"""
TMR Demo Lead Finder - core logic (no UI)
-----------------------------------------
Give it a booked demo address. It finds sports complexes, schools, stadiums and
youth sports clubs within a radius, fills in phones/emails/contacts from free
sources, flags anything already in HubSpot as a warm lead, lists HubSpot
companies near the demo, and returns the best 5-10 to call.

Free sources only:
  OpenStreetMap (fields, clubs, towns, addresses), each place's own website,
  and the federal public school directory (NCES, via the Urban Institute API).

Reuses complex_finder_v3.py, which must sit in the same folder.
"""

import difflib
import math
import re
import time
from concurrent.futures import ThreadPoolExecutor
from functools import lru_cache

import pandas as pd
import requests

import complex_finder_v3 as cf

MILE_M = 1609.344
OVERPASS_URLS = [cf.OVERPASS_URL,
                 "https://overpass.private.coffee/api/interpreter",
                 "https://overpass.kumi.systems/api/interpreter"]
SEARCH_URL = "https://nominatim.openstreetmap.org/search"
SCHOOLS_API = "https://educationdata.urban.org/api/v1/schools/ccd/directory/{year}/?fips={fips}"
HEADERS = {"User-Agent": cf.USER_AGENT}
SKIP_DEMO_SITE_M = 500        # anything this close to the demo address is the demo site itself
MAX_POOL = 20                 # candidates that get the full (slower) contact lookup
MAX_UNNAMED = 8               # unnamed field groups to label with a nearby road
CLUB_SLOTS = 2                # call-list spots kept for youth clubs when any exist
STADIUM_BONUS = 4
CLUB_BASE, SCHOOL_NO_FIELDS_BASE = 4, 4

WARM_LABELS = {
    "Already in HubSpot": "In HubSpot - warm lead",
    "Likely duplicate": "Likely in HubSpot - warm lead",
    "Possible duplicate": "Maybe in HubSpot - check",
    "Parent org in HubSpot": "Parent org in HubSpot",
    "New": "New",
    "Not checked": "Not checked",
}

CLUB_SPORTS = re.compile(r"soccer|american_football|baseball|softball|lacrosse|rugby|field_hockey|ultimate|"
                         r"cricket|multi", re.I)
CLUB_NAME = re.compile(r"soccer|futbol|football|baseball|softball|lacrosse|rugby|field hockey|ultimate|cricket|"
                       r"little league|youth|athletic association|sports association|\bfc\b|\bsc\b|\bsa\b", re.I)
CLUB_OPERATOR = re.compile(CLUB_NAME.pattern + r"|league|association|\bclub\b", re.I)

STATE_FIPS = {"AL": 1, "AK": 2, "AZ": 4, "AR": 5, "CA": 6, "CO": 8, "CT": 9, "DE": 10, "DC": 11, "FL": 12,
              "GA": 13, "HI": 15, "ID": 16, "IL": 17, "IN": 18, "IA": 19, "KS": 20, "KY": 21, "LA": 22,
              "ME": 23, "MD": 24, "MA": 25, "MI": 26, "MN": 27, "MS": 28, "MO": 29, "MT": 30, "NE": 31,
              "NV": 32, "NH": 33, "NJ": 34, "NM": 35, "NY": 36, "NC": 37, "ND": 38, "OH": 39, "OK": 40,
              "OR": 41, "PA": 42, "RI": 44, "SC": 45, "SD": 46, "TN": 47, "TX": 48, "UT": 49, "VT": 50,
              "VA": 51, "WA": 53, "WV": 54, "WI": 55, "WY": 56}


# ---------------------------------------------------------------- lookups (cached for the life of the app)
@lru_cache(maxsize=512)
def geocode_address(address):
    r = requests.get(SEARCH_URL, headers=HEADERS, timeout=30,
                     params={"q": address, "format": "jsonv2", "countrycodes": "us",
                             "limit": 1, "addressdetails": 1})
    r.raise_for_status()
    hits = r.json()
    if not hits:
        return None
    h = hits[0]
    a = h.get("address", {})
    return {"lat": float(h["lat"]), "lon": float(h["lon"]), "label": h.get("display_name", ""),
            "state": (a.get("ISO3166-2-lvl4") or "").replace("US-", "")}


def bbox(lat, lon, radius_m):
    """south,west,north,east box around a point (a box query is far faster than a circle)."""
    dlat = radius_m / 111320
    dlon = radius_m / (111320 * max(0.2, math.cos(math.radians(lat))))
    return f"({lat - dlat:.5f},{lon - dlon:.5f},{lat + dlat:.5f},{lon + dlon:.5f})"


@lru_cache(maxsize=128)
def fetch_area(lat, lon, radius_m):
    """Fields, tracks, named parks/schools/stadiums, sports clubs and towns. Trimmed to the exact radius later."""
    near = bbox(lat, lon, radius_m)
    wide = bbox(lat, lon, radius_m + cf.NAME_MATCH_M)       # names just past the edge
    q = f"""
    [out:json][timeout:120];
    (
      nwr["leisure"="pitch"]["sport"~"{cf.FIELD_SPORTS}|{cf.TRACK_SPORTS}"]{near};
      nwr["leisure"="track"]["sport"~"{cf.TRACK_SPORTS}"]{near};
      nwr["leisure"~"^(park|sports_centre|stadium|recreation_ground)$"]["name"]{wide};
      nwr["amenity"~"^(school|college|university)$"]["name"]{wide};
      nwr["club"="sport"]["name"]{wide};
      node["place"~"^(city|town|village|hamlet|suburb)$"]["name"]{wide};
    );
    out center tags;
    """
    problems = []
    for url in OVERPASS_URLS:
        for attempt in range(3):
            try:
                r = requests.post(url, data={"data": q}, headers=HEADERS, timeout=150)
                if r.status_code == 200:
                    return r.json()
                problems.append(f"{url.split('/')[2]}: {r.status_code}")
                if r.status_code not in (429, 502, 503, 504):
                    break                                  # a real error, not "busy" - try the next server
            except Exception as e:
                problems.append(f"{url.split('/')[2]}: {type(e).__name__}")
            time.sleep(5 * (attempt + 1))                  # busy: wait 5s, then 10s
    raise RuntimeError("The OpenStreetMap servers didn't answer (" + "; ".join(problems) +
                       "). Wait a few minutes and try again.")


_last_nominatim = [0.0]


@lru_cache(maxsize=4096)
def reverse_geocode(lat, lon):
    wait = 1.1 - (time.time() - _last_nominatim[0])      # Nominatim allows 1 request per second
    if wait > 0:
        time.sleep(wait)
    _last_nominatim[0] = time.time()
    try:
        r = requests.get(cf.NOMINATIM_URL, headers=HEADERS, timeout=30,
                         params={"lat": lat, "lon": lon, "format": "jsonv2", "zoom": 17})
        a = r.json().get("address", {})
    except Exception:
        a = {}
    street = " ".join(x for x in (a.get("house_number", ""), a.get("road", "")) if x)
    return {"street": street, "road": a.get("road", ""),
            "city": a.get("city") or a.get("town") or a.get("village") or a.get("hamlet", ""),
            "county": a.get("county", ""), "zip": (a.get("postcode") or "")[:5],
            "state": (a.get("ISO3166-2-lvl4") or "").replace("US-", "")}


@lru_cache(maxsize=16)
def school_directory(state):
    """Public high schools in a state from the federal school directory. Empty list if unavailable."""
    fips = STATE_FIPS.get(state)
    if not fips:
        return ()
    for year in (2023, 2022, 2021):
        url, rows = SCHOOLS_API.format(year=year, fips=fips), []
        try:
            while url:
                r = requests.get(url, headers=HEADERS, timeout=90)
                r.raise_for_status()
                j = r.json()
                rows += j.get("results", [])
                url = j.get("next")
        except Exception:
            continue
        out = []
        for s in rows:
            name = str(s.get("school_name") or "")
            level = s.get("school_level")
            if s.get("latitude") in (None, "") or s.get("school_status") == 2:
                continue
            if level == 3 or (level == 4 and re.search(r"\bhigh\b", name, re.I)):
                out.append({"name": name.title() if name.isupper() else name, "lat": float(s["latitude"]),
                            "lon": float(s["longitude"]), "phone": str(s.get("phone") or ""),
                            "street": str(s.get("street_location") or "").title(),
                            "city": str(s.get("city_location") or "").title(),
                            "zip": str(s.get("zip_location") or "")[:5], "district": str(s.get("lea_name") or "")})
        if out:
            return tuple(out)
    return ()


@lru_cache(maxsize=2048)
def site_info(url):
    try:
        return cf.site_contact_info(url)
    except Exception:
        return {"phones": [], "tel_links": [], "emails": []}


@lru_cache(maxsize=2048)
def site_people(url, kind):
    try:
        return cf.crawl_site(url, kind)
    except Exception:
        return {"ads": [], "pairs": [], "emails": []}


# ---------------------------------------------------------------- HubSpot (from uploaded files, nothing saved)
def load_hubspot_frames(frames, states):
    """Same matching index as complex_finder_v3.load_hubspot, built from uploaded DataFrames."""
    doms, regs, emails, phones, names = {}, {}, {}, {}, {}
    for h in frames:
        h = h.fillna("").astype(str)
        for rec in h.to_dict("records"):
            company = cf.pick(rec, "company name", "name", "associated company")
            label = company or " ".join(x for x in (cf.pick(rec, "first name"), cf.pick(rec, "last name")) if x)
            raw_st = cf.pick(rec, "state/region", "state", "company state/region").lower().strip()
            st = cf.STATE_ALIASES.get(raw_st, raw_st.upper() if raw_st else "")
            if st and states and st not in states:
                continue
            city = cf.pick(rec, "city", "company city").lower().strip()
            d = cf.domain(cf.pick(rec, "company domain name", "website url", "domain", "website"))
            e = cf.pick(rec, "email").lower()
            for dd, lab in ((d, label), (e.split("@")[-1] if e else "", label + " (contact email domain)")):
                if dd:
                    doms.setdefault(dd, lab)
                    regs.setdefault(cf.reg_domain(dd), lab)
            if e:
                emails.setdefault(e, label)
            ph = cf.digits10(cf.pick(rec, "phone number", "phone", "mobile phone number"))
            if ph:
                phones.setdefault(ph, label)
            k = cf.name_key(company)
            if company and len(k) >= 4:
                names.setdefault(k, []).append((cf.norm_match(company), city, label))
    return {"loaded": bool(frames), "doms": doms, "regs": regs, "emails": emails, "phones": phones, "names": names}


def _cols(df):
    return {c.lower().strip(): c for c in df.columns}


def hubspot_nearby(frames, towns, radius_mi):
    """HubSpot companies whose city is a town within the radius, with their contacts attached."""
    town_at = {}
    for t in towns:
        if t["miles"] <= radius_mi:
            k = t["name"].lower().strip()
            if k not in town_at or t["miles"] < town_at[k]["miles"]:
                town_at[k] = t
    companies, contacts = [], []
    for f in frames:
        c = _cols(f)
        if "first name" in c or "last name" in c:
            contacts.append(f)
        elif "company name" in c:
            companies.append(f)
    by_id, by_name = {}, {}
    for f in contacts:
        f = f.fillna("").astype(str)
        for rec in f.to_dict("records"):
            person = {"name": " ".join(x for x in (cf.pick(rec, "first name"), cf.pick(rec, "last name")) if x),
                      "email": cf.pick(rec, "email"), "phone": fmt_phone(cf.pick(rec, "phone number")),
                      "owner": cf.pick(rec, "contact owner"), "lead_status": cf.pick(rec, "lead status"),
                      "last_activity": cf.pick(rec, "last activity date")}
            cid = cf.pick(rec, "primary associated company id")
            if cid:
                by_id.setdefault(cid, []).append(person)
            elif cf.pick(rec, "associated company"):
                by_name.setdefault(cf.norm_match(cf.pick(rec, "associated company")), []).append(person)
    rows = []
    for f in companies:
        f = f.fillna("").astype(str)
        for rec in f.to_dict("records"):
            city = cf.pick(rec, "city")
            t = town_at.get(city.lower().strip())
            if not t:
                continue
            name = cf.pick(rec, "company name")
            people = by_id.get(cf.pick(rec, "record id"), []) + by_name.get(cf.norm_match(name), [])
            dates = pd.to_datetime([cf.pick(rec, "last activity date")] + [p["last_activity"] for p in people],
                                   errors="coerce")
            last = dates.max() if dates.notna().any() else pd.NaT
            rows.append({"company": name, "city": city, "miles": round(t["miles"], 1),
                         "lat": t["lat"], "lon": t["lon"], "phone": fmt_phone(cf.pick(rec, "phone number")),
                         "owner": cf.pick(rec, "company owner"), "industry": cf.pick(rec, "industry"),
                         "last_activity": last, "n_contacts": len(people), "contacts": people,
                         "record_id": cf.pick(rec, "record id")})
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    out["never_worked"] = out["last_activity"].isna()
    return out.sort_values(["never_worked", "last_activity", "miles"], ascending=[False, True, True],
                           na_position="first").reset_index(drop=True)


# ---------------------------------------------------------------- helpers
def fmt_phone(p):
    d = cf.digits10(p)
    return f"({d[:3]}) {d[3:6]}-{d[6:]}" if d else ""


def miles_from(demo, lat, lon):
    return cf.haversine_m(demo["lat"], demo["lon"], lat, lon) / MILE_M


def place_type(r):
    if r["matched_kind"] == "club":
        return "Youth sports club"
    if r["matched_kind"] == "stadium" or "stadium" in r["company"].lower():
        return "School stadium" if r["owner_type"] == "School" else "Stadium"
    if r["owner_type"] == "School" and r["units"] == 0:
        return "School (fields not mapped)"
    return {"Public": "Public park", "School": "School", "Private": "Private complex"}.get(r["owner_type"], "Other")


def people_kind(r):
    if r["owner_type"] == "School":
        return "ad"
    if r["owner_type"] == "Public":
        return "parks" if re.search(r"\.(gov|us)$|county|cityof|parks|recreation", cf.domain_of(r["website"])) else ""
    return "private"


def best_person(res, place_names):
    names = {}
    for nm, em in res.get("ads", []):
        key = cf.split_name(nm)
        if not all(key) or not cf.is_real_name(nm, em) or cf.norm_match(nm) in place_names:
            continue
        cur = names.setdefault(key, {"name": nm, "count": 0, "email": ""})
        cur["count"] += 1
        if em and not cur["email"]:
            cur["email"] = em
    if not names:
        return None
    return max(names.values(), key=lambda x: (x["count"], bool(x["email"])))


def same_place(a, b):
    ka, kb = cf.name_key(a), cf.name_key(b)
    if len(ka) >= 3 and ka == kb:
        return True
    return difflib.SequenceMatcher(None, cf.norm_match(a), cf.norm_match(b)).ratio() >= 0.8


def blank_row(**kw):
    row = {"cluster": -1, "name": "", "lat": 0.0, "lon": 0.0, "matched_kind": "", "operator": "",
           "osm_website": "", "osm_phone": "", "osm_email": "", "units": 0, "grass": 0, "turf": 0,
           "unknown": 0, "tracks": 0, "sports": "", "note": "", "nces_phone": "", "street": "", "city": "",
           "zip": "", "district": ""}
    row.update(kw)
    return row


def separate(data):
    """Pull towns, sports clubs and field operators out of the raw map data."""
    towns, clubs, ops, rest = [], [], {}, []
    for el in data.get("elements", []):
        tags = el.get("tags", {})
        c = el.get("center") or {"lat": el.get("lat"), "lon": el.get("lon")}
        if c.get("lat") is None:
            continue
        lat, lon = c["lat"], c["lon"]
        if tags.get("place"):
            towns.append({"name": tags.get("name", ""), "lat": lat, "lon": lon})
            continue
        if tags.get("club") == "sport" and not tags.get("leisure") and not tags.get("amenity"):
            name = tags.get("name", "")
            if CLUB_SPORTS.search(tags.get("sport", "")) or CLUB_NAME.search(name):
                clubs.append(blank_row(name=name, lat=lat, lon=lon, matched_kind="club",
                                       osm_website=tags.get("website", tags.get("contact:website", "")),
                                       osm_phone=tags.get("phone", tags.get("contact:phone", "")),
                                       osm_email=tags.get("email", tags.get("contact:email", "")),
                                       sports=tags.get("sport", "").replace(";", ", ")))
            continue
        rest.append(el)
        op = tags.get("operator", "")
        if tags.get("leisure") == "pitch" and op and CLUB_OPERATOR.search(op) \
                and not cf.PUBLIC_WORDS.search(op) and not cf.SCHOOL_WORDS.search(op):
            o = ops.setdefault(cf.norm_match(op), {"name": op, "pts": [], "sports": set()})
            o["pts"].append((lat, lon))
            o["sports"].update(s for s in tags.get("sport", "").split(";") if s)
    for o in ops.values():                                  # clubs known only from the fields they run
        lat = sum(p[0] for p in o["pts"]) / len(o["pts"])
        lon = sum(p[1] for p in o["pts"]) / len(o["pts"])
        if any(same_place(o["name"], c["name"]) for c in clubs):
            for c in clubs:
                if same_place(o["name"], c["name"]):
                    c["units"] = max(c["units"], len(o["pts"]))
            continue
        clubs.append(blank_row(name=o["name"], lat=lat, lon=lon, matched_kind="club", units=len(o["pts"]),
                               unknown=len(o["pts"]), sports=", ".join(sorted(o["sports"]))))
    return towns, clubs, {"elements": rest}


# ---------------------------------------------------------------- main pipeline
def find_leads(address, radius_mi=25, n_leads=10, min_fields=2, hubspot_frames=(), look_for_people=True,
               progress=print):
    notes = []
    demo = geocode_address(address.strip())
    if not demo:
        raise ValueError("Couldn't find that address. Try adding the city and state, or use the zip code.")
    radius_m = radius_mi * MILE_M

    progress("Searching OpenStreetMap for fields, schools, stadiums and clubs...")
    data = fetch_area(round(demo["lat"], 5), round(demo["lon"], 5), int(radius_m))
    towns, clubs, data = separate(data)
    for t in towns:
        t["miles"] = miles_from(demo, t["lat"], t["lon"])
    fields, tracks, places = cf.split_elements(data)
    df = pd.DataFrame()
    if fields or tracks:
        units = cf.cluster(cf.build_units(fields, tracks), 400)
        df = cf.summarise(units, places, 1)
    if df.empty:
        df = pd.DataFrame([blank_row()]).iloc[:0]
    for c, v in blank_row().items():
        if c not in df:
            df[c] = v

    # club fields: name the park they play at
    named = df[df["name"] != ""]
    for c in clubs:
        if c["units"] and not named.empty:
            d = [cf.haversine_m(c["lat"], c["lon"], la, lo) for la, lo in zip(named["lat"], named["lon"])]
            i = min(range(len(d)), key=d.__getitem__)
            if d[i] < 800:
                c["note"] = f"Plays at {named.iloc[i]['name']}"
    df = pd.concat([df, pd.DataFrame(clubs)], ignore_index=True) if clubs else df

    # federal school directory: phones for schools we found, plus high schools with no mapped fields
    if demo["state"]:
        progress("Checking the federal school directory...")
        directory = school_directory(demo["state"])
        if not directory:
            notes.append("The federal school directory didn't answer, so school phones come from websites only.")
        for s in directory:
            if miles_from(demo, s["lat"], s["lon"]) > radius_mi:
                continue
            cand = df[df["matched_kind"].isin(["school", "stadium"]) & (df["name"] != "")]
            hit = None
            for i, r in cand.iterrows():
                if cf.haversine_m(s["lat"], s["lon"], r["lat"], r["lon"]) < 1500 and same_place(s["name"], r["name"]):
                    hit = i
                    break
            if hit is not None:
                df.loc[hit, ["nces_phone", "district"]] = [s["phone"], s["district"]]
            else:
                df = pd.concat([df, pd.DataFrame([blank_row(name=s["name"], lat=s["lat"], lon=s["lon"],
                                                            matched_kind="school", nces_phone=s["phone"],
                                                            street=s["street"], city=s["city"], zip=s["zip"],
                                                            district=s["district"])])], ignore_index=True)

    df["miles"] = [miles_from(demo, la, lo) for la, lo in zip(df["lat"], df["lon"])]
    df = df[(df["miles"] <= radius_mi) & (df["miles"] * MILE_M > SKIP_DEMO_SITE_M)].copy()

    # unnamed field groups: listed separately, labeled by road
    unnamed = df[(df["name"] == "") & (df["units"] >= min_fields)].sort_values("units", ascending=False)
    unnamed = unnamed.head(MAX_UNNAMED).copy()

    keep = (df["name"] != "") & ((df["units"] >= min_fields) | df["matched_kind"].isin(["school", "stadium", "club"]))
    df = df[keep].sort_values("units", ascending=False).drop_duplicates("name").copy()
    if df.empty and unnamed.empty:
        return {"demo": demo, "leads": pd.DataFrame(), "others": pd.DataFrame(), "unnamed": pd.DataFrame(),
                "hubspot_near": hubspot_nearby(list(hubspot_frames), towns, radius_mi), "n_found": 0,
                "hubspot_checked": bool(hubspot_frames), "notes": notes}

    for c in ("g_name", "g_address", "g_phone", "g_website", "g_types", "g_status", "g_zip"):
        df[c] = ""
    df["company"] = df["name"]
    df["website"] = df["osm_website"]
    df["owner_type"] = df.apply(cf.classify, axis=1) if len(df) else []
    df.loc[df["matched_kind"] == "club", "owner_type"] = "Private"
    tracks_only = df["units"] - df["grass"] - df["turf"] - df["unknown"]
    df["field_score"] = (cf.W_GRASS * df["grass"] + cf.W_UNKNOWN * df["unknown"] + cf.W_TURF * df["turf"]
                         + cf.W_TRACK * tracks_only
                         + STADIUM_BONUS * ((df["matched_kind"] == "stadium") | df["name"].str.contains("stadium", case=False))
                         + CLUB_BASE * (df["matched_kind"] == "club")
                         + SCHOOL_NO_FIELDS_BASE * ((df["matched_kind"] == "school") & (df["units"] == 0)))
    df["pre_score"] = df["field_score"] * (1 - 0.4 * df["miles"] / radius_mi)     # closer is better
    df["type"] = [place_type({**r, "company": r["name"]}) for r in df.to_dict("records")]
    df = df.sort_values("pre_score", ascending=False).reset_index(drop=True)
    n_found = len(df)

    is_club = df["matched_kind"] == "club"
    size = min(MAX_POOL, max(n_leads * 2, n_leads))
    pool_idx = list(df[~is_club].index[:size - min(3, is_club.sum())]) + list(df[is_club].index[:3])
    pool = df.loc[pool_idx].copy()
    others = df.drop(index=pool_idx).copy()

    progress(f"Found {n_found} places. Looking up addresses for the top {len(pool)}...")
    addr = []
    for _, r in pool.iterrows():
        if r["street"] and r["city"]:
            addr.append({"street": r["street"], "road": "", "city": r["city"], "county": "", "zip": r["zip"],
                         "state": demo["state"]})
        else:
            addr.append(reverse_geocode(round(r["lat"], 6), round(r["lon"], 6)))
    pool = pd.concat([pool.drop(columns=["street", "city", "zip"]).reset_index(drop=True), pd.DataFrame(addr)],
                     axis=1).fillna("")
    pool["agency_key"] = pool.apply(cf.agency_key, axis=1)

    if not unnamed.empty:
        progress(f"Labeling {len(unnamed)} unnamed field groups...")
        labels = []
        for _, r in unnamed.iterrows():
            a = reverse_geocode(round(r["lat"], 6), round(r["lon"], 6))
            near = ", ".join(x for x in (a["road"], a["city"]) if x) or "unknown road"
            labels.append(f"{int(r['units'])} unnamed fields near {near}")
        unnamed["label"] = labels

    sites = sorted({u for u in pool["website"] if u and not re.search(r"facebook|instagram|google\.|yelp", u, re.I)})
    progress(f"Checking {len(sites)} websites for phone numbers and emails...")
    with ThreadPoolExecutor(max_workers=8) as ex:
        info = dict(zip(sites, ex.map(site_info, sites)))
    found_people = {}
    if look_for_people:
        jobs = sorted({(r["website"], people_kind(r)) for _, r in pool.iterrows()
                       if r["website"] in info and people_kind(r)})
        if jobs:
            progress(f"Looking for athletic directors, parks directors and club leaders on {len(jobs)} websites...")
            with ThreadPoolExecutor(max_workers=6) as ex:
                found_people = dict(zip(jobs, ex.map(lambda j: site_people(*j), jobs)))

    rows = []
    for _, r in pool.iterrows():
        r = r.to_dict()
        own = info.get(r["website"])
        r["phone"] = fmt_phone(r["osm_phone"]) or fmt_phone(r["nces_phone"]) or cf.best_phone(own, r["owner_type"])
        r["inbox"] = r["osm_email"] or cf.best_inbox(own, r["website"], r["owner_type"])
        r["contact_name"] = r["contact_title"] = r["contact_email"] = ""
        kind = people_kind(r)
        res = found_people.get((r["website"], kind))
        if res:
            p = best_person(res, cf.norm_match(r["company"]))
            if p:
                r["contact_name"] = p["name"]
                r["contact_title"] = ("Club leader" if r["matched_kind"] == "club" else cf.PROFILES[kind]["label"]) \
                    + " (per website)"
                r["contact_email"] = p["email"]
        if r["owner_type"] == "Public":
            r["buyer"] = r["operator"] or f"{r['agency_key']} parks & recreation"
        elif r["owner_type"] == "School" and r["district"]:
            r["buyer"] = f"{r['company']} ({r['district']})"
        else:
            r["buyer"] = r["company"]
        r["ask_for"] = ("Club president or field coordinator" if r["matched_kind"] == "club"
                        else cf.TARGET_ROLE.get(r["owner_type"], "Facility Manager"))
        rows.append(r)
    pool = pd.DataFrame(rows)

    # HubSpot: flag, never drop
    states = {s for s in list(pool["state"]) + [demo["state"]] if s}
    hs = load_hubspot_frames(list(hubspot_frames), states)
    reg_count, phone_count = {}, {}
    for _, r in pool.iterrows():
        d = cf.domain(r["website"])
        if d:
            reg_count[cf.reg_domain(d)] = reg_count.get(cf.reg_domain(d), 0) + 1
        p = cf.digits10(r["phone"])
        if p:
            phone_count[p] = phone_count.get(p, 0) + 1
    status = [cf.hubspot_status({**r, "contact_email": r["contact_email"] or r["inbox"]}, hs, reg_count, phone_count)
              for r in pool.to_dict("records")]
    pool["hubspot"] = [WARM_LABELS.get(s, s) for s, _ in status]
    pool["hubspot_match"] = [m for _, m in status]

    pool["score"] = (pool["pre_score"] + 6 * (pool["phone"] != "")
                     + 3 * ((pool["contact_email"] != "") | (pool["inbox"] != "")) + 3 * (pool["contact_name"] != ""))
    pool = pool.sort_values("score", ascending=False).reset_index(drop=True)
    club_rows = pool[pool["matched_kind"] == "club"].head(CLUB_SLOTS)
    rest = pool.drop(index=club_rows.index).head(n_leads - len(club_rows))
    leads = pd.concat([rest, club_rows]).sort_values("score", ascending=False).reset_index(drop=True)
    others = pd.concat([pool.drop(index=list(club_rows.index) + list(rest.index)), others], ignore_index=True)
    leads["maps_link"] = [f"https://www.google.com/maps/search/?api=1&query={la},{lo}"
                          for la, lo in zip(leads["lat"], leads["lon"])]

    near = hubspot_nearby(list(hubspot_frames), towns, radius_mi) if hubspot_frames else pd.DataFrame()
    return {"demo": demo, "leads": leads, "others": others, "unnamed": unnamed, "hubspot_near": near,
            "n_found": n_found, "hubspot_checked": hs["loaded"], "notes": notes}
