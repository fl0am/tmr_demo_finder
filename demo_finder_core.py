"""
TMR Demo Lead Finder - core logic (no UI)
-----------------------------------------
Give it a booked demo address. It finds sports complexes, schools and stadiums
within a radius (OpenStreetMap), looks up addresses, pulls phone numbers, emails
and named contacts from their websites (all free), flags anything already in
HubSpot as a warm lead, and returns the best 5-10 to call.

Reuses the field counting, owner typing, website search and HubSpot matching
from complex_finder_v3.py, which must sit in the same folder.
"""

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
HEADERS = {"User-Agent": cf.USER_AGENT}
SKIP_DEMO_SITE_M = 500        # anything this close to the demo address is the demo site itself
MAX_POOL = 20                 # how many candidates get the full (slower) contact lookup
STADIUM_BONUS = 4

WARM_LABELS = {
    "Already in HubSpot": "In HubSpot - warm lead",
    "Likely duplicate": "Likely in HubSpot - warm lead",
    "Possible duplicate": "Maybe in HubSpot - check",
    "Parent org in HubSpot": "Parent org in HubSpot",
    "New": "New",
    "Not checked": "Not checked",
}


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
    """Fields, tracks and named parks/schools/stadiums around a point. Trimmed to the exact radius later."""
    near = bbox(lat, lon, radius_m)
    wide = bbox(lat, lon, radius_m + cf.NAME_MATCH_M)       # names just past the edge
    q = f"""
    [out:json][timeout:120];
    (
      nwr["leisure"="pitch"]["sport"~"{cf.FIELD_SPORTS}|{cf.TRACK_SPORTS}"]{near};
      nwr["leisure"="track"]["sport"~"{cf.TRACK_SPORTS}"]{near};
      nwr["leisure"~"^(park|sports_centre|stadium|recreation_ground)$"]["name"]{wide};
      nwr["amenity"~"^(school|college|university)$"]["name"]{wide};
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
    return {"street": street,
            "city": a.get("city") or a.get("town") or a.get("village") or a.get("hamlet", ""),
            "county": a.get("county", ""), "zip": (a.get("postcode") or "")[:5],
            "state": (a.get("ISO3166-2-lvl4") or "").replace("US-", "")}


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


# ---------------------------------------------------------------- helpers
def fmt_phone(p):
    d = cf.digits10(p)
    return f"({d[:3]}) {d[3:6]}-{d[6:]}" if d else ""


def place_type(r):
    if r["matched_kind"] == "stadium" or "stadium" in r["company"].lower():
        return "Stadium" if r["owner_type"] != "School" else "School stadium"
    return {"Public": "Public park", "School": "School", "Private": "Private complex"}.get(r["owner_type"], "Unknown")


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


# ---------------------------------------------------------------- main pipeline
def find_leads(address, radius_mi=25, n_leads=10, min_fields=2, hubspot_frames=(), look_for_people=True,
               progress=print):
    demo = geocode_address(address.strip())
    if not demo:
        raise ValueError("Couldn't find that address. Try adding the city and state, or use the zip code.")
    radius_m = radius_mi * MILE_M

    progress("Searching OpenStreetMap for fields, schools and stadiums...")
    data = fetch_area(round(demo["lat"], 5), round(demo["lon"], 5), int(radius_m))
    fields, tracks, places = cf.split_elements(data)
    if not fields and not tracks:
        return {"demo": demo, "leads": pd.DataFrame(), "others": pd.DataFrame(), "n_found": 0}
    units = cf.cluster(cf.build_units(fields, tracks), 400)
    df = cf.summarise(units, places, 1)

    # keep named places: complexes with enough fields, plus any school or stadium with a field
    df = df[(df["name"] != "") & ((df["units"] >= min_fields) | df["matched_kind"].isin(["school", "stadium"]))].copy()
    df["miles"] = [cf.haversine_m(demo["lat"], demo["lon"], la, lo) / MILE_M for la, lo in zip(df["lat"], df["lon"])]
    df = df[(df["miles"] <= radius_mi) & (df["miles"] * MILE_M > SKIP_DEMO_SITE_M)]
    if df.empty:
        return {"demo": demo, "leads": pd.DataFrame(), "others": pd.DataFrame(), "n_found": 0}
    df = df.sort_values("units", ascending=False).drop_duplicates("name").copy()   # a school split into two clusters

    tracks_only = df["units"] - df["grass"] - df["turf"] - df["unknown"]
    df["field_score"] = (cf.W_GRASS * df["grass"] + cf.W_UNKNOWN * df["unknown"] + cf.W_TURF * df["turf"]
                         + cf.W_TRACK * tracks_only
                         + STADIUM_BONUS * ((df["matched_kind"] == "stadium") | df["name"].str.contains("stadium", case=False)))
    df["pre_score"] = df["field_score"] * (1 - 0.4 * df["miles"] / radius_mi)     # closer is better
    df = df.sort_values("pre_score", ascending=False).reset_index(drop=True)
    n_found = len(df)
    pool = df.head(min(MAX_POOL, max(n_leads * 2, n_leads))).copy()
    others = df.iloc[len(pool):].copy()

    progress(f"Found {n_found} places. Looking up addresses for the top {len(pool)}...")
    addr = [reverse_geocode(round(la, 6), round(lo, 6)) for la, lo in zip(pool["lat"], pool["lon"])]
    pool = pd.concat([pool.reset_index(drop=True), pd.DataFrame(addr)], axis=1)
    for c in ("g_name", "g_address", "g_phone", "g_website", "g_types", "g_status", "g_zip"):
        pool[c] = ""
    pool = pool.fillna("")
    pool["company"] = pool["name"]
    pool["website"] = pool["osm_website"]
    pool["owner_type"] = pool.apply(cf.classify, axis=1)
    pool["agency_key"] = pool.apply(cf.agency_key, axis=1)

    sites = sorted({u for u in pool["website"] if u and not re.search(r"facebook|instagram|google\.|yelp", u, re.I)})
    progress(f"Checking {len(sites)} websites for phone numbers and emails...")
    with ThreadPoolExecutor(max_workers=8) as ex:
        info = dict(zip(sites, ex.map(site_info, sites)))
    found_people = {}
    if look_for_people:
        jobs = sorted({(r["website"], people_kind(r)) for _, r in pool.iterrows()
                       if r["website"] in info and people_kind(r)})
        if jobs:
            progress(f"Looking for athletic directors, parks directors and owners on {len(jobs)} websites...")
            with ThreadPoolExecutor(max_workers=6) as ex:
                found_people = dict(zip(jobs, ex.map(lambda j: site_people(*j), jobs)))

    rows = []
    for _, r in pool.iterrows():
        r = r.to_dict()
        own = info.get(r["website"])
        r["phone"] = fmt_phone(r["osm_phone"]) or cf.best_phone(own, r["owner_type"])
        r["inbox"] = r["osm_email"] or cf.best_inbox(own, r["website"], r["owner_type"])
        r["contact_name"] = r["contact_title"] = r["contact_email"] = ""
        kind = people_kind(r)
        res = found_people.get((r["website"], kind))
        if res:
            p = best_person(res, cf.norm_match(r["company"]))
            if p:
                r["contact_name"] = p["name"]
                r["contact_title"] = cf.PROFILES[kind]["label"] + " (per website)"
                r["contact_email"] = p["email"]
        if r["owner_type"] == "Public":
            r["buyer"] = r["operator"] or f"{r['agency_key']} parks & recreation"
        else:
            r["buyer"] = r["company"]
        r["ask_for"] = cf.TARGET_ROLE.get(r["owner_type"], "Facility Manager")
        r["type"] = place_type(r)
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
    leads = pool.head(n_leads).copy()
    others = pd.concat([pool.iloc[n_leads:], others], ignore_index=True)
    leads["maps_link"] = [f"https://www.google.com/maps/search/?api=1&query={la},{lo}"
                          for la, lo in zip(leads["lat"], leads["lon"])]
    return {"demo": demo, "leads": leads, "others": others, "n_found": n_found, "hubspot_checked": hs["loaded"]}
