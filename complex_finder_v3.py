"""
Sports Complex Lead Finder  v3
------------------------------
Finds clusters of outdoor sports fields and tracks in a US state using
OpenStreetMap, groups them into complexes, enriches them with business and
contact info, checks them against HubSpot, and exports a ranked lead list.

Usage (from the tmr-leads folder):
    python complex_finder_v3.py --geocode --min-fields 4                        # basic run
    python complex_finder_v3.py --geocode --places --emails --min-fields 4      # + Google + website emails
    python complex_finder_v3.py --geocode --places --emails --find-contacts --google-budget 0 --min-fields 4
                                    # + athletic/parks directors from websites, saved Google answers only
    python complex_finder_v3.py --state FL --geocode                           # any state

  --google-budget N   max NEW Google calls this run (default 900). 0 = only reuse saved answers, no charges.
                      Works without an API key when it's 0.

Folders the script reads from (create them, drop CSVs in, re-run):
    hubspot/    your HubSpot Companies and Contacts exports  -> duplicate check
    seamless/   CSVs exported from Seamless.AI               -> named contacts
    agencies/   agencies_<STATE>.csv (auto-created)           -> parks dept contacts
                (a contact name or phone is used even without an email)

Requires: pip install requests pandas scikit-learn openpyxl
"""

import argparse
import html as _html
from concurrent.futures import ThreadPoolExecutor
import csv
import difflib
import glob
import json
import math
import os
import re
import time
from urllib.parse import urljoin, urlparse

import numpy as np
import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from sklearn.cluster import DBSCAN

# ---------------------------------------------------------------- settings
OVERPASS_URL = "https://overpass-api.de/api/interpreter"
NOMINATIM_URL = "https://nominatim.openstreetmap.org/reverse"
PLACES_URL = "https://places.googleapis.com/v1/places:searchText"
USER_AGENT = "TMR-complex-finder/0.3.1 (contact: michael.floam123@gmail.com)"  # put your email here

# Field sports that get painted. Tennis/basketball/pickleball are left out on purpose.
FIELD_SPORTS = ("soccer|american_football|baseball|softball|lacrosse|rugby|rugby_union|"
                "rugby_league|field_hockey|ultimate|cricket|gaelic_games|australian_football")
TRACK_SPORTS = "athletics|running"

TRACK_ABSORB_M = 80      # a track this close to a field is treated as wrapping it
NAME_MATCH_M = 600       # look this far for a named park/school/sports centre
PLACES_MATCH_M = 800     # accept a Google result only if it's this close

# Default scoring (also editable in the Settings tab of the spreadsheet)
W_GRASS, W_UNKNOWN, W_TURF, W_TRACK = 3, 2, 1, 2
TIER_A, TIER_B = 20, 10

TIER_COLORS = {"A": "C6EFCE", "B": "FFEB9C", "C": "EDEDED"}  # green / yellow / gray

GOOD_TITLES = ("director", "parks", "recreation", "facilit", "grounds", "athletic",
               "operations", "owner", "general manager", "president", "superintendent",
               "maintenance", "turf", "field")

HEADERS = {"User-Agent": USER_AGENT}


# ---------------------------------------------------------------- helpers
def haversine_m(lat1, lon1, lat2, lon2):
    r = 6371000
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = p2 - p1, math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def load_json(path, default):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return default


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f)


def domain_of(url):
    if not url or not isinstance(url, str):
        return ""
    if "://" not in url:
        url = "http://" + url
    d = urlparse(url.strip()).netloc.lower()
    return d[4:] if d.startswith("www.") else d


def digits10(phone):
    d = re.sub(r"\D", "", str(phone or ""))
    return d[-10:] if len(d) >= 10 else ""


STOP = {"the", "inc", "llc", "co", "corp", "of", "and", "&", "at", "complex", "sports",
        "park", "center", "centre", "fields", "field", "athletic", "athletics", "recreation",
        "rec", "department", "dept", "ga", "georgia"}


def norm_name(name):
    words = re.findall(r"[a-z0-9]+", str(name or "").lower())
    return " ".join(w for w in words if w not in STOP)


def clean(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    return str(v).strip()


# ---------------------------------------------------------------- 1. pull OSM
def fetch_osm(state, cache_dir):
    path = os.path.join(cache_dir, f"osm_v2_{state}.json")
    if os.path.exists(path):
        print(f"Using cached OpenStreetMap data ({path}).")
        return load_json(path, {})
    print(f"Querying OpenStreetMap for {state} (a few minutes)...")
    q = f"""
    [out:json][timeout:900];
    area["ISO3166-2"="US-{state}"]->.s;
    (
      nwr["leisure"="pitch"]["sport"~"{FIELD_SPORTS}|{TRACK_SPORTS}"](area.s);
      nwr["leisure"="track"]["sport"~"{TRACK_SPORTS}"](area.s);
      nwr["leisure"~"^(park|sports_centre|stadium|recreation_ground)$"]["name"](area.s);
      nwr["amenity"~"^(school|college|university)$"]["name"](area.s);
    );
    out center tags;
    """
    r = requests.post(OVERPASS_URL, data={"data": q}, headers=HEADERS, timeout=1000)
    r.raise_for_status()
    data = r.json()
    save_json(path, data)
    return data


def split_elements(data):
    fields, tracks, places = [], [], []
    for el in data.get("elements", []):
        tags = el.get("tags", {})
        c = el.get("center") or {"lat": el.get("lat"), "lon": el.get("lon")}
        if c.get("lat") is None:
            continue
        lat, lon = c["lat"], c["lon"]
        sport = tags.get("sport", "")
        leisure = tags.get("leisure", "")
        base = {"osm_id": f"{el['type']}/{el['id']}", "lat": lat, "lon": lon,
                "sport": sport, "surface": tags.get("surface", "")}
        if leisure == "track" or (leisure == "pitch" and re.fullmatch(f"({TRACK_SPORTS})(;.*)?", sport)):
            tracks.append({**base, "kind": "track"})
        elif leisure == "pitch":
            fields.append({**base, "kind": "field"})
        elif tags.get("name"):
            kind = "school" if tags.get("amenity") in ("school", "college", "university") else leisure
            places.append({"lat": lat, "lon": lon, "name": tags["name"], "kind": kind,
                           "operator": tags.get("operator", ""),
                           "website": tags.get("website", tags.get("contact:website", "")),
                           "phone": tags.get("phone", tags.get("contact:phone", "")),
                           "email": tags.get("email", tags.get("contact:email", ""))})
    return fields, tracks, places


def surface_class(s):
    s = (s or "").lower()
    if s in ("grass", "natural_grass"):
        return "grass"
    if "artificial" in s or "turf" in s or s in ("astroturf", "synthetic"):
        return "turf"
    return "unknown"


# ---------------------------------------------------------------- 2. cluster
def build_units(fields, tracks):
    """Fields count 1 each. A track wrapping a field is absorbed into it; a lone track counts 1."""
    units = [dict(f) for f in fields]
    for u in units:
        u["counted"] = True
    if fields:
        flat = np.array([f["lat"] for f in fields])
        flon = np.array([f["lon"] for f in fields])
    for t in tracks:
        t = dict(t)
        absorbed = False
        if fields:
            # quick box filter, then exact distance
            near = np.where((abs(flat - t["lat"]) < 0.002) & (abs(flon - t["lon"]) < 0.002))[0]
            absorbed = any(haversine_m(t["lat"], t["lon"], flat[i], flon[i]) <= TRACK_ABSORB_M for i in near)
        t["counted"] = not absorbed
        units.append(t)
    return pd.DataFrame(units)


def cluster(units, radius_m):
    coords = np.radians(units[["lat", "lon"]].to_numpy())
    labels = DBSCAN(eps=radius_m / 6371000, min_samples=1, metric="haversine").fit(coords).labels_
    units = units.copy()
    units["cluster"] = labels
    return units


# ---------------------------------------------------------------- 3. summarise
def summarise(units, places, min_fields):
    plat = np.array([p["lat"] for p in places]) if places else np.array([])
    plon = np.array([p["lon"] for p in places]) if places else np.array([])
    rows = []
    for cid, g in units.groupby("cluster"):
        counted = g[g["counted"]]
        n_fields = int((counted["kind"] == "field").sum())
        n_track_units = int((counted["kind"] == "track").sum())
        if n_fields + n_track_units < min_fields:
            continue
        surf = counted[counted["kind"] == "field"]["surface"].map(surface_class)
        lat, lon = g["lat"].mean(), g["lon"].mean()
        sports = sorted({s for v in g["sport"] for s in str(v).split(";") if s})
        best, best_d = None, NAME_MATCH_M
        if len(plat):
            near = np.where((abs(plat - lat) < 0.01) & (abs(plon - lon) < 0.012))[0]
            for i in near:
                d = haversine_m(lat, lon, plat[i], plon[i])
                if d < best_d:
                    best, best_d = places[i], d
        rows.append({
            "cluster": cid, "lat": round(lat, 6), "lon": round(lon, 6),
            "name": best["name"] if best else "", "matched_kind": best["kind"] if best else "",
            "operator": best["operator"] if best else "",
            "osm_website": best["website"] if best else "", "osm_phone": best["phone"] if best else "",
            "osm_email": best["email"] if best else "",
            "units": n_fields + n_track_units,
            "grass": int((surf == "grass").sum()), "turf": int((surf == "turf").sum()),
            "unknown": int((surf == "unknown").sum()),
            "tracks": int((g["kind"] == "track").sum()),
            "sports": ", ".join(sports),
        })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------- 4. geocode
def geocode(df, cache_dir):
    path = os.path.join(cache_dir, "geocode.json")
    cache = load_json(path, {})
    todo = sum(1 for _, r in df.iterrows() if f"{r.lat},{r.lon}" not in cache)
    if todo:
        print(f"Looking up addresses for {todo} complexes (~1/sec)...")
    out = []
    for i, r in df.iterrows():
        key = f"{r.lat},{r.lon}"
        if key not in cache:
            try:
                resp = requests.get(NOMINATIM_URL, headers=HEADERS, timeout=30,
                                    params={"lat": r.lat, "lon": r.lon, "format": "jsonv2", "zoom": 17})
                cache[key] = resp.json().get("address", {})
            except Exception as e:
                print(f"  geocode failed for {key}: {e}")
                cache[key] = {}
            time.sleep(1.1)
            if len(cache) % 25 == 0:
                save_json(path, cache)
        a = cache[key]
        street = " ".join(x for x in (a.get("house_number", ""), a.get("road", "")) if x)
        out.append({"street": street,
                    "city": a.get("city") or a.get("town") or a.get("village") or a.get("hamlet", ""),
                    "county": a.get("county", ""), "zip": (a.get("postcode") or "")[:5]})
    save_json(path, cache)
    return pd.concat([df.reset_index(drop=True), pd.DataFrame(out)], axis=1)


# ---------------------------------------------------------------- 5. Google Places
def places_enrich(df, cache_dir, key, state, budget=900):
    path = os.path.join(cache_dir, "places.json")
    cache = load_json(path, {})
    mask = ("places.displayName,places.formattedAddress,places.nationalPhoneNumber,"
            "places.websiteUri,places.types,places.businessStatus,places.location,"
            "places.addressComponents")
    todo = sum(1 for _, r in df.iterrows() if f"{r.lat},{r.lon}" not in cache)
    allowed = min(todo, budget) if key else 0
    if todo and allowed:
        print(f"Asking Google Places about {allowed} complexes (budget {budget}; {todo} not looked up yet)...")
    elif todo:
        print(f"Google Places: reusing saved answers only; {todo} complexes have none yet (no new calls, no charges).")
    calls = 0
    res = []
    for _, r in df.iterrows():
        ck = f"{r.lat},{r.lon}"
        if ck not in cache and calls >= allowed:
            res.append({})
            continue
        if ck not in cache:
            calls += 1
            query = r["name"] or f"sports complex near {r.get('city', '') or ''} {state}".strip()
            body = {"textQuery": query, "pageSize": 3,
                    "locationBias": {"circle": {"center": {"latitude": r.lat, "longitude": r.lon},
                                                "radius": 1000.0}}}
            try:
                resp = requests.post(PLACES_URL, json=body, timeout=30,
                                     headers={"X-Goog-Api-Key": key, "X-Goog-FieldMask": mask})
                if resp.status_code != 200:
                    print(f"  Places error {resp.status_code}: {resp.text[:200]}")
                    if resp.status_code in (401, 403):
                        raise SystemExit("Google rejected the API key. Check it and that Places API (New) is enabled.")
                cache[ck] = resp.json().get("places", [])
            except SystemExit:
                raise
            except Exception as e:
                print(f"  Places lookup failed: {e}")
                cache[ck] = []
            if len(cache) % 25 == 0:
                save_json(path, cache)
        best = None
        for p in cache[ck]:
            loc = p.get("location", {})
            if loc and haversine_m(r.lat, r.lon, loc["latitude"], loc["longitude"]) <= PLACES_MATCH_M:
                best = p
                break
        if best:
            zipc = next((c.get("shortText", "") for c in best.get("addressComponents", [])
                         if "postal_code" in c.get("types", [])), "")
            res.append({"g_name": best.get("displayName", {}).get("text", ""),
                        "g_address": best.get("formattedAddress", ""),
                        "g_phone": best.get("nationalPhoneNumber", ""),
                        "g_website": best.get("websiteUri", ""),
                        "g_types": ",".join(best.get("types", [])),
                        "g_status": best.get("businessStatus", ""), "g_zip": zipc})
        else:
            res.append({})
    save_json(path, cache)
    return pd.concat([df.reset_index(drop=True), pd.DataFrame(res)], axis=1)


# ---------------------------------------------------------------- 6. classify owner
PUBLIC_WORDS = re.compile(r"\b(county|city of|town of|village of|parks|recreation|government|"
                          r"department|municipal|authority|township|state of)\b", re.I)
SCHOOL_WORDS = re.compile(r"\b(school|high|middle|elementary|academy|college|university|"
                          r"isd|schools|board of education)\b", re.I)
PRIVATE_HINT = re.compile(r"\b(soccer|sports|athletic|baseball|softball|fc|sc|club|academy|"
                          r"llc|inc|complex|dome|plex)\b", re.I)


def classify(r):
    name = clean(r.get("name")) or clean(r.get("g_name"))
    op = clean(r.get("operator"))
    site = domain_of(clean(r.get("g_website")) or clean(r.get("osm_website")))
    types = clean(r.get("g_types"))
    if r.get("matched_kind") == "school" or "school" in types or "university" in types \
            or SCHOOL_WORDS.search(name) and not re.search(r"\bpark\b", name, re.I):
        return "School"
    if PUBLIC_WORDS.search(op) or site.endswith(".gov") or site.endswith(".us") \
            or re.search(r"(county|cityof|parks)", site):
        return "Public"
    if "park" in types.split(",") or (re.search(r"\bpark\b", name, re.I) and not PRIVATE_HINT.search(name)):
        return "Public"
    if not name:
        return "Unknown"
    return "Private"


# ---------------------------------------------------------------- 7. scrape emails
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
JUNK = ("example.", "sentry", "wixpress", ".png", ".jpg", ".gif", ".webp", "godaddy", "domain.com",
        "yourname", "email.com", "squarespace")


def scrape_emails(url, cache):
    if not url:
        return []
    if url in cache:
        return cache[url]
    found, pages = [], [url]
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        html = r.text
        for href in re.findall(r'href=["\']([^"\']+)["\']', html, re.I):
            if re.search(r"contact|about|staff|directory|team", href, re.I) and len(pages) < 4:
                full = urljoin(url, href)
                if domain_of(full) == domain_of(url) and full not in pages:
                    pages.append(full)
        texts = [html]
        for p in pages[1:]:
            try:
                texts.append(requests.get(p, headers={"User-Agent": "Mozilla/5.0"}, timeout=15).text)
            except Exception:
                pass
        for t in texts:
            for e in EMAIL_RE.findall(t.replace("&#64;", "@").replace("[at]", "@")):
                e = e.lower().strip(".")
                if not any(j in e for j in JUNK) and e not in found:
                    found.append(e)
    except Exception:
        pass
    cache[url] = found[:5]
    return cache[url]


def pick_email(emails, site_domain):
    if not emails:
        return ""
    same = [e for e in emails if site_domain and e.endswith(site_domain)]
    pool = same or emails
    pool.sort(key=lambda e: (e.split("@")[0] in ("info", "contact", "hello", "office", "admin"), len(e)))
    return pool[0]


# ---------------------------------------------------------------- 7b. shared contact tools (same as stadium finder v3)
BAD_EMAIL_PREFIX = r"^(schoolblocks|noreply-|do-not-reply|reporta|admissions?|privacy|parking|webmaster|web|esplost|splost|noreply|no-reply|donotreply|hr|jobs|careers|employment|registrar|financialaid|finaid|alumni|giving|development|communications|media|news|press|support|help|helpdesk|icmshelp|it|techsupport|accessibility|titleix|title9|compliance|records|transcripts|enrollment|sraoffice|courtinfo|ticketing|tickets|billing|accounts|payroll|purchasing|marketing)@"


def domain(url):
    if not url:
        return ""
    if "://" not in url:
        url = "http://" + url
    d = urlparse(url).netloc.lower()
    return d[4:] if d.startswith("www.") else d


EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")


AD_TITLE = re.compile(r"\b(athletic director|director of athletics|athletics director|athletic coordinator|"
                      r"director of athletics and activities|athletic administrator)\b", re.I)


ASSISTANT_AD = re.compile(r"\b(assistant|asst\.?|associate)\s+(athletic director|director of athletics|ad)\b", re.I)


PARKS_TITLE = re.compile(r"\b(parks?\s*(and|&)\s*rec(reation)?\s+director|director of parks|parks director|"
                         r"recreation director|director of recreation|parks and recreation director|"
                         r"leisure services director|athletics? (manager|supervisor|coordinator|superintendent)|"
                         r"recreation superintendent|parks superintendent|sports (manager|coordinator|supervisor)|"
                         r"director of community services|community services director|leisure services (director|manager)|"
                         r"recreation manager|parks manager|parks (and|&) facilities director)\b", re.I)
ASSISTANT_PARKS = re.compile(r"\b(assistant|asst\.?|deputy)\s+(director|superintendent|manager)\b", re.I)
PARKS_LINK_WORDS = re.compile(r"park|recreation|leisure|athlet|sports|staff|directory|contact|about|administration", re.I)
PARKS_FOCUS_WORDS = re.compile(r"park|recreation|leisure|athlet|sports|staff-directory|staff directory", re.I)
PROFILES = {}   # filled after the regexes below exist
PAGE_LINK_WORDS = re.compile(r"athlet|sports|staff|directory|faculty|coaches|contact|administration|about", re.I)


ATHLETIC_LINK_WORDS = re.compile(r"athlet|sports|coaches|staff-directory|staff directory", re.I)


SKIP_LINK = re.compile(r"facebook|twitter|x\.com|instagram|youtube|linkedin|tiktok|google\.|apple\.com|"
                       r"\.pdf$|\.jpg$|\.png$|\.docx?$|calendar|login|signin|javascript:|tel:", re.I)


NAME_RE = re.compile(r"(?:(?:Mr|Mrs|Ms|Miss|Dr|Coach)\.?\s+)?([A-Z][a-zA-Z'\-]+(?:\s+[A-Z]\.?)?\s+[A-Z][a-zA-Z'\-]+"
                     r"(?:\s+(?:Jr|Sr|II|III|IV)\.?)?)")


NOT_NAME_WORDS = {"fee", "fees", "accommodations", "accommodation", "go", "in", "this", "that", "war", "angry",
                  "your", "our", "with", "from", "for", "more", "click", "here", "welcome", "customer", "service", "search", "results", "returned", "visit", "executive", "associate", "vice",
                  "chair", "chairman", "teacher", "assessment", "technician", "local", "weather", "public", "utilities",
                  "rd", "road", "street", "ave", "avenue", "hwy", "highway", "commercial", "real", "report", "abuse",
                  "communication", "communications", "specialist", "senior", "custom", "university", "leadership",
                  "auxiliary", "media", "ela", "pe", "luther", "king", "secretary", "clerk", "nurse", "counselor",
                  "officer", "commissioner", "treasurer", "admin", "administrator", "webmaster", "support", "help",
                  "desk", "member", "members", "general", "main", "front", "reception", "receptionist", "calendar",
                  "privacy", "policy", "terms", "login", "portal", "online", "payments", "registration", "register",
                  "donate", "sponsor", "sponsors", "gallery", "photos", "video", "news", "alerts", "emergency",
                  "north", "south", "east", "west", "central", "memorial", "national", "state", "georgia", "atlanta",
                  "parks", "park", "recreation", "leisure", "services", "city", "government", "superintendent",
                  "manager", "supervisor", "commission", "commissioners", "mayor", "council", "hours", "facility",
                  "rentals", "programs", "events", "complex", "cultural", "affairs",  "park", "recreation", "leisure", "services", "city", "government", "superintendent",
                  "manager", "supervisor", "commission", "commissioners", "mayor", "council", "hours", "facility",
                  "rentals", "programs", "events", "complex", "cultural", "affairs", "athletic", "athletics", "director", "high", "school", "assistant", "department", "contact", "email",
                  "phone", "office", "football", "county", "staff", "home", "principal", "coach", "head", "the",
                  "sports", "coordinator", "of", "and", "booster", "club", "information", "directory", "news",
                  "schedule", "schedules", "tickets", "team", "teams", "varsity", "middle", "academy", "center",
                  "administration", "district", "board", "education", "welcome", "our", "meet", "for", "about",
                  "facilities", "stadium", "field", "activities", "georgia", "ghsa", "trainer", "athletes", "parent",
                  "parents", "student", "students", "forms", "physical", "click", "here", "read", "more", "view",
                  "all", "lady", "boys", "girls", "basketball", "baseball", "softball", "soccer", "track", "cross",
                  "country", "wrestling", "volleyball", "tennis", "golf", "swimming", "lacrosse", "cheer", "fax"}


EMAIL_PATTERNS = ["first.last", "flast", "firstlast", "first_last", "f.last", "lastf", "firstl", "last.first",
                  "lastfirst", "first.l", "first-last"]


PROFILES.update({
    "ad": {"title": AD_TITLE, "exclude": ASSISTANT_AD, "links": PAGE_LINK_WORDS, "focus": ATHLETIC_LINK_WORDS,
           "label": "Athletic Director"},
    "parks": {"title": PARKS_TITLE, "exclude": ASSISTANT_PARKS, "links": PARKS_LINK_WORDS, "focus": PARKS_FOCUS_WORDS,
              "label": "Parks & Recreation Director"},
})


def _clean_name_part(s):
    return re.sub(r"[^a-z]", "", s.lower())


def split_name(full):
    parts = [p for p in re.sub(r"\b(Mr|Mrs|Ms|Miss|Dr|Coach)\.?\s+", "", full).replace(",", " ").split()
             if not re.fullmatch(r"(Jr|Sr|II|III|IV)\.?|[A-Z]\.?", p)]
    if len(parts) < 2:
        return "", ""
    return _clean_name_part(parts[0]), _clean_name_part(parts[-1])


def make_local(pattern, first, last):
    if not first or not last:
        return ""
    return {"first.last": f"{first}.{last}", "flast": f"{first[0]}{last}", "firstlast": f"{first}{last}",
            "first_last": f"{first}_{last}", "f.last": f"{first[0]}.{last}", "lastf": f"{last}{first[0]}",
            "firstl": f"{first}{last[0]}", "last.first": f"{last}.{first}", "lastfirst": f"{last}{first}",
            "first.l": f"{first}.{last[0]}", "first-last": f"{first}-{last}"}[pattern]


def detect_pattern(full_name, email):
    first, last = split_name(full_name)
    local = email.split("@")[0].lower()
    local = re.sub(r"\d+$", "", local)          # jsmith2 -> jsmith
    for pat in EMAIL_PATTERNS:
        if make_local(pat, first, last) == local:
            return pat
    return ""


try:
    import gender_guesser.detector as _gg
    _FIRST = _gg.Detector(case_sensitive=False)
except Exception:          # optional: pip install gender-guesser
    _FIRST = None


def is_real_name(name, email=""):
    """Guard against page text read as a name ("Search Results", "PE Teacher", "Riverbend Rd").
    Accept when the first word is a known first name, or when an email on the page matches the name."""
    words = [w.strip(".").lower() for w in re.sub(r"\b(Mr|Mrs|Ms|Miss|Dr|Coach)\.?\s+", "", name).split()]
    words = [w for w in words if w not in ("jr", "sr", "ii", "iii", "iv")]
    if len(words) < 2 or any(w in NOT_NAME_WORDS for w in words):
        return False
    if email and detect_pattern(name, email):
        return True
    if _FIRST is not None:
        return _FIRST.get_gender(words[0]) != "unknown" and len(words[-1]) > 1
    return len(words[0]) > 2 and len(words[-1]) > 2   # without the name list: weaker check


def looks_like_name(s):
    s = s.strip(" :-–—|,\t")
    m = NAME_RE.fullmatch(s) or NAME_RE.search(s)
    if not m:
        return ""
    name = m.group(1).strip()
    words = [w.lower().strip(".") for w in name.split()]
    if any(w in NOT_NAME_WORDS for w in words) or len(name) > 40:
        return ""
    return name


def html_to_lines(text):
    text = unhide_emails(text)
    text = re.sub(r"(?is)<(script|style|noscript|svg)[^>]*>.*?</\1>", " ", text)
    text = re.sub(r'(?i)<a[^>]+href="mailto:([^"?]+)[^"]*"[^>]*>(.*?)</a>', r" \2 \1 ", text)
    text = re.sub(r"(?i)<(br|/p|/div|/li|/tr|/td|/th|/h\d|p|div|li|tr|h\d)[^>]*>", "\n", text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = _html.unescape(text)
    return [re.sub(r"\s+", " ", ln).strip() for ln in text.split("\n") if ln.strip()]


def fetch(url, session):
    try:
        r = session.get(url, headers={"User-Agent": USER_AGENT}, timeout=10, allow_redirects=True)
        if r.status_code == 200 and "html" in r.headers.get("content-type", "html"):
            return r.url, r.text
    except Exception:
        pass
    return url, ""


def links_from(base, text, same_site_only, prof=None):
    prof = prof or PROFILES["ad"]
    out = []
    base_reg = reg_domain(domain(base))
    for m in re.finditer(r'(?is)<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', text):
        href, label = m.group(1), re.sub(r"<[^>]+>", " ", m.group(2))
        if SKIP_LINK.search(href):
            continue
        url = urljoin(base, href)
        if not url.startswith("http"):
            continue
        same = reg_domain(domain(url)) == base_reg
        key = href + " " + label
        if same and prof["links"].search(key):
            out.append((0 if prof["focus"].search(key) else 1, url))
        elif not same and not same_site_only and prof["focus"].search(key):
            out.append((0, url))                     # separate athletics site (e.g. schoolnameathletics.com)
    seen, res = set(), []
    for _, u in sorted(out):
        u = u.split("#")[0].rstrip("/")
        if u not in seen:
            seen.add(u)
            res.append(u)
    return res


def extract_from_lines(lines, prof=None):
    prof = prof or PROFILES["ad"]
    TITLE, EXCL = prof["title"], prof["exclude"]
    """Returns (ad_candidates [(name, email_or_'')], name_email_pairs [(name, email)], emails)."""
    ads, pairs = [], []
    emails = []
    for i, ln in enumerate(lines):
        for e in EMAIL_RE.findall(ln):
            emails.append(e.lower())
    # name/email pairs anywhere on the page (for learning district email formats)
    for i, ln in enumerate(lines):
        for e in EMAIL_RE.findall(ln):
            window = lines[max(0, i - 2): i + 3]
            for w in window:
                for m in NAME_RE.finditer(w):
                    nm = looks_like_name(m.group(0))
                    if nm and detect_pattern(nm, e):
                        pairs.append((nm, e.lower()))
    for i, ln in enumerate(lines):
        if not TITLE.search(ln) or EXCL.search(ln) or len(ln) > 160:
            continue
        name = ""
        rest = TITLE.sub("|", ln)
        for chunk in re.split(r"[|:–—\-,/()]", rest):
            name = looks_like_name(chunk)
            if name:
                break
        if not name:
            for j in (i - 1, i + 1, i - 2, i + 2):
                if 0 <= j < len(lines) and len(lines[j]) <= 45 and not TITLE.search(lines[j]):
                    name = looks_like_name(lines[j])
                    if name:
                        break
        if not name:
            continue
        email = ""
        window = lines[max(0, i - 3): i + 4]
        near = [e.lower() for w in window for e in EMAIL_RE.findall(w)]
        for e in near:
            if detect_pattern(name, e):
                email = e
                break
        if not email:
            last = split_name(name)[1]
            for e in near:
                if last and last in e.split("@")[0].lower():
                    email = e
                    break
        ads.append((name, email))
    return ads, pairs, emails


def crawl_site(url, kind="ad"):
    """Look through a site for the person holding the profile's title (AD or parks director). ~12 page loads max."""
    prof = PROFILES[kind]
    session = requests.Session()
    result = {"ads": [], "pairs": [], "emails": [], "pages": 0}
    home_url, home = fetch(url, session)
    if not home:
        return result
    queue = links_from(home_url, home, same_site_only=False, prof=prof)[:8]
    pages = [(home_url, home)]
    visited = {home_url.rstrip("/")}
    second_level = []
    for u in queue:
        if u in visited:
            continue
        visited.add(u)
        fu, t = fetch(u, session)
        if t:
            pages.append((fu, t))
            if prof["focus"].search(u):
                second_level += [x for x in links_from(fu, t, same_site_only=True, prof=prof)
                                 if re.search(r"staff|directory|coaches|contact|administration", x, re.I)][:3]
    for u in second_level[:4]:
        if u in visited:
            continue
        visited.add(u)
        fu, t = fetch(u, session)
        if t:
            pages.append((fu, t))
    for _, t in pages:
        ads, pairs, emails = extract_from_lines(html_to_lines(t), prof)
        result["ads"] += ads
        result["pairs"] += pairs
        result["emails"] += emails
    result["pages"] = len(pages)
    return result


def learn_patterns(site_results, extra_pairs):
    """Email format per email domain, from real name+email examples. Needs 2+ agreeing examples."""
    votes = {}
    seen = set()
    for name, email in [p for res in site_results.values() for p in res.get("pairs", [])
                        if is_real_name(p[0], p[1])] + extra_pairs:
        key = (name.lower(), email.lower())
        if key in seen:
            continue
        seen.add(key)
        pat = detect_pattern(name, email)
        if pat:
            votes.setdefault(email.split("@")[-1].lower(), {}).setdefault(pat, 0)
            votes[email.split("@")[-1].lower()][pat] += 1
    patterns = {}
    for dom, v in votes.items():
        pat, n = max(v.items(), key=lambda x: x[1])
        if n >= 2 and n / sum(v.values()) >= 0.7:
            patterns[dom] = (pat, n)
    return patterns


def site_email_domain(site_url, res, known_map):
    """Which domain the school's staff emails use (website and email domains often differ)."""
    counts = {}
    for e in res.get("emails", []):
        d = e.split("@")[-1]
        if re.search(PERSONAL_EMAIL, e) or re.match(BAD_EMAIL_PREFIX, e):
            continue
        counts[d] = counts.get(d, 0) + 1
    for d, _ in sorted(counts.items(), key=lambda x: -x[1]):
        if re.search(DISTRICT_DOMAIN + r"|\.us$|\.org$|\.net$|\.com$", d):
            return d
    return known_map.get(reg_domain(domain(site_url)), "")


def outside_pairs():
    """Name+email examples from HubSpot and Seamless exports, plus website->email domain hints."""
    pairs, dom_map = [], {}
    for folder in ("hubspot", "seamless"):
        for fpath in glob.glob(os.path.join(folder, "*.csv")):
            try:
                df = pd.read_csv(fpath, dtype=str).fillna("")
            except Exception:
                continue
            for rec in df.to_dict("records"):
                name = pick(rec, "contact full name", "full name") or \
                    " ".join(x for x in (pick(rec, "first name"), pick(rec, "last name")) if x)
                for col in ("email 1", "email", "contact email", "email 2", "work email"):
                    e = pick(rec, col).lower()
                    if e and "@" in e and not re.search(PERSONAL_EMAIL, e):
                        if name:
                            pairs.append((name, e))
                        site = domain(pick(rec, "company website domain", "company domain name", "website", "website url"))
                        if site:
                            dom_map.setdefault(reg_domain(site), e.split("@")[-1])
    return pairs, dom_map


def pick(d, *names):
    low = {k.lower().strip(): k for k in d}
    for n in names:
        if n in low and d[low[n]]:
            return str(d[low[n]]).strip()
    return ""


PERSONAL_EMAIL = r"@(gmail|yahoo|hotmail|outlook|aol|icloud|comcast|bellsouth|att|live|msn|me)\.(com|net)$"


def title_score(title):
    t0 = (title or "").lower()
    if re.search(r"\b(owner|general manager|executive director|facility director|facilities director|"
                 r"facility manager|director of operations|president|ceo)\b", t0) and not re.search(r"assistant|vice", t0):
        return 5
    """How likely this person decides on field-painting equipment (higher is better)."""
    t = (title or "").lower()
    if re.search(r"athletic director|director of athletics|\bad\b|athletics director", t):
        return 5
    if re.search(r"turf|groundskeep|grounds|field operations|stadium operations|athletic facilities|facilities", t):
        return 5
    if re.search(r"park services|parks director|recreation director|director of parks|athletics manager|"
                 r"parks (and|&) rec|recreation superintendent|parks superintendent|leisure services", t):
        return 5
    if re.search(r"head (football )?coach|head football", t):
        return 3
    if re.search(r"recreation|parks|operations|athletic", t) and not re.search(r"trainer|training", t):
        return 2
    return 1


def pick_seamless_email(rec, company_domain):
    """Best email on a Seamless row: (email, quality) or ("", "")."""
    options = []
    for col, vcol in (("email 1", "email 1 validation"), ("email 2", "email 2 validation"),
                      ("contact email", "contact email validation"), ("email", ""), ("work email", "")):
        e = pick(rec, col).strip()
        if not e or "@" not in e:
            continue
        v = pick(rec, vcol).lower() if vcol else ""
        edom = e.split("@")[-1].lower()
        personal = bool(re.search(PERSONAL_EMAIL, e.lower()))
        related = (company_domain and reg_domain(edom) == reg_domain(company_domain)) or \
            re.search(DISTRICT_DOMAIN + r"|\.edu$|\.gov$|\.us$", edom)
        if not personal and not related:
            continue           # e.g. an "athletic director" with a real-estate company email
        rank = (0 if personal else 2) + (2 if v == "valid" else 1 if v in ("accept all", "catch all", "catch-all") else 0)
        quality = ("Personal email" if personal else "Work email") + (f" ({v})" if v else "")
        options.append((rank, e, quality))
    if not options:
        return "", ""
    options.sort(key=lambda x: -x[0])
    return options[0][1], options[0][2]


STATE_ALIASES = {"ga": "GA", "georgia": "GA", "fl": "FL", "florida": "FL", "al": "AL", "alabama": "AL",
                 "sc": "SC", "south carolina": "SC", "nc": "NC", "north carolina": "NC", "tn": "TN", "tennessee": "TN"}


NAME_STOPWORDS = {"ga", "government", "gov", "high", "school", "schools", "county", "stadium", "field", "athletics", "athletic", "university",
                  "college", "academy", "the", "inc", "middle", "public", "system", "district", "of", "and", "at",
                  "football", "track", "complex", "department", "dept", "sports", "hs", "ms", "city"}


PARENT_WORDS = r"district|public schools|county schools|city schools|school system|board of education|\bboe\b|county government|\bcounty\b$|parks|recreation|university system"


DISTRICT_DOMAIN = r"k12|schools|\bboe\b|boe\.|district|\.gov$|county|cityschools"


TYPE_WORDS = [("high", r"\bhigh\b|\bhs\b"), ("middle", r"\bmiddle\b|\bms\b"), ("elementary", r"elementary|primary"),
              ("college", r"university|college(?! (and|&) career)|seminary")]


def org_type(name):
    n = name.lower()
    return {t for t, pat in TYPE_WORDS if re.search(pat, n)}


def is_parent_name(label):
    """HubSpot record is a district, county, city or whole university rather than one school/department."""
    n = label.lower()
    return bool(re.search(PARENT_WORDS, n)) or bool(re.search(r"\bschools$|school system|^the university of|county$", n.strip()))


def norm_match(s):
    s = str(s or "").lower()
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    s = re.sub(r"\b(the|inc|llc|of|and|at)\b", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def name_key(s):
    """Distinctive part of a name: 'Warren County High School' -> 'warren'."""
    return " ".join(w for w in norm_match(s).split() if w not in NAME_STOPWORDS)


def reg_domain(d):
    """Registrable domain: graysonhs.gcpsk12.org -> gcpsk12.org, x.cobb.k12.ga.us -> cobb.k12.ga.us."""
    parts = d.split(".")
    if len(parts) >= 4 and parts[-3] == "k12":
        return ".".join(parts[-4:])
    return ".".join(parts[-2:]) if len(parts) >= 2 else d


# ---------------------------------------------------------------- 8. agencies table
AGENCY_COLS = ["Agency Key", "County", "City", "OSM Operator", "# Complexes",
               "Department Name", "Contact Name", "Contact Title", "Contact Email",
               "Contact Phone", "Department Website"]


def agency_key(r):
    op = clean(r.get("operator"))
    if op:
        return op
    city, county = clean(r.get("city")), clean(r.get("county"))
    return f"{city}, {county}" if city else county or "Unknown area"


def sync_agencies(df, state):
    os.makedirs("agencies", exist_ok=True)
    path = os.path.join("agencies", f"agencies_{state}.csv")
    existing = pd.read_csv(path, dtype=str).fillna("") if os.path.exists(path) \
        else pd.DataFrame(columns=AGENCY_COLS)
    pub = df[df["owner_type"] == "Public"]
    counts = pub.groupby("agency_key").size()
    rows = {r["Agency Key"]: dict(r) for _, r in existing.iterrows()}
    for r in rows.values():
        r["# Complexes"] = "0"   # recounted below; filled-in rows are kept even if unused
    for k, n in counts.items():
        sample = pub[pub["agency_key"] == k].iloc[0]
        if k not in rows:
            rows[k] = {c: "" for c in AGENCY_COLS}
            rows[k].update({"Agency Key": k, "County": clean(sample.get("county")),
                            "City": clean(sample.get("city")), "OSM Operator": clean(sample.get("operator"))})
            site = clean(sample.get("website"))
            if re.search(r"\.(gov|us)$", domain_of(site)):
                rows[k]["Department Website"] = site   # pre-fill a likely government site
        rows[k]["# Complexes"] = str(n)
    table = pd.DataFrame(rows.values(), columns=AGENCY_COLS)
    table["_n"] = pd.to_numeric(table["# Complexes"], errors="coerce").fillna(0)
    table = table.sort_values("_n", ascending=False).drop(columns="_n")
    table.to_csv(path, index=False)
    filled = (table["Contact Email"].str.strip() != "").sum()
    print(f"Agencies table: {len(table)} departments, {filled} with a contact email -> {path}")
    return table


# ---------------------------------------------------------------- 9. Seamless in/out
def load_seamless(state):
    """People from Seamless 'My Contacts' exports. Upload lists and company-only rows are ignored."""
    people, rejected = [], []
    for fpath in glob.glob(os.path.join("seamless", "*.csv")):
        if os.path.basename(fpath).lower().startswith("seamless_upload"):
            continue
        try:
            s = pd.read_csv(fpath, dtype=str).fillna("")
        except Exception as e:
            print(f"  Could not read {fpath}: {e}")
            continue
        for rec in s.to_dict("records"):
            name = pick(rec, "contact full name", "full name", "name") or \
                " ".join(x for x in (pick(rec, "first name"), pick(rec, "last name")) if x)
            if not name.strip():
                continue
            cdom = domain(pick(rec, "company website domain", "company domain", "website", "company website", "domain"))
            p = {"company": pick(rec, "company name - cleaned", "company name", "company", "organization"),
                 "domain": cdom, "name": name.strip(), "title": pick(rec, "title", "job title"),
                 "phone": pick(rec, "contact phone 1", "contact mobile phone", "direct phone", "phone", "phone 1",
                               "work phone", "mobile phone"),
                 "state": pick(rec, "company state abbr", "company state", "contact state abbr", "state").strip()}
            p["email"], p["email_quality"] = pick_seamless_email(rec, cdom)
            st = STATE_ALIASES.get(p["state"].lower(), p["state"].upper())
            if st and st != state:
                rejected.append({**p, "reason": f"company is in {st}, not {state}"})
                continue
            p["score"] = title_score(p["title"])
            people.append(p)
    if people:
        print(f"Loaded {len(people)} people from Seamless exports.")
    return people, rejected


PARKS_ROLE = re.compile(r"park|recreation|leisure|athletic|sports|facilit|grounds|turf|field", re.I)


def seamless_for(r, people, reg_count):
    """Seamless people who belong to this lead's buyer."""
    d = domain(r["website"])
    rd = reg_domain(d) if d else ""
    public = r["owner_type"] == "Public"
    names = [r["buyer"], r["company"]] + ([r.get("agency_dept", ""), r.get("agency_key", "")] if public else [])
    school_lead = r["owner_type"] == "School"
    # agency keys look like "County" (county-run) or "City, County" (city-run): match staff to the right one
    if public:
        places = [r.get("city", "")] if "," in (r.get("agency_key") or "") else [r.get("county", "")]
    else:
        places = []
    out = []
    for p in people:
        ok = False
        if not school_lead and org_type(p["company"]) & {"high", "middle", "elementary"}:
            continue          # a school's staff isn't the buyer for a park or private complex
        if d and p["domain"]:
            prd = reg_domain(p["domain"])
            if d == p["domain"]:
                ok = True
            elif rd == prd and (reg_count.get(rd, 0) <= 1 or (public and PARKS_ROLE.search(p["title"]))) \
                    and not re.search(r"k12|schools", rd):
                ok = True
        if not ok and r.get("agency_site") and p["domain"] and public and PARKS_ROLE.search(p["title"]):
            ok = reg_domain(domain(r["agency_site"])) == reg_domain(p["domain"])
        pk, pt = name_key(p["company"]), org_type(p["company"])
        if not ok and len(pk) >= 4:
            for c in names:
                ck, ct = name_key(c), org_type(c)
                if len(ck) >= 4 and (not pt or not ct or pt <= ct or ct <= pt) and \
                        (pk == ck or (len(pk) >= 8 and difflib.SequenceMatcher(None, pk, ck).ratio() >= 0.92)):
                    ok = True
                    break
        if not ok and public and len(pk) >= 4 and not org_type(p["company"]) \
                and re.search(r"park|recreation|leisure", p["title"], re.I):
            # "Chatham County, GA Government" + a parks title -> the Chatham County parks department
            ok = any(pk == name_key(x) for x in places if x)
        if ok:
            out.append(p)
    return out


def write_seamless_upload(df, agencies, state):
    path = f"seamless_upload_complexes_{state}.csv"
    rows, seen = [], set()
    need = df[(df["owner_type"].isin(["Private", "School"])) & (df["contact_email"] == "")
              & (df["company"] != "") & (~df["hubspot_status"].isin(["Already in HubSpot", "Likely duplicate"]))]
    for _, r in need.iterrows():
        if r["buyer"] in seen:
            continue
        seen.add(r["buyer"])
        rows.append({"Company Name": r["buyer"], "Company Domain": domain_of(r["website"]), "City": r["city"],
                     "State": state, "Target Titles": ("Athletic Director; Director of Athletics"
                                                      if r["owner_type"] == "School" else
                                                      "Owner; General Manager; Facility Director; Executive Director"),
                     "Lead IDs": r["lead_id"]})
    pub = df[(df["owner_type"] == "Public") & (df["contact_email"] == "")
             & (~df["hubspot_status"].isin(["Already in HubSpot", "Likely duplicate"]))]
    for _, r in pub.iterrows():
        if r["buyer"] in seen:
            continue
        seen.add(r["buyer"])
        rows.append({"Company Name": r["buyer"], "Company Domain": domain_of(r.get("agency_site", "")),
                     "City": r["city"], "State": state,
                     "Target Titles": "Parks Director; Recreation Director; Athletics Manager; Parks Superintendent",
                     "Lead IDs": r["lead_id"]})
    pd.DataFrame(rows, columns=["Company Name", "Company Domain", "City", "State", "Target Titles", "Lead IDs"]) \
        .to_csv(path, index=False)
    return path, len(rows)


# ---------------------------------------------------------------- 10. HubSpot check
def load_hubspot(state):
    doms, regs, emails, phones, names = {}, {}, {}, {}, {}
    files = glob.glob(os.path.join("hubspot", "*.csv"))
    for fpath in files:
        try:
            h = pd.read_csv(fpath, dtype=str).fillna("")
        except Exception as e:
            print(f"  Could not read {fpath}: {e}")
            continue
        for rec in h.to_dict("records"):
            company = pick(rec, "company name", "name", "associated company")
            label = company or " ".join(x for x in (pick(rec, "first name"), pick(rec, "last name")) if x)
            raw_st = pick(rec, "state/region", "state", "company state/region").lower().strip()
            st = STATE_ALIASES.get(raw_st, raw_st.upper() if raw_st else "")
            if st and st != state:
                continue
            city = pick(rec, "city", "company city").lower().strip()
            d = domain(pick(rec, "company domain name", "website url", "domain", "website"))
            e = pick(rec, "email").lower()
            for dd, lab in ((d, label), (e.split("@")[-1] if e else "", label + " (contact email domain)")):
                if dd:
                    doms.setdefault(dd, lab)
                    regs.setdefault(reg_domain(dd), lab)
            if e:
                emails.setdefault(e, label)
            ph = digits10(pick(rec, "phone number", "phone", "mobile phone number"))
            if ph:
                phones.setdefault(ph, label)
            k = name_key(company)
            if company and len(k) >= 4:
                names.setdefault(k, []).append((norm_match(company), city, label))
    return {"loaded": bool(files), "doms": doms, "regs": regs, "emails": emails, "phones": phones, "names": names}


GENERIC_DOMAINS = {"gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "aol.com", "icloud.com", "comcast.net",
                   "bellsouth.net", "att.net", "google.com", "facebook.com", "sites.google.com"}
PARKS_LABEL = re.compile(r"park|recreation|leisure", re.I)
GOV_LABEL = re.compile(r"\b(city|county|government|town|department|departments|consolidated|village|municipal|"
                       r"commission|augusta)\b", re.I)
ORG_WORDS = re.compile(r"\b(inc|llc|club|association|academy|school|church|ymca|soccer|baseball|football|sports|"
                       r"athletic|league|foundation|company|corp|group|university|college|county|city|park)\b", re.I)


def is_person_label(lab):
    """HubSpot contact with no company: label is just a person's name."""
    base = re.sub(r"\s*\(contact email domain\)", "", lab).strip()
    return bool(re.fullmatch(r"[A-Z][a-zA-Z'\-]+(\s+[A-Z][a-zA-Z'\-\.]+){1,2}", base)) and not ORG_WORDS.search(base)


def hubspot_status(r, hs, reg_count, phone_count):
    """
    Already in HubSpot     - same contact email, or the lead's own domain (for public parks: the parks dept's)
    Likely duplicate       - same distinctive name in the same city, or same phone as a single org/person
    Possible duplicate     - same distinctive name in this state
    Parent org in HubSpot  - HubSpot has the district, university or county government, not this buyer
    New
    """
    if not hs["loaded"]:
        return "Not checked", ""
    public = r["owner_type"] == "Public"
    d = domain(r["website"])
    rd = reg_domain(d) if d else ""
    e = (r["contact_email"] or "").lower()
    p = digits10(r["phone"])
    best = None
    for cand in (r["buyer"], r["company"]):
        k = name_key(cand)
        if len(k) < 4 or not cand:
            continue
        ctype = org_type(cand)
        for full, city, label in hs["names"].get(k, []):
            parent = is_parent_name(label) and not is_parent_name(cand)
            if not parent and difflib.SequenceMatcher(None, full, norm_match(cand)).ratio() < 0.6:
                continue
            htype = org_type(label)
            if not parent and ctype and htype and ctype != htype:
                continue
            same_city = bool(city) and city == (r["city"] or "").lower()
            score = (not parent, same_city)
            if best is None or score > (not best[2], best[1]):
                best = (label, same_city, parent)
    if e and e in hs["emails"]:
        return "Already in HubSpot", f"email: {hs['emails'][e]}"
    if d and d not in GENERIC_DOMAINS and d in hs["doms"]:
        lab = hs["doms"][d]
        if public:
            # only a parks-related record (or the park's own operator, e.g. a soccer association) is this buyer
            own = not is_parent_name(lab) and not GOV_LABEL.search(lab) and not is_person_label(lab)
            return ("Already in HubSpot", f"domain: {lab}") if PARKS_LABEL.search(lab) or own \
                else ("Parent org in HubSpot", f"domain {d}: {lab}")
        district_like = "(contact email domain)" in lab and re.search(DISTRICT_DOMAIN, d)
        if (reg_count.get(rd, 0) <= 1 or r["owner_type"] == "Private") and not is_parent_name(lab) and not district_like:
            return "Already in HubSpot", f"domain: {lab}"
    if best and not best[2] and best[1]:
        return "Likely duplicate", f"name + city: {best[0]}"
    if p and p in hs["phones"] and phone_count.get(p, 0) <= 1 and not is_parent_name(hs["phones"][p]):
        return "Likely duplicate", f"phone: {hs['phones'][p]}"
    if best and not best[2]:
        return "Possible duplicate", f"name: {best[0]}"
    if best:
        return "Parent org in HubSpot", f"name: {best[0]}"
    if rd and rd not in GENERIC_DOMAINS and rd in hs["regs"]:
        lab = hs["regs"][rd]
        if public and PARKS_LABEL.search(lab):
            return "Already in HubSpot", f"domain {rd}: {lab}"
        return "Parent org in HubSpot", f"domain {rd}: {lab}"
    if p and p in hs["phones"]:
        return "Parent org in HubSpot", f"shared phone: {hs['phones'][p]}"
    return "New", ""


# ---------------------------------------------------------------- 11. people on websites
def find_people(df, state, deep=False, deep_limit=0, use_browser=True):
    """Athletic directors on school sites; parks & rec directors on parks department sites.
    deep=True adds a second, deeper pass (browser, sitemaps, directories, search boxes, PDFs) on sites where
    the quick pass found nobody, plus private complex websites."""
    jobs = {}
    for _, r in df.iterrows():
        if r["owner_type"] == "School" and r["website"]:
            jobs.setdefault(("ad", r["website"]), None)
        elif r["owner_type"] == "Public":
            site = r.get("agency_site") or (r["website"] if re.search(r"\.(gov|us)$|county|cityof|parks|recreation",
                                                                      domain_of(r["website"])) else "")
            if site:
                jobs.setdefault(("parks", site), None)
        elif deep and r["owner_type"] in ("Private", "Unknown") and r["website"] \
                and not re.search(r"facebook|instagram|google\.|yelp", r["website"], re.I):
            jobs.setdefault(("private", r["website"]), None)
    order = {}
    for _, r in df.sort_values("score", ascending=False).iterrows():
        for k in ("ad", "parks", "private"):
            site = (r.get("agency_site") or r["website"]) if k == "parks" else r["website"]
            order.setdefault((k, site), len(order))
    results, n_new = {}, 0
    for kind in ("ad", "parks"):
        path = os.path.join("cache", f"{'ad_finder' if kind == 'ad' else 'parks_finder'}_{state}.json")
        cache = load_json(path, {})
        todo = [u for (k, u) in jobs if k == kind and u not in cache]
        if todo:
            label = "athletic directors on school" if kind == "ad" else "parks & rec directors on department"
            print(f"Looking for {label} websites: {len(todo)} sites (roughly {max(1, len(todo) // 25)} min)...")
            with ThreadPoolExecutor(max_workers=6) as ex:
                for n, (u, res) in enumerate(zip(todo, ex.map(lambda x: crawl_site(x, kind), todo)), 1):
                    cache[u] = res
                    if n % 25 == 0:
                        save_json(path, cache)
                        print(f"  {n}/{len(todo)} sites checked")
            save_json(path, cache)
            n_new += len(todo)
        for (k, u) in jobs:
            if k == kind and u in cache:
                results[(k, u)] = cache[u]
    deep_stats = {"tried": 0, "found": 0, "rendered": 0, "pdfs": 0}
    if deep:
        todo_all = [(k, u) for (k, u) in jobs if k == "private" or not results.get((k, u), {}).get("ads")]
        todo_all.sort(key=lambda x: order.get(x, 10 ** 9))
        if deep_limit:
            todo_all = todo_all[:deep_limit]
        caches = {k: load_json(os.path.join("cache", f"deep_finder_{k}_{state}.json"), {}) for k in ("ad", "parks", "private")}
        todo = [(k, u) for (k, u) in todo_all if u not in caches[k]]
        if todo:
            how = "with a browser" if use_browser else "without a browser"
            print(f"Deep search {how}: {len(todo)} sites (roughly {max(1, len(todo) // 8)} min)...")
            from concurrent.futures import as_completed
            with ThreadPoolExecutor(max_workers=4) as ex:
                futs = {ex.submit(crawl_site_deep, u, k, use_browser): (k, u) for (k, u) in todo}
                for n, f in enumerate(as_completed(futs), 1):
                    k, u = futs[f]
                    try:
                        caches[k][u] = f.result()
                    except Exception as e:
                        caches[k][u] = {"ads": [], "pairs": [], "emails": [], "error": str(e)[:200]}
                    if n % 5 == 0 or n == len(todo):
                        for kk in caches:
                            save_json(os.path.join("cache", f"deep_finder_{kk}_{state}.json"), caches[kk])
                        print(f"  {n}/{len(todo)} sites searched (saved)")
            for kk in caches:
                save_json(os.path.join("cache", f"deep_finder_{kk}_{state}.json"), caches[kk])
        for (k, u) in todo_all:
            d = caches[k].get(u)
            if not d:
                continue
            deep_stats["tried"] += 1
            deep_stats["rendered"] += d.get("rendered", 0)
            deep_stats["pdfs"] += d.get("pdfs", 0)
            if any(is_real_name(n, e) for n, e in d.get("ads", [])):
                deep_stats["found"] += 1
            base = results.get((k, u), {"ads": [], "pairs": [], "emails": []})
            results[(k, u)] = {"ads": base["ads"] + d.get("ads", []), "pairs": base["pairs"] + d.get("pairs", []),
                               "emails": base["emails"] + d.get("emails", []), "deep": bool(d.get("ads"))}
    extra, dom_map = outside_pairs()
    patterns = learn_patterns({f"{k}|{u}": v for (k, u), v in results.items()}, extra)
    people = {}
    for (k, u), res in results.items():
        names = {}
        for nm, em in res.get("ads", []):
            key = split_name(nm)
            if not all(key) or not is_real_name(nm, em):
                continue
            cur = names.setdefault(key, {"name": nm, "count": 0, "email": ""})
            cur["count"] += 1
            if em and not cur["email"]:
                cur["email"] = em
        if not names:
            continue
        best = max(names.values(), key=lambda x: (x["count"], bool(x["email"])))
        person = {"name": best["name"], "title": PROFILES[k]["label"] + " (per website)",
                  "email": best["email"], "status": "On website" if best["email"] else "",
                  "deep": bool(res.get("deep"))}
        if not person["email"]:
            edom = site_email_domain(u, res, dom_map)
            if edom in patterns:
                pat, n = patterns[edom]
                first, last = split_name(best["name"])
                person["email"] = f"{make_local(pat, first, last)}@{edom}"
                person["status"] = f"Guessed - {pat}@{edom} format ({n} examples)"
        people[(k, u)] = person
    used_domains = {site_email_domain(u, res, dom_map) for (k, u), res in results.items()} - {""}
    people["_patterns"] = patterns
    people["_site_domains"] = {u: site_email_domain(u, res, dom_map) for (k, u), res in results.items()}
    return people, len([1 for (k, _) in jobs if k == "ad"]), len([1 for (k, _) in jobs if k == "parks"]), \
        len([p for p in used_domains if p in patterns]), deep_stats


# ---------------------------------------------------------------- 11b. deep website search (all free)
# Opens sites in a real browser when needed, reads sitemaps and common directory pages, uses the site's own
# search box, reads PDF staff lists, and decodes emails that sites hide from spam bots.
import urllib.robotparser
import io

PRIVATE_TITLE = re.compile(r"\b(owner|co-owner|general manager|executive director|facility director|facilities director|"
                           r"facility manager|facilities manager|director of operations|operations director|"
                           r"club president|president|director of coaching|club director|sports director|"
                           r"program director|complex manager|park manager)\b", re.I)
ASSISTANT_PRIVATE = re.compile(r"\b(assistant|asst\.?|associate|vice|deputy)\s+(general manager|president|director|manager)\b"
                               r"|\bvice[- ]president\b", re.I)
PROFILES["private"] = {
    "title": PRIVATE_TITLE, "exclude": ASSISTANT_PRIVATE,
    "links": re.compile(r"about|staff|team|leadership|board|contact|directors|who-we-are|coaches|management", re.I),
    "focus": re.compile(r"staff|team|leadership|board|about|directory|contact|management", re.I),
    "label": "Owner / Director",
}
DEEP_PATHS = {
    "parks": ["/directory.aspx", "/Directory.aspx", "/staff-directory", "/directory", "/departments/parks-and-recreation",
              "/parks-and-recreation", "/parks", "/recreation", "/parks-recreation"],
    "ad": ["/athletics", "/staff-directory", "/directory", "/staff", "/athletics/staff", "/athletic-department"],
    "private": ["/about", "/about-us", "/staff", "/our-staff", "/our-team", "/team", "/leadership", "/board",
                "/contact", "/contact-us"],
}
SEARCH_TERMS = {"parks": ["parks director", "recreation director"], "ad": ["athletic director"],
                "private": ["director"]}
URL_WORDS = {
    "parks": re.compile(r"park|recreation|leisure|staff|directory|contact|department|athlet", re.I),
    "ad": re.compile(r"athlet|sports|staff|directory|faculty|coach|administration", re.I),
    "private": re.compile(r"about|staff|team|leadership|board|contact|director|management|who", re.I),
}
JS_HINTS = re.compile(r"__NEXT_DATA__|ng-app|data-reactroot|id=\"root\"|id=\"app\"|Loading\.\.\.|please enable javascript",
                      re.I)
CF_EMAIL = re.compile(r'data-cfemail="([0-9a-fA-F]+)"')
CF_HREF = re.compile(r'/cdn-cgi/l/email-protection#([0-9a-fA-F]+)')
AT_DOT = re.compile(r"([A-Za-z0-9._%+-]+)\s*(?:\[at\]|\(at\)|\{at\}|\s+at\s+)\s*([A-Za-z0-9-]+(?:\s*(?:\[dot\]|\(dot\)|\{dot\}|\.)\s*[A-Za-z0-9-]+)*)"
                    r"\s*(?:\[dot\]|\(dot\)|\{dot\})\s*([A-Za-z]{2,6})\b", re.I)


SITE_TIME_LIMIT = 120        # seconds per site, whatever happens
REQUEST_TIME_LIMIT = 20      # seconds per page download


def get_capped(session, url, timeout=REQUEST_TIME_LIMIT, max_bytes=3_000_000):
    """requests.get with a hard total-time and size limit.
    Plain timeouts don't stop servers that keep dripping a byte at a time, so the download reads in small
    pieces, checks the clock after each one, and runs in a helper thread that is abandoned if it overruns."""
    import threading
    box = {}

    def work():
        start = time.time()
        try:
            r = session.get(url, headers={"User-Agent": USER_AGENT}, timeout=(6, 10), stream=True, allow_redirects=True)
            box["resp"] = r
            chunks, size = [], 0
            read1 = getattr(r.raw, "read1", None)
            while size < max_bytes and time.time() - start < timeout:
                piece = read1(65536, decode_content=True) if read1 else r.raw.read(1024, decode_content=True)
                if not piece:
                    break
                chunks.append(piece)
                size += len(piece)
            r._content = b"".join(chunks)
            r._content_consumed = True
            box["ok"] = r
        except Exception:
            pass

    th = threading.Thread(target=work, daemon=True)
    th.start()
    th.join(timeout + 12)
    if th.is_alive():
        try:
            box["resp"].close()
        except Exception:
            pass
        return None
    return box.get("ok")


def fetch_deep(url, session):
    r = get_capped(session, url)
    if r is not None and r.status_code == 200 and "html" in r.headers.get("content-type", "html"):
        try:
            return r.url, r.content.decode(r.encoding or "utf-8", errors="replace")
        except Exception:
            return r.url, ""
    return url, ""


def _decode_cf(hexstr):
    try:
        key = int(hexstr[:2], 16)
        return "".join(chr(int(hexstr[i:i + 2], 16) ^ key) for i in range(2, len(hexstr), 2))
    except ValueError:
        return ""


def unhide_emails(text):
    """Turn spam-protected emails back into plain addresses."""
    text = re.sub(r'<(a|span)[^>]*data-cfemail="([0-9a-fA-F]+)"[^>]*>.*?</\1>',
                  lambda m: " " + _decode_cf(m.group(2)) + " ", text, flags=re.S)
    text = CF_HREF.sub(lambda m: "mailto:" + _decode_cf(m.group(1)), text)

    def at_dot(m):
        dom = re.sub(r"\s*(\[dot\]|\(dot\)|\{dot\})\s*", ".", m.group(2), flags=re.I).replace(" ", "")
        return f"{m.group(1)}@{dom}.{m.group(3)}"
    return AT_DOT.sub(at_dot, text)


_robots = {}


def allowed(url):
    """Respect robots.txt."""
    base = "{0.scheme}://{0.netloc}".format(urlparse(url))
    if base not in _robots:
        rp = urllib.robotparser.RobotFileParser()
        r = get_capped(requests.Session(), base + "/robots.txt", timeout=10, max_bytes=500_000)
        try:
            rp.parse(r.text.splitlines() if r is not None and r.status_code == 200 else [])
        except Exception:
            rp.parse([])
        _robots[base] = rp
    try:
        return _robots[base].can_fetch(USER_AGENT, url)
    except Exception:
        return True


def sitemap_urls(base, session, limit=3000):
    """Page addresses listed in the site's sitemap(s)."""
    root = "{0.scheme}://{0.netloc}".format(urlparse(base))
    maps, seen, urls = [], set(), []
    r = get_capped(session, root + "/robots.txt", timeout=10, max_bytes=500_000)
    if r is not None and r.status_code == 200:
        maps += re.findall(r"(?im)^sitemap:\s*(\S+)", r.text)
    maps += [root + "/sitemap.xml", root + "/sitemap_index.xml"]
    while maps and len(seen) < 6 and len(urls) < limit:
        m = maps.pop(0)
        if m in seen:
            continue
        seen.add(m)
        r = get_capped(session, m, timeout=15, max_bytes=5_000_000)
        if r is None or r.status_code != 200 or "<loc>" not in r.text:
            continue
        for loc in re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", r.text):
            loc = _html.unescape(loc)
            (maps if loc.lower().endswith(".xml") else urls).append(loc)
    return urls[:limit]


def pdf_lines(url, session):
    try:
        import pdfplumber
    except ImportError:
        return []
    try:
        r = get_capped(session, url, timeout=25, max_bytes=8_000_000)
        if r is None or r.status_code != 200:
            return []
        with pdfplumber.open(io.BytesIO(r.content)) as pdf:
            text = "\n".join((p.extract_text() or "") for p in pdf.pages[:15])
        return [re.sub(r"\s+", " ", ln).strip() for ln in unhide_emails(text).split("\n") if ln.strip()]
    except Exception:
        return []


class Browser:
    """A headless Chromium opened only if a page needs it. One per site, so it is thread-safe."""
    def __init__(self):
        self.pw = self.browser = None
        self.failed = False

    def page_html(self, url, search_term=None):
        if self.failed:
            return url, "", []
        try:
            if self.browser is None:
                from playwright.sync_api import sync_playwright
                self.pw = sync_playwright().start()
                self.browser = self.pw.chromium.launch()
            page = self.browser.new_page(user_agent=USER_AGENT)
            page.set_default_timeout(15000)
            page.set_default_navigation_timeout(20000)
            page.goto(url, wait_until="domcontentloaded")
            try:
                page.wait_for_load_state("networkidle", timeout=6000)
            except Exception:
                pass
            results = []
            if search_term:
                box = page.locator("input[type=search], input[name*=search i], input[placeholder*=search i], "
                                   "input[aria-label*=search i], input[id*=search i]").first
                if box.count():
                    box.fill(search_term)
                    box.press("Enter")
                    try:
                        page.wait_for_load_state("networkidle", timeout=8000)
                    except Exception:
                        pass
                    results = [a for a in page.eval_on_selector_all(
                        "a[href]", "els => els.map(e => [e.href, e.innerText])") if a[0].startswith("http")]
            html_ = page.content()
            final = page.url
            page.close()
            return final, html_, results
        except Exception:
            self.failed = self.browser is None     # browser missing -> stop trying for this site
            return url, "", []

    def close(self):
        try:
            if self.browser:
                self.browser.close()
            if self.pw:
                self.pw.stop()
        except Exception:
            pass


def crawl_site_deep(url, kind, use_browser=True, max_pages=25, max_renders=6):
    prof = PROFILES[kind]
    words = URL_WORDS[kind]
    session = requests.Session()
    browser = Browser() if use_browser else None
    result = {"ads": [], "pairs": [], "emails": [], "pages": 0, "rendered": 0, "pdfs": 0, "deep": True}
    root = "{0.scheme}://{0.netloc}".format(urlparse(url))
    texts, visited = [], set()
    deadline = time.time() + SITE_TIME_LIMIT

    def add_text(t):
        t = unhide_emails(t)
        texts.append(html_to_lines(t))
        for e in re.findall(r"mailto:([^\"'?\s>]+)", t):
            texts[-1].append(e)

    def get(u, render=False, search=None):
        nonlocal max_pages
        key = u.split("#")[0].rstrip("/")
        if key in visited or max_pages <= 0 or time.time() > deadline or not allowed(u):
            return u, "", []
        visited.add(key)
        max_pages -= 1
        fu, t = fetch_deep(u, session)
        needs = render or not t or len(re.sub(r"<[^>]+>", " ", t)) < 1500 or JS_HINTS.search(t or "")
        if browser and needs and result["rendered"] < max_renders:
            result["rendered"] += 1
            fu2, t2, links = browser.page_html(u, search)
            if t2:
                return fu2, t2, links
        return fu, t, []

    try:
        home_url, home, _ = get(url, render=False)
        if home:
            add_text(home)
        # 1. site search box (browser only)
        if browser:
            for term in SEARCH_TERMS[kind]:
                if result["rendered"] >= max_renders or time.time() > deadline:
                    break
                result["rendered"] += 1
                fu, t, links = browser.page_html(root if home_url.startswith(root) else home_url, term)
                if links:
                    add_text(t)
                    hits = [h for h, txt in links if words.search(h + " " + (txt or ""))
                            and reg_domain(domain(h)) == reg_domain(domain(root))][:3]
                    for h in hits:
                        fu2, t2, _ = get(h)
                        if t2:
                            add_text(t2)
        # 2. candidate pages: sitemap + common paths + links on the home page
        cands = []
        for u in (sitemap_urls(root, session) if time.time() < deadline else []):
            path = urlparse(u).path.lower()
            s = len(words.findall(path)) + (2 if prof["focus"].search(path) else 0)
            if s:
                cands.append((-s, u))
        cands.sort()
        queue = [u for _, u in cands[:8]]
        queue += [root + p for p in DEEP_PATHS[kind]]
        if home:
            queue += links_from(home_url, home, same_site_only=False, prof=prof)[:8]
        pdf_links = []
        for u in dict.fromkeys(queue):
            if max_pages <= 0 or time.time() > deadline:
                break
            if u.lower().endswith(".pdf"):
                pdf_links.append(u)
                continue
            fu, t, _ = get(u, render=bool(re.search(r"directory|staff", u, re.I)))
            if not t:
                continue
            add_text(t)
            for m in re.finditer(r'(?is)<a[^>]+href="([^"#]+\.pdf)"[^>]*>(.*?)</a>', t):
                if re.search(r"staff|directory|contact|athlet|park|recreation|coach|board|personnel",
                             m.group(1) + m.group(2), re.I):
                    pdf_links.append(urljoin(fu, m.group(1)))
            # one level down from directory-style pages (e.g. a county directory -> the parks department)
            if re.search(r"directory|departments?|staff", u, re.I):
                for sub in links_from(fu, t, same_site_only=True, prof=prof)[:4]:
                    fu2, t2, _ = get(sub)
                    if t2:
                        add_text(t2)
        # 3. PDF staff lists
        for p in list(dict.fromkeys(pdf_links))[:3]:
            if time.time() < deadline and allowed(p):
                lines = pdf_lines(p, session)
                if lines:
                    texts.append(lines)
                    result["pdfs"] += 1
    finally:
        if browser:
            browser.close()
    for lines in texts:
        ads, pairs, emails = extract_from_lines(lines, prof)
        result["ads"] += ads
        result["pairs"] += pairs
        result["emails"] += emails
    result["pages"] = len(visited)
    return result


def browser_available():
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch()
            b.close()
        return True
    except Exception:
        return False


# ---------------------------------------------------------------- 11c. phones + emails from every website (free)
PHONE_RE = re.compile(r"(?<![\d-])(?:\+?1[\s.\-]?)?\(?([2-9]\d{2})\)?[\s.\-]?([2-9]\d{2})[\s.\-]?(\d{4})(?![\d-])")
FAX_WORDS = re.compile(r"\bfax\b|\bf:|facsimile", re.I)
GOOD_PHONE_WORDS = {
    "Public": re.compile(r"park|recreation|leisure|athlet|sports|office|main|front desk|registration", re.I),
    "School": re.compile(r"athlet|sports|main office|front office|school|office|phone", re.I),
    "Private": re.compile(r"office|main|front desk|phone|call|contact|info", re.I),
    "Unknown": re.compile(r"office|main|phone|call|contact", re.I),
}
CONTACT_LINK_WORDS = re.compile(r"contact|about|staff|directory|location|hours|visit|parks|recreation|athlet|office", re.I)
VENDOR_EMAIL = re.compile(r"@(sentry|wix|wixpress|squarespace|godaddy|finalsite|schoolmessenger|blackboard|edlio|"
                          r"apptegy|civicplus|granicus|example|domain|email|yourdomain|sitelock)\.", re.I)


def _phones_from_lines(lines):
    """[(10-digit phone, context line)] skipping fax numbers."""
    out = []
    for i, ln in enumerate(lines):
        for m in PHONE_RE.finditer(ln):
            before = ln[max(0, m.start() - 25):m.start()]
            if FAX_WORDS.search(before) or (FAX_WORDS.search(ln) and not re.search(r"phone|tel|call", before, re.I)):
                continue
            ctx = " ".join(lines[max(0, i - 1): i + 1])
            out.append((m.group(1) + m.group(2) + m.group(3), ctx[:160]))
    return out


def site_contact_info(url):
    """Home page + up to 5 contact-type pages: phone numbers (with context) and emails. Plain requests, no browser."""
    session = requests.Session()
    info = {"phones": [], "tel_links": [], "emails": []}
    root = "{0.scheme}://{0.netloc}".format(urlparse(url))
    deadline = time.time() + 60
    fu, home = fetch_deep(url, session)
    if not home:
        return info
    queue = [fu]
    for m in re.finditer(r'(?is)<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', home):
        href, label = m.group(1), re.sub(r"<[^>]+>", " ", m.group(2))
        if SKIP_LINK.search(href) or not CONTACT_LINK_WORDS.search(href + " " + label):
            continue
        u = urljoin(fu, href)
        if u.startswith("http") and reg_domain(domain(u)) == reg_domain(domain(fu)):
            queue.append(u.split("#")[0])
    queue += [root + "/contact", root + "/contact-us"]
    seen, pages = set(), 0
    for u in dict.fromkeys(queue):
        if pages >= 6 or time.time() > deadline:
            break
        key = u.rstrip("/")
        if key in seen:
            continue
        seen.add(key)
        t = home if u == fu else fetch_deep(u, session)[1]
        if not t:
            continue
        pages += 1
        t = unhide_emails(t)
        info["tel_links"] += [re.sub(r"\D", "", x)[-10:] for x in re.findall(r'href="tel:([^"]+)"', t, re.I)]
        lines = html_to_lines(t)
        info["phones"] += _phones_from_lines(lines)
        info["emails"] += [e.lower() for e in EMAIL_RE.findall(" ".join(lines) + " " + " ".join(
            re.findall(r"mailto:([^\"'?\s>]+)", t)))]
    info["emails"] = sorted(set(info["emails"]))
    return info


def contact_details(df, state, workers=8):
    """Collect phones/emails for every lead site and parks department site (cached)."""
    path = os.path.join("cache", f"contact_info_{state}.json")
    cache = load_json(path, {})
    sites = set()
    for _, r in df.iterrows():
        for u in (r["website"], r.get("agency_site", "")):
            if u and not re.search(r"facebook|instagram|google\.|yelp|stadiumconnection", u, re.I):
                sites.add(u)
    todo = [u for u in sorted(sites) if u not in cache]
    if todo:
        print(f"Collecting phone numbers and emails from {len(todo)} websites (roughly {max(1, len(todo) // 60)} min)...")
        from concurrent.futures import as_completed
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = {ex.submit(site_contact_info, u): u for u in todo}
            for n, f in enumerate(as_completed(futs), 1):
                try:
                    cache[futs[f]] = f.result()
                except Exception:
                    cache[futs[f]] = {"phones": [], "tel_links": [], "emails": []}
                if n % 25 == 0 or n == len(todo):
                    save_json(path, cache)
                    print(f"  {n}/{len(todo)} websites checked (saved)")
    # emails already seen by the earlier searches are free extras
    for fname in ("emails.json", f"ad_finder_{state}.json", f"parks_finder_{state}.json",
                  f"deep_finder_ad_{state}.json", f"deep_finder_parks_{state}.json", f"deep_finder_private_{state}.json"):
        extra = load_json(os.path.join("cache", fname), {})
        for u, v in extra.items():
            ems = v if isinstance(v, list) else v.get("emails", [])
            if u in cache and ems:
                cache[u]["emails"] = sorted(set(cache[u]["emails"]) | {e.lower() for e in ems})
            elif u in sites and ems:
                cache[u] = {"phones": [], "tel_links": [], "emails": sorted({e.lower() for e in ems})}
    return cache


def best_phone(info, owner_type):
    if not info:
        return ""
    good = GOOD_PHONE_WORDS.get(owner_type, GOOD_PHONE_WORDS["Unknown"])
    score = {}
    for p in info.get("tel_links", []):
        if len(p) == 10:
            score[p] = score.get(p, 0) + 3
    for p, ctx in info.get("phones", []):
        score[p] = score.get(p, 0) + 1 + (2 if good.search(ctx) else 0)
    if not score:
        return ""
    p = max(score.items(), key=lambda x: x[1])[0]
    return f"({p[:3]}) {p[3:6]}-{p[6:]}"


def best_inbox(info, site, owner_type):
    if not info:
        return ""
    sd = reg_domain(domain(site)) if site else ""
    cands = []
    for e in info.get("emails", []):
        if re.match(BAD_EMAIL_PREFIX, e) or VENDOR_EMAIL.search(e) or re.search(r"\.(png|jpg|gif|webp|svg)$", e):
            continue
        personal = bool(re.search(PERSONAL_EMAIL, e))
        if personal and owner_type in ("Public", "School"):
            continue
        same = sd and reg_domain(e.split("@")[-1]) == sd
        role = re.search(r"park|rec|athlet|sport|info|office|contact|admin|director|field|facilit", e.split("@")[0])
        cands.append(((2 if same else 0) + (1 if role else 0) - (1 if personal else 0), e))
    if not cands:
        return ""
    return max(cands)[1]


# ---------------------------------------------------------------- 12. assemble contacts
TARGET_ROLE = {"Public": "Parks & Recreation Director / Athletics Manager",
               "School": "Athletic Director",
               "Private": "Owner / General Manager / Facility Director",
               "Unknown": "Facility Manager"}


def email_rank(quality, email):
    q = (quality or "").lower()
    if not email:
        return 0
    if "guessed" in q:
        return 2
    if re.match(BAD_EMAIL_PREFIX, email.lower()) or re.search(PERSONAL_EMAIL, email.lower()):
        return 1
    if "accept all" in q or "catch" in q:
        return 3
    return 4


def confidence(c):
    if not c["name"]:
        return "Low - general inbox"
    q = c["quality"].lower()
    if not c["email"]:
        base = "Medium - named, phone only" if c["phone"] else "Low - name only, no email"
    elif "guessed" in q:
        base = "Low - email guessed from district format, verify"
    elif "personal" in q:
        base = "Medium - personal email"
    elif "accept all" in q:
        base = "Medium - work email (accept all)"
    elif c["source"] == "Seamless" and "valid" in q:
        base = "High - verified work email"
    else:
        base = f"Medium - email from {c['source'].lower()}"
    return base + ("; check title" if c["title"] and title_score(c["title"]) < 4 and c["source"] == "Seamless" else "")


def assemble(df, agencies, seamless, hs, do_emails, people_found, cache_dir, state, contact_info=None):
    ag = {r["Agency Key"]: r for _, r in agencies.iterrows()}
    contact_info = contact_info or {}
    # parks-department phone + email, shared by all of that department's parks
    dept = {}
    for _, r in df.iterrows():
        if r["owner_type"] != "Public":
            continue
        d = dept.setdefault(r["agency_key"], {"phones": {}, "inboxes": {}})
        for site, w in ((r.get("agency_site", ""), 3), (r["website"], 1)):
            info = contact_info.get(site)
            ph, ib = best_phone(info, "Public"), best_inbox(info, site, "Public")
            if ph:
                d["phones"][ph] = d["phones"].get(ph, 0) + w
            if ib:
                d["inboxes"][ib] = d["inboxes"].get(ib, 0) + w
    dept = {k: (max(v["phones"], key=v["phones"].get) if v["phones"] else "",
                max(v["inboxes"], key=v["inboxes"].get) if v["inboxes"] else "") for k, v in dept.items()}
    email_cache = load_json(os.path.join(cache_dir, "emails.json"), {})
    if do_emails:
        n = sum(1 for _, r in df.iterrows() if r["owner_type"] != "Public" and r["website"]
                and r["website"] not in email_cache)
        if n:
            print(f"Checking {n} websites for email addresses...")
    reg_count, phone_count = {}, {}
    for _, r in df.iterrows():
        d = domain(r["website"])
        if d:
            reg_count[reg_domain(d)] = reg_count.get(reg_domain(d), 0) + 1
        p = digits10(r["phone"])
        if p:
            phone_count[p] = phone_count.get(p, 0) + 1
    out, used_seamless = [], set()
    for i, r in df.iterrows():
        r = r.to_dict()
        a = ag.get(r["agency_key"]) if r["owner_type"] == "Public" else None
        r["agency_dept"] = clean(a["Department Name"]) if a is not None else ""
        r["agency_site"] = clean(a["Department Website"]) if a is not None else ""
        if r["owner_type"] == "Public":
            r["buyer"] = r["agency_dept"] or r["agency_key"]
            if r["agency_site"] and not r["website"]:
                r["website"] = r["agency_site"]
        elif r["owner_type"] == "School":
            r["buyer"] = r["company"]
        else:
            r["buyer"] = r["company"]
        r["target_role"] = TARGET_ROLE.get(r["owner_type"], "Facility Manager")
        r["district"] = f"{r['county']} school district (verify)" if r["owner_type"] == "School" and r["county"] else ""
        cands = []
        if a is not None and (clean(a["Contact Name"]) or clean(a["Contact Email"])):
            em, quality = clean(a["Contact Email"]), "Work email" if clean(a["Contact Email"]) else ""
            nm = clean(a["Contact Name"])
            if nm and not em:
                patterns = people_found.get("_patterns", {})
                site_doms = people_found.get("_site_domains", {})
                edom = site_doms.get(r["agency_site"]) or site_doms.get(r["website"]) or ""
                if edom in patterns:
                    first, last = split_name(nm)
                    if first and last:
                        pat, n = patterns[edom]
                        em = f"{make_local(pat, first, last)}@{edom}"
                        quality = f"Guessed - {pat}@{edom} format ({n} examples)"
            cands.append({"name": nm, "title": clean(a["Contact Title"]), "email": em,
                          "phone": clean(a["Contact Phone"]), "source": "Agencies file",
                          "quality": quality, "tscore": 5})
        for p in seamless_for(r, seamless, reg_count):
            used_seamless.add(id(p))
            cands.append({"name": p["name"], "title": p["title"], "email": p["email"], "phone": p["phone"],
                          "source": "Seamless", "quality": p["email_quality"], "tscore": p["score"]})
        if r["owner_type"] == "School":
            key = ("ad", r["website"])
        elif r["owner_type"] == "Public":
            key = ("parks", r["agency_site"] or r["website"])
        else:
            key = ("private", r["website"])
        wp = people_found.get(key)
        if wp:
            cands.append({"name": wp["name"], "title": wp["title"], "email": wp["email"], "phone": "",
                          "source": "Website (deep search)" if wp.get("deep") else "Website",
                          "quality": wp["status"], "tscore": 5})
        inbox = ""
        if r["osm_email"]:
            inbox = r["osm_email"]
        own_info = contact_info.get(r["website"])
        r["phone_source"] = "Google / map data" if r["phone"] else ""
        if not r["phone"]:
            r["phone"] = best_phone(own_info, r["owner_type"])
            r["phone_source"] = "Website" if r["phone"] else ""
        if not r["phone"] and r["owner_type"] == "Public":
            r["phone"] = dept.get(r["agency_key"], ("", ""))[0]
            r["phone_source"] = "Parks department (shared)" if r["phone"] else ""
        inbox = inbox or best_inbox(own_info, r["website"], r["owner_type"])
        if do_emails and r["owner_type"] != "Public" and r["website"]:
            found = scrape_emails(r["website"], email_cache)
            good = [e for e in found if not re.match(BAD_EMAIL_PREFIX, e)]
            if good:
                inbox = inbox or pick_email(good, domain_of(r["website"]))
            if i % 20 == 0:
                save_json(os.path.join(cache_dir, "emails.json"), email_cache)
        elif r["website"] in email_cache:
            good = [e for e in email_cache[r["website"]] if not re.match(BAD_EMAIL_PREFIX, e)]
            if good:
                inbox = inbox or pick_email(good, domain_of(r["website"]))
        place_names = norm_match(f"{r['company']} {r['buyer']}")
        cands = [c for c in cands if not (c["name"] and c["source"].startswith("Website")
                                           and norm_match(c["name"]) in place_names)]
        # merge the same person from two sources (e.g. agencies file name + Seamless email)
        merged = []
        for c in cands:
            same = next((m for m in merged if split_name(m["name"])[1] and
                         split_name(m["name"])[1] == split_name(c["name"])[1]), None)
            if same:
                for f in ("email", "phone", "title"):
                    if not same[f] and c[f]:
                        same[f] = c[f]
                        if f == "email":
                            same["quality"] = c["quality"]
                same["source"] = same["source"] if c["source"] in same["source"] else same["source"] + " + " + c["source"]
            else:
                merged.append(dict(c))
        merged.sort(key=lambda c: (-email_rank(c["quality"], c["email"]) if c["email"] else 0,
                                   -min(c["tscore"], 5), -bool(c["phone"])))
        # a named decision-maker beats a generic inbox; the inbox stays in its own column
        persons = [c for c in merged if c["name"]]
        persons.sort(key=lambda c: (not c["source"].startswith("Agencies file"),     # your hand-checked contacts first
                                    -(email_rank(c["quality"], c["email"]) >= 2), -min(c["tscore"], 5),
                                    -email_rank(c["quality"], c["email"]), -bool(c["phone"])))
        if not inbox and r["owner_type"] == "Public":
            inbox = dept.get(r["agency_key"], ("", ""))[1]
        best = persons[0] if persons else None
        if best is None and inbox:
            best = {"name": "", "title": "", "email": inbox, "phone": "", "source": "Website inbox",
                    "quality": "inbox", "tscore": 0}
        r.update({"contact_name": "", "contact_title": "", "contact_email": "", "contact_phone": "",
                  "contact_source": "", "confidence": "None - needs research", "email_status": ""})
        if best:
            r.update({"contact_name": best["name"], "contact_title": best["title"], "contact_email": best["email"],
                      "contact_phone": best["phone"], "contact_source": best["source"], "confidence": confidence(best),
                      "email_status": best["quality"] if "guessed" in best["quality"].lower() else ""})
        r["other_contacts"] = "; ".join(
            f"{c['name']} ({c['title'] or c['source']}{', ' + c['email'] if c['email'] else ''}"
            f"{', ' + c['phone'] if c['phone'] else ''})" for c in persons[1:4] if c is not best)
        # a named person without a direct email gets the organization's general inbox
        if r["contact_name"] and not r["contact_email"] and inbox:
            r["contact_email"] = inbox
            r["email_status"] = "General inbox (no direct email found)"
            r["confidence"] = "Medium - named person, general inbox email"
        r["general_inbox"] = inbox if inbox and inbox != r["contact_email"] else ""
        # CRM ready = zip + a way to reach them (an email that isn't just a guess, or a phone number). A name is a bonus.
        guessed = "guessed" in r["confidence"].lower()
        real_email = (r["contact_email"] and not guessed) or r["general_inbox"]
        phone_any = r["contact_phone"] or r["phone"]
        if r["contact_name"] and r["contact_email"] and not guessed:
            level = "Named contact + email"
        elif r["contact_name"] and (real_email or phone_any):
            level = "Named contact + " + ("email" if real_email else "phone")
        elif real_email:
            level = "Email (no name)" + (" + phone" if phone_any else "")
        elif phone_any:
            level = "Phone only"
        elif r["contact_name"]:
            level = "Name only (verify guessed email)" if guessed else "Name only"
        else:
            level = "None"
        r["contact_level"] = level
        missing = [x for x, ok in (("email or phone", real_email or phone_any), ("zip", r["zip"])) if not ok]
        r["crm_ready"] = "Yes" if not missing else "Missing: " + ", ".join(missing)
        out.append(r)
    save_json(os.path.join(cache_dir, "emails.json"), email_cache)
    res = pd.DataFrame(out)
    res["hubspot_status"], res["hubspot_match"] = zip(*[hubspot_status(r, hs, reg_count, phone_count)
                                                        for r in res.to_dict("records")])
    return res, len(used_seamless)


# ---------------------------------------------------------------- 13. export
LEAD_COLS = [
    ("Rank", 7), ("Tier", 6), ("Score", 7), ("Complex Name", 34), ("Owner Type", 11),
    ("Buyer / Agency", 30), ("Target Role", 26), ("District (guess)", 24), ("Fields", 7), ("Grass", 7), ("Turf", 7), ("Unknown Surface", 9),
    ("Tracks", 7), ("Sports", 28), ("Street Address", 28), ("City", 16), ("County", 18),
    ("State", 6), ("Zip", 8), ("Phone", 15), ("Website", 30), ("Contact Name", 22),
    ("Contact Title", 22), ("Contact Email", 30), ("Contact Phone", 15), ("Contact Source", 14),
    ("Contact Confidence", 24), ("CRM Ready", 16), ("Contact Level", 22), ("Email Status", 24), ("Other Contacts", 40), ("General Inbox", 26), ("HubSpot Status", 18), ("HubSpot Match", 24),
    ("Satellite View", 12), ("Lat", 10), ("Lon", 11), ("Lead ID", 9),
    ("Status", 12), ("Notes", 30),
]


def export(df, fields_df, agencies, state, out_path):
    wb = Workbook()
    bold = Font(name="Arial", bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", start_color="1F3864")
    base = Font(name="Arial", size=10)
    input_fill = PatternFill("solid", start_color="FFFF00")

    # --- Settings
    st = wb.active
    st.title = "Settings"
    settings = [
        ("Scoring weights (points per field)", None, ""),
        ("Grass field", W_GRASS, "Natural grass needs frequent repainting: best fit for TMR."),
        ("Unknown surface", W_UNKNOWN, "Surface not tagged in OSM; often grass."),
        ("Artificial turf", W_TURF, "Lines usually permanent; lower value."),
        ("Track (not wrapping a field)", W_TRACK, "Standalone running track. Tracks around a field aren't counted separately."),
        ("", None, ""),
        ("Tier thresholds (score)", None, ""),
        ("Tier A: score at least", TIER_A, "Green rows"),
        ("Tier B: score at least", TIER_B, "Yellow rows. Everything below is Tier C (gray)."),
    ]
    for i, (label, val, note) in enumerate(settings, start=1):
        st.cell(i, 1, label).font = Font(name="Arial", bold=val is None, size=10)
        if val is not None:
            c = st.cell(i, 2, val)
            c.font = Font(name="Arial", size=10, color="0000FF")
            c.fill = input_fill
        st.cell(i, 3, note).font = Font(name="Arial", size=9, italic=True, color="666666")
    st.cell(11, 1, "Edit the yellow cells. Scores, ranks, tiers and colors on the Leads tab update automatically.") \
        .font = Font(name="Arial", size=10, italic=True)
    st.cell(12, 1, "Default values were chosen by Claude as a starting point, not from TMR data.") \
        .font = Font(name="Arial", size=9, italic=True, color="666666")
    st.column_dimensions["A"].width = 32
    st.column_dimensions["B"].width = 10
    st.column_dimensions["C"].width = 80

    # --- Leads
    ws = wb.create_sheet("Leads", 0)
    col = {name: i + 1 for i, (name, _) in enumerate(LEAD_COLS)}
    L = {name: get_column_letter(i) for name, i in col.items()}
    for name, width in LEAD_COLS:
        c = ws.cell(1, col[name], name)
        c.font, c.fill = bold, head_fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[L[name]].width = width
    ws.row_dimensions[1].height = 30
    n = len(df)
    last = n + 1
    for i, r in enumerate(df.itertuples(index=False), start=2):
        d = r._asdict()
        vals = {
            "Complex Name": d["company"] or "(unnamed complex)", "Owner Type": d["owner_type"],
            "Buyer / Agency": d["buyer"], "Target Role": d["target_role"], "District (guess)": d["district"], "Fields": d["units"], "Grass": d["grass"], "Turf": d["turf"],
            "Unknown Surface": d["unknown"], "Tracks": d["tracks"], "Sports": d["sports"],
            "Street Address": d["street"], "City": d["city"], "County": d["county"], "State": state,
            "Zip": d["zip"], "Phone": d["phone"], "Website": d["website"],
            "Contact Name": d["contact_name"], "Contact Title": d["contact_title"],
            "Contact Email": d["contact_email"], "Contact Phone": d["contact_phone"],
            "Contact Source": d["contact_source"], "Contact Confidence": d["confidence"],
            "CRM Ready": d["crm_ready"], "Contact Level": d["contact_level"], "Email Status": d["email_status"], "Other Contacts": d["other_contacts"], "General Inbox": d["general_inbox"],
            "HubSpot Status": d["hubspot_status"], "HubSpot Match": d["hubspot_match"],
            "Lat": d["lat"], "Lon": d["lon"], "Lead ID": d["lead_id"],
        }
        for k, v in vals.items():
            ws.cell(i, col[k], v).font = base
        # standalone tracks = counted units that aren't fields
        ws.cell(i, col["Score"],
                f"=Settings!$B$2*{L['Grass']}{i}+Settings!$B$3*{L['Unknown Surface']}{i}"
                f"+Settings!$B$4*{L['Turf']}{i}"
                f"+Settings!$B$5*({L['Fields']}{i}-{L['Grass']}{i}-{L['Turf']}{i}-{L['Unknown Surface']}{i})").font = base
        ws.cell(i, col["Rank"], f"=RANK({L['Score']}{i},${L['Score']}$2:${L['Score']}${last})").font = base
        ws.cell(i, col["Tier"],
                f'=IF({L["Score"]}{i}>=Settings!$B$8,"A",IF({L["Score"]}{i}>=Settings!$B$9,"B","C"))').font = base
        link = ws.cell(i, col["Satellite View"],
                       f'=HYPERLINK("https://www.google.com/maps/@{d["lat"]},{d["lon"]},500m/data=!3m1!1e3","Open map")')
        link.font = Font(name="Arial", size=10, color="0563C1", underline="single")
    ws.freeze_panes = "E2"
    if n:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(LEAD_COLS))}{last}"
        rng = f"A2:{get_column_letter(len(LEAD_COLS))}{last}"
        for tier, color in TIER_COLORS.items():
            ws.conditional_formatting.add(rng, FormulaRule(
                formula=[f'${L["Tier"]}2="{tier}"'], fill=PatternFill("solid", start_color=color, end_color=color)))
    ws.cell(1, col["Fields"]).comment = Comment(
        "Fields + standalone tracks. A track wrapped around a field counts once, as the field.", "TMR")
    ws.cell(1, col["Contact Confidence"]).comment = Comment(
        "High = Seamless-verified work email. Medium = named person from the agencies file, a website or Seamless "
        "without a verified email. Low = guessed email or general inbox. None = needs research.", "TMR")
    ws.cell(1, col["HubSpot Status"]).comment = Comment(
        "New = not in HubSpot. Parent org in HubSpot = the district, university or county government is, this buyer "
        "isn't. For public parks, Already in HubSpot means the parks department itself is there.", "TMR")

    # --- Agencies
    ag = wb.create_sheet("Agencies")
    for j, h in enumerate(AGENCY_COLS, start=1):
        c = ag.cell(1, j, h)
        c.font, c.fill = bold, head_fill
        ag.column_dimensions[get_column_letter(j)].width = 22
    for i, r in enumerate(agencies.itertuples(index=False), start=2):
        for j, v in enumerate(r, start=1):
            ag.cell(i, j, v).font = base
    ag.cell(1, len(AGENCY_COLS) + 2, "Read-only copy. Fill in agencies/agencies_" + state +
            ".csv instead, then re-run.").font = Font(name="Arial", italic=True, size=9)

    # --- Raw fields
    rf = wb.create_sheet("Raw Fields")
    cols = ["Lead ID", "OSM ID", "Kind", "Counted", "Sport", "Surface", "Lat", "Lon"]
    for j, h in enumerate(cols, start=1):
        c = rf.cell(1, j, h)
        c.font, c.fill = bold, head_fill
        rf.column_dimensions[get_column_letter(j)].width = 14
    for i, r in enumerate(fields_df.itertuples(index=False), start=2):
        for j, v in enumerate([r.lead_id, r.osm_id, r.kind, "yes" if r.counted else "no (wraps a field)",
                               r.sport, r.surface, r.lat, r.lon], start=1):
            rf.cell(i, j, v).font = base

    wb.save(out_path)


def export_hubspot_csv(df, state, path):
    """One row per buyer with a contact email, ready for HubSpot's import (companies + contacts in one file)."""
    ready = df[(df["hubspot_status"].isin(["New", "Not checked", "Parent org in HubSpot"])) & (df["company"] != "")
               & (df["crm_ready"] == "Yes")].copy()
    rows = {}
    for _, r in ready.iterrows():
        key = r["buyer"] if r["owner_type"] == "Public" else f"{r['company']}|{r['zip']}"
        first, _, last = r["contact_name"].partition(" ")
        if key not in rows:
            rows[key] = {
                "Company name": r["buyer"] or r["company"], "Company Domain Name": domain_of(r["website"]),
                "Website URL": r["website"], "Phone Number": r["phone"], "Street Address": r["street"],
                "City": r["city"], "State/Region": state, "Postal Code": r["zip"],
                "First Name": first, "Last Name": last, "Email": r["contact_email"] if "guessed" not in r["confidence"].lower()
                                                                    else (r["general_inbox"] or ""),
                "Contact Phone": r["contact_phone"] or r["phone"],
                "Job Title": r["contact_title"], "Owner Type": r["owner_type"], "Lead Tier": r["tier"],
                "Total Fields": 0, "Facilities": [], "Lead Source": "TMR Complex Finder",
            }
        rows[key]["Total Fields"] += int(r["units"])
        rows[key]["Facilities"].append(r["company"])
    out = pd.DataFrame(rows.values())
    if len(out):
        out["Facilities"] = out["Facilities"].map(lambda x: "; ".join(dict.fromkeys(x)))
    out.to_csv(path, index=False, quoting=csv.QUOTE_MINIMAL)
    return len(out)


# ---------------------------------------------------------------- 14. call list
def write_call_list(df, state):
    path = f"call_list_{state}.csv"
    need = df[(df["phone"] != "") & (~df["contact_level"].str.contains("email", case=False))
              & (~df["hubspot_status"].isin(["Already in HubSpot", "Likely duplicate"]))].copy()
    need = need.sort_values(["tier", "score"], ascending=[True, False]).drop_duplicates("buyer")
    ask = {"Public": "Who oversees your athletic fields and field maintenance?",
           "School": "Who is your athletic director, and what's the best email for them?",
           "Private": "Who manages your fields - the owner or a facility manager?",
           "Unknown": "Who manages the fields at this complex?"}
    out = pd.DataFrame({
        "Lead ID": need["lead_id"], "Tier": need["tier"], "Buyer": need["buyer"], "Complex": need["company"],
        "City": need["city"], "Main phone": need["phone"], "Ask for": need["target_role"],
        "Suggested question": need["owner_type"].map(ask), "Already have": need["contact_level"],
        "Name found": "", "Email found": "", "Notes": ""})
    out.to_csv(path, index=False)
    return path, len(out)


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--state", default="GA")
    ap.add_argument("--min-fields", type=int, default=2)
    ap.add_argument("--radius", type=float, default=400, help="grouping distance in meters")
    ap.add_argument("--geocode", action="store_true", help="look up street/city/county/zip")
    ap.add_argument("--places", action="store_true", help="enrich with Google Places (needs API key)")
    ap.add_argument("--places-key", default=os.environ.get("GOOGLE_PLACES_KEY", ""))
    ap.add_argument("--emails", action="store_true", help="scan private complex websites for emails")
    ap.add_argument("--google-budget", type=int, default=900, help="max NEW Google calls this run (0 = saved answers only)")
    ap.add_argument("--find-contacts", action="store_true",
                    help="find athletic directors (school sites) and parks & rec directors (department sites)")
    ap.add_argument("--deep-search", action="store_true",
                    help="free deeper website search: browser, sitemaps, directories, search boxes, PDFs, "
                         "hidden emails; also private complex sites (implies --find-contacts)")
    ap.add_argument("--deep-limit", type=int, default=0, help="deep-search only the N highest-scoring sites (for a test)")
    ap.add_argument("--no-browser", action="store_true", help="deep search without opening a browser")
    a = ap.parse_args()
    state = a.state.upper()
    cache_dir = "cache"
    os.makedirs(cache_dir, exist_ok=True)

    fields, tracks, places = split_elements(fetch_osm(state, cache_dir))
    print(f"Found {len(fields)} fields, {len(tracks)} tracks, {len(places)} named parks/schools/centres.")
    units = cluster(build_units(fields, tracks), a.radius)
    df = summarise(units, places, a.min_fields)
    if df.empty:
        print("No complexes found with those settings.")
        return
    print(f"{len(df)} complexes with {a.min_fields}+ fields.")

    if a.geocode:
        df = geocode(df, cache_dir)
    else:
        for c in ("street", "city", "county", "zip"):
            df[c] = ""
    if a.places:
        if not a.places_key and a.google_budget > 0:
            raise SystemExit("--places needs a key: set GOOGLE_PLACES_KEY, or add --google-budget 0 to reuse saved answers only")
        df = places_enrich(df, cache_dir, a.places_key, state, a.google_budget)
    for c in ("g_name", "g_address", "g_phone", "g_website", "g_types", "g_status", "g_zip"):
        if c not in df:
            df[c] = ""
    df = df.fillna("")

    # Drop permanently closed places, merge sources (Google wins, OSM fills gaps)
    closed = df["g_status"] == "CLOSED_PERMANENTLY"
    if closed.any():
        print(f"Dropping {closed.sum()} permanently closed places.")
        df = df[~closed].reset_index(drop=True)
    df["company"] = [n or g for n, g in zip(df["name"], df["g_name"])]
    df["phone"] = [g or o for g, o in zip(df["g_phone"], df["osm_phone"])]
    df["website"] = [g or o for g, o in zip(df["g_website"], df["osm_website"])]
    df["zip"] = [z or g for z, g in zip(df["zip"], df["g_zip"])]
    df["street"] = [g.split(",")[0] if g else s for g, s in zip(df["g_address"], df["street"])]
    df["owner_type"] = df.apply(classify, axis=1)
    df["agency_key"] = df.apply(agency_key, axis=1)

    # Sort by default score, assign stable IDs
    track_units = df["units"] - df["grass"] - df["turf"] - df["unknown"]
    df["score"] = W_GRASS * df["grass"] + W_UNKNOWN * df["unknown"] + W_TURF * df["turf"] + W_TRACK * track_units
    df["tier"] = np.where(df["score"] >= TIER_A, "A", np.where(df["score"] >= TIER_B, "B", "C"))
    df = df.sort_values("score", ascending=False).reset_index(drop=True)
    df["lead_id"] = [f"{state}-{i:04d}" for i in range(1, len(df) + 1)]

    agencies = sync_agencies(df, state)
    ag_site = {r["Agency Key"]: clean(r["Department Website"]) for _, r in agencies.iterrows()}
    df["agency_site"] = [ag_site.get(k, "") if t == "Public" else "" for k, t in zip(df["agency_key"], df["owner_type"])]
    seamless, rejected = load_seamless(state)
    hs = load_hubspot(state)
    people_found, n_ad_sites, n_parks_sites, n_formats, deep_stats = {}, 0, 0, 0, None
    use_browser = False
    if a.deep_search and not a.no_browser:
        use_browser = browser_available()
        if not use_browser:
            print("No browser found - deep search will run without one. To add it: pip install playwright, "
                  "then: playwright install chromium")
    if a.find_contacts or a.deep_search:
        people_found, n_ad_sites, n_parks_sites, n_formats, deep_stats = find_people(
            df, state, deep=a.deep_search, deep_limit=a.deep_limit, use_browser=use_browser)
    contact_info = contact_details(df, state) if a.emails else {}
    df, n_seamless_used = assemble(df, agencies, seamless, hs, a.emails, people_found, cache_dir, state, contact_info)
    up_path, n_up = write_seamless_upload(df, agencies, state)

    guessed = df[df["confidence"].str.contains("guessed", case=False)]
    verify_path = f"complex_emails_to_verify_{state}.csv"
    guessed[["lead_id", "buyer", "contact_name", "contact_email", "email_status"]].rename(columns={
        "lead_id": "Lead ID", "buyer": "Buyer", "contact_name": "Contact", "contact_email": "Guessed email",
        "email_status": "Format evidence"}).assign(**{"Verified? (valid/invalid)": ""}).to_csv(verify_path, index=False)
    if rejected:
        pd.DataFrame(rejected)[["name", "title", "company", "email", "phone", "state", "reason"]] \
            .to_csv(f"seamless_unmatched_complexes_{state}.csv", index=False)

    cmap = dict(zip(df["cluster"], df["lead_id"]))
    fields_df = units[units["cluster"].isin(cmap)].copy()
    fields_df["lead_id"] = fields_df["cluster"].map(cmap)
    fields_df = fields_df.sort_values("lead_id")

    call_path, n_call = write_call_list(df, state)
    out = f"{state}_sports_complex_leads_v3.xlsx"
    export(df, fields_df, agencies, state, out)
    hub_path = f"{state}_complex_hubspot_import.csv"
    n_hs = export_hubspot_csv(df, state, hub_path)

    named = (df["contact_name"] != "").sum()
    print("\n---------------- summary ----------------")
    print(f"Complexes: {len(df)}  (A: {(df.tier == 'A').sum()}, B: {(df.tier == 'B').sum()}, C: {(df.tier == 'C').sum()})")
    print("  Owner types: " + ", ".join(f"{k} {v}" for k, v in df["owner_type"].value_counts().items()))
    print(f"With a named contact: {named}   with contact email: {(df['contact_email'] != '').sum()}   "
          f"with zip: {(df['zip'] != '').sum()}")
    print("  Contact sources: " + ", ".join(f"{k or 'none'} {v}" for k, v in df["contact_source"].value_counts().items()))
    if a.find_contacts or a.deep_search:
        found_ad = sum(1 for key in people_found if isinstance(key, tuple) and key[0] == "ad")
        found_pk = sum(1 for key in people_found if isinstance(key, tuple) and key[0] == "parks")
        print(f"Website search: athletic directors on {found_ad} of {n_ad_sites} school sites; "
              f"parks & rec directors on {found_pk} of {n_parks_sites} department sites; "
              f"{len(guessed)} emails guessed from {n_formats} confirmed school/agency email formats -> {verify_path}")
    if deep_stats:
        found_pv = sum(1 for key in people_found if isinstance(key, tuple) and key[0] == "private")
        print(f"Deep search: {deep_stats['found']} of {deep_stats['tried']} sites yielded a name "
              f"(browser pages: {deep_stats['rendered']}, PDFs read: {deep_stats['pdfs']}); "
              f"private complex owners/directors found: {found_pv}")
    ready = (df["crm_ready"] == "Yes")
    target = df["hubspot_status"].isin(["New", "Parent org in HubSpot", "Not checked"])
    has_email = df["contact_level"].str.contains("email", case=False) & ~df["contact_level"].str.startswith("Name only")
    has_phone = (df["phone"] != "") | (df["contact_phone"] != "")
    print(f"Reachable: email {has_email.sum()}, phone {has_phone.sum()}, either {(has_email | has_phone).sum()} of {len(df)} "
          f"(phones from websites {(df['phone_source'] == 'Website').sum()}, shared from parks dept "
          f"{(df['phone_source'] == 'Parks department (shared)').sum()})")
    print("  Contact level: " + ", ".join(f"{k} {v}" for k, v in df["contact_level"].value_counts().items()))
    print(f"CRM ready (zip + email or phone): {ready.sum()} leads, {df[ready & target]['buyer'].nunique()} target buyers")
    print(f"Seamless: {n_seamless_used} people matched" +
          (f"   ({len(rejected)} set aside - see seamless_unmatched_complexes_{state}.csv)" if rejected else ""))
    if hs["loaded"]:
        print("HubSpot:")
        for st_ in ("New", "Parent org in HubSpot", "Possible duplicate", "Likely duplicate", "Already in HubSpot"):
            print(f"  {st_:<22} {(df['hubspot_status'] == st_).sum()}")
    else:
        print("HubSpot: not checked (no CSVs in the hubspot folder)")
    print(f"Spreadsheet:     {out}")
    print(f"Seamless upload: {up_path} ({n_up} buyers still needing a named contact)")
    print(f"HubSpot import:  {hub_path} ({n_hs} buyers: CRM ready, new or parent-org-only)")
    print(f"Call list:       {call_path} ({n_call} buyers we can phone but have no email for yet)")


if __name__ == "__main__":
    main()
