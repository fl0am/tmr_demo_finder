"""TMR Demo Lead Finder - web app.  Run locally with:  streamlit run app.py"""

import datetime as dt
import io

import folium
import pandas as pd
import streamlit as st
from streamlit_folium import st_folium

import demo_finder_core as core

st.set_page_config(page_title="TMR Demo Lead Finder", page_icon="📍", layout="wide")

WARM_COLORS = {"New": "#1f7a3a", "Not checked": "#1f7a3a", "Parent org in HubSpot": "#1f7a3a",
               "Maybe in HubSpot - check": "#c27c0e", "In HubSpot - warm lead": "#c2410c",
               "Likely in HubSpot - warm lead": "#c2410c"}


@st.cache_data(show_spinner=False)
def read_csv(data, name):
    return pd.read_csv(io.BytesIO(data), dtype=str).fillna("")


def opener(lead, date, venue, demo_label):
    where = venue or demo_label.split(",")[0]
    when = f"{date:%A, %B} {date.day}"
    return (f"Hi, this is [your name] with TinyMobileRobots. We're running a live demo of our line-marking robot "
            f"at {where} on {when}, about {lead['miles']:.0f} miles from you. Is the person who handles your "
            f"field painting around? We'd love to have them come see it, or bring it by your fields while "
            f"we're in the area.")


# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.header("HubSpot check")
    uploads = st.file_uploader("HubSpot Companies and Contacts exports (CSV)", type="csv",
                               accept_multiple_files=True)
    st.caption("Optional. Leads already in HubSpot are flagged as warm, not removed. "
               "Files are only used for this search and aren't saved.")
    st.header("Settings")
    radius = st.slider("Search radius (miles)", 5, 50, 25, 5)
    n_leads = st.slider("Leads to return", 5, 10, 10)
    min_fields = st.slider("Minimum fields for parks and private complexes", 1, 6, 2,
                           help="Schools and stadiums are included with any number of fields.")
    people = st.checkbox("Look for named contacts on websites", value=True,
                         help="Finds athletic directors, parks directors and owners. Adds about a minute.")

# ---------------------------------------------------------------- form
st.title("Demo Lead Finder")
st.write("Enter the address of a booked demo to get nearby fields, schools and stadiums to invite.")

with st.form("demo"):
    c1, c2 = st.columns([3, 1])
    address = c1.text_input("Demo address", placeholder="1234 Main St, Marietta, GA 30060")
    demo_date = c2.date_input("Demo date", dt.date.today() + dt.timedelta(days=7))
    venue = st.text_input("Demo venue name (optional, used in the call opener)",
                          placeholder="Cobb County Soccer Complex")
    go = st.form_submit_button("Find leads", type="primary")

if go:
    if not address.strip():
        st.error("Enter the demo address first.")
    else:
        frames = [read_csv(f.getvalue(), f.name) for f in uploads or []]
        with st.status("Finding leads (usually 1 to 2 minutes)...", expanded=True) as s:
            try:
                res = core.find_leads(address, radius, n_leads, min_fields, frames, people, progress=s.write)
                s.update(label="Done", state="complete", expanded=False)
                st.session_state["result"] = res
                st.session_state["meta"] = {"date": demo_date, "venue": venue.strip(), "radius": radius}
            except Exception as e:
                s.update(label="Search stopped", state="error")
                st.error(str(e))

res = st.session_state.get("result")
if not res:
    st.stop()

meta = st.session_state["meta"]
demo, leads, others = res["demo"], res["leads"], res["others"]
st.caption(f"Demo location: {demo['label']}")
if leads.empty:
    st.warning(f"No named fields, schools or stadiums found within {meta['radius']} miles. "
               "Try a bigger radius or a lower minimum field count.")
    st.stop()

m1, m2, m3, m4 = st.columns(4)
m1.metric("Places found nearby", res["n_found"])
m2.metric("With a phone number", int((leads["phone"] != "").sum()))
m3.metric("With a named contact", int((leads["contact_name"] != "").sum()))
warm = leads["hubspot"].str.contains("HubSpot - warm|check", regex=True).sum()
m4.metric("In HubSpot", int(warm) if res.get("hubspot_checked") else "Not checked")

# ---------------------------------------------------------------- map
fmap = folium.Map(location=[demo["lat"], demo["lon"]], zoom_start=10, tiles="OpenStreetMap")
folium.Circle([demo["lat"], demo["lon"]], radius=meta["radius"] * core.MILE_M, color="#1f3864",
              weight=1, fill=False).add_to(fmap)
folium.Marker([demo["lat"], demo["lon"]], tooltip="Demo location",
              icon=folium.Icon(color="red", icon="star")).add_to(fmap)
for _, r in others.iterrows():
    folium.CircleMarker([r["lat"], r["lon"]], radius=4, color="#888", fill=True, fill_opacity=0.6,
                        tooltip=f"{r['name']} ({r['miles']:.1f} mi, {int(r['units'])} fields)").add_to(fmap)
for i, r in leads.reset_index(drop=True).iterrows():
    color = WARM_COLORS.get(r["hubspot"], "#1f7a3a")
    folium.Marker(
        [r["lat"], r["lon"]],
        tooltip=f"{i + 1}. {r['company']} ({r['miles']:.1f} mi)",
        icon=folium.DivIcon(html=f'<div style="background:{color};color:#fff;border:2px solid #fff;'
                                 f'border-radius:50%;width:26px;height:26px;line-height:22px;text-align:center;'
                                 f'font:bold 13px Arial;box-shadow:0 1px 3px rgba(0,0,0,.4)">{i + 1}</div>',
                            icon_size=(26, 26), icon_anchor=(13, 13)),
    ).add_to(fmap)
st_folium(fmap, height=460, use_container_width=True, returned_objects=[])
st.caption("Numbered pins are your call list (green = new, orange = already in HubSpot). "
           "Gray dots are other nearby places.")

# ---------------------------------------------------------------- call list
st.subheader("Call list")
for i, r in leads.reset_index(drop=True).iterrows():
    with st.container(border=True):
        a, b = st.columns([3, 2])
        a.markdown(f"**{i + 1}. {r['company']}**  \n{r['type']}, {r['miles']:.1f} miles away, "
                   f"{int(r['units'])} field{'s' if int(r['units']) != 1 else ''}"
                   + (f" ({r['sports']})" if r["sports"] else ""))
        addr = ", ".join(x for x in (r["street"], r["city"], r["zip"]) if x)
        a.caption(addr or "Address not found")
        b.markdown(f"**HubSpot:** {r['hubspot']}" + (f"  \n{r['hubspot_match']}" if r["hubspot_match"] else ""))
        details = [f"**Phone:** {r['phone'] or 'not found'}"]
        if r["contact_name"]:
            details.append(f"**Contact:** {r['contact_name']}, {r['contact_title']}"
                           + (f", {r['contact_email']}" if r["contact_email"] else ""))
        else:
            details.append(f"**Ask for:** {r['ask_for']}")
        if r["inbox"] and r["inbox"] != r["contact_email"]:
            details.append(f"**Email:** {r['inbox']}")
        if r["owner_type"] == "Public":
            details.append(f"**Runs it:** {r['buyer']}")
        links = [f"[Map]({r['maps_link']})"] + ([f"[Website]({r['website']})"] if r["website"] else [])
        details.append(" | ".join(links))
        st.markdown("  \n".join(details))
        st.code(opener(r, meta["date"], meta["venue"], demo["label"]), language=None, wrap_lines=True)

# ---------------------------------------------------------------- downloads
out = leads.reset_index(drop=True)
csv = pd.DataFrame({
    "#": range(1, len(out) + 1), "Name": out["company"], "Type": out["type"], "Miles": out["miles"].round(1),
    "Fields": out["units"], "Sports": out["sports"], "Phone": out["phone"], "Contact Name": out["contact_name"],
    "Contact Title": out["contact_title"], "Contact Email": out["contact_email"], "General Email": out["inbox"],
    "Ask For": out["ask_for"], "Runs It": out["buyer"], "Street": out["street"], "City": out["city"],
    "State": out["state"], "Zip": out["zip"], "Website": out["website"], "HubSpot": out["hubspot"],
    "HubSpot Match": out["hubspot_match"], "Map": out["maps_link"],
    "Demo Date": meta["date"].isoformat(), "Demo Address": demo["label"],
}).to_csv(index=False)
st.download_button("Download call list (CSV)", csv, file_name=f"demo_leads_{meta['date']:%Y-%m-%d}.csv",
                   mime="text/csv")

if not others.empty:
    with st.expander(f"Other nearby places ({len(others)})"):
        st.dataframe(pd.DataFrame({"Name": others["name"], "Miles": others["miles"].round(1),
                                   "Fields": others["units"], "Sports": others["sports"]}),
                     hide_index=True, width="stretch")
