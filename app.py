"""TMR Demo Lead Finder - web app.  Run locally with:  streamlit run app.py"""

import datetime as dt
import io

import folium
import pandas as pd
import streamlit as st
from streamlit_folium import st_folium

import demo_finder_core as core

TMR_GREEN = "#083400"
WARM = "#C2410C"          # already in HubSpot
CLUB = "#4D7C0F"          # youth clubs
GRAY = "#8A8F86"

st.set_page_config(page_title="TMR Demo Lead Finder", page_icon="📍", layout="wide")
st.markdown(f"""
<style>
  .block-container {{ padding-top: 1.6rem; max-width: 1200px; }}
  h1 {{ color: {TMR_GREEN}; font-weight: 800; letter-spacing: -0.5px; margin-bottom: 0; }}
  div[data-testid="stMetricValue"] {{ color: {TMR_GREEN}; font-weight: 700; }}
  .lead-head {{ border-left: 6px solid {TMR_GREEN}; padding: 0.2rem 0 0.2rem 0.9rem; margin: 0.4rem 0 0.8rem; }}
  .lead-head.warm {{ border-left-color: {WARM}; }}
  .lead-head h3 {{ margin: 0; font-size: 1.3rem; }}
  .lead-head p {{ margin: 0.1rem 0 0; color: #4b5446; }}
  .pill {{ display: inline-block; padding: 1px 9px; border-radius: 10px; font-size: 0.8rem; font-weight: 600;
           background: #E3ECDF; color: {TMR_GREEN}; margin-right: 6px; }}
  .pill.warm {{ background: #FDE7DA; color: {WARM}; }}
  .legend span {{ display: inline-block; width: 11px; height: 11px; border-radius: 50%; margin: 0 5px 0 14px;
                  vertical-align: -1px; }}
</style>""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def read_csv(data, name):
    return pd.read_csv(io.BytesIO(data), dtype=str).fillna("")


def opener(lead, date, venue, demo_label):
    where = venue or demo_label.split(",")[0]
    when = f"{date:%A, %B} {date.day}"
    who = "your club's fields" if lead["type"] == "Youth sports club" else "your fields"
    return (f"Hi, this is [your name] with TinyMobileRobots. We're running a live demo of our line-marking robot "
            f"at {where} on {when}, about {lead['miles']:.0f} miles from you. Is the person who handles painting "
            f"{who} around? We'd love to have them come see it, or bring it by while we're in the area.")


def is_warm(status):
    return "warm" in status or "check" in status


# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.header("HubSpot")
    uploads = st.file_uploader("Companies and Contacts exports (CSV)", type="csv", accept_multiple_files=True)
    st.caption("Optional. Flags leads already in HubSpot and lists HubSpot companies near the demo. "
               "Files are only used for this search and aren't saved.")
    st.header("Settings")
    radius = st.slider("Search radius (miles)", 5, 50, 25, 5)
    n_leads = st.slider("Leads in the call list", 5, 10, 10)
    min_fields = st.slider("Minimum fields for parks and complexes", 1, 6, 2,
                           help="Schools, stadiums and clubs are included with any number of fields.")
    people = st.checkbox("Look for named contacts on websites", value=True,
                         help="Finds athletic directors, parks directors and club leaders. Adds about a minute.")

# ---------------------------------------------------------------- form
st.title("Demo Lead Finder")
st.caption("TinyMobileRobots. Enter a booked demo to get nearby fields, schools, stadiums and clubs to invite.")

with st.form("demo"):
    c1, c2, c3 = st.columns([3, 1.2, 2])
    address = c1.text_input("Demo address", placeholder="1234 Main St, Marietta, GA 30060")
    demo_date = c2.date_input("Demo date", dt.date.today() + dt.timedelta(days=7))
    venue = c3.text_input("Venue name (optional)", placeholder="Cobb County Soccer Complex")
    go = st.form_submit_button("Find leads", type="primary")

if go:
    if not address.strip():
        st.error("Enter the demo address first.")
    else:
        frames = [read_csv(f.getvalue(), f.name) for f in uploads or []]
        with st.status("Finding leads, usually 1 to 2 minutes...", expanded=True) as s:
            try:
                res = core.find_leads(address, radius, n_leads, min_fields, frames, people, progress=s.write)
                s.update(label="Leads ready", state="complete", expanded=False)
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
unnamed, near = res["unnamed"], res["hubspot_near"]
for n in res.get("notes", []):
    st.info(n)
st.caption(f"Demo: {demo['label']}")
if leads.empty:
    st.warning(f"No named fields, schools, stadiums or clubs found within {meta['radius']} miles. "
               "Try a bigger radius or a lower minimum field count.")
    st.stop()

leads = leads.reset_index(drop=True)
leads["#"] = range(1, len(leads) + 1)
leads["opener"] = [opener(r, meta["date"], meta["venue"], demo["label"]) for _, r in leads.iterrows()]

m1, m2, m3, m4 = st.columns(4)
m1.metric("Places found nearby", res["n_found"])
m2.metric("Call list with a phone", f"{int((leads['phone'] != '').sum())} of {len(leads)}")
m3.metric("Named contacts", int((leads["contact_name"] != "").sum()))
m4.metric("HubSpot companies nearby", len(near) if res.get("hubspot_checked") else "No export")

tab_calls, tab_map, tab_hs, tab_all = st.tabs(["Call list", "Map", "HubSpot nearby", "Everything nearby"])

# ---------------------------------------------------------------- call list
with tab_calls:
    f1, f2, f3 = st.columns([2, 2, 1])
    types = f1.multiselect("Type", sorted(leads["type"].unique()), placeholder="All types")
    hs_opts = sorted(leads["hubspot"].unique())
    hs_pick = f2.multiselect("HubSpot status", hs_opts, placeholder="All statuses")
    phone_only = f3.toggle("Has a phone", value=False)
    view = leads
    if types:
        view = view[view["type"].isin(types)]
    if hs_pick:
        view = view[view["hubspot"].isin(hs_pick)]
    if phone_only:
        view = view[view["phone"] != ""]

    table = pd.DataFrame({
        "#": view["#"], "Name": view["company"], "Type": view["type"], "Miles": view["miles"].round(1),
        "Fields": view["units"].astype(int), "Phone": view["phone"],
        "Contact": [n or "" for n in view["contact_name"]], "HubSpot": view["hubspot"],
    })
    st.caption("Click a row to see contact details and the call opener.")
    pick = st.dataframe(table, hide_index=True, width="stretch", on_select="rerun", selection_mode="single-row",
                        column_config={"#": st.column_config.NumberColumn(width="small"),
                                       "Miles": st.column_config.NumberColumn(format="%.1f", width="small"),
                                       "Fields": st.column_config.NumberColumn(width="small")})
    rows = pick.selection.rows if pick and pick.selection else []
    if rows:
        r = view.iloc[rows[0]]
        warm = is_warm(r["hubspot"])
        addr = ", ".join(x for x in (r["street"], r["city"], r["zip"]) if x) or "Address not found"
        st.markdown(f"""<div class="lead-head{' warm' if warm else ''}">
            <h3>{r['#']}. {r['company']}</h3>
            <p>{r['type']}, {r['miles']:.1f} miles away, {int(r['units'])} mapped field{'s' if int(r['units']) != 1 else ''}
            {('(' + r['sports'] + ')') if r['sports'] else ''}</p><p>{addr}</p></div>""", unsafe_allow_html=True)
        a, b = st.columns(2)
        with a:
            st.markdown(f"**Phone:** {r['phone'] or 'not found'}")
            if r["contact_name"]:
                st.markdown(f"**Contact:** {r['contact_name']}, {r['contact_title']}"
                            + (f"  \n**Email:** {r['contact_email']}" if r["contact_email"] else ""))
            else:
                st.markdown(f"**Ask for:** {r['ask_for']}")
            if r["inbox"] and r["inbox"] != r["contact_email"]:
                st.markdown(f"**General email:** {r['inbox']}")
        with b:
            st.markdown(f'<span class="pill{" warm" if warm else ""}">{r["hubspot"]}</span>', unsafe_allow_html=True)
            if r["hubspot_match"]:
                st.caption(f"Matched on {r['hubspot_match']}")
            if r["buyer"] != r["company"]:
                st.markdown(f"**Buyer:** {r['buyer']}")
            if r["note"]:
                st.markdown(r["note"])
            links = [f"[Map]({r['maps_link']})"] + ([f"[Website]({r['website']})"] if r["website"] else [])
            st.markdown(" | ".join(links))
        st.code(r["opener"], language=None, wrap_lines=True)

    csv = pd.DataFrame({
        "#": leads["#"], "Name": leads["company"], "Type": leads["type"], "Miles": leads["miles"].round(1),
        "Fields": leads["units"], "Sports": leads["sports"], "Phone": leads["phone"],
        "Contact Name": leads["contact_name"], "Contact Title": leads["contact_title"],
        "Contact Email": leads["contact_email"], "General Email": leads["inbox"], "Ask For": leads["ask_for"],
        "Buyer": leads["buyer"], "Street": leads["street"], "City": leads["city"], "State": leads["state"],
        "Zip": leads["zip"], "Website": leads["website"], "HubSpot": leads["hubspot"],
        "HubSpot Match": leads["hubspot_match"], "Map": leads["maps_link"], "Call Opener": leads["opener"],
        "Demo Date": meta["date"].isoformat(), "Demo Address": demo["label"],
    }).to_csv(index=False)
    st.download_button("Download call list (CSV)", csv, file_name=f"demo_leads_{meta['date']:%Y-%m-%d}.csv",
                       mime="text/csv")

# ---------------------------------------------------------------- map
with tab_map:
    fmap = folium.Map(location=[demo["lat"], demo["lon"]], zoom_start=10, tiles="OpenStreetMap")
    folium.Circle([demo["lat"], demo["lon"]], radius=meta["radius"] * core.MILE_M, color=TMR_GREEN,
                  weight=2, fill=True, fill_opacity=0.03).add_to(fmap)
    folium.Marker([demo["lat"], demo["lon"]], tooltip="Demo location",
                  icon=folium.Icon(color="red", icon="star")).add_to(fmap)
    for _, r in others.iterrows():
        folium.CircleMarker([r["lat"], r["lon"]], radius=4, color=GRAY, fill=True, fill_opacity=0.7, weight=1,
                            tooltip=f"{r['name']} ({r['type']}, {r['miles']:.1f} mi)").add_to(fmap)
    for _, r in unnamed.iterrows():
        folium.RegularPolygonMarker([r["lat"], r["lon"]], number_of_sides=4, radius=6, color=GRAY,
                                    fill=True, fill_opacity=0.4, tooltip=r["label"]).add_to(fmap)
    for _, r in leads.iterrows():
        color = WARM if is_warm(r["hubspot"]) else CLUB if r["type"] == "Youth sports club" else TMR_GREEN
        folium.Marker(
            [r["lat"], r["lon"]], tooltip=f"{r['#']}. {r['company']} ({r['miles']:.1f} mi)",
            icon=folium.DivIcon(html=f'<div style="background:{color};color:#fff;border:2px solid #fff;'
                                     f'border-radius:50%;width:26px;height:26px;line-height:22px;text-align:center;'
                                     f'font:bold 13px Arial;box-shadow:0 1px 3px rgba(0,0,0,.4)">{r["#"]}</div>',
                                icon_size=(26, 26), icon_anchor=(13, 13)),
        ).add_to(fmap)
    st.markdown(f'<div class="legend"><span style="background:{TMR_GREEN};margin-left:0"></span>New lead'
                f'<span style="background:{CLUB}"></span>Youth club<span style="background:{WARM}"></span>In HubSpot'
                f'<span style="background:{GRAY}"></span>Other nearby (squares are unnamed fields)</div>',
                unsafe_allow_html=True)
    st_folium(fmap, height=520, use_container_width=True, returned_objects=[])

# ---------------------------------------------------------------- HubSpot nearby
with tab_hs:
    if not res.get("hubspot_checked"):
        st.info("Upload your HubSpot Companies and Contacts exports in the sidebar, then search again to see "
                "HubSpot companies near this demo.")
    elif near.empty:
        st.info(f"No HubSpot companies list a city within {meta['radius']} miles of the demo.")
    else:
        st.caption("Placed by the company's city, so distances are to the town center. "
                   "Never-worked companies are listed first.")
        h1, h2 = st.columns([2, 1])
        inds = h1.multiselect("Industry", sorted(i for i in near["industry"].unique() if i),
                              placeholder="All industries")
        never = h2.toggle("Never worked only", value=False)
        hv = near
        if inds:
            hv = hv[hv["industry"].isin(inds)]
        if never:
            hv = hv[hv["never_worked"]]
        st.dataframe(pd.DataFrame({
            "Company": hv["company"], "City": hv["city"], "Miles": hv["miles"], "Phone": hv["phone"],
            "Owner": hv["owner"], "Last activity": hv["last_activity"].dt.date, "Contacts": hv["n_contacts"],
            "Industry": hv["industry"]}), hide_index=True, width="stretch",
            column_config={"Last activity": st.column_config.DateColumn(format="MMM D, YYYY"),
                           "Miles": st.column_config.NumberColumn(format="%.1f")})
        with_people = hv[hv["n_contacts"] > 0]
        if not with_people.empty:
            st.markdown("**Contacts at these companies**")
            for _, r in with_people.head(40).iterrows():
                with st.expander(f"{r['company']}, {r['city']} ({r['n_contacts']})"):
                    st.dataframe(pd.DataFrame(r["contacts"]).rename(columns={
                        "name": "Name", "email": "Email", "phone": "Phone", "owner": "Owner",
                        "lead_status": "Lead status", "last_activity": "Last activity"}),
                        hide_index=True, width="stretch")
        st.download_button("Download HubSpot nearby (CSV)", hv.drop(columns=["contacts"]).to_csv(index=False),
                           file_name=f"hubspot_near_demo_{meta['date']:%Y-%m-%d}.csv", mime="text/csv")

# ---------------------------------------------------------------- everything nearby
with tab_all:
    st.caption("Every named place found, closest first. Useful for extra calls or HubSpot searches.")
    allp = pd.concat([leads.assign(on_list="Yes"), others.assign(on_list="")], ignore_index=True)
    st.dataframe(pd.DataFrame({
        "Name": allp["name"], "Type": allp["type"], "Miles": allp["miles"].round(1),
        "Fields": allp["units"].astype(int), "Phone": allp.get("phone", pd.Series([""] * len(allp))).fillna(""),
        "Note": allp["note"], "On call list": allp["on_list"]}).sort_values("Miles"),
        hide_index=True, width="stretch")
    if not unnamed.empty:
        st.markdown("**Unnamed field groups**")
        st.caption("Fields with no name on the map. Open the satellite view to identify them.")
        st.dataframe(pd.DataFrame({
            "Fields": unnamed["label"], "Miles": unnamed["miles"].round(1),
            "Satellite view": [f"https://www.google.com/maps/@{la},{lo},500m/data=!3m1!1e3"
                               for la, lo in zip(unnamed["lat"], unnamed["lon"])]}),
            hide_index=True, width="stretch",
            column_config={"Satellite view": st.column_config.LinkColumn(display_text="Open")})
