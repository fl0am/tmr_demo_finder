# TMR Demo Lead Finder

Enter a booked demo address and get the 5-10 best nearby fields, schools and
stadiums to call, with a map, phone numbers, any contacts found on their
websites, a ready-made call opener, and a HubSpot warm-lead flag.

All data is free: OpenStreetMap for places and addresses, and each place's own
website for phones, emails and names. No Google key.

## Use it
1. Open the app link.
2. (Optional) Upload your HubSpot Companies/Contacts export CSVs in the sidebar.
3. Type the demo address and date, click **Find leads**. Takes 1-2 minutes.

HubSpot files are only used for that search. They are never saved or committed.

## Run it on your own computer
    pip install -r requirements.txt
    streamlit run app.py

## Files
- `app.py` - the web page
- `demo_finder_core.py` - the search logic
- `complex_finder_v3.py` - shared code from the statewide complex finder
