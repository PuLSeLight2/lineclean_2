VENUE WASTAGE APP — ADMIN & MULTI-VENUE EDITION

RUN LOCALLY
1. Install Python 3.10+.
2. Open a terminal in this folder.
3. Run: pip install -r requirements.txt
4. Run: streamlit run app.py
5. On first launch, create the first Admin account.

PHONE
For phone access on the same Wi-Fi, run:
  streamlit run app.py --server.address 0.0.0.0
Open the Network URL shown in the terminal on your phone.
For access from different locations/devices, deploy to a host with persistent storage and HTTPS.

ROLES
- Admin: full access; add/edit/deactivate venues; create accounts; assign venues; set roles and individual permissions.
- Manager: can record wastage, view history and manage stock/value lists for assigned venues by default.
- Team / Till: can record line-clean and general wastage by default. Admin can enable/disable individual permissions such as history or stock access.

VENUES
- Starts with Hotshots, Rosso, Centre and Reds.
- Admin can add as many additional venues as required, rename them or deactivate them.
- Each venue has separate stock lists, values and wastage records.

WASTAGE
- Line clean: Coors, Madri, Stowford, Carling, Carling Darkfruits, Guinness and Alpaca by default.
- Each tap = 5 pints; first and second pull-throughs shown separately by product.
- General wastage categories: Draught, Bottle beers, Soft bottle, Postmix, Spirits, Wine.
- Admin/Managers with stock permission can add/edit/deactivate items and set venue-specific units and values.

DATABASE
- Data is stored in venue_wastage.db beside app.py.
- The app attempts to keep existing data when replacing an earlier version that used the same database filename.
- Back up venue_wastage.db regularly. Cloud hosting must use persistent storage or data can be lost on restart.
- Passwords are stored as salted PBKDF2-HMAC-SHA256 hashes.


Dashboard and Excel update:
- Dashboard is available to Admin and any user granted History permission. It shows totals, daily trends, product/category breakdowns, team activity and recent entries, filtered by venue and period.
- Excel (.xlsx) history and dashboard exports separate Date and Time.
- General waste Quantity is the number of measures entered. Example: 2 x 25ml Jack = Quantity 2, Measure / unit 25ml, Total volume (ml) 50.
- Existing venue_wastage.db is retained when updating the app; back it up before replacing files.
