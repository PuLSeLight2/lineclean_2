# Render deployment — Venue Wastage App

## Before uploading
- Upload `app.py`, `requirements.txt`, and `.gitignore` to a **private** GitHub repository. Do not upload `venue_wastage.db` or passwords.
- Keep a safe backup of any existing local `venue_wastage.db` before moving online.
- Existing local database data is NOT transferred automatically by deploying the app. A new database is created on the disk unless you migrate it separately.

## Render web service
1. Create a new **Web Service** from your private GitHub repository.
2. Runtime: **Python**; choose a **paid** instance that supports persistent disks.
3. Build command: `pip install -r requirements.txt`
4. Start command: `streamlit run app.py --server.address 0.0.0.0 --server.port $PORT --server.headless true`
5. Add a **persistent disk** with mount path `/var/data` (select capacity in Render).
6. Add environment variable: `WASTAGE_DATA_DIR` = `/var/data`.
7. Deploy. Visit the HTTPS URL Render supplies and set up the first admin account.
8. Add a test record, restart/redeploy the service, and verify the record persists.

## Data safety
- Render only persists files written under the disk mount path. This package writes the SQLite database to `/var/data/venue_wastage.db` when configured.
- Use only **one** Render service instance with this SQLite disk.
- Keep separate SQLite-consistent backups; snapshots alone are not a substitute.
- Protect account credentials and grant venue access carefully before staff start using the public URL.
- If you have an existing local database, migrate it to the mounted disk before creating production accounts or records; otherwise you may need to merge data manually.

See https://render.com/docs/disks and https://render.com/docs/web-services for current Render instructions.
