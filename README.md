# WAN NOC Portal

A self-hosted dashboard for monitoring a 17-node WAN with OFC links, alternate media (microwave, satellite, BBR, BTS, copper, cellular), terminal equipment, radio equipment, and fault logs.

## Features

- **Dashboard** — KPI tiles, per-node connectivity state (Full / Degraded / Isolated), recent faults
- **Topology** — Geographic map (Leaflet) and logical network graph (vis-network), toggle between them, filter by media type and status
- **Links & Media** — Full CRUD for OFC links and alternate media, search/filter
- **Equipment** — Terminal and radio equipment inventory with age/lifecycle highlighting
- **Faults** — Log new faults, track open/resolved, severity filters, MTTR computation
- **Analytics** — Charts for fault categories, severity, timeline, OFC loss-vs-margin scatter, equipment age distribution
- **Daily Report** — Auto-generated communication state report, printable to PDF
- **Import/Export** — Excel template download, file upload (replace or append modes), full data export

## Quick Start

```bash
# 1. Install Python 3.9+ if not present
# 2. Install dependencies
pip install -r requirements.txt

# 3. Run
python app.py
```

Open `http://localhost:5000` in your browser.

The app auto-creates `instance/wan_portal.db` (SQLite) and seeds it with realistic mock data for 17 Indian cities on first run.

## Deployment on a Standalone Server

### Option A — Direct run with auto-start (Linux)

Create a systemd service file at `/etc/systemd/system/wan-portal.service`:

```ini
[Unit]
Description=WAN NOC Portal
After=network.target

[Service]
Type=simple
User=youruser
WorkingDirectory=/opt/wan_portal
ExecStart=/usr/bin/python3 /opt/wan_portal/app.py
Restart=always

[Install]
WantedBy=multi-user.target
```

Then:

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now wan-portal
```

### Option B — Production WSGI (recommended)

```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

### Option C — Bundle as a single executable

```bash
pip install pyinstaller
pyinstaller --onefile --add-data "templates:templates" --add-data "static:static" app.py
# Output binary in dist/app — copy to your standalone server
```

### Option D — Offline deployment (no internet on server)

The portal uses a few CDN libs (Leaflet, vis-network, Chart.js). To run fully offline:

1. Download these files on a machine with internet:
   - `https://unpkg.com/leaflet@1.9.4/dist/leaflet.js` and `leaflet.css`
   - `https://unpkg.com/vis-network@9.1.9/standalone/umd/vis-network.min.js`
   - `https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js`
2. Place them under `static/vendor/` and update the `<script>` and `<link>` tags in `templates/base.html`, `templates/dashboard.html`, and `templates/analytics.html` to point at `{{ url_for('static', filename='vendor/...') }}`
3. Fonts: the app already uses Arial (system font), so no font download is needed.
4. **For the map**: upload your offline OSM map as an `.mbtiles` file via **Import / Export → Offline Map**. The dashboard will switch to your local tiles automatically.

#### Creating a North-India MBTiles file

Several ways to do this:

**Easiest — MapTiler Desktop / TileMill:**
   - Download a North-India OSM extract from <https://download.geofabrik.de/asia/india.html> or use a sub-region clip
   - Open in MapTiler/TileMill, set bounds to roughly lat 26°–37°N, lng 72°–82°E
   - Export as MBTiles, zoom levels 5–12 (good balance of detail vs. file size)

**Command-line — `tilemaker` (open-source):**
   ```bash
   tilemaker --input india-latest.osm.pbf --output north-india.mbtiles \
     --bbox 72,26,82,37 --config tilemaker/resources/config-openmaptiles.json
   ```

**From an existing tile set — `mb-util`:**
   ```bash
   pip install mbutil
   mb-util tiles_folder north-india.mbtiles --scheme=xyz
   ```

Once you have the file, just upload it on the **Import / Export** page. No restart needed — refresh the dashboard and you'll see your offline map. A green "OFFLINE MAP" badge appears in the top-right of the topology stage.

The mbtiles file is stored in `instance/tiles/active.mbtiles` — you can also drop it there directly via SCP/file copy.

## Data Format

Use **Import/Export → Download Template** to get a blank Excel with these sheets:

| Sheet | Columns |
|---|---|
| Nodes | name, latitude, longitude, region, node_type, remarks |
| OFC_Links | link_id, from_node, to_node, distance_km, year_laid, no_of_fiber, loss_db, no_dark_fiber, cable_type, margin_db, status, remarks |
| Alternate_Media | media_type, from_node, to_node, spec, hop_distance_km, status, remarks |
| Terminal_Equipment | equipment_id, location, eqpt_type, eth_ports, e1_voice_ports, capacity, year_purchased, status, remarks |
| Radio_Equipment | equipment_id, radio_type, location, frequency, year_purchased, status, remarks |
| Fault_Logs | fault_id, category, affected_link, affected_node, reported_at, resolved_at, severity, description, action_taken, status |

**Status conventions:**
- OFC / Alt Media: `UP`, `DEGRADED`, `DOWN`
- Equipment: `OPERATIONAL`, `FAULTY`, `EOL`
- Faults: `OPEN`, `RESOLVED` (severity: `HIGH`, `MEDIUM`, `LOW`)

Datetime fields use ISO format: `2025-05-04T14:30:00`

## How Connectivity States Are Computed

For each of the 17 nodes:
- **FULL** — at least 2 OFC links UP, OR (1 OFC UP AND 1 alt media UP)
- **DEGRADED** — at least 1 path UP/degraded but not enough redundancy
- **ISOLATED** — no UP path

You can adjust this logic in `compute_node_states()` in `app.py`.

## Customizing for Your Real Data

1. Run the app once to create the DB schema
2. Go to **Import/Export → Reset to Mock Data** is optional — you can skip it
3. Wipe the mock data: **Import/Export → Upload Excel** with your real file in **Replace** mode

Or programmatically:

```bash
rm instance/wan_portal.db   # wipe everything
python app.py               # re-creates schema, seeds mock
# then upload your real Excel via the UI
```

## Security Notes

- The app has no authentication built in. For production, put it behind nginx with HTTP basic auth, or add Flask-Login.
- Change the `app.secret_key` in `app.py` before exposing to a network.
- SQLite is fine for single-server deployments. For multi-user heavy writes, switch to PostgreSQL by changing the connection logic in `get_db()`.

## File Structure

```
wan_portal/
├── app.py                  # Flask app, all routes, DB logic
├── requirements.txt
├── README.md
├── instance/
│   └── wan_portal.db       # auto-created SQLite database
├── static/
│   └── css/style.css       # all styling
└── templates/
    ├── base.html           # nav + layout
    ├── dashboard.html      # landing page
    ├── topology.html       # geographic + logical map
    ├── links.html          # OFC + alt media tables
    ├── equipment.html      # terminal + radio equipment
    ├── faults.html         # fault log
    ├── analytics.html      # charts
    ├── daily_report.html   # printable report
    └── import_export.html  # Excel I/O
```

## Upload size limits (PMTiles / MBTiles maps)

The default upload cap is **2 GB**, which covers most PMTiles regional builds.
To change it, set the `MAX_UPLOAD_MB` environment variable before starting the
server:

```bash
# Linux / macOS
MAX_UPLOAD_MB=4096 python app.py

# Windows PowerShell
$env:MAX_UPLOAD_MB="4096"; python app.py
```

If you're seeing a **413 Payload Too Large** error and the file is below the cap
shown on the Import/Export page, you are almost certainly running behind a
reverse proxy (nginx, Apache, IIS) that has its own upload limit. The proxy
caps are independent of Flask's and need to be raised separately:

| Proxy | Setting | Where |
|---|---|---|
| **nginx** | `client_max_body_size 2g;` | inside the `server` or `location` block |
| **Apache** | `LimitRequestBody 2147483648` | in the `<Location>` or `<Directory>` block |
| **IIS** | `maxAllowedContentLength` and `maxRequestEntityAllowed` | in `web.config` |
| **Caddy** | `request_body { max_size 2GB }` | in your site block |

Restart the proxy after the change.

## License

Use freely. No warranty.
