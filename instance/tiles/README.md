# Map tiles

Place your offline basemap here as one of:

- `active.pmtiles` — PMTiles (vector or raster). Preferred.
- `active.mbtiles` — MBTiles (raster, SQLite).

The tile files themselves are **not** committed to the repository (they can be
several GB). The application serves whatever tile file is present here via
`/api/tiles/info`, `/tiles/<z>/<x>/<y>.png`, and `/api/pmtiles`.

If no tile file is present, the dashboard falls back to online tiles.
