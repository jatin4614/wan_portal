"""Production launcher: serves the Flask app via Waitress.

Behind IIS + HttpPlatformHandler, IIS launches this process and tells it which
local port to listen on via the HTTP_PLATFORM_PORT environment variable, then
reverse-proxies public traffic to that port. We MUST bind to that port (not a
hardcoded one) or IIS proxies to a dead port and the request hangs / 502s.
When run by hand (no IIS) it falls back to 5000.
"""
import os
from waitress import serve
from app import app

if __name__ == "__main__":
    # IIS HttpPlatformHandler assigns the port; fall back to 5000 for manual runs.
    port = int(os.environ.get("HTTP_PLATFORM_PORT", "5000"))
    print(f"Waitress serving the WAN Portal on 127.0.0.1:{port}", flush=True)
    # Listen only on localhost; IIS proxies public traffic to us.
    serve(
        app,
        host="127.0.0.1",
        port=port,
        threads=8,                # tune based on concurrent users
        max_request_body_size=17_179_869_184,  # 16 GB (same as Flask cap)
        channel_timeout=600,      # 10 min for slow large uploads
    )