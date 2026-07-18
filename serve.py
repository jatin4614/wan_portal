"""Production launcher: serves the Flask app via Waitress."""
from waitress import serve
from app import app

if __name__ == "__main__":
    # Listen only on localhost; IIS will proxy public traffic to us
    serve(
        app,
        host="127.0.0.1",
        port=5000,
        threads=8,                # tune based on concurrent users
        max_request_body_size=17_179_869_184,  # 16 GB (same as Flask cap)
        channel_timeout=600,      # 10 min for slow large uploads
    )