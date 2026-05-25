"""
Minimal ArcGIS Enterprise OAuth2 Login Example
===============================================

PURPOSE
-------

This example demonstrates a VERY SIMPLE browser-based OAuth login flow
for ArcGIS Enterprise / Portal for ArcGIS.

The script:

1. Opens the ArcGIS Enterprise OAuth login page in the browser
2. User authenticates interactively
3. ArcGIS redirects back to a LOCAL temporary HTTP server
4. The script extracts the OAuth access token
5. The token is used to access a protected ImageServer resource

This is the MODERN recommended authentication model for:
- ArcGIS Enterprise
- Federated ArcGIS Server
- Portal for ArcGIS
- OAuth-secured ImageServer services

This avoids:
- manually handling passwords
- storing credentials
- dealing with expiring session cookies manually

-----------------------------------------------------------------------

IMPORTANT REQUIREMENTS
----------------------

Before running this script:

1. Register an OAuth Application in Portal for ArcGIS

Portal:
    Content
      -> New Item
      -> Application
      -> Register

2. Configure the redirect URI EXACTLY as:

    http://localhost:8765/callback

3. Copy the generated CLIENT ID

-----------------------------------------------------------------------

INSTALLATION
------------

Requires:

    pip install requests python-dotenv

-----------------------------------------------------------------------

TESTED CONCEPTS
---------------

This example demonstrates:

- OAuth2 Implicit Grant flow
- Browser popup authentication
- Automatic token retrieval
- Protected ArcGIS REST access
- DEM/ImageServer integration

-----------------------------------------------------------------------
"""

import http.server
import os
import socketserver
import threading
import time
import urllib.parse
import webbrowser

import requests
from dotenv import load_dotenv

# ---------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------

ENV_FILE = ".env"
load_dotenv(ENV_FILE)


def get_required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise ValueError(
            f"Missing required environment variable: {name}. "
            "Set it in the .env file before running this script."
        )
    return value


def write_env_value(name: str, value: str, env_file: str = ENV_FILE) -> None:
    lines = []
    if os.path.exists(env_file):
        with open(env_file, "r", encoding="utf-8") as f:
            lines = f.readlines()

    updated = False
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("#") or "=" not in stripped:
            continue
        key, _ = stripped.split("=", 1)
        if key == name:
            lines[i] = f"{name}={value}\n"
            updated = True
            break

    if not updated:
        if lines and not lines[-1].endswith("\n"):
            lines[-1] = lines[-1] + "\n"
        lines.append(f"{name}={value}\n")

    with open(env_file, "w", encoding="utf-8") as f:
        f.writelines(lines)


# Your Portal URL
PORTAL_URL = get_required_env("ARCGIS_PORTAL_URL")

# OAuth Application Client ID
CLIENT_ID = get_required_env("ARCGIS_CLIENT_ID")

# Local redirect URI
# MUST MATCH the registered URI in Portal
REDIRECT_URI = get_required_env("ARCGIS_REDIRECT_URI")

# Example protected ArcGIS ImageServer
IMAGE_SERVER_URL = get_required_env("ARCGIS_IMAGE_SERVER_URL")

# Global token state. Initialized from .env and updated after OAuth login.
oauth_token = os.getenv("ARCGIS_OAUTH_TOKEN")

# ---------------------------------------------------------------------
# LOCAL HTTP SERVER
# ---------------------------------------------------------------------
#
# ArcGIS OAuth will redirect the browser to:
#
#     http://localhost:8765/callback#access_token=....
#
# Unfortunately:
# - the token appears after "#"
# - browsers DO NOT send fragment parts to the server
#
# Therefore:
# - we serve a tiny HTML page
# - JavaScript extracts the fragment
# - JavaScript sends token back to Python
#
# ---------------------------------------------------------------------


class OAuthHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):

        global oauth_token

        parsed = urllib.parse.urlparse(self.path)

        # -------------------------------------------------------------
        # STEP 1:
        # User is redirected here after login
        #
        # Example:
        #
        #   /callback#access_token=....
        #
        # BUT:
        #   Python server never sees fragment after "#"
        #
        # So we return HTML+JavaScript to extract it.
        # -------------------------------------------------------------

        if parsed.path == "/callback":
            html = """
            <html>
            <body>
                <h2>ArcGIS OAuth Login Successful</h2>

                <p>You may close this window.</p>

                <script>

                // ----------------------------------------------------
                // Extract token from URL fragment
                // ----------------------------------------------------

                const fragment = window.location.hash.substring(1);

                const params = new URLSearchParams(fragment);

                const token = params.get("access_token");

                // ----------------------------------------------------
                // Send token back to Python server
                // ----------------------------------------------------

                fetch("/store_token?token=" + token)
                    .then(() => {
                        console.log("Token sent to Python.");
                    });

                </script>
            </body>
            </html>
            """

            self.send_response(200)
            self.send_header("Content-type", "text/html")
            self.end_headers()

            self.wfile.write(html.encode("utf-8"))

            return

        # -------------------------------------------------------------
        # STEP 2:
        # Browser JavaScript sends token here
        # -------------------------------------------------------------

        if parsed.path == "/store_token":
            query = urllib.parse.parse_qs(parsed.query)

            token = query.get("token", [None])[0]

            if token:
                oauth_token = token
                write_env_value("ARCGIS_OAUTH_TOKEN", token)
                print("\n[+] OAuth token received and stored in .env.\n")
            else:
                print("\n[-] Token missing in callback.\n")

            self.send_response(200)
            self.end_headers()

            self.wfile.write(b"OK")

            return

        # -------------------------------------------------------------
        # Unknown paths
        # -------------------------------------------------------------

        self.send_response(404)
        self.end_headers()


# ---------------------------------------------------------------------
# START LOCAL HTTP SERVER
# ---------------------------------------------------------------------

PORT = urllib.parse.urlparse(REDIRECT_URI).port or 8765

httpd = socketserver.TCPServer(("localhost", PORT), OAuthHandler)

server_thread = threading.Thread(target=httpd.serve_forever)

server_thread.daemon = True

server_thread.start()

print(f"[+] Local callback server started on port {PORT}")

# ---------------------------------------------------------------------
# BUILD OAUTH LOGIN URL
# ---------------------------------------------------------------------
#
# response_type=token
#
# means:
# - implicit flow
# - access token returned directly
#
# ---------------------------------------------------------------------

oauth_url = (
    f"{PORTAL_URL}/sharing/rest/oauth2/authorize"
    f"?client_id={CLIENT_ID}"
    f"&response_type=token"
    f"&redirect_uri={urllib.parse.quote(REDIRECT_URI)}"
)

if oauth_token:
    print("\n[+] ARCGIS_OAUTH_TOKEN found in .env. Reusing existing token.\n")
else:
    print("\n[+] Opening browser for ArcGIS login...\n")

    # -----------------------------------------------------------------
    # OPEN BROWSER LOGIN WINDOW
    # -----------------------------------------------------------------

    webbrowser.open(oauth_url)

    # -------------------------------------------------------------
    # WAIT FOR TOKEN
    # -------------------------------------------------------------

    print("[+] Waiting for OAuth authentication...")

    while oauth_token is None:
        time.sleep(1)

print("\n[+] Authentication complete.")

print("\n[+] Access Token:")
print(oauth_token)

# ---------------------------------------------------------------------
# TEST PROTECTED IMAGE SERVER ACCESS
# ---------------------------------------------------------------------
#
# We now call the protected ArcGIS ImageServer
# using the OAuth token.
#
# ---------------------------------------------------------------------

test_url = f"{IMAGE_SERVER_URL}?f=pjson&token={oauth_token}"

print("\n[+] Testing authenticated ImageServer access...\n")

response = requests.get(test_url, timeout=30)

print(f"HTTP Status: {response.status_code}")

print("\nResponse:\n")

print(response.text)

# ---------------------------------------------------------------------
# OPTIONAL:
# TEST exportImage
# ---------------------------------------------------------------------

export_url = (
    "https://mapas.pd.anatel.gov.br/image/rest/services/"
    "DEM/ImageServer/exportImage"
    "?bbox=-40.8,-19.8,-40.7,-19.7"
    "&bboxSR=4326"
    "&imageSR=4326"
    "&size=512,512"
    "&format=tiff"
    "&f=pjson"
    f"&token={oauth_token}"
)

print("\n[+] Testing authenticated exportImage...\n")

response = requests.get(export_url, timeout=30)

print(f"HTTP Status: {response.status_code}")

print("\nResponse:\n")

print(response.text)

# ---------------------------------------------------------------------
# CLEANUP
# ---------------------------------------------------------------------

httpd.shutdown()

print("\n[+] Finished.")
