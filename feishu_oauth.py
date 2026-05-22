#!/usr/bin/env python3
"""
One-time Feishu OAuth setup — run locally to obtain FEISHU_USER_REFRESH_TOKEN.

Prerequisites:
  1. In Feishu developer console → your app → 安全设置 → 重定向URL
     add:  http://127.0.0.1:8080/callback
  2. Run this script with env vars set:
     Windows:
       set FEISHU_APP_ID=cli_xxx && set FEISHU_APP_SECRET=xxx && python feishu_oauth.py
     macOS/Linux:
       FEISHU_APP_ID=cli_xxx FEISHU_APP_SECRET=xxx python feishu_oauth.py
  3. Copy the printed refresh_token → add as GitHub Secret FEISHU_USER_REFRESH_TOKEN
"""
import json
import os
import urllib.parse
import urllib.request
import urllib.error
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer

APP_ID = os.environ.get("FEISHU_APP_ID", "")
APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "")
REDIRECT_URI = "http://127.0.0.1:8080/callback"
FEISHU_API = "https://open.feishu.cn/open-apis"

_code = []


class _Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        params = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        if "code" in params:
            _code.append(params["code"][0])
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"<h1>Authorization successful! Close this tab.</h1>")
        else:
            self.send_response(400)
            self.end_headers()
            error = params.get("error", ["unknown"])[0]
            self.wfile.write(f"<h1>Error: {error}</h1>".encode())

    def log_message(self, *args):
        pass


def main():
    if not APP_ID or not APP_SECRET:
        print("ERROR: FEISHU_APP_ID and FEISHU_APP_SECRET must be set as environment variables.")
        return

    auth_url = (
        f"{FEISHU_API}/authen/v1/authorize"
        f"?app_id={APP_ID}"
        f"&redirect_uri={urllib.parse.quote(REDIRECT_URI, safe='')}"
        f"&scope={urllib.parse.quote('sheets:spreadsheet:readonly', safe='')}"
        f"&state=ctr_setup"
    )
    print(f"Opening browser for Feishu authorization...")
    print(f"(If browser doesn't open, visit this URL manually:)\n{auth_url}\n")
    webbrowser.open(auth_url)

    server = HTTPServer(("127.0.0.1", 8080), _Handler)
    print("Waiting for callback on http://127.0.0.1:8080 (timeout: 3 min)...")
    server.timeout = 180
    while not _code:
        server.handle_request()
    server.server_close()

    code = _code[0]
    print(f"Got authorization code.")

    # Exchange code for tokens (v2 OIDC endpoint)
    url = f"{FEISHU_API}/authen/v2/oauth/token"
    body = json.dumps({
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "client_id": APP_ID,
        "client_secret": APP_SECRET,
    }).encode()
    req = urllib.request.Request(url, data=body, method="POST",
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req) as r:
            resp = json.loads(r.read())
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="replace")
        print(f"Token exchange failed: HTTP {e.code} — {err_body[:500]}")
        return

    refresh_token = resp.get("refresh_token")
    access_token = resp.get("access_token")

    if refresh_token:
        print(f"\n{'='*60}")
        print("SUCCESS! Add the following value to GitHub Secrets")
        print("Secret name:  FEISHU_USER_REFRESH_TOKEN")
        print(f"Secret value: {refresh_token}")
        print(f"{'='*60}\n")
        print(f"(access_token also obtained, expires in {resp.get('expires_in', '?')}s)")
    else:
        print(f"No refresh_token in response: {resp}")


if __name__ == "__main__":
    main()
