"""
proxypics_probe_url.py — focused probe of api.proxypics.com only.
Filters out HTML responses (SPA noise) and surfaces real JSON answers.
"""
import json, os, sys, urllib.parse, urllib.request, urllib.error

DEFAULT_API_KEY = "9vZwjFnMGN1JYgERfCMGfqEA"  # paste here or set PROXYPICS_API_KEY

BASE = "https://api.proxypics.com"

# Paths most likely to exist on a public API rooted at /.
PATHS = [
    "/photo_requests", "/photo-requests",
    "/v1/photo_requests", "/v1/photo-requests",
    "/v2/photo_requests", "/v2/photo-requests",
    "/api/photo_requests", "/api/photo-requests",
    "/requests", "/v1/requests", "/v2/requests",
    "/draws", "/v1/draws", "/v2/draws",
    "/integrations/photo_requests",
    "/integrations/v1/photo_requests",
    "/external/photo_requests", "/external/v1/photo_requests",
    "/public/photo_requests", "/public/v1/photo_requests",
    "/me", "/v1/me",
    "/account", "/v1/account",
    "/openapi.json", "/swagger.json", "/api-docs",
]

AUTHS = [
    ("Bearer",       lambda k: {"Authorization": f"Bearer {k}"}),
    ("X-API-Key",    lambda k: {"X-API-Key": k}),
    ("X-Api-Key",    lambda k: {"X-Api-Key": k}),
    ("api-key",      lambda k: {"api-key": k}),
    ("access-token", lambda k: {"access-token": k}),
    ("Token token=", lambda k: {"Authorization": f"Token token={k}"}),
    ("no auth",      lambda k: {}),  # baseline — proves whether auth actually matters
]


def get_key():
    k = os.environ.get("PROXYPICS_API_KEY") or DEFAULT_API_KEY
    if not k: sys.exit("Set DEFAULT_API_KEY or PROXYPICS_API_KEY.")
    return k.strip()


def call(url, headers):
    req = urllib.request.Request(url, headers=headers, method="GET")
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            body = r.read()[:600].decode("utf-8", errors="replace")
            return r.status, dict(r.headers), body
    except urllib.error.HTTPError as e:
        body = e.read()[:400].decode("utf-8", errors="replace")
        return e.code, dict(e.headers), body
    except Exception as e:
        return None, {}, f"{type(e).__name__}: {e}"


def is_html(body, content_type):
    if "text/html" in (content_type or "").lower(): return True
    s = body.lstrip()[:50].lower()
    return s.startswith("<!doctype") or s.startswith("<html")


def main():
    key = get_key()
    headers_base = {"Accept": "application/json", "User-Agent": "proxypics-probe/0.3"}
    print(f"Probing {BASE} — JSON responses only (HTML ignored)\n")
    seen = []
    for path in PATHS:
        url = BASE + path
        for label, recipe in AUTHS:
            h = {**headers_base, **recipe(key)}
            status, resp_headers, body = call(url, h)
            if status is None:
                continue
            ct = resp_headers.get("Content-Type", "")
            if is_html(body, ct):
                continue
            # Now this is a real JSON / text response — show it.
            marker = "✓" if status == 200 else ("·" if status in (401, 403) else "?")
            short = body.replace("\n", " ").strip()[:140]
            print(f"  {marker} {status}  {path:36s}  ({label:14s})  → {short}")
            seen.append((status, path, label, body))
    if not seen:
        print("\nNo non-HTML responses anywhere. The key endpoints are NOT on api.proxypics.com.")
    else:
        # Highlight: any 200 + JSON containing array/object that LOOKS like data.
        data_hits = [s for s in seen if s[0] == 200 and ("[" in s[3] or '"id"' in s[3])]
        if data_hits:
            print("\n*** Likely real data endpoints:")
            for status, path, label, body in data_hits:
                print(f"   {BASE}{path}   auth: {label}")
