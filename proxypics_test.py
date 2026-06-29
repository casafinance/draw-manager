"""
proxypics_test.py — verify the ProxyPics v3 API works with your API key.

Now using the REAL documented endpoint and auth:
  - Base URL: https://api.proxypics.com/api/v3 (live)
              https://sandbox.proxypics.com/api/v3 (sandbox)
  - Auth:    header "x-api-key: <your key>"

SETUP
-----
1. Paste your API key into DEFAULT_API_KEY below, or set PROXYPICS_API_KEY env var.
2. (Optional) flip USE_SANDBOX to True for testing against the sandbox.
3. Run:  python proxypics_test.py

This script does READ-ONLY GETs:
  - GET /photo-requests          (list, page 1)
  - GET /photo-requests/<id>     (detail of the first one)
  - GET /photo-request-templates (your templates / products)

Use --address "fragment" to filter by address.
Use --id 1234567 to fetch a specific request.
Use --templates to list your templates (useful — we'll need a template_token to create).
Use --shape to print a structure summary.
"""

import argparse
import json
import os
import sys
import urllib.parse
import urllib.request
import urllib.error


DEFAULT_API_KEY = "9vZwjFnMGN1JYgERfCMGfqEA"              # paste here or set PROXYPICS_API_KEY
USE_SANDBOX     = False           # True -> sandbox.proxypics.com (safe for testing)

LIVE_BASE    = "https://api.proxypics.com/api/v3"
SANDBOX_BASE = "https://sandbox.proxypics.com/api/v3"


def base_url():
    return SANDBOX_BASE if USE_SANDBOX else LIVE_BASE


def get_api_key():
    k = os.environ.get("PROXYPICS_API_KEY") or DEFAULT_API_KEY
    if not k:
        sys.exit(
            "ERROR: no API key set.\n"
            "Paste your key into DEFAULT_API_KEY at the top of this script "
            "or set PROXYPICS_API_KEY in your environment.")
    return k.strip()


def call(method, path, params=None, body=None):
    url = base_url() + path
    if params:
        qs = urllib.parse.urlencode({k: v for k, v in params.items() if v is not None})
        if qs:
            url = f"{url}?{qs}"
    headers = {
        "Accept":       "application/json",
        "Content-Type": "application/json",
        "User-Agent":   "proxypics-test/1.0",
        "x-api-key":    get_api_key(),
    }
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, headers=headers, method=method, data=data)
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            raw = r.read().decode("utf-8", errors="replace")
            try:    return r.status, json.loads(raw)
            except: return r.status, raw
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", errors="replace")
        try:    return e.code, json.loads(raw)
        except: return e.code, raw


def show(label, status, body):
    print(f"\n{'='*72}\n{label}\nHTTP {status}\n{'-'*72}")
    if isinstance(body, (dict, list)):
        print(json.dumps(body, indent=2, default=str)[:6000])
    else:
        print(str(body)[:2000])


def describe(obj, prefix="", maxd=4):
    if maxd <= 0:
        print(prefix + "…"); return
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, dict):
                print(f"{prefix}{k}: object"); describe(v, prefix + "  ", maxd - 1)
            elif isinstance(v, list):
                print(f"{prefix}{k}: list[{len(v)}]")
                if v: describe(v[0], prefix + "  [0]: ", maxd - 1)
            else:
                print(f"{prefix}{k}: {type(v).__name__} = {repr(v)[:60]}")
    elif isinstance(obj, list):
        print(f"{prefix}list[{len(obj)}]")
        if obj: describe(obj[0], prefix + "  [0]: ", maxd - 1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--address", help="Filter by address fragment")
    ap.add_argument("--id", help="Fetch a specific photo-request ID")
    ap.add_argument("--templates", action="store_true",
                    help="List your photo request templates (products)")
    ap.add_argument("--shape", action="store_true",
                    help="Print field-type summary instead of raw JSON")
    args = ap.parse_args()

    print(f"Base URL: {base_url()}")
    print(f"Auth:     x-api-key (key length {len(get_api_key())})\n")

    if args.templates:
        status, body = call("GET", "/photo-request-templates", params={"page": 1, "per_page": 50})
        show("LIST /photo-request-templates", status, body)
        if args.shape and isinstance(body, (dict, list)):
            print("\n--- shape ---"); describe(body)
        return

    # 1. LIST photo-requests
    params = {"page": 1, "per_page": 20}
    if args.address:
        # The v3 API may use different filter syntax than v2's `filters` blob.
        # Try common variants — the server ignores unknown params, so safe.
        params["address"] = args.address
        params["q"] = args.address
    status, body = call("GET", "/photo-requests", params=params)
    show("LIST /photo-requests" + (f' (search "{args.address}")' if args.address else ""),
         status, body)
    if args.shape and isinstance(body, (dict, list)):
        print("\n--- shape ---"); describe(body)

    # 2. SHOW one — either --id or the first from the list
    first_id = None
    if isinstance(body, dict):
        data = body.get("data") or []
        if isinstance(data, list) and data and isinstance(data[0], dict):
            first_id = data[0].get("id")
    target = args.id or first_id
    if target:
        status, body = call("GET", f"/photo-requests/{target}")
        show(f"GET /photo-requests/{target}", status, body)
        if args.shape and isinstance(body, (dict, list)):
            print("\n--- shape ---"); describe(body)
    else:
        print("\n(No items in list to drill into. Try --address with a known fragment.)")


if __name__ == "__main__":
    main()
