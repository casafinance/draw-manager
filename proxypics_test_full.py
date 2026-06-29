"""
proxypics_test_full.py — full round-trip test of the ProxyPics client.

What it does (in order):
  1. Tests connection (proves auth works).
  2. Lists your existing photo-requests (read-only, safe).
  3. Lists templates and finds the "draw inspection without floorplan" one.
  4. With --create:   creates a real test photo-request in SANDBOX.
  5. With --duplicate: creates one, then duplicates it (Draw 2 from Draw 1).
  6. With --cleanup:  cancels everything it just created.

SAFETY
------
Defaults to SANDBOX (https://sandbox.proxypics.com). You MUST also pass --live
to point at production, and even then the script refuses --create / --duplicate
against live unless you also pass --i-really-mean-live.

SANDBOX SETUP (one-time)
------------------------
1. Create an account at https://sandbox-app.proxypics.com
2. Go to Profile → Integrations → generate an API Key for sandbox
3. In ProxyPics sandbox, save a payment method:
   card 4111 1111 1111 1111, any CVV, any future expiry
4. Put the SANDBOX api key in DEFAULT_API_KEY below or PROXYPICS_API_KEY env var.

USAGE
-----
  py proxypics_test_full.py                          # connection + list only
  py proxypics_test_full.py --create                 # also creates 1 test request
  py proxypics_test_full.py --create --duplicate     # create + duplicate it
  py proxypics_test_full.py --create --cleanup       # create then cancel
  py proxypics_test_full.py --create --duplicate --cleanup   # full lifecycle
"""

import argparse
import json
import os
import sys
import time

from proxypics import ProxyPicsClient, ProxyPicsError


DEFAULT_API_KEY = "9vZwjFnMGN1JYgERfCMGfqEA"   # sandbox key (or set PROXYPICS_API_KEY env var)

# A safe, real-looking address for the test photo request.
TEST_ADDRESS = "1600 Amphitheatre Pkwy, Mountain View, CA 94043, USA"

# Substring used to find your draw-inspection template by name.
TEMPLATE_NAME_HINT = "draw inspection without floorplan"


def get_key() -> str:
    k = os.environ.get("PROXYPICS_API_KEY") or DEFAULT_API_KEY
    if not k:
        sys.exit(
            "ERROR: no API key set. Paste your SANDBOX key into DEFAULT_API_KEY "
            "at the top of this script, or set PROXYPICS_API_KEY in your env."
        )
    return k.strip()


def header(text: str) -> None:
    print(f"\n{'='*72}\n  {text}\n{'='*72}")


def pretty(obj, limit: int = 1600) -> None:
    s = json.dumps(obj, indent=2, default=str) if isinstance(obj, (dict, list)) else str(obj)
    if len(s) > limit:
        s = s[:limit] + f"\n… (truncated, {len(s)-limit} more chars)"
    print(s)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--live", action="store_true",
                    help="Use the LIVE ProxyPics API (default is sandbox)")
    ap.add_argument("--i-really-mean-live", action="store_true",
                    help="Required guard alongside --live for write operations")
    ap.add_argument("--create", action="store_true",
                    help="Create a real test photo-request")
    ap.add_argument("--duplicate", action="store_true",
                    help="After creating, duplicate it (mimics Draw 2 from Draw 1)")
    ap.add_argument("--cleanup", action="store_true",
                    help="Cancel anything this script created at the end")
    ap.add_argument("--address", default=TEST_ADDRESS,
                    help=f"Address for the test request (default: {TEST_ADDRESS!r})")
    args = ap.parse_args()

    use_sandbox = not args.live
    if args.live and (args.create or args.duplicate) and not args.i_really_mean_live:
        sys.exit("Refusing to write to LIVE without --i-really-mean-live as well.")

    client = ProxyPicsClient(api_key=get_key(), sandbox=use_sandbox)

    # ---- 1. Connection test ------------------------------------------------
    header(f"1. Connection test  ({'SANDBOX' if use_sandbox else 'LIVE'})")
    try:
        info = client.test_connection()
        print(f"  ✓ Authenticated. base_url={info['base_url']}")
        print(f"  Total photo-requests on this account: {info['total_requests']}")
    except ProxyPicsError as e:
        print(f"  ✗ FAILED: {e}")
        return 1

    # ---- 2. List existing --------------------------------------------------
    header("2. List existing photo-requests (first 5)")
    try:
        listing = client.list_photo_requests(page=1, per_page=5)
        data = (listing.get("data") if isinstance(listing, dict) else listing) or []
        print(f"  Got {len(data)} item(s).")
        for r in data[:5]:
            if isinstance(r, dict):
                print(f"    #{r.get('id')}  status={r.get('status'):<12}  "
                      f"loan={r.get('loan_number')!r:<10}  {r.get('small_address')!r}")
    except ProxyPicsError as e:
        print(f"  ✗ FAILED: {e}")
        return 1

    # ---- 3. Find the draw-inspection template -------------------------------
    header(f"3. Find template matching {TEMPLATE_NAME_HINT!r}")
    template_token = None
    try:
        tpl = client.find_template_by_name(TEMPLATE_NAME_HINT)
        if tpl:
            template_token = tpl.get("token")
            print(f"  ✓ Found: {tpl.get('name')!r}")
            print(f"    id={tpl.get('id')}  token={template_token}")
            print(f"    platform={tpl.get('photo_request_platform')}  "
                  f"tasks={len(tpl.get('tasks') or [])}")
        else:
            print(f"  ✗ No template matched {TEMPLATE_NAME_HINT!r}. "
                  "Set up a Draw Inspection template in ProxyPics first, "
                  "or change TEMPLATE_NAME_HINT at the top of this script.")
            if args.create or args.duplicate:
                return 1
    except ProxyPicsError as e:
        print(f"  ✗ FAILED: {e}")
        if args.create or args.duplicate:
            return 1

    # ---- 4. Create ---------------------------------------------------------
    created_ids: list = []
    new_id = None
    if args.create:
        if not template_token:
            print("\nCannot --create without a template token. Skipping.")
        else:
            header(f"4. CREATE a test photo-request at {args.address!r}")
            try:
                new_req = client.create_photo_request(
                    address=args.address,
                    template_token=template_token,
                    loan_number="TEST-PP-API",
                    additional_notes="Created by Draw Manager integration test. Safe to ignore.",
                    photo_request_platform="crowdsource",
                )
                new_id = new_req.get("id") if isinstance(new_req, dict) else None
                if new_id:
                    created_ids.append(new_id)
                    print(f"  ✓ Created request #{new_id}")
                    print(f"    address: {new_req.get('address')!r}")
                    print(f"    status:  {new_req.get('status')!r}")
                    print(f"    cost:    {new_req.get('cost')}  (cents)")
                    print(f"    expires: {new_req.get('expires_at')!r}")
                else:
                    print("  ? Response had no id field.")
                    pretty(new_req, limit=600)
            except ProxyPicsError as e:
                print(f"  ✗ CREATE FAILED: {e}")

    # ---- 5. Duplicate ------------------------------------------------------
    dup_id = None
    if args.duplicate:
        parent = new_id
        if not parent:
            # Fall back to the most recent existing request.
            try:
                listing = client.list_photo_requests(page=1, per_page=1)
                data = (listing.get("data") if isinstance(listing, dict) else listing) or []
                if data and isinstance(data[0], dict):
                    parent = data[0].get("id")
                    print(f"\n  (no --create result; using most recent existing #{parent})")
            except ProxyPicsError as e:
                print(f"  ✗ Couldn't find a parent to duplicate: {e}")

        if parent:
            header(f"5. DUPLICATE photo-request #{parent} (Draw N from Draw N-1)")
            try:
                # Sandbox is sometimes briefly inconsistent right after a create.
                time.sleep(1.0)
                dup = client.duplicate_photo_request(parent)
                dup_id = dup.get("id") if isinstance(dup, dict) else None
                if dup_id:
                    created_ids.append(dup_id)
                    print(f"  ✓ Duplicated → new request #{dup_id}")
                    print(f"    address: {dup.get('address')!r}")
                    print(f"    status:  {dup.get('status')!r}")
                else:
                    print("  ? Duplicate response had no id field.")
                    pretty(dup, limit=600)
            except ProxyPicsError as e:
                print(f"  ✗ DUPLICATE FAILED: {e}")

    # ---- 6. Cleanup --------------------------------------------------------
    if args.cleanup and created_ids:
        header(f"6. CLEANUP — cancelling {len(created_ids)} test request(s)")
        for rid in created_ids:
            try:
                resp = client.cancel_photo_request(rid)
                status = resp.get("status") if isinstance(resp, dict) else "?"
                print(f"  ✓ Cancelled #{rid}  (final status: {status!r})")
            except ProxyPicsError as e:
                print(f"  ✗ Cancel #{rid} failed: {e}")
                print("    (If it wasn't 'unassigned' yet, ProxyPics support has been notified.)")

    header("DONE")
    if created_ids and not args.cleanup:
        print(f"  Created request id(s): {created_ids}  "
              f"(re-run with --cleanup to cancel them).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
