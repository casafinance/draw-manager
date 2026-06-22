"""
ProxyPics draw-inspection request automation.

Flow:
    1. Prompt for a property address (or pass --address "...").
    2. Look up the loan in Casa Finance (loans → falls back to servicing).
    3. Extract CF# and the Borrower / Guarantor contact rows.
    4. Ask which contact to use.
    5. Open ProxyPics and fill the new-draw-inspection form.
    6. Stop after clicking "Next" (does NOT submit).

First run on each site will need a manual login; sessions persist in
./.pw-profile so subsequent runs go straight through.
"""

from __future__ import annotations

import argparse
import re
import sqlite3
import sys
import time
from pathlib import Path
from playwright.sync_api import (
    sync_playwright, Page, TimeoutError as PWTimeout, Locator,
)

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
PROXYPICS_URL    = "https://app.proxypics.com/photo-requests/draw-inspections/new"
CASA_LOANS_URL   = "https://myportal.casa.finance/admin/loans"
CASA_SERV_URL    = "https://myportal.casa.finance/admin/servicing"

_SCRIPT_DIR      = (Path(sys.executable).parent
                    if getattr(sys, "frozen", False)
                    else Path(__file__).parent)
PROFILE_DIR      = _SCRIPT_DIR / ".pw-profile"   # may be overridden by --profile-dir
LOGIN_WAIT_S     = 300

# Sentinels used to detect "page is ready, not on login".
PROXYPICS_READY  = 'button:has-text("Add contact")'
CASA_READY       = 'input[placeholder="Search"]'

DEFAULTS = {
    "inspection_type_search": "draw inspection without floorplan",
    "csv_path":               "",
    "progress_step":          "25",
}


# ===========================================================================
# Generic helpers
# ===========================================================================
def digits_only(s: str) -> str:
    """Strip everything but digits; drop leading 1 if 11 digits long."""
    d = re.sub(r"\D", "", s or "")
    if len(d) == 11 and d.startswith("1"):
        d = d[1:]
    return d


def wait_for_ready(page: Page, target_url: str, ready_selector: str,
                   label: str) -> None:
    """Navigate to target_url; poll until ready_selector is visible.
    Handles login redirects by periodically re-navigating to target_url."""
    print(f">>> [{label}] Navigating to {target_url}")
    try:
        page.goto(target_url, wait_until="domcontentloaded", timeout=30_000)
    except Exception as e:
        print(f"    (initial goto warning: {e})")

    deadline       = time.time() + LOGIN_WAIT_S
    next_renav_at  = 0.0
    next_status_at = 0.0
    target_path    = target_url.split("//", 1)[1].split("/", 1)[1]

    while time.time() < deadline:
        try:
            if page.locator(ready_selector).count() > 0:
                page.locator(ready_selector).first.wait_for(
                    state="visible", timeout=5000
                )
                print(f">>> [{label}] Page ready.")
                return
        except PWTimeout:
            pass
        except Exception as e:
            print(f"    (poll warning: {e})")

        now = time.time()
        if now >= next_status_at:
            try:
                print(f"    [{label}] waiting…  url={page.url}")
            except Exception:
                print(f"    [{label}] waiting…")
            next_status_at = now + 5

        if target_path not in page.url and now >= next_renav_at:
            next_renav_at = now + 8
            try:
                page.goto(target_url, wait_until="domcontentloaded",
                          timeout=15_000)
            except Exception:
                pass

        page.wait_for_timeout(1000)

    shot = Path(__file__).parent / f"timeout_{label}.png"
    try:
        page.screenshot(path=str(shot), full_page=True)
        print(f">>> Saved screenshot to {shot}")
    except Exception:
        pass
    raise TimeoutError(
        f"[{label}] page not ready within {LOGIN_WAIT_S}s. "
        f"Last URL: {page.url}."
    )


# ===========================================================================
# Local DB fast path
# ===========================================================================
# Casa Finance address search is brittle: minor formatting differences, closed
# loans hiding on the wrong tab, and React detail-page race conditions all make
# it fail intermittently. The Draw Manager app already maintains a synced
# sqlite DB with column A (CF#) → column C (address) mappings for every loan
# we know about. We resolve address → CF# locally first, then drive the Casa
# search by CF# (which is unique and deterministic) instead of address text.

_ADDR_PUNCT = re.compile(r"[,\.]")
_ADDR_WS    = re.compile(r"\s+")

def _norm_addr(s: str) -> str:
    """Loose normalization so '2307 Chestnut St, Houston TX' matches
    '2307 chestnut st' — lowercase, strip punctuation, collapse whitespace."""
    if not s:
        return ""
    s = s.lower().strip()
    s = _ADDR_PUNCT.sub(" ", s)
    s = _ADDR_WS.sub(" ", s)
    return s.strip()


def _db_lookup_cf(db_path: str, address: str) -> str | None:
    """Look up a CF# in the local Draw Manager DB by fuzzy address match.
    Returns the CF# (e.g. 'CF008') or None. Read-only, never raises."""
    if not db_path or not Path(db_path).exists():
        return None
    needle = _norm_addr(address)
    if not needle:
        return None
    try:
        # `mode=ro` so we never lock the file the app might also be reading.
        conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
        conn.row_factory = sqlite3.Row
        try:
            rows = conn.execute(
                "SELECT cf_number, address FROM loans "
                "WHERE address IS NOT NULL AND address != ''"
            ).fetchall()
        finally:
            conn.close()
    except Exception as e:
        print(f"    (DB lookup warning: {e})")
        return None

    # Three-tier match, most-strict first. Returning early on first hit means
    # closer matches win even when the DB has multiple loans on the same street.
    normalized = [(r["cf_number"], _norm_addr(r["address"])) for r in rows]

    # 1. Exact normalized equality.
    for cf, a in normalized:
        if a == needle:
            return cf
    # 2. Stored address starts with what the user typed
    #    (user gave street; DB has 'street, city, st zip').
    for cf, a in normalized:
        if a.startswith(needle + " "):
            return cf
    # 3. What the user typed starts with the stored address
    #    (user pasted in extras; DB has just the street).
    for cf, a in normalized:
        if a and needle.startswith(a + " "):
            return cf
    return None


# ===========================================================================
# Casa Finance lookup
# ===========================================================================
def casa_lookup(page: Page, address: str, cf_hint: str | None = None) -> dict:
    """Find the loan in Casa, return {cf_number, contacts: [...]}.
    If cf_hint is provided, Casa is searched by CF# first (unique, deterministic)
    before falling back to address text."""
    for section_url, label in [(CASA_LOANS_URL,  "casa-loans"),
                               (CASA_SERV_URL,   "casa-servicing")]:
        result = _try_casa_section(page, section_url, label, address, cf_hint)
        if result is not None:
            return result
    raise RuntimeError(f"Could not find a loan for: {address!r}")


def _try_casa_section(page: Page, section_url: str, label: str,
                      address: str, cf_hint: str | None = None) -> dict | None:
    """Search Casa by address, click the matching row, extract CF# + contacts.
    If cf_hint is provided (from the local DB), it OVERRIDES whatever the
    page extraction would produce — Casa's search box only accepts addresses,
    so we still search by address, but we can skip the brittle CF# extraction
    step when we already know the answer."""
    wait_for_ready(page, section_url, CASA_READY, label)

    queries = [address]
    street_only = address.split(",", 1)[0].strip()
    if street_only and street_only != address:
        queries.append(street_only)

    for query in queries:
        # Make sure we're back on the section page (we may have navigated to a
        # row's detail on a previous attempt).
        if section_url.rsplit("/", 1)[-1] not in page.url:
            print(f"    Returning to {section_url}")
            try:
                page.goto(section_url, wait_until="domcontentloaded")
                page.locator(CASA_READY).first.wait_for(
                    state="visible", timeout=15_000
                )
            except Exception:
                pass

        print(f">>> [{label}] Searching: {query!r}")
        if not _casa_search_and_click(page, query):
            continue

        # Wait for the loan detail to be visibly populated before extracting.
        contacts_tab_present = True
        try:
            page.locator('[data-testid="Contacts"]').first.wait_for(
                state="visible", timeout=15_000
            )
        except PWTimeout:
            contacts_tab_present = False
            print("    Loan detail didn't surface a Contacts tab within 15s.")

        # CF# resolution: prefer the DB-known value (cf_hint) over scraping
        # the page. The hint is what made the address resolvable in the first
        # place, so we trust it.
        if cf_hint:
            cf = cf_hint
            print(f">>> [{label}] Using DB-known CF# (skipped page extraction): {cf}")
        else:
            try:
                cf = _extract_cf_number(page)
            except Exception as e:
                shot = Path(__file__).parent / f"cf_extract_failed_{label}.png"
                try:
                    page.screenshot(path=str(shot), full_page=True)
                except Exception:
                    pass
                print(f"    Couldn't read CF# ({e}). Screenshot: {shot}")
                # Same row won't work better next loop — bail out of this section.
                return None
            print(f">>> [{label}] Found CF#: {cf}")

        # If the Contacts tab never appeared, don't sit on a hung .click() —
        # short-circuit with an empty contact list and let the caller decide.
        if not contacts_tab_present:
            print("    Skipping contacts extraction (no Contacts tab).")
            contacts: list[dict] = []
        else:
            contacts = _extract_contacts(page)
        return {"cf_number": cf, "contacts": contacts}
    return None


def _casa_search_and_click(page: Page, query: str) -> bool:
    """Type into the search box, wait for results, click the first row.
    Returns True if a row was clicked, False otherwise."""
    search = page.locator('input[placeholder="Search"]').first
    search.click()
    # Clear any previous text.
    search.fill("")
    search.fill(query)
    # Give the table a moment to refresh.
    page.wait_for_timeout(1500)

    rows = page.locator('tr[data-row-id]')
    n = rows.count()
    print(f"    {n} row(s) returned.")
    if n == 0:
        return False

    rows.first.click()
    # SPA route — domcontentloaded won't re-fire. Wait for a sentinel that
    # only exists on the loan-detail page (breadcrumb-1 holds the address).
    try:
        page.locator('[data-testid="breadcrumb-1"]').first.wait_for(
            state="visible", timeout=15_000
        )
    except PWTimeout:
        print("    Detail page didn't render breadcrumb in 15s.")
        return False
    # Small settle so React has a beat to populate values.
    page.wait_for_timeout(800)
    return True


CF_REGEX = re.compile(r"\bCF\s*\d+\b", re.IGNORECASE)


def _normalize_cf(text: str) -> str | None:
    """Pull a CF### out of `text`, normalize to e.g. 'CF630'."""
    if not text:
        return None
    cleaned = text.replace("\xa0", " ")
    m = CF_REGEX.search(cleaned)
    if not m:
        return None
    return re.sub(r"\s+", "", m.group(0)).upper()


def _extract_cf_number(page: Page) -> str:
    """Find the loan's CF# on the detail page. Tries several strategies."""
    try:
        page.wait_for_load_state("networkidle", timeout=8000)
    except PWTimeout:
        pass

    # Strategy 1: structural — the page uses CSS-modules with the pattern
    #   <div class="_summary_xxx"><div class="_key_xxx">…label…</div>
    #                            <div class="_value_xxx">CF772</div></div>
    # The hashes change but the prefixes don't.
    xpath_candidates = [
        # Element containing label → climb to summary container → grab value child
        '//*[contains(normalize-space(.), "Casa Finance File Number")]'
        '/ancestor::*[contains(@class, "_summary_")][1]'
        '//*[contains(@class, "_value_")]',
        # Key div with label → next sibling div (the value)
        '//div[contains(@class, "_key_") and .//*[contains(text(), "Casa Finance File Number")]]'
        '/following-sibling::div[1]',
        # Same but label is direct text, not nested
        '//div[contains(@class, "_key_") and contains(normalize-space(.), "Casa Finance File Number")]'
        '/following-sibling::div[1]',
        # Original fallbacks (in case some pages still use contenteditable)
        '//*[contains(normalize-space(.), "Casa Finance File Number")]'
        '/following::*[@contenteditable][1]',
        '//*[contains(normalize-space(.), "File Number")]'
        '/following::*[contains(@class, "_value_")][1]',
    ]
    for xp in xpath_candidates:
        try:
            loc = page.locator(f"xpath={xp}").first
            loc.wait_for(state="attached", timeout=3000)
            cf = _normalize_cf(loc.text_content() or "")
            if cf:
                return cf
        except Exception:
            continue

    # Strategy 2: brute-force — scan every `_value_*` div on the page.
    try:
        values = page.locator('[class*="_value_"]')
        for i in range(min(values.count(), 60)):
            cf = _normalize_cf(values.nth(i).text_content() or "")
            if cf:
                return cf
    except Exception:
        pass

    # Strategy 3: full body text regex sweep.
    try:
        body_text = page.locator("body").inner_text(timeout=5000)
        cf = _normalize_cf(body_text)
        if cf:
            return cf
    except Exception:
        pass

    raise RuntimeError("No CF# found anywhere on the loan detail page.")


def _extract_contacts(page: Page) -> list[dict]:
    """Click Contacts tab, read Borrower and Guarantor rows.
    Returns empty list if the Contacts tab isn't on the page — never hangs
    waiting for a click to resolve on an absent element."""
    tab = page.locator('[data-testid="Contacts"]')
    if tab.count() == 0:
        print("    No Contacts tab present — returning empty contacts list.")
        return []
    tab.first.click()
    page.wait_for_timeout(1200)

    contacts: list[dict] = []
    for role in ["Borrower", "Guarantor"]:
        # A contact row whose Role cell contains the role text.
        row = page.locator(
            f'tr[data-row-id]:has(td[data-label="Role"]:has-text("{role}"))'
        ).first
        if row.count() == 0:
            print(f"    No {role} row found.")
            continue
        name  = _cell_text(row, "Name")
        email = _cell_text(row, "Email")
        phone = _cell_text(row, "Phone")
        contacts.append({
            "role":  role,
            "name":  name,
            "email": "" if email in ("", "-") else email,
            "phone": "" if phone in ("", "-") else phone,
        })
    return contacts


def _cell_text(row: Locator, data_label: str) -> str:
    cell = row.locator(f'td[data-label="{data_label}"]')
    if cell.count() == 0:
        return ""
    return (cell.first.text_content() or "").strip()


# ===========================================================================
# ProxyPics form filling
# ===========================================================================
def fill_proxypics(page: Page, data: dict, submit: bool = False) -> None:
    wait_for_ready(page, PROXYPICS_URL, PROXYPICS_READY, "proxypics")

    # Each step is wrapped so one missing field/selector can't abort the whole
    # run — for a demo we want it to fill what it can and still reach "Next".
    def step(label, fn):
        print(f">>> {label}")
        try:
            fn()
        except Exception as e:
            print(f"    (skipped {label!r}: {type(e).__name__}: {e})")

    def _inspection_type():
        field = page.locator(".ant-select-auto-complete").first
        field.click()
        field.locator("input.ant-select-selection-search-input").fill(
            data["inspection_type_search"]
        )
        page.locator(".ant-select-item-option-content").first.wait_for(
            state="visible", timeout=10_000
        )
        page.locator(".ant-select-item-option-content").first.click()
    step("Filling: inspection type", _inspection_type)

    step("Filling: loan number",
         lambda: page.locator('input[id$="_loanNumber"]').fill(data["loan_number"]))

    def _address():
        addr = page.locator('input[id$="_address"]')
        addr.click()
        addr.fill(data["address"])
        try:
            page.wait_for_selector(".pac-item", timeout=5000)
            page.keyboard.press("ArrowDown")
            page.keyboard.press("Enter")
        except PWTimeout:
            pass
    step("Filling: address", _address)

    # Contact: only add one if we actually have any contact data. With no
    # contacts on the loan we skip this entirely rather than creating a blank.
    has_contact = any((data.get("contact_name"), data.get("contact_phone"),
                       data.get("contact_email")))
    if has_contact:
        def _contact():
            page.get_by_role("button", name="Add contact").click()
            if data.get("contact_name"):
                page.get_by_placeholder("Contact Name").fill(data["contact_name"])
            if data.get("contact_phone"):
                phone_input = page.locator('input[type="tel"]').last
                phone_input.click()
                phone_input.press("End")
                phone_input.type(data["contact_phone"])
            if data.get("contact_email"):
                page.get_by_placeholder("Email address").fill(data["contact_email"])
        step("Filling: contact", _contact)
    else:
        print(">>> No contact info — skipping the Add contact step.")

    if data.get("csv_path"):
        def _csv():
            with page.expect_file_chooser() as fc_info:
                page.get_by_role("button", name="Import line items").click()
            fc_info.value.set_files(data["csv_path"])
        step("Uploading: line items CSV", _csv)
    else:
        print(">>> No line-items CSV provided — skipping import.")

    def _progress():
        wrapper = page.locator('input[id$="_drawProgressStep"]').locator(
            'xpath=ancestor::div[contains(@class,"ant-select")][1]'
        )
        wrapper.click()
        page.locator(f'.ant-select-item-option[title="{data["progress_step"]}"]').click()
    step("Selecting: progress step", _progress)

    step("Clicking: Next",
         lambda: page.get_by_role("button", name="Next").click())

    if submit:
        def _submit():
            # The Submit button on the review screen:
            #   <button type="submit" class="... ant-btn-primary"><span>Submit</span></button>
            btn = page.get_by_role("button", name="Submit", exact=True)
            try:
                btn.first.wait_for(state="visible", timeout=20_000)
            except Exception:
                btn = page.locator('button[type="submit"].ant-btn-primary'
                                   ':has(span:text-is("Submit"))')
                btn.first.wait_for(state="visible", timeout=10_000)
            btn.first.click()
        step("Clicking: Submit (LIVE)", _submit)
        print(">>> LIVE submit attempted.")
    else:
        print(">>> TEST MODE — form filled and Next clicked; stopping before Submit.")


# ===========================================================================
# CLI
# ===========================================================================
def choose_contact(contacts: list[dict], preselect: str = "") -> dict:
    if not contacts:
        # No contacts on the loan — don't abort. Return an empty placeholder so
        # the run still proceeds, fills everything else, and stops at Next.
        print(">>> No contacts found on the loan — continuing without one.")
        return {"role": "", "name": "", "email": "", "phone": ""}

    preselect = (preselect or "").strip().lower()

    # Explicit numeric index
    if preselect.isdigit():
        i = int(preselect)
        if 1 <= i <= len(contacts):
            c = contacts[i - 1]
            print(f">>> Auto-picked #{i}: {c['role']} — {c['name']}")
            return c
        print(f"    --contact-choice {preselect} out of range; falling through to auto.")
        preselect = "auto"

    # Auto-pick: prefer one with both email AND phone, else first with either,
    # else just the first contact.
    if preselect == "auto":
        for c in contacts:
            if c["email"] and c["phone"]:
                print(f">>> Auto-picked (has email+phone): {c['role']} — {c['name']}")
                return c
        for c in contacts:
            if c["email"] or c["phone"]:
                print(f">>> Auto-picked (has some contact info): {c['role']} — {c['name']}")
                return c
        print(f">>> Auto-picked first: {contacts[0]['role']} — {contacts[0]['name']}")
        return contacts[0]

    # Interactive
    print("\nContacts on file:")
    for i, c in enumerate(contacts, 1):
        print(f"  {i}. {c['role']:9s} {c['name']:30s}  "
              f"email={c['email'] or '—':30s}  phone={c['phone'] or '—'}")
    while True:
        raw = input(f"Choose contact [1-{len(contacts)}]: ").strip()
        if raw.isdigit() and 1 <= int(raw) <= len(contacts):
            return contacts[int(raw) - 1]
        print("  ??? try again.")


def ensure_chromium_installed() -> None:
    """Verify Playwright's Chromium is on disk; install it if not.
    Handles both dev mode (sys.executable is python) and the frozen .exe
    (where we have to invoke playwright's bundled driver directly because
    `python -m` isn't available)."""
    import os, subprocess
    # Common install locations playwright checks in order.
    candidates = [
        os.environ.get("PLAYWRIGHT_BROWSERS_PATH"),
        os.path.expandvars(r"%LOCALAPPDATA%\ms-playwright"),
        os.path.expanduser("~/.cache/ms-playwright"),
        os.path.expanduser("~/Library/Caches/ms-playwright"),
    ]
    for d in candidates:
        if d and os.path.isdir(d):
            try:
                if any(n.startswith("chromium") for n in os.listdir(d)):
                    return  # already installed
            except OSError:
                pass

    print(">>> Playwright Chromium not found on this machine.")
    print(">>> Installing now (one-time, ~150 MB download)…")

    if getattr(sys, "frozen", False):
        # Inside the PyInstaller .exe — sys.executable is OUR exe, not python.
        # Invoke playwright's bundled driver instead.
        try:
            from playwright._impl._driver import compute_driver_executable, get_driver_env
            driver = compute_driver_executable()
            cmd = list(driver) if isinstance(driver, (list, tuple)) else [str(driver)]
            env = get_driver_env() if callable(get_driver_env) else os.environ
            subprocess.run(cmd + ["install", "chromium"], env=env, check=True)
        except Exception as e:
            print(f"ERROR: auto-install failed inside frozen exe: {e}")
            print("Manual fix: open PowerShell and run:")
            print("    python -m playwright install chromium")
            sys.exit(2)
    else:
        # Dev mode — sys.executable IS python.
        try:
            subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"],
                           check=True)
        except Exception as e:
            print(f"ERROR: auto-install failed: {e}")
            print("Manual fix: run `python -m playwright install chromium`")
            sys.exit(2)
    print(">>> Chromium installed. Continuing.")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--address", help="Property address (prompted if omitted)")
    ap.add_argument("--csv",     default=DEFAULTS["csv_path"])
    ap.add_argument("--step",    default=DEFAULTS["progress_step"])
    ap.add_argument("--type",    default=DEFAULTS["inspection_type_search"],
                    dest="insp_type")
    ap.add_argument("--contact-choice", default="",
                    help="'auto', or 1-based index. Empty = prompt.")
    ap.add_argument("--profile-dir", default="",
                    help="Browser user data directory (for persistent login).")
    ap.add_argument("--db", default="",
                    help="Path to Draw Manager's draws.db for the local "
                         "address→CF# fast path (skips brittle Casa scraping).")
    ap.add_argument("--submit", action="store_true",
                    help="LIVE mode: actually click Submit at the end. Without "
                         "this flag the run fills everything and stops at Next.")
    ap.add_argument("--headless", action="store_true",
                    help="Run the browser with no visible window.")
    args = ap.parse_args()

    # Allow override of the browser profile location (used when launched from
    # the Draw Manager app so dev-mode and exe-mode share a profile if desired).
    global PROFILE_DIR
    if args.profile_dir:
        PROFILE_DIR = Path(args.profile_dir)

    address = (args.address or input("Property address: ")).strip()
    if not address:
        print("No address — exiting.")
        sys.exit(1)

    # The line-items CSV is optional. If a path is given but missing, warn and
    # carry on without it rather than aborting a session that may burn a login.
    if args.csv and not Path(args.csv).exists():
        print(f"WARNING: draw-items CSV not found, continuing without it: {args.csv}")
        args.csv = ""

    ensure_chromium_installed()

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=args.headless,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        print(f">>> Looking up: {address!r}")

        # Local fast path: look the address up in the Draw Manager DB so we
        # can SKIP the brittle CF#-extraction step on the loan-detail page.
        # Casa search itself still uses the address (its search box only
        # accepts addresses, not CF#s).
        cf_hint: str | None = None
        if args.db:
            cf_hint = _db_lookup_cf(args.db, address)
            if cf_hint:
                print(f">>> Local DB match: {cf_hint} — will skip page CF# extraction.")
            else:
                print(">>> No local DB match — will extract CF# from the Casa page.")
        else:
            print(">>> No --db provided — will extract CF# from the Casa page.")

        info = casa_lookup(page, address, cf_hint=cf_hint)
        contact = choose_contact(info["contacts"], preselect=args.contact_choice)

        data = {
            "inspection_type_search": args.insp_type,
            "loan_number":            info["cf_number"],
            "address":                address,
            "contact_name":           contact["name"],
            "contact_phone":          digits_only(contact["phone"]),
            "contact_email":          contact["email"],
            "csv_path":               args.csv,
            "progress_step":          args.step,
        }
        print(f"\n>>> {'LIVE — will SUBMIT' if args.submit else 'TEST — stops before Submit'}")
        print(f">>> Submitting to ProxyPics with:")
        for k, v in data.items():
            print(f"      {k}: {v}")
        print()

        fill_proxypics(page, data, submit=args.submit)

        print(">>> Done. Close the browser window when you're finished reviewing.")
        try:
            page.wait_for_event("close", timeout=0)
        except Exception:
            pass
        ctx.close()


if __name__ == "__main__":
    main()