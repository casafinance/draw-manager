"""
fci_funding_probe.py — dump the RAW funding-history fields FCI returns for a loan,
so we know the real field names (instead of guessing percentOwned/currentBalance).

USAGE:
  1. Paste your FCI API key into API_KEY below (or set FCI_API_KEY env var).
  2. Find the loan account number for 850 SW 5th St (from your sheet/Casa) and
     put it in LOAN_ACCOUNT, OR pass it: py fci_funding_probe.py <loan_account>
  3. Run it. It prints every field on every investor row.

This is READ-ONLY.
"""
import json, os, sys
import casa_logic as L

API_KEY = ""          # paste FCI key here or set FCI_API_KEY
LOAN_ACCOUNT = ""     # put the loan account # for 850 SW 5th St here, or pass as arg


def main():
    key = os.environ.get("FCI_API_KEY") or API_KEY
    if not key:
        # try reading from the casa config
        try:
            cfg = L.load_config()
            key = cfg.get("fci_api_key", "")
        except Exception:
            pass
    if not key:
        sys.exit("Set API_KEY at top, or FCI_API_KEY env, or have it in config.json")

    loan = (sys.argv[1] if len(sys.argv) > 1 else LOAN_ACCOUNT).strip()
    if not loan:
        sys.exit("Provide the loan account #: py fci_funding_probe.py <loan_account>")

    # First, introspect what fields getFundingHistory actually exposes.
    print(f"=== Probing funding history for loan {loan} ===\n")

    # Ask for a broad set of likely field names; FCI ignores unknown ones or errors.
    # We'll try them one at a time to discover which exist.
    candidate_fields = [
        "lenderAccount", "percentOwned", "percentageOwned", "ownership",
        "ownershipPercent", "pctOwned", "percent", "investorPercent",
        "currentBalance", "principalBalance", "originalBalance",
        "amountFunded", "fundedAmount", "amount", "investmentAmount",
        "lenderName", "investorName", "name", "isEnabled", "enabled",
        "propertyCode", "investorRate", "rate",
    ]
    working = []
    for f in candidate_fields:
        q = f'query Q($la:String!){{getFundingHistory(loanaccount:$la){{lenderAccount {f}}}}}'
        try:
            data = L.fci_api_call(key, q, {"la": loan})
            rows = data.get("getFundingHistory")
            if isinstance(rows, list):
                working.append(f)
        except Exception:
            pass  # field doesn't exist
    print("Fields that EXIST on getFundingHistory:")
    print("  " + ", ".join(working) + "\n")

    # Now fetch all working fields together and dump every row.
    fieldset = " ".join(dict.fromkeys(["lenderAccount"] + working))
    q = f'query Q($la:String!){{getFundingHistory(loanaccount:$la){{{fieldset}}}}}'
    data = L.fci_api_call(key, q, {"la": loan})
    rows = data.get("getFundingHistory", [])
    print(f"=== {len(rows)} investor row(s) ===")
    for i, r in enumerate(rows, 1):
        print(f"\nRow {i}:")
        print(json.dumps(r, indent=2, default=str))


if __name__ == "__main__":
    main()
