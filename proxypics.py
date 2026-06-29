"""
proxypics.py — thin ProxyPics API v3 client.

Auth: x-api-key header. JSON in, JSON out.
Endpoints used:
  GET    /photo-requests                      list (paginated)
  GET    /photo-requests/{id}                 show
  POST   /photo-requests                      create (new draw)
  POST   /photo-requests/{id}/duplicate       duplicate (new Draw N from Draw N-1)
  DELETE /photo-requests/{id}                 cancel (only safe when 'unassigned')
  GET    /photo-request-templates             list templates (products)

Used both by the standalone test harness and by Draw Manager's app.py.
"""

import json
import urllib.parse
import urllib.request
import urllib.error
from typing import Any, Optional


LIVE_BASE    = "https://api.proxypics.com/api/v3"
SANDBOX_BASE = "https://sandbox.proxypics.com/api/v3"


class ProxyPicsError(Exception):
    """Raised when the API returns a non-2xx response."""
    def __init__(self, status: int, body: Any, url: str):
        self.status = status
        self.body = body
        self.url = url
        # Pull useful detail from the body if it's the documented error shape.
        msg = body
        if isinstance(body, dict):
            errs = body.get("errors")
            if errs:
                msg = errs if isinstance(errs, str) else "; ".join(map(str, errs))
            else:
                msg = body.get("message") or body
        super().__init__(f"HTTP {status} @ {url}: {msg}")


class ProxyPicsClient:
    """ProxyPics REST client. Stateless — one instance per session/process is fine."""

    def __init__(self, api_key: str, sandbox: bool = False, timeout: float = 30.0):
        if not api_key or not api_key.strip():
            raise ValueError("ProxyPics API key is required.")
        self.api_key = api_key.strip()
        self.sandbox = bool(sandbox)
        self.base_url = SANDBOX_BASE if self.sandbox else LIVE_BASE
        self.timeout = timeout

    # ---- core request plumbing ------------------------------------------------
    def _request(self, method: str, path: str, *,
                 params: Optional[dict] = None,
                 body: Optional[dict] = None) -> Any:
        url = self.base_url + path
        if params:
            clean = {k: v for k, v in params.items() if v is not None}
            if clean:
                url = f"{url}?{urllib.parse.urlencode(clean)}"
        headers = {
            "Accept":       "application/json",
            "Content-Type": "application/json",
            "User-Agent":   "draw-manager-proxypics/1.0",
            "x-api-key":    self.api_key,
        }
        data = json.dumps(body).encode("utf-8") if body is not None else None
        req = urllib.request.Request(url, headers=headers, method=method, data=data)
        try:
            with urllib.request.urlopen(req, timeout=self.timeout) as r:
                raw = r.read().decode("utf-8", errors="replace")
                try:    return json.loads(raw)
                except json.JSONDecodeError: return raw
        except urllib.error.HTTPError as e:
            raw = e.read().decode("utf-8", errors="replace")
            try:    body_obj = json.loads(raw)
            except json.JSONDecodeError: body_obj = raw
            raise ProxyPicsError(e.code, body_obj, url) from None

    # ---- connection test -------------------------------------------------------
    def test_connection(self) -> dict:
        """Light call that proves auth works. Returns a small summary."""
        data = self._request("GET", "/photo-requests", params={"page": 1, "per_page": 1})
        if not isinstance(data, dict):
            raise ProxyPicsError(200, data, self.base_url + "/photo-requests")
        meta = (data.get("meta") or {}).get("pagination") or {}
        return {
            "ok": True,
            "sandbox": self.sandbox,
            "base_url": self.base_url,
            "total_requests": meta.get("total"),
        }

    # ---- photo-requests --------------------------------------------------------
    def list_photo_requests(self, page: int = 1, per_page: int = 20) -> dict:
        return self._request("GET", "/photo-requests",
                             params={"page": page, "per_page": per_page})

    def get_photo_request(self, request_id) -> dict:
        return self._request("GET", f"/photo-requests/{request_id}")

    def find_by_loan_number(self, loan_number: str, max_pages: int = 10) -> Optional[dict]:
        """Search all your requests for one whose loan_number matches.
        The API doesn't expose a documented filter, so we paginate client-side.
        Returns the MOST RECENT match (highest id), or None.

        max_pages caps how far back we'll look — at per_page=50 that's 500 records."""
        target = (loan_number or "").strip()
        if not target:
            return None
        newest = None
        for page in range(1, max_pages + 1):
            resp = self.list_photo_requests(page=page, per_page=50)
            data = (resp or {}).get("data") or []
            if not data:
                break
            for r in data:
                if not isinstance(r, dict):
                    continue
                if (r.get("loan_number") or "").strip() == target:
                    if newest is None or (r.get("id") or 0) > (newest.get("id") or 0):
                        newest = r
            # Stop early if we've passed the last page per pagination meta.
            meta = ((resp or {}).get("meta") or {}).get("pagination") or {}
            total = meta.get("total")
            cur   = meta.get("current_page", page)
            per   = meta.get("per_page", 50)
            if total is not None and cur * per >= total:
                break
        return newest

    def create_photo_request(self, *,
                             address: str,
                             template_token: str,
                             loan_number: Optional[str] = None,
                             borrower_name: Optional[str] = None,
                             contacts: Optional[list] = None,
                             additional_notes: Optional[str] = None,
                             external_id: Optional[str] = None,
                             expires_at: Optional[str] = None,
                             photo_request_platform: str = "crowdsource",
                             extra: Optional[dict] = None) -> dict:
        """Create a new photo request from a template.

        `contacts` is a list of dicts in the API's contacts_attributes shape,
        e.g. [{"name": "Jane", "number": "+1 555-…", "contact_type": "point_of_contact"}].
        `extra` is merged in last for forward-compat with new API fields.
        """
        if not address:
            raise ValueError("address is required")
        if not template_token:
            raise ValueError("template_token is required")
        body: dict = {
            "address":                address.strip(),
            "template_token":         template_token,
            "photo_request_platform": photo_request_platform,
        }
        if loan_number:      body["loan_number"]      = loan_number
        if borrower_name:    body["borrower_name"]    = borrower_name
        if additional_notes: body["additional_notes"] = additional_notes
        if external_id:      body["external_id"]      = external_id
        if expires_at:       body["expires_at"]       = expires_at
        if contacts:         body["contacts_attributes"] = contacts
        if extra:            body.update(extra)
        return self._request("POST", "/photo-requests", body=body)

    def duplicate_photo_request(self, parent_id) -> dict:
        """Create a new request that mirrors {parent_id} (Draw N from Draw N-1).

        Per the docs: 'Useful for Draw Inspections, it creates a new Photo
        Request exactly the same as the provided ID, creating a Draw 2 from
        Draw 1 for example.' The new request gets a parent_id link to the
        original — multiple draws against one address chain naturally."""
        return self._request("POST", f"/photo-requests/{parent_id}/duplicate")

    def cancel_photo_request(self, request_id) -> dict:
        """Cancel a request. Per docs, succeeds immediately only if status is
        'unassigned'; otherwise ProxyPics support is notified to handle it."""
        return self._request("DELETE", f"/photo-requests/{request_id}")

    # ---- templates -------------------------------------------------------------
    def list_templates(self, page: int = 1, per_page: int = 50) -> Any:
        return self._request("GET", "/photo-request-templates",
                             params={"page": page, "per_page": per_page})

    def find_template_by_name(self, name_substring: str) -> Optional[dict]:
        """Find a template whose name contains the given substring (case-insensitive).
        Returns the full template dict (you'll usually want `.token` from it)."""
        needle = (name_substring or "").strip().lower()
        if not needle:
            return None
        for page in range(1, 6):  # 250 templates is more than anyone has
            resp = self.list_templates(page=page, per_page=50)
            data = resp if isinstance(resp, list) else (resp.get("data") if isinstance(resp, dict) else None)
            if not data:
                break
            for t in data:
                if isinstance(t, dict) and needle in (t.get("name") or "").lower():
                    return t
        return None
