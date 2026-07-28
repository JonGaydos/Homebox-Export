"""Homebox API client with retrying GETs, adapted from homebox-label.

Authenticates with a Homebox API key sent as a Bearer token. Uses the
current entities-based API (Homebox removed /api/v1/items).
"""
import re
import time

import requests

from hbx.retry import retry_policy

TIMEOUT = 15
ASSET_ID_RE = re.compile(r"^\d{3}-\d{3}$")


def format_asset_id(raw) -> str:
    """Normalize input to Homebox's ###-### asset id form.

    Already-formatted ids pass through; digit strings are zero-padded.
    Non-numeric input is returned unchanged so lookups fail visibly.
    """
    s = str(raw if raw is not None else "").strip()
    if ASSET_ID_RE.match(s):
        return s
    digits = re.sub(r"\D", "", s)
    if not digits:
        return s
    digits = digits.zfill(6)
    return f"{digits[:-3]}-{digits[-3:]}"


def _asset_num(asset_id):
    return int(asset_id.replace("-", ""))


def in_asset_range(asset_id, lo, hi):
    """True if asset_id falls in [lo, hi]. Blank or invalid bounds are open."""
    if not ASSET_ID_RE.match(asset_id or ""):
        return False
    n = _asset_num(asset_id)
    if lo and ASSET_ID_RE.match(lo) and n < _asset_num(lo):
        return False
    if hi and ASSET_ID_RE.match(hi) and n > _asset_num(hi):
        return False
    return True


def is_item(entity) -> bool:
    return not (entity.get("entityType") or {}).get("isLocation", False)


def location_ids(node):
    """A location node's id plus all location-type descendant ids."""
    ids = [node["id"]]
    for child in node.get("children", []):
        if child.get("type") == "location":
            ids.extend(location_ids(child))
    return ids


class HomeboxError(Exception):
    """Terminal request failure (network down, HTTP error after retries)."""


class AuthError(HomeboxError):
    """401 from Homebox; the API key is wrong or expired."""


class HomeboxClient:
    def __init__(self, base_url, api_key, max_retries=4, session=None, sleep=time.sleep):
        self.base = base_url.rstrip("/")
        self.api_key = api_key
        self.max_retries = max_retries
        self.session = session or requests.Session()
        self.sleep = sleep

    def _headers(self):
        return {"Authorization": f"Bearer {self.api_key}"}

    def _get(self, path, params=None, raw=False):
        attempt = 0
        while True:
            exc = None
            resp = None
            try:
                resp = self.session.get(
                    self.base + path, headers=self._headers(),
                    params=params, timeout=TIMEOUT,
                )
            except requests.exceptions.RequestException as e:
                exc = e

            status = resp.status_code if resp is not None else None
            if status == 401:
                raise AuthError("Homebox rejected the API key (HTTP 401).")

            retry_after = resp.headers.get("Retry-After") if resp is not None else None
            should_retry, delay = retry_policy(
                exception=exc, status_code=status, retry_after=retry_after,
                attempt=attempt, max_retries=self.max_retries,
            )
            if should_retry:
                self.sleep(delay)
                attempt += 1
                continue

            if exc is not None:
                raise HomeboxError(f"Homebox unreachable: {exc}")
            if status != 200:
                raise HomeboxError(f"Homebox returned HTTP {status}.")
            if raw:
                return resp.content
            try:
                return resp.json()
            except ValueError as e:
                raise HomeboxError(f"Homebox returned invalid JSON: {e}")

    def get_asset(self, asset_id):
        data = self._get(f"/api/v1/assets/{asset_id}")
        if not data.get("total") or not data.get("items"):
            return None
        return data["items"][0]

    def search_items(self, query, page_size=100, tags=None, parent_ids=None,
                     page=None):
        params = {"q": query, "pageSize": page_size}
        if tags:
            params["tags"] = tags
        if parent_ids:
            params["parentIds"] = parent_ids
        if page is not None:
            params["page"] = page
        data = self._get("/api/v1/entities", params=params)
        return data.get("items", [])

    def search_items_all(self, query, tags=None, parent_ids=None,
                         max_items=10000, page_size=500):
        """Page through results, keeping only items (not locations).

        Returns (items, truncated).
        """
        items = []
        page = 1
        while len(items) < max_items:
            batch = self.search_items(query, page_size=page_size, tags=tags,
                                      parent_ids=parent_ids, page=page)
            items.extend(e for e in batch if is_item(e))
            if len(batch) < page_size:
                return items, False
            page += 1
        return items[:max_items], True

    def get_entity(self, entity_id):
        return self._get(f"/api/v1/entities/{entity_id}")

    def get_maintenance(self, entity_id):
        data = self._get(f"/api/v1/entities/{entity_id}/maintenance")
        if isinstance(data, list):
            return data
        return data.get("entries", data.get("data", []))

    def get_attachment(self, entity_id, attachment_id) -> bytes:
        return self._get(
            f"/api/v1/entities/{entity_id}/attachments/{attachment_id}", raw=True)

    def list_tags(self):
        return self._get("/api/v1/tags")

    def location_tree(self):
        return self._get("/api/v1/entities/tree")

    def test_connection(self):
        try:
            self.search_items("", page_size=1)
            return True, "Connected to Homebox."
        except AuthError:
            return False, "Connected, but the API key was rejected."
        except HomeboxError as e:
            return False, str(e)
