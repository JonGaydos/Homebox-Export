import pytest
import requests

from hbx.client import (
    AuthError, HomeboxClient, HomeboxError, format_asset_id, in_asset_range,
)


class FakeResponse:
    def __init__(self, status_code=200, payload=None, headers=None, content=b""):
        self.status_code = status_code
        self._payload = payload if payload is not None else {}
        self.headers = headers or {}
        self.content = content

    def json(self):
        return self._payload


class FakeSession:
    """Returns queued responses; a requests exception instance raises instead."""

    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []

    def get(self, url, headers=None, params=None, timeout=None):
        self.calls.append({"url": url, "headers": headers, "params": params,
                           "timeout": timeout})
        r = self.responses.pop(0)
        if isinstance(r, Exception):
            raise r
        return r


def make_client(responses, max_retries=4):
    session = FakeSession(responses)
    client = HomeboxClient(
        "http://homebox.example:3100/", "testkey",
        max_retries=max_retries, session=session, sleep=lambda s: None,
    )
    return client, session


ITEM_TYPE = {"isLocation": False}
LOC_TYPE = {"isLocation": True}


def test_get_asset_found_sends_bearer_key():
    payload = {"total": 1, "items": [{"name": "Laptop", "assetId": "000-001"}]}
    client, session = make_client([FakeResponse(200, payload)])
    item = client.get_asset("000-001")
    assert item["name"] == "Laptop"
    assert session.calls[0]["url"] == "http://homebox.example:3100/api/v1/assets/000-001"
    assert session.calls[0]["headers"]["Authorization"] == "Bearer testkey"
    assert session.calls[0]["timeout"] is not None


def test_get_asset_not_found_returns_none():
    client, _ = make_client([FakeResponse(200, {"total": 0, "items": []})])
    assert client.get_asset("999-999") is None


def test_401_raises_auth_error_without_retry():
    client, session = make_client([FakeResponse(401)])
    with pytest.raises(AuthError):
        client.get_asset("000-001")
    assert len(session.calls) == 1


def test_500_retries_then_succeeds():
    payload = {"total": 1, "items": [{"name": "X"}]}
    client, session = make_client([FakeResponse(500), FakeResponse(200, payload)])
    assert client.get_asset("000-001")["name"] == "X"
    assert len(session.calls) == 2


def test_connection_error_retries_then_raises_homebox_error():
    errs = [requests.exceptions.ConnectionError("refused")] * 3
    client, session = make_client(errs, max_retries=2)
    with pytest.raises(HomeboxError):
        client.get_asset("000-001")
    assert len(session.calls) == 3


def test_search_items_hits_entities_endpoint():
    payload = {"items": [{"name": "Drill", "id": "abc"}]}
    client, session = make_client([FakeResponse(200, payload)])
    items = client.search_items("dri")
    assert items[0]["name"] == "Drill"
    assert session.calls[0]["url"] == "http://homebox.example:3100/api/v1/entities"
    assert session.calls[0]["params"] == {"q": "dri", "pageSize": 100}


def test_search_items_passes_filters():
    client, session = make_client([FakeResponse(200, {"items": []})])
    client.search_items("x", tags=["t1", "t2"], parent_ids=["p1"], page=3)
    assert session.calls[0]["params"] == {
        "q": "x", "pageSize": 100, "tags": ["t1", "t2"],
        "parentIds": ["p1"], "page": 3,
    }


def test_search_items_all_paginates_and_drops_locations():
    page1 = {"items": [{"name": f"i{i}", "entityType": ITEM_TYPE} for i in range(500)]}
    page2 = {"items": [{"name": "garage", "entityType": LOC_TYPE},
                       {"name": "last", "entityType": ITEM_TYPE}]}
    client, session = make_client([FakeResponse(200, page1), FakeResponse(200, page2)])
    items, truncated = client.search_items_all("x")
    assert len(items) == 501
    assert truncated is False
    assert all(not it["entityType"]["isLocation"] for it in items)
    assert session.calls[0]["params"]["page"] == 1
    assert session.calls[1]["params"]["page"] == 2


def test_search_items_all_truncates_at_cap():
    responses = [
        FakeResponse(200, {"items": [{"name": "x", "entityType": ITEM_TYPE}] * 500})
        for _ in range(10)
    ]
    client, _ = make_client(responses)
    items, truncated = client.search_items_all("x", max_items=1000)
    assert len(items) == 1000
    assert truncated is True


def test_get_entity():
    client, session = make_client([FakeResponse(200, {"id": "e1", "name": "Thing"})])
    assert client.get_entity("e1")["name"] == "Thing"
    assert session.calls[0]["url"] == "http://homebox.example:3100/api/v1/entities/e1"


def test_get_maintenance_accepts_list_or_dict():
    client, _ = make_client([FakeResponse(200, [{"name": "Oil change"}])])
    assert client.get_maintenance("e1") == [{"name": "Oil change"}]
    client, _ = make_client([FakeResponse(200, {"entries": [{"name": "Filter"}]})])
    assert client.get_maintenance("e1") == [{"name": "Filter"}]


def test_get_attachment_returns_bytes():
    client, session = make_client([FakeResponse(200, content=b"\xff\xd8jpeg")])
    data = client.get_attachment("e1", "a1")
    assert data == b"\xff\xd8jpeg"
    assert session.calls[0]["url"] == (
        "http://homebox.example:3100/api/v1/entities/e1/attachments/a1")


def test_list_tags_and_location_tree():
    client, session = make_client([
        FakeResponse(200, [{"id": "t1"}]),
        FakeResponse(200, [{"id": "l1", "type": "location"}]),
    ])
    assert client.list_tags() == [{"id": "t1"}]
    assert client.location_tree() == [{"id": "l1", "type": "location"}]
    assert session.calls[0]["url"] == "http://homebox.example:3100/api/v1/tags"
    assert session.calls[1]["url"] == "http://homebox.example:3100/api/v1/entities/tree"


def test_test_connection_ok_and_bad_key():
    client, _ = make_client([FakeResponse(200, {"items": []})])
    ok, _ = client.test_connection()
    assert ok is True
    client, _ = make_client([FakeResponse(401)])
    ok, msg = client.test_connection()
    assert ok is False
    assert "key" in msg.lower()


def test_format_asset_id():
    assert format_asset_id("002-062") == "002-062"
    assert format_asset_id("2062") == "002-062"
    assert format_asset_id(62) == "000-062"
    assert format_asset_id("1234567") == "1234-567"
    assert format_asset_id("") == ""
    assert format_asset_id("garbage") == "garbage"


def test_in_asset_range():
    assert in_asset_range("001-050", "001-001", "001-100") is True
    assert in_asset_range("001-101", "001-001", "001-100") is False
    assert in_asset_range("000-999", "001-001", "001-100") is False
    assert in_asset_range("001-050", "", "") is True
    assert in_asset_range("garbage", "", "") is False
