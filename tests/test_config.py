import json

from hbx import config


def test_load_config_returns_defaults_when_missing(tmp_path):
    cfg = config.load_config(tmp_path / "nope.json")
    assert cfg == config.DEFAULTS


def test_save_and_load_roundtrip(tmp_path):
    path = tmp_path / "cfg.json"
    cfg = dict(config.DEFAULTS)
    cfg["homebox_url"] = "http://hb:3100"
    cfg["owner"] = "Jon"
    config.save_config(cfg, path)
    assert config.load_config(path)["homebox_url"] == "http://hb:3100"
    assert config.load_config(path)["owner"] == "Jon"


def test_load_config_migrates_legacy_url_key(tmp_path):
    path = tmp_path / "cfg.json"
    path.write_text(json.dumps({"url": "http://old:3100", "owner": "Jon"}))
    cfg = config.load_config(path)
    assert cfg["homebox_url"] == "http://old:3100"


def test_load_config_ignores_corrupt_file(tmp_path):
    path = tmp_path / "cfg.json"
    path.write_text("{not json")
    assert config.load_config(path) == config.DEFAULTS


class FakeKeyring:
    def __init__(self):
        self.store = {}

    def get_password(self, service, user):
        return self.store.get((service, user))

    def set_password(self, service, user, value):
        self.store[(service, user)] = value


def test_api_key_saved_to_keyring(monkeypatch):
    fake = FakeKeyring()
    monkeypatch.setattr(config, "keyring", fake)
    config.save_api_key("sekret")
    assert config.load_api_key() == "sekret"
    assert fake.store[(config.KEYRING_SERVICE, config.KEYRING_USER)] == "sekret"


def test_load_api_key_returns_empty_on_keyring_failure(monkeypatch):
    class Broken:
        def get_password(self, *a):
            raise RuntimeError("no backend")

    monkeypatch.setattr(config, "keyring", Broken())
    assert config.load_api_key() == ""
