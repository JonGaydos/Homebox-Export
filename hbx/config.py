"""Portable settings JSON next to the exe; API key in Windows Credential Manager."""
import json
import sys
from pathlib import Path

import keyring

KEYRING_SERVICE = "HomeboxExport"
KEYRING_USER = "api_key"

DEFAULTS = {
    "homebox_url": "",
    "owner": "",
}


def config_path() -> Path:
    """Config lives next to the .exe (or the repo root) so the app stays portable."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / "homebox_export_config.json"
    return Path(__file__).resolve().parent.parent / "homebox_export_config.json"


def load_config(path=None) -> dict:
    cfg = dict(DEFAULTS)
    try:
        data = json.loads(Path(path or config_path()).read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return cfg
    if isinstance(data, dict):
        # v1 config used "url"; carry it over
        if "url" in data and "homebox_url" not in data:
            data["homebox_url"] = data.pop("url")
        cfg.update({k: v for k, v in data.items() if k in DEFAULTS})
    return cfg


def save_config(cfg: dict, path=None):
    Path(path or config_path()).write_text(
        json.dumps(cfg, indent=2), encoding="utf-8")


def load_api_key() -> str:
    try:
        return keyring.get_password(KEYRING_SERVICE, KEYRING_USER) or ""
    except Exception:
        return ""


def save_api_key(key: str):
    keyring.set_password(KEYRING_SERVICE, KEYRING_USER, key)
