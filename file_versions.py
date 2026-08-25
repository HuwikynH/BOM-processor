"""
file_versions.py
----------------
Shared helpers for versioned JSON config files.

The original file (attrition_rules.json / component_dictionary.json) is NEVER
overwritten. Saving always creates a new version file:
    attrition_rules_v<timestamp>.json
    component_dictionary_v<timestamp>.json
An active-pointer file (_active_rules.json / _active_dictionary.json) records
which version the engine currently uses.
"""

import datetime
import glob
import json
import os

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))

_PTR_NAMES = {
    "attrition_rules.json": "_active_rules.json",
    "component_dictionary.json": "_active_dictionary.json",
}


def _ptr_path(base_name: str) -> str:
    return os.path.join(_BASE_DIR, _PTR_NAMES[base_name])


def _is_base(name: str) -> bool:
    return not name.startswith("_") and "_v" not in name


def active_filename(base_name: str) -> str:
    """Resolve the currently active file for a base config. Falls back to base."""
    try:
        with open(_ptr_path(base_name), encoding="utf-8") as f:
            name = json.load(f).get("active_file")
        if name and os.path.exists(os.path.join(_BASE_DIR, name)):
            return name
    except Exception:
        pass
    return base_name


def set_active_filename(base_name: str, filename: str):
    with open(_ptr_path(base_name), "w", encoding="utf-8") as f:
        json.dump({"active_file": filename}, f, indent=2)


def list_files(base_name: str) -> list[str]:
    """[original, v_oldest, ..., v_newest]"""
    pattern = os.path.join(_BASE_DIR, os.path.splitext(base_name)[0] + "_v*.json")
    versions = sorted(os.path.basename(p) for p in glob.glob(pattern))
    return [base_name] + versions


def display_name(filename: str) -> str:
    """'attrition_rules.json' → 'Original (default)';
    'attrition_rules_v2026-08-25_141530.json' → 'v2026-08-25 14:15:30'"""
    if filename in _PTR_NAMES:
        return "Original (default)"
    stem = os.path.splitext(filename)[0]
    v = stem.split("_v", 1)[1] if "_v" in stem else stem
    try:
        date, time = v.split("_", 1)
        return f"v{date} {time[:2]}:{time[2:4]}:{time[4:6]}"
    except ValueError:
        return f"v{v}"


def load_active(base_name: str, filename: str | None = None) -> dict:
    name = filename or active_filename(base_name)
    with open(os.path.join(_BASE_DIR, name), encoding="utf-8") as f:
        return json.load(f)


def save_new_version(base_name: str, data: dict, saved_from: str = "") -> str:
    """Write data to a NEW timestamped version file, make it active.
    Returns the new filename."""
    stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    fname = f"{os.path.splitext(base_name)[0]}_v{stamp}.json"
    payload = dict(data)
    meta = dict(payload.get("_meta", {}))
    if saved_from:
        meta["saved_from"] = saved_from
    meta["saved_at"] = stamp.replace("_", " ")
    meta["base_version"] = active_filename(base_name)
    payload["_meta"] = meta
    with open(os.path.join(_BASE_DIR, fname), "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)
    set_active_filename(base_name, fname)
    return fname
