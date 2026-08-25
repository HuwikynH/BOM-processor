"""
attrition_engine.py
-------------------
Core parsing engine that reads component_dictionary.json and attrition_rules.json,
then analyzes BOM rows (Part Type + Description + Unit) to produce:
  - canonical_type   : e.g. "RESISTOR", "IC", "WIRE"
  - full_name        : e.g. "Resistor (Chip 0603)"
  - package_detected : e.g. "0603", "SOP"
  - attrition_rate   : e.g. 0.10
  - attrition_pct    : e.g. "10%"
  - qty_bom          : original quantity
  - qty_with_attrition : rounded quantity including attrition
"""

import json
import math
import re
import os

# ── Load JSON databases ────────────────────────────────────────────────────────
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _load_json(filename):
    with open(os.path.join(_BASE_DIR, filename), encoding="utf-8") as f:
        return json.load(f)


def _active_rules_filename() -> str:
    """Resolve the currently active rules file via _active_rules.json pointer.
    Falls back to the original attrition_rules.json."""
    from file_versions import active_filename
    return active_filename("attrition_rules.json")


from file_versions import active_filename as _active_filename
DICTIONARY = _load_json(_active_filename("component_dictionary.json"))
RULES      = _load_json(_active_rules_filename())

# ── Flatten keyword → type map (ignore _comment* keys) ───────────────────────
_KEYWORD_MAP: dict[str, str] = {}
for _group in DICTIONARY["keyword_to_type"].values():
    if isinstance(_group, dict):
        for k, v in _group.items():
            if not k.startswith("_"):
                _KEYWORD_MAP[k.upper()] = v

# ── Package keyword list (ordered longest-first to avoid partial matches) ─────
_PACKAGE_MAP: dict[str, str] = {
    k.upper(): v
    for k, v in DICTIONARY["package_keywords"].items()
    if not k.startswith("_")
}
_PACKAGE_KEYS_SORTED = sorted(_PACKAGE_MAP.keys(), key=len, reverse=True)

# ── Unit map ──────────────────────────────────────────────────────────────────
_UNIT_MAP: dict[str, str] = {
    k.upper(): v
    for k, v in DICTIONARY["unit_aliases"].items()
    if not k.startswith("_")
}
_LENGTH_UNITS = set(RULES["assembly_context"]["length_units"])


# ── Public API ────────────────────────────────────────────────────────────────

def parse_description(description: str) -> str | None:
    """
    Scan `description` tokens to detect canonical component type.
    Strategy:
      1. Try multi-word phrases first (e.g. "HEAT SHRINK", "OP AMP")
      2. Try each comma/space-separated token
    Returns canonical type string or None if unknown.
    """
    if not description:
        return None

    desc_upper = description.upper()

    # Phase 1 – try multi-word keyword matches (phrases)
    for kw in sorted(_KEYWORD_MAP.keys(), key=len, reverse=True):
        if " " in kw and kw in desc_upper:
            return _KEYWORD_MAP[kw]

    # Phase 2 – tokenize by comma then whitespace, try each token (keeps things like CE-WIRE intact)
    tokens = [t.strip() for part in desc_upper.split(",") for t in part.split()]
    for token in tokens:
        if token in _KEYWORD_MAP:
            return _KEYWORD_MAP[token]

    # Phase 3 – aggressive tokenization by hyphen, underscore, slash (for cases like CAP-CERAM or RES_SMD)
    aggressive_tokens = re.split(r'[,\s\-_/]+', desc_upper)
    for token in aggressive_tokens:
        if token in _KEYWORD_MAP:
            return _KEYWORD_MAP[token]

    return None


def detect_package(description: str) -> str | None:
    """
    Scan `description` for known package/size keywords.
    Returns canonical package string (e.g. '0603', 'SOP', 'QFN') or None.
    """
    if not description:
        return None

    desc_upper = description.upper()

    # Try exact token matches (comma/space split)
    tokens = {t.strip() for part in desc_upper.split(",") for t in part.split()}
    for pkg_key in _PACKAGE_KEYS_SORTED:
        if pkg_key in tokens:
            return _PACKAGE_MAP[pkg_key]

    # Aggressive token match (split by hyphen, underscore, parentheses)
    aggressive_tokens = set(re.split(r'[,\s\-_/()]+', desc_upper))
    for pkg_key in _PACKAGE_KEYS_SORTED:
        if pkg_key in aggressive_tokens:
            return _PACKAGE_MAP[pkg_key]

    # Fallback: substring search (for cases like "1206SMD" or "SOD-323")
    for pkg_key in _PACKAGE_KEYS_SORTED:
        pattern = r"\b" + re.escape(pkg_key) + r"\b"
        if re.search(pattern, desc_upper):
            return _PACKAGE_MAP[pkg_key]

    # Last resort: strip leading digits from tokens (e.g. "64QFN" → "QFN", "144TQFP" → "TQFP")
    for token in tokens:
        stripped = re.sub(r"^\d+", "", token)
        if stripped and stripped in _PACKAGE_MAP:
            return _PACKAGE_MAP[stripped]

    return None


def get_attrition_rate(
    canonical_type: str,
    package: str | None,
    unit: str | None,
    description: str | None = None,
) -> float:
    """
    Look up attrition rate from attrition_rules.json.
    Decision logic:
      - If unit is a length unit → cable_box_rules (e.g. WIRE)
      - CONNECTOR: cable context (housing/crimp in description) → cable_box_rules,
        otherwise smt_rules (PCBA headers)
      - Else → smt_rules, with package sub-lookup for RES/CAP
    Returns rate as decimal (e.g. 0.10 for 10%).
    """
    if not canonical_type:
        return 0.0

    unit_upper = (unit or "").upper()
    is_length = unit_upper in _LENGTH_UNITS

    # WIRE always uses cable_box_rules
    # Component types that ALWAYS use cable_box_rules
    _ALWAYS_CABLE = {"WIRE", "TERMINAL", "HEAT_SHRINK", "CABLE_TIE",
                     "LABEL", "POWER_MONITOR"}

    desc_upper = (description or "").upper()
    if canonical_type == "CONNECTOR" and any(
            k in desc_upper for k in ("HOUSING", "CRIMP")):
        # Connector / Housing in a cable assembly → 0.5%
        table = RULES["cable_box_rules"]
    elif canonical_type in _ALWAYS_CABLE:
        table = RULES["cable_box_rules"]
    elif is_length:
        # Length unit + non-wire → cable context (heat shrink sold by length, etc.)
        table = RULES["cable_box_rules"]
    else:
        table = RULES["smt_rules"]

    rule_entry = table.get(canonical_type)
    if rule_entry is None:
        # Fallback: try the other table
        other_table = RULES["smt_rules"] if table is RULES["cable_box_rules"] else RULES["cable_box_rules"]
        rule_entry = other_table.get(canonical_type)
    if rule_entry is None:
        return 0.0  # Unknown type → 0% (safe default)

    # Sub-lookup by package (for RESISTOR, CAPACITOR, IC, etc.)
    if package and package in rule_entry:
        return rule_entry[package]

    return rule_entry.get("_default", 0.0)


def apply_attrition(qty: float, rate: float, unit: str) -> float:
    """
    Calculate quantity with attrition applied.
      - PCS units → round UP to nearest integer
      - Length units (IN, FT, M...) → keep 3 decimal places
    """
    result = qty * (1.0 + rate)
    unit_upper = (unit or "").upper()
    if unit_upper in _LENGTH_UNITS:
        return round(result, 3)
    else:
        return math.ceil(result)


def analyze_row(
    part_type: str | None,
    description: str | None,
    qty: float | None,
    unit: str | None,
) -> dict:
    """
    Full analysis of a single BOM row.

    Parameters
    ----------
    part_type   : from 'Part Type' column (can be None for CSV-only BOMs)
    description : from 'Description' column
    qty         : numeric quantity
    unit        : unit string (EA, IN, FT, ...)

    Returns
    -------
    dict with keys:
      canonical_type, full_name, package_detected,
      attrition_rate, attrition_pct,
      qty_bom, unit, qty_with_attrition, resolved_via
    """
    qty = qty or 0.0
    unit = (unit or "EA").strip()

    # ── Step 1: Resolve canonical type ────────────────────────────────────────
    canonical_type = None
    resolved_via = "unknown"

    # Priority 1: use Part Type column if available and recognized
    if part_type:
        pt_upper = part_type.strip().upper()
        if pt_upper in _KEYWORD_MAP:
            canonical_type = _KEYWORD_MAP[pt_upper]
            resolved_via = "part_type_col"
        else:
            # Normalize Part Type directly (many are already canonical)
            _DIRECT_MAP = {
                "RESISTOR": "RESISTOR", "CAPACITOR": "CAPACITOR",
                "IC": "IC", "DIODE": "DIODE", "TRANSISTOR": "TRANSISTOR",
                "INDUCTOR": "INDUCTOR", "LED": "LED", "RELAY": "RELAY",
                "CRYSTAL": "CRYSTAL", "XTAL": "CRYSTAL", "OSCILLATOR": "OSCILLATOR",
                "CONNECTOR": "CONNECTOR", "TERMINAL": "TERMINAL",
                "WIRE": "WIRE", "CABLE": "WIRE",
                "SHRINK TUBE": "HEAT_SHRINK", "HEAT SHRINK": "HEAT_SHRINK",
                "LABEL": "LABEL", "SCREW": "SCREW_NUT_WASHER",
                "NUT": "SCREW_NUT_WASHER", "WASHER": "SCREW_NUT_WASHER",
                "STANDOFF": "SCREW_NUT_WASHER", "TIE CABLE": "CABLE_TIE",
                "SWITCH": "SWITCH", "JUMPER": "JUMPER",
                "MECH": "SHEET_METAL", "MECH ASSY": "SHEET_METAL",
                "PLASTIC": "OTHER_SPECIAL", "CABLE ASSY": "OTHER_SPECIAL",
                "PCBA": "OTHER_SPECIAL", "PCBA ASSY": "OTHER_SPECIAL",
                "FAB DWG": "OTHER_SPECIAL", "PCB SCHEMATIC": "OTHER_SPECIAL",
                "SUB ASSEMBLY - ELECTRICAL": "OTHER_SPECIAL",
                "MANUFACTURING DOCUMENTATION PACKAGE.": "OTHER_SPECIAL",
                "ASSY PROGRAMMABLE": "OTHER_SPECIAL",
                "ASSY PROGRAMMABLE.": "OTHER_SPECIAL",
                "TEST POINT": "TEST_POINT",
                "TEST SPECS PCB": "OTHER_SPECIAL",
                "ADHESIVE": "OTHER_SPECIAL",
            }
            canonical_type = _DIRECT_MAP.get(pt_upper)
            if canonical_type:
                resolved_via = "part_type_col"

    # Priority 2: fallback to Description parsing
    if not canonical_type:
        canonical_type = parse_description(description)
        if canonical_type:
            resolved_via = "description_parse"

    # ── Step 2: Detect package from Description ────────────────────────────────
    package = detect_package(description)

    # ── Step 3: Get attrition rate ────────────────────────────────────────────
    rate = get_attrition_rate(canonical_type, package, unit, description)

    # ── Step 4: Calculate Qty with attrition ──────────────────────────────────
    qty_final = apply_attrition(qty, rate, unit)

    # ── Step 5: Build human-readable full name ────────────────────────────────
    FULL_NAMES = {
        "RESISTOR": "Resistor",
        "CAPACITOR": "Capacitor",
        "INDUCTOR": "Inductor",
        "DIODE": "Diode",
        "TRANSISTOR": "Transistor",
        "IC": "IC / Integrated Circuit",
        "CRYSTAL": "Crystal",
        "OSCILLATOR": "Oscillator",
        "LED": "LED",
        "RELAY": "Relay",
        "JUMPER": "Jumper",
        "SWITCH": "Switch",
        "TRANSFORMER": "Transformer",
        "TEST_POINT": "Test Point",
        "CONNECTOR": "Connector",
        "TERMINAL": "Terminal / Crimp / Ferrule",
        "WIRE": "Wire / Cable",
        "HEAT_SHRINK": "Heat Shrink Tubing",
        "CABLE_TIE": "Cable Tie",
        "LABEL": "Label / Marker",
        "SCREW_NUT_WASHER": "Screw / Nut / Washer / Standoff",
        "SHEET_METAL": "Sheet Metal / Mechanical",
        "POWER_MONITOR": "Power / Monitor / CB",
        "OTHER_SPECIAL": "Other / Special",
    }
    base_name = FULL_NAMES.get(canonical_type, canonical_type or "Unknown")
    if package and canonical_type in ("RESISTOR", "CAPACITOR", "INDUCTOR", "LED"):
        full_name = f"{base_name} [{package}]"
    elif package and canonical_type in ("IC", "DIODE", "TRANSISTOR"):
        full_name = f"{base_name} [{package}]"
    else:
        full_name = base_name

    return {
        "canonical_type":     canonical_type or "UNKNOWN",
        "full_name":          full_name,
        "package_detected":   package or "-",
        "attrition_rate":     rate,
        "attrition_pct":      f"{rate*100:.1f}%",
        "qty_bom":            qty,
        "unit":               unit,
        "qty_with_attrition": qty_final,
        "resolved_via":       resolved_via,
    }


# ── Quick self-test ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    TEST_CASES = [
        # (part_type, description, qty, unit)
        ("RESISTOR", "RES D,0402,100.00 OHM,1.00%,1/16W,YES", 100, "EA"),
        ("RESISTOR", "RESISTOR, 470 OHM, 1/4W, 5%, THK FILM, 1206 SMD", 50, "EA"),
        ("CAPACITOR", "CAP, 1UF, 6.3V, 10%, X5R, CER, 0402 SMD", 200, "EA"),
        ("IC",        "IC, CYCLONE 5, 5CEFA7F31C6N, 896-PIN FBGA", 5, "EA"),
        ("IC",        "IC, SN74HCT14, HEX SCHMITT-TRIGGER INVER, SOIC", 20, "EA"),
        ("IC",        "IC,ADC,DUAL 14-BIT,PARALLEL/SERIAL,64QFN", 3, "EA"),
        ("WIRE",      "CABLE, 3 COND, 18AWG, SHLD, UL STYLE 2501", 32, "IN"),
        ("TERMINAL",  "TERMINAL, RING LUG, 22-18 AWG, #8 STUD", 10, "EA"),
        ("SHRINK TUBE", "TUBING, HEAT SHRINK, 1/8 IN, BLACK", 3.5, "IN"),
        ("LABEL",     "LABEL, WHITE/CLEAR, 1 IN X 3 IN", 20, "EA"),
        ("SCREW",     "SCREW, BHCS, 4-40 X 1/4IN, 18-8 SS", 6, "EA"),
        ("LED",       "LED GREEN 569nm 6MCD 2.1V 10MA 0603 SMD", 10, "EA"),
        ("DIODE",     "DIODE, FAST SWITCHING, 1N4148, SOD-323 S", 30, "EA"),
        # CSV-style description without Part Type
        (None,        "RES SMD 2.4K OHM 5% 1/4W 1206", 100, "EA"),
        (None,        "CAP, CERAMIC 33PF 50V NP0 0603", 50, "EA"),
        (None,        "TBG SHRINK 3/8 ID X .025 WALL POLYOLEFIN BLK", 0.167, "FT"),
    ]

    print(f"{'Part Type':<20} {'Description':<45} {'Qty':>6} {'Unit':<5} | "
          f"{'Full Name':<35} {'Pkg':<10} {'Att%':<6} {'Qty+Att':>9} {'Via'}")
    print("-" * 170)

    for pt, desc, qty, unit in TEST_CASES:
        r = analyze_row(pt, desc, qty, unit)
        print(
            f"{(pt or 'N/A'):<20} "
            f"{desc[:44]:<45} "
            f"{qty:>6} {unit:<5} | "
            f"{r['full_name']:<35} "
            f"{r['package_detected']:<10} "
            f"{r['attrition_pct']:<6} "
            f"{r['qty_with_attrition']:>9} "
            f"{r['resolved_via']}"
        )
