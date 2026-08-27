"""
app.py
------
BOM Processor – Desktop GUI built with CustomTkinter.

Flow:
  1. User clicks Browse → selects xlsx file
  2. App reads row 1 to detect column headers
  3. App processes all data rows → shows table:
       [Part Type | Description (original) | Attrition %]
  4. Status bar shows summary (total rows, unmatched count)
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, ttk
import customtkinter as ctk
import openpyxl

from bom_reader import read_bom_file, get_available_sheets
from attrition_engine import analyze_row, apply_attrition
from attrition_editor import AttritionEditorWindow

# ── Theme ──────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# ── Colour palette ─────────────────────────────────────────────────────────────
CLR_BG        = "#ffffff"
CLR_SURFACE   = "#f4f4f5"
CLR_ACCENT    = "#16a34a"
CLR_ACCENT2   = "#dc2626"
CLR_TEXT      = "#18181b"
CLR_MUTED     = "#52525b"
CLR_WARN      = "#dc2626"
CLR_ROW_EVEN  = "#f4f4f5"
CLR_ROW_ODD   = "#ffffff"
CLR_HEADER_BG = "#e4e4e7"

# ── Attrition band colours ─────────────────────────────────────────────────────
# Continuous gradient anchors: (pct, RGB). Any custom rate from edited rules
# gets a smooth colour between anchors instead of falling into fixed bands.
_ATT_ANCHORS = [
    (0.0,  (0x71, 0x71, 0x7a)),   # grey    – 0%
    (0.5,  (0x0d, 0x94, 0x88)),   # teal    – 0.5%
    (1.0,  (0x05, 0x96, 0x69)),   # green   – 1%
    (2.0,  (0x25, 0x63, 0xeb)),   # blue    – 2%
    (3.0,  (0x7c, 0x3a, 0xed)),   # violet  – 3%
    (5.0,  (0xd9, 0x77, 0x06)),   # orange  – 5%
    (10.0, (0xdc, 0x26, 0x26)),   # red     – 10%+
]


def _att_gradient_hex(pct: float) -> str:
    if pct <= 0:
        return "#71717a"
    if pct >= 10:
        return "#dc2626"
    for (p1, c1), (p2, c2) in zip(_ATT_ANCHORS, _ATT_ANCHORS[1:]):
        if p1 <= pct <= p2:
            t = (pct - p1) / (p2 - p1)
            rgb = tuple(round(a + (b - a) * t) for a, b in zip(c1, c2))
            return "#%02x%02x%02x" % rgb
    return "#71717a"



FRIENDLY_TYPES = {
    "Resistor": "RESISTOR",
    "Capacitor": "CAPACITOR",
    "IC": "IC",
    "Diode": "DIODE",
    "Transistor": "TRANSISTOR",
    "Inductor": "INDUCTOR",
    "LED": "LED",
    "Relay": "RELAY",
    "Crystal": "CRYSTAL",
    "Oscillator": "OSCILLATOR",
    "Connector": "CONNECTOR",
    "Terminal / Crimp / Ferrule": "TERMINAL",
    "Wire / Cable": "WIRE",
    "Heat Shrink Tubing": "HEAT_SHRINK",
    "Label / Marker": "LABEL",
    "Screw / Nut / Washer / Standoff": "SCREW_NUT_WASHER",
    "Cable Tie": "CABLE_TIE",
    "Switch": "SWITCH",
    "Jumper": "JUMPER",
    "Sheet Metal / Mechanical": "SHEET_METAL",
    "Test Point": "TEST_POINT",
    "Other / Special": "OTHER_SPECIAL",
    "Unknown (???)": "UNKNOWN"
}

class BOMApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("BOM Processor  v1.5")
        self.geometry("1100x680")
        self.minsize(900, 540)
        self.configure(fg_color=CLR_BG)
        self.after(0, lambda: self.state("zoomed") if hasattr(self, "state") else None)

        self._filepath: str | None = None
        self._all_rows: list[dict] = []   # processed result rows
        self._filter_text: str = ""
        self._sort_col = None
        self._sort_reverse = False

        self._build_ui()

    # ── UI Layout ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── Top bar ──
        top = ctk.CTkFrame(self, fg_color=CLR_SURFACE, corner_radius=0, height=72)
        top.pack(fill="x", side="top")
        top.pack_propagate(False)

        ctk.CTkLabel(
            top, text="BOM Processor",
            font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
            text_color=CLR_ACCENT,
        ).pack(side="left", padx=18, pady=10)

        # Browse button
        self.btn_browse = ctk.CTkButton(
            top, text="📂  Open File",
            width=130, height=40,
            corner_radius=8,
            fg_color=CLR_ACCENT2, hover_color="#b91c1c", text_color="#ffffff",
            command=self._on_browse,
        )
        self.btn_browse.pack(side="left", padx=8, pady=10)

        # File label
        self.lbl_file = ctk.CTkLabel(
            top, text="No file selected",
            font=ctk.CTkFont(size=14),
            text_color=CLR_MUTED,
        )
        self.lbl_file.pack(side="left", padx=4)

        # Process button (right side)
        self.btn_process = ctk.CTkButton(
            top, text="▶  Process",
            width=120, height=40,
            corner_radius=8,
            fg_color=CLR_ACCENT, hover_color="#15803d", text_color="#ffffff",
            state="disabled",
            command=self._on_process,
        )
        self.btn_process.pack(side="right", padx=(4, 18), pady=10)

        
        # Export Excel button
        self.btn_export = ctk.CTkButton(
            top, text="📥  Export",
            width=110, height=40,
            corner_radius=8,
            fg_color=CLR_ACCENT, hover_color="#15803d", text_color="#ffffff",
            command=self._on_export,
            state="disabled"
        )
        self.btn_export.pack(side="right", padx=4, pady=10)

        # Edit Attrition Rules button
        self.btn_edit_rules = ctk.CTkButton(
            top, text="⚙  Edit Rules",
            width=120, height=40,
            corner_radius=8,
            fg_color="#e4e4e7", hover_color="#d4d4d8", text_color="#18181b",
            command=self._on_open_editor,
        )
        self.btn_edit_rules.pack(side="right", padx=4, pady=10)

        # Edit Keyword Dictionary button
        self.btn_edit_dict = ctk.CTkButton(
            top, text="📚  Edit Dictionary",
            width=140, height=40,
            corner_radius=8,
            fg_color="#e4e4e7", hover_color="#d4d4d8", text_color="#18181b",
            command=self._on_open_dict_editor,
        )
        self.btn_edit_dict.pack(side="right", padx=4, pady=10)

        # ── Filter / search bar ──
        bar = ctk.CTkFrame(self, fg_color=CLR_SURFACE, corner_radius=0, height=54)
        bar.pack(fill="x", side="top")
        bar.pack_propagate(False)

        ctk.CTkLabel(
            bar, text="Filter:", font=ctk.CTkFont(size=14), text_color=CLR_MUTED
        ).pack(side="left", padx=(16, 4), pady=8)

        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", self._on_filter_change)
        filter_entry = ctk.CTkEntry(
            bar, textvariable=self.filter_var,
            width=280, height=28,
            placeholder_text="Search Part Type or Description…",
            fg_color="#ffffff", border_color="#d4d4d8", text_color="#18181b",
        )
        filter_entry.pack(side="left", padx=4, pady=6)

        # Attrition filter dropdown
        ctk.CTkLabel(
            bar, text="Attrition:", font=ctk.CTkFont(size=14), text_color=CLR_MUTED
        ).pack(side="left", padx=(20, 4))
        self.att_filter_var = tk.StringVar(value="All")
        self.att_combo = ctk.CTkComboBox(
            bar,
            values=["All", "Unknown"],
            variable=self.att_filter_var,
            width=110, height=28,
            fg_color="#ffffff", border_color="#d4d4d8", text_color="#18181b",
            command=self._on_filter_change,
        )
        self.att_combo.pack(side="left", padx=4)

        # Internal P/N column selector
        ctk.CTkLabel(
            bar, text="Internal P/N Col:", font=ctk.CTkFont(size=13), text_color=CLR_MUTED
        ).pack(side="left", padx=(16, 4))
        self.int_pn_col_var = tk.StringVar(value="Auto")
        self.int_pn_col_combo = ctk.CTkComboBox(
            bar,
            values=["Auto"],
            variable=self.int_pn_col_var,
            width=160, height=28,
            fg_color="#ffffff", border_color="#d4d4d8", text_color="#18181b",
            command=self._on_int_pn_col_change,
        )
        self.int_pn_col_combo.pack(side="left", padx=4)

        # MFR P/N column selector
        ctk.CTkLabel(
            bar, text="MFR P/N Col:", font=ctk.CTkFont(size=13), text_color=CLR_MUTED
        ).pack(side="left", padx=(16, 4))
        self.mfr_pn_col_var = tk.StringVar(value="Auto")
        self.mfr_pn_col_combo = ctk.CTkComboBox(
            bar,
            values=["Auto"],
            variable=self.mfr_pn_col_var,
            width=160, height=28,
            fg_color="#ffffff", border_color="#d4d4d8", text_color="#18181b",
            command=self._on_mfr_pn_col_change,
        )
        self.mfr_pn_col_combo.pack(side="left", padx=4)

        # Row counter label (right)
        self.lbl_count = ctk.CTkLabel(
            bar, text="",
            font=ctk.CTkFont(size=13), text_color=CLR_MUTED,
        )
        self.lbl_count.pack(side="right", padx=18)

        # ── Main table ──
        table_frame = ctk.CTkFrame(self, fg_color=CLR_BG, corner_radius=0)
        table_frame.pack(fill="both", expand=True, padx=0, pady=0)

        # ttk style for the Treeview
        style = ttk.Style()
        style.theme_use("default")
        style.configure("BOM.Treeview",
                        background=CLR_ROW_EVEN,
                        foreground=CLR_TEXT,
                        rowheight=26,
                        fieldbackground=CLR_ROW_EVEN,
                        bordercolor=CLR_BG,
                        borderwidth=0,
                        font=("Segoe UI", 13))
        style.configure("BOM.Treeview.Heading",
                        background=CLR_HEADER_BG,
                        foreground=CLR_TEXT,
                        relief="flat",
                        font=("Segoe UI", 13, "bold"))
        style.map("BOM.Treeview",
                  background=[("selected", "#bbf7d0")],
                  foreground=[("selected", "#000000")])
        style.map("BOM.Treeview.Heading",
                  background=[("active", "#d4d4d8")])

        cols = ("no", "part_type", "description", "mfr_pn", "int_pn", "qty", "attrition")
        self.tree = ttk.Treeview(
            table_frame,
            columns=cols,
            show="headings",
            style="BOM.Treeview",
            selectmode="extended",
        )

        # Column definitions
        self.tree.heading("no",          text="#",              anchor="center", command=lambda: self._sort_by("no"))
        self.tree.heading("part_type",   text="Part Type",      anchor="w", command=lambda: self._sort_by("part_type"))
        self.tree.heading("description", text="Description (original)", anchor="w", command=lambda: self._sort_by("description"))
        self.tree.heading("mfr_pn",      text="MFR P/N",        anchor="w",     command=lambda: self._sort_by("mfr_pn"))
        self.tree.heading("int_pn",      text="Internal P/N",   anchor="w",     command=lambda: self._sort_by("int_pn"))
        self.tree.heading("qty",          text="Qty",            anchor="e",     command=lambda: self._sort_by("qty"))
        self.tree.heading("attrition",   text="Attrition %",    anchor="center", command=lambda: self._sort_by("attrition"))

        self.tree.column("no",          width=46,  minwidth=40,  stretch=True,  anchor="center")
        self.tree.column("part_type",   width=180, minwidth=120, stretch=True,  anchor="w")
        self.tree.column("description", width=420, minwidth=250, stretch=True,  anchor="w")
        self.tree.column("mfr_pn",      width=180, minwidth=120, stretch=True,  anchor="w")
        self.tree.column("int_pn",      width=180, minwidth=120, stretch=True,  anchor="w")
        self.tree.column("qty",         width=80,  minwidth=70,  stretch=True,  anchor="e")
        self.tree.column("attrition",   width=110, minwidth=90,  stretch=True,  anchor="center")

        # Tag colours for attrition — configured dynamically per unique rate
        self.tree.tag_configure("unknown", foreground="#be185d")
        self.tree.tag_configure("row_odd", background=CLR_ROW_ODD)

        # Scrollbars
        vsb = ttk.Scrollbar(table_frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        # Copy support: Ctrl+C = whole row(s), right-click = single cell
        self.tree.bind("<Control-c>", self._copy_selection)
        self.tree.bind("<Control-C>", self._copy_selection)
        self.tree.bind("<Button-3>", self._on_right_click)

        # ── Status bar ──
        status = ctk.CTkFrame(self, fg_color="#fef9c3", corner_radius=0, height=44)
        status.pack(fill="x", side="bottom")
        status.pack_propagate(False)

        self.lbl_status = ctk.CTkLabel(
            status, text="Ready. Open an xlsx file to begin.",
            font=ctk.CTkFont(size=15, weight="bold"), text_color=CLR_TEXT,
            anchor="w",
        )
        self.lbl_status.pack(side="left", padx=14, pady=4)

    # ── Event handlers ─────────────────────────────────────────────────────────
    def _on_browse(self):
        path = filedialog.askopenfilename(
            title="Select BOM Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not path:
            return
        self._filepath = path
        fname = os.path.basename(path)
        self.lbl_file.configure(text=fname, text_color=CLR_TEXT)
        self.btn_process.configure(state="normal")
        self.lbl_status.configure(
            text=f"File loaded: {fname}  –  Click ▶ Process to analyse.",
            text_color=CLR_ACCENT2,
        )

    def _on_open_editor(self):
        """Open the Attrition Rules editor window."""
        AttritionEditorWindow(self, on_save_callback=self._on_rules_saved)

    def _on_open_dict_editor(self):
        """Open the Keyword Dictionary editor window."""
        from dictionary_editor import DictionaryEditorWindow
        DictionaryEditorWindow(self, on_save_callback=self._on_dict_saved)

    def _on_dict_saved(self):
        """Called after dictionary editor saves/switches versions."""
        import importlib
        import bom_reader as _br
        import attrition_engine as _eng
        importlib.reload(_br)
        importlib.reload(_eng)
        # Patch module-level references in this file's scope
        global read_bom_file, get_available_sheets, analyze_row
        from bom_reader import read_bom_file, get_available_sheets  # noqa: F811
        from attrition_engine import analyze_row  # noqa: F811
        self.lbl_status.configure(
            text="  Dictionary updated. Click ▶ Process to re-apply.",
            text_color="#cba6f7",
        )

    def _on_rules_saved(self):
        """Called after editor saves rules. If BOM already processed, re-run."""
        import importlib
        import attrition_engine as _eng
        importlib.reload(_eng)
        # Patch the module-level function reference in this file's scope
        global analyze_row
        from attrition_engine import analyze_row  # noqa: F811
        self.lbl_status.configure(
            text="  Attrition rules updated. Click ▶ Process to re-apply.",
            text_color="#cba6f7",
        )

    def _on_process(self):
        if not self._filepath:
            return
        self.btn_process.configure(state="disabled", text="⏳ Processing…")
        self.lbl_status.configure(text="Reading file…", text_color=CLR_MUTED)
        # Run in thread so UI stays responsive
        threading.Thread(target=self._process_thread, daemon=True).start()

    def _process_thread(self):
        try:
            rows, col_map, sheet, header_values = read_bom_file(self._filepath)
            if not rows:
                self.after(0, self._show_error,
                           "No suitable sheet found.\n"
                           "Make sure row 1 contains column headers "
                           "(Description / Part Type / Qty / Unit).")
                return

            # Store raw rows for column remapping
            self._raw_rows = rows
            self._col_map = col_map

            results = []
            for r in rows:
                analysis = analyze_row(
                    part_type=r["part_type"],
                    description=r["description"],
                    qty=r["quantity"],
                    unit=r["unit"],
                )
                results.append({
                    "_id": len(results),
                    "part_type_raw":  r["part_type"],
                    "part_type_full": analysis["full_name"],
                    "description":    r["description"],
                    "attrition_rate": analysis["attrition_rate"],
                    "attrition_pct":  analysis["attrition_pct"],
                    "canonical_type": analysis["canonical_type"],
                    "resolved_via":   analysis["resolved_via"],
                    "qty_bom":        r.get("quantity", 0),
                    "unit":           r.get("unit", "EA"),
                    "mpn":            r.get("mpn"),
                    "internal_pn":    r.get("internal_pn"),
                })

            self._all_rows = results
            self.after(0, self._populate_table, results, sheet, col_map, header_values)

        except Exception as exc:
            self.after(0, self._show_error, str(exc))

    def _populate_table(self, results: list[dict], sheet: str, col_map: dict, header_values: list):
        # Store header values for column remapping
        self._header_values = header_values
        # Populate column selector dropdowns
        self._update_pn_column_combos(col_map, header_values)

        # Apply current filter
        self._render_rows(results)
        matched   = len(results)
        unmatched = sum(1 for r in results if r["canonical_type"] == "UNKNOWN")
        fname = os.path.basename(self._filepath)
        self.lbl_status.configure(
            text=f"  {fname}  |  Sheet: '{sheet}'  |  "
                 f"{matched} rows  |  {unmatched} unrecognised",
            text_color=CLR_ACCENT2,
        )
        self.btn_process.configure(state="normal", text="▶  Process")
        self.btn_export.configure(state="normal")

    def _on_int_pn_col_change(self, selected: str):
        """Re-extract Internal P/N from selected column."""
        self._remap_pn_column(selected, "internal_pn")
        self._render_rows()

    def _on_mfr_pn_col_change(self, selected: str):
        """Re-extract MFR P/N from selected column."""
        self._remap_pn_column(selected, "mpn")
        self._render_rows()

    def _remap_pn_column(self, selected_header: str, target_field: str):
        """Re-extract a P/N field from the selected column in raw rows."""
        if not hasattr(self, "_raw_rows") or not self._raw_rows:
            return
        if selected_header == "Auto":
            self._reprocess_all_rows()
            return

        if not hasattr(self, "_header_values") or not self._header_values:
            return

        # Find column index from actual header row
        headers = [str(v) if v is not None else f"Col_{i}" for i, v in enumerate(self._header_values)]
        try:
            col_idx = headers.index(selected_header)
        except ValueError:
            return

        # Extract values from selected column for each row
        for i, row_dict in enumerate(self._raw_rows):
            raw_row = row_dict["raw_row"]
            if col_idx < len(raw_row):
                val = raw_row[col_idx]
                if val is not None and str(val).strip():
                    self._all_rows[i][target_field] = str(val).strip()
                else:
                    self._all_rows[i][target_field] = None

    def _reprocess_all_rows(self):
        """Full re-process using original column detection."""
        if not hasattr(self, "_raw_rows") or not self._raw_rows:
            return
        # Re-run analysis on all raw rows
        for i, r in enumerate(self._raw_rows):
            analysis = analyze_row(
                part_type=r["part_type"],
                description=r["description"],
                qty=r["quantity"],
                unit=r["unit"],
            )
            self._all_rows[i].update({
                "part_type_raw":  r["part_type"],
                "part_type_full": analysis["full_name"],
                "description":    r["description"],
                "attrition_rate": analysis["attrition_rate"],
                "attrition_pct":  analysis["attrition_pct"],
                "canonical_type": analysis["canonical_type"],
                "resolved_via":   analysis["resolved_via"],
                "qty_bom":        r.get("quantity", 0),
                "unit":           r.get("unit", "EA"),
                "mpn":            r.get("mpn"),
                "internal_pn":    r.get("internal_pn"),
            })

    def _update_pn_column_combos(self, col_map: dict, header_values: list):
        """Build list of available columns for P/N selectors."""
        if not header_values:
            return
        # Use provided header values directly
        headers = [str(v) if v is not None else f"Col_{i}" for i, v in enumerate(header_values)]

        # Mark detected columns
        detected = {}
        for role, idx in col_map.items():
            if idx < len(headers):
                detected[role] = headers[idx]

        # Build display list: "Auto" + header names
        options = ["Auto"] + headers

        # Update Internal P/N combo
        self.int_pn_col_combo.configure(values=options)
        # Set default: prefer internal_pn_col, then partnumber_col
        default_int = detected.get("internal_pn_col") or detected.get("partnumber_col") or "Auto"
        if default_int in options:
            self.int_pn_col_var.set(default_int)
        else:
            self.int_pn_col_var.set("Auto")

        # Update MFR P/N combo
        self.mfr_pn_col_combo.configure(values=options)
        default_mfr = detected.get("mpn_col") or detected.get("partnumber_col") or "Auto"
        if default_mfr in options:
            self.mfr_pn_col_var.set(default_mfr)
        else:
            self.mfr_pn_col_var.set("Auto")

    def _render_rows(self, results: list[dict] | None = None):
        if results is None:
            results = self._all_rows

        ft = self.filter_var.get().strip().lower()
        af = self.att_filter_var.get()

        # First pass: collect visible rows + unique pct values
        visible = []
        pct_set: set[float] = set()
        for r in results:
            if ft and ft not in r["description"].lower() and ft not in r["part_type_full"].lower():
                continue
            if af == "Unknown":
                if r["canonical_type"] != "UNKNOWN":
                    continue
            elif af != "All" and r["attrition_pct"] != af:
                continue
            visible.append(r)
            if r["canonical_type"] != "UNKNOWN":
                pct_set.add(round(r["attrition_rate"] * 100, 1))

        # Configure a tag for each unique pct (gradient)
        for pct in pct_set:
            tag = f"att_{pct:.1f}"
            if not self.tree.tag_has(tag):
                self.tree.tag_configure(tag, foreground=_att_gradient_hex(pct))

        self.tree.delete(*self.tree.get_children())
        self._visible_rows = []
        for idx, r in enumerate(visible, start=1):
            if r["canonical_type"] == "UNKNOWN":
                tag = "unknown"
            else:
                pct = round(r["attrition_rate"] * 100, 1)
                tag = f"att_{pct:.1f}"
            tags = (tag,) if idx % 2 == 0 else (tag, "row_odd")

            # Part Type display
            if r.get("canonical_type") == "UNKNOWN":
                pt_display = "???"
            elif r.get("canonical_type") == "OTHER_SPECIAL" and r.get("part_type_raw"):
                pt_display = str(r["part_type_raw"]).strip().title()
            else:
                pt_display = r["part_type_full"]

            qty_val = float(r.get("qty_bom", 0) or 0)
            qty_display = str(int(qty_val)) if qty_val == int(qty_val) else f"{qty_val:g}"

            self.tree.insert(
                "", "end", iid=str(r["_id"]),
                values=(idx, pt_display, r["description"],
                        r.get("mpn") or "", r.get("internal_pn") or "",
                        qty_display, r["attrition_pct"]),
                tags=tags,
            )
            self._visible_rows.append(r)

        total = len(self._all_rows)
        shown = len(self._visible_rows)
        self.lbl_count.configure(text=f"Showing {shown} / {total} rows")

        # Update attrition filter dropdown with actual rates present
        if hasattr(self, "att_combo"):
            rate_vals = ["All", "Unknown"] + sorted(f"{p:.1f}%" for p in pct_set)
            self.att_combo.configure(values=rate_vals)

    def _rate_tag(self, rate: float, ctype: str) -> str:
        if ctype == "UNKNOWN":
            return "unknown"
        pct = round(rate * 100, 1)
        return f"att_{pct:.1f}"

    def _on_filter_change(self, *_):
        self._render_rows()

    def _show_error(self, msg: str):
        self.btn_process.configure(state="normal", text="▶  Process")
        self.btn_export.configure(state="normal")
        self.lbl_status.configure(text=f"Error: {msg}", text_color=CLR_WARN)
        # Also show popup
        win = ctk.CTkToplevel(self)
        win.title("Error")
        win.geometry("460x180")
        win.grab_set()
        ctk.CTkLabel(win, text="⚠  Error", font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=CLR_WARN).pack(pady=(20, 8))
        ctk.CTkLabel(win, text=msg, wraplength=400,
                     font=ctk.CTkFont(size=14)).pack(padx=20)
        ctk.CTkButton(win, text="OK", command=win.destroy,
                      width=80).pack(pady=16)

    def _sort_by(self, col):
        if self._sort_col == col:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_reverse = False
        self._sort_col = col
        
        # Sort _all_rows based on the column
        if col == "no":
            self._all_rows.sort(key=lambda r: r["_id"], reverse=self._sort_reverse)
        elif col == "part_type":
            self._all_rows.sort(key=lambda r: r["part_type_full"], reverse=self._sort_reverse)
        elif col == "description":
            self._all_rows.sort(key=lambda r: r["description"] or "", reverse=self._sort_reverse)
        elif col == "mfr_pn":
            self._all_rows.sort(key=lambda r: r.get("mpn") or "", reverse=self._sort_reverse)
        elif col == "int_pn":
            self._all_rows.sort(key=lambda r: r.get("internal_pn") or "", reverse=self._sort_reverse)
        elif col == "qty":
            self._all_rows.sort(key=lambda r: float(r.get("qty_bom", 0) or 0), reverse=self._sort_reverse)
        elif col == "attrition":
            self._all_rows.sort(key=lambda r: r["attrition_rate"], reverse=self._sort_reverse)
            
        self._render_rows()

    def _on_export(self):
        if not hasattr(self, "_visible_rows") or not self._visible_rows:
            return
        path = filedialog.asksaveasfilename(
            title="Export Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not path:
            return
            
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "BOM Export"
        
        headers = ["Row", "Original Part Type", "Canonical Type", "Description",
                   "MFR Part Number", "Internal Part Number",
                   "BOM Qty", "Unit", "Attrition %", "Attrition Rate", "Final Qty"]
        ws.append(headers)
        
        for r in self._visible_rows:
            raw_pt = r.get("part_type_raw", "")
            can_pt = r.get("part_type_full", "")
            if r.get("canonical_type") == "OTHER_SPECIAL" and raw_pt:
                can_pt = str(raw_pt).strip().title()
                
            qty = float(r.get("qty_bom", 0) or 0)
            rate = float(r.get("attrition_rate", 0))
            final_qty = apply_attrition(qty, rate, r.get("unit", "EA"))
            
            ws.append([
                r["_id"] + 1,
                raw_pt,
                can_pt,
                r.get("description", ""),
                r.get("mpn") or "",
                r.get("internal_pn") or "",
                qty,
                r.get("unit", "EA"),
                r.get("attrition_pct", "0.0%"),
                rate,
                final_qty
            ])
            
        wb.save(path)
        self.lbl_status.configure(text=f"Exported successfully to {os.path.basename(path)}", text_color="#10b981")

    def _copy_selection(self, event=None):
        """Copy selected row(s) to clipboard, cells separated by tabs."""
        items = self.tree.selection()
        if not items:
            return "break"
        lines = ["\t".join(str(v) for v in self.tree.item(i, "values")) for i in items]
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))
        self.lbl_status.configure(
            text=f"Đã copy {len(items)} dòng vào clipboard (Ctrl+V để dán)",
            text_color="#2563eb")
        return "break"

    def _on_right_click(self, event):
        """Right-click a cell → copy that cell's value."""
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row_id or not col:
            return
        idx = int(col[1:]) - 1
        values = self.tree.item(row_id, "values")
        if idx >= len(values):
            return
        menu = tk.Menu(self, tearoff=0)
        cell_val = str(values[idx])
        menu.add_command(
            label=f"Copy: {cell_val[:40]}{'…' if len(cell_val) > 40 else ''}",
            command=lambda: (
                self.clipboard_clear(),
                self.clipboard_append(cell_val),
                self.lbl_status.configure(
                    text=f"Đã copy vào clipboard: {cell_val[:60]}",
                    text_color="#2563eb"),
            ))
        menu.tk_popup(event.x_root, event.y_root)
        return "break"

    def _on_tree_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell": return
        column = self.tree.identify_column(event.x)
        if column != "#2": return # Only Part Type
        
        item_id = self.tree.identify_row(event.y)
        if not item_id: return
        
        x, y, w, h = self.tree.bbox(item_id, column)
        
        # Get current display value
        r_id = int(item_id)
        target_row = next((r for r in self._all_rows if r["_id"] == r_id), None)
        if not target_row: return
        
        # Build dropdown
        cb = ttk.Combobox(self.tree, values=list(FRIENDLY_TYPES.keys()), state="readonly")
        cb.place(x=x, y=y, width=w, height=h)
        
        # Find current friendly name if possible
        reverse_map = {v: k for k, v in FRIENDLY_TYPES.items()}
        curr_friendly = reverse_map.get(target_row["canonical_type"], "")
        if curr_friendly:
            cb.set(curr_friendly)
            
        def on_select(e):
            sel = cb.get()
            cb.destroy()
            if not sel: return
            can_type = FRIENDLY_TYPES.get(sel)
            if not can_type: return
            
            # Update row
            from attrition_engine import get_attrition_rate
            target_row["canonical_type"] = can_type
            
            if can_type == "UNKNOWN":
                target_row["part_type_full"] = "Unknown"
                target_row["attrition_rate"] = 0.0
                target_row["attrition_pct"] = "0.0%"
            else:
                target_row["part_type_full"] = sel
                rate = get_attrition_rate(can_type, "-",
                                          target_row.get("unit", "EA"),
                                          target_row.get("description"))
                target_row["attrition_rate"] = rate
                target_row["attrition_pct"] = f"{rate*100:.1f}%"
                
            self._render_rows()
            self.lbl_status.configure(text=f"Updated row {r_id+1} to {sel} ({target_row['attrition_pct']})", text_color="#10b981")
            
        cb.bind("<<ComboboxSelected>>", on_select)
        cb.bind("<FocusOut>", lambda e: cb.destroy())
        cb.focus_set()
