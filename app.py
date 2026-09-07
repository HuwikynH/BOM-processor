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
import re

from bom_reader import read_bom_file, get_available_sheets
from attrition_engine import analyze_row, apply_attrition
from attrition_editor import AttritionEditorWindow

# ── Searchable Combobox Helper ──────────────────────────────────────────────────
class SearchableCombobox(tk.Entry):
    """
    Autocomplete entry that keeps keyboard focus while showing suggestions.
    Native ttk.Combobox popdowns take focus on Windows, which hides the caret.
    """
    def __init__(self, master=None, **kwargs):
        values = kwargs.pop("values", [])
        kwargs.pop("state", None)
        super().__init__(master, **kwargs)
        self._all_values = list(values)
        self._filtered_values = list(values)
        self._filter_after_id = None
        self._popup = None
        self._listbox = None
        self.bind("<KeyRelease>", self._on_keyrelease)
        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        self.bind("<Down>", self._on_down)
        self.bind("<Up>", self._on_up)

    def _on_keyrelease(self, event):
        if event.keysym in {"Up", "Down", "Left", "Right", "Return", "Escape", "Tab"}:
            return
        if self._filter_after_id:
            self.after_cancel(self._filter_after_id)
        self._filter_after_id = self.after(80, self._filter)

    def _filter(self):
        typed = self.get().strip().lower()
        if not typed:
            self._filtered_values = self._all_values
        else:
            self._filtered_values = [
                v for v in self._all_values if typed in str(v).lower()
            ]
        self._show_popup()

    def _show_popup(self):
        if not self._filtered_values:
            self._hide_popup()
            return
        if self._popup is None or not self._popup.winfo_exists():
            self._popup = tk.Toplevel(self)
            self._popup.overrideredirect(True)
            self._popup.transient(self.winfo_toplevel())
            self._listbox = tk.Listbox(
                self._popup,
                height=min(8, len(self._filtered_values)),
                font=self.cget("font"),
                bg="#ffffff",
                fg="#18181b",
                selectbackground="#2563eb",
                selectforeground="#ffffff",
                activestyle="none",
                exportselection=False,
            )
            self._listbox.pack(fill="both", expand=True)
            self._listbox.bind("<ButtonRelease-1>", self._select_from_popup)
            self._listbox.bind("<Return>", self._select_from_popup)
            self._listbox.bind("<Escape>", lambda e: self._hide_popup())
        self._listbox.delete(0, "end")
        for value in self._filtered_values:
            self._listbox.insert("end", value)
        self._listbox.configure(height=min(8, len(self._filtered_values)))
        self._popup.update_idletasks()
        x            = self.winfo_rootx()
        entry_bottom = self.winfo_rooty() + self.winfo_height()
        entry_top    = self.winfo_rooty()
        width        = max(self.winfo_width(), 180)
        popup_h      = self._listbox.winfo_reqheight()
        screen_h     = self.winfo_screenheight()
        # Flip upward when there is not enough space below the widget
        if entry_bottom + popup_h > screen_h and entry_top - popup_h >= 0:
            y = entry_top - popup_h
        else:
            y = entry_bottom
        self._popup.geometry(f"{width}x{popup_h}+{x}+{y}")
        self._popup.deiconify()
        self._take_focus()

    def _take_focus(self):
        def apply_focus():
            if not self.winfo_exists():
                return
            try:
                self.focus_force()
                self.icursor("end")
            except Exception:
                pass
        self.after_idle(apply_focus)

    def _hide_popup(self):
        if self._popup is not None and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None
        self._listbox = None

    def _select_from_popup(self, event=None):
        if self._listbox is None:
            return "break"
        selection = self._listbox.curselection()
        if not selection:
            return "break"
        self.set(self._listbox.get(selection[0]))
        self._hide_popup()
        self.event_generate("<<ComboboxSelected>>")
        return "break"

    def get(self):
        if self._listbox is not None:
            selection = self._listbox.curselection()
            if selection:
                return self._listbox.get(selection[0])
        return super().get()

    def _on_focus_in(self, event):
        if not self.get():
            self._filtered_values = self._all_values
            self._show_popup()

    def _on_focus_out(self, event):
        if self._filter_after_id:
            self.after_cancel(self._filter_after_id)

    def _on_down(self, event=None):
        if self._popup is None or not self._popup.winfo_exists():
            self._filter()
        if self._listbox is not None and self._listbox.size():
            current = self._listbox.curselection()
            idx = current[0] + 1 if current and current[0] + 1 < self._listbox.size() else 0
            self._listbox.selection_clear(0, "end")
            self._listbox.selection_set(idx)
            self._listbox.activate(idx)
        return "break"

    def _on_up(self, event=None):
        if self._listbox is not None and self._listbox.size():
            current = self._listbox.curselection()
            idx = current[0] - 1 if current and current[0] > 0 else self._listbox.size() - 1
            self._listbox.selection_clear(0, "end")
            self._listbox.selection_set(idx)
            self._listbox.activate(idx)
        return "break"

    def set_values(self, values):
        """Update the complete list of values."""
        self._all_values = list(values)
        self._filtered_values = list(values)

    def set(self, value):
        self.delete(0, "end")
        self.insert(0, value)
        self.icursor("end")

    def destroy(self):
        self._hide_popup()
        super().destroy()

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


def _build_friendly_with_packages() -> list[str]:
    """
    Build dropdown options with package variants from active attrition rules.
    Returns list like: ["Resistor [0201]", "Resistor [0402]", ..., "Capacitor [0402]", ...]
    Falls back to FRIENDLY_TYPES if rules not available.
    """
    try:
        from attrition_engine import RULES
        smt = RULES.get("smt_rules", {})
        cable = RULES.get("cable_box_rules", {})
    except Exception:
        return list(FRIENDLY_TYPES.keys())

    # Map canonical type -> list of packages (excluding _default, _comment)
    pkg_map: dict[str, list[str]] = {}
    for table in (smt, cable):
        for comp_type, entry in table.items():
            if comp_type.startswith("_"):
                continue
            if isinstance(entry, dict):
                pkgs = [p for p in entry.keys() if not p.startswith("_")]
                if pkgs:
                    pkg_map.setdefault(comp_type, []).extend(pkgs)

    # Build options: friendly_name [package] for each package
    options = []
    for friendly, canon in FRIENDLY_TYPES.items():
        pkgs = sorted(set(pkg_map.get(canon, [])))
        if pkgs:
            for pkg in pkgs:
                options.append(f"{friendly} [{pkg}]")
        else:
            options.append(friendly)

    return options

class BOMApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("BOM Processor  v1.5")
        self.geometry("1100x680")
        self.minsize(900, 540)
        self.configure(fg_color=CLR_BG)
        self.after(0, lambda: self.state("zoomed") if hasattr(self, "state") else None)

        # Custom icon
        _icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
        if os.path.exists(_icon_path):
            try:
                self.iconbitmap(_icon_path)
            except Exception:
                pass

        self._filepath: str | None = None
        self._all_rows: list[dict] = []   # processed result rows
        self._filter_text: str = ""
        self._sort_col = None
        self._sort_reverse = False

        # P/N column selection state (for header dropdown menus)
        self.int_pn_col_var = tk.StringVar(value="Auto")
        self.mfr_pn_col_var = tk.StringVar(value="Auto")
        self._active_cell_editor = None

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
        self.tree.heading("mfr_pn",      text="MFR P/N  ▼",     anchor="w",     command=lambda: self._show_pn_column_menu("mfr_pn"))
        self.tree.heading("int_pn",      text="Internal P/N  ▼", anchor="w",     command=lambda: self._show_pn_column_menu("int_pn"))
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
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
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
                manual_pkg = None
                # Check if existing row has manual package (for re-process)
                existing = next((ar for ar in self._all_rows if ar.get("_id") == len(results)), None)
                if existing and existing.get("_manual_package"):
                    manual_pkg = existing["_manual_package"]
                analysis = analyze_row(
                    part_type=r["part_type"],
                    description=r["description"],
                    qty=r["quantity"],
                    unit=r["unit"],
                    manual_package=manual_pkg,
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
            manual_pkg = self._all_rows[i].get("_manual_package") if i < len(self._all_rows) else None
            analysis = analyze_row(
                part_type=r["part_type"],
                description=r["description"],
                qty=r["quantity"],
                unit=r["unit"],
                manual_package=manual_pkg,
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

    def _show_pn_column_menu(self, col_key: str):
        """Show dropdown menu for P/N column selection at the column header."""
        if not hasattr(self, "_header_values") or not self._header_values:
            return

        headers = [str(v) if v is not None else f"Col_{i}" for i, v in enumerate(self._header_values)]

        # Detect current column mapping
        detected = {}
        if hasattr(self, "_col_map"):
            for role, idx in self._col_map.items():
                if idx < len(headers):
                    detected[role] = headers[idx]

        detected_roles = {
            "mfr_pn": ("mpn_col", "partnumber_col"),
            "int_pn": ("internal_pn_col", "partnumber_col"),
        }.get(col_key, ())

        # Create dropdown menu
        menu = tk.Menu(self, tearoff=0)
        menu.configure(
            bg="#ffffff", fg="#18181b",
            activebackground="#16a34a", activeforeground="#ffffff",
            font=("Segoe UI", 11), bd=1, relief="solid"
        )

        # Current selection
        current_var = self.mfr_pn_col_var if col_key == "mfr_pn" else self.int_pn_col_var
        current_val = current_var.get()

        # Add "Auto" option
        menu.add_command(
            label="⟲  Auto (tự động detect)",
            command=lambda: self._on_pn_col_menu_select(col_key, "Auto"),
            font=("Segoe UI", 11, "bold" if current_val == "Auto" else "normal")
        )
        menu.add_command(
            label="Sort A → Z",
            command=lambda: self._sort_by(col_key, reverse=False),
        )
        menu.add_command(
            label="Sort Z → A",
            command=lambda: self._sort_by(col_key, reverse=True),
        )
        menu.add_separator()

        # Add each header as option
        for h_str in headers:
            is_current = (h_str == current_val)
            is_detected = any(detected.get(role) == h_str for role in detected_roles)
            
            label = f"  {h_str}"
            if is_detected:
                label += "  ✓ (auto-detect)"
            if is_current:
                label = "✓ " + label
            
            menu.add_command(
                label=label,
                command=lambda h=h_str: self._on_pn_col_menu_select(col_key, h),
                font=("Segoe UI", 11, "bold" if is_current else "normal")
            )

        # Show menu at cursor position
        try:
            x = self.winfo_pointerx()
            y = self.winfo_pointery()
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def _on_pn_col_menu_select(self, col_key: str, selected: str):
        """Handle selection from P/N column dropdown menu."""
        if col_key == "mfr_pn":
            self.mfr_pn_col_var.set(selected)
            self._on_mfr_pn_col_change(selected)
        else:
            self.int_pn_col_var.set(selected)
            self._on_int_pn_col_change(selected)

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

    def _sort_by(self, col, reverse: bool | None = None):
        if reverse is not None:
            self._sort_reverse = reverse
        elif self._sort_col == col:
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
        from tkinter import filedialog
        import openpyxl
        import os
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
        
        mfr_col_name = self.mfr_pn_col_var.get()
        if mfr_col_name == "Auto":
            mfr_col_name = "MFR P/N"
            
        int_col_name = self.int_pn_col_var.get()
        if int_col_name == "Auto":
            int_col_name = "Internal P/N"
            
        headers = ["#", "Part Type", "Description (original)",
                   mfr_col_name, int_col_name, "Qty", "Attrition %"]
        ws.append(headers)
        
        for idx, r in enumerate(self._visible_rows, start=1):
            if r.get("canonical_type") == "UNKNOWN":
                pt_display = "???"
            elif r.get("canonical_type") == "OTHER_SPECIAL" and r.get("part_type_raw"):
                pt_display = str(r.get("part_type_raw", "")).strip().title()
            else:
                pt_display = r.get("part_type_full", "")
                
            qty_val = float(r.get("qty_bom", 0) or 0)
            qty_display = int(qty_val) if qty_val == int(qty_val) else qty_val
            
            ws.append([
                idx,
                pt_display,
                r.get("description", ""),
                r.get("mpn") or "",
                r.get("internal_pn") or "",
                qty_display,
                r.get("attrition_pct", "0.0%")
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

        self._destroy_active_cell_editor()
        
        # Build dropdown with package variants
        options = _build_friendly_with_packages()
        cb = SearchableCombobox(self.tree, values=options, state="normal")
        self._active_cell_editor = cb
        cb.place(x=x, y=y, width=w, height=h)
        
        # Find current friendly name (with package) if possible
        current_full = target_row.get("part_type_full", "")
        reverse_map = {v: k for k, v in FRIENDLY_TYPES.items()}
        curr_friendly = reverse_map.get(target_row["canonical_type"], "")
        if curr_friendly and current_full:
            # If current has package like "Resistor [0805]", use it
            cb.set(current_full)
        elif curr_friendly:
            cb.set(curr_friendly)
            
        def on_select(e):
            sel = cb.get()
            if not sel:
                cb.destroy()
                return
            # Parse selection: "Resistor [0805]" -> canon=RESISTOR, pkg=0805
            import re
            m = re.match(r"^(.+?)\s*\[(.+?)\]$", sel.strip())
            if m:
                friendly, pkg = m.group(1).strip(), m.group(2).strip()
                can_type = FRIENDLY_TYPES.get(friendly)
                new_pkg = pkg
            else:
                friendly = sel.strip()
                can_type = FRIENDLY_TYPES.get(friendly)
                new_pkg = None
            if not can_type:
                cb.destroy()
                return
            
            # Update row
            from attrition_engine import get_attrition_rate
            target_row["canonical_type"] = can_type
            if new_pkg:
                # Store package for attrition lookup
                target_row["_manual_package"] = new_pkg
            else:
                target_row.pop("_manual_package", None)
            
            if can_type == "UNKNOWN":
                target_row["part_type_full"] = "Unknown"
                target_row["attrition_rate"] = 0.0
                target_row["attrition_pct"] = "0.0%"
            else:
                target_row["part_type_full"] = sel
                rate = get_attrition_rate(can_type, new_pkg,
                                          target_row.get("unit", "EA"),
                                          target_row.get("description"))
                target_row["attrition_rate"] = rate
                target_row["attrition_pct"] = f"{rate*100:.1f}%"

            self._destroy_active_cell_editor(cb)
            self._render_rows()
            self.lbl_status.configure(text=f"Updated row {r_id+1} to {sel} ({target_row['attrition_pct']})", text_color="#10b981")
            
        cb.bind("<<ComboboxSelected>>", on_select)
        cb.bind("<Return>", on_select)
        cb.bind("<KP_Enter>", on_select)
        cb.bind("<Tab>", on_select)
        cb.bind("<Escape>", lambda e: self._destroy_active_cell_editor(cb))
        cb.focus_set()
        cb._take_focus()

    def _destroy_active_cell_editor(self, editor=None):
        active = self._active_cell_editor
        if editor is not None and active is not editor:
            return
        if active is not None:
            try:
                if active.winfo_exists():
                    active.destroy()
            except Exception:
                pass
        self._active_cell_editor = None
