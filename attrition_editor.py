"""
attrition_editor.py
--------------------
Popup window to VIEW and EDIT attrition rates.

Layout:
  ┌─────────────────────────────────────┐
  │  ⚙ Attrition Rules Editor          │
  │  [SMT] [Cable / Box]               │
  │  ┌──────────────┬────────┬───────┐  │
  │  │ Component    │ Package│  Att% │  │
  │  ├──────────────┼────────┼───────┤  │
  │  │ Resistor     │ 0201   │  10%  │  │
  │  │ Resistor     │ 0402   │  10%  │  │
  │  │ …            │ …      │  …    │  │
  │  └──────────────┴────────┴───────┘  │
  │  Double-click Att% cell to edit.    │
  │  [Reset]                   [Save]   │
  └─────────────────────────────────────┘
"""

import json
import os
import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_RULES_PATH = os.path.join(_BASE_DIR, "attrition_rules.json")

# ── Colours (match main app) ──────────────────────────────────────────────────
CLR_BG        = "#ffffff"
CLR_SURFACE   = "#f4f4f5"
CLR_ACCENT    = "#16a34a"
CLR_TEXT      = "#18181b"
CLR_MUTED     = "#52525b"
CLR_WARN      = "#dc2626"
CLR_GREEN     = "#16a34a"
CLR_HEADER_BG = "#e4e4e7"
CLR_ROW_EVEN  = "#f4f4f5"
CLR_ROW_ODD   = "#ffffff"
CLR_EDIT_HL   = "#d97706"   # orange highlight for edited cells


def _load_rules() -> dict:
    with open(_RULES_PATH, encoding="utf-8") as f:
        return json.load(f)


def _save_rules(rules: dict):
    with open(_RULES_PATH, "w", encoding="utf-8") as f:
        json.dump(rules, f, indent=2, ensure_ascii=False)


def _flatten_smt(smt_rules: dict) -> list[tuple]:
    """
    Returns list of (component_type, package_label, rate_float, json_key_path)
    json_key_path = ("smt_rules", comp_type, pkg_key)
    """
    rows = []
    for comp, entry in smt_rules.items():
        if comp.startswith("_") or not isinstance(entry, dict):
            continue
        for pkg, val in entry.items():
            if pkg.startswith("_"):
                # _default → show as "(default)", _comment → skip
                if pkg == "_default":
                    pkg_label = "(default)"
                else:
                    continue
            else:
                pkg_label = pkg
            try:
                rate = float(val)
            except (TypeError, ValueError):
                continue   # skip non-numeric values like _comment strings
            rows.append((comp, pkg_label, rate, ("smt_rules", comp, pkg)))
    return rows


def _flatten_cable(cable_rules: dict) -> list[tuple]:
    """
    Returns list of (component_type, "(all)", rate_float, json_key_path)
    """
    rows = []
    for comp, entry in cable_rules.items():
        if comp.startswith("_") or not isinstance(entry, dict):
            continue
        val = entry.get("_default", 0.0)
        try:
            rate = float(val)
        except (TypeError, ValueError):
            rate = 0.0
        rows.append((comp, "(all)", rate, ("cable_box_rules", comp, "_default")))
    return rows


# ── Editor window ─────────────────────────────────────────────────────────────

class AttritionEditorWindow(ctk.CTkToplevel):
    def __init__(self, parent, on_save_callback=None):
        super().__init__(parent)
        self.title("⚙  Attrition Rules Editor")
        self.geometry("660x540")
        self.minsize(560, 400)
        self.configure(fg_color=CLR_BG)
        self.grab_set()   # modal

        self._on_save_callback = on_save_callback
        self._rules: dict = _load_rules()
        self._dirty: bool = False          # unsaved changes flag
        self._edit_entry: tk.Entry | None = None  # floating entry widget

        self._build_ui()
        self._populate_current_tab()

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Title bar
        hdr = ctk.CTkFrame(self, fg_color=CLR_SURFACE, corner_radius=0, height=64)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="⚙  Attrition Rules Editor",
                     font=ctk.CTkFont(size=20, weight="bold"),
                     text_color=CLR_ACCENT).pack(side="left", padx=16, pady=10)
        self.lbl_hint = ctk.CTkLabel(
            hdr, text="Double-click  Att %  to edit",
            font=ctk.CTkFont(size=12, slant="italic"), text_color=CLR_MUTED)
        self.lbl_hint.pack(side="right", padx=16)

        # Tab bar
        tab_bar = ctk.CTkFrame(self, fg_color=CLR_SURFACE, corner_radius=0, height=48)
        tab_bar.pack(fill="x")
        tab_bar.pack_propagate(False)

        self._active_tab = tk.StringVar(value="smt")
        for label, key in [("  🔲  SMT / PCBA  ", "smt"),
                            ("  🔌  Cable / Box  ", "cable")]:
            btn = ctk.CTkButton(
                tab_bar, text=label, height=32, corner_radius=0,
                fg_color=CLR_ACCENT if key == "smt" else CLR_SURFACE,
                hover_color="#15803d",
                command=lambda k=key: self._switch_tab(k),
            )
            btn.pack(side="left", padx=2, pady=3)
            setattr(self, f"_tab_btn_{key}", btn)

        # Table frame
        self._table_frame = ctk.CTkFrame(self, fg_color=CLR_BG, corner_radius=0)
        self._table_frame.pack(fill="both", expand=True)
        self._build_tree()

        # Bottom bar
        btm = ctk.CTkFrame(self, fg_color=CLR_SURFACE, corner_radius=0, height=56)
        btm.pack(fill="x", side="bottom")
        btm.pack_propagate(False)

        self.lbl_status = ctk.CTkLabel(btm, text="", font=ctk.CTkFont(size=13),
                                       text_color=CLR_MUTED, anchor="w")
        self.lbl_status.pack(side="left", padx=14, pady=8)

        ctk.CTkButton(btm, text="💾  Save Changes", width=148, height=36,
                      fg_color=CLR_ACCENT, hover_color="#15803d", text_color="#ffffff",
                      command=self._on_save).pack(side="right", padx=14, pady=8)
        ctk.CTkButton(btm, text="↺  Reset", width=90, height=36,
                      fg_color=CLR_SURFACE, hover_color="#d4d4d8", text_color="#18181b",
                      border_width=1, border_color=CLR_MUTED,
                      command=self._on_reset).pack(side="right", padx=4, pady=8)

    def _build_tree(self):
        """Create / recreate the Treeview."""
        # Destroy old tree if rebuilding
        for w in self._table_frame.winfo_children():
            w.destroy()

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Ed.Treeview",
                        background=CLR_ROW_EVEN, foreground=CLR_TEXT,
                        rowheight=40, fieldbackground=CLR_ROW_EVEN,
                        borderwidth=0, font=("Segoe UI", 13))
        style.configure("Ed.Treeview.Heading",
                        background=CLR_HEADER_BG, foreground=CLR_TEXT,
                        relief="flat", font=("Segoe UI", 13, "bold"))
        style.map("Ed.Treeview",
                  background=[("selected", "#bbf7d0")],
                  foreground=[("selected", "#000000")])

        cols = ("component", "package", "attrition")
        self.tree = ttk.Treeview(self._table_frame, columns=cols,
                                  show="headings", style="Ed.Treeview",
                                  selectmode="browse")

        self.tree.heading("component",  text="Component Type",  anchor="w")
        self.tree.heading("package",    text="Package / Size",  anchor="center")
        self.tree.heading("attrition",  text="Attrition %",     anchor="center")

        self.tree.column("component",  width=240, minwidth=160, stretch=True,  anchor="w")
        self.tree.column("package",    width=150, minwidth=100, stretch=False, anchor="center")
        self.tree.column("attrition",  width=130, minwidth=90,  stretch=False, anchor="center")

        # Tag colours
        self.tree.tag_configure("edited",   foreground=CLR_EDIT_HL, font=("Segoe UI", 13, "bold"))
        self.tree.tag_configure("row_odd",  background=CLR_ROW_ODD)
        self.tree.tag_configure("zero",     foreground=CLR_MUTED)

        vsb = ttk.Scrollbar(self._table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        # Bind double-click for editing
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-1>", self._cancel_edit)

    # ── Tab switching ─────────────────────────────────────────────────────────
    def _switch_tab(self, tab: str):
        self._cancel_edit()
        self._active_tab.set(tab)
        self._tab_btn_smt.configure(fg_color=CLR_ACCENT if tab == "smt" else CLR_SURFACE)
        self._tab_btn_cable.configure(fg_color=CLR_ACCENT if tab == "cable" else CLR_SURFACE)
        self._populate_current_tab()

    def _populate_current_tab(self):
        """Fill tree with rows for current tab."""
        self.tree.delete(*self.tree.get_children())
        tab = self._active_tab.get()
        if tab == "smt":
            rows = _flatten_smt(self._rules["smt_rules"])
        else:
            rows = _flatten_cable(self._rules["cable_box_rules"])

        for idx, (comp, pkg, rate, key_path) in enumerate(rows):
            pct_str = f"{rate*100:.1f}%"
            tags = []
            if idx % 2 != 0:
                tags.append("row_odd")
            if rate == 0.0:
                tags.append("zero")
            self.tree.insert(
                "", "end",
                iid=str(idx),
                values=(comp, pkg, pct_str),
                tags=tuple(tags),
            )
            # Store key_path on the item (as hidden data via iid→key_path map)
        self._key_paths = {
            str(idx): data[3]
            for idx, data in enumerate(rows)
        }

    # ── Inline editing ────────────────────────────────────────────────────────
    def _on_double_click(self, event: tk.Event):
        """Start inline edit when user double-clicks the Attrition % column."""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        if col != "#3":       # only 3rd column (attrition)
            return
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return

        self._cancel_edit()   # close any existing editor
        self._start_edit(row_id)

    def _start_edit(self, row_id: str):
        """Overlay a small Entry widget on top of the selected cell."""
        # Get bounding box of the attrition cell (#3)
        bbox = self.tree.bbox(row_id, "#3")
        if not bbox:
            return
        x, y, w, h = bbox

        current_val = self.tree.item(row_id, "values")[2]   # e.g. "10.0%"
        # Strip % sign for editing
        edit_val = current_val.rstrip("%").strip()

        var = tk.StringVar(value=edit_val)
        entry = tk.Entry(
            self.tree,
            textvariable=var,
            font=("Segoe UI", 13, "bold"),
            bg="#ffffff",
            fg=CLR_EDIT_HL,
            insertbackground=CLR_TEXT,
            relief="flat",
            bd=2,
            highlightthickness=2,
            highlightcolor=CLR_ACCENT,
            highlightbackground=CLR_ACCENT,
            justify="center",
        )
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        entry.select_range(0, "end")

        self._edit_entry = entry
        self._edit_row_id = row_id
        self._edit_var = var

        entry.bind("<Return>",  self._commit_edit)
        entry.bind("<KP_Enter>", self._commit_edit)
        entry.bind("<Escape>",  self._cancel_edit)
        entry.bind("<Tab>",     self._commit_edit)
        entry.bind("<FocusOut>", self._commit_edit)

    def _commit_edit(self, event=None):
        """Validate and apply the edited value."""
        if self._edit_entry is None:
            return
        raw = self._edit_var.get().strip().rstrip("%").strip()
        try:
            new_pct = float(raw)
            if not (0.0 <= new_pct <= 100.0):
                raise ValueError("out of range")
        except ValueError:
            self._flash_error("Enter a number between 0 and 100")
            return

        row_id = self._edit_row_id
        new_rate = new_pct / 100.0
        key_path = self._key_paths[row_id]

        # Update JSON in memory
        section, comp, pkg = key_path
        self._rules[section][comp][pkg] = new_rate

        # Update treeview cell
        old_vals = list(self.tree.item(row_id, "values"))
        old_vals[2] = f"{new_pct:.1f}%"
        self.tree.item(row_id, values=old_vals)

        # Mark as edited
        existing_tags = list(self.tree.item(row_id, "tags"))
        if "edited" not in existing_tags:
            existing_tags.append("edited")
        if new_rate == 0.0 and "zero" not in existing_tags:
            existing_tags.append("zero")
        elif new_rate != 0.0 and "zero" in existing_tags:
            existing_tags.remove("zero")
        self.tree.item(row_id, tags=tuple(existing_tags))

        self._dirty = True
        self.lbl_status.configure(
            text=f"  Unsaved changes  –  click Save to apply.",
            text_color=CLR_EDIT_HL,
        )
        self._destroy_entry()

    def _cancel_edit(self, event=None):
        self._destroy_entry()

    def _destroy_entry(self):
        if self._edit_entry:
            self._edit_entry.destroy()
            self._edit_entry = None

    def _flash_error(self, msg: str):
        self.lbl_status.configure(text=f"  ⚠  {msg}", text_color=CLR_WARN)
        if self._edit_entry:
            self._edit_entry.configure(highlightcolor=CLR_WARN,
                                       highlightbackground=CLR_WARN)

    # ── Save / Reset ─────────────────────────────────────────────────────────
    def _on_save(self):
        self._cancel_edit()
        try:
            _save_rules(self._rules)
            self._dirty = False
            self.lbl_status.configure(
                text="  Saved successfully. Rules will apply on next Process.",
                text_color=CLR_GREEN,
            )
            # Reload engine in-place
            _reload_engine()
            if self._on_save_callback:
                self._on_save_callback()
        except Exception as exc:
            self.lbl_status.configure(
                text=f"  Save failed: {exc}", text_color=CLR_WARN)

    def _on_reset(self):
        self._cancel_edit()
        if self._dirty:
            if not messagebox.askyesno(
                "Reset",
                "Discard all unsaved changes and reload from file?",
                parent=self
            ):
                return
        self._rules = _load_rules()
        self._dirty = False
        self._populate_current_tab()
        self.lbl_status.configure(text="  Rules reloaded from file.", text_color=CLR_MUTED)

    # ── Close guard ───────────────────────────────────────────────────────────
    def destroy(self):
        if self._dirty:
            if not messagebox.askyesno(
                "Unsaved Changes",
                "You have unsaved changes.\nClose without saving?",
                parent=self
            ):
                return
        super().destroy()


# ── Engine hot-reload ─────────────────────────────────────────────────────────
def _reload_engine():
    """Force attrition_engine to reload its JSON rules from disk."""
    import attrition_engine as _eng
    import importlib
    importlib.reload(_eng)
