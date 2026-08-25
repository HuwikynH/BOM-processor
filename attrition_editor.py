"""
attrition_editor.py
--------------------
Popup window to VIEW and EDIT attrition rates.

Layout:
   ┌───────────────────────────────────────────────┐
   │  ⚙ Attrition Rules Editor                     │
   │  Version: [ dropdown ▼ ]        (hint)        │
   │  [SMT / PCBA] [Cable / Box]                   │
   │  ┌──────────────┬────────┬───────┐  Filter:  │
   │  │ Component    │ Package│  Att% │  [____]   │
   │  └──────────────┴────────┴───────┘           │
   │  Double-click Att% cell to edit.              │
   │  (component description shown at bottom)      │
   │  [💾 Save as new version]                     │
   └───────────────────────────────────────────────┘

Saving ALWAYS creates a new version file (attrition_rules_v*.json).
The original attrition_rules.json is never modified.
The dropdown switches the active version used by the engine.
"""

import glob
import json
import os
import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_RULES_BASE_NAME = "attrition_rules.json"
_RULES_BASE_PATH = os.path.join(_BASE_DIR, _RULES_BASE_NAME)
_ACTIVE_PTR_PATH = os.path.join(_BASE_DIR, "_active_rules.json")
_VERSION_GLOB = os.path.join(_BASE_DIR, "attrition_rules_v*.json")

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


# ── Rule-file management ──────────────────────────────────────────────────────
def _version_display_name(filename: str) -> str:
    """'attrition_rules.json' → 'Original (default)'
    'attrition_rules_v2026-08-25_141530.json' → 'v2026-08-25 14:15:30'"""
    if filename == _RULES_BASE_NAME:
        return "Original (default)"
    stem = os.path.splitext(filename)[0]
    v = stem.replace("attrition_rules_v", "", 1)
    try:
        date, time = v.split("_", 1)
        return f"v{date} {time[:2]}:{time[2:4]}:{time[4:6]}"
    except ValueError:
        return f"v{v}"


def _list_rule_files() -> list[str]:
    """All rule files: original first, then versions oldest → newest."""
    versions = sorted(
        os.path.basename(p) for p in glob.glob(_VERSION_GLOB)
    )
    return [_RULES_BASE_NAME] + versions


def _get_active_rules_name() -> str:
    try:
        with open(_ACTIVE_PTR_PATH, encoding="utf-8") as f:
            name = json.load(f).get("active_file")
        if name and os.path.exists(os.path.join(_BASE_DIR, name)):
            return name
    except Exception:
        pass
    return _RULES_BASE_NAME


def _set_active_rules_name(filename: str):
    with open(_ACTIVE_PTR_PATH, "w", encoding="utf-8") as f:
        json.dump({"active_file": filename}, f, indent=2)


def _load_rules(filename: str | None = None) -> dict:
    name = filename or _get_active_rules_name()
    with open(os.path.join(_BASE_DIR, name), encoding="utf-8") as f:
        return json.load(f)


def _save_rules_as_version(rules: dict) -> str:
    """Write rules to a NEW version file, make it active. Returns filename."""
    import datetime
    stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    fname = f"attrition_rules_v{stamp}.json"
    payload = dict(rules)
    meta = dict(payload.get("_meta", {}))
    meta["saved_from"] = "Attrition Rules Editor"
    meta["saved_at"] = stamp.replace("_", " ")
    meta["base_version"] = _get_active_rules_name()
    payload["_meta"] = meta
    with open(os.path.join(_BASE_DIR, fname), "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)
    _set_active_rules_name(fname)
    return fname


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


def _component_comment(rules: dict, section: str, comp: str) -> str:
    entry = rules.get(section, {}).get(comp, {})
    if isinstance(entry, dict):
        return str(entry.get("_comment", "") or "")
    return ""


# ── Editor window ─────────────────────────────────────────────────────────────

class AttritionEditorWindow(ctk.CTkToplevel):
    def __init__(self, parent, on_save_callback=None):
        super().__init__(parent)
        self.title("⚙  Attrition Rules Editor")
        self.geometry("760x600")
        self.minsize(640, 460)
        self.configure(fg_color=CLR_BG)
        self.grab_set()   # modal

        self._on_save_callback = on_save_callback
        self._rules: dict = _load_rules()
        self._dirty: bool = False          # unsaved changes flag
        self._edit_entry: tk.Entry | None = None  # floating entry widget
        self._all_rows: list[tuple] = []   # unfiltered rows of current tab
        self._key_paths: dict = {}

        self._build_ui()
        self._populate_current_tab()

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Title bar
        hdr = ctk.CTkFrame(self, fg_color=CLR_SURFACE, corner_radius=0, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="⚙  Attrition Rules Editor",
                     font=ctk.CTkFont(size=20, weight="bold"),
                     text_color=CLR_ACCENT).pack(side="left", padx=16, pady=8)
        ctk.CTkLabel(hdr,
                     text="Double-click a value in  Attrition %  to edit",
                     font=ctk.CTkFont(size=12, slant="italic"),
                     text_color=CLR_MUTED).pack(side="right", padx=16)

        # Version selector row
        ver = ctk.CTkFrame(self, fg_color=CLR_BG, corner_radius=0, height=44)
        ver.pack(fill="x")
        ver.pack_propagate(False)
        ctk.CTkLabel(ver, text="Rules version:",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=CLR_TEXT).pack(side="left", padx=(16, 8))
        self._version_files = _list_rule_files()
        self._version_menu = ctk.CTkOptionMenu(
            ver, width=240, height=28,
            values=[_version_display_name(f) for f in self._version_files],
            fg_color=CLR_SURFACE, text_color=CLR_TEXT,
            button_color=CLR_HEADER_BG, button_hover_color="#d4d4d8",
            command=self._on_version_selected,
        )
        self._version_menu.set(_version_display_name(_get_active_rules_name()))
        self._version_menu.pack(side="left", pady=8)
        self._btn_delete_ver = ctk.CTkButton(
            ver, text="🗑  Delete", width=90, height=28,
            fg_color=CLR_SURFACE, hover_color="#fecaca",
            text_color=CLR_WARN, border_width=1, border_color=CLR_WARN,
            state="disabled", command=self._on_delete_version,
        )
        self._btn_delete_ver.pack(side="left", padx=(8, 0), pady=8)
        self._update_delete_btn_state()
        ctk.CTkLabel(ver,
                     text="Save always creates a new version — the original file is never overwritten",
                     font=ctk.CTkFont(size=11, slant="italic"),
                     text_color=CLR_MUTED).pack(side="left", padx=12)

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
                hover_color="#15803d" if key == "smt" else "#e4e4e7",
                text_color="#ffffff" if key == "smt" else CLR_TEXT,
                command=lambda k=key: self._switch_tab(k),
            )
            btn.pack(side="left", padx=2, pady=3)
            setattr(self, f"_tab_btn_{key}", btn)

        # Search row
        search_row = ctk.CTkFrame(self, fg_color=CLR_BG, corner_radius=0, height=40)
        search_row.pack(fill="x")
        search_row.pack_propagate(False)
        ctk.CTkLabel(search_row, text="🔍 Filter:",
                     font=ctk.CTkFont(size=12),
                     text_color=CLR_MUTED).pack(side="right", padx=(8, 4))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._populate_current_tab())
        ctk.CTkEntry(search_row, width=200, height=26,
                     textvariable=self._search_var,
                     placeholder_text="component or package…").pack(side="right", pady=6)

        # Table frame
        self._table_frame = ctk.CTkFrame(self, fg_color=CLR_BG, corner_radius=0)
        self._table_frame.pack(fill="both", expand=True)
        self._build_tree()

        # Bottom bar
        btm = ctk.CTkFrame(self, fg_color=CLR_SURFACE, corner_radius=0, height=64)
        btm.pack(fill="x", side="bottom")
        btm.pack_propagate(False)

        self.lbl_desc = ctk.CTkLabel(btm, text="", justify="left", anchor="w",
                                     font=ctk.CTkFont(size=12, slant="italic"),
                                     text_color=CLR_MUTED, wraplength=440)
        self.lbl_desc.pack(side="left", padx=14, pady=6)

        self._btn_save = ctk.CTkButton(btm, text="💾  Save as New Version", width=190, height=36,
                      fg_color=CLR_ACCENT, hover_color="#15803d", text_color="#ffffff",
                      state="disabled",
                      command=self._on_save)
        self._btn_save.pack(side="right", padx=14, pady=8)

        self.lbl_status = ctk.CTkLabel(btm, text="", font=ctk.CTkFont(size=12),
                                       text_color=CLR_MUTED, anchor="e", wraplength=180)
        self.lbl_status.pack(side="right", padx=4, pady=6)

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

        self.tree.column("component",  width=280, minwidth=160, stretch=True,  anchor="w")
        self.tree.column("package",    width=160, minwidth=100, stretch=False, anchor="center")
        self.tree.column("attrition",  width=140, minwidth=90,  stretch=False, anchor="center")

        # Tag colours
        self.tree.tag_configure("edited",   foreground=CLR_EDIT_HL, font=("Segoe UI", 13, "bold"))
        self.tree.tag_configure("row_odd",  background=CLR_ROW_ODD)
        self.tree.tag_configure("zero",     foreground=CLR_MUTED)
        self.tree.tag_configure("group_start", background="#eef2ff")

        vsb = ttk.Scrollbar(self._table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        # Bind double-click for editing + selection for description
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-1>", self._cancel_edit)
        self.tree.bind("<<TreeviewSelect>>", self._on_row_selected)

    # ── Tab switching / population ────────────────────────────────────────────
    def _switch_tab(self, tab: str):
        self._cancel_edit()
        self._active_tab.set(tab)
        self._tab_btn_smt.configure(
            fg_color=CLR_ACCENT if tab == "smt" else CLR_SURFACE,
            text_color="#ffffff" if tab == "smt" else CLR_TEXT)
        self._tab_btn_cable.configure(
            fg_color=CLR_ACCENT if tab == "cable" else CLR_SURFACE,
            text_color="#ffffff" if tab == "cable" else CLR_TEXT)
        self._populate_current_tab()

    def _current_rows(self) -> list[tuple]:
        tab = self._active_tab.get()
        if tab == "smt":
            return _flatten_smt(self._rules["smt_rules"])
        return _flatten_cable(self._rules["cable_box_rules"])

    def _populate_current_tab(self):
        """Fill tree with rows for current tab (respecting search filter)."""
        if not hasattr(self, "tree"):
            return
        self._cancel_edit()
        self._all_rows = self._current_rows()

        query = ""
        if hasattr(self, "_search_var"):
            query = self._search_var.get().strip().lower()

        self.tree.delete(*self.tree.get_children())
        self._key_paths = {}
        prev_comp = None
        iid = 0
        shown = 0
        for comp, pkg, rate, key_path in self._all_rows:
            if query and query not in comp.lower() and query not in pkg.lower():
                continue
            pct_str = f"{rate*100:.1f}%"
            tags = []
            if comp != prev_comp:            # first row of each component group
                tags.append("group_start")
            prev_comp = comp
            if iid % 2 != 0:
                tags.append("row_odd")
            if rate == 0.0:
                tags.append("zero")
            row_iid = str(iid)
            self.tree.insert(
                "", "end", iid=row_iid,
                values=(comp, pkg, pct_str),
                tags=tuple(tags),
            )
            self._key_paths[row_iid] = key_path
            iid += 1
            shown += 1

        if hasattr(self, "lbl_status") and not self._dirty:
            total = len(self._all_rows)
            extra = f" (filter: {shown}/{total})" if query else ""
            self.lbl_status.configure(
                text=f"  {total} rules{extra}", text_color=CLR_MUTED)

    def _on_row_selected(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        if not vals:
            return
        section = "smt_rules" if self._active_tab.get() == "smt" else "cable_box_rules"
        comment = _component_comment(self._rules, section, str(vals[0]))
        if comment:
            self.lbl_desc.configure(text=f"{vals[0]}: {comment}")
        else:
            self.lbl_desc.configure(text=str(vals[0]))

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

        self._set_dirty(True)
        self.lbl_status.configure(
            text="  Unsaved — Save as\n  New Version to apply.",
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

    def _set_dirty(self, dirty: bool):
        """Toggle unsaved-changes state and the Save button availability."""
        self._dirty = dirty
        self._btn_save.configure(state="normal" if dirty else "disabled")

    def _update_delete_btn_state(self):
        """Delete button only active when a saved version (not Original) is selected."""
        current = self._version_menu.get()
        can_delete = current != _version_display_name(_RULES_BASE_NAME)
        self._btn_delete_ver.configure(state="normal" if can_delete else "disabled")

    # ── Save / version switching ──────────────────────────────────────────────
    def _on_save(self):
        self._cancel_edit()
        try:
            fname = _save_rules_as_version(self._rules)
            self._set_dirty(False)
            # Refresh dropdown and point at the new version
            self._version_files = _list_rule_files()
            self._version_menu.configure(
                values=[_version_display_name(f) for f in self._version_files])
            self._version_menu.set(_version_display_name(fname))
            self.lbl_status.configure(
                text=f"  Saved → {fname[:24]}…\n  Applied on next Process.",
                text_color=CLR_GREEN,
            )
            # Reload engine in-place
            _reload_engine()
            if self._on_save_callback:
                self._on_save_callback()
        except Exception as exc:
            self.lbl_status.configure(
                text=f"  Save failed: {exc}", text_color=CLR_WARN)

    def _on_version_selected(self, display_name: str):
        """Switch the active rules version."""
        # Map display name back to filename
        fname = next(
            (f for f in self._version_files
             if _version_display_name(f) == display_name),
            _RULES_BASE_NAME,
        )
        if fname == _get_active_rules_name():
            self._update_delete_btn_state()
            return
        if self._dirty:
            if not messagebox.askyesno(
                "Unsaved Changes",
                "You have unsaved changes that will be discarded.\nSwitch version anyway?",
                parent=self
            ):
                self._version_menu.set(_version_display_name(_get_active_rules_name()))
                self._update_delete_btn_state()
                return
        try:
            self._rules = _load_rules(fname)
        except Exception as exc:
            messagebox.showerror("Load Failed", str(exc), parent=self)
            return
        self._set_dirty(False)
        _set_active_rules_name(fname)
        _reload_engine()
        self._populate_current_tab()
        self._update_delete_btn_state()
        self.lbl_status.configure(
            text=f"  Switched to {_version_display_name(fname)}",
            text_color=CLR_MUTED,
        )
        if self._on_save_callback:
            self._on_save_callback()

    def _on_delete_version(self):
        """Delete the selected (non-original) version file."""
        current = self._version_menu.get()
        fname = next(
            (f for f in self._version_files
             if _version_display_name(f) == current),
            None,
        )
        if fname is None or fname == _RULES_BASE_NAME:
            return
        if self._dirty:
            if not messagebox.askyesno(
                "Unsaved Changes",
                "You have unsaved changes that will be discarded.\nContinue?",
                parent=self
            ):
                return
        if not messagebox.askyesno(
            "Delete Version",
            f"Delete rule version '{_version_display_name(fname)}'?\n\n"
            + ("The engine will fall back to the Original rules."
               if fname == _get_active_rules_name()
               else "Other versions are not affected."),
            parent=self
        ):
            return
        try:
            os.remove(os.path.join(_BASE_DIR, fname))
        except Exception as exc:
            messagebox.showerror("Delete Failed", str(exc), parent=self)
            return
        # If the deleted version was active, fall back to the original
        if _get_active_rules_name() == fname:
            self._rules = _load_rules(_RULES_BASE_NAME)
            _set_active_rules_name(_RULES_BASE_NAME)
            _reload_engine()
            if self._on_save_callback:
                self._on_save_callback()
        self._set_dirty(False)
        self._version_files = _list_rule_files()
        self._version_menu.configure(
            values=[_version_display_name(f) for f in self._version_files])
        self._version_menu.set(_version_display_name(_get_active_rules_name()))
        self._update_delete_btn_state()
        self._populate_current_tab()
        self.lbl_status.configure(
            text=f"  Deleted {_version_display_name(fname)}",
            text_color=CLR_MUTED,
        )

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
