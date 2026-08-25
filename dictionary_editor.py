"""
dictionary_editor.py
--------------------
Popup window to VIEW and EDIT the component keyword dictionary
(component_dictionary.json).

Tabs (by category):
  🔲 SMT / PCBA   – keyword → component type (SMT group)
  🔌 Cable / Box  – keyword → component type (cable group)
  📦 Misc / Added – keyword → component type (misc group)
  📐 Package      – keyword → SMT package group
  📏 Units        – unit alias → PCS / LENGTH
  📋 Headers      – column header alias → role

Saving ALWAYS creates a new version file (component_dictionary_v*.json).
The original component_dictionary.json is never modified.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk

from file_versions import (
    list_files, display_name, active_filename, load_active, save_new_version,
)

_BASE_DIR_NAME = "component_dictionary.json"

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
CLR_EDIT_HL   = "#d97706"

# ── Tab definitions ───────────────────────────────────────────────────────────
# key → (json_section, json_group_or_None, col1_title, col2_title)
_TABS = {
    "smt":     ("keyword_to_type", "SMT_COMPONENTS",       "Keyword (viết tắt)", "Component Type"),
    "cable":   ("keyword_to_type", "CABLE_BOX_COMPONENTS", "Keyword (viết tắt)", "Component Type"),
    "misc":    ("keyword_to_type", "MISC_ADDED",           "Keyword (viết tắt)", "Component Type"),
    "package": ("package_keywords", None,                  "Keyword",            "Package Group"),
    "unit":    ("unit_aliases", None,                      "Unit (file BOM)",    "Canonical Unit"),
    "headers": ("column_header_aliases", None,             "Header Alias",       "Column Role"),
}

_UNIT_TARGETS = ["PCS", "LENGTH"]
_HEADER_ROLES = ["description_col", "part_type_col", "quantity_col",
                 "unit_col", "partnumber_col", "mpn_col", "internal_pn_col"]


def _canonical_types() -> list[str]:
    """Component types available in the active attrition rules."""
    try:
        rules = load_active("attrition_rules.json")
        types: list[str] = []
        for section in ("smt_rules", "cable_box_rules"):
            for key in rules.get(section, {}):
                if not key.startswith("_") and key not in types:
                    types.append(key)
        return sorted(types)
    except Exception:
        return []


class DictionaryEditorWindow(ctk.CTkToplevel):
    def __init__(self, parent, on_save_callback=None):
        super().__init__(parent)
        self.title("📚  Keyword Dictionary Editor")
        self.geometry("820x640")
        self.minsize(700, 480)
        self.configure(fg_color=CLR_BG)
        self.grab_set()

        self._on_save_callback = on_save_callback
        self._dict: dict = load_active(_BASE_DIR_NAME)
        self._dirty: bool = False
        self._edit_widget: tk.Widget | None = None
        self._types = _canonical_types()

        self._build_ui()
        self._populate_current_tab()

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        hdr = ctk.CTkFrame(self, fg_color=CLR_SURFACE, corner_radius=0, height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="📚  Keyword Dictionary Editor",
                     font=ctk.CTkFont(size=20, weight="bold"),
                     text_color=CLR_ACCENT).pack(side="left", padx=16, pady=8)
        ctk.CTkLabel(hdr,
                     text="Double-click a cell to edit  •  chọn tab theo danh mục",
                     font=ctk.CTkFont(size=12, slant="italic"),
                     text_color=CLR_MUTED).pack(side="right", padx=16)

        # Version selector row
        ver = ctk.CTkFrame(self, fg_color=CLR_BG, corner_radius=0, height=44)
        ver.pack(fill="x")
        ver.pack_propagate(False)
        ctk.CTkLabel(ver, text="Dictionary version:",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=CLR_TEXT).pack(side="left", padx=(16, 8))
        self._version_files = list_files(_BASE_DIR_NAME)
        self._version_menu = ctk.CTkOptionMenu(
            ver, width=240, height=28,
            values=[display_name(f) for f in self._version_files],
            fg_color=CLR_SURFACE, text_color=CLR_TEXT,
            button_color=CLR_HEADER_BG, button_hover_color="#d4d4d8",
            command=self._on_version_selected,
        )
        self._version_menu.set(display_name(active_filename(_BASE_DIR_NAME)))
        self._version_menu.pack(side="left", pady=8)
        self._btn_delete_ver = ctk.CTkButton(
            ver, text="🗑  Delete", width=90, height=28,
            fg_color=CLR_SURFACE, hover_color="#fecaca",
            text_color=CLR_WARN, border_width=1, border_color=CLR_WARN,
            state="disabled", command=self._on_delete_version,
        )
        self._btn_delete_ver.pack(side="left", padx=(8, 0), pady=8)
        ctk.CTkLabel(ver,
                     text="Save tạo bản ghi mới — từ điển gốc không bao giờ bị ghi đè",
                     font=ctk.CTkFont(size=11, slant="italic"),
                     text_color=CLR_MUTED).pack(side="left", padx=12)

        # Tab bar
        tab_bar = ctk.CTkFrame(self, fg_color=CLR_SURFACE, corner_radius=0, height=48)
        tab_bar.pack(fill="x")
        tab_bar.pack_propagate(False)

        self._active_tab = tk.StringVar(value="smt")
        labels = {
            "smt":     "  🔲  SMT / PCBA  ",
            "cable":   "  🔌  Cable / Box  ",
            "misc":    "  📦  Misc / Added  ",
            "package": "  📐  Package  ",
            "unit":    "  📏  Units  ",
            "headers": "  📋  Headers  ",
        }
        for key, label in labels.items():
            btn = ctk.CTkButton(
                tab_bar, text=label, height=32, corner_radius=0,
                fg_color=CLR_ACCENT if key == "smt" else CLR_SURFACE,
                hover_color="#15803d" if key == "smt" else "#e4e4e7",
                text_color="#ffffff" if key == "smt" else CLR_TEXT,
                command=lambda k=key: self._switch_tab(k),
            )
            btn.pack(side="left", padx=2, pady=3)
            setattr(self, f"_tab_btn_{key}", btn)

        # Search + add row
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
                     placeholder_text="từ khóa hoặc giá trị…").pack(side="right", pady=6)

        # Table frame
        self._table_frame = ctk.CTkFrame(self, fg_color=CLR_BG, corner_radius=0)
        self._table_frame.pack(fill="both", expand=True)
        self._build_tree()

        # Bottom bar
        btm = ctk.CTkFrame(self, fg_color=CLR_SURFACE, corner_radius=0, height=60)
        btm.pack(fill="x", side="bottom")
        btm.pack_propagate(False)

        self.lbl_status = ctk.CTkLabel(btm, text="", font=ctk.CTkFont(size=12),
                                       text_color=CLR_MUTED, anchor="w", wraplength=300)
        self.lbl_status.pack(side="left", padx=14, pady=8)

        self._btn_save = ctk.CTkButton(btm, text="💾  Save as New Version",
                                       width=190, height=36,
                                       fg_color=CLR_ACCENT, hover_color="#15803d",
                                       text_color="#ffffff", state="disabled",
                                       command=self._on_save)
        self._btn_save.pack(side="right", padx=14, pady=8)

        self._btn_add = ctk.CTkButton(btm, text="＋  Add Row", width=110, height=36,
                                      fg_color="#e4e4e7", hover_color="#d4d4d8",
                                      text_color="#18181b",
                                      command=self._on_add_row)
        self._btn_add.pack(side="right", padx=4, pady=8)

        self._btn_del_row = ctk.CTkButton(btm, text="🗑  Delete Row", width=110, height=36,
                                          fg_color="#e4e4e7", hover_color="#fecaca",
                                          text_color=CLR_WARN,
                                          state="disabled",
                                          command=self._on_delete_row)
        self._btn_del_row.pack(side="right", padx=4, pady=8)

        self._update_delete_btn_state()

    def _build_tree(self):
        for w in self._table_frame.winfo_children():
            w.destroy()

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Dic.Treeview",
                        background=CLR_ROW_EVEN, foreground=CLR_TEXT,
                        rowheight=36, fieldbackground=CLR_ROW_EVEN,
                        borderwidth=0, font=("Segoe UI", 13))
        style.configure("Dic.Treeview.Heading",
                        background=CLR_HEADER_BG, foreground=CLR_TEXT,
                        relief="flat", font=("Segoe UI", 13, "bold"))
        style.map("Dic.Treeview",
                  background=[("selected", "#bbf7d0")],
                  foreground=[("selected", "#000000")])

        cols = ("keyword", "target")
        self.tree = ttk.Treeview(self._table_frame, columns=cols,
                                 show="headings", style="Dic.Treeview",
                                 selectmode="browse")
        self.tree.heading("keyword", text="Keyword", anchor="w")
        self.tree.heading("target",  text="Maps To", anchor="w")
        self.tree.column("keyword", width=300, minwidth=180, stretch=True, anchor="w")
        self.tree.column("target",  width=300, minwidth=180, stretch=True, anchor="w")

        self.tree.tag_configure("edited", foreground=CLR_EDIT_HL,
                                font=("Segoe UI", 13, "bold"))
        self.tree.tag_configure("new", background="#ecfdf5")
        self.tree.tag_configure("row_odd", background=CLR_ROW_ODD)

        vsb = ttk.Scrollbar(self._table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-1>", self._cancel_edit)
        self.tree.bind("<<TreeviewSelect>>", self._on_row_selected)

    # ── Data flatten / update ─────────────────────────────────────────────────
    def _tab_info(self) -> tuple[str, str | None]:
        section, group, _c1, _c2 = _TABS[self._active_tab.get()]
        return section, group

    def _current_rows(self) -> list[tuple[str, str, tuple]]:
        """[(keyword, target, locator)] — locator identifies the entry in JSON."""
        section, group = self._tab_info()
        rows = []
        if section == "keyword_to_type":
            for kw, target in self._dict[section].get(group, {}).items():
                if kw.startswith("_"):
                    continue
                rows.append((kw, str(target), (section, group, kw)))
        elif section == "column_header_aliases":
            for role, aliases in self._dict[section].items():
                if role.startswith("_"):
                    continue
                for alias in aliases:
                    rows.append((str(alias), role, (section, role, alias)))
        else:
            for kw, target in self._dict[section].items():
                if kw.startswith("_"):
                    continue
                rows.append((kw, str(target), (section, None, kw)))
        return rows

    def _populate_current_tab(self):
        if not hasattr(self, "tree"):
            return
        self._cancel_edit()
        section, group = self._tab_info()
        col1 = _TABS[self._active_tab.get()][2]
        col2 = _TABS[self._active_tab.get()][3]
        self.tree.heading("keyword", text=col1, anchor="w")
        self.tree.heading("target",  text=col2, anchor="w")

        query = self._search_var.get().strip().lower() if hasattr(self, "_search_var") else ""

        self.tree.delete(*self.tree.get_children())
        self._locators: dict[str, tuple] = {}
        iid = 0
        for kw, target, locator in self._current_rows():
            if query and query not in kw.lower() and query not in target.lower():
                continue
            tags = []
            if iid % 2 != 0:
                tags.append("row_odd")
            row_iid = str(iid)
            self.tree.insert("", "end", iid=row_iid,
                             values=(kw, target), tags=tuple(tags))
            self._locators[row_iid] = locator
            iid += 1

        if hasattr(self, "lbl_status") and not self._dirty:
            self.lbl_status.configure(text=f"  {iid} entries", text_color=CLR_MUTED)

    # ── Tab switching ─────────────────────────────────────────────────────────
    def _switch_tab(self, tab: str):
        self._cancel_edit()
        self._active_tab.set(tab)
        for k in _TABS:
            btn = getattr(self, f"_tab_btn_{k}")
            btn.configure(fg_color=CLR_ACCENT if k == tab else CLR_SURFACE,
                          text_color="#ffffff" if k == tab else CLR_TEXT,
                          hover_color="#15803d" if k == tab else "#e4e4e7")
        self._populate_current_tab()

    # ── Selection ─────────────────────────────────────────────────────────────
    def _on_row_selected(self, event=None):
        has_sel = bool(self.tree.selection())
        self._btn_del_row.configure(state="normal" if has_sel else "disabled")

    # ── Inline editing ────────────────────────────────────────────────────────
    def _on_double_click(self, event: tk.Event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        col = self.tree.identify_column(event.x)
        if col not in ("#1", "#2"):
            return
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self._cancel_edit()
        self._start_edit(row_id, col)

    def _target_choices(self) -> list[str] | None:
        """Fixed choice list for column #2, or None for free text."""
        tab = self._active_tab.get()
        if tab in ("smt", "cable", "misc"):
            return self._types
        if tab == "unit":
            return _UNIT_TARGETS
        if tab == "headers":
            return _HEADER_ROLES
        return None   # package: free text

    def _start_edit(self, row_id: str, col: str):
        bbox = self.tree.bbox(row_id, col)
        if not bbox:
            return
        x, y, w, h = bbox
        current_val = self.tree.item(row_id, "values")[int(col[1]) - 1]

        choices = self._target_choices() if col == "#2" else None
        if choices:
            widget = ttk.Combobox(self.tree, values=choices, state="readonly",
                                  font=("Segoe UI", 12), justify="left")
            widget.set(current_val)
            widget.place(x=x, y=y, width=w, height=h)
            widget.bind("<<ComboboxSelected>>", lambda e: self._commit_edit())
        else:
            var = tk.StringVar(value=current_val)
            widget = tk.Entry(self.tree, textvariable=var,
                              font=("Segoe UI", 12, "bold"),
                              bg="#ffffff", fg=CLR_EDIT_HL,
                              insertbackground=CLR_TEXT, relief="flat", bd=2,
                              highlightthickness=2, highlightcolor=CLR_ACCENT,
                              highlightbackground=CLR_ACCENT, justify="left")
            widget.place(x=x, y=y, width=w, height=h)
            widget.focus_set()
            widget.select_range(0, "end")
            widget.bind("<Return>", self._commit_edit)
            widget.bind("<KP_Enter>", self._commit_edit)
            widget.bind("<Escape>", self._cancel_edit)
            widget.bind("<FocusOut>", self._commit_edit)

        self._edit_widget = widget
        self._edit_row_id = row_id
        self._edit_col = col

    def _commit_edit(self, event=None):
        if self._edit_widget is None:
            return
        row_id = self._edit_row_id
        col = self._edit_col
        new_val = str(self._edit_widget.get()).strip()
        old_kw, old_target = self.tree.item(row_id, "values")
        self._destroy_edit()

        if col == "#1":
            if not new_val or new_val == old_kw:
                return
            if self._key_exists(new_val):
                self.lbl_status.configure(
                    text=f"  ⚠  '{new_val}' đã tồn tại trong tab này", text_color=CLR_WARN)
                return
            self._update_key(str(old_kw), new_val)
            self.tree.set(row_id, "keyword", new_val)
        else:
            if new_val == old_target:
                return
            self._update_target(str(old_kw), str(old_target), new_val)
            self.tree.set(row_id, "target", new_val)

        # highlight edited row
        tags = list(self.tree.item(row_id, "tags"))
        if "edited" not in tags:
            tags.append("edited")
        self.tree.item(row_id, tags=tuple(tags))

        self._set_dirty(True)
        self.lbl_status.configure(
            text="  Chưa lưu — Save as New Version để áp dụng.",
            text_color=CLR_EDIT_HL)

    def _key_exists(self, kw: str) -> bool:
        section, group = self._tab_info()
        if section == "keyword_to_type":
            return kw in self._dict[section].get(group, {})
        if section == "column_header_aliases":
            return any(kw in aliases for aliases in self._dict[section].values())
        return kw in self._dict[section]

    def _update_key(self, old_kw: str, new_kw: str):
        section, group = self._tab_info()
        if section == "keyword_to_type":
            g = self._dict[section][group]
            g[new_kw] = g.pop(old_kw)
        elif section == "column_header_aliases":
            for role, aliases in self._dict[section].items():
                if role.startswith("_"):
                    continue
                if old_kw in aliases:
                    aliases[aliases.index(old_kw)] = new_kw
                    return
        else:
            self._dict[section][new_kw] = self._dict[section].pop(old_kw)

    def _update_target(self, kw: str, old_target: str, new_target: str):
        section, group = self._tab_info()
        if section == "keyword_to_type":
            self._dict[section][group][kw] = new_target
        elif section == "column_header_aliases":
            # move alias to another role list
            for role, aliases in self._dict[section].items():
                if role.startswith("_"):
                    continue
                if kw in aliases:
                    aliases.remove(kw)
                    break
            self._dict[section].setdefault(new_target, []).append(kw)
        else:
            self._dict[section][kw] = new_target

    def _cancel_edit(self, event=None):
        self._destroy_edit()

    def _destroy_edit(self):
        if self._edit_widget:
            self._edit_widget.destroy()
            self._edit_widget = None

    # ── Add / delete rows ─────────────────────────────────────────────────────
    def _on_add_row(self):
        section, group = self._tab_info()
        base = "NEW_KEYWORD"
        kw, i = base, 1
        while self._key_exists(kw):
            i += 1
            kw = f"{base}{i}"

        if section == "keyword_to_type":
            default_target = "OTHER_SPECIAL"
            self._dict[section][group][kw] = default_target
        elif section == "column_header_aliases":
            default_target = "description_col"
            self._dict[section].setdefault(default_target, []).append(kw)
        else:
            default_target = "PCS" if section == "unit_aliases" else ""
            self._dict[section][kw] = default_target

        self._populate_current_tab()
        self._set_dirty(True)
        # select & scroll to the new row
        for iid, locator in self._locators.items():
            if locator[2] == kw:
                self.tree.see(iid)
                self.tree.selection_set(iid)
                break
        self.lbl_status.configure(
            text=f"  Đã thêm '{kw}' — double-click để sửa.", text_color=CLR_MUTED)

    def _on_delete_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        row_id = sel[0]
        kw, target = self.tree.item(row_id, "values")
        if not messagebox.askyesno("Delete Entry",
                                   f"Xóa '{kw}' → '{target}'?", parent=self):
            return
        self._cancel_edit()
        section, group = self._tab_info()
        locator = self._locators[row_id]
        if section == "keyword_to_type":
            self._dict[section][group].pop(str(kw), None)
        elif section == "column_header_aliases":
            aliases = self._dict[section].get(str(target), [])
            if str(kw) in aliases:
                aliases.remove(str(kw))
        else:
            self._dict[section].pop(str(kw), None)

        self._populate_current_tab()
        self._set_dirty(True)
        self.lbl_status.configure(text=f"  Đã xóa '{kw}'.", text_color=CLR_MUTED)

    # ── Dirty / save / versions ───────────────────────────────────────────────
    def _set_dirty(self, dirty: bool):
        self._dirty = dirty
        self._btn_save.configure(state="normal" if dirty else "disabled")

    def _on_save(self):
        self._cancel_edit()
        try:
            fname = save_new_version(_BASE_DIR_NAME, self._dict,
                                     saved_from="Keyword Dictionary Editor")
            self._set_dirty(False)
            self._version_files = list_files(_BASE_DIR_NAME)
            self._version_menu.configure(
                values=[display_name(f) for f in self._version_files])
            self._version_menu.set(display_name(fname))
            self._update_delete_btn_state()
            self.lbl_status.configure(
                text=f"  Đã lưu → {fname}\n  Áp dụng ở lần Process sau.",
                text_color=CLR_GREEN)
            self._reload_consumers()
            if self._on_save_callback:
                self._on_save_callback()
        except Exception as exc:
            self.lbl_status.configure(text=f"  Save failed: {exc}", text_color=CLR_WARN)

    def _on_version_selected(self, display: str):
        fname = next((f for f in self._version_files
                      if display_name(f) == display), _BASE_DIR_NAME)
        if fname == active_filename(_BASE_DIR_NAME):
            self._update_delete_btn_state()
            return
        if self._dirty:
            if not messagebox.askyesno(
                "Unsaved Changes",
                "Có thay đổi chưa lưu sẽ bị bỏ qua.\nChuyển phiên bản anyway?",
                parent=self
            ):
                self._version_menu.set(display_name(active_filename(_BASE_DIR_NAME)))
                self._update_delete_btn_state()
                return
        try:
            self._dict = load_active(_BASE_DIR_NAME, fname)
        except Exception as exc:
            messagebox.showerror("Load Failed", str(exc), parent=self)
            return
        from file_versions import set_active_filename
        set_active_filename(_BASE_DIR_NAME, fname)
        self._set_dirty(False)
        self._reload_consumers()
        self._populate_current_tab()
        self._update_delete_btn_state()
        self.lbl_status.configure(
            text=f"  Đã chuyển sang {display_name(fname)}", text_color=CLR_MUTED)
        if self._on_save_callback:
            self._on_save_callback()

    def _update_delete_btn_state(self):
        can_delete = self._version_menu.get() != display_name(_BASE_DIR_NAME)
        self._btn_delete_ver.configure(state="normal" if can_delete else "disabled")

    def _on_delete_version(self):
        current = self._version_menu.get()
        fname = next((f for f in self._version_files
                      if display_name(f) == current), None)
        if fname is None or fname == _BASE_DIR_NAME:
            return
        if self._dirty and not messagebox.askyesno(
                "Unsaved Changes",
                "Có thay đổi chưa lưu sẽ bị bỏ qua.\nContinue?", parent=self):
            return
        if not messagebox.askyesno(
            "Delete Version",
            f"Xóa phiên bản từ điển '{display_name(fname)}'?\n\n"
            + ("Engine sẽ quay về dùng bản Original." if fname == active_filename(_BASE_DIR_NAME)
               else "Các phiên bản khác không bị ảnh hưởng."),
            parent=self
        ):
            return
        import os
        try:
            os.remove(os.path.join(os.path.dirname(os.path.abspath(__file__)), fname))
        except Exception as exc:
            messagebox.showerror("Delete Failed", str(exc), parent=self)
            return
        from file_versions import set_active_filename
        if active_filename(_BASE_DIR_NAME) == fname:
            self._dict = load_active(_BASE_DIR_NAME, _BASE_DIR_NAME)
            set_active_filename(_BASE_DIR_NAME, _BASE_DIR_NAME)
            self._reload_consumers()
            if self._on_save_callback:
                self._on_save_callback()
        self._set_dirty(False)
        self._version_files = list_files(_BASE_DIR_NAME)
        self._version_menu.configure(
            values=[display_name(f) for f in self._version_files])
        self._version_menu.set(display_name(active_filename(_BASE_DIR_NAME)))
        self._update_delete_btn_state()
        self._populate_current_tab()
        self.lbl_status.configure(
            text=f"  Đã xóa {display_name(fname)}", text_color=CLR_MUTED)

    # ── Hot reload ────────────────────────────────────────────────────────────
    def _reload_consumers(self):
        """Reload modules that read the dictionary at import time."""
        import importlib
        import attrition_engine as _eng
        import bom_reader as _br
        importlib.reload(_eng)
        importlib.reload(_br)

    # ── Close guard ───────────────────────────────────────────────────────────
    def destroy(self):
        if self._dirty:
            if not messagebox.askyesno(
                "Unsaved Changes",
                "Bạn có thay đổi chưa lưu.\nĐóng mà không lưu?",
                parent=self
            ):
                return
        super().destroy()
