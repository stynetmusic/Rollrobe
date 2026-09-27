"""Тёмное рабочее окно Rollrobe: смета, прайс, настройки."""

from __future__ import annotations

import sys
import traceback
import tkinter as tk
import tkinter.font as tkfont
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, ttk

import customtkinter as ctk

from rollrobe.db import UNIT_CODES, Database
from rollrobe.dialogs import ask, pick_estimate, tree_text
from rollrobe.excel_io import (
    SEED_FILES,
    ImportFormatError,
    ensure_seed_prices,
    export_estimate,
    export_price,
    import_workbook,
)
from rollrobe.money import format_edit, format_number, line_total, parse_decimal, q2, sum_lines
from rollrobe.paths import app_dir, database_path
from rollrobe.search import filter_items

TREE_BG = "#232323"
TREE_ALT = "#2A2A2A"
TREE_FG = "#EDEDED"
TREE_HEAD = "#333333"
TREE_SEL = "#3E5F86"
ZERO_FG = "#8E8E8E"
EDITED_FG = "#E2C27A"
BTN = "#3A3D42"
BTN_HOVER = "#4A4E54"
BTN_OK = "#2F6F4E"
BTN_OK_HOVER = "#3B8660"
ESTIMATE_LIMIT = 400
PRICE_LIMIT = 20000


class CellEditor:
    """Поле поверх ячейки таблицы. Enter пишет, Escape отменяет."""

    def __init__(self, tree: ttk.Treeview, font: tuple, on_commit):
        self.tree = tree
        self.font = font
        self.on_commit = on_commit
        self.entry: tk.Entry | None = None
        self.iid = ""
        self.column = ""
        self._closing = False

    def is_open(self) -> bool:
        return self.entry is not None

    def open(self, iid: str, column: str, initial: str) -> None:
        if self.entry is not None:
            self.close(True)
        bbox = self.tree.bbox(iid, column)
        if not bbox:
            self.tree.see(iid)
            bbox = self.tree.bbox(iid, column)
        if not bbox:
            return
        x, y, width, height = bbox
        entry = tk.Entry(
            self.tree,
            font=self.font,
            bg="#101010",
            fg="#FFFFFF",
            insertbackground="#FFFFFF",
            relief="flat",
            justify="right",
            highlightthickness=1,
            highlightbackground="#F3C15A",
        )
        entry.insert(0, initial)
        entry.select_range(0, "end")
        entry.place(x=x, y=y, width=max(width, 70), height=height)
        entry.focus_set()
        entry.bind("<Return>", self._commit_event)
        entry.bind("<KP_Enter>", self._commit_event)
        entry.bind("<Escape>", self._cancel_event)
        entry.bind("<FocusOut>", lambda _event: self.close(True))
        self.entry = entry
        self.iid = iid
        self.column = column

    def close(self, commit: bool) -> None:
        if self._closing or self.entry is None:
            return
        self._closing = True
        try:
            value = self.entry.get()
            iid, column = self.iid, self.column
            widget = self.entry
            self.entry = None
            widget.destroy()
            if commit:
                self.on_commit(iid, column, value)
        finally:
            self._closing = False

    def _commit_event(self, _event=None):
        self.close(True)
        return "break"

    def _cancel_event(self, _event=None):
        self.close(False)
        return "break"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self._booting = True
        self._suspend = False
        self._est_after = None
        self._price_after = None
        self.estimate_id: int | None = None
        self.lines: list[dict] = []
        self.catalog: list[dict] = []
        self._by_id: dict[int, dict] = {}
        self._estimate_hits: dict[str, dict] = {}
        self._lists_by_name: dict[str, dict] = {}
        self.active_list: dict | None = None
        self.dirty = False
        self._clean = None
        self.db = Database(database_path())
        self.unit_labels = self.db.unit_labels()
        self.active_choice = tk.IntVar(value=0)
        self.estimate_show_zero = tk.IntVar(value=0)
        self.price_show_zero = tk.IntVar(value=0)

        family = _font_family()
        self.font_family = family
        self.table_font = (family, 13)
        self.font = ctk.CTkFont(family=family, size=13)
        self.font_title = ctk.CTkFont(family=family, size=20, weight="bold")
        self.font_total = ctk.CTkFont(family=family, size=28, weight="bold")
        self.font_small = ctk.CTkFont(family=family, size=12)
        self.font_hint = ctk.CTkFont(family=family, size=11)

        self.title("Rollrobe")
        self.minsize(1040, 680)
        self.geometry("1240x800")
        self._center(1240, 800)
        self._configure_table_style()
        self._build()
        self._bind_shortcuts()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.report_callback_exception = self._report_error
        self._startup()

    def destroy(self):
        for handle in (self._est_after, self._price_after):
            if handle is not None:
                try:
                    self.after_cancel(handle)
                except tk.TclError:
                    pass
        try:
            self.line_editor.close(False)
            self.price_editor.close(False)
        except Exception:
            pass
        try:
            self.db.close()
        finally:
            super().destroy()

    def _center(self, width: int, height: int) -> None:
        self.update_idletasks()
        x = max(0, (self.winfo_screenwidth() - width) // 2)
        y = max(0, (self.winfo_screenheight() - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _configure_table_style(self) -> None:
        style = ttk.Style(self)
        # aqua на Mac не красит Treeview. clam слушается цветов, окно остаётся тёмным.
        if sys.platform == "darwin":
            try:
                style.theme_use("clam")
            except tk.TclError:
                pass
        style.configure(
            "Roll.Treeview",
            background=TREE_BG,
            fieldbackground=TREE_BG,
            foreground=TREE_FG,
            borderwidth=0,
            relief="flat",
            rowheight=26,
            font=self.table_font,
        )
        style.configure(
            "Roll.Treeview.Heading",
            background=TREE_HEAD,
            foreground="#E4E4E4",
            relief="flat",
            borderwidth=0,
            padding=(6, 4),
            font=(self.font_family, 12, "bold"),
        )
        style.map(
            "Roll.Treeview",
            background=[("selected", TREE_SEL)],
            foreground=[("selected", "#FFFFFF")],
        )
        # macOS aqua прячет цвета Treeview, пока в карте есть фильтр !selected.
        style.map(
            "Roll.Treeview",
            foreground=_drop_mac_filter(style, "Roll.Treeview", "foreground"),
            background=_drop_mac_filter(style, "Roll.Treeview", "background"),
        )
        style.map("Roll.Treeview.Heading", background=[("active", "#3C3C3C")])
        style.configure(
            "Vertical.TScrollbar",
            background="#3A3A3A",
            troughcolor="#1E1E1E",
            bordercolor="#1E1E1E",
            arrowcolor="#D0D0D0",
            relief="flat",
        )
        style.map("Vertical.TScrollbar", background=[("active", "#4A4A4A"), ("!active", "#3A3A3A")])

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        ctk.CTkLabel(bar, text="Rollrobe", font=self.font_title).pack(side="left")
        self.tabs = ctk.CTkSegmentedButton(
            bar,
            values=["Смета", "Прайс", "Настройки"],
            command=self._on_tab,
            font=self.font,
            height=32,
            selected_color=BTN_OK,
            selected_hover_color=BTN_OK_HOVER,
            unselected_color="#2E2E2E",
            unselected_hover_color="#3A3A3A",
            text_color="#F2F2F2",
        )
        self.tabs.pack(side="right")

        self.content = ctk.CTkFrame(self, fg_color="transparent")
        self.content.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 4))
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)

        self.screens = {
            "Смета": self._build_estimate(),
            "Прайс": self._build_price(),
            "Настройки": self._build_settings(),
        }
        self.tabs.set("Смета")
        self._show("Смета")

        status = ctk.CTkFrame(self, fg_color="transparent")
        status.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 8))
        self.status_var = tk.StringVar(value="")
        ctk.CTkLabel(
            status, textvariable=self.status_var, font=self.font_hint, text_color="#A0A0A0", anchor="w"
        ).pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(
            status,
            text="Ctrl/⌘  F поиск · S сохранить · N новая · Enter в смету · Delete удалить строку",
            font=self.font_hint,
            text_color="#7A7A7A",
        ).pack(side="right")

    def _build_estimate(self) -> ctk.CTkFrame:
        screen = ctk.CTkFrame(self.content, fg_color="transparent")
        screen.grid_columnconfigure(0, weight=1)
        screen.grid_rowconfigure(3, weight=2)
        screen.grid_rowconfigure(5, weight=3)

        self.var_number = tk.StringVar()
        self.var_date = tk.StringVar()
        self.var_client = tk.StringVar()
        self.var_object = tk.StringVar()
        self.var_width = tk.StringVar()
        self.var_height = tk.StringVar()
        self.var_comment = tk.StringVar()
        for var in (
            self.var_number, self.var_date, self.var_client, self.var_object,
            self.var_width, self.var_height, self.var_comment,
        ):
            var.trace_add("write", self._on_header_write)

        form = ctk.CTkFrame(screen, fg_color="transparent")
        form.grid(row=0, column=0, sticky="ew")
        for column, weight in ((2, 1), (3, 1)):
            form.grid_columnconfigure(column, weight=weight)
        self._field(form, 0, 0, "Номер", self.var_number, 120)
        self._field(form, 0, 1, "Дата", self.var_date, 110)
        self._field(form, 0, 2, "Клиент", self.var_client, 220)
        self._field(form, 0, 3, "Объект", self.var_object, 220)
        self._field(form, 2, 0, "Ширина, мм", self.var_width, 110)
        self._field(form, 2, 1, "Высота, мм", self.var_height, 110)
        self._field(form, 2, 2, "Комментарий", self.var_comment, 280, columnspan=2)
        ctk.CTkLabel(
            form,
            text="Ширина и высота в итог не входят",
            font=self.font_hint,
            text_color="#8A8A8A",
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 4))

        search = ctk.CTkFrame(screen, fg_color="transparent")
        search.grid(row=1, column=0, sticky="ew", pady=(2, 0))
        search.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(search, text="Поиск", font=self.font_small).grid(row=0, column=0, padx=(0, 6))
        self.estimate_query = tk.StringVar()
        self.estimate_query_entry = ctk.CTkEntry(
            search,
            textvariable=self.estimate_query,
            font=self.font,
            height=32,
            placeholder_text="Код, имя или цвет",
        )
        self.estimate_query_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.estimate_query.trace_add("write", self._schedule_estimate_search)
        ctk.CTkSwitch(
            search,
            text="Показывать без цены",
            variable=self.estimate_show_zero,
            command=self.run_estimate_search,
            font=self.font_small,
            width=210,
            progress_color=BTN_OK,
        ).grid(row=0, column=2, padx=(0, 10))
        ctk.CTkLabel(search, text="Кол-во", font=self.font_small).grid(row=0, column=3, padx=(0, 6))
        self.qty_var = tk.StringVar(value="1,00")
        self.qty_entry = ctk.CTkEntry(search, textvariable=self.qty_var, font=self.font, width=80, height=32)
        self.qty_entry.grid(row=0, column=4, padx=(0, 8))
        self._button(search, "В смету", self._add_from_search, width=110, primary=True).grid(row=0, column=5)

        self.search_caption = tk.StringVar(value="")
        ctk.CTkLabel(
            screen, textvariable=self.search_caption, font=self.font_hint, text_color="#A0A0A0", anchor="w"
        ).grid(row=2, column=0, sticky="ew", pady=(4, 2))

        search_wrap, self.search_tree = self._tree(
            screen,
            ("code", "name", "color", "unit", "price"),
            (("code", "Код", 120, "w", False),
             ("name", "Имя", 360, "w", True),
             ("color", "Цвет", 70, "center", False),
             ("unit", "Ед.", 80, "w", False),
             ("price", "Цена", 110, "e", False)),
        )
        search_wrap.grid(row=3, column=0, sticky="nsew", pady=(0, 6))
        self.search_tree.bind("<Double-1>", self._add_from_search)
        self.search_tree.bind("<Return>", self._add_from_search)
        self.estimate_query_entry.bind("<Return>", self._add_from_search)
        self.estimate_query_entry.bind("<KP_Enter>", self._add_from_search)
        self.qty_entry.bind("<Return>", self._add_from_search)
        self.qty_entry.bind("<KP_Enter>", self._add_from_search)

        self.lines_caption = tk.StringVar(value="Позиции сметы: пока пусто")
        ctk.CTkLabel(
            screen, textvariable=self.lines_caption, font=self.font_small, anchor="w"
        ).grid(row=4, column=0, sticky="ew", pady=(0, 2))

        lines_wrap, self.lines_tree = self._tree(
            screen,
            ("code", "name", "color", "unit", "qty", "price", "sum"),
            (("code", "Код", 110, "w", False),
             ("name", "Имя", 300, "w", True),
             ("color", "Цвет", 70, "center", False),
             ("unit", "Ед.", 80, "w", False),
             ("qty", "Кол-во", 90, "e", False),
             ("price", "Цена", 110, "e", False),
             ("sum", "Сумма", 120, "e", False)),
        )
        lines_wrap.grid(row=5, column=0, sticky="nsew")
        self.lines_tree.bind("<Double-1>", self._on_line_double_click)
        self.lines_tree.bind("<Return>", self._edit_selected_qty)
        self.lines_tree.bind("<Delete>", self._on_delete_key)
        self.lines_tree.bind("<BackSpace>", self._on_delete_key)
        self.line_editor = CellEditor(self.lines_tree, self.table_font, self._commit_line)

        bottom = ctk.CTkFrame(screen, fg_color="transparent")
        bottom.grid(row=6, column=0, sticky="ew", pady=(8, 0))
        actions = ctk.CTkFrame(bottom, fg_color="transparent")
        actions.pack(side="left")
        self._button(actions, "Новая смета", self.new_estimate, 120).pack(side="left", padx=(0, 6))
        self._button(actions, "Открыть", self.open_estimate_dialog, 100).pack(side="left", padx=(0, 6))
        self._button(actions, "Сохранить", self.save_estimate, 110, primary=True).pack(side="left", padx=(0, 6))
        self._button(actions, "Дублировать", self.duplicate_estimate, 120).pack(side="left", padx=(0, 6))
        self._button(actions, "Экспорт в xlsx", self.export_estimate_xlsx, 140).pack(side="left", padx=(0, 6))
        self._button(actions, "Удалить строку", self.delete_selected_line, 130).pack(side="left")

        total_box = ctk.CTkFrame(bottom, fg_color="#322C1C", corner_radius=6)
        total_box.pack(side="right")
        ctk.CTkLabel(total_box, text="Итого", font=self.font_small, text_color="#C8B88A").pack(
            side="left", padx=(12, 8), pady=8
        )
        self.total_label = ctk.CTkLabel(
            total_box, text="0,00", font=self.font_total, text_color="#F3C15A", anchor="e"
        )
        self.total_label.pack(side="left", padx=(0, 14), pady=4)
        return screen

    def _build_price(self) -> ctk.CTkFrame:
        screen = ctk.CTkFrame(self.content, fg_color="transparent")
        screen.grid_columnconfigure(0, weight=1)
        screen.grid_rowconfigure(3, weight=1)

        top = ctk.CTkFrame(screen, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew")
        top.grid_columnconfigure(3, weight=1)
        ctk.CTkLabel(top, text="Прайс", font=self.font_small).grid(row=0, column=0, padx=(0, 6))
        self.price_menu = ctk.CTkOptionMenu(
            top,
            values=["Нет прайса"],
            command=self._on_price_menu,
            font=self.font,
            width=240,
            height=32,
            fg_color="#2E2E2E",
            button_color=BTN_OK,
            button_hover_color=BTN_OK_HOVER,
        )
        self.price_menu.grid(row=0, column=1, padx=(0, 12))
        ctk.CTkLabel(top, text="Поиск", font=self.font_small).grid(row=0, column=2, padx=(0, 6))
        self.price_query = tk.StringVar()
        self.price_query_entry = ctk.CTkEntry(
            top,
            textvariable=self.price_query,
            font=self.font,
            height=32,
            placeholder_text="Код, имя или цвет",
        )
        self.price_query_entry.grid(row=0, column=3, sticky="ew", padx=(0, 8))
        self.price_query.trace_add("write", self._schedule_price_search)
        ctk.CTkSwitch(
            top,
            text="Показывать без цены",
            variable=self.price_show_zero,
            command=self.run_price_search,
            font=self.font_small,
            width=210,
            progress_color=BTN_OK,
        ).grid(row=0, column=4)

        actions = ctk.CTkFrame(screen, fg_color="transparent")
        actions.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        self._button(actions, "Импорт xlsx", self.import_price_dialog, 130).pack(side="left", padx=(0, 6))
        self._button(actions, "Экспорт xlsx", self.export_price_xlsx, 130).pack(side="left", padx=(0, 6))
        self._button(actions, "Сбросить позицию к файлу", self.reset_selected_price, 220).pack(
            side="left", padx=(0, 6)
        )
        self._button(
            actions, "Заменить мои цены ценами из файла", self.replace_all_prices, 280
        ).pack(side="left")

        self.price_caption = tk.StringVar(value="")
        ctk.CTkLabel(
            screen, textvariable=self.price_caption, font=self.font_hint, text_color="#A0A0A0", anchor="w"
        ).grid(row=2, column=0, sticky="ew", pady=(6, 2))

        wrap, self.price_tree = self._tree(
            screen,
            ("code", "name", "color", "unit", "file", "mine", "edited"),
            (("code", "Код", 110, "w", False),
             ("name", "Имя", 300, "w", True),
             ("color", "Цвет", 70, "center", False),
             ("unit", "Ед.", 80, "w", False),
             ("file", "Цена файла", 120, "e", False),
             ("mine", "Моя цена", 120, "e", False),
             ("edited", "Изменена", 90, "center", False)),
        )
        wrap.grid(row=3, column=0, sticky="nsew")
        self.price_tree.bind("<Double-1>", self._on_price_double_click)
        self.price_editor = CellEditor(self.price_tree, self.table_font, self._commit_price)
        ctk.CTkLabel(
            screen,
            text="Двойной щелчок по «Моя цена» — правка, она сразу пишется в базу. Цена файла не меняется. Смета хранит свою цену.",
            font=self.font_hint,
            text_color="#8A8A8A",
            anchor="w",
        ).grid(row=4, column=0, sticky="ew", pady=(6, 0))
        return screen

    def _build_settings(self) -> ctk.CTkFrame:
        screen = ctk.CTkFrame(self.content, fg_color="transparent")
        inner = ctk.CTkFrame(screen, fg_color="transparent")
        inner.pack(anchor="nw", fill="x", padx=8, pady=4)

        ctk.CTkLabel(inner, text="Подписи единиц", font=self.font_title).pack(anchor="w", pady=(0, 4))
        ctk.CTkLabel(
            inner,
            text="В файле ПГИ нет названий единиц, только коды 0, 2, 3 и 4. Здесь меняется подпись, не код.",
            font=self.font_small,
            text_color="#B0B0B0",
            anchor="w",
        ).pack(anchor="w", pady=(0, 8))

        self.unit_vars: dict[int, tk.StringVar] = {}
        for code in UNIT_CODES:
            row = ctk.CTkFrame(inner, fg_color="transparent")
            row.pack(anchor="w", pady=3)
            ctk.CTkLabel(row, text=f"Код {code}", font=self.font, width=70, anchor="w").pack(side="left")
            var = tk.StringVar(value=self.unit_labels[code])
            self.unit_vars[code] = var
            ctk.CTkEntry(row, textvariable=var, font=self.font, width=220, height=32).pack(side="left", padx=(8, 0))

        ctk.CTkLabel(inner, text="Активный прайс", font=self.font_title).pack(anchor="w", pady=(18, 4))
        ctk.CTkLabel(
            inner,
            text="Поиск на смете идёт только по нему. Переключение сразу записывается.",
            font=self.font_small,
            text_color="#B0B0B0",
            anchor="w",
        ).pack(anchor="w", pady=(0, 6))
        self.settings_list_box = ctk.CTkFrame(inner, fg_color="transparent")
        self.settings_list_box.pack(anchor="w", fill="x")

        ctk.CTkLabel(inner, text="База", font=self.font_title).pack(anchor="w", pady=(18, 4))
        ctk.CTkLabel(
            inner,
            text=str(database_path()),
            font=self.font_small,
            text_color="#B0B0B0",
            anchor="w",
        ).pack(anchor="w", pady=(0, 12))
        self._button(inner, "Сохранить подписи", self.save_settings, 180, primary=True).pack(anchor="w")
        return screen

    def _field(self, parent, row, column, label, variable, width, columnspan=1):
        ctk.CTkLabel(parent, text=label, font=self.font_small, anchor="w").grid(
            row=row, column=column, columnspan=columnspan, sticky="w", padx=(0, 8)
        )
        entry = ctk.CTkEntry(parent, textvariable=variable, font=self.font, height=30)
        if columnspan > 1:
            entry.grid(row=row + 1, column=column, columnspan=columnspan, sticky="ew", padx=(0, 10), pady=(0, 6))
        else:
            entry.configure(width=width)
            entry.grid(row=row + 1, column=column, sticky="ew", padx=(0, 10), pady=(0, 6))
        return entry

    def _button(self, parent, text, command, width=128, primary=False):
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=32,
            font=self.font,
            fg_color=BTN_OK if primary else BTN,
            hover_color=BTN_OK_HOVER if primary else BTN_HOVER,
        )

    def _tree(self, parent, columns, spec):
        wrap = ctk.CTkFrame(parent, fg_color=TREE_BG, corner_radius=4)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(0, weight=1)
        tree = ttk.Treeview(
            wrap, columns=columns, show="headings", style="Roll.Treeview", selectmode="browse"
        )
        for key, title, width, anchor, stretch in spec:
            tree.heading(key, text=title)
            tree.column(key, width=width, minwidth=48, anchor=anchor, stretch=stretch)
        scroll = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")
        tree.tag_configure("odd", background=TREE_ALT)
        tree.tag_configure("even", background=TREE_BG)
        tree.tag_configure("zero", foreground=ZERO_FG)
        tree.tag_configure("edited", foreground=EDITED_FG)
        return wrap, tree

    def _bind_shortcuts(self) -> None:
        pairs = (
            (("f", "F", "Cyrillic_a", "Cyrillic_A"), self._shortcut_find),
            (("s", "S", "Cyrillic_yeru", "Cyrillic_YERU"), self._shortcut_save),
            (("n", "N", "Cyrillic_te", "Cyrillic_TE"), self._shortcut_new),
        )
        for keys, handler in pairs:
            for key in keys:
                self.bind_all(f"<Control-{key}>", handler)
                self.bind_all(f"<Command-{key}>", handler)

    def _startup(self) -> None:
        def progress(text: str) -> None:
            self.status(text)
            self.update_idletasks()

        notes: list[str] = []
        try:
            notes = ensure_seed_prices(self.db, on_progress=progress)
        except Exception as exc:
            traceback.print_exc()
            notes = [f"Ошибка импорта: {exc}"]
        self.reload_catalog()
        self.refresh_settings_lists()
        self.sync_price_menu()
        self._load_blank()
        self.run_estimate_search()
        if self.active_list:
            text = f"Активный прайс: {self.active_list['name']}. Позиций: {len(self.catalog)}."
            if notes:
                text = " ".join(notes) + " " + text
            self.status(text)
        else:
            self.status("Прайс не загружен. Импортируйте xlsx ПГИ на вкладке «Прайс».")
        self._booting = False

    def status(self, text: str) -> None:
        self.status_var.set(text)

    def current_total(self) -> Decimal:
        return sum_lines(self.lines)

    def reload_catalog(self) -> None:
        self.active_list = self.db.active_list()
        self.unit_labels = self.db.unit_labels()
        if self.active_list is None:
            self.catalog = []
        else:
            self.catalog = self.db.items(self.active_list["id"])
        self._by_id = {item["id"]: item for item in self.catalog}

    def _on_tab(self, name: str) -> None:
        self._show(name)

    def go_tab(self, name: str) -> None:
        self.tabs.set(name)
        self._show(name)

    def _show(self, name: str) -> None:
        for screen in self.screens.values():
            screen.grid_remove()
        self.screens[name].grid(row=0, column=0, sticky="nsew")
        if name == "Прайс":
            self.sync_price_menu()
            self.run_price_search()

    def _on_header_write(self, *_args) -> None:
        self._refresh_dirty()

    def _refresh_dirty(self) -> None:
        if self._suspend:
            return
        self.dirty = self._capture() != self._clean
        self._update_title()

    def _update_title(self) -> None:
        number = self.var_number.get().strip()
        title = "Rollrobe" if not number else f"Rollrobe — {number}"
        if self.dirty:
            title += " *"
        self.title(title)

    def _header_from_form(self) -> dict:
        return {
            "number": self.var_number.get().strip(),
            "date": self.var_date.get().strip(),
            "client": self.var_client.get().strip(),
            "object": self.var_object.get().strip(),
            "width": self.var_width.get().strip(),
            "height": self.var_height.get().strip(),
            "comment": self.var_comment.get().strip(),
        }

    def _capture(self):
        header = self._header_from_form()
        lines = tuple(
            (
                line["code"],
                line["name"],
                line["color"],
                int(line["unit"]),
                f"{q2(line['qty']):.2f}",
                f"{q2(line['price']):.2f}",
            )
            for line in self.lines
        )
        return (
            header["number"], header["date"], header["client"], header["object"],
            header["width"], header["height"], header["comment"], lines,
        )

    def _load_blank(self) -> None:
        self._suspend = True
        self.estimate_id = None
        self.var_number.set("")
        self.var_date.set(datetime.now().strftime("%d.%m.%Y"))
        self.var_client.set("")
        self.var_object.set("")
        self.var_width.set("")
        self.var_height.set("")
        self.var_comment.set("")
        self.lines = []
        self._rebuild_lines()
        self._suspend = False
        self._clean = self._capture()
        self.dirty = False
        self._update_title()

    def load_estimate_by_id(self, estimate_id: int) -> bool:
        loaded = self.db.load_estimate(estimate_id)
        if loaded is None:
            self.status("Смета не найдена")
            return False
        header, lines = loaded
        self._suspend = True
        self.estimate_id = header["id"]
        self.var_number.set(header["number"])
        self.var_date.set(header["date"])
        self.var_client.set(header["client"])
        self.var_object.set(header["object"])
        self.var_width.set(header["width"])
        self.var_height.set(header["height"])
        self.var_comment.set(header["comment"])
        self.lines = lines
        self._rebuild_lines()
        self._suspend = False
        self._clean = self._capture()
        self.dirty = False
        self._update_title()
        return True

    def _rebuild_lines(self) -> None:
        if hasattr(self, "line_editor"):
            self.line_editor.close(False)
        _clear_tree(self.lines_tree)
        for index, line in enumerate(self.lines):
            self.lines_tree.insert(
                "",
                "end",
                iid=str(index),
                tags=("odd" if index % 2 else "even",),
                values=self._line_values(line),
            )
        count = len(self.lines)
        self.lines_caption.set("Позиции сметы: пока пусто" if count == 0 else f"Позиции сметы: {count}")
        self.total_label.configure(text=format_number(sum_lines(self.lines)))

    def _line_values(self, line: dict):
        return (
            tree_text(line["code"]),
            tree_text(line["name"]),
            tree_text(line["color"]),
            self.unit_labels.get(int(line["unit"]), f"ед. {line['unit']}"),
            format_number(line["qty"]),
            format_number(line["price"]),
            format_number(line_total(line["qty"], line["price"])),
        )

    def _schedule_estimate_search(self, *_args) -> None:
        if self._est_after is not None:
            self.after_cancel(self._est_after)
        self._est_after = self.after(120, self._run_estimate_search_later)

    def _run_estimate_search_later(self) -> None:
        self._est_after = None
        self.run_estimate_search()

    def run_estimate_search(self, *_args) -> None:
        query = self.estimate_query.get().strip()
        self._estimate_hits = {}
        _clear_tree(self.search_tree)
        prefix = f"{self.active_list['name']} · " if self.active_list else ""
        if not query:
            hidden = "Нули скрыты." if not self.estimate_show_zero.get() else "Позиции без цены тоже видны."
            self.search_caption.set(prefix + "Введите код, имя или цвет. " + hidden)
            return
        if not self.catalog:
            self.search_caption.set("Нет активного прайса.")
            return
        result = filter_items(
            self.catalog,
            query,
            show_zero=bool(self.estimate_show_zero.get()),
            limit=ESTIMATE_LIMIT,
        )
        for index, item in enumerate(result["shown"]):
            iid = str(item["id"])
            self._estimate_hits[iid] = item
            tags = ["odd" if index % 2 else "even"]
            if item["user_price"] == 0:
                tags.append("zero")
            self.search_tree.insert("", "end", iid=iid, tags=tuple(tags), values=self._search_values(item))
        self.search_caption.set(prefix + _result_text(result))

    def _search_values(self, item: dict):
        return (
            tree_text(item["code"]),
            tree_text(item["name"]),
            tree_text(item["color"]),
            self.unit_labels.get(item["unit"], f"ед. {item['unit']}"),
            format_number(item["user_price"]),
        )

    def _selected_search_item(self) -> dict | None:
        selected = self.search_tree.selection()
        if selected and selected[0] in self._estimate_hits:
            return self._estimate_hits[selected[0]]
        children = self.search_tree.get_children()
        if children and children[0] in self._estimate_hits:
            return self._estimate_hits[children[0]]
        return None

    def _add_from_search(self, _event=None):
        item = self._selected_search_item()
        if item is None:
            self.status("Ничего не найдено. Введите код, имя или цвет.")
            return "break"
        self.add_item_to_estimate(item)
        return "break"

    def add_item_to_estimate(self, item: dict, qty: Decimal | None = None) -> bool:
        if qty is None:
            try:
                qty = parse_decimal(self.qty_var.get())
            except Exception as exc:
                self.status(f"Количество: {exc}. Пример: 1,50")
                return False
        else:
            qty = q2(qty)
        if qty == 0:
            self.status("Количество должно быть больше нуля")
            return False
        self.lines.append(
            {
                "code": item["code"],
                "name": item["name"],
                "color": item["color"],
                "unit": int(item["unit"]),
                "qty": qty,
                "price": q2(item["user_price"]),
            }
        )
        self._rebuild_lines()
        last = str(len(self.lines) - 1)
        if self.lines_tree.exists(last):
            self.lines_tree.selection_set(last)
            self.lines_tree.see(last)
        self._refresh_dirty()
        self.status(f"В смете: {item['code']} {item['name']} × {format_number(qty)}")
        return True

    def _on_line_double_click(self, event):
        if self.lines_tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.lines_tree.identify_row(event.y)
        logical = {"#5": "qty", "#6": "price"}.get(self.lines_tree.identify_column(event.x))
        if not iid or not logical:
            return
        index = int(iid)
        self.line_editor.open(iid, logical, format_edit(self.lines[index][logical]))
        return "break"

    def _edit_selected_qty(self, _event=None):
        selected = self.lines_tree.selection()
        if not selected:
            return "break"
        index = int(selected[0])
        self.line_editor.open(selected[0], "qty", format_edit(self.lines[index]["qty"]))
        return "break"

    def _commit_line(self, iid: str, column: str, raw: str) -> None:
        index = int(iid)
        if index >= len(self.lines):
            return
        try:
            value = parse_decimal(raw)
        except Exception as exc:
            label = "Количество" if column == "qty" else "Цена"
            self.status(f"{label}: {exc}. Пример: 12,50")
            return
        if column == "qty" and value == 0:
            self.status("Количество должно быть больше нуля")
            return
        self.lines[index][column] = value
        self._rebuild_lines()
        if self.lines_tree.exists(str(index)):
            self.lines_tree.selection_set(str(index))
        self._refresh_dirty()
        self.status("Строка обновлена")

    def _on_delete_key(self, _event=None):
        self.delete_selected_line()
        return "break"

    def delete_selected_line(self) -> None:
        if self.line_editor.is_open():
            self.line_editor.close(False)
        selected = self.lines_tree.selection()
        if not selected:
            self.status("Выберите строку сметы")
            return
        index = int(selected[0])
        if index < 0 or index >= len(self.lines):
            return
        removed = self.lines.pop(index)
        self._rebuild_lines()
        if self.lines:
            next_index = min(index, len(self.lines) - 1)
            self.lines_tree.selection_set(str(next_index))
        self._refresh_dirty()
        self.status(f"Удалено: {removed['code']} {removed['name']}".strip())

    def new_estimate(self) -> None:
        if not self._confirm_discard_if_dirty():
            return
        self._load_blank()
        self.go_tab("Смета")
        self.status("Новая смета")

    def open_estimate_dialog(self) -> None:
        if not self._confirm_discard_if_dirty():
            return
        chosen = pick_estimate(self, self.db.list_estimates(), self.font)
        if chosen is None:
            return
        if self.load_estimate_by_id(chosen):
            self.go_tab("Смета")
            self.status("Смета открыта")

    def save_estimate(self) -> bool:
        if self.line_editor.is_open():
            self.line_editor.close(True)
        try:
            self.estimate_id = self.db.save_estimate(self.estimate_id, self._header_from_form(), self.lines)
        except Exception as exc:
            traceback.print_exc()
            self.status(f"Не удалось сохранить: {exc}")
            ask(self, "Ошибка", str(exc), [("Понятно", "ok")], self.font)
            return False
        self._clean = self._capture()
        self.dirty = False
        self._update_title()
        self.status("Смета сохранена")
        return True

    def duplicate_estimate(self) -> None:
        had_saved = self.estimate_id is not None
        if had_saved and self.dirty:
            if not self.save_estimate():
                return
        header = self._header_from_form()
        base = header["number"].strip()
        header["number"] = f"{base} копия".strip()
        try:
            new_id = self.db.save_estimate(None, header, self.lines)
        except Exception as exc:
            traceback.print_exc()
            self.status(f"Не удалось дублировать: {exc}")
            return
        self.load_estimate_by_id(new_id)
        self.go_tab("Смета")
        if had_saved:
            self.status("Создана копия. Исходная смета не изменена.")
        else:
            self.status("Копия сохранена")

    def export_estimate_xlsx(self) -> None:
        number = self.var_number.get().strip()
        initial = _safe_filename(f"Смета {number}" if number else "Смета") + ".xlsx"
        chosen = filedialog.asksaveasfilename(
            parent=self,
            title="Экспорт сметы",
            initialdir=str(app_dir()),
            initialfile=initial,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not chosen:
            return
        path = _xlsx_path(chosen)
        if not self._confirm_not_price_file(path):
            return
        try:
            export_estimate(path, self._header_from_form(), self.lines, self.unit_labels)
        except Exception as exc:
            traceback.print_exc()
            self.status(f"Не удалось экспортировать: {exc}")
            ask(self, "Ошибка", str(exc), [("Понятно", "ok")], self.font)
            return
        self.status(f"Смета выгружена: {path.name}")

    def _confirm_discard_if_dirty(self) -> bool:
        if not self.dirty:
            return True
        choice = ask(
            self,
            "Смета не сохранена",
            "Сохранить текущую смету?",
            [("Сохранить", "save"), ("Не сохранять", "drop"), ("Отмена", "cancel")],
            self.font,
        )
        if choice == "save":
            return self.save_estimate()
        return choice == "drop"

    def _schedule_price_search(self, *_args) -> None:
        if self._price_after is not None:
            self.after_cancel(self._price_after)
        self._price_after = self.after(120, self._run_price_search_later)

    def _run_price_search_later(self) -> None:
        self._price_after = None
        self.run_price_search()

    def run_price_search(self, keep_id=None) -> None:
        if self.tabs.get() != "Прайс":
            return
        if self.price_editor.is_open():
            # Правка сама вызовет поиск после записи. Здесь её не выбрасываем.
            self.price_editor.close(True)
            return
        _clear_tree(self.price_tree)
        if not self.catalog:
            self.price_caption.set("Нет позиций. Импортируйте прайс ПГИ.")
            return
        result = filter_items(
            self.catalog,
            self.price_query.get(),
            show_zero=bool(self.price_show_zero.get()),
            limit=PRICE_LIMIT,
        )
        for index, item in enumerate(result["shown"]):
            tags = ["odd" if index % 2 else "even"]
            if item["user_price"] == 0:
                tags.append("zero")
            if item["user_edited"]:
                tags.append("edited")
            self.price_tree.insert(
                "",
                "end",
                iid=str(item["id"]),
                tags=tuple(tags),
                values=self._price_values(item),
            )
        prefix = f"{self.active_list['name']} · " if self.active_list else ""
        self.price_caption.set(prefix + _result_text(result))
        if keep_id is not None and self.price_tree.exists(str(keep_id)):
            self.price_tree.selection_set(str(keep_id))
            self.price_tree.see(str(keep_id))

    def _price_values(self, item: dict):
        return (
            tree_text(item["code"]),
            tree_text(item["name"]),
            tree_text(item["color"]),
            self.unit_labels.get(item["unit"], f"ед. {item['unit']}"),
            format_number(item["file_price"]),
            format_number(item["user_price"]),
            "да" if item["user_edited"] else "",
        )

    def _on_price_menu(self, name: str) -> None:
        listed = self._lists_by_name.get(name)
        if listed is None:
            return
        if self.active_list and listed["id"] == self.active_list["id"]:
            return
        self.db.set_active(listed["id"])
        self.reload_catalog()
        self.active_choice.set(listed["id"])
        self.run_price_search()
        self.run_estimate_search()
        self.status(f"Активный прайс: {name}")

    def sync_price_menu(self) -> None:
        lists = self.db.lists()
        self._lists_by_name = {item["name"]: item for item in lists}
        if not lists:
            self.price_menu.configure(values=["Нет прайса"])
            self.price_menu.set("Нет прайса")
            return
        names = [item["name"] for item in lists]
        self.price_menu.configure(values=names)
        active_name = self.active_list["name"] if self.active_list else names[0]
        self.price_menu.set(active_name)

    def _on_price_double_click(self, event):
        region = self.price_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.price_tree.identify_column(event.x)
        iid = self.price_tree.identify_row(event.y)
        if not iid:
            return
        if column == "#5":
            self.status("Цена файла не редактируется. Меняйте «Моя цена» или сбросьте позицию к файлу.")
            return "break"
        if column != "#6":
            return
        item = self._by_id.get(int(iid))
        if item is None:
            return
        self.price_editor.open(iid, "price", format_edit(item["user_price"]))
        return "break"

    def _commit_price(self, iid: str, _column: str, raw: str) -> None:
        item = self._by_id.get(int(iid))
        if item is None:
            return
        try:
            price = parse_decimal(raw)
        except Exception as exc:
            self.status(f"Цена: {exc}. Пример: 2296,26")
            return
        self.db.set_user_price(item["id"], price)
        item["user_price"] = price
        item["user_edited"] = price != item["file_price"]
        self.run_price_search(keep_id=item["id"])
        self.run_estimate_search()
        note = "отличается от файла" if item["user_edited"] else "как в файле"
        self.status(f"Цена сохранена, {note}: {item['code']}")

    def reset_selected_price(self) -> None:
        selected = self.price_tree.selection()
        if not selected:
            self.status("Выберите позицию прайса")
            return
        item = self._by_id.get(int(selected[0]))
        if item is None:
            return
        self.db.reset_user_price(item["id"])
        item["user_price"] = item["file_price"]
        item["user_edited"] = False
        self.run_price_search(keep_id=item["id"])
        self.run_estimate_search()
        self.status(f"Цена {item['code']} сброшена к файлу")

    def replace_all_prices(self) -> None:
        if self.active_list is None:
            self.status("Нет активного прайса")
            return
        name = self.active_list["name"]
        choice = ask(
            self,
            "Заменить мои цены",
            f"Заменить все мои цены в прайсе «{name}» ценами из файла? Ручные правки будут сброшены.",
            [("Заменить", "yes"), ("Отмена", "cancel")],
            self.font,
        )
        if choice != "yes":
            return
        count = self.db.replace_user_prices(self.active_list["id"])
        self.reload_catalog()
        self.run_price_search()
        self.run_estimate_search()
        self.status(f"Мои цены в «{name}» заменены ценами из файла ({count} поз.)")

    def import_price_dialog(self) -> None:
        chosen = filedialog.askopenfilename(
            parent=self,
            title="Импорт прайса ПГИ",
            initialdir=str(app_dir()),
            filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")],
        )
        if not chosen:
            return
        path = Path(chosen)
        self.status(f"Импорт {path.name}…")
        self.update_idletasks()
        try:
            message = import_workbook(self.db, path, make_active=self.db.active_list() is None)
        except ImportFormatError as exc:
            self.status(str(exc))
            ask(self, "Импорт", str(exc), [("Понятно", "ok")], self.font)
            return
        except Exception as exc:
            traceback.print_exc()
            self.status(f"Не удалось импортировать: {exc}")
            ask(self, "Импорт", str(exc), [("Понятно", "ok")], self.font)
            return
        self.reload_catalog()
        self.refresh_settings_lists()
        self.sync_price_menu()
        self.run_price_search()
        self.run_estimate_search()
        self.status(message)

    def export_price_xlsx(self) -> None:
        if self.active_list is None or not self.catalog:
            self.status("Нечего экспортировать")
            return
        initial = _safe_filename(f"Прайс-{self.active_list['name']}") + ".xlsx"
        chosen = filedialog.asksaveasfilename(
            parent=self,
            title="Экспорт прайса",
            initialdir=str(app_dir()),
            initialfile=initial,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not chosen:
            return
        path = _xlsx_path(chosen)
        if not self._confirm_not_price_file(path):
            return
        try:
            export_price(path, self.active_list["name"], self.catalog, self.unit_labels)
        except Exception as exc:
            traceback.print_exc()
            self.status(f"Не удалось экспортировать: {exc}")
            return
        self.status(f"Прайс выгружен: {path.name}")

    def _confirm_not_price_file(self, path: Path) -> bool:
        reserved = {row["filename"] for row in self.db.lists()}
        reserved.update(name for name, _active in SEED_FILES)
        if path.name not in reserved:
            return True
        ask(
            self,
            "Это файл прайса",
            f"«{path.name}» — исходный прайс ПГИ. Экспорт его перезапишет и сломает повторный импорт. Выберите другое имя.",
            [("Отмена", "cancel")],
            self.font,
        )
        self.status("Экспорт отменён: выберите другое имя файла")
        return False

    def refresh_settings_lists(self) -> None:
        for child in self.settings_list_box.winfo_children():
            child.destroy()
        lists = self.db.lists()
        current = self.active_list["id"] if self.active_list else 0
        self.active_choice.set(current)
        if not lists:
            ctk.CTkLabel(
                self.settings_list_box, text="Прайсы ещё не импортированы.", font=self.font
            ).pack(anchor="w")
            return
        for item in lists:
            ctk.CTkRadioButton(
                self.settings_list_box,
                text=item["name"],
                variable=self.active_choice,
                value=item["id"],
                font=self.font,
                command=self._on_settings_radio,
                radiobutton_width=18,
                radiobutton_height=18,
            ).pack(anchor="w", pady=2)

    def _on_settings_radio(self) -> None:
        if self._suspend:
            return
        list_id = int(self.active_choice.get())
        if self.active_list and list_id == self.active_list["id"]:
            return
        self.db.set_active(list_id)
        self.reload_catalog()
        self.sync_price_menu()
        self.run_price_search()
        self.run_estimate_search()
        name = self.active_list["name"] if self.active_list else ""
        self.status(f"Активный прайс: {name}")

    def save_settings(self) -> None:
        labels = {}
        for code, var in self.unit_vars.items():
            text = var.get().strip() or f"ед. {code}"
            var.set(text)
            labels[code] = text
        self.db.set_unit_labels(labels)
        self.unit_labels = labels
        self._rebuild_lines()
        self.run_estimate_search()
        self.run_price_search()
        self.status("Подписи единиц сохранены")

    def _shortcut_find(self, event=None):
        if not self._shortcut_here(event):
            return
        if self.tabs.get() == "Прайс":
            self.price_query_entry.focus_set()
            self.price_query_entry.select_range(0, "end")
        else:
            if self.tabs.get() != "Смета":
                self.go_tab("Смета")
            self.estimate_query_entry.focus_set()
            self.estimate_query_entry.select_range(0, "end")
        return "break"

    def _shortcut_save(self, event=None):
        if not self._shortcut_here(event):
            return
        tab = self.tabs.get()
        if tab == "Настройки":
            self.save_settings()
        elif tab == "Прайс":
            if self.price_editor.is_open():
                self.price_editor.close(True)
            else:
                self.status("Цены прайса записываются сразу, отдельное сохранение не нужно")
        else:
            self.save_estimate()
        return "break"

    def _shortcut_new(self, event=None):
        if not self._shortcut_here(event):
            return
        self.new_estimate()
        return "break"

    def _shortcut_here(self, event) -> bool:
        if event is None:
            return True
        try:
            return event.widget.winfo_toplevel() is self
        except tk.TclError:
            return False

    def _on_close(self) -> None:
        if self._booting:
            self.destroy()
            return
        if not self._confirm_discard_if_dirty():
            return
        self.destroy()

    def _report_error(self, exc_type, exc, tb) -> None:
        traceback.print_exception(exc_type, exc, tb)
        try:
            self.status(str(exc))
        except tk.TclError:
            pass


def _font_family() -> str:
    if "Segoe UI" in set(tkfont.families()):
        return "Segoe UI"
    return tkfont.nametofont("TkDefaultFont").actual()["family"]


def _drop_mac_filter(style: ttk.Style, style_name: str, option: str):
    return [
        item
        for item in style.map(style_name, query_opt=option)
        if item[:2] != ("!disabled", "!selected")
    ]


def _clear_tree(tree: ttk.Treeview) -> None:
    children = tree.get_children()
    if children:
        tree.delete(*children)


def _result_text(result: dict) -> str:
    shown = len(result["shown"])
    text = f"показано {shown}"
    if result["visible_total"] > shown:
        text += f" из {result['visible_total']}, уточните поиск"
    if result["hidden"]:
        text += f" · скрыто без цены: {result['hidden']}"
    return text


def _safe_filename(text: str) -> str:
    cleaned = "".join(ch for ch in text if ch not in '\\/:*?"<>|').strip()
    return cleaned or "Смета"


def _xlsx_path(chosen: str) -> Path:
    path = Path(chosen)
    if path.suffix.lower() != ".xlsx":
        path = path.with_suffix(".xlsx")
    return path


def main() -> None:
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")
    app = App()
    app.mainloop()
