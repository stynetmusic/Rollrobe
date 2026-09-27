"""Короткие тёмные диалоги. Кнопки по-русски, без системных Yes/No."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

import customtkinter as ctk

from rollrobe.money import format_number


def tree_text(value) -> str:
    """Tk превращает «01» и «00» в числа и съедает ноль. Невидимый символ оставляет текст."""
    text = "" if value is None else str(value)
    if text.isascii() and text.isdigit():
        return "\u200b" + text
    return text


def ask(master, title: str, text: str, buttons: list[tuple[str, str]], font) -> str | None:
    dialog = _Ask(master, title, text, buttons, font)
    return dialog.result


def pick_estimate(master, rows: list[dict], font) -> int | None:
    dialog = _Picker(master, rows, font)
    return dialog.result


class _Ask(ctk.CTkToplevel):
    def __init__(self, master, title: str, text: str, buttons: list[tuple[str, str]], font):
        super().__init__(master)
        self.result = None
        self.title(title)
        self.resizable(False, False)
        self.transient(master)
        label = ctk.CTkLabel(self, text=text, font=font, wraplength=440, justify="left", anchor="w")
        label.pack(fill="x", padx=18, pady=(16, 12))
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=(0, 16))
        for caption, value in buttons:
            primary = value not in (None, "cancel", "drop")
            button = ctk.CTkButton(
                row,
                text=caption,
                font=font,
                width=130,
                height=32,
                fg_color="#2F6F4E" if primary else "#3A3D42",
                hover_color="#3B8660" if primary else "#4A4E54",
                command=lambda chosen=value: self._close(chosen),
            )
            button.pack(side="left", padx=4)
        self.protocol("WM_DELETE_WINDOW", lambda: self._close(None))
        self.bind("<Escape>", lambda _event: self._close(None))
        self.grab_set()
        self._place(master, 480, 180)
        self.after(30, self.lift)
        self.after(30, self.focus_force)
        self.wait_window()

    def _close(self, value) -> None:
        self.result = value
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()

    def _place(self, master, width: int, height: int) -> None:
        self.update_idletasks()
        width = max(width, self.winfo_reqwidth())
        height = max(height, self.winfo_reqheight())
        x = master.winfo_rootx() + max(0, (master.winfo_width() - width) // 2)
        y = master.winfo_rooty() + max(0, (master.winfo_height() - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")


class _Picker(ctk.CTkToplevel):
    def __init__(self, master, rows: list[dict], font):
        super().__init__(master)
        self.result = None
        self.title("Открыть смету")
        self.transient(master)
        self.minsize(720, 360)
        self.geometry("860x440")
        if not rows:
            ctk.CTkLabel(self, text="Сохранённых смет нет.", font=font, anchor="w").pack(
                fill="x", padx=16, pady=16
            )
        else:
            wrap = ctk.CTkFrame(self, fg_color="#1E1E1E", corner_radius=4)
            wrap.pack(fill="both", expand=True, padx=12, pady=(12, 8))
            columns = ("number", "date", "client", "object", "positions", "total")
            self.tree = ttk.Treeview(
                wrap, columns=columns, show="headings", style="Roll.Treeview", selectmode="browse"
            )
            headings = (
                ("number", "Номер", 120, "w"),
                ("date", "Дата", 110, "w"),
                ("client", "Клиент", 180, "w"),
                ("object", "Объект", 180, "w"),
                ("positions", "Позиции", 80, "e"),
                ("total", "Итого", 120, "e"),
            )
            for key, title, width, anchor in headings:
                self.tree.heading(key, text=title)
                self.tree.column(key, width=width, anchor=anchor, stretch=(key in ("client", "object")))
            scroll = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview, style="Roll.Vertical.TScrollbar")
            self.tree.configure(yscrollcommand=scroll.set)
            self.tree.grid(row=0, column=0, sticky="nsew")
            scroll.grid(row=0, column=1, sticky="ns")
            wrap.grid_columnconfigure(0, weight=1)
            wrap.grid_rowconfigure(0, weight=1)
            for index, row in enumerate(rows):
                self.tree.insert(
                    "",
                    "end",
                    iid=str(row["id"]),
                    tags=("odd" if index % 2 else "even",),
                    values=(
                        tree_text(row["number"]),
                        row["date"],
                        tree_text(row["client"]),
                        tree_text(row["object"]),
                        row["positions"],
                        format_number(row["total"]),
                    ),
                )
            self.tree.bind("<Double-1>", self._open_selected)
            self.tree.bind("<Return>", self._open_selected)
            children = self.tree.get_children()
            if children:
                self.tree.selection_set(children[0])
                self.tree.focus(children[0])

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(
            bar,
            text="Отмена",
            font=font,
            width=120,
            height=32,
            fg_color="#3A3D42",
            hover_color="#4A4E54",
            command=lambda: self._close(None),
        ).pack(side="right", padx=(8, 0))
        open_button = ctk.CTkButton(
            bar,
            text="Открыть",
            font=font,
            width=120,
            height=32,
            fg_color="#2F6F4E",
            hover_color="#3B8660",
            command=self._open_selected,
            state="normal" if rows else "disabled",
        )
        open_button.pack(side="right")
        self.protocol("WM_DELETE_WINDOW", lambda: self._close(None))
        self.bind("<Escape>", lambda _event: self._close(None))
        self.grab_set()
        self._place(master)
        self.after(30, self.lift)
        self.after(30, self.focus_force)
        self.wait_window()

    def _open_selected(self, event=None):
        if not hasattr(self, "tree"):
            return "break"
        if event is not None and getattr(event, "keysym", "") != "Return":
            region = self.tree.identify_region(event.x, event.y)
            if region in ("heading", "separator", "nothing"):
                return "break"
        selected = self.tree.selection()
        if not selected:
            return "break"
        self._close(int(selected[0]))
        return "break"

    def _close(self, value) -> None:
        self.result = value
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()

    def _place(self, master) -> None:
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = master.winfo_rootx() + max(0, (master.winfo_width() - width) // 2)
        y = master.winfo_rooty() + max(0, (master.winfo_height() - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
