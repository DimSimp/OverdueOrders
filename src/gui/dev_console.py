from __future__ import annotations

from datetime import datetime
from tkinter import filedialog

import customtkinter as ctk

import src.log_memory as log_memory

_instance: DevConsoleWindow | None = None


def open_dev_console(master: ctk.CTk | ctk.CTkToplevel) -> None:
    """Open (or raise) the dev console window."""
    global _instance
    try:
        if _instance is not None and _instance.winfo_exists():
            _instance.lift()
            _instance.focus_force()
            return
    except Exception:
        pass
    _instance = DevConsoleWindow(master)


class DevConsoleWindow(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Dev Console — Log Viewer")
        self.geometry("960x500")
        self.minsize(600, 300)

        # ── Toolbar ───────────────────────────────────────────────────────
        toolbar = ctk.CTkFrame(self, height=40, corner_radius=0, fg_color=("gray85", "gray20"))
        toolbar.pack(fill="x", side="top")
        toolbar.pack_propagate(False)

        ctk.CTkButton(
            toolbar, text="Export log…", width=110, height=28,
            command=self._export,
        ).pack(side="left", padx=8, pady=6)

        ctk.CTkButton(
            toolbar, text="Clear view", width=100, height=28,
            fg_color=("gray70", "gray30"), hover_color=("gray60", "gray40"),
            command=self._clear_view,
        ).pack(side="left", padx=(0, 8), pady=6)

        self._line_count_label = ctk.CTkLabel(
            toolbar, text="",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        )
        self._line_count_label.pack(side="right", padx=12)

        # ── Log textbox ───────────────────────────────────────────────────
        self._textbox = ctk.CTkTextbox(
            self,
            font=ctk.CTkFont(family="Courier New", size=11),
            wrap="none",
            state="disabled",
        )
        self._textbox.pack(fill="both", expand=True, padx=6, pady=(4, 6))

        self._cleared_at = 0  # skip lines before this index when rendering
        self._schedule_refresh()

    # ── Refresh loop ──────────────────────────────────────────────────────

    def _refresh(self) -> None:
        lines = log_memory.get_lines()
        visible = lines[self._cleared_at:]
        text = "\n".join(visible)
        self._textbox.configure(state="normal")
        self._textbox.delete("1.0", "end")
        self._textbox.insert("end", text)
        self._textbox.configure(state="disabled")
        self._textbox.see("end")
        self._line_count_label.configure(text=f"{len(visible)} lines")

    def _schedule_refresh(self) -> None:
        try:
            if not self.winfo_exists():
                return
            self._refresh()
            self.after(500, self._schedule_refresh)
        except Exception:
            pass

    # ── Actions ───────────────────────────────────────────────────────────

    def _clear_view(self) -> None:
        """Hide all lines logged so far (doesn't discard them from memory)."""
        self._cleared_at = len(log_memory.get_lines())
        self._refresh()

    def _export(self) -> None:
        lines = log_memory.get_lines()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = filedialog.asksaveasfilename(
            parent=self,
            title="Export log",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"scarlett_log_{ts}.txt",
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
