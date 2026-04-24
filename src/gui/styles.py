from __future__ import annotations

from tkinter import ttk

import customtkinter as ctk


def apply_styles() -> None:
    """Register all ttk Treeview styles for the application.

    Must be called once after the Tk root exists, before any windows open.
    Using 'clam' throughout because the Windows 'vista'/'default' themes
    ignore fieldbackground, leaving treeview interiors white.
    """
    dark = ctk.get_appearance_mode() == "Dark"
    style = ttk.Style()
    style.theme_use("clam")

    # ── Orders.Treeview — Daily / Afternoon ops results ───────────────────
    style.configure(
        "Orders.Treeview",
        background="#2b2b2b" if dark else "#f5f5f5",
        foreground="#ffffff" if dark else "#1a1a1a",
        fieldbackground="#2b2b2b" if dark else "#f5f5f5",
        rowheight=28,
        font=("", 12),
        borderwidth=0,
    )
    style.configure(
        "Orders.Treeview.Heading",
        background="#1f1f1f" if dark else "#d8d8d8",
        foreground="#cccccc" if dark else "#333333",
        font=("", 12, "bold"),
        relief="flat",
    )
    style.map(
        "Orders.Treeview",
        background=[("selected", "#3d5a80" if dark else "#cde0f5")],
        foreground=[("selected", "#ffffff" if dark else "#1a1a1a")],
    )

    # ── POS.Treeview — Till tab cart ──────────────────────────────────────
    style.configure(
        "POS.Treeview",
        rowheight=30, font=("Segoe UI", 11),
        background="#1c1c1c", foreground="#e0e0e0", fieldbackground="#1c1c1c",
    )
    style.configure(
        "POS.Treeview.Heading",
        font=("Segoe UI", 11, "bold"),
        background="#2a2a2a", foreground="#cccccc",
    )
    style.map(
        "POS.Treeview",
        background=[("selected", "#1f538d")],
        foreground=[("selected", "white")],
    )

    # ── Inv.Treeview — Inventory browse ──────────────────────────────────
    style.configure(
        "Inv.Treeview",
        rowheight=26, font=("Segoe UI", 11),
        background="#1c1c1c", foreground="#e0e0e0", fieldbackground="#1c1c1c",
    )
    style.configure(
        "Inv.Treeview.Heading",
        font=("Segoe UI", 11, "bold"),
        background="#2a2a2a", foreground="#cccccc",
    )
    style.map(
        "Inv.Treeview",
        background=[("selected", "#1f538d")],
        foreground=[("selected", "white")],
    )

    # ── Sales.Treeview — Daily Sales report (collapsible tx rows) ────────
    style.configure(
        "Sales.Treeview",
        rowheight=24, font=("Segoe UI", 10),
        background="#1c1c1c", foreground="#e0e0e0", fieldbackground="#1c1c1c",
    )
    style.configure(
        "Sales.Treeview.Heading",
        font=("Segoe UI", 10, "bold"),
        background="#2a2a2a", foreground="#cccccc",
    )
    style.map(
        "Sales.Treeview",
        background=[("selected", "#1f538d")],
        foreground=[("selected", "white")],
    )

    # ── Cust.Treeview — Customer browse ──────────────────────────────────
    style.configure(
        "Cust.Treeview",
        rowheight=26, font=("Segoe UI", 11),
        background="#1c1c1c", foreground="#e0e0e0", fieldbackground="#1c1c1c",
    )
    style.configure(
        "Cust.Treeview.Heading",
        font=("Segoe UI", 11, "bold"),
        background="#2a2a2a", foreground="#cccccc",
    )
    style.map(
        "Cust.Treeview",
        background=[("selected", "#1f538d")],
        foreground=[("selected", "white")],
    )

    # ── Prev.Treeview — CSV import preview ───────────────────────────────
    style.configure("Prev.Treeview", rowheight=22, font=("Segoe UI", 10))
    style.configure("Prev.Treeview.Heading", font=("Segoe UI", 10, "bold"))
