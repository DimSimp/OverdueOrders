from __future__ import annotations

import customtkinter as ctk


class AftMenuFrame(ctk.CTkFrame):
    """
    Afternoon Operations menu — shown when entering the workflow from the launcher.
    Offers two actions: start a fresh session or restore the last saved one.
    """

    def __init__(self, master, on_new_session, on_load_session, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._on_new_session = on_new_session
        self._on_load_session = on_load_session
        self._build_ui()

    def _build_ui(self):
        center = ctk.CTkFrame(self, fg_color="transparent")
        center.pack(expand=True, fill="both", padx=80, pady=40)

        ctk.CTkLabel(
            center,
            text="What would you like to do?",
            font=ctk.CTkFont(size=13),
            text_color=("gray40", "gray70"),
        ).pack(pady=(0, 24))

        # Start New Session
        ctk.CTkButton(
            center,
            text="Start New Session",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=64,
            command=self._on_new_session,
        ).pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(
            center,
            text="Compare today's inventory reports, fetch overdue orders, and view results",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(pady=(0, 20))

        # Load Session
        ctk.CTkButton(
            center,
            text="Load Session",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=64,
            command=self._on_load_session,
        ).pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(
            center,
            text="Resume the last saved session and go straight to the results screen",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack()
