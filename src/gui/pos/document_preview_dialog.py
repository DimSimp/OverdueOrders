from __future__ import annotations

import shutil
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox
from urllib.parse import quote

import customtkinter as ctk
from PIL import Image


class DocumentPreviewDialog(ctk.CTkToplevel):
    """Preview an A4 customer PDF and expose document actions."""

    def __init__(
        self,
        master,
        pdf_path: str,
        title: str,
        customer: dict | None = None,
        email_subject: str = "",
        **kwargs,
    ):
        super().__init__(master, **kwargs)
        self._pdf_path = Path(pdf_path)
        self._customer = customer or {}
        self._email_subject = email_subject or title
        self._preview_image = None

        self.title(title)
        self.geometry("760x860")
        self.minsize(620, 680)
        self.grab_set()
        self.after(50, self.lift)

        self._build_ui(title)

    def _build_ui(self, title: str) -> None:
        header = ctk.CTkFrame(self, fg_color=("gray90", "gray18"), corner_radius=0)
        header.pack(fill="x")

        ctk.CTkLabel(
            header,
            text=title,
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(side="left", padx=16, pady=10)

        ctk.CTkButton(header, text="Print", width=82, command=self._print).pack(
            side="right", padx=(0, 12), pady=8
        )
        ctk.CTkButton(header, text="Save PDF", width=92, command=self._save_as).pack(
            side="right", padx=(0, 8), pady=8
        )
        ctk.CTkButton(header, text="Email", width=82, command=self._email).pack(
            side="right", padx=(0, 8), pady=8
        )

        body = ctk.CTkFrame(self, fg_color=("gray80", "gray12"), corner_radius=0)
        body.pack(fill="both", expand=True)

        scroll = ctk.CTkScrollableFrame(body, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=14, pady=14)

        image = self._render_preview()
        if image:
            self._preview_image = ctk.CTkImage(light_image=image, dark_image=image, size=image.size)
            ctk.CTkLabel(scroll, image=self._preview_image, text="").pack()
        else:
            ctk.CTkLabel(
                scroll,
                text=f"Preview unavailable.\n\n{self._pdf_path}",
                text_color=("gray35", "gray70"),
                font=ctk.CTkFont(size=12),
            ).pack(expand=True, pady=80)

        footer = ctk.CTkFrame(self, fg_color=("gray90", "gray18"), corner_radius=0)
        footer.pack(fill="x")
        ctk.CTkLabel(
            footer,
            text=str(self._pdf_path),
            font=ctk.CTkFont(size=10),
            text_color=("gray45", "gray60"),
        ).pack(side="left", padx=16, pady=8)
        ctk.CTkButton(footer, text="Close", width=90, command=self.destroy).pack(
            side="right", padx=12, pady=8
        )

    def _render_preview(self) -> Image.Image | None:
        try:
            import fitz

            doc = fitz.open(str(self._pdf_path))
            page = doc[0]
            zoom = 1.25
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()

            max_w = 680
            if image.width > max_w:
                ratio = max_w / image.width
                image = image.resize((max_w, int(image.height * ratio)), Image.LANCZOS)
            return image
        except Exception:
            return None

    def _print(self) -> None:
        from src.config import config
        from src.printer_utils import print_pdf

        printer = config.device.a4_printer
        if not printer:
            messagebox.showwarning(
                "No A4 Printer Configured",
                "No A4 printer is configured.\nGo to Settings -> Printers to set one up.",
                parent=self,
            )
            return
        try:
            print_pdf(str(self._pdf_path), printer)
        except Exception as exc:
            messagebox.showerror(
                "Print Failed",
                f"Could not print document:\n{exc}",
                parent=self,
            )

    def _save_as(self) -> None:
        target = filedialog.asksaveasfilename(
            parent=self,
            title="Save PDF",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=self._pdf_path.name,
        )
        if not target:
            return
        try:
            shutil.copyfile(self._pdf_path, target)
        except Exception as exc:
            messagebox.showerror(
                "Save Failed",
                f"Could not save PDF:\n{exc}",
                parent=self,
            )

    def _email(self) -> None:
        email = (self._customer.get("email") or "").strip()
        if not email:
            messagebox.showwarning(
                "No Email Address",
                "This customer does not have an email address saved.",
                parent=self,
            )
            return

        subject = quote(self._email_subject)
        body = quote(
            "Please find your quote attached.\n\n"
            f"PDF location: {self._pdf_path}\n\n"
            "Attach the PDF above before sending."
        )
        webbrowser.open(f"mailto:{quote(email)}?subject={subject}&body={body}")
