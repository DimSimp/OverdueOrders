import logging
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

# When packaged as a windowed exe there is no console, so log to a file instead.
if getattr(sys, "frozen", False):
    _log_path = Path(sys.executable).parent / "scarlett_aio.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
        datefmt="%H:%M:%S",
        filename=str(_log_path),
        filemode="a",
        encoding="utf-8",
    )
else:
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
        datefmt="%H:%M:%S",
    )

from src.config import config
from src.gui.app import App


def _register_if_needed():
    """
    When running as a packaged exe, silently register the .scar file association
    the first time (or whenever the exe has moved to a new path).
    Does nothing when running from source.
    """
    if not getattr(sys, "frozen", False):
        return

    import winreg

    EXTENSION = ".scar"
    PROG_ID   = "ScarlettsOverdueOrders.Session"
    APP_NAME  = "Scarlett Music Overdue Orders Session"
    exe       = sys.executable
    command   = f'"{exe}" "%1"'
    scar_ico  = os.path.join(os.path.dirname(exe), "scar.ico")

    # Check if already registered to this exact exe path
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            rf"Software\Classes\{PROG_ID}\shell\open\command",
        ) as k:
            current = winreg.QueryValueEx(k, "")[0]
            if current.lower() == command.lower():
                return  # Already up to date
    except FileNotFoundError:
        pass

    # Register (or re-register after the exe moved)
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"Software\Classes\{EXTENSION}") as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, PROG_ID)

    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"Software\Classes\{PROG_ID}") as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, APP_NAME)

    if os.path.exists(scar_ico):
        with winreg.CreateKey(
            winreg.HKEY_CURRENT_USER, rf"Software\Classes\{PROG_ID}\DefaultIcon"
        ) as k:
            winreg.SetValueEx(k, "", 0, winreg.REG_SZ, scar_ico)

    with winreg.CreateKey(
        winreg.HKEY_CURRENT_USER, rf"Software\Classes\{PROG_ID}\shell\open\command"
    ) as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, command)


def main():
    _register_if_needed()

    try:
        config.load()
    except FileNotFoundError:
        import customtkinter as ctk
        from tkinter import messagebox
        root = ctk.CTk()
        root.withdraw()
        messagebox.showerror(
            "Configuration Missing",
            "config.json not found.\n\n"
            "Copy config.example.json to config.json and fill in your "
            "Neto API key and eBay credentials, then restart the application."
        )
        sys.exit(1)
    except Exception as e:
        import customtkinter as ctk
        from tkinter import messagebox
        root = ctk.CTk()
        root.withdraw()
        messagebox.showerror("Configuration Error", f"Failed to load config.json:\n\n{e}")
        sys.exit(1)

    # If a .scar file was passed as an argument (e.g. via file association),
    # open it automatically on startup
    startup_session = sys.argv[1] if len(sys.argv) > 1 else None

    app = App(config, startup_session=startup_session)
    app.mainloop()


if __name__ == "__main__":
    main()
