import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
    datefmt="%H:%M:%S",
)

from src.config import config
from src.gui.app import App


def main():
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
