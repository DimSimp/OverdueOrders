"""
Register the .scar file extension so double-clicking a session file
opens this application directly to the results screen.

Run this script ONCE (as administrator if needed):
    python register_file_association.py

To undo the registration, run:
    python register_file_association.py --unregister
"""
import os
import sys
import winreg

EXTENSION = ".scar"
PROG_ID   = "ScarlettsOverdueOrders.Session"
APP_NAME  = "Scarlett Music Overdue Orders Session"

def _root() -> str:
    return os.path.dirname(os.path.abspath(__file__))

def _main_py() -> str:
    return os.path.join(_root(), "main.py")

def _scar_ico() -> str:
    return os.path.join(_root(), "scar.ico")

def _pythonw() -> str:
    # Use pythonw.exe so no console window appears when double-clicking
    exe = sys.executable
    pythonw = os.path.join(os.path.dirname(exe), "pythonw.exe")
    return pythonw if os.path.exists(pythonw) else exe

def register():
    main_py  = _main_py()
    scar_ico = _scar_ico()
    pythonw  = _pythonw()
    command  = f'"{pythonw}" "{main_py}" "%1"'

    # HKCU\Software\Classes\.scar  →  ScarlettsOverdueOrders.Session
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                          rf"Software\Classes\{EXTENSION}") as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, PROG_ID)

    # HKCU\Software\Classes\<ProgID>
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                          rf"Software\Classes\{PROG_ID}") as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, APP_NAME)

    # HKCU\Software\Classes\<ProgID>\DefaultIcon  →  scar.ico
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                          rf"Software\Classes\{PROG_ID}\DefaultIcon") as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, scar_ico)

    # HKCU\Software\Classes\<ProgID>\shell\open\command
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                          rf"Software\Classes\{PROG_ID}\shell\open\command") as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, command)

    print(f"Registered: {EXTENSION}  →  {command}")
    print(f"File icon:  {scar_ico}")
    print("You may need to log out and back in for the icon to refresh in Explorer.")

def unregister():
    for path in [
        rf"Software\Classes\{PROG_ID}\shell\open\command",
        rf"Software\Classes\{PROG_ID}\shell\open",
        rf"Software\Classes\{PROG_ID}\shell",
        rf"Software\Classes\{PROG_ID}\DefaultIcon",
        rf"Software\Classes\{PROG_ID}",
        rf"Software\Classes\{EXTENSION}",
    ]:
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
        except FileNotFoundError:
            pass
    print("File association removed.")

if __name__ == "__main__":
    if "--unregister" in sys.argv:
        unregister()
    else:
        register()
