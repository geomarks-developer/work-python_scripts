import os
import sys
from streamlit.web import cli as stcli

if getattr(sys, "frozen", False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

app_path = os.path.join(base_dir, "app.py")

if __name__ == "__main__":
    sys.argv = [
        "streamlit",
        "run",
        app_path,
    ]

    raise SystemExit(stcli.main())