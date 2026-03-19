"""
run.py — launcher for the packaged (PyInstaller) build.

When running from source, use: streamlit run app.py
When running the packaged exe, this file is the entrypoint.
"""
import os
import sys
from pathlib import Path

# Ensure the bundled app directory is on sys.path
if getattr(sys, "frozen", False):
    # Running as PyInstaller bundle
    bundle_dir = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    sys.path.insert(0, str(bundle_dir))
    os.environ.setdefault("STREAMLIT_SERVER_HEADLESS", "true")
    os.environ.setdefault("STREAMLIT_SERVER_PORT", "8501")
    os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")

from streamlit.web import cli as stcli

app_path = Path(__file__).parent / "app.py"
sys.argv = ["streamlit", "run", str(app_path), "--server.headless=true"]
stcli.main()
