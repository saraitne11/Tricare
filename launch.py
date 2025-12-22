import os
import sys
from pathlib import Path

from streamlit.web import cli as stcli


if __name__ == "__main__":
    # Ensure working directory is the folder containing the bundled files.
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    os.chdir(base_dir)

    # Disable dev mode so we can set server.port in a frozen app.
    os.environ.setdefault("STREAMLIT_GLOBAL_DEVELOPMENT_MODE", "false")

    app_path = base_dir / "app.py"
    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--server.headless=false",
        "--server.port=8501",
        "--browser.serverAddress=localhost",
    ]
    stcli.main()
import os
import sys
import subprocess


def main() -> None:
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    app_path = os.path.join(base_dir, "app.py")

    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        app_path,
        "--server.headless=true",
        "--server.port=8501",
    ]

    env = os.environ.copy()

    # Launch Streamlit and forward exit code; swallow Ctrl+C to avoid PyInstaller error.
    proc = subprocess.Popen(cmd, cwd=base_dir, env=env)
    try:
        proc.wait()
    except KeyboardInterrupt:
        proc.terminate()
        proc.wait()
        sys.exit(0)
    sys.exit(proc.returncode)


if __name__ == "__main__":
    main()

