"""Launch StudentsGUIDE locally and open its browser interface when ready."""
from __future__ import annotations

import os
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")
HOST = "127.0.0.1"


def find_available_port(preferred: int) -> int:
    """Choose a free loopback port, starting with the configured preference."""
    for candidate in range(preferred, preferred + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind((HOST, candidate))
                return candidate
            except OSError:
                continue
    raise RuntimeError("No free local port was found. Close another local server and try again.")


def wait_until_ready(process: subprocess.Popen, url: str, timeout: float = 30.0) -> bool:
    deadline = time.monotonic() + timeout
    request = urllib.request.Request(f"{url}/api/status", headers={"User-Agent": "StudentsGUIDE-Launcher"})
    while time.monotonic() < deadline:
        if process.poll() is not None:
            return False
        try:
            with urllib.request.urlopen(request, timeout=1.0) as response:
                if response.status == 200:
                    return True
        except (urllib.error.URLError, TimeoutError, OSError):
            time.sleep(0.35)
    return False


def main() -> int:
    preferred_port = int(os.getenv("PORT", "8000"))
    try:
        port = find_available_port(preferred_port)
    except RuntimeError as exc:
        print(str(exc))
        return 1
    url = f"http://{HOST}:{port}"
    command = [
        sys.executable, "-m", "uvicorn", "app:app",
        "--host", HOST, "--port", str(port), "--log-level", "warning",
    ]
    try:
        server = subprocess.Popen(command, cwd=BASE_DIR)
    except OSError as exc:
        print(f"Could not start StudentsGUIDE: {exc}")
        return 1

    print("Starting StudentsGUIDE.com...")
    print(f"Local app address: {url}")
    print("Keep this window open while using the application.")

    if wait_until_ready(server, url):
        print("StudentsGUIDE is ready. Opening your browser.")
        webbrowser.open(url)
    else:
        print("StudentsGUIDE did not become ready. Check the error shown above.")
        if server.poll() is None:
            server.terminate()
        return 1

    try:
        return server.wait()
    except KeyboardInterrupt:
        print("\nStopping StudentsGUIDE...")
        server.terminate()
        try:
            return server.wait(timeout=8)
        except subprocess.TimeoutExpired:
            server.kill()
            return server.wait()


if __name__ == "__main__":
    raise SystemExit(main())
