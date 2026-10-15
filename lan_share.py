#!/usr/bin/env python3
"""
Local network transfer page (Streamlit edition, stealth UI).

    pip install streamlit
    python lan_share.py

Asks for a port (Enter = random free port), then serves on ALL network
interfaces so any device on the same Wi-Fi / hotspot can open:

    http://<this-machine's-ip>:<port>

The page itself has no readable text: just icon buttons. Files live in ./storage/.
No password. Ctrl+C to stop.
"""

import base64
import importlib.util
import io
import os
import platform
import random
import re
import shutil
import socket
import struct
import subprocess
import sys
import threading
import time
import urllib.request
import zipfile
import zlib
from datetime import datetime
from pathlib import Path

SHARE_DIR = Path(__file__).resolve().parent / "storage"

# Hides every Streamlit-branded / text element and swaps the upload button for icons.
STEALTH_CSS = r"""
#MainMenu, footer, header, [data-testid="stHeader"], [data-testid="stToolbar"],
[data-testid="stDecoration"], [data-testid="stStatusWidget"],
[data-testid="stAppDeployButton"], .stDeployButton {display:none !important}
.block-container {padding-top:1.2rem !important; max-width:640px}
[data-testid="stFileUploaderDropzoneInstructions"] {display:none !important}
[data-testid="stFileUploaderDropzone"] {justify-content:center; padding:.4rem .6rem;
    min-height:0; background:transparent; border-style:dashed; border-radius:12px}
[data-testid="stFileUploaderDropzone"] button {font-size:0 !important; background:transparent;
    border:0; box-shadow:none; min-width:0; width:100%}
[data-testid="stFileUploaderDropzone"] button > * {display:none !important}
[data-testid="stFileUploaderDropzone"] button::after {content:"\002B"; font-size:24px !important;
    line-height:1.2; opacity:.6}
[class*="st-key-dir"] [data-testid="stFileUploaderDropzone"] button::after {content:"\1F4C1";
    font-size:20px !important}
"""


# ----------------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------------
def tiny_png(rgb=(236, 236, 236), size=16):
    """A plain solid-colour PNG (used as a neutral favicon)."""
    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)
    raw = b"".join(b"\x00" + bytes(rgb) * size for _ in range(size))
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw))
            + chunk(b"IEND", b""))


def safe_join(rel):
    """Map an uploaded name (may include folders) to a path INSIDE SHARE_DIR."""
    rel = rel.replace("\\", "/")
    parts = [re.sub(r'[<>:"|?*\x00-\x1f]', "_", p) for p in rel.split("/") if p not in ("", ".")]
    if not parts or ".." in parts:
        return None
    target = SHARE_DIR.joinpath(*parts).resolve()
    try:
        target.relative_to(SHARE_DIR.resolve())
    except ValueError:
        return None
    return target


# ----------------------------------------------------------------------------
# Network helpers
# ----------------------------------------------------------------------------
def get_default_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))  # nothing is sent; OS just picks the route
        return s.getsockname()[0]
    except OSError:
        return None
    finally:
        s.close()


def list_ipv4_addresses():
    found = {}
    try:
        import psutil
        for iface, addrs in psutil.net_if_addrs().items():
            for a in addrs:
                if a.family == socket.AF_INET:
                    found[a.address] = iface
    except ImportError:
        try:
            for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
                found.setdefault(info[4][0], "?")
        except OSError:
            pass
    default_ip = get_default_ip()
    if default_ip:
        found.setdefault(default_ip, "default route")
    found.pop("127.0.0.1", None)
    return found


def port_is_free(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(("0.0.0.0", port))
            return True
        except OSError:
            return False


def pick_random_port():
    for _ in range(200):
        p = random.randint(8000, 9999)
        if port_is_free(p):
            return p
    sys.exit("Could not find a free port.")


# ----------------------------------------------------------------------------
# Firewall: the #1 reason other devices can't connect
# ----------------------------------------------------------------------------
def ensure_firewall_open(port):
    system = platform.system()
    try:
        if system == "Windows":
            _open_windows_firewall(port)
        elif system == "Darwin":
            out = subprocess.run(
                ["/usr/libexec/ApplicationFirewall/socketfilterfw", "--getglobalstate"],
                capture_output=True, text=True).stdout
            if "enabled" in out.lower():
                print("  [firewall] macOS firewall is ON. If a popup asks whether Python may\n"
                      "             accept incoming connections, click ALLOW.")
        elif system == "Linux":
            out = subprocess.run(["ufw", "status"], capture_output=True, text=True).stdout
            if "Status: active" in out:
                print(f"  [firewall] ufw is active. Run:  sudo ufw allow {port}/tcp")
    except Exception:
        pass


def _open_windows_firewall(port):
    import ctypes
    name = f"Local Web {port}"

    def rule_exists():
        r = subprocess.run(["netsh", "advfirewall", "firewall", "show", "rule",
                            f"name={name}"], capture_output=True, text=True)
        return r.returncode == 0 and "No rules match" not in r.stdout

    if rule_exists():
        print(f"  [firewall] Windows Firewall rule '{name}' already present.")
        return
    args = ["advfirewall", "firewall", "add", "rule", f"name={name}", "dir=in",
            "action=allow", "protocol=TCP", f"localport={port}", "profile=any"]
    if ctypes.windll.shell32.IsUserAnAdmin():
        subprocess.run(["netsh"] + args, capture_output=True)
    else:
        print("  [firewall] Windows Firewall usually blocks other devices. Asking for\n"
              "             permission (UAC popup) to open this port - click YES.")
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", "netsh", subprocess.list2cmdline(args), None, 0)
        for _ in range(20):
            time.sleep(1)
            if rule_exists():
                break
    if rule_exists():
        print(f"  [firewall] Opened TCP port {port} (all network profiles).")
        print(f"             Remove later:  netsh advfirewall firewall delete rule name=\"{name}\"")
    else:
        print("  [firewall] Could NOT open the port automatically. Run this in an\n"
              "             Administrator Command Prompt, then retry:\n"
              f"             netsh advfirewall firewall add rule name=\"{name}\" dir=in "
              f"action=allow protocol=TCP localport={port} profile=any")


# ----------------------------------------------------------------------------
# Stealth: neutralise the HTML shell Streamlit serves before the app even loads
# ----------------------------------------------------------------------------
def patch_streamlit_shell():
    """Set a plain tab title + favicon, drop the 'Streamlit' comment, inject the
    hide-everything CSS into <head> (so nothing flashes before the app loads).
    Idempotent; a backup is kept as index.html.orig."""
    import streamlit
    idx = Path(streamlit.__file__).parent / "static" / "index.html"
    try:
        backup = idx.with_name("index.html.orig")
        if not backup.exists():
            shutil.copy2(idx, backup)
        html = backup.read_text(encoding="utf-8")          # always patch from the original
        icon = base64.b64encode(tiny_png()).decode()
        html = re.sub(r"<!--.*?-->", "", html, flags=re.S)  # licence comment names the framework
        html = re.sub(r'<link rel="modulepreload"[^>]*Streamlit[^>]*>\s*', "", html)
        html = html.replace("<title>Streamlit</title>", "<title>Home</title>")
        html = re.sub(r'<link rel="shortcut icon"[^>]*/>',
                      f'<link rel="icon" href="data:image/png;base64,{icon}" />', html)
        html = html.replace("</head>", f"<style>{STEALTH_CSS}</style></head>", 1)
        if idx.read_text(encoding="utf-8") == html:
            return "already applied"
        idx.write_text(html, encoding="utf-8")
        return "applied"
    except Exception as e:
        return f"skipped ({e.__class__.__name__}: {e})"


# ----------------------------------------------------------------------------
# Launcher
# ----------------------------------------------------------------------------
def self_test(ip, port):
    for _ in range(40):
        time.sleep(1)
        try:
            with urllib.request.urlopen(f"http://{ip}:{port}/_stcore/health", timeout=2) as r:
                if r.status == 200:
                    print(f"  [self-test] Server reachable on {ip}:{port} from this machine: YES")
                    return
        except Exception:
            pass
    print(f"  [self-test] Server reachable on {ip}:{port} from this machine: NO")


def launcher():
    if importlib.util.find_spec("streamlit") is None:
        sys.exit("Streamlit is not installed. Run:  pip install streamlit")

    print("=" * 64)
    print(" Local web page  -  open to everyone on the same Wi-Fi / hotspot")
    print("=" * 64)
    raw = input("Port number (press Enter for a random free port): ").strip()
    if raw:
        if not raw.isdigit() or not (1024 <= int(raw) <= 65535):
            sys.exit("Please enter a number between 1024 and 65535.")
        port = int(raw)
        if not port_is_free(port):
            sys.exit(f"Port {port} is already in use. Try another one.")
    else:
        port = pick_random_port()

    SHARE_DIR.mkdir(exist_ok=True)
    addrs = list_ipv4_addresses()
    default_ip = get_default_ip()

    print()
    ensure_firewall_open(port)
    print(f"  [stealth]  page shell {patch_streamlit_shell()}")

    print("\n  Open one of these on the OTHER device (same Wi-Fi / hotspot):\n")
    for ip, iface in sorted(addrs.items(), key=lambda kv: kv[0] != default_ip):
        star = "   <-- most likely" if ip == default_ip else ""
        print(f"     http://{ip}:{port}    ({iface}){star}")
    print(f"\n     http://localhost:{port}    (this machine only)\n")

    if any(ip.startswith("172.20.10.") for ip in addrs):
        print("  NOTE: 172.20.10.x is an iPhone hotspot. iPhone hotspots are widely\n"
              "        reported to isolate connected devices from each other, which\n"
              "        blocks this kind of test. If nothing shows up here when the\n"
              "        other device opens the link, that is the likely reason.\n")

    print("  A line is printed here whenever a device opens the page. If you open the\n"
          "  link on the other device and NOTHING appears, the request never arrived\n"
          "  (firewall, VPN, wrong IP, or the network blocks device-to-device traffic).\n"
          "  Ctrl+C to stop.\n")

    if default_ip:
        threading.Thread(target=self_test, args=(default_ip, port), daemon=True).start()

    env = os.environ.copy()
    env["LAN_SHARE_CHILD"] = "1"
    cmd = [
        sys.executable, "-m", "streamlit", "run", str(Path(__file__).resolve()),
        "--server.address", "0.0.0.0",          # every network interface
        "--server.port", str(port),
        "--server.headless", "true",
        "--server.maxUploadSize", "2048",       # MB
        "--browser.gatherUsageStats", "false",
        "--client.toolbarMode", "minimal",
    ]
    try:
        subprocess.run(cmd, env=env)
    except KeyboardInterrupt:
        print("\nStopped.")


# ----------------------------------------------------------------------------
# The page (runs inside Streamlit). Icons only - no readable words.
# ----------------------------------------------------------------------------
def app():
    import streamlit as st

    st.set_page_config(page_title="Home", page_icon=tiny_png(), layout="centered",
                       initial_sidebar_state="collapsed",
                       menu_items={"Get help": None, "Report a bug": None, "About": None})
    st.markdown(f"<style>{STEALTH_CSS}</style>", unsafe_allow_html=True)
    SHARE_DIR.mkdir(exist_ok=True)

    # Terminal diagnostic: one line per newly opened page
    if not st.session_state.get("_seen"):
        st.session_state["_seen"] = True
        try:
            ip = st.context.ip_address or "?"
        except Exception:
            ip = "?"
        print(f"  [{datetime.now():%H:%M:%S}] page opened by {ip}", flush=True)

    n = st.session_state.setdefault("n", 0)
    blank = "\u200b"  # invisible label

    # ---- add: files (+) and folders (folder icon) ----------------------------
    can_dir = "directory" in (st.file_uploader.__doc__ or "")
    col1, col2 = st.columns(2) if can_dir else (st.container(), None)
    with col1:
        files = st.file_uploader(blank, accept_multiple_files=True, key=f"fil{n}",
                                 label_visibility="collapsed")
    folder = []
    if col2 is not None:
        with col2:
            folder = st.file_uploader(blank, accept_multiple_files="directory", key=f"dir{n}",
                                      label_visibility="collapsed")

    saved = False
    for f in list(files or []) + list(folder or []):
        target = safe_join(f.name)          # folder uploads arrive as "dir/sub/name"
        if target is not None:
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(f.getbuffer())
            saved = True
    if saved:
        st.session_state["n"] += 1          # resets the pickers
        st.rerun()

    # ---- list ---------------------------------------------------------------
    items = sorted(p for p in SHARE_DIR.rglob("*") if p.is_file())
    if not items:
        return

    @st.cache_data(show_spinner=False, max_entries=2)
    def build_zip(signature):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for path_str, _ in signature:
                p = Path(path_str)
                z.write(p, p.relative_to(SHARE_DIR).as_posix())
        return buf.getvalue()

    signature = tuple((str(p), p.stat().st_mtime_ns) for p in items)
    st.download_button("\u2B07", build_zip(signature), file_name="archive.zip", key="zip_all")

    for i, p in enumerate(items):
        rel = p.relative_to(SHARE_DIR).as_posix()
        c1, c2, c3 = st.columns([7, 1, 1])
        c1.write(f"{rel}  \n{p.stat().st_size / 1024:,.1f} KB")
        c2.download_button("\u2B07", p.read_bytes(), file_name=p.name, key=f"get_{i}")
        if c3.button("\u2715", key=f"rm_{i}"):
            p.unlink()
            st.rerun()


if __name__ == "__main__":
    if os.environ.get("LAN_SHARE_CHILD") == "1":
        app()
    else:
        launcher()
