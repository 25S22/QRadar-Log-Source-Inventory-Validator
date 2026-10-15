"""
LAN file-share feasibility test (single file), v2 - works with ANY IP range.

Usage:
    pip install streamlit psutil      # psutil is optional but lists all your networks
    python lan_share.py

What it does:
  1. Asks for a port (Enter = random free port)
  2. Asks for an optional PIN
  3. Lists every network address on this machine (any range: 10.x, 172.x,
     192.168.x, hotspot ranges, Tailscale 100.x ...) and lets you choose which
     one to serve on, or all of them
  4. Starts a Streamlit site and prints the URL(s) to open from your other device

Shared files live in ./lan_shared/ next to this script.
"""

import io
import ipaddress
import os
import random
import socket
import subprocess
import sys
import zipfile
from pathlib import Path

SHARE_DIR = Path(__file__).resolve().parent / "lan_shared"


# ----------------------------------------------------------------------------
# Helpers used by the launcher
# ----------------------------------------------------------------------------
def get_lan_ip() -> str:
    """IP this machine uses for its default route (the 'main' network)."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))  # no packet is actually sent
        return s.getsockname()[0]
    except OSError:
        try:
            return socket.gethostbyname(socket.gethostname())
        except OSError:
            return "127.0.0.1"
    finally:
        s.close()


def classify_ip(ip: str) -> str:
    """Human-readable label for which range an address belongs to."""
    try:
        a = ipaddress.ip_address(ip)
    except ValueError:
        return "invalid"
    if a.is_loopback:
        return "loopback (this machine only)"
    if a.is_link_local:
        return "link-local 169.254.x.x (no DHCP / direct cable)"
    if a in ipaddress.ip_network("100.64.0.0/10"):
        return "100.64.0.0/10 (Tailscale / carrier-grade NAT)"
    if a in ipaddress.ip_network("10.0.0.0/8"):
        return "private 10.0.0.0/8"
    if a in ipaddress.ip_network("172.16.0.0/12"):
        return "private 172.16.0.0/12"
    if a in ipaddress.ip_network("192.168.0.0/16"):
        return "private 192.168.0.0/16"
    if a.is_private:
        return "private (other)"
    return "PUBLIC address (not a LAN range)"


def list_ipv4_addresses() -> dict:
    """Return {ip: interface_name} for this machine, excluding loopback."""
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
        found.setdefault(get_lan_ip(), "default route")
        print("(Tip: 'pip install psutil' to list every network adapter.)\n")
    found.pop("127.0.0.1", None)
    return found


def can_bind(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind((host, port))
            return True
        except OSError:
            return False


def pick_random_port() -> int:
    for _ in range(200):
        p = random.randint(8000, 9999)
        if can_bind("0.0.0.0", p):
            return p
    raise RuntimeError("Could not find a free port")


def launcher() -> None:
    print("=" * 64)
    print(" LAN file-share test (any IP range)")
    print("=" * 64)

    # ---- port ----
    raw = input("Port number (press Enter for a random free port): ").strip()
    if raw:
        if not raw.isdigit() or not (1024 <= int(raw) <= 65535):
            sys.exit("Please enter a number between 1024 and 65535.")
        port = int(raw)
        if not can_bind("0.0.0.0", port):
            sys.exit(f"Port {port} is already in use. Try another one.")
    else:
        port = pick_random_port()

    # ---- pin ----
    pin = input("Optional PIN to protect the site (Enter for none): ").strip()

    # ---- which network / IP range to serve on ----
    addrs = list(list_ipv4_addresses().items())
    print("\nNetwork addresses found on this machine:")
    for i, (ip, iface) in enumerate(addrs, 1):
        print(f"  [{i}] {ip:<16} {iface:<14} {classify_ip(ip)}")
    print("  [A] all interfaces (0.0.0.0)   <- default")
    choice = input("Serve on which? Enter = all, a number, or type an IP: ").strip()

    if choice == "" or choice.lower() == "a":
        bind = "0.0.0.0"
    elif choice.isdigit() and 1 <= int(choice) <= len(addrs):
        bind = addrs[int(choice) - 1][0]
    else:
        bind = choice
        try:
            ipaddress.ip_address(bind)
        except ValueError:
            sys.exit(f"'{bind}' is not a valid IP address.")
    if not can_bind(bind, port):
        sys.exit(f"Can't bind to {bind}:{port}. That IP isn't assigned to this "
                 f"machine, or the port is taken.")

    SHARE_DIR.mkdir(exist_ok=True)

    print("\n  Open from this machine:  http://localhost:%d" % port)
    urls = [ip for ip, _ in addrs] if bind == "0.0.0.0" else [bind]
    for ip in urls:
        print(f"  From another device:     http://{ip}:{port}   ({classify_ip(ip)})")
    print("\n  The other device must be on the SAME network/subnet as the")
    print("  address you use (e.g. both on 10.20.30.x). Ctrl+C to stop.\n")

    env = os.environ.copy()
    env["LAN_SHARE_CHILD"] = "1"
    env["LAN_SHARE_PIN"] = pin

    cmd = [
        sys.executable, "-m", "streamlit", "run", str(Path(__file__).resolve()),
        "--server.address", bind,
        "--server.port", str(port),
        "--server.headless", "true",
        "--server.maxUploadSize", "2048",  # MB
        "--browser.gatherUsageStats", "false",
    ]
    try:
        subprocess.run(cmd, env=env)
    except KeyboardInterrupt:
        print("\nStopped.")


# ----------------------------------------------------------------------------
# The Streamlit app (only runs when launched by the launcher above)
# ----------------------------------------------------------------------------
def app() -> None:
    import streamlit as st

    st.set_page_config(page_title="LAN Share", page_icon="📡")
    SHARE_DIR.mkdir(exist_ok=True)

    pin = os.environ.get("LAN_SHARE_PIN", "")
    if pin and not st.session_state.get("authed"):
        st.title("📡 LAN Share")
        entered = st.text_input("Enter PIN", type="password")
        if st.button("Unlock"):
            if entered == pin:
                st.session_state["authed"] = True
                st.rerun()
            else:
                st.error("Wrong PIN")
        st.stop()

    st.title("📡 LAN Share")

    client_ip = "unknown"
    try:
        client_ip = st.context.ip_address or "unknown"
    except Exception:
        pass
    st.success(f"Connected. Your device's IP as seen by the server: **{client_ip}**")
    st.caption(f"Server host: {socket.gethostname()} · default LAN IP: {get_lan_ip()}")

    st.subheader("⬆️ Upload")
    st.caption("Folders can't be picked directly in a browser upload box. "
               "Zip the folder first, or select many files at once.")
    uploader_key = st.session_state.setdefault("uk", 0)
    files = st.file_uploader("Choose files", accept_multiple_files=True,
                             key=f"up_{uploader_key}")
    if files and st.button(f"Save {len(files)} file(s) to shared folder"):
        for f in files:
            safe_name = os.path.basename(f.name)  # block path traversal
            (SHARE_DIR / safe_name).write_bytes(f.getbuffer())
        st.session_state["uk"] += 1
        st.rerun()

    st.subheader("⬇️ Shared files")
    items = sorted(p for p in SHARE_DIR.rglob("*") if p.is_file())
    if not items:
        st.info("Nothing shared yet.")
    else:
        for i, p in enumerate(items):
            rel = p.relative_to(SHARE_DIR)
            size_kb = p.stat().st_size / 1024
            c1, c2, c3 = st.columns([5, 2, 1])
            c1.write(f"**{rel}**  \n{size_kb:,.1f} KB")
            c2.download_button("Download", data=p.read_bytes(),
                               file_name=p.name, key=f"dl_{i}")
            if c3.button("🗑", key=f"rm_{i}", help="Delete"):
                p.unlink()
                st.rerun()

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for p in items:
                z.write(p, p.relative_to(SHARE_DIR))
        st.download_button("Download everything as .zip", buf.getvalue(),
                           file_name="lan_shared.zip")


if __name__ == "__main__":
    if os.environ.get("LAN_SHARE_CHILD") == "1":
        app()
    else:
        launcher()
