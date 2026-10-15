#!/usr/bin/env python3
"""
LAN file share - zero dependencies (pure Python standard library).

    python lan_share.py

Asks for a port (Enter = random free port), then serves on ALL network
interfaces so any device on the same Wi-Fi / hotspot can open it:

    http://<this-machine's-ip>:<port>

Features: upload files OR whole folders, download files, download all as .zip,
delete files. Every request is logged here, so you can SEE whether the other
device is actually reaching this machine.

Shared files live in ./storage/ next to this script.
No password: anyone on the same network can use it. Ctrl+C to stop.
"""

import http.server
import json
import mimetypes
import platform
import random
import re
import shutil
import socket
import socketserver
import subprocess
import sys
import tempfile
import threading
import time
import urllib.parse
import urllib.request
import zipfile
from datetime import datetime
from pathlib import Path

SHARE_DIR = Path(__file__).resolve().parent / "storage"
SEEN_CLIENTS = set()
OWN_IPS = set()


# ----------------------------------------------------------------------------
# Network helpers
# ----------------------------------------------------------------------------
def get_default_ip():
    """IP of the interface this machine uses for its main network connection."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))  # nothing is sent; OS just picks the route
        return s.getsockname()[0]
    except OSError:
        return None
    finally:
        s.close()


def list_ipv4_addresses():
    """{ip: interface_name} for this machine, excluding loopback."""
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
                print("  [firewall] macOS firewall is ON. If a popup asks whether Python\n"
                      "             may accept incoming connections, click ALLOW.")
        elif system == "Linux":
            out = subprocess.run(["ufw", "status"], capture_output=True, text=True).stdout
            if "Status: active" in out:
                print(f"  [firewall] ufw is active. Run:  sudo ufw allow {port}/tcp")
    except Exception:
        pass


def _open_windows_firewall(port):
    import ctypes
    name = f"LAN Share {port}"

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
        print(f"             Remove later with:  netsh advfirewall firewall delete rule name=\"{name}\"")
    else:
        print("  [firewall] Could NOT open the port automatically. Run this in an\n"
              "             Administrator Command Prompt, then retry:\n"
              f"             netsh advfirewall firewall add rule name=\"{name}\" dir=in "
              f"action=allow protocol=TCP localport={port} profile=any")


# ----------------------------------------------------------------------------
# Safe path handling (blocks ../ tricks)
# ----------------------------------------------------------------------------
def safe_join(rel):
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
# Web page
# ----------------------------------------------------------------------------
PAGE = """<!doctype html>
<html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Home</title>
<style>
 :root{color-scheme:light dark;--bg:#fff;--fg:#222;--line:#e6e6e6;--mut:#999}
 @media (prefers-color-scheme:dark){:root{--bg:#121212;--fg:#ddd;--line:#2a2a2a;--mut:#777}}
 body{font-family:-apple-system,system-ui,sans-serif;background:var(--bg);color:var(--fg);
      margin:0;padding:20px;max-width:640px;margin-inline:auto}
 .bar{display:flex;gap:8px;justify-content:flex-end;margin-bottom:10px}
 .ic{width:38px;height:38px;border-radius:50%;border:1px solid var(--line);display:flex;
     align-items:center;justify-content:center;cursor:pointer;font-size:17px;color:var(--mut);
     text-decoration:none;user-select:none;background:none}
 .row{display:flex;gap:10px;align-items:center;padding:10px 0;border-top:1px solid var(--line)}
 .row a{flex:1;word-break:break-all;color:var(--fg);text-decoration:none}
 .row span{color:var(--mut);font-size:.8rem;white-space:nowrap}
 .row button{background:none;border:0;font-size:1rem;cursor:pointer;color:var(--mut)}
 progress{width:100%;height:4px} #p div{margin:6px 0;font-size:.85rem;color:var(--mut);word-break:break-all}
</style></head><body>
<div class="bar">
 <label class="ic">&#128196;<input type="file" id="a" multiple hidden></label>
 <label class="ic">&#128193;<input type="file" id="b" webkitdirectory multiple hidden></label>
 <a class="ic" id="z" href="/z" style="display:none">&#11015;</a>
</div>
<div id="p"></div>
<div id="l"></div>

<script>
const $ = id => document.getElementById(id);
function fmt(n){ if(n<1024) return n+' B'; if(n<1048576) return (n/1024).toFixed(1)+' KB';
  if(n<1073741824) return (n/1048576).toFixed(1)+' MB'; return (n/1073741824).toFixed(2)+' GB'; }

async function refresh(){
  let items = [];
  try{ items = await (await fetch('/l')).json(); }catch(e){ return; }
  const box = $('l'); box.textContent = '';
  $('z').style.display = items.length ? 'flex' : 'none';
  for(const it of items){
    const row = document.createElement('div'); row.className = 'row';
    const a = document.createElement('a');
    a.href = '/f/' + it.path.split('/').map(encodeURIComponent).join('/');
    a.textContent = it.path;
    const s = document.createElement('span'); s.textContent = fmt(it.size);
    const x = document.createElement('button'); x.textContent = '\\u2715';
    x.onclick = async () => {
      if(confirm(it.path + ' ?')){
        await fetch('/x?q=' + encodeURIComponent(it.path), {method:'POST'}); refresh(); }
    };
    row.append(a, s, x); box.append(row);
  }
}

function sendOne(item, rel, label, bar){
  return new Promise(res => {
    const x = new XMLHttpRequest();
    x.open('POST', '/u?q=' + encodeURIComponent(rel));
    x.upload.onprogress = e => { if(e.lengthComputable) bar.value = e.loaded / e.total * 100; };
    x.onload = () => { label.textContent = (x.status == 200 ? '\\u2713 ' : '\\u2717 ') + rel +
                                           ' (' + fmt(item.size) + ')'; bar.remove(); res(x.status == 200); };
    x.onerror = () => { label.textContent = '\\u2717 ' + rel; res(false); };
    x.send(item);
  });
}

async function addAll(list){
  for(const item of list){
    const rel = item.webkitRelativePath || item.name;
    const wrap = document.createElement('div'), label = document.createElement('div'),
          bar = document.createElement('progress');
    bar.max = 100; bar.value = 0; label.textContent = rel;
    wrap.append(label, bar); $('p').prepend(wrap);
    await sendOne(item, rel, label, bar);
  }
  refresh();
}
$('a').onchange = e => { addAll([...e.target.files]); e.target.value = ''; };
$('b').onchange = e => { addAll([...e.target.files]); e.target.value = ''; };

refresh(); setInterval(refresh, 5000);
</script></body></html>
"""


# ----------------------------------------------------------------------------
# HTTP server
# ----------------------------------------------------------------------------
class Handler(http.server.BaseHTTPRequestHandler):
    server_version = "S"
    sys_version = ""

    # -- logging: this is the key diagnostic --------------------------------
    def log_message(self, fmt, *args):
        if self.headers.get("X-Selftest"):
            return
        ip = self.client_address[0]
        is_local = ip.startswith("127.") or ip in OWN_IPS
        if not is_local and ip not in SEEN_CLIENTS:
            SEEN_CLIENTS.add(ip)
            print(f"\n  >>> First request from ANOTHER device ({ip}) - the network path WORKS! <<<\n")
        print(f"  [{datetime.now():%H:%M:%S}] {ip}  {self.command} {self.path}")

    # -- small response helpers -----------------------------------------------
    def _send(self, code, body=b"", ctype="text/plain; charset=utf-8", extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    # -- routes ---------------------------------------------------------------
    def do_GET(self):
        try:
            url = urllib.parse.urlsplit(self.path)
            route = url.path
            if route == "/":
                self._send(200, PAGE.encode("utf-8"), "text/html; charset=utf-8")
            elif route == "/p":
                self._send(200, b"ok")
            elif route == "/l":
                items = [{"path": p.relative_to(SHARE_DIR).as_posix(), "size": p.stat().st_size}
                         for p in sorted(SHARE_DIR.rglob("*")) if p.is_file()]
                self._send(200, json.dumps(items).encode(), "application/json")
            elif route == "/z":
                self._serve_zip()
            elif route.startswith("/f/"):
                self._serve_file(urllib.parse.unquote(route[len("/f/"):]))
            else:
                self._send(404, b"Not found")
        except (BrokenPipeError, ConnectionResetError):
            pass

    do_HEAD = do_GET

    def do_POST(self):
        try:
            url = urllib.parse.urlsplit(self.path)
            query = urllib.parse.parse_qs(url.query)
            rel = query.get("q", [""])[0]
            if url.path == "/u":
                self._receive_upload(rel)
            elif url.path == "/x":
                target = safe_join(rel)
                if target and target.is_file():
                    target.unlink()
                    self._send(200, b"deleted")
                else:
                    self._send(404, b"Not found")
            else:
                self._send(404, b"Not found")
        except (BrokenPipeError, ConnectionResetError):
            pass

    # -- actions --------------------------------------------------------------
    def _receive_upload(self, rel):
        target = safe_join(rel)
        length = int(self.headers.get("Content-Length", 0))
        if target is None:
            remaining = length
            while remaining > 0:                    # drain so the client sees our reply
                remaining -= len(self.rfile.read(min(65536, remaining)))
            self._send(400, b"Bad path")
            return
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = target.with_name(target.name + ".part")
        remaining = length
        with open(tmp, "wb") as out:
            while remaining > 0:
                chunk = self.rfile.read(min(1 << 20, remaining))
                if not chunk:
                    break
                out.write(chunk)
                remaining -= len(chunk)
        if remaining:                               # connection dropped mid-upload
            tmp.unlink(missing_ok=True)
            return
        tmp.replace(target)
        self._send(200, b"saved")

    def _serve_file(self, rel):
        target = safe_join(rel)
        if target is None or not target.is_file():
            self._send(404, b"Not found")
            return
        ctype = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(target.stat().st_size))
        self.send_header("Content-Disposition",
                         "attachment; filename*=UTF-8''" + urllib.parse.quote(target.name))
        self.end_headers()
        if self.command != "HEAD":
            with open(target, "rb") as f:
                shutil.copyfileobj(f, self.wfile, 1 << 20)

    def _serve_zip(self):
        with tempfile.TemporaryFile() as tmp:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
                for p in sorted(SHARE_DIR.rglob("*")):
                    if p.is_file() and not p.name.endswith(".part"):
                        z.write(p, p.relative_to(SHARE_DIR).as_posix())
            size = tmp.tell()
            tmp.seek(0)
            self.send_response(200)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Length", str(size))
            self.send_header("Content-Disposition", 'attachment; filename="archive.zip"')
            self.end_headers()
            if self.command != "HEAD":
                shutil.copyfileobj(tmp, self.wfile, 1 << 20)


class Server(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True
    allow_reuse_address = True


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def self_test(ip, port):
    """Confirm the server answers on this machine's network address."""
    time.sleep(0.5)
    try:
        req = urllib.request.Request(f"http://{ip}:{port}/p", headers={"X-Selftest": "1"})
        with urllib.request.urlopen(req, timeout=3) as r:
            ok = r.read() == b"ok"
    except Exception:
        ok = False
    print(f"  [self-test] Server reachable on {ip}:{port} from this machine: "
          f"{'YES' if ok else 'NO'}")


def main():
    print("=" * 64)
    print(" LAN Share  -  open to everyone on the same Wi-Fi / hotspot")
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
    try:
        server = Server(("0.0.0.0", port), Handler)   # 0.0.0.0 = every network interface
    except OSError as e:
        sys.exit(f"Could not start server on port {port}: {e}")

    addrs = list_ipv4_addresses()
    OWN_IPS.update(addrs)
    default_ip = get_default_ip()

    print()
    ensure_firewall_open(port)

    print("\n  Open one of these on the OTHER device (same Wi-Fi / hotspot):\n")
    for ip, iface in sorted(addrs.items(), key=lambda kv: kv[0] != default_ip):
        star = "   <-- most likely" if ip == default_ip else ""
        print(f"     http://{ip}:{port}    ({iface}){star}")
    print(f"\n     http://localhost:{port}    (this machine only)\n")

    if any(ip.startswith("172.20.10.") for ip in addrs):
        print("  NOTE: 172.20.10.x is an iPhone hotspot. iPhone hotspots are widely\n"
              "        reported to isolate connected devices from each other, which\n"
              "        blocks this kind of test. If nothing shows up below when the\n"
              "        other device opens the link, that is the likely reason.\n")

    print("  Every request will be logged below. If you open the link on the other")
    print("  device and NOTHING appears here, the request never arrived (firewall,\n"
          "  VPN, wrong IP, or the network blocks device-to-device traffic).")
    print("  Ctrl+C to stop.\n")

    if default_ip:
        threading.Thread(target=self_test, args=(default_ip, port), daemon=True).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
    
