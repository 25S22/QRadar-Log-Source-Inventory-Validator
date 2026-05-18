"""
SIEM Signoff Dashboard  v4.0
────────────────────────────────────────────────────────────────────────────
Standalone dashboard for signoff_data.json.
Zero external dependencies — just Python stdlib + a browser.

PERSISTENCE MODEL
  Every edit and every delete is written to signoff_data.json immediately
  via an HTTP POST to /save.  There is no separate "Save" step.
  The file is written atomically (temp → os.replace) so a crash mid-write
  leaves the original intact.

FEATURES
  • Period filter: 7D / 30D / 90D
  • Full-text search across hosts, sender, subject, notes
  • Sort on any column
  • Paginated table (15 rows per page)
  • Stat cards + status/host bar charts refresh with every filter
  • Per-row EDIT panel (slide-in): status override, resolved flag, free-text notes
    plus read-only view of sender, subject, host type details
  • Notes auto-save as you type (debounced 600ms)
  • Per-row DELETE with confirmation modal
  • All changes auto-saved to signoff_data.json instantly
  • Single-level Ctrl+Z undo per session
  • JSON export of full dataset
  • Keyboard shortcuts: Esc closes panel/modal, Ctrl+Z undoes last action

USAGE
  python signoff_dashboard.py
  python signoff_dashboard.py --port 9000
  python signoff_dashboard.py --minutes 0   # 0 = run until Ctrl+C (default)
────────────────────────────────────────────────────────────────────────────
"""

import argparse
import json
import os
import socket
import tempfile
import threading
import time
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer

# ─── Paths ────────────────────────────────────────────────────────────────────
_DIR              = os.path.dirname(os.path.abspath(__file__))
SIGNOFF_DATA_PATH = os.path.join(_DIR, 'signoff_data.json')
DEFAULT_PORT      = 8745
DEFAULT_MINUTES   = 0     # 0 = run until Ctrl+C


# ══════════════════════════════════════════════════════════════════════════════
# PYTHON HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


def _atomic_write(path: str, data: dict) -> None:
    """Write JSON atomically: temp file → os.replace. Crash-safe."""
    dir_ = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(dir=dir_, suffix='.tmp')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as fh:
            json.dump(data, fh, indent=2, ensure_ascii=False)
        os.replace(tmp, path)
    except Exception:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def _find_free_port(preferred: int) -> int:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        s.bind(('127.0.0.1', preferred))
        s.close()
        return preferred
    except OSError:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.bind(('127.0.0.1', 0))
        p = s.getsockname()[1]
        s.close()
        return p


# ══════════════════════════════════════════════════════════════════════════════
# HTTP HANDLER
# ══════════════════════════════════════════════════════════════════════════════

class Handler(BaseHTTPRequestHandler):

    def do_GET(self):
        if self.path in ('/', '/index.html'):
            self._ok('text/html; charset=utf-8', HTML.encode())
        elif self.path == '/data.json':
            try:
                with open(SIGNOFF_DATA_PATH, 'rb') as f:
                    body = f.read()
                self._ok('application/json', body)
            except FileNotFoundError:
                self._send(404, 'text/plain', b'signoff_data.json not found')
        else:
            self._send(404, 'text/plain', b'not found')

    def do_POST(self):
        length = int(self.headers.get('Content-Length', 0))
        raw    = self.rfile.read(length)
        if self.path == '/save':
            try:
                data = json.loads(raw.decode('utf-8'))
                if not isinstance(data.get('entries'), list):
                    raise ValueError("payload missing 'entries' list")
                _atomic_write(SIGNOFF_DATA_PATH, data)
                n = len(data['entries'])
                _log(f"[save] {n} entries written → {SIGNOFF_DATA_PATH}")
                self._json(200, {'ok': True, 'entries': n,
                                 'saved_at': datetime.now().isoformat()})
            except Exception as exc:
                _log(f"[save ERROR] {exc}")
                self._json(500, {'error': str(exc)})
        else:
            self._send(404, 'text/plain', b'not found')

    def do_OPTIONS(self):
        self.send_response(200)
        for k, v in [('Access-Control-Allow-Origin', '*'),
                     ('Access-Control-Allow-Methods', 'GET, POST, OPTIONS'),
                     ('Access-Control-Allow-Headers', 'Content-Type')]:
            self.send_header(k, v)
        self.end_headers()

    def _ok(self, ct, body):
        self._send(200, ct, body)

    def _send(self, code, ct, body):
        self.send_response(code)
        self.send_header('Content-Type', ct)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-cache')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code, obj):
        self._send(code, 'application/json', json.dumps(obj).encode())

    def log_message(self, *_):
        pass  # silence per-request console noise


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD HTML  — fully self-contained, zero external files
# ══════════════════════════════════════════════════════════════════════════════

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SIEM Signoff Dashboard</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
:root{
  --bg:#08090f;
  --s1:#0f1118;
  --s2:#14171f;
  --s3:#1b1f2c;
  --bd:#1e2336;
  --bd2:#262d44;
  --green:#00d48a; --green-g:#00d48a18;
  --amber:#f5a623; --amber-g:#f5a62318;
  --red:#f03e5a;   --red-g:#f03e5a18;
  --blue:#3d8ef0;  --blue-g:#3d8ef018;
  --purple:#8b5cf6;--purple-g:#8b5cf618;
  --text:#c9d1e8;
  --muted:#46527a;
  --muted2:#272f4a;
  --mono:'Consolas','Courier New',monospace;
  --sans:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
  --r:8px;
  --pw:500px;
}
html{scroll-behavior:smooth;}
body{background:var(--bg);color:var(--text);font-family:var(--sans);
     font-size:13px;min-height:100vh;overflow-x:hidden;}
::-webkit-scrollbar{width:4px;height:4px;}
::-webkit-scrollbar-track{background:var(--bg);}
::-webkit-scrollbar-thumb{background:var(--bd2);border-radius:2px;}

/* ── Header ──────────────────────────────────────────────── */
header{
  height:56px;background:var(--s1);border-bottom:1px solid var(--bd);
  display:flex;align-items:center;gap:14px;padding:0 24px;
  position:sticky;top:0;z-index:100;
}
.logo{display:flex;align-items:center;gap:9px;flex-shrink:0;}
.pulse{width:7px;height:7px;border-radius:50%;background:var(--green);
       box-shadow:0 0 6px var(--green);animation:pu 2.5s ease-in-out infinite;}
@keyframes pu{0%,100%{opacity:1;}50%{opacity:.2;}}
.logo h1{font-family:var(--mono);font-size:11px;font-weight:700;
          letter-spacing:2px;text-transform:uppercase;color:var(--text);}
.logo small{display:block;font-size:9px;color:var(--muted);letter-spacing:1px;font-family:var(--mono);}
.spacer{flex:1;}
.hdr-r{display:flex;align-items:center;gap:8px;flex-wrap:wrap;}

/* ── Buttons ─────────────────────────────────────────────── */
.btn{
  border:1px solid var(--bd2);background:transparent;color:var(--muted);
  border-radius:6px;font-size:11px;padding:5px 12px;cursor:pointer;
  font-family:var(--mono);font-weight:600;letter-spacing:.3px;
  transition:all .14s;white-space:nowrap;
}
.btn:hover:not(:disabled){border-color:var(--blue);color:var(--blue);background:var(--blue-g);}
.btn:disabled{opacity:.3;cursor:not-allowed;}
.btn-green{border-color:var(--green);color:var(--green);}
.btn-green:hover:not(:disabled){background:var(--green-g);}
.btn-amber{border-color:var(--amber);color:var(--amber);}
.btn-amber:hover:not(:disabled){background:var(--amber-g);}
.btn-red{border-color:var(--red);color:var(--red);}
.btn-red:hover:not(:disabled){background:var(--red);color:#fff;}
.btn-ghost{border-color:transparent;color:var(--muted);}
.btn-ghost:hover:not(:disabled){border-color:var(--bd2);color:var(--text);background:transparent;}

/* ── Save indicator ──────────────────────────────────────── */
#saveInd{
  font-family:var(--mono);font-size:10px;padding:4px 10px;
  border-radius:5px;border:1px solid transparent;transition:all .2s;white-space:nowrap;
}
#saveInd.idle  {color:var(--muted);}
#saveInd.saving{color:var(--amber);background:var(--amber-g);border-color:var(--amber);}
#saveInd.saved {color:var(--green);background:var(--green-g);border-color:var(--green);}
#saveInd.error {color:var(--red);background:var(--red-g);border-color:var(--red);}

/* ── Pills ───────────────────────────────────────────────── */
.pills{display:flex;border:1px solid var(--bd2);border-radius:6px;overflow:hidden;}
.pills button{
  background:transparent;border:none;color:var(--muted);
  font-family:var(--mono);font-size:11px;font-weight:700;
  padding:5px 12px;cursor:pointer;transition:all .14s;letter-spacing:.3px;
}
.pills button.on{background:var(--blue);color:#fff;}
.pills button:not(.on):hover{background:var(--s3);color:var(--text);}

/* ── Search ──────────────────────────────────────────────── */
.srch{position:relative;}
.srch svg{position:absolute;left:8px;top:50%;transform:translateY(-50%);
          color:var(--muted);pointer-events:none;}
.srch input{
  background:var(--s2);border:1px solid var(--bd2);border-radius:6px;
  color:var(--text);font-size:11px;font-family:var(--mono);
  padding:5px 10px 5px 26px;width:190px;outline:none;transition:all .15s;
}
.srch input:focus{border-color:var(--blue);background:var(--s3);width:220px;}

/* ── Main ────────────────────────────────────────────────── */
main{padding:20px 24px;max-width:1440px;margin:0 auto;}

/* ── Stat cards ──────────────────────────────────────────── */
.cards{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:18px;}
@media(max-width:800px){.cards{grid-template-columns:repeat(2,1fr);}}
.card{background:var(--s1);border:1px solid var(--bd);border-radius:var(--r);
      padding:16px 18px;position:relative;overflow:hidden;transition:transform .14s;}
.card:hover{transform:translateY(-1px);}
.card-line{position:absolute;top:0;left:0;right:0;height:2px;}
.cg .card-line{background:linear-gradient(90deg,var(--green),transparent);}
.ca .card-line{background:linear-gradient(90deg,var(--amber),transparent);}
.cr .card-line{background:linear-gradient(90deg,var(--red),transparent);}
.cb .card-line{background:linear-gradient(90deg,var(--blue),transparent);}
.cp .card-line{background:linear-gradient(90deg,var(--purple),transparent);}
.card .num{font-family:var(--mono);font-size:34px;font-weight:700;
           line-height:1;letter-spacing:-1px;margin-bottom:5px;}
.cg .num{color:var(--green);}.ca .num{color:var(--amber);}
.cr .num{color:var(--red);} .cb .num{color:var(--blue);}
.cp .num{color:var(--purple);}
.card .lbl{font-size:9px;text-transform:uppercase;letter-spacing:1px;
           font-weight:700;color:var(--muted);}
.card .sub{font-size:10px;color:var(--muted2);margin-top:3px;}

/* ── Charts ──────────────────────────────────────────────── */
.charts{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:18px;}
@media(max-width:580px){.charts{grid-template-columns:1fr;}}
.cbox{background:var(--s1);border:1px solid var(--bd);border-radius:var(--r);padding:16px 18px;}
.cbox h3{font-family:var(--mono);font-size:9px;color:var(--muted);
          letter-spacing:1.5px;text-transform:uppercase;margin-bottom:12px;}
.br{display:flex;align-items:center;gap:8px;margin-bottom:6px;}
.bk{width:86px;font-size:10px;color:var(--text);font-family:var(--mono);
    white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex-shrink:0;}
.bt{flex:1;background:var(--s3);border-radius:3px;height:5px;overflow:hidden;}
.bf{height:100%;border-radius:3px;transition:width .45s cubic-bezier(.4,0,.2,1);}
.bv{width:22px;text-align:right;font-size:10px;color:var(--muted);font-family:var(--mono);}
.ce{color:var(--muted);font-family:var(--mono);font-size:11px;text-align:center;padding:14px 0;}

/* ── Table ───────────────────────────────────────────────── */
.tw{background:var(--s1);border:1px solid var(--bd);border-radius:var(--r);overflow:hidden;}
.th{display:flex;align-items:center;gap:10px;padding:11px 16px;border-bottom:1px solid var(--bd);}
.th h2{font-family:var(--mono);font-size:10px;font-weight:700;
        letter-spacing:1.5px;text-transform:uppercase;color:var(--muted);}
.cnt{background:var(--s3);border:1px solid var(--bd2);border-radius:20px;
     padding:2px 8px;font-family:var(--mono);font-size:10px;color:var(--muted);}
table{width:100%;border-collapse:collapse;}
th{background:var(--s2);color:var(--muted);font-family:var(--mono);font-size:9px;
   font-weight:700;letter-spacing:1px;text-transform:uppercase;
   padding:9px 14px;border-bottom:1px solid var(--bd);
   cursor:pointer;user-select:none;white-space:nowrap;transition:color .14s;text-align:left;}
th:hover{color:var(--text);}
th.sorted{color:var(--blue);}
th .arr{opacity:.3;margin-left:3px;font-size:8px;}
th.sorted .arr{opacity:1;}
td{padding:9px 14px;border-bottom:1px solid var(--bd);
   vertical-align:middle;font-size:12px;transition:background .1s;}
tr:last-child td{border-bottom:none;}
tr:hover td{background:var(--s2);}
tr.editing-row td{background:var(--blue-g)!important;border-left:2px solid var(--blue);}

/* ── Badges ──────────────────────────────────────────────── */
.badge{display:inline-flex;align-items:center;gap:3px;padding:2px 8px;
       border-radius:20px;font-size:10px;font-family:var(--mono);
       font-weight:700;white-space:nowrap;}
.bg{background:var(--green-g);color:var(--green);border:1px solid #00d48a28;}
.ba{background:var(--amber-g);color:var(--amber);border:1px solid #f5a62328;}
.br2{background:var(--red-g);color:var(--red);border:1px solid #f03e5a28;}
.bb{background:var(--blue-g);color:var(--blue);border:1px solid #3d8ef028;}
.bm{background:var(--muted2);color:var(--muted);border:1px solid var(--bd2);}
.hp{display:inline-block;background:var(--s3);border:1px solid var(--bd2);
    border-radius:4px;padding:1px 6px;font-family:var(--mono);
    font-size:10px;margin:1px 2px 1px 0;color:var(--text);}
.nc{display:inline-block;background:var(--blue-g);border:1px solid #3d8ef025;
    border-radius:4px;padding:1px 7px;font-size:10px;color:var(--blue);
    font-family:var(--mono);max-width:180px;white-space:nowrap;
    overflow:hidden;text-overflow:ellipsis;cursor:help;vertical-align:middle;}

/* ── Empty ───────────────────────────────────────────────── */
.empty{padding:48px;text-align:center;color:var(--muted);
       font-family:var(--mono);font-size:11px;}

/* ── Pagination ──────────────────────────────────────────── */
.pag{display:flex;gap:4px;align-items:center;flex-wrap:wrap;
     padding:10px 14px;border-top:1px solid var(--bd);}
.pag button{background:var(--s2);border:1px solid var(--bd);color:var(--muted);
            border-radius:5px;padding:3px 9px;font-size:11px;cursor:pointer;
            font-family:var(--mono);transition:all .14s;}
.pag button.on{background:var(--blue);border-color:var(--blue);color:#fff;}
.pag button:disabled{opacity:.25;cursor:not-allowed;}
.pag button:not(.on):not(:disabled):hover{border-color:var(--blue);color:var(--blue);}
.pi{margin-left:auto;font-size:10px;color:var(--muted);font-family:var(--mono);}

/* ── Edit panel ──────────────────────────────────────────── */
#backdrop{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);
          z-index:200;backdrop-filter:blur(3px);}
#backdrop.open{display:block;}
#panel{
  position:fixed;right:0;top:0;bottom:0;width:var(--pw);max-width:96vw;
  background:var(--s1);border-left:1px solid var(--bd2);
  z-index:201;display:flex;flex-direction:column;
  transform:translateX(100%);transition:transform .22s cubic-bezier(.4,0,.2,1);
}
#panel.open{transform:translateX(0);}
.ph{padding:15px 20px;border-bottom:1px solid var(--bd);
    display:flex;align-items:flex-start;gap:10px;flex-shrink:0;}
.ph-left h2{font-family:var(--mono);font-size:10px;font-weight:700;
            letter-spacing:1.5px;text-transform:uppercase;color:var(--blue);margin-bottom:3px;}
.ph-left p{font-size:11px;color:var(--muted);word-break:break-all;line-height:1.5;}
.pb{flex:1;overflow-y:auto;padding:18px 20px;display:flex;flex-direction:column;gap:14px;}

/* field group */
.fg{display:flex;flex-direction:column;gap:5px;}
.fg label{font-family:var(--mono);font-size:9px;font-weight:700;
          text-transform:uppercase;letter-spacing:1px;color:var(--muted);}
.fg select,.fg textarea{
  background:var(--s2);border:1px solid var(--bd2);border-radius:6px;
  color:var(--text);font-family:var(--mono);font-size:12px;
  padding:8px 10px;outline:none;transition:border-color .14s,background .14s;width:100%;
}
.fg select:focus,.fg textarea:focus{border-color:var(--blue);background:var(--s3);}
.fg textarea{resize:vertical;min-height:80px;line-height:1.5;}
.fg select option{background:var(--s2);}

/* info boxes (read-only) */
.igrid{display:grid;grid-template-columns:1fr 1fr;gap:8px;}
.ibox{background:var(--s2);border:1px solid var(--bd);border-radius:6px;padding:8px 10px;}
.ibox .il{font-size:9px;color:var(--muted);text-transform:uppercase;
          letter-spacing:1px;font-family:var(--mono);margin-bottom:3px;}
.ibox .iv{font-size:11px;color:var(--text);font-family:var(--mono);
          word-break:break-all;line-height:1.4;}

/* host accordion */
.hcard{background:var(--s2);border:1px solid var(--bd);border-radius:6px;overflow:hidden;margin-bottom:4px;}
.hcard-h{padding:8px 10px;display:flex;align-items:center;gap:6px;cursor:pointer;transition:background .14s;}
.hcard-h:hover{background:var(--s3);}
.hcard-h .hn{font-family:var(--mono);font-size:11px;font-weight:700;flex:1;}
.hcard-b{border-top:1px solid var(--bd);padding:8px 10px;display:none;
         flex-direction:column;gap:4px;}
.hcard-b.open{display:flex;}
.tr{display:flex;align-items:center;gap:6px;font-family:var(--mono);font-size:10px;padding:2px 0;}
.tr .ti{width:14px;text-align:center;flex-shrink:0;}
.tr .tn{flex:1;color:var(--text);}
.tr .td{color:var(--muted);}

/* panel footer */
.pf{padding:13px 20px;border-top:1px solid var(--bd);flex-shrink:0;
    display:flex;gap:8px;align-items:center;}
.pss{font-family:var(--mono);font-size:10px;margin-left:auto;transition:color .2s;}
.pss.idle  {color:var(--muted);}
.pss.saving{color:var(--amber);}
.pss.saved {color:var(--green);}
.pss.error {color:var(--red);}

/* ── Delete modal ────────────────────────────────────────── */
#delmod{display:none;position:fixed;inset:0;z-index:300;
        background:rgba(0,0,0,.65);backdrop-filter:blur(4px);
        align-items:center;justify-content:center;}
#delmod.open{display:flex;}
.dbox{background:var(--s2);border:1px solid var(--bd2);border-radius:10px;
      padding:24px 26px;width:360px;max-width:92vw;
      animation:pop .15s cubic-bezier(.4,0,.2,1);}
@keyframes pop{from{opacity:0;transform:scale(.95);}to{opacity:1;transform:none;}}
.dbox h3{font-family:var(--mono);font-size:12px;color:var(--red);margin-bottom:6px;}
.dbox p{font-size:12px;color:var(--muted);margin-bottom:18px;line-height:1.5;}
.da{display:flex;gap:8px;justify-content:flex-end;}

/* ── Undo bar ────────────────────────────────────────────── */
#ubar{display:none;position:fixed;bottom:20px;left:50%;transform:translateX(-50%);
      background:var(--s2);border:1px solid var(--bd2);border-radius:8px;
      padding:9px 16px;z-index:300;align-items:center;gap:12px;
      font-family:var(--mono);font-size:11px;color:var(--text);
      box-shadow:0 4px 24px rgba(0,0,0,.5);}
#ubar.show{display:flex;}

/* ── Toasts ──────────────────────────────────────────────── */
#toasts{position:fixed;bottom:18px;right:18px;display:flex;flex-direction:column;
        gap:6px;z-index:400;}
.toast{background:var(--s2);border:1px solid var(--bd2);border-radius:7px;
       padding:8px 14px;font-family:var(--mono);font-size:11px;
       opacity:0;transform:translateX(8px);transition:all .18s;
       pointer-events:none;max-width:280px;}
.toast.show{opacity:1;transform:none;}
.tok {border-left:3px solid var(--green);color:var(--green);}
.terr{border-left:3px solid var(--red);color:var(--red);}
.twn {border-left:3px solid var(--amber);color:var(--amber);}
.tinf{border-left:3px solid var(--blue);color:var(--blue);}
</style>
</head>
<body>

<header>
  <div class="logo">
    <div class="pulse"></div>
    <div>
      <h1>SIEM Signoff</h1>
      <small id="lastEntry">Loading...</small>
    </div>
  </div>
  <div class="spacer"></div>
  <div class="hdr-r">
    <div class="pills" id="pills">
      <button class="on" onclick="setPeriod('7d',this)">7D</button>
      <button onclick="setPeriod('30d',this)">30D</button>
      <button onclick="setPeriod('90d',this)">90D</button>
    </div>
    <div class="srch">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none"
           stroke="currentColor" stroke-width="2.5">
        <circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/>
      </svg>
      <input type="text" id="q" placeholder="Host / sender / notes..." oninput="render()">
    </div>
    <button class="btn btn-amber" onclick="doExport()">&#x2B07; Export</button>
    <div id="saveInd" class="idle">&#x2714; Saved</div>
  </div>
</header>

<main>
  <div class="cards" id="cards"></div>
  <div class="charts">
    <div class="cbox"><h3>Status Breakdown</h3><div id="cStatus"></div></div>
    <div class="cbox"><h3>Top Hostnames</h3>  <div id="cHosts"></div></div>
  </div>
  <div class="tw">
    <div class="th">
      <h2>Signoff Log</h2>
      <span class="cnt" id="cnt">—</span>
      <div class="spacer"></div>
    </div>
    <table>
      <thead><tr>
        <th onclick="doSort('timestamp')"      id="th-timestamp">     Timestamp<span class="arr">&#x25BC;</span></th>
        <th>Hostnames</th>
        <th onclick="doSort('overall_status')" id="th-overall_status">Status<span class="arr">&#x25BC;</span></th>
        <th onclick="doSort('sender')"         id="th-sender">        Sender<span class="arr">&#x25BC;</span></th>
        <th>Notes</th>
        <th style="text-align:right;min-width:110px;">Actions</th>
      </tr></thead>
      <tbody id="tbody"></tbody>
    </table>
    <div class="pag" id="pag"></div>
  </div>
</main>

<!-- Edit panel -->
<div id="backdrop" onclick="closePanel()"></div>
<aside id="panel">
  <div class="ph">
    <div class="ph-left">
      <h2>Edit Record</h2>
      <p id="panel-hosts"></p>
    </div>
    <button class="btn btn-ghost" onclick="closePanel()" style="flex-shrink:0;">&#x2715; Close</button>
  </div>
  <div class="pb">
    <div class="igrid" id="pInfo"></div>
    <div class="fg">
      <label>Status</label>
      <select id="fStatus">
        <option value="active">Active</option>
        <option value="partial">Partial</option>
        <option value="not_found">Not Found</option>
      </select>
    </div>
    <div class="fg">
      <label>Resolved</label>
      <select id="fResolved">
        <option value="false">No — still open</option>
        <option value="true">Yes — resolved (skip future dedup)</option>
      </select>
    </div>
    <div class="fg">
      <label>Ticket / Notes &nbsp;<span style="color:var(--muted);font-weight:400;">(auto-saved as you type)</span></label>
      <textarea id="fNotes" placeholder="Ticket ID, remediation steps, context..."></textarea>
    </div>
    <div>
      <div style="font-family:var(--mono);font-size:9px;color:var(--muted);
                  text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">
        QRadar Host Detail
      </div>
      <div id="pDetail"></div>
    </div>
    <div class="igrid" id="pMeta"></div>
  </div>
  <div class="pf">
    <button class="btn btn-green" onclick="applyEdit()">&#x2713; Apply Changes</button>
    <button class="btn btn-ghost" onclick="closePanel()">Cancel</button>
    <span class="pss idle" id="pss"></span>
  </div>
</aside>

<!-- Delete confirm -->
<div id="delmod">
  <div class="dbox">
    <h3>&#x26A0; Confirm Delete</h3>
    <p id="delMsg"></p>
    <div class="da">
      <button class="btn" onclick="closeDel()">Cancel</button>
      <button class="btn btn-red" id="delOk">Delete</button>
    </div>
  </div>
</div>

<!-- Undo bar -->
<div id="ubar">
  <span id="uMsg"></span>
  <button class="btn btn-amber" onclick="doUndo()"
          style="padding:3px 10px;font-size:10px;">&#x21A9; Undo</button>
</div>

<!-- Toasts -->
<div id="toasts"></div>

<script>
/* ═══════════════════════════════════════════════════════════
   STATE
   ═══════════════════════════════════════════════════════════ */
let D          = null;
let _undoSnap  = null;
let _undoLabel = '';
let _undoTimer = null;
let _sf        = 'timestamp';
let _sasc      = false;
let _page      = 1;
const PS       = 15;
let _period    = '7d';
let _editIdx   = null;
let _delIdx    = null;
let _saveTimer = null;
let _autoTimer = null;

/* ═══════════════════════════════════════════════════════════
   BOOT
   ═══════════════════════════════════════════════════════════ */
async function init() {
  try {
    const r = await fetch('/data.json',{cache:'no-cache'});
    if (!r.ok) throw new Error('HTTP '+r.status);
    D = await r.json();
    if (!Array.isArray(D.entries)) D.entries=[];
  } catch(e) {
    D = {schema_version:3,entries:[]};
    toast('Could not load data.json — '+e.message,'err');
  }
  refreshLastEntry();
  render();
}
function clone(o){return JSON.parse(JSON.stringify(o));}
function refreshLastEntry(){
  const el = document.getElementById('lastEntry');
  const entries = D.entries||[];
  if (!entries.length){el.textContent='No data';return;}
  const last = [...entries].sort((a,b)=>(b.timestamp||'').localeCompare(a.timestamp||''))[0];
  el.textContent = 'Last entry: '+fmtShort(last.timestamp);
}

/* ═══════════════════════════════════════════════════════════
   PERSIST — atomic write via /save, debounced
   ═══════════════════════════════════════════════════════════ */
function persist(label, delay=400){
  clearTimeout(_saveTimer);
  _saveTimer = setTimeout(()=>_doSave(label), delay);
}
async function _doSave(label){
  setInd('saving','&#x21BA; Saving...');
  try {
    const r = await fetch('/save',{
      method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify(D),
    });
    if (!r.ok) throw new Error('HTTP '+r.status);
    const res = await r.json();
    if (!res.ok) throw new Error(res.error||'unknown');
    setInd('saved','&#x2714; Saved');
    if (label) toast(label,'ok');
    setTimeout(()=>setInd('idle','&#x2714; Saved'),2200);
  } catch(e){
    setInd('error','&#x2715; Save failed');
    toast('Save failed: '+e.message+' — use Export as backup','err');
  }
}
function setInd(cls,html){
  const el=document.getElementById('saveInd');
  el.className=cls; el.innerHTML=html;
}

/* ═══════════════════════════════════════════════════════════
   UNDO (single level, in-memory only)
   ═══════════════════════════════════════════════════════════ */
function snap(label){
  _undoSnap=clone(D); _undoLabel=label;
  clearTimeout(_undoTimer);
  document.getElementById('uMsg').textContent=label+' — ';
  document.getElementById('ubar').classList.add('show');
  _undoTimer=setTimeout(hideUndo,8000);
}
function hideUndo(){
  document.getElementById('ubar').classList.remove('show');
  _undoSnap=null;
}
function doUndo(){
  if (!_undoSnap) return;
  D=_undoSnap; _undoSnap=null;
  hideUndo(); closePanel();
  render(); refreshLastEntry();
  persist(null,0);
  toast('Undone: '+_undoLabel,'warn');
}

/* ═══════════════════════════════════════════════════════════
   FILTERING / SORTING
   ═══════════════════════════════════════════════════════════ */
function cutoff(){
  const days={'7d':7,'30d':30,'90d':90}[_period]||7;
  const d=new Date(); d.setDate(d.getDate()-days); return d;
}
function filtered(){
  const q=(document.getElementById('q').value||'').toLowerCase().trim();
  const co=cutoff();
  return (D.entries||[]).filter(e=>{
    if (new Date(e.timestamp)<co) return false;
    if (!q) return true;
    const hosts=(e.hosts||[]).map(h=>h.hostname||'').join(' ').toLowerCase();
    return [e.sender||'',e.notes||'',e.email_subject||'',hosts]
            .some(s=>s.toLowerCase().includes(q));
  });
}
function sorted(arr){
  return [...arr].sort((a,b)=>{
    const av=a[_sf]||'',bv=b[_sf]||'';
    return av<bv?(_sasc?-1:1):av>bv?(_sasc?1:-1):0;
  });
}

/* ═══════════════════════════════════════════════════════════
   RENDER
   ═══════════════════════════════════════════════════════════ */
function render(){
  const all=sorted(filtered());
  const total=all.length;
  const pages=Math.max(1,Math.ceil(total/PS));
  if (_page>pages) _page=1;
  const slice=all.slice((_page-1)*PS,_page*PS);

  document.getElementById('cnt').textContent=
    total+' record'+(total!==1?'s':'');

  // Sort indicators
  ['timestamp','overall_status','sender'].forEach(f=>{
    const th=document.getElementById('th-'+f);
    if (!th) return;
    th.classList.toggle('sorted',_sf===f);
    th.querySelector('.arr').textContent=(_sf===f&&_sasc)?'▲':'▼';
  });

  // Rows
  const tbody=document.getElementById('tbody');
  if (!slice.length){
    tbody.innerHTML='<tr><td colspan="6"><div class="empty">No records match this filter.</div></td></tr>';
  } else {
    tbody.innerHTML=slice.map((x,i)=>{
      const gi=(D.entries||[]).indexOf(all[(_page-1)*PS+i]);
      const hosts=(x.hosts||[]).map(h=>
        `<span class="hp" title="${esc(h.hostname||'')}">${esc(h.hostname||'?')}</span>`
      ).join('');
      const note=x.notes
        ?`<span class="nc" title="${esc(x.notes)}">${esc(x.notes)}</span>`
        :`<span style="color:var(--muted2)">—</span>`;
      const rowCls=gi===_editIdx?' class="editing-row"':'';
      return `<tr${rowCls}>
        <td style="font-family:var(--mono);font-size:11px;white-space:nowrap;">${fmtDate(x.timestamp)}</td>
        <td>${hosts}</td>
        <td>${sbadge(x.overall_status,x.manually_resolved,x.is_revalidation)}</td>
        <td style="white-space:nowrap;">${fmtSender(x.sender)}</td>
        <td>${note}</td>
        <td style="text-align:right;white-space:nowrap;">
          <button class="btn" onclick="openPanel(${gi})"
            style="padding:4px 10px;font-size:10px;">&#x270E; Edit</button>
          <button class="btn btn-red" onclick="confirmDel(${gi})"
            style="padding:4px 10px;font-size:10px;" title="Delete">&#x2715;</button>
        </td>
      </tr>`;
    }).join('');
  }

  // Pagination
  let ph=`<button onclick="goP(${_page-1})" ${_page===1?'disabled':''}>&#x2039;</button>`;
  const s=Math.max(1,_page-2),e=Math.min(pages,_page+2);
  if(s>1) ph+=`<button onclick="goP(1)">1</button>${s>2?'<span style="color:var(--muted);padding:0 2px">&#x2026;</span>':''}`;
  for(let p=s;p<=e;p++) ph+=`<button onclick="goP(${p})" class="${p===_page?'on':''}">${p}</button>`;
  if(e<pages) ph+=`${e<pages-1?'<span style="color:var(--muted);padding:0 2px">&#x2026;</span>':''}<button onclick="goP(${pages})">${pages}</button>`;
  ph+=`<button onclick="goP(${_page+1})" ${_page===pages?'disabled':''}>&#x203A;</button>
       <span class="pi">${total} record${total!==1?'s':''} &middot; page ${_page}/${pages}</span>`;
  document.getElementById('pag').innerHTML=ph;

  renderCards(filtered());
  renderCharts(filtered());
}
function goP(p){_page=p;render();window.scrollTo({top:0,behavior:'smooth'});}
function doSort(f){_sf===f?_sasc=!_sasc:(_sf=f,_sasc=false);render();}
function setPeriod(p,btn){
  _period=p;_page=1;
  document.querySelectorAll('.pills button').forEach(b=>b.classList.remove('on'));
  btn.classList.add('on');render();
}

/* ═══════════════════════════════════════════════════════════
   STAT CARDS
   ═══════════════════════════════════════════════════════════ */
function renderCards(e){
  const tot=e.length;
  const act=e.filter(x=>x.overall_status==='active'    &&!x.manually_resolved).length;
  const par=e.filter(x=>x.overall_status==='partial'   &&!x.manually_resolved).length;
  const nf =e.filter(x=>x.overall_status==='not_found' &&!x.manually_resolved).length;
  const res=e.filter(x=>x.manually_resolved).length;
  const pct=tot?Math.round(act/tot*100):0;
  document.getElementById('cards').innerHTML=`
    <div class="card cb"><div class="card-line"></div>
      <div class="num">${tot}</div><div class="lbl">Total</div><div class="sub">in period</div></div>
    <div class="card cg"><div class="card-line"></div>
      <div class="num">${act}</div><div class="lbl">Active</div><div class="sub">${pct}% success rate</div></div>
    <div class="card ca"><div class="card-line"></div>
      <div class="num">${par}</div><div class="lbl">Partial</div><div class="sub">missing sources</div></div>
    <div class="card cr"><div class="card-line"></div>
      <div class="num">${nf}</div><div class="lbl">Not Found</div><div class="sub">not onboarded</div></div>
    <div class="card cp"><div class="card-line"></div>
      <div class="num">${res}</div><div class="lbl">Resolved</div><div class="sub">manually closed</div></div>`;
}

/* ═══════════════════════════════════════════════════════════
   CHARTS
   ═══════════════════════════════════════════════════════════ */
function renderCharts(e){
  const sm={};
  e.forEach(x=>{
    const k=x.manually_resolved?'resolved':(x.overall_status||'unknown');
    sm[k]=(sm[k]||0)+1;
  });
  const clr={active:'var(--green)',partial:'var(--amber)',not_found:'var(--red)',
             resolved:'var(--muted)',unknown:'var(--bd2)'};
  barChart(document.getElementById('cStatus'),
    Object.entries(sm).map(([k,v])=>({k,v})).sort((a,b)=>b.v-a.v),
    k=>clr[k]||'var(--blue)');

  const hm={};
  e.forEach(x=>(x.hosts||[]).forEach(h=>{
    if(h.hostname)hm[h.hostname]=(hm[h.hostname]||0)+1;
  }));
  barChart(document.getElementById('cHosts'),
    Object.entries(hm).map(([k,v])=>({k,v})).sort((a,b)=>b.v-a.v).slice(0,8),
    ()=>'var(--blue)');
}
function barChart(el,items,colorFn){
  if(!items.length){el.innerHTML='<div class="ce">No data</div>';return;}
  const mx=Math.max(...items.map(i=>i.v),1);
  el.innerHTML=items.map(i=>`
    <div class="br">
      <span class="bk" title="${esc(i.k)}">${esc(i.k)}</span>
      <div class="bt"><div class="bf" style="width:${Math.round(i.v/mx*100)}%;background:${colorFn(i.k)}"></div></div>
      <span class="bv">${i.v}</span>
    </div>`).join('');
}

/* ═══════════════════════════════════════════════════════════
   EDIT PANEL
   ═══════════════════════════════════════════════════════════ */
function openPanel(gi){
  _editIdx=gi;
  const e=D.entries[gi];
  if (!e) return;

  // Header
  const hostNames=(e.hosts||[]).map(h=>h.hostname||'?').join(' · ');
  document.getElementById('panel-hosts').textContent=hostNames||e.email_subject||'—';

  // Info grid (read-only)
  document.getElementById('pInfo').innerHTML=`
    <div class="ibox">
      <div class="il">Received</div>
      <div class="iv">${fmtFull(e.timestamp)}</div>
    </div>
    <div class="ibox">
      <div class="il">Run Type</div>
      <div class="iv">${e.is_revalidation?'&#x267B; Revalidation':'&#x2726; New signoff'}</div>
    </div>`;

  // Editable fields
  document.getElementById('fStatus').value  =e.overall_status||'active';
  document.getElementById('fResolved').value=e.manually_resolved?'true':'false';
  document.getElementById('fNotes').value   =e.notes||'';

  // Host detail
  const det=document.getElementById('pDetail');
  det.innerHTML=(e.hosts||[]).map((h,idx)=>{
    const hbadge=hsBadge(h.status);
    const types=(h.type_results||[]).map(tr=>{
      const icon=tr.found
        ?(tr.days_ago===null
          ?'<span style="color:var(--amber)">&#x26A0;</span>'
          :'<span style="color:var(--green)">&#x2714;</span>')
        :'<span style="color:var(--red)">&#x2716;</span>';
      const dstr=tr.found&&tr.days_ago!==null&&tr.days_ago!==undefined
        ?(tr.days_ago===0?'today':tr.days_ago+'d ago')
        :(tr.found?'no events':'missing');
      return `<div class="tr">
        <span class="ti">${icon}</span>
        <span class="tn">${esc(tr.expected||'')}</span>
        <span class="td">${dstr}</span>
      </div>`;
    }).join('');
    const id='hcb'+idx;
    return `<div class="hcard">
      <div class="hcard-h" onclick="toggleHc('${id}')">
        <span class="hn">${esc(h.hostname||'?')}</span>
        ${hbadge}
        <span style="color:var(--muted);font-size:10px;margin-left:4px;">&#x25BE;</span>
      </div>
      <div class="hcard-b" id="${id}">
        ${types||'<span style="color:var(--muted);font-size:10px;">No type detail recorded</span>'}
      </div>
    </div>`;
  }).join('');

  // Meta (read-only)
  document.getElementById('pMeta').innerHTML=`
    <div class="ibox" style="grid-column:1/-1;">
      <div class="il">Sender</div>
      <div class="iv">${esc(e.sender||'—')}</div>
    </div>
    <div class="ibox" style="grid-column:1/-1;">
      <div class="il">Subject</div>
      <div class="iv">${esc(e.email_subject||'—')}</div>
    </div>`;

  setPss('idle','');
  document.getElementById('backdrop').classList.add('open');
  document.getElementById('panel').classList.add('open');
  render();
}
function closePanel(){
  _editIdx=null;
  document.getElementById('backdrop').classList.remove('open');
  document.getElementById('panel').classList.remove('open');
  render();
}
function toggleHc(id){
  const el=document.getElementById(id);
  if(el) el.classList.toggle('open');
}

// Apply button: status + resolved override, then persist
function applyEdit(){
  if (_editIdx===null) return;
  snap('Edit record');
  const e=D.entries[_editIdx];
  e.overall_status   =document.getElementById('fStatus').value;
  e.manually_resolved=document.getElementById('fResolved').value==='true';
  e.notes            =document.getElementById('fNotes').value.trim();
  e.last_edited      =new Date().toISOString();
  setPss('saving','&#x21BA; Saving...');
  persist('Record updated',0);
  render();
  toast('Changes applied and saved','ok');
  setTimeout(()=>setPss('saved','&#x2714; Saved to disk'),500);
  setTimeout(()=>setPss('idle',''),2500);
}

// Auto-save notes as user types (debounced, no undo snap for typing)
document.getElementById('fNotes').addEventListener('input',()=>{
  if (_editIdx===null) return;
  D.entries[_editIdx].notes=document.getElementById('fNotes').value.trim();
  setPss('saving','&#x21BA; Auto-saving...');
  clearTimeout(_autoTimer);
  _autoTimer=setTimeout(()=>{
    persist(null,0);
    setPss('saved','&#x2714; Auto-saved');
    setTimeout(()=>setPss('idle',''),2000);
  },600);
});

function setPss(cls,html){
  const el=document.getElementById('pss');
  el.className='pss '+cls; el.innerHTML=html;
}

/* ═══════════════════════════════════════════════════════════
   DELETE
   ═══════════════════════════════════════════════════════════ */
function confirmDel(gi){
  _delIdx=gi;
  const e=(D.entries||[])[gi];
  if (!e) return;
  const label=(e.hosts||[]).map(h=>h.hostname||'?').join(', ')||e.email_subject||'this record';
  document.getElementById('delMsg').textContent=
    `Permanently delete the record for "${label}"? `+
    `This writes immediately to signoff_data.json. You can undo within 8 seconds.`;
  document.getElementById('delOk').onclick=execDel;
  document.getElementById('delmod').classList.add('open');
}
function closeDel(){
  _delIdx=null;
  document.getElementById('delmod').classList.remove('open');
}
function execDel(){
  if (_delIdx===null) return;
  const e=D.entries[_delIdx];
  const label=(e.hosts||[]).map(h=>h.hostname||'?').join(', ')||'record';
  snap('Delete: '+label);
  D.entries.splice(_delIdx,1);
  if (_editIdx===_delIdx) closePanel();
  else if (_editIdx!==null && _editIdx>_delIdx) _editIdx--;
  _delIdx=null;
  closeDel();
  const pages=Math.max(1,Math.ceil(filtered().length/PS));
  if (_page>pages) _page=pages;
  render(); refreshLastEntry();
  persist(null,0);
  toast('Deleted — '+label+' (Ctrl+Z to undo)','warn');
}

/* ═══════════════════════════════════════════════════════════
   FORMAT HELPERS
   ═══════════════════════════════════════════════════════════ */
function sbadge(status,resolved,reval){
  if (resolved) return '<span class="badge bm">&#x2714; Resolved</span>';
  const m={active:['bg','&#x2714;','Active'],partial:['ba','&#x26A0;','Partial'],not_found:['br2','&#x2716;','Not Found']};
  const [c,ic,lb]=m[status]||['bm','?',status||'—'];
  const rv=reval?` <span class="badge bb" style="margin-left:3px;font-size:9px;">&#x21BB; Reval</span>`:'';
  return `<span class="badge ${c}">${ic} ${lb}</span>${rv}`;
}
function hsBadge(s){
  const m={active:'bg',partial:'ba',not_found:'br2'};
  return `<span class="badge ${m[s]||'bm'}" style="font-size:9px;">${s||'—'}</span>`;
}
function fmtDate(iso){
  if(!iso)return'—';
  const d=new Date(iso);
  const dt=d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});
  const t=d.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'});
  return `${dt} <span style="color:var(--muted)">${t}</span>`;
}
function fmtShort(iso){
  if(!iso)return'—';
  return new Date(iso).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});
}
function fmtFull(iso){
  if(!iso)return'—';
  const d=new Date(iso);
  return d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'})
       +' '+d.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit',second:'2-digit'});
}
function fmtSender(s){
  if(!s)return'<span style="color:var(--muted2)">—</span>';
  const p=s.split('@');
  if(p.length<2)return`<span style="font-family:var(--mono);font-size:11px">${esc(s)}</span>`;
  return`<span style="font-family:var(--mono);font-size:11px">${esc(p[0])}</span>`
       +`<span style="font-family:var(--mono);font-size:10px;color:var(--muted)">@${esc(p[1])}</span>`;
}
function esc(s){
  return String(s||'')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/* ═══════════════════════════════════════════════════════════
   EXPORT
   ═══════════════════════════════════════════════════════════ */
function doExport(){
  const b=new Blob([JSON.stringify(D,null,2)],{type:'application/json'});
  Object.assign(document.createElement('a'),{
    href:URL.createObjectURL(b),download:'signoff_data_export.json'
  }).click();
  toast('Exported signoff_data_export.json','info');
}

/* ═══════════════════════════════════════════════════════════
   TOASTS
   ═══════════════════════════════════════════════════════════ */
function toast(msg,type='ok'){
  const el=document.createElement('div');
  el.className=`toast t${type[0]=='e'?'err':type[0]=='w'?'wn':type[0]=='i'?'inf':'ok'}`;
  el.textContent=msg;
  document.getElementById('toasts').appendChild(el);
  requestAnimationFrame(()=>el.classList.add('show'));
  setTimeout(()=>{el.classList.remove('show');setTimeout(()=>el.remove(),200);},3800);
}

/* ═══════════════════════════════════════════════════════════
   KEYBOARD
   ═══════════════════════════════════════════════════════════ */
document.addEventListener('keydown',e=>{
  if (e.key==='Escape'){closePanel();closeDel();}
  if ((e.ctrlKey||e.metaKey)&&e.key==='z'&&_undoSnap){e.preventDefault();doUndo();}
});

/* boot */
init();
</script>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════════════════════
# SERVER
# ══════════════════════════════════════════════════════════════════════════════

def run(port: int, minutes: int) -> None:
    port   = _find_free_port(port)
    server = HTTPServer(('127.0.0.1', port), Handler)
    url    = f'http://127.0.0.1:{port}/'

    if not os.path.exists(SIGNOFF_DATA_PATH):
        _log(f"WARNING: {SIGNOFF_DATA_PATH} not found — run signoff_runner.py first.")
    else:
        try:
            with open(SIGNOFF_DATA_PATH, 'r', encoding='utf-8') as f:
                data = json.load(f)
            _log(f"Data      : {SIGNOFF_DATA_PATH} ({len(data.get('entries',[]))} entries)")
        except Exception as exc:
            _log(f"WARNING   : Could not parse data file: {exc}")

    _log(f"Dashboard : {url}")
    _log("Persist   : every edit/delete auto-writes to signoff_data.json (atomic)")
    _log("Shortcuts : Esc — close panel | Ctrl+Z — undo last action")
    _log("Running until Ctrl+C." if minutes == 0 else f"Shutting down after {minutes} min.")

    t = threading.Thread(target=server.serve_forever, daemon=True)
    t.start()

    def _open():
        time.sleep(0.5)
        try: webbrowser.open(url)
        except Exception as exc: _log(f"WARNING: could not open browser: {exc}")

    threading.Thread(target=_open, daemon=True).start()

    try:
        if minutes > 0:
            time.sleep(minutes * 60)
            _log(f"Timeout ({minutes} min) — shutting down.")
        else:
            while True:
                time.sleep(3600)
    except KeyboardInterrupt:
        _log("\nCtrl+C — shutting down.")
    finally:
        server.shutdown()


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    ap = argparse.ArgumentParser(
        description='SIEM Signoff Dashboard — auto-persists every change to signoff_data.json.'
    )
    ap.add_argument('--port', type=int, default=DEFAULT_PORT,
                    help=f'Preferred local port (default: {DEFAULT_PORT})')
    ap.add_argument('--minutes', type=int, default=DEFAULT_MINUTES,
                    help='Runtime in minutes; 0 = run until Ctrl+C (default: 0)')
    args = ap.parse_args()
    run(port=args.port, minutes=args.minutes)
