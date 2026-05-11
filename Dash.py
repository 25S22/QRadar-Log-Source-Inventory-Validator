"""
SIEM Signoff Dashboard  v3.0
────────────────────────────────────────────────────────────────────────────
Standalone viewer for signoff_data.json produced by signoff_runner.py.
Run this script independently at any time — the runner does not need to be
running concurrently.

Usage:
    python signoff_dashboard.py
    python signoff_dashboard.py --port 9000 --minutes 120
    python signoff_dashboard.py --minutes 0     # run until Ctrl+C

The dashboard opens automatically in your default browser.
Edits (notes, status overrides, resolved flags) are saved back to
signoff_data.json via the in-page Save button.
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

# ─── Config ───────────────────────────────────────────────────────────────────
_DIR              = os.path.dirname(os.path.abspath(__file__))
SIGNOFF_DATA_PATH = os.path.join(_DIR, 'signoff_data.json')
DEFAULT_PORT      = 8745
DEFAULT_MINUTES   = 60   # 0 = run until Ctrl+C


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _ts() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _log(msg: str) -> None:
    print(f"[{_ts()}] {msg}")


def _atomic_write(path: str, data: dict) -> None:
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

class DashboardHandler(BaseHTTPRequestHandler):

    def do_GET(self):
        if self.path in ('/', '/index.html'):
            body = DASHBOARD_HTML.encode('utf-8')
            self._send(200, 'text/html; charset=utf-8', body)

        elif self.path == '/data.json':
            try:
                with open(SIGNOFF_DATA_PATH, 'rb') as f:
                    body = f.read()
                self._send(200, 'application/json', body)
            except FileNotFoundError:
                self._send(404, 'text/plain', b'signoff_data.json not found')

        else:
            self._send(404, 'text/plain', b'not found')

    def do_POST(self):
        if self.path == '/save':
            try:
                length = int(self.headers.get('Content-Length', 0))
                body   = self.rfile.read(length)
                data   = json.loads(body.decode('utf-8'))

                # Basic schema guard — must have an 'entries' list
                if not isinstance(data.get('entries'), list):
                    raise ValueError("Payload missing 'entries' list")

                _atomic_write(SIGNOFF_DATA_PATH, data)
                n = len(data['entries'])
                _log(f"[Save] {n} entries written to {SIGNOFF_DATA_PATH}")
                self._json(200, {'ok': True, 'entries': n})
            except Exception as exc:
                _log(f"[Save ERROR] {exc}")
                self._json(500, {'error': str(exc)})
        else:
            self._send(404, 'text/plain', b'not found')

    def do_OPTIONS(self):
        self.send_response(200)
        for k, v in [
            ('Access-Control-Allow-Origin',  '*'),
            ('Access-Control-Allow-Methods', 'GET, POST, OPTIONS'),
            ('Access-Control-Allow-Headers', 'Content-Type'),
        ]:
            self.send_header(k, v)
        self.end_headers()

    # ── Internal helpers ──────────────────────────────────────────────────────

    def _send(self, code: int, ctype: str, body: bytes) -> None:
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-cache')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code: int, obj: dict) -> None:
        self._send(code, 'application/json', json.dumps(obj).encode())

    def log_message(self, fmt, *args):
        pass  # suppress per-request console noise


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD HTML  (embedded — no external files required)
# ══════════════════════════════════════════════════════════════════════════════

DASHBOARD_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SIEM Signoff Dashboard</title>
<style>
/* ── Reset & variables ───────────────────────────────────────── */
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
:root{
  --bg:#0b0d12;
  --surface:#12151c;
  --surface2:#181c26;
  --surface3:#1e2330;
  --border:#222840;
  --border2:#2c3350;
  --green:#00c27a;   --green-d:#00c27a22;
  --amber:#f0a030;   --amber-d:#f0a03022;
  --red:#f04060;     --red-d:#f0406022;
  --blue:#4a90e8;    --blue-d:#4a90e822;
  --purple:#9b6cf0;  --purple-d:#9b6cf022;
  --text:#ccd3e8;
  --muted:#4e5a7a;
  --muted2:#2e3650;
  --mono:'Courier New',Courier,monospace;
  --sans:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
  --r:8px;
}

html{scroll-behavior:smooth;}
body{
  background:var(--bg);color:var(--text);
  font-family:var(--sans);font-size:13px;
  min-height:100vh;overflow-x:hidden;
}

::-webkit-scrollbar{width:5px;}
::-webkit-scrollbar-track{background:var(--bg);}
::-webkit-scrollbar-thumb{background:var(--border2);border-radius:3px;}

/* ── Dirty banner ───────────────────────────────────────────── */
#dirtyBanner{
  display:none;
  background:linear-gradient(90deg,#1a1100,#110800);
  border-bottom:1px solid var(--amber);
  padding:7px 28px;gap:12px;align-items:center;
  font-family:var(--mono);font-size:11px;color:var(--amber);
}
#dirtyBanner.show{display:flex;}

/* ── Header ─────────────────────────────────────────────────── */
header{
  background:var(--surface);
  border-bottom:1px solid var(--border);
  padding:0 28px;height:58px;
  display:flex;align-items:center;gap:16px;
  position:sticky;top:0;z-index:200;
  backdrop-filter:blur(10px);
}
.logo{display:flex;align-items:center;gap:10px;flex-shrink:0;}
.logo-pulse{
  width:7px;height:7px;border-radius:50%;
  background:var(--green);box-shadow:0 0 6px var(--green);
  animation:pulse 2.6s ease-in-out infinite;
}
@keyframes pulse{0%,100%{opacity:1;}50%{opacity:.3;}}
.logo h1{
  font-family:var(--mono);font-size:11px;font-weight:700;
  letter-spacing:2px;text-transform:uppercase;color:var(--text);
}
.logo-sub{
  font-family:var(--mono);font-size:9px;color:var(--muted);
  letter-spacing:1px;margin-top:2px;
}
.spacer{flex:1;}
.hdr-right{display:flex;gap:8px;align-items:center;flex-wrap:wrap;}

/* ── Buttons ─────────────────────────────────────────────────── */
.btn{
  background:transparent;border:1px solid var(--border2);
  color:var(--muted);border-radius:6px;
  font-size:11px;padding:5px 13px;cursor:pointer;
  font-family:var(--mono);font-weight:600;letter-spacing:.3px;
  transition:all .15s ease;white-space:nowrap;
}
.btn:hover:not(:disabled){border-color:var(--blue);color:var(--blue);background:var(--blue-d);}
.btn:disabled{opacity:.3;cursor:not-allowed;}

.btn-save{border-color:var(--green);color:var(--green);}
.btn-save:hover:not(:disabled){background:var(--green);color:#000;box-shadow:0 0 10px var(--green-d);}
.btn-save.dirty{
  background:var(--green-d);animation:savepulse 1.8s ease-in-out infinite;
}
@keyframes savepulse{0%,100%{box-shadow:0 0 0 0 var(--green-d);}
  50%{box-shadow:0 0 0 4px transparent;}}

.btn-discard{border-color:var(--red);color:var(--red);}
.btn-discard:hover:not(:disabled){background:var(--red-d);}
.btn-danger{border-color:var(--red-d);color:var(--red);}
.btn-danger:hover{background:var(--red);color:#fff;border-color:var(--red);}
.btn-amber{border-color:var(--amber-d);color:var(--amber);}
.btn-amber:hover{background:var(--amber-d);border-color:var(--amber);}

/* ── Period pills ───────────────────────────────────────────── */
.pills{
  display:flex;border:1px solid var(--border2);
  border-radius:6px;overflow:hidden;
}
.pills button{
  background:transparent;border:none;color:var(--muted);
  font-family:var(--mono);font-size:11px;font-weight:600;
  padding:5px 13px;cursor:pointer;transition:all .15s;letter-spacing:.3px;
}
.pills button.active{background:var(--blue);color:#fff;}
.pills button:not(.active):hover{background:var(--surface3);color:var(--text);}

/* ── Search ─────────────────────────────────────────────────── */
.search-wrap{position:relative;}
.search-wrap svg{
  position:absolute;left:9px;top:50%;transform:translateY(-50%);
  color:var(--muted);pointer-events:none;
}
input[type=text]{
  background:var(--surface2);border:1px solid var(--border2);
  border-radius:6px;color:var(--text);font-size:11px;
  padding:5px 11px 5px 28px;width:200px;outline:none;
  transition:all .15s;font-family:var(--mono);
}
input[type=text]:focus{border-color:var(--blue);background:var(--surface3);width:230px;}

/* ── Main ────────────────────────────────────────────────────── */
main{padding:24px 28px;max-width:1400px;margin:0 auto;}

/* ── Stat cards ──────────────────────────────────────────────── */
.stats{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:22px;}
@media(max-width:860px){.stats{grid-template-columns:repeat(2,1fr);}}
.card{
  background:var(--surface);border:1px solid var(--border);
  border-radius:var(--r);padding:18px 20px;position:relative;overflow:hidden;
  transition:transform .15s;
}
.card:hover{transform:translateY(-1px);}
.card-bar{position:absolute;top:0;left:0;right:0;height:2px;}
.g .card-bar{background:linear-gradient(90deg,var(--green),transparent);}
.a .card-bar{background:linear-gradient(90deg,var(--amber),transparent);}
.r .card-bar{background:linear-gradient(90deg,var(--red),transparent);}
.b .card-bar{background:linear-gradient(90deg,var(--blue),transparent);}
.p .card-bar{background:linear-gradient(90deg,var(--purple),transparent);}
.card .num{
  font-family:var(--mono);font-size:36px;font-weight:700;
  line-height:1;margin-bottom:6px;letter-spacing:-1px;
}
.g .num{color:var(--green);} .a .num{color:var(--amber);}
.r .num{color:var(--red);}   .b .num{color:var(--blue);}
.p .num{color:var(--purple);}
.card .clabel{
  color:var(--muted);font-size:10px;letter-spacing:1px;
  text-transform:uppercase;font-weight:600;
}
.card .sub{color:var(--muted2);font-size:10px;margin-top:4px;}

/* ── Chart bar ───────────────────────────────────────────────── */
.charts{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:22px;}
@media(max-width:640px){.charts{grid-template-columns:1fr;}}
.chart-box{
  background:var(--surface);border:1px solid var(--border);
  border-radius:var(--r);padding:18px 20px;
}
.chart-box h3{
  font-family:var(--mono);font-size:9px;color:var(--muted);
  letter-spacing:1.5px;text-transform:uppercase;margin-bottom:14px;
}
.bar-row{display:flex;align-items:center;gap:8px;margin-bottom:7px;}
.bar-key{width:90px;font-size:10px;color:var(--text);font-family:var(--mono);
         white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.bar-track{flex:1;background:var(--surface3);border-radius:3px;height:5px;overflow:hidden;}
.bar-fill{height:100%;border-radius:3px;transition:width .5s cubic-bezier(.4,0,.2,1);}
.bar-val{width:22px;text-align:right;font-size:10px;color:var(--muted);font-family:var(--mono);}
.no-data{
  color:var(--muted);font-family:var(--mono);font-size:11px;
  text-align:center;padding:18px 0;
}

/* ── Table ───────────────────────────────────────────────────── */
.table-wrap{
  background:var(--surface);border:1px solid var(--border);
  border-radius:var(--r);overflow:hidden;
}
.table-hdr{
  padding:12px 18px;border-bottom:1px solid var(--border);
  display:flex;align-items:center;gap:10px;
}
.table-hdr h2{
  font-family:var(--mono);font-size:10px;font-weight:700;
  letter-spacing:1.5px;text-transform:uppercase;color:var(--muted);
}
.rec-badge{
  background:var(--surface3);border:1px solid var(--border2);
  border-radius:20px;padding:2px 8px;
  font-family:var(--mono);font-size:10px;color:var(--muted);
}
table{width:100%;border-collapse:collapse;}
th{
  background:var(--surface2);color:var(--muted);
  font-family:var(--mono);font-size:9px;font-weight:700;
  letter-spacing:1px;text-align:left;text-transform:uppercase;
  padding:9px 14px;border-bottom:1px solid var(--border);
  cursor:pointer;user-select:none;white-space:nowrap;transition:color .15s;
}
th:hover{color:var(--text);}
th.sorted{color:var(--blue);}
th .arr{opacity:.3;margin-left:3px;font-size:8px;}
th.sorted .arr{opacity:1;}
td{
  padding:10px 14px;border-bottom:1px solid var(--border);
  vertical-align:middle;font-size:12px;transition:background .1s;
}
tr:last-child td{border-bottom:none;}
tr:hover td{background:var(--surface2);}

/* ── Status badges ───────────────────────────────────────────── */
.badge{
  display:inline-flex;align-items:center;gap:3px;
  padding:2px 8px;border-radius:20px;
  font-size:10px;font-family:var(--mono);font-weight:700;
  letter-spacing:.3px;white-space:nowrap;
}
.ba{background:var(--green-d);color:var(--green);border:1px solid #00c27a33;}
.bp{background:var(--amber-d);color:var(--amber);border:1px solid #f0a03033;}
.bn{background:var(--red-d);color:var(--red);border:1px solid #f0406033;}
.bx{background:var(--muted2);color:var(--muted);border:1px solid var(--border2);}
.br{background:var(--blue-d);color:var(--blue);border:1px solid #4a90e833;}
.hp{
  display:inline-block;
  background:var(--surface3);border:1px solid var(--border2);
  border-radius:4px;padding:1px 7px;
  font-family:var(--mono);font-size:10px;
  margin:1px 2px 1px 0;color:var(--text);
}

/* ── Note chip ───────────────────────────────────────────────── */
.note-chip{
  display:inline-block;
  background:var(--blue-d);border:1px solid var(--blue-d);
  border-radius:4px;padding:1px 7px;
  font-size:10px;color:var(--blue);font-family:var(--mono);
  max-width:160px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
  cursor:help;
}

/* ── Pagination ──────────────────────────────────────────────── */
.pag{
  padding:10px 14px;display:flex;gap:4px;
  align-items:center;flex-wrap:wrap;
  border-top:1px solid var(--border);
}
.pag button{
  background:var(--surface2);border:1px solid var(--border);
  color:var(--muted);border-radius:5px;padding:3px 9px;
  font-size:11px;cursor:pointer;font-family:var(--mono);
  transition:all .15s;
}
.pag button.active{background:var(--blue);border-color:var(--blue);color:#fff;}
.pag button:disabled{opacity:.25;cursor:not-allowed;}
.pag button:not(.active):not(:disabled):hover{border-color:var(--blue);color:var(--blue);}
.pag-info{color:var(--muted);font-size:10px;margin-left:auto;font-family:var(--mono);}
.empty{
  padding:52px;text-align:center;
  color:var(--muted);font-family:var(--mono);font-size:11px;
}

/* ── Modal ───────────────────────────────────────────────────── */
.modal-bg{
  display:none;position:fixed;inset:0;
  background:rgba(0,0,0,.65);z-index:300;
  align-items:center;justify-content:center;
  backdrop-filter:blur(4px);
}
.modal-bg.open{display:flex;}
.modal{
  background:var(--surface);border:1px solid var(--border2);
  border-radius:12px;padding:28px 30px;
  width:460px;max-width:94vw;max-height:88vh;overflow-y:auto;
  animation:min .18s cubic-bezier(.4,0,.2,1);
}
@keyframes min{from{opacity:0;transform:translateY(-8px) scale(.98);}to{opacity:1;transform:none;}}
.modal-title{
  font-family:var(--mono);font-size:11px;font-weight:700;
  color:var(--blue);letter-spacing:1.5px;text-transform:uppercase;margin-bottom:4px;
}
.modal-sub{font-size:11px;color:var(--muted);margin-bottom:20px;}
.fr{margin-bottom:14px;}
.fr label{
  display:block;color:var(--muted);font-size:9px;font-weight:700;
  margin-bottom:5px;text-transform:uppercase;letter-spacing:1px;font-family:var(--mono);
}
.fr select,.fr textarea{
  width:100%;background:var(--surface2);border:1px solid var(--border2);
  border-radius:6px;color:var(--text);font-family:var(--mono);font-size:12px;
  padding:8px 11px;outline:none;resize:vertical;transition:all .15s;
}
.fr select:focus,.fr textarea:focus{border-color:var(--blue);background:var(--surface3);}
.fr select option{background:var(--surface2);}
.modal-actions{display:flex;gap:8px;margin-top:18px;justify-content:flex-end;}

/* ── Toasts ──────────────────────────────────────────────────── */
.toast-stack{position:fixed;bottom:20px;right:20px;display:flex;flex-direction:column;gap:6px;z-index:999;}
.toast{
  background:var(--surface);border:1px solid var(--border2);border-radius:7px;
  padding:9px 16px;font-family:var(--mono);font-size:11px;
  opacity:0;transform:translateX(10px);transition:all .2s;color:var(--text);
  pointer-events:none;max-width:300px;
}
.toast.show{opacity:1;transform:none;}
.t-ok  {border-left:3px solid var(--green);color:var(--green);}
.t-err {border-left:3px solid var(--red);color:var(--red);}
.t-warn{border-left:3px solid var(--amber);color:var(--amber);}
.t-info{border-left:3px solid var(--blue);color:var(--blue);}
</style>
</head>
<body>

<!-- Unsaved-changes banner -->
<div id="dirtyBanner">
  <span>&#x26A0;&nbsp; Unsaved changes — exists only in this tab</span>
  <div class="spacer"></div>
  <button class="btn btn-discard" onclick="discard()">Discard</button>
</div>

<header>
  <div class="logo">
    <div class="logo-pulse"></div>
    <div>
      <h1>SIEM Signoff Dashboard</h1>
      <div class="logo-sub" id="lastRun">Loading...</div>
    </div>
  </div>
  <div class="spacer"></div>
  <div class="hdr-right">
    <div class="pills" id="pills">
      <button onclick="setPeriod('7d',this)" class="active">7D</button>
      <button onclick="setPeriod('30d',this)">30D</button>
      <button onclick="setPeriod('90d',this)">90D</button>
    </div>
    <div class="search-wrap">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none"
           stroke="currentColor" stroke-width="2.5">
        <circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/>
      </svg>
      <input type="text" id="search" placeholder="Host / sender / notes..." oninput="render()">
    </div>
    <button class="btn btn-amber" onclick="exportJSON()">&#x2B07; Export</button>
    <button class="btn btn-save" id="saveBtn" onclick="save()" disabled>
      &#x2713; Save Changes
    </button>
  </div>
</header>

<main>
  <div class="stats" id="stats"></div>
  <div class="charts">
    <div class="chart-box"><h3>Status Breakdown</h3><div id="statusChart"></div></div>
    <div class="chart-box"><h3>Top Hostnames</h3><div id="hostChart"></div></div>
  </div>
  <div class="table-wrap">
    <div class="table-hdr">
      <h2>Signoff Log</h2>
      <span class="rec-badge" id="recBadge">0</span>
      <div class="spacer"></div>
    </div>
    <table>
      <thead><tr>
        <th onclick="sortBy('timestamp')" id="th-timestamp">
          Timestamp<span class="arr">&#x25BC;</span>
        </th>
        <th>Hostnames</th>
        <th onclick="sortBy('overall_status')" id="th-overall_status">
          Status<span class="arr">&#x25BC;</span>
        </th>
        <th onclick="sortBy('sender')" id="th-sender">
          Sender<span class="arr">&#x25BC;</span>
        </th>
        <th>Flags</th>
        <th>Notes</th>
        <th style="text-align:right;min-width:100px;">Actions</th>
      </tr></thead>
      <tbody id="tbody"></tbody>
    </table>
    <div class="pag" id="pag"></div>
  </div>
</main>

<!-- Edit modal -->
<div class="modal-bg" id="modalBg" onclick="if(event.target===this)closeModal()">
  <div class="modal">
    <div class="modal-title">Edit Record</div>
    <div class="modal-sub" id="modalSub"></div>
    <div class="fr">
      <label>Status Override</label>
      <select id="eStatus">
        <option value="">— keep current —</option>
        <option value="active">Active</option>
        <option value="partial">Partial</option>
        <option value="not_found">Not Found</option>
      </select>
    </div>
    <div class="fr">
      <label>Manually Resolved</label>
      <select id="eResolved">
        <option value="false">No — still open</option>
        <option value="true">Yes — exclude from future dedup</option>
      </select>
    </div>
    <div class="fr">
      <label>Notes</label>
      <textarea id="eNotes" rows="3" placeholder="Ticket ID, actions taken..."></textarea>
    </div>
    <div class="modal-actions">
      <button class="btn" onclick="closeModal()">Cancel</button>
      <button class="btn btn-save" onclick="applyEdit()">&#x2713; Apply</button>
    </div>
  </div>
</div>

<div class="toast-stack" id="toasts"></div>

<script>
// ═══════════════════════════════════════════════════════
// State
// ═══════════════════════════════════════════════════════
let RAW = null;   // last persisted snapshot
let D   = null;   // working copy (may differ when dirty)
let _dirty = false;

let _sf   = 'timestamp';  // sort field
let _sasc = false;        // sort ascending?
let _page = 1;
const PAGE_SIZE = 15;
let _period = '7d';
let _editIdx = null;       // index in D.entries being edited

// ═══════════════════════════════════════════════════════
// Boot
// ═══════════════════════════════════════════════════════
async function init() {
  try {
    const r = await fetch('/data.json', {cache:'no-cache'});
    if (!r.ok) throw new Error('HTTP ' + r.status);
    RAW = await r.json();
  } catch(e) {
    console.warn('Fetch failed:', e);
    RAW = {schema_version:3, entries:[]};
    toast('Could not load data.json — ' + e.message, 'err');
  }
  D = clone(RAW);
  updateLastRun();
  render();
}

function clone(o) { return JSON.parse(JSON.stringify(o)); }

function updateLastRun() {
  const entries = D.entries || [];
  if (!entries.length) {
    document.getElementById('lastRun').textContent = 'No data yet';
    return;
  }
  const last = entries.slice().sort((a,b) => (b.timestamp||'').localeCompare(a.timestamp||''))[0];
  const d = new Date(last.timestamp);
  document.getElementById('lastRun').textContent =
    'Last entry: ' + d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});
}

// ═══════════════════════════════════════════════════════
// Dirty state
// ═══════════════════════════════════════════════════════
function markDirty() {
  _dirty = true;
  document.getElementById('dirtyBanner').classList.add('show');
  const b = document.getElementById('saveBtn');
  b.disabled = false;
  b.classList.add('dirty');
}
function markClean() {
  _dirty = false;
  document.getElementById('dirtyBanner').classList.remove('show');
  const b = document.getElementById('saveBtn');
  b.disabled = true;
  b.classList.remove('dirty');
}

// ═══════════════════════════════════════════════════════
// Save / discard
// ═══════════════════════════════════════════════════════
async function save() {
  if (!_dirty) return;
  const btn = document.getElementById('saveBtn');
  btn.textContent = 'Saving...';
  btn.disabled = true;
  try {
    const r = await fetch('/save', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify(D),
    });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    RAW = clone(D);
    markClean();
    btn.textContent = '✓ Save Changes';
    toast('Saved to signoff_data.json', 'ok');
  } catch(e) {
    btn.disabled = false;
    btn.classList.add('dirty');
    btn.textContent = '✓ Save Changes';
    toast('Server save failed — downloading JSON instead', 'warn');
    downloadJSON();
  }
}

function discard() {
  if (!confirm('Discard all unsaved changes?')) return;
  D = clone(RAW);
  markClean();
  _page = 1;
  render();
  toast('Changes discarded', 'warn');
}

// ═══════════════════════════════════════════════════════
// Filtering & sorting
// ═══════════════════════════════════════════════════════
function cutoffDate() {
  const days = {'7d':7,'30d':30,'90d':90}[_period] || 7;
  const d = new Date();
  d.setDate(d.getDate() - days);
  return d;
}

function filtered() {
  const q = (document.getElementById('search').value || '').toLowerCase().trim();
  const co = cutoffDate();
  return (D.entries || []).filter(e => {
    if (new Date(e.timestamp) < co) return false;
    if (!q) return true;
    const hosts  = (e.hosts || []).map(h => h.hostname || '').join(' ').toLowerCase();
    const sender = (e.sender || '').toLowerCase();
    const notes  = (e.notes  || '').toLowerCase();
    const subj   = (e.email_subject || '').toLowerCase();
    return hosts.includes(q) || sender.includes(q) || notes.includes(q) || subj.includes(q);
  });
}

function sorted(arr) {
  return [...arr].sort((a,b) => {
    const av = a[_sf] || '';
    const bv = b[_sf] || '';
    if (av < bv) return _sasc ? -1 : 1;
    if (av > bv) return _sasc ?  1 : -1;
    return 0;
  });
}

// ═══════════════════════════════════════════════════════
// Render
// ═══════════════════════════════════════════════════════
function render() {
  const all   = sorted(filtered());
  const total = all.length;
  const pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
  if (_page > pages) _page = 1;
  const slice = all.slice((_page-1)*PAGE_SIZE, _page*PAGE_SIZE);

  document.getElementById('recBadge').textContent =
    total + ' record' + (total !== 1 ? 's' : '');

  // Update sort headers
  ['timestamp','overall_status','sender'].forEach(f => {
    const th = document.getElementById('th-' + f);
    if (!th) return;
    th.classList.toggle('sorted', _sf === f);
    const a = th.querySelector('.arr');
    if (a) a.textContent = (_sf === f && _sasc) ? '▲' : '▼';
  });

  // Table body
  const tbody = document.getElementById('tbody');
  if (!slice.length) {
    tbody.innerHTML = '<tr><td colspan="7"><div class="empty">No records match this filter.</div></td></tr>';
  } else {
    tbody.innerHTML = slice.map((x, i) => {
      // Stable global index into D.entries
      const gi = (D.entries || []).indexOf(all[(_page-1)*PAGE_SIZE + i]);

      const hosts = (x.hosts || []).map(h =>
        `<span class="hp">${esc(h.hostname || '?')}</span>`
      ).join('');

      let flag = '';
      if      (x.is_revalidation) flag = '<span class="badge br">&#x21BB; Reval</span>';
      else if (!x.is_revalidation && !x.manually_resolved) flag = '<span style="color:var(--muted2);font-size:10px;">new</span>';

      const noteCell = x.notes
        ? `<span class="note-chip" title="${esc(x.notes)}">${esc(x.notes)}</span>`
        : '<span style="color:var(--muted2);">—</span>';

      return `<tr>
        <td style="white-space:nowrap;font-family:var(--mono);font-size:11px;">${fmtDate(x.timestamp)}</td>
        <td>${hosts}</td>
        <td>${statusBadge(x.overall_status, x.manually_resolved)}</td>
        <td style="white-space:nowrap;">${fmtSender(x.sender)}</td>
        <td>${flag}</td>
        <td>${noteCell}</td>
        <td style="text-align:right;white-space:nowrap;">
          <button class="btn" onclick="openEdit(${gi})">Edit</button>
          <button class="btn btn-danger" onclick="delRow(${gi})" title="Delete">&#x2715;</button>
        </td>
      </tr>`;
    }).join('');
  }

  // Pagination
  let ph = `<button onclick="goPage(${_page-1})" ${_page===1?'disabled':''}>&#x2039; Prev</button>`;
  const s = Math.max(1, _page-2), e = Math.min(pages, _page+2);
  if (s > 1) ph += `<button onclick="goPage(1)">1</button>${s>2?'<span style="color:var(--muted);padding:0 3px">…</span>':''}`;
  for (let p=s; p<=e; p++)
    ph += `<button onclick="goPage(${p})" class="${p===_page?'active':''}">${p}</button>`;
  if (e < pages) ph += `${e<pages-1?'<span style="color:var(--muted);padding:0 3px">…</span>':''}<button onclick="goPage(${pages})">${pages}</button>`;
  ph += `<button onclick="goPage(${_page+1})" ${_page===pages?'disabled':''}>Next &#x203A;</button>
         <span class="pag-info">${total} record${total!==1?'s':''} &middot; page ${_page}/${pages}</span>`;
  document.getElementById('pag').innerHTML = ph;

  renderStats(filtered());
  renderCharts(filtered());
}

function goPage(p)   { _page = p; render(); window.scrollTo({top:0,behavior:'smooth'}); }
function sortBy(f)   { _sf === f ? _sasc = !_sasc : (_sf=f, _sasc=false); render(); }
function setPeriod(p, btn) {
  _period = p; _page = 1;
  document.querySelectorAll('.pills button').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  render();
}

// ═══════════════════════════════════════════════════════
// Stats
// ═══════════════════════════════════════════════════════
function renderStats(e) {
  const total    = e.length;
  const active   = e.filter(x => x.overall_status==='active'    && !x.manually_resolved).length;
  const partial  = e.filter(x => x.overall_status==='partial'   && !x.manually_resolved).length;
  const notFound = e.filter(x => x.overall_status==='not_found' && !x.manually_resolved).length;
  const resolved = e.filter(x => x.manually_resolved).length;
  const pct      = total ? Math.round(active/total*100) : 0;

  document.getElementById('stats').innerHTML = `
    <div class="card b"><div class="card-bar"></div>
      <div class="num">${total}</div>
      <div class="clabel">Total</div>
      <div class="sub">in period</div></div>
    <div class="card g"><div class="card-bar"></div>
      <div class="num">${active}</div>
      <div class="clabel">Active</div>
      <div class="sub">${pct}% of period</div></div>
    <div class="card a"><div class="card-bar"></div>
      <div class="num">${partial}</div>
      <div class="clabel">Partial</div>
      <div class="sub">missing sources</div></div>
    <div class="card r"><div class="card-bar"></div>
      <div class="num">${notFound}</div>
      <div class="clabel">Not Found</div>
      <div class="sub">not onboarded</div></div>
    <div class="card p"><div class="card-bar"></div>
      <div class="num">${resolved}</div>
      <div class="clabel">Resolved</div>
      <div class="sub">manually closed</div></div>`;
}

// ═══════════════════════════════════════════════════════
// Charts
// ═══════════════════════════════════════════════════════
function renderCharts(e) {
  // Status breakdown
  const sm = {};
  e.forEach(x => {
    const k = x.manually_resolved ? 'resolved' : (x.overall_status || 'unknown');
    sm[k] = (sm[k]||0) + 1;
  });
  const colors = {active:'var(--green)',partial:'var(--amber)',
                  not_found:'var(--red)',resolved:'var(--muted)',unknown:'var(--border2)'};
  barChart(
    document.getElementById('statusChart'),
    Object.entries(sm).map(([k,v])=>({k,v})).sort((a,b)=>b.v-a.v),
    k => colors[k] || 'var(--blue)'
  );

  // Top hostnames
  const hm = {};
  e.forEach(x => (x.hosts||[]).forEach(h => {
    if (h.hostname) hm[h.hostname] = (hm[h.hostname]||0)+1;
  }));
  barChart(
    document.getElementById('hostChart'),
    Object.entries(hm).map(([k,v])=>({k,v})).sort((a,b)=>b.v-a.v).slice(0,8),
    () => 'var(--blue)'
  );
}

function barChart(el, items, colorFn) {
  if (!items.length) { el.innerHTML = '<div class="no-data">No data</div>'; return; }
  const mx = Math.max(...items.map(i=>i.v), 1);
  el.innerHTML = items.map(i => `
    <div class="bar-row">
      <span class="bar-key" title="${esc(i.k)}">${esc(i.k)}</span>
      <div class="bar-track">
        <div class="bar-fill" style="width:${Math.round(i.v/mx*100)}%;background:${colorFn(i.k)}"></div>
      </div>
      <span class="bar-val">${i.v}</span>
    </div>`).join('');
}

// ═══════════════════════════════════════════════════════
// Row actions
// ═══════════════════════════════════════════════════════
function delRow(gi) {
  const e = (D.entries||[])[gi];
  if (!e) return;
  const label = (e.hosts||[]).map(h=>h.hostname||'?').join(', ') || 'this record';
  if (!confirm(`Delete record for "${label}"?\n\nClick Save Changes to persist.`)) return;
  D.entries.splice(gi, 1);
  markDirty();
  const pages = Math.max(1, Math.ceil(filtered().length / PAGE_SIZE));
  if (_page > pages) _page = pages;
  render();
  toast('Deleted — click Save Changes to persist', 'warn');
}

function openEdit(gi) {
  _editIdx = gi;
  const e = (D.entries||[])[gi];
  if (!e) return;
  const hosts = (e.hosts||[]).map(h=>h.hostname||'?').join(', ');
  document.getElementById('modalSub').textContent = hosts || e.email_subject || '';
  document.getElementById('eStatus').value   = e.overall_status || '';
  document.getElementById('eResolved').value = e.manually_resolved ? 'true' : 'false';
  document.getElementById('eNotes').value    = e.notes || '';
  document.getElementById('modalBg').classList.add('open');
}

function closeModal() {
  document.getElementById('modalBg').classList.remove('open');
  _editIdx = null;
}

function applyEdit() {
  if (_editIdx === null) return;
  const e  = D.entries[_editIdx];
  const ns = document.getElementById('eStatus').value;
  if (ns) e.overall_status = ns;
  e.manually_resolved = document.getElementById('eResolved').value === 'true';
  e.notes = document.getElementById('eNotes').value.trim();
  closeModal();
  markDirty();
  render();
  toast('Edit applied — click Save Changes to persist', 'info');
}

// ═══════════════════════════════════════════════════════
// Formatting helpers
// ═══════════════════════════════════════════════════════
function statusBadge(status, resolved) {
  if (resolved) return '<span class="badge bx">&#x2714; Resolved</span>';
  const map = {
    active:    ['ba','&#x2714;','Active'],
    partial:   ['bp','&#x26A0;','Partial'],
    not_found: ['bn','&#x2716;','Not Found'],
  };
  const [cls, icon, label] = map[status] || ['bx','?', status||'—'];
  return `<span class="badge ${cls}">${icon} ${label}</span>`;
}

function fmtDate(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  const date = d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});
  const time = d.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'});
  return `${date} <span style="color:var(--muted);font-size:10px">${time}</span>`;
}

function fmtSender(s) {
  if (!s) return '<span style="color:var(--muted2)">—</span>';
  const parts = s.split('@');
  if (parts.length < 2)
    return `<span style="font-family:var(--mono);font-size:11px">${esc(s)}</span>`;
  return `<span style="font-family:var(--mono);font-size:11px">${esc(parts[0])}</span>`
       + `<span style="font-family:var(--mono);font-size:10px;color:var(--muted)">@${esc(parts[1])}</span>`;
}

function esc(s) {
  return String(s||'')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}

// ═══════════════════════════════════════════════════════
// Export
// ═══════════════════════════════════════════════════════
function downloadJSON() {
  const b = new Blob([JSON.stringify(D,null,2)],{type:'application/json'});
  Object.assign(document.createElement('a'),{
    href: URL.createObjectURL(b),
    download: 'signoff_data.json',
  }).click();
}
function exportJSON() {
  downloadJSON();
  toast('Exported signoff_data.json', 'info');
}

// ═══════════════════════════════════════════════════════
// Toasts
// ═══════════════════════════════════════════════════════
function toast(msg, type='ok') {
  const stack = document.getElementById('toasts');
  const el = document.createElement('div');
  el.className = `toast t-${type}`;
  el.textContent = msg;
  stack.appendChild(el);
  requestAnimationFrame(() => el.classList.add('show'));
  setTimeout(() => {
    el.classList.remove('show');
    setTimeout(() => el.remove(), 300);
  }, 4000);
}

// ═══════════════════════════════════════════════════════
// Keyboard shortcuts
// ═══════════════════════════════════════════════════════
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') closeModal();
  if ((e.ctrlKey||e.metaKey) && e.key==='s' && _dirty) {
    e.preventDefault();
    save();
  }
});

// Boot
init();
</script>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════════════════════
# SERVER ORCHESTRATION
# ══════════════════════════════════════════════════════════════════════════════

def run(port: int, minutes: int) -> None:
    """Start the HTTP server, open the browser, wait for timeout or Ctrl+C."""
    port   = _find_free_port(port)
    server = HTTPServer(('127.0.0.1', port), DashboardHandler)
    url    = f'http://127.0.0.1:{port}/'

    # Validate data file on startup
    if not os.path.exists(SIGNOFF_DATA_PATH):
        _log(f"WARNING: {SIGNOFF_DATA_PATH} not found — dashboard will show empty data.")
        _log("Run signoff_runner.py first to populate data.")
    else:
        try:
            with open(SIGNOFF_DATA_PATH, 'r', encoding='utf-8') as f:
                data = json.load(f)
            n = len(data.get('entries', []))
            _log(f"Data file: {SIGNOFF_DATA_PATH} ({n} entries)")
        except Exception as exc:
            _log(f"WARNING: Could not parse data file: {exc}")

    _log(f"Dashboard: {url}")
    if minutes > 0:
        _log(f"Server will close automatically after {minutes} minute(s).  Ctrl+C to exit early.")
    else:
        _log("Server running until Ctrl+C.")

    # Serve in background thread
    t = threading.Thread(target=server.serve_forever, daemon=True)
    t.start()

    # Open browser after a short delay (let server bind)
    def _open():
        time.sleep(0.4)
        try:
            webbrowser.open(url)
        except Exception as exc:
            _log(f"WARNING: Could not open browser: {exc}")

    threading.Thread(target=_open, daemon=True).start()

    try:
        if minutes > 0:
            time.sleep(minutes * 60)
            _log(f"Timeout ({minutes} min) reached — shutting down.")
        else:
            # Block forever until Ctrl+C
            while True:
                time.sleep(3600)
    except KeyboardInterrupt:
        _log("\nInterrupted by user — shutting down.")
    finally:
        server.shutdown()


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='SIEM Signoff Dashboard — serves signoff_data.json in your browser.'
    )
    parser.add_argument(
        '--port', type=int, default=DEFAULT_PORT,
        help=f'Preferred local port (default: {DEFAULT_PORT}; auto-picks if in use)',
    )
    parser.add_argument(
        '--minutes', type=int, default=DEFAULT_MINUTES,
        help=f'Minutes to keep server running; 0 = run until Ctrl+C (default: {DEFAULT_MINUTES})',
    )
    args = parser.parse_args()
    run(port=args.port, minutes=args.minutes)
