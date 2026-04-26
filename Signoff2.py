"""
QRadar Signoff Auto-Draft — Production Edition

Fixes and improvements in this version:
  · FIXED  Stale lockfile: age-based detection instead of simple existence check
  · FIXED  Hostname deduplication: "HOST1 | HOST1" no longer double-queries QRadar
  · FIXED  All-time bar chart: dynamically computes date span from earliest record
  · FIXED  ESCALATION_CC wildcards: documented and removed from routing (only literal addrs)
  · FIXED  Subject guard used '[processed]' (closed bracket) missing [Processed-Active] etc — now uses '[processed' (open)
  · FIXED  build_reply_for_all_hosts: overall_status now derived correctly even with empty host list
  · FIXED  QRadar API calls now have retry + 429 rate-limit back-off via _api_get()
  · ADDED  OVERRIDES_FILE: persistent manual overrides (status, note, delete) via export/import cycle
  · ADDED  Log rotation: auto-trims log file at LOG_MAX_MB to avoid unbounded growth
  · ADDED  Dashboard: soft-delete with restore, status override (purple badge), note field
  · ADDED  Dashboard: host history timeline drawer, copy-hostname, revalidation count column
  · ADDED  Dashboard: status filter chips, search covers hostname + sender
  · ADDED  Dashboard: import/export overrides JSON, unsaved-changes banner, localStorage session
  · ADDED  Dashboard: Esc key closes modals, toast notifications for all actions
  · ADDED  Dashboard: grid lines + Y-axis labels on bar chart
"""

import requests
import urllib3
import os
import json
import uuid
import time
import win32com.client
from datetime import datetime, timedelta


# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════

QRADAR_HOST     = os.environ.get('QRADAR_HOST',     'https://your-qradar-host')
QRADAR_USERNAME = os.environ.get('QRADAR_USERNAME', 'your-username')
QRADAR_PASSWORD = os.environ.get('QRADAR_PASSWORD', 'your-password')
VERIFY_SSL      = False

# Subject: "Security Signoff | HOST1 | HOST2 | HOST3"
SUBJECT_KEYWORD   = 'Security Signoff'
SUBJECT_SEPARATOR = '|'

LOOKBACK_HOURS      = 24
ALLOWED_SENDERS     = []           # exact addresses or @domain wildcards; [] = allow all
YOUR_EMAIL_ADDRESS  = 'youremail@yourorg.com'
TRIGGER_DL          = '@SOC-DL@yourorg.com'   # must appear in body; '' = disabled
SIGNOFF_FOLDER_NAME = 'SIEM Signoffs'          # None = full Inbox

RUN_LOG_PATH   = r'C:\path\to\signoff_runner.log'
LOCKFILE_PATH  = r'C:\path\to\signoff.lock'
TRACKING_FILE  = r'C:\path\to\signoff_tracking.json'
OVERRIDES_FILE = r'C:\path\to\signoff_overrides.json'   # UI overrides (edit/delete)
DASHBOARD_FILE = r'C:\path\to\signoff_dashboard.html'

LOCKFILE_STALE_MINUTES = 120   # lockfile older than this is treated as stale/orphaned
LOG_MAX_MB             = 5     # rotate log when it exceeds this size
REQUEST_TIMEOUT        = 30
API_RETRIES            = 3
API_BACKOFF_SECONDS    = 2.0

# Partial / Not Found → override recipients with these lists only.
# Active              → ReplyAll to original thread, no override.
# NOTE: Literal email addresses only. @domain wildcards are NOT supported here.
ESCALATION_TO      = ['onboarding-owner@yourorg.com']
ESCALATION_CC      = ['soc-dl@yourorg.com']
ESCALATION_CONTACT = '@xyz'   # name shown inline in email body for missing/silent rows

REVALIDATION_WINDOW_DAYS   = 14
REVALIDATION_COOLDOWN_DAYS = 3

TAG_ACTIVE         = '[Processed-Active]'
TAG_PARTIAL        = '[Processed-Partial]'
TAG_NOT_FOUND      = '[Processed-NotFound]'
REVALIDATABLE_TAGS = {TAG_PARTIAL, TAG_NOT_FOUND}

OS_TYPE_GROUPS = {
    'Windows': {'required': ['Microsoft Security', 'WinCollect']},
    'Linux':   {'required': ['Linux OS']},
}

ACTIVITY_THRESHOLD_DAYS = 7
_MIN_TS = 0
_MAX_TS = 2147483647
LOG_SOURCE_TYPES_CACHE = {}


# ══════════════════════════════════════════════════════════════════════════════
#  DASHBOARD HTML TEMPLATE
#  Placeholders replaced at runtime: __TRACKING_DATA__  __OVERRIDES_DATA__  __GENERATED__
# ══════════════════════════════════════════════════════════════════════════════

_DASHBOARD_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SIEM Signoff Dashboard</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg0:#070d1a;--bg1:#0d1525;--bg2:#111e30;--bg3:#192540;
  --bd:#1f3050;--bd2:#2a3f60;
  --t1:#e8f0fe;--t2:#8fa8c8;--t3:#4d6480;
  --blue:#4d8cff;--blue-d:#1a4fbf;--blue-bg:#1a4fbf18;
  --green:#3dd68c;--green-bg:#0d3d2518;--green-bd:#145232;
  --amber:#f5a623;--amber-bg:#4a2c0018;--amber-bd:#7a4800;
  --red:#f06060;--red-bg:#4a0d0d18;--red-bd:#7a1515;
  --purple:#b388ff;--purple-bg:#2d1a5018;--purple-bd:#5c3494;
  --r:8px;--rs:5px;--sh:0 8px 32px #00000060;--tr:all .14s ease;
  --fnt:'Segoe UI',system-ui,sans-serif;--mono:'Consolas','Courier New',monospace;
}
html{scroll-behavior:smooth}
body{font-family:var(--fnt);background:var(--bg0);color:var(--t1);font-size:13px;line-height:1.55;min-height:100vh}
input,select,textarea,button{font-family:var(--fnt)}
button{cursor:pointer}

/* LAYOUT */
.topbar{background:var(--bg2);border-bottom:1px solid var(--bd);padding:11px 24px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}
.topbar-l{display:flex;align-items:center;gap:10px}
.ttl{font-size:15px;font-weight:600}
.qbadge{background:var(--blue-bg);color:var(--blue);font-size:10px;padding:2px 8px;border-radius:4px;font-weight:700;border:1px solid var(--blue-d);letter-spacing:.3px}
.topbar-r{font-size:11px;color:var(--t3);display:flex;gap:14px;align-items:center}
.main{padding:20px 24px 40px;max-width:1400px;margin:0 auto;width:100%}

/* DIRTY BANNER */
.dirty{background:var(--amber-bg);border:1px solid var(--amber-bd);border-radius:var(--rs);padding:9px 14px;margin-bottom:14px;display:flex;align-items:center;justify-content:space-between;gap:12px}
.dirty.hidden{display:none}
.dirty-txt{color:var(--amber);font-size:12px;font-weight:500}
.dirty-acts{display:flex;gap:8px}

/* PERIOD */
.period{display:flex;gap:6px;margin-bottom:16px;flex-wrap:wrap}
.pbtn{background:var(--bg2);border:1px solid var(--bd);color:var(--t2);padding:5px 14px;border-radius:var(--rs);font-size:12px;font-weight:500;transition:var(--tr)}
.pbtn:hover{border-color:var(--blue);color:var(--blue)}
.pbtn.on{background:var(--blue-d);border-color:var(--blue);color:#fff}

/* CARDS */
.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:16px}
.card{background:var(--bg2);border:1px solid var(--bd);border-radius:var(--r);padding:16px}
.clbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--t3);margin-bottom:8px}
.cval{font-size:28px;font-weight:600;line-height:1}
.csub{font-size:11px;color:var(--t3);margin-top:4px}
.c-b{color:var(--blue)}.c-g{color:var(--green)}.c-a{color:var(--amber)}.c-r{color:var(--red)}

/* CHARTS */
.charts{display:grid;grid-template-columns:230px 1fr;gap:12px;margin-bottom:16px}
.cbox{background:var(--bg2);border:1px solid var(--bd);border-radius:var(--r);padding:16px}
.cttl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--t3);margin-bottom:12px}
.donut-wrap{position:relative;width:100%}
.donut-center{position:absolute;top:50%;left:50%;transform:translate(-50%,-55%);text-align:center;pointer-events:none}
.dc-val{font-size:22px;font-weight:600;line-height:1}
.dc-lbl{font-size:10px;color:var(--t3);margin-top:2px}
.legend{margin-top:12px;display:flex;flex-direction:column;gap:6px}
.leg-row{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--t2);padding:1px 0}
.leg-dot{width:8px;height:8px;border-radius:2px;flex-shrink:0}

/* TOOLBAR */
.toolbar{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:10px;flex-wrap:wrap}
.tbl-l{display:flex;align-items:center;gap:7px;flex-wrap:wrap}
.tbl-r{display:flex;align-items:center;gap:6px}
.sbox{background:var(--bg2);border:1px solid var(--bd);border-radius:var(--rs);color:var(--t1);padding:6px 12px;font-size:12px;width:220px;transition:var(--tr)}
.sbox:focus{outline:none;border-color:var(--blue)}
.chip{background:var(--bg2);border:1px solid var(--bd);color:var(--t3);padding:4px 12px;border-radius:20px;font-size:11px;font-weight:700;cursor:pointer;transition:var(--tr);opacity:.45}
.chip.on{opacity:1}
.chip:hover{opacity:.8}
.chip-g.on{background:var(--green-bg);border-color:var(--green-bd);color:var(--green)}
.chip-a.on{background:var(--amber-bg);border-color:var(--amber-bd);color:var(--amber)}
.chip-r.on{background:var(--red-bg);border-color:var(--red-bd);color:var(--red)}
.chip-d.on{color:var(--t2);border-color:var(--bd2)}
.bsm{background:var(--bg2);border:1px solid var(--bd);color:var(--t2);padding:6px 11px;border-radius:var(--rs);font-size:12px;font-weight:500;display:flex;align-items:center;gap:5px;transition:var(--tr);white-space:nowrap}
.bsm:hover{border-color:var(--blue);color:var(--blue)}
.bsm.pri{background:var(--blue-d);border-color:var(--blue);color:#fff}
.bsm.pri:hover{background:var(--blue)}
.bsm.danger:hover{border-color:var(--red-bd);color:var(--red)}

/* TABLE */
.twrap{background:var(--bg2);border:1px solid var(--bd);border-radius:var(--r);overflow:hidden}
.tscroll{overflow-x:auto}
table{width:100%;border-collapse:collapse}
th{background:var(--bg0);color:var(--t3);font-size:10px;font-weight:700;text-align:left;padding:9px 14px;border-bottom:1px solid var(--bd);text-transform:uppercase;letter-spacing:.5px;cursor:pointer;white-space:nowrap;user-select:none}
th:hover{color:var(--t2)}
td{padding:10px 14px;border-bottom:1px solid var(--bg1);vertical-align:middle;font-size:12px}
tr:last-child td{border-bottom:none}
tr.del-row td{opacity:.38}
tr.del-row .hn-txt{text-decoration:line-through}
tr:not(.del-row):hover td{background:var(--bg3)}
.hn{display:flex;align-items:center;gap:6px;cursor:pointer}
.hn-txt{font-family:var(--mono);font-size:12px;font-weight:500;color:var(--t1)}
.cp-icon{opacity:0;font-size:11px;color:var(--t3);transition:var(--tr)}
.hn:hover .cp-icon{opacity:1}
.sbadge{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:20px;font-size:10px;font-weight:700;letter-spacing:.3px;white-space:nowrap}
.s-active{background:var(--green-bg);color:var(--green);border:1px solid var(--green-bd)}
.s-partial{background:var(--amber-bg);color:var(--amber);border:1px solid var(--amber-bd)}
.s-nf{background:var(--red-bg);color:var(--red);border:1px solid var(--red-bd)}
.s-override{background:var(--purple-bg);color:var(--purple);border:1px solid var(--purple-bd)}
.s-del{background:var(--bg3);color:var(--t3);border:1px solid var(--bd2)}
.dot6{width:6px;height:6px;border-radius:50%;display:inline-block;flex-shrink:0}
.obadge{background:var(--blue-bg);color:var(--blue);font-size:10px;padding:2px 7px;border-radius:4px;border:1px solid var(--blue-d);white-space:nowrap}
.rbadge{background:var(--purple-bg);color:var(--purple);border:1px solid var(--purple-bd);font-size:10px;padding:1px 6px;border-radius:4px;font-weight:700}
.tr-up{color:var(--green);font-size:11px}.tr-dn{color:var(--red);font-size:11px}.tr-eq{color:var(--t3);font-size:11px}
.act-btn{background:none;border:1px solid transparent;color:var(--t3);padding:3px 7px;border-radius:var(--rs);font-size:11px;transition:var(--tr)}
.act-btn:hover{border-color:var(--bd2);color:var(--t1);background:var(--bg3)}
.act-btn.del:hover{border-color:var(--red-bd);color:var(--red)}
.act-btn.res:hover{border-color:var(--green-bd);color:var(--green)}
.no-data{text-align:center;padding:40px;color:var(--t3);font-size:13px}

/* MODAL */
.overlay{position:fixed;inset:0;background:#00000085;backdrop-filter:blur(3px);z-index:200;display:flex;align-items:center;justify-content:center;animation:fadeIn .15s ease}
.overlay.hidden{display:none}
.modal{background:var(--bg2);border:1px solid var(--bd);border-radius:var(--r);width:460px;max-width:96vw;max-height:90vh;overflow-y:auto;box-shadow:var(--sh)}
.mhdr{padding:16px 20px 12px;border-bottom:1px solid var(--bd);display:flex;align-items:center;justify-content:space-between}
.mttl{font-size:14px;font-weight:600}
.mclose{background:none;border:none;color:var(--t3);font-size:18px;padding:2px 6px;border-radius:4px;transition:var(--tr)}
.mclose:hover{color:var(--t1);background:var(--bg3)}
.mbody{padding:16px 20px;display:flex;flex-direction:column;gap:14px}
.mfoot{padding:12px 20px;border-top:1px solid var(--bd);display:flex;justify-content:flex-end;gap:8px}
.flbl{font-size:11px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.4px;margin-bottom:6px}
.finput{background:var(--bg1);border:1px solid var(--bd);border-radius:var(--rs);color:var(--t1);padding:8px 10px;font-size:13px;width:100%;transition:var(--tr)}
.finput:focus{outline:none;border-color:var(--blue)}
.finput:disabled{opacity:.45;cursor:not-allowed}
select.finput option{background:var(--bg1)}
textarea.finput{resize:vertical;min-height:72px}

/* DRAWER */
.drw-ov{position:fixed;inset:0;background:#00000060;z-index:200;animation:fadeIn .15s ease}
.drw-ov.hidden{display:none}
.drawer{position:fixed;right:0;top:0;bottom:0;width:440px;max-width:96vw;background:var(--bg2);border-left:1px solid var(--bd);box-shadow:var(--sh);z-index:201;overflow-y:auto;animation:slideIn .18s ease}
.drw-hdr{padding:14px 20px;border-bottom:1px solid var(--bd);display:flex;align-items:flex-start;justify-content:space-between;position:sticky;top:0;background:var(--bg2);z-index:1}
.drw-body{padding:16px 20px}
.tl{display:flex;flex-direction:column;gap:0}
.tl-item{display:flex;gap:12px}
.tl-lc{display:flex;flex-direction:column;align-items:center}
.tl-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0;margin-top:4px}
.tl-line{width:1px;flex:1;background:var(--bd);min-height:24px}
.tl-content{padding-bottom:18px;flex:1}
.tl-ts{font-size:10px;color:var(--t3);margin-bottom:3px}
.tl-reval{font-size:10px;color:var(--purple);margin-top:3px}

/* TOAST */
.toast-wrap{position:fixed;bottom:20px;right:20px;display:flex;flex-direction:column;gap:8px;z-index:300}
.toast{background:var(--bg2);border:1px solid var(--bd);border-radius:var(--rs);padding:10px 14px;font-size:12px;box-shadow:var(--sh);animation:slideUp .18s ease;display:flex;align-items:center;gap:8px;min-width:240px;max-width:360px}
.toast.success{border-color:var(--green-bd);color:var(--green)}
.toast.warning{border-color:var(--amber-bd);color:var(--amber)}
.toast.error{border-color:var(--red-bd);color:var(--red)}
.toast.info{border-color:var(--blue-d);color:var(--blue)}

/* ANIM */
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
@keyframes slideIn{from{transform:translateX(100%)}to{transform:translateX(0)}}
@keyframes slideUp{from{transform:translateY(8px);opacity:0}to{transform:translateY(0);opacity:1}}

@media(max-width:900px){
  .cards{grid-template-columns:repeat(2,1fr)}
  .charts{grid-template-columns:1fr}
  .topbar-r{display:none}
}
</style>
</head>
<body>

<!-- TOPBAR -->
<div class="topbar">
  <div class="topbar-l">
    <span class="ttl">SIEM Signoff Dashboard</span>
    <span class="qbadge">QRadar</span>
  </div>
  <div class="topbar-r">
    <span id="gen-ts"></span>
    <span id="rec-count"></span>
  </div>
</div>

<div class="main">

<!-- DIRTY BANNER -->
<div class="dirty hidden" id="dirty-banner">
  <span class="dirty-txt">&#9888; Unsaved override changes. Export &#8594; save as <code>signoff_overrides.json</code> &#8594; re-run script to persist.</span>
  <div class="dirty-acts">
    <button class="bsm pri" onclick="exportOverrides()">&#8595; Export Overrides</button>
    <button class="bsm danger" onclick="resetOverrides()">&#8617; Discard</button>
  </div>
</div>

<!-- PERIOD -->
<div class="period">
  <button class="pbtn" data-p="7"  onclick="setPeriod(7)">Last 7 days</button>
  <button class="pbtn" data-p="15" onclick="setPeriod(15)">Last 15 days</button>
  <button class="pbtn on" data-p="30" onclick="setPeriod(30)">Last 30 days</button>
  <button class="pbtn" data-p="0"  onclick="setPeriod(0)">All time</button>
</div>

<!-- CARDS -->
<div class="cards">
  <div class="card"><div class="clbl">Signoff Emails</div><div class="cval c-b" id="ct-tot">&#8212;</div><div class="csub" id="ct-hosts">&#8212;</div></div>
  <div class="card"><div class="clbl">Active</div><div class="cval c-g" id="ct-a">&#8212;</div><div class="csub" id="ct-ap">&#8212;</div></div>
  <div class="card"><div class="clbl">Partial</div><div class="cval c-a" id="ct-p">&#8212;</div><div class="csub" id="ct-pp">&#8212;</div></div>
  <div class="card"><div class="clbl">Not Found</div><div class="cval c-r" id="ct-n">&#8212;</div><div class="csub" id="ct-np">&#8212;</div></div>
</div>

<!-- CHARTS -->
<div class="charts">
  <div class="cbox">
    <div class="cttl">Status Distribution</div>
    <div class="donut-wrap">
      <svg id="donut-svg" viewBox="0 0 200 180" width="100%" role="img" aria-label="Status distribution donut chart"><title>Status distribution</title></svg>
      <div class="donut-center"><div class="dc-val" id="dc-val">&#8212;</div><div class="dc-lbl">active</div></div>
    </div>
    <div class="legend" id="donut-leg"></div>
  </div>
  <div class="cbox">
    <div class="cttl">Signoffs Over Time</div>
    <svg id="bar-svg" viewBox="0 0 580 210" width="100%" role="img" aria-label="Signoffs over time bar chart"><title>Signoffs over time</title></svg>
  </div>
</div>

<!-- TOOLBAR -->
<div class="toolbar">
  <div class="tbl-l">
    <input type="text" class="sbox" placeholder="Search hostname, sender&#8230;" oninput="setSearch(this.value)">
    <button class="chip chip-g on" id="chip-active"   onclick="toggleChip('active')">&#9679; Active</button>
    <button class="chip chip-a on" id="chip-partial"  onclick="toggleChip('partial')">&#9679; Partial</button>
    <button class="chip chip-r on" id="chip-nf"       onclick="toggleChip('not_found')">&#9679; Not Found</button>
    <button class="chip chip-d"   id="chip-del"       onclick="toggleChip('deleted')">&#9675; Deleted</button>
  </div>
  <div class="tbl-r">
    <label class="bsm" style="cursor:pointer">
      &#8593; Import Overrides
      <input type="file" accept=".json" style="display:none" onchange="importOverrides(event)">
    </label>
    <button class="bsm" onclick="exportOverrides()">&#8595; Export Overrides</button>
  </div>
</div>

<!-- TABLE -->
<div class="twrap">
  <div class="tscroll">
    <table>
      <thead><tr>
        <th onclick="sortBy('hostname')">Hostname &#8597;</th>
        <th onclick="sortBy('os_group')">OS &#8597;</th>
        <th onclick="sortBy('display_status')">Status &#8597;</th>
        <th onclick="sortBy('last_checked')">Last Checked &#8597;</th>
        <th onclick="sortBy('days_ago')">QRadar Event &#8597;</th>
        <th onclick="sortBy('checks')">Checks &#8597;</th>
        <th onclick="sortBy('revals')">Revals &#8597;</th>
        <th>Trend</th>
        <th>Note</th>
        <th style="text-align:center">Actions</th>
      </tr></thead>
      <tbody id="tbody"></tbody>
    </table>
  </div>
</div>

<div style="font-size:11px;color:var(--t3);text-align:right;margin-top:10px">
  QRadar Signoff Automation &#183; SOC use only &#183; Refreshed on each script run
</div>
</div><!-- /main -->

<!-- EDIT MODAL -->
<div class="overlay hidden" id="edit-ov" onclick="closeEdit(event)">
  <div class="modal" onclick="event.stopPropagation()">
    <div class="mhdr"><span class="mttl">Override Host Status</span><button class="mclose" onclick="closeEdit()">&#10005;</button></div>
    <div class="mbody">
      <div><div class="flbl">Hostname</div><input class="finput" id="e-hn" disabled></div>
      <div><div class="flbl">Actual QRadar Status</div><input class="finput" id="e-actual" disabled></div>
      <div>
        <div class="flbl">Override Status</div>
        <select class="finput" id="e-override">
          <option value="">&#8212; No override (use actual) &#8212;</option>
          <option value="active">Active</option>
          <option value="partial">Partial</option>
          <option value="not_found">Not Found</option>
        </select>
      </div>
      <div><div class="flbl">Note / Exception Reason</div><textarea class="finput" id="e-note" placeholder="e.g. Approved exception &#8212; firewall blocks syslog forwarding until Q3"></textarea></div>
      <div><div class="flbl">Override By</div><input class="finput" id="e-by" placeholder="Name or email of person making this override"></div>
    </div>
    <div class="mfoot">
      <button class="bsm" onclick="closeEdit()">Cancel</button>
      <button class="bsm pri" onclick="saveEdit()">Save Override</button>
    </div>
  </div>
</div>

<!-- DELETE CONFIRM MODAL -->
<div class="overlay hidden" id="del-ov" onclick="closeDel(event)">
  <div class="modal" style="width:400px" onclick="event.stopPropagation()">
    <div class="mhdr"><span class="mttl">Hide Host from Dashboard?</span><button class="mclose" onclick="closeDel()">&#10005;</button></div>
    <div class="mbody">
      <p style="color:var(--t2);font-size:13px;line-height:1.65">
        <strong id="d-hn" style="color:var(--t1)"></strong> will be hidden from dashboard views. This is a <strong>soft delete</strong> &#8212; the full audit trail is preserved and you can restore the host at any time using the <em>Deleted</em> filter.
      </p>
      <div><div class="flbl">Deletion Reason (optional)</div><input class="finput" id="d-reason" placeholder="e.g. Test asset &#8212; not a production host"></div>
    </div>
    <div class="mfoot">
      <button class="bsm" onclick="closeDel()">Cancel</button>
      <button class="bsm" onclick="confirmDel()" style="color:var(--red);border-color:var(--red-bd)">Delete from Dashboard</button>
    </div>
  </div>
</div>

<!-- HISTORY DRAWER -->
<div class="drw-ov hidden" id="drw-ov" onclick="closeDrawer()">
  <div class="drawer" onclick="event.stopPropagation()">
    <div class="drw-hdr">
      <div>
        <div class="mttl" id="drw-ttl">Host History</div>
        <div style="font-size:11px;color:var(--t3);margin-top:2px" id="drw-sub"></div>
      </div>
      <button class="mclose" onclick="closeDrawer()">&#10005;</button>
    </div>
    <div class="drw-body" id="drw-body"></div>
  </div>
</div>

<!-- TOAST -->
<div class="toast-wrap" id="toast-wrap"></div>

<script>
// ── DATA ─────────────────────────────────────────────────────────────────────
const ALL      = __TRACKING_DATA__;
const BASE_OVR = __OVERRIDES_DATA__;
const GEN      = '__GENERATED__';

document.getElementById('gen-ts').textContent   = 'Generated ' + GEN;
document.getElementById('rec-count').textContent = ALL.length + ' records';

const LS_KEY = 'siem_ovr_v3';
let localOvr = {};
try { localOvr = JSON.parse(localStorage.getItem(LS_KEY) || '{}'); } catch(e) {}

function merged() {
  const m = JSON.parse(JSON.stringify(BASE_OVR));
  Object.entries(localOvr).forEach(([k,v]) => { m[k] = Object.assign({}, m[k]||{}, v); });
  return m;
}
function dirty() { return JSON.stringify(merged()) !== JSON.stringify(BASE_OVR); }
function saveLoc() { localStorage.setItem(LS_KEY, JSON.stringify(localOvr)); updateDirtyBanner(); }
function updateDirtyBanner() { document.getElementById('dirty-banner').classList.toggle('hidden', !dirty()); }

// ── STATE ────────────────────────────────────────────────────────────────────
let period = 30, sortCol = 'last_checked', sortAsc = false, srch = '';
let activeFilters = new Set(['active','partial','not_found']);
let showDel = false, allHosts = [];
let editHn = null, delHn = null;

const SLBL  = {active:'Active', partial:'Partial', not_found:'Not Found'};
const SDOT  = {active:'#3dd68c', partial:'#f5a623', not_found:'#f06060'};
const SCLS  = {active:'s-active', partial:'s-partial', not_found:'s-nf'};
const SRANK = {not_found:0, partial:1, active:2};

// ── PERIOD / FILTERS ─────────────────────────────────────────────────────────
function setPeriod(d) {
  period = d;
  document.querySelectorAll('.pbtn').forEach(b => b.classList.toggle('on', +b.dataset.p === d));
  render();
}
function setSearch(v) { srch = v.toLowerCase(); drawTable(); }
function toggleChip(key) {
  if (key === 'deleted') {
    showDel = !showDel;
    document.getElementById('chip-del').classList.toggle('on', showDel);
  } else {
    activeFilters.has(key) ? activeFilters.delete(key) : activeFilters.add(key);
    const ids = {active:'chip-active', partial:'chip-partial', not_found:'chip-nf'};
    document.getElementById(ids[key]).classList.toggle('on', activeFilters.has(key));
  }
  drawTable();
}
function filtRecs() {
  if (!period) return ALL;
  const c = new Date(); c.setDate(c.getDate() - period);
  return ALL.filter(r => new Date(r.timestamp) >= c);
}

// ── HOST MAP ─────────────────────────────────────────────────────────────────
function buildMap(recs) {
  const m = {};
  recs.forEach(r => {
    r.host_results.forEach(h => {
      const k = h.hostname;
      if (!m[k]) m[k] = {hostname:k, os_group:h.os_group||'&#8212;', status:h.status,
        last_checked:r.timestamp, days_ago:h.days_ago, checks:0, revals:0,
        prev_status:null, history:[], senders:new Set()};
      const e = m[k];
      if (r.timestamp > e.last_checked) {
        e.prev_status = e.status; e.status = h.status; e.last_checked = r.timestamp;
        e.days_ago = h.days_ago; e.os_group = h.os_group || e.os_group || '&#8212;';
      }
      e.checks++;
      if (r.is_revalidation) e.revals++;
      if (r.sender) e.senders.add(r.sender);
      e.history.push({ts:r.timestamp, status:h.status, reval:r.is_revalidation, sender:r.sender||''});
    });
  });
  return Object.values(m).map(h => ({...h, senders:[...h.senders]}));
}

function applyOvr(hosts) {
  const o = merged();
  return hosts.map(h => {
    const v = o[h.hostname] || {};
    return {...h,
      deleted:         v.deleted || false,
      override_status: v.override || null,
      note:            v.note    || '',
      override_by:     v.by      || '',
      override_at:     v.at      || '',
      display_status:  v.override || h.status,
    };
  });
}

// ── RENDER ───────────────────────────────────────────────────────────────────
function render() {
  const recs  = filtRecs();
  const hosts = applyOvr(buildMap(recs));
  allHosts = hosts;
  const vis = hosts.filter(h => !h.deleted);
  const a = vis.filter(h => h.display_status==='active').length;
  const p = vis.filter(h => h.display_status==='partial').length;
  const n = vis.filter(h => h.display_status==='not_found').length;
  const u = vis.length;
  const pct = v => u ? Math.round(v/u*100)+'% of hosts' : '&#8212;';

  document.getElementById('ct-tot').textContent = recs.length;
  document.getElementById('ct-hosts').textContent = u+' unique host'+(u!==1?'s':'');
  document.getElementById('ct-a').textContent=a; document.getElementById('ct-ap').textContent=pct(a);
  document.getElementById('ct-p').textContent=p; document.getElementById('ct-pp').textContent=pct(p);
  document.getElementById('ct-n').textContent=n; document.getElementById('ct-np').textContent=pct(n);
  document.getElementById('dc-val').textContent=a;
  renderDonut(a,p,n); renderBar(recs); drawTable(); updateDirtyBanner();
}

// ── DONUT ────────────────────────────────────────────────────────────────────
function renderDonut(a,p,n) {
  const svg=document.getElementById('donut-svg'), total=a+p+n;
  if(!total){svg.innerHTML='<text x="100" y="95" text-anchor="middle" fill="#4d6480" font-size="12" font-family="Segoe UI,sans-serif">No data in period</text>';return;}
  const cx=100,cy=90,R=68,r=48, colors=['#3dd68c','#f5a623','#f06060'], vals=[a,p,n];
  let paths='', angle=-Math.PI/2;
  vals.forEach((v,i)=>{
    if(!v) return;
    const sweep=2*Math.PI*(v/total);
    const x1=cx+R*Math.cos(angle),y1=cy+R*Math.sin(angle); angle+=sweep;
    const x2=cx+R*Math.cos(angle),y2=cy+R*Math.sin(angle);
    const xi1=cx+r*Math.cos(angle-sweep),yi1=cy+r*Math.sin(angle-sweep);
    const xi2=cx+r*Math.cos(angle),yi2=cy+r*Math.sin(angle);
    paths+=`<path d="M${x1},${y1} A${R},${R} 0 ${sweep>Math.PI?1:0},1 ${x2},${y2} L${xi2},${yi2} A${r},${r} 0 ${sweep>Math.PI?1:0},0 ${xi1},${yi1} Z" fill="${colors[i]}" opacity="0.82"/>`;
  });
  svg.innerHTML=paths;
  document.getElementById('donut-leg').innerHTML=
    [['#3dd68c','Active',a],['#f5a623','Partial',p],['#f06060','Not Found',n]].map(([c,l,v])=>
      `<div class="leg-row"><div class="leg-dot" style="background:${c}"></div><span style="flex:1">${l}</span><span style="font-weight:600;color:var(--t1)">${v}</span></div>`
    ).join('');
}

// ── BAR CHART (fixes All-time range) ─────────────────────────────────────────
function renderBar(recs) {
  const el=document.getElementById('bar-svg');
  const W=580,H=210,PL=30,PR=8,PT=10,PB=34;
  const cH=H-PT-PB, cW=W-PL-PR;

  // Compute day count: fixed period or full data span for "All time"
  let days=period;
  if(!days){
    if(!recs.length){el.innerHTML='';return;}
    const dates=recs.map(r=>r.timestamp.slice(0,10)).sort();
    days=Math.max(7, Math.ceil((Date.now()-new Date(dates[0]))/86400000)+1);
  }

  const buckets={};
  for(let i=days-1;i>=0;i--){
    const d=new Date(); d.setDate(d.getDate()-i);
    buckets[d.toISOString().slice(0,10)]={a:0,p:0,n:0};
  }
  recs.forEach(r=>{
    const d=r.timestamp.slice(0,10);
    if(!buckets[d]) return;
    if(r.overall_status==='active') buckets[d].a++;
    else if(r.overall_status==='partial') buckets[d].p++;
    else buckets[d].n++;
  });

  const keys=Object.keys(buckets);
  const maxVal=Math.max(1,...Object.values(buckets).map(b=>b.a+b.p+b.n));
  const bw=Math.max(2,Math.floor(cW/keys.length)-1);
  const gap=cW/keys.length;
  const sc=v=>(v/maxVal)*cH;

  let grid='',bars='',lbls='';
  const gc=4;
  for(let g=0;g<=gc;g++){
    const y=PT+cH-(g/gc)*cH, lv=Math.round((g/gc)*maxVal);
    grid+=`<line x1="${PL}" y1="${y}" x2="${W-PR}" y2="${y}" stroke="#1f3050" stroke-width="0.5"/>`;
    if(g>0||maxVal>0) grid+=`<text x="${PL-3}" y="${y+3.5}" text-anchor="end" font-size="8" fill="#4d6480" font-family="Segoe UI,sans-serif">${lv}</text>`;
  }
  const tick=Math.max(1,Math.ceil(keys.length/14));
  keys.forEach((k,i)=>{
    const {a,p,n}=buckets[k], x=PL+i*gap+(gap-bw)/2;
    let y=H-PB;
    if(n){const h=Math.max(1,sc(n));bars+=`<rect x="${x}" y="${y-h}" width="${bw}" height="${h}" fill="#f0606060" rx="1"/>`;y-=h;}
    if(p){const h=Math.max(1,sc(p));bars+=`<rect x="${x}" y="${y-h}" width="${bw}" height="${h}" fill="#f5a62360" rx="1"/>`;y-=h;}
    if(a){const h=Math.max(1,sc(a));bars+=`<rect x="${x}" y="${y-h}" width="${bw}" height="${h}" fill="#3dd68c60" rx="1"/>`;y-=h;}
    if(i%tick===0||i===keys.length-1)
      lbls+=`<text x="${x+bw/2}" y="${H-PB+14}" text-anchor="middle" font-size="8" fill="#4d6480" font-family="Segoe UI,sans-serif">${k.slice(5)}</text>`;
  });
  el.innerHTML=grid+bars+lbls;
}

// ── TABLE ────────────────────────────────────────────────────────────────────
function sortBy(col){
  sortCol===col ? sortAsc=!sortAsc : (sortCol=col, sortAsc=col==='hostname');
  drawTable();
}
function rel(iso){
  if(!iso) return '&#8212;';
  const d=(Date.now()-new Date(iso))/1000;
  if(d<60) return 'Just now';
  if(d<3600) return Math.floor(d/60)+'m ago';
  if(d<86400) return Math.floor(d/3600)+'h ago';
  const n=Math.floor(d/86400); return n===1?'Yesterday':n+'d ago';
}
function trend(h){
  if(!h.prev_status||h.prev_status===h.display_status) return '<span class="tr-eq">&#8212;</span>';
  return SRANK[h.display_status]>SRANK[h.prev_status]
    ?'<span class="tr-up">&#9650; Improved</span>'
    :'<span class="tr-dn">&#9660; Degraded</span>';
}

function drawTable(){
  let rows=allHosts.filter(h=>{
    if(h.deleted && !showDel) return false;
    if(h.deleted && showDel) return true;
    if(!activeFilters.has(h.display_status)) return false;
    if(srch && !h.hostname.toLowerCase().includes(srch) &&
       !h.senders.some(s=>s.toLowerCase().includes(srch))) return false;
    return true;
  });
  rows.sort((a,b)=>{
    let av=a[sortCol]??'zzz', bv=b[sortCol]??'zzz';
    if(typeof av==='string') av=av.toLowerCase();
    if(typeof bv==='string') bv=bv.toLowerCase();
    return(av<bv?-1:av>bv?1:0)*(sortAsc?1:-1);
  });

  const tb=document.getElementById('tbody');
  if(!rows.length){tb.innerHTML='<tr><td colspan="10" class="no-data">No hosts match the current filters.</td></tr>';return;}

  tb.innerHTML=rows.map(h=>{
    const da=h.days_ago!=null?(h.days_ago===0?'Today':h.days_ago+'d ago'):'&#8212;';
    const daStyle=h.days_ago!=null&&h.days_ago>7?'color:var(--red)':'color:var(--t2)';
    const og=h.os_group&&h.os_group!=='&#8212;'
      ?`<span class="obadge">${h.os_group}</span>`
      :'<span style="color:var(--t3)">&#8212;</span>';
    const rv=h.revals?`<span class="rbadge">${h.revals}</span>`:'<span style="color:var(--t3)">&#8212;</span>';

    let badge;
    if(h.deleted){
      badge=`<span class="sbadge s-del"><span class="dot6" style="background:var(--t3)"></span>Deleted</span>`;
    } else if(h.override_status){
      const tip=`Override by ${h.override_by||'unknown'} &#183; ${h.override_at||''}`;
      badge=`<span class="sbadge s-override" title="${tip}"><span class="dot6" style="background:var(--purple)"></span>${SLBL[h.override_status]||h.override_status} &#9881;</span>`;
    } else {
      badge=`<span class="sbadge ${SCLS[h.display_status]||'s-partial'}"><span class="dot6" style="background:${SDOT[h.display_status]||'#8fa8c8'}"></span>${SLBL[h.display_status]||h.display_status}</span>`;
    }

    const noteHtml=h.note
      ?`<span title="${h.note.replace(/"/g,'&quot;')}" style="cursor:help;font-size:13px">&#128221;</span>`
      :'<span style="color:var(--t3)">&#8212;</span>';

    const esc=s=>s.replace(/'/g,"\\'");
    const acts=h.deleted
      ?`<button class="act-btn res" onclick="restoreHost('${esc(h.hostname)}')" title="Restore host">&#8617; Restore</button>`
      :`<button class="act-btn" onclick="openDrawer('${esc(h.hostname)}')" title="View history">&#128203;</button>
        <button class="act-btn" onclick="openEdit('${esc(h.hostname)}')" title="Override status">&#9998;</button>
        <button class="act-btn del" onclick="openDel('${esc(h.hostname)}')" title="Hide from dashboard">&#128465;</button>`;

    return `<tr class="${h.deleted?'del-row':''}" data-hn="${h.hostname}">
      <td><div class="hn" onclick="copyHn('${esc(h.hostname)}')" title="Click to copy hostname">
        <span class="hn-txt">${h.hostname}</span>
        <span class="cp-icon">&#10098;</span>
      </div></td>
      <td>${og}</td>
      <td>${badge}</td>
      <td style="color:var(--t2)">${rel(h.last_checked)}</td>
      <td style="${daStyle}">${da}</td>
      <td style="color:var(--t3);text-align:center">${h.checks}</td>
      <td style="text-align:center">${rv}</td>
      <td>${trend(h)}</td>
      <td>${noteHtml}</td>
      <td style="text-align:center;white-space:nowrap">${acts}</td>
    </tr>`;
  }).join('');
}

// ── EDIT MODAL ────────────────────────────────────────────────────────────────
function openEdit(hn){
  editHn=hn;
  const h=allHosts.find(x=>x.hostname===hn);
  const o=merged()[hn]||{};
  document.getElementById('e-hn').value       = hn;
  document.getElementById('e-actual').value   = SLBL[h?h.status:'']||'&#8212;';
  document.getElementById('e-override').value = o.override||'';
  document.getElementById('e-note').value     = o.note||'';
  document.getElementById('e-by').value       = o.by||'';
  document.getElementById('edit-ov').classList.remove('hidden');
}
function closeEdit(e){
  if(e&&e.target!==document.getElementById('edit-ov')) return;
  document.getElementById('edit-ov').classList.add('hidden'); editHn=null;
}
function saveEdit(){
  if(!editHn) return;
  const existing=localOvr[editHn]||{};
  localOvr[editHn]={
    ...existing,
    override: document.getElementById('e-override').value||null,
    note:     document.getElementById('e-note').value.trim(),
    by:       document.getElementById('e-by').value.trim(),
    at:       new Date().toISOString().slice(0,16).replace('T',' '),
    deleted:  existing.deleted||false,
  };
  saveLoc(); render(); closeEdit();
  toast('Override saved for '+editHn,'success');
}

// ── DELETE MODAL ──────────────────────────────────────────────────────────────
function openDel(hn){
  delHn=hn;
  document.getElementById('d-hn').textContent=hn;
  document.getElementById('d-reason').value='';
  document.getElementById('del-ov').classList.remove('hidden');
}
function closeDel(e){
  if(e&&e.target!==document.getElementById('del-ov')) return;
  document.getElementById('del-ov').classList.add('hidden'); delHn=null;
}
function confirmDel(){
  if(!delHn) return;
  const existing=localOvr[delHn]||{};
  localOvr[delHn]={
    ...existing,
    deleted:    true,
    del_reason: document.getElementById('d-reason').value.trim(),
    del_at:     new Date().toISOString().slice(0,16).replace('T',' '),
  };
  saveLoc(); closeDel(); render();
  toast(delHn+' hidden from dashboard','warning'); delHn=null;
}
function restoreHost(hn){
  const existing=localOvr[hn]||{};
  localOvr[hn]={...existing, deleted:false};
  saveLoc(); render(); toast(hn+' restored','success');
}

// ── HISTORY DRAWER ────────────────────────────────────────────────────────────
function openDrawer(hn){
  const h=allHosts.find(x=>x.hostname===hn);
  if(!h) return;
  document.getElementById('drw-ttl').textContent=hn;
  document.getElementById('drw-sub').textContent=
    h.checks+' check'+(h.checks!==1?'s':'')+' \u00b7 '+h.revals+' reval'+(h.revals!==1?'s':'')+' \u00b7 OS: '+h.os_group.replace('&#8212;','—');

  const o=merged()[hn]||{};
  let ovSection='';
  if(o.override||o.note){
    ovSection=`<div style="background:var(--bg3);border:1px solid var(--bd);border-radius:6px;padding:12px;margin-bottom:16px">
      <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--t3);margin-bottom:8px">Manual Override</div>
      ${o.override?`<div style="margin-bottom:6px;font-size:12px">Status override: <span class="sbadge s-override" style="margin-left:4px">&#9881; ${SLBL[o.override]||o.override}</span></div>`:''}
      ${o.note?`<div style="color:var(--t2);font-size:12px;font-style:italic;margin-bottom:6px">"${o.note}"</div>`:''}
      ${o.by?`<div style="font-size:11px;color:var(--t3)">By ${o.by} &#183; ${o.at||''}</div>`:''}
    </div>`;
  }

  const hostRecs=ALL.filter(r=>r.host_results.some(hh=>hh.hostname===hn))
    .sort((a,b)=>b.timestamp.localeCompare(a.timestamp));
  const DC={active:'#3dd68c', partial:'#f5a623', not_found:'#f06060'};

  const timeline=hostRecs.map((r,i)=>{
    const hh=r.host_results.find(x=>x.hostname===hn);
    const st=hh?hh.status:'?';
    const da=hh&&hh.days_ago!=null?' \u00b7 Event '+(hh.days_ago===0?'today':hh.days_ago+'d ago'):'';
    return `<div class="tl-item">
      <div class="tl-lc">
        <div class="tl-dot" style="background:${DC[st]||'#4d6480'}"></div>
        ${i<hostRecs.length-1?'<div class="tl-line"></div>':''}
      </div>
      <div class="tl-content">
        <div class="tl-ts">${r.timestamp.replace('T',' ').slice(0,16)}</div>
        <div style="margin:2px 0"><span class="sbadge ${SCLS[st]||'s-partial'}">${SLBL[st]||st}</span>${da}</div>
        <div style="font-size:11px;color:var(--t3)">From: ${r.sender||'(unknown)'}</div>
        ${r.is_revalidation?'<div class="tl-reval">\u27f3 Revalidation run</div>':''}
      </div>
    </div>`;
  }).join('');

  document.getElementById('drw-body').innerHTML =
    ovSection +
    `<div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--t3);margin-bottom:12px">Signoff History (${hostRecs.length} record${hostRecs.length!==1?'s':''})</div>` +
    (timeline||'<div style="color:var(--t3)">No history records found.</div>') +
    `<div class="tl">${timeline}</div>`;

  document.getElementById('drw-ov').classList.remove('hidden');
}
function closeDrawer(){ document.getElementById('drw-ov').classList.add('hidden'); }

// ── OVERRIDES IMPORT / EXPORT ─────────────────────────────────────────────────
function exportOverrides(){
  const blob=new Blob([JSON.stringify(merged(),null,2)],{type:'application/json'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download='signoff_overrides.json'; a.click(); URL.revokeObjectURL(a.href);
  toast('Exported. Save as OVERRIDES_FILE and re-run script to persist.','info');
}
function importOverrides(ev){
  const f=ev.target.files[0]; if(!f) return;
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const imp=JSON.parse(e.target.result);
      if(typeof imp!=='object'||Array.isArray(imp)) throw new Error();
      localOvr=imp; saveLoc(); render();
      toast('Imported '+Object.keys(imp).length+' override(s)','success');
    }catch(err){ toast('Import failed: invalid JSON format','error'); }
  };
  reader.readAsText(f); ev.target.value='';
}
function resetOverrides(){
  if(!confirm('Discard all unsaved changes?')) return;
  localStorage.removeItem(LS_KEY); localOvr={}; render();
  toast('Overrides reset to last saved state','info');
}

// ── UTILS ─────────────────────────────────────────────────────────────────────
function copyHn(hn){
  navigator.clipboard.writeText(hn)
    .then(()=>toast('Copied: '+hn,'info'))
    .catch(()=>toast('Copy not available in this browser','warning'));
}
function toast(msg,type='info'){
  const c=document.getElementById('toast-wrap');
  const t=document.createElement('div'); t.className='toast '+type;
  const icons={success:'&#10003;',warning:'&#9888;',error:'&#10005;',info:'&#8505;'};
  t.innerHTML=`<span>${icons[type]||'&#8505;'}</span><span style="flex:1">${msg}</span>`;
  c.appendChild(t);
  setTimeout(()=>{t.style.opacity='0';t.style.transform='translateY(4px)';t.style.transition='.2s';setTimeout(()=>t.remove(),200);},3500);
}
document.addEventListener('keydown',e=>{
  if(e.key==='Escape'){ closeEdit(); closeDel(); closeDrawer(); }
});

// ── INIT ──────────────────────────────────────────────────────────────────────
render();
</script>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════════════════════
#  LOGGING
# ══════════════════════════════════════════════════════════════════════════════

def _rotate_log_if_needed():
    try:
        if os.path.exists(RUN_LOG_PATH) and \
                os.path.getsize(RUN_LOG_PATH) > LOG_MAX_MB * 1024 * 1024:
            with open(RUN_LOG_PATH, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()
            with open(RUN_LOG_PATH, 'w', encoding='utf-8') as f:
                f.write(f"[Log rotated at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]\n")
                f.write(content[len(content) // 2:])
    except Exception:
        pass


def _log(message):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    print(line)
    try:
        with open(RUN_LOG_PATH, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception as e:
        print(f"WARNING: Log write failed: {e}")


# ══════════════════════════════════════════════════════════════════════════════
#  LOCKFILE — stale lockfile detection via file age
# ══════════════════════════════════════════════════════════════════════════════

def acquire_lock():
    if os.path.exists(LOCKFILE_PATH):
        try:
            age_minutes = (
                datetime.now() - datetime.fromtimestamp(os.path.getmtime(LOCKFILE_PATH))
            ).total_seconds() / 60
        except Exception:
            age_minutes = 0

        if age_minutes < LOCKFILE_STALE_MINUTES:
            _log(f"WARNING: Lockfile exists ({age_minutes:.0f}m old) — another instance may be running. Exiting.")
            return False
        else:
            _log(f"INFO: Stale lockfile detected ({age_minutes:.0f}m old) — removing and continuing.")
            try:
                os.remove(LOCKFILE_PATH)
            except Exception:
                pass

    try:
        with open(LOCKFILE_PATH, 'w') as f:
            f.write(str(os.getpid()))
        return True
    except Exception as e:
        _log(f"ERROR: Could not create lockfile: {e}")
        return False


def release_lock():
    try:
        if os.path.exists(LOCKFILE_PATH):
            os.remove(LOCKFILE_PATH)
    except Exception as e:
        _log(f"WARNING: Could not remove lockfile: {e}")


# ══════════════════════════════════════════════════════════════════════════════
#  DATETIME HELPER
# ══════════════════════════════════════════════════════════════════════════════

def _com_dt_to_py(com_dt):
    """Converts a pywintypes COM datetime to a naive Python datetime (local time)."""
    if com_dt is None:
        return None
    try:
        return datetime(
            com_dt.year, com_dt.month, com_dt.day,
            com_dt.hour, com_dt.minute, com_dt.second,
        )
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════════════════════
#  TRACKING & DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════

def load_tracking():
    if not os.path.exists(TRACKING_FILE):
        return []
    try:
        with open(TRACKING_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        _log(f"WARNING: Could not load tracking file: {e}")
        return []


def append_tracking_record(record):
    """Append-only audit log. Records are never modified or deleted."""
    records = load_tracking()
    records.append(record)
    try:
        with open(TRACKING_FILE, 'w', encoding='utf-8') as f:
            json.dump(records, f, indent=2, ensure_ascii=False)
    except Exception as e:
        _log(f"WARNING: Could not write tracking file: {e}")


def load_overrides():
    """
    Loads manual dashboard overrides (status, note, delete).
    Returns empty dict if file missing or corrupt.
    Overrides are written by the user via the dashboard Export button
    and saved to OVERRIDES_FILE before the next script run.
    """
    if not os.path.exists(OVERRIDES_FILE):
        return {}
    try:
        with open(OVERRIDES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        _log(f"WARNING: Could not load overrides file: {e}")
        return {}


def generate_dashboard():
    """
    Regenerates the self-contained HTML dashboard.
    Embeds all tracking records and current overrides so no server is needed.
    """
    records   = load_tracking()
    overrides = load_overrides()
    try:
        data_json     = json.dumps(records,   ensure_ascii=False)
        overrides_json = json.dumps(overrides, ensure_ascii=False)
        generated     = datetime.now().strftime('%d %B %Y at %H:%M')

        html = _DASHBOARD_HTML
        html = html.replace('__TRACKING_DATA__', data_json)
        html = html.replace('__OVERRIDES_DATA__', overrides_json)
        html = html.replace('__GENERATED__',      generated)

        with open(DASHBOARD_FILE, 'w', encoding='utf-8') as f:
            f.write(html)
        _log(f"Dashboard written: {DASHBOARD_FILE}  ({len(records)} records, "
             f"{len(overrides)} override(s))")
    except Exception as e:
        _log(f"WARNING: Could not write dashboard: {e}")


# ══════════════════════════════════════════════════════════════════════════════
#  QRADAR — HTTP GET HELPER WITH RETRY & RATE-LIMIT BACK-OFF
# ══════════════════════════════════════════════════════════════════════════════

def _api_get(url, params=None):
    """
    GET request with retry on connection errors and 429 rate-limit back-off.
    Returns a Response object or None on total failure.
    """
    for attempt in range(API_RETRIES):
        try:
            resp = requests.get(
                url, params=params,
                auth=(QRADAR_USERNAME, QRADAR_PASSWORD),
                verify=VERIFY_SSL, timeout=REQUEST_TIMEOUT,
                headers={'Accept': 'application/json', 'Version': '14.0'}
            )
            if resp.status_code == 429:
                wait = API_BACKOFF_SECONDS * (2 ** attempt)
                _log(f"WARNING: QRadar rate limited (429) — waiting {wait:.0f}s (attempt {attempt+1}/{API_RETRIES})")
                time.sleep(wait)
                continue
            return resp
        except requests.exceptions.ConnectionError:
            if attempt < API_RETRIES - 1:
                time.sleep(API_BACKOFF_SECONDS)
                continue
            _log(f"ERROR: Connection error after {API_RETRIES} attempts: {url}")
            return None
        except Exception as e:
            _log(f"ERROR: Unexpected API error ({url}): {e}")
            return None
    return None


def test_qradar_connection():
    _log("Testing QRadar connection...")
    resp = _api_get(f"{QRADAR_HOST.rstrip('/')}/api/help/versions")
    if resp is None:
        _log("ERROR: QRadar unreachable after retries.")
        return False
    if resp.status_code == 200:
        _log("QRadar connection OK.")
        return True
    if resp.status_code == 401:
        _log("ERROR: Authentication failed. Check QRADAR_USERNAME / QRADAR_PASSWORD env vars.")
        return False
    _log(f"WARNING: Unexpected response HTTP {resp.status_code}")
    return False


def fetch_log_source_types():
    _log("Fetching Log Source Types...")
    resp = _api_get(
        f"{QRADAR_HOST.rstrip('/')}/api/config/event_sources"
        f"/log_source_management/log_source_types"
    )
    if resp and resp.status_code == 200:
        for t in resp.json():
            ls_id, ls_name = t.get('id'), t.get('name')
            if ls_id is not None and ls_name is not None:
                LOG_SOURCE_TYPES_CACHE[ls_id] = ls_name
        _log(f"Cached {len(LOG_SOURCE_TYPES_CACHE)} Log Source Types.")
    else:
        _log(f"WARNING: Failed to fetch Log Source Types: {resp.status_code if resp else 'no response'}")


# ══════════════════════════════════════════════════════════════════════════════
#  QRADAR QUERIES — STRICTLY READ-ONLY (HTTP GET ONLY)
# ══════════════════════════════════════════════════════════════════════════════

def _safe_timestamp(timestamp_ms):
    if not timestamp_ms:
        return 'No events recorded', 'No Activity', None
    try:
        if isinstance(timestamp_ms, float):
            timestamp_ms = int(timestamp_ms)
        epoch_s = timestamp_ms / 1000.0 if timestamp_ms > 4102444800 else timestamp_ms
        if epoch_s <= _MIN_TS or epoch_s > _MAX_TS:
            return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None
        last_event_dt = datetime.fromtimestamp(epoch_s)
        days_ago      = (datetime.now() - last_event_dt).days
        activity      = 'Active' if days_ago <= ACTIVITY_THRESHOLD_DAYS else 'Inactive'
        return last_event_dt.strftime('%Y-%m-%d %H:%M:%S'), activity, days_ago
    except Exception:
        return f'Invalid timestamp: {timestamp_ms}', 'Unknown', None


def query_all_log_sources_readonly(hostname):
    """
    READ-ONLY — fetches every log source matching hostname.
    Only HTTP GET via _api_get() which includes retry logic.
    Nothing in QRadar is modified, created, or deleted.
    """
    clean = str(hostname).replace('"', '').replace("'", "").strip()
    resp  = _api_get(
        f"{QRADAR_HOST.rstrip('/')}/api/config/event_sources"
        f"/log_source_management/log_sources",
        params={'filter': f'name ilike "%{clean}%"'}
    )
    if resp is None:
        return {'status': 'Connection Error', 'sources': []}
    if resp.status_code != 200:
        return {'status': f'API Error {resp.status_code}', 'sources': []}
    ls_data = resp.json()
    if not ls_data:
        return {'status': 'Not Found', 'sources': []}
    sources = []
    for src in ls_data:
        type_id      = src.get('type_id')
        ls_type_name = LOG_SOURCE_TYPES_CACHE.get(type_id, f'Unknown Type ID: {type_id}')
        last_seen, activity, days_ago = _safe_timestamp(src.get('last_event_time'))
        sources.append({
            'name':      src.get('name', hostname),
            'ls_type':   ls_type_name,
            'enabled':   src.get('enabled', False),
            'last_seen': last_seen,
            'activity':  activity,
            'days_ago':  days_ago,
        })
    return {'status': 'Found', 'sources': sources}


def validate_expected_types(all_sources_result, required_types):
    results = []
    sources = all_sources_result.get('sources', [])
    for expected_kw in required_types:
        exp_words = str(expected_kw).lower().split()
        matched   = [
            s for s in sources
            if all(w in str(s.get('ls_type', '')).lower() for w in exp_words)
        ]
        if not matched:
            results.append({'expected': expected_kw, 'found': False,
                            'ls_type': None, 'ls_name': None,
                            'last_seen': None, 'days_ago': None})
            continue
        me = sorted([s for s in matched if s.get('enabled')],
                    key=lambda x: x.get('days_ago') or 99999)
        md = sorted([s for s in matched if not s.get('enabled')],
                    key=lambda x: x.get('days_ago') or 99999)
        best = me[0] if me else md[0]
        results.append({'expected': expected_kw, 'found': True,
                        'ls_type': best.get('ls_type'), 'ls_name': best.get('name'),
                        'last_seen': best.get('last_seen'), 'days_ago': best.get('days_ago')})
    return results


def detect_os_group(sources):
    if not OS_TYPE_GROUPS:
        return None, None
    for group_name, rules in OS_TYPE_GROUPS.items():
        required = rules.get('required', [])
        if not required:
            continue
        sig_words = str(required[0]).lower().split()
        if any(
            all(w in str(s.get('ls_type', '')).lower() for w in sig_words)
            for s in sources
        ):
            return group_name, rules
    return None, None


# ══════════════════════════════════════════════════════════════════════════════
#  SUBJECT PARSING — MULTI-HOSTNAME WITH DEDUPLICATION
# ══════════════════════════════════════════════════════════════════════════════

def passes_subject_guards(subject):
    """
    '[processed' (no closing bracket) catches all outcome tag variants:
    [Processed-Active], [Processed-Partial], [Processed-NotFound].
    Using '[processed]' with a closing bracket would miss these — fixed.
    """
    if not subject:
        return False, "empty subject"
    s, sl = subject.strip(), subject.strip().lower()
    if any(sl.startswith(p) for p in ('re:', 'fw:', 'fwd:')):
        return False, f"reply/forward prefix: '{s[:30]}'"
    if '[processed' in sl:
        return False, "subject already bears an outcome tag"
    if SUBJECT_SEPARATOR not in s:
        return False, f"separator '{SUBJECT_SEPARATOR}' not found"
    left = s.split(SUBJECT_SEPARATOR)[0].strip().lower()
    if SUBJECT_KEYWORD.lower() not in left:
        return False, f"keyword '{SUBJECT_KEYWORD}' not found left of separator"
    return True, "ok"


def extract_hostnames(subject):
    """
    Parses hostnames from right of separator. Deduplicates preserving order.

    "Security Signoff | HOST1 | HOST2"        -> ['HOST1', 'HOST2']
    "Security Signoff | HOST1 | HOST1 | HOST2" -> ['HOST1', 'HOST2']  (deduped)
    """
    parts = subject.split(SUBJECT_SEPARATOR)
    seen, result = set(), []
    for p in parts[1:]:
        h = p.strip()
        if h and h.lower() not in seen:
            seen.add(h.lower())
            result.append(h)
    return result


# ══════════════════════════════════════════════════════════════════════════════
#  SENDER VALIDATION
# ══════════════════════════════════════════════════════════════════════════════

def is_sender_allowed(sender_address):
    if not ALLOWED_SENDERS:
        return True
    if not sender_address:
        return False
    sc = sender_address.strip().lower()
    for entry in ALLOWED_SENDERS:
        ec = entry.strip().lower()
        if ec.startswith('@') and sc.endswith(ec):
            return True
        if ec == sc:
            return True
    return False


# ══════════════════════════════════════════════════════════════════════════════
#  CONVERSATION STATUS — state machine via Outlook subject tags
# ══════════════════════════════════════════════════════════════════════════════

def check_conversation_status(mail_item, sent_folder, drafts_folder):
    """
    Scans Sent Items (restricted to revalidation window) and Drafts for the
    most recent outcome tag in this conversation thread.

    Returns (tag, datetime) or (None, None).
    """
    conv_id  = mail_item.ConversationID
    last_tag = None
    last_dt  = None

    def _extract_tag(subject):
        s = (subject or '').lower()
        if TAG_NOT_FOUND.lower() in s: return TAG_NOT_FOUND
        if TAG_PARTIAL.lower()   in s: return TAG_PARTIAL
        if TAG_ACTIVE.lower()    in s: return TAG_ACTIVE
        if '[processed]'         in s: return 'legacy'
        return None

    def _update(tag, item_dt):
        nonlocal last_tag, last_dt
        if tag and (last_dt is None or (item_dt and item_dt > last_dt)):
            last_tag, last_dt = tag, item_dt

    cutoff_str = (datetime.now() - timedelta(days=REVALIDATION_WINDOW_DAYS)
                  ).strftime('%m/%d/%Y %I:%M %p')
    try:
        for item in sent_folder.Items.Restrict(f"[SentOn] >= '{cutoff_str}'"):
            try:
                if item.ConversationID == conv_id:
                    _update(_extract_tag(item.Subject), _com_dt_to_py(item.SentOn))
            except Exception:
                continue
    except Exception as e:
        _log(f"WARNING: Could not scan Sent Items: {e}")

    try:
        for item in drafts_folder.Items:
            try:
                if item.ConversationID == conv_id:
                    _update(_extract_tag(item.Subject),
                            _com_dt_to_py(item.LastModificationTime))
            except Exception:
                continue
    except Exception as e:
        _log(f"WARNING: Could not scan Drafts: {e}")

    return last_tag, last_dt


# ══════════════════════════════════════════════════════════════════════════════
#  BODY DL CHECK
# ══════════════════════════════════════════════════════════════════════════════

def body_contains_dl(mail_item):
    if not TRIGGER_DL.strip():
        return True
    dl_lower = TRIGGER_DL.strip().lower()
    try:
        if dl_lower in (mail_item.Body or '').lower():
            return True
        if dl_lower in (mail_item.HTMLBody or '').lower():
            return True
        return False
    except Exception as e:
        _log(f"WARNING: Body DL check failed: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
#  HTML REPLY BUILDER — MULTI-HOSTNAME
# ══════════════════════════════════════════════════════════════════════════════

def _status_for_host(hostname, qr_result):
    """Returns (host_status, type_validation, os_group) for one hostname."""
    status  = qr_result.get('status')
    sources = qr_result.get('sources', [])

    if status != 'Found' or not sources:
        return 'not_found', None, None

    if not OS_TYPE_GROUPS:
        return 'active', None, None

    group_name, group_rules = detect_os_group(sources)
    if group_name is None:
        _log(f"         OS group undetected for {hostname} — found but no type validation applied")
        return 'active', None, None

    validation  = validate_expected_types(qr_result, group_rules.get('required', []))
    any_problem = any(
        not r['found'] or (r['found'] and r['days_ago'] is None)
        for r in validation
    )
    return ('partial' if any_problem else 'active'), validation, group_name


def _host_section_html(hostname, host_status, type_validation, os_group, sources):
    """Builds inline HTML section for one hostname in the multi-host reply email."""
    section_color = {'active': '#1a7a4a', 'partial': '#c87800',
                     'not_found': '#c0392b'}.get(host_status, '#555')
    os_label = f' &mdash; {os_group}' if os_group else ''

    header = (
        f'<div style="margin:18px 0 6px;padding:7px 14px;'
        f'background:{section_color}12;border-left:3px solid {section_color};'
        f'border-radius:0 6px 6px 0;">'
        f'<span style="font-family:monospace;font-weight:700;font-size:12px;color:{section_color}">'
        f'{hostname}</span>'
        f'<span style="color:#999;font-size:11px;">{os_label}</span>'
        f'</div>'
    )

    if host_status == 'not_found':
        return (header +
                f'<p style="font-size:12px;color:#c0392b;margin:4px 0 0 14px;">'
                f'Not found in QRadar log source inventory.</p>')

    if type_validation is not None:
        rows = ''
        for r in type_validation:
            days_str = (
                f" <span style='color:#888;font-size:10px;'>"
                f"({'Today' if r['days_ago']==0 else str(r['days_ago'])+'d ago'})</span>"
                if r['days_ago'] is not None else ''
            )
            if not r['found']:
                note = (f"{ESCALATION_CONTACT} please onboard this log source."
                        if ESCALATION_CONTACT.strip() else
                        "Not found — please onboard this log source.")
                rows += (
                    f'<tr style="background:#fff5f5"><td style="color:#c0392b;font-weight:700;'
                    f'text-align:center;padding:7px 10px;border-bottom:1px solid #ffe0e0;'
                    f'font-size:11px;width:20px">&#10006;</td>'
                    f'<td style="color:#c0392b;font-weight:600;padding:7px 10px;'
                    f'border-bottom:1px solid #ffe0e0;font-size:11px">{r["expected"]}</td>'
                    f'<td style="color:#c0392b;padding:7px 10px;border-bottom:1px solid #ffe0e0;'
                    f'font-size:11px">&#8212;</td>'
                    f'<td style="color:#c0392b;padding:7px 10px;border-bottom:1px solid #ffe0e0;'
                    f'font-size:11px">&#8212;</td>'
                    f'<td style="color:#c0392b;font-style:italic;padding:7px 10px;'
                    f'border-bottom:1px solid #ffe0e0;font-size:11px">{note}</td></tr>'
                )
            elif r['days_ago'] is None:
                note = (f"{ESCALATION_CONTACT} no events received yet — please investigate."
                        if ESCALATION_CONTACT.strip() else
                        "No events received yet — please investigate.")
                rows += (
                    f'<tr style="background:#fffbf0"><td style="color:#c87800;font-weight:700;'
                    f'text-align:center;padding:7px 10px;border-bottom:1px solid #ffefc0;'
                    f'font-size:11px;width:20px">&#9888;</td>'
                    f'<td style="color:#c87800;font-weight:600;padding:7px 10px;'
                    f'border-bottom:1px solid #ffefc0;font-size:11px">{r["expected"]}</td>'
                    f'<td style="color:#555;padding:7px 10px;border-bottom:1px solid #ffefc0;'
                    f'font-size:11px">{r.get("ls_name","N/A")}</td>'
                    f'<td style="color:#c87800;font-style:italic;padding:7px 10px;'
                    f'border-bottom:1px solid #ffefc0;font-size:11px">No events recorded</td>'
                    f'<td style="color:#c87800;font-style:italic;padding:7px 10px;'
                    f'border-bottom:1px solid #ffefc0;font-size:11px">{note}</td></tr>'
                )
            else:
                rows += (
                    f'<tr style="background:#f0faf4"><td style="color:#1a7a4a;font-weight:700;'
                    f'text-align:center;padding:7px 10px;border-bottom:1px solid #d0f0e0;'
                    f'font-size:11px;width:20px">&#10004;</td>'
                    f'<td style="color:#333;font-weight:600;padding:7px 10px;'
                    f'border-bottom:1px solid #d0f0e0;font-size:11px">{r["expected"]}</td>'
                    f'<td style="color:#555;padding:7px 10px;border-bottom:1px solid #d0f0e0;'
                    f'font-size:11px">{r.get("ls_name","N/A")}</td>'
                    f'<td style="color:#555;padding:7px 10px;border-bottom:1px solid #d0f0e0;'
                    f'font-size:11px">{r.get("last_seen","N/A")}{days_str}</td>'
                    f'<td style="color:#1a7a4a;font-weight:600;padding:7px 10px;'
                    f'border-bottom:1px solid #d0f0e0;font-size:11px">Confirmed</td></tr>'
                )
        tbl = (
            f'<table style="width:100%;border-collapse:collapse;margin:4px 0 0;'
            f'border:1px solid #e0e0e0;border-radius:4px;overflow:hidden;font-size:11px">'
            f'<tr style="background:#f5f5f5">'
            f'<th style="padding:6px 10px;border-bottom:2px solid #ddd;text-align:left;width:20px"></th>'
            f'<th style="padding:6px 10px;border-bottom:2px solid #ddd;text-align:left">Log Source Type</th>'
            f'<th style="padding:6px 10px;border-bottom:2px solid #ddd;text-align:left">Log Source Name</th>'
            f'<th style="padding:6px 10px;border-bottom:2px solid #ddd;text-align:left">Last Event</th>'
            f'<th style="padding:6px 10px;border-bottom:2px solid #ddd;text-align:left">Status</th>'
            f'</tr>{rows}</table>'
        )
        return header + tbl

    # Simple mode — no type validation
    best = None
    if sources:
        en = sorted([s for s in sources if s.get('enabled')],
                    key=lambda x: x.get('days_ago') or 99999)
        di = sorted([s for s in sources if not s.get('enabled')],
                    key=lambda x: x.get('days_ago') or 99999)
        best = en[0] if en else (di[0] if di else None)

    if best:
        da  = best.get('days_ago')
        dsp = 'Today' if da == 0 else (f"{da}d ago" if da is not None else 'N/A')
        return header + (
            f'<table style="width:100%;max-width:480px;border-collapse:collapse;'
            f'margin:4px 0 0;border:1px solid #e0e0e0;border-radius:4px;overflow:hidden">'
            f'<tr><td style="padding:6px 12px;color:#555;font-size:12px;'
            f'border-bottom:1px solid #eee;width:140px">Log Source Name</td>'
            f'<td style="padding:6px 12px;font-size:12px;border-bottom:1px solid #eee;'
            f'font-weight:600">{best.get("name","N/A")}</td></tr>'
            f'<tr><td style="padding:6px 12px;color:#555;font-size:12px;'
            f'border-bottom:1px solid #eee">Log Source Type</td>'
            f'<td style="padding:6px 12px;font-size:12px;'
            f'border-bottom:1px solid #eee">{best.get("ls_type","N/A")}</td></tr>'
            f'<tr><td style="padding:6px 12px;color:#555;font-size:12px">Last Event</td>'
            f'<td style="padding:6px 12px;font-size:12px">{best.get("last_seen","N/A")} '
            f'<span style="color:#888;font-size:11px">({dsp})</span></td></tr>'
            f'</table>'
        )
    return header


def build_reply_for_all_hosts(hostname_qr_pairs):
    """
    Processes all (hostname, qr_result) pairs and builds the combined reply.

    Returns: (html_body, overall_status, host_tracking)

    overall_status : worst-case across all hosts
        'active'    -> ReplyAll to original thread
        'partial'   -> ESCALATION_TO / ESCALATION_CC only
        'not_found' -> ESCALATION_TO / ESCALATION_CC only
    """
    STATUS_RANK  = {'not_found': 0, 'partial': 1, 'active': 2}
    run_time     = datetime.now().strftime('%d %B %Y, %H:%M')
    host_sections = []
    host_tracking = []
    statuses      = []

    for hostname, qr_result in hostname_qr_pairs:
        host_status, type_validation, os_group = _status_for_host(hostname, qr_result)
        sources = qr_result.get('sources', [])
        statuses.append(host_status)

        host_sections.append(
            _host_section_html(hostname, host_status, type_validation,
                               os_group, sources)
        )

        best_da, best_seen = None, None
        if sources:
            active_with_events = [
                s for s in sources
                if s.get('enabled') and s.get('days_ago') is not None
            ]
            if active_with_events:
                best      = min(active_with_events, key=lambda x: x['days_ago'])
                best_da   = best['days_ago']
                best_seen = best['last_seen']
        host_tracking.append({
            'hostname': hostname,
            'status':   host_status,
            'os_group': os_group,
            'last_seen': best_seen,
            'days_ago':  best_da,
        })

    # Worst-case overall status; guard against empty list
    overall_status = min(statuses, key=lambda s: STATUS_RANK.get(s, 1)) \
        if statuses else 'not_found'

    n_hosts  = len(hostname_qr_pairs)
    n_ok     = sum(1 for s in statuses if s == 'active')
    n_issues = n_hosts - n_ok

    BANNERS = {
        'active':    ('#1a7a4a', '&#10004;&nbsp; All Hosts Confirmed Reporting on SIEM'),
        'partial':   ('#c87800', f'&#9888;&nbsp; {n_issues} of {n_hosts} '
                      f'Host{"s" if n_hosts > 1 else ""} Require Attention'),
        'not_found': ('#c0392b', f'&#10006;&nbsp; {"Some hosts not" if n_ok else "No hosts"} '
                      f'found in QRadar'),
    }
    banner_color, banner_label = BANNERS.get(overall_status, BANNERS['partial'])

    if overall_status == 'active' and n_hosts == 1:
        summary = f'<b>{hostname_qr_pairs[0][0]}</b> is confirmed reporting on our SIEM.'
    elif overall_status == 'active':
        names   = ', '.join(f'<b>{h}</b>' for h, _ in hostname_qr_pairs)
        summary = f'All {n_hosts} requested hosts ({names}) are confirmed reporting on our SIEM.'
    else:
        summary = (f'{n_ok} of {n_hosts} host{"s" if n_hosts>1 else ""} confirmed active. '
                   f'Issues are highlighted below.')

    body = f"""
<html>
<body style="font-family:'Segoe UI',Arial,sans-serif;color:#222;font-size:13px;
             line-height:1.6;margin:0;padding:0;">
  <div style="max-width:680px;padding:20px 0;">
    <p style="margin:0 0 16px 0;">Hi,</p>
    <div style="background:{banner_color};color:#fff;padding:10px 16px;
                border-radius:6px;font-size:13px;font-weight:600;
                margin-bottom:12px;letter-spacing:0.2px;">
      {banner_label}
    </div>
    <p style="margin:0 0 4px 0;">{summary}</p>
    {''.join(host_sections)}
    <p style="margin:20px 0 4px 0;color:#555;font-size:12px;">
      This is an automated response from the SIEM monitoring system.<br>
      Checked against QRadar on {run_time}.
    </p>
    <p style="margin:16px 0 0 0;">Regards,<br>
    <span style="font-weight:600;">Cyberdefence</span></p>
  </div>
</body>
</html>"""

    return body, overall_status, host_tracking


# ══════════════════════════════════════════════════════════════════════════════
#  DRAFT CREATOR — SIMPLIFIED ROUTING
# ══════════════════════════════════════════════════════════════════════════════

def create_draft_reply(mail_item, html_body, overall_status,
                       hostnames_str, is_revalidation=False):
    """
    Creates and saves a draft reply. Routing driven by overall_status.

        'active'              -> ReplyAll, original recipients unchanged.
        'partial'/'not_found' -> To = ESCALATION_TO, CC = ESCALATION_CC only.

    The outcome tag written to the subject is the permanent state record used
    by check_conversation_status() on all future runs.

    THIS IS DRAFT ONLY. reply.Save() is called. reply.Send() is NEVER called.
    """
    tag_map = {'active': TAG_ACTIVE, 'partial': TAG_PARTIAL, 'not_found': TAG_NOT_FOUND}
    tag     = tag_map.get(overall_status, TAG_ACTIVE)
    prefix  = '[Revalidated] ' if is_revalidation else ''

    try:
        reply          = mail_item.ReplyAll()
        reply.HTMLBody = html_body
        reply.Subject  = f"{prefix}{tag} {mail_item.Subject}"

        if overall_status in ('partial', 'not_found'):
            if not ESCALATION_TO and not ESCALATION_CC:
                _log("         WARNING: ESCALATION_TO and ESCALATION_CC both empty "
                     "— falling back to ReplyAll for this escalation draft.")
            else:
                reply.To = '; '.join(ESCALATION_TO) if ESCALATION_TO else ''
                reply.CC = '; '.join(ESCALATION_CC) if ESCALATION_CC else ''
                _log(f"         Escalation routing  To: {reply.To or '(none)'}  "
                     f"CC: {reply.CC or '(none)'}")
        else:
            _log("         ReplyAll to original thread (Active)")

        reply.Save()
        _log(f"         Draft saved [{tag}]{' [revalidation]' if is_revalidation else ''}"
             f" — {hostnames_str}")
        return True

    except Exception as e:
        _log(f"         ERROR: Draft creation failed for {hostnames_str}: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
#  OUTLOOK SETUP
# ══════════════════════════════════════════════════════════════════════════════

def get_outlook_folders():
    try:
        outlook    = win32com.client.Dispatch('Outlook.Application')
        ns         = outlook.GetNamespace('MAPI')
        main_inbox = ns.GetDefaultFolder(6)    # 6  = Inbox
        drafts     = ns.GetDefaultFolder(16)   # 16 = Drafts
        sent       = ns.GetDefaultFolder(5)    # 5  = Sent Items

        if SIGNOFF_FOLDER_NAME:
            try:
                inbox = main_inbox.Folders[SIGNOFF_FOLDER_NAME]
                _log(f"Monitoring: Inbox\\{SIGNOFF_FOLDER_NAME}")
            except Exception:
                _log(f"WARNING: Subfolder '{SIGNOFF_FOLDER_NAME}' not found — "
                     f"falling back to full Inbox.")
                inbox = main_inbox
        else:
            inbox = main_inbox
            _log("Monitoring: Full Inbox")

        return inbox, drafts, sent

    except Exception as e:
        _log(f"ERROR: Could not connect to Outlook: {e}")
        return None, None, None


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    if not VERIFY_SSL:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    _rotate_log_if_needed()

    _log("=" * 60)
    _log("QRadar Signoff Auto-Draft starting...")
    _log(f"   Normal window   : {LOOKBACK_HOURS}h")
    _log(f"   Reval window    : {REVALIDATION_WINDOW_DAYS}d  "
         f"Cooldown: {REVALIDATION_COOLDOWN_DAYS}d")
    _log(f"   Trigger DL      : '{TRIGGER_DL}'")
    _log(f"   Escalation To   : {ESCALATION_TO}")
    _log(f"   Escalation CC   : {ESCALATION_CC}")
    _log(f"   Dashboard       : {DASHBOARD_FILE}")
    _log(f"   Overrides       : {OVERRIDES_FILE}")
    _log(f"   MODE            : DRAFT ONLY — nothing is sent automatically")

    if not acquire_lock():
        return

    try:
        inbox, drafts, sent = get_outlook_folders()
        if inbox is None:
            return

        if not test_qradar_connection():
            _log("ERROR: QRadar unreachable — exiting. All emails left untouched.")
            return

        fetch_log_source_types()

        scan_hours      = max(LOOKBACK_HOURS, REVALIDATION_WINDOW_DAYS * 24)
        normal_cutoff   = datetime.now() - timedelta(hours=LOOKBACK_HOURS)
        cooldown_cutoff = datetime.now() - timedelta(days=REVALIDATION_COOLDOWN_DAYS)
        cutoff_str      = (datetime.now() - timedelta(hours=scan_hours)
                           ).strftime('%m/%d/%Y %I:%M %p')

        inbox_items = list(inbox.Items.Restrict(f"[ReceivedTime] >= '{cutoff_str}'"))
        _log(f"\nScanned {len(inbox_items)} email(s) in last {scan_hours // 24}d window.")

        processed = skipped = drafted = revalidated = 0

        for mail_item in inbox_items:
            try:
                if mail_item.Class != 43:    # 43 = olMail
                    continue
            except Exception:
                continue

            subject = ''
            try:
                subject = mail_item.Subject or ''
            except Exception:
                continue

            passed, reason = passes_subject_guards(subject)
            if not passed:
                skipped += 1
                _log(f"   SKIP (subject — {reason}): '{subject[:60]}'")
                continue

            try:
                sender = mail_item.SenderEmailAddress or ''
            except Exception:
                sender = ''

            if sender.strip().lower() == YOUR_EMAIL_ADDRESS.strip().lower():
                skipped += 1
                continue

            if not is_sender_allowed(sender):
                skipped += 1
                _log(f"   SKIP (sender not allowed — {sender}): '{subject[:60]}'")
                continue

            if not body_contains_dl(mail_item):
                skipped += 1
                _log(f"   SKIP ('{TRIGGER_DL}' not in body): '{subject[:60]}'")
                continue

            # Multi-hostname extraction with deduplication
            hostnames = extract_hostnames(subject)
            if not hostnames:
                skipped += 1
                _log(f"   SKIP (no hostnames found): '{subject[:60]}'")
                continue

            received_dt         = _com_dt_to_py(mail_item.ReceivedTime)
            is_in_normal_window = received_dt is not None and received_dt >= normal_cutoff

            _log(f"\nCandidate: '{subject[:70]}'")
            _log(f"   Sender : {sender}")
            _log(f"   Hosts  : {hostnames}")
            _log(f"   Window : {'normal' if is_in_normal_window else 'revalidation candidate'}")

            last_tag, last_dt = check_conversation_status(mail_item, sent, drafts)

            if last_tag is None:
                if not is_in_normal_window:
                    skipped += 1
                    _log(f"   SKIP (old email, never processed — out of scope)")
                    continue
                is_revalidation = False
                _log(f"   -> New signoff")

            elif last_tag in (TAG_ACTIVE, 'legacy'):
                skipped += 1
                last_dt_str = last_dt.strftime('%Y-%m-%d') if last_dt else 'unknown'
                _log(f"   SKIP (previously Active on {last_dt_str} — permanent skip)")
                continue

            elif last_tag in REVALIDATABLE_TAGS:
                if last_dt and last_dt >= cooldown_cutoff:
                    days_ago = (datetime.now() - last_dt).days
                    skipped += 1
                    _log(f"   SKIP (cooldown — last checked {days_ago}d ago, "
                         f"cooldown is {REVALIDATION_COOLDOWN_DAYS}d)")
                    continue
                last_dt_str = last_dt.strftime('%Y-%m-%d') if last_dt else 'unknown'
                _log(f"   -> Revalidating {last_tag} from {last_dt_str}")
                is_revalidation = True

            else:
                skipped += 1
                _log(f"   SKIP (unrecognised tag state: '{last_tag}')")
                continue

            # QRadar query — one call per hostname
            _log(f"   Querying QRadar for {len(hostnames)} host(s)...")
            hostname_qr_pairs = []
            for hn in hostnames:
                qr = query_all_log_sources_readonly(hn)
                _log(f"      {hn}: {qr['status']} "
                     f"({len(qr.get('sources', []))} sources)")
                hostname_qr_pairs.append((hn, qr))

            body, overall_status, host_tracking = build_reply_for_all_hosts(
                hostname_qr_pairs
            )
            _log(f"   Overall: {overall_status.upper()}")

            hostnames_str = ' | '.join(hostnames)
            success = create_draft_reply(
                mail_item, body, overall_status,
                hostnames_str=hostnames_str,
                is_revalidation=is_revalidation,
            )

            if success:
                drafted += 1
                if is_revalidation:
                    revalidated += 1
                append_tracking_record({
                    'id':               str(uuid.uuid4()),
                    'timestamp':        datetime.now().isoformat(timespec='seconds'),
                    'subject':          subject,
                    'sender':           sender,
                    'is_revalidation':  is_revalidation,
                    'overall_status':   overall_status,
                    'host_results':     host_tracking,
                })

            processed += 1

        generate_dashboard()

        _log(f"\n{'=' * 60}")
        _log(f"Run complete.")
        _log(f"   Processed  : {processed}  Drafted: {drafted} "
             f"({revalidated} revalidation{'s' if revalidated != 1 else ''})")
        _log(f"   Skipped    : {skipped}")
        _log(f"   Dashboard  : {DASHBOARD_FILE}")
        _log(f"{'=' * 60}\n")

    finally:
        release_lock()


if __name__ == '__main__':
    main()
