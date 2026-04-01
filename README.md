# QRadar Log Source Inventory Validator

Repository of operational QRadar scripts used to keep log-source coverage healthy, reduce blind spots, and surface high-impact governance issues for SOC and leadership.

## What this repository is trying to achieve

This repo is focused on one outcome: **confidence that QRadar is ingesting the right data, at the right volume, with useful rules behind it**.

It does that by combining:

- **Inventory validation** (is each documented source present and active?)
- **Operational hygiene checks** (what’s stale, missing, or erroring?)
- **Governance visibility** (who is burning EPS license and which rules are dead/noisy?)
- **Stakeholder reporting/signoff** (email-style closure workflow)

## Core scripts

### 1) `Qradar.py` — Main validator

Primary inventory validator that compares an Excel inventory against QRadar log source configuration and activity.

Typical outputs:
- Updated source status data
- Highlighted inactive / not found / API-error records
- Filtered issue workbook for action tracking

---

### 2) `New_check.py` — New checker

A streamlined checker variant for fast validation runs with cleaner activity interpretation.

Use this when you want:
- A lighter check flow
- Fast “what changed?” visibility
- A practical daily/weekly hygiene pass

---

### 3) `Signoff.py` — Signoff workflow

Supports operational signoff/reply style workflows (including Outlook-driven processes where available) so teams can close findings and share status quickly.

Use this when you need:
- Communication-friendly reporting
- A simple closure/signoff cycle for log source reviews

---

### 4) `eps_burn_rate_monitor.py` — EPS burn-rate monitor (new)

Runs AQL analytics to rank log sources by EPS consumption and estimate license risk trajectory.

What it provides:
- Ranked EPS consumers by log source
- Trend direction indicators
- “Days until cap” style projection (when cap is configured)

Output:
- `eps_burn_rate_report.xlsx` (`Summary`, `Source_Ranking`)

---

### 5) `rule_effectiveness_auditor.py` — Rule effectiveness auditor (new)

Pulls enabled rules and correlates with fire-count data to expose:
- **Dead rules** (0 fires in lookback)
- **Noise generators** (high fire volume likely ignored operationally)

Output:
- `rule_effectiveness_audit.xlsx`
  - `Summary`
  - `All_Enabled_Rules`
  - `Dead_Rules`
  - `Noise_Generators`

## Why these checks matter together

The strongest story this repo tells is the combination of:

- **EPS burn by source** + **rule effectiveness**

This helps answer leadership questions like:

- “Which sources consume the most license?”
- “Are those sources feeding rules that actually detect incidents?”
- “Are we paying for ingestion that creates mostly noise?”

## Requirements

- Python 3.6+
- Network/API access to QRadar
- Valid QRadar credentials with read access
- Optional: Outlook integration on Windows for signoff/report workflows

Install common dependencies:

```bash
pip install pandas requests urllib3 pywin32 numpy openpyxl matplotlib
```

## Quick usage

From repository root:

```bash
python Qradar.py
python New_check.py
python Signoff.py
python eps_burn_rate_monitor.py
python rule_effectiveness_auditor.py
```

Each script has a config section near the top where you set:
- `QRADAR_HOST`
- `QRADAR_USERNAME`
- `QRADAR_PASSWORD`
- Input/output paths and thresholds

## Validation approach in this repo

This repository is script-first (no formal test suite is currently bundled).  
For quick validation after edits:

```bash
python -m py_compile Qradar.py New_check.py Signoff.py eps_burn_rate_monitor.py rule_effectiveness_auditor.py
```

## Repository structure (key files)

- `Qradar.py` — validator
- `New_check.py` — new checker
- `Signoff.py` — signoff workflow
- `eps_burn_rate_monitor.py` — EPS risk monitoring
- `rule_effectiveness_auditor.py` — rule-value/noise auditing
- `New_Radar.py`, `fun.py`, `latest_r.py` — additional variants/utilities

---

If you want, I can also add a short **“which script should I run?” decision table** for SOC analysts and management audiences.
