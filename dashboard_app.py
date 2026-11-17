#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import html
import json
import os
import re
import sys
import threading
import traceback
import webbrowser
from datetime import datetime
from pathlib import Path
from tkinter import (
    Tk, StringVar, BooleanVar, Text, END, DISABLED, NORMAL, filedialog, messagebox
)
from tkinter import ttk

import pandas as pd

NEWLINE = chr(10)
TAB = chr(9)

CARRIAGE_RETURN = chr(13)

DEFAULT_SHEET = "query 2"
DEFAULT_PREVIOUS_SHEET = "query"
MAX_ASK_LENGTH = 600
DASHBOARD_TITLE = "Action Tracker Dashboard"
APP_TITLE = "Action Tracker Dashboard Generator"
BG_COLOR = "#1E2A47"
ACCENT_COLOR = "#5EEAD4"
TEXT_COLOR = "#FFFFFF"

SOURCE_GLOBAL_PM = "GLOBAL_PM"
SOURCE_ACTION_LOG = "ACTION_LOG"

GLOBAL_PM_COLUMN_MAP = {
    "project":      "Project Name",
    "wbs":          "WBS - Equipment",
    "ask_raw":      "Project Ask",
    "impact":       "Impact",
    "due_date":     "Due Date",
    "owner_raw":    "Owner",
    "forecast":     "Forecast Conclusion Date",
    "status":       "Status",
    "plan_raw":     "Action Plan Detail",
    "days_open":    "L2 Days Open",
    "level":        "Escalation Level",
    "created_at":   "L2 Created",
    "item_type":    "Item Type",
    "path":         "Path",
}

GLOBAL_PM_REQUIRED_KEYS = [
    "project", "wbs", "ask_raw", "impact", "due_date", "owner_raw",
    "forecast", "status", "plan_raw", "days_open", "level", "created_at",
]

ACTION_LOG_COLUMN_MAP = {
    "project":          "Project",
    "wbs":              "Title",
    "ask_raw":          "Action Description",
    "impact":           "Project Impact",
    "due_date":         "BL Due Date",
    "owner_raw":        "Owner",
    "forecast":         "NEW Due Date",
    "status":           "Status",
    "plan_raw":         "Comments",
    "level":            "Escalation Level",
    "item_type":        "Item Type",
    "path":             "Path",
    "action_type":      "Action Type",
    "functional_owner": "Functional Owner",
    "record_id":        "Record ID (R&O)",
    "project_code":     "Project Header Code",
}

ACTION_LOG_REQUIRED_KEYS = [
    "project", "wbs", "ask_raw", "impact", "due_date", "owner_raw",
    "forecast", "status", "plan_raw", "level",
]

COMPARISON_FIELDS = [
    ("status", "Status"),
    ("dueDate", "Due Date"),
    ("forecastDate", "Forecast Conclusion Date"),
    ("daysOpen", "Days Open"),
    ("owner", "Owner"),
    ("coOwner", "Co-Owner"),
    ("impact", "Impact"),
    ("declaredLevel", "Declared Escalation Level"),
    ("reportLevel", "Reported Level"),
    ("alertLevel", "Alert Level"),
    ("plan", "Action Plan / Comments"),
]


def _is_nan(value):
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def clean_text(value, max_len=None, escape_html=True):
    if _is_nan(value):
        return ""
    text = str(value).replace(CARRIAGE_RETURN, NEWLINE)
    text = re.sub(NEWLINE + "{2,}", NEWLINE, text)
    whitespace_pattern = "[ " + TAB + "]{2,}"
    text = re.sub(whitespace_pattern, " ", text)
    text = text.strip()
    if escape_html:
        text = html.escape(text, quote=False)
    if max_len and len(text) > max_len:
        text = text[:max_len].rsplit(" ", 1)[0].rstrip(",.;:") + "..."
    return text


# Some source spreadsheets export the "Owner" field as a people-picker value
# that appends the company name after each person's name (e.g. "Jane Doe Example Corp").
# Set this to match your organization's name, or leave it empty to disable the strip.
OWNER_COMPANY_SUFFIX = "Example Corp"


def extract_owners(value):
    if _is_nan(value):
        return []

    text = str(value).strip()
    if not text:
        return []

    suffix = re.escape(OWNER_COMPANY_SUFFIX) if OWNER_COMPANY_SUFFIX else None
    names = re.findall(rf"([^;#]+?)\s*{suffix}", text, flags=re.IGNORECASE) if suffix else []
    names = [n.strip(" ;#-()") for n in names if n.strip(" ;#-()")]

    if not names:
        parts = [p.strip() for p in re.split(r";#", text) if p.strip()]
        cleaned = []
        for part in parts:
            if suffix:
                part = re.sub(rf"\s*{suffix}\s*$", "", part, flags=re.IGNORECASE).strip()
            part = re.sub(r"^\d+\s*", "", part).strip()
            if part:
                cleaned.append(part)
        names = cleaned

    unique_names = []
    seen = set()
    for name in names:
        key = name.lower()
        if key not in seen:
            seen.add(key)
            unique_names.append(name)

    return unique_names


def clean_owner(value):
    owners = extract_owners(value)
    if not owners:
        return "Unassigned"
    return owners[0]


def clean_co_owner(value):
    owners = extract_owners(value)
    if len(owners) <= 1:
        return ""
    return " + ".join(owners[1:])


def default_wbs(value):
    if _is_nan(value):
        return "(General)"
    text = str(value).strip()
    return text if text else "(General)"


def parse_date(value):
    if _is_nan(value):
        return None
    try:
        dt = pd.to_datetime(value)
        if pd.isna(dt):
            return None
        return dt.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return None


def parse_days_open(value):
    if _is_nan(value):
        return None
    try:
        return int(round(float(value)))
    except (ValueError, TypeError):
        return None


def calculate_days_open_from_date(date_str):
    if not date_str:
        return None
    try:
        reference_date = datetime.strptime(date_str, "%Y-%m-%d")
    except (ValueError, TypeError):
        return None

    delta_days = (datetime.now() - reference_date).days
    return max(delta_days, 0)


def normalize_status_global(value):
    text = clean_text(value, escape_html=False)
    if not text:
        return "Plan to be defined"
    if "plan in place" in text.lower():
        return "Plan in place"
    return "Plan to be defined"


def normalize_level_global(value):
    text = clean_text(value, escape_html=False).upper()
    if not text:
        return "GENERAL"
    if text in {"L1", "L2", "L3"}:
        return text
    return "GENERAL"


def normalize_status_action_log(value):
    text = clean_text(value, escape_html=True)
    return text if text else "Open"


def normalize_level_action_log(value):
    text = clean_text(value, escape_html=False).upper()
    if not text:
        return "GENERAL"
    if text.startswith("L1"):
        return "L1"
    if text.startswith("L2"):
        return "L2"
    if text.startswith("L3"):
        return "L3"
    return "GENERAL"


def classify_report_level(record):
    declared = (record.get("declaredLevel") or "GENERAL").upper()
    if declared in ("L1", "L2", "L3"):
        return declared
    return "GENERAL"


def has_defined_plan(record):
    plan_text = (record.get("plan") or "").strip()
    return bool(plan_text)


def compute_alert(record):
    level = record.get("reportLevel")
    days_open = record.get("daysOpen")
    plan_defined = has_defined_plan(record)

    if level == "L3" and days_open is not None and days_open > 30:
        return "critical", (
            "Critical: L3 item open for more than 30 days, regardless "
            "of action plan status. Immediate attention required."
        )

    if level == "L2" and days_open is not None and days_open >= 7 and not plan_defined:
        return "urgent", (
            "Urgent: L2 item with 7 or more days open and no text found "
            "in the action plan / comments field yet. If not addressed, "
            "this item will likely need to be escalated to L3."
        )

    if level == "L2" and days_open is not None and days_open > 30:
        return "alert", (
            "Alert: L2 item open for more than 30 days. An action plan "
            "is already defined (text present in Action Plan Detail / "
            "Comments), but the item remains open longer than expected "
            "for this level - consider reviewing the plan or escalating."
        )

    if level in ("L2", "L3") and plan_defined and days_open is not None and 8 <= days_open <= 30:
        return "attention", (
            "Attention: item has a defined action plan (text present in "
            "Action Plan Detail / Comments) and has been open between 8 "
            "and 30 days. Keep monitoring progress."
        )

    return "na", ""


def validate_required_columns(df, column_map, required_keys):
    missing = [
        column_map[k] for k in required_keys
        if column_map.get(k) and column_map[k] not in df.columns
    ]
    if missing:
        raise ValueError(
            "The selected sheet does not contain the required columns: "
            + ", ".join(missing)
        )


def load_data(xlsx_path, sheet_name, column_map, required_keys, source_base):
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, engine="openpyxl")
    df = df.dropna(how="all")

    validate_required_columns(df, column_map, required_keys)

    source_file = Path(xlsx_path).name
    all_records = []

    item_type_col = column_map.get("item_type")
    path_col = column_map.get("path")
    action_type_col = column_map.get("action_type")
    functional_owner_col = column_map.get("functional_owner")
    record_id_col = column_map.get("record_id")
    project_code_col = column_map.get("project_code")

    for idx, (_, row) in enumerate(df.iterrows()):
        raw_owner_value = row.get(column_map["owner_raw"])
        owners_list = extract_owners(raw_owner_value)

        due_date = parse_date(row.get(column_map["due_date"]))

        record = {
            "id": idx,
            "sourceSheet":  sheet_name,
            "sourceFile":   source_file,
            "sourceBase":   source_base,
            "project":      clean_text(row.get(column_map["project"]), escape_html=True) or "-",
            "wbs":          default_wbs(row.get(column_map["wbs"])),
            "impact":       clean_text(row.get(column_map["impact"]), escape_html=True) or "N/A",
            "owner":        clean_owner(raw_owner_value),
            "coOwner":      clean_co_owner(raw_owner_value),
            "ownersList":   owners_list if owners_list else ["Unassigned"],
            "dueDate":      due_date,
            "forecastDate": parse_date(row.get(column_map["forecast"])),
            "ask":          clean_text(row.get(column_map["ask_raw"]), max_len=MAX_ASK_LENGTH),
            "plan":         clean_text(row.get(column_map["plan_raw"])),
            "itemType":     clean_text(row.get(item_type_col), escape_html=True) if item_type_col in df.columns else "",
            "path":         clean_text(row.get(path_col), escape_html=True) if path_col in df.columns else "",
            "actionType":       "",
            "functionalOwner":  "",
            "recordId":         "",
            "projectCode":      "",
            "daysOpenSource":   "",
        }

        if source_base == SOURCE_GLOBAL_PM:
            record["status"] = normalize_status_global(row.get(column_map["status"]))
            record["declaredLevel"] = normalize_level_global(row.get(column_map["level"]))
            created_at = parse_date(row.get(column_map["created_at"]))
            record["createdAt"] = created_at

            raw_days_open = parse_days_open(row.get(column_map["days_open"]))
            if raw_days_open is not None:
                record["daysOpen"] = raw_days_open
                record["daysOpenSource"] = "L2 Days Open (source)"
            else:
                calculated_days_open = calculate_days_open_from_date(created_at)
                record["daysOpen"] = calculated_days_open
                record["daysOpenSource"] = (
                    "Calculated from L2 Created"
                    if calculated_days_open is not None
                    else "N/A - L2 Created not available"
                )
        else:
            record["status"] = normalize_status_action_log(row.get(column_map["status"]))
            record["declaredLevel"] = normalize_level_action_log(row.get(column_map["level"]))
            record["createdAt"] = None

            calculated_days_open = calculate_days_open_from_date(due_date)
            record["daysOpen"] = calculated_days_open
            record["daysOpenSource"] = (
                "Calculated from BL Due Date"
                if calculated_days_open is not None
                else "N/A - BL Due Date not available"
            )

            if action_type_col in df.columns:
                record["actionType"] = clean_text(row.get(action_type_col), escape_html=True)
            if functional_owner_col in df.columns:
                record["functionalOwner"] = clean_text(row.get(functional_owner_col), escape_html=True)
            if record_id_col in df.columns:
                record["recordId"] = clean_text(row.get(record_id_col), escape_html=True)
            if project_code_col in df.columns:
                record["projectCode"] = clean_text(row.get(project_code_col), escape_html=True)

            # For L1 items in Action Log, show Functional Owner (e.g. ENG, SOU, PMO)
            # instead of the individual Owner. If Functional Owner is empty,
            # keep the original Owner as fallback.
            declared_level_upper = (record.get("declaredLevel") or "").upper()
            functional_owner_value = (record.get("functionalOwner") or "").strip()
            if declared_level_upper == "L1" and functional_owner_value:
                record["owner"] = functional_owner_value
                record["coOwner"] = ""
                record["ownersList"] = [functional_owner_value]

        record["reportLevel"] = classify_report_level(record)
        record["level"] = record["reportLevel"]

        alert_level, alert_text = compute_alert(record)
        record["alertLevel"] = alert_level
        record["alertText"] = alert_text
        record["planDefined"] = has_defined_plan(record)

        all_records.append(record)

    return all_records


def split_dashboard_views(records):
    return {
        "general": records,
        "l1": [r for r in records if r.get("reportLevel") == "L1"],
        "l2": [r for r in records if r.get("reportLevel") == "L2"],
        "l3": [r for r in records if r.get("reportLevel") == "L3"],
    }


def build_dashboard_payload(records):
    views = split_dashboard_views(records)
    return {
        "views": views,
        "totalCount": len(records),
    }


def renumber_records(records):
    for idx, r in enumerate(records):
        r["id"] = idx
    return records


def make_compare_key(record):
    project = (record.get("project") or "").strip().lower()
    wbs = (record.get("wbs") or "").strip().lower()
    return project + "||" + wbs


def prepare_week_data(xlsx_path, sheet_name, column_map, required_keys, source_base):
    records = load_data(xlsx_path, sheet_name, column_map, required_keys, source_base)
    payload = build_dashboard_payload(records)
    return payload, records


def build_comparison_data(current_records, previous_records):
    current_map = {}
    for r in current_records:
        current_map[make_compare_key(r)] = r

    previous_map = {}
    for r in previous_records:
        previous_map[make_compare_key(r)] = r

    current_keys = set(current_map.keys())
    previous_keys = set(previous_map.keys())

    new_keys = current_keys - previous_keys
    closed_keys = previous_keys - current_keys
    common_keys = current_keys & previous_keys

    new_items = [current_map[k] for k in new_keys]
    closed_items = [previous_map[k] for k in closed_keys]

    changed_items = []
    unchanged_count = 0

    for k in common_keys:
        cur = current_map[k]
        prev = previous_map[k]
        diffs = {}
        for field_key, field_label in COMPARISON_FIELDS:
            cur_val = cur.get(field_key)
            prev_val = prev.get(field_key)
            if cur_val != prev_val:
                diffs[field_key] = {
                    "label": field_label,
                    "previous": prev_val,
                    "current": cur_val,
                }
        if diffs:
            changed_items.append({
                "project": cur.get("project"),
                "wbs": cur.get("wbs"),
                "id": cur.get("id"),
                "diffs": diffs,
            })
        else:
            unchanged_count += 1

    changed_items.sort(
        key=lambda x: (x["diffs"].get("daysOpen", {}).get("current") or 0),
        reverse=True
    )
    new_items.sort(key=lambda r: r.get("daysOpen") or 0, reverse=True)
    closed_items.sort(key=lambda r: (r.get("project") or "").lower())

    return {
        "newItems": new_items,
        "closedItems": closed_items,
        "changedItems": changed_items,
        "unchangedCount": unchanged_count,
        "totalCurrent": len(current_records),
        "totalPrevious": len(previous_records),
    }


HTML_TEMPLATE = r"""<html lang="en"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>__TITLE__</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&amp;display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root {
  --color-dk1: #1E2A47; --color-lt1: #FFFFFF; --color-dk2: #212121; --color-lt2: #5EEAD4;
  --color-accent1: #16213A; --color-accent6: #4C6A92; --color-tint-2: #93A9C9; --color-tint-3: #C7D3E6;
  --color-gray-1: #FAFAFA; --color-accent2: #E3E3E3; --color-accent3: #CCCCCC; --color-accent4: #696969;
  --color-accent5: #484848; --color-hlink: #2176FF; --color-sky: #5FB7FF; --color-fol-hlink:#3E9E00;
  --color-forest: #7BDD3C; --color-dark-earth:#4A2419; --color-earth: #774936;
  --color-critical: #E63B2E; --color-urgent: #F5A623; --color-attention:#F2C94C; --color-alert: #5FB7FF;
  --font-heading: "Inter", "Segoe UI", sans-serif;
  --font-body: "Inter", "Segoe UI", sans-serif;
}
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: var(--font-body); background: var(--color-dk1); color: var(--color-lt1); min-height: 100vh; font-size: 14px; line-height: 1.5; margin-bottom: 20px; }
.dash-wrapper { max-width: 1440px; margin: 0 auto; padding: 0 0 40px 0; }
.dash-header { background: var(--color-accent1); border-bottom: 3px solid var(--color-lt2); padding: 20px 32px; display: flex; align-items: center; justify-content: space-between; gap: 20px; flex-wrap: wrap; }
.header-left { display: flex; align-items: center; gap: 20px; }
.logo-svg { height: 44px; width: auto; flex-shrink: 0; }
.header-divider { width: 2px; height: 44px; background: var(--color-lt2); opacity: 0.6; }
.header-titles h1 { font-family: var(--font-heading); font-size: 22px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.04em; color: var(--color-lt1); line-height: 1.1; }
.header-titles p { font-size: 12px; color: var(--color-tint-2); font-weight: 400; margin-top: 3px; letter-spacing: 0.02em; }
.header-right { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
.header-badge { background: var(--color-lt2); color: var(--color-dk2); font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 2px; text-transform: uppercase; letter-spacing: 0.06em; }
.header-timestamp { font-size: 11px; color: var(--color-tint-2); font-weight: 400; }
.filter-bar { background: var(--color-accent1); border-bottom: 1px solid rgba(147,169,201,0.2); padding: 14px 32px; }
.filter-bar-row { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.filter-label { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; color: var(--color-tint-2); white-space: nowrap; }
.filter-select, .filter-search { background: var(--color-dk1); border: 1px solid rgba(147,169,201,0.3); color: var(--color-lt1); font-family: var(--font-body); font-size: 13px; padding: 7px 12px; border-radius: 2px; outline: none; cursor: pointer; transition: border-color 0.2s; }
.filter-select:focus, .filter-search:focus { border-color: var(--color-lt2); }
.filter-select option { background: var(--color-accent1); color: var(--color-lt1); }
.filter-search { min-width: 220px; flex: 1; }
.filter-search::placeholder { color: var(--color-accent6); opacity: 0.7; }
.filter-divider { width: 1px; height: 22px; background: rgba(147,169,201,0.25); margin: 0 4px; }
.btn-reset { background: transparent; border: 1px solid var(--color-lt2); color: var(--color-lt2); font-family: var(--font-body); font-size: 12px; font-weight: 600; padding: 7px 14px; border-radius: 2px; cursor: pointer; text-transform: uppercase; letter-spacing: 0.06em; transition: background 0.2s, color 0.2s; white-space: nowrap; }
.btn-reset:hover { background: var(--color-lt2); color: var(--color-dk2); }
.source-info-bar { background: var(--color-accent1); border-bottom: 1px solid rgba(147,169,201,0.15); padding: 10px 32px; display: flex; gap: 20px; flex-wrap: wrap; font-size: 11px; color: var(--color-tint-2); align-items: center; }
.source-info-item { display: flex; align-items: center; gap: 6px; white-space: nowrap; }
.source-info-item strong { color: var(--color-lt1); }
.tab-nav { display: flex; gap: 4px; background: var(--color-accent1); border-bottom: 1px solid rgba(147,169,201,0.2); padding: 0 32px; flex-wrap: wrap; }
.tab-btn { background: transparent; border: none; border-bottom: 3px solid transparent; color: var(--color-tint-2); font-family: var(--font-heading); font-size: 12px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.08em; padding: 14px 18px; cursor: pointer; transition: color 0.2s, border-color 0.2s; }
.tab-btn:hover { color: var(--color-lt1); }
.tab-btn.active { color: var(--color-lt2); border-bottom-color: var(--color-lt2); }
.tab-content { display: none; flex-direction: column; gap: 24px; }
.tab-content.active { display: flex; }
.dash-main { padding: 24px 32px; display: flex; flex-direction: column; gap: 24px; }
.kpi-row { display: grid; grid-template-columns: repeat(6, 1fr); gap: 14px; }
.kpi-row.kpi-row-compact { grid-template-columns: repeat(2, 1fr); }
.kpi-card { background: var(--color-accent1); border: 1px solid rgba(147,169,201,0.2); border-top: 3px solid var(--color-accent6); padding: 16px 18px; border-radius: 2px; position: relative; overflow: hidden; }
.kpi-card.kpi-critical { border-top-color: var(--color-critical); }
.kpi-card.kpi-urgent { border-top-color: var(--color-urgent); }
.kpi-card.kpi-attention { border-top-color: var(--color-attention); }
.kpi-card.kpi-alert { border-top-color: var(--color-alert); }
.kpi-card.kpi-plan { border-top-color: var(--color-lt2); }
.kpi-card.kpi-new { border-top-color: #5FB7FF; }
.kpi-card.kpi-resolved { border-top-color: #7BDD3C; }
.kpi-label { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.1em; color: var(--color-tint-2); margin-bottom: 8px; }
.kpi-value { font-family: var(--font-heading); font-size: 36px; font-weight: 800; line-height: 1; color: var(--color-lt1); }
.kpi-card.kpi-critical .kpi-value { color: var(--color-critical); }
.kpi-card.kpi-urgent .kpi-value { color: var(--color-urgent); }
.kpi-card.kpi-attention .kpi-value { color: var(--color-attention); }
.kpi-card.kpi-alert .kpi-value { color: var(--color-alert); }
.kpi-card.kpi-plan .kpi-value { color: var(--color-lt2); }
.kpi-card.kpi-new .kpi-value { color: #5FB7FF; }
.kpi-card.kpi-resolved .kpi-value { color: #7BDD3C; }
.kpi-sub { font-size: 10.5px; color: var(--color-accent6); margin-top: 4px; }
.section-header { font-family: var(--font-heading); font-size: 13px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.1em; color: var(--color-lt2); padding-bottom: 8px; border-bottom: 1px solid rgba(94,234,212,0.2); margin-bottom: 16px; }
.charts-row { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
.chart-panel { background: var(--color-accent1); border: 1px solid rgba(147,169,201,0.2); border-radius: 2px; padding: 20px; }
.chart-panel-title { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: var(--color-tint-2); margin-bottom: 14px; }
.chart-canvas-wrapper { position: relative; height: 260px; }
.chart-canvas-wrapper.doughnut-box { height: 220px; }
.alert-banner { background: rgba(230,59,46,0.12); border: 1px solid rgba(230,59,46,0.4); border-left: 4px solid var(--color-critical); border-radius: 2px; padding: 14px 20px; }
.alert-banner-header { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }
.alert-banner-title { font-family: var(--font-heading); font-size: 13px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.08em; color: var(--color-critical); }
.alert-items { display: flex; flex-wrap: wrap; gap: 8px; }
.alert-item { background: rgba(230,59,46,0.15); border: 1px solid rgba(230,59,46,0.3); border-radius: 2px; padding: 6px 12px; font-size: 11px; color: var(--color-lt1); cursor: pointer; transition: background 0.2s; }
.alert-item:hover { background: rgba(230,59,46,0.3); }
.alert-item strong { color: var(--color-critical); font-weight: 700; }
.table-panel { background: var(--color-accent1); border: 1px solid rgba(147,169,201,0.2); border-radius: 2px; overflow: hidden; }
.table-toolbar { padding: 14px 20px; display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid rgba(147,169,201,0.15); flex-wrap: wrap; gap: 10px; }
.table-count { font-size: 12px; color: var(--color-tint-2); }
.table-count span { color: var(--color-lt2); font-weight: 700; }
.table-scroll { overflow-x: auto; max-height: 600px; overflow-y: auto; }
table { width: 100%; border-collapse: collapse; font-size: 12px; }
thead th { background: var(--color-dk1); color: var(--color-tint-2); font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; padding: 10px 12px; text-align: left; white-space: nowrap; position: sticky; top: 0; z-index: 10; border-bottom: 2px solid rgba(147,169,201,0.2); cursor: pointer; user-select: none; }
thead th:hover { color: var(--color-lt2); }
thead th .sort-icon { margin-left: 4px; opacity: 0.5; font-size: 9px; }
thead th.sorted .sort-icon { opacity: 1; color: var(--color-lt2); }
tbody tr { border-bottom: 1px solid rgba(147,169,201,0.08); transition: background 0.15s; cursor: pointer; }
tbody tr:nth-child(even) { background: rgba(22,33,58,0.3); }
tbody tr:hover { background: rgba(76,106,146,0.2); }
tbody tr.expanded { background: rgba(76,106,146,0.15); }
tbody td { padding: 9px 12px; color: var(--color-lt1); vertical-align: middle; }
tbody td.td-project { font-weight: 600; color: var(--color-tint-2); white-space: nowrap; }
tbody td.td-wbs { max-width: 160px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--color-lt1); }
.badge { display: inline-flex; align-items: center; gap: 4px; padding: 3px 8px; border-radius: 2px; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; white-space: nowrap; }
.badge-critical { background: rgba(230,59,46,0.2); color: var(--color-critical); border: 1px solid rgba(230,59,46,0.4); }
.badge-urgent { background: rgba(245,166,35,0.2); color: var(--color-urgent); border: 1px solid rgba(245,166,35,0.4); }
.badge-attention { background: rgba(242,201,76,0.2); color: var(--color-attention); border: 1px solid rgba(242,201,76,0.45); }
.badge-alert { background: rgba(95,183,255,0.18); color: var(--color-alert); border: 1px solid rgba(95,183,255,0.4); }
.badge-na { background: rgba(105,105,105,0.2); color: var(--color-accent4); border: 1px solid rgba(105,105,105,0.3); }
.badge-plan { background: rgba(94,234,212,0.12); color: var(--color-lt2); border: 1px solid rgba(94,234,212,0.3); }
.badge-plan-overdue { background: rgba(230,59,46,0.2); color: var(--color-critical); border: 1px solid rgba(230,59,46,0.45); }
.badge-define { background: rgba(147,169,201,0.15); color: var(--color-tint-2); border: 1px solid rgba(147,169,201,0.3); }
.badge-new { background: rgba(95,183,255,0.15); color: #5FB7FF; border: 1px solid rgba(95,183,255,0.4); }
.badge-resolved { background: rgba(123,221,60,0.15); color: #7BDD3C; border: 1px solid rgba(123,221,60,0.4); }
.badge-source-gpm { background: rgba(147,169,201,0.18); color: #93A9C9; border: 1px solid rgba(147,169,201,0.4); }
.badge-source-al { background: rgba(123,221,60,0.15); color: #7BDD3C; border: 1px solid rgba(123,221,60,0.4); }
.btn-expand { background: transparent; border: 1px solid rgba(147,169,201,0.3); color: var(--color-tint-2); font-size: 11px; padding: 4px 8px; border-radius: 2px; cursor: pointer; transition: all 0.2s; white-space: nowrap; }
.btn-expand:hover { border-color: var(--color-lt2); color: var(--color-lt2); }
.expand-row td { background: rgba(22,33,58,0.4); padding: 0; }
.expand-content { padding: 16px 20px; display: grid; grid-template-columns: 1fr 1fr; gap: 16px; border-top: 1px solid rgba(147,169,201,0.15); }
.expand-section h4 { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: var(--color-lt2); margin-bottom: 6px; }
.expand-section p { font-size: 12px; color: var(--color-accent3); line-height: 1.6; white-space: pre-line; }
.days-critical { color: var(--color-critical); font-weight: 700; }
.days-urgent { color: var(--color-urgent); font-weight: 700; }
.days-attention { color: var(--color-attention); font-weight: 700; }
.days-alert { color: var(--color-alert); font-weight: 700; }
.days-na { color: var(--color-accent4); }
.dash-footer { text-align: center; padding: 20px 32px; font-size: 11px; color: var(--color-accent6); border-top: 1px solid rgba(147,169,201,0.15); }
.dash-footer strong { color: var(--color-lt2); }
.diff-prev { color: var(--color-accent3); text-decoration: line-through; opacity: 0.75; }
.diff-cur { color: var(--color-lt2); font-weight: 700; }
@media (max-width: 1200px) { .kpi-row { grid-template-columns: repeat(3, 1fr); } }
@media (max-width: 768px) { .dash-header { padding: 16px 16px; } .dash-main { padding: 16px 16px; } .filter-bar { padding: 12px 16px; } .tab-nav { padding: 0 16px; } .kpi-row { grid-template-columns: repeat(2, 1fr); } .charts-row { grid-template-columns: 1fr; } .expand-content { grid-template-columns: 1fr; } }
@media (max-width: 480px) { .kpi-row { grid-template-columns: 1fr 1fr; } .header-titles h1 { font-size: 16px; } }
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: var(--color-accent1); }
::-webkit-scrollbar-thumb { background: var(--color-accent6); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: var(--color-tint-2); }
@keyframes fadeIn { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: translateY(0); } }
.fade-in { animation: fadeIn 0.25s ease forwards; }
.no-data { text-align: center; padding: 40px; color: var(--color-accent6); font-size: 13px; }
#print-footer { display: flex; justify-content: space-between; align-items: flex-end; font-family: Inter, "Segoe UI", sans-serif; font-size: 0.8rem; min-height: 20px; width: 100%; color: grey; margin-top: 4px; padding: 0 12px; box-sizing: border-box; background-color: white; }
#print-footer .source-files { text-align: left; flex: 1 1 auto; margin-right: 16px; }
#print-footer .generated-date { text-align: right; flex: 0 0 auto; white-space: nowrap; }
.floating-hscroll {
  position: fixed;
  bottom: 0;
  height: 14px;
  overflow-x: scroll;
  overflow-y: hidden;
  background: var(--color-accent1);
  border-top: 1px solid rgba(147,169,201,0.35);
  box-shadow: 0 -2px 6px rgba(0,0,0,0.25);
  z-index: 999;
  display: none;
}
.floating-hscroll.visible { display: block; }
.floating-hscroll-inner { height: 1px; }
.floating-hscroll::-webkit-scrollbar { height: 10px; }
.floating-hscroll::-webkit-scrollbar-thumb { background: var(--color-accent6); border-radius: 5px; }
</style>
</head>
<body>
<div class="dash-wrapper">
<header class="dash-header">
  <div class="header-left">
    <div class="header-titles">
      <h1>Action Tracker</h1>
      <p>Critical Actions Dashboard</p>
    </div>
  </div>
  <div class="header-right">
    <span class="header-badge">LIVE TRACKING</span>
    <span class="header-timestamp">Last Updated: __LAST_UPDATED__</span>
  </div>
</header>

<div class="filter-bar">
  <div class="filter-bar-row">
    <span class="filter-label">Filter:</span>
    <select class="filter-select" id="filterSource"><option value="">All Sources</option></select>
    <select class="filter-select" id="filterProject"><option value="">All Projects</option></select>
    <select class="filter-select" id="filterOwner"><option value="">All Owners</option></select>
    <select class="filter-select" id="filterImpact"><option value="">All Impacts</option></select>
    <select class="filter-select" id="filterUrgency">
      <option value="">All Alerts</option>
      <option value="critical">🔴 Critical</option>
      <option value="urgent">🟠 Urgent</option>
      <option value="attention">🟡 Attention</option>
      <option value="alert">🔵 Alert</option>
      <option value="na">⚪ N/A</option>
    </select>
    <div class="filter-divider"></div>
    <input class="filter-search" id="filterSearch" type="text" placeholder="Search project, WBS, owner, action type...">
    <button class="btn-reset" id="btnReset">Reset Filters</button>
  </div>
</div>

<div class="source-info-bar" id="sourceInfoBar"></div>

<div class="tab-nav" id="tabNav">
  <button class="tab-btn active" data-view="general">📊 Overview</button>
  <button class="tab-btn" data-view="l1">L1</button>
  <button class="tab-btn" data-view="l2">L2</button>
  <button class="tab-btn" data-view="l3">L3</button>
  <button class="tab-btn" data-target="tabContentComparisonGPM" id="tabBtnComparisonGPM" style="display:none;">Comparison — Global PM</button>
  <button class="tab-btn" data-target="tabContentComparisonAL" id="tabBtnComparisonAL" style="display:none;">🔄 Comparison — Action Logs</button>
</div>

<main class="dash-main">

<div class="tab-content active" id="tabContentDashboard">

<div class="kpi-row" id="kpiRow">
  <div class="kpi-card" id="kpiCardTotal">
    <div class="kpi-label">Total Items</div>
    <div class="kpi-value" id="kpiTotal">__TOTAL_COUNT__</div>
    <div class="kpi-sub">Active critical actions</div>
  </div>
  <div class="kpi-card kpi-critical" id="kpiCardCritical">
    <div class="kpi-label">🔴 Critical</div>
    <div class="kpi-value" id="kpiCritical">-</div>
    <div class="kpi-sub">L3, open &gt; 30 days</div>
  </div>
  <div class="kpi-card kpi-urgent" id="kpiCardUrgent">
    <div class="kpi-label">🟠 Urgent</div>
    <div class="kpi-value" id="kpiUrgent">-</div>
    <div class="kpi-sub">L2, &ge;7 days, no plan defined</div>
  </div>
  <div class="kpi-card kpi-attention" id="kpiCardAttention">
    <div class="kpi-label">🟡 Attention</div>
    <div class="kpi-value" id="kpiAttention">-</div>
    <div class="kpi-sub">L2/L3 with plan, 8–30 days</div>
  </div>
  <div class="kpi-card kpi-alert" id="kpiCardAlert">
    <div class="kpi-label">🔵 Alert</div>
    <div class="kpi-value" id="kpiAlert">-</div>
    <div class="kpi-sub">L2, open &gt; 30 days (plan defined)</div>
  </div>
  <div class="kpi-card kpi-plan" id="kpiCardPlan">
    <div class="kpi-label">Plan / Comments Defined</div>
    <div class="kpi-value" id="kpiPlan">-</div>
    <div class="kpi-sub">Items with text in Action Plan Detail / Comments</div>
  </div>
</div>

<div>
  <div class="section-header">DISTRIBUTION OVERVIEW</div>
  <div class="charts-row">
    <div class="chart-panel">
      <div class="chart-panel-title">Items by Project</div>
      <div class="chart-canvas-wrapper"><canvas id="chartProject"></canvas></div>
    </div>
    <div class="chart-panel">
      <div class="chart-panel-title">Impact Distribution</div>
      <div class="chart-canvas-wrapper doughnut-box"><canvas id="chartImpact"></canvas></div>
    </div>
  </div>
</div>

<div>
  <div class="charts-row">
    <div class="chart-panel">
      <div class="chart-panel-title">Items by Owner</div>
      <div class="chart-canvas-wrapper"><canvas id="chartOwner"></canvas></div>
    </div>
    <div class="chart-panel">
      <div class="chart-panel-title">Days Open by Item (Top 20, sorted desc)</div>
      <div class="chart-canvas-wrapper"><canvas id="chartDays"></canvas></div>
    </div>
  </div>
</div>

<div class="alert-banner" id="alertBanner">
  <div class="alert-banner-header">
    <span style="font-size:18px;">!</span>
    <span class="alert-banner-title">MOST CRITICAL ITEMS - IMMEDIATE ATTENTION REQUIRED</span>
  </div>
  <div class="alert-items" id="alertItems"></div>
</div>

<div class="table-panel">
  <div class="table-toolbar">
    <div class="section-header" style="margin-bottom:0; border-bottom:none;">CRITICAL ACTIONS REGISTER</div>
    <div class="table-count">Showing <span id="tableCount">0</span> of <span id="tableTotal">__TOTAL_COUNT__</span> items</div>
  </div>
  <div class="table-scroll">
    <table id="dataTable">
      <thead>
        <tr>
          <th data-col="id">Num <span class="sort-icon">up-down</span></th>
          <th data-col="sourceBase">Source <span class="sort-icon">up-down</span></th>
          <th data-col="project">Project <span class="sort-icon">up-down</span></th>
          <th data-col="wbs">WBS / Title <span class="sort-icon">up-down</span></th>
          <th data-col="impact">Impact <span class="sort-icon">up-down</span></th>
          <th data-col="owner">Owner <span class="sort-icon">up-down</span></th>
          <th data-col="coOwner">Co-Owner <span class="sort-icon">up-down</span></th>
          <th data-col="status">Status <span class="sort-icon">up-down</span></th>
          <th data-col="reportLevel">Escalation Level <span class="sort-icon">up-down</span></th>
          <th data-col="daysOpen">Days Open <span class="sort-icon">up-down</span></th>
          <th data-col="dueDate">Due Date <span class="sort-icon">up-down</span></th>
          <th data-col="forecastDate">Forecast <span class="sort-icon">up-down</span></th>
          <th>Alert</th>
          <th>Details</th>
        </tr>
      </thead>
      <tbody id="tableBody"></tbody>
    </table>
    <div class="no-data" id="noData" style="display:none;">No items match the current filters.</div>
  </div>
</div>

</div>

<div class="tab-content" id="tabContentComparisonGPM">
  <div>
    <div class="section-header">WEEKLY COMPARISON — GLOBAL PROJECT MANAGEMENT</div>
    <div class="section-sub-header" style="font-size:11px; color:var(--color-tint-2); margin-bottom:16px;">
      Both weeks are processed through the same rules (Escalation Level + Alert). New / Closed items below already reflect the current classification.
    </div>
    <div class="kpi-row" style="grid-template-columns: repeat(4, 1fr);">
      <div class="kpi-card kpi-new">
        <div class="kpi-label">New Items</div>
        <div class="kpi-value" id="cmpNewGPM">0</div>
        <div class="kpi-sub">Added this week</div>
      </div>
      <div class="kpi-card kpi-resolved">
        <div class="kpi-label">Resolved / Closed</div>
        <div class="kpi-value" id="cmpClosedGPM">0</div>
        <div class="kpi-sub">Removed since last week</div>
      </div>
      <div class="kpi-card kpi-attention">
        <div class="kpi-label">Changed</div>
        <div class="kpi-value" id="cmpChangedGPM">0</div>
        <div class="kpi-sub">Updated status / dates / owner / level / alert</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Unchanged</div>
        <div class="kpi-value" id="cmpUnchangedGPM">0</div>
        <div class="kpi-sub">No changes detected</div>
      </div>
    </div>
  </div>

  <div class="table-panel">
    <div class="table-toolbar"><div class="section-header" style="margin-bottom:0; border-bottom:none;">🆕 NEW ITEMS THIS WEEK</div></div>
    <div class="table-scroll">
      <table>
        <thead><tr><th>Project</th><th>WBS / Title</th><th>Impact</th><th>Owner</th><th>Days Open</th><th>Status</th><th>Level</th><th>Alert</th></tr></thead>
        <tbody id="cmpNewBodyGPM"></tbody>
      </table>
      <div class="no-data" id="cmpNewNoDataGPM" style="display:none;">No new items this week.</div>
    </div>
  </div>

  <div class="table-panel">
    <div class="table-toolbar"><div class="section-header" style="margin-bottom:0; border-bottom:none;">✅ RESOLVED / CLOSED ITEMS</div></div>
    <div class="table-scroll">
      <table>
        <thead><tr><th>Project</th><th>WBS / Title</th><th>Impact</th><th>Owner</th><th>Days Open (Last Week)</th><th>Status (Last Week)</th><th>Level (Last Week)</th><th>Alert (Last Week)</th></tr></thead>
        <tbody id="cmpClosedBodyGPM"></tbody>
      </table>
      <div class="no-data" id="cmpClosedNoDataGPM" style="display:none;">No items were closed this week.</div>
    </div>
  </div>

  <div class="table-panel">
    <div class="table-toolbar"><div class="section-header" style="margin-bottom:0; border-bottom:none;">🔄 CHANGED ITEMS</div></div>
    <div class="table-scroll">
      <table>
        <thead><tr><th>Project</th><th>WBS / Title</th><th>Field</th><th>Last Week</th><th>This Week</th></tr></thead>
        <tbody id="cmpChangedBodyGPM"></tbody>
      </table>
      <div class="no-data" id="cmpChangedNoDataGPM" style="display:none;">No changes detected between weeks.</div>
    </div>
  </div>
</div>

<div class="tab-content" id="tabContentComparisonAL">
  <div>
    <div class="section-header">WEEKLY COMPARISON — PROJECT ACTION LOGS</div>
    <div class="section-sub-header" style="font-size:11px; color:var(--color-tint-2); margin-bottom:16px;">
      Both weeks are processed through the same rules (Escalation Level + Alert). New / Closed items below already reflect the current classification.
    </div>
    <div class="kpi-row" style="grid-template-columns: repeat(4, 1fr);">
      <div class="kpi-card kpi-new">
        <div class="kpi-label">New Items</div>
        <div class="kpi-value" id="cmpNewAL">0</div>
        <div class="kpi-sub">Added this week</div>
      </div>
      <div class="kpi-card kpi-resolved">
        <div class="kpi-label">Resolved / Closed</div>
        <div class="kpi-value" id="cmpClosedAL">0</div>
        <div class="kpi-sub">Removed since last week</div>
      </div>
      <div class="kpi-card kpi-attention">
        <div class="kpi-label">Changed</div>
        <div class="kpi-value" id="cmpChangedAL">0</div>
        <div class="kpi-sub">Updated status / dates / owner / level / alert</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Unchanged</div>
        <div class="kpi-value" id="cmpUnchangedAL">0</div>
        <div class="kpi-sub">No changes detected</div>
      </div>
    </div>
  </div>

  <div class="table-panel">
    <div class="table-toolbar"><div class="section-header" style="margin-bottom:0; border-bottom:none;">🆕 NEW ITEMS THIS WEEK</div></div>
    <div class="table-scroll">
      <table>
        <thead><tr><th>Project</th><th>WBS / Title</th><th>Impact</th><th>Owner</th><th>Days Open</th><th>Status</th><th>Level</th><th>Alert</th></tr></thead>
        <tbody id="cmpNewBodyAL"></tbody>
      </table>
      <div class="no-data" id="cmpNewNoDataAL" style="display:none;">No new items this week.</div>
    </div>
  </div>

  <div class="table-panel">
    <div class="table-toolbar"><div class="section-header" style="margin-bottom:0; border-bottom:none;">✅ RESOLVED / CLOSED ITEMS</div></div>
    <div class="table-scroll">
      <table>
        <thead><tr><th>Project</th><th>WBS / Title</th><th>Impact</th><th>Owner</th><th>Days Open (Last Week)</th><th>Status (Last Week)</th><th>Level (Last Week)</th><th>Alert (Last Week)</th></tr></thead>
        <tbody id="cmpClosedBodyAL"></tbody>
      </table>
      <div class="no-data" id="cmpClosedNoDataAL" style="display:none;">No items were closed this week.</div>
    </div>
  </div>

  <div class="table-panel">
    <div class="table-toolbar"><div class="section-header" style="margin-bottom:0; border-bottom:none;">🔄 CHANGED ITEMS</div></div>
    <div class="table-scroll">
      <table>
        <thead><tr><th>Project</th><th>WBS / Title</th><th>Field</th><th>Last Week</th><th>This Week</th></tr></thead>
        <tbody id="cmpChangedBodyAL"></tbody>
      </table>
      <div class="no-data" id="cmpChangedNoDataAL" style="display:none;">No changes detected between weeks.</div>
    </div>
  </div>
</div>

</main>

<footer class="dash-footer">
  <strong>Action Tracker Dashboard</strong> &nbsp;|&nbsp; <strong>Last Updated: __LAST_UPDATED__</strong>
</footer>
</div>

<div class="floating-hscroll" id="floatingHScroll">
  <div class="floating-hscroll-inner" id="floatingHScrollInner"></div>
</div>

<script>
const dashboardPayload = __DATA_JSON__;
const comparisonDataGPM = __COMPARISON_GPM_JSON__;
const comparisonDataAL = __COMPARISON_AL_JSON__;

let currentView = 'general';
let allData = [...((dashboardPayload.views || {})[currentView] || [])];
let filteredData = [...allData];
let sortCol = null;
let sortDir = 1;
let expandedRows = new Set();
let charts = {};

const BORDER_COLORS = {
  critical: '#E63B2E',
  urgent: '#F5A623',
  attention: '#F2C94C',
  alert: '#5FB7FF'
};

function hasDefinedPlan(item) {
  const planText = (item.plan || '').trim();
  return planText.length > 0;
}

function isForecastOverdue(item) {
  if (!item.forecastDate) return false;
  const forecast = new Date(item.forecastDate + 'T00:00:00');
  if (isNaN(forecast.getTime())) return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return forecast.getTime() < today.getTime();
}

function getAlert(item) {
  return item.alertLevel || 'na';
}

function formatDate(s) {
  if (!s) return '-';
  const parts = s.split('-');
  if (parts.length !== 3) return s;
  return parts[2] + '/' + parts[1] + '/' + parts[0].slice(2);
}

function formatDays(d, item=null) {
  if (d === null || d === undefined) return '<span class="days-na">N/A</span>';
  if (d > 9999) {
    return '<span class="days-na" title="Possible invalid date in source data">⚠ ' + d.toLocaleString() + '</span>';
  }
  const u = item ? getAlert(item) : 'na';
  if (u === 'critical') return '<span class="days-critical">' + d + '</span>';
  if (u === 'urgent') return '<span class="days-urgent">' + d + '</span>';
  if (u === 'attention') return '<span class="days-attention">' + d + '</span>';
  if (u === 'alert') return '<span class="days-alert">' + d + '</span>';
  return '<span class="days-na">' + d + '</span>';
}

function alertBadge(item) {
  const u = getAlert(item);
  if (u === 'critical') return '<span class="badge badge-critical">🔴 CRITICAL</span>';
  if (u === 'urgent') return '<span class="badge badge-urgent">🟠 URGENT</span>';
  if (u === 'attention') return '<span class="badge badge-attention">🟡 ATTENTION</span>';
  if (u === 'alert') return '<span class="badge badge-alert">🔵 ALERT</span>';
  return '<span class="badge badge-na">N/A</span>';
}

function statusBadge(item) {
  const defined = hasDefinedPlan(item);
  const label = item.sourceBase === 'ACTION_LOG'
    ? (item.status || 'Open')
    : (defined ? 'Plan in Place' : 'To be Defined');

  if (!defined) {
    return '<span class="badge badge-define">' + label + '</span>';
  }

  const isPlanInPlace = label.trim().toLowerCase() === 'plan in place';

  if (isPlanInPlace && isForecastOverdue(item)) {
    return '<span class="badge badge-plan-overdue">' + label + '</span>';
  }

  return '<span class="badge badge-plan">' + label + '</span>';
}

function sourceBadge(item) {
  if (item.sourceBase === 'ACTION_LOG') {
    return '<span class="badge badge-source-al" title="' + (item.sourceFile || '') + ' / ' + (item.sourceSheet || '') + '">ACTION LOG</span>';
  }
  return '<span class="badge badge-source-gpm" title="' + (item.sourceFile || '') + ' / ' + (item.sourceSheet || '') + '">GLOBAL PM</span>';
}

function sourceBadgeSimple(sourceBase) {
  if (sourceBase === 'ACTION_LOG') return '<span class="badge badge-source-al">AL</span>';
  return '<span class="badge badge-source-gpm">GPM</span>';
}

function truncate(str, n) {
  if (!str) return '-';
  return str.length > n ? str.slice(0, n) + '...' : str;
}

function updateKpiVisibilityForView(viewName) {
  const kpiRow = document.getElementById('kpiRow');
  const alertCardIds = ['kpiCardCritical', 'kpiCardUrgent', 'kpiCardAttention', 'kpiCardAlert'];

  if (viewName === 'l1') {
    alertCardIds.forEach(id => {
      const el = document.getElementById(id);
      if (el) el.style.display = 'none';
    });
    if (kpiRow) kpiRow.classList.add('kpi-row-compact');
  } else {
    alertCardIds.forEach(id => {
      const el = document.getElementById(id);
      if (el) el.style.display = '';
    });
    if (kpiRow) kpiRow.classList.remove('kpi-row-compact');
  }
}

function switchView(viewName) {
  currentView = viewName;
  allData = [...((dashboardPayload.views || {})[currentView] || [])];
  filteredData = [...allData];
  sortCol = null;
  sortDir = 1;
  document.querySelectorAll('thead th').forEach(t => {
    t.classList.remove('sorted');
    const icon = t.querySelector('.sort-icon');
    if (icon) icon.textContent = 'up-down';
  });
  updateKpiVisibilityForView(viewName);
  populateFilters();
  applyFilters();
}

function populateFilters() {
  const sources = [...new Set(allData.map(d => d.sourceBase))].sort();
  const projects = [...new Set(allData.map(d => d.project))].sort();
  const owners = [...new Set(allData.flatMap(d => d.ownersList || [d.owner]))].sort();
  const impacts = [...new Set(allData.map(d => d.impact))].sort();

  const sSel = document.getElementById('filterSource');
  const pSel = document.getElementById('filterProject');
  const oSel = document.getElementById('filterOwner');
  const iSel = document.getElementById('filterImpact');

  const prevSourceVal = sSel.value;
  const prevProjVal = pSel.value;
  const prevOwnerVal = oSel.value;
  const prevImpactVal = iSel.value;

  sSel.innerHTML = '<option value="">All Sources</option>';
  pSel.innerHTML = '<option value="">All Projects</option>';
  oSel.innerHTML = '<option value="">All Owners</option>';
  iSel.innerHTML = '<option value="">All Impacts</option>';

  sources.forEach(s => {
    const o = document.createElement('option');
    o.value = s;
    o.textContent = s === 'ACTION_LOG' ? 'Action Logs' : 'Global PM';
    sSel.appendChild(o);
  });

  projects.forEach(p => {
    const o = document.createElement('option');
    o.value = p;
    o.textContent = p;
    pSel.appendChild(o);
  });

  owners.forEach(p => {
    const o = document.createElement('option');
    o.value = p;
    o.textContent = p;
    oSel.appendChild(o);
  });

  impacts.forEach(p => {
    const o = document.createElement('option');
    o.value = p;
    o.textContent = p;
    iSel.appendChild(o);
  });

  if (sources.includes(prevSourceVal)) sSel.value = prevSourceVal;
  if (projects.includes(prevProjVal)) pSel.value = prevProjVal;
  if (owners.includes(prevOwnerVal)) oSel.value = prevOwnerVal;
  if (impacts.includes(prevImpactVal)) iSel.value = prevImpactVal;
}

function matchesSearch(item, search) {
  if (!search) return true;
  const ownerNames = (item.ownersList && item.ownersList.length) ? item.ownersList : [item.owner];
  const haystack = [
    item.project,
    item.wbs,
    item.impact,
    item.reportLevel,
    item.declaredLevel,
    item.sourceSheet,
    item.sourceFile,
    item.actionType,
    item.functionalOwner,
    item.recordId,
    item.projectCode,
    item.alertText,
    ...ownerNames
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();

  const searchTerms = search.toLowerCase().split(/\s+/).filter(Boolean);
  return searchTerms.every(term => {
    const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp('\\b' + escaped, 'i');
    return regex.test(haystack);
  });
}

function applyFilters() {
  const source = document.getElementById('filterSource').value;
  const proj = document.getElementById('filterProject').value;
  const owner = document.getElementById('filterOwner').value;
  const impact = document.getElementById('filterImpact').value;
  const urgency = document.getElementById('filterUrgency').value;
  const search = document.getElementById('filterSearch').value.trim();

  filteredData = allData.filter(item => {
    if (source && item.sourceBase !== source) return false;
    if (proj && item.project !== proj) return false;
    if (owner && !(item.ownersList || [item.owner]).includes(owner)) return false;
    if (impact && item.impact !== impact) return false;
    if (urgency && getAlert(item) !== urgency) return false;
    if (!matchesSearch(item, search)) return false;
    return true;
  });

  if (sortCol) {
    filteredData.sort((a, b) => {
      let va = a[sortCol], vb = b[sortCol];
      if (va === null || va === undefined) va = -Infinity;
      if (vb === null || vb === undefined) vb = -Infinity;
      if (typeof va === 'string') va = va.toLowerCase();
      if (typeof vb === 'string') vb = vb.toLowerCase();
      if (va < vb) return -sortDir;
      if (va > vb) return sortDir;
      return 0;
    });
  }

  renderKPIs();
  renderTable();
  renderCharts();
  renderAlertBanner();
}

function renderKPIs() {
  const total = filteredData.length;
  const critical = filteredData.filter(d => getAlert(d) === 'critical').length;
  const urgent = filteredData.filter(d => getAlert(d) === 'urgent').length;
  const attention = filteredData.filter(d => getAlert(d) === 'attention').length;
  const alertCount = filteredData.filter(d => getAlert(d) === 'alert').length;
  const plan = filteredData.filter(d => hasDefinedPlan(d)).length;

  document.getElementById('kpiTotal').textContent = total;
  document.getElementById('kpiCritical').textContent = critical;
  document.getElementById('kpiUrgent').textContent = urgent;
  document.getElementById('kpiAttention').textContent = attention;
  document.getElementById('kpiAlert').textContent = alertCount;
  document.getElementById('kpiPlan').textContent = plan;
  document.getElementById('tableCount').textContent = total;
  document.getElementById('tableTotal').textContent = allData.length;
}

function renderTable() {
  const tbody = document.getElementById('tableBody');
  const noData = document.getElementById('noData');
  tbody.innerHTML = '';
  expandedRows.clear();

  if (filteredData.length === 0) {
    noData.style.display = 'block';
    return;
  }
  noData.style.display = 'none';

  filteredData.forEach((item) => {
    const alertLevel = getAlert(item);
    const tr = document.createElement('tr');
    tr.className = 'fade-in';
    tr.dataset.id = item.id;

    const borderColor = BORDER_COLORS[alertLevel] || 'transparent';
    tr.setAttribute('style', 'border-left: 3px solid ' + borderColor + ';');

    tr.innerHTML = `
      <td style="color:var(--color-accent6);font-size:11px;">${item.id + 1}</td>
      <td>${sourceBadge(item)}</td>
      <td class="td-project">${item.project}</td>
      <td class="td-wbs" title="${item.wbs}">${truncate(item.wbs, 28)}</td>
      <td><span style="font-size:11px;color:var(--color-tint-2);">${item.impact}</span></td>
      <td style="white-space:nowrap;font-size:11px;">${item.owner}</td>
      <td style="white-space:nowrap;font-size:11px;">${item.coOwner || '-'}</td>
      <td>${statusBadge(item)}</td>
      <td style="white-space:nowrap;font-size:11px;">${item.reportLevel || '-'}</td>
      <td>${formatDays(item.daysOpen, item)}</td>
      <td style="white-space:nowrap;font-size:11px;">${formatDate(item.dueDate)}</td>
      <td style="white-space:nowrap;font-size:11px;">${formatDate(item.forecastDate)}</td>
      <td>${alertBadge(item)}</td>
      <td><button class="btn-expand" data-id="${item.id}">Details</button></td>
    `;

    tr.querySelector('.btn-expand').addEventListener('click', (e) => { e.stopPropagation(); toggleExpand(item.id, tr); });
    tr.addEventListener('click', () => toggleExpand(item.id, tr));
    tbody.appendChild(tr);
  });
}

function toggleExpand(id, tr) {
  const existingExpand = document.getElementById('expand-' + id);
  if (existingExpand) {
    existingExpand.remove();
    expandedRows.delete(id);
    tr.classList.remove('expanded');
    const btn = tr.querySelector('.btn-expand');
    if (btn) btn.textContent = 'Details';
    return;
  }

  const item = allData.find(d => d.id === id);
  if (!item) return;

  expandedRows.add(id);
  tr.classList.add('expanded');
  const btn = tr.querySelector('.btn-expand');
  if (btn) btn.textContent = 'Close';

  const extraSections = item.sourceBase === 'ACTION_LOG' ? `
        <div class="expand-section">
          <h4>Action Type</h4>
          <p>${item.actionType || '-'}</p>
        </div>
        <div class="expand-section">
          <h4>Functional Owner</h4>
          <p>${item.functionalOwner || '-'}</p>
        </div>
        <div class="expand-section">
          <h4>Record ID</h4>
          <p>${item.recordId || '-'}</p>
        </div>
        <div class="expand-section">
          <h4>Project Header Code</h4>
          <p>${item.projectCode || '-'}</p>
        </div>
  ` : '';

  const defined = hasDefinedPlan(item);
  const label = item.sourceBase === 'ACTION_LOG'
    ? (item.status || 'Open')
    : (defined ? 'Plan in Place' : 'To be Defined');
  const isPlanInPlace = defined && label.trim().toLowerCase() === 'plan in place';

  let forecastNote;
  if (!isPlanInPlace) {
    forecastNote = 'Not applicable - this rule only applies to items marked as "Plan in Place".';
  } else if (isForecastOverdue(item)) {
    forecastNote = 'Forecast Conclusion Date has already passed.';
  } else {
    forecastNote = 'Forecast Conclusion Date has not passed yet.';
  }

  const expandTr = document.createElement('tr');
  expandTr.className = 'expand-row';
  expandTr.id = 'expand-' + id;
  expandTr.innerHTML = `
    <td colspan="14">
      <div class="expand-content">
        <div class="expand-section">
          <h4>${item.sourceBase === 'ACTION_LOG' ? 'Action Description' : 'Ask / Critical Issue'}</h4>
          <p>${item.ask || '-'}</p>
        </div>
        <div class="expand-section">
          <h4>${item.sourceBase === 'ACTION_LOG' ? 'Comments' : 'Action Plan Detail'}</h4>
          <p>${item.plan || '(empty - plan not considered defined)'}</p>
        </div>
        <div class="expand-section">
          <h4>Alert</h4>
          <p>${item.alertText || 'No alert - item does not match Critical, Urgent, Attention or Alert criteria.'}</p>
        </div>
        <div class="expand-section">
          <h4>Escalation Level</h4>
          <p>${item.reportLevel || '-'} (fixed - does not change automatically)</p>
        </div>
        <div class="expand-section">
          <h4>Days Open</h4>
          <p>${item.daysOpen !== null && item.daysOpen !== undefined ? item.daysOpen : 'N/A'} — ${item.daysOpenSource || '-'}</p>
        </div>
        <div class="expand-section">
          <h4>Forecast Conclusion Date</h4>
          <p>${formatDate(item.forecastDate)} — ${forecastNote}</p>
        </div>
        <div class="expand-section">
          <h4>Source</h4>
          <p>${item.sourceBase === 'ACTION_LOG' ? 'Action Logs' : 'Global PM'} — ${item.sourceFile || '-'} / ${item.sourceSheet || '-'}</p>
        </div>
        ${extraSections}
      </div>
    </td>
  `;
  tr.after(expandTr);
}

function renderAlertBanner() {
  const criticalItems = filteredData
    .filter(d => {
      const u = getAlert(d);
      return u === 'critical' || u === 'urgent' || u === 'alert';
    })
    .sort((a, b) => (b.daysOpen || 0) - (a.daysOpen || 0))
    .slice(0, 8);

  const container = document.getElementById('alertItems');
  container.innerHTML = '';

  if (criticalItems.length === 0) {
    container.innerHTML = '<span style="font-size:12px;color:var(--color-accent6);">No critical items in current filter.</span>';
    return;
  }

  criticalItems.forEach(item => {
    const div = document.createElement('div');
    div.className = 'alert-item';
    div.innerHTML = `<strong>${item.project}</strong> - ${truncate(item.wbs, 30)} - <strong>${item.daysOpen ? item.daysOpen + 'd' : 'N/A'}</strong>`;
    div.addEventListener('click', () => {
      document.getElementById('filterProject').value = item.project;
      applyFilters();
    });
    container.appendChild(div);
  });
}

function renderSourceInfo() {
  const bar = document.getElementById('sourceInfoBar');
  const base = dashboardPayload.views.general || [];
  const groups = {};
  base.forEach(item => {
    const key = item.sourceBase + '|' + (item.sourceFile || '') + '|' + (item.sourceSheet || '');
    if (!groups[key]) {
      groups[key] = { sourceBase: item.sourceBase, sourceFile: item.sourceFile, sourceSheet: item.sourceSheet, count: 0 };
    }
    groups[key].count++;
  });
  const entries = Object.values(groups);
  if (entries.length === 0) {
    bar.innerHTML = '<span>No data sources loaded.</span>';
    return;
  }
  bar.innerHTML = entries.map(g => {
    const label = g.sourceBase === 'ACTION_LOG' ? 'Action Logs' : 'Global PM';
    return '<div class="source-info-item">' + sourceBadgeSimple(g.sourceBase) + '<strong>' + label + '</strong>: ' +
      (g.sourceFile || '-') + ' (' + (g.sourceSheet || '-') + ') — <strong>' + g.count + '</strong> items</div>';
  }).join('');
}

function renderComparisonBlock(data, suffix) {
  const tabBtn = document.getElementById('tabBtnComparison' + suffix);
  if (!data) {
    if (tabBtn) tabBtn.style.display = 'none';
    return;
  }
  if (tabBtn) tabBtn.style.display = 'inline-block';

  document.getElementById('cmpNew' + suffix).textContent = data.newItems.length;
  document.getElementById('cmpClosed' + suffix).textContent = data.closedItems.length;
  document.getElementById('cmpChanged' + suffix).textContent = data.changedItems.length;
  document.getElementById('cmpUnchanged' + suffix).textContent = data.unchangedCount;

  const newBody = document.getElementById('cmpNewBody' + suffix);
  newBody.innerHTML = '';
  if (data.newItems.length === 0) {
    document.getElementById('cmpNewNoData' + suffix).style.display = 'block';
  } else {
    document.getElementById('cmpNewNoData' + suffix).style.display = 'none';
    data.newItems.forEach(item => {
      const tr = document.createElement('tr');
      const borderColor = BORDER_COLORS[getAlert(item)] || 'transparent';
      tr.setAttribute('style', 'border-left: 3px solid ' + borderColor + ';');
      tr.innerHTML = `
        <td class="td-project">${item.project}</td>
        <td class="td-wbs" title="${item.wbs}">${truncate(item.wbs, 28)}</td>
        <td style="font-size:11px;color:var(--color-tint-2);">${item.impact}</td>
        <td style="white-space:nowrap;font-size:11px;">${item.owner}</td>
        <td>${formatDays(item.daysOpen, item)}</td>
        <td>${statusBadge(item)}</td>
        <td style="white-space:nowrap;font-size:11px;">${item.reportLevel || '-'}</td>
        <td>${alertBadge(item)}</td>
      `;
      newBody.appendChild(tr);
    });
  }

  const closedBody = document.getElementById('cmpClosedBody' + suffix);
  closedBody.innerHTML = '';
  if (data.closedItems.length === 0) {
    document.getElementById('cmpClosedNoData' + suffix).style.display = 'block';
  } else {
    document.getElementById('cmpClosedNoData' + suffix).style.display = 'none';
    data.closedItems.forEach(item => {
      const tr = document.createElement('tr');
      const borderColor = BORDER_COLORS[getAlert(item)] || 'transparent';
      tr.setAttribute('style', 'border-left: 3px solid ' + borderColor + ';');
      tr.innerHTML = `
        <td class="td-project">${item.project}</td>
        <td class="td-wbs" title="${item.wbs}">${truncate(item.wbs, 28)}</td>
        <td style="font-size:11px;color:var(--color-tint-2);">${item.impact}</td>
        <td style="white-space:nowrap;font-size:11px;">${item.owner}</td>
        <td>${formatDays(item.daysOpen, item)}</td>
        <td>${statusBadge(item)}</td>
        <td style="white-space:nowrap;font-size:11px;">${item.reportLevel || '-'}</td>
        <td>${alertBadge(item)}</td>
      `;
      closedBody.appendChild(tr);
    });
  }

  const changedBody = document.getElementById('cmpChangedBody' + suffix);
  changedBody.innerHTML = '';
  let totalDiffRows = 0;
  const dateFields = ['Due Date', 'Forecast Conclusion Date'];

  data.changedItems.forEach(entry => {
    Object.values(entry.diffs).forEach(diff => {
      totalDiffRows++;
      const tr = document.createElement('tr');
      let prevVal = (diff.previous === null || diff.previous === undefined || diff.previous === '') ? '-' : diff.previous;
      let curVal = (diff.current === null || diff.current === undefined || diff.current === '') ? '-' : diff.current;
      if (dateFields.includes(diff.label)) {
        prevVal = formatDate(diff.previous);
        curVal = formatDate(diff.current);
      }
      tr.innerHTML = `
        <td class="td-project">${entry.project}</td>
        <td class="td-wbs" title="${entry.wbs}">${truncate(entry.wbs, 28)}</td>
        <td style="font-size:11px;color:var(--color-tint-2);">${diff.label}</td>
        <td class="diff-prev">${prevVal}</td>
        <td class="diff-cur">${curVal}</td>
      `;
      changedBody.appendChild(tr);
    });
  });

  document.getElementById('cmpChangedNoData' + suffix).style.display = totalDiffRows === 0 ? 'block' : 'none';
}

function initTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      if (btn.style.display === 'none') return;

      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      if (btn.dataset.view) {
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        const dashboardTab = document.getElementById('tabContentDashboard');
        if (dashboardTab) dashboardTab.classList.add('active');
        switchView(btn.dataset.view);
        return;
      }

      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      const target = document.getElementById(btn.dataset.target);
      if (target) target.classList.add('active');
    });
  });
}

const CHART_COLORS = {
  primary: '#1E2A47', tint1: '#4C6A92', tint2: '#93A9C9', tint3: '#C7D3E6',
  critical: '#E63B2E', urgent: '#F5A623', attention: '#F2C94C', alert: '#5FB7FF',
  forest: '#7BDD3C', earth: '#774936', darkSky: '#2176FF'
};
const IMPACT_COLORS = [CHART_COLORS.critical, CHART_COLORS.urgent, CHART_COLORS.attention, CHART_COLORS.alert, CHART_COLORS.tint1, CHART_COLORS.forest];

const chartDefaults = {
  plugins: {
    legend: { labels: { color: '#93A9C9', font: { family: 'Inter', size: 11 }, boxWidth: 12, padding: 10 } },
    tooltip: { backgroundColor: '#16213A', titleColor: '#5EEAD4', bodyColor: '#FFFFFF', borderColor: '#4C6A92', borderWidth: 1, padding: 10, titleFont: { family: 'Inter', weight: '700', size: 12 }, bodyFont: { family: 'Inter', size: 11 } }
  },
  scales: {
    x: { ticks: { color: '#93A9C9', font: { family: 'Inter', size: 10 } }, grid: { color: 'rgba(147,169,201,0.1)' } },
    y: { ticks: { color: '#93A9C9', font: { family: 'Inter', size: 10 } }, grid: { color: 'rgba(147,169,201,0.1)' } }
  }
};

function destroyChart(key) {
  if (charts[key]) {
    charts[key].destroy();
    delete charts[key];
  }
}

function renderCharts() {
  renderChartProject();
  renderChartImpact();
  renderChartOwner();
  renderChartDays();
}

function renderChartProject() {
  destroyChart('project');
  const projects = [...new Set(allData.map(d => d.project))].sort();
  const critCounts = projects.map(p => filteredData.filter(d => d.project === p && getAlert(d) === 'critical').length);
  const urgCounts = projects.map(p => filteredData.filter(d => d.project === p && getAlert(d) === 'urgent').length);
  const attCounts = projects.map(p => filteredData.filter(d => d.project === p && getAlert(d) === 'attention').length);
  const alertCounts = projects.map(p => filteredData.filter(d => d.project === p && getAlert(d) === 'alert').length);
  const naCounts = projects.map(p => filteredData.filter(d => d.project === p && getAlert(d) === 'na').length);

  const ctx = document.getElementById('chartProject').getContext('2d');
  charts['project'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: projects.map(p => p.length > 16 ? p.slice(0, 16) + '...' : p),
      datasets: [
        { label: 'Critical', data: critCounts, backgroundColor: CHART_COLORS.critical, borderRadius: 2 },
        { label: 'Urgent', data: urgCounts, backgroundColor: CHART_COLORS.urgent, borderRadius: 2 },
        { label: 'Attention', data: attCounts, backgroundColor: CHART_COLORS.attention, borderRadius: 2 },
        { label: 'Alert', data: alertCounts, backgroundColor: CHART_COLORS.alert, borderRadius: 2 },
        { label: 'N/A', data: naCounts, backgroundColor: CHART_COLORS.tint1, borderRadius: 2 }
      ]
    },
    options: {
      maintainAspectRatio: false, indexAxis: 'y',
      plugins: { ...chartDefaults.plugins },
      scales: {
        x: { stacked: true, ticks: { color: '#93A9C9', font: { family: 'Inter', size: 10 } }, grid: { color: 'rgba(147,169,201,0.1)' } },
        y: { stacked: true, ticks: { color: '#93A9C9', font: { family: 'Inter', size: 10 } }, grid: { color: 'rgba(147,169,201,0.08)' } }
      },
      onClick: (e, elements) => {
        if (elements.length > 0) {
          const idx = elements[0].index;
          document.getElementById('filterProject').value = projects[idx];
          applyFilters();
        }
      }
    }
  });
}

function renderChartImpact() {
  destroyChart('impact');
  const impacts = [...new Set(allData.map(d => d.impact))];
  const counts = impacts.map(i => filteredData.filter(d => d.impact === i).length);
  const ctx = document.getElementById('chartImpact').getContext('2d');
  charts['impact'] = new Chart(ctx, {
    type: 'doughnut',
    data: { labels: impacts, datasets: [{ data: counts, backgroundColor: IMPACT_COLORS, borderColor: '#16213A', borderWidth: 2, hoverOffset: 6 }] },
    options: {
      maintainAspectRatio: false, cutout: '62%',
      plugins: {
        legend: { position: 'right', labels: { color: '#93A9C9', font: { family: 'Inter', size: 10 }, boxWidth: 10, padding: 8 } },
        tooltip: { ...chartDefaults.plugins.tooltip }
      },
      onClick: (e, elements) => {
        if (elements.length > 0) {
          const idx = elements[0].index;
          document.getElementById('filterImpact').value = impacts[idx];
          applyFilters();
        }
      }
    }
  });
}

function renderChartOwner() {
  destroyChart('owner');

  const ownerCounts = {};
  filteredData.forEach(item => {
    const owners = item.ownersList || [item.owner];
    owners.forEach(owner => {
      if (!owner) return;
      ownerCounts[owner] = (ownerCounts[owner] || 0) + 1;
    });
  });

  const owners = Object.keys(ownerCounts).sort();
  const counts = owners.map(o => ownerCounts[o]);

  const ctx = document.getElementById('chartOwner').getContext('2d');
  charts['owner'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: owners.map(o => o.length > 18 ? o.slice(0, 18) + '...' : o),
      datasets: [{ label: 'Items', data: counts, backgroundColor: counts.map((c, i) => i % 2 === 0 ? CHART_COLORS.primary : CHART_COLORS.tint1), borderRadius: 2 }]
    },
    options: {
      maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: { ...chartDefaults.plugins.tooltip } },
      scales: {
        x: { ticks: { color: '#93A9C9', font: { family: 'Inter', size: 10 } }, grid: { color: 'rgba(147,169,201,0.1)' } },
        y: { ticks: { color: '#93A9C9', font: { family: 'Inter', size: 10 } }, grid: { color: 'rgba(147,169,201,0.08)' } }
      },
      onClick: (e, elements) => {
        if (elements.length > 0) {
          const idx = elements[0].index;
          document.getElementById('filterOwner').value = owners[idx];
          applyFilters();
        }
      }
    }
  });
}

function renderChartDays() {
  destroyChart('days');
  const sorted = [...filteredData]
    .filter(d => d.daysOpen !== null && d.daysOpen !== undefined && d.daysOpen < 50000)
    .sort((a, b) => b.daysOpen - a.daysOpen)
    .slice(0, 20);

  const labels = sorted.map(d => truncate(d.wbs, 18));
  const values = sorted.map(d => d.daysOpen);
  const colors = sorted.map(d => {
    const u = getAlert(d);
    if (u === 'critical') return CHART_COLORS.critical;
    if (u === 'urgent') return CHART_COLORS.urgent;
    if (u === 'attention') return CHART_COLORS.attention;
    if (u === 'alert') return CHART_COLORS.alert;
    return CHART_COLORS.tint2;
  });

  const ctx = document.getElementById('chartDays').getContext('2d');
  charts['days'] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: 'Days Open', data: values, backgroundColor: colors, borderRadius: 2 }] },
    options: {
      maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: { ...chartDefaults.plugins.tooltip } },
      scales: {
        x: { ticks: { color: '#93A9C9', font: { family: 'Inter', size: 10 } }, grid: { color: 'rgba(147,169,201,0.1)' } },
        y: { ticks: { color: '#93A9C9', font: { family: 'Inter', size: 9 } }, grid: { color: 'rgba(147,169,201,0.08)' } }
      }
    }
  });
}

document.querySelectorAll('thead th[data-col]').forEach(th => {
  th.addEventListener('click', () => {
    const col = th.dataset.col;
    if (sortCol === col) { sortDir *= -1; } else { sortCol = col; sortDir = 1; }
    document.querySelectorAll('thead th').forEach(t => t.classList.remove('sorted'));
    th.classList.add('sorted');
    const icon = th.querySelector('.sort-icon');
    if (icon) icon.textContent = sortDir === 1 ? 'up' : 'down';
    applyFilters();
  });
});

document.getElementById('filterSource').addEventListener('change', applyFilters);
document.getElementById('filterProject').addEventListener('change', applyFilters);
document.getElementById('filterOwner').addEventListener('change', applyFilters);
document.getElementById('filterImpact').addEventListener('change', applyFilters);
document.getElementById('filterUrgency').addEventListener('change', applyFilters);
document.getElementById('filterSearch').addEventListener('input', applyFilters);

document.getElementById('btnReset').addEventListener('click', () => {
  document.getElementById('filterSource').value = '';
  document.getElementById('filterProject').value = '';
  document.getElementById('filterOwner').value = '';
  document.getElementById('filterImpact').value = '';
  document.getElementById('filterUrgency').value = '';
  document.getElementById('filterSearch').value = '';
  sortCol = null;
  sortDir = 1;
  document.querySelectorAll('thead th').forEach(t => {
    t.classList.remove('sorted');
    const icon = t.querySelector('.sort-icon');
    if (icon) icon.textContent = 'up-down';
  });
  applyFilters();
});

function initFloatingScrollbar() {
  const floatBar = document.getElementById('floatingHScroll');
  const floatInner = document.getElementById('floatingHScrollInner');
  if (!floatBar || !floatInner) return;

  let syncing = false;
  let activeTarget = null;

  function pickTarget() {
    const candidates = Array.from(document.querySelectorAll('.table-scroll'));
    const viewportHeight = window.innerHeight;
    let best = null;
    let bestVisibleHeight = 0;

    candidates.forEach(el => {
      const needsScroll = el.scrollWidth > el.clientWidth + 2;
      if (!needsScroll) return;
      const rect = el.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) return;
      const visibleTop = Math.max(rect.top, 0);
      const visibleBottom = Math.min(rect.bottom, viewportHeight);
      const visibleHeight = visibleBottom - visibleTop;
      const scrollbarAlreadyVisible = rect.bottom <= viewportHeight;
      if (visibleHeight > 0 && !scrollbarAlreadyVisible && visibleHeight > bestVisibleHeight) {
        bestVisibleHeight = visibleHeight;
        best = el;
      }
    });
    return best;
  }

  function detachFromTarget(target) {
    if (target && target._floatSyncHandler) {
      target.removeEventListener('scroll', target._floatSyncHandler);
      target._floatSyncHandler = null;
    }
  }

  function attachToTarget(target) {
    const handler = () => {
      if (syncing) return;
      syncing = true;
      floatBar.scrollLeft = target.scrollLeft;
      syncing = false;
    };
    target._floatSyncHandler = handler;
    target.addEventListener('scroll', handler);
  }

  function updateBar() {
    const target = pickTarget();

    if (target !== activeTarget) {
      detachFromTarget(activeTarget);
      activeTarget = target;
      if (activeTarget) attachToTarget(activeTarget);
    }

    if (!activeTarget) {
      floatBar.classList.remove('visible');
      return;
    }

    const rect = activeTarget.getBoundingClientRect();
    floatBar.style.left = rect.left + 'px';
    floatBar.style.width = rect.width + 'px';
    floatInner.style.width = activeTarget.scrollWidth + 'px';
    floatBar.classList.add('visible');

    if (!syncing) {
      syncing = true;
      floatBar.scrollLeft = activeTarget.scrollLeft;
      syncing = false;
    }
  }

  floatBar.addEventListener('scroll', () => {
    if (syncing || !activeTarget) return;
    syncing = true;
    activeTarget.scrollLeft = floatBar.scrollLeft;
    syncing = false;
  });

  window.addEventListener('scroll', updateBar, { passive: true });
  window.addEventListener('resize', updateBar);

  const observer = new MutationObserver(() => updateBar());
  document.querySelectorAll('.table-scroll').forEach(el => {
    observer.observe(el, { childList: true, subtree: true });
  });

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => setTimeout(updateBar, 50));
  });

  updateBar();
  setInterval(updateBar, 400);
}

populateFilters();
applyFilters();
renderSourceInfo();
renderComparisonBlock(comparisonDataGPM, 'GPM');
renderComparisonBlock(comparisonDataAL, 'AL');
initTabs();
initFloatingScrollbar();
</script>
<footer id="print-footer"><div class="source-files">Source files: __SOURCE_FILES__</div><div class="generated-date">Generated by Action Tracker Dashboard Generator on __GENERATED_DATE__</div></footer>
</body></html>
"""


def render_html(dashboard_payload, source_files_label,
                 comparison_gpm=None, comparison_al=None,
                 current_sheet_gpm=None, previous_sheet_gpm=None,
                 current_sheet_al=None, previous_sheet_al=None):
    data_json = json.dumps(dashboard_payload, ensure_ascii=False)
    comparison_gpm_json = json.dumps(comparison_gpm, ensure_ascii=False) if comparison_gpm else "null"
    comparison_al_json = json.dumps(comparison_al, ensure_ascii=False) if comparison_al else "null"
    now = datetime.now()
    out = HTML_TEMPLATE
    out = out.replace("__DATA_JSON__", data_json)
    out = out.replace("__COMPARISON_GPM_JSON__", comparison_gpm_json)
    out = out.replace("__COMPARISON_AL_JSON__", comparison_al_json)
    out = out.replace("__CURRENT_SHEET_GPM__", html.escape(str(current_sheet_gpm or "N/A")))
    out = out.replace("__PREVIOUS_SHEET_GPM__", html.escape(str(previous_sheet_gpm or "N/A")))
    out = out.replace("__CURRENT_SHEET_AL__", html.escape(str(current_sheet_al or "N/A")))
    out = out.replace("__PREVIOUS_SHEET_AL__", html.escape(str(previous_sheet_al or "N/A")))
    out = out.replace("__LAST_UPDATED__", now.strftime("%B %Y"))
    out = out.replace("__GENERATED_DATE__", now.strftime("%d/%m/%Y %H:%M"))
    out = out.replace("__SOURCE_FILES__", html.escape(source_files_label))
    out = out.replace("__TOTAL_COUNT__", str(len(dashboard_payload["views"]["general"])))
    out = out.replace("__TITLE__", DASHBOARD_TITLE)
    return out


def load_base_current_week(xlsx_path, sheet_name, base_key):
    if not xlsx_path:
        return None, []

    if base_key == SOURCE_GLOBAL_PM:
        column_map = GLOBAL_PM_COLUMN_MAP
        required_keys = GLOBAL_PM_REQUIRED_KEYS
    else:
        column_map = ACTION_LOG_COLUMN_MAP
        required_keys = ACTION_LOG_REQUIRED_KEYS

    payload, records = prepare_week_data(xlsx_path, sheet_name, column_map, required_keys, base_key)
    return payload, records


def load_base_previous_week(xlsx_path, sheet_name, base_key):
    return load_base_current_week(xlsx_path, sheet_name, base_key)


def build_combined_payload(gpm_records, al_records):
    combined = list(gpm_records) + list(al_records)
    combined = renumber_records(combined)
    return build_dashboard_payload(combined)


class DashboardApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("820x760")
        self.root.minsize(760, 640)
        self.root.configure(bg=BG_COLOR)
        self.root.resizable(True, True)

        self.gpm_xlsx_path = None
        self.al_xlsx_path = None
        self.output_path = None

        self.gpm_sheet_var = StringVar(value=DEFAULT_SHEET)
        self.gpm_prev_sheet_var = StringVar(value=DEFAULT_PREVIOUS_SHEET)
        self.al_sheet_var = StringVar(value=DEFAULT_SHEET)
        self.al_prev_sheet_var = StringVar(value=DEFAULT_PREVIOUS_SHEET)

        self.compare_var = BooleanVar(value=True)
        self.use_action_log_var = BooleanVar(value=False)

        self._center_window()
        self._build_ui()

    def _center_window(self):
        self.root.update_idletasks()
        width = 820
        height = 760
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = max((screen_width // 2) - (width // 2), 0)
        y = max((screen_height // 2) - (height // 2), 0)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _build_ui(self):
        APP_BG = BG_COLOR
        CARD_BG = "#16213A"
        CARD_INNER_BG = "#1C2D4D"
        ACCENT = ACCENT_COLOR
        TEXT = TEXT_COLOR
        MUTED = "#C7D3E6"
        SOFT = "#93A9C9"
        ENTRY_BG = "#F5F7F7"
        ENTRY_FG = "#1E293B"
        LOG_BG = "#0F1B33"

        style = ttk.Style()
        style.theme_use("clam")

        style.configure(".", background=APP_BG, foreground=TEXT, font=("Segoe UI", 10))
        style.configure("App.TFrame", background=APP_BG)
        style.configure("Card.TFrame", background=CARD_BG)
        style.configure("Inner.TFrame", background=CARD_BG)

        style.configure("Header.TLabel", background=APP_BG, foreground=ACCENT, font=("Segoe UI", 18, "bold"))
        style.configure("SubHeader.TLabel", background=APP_BG, foreground=MUTED, font=("Segoe UI", 9))
        style.configure("Body.TLabel", background=CARD_BG, foreground=TEXT, font=("Segoe UI", 10))
        style.configure("Hint.TLabel", background=CARD_BG, foreground=SOFT, font=("Segoe UI", 9))
        style.configure("StatusTitle.TLabel", background=APP_BG, foreground=TEXT, font=("Segoe UI", 10, "bold"))

        style.configure(
            "Card.TLabelframe",
            background=CARD_BG,
            bordercolor="#5C7A9E",
            relief="solid",
            borderwidth=1
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=CARD_BG,
            foreground=ACCENT,
            font=("Segoe UI", 10, "bold")
        )

        style.configure(
            "TEntry",
            fieldbackground=ENTRY_BG,
            foreground=ENTRY_FG,
            bordercolor="#6C86A8",
            lightcolor="#6C86A8",
            darkcolor="#6C86A8",
            padding=5
        )

        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(14, 8),
            foreground="#1A1A1A",
            background=ACCENT,
            borderwidth=0
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#7FF4DC"), ("disabled", "#6B7A8F")],
            foreground=[("disabled", "#EEEEEE")]
        )

        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 9, "bold"),
            padding=(10, 6),
            foreground=TEXT,
            background=CARD_INNER_BG,
            borderwidth=1,
            bordercolor="#6C86A8"
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#233955"), ("disabled", "#5A6B85")],
            foreground=[("disabled", "#D7E0E0")]
        )

        style.configure(
            "TCheckbutton",
            background=CARD_BG,
            foreground=TEXT,
            font=("Segoe UI", 10)
        )

        outer = ttk.Frame(self.root, style="App.TFrame")
        outer.pack(fill="both", expand=True)

        self.canvas = __import__("tkinter").Canvas(
            outer,
            bg=APP_BG,
            highlightthickness=0,
            bd=0
        )
        self.v_scroll = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set)

        self.v_scroll.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.scrollable_frame = ttk.Frame(self.canvas, style="App.TFrame", padding=12)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.scrollable_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self._bind_mousewheel()

        main = self.scrollable_frame

        ttk.Label(main, text="Action Tracker Dashboard Generator", style="Header.TLabel").pack(pady=(0, 2))
        ttk.Label(
            main,
            text="Load Global Project Management and (optionally) Project Action Logs to generate the dashboard.",
            style="SubHeader.TLabel"
        ).pack()
        ttk.Label(
            main,
            text="No data is sent to the internet — everything runs on your computer.",
            style="SubHeader.TLabel"
        ).pack(pady=(0, 10))

        frame_gpm = ttk.LabelFrame(
            main,
            text="Base 1 — GLOBAL PROJECT MANAGEMENT (.xlsx)",
            style="Card.TLabelframe",
            padding=10
        )
        frame_gpm.pack(fill="x", pady=(0, 8))

        ttk.Label(
            frame_gpm,
            text="Select the GLOBAL- PROJECT MANAGEMENT.xlsx spreadsheet.",
            style="Hint.TLabel"
        ).pack(anchor="w", pady=(0, 6))

        gpm_file_row = ttk.Frame(frame_gpm, style="Inner.TFrame")
        gpm_file_row.pack(fill="x")

        self.lbl_gpm_file = ttk.Label(
            gpm_file_row, text="No file selected.", style="Body.TLabel",
            wraplength=500, justify="left"
        )
        self.lbl_gpm_file.pack(side="left", fill="x", expand=True, padx=(0, 12))

        ttk.Button(
            gpm_file_row, text="Select .xlsx", style="Secondary.TButton",
            command=self.select_gpm_file
        ).pack(side="right")

        gpm_grid = ttk.Frame(frame_gpm, style="Inner.TFrame")
        gpm_grid.pack(fill="x", pady=(8, 0))
        gpm_grid.columnconfigure(1, weight=1)

        ttk.Label(gpm_grid, text="Current Week Sheet:", style="Body.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 12), pady=4
        )
        ttk.Entry(gpm_grid, textvariable=self.gpm_sheet_var).grid(row=0, column=1, sticky="ew", pady=4)

        self.lbl_gpm_prev_sheet = ttk.Label(gpm_grid, text="Previous Week Sheet:", style="Body.TLabel")
        self.lbl_gpm_prev_sheet.grid(row=1, column=0, sticky="w", padx=(0, 12), pady=4)
        self.entry_gpm_prev_sheet = ttk.Entry(gpm_grid, textvariable=self.gpm_prev_sheet_var)
        self.entry_gpm_prev_sheet.grid(row=1, column=1, sticky="ew", pady=4)

        frame_al = ttk.LabelFrame(
            main,
            text="Base 2 — PROJECT ACTION LOGS (.xlsx) — optional",
            style="Card.TLabelframe",
            padding=10
        )
        frame_al.pack(fill="x", pady=(0, 8))

        self.chk_use_al = ttk.Checkbutton(
            frame_al,
            text="Include Project Action Logs in this dashboard",
            variable=self.use_action_log_var,
            style="TCheckbutton",
            command=self._on_toggle_use_al
        )
        self.chk_use_al.pack(anchor="w", pady=(0, 6))

        al_file_row = ttk.Frame(frame_al, style="Inner.TFrame")
        al_file_row.pack(fill="x")

        self.lbl_al_file = ttk.Label(
            al_file_row, text="No file selected.", style="Body.TLabel",
            wraplength=500, justify="left"
        )
        self.lbl_al_file.pack(side="left", fill="x", expand=True, padx=(0, 12))

        self.btn_select_al = ttk.Button(
            al_file_row, text="Select .xlsx", style="Secondary.TButton",
            command=self.select_al_file
        )
        self.btn_select_al.pack(side="right")

        al_grid = ttk.Frame(frame_al, style="Inner.TFrame")
        al_grid.pack(fill="x", pady=(8, 0))
        al_grid.columnconfigure(1, weight=1)

        self.lbl_al_sheet = ttk.Label(al_grid, text="Current Week Sheet:", style="Body.TLabel")
        self.lbl_al_sheet.grid(row=0, column=0, sticky="w", padx=(0, 12), pady=4)
        self.entry_al_sheet = ttk.Entry(al_grid, textvariable=self.al_sheet_var)
        self.entry_al_sheet.grid(row=0, column=1, sticky="ew", pady=4)

        self.lbl_al_prev_sheet = ttk.Label(al_grid, text="Previous Week Sheet:", style="Body.TLabel")
        self.lbl_al_prev_sheet.grid(row=1, column=0, sticky="w", padx=(0, 12), pady=4)
        self.entry_al_prev_sheet = ttk.Entry(al_grid, textvariable=self.al_prev_sheet_var)
        self.entry_al_prev_sheet.grid(row=1, column=1, sticky="ew", pady=4)

        frame_compare = ttk.LabelFrame(
            main,
            text="Weekly Comparison Settings",
            style="Card.TLabelframe",
            padding=10
        )
        frame_compare.pack(fill="x", pady=(0, 10))

        self.chk_compare = ttk.Checkbutton(
            frame_compare,
            text="Include weekly comparison",
            variable=self.compare_var,
            style="TCheckbutton",
            command=self._on_toggle_compare
        )
        self.chk_compare.pack(anchor="w")

        self.btn_generate = ttk.Button(
            main,
            text="Generate HTML Dashboard",
            style="Primary.TButton",
            command=self.on_generate_clicked
        )
        self.btn_generate.pack(pady=(0, 10))
        self.btn_generate.state(["disabled"])

        ttk.Label(main, text="Status", style="StatusTitle.TLabel").pack(anchor="w", pady=(0, 4))

        self.txt_log = Text(
            main,
            height=7,
            bg=LOG_BG,
            fg="#DCE6F5",
            insertbackground="#FFFFFF",
            font=("Consolas", 9),
            state=DISABLED,
            wrap="word",
            relief="solid",
            bd=1,
            padx=8,
            pady=8
        )
        self.txt_log.pack(fill="both", expand=True, pady=(0, 8))

        frame_actions = ttk.Frame(main, style="App.TFrame")
        frame_actions.pack(pady=(0, 4))

        self.btn_open = ttk.Button(
            frame_actions, text="Open Dashboard", style="Secondary.TButton",
            command=self.open_output
        )
        self.btn_open.pack(side="left", padx=6)
        self.btn_open.state(["disabled"])

        self.btn_folder = ttk.Button(
            frame_actions, text="Open Folder", style="Secondary.TButton",
            command=self.open_output_folder
        )
        self.btn_folder.pack(side="left", padx=6)
        self.btn_folder.state(["disabled"])

        self._on_toggle_compare()
        self._on_toggle_use_al()

    def _on_frame_configure(self, event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _bind_mousewheel(self):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_linux)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_linux)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")

    def _on_toggle_compare(self):
        state = ["!disabled"] if self.compare_var.get() else ["disabled"]
        style_on = "Body.TLabel"
        style_off = "Hint.TLabel"

        self.entry_gpm_prev_sheet.state(state)
        self.lbl_gpm_prev_sheet.configure(style=style_on if self.compare_var.get() else style_off)

        if self.use_action_log_var.get():
            self.entry_al_prev_sheet.state(state)
            self.lbl_al_prev_sheet.configure(style=style_on if self.compare_var.get() else style_off)

    def _on_toggle_use_al(self):
        enabled = self.use_action_log_var.get()
        state = ["!disabled"] if enabled else ["disabled"]

        self.btn_select_al.state(state)
        self.entry_al_sheet.state(state)
        self.lbl_al_sheet.configure(style="Body.TLabel" if enabled else "Hint.TLabel")

        compare_enabled = enabled and self.compare_var.get()
        prev_state = ["!disabled"] if compare_enabled else ["disabled"]
        self.entry_al_prev_sheet.state(prev_state)
        self.lbl_al_prev_sheet.configure(style="Body.TLabel" if compare_enabled else "Hint.TLabel")

        self._update_generate_button_state()

    def _update_generate_button_state(self):
        if self.gpm_xlsx_path:
            self.btn_generate.state(["!disabled"])
        else:
            self.btn_generate.state(["disabled"])

    def log(self, message):
        self.txt_log.configure(state=NORMAL)
        self.txt_log.insert(END, message + NEWLINE)
        self.txt_log.see(END)
        self.txt_log.configure(state=DISABLED)
        self.root.update_idletasks()

    def select_gpm_file(self):
        path = filedialog.askopenfilename(
            title="Select the GLOBAL- PROJECT MANAGEMENT.xlsx file",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not path:
            return
        self.gpm_xlsx_path = Path(path)
        self.lbl_gpm_file.configure(text="Selected file: " + self.gpm_xlsx_path.name)
        self.log("Selected Global PM file: " + str(self.gpm_xlsx_path))
        self._update_generate_button_state()

    def select_al_file(self):
        path = filedialog.askopenfilename(
            title="Select the PROJECT ACTION LOGS.xlsx file",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not path:
            return
        self.al_xlsx_path = Path(path)
        self.lbl_al_file.configure(text="Selected file: " + self.al_xlsx_path.name)
        self.log("Selected Action Logs file: " + str(self.al_xlsx_path))

    def on_generate_clicked(self):
        if not self.gpm_xlsx_path:
            messagebox.showwarning(APP_TITLE, "Please select the Global Project Management .xlsx file first.")
            return

        if self.use_action_log_var.get() and not self.al_xlsx_path:
            messagebox.showwarning(
                APP_TITLE,
                "You enabled Project Action Logs but did not select its .xlsx file."
            )
            return

        save_path = filedialog.asksaveasfilename(
            title="Save dashboard as...",
            defaultextension=".html",
            initialfile="dashboard_critical_actions.html",
            filetypes=[("HTML files", "*.html")]
        )
        if not save_path:
            return

        self.output_path = Path(save_path)
        self.btn_generate.state(["disabled"])
        self.btn_open.state(["disabled"])
        self.btn_folder.state(["disabled"])
        self.log("Processing... this may take a few seconds.")
        threading.Thread(target=self._generate_worker, daemon=True).start()

    def _generate_worker(self):
        try:
            gpm_sheet = self.gpm_sheet_var.get().strip() or DEFAULT_SHEET
            gpm_prev_sheet = self.gpm_prev_sheet_var.get().strip()
            include_comparison = self.compare_var.get()
            use_al = self.use_action_log_var.get()

            self.log("Reading Global PM current week sheet: " + gpm_sheet)
            gpm_payload, gpm_clean = load_base_current_week(self.gpm_xlsx_path, gpm_sheet, SOURCE_GLOBAL_PM)
            self.log("Global PM items loaded: " + str(gpm_payload["totalCount"]))
            self.log(
                "Global PM views: Overview=" + str(len(gpm_payload["views"]["general"])) +
                " | L1=" + str(len(gpm_payload["views"]["l1"])) +
                " | L2=" + str(len(gpm_payload["views"]["l2"])) +
                " | L3=" + str(len(gpm_payload["views"]["l3"]))
            )

            gpm_calculated = sum(1 for r in gpm_clean if r.get("daysOpenSource", "").startswith("Calculated"))
            if gpm_calculated:
                self.log(
                    "Global PM: Days Open calculated automatically (from L2 Created) for " +
                    str(gpm_calculated) + " item(s) missing L2 Days Open."
                )
            gpm_plan_defined = sum(1 for r in gpm_clean if r.get("planDefined"))
            self.log(
                "Global PM: " + str(gpm_plan_defined) + " item(s) have text in Action Plan Detail "
                "(considered as plan defined)."
            )

            al_clean = []
            if use_al:
                al_sheet = self.al_sheet_var.get().strip() or DEFAULT_SHEET
                self.log("Reading Action Logs current week sheet: " + al_sheet)
                al_payload, al_clean = load_base_current_week(self.al_xlsx_path, al_sheet, SOURCE_ACTION_LOG)
                self.log("Action Logs items loaded: " + str(al_payload["totalCount"]))
                self.log(
                    "Action Logs views: Overview=" + str(len(al_payload["views"]["general"])) +
                    " | L1=" + str(len(al_payload["views"]["l1"])) +
                    " | L2=" + str(len(al_payload["views"]["l2"])) +
                    " | L3=" + str(len(al_payload["views"]["l3"]))
                )
                al_calculated = sum(1 for r in al_clean if r.get("daysOpenSource", "").startswith("Calculated"))
                self.log(
                    "Action Logs: Days Open calculated from BL Due Date for " +
                    str(al_calculated) + " item(s)."
                )
                al_plan_defined = sum(1 for r in al_clean if r.get("planDefined"))
                self.log(
                    "Action Logs: " + str(al_plan_defined) + " item(s) have text in Comments "
                    "(considered as plan defined)."
                )
                al_functional_owner_l1 = sum(
                    1 for r in al_clean
                    if (r.get("declaredLevel") or "").upper() == "L1" and (r.get("functionalOwner") or "").strip()
                )
                self.log(
                    "Action Logs: " + str(al_functional_owner_l1) + " L1 item(s) now displayed with "
                    "Functional Owner instead of individual Owner."
                )
            else:
                al_sheet = None

            combined_payload = build_combined_payload(gpm_clean, al_clean)
            if use_al:
                self.log(
                    "Combined dashboard: " + str(combined_payload["totalCount"]) +
                    " items total (Global PM + Action Logs)."
                )
            else:
                self.log(
                    "Combined dashboard: " + str(combined_payload["totalCount"]) +
                    " items total (Global PM only)."
                )

            comparison_gpm = None
            used_gpm_prev_sheet = None
            if include_comparison and gpm_prev_sheet:
                try:
                    self.log("Reading Global PM previous week sheet: " + gpm_prev_sheet)
                    _, gpm_previous_clean = load_base_previous_week(
                        self.gpm_xlsx_path, gpm_prev_sheet, SOURCE_GLOBAL_PM
                    )
                    comparison_gpm = build_comparison_data(gpm_clean, gpm_previous_clean)
                    used_gpm_prev_sheet = gpm_prev_sheet
                    self.log(
                        "Global PM comparison: " + str(len(comparison_gpm["newItems"])) + " new, " +
                        str(len(comparison_gpm["closedItems"])) + " closed, " +
                        str(len(comparison_gpm["changedItems"])) + " changed, " +
                        str(comparison_gpm["unchangedCount"]) + " unchanged."
                    )
                except Exception as comp_error:
                    self.log(
                        "Warning: could not generate Global PM comparison (" + str(comp_error) +
                        "). Global PM comparison tab will be hidden."
                    )
                    comparison_gpm = None
                    used_gpm_prev_sheet = None

            comparison_al = None
            used_al_prev_sheet = None
            if include_comparison and use_al:
                al_prev_sheet = self.al_prev_sheet_var.get().strip()
                if al_prev_sheet:
                    try:
                        self.log("Reading Action Logs previous week sheet: " + al_prev_sheet)
                        _, al_previous_clean = load_base_previous_week(
                            self.al_xlsx_path, al_prev_sheet, SOURCE_ACTION_LOG
                        )
                        comparison_al = build_comparison_data(al_clean, al_previous_clean)
                        used_al_prev_sheet = al_prev_sheet
                        self.log(
                            "Action Logs comparison: " + str(len(comparison_al["newItems"])) + " new, " +
                            str(len(comparison_al["closedItems"])) + " closed, " +
                            str(len(comparison_al["changedItems"])) + " changed, " +
                            str(comparison_al["unchangedCount"]) + " unchanged."
                        )
                    except Exception as comp_error:
                        self.log(
                            "Warning: could not generate Action Logs comparison (" + str(comp_error) +
                            "). Action Logs comparison tab will be hidden."
                        )
                        comparison_al = None
                        used_al_prev_sheet = None

            source_files = [self.gpm_xlsx_path.name]
            if use_al and self.al_xlsx_path:
                source_files.append(self.al_xlsx_path.name)
            source_files_label = ", ".join(source_files)

            html_out = render_html(
                combined_payload,
                source_files_label,
                comparison_gpm=comparison_gpm,
                comparison_al=comparison_al,
                current_sheet_gpm=gpm_sheet,
                previous_sheet_gpm=used_gpm_prev_sheet,
                current_sheet_al=(self.al_sheet_var.get().strip() if use_al else None),
                previous_sheet_al=used_al_prev_sheet,
            )
            self.output_path.write_text(html_out, encoding="utf-8")

            self.log("Dashboard generated successfully at:" + NEWLINE + str(self.output_path))
            self.root.after(0, self._on_success)

        except Exception as e:
            error_msg = "Error generating dashboard:" + NEWLINE + str(e) + NEWLINE + NEWLINE + traceback.format_exc()
            self.log(error_msg)
            self.root.after(0, lambda: messagebox.showerror(APP_TITLE, str(e)))
            self.root.after(0, lambda: self.btn_generate.state(["!disabled"]))

    def _on_success(self):
        self.btn_generate.state(["!disabled"])
        self.btn_open.state(["!disabled"])
        self.btn_folder.state(["!disabled"])
        messagebox.showinfo(APP_TITLE, "Dashboard generated successfully!")

    def open_output(self):
        if self.output_path and self.output_path.exists():
            webbrowser.open(self.output_path.as_uri())

    def open_output_folder(self):
        if self.output_path and self.output_path.exists():
            folder = self.output_path.parent
            if sys.platform == "win32":
                os.startfile(folder)
            else:
                webbrowser.open(folder.as_uri())


def main():
    root = Tk()
    DashboardApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()