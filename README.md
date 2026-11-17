# Action Tracker Dashboard Generator

A desktop tool that turns one or two Excel action-tracking exports into a single
interactive, fully offline HTML dashboard — no server, no internet connection,
no data leaving your computer.

Point it at a weekly "open items" export (and optionally a second source with a
different schema) and it produces a self-contained `.html` file with KPI cards,
charts, a sortable/filterable table, and a week-over-week "what changed" view.

## Why this exists

Tracking open action items / risks / escalations across a project in raw Excel
gets unreadable fast: no severity grouping, no trend, no easy way to see what
changed since last week. This tool reads the export(s) you already produce,
applies a consistent escalation/severity model, and renders a dashboard you can
open in any browser or drop into an email/SharePoint link — all without
standing up any infrastructure.

## Features

- **Two source schemas supported out of the box** — a "Global PM" style export
  and a "Project Action Log" style export, each with its own column mapping.
  Combine both into one dashboard, or use just one.
- **Automatic escalation classification** (L1 / L2 / L3) and **alert severity**
  (`critical` / `urgent` / `attention` / `alert` / `n/a`) computed from days-open
  and whether an action plan / comment has been filled in:
  - **Critical** — an L3 item open more than 30 days, regardless of plan status.
  - **Urgent** — an L2 item open 7+ days with no action plan text yet.
  - **Alert** — an L2 item open more than 30 days that *does* have a plan.
  - **Attention** — an L2/L3 item with a plan, open between 8 and 30 days.
- **Automatic "days open" calculation** when the source doesn't provide it,
  falling back to the creation/due date.
- **Week-over-week comparison** — feed in last week's sheet and this week's,
  and the dashboard gets a "New / Closed / Changed / Unchanged" tab showing
  exactly which fields changed per item.
- **Self-contained output** — the generated HTML embeds all data as inline JSON
  and uses Chart.js from a CDN; everything else (filtering, sorting, tabs,
  row-expand, a synced floating horizontal scrollbar for wide tables) is vanilla
  JS baked into the file.
- **Simple desktop UI** (Tkinter) — pick your file(s) and sheet names, tick a
  couple of checkboxes, generate.

## Tech stack

- Python 3, [pandas](https://pandas.pydata.org/) + [openpyxl](https://openpyxl.readthedocs.io/) for reading Excel
- Tkinter for the desktop UI
- Vanilla JS + [Chart.js](https://www.chartjs.org/) (loaded from CDN) for the generated dashboard

## Getting started

```bash
pip install pandas openpyxl
python dashboard_app.py
```

1. Select your "Global PM"-style `.xlsx` export and its current-week sheet name.
2. (Optional) Enable "Project Action Logs" and select a second `.xlsx` with that schema.
3. (Optional) Enable weekly comparison and fill in the previous week's sheet name(s).
4. Click **Generate HTML Dashboard** and choose where to save it.
5. Open the generated `.html` in any browser.

### Expected columns

The two supported source schemas are defined near the top of `dashboard_app.py`
as `GLOBAL_PM_COLUMN_MAP` and `ACTION_LOG_COLUMN_MAP`. Adjust those dictionaries
to match your own spreadsheet's column headers — the rest of the logic is
schema-agnostic and works off the mapped keys.

### Owner name cleanup

If your "Owner" column exports people-picker values with a trailing company
name (e.g. `"Jane Doe Example Corp"`), set `OWNER_COMPANY_SUFFIX` near the top
of the file to match your organization so names get cleaned up automatically.
Leave it as an empty string to disable this step entirely.

## Packaging as a standalone .exe

A PyInstaller spec (`action_tracker_dashboard.spec`) is included:

```bash
pip install pyinstaller
pyinstaller action_tracker_dashboard.spec
```

## License

MIT — see [LICENSE](LICENSE).
