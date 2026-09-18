# SRM AP Strategy Execution & Performance Management System

Digital system for SRM University AP's Strategic Goal → Pillar → Cluster →
Office → KRA/KPI cascade, modelled directly on the SRMAP_KRA-KPI_AY26-27
workbook: Pillars (P1–P6) → 5-Year Goals → Clusters (Tier 1) → Offices
(Tier 2) → Verticals → KRAs → KPIs → quarterly milestones/actuals/RAG
status/ATR, with the five-role RACI model (Owner, R, A, C, I, Responsible
Officer) and weight-sum-to-100 validation preserved throughout.

**Stack:** Google Sheets + Google Apps Script (chosen to stay closest to the
existing workbook-based workflow stakeholders already use).

**Status:** all seven build phases have a working implementation --
data model, import pipeline, weight validation, RAG computation, quarterly
entry, the roll-up engine, dashboards/drill-down, escalation/notifications,
and reporting/Excel export. The import pipeline has been validated against
the real `SRMAP_KRA-KPI_AY26-27_Updated.xlsx` (39 sheets, 973 KPI rows across
30 offices + 5 Tier 0/1 scorecards -- see `tools/validate_import.py` and
"Validated against the real workbook" below). None of this code has been
executed inside an actual Google Sheet yet (this environment has no Google
account access) -- see `DEPLOY.md` and NEXT_STEPS.md's "First real run"
checklist before treating it as production-ready.

## Validated against the real workbook

`tools/validate_import.py` re-implements the same header-normalization /
header-discovery / row-filtering logic `src/03_ImportPipeline.gs` uses, in
Python, and runs it directly against the `.xlsx` (no Google Sheets needed).
Run it yourself after any change to the import logic:

```bash
pip install openpyxl
python3 tools/validate_import.py /path/to/SRMAP_KRA-KPI_AY26-27_Updated.xlsx
```

Last run against the real file: all 30 Tier 2 office KPI row counts match
`01_Index` exactly, all 5 Tier 0/1 scorecards match, every office's/scorecard's
Wt(%) sums to 100, 0 Goal ID extraction failures, 44 KPI rows carry
`Goal Role = Co-owner` (captured as many-to-many `KPI_Office_Links`, e.g.
AiTI/QuTI/Research co-owning P3.1). Two real-workbook quirks the importer
handles that are worth knowing about:
- Header row position varies by sheet (row 2, 3 or 4) -- located by content,
  not a fixed row number.
- Every Tier 0/1/2 sheet carries a trailing "TOTAL" row (and one sheet,
  Admissions, has a few stray leftover cells below its data) -- filtered out
  by requiring a real `KPI` value on each row.

`02_SampleData.gs`'s `SAMPLE_SRC_*` fixture (placeholder data, built before
the real file was available) is kept only as a fast smoke test independent
of the large real workbook; use "3. Import Real Workbook..." for real data.

## Layout

```
src/                        Apps Script source (flat, matches Apps Script's project layout)
  appsscript.json            Manifest (scopes, runtime)
  00_Constants.gs             Sheet names + column headers + header aliases (schema source of truth)
  01_SchemaSetup.gs           Builds the canonical sheets
  02_SampleData.gs            PLACEHOLDER sample source data -- smoke-test fixture only
  03_ImportPipeline.gs        Reads the 39-sheet workbook by header name into the canonical model
  04_Validation.gs            Office/Cluster/Goal weight-sum-to-100 checks
  05_RAGEngine.gs             Green/Amber/Red/Grey status computation
  06_QuarterlyEntry.gs        Server side of the KPI Owner quarterly entry sidebar
  07_Menu.gs                  Custom spreadsheet menu wiring everything below
  08_RollupEngine.gs          KPI -> Office -> Cluster -> Goal -> Pillar achievement roll-up
  09_Dashboards.gs            VC / Cluster / Office / Goal / My KPIs dashboards
  10_Escalation.gs            Red/overdue-ATR/missed-milestone/weight-drift escalation + reminders
  11_Reporting.gs             Quarterly Scorecard / Review Pack / Owner reports + Excel-export tabs
  99_Triggers.gs               onEdit audit logging + live weight-total checks + Red-status alerts
  QuarterlyEntryForm.html     Quarterly entry sidebar UI
tools/
  validate_import.py          Standalone import validator (see above) -- run against the real .xlsx
DATA_MODEL.md                 What every canonical sheet/column means
DEPLOY.md                     How to push this to a real Google Sheet with clasp, and first-run steps
NEXT_STEPS.md                 What's still rough / needs real people to check the numbers
```

## Quick start

See `DEPLOY.md` for full steps. Summary: `clasp push` this into a Google
Sheet, then in that Sheet use the **"SRM AP Strategy System"** menu, roughly
top to bottom: Setup Schema → Import Real Workbook → Validate Weights →
Recompute Roll-up Achievement % → Refresh Dashboards → Open Quarterly Entry
Sidebar (this is the one KPI Owners use every quarter).
