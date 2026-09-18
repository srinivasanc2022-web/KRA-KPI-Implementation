# SRM AP Strategy Execution & Performance Management System

Digital system for SRM University AP's Strategic Goal → Pillar → Cluster →
Office → KRA/KPI cascade, modelled directly on the SRMAP_KRA-KPI_AY26-27
workbook: Pillars (P1–P6) → 5-Year Goals → Clusters (Tier 1) → Offices
(Tier 2) → Verticals → KRAs → KPIs → quarterly milestones/actuals/RAG
status/ATR, with the five-role RACI model (Owner, R, A, C, I, Responsible
Officer) and weight-sum-to-100 validation preserved throughout.

**Stack:** Google Sheets + Google Apps Script (chosen to stay closest to the
existing workbook-based workflow stakeholders already use).

**Status:** foundation phase — data model, import pipeline, weight
validation, RAG computation, and the quarterly entry flow are built. Roll-up
computation, dashboards, escalation/notifications, and Excel-shaped
reporting are not yet built. See `NEXT_STEPS.md`.

## ⚠️ No real workbook data included

The actual `SRMAP_KRA-KPI_AY26-27.xlsx` was not available when this was
built. `src/02_SampleData.gs` seeds small **placeholder** source sheets
(prefixed `SAMPLE_SRC_`) shaped like the real workbook, purely so the import
pipeline can be exercised end to end. Do not treat any pillar/office/KPI
text in that file as real — replace it via **Menu → 3. Import Real
Workbook...** once the actual file is available, and validate the import
against `01_Index`'s own KPI counts per office before trusting it.

## Layout

```
src/                     Apps Script source (flat, matches Apps Script's project layout)
  appsscript.json        Manifest (scopes, runtime)
  00_Constants.gs         Sheet names + column headers (single source of truth for the schema)
  01_SchemaSetup.gs       Builds the canonical sheets
  02_SampleData.gs        PLACEHOLDER sample source data (see warning above)
  03_ImportPipeline.gs    Reads the 38-sheet workbook by header name into the canonical model
  04_Validation.gs        Office/Cluster/Goal weight-sum-to-100 checks
  05_RAGEngine.gs         Green/Amber/Red/Grey status computation
  06_QuarterlyEntry.gs    Server side of the KPI Owner quarterly entry sidebar
  07_Menu.gs              Custom spreadsheet menu wiring everything above
  99_Triggers.gs          onEdit audit logging + live weight-total checks
  QuarterlyEntryForm.html Quarterly entry sidebar UI
DATA_MODEL.md             What every canonical sheet/column means
DEPLOY.md                 How to push this to a real Google Sheet with clasp
NEXT_STEPS.md             What's built vs. what's left (phases 3-7)
```

## Quick start

See `DEPLOY.md` for full steps. Summary: `clasp push` this into a Google
Sheet, then in that Sheet use the **"SRM AP Strategy System"** menu:
Setup Schema → (test with sample data, or) Import Real Workbook →
Validate Weights → Open Quarterly Entry Sidebar.
