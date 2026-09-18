# Deploying this system

This is a Google Apps Script project (bound to a Google Sheet), managed as
code via [`clasp`](https://github.com/google/clasp). It was written and
committed in a sandboxed environment with **no Google account access**, so
nothing here has been pushed to, or executed inside, an actual Google Sheet
yet. Follow these steps to stand it up for real.

## 1. Prerequisites

```bash
npm install -g @google/clasp
clasp login
```

`clasp login` opens a browser OAuth flow against your Google account — do
this on your own machine, not in an unattended environment.

## 2. Create (or bind to) the target Google Sheet

Option A — brand new sheet, script created alongside it:

```bash
clasp create --title "SRM AP Strategy Execution System" --type sheets --rootDir ./src
```

Option B — bind to an existing Google Sheet (e.g. a copy of the real
SRMAP_KRA-KPI_AY26-27 workbook, so the KPI system lives right next to the
source data):

1. Open the target Sheet in a browser → Extensions → Apps Script.
2. Project Settings → copy the **Script ID**.
3. In this repo, create `.clasp.json` (see `.clasp.json.example`) with that
   Script ID and `"rootDir": "src"`.

Either way, `clasp` needs `rootDir: src` — that directory holds every `.gs`
file, `appsscript.json` (the manifest), and `QuarterlyEntryForm.html`, in the
flat layout Apps Script projects expect.

## 3. Push the code

```bash
clasp push
```

This uploads every file under `src/` to the Apps Script project.

## 4. First run, inside the Sheet

1. Open the bound Google Sheet. Reload it — the **"SRM AP Strategy System"**
   custom menu should appear (wired by `onOpen()` in `src/07_Menu.gs`).
2. **Menu → 1. Setup Schema.** Creates all canonical tabs (Pillars, Goals,
   Clusters, Offices, Verticals, KRAs, KPIs, KPI_Office_Links, Actions,
   Audit_Log, Validation_Log, Import_Log, Escalations, Notification_Log,
   Office_Contacts, Config) with headers, dropdowns and the default
   thresholds.
3. On first use of any menu item, Google will prompt for authorization
   (Spreadsheets, sending mail, UI) — approve it.
4. **Smoke test without the real workbook (optional):**
   - Menu → 2a. Build Sample Source Data — creates placeholder
     `SAMPLE_SRC_*` tabs shaped like the real workbook.
   - Menu → 2b. Import Sample Source Data → Canonical Model — runs the full
     import pipeline against that fixture. Confirm rows land correctly, then
     Menu → 2c to remove the fixture tabs.
5. **Import the real workbook:**
   - Either paste its actual tabs (`01_Index`, `03_Strategic_Pillars`,
     `04_Cascade_Map`, the five Tier 0/1 scorecards, and the 30 office
     scorecards) directly into this spreadsheet under their real names, or
     keep the real workbook as a separate Google Sheet and use its Sheet ID.
   - Menu → 3. Import Real Workbook..., paste the Sheet ID/URL (or leave
     blank to import from tabs already in this spreadsheet).
   - **Validate against the source workbook numbers before trusting the
     import**: compare `Import_Log` row counts against `01_Index`'s own KPI
     counts per office. `tools/validate_import.py` already did this offline
     against the real file (see README) and found an exact match on every
     office/scorecard -- if a live import in Sheets disagrees, something
     about the live environment differs and is worth tracking down before
     going further.
6. Menu → 4. Validate Weights — confirms every office/cluster sums to 100
   and every goal with contributing offices has an Owner Weight % set.
   Fix any `FAIL` rows in `Validation_Log` before treating a scorecard as
   published.
7. Menu → 7. Recompute Roll-up Achievement % — computes KPI → Office →
   Cluster → Goal → Pillar achievement (see `08_RollupEngine.gs` and
   `NEXT_STEPS.md` for the formulas and their caveats). Re-run any time
   after entering new Actuals.
8. Menu → 8a. Refresh Dashboards — builds/refreshes the VC, Cluster, Office,
   Goal and My KPIs dashboard tabs. Office/Goal/My KPIs use a picker cell
   (e.g. type an office code into `Dashboard_Office!B3`) and update live via
   `QUERY()` formulas after that.
9. Menu → 6. Open Quarterly Entry Sidebar — this is the screen KPI Owners
   use every quarter to log Actuals; it auto-computes RAG status and blocks
   saving an Amber/Red quarter without an ATR note.
10. Fill in the `Office_Contacts` sheet with real email addresses, then
    Menu → 10b. Install Daily Escalation/Reminder Triggers to automate
    `runEscalationScan()` and `notifyKpiOwners()` (both also run on demand
    from the menu). Without contacts filled in, reminders log
    "SKIPPED - no contact email" instead of sending.
11. Menu → 11a-d for the spec's reports: Quarterly Scorecard Report (one
    office), Cluster/VC Review Pack, Owner Performance Report, and
    Regenerate All Scorecard Tabs (rebuilds one tab per Tier 0/1/2
    scorecard from canonical data -- do this before File > Download >
    Microsoft Excel for a drop-in export).

## 5. Ongoing changes

Edit files locally, then `clasp push` again. `clasp pull` will pull down any
changes made directly in the Apps Script IDE (avoid doing that long-term —
keep this repo the source of truth).

## What's rough / needs real people to check it

All seven build phases have a working implementation now. See
`NEXT_STEPS.md` for exactly which parts (mainly the roll-up engine's
achievement-% formulas, and escalation/notification's due-date and contact
assumptions) are first-pass approximations that need validation against how
the university actually wants these numbers to behave, before this goes
into real use.
