# What's built, and what still needs real eyes on it

All seven phases of the build plan now have a working implementation. This
document is the honest list of what's genuinely solid vs. what's a
first-pass approximation that needs to be checked against how the
university actually wants these numbers to work.

## Built and validated against real data

- **Data model & import** (phases 1-2): canonical Pillars/Goals/Clusters/
  Offices/Verticals/KRAs/KPIs, full spec §3.7 KPI field set, Actions/
  milestone schema, many-to-many KPI↔Office co-ownership. Import pipeline
  reads the real 39-sheet workbook by header name (normalized, alias-mapped)
  and has been checked row-for-row against `01_Index` -- see README's
  "Validated against the real workbook" section and `tools/validate_import.py`.
- **CRUD / quarterly entry** (phase 3): weight validation, RAG computation,
  Quarterly Entry sidebar with ATR-required-on-Amber/Red, audit logging.

## Built, but first-pass and likely to need correction

These are exactly the parts the original build plan flagged as needing
correction once real people check the numbers ("the part most likely to
need correcting"). Each is isolated and documented so it's fixable without
touching the rest of the system.

- **Roll-up engine** (`08_RollupEngine.gs`, phase 4). The real workbook's
  quarterly targets are frequently narrative text, not numbers, even on
  numeric-unit KPIs -- so per-KPI achievement % falls back through
  `Score (Actual vs Target)` → numeric Q4 Actual/Target → RAG-status score
  (Green=100/Amber=70/Red=40/Grey=0), in that order. Cluster roll-up
  averages the named "Rolls Up From" offices equally (no finer split is in
  the source data); Goal roll-up defaults to an even owner/contributor split
  until an admin sets `Owner Weight %` on the Goals sheet. **Before trusting
  any Achievement % number, have someone who owns a KPI check it against
  what they'd expect.**
- **Escalation & notifications** (`10_Escalation.gs`, phase 6). Due-date
  reminders key off four Config-configurable quarter-close dates for every
  KPI, because the workbook only ever carries four milestone checkpoints
  regardless of a KPI's own Frequency (a "Monthly" or "Per exam cycle" KPI
  still only has Q1-Q4 slots to be measured against) -- a true per-Frequency
  calendar would need the workbook itself to carry more checkpoints.
  Notifications need real email addresses: fill in the `Office_Contacts`
  sheet (Office Code → email) -- until then, `notifyKpiOwners()` logs
  "SKIPPED - no contact email" per KPI instead of silently doing nothing.
- **Dashboards** (`09_Dashboards.gs`, phase 5): Office/Goal/My KPIs use live
  `QUERY()` formulas; VC/Cluster are script-computed snapshots refreshed by
  "Refresh Dashboards". Multi-dimension filtering (KRA Type + Metric Type +
  Indicator Nature + Status all at once) is left to Sheets' own native
  filter views (Data > Create a filter) over the QUERY output rather than a
  custom filter UI -- deliberate, not a shortcut, but worth knowing.
- **Reporting/export** (`11_Reporting.gs`, phase 7): `regenerateAllScorecardTabs()`
  rebuilds one tab per Tier 0/1/2 scorecard from canonical data using the
  canonical column set/order, not a byte-for-byte reproduction of the
  original workbook's exact header wording and formatting. After
  regenerating, File > Download > Microsoft Excel is the actual export step
  (native Sheets capability, no custom code needed for it).

## Not built at all

- **Actions/milestone UI**: the `Actions` sheet schema exists (§3.8) but has
  no entry screen. Reports reference it (e.g. Owner Performance Report's
  "open actions") but it'll be empty until KPI Owners have somewhere to log
  actions against a KPI.
- **Per-KPI configurable RAG thresholds**: thresholds are global (Config
  sheet), not per-KPI as the spec's stretch goal describes.
- **Fine-grained per-recipient RACI notification routing**: `Office_Contacts`
  is one email per role per OFFICE, not per named individual in the R/A/C/I
  columns -- those columns are free text (e.g. "Vice-Chancellor; Pro
  Vice-Chancellor; Governing Council") and aren't parsed into addressable
  recipients.

## First real run checklist

Nothing in this repo has executed inside an actual Google Sheet yet (this
build environment has no Google account access -- see `DEPLOY.md`). Before
treating any of it as production-ready:

1. `clasp push` into a real Google Sheet, run "1. Setup Schema", then
   "3. Import Real Workbook..." against the actual file.
2. Compare `Import_Log` row counts against what `tools/validate_import.py`
   predicted (they should match exactly -- if not, something about the live
   Sheets environment is behaving differently than openpyxl did).
3. Run "4. Validate Weights" and fix any FAIL rows before going further.
4. Run "7. Recompute Roll-up Achievement %" and have a handful of KPI
   Owners sanity-check their own office's number.
5. Pilot with one office + one cluster for a live quarter (per the original
   build plan's step 7) before rolling out to all 30 offices.
