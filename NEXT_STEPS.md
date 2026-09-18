# What's built vs. what's next

This pass covers **phases 1–2** of the build plan (data model + import
pipeline, core CRUD/quarterly entry flow). Phases 3–7 are not built yet.

## Built in this pass

- Canonical data model for Pillars → Goals → Clusters → Offices → Verticals →
  KRAs → KPIs, with the full KPI field set (spec section 3.7), the
  Actions/milestone layer schema (3.8), and a genuine many-to-many
  KPI↔Office co-ownership table (section 4/12).
- Import pipeline that reads `01_Index`, `03_Strategic_Pillars`,
  `04_Cascade_Map`, the five Tier 0/1 scorecards, and every Tier 2 office
  sheet `01_Index` lists — by header name, tolerant of the known
  blank-header-row quirk.
- A placeholder sample-data fixture (`02_SampleData.gs`) so the importer is
  testable without the real workbook. **This is not real data** — see
  `README.md`.
- Office/Cluster weight-sum-to-100 validation, and Goal-level owner-weight
  presence validation, logged to `Validation_Log`.
- RAG status computation (Green ≥95% / Amber 80–94% / Red <80% / Grey =
  no actual) with configurable thresholds and a manual-entry fallback for
  Band/Grade-unit KPIs.
- Quarterly Entry sidebar for KPI Owners: pick office → KPI → quarter, enter
  Actual (or manual Status for non-numeric units), server enforces "ATR
  required when Amber/Red" before saving.
- Audit logging of edits to targets/weights/owners/status, both from the
  entry form and from direct in-sheet edits (`onEdit` trigger).

## Not built yet (phases 3–7 of the plan)

1. **Roll-up engine** (phase 4): KPI → KRA → Vertical → Office → Cluster →
   Goal → Pillar achievement computation, using `Rolls Up From` links and the
   Goal `Owner Weight %` split rather than an even split. This is explicitly
   called out as "the part most likely to need correcting once real people
   check the numbers" — build it as an isolated, unit-testable calculation
   layer before wiring it into any dashboard.
2. **Dashboards & drill-down** (phase 5): VC/Institutional, Cluster, Office,
   Goal, and individual-accountability views, plus the Pillar → Goal →
   Cluster → Office → Vertical → KRA → KPI → Quarterly Milestone → Actual →
   ATR click-through. Depends on (1).
3. **Escalation & notifications** (phase 6): auto-flag Red KPIs, ATRs open
   past quarter-close, missed milestones, weight totals drifting from 100
   (the `onEdit` toast in `99_Triggers.gs` is a start, but not a real
   notification/escalation system); per-KPI `Frequency`-aware reminders to
   Owners/Responsible Officers/Informed parties.
4. **Reporting & Excel export** (spec section 11): Quarterly Scorecard
   Report, Cluster/VC Review Pack, Owner Performance Report, and export back
   to the original 38-tab structure for continuity with existing
   stakeholders.
5. **Actions/milestone UI**: the `Actions` sheet schema exists (3.8) but has
   no entry screen yet.
6. **Real-workbook validation**: re-run the import against the actual
   SRMAP_KRA-KPI_AY26-27 workbook once available, and reconcile
   `Import_Log` row counts against `01_Index`'s own KPI-count column per
   office — this was the explicit gate before moving past phase 2.

## Suggested order for the next session

Roll-up engine first (it's pure calculation, testable without any UI), then
dashboards on top of it, then escalation/notifications, then reporting/export
— matching the phase order in the original build plan.
