# Data model

Canonical "system of record" sheets, built by `setupSchema()`
(`src/01_SchemaSetup.gs`) and populated by the import pipeline
(`src/03_ImportPipeline.gs`). Column lists are defined once, in
`src/00_Constants.gs` (`COLS`), and every module reads/writes by header
name — never by column index — so this layout can evolve safely.

This mirrors spec section 3 (entity model) and section 2 (reference data to
import, not recreate).

## Pillars
`Pillar ID` (P1–P6), `Pillar Name`.

## Goals
`Goal ID` (e.g. P1.1), `Parent Pillar`, `Five-Year Goal / Target`,
`Owner Office`, `Contributing Offices`, `Owner Weight %`, `Overall Progress`.

`Owner Weight %` is the admin-set owner/contributor roll-up split described
in spec section 4 ("the system should let an admin set that split explicitly
per goal rather than hard-coding an even split") — it's intentionally left
blank on import and flagged by `Validate Weights` until an admin sets it.
`Overall Progress` is a placeholder column for the roll-up engine (phase 2 —
see `NEXT_STEPS.md`); it is not computed yet.

## Clusters
`Cluster ID`, `Cluster Name`, `Member Offices`, `Aggregate Weight (Check)`.

## Offices
`Office Code`, `Full Name`, `Cluster`, `Tier`, `Goals Owned`,
`Goals Contributed To`, `KPI Count (Strategic)`, `KPI Count (Operational)`,
`Vertical Count`, `Total Weight (Check)`.

Populated from `04_Cascade_Map` (goals owned/contributed, KPI split) and
recomputed (`Total Weight (Check)`) from the actual KPI rows after every
import, per office.

## Verticals / KRAs
Derived automatically from the KPI rows' `Vertical` / `KRA` text columns
during import (`deriveVerticalsAndKras_()`) — unique `(Office, Vertical)` and
`(Vertical, KRA)` combinations become rows with generated IDs (`V001`,
`K001`, ...). Not hand-maintained.

## KPIs
The full record per spec section 3.7, one row per KPI per scorecard
appearance. All 49 spec fields are present, plus two system fields:

- `KPI_ID` — generated as `<OfficeCode>-<seq>` on import; the join key every
  other table uses.
- `Co-Owner Offices (Name:%)` — optional source column (`"AiTI:30, QuTI:20"`)
  that the importer expands into `KPI_Office_Links` rows instead of
  duplicating the KPI (spec section 4 / 12, the AiTI/QuTI/Research
  co-ownership case).

`Rolls Up From` is populated on Tier 0/1 rows (the literal parent-child link
described in spec section 2) and is what the future roll-up engine will key
off instead of assuming an even split across contributing offices.

Status (`Q1–Q4 Status`) is intentionally left blank by the importer, exactly
as the source workbook ships it — filling it in is the system's core job
every quarter (spec section 6), via the Quarterly Entry sidebar or the bulk
RAG recompute menu item.

## KPI_Office_Links
`KPI_ID`, `Office Code`, `Role` (`Owner` / `Contributor`), `Co-Ownership %`.

A genuine many-to-many join table so a shared KPI (e.g. AiTI/QuTI co-owning
P3.1 with Research) is represented once, not duplicated — required by build
note #12 ("Build the co-ownership case ... as a genuine many-to-many
KPI-office link from the start").

## Actions
The milestone/action execution layer from spec section 3.8, one level below
KPI: `Action ID`, `Description`, `Linked KPI_ID`, `Owner`, `Contributors`,
`Priority`, `Start Date`, `Due Date`, `Dependencies`, `Status`,
`% Complete`, `Closure Criteria`, `Remarks`. Schema only in this pass — no
UI yet (see `NEXT_STEPS.md`).

## Audit_Log
`Timestamp`, `User`, `Sheet`, `Row`, `Field`, `Old Value`, `New Value`.
Written by every quarterly-entry save (`06_QuarterlyEntry.gs`) and by direct
in-sheet edits to targets/weights/owners/status via the `onEdit` trigger
(`99_Triggers.gs`), per spec section 7 ("Maintain a full audit trail of
changes to targets, weights, owners and status").

## Validation_Log
`Timestamp`, `Scope` (`Office`/`Cluster`/`Goal`), `Entity ID`,
`Total Weight`, `Status` (`OK`/`FAIL`), `Details`. Rewritten on every
`Validate Weights` run; failing rows are highlighted.

## Import_Log
`Timestamp`, `Source Sheet`, `Tier`, `Cluster`, `Rows Imported`, `Status`,
`Details`. One row per source tab processed by the import pipeline — this is
what you check row counts against to validate an import (spec build note:
"validate it against the workbook numbers ... before moving on").

## Config
Key/value store for tunables: `Current_AY`, `Green_Threshold` (0.95),
`Amber_Threshold` (0.80), `Weight_Tolerance` (0.01). Read by
`05_RAGEngine.gs` and `04_Validation.gs` — change values here rather than in
code.
