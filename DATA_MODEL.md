# Data model

Canonical "system of record" sheets, built by `setupSchema()`
(`src/01_SchemaSetup.gs`) and populated by the import pipeline
(`src/03_ImportPipeline.gs`). Column lists are defined once, in
`src/00_Constants.gs` (`COLS`), and every module reads/writes by header
name — never by column index — so this layout can evolve safely.

This mirrors spec section 3 (entity model) and section 2 (reference data to
import, not recreate). Reconciled against the real
`SRMAP_KRA-KPI_AY26-27_Updated.xlsx` (39 sheets) -- see README's "Validated
against the real workbook" section.

## Pillars
`Pillar ID` (P1–P6), `Pillar Name`, `Achievement %`.

`Pillar`/`Pillar name` are only filled on the first goal row of each pillar
group in the real `03_Strategic_Pillars` sheet (a visual merge); the
importer forward-fills them. `Achievement %` is written by the roll-up
engine (`08_RollupEngine.gs`) -- a simple average of the pillar's goals'
achievement.

## Goals
`Goal ID` (e.g. P1.1), `Parent Pillar`, `Five-Year Goal / Target`,
`Owner Office`, `Contributing Offices`, `Owner Weight %`, `Overall Progress`,
`Achievement %`.

`Owner Office` is sometimes compound in the real data (e.g. `"SEAS / PSB"`
on P1.4, `"Research / DEEPS"` on P3.5) -- the roll-up engine splits on `/`
and treats each as a co-owner for weighting purposes. `Owner Weight %` is
the admin-set owner/contributor roll-up split described in spec section 4
("the system should let an admin set that split explicitly per goal rather
than hard-coding an even split") — it's intentionally left blank on import;
until it's set, the roll-up engine defaults to an even split across owner(s)
and contributor(s). `Achievement %` is written by the roll-up engine.

## Clusters
`Cluster ID`, `Cluster Name`, `Member Offices`, `Aggregate Weight (Check)`,
`Achievement %`. Derived from the Offices sheet's `Cluster` column after
import (`deriveClusters_()`), not hand-maintained.

## Offices
`Office Code`, `Full Name`, `Cluster`, `Tier`, `Goals Owned`,
`Goals Contributed To`, `KPI Count (Strategic)`, `KPI Count (Operational)`,
`Vertical Count`, `Total Weight (Check)`, `Achievement %`.

`Cluster`/`Goals Owned`/`Goals Contributed To`/KPI-count columns come from
`04_Cascade_Map`; `Total Weight (Check)` and `Achievement %` are recomputed
from the actual Tier 2 KPI rows after every import / roll-up run.

## Verticals / KRAs
Derived automatically from the KPI rows' `Vertical` / `KRA` text columns
during import (`deriveVerticalsAndKras_()`) — unique `(Office, Vertical)` and
`(Vertical, KRA)` combinations become rows with generated IDs (`V001`,
`K001`, ...). Not hand-maintained.

## KPIs
The full record per spec section 3.7, one row per KPI per scorecard
appearance (973 rows on the real AY26-27 workbook: 72 Tier 0/1 + 901 Tier 2).
All spec fields are present, plus system fields:

- `KPI_ID` — generated as `<OfficeCode-or-TierCode>-<seq>` on import (e.g.
  `AA-001`, `VC-001`, `C1-001`); the join key every other table uses.
- `Goal ID` — EXTRACTED, not read directly: the real source column is a
  single compound field (`5-Yr Goal (ID + target)`, e.g. `"P1.4 —
  International Accreditation: AACSB and ABET"`); `extractGoalId_()` pulls
  the leading `P<n>.<n>` out of it via regex.
- `Achievement %` — written by the roll-up engine, not imported.
- `Rolls Up From` — populated on Tier 0/1 rows only (Tier 2 sheets don't
  have this column); the literal parent-child link the roll-up engine keys
  off for cluster-level achievement, per spec section 2/4.
- `Owner Sign-off` — present on Tier 0/1 rows only in the real workbook;
  carried through as-is, not otherwise used yet.

**Co-ownership** (spec section 4/12, e.g. AiTI/QuTI/Research co-owning
P3.1): the real workbook has NO per-KPI "co-owner %" column. Instead, each
co-owning office has its OWN KPI row tagged `Goal Role = "Co-owner"` for
that goal. The importer turns every such row into a `KPI_Office_Links` row
automatically — a genuine many-to-many link, never a duplicated KPI.

`Goal Role` values found in the real data: `Owner`, `Contributor`,
`Institutional` (Tier 0), `Cluster aggregate` (Tier 1), `Co-owner`, and
`Enabler` (an office's standard "University Impact" operational-hygiene KRA,
supporting a goal it neither owns nor is a named contributor to).

Status (`Q1–Q4 Status`) is intentionally left blank by the importer, exactly
as the source workbook ships it — filling it in is the system's core job
every quarter (spec section 6), via the Quarterly Entry sidebar or the bulk
RAG recompute menu item.

## KPI_Office_Links
`KPI_ID`, `Goal ID`, `Office Code`, `Role` (`Co-owner` today; the column
supports any role string), `Co-Ownership %` (left blank -- the real data
gives no percentage split, only the fact of co-ownership; an admin can fill
this in).

A genuine many-to-many join table so a shared goal (e.g. AiTI/QuTI/Research
all owning P3.1) is represented via links, not by duplicating KPI rows or
guessing a percentage — required by build note #12. The roll-up engine
(`08_RollupEngine.gs`) reads this table to extend a Goal's contributor list
beyond `03_Strategic_Pillars`' own `Contributing Offices` text.

## Actions
The milestone/action execution layer from spec section 3.8, one level below
KPI: `Action ID`, `Description`, `Linked KPI_ID`, `Owner`, `Contributors`,
`Priority`, `Start Date`, `Due Date`, `Dependencies`, `Status`,
`% Complete`, `Closure Criteria`, `Remarks`. Schema only — no entry UI yet
(see `NEXT_STEPS.md`).

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
"validate it against the workbook numbers ... before moving on"; see also
`tools/validate_import.py`).

## Escalations
`Timestamp`, `KPI_ID`, `Office`, `Reason` (`Red status` / `ATR overdue` /
`Milestone missed` / `Weight drift`), `Detail`, `Status` (`Open` by default;
edit to anything else to mark resolved and stop it re-triggering). Written
by `runEscalationScan()` (`10_Escalation.gs`), spec section 10.

## Notification_Log
`Timestamp`, `KPI_ID`, `Recipient Role`, `Recipient`, `Reason`, `Sent`
(`Yes`/`No`). Written by `notifyKpiOwners()` and `notifyOnStatusChange_()` —
every attempt is logged, including skips when no contact email is on file,
so gaps in `Office_Contacts` are visible rather than silent.

## Office_Contacts
`Office Code`, `KPI Owner Email`, `Responsible Officer Email`,
`Escalation Email`. Starts empty — the real workbook has no email
addresses, only role/name text (e.g. `"Dean AA"`) in the `KPI Owner` /
`Responsible Officer` columns. Fill this in to make notifications actually
send; see `NEXT_STEPS.md`.

## Config
Key/value store for tunables: `Current_AY`, `Green_Threshold` (0.95),
`Amber_Threshold` (0.80), `Weight_Tolerance` (0.01), `Q1-4_Close_Date`
(AY26-27 quarter-close defaults, used by escalation/reminder logic),
`Reminder_Lead_Days` (7). Read across `05_RAGEngine.gs`, `04_Validation.gs`
and `10_Escalation.gs` — change values here rather than in code.

## Dashboard_* sheets
Not part of the data model itself (no data is stored there that doesn't
already exist elsewhere) — see `09_Dashboards.gs` and README's Layout
section. `Dashboard_VC` / `Dashboard_Cluster` are script-computed snapshots;
`Dashboard_Office` / `Dashboard_Goal` / `Dashboard_MyKPIs` are live
`QUERY()`-formula sheets driven by a picker cell.
