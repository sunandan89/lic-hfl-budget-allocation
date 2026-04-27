# LIC HFL — Developer Handover

**Author:** Sunandan
**Date:** 27 April 2026
**Target instance:** `stg.lichfl.mgrant.in` (Frappe v16, mGrant app)
**GitHub repo:** `https://github.com/sunandan89/lic-hfl-budget-allocation` (branch: `main`)
**Sample GAF used for end-to-end validation:** `GAF-0545` → `Project P-GAF-0545` → `Year 1 / FY-2025`

> **Read this entire document before pushing anything from the GitHub repo to staging.** Several of the changes below live ONLY on staging (Custom Fields, Property Setters, Workflow vocabulary widening). If you sync the repo without recreating these on the target environment first, the application will break — the Project Proposal workflow will reject every transition, the budget grids will render with empty convergence columns, and the QUR submit will throw "Invalid workflow action".
>
> ### Repo sync status — 27 April 2026
>
> All actively-deployed code is now on branch `chore/sync-staging-2026-04-27` of the repo, awaiting merge to `main`. Open PR at:
> `https://github.com/sunandan89/lic-hfl-budget-allocation/pull/new/chore/sync-staging-2026-04-27`
>
> Verified by reading the branch contents via the GitHub API:
>
> - `budget_allocation_client_script.js` is now **184,149 B** (was 101 KB / v8c cut on `main`) — current v9 final.
> - `lic_hfl_qur_v2.js` is now in the repo at top level (51,677 B, v2.3).
> - `legacy/budget_utilisation_v3.js` archived for fallback.
> - `HANDOVER_DEVELOPER.md`, `QUR_user_guide.md`, `GIT_PUSH_PLAN.md` added.
> - `README.md` rewritten as a real entry-point.
>
> **Once this branch is merged, a developer can `git clone` and reproduce the deployed JS.** Fixtures (Custom Fields + Property Setters) are still a separate workstream — see Section D.
>
> The document is written in **two layers**:
> 1. *Things that MUST be carried over* (Section A → Section E).
> 2. *Bypasses, hardcodes, and one-shot patches that MUST NOT be carried over* (Section F).

---

## A. Scope — what I built

Two end-to-end surfaces on the Project Proposal (GAF) → Project → Grant chain:

1. **Budget Allocation** — interactive spreadsheet-style budget grid on the Project Proposal form, replacing the legacy 3-fund-source layout with a "convergence" layout matching the LIC HFL Excel template (Programmatic + Non-Programmatic, lumpsum totals, proportional quarterly allocation, multi-year, Excel export).
2. **Budget Utilisation / Quarterly Utilisation Report (QUR)** — units-based quarterly reporting surface on the Grant form, a YTD summary widget, and the quarter-entry widget that triggers the native Frappe QUR workflow.

Everything is implemented as **Client Scripts + Custom Fields + Property Setters**. No app code was added. Some pieces *should* eventually be Server Scripts / app hooks — flagged in Section E.

---

## B. Inventory of changes on staging

### B.1 Client Scripts (DocType: Project proposal)

| Client Script name | DocType | Repo file (canonical source) | Status |
|---|---|---|---|
| `LIC Budget Allocation v4` | Project proposal | `budget_allocation_v9_final.js` (~184 KB, ~4,900 lines) | **Active** |
| `Activity KPI NGO Filter` | Project proposal | `activity_kpi_ngo_filter_client_script.js` (~19 lines) | **Active — dependency** |

> The *name* `LIC Budget Allocation v4` is legacy. The *content* is v9 (final). Don't rename — the override is keyed by name in places.

### B.2 Client Scripts (DocType: Grant)

| Client Script name | DocType | Repo file | Status |
|---|---|---|---|
| `LIC HFL QUR v2` | Grant | `lic_hfl_qur_v2.js` (~52 KB, ~1,200 lines, v2.3) | **Active** — replaces both the `budget` HTML field render and the `quarterly_utilisation` HTML field render |
| `LIC Activity Utilisation` | Grant | (kept disabled) | **Disabled.** Original v1 BPU-direct-write render. Keep in DB as a fallback only; do NOT enable. |
| `LIC Budget Utilisation v3` | Grant | `budget_utilisation_v3.js` | **Superseded by QUR v2.** Probably delete on the new environment unless the `budget_report_html` field exists separately and you want the simplified single-quarter view. |

### B.3 Custom Fields (added via Custom Field doctype)

These are **not** in the GitHub repo. They MUST be created on the target environment **before** the Client Scripts will work — both the budget allocation and QUR scripts read/write these fields.

| DocType | Fieldname | Fieldtype | Inserted after | Purpose |
|---|---|---|---|---|
| Project Budget Planning | `custom_total_lic_contribution` | Currency | `total_planned_budget` | LIC HFL portion of an activity's lumpsum |
| Project Budget Planning | `custom_govt_contribution` | Currency | `custom_total_lic_contribution` | Government convergence portion |
| Project Budget Planning | `custom_benf_contribution` | Currency | `custom_govt_contribution` | Beneficiary convergence portion |
| Project Budget Planning | `custom_task_details` | Small Text | `custom_benf_contribution` | Long-form task description (replaces overloading `assumption`) |
| Project Budget Planning | `custom_activity` | Link → Activity Master | (anywhere on form) | Stable activity ID reference (string-name fallback still in place) |
| Project Budget Planning | `custom_unit_of_measurement` | Link → Units | near `custom_activity` | UoM denormalisation for fast read in QUR |
| Budget Planning Child (BPU child) | `unit` | Float | `planned_amount` | Planned units per quarter (drives QUR) |
| Budget Planning Child (BPU child) | `unit_cost` | Currency | `unit` | Unit cost per quarter (LIC-ratio-adjusted on convergence rows) |
| Quarterly Utilisation Report Child | `custom_actual_units` | Float | `actuals` | Partner input — units delivered |
| Quarterly Utilisation Report Child | `custom_planned_units` | Float (read-only) | `custom_actual_units` | Pulled forward from BPU planning_table.unit |
| Quarterly Utilisation Report Child | `custom_unit_cost` | Currency (read-only) | `custom_planned_units` | Pulled forward from BPU planning_table.unit_cost |

> **How to export these for the new environment:** add the Custom Field doctype to `fixtures` in `hooks.py` of whichever app you control on staging (or just `fixtures = ["Custom Field"]` and prune the JSON down to only LIC-HFL-prefixed fields). Then `bench export-fixtures` and commit the JSON. See Section D for full fixture plan.

### B.4 Property Setters

These also live ONLY on staging — recreate via fixtures.

| Property Setter name | Target | What it changes | Why |
|---|---|---|---|
| `Project proposal-stage-options` | `Project proposal.stage` (Select options) | **Widened** to include both old vocabulary ("GAF Submitted / GAF Under Review / GAF Approved") AND new vocabulary ("Project Application Started / Submitted / Under Review / Approved / Board Note Signed") | The custom workflow override `frappe_theme.overrides.workflow.custom_apply_workflow` writes the **old** labels; the migration introduced the **new** labels. Without this widening, every workflow transition fails with `ValidationError: Stage cannot be "GAF Submitted"…`. **Do not shrink this back.** |
| `upload_signed_mou` label | `Project proposal.upload_signed_mou` | Label → "Upload Board Note" | Final-approve gate cosmetics |
| `mou_signing_date` label | `Project proposal.mou_signing_date` | Label → "Board Note Signing Date" | Same |
| `mou_verified` label | `Project proposal.mou_verified` | Label → "I have uploaded the final signed Board Note…" | Same |
| `Project proposal.total_planned_budget` allow-on-submit | (attempted Property Setter) | `allow_on_submit = 1` | **Does not work** — mgrant app's "Cannot Update After Submit" hook overrides it. Don't bother re-creating. We use a Server Script to bypass on staging (Section F.4). On prod the right fix is a hook patch in mgrant app code. |

### B.5 Workflow

`Project proposal` workflow on this instance has the per-transition popup-required fields documented in `project_lic_hfl_gaf_approval_chain.md` (originSession). Summary:

| From → To | Action role | Required fields collected via popup |
|---|---|---|
| Pending at RPL → Pending at PL | RPL Approve | `custom_project_code` |
| Pending at PL → GAF Approved | PL Approve | comment only |
| GAF Approved → Approved | RPL Approve | `upload_signed_mou` (Board Note), `custom_upload_mou`, `custom_upload_sanction_letter`, `mou_signing_date`, `mou_verified=1` |

Final Approve also has a **server-side check** (in `frappe_theme/overrides/workflow.py → custom_apply_workflow`) that `Project proposal.total_planned_budget > 0`. The right value to populate is `sum(custom_total_lic_contribution)` across all PBPs on the proposal. **If you haven't migrated the convergence custom fields, this check will fail and block every approval.**

`Quarterly Utilisation Report` workflow is the standard mGrant one: Pending → Pending at PM → Pending at SPM → Pending at PL → Approved. The first transition's condition is `custom_spm` set on the Grant (and routes via PM if both PM+SPM are set). **Prerequisite: `Grant.custom_pm` and `Grant.custom_spm` (Link → SVA User) must be set before any QUR Submit will succeed.**

### B.6 Server Script

| Server Script | Status | Purpose |
|---|---|---|
| `patch_gaf_0545_budget` | Disabled, kept in DB | One-shot reusable patch (RestrictedPython sandbox) for direct DB writes that bypass the "Cannot Update After Submit" hook, BPU bulk creation/update from PBPs, and approval state writes. **See Section F.1 — this is a bypass, do NOT push to prod as-is.** |

### B.7 Master data assumptions baked into the client scripts

| Assumption | Where in code | Risk if not true on target |
|---|---|---|
| Donor `D-0001` exists | `budget_allocation_v9_final.js` lines 1488, 1614 (`donor: 'D-0001'`) | Save fails with "Could not find Donor D-0001" |
| Sub Budget Heads named *exactly* `"Programmatic Costs"`, `"Human Resource Costs"`, `"Administration Costs"`, `"NGO Management Costs"` | Throughout `ab_organize*` and `ab_save*` | Section banners empty / save misroutes |
| Fund source naming includes substrings "LIC"/"HFL", "Gov", "Ben" | `ab_resolveFundSource` | Default to LIC column |
| `Project proposal.cummulative_budget` exists as HTML field inside a Budget Allocation tab section | Form layout | Script silently exits |
| `Grant.budget` (HTML) and `Grant.quarterly_utilisation` (HTML) exist | Form layout | QUR widgets don't render |
| `quaterly_project` child table on Project proposal stores per-quarter rows; `annual_project` stores per-year rows | Throughout | Column generation breaks |
| Geography Details doctype name pattern: `GD-Project proposal-{GAF}` linked back via `document_type='Project proposal'` + `docname='{GAF}'` | `ab_extractStatesString`, `ab_extractDistrictCount` (in budget script) and Excel export header | States/districts blank in Excel summary |

---

## C. What is in the GitHub repo vs what is NOT

### C.1 On branch `chore/sync-staging-2026-04-27` (after the 27 Apr sync push)

Verified via GitHub API on the branch:

| File on GitHub | Size | What it is |
|---|---|---|
| `README.md` | 1,880 B | Repo entry-point — points to `HANDOVER_DEVELOPER.md` |
| `HANDOVER_DEVELOPER.md` | this doc | **Primary developer handover** |
| `HANDOVER_v9.md` | 8,065 B | Earlier (25 Apr) snapshot — kept for context |
| `LIC_HFL_Budget_Allocation_Handover.md` | 13,481 B | v8c-era handover — DocType setup notes still relevant |
| `QUR_user_guide.md` | 9,643 B | Partner-facing user guide |
| `GIT_PUSH_PLAN.md` | 8,313 B | Audit trail of the 27 Apr sync push |
| `activity_kpi_ngo_filter_client_script.js` | 690 B | Deploy as Client Script "Activity KPI NGO Filter" |
| `budget_allocation_client_script.js` | **184,149 B** | Deploy as Client Script "LIC Budget Allocation v4" |
| `lic_hfl_qur_v2.js` | 51,677 B | Deploy as Client Script "LIC HFL QUR v2" |
| `legacy/budget_utilisation_v3.js` | 31,448 B | Archived — superseded by QUR v2 |

### C.1b Things deliberately NOT in the repo

- **Custom Field definitions** — kept as schema, must be exported as fixtures (see Section D).
- **Property Setter definitions** — same, fixtures only.
- **`patch_gaf_0545_budget` Server Script** — DELIBERATELY excluded. It is a one-shot bypass (see Section F.1) and replicating it to prod is exactly what we want to prevent.
- **Diagnostic / iteration drafts** that never reached production. Listed in Section C.1c so the developer knows they're local-only and not load-bearing.

### C.1c Local-only artefacts in `~/Documents/Claude/Projects/LIC HFL Changes/`

These are **not** in the repo and should not be deployed. Listed for completeness so the developer / future you doesn't try to use them:

| Local file | Why it's not pushed |
|---|---|
| `budget_allocation_v7c..v9_activity_fix.js` (9 files) | Earlier iterations of the budget allocation script. Replaced by `v9_final`. Museum only. |
| `budget_utilisation_v1.js`, `v2.js` | Diagnostic drafts (Apr 25–26) used to discover schema. Replaced by `v3` then by `lic_hfl_qur_v2.js`. |
| `lic-budget-wireframe*.html` (3 files) | Static design wireframes. Reference only. |
| `pattern-visuals/` | PNGs documenting Frappe patterns — already published on the patterns repo. |
| `test_export_preview.xlsx`, `qurv2_*.txt`, `.qurv2_chunk*.js`, `.deploy/` | Deployment intermediates / test artefacts. |

### C.1d On staging but no local copy + not in repo

| Item | Status | Why it's not in the repo |
|---|---|---|
| `LIC Activity Utilisation` Client Script (Grant, v1, disabled) | Original Apr-14 activity-based QUR, kept disabled in DB on staging as a rollback safety net | We have no local copy of the script body (predates the local-file workflow). If you want it preserved, fetch the script body from staging via `frappe.client.get` on `Client Script` and commit it. Otherwise leave it on staging only — it's already disabled. |
| `patch_gaf_0545_budget` Server Script | Disabled in staging DB | DELIBERATELY excluded — see Section F.1. |

### C.2 NOT in the repo, lives only on staging — recreate via fixtures or manual setup

- All 11 Custom Fields in **Section B.3**
- All 4 active Property Setters in **Section B.4**
- Workflow definition for `Project proposal` (custom labels and per-transition required fields) — currently relies on `frappe_theme` app code, but the **vocabulary widening** must be carried via Property Setter export
- The disabled Server Script `patch_gaf_0545_budget` (you can skip this on prod — see Section F.1)
- The standing assumption that Donor `D-0001` exists, the four Sub Budget Heads are named exactly as listed, the Geography Details doctype + child structure exist

### C.3 Repo hygiene — recommended commits before next deploy

- Drop the v7/v8/v9_activity_fix/budget_utilisation_v1/v2 archives into a `legacy/` folder so the repo root reads cleanly
- Add a top-level `README.md` with the deployment order from Section G of this doc
- Add a `fixtures/` folder with the JSON exports from Section D

---

## C.2 Frappe Patterns library — load-bearing dependency

The work draws heavily on `https://github.com/sunandan89/mgrant-frappe-patterns` (12 patterns published, verified on 27 Apr 2026). The developer should read at least the patterns marked **load-bearing** below before touching the scripts — these are not optional helpers, they are the *reasons the scripts work at all* on Frappe v15/v16.

| Pattern (folder name) | Load-bearing for | Why |
|---|---|---|
| `frappe-wrapper-guard` | Both Budget Allocation + QUR | Monkey-patch on `$wrapper.html()` so Frappe's form refresh doesn't blow away our custom HTML mid-render. Without it, the budget grid disappears the moment the user changes any other field on the form. |
| `frappe-tab-wrapper-access` | Both | How we reach into the `cummulative_budget` (Project proposal) and `budget` + `quarterly_utilisation` (Grant) HTML fields from a Client Script. |
| `frappe-sequential-save` | Budget Allocation NP save path | Avoids Frappe document-lock conflicts when saving 7+ Non-Programmatic PBPs. Parallel `frappe.client.save` calls hit `TimestampMismatchError` immediately. |
| `client-side-xlsx-export` | Budget Allocation Excel export | The xlsx-js-style + post-process workflow that finally produced an Excel file Microsoft doesn't try to "repair". |
| `jszip-bulk-download` | Excel export image embedding | LIC HFL CSR JPEG embedded via twoCellAnchor drawing XML — the JSZip post-process step. |
| `sticky-table-freeze` | Both | The 5/6 frozen-column layout, `table-layout: fixed` + `max-width`, merged `<th colspan>` with `position: sticky; left: 0`. **Note the `feedback_css_sticky_overflow` finding:** `overflow:hidden` and `position:sticky` cannot coexist on the same `<td>`. The pattern's `Known limitations` section needs updating with this. |
| `frappe-dom-data-scraper` | Budget Utilisation v1/v2 (legacy) | Used for the original DOM-scraping approach before we moved to API-driven data fetch. Not load-bearing on QUR v2. |
| `link-field-parent-filter` | `Activity KPI NGO Filter` script | The whole point of that 690-byte script. |
| `remote-client-script-deploy` | Deployment ergonomics | How I pushed scripts to staging via REST API without bench/SSH. The developer doesn't strictly need this if they have bench access on prod. |
| `shadow-dom-keyboard-fix` | (used in some Custom HTML Block work, not directly here) | Ignore unless deploying the same scripts inside a Workspace CHB. |
| `form-read-view-overlay` | Not used in this work | Skip. |
| `fuzzy-search` | Not used in this work | Skip. |

> The HANDOVER_v9.md still references "7 patterns". That's outdated — there are 12 today. The developer should clone the patterns repo and read the load-bearing six before deploying.

---

## D. Fixtures plan — what to export and how

I deliberately DID NOT make this a separate Frappe app, because we don't own a custom app on this client. But the developer (you) can either:

**Option 1 — Add a fixtures folder to mgrant or frappe_theme app** (whichever is more appropriate to own these client-specific overrides). In that app's `hooks.py`:

```python
fixtures = [
    {
        "dt": "Custom Field",
        "filters": [
            ["name", "in", [
                "Project Budget Planning-custom_total_lic_contribution",
                "Project Budget Planning-custom_govt_contribution",
                "Project Budget Planning-custom_benf_contribution",
                "Project Budget Planning-custom_task_details",
                "Project Budget Planning-custom_activity",
                "Project Budget Planning-custom_unit_of_measurement",
                "Budget Planning Child-unit",
                "Budget Planning Child-unit_cost",
                "Quarterly Utilisation Report Child-custom_actual_units",
                "Quarterly Utilisation Report Child-custom_planned_units",
                "Quarterly Utilisation Report Child-custom_unit_cost"
            ]]
        ]
    },
    {
        "dt": "Property Setter",
        "filters": [
            ["name", "in", [
                "Project proposal-stage-options",
                "Project proposal-upload_signed_mou-label",
                "Project proposal-mou_signing_date-label",
                "Project proposal-mou_verified-label"
            ]]
        ]
    },
    {
        "dt": "Client Script",
        "filters": [
            ["name", "in", [
                "LIC Budget Allocation v4",
                "Activity KPI NGO Filter",
                "LIC HFL QUR v2"
            ]]
        ]
    }
]
```

Then `bench --site stg.lichfl.mgrant.in export-fixtures` will dump JSON into the app's `fixtures/` folder. Commit it. On any other site, `bench --site <target> migrate` (with that app installed) will replay them.

**Option 2 — Manual recreation.** Use the Custom Field, Property Setter, and Client Script lists in the desk, manually create each one. Acceptable for a one-time prod cutover, but error-prone and not idempotent.

> Either way, the fixtures **must land on the target environment BEFORE you push the JS files** — because the JS reads `custom_total_lic_contribution` etc. and assumes they exist.

### What does NOT belong in fixtures

- Workflow definition document (`Workflow` doctype) — currently lives in `frappe_theme` app code. If we needed to override it, we'd export the Workflow doc itself, but the override happens at the Python override level. **Don't fixturise this without team alignment** — it would be an environment-specific override.
- Donor records, Sub Budget Heads, Activity Masters — these are *master data*, owned by the client team, not by code. The handover assumes they already exist on every environment with the right names.

---

## E. Items that should be Server Scripts / app hooks (currently Client Scripts)

The reason these ended up as client scripts is **time-to-iterate** on staging without bench/SSH. They work — but they have hidden costs (cannot be unit-tested, validation logic runs in browser only, can be bypassed by anyone with REST API access). For prod hardening, please consider promoting:

| Currently Client Script | Should be | Reason |
|---|---|---|
| Auto-population of PBP rows from Activity KPI selections (in `budget_allocation_v9_final.js`) | Server hook on `Project proposal.before_save` | Currently runs on form refresh; if a user creates the GAF via REST or import, no rows are created. Server hook would guarantee parity. |
| BPU creation from PBPs after GAF approval | Server hook on `Project proposal.on_workflow_action` (final Approve transition) | Currently I created BPUs **manually** for GAF-0545 via the `patch_gaf_0545_budget` Server Script. The clean-path workflow probably triggers this server-side already, but our path bypassed it. **Verify on a fresh GAF — does the Approve flow actually create BPUs in this app version?** If not, this is a critical bug, not a nice-to-have. |
| BPU `workflow_state = "Approved"` after creation | Server hook | The QUR API `mgrant.apis.quarterly_utilisation_report.quarterly_utilisation_report.get_quarterly_utilisation_report` filters BPUs by `workflow_state = "Approved"`. Created-via-insert BPUs start at "Pending", and the QUR dialog says "No plannings found". Need a server-side default-to-Approved on BPU after PBP-derived insert. |
| Over-spend handling in QUR (`actuals = min(real, plan)` cap, real units in `custom_actual_units`) | Relax the Python validation in mgrant app (the `Actual Amount cannot be greater than Budget Amount` check) | The cap-at-budget workaround works for honest reporting, but downstream Frappe reports reading `actuals` undercount. The right fix is one line of Python in the mgrant app — see Section F.5. |
| `Project proposal.total_planned_budget` allow-on-submit override | Hook patch in mgrant app's "Cannot Update After Submit" hook | Property Setter `allow_on_submit=1` is overridden by the hook. We currently bypass via `frappe.db.set_value`. Right fix: app-level allowlist for this fieldname. |
| `Grant.total_planned_budget` is `read_only=1` and `frappe.client.set_value` no-ops on it | Schema change OR Server Script that uses `frappe.db.set_value` directly, called from a hook | After bulk BPU creation, BPU save hooks reset `Grant.total_planned_budget` to 0. We currently re-set via direct DB write. Right fix: BPU save hook should *increment* not *reset*. |

---

## F. Bypasses, hardcodes, and one-shot patches — DO NOT replicate to prod

These are the scariest bits. Read carefully.

### F.1 `patch_gaf_0545_budget` Server Script

A reusable **one-shot Server Script** in DB on staging, currently disabled. API method = same name. Body is edited in-place: I set `disabled=0`, POST `/api/method/patch_gaf_0545_budget`, then set `disabled=1`. I used it for:

- Direct DB patches that bypass mgrant's "Cannot Update After Submit" hook
- BPU bulk creation/update from PBPs (because the workflow path didn't trigger app-level BPU creation)
- Approval state writes on BPUs (because they default to "Pending" and the QUR API filters on "Approved")

**Do NOT push this to prod.** It is GAF-0545-specific and the mechanism (toggle enabled on/off, POST, toggle off) is brittle. The work this script did should be replaced by:
- Server hook on Project proposal final-approve → creates BPUs from PBPs server-side
- Server hook on BPU after_insert → set `workflow_state = "Approved"` (or use proper workflow apply)

### F.2 Hardcoded `D-0001` donor / fund_source

In `budget_allocation_v9_final.js`:
- Line ~1488: `donor: 'D-0001'`
- Line ~1614: `donor: 'D-0001'`
- Line ~1618: `fund_source: 'D-0001'` ← **bug, this should be a Fund Source ID, not a Donor ID. Inherited from v8c. Verify whether `fund_source` is consumed downstream — if yes, fix; if no, remove.**

If the prod environment has a different donor ID, change these to match. Better fix: read the donor from the GAF's `ngo` profile or accept it as form input.

### F.3 Quarter convention (4-row override)

The system's auto-generator on save slices `quaterly_project` along **Indian fiscal-year boundaries** (Apr–Mar). LIC HFL source budgets use **project-start-based 3-month quarters**. If `project_start_date` isn't on a fiscal boundary (Apr/Jul/Oct/Jan), the system creates **5+ quarter rows** (with an orphan tail), and the budget UI renders an extra column band before Y_n Total.

**Per-GAF workaround applied on GAF-0545:** I replaced `quaterly_project` via `frappe.client.save` with 4 manual rows (project-start + 3 months each). PBP `planning_table` re-binds because it stores quarter labels (`Q1`–`Q4`) not link IDs.

**The right fix is a server-side hook patch** on Project proposal save that computes quarters from project_start_date instead of fiscal year. I have not done this.

> ⚠️ Don't overwrite `annual_project` simultaneously — that table has a different doctype name and the save will fail.

### F.4 `frappe.db.set_value` for read-only / allow-on-submit fields

Multiple places (see memory `project_lic_hfl_gaf_approval_chain.md`) where `frappe.client.set_value` returns 200 but no-ops because the field is `read_only=1` or post-submit-locked. I worked around this by running `frappe.db.set_value(...)` inside the `patch_gaf_0545_budget` Server Script. This is a **sandbox bypass**, not a real fix.

Fields affected:
- `Grant.total_planned_budget`
- `Project proposal.total_planned_budget` and `overall_planned_budget` (post-submit)

### F.5 Over-spend cap-at-budget in QUR

`lic_hfl_qur_v2.js` ~line 880:
```js
const cappedActuals = Math.min(realActuals, r.planAmt);
// Real units stored in custom_actual_units; comment mandatory if real > plan
```

Real units stored in `custom_actual_units`; server `actuals` is capped at plan to satisfy the validator. Reviewer sees over-spend visually (red row + comment). Downstream Frappe reports reading `actuals` undercount.

**Fix in mgrant app:** relax (or remove) the `Actual Amount cannot be greater than Budget Amount` Python validator. Once relaxed, the `Math.min` cap can be removed and `actuals` will hold the truth. No data migration needed — `custom_actual_units` already has it.

### F.6 RestrictedPython sandbox quirks (only if you ever revive `patch_gaf_0545_budget`)

- No `import` — use the pre-injected `frappe`
- No augmented item assignment (`d[k] += x` fails)
- No `frappe.get_roles()`
- Only SELECT in `frappe.db.sql` — use `frappe.db.set_value` in a loop instead

### F.7 GAF-0545-specific data state

These are *demo data* state, not behaviours, but flagging so you don't accidentally copy them:
- Approved end-to-end on staging: Project P-GAF-0545 (Active), Grant P-GAF-0545 / Year 1 / FY-2025 (Active), 29 BPUs (Approved, total ₹1,28,97,600 LIC portion), QU-0672 (Q1, ₹38.33L), QU-0706 (Q2, ₹26.28L)
- Board Note + MOU + Sanction Letter are **placeholder PDFs** at `/files/GAF-0545_BoardNote_placeholder.pdf` — replace with real signed docs when going to prod
- `Grant.custom_pm` and `Grant.custom_spm` set to `USER-25-0006` (placeholder, full_name "PL")
- `Mahasabha` Q2 row was over-spent (real 2 vs plan 1) — server `actuals` capped at plan, real units preserved in `custom_actual_units`; this is a deliberate test artifact, not data corruption

---

## G. Recommended deployment order on a new environment

If the developer is doing the prod cutover (or even setting up a UAT clone), do this strictly in order:

1. **Pre-flight: master data check.** Confirm Donor `D-0001`, the four named Sub Budget Heads, Activity Master + Units exist with the expected naming. If not, create them or update the hardcoded references in F.2.
2. **Apply Custom Field fixtures** (Section B.3 / Section D). Verify on the desk: `Project Budget Planning` form should now show the four `custom_*` fields; `Budget Planning Child` should show `unit` + `unit_cost`; `Quarterly Utilisation Report Child` should show the three `custom_*` fields.
3. **Apply Property Setter fixtures** (Section B.4). Verify the Project proposal `stage` dropdown contains BOTH old and new label sets.
4. **Verify the workflow override.** Trigger a dummy GAF transition Pending at RPL → Pending at PL with the popup. If you get `Stage cannot be "GAF Submitted"…`, the Property Setter import didn't widen the stage options — fix before continuing.
5. **Merge branch `chore/sync-staging-2026-04-27` to `main`** if not already done, so `git clone` gives you the current code. (As of 27 Apr 2026 the branch is pushed but not yet merged.)
6. **Deploy Client Scripts** in this order, pasting contents from the merged `main`:
   1. `Activity KPI NGO Filter` ← `activity_kpi_ngo_filter_client_script.js`
   2. `LIC Budget Allocation v4` ← `budget_allocation_client_script.js` (now 184 KB v9 final)
   3. `LIC HFL QUR v2` ← `lic_hfl_qur_v2.js`
7. **Smoke test on a fresh GAF:**
   - Create a GAF, select an NGO, ensure Activity KPIs auto-populate.
   - Open Budget Allocation tab → Programmatic + Non-Programmatic should render.
   - Enter values, Save. Confirm PBPs are created with `custom_total_lic_contribution` populated.
   - Run the GAF through the workflow. **CRITICAL: confirm BPUs are auto-created on the Grant when the GAF reaches "Approved".** If not — that's the missing server hook from Section E. Manually create BPUs for the smoke test only.
   - Open the Grant page → Budget Allocation tab → both YTD summary and QUR widget should render with 0 actuals.
   - Set `Grant.custom_pm` and `Grant.custom_spm`. File Q1 actuals, Submit. Walk through the QUR approval workflow.
8. **Only after all above pass on UAT/staging clone, push to prod.**

---

## H. Open questions / things I'm not sure about

These are decisions I made on staging that the developer should validate against application owners / mgrant codebase before prod:

1. **Does the standard mGrant flow auto-create BPUs from PBPs on GAF approval?** I bypassed the workflow with direct DB writes for GAF-0545, so I never observed the clean path. If yes — great, drop the BPU creation block from F.1. If no — that's a critical missing feature and needs an app-level hook.
2. **Is `fund_source: 'D-0001'` (a Donor ID) being saved into `Project Budget Planning.fund_source` (a Link to Fund Sources) actually wrong?** Inherited from v8c. I haven't seen it cause errors, but the type mismatch is suspicious. Worth one Sentry/log search.
3. **Activity matching by name string vs ID.** v9 still falls back to `description` matching when the new `custom_activity` Link isn't populated. Backfill plan needed for legacy proposals.
4. **The "Budget Summary" tab on the Project proposal form.** Hidden via `ab_hideBudgetSummaryTab`. Exists in Excel export only. Future feature — not blocking.
5. **Excel export library (`xlsx-js-style@1.2.0`) loaded from `cdn.jsdelivr.net`.** Will the prod environment have outbound internet to fetch it? If not, host locally or remove the download button.
6. **The disabled `LIC Activity Utilisation` v1 Client Script on Grant.** Should we delete it on prod, or keep as a rollback safety net? My vote: delete on prod, keep on staging until QUR v2 has 1–2 quarters of stable use.

---

## I. Files in the repo (post-sync)

After branch `chore/sync-staging-2026-04-27` merges to `main`, the developer pulls these:

| File | Purpose |
|---|---|
| `budget_allocation_client_script.js` | **Deploy as `LIC Budget Allocation v4` Client Script** (Project proposal) |
| `activity_kpi_ngo_filter_client_script.js` | **Deploy as `Activity KPI NGO Filter` Client Script** (Project proposal) |
| `lic_hfl_qur_v2.js` | **Deploy as `LIC HFL QUR v2` Client Script** (Grant) |
| `legacy/budget_utilisation_v3.js` | Archived fallback — do not deploy |
| `HANDOVER_DEVELOPER.md` | This document |
| `HANDOVER_v9.md` | Earlier handover focused on v9 budget allocation only |
| `LIC_HFL_Budget_Allocation_Handover.md` | Earlier (v8c-era) handover, retained for DocType setup notes |
| `QUR_user_guide.md` | Partner-facing user guide for QUR — share with ops/PM team, not the developer |
| `GIT_PUSH_PLAN.md` | Audit trail of the 27 Apr sync push |
| `README.md` | Repo entry-point |

---

## J. Contact / context for the developer

If anything in this document is unclear, escalate questions in this order:
1. The two project memory files (paths in my head, will share separately) — they have the why behind decisions.
2. Me (Sunandan).
3. The mgrant / frappe_theme app maintainers — for anything in Section E (server-side hardening) or Section F.5 (validator relaxation).

> **Bottom line:** Do not assume "git pull = production-ready". The code in the repo is necessary but not sufficient. Custom Fields + Property Setters in Section B must be carried over via fixtures (Section D) BEFORE the JS will work, and the bypasses in Section F must be replaced by proper server-side fixes BEFORE we sign off prod.
