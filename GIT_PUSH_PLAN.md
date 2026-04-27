# Git Push Plan — close the gap between local + staging vs `lic-hfl-budget-allocation` repo

**Repo:** `https://github.com/sunandan89/lic-hfl-budget-allocation` (branch: `main`)
**Verified state of `main` on 27 April 2026:** 5 files only (see HANDOVER_DEVELOPER.md → C.1).
**Goal of this plan:** make `main` reflect what's actually deployed on `stg.lichfl.mgrant.in` so the developer can `git pull` and reproduce staging.

> Run these commands from a clone of the repo. I have not committed anything from this session — that's your call with your token. The plan below assumes the repo is cloned at `~/lic-hfl-budget-allocation/` and your local working files are in `~/Documents/Claude/Projects/LIC HFL Changes/`. Adjust paths as needed.

---

## Pre-flight

```bash
cd ~/lic-hfl-budget-allocation
git status               # confirm clean
git pull origin main     # confirm up to date
git checkout -b chore/sync-staging-2026-04-27
```

Working on a branch (not pushing direct to `main`) gives you a chance to PR-review the diff before sharing the repo.

---

## Step 1 — Replace the stale Budget Allocation script

The repo has `budget_allocation_client_script.js` at 101 KB (v8c / early-v9). The current shipped script is 184 KB.

```bash
cp "~/Documents/Claude/Projects/LIC HFL Changes/budget_allocation_v9_final.js" \
   "~/lic-hfl-budget-allocation/budget_allocation_client_script.js"

git add budget_allocation_client_script.js
git commit -m "Replace budget_allocation_client_script.js with v9 final (staging parity)

- Convergence custom-field handling (custom_total_lic_contribution, custom_govt_contribution, custom_benf_contribution, custom_task_details)
- NP single-table merge with section banners + grand total
- Excel export rewrite: xlsx-js-style + JSZip post-process, no Microsoft 'repair' warning
- Geography Details extraction via GD-Project proposal-{GAF} doctype
- Budget Summary field-name corrections (project_start_date, custom_direct_beneficiary)
- PBP custom_activity / custom_unit_of_measurement support
- Quarter convention workarounds for project-start-based 4-quarter slicing

Source: budget_allocation_v9_final.js (184 KB, ~4,900 lines) deployed as
'LIC Budget Allocation v4' Client Script on stg.lichfl.mgrant.in."
```

> Optional: if you want to preserve the old 101 KB version for archival, `git mv budget_allocation_client_script.js legacy/budget_allocation_v8c.js` BEFORE the cp+commit above, in a separate commit. But this is just bookkeeping.

---

## Step 2 — Add the QUR / Budget Utilisation Reporting script

This has never been pushed.

```bash
cp "~/Documents/Claude/Projects/LIC HFL Changes/lic_hfl_qur_v2.js" \
   "~/lic-hfl-budget-allocation/lic_hfl_qur_v2.js"

git add lic_hfl_qur_v2.js
git commit -m "Add lic_hfl_qur_v2.js (Quarterly Utilisation Reporting on Grant)

Client Script for the Grant doctype, deployed as 'LIC HFL QUR v2' on
stg.lichfl.mgrant.in.

Two stacked widgets on the Budget Allocation tab:
1. Budget · Year-to-date summary (replaces default mGrant SVADatatable on
   the 'budget' HTML field)
2. Quarterly Utilisation Report widget (replaces default render of
   'quarterly_utilisation' HTML field)

Schema dependencies (must exist on target before deploy):
- Budget Planning Child custom: unit (Float), unit_cost (Currency)
- Quarterly Utilisation Report Child custom: custom_actual_units (Float),
  custom_planned_units (Float, ro), custom_unit_cost (Currency, ro)

v2.3 (~52 KB, ~1,200 lines)."
```

---

## Step 3 — Add archived budget_utilisation_v3.js (optional but tidy)

```bash
mkdir -p legacy
cp "~/Documents/Claude/Projects/LIC HFL Changes/budget_utilisation_v3.js" \
   "~/lic-hfl-budget-allocation/legacy/budget_utilisation_v3.js"

git add legacy/budget_utilisation_v3.js
git commit -m "Archive budget_utilisation_v3.js (superseded by lic_hfl_qur_v2.js)

Earlier single-quarter Budget Report tab render on the Grant doctype.
Superseded by the QUR v2 stacked-widgets approach. Retained for reference."
```

---

## Step 4 — Add the new handover docs

```bash
cp "~/Documents/Claude/Projects/LIC HFL Changes/HANDOVER_DEVELOPER.md" \
   "~/lic-hfl-budget-allocation/HANDOVER_DEVELOPER.md"
cp "~/Documents/Claude/Projects/LIC HFL Changes/QUR_user_guide.md" \
   "~/lic-hfl-budget-allocation/QUR_user_guide.md"
cp "~/Documents/Claude/Projects/LIC HFL Changes/GIT_PUSH_PLAN.md" \
   "~/lic-hfl-budget-allocation/GIT_PUSH_PLAN.md"

git add HANDOVER_DEVELOPER.md QUR_user_guide.md GIT_PUSH_PLAN.md
git commit -m "Add developer handover, QUR user guide, and git sync plan

HANDOVER_DEVELOPER.md — primary developer-facing handover. Covers Custom
Fields + Property Setters that live ONLY on staging (must fixturise),
bypasses/hardcodes that must NOT be replicated to prod, deployment order,
and Frappe Patterns library dependencies.

QUR_user_guide.md — partner-facing user guide for filing the quarterly
utilisation report.

GIT_PUSH_PLAN.md — this file. Documents the gap between repo and staging
as of 27 Apr 2026 and the steps to close it."
```

---

## Step 5 — Update README.md to reflect the new file inventory

The current 1,899-byte README is a stub. Replace it with a real entry-point. Suggested content (drop into `README.md` then commit):

```markdown
# LIC HFL — Budget Allocation + Quarterly Utilisation Reporting

Client Scripts and handover documentation for the LIC HFL CSR grant management
platform on Frappe v16 (mGrant app), deployed at `stg.lichfl.mgrant.in`.

## What's in this repo

| File | Deploy as | DocType |
|---|---|---|
| `budget_allocation_client_script.js` | Client Script "LIC Budget Allocation v4" | Project proposal |
| `activity_kpi_ngo_filter_client_script.js` | Client Script "Activity KPI NGO Filter" | Project proposal |
| `lic_hfl_qur_v2.js` | Client Script "LIC HFL QUR v2" | Grant |

## What's NOT in this repo (must recreate on target environment via fixtures)

- 11 Custom Fields (Project Budget Planning, Budget Planning Child, Quarterly Utilisation Report Child)
- 4 Property Setters (Project proposal stage options widening, MOU/Board Note labels)

See `HANDOVER_DEVELOPER.md` Section B.3 + B.4 for the exhaustive list.

## Where to start

1. Read `HANDOVER_DEVELOPER.md` end-to-end before deploying anything.
2. Read `QUR_user_guide.md` if you need to brief partners on the QUR widget.
3. Frappe Patterns library used: `https://github.com/sunandan89/mgrant-frappe-patterns` (12 patterns; load-bearing ones listed in HANDOVER_DEVELOPER.md Section C.2).

## Status

- Last sync to staging: 27 April 2026
- End-to-end validated on: GAF-0545 → Project P-GAF-0545 → Year 1 / FY-2025
- Open work: see HANDOVER_DEVELOPER.md Section H.
```

```bash
# After editing README.md:
git add README.md
git commit -m "Rewrite README.md as repo entry-point, point to HANDOVER_DEVELOPER.md"
```

---

## Step 6 — Push and open PR (or merge to main directly if you prefer)

```bash
git push -u origin chore/sync-staging-2026-04-27
# Then on GitHub: open PR, review the diff, merge.
```

---

## Sanity checks before sharing the repo

After merging:

```bash
git checkout main
git pull
ls -la                     # 4 .js files (or 3 + legacy/)
git log --oneline | head   # see your sync commits
```

Then visit `https://github.com/sunandan89/lic-hfl-budget-allocation` in a browser and confirm:
- `budget_allocation_client_script.js` is **~184 KB** (not 101 KB)
- `lic_hfl_qur_v2.js` is present at top level
- `HANDOVER_DEVELOPER.md` is present and renders cleanly
- The README points to HANDOVER_DEVELOPER.md

Once that's clean, share the repo URL with your developer.

---

## What I deliberately left out

- I have NOT updated the patterns repo (`mgrant-frappe-patterns`). The HANDOVER_v9.md inside the budget repo still says "7 patterns" but the patterns repo actually has 12. That's a stale reference inside a snapshot doc — leave it; the new HANDOVER_DEVELOPER.md (Section C.2) supersedes it.
- I have NOT pushed the older v7c..v9_activity_fix archives into the repo. They serve no deployment purpose; if you want them as a museum, drop them into `legacy/` in a separate commit.
- I have NOT touched fixtures JSON. That's a separate workstream — see HANDOVER_DEVELOPER.md Section D for the `hooks.py` snippet your developer needs to add to whichever app owns these overrides.
