# LIC HFL — Budget Allocation + Quarterly Utilisation Reporting

Client Scripts and handover documentation for the LIC HFL CSR grant management
platform on Frappe v16 (mGrant app), deployed at `stg.lichfl.mgrant.in`.

## What's in this repo

| File | Deploy as | DocType |
|---|---|---|
| `budget_allocation_client_script.js` | Client Script "LIC Budget Allocation v4" | Project proposal |
| `activity_kpi_ngo_filter_client_script.js` | Client Script "Activity KPI NGO Filter" | Project proposal |
| `lic_hfl_qur_v2.js` | Client Script "LIC HFL QUR v2" | Grant |
| `legacy/budget_utilisation_v3.js` | (archived — superseded by QUR v2) | — |

## What's NOT in this repo (must recreate on target environment via fixtures)

- 11 Custom Fields (Project Budget Planning, Budget Planning Child, Quarterly Utilisation Report Child)
- 4 Property Setters (Project proposal stage options widening, MOU/Board Note labels)

See `HANDOVER_DEVELOPER.md` Sections B.3 and B.4 for the exhaustive list, and
Section D for a fixtures `hooks.py` snippet.

## Where to start

1. Read `HANDOVER_DEVELOPER.md` end-to-end before deploying anything.
2. Read `QUR_user_guide.md` if you need to brief partners on the QUR widget.
3. Frappe Patterns library used: `https://github.com/sunandan89/mgrant-frappe-patterns` (12 patterns; load-bearing ones listed in HANDOVER_DEVELOPER.md Section C.2).

## Status

- Last sync to staging: 27 April 2026
- End-to-end validated on: GAF-0545 → Project P-GAF-0545 → Year 1 / FY-2025
- Open questions and unresolved items: see HANDOVER_DEVELOPER.md Section H

## Older handover snapshots (kept for context)

- `HANDOVER_v9.md` — point-in-time handover from 25 April (budget allocation only)
- `LIC_HFL_Budget_Allocation_Handover.md` — earlier v8c-era handover (DocType setup notes still relevant)

For all current guidance, prefer `HANDOVER_DEVELOPER.md`.
