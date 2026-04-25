# LIC HFL — Budget Allocation v9: Session Handover

**Date:** 25 April 2026
**Prepared by:** Sunandan (via Claude AI pair-programming)
**Staging:** `stg.lichfl.mgrant.in` → GAF-0322
**Git repo:** https://github.com/sunandan89/lic-hfl-budget-allocation.git (branch: `main`)
**Script:** `budget_allocation_client_script.js` (~2400 lines, ~101KB)
**Deployed as:** Client Script "LIC Budget Allocation v4" on staging

---

## What Was Done (v9 — complete rework from v8c)

### 1. Programmatic Costs Tab — Column Layout Rework
- **Old layout (v8c):** 5 frozen cols + 7 per-quarter sub-columns (LIC Units/Amt, Govt Units/Amt, Benf Units/Amt, Total Amt) + Year Totals + Grand Total Y1/Y2/Combined
- **New layout (v9):** 5 frozen cols (Sr, Activity, Task Details, UoM, Unit Cost) + 5 Convergence cols (Total Units, Total Cost, LIC HFL Contribution, Govt Contrib, Benf Contrib) + 4 per-quarter sub-columns (Units, Cost, LIC HFL, Benf) + Year Totals (Y1 Total, Y2 Total) + Remarks

### 2. Auto-Calculation Formulas (verified against client Excel)
- Total Cost = Unit Cost × Total Units
- Total LIC HFL = Total Cost − Govt Contribution − Beneficiary Contribution
- Quarterly Cost = Unit Cost × Units to Cover
- Quarterly LIC HFL = (Total LIC HFL / Total Units) × Units to Cover
- Quarterly Beneficiary = Quarterly Cost − Quarterly LIC HFL
- Y1/Y2 Totals = Sum of quarterly costs per year

### 3. Non-Programmatic Costs Tab Updates
- Added Total Units (auto-calc) and Total Cost (auto-calc) columns
- 6 frozen columns: Sr, Particulars, UoM, Unit Cost, Total Units, Total Cost
- Year-grouped quarter headers (Option B nested style)
- Y1 Total / Y2 Total columns (replaced old "Combined" columns)
- Remarks column added

### 4. Task Details — Editable
- Changed from read-only text to editable `<input type="text">`
- Pre-populated from PBP `assumption` field

### 5. Activity Bug Fix
- Added `after_save` handler that clears `frm._ab_rendered_for`
- New activities added after saving now appear without hard reload

### 6. Year-Grouped Headers (Option B)
- Row 1: Maroon strip with "Activity Details" (frozen) + Convergence + Year 1 (thin label above Q1-Q4) + Year 2 + Grand Total (Y1/Y2 Total) + Remarks
- Row 2: Column sub-headers
- Same pattern for NP sections

### 7. UI/UX
- Platform-aligned maroon (#8B1A1A) colour scheme — no blue/yellow on web UI
- Frozen column separator (2px border + shadow)
- `table-layout: fixed` with explicit `max-width` on frozen columns for stable widths
- Vertical scroll with sticky headers (`max-height: 70vh; overflow: auto`)
- Horizontal scroll with frozen columns (`position: sticky; left: Npx`)
- Indian number formatting (toLocaleString en-IN)

### 8. Excel Export
- 3-sheet workbook: Budget Summary, Programmatic Costs, Non-Programmatic Costs
- Client branding colours in Excel: blue (#00529C) for Q1/Q3, yellow (#FFCB05) for Q2/Q4
- Convergence columns (H, I, J) included in Programmatic sheet
- Budget Summary with project details, quarterly breakup, cost sharing
- NP sheet with section summaries at top

### 9. Save Logic
- Programmatic: single PBP record per activity (not 3 per fund source as in v8c)
- Non-Programmatic: sequential saves to avoid document lock conflicts
- Both use `frappe.client.save` via API

### 10. Patterns Library
- 5 patterns pushed to https://github.com/sunandan89/mgrant-frappe-patterns
- Updated pattern discovery prompt

---

## What Was NOT Achieved / Known Issues

### 1. Hover-Expand for Activity/Task Details
- **Attempted:** CSS hover-expand to show full text on hover
- **Failed:** `position: sticky` (for freeze) conflicts with `overflow: hidden` (for truncation) on the same `<td>`. Inner `<div>` wrappers also caused layout breaks.
- **Current state:** Reverted to native browser `title` tooltip on Activity cells. Task Details and NP Particulars are plain inputs with no expand.
- **Recommendation:** Implement via JavaScript tooltip (e.g., Frappe's `frappe.ui.Tooltip` or a custom positioned `<div>` created on mouseenter) instead of CSS-only approach.

### 2. Excel File Corruption
- **Issue:** Downloaded XLSX has "repair" warning in Excel. Root cause: `xlsx-js-style` library generates invalid merge references for columns beyond Z (AA+) and includes XLDAPR metadata.
- **Partial fix:** All merges use numeric indices now. But the metadata.xml issue is a library limitation.
- **Recommendation:** Either suppress metadata generation in xlsx-js-style config, or post-process the ZIP to remove `xl/metadata.xml` before download.

### 3. Activity Matching by ID (not name)
- **Parked:** Current matching between PBP records and Activity KPIs is by `description` (name string). If an activity is renamed in the Activity Master, the old PBP record won't match.
- **Recommendation:** Add an `activity_master` Link field to the PBP DocType, store the Activity Master ID, and match by ID instead of name.

### 4. Data Migration from v8c
- **Issue:** Existing PBP records from v8c store data per-fund-source (3 records per activity: LIC, Govt, Benf). The v9 model expects 1 record per activity with lumpsum convergence fields.
- **Current behavior:** v9 reads the first PBP record it finds per activity description. Old quarterly units show up but convergence fields (Total Units, Govt, Benf) default to 0.
- **Recommendation:** Run a one-time migration script to consolidate old 3-record-per-activity data into the v9 single-record format.

### 5. Non-Programmatic Recalc Cell Indices
- **Risk:** The `ab_recalcNonProgRow` and `ab_recalcNonProgSectionTotal` functions use hardcoded cell indices. After adding frozen columns and year totals, these indices may be off.
- **Recommendation:** Test NP recalc thoroughly — enter values and verify auto-calculations update correctly. If broken, update cell index offsets.

### 6. Budget Summary Tab
- Hidden via `ab_hideBudgetSummaryTab`. Not implemented on the web UI — only exists in the Excel export.

---

## File Manifest

| File | Purpose |
|------|---------|
| `budget_allocation_client_script.js` | Main v9 script (~2400 lines) — paste into Frappe Client Script |
| `activity_kpi_ngo_filter_client_script.js` | Activity filter dependency (19 lines) |
| `HANDOVER_v9.md` | This document |
| `LIC_HFL_Budget_Allocation_Handover.md` | Original v8c handover (still relevant for DocType setup) |

---

## Git History (key commits)

```
ab0aa4d Freeze NP columns through Total Cost (6 frozen cols)
8b71b22 Fix: pass years param to buildProgRow and buildProgGrandTotal
301bca9 Add Year spans + Grand Total columns (Option B nested headers)
eaaa84d Revert hover-expand — back to stable version
f4bf4f6 Fix frozen column width stability — add max-width + table-layout:fixed
eb8c8ad Merge frozen header cells into single sticky cell with left:0
d0141eb Fix quarterly column widths — Units 65px, Cost 75px
c5af313 Fix scroll, reduce frozen cols, compress widths
8f0fe89 Fix header row column alignment — Convergence spans cols 7-9
7ca467a v9 complete: column rework + NP update + Excel export + UI refresh
d5caad3 Fix bug: new activities not appearing after budget is saved
```

---

## Frappe Patterns Library Status

The `mgrant-frappe-patterns` repo has 7 patterns. Two learnings from this session that could update existing patterns:

1. **sticky-table-freeze** — needs update:
   - Document that `overflow: hidden` on `<td>` conflicts with `position: sticky`
   - Add the `table-layout: fixed` + `max-width` technique for stable frozen column widths
   - Add the merged header cell approach (`colspan` with `left:0`) for multi-row header freeze
   - Note: `max-height: 70vh` + `overflow: auto` on wrapper for vertical scroll

2. **New potential pattern: nested-year-quarter-header**
   - The Option B nested header technique (year label as thin sub-row within a `<th>` using flexbox `<div>` layout) is reusable for any multi-year quarterly table.

**Recommendation:** Update the sticky-table-freeze pattern with the learnings before the next project. The nested header could be a new pattern if it proves stable.
