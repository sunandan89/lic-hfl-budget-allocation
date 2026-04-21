# LIC HFL — Budget Allocation Feature: Developer Handover

**Date:** 21 April 2026
**Prepared by:** Sunandan (via Claude AI pair-programming)
**Staging reference:** `stg.lichfl.mgrant.in`
**Git repo:** https://github.com/sunandan89/lic-hfl-budget-allocation.git (branch: `main`)

---

## What This Feature Does

The Budget Allocation feature adds an interactive, spreadsheet-like budget planning interface inside the **Project Proposal (GAF)** form. It renders under the **Budget Allocation** tab and has two sub-tabs:

- **Programmatic Costs** — activity-level budget rows (auto-populated from the NGO's activity profile), with quarterly breakdowns across LIC HFL, Government, and Beneficiary funding sources.
- **Non-Programmatic Costs** — three collapsible sections: Human Resource Costs, Administration Costs, and NGO Management Costs, each with quarterly unit/cost breakdowns.

Users can edit values inline, save to the backend, and download a styled Excel workbook matching the LIC HFL budget format.

---

## Architecture Overview

This feature is built entirely as **client-side configuration** — no server-side Python code, no custom API endpoints, no bench commands. It consists of:

| Layer | What | How |
|-------|------|-----|
| **No-code** | DocTypes, fields, master data | Frappe Form Builder / Admin Setup |
| **Low-code** | 2 Client Scripts | Site-level Client Scripts (paste JS) |
| **External** | Excel export library | CDN (loaded at runtime in browser) |

---

## Step-by-Step Deployment Guide

### Step 1: Verify DocTypes and Fields Exist

These DocTypes and fields must already exist on the target server. If you're running the same mGrant app version as staging, they should all be present. Verify each one.

#### 1a. Project Proposal DocType

The `Project proposal` DocType must have these fields:

| Field Name | Type | Purpose |
|------------|------|---------|
| `cummulative_budget` | HTML | Container where the budget tab renders |
| `quaterly_project` | Table (child) | Quarters config — each row has: `year`, `quarter`, `year_sequence`, `quarter_sequence`, `start_date`, `end_date` |
| `annual_project` | Table (child) | Years config — each row has: `year`, `year_sequence` |
| `custom_activity` | Table (child) | Activity selections — each row has: `activity` (Link to Activity Master) |
| `project_name` | Data | Project name (used in Excel export) |
| `start_date` | Date | Project start date |
| `end_date` | Date | Project end date |
| `total_beneficiaries` | Int | Beneficiary count |
| `state` | Data/Link | State name |
| `ngo` | Link | NGO/Partner link |

**Critical:** The `cummulative_budget` HTML field must be placed inside a tab section called "Budget Allocation" (or similar) on the form layout.

#### 1b. Project Budget Planning (PBP) DocType

This is the backend storage for budget data. Each row in the budget tab maps to one or more PBP records.

| Field Name | Type | Purpose |
|------------|------|---------|
| `project_proposal` | Link → Project proposal | Parent proposal |
| `description` | Data | Activity/line-item name |
| `budget_head` | Link → Budget heads | Budget head category |
| `sub_budget_head` | Link → Sub Budget Head | Determines Programmatic vs Non-Programmatic placement |
| `fund_source` | Link → Fund Sources | LIC HFL / Government / Beneficiary |
| `donor` | Link → Donor | Always set to "D-0001" |
| `total_planned_budget` | Currency | Computed total |
| `assumption` | Data | Task details / UoM text |
| `planning_table` | Table (child of PBP Child) | Quarterly breakdown |

#### 1c. PBP Child DocType

Child table of Project Budget Planning, stores per-quarter data:

| Field Name | Type | Purpose |
|------------|------|---------|
| `timespan` | Data | Quarter identifier (e.g., "Q1", "Q2") |
| `planned_amount` | Currency | Budget amount for that quarter |

#### 1d. Master DocTypes (read-only references)

| DocType | Key Fields | Purpose |
|---------|-----------|---------|
| `Budget heads` | `name`, `budget_head_name` | Top-level budget categories |
| `Sub Budget Head` | `name`, `sub_budget_head`, `budget_head` (Link) | Sub-categories that determine tab placement |
| `Fund Sources` | `name`, `source_name` | Funding source names |
| `Activity Master` | `name`, `activity_name`, `custom_unit_of_measurement` (Link to Units) | Activity definitions |
| `Units` | `name`, `unit_name` | Unit of measurement master |
| `Donor` | `name` | Donor records |

---

### Step 2: Create Required Master Data

If master data doesn't exist, create it through the Frappe desk UI.

#### 2a. Donor Record

Go to **Donor list** → Create new:
- **Name/ID:** `D-0001`
- This is hardcoded in the save logic. The script sets `donor: 'D-0001'` on every new Project Budget Planning record.

#### 2b. Sub Budget Heads

Go to **Sub Budget Head list** → Ensure these 4 records exist with **exact names**:

| Sub Budget Head Name | Parent Budget Head | Tab Placement |
|---------------------|--------------------|---------------|
| `Programmatic Costs` | (your programmatic budget head) | Programmatic Costs tab |
| `Human Resource Costs` | (your non-prog budget head) | Non-Programmatic → Section A |
| `Administration Costs` | (your non-prog budget head) | Non-Programmatic → Section B |
| `NGO Management Costs` | (your non-prog budget head) | Non-Programmatic → Section C |

**These names are exact-match strings in the code.** If a sub-budget head is named differently (e.g., "HR Costs" instead of "Human Resource Costs"), it won't appear in the correct section.

#### 2c. Fund Sources

Go to **Fund Sources list** → Ensure at least these exist:

| Source Name | Mapped Column |
|-------------|--------------|
| Any name containing "LIC" or "HFL" | LIC HFL column |
| Any name containing "Gov" | Government column |
| Any name containing "Ben" | Beneficiary column |

The script normalizes fund source names using substring matching (case-insensitive). For example, "LIC HFL Contribution" → LIC column, "Government" → Govt column, "Beneficiary Contribution" → Benf column.

#### 2d. Activities and Units

- **Activity Master:** Create activity records that represent the programmatic activities (e.g., "Establishment of Backyard Poultry", "Health Camps", etc.). Link each to a Unit of Measurement.
- **Units:** Create UoM records (e.g., "Numbers", "Farmers", "No. of Schools").
- **NGO Profile:** Assign activities to the NGO/partner profile so they can be auto-populated into proposals.

---

### Step 3: Configure Project Proposal Form Layout

In the **Form Builder** for `Project proposal`:

1. Ensure there is a tab section for "Budget Allocation" (or equivalent)
2. The `cummulative_budget` (HTML) field must be inside this tab
3. The `quaterly_project` and `annual_project` child tables must be populated (typically in the "Timeframe" tab) — these define how many quarters and years the budget grid shows

---

### Step 4: Deploy Client Script 1 — Activity KPI NGO Filter

This script is a **pre-existing dependency** (not written as part of this project). It handles:

- Filtering activities from the selected NGO/partner profile
- Populating the `custom_activity` child table on the Project Proposal form
- This is what causes the Programmatic Costs tab to show pre-selected activity rows

**How to deploy:**
1. Get the script from either:
   - Git repo: `https://github.com/sunandan89/lic-hfl-budget-allocation.git` → file `activity_kpi_ngo_filter_client_script.js`
   - Or copy from staging: `stg.lichfl.mgrant.in` → Client Script list → open **"Activity KPI NGO Filter"** → copy script
2. On the target server: Client Script list → **+ Add Client Script**
   - **Name:** `Activity KPI NGO Filter`
   - **DocType:** `Project proposal`
   - **Script Type:** Client
   - **Enabled:** ✓
   - Paste the script
3. Save

**Without this script:** The budget tab will still work, but the Programmatic Costs tab will only show rows from existing PBP records — no auto-population of activities from the NGO profile.

The script is available in the Git repo as `activity_kpi_ngo_filter_client_script.js` (19 lines). It sets a `get_query` filter on the `custom_activity` child table so the Activity dropdown only shows activities linked to the selected NGO partner. It re-applies the filter when the `ngo` field changes.

---

### Step 5: Deploy Client Script 2 — LIC Budget Allocation v4

This is the main budget allocation script (v8c, ~2,350 lines).

**How to deploy:**
1. Get the latest code from either:
   - Git repo: `https://github.com/sunandan89/lic-hfl-budget-allocation.git` → file `budget_allocation_client_script.js`
   - Or the file `budget_allocation_v8c.js` provided alongside this handover
2. On the target server: Client Script list → **+ Add Client Script**
   - **Name:** `LIC Budget Allocation v4` (or any name you prefer)
   - **DocType:** `Project proposal`
   - **Script Type:** Client
   - **Enabled:** ✓
   - Paste the full script code
3. Save
4. Hard-reload any open Project Proposal form (Ctrl+Shift+R) to pick up the new script

---

### Step 6: Verify

1. Open any Project Proposal that has:
   - Quarters and years configured (in the Timeframe tab)
   - An NGO selected (with activities in their profile)
2. Click the **Budget Allocation** tab
3. You should see:
   - **Programmatic Costs** sub-tab with activity rows auto-populated
   - **Non-Programmatic Costs** sub-tab with 3 collapsible sections
   - Editable input fields for units and costs
   - "Save Programmatic" / "Save Non-Programmatic" buttons
   - "Download Budget Sheet" button (top-right)
4. Test horizontal scroll — the first 5 columns (Sr, Activity, Task Details, UoM, Unit Cost) should remain frozen
5. Test the Excel download — should produce a 3-sheet workbook with styling

---

## What Each Component Does

### Client Script: Activity KPI NGO Filter
- **Trigger:** When NGO field changes on the Project Proposal form
- **Action:** Fetches activities from the NGO's profile, populates `custom_activity` child table
- **Impact on budget:** The budget script reads `custom_activity` to auto-populate Programmatic Costs rows

### Client Script: LIC Budget Allocation v4 (v8c)
- **Trigger:** `refresh` and `onload` events on Project Proposal form
- **Renders into:** `cummulative_budget` HTML field
- **Key behaviors:**
  - Fetches quarters, years, budget heads, sub-budget heads, fund sources, PBP records, and activity KPIs in parallel
  - Builds a tabbed interface (Programmatic / Non-Programmatic)
  - Programmatic tab: auto-populates rows from Activity KPI selections, shows 7 sub-columns per quarter (LIC Units, LIC Amt, Govt Units, Govt Amt, Benf Units, Benf Amt, Total Amt), plus year totals and grand totals
  - Non-Programmatic tab: 3 sections with 2 sub-columns per quarter (Units, Cost), allows adding/removing rows
  - Saves data as Project Budget Planning records with PBP Child quarterly entries
  - Excel export: generates a styled 3-sheet XLSX matching the LIC HFL budget format, with sheet protection enabled
  - Monkey-patches `$wrapper.html()` to prevent Frappe from overwriting the custom HTML on form refresh

---

## External Dependency

The Excel export feature loads the **xlsx-js-style** library from CDN at runtime:

```
https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js
```

The target server's browser must have internet access for the download button to work. If the environment is air-gapped, you would need to either host this library locally or remove the download feature.

---

## Known Limitations / Notes

1. **Donor is hardcoded:** The save logic always sets `donor: 'D-0001'`. If your donor ID differs, update lines 1239 and 1384 in the script.

2. **Sub-budget head names are exact-match:** The strings "Programmatic Costs", "Human Resource Costs", "Administration Costs", "NGO Management Costs" must match exactly in the Sub Budget Head master.

3. **Fund source matching is substring-based:** Any fund source containing "LIC" or "HFL" maps to the LIC column; "Gov" → Government; "Ben" → Beneficiary. All others default to LIC.

4. **Budget Summary tab:** There is a third tab (Budget Summary) that is currently hidden/not implemented. The function `ab_hideBudgetSummaryTab` hides it. This is planned for future development.

5. **Sequential saves:** The Non-Programmatic save uses sequential (not parallel) saves to avoid Frappe document locking conflicts. This means saving many rows may take a few seconds.

6. **No server-side validation:** All logic runs client-side. There is no server-side script enforcing budget rules or totals.

---

## File Manifest

| File | Purpose |
|------|---------|
| `budget_allocation_client_script.js` | Main budget allocation script — v8c (paste into Frappe as "LIC Budget Allocation v4") |
| `activity_kpi_ngo_filter_client_script.js` | Activity filter script (paste into Frappe as "Activity KPI NGO Filter") |
| `LIC_HFL_Budget_Allocation_Handover.md` | This document |
| `budget_allocation_v8c.js` | Copy of main script for quick access |

---

## Version History

| Version | Commit | Changes |
|---------|--------|---------|
| v5 | `b787262` | Initial 3-fund-source layout |
| v7f | `cc042f6` | Save buttons, UoM dropdown, compact UI |
| v8 | `e3ba5e5` | Excel export with styled 3-sheet workbook |
| v8b | `a1dd21c` | Sticky header fix attempt + sheet protection |
| v8c | `40b2d62` | Fixed frozen columns via `border-collapse: separate` |
