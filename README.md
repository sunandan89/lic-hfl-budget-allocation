# LIC HFL Budget Allocation - Client Script

Custom Frappe Client Script for the **Budget Allocation** tab on the **Project Proposal** form in the mGrant application (LIC HFL grant management system).

## Overview

This script renders a custom budget planning UI inside the `cummulative_budget` HTML field on the Project Proposal form. It replaces Frappe's default rendering with a 3-fund-source layout (LIC HFL, Government, Beneficiary).

## Features

- **Two-tab layout**: Programmatic Costs (Direct) and Non-Programmatic Costs (Indirect)
- **7 sub-columns per quarter**: LIC Units, LIC Amt, Govt Units, Govt Amt, Benf Units, Benf Amt, Total Amt
- **Frozen columns**: Sr, Activity, Task Details, UoM, Unit Cost stay visible while scrolling
- **Auto-calculated amounts**: Unit Cost x Units = Amount, with quarterly/yearly/grand totals
- **Indian number formatting** (en-IN locale)
- **Collapsible sections** for non-programmatic categories (HR, Admin, NGO Management)
- **Year and quarter headers** with date ranges from the proposal's timeframe config

## Technical Details

- **DocType**: Project proposal (form target)
- **Script Type**: Site-level Client Script (`LIC Budget Allocation v4`)
- **Data Source**: Project Budget Planning records linked to the proposal
- **Fund Source Mapping**: Normalized from Fund Sources DocType (`LIC HFL` -> lic, `Government` -> govt, `Beneficiary` -> benf)
- **Monkey-patch**: jQuery `.html()` on the target wrapper is patched to prevent Frappe's form renderer from overwriting the custom UI

## Deployment

This script is deployed as a Client Script record on the Frappe instance:
- **Name**: `LIC Budget Allocation v4`
- **DocType**: `Project proposal`
- **Enabled**: Yes

To deploy, update the `script` field of the Client Script document via Frappe API or admin panel.

## Staging

- **URL**: `https://stg.lichfl.mgrant.in`
- **Test Proposal**: GAF-0322
