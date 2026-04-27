# Quarterly Utilisation Report — Partner User Guide

**Platform:** mGrant (LIC HFL staging — `stg.lichfl.mgrant.in`)
**Module:** Grant → Budget Allocation tab
**Audience:** Partner finance/programme team filing the quarterly utilisation report (QUR)
**Last updated:** 26 April 2026

---

## What this is

Every quarter, the partner reports how much of the planned budget was actually utilised. The report is filed activity-by-activity in **units delivered** — the system converts units into rupees automatically using the unit cost agreed in the budget plan.

A quarter is not closed until the LIC HFL programme manager **approves** the submitted report. Until then, the partner can keep editing.

---

## Before you begin

You will need:

1. The **Grant URL** for your project. For Year 1 of *Regenerative Natural Resource Based Livelihoods*, this is:
   `https://stg.lichfl.mgrant.in/app/grant/P-GAF-0545`

2. A **logged-in user** with edit access to that grant.

3. Activity-wise actuals for the quarter — for each programmatic activity and each non-programmatic line, how many units were delivered. (Hours of staff time, number of trainings, person-days, etc., depending on the unit of measurement (UoM) defined for that line.)

4. Supporting documents (vouchers, invoices, photos) ready to attach if relevant — optional but useful.

---

## The page at a glance

When you open the grant and click the **Budget Allocation** tab you see two stacked panels:

### Panel 1 — Year-to-date summary (top)

Aggregated picture of where the project stands across all *approved* quarters so far.

- Four KPI cards: Total Sanctioned Amount, Total LIC Disbursed, Total Utilised, Unutilised
- A consolidated table titled **Budget · Year-to-date summary** with rows for each activity:
  - `Plan units` (annual plan), `Actual units` (cumulative through last approved quarter), `% Delivered`
  - `Plan ₹`, `Actual ₹`, `Variance ₹`, `% Spent` (with a coloured progress bar)
- A grouped layout: **A · Programmatic activities** (with a subtotal) followed by **B · Non-programmatic** (with a subtotal), then a yellow **Grand total · YTD** row at the bottom.

This panel is read-only — it reflects what has already been approved.

### Panel 2 — Quarterly entry (below)

This is where you actually file the report. It has:

- A row of **quarter tabs**: `Q1 · May–Jul`, `Q2 · Aug–Oct`, `Q3 · Nov–Jan`, `Q4 · Feb–Apr`. Filed quarters show a green ✓.
- Three context KPI cards: **YTD Plan (through Q<sub>n−1</sub>)**, **YTD Actuals**, **YTD Variance** — these give you running context.
- The active-quarter **entry table** with one row per activity, an `Actual units` input, auto-calculated rupee columns, a `Comment` field, and a `📎` attachment column.
- A footer with `Save draft` and `Submit Q<sub>n</sub> for approval` buttons.

---

## Step-by-step: filing a quarter

### 1. Open the project

Navigate to the grant URL. Click **Budget Allocation** in the tab strip at the top of the form.

### 2. Pick the quarter

Click the quarter tab you want to file (e.g. **Q3 · Nov–Jan**). The active tab is underlined in maroon. Filed quarters carry a green ✓ next to the label and (in v2.3+) open in read-only mode if clicked.

### 3. Read the YTD context

Glance at the three KPI cards above the entry table:

- **YTD Plan (through Q<sub>n−1</sub>)** — what the plan said cumulative spend should be
- **YTD Actuals** — what was actually utilised through prior approved quarters
- **YTD Variance** — gap between plan and actuals; `+X% under` (good) or `−X% over` (red)

This tells you whether the project is on, ahead of, or behind plan *before* you file the new quarter.

### 4. Fill `Actual units` per row

For each activity row in the entry table:

- The `Plan units` column shows the **remaining quota** for that activity in this quarter (annual plan minus YTD actuals). 0 means the activity is fully delivered for the year.
- The grey caption under the activity name shows context like `ACT-0608 · YTD 0.45 / 0.5 units · +90% delivery`.
- Type the actual units delivered this quarter into `Actual units`.

The system **auto-calculates** as you type:

- `Actual ₹` = `Actual units × Unit cost`
- `Variance ₹` = `Plan ₹ − Actual ₹`
- `Var %` = `Variance ÷ Plan` (over-spend shown in red)

### 5. Add a comment (mandatory for over-spend)

If for any row you delivered more units than planned for the quarter (over-spend), the system requires a `Comment` explaining why. Empty comment on an over-spend row will block Save Draft / Submit with a clear message.

For under-spend or on-plan rows, the comment is optional but recommended.

### 6. Attach evidence (optional)

Click the `📎` icon at the right edge of any row to attach a voucher, invoice, photo, or any supporting file for that activity.

### 7. Save draft

When you're done (or before stepping away), click **Save draft** in the footer. The button is dim by default; it becomes maroon when there are unsaved changes. The footer caption switches between *“No unsaved changes”* and *“X rows changed”*.

Saved drafts persist server-side. You can come back and continue any time before submission.

### 8. Submit for approval

Once the quarter is final, click **Submit Q<sub>n</sub> for approval**. This triggers the Frappe workflow — the QUR moves to *“Pending PM Approval”* and you cannot edit it further until the PM either approves or rejects.

The PM (or SPM) will receive the QUR in their queue. After approval, the quarter shows **✓ filed** on the tab strip and its actuals roll up into the YTD summary panel at the top.

---

## Special cases

### Filing a previous quarter you forgot

Click the quarter tab. If it's not yet filed it will be editable. If a future quarter is already approved, the system may block back-filing — talk to the PM in that case.

### Over-spending an activity

You can log actual units higher than the planned quota. The system caps the rupee value at the budget for accounting purposes, but the **real units** are preserved on the row in the `custom_actual_units` field. A `Comment` is mandatory to explain the over-spend.

### A row I need to report on isn't there

The entry table is built from the approved Budget Plan. If an activity is missing, the budget itself needs to be revised. Speak to the LIC HFL PM — budget revisions are out of scope for QUR.

### I clicked Submit by mistake

Once submitted, you cannot edit. Ask the PM to **reject** the QUR; this returns it to draft state and you can resubmit after corrections.

---

## What the PM does next

The PM opens the same Grant page, sees the new submitted QUR in the workflow queue, reviews the entries against the comments and any attachments, and either:

- **Approves** — the quarter rolls into the YTD summary and is locked.
- **Rejects** — the QUR returns to draft for the partner to revise.

The PM has the full audit trail: who submitted, when, what changed, and any comments per row.

---

## FAQ

**Q. Can I enter rupee amounts directly instead of units?**
A. No. The unit-based entry was deliberately chosen so that physical delivery is what's tracked. Rupees are a derived view of `units × unit cost`. If the unit cost is wrong on a row, that's a budget plan issue — not a reporting issue.

**Q. Why is `Plan units` 0 on some rows?**
A. Because the annual plan for that activity is fully delivered in earlier quarters. You can still enter actuals (over-delivery) but you'll need a comment.

**Q. The YTD Variance shows `+3% under` — is that bad?**
A. No. `+X% under` means actual spend is below plan, which is fine for an in-progress year. Only `−X% over` (red) flags a real over-spend that needs explanation.

**Q. Can two people edit the same QUR at once?**
A. The system uses optimistic locking. If two people save in parallel, the second save will fail with a timestamp mismatch. Coordinate offline; only one person should be filing at a time.

**Q. What's the difference between Save draft and Submit?**
A. **Save draft** persists your work but the QUR stays editable and is *not* sent for approval. **Submit** locks the QUR and routes it to the PM workflow queue. Always Save draft as you go; Submit only when you're done.

---

## Need help?

If something looks off — a row missing, a unit cost wrong, the wrong quarter active — capture a screenshot and ping the LIC HFL programme manager. Most issues trace back to the budget plan that was approved at GAF stage.

For technical / platform issues, contact the mGrant support team through the **Help** menu in the top-right of the page.

---

## Appendix · Reference screenshots

Five reference views captured from `P-GAF-0545` on 26 April 2026:

1. **Top of Budget Allocation tab** — KPI cards (Total Sanctioned ₹1.29 Cr, Disbursed ₹0, Utilised ₹64.61 L, Unutilised −₹64.61 L) and start of YTD summary table.
2. **YTD summary mid-table** — programmatic subtotal ₹52.28 L on plan ₹1.04 Cr (+50%), then non-programmatic banner.
3. **Quarter tab strip + entry header** — Q1 ✓ filed · Q2 ✓ filed · **Q3 · Nov–Jan** active · Q4 · Feb–Apr; YTD Plan ₹66.65 L / Actuals ₹64.60 L / Variance +3% under.
4. **Entry table for Q3** — empty actuals across 22 programmatic + 7 non-programmatic activities, `Add a note` placeholders, `📎` icons on the right.
5. **Footer** — *No unsaved changes* caption with disabled Save draft and active maroon **Submit Q3 for approval** button.

These were captured live during this session. Drop them into this guide as inline images by exporting from the chat.
