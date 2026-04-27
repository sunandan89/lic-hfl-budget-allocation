/**
 * LIC HFL — Quarterly Utilisation Report v2.3
 * Client Script for "Grant" DocType (Form view)
 *
 * v2.3 changes (2026-04-27) — partner-clarity polish:
 * - Var % column header → "% Spent" (consistency with YTD summary)
 * - Attach button: `+` glyph → `📎` paperclip (more legible affordance)
 * - Comment column widened 200 → 260px (room for one-line explanations)
 * - Save Draft button: neutral grey when no changes, amber only when dirty
 * - YTD summary row accent strip: 3px → 4px (more visible state colour)
 * - Post-submit toast extended: "Continue with Q{n+1} →" CTA
 *
 * v2.2 (2026-04-26): YTD Budget Planning summary, over-spend cap-at-budget,
 * stacked widgets layout, view past filed quarters in read-only mode.
 *
 * v2.1 (2026-04-26 evening): Manual Save Draft + Submit, dirty tracking,
 * TimestampMismatch refetch+retry, zero-actuals confirm, single-row header.
 *
 * Schema dependencies (already deployed):
 * - Budget Planning Child custom: unit (Float), unit_cost (Currency)
 * - Quarterly Utilisation Report Child custom: custom_actual_units (Float),
 *   custom_planned_units (Float, ro), custom_unit_cost (Currency, ro)
 */

frappe.ui.form.on('Grant', {
    refresh(frm) {
        if (!frm.doc.name || frm.doc.__islocal) return;
        qurv2_setup(frm);
    }
});

const QURV2_NS = '__qur_v2__';

function qurv2_setup(frm) {
    if (!frm.fields_dict.quarterly_utilisation) return;
    const wrapper = frm.fields_dict.quarterly_utilisation.$wrapper;
    if (!wrapper || !wrapper.length) return;

    if (!wrapper.data('qurv2-patched')) {
        const $w = wrapper;
        const origHtml = $w.html.bind($w);
        $w.html = function () { return $w; };
        $w.empty = function () { return $w; };
        $w.data('qurv2-patched', true);
        $w.data('qurv2-orig-html', origHtml);
    }

    qurv2_render(frm);
    qurv2_bindGlobalKeys(frm);
    qurv2_setupBudgetPlanningSummary(frm);
}

// --- Budget Planning YTD summary -----------------------------------------
// Replaces the default mGrant SVADatatable on the `budget` HTML field with
// a 29-row YTD summary (one row per BPU, aggregated across approved QURs).

function qurv2_setupBudgetPlanningSummary(frm) {
    let tries = 0;
    const interval = setInterval(() => {
        tries += 1;
        const fld = frm.fields_dict.budget;
        const $section = fld && fld.$wrapper;
        if ($section && $section.length && $section.find('table, .datatable, .frappe-control').length) {
            clearInterval(interval);
            qurv2_patchBudgetWrapper($section);
            if (frm[QURV2_NS] && frm[QURV2_NS].ctx) qurv2_renderBudgetPlanningSummary(frm);
        } else if (tries >= 30) {
            clearInterval(interval);
        }
    }, 500);
}

function qurv2_patchBudgetWrapper($w) {
    if ($w.data('qurv2-bp-patched')) return;
    const origHtml = $w.html.bind($w);
    $w.html = function () { return $w; };
    $w.empty = function () { return $w; };
    $w.data('qurv2-bp-patched', true);
    $w.data('qurv2-bp-orig-html', origHtml);
}

function qurv2_renderBudgetPlanningSummary(frm) {
    const fld = frm.fields_dict.budget;
    if (!fld || !fld.$wrapper) return;
    const $w = fld.$wrapper;
    const orig = $w.data('qurv2-bp-orig-html');
    if (!orig) return;
    const ctx = frm[QURV2_NS] && frm[QURV2_NS].ctx;
    if (!ctx) return;
    if ($w[0]) $w[0].innerHTML = '';
    orig(qurv2_buildBudgetSummaryHtml(ctx));
}

function qurv2_buildBudgetSummaryHtml(ctx) {
    const css = `
<style id="qurv2-bp-style">
.qurv2-bp-root { font-family: inherit; padding: 0 0 1rem; }
.qurv2-bp-header { display: flex; align-items: baseline; padding: 0 0 8px; gap: 12px; }
.qurv2-bp-title { font-size: 14px; font-weight: 500; color: #2C2C2A; }
.qurv2-bp-sub { font-size: 11px; color: #888; }
.qurv2-bp-table { width: 100%; border-collapse: collapse; font-size: 12px; table-layout: fixed; }
.qurv2-bp-table th { background: #EFEEEA; padding: 8px 10px; font-weight: 500; color: #555; border-bottom: 0.5px solid #C9C7C0; text-align: right; white-space: nowrap; }
.qurv2-bp-table th.qurv2-bp-th-text { text-align: left; }
.qurv2-bp-table td { padding: 8px 10px; border-bottom: 0.5px solid #EEEEEC; vertical-align: middle; }
.qurv2-bp-table td:first-child { border-left: 4px solid transparent; }
.qurv2-bp-table td.qurv2-bp-num { text-align: right; font-variant-numeric: tabular-nums; }
.qurv2-bp-row.qurv2-bp-not-started td:first-child { border-left-color: #DEDCD6; }
.qurv2-bp-row.qurv2-bp-on-track td:first-child { border-left-color: #3B6D11; }
.qurv2-bp-row.qurv2-bp-behind td:first-child { border-left-color: #BA7517; }
.qurv2-bp-row.qurv2-bp-at-plan td:first-child { border-left-color: #3B6D11; }
.qurv2-bp-row.qurv2-bp-over td { background: #FCEBEB; }
.qurv2-bp-row.qurv2-bp-over td:first-child { border-left-color: #A32D2D; }
.qurv2-bp-banner td { background: #333333; color: #FFFFFF; padding: 7px 12px; font-size: 11px; font-weight: 500; letter-spacing: 0.5px; text-transform: uppercase; border: none; }
.qurv2-bp-subtotal td { background: #BCBDC0; font-weight: 500; color: #2C2C2A; padding: 8px 10px; }
.qurv2-bp-grandtotal td { background: #FFCB05; font-weight: 500; color: #4A1B0C; padding: 10px; font-size: 13px; }
.qurv2-bp-pct-cell { white-space: nowrap; }
.qurv2-bp-pct-bar { display: inline-block; width: 56px; height: 4px; background: #E6E6E6; border-radius: 2px; vertical-align: middle; margin-right: 6px; overflow: hidden; }
.qurv2-bp-pct-fill { display: block; height: 100%; background: #3B6D11; border-radius: 2px; }
.qurv2-bp-pct-fill.amber { background: #BA7517; }
.qurv2-bp-pct-fill.red { background: #A32D2D; }
.qurv2-bp-row-act { color: #2C2C2A; font-weight: 400; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.qurv2-bp-row-sub { font-size: 10.5px; color: #999; margin-top: 2px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.qurv2-bp-row.qurv2-bp-not-started .qurv2-bp-row-act { color: #888; }
.qurv2-bp-over { color: #A32D2D; }
</style>`;

    function annualFor(bpu) {
        let planUnits = 0, planAmt = 0;
        for (const c of (bpu.planning_table || [])) {
            planUnits += +c.unit || 0;
            planAmt += +c.planned_amount || 0;
        }
        return { planUnits, planAmt };
    }
    function ytdFor(bpu) {
        let actualUnits = 0, actualAmt = 0, lastComment = '', lastQ = null;
        for (const q of ctx.quarters) {
            const qur = ctx.qStates[q.n].qur;
            if (!qur || qur.workflow_state !== 'Approved') continue;
            const ch = (qur.quarterly_utilisation || []).find(c => c.budget_plan === bpu.name);
            if (ch) {
                actualUnits += +ch.custom_actual_units || 0;
                actualAmt += +ch.actuals || 0;
                if (ch.comment) lastComment = ch.comment;
                lastQ = q.n;
            }
        }
        return { actualUnits, actualAmt, lastComment, lastQ };
    }
    function rowState(annual, ytd) {
        if (ytd.actualAmt === 0 && ytd.actualUnits === 0) return 'not-started';
        const sp = annual.planAmt > 0 ? (ytd.actualAmt / annual.planAmt * 100) : 0;
        const dp = annual.planUnits > 0 ? (ytd.actualUnits / annual.planUnits * 100) : 0;
        if (sp > 100) return 'over';
        if (sp >= 96) return 'at-plan';
        if (dp < 60 && sp >= 1) return 'behind';
        return 'on-track';
    }
    function pctBarHtml(spentPct) {
        const fill = Math.min(spentPct, 100);
        let cls = '';
        if (spentPct > 100) cls = 'red';
        else if (spentPct < 60 && spentPct > 0) cls = 'amber';
        return `<span class="qurv2-bp-pct-bar"><span class="qurv2-bp-pct-fill ${cls}" style="width: ${fill.toFixed(0)}%;"></span></span>${qurv2_pct(spentPct)}`;
    }
    function buildRow(bpu, sno) {
        const a = annualFor(bpu), y = ytdFor(bpu);
        const variance = a.planAmt - y.actualAmt;
        const sp = a.planAmt > 0 ? (y.actualAmt / a.planAmt * 100) : 0;
        const dp = a.planUnits > 0 ? (y.actualUnits / a.planUnits * 100) : 0;
        const state = rowState(a, y);
        const uomLabel = bpu.custom_unit_of_measurement_title || __('Nos');
        const itemTitle = frappe.utils.escape_html((bpu.item_name || bpu.particulars || '—') + ' · ' + uomLabel);
        const taskTitle = bpu.custom_task_details ? frappe.utils.escape_html(bpu.custom_task_details) : '';
        const lastQNote = y.lastQ ? __('Last reported Q{0}', [y.lastQ]) : __('Not reported yet');
        const tooltip = y.lastComment ? frappe.utils.escape_html(y.lastComment) : __('No comments yet');
        return `<tr class="qurv2-bp-row qurv2-bp-${state}" title="${tooltip}">
            <td style="text-align: center; color: #AAA;">${sno}</td>
            <td>
                <div class="qurv2-bp-row-act" title="${itemTitle}">${itemTitle}</div>
                <div class="qurv2-bp-row-sub" title="${taskTitle}">${bpu.custom_activity ? frappe.utils.escape_html(bpu.custom_activity) + ' · ' : ''}${lastQNote}</div>
            </td>
            <td class="qurv2-bp-num">${qurv2_num(a.planUnits)}</td>
            <td class="qurv2-bp-num">${qurv2_num(y.actualUnits)}</td>
            <td class="qurv2-bp-num">${qurv2_pct(dp)}</td>
            <td class="qurv2-bp-num">${qurv2_inr(a.planAmt)}</td>
            <td class="qurv2-bp-num">${qurv2_inr(y.actualAmt)}</td>
            <td class="qurv2-bp-num ${variance < 0 ? 'qurv2-bp-over' : ''}">${qurv2_inr(variance)}</td>
            <td class="qurv2-bp-num qurv2-bp-pct-cell">${pctBarHtml(sp)}</td>
        </tr>`;
    }
    function subtotalRow(bpus, label) {
        let pPlan = 0, pActual = 0;
        for (const b of bpus) {
            const a = annualFor(b), y = ytdFor(b);
            pPlan += a.planAmt;
            pActual += y.actualAmt;
        }
        const v = pPlan - pActual;
        const sp = pPlan > 0 ? (pActual / pPlan * 100) : 0;
        return {
            html: `<tr class="qurv2-bp-subtotal"><td></td><td colspan="2">${label}</td><td></td><td></td><td class="qurv2-bp-num">${qurv2_inr(pPlan)}</td><td class="qurv2-bp-num">${qurv2_inr(pActual)}</td><td class="qurv2-bp-num">${qurv2_inr(v)}</td><td class="qurv2-bp-num">${qurv2_pct(sp)}</td></tr>`,
            pPlan, pActual
        };
    }

    const prog = ctx.bpuFullList.filter(b => !!b.custom_activity);
    const np = ctx.bpuFullList.filter(b => !b.custom_activity);
    const progRows = prog.map((b, i) => buildRow(b, i + 1));
    const npRows = np.map((b, i) => buildRow(b, prog.length + i + 1));
    const progSub = subtotalRow(prog, __('Subtotal — Programmatic'));
    const npSub = subtotalRow(np, __('Subtotal — Non-programmatic'));
    const gtPlan = progSub.pPlan + npSub.pPlan;
    const gtActual = progSub.pActual + npSub.pActual;
    const gtVar = gtPlan - gtActual;
    const gtSp = gtPlan > 0 ? (gtActual / gtPlan * 100) : 0;

    return `${css}
<div class="qurv2-bp-root">
    <div class="qurv2-bp-header">
        <div class="qurv2-bp-title">${__('Budget · Year-to-date summary')}</div>
        <div class="qurv2-bp-sub">${__('Aggregated across approved quarterly reports')}</div>
    </div>
    <table class="qurv2-bp-table">
        <colgroup>
            <col style="width: 36px;"><col style="width: 280px;">
            <col style="width: 80px;"><col style="width: 80px;"><col style="width: 80px;">
            <col style="width: 110px;"><col style="width: 110px;"><col style="width: 110px;">
            <col style="width: 130px;">
        </colgroup>
        <thead>
            <tr>
                <th style="text-align: center;">#</th>
                <th class="qurv2-bp-th-text">${__('Activity · UoM')}</th>
                <th>${__('Plan units')}</th>
                <th>${__('Actual units')}</th>
                <th>${__('% Delivered')}</th>
                <th>${__('Plan ₹')}</th>
                <th>${__('Actual ₹')}</th>
                <th>${__('Variance ₹')}</th>
                <th>${__('% Spent')}</th>
            </tr>
        </thead>
        <tbody>
            <tr class="qurv2-bp-banner"><td colspan="9">${__('A · Programmatic activities')}</td></tr>
            ${progRows.join('')}
            ${progSub.html}
            <tr class="qurv2-bp-banner"><td colspan="9">${__('B · Non-programmatic')}</td></tr>
            ${npRows.join('')}
            ${npSub.html}
            <tr class="qurv2-bp-grandtotal"><td></td><td colspan="2">${__('Grand total · YTD')}</td><td></td><td></td><td class="qurv2-bp-num">${qurv2_inr(gtPlan)}</td><td class="qurv2-bp-num">${qurv2_inr(gtActual)}</td><td class="qurv2-bp-num">${qurv2_inr(gtVar)}</td><td class="qurv2-bp-num">${qurv2_pct(gtSp)}</td></tr>
        </tbody>
    </table>
</div>`;
}

async function qurv2_render(frm) {
    const wrapper = frm.fields_dict.quarterly_utilisation.$wrapper;
    const orig = wrapper.data('qurv2-orig-html');
    if (!frm[QURV2_NS]) frm[QURV2_NS] = { dirty: false };

    orig(`<div class="qurv2-root"><div class="qurv2-loading" style="padding: 24px; text-align: center; color: #888;">${__('Loading utilisation data...')}</div></div>`);

    try {
        const ctx = await qurv2_loadContext(frm);
        frm[QURV2_NS].ctx = ctx;
        frm[QURV2_NS].dirty = false;
        qurv2_paint(frm);
        qurv2_renderBudgetPlanningSummary(frm);
    } catch (e) {
        console.error('[QURv2] load failed', e);
        const msg = (e && e.message) || String(e);
        orig(`<div class="qurv2-root"><div style="padding: 16px; color: #A32D2D; background: #FCEBEB; border-radius: 6px;">${__('Could not load utilisation data')}: ${frappe.utils.escape_html(msg)}</div></div>`);
    }
}

async function qurv2_loadContext(frm) {
    const grantName = frm.doc.name;

    const bpus = await frappe.db.get_list('Budget Plan and Utilisation', {
        filters: { grant: grantName },
        fields: ['name'],
        limit: 0,
        order_by: 'idx asc'
    });
    const bpuFullList = [];
    for (const b of bpus) {
        const full = await frappe.db.get_doc('Budget Plan and Utilisation', b.name);
        bpuFullList.push(full);
    }

    const qurs = await frappe.db.get_list('Quarterly Utilisation Report', {
        filters: { grant: grantName },
        fields: ['name'],
        limit: 0,
        order_by: 'modified desc'
    });
    const qurFullByQ = {};
    const dupQuarters = [];
    for (const q of qurs) {
        const full = await frappe.db.get_doc('Quarterly Utilisation Report', q.name);
        if (qurFullByQ[full.quarter]) {
            dupQuarters.push({ kept: qurFullByQ[full.quarter].name, dropped: full.name, q: full.quarter });
        } else {
            qurFullByQ[full.quarter] = full;
        }
    }
    if (dupQuarters.length) {
        dupQuarters.forEach(d => {
            frappe.show_alert({
                message: __('Found duplicate draft for Q{0}; keeping the most recent ({1}). Older draft {2} ignored.', [d.q, d.kept, d.dropped]),
                indicator: 'orange'
            }, 8);
        });
    }

    const quarters = (frm.doc.quaterly_project || []).map(q => ({
        n: parseInt(String(q.quarter).replace('Q', ''), 10) || q.quarter_sequence,
        timespan: q.timespan || ('Q' + (q.quarter_sequence || q.quarter)),
        start: q.start_date, end: q.end_date,
        year: q.year_sequence || 1
    })).sort((a, b) => a.n - b.n);

    const qStates = {};
    for (const q of quarters) {
        const qur = qurFullByQ[q.n];
        let state = 'empty';
        if (qur) {
            if (qur.workflow_state === 'Approved') state = 'filed';
            else if (qur.workflow_state === 'Sent Back To Partner') state = 'sent_back';
            else if (qur.docstatus === 1) state = 'filed';
            else if (['Pending at PM', 'Pending at SPM', 'Pending at PL'].includes(qur.workflow_state)) state = 'review';
            else state = 'drafting';
        }
        qStates[q.n] = { state, qur };
    }

    let currentQ = quarters[0] && quarters[0].n;
    for (const q of quarters) {
        const s = qStates[q.n].state;
        if (s !== 'filed') { currentQ = q.n; break; }
    }

    return {
        grantName, bpuFullList, qurFullByQ, qStates, quarters, currentQ,
        annualBudget: frm.doc.total_planned_budget || 0,
        ngo: frm.doc.ngo, donor: frm.doc.donor
    };
}

function qurv2_paint(frm) {
    const wrapper = frm.fields_dict.quarterly_utilisation.$wrapper;
    const orig = wrapper.data('qurv2-orig-html');
    const ctx = frm[QURV2_NS].ctx;
    if (!ctx) return;

    // Hard-clear any stale content (defensive against multiple-render races)
    if (wrapper[0]) wrapper[0].innerHTML = '';
    orig(qurv2_buildHtml(ctx));
    qurv2_bindEvents(frm);
    qurv2_recalcAll(frm);
    qurv2_updateDirtyChrome(frm);
}

function qurv2_buildHtml(ctx) {
    const css = `
<style id="qurv2-style">
.qurv2-root { font-family: inherit; }
.qurv2-tabs { display: flex; gap: 0; border-bottom: 1px solid #E6E6E6; padding: 0 4px; }
.qurv2-tab { padding: 10px 18px; font-size: 13px; cursor: pointer; border-bottom: 2px solid transparent; color: #888; user-select: none; }
.qurv2-tab.active { color: #8B1A1A; font-weight: 500; border-bottom-color: #8B1A1A; }
.qurv2-tab .qurv2-tab-status { font-size: 11px; margin-left: 6px; color: #999; }
.qurv2-tab.filed .qurv2-tab-status { color: #3B6D11; }
.qurv2-tab.drafting .qurv2-tab-status, .qurv2-tab.review .qurv2-tab-status { color: #8B1A1A; }
.qurv2-tab.locked { cursor: default; opacity: 0.6; }
.qurv2-kpi-band { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; padding: 12px 4px; }
.qurv2-kpi { background: #F7F7F4; border-radius: 6px; padding: 10px 14px; }
.qurv2-kpi-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px; color: #777; }
.qurv2-kpi-value { font-size: 17px; font-weight: 500; margin-top: 4px; color: #2C2C2A; }
.qurv2-toolbar { padding: 8px 4px; font-size: 11px; color: #999; display: flex; align-items: center; gap: 12px; }
.qurv2-toolbar .qurv2-dirty { margin-left: auto; }
.qurv2-toolbar .qurv2-dirty.has-unsaved { color: #854F0B; font-weight: 500; }
.qurv2-table-wrap { overflow: auto; max-height: 65vh; border: 0.5px solid #E6E6E6; border-radius: 6px; }
.qurv2-table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 12px; }
.qurv2-table th { position: sticky; top: 0; background: #EFEEEA; padding: 8px 10px; font-weight: 500; color: #555; border-bottom: 0.5px solid #C9C7C0; text-align: right; white-space: nowrap; z-index: 2; }
.qurv2-table th.qurv2-th-text { text-align: left; }
.qurv2-table th.qurv2-th-input { background: #FAEEDA; color: #854F0B; }
.qurv2-table th.qurv2-stickyL { left: 0; z-index: 3; }
.qurv2-table th.qurv2-stickyL2 { z-index: 3; }
.qurv2-table th.qurv2-stickyL3 { z-index: 3; border-right: 0.5px solid #C9C7C0; }
.qurv2-table td { padding: 7px 10px; border-bottom: 0.5px solid #EEEEEC; vertical-align: top; }
.qurv2-table td.qurv2-num { text-align: right; font-variant-numeric: tabular-nums; }
.qurv2-table td.qurv2-stickyL { position: sticky; left: 0; background: #FFFFFF; z-index: 1; }
.qurv2-table td.qurv2-stickyL2 { position: sticky; background: #FFFFFF; z-index: 1; }
.qurv2-table td.qurv2-stickyL3 { position: sticky; background: #FFFFFF; z-index: 1; border-right: 0.5px solid #DEDCD6; }
.qurv2-row-srno { color: #AAA; }
.qurv2-row-act { font-weight: 400; color: #2C2C2A; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.qurv2-row-sub { font-size: 10.5px; color: #999; margin-top: 2px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.qurv2-banner td { background: #333333; color: #FFFFFF; padding: 7px 12px; font-size: 11px; font-weight: 500; letter-spacing: 0.5px; text-transform: uppercase; border: none; }
.qurv2-subtotal td { background: #BCBDC0; font-weight: 500; color: #2C2C2A; padding: 8px 10px; }
.qurv2-grandtotal td { background: #FFCB05; font-weight: 500; color: #4A1B0C; padding: 10px; font-size: 13px; }
.qurv2-input { width: 100%; height: 28px; padding: 0 6px; border: 0.5px solid #BA7517; border-radius: 4px; text-align: right; font-size: 12.5px; font-weight: 500; background: #FFFFFF; font-family: inherit; }
.qurv2-input:focus { outline: 1.5px solid #8B1A1A; outline-offset: -1px; }
.qurv2-input.readonly { border-color: transparent; background: transparent; color: #888; cursor: default; }
.qurv2-cell-input-bg { background: #FFFEF7; }
.qurv2-comment-input { width: 100%; height: 26px; padding: 0 8px; border: 0.5px solid #DEDCD6; border-radius: 4px; font-size: 12px; font-family: inherit; }
.qurv2-comment-input:focus { outline: 1.5px solid #8B1A1A; outline-offset: -1px; border-color: transparent; }
.qurv2-attach-btn { width: 24px; height: 24px; border: 0.5px dashed #C9C7C0; border-radius: 4px; background: transparent; color: #999; cursor: pointer; font-size: 14px; line-height: 1; display: inline-flex; align-items: center; justify-content: center; }
.qurv2-attach-btn.has-file { border-style: solid; color: #185FA5; border-color: #B5D4F4; background: #E6F1FB; }
.qurv2-over { color: #A32D2D; }
.qurv2-row.qurv2-overspent td { background: #FCEBEB; }
.qurv2-row.qurv2-overspent td.qurv2-stickyL,
.qurv2-row.qurv2-overspent td.qurv2-stickyL2,
.qurv2-row.qurv2-overspent td.qurv2-stickyL3 { background: #FCEBEB; }
.qurv2-row.qurv2-overspent input[data-cell="comment"]:placeholder-shown { border-color: #E24B4A; background: #FFF; }
.qurv2-actions-bar { display: flex; gap: 10px; justify-content: flex-end; padding: 12px 4px; align-items: center; }
.qurv2-actions-bar .qurv2-status { margin-right: auto; font-size: 11px; color: #999; }
.qurv2-actions-bar .qurv2-status.has-unsaved { color: #854F0B; font-weight: 500; }
.qurv2-btn-primary { background: #8B1A1A; color: #FFFFFF; border: none; border-radius: 6px; padding: 8px 16px; font-size: 13px; font-weight: 500; cursor: pointer; font-family: inherit; }
.qurv2-btn-primary:hover { background: #6F1414; }
.qurv2-btn-primary:disabled { background: #BCBDC0; cursor: not-allowed; }
.qurv2-btn-ghost { background: transparent; color: #777; border: 0.5px solid #C9C7C0; border-radius: 6px; padding: 8px 14px; font-size: 13px; cursor: pointer; font-family: inherit; font-weight: 500; transition: all 0.15s; }
.qurv2-btn-ghost:hover { background: #F4F4F0; color: #2C2C2A; border-color: #888; }
.qurv2-btn-ghost.has-changes { color: #854F0B; border-color: #BA7517; }
.qurv2-btn-ghost.has-changes:hover { background: #FAEEDA; }
.qurv2-btn-ghost:disabled { opacity: 0.5; cursor: not-allowed; }
</style>`;
    return `${css}
<div class="qurv2-root">
  ${qurv2_buildTabs(ctx)}
  ${qurv2_buildKpis(ctx)}
  <div class="qurv2-toolbar">
    <span>${__('Enter')} <strong>${__('actual units')}</strong> ${__('— financial values auto-calculate. Click Save draft to persist; Submit applies the workflow.')}</span>
    <span class="qurv2-dirty"></span>
  </div>
  ${qurv2_buildTable(ctx)}
  ${qurv2_buildActions(ctx)}
</div>`;
}

function qurv2_buildTabs(ctx) {
    const labelMap = { empty: '', drafting: __('drafting'), filed: __('✓ filed'), sent_back: __('sent back'), review: __('under review') };
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const tabs = ctx.quarters.map(q => {
        const st = ctx.qStates[q.n];
        const cls = ['qurv2-tab'];
        if (q.n === ctx.currentQ) cls.push('active');
        if (st.state === 'filed') cls.push('filed');
        if (st.state === 'drafting') cls.push('drafting');
        if (st.state === 'review') cls.push('review');
        const sd = q.start ? new Date(q.start) : null;
        const ed = q.end ? new Date(q.end) : null;
        const sm = sd ? monthNames[sd.getMonth()] : '';
        const em = ed ? monthNames[ed.getMonth()] : '';
        return `<div class="${cls.join(' ')}" data-q="${q.n}">Q${q.n}${sm ? ' · ' + sm + '–' + em : ''}<span class="qurv2-tab-status">${labelMap[st.state]}</span></div>`;
    });
    return `<div class="qurv2-tabs">${tabs.join('')}</div>`;
}

function qurv2_buildKpis(ctx) {
    const priorQs = ctx.quarters.filter(q => q.n < ctx.currentQ);
    let ytdPlan = 0, ytdActual = 0;
    for (const b of ctx.bpuFullList) {
        for (const c of (b.planning_table || [])) {
            const cn = qurv2_quarterNum(c);
            if (priorQs.find(p => p.n === cn)) ytdPlan += +c.planned_amount || 0;
        }
    }
    for (const q of priorQs) {
        const qur = ctx.qStates[q.n].qur;
        if (qur && qur.workflow_state === 'Approved') {
            for (const ch of (qur.quarterly_utilisation || [])) ytdActual += +ch.actuals || 0;
        }
    }
    const ytdVar = ytdPlan - ytdActual;
    const ytdVarPct = ytdPlan > 0 ? (ytdVar / ytdPlan * 100) : 0;
    const through = priorQs.length ? __('through Q{0}', [priorQs[priorQs.length - 1].n]) : __('no prior quarter');
    const varColor = ytdVar < 0 ? '#A32D2D' : '#2C2C2A';
    const overUnder = ytdVar < 0 ? __('over') : __('under');
    return `<div class="qurv2-kpi-band">
        <div class="qurv2-kpi"><div class="qurv2-kpi-label">${__('YTD plan')} (${through})</div>
            <div class="qurv2-kpi-value">${qurv2_inr(ytdPlan)}</div></div>
        <div class="qurv2-kpi"><div class="qurv2-kpi-label">${__('YTD actuals')} (${through})</div>
            <div class="qurv2-kpi-value">${qurv2_inr(ytdActual)}</div></div>
        <div class="qurv2-kpi"><div class="qurv2-kpi-label">${__('YTD variance')}</div>
            <div class="qurv2-kpi-value" style="color: ${varColor};">${qurv2_inr(ytdVar)}
                <span style="font-size: 11px; font-weight: 400; color: #777;">${qurv2_pct(ytdVarPct)} ${ytdPlan ? overUnder : ''}</span>
            </div></div>
    </div>`;
}

function qurv2_buildTable(ctx) {
    const cur = ctx.currentQ;
    const curQur = ctx.qStates[cur] && ctx.qStates[cur].qur;
    const isReadOnly = curQur && (curQur.workflow_state === 'Approved' || curQur.docstatus === 1 ||
                                   ['Pending at PM', 'Pending at SPM', 'Pending at PL'].includes(curQur.workflow_state));
    const childByBpu = {};
    if (curQur) for (const ch of (curQur.quarterly_utilisation || [])) childByBpu[ch.budget_plan] = ch;

    const colgroup = `<colgroup>
        <col style="width: 36px;"><col style="width: 240px;"><col style="width: 70px;">
        <col style="width: 90px;">
        <col style="width: 110px;"><col style="width: 110px;"><col style="width: 110px;"><col style="width: 90px;">
        <col style="width: 260px;"><col style="width: 50px;">
    </colgroup>`;

    const thead = `<thead><tr>
        <th class="qurv2-stickyL" style="left: 0;">#</th>
        <th class="qurv2-th-text qurv2-stickyL2" style="left: 36px;">${__('Activity · UoM')}</th>
        <th class="qurv2-stickyL3" style="left: 276px;">${__('Plan units')}</th>
        <th class="qurv2-th-input">${__('Actual units')}</th>
        <th>${__('Plan ₹')}</th>
        <th>${__('Actual ₹')}</th>
        <th>${__('Variance ₹')}</th>
        <th>${__('% Spent')}</th>
        <th class="qurv2-th-text">${__('Comment')}</th>
        <th>📎</th>
    </tr></thead>`;

    const prog = ctx.bpuFullList.filter(b => !!b.custom_activity);
    const np = ctx.bpuFullList.filter(b => !b.custom_activity);
    const progRows = prog.map((b, i) => qurv2_buildRow(b, i + 1, ctx, cur, childByBpu, isReadOnly));
    const npRows = np.map((b, i) => qurv2_buildRow(b, prog.length + i + 1, ctx, cur, childByBpu, isReadOnly));

    const tbody = `<tbody>
        <tr class="qurv2-banner"><td colspan="10">${__('A · Programmatic activities')}</td></tr>
        ${progRows.join('')}
        <tr class="qurv2-subtotal" data-subtotal="prog">
            <td class="qurv2-stickyL"></td>
            <td class="qurv2-stickyL2" style="left: 36px;" colspan="2">${__('Subtotal — Programmatic')}</td>
            <td></td>
            <td class="qurv2-num" data-st="prog-plan">0</td>
            <td class="qurv2-num" data-st="prog-actual">0</td>
            <td class="qurv2-num" data-st="prog-var">0</td>
            <td class="qurv2-num" data-st="prog-pct">0%</td>
            <td></td><td></td>
        </tr>
        <tr class="qurv2-banner"><td colspan="10">${__('B · Non-programmatic (HR / Admin / NGO management)')}</td></tr>
        ${npRows.join('')}
        <tr class="qurv2-subtotal" data-subtotal="np">
            <td class="qurv2-stickyL"></td>
            <td class="qurv2-stickyL2" style="left: 36px;" colspan="2">${__('Subtotal — Non-programmatic')}</td>
            <td></td>
            <td class="qurv2-num" data-st="np-plan">0</td>
            <td class="qurv2-num" data-st="np-actual">0</td>
            <td class="qurv2-num" data-st="np-var">0</td>
            <td class="qurv2-num" data-st="np-pct">0%</td>
            <td></td><td></td>
        </tr>
        <tr class="qurv2-grandtotal" data-grandtotal="1">
            <td class="qurv2-stickyL"></td>
            <td class="qurv2-stickyL2" style="left: 36px;" colspan="2">${__('Grand total · Q{0}', [cur])}</td>
            <td></td>
            <td class="qurv2-num" data-st="gt-plan">0</td>
            <td class="qurv2-num" data-st="gt-actual">0</td>
            <td class="qurv2-num" data-st="gt-var">0</td>
            <td class="qurv2-num" data-st="gt-pct">0%</td>
            <td></td><td></td>
        </tr>
    </tbody>`;

    return `<div class="qurv2-table-wrap">
        <table class="qurv2-table">${colgroup}${thead}${tbody}</table>
    </div>`;
}

function qurv2_buildRow(bpu, sno, ctx, currentQ, childByBpu, isReadOnly) {
    const ptCur = (bpu.planning_table || []).find(c => qurv2_quarterNum(c) === currentQ);
    const planUnits = ptCur ? +ptCur.unit || 0 : 0;
    const unitCost = ptCur ? +ptCur.unit_cost || 0 : 0;
    const planAmt = ptCur ? +ptCur.planned_amount || (planUnits * unitCost) : 0;

    const priorQs = ctx.quarters.filter(q => q.n < currentQ).map(q => q.n);
    let ytdPlanUnits = 0, ytdActualUnits = 0;
    for (const c of (bpu.planning_table || [])) {
        if (priorQs.includes(qurv2_quarterNum(c))) ytdPlanUnits += +c.unit || 0;
    }
    for (const q of ctx.quarters) {
        if (q.n >= currentQ) continue;
        const qur = ctx.qStates[q.n].qur;
        if (!qur || qur.workflow_state !== 'Approved') continue;
        const ch = (qur.quarterly_utilisation || []).find(c => c.budget_plan === bpu.name);
        if (ch) ytdActualUnits += +ch.custom_actual_units || 0;
    }
    const deliveryPct = ytdPlanUnits > 0 ? (ytdActualUnits / ytdPlanUnits * 100) : 0;

    const draft = childByBpu[bpu.name];
    const draftActualUnits = draft ? (+draft.custom_actual_units || 0) : 0;
    const draftComment = draft ? (draft.comment || '') : '';
    const draftSupport = draft ? (draft.support_documents || '') : '';

    const uomLabel = bpu.custom_unit_of_measurement_title || __('Nos');
    const actId = bpu.custom_activity || '';
    const taskBits = [];
    if (actId) taskBits.push(actId);
    if (priorQs.length && ytdPlanUnits > 0) {
        taskBits.push(__('YTD {0} / {1} units · {2} delivery', [qurv2_num(ytdActualUnits), qurv2_num(ytdPlanUnits), qurv2_pct(deliveryPct)]));
    }

    const inputAttrs = isReadOnly ? 'readonly' : '';
    const inputCls = isReadOnly ? 'qurv2-input readonly' : 'qurv2-input';
    const cellInputBg = isReadOnly ? '' : 'qurv2-cell-input-bg';
    const taskTitle = bpu.custom_task_details ? frappe.utils.escape_html(bpu.custom_task_details) : '';
    const itemTitle = frappe.utils.escape_html((bpu.item_name || bpu.particulars || '—') + ' · ' + uomLabel);

    return `<tr class="qurv2-row" data-bpu="${bpu.name}" data-plan-units="${planUnits}" data-unit-cost="${unitCost}" data-plan-amt="${planAmt}" data-support="${frappe.utils.escape_html(draftSupport)}">
        <td class="qurv2-stickyL qurv2-row-srno" style="text-align: center;">${sno}</td>
        <td class="qurv2-stickyL2" style="left: 36px;">
            <div class="qurv2-row-act" title="${itemTitle}">${itemTitle}</div>
            <div class="qurv2-row-sub" title="${taskTitle}">${frappe.utils.escape_html(taskBits.join(' · '))}</div>
        </td>
        <td class="qurv2-stickyL3 qurv2-num" style="left: 276px; color: #777;">${qurv2_num(planUnits)}</td>
        <td class="${cellInputBg}" style="padding: 4px 8px;">
            <input type="text" class="${inputCls}" data-cell="actual-units-input" value="${draftActualUnits || ''}" placeholder="0" ${inputAttrs}>
        </td>
        <td class="qurv2-num" data-cell="plan-amt" style="color: #777;">${qurv2_inr(planAmt)}</td>
        <td class="qurv2-num" data-cell="actual-amt">${qurv2_inr(draftActualUnits * unitCost)}</td>
        <td class="qurv2-num" data-cell="variance">${qurv2_inr(planAmt - draftActualUnits * unitCost)}</td>
        <td class="qurv2-num" data-cell="var-pct">${qurv2_pct_spent(planAmt > 0 ? (draftActualUnits * unitCost / planAmt * 100) : 0)}</td>
        <td style="padding: 4px 8px;">
            <input type="text" class="qurv2-comment-input" data-cell="comment" value="${frappe.utils.escape_html(draftComment)}" placeholder="${__('Add a note')}" ${inputAttrs}>
        </td>
        <td style="text-align: center;">
            <button class="qurv2-attach-btn ${draftSupport ? 'has-file' : ''}" data-cell="attach" ${isReadOnly ? 'disabled' : ''} title="${draftSupport ? __('Attached') : __('Attach evidence')}">${draftSupport ? '✓' : '📎'}</button>
        </td>
    </tr>`;
}

function qurv2_buildActions(ctx) {
    const cur = ctx.currentQ;
    const curState = ctx.qStates[cur] && ctx.qStates[cur].state;
    const curQur = ctx.qStates[cur] && ctx.qStates[cur].qur;
    let primaryLabel = __('Submit Q{0} for approval', [cur]);
    let primaryDisabled = false;
    let saveDisabled = false;
    if (curState === 'filed') {
        primaryLabel = __('Q{0} approved', [cur]);
        primaryDisabled = true;
        saveDisabled = true;
    } else if (curState === 'review') {
        primaryLabel = __('Q{0} under review', [cur]);
        primaryDisabled = true;
        saveDisabled = true;
    }
    return `<div class="qurv2-actions-bar">
        <span class="qurv2-status">${__('No unsaved changes')}</span>
        <button class="qurv2-btn-ghost qurv2-save-draft" ${saveDisabled ? 'disabled' : ''}>${__('Save draft')}</button>
        <button class="qurv2-btn-primary qurv2-submit" ${primaryDisabled ? 'disabled' : ''}>${primaryLabel}</button>
    </div>`;
}

// --- Events + computation ------------------------------------------------

function qurv2_bindEvents(frm) {
    const wrapper = frm.fields_dict.quarterly_utilisation.$wrapper;

    wrapper.find('.qurv2-tab').off('click.qurv2').on('click.qurv2', function () {
        const $t = $(this);
        const q = parseInt($t.data('q'), 10);
        if (!q || q === frm[QURV2_NS].ctx.currentQ) return;
        if (frm[QURV2_NS].dirty) {
            frappe.confirm(
                __('You have unsaved changes for Q{0}. Discard and switch to Q{1}?', [frm[QURV2_NS].ctx.currentQ, q]),
                () => { frm[QURV2_NS].dirty = false; frm[QURV2_NS].ctx.currentQ = q; qurv2_paint(frm); }
            );
        } else {
            frm[QURV2_NS].ctx.currentQ = q;
            qurv2_paint(frm);
        }
    });

    wrapper.find('input[data-cell="actual-units-input"], input[data-cell="comment"]').off('input.qurv2')
        .on('input.qurv2', function () {
            qurv2_markDirty(frm);
            if ($(this).attr('data-cell') === 'actual-units-input') {
                qurv2_recalcRow($(this).closest('.qurv2-row'));
                qurv2_recalcTotals(frm);
            }
        });

    wrapper.find('button[data-cell="attach"]').off('click.qurv2').on('click.qurv2', function () {
        const $btn = $(this);
        if ($btn.is(':disabled')) return;
        qurv2_pickFile(frm, $btn);
    });

    wrapper.find('.qurv2-save-draft').off('click.qurv2').on('click.qurv2', () => qurv2_saveDraft(frm));
    wrapper.find('.qurv2-submit').off('click.qurv2').on('click.qurv2', () => qurv2_submit(frm));
}

function qurv2_bindGlobalKeys(frm) {
    if (window.__qurv2_keys_bound) return;
    window.__qurv2_keys_bound = true;
    $(document).on('keydown.qurv2_keys', function (e) {
        const ctxOk = frm && frm[QURV2_NS] && frm[QURV2_NS].ctx;
        if (!ctxOk) return;
        if ((e.metaKey || e.ctrlKey) && (e.key === 's' || e.key === 'S')) {
            const wrapper = frm.fields_dict.quarterly_utilisation && frm.fields_dict.quarterly_utilisation.$wrapper;
            if (wrapper && wrapper.find('.qurv2-root').length) {
                e.preventDefault();
                qurv2_saveDraft(frm);
            }
        }
    });
    $(window).on('beforeunload.qurv2', function () {
        if (frm && frm[QURV2_NS] && frm[QURV2_NS].dirty) return __('You have unsaved utilisation changes. Leave anyway?');
    });
}

function qurv2_recalcRow($row) {
    const planAmt = +$row.data('plan-amt') || 0;
    const unitCost = +$row.data('unit-cost') || 0;
    const $input = $row.find('input[data-cell="actual-units-input"]');
    const actualUnits = qurv2_parseNum($input.val());
    const actualAmt = actualUnits * unitCost;
    const variance = planAmt - actualAmt;
    const spentPct = planAmt > 0 ? (actualAmt / planAmt * 100) : 0;

    $row.find('[data-cell="actual-amt"]').text(qurv2_inr(actualAmt));
    $row.find('[data-cell="variance"]').text(qurv2_inr(variance)).toggleClass('qurv2-over', variance < 0);
    $row.find('[data-cell="var-pct"]').text(qurv2_pct_spent(spentPct)).toggleClass('qurv2-over', spentPct > 100);
    $row.toggleClass('qurv2-overspent', actualAmt > planAmt);
    $row.data('actual-units', actualUnits);
}

function qurv2_recalcAll(frm) {
    const wrapper = frm.fields_dict.quarterly_utilisation.$wrapper;
    wrapper.find('.qurv2-row').each(function () { qurv2_recalcRow($(this)); });
    qurv2_recalcTotals(frm);
}

function qurv2_recalcTotals(frm) {
    const wrapper = frm.fields_dict.quarterly_utilisation.$wrapper;
    const ctx = frm[QURV2_NS].ctx;

    let progPlan = 0, progActual = 0, npPlan = 0, npActual = 0;
    wrapper.find('.qurv2-row').each(function () {
        const $r = $(this);
        const bpuName = $r.data('bpu');
        const bpu = ctx.bpuFullList.find(b => b.name === bpuName);
        const isProg = bpu && bpu.custom_activity;
        const planAmt = +$r.data('plan-amt') || 0;
        const actualUnits = +$r.data('actual-units') || 0;
        const unitCost = +$r.data('unit-cost') || 0;
        const actualAmt = actualUnits * unitCost;
        if (isProg) { progPlan += planAmt; progActual += actualAmt; }
        else { npPlan += planAmt; npActual += actualAmt; }
    });
    const gtPlan = progPlan + npPlan, gtActual = progActual + npActual;

    function setRow(prefix, plan, actual) {
        const v = plan - actual;
        const sp = plan > 0 ? (actual / plan * 100) : 0;
        wrapper.find(`[data-st="${prefix}-plan"]`).text(qurv2_inr(plan));
        wrapper.find(`[data-st="${prefix}-actual"]`).text(qurv2_inr(actual));
        wrapper.find(`[data-st="${prefix}-var"]`).text(qurv2_inr(v)).toggleClass('qurv2-over', v < 0);
        wrapper.find(`[data-st="${prefix}-pct"]`).text(qurv2_pct_spent(sp)).toggleClass('qurv2-over', sp > 100);
    }
    setRow('prog', progPlan, progActual);
    setRow('np', npPlan, npActual);
    setRow('gt', gtPlan, gtActual);
}

// --- Dirty tracking ------------------------------------------------------

function qurv2_markDirty(frm) {
    if (!frm[QURV2_NS].dirty) {
        frm[QURV2_NS].dirty = true;
        qurv2_updateDirtyChrome(frm);
    } else {
        qurv2_updateDirtyChrome(frm); // recount
    }
}

function qurv2_clearDirty(frm) {
    frm[QURV2_NS].dirty = false;
    qurv2_updateDirtyChrome(frm);
}

function qurv2_updateDirtyChrome(frm) {
    const wrapper = frm.fields_dict.quarterly_utilisation.$wrapper;
    const dirty = !!frm[QURV2_NS].dirty;
    const $dirty = wrapper.find('.qurv2-dirty');
    const $status = wrapper.find('.qurv2-actions-bar .qurv2-status');
    const $saveBtn = wrapper.find('.qurv2-save-draft');
    if (dirty) {
        const n = qurv2_countDirtyRows(frm);
        $dirty.text(__('{0} unsaved change(s)', [n])).addClass('has-unsaved');
        $status.text(__('{0} unsaved change(s)', [n])).addClass('has-unsaved');
        $saveBtn.addClass('has-changes');
    } else {
        $dirty.text('').removeClass('has-unsaved');
        $status.text(__('No unsaved changes')).removeClass('has-unsaved');
        $saveBtn.removeClass('has-changes');
    }
}

function qurv2_countDirtyRows(frm) {
    const ctx = frm[QURV2_NS].ctx;
    const cur = ctx.currentQ;
    const curQur = ctx.qStates[cur] && ctx.qStates[cur].qur;
    const childByBpu = {};
    if (curQur) for (const ch of (curQur.quarterly_utilisation || [])) childByBpu[ch.budget_plan] = ch;
    const wrapper = frm.fields_dict.quarterly_utilisation.$wrapper;
    let n = 0;
    wrapper.find('.qurv2-row').each(function () {
        const $r = $(this);
        const bpu = $r.data('bpu');
        const draft = childByBpu[bpu];
        const savedUnits = draft ? (+draft.custom_actual_units || 0) : 0;
        const savedComment = draft ? (draft.comment || '') : '';
        const savedSupport = draft ? (draft.support_documents || '') : '';
        const curUnits = qurv2_parseNum($r.find('input[data-cell="actual-units-input"]').val());
        const curComment = $r.find('input[data-cell="comment"]').val() || '';
        const curSupport = $r.data('support') || '';
        if (savedUnits !== curUnits || savedComment !== curComment || savedSupport !== curSupport) n++;
    });
    return n;
}

// --- Save / Submit -------------------------------------------------------

async function qurv2_saveDraft(frm, opts) {
    opts = opts || {};
    const ctx = frm[QURV2_NS].ctx;
    const cur = ctx.currentQ;
    const curState = ctx.qStates[cur].state;
    if (curState === 'filed' || curState === 'review') {
        if (!opts.silent) frappe.show_alert({ message: __('Q{0} is already submitted; cannot edit.', [cur]), indicator: 'orange' }, 5);
        return null;
    }
    const curQ = ctx.quarters.find(q => q.n === cur);
    const wrapper = frm.fields_dict.quarterly_utilisation.$wrapper;
    const rows = qurv2_collectRows(wrapper);

    // Validation: over-spend rows must have a Comment explaining the variance
    const overspendMissingComment = rows.filter(r => (r.actualUnits * r.unitCost) > r.planAmt && !(r.comment || '').trim());
    if (overspendMissingComment.length > 0) {
        const items = overspendMissingComment.slice(0, 3).map(r => {
            const bpu = ctx.bpuFullList.find(b => b.name === r.bpu);
            return bpu ? bpu.item_name : r.bpu;
        });
        const moreNote = overspendMissingComment.length > 3 ? __(' (+ {0} more)', [overspendMissingComment.length - 3]) : '';
        frappe.msgprint({
            title: __('Comment required for over-spend'),
            message: __('{0} row(s) report actuals above plan. Please add a Comment explaining each variance:', [overspendMissingComment.length])
                     + '<br><br><strong>' + items.map(s => frappe.utils.escape_html(s)).join('<br>') + '</strong>' + moreNote,
            indicator: 'orange'
        });
        return null;
    }

    frappe.dom.freeze(__('Saving Q{0} draft…', [cur]));
    try {
        const saved = await qurv2_persist(ctx, cur, curQ, rows, /*retry=*/true);
        ctx.qStates[cur].qur = saved;
        ctx.qStates[cur].state = 'drafting';
        ctx.qurFullByQ[cur] = saved;
        qurv2_clearDirty(frm);
        if (!opts.silent) {
            frappe.show_alert({
                message: __('Q{0} draft saved · Total ₹ {1}', [cur, qurv2_inr_plain(rows.reduce((s, r) => s + r.actualUnits * r.unitCost, 0))]),
                indicator: 'green'
            }, 4);
        }
        return saved;
    } catch (e) {
        const msg = qurv2_extractErrorMessage(e);
        frappe.msgprint({ title: __('Save failed'), message: frappe.utils.escape_html(msg), indicator: 'red' });
        return null;
    } finally {
        frappe.dom.unfreeze();
    }
}

async function qurv2_persist(ctx, cur, curQ, rows, allowRetry) {
    let qur = ctx.qStates[cur].qur;
    // Server constraint: actuals <= budget. To allow honest over-spend reporting,
    // we cap server-side `actuals` at budget but preserve real units in custom_actual_units.
    // The dialog reads custom_actual_units so display always shows real over-spend.
    const childPayload = rows.map(r => {
        const realActuals = r.actualUnits * r.unitCost;
        const cappedActuals = Math.min(realActuals, r.planAmt);
        return {
            budget_plan: r.bpu,
            fund_source: 'Primary Donor',
            custom_actual_units: r.actualUnits,
            custom_planned_units: r.planUnits,
            custom_unit_cost: r.unitCost,
            budget: r.planAmt,
            actuals: cappedActuals,
            variance: r.planAmt - cappedActuals,
            comment: r.comment,
            support_documents: r.support
        };
    });

    try {
        if (!qur || !qur.name) {
            const ins = await frappe.call({
                method: 'frappe.client.insert',
                args: { doc: {
                    doctype: 'Quarterly Utilisation Report',
                    grant: ctx.grantName, ngo: ctx.ngo, donor: ctx.donor,
                    quarter: cur,
                    timespan: curQ.timespan || ('Q' + cur),
                    start_date: curQ.start, end_date: curQ.end,
                    workflow_state: 'Pending',
                    utilisation_type: 'Utilisation',
                    quarterly_utilisation: childPayload
                } }
            });
            return ins.message;
        } else {
            qur.quarterly_utilisation = childPayload;
            const u = await frappe.call({ method: 'frappe.client.save', args: { doc: qur } });
            return u.message;
        }
    } catch (e) {
        const msg = qurv2_extractErrorMessage(e);
        if (allowRetry && /TimestampMismatchError|has been modified/i.test(msg)) {
            // Refetch and retry once
            if (qur && qur.name) {
                const fresh = await frappe.db.get_doc('Quarterly Utilisation Report', qur.name);
                ctx.qStates[cur].qur = fresh;
                return await qurv2_persist(ctx, cur, curQ, rows, /*allowRetry=*/false);
            }
        }
        throw e;
    }
}

function qurv2_collectRows(wrapper) {
    const rows = [];
    wrapper.find('.qurv2-row').each(function () {
        const $r = $(this);
        rows.push({
            bpu: $r.data('bpu'),
            actualUnits: qurv2_parseNum($r.find('input[data-cell="actual-units-input"]').val()),
            comment: $r.find('input[data-cell="comment"]').val() || '',
            unitCost: +$r.data('unit-cost') || 0,
            planAmt: +$r.data('plan-amt') || 0,
            planUnits: +$r.data('plan-units') || 0,
            support: $r.data('support') || ''
        });
    });
    return rows;
}

async function qurv2_submit(frm) {
    const ctx = frm[QURV2_NS].ctx;
    const cur = ctx.currentQ;

    // Save draft first
    const saved = await qurv2_saveDraft(frm, { silent: true });
    if (!saved) return; // save failed; error already shown

    // Confirm if all-zero
    const totalActuals = (saved.quarterly_utilisation || []).reduce((s, c) => s + (+c.actuals || 0), 0);
    const proceed = totalActuals === 0
        ? await new Promise(resolve => frappe.confirm(
            __('All actuals for Q{0} are zero. Submit a zero-utilisation report?', [cur]),
            () => resolve(true), () => resolve(false)))
        : true;
    if (!proceed) return;

    frappe.dom.freeze(__('Submitting Q{0}…', [cur]));
    try {
        const r = await frappe.xcall('frappe.model.workflow.apply_workflow', { doc: saved, action: 'Submit' });
        // Find next quarter (if any) so we can offer a Continue CTA
        const nextQ = ctx.quarters.find(q => q.n > cur);
        const baseMsg = __('Q{0} submitted for approval', [cur]);
        if (nextQ) {
            frappe.show_alert({
                message: baseMsg + ' · <a class="qurv2-continue-link" data-q="' + nextQ.n + '" style="color: #fff; text-decoration: underline; cursor: pointer;">'
                         + __('Continue with Q{0} →', [nextQ.n]) + '</a>',
                indicator: 'green'
            }, 8);
            // Wire the link click — re-render to next quarter
            setTimeout(() => {
                $(document).off('click.qurv2_continue').on('click.qurv2_continue', '.qurv2-continue-link', function () {
                    const target = parseInt($(this).data('q'), 10);
                    if (target && frm[QURV2_NS].ctx) {
                        frm[QURV2_NS].ctx.currentQ = target;
                        qurv2_paint(frm);
                    }
                });
            }, 50);
        } else {
            frappe.show_alert({ message: baseMsg, indicator: 'green' }, 5);
        }
        await qurv2_render(frm);
    } catch (e) {
        const msg = qurv2_extractErrorMessage(e);
        frappe.msgprint({ title: __('Submit failed'), message: frappe.utils.escape_html(msg), indicator: 'red' });
    } finally {
        frappe.dom.unfreeze();
    }
}

function qurv2_extractErrorMessage(e) {
    if (!e) return 'Unknown error';
    if (e.responseJSON && e.responseJSON.exception) {
        const m = String(e.responseJSON.exception).split(':');
        return m.slice(1).join(':').trim() || e.responseJSON.exception;
    }
    if (e.message) return e.message;
    return String(e);
}

// --- File picker ---------------------------------------------------------

function qurv2_pickFile(frm, $btn) {
    new frappe.ui.FileUploader({
        method: 'frappe.client.attach_file',
        on_success: (file_doc) => {
            const url = file_doc.file_url;
            const $row = $btn.closest('.qurv2-row');
            $row.data('support', url);
            $btn.addClass('has-file').text('✓').attr('title', __('Attached'));
            qurv2_markDirty(frm);
        }
    });
}

// --- Helpers -------------------------------------------------------------

function qurv2_quarterNum(c) {
    if (typeof c.quarter === 'number') return c.quarter;
    if (c.timespan && c.timespan.match(/^Q(\d)/)) return parseInt(RegExp.$1, 10);
    if (typeof c.quarter === 'string') {
        const m = c.quarter.match(/(\d)/);
        if (m) return parseInt(m[1], 10);
    }
    return 0;
}

function qurv2_inr(n) {
    n = +n || 0;
    const sign = n < 0 ? '-' : '';
    n = Math.abs(n);
    return sign + '₹ ' + new Intl.NumberFormat('en-IN', { maximumFractionDigits: 0 }).format(Math.round(n));
}

function qurv2_inr_plain(n) {
    n = +n || 0;
    const sign = n < 0 ? '-' : '';
    n = Math.abs(n);
    return sign + new Intl.NumberFormat('en-IN', { maximumFractionDigits: 0 }).format(Math.round(n));
}

function qurv2_num(n) {
    n = +n || 0;
    if (n === 0) return '0';
    if (n === Math.round(n)) return String(Math.round(n));
    return n.toFixed(2).replace(/\.?0+$/, '');
}

function qurv2_pct(n) {
    n = +n || 0;
    const sign = n > 0 ? '+' : (n < 0 ? '−' : '');
    return sign + Math.abs(n).toFixed(0) + '%';
}

// Spent% formatting: 0%, 45%, 100% — when over plan, shows "+X% over"
function qurv2_pct_spent(n) {
    n = +n || 0;
    if (n > 100) return '+' + (n - 100).toFixed(0) + '% over';
    return n.toFixed(0) + '%';
}

function qurv2_parseNum(s) {
    if (!s) return 0;
    const cleaned = String(s).replace(/[^0-9.\-]/g, '');
    const n = parseFloat(cleaned);
    return isNaN(n) ? 0 : n;
}
