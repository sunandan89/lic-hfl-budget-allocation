// ============================================================================
// LIC Budget Utilisation Reporting v3 — Grant · Budget Report tab
// Client Script for: Grant
// ----------------------------------------------------------------------------
// Renders the simplified single-quarter utilisation surface on
// `budget_report_html` (inside the Budget Report tab on Grant). Replaces the
// broken `frappe_theme.dt_api.get_dt_list` Task Tracker view that was throwing
// OperationalError 1054 (`tp.custom_milestone_display_order` missing).
//
// Design (mirrors v9 budget allocation):
//   - Quarter pills Q1..Q4 — pick the quarter being reported
//   - Activity row: name · unit · target units · unit cost · plan ₹ · up to
//     Q(n-1) ₹ locked · this Q units (input) · this Q ₹ auto · % plan
//   - Annual review toggle: read-only Q1..Q4 breakdown for audit / Excel
//
// Data model bridge:
//   - Budget Plan and Utilisation (BPU) holds `planning_table` (per-quarter
//     planned + utilised ₹) for this grant.
//   - The matching Project Budget Planning (PBP) record (matched by
//     description) holds per-quarter `unit_cost` on its PBP Child rows.
//   - Save writes back `utilised_amount` on Budget Planning Child:
//       utilised_amount = units_entered × unit_cost
// ============================================================================

frappe.ui.form.on('Grant', {
  onload(frm)  { bu_install_intercept(); bu_setup(frm); },
  refresh(frm) { bu_install_intercept(); bu_setup(frm); },
  after_save(frm) { frm._bu_rendered_for = null; }
});

// ---------------------------------------------------------------------------
// 1) Suppress the broken Task Tracker report query at the network layer
// ---------------------------------------------------------------------------
function bu_install_intercept() {
  if (window._bu_call_patched) return;
  var orig = frappe.call;
  frappe.call = function(opts) {
    try {
      if (opts && opts.method === 'frappe_theme.dt_api.get_dt_list') {
        var a = opts.args || {};
        if (a.doctype === 'Task Tracker' && a.ref_doctype === 'Grant') {
          return Promise.resolve({ message: { columns: [], result: [], chart: null,
            report_summary: [], skip_total_row: 1, status: null,
            execution_time: 0, add_total_row: 0 } });
        }
      }
    } catch (e) {}
    return orig.apply(this, arguments);
  };
  window._bu_call_patched = true;
}

// ---------------------------------------------------------------------------
// 2) Setup gate
// ---------------------------------------------------------------------------
function bu_setup(frm) {
  if (frm.doc.__islocal) return;
  var $w = frm.fields_dict.budget_report_html && frm.fields_dict.budget_report_html.$wrapper;
  if (!$w || !$w.length) return;

  if (!$w._bu_patched) {
    var origHtml = $.fn.html;
    $w.html = function() {
      if (arguments.length === 0) return origHtml.apply(this, arguments);
      if (frm._bu_self_render) return origHtml.apply(this, arguments);
      return this;
    };
    $w._bu_patched = true;
  }

  if (frm._bu_rendered_for === frm.doc.name) return;
  bu_render(frm).catch(function(e){ console.error('[BU] render error:', e); });
}

// ---------------------------------------------------------------------------
// 3) Main render
// ---------------------------------------------------------------------------
async function bu_render(frm) {
  var $w = frm.fields_dict.budget_report_html.$wrapper;
  bu_paint($w, '<div style="padding:24px;text-align:center;color:#888;font:13px/1.4 -apple-system,system-ui,sans-serif;">Loading budget utilisation…</div>', frm);

  // BPU rows for this Grant
  var bpuList = await frappe.call({
    method: 'frappe.client.get_list',
    args: {
      doctype: 'Budget Plan and Utilisation',
      fields: ['name','item_name','budget_head','sub_budget_head','custom_activity','custom_unit_of_measurement','custom_task_details','total_planned_budget','total_utilisation'],
      filters: { grant: frm.doc.name },
      limit_page_length: 500,
      order_by: 'budget_head asc, item_name asc'
    }
  }).then(function(r){return (r&&r.message)||[];}).catch(function(){return [];});

  if (!bpuList.length) {
    bu_paint($w, bu_emptyState('No budget items have been allocated for this grant yet. Add activities under the Budget Allocation tab first.'), frm);
    frm._bu_rendered_for = frm.doc.name;
    return;
  }

  // Fetch full BPU records (with planning_table)
  var bpuFull = {};
  for (var i = 0; i < bpuList.length; i += 8) {
    var chunk = bpuList.slice(i, i + 8);
    await Promise.all(chunk.map(function(b){
      return frappe.call({method:'frappe.client.get', args:{doctype:'Budget Plan and Utilisation', name:b.name}})
        .then(function(r){if (r && r.message) bpuFull[b.name] = r.message;}).catch(function(){});
    }));
  }

  // PBP records by description for unit cost bridge
  var descs = bpuList.map(function(b){return b.item_name;}).filter(Boolean).filter(function(v,i,a){return a.indexOf(v)===i;});
  var pbpByDesc = {};
  await Promise.all(descs.map(function(d){
    return frappe.call({method:'frappe.client.get_list', args:{doctype:'Project Budget Planning', fields:['name','description','total_planned_budget','custom_total_lic_contribution'], filters:{description:d}, limit_page_length:5}})
      .then(function(r){
        var hits = (r&&r.message)||[];
        if (hits.length) pbpByDesc[d] = hits[0].name;
      }).catch(function(){});
  }));

  // Pull each matched PBP's planning_table → unit_cost map per (year, quarter)
  var pbpFull = {};
  var pbpNames = Object.values(pbpByDesc).filter(function(v,i,a){return a.indexOf(v)===i;});
  for (var k = 0; k < pbpNames.length; k += 8) {
    var chunk2 = pbpNames.slice(k, k + 8);
    await Promise.all(chunk2.map(function(n){
      return frappe.call({method:'frappe.client.get', args:{doctype:'Project Budget Planning', name:n}})
        .then(function(r){if (r && r.message) pbpFull[n] = r.message;}).catch(function(){});
    }));
  }

  // Resolve unit-of-measurement names
  var uomIds = bpuList.map(function(b){return b.custom_unit_of_measurement;}).filter(Boolean).filter(function(v,i,a){return a.indexOf(v)===i;});
  var uomMap = {};
  await Promise.all(uomIds.map(function(uid){
    return frappe.call({method:'frappe.client.get_value', args:{doctype:'Units', filters:{name:uid}, fieldname:'unit_name'}})
      .then(function(r){if (r && r.message && r.message.unit_name) uomMap[uid] = r.message.unit_name;}).catch(function(){});
  }));

  // Resolve sub-budget-head names
  var sbhIds = bpuList.map(function(b){return b.sub_budget_head;}).filter(Boolean).filter(function(v,i,a){return a.indexOf(v)===i;});
  var sbhMap = {};
  await Promise.all(sbhIds.map(function(id){
    return frappe.call({method:'frappe.client.get_value', args:{doctype:'Sub Budget Head', filters:{name:id}, fieldname:'sub_budget_head'}})
      .then(function(r){if (r && r.message && r.message.sub_budget_head) sbhMap[id] = r.message.sub_budget_head;}).catch(function(){});
  }));

  // Build organised data
  var rows = bpuList.map(function(b){
    var full = bpuFull[b.name] || b;
    var planning = (full.planning_table || []).slice().sort(function(a,c){
      if (a.year !== c.year) return (a.year||0) - (c.year||0);
      return (a.quarter||0) - (c.quarter||0);
    });

    // Map of (year,quarter) -> unit_cost from matching PBP
    var unitCostMap = {};
    var pbpName = pbpByDesc[b.item_name];
    if (pbpName && pbpFull[pbpName]) {
      var pbpRows = pbpFull[pbpName].planning_table || [];
      pbpRows.forEach(function(p){
        var qNum = bu_parseQ(p.quarter);
        var key = (p.year||1) + ':' + qNum;
        if (Number(p.unit_cost||0) > 0) unitCostMap[key] = Number(p.unit_cost);
      });
    }

    // Pick the most common unit_cost as the row's display unit_cost
    var costsList = Object.values(unitCostMap);
    var dominantUC = costsList.length ? bu_mode(costsList) : 0;
    var totalPlanned = Number(b.total_planned_budget||0);
    var targetUnits = (dominantUC > 0) ? (totalPlanned / dominantUC) : 0;

    return {
      bpu: b.name,
      item_name: b.item_name || b.name,
      task_details: full.custom_task_details || '',
      sbh_name: sbhMap[b.sub_budget_head] || b.sub_budget_head || '',
      uom: uomMap[b.custom_unit_of_measurement] || '',
      total_planned: totalPlanned,
      total_utilised: Number(b.total_utilisation||0),
      planning: planning, // {name, year, quarter, planned_amount, utilised_amount, ...}
      unit_cost_map: unitCostMap,
      unit_cost: dominantUC,
      target_units: targetUnits
    };
  });

  // Years present
  var yearSet = {};
  rows.forEach(function(r){ r.planning.forEach(function(p){ if (p.year) yearSet[p.year] = true; }); });
  var years = Object.keys(yearSet).map(Number).sort(function(a,b){return a-b;});
  if (!years.length) years = [1];

  // Default quarter: earliest with no utilisation entered yet
  function pickDefaultQuarter(year) {
    for (var q = 1; q <= 4; q++) {
      var has = rows.some(function(r){
        var c = r.planning.find(function(p){return p.year===year && p.quarter===q;});
        return c && Number(c.utilised_amount||0) > 0;
      });
      if (!has) return q;
    }
    return 4;
  }

  var state = {
    activeYear: years[0],
    activeQuarter: pickDefaultQuarter(years[0]),
    activeView: 'reporting',
    rows: rows,
    years: years,
    dirty: {} // {bpuName: {childName: {units, amount}}}
  };

  bu_paint($w, bu_buildHTML(state, frm), frm);
  bu_attachEvents($w, state, frm);

  frm._bu_rendered_for = frm.doc.name;
}

// ---------------------------------------------------------------------------
// 4) HTML
// ---------------------------------------------------------------------------
function bu_buildHTML(state, frm) {
  return '<div class="bu-root">' +
    bu_styles() +
    bu_topbar(state) +
    (state.activeView === 'reporting' ? bu_reportingView(state) : bu_annualView(state)) +
    bu_bottomBar(state) +
  '</div>';
}

function bu_styles() {
  return '<style>' +
    '.bu-root{padding:16px 18px;background:#FAFAF7;border-radius:6px;font:13px/1.5 -apple-system,system-ui,Segoe UI,sans-serif;color:#222;}' +
    '.bu-topbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:14px;}' +
    '.bu-toplabel{font-size:12px;color:#666;margin-right:4px;}' +
    '.bu-pill{font-size:12px;padding:5px 12px;border-radius:6px;border:1px solid #D8D7D2;background:#FFF;color:#444;cursor:pointer;user-select:none;}' +
    '.bu-pill:hover{border-color:#00529C;color:#00529C;}' +
    '.bu-pill.active{background:#00529C;color:#FFF;border-color:#00529C;font-weight:500;}' +
    '.bu-pill.locked{background:#F1EFE8;color:#888;cursor:default;}' +
    '.bu-viewtoggle{margin-left:auto;display:flex;gap:6px;}' +
    '.bu-table-wrap{background:#FFF;border:1px solid #E2E0DA;border-radius:6px;overflow:auto;max-height:62vh;}' +
    '.bu-table{border-collapse:collapse;font-size:12px;width:100%;table-layout:auto;}' +
    '.bu-table th{padding:8px 10px;text-align:left;font-weight:500;background:#00529C;color:#FFF;border-right:1px solid #003B70;position:sticky;top:0;z-index:2;white-space:nowrap;}' +
    '.bu-table th.bu-sub{background:#003B70;font-size:11px;font-weight:400;}' +
    '.bu-table th.bu-yt{background:#FFCB05;color:#4A1B0C;}' +
    '.bu-table th.bu-locked{background:#444441;}' +
    '.bu-table td{padding:6px 10px;border-bottom:1px solid #EFEDE6;vertical-align:middle;}' +
    '.bu-table tr.bu-grouphdr td{background:#F1EFE8;font-weight:500;color:#333;}' +
    '.bu-num{text-align:right;font-variant-numeric:tabular-nums;}' +
    '.bu-num.muted{color:#888;}' +
    '.bu-locked-col{background:#F8F7F2;color:#444;}' +
    '.bu-input-cell{padding:4px 6px;}' +
    '.bu-input{width:100%;min-width:70px;padding:5px 8px;border:1px solid #00529C;background:#E6F1FB;color:#042C53;border-radius:4px;font:inherit;font-weight:500;text-align:right;font-variant-numeric:tabular-nums;}' +
    '.bu-input:focus{outline:none;border-color:#003B70;background:#FFF;}' +
    '.bu-input[disabled]{background:#F1EFE8;border-color:#D8D7D2;color:#888;cursor:not-allowed;}' +
    '.bu-pct{background:#FFFCEB;font-weight:500;}' +
    '.bu-pct.good{color:#3B6D11;} .bu-pct.warn{color:#854F0B;} .bu-pct.bad{color:#791F1F;}' +
    '.bu-totalrow td{background:#BCBDC0;font-weight:500;color:#2C2C2A;}' +
    '.bu-totalrow td.bu-yt{background:#FFCB05;}' +
    '.bu-bottombar{display:flex;align-items:center;gap:10px;margin-top:14px;}' +
    '.bu-btn{padding:7px 16px;border-radius:6px;border:1px solid #00529C;background:#00529C;color:#FFF;font-size:12px;font-weight:500;cursor:pointer;}' +
    '.bu-btn:hover{background:#003B70;}' +
    '.bu-btn.ghost{background:#FFF;color:#00529C;}' +
    '.bu-btn.ghost:hover{background:#E6F1FB;}' +
    '.bu-btn[disabled]{opacity:0.5;cursor:not-allowed;}' +
    '.bu-meta{font-size:11px;color:#888;margin-left:auto;}' +
    '.bu-empty{padding:30px 24px;text-align:center;color:#888;background:#FFF;border:1px dashed #D8D7D2;border-radius:6px;}' +
    '.bu-activity{max-width:280px;}' +
    '.bu-activity-name{font-weight:500;color:#2C2C2A;}' +
    '.bu-activity-task{font-size:11px;color:#888;margin-top:1px;line-height:1.3;}' +
    '.bu-warn-banner{background:#FFFCEB;border:1px solid #FAC775;border-radius:4px;padding:8px 12px;margin-bottom:10px;font-size:12px;color:#854F0B;}' +
  '</style>';
}

function bu_topbar(state) {
  var qPills = '';
  for (var q = 1; q <= 4; q++) {
    var cls = 'bu-pill' + (q === state.activeQuarter ? ' active' : '');
    qPills += '<span class="' + cls + '" data-action="setq" data-q="' + q + '">Q' + q + '</span>';
  }
  return '<div class="bu-topbar">' +
    '<span class="bu-toplabel">Reporting for</span>' +
    qPills +
    '<span class="bu-toplabel" style="margin-left:8px;">·</span>' +
    '<span class="bu-pill active">Year ' + state.activeYear + '</span>' +
    '<div class="bu-viewtoggle">' +
      '<span class="bu-pill ' + (state.activeView==='reporting'?'active':'') + '" data-action="setview" data-v="reporting">Reporting</span>' +
      '<span class="bu-pill ' + (state.activeView==='annual'?'active':'') + '" data-action="setview" data-v="annual">Annual review</span>' +
    '</div>' +
  '</div>';
}

function bu_reportingView(state) {
  var year = state.activeYear, q = state.activeQuarter;
  var groups = {};
  state.rows.forEach(function(r){
    var key = r.sbh_name || 'Other';
    if (!groups[key]) groups[key] = [];
    groups[key].push(r);
  });

  var totalsPlan = 0, totalsPrior = 0, totalsCurrent = 0;
  var anyMissingCost = false;
  var body = '';

  Object.keys(groups).forEach(function(grp){
    body += '<tr class="bu-grouphdr"><td colspan="10">' + bu_esc(grp) + '</td></tr>';
    groups[grp].forEach(function(r){
      var planRow = r.planning.filter(function(p){return p.year===year;});
      var planned = planRow.reduce(function(s,p){return s + Number(p.planned_amount||0);}, 0);
      var prior = planRow.filter(function(p){return p.quarter < q;}).reduce(function(s,p){return s + Number(p.utilised_amount||0);}, 0);
      var currentEntry = planRow.find(function(p){return p.quarter === q;});
      var currentInit = currentEntry ? Number(currentEntry.utilised_amount||0) : 0;

      // Per-quarter unit_cost (fall back to dominant)
      var qKey = year + ':' + q;
      var uc = r.unit_cost_map[qKey] || r.unit_cost || 0;
      if (!uc) anyMissingCost = true;
      var unitsInit = uc > 0 ? (currentInit / uc) : 0;
      var newYTD = prior + currentInit;
      var pct = planned ? (newYTD / planned * 100) : 0;
      var pctCls = pct >= 95 ? 'good' : pct >= 70 ? 'warn' : 'bad';

      totalsPlan += planned; totalsPrior += prior; totalsCurrent += currentInit;

      var ce = currentEntry ? currentEntry.name : '';
      body += '<tr data-bpu="' + bu_esc(r.bpu) + '" data-child="' + bu_esc(ce) + '" data-uc="' + uc + '">' +
        '<td class="bu-activity">' +
          '<div class="bu-activity-name">' + bu_esc(r.item_name) + '</div>' +
          (r.task_details ? '<div class="bu-activity-task">' + bu_esc(r.task_details) + '</div>' : '') +
        '</td>' +
        '<td>' + bu_esc(r.uom || '—') + '</td>' +
        '<td class="bu-num">' + (r.target_units > 0 ? bu_num(r.target_units) : '—') + '</td>' +
        '<td class="bu-num">' + (uc > 0 ? bu_inr(uc) : '—') + '</td>' +
        '<td class="bu-num muted">' + bu_inr(planned) + '</td>' +
        '<td class="bu-num bu-locked-col">' + (q === 1 ? '<span style="color:#bbb;">—</span>' : bu_inr(prior)) + '</td>' +
        '<td class="bu-input-cell">' +
          (currentEntry
            ? '<input type="number" min="0" step="any" class="bu-input bu-units" data-init="' + unitsInit + '" value="' + (unitsInit ? bu_round(unitsInit, 4) : '') + '" ' + (uc > 0 ? '' : 'disabled') + ' />'
            : '<span style="color:#bbb;">—</span>') +
        '</td>' +
        '<td class="bu-num"><span class="bu-q-amount">' + bu_inr(currentInit) + '</span></td>' +
        '<td class="bu-num"><span class="bu-newytd">' + bu_inr(newYTD) + '</span></td>' +
        '<td class="bu-num bu-pct ' + pctCls + '"><span class="bu-pct-out">' + pct.toFixed(1) + '%</span></td>' +
      '</tr>';
    });
  });

  var totalNewYTD = totalsPrior + totalsCurrent;
  var totalPct = totalsPlan ? (totalNewYTD / totalsPlan * 100) : 0;
  var totalPctCls = totalPct >= 95 ? 'good' : totalPct >= 70 ? 'warn' : 'bad';

  body += '<tr class="bu-totalrow">' +
    '<td colspan="4">Total · all activities</td>' +
    '<td class="bu-num">' + bu_inr(totalsPlan) + '</td>' +
    '<td class="bu-num">' + (q === 1 ? '—' : bu_inr(totalsPrior)) + '</td>' +
    '<td></td>' +
    '<td class="bu-num"><span class="bu-total-q-amount">' + bu_inr(totalsCurrent) + '</span></td>' +
    '<td class="bu-num bu-yt"><span class="bu-total-newytd">' + bu_inr(totalNewYTD) + '</span></td>' +
    '<td class="bu-num bu-yt ' + totalPctCls + '"><span class="bu-total-pct">' + totalPct.toFixed(1) + '%</span></td>' +
  '</tr>';

  var warn = anyMissingCost ? '<div class="bu-warn-banner">Some activities don\'t yet have a unit cost defined in their Project Budget Planning record. Those rows are read-only — set unit cost upstream first.</div>' : '';

  return warn + '<div class="bu-table-wrap"><table class="bu-table">' +
    '<thead>' +
      '<tr>' +
        '<th rowspan="2">Activity</th>' +
        '<th rowspan="2">Unit</th>' +
        '<th rowspan="2">Target units</th>' +
        '<th rowspan="2">Unit cost ₹</th>' +
        '<th rowspan="2">Plan ₹</th>' +
        '<th rowspan="2" class="bu-locked">Up to Q' + (q-1 || '—') + ' ₹</th>' +
        '<th colspan="2">Q' + q + ' utilisation</th>' +
        '<th rowspan="2">New YTD ₹</th>' +
        '<th rowspan="2" class="bu-yt">% plan</th>' +
      '</tr>' +
      '<tr>' +
        '<th class="bu-sub">Units done</th>' +
        '<th class="bu-sub">₹ auto</th>' +
      '</tr>' +
    '</thead>' +
    '<tbody>' + body + '</tbody>' +
  '</table></div>';
}

function bu_annualView(state) {
  var year = state.activeYear;
  var totals = { plan:0, q1:0, q2:0, q3:0, q4:0 };
  var groups = {};
  state.rows.forEach(function(r){
    var key = r.sbh_name || 'Other';
    if (!groups[key]) groups[key] = [];
    groups[key].push(r);
  });

  var body = '';
  Object.keys(groups).forEach(function(grp){
    body += '<tr class="bu-grouphdr"><td colspan="9">' + bu_esc(grp) + '</td></tr>';
    groups[grp].forEach(function(r){
      var pr = r.planning.filter(function(p){return p.year===year;});
      var planned = pr.reduce(function(s,p){return s + Number(p.planned_amount||0);}, 0);
      var qa = [0,0,0,0,0];
      pr.forEach(function(p){if (p.quarter>=1 && p.quarter<=4) qa[p.quarter] = Number(p.utilised_amount||0);});
      var yt = qa[1]+qa[2]+qa[3]+qa[4];
      var pct = planned ? (yt / planned * 100) : 0;
      var pctCls = pct >= 95 ? 'good' : pct >= 70 ? 'warn' : 'bad';
      totals.plan += planned; totals.q1 += qa[1]; totals.q2 += qa[2]; totals.q3 += qa[3]; totals.q4 += qa[4];
      body += '<tr>' +
        '<td class="bu-activity"><div class="bu-activity-name">' + bu_esc(r.item_name) + '</div></td>' +
        '<td>' + bu_esc(r.uom || '—') + '</td>' +
        '<td class="bu-num muted">' + bu_inr(planned) + '</td>' +
        '<td class="bu-num">' + bu_inr(qa[1]) + '</td>' +
        '<td class="bu-num">' + bu_inr(qa[2]) + '</td>' +
        '<td class="bu-num">' + bu_inr(qa[3]) + '</td>' +
        '<td class="bu-num">' + bu_inr(qa[4]) + '</td>' +
        '<td class="bu-num">' + bu_inr(yt) + '</td>' +
        '<td class="bu-num bu-pct ' + pctCls + '">' + pct.toFixed(1) + '%</td>' +
      '</tr>';
    });
  });

  var grand = totals.q1+totals.q2+totals.q3+totals.q4;
  var grandPct = totals.plan ? (grand / totals.plan * 100) : 0;
  var grandPctCls = grandPct >= 95 ? 'good' : grandPct >= 70 ? 'warn' : 'bad';
  body += '<tr class="bu-totalrow">' +
    '<td colspan="2">Total</td>' +
    '<td class="bu-num">' + bu_inr(totals.plan) + '</td>' +
    '<td class="bu-num">' + bu_inr(totals.q1) + '</td>' +
    '<td class="bu-num">' + bu_inr(totals.q2) + '</td>' +
    '<td class="bu-num">' + bu_inr(totals.q3) + '</td>' +
    '<td class="bu-num">' + bu_inr(totals.q4) + '</td>' +
    '<td class="bu-num">' + bu_inr(grand) + '</td>' +
    '<td class="bu-num bu-yt ' + grandPctCls + '">' + grandPct.toFixed(1) + '%</td>' +
  '</tr>';

  return '<div class="bu-table-wrap"><table class="bu-table">' +
    '<thead><tr>' +
      '<th>Activity</th><th>Unit</th><th>Plan ₹</th>' +
      '<th>Q1 ₹</th><th>Q2 ₹</th><th>Q3 ₹</th><th>Q4 ₹</th>' +
      '<th>Year total</th><th class="bu-yt">% plan</th>' +
    '</tr></thead>' +
    '<tbody>' + body + '</tbody>' +
  '</table></div>';
}

function bu_bottomBar(state) {
  if (state.activeView !== 'reporting') {
    return '<div class="bu-bottombar">' +
      '<button class="bu-btn ghost" data-action="export">Download CSV</button>' +
      '<span class="bu-meta">Annual review · read-only</span>' +
    '</div>';
  }
  return '<div class="bu-bottombar">' +
    '<button class="bu-btn" data-action="save">Save Q' + state.activeQuarter + ' utilisation</button>' +
    '<button class="bu-btn ghost" data-action="reset">Reset changes</button>' +
    '<span class="bu-meta">Edits to Q' + state.activeQuarter + ' only · units × unit cost = ₹</span>' +
  '</div>';
}

function bu_emptyState(msg) {
  return '<div class="bu-root">' + bu_styles() + '<div class="bu-empty">' + bu_esc(msg) + '</div></div>';
}

// ---------------------------------------------------------------------------
// 5) Events
// ---------------------------------------------------------------------------
function bu_attachEvents($w, state, frm) {
  $w.off('.bu');

  $w.on('click.bu', '.bu-pill', function(){
    var $t = $(this);
    var act = $t.data('action');
    if (act === 'setq') {
      var q = Number($t.data('q'));
      if (q && q !== state.activeQuarter) {
        if (Object.keys(state.dirty).length) {
          if (!confirm('You have unsaved changes. Switching quarter will discard them. Continue?')) return;
          state.dirty = {};
        }
        state.activeQuarter = q;
        bu_paint($w, bu_buildHTML(state, frm), frm);
        bu_attachEvents($w, state, frm);
      }
    } else if (act === 'setview') {
      var v = $t.data('v');
      if (v && v !== state.activeView) {
        state.activeView = v;
        bu_paint($w, bu_buildHTML(state, frm), frm);
        bu_attachEvents($w, state, frm);
      }
    }
  });

  $w.on('input.bu change.bu', '.bu-units', function(){
    bu_recompute($w, state);
  });

  $w.on('click.bu', '[data-action="save"]', function(){ bu_save($w, state, frm); });
  $w.on('click.bu', '[data-action="reset"]', function(){
    state.dirty = {};
    bu_paint($w, bu_buildHTML(state, frm), frm);
    bu_attachEvents($w, state, frm);
  });
  $w.on('click.bu', '[data-action="export"]', function(){ bu_exportCSV(state, frm); });
}

function bu_recompute($w, state) {
  var totalPlan = 0, totalPrior = 0, totalCurrent = 0;
  $w.find('tbody tr[data-bpu]').each(function(){
    var $tr = $(this);
    var bpuName = $tr.attr('data-bpu');
    var rowMeta = state.rows.find(function(r){return r.bpu===bpuName;});
    if (!rowMeta) return;
    var year = state.activeYear, q = state.activeQuarter;
    var pr = rowMeta.planning.filter(function(p){return p.year===year;});
    var planned = pr.reduce(function(s,p){return s + Number(p.planned_amount||0);}, 0);
    var prior = pr.filter(function(p){return p.quarter < q;}).reduce(function(s,p){return s + Number(p.utilised_amount||0);}, 0);
    var $inp = $tr.find('.bu-units');
    var uc = Number($tr.attr('data-uc') || 0);
    var units = $inp.length ? Number($inp.val() || 0) : 0;
    var amount = units * uc;
    var newYTD = prior + amount;
    var pct = planned ? (newYTD / planned * 100) : 0;
    $tr.find('.bu-q-amount').text(bu_inr(amount));
    $tr.find('.bu-newytd').text(bu_inr(newYTD));
    var $pct = $tr.find('.bu-pct-out');
    $pct.text(pct.toFixed(1) + '%');
    $pct.closest('.bu-pct').removeClass('good warn bad').addClass(pct >= 95 ? 'good' : pct >= 70 ? 'warn' : 'bad');

    var initUnits = Number($inp.attr('data-init') || 0);
    if ($inp.length) {
      var childName = $tr.attr('data-child');
      if (childName) {
        if (Math.abs(units - initUnits) > 1e-9) {
          if (!state.dirty[bpuName]) state.dirty[bpuName] = {};
          state.dirty[bpuName][childName] = { units: units, amount: amount };
        } else if (state.dirty[bpuName] && state.dirty[bpuName][childName] !== undefined) {
          delete state.dirty[bpuName][childName];
          if (!Object.keys(state.dirty[bpuName]).length) delete state.dirty[bpuName];
        }
      }
    }

    totalPlan += planned; totalPrior += prior; totalCurrent += amount;
  });

  var totalNewYTD = totalPrior + totalCurrent;
  var totalPct = totalPlan ? (totalNewYTD / totalPlan * 100) : 0;
  $w.find('.bu-total-q-amount').text(bu_inr(totalCurrent));
  $w.find('.bu-total-newytd').text(bu_inr(totalNewYTD));
  var $tp = $w.find('.bu-total-pct');
  $tp.text(totalPct.toFixed(1) + '%');
  $tp.closest('td').removeClass('good warn bad').addClass(totalPct >= 95 ? 'good' : totalPct >= 70 ? 'warn' : 'bad');
}

// ---------------------------------------------------------------------------
// 6) Save: sequential per BPU (avoid Frappe document-lock collisions)
// ---------------------------------------------------------------------------
async function bu_save($w, state, frm) {
  var dirty = state.dirty;
  var bpuNames = Object.keys(dirty);
  if (!bpuNames.length) {
    frappe.show_alert({ message: 'No changes to save.', indicator: 'orange' });
    return;
  }

  var $btn = $w.find('[data-action="save"]');
  $btn.attr('disabled', true).text('Saving…');

  var ok = 0, fail = 0;
  for (var i = 0; i < bpuNames.length; i++) {
    var bpu = bpuNames[i];
    try {
      var doc = await frappe.call({method:'frappe.client.get', args:{doctype:'Budget Plan and Utilisation', name:bpu}}).then(function(r){return r&&r.message;});
      if (!doc) { fail++; continue; }
      var changedAny = false;
      (doc.planning_table || []).forEach(function(p){
        if (dirty[bpu] && dirty[bpu][p.name] !== undefined) {
          p.utilised_amount = dirty[bpu][p.name].amount;
          changedAny = true;
        }
      });
      if (!changedAny) continue;
      doc.total_utilisation = (doc.planning_table || []).reduce(function(s,p){return s + Number(p.utilised_amount||0);}, 0);
      await frappe.call({method:'frappe.client.save', args:{doc:doc}});
      ok++;
    } catch (e) {
      console.error('[BU] save error for', bpu, e);
      fail++;
    }
  }

  state.dirty = {};
  $btn.attr('disabled', false).text('Save Q' + state.activeQuarter + ' utilisation');

  if (fail === 0) frappe.show_alert({ message: 'Saved utilisation for ' + ok + ' activit' + (ok===1?'y':'ies') + '.', indicator: 'green' });
  else frappe.show_alert({ message: ok + ' saved · ' + fail + ' failed. Check console.', indicator: 'red' });

  frm._bu_rendered_for = null;
  bu_setup(frm);
}

// ---------------------------------------------------------------------------
// 7) Excel/CSV export (annual)
// ---------------------------------------------------------------------------
function bu_exportCSV(state, frm) {
  var year = state.activeYear;
  var lines = [['Activity','Unit','Sub-budget head','Plan ₹','Q1 ₹','Q2 ₹','Q3 ₹','Q4 ₹','Year total ₹','% plan']];
  state.rows.forEach(function(r){
    var pr = r.planning.filter(function(p){return p.year===year;});
    var planned = pr.reduce(function(s,p){return s + Number(p.planned_amount||0);}, 0);
    var qa = [0,0,0,0,0];
    pr.forEach(function(p){if (p.quarter>=1 && p.quarter<=4) qa[p.quarter] = Number(p.utilised_amount||0);});
    var yt = qa[1]+qa[2]+qa[3]+qa[4];
    var pct = planned ? (yt / planned * 100) : 0;
    lines.push([r.item_name, r.uom||'', r.sbh_name||'', planned, qa[1], qa[2], qa[3], qa[4], yt, pct.toFixed(2)]);
  });
  var csv = '\uFEFF' + lines.map(function(row){
    return row.map(function(c){
      var s = String(c==null?'':c);
      if (/[",\n]/.test(s)) s = '"' + s.replace(/"/g,'""') + '"';
      return s;
    }).join(',');
  }).join('\n');
  var blob = new Blob([csv], { type:'text/csv;charset=utf-8;' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = 'utilisation_' + frm.doc.name + '_Y' + year + '.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
}

// ---------------------------------------------------------------------------
// 8) Helpers
// ---------------------------------------------------------------------------
function bu_paint($w, html, frm) { frm._bu_self_render = true; $w.html(html); frm._bu_self_render = false; }

function bu_inr(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  var num = Number(n);
  if (num === 0) return '0';
  var sign = num < 0 ? '-' : '';
  var abs = Math.abs(num);
  var int = Math.floor(abs);
  var dec = abs - int;
  var s = String(int);
  var lastThree = s.length > 3 ? s.slice(-3) : s;
  var rest = s.length > 3 ? s.slice(0, -3) : '';
  if (rest) lastThree = ',' + lastThree;
  rest = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',');
  var out = sign + rest + lastThree;
  if (dec > 0) out += '.' + Math.round(dec * 100).toString().padStart(2,'0');
  return out;
}

function bu_num(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  var v = Number(n);
  if (Math.abs(v - Math.round(v)) < 1e-9) return String(Math.round(v));
  return v.toFixed(2);
}

function bu_round(n, places) {
  var p = Math.pow(10, places || 2);
  return Math.round(Number(n) * p) / p;
}

function bu_esc(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/[&<>"']/g, function(c){ return { '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]; });
}

function bu_parseQ(q) {
  if (typeof q === 'number') return q;
  if (typeof q !== 'string') return 0;
  var m = q.match(/(\d+)/);
  return m ? Number(m[1]) : 0;
}

function bu_mode(arr) {
  var c = {}, best = arr[0], n = 0;
  arr.forEach(function(v){ c[v] = (c[v]||0) + 1; if (c[v] > n) { n = c[v]; best = v; } });
  return best;
}
