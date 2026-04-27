// ============================================================================
// LIC Budget Allocation v9 — New column layout with convergence
// Client Script for Project proposal
// ============================================================================

frappe.ui.form.on('Project proposal', {
  refresh(frm) {
    if (!frm.doc.__islocal) setup_budget_tab(frm);
  },
  onload(frm) {
    if (!frm.doc.__islocal) setup_budget_tab(frm);
  },
  after_save(frm) {
    // Clear render cache so the budget tab re-renders on next refresh.
    // This ensures newly added activities (custom_activity rows) appear
    // without requiring a hard page reload.
    frm._ab_rendered_for = null;
  }
});

function setup_budget_tab(frm) {
  if (frm._ab_rendering) return;
  const $w = frm.fields_dict.cummulative_budget && frm.fields_dict.cummulative_budget.$wrapper;
  if (!$w || !$w.length) return;

  // Monkey-patch: block external .html() calls, allow our own via flag
  if (!$w._ab_patched) {
    const origHtml = $.fn.html;
    $w.html = function() {
      if (arguments.length === 0) return origHtml.apply(this, arguments);
      if (frm._ab_self_render) return origHtml.apply(this, arguments);
      return this;
    };
    $w._ab_patched = true;
  }

  // Skip re-render if already rendered for this document
  if (frm._ab_rendered_for === frm.doc.name) return;

  ab_render(frm).catch(function(err) { console.error('[AB] render error:', err); });
}

// ============================================================================
// MAIN RENDER
// ============================================================================

async function ab_render(frm) {
  frm._ab_rendering = true;
  const $w = frm.fields_dict.cummulative_budget.$wrapper;

  try {
    frm._ab_self_render = true;
    $w.html('<div style="padding:24px;text-align:center;color:#888;">Loading budget data…</div>');
    frm._ab_self_render = false;

    // Step 1: Parallel fetch of all reference data + PBP list + Units master
    const [quarters, years, bhList, sbhList, fsList, pbpList, unitsList] = await Promise.all([
      ab_fetchQuarters(frm),
      ab_fetchYears(frm),
      ab_getList('Budget heads', ['name', 'budget_head_name']),
      ab_getList('Sub Budget Head', ['name', 'sub_budget_head', 'budget_head']),
      ab_getList('Fund Sources', ['name', 'source_name']),
      ab_getList('Project Budget Planning', ['name', 'description', 'budget_head', 'sub_budget_head', 'fund_source', 'total_planned_budget', 'assumption', 'custom_task_details', 'custom_total_lic_contribution', 'custom_govt_contribution', 'custom_benf_contribution'], { project_proposal: frm.doc.name }, 200, 'budget_head asc, description asc'),
      ab_getList('Units', ['name', 'unit_name'], null, 200, 'unit_name asc')
    ]);

    if (!quarters.length || !years.length) {
      frm._ab_self_render = true;
      $w.html('<div style="padding:24px;text-align:center;color:#999;">No quarters or years configured for this proposal.</div>');
      frm._ab_self_render = false;
      return;
    }

    // Step 2: Build lookup maps
    var bhMap = {};  bhList.forEach(function(b) { bhMap[b.name] = b.budget_head_name; });
    var sbhMap = {}; sbhList.forEach(function(s) { sbhMap[s.name] = s.sub_budget_head; });
    var fsMap = {};  fsList.forEach(function(f) { fsMap[f.name] = f.source_name; });

    // Reverse map: sub_budget_head name → { sbhId, bhId }
    var sbhRevMap = {};
    sbhList.forEach(function(s) { sbhRevMap[s.sub_budget_head] = { sbhId: s.name, bhId: s.budget_head }; });

    // Step 3: Fetch full PBP records (with planning_table children) in batches
    var pbpFull = await ab_fetchAllFull(pbpList.map(function(r) { return r.name; }));

    // Merge list-level fields into full records
    pbpList.forEach(function(r) {
      if (pbpFull[r.name]) {
        pbpFull[r.name]._list = r;
      }
    });

    // Step 3b: Fetch Activity KPI selections for auto-populate
    var activityKPIs = await ab_fetchActivityKPIs(frm);

    // Step 4: Organize into display structures
    var progData = ab_organizeProgData(pbpFull, bhMap, sbhMap, fsMap, quarters, years, activityKPIs);
    var nonProgData = ab_organizeNonProgData(pbpFull, bhMap, sbhMap, fsMap, quarters, years);

    // Step 5: Build HTML
    var html = ab_buildHTML(frm, quarters, years, progData, nonProgData, unitsList);

    frm._ab_self_render = true;
    $w.html(html);
    frm._ab_self_render = false;

    // Step 6: Wire up events
    ab_attachEvents(frm, quarters, years, bhMap, sbhMap, sbhRevMap, fsMap, pbpFull, progData, nonProgData, unitsList);

    // Mark as rendered so refresh events don't re-render
    frm._ab_rendered_for = frm.doc.name;

    // Hide Budget Summary tab (not yet implemented)
    ab_hideBudgetSummaryTab(frm);

  } catch (err) {
    console.error('[AB] Error:', err);
    frm._ab_self_render = true;
    $w.html('<div style="padding:20px;background:#ffebee;color:#c62828;border-radius:4px;">Error loading budget: ' + (err.message || err) + '</div>');
    frm._ab_self_render = false;
  } finally {
    frm._ab_rendering = false;
  }
}

// ============================================================================
// DATA FETCHING
// ============================================================================

function ab_fetchQuarters(frm) {
  var q = (frm.doc.quaterly_project || []).map(function(x) {
    return { year: x.year, quarter: x.quarter, year_sequence: x.year_sequence || 0, quarter_sequence: x.quarter_sequence || 0, start_date: x.start_date, end_date: x.end_date };
  });
  q.sort(function(a, b) { return a.year_sequence !== b.year_sequence ? a.year_sequence - b.year_sequence : a.quarter_sequence - b.quarter_sequence; });
  return Promise.resolve(q);
}

function ab_fetchYears(frm) {
  var y = (frm.doc.annual_project || []).map(function(x) {
    return { year: x.year, year_sequence: x.year_sequence || 0 };
  });
  y.sort(function(a, b) { return a.year_sequence - b.year_sequence; });
  return Promise.resolve(y);
}

function ab_getList(doctype, fields, filters, limit, orderBy) {
  var args = { doctype: doctype, fields: fields, limit_page_length: limit || 100 };
  if (filters) args.filters = filters;
  if (orderBy) args.order_by = orderBy;
  return frappe.call({ method: 'frappe.client.get_list', args: args })
    .then(function(r) { return r.message || []; })
    .catch(function(e) { console.warn('[AB] getList error for ' + doctype + ':', e); return []; });
}

async function ab_fetchAllFull(names) {
  var results = {};
  var batch = 10;
  for (var i = 0; i < names.length; i += batch) {
    var chunk = names.slice(i, i + batch);
    var promises = chunk.map(function(n) {
      return frappe.call({ method: 'frappe.client.get', args: { doctype: 'Project Budget Planning', name: n } })
        .then(function(r) { if (r.message) results[n] = r.message; })
        .catch(function(e) { console.warn('[AB] fetchFull error:', n, e); });
    });
    await Promise.all(promises);
  }
  return results;
}

// ============================================================================
// ACTIVITY KPI AUTO-POPULATE
// ============================================================================

async function ab_fetchActivityKPIs(frm) {
  // Read the proposal's custom_activity child table (Activity Child rows)
  var activityRows = frm.doc.custom_activity || [];
  if (!activityRows.length) return [];

  // Fetch Activity Master details for each selected activity
  var actMasterIds = activityRows.map(function(r) { return r.activity; }).filter(Boolean);
  if (!actMasterIds.length) return [];

  // Fetch Activity Master records with their UoM links
  var activityMasters = [];
  var batch = 10;
  for (var i = 0; i < actMasterIds.length; i += batch) {
    var chunk = actMasterIds.slice(i, i + batch);
    var promises = chunk.map(function(actId) {
      return frappe.call({
        method: 'frappe.client.get',
        args: { doctype: 'Activity Master', name: actId }
      }).then(function(r) {
        if (r.message) activityMasters.push(r.message);
      }).catch(function(e) {
        console.warn('[AB] fetchActivityMaster error:', actId, e);
      });
    });
    await Promise.all(promises);
  }

  // For each activity master, resolve the UoM name from the Units DocType
  var uomIds = activityMasters.map(function(a) { return a.custom_unit_of_measurement; }).filter(Boolean);
  var uniqueUomIds = uomIds.filter(function(v, i, arr) { return arr.indexOf(v) === i; });
  var uomMap = {};

  if (uniqueUomIds.length) {
    var uomPromises = uniqueUomIds.map(function(uid) {
      return frappe.call({
        method: 'frappe.client.get',
        args: { doctype: 'Units', name: uid }
      }).then(function(r) {
        if (r.message) uomMap[uid] = r.message.unit_name;
      }).catch(function(e) {
        console.warn('[AB] fetchUnit error:', uid, e);
      });
    });
    await Promise.all(uomPromises);
  }

  // Return enriched activity list
  return activityMasters.map(function(a) {
    return {
      actId: a.name,
      activityName: a.activity_name,
      uomId: a.custom_unit_of_measurement || '',
      uomName: uomMap[a.custom_unit_of_measurement] || ''
    };
  });
}

// ============================================================================
// DATA ORGANIZATION
// ============================================================================

function ab_normFS(sourceName) {
  if (!sourceName) return 'lic';
  var l = sourceName.toLowerCase();
  if (l.indexOf('lic') >= 0 || l.indexOf('hfl') >= 0) return 'lic';
  if (l.indexOf('gov') >= 0) return 'govt';
  if (l.indexOf('ben') >= 0) return 'benf';
  return 'lic';
}

function ab_qKey(year, quarter) { return year + '::' + quarter; }

function ab_organizeProgData(pbpFull, bhMap, sbhMap, fsMap, quarters, years, activityKPIs) {
  // Build quarter index: qKey → array position
  var qIndex = {};
  quarters.forEach(function(q, i) { qIndex[ab_qKey(q.year, q.quarter)] = i; });

  // Group PBP records by activity description (programmatic = sub-budget head 'Programmatic Costs')
  var activities = {};
  var activityOrder = [];

  Object.keys(pbpFull).forEach(function(pbpName) {
    var rec = pbpFull[pbpName];
    var sbhName = sbhMap[rec.sub_budget_head] || '';
    if (sbhName !== 'Programmatic Costs') return;

    var desc = rec.description || pbpName;

    if (!activities[desc]) {
      activities[desc] = {
        description: desc,
        // Prefer the new custom_task_details field for the long task description;
        // fall back to assumption (legacy) so old data still renders.
        assumption: rec.custom_task_details || rec.assumption || '',
        uomName: 'Numbers',
        unit_cost: 0,
        total_units: 0,
        govt_contribution: 0,
        benf_contribution: 0,
        remarks: '',
        quarters: {},
        pbpName: pbpName
      };
      activityOrder.push(desc);

      // Initialize all quarters to 0
      for (var i = 0; i < quarters.length; i++) {
        activities[desc].quarters[i] = { units: 0 };
      }
    }

    var act = activities[desc];
    act.pbpName = pbpName;

    // Populate quarters from planning_table (sum LIC units as "units to cover")
    (rec.planning_table || []).forEach(function(row) {
      var qi = qIndex[ab_qKey(row.year, row.quarter)];
      if (qi !== undefined) {
        act.quarters[qi] = { units: row.unit || 0 };
        if (row.unit_cost && !act.unit_cost) act.unit_cost = row.unit_cost;
      }
    });

    // Load convergence breakdown from custom fields (added Apr 2026 — see HANDOVER notes).
    // PBP doctype now stores LIC HFL / Govt / Beneficiary split per row.
    if (rec.custom_govt_contribution != null) act.govt_contribution = rec.custom_govt_contribution;
    if (rec.custom_benf_contribution != null) act.benf_contribution = rec.custom_benf_contribution;
    if (rec.custom_total_lic_contribution != null) act.lic_contribution = rec.custom_total_lic_contribution;

    // Derive total_units from PBP's total_planned_budget / unit_cost (so the input shows
    // the planner's intended figure on first paint). Falls back to sum of quarter units
    // if total_planned_budget isn't set yet.
    if (rec.total_planned_budget && act.unit_cost) {
      act.total_units = Math.round((rec.total_planned_budget / act.unit_cost) * 100) / 100;
    } else {
      var qSum = 0;
      for (var qi2 = 0; qi2 < quarters.length; qi2++) {
        qSum += (act.quarters[qi2] && act.quarters[qi2].units) || 0;
      }
      act.total_units = qSum;
    }
  });

  // Auto-populate from Activity KPI selections (if any).
  // The Activity KPI child table on the Project Proposal defines which Activity Master entries
  // apply to this proposal. Each one renders as a budget row — pre-existing PBPs (matched by
  // description) are enriched in place; new ones get blank-row stubs for the user to fill.
  if (activityKPIs && activityKPIs.length) {
    activityKPIs.forEach(function(kpi) {
      var desc = kpi.activityName;
      if (!desc) return;

      // Only add if not already present from PBP records
      if (!activities[desc]) {
        activities[desc] = {
          description: desc,
          assumption: '',
          uomName: kpi.uomName || 'Numbers',
          unit_cost: 0,
          total_units: 0,
          govt_contribution: 0,
          benf_contribution: 0,
          remarks: '',
          activityMasterId: kpi.actId,
          autoPopulated: true,
          quarters: {},
          pbpName: null
        };
        activityOrder.push(desc);

        // Initialize all quarters to 0
        for (var i = 0; i < quarters.length; i++) {
          activities[desc].quarters[i] = { units: 0 };
        }
      } else {
        // Activity exists from PBP — enrich with Activity Master ID and UoM
        activities[desc].activityMasterId = kpi.actId;
        activities[desc].uomName = kpi.uomName || activities[desc].uomName;
      }
    });
  }

  return { rows: activityOrder.map(function(d) { return activities[d]; }) };
}

function ab_organizeNonProgData(pbpFull, bhMap, sbhMap, fsMap, quarters, years) {
  var qIndex = {};
  quarters.forEach(function(q, i) { qIndex[ab_qKey(q.year, q.quarter)] = i; });

  var sections = {};
  var sectionOrder = ['Human Resource Costs', 'Administration Costs', 'NGO Management Costs'];

  var nonProgSections = ['Human Resource Costs', 'Administration Costs', 'NGO Management Costs'];

  Object.keys(pbpFull).forEach(function(pbpName) {
    var rec = pbpFull[pbpName];
    var sbhName = sbhMap[rec.sub_budget_head] || '';
    if (nonProgSections.indexOf(sbhName) < 0) return;

    var section = sbhName;
    var desc = rec.description || pbpName;

    if (!sections[section]) sections[section] = {};

    if (!sections[section][desc]) {
      sections[section][desc] = {
        description: desc,
        assumption: rec.assumption || '',
        unit_cost: 0,
        total_units: 0,           // user-editable scalar (separate from quarter sums)
        pbpName: pbpName,
        quarters: {}
      };
      for (var i = 0; i < quarters.length; i++) {
        sections[section][desc].quarters[i] = { units: 0 };
      }
    }

    var item = sections[section][desc];
    item.pbpName = pbpName;

    (rec.planning_table || []).forEach(function(row) {
      var qi = qIndex[ab_qKey(row.year, row.quarter)];
      if (qi !== undefined) {
        item.quarters[qi] = { units: row.unit || 0 };
        if (row.unit_cost && !item.unit_cost) item.unit_cost = row.unit_cost;
      }
    });

    // Derive total_units from PBP's stored total_planned_budget / unit_cost.
    // Falls back to sum of quarter units if no stored total budget (legacy / new rows).
    if (rec.total_planned_budget && item.unit_cost) {
      item.total_units = Math.round((rec.total_planned_budget / item.unit_cost) * 100) / 100;
    } else {
      var qSum = 0;
      for (var qi2 = 0; qi2 < quarters.length; qi2++) qSum += (item.quarters[qi2] && item.quarters[qi2].units) || 0;
      item.total_units = qSum;
    }
  });

  // Convert to ordered array structure — always include all 3 sections
  var result = {};
  sectionOrder.forEach(function(sec) {
    if (sections[sec]) {
      result[sec] = Object.values(sections[sec]);
    } else {
      // Create a default blank row so the section is visible for NGO input
      var blankRow = {
        description: '',
        assumption: '',
        unit_cost: 0,
        total_units: 0,
        pbpName: null,
        quarters: {}
      };
      for (var i = 0; i < quarters.length; i++) {
        blankRow.quarters[i] = { units: 0 };
      }
      result[sec] = [blankRow];
    }
  });
  // Add any extra sections not in the predefined order
  Object.keys(sections).forEach(function(sec) {
    if (!result[sec]) result[sec] = Object.values(sections[sec]);
  });

  return { sections: result };
}

// ============================================================================
// HTML BUILDING
// ============================================================================

function ab_buildHTML(frm, quarters, years, progData, nonProgData, unitsList) {
  return '<div class="ab-container"><style>' + ab_getStyles() + '</style>' +
    '<div class="ab-tabs">' +
      '<button class="ab-tab-btn ab-tab-active" data-tab="programmatic">Programmatic Costs</button>' +
      '<button class="ab-tab-btn" data-tab="non-programmatic">Non-Programmatic Costs</button>' +
    '</div>' +
    '<div class="ab-tab-content" id="ab-programmatic">' +
      ab_buildProgTab(frm, quarters, years, progData) +
      '<div style="padding:10px 0;text-align:right;"><button class="btn btn-sm btn-primary ab-save-prog-btn">Save Programmatic</button></div>' +
    '</div>' +
    '<div class="ab-tab-content ab-hidden" id="ab-non-programmatic">' +
      ab_buildNonProgTab(frm, quarters, years, nonProgData, unitsList) +
      '<div style="padding:10px 0;text-align:right;"><button class="btn btn-sm btn-primary ab-save-nonprog-btn">Save Non-Programmatic</button></div>' +
    '</div>' +
    '<div class="ab-footer">' +
      '<span class="ab-legend"><span style="background:#FFFDE7;padding:2px 6px;border:1px solid #ddd;">■</span> Editable</span>' +
      '<span class="ab-legend"><span style="background:#E8EAF6;padding:2px 6px;border:1px solid #ddd;">■</span> Auto-calculated</span>' +
      '<span class="ab-legend"><span style="background:#FFF9E6;padding:2px 6px;border:1px solid #ddd;">■</span> Grand Total</span>' +
      '<button class="btn btn-sm btn-default ab-download-btn" style="margin-left:auto;border:1px solid #8B1A1A;color:#8B1A1A;"><i class="fa fa-download"></i> Download Budget Sheet</button>' +
      '<span class="ab-saved-indicator" id="ab-saved"></span>' +
    '</div>' +
  '</div>';
}

// ---- Programmatic Tab (v9 layout) ----

function ab_buildProgTab(frm, quarters, years, data) {
  var rows = data.rows || [];

  var html = '<div class="ab-scroll-wrapper"><table class="ab-table ab-prog-table" style="border-collapse: separate; border-spacing: 0;">';

  html += '<thead>';

  // Row 1: Frozen headers (empty for cols 0-4) + Convergence spanning cols 5-9 + Year-grouped quarters + Grand Total + Remarks
  html += '<tr class="ab-header-row-1">';
  html += '<th colspan="5" class="ab-frozen-header" style="position:sticky;left:0;top:0;z-index:20;background:#8B1A1A;color:white;font-weight:700;">Activity Details</th>';
  html += '<th colspan="5" class="ab-convergence-header">Convergence</th>';

  // Year-grouped headers with nested quarter labels
  years.forEach(function(y, yi) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var colspan = yqList.length * 4;
    html += '<th colspan="' + colspan + '" class="ab-year-header-cell">' +
      '<div class="ab-year-top">Year ' + (yi + 1) + '</div>' +
      '<div class="ab-quarter-row">';
    yqList.forEach(function(q) {
      html += '<span class="ab-q-label">' + q.quarter + '</span>';
    });
    html += '</div></th>';
  });

  // Grand Total column spanning all years
  html += '<th colspan="' + years.length + '" class="ab-gt-header-cell">' +
    '<div class="ab-year-top">Grand Total</div>' +
    '<div class="ab-quarter-row">';
  years.forEach(function(y, yi) {
    html += '<span class="ab-q-label">Y' + (yi + 1) + ' Total</span>';
  });
  html += '</div></th>';

  html += '<th class="ab-quarter-header">&nbsp;</th></tr>';

  // Row 2: Column headers
  html += '<tr class="ab-header-row-2">';
  html += '<th class="ab-frozen ab-sr-hdr ab-frozen-last" style="left:0px;width:35px;min-width:35px;">Sr.</th>';
  html += '<th class="ab-frozen ab-col-hdr" style="left:35px;width:150px;min-width:150px;max-width:150px;">Activity</th>';
  html += '<th class="ab-frozen ab-col-hdr" style="left:185px;width:120px;min-width:120px;max-width:120px;">Task Details</th>';
  html += '<th class="ab-frozen ab-col-hdr" style="left:305px;width:55px;min-width:55px;max-width:55px;">UoM</th>';
  html += '<th class="ab-frozen ab-col-hdr ab-frozen-last" style="left:360px;width:75px;min-width:75px;max-width:75px;">Unit Cost</th>';
  html += '<th class="ab-col-hdr" style="width:70px;min-width:70px;">Total Units</th>';
  html += '<th class="ab-col-hdr" style="width:85px;min-width:85px;">Total Cost</th>';
  html += '<th class="ab-col-hdr" style="width:100px;min-width:100px;">LIC HFL Contribution</th>';
  html += '<th class="ab-col-hdr" style="width:85px;min-width:85px;">Govt Contrib (₹)</th>';
  html += '<th class="ab-col-hdr" style="width:85px;min-width:85px;">Benf Contrib (₹)</th>';

  // Quarterly column headers (per quarter)
  quarters.forEach(function(q, qi) {
    var qClass = qi % 2 === 0 ? 'ab-q-odd' : 'ab-q-even';
    html += '<th class="ab-subcol-hdr ' + qClass + '" style="width:65px;min-width:65px;">Units</th>';
    html += '<th class="ab-subcol-hdr ' + qClass + '" style="width:75px;min-width:75px;">Cost</th>';
    html += '<th class="ab-subcol-hdr ' + qClass + '" style="width:75px;min-width:75px;">LIC HFL</th>';
    html += '<th class="ab-subcol-hdr ' + qClass + '" style="width:65px;min-width:65px;">Benf</th>';
  });

  // Year total columns
  years.forEach(function(y, yi) {
    html += '<th class="ab-gt-subcol" style="width:80px;min-width:80px;">Y' + (yi + 1) + ' Total</th>';
  });

  html += '<th class="ab-remarks-hdr" style="width:100px;min-width:100px;">Remarks</th>';
  html += '</tr></thead><tbody>';

  // Data rows
  rows.forEach(function(row, idx) {
    html += ab_buildProgRow(row, quarters, years, idx);
  });

  // Grand total row
  html += ab_buildProgGrandTotal(rows, quarters, years);

  html += '</tbody></table></div>';
  return html;
}

function ab_buildProgRow(row, quarters, years, idx) {
  var html = '<tr class="ab-data-row" data-idx="' + idx + '">';

  // Frozen columns (0-4): Sr, Activity, Task Details, UoM, Unit Cost
  html += '<td class="ab-frozen ab-sr" style="left:0px;width:35px;min-width:35px;max-width:35px;">' + (idx + 1) + '</td>';
  html += '<td class="ab-frozen" style="left:35px;width:150px;min-width:150px;max-width:150px;text-align:left;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="' + ab_he(row.description) + '">' + ab_he(row.description) + '</td>';
  html += '<td class="ab-frozen ab-editable" style="left:185px;width:120px;min-width:120px;max-width:120px;text-align:left;"><input type="text" class="ab-inp ab-task-inp" style="width:110px;" data-idx="' + idx + '" value="' + ab_he(row.assumption || '') + '" placeholder="Task details..." /></td>';
  html += '<td class="ab-frozen" style="left:305px;width:55px;min-width:55px;max-width:55px;overflow:hidden;text-overflow:ellipsis;">' + ab_he(row.uomName || 'Numbers') + '</td>';
  html += '<td class="ab-frozen ab-editable ab-frozen-last" style="left:360px;width:75px;min-width:75px;max-width:75px;"><input type="number" class="ab-inp ab-uc-inp" style="width:65px;" data-idx="' + idx + '" value="' + (row.unit_cost || 0) + '" /></td>';

  // Non-frozen columns (5-9): Total Units, Total Cost, LIC HFL, Govt, Benf
  html += '<td class="ab-editable" style="width:70px;min-width:70px;"><input type="number" class="ab-inp ab-tu-inp" data-idx="' + idx + '" value="' + (row.total_units || 0) + '" /></td>';

  // Auto-calc: Total Cost = Unit Cost × Total Units
  var tc = (row.unit_cost || 0) * (row.total_units || 0);
  html += '<td class="ab-calc" style="width:85px;min-width:85px;">' + ab_fc(tc) + '</td>';

  // Auto-calc: Total LIC HFL = Total Cost - Govt - Benf
  var tlhfl = tc - (row.govt_contribution || 0) - (row.benf_contribution || 0);
  html += '<td class="ab-calc" style="width:100px;min-width:100px;">' + ab_fc(tlhfl) + '</td>';

  html += '<td class="ab-editable" style="width:85px;min-width:85px;"><input type="number" class="ab-inp ab-govt-inp" data-idx="' + idx + '" value="' + (row.govt_contribution || 0) + '" /></td>';
  html += '<td class="ab-editable" style="width:85px;min-width:85px;"><input type="number" class="ab-inp ab-benf-inp" data-idx="' + idx + '" value="' + (row.benf_contribution || 0) + '" /></td>';

  // Per-quarter columns (4 per quarter)
  quarters.forEach(function(q, qi) {
    var qUnits = (row.quarters[qi] || {}).units || 0;
    var qCost = (row.unit_cost || 0) * qUnits;
    var qLicHfl = tlhfl > 0 && (row.total_units || 0) > 0 ? (tlhfl / (row.total_units || 0)) * qUnits : 0;
    var qBenf = qCost - qLicHfl;
    var qClass = qi % 2 === 0 ? 'ab-q-odd' : 'ab-q-even';

    html += '<td class="ab-editable ' + qClass + '" style="width:65px;"><input type="number" class="ab-inp ab-q-units-inp" style="width:55px;" data-idx="' + idx + '" data-qi="' + qi + '" value="' + qUnits + '" /></td>';
    html += '<td class="ab-calc ' + qClass + '" style="width:75px;">' + ab_fc(qCost) + '</td>';
    html += '<td class="ab-calc ' + qClass + '" style="width:75px;">' + ab_fc(qLicHfl) + '</td>';
    html += '<td class="ab-calc ' + qClass + '" style="width:65px;">' + ab_fc(qBenf) + '</td>';
  });

  // Year total columns (auto-calculated)
  var yuc = row.unit_cost || 0;
  years.forEach(function(y) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yCost = 0;
    yqList.forEach(function(q) {
      var qi = quarters.indexOf(q);
      var qUnits = (row.quarters[qi] || {}).units || 0;
      yCost += yuc * qUnits;
    });
    html += '<td class="ab-gt-cell" style="width:80px;">' + ab_fc(yCost) + '</td>';
  });

  html += '<td class="ab-editable"><input type="text" class="ab-inp ab-remarks-inp" data-idx="' + idx + '" value="' + ab_he(row.remarks || '') + '" placeholder="Remarks..." /></td>';
  html += '</tr>';
  return html;
}

function ab_buildProgGrandTotal(rows, quarters, years) {
  // Compute initial grand-total figures from the loaded row data so the
  // first paint shows real numbers (no need to wait for a user input event).
  var sumTU = 0, sumTC = 0, sumLIC = 0, sumGovt = 0, sumBenf = 0;
  var qUnits = quarters.map(function() { return 0; });
  var qCosts = quarters.map(function() { return 0; });
  var qLics  = quarters.map(function() { return 0; });
  var qBenfs = quarters.map(function() { return 0; });
  rows.forEach(function(row) {
    var uc = row.unit_cost || 0;
    var tu = row.total_units || 0;
    sumTU += tu;
    sumTC += uc * tu;
    var govt = row.govt_contribution || 0;
    var benf = row.benf_contribution || 0;
    sumGovt += govt;
    sumBenf += benf;
    sumLIC  += (uc * tu) - govt - benf;
    quarters.forEach(function(q, qi) {
      var u = (row.quarters[qi] || {}).units || 0;
      var c = u * uc;
      var qLicShare = (uc * tu) > 0 ? (c * ((uc*tu) - govt - benf) / (uc*tu)) : 0;
      var qBenfShare = (uc * tu) > 0 ? (c * benf / (uc*tu)) : 0;
      qUnits[qi] += u;
      qCosts[qi] += c;
      qLics[qi]  += qLicShare;
      qBenfs[qi] += qBenfShare;
    });
  });

  var yearSet = {};
  quarters.forEach(function(q) { yearSet[q.year_sequence] = true; });
  var yearCount = Object.keys(yearSet).length;
  var yearTotals = years.map(function(y) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var t = 0;
    rows.forEach(function(row) {
      var uc = row.unit_cost || 0;
      yqList.forEach(function(q) {
        var qi = quarters.indexOf(q);
        var u = (row.quarters[qi] || {}).units || 0;
        t += uc * u;
      });
    });
    return t;
  });

  var html = '<tr class="ab-grand-total-row">';

  // Frozen columns (5): Sr, Activity, Task Details, UoM, Unit Cost
  html += '<td class="ab-frozen ab-sr" style="left:0px;width:35px;min-width:35px;max-width:35px;">GT</td>';
  html += '<td class="ab-frozen" style="left:35px;width:150px;min-width:150px;max-width:150px;text-align:left;font-weight:700;">GRAND TOTAL</td>';
  html += '<td class="ab-frozen" style="left:185px;width:120px;min-width:120px;max-width:120px;"></td>';
  html += '<td class="ab-frozen" style="left:305px;width:55px;min-width:55px;max-width:55px;"></td>';
  html += '<td class="ab-frozen ab-frozen-last" style="left:360px;width:75px;min-width:75px;max-width:75px;"></td>';

  // Non-frozen columns: Total Units, Total Cost, LIC HFL, Govt, Benf
  html += '<td class="ab-gt-cell" style="width:70px;min-width:70px;">' + sumTU + '</td>';
  html += '<td class="ab-gt-cell" style="width:85px;min-width:85px;">' + ab_fc(sumTC) + '</td>';
  html += '<td class="ab-gt-cell" style="width:100px;min-width:100px;">' + ab_fc(sumLIC) + '</td>';
  html += '<td class="ab-gt-cell" style="width:85px;min-width:85px;">' + ab_fc(sumGovt) + '</td>';
  html += '<td class="ab-gt-cell" style="width:85px;min-width:85px;">' + ab_fc(sumBenf) + '</td>';

  // Per-quarter totals
  quarters.forEach(function(q, qi) {
    var qClass = qi % 2 === 0 ? 'ab-q-odd' : 'ab-q-even';
    html += '<td class="ab-gt-cell ' + qClass + '">' + qUnits[qi] + '</td>';
    html += '<td class="ab-gt-cell ' + qClass + '">' + ab_fc(qCosts[qi]) + '</td>';
    html += '<td class="ab-gt-cell ' + qClass + '">' + ab_fc(qLics[qi]) + '</td>';
    html += '<td class="ab-gt-cell ' + qClass + '">' + ab_fc(qBenfs[qi]) + '</td>';
  });

  // Year total columns
  yearTotals.forEach(function(yt) {
    html += '<td class="ab-gt-cell">' + ab_fc(yt) + '</td>';
  });

  html += '<td class="ab-gt-cell"></td>';
  html += '</tr>';
  return html;
}

// ---- Non-Programmatic Tab ----

function ab_buildNonProgTab(frm, quarters, years, data, unitsList) {
  var sections = data.sections || {};
  var sectionKeys = Object.keys(sections);
  if (!sectionKeys.length) return '<div style="padding:24px;text-align:center;color:#999;">No non-programmatic data found.</div>';

  // ONE merged table. Section dividers, section TOTAL rows, "+ Add Row" affordances and the final
  // GRAND TOTAL row all live as full-width rows inside this table's tbody. Single sticky header at top.
  // Section letters (A/B/C). Sr resets per section.
  var sectionLetter = { 'Human Resource Costs': 'A', 'Administration Costs': 'B', 'NGO Management Costs': 'C' };

  var nQuarterCols = quarters.length * 2;
  var nYearCols = years.length;
  var totalCols = 6 + nQuarterCols + nYearCols + 1;

  var html = '<div class="ab-nonprog-wrapper"><div class="ab-nonprog-scroll">';
  html += '<table class="ab-table ab-nonprog-table"><thead>';

  // Header row 1 — Particulars (left frozen merged) · Year N (per year) · Grand Total · (blank for Remarks merged 2 rows)
  html += '<tr class="ab-header-row-1">' +
    '<th colspan="6" class="ab-frozen-header" style="position:sticky;left:0;top:0;z-index:20;background:#8B1A1A;color:white;font-weight:700;">Particulars</th>';
  years.forEach(function(y, yi) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var colspan = yqList.length * 2;
    html += '<th colspan="' + colspan + '" class="ab-year-header-cell">' +
      '<div class="ab-year-top">Year ' + (yi + 1) + '</div>' +
      '<div class="ab-quarter-row">';
    yqList.forEach(function(q) { html += '<span class="ab-q-label">' + q.quarter + '</span>'; });
    html += '</div></th>';
  });
  html += '<th colspan="' + nYearCols + '" class="ab-gt-header-cell">' +
    '<div class="ab-year-top">Grand Total</div>' +
    '<div class="ab-quarter-row">';
  years.forEach(function(y, yi) { html += '<span class="ab-q-label">Y' + (yi + 1) + ' Total</span>'; });
  html += '</div></th>';
  html += '<th class="ab-quarter-header">&nbsp;</th></tr>';

  // Header row 2 — actual column field names (frozen left)
  html += '<tr class="ab-header-row-2">' +
    '<th class="ab-frozen" style="left:0;width:30px;min-width:30px;max-width:30px;">Sr.</th>' +
    '<th class="ab-frozen" style="left:30px;width:160px;min-width:160px;max-width:160px;">Particulars</th>' +
    '<th class="ab-frozen" style="left:190px;width:60px;min-width:60px;max-width:60px;">UoM</th>' +
    '<th class="ab-frozen" style="left:250px;width:75px;min-width:75px;max-width:75px;">Unit Cost</th>' +
    '<th class="ab-frozen ab-calc" style="left:325px;width:70px;min-width:70px;max-width:70px;">Total Units</th>' +
    '<th class="ab-frozen ab-calc ab-frozen-last" style="left:395px;width:85px;min-width:85px;max-width:85px;">Total Cost</th>';
  quarters.forEach(function(q, qi) {
    var qClass = qi % 2 === 0 ? 'ab-q-odd' : 'ab-q-even';
    html += '<th class="ab-subcol ' + qClass + '" style="width:65px;min-width:65px;">Units</th>';
    html += '<th class="ab-subcol ' + qClass + '" style="width:75px;min-width:75px;">Cost</th>';
  });
  years.forEach(function(y, yi) {
    html += '<th class="ab-gt-subcol" style="width:80px;min-width:80px;">Y' + (yi + 1) + ' Total</th>';
  });
  html += '<th class="ab-remarks-hdr" style="width:100px;min-width:100px;">Remarks</th>';
  html += '</tr></thead><tbody>';

  // Body — for each section: banner row, data rows, total row, add-row affordance row
  sectionKeys.forEach(function(secTitle) {
    var letter = sectionLetter[secTitle] || '';
    var rows = sections[secTitle] || [];

    // Section banner row (full-width, dark grey — matches source Excel)
    html += '<tr class="ab-np-section-banner" data-section="' + ab_he(secTitle) + '">' +
      '<td colspan="' + totalCols + '" class="ab-np-section-banner-cell">' +
        '<span class="ab-np-section-letter">' + (letter ? letter + '   ' : '') + '</span>' +
        ab_he(secTitle) +
      '</td>' +
    '</tr>';

    // Data rows (Sr resets per section: idx is 0-based within rows[])
    rows.forEach(function(row, idx) {
      html += ab_buildNonProgRow(row, quarters, years, idx, secTitle, unitsList);
    });

    // Section TOTAL (X) row
    html += ab_buildNonProgSectionTotal(rows, quarters, years, secTitle, letter);

    // + Add Row affordance row (clearly scoped per section)
    html += '<tr class="ab-np-addrow-row" data-section="' + ab_he(secTitle) + '">' +
      '<td colspan="' + totalCols + '" class="ab-np-addrow-cell">' +
        '<button class="btn btn-xs btn-default ab-add-row-btn" data-section="' + ab_he(secTitle) + '" type="button">' +
          '+ Add Row to ' + ab_he(secTitle) +
        '</button>' +
      '</td>' +
    '</tr>';
  });

  // Final GRAND TOTAL row (yellow)
  html += ab_buildNonProgGrandTotal(sections, quarters, years, sectionKeys, totalCols);

  html += '</tbody></table>';
  html += '</div></div>'; // close ab-nonprog-scroll, ab-nonprog-wrapper
  return html;
}

// Build the final GRAND TOTAL row (sums across all sections)
function ab_buildNonProgGrandTotal(sections, quarters, years, sectionKeys, totalCols) {
  // Aggregate across sections
  var grandUnits = 0, grandCost = 0;
  var qSums = []; // [{units, cost}, ...] per quarter
  for (var qi = 0; qi < quarters.length; qi++) qSums.push({ units: 0, cost: 0 });
  var ySums = []; // per year cost
  for (var yi = 0; yi < years.length; yi++) ySums.push(0);

  sectionKeys.forEach(function(secTitle) {
    (sections[secTitle] || []).forEach(function(row) {
      var uc = row.unit_cost || 0;
      var rowTU = (typeof row.total_units === 'number') ? row.total_units : 0;
      if (!rowTU) {
        quarters.forEach(function(q, qi) { rowTU += (row.quarters[qi] || {}).units || 0; });
      }
      // Grand totals for the frozen left "Total Units" / "Total Cost" cells (independent of quarter cells)
      grandUnits += rowTU;
      grandCost  += uc * rowTU;
      // Quarter totals + year totals are still based on per-quarter inputs
      quarters.forEach(function(q, qi) {
        var u = (row.quarters[qi] || {}).units || 0;
        qSums[qi].units += u;
        qSums[qi].cost  += u * uc;
      });
      years.forEach(function(y, yi) {
        var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
        yqList.forEach(function(q) {
          var qi = quarters.indexOf(q);
          var u = (row.quarters[qi] || {}).units || 0;
          ySums[yi] += u * uc;
        });
      });
    });
  });

  var html = '<tr class="ab-np-grand-total">' +
    '<td class="ab-frozen" style="left:0;width:30px;min-width:30px;max-width:30px;"></td>' +
    '<td class="ab-frozen" style="left:30px;width:160px;min-width:160px;max-width:160px;text-align:left;font-weight:700;">GRAND TOTAL</td>' +
    '<td class="ab-frozen" style="left:190px;width:60px;min-width:60px;max-width:60px;"></td>' +
    '<td class="ab-frozen" style="left:250px;width:75px;min-width:75px;max-width:75px;"></td>' +
    '<td class="ab-frozen ab-calc" style="left:325px;width:70px;min-width:70px;max-width:70px;">' + grandUnits + '</td>' +
    '<td class="ab-frozen ab-calc ab-frozen-last" style="left:395px;width:85px;min-width:85px;max-width:85px;">' + ab_fc(grandCost) + '</td>';
  quarters.forEach(function(q, qi) {
    var qClass = qi % 2 === 0 ? 'ab-q-odd' : 'ab-q-even';
    html += '<td class="' + qClass + '">' + qSums[qi].units + '</td>';
    html += '<td class="' + qClass + '">' + ab_fc(qSums[qi].cost) + '</td>';
  });
  years.forEach(function(y, yi) {
    html += '<td>' + ab_fc(ySums[yi]) + '</td>';
  });
  html += '<td></td></tr>';
  return html;
}

function ab_buildNonProgRow(row, quarters, years, idx, secTitle, unitsList) {
  var uc = row.unit_cost || 0;
  var qData = row.quarters || {};
  var pbpId = row.pbpName || '';

  // Build UoM dropdown
  var uomOptions = '<option value="">-- Select --</option>';
  (unitsList || []).forEach(function(u) {
    var sel = (row.assumption === u.unit_name) ? ' selected' : '';
    uomOptions += '<option value="' + ab_he(u.unit_name) + '"' + sel + '>' + ab_he(u.unit_name) + '</option>';
  });

  var html = '<tr class="ab-data-row" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" data-pbp="' + ab_he(pbpId) + '">' +
    '<td class="ab-frozen ab-sr" style="left:0;width:30px;min-width:30px;max-width:30px;">' + (idx + 1) + '</td>' +
    '<td class="ab-frozen ab-editable" style="left:30px;width:160px;min-width:160px;max-width:160px;text-align:left;"><input type="text" class="ab-inp ab-desc-inp" style="width:150px;text-align:left;" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" value="' + ab_he(row.description) + '" placeholder="Enter particulars..." /></td>' +
    '<td class="ab-frozen ab-editable" style="left:190px;width:60px;min-width:60px;max-width:60px;"><select class="ab-inp ab-uom-sel" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" style="width:55px;text-align:left;">' + uomOptions + '</select></td>' +
    '<td class="ab-frozen ab-editable" style="left:250px;width:75px;min-width:75px;max-width:75px;"><input type="number" class="ab-inp ab-uc-inp" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" value="' + uc + '" /></td>';

  // Total Units (user-editable, like UoM and Unit Cost) — frozen
  var totalUnitsInit = (typeof row.total_units === 'number') ? row.total_units : 0;
  if (!totalUnitsInit) {
    // Initial fallback: derive from quarter sum so existing data isn't lost on first paint
    quarters.forEach(function(q, qi) { totalUnitsInit += (qData[qi] || {}).units || 0; });
  }
  html += '<td class="ab-frozen ab-editable" style="left:325px;width:70px;min-width:70px;max-width:70px;">' +
    '<input type="number" class="ab-inp ab-tu-inp" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" value="' + totalUnitsInit + '" step="any" />' +
    '</td>';

  // Total Cost (auto-calc from Unit Cost × Total Units) — frozen last
  var totalCost = uc * totalUnitsInit;
  html += '<td class="ab-frozen ab-calc ab-frozen-last" style="left:395px;width:85px;min-width:85px;max-width:85px;">' + ab_fc(totalCost) + '</td>';

  // Per-quarter data
  quarters.forEach(function(q, qi) {
    var u = (qData[qi] || {}).units || 0;
    var c = u * uc;
    var qClass = qi % 2 === 0 ? 'ab-q-odd' : 'ab-q-even';

    html += '<td class="ab-editable ' + qClass + '" style="width:65px;"><input type="number" class="ab-inp ab-np-inp" style="width:55px;" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" data-qi="' + qi + '" value="' + u + '" /></td>' +
            '<td class="ab-calc ' + qClass + '" style="width:75px;">' + ab_fc(c) + '</td>';
  });

  // Year total columns (auto-calculated)
  years.forEach(function(y) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yCost = 0;
    yqList.forEach(function(q) {
      var qi = quarters.indexOf(q);
      var u = (qData[qi] || {}).units || 0;
      yCost += u * uc;
    });
    html += '<td class="ab-gt-cell" style="width:80px;">' + ab_fc(yCost) + '</td>';
  });

  // Remarks
  html += '<td class="ab-editable"><input type="text" class="ab-inp ab-remarks-inp" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" value="' + ab_he(row.remarks || '') + '" placeholder="Remarks..." /></td>';

  html += '</tr>';
  return html;
}

function ab_buildNonProgSectionTotal(rows, quarters, years, secTitle, letter) {
  var label = letter ? ('TOTAL (' + letter + ')') : 'Section Total';
  var html = '<tr class="ab-section-total-row" data-section="' + ab_he(secTitle) + '">' +
    '<td class="ab-frozen" style="left:0;width:30px;min-width:30px;max-width:30px;"></td>' +
    '<td class="ab-frozen" style="left:30px;width:160px;min-width:160px;max-width:160px;text-align:left;font-weight:700;">' + ab_he(label) + '</td>' +
    '<td class="ab-frozen" style="left:190px;width:60px;min-width:60px;max-width:60px;"></td>' +
    '<td class="ab-frozen" style="left:250px;width:75px;min-width:75px;max-width:75px;"></td>';

  // Total Units (now derives from each row's stored total_units, not quarter sum)
  var totalUnits = 0;
  rows.forEach(function(row) {
    var tu = (typeof row.total_units === 'number') ? row.total_units : 0;
    if (!tu) {
      // Initial paint fallback when total_units hasn't been derived yet
      quarters.forEach(function(q, qi) { tu += (row.quarters[qi] || {}).units || 0; });
    }
    totalUnits += tu;
  });
  html += '<td class="ab-frozen ab-calc" style="left:325px;width:70px;min-width:70px;max-width:70px;">' + totalUnits + '</td>';

  // Total Cost = sum(unit_cost × total_units) per row
  var totalCost = 0;
  rows.forEach(function(row) {
    var uc = row.unit_cost || 0;
    var tu = (typeof row.total_units === 'number') ? row.total_units : 0;
    if (!tu) { quarters.forEach(function(q, qi) { tu += (row.quarters[qi] || {}).units || 0; }); }
    totalCost += uc * tu;
  });
  html += '<td class="ab-frozen ab-calc ab-frozen-last" style="left:395px;width:85px;min-width:85px;max-width:85px;">' + ab_fc(totalCost) + '</td>';

  // Per-quarter totals
  quarters.forEach(function(q, qi) {
    var qUnits = 0, qCost = 0;
    rows.forEach(function(row) {
      var u = (row.quarters[qi] || {}).units || 0;
      qUnits += u;
      qCost += u * (row.unit_cost || 0);
    });
    var qClass = qi % 2 === 0 ? 'ab-q-odd' : 'ab-q-even';
    html += '<td class="ab-gt-cell ' + qClass + '">' + qUnits + '</td>';
    html += '<td class="ab-gt-cell ' + qClass + '">' + ab_fc(qCost) + '</td>';
  });

  // Year total columns
  years.forEach(function(y) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yTotal = 0;
    rows.forEach(function(row) {
      var uc = row.unit_cost || 0;
      yqList.forEach(function(q) {
        var qi = quarters.indexOf(q);
        var u = (row.quarters[qi] || {}).units || 0;
        yTotal += u * uc;
      });
    });
    html += '<td class="ab-gt-cell">' + ab_fc(yTotal) + '</td>';
  });

  html += '<td class="ab-gt-cell"></td>';

  html += '</tr>';
  return html;
}

// ============================================================================
// EVENTS
// ============================================================================

function ab_attachEvents(frm, quarters, years, bhMap, sbhMap, sbhRevMap, fsMap, pbpFull, progData, nonProgData, unitsList) {
  // Tab switching
  document.querySelectorAll('.ab-tab-btn').forEach(function(btn) {
    btn.addEventListener('click', function(e) {
      e.preventDefault();
      document.querySelectorAll('.ab-tab-btn').forEach(function(b) { b.classList.remove('ab-tab-active'); });
      document.querySelectorAll('.ab-tab-content').forEach(function(c) { c.classList.add('ab-hidden'); });
      this.classList.add('ab-tab-active');
      document.getElementById('ab-' + this.dataset.tab).classList.remove('ab-hidden');
    });
  });

  // (Section banners are inline rows now — no toggle handler needed.)

  // Input change handlers — recalculate row on change (no auto-save, use Save button)
  document.querySelectorAll('.ab-inp').forEach(function(inp) {
    inp.addEventListener('change', function() {
      ab_recalcRow(this, quarters, years, progData, nonProgData);
    });
    inp.addEventListener('input', function() {
      ab_recalcRow(this, quarters, years, progData, nonProgData);
    });
  });

  // Add Row buttons for non-programmatic sections
  document.querySelectorAll('.ab-add-row-btn').forEach(function(btn) {
    btn.addEventListener('click', function(e) {
      e.preventDefault();
      e.stopPropagation();
      var secTitle = this.dataset.section;
      ab_addNonProgRow(frm, secTitle, quarters, years, sbhRevMap, nonProgData, progData, unitsList);
    });
  });

  // Save Non-Programmatic button
  var saveNPBtn = document.querySelector('.ab-save-nonprog-btn');
  if (saveNPBtn) {
    saveNPBtn.addEventListener('click', function(e) {
      e.preventDefault();
      var allRows = Array.from(document.querySelectorAll('.ab-nonprog-table .ab-data-row'));
      var rowsToSave = allRows.filter(function(tr) {
        var descInp = tr.querySelector('.ab-desc-inp');
        return descInp && descInp.value.trim();
      });
      if (!rowsToSave.length) { frappe.show_alert({ message: 'No rows to save (add descriptions first)', indicator: 'orange' }); return; }
      frappe.show_alert({ message: 'Saving ' + rowsToSave.length + ' non-programmatic rows...', indicator: 'blue' });
      ab_saveNonProgSequential(frm, rowsToSave, 0, quarters, years, sbhRevMap);
    });
  }

  // Save Programmatic button
  var savePBtn = document.querySelector('.ab-save-prog-btn');
  if (savePBtn) {
    savePBtn.addEventListener('click', function(e) {
      e.preventDefault();
      ab_saveProgData(frm, quarters, years, progData, sbhRevMap);
    });
  }

  // Download Budget Sheet button
  var dlBtn = document.querySelector('.ab-download-btn');
  if (dlBtn) {
    dlBtn.addEventListener('click', function(e) {
      e.preventDefault();
      ab_downloadBudget(frm, quarters, years, progData, nonProgData, unitsList);
    });
  }
}

function ab_recalcRow(inputEl, quarters, years, progData, nonProgData) {
  var tr = inputEl.closest('tr');
  if (!tr) return;

  // Check if this is a non-programmatic row
  var section = tr.dataset.section;
  if (section !== undefined && section !== '') {
    ab_recalcNonProgRow(tr, quarters, years);
    return;
  }

  // For programmatic rows (v9 layout)
  var idx = parseInt(tr.dataset.idx);
  if (!isNaN(idx) && progData.rows[idx]) {
    var row = progData.rows[idx];
    var cells = tr.querySelectorAll('td');

    // Read inputs from frozen columns
    var ucInp = cells[4] && cells[4].querySelector('input');
    var uc = ucInp ? (parseFloat(ucInp.value) || 0) : (row.unit_cost || 0);
    row.unit_cost = uc;

    var tuInp = cells[5] && cells[5].querySelector('input');
    var tu = tuInp ? (parseFloat(tuInp.value) || 0) : (row.total_units || 0);
    row.total_units = tu;

    var govtInp = cells[8] && cells[8].querySelector('input');
    var govt = govtInp ? (parseFloat(govtInp.value) || 0) : (row.govt_contribution || 0);
    row.govt_contribution = govt;

    var benfInp = cells[9] && cells[9].querySelector('input');
    var benf = benfInp ? (parseFloat(benfInp.value) || 0) : (row.benf_contribution || 0);
    row.benf_contribution = benf;

    // Update task details and remarks
    var taskInp = cells[2] && cells[2].querySelector('input');
    if (taskInp) row.assumption = taskInp.value;
    var remarksInp = cells[cells.length - 1] && cells[cells.length - 1].querySelector('input');
    if (remarksInp) row.remarks = remarksInp.value;

    // Recalculate auto-calc cells
    var tc = uc * tu; // Total Cost
    var tlhfl = tc - govt - benf; // Total LIC HFL
    cells[6].textContent = ab_fc(tc);
    cells[7].textContent = ab_fc(tlhfl);

    // Recalculate quarterly cells (starting at column 10)
    var ci = 10;
    quarters.forEach(function(q, qi) {
      var qUnitsInp = cells[ci] && cells[ci].querySelector('input');
      var qUnits = qUnitsInp ? (parseFloat(qUnitsInp.value) || 0) : 0;
      row.quarters[qi] = { units: qUnits };

      var qCost = uc * qUnits;
      var qLicHfl = tu > 0 ? (tlhfl / tu) * qUnits : 0;
      var qBenf = qCost - qLicHfl;

      cells[ci + 1].textContent = ab_fc(qCost);
      cells[ci + 2].textContent = ab_fc(qLicHfl);
      cells[ci + 3].textContent = ab_fc(qBenf);

      ci += 4;
    });

    // Recalculate year total cells (after quarterly cells)
    // ci now points to the first year total cell
    years.forEach(function(y) {
      var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
      var yCost = 0;
      yqList.forEach(function(q) {
        var qi = quarters.indexOf(q);
        var qUnits = (row.quarters[qi] || {}).units || 0;
        yCost += uc * qUnits;
      });
      if (cells[ci]) cells[ci].textContent = ab_fc(yCost);
      ci += 1;
    });

    // Recalculate grand total
    ab_recalcProgGrandTotal(quarters, progData, years);
  }
}

function ab_recalcProgGrandTotal(quarters, progData, years) {
  var table = document.querySelector('.ab-prog-table');
  if (!table) return;
  var gtRow = table.querySelector('.ab-grand-total-row');
  if (!gtRow) return;
  var dataRows = table.querySelectorAll('.ab-data-row');
  if (!dataRows.length) return;

  var gtCells = gtRow.querySelectorAll('td');
  var gci = 6; // Start at Total Cost column

  // Total Cost, Total LIC HFL, Total Govt, Total Benf
  var sumTc = 0, sumLhfl = 0, sumGovt = 0, sumBenf = 0;
  dataRows.forEach(function(dr) {
    var drCells = dr.querySelectorAll('td');
    var tcCell = drCells[6];
    var lhflCell = drCells[7];
    var govtCell = drCells[8] && drCells[8].querySelector('input');
    var benfCell = drCells[9] && drCells[9].querySelector('input');

    if (tcCell && tcCell.textContent) sumTc += parseFloat(tcCell.textContent.replace(/[^0-9.-]/g, '')) || 0;
    if (lhflCell && lhflCell.textContent) sumLhfl += parseFloat(lhflCell.textContent.replace(/[^0-9.-]/g, '')) || 0;
    if (govtCell) sumGovt += parseFloat(govtCell.value) || 0;
    if (benfCell) sumBenf += parseFloat(benfCell.value) || 0;
  });

  gtCells[6].textContent = ab_fc(sumTc);
  gtCells[7].textContent = ab_fc(sumLhfl);
  gtCells[8].textContent = ab_fc(sumGovt);
  gtCells[9].textContent = ab_fc(sumBenf);

  // Per-quarter totals
  var gci_q = 10;
  quarters.forEach(function() {
    var qUnits = 0, qCost = 0, qLicHfl = 0, qBenf = 0;

    dataRows.forEach(function(dr) {
      var drCells = dr.querySelectorAll('td');
      var qUnitsInp = drCells[gci_q] && drCells[gci_q].querySelector('input');
      var qCostCell = drCells[gci_q + 1];
      var qLhflCell = drCells[gci_q + 2];
      var qBenfCell = drCells[gci_q + 3];

      if (qUnitsInp) qUnits += parseFloat(qUnitsInp.value) || 0;
      if (qCostCell && qCostCell.textContent) qCost += parseFloat(qCostCell.textContent.replace(/[^0-9.-]/g, '')) || 0;
      if (qLhflCell && qLhflCell.textContent) qLicHfl += parseFloat(qLhflCell.textContent.replace(/[^0-9.-]/g, '')) || 0;
      if (qBenfCell && qBenfCell.textContent) qBenf += parseFloat(qBenfCell.textContent.replace(/[^0-9.-]/g, '')) || 0;
    });

    gtCells[gci_q].textContent = qUnits;
    gtCells[gci_q + 1].textContent = ab_fc(qCost);
    gtCells[gci_q + 2].textContent = ab_fc(qLicHfl);
    gtCells[gci_q + 3].textContent = ab_fc(qBenf);

    gci_q += 4;
  });

  // Year total cells
  var gci_y = 10 + (quarters.length * 4);
  years.forEach(function(y) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yTotal = 0;
    dataRows.forEach(function(dr) {
      var drCells = dr.querySelectorAll('td');
      var drYCostCell = drCells[gci_y];
      if (drYCostCell && drYCostCell.textContent) {
        yTotal += parseFloat(drYCostCell.textContent.replace(/[^0-9.-]/g, '')) || 0;
      }
    });
    if (gtCells[gci_y]) gtCells[gci_y].textContent = ab_fc(yTotal);
    gci_y += 1;
  });
}

function ab_recalcNonProgRow(tr, quarters, years) {
  var cells = tr.querySelectorAll('td');
  var ucInput = cells[3] && cells[3].querySelector('input');
  var uc = ucInput ? (parseFloat(ucInput.value) || 0) : 0;

  // Recalc Total Units and Total Cost
  // Total Units is now a user input (cell index 4) — read it directly. Total Cost = uc × totalUnits.
  var tuInp = cells[4] && cells[4].querySelector('input');
  var totalUnits = tuInp ? (parseFloat(tuInp.value) || 0) : 0;
  var totalCost = uc * totalUnits;
  cells[5].textContent = ab_fc(totalCost); // Total Cost cell auto-calc

  // Recalc per-quarter costs (independent from Total Units; quarter units × unit cost)
  var ci = 6;
  quarters.forEach(function() {
    var uInp = cells[ci] && cells[ci].querySelector('input');
    var u = uInp ? (parseFloat(uInp.value) || 0) : 0;
    var c = u * uc;
    cells[ci + 1].textContent = ab_fc(c);
    ci += 2;
  });

  // Recalc year total columns
  // ci now points to the first year total cell
  years.forEach(function(y) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yCost = 0;
    var yCi = 6;
    quarters.forEach(function(q, qi) {
      var u = parseInt(cells[yCi].querySelector('input').value) || 0;
      if (yqList.indexOf(q) >= 0) {
        yCost += u * uc;
      }
      yCi += 2;
    });
    if (cells[ci]) cells[ci].textContent = ab_fc(yCost);
    ci += 1;
  });

  ab_recalcNonProgSectionTotal(tr, quarters, years);
  ab_recalcNonProgGrandTotal(quarters, years);
}

// Sum every section's data rows into the GRAND TOTAL row at the bottom of the merged tbody.
function ab_recalcNonProgGrandTotal(quarters, years) {
  var tbody = document.querySelector('.ab-nonprog-table tbody');
  if (!tbody) return;
  var grandRow = tbody.querySelector('.ab-np-grand-total');
  if (!grandRow) return;
  var dataRows = tbody.querySelectorAll('.ab-data-row');
  var gtCells = grandRow.querySelectorAll('td');

  var totalUnits = 0, totalCost = 0;
  dataRows.forEach(function(dr) {
    var drCells = dr.querySelectorAll('td');
    var ucInp = drCells[3] && drCells[3].querySelector('input');
    var uc = ucInp ? (parseFloat(ucInp.value) || 0) : 0;
    var tuInp = drCells[4] && drCells[4].querySelector('input');
    var rowTU = tuInp ? (parseFloat(tuInp.value) || 0) : 0;
    totalUnits += rowTU;
    totalCost  += uc * rowTU;
  });
  gtCells[4].textContent = totalUnits;
  gtCells[5].textContent = ab_fc(totalCost);

  var ci = 6;
  quarters.forEach(function() {
    var qUnits = 0, qCost = 0;
    dataRows.forEach(function(dr) {
      var drCells = dr.querySelectorAll('td');
      var ucInp = drCells[3] && drCells[3].querySelector('input');
      var uc = ucInp ? (parseFloat(ucInp.value) || 0) : 0;
      var uInp = drCells[ci] && drCells[ci].querySelector('input');
      var u = uInp ? (parseInt(uInp.value) || 0) : 0;
      qUnits += u;
      qCost  += u * uc;
    });
    if (gtCells[ci]) gtCells[ci].textContent = qUnits;
    if (gtCells[ci + 1]) gtCells[ci + 1].textContent = ab_fc(qCost);
    ci += 2;
  });

  years.forEach(function(y) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yTotal = 0;
    dataRows.forEach(function(dr) {
      var drCells = dr.querySelectorAll('td');
      var ucInp = drCells[3] && drCells[3].querySelector('input');
      var uc = ucInp ? (parseFloat(ucInp.value) || 0) : 0;
      var qCi = 6;
      quarters.forEach(function(q) {
        if (yqList.indexOf(q) >= 0) {
          var uInp = drCells[qCi] && drCells[qCi].querySelector('input');
          var u = uInp ? (parseInt(uInp.value) || 0) : 0;
          yTotal += u * uc;
        }
        qCi += 2;
      });
    });
    if (gtCells[ci]) gtCells[ci].textContent = ab_fc(yTotal);
    ci += 1;
  });
}

function ab_recalcNonProgSectionTotal(dataRowTr, quarters, years) {
  var tbody = dataRowTr.closest('tbody');
  if (!tbody) return;
  // Merged-table mode: scope to the section the changed row belongs to.
  var section = dataRowTr.dataset.section;
  if (!section) return;
  var totalRow = tbody.querySelector('.ab-section-total-row[data-section="' + section.replace(/"/g, '\\"') + '"]');
  if (!totalRow) return;
  var dataRows = tbody.querySelectorAll('.ab-data-row[data-section="' + section.replace(/"/g, '\\"') + '"]');
  var gtCells = totalRow.querySelectorAll('td');

  // Total Units = sum of each row's total_units input. Total Cost = sum of (uc × total_units).
  var totalUnits = 0, totalCost = 0;
  dataRows.forEach(function(dr) {
    var drCells = dr.querySelectorAll('td');
    var ucInp = drCells[3] && drCells[3].querySelector('input');
    var uc = ucInp ? (parseFloat(ucInp.value) || 0) : 0;
    var tuInp = drCells[4] && drCells[4].querySelector('input');
    var rowTU = tuInp ? (parseFloat(tuInp.value) || 0) : 0;
    totalUnits += rowTU;
    totalCost  += uc * rowTU;
  });

  gtCells[4].textContent = totalUnits;
  gtCells[5].textContent = ab_fc(totalCost);

  // Quarterly totals
  var ci = 6;
  quarters.forEach(function() {
    var qUnits = 0, qCost = 0;
    dataRows.forEach(function(dr) {
      var drCells = dr.querySelectorAll('td');
      var ucInp = drCells[3] && drCells[3].querySelector('input');
      var uc = ucInp ? (parseFloat(ucInp.value) || 0) : 0;
      var uInp = drCells[ci] && drCells[ci].querySelector('input');
      var u = uInp ? (parseInt(uInp.value) || 0) : 0;
      qUnits += u;
      qCost += u * uc;
    });
    gtCells[ci].textContent = qUnits;
    ci++;
    gtCells[ci].textContent = ab_fc(qCost);
    ci++;
  });

  // Year total columns
  years.forEach(function(y) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yTotal = 0;
    dataRows.forEach(function(dr) {
      var drCells = dr.querySelectorAll('td');
      var ucInp = drCells[3] && drCells[3].querySelector('input');
      var uc = ucInp ? (parseFloat(ucInp.value) || 0) : 0;
      var qCi = 6;
      quarters.forEach(function(q) {
        if (yqList.indexOf(q) >= 0) {
          var uInp = drCells[qCi] && drCells[qCi].querySelector('input');
          var u = uInp ? (parseInt(uInp.value) || 0) : 0;
          yTotal += u * uc;
        }
        qCi += 2;
      });
    });
    gtCells[ci].textContent = ab_fc(yTotal);
    ci += 1;
  });
}

// ============================================================================
// ADD ROW + SAVE (Non-Programmatic)
// ============================================================================

function ab_addNonProgRow(frm, secTitle, quarters, years, sbhRevMap, nonProgData, progData, unitsList) {
  // Merged-table mode: there is one shared tbody. Locate the section's TOTAL row by data-section
  // and insert the new data row before it. Section-scoped Sr renumber afterwards.
  var tbody = document.querySelector('.ab-nonprog-table tbody');
  if (!tbody) return;
  var qSec = secTitle.replace(/"/g, '\\"');
  var totalRow = tbody.querySelector('.ab-section-total-row[data-section="' + qSec + '"]');
  if (!totalRow) return;
  var dataRows = tbody.querySelectorAll('.ab-data-row[data-section="' + qSec + '"]');
  var newIdx = dataRows.length;

  var newRow = {
    description: '',
    assumption: '',
    remarks: '',
    unit_cost: 0,
    total_units: 0,
    pbpName: null,
    quarters: {}
  };
  for (var i = 0; i < quarters.length; i++) {
    newRow.quarters[i] = { units: 0 };
  }

  if (nonProgData.sections[secTitle]) {
    nonProgData.sections[secTitle].push(newRow);
  }

  var rowHtml = ab_buildNonProgRow(newRow, quarters, years, newIdx, secTitle, unitsList);
  var tempDiv = document.createElement('tbody');
  tempDiv.innerHTML = rowHtml;
  var newTr = tempDiv.querySelector('tr');

  tbody.insertBefore(newTr, totalRow);

  // Sr resets per section — only renumber rows in this section.
  tbody.querySelectorAll('.ab-data-row[data-section="' + qSec + '"]').forEach(function(tr, i) {
    var srCell = tr.querySelector('.ab-sr');
    if (srCell) srCell.textContent = i + 1;
  });

  newTr.querySelectorAll('.ab-inp').forEach(function(inp) {
    inp.addEventListener('change', function() {
      ab_recalcRow(this, quarters, years, progData, nonProgData);
    });
    inp.addEventListener('input', function() {
      ab_recalcRow(this, quarters, years, progData, nonProgData);
    });
  });

  var descInp = newTr.querySelector('.ab-desc-inp');
  if (descInp) descInp.focus();
}

function ab_saveNonProgRow(frm, tr, quarters, years, sbhRevMap) {
  var secTitle = tr.dataset.section;
  var pbpName = tr.dataset.pbp || '';
  var cells = tr.querySelectorAll('td');

  var descInput = tr.querySelector('.ab-desc-inp');
  var description = descInput ? descInput.value.trim() : '';
  if (!description) return Promise.resolve();

  var uomSelect = tr.querySelector('.ab-uom-sel');
  var uomValue = uomSelect ? uomSelect.value : '';

  var ucInput = tr.querySelector('.ab-uc-inp');
  var unitCost = ucInput ? (parseFloat(ucInput.value) || 0) : 0;

  // Total Units is now a user input — persist via total_planned_budget so we can derive it back on load.
  var tuInput = tr.querySelector('.ab-tu-inp');
  var totalUnits = tuInput ? (parseFloat(tuInput.value) || 0) : 0;
  var totalPlannedBudget = unitCost * totalUnits;

  var planningRows = [];
  var ci = 6; // After frozen cols + Total Units + Total Cost
  quarters.forEach(function(q) {
    var uInput = cells[ci] && cells[ci].querySelector('input');
    var units = uInput ? (parseFloat(uInput.value) || 0) : 0;
    ci += 2;

    planningRows.push({
      doctype: 'PBP Child',
      year: q.year,
      quarter: q.quarter,
      timespan: q.quarter,
      unit: units,
      unit_cost: unitCost,
      planned_amount: units * unitCost,
      start_date: q.start_date,
      end_date: q.end_date
    });
  });

  var sbhInfo = sbhRevMap[secTitle] || {};
  if (!sbhInfo.sbhId) {
    console.warn('[AB] No sbh mapping for section:', secTitle);
    return Promise.resolve();
  }

  if (pbpName) {
    return frappe.call({ method: 'frappe.client.get', args: { doctype: 'Project Budget Planning', name: pbpName } })
      .then(function(r) {
        if (!r.message) return;
        var doc = r.message;
        doc.description = description;
        doc.assumption = uomValue;
        doc.total_planned_budget = totalPlannedBudget;
        doc.planning_table = planningRows;
        return frappe.call({ method: 'frappe.client.save', args: { doc: doc } });
      })
      .then(function() { console.log('[AB] Updated PBP:', pbpName); })
      .catch(function(e) { console.error('[AB] Save error:', pbpName, e); });
  } else {
    var newDoc = {
      doctype: 'Project Budget Planning',
      project_proposal: frm.doc.name,
      donor: 'D-0001',
      description: description,
      assumption: uomValue,
      budget_head: sbhInfo.bhId,
      sub_budget_head: sbhInfo.sbhId,
      total_planned_budget: totalPlannedBudget,
      planning_table: planningRows
    };
    return frappe.call({ method: 'frappe.client.save', args: { doc: newDoc } })
      .then(function(r) {
        if (r.message) {
          tr.dataset.pbp = r.message.name;
          console.log('[AB] Created PBP:', r.message.name);
        }
      })
      .catch(function(e) { console.error('[AB] Create error:', e); });
  }
}

function ab_saveNonProgSequential(frm, rows, idx, quarters, years, sbhRevMap) {
  if (idx >= rows.length) {
    ab_showSaved();
    frappe.show_alert({ message: 'Saved ' + rows.length + ' rows successfully', indicator: 'green' });
    return;
  }
  var p = ab_saveNonProgRow(frm, rows[idx], quarters, years, sbhRevMap);
  (p || Promise.resolve()).then(function() {
    ab_saveNonProgSequential(frm, rows, idx + 1, quarters, years, sbhRevMap);
  }).catch(function(e) {
    console.error('[AB] Sequential save error at row ' + idx + ':', e);
    frappe.show_alert({ message: 'Error saving row ' + (idx + 1) + '. Stopped.', indicator: 'red' });
  });
}

async function ab_saveProgData(frm, quarters, years, progData, sbhRevMap) {
  var rows = progData.rows || [];
  if (!rows.length) { frappe.show_alert({ message: 'No programmatic rows to save', indicator: 'orange' }); return; }

  var progSbh = sbhRevMap['Programmatic Costs'] || {};
  if (!progSbh.sbhId) {
    frappe.show_alert({ message: 'No sub-budget head mapping for "Programmatic Costs"', indicator: 'red' });
    return;
  }

  var table = document.querySelector('.ab-prog-table');
  if (!table) return;
  var dataRows = table.querySelectorAll('.ab-data-row');

  frappe.show_alert({ message: 'Saving programmatic data...', indicator: 'blue' });

  var saveCount = 0;
  var errCount = 0;

  for (var ri = 0; ri < rows.length; ri++) {
    var row = rows[ri];
    var drEl = dataRows[ri];
    if (!drEl) continue;
    var drCells = drEl.querySelectorAll('td');

    // Read inputs
    var ucInp = drCells[4] && drCells[4].querySelector('input');
    var uc = ucInp ? (parseFloat(ucInp.value) || 0) : 0;
    var tuInp = drCells[5] && drCells[5].querySelector('input');
    var tu = tuInp ? (parseFloat(tuInp.value) || 0) : 0;
    var govtInp = drCells[8] && drCells[8].querySelector('input');
    var govt = govtInp ? (parseFloat(govtInp.value) || 0) : 0;
    var benfInp = drCells[9] && drCells[9].querySelector('input');
    var benf = benfInp ? (parseFloat(benfInp.value) || 0) : 0;
    var taskInp = drCells[2] && drCells[2].querySelector('input');
    var task = taskInp ? taskInp.value : '';
    var remarksInp = drCells[drCells.length - 1] && drCells[drCells.length - 1].querySelector('input');
    var remarks = remarksInp ? remarksInp.value : '';

    // Collect quarterly units
    var planningRows = [];
    var ci = 10;
    quarters.forEach(function(q) {
      var qUnitsInp = drCells[ci] && drCells[ci].querySelector('input');
      var qUnits = qUnitsInp ? (parseFloat(qUnitsInp.value) || 0) : 0;
      ci += 4;

      var qCost = uc * qUnits;
      planningRows.push({
        doctype: 'PBP Child',
        year: q.year,
        quarter: q.quarter,
        timespan: q.quarter,
        unit: qUnits,
        unit_cost: uc,
        planned_amount: qCost,
        start_date: q.start_date,
        end_date: q.end_date
      });
    });

    var totalBudget = planningRows.reduce(function(s, r) { return s + r.planned_amount; }, 0);

    // Persist convergence breakdown to custom fields (added Apr 2026).
    var govt = row.govt_contribution || 0;
    var benf = row.benf_contribution || 0;
    var licContrib = totalBudget - govt - benf;

    // Save or create PBP record
    var pbpName = row.pbpName;
    if (pbpName) {
      try {
        var r = await frappe.call({ method: 'frappe.client.get', args: { doctype: 'Project Budget Planning', name: pbpName } });
        if (r.message) {
          var doc = r.message;
          doc.description = row.description;
          doc.assumption = task;
          doc.custom_task_details = task;
          doc.total_planned_budget = totalBudget;
          doc.custom_total_lic_contribution = licContrib;
          doc.custom_govt_contribution = govt;
          doc.custom_benf_contribution = benf;
          doc.planning_table = planningRows;
          await frappe.call({ method: 'frappe.client.save', args: { doc: doc } });
          saveCount++;
        }
      } catch (e) { console.error('[AB] Prog save error:', pbpName, e); errCount++; }
    } else if (totalBudget > 0) {
      try {
        var newDoc = {
          doctype: 'Project Budget Planning',
          project_proposal: frm.doc.name,
          donor: 'D-0001',
          description: row.description,
          budget_head: progSbh.bhId,
          sub_budget_head: progSbh.sbhId,
          fund_source: 'D-0001',
          assumption: task,
          custom_task_details: task,
          total_planned_budget: totalBudget,
          custom_total_lic_contribution: licContrib,
          custom_govt_contribution: govt,
          custom_benf_contribution: benf,
          planning_table: planningRows
        };
        var cr = await frappe.call({ method: 'frappe.client.save', args: { doc: newDoc } });
        if (cr.message) {
          row.pbpName = cr.message.name;
          saveCount++;
        }
      } catch (e) { console.error('[AB] Prog create error:', e); errCount++; }
    }
  }

  if (errCount) frappe.show_alert({ message: 'Saved ' + saveCount + ' records with ' + errCount + ' errors', indicator: 'orange' });
  else frappe.show_alert({ message: 'Saved ' + saveCount + ' programmatic records', indicator: 'green' });
  ab_showSaved();
}

function ab_showSaved() {
  var indicator = document.getElementById('ab-saved');
  if (indicator) {
    indicator.textContent = '✓ Saved';
    indicator.style.display = 'inline';
    setTimeout(function() { indicator.style.display = 'none'; }, 2000);
  }
}

// ============================================================================
// HIDE BUDGET SUMMARY TAB
// ============================================================================

function ab_hideBudgetSummaryTab(frm) {
  try {
    var tabLink = document.querySelector('[data-fieldname="custom_budget_summary_tab"]');
    if (tabLink) {
      var tabEl = tabLink.closest('.form-clickable-section') || tabLink.closest('.nav-item') || tabLink;
      if (tabEl) tabEl.style.display = 'none';
    }
    var allTabs = document.querySelectorAll('.form-tabs .nav-link, .form-tabs .tab-link');
    allTabs.forEach(function(tab) {
      if (tab.textContent.trim() === 'Budget Summary') {
        var parent = tab.closest('li') || tab.closest('.nav-item') || tab;
        parent.style.display = 'none';
      }
    });
  } catch (e) {
    console.warn('[AB] Could not hide Budget Summary tab:', e);
  }
}

// ============================================================================
// UTILITIES
// ============================================================================

function ab_fc(n) {
  if (n === 0) return '0';
  return n.toLocaleString('en-IN', { maximumFractionDigits: 0 });
}

function ab_he(s) {
  if (!s) return '';
  var d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

function ab_dateRange(s, e) {
  if (!s || !e) return '';
  var fmt = function(d) {
    var dt = new Date(d);
    return dt.toLocaleDateString('en-IN', { day: '2-digit', month: 'short' });
  };
  return '(' + fmt(s) + ' - ' + fmt(e) + ')';
}

// ============================================================================
// STYLES
// ============================================================================

function ab_getStyles() {
  return `
.ab-container { font-family: Arial, sans-serif; font-size: 13px; color: #333; }
.ab-tabs { display: flex; gap: 8px; margin-bottom: 12px; border-bottom: 2px solid #ddd; }
.ab-tab-btn { padding: 10px 20px; background: transparent; border: none; border-bottom: 3px solid transparent; cursor: pointer; font-weight: 600; color: #666; }
.ab-tab-active { color: #8B1A1A; border-bottom-color: #8B1A1A; }
.ab-tab-content { display: block; }
.ab-tab-content.ab-hidden { display: none; }
.ab-footer { display: flex; gap: 16px; align-items: center; padding: 12px 0; margin-top: 16px; }
.ab-legend { display: inline-flex; gap: 6px; align-items: center; font-size: 12px; }
.ab-scroll-wrapper { overflow: auto; max-height: 70vh; border: 1px solid #ddd; margin: 12px 0; position: relative; -webkit-overflow-scrolling: touch; }
.ab-table { width: max-content; border-collapse: separate; border-spacing: 0; background: white; table-layout: fixed; }
.ab-table td, .ab-table th { border: 1px solid #e0e0e0; padding: 6px 8px; text-align: center; font-size: 12px; white-space: nowrap; }
.ab-table thead th { position: sticky; z-index: 9; background: #f5f5f5; }
.ab-table thead tr:nth-child(1) th { top: 0; z-index: 11; }
.ab-table thead tr:nth-child(2) th { top: 33px; z-index: 11; }
.ab-frozen { position: sticky; z-index: 10; background: #fafafa; }
.ab-table thead .ab-frozen { z-index: 20 !important; }
.ab-table thead tr:nth-child(1) .ab-frozen-header { z-index: 20 !important; top: 0; }
.ab-table thead tr:nth-child(2) .ab-frozen { top: 33px; }
.ab-frozen-last { border-right: 2px solid #999; box-shadow: 2px 0 4px rgba(0,0,0,0.08); }
.ab-frozen-header { position: sticky; z-index: 15; background: #f5f5f5; font-weight: 700; }
.ab-hdr-r1, .ab-hdr-r2, .ab-header-row-1, .ab-header-row-2 { background: #f5f5f5; font-weight: 700; }
.ab-convergence-header { background: #8B1A1A !important; color: white !important; font-weight: 700; }
.ab-quarter-header { background: #8B1A1A !important; color: white !important; font-weight: 700; }
.ab-remarks-header, .ab-combined-header { background: #f5f5f5; font-weight: 700; }
.ab-col-hdr, .ab-sr-hdr { background: #f5f5f5; font-weight: 700; }
.ab-subcol-hdr, .ab-subcol { background: #F5E6E6; font-weight: 600; font-size: 11px; color: #333; }
.ab-remarks-hdr { background: #f5f5f5; font-weight: 600; font-size: 11px; }
.ab-q-odd, .ab-q-even { }
.ab-editable { background: white; }
.ab-editable input { background: white; }
.ab-calc { background: #f5f5f5; color: #555; }
.ab-inp { width: 95%; padding: 5px 8px; border: 1px solid #d1d8dd; border-radius: 3px; font-size: 12px; background: white; }
.ab-inp:focus { border-color: #8B1A1A; outline: none; box-shadow: 0 0 0 1px rgba(139,26,26,0.2); }
.ab-task-inp, .ab-remarks-inp { text-align: left; }
.ab-desc-cell { padding: 6px 8px; text-align: left; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ab-sr { width: 40px; text-align: center; }
.ab-gt-cell { background: #F5E6E6; font-weight: 600; }
.ab-gt-final { font-weight: 700; }
.ab-grand-total-row { background: #F5E6E6; font-weight: 700; border-top: 2px solid #8B1A1A; }
.ab-grand-total-header { background: #F5E6E6; font-weight: 700; }
.ab-year-header { background: #f5f5f5; font-weight: 700; }
.ab-year-total-header { background: #f5f5f5; }
.ab-yt-cell, .ab-yt-subcol { background: #f9f9f9; }
.ab-year-header-cell { background: #8B1A1A !important; color: white !important; font-weight: 700; padding: 0 !important; vertical-align: top; }
.ab-year-top { background: rgba(255,255,255,0.15); padding: 2px 4px; text-align: center; font-size: 10px; font-weight: 600; border-bottom: 1px solid rgba(255,255,255,0.2); }
.ab-quarter-row { display: flex; }
.ab-q-label { flex: 1; text-align: center; padding: 4px 2px; font-size: 11px; font-weight: 700; }
.ab-gt-header-cell { background: #F5E6E6 !important; font-weight: 700; padding: 0 !important; vertical-align: top; }
.ab-gt-header-cell .ab-year-top { background: rgba(139,26,26,0.1); color: #333; }
.ab-gt-header-cell .ab-q-label { color: #333; font-size: 10px; }
.ab-gt-subcol { background: #F5E6E6; font-weight: 600; }
.ab-section-total-row { background: #BCBDC0; font-weight: 700; }
.ab-section-total-row .ab-frozen { background: #BCBDC0; }
.ab-add-row-btn { font-size: 12px; padding: 4px 12px; cursor: pointer; }
.ab-saved-indicator { color: #8B1A1A; font-weight: 700; display: none; }
.ab-nonprog-wrapper { }

/* Single shared vertical-scroll wrapper for the whole Non-Programmatic tab.
   ONE table inside, all sections merged. Sticky headers stay trapped here. */
.ab-nonprog-scroll { overflow: auto; max-height: 70vh; border: 1px solid #ddd; -webkit-overflow-scrolling: touch; }

/* Inline section banner row (full-width dark grey, matches source Excel #333333). */
.ab-np-section-banner td { background: #333333 !important; padding: 0 !important; border: none !important; }
.ab-np-section-banner-cell {
  padding: 10px 16px !important;
  color: #fff !important;
  font-weight: 700 !important;
  font-size: 13px;
  letter-spacing: 0.3px;
  text-align: left !important;
  white-space: nowrap;
}
.ab-np-section-letter { display: inline-block; min-width: 22px; font-weight: 700; }

/* Inline "+ Add Row to <Section>" affordance row */
.ab-np-addrow-row td { background: #fafafa; padding: 6px 16px !important; border: none !important; border-bottom: 1px solid #eee !important; text-align: left !important; }
.ab-np-addrow-cell .ab-add-row-btn {
  border: 1px dashed #8B1A1A;
  color: #8B1A1A;
  background: #fff;
  border-radius: 4px;
  font-weight: 600;
}
.ab-np-addrow-cell .ab-add-row-btn:hover { background: #8B1A1A; color: #fff; }

/* Final GRAND TOTAL row at the very bottom (yellow, matches source Excel #FFCB05) */
.ab-np-grand-total td { background: #FFCB05 !important; font-weight: 700; color: #222; }
.ab-np-grand-total .ab-frozen { background: #FFCB05 !important; }
.ab-table tbody tr:nth-child(even) td { background: #fafafa; }
.ab-table tbody tr:nth-child(even) .ab-frozen { background: #f5f5f5; }
.ab-table tbody tr:nth-child(even) .ab-calc { background: #f0f0f0; }
.ab-table tbody tr:nth-child(even) .ab-gt-cell { background: #F0DCDC; }
  `;
}

// ============================================================================
// EXCEL EXPORT  (rewritten — see HANDOVER_v9.md)
// Library: xlsx-js-style@1.2.0  (per mgrant-frappe-patterns/client-side-xlsx-export)
// + JSZip post-process to inject the LIC HFL CSR logo (xlsx-js-style cannot embed images natively).
// Plain values only — no formulas, no sheet protection, no calcChain.
// Indian number format (lakh-style grouping). Multi-year extension with Option C
// (per-year Q1-Q4 blocks first, then a "Grand Total" band with one column per year, then Remarks).
// ============================================================================

// LIC HFL CSR logo (extracted from the Proposed Budget Format source workbook — xl/media/image1.jpeg).
// Embedded as base64 so the export is self-contained.
var AB_LOGO_JPEG_B64 = "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";

// Source-matching colour palette (extracted from the proposed budget xlsx).
var AB_C_DARK      = '333333';   // section banner / Project Details / Budget Breakup heads
var AB_C_GREY_HDR  = 'BCBDC0';   // column-name strip (Sr. No., Activity, ...)
var AB_C_LIGHT     = 'E6E6E6';   // label cells (I, II, III ... Programmatic Costs ...)
var AB_C_ALT       = 'F4F5F6';   // alternate body rows
var AB_C_YELLOW    = 'FFCB05';   // grand total / total banners (LIC HFL yellow)
var AB_C_BLUE      = '00529C';   // Q1/Q3 quarter pills, title accents (LIC HFL blue)
var AB_C_WHITE     = 'FFFFFF';
var AB_INR_FMT     = '_(₹* #,##,##,##0.00_);_(₹* (#,##,##,##0.00);_(₹* "-"??_);_(@_)';
var AB_NUM_FMT     = '#,##,##0';
var AB_PCT_FMT     = '0.00%';

function ab_loadXlsxLib() {
  return new Promise(function(resolve, reject) {
    if (window.XLSX && window.XLSX.utils && window.XLSX.utils.aoa_to_sheet) { resolve(); return; }
    var s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';
    s.onload = function() { resolve(); };
    s.onerror = function() { reject(new Error('Failed to load xlsx-js-style')); };
    document.head.appendChild(s);
  });
}

function ab_loadJSZip() {
  return new Promise(function(resolve, reject) {
    if (window.JSZip) { resolve(window.JSZip); return; }
    var s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
    s.onload = function() { resolve(window.JSZip); };
    s.onerror = function() { reject(new Error('Failed to load JSZip')); };
    document.head.appendChild(s);
  });
}

// Styled-cell helper (Calibri matches source). Dropping borders default to thin black.
function ab_sc(v, opts) {
  opts = opts || {};
  var cell = { v: v, t: typeof v === 'number' ? 'n' : 's' };
  var s = {};
  s.font = opts.font || { name: 'Calibri', sz: 11 };
  if (!s.font.name) s.font.name = 'Calibri';
  if (opts.fill) s.fill = { fgColor: { rgb: opts.fill } };
  s.alignment = opts.alignment || { vertical: 'center', wrapText: true };
  if (opts.border !== false) {
    var thin = { style: 'thin', color: { rgb: '000000' } };
    s.border = opts.border || { top: thin, bottom: thin, left: thin, right: thin };
  }
  if (opts.numFmt) { s.numFmt = opts.numFmt; cell.z = opts.numFmt; }
  cell.s = s;
  return cell;
}

// ============================================================================
// Project meta resolution — NGO name lookup, ordinal date formatting, states + districts
// ============================================================================

// Format a Frappe date as ordinal English: '1st April 2026'.
// Handles 'YYYY-MM-DD' (Frappe storage), 'DD-MM-YYYY' (Frappe display), 'DD/MM/YYYY', and Date objects.
function ab_parseDate(s) {
  if (!s) return null;
  if (s instanceof Date) return isNaN(s.getTime()) ? null : s;
  s = String(s).trim();
  if (!s) return null;
  // YYYY-MM-DD or YYYY/MM/DD
  var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) {
    var d = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
    return isNaN(d.getTime()) ? null : d;
  }
  // DD-MM-YYYY or DD/MM/YYYY
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  if (m) {
    var d2 = new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
    return isNaN(d2.getTime()) ? null : d2;
  }
  // Last-resort: native parser (handles ISO timestamps, RFC strings)
  var d3 = new Date(s);
  return isNaN(d3.getTime()) ? null : d3;
}

function ab_formatOrdinalDate(s) {
  var d = ab_parseDate(s);
  if (!d) return s ? String(s) : '';
  var day = d.getDate();
  var months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  var month = months[d.getMonth()];
  var year = d.getFullYear();
  var suffix = 'th';
  var lastTwo = day % 100;
  if (lastTwo < 11 || lastTwo > 13) {
    var lastOne = day % 10;
    if (lastOne === 1) suffix = 'st';
    else if (lastOne === 2) suffix = 'nd';
    else if (lastOne === 3) suffix = 'rd';
  }
  return day + suffix + ' ' + month + ' ' + year;
}

// Pull the comma-separated list of intervention states from frm.doc.
// Tries common mGrant fieldnames, then introspects child tables containing a 'state'-ish field.
function ab_extractStatesString(frm) {
  var named = ['proposal_state', 'project_intervention_state', 'intervention_state', 'state_table', 'project_state', 'location_state', 'states', 'state_details', 'project_states', 'geography_state', 'geography', 'project_location', 'project_geography', 'state_intervention'];
  var stateKeys = ['state', 'state_name', 'state1', 'project_state', 'state_link', 'name_of_state'];
  function readFrom(arr) {
    for (var k = 0; k < stateKeys.length; k++) {
      var key = stateKeys[k];
      if (arr.length && arr[0] && arr[0][key] !== undefined && arr[0][key] !== null) {
        var vals = arr.map(function(r) { return (r[key] || '').toString().trim(); }).filter(Boolean);
        var seen = {}, uniq = [];
        vals.forEach(function(v) { if (!seen[v]) { seen[v] = 1; uniq.push(v); } });
        if (uniq.length) return uniq;
      }
    }
    return null;
  }
  for (var i = 0; i < named.length; i++) {
    var arr = frm.doc[named[i]];
    if (Array.isArray(arr) && arr.length) {
      var vals = readFrom(arr);
      if (vals && vals.length) { console.log('[AB] States from frm.doc.' + named[i] + ':', vals); return vals.join(', '); }
    }
  }
  // Fallback: scan EVERY child table on the doc for any state-like key
  for (var k in frm.doc) {
    var v = frm.doc[k];
    if (Array.isArray(v) && v.length && v[0] && typeof v[0] === 'object') {
      var vals2 = readFrom(v);
      if (vals2 && vals2.length) { console.log('[AB] States introspected from frm.doc.' + k + ':', vals2); return vals2.join(', '); }
    }
  }
  if (frm.doc.state) { console.log('[AB] States from flat frm.doc.state'); return frm.doc.state; }
  console.warn('[AB] No states field/child-table found on frm.doc. Child tables present:', ab_listChildTables(frm));
  return '';
}

// Count distinct districts on the GAF — used as "No. of Intervention Site(s)" in Budget Summary.
function ab_extractDistrictCount(frm) {
  var named = ['proposal_district', 'project_intervention_district', 'intervention_district', 'district_table', 'project_district', 'districts', 'district_details', 'block_table', 'project_block', 'geography_district', 'geography', 'project_location', 'project_geography', 'district_intervention', 'project_districts'];
  var districtKeys = ['district', 'district_name', 'block', 'block_name', 'district_link', 'name_of_district'];
  function readFrom(arr) {
    for (var k = 0; k < districtKeys.length; k++) {
      var key = districtKeys[k];
      if (arr.length && arr[0] && arr[0][key] !== undefined && arr[0][key] !== null) {
        var vals = arr.map(function(r) { return (r[key] || '').toString().trim(); }).filter(Boolean);
        var seen = {}, uniq = [];
        vals.forEach(function(v) { if (!seen[v]) { seen[v] = 1; uniq.push(v); } });
        if (uniq.length) return uniq;
      }
    }
    return null;
  }
  for (var i = 0; i < named.length; i++) {
    var arr = frm.doc[named[i]];
    if (Array.isArray(arr) && arr.length) {
      var vals = readFrom(arr);
      if (vals && vals.length) { console.log('[AB] Districts from frm.doc.' + named[i] + ', count:', vals.length, 'values:', vals); return vals.length; }
    }
  }
  for (var k in frm.doc) {
    var v = frm.doc[k];
    if (Array.isArray(v) && v.length && v[0] && typeof v[0] === 'object') {
      var vals2 = readFrom(v);
      if (vals2 && vals2.length) { console.log('[AB] Districts introspected from frm.doc.' + k + ', count:', vals2.length, 'values:', vals2); return vals2.length; }
    }
  }
  console.warn('[AB] No district field/child-table found on frm.doc. Child tables present:', ab_listChildTables(frm));
  return 0;
}

// Helper for diagnostics: list every child-table fieldname and the keys of its first row.
function ab_listChildTables(frm) {
  var out = {};
  for (var k in frm.doc) {
    var v = frm.doc[k];
    if (Array.isArray(v) && v.length && v[0] && typeof v[0] === 'object') {
      out[k] = { rows: v.length, sampleKeys: Object.keys(v[0]).filter(function(x) { return !x.startsWith('_') && !['name','owner','creation','modified','modified_by','docstatus','idx','parent','parentfield','parenttype'].includes(x); }) };
    }
  }
  return out;
}

// (Removed: ab_discoverGeographyViaAPI — the multi-doctype probe loop hung the download.
// Geography is now extracted ONLY from frm.doc via ab_extractStatesString / ab_extractDistrictCount.
// If the values come up blank, the diagnostic dump in ab_resolveProjectMeta tells us which
// child tables are present so we can hard-code the correct fieldname in the next iteration.)

// Geography loader — fetches the linked `Geography Details` doc (name pattern GD-Project proposal-{GAF}).
// Returns { states_str, intervention_sites } or empty values if no doc / lookup fails.
async function ab_loadGeographyFromDetails(frm) {
  if (!frappe || !frappe.db || !frappe.db.get_doc) return { states_str: '', intervention_sites: 0 };
  var gdName = 'GD-Project proposal-' + frm.doc.name;
  try {
    var gd = await frappe.db.get_doc('Geography Details', gdName);
    var rows = (gd && gd.geography_details) || [];
    var stateSeen = {}, states = [];
    var districtSeen = {}, districts = [];
    rows.forEach(function(r) {
      var s = (r.state_name || '').toString().trim();
      var d = (r.district_name || '').toString().trim();
      if (s && !stateSeen[s]) { stateSeen[s] = 1; states.push(s); }
      if (d && !districtSeen[d]) { districtSeen[d] = 1; districts.push(d); }
    });
    console.log('[AB] Geography Details ' + gdName + ': ' + states.length + ' states, ' + districts.length + ' districts');
    return { states_str: states.join(', '), intervention_sites: districts.length };
  } catch (e) {
    console.warn('[AB] Geography Details lookup failed for ' + gdName + ':', (e && e.message) || e);
    return { states_str: '', intervention_sites: 0 };
  }
}

// Resolve everything Budget Summary needs that isn't already on frm.doc.
// Loads NGO name + Geography Details once per export.
async function ab_resolveProjectMeta(frm) {
  // Diagnostic dump
  console.log('[AB] frm.doc keys:', Object.keys(frm.doc).sort());
  console.log('[AB] Dates:', { project_start_date: frm.doc.project_start_date, project_end_date: frm.doc.project_end_date });
  console.log('[AB] Beneficiaries:', { custom_direct_beneficiary: frm.doc.custom_direct_beneficiary, custom_indirect_beneficiary: frm.doc.custom_indirect_beneficiary });
  console.log('[AB] Partner field:', frm.doc.ngo);

  // NGO name lookup — single targeted call against NGO doctype.
  var ngoId = frm.doc.ngo || frm.doc.partner || frm.doc.implementing_partner || '';
  var ngoName = frm.doc.ngo_name || frm.doc.partner_name || ngoId;
  if (ngoId && frappe && frappe.db && frappe.db.get_value) {
    try {
      var resp = await frappe.db.get_value('NGO', ngoId, 'ngo_name');
      var n = resp && resp.message && resp.message.ngo_name;
      if (n) ngoName = n;
    } catch (e) { console.warn('[AB] NGO name lookup failed:', e); }
  }
  console.log('[AB] Partner name resolved:', ngoName);

  // Project dates — Frappe stores them on `project_start_date` / `project_end_date`,
  // not the legacy `start_date` / `end_date`.
  var startStr = ab_formatOrdinalDate(frm.doc.project_start_date || frm.doc.start_date);
  var endStr   = ab_formatOrdinalDate(frm.doc.project_end_date   || frm.doc.end_date);
  var durationStr = (startStr || endStr) ? (startStr + ' to ' + endStr) : '';
  console.log('[AB] Dates parsed:', { start: startStr, end: endStr, duration: durationStr });

  // Geography — fetched from the linked Geography Details doctype.
  var geo = await ab_loadGeographyFromDetails(frm);
  console.log('[AB] Geography:', geo);

  // Beneficiaries — Frappe stores them on `custom_direct_beneficiary`.
  var beneficiaries = parseFloat(frm.doc.custom_direct_beneficiary) || parseFloat(frm.doc.total_beneficiaries) || 0;

  return {
    ngo_name: ngoName,
    start_date_str: startStr,
    end_date_str: endStr,
    duration_str: durationStr,
    states_str: geo.states_str,
    intervention_sites: geo.intervention_sites,
    total_beneficiaries: beneficiaries
  };
}

// Group `quarters` (the flat list across years) by year_sequence so we can
// build Year 1, Year 2, ... blocks. Returns [{year, qs:[q1..q4]}, ...]
function ab_groupQuartersByYear(quarters, years) {
  return (years || []).map(function(y, yi) {
    var ys = (y && y.year_sequence != null) ? y.year_sequence : (yi + 1);
    var label = (y && y.year_label) ? y.year_label : ('Year ' + (yi + 1));
    return {
      year: y,
      yIdx: yi,
      label: label,
      qs: quarters.filter(function(q) { return q.year_sequence === ys; })
    };
  });
}

// Build the orchestrator.
async function ab_downloadBudget(frm, quarters, years, progData, nonProgData, unitsList) {
  try {
    frappe.show_alert({ message: 'Preparing Excel download...', indicator: 'blue' });
    await ab_loadXlsxLib();
    var JSZip = await ab_loadJSZip();

    var progLive = ab_collectProgDataFromDOM(quarters, progData);
    var npLive   = ab_collectNonProgDataFromDOM(quarters, nonProgData);
    var yearGroups = ab_groupQuartersByYear(quarters, years);
    var meta = await ab_resolveProjectMeta(frm);

    var wb = XLSX.utils.book_new();
    var summary = ab_buildBudgetSummarySheet(frm, quarters, yearGroups, progLive, npLive, meta);
    XLSX.utils.book_append_sheet(wb, summary.ws, 'Budget Summary');
    var prog = ab_buildProgSheet(frm, quarters, yearGroups, progLive);
    XLSX.utils.book_append_sheet(wb, prog.ws, 'Programmatic Costs');
    var np = ab_buildNonProgSheet(frm, quarters, yearGroups, npLive);
    XLSX.utils.book_append_sheet(wb, np.ws, 'Non-Programmatic Costs');

    // Render to ArrayBuffer (no protection, no calcChain — clean file)
    var ab = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

    // Inject the logo via JSZip post-process.
    var blob = await ab_injectLogo(ab, JSZip, [
      { sheetIdx: 1, fromCol: 0, toCol: 5 },  // Budget Summary: A1:F2
      { sheetIdx: 2, fromCol: 0, toCol: 2 },  // Programmatic:    A1:C2
      { sheetIdx: 3, fromCol: 0, toCol: 2 }   // Non-Programmatic: A1:C2
    ]);

    var filename = (frm.doc.name || 'Budget') + '_Budget_' + new Date().toISOString().split('T')[0] + '.xlsx';
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    setTimeout(function() { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);

    frappe.show_alert({ message: 'Downloaded: ' + filename, indicator: 'green' });
  } catch (err) {
    console.error('[AB] Download error:', err);
    frappe.show_alert({ message: 'Excel error: ' + (err.message || err), indicator: 'red' });
  }
}

// ── JSZip post-process: insert image1.jpeg + drawing.xml + relationships ──
async function ab_injectLogo(arrayBuffer, JSZip, anchors) {
  var zip = await JSZip.loadAsync(arrayBuffer);

  // 1. Add the JPEG bytes
  var bin = atob(AB_LOGO_JPEG_B64);
  var bytes = new Uint8Array(bin.length);
  for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  zip.file('xl/media/image1.jpeg', bytes, { binary: true });

  for (var s = 0; s < anchors.length; s++) {
    var anc = anchors[s];
    var sIdx = anc.sheetIdx;
    var drawingXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<xdr:twoCellAnchor editAs="oneCell">' +
          '<xdr:from><xdr:col>' + anc.fromCol + '</xdr:col><xdr:colOff>57150</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>57150</xdr:rowOff></xdr:from>' +
          '<xdr:to><xdr:col>' + anc.toCol + '</xdr:col><xdr:colOff>647700</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>714375</xdr:rowOff></xdr:to>' +
          '<xdr:pic>' +
            '<xdr:nvPicPr><xdr:cNvPr id="' + (1024 + sIdx) + '" name="LIC HFL Logo"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
            '<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
            '<xdr:spPr bwMode="auto"><a:xfrm><a:off x="57150" y="57150"/><a:ext cx="3114675" cy="847725"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>' +
          '</xdr:pic>' +
          '<xdr:clientData/>' +
        '</xdr:twoCellAnchor>' +
      '</xdr:wsDr>';
    zip.file('xl/drawings/drawing' + sIdx + '.xml', drawingXml);

    var drawingRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/>' +
      '</Relationships>';
    zip.file('xl/drawings/_rels/drawing' + sIdx + '.xml.rels', drawingRels);

    // Patch sheet xml — append <drawing r:id="rIdN"/> before </worksheet>
    var sheetPath = 'xl/worksheets/sheet' + sIdx + '.xml';
    var sheetFile = zip.file(sheetPath);
    if (!sheetFile) continue;
    var sheetXml = await sheetFile.async('string');
    // Find an unused rId — append to existing rels list
    var sheetRelsPath = 'xl/worksheets/_rels/sheet' + sIdx + '.xml.rels';
    var existingRelsFile = zip.file(sheetRelsPath);
    var newRelsXml;
    var nextRid = 'rId1';
    if (existingRelsFile) {
      var existingRelsXml = await existingRelsFile.async('string');
      // Find max rId
      var maxId = 0;
      var rgx = /Id="rId(\d+)"/g;
      var m;
      while ((m = rgx.exec(existingRelsXml)) !== null) {
        if (parseInt(m[1]) > maxId) maxId = parseInt(m[1]);
      }
      nextRid = 'rId' + (maxId + 1);
      newRelsXml = existingRelsXml.replace(
        '</Relationships>',
        '<Relationship Id="' + nextRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + sIdx + '.xml"/></Relationships>'
      );
    } else {
      newRelsXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + sIdx + '.xml"/>' +
        '</Relationships>';
    }
    zip.file(sheetRelsPath, newRelsXml);

    if (sheetXml.indexOf('<drawing ') === -1) {
      sheetXml = sheetXml.replace('</worksheet>', '<drawing r:id="' + nextRid + '"/></worksheet>');
      zip.file(sheetPath, sheetXml);
    }
  }

  // 2. Update [Content_Types].xml — add jpeg Default + drawing Overrides
  var ctFile = zip.file('[Content_Types].xml');
  var ct = await ctFile.async('string');
  if (ct.indexOf('Extension="jpeg"') === -1) {
    ct = ct.replace('<Default ', '<Default Extension="jpeg" ContentType="image/jpeg"/><Default ');
  }
  for (var d = 0; d < anchors.length; d++) {
    var di = anchors[d].sheetIdx;
    if (ct.indexOf('drawing' + di + '.xml') === -1) {
      ct = ct.replace(
        '</Types>',
        '<Override PartName="/xl/drawings/drawing' + di + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'
      );
    }
  }
  zip.file('[Content_Types].xml', ct);

  return await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

// ============================================================================
// Live data collection from DOM (web UI is the source of truth)
// ============================================================================

function ab_collectProgDataFromDOM(quarters, progData) {
  var rows = [];
  var table = document.querySelector('.ab-prog-table');
  if (!table) return { rows: rows };
  var dataRows = table.querySelectorAll('.ab-data-row');

  dataRows.forEach(function(tr, idx) {
    var origRow = (progData.rows || [])[idx] || {};
    var cells = tr.querySelectorAll('td');
    var desc = origRow.description || '';
    var task = '';
    var taskInp = cells[2] && cells[2].querySelector('input');
    if (taskInp) task = taskInp.value;

    var ucInp = cells[4] && cells[4].querySelector('input');
    var uc = ucInp ? parseFloat(ucInp.value) || 0 : 0;
    var tuInp = cells[5] && cells[5].querySelector('input');
    var tu = tuInp ? parseFloat(tuInp.value) || 0 : 0;
    var govtInp = cells[8] && cells[8].querySelector('input');
    var govt = govtInp ? parseFloat(govtInp.value) || 0 : 0;
    var benfInp = cells[9] && cells[9].querySelector('input');
    var benf = benfInp ? parseFloat(benfInp.value) || 0 : 0;
    var uom = origRow.uomName || 'Numbers';

    var qtrData = [];
    var ci = 10;
    quarters.forEach(function() {
      var qUnitsInp = cells[ci] && cells[ci].querySelector('input');
      var qUnits = qUnitsInp ? parseFloat(qUnitsInp.value) || 0 : 0;
      var qCostCell = cells[ci + 1];
      var qCost = qCostCell && qCostCell.textContent ? parseFloat(qCostCell.textContent.replace(/[^0-9.-]/g, '')) || 0 : 0;
      var qLhflCell = cells[ci + 2];
      var qLhfl = qLhflCell && qLhflCell.textContent ? parseFloat(qLhflCell.textContent.replace(/[^0-9.-]/g, '')) || 0 : 0;
      var qBenfCell = cells[ci + 3];
      var qBenf = qBenfCell && qBenfCell.textContent ? parseFloat(qBenfCell.textContent.replace(/[^0-9.-]/g, '')) || 0 : 0;
      qtrData.push({ units: qUnits, cost: qCost, licHfl: qLhfl, benf: qBenf });
      ci += 4;
    });
    rows.push({ description: desc, task: task, uom: uom, unitCost: uc, totalUnits: tu, govt: govt, benf: benf, quarters: qtrData, remarks: '' });
  });
  return { rows: rows };
}

function ab_collectNonProgDataFromDOM(quarters, nonProgData) {
  var sections = {};
  var sectionOrder = ['Human Resource Costs', 'Administration Costs', 'NGO Management Costs'];
  sectionOrder.forEach(function(sec) {
    var sRows = [];
    var allTrs = document.querySelectorAll('.ab-nonprog-table .ab-data-row[data-section="' + sec + '"]');
    allTrs.forEach(function(tr) {
      var descInp = tr.querySelector('.ab-desc-inp');
      var uomSel  = tr.querySelector('.ab-uom-sel');
      var ucInp   = tr.querySelector('.ab-uc-inp');
      var tuInp   = tr.querySelector('.ab-tu-inp');
      var rmkInp  = tr.querySelector('.ab-remarks-inp');
      var desc = descInp ? descInp.value.trim() : '';
      var uom  = uomSel  ? uomSel.value : '';
      var uc   = ucInp   ? parseFloat(ucInp.value) || 0 : 0;
      var tu   = tuInp   ? parseFloat(tuInp.value) || 0 : 0;
      var rmk  = rmkInp  ? rmkInp.value : '';
      var qtrUnits = [], qtrCosts = [];
      quarters.forEach(function(q, qi) {
        var inp = tr.querySelector('.ab-np-inp[data-qi="' + qi + '"]');
        var u = inp ? parseFloat(inp.value) || 0 : 0;
        qtrUnits.push(u); qtrCosts.push(u * uc);
      });
      if (desc) sRows.push({ description: desc, uom: uom, unitCost: uc, totalUnits: tu, quarterUnits: qtrUnits, quarterCosts: qtrCosts, remarks: rmk });
    });
    sections[sec] = sRows;
  });
  return { sections: sections };
}

// ============================================================================
// SHEET 2: Programmatic Costs  (Option C multi-year layout)
// ============================================================================
// Columns:
//   A: Sr. No.   B: Activity   C: Task Details (Mandatory)   D: UoM   E: Unit Cost   F: Total Units   G: Total Cost
//   H: Total LIC HFL Contribution   I: Government Contribution   J: Beneficiary Contribution    (Convergence)
//   then per year: Q1[U,C,L,B] Q2[U,C,L,B] Q3[U,C,L,B] Q4[U,C,L,B]   (16 cols/year)
//   then "Grand Total" band: one Cost column per year (Y1 Total, Y2 Total, ...)
//   then Remarks
// ============================================================================
function ab_buildProgSheet(frm, quarters, yearGroups, liveData) {
  var rows = liveData.rows || [];
  var Y = yearGroups.length || 1;
  var nQuarters = 4;  // assume 4 quarters per year — yearGroup may have fewer; we pad later

  var firstYearCol = 10;                                    // K = index 10
  var yearBlockSize = nQuarters * 4;                         // 16 cols per year
  var grandStartCol = firstYearCol + Y * yearBlockSize;     // first Y_n Total col
  var remarksCol = grandStartCol + Y;                       // last col
  var totalCols = remarksCol + 1;

  var aoa = [];
  var merges = [];
  var emptyRow = function() { var r = []; for (var i = 0; i < totalCols; i++) r.push(ab_sc('')); return r; };

  // Row 1-2: Logo (A1:C2) + Title banner (D1:J2) — match source.
  var r1 = emptyRow();
  r1[3] = ab_sc('Programmatic Costs', { font: { name: 'Calibri', sz: 14, bold: true, color: { rgb: AB_C_BLUE } }, fill: AB_C_YELLOW, alignment: { horizontal: 'center', vertical: 'center' } });
  aoa.push(r1);
  aoa.push(emptyRow());
  merges.push({ s: { r: 0, c: 0 }, e: { r: 1, c: 2 } });   // A1:C2 logo area
  merges.push({ s: { r: 0, c: 3 }, e: { r: 1, c: 9 } });   // D1:J2 title

  // Row 3: super-group band — Convergence (H3:J3), Year N (per year), Grand Total, Remarks (vertical to row 5)
  var r3 = emptyRow();
  r3[7] = ab_sc('Convergence', { font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_LIGHT, alignment: { horizontal: 'center', vertical: 'center' } });
  merges.push({ s: { r: 2, c: 7 }, e: { r: 3, c: 9 } });   // H3:J4 — Convergence band spans rows 3-4 (sub-cols are in row 5)

  yearGroups.forEach(function(yg, yi) {
    var startC = firstYearCol + yi * yearBlockSize;
    var endC = startC + yearBlockSize - 1;
    r3[startC] = ab_sc(yg.label, { font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: AB_C_WHITE } }, fill: AB_C_BLUE, alignment: { horizontal: 'center', vertical: 'center' } });
    merges.push({ s: { r: 2, c: startC }, e: { r: 2, c: endC } });
  });

  r3[grandStartCol] = ab_sc('Grand Total', { font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_YELLOW, alignment: { horizontal: 'center', vertical: 'center' } });
  if (Y > 1) merges.push({ s: { r: 2, c: grandStartCol }, e: { r: 2, c: grandStartCol + Y - 1 } });

  r3[remarksCol] = ab_sc('Remarks', { font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_GREY_HDR, alignment: { horizontal: 'center', vertical: 'center' } });
  merges.push({ s: { r: 2, c: remarksCol }, e: { r: 4, c: remarksCol } }); // remarks merged across rows 3-5
  aoa.push(r3);

  // Row 4: sub-group — Q1/Q2/Q3/Q4 per year, Y1 Total / Y2 Total under Grand Total
  var r4 = emptyRow();
  yearGroups.forEach(function(yg, yi) {
    var startC = firstYearCol + yi * yearBlockSize;
    [0,1,2,3].forEach(function(qi) {
      var sc_ = startC + qi * 4;
      var fill = (qi % 2 === 0) ? AB_C_BLUE : AB_C_YELLOW;
      var color = (qi % 2 === 0) ? AB_C_WHITE : AB_C_DARK;
      var qLabel = (yg.qs[qi] && yg.qs[qi].quarter) ? yg.qs[qi].quarter : ('Q' + (qi + 1));
      r4[sc_] = ab_sc(qLabel, { font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: color } }, fill: fill, alignment: { horizontal: 'center', vertical: 'center' } });
      merges.push({ s: { r: 3, c: sc_ }, e: { r: 3, c: sc_ + 3 } });
    });
  });
  yearGroups.forEach(function(yg, yi) {
    r4[grandStartCol + yi] = ab_sc(yg.label.replace(/^Year /, 'Y') + ' Total', { font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_YELLOW, alignment: { horizontal: 'center', vertical: 'center' } });
  });
  aoa.push(r4);

  // Row 5: column-name strip (BCBDC0)
  var hdrCellOpts = { font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_GREY_HDR, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } };
  var r5 = emptyRow();
  r5[0] = ab_sc('Sr. No.', hdrCellOpts);
  r5[1] = ab_sc('Activity', hdrCellOpts);
  r5[2] = ab_sc('Task Details (Mandatory)', hdrCellOpts);
  r5[3] = ab_sc('Unit of Measurement', hdrCellOpts);
  r5[4] = ab_sc('Unit Cost', hdrCellOpts);
  r5[5] = ab_sc('Total Units', hdrCellOpts);
  r5[6] = ab_sc('Total Cost', hdrCellOpts);
  r5[7] = ab_sc('Total LIC HFL Contribution', hdrCellOpts);
  r5[8] = ab_sc('Government Contribution', hdrCellOpts);
  r5[9] = ab_sc('Beneficiary Contribution', hdrCellOpts);
  yearGroups.forEach(function(yg, yi) {
    var startC = firstYearCol + yi * yearBlockSize;
    [0,1,2,3].forEach(function(qi) {
      var sc_ = startC + qi * 4;
      r5[sc_]     = ab_sc('Units to Cover',           hdrCellOpts);
      r5[sc_ + 1] = ab_sc('Cost',                     hdrCellOpts);
      r5[sc_ + 2] = ab_sc('LIC HFL Contribution',     hdrCellOpts);
      r5[sc_ + 3] = ab_sc('Beneficiary Contribution', hdrCellOpts);
    });
  });
  yearGroups.forEach(function(yg, yi) {
    r5[grandStartCol + yi] = ab_sc('Cost', hdrCellOpts);
  });
  // r5[remarksCol] — already merged from row 3, leave blank
  aoa.push(r5);

  // Row 6+: Data
  var altRow = AB_C_ALT;
  rows.forEach(function(row, idx) {
    var altOpts = (idx % 2 === 1) ? { fill: altRow } : {};
    var cell = function(v, extra) {
      var o = {};
      if (idx % 2 === 1) o.fill = altRow;
      if (extra) for (var k in extra) o[k] = extra[k];
      return ab_sc(v, o);
    };
    var totalCost = (row.unitCost || 0) * (row.totalUnits || 0);
    var licHflContrib = totalCost - (row.govt || 0) - (row.benf || 0);
    var r = emptyRow();
    r[0] = cell(idx + 1, { alignment: { horizontal: 'center' } });
    r[1] = cell(row.description, { alignment: { horizontal: 'left', wrapText: true } });
    r[2] = cell(row.task, { alignment: { horizontal: 'left', wrapText: true } });
    r[3] = cell(row.uom, { alignment: { horizontal: 'center' } });
    r[4] = cell(row.unitCost, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
    r[5] = cell(row.totalUnits, { alignment: { horizontal: 'right' } });
    r[6] = cell(totalCost, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
    r[7] = cell(licHflContrib, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
    r[8] = cell(row.govt, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
    r[9] = cell(row.benf, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });

    yearGroups.forEach(function(yg, yi) {
      var startC = firstYearCol + yi * yearBlockSize;
      yg.qs.forEach(function(q, qi) {
        var globalQi = quarters.indexOf(q);
        var qd = row.quarters[globalQi] || { units: 0, cost: 0, licHfl: 0, benf: 0 };
        var sc_ = startC + qi * 4;
        r[sc_]     = cell(qd.units || 0, { alignment: { horizontal: 'right' } });
        r[sc_ + 1] = cell(qd.cost  || 0, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
        r[sc_ + 2] = cell(qd.licHfl|| 0, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
        r[sc_ + 3] = cell(qd.benf  || 0, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
      });
      // Year-total cost = sum of this year's quarter costs for the row
      var yTotal = yg.qs.reduce(function(s, q) {
        var globalQi = quarters.indexOf(q);
        var qd = row.quarters[globalQi] || {};
        return s + (qd.cost || 0);
      }, 0);
      r[grandStartCol + yi] = cell(yTotal, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' }, fill: AB_C_LIGHT });
    });
    r[remarksCol] = cell(row.remarks || '', { alignment: { horizontal: 'left', wrapText: true } });
    aoa.push(r);
  });

  // Grand Total row (yellow band)
  var gtFontOpts = { font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_YELLOW, alignment: { horizontal: 'right', vertical: 'center' } };
  var gtTextOpts = { font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_YELLOW, alignment: { horizontal: 'left', vertical: 'center' } };

  var gtTotalCost = 0, gtLicHfl = 0, gtGovt = 0, gtBenf = 0;
  rows.forEach(function(row) {
    var tc = (row.unitCost || 0) * (row.totalUnits || 0);
    gtTotalCost += tc;
    gtLicHfl    += tc - (row.govt || 0) - (row.benf || 0);
    gtGovt      += (row.govt || 0);
    gtBenf      += (row.benf || 0);
  });

  var gt = emptyRow();
  gt[0] = ab_sc('', { fill: AB_C_YELLOW });
  gt[1] = ab_sc('Grand Total', gtTextOpts);
  merges.push({ s: { r: aoa.length, c: 1 }, e: { r: aoa.length, c: 5 } });
  for (var i = 2; i <= 5; i++) gt[i] = ab_sc('', { fill: AB_C_YELLOW });
  gt[6] = ab_sc(gtTotalCost, { numFmt: AB_INR_FMT, font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });
  gt[7] = ab_sc(gtLicHfl,    { numFmt: AB_INR_FMT, font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });
  gt[8] = ab_sc(gtGovt,      { numFmt: AB_INR_FMT, font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });
  gt[9] = ab_sc(gtBenf,      { numFmt: AB_INR_FMT, font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });

  yearGroups.forEach(function(yg, yi) {
    var startC = firstYearCol + yi * yearBlockSize;
    yg.qs.forEach(function(q, qi) {
      var globalQi = quarters.indexOf(q);
      var qU = 0, qC = 0, qL = 0, qB = 0;
      rows.forEach(function(row) {
        var qd = row.quarters[globalQi] || {};
        qU += qd.units  || 0;
        qC += qd.cost   || 0;
        qL += qd.licHfl || 0;
        qB += qd.benf   || 0;
      });
      var sc_ = startC + qi * 4;
      gt[sc_]     = ab_sc(qU, { font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });
      gt[sc_ + 1] = ab_sc(qC, { numFmt: AB_INR_FMT, font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });
      gt[sc_ + 2] = ab_sc(qL, { numFmt: AB_INR_FMT, font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });
      gt[sc_ + 3] = ab_sc(qB, { numFmt: AB_INR_FMT, font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });
    });
    var yTotal = 0;
    rows.forEach(function(row) {
      yg.qs.forEach(function(q) {
        var globalQi = quarters.indexOf(q);
        var qd = row.quarters[globalQi] || {};
        yTotal += qd.cost || 0;
      });
    });
    gt[grandStartCol + yi] = ab_sc(yTotal, { numFmt: AB_INR_FMT, font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });
  });
  gt[remarksCol] = ab_sc('', { fill: AB_C_YELLOW });
  aoa.push(gt);

  var ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = merges;
  ws['!cols'] = ab_buildProgColWidths(Y);
  ws['!rows'] = [{ hpt: 75 }, { hpt: 0 }, { hpt: 26 }, { hpt: 24 }, { hpt: 38 }];
  ws['!freeze'] = { xSplit: 7, ySplit: 5 };
  return { ws: ws };
}

function ab_buildProgColWidths(Y) {
  var cols = [
    { wch: 6 },   // Sr. No.
    { wch: 30 },  // Activity
    { wch: 35 },  // Task Details
    { wch: 14 },  // UoM
    { wch: 14 },  // Unit Cost
    { wch: 11 },  // Total Units
    { wch: 16 },  // Total Cost
    { wch: 18 },  // Total LIC HFL Contribution
    { wch: 16 },  // Govt Contribution
    { wch: 16 }   // Beneficiary Contribution
  ];
  for (var y = 0; y < Y; y++) {
    for (var q = 0; q < 4; q++) {
      cols.push({ wch: 10 }, { wch: 14 }, { wch: 14 }, { wch: 14 });
    }
  }
  for (var y = 0; y < Y; y++) cols.push({ wch: 14 });   // Y_n Total
  cols.push({ wch: 30 });                                // Remarks
  return cols;
}

// ============================================================================
// SHEET 3: Non-Programmatic Costs  (Option C multi-year layout)
// ============================================================================
function ab_buildNonProgSheet(frm, quarters, yearGroups, liveData) {
  var Y = yearGroups.length || 1;
  var firstYearCol = 6;                  // G
  var yearBlockSize = 4 * 2;             // 4 quarters × 2 sub-cols (Units, Cost)
  var grandStartCol = firstYearCol + Y * yearBlockSize;
  var remarksCol = grandStartCol + Y;
  var totalCols = remarksCol + 1;

  var aoa = [];
  var merges = [];
  var emptyRow = function() { var r = []; for (var i = 0; i < totalCols; i++) r.push(ab_sc('')); return r; };

  // Compute section totals first (for the small Cost-Component summary block at top)
  var sectionOrder = ['Human Resource Costs', 'Administration Costs', 'NGO Management Costs'];
  var sectionLabel = { 'Human Resource Costs': 'Human Resource', 'Administration Costs': 'Admin', 'NGO Management Costs': 'NGO Management' };
  var sectionLetter = { 'Human Resource Costs': 'A', 'Administration Costs': 'B', 'NGO Management Costs': 'C' };
  var sections = liveData.sections || {};
  var sectionTotals = {};
  var npGrand = 0;
  sectionOrder.forEach(function(sec) {
    var t = 0;
    (sections[sec] || []).forEach(function(row) {
      var tu = (typeof row.totalUnits === 'number' && row.totalUnits) ? row.totalUnits : (row.quarterUnits || []).reduce(function(s, u) { return s + u; }, 0);
      t += (row.unitCost || 0) * tu;
    });
    sectionTotals[sec] = t;
    npGrand += t;
  });

  // Row 1-2: Logo (A1:C2) + Title banner (D1:end-row2 yellow "Non-Programmatic Costs")
  var r1 = emptyRow();
  r1[3] = ab_sc('Non-Programmatic Costs', { font: { name: 'Calibri', sz: 14, bold: true, color: { rgb: AB_C_BLUE } }, fill: AB_C_YELLOW, alignment: { horizontal: 'center', vertical: 'center' } });
  aoa.push(r1);
  aoa.push(emptyRow());
  merges.push({ s: { r: 0, c: 0 }, e: { r: 1, c: 2 } });               // A1:C2 logo area
  merges.push({ s: { r: 0, c: 3 }, e: { r: 1, c: totalCols - 1 } });   // D1:last2 title

  // Rows 3-6: small Cost-Component summary block at top (matches source rows 1-4)
  // Source layout: B = "" / C-D = Cost Component / E = Amount / F = %ge
  var r3 = emptyRow();
  r3[2] = ab_sc('Cost Component', { font: { name: 'Calibri', sz: 11, bold: true }, fill: AB_C_GREY_HDR, alignment: { horizontal: 'center', vertical: 'center' } });
  r3[4] = ab_sc('Amount', { font: { name: 'Calibri', sz: 11, bold: true }, fill: AB_C_GREY_HDR, alignment: { horizontal: 'center', vertical: 'center' } });
  r3[5] = ab_sc('%age', { font: { name: 'Calibri', sz: 11, bold: true }, fill: AB_C_GREY_HDR, alignment: { horizontal: 'center', vertical: 'center' } });
  merges.push({ s: { r: 2, c: 2 }, e: { r: 2, c: 3 } });
  // we won't push r3 yet — push with header cells filled below

  // Push three summary rows (rows 4,5,6 in 1-indexed → indices 3,4,5)
  // Actually keep aoa simple: we already pushed rows 1-2. Push row 3 = summary header.
  aoa.push(r3);
  sectionOrder.forEach(function(sec, si) {
    var fill = (si % 2 === 0) ? AB_C_WHITE : AB_C_ALT;
    var pct = npGrand > 0 ? (sectionTotals[sec] / npGrand) : 0;
    var sr = emptyRow();
    sr[2] = ab_sc(sectionLabel[sec], { fill: fill, alignment: { horizontal: 'left', vertical: 'center' } });
    merges.push({ s: { r: 2 + si + 1, c: 2 }, e: { r: 2 + si + 1, c: 3 } });
    sr[4] = ab_sc(sectionTotals[sec], { numFmt: AB_INR_FMT, fill: fill, alignment: { horizontal: 'right', vertical: 'center' } });
    sr[5] = ab_sc(pct, { numFmt: AB_PCT_FMT, fill: fill, alignment: { horizontal: 'center', vertical: 'center' } });
    aoa.push(sr);
  });

  // Now build the main table. Header rows (super-band, Q labels, column names) start at the next row.
  // (Row 7 in 1-indexed → index 6)
  var headerStartIdx = aoa.length;

  // Header Row A: super-group band (Year N · Grand Total · Remarks)
  var hA = emptyRow();
  yearGroups.forEach(function(yg, yi) {
    var startC = firstYearCol + yi * yearBlockSize;
    var endC = startC + yearBlockSize - 1;
    hA[startC] = ab_sc(yg.label, { font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: AB_C_WHITE } }, fill: AB_C_BLUE, alignment: { horizontal: 'center', vertical: 'center' } });
    merges.push({ s: { r: headerStartIdx, c: startC }, e: { r: headerStartIdx, c: endC } });
  });
  hA[grandStartCol] = ab_sc('Grand Total', { font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_YELLOW, alignment: { horizontal: 'center', vertical: 'center' } });
  if (Y > 1) merges.push({ s: { r: headerStartIdx, c: grandStartCol }, e: { r: headerStartIdx, c: grandStartCol + Y - 1 } });
  hA[remarksCol] = ab_sc('Remarks', { font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_GREY_HDR, alignment: { horizontal: 'center', vertical: 'center' } });
  merges.push({ s: { r: headerStartIdx, c: remarksCol }, e: { r: headerStartIdx + 2, c: remarksCol } });
  aoa.push(hA);

  // Header Row B: Q1/Q2/Q3/Q4 labels per year + Y_n Total under Grand Total
  var hB = emptyRow();
  yearGroups.forEach(function(yg, yi) {
    var startC = firstYearCol + yi * yearBlockSize;
    [0,1,2,3].forEach(function(qi) {
      var sc_ = startC + qi * 2;
      var fill = (qi % 2 === 0) ? AB_C_BLUE : AB_C_YELLOW;
      var color = (qi % 2 === 0) ? AB_C_WHITE : AB_C_DARK;
      var qLabel = (yg.qs[qi] && yg.qs[qi].quarter) ? yg.qs[qi].quarter : ('Q' + (qi + 1));
      hB[sc_] = ab_sc(qLabel, { font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: color } }, fill: fill, alignment: { horizontal: 'center', vertical: 'center' } });
      merges.push({ s: { r: headerStartIdx + 1, c: sc_ }, e: { r: headerStartIdx + 1, c: sc_ + 1 } });
    });
  });
  yearGroups.forEach(function(yg, yi) {
    hB[grandStartCol + yi] = ab_sc(yg.label.replace(/^Year /, 'Y') + ' Total', { font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_YELLOW, alignment: { horizontal: 'center', vertical: 'center' } });
  });
  aoa.push(hB);

  // Header Row C: column-name strip
  var hdrCellOpts = { font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_GREY_HDR, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } };
  var hC = emptyRow();
  hC[0] = ab_sc('Sr. No.', hdrCellOpts);
  hC[1] = ab_sc('Particulars', hdrCellOpts);
  hC[2] = ab_sc('Unit of Measurement', hdrCellOpts);
  hC[3] = ab_sc('Unit Cost', hdrCellOpts);
  hC[4] = ab_sc('Total Units', hdrCellOpts);
  hC[5] = ab_sc('Total Cost', hdrCellOpts);
  yearGroups.forEach(function(yg, yi) {
    var startC = firstYearCol + yi * yearBlockSize;
    [0,1,2,3].forEach(function(qi) {
      var sc_ = startC + qi * 2;
      hC[sc_]     = ab_sc('Units to Cover', hdrCellOpts);
      hC[sc_ + 1] = ab_sc('Cost',           hdrCellOpts);
    });
  });
  yearGroups.forEach(function(yg, yi) {
    hC[grandStartCol + yi] = ab_sc('Cost', hdrCellOpts);
  });
  aoa.push(hC);

  // Section A / B / C — banner + rows + TOTAL (X)
  sectionOrder.forEach(function(sec) {
    var sRows = sections[sec] || [];
    var letter = sectionLetter[sec];

    // Section banner (full-width, dark grey)
    var sb = emptyRow();
    sb[0] = ab_sc(letter, { font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: AB_C_WHITE } }, fill: AB_C_DARK, alignment: { horizontal: 'center', vertical: 'center' } });
    sb[1] = ab_sc(sec,    { font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: AB_C_WHITE } }, fill: AB_C_DARK, alignment: { horizontal: 'left', vertical: 'center' } });
    for (var i = 2; i < totalCols; i++) sb[i] = ab_sc('', { fill: AB_C_DARK });
    merges.push({ s: { r: aoa.length, c: 1 }, e: { r: aoa.length, c: totalCols - 1 } });
    aoa.push(sb);

    // Data rows
    sRows.forEach(function(row, idx) {
      var fillBase = (idx % 2 === 1) ? AB_C_ALT : AB_C_WHITE;
      var cell = function(v, extra) {
        var o = { fill: fillBase };
        if (extra) for (var k in extra) o[k] = extra[k];
        return ab_sc(v, o);
      };
      var tu = (typeof row.totalUnits === 'number' && row.totalUnits) ? row.totalUnits : (row.quarterUnits || []).reduce(function(s, u) { return s + u; }, 0);
      var tc = (row.unitCost || 0) * tu;
      var r = emptyRow();
      r[0] = cell(idx + 1, { alignment: { horizontal: 'center' } });
      r[1] = cell(row.description, { alignment: { horizontal: 'left', wrapText: true } });
      r[2] = cell(row.uom, { alignment: { horizontal: 'center' } });
      r[3] = cell(row.unitCost, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
      r[4] = cell(tu, { alignment: { horizontal: 'right' } });
      r[5] = cell(tc, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });

      yearGroups.forEach(function(yg, yi) {
        var startC = firstYearCol + yi * yearBlockSize;
        var yTotalCost = 0;
        yg.qs.forEach(function(q, qi) {
          var globalQi = quarters.indexOf(q);
          var u = (row.quarterUnits || [])[globalQi] || 0;
          var c = u * (row.unitCost || 0);
          yTotalCost += c;
          var sc_ = startC + qi * 2;
          r[sc_]     = cell(u, { alignment: { horizontal: 'right' } });
          r[sc_ + 1] = cell(c, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
        });
        r[grandStartCol + yi] = cell(yTotalCost, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' }, fill: AB_C_LIGHT });
      });
      r[remarksCol] = cell(row.remarks || '', { alignment: { horizontal: 'left', wrapText: true } });
      aoa.push(r);
    });

    // Section TOTAL (X) row — light grey fill
    var totalOpts = { font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_GREY_HDR, alignment: { horizontal: 'right', vertical: 'center' } };
    var totalLabelOpts = { font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_GREY_HDR, alignment: { horizontal: 'left', vertical: 'center' } };
    var st = emptyRow();
    st[0] = ab_sc('', { fill: AB_C_GREY_HDR });
    st[1] = ab_sc('TOTAL (' + letter + ')', totalLabelOpts);
    for (var i = 2; i <= 3; i++) st[i] = ab_sc('', { fill: AB_C_GREY_HDR });
    var sTotalUnits = 0, sTotalCost = 0;
    sRows.forEach(function(row) {
      var tu = (typeof row.totalUnits === 'number' && row.totalUnits) ? row.totalUnits : (row.quarterUnits || []).reduce(function(s, u) { return s + u; }, 0);
      sTotalUnits += tu;
      sTotalCost  += (row.unitCost || 0) * tu;
    });
    st[4] = ab_sc(sTotalUnits, { font: totalOpts.font, fill: AB_C_GREY_HDR, alignment: totalOpts.alignment });
    st[5] = ab_sc(sTotalCost,  { numFmt: AB_INR_FMT, font: totalOpts.font, fill: AB_C_GREY_HDR, alignment: totalOpts.alignment });

    yearGroups.forEach(function(yg, yi) {
      var startC = firstYearCol + yi * yearBlockSize;
      var yTotal = 0;
      yg.qs.forEach(function(q, qi) {
        var globalQi = quarters.indexOf(q);
        var qU = 0, qC = 0;
        sRows.forEach(function(row) {
          var u = (row.quarterUnits || [])[globalQi] || 0;
          qU += u; qC += u * (row.unitCost || 0);
        });
        yTotal += qC;
        var sc_ = startC + qi * 2;
        st[sc_]     = ab_sc(qU, { font: totalOpts.font, fill: AB_C_GREY_HDR, alignment: totalOpts.alignment });
        st[sc_ + 1] = ab_sc(qC, { numFmt: AB_INR_FMT, font: totalOpts.font, fill: AB_C_GREY_HDR, alignment: totalOpts.alignment });
      });
      st[grandStartCol + yi] = ab_sc(yTotal, { numFmt: AB_INR_FMT, font: totalOpts.font, fill: AB_C_GREY_HDR, alignment: totalOpts.alignment });
    });
    st[remarksCol] = ab_sc('', { fill: AB_C_GREY_HDR });
    aoa.push(st);
  });

  // Final Grand Total row (yellow band)
  var gtFontOpts = { font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_YELLOW, alignment: { horizontal: 'right', vertical: 'center' } };
  var gtLabelOpts = { font: { name: 'Calibri', sz: 13, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_YELLOW, alignment: { horizontal: 'left', vertical: 'center' } };
  var gt = emptyRow();
  gt[0] = ab_sc('', { fill: AB_C_YELLOW });
  gt[1] = ab_sc('Grand Total', gtLabelOpts);
  for (var i = 2; i <= 3; i++) gt[i] = ab_sc('', { fill: AB_C_YELLOW });
  // Grand-total Total Units = sum of every section's user-input total_units
  var npGrandUnits = 0;
  sectionOrder.forEach(function(sec) {
    (sections[sec] || []).forEach(function(row) {
      var tu = (typeof row.totalUnits === 'number' && row.totalUnits) ? row.totalUnits : (row.quarterUnits || []).reduce(function(s, u) { return s + u; }, 0);
      npGrandUnits += tu;
    });
  });
  gt[4] = ab_sc(npGrandUnits, { font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });
  gt[5] = ab_sc(npGrand,      { numFmt: AB_INR_FMT, font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });

  yearGroups.forEach(function(yg, yi) {
    var startC = firstYearCol + yi * yearBlockSize;
    var yTotal = 0;
    yg.qs.forEach(function(q, qi) {
      var globalQi = quarters.indexOf(q);
      var qU = 0, qC = 0;
      sectionOrder.forEach(function(sec) {
        (sections[sec] || []).forEach(function(row) {
          var u = (row.quarterUnits || [])[globalQi] || 0;
          qU += u; qC += u * (row.unitCost || 0);
        });
      });
      yTotal += qC;
      var sc_ = startC + qi * 2;
      gt[sc_]     = ab_sc(qU, { font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });
      gt[sc_ + 1] = ab_sc(qC, { numFmt: AB_INR_FMT, font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });
    });
    gt[grandStartCol + yi] = ab_sc(yTotal, { numFmt: AB_INR_FMT, font: gtFontOpts.font, fill: AB_C_YELLOW, alignment: gtFontOpts.alignment });
  });
  gt[remarksCol] = ab_sc('', { fill: AB_C_YELLOW });
  aoa.push(gt);

  var ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = merges;
  ws['!cols'] = ab_buildNonProgColWidths(Y);
  ws['!rows'] = [
    { hpt: 38 }, { hpt: 38 },        // logo rows
    { hpt: 22 }, { hpt: 22 }, { hpt: 22 }, { hpt: 22 },  // top mini summary
    { hpt: 26 }, { hpt: 24 }, { hpt: 38 }                 // header band rows
  ];
  ws['!freeze'] = { xSplit: 6, ySplit: headerStartIdx + 3 };
  return { ws: ws };
}

function ab_buildNonProgColWidths(Y) {
  var cols = [
    { wch: 6 },   // Sr. No.
    { wch: 35 },  // Particulars
    { wch: 16 },  // UoM
    { wch: 14 },  // Unit Cost
    { wch: 11 },  // Total Units
    { wch: 16 }   // Total Cost
  ];
  for (var y = 0; y < Y; y++) {
    for (var q = 0; q < 4; q++) cols.push({ wch: 10 }, { wch: 14 });
  }
  for (var y = 0; y < Y; y++) cols.push({ wch: 14 });
  cols.push({ wch: 30 });
  return cols;
}

// ============================================================================
// SHEET 1: Budget Summary  (matches source rows 1-28 + multi-year quarterly stack)
// ============================================================================
function ab_buildBudgetSummarySheet(frm, quarters, yearGroups, progLive, npLive, meta) {
  meta = meta || {};
  // Pre-compute totals
  var sectionTotals = { 'Human Resource Costs': 0, 'Administration Costs': 0, 'NGO Management Costs': 0 };
  Object.keys(sectionTotals).forEach(function(sec) {
    (npLive.sections[sec] || []).forEach(function(row) {
      var tu = (typeof row.totalUnits === 'number' && row.totalUnits) ? row.totalUnits : (row.quarterUnits || []).reduce(function(s, u) { return s + u; }, 0);
      sectionTotals[sec] += (row.unitCost || 0) * tu;
    });
  });
  // PROGRAMMATIC: gross total cost (used internally for some calcs) AND LIC-only (the source convention).
  var progGrossTotal = (progLive.rows || []).reduce(function(s, r) { return s + (r.unitCost || 0) * (r.totalUnits || 0); }, 0);
  var progLicTotal   = (progLive.rows || []).reduce(function(s, r) {
    var grossRow = (r.unitCost || 0) * (r.totalUnits || 0);
    var govt = r.govt || 0;
    var benf = r.benf || 0;
    return s + (grossRow - govt - benf);
  }, 0);
  // Source convention: Programmatic in Budget Summary = LIC HFL Contribution only.
  var progTotal  = progLicTotal;
  var hrTotal    = sectionTotals['Human Resource Costs'];
  var adminTotal = sectionTotals['Administration Costs'];
  var ngoTotal   = sectionTotals['NGO Management Costs'];
  var totalDirect   = progTotal + hrTotal;
  var totalIndirect = adminTotal + ngoTotal;
  var grandTotal    = totalDirect + totalIndirect;  // = LIC HFL Contribution total (per source convention)

  var govtContrib = (progLive.rows || []).reduce(function(s, r) { return s + (r.govt || 0); }, 0);
  var benfContrib = (progLive.rows || []).reduce(function(s, r) { return s + (r.benf || 0); }, 0);
  var licHflContrib = grandTotal - govtContrib - benfContrib;

  // Per-year quarterly totals
  var perYearQ = yearGroups.map(function(yg) {
    return yg.qs.map(function(q) {
      var qi = quarters.indexOf(q);
      var qProg = 0, qNp = 0;
      (progLive.rows || []).forEach(function(r) { qProg += ((r.quarters[qi] || {}).cost) || 0; });
      Object.keys(npLive.sections || {}).forEach(function(sec) {
        (npLive.sections[sec] || []).forEach(function(row) {
          var u = (row.quarterUnits || [])[qi] || 0;
          qNp += u * (row.unitCost || 0);
        });
      });
      return { quarter: q.quarter, total: qProg + qNp };
    });
  });

  var aoa = [];
  var merges = [];
  var totalCols = 13;
  var emptyRow = function() { var r = []; for (var i = 0; i < totalCols; i++) r.push(ab_sc('')); return r; };

  // Row 1-2: Logo (A1:F2 — wider on Summary) + Title banner G1:M2 yellow
  var r1 = emptyRow();
  r1[6] = ab_sc('Proposed Budget Statement', { font: { name: 'Calibri', sz: 16, bold: true, color: { rgb: AB_C_BLUE } }, fill: AB_C_YELLOW, alignment: { horizontal: 'center', vertical: 'center' } });
  aoa.push(r1);
  aoa.push(emptyRow());
  merges.push({ s: { r: 0, c: 0 }, e: { r: 1, c: 5 } });   // A1:F2 logo
  merges.push({ s: { r: 0, c: 6 }, e: { r: 1, c: 12 } });  // G1:M2 title

  // Row 3 (blank)
  aoa.push(emptyRow());

  // Row 4: Section headers (dark grey)
  var darkHdrOpts = { font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: AB_C_WHITE } }, fill: AB_C_DARK, alignment: { horizontal: 'center', vertical: 'center' } };
  var r4 = emptyRow();
  r4[1] = ab_sc('Project Details', darkHdrOpts);
  merges.push({ s: { r: 3, c: 1 }, e: { r: 3, c: 7 } });   // B4:H4
  r4[9] = ab_sc('Cost Per Beneficiary', darkHdrOpts);
  merges.push({ s: { r: 3, c: 9 }, e: { r: 3, c: 11 } });  // J4:L4
  aoa.push(r4);

  var labelOpts  = { fill: AB_C_LIGHT, alignment: { horizontal: 'center', vertical: 'center' } };
  var labelLeft  = { fill: AB_C_LIGHT, alignment: { horizontal: 'left', vertical: 'center' } };
  var valLeft    = { alignment: { horizontal: 'left', vertical: 'center', wrapText: true } };
  var valRight   = { alignment: { horizontal: 'right', vertical: 'center' } };
  var totalBenef = (typeof meta.total_beneficiaries === 'number' && meta.total_beneficiaries)
    ? meta.total_beneficiaries
    : (parseFloat(frm.doc.custom_direct_beneficiary) || parseFloat(frm.doc.total_beneficiaries) || 0);

  // Row 5: I  Name of the Project    +  Total Project Costs
  var r5 = emptyRow();
  r5[1] = ab_sc('I', labelOpts);
  r5[2] = ab_sc('Name of the Project', labelLeft);
  merges.push({ s: { r: 4, c: 2 }, e: { r: 4, c: 4 } });
  r5[5] = ab_sc(frm.doc.project_name || '', valLeft);
  merges.push({ s: { r: 4, c: 5 }, e: { r: 4, c: 7 } });
  r5[9]  = ab_sc('Total Project Costs', labelLeft);
  r5[10] = ab_sc(grandTotal, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right', vertical: 'center' } });
  merges.push({ s: { r: 4, c: 10 }, e: { r: 4, c: 11 } });
  aoa.push(r5);

  // Row 6: II Project Duration       +  Total Beneficiaries
  var r6 = emptyRow();
  r6[1] = ab_sc('II', labelOpts);
  r6[2] = ab_sc('Project Duration', labelLeft);
  merges.push({ s: { r: 5, c: 2 }, e: { r: 5, c: 4 } });
  r6[5] = ab_sc(meta.duration_str || ((frm.doc.start_date || '') + ' to ' + (frm.doc.end_date || '')), valLeft);
  merges.push({ s: { r: 5, c: 5 }, e: { r: 5, c: 7 } });
  r6[9]  = ab_sc('Total Beneficiaries', labelLeft);
  r6[10] = ab_sc(totalBenef, valRight);
  r6[11] = ab_sc('Individuals', valLeft);
  aoa.push(r6);

  // Row 7: III Intervention States    +  Cost per Beneficiary
  var r7 = emptyRow();
  r7[1] = ab_sc('III', labelOpts);
  r7[2] = ab_sc('Project Intervention States', labelLeft);
  merges.push({ s: { r: 6, c: 2 }, e: { r: 6, c: 4 } });
  r7[5] = ab_sc(meta.states_str || frm.doc.state || '', valLeft);
  merges.push({ s: { r: 6, c: 5 }, e: { r: 6, c: 7 } });
  r7[9]  = ab_sc('Cost per Beneficiary', labelLeft);
  r7[10] = ab_sc(totalBenef > 0 ? grandTotal / totalBenef : 0, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right', vertical: 'center' } });
  merges.push({ s: { r: 6, c: 10 }, e: { r: 6, c: 11 } });
  aoa.push(r7);

  // Row 8: IV  No. of Intervention Site (count of districts on the GAF)
  var r8 = emptyRow();
  r8[1] = ab_sc('IV', labelOpts);
  r8[2] = ab_sc('No. of Intervention Sites', labelLeft);
  merges.push({ s: { r: 7, c: 2 }, e: { r: 7, c: 4 } });
  var sitesLabel = (meta.intervention_sites || 0) + (meta.intervention_sites === 1 ? ' District' : ' Districts');
  r8[5] = ab_sc(sitesLabel, valLeft);
  merges.push({ s: { r: 7, c: 5 }, e: { r: 7, c: 7 } });
  aoa.push(r8);

  // Row 9: V  Implementing Partner (NGO name, not the NGO id)
  var r9 = emptyRow();
  r9[1] = ab_sc('V', labelOpts);
  r9[2] = ab_sc('Implementing Partner', labelLeft);
  merges.push({ s: { r: 8, c: 2 }, e: { r: 8, c: 4 } });
  r9[5] = ab_sc(meta.ngo_name || frm.doc.ngo || '', valLeft);
  merges.push({ s: { r: 8, c: 5 }, e: { r: 8, c: 7 } });
  aoa.push(r9);

  // Row 10: blank
  aoa.push(emptyRow());

  // Row 11: Budget Breakup (full-width dark band on left)
  var r11 = emptyRow();
  r11[1] = ab_sc('Budget Breakup', darkHdrOpts);
  merges.push({ s: { r: 10, c: 1 }, e: { r: 10, c: 6 } });
  r11[9] = ab_sc('Quarterly Breakup of Costs (per year)', darkHdrOpts);
  merges.push({ s: { r: 10, c: 9 }, e: { r: 10, c: 11 } });
  aoa.push(r11);

  // Row 12: Direct Costs heading row + Quarterly Breakup column heads
  var grHdr = { font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_GREY_HDR, alignment: { horizontal: 'center', vertical: 'center' } };
  var r12 = emptyRow();
  r12[1] = ab_sc('A', grHdr);
  r12[2] = ab_sc('DIRECT COSTS', { font: grHdr.font, fill: AB_C_GREY_HDR, alignment: { horizontal: 'left', vertical: 'center' } });
  merges.push({ s: { r: 11, c: 2 }, e: { r: 11, c: 4 } });
  r12[5] = ab_sc('Amount', grHdr);
  r12[6] = ab_sc('%age', grHdr);
  r12[8] = ab_sc('Year', grHdr);
  r12[9] = ab_sc('Quarter', grHdr);
  r12[10] = ab_sc('Amount', grHdr);
  r12[11] = ab_sc('%age', grHdr);
  aoa.push(r12);

  var labelLeftPlain = { alignment: { horizontal: 'left', vertical: 'center' } };
  // Row 13: i  Programmatic Costs
  var r13 = emptyRow();
  r13[1] = ab_sc('i', labelOpts);
  r13[2] = ab_sc('Programmatic Costs', labelLeftPlain);
  merges.push({ s: { r: 12, c: 2 }, e: { r: 12, c: 4 } });
  r13[5] = ab_sc(progTotal, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
  r13[6] = ab_sc(grandTotal > 0 ? progTotal / grandTotal : 0, { numFmt: AB_PCT_FMT, alignment: { horizontal: 'center' } });
  aoa.push(r13);

  // Row 14: ii Human Resource Costs
  var r14 = emptyRow();
  r14[1] = ab_sc('ii', labelOpts);
  r14[2] = ab_sc('Human Resource Costs', labelLeftPlain);
  merges.push({ s: { r: 13, c: 2 }, e: { r: 13, c: 4 } });
  r14[5] = ab_sc(hrTotal, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
  r14[6] = ab_sc(grandTotal > 0 ? hrTotal / grandTotal : 0, { numFmt: AB_PCT_FMT, alignment: { horizontal: 'center' } });
  aoa.push(r14);

  // Row 15: TOTAL DIRECT COST (A)
  var totLine = { font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_LIGHT };
  var r15 = emptyRow();
  r15[2] = ab_sc('TOTAL DIRECT COST (A)', { font: totLine.font, fill: AB_C_LIGHT, alignment: { horizontal: 'left', vertical: 'center' } });
  merges.push({ s: { r: 14, c: 2 }, e: { r: 14, c: 4 } });
  r15[5] = ab_sc(totalDirect, { numFmt: AB_INR_FMT, font: totLine.font, fill: AB_C_LIGHT, alignment: { horizontal: 'right', vertical: 'center' } });
  r15[6] = ab_sc(grandTotal > 0 ? totalDirect / grandTotal : 0, { numFmt: AB_PCT_FMT, font: totLine.font, fill: AB_C_LIGHT, alignment: { horizontal: 'center', vertical: 'center' } });
  aoa.push(r15);

  // Row 16: B  INDIRECT COSTS heading
  var r16 = emptyRow();
  r16[1] = ab_sc('B', grHdr);
  r16[2] = ab_sc('INDIRECT COSTS', { font: grHdr.font, fill: AB_C_GREY_HDR, alignment: { horizontal: 'left', vertical: 'center' } });
  merges.push({ s: { r: 15, c: 2 }, e: { r: 15, c: 4 } });
  r16[5] = ab_sc('Amount', grHdr);
  r16[6] = ab_sc('%age', grHdr);
  aoa.push(r16);

  // Row 17: i  Admin Costs
  var r17 = emptyRow();
  r17[1] = ab_sc('i', labelOpts);
  r17[2] = ab_sc('Admin Costs', labelLeftPlain);
  merges.push({ s: { r: 16, c: 2 }, e: { r: 16, c: 4 } });
  r17[5] = ab_sc(adminTotal, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
  r17[6] = ab_sc(grandTotal > 0 ? adminTotal / grandTotal : 0, { numFmt: AB_PCT_FMT, alignment: { horizontal: 'center' } });
  aoa.push(r17);

  // Row 18: ii NGO Management Costs
  var r18 = emptyRow();
  r18[1] = ab_sc('ii', labelOpts);
  r18[2] = ab_sc('NGO Management Costs', labelLeftPlain);
  merges.push({ s: { r: 17, c: 2 }, e: { r: 17, c: 4 } });
  r18[5] = ab_sc(ngoTotal, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
  r18[6] = ab_sc(grandTotal > 0 ? ngoTotal / grandTotal : 0, { numFmt: AB_PCT_FMT, alignment: { horizontal: 'center' } });
  aoa.push(r18);

  // Row 19: TOTAL INDIRECT COSTS (B)
  var r19 = emptyRow();
  r19[2] = ab_sc('TOTAL INDIRECT COSTS (B)', { font: totLine.font, fill: AB_C_LIGHT, alignment: { horizontal: 'left', vertical: 'center' } });
  merges.push({ s: { r: 18, c: 2 }, e: { r: 18, c: 4 } });
  r19[5] = ab_sc(totalIndirect, { numFmt: AB_INR_FMT, font: totLine.font, fill: AB_C_LIGHT, alignment: { horizontal: 'right', vertical: 'center' } });
  r19[6] = ab_sc(grandTotal > 0 ? totalIndirect / grandTotal : 0, { numFmt: AB_PCT_FMT, font: totLine.font, fill: AB_C_LIGHT, alignment: { horizontal: 'center', vertical: 'center' } });
  aoa.push(r19);

  // Row 20: GRAND TOTAL (A+B) — yellow band
  var ytLine = { font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: AB_C_DARK } }, fill: AB_C_YELLOW };
  var r20 = emptyRow();
  r20[2] = ab_sc('GRAND TOTAL (A+B)', { font: ytLine.font, fill: AB_C_YELLOW, alignment: { horizontal: 'left', vertical: 'center' } });
  merges.push({ s: { r: 19, c: 2 }, e: { r: 19, c: 4 } });
  r20[5] = ab_sc(grandTotal, { numFmt: AB_INR_FMT, font: ytLine.font, fill: AB_C_YELLOW, alignment: { horizontal: 'right', vertical: 'center' } });
  r20[6] = ab_sc(1, { numFmt: AB_PCT_FMT, font: ytLine.font, fill: AB_C_YELLOW, alignment: { horizontal: 'center', vertical: 'center' } });
  aoa.push(r20);

  // Now stitch the right-side Quarterly Breakup column data into the rows we already pushed.
  // We placed headers in row 12 (idx 11). Quarter rows start at row 13 (idx 12).
  var qRowStart = 12;
  yearGroups.forEach(function(yg, yi) {
    var qLines = perYearQ[yi] || [];
    var quarterNames = ['First Quarter', 'Second Quarter', 'Third Quarter', 'Fourth Quarter'];
    var yearTotal = qLines.reduce(function(s, x) { return s + x.total; }, 0);

    // Year label cell (merged across 4 rows)
    var yLabelRow = qRowStart;
    if (aoa[yLabelRow]) {
      aoa[yLabelRow][8] = ab_sc(yg.label, { fill: AB_C_LIGHT, font: { name: 'Calibri', sz: 11, bold: true }, alignment: { horizontal: 'center', vertical: 'center' } });
      merges.push({ s: { r: yLabelRow, c: 8 }, e: { r: yLabelRow + 3, c: 8 } });
    }
    qLines.forEach(function(qLine, qi) {
      var rIdx = qRowStart + qi;
      if (!aoa[rIdx]) return;
      var fill = (qi % 2 === 0) ? AB_C_WHITE : AB_C_ALT;
      aoa[rIdx][9]  = ab_sc(quarterNames[qi] || ('Quarter ' + (qi + 1)), { fill: fill, alignment: { horizontal: 'left', vertical: 'center' } });
      aoa[rIdx][10] = ab_sc(qLine.total, { numFmt: AB_INR_FMT, fill: fill, alignment: { horizontal: 'right', vertical: 'center' } });
      aoa[rIdx][11] = ab_sc(grandTotal > 0 ? qLine.total / grandTotal : 0, { numFmt: AB_PCT_FMT, fill: fill, alignment: { horizontal: 'center', vertical: 'center' } });
    });
    // Year total subtotal row — append after the 4 quarters
    var ytIdx = qRowStart + 4;
    if (!aoa[ytIdx]) aoa[ytIdx] = emptyRow();
    aoa[ytIdx][8] = ab_sc(yg.label + ' Total', { font: ytLine.font, fill: AB_C_YELLOW, alignment: { horizontal: 'center', vertical: 'center' } });
    aoa[ytIdx][9] = ab_sc('', { fill: AB_C_YELLOW });
    aoa[ytIdx][10] = ab_sc(yearTotal, { numFmt: AB_INR_FMT, font: ytLine.font, fill: AB_C_YELLOW, alignment: { horizontal: 'right', vertical: 'center' } });
    aoa[ytIdx][11] = ab_sc(grandTotal > 0 ? yearTotal / grandTotal : 0, { numFmt: AB_PCT_FMT, font: ytLine.font, fill: AB_C_YELLOW, alignment: { horizontal: 'center', vertical: 'center' } });
    qRowStart = ytIdx + 1;
  });

  // Row 21: blank
  aoa.push(emptyRow());

  // Row 22: Cost Sharing Details header (full row)
  var rCS = emptyRow();
  rCS[1] = ab_sc('Cost Sharing Details', darkHdrOpts);
  merges.push({ s: { r: aoa.length, c: 1 }, e: { r: aoa.length, c: 6 } });
  aoa.push(rCS);

  // Cost Sharing column heads row
  var rCSh = emptyRow();
  rCSh[1] = ab_sc('Cost Sharing', { font: grHdr.font, fill: AB_C_GREY_HDR, alignment: { horizontal: 'left', vertical: 'center' } });
  merges.push({ s: { r: aoa.length, c: 1 }, e: { r: aoa.length, c: 4 } });
  rCSh[5] = ab_sc('Amount', grHdr);
  rCSh[6] = ab_sc('%age', grHdr);
  aoa.push(rCSh);

  // Government Convergence
  var csRow = function(label, amount) {
    var r = emptyRow();
    r[1] = ab_sc(label, labelLeftPlain);
    merges.push({ s: { r: aoa.length, c: 1 }, e: { r: aoa.length, c: 4 } });
    r[5] = ab_sc(amount, { numFmt: AB_INR_FMT, alignment: { horizontal: 'right' } });
    r[6] = ab_sc(grandTotal > 0 ? amount / grandTotal : 0, { numFmt: AB_PCT_FMT, alignment: { horizontal: 'center' } });
    aoa.push(r);
  };
  csRow('Government Convergence', govtContrib);
  csRow('Community/Beneficiary/NGO/Other', benfContrib);
  csRow('LIC HFL Contribution', licHflContrib);

  // TOTAL row (yellow)
  var rT = emptyRow();
  rT[1] = ab_sc('TOTAL', { font: ytLine.font, fill: AB_C_YELLOW, alignment: { horizontal: 'left', vertical: 'center' } });
  merges.push({ s: { r: aoa.length, c: 1 }, e: { r: aoa.length, c: 4 } });
  rT[5] = ab_sc(grandTotal, { numFmt: AB_INR_FMT, font: ytLine.font, fill: AB_C_YELLOW, alignment: { horizontal: 'right', vertical: 'center' } });
  rT[6] = ab_sc(1, { numFmt: AB_PCT_FMT, font: ytLine.font, fill: AB_C_YELLOW, alignment: { horizontal: 'center', vertical: 'center' } });
  aoa.push(rT);

  // Column widths: A,B narrow (numerals), C-H content, I-M for quarterly breakup
  var ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = merges;
  ws['!cols'] = [
    { wch: 4 },   // A
    { wch: 5 },   // B (Roman numerals)
    { wch: 26 },  // C  Project Details labels
    { wch: 14 },  // D
    { wch: 14 },  // E
    { wch: 14 },  // F
    { wch: 4 },   // G
    { wch: 4 },   // H spacer
    { wch: 8 },   // I  Year
    { wch: 18 },  // J  Quarter
    { wch: 16 },  // K  Amount
    { wch: 10 },  // L  %age
    { wch: 4 }    // M
  ];
  ws['!rows'] = [{ hpt: 38 }, { hpt: 38 }];
  return { ws: ws };
}
