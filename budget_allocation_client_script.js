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
      ab_getList('Project Budget Planning', ['name', 'description', 'budget_head', 'sub_budget_head', 'fund_source', 'total_planned_budget', 'assumption'], { project_proposal: frm.doc.name }, 200, 'budget_head asc, description asc'),
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
        assumption: rec.assumption || '',
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

    // Try to extract govt and benf contributions from total_planned_budget if available
    // For now, default to 0 (these should be entered by user)
  });

  // Auto-populate from Activity KPI selections (if any)
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

  // Row 1: Frozen headers (empty for cols 0-4) + Convergence spanning cols 5-9 + Quarter spans + Remarks
  html += '<tr class="ab-header-row-1">';
  html += '<th colspan="5" class="ab-frozen-header" style="position:sticky;left:0;top:0;z-index:20;background:#8B1A1A;color:white;font-weight:700;">Activity Details</th>';
  html += '<th colspan="5" class="ab-convergence-header">Convergence</th>';
  quarters.forEach(function(q, qi) {
    html += '<th colspan="4" class="ab-quarter-header">' + q.quarter + '</th>';
  });
  html += '<th class="ab-quarter-header">&nbsp;</th></tr>';

  // Row 2: Column headers
  html += '<tr class="ab-header-row-2">';
  html += '<th class="ab-frozen ab-sr-hdr ab-frozen-last" style="left:0px;width:35px;min-width:35px;">Sr.</th>';
  html += '<th class="ab-frozen ab-col-hdr" style="left:35px;width:150px;min-width:150px;">Activity</th>';
  html += '<th class="ab-frozen ab-col-hdr" style="left:185px;width:120px;min-width:120px;">Task Details</th>';
  html += '<th class="ab-frozen ab-col-hdr" style="left:305px;width:55px;min-width:55px;">UoM</th>';
  html += '<th class="ab-frozen ab-col-hdr ab-frozen-last" style="left:360px;width:75px;min-width:75px;">Unit Cost</th>';
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

  html += '<th class="ab-remarks-hdr" style="width:100px;min-width:100px;">Remarks</th>';
  html += '</tr></thead><tbody>';

  // Data rows
  rows.forEach(function(row, idx) {
    html += ab_buildProgRow(row, quarters, idx);
  });

  // Grand total row
  html += ab_buildProgGrandTotal(rows, quarters);

  html += '</tbody></table></div>';
  return html;
}

function ab_buildProgRow(row, quarters, idx) {
  var html = '<tr class="ab-data-row" data-idx="' + idx + '">';

  // Frozen columns (0-4): Sr, Activity, Task Details, UoM, Unit Cost
  html += '<td class="ab-frozen ab-sr ab-frozen-last" style="left:0px;width:35px;min-width:35px;">' + (idx + 1) + '</td>';
  html += '<td class="ab-frozen ab-desc-cell" style="left:35px;width:150px;min-width:150px;text-align:left;" title="' + ab_he(row.description) + '">' + ab_he(row.description) + '</td>';
  html += '<td class="ab-frozen ab-editable" style="left:185px;width:120px;min-width:120px;text-align:left;"><input type="text" class="ab-inp ab-task-inp" data-idx="' + idx + '" value="' + ab_he(row.assumption || '') + '" placeholder="Task details..." /></td>';
  html += '<td class="ab-frozen" style="left:305px;width:55px;min-width:55px;">' + ab_he(row.uomName || 'Numbers') + '</td>';
  html += '<td class="ab-frozen ab-editable ab-frozen-last" style="left:360px;width:75px;min-width:75px;"><input type="number" class="ab-inp ab-uc-inp" data-idx="' + idx + '" value="' + (row.unit_cost || 0) + '" /></td>';

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

  html += '<td class="ab-editable"><input type="text" class="ab-inp ab-remarks-inp" data-idx="' + idx + '" value="' + ab_he(row.remarks || '') + '" placeholder="Remarks..." /></td>';
  html += '</tr>';
  return html;
}

function ab_buildProgGrandTotal(rows, quarters) {
  var html = '<tr class="ab-grand-total-row">';

  // Frozen columns (5): Sr, Activity, Task Details, UoM, Unit Cost
  html += '<td class="ab-frozen ab-sr ab-frozen-last" style="left:0px;width:35px;min-width:35px;">GT</td>';
  html += '<td class="ab-frozen" style="left:35px;width:150px;min-width:150px;text-align:left;font-weight:700;">GRAND TOTAL</td>';
  html += '<td class="ab-frozen" style="left:185px;width:120px;min-width:120px;"></td>';
  html += '<td class="ab-frozen" style="left:305px;width:55px;min-width:55px;"></td>';
  html += '<td class="ab-frozen" style="left:360px;width:75px;min-width:75px;"></td>';

  // Non-frozen columns (5): Total Units, Total Cost, LIC HFL, Govt, Benf
  html += '<td class="ab-gt-cell" style="width:70px;min-width:70px;"></td>';
  html += '<td class="ab-gt-cell" style="width:85px;min-width:85px;">0</td>';
  html += '<td class="ab-gt-cell" style="width:100px;min-width:100px;">0</td>';
  html += '<td class="ab-gt-cell" style="width:85px;min-width:85px;">0</td>';
  html += '<td class="ab-gt-cell" style="width:85px;min-width:85px;">0</td>';

  // Per-quarter totals
  quarters.forEach(function(q, qi) {
    var qClass = qi % 2 === 0 ? 'ab-q-odd' : 'ab-q-even';
    html += '<td class="ab-gt-cell ' + qClass + '">0</td>';
    html += '<td class="ab-gt-cell ' + qClass + '">0</td>';
    html += '<td class="ab-gt-cell ' + qClass + '">0</td>';
    html += '<td class="ab-gt-cell ' + qClass + '">0</td>';
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

  var html = '<div class="ab-nonprog-wrapper">';

  sectionKeys.forEach(function(secTitle) {
    var rows = sections[secTitle] || [];
    html += '<div class="ab-section">' +
      '<div class="ab-section-header" data-section="' + ab_he(secTitle) + '">' + ab_he(secTitle) + ' <span class="ab-toggle">▼</span></div>' +
      '<div class="ab-section-content"><div class="ab-scroll-wrapper">' +
      '<table class="ab-table ab-nonprog-table"><thead>';

    // Row 1: Headers with frozen columns (4) merged + quarterly + Combined + Remarks — all maroon
    html += '<tr class="ab-header-row-1">' +
      '<th colspan="4" class="ab-frozen-header" style="position:sticky;left:0;top:0;z-index:20;background:#8B1A1A;color:white;font-weight:700;">Particulars</th>' +
      '<th class="ab-quarter-header">&nbsp;</th>' +
      '<th class="ab-quarter-header">&nbsp;</th>';
    quarters.forEach(function(q, qi) {
      html += '<th colspan="2" class="ab-quarter-header">' + q.quarter + '</th>';
    });
    html += '<th colspan="2" class="ab-quarter-header">Combined</th><th class="ab-quarter-header">&nbsp;</th></tr>';

    // Row 2: Column headers
    html += '<tr class="ab-header-row-2">' +
      '<th class="ab-frozen" style="left:0;width:30px;min-width:30px;">Sr.</th>' +
      '<th class="ab-frozen" style="left:30px;width:160px;min-width:160px;">Particulars</th>' +
      '<th class="ab-frozen" style="left:190px;width:60px;min-width:60px;">UoM</th>' +
      '<th class="ab-frozen ab-frozen-last" style="left:250px;width:75px;min-width:75px;">Unit Cost</th>' +
      '<th class="ab-calc" style="width:70px;min-width:70px;">Total Units</th>' +
      '<th class="ab-calc" style="width:85px;min-width:85px;">Total Cost</th>';

    quarters.forEach(function(q, qi) {
      var qClass = qi % 2 === 0 ? 'ab-q-odd' : 'ab-q-even';
      html += '<th class="ab-subcol ' + qClass + '" style="width:65px;min-width:65px;">Units</th>';
      html += '<th class="ab-subcol ' + qClass + '" style="width:75px;min-width:75px;">Cost</th>';
    });
    html += '<th class="ab-subcol" style="width:65px;min-width:65px;">Units</th><th class="ab-subcol" style="width:75px;min-width:75px;">Cost</th><th class="ab-remarks-hdr" style="width:100px;min-width:100px;">Remarks</th>';
    html += '</tr></thead><tbody>';

    // Data rows
    rows.forEach(function(row, idx) {
      html += ab_buildNonProgRow(row, quarters, idx, secTitle, unitsList);
    });

    // Section total
    html += ab_buildNonProgSectionTotal(rows, quarters, secTitle);

    html += '</tbody></table></div>' +
      '<div style="padding:8px 16px;">' +
        '<button class="btn btn-xs btn-default ab-add-row-btn" data-section="' + ab_he(secTitle) + '" style="font-size:12px;padding:4px 12px;cursor:pointer;">' +
          '+ Add Row' +
        '</button>' +
      '</div>' +
      '</div></div>';
  });

  html += '</div>';
  return html;
}

function ab_buildNonProgRow(row, quarters, idx, secTitle, unitsList) {
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
    '<td class="ab-frozen ab-sr" style="left:0;width:30px;min-width:30px;">' + (idx + 1) + '</td>' +
    '<td class="ab-frozen ab-editable" style="left:30px;width:160px;min-width:160px;text-align:left;"><input type="text" class="ab-inp ab-desc-inp" style="width:150px;text-align:left;" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" value="' + ab_he(row.description) + '" placeholder="Enter particulars..." /></td>' +
    '<td class="ab-frozen ab-editable" style="left:190px;width:60px;min-width:60px;"><select class="ab-inp ab-uom-sel" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" style="width:55px;text-align:left;">' + uomOptions + '</select></td>' +
    '<td class="ab-frozen ab-editable ab-frozen-last" style="left:250px;width:75px;min-width:75px;"><input type="number" class="ab-inp ab-uc-inp" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" value="' + uc + '" /></td>';

  // Total Units (auto-calc) - non-frozen
  var totalUnits = 0;
  quarters.forEach(function(q, qi) {
    totalUnits += (qData[qi] || {}).units || 0;
  });
  html += '<td class="ab-calc" style="width:70px;min-width:70px;">' + totalUnits + '</td>';

  // Total Cost (auto-calc) - non-frozen
  var totalCost = uc * totalUnits;
  html += '<td class="ab-calc" style="width:85px;min-width:85px;">' + ab_fc(totalCost) + '</td>';

  // Per-quarter data
  var combinedCost = 0;
  quarters.forEach(function(q, qi) {
    var u = (qData[qi] || {}).units || 0;
    var c = u * uc;
    combinedCost += c;
    var qClass = qi % 2 === 0 ? 'ab-q-odd' : 'ab-q-even';

    html += '<td class="ab-editable ' + qClass + '" style="width:65px;"><input type="number" class="ab-inp ab-np-inp" style="width:55px;" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" data-qi="' + qi + '" value="' + u + '" /></td>' +
            '<td class="ab-calc ' + qClass + '" style="width:75px;">' + ab_fc(c) + '</td>';
  });

  // Combined Total
  html += '<td class="ab-calc">' + totalUnits + '</td>';
  html += '<td class="ab-calc">' + ab_fc(combinedCost) + '</td>';

  // Remarks
  html += '<td class="ab-editable"><input type="text" class="ab-inp ab-remarks-inp" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" value="' + ab_he(row.remarks || '') + '" placeholder="Remarks..." /></td>';

  html += '</tr>';
  return html;
}

function ab_buildNonProgSectionTotal(rows, quarters, secTitle) {
  var html = '<tr class="ab-section-total-row">' +
    '<td class="ab-frozen" style="left:0;width:30px;min-width:30px;"></td>' +
    '<td class="ab-frozen" style="left:30px;width:160px;min-width:160px;text-align:left;font-weight:700;">Section Total</td>' +
    '<td class="ab-frozen" style="left:190px;width:60px;min-width:60px;"></td>' +
    '<td class="ab-frozen" style="left:250px;width:75px;min-width:75px;"></td>';

  // Total Units - non-frozen
  var totalUnits = 0;
  rows.forEach(function(row) {
    quarters.forEach(function(q, qi) {
      totalUnits += (row.quarters[qi] || {}).units || 0;
    });
  });
  html += '<td class="ab-calc" style="width:70px;min-width:70px;">' + totalUnits + '</td>';

  // Total Cost - non-frozen
  var totalCost = 0;
  rows.forEach(function(row) {
    var uc = row.unit_cost || 0;
    quarters.forEach(function(q, qi) {
      var u = (row.quarters[qi] || {}).units || 0;
      totalCost += u * uc;
    });
  });
  html += '<td class="ab-calc" style="width:85px;min-width:85px;">' + ab_fc(totalCost) + '</td>';

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

  // Combined Total
  html += '<td class="ab-gt-cell">' + totalUnits + '</td>';
  html += '<td class="ab-gt-cell">' + ab_fc(totalCost) + '</td>';
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

  // Section toggle
  document.querySelectorAll('.ab-section-header').forEach(function(hdr) {
    hdr.addEventListener('click', function() {
      var content = this.nextElementSibling;
      var toggle = this.querySelector('.ab-toggle');
      if (content.style.display === 'none') {
        content.style.display = 'block';
        toggle.textContent = '▼';
      } else {
        content.style.display = 'none';
        toggle.textContent = '▶';
      }
    });
  });

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

    // Recalculate grand total
    ab_recalcProgGrandTotal(quarters, progData);
  }
}

function ab_recalcProgGrandTotal(quarters, progData) {
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
}

function ab_recalcNonProgRow(tr, quarters, years) {
  var cells = tr.querySelectorAll('td');
  var ucInput = cells[3] && cells[3].querySelector('input');
  var uc = ucInput ? (parseFloat(ucInput.value) || 0) : 0;

  // Recalc Total Units and Total Cost
  var totalUnits = 0;
  var ci = 6; // Start at per-quarter data
  quarters.forEach(function() {
    var u = parseInt(cells[ci].querySelector('input').value) || 0;
    totalUnits += u;
    ci += 2;
  });

  cells[4].textContent = totalUnits; // Total Units cell
  var totalCost = uc * totalUnits;
  cells[5].textContent = ab_fc(totalCost); // Total Cost cell

  // Recalc quarterly costs
  ci = 6;
  var combinedCost = 0;
  quarters.forEach(function() {
    var u = parseInt(cells[ci].querySelector('input').value) || 0;
    var c = u * uc;
    cells[ci + 1].textContent = ab_fc(c);
    combinedCost += c;
    ci += 2;
  });

  // Recalc Combined Total
  ci = 6 + (quarters.length * 2); // After all quarterly columns
  cells[ci].textContent = totalUnits;
  cells[ci + 1].textContent = ab_fc(combinedCost);

  ab_recalcNonProgSectionTotal(tr, quarters, years);
}

function ab_recalcNonProgSectionTotal(dataRowTr, quarters, years) {
  var tbody = dataRowTr.closest('tbody');
  if (!tbody) return;
  var totalRow = tbody.querySelector('.ab-section-total-row');
  if (!totalRow) return;
  var dataRows = tbody.querySelectorAll('.ab-data-row');
  var gtCells = totalRow.querySelectorAll('td');

  // Total Units and Total Cost across all rows
  var totalUnits = 0, totalCost = 0;
  dataRows.forEach(function(dr) {
    var drCells = dr.querySelectorAll('td');
    var ucInp = drCells[3] && drCells[3].querySelector('input');
    var uc = ucInp ? (parseFloat(ucInp.value) || 0) : 0;

    // Sum from each row's quarterly data
    var ci = 6;
    quarters.forEach(function() {
      var u = parseInt(drCells[ci].querySelector('input').value) || 0;
      totalUnits += u;
      totalCost += u * uc;
      ci += 2;
    });
  });

  gtCells[4].textContent = totalUnits; // Total Units in section total
  gtCells[5].textContent = ab_fc(totalCost); // Total Cost in section total

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

  // Combined Total
  gtCells[ci].textContent = totalUnits;
  ci++;
  gtCells[ci].textContent = ab_fc(totalCost);
}

// ============================================================================
// ADD ROW + SAVE (Non-Programmatic)
// ============================================================================

function ab_addNonProgRow(frm, secTitle, quarters, years, sbhRevMap, nonProgData, progData, unitsList) {
  var sectionEl = null;
  document.querySelectorAll('.ab-section-header').forEach(function(hdr) {
    if (hdr.dataset.section === secTitle) sectionEl = hdr.closest('.ab-section');
  });
  if (!sectionEl) return;

  var tbody = sectionEl.querySelector('tbody');
  if (!tbody) return;

  var totalRow = tbody.querySelector('.ab-section-total-row');
  var dataRows = tbody.querySelectorAll('.ab-data-row');
  var newIdx = dataRows.length;

  var newRow = {
    description: '',
    assumption: '',
    remarks: '',
    unit_cost: 0,
    pbpName: null,
    quarters: {}
  };
  for (var i = 0; i < quarters.length; i++) {
    newRow.quarters[i] = { units: 0 };
  }

  if (nonProgData.sections[secTitle]) {
    nonProgData.sections[secTitle].push(newRow);
  }

  var rowHtml = ab_buildNonProgRow(newRow, quarters, newIdx, secTitle, unitsList);
  var tempDiv = document.createElement('tbody');
  tempDiv.innerHTML = rowHtml;
  var newTr = tempDiv.querySelector('tr');

  if (totalRow) {
    tbody.insertBefore(newTr, totalRow);
  } else {
    tbody.appendChild(newTr);
  }

  tbody.querySelectorAll('.ab-data-row').forEach(function(tr, i) {
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

  var planningRows = [];
  var ci = 6; // After frozen cols + Total Units + Total Cost
  quarters.forEach(function(q) {
    var uInput = cells[ci] && cells[ci].querySelector('input');
    var units = uInput ? (parseInt(uInput.value) || 0) : 0;
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

    // Save or create PBP record
    var pbpName = row.pbpName;
    if (pbpName) {
      try {
        var r = await frappe.call({ method: 'frappe.client.get', args: { doctype: 'Project Budget Planning', name: pbpName } });
        if (r.message) {
          var doc = r.message;
          doc.description = row.description;
          doc.assumption = task;
          doc.total_planned_budget = totalBudget;
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
          total_planned_budget: totalBudget,
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
.ab-table { width: max-content; border-collapse: separate; border-spacing: 0; background: white; }
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
.ab-section { border: 1px solid #ddd; margin-bottom: 12px; border-radius: 4px; overflow: hidden; }
.ab-section-header { background: #8B1A1A; color: white; padding: 12px 16px; cursor: pointer; font-weight: 700; display: flex; justify-content: space-between; }
.ab-toggle { cursor: pointer; user-select: none; }
.ab-section-content { display: block; }
.ab-section-total-row { background: #F5E6E6; font-weight: 700; }
.ab-add-row-btn { margin-top: 8px; }
.ab-saved-indicator { color: #8B1A1A; font-weight: 700; display: none; }
.ab-nonprog-wrapper { }
.ab-table tbody tr:nth-child(even) td { background: #fafafa; }
.ab-table tbody tr:nth-child(even) .ab-frozen { background: #f5f5f5; }
.ab-table tbody tr:nth-child(even) .ab-calc { background: #f0f0f0; }
.ab-table tbody tr:nth-child(even) .ab-gt-cell { background: #F0DCDC; }
  `;
}

// ============================================================================
// EXCEL EXPORT
// ============================================================================

function ab_loadXlsxLib() {
  return new Promise(function(resolve, reject) {
    if (window.XLSX && window.XLSX.utils && window.XLSX.utils.aoa_to_sheet) { resolve(); return; }
    var s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';
    s.onload = function() { resolve(); };
    s.onerror = function() { reject(new Error('Failed to load XLSX library')); };
    document.head.appendChild(s);
  });
}

async function ab_downloadBudget(frm, quarters, years, progData, nonProgData, unitsList) {
  try {
    frappe.show_alert({ message: 'Preparing Excel download...', indicator: 'blue' });
    await ab_loadXlsxLib();

    var wb = XLSX.utils.book_new();

    // Color constants
    var C_DARK = '333333';
    var C_GREY_HDR = 'BCBDC0';
    var C_LIGHT_GREY = 'E6E6E6';
    var C_ALT_ROW = 'F4F5F6';
    var C_YELLOW = 'FFCB05';
    var C_BLUE = '00529C';
    var C_WHITE = 'FFFFFF';
    var INR_FMT = '_(₹* #,##0.00_);_(₹* (#,##0.00);_(₹* "-"??_);_(@_)';
    var NUM_FMT = '#,##0';

    function sc(v, opts) {
      opts = opts || {};
      var cell = { v: v, t: typeof v === 'number' ? 'n' : 's' };
      var s = {};
      if (opts.font) s.font = opts.font;
      else s.font = { name: 'Arial', sz: 11 };
      if (opts.fill) s.fill = { fgColor: { rgb: opts.fill } };
      if (opts.alignment) s.alignment = opts.alignment;
      else s.alignment = { vertical: 'center' };
      if (opts.border !== false) {
        s.border = {
          top: { style: 'thin', color: { rgb: '000000' } },
          bottom: { style: 'thin', color: { rgb: '000000' } },
          left: { style: 'thin', color: { rgb: '000000' } },
          right: { style: 'thin', color: { rgb: '000000' } }
        };
      }
      if (opts.numFmt) { s.numFmt = opts.numFmt; cell.z = opts.numFmt; }
      cell.s = s;
      return cell;
    }

    var progLiveData = ab_collectProgDataFromDOM(quarters, progData);
    var nonProgLiveData = ab_collectNonProgDataFromDOM(quarters, nonProgData);

    var summarySheet = ab_buildBudgetSummarySheet(frm, quarters, progLiveData, nonProgLiveData, sc, C_DARK, C_GREY_HDR, C_BLUE, C_YELLOW, C_WHITE, INR_FMT);
    XLSX.utils.book_append_sheet(wb, summarySheet.ws, 'Budget Summary');

    var progSheet = ab_buildProgSheet(frm, quarters, progLiveData, sc, C_DARK, C_GREY_HDR, C_LIGHT_GREY, C_YELLOW, C_BLUE, C_WHITE, INR_FMT, NUM_FMT);
    XLSX.utils.book_append_sheet(wb, progSheet.ws, 'Programmatic Costs');

    var npSheet = ab_buildNonProgSheet(frm, quarters, nonProgLiveData, sc, C_DARK, C_GREY_HDR, C_LIGHT_GREY, C_YELLOW, C_BLUE, C_WHITE, INR_FMT, NUM_FMT);
    XLSX.utils.book_append_sheet(wb, npSheet.ws, 'Non-Programmatic Costs');

    // Set sheet protection on Prog and NP sheets
    if (progSheet.ws) progSheet.ws['!protect'] = { password: '' };
    if (npSheet.ws) npSheet.ws['!protect'] = { password: '' };

    var filename = frm.doc.name + '_Budget_' + new Date().toISOString().split('T')[0] + '.xlsx';
    XLSX.writeFile(wb, filename);

    frappe.show_alert({ message: 'Downloaded: ' + filename, indicator: 'green' });
  } catch (err) {
    console.error('[AB] Download error:', err);
    frappe.show_alert({ message: 'Error: ' + (err.message || err), indicator: 'red' });
  }
}

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

    rows.push({ description: desc, task: task, uom: uom, unitCost: uc, totalUnits: tu, govt: govt, benf: benf, quarters: qtrData });
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
      var uomSel = tr.querySelector('.ab-uom-sel');
      var ucInp = tr.querySelector('.ab-uc-inp');
      var desc = descInp ? descInp.value.trim() : '';
      var uom = uomSel ? uomSel.value : '';
      var uc = ucInp ? parseFloat(ucInp.value) || 0 : 0;

      var qtrUnits = [];
      var qtrCosts = [];
      quarters.forEach(function(q, qi) {
        var inp = tr.querySelector('.ab-np-inp[data-qi="' + qi + '"]');
        var u = inp ? parseInt(inp.value) || 0 : 0;
        qtrUnits.push(u);
        qtrCosts.push(u * uc);
      });

      if (desc) sRows.push({ description: desc, uom: uom, unitCost: uc, quarterUnits: qtrUnits, quarterCosts: qtrCosts });
    });
    sections[sec] = sRows;
  });

  return { sections: sections };
}

function ab_buildProgSheet(frm, quarters, liveData, sc, C_DARK, C_GREY_HDR, C_LIGHT_GREY, C_YELLOW, C_BLUE, C_WHITE, INR_FMT, NUM_FMT) {
  var ws = {};
  var aoa = [];
  var merges = [];
  var rows = liveData.rows || [];

  // Column widths (27 columns: A-AA)
  var cols = [
    { wch: 8 },   // A: Sr. No.
    { wch: 25 },  // B: Activity
    { wch: 25 },  // C: Task Details
    { wch: 12 },  // D: Unit of Measurement
    { wch: 12 },  // E: Unit Cost
    { wch: 12 },  // F: Total Units
    { wch: 12 }   // G: Total Cost
  ];
  // H-J: Convergence (Total LIC HFL, Govt, Benf)
  cols.push({ wch: 12 }, { wch: 12 }, { wch: 12 });
  // Quarterly columns (K-Z): Q1, Q2, Q3, Q4 (4 cols each)
  quarters.forEach(function() {
    cols.push({ wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 });
  });
  // AA: Remarks
  cols.push({ wch: 12 });
  ws['!cols'] = cols;

  // Row 1-2: Title (merged D1:J2)
  aoa.push([sc('', { fill: C_LIGHT_GREY }), sc('', { fill: C_LIGHT_GREY }), sc('', { fill: C_LIGHT_GREY }), sc('Programmatic Costs', { font: { bold: true, sz: 13 }, fill: C_GREY_HDR, alignment: { horizontal: 'center', vertical: 'center' } }), sc(''), sc(''), sc('')]);
  merges.push({ s: { r: 0, c: 0 }, e: { r: 1, c: 2 } }); // A1:C2 merged empty
  merges.push({ s: { r: 0, c: 3 }, e: { r: 1, c: 9 } }); // D1:J2 merged title

  aoa.push(Array(27).fill(sc('')));

  // Row 3: Group headers
  var row3 = Array(27).fill(sc(''));
  row3[7] = sc('Convergence', { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
  merges.push({ s: { r: 2, c: 7 }, e: { r: 2, c: 9 } }); // H3:J3 merged

  // Quarter headers: K3 Q1, O3 Q2, S3 Q3, W3 Q4
  var qStartCols = [10, 14, 18, 22]; // K, O, S, W (0-indexed)
  quarters.forEach(function(q, qi) {
    var colIdx = qStartCols[qi];
    row3[colIdx] = sc('Q' + (qi + 1), { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
    merges.push({ s: { r: 2, c: colIdx }, e: { r: 2, c: colIdx + 3 } });
  });

  row3[26] = sc('Remarks', { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
  merges.push({ s: { r: 2, c: 26 }, e: { r: 3, c: 26 } }); // AA3:AA4 merged
  aoa.push(row3);

  // Row 4: Column headers (27 columns)
  var hdr4 = [
    sc('Sr. No.', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } } }),
    sc('Activity', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } } }),
    sc('Task Details (Mandatory)', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } } }),
    sc('Unit of Measurement', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } } }),
    sc('Unit Cost', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } }, numFmt: INR_FMT }),
    sc('Total Units', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } } }),
    sc('Total Cost', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } }, numFmt: INR_FMT }),
    sc('Total LIC HFL Contribution', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } }, numFmt: INR_FMT }),
    sc('Government Contribution', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } }, numFmt: INR_FMT }),
    sc('Beneficiary Contribution', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } }, numFmt: INR_FMT })
  ];

  // Quarterly headers
  quarters.forEach(function(q, qi) {
    var qFill = qi % 2 === 0 ? C_BLUE : C_YELLOW;
    var qColor = qi % 2 === 0 ? C_WHITE : C_DARK;
    hdr4.push(
      sc('Units', { fill: qFill, font: { bold: true, color: { rgb: qColor } } }),
      sc('Cost', { fill: qFill, font: { bold: true, color: { rgb: qColor } }, numFmt: INR_FMT }),
      sc('LIC HFL Contribution', { fill: qFill, font: { bold: true, color: { rgb: qColor } }, numFmt: INR_FMT }),
      sc('Beneficiary Contribution', { fill: qFill, font: { bold: true, color: { rgb: qColor } }, numFmt: INR_FMT })
    );
  });

  hdr4.push(sc('Remarks', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } } }));
  aoa.push(hdr4);

  // Data rows (Row 5+)
  rows.forEach(function(row, idx) {
    var totalCost = row.unitCost * row.totalUnits;
    var licHflContrib = totalCost - row.govt - row.benf;

    var r = [
      sc(idx + 1, { alignment: { horizontal: 'center' } }),
      sc(row.description),
      sc(row.task),
      sc(row.uom),
      sc(row.unitCost, { numFmt: INR_FMT }),
      sc(row.totalUnits),
      sc(totalCost, { numFmt: INR_FMT }),
      sc(licHflContrib, { numFmt: INR_FMT }),
      sc(row.govt, { numFmt: INR_FMT }),
      sc(row.benf, { numFmt: INR_FMT })
    ];

    // Quarterly data
    var qData = row.quarters || [];
    qData.forEach(function(q) {
      r.push(
        sc(q.units),
        sc(q.cost, { numFmt: INR_FMT }),
        sc(q.licHfl, { numFmt: INR_FMT }),
        sc(q.benf, { numFmt: INR_FMT })
      );
    });

    r.push(sc(''));
    aoa.push(r);
  });

  // Grand Total row
  var gtRow = Array(27).fill(sc(''));
  gtRow[0] = sc('GT', { font: { bold: true }, fill: C_YELLOW });
  gtRow[1] = sc('GRAND TOTAL', { font: { bold: true }, fill: C_YELLOW });

  var gtTotalCost = 0, gtLicHfl = 0, gtGovt = 0, gtBenf = 0;
  rows.forEach(function(row) {
    var tc = row.unitCost * row.totalUnits;
    gtTotalCost += tc;
    gtLicHfl += (tc - row.govt - row.benf);
    gtGovt += row.govt;
    gtBenf += row.benf;
  });

  gtRow[6] = sc(gtTotalCost, { numFmt: INR_FMT, fill: C_YELLOW, font: { bold: true } });
  gtRow[7] = sc(gtLicHfl, { numFmt: INR_FMT, fill: C_YELLOW, font: { bold: true } });
  gtRow[8] = sc(gtGovt, { numFmt: INR_FMT, fill: C_YELLOW, font: { bold: true } });
  gtRow[9] = sc(gtBenf, { numFmt: INR_FMT, fill: C_YELLOW, font: { bold: true } });

  // Quarterly totals
  quarters.forEach(function(q, qi) {
    var qUnits = 0, qCost = 0, qLicHfl = 0, qBenf = 0;
    rows.forEach(function(row) {
      var qd = row.quarters[qi] || {};
      qUnits += qd.units || 0;
      qCost += qd.cost || 0;
      qLicHfl += qd.licHfl || 0;
      qBenf += qd.benf || 0;
    });
    var colBase = 10 + (qi * 4);
    gtRow[colBase] = sc(qUnits, { fill: C_YELLOW, font: { bold: true } });
    gtRow[colBase + 1] = sc(qCost, { numFmt: INR_FMT, fill: C_YELLOW, font: { bold: true } });
    gtRow[colBase + 2] = sc(qLicHfl, { numFmt: INR_FMT, fill: C_YELLOW, font: { bold: true } });
    gtRow[colBase + 3] = sc(qBenf, { numFmt: INR_FMT, fill: C_YELLOW, font: { bold: true } });
  });

  gtRow[26] = sc('', { fill: C_YELLOW });
  aoa.push(gtRow);

  ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = merges;

  return { ws: ws };
}

function ab_buildNonProgSheet(frm, quarters, liveData, sc, C_DARK, C_GREY_HDR, C_LIGHT_GREY, C_YELLOW, C_BLUE, C_WHITE, INR_FMT, NUM_FMT) {
  var ws = {};
  var aoa = [];
  var merges = [];
  var rowIdx = 0;

  // Column widths (15 columns: A-O)
  var cols = [
    { wch: 8 },   // A: Sr. No.
    { wch: 20 },  // B: Particulars
    { wch: 15 },  // C: Unit of Measurement
    { wch: 12 },  // D: Unit Cost
    { wch: 12 },  // E: Total Units
    { wch: 12 }   // F: Total Cost
  ];
  // Quarterly columns (G-N): 4 quarters, 2 cols each
  quarters.forEach(function() {
    cols.push({ wch: 12 }, { wch: 12 });
  });
  // O: Remarks
  cols.push({ wch: 12 });
  ws['!cols'] = cols;

  var sectionOrder = ['Human Resource Costs', 'Administration Costs', 'NGO Management Costs'];
  var sections = liveData.sections || {};
  var allCosts = {};
  var allUnits = {};
  var totalCosts = {};

  // Calculate section totals (for summary in rows 1-4)
  sectionOrder.forEach(function(sec) {
    allCosts[sec] = 0;
    allUnits[sec] = 0;
    totalCosts[sec] = 0;
    var sRows = sections[sec] || [];
    sRows.forEach(function(row) {
      var uc = row.unitCost || 0;
      var totalUnits = (row.quarterUnits || []).reduce(function(s, u) { return s + u; }, 0);
      var rowTotal = uc * totalUnits;
      totalCosts[sec] += rowTotal;
    });
  });

  // Rows 1-4: Summary block (TOP LEFT)
  // Row 1
  aoa.push([
    sc(''), sc(''),
    sc('Cost Component', { fill: C_LIGHT_GREY, font: { bold: true }, alignment: { horizontal: 'center' } }),
    sc('', { fill: C_LIGHT_GREY }),
    sc('Amount', { fill: C_LIGHT_GREY, font: { bold: true }, alignment: { horizontal: 'center' } }),
    sc('%ge', { fill: C_LIGHT_GREY, font: { bold: true }, alignment: { horizontal: 'center' } })
  ]);
  merges.push({ s: { r: 0, c: 2 }, e: { r: 0, c: 3 } }); // C1:D1 merged "Cost Component"
  merges.push({ s: { r: 0, c: 6 }, e: { r: 3, c: 14 } }); // G1:O4 merged title block
  rowIdx++;

  var grandTotal = 0;
  sectionOrder.forEach(function(sec) { grandTotal += totalCosts[sec]; });

  // Rows 2-4: Section summary rows
  sectionOrder.forEach(function(sec) {
    var pct = grandTotal > 0 ? (totalCosts[sec] / grandTotal * 100).toFixed(2) : 0;
    aoa.push([
      sc(''), sc(''),
      sc(sec, { fill: C_LIGHT_GREY, alignment: { horizontal: 'center' } }),
      sc('', { fill: C_LIGHT_GREY }),
      sc(totalCosts[sec], { fill: C_LIGHT_GREY, numFmt: INR_FMT, alignment: { horizontal: 'right' } }),
      sc(pct + '%', { fill: C_LIGHT_GREY, alignment: { horizontal: 'center' } })
    ]);
    merges.push({ s: { r: rowIdx, c: 2 }, e: { r: rowIdx, c: 3 } });
    rowIdx++;
  });

  // Title block on right (G1:N4)
  aoa[0][6] = sc('Non-Programmatic Costs', { fill: C_GREY_HDR, font: { bold: true, sz: 12 }, alignment: { horizontal: 'center', vertical: 'center' } });

  // Row 5: Quarter headers (G5-N5)
  aoa.push(Array(15).fill(sc('')));
  var row5 = aoa[4];
  quarters.forEach(function(q, qi) {
    var colIdx = 6 + (qi * 2); // G=6, I=8, K=10, M=12
    row5[colIdx] = sc('Q' + (qi + 1), { fill: C_GREY_HDR, font: { bold: true }, alignment: { horizontal: 'center' } });
    merges.push({ s: { r: 4, c: colIdx }, e: { r: 4, c: colIdx + 1 } });
  });
  row5[14] = sc('Remarks', { fill: C_GREY_HDR, font: { bold: true }, alignment: { horizontal: 'center' } });
  merges.push({ s: { r: 4, c: 14 }, e: { r: 5, c: 14 } }); // O5:O6 merged
  rowIdx++;

  // Row 6: Column headers
  var hdr6 = [
    sc('Sr. No.', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } } }),
    sc('Particulars', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } } }),
    sc('Unit of Measurement', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } } }),
    sc('Unit Cost', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } }, numFmt: INR_FMT }),
    sc('Total Units', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } } }),
    sc('Total Cost', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } }, numFmt: INR_FMT })
  ];

  quarters.forEach(function(q, qi) {
    var qFill = qi % 2 === 0 ? C_BLUE : C_YELLOW;
    var qColor = qi % 2 === 0 ? C_WHITE : C_DARK;
    hdr6.push(
      sc('Units', { fill: qFill, font: { bold: true, color: { rgb: qColor } } }),
      sc('Cost', { fill: qFill, font: { bold: true, color: { rgb: qColor } }, numFmt: INR_FMT })
    );
  });

  hdr6.push(sc('Remarks', { fill: C_GREY_HDR, font: { bold: true, color: { rgb: C_DARK } } }));
  aoa.push(hdr6);
  rowIdx++;

  // Section data tables
  sectionOrder.forEach(function(sec) {
    // Section header row
    aoa.push([
      sc('A', { fill: C_DARK, font: { bold: true, color: { rgb: C_WHITE } } }),
      sc(sec, { fill: C_DARK, font: { bold: true, color: { rgb: C_WHITE } } })
    ]);
    // Fill rest with empty cells
    for (var i = 2; i < 15; i++) {
      aoa[rowIdx].push(sc('', { fill: C_DARK }));
    }
    rowIdx++;

    var sRows = sections[sec] || [];
    sRows.forEach(function(row, idx) {
      var totalUnits = (row.quarterUnits || []).reduce(function(s, u) { return s + u; }, 0);
      var totalCost = row.unitCost * totalUnits;

      var r = [
        sc(idx + 1, { alignment: { horizontal: 'center' } }),
        sc(row.description),
        sc(row.uom),
        sc(row.unitCost, { numFmt: INR_FMT }),
        sc(totalUnits),
        sc(totalCost, { numFmt: INR_FMT })
      ];

      // Quarterly data
      (row.quarterUnits || []).forEach(function(u, qi) {
        r.push(
          sc(u),
          sc(u * row.unitCost, { numFmt: INR_FMT })
        );
      });

      r.push(sc(''));
      aoa.push(r);
      rowIdx++;
    });

    // Section total row
    var stotal = [
      sc(''),
      sc('TOTAL(' + sec.charAt(0).toUpperCase() + ')', { font: { bold: true }, fill: C_LIGHT_GREY })
    ];
    for (var i = 2; i < 6; i++) stotal.push(sc('', { fill: C_LIGHT_GREY }));

    quarters.forEach(function(q, qi) {
      var qUnits = 0, qCost = 0;
      sRows.forEach(function(row) {
        qUnits += row.quarterUnits[qi] || 0;
        qCost += (row.quarterUnits[qi] || 0) * row.unitCost;
      });
      stotal.push(
        sc(qUnits, { fill: C_LIGHT_GREY, font: { bold: true } }),
        sc(qCost, { numFmt: INR_FMT, fill: C_LIGHT_GREY, font: { bold: true } })
      );
    });

    stotal.push(sc('', { fill: C_LIGHT_GREY }));
    aoa.push(stotal);
    rowIdx++;
  });

  // Grand Total row
  var gtRow = [
    sc(''),
    sc('Grand Total', { font: { bold: true }, fill: C_YELLOW })
  ];
  for (var i = 2; i < 6; i++) gtRow.push(sc('', { fill: C_YELLOW }));

  var gtQtrCosts = Array(quarters.length).fill(0);
  sectionOrder.forEach(function(sec) {
    var sRows = sections[sec] || [];
    sRows.forEach(function(row) {
      (row.quarterUnits || []).forEach(function(u, qi) {
        gtQtrCosts[qi] += u * row.unitCost;
      });
    });
  });

  quarters.forEach(function(q, qi) {
    gtRow.push(
      sc(0, { fill: C_YELLOW }),
      sc(gtQtrCosts[qi], { numFmt: INR_FMT, fill: C_YELLOW, font: { bold: true } })
    );
  });

  gtRow.push(sc('', { fill: C_YELLOW }));
  aoa.push(gtRow);

  ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = merges;

  return { ws: ws };
}

function ab_buildBudgetSummarySheet(frm, quarters, progData, nonProgData, sc, C_DARK, C_GREY_HDR, C_BLUE, C_YELLOW, C_WHITE, INR_FMT) {
  var ws = {};
  var aoa = [];
  var merges = [];

  // Column widths (13 columns: A-M)
  ws['!cols'] = [{ wch: 3 }, { wch: 3 }, { wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 3 }, { wch: 3 }, { wch: 3 }, { wch: 12 }, { wch: 15 }, { wch: 15 }, { wch: 12 }];

  // Row 1: Title (G1:M2)
  aoa.push(Array(13).fill(sc('')));
  merges.push({ s: { r: 0, c: 0 }, e: { r: 1, c: 5 } }); // A1:F2 merged empty
  aoa[0][6] = sc('Proposed Budget Statement - Summary Sheet', { font: { bold: true, sz: 12 }, fill: C_GREY_HDR, alignment: { horizontal: 'center', vertical: 'center' } });
  merges.push({ s: { r: 0, c: 6 }, e: { r: 1, c: 12 } }); // G1:M2 merged title

  aoa.push(Array(13).fill(sc('')));
  aoa.push(Array(13).fill(sc(''))); // Row 3: blank

  var rowIdx = 3;

  // Row 4: "Project Details" header (B4, merged B4:H4) and "Cost Per Beneficiary" (J4, merged J4:L4)
  var row4 = Array(13).fill(sc(''));
  row4[1] = sc('Project Details', { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
  row4[9] = sc('Cost Per Beneficiary', { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
  merges.push({ s: { r: 3, c: 1 }, e: { r: 3, c: 7 } }); // B4:H4
  merges.push({ s: { r: 3, c: 9 }, e: { r: 3, c: 11 } }); // J4:L4
  aoa.push(row4);
  rowIdx++;

  // Row 5: Name of Project
  var progTotal = (progData.rows || []).reduce(function(s, r) { return s + (r.unitCost * r.totalUnits); }, 0);
  var hrTotal = (nonProgData.sections['Human Resource Costs'] || []).reduce(function(s, r) {
    var tu = (r.quarterUnits || []).reduce(function(sum, u) { return sum + u; }, 0);
    return s + (r.unitCost * tu);
  }, 0);
  var adminTotal = (nonProgData.sections['Administration Costs'] || []).reduce(function(s, r) {
    var tu = (r.quarterUnits || []).reduce(function(sum, u) { return sum + u; }, 0);
    return s + (r.unitCost * tu);
  }, 0);
  var ngoTotal = (nonProgData.sections['NGO Management Costs'] || []).reduce(function(s, r) {
    var tu = (r.quarterUnits || []).reduce(function(sum, u) { return sum + u; }, 0);
    return s + (r.unitCost * tu);
  }, 0);
  var grandTotal = progTotal + hrTotal + adminTotal + ngoTotal;

  var row5 = Array(13).fill(sc(''));
  row5[1] = sc('I');
  row5[2] = sc('Name of the Project', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 4, c: 2 }, e: { r: 4, c: 4 } }); // C5:E5
  row5[5] = sc(frm.doc.project_name || '', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 4, c: 5 }, e: { r: 4, c: 7 } }); // F5:H5
  row5[9] = sc('Total Project Costs');
  row5[10] = sc(grandTotal, { numFmt: INR_FMT, alignment: { horizontal: 'right' } });
  merges.push({ s: { r: 4, c: 10 }, e: { r: 4, c: 11 } }); // K5:L5
  aoa.push(row5);
  rowIdx++;

  // Row 6: Project Duration
  var totalBenef = frm.doc.total_beneficiaries || 0;
  var row6 = Array(13).fill(sc(''));
  row6[1] = sc('II');
  row6[2] = sc('Project Duration', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 5, c: 2 }, e: { r: 5, c: 4 } }); // C6:E6
  row6[5] = sc(frm.doc.start_date || '', { alignment: { horizontal: 'left' } });
  row6[6] = sc('TO');
  row6[7] = sc(frm.doc.end_date || '', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 5, c: 5 }, e: { r: 5, c: 7 } }); // F6:H6
  row6[9] = sc('Total Beneficiaries');
  row6[10] = sc(totalBenef);
  row6[11] = sc('Individuals');
  aoa.push(row6);
  rowIdx++;

  // Row 7: Project Intervention States
  var row7 = Array(13).fill(sc(''));
  row7[1] = sc('III');
  row7[2] = sc('Project Intervention States', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 6, c: 2 }, e: { r: 6, c: 4 } }); // C7:E7
  row7[5] = sc(frm.doc.state || '', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 6, c: 5 }, e: { r: 6, c: 7 } }); // F7:H7
  row7[9] = sc('Cost per Beneficiary');
  row7[10] = sc(totalBenef > 0 ? grandTotal / totalBenef : 0, { numFmt: INR_FMT, alignment: { horizontal: 'right' } });
  merges.push({ s: { r: 6, c: 10 }, e: { r: 6, c: 11 } }); // K7:L7
  aoa.push(row7);
  rowIdx++;

  // Row 8: No. of Intervention Site
  var row8 = Array(13).fill(sc(''));
  row8[1] = sc('IV');
  row8[2] = sc('No. of Intervention Site', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 7, c: 2 }, e: { r: 7, c: 4 } }); // C8:E8
  row8[5] = sc('', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 7, c: 5 }, e: { r: 7, c: 7 } }); // F8:H8
  aoa.push(row8);
  rowIdx++;

  // Row 9: Implementing Partner
  var row9 = Array(13).fill(sc(''));
  row9[1] = sc('V');
  row9[2] = sc('Implementing Partner', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 8, c: 2 }, e: { r: 8, c: 4 } }); // C9:E9
  row9[5] = sc(frm.doc.ngo || '', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 8, c: 5 }, e: { r: 8, c: 7 } }); // F9:H9
  aoa.push(row9);
  rowIdx++;

  // Row 10: blank
  aoa.push(Array(13).fill(sc('')));
  rowIdx++;

  // Row 11: "Budget Breakup" header (B11, merged B11:G11)
  var row11 = Array(13).fill(sc(''));
  row11[1] = sc('Budget Breakup', { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
  merges.push({ s: { r: 10, c: 1 }, e: { r: 10, c: 6 } }); // B11:G11
  aoa.push(row11);
  rowIdx++;

  // Row 12: "A" + "DIRECT COSTS" header
  var row12 = Array(13).fill(sc(''));
  row12[1] = sc('A');
  row12[2] = sc('DIRECT COSTS', { font: { bold: true }, fill: C_GREY_HDR });
  merges.push({ s: { r: 11, c: 2 }, e: { r: 11, c: 4 } }); // C12:E12
  row12[5] = sc('Amount', { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
  row12[6] = sc('%ge', { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
  row12[9] = sc('Quarterly Breakup of Costs', { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
  merges.push({ s: { r: 11, c: 9 }, e: { r: 11, c: 11 } }); // J12:L12
  aoa.push(row12);
  rowIdx++;

  // Calculate quarterly costs for prog and NP
  var progQtrCosts = Array(quarters.length).fill(0);
  var npQtrCosts = Array(quarters.length).fill(0);

  (progData.rows || []).forEach(function(row) {
    (row.quarters || []).forEach(function(q, qi) {
      progQtrCosts[qi] += q.cost || 0;
    });
  });

  Object.keys(nonProgData.sections || {}).forEach(function(sec) {
    (nonProgData.sections[sec] || []).forEach(function(row) {
      (row.quarterUnits || []).forEach(function(u, qi) {
        npQtrCosts[qi] += u * row.unitCost;
      });
    });
  });

  // Row 13: Programmatic Costs
  var q1ProgNpCost = progQtrCosts[0] + npQtrCosts[0];
  var row13 = Array(13).fill(sc(''));
  row13[1] = sc('i');
  row13[2] = sc('Programmatic Costs', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 12, c: 2 }, e: { r: 12, c: 4 } }); // C13:E13
  row13[5] = sc(progTotal, { numFmt: INR_FMT, alignment: { horizontal: 'right' } });
  row13[6] = sc(grandTotal > 0 ? ((progTotal / grandTotal) * 100).toFixed(2) + '%' : '0%', { alignment: { horizontal: 'center' } });
  row13[9] = sc('First Quarter');
  row13[10] = sc(q1ProgNpCost, { numFmt: INR_FMT, alignment: { horizontal: 'right' } });
  row13[11] = sc(grandTotal > 0 ? ((q1ProgNpCost / grandTotal) * 100).toFixed(2) + '%' : '0%', { alignment: { horizontal: 'center' } });
  aoa.push(row13);
  rowIdx++;

  // Row 14: Human Resource Costs
  var q2ProgNpCost = progQtrCosts[1] + npQtrCosts[1];
  var row14 = Array(13).fill(sc(''));
  row14[1] = sc('ii');
  row14[2] = sc('Human Resource Costs', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 13, c: 2 }, e: { r: 13, c: 4 } }); // C14:E14
  row14[5] = sc(hrTotal, { numFmt: INR_FMT, alignment: { horizontal: 'right' } });
  row14[6] = sc(grandTotal > 0 ? ((hrTotal / grandTotal) * 100).toFixed(2) + '%' : '0%', { alignment: { horizontal: 'center' } });
  row14[9] = sc('Second Quarter');
  row14[10] = sc(q2ProgNpCost, { numFmt: INR_FMT, alignment: { horizontal: 'right' } });
  row14[11] = sc(grandTotal > 0 ? ((q2ProgNpCost / grandTotal) * 100).toFixed(2) + '%' : '0%', { alignment: { horizontal: 'center' } });
  aoa.push(row14);
  rowIdx++;

  // Row 15: TOTAL DIRECT COST (A)
  var totalDirect = progTotal + hrTotal;
  var q3ProgNpCost = progQtrCosts[2] + npQtrCosts[2];
  var row15 = Array(13).fill(sc(''));
  row15[2] = sc('TOTAL DIRECT COST (A)', { font: { bold: true }, alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 14, c: 2 }, e: { r: 14, c: 4 } }); // C15:E15
  row15[5] = sc(totalDirect, { numFmt: INR_FMT, font: { bold: true }, alignment: { horizontal: 'right' } });
  row15[6] = sc(grandTotal > 0 ? ((totalDirect / grandTotal) * 100).toFixed(2) + '%' : '0%', { font: { bold: true }, alignment: { horizontal: 'center' } });
  row15[9] = sc('Third Quarter');
  row15[10] = sc(q3ProgNpCost, { numFmt: INR_FMT, alignment: { horizontal: 'right' } });
  row15[11] = sc(grandTotal > 0 ? ((q3ProgNpCost / grandTotal) * 100).toFixed(2) + '%' : '0%', { alignment: { horizontal: 'center' } });
  aoa.push(row15);
  rowIdx++;

  // Row 16: "B" + "INDIRECT COSTS" header
  var row16 = Array(13).fill(sc(''));
  row16[1] = sc('B');
  row16[2] = sc('INDIRECT COSTS', { font: { bold: true }, fill: C_GREY_HDR });
  merges.push({ s: { r: 15, c: 2 }, e: { r: 15, c: 4 } }); // C16:E16
  row16[5] = sc('Amount', { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
  row16[6] = sc('%ge', { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
  row16[9] = sc('Fourth Quarter');
  var q4ProgNpCost = progQtrCosts[3] + npQtrCosts[3];
  row16[10] = sc(q4ProgNpCost, { numFmt: INR_FMT, alignment: { horizontal: 'right' } });
  row16[11] = sc(grandTotal > 0 ? ((q4ProgNpCost / grandTotal) * 100).toFixed(2) + '%' : '0%', { alignment: { horizontal: 'center' } });
  aoa.push(row16);
  rowIdx++;

  // Row 17: Admin Costs + Total (merged J17:J18, K17:K18, L17:L18)
  var row17 = Array(13).fill(sc(''));
  row17[1] = sc('i');
  row17[2] = sc('Admin Costs', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 16, c: 2 }, e: { r: 16, c: 4 } }); // C17:E17
  row17[5] = sc(adminTotal, { numFmt: INR_FMT, alignment: { horizontal: 'right' } });
  row17[6] = sc(grandTotal > 0 ? ((adminTotal / grandTotal) * 100).toFixed(2) + '%' : '0%', { alignment: { horizontal: 'center' } });
  row17[9] = sc('Total', { font: { bold: true } });
  row17[10] = sc(grandTotal, { numFmt: INR_FMT, font: { bold: true }, alignment: { horizontal: 'right' } });
  row17[11] = sc('100%', { font: { bold: true }, alignment: { horizontal: 'center' } });
  merges.push({ s: { r: 16, c: 9 }, e: { r: 17, c: 9 } }); // J17:J18
  merges.push({ s: { r: 16, c: 10 }, e: { r: 17, c: 10 } }); // K17:K18
  merges.push({ s: { r: 16, c: 11 }, e: { r: 17, c: 11 } }); // L17:L18
  aoa.push(row17);
  rowIdx++;

  // Row 18: NGO Management Costs
  var row18 = Array(13).fill(sc(''));
  row18[1] = sc('ii');
  row18[2] = sc('NGO Management Costs', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 17, c: 2 }, e: { r: 17, c: 4 } }); // C18:E18
  row18[5] = sc(ngoTotal, { numFmt: INR_FMT, alignment: { horizontal: 'right' } });
  row18[6] = sc(grandTotal > 0 ? ((ngoTotal / grandTotal) * 100).toFixed(2) + '%' : '0%', { alignment: { horizontal: 'center' } });
  aoa.push(row18);
  rowIdx++;

  // Row 19: TOTAL INDIRECT COSTS (B)
  var totalIndirect = adminTotal + ngoTotal;
  var row19 = Array(13).fill(sc(''));
  row19[2] = sc('TOTAL INDIRECT COSTS (B)', { font: { bold: true }, alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 18, c: 2 }, e: { r: 18, c: 4 } }); // C19:E19
  row19[5] = sc(totalIndirect, { numFmt: INR_FMT, font: { bold: true }, alignment: { horizontal: 'right' } });
  row19[6] = sc(grandTotal > 0 ? ((totalIndirect / grandTotal) * 100).toFixed(2) + '%' : '0%', { font: { bold: true }, alignment: { horizontal: 'center' } });
  aoa.push(row19);
  rowIdx++;

  // Row 20: GRANT TOTAL (A+B)
  var row20 = Array(13).fill(sc(''));
  row20[2] = sc('GRANT TOTAL (A+B)', { font: { bold: true }, alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 19, c: 2 }, e: { r: 19, c: 4 } }); // C20:E20
  row20[5] = sc(grandTotal, { numFmt: INR_FMT, font: { bold: true }, fill: C_YELLOW, alignment: { horizontal: 'right' } });
  row20[6] = sc('100%', { font: { bold: true }, fill: C_YELLOW, alignment: { horizontal: 'center' } });
  aoa.push(row20);
  rowIdx++;

  // Row 21: blank
  aoa.push(Array(13).fill(sc('')));
  rowIdx++;

  // Row 22: "Cost Sharing Details" header (B22, merged B22:G22)
  var row22 = Array(13).fill(sc(''));
  row22[1] = sc('Cost Sharing Details', { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
  merges.push({ s: { r: 21, c: 1 }, e: { r: 21, c: 6 } }); // B22:G22
  aoa.push(row22);
  rowIdx++;

  // Row 23: Column headers for cost sharing
  var row23 = Array(13).fill(sc(''));
  row23[1] = sc('Cost Sharing', { font: { bold: true }, fill: C_GREY_HDR });
  merges.push({ s: { r: 22, c: 1 }, e: { r: 22, c: 4 } }); // B23:E23
  row23[5] = sc('Amount', { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
  row23[6] = sc('%ge', { font: { bold: true }, fill: C_GREY_HDR, alignment: { horizontal: 'center' } });
  aoa.push(row23);
  rowIdx++;

  // Row 24: Government Convergence
  var govtContrib = (progData.rows || []).reduce(function(s, r) { return s + (r.govt || 0); }, 0);
  var row24 = Array(13).fill(sc(''));
  row24[1] = sc('Government Convergence', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 23, c: 1 }, e: { r: 23, c: 4 } }); // B24:E24
  row24[5] = sc(govtContrib, { numFmt: INR_FMT, alignment: { horizontal: 'right' } });
  row24[6] = sc(grandTotal > 0 ? ((govtContrib / grandTotal) * 100).toFixed(2) + '%' : '0%', { alignment: { horizontal: 'center' } });
  aoa.push(row24);
  rowIdx++;

  // Row 25: Community/Beneficiary
  var benfContrib = (progData.rows || []).reduce(function(s, r) { return s + (r.benf || 0); }, 0);
  var row25 = Array(13).fill(sc(''));
  row25[1] = sc('Community/Beneficiary/NGO/Other', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 24, c: 1 }, e: { r: 24, c: 4 } }); // B25:E25
  row25[5] = sc(benfContrib, { numFmt: INR_FMT, alignment: { horizontal: 'right' } });
  row25[6] = sc(grandTotal > 0 ? ((benfContrib / grandTotal) * 100).toFixed(2) + '%' : '0%', { alignment: { horizontal: 'center' } });
  aoa.push(row25);
  rowIdx++;

  // Row 26: LIC HFL Contribution
  var licHflContrib = grandTotal - govtContrib - benfContrib;
  var row26 = Array(13).fill(sc(''));
  row26[1] = sc('LIC HFL Contribution', { alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 25, c: 1 }, e: { r: 25, c: 4 } }); // B26:E26
  row26[5] = sc(licHflContrib, { numFmt: INR_FMT, alignment: { horizontal: 'right' } });
  row26[6] = sc(grandTotal > 0 ? ((licHflContrib / grandTotal) * 100).toFixed(2) + '%' : '0%', { alignment: { horizontal: 'center' } });
  aoa.push(row26);
  rowIdx++;

  // Row 27: TOTAL
  var row27 = Array(13).fill(sc(''));
  row27[1] = sc('TOTAL', { font: { bold: true }, alignment: { horizontal: 'left' } });
  merges.push({ s: { r: 26, c: 1 }, e: { r: 26, c: 4 } }); // B27:E27
  row27[5] = sc(grandTotal, { numFmt: INR_FMT, font: { bold: true }, fill: C_YELLOW, alignment: { horizontal: 'right' } });
  row27[6] = sc('100%', { font: { bold: true }, fill: C_YELLOW, alignment: { horizontal: 'center' } });
  aoa.push(row27);

  ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = merges;

  return { ws: ws };
}
