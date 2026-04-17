// ============================================================================
// LIC Budget Allocation v7f — No auto-save + compact UI + scroll fix + hide Budget Summary
// Client Script for Project proposal
// ============================================================================

frappe.ui.form.on('Project proposal', {
  refresh(frm) {
    if (!frm.doc.__islocal) setup_budget_tab(frm);
  },
  onload(frm) {
    if (!frm.doc.__islocal) setup_budget_tab(frm);
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
    var fsKey = ab_normFS(fsMap[rec.fund_source] || '');

    if (!activities[desc]) {
      activities[desc] = {
        description: desc,
        assumption: rec.assumption || '',
        unit_cost: 0,
        fundSources: { lic: { pbpName: null, quarters: {} }, govt: { pbpName: null, quarters: {} }, benf: { pbpName: null, quarters: {} } }
      };
      activityOrder.push(desc);

      // Initialize all quarters to 0
      for (var i = 0; i < quarters.length; i++) {
        activities[desc].fundSources.lic.quarters[i] = { units: 0 };
        activities[desc].fundSources.govt.quarters[i] = { units: 0 };
        activities[desc].fundSources.benf.quarters[i] = { units: 0 };
      }
    }

    var act = activities[desc];
    act.fundSources[fsKey].pbpName = pbpName;

    // Populate quarters from planning_table
    (rec.planning_table || []).forEach(function(row) {
      var qi = qIndex[ab_qKey(row.year, row.quarter)];
      if (qi !== undefined) {
        act.fundSources[fsKey].quarters[qi] = { units: row.unit || 0 };
        if (row.unit_cost && !act.unit_cost) act.unit_cost = row.unit_cost;
      }
    });
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
          assumption: kpi.uomName || '',
          unit_cost: 0,
          activityMasterId: kpi.actId,
          autoPopulated: true,
          fundSources: { lic: { pbpName: null, quarters: {} }, govt: { pbpName: null, quarters: {} }, benf: { pbpName: null, quarters: {} } }
        };
        activityOrder.push(desc);

        // Initialize all quarters to 0
        for (var i = 0; i < quarters.length; i++) {
          activities[desc].fundSources.lic.quarters[i] = { units: 0 };
          activities[desc].fundSources.govt.quarters[i] = { units: 0 };
          activities[desc].fundSources.benf.quarters[i] = { units: 0 };
        }
      } else {
        // Activity exists from PBP — enrich with Activity Master ID
        activities[desc].activityMasterId = kpi.actId;
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
      '<span class="ab-legend"><span style="background:#fffef5;padding:2px 6px;border:1px solid #ddd;">■</span> Editable</span>' +
      '<span class="ab-legend"><span style="background:#f4f7fa;padding:2px 6px;border:1px solid #ddd;">■</span> Auto-calculated</span>' +
      '<span class="ab-legend"><span style="background:#fff9e6;padding:2px 6px;border:1px solid #ddd;">■</span> Grand Total</span>' +
      '<span class="ab-saved-indicator" id="ab-saved"></span>' +
    '</div>' +
  '</div>';
}

// ---- Programmatic Tab ----

function ab_buildProgTab(frm, quarters, years, data) {
  var rows = data.rows || [];
  var nq = 7; // 7 subcols per quarter
  var ny = 5; // 5 year-total subcols

  var html = '<div class="ab-scroll-wrapper"><table class="ab-table ab-prog-table"><thead>';

  // Row 1: Year spans + Grand Total
  html += '<tr class="ab-header-row-1"><th colspan="5" class="ab-frozen-header" style="left:0;">.</th>';
  years.forEach(function(y, yi) {
    var yqCount = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; }).length;
    var colspan = yqCount * nq + ny;
    html += '<th colspan="' + colspan + '" class="ab-year-header ab-year-' + yi + '">Year ' + (yi + 1) + '</th>';
  });
  html += '<th colspan="3" class="ab-grand-total-header">Grand Total</th></tr>';

  // Row 2: Quarter labels + Year Total label
  html += '<tr class="ab-header-row-2"><th colspan="5" class="ab-frozen-header" style="left:0;">.</th>';
  years.forEach(function(y, yi) {
    quarters.filter(function(q) { return q.year_sequence === y.year_sequence; }).forEach(function(q) {
      html += '<th colspan="7" class="ab-quarter-header ab-year-' + yi + '">' + q.quarter + ' ' + ab_dateRange(q.start_date, q.end_date) + '</th>';
    });
    html += '<th colspan="5" class="ab-year-total-header ab-year-' + yi + '">Year ' + (yi + 1) + ' Total</th>';
  });
  html += '<th colspan="3" class="ab-grand-total-header">.</th></tr>';

  // Row 3: Sub-column headers
  html += '<tr class="ab-header-row-3">' +
    '<th class="ab-frozen" style="left:0;width:40px;min-width:40px;">Sr.</th>' +
    '<th class="ab-frozen" style="left:40px;width:200px;min-width:200px;">Activity</th>' +
    '<th class="ab-frozen" style="left:240px;width:180px;min-width:180px;">Task Details</th>' +
    '<th class="ab-frozen" style="left:420px;width:80px;min-width:80px;">UoM</th>' +
    '<th class="ab-frozen" style="left:500px;width:90px;min-width:90px;">Unit Cost</th>';

  years.forEach(function(y, yi) {
    quarters.filter(function(q) { return q.year_sequence === y.year_sequence; }).forEach(function() {
      html += '<th class="ab-subcol ab-year-' + yi + '">LIC Units</th>' +
              '<th class="ab-subcol ab-year-' + yi + '">LIC Amt</th>' +
              '<th class="ab-subcol ab-year-' + yi + '">Govt Units</th>' +
              '<th class="ab-subcol ab-year-' + yi + '">Govt Amt</th>' +
              '<th class="ab-subcol ab-year-' + yi + '">Benf Units</th>' +
              '<th class="ab-subcol ab-year-' + yi + '">Benf Amt</th>' +
              '<th class="ab-subcol ab-year-' + yi + '">Total Amt</th>';
    });
    html += '<th class="ab-yt-subcol ab-year-' + yi + '">Total Units</th>' +
            '<th class="ab-yt-subcol ab-year-' + yi + '">Total Amt</th>' +
            '<th class="ab-yt-subcol ab-year-' + yi + '">LIC Cont</th>' +
            '<th class="ab-yt-subcol ab-year-' + yi + '">Govt Cont</th>' +
            '<th class="ab-yt-subcol ab-year-' + yi + '">Benf Cont</th>';
  });
  html += '<th class="ab-gt-subcol">Y1 Total</th><th class="ab-gt-subcol">Y2 Total</th><th class="ab-gt-subcol">Combined</th>';
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
  var uc = row.unit_cost || 0;
  var lic = row.fundSources.lic;
  var govt = row.fundSources.govt;
  var benf = row.fundSources.benf;

  var html = '<tr class="ab-data-row" data-idx="' + idx + '">' +
    '<td class="ab-frozen ab-sr" style="left:0;">' + (idx + 1) + '</td>' +
    '<td class="ab-frozen" style="left:40px;text-align:left;"><div class="ab-desc-cell" title="' + ab_he(row.description) + '">' + ab_he(row.description) + '</div></td>' +
    '<td class="ab-frozen" style="left:240px;text-align:left;">' + ab_he(row.assumption || '') + '</td>' +
    '<td class="ab-frozen" style="left:420px;">Numbers</td>' +
    '<td class="ab-frozen ab-editable" style="left:500px;"><input type="number" class="ab-inp ab-uc-inp" data-idx="' + idx + '" value="' + uc + '" /></td>';

  // Track year totals for grand total cols
  var yearTotals = [];

  years.forEach(function(y, yi) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yTotalUnits = 0, yLic = 0, yGovt = 0, yBenf = 0;

    yqList.forEach(function(q) {
      // Find the index of this quarter in the full quarters array
      var qi = quarters.indexOf(q);
      var lu = (lic.quarters[qi] || {}).units || 0;
      var gu = (govt.quarters[qi] || {}).units || 0;
      var bu = (benf.quarters[qi] || {}).units || 0;
      var la = lu * uc, ga = gu * uc, ba = bu * uc;
      var ta = la + ga + ba;

      yTotalUnits += lu + gu + bu;
      yLic += la; yGovt += ga; yBenf += ba;

      html += '<td class="ab-editable"><input type="number" class="ab-inp" data-idx="' + idx + '" data-fs="lic" data-qi="' + qi + '" value="' + lu + '" /></td>' +
              '<td class="ab-calc">' + ab_fc(la) + '</td>' +
              '<td class="ab-editable"><input type="number" class="ab-inp" data-idx="' + idx + '" data-fs="govt" data-qi="' + qi + '" value="' + gu + '" /></td>' +
              '<td class="ab-calc">' + ab_fc(ga) + '</td>' +
              '<td class="ab-editable"><input type="number" class="ab-inp" data-idx="' + idx + '" data-fs="benf" data-qi="' + qi + '" value="' + bu + '" /></td>' +
              '<td class="ab-calc">' + ab_fc(ba) + '</td>' +
              '<td class="ab-calc ab-total-col">' + ab_fc(ta) + '</td>';
    });

    var yTotal = yLic + yGovt + yBenf;
    yearTotals.push(yTotal);

    html += '<td class="ab-yt-cell">' + yTotalUnits + '</td>' +
            '<td class="ab-yt-cell">' + ab_fc(yTotal) + '</td>' +
            '<td class="ab-yt-cell">' + ab_fc(yLic) + '</td>' +
            '<td class="ab-yt-cell">' + ab_fc(yGovt) + '</td>' +
            '<td class="ab-yt-cell">' + ab_fc(yBenf) + '</td>';
  });

  // Grand total: Y1, Y2, Combined
  var y1 = yearTotals[0] || 0;
  var y2 = yearTotals[1] || 0;
  html += '<td class="ab-gt-cell">' + ab_fc(y1) + '</td>' +
          '<td class="ab-gt-cell">' + ab_fc(y2) + '</td>' +
          '<td class="ab-gt-cell">' + ab_fc(y1 + y2) + '</td>';

  html += '</tr>';
  return html;
}

function ab_buildProgGrandTotal(rows, quarters, years) {
  // Sum across all rows
  var html = '<tr class="ab-grand-total-row">' +
    '<td class="ab-frozen ab-sr" style="left:0;">GT</td>' +
    '<td class="ab-frozen" style="left:40px;text-align:left;font-weight:700;">GRAND TOTAL</td>' +
    '<td class="ab-frozen" style="left:240px;"></td>' +
    '<td class="ab-frozen" style="left:420px;"></td>' +
    '<td class="ab-frozen" style="left:500px;"></td>';

  var grandY = [];

  years.forEach(function(y, yi) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yTotalUnits = 0, yLic = 0, yGovt = 0, yBenf = 0;

    yqList.forEach(function(q) {
      var qi = quarters.indexOf(q);
      var qLu = 0, qLa = 0, qGu = 0, qGa = 0, qBu = 0, qBa = 0;

      rows.forEach(function(row) {
        var uc = row.unit_cost || 0;
        var lu = (row.fundSources.lic.quarters[qi] || {}).units || 0;
        var gu = (row.fundSources.govt.quarters[qi] || {}).units || 0;
        var bu = (row.fundSources.benf.quarters[qi] || {}).units || 0;
        qLu += lu; qLa += lu * uc;
        qGu += gu; qGa += gu * uc;
        qBu += bu; qBa += bu * uc;
      });

      var qTa = qLa + qGa + qBa;
      yTotalUnits += qLu + qGu + qBu;
      yLic += qLa; yGovt += qGa; yBenf += qBa;

      html += '<td class="ab-gt-cell">' + qLu + '</td>' +
              '<td class="ab-gt-cell">' + ab_fc(qLa) + '</td>' +
              '<td class="ab-gt-cell">' + qGu + '</td>' +
              '<td class="ab-gt-cell">' + ab_fc(qGa) + '</td>' +
              '<td class="ab-gt-cell">' + qBu + '</td>' +
              '<td class="ab-gt-cell">' + ab_fc(qBa) + '</td>' +
              '<td class="ab-gt-cell">' + ab_fc(qTa) + '</td>';
    });

    var yTotal = yLic + yGovt + yBenf;
    grandY.push(yTotal);

    html += '<td class="ab-gt-cell">' + yTotalUnits + '</td>' +
            '<td class="ab-gt-cell">' + ab_fc(yTotal) + '</td>' +
            '<td class="ab-gt-cell">' + ab_fc(yLic) + '</td>' +
            '<td class="ab-gt-cell">' + ab_fc(yGovt) + '</td>' +
            '<td class="ab-gt-cell">' + ab_fc(yBenf) + '</td>';
  });

  var gy1 = grandY[0] || 0;
  var gy2 = grandY[1] || 0;
  html += '<td class="ab-gt-cell ab-gt-final">' + ab_fc(gy1) + '</td>' +
          '<td class="ab-gt-cell ab-gt-final">' + ab_fc(gy2) + '</td>' +
          '<td class="ab-gt-cell ab-gt-final">' + ab_fc(gy1 + gy2) + '</td>';

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

    // Row 1: Year spans
    html += '<tr class="ab-header-row-1"><th colspan="4" class="ab-frozen-header" style="left:0;">.</th>';
    years.forEach(function(y, yi) {
      var yqCount = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; }).length;
      html += '<th colspan="' + (yqCount * 2) + '" class="ab-year-header ab-year-' + yi + '">Year ' + (yi + 1) + '</th>';
    });
    html += '<th colspan="' + (years.length * 2) + '" class="ab-year-header">Year Totals</th>';
    html += '<th colspan="3" class="ab-grand-total-header">Grand Total</th></tr>';

    // Row 2: Quarter labels
    html += '<tr class="ab-header-row-2"><th colspan="4" class="ab-frozen-header" style="left:0;">.</th>';
    years.forEach(function(y, yi) {
      quarters.filter(function(q) { return q.year_sequence === y.year_sequence; }).forEach(function(q) {
        html += '<th colspan="2" class="ab-quarter-header ab-year-' + yi + '">' + q.quarter + '</th>';
      });
    });
    years.forEach(function(y, yi) {
      html += '<th colspan="2" class="ab-year-total-header ab-year-' + yi + '">Year ' + (yi + 1) + '</th>';
    });
    html += '<th colspan="3" class="ab-grand-total-header">.</th></tr>';

    // Row 3: Sub-columns
    html += '<tr class="ab-header-row-3">' +
      '<th class="ab-frozen" style="left:0;width:30px;min-width:30px;">Sr.</th>' +
      '<th class="ab-frozen" style="left:30px;width:180px;min-width:180px;">Particulars</th>' +
      '<th class="ab-frozen" style="left:210px;width:72px;min-width:72px;">UoM</th>' +
      '<th class="ab-frozen" style="left:282px;width:70px;min-width:70px;">Unit Cost</th>';

    years.forEach(function(y, yi) {
      quarters.filter(function(q) { return q.year_sequence === y.year_sequence; }).forEach(function() {
        html += '<th class="ab-subcol ab-year-' + yi + '">Units</th><th class="ab-subcol ab-year-' + yi + '">Cost</th>';
      });
    });
    years.forEach(function(y, yi) {
      html += '<th class="ab-yt-subcol ab-year-' + yi + '">Units</th><th class="ab-yt-subcol ab-year-' + yi + '">Cost</th>';
    });
    html += '<th class="ab-gt-subcol">Y1 Cost</th><th class="ab-gt-subcol">Y2 Cost</th><th class="ab-gt-subcol">Combined</th>';
    html += '</tr></thead><tbody>';

    // Data rows
    rows.forEach(function(row, idx) {
      html += ab_buildNonProgRow(row, quarters, years, idx, secTitle, unitsList);
    });

    // Section total
    html += ab_buildNonProgSectionTotal(rows, quarters, years, secTitle);

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
    '<td class="ab-frozen ab-sr" style="left:0;">' + (idx + 1) + '</td>' +
    '<td class="ab-frozen ab-editable" style="left:30px;text-align:left;"><input type="text" class="ab-inp ab-desc-inp" style="width:170px;text-align:left;" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" value="' + ab_he(row.description) + '" placeholder="Enter particulars..." /></td>' +
    '<td class="ab-frozen ab-editable" style="left:210px;"><select class="ab-inp ab-uom-sel" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" style="width:75px;text-align:left;">' + uomOptions + '</select></td>' +
    '<td class="ab-frozen ab-editable" style="left:282px;"><input type="number" class="ab-inp ab-uc-inp" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" value="' + uc + '" /></td>';

  var yearTotals = [];

  years.forEach(function(y, yi) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yUnits = 0, yCost = 0;

    yqList.forEach(function(q) {
      var qi = quarters.indexOf(q);
      var u = (qData[qi] || {}).units || 0;
      var c = u * uc;
      yUnits += u;
      yCost += c;

      html += '<td class="ab-editable"><input type="number" class="ab-inp ab-np-inp" data-section="' + ab_he(secTitle) + '" data-ridx="' + idx + '" data-qi="' + qi + '" value="' + u + '" /></td>' +
              '<td class="ab-calc">' + ab_fc(c) + '</td>';
    });

    yearTotals.push({ units: yUnits, cost: yCost });
  });

  // Year totals
  years.forEach(function(y, yi) {
    html += '<td class="ab-yt-cell">' + yearTotals[yi].units + '</td>' +
            '<td class="ab-yt-cell">' + ab_fc(yearTotals[yi].cost) + '</td>';
  });

  // Grand total
  var y1c = (yearTotals[0] || {}).cost || 0;
  var y2c = (yearTotals[1] || {}).cost || 0;
  html += '<td class="ab-gt-cell">' + ab_fc(y1c) + '</td>' +
          '<td class="ab-gt-cell">' + ab_fc(y2c) + '</td>' +
          '<td class="ab-gt-cell">' + ab_fc(y1c + y2c) + '</td>';

  html += '</tr>';
  return html;
}

function ab_buildNonProgSectionTotal(rows, quarters, years, secTitle) {
  var html = '<tr class="ab-section-total-row">' +
    '<td class="ab-frozen" style="left:0;"></td>' +
    '<td class="ab-frozen" style="left:30px;text-align:left;font-weight:700;">Section Total</td>' +
    '<td class="ab-frozen" style="left:210px;"></td>' +
    '<td class="ab-frozen" style="left:282px;"></td>';

  var yearTotals = [];

  years.forEach(function(y, yi) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yUnits = 0, yCost = 0;

    yqList.forEach(function(q) {
      var qi = quarters.indexOf(q);
      var qUnits = 0, qCost = 0;
      rows.forEach(function(row) {
        var u = (row.quarters[qi] || {}).units || 0;
        qUnits += u;
        qCost += u * (row.unit_cost || 0);
      });
      yUnits += qUnits;
      yCost += qCost;
      html += '<td class="ab-gt-cell">' + qUnits + '</td><td class="ab-gt-cell">' + ab_fc(qCost) + '</td>';
    });

    yearTotals.push({ units: yUnits, cost: yCost });
  });

  years.forEach(function(y, yi) {
    html += '<td class="ab-gt-cell">' + yearTotals[yi].units + '</td>' +
            '<td class="ab-gt-cell">' + ab_fc(yearTotals[yi].cost) + '</td>';
  });

  var y1c = (yearTotals[0] || {}).cost || 0;
  var y2c = (yearTotals[1] || {}).cost || 0;
  html += '<td class="ab-gt-cell ab-gt-final">' + ab_fc(y1c) + '</td>' +
          '<td class="ab-gt-cell ab-gt-final">' + ab_fc(y2c) + '</td>' +
          '<td class="ab-gt-cell ab-gt-final">' + ab_fc(y1c + y2c) + '</td>';

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

  // Save Non-Programmatic button — sequential saves to avoid server conflicts
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
      ab_saveProgData(frm, quarters, years, progData, sbhRevMap, fsMap);
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

  // For programmatic rows
  var idx = parseInt(tr.dataset.idx);
  if (!isNaN(idx) && progData.rows[idx]) {
    var row = progData.rows[idx];
    // Read unit_cost from the input field (cell index 4)
    var cells = tr.querySelectorAll('td');
    var ucInput = cells[4] && cells[4].querySelector('input');
    var uc = ucInput ? (parseFloat(ucInput.value) || 0) : (row.unit_cost || 0);
    row.unit_cost = uc; // Update stored value
    var ci = 5; // Start after 5 frozen cols

    var yearTotals = [];

    years.forEach(function(y) {
      var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
      var yTotalUnits = 0, yLic = 0, yGovt = 0, yBenf = 0;

      yqList.forEach(function(q) {
        var lu = parseInt(cells[ci].querySelector('input').value) || 0; ci++;
        cells[ci].textContent = ab_fc(lu * uc); ci++;
        var gu = parseInt(cells[ci].querySelector('input').value) || 0; ci++;
        cells[ci].textContent = ab_fc(gu * uc); ci++;
        var bu = parseInt(cells[ci].querySelector('input').value) || 0; ci++;
        cells[ci].textContent = ab_fc(bu * uc); ci++;
        cells[ci].textContent = ab_fc((lu + gu + bu) * uc); ci++;

        yTotalUnits += lu + gu + bu;
        yLic += lu * uc; yGovt += gu * uc; yBenf += bu * uc;
      });

      var yTotal = yLic + yGovt + yBenf;
      yearTotals.push(yTotal);

      cells[ci].textContent = yTotalUnits; ci++;
      cells[ci].textContent = ab_fc(yTotal); ci++;
      cells[ci].textContent = ab_fc(yLic); ci++;
      cells[ci].textContent = ab_fc(yGovt); ci++;
      cells[ci].textContent = ab_fc(yBenf); ci++;
    });

    // Grand total cols
    var y1 = yearTotals[0] || 0, y2 = yearTotals[1] || 0;
    cells[ci].textContent = ab_fc(y1); ci++;
    cells[ci].textContent = ab_fc(y2); ci++;
    cells[ci].textContent = ab_fc(y1 + y2);
  }

  // Recalculate the programmatic grand total row
  ab_recalcProgGrandTotal(quarters, years);
}

function ab_recalcProgGrandTotal(quarters, years) {
  var table = document.querySelector('.ab-prog-table');
  if (!table) return;
  var gtRow = table.querySelector('.ab-grand-total-row');
  if (!gtRow) return;
  var dataRows = table.querySelectorAll('.ab-data-row');
  if (!dataRows.length) return;

  var gtCells = gtRow.querySelectorAll('td');
  var gci = 5; // Skip 5 frozen cols (Sr, Activity, Task, UoM, UnitCost)
  var grandY = [];

  years.forEach(function(y) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yTotalUnits = 0, yLic = 0, yGovt = 0, yBenf = 0;

    yqList.forEach(function() {
      var qLu = 0, qLa = 0, qGu = 0, qGa = 0, qBu = 0, qBa = 0;
      dataRows.forEach(function(dr) {
        var drCells = dr.querySelectorAll('td');
        var ucInp = drCells[4] && drCells[4].querySelector('input');
        var uc = ucInp ? (parseFloat(ucInp.value) || 0) : 0;
        var licInp = drCells[gci] && drCells[gci].querySelector('input');
        var govtInp = drCells[gci + 2] && drCells[gci + 2].querySelector('input');
        var benfInp = drCells[gci + 4] && drCells[gci + 4].querySelector('input');
        var lu = licInp ? (parseInt(licInp.value) || 0) : 0;
        var gu = govtInp ? (parseInt(govtInp.value) || 0) : 0;
        var bu = benfInp ? (parseInt(benfInp.value) || 0) : 0;
        qLu += lu; qLa += lu * uc;
        qGu += gu; qGa += gu * uc;
        qBu += bu; qBa += bu * uc;
      });

      var qTa = qLa + qGa + qBa;
      yTotalUnits += qLu + qGu + qBu;
      yLic += qLa; yGovt += qGa; yBenf += qBa;

      gtCells[gci].textContent = qLu; gci++;
      gtCells[gci].textContent = ab_fc(qLa); gci++;
      gtCells[gci].textContent = qGu; gci++;
      gtCells[gci].textContent = ab_fc(qGa); gci++;
      gtCells[gci].textContent = qBu; gci++;
      gtCells[gci].textContent = ab_fc(qBa); gci++;
      gtCells[gci].textContent = ab_fc(qTa); gci++;
    });

    var yTotal = yLic + yGovt + yBenf;
    grandY.push(yTotal);

    gtCells[gci].textContent = yTotalUnits; gci++;
    gtCells[gci].textContent = ab_fc(yTotal); gci++;
    gtCells[gci].textContent = ab_fc(yLic); gci++;
    gtCells[gci].textContent = ab_fc(yGovt); gci++;
    gtCells[gci].textContent = ab_fc(yBenf); gci++;
  });

  var gy1 = grandY[0] || 0, gy2 = grandY[1] || 0;
  gtCells[gci].textContent = ab_fc(gy1); gci++;
  gtCells[gci].textContent = ab_fc(gy2); gci++;
  gtCells[gci].textContent = ab_fc(gy1 + gy2);
}

function ab_recalcNonProgRow(tr, quarters, years) {
  var cells = tr.querySelectorAll('td');
  // Read unit cost from input (cell index 3: Sr=0, Particulars=1, UoM=2, UnitCost=3)
  var ucInput = cells[3] && cells[3].querySelector('input');
  var uc = ucInput ? (parseFloat(ucInput.value) || 0) : 0;
  var ci = 4; // Start after 4 frozen cols

  var yearTotals = [];

  years.forEach(function(y) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yUnits = 0, yCost = 0;

    yqList.forEach(function(q) {
      var u = parseInt(cells[ci].querySelector('input').value) || 0; ci++;
      var c = u * uc;
      cells[ci].textContent = ab_fc(c); ci++;
      yUnits += u;
      yCost += c;
    });

    yearTotals.push({ units: yUnits, cost: yCost });
  });

  // Year totals
  years.forEach(function(y, yi) {
    cells[ci].textContent = yearTotals[yi].units; ci++;
    cells[ci].textContent = ab_fc(yearTotals[yi].cost); ci++;
  });

  // Grand total
  var y1c = (yearTotals[0] || {}).cost || 0;
  var y2c = (yearTotals[1] || {}).cost || 0;
  cells[ci].textContent = ab_fc(y1c); ci++;
  cells[ci].textContent = ab_fc(y2c); ci++;
  cells[ci].textContent = ab_fc(y1c + y2c);

  // Recalculate section total
  ab_recalcNonProgSectionTotal(tr, quarters, years);
}

function ab_recalcNonProgSectionTotal(dataRowTr, quarters, years) {
  var tbody = dataRowTr.closest('tbody');
  if (!tbody) return;
  var totalRow = tbody.querySelector('.ab-section-total-row');
  if (!totalRow) return;
  var dataRows = tbody.querySelectorAll('.ab-data-row');
  var gtCells = totalRow.querySelectorAll('td');
  var gci = 4; // Skip 4 frozen cols (Sr, Particulars, UoM, UnitCost)
  var yearTotals = [];

  years.forEach(function(y) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    var yUnits = 0, yCost = 0;

    yqList.forEach(function() {
      var qUnits = 0, qCost = 0;
      dataRows.forEach(function(dr) {
        var drCells = dr.querySelectorAll('td');
        var ucInp = drCells[3] && drCells[3].querySelector('input');
        var uc = ucInp ? (parseFloat(ucInp.value) || 0) : 0;
        var uInp = drCells[gci] && drCells[gci].querySelector('input');
        var u = uInp ? (parseInt(uInp.value) || 0) : 0;
        qUnits += u;
        qCost += u * uc;
      });
      yUnits += qUnits; yCost += qCost;
      gtCells[gci].textContent = qUnits; gci++;
      gtCells[gci].textContent = ab_fc(qCost); gci++;
    });

    yearTotals.push({ units: yUnits, cost: yCost });
  });

  years.forEach(function(y, yi) {
    gtCells[gci].textContent = yearTotals[yi].units; gci++;
    gtCells[gci].textContent = ab_fc(yearTotals[yi].cost); gci++;
  });

  var y1c = (yearTotals[0] || {}).cost || 0;
  var y2c = (yearTotals[1] || {}).cost || 0;
  gtCells[gci].textContent = ab_fc(y1c); gci++;
  gtCells[gci].textContent = ab_fc(y2c); gci++;
  gtCells[gci].textContent = ab_fc(y1c + y2c);
}

// ============================================================================
// ADD ROW + SAVE (Non-Programmatic)
// ============================================================================

function ab_addNonProgRow(frm, secTitle, quarters, years, sbhRevMap, nonProgData, progData, unitsList) {
  // Find the table for this section
  var sectionEl = null;
  document.querySelectorAll('.ab-section-header').forEach(function(hdr) {
    if (hdr.dataset.section === secTitle) sectionEl = hdr.closest('.ab-section');
  });
  if (!sectionEl) return;

  var tbody = sectionEl.querySelector('tbody');
  if (!tbody) return;

  // Find the section-total row (last row in tbody) to insert before it
  var totalRow = tbody.querySelector('.ab-section-total-row');

  // Determine new row index
  var dataRows = tbody.querySelectorAll('.ab-data-row');
  var newIdx = dataRows.length;

  // Build new row HTML
  var newRow = {
    description: '',
    assumption: '',
    unit_cost: 0,
    pbpName: null,
    quarters: {}
  };
  for (var i = 0; i < quarters.length; i++) {
    newRow.quarters[i] = { units: 0 };
  }

  // Add to nonProgData
  if (nonProgData.sections[secTitle]) {
    nonProgData.sections[secTitle].push(newRow);
  }

  var rowHtml = ab_buildNonProgRow(newRow, quarters, years, newIdx, secTitle, unitsList);
  var tempDiv = document.createElement('tbody');
  tempDiv.innerHTML = rowHtml;
  var newTr = tempDiv.querySelector('tr');

  if (totalRow) {
    tbody.insertBefore(newTr, totalRow);
  } else {
    tbody.appendChild(newTr);
  }

  // Update Sr numbers
  tbody.querySelectorAll('.ab-data-row').forEach(function(tr, i) {
    var srCell = tr.querySelector('.ab-sr');
    if (srCell) srCell.textContent = i + 1;
  });

  // Wire up events on the new row inputs (recalc only, no auto-save)
  newTr.querySelectorAll('.ab-inp').forEach(function(inp) {
    inp.addEventListener('change', function() {
      ab_recalcRow(this, quarters, years, progData, nonProgData);
    });
    inp.addEventListener('input', function() {
      ab_recalcRow(this, quarters, years, progData, nonProgData);
    });
  });

  // Focus the description input
  var descInp = newTr.querySelector('.ab-desc-inp');
  if (descInp) descInp.focus();
}

function ab_saveNonProgRow(frm, tr, quarters, years, sbhRevMap) {
  var secTitle = tr.dataset.section;
  var pbpName = tr.dataset.pbp || '';
  var cells = tr.querySelectorAll('td');

  // Read description
  var descInput = tr.querySelector('.ab-desc-inp');
  var description = descInput ? descInput.value.trim() : '';
  if (!description) return Promise.resolve(); // Don't save rows without a description

  // Read UoM
  var uomSelect = tr.querySelector('.ab-uom-sel');
  var uomValue = uomSelect ? uomSelect.value : '';

  // Read unit cost
  var ucInput = tr.querySelector('.ab-uc-inp');
  var unitCost = ucInput ? (parseFloat(ucInput.value) || 0) : 0;

  // Read quarterly units
  var planningRows = [];
  var ci = 4; // Start after 4 frozen cols
  years.forEach(function(y) {
    var yqList = quarters.filter(function(q) { return q.year_sequence === y.year_sequence; });
    yqList.forEach(function(q) {
      var uInput = cells[ci] && cells[ci].querySelector('input');
      var units = uInput ? (parseInt(uInput.value) || 0) : 0;
      ci += 2; // Skip units cell + cost cell

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
  });

  // Get sub-budget head and budget head IDs
  var sbhInfo = sbhRevMap[secTitle] || {};
  if (!sbhInfo.sbhId) {
    console.warn('[AB] No sbh mapping for section:', secTitle);
    return Promise.resolve();
  }

  if (pbpName) {
    // UPDATE existing PBP record
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
    // CREATE new PBP record
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

// Save non-prog rows one at a time to avoid server conflicts
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

async function ab_saveProgData(frm, quarters, years, progData, sbhRevMap, fsMap) {
  var rows = progData.rows || [];
  if (!rows.length) { frappe.show_alert({ message: 'No programmatic rows to save', indicator: 'orange' }); return; }

  var progSbh = sbhRevMap['Programmatic Costs'] || {};
  if (!progSbh.sbhId) {
    frappe.show_alert({ message: 'No sub-budget head mapping for "Programmatic Costs"', indicator: 'red' });
    return;
  }

  // Read current values from the DOM
  var table = document.querySelector('.ab-prog-table');
  if (!table) return;
  var dataRows = table.querySelectorAll('.ab-data-row');

  frappe.show_alert({ message: 'Saving programmatic data...', indicator: 'blue' });

  // Reverse-lookup fund source IDs from fsMap
  var fsIdByKey = {};
  Object.keys(fsMap).forEach(function(fsId) {
    var key = ab_normFS(fsMap[fsId]);
    if (!fsIdByKey[key]) fsIdByKey[key] = fsId;
  });

  var saveCount = 0;
  var errCount = 0;

  for (var ri = 0; ri < rows.length; ri++) {
    var row = rows[ri];
    var drEl = dataRows[ri];
    if (!drEl) continue;
    var drCells = drEl.querySelectorAll('td');

    // Read unit cost from input
    var ucInp = drCells[4] && drCells[4].querySelector('input');
    var uc = ucInp ? (parseFloat(ucInp.value) || 0) : 0;

    // For each fund source, collect quarterly units and save
    var fsKeys = ['lic', 'govt', 'benf'];
    var ci = 5; // Start after 5 frozen cols

    for (var yi = 0; yi < years.length; yi++) {
      var yqList = quarters.filter(function(q) { return q.year_sequence === years[yi].year_sequence; });

      for (var qj = 0; qj < yqList.length; qj++) {
        // Read LIC units, skip LIC amt, Govt units, skip Govt amt, Benf units, skip Benf amt, skip Total amt
        var licInp = drCells[ci] && drCells[ci].querySelector('input');
        var lu = licInp ? (parseInt(licInp.value) || 0) : 0;
        // ci+1 = LIC amt (calc)
        var govtInp = drCells[ci + 2] && drCells[ci + 2].querySelector('input');
        var gu = govtInp ? (parseInt(govtInp.value) || 0) : 0;
        // ci+3 = Govt amt (calc)
        var benfInp = drCells[ci + 4] && drCells[ci + 4].querySelector('input');
        var bu = benfInp ? (parseInt(benfInp.value) || 0) : 0;
        // ci+5 = Benf amt, ci+6 = Total amt

        var qi = quarters.indexOf(yqList[qj]);

        // Update in-memory data
        row.fundSources.lic.quarters[qi] = { units: lu };
        row.fundSources.govt.quarters[qi] = { units: gu };
        row.fundSources.benf.quarters[qi] = { units: bu };

        ci += 7;
      }
      ci += 5; // Skip year total subcols
    }
    row.unit_cost = uc;

    // Save each fund source that has a PBP record or units > 0
    for (var fi = 0; fi < fsKeys.length; fi++) {
      var fsKey = fsKeys[fi];
      var fs = row.fundSources[fsKey];
      var fsId = fsIdByKey[fsKey];

      // Build planning rows
      var planningRows = [];
      quarters.forEach(function(q, qi) {
        var units = (fs.quarters[qi] || {}).units || 0;
        planningRows.push({
          doctype: 'PBP Child',
          year: q.year,
          quarter: q.quarter,
          timespan: q.quarter,
          unit: units,
          unit_cost: uc,
          planned_amount: units * uc,
          start_date: q.start_date,
          end_date: q.end_date
        });
      });

      var totalBudget = planningRows.reduce(function(s, r) { return s + r.planned_amount; }, 0);

      // Only save if there's a PBP record or there are actual units
      if (fs.pbpName) {
        try {
          var r = await frappe.call({ method: 'frappe.client.get', args: { doctype: 'Project Budget Planning', name: fs.pbpName } });
          if (r.message) {
            var doc = r.message;
            doc.planning_table = planningRows;
            doc.total_planned_budget = totalBudget;
            await frappe.call({ method: 'frappe.client.save', args: { doc: doc } });
            saveCount++;
          }
        } catch (e) { console.error('[AB] Prog save error:', fs.pbpName, e); errCount++; }
      } else if (totalBudget > 0 && fsId) {
        try {
          var newDoc = {
            doctype: 'Project Budget Planning',
            project_proposal: frm.doc.name,
            donor: 'D-0001',
            description: row.description,
            budget_head: progSbh.bhId,
            sub_budget_head: progSbh.sbhId,
            fund_source: fsId,
            total_planned_budget: totalBudget,
            planning_table: planningRows
          };
          var cr = await frappe.call({ method: 'frappe.client.save', args: { doc: newDoc } });
          if (cr.message) {
            fs.pbpName = cr.message.name;
            saveCount++;
          }
        } catch (e) { console.error('[AB] Prog create error:', e); errCount++; }
      }
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
// HIDE BUDGET SUMMARY TAB (not yet implemented)
// ============================================================================

function ab_hideBudgetSummaryTab(frm) {
  // Hide the Budget Summary tab in Frappe's tab bar
  try {
    var tabLink = document.querySelector('[data-fieldname="custom_budget_summary_tab"]');
    if (tabLink) {
      var tabEl = tabLink.closest('.form-clickable-section') || tabLink.closest('.nav-item') || tabLink;
      if (tabEl) tabEl.style.display = 'none';
    }
    // Also try the tab content approach used by Frappe v16
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
  return '\
.ab-container { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; font-size: 12px; color: #333; padding: 8px; }\
.ab-tabs { display: flex; gap: 4px; margin-bottom: 10px; border-bottom: 2px solid #e0e0e0; }\
.ab-tab-btn { padding: 8px 16px; border: none; background: none; cursor: pointer; font-size: 13px; color: #666; border-bottom: 3px solid transparent; transition: all 0.2s; }\
.ab-tab-btn:hover { color: #333; }\
.ab-tab-btn.ab-tab-active { color: #1a5490; border-bottom-color: #1a5490; font-weight: 600; }\
.ab-tab-content { display: block; }\
.ab-tab-content.ab-hidden { display: none; }\
.ab-scroll-wrapper { overflow-x: auto; overflow-y: auto; max-height: 70vh; background: #fff; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }\
.ab-table { border-collapse: collapse; background: #fff; }\
.ab-prog-table { min-width: 100%; }\
.ab-nonprog-table { min-width: 100%; }\
.ab-table th, .ab-table td { border: 1px solid #ddd; padding: 3px 4px; font-size: 11px; white-space: nowrap; }\
.ab-table th { text-align: center; font-weight: 500; position: sticky; top: 0; z-index: 8; }\
.ab-table td { text-align: right; }\
.ab-table thead tr:nth-child(1) th { top: 0; z-index: 8; }\
.ab-table thead tr:nth-child(2) th { top: 26px; z-index: 8; }\
.ab-table thead tr:nth-child(3) th { top: 52px; z-index: 8; }\
.ab-frozen { position: sticky; z-index: 10; background: #f8f9fa; }\
.ab-frozen-header { position: sticky; z-index: 12; background: #f8f9fa; }\
.ab-table thead .ab-frozen { z-index: 13; }\
.ab-sr { text-align: center; font-weight: 500; width: 30px; min-width: 30px; }\
.ab-desc-cell { max-width: 180px; overflow: hidden; text-overflow: ellipsis; }\
.ab-header-row-1 th { background: #f0f4ff; font-weight: 600; padding: 5px 3px; }\
.ab-year-header { color: #1a5490; }\
.ab-year-header.ab-year-0 { background: #eef5ff; }\
.ab-year-header.ab-year-1 { background: #e8f5e9; color: #1b5e20; }\
.ab-grand-total-header { background: #fff9e6; color: #f57f17; }\
.ab-header-row-2 th, .ab-header-row-3 th { background: #f8f9fa; font-weight: 500; padding: 3px 3px; }\
.ab-quarter-header.ab-year-0 { background: #eef5ff; border-color: #c5d9f1; }\
.ab-quarter-header.ab-year-1 { background: #e8f5e9; border-color: #c8e6c9; }\
.ab-year-total-header.ab-year-0 { background: #dce8f8; font-weight: 600; }\
.ab-year-total-header.ab-year-1 { background: #d5ecd8; font-weight: 600; }\
.ab-subcol { background: #fafafa; font-size: 10px; }\
.ab-yt-subcol { background: #eef5ff; font-weight: 600; font-size: 10px; }\
.ab-yt-subcol.ab-year-1 { background: #e8f5e9; }\
.ab-gt-subcol { background: #fff3cd; font-weight: 700; font-size: 10px; }\
.ab-calc { background: #f4f7fa; font-family: "Monaco","Menlo",monospace; font-size: 11px; }\
.ab-total-col { background: #edf2f7; font-weight: 500; }\
.ab-yt-cell { background: #f0f5ff; font-weight: 500; font-family: "Monaco","Menlo",monospace; font-size: 11px; }\
.ab-gt-cell { background: #fffbf0; font-weight: 600; font-family: "Monaco","Menlo",monospace; font-size: 11px; border-color: #ffe0b2; }\
.ab-gt-final { background: #fff3cd; font-weight: 700; }\
.ab-grand-total-row { border-top: 2px solid #ffc107; }\
.ab-grand-total-row td { background: #fffbf0; font-weight: 700; }\
.ab-section-total-row td { background: #f5f5f5; font-weight: 600; border-top: 2px solid #ddd; }\
.ab-editable { background: #fffef5; }\
.ab-inp { width: 50px; padding: 2px 3px; border: 1px solid #ddd; background: #fffef5; border-radius: 2px; font-size: 11px; text-align: right; font-family: inherit; }\
.ab-inp:focus { outline: none; border-color: #4285f4; background: #fff; box-shadow: 0 0 3px rgba(66,133,244,0.3); }\
.ab-section { border: 1px solid #ddd; margin-bottom: 8px; border-radius: 4px; overflow: hidden; }\
.ab-section-header { background: #f5f5f5; padding: 8px 12px; cursor: pointer; font-weight: 600; font-size: 12px; user-select: none; display: flex; justify-content: space-between; align-items: center; }\
.ab-section-header:hover { background: #eee; }\
.ab-toggle { font-size: 10px; color: #999; }\
.ab-section-content { display: block; }\
.ab-data-row:hover td { background: #f0f7ff; }\
.ab-footer { margin-top: 10px; padding: 8px 12px; background: #f8f9fa; border-radius: 4px; font-size: 11px; display: flex; gap: 16px; align-items: center; }\
.ab-legend { display: flex; align-items: center; gap: 4px; }\
.ab-saved-indicator { margin-left: auto; color: #4caf50; font-weight: 500; display: none; }\
.ab-nonprog-wrapper { }\
.ab-add-row-btn { background: #fff; border: 1px dashed #aaa; color: #666; transition: all 0.2s; font-size: 11px; padding: 3px 10px; }\
.ab-add-row-btn:hover { background: #f0f7ff; border-color: #4285f4; color: #4285f4; }\
.ab-desc-inp { width: 170px; padding: 2px 3px; border: 1px solid #ddd; background: #fffef5; border-radius: 2px; font-size: 11px; text-align: left; font-family: inherit; }\
.ab-desc-inp:focus { outline: none; border-color: #4285f4; background: #fff; box-shadow: 0 0 3px rgba(66,133,244,0.3); }\
.ab-uom-sel { width: 70px; padding: 2px 1px; border: 1px solid #ddd; background: #fffef5; border-radius: 2px; font-size: 10px; font-family: inherit; cursor: pointer; }\
.ab-uom-sel:focus { outline: none; border-color: #4285f4; background: #fff; box-shadow: 0 0 3px rgba(66,133,244,0.3); }\
';
}
