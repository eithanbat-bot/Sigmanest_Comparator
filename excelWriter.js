const REPORT_SHEETS = [
  'Summary',
  'Tab Selection',
  'Mismatch',
  'Parts in CL but not in WS',
  'Parts in WS but not in CL',
  'Part Review',
  'Material Rules',
  'Thickness Rules',
  'Comparison Data'
];

const BLUE = '#1F4E78';
const BLUE_LIGHT = '#D9EAF7';
const BLUE_PALE = '#F4F8FB';
const BUTTON_BLUE = '#2E75B6';
const GREEN = '#107C10';
const GREEN_PALE = '#EAF4EA';
const RED = '#A4262C';
const RED_PALE = '#FCECEC';
const AMBER = '#805600';
const AMBER_PALE = '#FFF6D6';
const GREY = '#595959';
const BORDER = '#D9E2F3';
const WHITE = '#FFFFFF';

function stringify(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  return String(v);
}

async function freshSheet(context, name) {
  const old = context.workbook.worksheets.getItemOrNullObject(name);
  old.load('isNullObject');
  await context.sync();
  if (!old.isNullObject) {
    old.delete();
    await context.sync();
  }
  return context.workbook.worksheets.add(name);
}

async function existingOrNewSheet(context, name) {
  const sheet = context.workbook.worksheets.getItemOrNullObject(name);
  sheet.load('isNullObject');
  await context.sync();
  if (sheet.isNullObject) return context.workbook.worksheets.add(name);
  return sheet;
}

/*
 * Office.js RangeFormat.columnWidth is intentionally given larger values
 * than the old writer used. The previous writer treated character-style
 * widths (10, 16, 34, etc.) as if they were display widths, producing the
 * very narrow columns shown in the screenshots.
 */
function setWidths(sheet, widths) {
  for (const [col, width] of Object.entries(widths)) {
    sheet.getRange(`${col}:${col}`).format.columnWidth = width;
  }
}

function setRows(sheet, rangeAddress, height) {
  sheet.getRange(rangeAddress).format.rowHeight = height;
}

function cleanSheet(sheet) {
  sheet.showGridlines = false;
  sheet.showHeadings = true;
}

function title(sheet, text, lastCol, subtitle = '') {
  const r = sheet.getRange(`A1:${lastCol}1`);
  r.merge(false);
  r.getCell(0, 0).values = [[text]];
  r.format.font.name = 'Calibri';
  r.format.font.size = 18;
  r.format.font.bold = true;
  r.format.font.color = BLUE;
  r.format.horizontalAlignment = 'Left';
  r.format.verticalAlignment = 'Center';
  r.format.wrapText = false;
  r.format.rowHeight = 30;

  if (subtitle) {
    const s = sheet.getRange(`A2:${lastCol}2`);
    s.merge(false);
    s.getCell(0, 0).values = [[subtitle]];
    s.format.font.name = 'Calibri';
    s.format.font.size = 10;
    s.format.font.color = GREY;
    s.format.horizontalAlignment = 'Left';
    s.format.verticalAlignment = 'Center';
    s.format.wrapText = true;
    s.format.rowHeight = 28;
  } else {
    sheet.getRange(`A2:${lastCol}2`).format.rowHeight = 8;
  }
  sheet.getRange(`A3:${lastCol}3`).format.rowHeight = 8;
}

function sectionHeader(sheet, address, text) {
  const r = sheet.getRange(address);
  r.merge(false);
  r.getCell(0, 0).values = [[text]];
  r.format.fill.color = BLUE_PALE;
  r.format.font.name = 'Calibri';
  r.format.font.size = 11;
  r.format.font.bold = true;
  r.format.font.color = BLUE;
  r.format.horizontalAlignment = 'Left';
  r.format.verticalAlignment = 'Center';
  r.format.rowHeight = 22;
}

function header(sheet, row, cols, theme) {
  const r = sheet.getRangeByIndexes(row - 1, 0, 1, cols.length);
  r.values = [cols];
  r.format.fill.color = theme.fill;
  r.format.font.color = theme.text;
  r.format.font.name = theme.font;
  r.format.font.bold = true;
  r.format.font.size = theme.size;
  r.format.horizontalAlignment = 'Center';
  r.format.verticalAlignment = 'Center';
  r.format.wrapText = true;
  r.format.borders.getItem('EdgeTop').style = 'Continuous';
  r.format.borders.getItem('EdgeTop').color = theme.fill;
  r.format.borders.getItem('EdgeBottom').style = 'Continuous';
  r.format.borders.getItem('EdgeBottom').color = theme.fill;
  r.format.rowHeight = 34;
}

function body(sheet, startRow, values, cols) {
  if (!values.length) return;
  const normalized = values.map(row => {
    const out = Array.isArray(row) ? row.slice(0, cols) : [];
    while (out.length < cols) out.push('');
    return out.map(stringify);
  });
  const r = sheet.getRangeByIndexes(startRow - 1, 0, normalized.length, cols);
  r.values = normalized;
  r.format.font.name = 'Calibri';
  r.format.font.size = 10;
  r.format.horizontalAlignment = 'Center';
  r.format.verticalAlignment = 'Center';
  r.format.wrapText = false;
  r.format.borders.getItem('EdgeBottom').style = 'Continuous';
  r.format.borders.getItem('EdgeBottom').color = BORDER;
  r.format.rowHeight = 21;

  for (let i = 0; i < normalized.length; i += 2) {
    sheet.getRangeByIndexes(startRow - 1 + i, 0, 1, cols).format.fill.color = BLUE_PALE;
  }
}

function wrapColumns(sheet, columns, startRow, rowCount) {
  if (!rowCount) return;
  for (const col of columns) {
    sheet.getRange(`${col}${startRow}:${col}${startRow + rowCount - 1}`).format.wrapText = true;
  }
}

function align(sheet, columns, startRow, rowCount, direction) {
  if (!rowCount) return;
  for (const col of columns) {
    sheet.getRange(`${col}${startRow}:${col}${startRow + rowCount - 1}`).format.horizontalAlignment = direction;
  }
}

function freeze(sheet, row) {
  sheet.freezePanes.unfreeze();
  sheet.freezePanes.freezeRows(row);
}

function summaryButton(sheet, address, text) {
  const r = sheet.getRange(address);
  r.merge(false);
  r.getCell(0, 0).values = [[text]];
  r.format.fill.color = BUTTON_BLUE;
  r.format.font.name = 'Calibri';
  r.format.font.size = 10;
  r.format.font.color = WHITE;
  r.format.font.bold = true;
  r.format.horizontalAlignment = 'Center';
  r.format.verticalAlignment = 'Center';
  r.format.wrapText = true;
  r.format.borders.getItem('EdgeTop').style = 'Continuous';
  r.format.borders.getItem('EdgeTop').color = BUTTON_BLUE;
  r.format.borders.getItem('EdgeBottom').style = 'Continuous';
  r.format.borders.getItem('EdgeBottom').color = BUTTON_BLUE;
  r.format.borders.getItem('EdgeLeft').style = 'Continuous';
  r.format.borders.getItem('EdgeLeft').color = BUTTON_BLUE;
  r.format.borders.getItem('EdgeRight').style = 'Continuous';
  r.format.borders.getItem('EdgeRight').color = BUTTON_BLUE;
}

function mergedValue(sheet, address, value, options = {}) {
  const r = sheet.getRange(address);
  r.merge(false);
  r.getCell(0, 0).values = [[stringify(value)]];
  r.format.font.name = 'Calibri';
  r.format.font.size = options.size || 10;
  r.format.font.color = options.color || '#1F1F1F';
  r.format.font.bold = !!options.bold;
  r.format.horizontalAlignment = options.align || 'Left';
  r.format.verticalAlignment = 'Center';
  r.format.wrapText = options.wrap !== false;
  if (options.fill) r.format.fill.color = options.fill;
  if (options.height) r.format.rowHeight = options.height;
}

function styleActionColumn(sheet, startRow, values, col) {
  for (let i = 0; i < values.length; i++) {
    const v = String(values[i] || '').toUpperCase();
    const cell = sheet.getRange(`${col}${startRow + i}`);
    cell.format.font.bold = true;
    cell.format.horizontalAlignment = 'Center';
    if (v === 'ADD') {
      cell.format.fill.color = GREEN_PALE;
      cell.format.font.color = GREEN;
    } else if (v === 'REMOVE') {
      cell.format.fill.color = RED_PALE;
      cell.format.font.color = RED;
    } else if (v === 'REVIEW') {
      cell.format.fill.color = AMBER_PALE;
      cell.format.font.color = AMBER;
    }
  }
}

function styleTabStatuses(sheet, startRow, rows) {
  for (let i = 0; i < rows.length; i++) {
    const status = String(rows[i]?.[4] ?? '').toUpperCase();
    const cell = sheet.getRange(`E${startRow + i}`);
    cell.format.font.bold = true;
    cell.format.horizontalAlignment = 'Center';
    if (status === 'SELECTED' || status === 'DETECTED') {
      cell.format.fill.color = GREEN_PALE;
      cell.format.font.color = GREEN;
    } else if (status.includes('NO TABLE')) {
      cell.format.fill.color = AMBER_PALE;
      cell.format.font.color = AMBER;
    } else {
      cell.format.fill.color = BLUE_PALE;
      cell.format.font.color = GREY;
    }
  }
}

function styleTheme(theme) {
  return {
    fill: theme.fill || BLUE,
    text: theme.text || WHITE,
    font: theme.font || 'Calibri',
    size: Number(theme.size) || 10,
  };
}

export async function writeReports(
  result,
  rules = { materialRules: [], thicknessRules: [] },
  settings = {
    headerFillColor: BLUE,
    headerTextColor: WHITE,
    headerFontFamily: 'Calibri',
    headerFontSize: 10
  }
) {
  await Excel.run(async context => {
    const theme = styleTheme({
      fill: settings.headerFillColor,
      text: settings.headerTextColor,
      font: settings.headerFontFamily,
      size: settings.headerFontSize
    });

    // Rebuild user-facing report tabs. Audit is deliberately preserved.
    for (const name of REPORT_SHEETS.filter(x => x !== 'Audit' && x !== 'Comparison Data')) {
      await freshSheet(context, name);
    }

    const summary = context.workbook.worksheets.getItem('Summary');
    cleanSheet(summary);
    title(summary, 'SigmaNEST CL / WS Comparator', 'J', 'Control panel, run status and headline counts. See the other tabs for full detail.');
    sectionHeader(summary, 'A4:J4', 'Controls');
    summaryButton(summary, 'A5:B6', '1. Select CL Workbook');
    summaryButton(summary, 'C5:D6', '2. Select WS Workbook');
    summaryButton(summary, 'E5:F6', '3. Refresh Inputs');
    summaryButton(summary, 'G5:H6', '4. Run Comparison');
    summaryButton(summary, 'I5:J6', '5. Generate Revised WS');
    setRows(summary, 'A5:J6', 28);

    sectionHeader(summary, 'A8:J8', 'Run Information');
    mergedValue(summary, 'A10:C10', 'CL Workbook', { bold: true, fill: BLUE_PALE, wrap: false });
    mergedValue(summary, 'D10:J10', result.clFileName, { wrap: false });
    mergedValue(summary, 'A11:C11', 'WS Workbook', { bold: true, fill: BLUE_PALE, wrap: false });
    mergedValue(summary, 'D11:J11', result.wsFileName, { wrap: false });
    mergedValue(summary, 'A12:C12', 'Job Number (detected)', { bold: true, fill: BLUE_PALE, wrap: false });
    mergedValue(summary, 'D12:J12', result.jobNumber, { wrap: false });
    mergedValue(summary, 'A13:C13', 'Selected CL Tabs', { bold: true, fill: BLUE_PALE, wrap: false });
    mergedValue(summary, 'D13:J13', (result.selectedTabs || []).join(', '), { height: 42, wrap: true });
    mergedValue(summary, 'A14:C14', 'Run Date / Time', { bold: true, fill: BLUE_PALE, wrap: false });
    mergedValue(summary, 'D14:J14', new Date(result.runDateTime).toLocaleString(), { wrap: false });
    summary.getRange('A10:C14').format.borders.getItem('EdgeBottom').style = 'Continuous';
    summary.getRange('A10:C14').format.borders.getItem('EdgeBottom').color = BORDER;
    summary.getRange('D10:J14').format.borders.getItem('EdgeBottom').style = 'Continuous';
    summary.getRange('D10:J14').format.borders.getItem('EdgeBottom').color = BORDER;

    sectionHeader(summary, 'A16:J16', 'Headline Counts');
    const countRows = [
      ['CL part rows processed', result.validIncludedClRecords],
      ['Unique CL requirements', result.uniqueCLRequirements],
      ['WS part records (valid)', result.validWsRecords],
      ['Exact matches', result.exactMatches],
      ['Quantity-only mismatches', result.quantityOnlyMismatches],
      ['Material-involved mismatches', result.materialMismatches],
      ['Thickness-involved mismatches', result.thicknessMismatches],
      ['Total mismatches', result.totalMismatches],
      ['Parts to Add', result.addCount],
      ['Parts to Remove', result.removeCount],
      ['Parts requiring Part Review', result.partReviewCount],
      ['CL-side thickness/material conflicts', result.dataConflicts]
    ];
    for (let i = 0; i < countRows.length; i++) {
      const row = 17 + i;
      const label = summary.getRange(`A${row}:C${row}`);
      label.merge(false);
      label.getCell(0, 0).values = [[countRows[i][0]]];
      label.format.fill.color = BLUE_PALE;
      label.format.font.name = 'Calibri';
      label.format.font.size = 10;
      label.format.font.bold = true;
      label.format.horizontalAlignment = 'Left';
      label.format.verticalAlignment = 'Center';

      const value = summary.getRange(`D${row}:E${row}`);
      value.merge(false);
      value.getCell(0, 0).values = [[countRows[i][1]]];
      value.format.font.name = 'Calibri';
      value.format.font.size = 11;
      value.format.font.bold = true;
      value.format.horizontalAlignment = 'Right';
      value.format.verticalAlignment = 'Center';
      value.format.rowHeight = 22;

      summary.getRange(`A${row}:E${row}`).format.borders.getItem('EdgeBottom').style = 'Continuous';
      summary.getRange(`A${row}:E${row}`).format.borders.getItem('EdgeBottom').color = BORDER;
    }

    mergedValue(summary, 'A30:J30', 'Rule configuration used for this run: Material Rules and Thickness Rules sheets below.', { color: GREY, size: 9, height: 22 });
    setWidths(summary, {
      A: 135, B: 80, C: 80, D: 135, E: 90, F: 90, G: 110, H: 110, I: 130, J: 130
    });

    const tabRows = result.tabSelection || (result.selectedTabs || []).map(x => [true, x, '', '', 'SELECTED']);
    const tab = context.workbook.worksheets.getItem('Tab Selection');
    cleanSheet(tab);
    title(tab, 'CL Worksheet Tab Selection', 'E', 'Every sheet found in the selected CL workbook is listed below. Tick Include for the sheets that should feed the comparison.');
    header(tab, 4, ['Include', 'Sheet Name', 'Detected Header Row', 'Detected Part Rows', 'Status'], theme);
    body(tab, 5, tabRows, 5);
    setWidths(tab, { A: 65, B: 225, C: 125, D: 125, E: 170 });
    align(tab, ['A', 'C', 'D', 'E'], 5, tabRows.length, 'Center');
    styleTabStatuses(tab, 5, tabRows);
    setRows(tab, 'A2:E2', 30);
    freeze(tab, 4);

    const mismatchRows = result.rows.filter(x => x.status === 'Mismatch');
    const mismatch = context.workbook.worksheets.getItem('Mismatch');
    cleanSheet(mismatch);
    title(mismatch, 'Mismatch — parts present in both CL and WS with one or more differences', 'K', 'Review material, thickness and quantity differences before generating a revised WS.');
    header(mismatch, 4, ['Part Number', 'WS Material', 'CL Material', 'WS Thickness', 'CL Thickness', 'WS Quantity', 'CL Quantity', 'Mismatch Type', 'CL Source Tabs', 'Reviewed?', 'Notes'], theme);
    body(mismatch, 5, mismatchRows.map(x => [
      x.partNoRaw, x.wsMaterial, x.clMaterial, x.wsThickness ?? '', x.clThickness, x.wsQty ?? '', x.clQty,
      x.mismatchType, x.sourceTabs, '', ''
    ]), 11);
    setWidths(mismatch, { A: 105, B: 135, C: 135, D: 75, E: 75, F: 78, G: 78, H: 190, I: 205, J: 78, K: 165 });
    wrapColumns(mismatch, ['H', 'I', 'K'], 5, mismatchRows.length);
    align(mismatch, ['D', 'E', 'F', 'G', 'J'], 5, mismatchRows.length, 'Center');
    freeze(mismatch, 4);

    const addRows = result.rows.filter(x => x.status === 'Add');
    const add = context.workbook.worksheets.getItem('Parts in CL but not in WS');
    cleanSheet(add);
    title(add, 'Parts required by the selected CL tabs but absent from the WS', 'H', 'These items are candidates for addition to the WS when the match is unambiguous.');
    header(add, 4, ['Part Number', 'Description', 'Material', 'Thickness', 'Required Quantity', 'Source CL Tabs', 'Action', 'Notes'], theme);
    body(add, 5, addRows.map(x => [
      x.partNoRaw, x.description, x.clMaterial, x.clThickness, x.clQty, x.sourceTabs, 'ADD', x.mismatchType
    ]), 8);
    setWidths(add, { A: 125, B: 250, C: 165, D: 90, E: 105, F: 230, G: 90, H: 220 });
    wrapColumns(add, ['B', 'F', 'H'], 5, addRows.length);
    align(add, ['D', 'E', 'G'], 5, addRows.length, 'Center');
    styleActionColumn(add, 5, addRows.map(() => 'ADD'), 'G');
    freeze(add, 4);

    const removeRows = result.removals || [];
    const rem = context.workbook.worksheets.getItem('Parts in WS but not in CL');
    cleanSheet(rem);
    title(rem, 'Parts currently in the WS which are not required by any selected CL tab', 'H', 'Only unambiguous removals are listed here. Family/variant review items stay out of this list.');
    header(rem, 4, ['Part Number', 'Material', 'Thickness', 'Current Quantity', 'WS Row', 'Action', 'Notes', 'Match Row'], theme);
    body(rem, 5, removeRows.map(x => [
      x.partNoRaw, x.materialCanonical, x.thicknessRaw ?? '', x.qtyRaw ?? '', x.sourceRow, 'REMOVE', x.reason, x.removeRank
    ]), 8);
    setWidths(rem, { A: 125, B: 165, C: 90, D: 110, E: 80, F: 90, G: 260, H: 100 });
    wrapColumns(rem, ['G'], 5, removeRows.length);
    align(rem, ['C', 'D', 'E', 'F', 'H'], 5, removeRows.length, 'Center');
    styleActionColumn(rem, 5, removeRows.map(() => 'REMOVE'), 'F');
    freeze(rem, 4);

    const reviewRows = result.rows.filter(x => x.status === 'Part Review');
    const review = context.workbook.worksheets.getItem('Part Review');
    cleanSheet(review);
    title(review, 'Part Review — CL parts with no exact WS match but a likely family/variant relative', 'N', 'These are NEVER auto-merged, auto-added, or auto-removed. Review them before changing the WS.');
    header(review, 4, ['CL Part Number', 'WS Part Number', 'Description', 'CL Material', 'WS Material', 'CL Thickness', 'WS Thickness', 'CL Qty', 'WS Qty', 'Family Key', 'Reason', 'Reviewed?', 'Notes', 'Action'], theme);
    body(review, 5, reviewRows.map(x => [
      x.partNoRaw, x.partReviewWsPartNo ?? '', x.description, x.clMaterial, x.partReviewWsMaterial ?? '',
      x.clThickness, x.partReviewWsThickness ?? '', x.clQty, x.partReviewWsQty ?? '', x.familyKey, x.mismatchType, '', '', 'REVIEW'
    ]), 14);
    setWidths(review, { A: 125, B: 135, C: 240, D: 165, E: 165, F: 90, G: 90, H: 80, I: 80, J: 125, K: 300, L: 85, M: 220, N: 85 });
    wrapColumns(review, ['C', 'K', 'M'], 5, reviewRows.length);
    align(review, ['F', 'G', 'H', 'I', 'L', 'N'], 5, reviewRows.length, 'Center');
    styleActionColumn(review, 5, reviewRows.map(() => 'REVIEW'), 'N');
    freeze(review, 4);

    const audit = await existingOrNewSheet(context, 'Audit');
    cleanSheet(audit);
    title(audit, 'Audit Log — one row is appended by the Office add-in at the end of every comparison run', 'N', 'Historical comparison runs remain in this sheet.');
    header(audit, 4, ['Run Date/Time', 'CL Filename', 'WS Filename', 'Job Number', 'Selected CL Tabs', 'CL Records', 'WS Records', 'Matched', 'Mismatched', 'Added', 'Removed', 'Part Review', 'Status', 'Notes'], theme);
    const usedAudit = audit.getUsedRangeOrNullObject();
    usedAudit.load(['isNullObject', 'rowCount']);
    await context.sync();
    const nextAuditRow = usedAudit.isNullObject ? 5 : Math.max(5, usedAudit.rowCount + 1);
    body(audit, nextAuditRow, [[
      new Date(result.runDateTime).toLocaleString(), result.clFileName, result.wsFileName, result.jobNumber,
      (result.selectedTabs || []).join(', '), result.validIncludedClRecords, result.validWsRecords,
      result.exactMatches, result.totalMismatches, result.addCount, result.removeCount,
      result.partReviewCount, 'READY FOR REVIEW', 'Comparison run by SigmaNEST Comparator Office Add-in'
    ]], 14);
    setWidths(audit, { A: 140, B: 280, C: 190, D: 90, E: 330, F: 90, G: 90, H: 85, I: 100, J: 80, K: 80, L: 95, M: 135, N: 250 });
    wrapColumns(audit, ['E', 'N'], 5, Math.max(1, nextAuditRow - 4));
    align(audit, ['A', 'D', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'], 5, Math.max(1, nextAuditRow - 4), 'Center');
    freeze(audit, 4);

    const mr = context.workbook.worksheets.getItem('Material Rules');
    cleanSheet(mr);
    title(mr, 'Material Equivalence Rules', 'E', 'Define which materials may be treated as interchangeable, and at which thickness. Only ACTIVE rows are used.');
    const materialRuleRows = (rules.materialRules || []).map(x => [!!x.active, x.materialA || '', x.materialB || '', x.thicknessRule ?? '', x.reason || '']);
    header(mr, 4, ['Active', 'Material A', 'Material B', 'Thickness Rule (mm, blank = any)', 'Reason / Notes'], theme);
    body(mr, 5, materialRuleRows, 5);
    setWidths(mr, { A: 65, B: 180, C: 180, D: 250, E: 330 });
    align(mr, ['A', 'D'], 5, materialRuleRows.length, 'Center');
    wrapColumns(mr, ['E'], 5, materialRuleRows.length);
    freeze(mr, 4);

    const tr = context.workbook.worksheets.getItem('Thickness Rules');
    cleanSheet(tr);
    title(tr, 'Thickness Normalization Rules', 'E', 'Define thickness roundings/equivalences. Material Filter is matched as a substring; blank applies to any material. Only ACTIVE rows are used.');
    const thicknessRuleRows = (rules.thicknessRules || []).map(x => [!!x.active, x.materialFilter || '', x.fromThickness ?? '', x.toThickness ?? '', x.reason || '']);
    header(tr, 4, ['Active', 'Material Filter (blank = any)', 'From Thickness (mm)', 'To Thickness (mm)', 'Reason / Notes'], theme);
    body(tr, 5, thicknessRuleRows, 5);
    setWidths(tr, { A: 65, B: 220, C: 150, D: 150, E: 330 });
    align(tr, ['A', 'C', 'D'], 5, thicknessRuleRows.length, 'Center');
    wrapColumns(tr, ['B', 'E'], 5, thicknessRuleRows.length);
    freeze(tr, 4);

    const data = await freshSheet(context, 'Comparison Data');
    cleanSheet(data);
    header(data, 1, [
      'ReqIndex', 'PartNo_Norm', 'PartNo_Raw', 'Description', 'CL_Material', 'CL_Thickness', 'CL_Qty', 'SourceTabs',
      'CL_ThicknessConflict', 'CL_MaterialConflict', 'WS_MatchFound', 'WS_RecordKey', 'WS_Material', 'WS_Thickness',
      'WS_Qty', 'WS_SourceRow', 'WS_ThicknessConflict', 'WS_MaterialConflict', 'MaterialMatch', 'ThicknessMatch',
      'QtyMatch', 'MismatchType', 'Status', 'FamilyKey'
    ], theme);
    const dataRows = result.rows.map(x => [
      x.reqIndex, x.partNoNorm, x.partNoRaw, x.description, x.clMaterial, x.clThickness, x.clQty, x.sourceTabs,
      x.clThicknessConflict, x.clMaterialConflict, x.wsMatchFound, x.wsRecordKey, x.wsMaterial, x.wsThickness ?? '',
      x.wsQty ?? '', x.wsSourceRow ?? '', x.wsThicknessConflict, x.wsMaterialConflict, x.materialMatch,
      x.thicknessMatch, x.qtyMatch, x.mismatchType, x.status, x.familyKey
    ]);
    body(data, 2, dataRows, 24);
    setWidths(data, {
      A: 70, B: 140, C: 125, D: 230, E: 165, F: 90, G: 80, H: 220, I: 150, J: 150,
      K: 110, L: 150, M: 165, N: 90, O: 80, P: 90, Q: 150, R: 150, S: 110, T: 110, U: 100, V: 250, W: 120, X: 140
    });
    wrapColumns(data, ['D', 'H', 'V'], 2, dataRows.length);
    align(data, ['A', 'F', 'G', 'I', 'J', 'K', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'W'], 2, dataRows.length, 'Center');
    freeze(data, 1);
    data.visibility = 'Hidden';

    // Final pass: center every used/populated cell horizontally and vertically.
    // This intentionally overrides sheet-specific alignments so every report tab
    // has one consistent visual language.
    for (const sheet of [summary, tab, mismatch, add, rem, review, audit, mr, tr]) {
      const used = sheet.getUsedRangeOrNullObject();
      used.load('isNullObject');
      await context.sync();
      if (!used.isNullObject) {
        used.format.horizontalAlignment = 'Center';
        used.format.verticalAlignment = 'Center';
      }
    }

    summary.activate();
    await context.sync();
  });
}
