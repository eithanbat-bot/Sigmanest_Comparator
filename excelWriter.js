const managedSheets = [
    'Summary', 'Tab Selection', 'Mismatch', 'Parts in CL but not in WS', 'Parts in WS but not in CL',
    'Part Review', 'Audit', 'Material Rules', 'Thickness Rules', 'Comparison Data'
];
const BLUE = '#1F4E78';
const BUTTON_BLUE = '#2E75B6';
const GREEN = '#107C10';
const GREY = '#595959';
function stringify(v) { if (v === null || v === undefined) return ''; if (typeof v === 'number' || typeof v === 'boolean') return v; return String(v); }
async function getOrCreateSheet(context, name) {
    const item = context.workbook.worksheets.getItemOrNullObject(name); item.load('isNullObject'); await context.sync();
    if (!item.isNullObject) { item.delete(); await context.sync(); }
    return context.workbook.worksheets.add(name);
}
function setWidths(sheet, widths) { for (const [col, width] of Object.entries(widths)) sheet.getRange(`${col}:${col}`).format.columnWidth = width; }
function title(sheet, text, lastCol, subtitle='') {
    const r = sheet.getRange(`A1:${lastCol}1`);
    r.merge(false);
    r.getCell(0,0).values=[[text]];
    r.format.font.name='Calibri'; r.format.font.size=16; r.format.font.bold=true; r.format.font.color=BLUE; r.format.horizontalAlignment='Center'; r.format.verticalAlignment='Center';
    if (subtitle) {
        const s=sheet.getRange(`A2:${lastCol}2`);
        s.merge(false);
        s.getCell(0,0).values=[[subtitle]];
        s.format.font.name='Calibri'; s.format.font.size=10; s.format.font.color=GREY; s.format.wrapText=true;
    }
}
function header(sheet, row, cols) { const r=sheet.getRangeByIndexes(row-1,0,1,cols.length); r.values=[cols]; r.format.fill.color=BLUE; r.format.font.color='#FFFFFF'; r.format.font.bold=true; r.format.font.size=10; r.format.horizontalAlignment='Center'; r.format.verticalAlignment='Center'; r.format.wrapText=true; }
function body(sheet, startRow, values, cols) {
    if(!values.length)return;
    const normalized = values.map(row => {
        const out = Array.isArray(row) ? row.slice(0, cols) : [];
        while (out.length < cols) out.push('');
        return out.map(stringify);
    });
    const r=sheet.getRangeByIndexes(startRow-1,0,normalized.length,cols);
    r.values=normalized;
    r.format.font.name='Calibri'; r.format.font.size=11; r.format.wrapText=true; r.format.verticalAlignment='Center';
}
function addGrid(sheet, rangeAddress) { const r=sheet.getRange(rangeAddress); r.format.borders.getItem('EdgeTop').style='Continuous'; r.format.borders.getItem('EdgeBottom').style='Continuous'; r.format.borders.getItem('EdgeLeft').style='Continuous'; r.format.borders.getItem('EdgeRight').style='Continuous'; r.format.borders.getItem('InsideHorizontal').style='Continuous'; r.format.borders.getItem('InsideVertical').style='Continuous'; }
function freeze(sheet,row){ sheet.freezePanes.freezeRows(row); }
function summaryButton(sheet, address, text) { const r=sheet.getRange(address); r.merge(false); r.getCell(0,0).values=[[text]]; r.format.fill.color=BUTTON_BLUE; r.format.font.color='#FFFFFF'; r.format.font.bold=true; r.format.horizontalAlignment='Center'; r.format.verticalAlignment='Center'; }
export async function writeReports(result, rules={materialRules:[], thicknessRules:[]}) {
 await Excel.run(async context=>{
  for(const n of managedSheets){const old=context.workbook.worksheets.getItemOrNullObject(n);old.load('isNullObject');await context.sync();if(!old.isNullObject){old.delete();await context.sync();}}
  const summary=await getOrCreateSheet(context,'Summary');
  title(summary,'SigmaNEST CL / WS Comparator','I','Control panel, run status and headline counts.');
  summary.getRange('A4:I4').merge(false); summary.getRange('A4').values=[['Controls']]; summary.getRange('A4').format.font.bold=true;
  summaryButton(summary,'A5:B6','1. Select CL Workbook'); summaryButton(summary,'C5:D6','2. Select WS Workbook'); summaryButton(summary,'E5:F6','3. Refresh Inputs'); summaryButton(summary,'G5:H6','4. Run Comparison'); summaryButton(summary,'I5:I6','5. Generate Revised WS');
  summary.getRange('A8:I8').merge(false); summary.getRange('A8').values=[['Run Information']]; summary.getRange('A8').format.font.bold=true;
  summary.getRange('A10:B14').values=[['CL Workbook',result.clFileName],['WS Workbook',result.wsFileName],['Job Number (detected)',result.jobNumber],['Selected CL Tabs',result.selectedTabs.join(', ')],['Run Date/Time',new Date(result.runDateTime).toLocaleString()]];
  summary.getRange('A16:B16').merge(false); summary.getRange('A16').values=[['Headline Counts']]; summary.getRange('A16').format.font.bold=true;
  summary.getRange('A17:B28').values=[['CL part rows processed',result.validIncludedClRecords],['Unique CL requirements',result.uniqueClRequirements],['WS part records (valid)',result.validWsRecords],['Exact matches',result.exactMatches],['Quantity-only mismatches',result.quantityOnlyMismatches],['Material-involved mismatches',result.materialMismatches],['Thickness-involved mismatches',result.thicknessMismatches],['Total mismatches',result.totalMismatches],['Parts to Add',result.addCount],['Parts to Remove',result.removeCount],['Parts requiring Part Review',result.partReviewCount],['CL-side thickness/material conflicts',result.dataConflicts]];
  summary.getRange('A30:I30').merge(false); summary.getRange('A30').values=[['Rule configuration used for this run: Material Rules and Thickness Rules sheets below.']]; summary.getRange('A30').format.font.color=GREY;
  setWidths(summary,{A:46,B:22,C:10,D:10,E:10,F:10,G:20,H:20,I:20});
  summary.getRange('A1').format.font.size=20; summary.getRange('A1').format.horizontalAlignment='Left'; summary.getRange('A2').format.horizontalAlignment='Left'; summary.getRange('A10:B28').format.font.size=10; addGrid(summary,'A10:B28');

  const tab=await getOrCreateSheet(context,'Tab Selection'); title(tab,'CL Worksheet Tab Selection','E','Every sheet found in the selected CL workbook is listed below. Tick Include for the sheets that should feed the comparison.'); header(tab,4,['Include','Sheet Name','Detected Header Row','Detected Part Rows','Status']);
  body(tab,5,result.tabSelection||result.selectedTabs.map(x=>[true,x,'','','SELECTED']),5); setWidths(tab,{A:10,B:34,C:18,D:16,E:40}); freeze(tab,4); addGrid(tab,`A4:E${Math.max(4,(result.tabSelection||[]).length+4)}`);

  const mismatch=await getOrCreateSheet(context,'Mismatch'); title(mismatch,'Mismatch — parts present in both CL and WS with one or more differences','K'); header(mismatch,3,['Part Number','WS Material','CL Material','WS Thickness','CL Thickness','WS Quantity','CL Quantity','Mismatch Type','CL Source Tabs','Reviewed?','Notes']);
  body(mismatch,4,result.rows.filter(x=>x.status==='Mismatch').map(x=>[x.partNoRaw,x.wsMaterial,x.clMaterial,x.wsThickness??'',x.clThickness,x.wsQty??'',x.clQty,x.mismatchType,x.sourceTabs,'','']),11); setWidths(mismatch,{A:16,B:16,C:16,D:12,E:12,F:12,G:12,H:38,I:46,J:14,K:30}); freeze(mismatch,3); addGrid(mismatch,`A3:K${Math.max(3,result.rows.filter(x=>x.status==='Mismatch').length+3)}`);

  const add=await getOrCreateSheet(context,'Parts in CL but not in WS'); title(add,'Parts required by the selected CL tabs but absent from the WS','H'); header(add,3,['Part Number','Description','Material','Thickness','Required Quantity','Source CL Tabs','Action','Notes']);
  body(add,4,result.rows.filter(x=>x.status==='Add').map(x=>[x.partNoRaw,x.description,x.clMaterial,x.clThickness,x.clQty,x.sourceTabs,'ADD',x.mismatchType]),8); setWidths(add,{A:16,B:34,C:16,D:12,E:16,F:26,G:10,H:30}); freeze(add,3); addGrid(add,`A3:H${Math.max(3,result.rows.filter(x=>x.status==='Add').length+3)}`);

  const rem=await getOrCreateSheet(context,'Parts in WS but not in CL'); title(rem,'Parts currently in the WS which are not required by any selected CL tab','H'); header(rem,3,['Part Number','Material','Thickness','Current Quantity','WS Row','Action','Notes','Match Row']);
  body(rem,4,result.removals.map(x=>[x.partNoRaw,x.materialCanonical,x.thicknessRaw??'',x.qtyRaw??'',x.sourceRow,'REMOVE',x.reason,x.removeRank]),8); setWidths(rem,{A:16,B:16,C:12,D:14,E:10,F:10,G:30,H:13}); freeze(rem,3); addGrid(rem,`A3:H${Math.max(3,result.removals.length+3)}`);

  const review=await getOrCreateSheet(context,'Part Review'); title(review,'Part Review — CL parts with no exact WS match but a likely family/variant relative','N','These are NEVER auto-merged, auto-added, or auto-removed. Review them before changing the WS.'); header(review,4,['CL Part Number','WS Part Number','Description','CL Material','WS Material','CL Thickness','WS Thickness','CL Qty','WS Qty','Family Key','Reason','Reviewed?','Notes','Action']);
  body(review,5,result.rows.filter(x=>x.status==='Part Review').map(x=>[x.partNoRaw,x.partReviewWsPartNo??'',x.description,x.clMaterial,x.partReviewWsMaterial??'',x.clThickness,x.partReviewWsThickness??'',x.clQty,x.partReviewWsQty??'',x.familyKey,x.mismatchType,'','','REVIEW']),14); setWidths(review,{A:16,B:18,C:30,D:16,E:24,F:14,G:14,H:10,I:10,J:16,K:48,L:13,M:26,N:13}); freeze(review,4); addGrid(review,`A4:N${Math.max(4,result.rows.filter(x=>x.status==='Part Review').length+4)}`);

  const audit=await getOrCreateSheet(context,'Audit'); title(audit,'Audit Log — one row is appended by the Office add-in at the end of every comparison run','N'); header(audit,3,['Run Date/Time','CL Filename','WS Filename','Job Number','Selected CL Tabs','CL Records','WS Records','Matched','Mismatched','Added','Removed','Part Review','Status','Notes']);
  body(audit,4,[[new Date(result.runDateTime).toLocaleString(),result.clFileName,result.wsFileName,result.jobNumber,result.selectedTabs.join(', '),result.validIncludedClRecords,result.validWsRecords,result.exactMatches,result.totalMismatches,result.addCount,result.removeCount,result.partReviewCount,'READY FOR REVIEW','Comparison run by SigmaNEST Comparator Office Add-in']],14); setWidths(audit,{A:18,B:40,C:20,D:12,E:50,F:10,G:10,H:10,I:12,J:10,K:10,L:12,M:22,N:30}); freeze(audit,3); addGrid(audit,'A3:N4');

  const mr=await getOrCreateSheet(context,'Material Rules'); title(mr,'Material Equivalence Rules','E','Define which materials may be treated as interchangeable, and at which thickness. Only ACTIVE rows are used by the comparison engine.'); header(mr,4,['Active','Material A','Material B','Thickness Rule (mm, blank = any)','Reason / Notes']);
  body(mr,5,(rules.materialRules||[]).map(x=>[!!x.active,x.materialA||'',x.materialB||'',x.thicknessRule??'',x.reason||''] ),5); setWidths(mr,{A:10,B:22,C:22,D:30,E:55}); freeze(mr,4); addGrid(mr,`A4:E${Math.max(4,(rules.materialRules||[]).length+4)}`);

  const tr=await getOrCreateSheet(context,'Thickness Rules'); title(tr,'Thickness Normalization Rules','E','Define thickness roundings/equivalences. Material Filter is matched as a substring against the canonical material text (blank = applies to any material). Only ACTIVE rows are used.'); header(tr,4,['Active','Material Filter (blank = any)','From Thickness (mm)','To Thickness (mm)','Reason / Notes']);
  body(tr,5,(rules.thicknessRules||[]).map(x=>[!!x.active,x.materialFilter||'',x.fromThickness??'',x.toThickness??'',x.reason||'']),5); setWidths(tr,{A:10,B:26,C:20,D:20,E:45}); freeze(tr,4); addGrid(tr,`A4:E${Math.max(4,(rules.thicknessRules||[]).length+4)}`);

  const data=await getOrCreateSheet(context,'Comparison Data'); header(data,1,['ReqIndex','PartNo_Norm','PartNo_Raw','Description','CL_Material','CL_Thickness','CL_Qty','SourceTabs','CL_ThicknessConflict','CL_MaterialConflict','WS_MatchFound','WS_RecordKey','WS_Material','WS_Thickness','WS_Qty','WS_SourceRow','WS_ThicknessConflict','WS_MaterialConflict','MaterialMatch','ThicknessMatch','QtyMatch','MismatchType','Status','FamilyKey']); body(data,2,result.rows.map(x=>[x.reqIndex,x.partNoNorm,x.partNoRaw,x.description,x.clMaterial,x.clThickness,x.clQty,x.sourceTabs,x.clThicknessConflict,x.clMaterialConflict,x.wsMatchFound,x.wsRecordKey,x.wsMaterial,x.wsThickness??'',x.wsQty??'',x.wsSourceRow??'',x.wsThicknessConflict,x.wsMaterialConflict,x.materialMatch,x.thicknessMatch,x.qtyMatch,x.mismatchType,x.status,x.familyKey]),24); data.getRange('A1:X1').format.fill.color=BLUE; data.getRange('A1:X1').format.font.color='#FFFFFF'; data.getRange('A1:X1').format.font.bold=true; data.getRange('A1:X1').format.wrapText=true; data.getRange('A:X').format.autofitColumns(); freeze(data,1);

  for(const s of [summary,tab,mismatch,add,rem,review,audit,mr,tr,data]) s.getUsedRange().format.font.name='Calibri';
  summary.activate(); await context.sync();
 });
}
