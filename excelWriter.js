const managedSheets = [
    'SN Summary', 'SN Mismatch', 'SN Add', 'SN Remove', 'SN Part Review', 'SN Comparison Data', 'SN Audit'
];
function stringify(v) {
    if (v === null || v === undefined)
        return '';
    if (typeof v === 'number' || typeof v === 'boolean')
        return v;
    return String(v);
}
async function replaceSheet(context, name) {
    let sheet;
    const item = context.workbook.worksheets.getItemOrNullObject(name);
    item.load('isNullObject');
    await context.sync();
    if (!item.isNullObject) {
        sheet = item;
        sheet.delete();
        await context.sync();
    }
    sheet = context.workbook.worksheets.add(name);
    return sheet;
}
function writeTable(sheet, values) {
    if (!values.length)
        return;
    const maxCols = Math.max(...values.map(r => r.length));
    const padded = values.map(r => [...r, ...Array(maxCols - r.length).fill('')]);
    const range = sheet.getRangeByIndexes(0, 0, padded.length, maxCols);
    range.values = padded;
    const header = sheet.getRangeByIndexes(0, 0, 1, maxCols);
    header.format.font.bold = true;
    header.format.fill.color = '#1F4E78';
    header.format.font.color = '#FFFFFF';
    range.format.autofitColumns();
    range.format.autofitRows();
    sheet.freezePanes.freezeRows(1);
}
export async function writeReports(result) {
    await Excel.run(async (context) => {
        for (const n of managedSheets) {
            const old = context.workbook.worksheets.getItemOrNullObject(n);
            old.load('isNullObject');
            await context.sync();
            if (!old.isNullObject)
                old.delete();
        }
        await context.sync();
        const summary = context.workbook.worksheets.add('SN Summary');
        const summaryRows = [
            ['SigmaNEST CL / WS Comparator', ''],
            ['Run date/time', new Date(result.runDateTime).toLocaleString()],
            ['CL workbook', result.clFileName],
            ['WS workbook', result.wsFileName],
            ['Job number', result.jobNumber],
            ['Selected CL tabs', result.selectedTabs.join(', ')],
            ['Status', 'READY FOR REVIEW'],
            [''],
            ['Comparison Counts', ''],
            ['CL part rows processed', result.validIncludedClRecords],
            ['Unique CL requirements', result.uniqueClRequirements],
            ['WS part records (valid)', result.validWsRecords],
            ['Exact matches', result.exactMatches],
            ['Quantity-only mismatches', result.quantityOnlyMismatches],
            ['Material-involved mismatches', result.materialMismatches],
            ['Thickness-involved mismatches', result.thicknessMismatches],
            ['Total mismatches', result.totalMismatches],
            ['Parts to Add', result.addCount],
            ['Parts to Remove', result.removeCount],
            ['Parts requiring Part Review', result.partReviewCount],
            ['CL-side thickness/material conflicts', result.dataConflicts],
            ['CL rows skipped - invalid/zero quantity', result.invalidIncludedClRecords],
            ['Total unresolved issues', result.totalMismatches + result.addCount + result.removeCount + result.partReviewCount],
            [''],
            ['Rule: 4 mm Armox is treated as 4 mm Ramor 500.', ''],
            ['Safe revised-WS changes only: quantity-only updates, definitive removes, and definitive adds. Material/thickness mismatches and Part Review remain untouched.', ''],
        ].map(r => r.map(stringify));
        writeTable(summary, summaryRows);
        summary.getRange('A1').format.font.size = 16;
        summary.getRange('A1').format.font.bold = true;
        const mismatch = context.workbook.worksheets.add('SN Mismatch');
        writeTable(mismatch, [
            ['Part No', 'Description', 'CL Material', 'CL Thickness', 'CL Qty', 'WS Material', 'WS Thickness', 'WS Qty', 'WS Source Row', 'Mismatch Type', 'Source Tabs', 'Family Key'],
            ...result.rows.filter(x => x.status === 'Mismatch').map(x => [x.partNoRaw, x.description, x.clMaterial, x.clThickness, x.clQty, x.wsMaterial, x.wsThickness ?? '', x.wsQty ?? '', x.wsSourceRow ?? '', x.mismatchType, x.sourceTabs, x.familyKey])
        ]);
        const add = context.workbook.worksheets.add('SN Add');
        writeTable(add, [
            ['Part No', 'Description', 'Material', 'Thickness', 'Qty', 'Source Tabs', 'Reason'],
            ...result.rows.filter(x => x.status === 'Add').map(x => [x.partNoRaw, x.description, x.clMaterial, x.clThickness, x.clQty, x.sourceTabs, x.mismatchType])
        ]);
        const remove = context.workbook.worksheets.add('SN Remove');
        writeTable(remove, [
            ['Remove Rank', 'Source Row', 'Part No', 'Material', 'Thickness', 'Qty', 'Reason'],
            ...result.removals.map(x => [x.removeRank, x.sourceRow, x.partNoRaw, x.materialCanonical, x.thicknessRaw ?? '', x.qtyRaw ?? '', x.reason])
        ]);
        const review = context.workbook.worksheets.add('SN Part Review');
        writeTable(review, [
            ['Part No (CL)', 'Description', 'CL Material', 'CL Thickness', 'CL Qty', 'Similar WS Part No', 'WS Material', 'WS Thickness', 'WS Qty', 'Reason', 'Family Key'],
            ...result.rows.filter(x => x.status === 'Part Review').map(x => [x.partNoRaw, x.description, x.clMaterial, x.clThickness, x.clQty, x.partReviewWsPartNo ?? '', x.partReviewWsMaterial ?? '', x.partReviewWsThickness ?? '', x.partReviewWsQty ?? '', x.mismatchType, x.familyKey])
        ]);
        const data = context.workbook.worksheets.add('SN Comparison Data');
        writeTable(data, [
            ['ReqIndex', 'PartNo_Norm', 'PartNo_Raw', 'Description', 'CL_Material', 'CL_Thickness', 'CL_Qty', 'SourceTabs', 'CL_ThicknessConflict', 'CL_MaterialConflict', 'WS_MatchFound', 'WS_RecordKey', 'WS_Material', 'WS_Thickness', 'WS_Qty', 'WS_SourceRow', 'WS_ThicknessConflict', 'WS_MaterialConflict', 'MaterialMatch', 'ThicknessMatch', 'QtyMatch', 'MismatchType', 'Status', 'FamilyKey'],
            ...result.rows.map(x => [x.reqIndex, x.partNoNorm, x.partNoRaw, x.description, x.clMaterial, x.clThickness, x.clQty, x.sourceTabs, x.clThicknessConflict, x.clMaterialConflict, x.wsMatchFound, x.wsRecordKey, x.wsMaterial, x.wsThickness ?? '', x.wsQty ?? '', x.wsSourceRow ?? '', x.wsThicknessConflict, x.wsMaterialConflict, x.materialMatch, x.thicknessMatch, x.qtyMatch, x.mismatchType, x.status, x.familyKey])
        ]);
        const audit = context.workbook.worksheets.add('SN Audit');
        writeTable(audit, [
            ['Date/Time', 'CL Workbook', 'WS Workbook', 'Job Number', 'Selected Tabs', 'CL Rows', 'WS Records', 'Exact Matches', 'Total Mismatches', 'Parts Add', 'Parts Remove', 'Part Review', 'Status', 'Notes'],
            [new Date(result.runDateTime).toLocaleString(), result.clFileName, result.wsFileName, result.jobNumber, result.selectedTabs.join(', '), result.validIncludedClRecords, result.validWsRecords, result.exactMatches, result.totalMismatches, result.addCount, result.removeCount, result.partReviewCount, 'READY FOR REVIEW', 'Comparison run by SigmaNEST Comparator Office Add-in']
        ]);
        summary.activate();
        await context.sync();
    });
}
