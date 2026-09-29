import { normHeader, parseNumber, parseThickness, canonicalizeMaterial } from './textRules.js';
function asMatrix(workbook, sheetName) {
    const ws = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
}
function findHeader(rows, maxRows, required, alternateAny = []) {
    const limit = Math.min(maxRows, rows.length);
    for (let r = 0; r < limit; r++) {
        const headers = (rows[r] || []).map(normHeader);
        const hasRequired = required.every(x => headers.includes(x));
        const hasAny = alternateAny.length === 0 || alternateAny.some(x => headers.includes(x));
        if (!hasRequired || !hasAny)
            continue;
        const cols = {};
        for (const h of required)
            cols[h] = headers.indexOf(h);
        if (alternateAny.length) {
            const found = alternateAny.find(x => headers.includes(x));
            if (found)
                cols[found] = headers.indexOf(found);
        }
        return { row: r, cols };
    }
    return undefined;
}
export async function parseClFile(file) {
    const sourceArrayBuffer = await file.arrayBuffer();
    const wb = XLSX.read(sourceArrayBuffer, { type: 'array', raw: true });
    const tabs = wb.SheetNames;
    const records = [];
    const tabMeta = {};
    for (const sheet of tabs) {
        const rows = asMatrix(wb, sheet);
        const header = findHeader(rows, 40, ['part no', 'material'], ['qty', 'quantity']);
        if (!header) {
            tabMeta[sheet] = { detectedPartRows: 0, status: 'NO DETECTABLE PART TABLE' };
            continue;
        }
        const qtyCol = header.cols.qty ?? header.cols.quantity;
        const partCol = header.cols['part no'];
        const matCol = header.cols.material;
        const descCol = (rows[header.row] || []).map(normHeader).indexOf('part description');
        let detectedPartRows = 0;
        for (let r = header.row + 1; r < rows.length; r++) {
            const row = rows[r] || [];
            const part = row[partCol];
            const mat = row[matCol];
            const qty = parseNumber(row[qtyCol]);
            if (part === null || part === undefined || String(part).trim() === '')
                continue;
            if (mat === null || mat === undefined || String(mat).trim() === '' || qty === undefined)
                continue;
            const parsed = parseThickness(mat);
            records.push({
                sheet,
                sourceRow: r + 1,
                partNoRaw: String(part).trim(),
                description: descCol >= 0 ? String(row[descCol] ?? '').trim() : '',
                qty,
                materialRaw: String(mat).trim(),
                thicknessParsed: parsed.thickness,
                materialCanonical: parsed.materialCanonical,
                thicknessParseError: parsed.thickness === undefined,
            });
            detectedPartRows++;
        }
        tabMeta[sheet] = { headerRow: header.row + 1, detectedPartRows, status: 'DETECTED' };
    }
    return { fileName: file.name, tabs, records, tabMeta, sourceArrayBuffer };
}
export async function parseWsFile(file) {
    const sourceArrayBuffer = await file.arrayBuffer();
    const wb = XLSX.read(sourceArrayBuffer, { type: 'array', raw: true });
    const firstSheet = wb.SheetNames[0];
    if (!firstSheet)
        throw new Error('The WS workbook has no worksheets.');
    const rows = asMatrix(wb, firstSheet);
    const header = findHeader(rows, 10, ['part name', 'material'], ['qty', 'quantity']);
    if (!header)
        throw new Error("Could not find a WS header containing 'Part Name', 'Material', and 'Qty/Quantity' in the first 10 rows.");
    const headers = (rows[header.row] || []).map(normHeader);
    const partCol = headers.indexOf('part name');
    const materialCol = headers.indexOf('material');
    const thicknessCol = headers.indexOf('thickness');
    const qtyCol = headers.includes('qty') ? headers.indexOf('qty') : headers.indexOf('quantity');
    const records = [];
    for (let r = header.row + 1; r < rows.length; r++) {
        const row = rows[r] || [];
        const part = row[partCol];
        if (part === null || part === undefined || String(part).trim() === '')
            continue;
        const mat = materialCol >= 0 ? row[materialCol] : '';
        const thk = thicknessCol >= 0 ? parseNumber(row[thicknessCol]) : undefined;
        const qty = qtyCol >= 0 ? parseNumber(row[qtyCol]) : undefined;
        records.push({
            sourceRow: r + 1,
            partNoRaw: String(part).trim(),
            materialRaw: String(mat ?? '').trim(),
            thicknessRaw: thk,
            qtyRaw: qty,
            materialCanonical: canonicalizeMaterial(mat ?? ''),
        });
    }
    return {
        fileName: file.name,
        records,
        headerRow: header.row + 1,
        partCol,
        materialCol,
        thicknessCol,
        qtyCol,
        sourceArrayBuffer,
    };
}
