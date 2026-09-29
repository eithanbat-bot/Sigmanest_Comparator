function setCellValue(ws, r, c, value) {
    const addr = XLSX.utils.encode_cell({ r, c });
    if (value === undefined || value === null)
        delete ws[addr];
    else
        ws[addr] = { t: typeof value === 'number' ? 'n' : 's', v: value };
}
function findLastRow(ws) {
    const ref = ws['!ref'];
    if (!ref)
        return 0;
    return XLSX.utils.decode_range(ref).e.r;
}
export function buildRevisedWorkbook(parsed, result) {
    const wb = XLSX.read(parsed.sourceArrayBuffer, { type: 'array', raw: true });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    // Quantity-only updates: use original source row from WS_Staging behavior.
    for (const row of result.rows.filter(x => x.status === 'Mismatch' && x.materialMatch && x.thicknessMatch && !x.qtyMatch)) {
        if (row.wsSourceRow && parsed.qtyCol >= 0)
            setCellValue(ws, row.wsSourceRow - 1, parsed.qtyCol, row.clQty);
    }
    // Remove only the first WS row for each definitive key, matching the VBA logic.
    const removeRows = [...result.removals].map(x => x.sourceRow - 1).sort((a, b) => b - a);
    let matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    for (const zeroBased of removeRows) {
        if (zeroBased >= 0 && zeroBased < matrix.length)
            matrix.splice(zeroBased, 1);
    }
    // Rebuild matrix after removals so source-row updates have already been applied.
    // Add definitive missing parts to the bottom.
    for (const row of result.rows.filter(x => x.status === 'Add')) {
        const newRow = [];
        newRow[parsed.partCol] = row.partNoRaw;
        if (parsed.materialCol >= 0)
            newRow[parsed.materialCol] = row.clMaterial;
        if (parsed.thicknessCol >= 0)
            newRow[parsed.thicknessCol] = row.clThickness;
        if (parsed.qtyCol >= 0)
            newRow[parsed.qtyCol] = row.clQty;
        matrix.push(newRow);
    }
    const outWs = XLSX.utils.aoa_to_sheet(matrix);
    // Keep simple column widths where possible.
    if (range.e.c >= 0)
        outWs['!cols'] = Array.from({ length: range.e.c + 1 }, () => ({ wch: 14 }));
    wb.Sheets[sheetName] = outWs;
    return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
}
export function triggerDownload(buffer, filename) {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
}
export async function openRevisedWorkbookInExcel(buffer) {
    const bytes = new Uint8Array(buffer);
    let binary = '';
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk)
        binary += String.fromCharCode(...bytes.subarray(i, Math.min(i + chunk, bytes.length)));
    const base64 = btoa(binary);
    await Excel.createWorkbook(base64);
}
