import { familyKey, normalizePartNumber } from './textRules.js';
function eq(a, b) {
    return (a ?? '').toLocaleLowerCase() === (b ?? '').toLocaleLowerCase();
}
function applyThicknessRules(original, material, rules) {
    if (original === undefined)
        return undefined;
    for (const rule of rules.filter(r => r.active)) {
        if (Math.abs(rule.fromThickness - original) > 1e-6)
            continue;
        if (!rule.materialFilter || material.toLocaleLowerCase().includes(rule.materialFilter.toLocaleLowerCase()))
            return rule.toThickness;
    }
    return original;
}
function applyMaterialRules(original, thickness, rules) {
    for (const rule of rules.filter(r => r.active)) {
        const matchesMaterial = eq(rule.materialA, original) || eq(rule.materialB, original);
        const matchesThickness = rule.thicknessRule === undefined || (thickness !== undefined && Math.abs(rule.thicknessRule - thickness) < 1e-6);
        if (matchesMaterial && matchesThickness)
            return rule.materialA;
    }
    return original;
}
function key(part, thickness, material) {
    const t = thickness === undefined ? '' : String(thickness);
    return `${part}|${t}|${material.toUpperCase()}`;
}
function mismatchType(r) {
    const parts = [];
    if (!r.materialMatch)
        parts.push('Material');
    if (!r.thicknessMatch)
        parts.push('Thickness');
    if (!r.qtyMatch)
        parts.push('Quantity');
    return `${parts.join(' + ')} mismatch`;
}
export function compare(allClRecords, selectedTabs, wsRecords, clFileName, wsFileName, jobNumber, materialRules, thicknessRules) {
    if (!selectedTabs.length)
        throw new Error('No CL tabs are selected. Select at least one tab before running the comparison.');
    const selectedSet = new Set(selectedTabs.map(x => x.toLowerCase()));
    const clStaged = allClRecords.map(source => {
        const include = selectedSet.has((source.sheet || '').toLowerCase());
        const effectiveThickness = applyThicknessRules(source.thicknessParsed, source.materialCanonical, thicknessRules);
        const effectiveMaterial = applyMaterialRules(source.materialCanonical, effectiveThickness, materialRules);
        const valid = include && source.qty > 0 && !!source.partNoRaw.trim() && !!effectiveMaterial.trim() && !source.thicknessParseError && source.thicknessParsed !== undefined;
        const partNoNorm = normalizePartNumber(source.partNoRaw);
        const recordKey = valid ? key(partNoNorm, effectiveThickness, effectiveMaterial) : '';
        return {
            source,
            include,
            effectiveThickness,
            effectiveMaterial,
            valid,
            partNoNorm,
            recordKey,
            familyKey: familyKey(partNoNorm),
            isFirstOfKey: false,
            aggQty: 0,
            reqRank: undefined,
            thicknessConflict: false,
            materialConflict: false,
        };
    });
    const validCl = clStaged.filter(x => x.valid);
    const clByPart = new Map();
    for (const x of validCl) {
        const arr = clByPart.get(x.partNoNorm) ?? [];
        arr.push(x);
        clByPart.set(x.partNoNorm, arr);
    }
    for (const x of validCl) {
        const group = clByPart.get(x.partNoNorm);
        x.thicknessConflict = group.some(y => Math.abs((y.effectiveThickness ?? 0) - (x.effectiveThickness ?? 0)) > 1e-6);
        x.materialConflict = group.some(y => !eq(y.effectiveMaterial, x.effectiveMaterial));
    }
    const seenClKeys = new Set();
    let rank = 0;
    for (const x of clStaged) {
        if (!x.valid)
            continue;
        x.isFirstOfKey = !seenClKeys.has(x.recordKey);
        if (!x.isFirstOfKey)
            continue;
        seenClKeys.add(x.recordKey);
        rank++;
        x.reqRank = rank;
        x.aggQty = validCl.filter(y => y.recordKey.toLowerCase() === x.recordKey.toLowerCase()).reduce((sum, y) => sum + y.source.qty, 0);
    }
    const wsStaged = wsRecords.map(source => {
        const effectiveThickness = applyThicknessRules(source.thicknessRaw, source.materialCanonical, thicknessRules);
        const effectiveMaterial = applyMaterialRules(source.materialCanonical, effectiveThickness, materialRules);
        const valid = !!source.partNoRaw.trim() && (source.qtyRaw ?? 0) > 0 && !!effectiveMaterial.trim();
        const partNoNorm = normalizePartNumber(source.partNoRaw);
        const recordKey = valid ? key(partNoNorm, effectiveThickness, effectiveMaterial) : '';
        return {
            source,
            effectiveThickness,
            effectiveMaterial,
            valid,
            partNoNorm,
            recordKey,
            familyKey: familyKey(partNoNorm),
            isFirstOfKey: false,
            aggQty: 0,
            thicknessConflict: false,
            materialConflict: false,
            clPresent: false,
        };
    });
    const validWs = wsStaged.filter(x => x.valid);
    const wsByPart = new Map();
    const wsByFamily = new Map();
    const clPartSet = new Set(validCl.map(x => x.partNoNorm.toLowerCase()));
    for (const x of validWs) {
        const a = wsByPart.get(x.partNoNorm.toLowerCase()) ?? [];
        a.push(x);
        wsByPart.set(x.partNoNorm.toLowerCase(), a);
        const f = x.familyKey.toLowerCase();
        const b = wsByFamily.get(f) ?? [];
        b.push(x);
        wsByFamily.set(f, b);
    }
    for (const x of validWs) {
        const group = wsByPart.get(x.partNoNorm.toLowerCase());
        x.thicknessConflict = group.some(y => Math.abs((y.effectiveThickness ?? 0) - (x.effectiveThickness ?? 0)) > 1e-6);
        x.materialConflict = group.some(y => !eq(y.effectiveMaterial, x.effectiveMaterial));
        x.clPresent = clPartSet.has(x.partNoNorm.toLowerCase());
    }
    const seenWsKeys = new Set();
    for (const x of validWs) {
        x.isFirstOfKey = !seenWsKeys.has(x.recordKey);
        if (x.isFirstOfKey) {
            seenWsKeys.add(x.recordKey);
            x.aggQty = validWs.filter(y => y.recordKey.toLowerCase() === x.recordKey.toLowerCase()).reduce((sum, y) => sum + (y.source.qtyRaw ?? 0), 0);
        }
    }
    const removals = validWs
        .filter(x => x.isFirstOfKey && !x.clPresent)
        .sort((a, b) => a.source.sourceRow - b.source.sourceRow)
        .map((x, i) => ({
        sourceRow: x.source.sourceRow,
        partNoRaw: x.source.partNoRaw,
        materialRaw: x.source.materialRaw,
        thicknessRaw: x.source.thicknessRaw,
        qtyRaw: x.aggQty,
        materialCanonical: x.effectiveMaterial,
        reason: 'Part not present in selected CL',
        removeRank: i + 1,
    }));
    const requirements = clStaged.filter(x => x.valid && x.isFirstOfKey).sort((a, b) => (a.reqRank ?? 0) - (b.reqRank ?? 0));
    const rows = [];
    let exactMatches = 0, qtyOnly = 0, materialMismatch = 0, thicknessMismatch = 0, totalMismatch = 0, addCount = 0, reviewCount = 0;
    for (const req of requirements) {
        const row = {
            reqIndex: req.reqRank,
            partNoNorm: req.partNoNorm,
            partNoRaw: req.source.partNoRaw,
            description: req.source.description,
            clMaterial: req.effectiveMaterial,
            clThickness: req.effectiveThickness ?? 0,
            clQty: req.aggQty,
            sourceTabs: [...new Set(clStaged.filter(x => x.valid && x.recordKey.toLowerCase() === req.recordKey.toLowerCase()).map(x => x.source.sheet))].join(', '),
            clThicknessConflict: req.thicknessConflict,
            clMaterialConflict: req.materialConflict,
            wsMatchFound: false,
            wsRecordKey: '',
            wsMaterial: '',
            wsThickness: undefined,
            wsQty: undefined,
            wsSourceRow: undefined,
            wsThicknessConflict: false,
            wsMaterialConflict: false,
            materialMatch: false,
            thicknessMatch: false,
            qtyMatch: false,
            mismatchType: '',
            status: 'Add',
            familyKey: req.familyKey,
        };
        const exactCandidates = wsByPart.get(req.partNoNorm.toLowerCase()) ?? [];
        if (exactCandidates.length) {
            const first = exactCandidates[0];
            row.wsMatchFound = true;
            row.wsRecordKey = first.recordKey;
            row.wsMaterial = first.effectiveMaterial;
            row.wsThickness = first.effectiveThickness;
            row.wsQty = first.aggQty;
            row.wsSourceRow = first.source.sourceRow;
            row.wsThicknessConflict = exactCandidates.some(x => x.thicknessConflict);
            row.wsMaterialConflict = exactCandidates.some(x => x.materialConflict);
            row.materialMatch = eq(row.clMaterial, row.wsMaterial);
            row.thicknessMatch = row.wsThickness !== undefined && Math.abs(row.clThickness - row.wsThickness) < 1e-6;
            row.qtyMatch = row.wsQty !== undefined && Math.abs(row.clQty - row.wsQty) < 1e-6;
            if (row.materialMatch && row.thicknessMatch && row.qtyMatch) {
                row.status = 'Exact Match';
                row.mismatchType = 'Exact Match';
                exactMatches++;
            }
            else {
                row.status = 'Mismatch';
                row.mismatchType = mismatchType(row);
                totalMismatch++;
                if (!row.materialMatch)
                    materialMismatch++;
                if (!row.thicknessMatch)
                    thicknessMismatch++;
                if (row.materialMatch && row.thicknessMatch && !row.qtyMatch)
                    qtyOnly++;
            }
        }
        else {
            const familyCandidates = wsByFamily.get(req.familyKey.toLowerCase()) ?? [];
            if (familyCandidates.length) {
                const family = familyCandidates[0];
                row.status = 'Part Review';
                row.mismatchType = 'Similar part number - not an exact normalized match';
                row.partReviewWsPartNo = family.source.partNoRaw;
                row.partReviewWsMaterial = family.effectiveMaterial;
                row.partReviewWsThickness = family.effectiveThickness;
                row.partReviewWsQty = family.source.qtyRaw;
                reviewCount++;
            }
            else {
                row.status = 'Add';
                row.mismatchType = 'Not found in WS';
                addCount++;
            }
        }
        rows.push(row);
    }
    return {
        rows, removals,
        selectedTabs,
        clRecords: allClRecords,
        wsRecords,
        runDateTime: new Date().toISOString(),
        clFileName,
        wsFileName,
        jobNumber,
        validIncludedClRecords: clStaged.filter(x => x.include && x.valid).length,
        uniqueClRequirements: requirements.length,
        validWsRecords: validWs.length,
        exactMatches,
        quantityOnlyMismatches: qtyOnly,
        materialMismatches: materialMismatch,
        thicknessMismatches: thicknessMismatch,
        totalMismatches: totalMismatch,
        addCount,
        removeCount: removals.length,
        partReviewCount: reviewCount,
        dataConflicts: clStaged.filter(x => x.valid && x.materialConflict).length + clStaged.filter(x => x.valid && x.thicknessConflict).length,
        invalidIncludedClRecords: Math.max(0, clStaged.filter(x => x.include && !!x.source.materialCanonical && !x.source.thicknessParseError).length - clStaged.filter(x => x.include && x.valid).length),
    };
}
