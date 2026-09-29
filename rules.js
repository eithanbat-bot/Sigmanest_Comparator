// Seeded from the user's supplied workbook. The existing workbook has the
// special Armox/Ramor entries disabled, but the user's stated operating rule
// is that 4 mm Armox is treated as 4 mm Ramor 500, so this explicit 4 mm rule
// is enabled in the add-in.
export const MATERIAL_RULES = [
    { active: true, materialA: 'MS', materialB: 'MILD STEEL', reason: 'General' },
    { active: true, materialA: 'Perspex', materialB: 'BLACK PERSPEX', reason: 'General' },
    { active: true, materialA: 'MS', materialB: 'Mild Steel Sheet', reason: 'General' },
    { active: true, materialA: 'Armox 500', materialB: 'Armox', reason: 'General' },
    { active: true, materialA: 'Ramor 500', materialB: 'Armox 500', thicknessRule: 4, reason: 'Approved 4 mm substitution' },
    { active: true, materialA: 'ALTread', materialB: 'Al Tread 1200H4' },
    { active: true, materialA: 'Strenx 700', materialB: 'Strenx' },
    { active: true, materialA: 'Chromodeck', materialB: 'Chromodeck Sheet' },
    { active: true, materialA: 'Gr304 SS', materialB: 'SS' },
    { active: true, materialA: 'Ramor 500', materialB: 'Ramort' },
    { active: true, materialA: 'Ramor 500', materialB: 'Ramort 500' },
    { active: true, materialA: 'ALTread', materialB: 'Aluminium Tread' },
    { active: true, materialA: 'Hardox', materialB: 'Hardox 500' },
    { active: true, materialA: 'Ramor 500', materialB: 'Ramor® 500', reason: 'General' },
    { active: true, materialA: 'MS', materialB: 'Mild Steel Plate', reason: 'General' },
    { active: true, materialA: 'Aluminium', materialB: 'Aluminium Sheet', reason: 'General' },
    { active: true, materialA: 'ALTread', materialB: 'AL Chequer Plate', reason: 'General' },
    { active: true, materialA: 'Gr304 SS', materialB: 'GR 304 SS', reason: 'General' },
];
export const THICKNESS_RULES = [
    { active: true, materialFilter: 'Armox', fromThickness: 9, toThickness: 10, reason: 'Spec 10/31: 8.5mm Armox treated as 9mm' },
    { active: true, materialFilter: 'Ramor', fromThickness: 8.5, toThickness: 9, reason: 'Spec 10/31: 8.5mm Ramor treated as 9mm' },
    { active: true, fromThickness: 3.5, toThickness: 4, reason: 'General half-mm normalization' },
];
