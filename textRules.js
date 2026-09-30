export function clean(value) {
  return [...value]
    .filter(ch => ch === '\t' || ch === '\n' || ch === '\r' || ch.charCodeAt(0) >= 32)
    .join('');
}

export function normalizePartNumber(raw) {
  if (raw === null || raw === undefined) return '';
  const s = clean(String(raw)).replace(/\u00a0/g, ' ').trim();
  return s.replace(/\s+/g, '').replace(/-/g, '').replace(/[\u2013\u2014]/g, '').toUpperCase();
}

export function familyKey(normalizedPartNumber) {
  const s = normalizedPartNumber || '';
  if (s.length < 2) return s;
  const last = s[s.length - 1];
  const prev = s[s.length - 2];
  const isLetter = last >= 'A' && last <= 'Z';
  const prevIsDigit = /[0-9]/.test(prev);
  return isLetter && prevIsDigit ? s.slice(0, -1) : s;
}

export function canonicalizeMaterial(raw) {
  const input = raw === null || raw === undefined ? '' : String(raw).trim();
  return input
    .replace(/armoxt/gi, 'Armox')
    .replace(/\s+/g, ' ')
    .trim();
}

export function normHeader(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim().toLowerCase().split(/\s+/).filter(Boolean).join(' ');
}

export function parseNumber(value) {
  if (value === null || value === undefined || String(value).trim() === '') return undefined;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const t = String(value).trim().replace(/,/g, '.');
  const n = Number(t);
  return Number.isFinite(n) ? n : undefined;
}

export function parseThickness(raw) {
  const txt = raw === null || raw === undefined ? '' : String(raw).trim();
  const mmPos = txt.toLowerCase().indexOf('mm');
  if (mmPos < 0) return { materialCanonical: canonicalizeMaterial(txt) };
  const before = txt.slice(0, mmPos).trimEnd();
  let p = before.length - 1;
  while (p >= 0 && /[0-9.,]/.test(before[p])) p--;
  const token = before.slice(p + 1);
  const prefix = before.slice(0, p + 1);
  const remainder = (prefix + ' ' + txt.slice(mmPos + 2)).trim().split(/\s+/).filter(Boolean).join(' ');
  const parsed = Number(token.replace(/,/g, '.'));
  if (!Number.isFinite(parsed)) return { materialCanonical: canonicalizeMaterial(remainder) };
  return { thickness: parsed, materialCanonical: canonicalizeMaterial(remainder) };
}

export function detectJobNumber(text) {
  const m = String(text || '').match(/\b(?:JO?|JOB)?(\d{2,4})[-/](\d{2})\b/i);
  if (!m) return '';
  return 'J' + String(m[1]).padStart(3, '0') + '/' + m[2];
}
