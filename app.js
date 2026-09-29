import { APP_NAME, DEV_MODE, VERSION, defaultLicense } from './config.js';
import { parseClFile, parseWsFile } from './parser.js';
import { compare } from './comparisonEngine.js';
import { MATERIAL_RULES, THICKNESS_RULES } from './rules.js';
import { checkLicense, activateLicense } from './license.js';
import { loadState, persistCl, persistWs, persistSelectionAndResult } from './state.js';
import { writeReports } from './excelWriter.js';
import { buildRevisedWorkbook, triggerDownload, openRevisedWorkbookInExcel } from './revisedWs.js';
import { detectJobNumber } from './textRules.js';
let clData;
let wsData;
let lastResult;
let license = defaultLicense;
function el(id) { return document.getElementById(id); }
function setStatus(text, kind = 'info') {
    const node = el('status');
    node.textContent = text;
    node.dataset.kind = kind;
}
function setStep(id, text) { el(id).textContent = text; }
function countBadge(id, n) { el(id).textContent = String(n); }
function renderTabs() {
    const container = el('tab-list');
    container.innerHTML = '';
    if (!clData)
        return;
    const defaultSelected = new Set(clData.tabs.filter(t => /^BAT/i.test(t)));
    const previous = new Set((lastResult?.selectedTabs ?? []).length ? lastResult.selectedTabs : [...defaultSelected]);
    for (const tab of clData.tabs) {
        const meta = clData.tabMeta[tab];
        const row = document.createElement('label');
        row.className = 'tab-row';
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = previous.has(tab);
        checkbox.dataset.sheet = tab;
        checkbox.addEventListener('change', () => updateSelectedCount());
        const title = document.createElement('span');
        title.textContent = tab;
        const badge = document.createElement('span');
        badge.className = `pill ${meta?.status === 'DETECTED' ? 'pill-ok' : 'pill-warn'}`;
        badge.textContent = meta?.status === 'DETECTED' ? `${meta.detectedPartRows} rows` : 'No table';
        row.append(checkbox, title, badge);
        container.appendChild(row);
    }
    updateSelectedCount();
}
function selectedTabs() {
    return Array.from(el('tab-list').querySelectorAll('input[type=checkbox]:checked')).map(c => c.dataset.sheet || '').filter(Boolean);
}
function updateSelectedCount() { countBadge('selected-count', selectedTabs().length); }
async function requireLicense() {
    license = await checkLicense();
    el('license-status').textContent = license.active ? `${license.plan || 'ACTIVE'}${license.expiresAt ? ` · Expires ${new Date(license.expiresAt).toLocaleDateString()}` : ''}` : 'INACTIVE';
    el('license-status').className = `license ${license.active ? 'active' : 'inactive'}`;
    return license.active;
}
export async function chooseCl() {
    if (!(await requireLicense()))
        return setStatus('Subscription is not active.', 'error');
    el('cl-file').click();
}
export async function chooseWs() {
    if (!(await requireLicense()))
        return setStatus('Subscription is not active.', 'error');
    el('ws-file').click();
}
async function onClFile(file) {
    if (!file)
        return;
    setStatus('Reading CL workbook…');
    try {
        clData = await parseClFile(file);
        await persistCl(clData);
        setStep('cl-name', file.name);
        setStep('cl-meta', `${clData.tabs.length} sheets · ${clData.records.length} parsed rows`);
        renderTabs();
        setStatus('CL workbook loaded. Choose the tabs to include.', 'ok');
    }
    catch (e) {
        setStatus(e.message, 'error');
    }
}
async function onWsFile(file) {
    if (!file)
        return;
    setStatus('Reading WS workbook…');
    try {
        wsData = await parseWsFile(file);
        await persistWs(wsData);
        setStep('ws-name', file.name);
        setStep('ws-meta', `${wsData.records.length} parsed records · header row ${wsData.headerRow}`);
        setStatus('WS workbook loaded.', 'ok');
    }
    catch (e) {
        setStatus(e.message, 'error');
    }
}
async function ensureLoadedFromState() {
    const s = await loadState();
    // File bytes are intentionally not persisted in this MVP. The add-in requires selecting files once per Excel session.
    lastResult = s.result;
    license = s.license || defaultLicense;
}
export async function runComparison() {
    if (!(await requireLicense()))
        return setStatus('Subscription is not active.', 'error');
    if (!clData || !wsData)
        return setStatus('Select both the CL and WS workbook first.', 'error');
    const tabs = selectedTabs();
    if (!tabs.length)
        return setStatus('Select at least one CL tab.', 'error');
    setStatus('Running comparison…');
    try {
        const job = detectJobNumber(`${clData.fileName} ${wsData.fileName}`);
        lastResult = compare(clData.records, tabs, wsData.records, clData.fileName, wsData.fileName, job, MATERIAL_RULES, THICKNESS_RULES);
        await persistSelectionAndResult(tabs, lastResult);
        await writeReports(lastResult);
        renderResults(lastResult);
        setStatus('Comparison complete. See the SN Summary and result sheets.', 'ok');
    }
    catch (e) {
        setStatus(e.message, 'error');
    }
}
function renderResults(r) {
    const grid = el('result-grid');
    grid.innerHTML = '';
    const metrics = [
        ['Exact', r.exactMatches], ['Mismatches', r.totalMismatches], ['Add', r.addCount], ['Remove', r.removeCount], ['Review', r.partReviewCount]
    ];
    for (const [label, n] of metrics) {
        const card = document.createElement('div');
        card.className = 'metric';
        const v = document.createElement('strong');
        v.textContent = String(n);
        const t = document.createElement('span');
        t.textContent = label;
        card.append(v, t);
        grid.appendChild(card);
    }
    countBadge('mismatch-count', r.totalMismatches);
}
export async function generateRevisedWs() {
    if (!(await requireLicense()))
        return setStatus('Subscription is not active.', 'error');
    if (!lastResult || !wsData)
        return setStatus('Run a comparison first.', 'error');
    setStatus('Building safe revised WS copy…');
    try {
        const buffer = buildRevisedWorkbook(wsData, lastResult);
        const baseName = wsData.fileName.replace(/\.[^.]+$/, '');
        const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}Z$/, 'Z').replace('T', '_');
        const filename = `${baseName}_REVISED_${stamp}.xlsx`;
        triggerDownload(buffer, filename);
        el('revised-name').textContent = filename;
        const open = confirm(`Revised WS created as ${filename}.\n\nOK = also open it as a new Excel workbook.\nCancel = keep the download only.`);
        if (open)
            await openRevisedWorkbookInExcel(buffer);
        setStatus(`Revised WS created: ${filename}. Material/thickness mismatches and Part Review items were left untouched.`, 'ok');
    }
    catch (e) {
        setStatus(e.message, 'error');
    }
}
export async function refreshInputs() {
    if (!(await requireLicense()))
        return setStatus('Subscription is not active.', 'error');
    if (!clData && !wsData)
        return setStatus('Select the CL and WS files first.', 'error');
    setStatus('Re-reading current session inputs…');
    try {
        if (clData) {
            const freshCl = new File([clData.sourceArrayBuffer], clData.fileName);
            await onClFile(freshCl);
        }
        if (wsData) {
            const freshWs = new File([wsData.sourceArrayBuffer], wsData.fileName);
            await onWsFile(freshWs);
        }
        setStatus('Inputs refreshed.', 'ok');
    }
    catch (e) {
        setStatus(e.message, 'error');
    }
}
export async function openLicenseDialog() {
    const code = prompt('Enter your license / activation code:');
    if (code === null)
        return;
    if (DEV_MODE) {
        license = defaultLicense;
        setStatus('Development mode: test license is active.', 'ok');
        return;
    }
    const state = await activateLicense(code.trim());
    license = state;
    await requireLicense();
    setStatus(state.active ? 'License activated.' : (state.message || 'License activation failed.'), state.active ? 'ok' : 'error');
}
export function initUi() {
    el('version').textContent = `v${VERSION}`;
    el('app-title').textContent = APP_NAME;
    el('cl-file').addEventListener('change', e => onClFile(e.target.files?.[0]));
    el('ws-file').addEventListener('change', e => onWsFile(e.target.files?.[0]));
    el('btn-cl').addEventListener('click', () => void chooseCl());
    el('btn-ws').addEventListener('click', () => void chooseWs());
    el('btn-run').addEventListener('click', () => void runComparison());
    el('btn-revised').addEventListener('click', () => void generateRevisedWs());
    el('btn-refresh').addEventListener('click', () => void refreshInputs());
    el('btn-license').addEventListener('click', () => void openLicenseDialog());
    el('btn-select-all').addEventListener('click', () => { el('tab-list').querySelectorAll('input').forEach(x => x.checked = true); updateSelectedCount(); });
    el('btn-select-bat').addEventListener('click', () => { el('tab-list').querySelectorAll('input').forEach(x => x.checked = /^BAT/i.test(x.dataset.sheet || '')); updateSelectedCount(); });
    el('btn-clear-all').addEventListener('click', () => { el('tab-list').querySelectorAll('input').forEach(x => x.checked = false); updateSelectedCount(); });
    void ensureLoadedFromState().then(() => void requireLicense());
}
// Command entry points used by the Ribbon.
export async function commandChooseCl(event) { await chooseCl(); event.completed(); }
export async function commandChooseWs(event) { await chooseWs(); event.completed(); }
export async function commandRunComparison(event) { await runComparison(); event.completed(); }
export async function commandGenerateRevisedWs(event) { await generateRevisedWs(); event.completed(); }
export async function commandShowTaskPane(event) { event.completed(); }
