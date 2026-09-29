import { APP_NAME, DEV_MODE, VERSION, defaultLicense } from './config.js';
import { parseClFile, parseWsFile } from './parser.js';
import { compare } from './comparisonEngine.js';
import { MATERIAL_RULES, THICKNESS_RULES, cloneDefaultRules } from './rules.js';
import { checkLicense, activateLicense } from './license.js';
import { loadState, persistCl, persistWs, persistSelectionAndResult, persistRules } from './state.js';
import { writeReports } from './excelWriter.js';
import { buildRevisedWorkbook, triggerDownload, openRevisedWorkbookInExcel } from './revisedWs.js';
import { detectJobNumber } from './textRules.js';
let clData;
let wsData;
let lastResult;
let license = defaultLicense;
let materialRules = cloneDefaultRules().materialRules;
let thicknessRules = cloneDefaultRules().thicknessRules;
let activeRuleTab = 'material';
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
    if (Array.isArray(s.materialRules)) materialRules = s.materialRules;
    if (Array.isArray(s.thicknessRules)) thicknessRules = s.thicknessRules;
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
        lastResult = compare(clData.records, tabs, wsData.records, clData.fileName, wsData.fileName, job, materialRules, thicknessRules);
        await persistSelectionAndResult(tabs, lastResult);
        await writeReports(lastResult, { materialRules, thicknessRules });
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

function renderRuleTables() {
 const m=el('rules-material'), t=el('rules-thickness'); m.innerHTML=''; t.innerHTML='';
 const mt=document.createElement('table'); mt.className='rule-table'; mt.innerHTML='<thead><tr><th>Active</th><th>Material A</th><th>Material B</th><th>Thickness</th><th>Reason / Notes</th><th></th></tr></thead><tbody></tbody>'; const mb=mt.querySelector('tbody');
 materialRules.forEach((r,i)=>{const tr=document.createElement('tr'); tr.innerHTML=`<td><input type="checkbox" data-field="active" ${r.active?'checked':''}></td><td><input data-field="materialA" value="${esc(r.materialA)}"></td><td><input data-field="materialB" value="${esc(r.materialB)}"></td><td><input data-field="thicknessRule" type="number" step="0.1" value="${r.thicknessRule??''}"></td><td><input data-field="reason" value="${esc(r.reason||'')}"></td><td><button class="btn small danger" data-delete="${i}">×</button></td>`; tr.querySelectorAll('[data-field]').forEach(inp=>inp.addEventListener('change',()=>updateMaterialRule(i,tr))); tr.querySelector('[data-delete]').addEventListener('click',()=>{materialRules.splice(i,1);renderRuleTables()}); mb.appendChild(tr)}); m.appendChild(mt);
 const tt=document.createElement('table'); tt.className='rule-table'; tt.innerHTML='<thead><tr><th>Active</th><th>Material Filter</th><th>From</th><th>To</th><th>Reason / Notes</th><th></th></tr></thead><tbody></tbody>'; const tb=tt.querySelector('tbody');
 thicknessRules.forEach((r,i)=>{const tr=document.createElement('tr'); tr.innerHTML=`<td><input type="checkbox" data-field="active" ${r.active?'checked':''}></td><td><input data-field="materialFilter" value="${esc(r.materialFilter||'')}"></td><td><input data-field="fromThickness" type="number" step="0.1" value="${r.fromThickness??''}"></td><td><input data-field="toThickness" type="number" step="0.1" value="${r.toThickness??''}"></td><td><input data-field="reason" value="${esc(r.reason||'')}"></td><td><button class="btn small danger" data-delete="${i}">×</button></td>`; tr.querySelectorAll('[data-field]').forEach(inp=>inp.addEventListener('change',()=>updateThicknessRule(i,tr))); tr.querySelector('[data-delete]').addEventListener('click',()=>{thicknessRules.splice(i,1);renderRuleTables()}); tb.appendChild(tr)}); t.appendChild(tt);
}
function esc(v){return String(v??'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}
function updateMaterialRule(i,tr){const q=f=>tr.querySelector(`[data-field="${f}"]`); materialRules[i]={...materialRules[i],active:q('active').checked,materialA:q('materialA').value.trim(),materialB:q('materialB').value.trim(),thicknessRule:q('thicknessRule').value===''?undefined:Number(q('thicknessRule').value),reason:q('reason').value.trim()};}
function updateThicknessRule(i,tr){const q=f=>tr.querySelector(`[data-field="${f}"]`); thicknessRules[i]={...thicknessRules[i],active:q('active').checked,materialFilter:q('materialFilter').value.trim(),fromThickness:q('fromThickness').value===''?undefined:Number(q('fromThickness').value),toThickness:q('toThickness').value===''?undefined:Number(q('toThickness').value),reason:q('reason').value.trim()};}
async function saveRules(){await persistRules(materialRules,thicknessRules);setStatus('Comparison rules saved. They will be used on the next run.','ok');}
function switchRuleTab(tab){activeRuleTab=tab;el('rules-material').classList.toggle('hidden',tab!=='material');el('rules-thickness').classList.toggle('hidden',tab!=='thickness');el('rule-tab-material').classList.toggle('active',tab==='material');el('rule-tab-thickness').classList.toggle('active',tab==='thickness');el('btn-add-rule').textContent=tab==='material'?'+ Add Material Rule':'+ Add Thickness Rule';}
function addRule(){if(activeRuleTab==='material') materialRules.push({active:true,materialA:'',materialB:'',reason:''}); else thicknessRules.push({active:true,materialFilter:'',fromThickness:'',toThickness:'',reason:''});renderRuleTables();}
async function resetRules(){const d=cloneDefaultRules();materialRules=d.materialRules;thicknessRules=d.thicknessRules;renderRuleTables();await persistRules(materialRules,thicknessRules);setStatus('Rules reset to the original comparator defaults.','ok');}

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
    el('rule-tab-material').addEventListener('click',()=>switchRuleTab('material'));
    el('rule-tab-thickness').addEventListener('click',()=>switchRuleTab('thickness'));
    el('btn-add-rule').addEventListener('click',addRule);
    el('btn-save-rules').addEventListener('click',()=>void saveRules());
    el('btn-reset-rules').addEventListener('click',()=>void resetRules());
    renderRuleTables();
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
