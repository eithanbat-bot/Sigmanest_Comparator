import { STATE_KEY } from './config.js';
let memory = { selectedTabs: [], license: { active: false } };
function stripCl(cl) {
    const { sourceArrayBuffer: _buffer, ...rest } = cl;
    return rest;
}
function stripWs(ws) {
    const { sourceArrayBuffer: _buffer, ...rest } = ws;
    return rest;
}
export async function loadState() {
    try {
        if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
            const raw = await OfficeRuntime.storage.getItem(STATE_KEY);
            if (raw)
                memory = { ...memory, ...JSON.parse(raw) };
        }
        else {
            const raw = localStorage.getItem(STATE_KEY);
            if (raw)
                memory = { ...memory, ...JSON.parse(raw) };
        }
    }
    catch { /* use memory */ }
    return memory;
}
export async function saveState(state) {
    memory = state;
    const json = JSON.stringify(state);
    try {
        if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage)
            await OfficeRuntime.storage.setItem(STATE_KEY, json);
        else
            localStorage.setItem(STATE_KEY, json);
    }
    catch {
        try {
            localStorage.setItem(STATE_KEY, json);
        }
        catch { /* best effort */ }
    }
}
export async function persistSelectionAndResult(selectedTabs, result) {
    const current = await loadState();
    current.selectedTabs = selectedTabs;
    if (result)
        current.result = result;
    await saveState(current);
}
export async function persistCl(cl) {
    const current = await loadState();
    current.cl = stripCl(cl);
    await saveState(current);
}
export async function persistWs(ws) {
    const current = await loadState();
    current.ws = stripWs(ws);
    await saveState(current);
}
export async function persistLicense(license) {
    const current = await loadState();
    current.license = license;
    await saveState(current);
}
