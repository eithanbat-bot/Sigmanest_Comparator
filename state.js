import { STATE_KEY } from './config.js';

let memory = {
  selectedTabs: [],
  license: { active: false },
  materialRules: null,
  thicknessRules: null,
  settings: null,
  result: null,
};

function stripBuffer(obj) {
  if (!obj || typeof obj !== 'object') return obj;
  const { sourceArrayBuffer: _buffer, ...rest } = obj;
  return rest;
}

async function readPersisted() {
  try {
    if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
      return await OfficeRuntime.storage.getItem(STATE_KEY);
    }
    return localStorage.getItem(STATE_KEY);
  } catch {
    return null;
  }
}

async function writePersisted(json) {
  try {
    if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
      await OfficeRuntime.storage.setItem(STATE_KEY, json);
      return;
    }
    localStorage.setItem(STATE_KEY, json);
  } catch {
    try { localStorage.setItem(STATE_KEY, json); } catch { /* best effort */ }
  }
}

export async function loadState() {
  const raw = await readPersisted();
  if (raw) {
    try { memory = { ...memory, ...JSON.parse(raw) }; } catch { /* retain memory */ }
  }
  return memory;
}

export async function saveState(state) {
  memory = state;
  await writePersisted(JSON.stringify(state));
}

export async function persistSelectionAndResult(selectedTabs, result) {
  const current = await loadState();
  current.selectedTabs = [...selectedTabs];
  if (result) current.result = result;
  await saveState(current);
}

export async function persistCl(cl) {
  const current = await loadState();
  current.cl = stripBuffer(cl);
  await saveState(current);
}

export async function persistWs(ws) {
  const current = await loadState();
  current.ws = stripBuffer(ws);
  await saveState(current);
}

export async function persistRules(materialRules, thicknessRules) {
  const current = await loadState();
  current.materialRules = materialRules;
  current.thicknessRules = thicknessRules;
  await saveState(current);
}

export async function persistLicense(license) {
  const current = await loadState();
  current.license = license;
  await saveState(current);
}

export async function persistSettings(settings) {
  const current = await loadState();
  current.settings = settings;
  await saveState(current);
}
