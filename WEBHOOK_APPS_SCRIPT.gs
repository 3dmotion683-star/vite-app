/**
 * AquaBiz Pro webhook (access sync only)
 *
 * Deploy:
 * 1) Extensions -> Apps Script
 * 2) Deploy -> New deployment -> Web app
 * 3) Execute as: Me
 * 4) Who has access: Anyone with the link
 */

const ACCESS_CONFIG_KEY = 'AQ_ACCESS_CONFIG_V1';
const OBZVON_NEW_ROWS_KEY = 'AQ_OBZVON_NEW_ROWS_V1';
const MAX_OBZVON_NEW_ROWS = 20000;

function doGet(e) {
  const action = safe((e && e.parameter && e.parameter.action) || '');
  if (action === 'access_get') return getAccessConfig_();
  if (action === 'obzvon_new_get') return getObzvonNewRows_();
  return json_({ ok: true, service: 'AquaBiz Webhook', mode: 'access_only' });
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '';
    const data = parseJsonSafe_(raw, null);
    const action = safe(data && data.action);
    if (action === 'access_set') return setAccessConfig_(data && data.accessConfig);
    if (action === 'obzvon_new_upsert') return upsertObzvonNewRows_(data && data.rows, data && data.by);
    if (action === 'obzvon_new_sheet_last_id') return getObzvonNewSheetLastId_(data);
    if (action === 'obzvon_new_sheet_replace') return replaceObzvonNewRowsSheet_(data);
    if (action === 'obzvon_new_sheet_append' || action === 'obzvon_new_export') {
      return appendObzvonNewRowsToSheet_(data);
    }
    return json_({ ok: false, error: 'Unsupported action', mode: 'access_only' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function getAccessConfig_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(ACCESS_CONFIG_KEY);
  const accessConfig = parseJsonSafe_(raw, {});
  return json_({ ok: true, accessConfig: accessConfig || {} });
}

function setAccessConfig_(cfg) {
  const payload = normalizeAccessConfigPayload_(cfg);
  if (!payload) return json_({ ok: false, error: "accessConfig noto'g'ri formatda" });

  const withMeta = {
    users: payload.users,
    access: payload.access,
    userCreds: payload.userCreds,
    updatedAt: safe(payload.updatedAt) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX"),
    updatedBy: safe(payload.updatedBy),
  };
  PropertiesService.getScriptProperties().setProperty(ACCESS_CONFIG_KEY, JSON.stringify(withMeta));
  return json_({ ok: true, saved: true, mode: 'access_set' });
}

function getObzvonNewRows_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(OBZVON_NEW_ROWS_KEY);
  const rows = parseJsonSafe_(raw, []);
  return json_({ ok: true, rows: Array.isArray(rows) ? rows : [] });
}

function upsertObzvonNewRows_(rows, by) {
  const incomingRaw = Array.isArray(rows) ? rows : [];
  const incoming = incomingRaw
    .map((r) => normalizeObzvonNewRow_(r))
    .filter(Boolean);
  if (!incoming.length) return json_({ ok: true, saved: true, count: 0 });

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(OBZVON_NEW_ROWS_KEY);
  const existing = Array.isArray(parseJsonSafe_(raw, [])) ? parseJsonSafe_(raw, []) : [];

  const map = {};
  existing.forEach((r) => {
    const n = normalizeObzvonNewRow_(r);
    if (n) map[n.rid] = n;
  });

  incoming.forEach((r) => {
    const prev = map[r.rid];
    if (!prev || ts_(r.updatedAt) >= ts_(prev.updatedAt)) {
      map[r.rid] = r;
    }
  });

  let merged = Object.keys(map).map((k) => map[k]);
  merged.sort((a, b) => ts_(b.updatedAt) - ts_(a.updatedAt));
  if (merged.length > MAX_OBZVON_NEW_ROWS) merged = merged.slice(0, MAX_OBZVON_NEW_ROWS);

  props.setProperty(OBZVON_NEW_ROWS_KEY, JSON.stringify(merged));
  return json_({
    ok: true,
    saved: true,
    by: safe(by),
    count: merged.length,
  });
}

function getObzvonSheetConfig_(data) {
  const sheetId = safe(data && data.sheetId);
  const sheetName = safe(data && data.sheetName) || 'Barcha_obzvon_yangi';
  const fallbackHeaders = ['No', 'ID', 'Mijoz', 'Sana', 'Mavzu', 'Izoh', 'Keyingi sana', 'Z.soni / summa', 'Zakaz sanasi', 'Operator', 'Kompaniya', 'RID', 'UpdatedAt'];
  const fallbackKeys = ['no', 'customerId', 'customer', 'callDate', 'topic', 'note', 'nextDate', 'orderCount', 'orderDate', 'operator', 'company', 'rid', 'updatedAt'];
  const labelToKey = {
    'No': 'no',
    'ID': 'customerId',
    'Mijoz': 'customer',
    'Sana': 'callDate',
    'Mavzu': 'topic',
    'Izoh': 'note',
    'Keyingi sana': 'nextDate',
    'Z.soni / summa': 'orderCount',
    'Zakaz sanasi': 'orderDate',
    'Operator': 'operator',
    'Kompaniya': 'company',
    'SyncID': 'stableId',
    'RID': 'rid',
    'UpdatedAt': 'updatedAt',
  };
  let headers = Array.isArray(data && data.headers)
    ? data.headers.map((h) => safe(h)).filter(Boolean)
    : [];
  if (!headers.length) headers = fallbackHeaders;
  const keys = headers.map((h) => labelToKey[h] || h);
  const ridIndex = Math.max(0, headers.findIndex((h) => safe(h).toUpperCase() === 'RID'));
  // Eski strukturada SyncID alohida ustun edi, yangi strukturada esa No ustuniga yoziladi
  const stableIdIndex = (() => {
    const syncIdx = headers.findIndex((h) => safe(h).toUpperCase() === 'SYNCID');
    if (syncIdx >= 0) return syncIdx;
    return headers.findIndex((h) => safe(h).toUpperCase() === 'NO');
  })();
  return { sheetId, sheetName, headers, keys, ridIndex, stableIdIndex, fallbackKeys };
}

function prepareObzvonSheetRows_(rows, conf) {
  const headers = conf.headers;
  const keys = conf.keys;
  const fallbackKeys = conf.fallbackKeys;
  return (Array.isArray(rows) ? rows : []).map((r, i) => {
    let arr;
    if (Array.isArray(r)) {
      arr = [];
      for (let j = 0; j < headers.length; j++) arr.push(safe(r[j]));
    } else {
      const obj = (r && typeof r === 'object') ? r : {};
      const useKeys = keys.length === headers.length ? keys : fallbackKeys;
      arr = useKeys.map((k) => safe(obj[k]));
    }
    if (!arr.some((v) => safe(v))) return null;
    if (!safe(arr[conf.ridIndex])) {
      // RID bo'sh bo'lsa vaqtinchalik stabil identifikator yaratamiz
      arr[conf.ridIndex] = 'rid_' + Utilities.base64EncodeWebSafe(String(new Date().getTime()) + '_' + i).replace(/=+$/g, '');
    }
    return arr;
  }).filter(Boolean);
}

function getObzvonNewSheetLastId_(data) {
  const conf = getObzvonSheetConfig_(data);
  if (!conf.sheetId) return json_({ ok: false, error: 'sheetId berilmagan' });
  const ss = SpreadsheetApp.openById(conf.sheetId);
  let sh = ss.getSheetByName(conf.sheetName);
  if (!sh) return json_({ ok: true, sheetName: conf.sheetName, lastRid: '', lastStableId: 0, rows: 0 });
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return json_({ ok: true, sheetName: conf.sheetName, lastRid: '', lastStableId: 0, rows: 0 });

  const ridVals = sh.getRange(2, conf.ridIndex + 1, lastRow - 1, 1).getValues();
  let lastRid = '';
  for (let i = ridVals.length - 1; i >= 0; i--) {
    const v = safe(ridVals[i][0]);
    if (v) { lastRid = v; break; }
  }

  let lastStableId = 0;
  if (conf.stableIdIndex >= 0) {
    const stableVals = sh.getRange(2, conf.stableIdIndex + 1, lastRow - 1, 1).getValues();
    for (let i = stableVals.length - 1; i >= 0; i--) {
      const n = Number(safe(stableVals[i][0]));
      if (n > 0) { lastStableId = n; break; }
    }
  }
  return json_({ ok: true, sheetName: conf.sheetName, lastRid: lastRid, lastStableId: lastStableId, rows: lastRow - 1 });
}

function replaceObzvonNewRowsSheet_(data) {
  const conf = getObzvonSheetConfig_(data);
  if (!conf.sheetId) return json_({ ok: false, error: 'sheetId berilmagan' });
  const rows = prepareObzvonSheetRows_(data && data.rows, conf);
  const ss = SpreadsheetApp.openById(conf.sheetId);
  let sh = ss.getSheetByName(conf.sheetName);
  if (!sh) sh = ss.insertSheet(conf.sheetName);
  sh.clearContents();
  sh.getRange(1, 1, 1, conf.headers.length).setValues([conf.headers]);
  if (rows.length) {
    // Bir xil RID kelib qolsa, birinchisini qoldiramiz
    const uniq = [];
    const seen = {};
    for (let i = 0; i < rows.length; i++) {
      const rid = safe(rows[i][conf.ridIndex]);
      if (rid && seen[rid]) continue;
      if (rid) seen[rid] = true;
      uniq.push(rows[i]);
    }
    if (uniq.length) {
      sh.getRange(2, 1, uniq.length, conf.headers.length).setValues(uniq);
      return json_({ ok: true, saved: true, replaced: true, inserted: uniq.length, sheetName: conf.sheetName });
    }
  }
  return json_({ ok: true, saved: true, replaced: true, inserted: 0, sheetName: conf.sheetName });
}

function appendObzvonNewRowsToSheet_(data) {
  const conf = getObzvonSheetConfig_(data);
  if (!conf.sheetId) return json_({ ok: false, error: 'sheetId berilmagan' });
  const withHeaders = Boolean(data && data.withHeaders);
  const rows = prepareObzvonSheetRows_(data && data.rows, conf);

  const ss = SpreadsheetApp.openById(conf.sheetId);
  let sh = ss.getSheetByName(conf.sheetName);
  if (!sh) sh = ss.insertSheet(conf.sheetName);

  const lastRow = sh.getLastRow();
  if (withHeaders && lastRow === 0) {
    sh.getRange(1, 1, 1, conf.headers.length).setValues([conf.headers]);
  }

  const existingRidSet = {};
  const rowsInSheet = sh.getLastRow();
  if (rowsInSheet > 1) {
    const ridVals = sh.getRange(2, conf.ridIndex + 1, rowsInSheet - 1, 1).getValues();
    ridVals.forEach((v) => {
      const rid = safe(v[0]);
      if (rid) existingRidSet[rid] = true;
    });
  }

  const prepared = [];
  rows.forEach((arr) => {
    const rid = safe(arr[conf.ridIndex]);
    if (rid && existingRidSet[rid]) return;
    if (rid) existingRidSet[rid] = true;
    prepared.push(arr);
  });

  if (prepared.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, prepared.length, conf.headers.length).setValues(prepared);
  }

  return json_({
    ok: true,
    saved: true,
    inserted: prepared.length,
    sheetName: conf.sheetName,
  });
}

function normalizeObzvonNewRow_(row) {
  if (!row || typeof row !== 'object') return null;
  const rid = safe(row.rid || row._rid);
  if (!rid) return null;
  const out = {
    rid: rid,
    no: safe(row.no),
    customer: safe(row.customer),
    callDate: safe(row.callDate),
    topic: safe(row.topic),
    note: safe(row.note),
    nextDate: safe(row.nextDate),
    orderCount: safe(row.orderCount),
    operator: safe(row.operator),
    customerId: safe(row.customerId || row.id),
    orderDate: safe(row.orderDate),
    updatedAt: safe(row.updatedAt) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX"),
  };
  if (!out.customer && !out.customerId) return null;
  if (!out.note && !out.orderCount && !out.nextDate && !out.orderDate) return null;
  return out;
}

function ts_(v) {
  const s = safe(v);
  if (!s) return 0;
  const t = new Date(s).getTime();
  return isNaN(t) ? 0 : t;
}

function normalizeAccessConfigPayload_(cfg) {
  if (!cfg || typeof cfg !== 'object') return null;
  const users = Array.isArray(cfg.users)
    ? cfg.users.map((u) => safe(u)).filter(Boolean)
    : [];
  if (!users.length) return null;
  const access = (cfg.access && typeof cfg.access === 'object') ? cfg.access : {};
  const userCreds = (cfg.userCreds && typeof cfg.userCreds === 'object') ? cfg.userCreds : {};
  return {
    users: users,
    access: access,
    userCreds: userCreds,
    updatedAt: safe(cfg.updatedAt),
    updatedBy: safe(cfg.updatedBy),
  };
}

function safe(v) {
  return v == null ? '' : String(v).trim();
}

function parseJsonSafe_(raw, fallback) {
  try {
    if (!raw) return fallback;
    return JSON.parse(raw);
  } catch (e) {
    return fallback;
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
