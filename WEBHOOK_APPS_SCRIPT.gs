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

function appendObzvonNewRowsToSheet_(data) {
  const sheetId = safe(data && data.sheetId);
  const sheetName = safe(data && data.sheetName) || 'Barcha_obzvon_yangi';
  const withHeaders = Boolean(data && data.withHeaders);
  const rows = Array.isArray(data && data.rows) ? data.rows : [];
  if (!sheetId) return json_({ ok: false, error: 'sheetId berilmagan' });

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
    'RID': 'rid',
    'UpdatedAt': 'updatedAt',
  };

  let headers = Array.isArray(data && data.headers)
    ? data.headers.map((h) => safe(h)).filter(Boolean)
    : [];
  if (!headers.length) headers = fallbackHeaders;
  const keys = headers.map((h) => labelToKey[h] || h);

  const ss = SpreadsheetApp.openById(sheetId);
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);

  const lastRow = sh.getLastRow();
  if (withHeaders && lastRow === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const prepared = rows.map((r) => {
    if (Array.isArray(r)) {
      const arr = [];
      for (let i = 0; i < headers.length; i++) arr.push(safe(r[i]));
      return arr;
    }
    const obj = (r && typeof r === 'object') ? r : {};
    if (keys.length !== headers.length) {
      return fallbackKeys.map((k) => safe(obj[k]));
    }
    return keys.map((k) => safe(obj[k]));
  }).filter((arr) => arr.some((v) => safe(v)));

  if (prepared.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, prepared.length, headers.length).setValues(prepared);
  }

  return json_({
    ok: true,
    saved: true,
    inserted: prepared.length,
    sheetName: sheetName,
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
