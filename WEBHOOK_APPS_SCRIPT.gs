/**
 * AquaBiz Pro webhook
 *
 * Deploy:
 * 1) Extensions -> Apps Script
 * 2) Deploy -> New deployment -> Web app
 * 3) Execute as: Me
 * 4) Who has access: Anyone with the link
 */

const TARGET_SHEET_NAME = '\u041E\u0431\u0437\u0432\u043E\u043D \u0412\u0421\u0415'; // "Обзвон ВСЕ"
const ACCESS_CONFIG_KEY = 'AQ_ACCESS_CONFIG_V1';

function doGet(e) {
  const action = safe((e && e.parameter && e.parameter.action) || '');
  if (action === 'access_get') {
    return getAccessConfig_();
  }
  return json_({ ok: true, service: 'AquaBiz Webhook', sheet: TARGET_SHEET_NAME });
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '';
    const data = parseJsonSafe_(raw, null);

    // Access config sync (global permissions for all devices)
    const action = safe(data && data.action);
    if (action === 'access_set') {
      return setAccessConfig_(data && data.accessConfig);
    }

    // Legacy obzvon rows sync
    return appendObzvonRows_(data);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function appendObzvonRows_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sh) return json_({ ok: false, error: 'Sheet topilmadi: ' + TARGET_SHEET_NAME });

  const rows = Array.isArray(data) ? data : (data ? [data] : []);
  if (!rows.length) return json_({ ok: true, inserted: 0 });

  const startNo = nextNo_(sh);
  const values = rows.map((r, i) => {
    const no = startNo + i;
    // A № | B Mijoz ID | C Mijoz ismi | D Sana | E Maqsad | F Izoh
    // G Keyingi sana | H Operator | I Zakaz soni | J bo'sh | K Zakaz sanasi
    return [
      no,
      safe(r.mijozId),
      safe(r.mijozIsmi),
      safeDate_(r.sana),
      safe(r.maqsad),
      safe(r.izoh),
      safeDate_(r.keyingiSana),
      safe(r.operator),
      safe(r.zakazSoni),
      '',
      safeDate_(r.zakazSanasi),
    ];
  });

  sh.getRange(sh.getLastRow() + 1, 1, values.length, 11).setValues(values);
  return json_({ ok: true, inserted: values.length });
}

function getAccessConfig_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(ACCESS_CONFIG_KEY);
  const accessConfig = parseJsonSafe_(raw, {});
  return json_({
    ok: true,
    accessConfig: accessConfig || {},
  });
}

function setAccessConfig_(cfg) {
  const payload = normalizeAccessConfigPayload_(cfg);
  if (!payload) return json_({ ok: false, error: 'accessConfig noto\'g\'ri formatda' });

  const withMeta = {
    users: payload.users,
    access: payload.access,
    userCreds: payload.userCreds,
    updatedAt: safe(payload.updatedAt) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX"),
    updatedBy: safe(payload.updatedBy),
  };
  PropertiesService.getScriptProperties().setProperty(ACCESS_CONFIG_KEY, JSON.stringify(withMeta));
  return json_({ ok: true, saved: true });
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

function nextNo_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return 1;
  const v = sh.getRange(last, 1).getValue();
  const n = Number(v);
  return Number.isFinite(n) && n > 0 ? n + 1 : last;
}

function safe(v) {
  return v == null ? '' : String(v).trim();
}

function safeDate_(v) {
  if (!v) return '';
  const d = new Date(v);
  if (!isNaN(d)) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  return String(v);
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
