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

function doGet(e) {
  const action = safe((e && e.parameter && e.parameter.action) || '');
  if (action === 'access_get') return getAccessConfig_();
  return json_({ ok: true, service: 'AquaBiz Webhook', mode: 'access_only' });
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '';
    const data = parseJsonSafe_(raw, null);
    const action = safe(data && data.action);
    if (action === 'access_set') return setAccessConfig_(data && data.accessConfig);
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
