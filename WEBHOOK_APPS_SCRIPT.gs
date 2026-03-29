/**
 * AquaBiz Pro - Obzvon webhook
 * Deploy: Execute as "Me", access "Anyone with the link".
 * Sheet name: "Обзвон ВСЕ"
 */

const TARGET_SHEET_NAME = 'Обзвон ВСЕ';

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(TARGET_SHEET_NAME);
    if (!sh) return json({ ok: false, error: `Sheet topilmadi: ${TARGET_SHEET_NAME}` });

    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '[]';
    const data = JSON.parse(raw);
    const rows = Array.isArray(data) ? data : [data];
    if (!rows.length) return json({ ok: true, inserted: 0 });

    const startNo = nextNo(sh);
    const values = rows.map((r, i) => {
      const no = startNo + i;
      // A № | B Mijoz ID | C Mijoz ismi | D Sana | E Maqsad | F Izoh | G Keyingi sana | H Operator | I Zakaz soni | J bo'sh | K Zakaz sanasi
      return [
        no,
        safe(r.mijozId),
        safe(r.mijozIsmi),
        safeDate(r.sana),
        safe(r.maqsad),
        safe(r.izoh),
        safeDate(r.keyingiSana),
        safe(r.operator),
        safe(r.zakazSoni),
        '',
        safeDate(r.zakazSanasi),
      ];
    });

    sh.getRange(sh.getLastRow() + 1, 1, values.length, 11).setValues(values);
    return json({ ok: true, inserted: values.length });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function doGet() {
  return json({ ok: true, service: 'AquaBiz Obzvon Webhook', sheet: TARGET_SHEET_NAME });
}

function nextNo(sh) {
  const last = sh.getLastRow();
  if (last < 2) return 1;
  const v = sh.getRange(last, 1).getValue();
  const n = Number(v);
  return Number.isFinite(n) && n > 0 ? n + 1 : last;
}

function safe(v) {
  return v == null ? '' : String(v).trim();
}

function safeDate(v) {
  if (!v) return '';
  const d = new Date(v);
  if (!isNaN(d)) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM.yyyy');
  return String(v);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

