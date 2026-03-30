import { useState, useMemo, useCallback, useEffect, useDeferredValue } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, AreaChart, Area,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
} from 'recharts';

/* ═══════════════ EXCLUDED MERCHANT KEYWORDS ═══════════════ */
const EXCLUDED_KEYWORDS = [
  'Dividend','UOS','Rasxod','Rasxodnik','Postavshik',
  'Personal','Dolg','VIP','KULER','Ofis','Neus',
];
const isExcludedMerchant = (name) => {
  const nl = (name || '').toLowerCase();
  return EXCLUDED_KEYWORDS.some((k) => nl.includes(k.toLowerCase()));
};
const ALLOWED_AA_KEYWORD = 'murodbash';
const EXCLUDED_Z_CATEGORIES = [
  'mijozlar qatoriga kiritilmaganlar',
  'dividend',
  'uos',
  'rasxod',
  'rasxodnik',
  'postavshik',
  'personal',
  'dolg',
  'vip',
  'kuler',
  'ofis',
  'neus',
];
const isAllowedAA = (v) => {
  const s = String(v || '').trim().toLowerCase();
  return s === '' || s.includes(ALLOWED_AA_KEYWORD);
};
const isExcludedZCategory = (v) => {
  const s = String(v || '').trim().toLowerCase();
  return EXCLUDED_Z_CATEGORIES.some((k) => s === k || s.includes(k));
};
const normCurrency = (v) => {
  const s = String(v || '').trim().toUpperCase();
  if (['USD', '$', 'ДОЛЛАР'].includes(s)) return 'USD';
  return 'UZS';
};
function getDebtStats(customers = []) {
  const debtorsUZS = customers.filter((c) => c.balanceUZS < 0);
  return {
    anyCount: debtorsUZS.length,
    uzsCount: debtorsUZS.length,
    uzsSum: debtorsUZS.reduce((s, c) => s + Math.abs(c.balanceUZS), 0),
  };
}
function recalcInstallment(row, monthsRaw) {
  const months = Math.max(1, Math.min(60, Number(monthsRaw) || 6));
  const principal = Math.max(0, row.principal || 0);
  const paid = Math.min(Math.max(0, row.paid || 0), principal);
  const monthly = months > 0 ? Math.ceil(principal / months) : 0;
  const remaining = Math.max(0, principal - paid);
  const paidMonths = monthly > 0 ? Math.floor(paid / monthly) : 0;
  const monthsLeft = monthly > 0 ? Math.ceil(remaining / monthly) : 0;

  let dueCount = 0;
  let overdueAmount = 0;
  if (row.purchaseDate && monthly > 0) {
    const start = toDate(row.purchaseDate);
    if (start) {
      const now = new Date();
      const elapsed = Math.max(0, (now.getFullYear() - start.getFullYear()) * 12 + (now.getMonth() - start.getMonth()) + 1);
      const shouldPayMonths = Math.min(months, elapsed);
      const shouldPaid = shouldPayMonths * monthly;
      overdueAmount = Math.max(0, Math.min(principal, shouldPaid) - paid);
      dueCount = overdueAmount > 0 ? 1 : 0;
    }
  }
  return { ...row, months, monthly, remaining, paidMonths, monthsLeft, dueCount, overdueAmount };
}

/* ═══════════════ CONFIG ═══════════════ */
const SHEET_CONFIG = {
  url: 'https://docs.google.com/spreadsheets/d/1RND6D5JIWh6vnk8i_FO1mx1if4hx182VNaPzGQtKuQM/edit?gid=1272423090#gid=1272423090',
  gids: { merchants: '', balans: '', orders: '', cashbox: '', integration: '', mijozlar: '' },
};
const OBZVON_ALL_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1lfjqNFaD2Gy-tyKrmWVX4ja2DvsyLcACaGlW2ZJrkf8/edit?pli=1&gid=0#gid=0';
const OBZVON_ALL_LIMIT = 100;
const OBZVON_ALL_INSTALLED_KEY = 'aq-obzvon-all-installed';
// Apps Script Web App URL ni shu yerga qo'ying (Deploy qilingandan keyin):
// Misol: https://script.google.com/macros/s/AKfycb.../exec
const OBZVON_WEBHOOK_DEFAULT = 'https://script.google.com/macros/s/AKfycbyi1OP_a_5C7-TBujLuo9dDast0RLVelhQsTiO6dlN_mefi55vnTHm_ZRXpDFFTTXb7qA/exec';

/* ═══════════════ HELPERS ═══════════════ */
const fmt  = (n) => new Intl.NumberFormat('uz-UZ').format(Math.round(n || 0));
const fmtM = (n) =>
  n >= 1e6 ? (n / 1e6).toFixed(1) + 'M' :
  n >= 1e3 ? (n / 1e3).toFixed(0) + 'K' : fmt(n);

function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  const n = Number(v);
  if (!isNaN(n) && n >= 40000 && n <= 60000) {
    const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
    return isNaN(d) ? null : d;
  }
  const s = String(v).trim();
  let m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/)))  return new Date(+m[1], +m[2]-1, +m[3]);
  if ((m = s.match(/^(\d{2})[./](\d{2})[./](\d{4})/))) return new Date(+m[3], +m[2]-1, +m[1]);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
const fmtD = (v) => { const d = toDate(v); return d ? d.toLocaleDateString('ru-RU') : '—'; };
const daysAgo = (v) => {
  const d = toDate(v); if (!d) return null;
  const n = new Date(); n.setHours(0,0,0,0); d.setHours(0,0,0,0);
  return Math.floor((n - d) / 864e5);
};
const monthKey = (v) => {
  const d = toDate(v); if (!d) return null;
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
};
const toNum = (v) => {
  if (v == null || v === '' || v === '-') return 0;
  const n = parseFloat(String(v).replace(/[\s\xa0]/g,'').replace(',','.'));
  return isNaN(n) ? 0 : n;
};
const normText = (v) => String(v || '').trim().toLowerCase();
const isOrderDoc = (v) => {
  const t = normText(v);
  return t === 'заказ' || t === 'zakaz' || t === 'р—р°рєр°р·';
};
const isReturnDoc = (v) => {
  const t = normText(v);
  return t === 'возврат' || t === 'vozvrat' || t === 'р’рѕр·рір‚р°с‚';
};
const isPaymentFromCounterparty = (v) => {
  const t = normText(v);
  return t === 'оплата от контрагента' || t === 'oplata ot kontragenta' || t.includes('контрагента');
};
const isPaymentToCounterparty = (v) => {
  const t = normText(v);
  return t === 'оплата контрагенту' || t === 'oplata kontragentu' || t.includes('контрагенту');
};
function parseObzvonAllRows(rows = []) {
  if (!Array.isArray(rows) || rows.length < 2) return [];
  const headerIdx = rows.findIndex((r) => {
    const c0 = String((r || [])[0] || '').trim().toLowerCase();
    const c1 = String((r || [])[1] || '').trim().toLowerCase();
    return c0 === '№' || c0 === 'no' || c1.includes('контрагент') || c1.includes('kontragent');
  });
  const start = headerIdx >= 0 ? headerIdx + 1 : 0;
  return rows.slice(start)
    .map((r) => ({
      no: String(r[0] || '').trim(),
      customer: String(r[1] || '').trim(),
      callDate: r[2] || '',
      topic: String(r[3] || '').trim(),
      note: String(r[4] || '').trim(),
      nextDate: r[5] || '',
      orderCount: String(r[6] || '').trim(),
      operator: String(r[7] || '').trim(),
      customerId: String(r[8] || '').trim(),
      orderDate: r[10] || '',
    }))
    .filter((x) => x.customer || x.customerId || x.no);
}

/* ═══════════════ STATUS HELPERS ═══════════════ */
// ДОСТАВЛЕН va ПОЛУЧЕН НА СКЛАД — ikkisi ham yetkazilgan hisoblanadi
const isDeliveredStatus = (s) => {
  const st = (s || '').trim();
  return st === 'ДОСТАВЛЕН' || st === 'ПОЛУЧЕН НА СКЛАД';
};
// Bekor qilingan
const isCancelledStatus = (s) => (s || '').trim() === 'ОТМЕНЕНО';

/* ═══════════════ MAHSULOT ANIQLASH ═══════════════ */
// Tara ga ta'sir qiladigan 3 ta mahsulot (zakaz=+, vozvrat=-)
const isTaraAffectingProduct = (p) => {
  const pl = (p || '').trim();
  return (
    pl === 'Murodbaxsh 18.9L' ||
    pl === 'БОНУС Murodbaxsh 18.9L' ||
    pl === 'Тара 18.9L (пустой)'
  );
};
// Faqat suv mahsulotlari (daromad hisoblash uchun)
const isWaterProduct = (p) => {
  const pl = (p || '').trim();
  return pl === 'Murodbaxsh 18.9L' || pl === 'БОНУС Murodbaxsh 18.9L';
};

/* ═══════════════ EXCEL READER ═══════════════ */
function readExcelFile(file) {
  return new Promise((res, rej) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type:'array', cellDates:false, raw:true });
        const result = {};
        wb.SheetNames.forEach((name) => {
          result[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { header:1, defval:'', raw:true });
        });
        res({ fileName: file.name, sheets: result });
      } catch (err) { rej(err); }
    };
    reader.onerror = rej;
    reader.readAsArrayBuffer(file);
  });
}

/* ═══════════════ GOOGLE SHEETS ═══════════════ */
function buildGsUrl(sheetId, gid) {
  return `https://docs.google.com/spreadsheets/d/${sheetId}/pub?gid=${gid}&single=true&output=csv`;
}
function buildGsUrlByName(sheetId, sheetName) {
  return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
}
function parseGsCsv(text) {
  try {
    const wb = XLSX.read(text, { type:'string', raw:true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'', raw:true });
    return rows
      .map((r) => (r || []).map((v) => String(v ?? '').trim()))
      .filter((r) => r.some((v) => v !== ''));
  } catch {
    const lines = text.replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n');
    return lines.map((line) => {
      const cells = [];
      let cur = '';
      let inQ = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          if (inQ && line[i + 1] === '"') {
            cur += '"';
            i++;
          } else {
            inQ = !inQ;
          }
        } else if (ch === ',' && !inQ) {
          cells.push(cur);
          cur = '';
        } else {
          cur += ch;
        }
      }
      cells.push(cur);
      return cells.map((v) => String(v || '').trim());
    }).filter((r) => r.some((v) => v));
  }
}
async function fetchSheetCsv(sheetId, gid, label) {
  const sources = [
    buildGsUrl(sheetId, gid),
    `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`,
    `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${gid}`,
  ];
  const urls = sources.flatMap((url) => [
    url,
    `https://corsproxy.io/?${encodeURIComponent(url)}`,
    `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`,
  ]);
  for (const u of urls) {
    try {
      const r = await fetch(u, { signal: AbortSignal.timeout(12000) });
      if (!r.ok) continue;
      const text = await r.text();
      if (text.trim().startsWith('<')) continue;
      const rows = parseGsCsv(text);
      if (rows.length < 2) continue;
      return rows;
    } catch { continue; }
  }
  throw new Error(`"${label}" varaqi yuklanmadi. GID: ${gid}`);
}
async function fetchSheetCsvByName(sheetId, sheetName, label) {
  const sources = [
    buildGsUrlByName(sheetId, sheetName),
    `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&sheet=${encodeURIComponent(sheetName)}`,
  ];
  const urls = sources.flatMap((url) => [
    url,
    `https://corsproxy.io/?${encodeURIComponent(url)}`,
    `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`,
  ]);
  for (const u of urls) {
    try {
      const r = await fetch(u, { signal: AbortSignal.timeout(12000) });
      if (!r.ok) continue;
      const text = await r.text();
      if (text.trim().startsWith('<')) continue;
      const rows = parseGsCsv(text);
      if (rows.length < 2) continue;
      return rows;
    } catch { continue; }
  }
  throw new Error(`"${label}" varaqi yuklanmadi. Sheet: ${sheetName}`);
}
async function fetchSheetOpenSheet(sheetId, sheetName, label) {
  const base = `https://opensheet.elk.sh/${sheetId}/${encodeURIComponent(sheetName)}`;
  const urls = [
    base,
    `https://corsproxy.io/?${encodeURIComponent(base)}`,
    `https://api.allorigins.win/raw?url=${encodeURIComponent(base)}`,
  ];
  for (const u of urls) {
    try {
      const r = await fetch(u, { signal: AbortSignal.timeout(12000) });
      if (!r.ok) continue;
      const js = await r.json();
      if (!Array.isArray(js) || js.length === 0) continue;
      const headers = Object.keys(js[0] || {});
      if (!headers.length) continue;
      const rows = [headers, ...js.map((obj) => headers.map((h) => String(obj?.[h] ?? '').trim()))];
      if (rows.length < 2) continue;
      return rows;
    } catch { continue; }
  }
  throw new Error(`"${label}" varaqi opensheet orqali yuklanmadi`);
}
async function loadFromGoogleSheets(sheetId, gids, onProgress, mode='gid') {
  if (mode === 'named') {
    const byName = [
      { key:'public.view_merchants', label:'Mijozlar', sheet:'public.view_merchants' },
      { key:'public.view_current_merchant_event_balance', label:'Balans', sheet:'public.view_current_merchant_event_balance' },
      { key:'public.view_item_basket', label:'Zakazlar', sheet:'public.view_item_basket' },
      { key:'public.view_cashbox_documents', label:'Kassa', sheet:'public.view_cashbox_documents' },
      { key:'intigratsiya', label:'Integratsiya', sheet:'intigratsiya' },
      { key:'mijozlar', label:'Mijoz biriktiruv', sheet:'mijozlar' },
    ];
    const sheets = {};
    for (const s of byName) {
      onProgress(`${s.label} yuklanmoqda...`);
      sheets[s.key] = await fetchSheetCsvByName(sheetId, s.sheet, s.label);
    }
    if (!sheets['public.view_merchants']) throw new Error('Mijozlar varaqi topilmadi!');
    return { sheets };
  }

  const sheetNames = [
    { key: 'public.view_merchants',           gid: gids.merchants,   label: 'Mijozlar' },
    { key: 'Balans',                           gid: gids.balans,      label: 'Balans' },
    { key: 'public.view_item_basket',          gid: gids.orders,      label: 'Zakazlar' },
    { key: 'public.view_cashbox_documents',    gid: gids.cashbox,     label: 'Kassa' },
    { key: 'intigratsiya',                     gid: gids.integration, label: 'Integratsiya' },
    { key: 'mijozlar',                         gid: gids.mijozlar,    label: 'Mijoz biriktiruv' },
  ];
  const sheets = {};
  for (const s of sheetNames) {
    if (!s.gid) continue;
    onProgress(`${s.label} yuklanmoqda...`);
    sheets[s.key] = await fetchSheetCsv(sheetId, s.gid, s.label);
  }
  if (!sheets['public.view_merchants']) throw new Error('Mijozlar varaqi topilmadi!');
  return { sheets };
}
function extractSheetId(url) {
  const m = (url || '').match(/\/spreadsheets\/d\/([^/&#?]+)/);
  return m ? m[1] : '';
}
function extractGid(url) {
  const m = (url || '').match(/[?&#]gid=([0-9]+)/);
  return m ? m[1] : '';
}

/* ═══════════════ DATA PROCESSING ═══════════════ */
function processAll(mainData) {
  if (!mainData) return { customers:[], orders:[], cashbox:[], contacts:[] };

  /* ── 1. MERCHANTS ──
     A(r[0])=ID  C(r[2])=Nom  D(r[3])=Tel  E(r[4])=Kontakt
     O(r[14])=Manzil  V(r[21])=Rayon  Y(r[24])=EskiID  Z(r[25])=Manba */
  const merchantSheet = mainData.sheets['public.view_merchants'];
  const contacts = [];
  const allMerchants = [];
  if (merchantSheet) {
    merchantSheet.slice(1).forEach((r) => {
      const id = String(r[0] || '').trim();
      if (!id) return;
      if (!/^\d+$/.test(id)) return;
      const name = String(r[2] || '').trim();
      const source = String(r[25] || '').trim(); // Z ustun
      const merchNote = String(r[19] || '').trim(); // T ustun = primechaniya
      allMerchants.push({ id, name, source, merchNote });
      if (isExcludedMerchant(name)) return;
      const aaValue = String(r[26] || '').trim(); // AA ustun
      if (isExcludedZCategory(source)) return;
      if (!isAllowedAA(aaValue)) return;
      contacts.push({
        id,
        name,
        phone:    String(r[3]  || '').trim().replace(/[^+\d]/g,'').slice(0,13),
        contact:  String(r[4]  || '').trim(),
        address:  String(r[14] || '').trim(),
        district: String(r[21] || '').split(';').map((s)=>s.trim()).filter(Boolean)
                    .find((s) => !s.includes('Общая')) ||
                  String(r[21] || '').split(';')[0].trim(),
        source,
        merchantNote: merchNote,
        eskiId:   String(r[24] || '').trim(), // Y ustun — eski ID (integratsiya uchun)
      });
    });
  }
  const contactById = {};
  contacts.forEach((c) => { contactById[c.id] = c; });

  /* ── 2. BALANS ──
     8-qatordan boshlanadi (slice(7))
     G(r[6])=MijozID  D(r[3])=BalansUZS  C(r[2])=BalansUSD */
  const balMapUZS = {};
  const balMapUSD = {};
  const balSheet =
    mainData.sheets['public.view_current_merchant_event_balance'] ||
    mainData.sheets['Balans'];
  if (balSheet) {
    balSheet.slice(7).forEach((r) => {
      const mid = String(r[6] || '').trim(); // G ustun
      if (!mid) return;
      if (!/^\d+$/.test(mid)) return;
      balMapUZS[mid] = toNum(r[3]); // D ustun = UZS balans
      balMapUSD[mid] = toNum(r[2]); // C ustun = USD balans
    });
  }

  /* ── 3. INTIGRATSIYA ──
     2-qatordan boshlanadi (slice(1))
     A(r[0])=EskiID  H(r[7])=TaraOstatka
     Logika: merchants.eskiId === intigratsiya.eskiId → boshlang'ich tara */
  const integTaraByEskiId = {};
  const integSheet = mainData.sheets['intigratsiya'];
  if (integSheet) {
    integSheet.slice(1).forEach((r) => {
      const eskiId = String(r[0] || '').trim(); // A ustun = eski ID
      if (!eskiId) return;
      const taraOstatka = toNum(r[7]); // H ustun = tara ostatka
      if (taraOstatka > 0) integTaraByEskiId[eskiId] = taraOstatka;
    });
  }

  /* ── 4. ZAKAZLAR ──
     A(r[0])=ZakazNom  D(r[3])=Mahsulot  H(r[7])=Miqdor  I(r[8])=Summa
     J(r[9])=Kat  K(r[10])=KontragentNom  L(r[11])=Valyuta
     M(r[12])=HujjatTuri  N(r[13])=Status  R(r[17])=UniqueID
     V(r[21])=Agent  X(r[23])=DostavchikIsmi  Y(r[24])=Sana  AD(r[29])=MijozID */
  const orderSheet = mainData.sheets['public.view_item_basket'];
  const rawOrders = [];
  if (orderSheet) {
    orderSheet.slice(1).forEach((r) => {
      const soNum      = String(r[0]  || '').trim();
      const product    = String(r[3]  || '').trim();
      const qty        = toNum(r[7]);
      const sum        = toNum(r[8]);
      const cat        = String(r[9]  || '').trim();
      const contName   = String(r[10] || '').trim();
      const currency   = String(r[11] || '').trim();
      const docType    = String(r[12] || '').trim();
      const status     = String(r[13] || '').trim();
      const uniqueId   = String(r[17] || '').trim();
      const agent      = String(r[21] || '').trim();
      const delivPerson= String(r[23] || '').trim(); // X ustun = dostavchik ismi
      const orderDate  = String(r[24] || '').trim(); // Y ustun = zakaz/vozvrat sanasi
      const mId        = String(r[29] || '').trim();
      const price      = qty && qty !== 0 ? Math.abs(sum / qty) : 0;

      if (!soNum || soNum === 'id') return;
      rawOrders.push({
        soNum, product, qty, sum, cat, contName, currency,
        docType, status, uniqueId, agent,
        delivPerson, // dostavchik ismi
        orderDate,   // zakaz/vozvrat sanasi
        mId, price,
      });
    });
  }

  /* ── 5. KASSA ──
     A(r[0])=OpStatus  B(r[1])=OpNom  E(r[4])=Summa  N(r[13])=Izoh
     R(r[17])=OpTuri  S(r[18])=Operator  T(r[19])=KontragentNom
     U(r[20])=Kassa   V(r[21])=Sana   X(r[23])=MijozID */
  const cashSheet = mainData.sheets['public.view_cashbox_documents'];
  const rawCash = [];
  if (cashSheet) {
    cashSheet.slice(1).forEach((r) => {
      const opStatus = String(r[0]  || '').trim();
      const opNum    = String(r[1]  || '').trim();
      const sana     = String(r[21] || r[2] || '').trim();
      const amount   = toNum(r[4]);
      const currency = normCurrency(r[11] || r[12] || r[10] || r[22] || '');
      const opType   = String(r[17] || '').trim();
      const operator = String(r[18] || '').trim();
      const contName = String(r[19] || r[3] || '').trim();
      const kassa    = String(r[20] || '').trim();
      const mId      = String(r[23] || '').trim();
      const note     = String(r[13] || '').trim();

      if (!opNum || opNum === 'Номер операции') return;
      rawCash.push({ opStatus, opNum, sana, amount, currency, opType, operator, contName, kassa, mId, note });
    });
  }

  /* ── 6. GURUHLASH ── */
  const ordersByMId = {};
  rawOrders.forEach((o) => {
    const mid = o.mId || o.contName;
    if (!ordersByMId[mid]) ordersByMId[mid] = [];
    ordersByMId[mid].push(o);
  });
  const cashByMId = {};
  rawCash.forEach((c) => {
    const mid = c.mId || c.contName;
    if (!cashByMId[mid]) cashByMId[mid] = [];
    cashByMId[mid].push(c);
  });

  const assignmentById = {};
  const assignSheet = mainData.sheets['mijozlar'];
  if (assignSheet) {
    assignSheet.slice(1).forEach((r) => {
      const code = String(r[0] || '').trim();
      const id = String(r[1] || '').trim();
      if (!id) return;
      if (!/^\d+$/.test(id)) return;
      let operator = '';
      if (code === 'Op3' || code === '5') operator = 'Dildora';
      if (code === 'Op2' || code === '4') operator = 'Dilfuza';
      if (operator) assignmentById[id] = operator;
    });
  }

  /* ── 7. MIJOZLAR YARATISH ── */
  const customers = contacts.map((c) => {
    const myOrders = ordersByMId[c.id] || [];
    const myCash   = cashByMId[c.id]   || [];

    // Suv yetkazilgan (daromad, sotish statistikasi uchun)
    const waterDelivered = myOrders.filter(
      (o) => isWaterProduct(o.product) && isOrderDoc(o.docType) && isDeliveredStatus(o.status)
    );

    // TARA +: 3 mahsulotdan biri, zakaz turi, yetkazilgan → mijozda tara ko'payadi
    const taraPlus = myOrders.filter(
      (o) => isTaraAffectingProduct(o.product) && isOrderDoc(o.docType) && isDeliveredStatus(o.status)
    );
    // TARA -: 3 mahsulotdan biri, vozvrat turi, qabul qilingan → mijozda tara kamayadi
    const taraMinus = myOrders.filter(
      (o) => isTaraAffectingProduct(o.product) && isReturnDoc(o.docType) && isDeliveredStatus(o.status)
    );

    // Boshlang'ich tara: merchants.eskiId → intigratsiya.A → intigratsiya.H
    const integTara  = c.eskiId ? (integTaraByEskiId[c.eskiId] || 0) : 0;
    const taraPlusQ  = taraPlus.reduce((s, o) => s + Math.abs(o.qty), 0);
    const taraMinusQ = taraMinus.reduce((s, o) => s + Math.abs(o.qty), 0);
    const tara       = integTara + taraPlusQ - taraMinusQ;

    const totalWaterQ     = waterDelivered.reduce((s, o) => s + Math.abs(o.qty), 0);
    const totalWaterS_uzs = waterDelivered.filter((o) => o.currency !== 'USD').reduce((s, o) => s + o.sum, 0);
    const totalWaterS_usd = waterDelivered.filter((o) => o.currency === 'USD').reduce((s, o) => s + o.sum, 0);

    const kulerOrds = myOrders.filter(
      (o) => (o.product || '').toLowerCase().includes('kuler') ||
              (o.product || '').toLowerCase().includes('кулер')
    );
    const kulers = kulerOrds.filter((o) => isOrderDoc(o.docType)).length;

    const sortedWater = [...waterDelivered].sort((a, b) => {
      const da = toDate(a.orderDate), db = toDate(b.orderDate);
      return da && db ? db - da : 0;
    });
    const last = sortedWater[0];
    const days = last ? daysAgo(last.orderDate) : null;

    const balanceUZS = balMapUZS[c.id] ?? 0;
    const balanceUSD = balMapUSD[c.id] ?? 0;

    const totalPaidUZS = myCash
      .filter((p) => isPaymentFromCounterparty(p.opType) && p.currency !== 'USD')
      .reduce((s, p) => s + p.amount, 0);
    const totalPaidUSD = myCash
      .filter((p) => isPaymentFromCounterparty(p.opType) && p.currency === 'USD')
      .reduce((s, p) => s + p.amount, 0);

    return {
      id: c.id,
      name: c.name,
      phone: c.phone || '—',
      district: c.district,
      source: c.source,
      merchantNote: c.merchantNote || '',
      address: c.address,
      balanceUZS,
      balanceUSD,
      balance: balanceUZS,
      tara,
      integTara,
      kulers,
      totalWaterQ,
      totalWaterS_uzs,
      totalWaterS_usd,
      vozvratQ: taraMinusQ,
      totalPaidUZS,
      totalPaidUSD,
      lastOrderDate:   last?.orderDate    || '',
      lastDelivPerson: last?.delivPerson  || '',
      lastDocNum:      last?.soNum        || '',
      lastQty:         last ? Math.abs(last.qty) : 0,
      lastSum:         last?.sum          || 0,
      lastAgent:       last?.agent        || '',
      lastStatus:      last?.status       || '',
      daysAgo: days,
      hasOrders: waterDelivered.length > 0,
      isDebtor: balanceUZS < 0,
    };
  });

  const merchantById = {};
  allMerchants.forEach((m) => { merchantById[m.id] = m; });

  const kulerIds = new Set(
    allMerchants
      .filter((m) => (m.source || '').toLowerCase().includes('kuler'))
      .map((m) => m.id)
  );
  const isKulerProduct = (p) => {
    const pl = (p || '').toLowerCase();
    return pl.includes('kuler') || pl.includes('кулер');
  };

  const deliveredKulerOrders = rawOrders.filter((o) =>
    kulerIds.has(o.mId) &&
    isOrderDoc(o.docType) &&
    isDeliveredStatus(o.status) &&
    isKulerProduct(o.product)
  );
  const returnedKulerOrders = rawOrders.filter((o) =>
    kulerIds.has(o.mId) &&
    isReturnDoc(o.docType) &&
    isDeliveredStatus(o.status) &&
    isKulerProduct(o.product)
  );

  const kulerByCustomer = {};
  deliveredKulerOrders.forEach((o) => {
    const key = o.mId;
    if (!kulerByCustomer[key]) {
      kulerByCustomer[key] = {
        customerId: key,
        plans: [],
        payments: [],
        principal: 0,
        returned: 0,
        paid: 0,
      };
    }
    kulerByCustomer[key].plans.push(o);
    kulerByCustomer[key].principal += Math.abs(o.sum || 0);
  });
  returnedKulerOrders.forEach((o) => {
    const key = o.mId;
    if (!kulerByCustomer[key]) return;
    kulerByCustomer[key].returned += Math.abs(o.sum || 0);
  });
  rawCash.forEach((p) => {
    const key = p.mId;
    if (!kulerByCustomer[key]) return;
    if (!isPaymentFromCounterparty(p.opType)) return;
    if (p.currency === 'USD') return;
    kulerByCustomer[key].paid += Math.abs(p.amount || 0);
    kulerByCustomer[key].payments.push({
      sana: p.sana,
      amount: Math.abs(p.amount || 0),
      operator: p.operator || '',
      opNum: p.opNum || '',
    });
  });

  const kulerInstallments = Object.values(kulerByCustomer).map((k) => {
    const customer = merchantById[k.customerId] || contacts.find((c) => c.id === k.customerId);
    const firstOrder = [...k.plans].sort((a,b) => (toDate(a.orderDate)||0) - (toDate(b.orderDate)||0))[0];
    const months = 6;
    const net = Math.max(0, k.principal - k.returned);
    const monthly = months > 0 ? Math.ceil(net / months) : 0;
    const paid = Math.min(k.paid, net);
    const remaining = Math.max(0, net - paid);
    const paidMonths = monthly > 0 ? Math.floor(paid / monthly) : 0;
    const monthsLeft = monthly > 0 ? Math.ceil(remaining / monthly) : 0;

    let dueCount = 0;
    let overdueAmount = 0;
    if (firstOrder?.orderDate && monthly > 0) {
      const start = toDate(firstOrder.orderDate);
      const now = new Date();
      const elapsed = Math.max(0, (now.getFullYear() - start.getFullYear()) * 12 + (now.getMonth() - start.getMonth()) + 1);
      const shouldPayMonths = Math.min(months, elapsed);
      const shouldPaid = shouldPayMonths * monthly;
      overdueAmount = Math.max(0, Math.min(net, shouldPaid) - paid);
      dueCount = overdueAmount > 0 ? 1 : 0;
    }

    return {
      customerId: k.customerId,
      customerName: customer?.name || `ID: ${k.customerId}`,
      source: customer?.source || 'KULER',
      note: customer?.merchantNote || '',
      purchaseDate: firstOrder?.orderDate || '',
      orderNo: firstOrder?.soNum || '—',
      operator: firstOrder?.agent || '—',
      product: firstOrder?.product || 'Kuler',
      payments: [...k.payments].sort((a,b) => (toDate(a.sana)||0) - (toDate(b.sana)||0)),
      months,
      monthly,
      principal: net,
      paid,
      remaining,
      paidMonths,
      monthsLeft,
      dueCount,
      overdueAmount,
    };
  }).filter((x) => x.principal > 0);

  return {
    customers, orders: rawOrders, cashbox: rawCash, contacts,
    rawOrders, rawCash, ordersByMId, cashByMId, kulerInstallments, assignmentById,
  };
}

/* ═══════════════ SVERKA ═══════════════ */
function buildSverka(customer, ordersByMId, cashByMId) {
  const cid = customer.id;
  const myOrders = (ordersByMId[cid] || []).map((o) => ({
    ...o,
    _type: 'order',
    _date: toDate(o.orderDate) || new Date(0),
    _dateStr: o.orderDate,
  }));
  const myCash = (cashByMId[cid] || [])
    .filter((p) => isPaymentFromCounterparty(p.opType))
    .map((p) => ({
      ...p,
      _type: 'payment',
      _date: toDate(p.sana) || new Date(0),
      _dateStr: p.sana,
    }));

  const combined = [...myOrders, ...myCash].sort((a, b) => a._date - b._date);

  let runBalUZS = 0, runBalUSD = 0;
  return combined.map((row) => {
    if (row._type === 'order') {
      const isVoz  = isReturnDoc(row.docType);
      const isTara = isTaraAffectingProduct(row.product);
      const isWater = isWaterProduct(row.product);
      // ОТМЕНЕНО bo'lmagan va ДОСТАВЛЕН/ПОЛУЧЕН НА СКЛАД bo'lgan → balansga ta'sir qiladi
      const countsForBalance = isDeliveredStatus(row.status);
      const orderSum  = row.sum || 0;
      const isCancelled = isCancelledStatus(row.status);

      if (countsForBalance && !isCancelled) {
        if (row.currency === 'USD') runBalUSD -= orderSum;
        else runBalUZS -= orderSum;
      }
      return {
        kod: row.soNum,
        sana: fmtD(row._dateStr),
        dokument: row.docType || 'Заказ',
        produkt: row.product,
        qty: row.qty,
        narx: row.price,
        summa: countsForBalance && !isCancelled ? orderSum : 0,
        tolov: 0,
        balansUZS: runBalUZS,
        balansUSD: runBalUSD,
        currency: row.currency || 'UZS',
        agent: row.agent,
        note: '',
        driver: row.delivPerson, // X ustun — dostavchik ismi
        status: row.status,
        _isVoz: isVoz,
        _isTara: isTara,
        _isWater: isWater,
        _isCancelled: isCancelled,
        _type: 'order',
      };
    } else {
      if (row.currency === 'USD') runBalUSD += row.amount;
      else runBalUZS += row.amount;
      return {
        kod: row.opNum,
        sana: fmtD(row._dateStr),
        dokument: "Оплата от контрагента",
        produkt: '',
        qty: null,
        narx: null,
        summa: 0,
        tolov: row.amount,
        balansUZS: runBalUZS,
        balansUSD: runBalUSD,
        currency: row.currency || 'UZS',
        agent: row.operator || '',
        note: row.note || '',
        driver: row.kassa || '',
        status: row.opStatus || '',
        _isVoz: false,
        _isTara: false,
        _isWater: false,
        _isCancelled: false,
        _type: 'payment',
      };
    }
  });
}

/* ═══════════════ EXCEL EXPORT ═══════════════ */
function exportSverkaExcel(customer, rows) {
  const wb = XLSX.utils.book_new();
  const headers = [
    'Kod','Sana','Dokument','Mahsulot','Kol-vo','Narx',
    'Summa',"To'lov (UZS)","To'lov (USD)",'Balans UZS','Balans USD',
    'Valyuta','Agent','Dostavchik','Status',
  ];
  const title = [
    [`SVERKA: ${customer.name}`],
    [`ID: ${customer.id}`],
    [`Joriy balans UZS: ${fmt(customer.balanceUZS)} so'm`],
    [`Joriy balans USD: ${fmt(customer.balanceUSD)} $`],
    [`Idish (tara): ${customer.tara} ta`],
    [],
  ];
  const data = [
    ...title,
    headers,
    ...rows.map((r) => [
      r.kod, r.sana, r.dokument, r.produkt,
      r.qty != null ? r.qty : '',
      r.narx ? Math.round(r.narx) : '',
      r.summa || '',
      r._type === 'payment' && r.currency !== 'USD' ? r.tolov : '',
      r._type === 'payment' && r.currency === 'USD'  ? r.tolov : '',
      r.balansUZS, r.balansUSD, r.currency,
      r.agent, r.driver, r.status,
    ]),
  ];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [12,14,22,22,7,12,12,12,12,14,12,8,12,18,14].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws, 'Sverka');
  XLSX.writeFile(wb, `Sverka_${customer.name.slice(0,40).replace(/[/\\:*?"<>|]/g,'_')}.xlsx`);
}
function exportAllReport(customers) {
  const wb = XLSX.utils.book_new();
  const headers = [
    'ID','Mijoz','Telefon','Rayon','Balans UZS','Balans USD',
    'Idish (ta)','Kuler','Jami suv (ta)','Summa UZS','Summa USD',
    'Oxirgi zakaz','Kun','Agent','Manba',
  ];
  const ws = XLSX.utils.aoa_to_sheet([
    headers,
    ...customers.map((c) => [
      c.id, c.name, c.phone, c.district, c.balanceUZS, c.balanceUSD,
      c.tara, c.kulers, c.totalWaterQ, c.totalWaterS_uzs, c.totalWaterS_usd,
      fmtD(c.lastOrderDate), c.daysAgo != null ? c.daysAgo : '', c.lastAgent, c.source,
    ]),
  ]);
  ws['!cols'] = [10,35,14,16,14,12,10,7,12,16,12,14,8,12,14].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws, 'Mijozlar');

  const debtors = customers.filter((c) => c.balanceUZS < 0 || c.balanceUSD < 0)
    .sort((a,b) => a.balanceUZS - b.balanceUZS);
  const dh = ['ID','Mijoz','Telefon','Rayon','Qarz UZS','Qarz USD','Idish','Oxirgi zakaz','Kun','Agent'];
  const ws2 = XLSX.utils.aoa_to_sheet([
    dh,
    ...debtors.map((c) => [
      c.id, c.name, c.phone, c.district, c.balanceUZS, c.balanceUSD,
      c.tara, fmtD(c.lastOrderDate), c.daysAgo != null ? c.daysAgo : '', c.lastAgent,
    ]),
  ]);
  ws2['!cols'] = [10,35,14,16,14,12,8,12,6,12].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws2, 'Qarzdorlar');
  XLSX.writeFile(wb, `AquaBiz_Hisobot_${new Date().toISOString().slice(0,10)}.xlsx`);
}

/* ═══════════════ CSS ═══════════════ */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500&display=swap');
:root{
  --bg:#080c10;--s1:#0d1117;--s2:#161b22;--s3:#1c2128;
  --b1:#30363d;--b2:#21262d;
  --t1:#e6edf3;--t2:#8b949e;--t3:#656d76;--t4:#3d444d;
  --bl:#58a6ff;--bl2:#1f3d6b;--bl3:#0d1f3c;
  --gr:#3fb950;--gr2:#1a3a24;--gr3:#0d1f14;
  --rd:#f85149;--rd2:#3d1b1a;
  --yl:#d29922;--yl2:#3a2d0d;
  --or:#f0883e;--or2:#3d2210;
  --pu:#bc8cff;--pu2:#2d1f4e;
  --r:8px;--rl:12px;
  --sans:'IBM Plex Sans',sans-serif;
  --mono:'IBM Plex Mono',monospace;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden;background:var(--bg);color:var(--t1)}
body,input,select,button{font-family:var(--sans)}
::-webkit-scrollbar{width:10px;height:10px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:var(--b1);border-radius:6px}
.btn{cursor:pointer;border:none;font-size:12.5px;font-weight:600;border-radius:var(--r);
  padding:7px 13px;transition:all .13s;display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
.btn-bl{background:var(--bl);color:#000}.btn-bl:hover{opacity:.85}
.btn-gr{background:var(--gr);color:#000}.btn-gr:hover{opacity:.85}
.btn-gh{background:transparent;color:var(--t2);border:1px solid var(--b1)}
.btn-gh:hover{background:var(--s3);color:var(--t1)}
.btn-sm{padding:5px 10px;font-size:11.5px}
.input{width:100%;padding:8px 11px;border:1px solid var(--b1);border-radius:var(--r);
  font-size:13px;color:var(--t1);background:var(--s3);outline:none}
.input:focus{border-color:var(--bl);box-shadow:0 0 0 3px rgba(88,166,255,.1)}
.select{padding:7px 10px;border:1px solid var(--b1);border-radius:var(--r);
  font-size:12.5px;color:var(--t1);background:var(--s3);cursor:pointer;outline:none}
.select:focus{border-color:var(--bl)}
.card{background:var(--s1);border:1px solid var(--b2);border-radius:var(--rl)}
.tag{display:inline-flex;align-items:center;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600;font-family:var(--mono)}
.modal-ov{position:fixed;inset:0;background:rgba(0,0,0,.85);display:flex;align-items:center;justify-content:center;z-index:300;padding:16px;backdrop-filter:blur(5px)}
.modal{background:var(--s1);border:1px solid var(--b1);border-radius:var(--rl);width:100%;max-width:900px;max-height:94vh;overflow-y:auto;box-shadow:0 32px 80px rgba(0,0,0,.8)}
.mhdr{padding:14px 20px;border-bottom:1px solid var(--b2);display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;background:var(--s1);z-index:1}
.mbdy{padding:16px 20px}
@keyframes up{from{transform:translateY(10px);opacity:0}to{transform:translateY(0);opacity:1}}
@keyframes fd{from{opacity:0}to{opacity:1}}
@keyframes spin{to{transform:rotate(360deg)}}
.ani{animation:up .2s cubic-bezier(.2,.8,.2,1)}
.fade{animation:fd .15s ease}
.tbl{width:100%;border-collapse:separate;border-spacing:0;font-size:12px}
.tbl th{background:var(--s2);color:var(--t3);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;padding:7px 8px;text-align:left;border-bottom:1px solid var(--b1);position:sticky;top:0;z-index:2;white-space:nowrap;user-select:none;cursor:pointer}
.tbl th:hover,.tbl th.act{color:var(--bl)}
.tbl td{padding:7px 8px;border-bottom:1px solid var(--b2);vertical-align:middle;color:var(--t2);white-space:nowrap}
.tbl tbody tr:hover td{background:var(--s2);cursor:pointer}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.g3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
.g4{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
.g5{display:grid;grid-template-columns:repeat(5,1fr);gap:10px}
.stat{background:var(--s1);border:1px solid var(--b2);border-radius:var(--rl);padding:14px 16px;position:relative;overflow:hidden}
.sb{display:flex;align-items:center;gap:8px;background:var(--s3);border:1px solid var(--b1);border-radius:var(--r);padding:7px 11px}
.sb input{border:none;background:transparent;outline:none;color:var(--t1);font-size:13px;width:100%}
.drop-z{border:2px dashed var(--b1);border-radius:var(--rl);padding:16px 18px;cursor:pointer;transition:all .14s;background:var(--s3)}
.drop-z:hover,.drop-z.drag{border-color:var(--bl);background:var(--bl3)}
.drop-z.done{border-color:var(--gr);background:var(--gr3)}
.tabs{display:flex;gap:2px;background:var(--s2);border-radius:var(--r);padding:3px}
.tab{padding:5px 13px;border-radius:6px;cursor:pointer;font-size:12.5px;font-weight:600;color:var(--t2);transition:all .13s;border:none;background:none}
.tab.on{background:var(--s1);color:var(--t1);box-shadow:0 1px 4px rgba(0,0,0,.4)}
.nav-i{display:flex;align-items:center;gap:9px;padding:8px 10px;border-radius:var(--r);cursor:pointer;color:var(--t2);font-size:13px;font-weight:500;transition:all .13s;user-select:none;position:relative}
.nav-i:hover{background:var(--s2);color:var(--t1)}
.nav-i.on{background:var(--bl2);color:var(--bl)}
.notif{position:fixed;bottom:18px;right:18px;z-index:999;padding:10px 16px;border-radius:var(--r);display:flex;align-items:center;gap:8px;font-size:13px;font-weight:600;box-shadow:0 4px 24px rgba(0,0,0,.6);animation:up .2s ease}
.sv-tbl td{font-size:11.5px}
.sv-tbl tr.payment td{background:rgba(63,185,80,.04)}
.sv-tbl tr.vozvrat td{background:rgba(240,136,62,.04)}
.sv-tbl tr.cancelled td{background:rgba(248,81,73,.04);opacity:.6}
.cur-uzs{color:var(--gr);font-size:10px;font-family:var(--mono)}
.cur-usd{color:var(--yl);font-size:10px;font-family:var(--mono)}
.dot{display:inline-block;width:8px;height:8px;border-radius:50%;flex-shrink:0}
.toggle{position:relative;width:36px;height:20px;border-radius:10px;background:var(--b1);border:1px solid var(--b2);cursor:pointer;transition:all .16s}
.toggle.on{background:var(--gr2);border-color:var(--gr)}
.toggle .knob{position:absolute;top:1px;left:1px;width:16px;height:16px;border-radius:50%;background:#fff;transition:left .16s}
.toggle.on .knob{left:17px}
`;

/* ═══════════════ LOCAL STORAGE ═══════════════ */
const S = {
  get: (k,d) => { try { const v=localStorage.getItem(k); return v?JSON.parse(v):d; } catch { return d; } },
  set: (k,v) => { try { localStorage.setItem(k,JSON.stringify(v)); } catch {} },
};
const DEFAULT_USERS = ['Dildora', 'Dilfuza', 'Admin'];
const DEFAULT_USER_CREDS = { Dildora:'Dildora', Dilfuza:'Dilfuza', Admin:'12345' };
const DEFAULT_ACCESS = {
  Dildora: { scope: 'own', visible: { dash:true, cust:true, orders:true, kassa:true, obzvon:true, doljniki:true, reports:true, settings:false } },
  Dilfuza: { scope: 'own', visible: { dash:true, cust:true, orders:true, kassa:true, obzvon:true, doljniki:true, reports:true, settings:false } },
  Admin:   { scope: 'all', visible: { dash:true, cust:true, orders:true, kassa:true, obzvon:true, doljniki:true, reports:true, settings:true } },
};

/* ═══════════════ UPLOAD MODAL ═══════════════ */
function UploadModal({
  onLoad,
  onLoadObzvonAll,
  hasData,
  onClose,
  mainSheetUrl,
  setMainSheetUrl,
  obzvonSheetUrl,
  setObzvonSheetUrl,
}) {
  const [tab, setTab]   = useState('excel');
  const [files, setFiles] = useState({ main: null, integration: null });
  const [loading, setLoad] = useState(false);
  const [progress, setProg] = useState('');
  const [gids, setGids] = useState(() => S.get('aq-gs-gids', {
    merchants:'', balans:'', orders:'', cashbox:'', integration:'', mijozlar:'',
  }));
  const fixedMainUrl = SHEET_CONFIG.url || '';
  const fixedObzvonUrl = OBZVON_ALL_SHEET_URL || '';
  const sheetId = extractSheetId(fixedMainUrl);

  const pickFile = (key) => {
    const i = document.createElement('input');
    i.type='file'; i.accept='.xlsx,.xls';
    i.onchange = (e) => setFiles((p) => ({ ...p, [key]: e.target.files[0] }));
    i.click();
  };

  const doLoadExcel = async () => {
    if (!files.main) { alert('Asosiy faylni tanlang!'); return; }
    setLoad(true);
    try {
      setProg("Asosiy fayl o'qilmoqda...");
      const main = await readExcelFile(files.main);
      if (!main.sheets['public.view_merchants'])
        throw new Error(`Noto'g'ri fayl! Varaqlar: ${Object.keys(main.sheets).join(', ')}`);

      let integSheets = {};
      if (files.integration) {
        setProg("Integratsiya fayli o'qilmoqda...");
        const integ = await readExcelFile(files.integration);
        const firstSheet = Object.values(integ.sheets)[0];
        if (firstSheet) integSheets['intigratsiya'] = firstSheet; // nom: intigratsiya
      }

      setProg('Hisoblanmoqda...');
      setTimeout(() => {
        try {
          const merged = { sheets: { ...main.sheets, ...integSheets } };
          const data = processAll(merged);
          onLoad(data);
          setProg('');
        } catch (e2) { alert('Hisoblash xatosi: '+e2.message); setLoad(false); setProg(''); }
      }, 50);
    } catch (e) { alert('Xato: '+e.message); setLoad(false); setProg(''); }
  };

  const doLoadSheets = async () => {
    const usedMainUrl = fixedMainUrl || mainSheetUrl || '';
    const usedObzvonUrl = fixedObzvonUrl || obzvonSheetUrl || '';
    const usedSheetId = extractSheetId(usedMainUrl);
    const usedObzvonSheetId = extractSheetId(usedObzvonUrl);
    if (!usedSheetId) { alert("Google Sheets URL kodda sozlanmagan!"); return; }
    setLoad(true);
    try {
      if (setMainSheetUrl) setMainSheetUrl(usedMainUrl);
      if (setObzvonSheetUrl) setObzvonSheetUrl(usedObzvonUrl);
      S.set('aq-main-url', usedMainUrl);
      S.set('aq-obzvon-url', usedObzvonUrl);
      const raw = await loadFromGoogleSheets(usedSheetId, gids, setProg, 'named');
      setProg('Hisoblanmoqda...');
      setTimeout(() => {
        try {
          const data = processAll(raw);
          onLoad(data);
          if (usedObzvonSheetId && onLoadObzvonAll) onLoadObzvonAll(usedObzvonUrl);
          setProg('');
        } catch (e2) { alert('Hisoblash xatosi: '+e2.message); setLoad(false); setProg(''); }
      }, 50);
    } catch (e) { alert('Xato: '+e.message); setLoad(false); setProg(''); }
  };

  return (
    <div className="modal-ov fade"
      onClick={(e) => hasData && onClose && e.target===e.currentTarget && onClose()}>
      <div className="modal ani" style={{maxWidth:600}}>
        <div className="mhdr">
          <div>
            <div style={{fontWeight:800,fontSize:16}}>📂 Ma'lumot yuklash</div>
            <div style={{fontSize:12,color:'var(--t3)',marginTop:3}}>Excel yoki Google Sheets dan</div>
          </div>
          {hasData && onClose && <button className="btn btn-gh btn-sm" onClick={onClose}>✕</button>}
        </div>
        <div className="mbdy">
          <div className="tabs" style={{marginBottom:18}}>
            {[['excel','📊 Excel fayl'],['sheets','🔗 Google Sheets']].map(([t,l]) => (
              <button key={t} className={`tab${tab===t?' on':''}`} onClick={()=>setTab(t)}>{l}</button>
            ))}
          </div>

          {tab==='excel' && (
            <>
              <div style={{marginBottom:12}}>
                <div style={{fontSize:11.5,fontWeight:700,color:'var(--t3)',textTransform:'uppercase',letterSpacing:'.04em',marginBottom:8}}>
                  Asosiy fayl (majburiy)
                </div>
                <div className={`drop-z${files.main?' done':''}`}
                  onClick={()=>!loading&&pickFile('main')} style={{padding:'20px',textAlign:'center'}}>
                  <div style={{fontSize:28,marginBottom:6}}>{files.main?'✅':'📊'}</div>
                  {files.main
                    ? <div style={{fontWeight:700,color:'var(--gr)',fontSize:13}}>{files.main.name}</div>
                    : <div>
                        <div style={{fontWeight:700,fontSize:13,marginBottom:4}}>Asosiy Excel faylni tanlang</div>
                        <div style={{fontSize:11.5,color:'var(--t3)'}}>public.view_merchants, Balans, Zakazlar, Kassa varaqlari</div>
                      </div>
                  }
                </div>
              </div>
              <div style={{marginBottom:14}}>
                <div style={{fontSize:11.5,fontWeight:700,color:'var(--t3)',textTransform:'uppercase',letterSpacing:'.04em',marginBottom:8}}>
                  Integratsiya fayli (ixtiyoriy — eski baza ostatkasi)
                </div>
                <div className={`drop-z${files.integration?' done':''}`}
                  onClick={()=>!loading&&pickFile('integration')} style={{padding:'16px',textAlign:'center'}}>
                  <div style={{fontSize:22,marginBottom:4}}>{files.integration?'✅':'🔗'}</div>
                  {files.integration
                    ? <div style={{fontWeight:700,color:'var(--gr)',fontSize:13}}>{files.integration.name}</div>
                    : <div>
                        <div style={{fontWeight:700,fontSize:13,marginBottom:3}}>Integratsiya faylini tanlang</div>
                        <div style={{fontSize:11,color:'var(--t3)'}}>A=eski ID · H=tara ostatkasi · Varaq nomi: "intigratsiya"</div>
                      </div>
                  }
                </div>
              </div>
              {loading && (
                <div style={{background:'var(--bl3)',border:'1px solid var(--bl2)',borderRadius:8,padding:'9px 14px',marginBottom:12,display:'flex',alignItems:'center',gap:10,color:'var(--bl)'}}>
                  <span style={{display:'inline-block',animation:'spin .7s linear infinite'}}>⟳</span>
                  <span style={{fontSize:13,fontWeight:600}}>{progress}</span>
                </div>
              )}
              <div style={{display:'flex',justifyContent:'flex-end',gap:8}}>
                {hasData && onClose && <button className="btn btn-gh" onClick={onClose}>Bekor</button>}
                <button className="btn btn-bl" onClick={doLoadExcel} disabled={loading||!files.main} style={{minWidth:140}}>
                  {loading?'⟳ Yuklanmoqda...':'✓ Yuklash'}
                </button>
              </div>
            </>
          )}

          {tab==='sheets' && (
            <>
              <div style={{marginBottom:14}}>
                <div style={{fontSize:11.5,fontWeight:700,color:'var(--t3)',textTransform:'uppercase',letterSpacing:'.04em',marginBottom:6}}>
                  Asosiy ma'lumot fayli URL
                </div>
                <div className="input" style={{fontFamily:'var(--mono)',fontSize:12,padding:'10px 11px',whiteSpace:'nowrap',overflow:'hidden',textOverflow:'ellipsis'}}>
                  {fixedMainUrl || "URL topilmadi"}
                </div>
                <div style={{fontSize:11,color:'var(--t3)',marginTop:5}}>
                  Bu URL avtomatik ishlatiladi. Barcha listlar bir faylda bo'lishi kerak.
                </div>
                {sheetId && <div style={{fontSize:11,color:'var(--gr)',marginTop:5,fontFamily:'var(--mono)'}}>✓ ID: {sheetId}</div>}
              </div>
              <div style={{background:'var(--s3)',border:'1px solid var(--b1)',borderRadius:8,padding:'10px 12px',marginBottom:12,fontSize:12,color:'var(--t2)'}}>
                Kerakli listlar:
                <div style={{marginTop:6,fontFamily:'var(--mono)',fontSize:11}}>
                  public.view_merchants, public.view_current_merchant_event_balance, public.view_item_basket, public.view_cashbox_documents, intigratsiya, mijozlar
                </div>
              </div>
              <div style={{marginBottom:14}}>
                <div style={{fontSize:11.5,fontWeight:700,color:'var(--t3)',textTransform:'uppercase',letterSpacing:'.04em',marginBottom:6}}>
                  Obzvon fayli URL
                </div>
                <input className="input" style={{fontFamily:'var(--mono)',fontSize:12}}
                  placeholder="https://docs.google.com/spreadsheets/d/..."
                  value={fixedObzvonUrl} readOnly/>
                <div style={{fontSize:11,color:'var(--t3)',marginTop:5}}>
                  "Обзвон ВСЕ" listi shu faylda bo'lishi kerak
                </div>
              </div>
              <div style={{background:'var(--yl2)',border:'1px solid var(--yl)',borderRadius:8,padding:'9px 14px',marginBottom:14,fontSize:12,color:'var(--yl)'}}>
                ⚠️ Fayl <strong>Fayl → Veb-da nashr qilish</strong> qilingan bo'lishi kerak
              </div>
              {loading && (
                <div style={{background:'var(--bl3)',border:'1px solid var(--bl2)',borderRadius:8,padding:'9px 14px',marginBottom:12,display:'flex',alignItems:'center',gap:10,color:'var(--bl)'}}>
                  <span style={{display:'inline-block',animation:'spin .7s linear infinite'}}>⟳</span>
                  <span style={{fontSize:13,fontWeight:600}}>{progress}</span>
                </div>
              )}
              <div style={{display:'flex',justifyContent:'flex-end',gap:8}}>
                {hasData && onClose && <button className="btn btn-gh" onClick={onClose}>Bekor</button>}
                <button className="btn btn-bl" onClick={doLoadSheets} disabled={loading||!sheetId} style={{minWidth:160}}>
                  {loading?'⟳ Ulanmoqda...':'🔗 Ulash'}
                </button>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

/* ═══════════════ CUSTOMER DETAIL ═══════════════ */
function CustomerDetail({ c, D, onClose }) {
  const { ordersByMId, cashByMId } = D;
  const [activeTab, setTab] = useState('sverka');

  const sverkaRows = useMemo(
    () => buildSverka(c, ordersByMId || {}, cashByMId || {}),
    [c, ordersByMId, cashByMId]
  );
  const finalBalUZS = sverkaRows.length > 0 ? sverkaRows[sverkaRows.length-1].balansUZS : 0;
  const finalBalUSD = sverkaRows.length > 0 ? sverkaRows[sverkaRows.length-1].balansUSD : 0;

  const myOrders = (ordersByMId?.[c.id] || []).filter(
    (o) => isWaterProduct(o.product) && isOrderDoc(o.docType)
  );
  const myReturns = (ordersByMId?.[c.id] || []).filter(
    (o) => isWaterProduct(o.product) && isReturnDoc(o.docType)
  );
  const myPays = (cashByMId?.[c.id] || []).filter(
    (p) => isPaymentFromCounterparty(p.opType)
  );

  const balColor = (v) => v < 0 ? 'var(--rd)' : v > 0 ? 'var(--gr)' : 'var(--t3)';
  const taraColor = c.tara < 0 ? 'var(--rd)' : 'var(--bl)';

  return (
    <div className="modal-ov fade" onClick={(e)=>e.target===e.currentTarget&&onClose()}>
      <div className="modal ani" style={{maxWidth:1000}}>
        <div className="mhdr">
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontWeight:700,fontSize:14,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
              {c.name}
            </div>
            <div style={{fontSize:11,color:'var(--t3)',fontFamily:'var(--mono)',marginTop:2}}>
              ID: {c.id} · {c.district} · {c.phone}
            </div>
          </div>
          <div style={{display:'flex',gap:8,flexShrink:0,marginLeft:12}}>
            <button className="btn btn-gr btn-sm" onClick={()=>exportSverkaExcel(c,sverkaRows)}>⬇ Excel</button>
            <button className="btn btn-gh btn-sm" onClick={onClose}>✕</button>
          </div>
        </div>
        <div className="mbdy">
          <div style={{display:'grid',gridTemplateColumns:'repeat(6,1fr)',gap:10,marginBottom:14}}>
            {[
              { l:'BALANS UZS', v:fmt(c.balanceUZS)+" so'm", col:balColor(c.balanceUZS) },
              { l:'BALANS USD', v:fmt(c.balanceUSD)+' $',    col:balColor(c.balanceUSD) },
              { l:'IDISH (TARA)', v:(c.tara<0?'-':'')+Math.abs(c.tara)+' ta', col:taraColor,
                s:c.integTara>0?`Boshlang'ich: ${c.integTara}`:null },
              { l:'JAMI ZAKAZ',  v:c.totalWaterQ+' ta', col:'var(--pu)' },
              { l:"TO'LOV UZS", v:fmt(c.totalPaidUZS)+" so'm", col:'var(--gr)' },
              { l:"TO'LOV USD", v:fmt(c.totalPaidUSD)+' $',    col:'var(--yl)' },
            ].map((s,i) => (
              <div key={i} style={{background:'var(--s3)',borderRadius:9,padding:'10px 12px',textAlign:'center'}}>
                <div style={{fontSize:9,color:'var(--t3)',fontWeight:700,textTransform:'uppercase',marginBottom:4}}>{s.l}</div>
                <div style={{fontSize:15,fontWeight:800,color:s.col,fontFamily:'var(--mono)'}}>{s.v}</div>
                {s.s && <div style={{fontSize:9.5,color:'var(--t3)',marginTop:2}}>{s.s}</div>}
              </div>
            ))}
          </div>

          <div className="tabs" style={{marginBottom:14,display:'inline-flex'}}>
            {[['sverka','📋 Sverka'],['orders','📦 Zakazlar'],['returns','↩ Vozvratlar'],['pays',"💰 To'lovlar"]].map(([t,l]) => (
              <button key={t} className={`tab${activeTab===t?' on':''}`} onClick={()=>setTab(t)}>{l}</button>
            ))}
          </div>

          {activeTab==='sverka' && (
            <div>
              <div style={{background:'var(--s3)',border:'1px solid var(--b1)',borderRadius:8,padding:'8px 14px',marginBottom:12,fontSize:12,color:'var(--t3)',display:'flex',gap:16,flexWrap:'wrap'}}>
                <span>📦 Yetkazilgan: <strong style={{color:'var(--bl)'}}>{c.totalWaterQ} ta</strong></span>
                <span>↩ Qaytarilgan: <strong style={{color:'var(--or)'}}>{c.vozvratQ} ta</strong></span>
                <span>🏠 Mijozda: <strong style={{color:c.tara<0?'var(--rd)':'var(--gr)'}}>{c.tara} ta</strong></span>
                <span>💴 Balans UZS: <strong style={{color:balColor(finalBalUZS)}}>{fmt(Math.abs(finalBalUZS))} so'm</strong></span>
                <span>💵 Balans USD: <strong style={{color:balColor(finalBalUSD)}}>{fmt(Math.abs(finalBalUSD))} $</strong></span>
              </div>
              <div style={{overflow:'auto',maxHeight:'52vh',borderRadius:9,border:'1px solid var(--b2)'}}>
                <table className="tbl sv-tbl" style={{minWidth:1000}}>
                  <thead>
                    <tr>
                      <th style={{minWidth:80}}>Kod</th>
                      <th style={{minWidth:85}}>Sana</th>
                      <th style={{minWidth:90}}>Hujjat</th>
                      <th style={{minWidth:170}}>Mahsulot</th>
                      <th style={{minWidth:55,textAlign:'center'}}>Kol-vo</th>
                      <th style={{minWidth:80,textAlign:'right'}}>Narx</th>
                      <th style={{minWidth:95,textAlign:'right'}}>Summa</th>
                      <th style={{minWidth:65,textAlign:'center'}}>Valyuta</th>
                      <th style={{minWidth:90,textAlign:'right'}}>To'lov</th>
                      <th style={{minWidth:100,textAlign:'right'}}>Balans UZS</th>
                      <th style={{minWidth:90,textAlign:'right'}}>Balans USD</th>
                      <th style={{minWidth:80}}>Agent</th>
                      <th style={{minWidth:110}}>Dostavchik</th>
                      <th style={{minWidth:85}}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sverkaRows.length===0 ? (
                      <tr><td colSpan={14} style={{textAlign:'center',padding:32,color:'var(--t3)'}}>Ma'lumot yo'q</td></tr>
                    ) : (
                      sverkaRows.map((row,i) => (
                        <tr key={i} className={row._type==='payment'?'payment':row._isVoz?'vozvrat':row._isCancelled?'cancelled':''}>
                          <td style={{fontFamily:'var(--mono)',fontSize:10.5,color:'var(--t3)'}}>{row.kod}</td>
                          <td style={{fontFamily:'var(--mono)',fontSize:11}}>{row.sana}</td>
                          <td>
                            {row._type==='payment'
                              ? <span className="tag" style={{background:'var(--gr2)',color:'var(--gr)',fontSize:10}}>To'lov</span>
                              : row._isCancelled
                              ? <span className="tag" style={{background:'var(--rd2)',color:'var(--t3)',fontSize:10}}>Otmena</span>
                              : row._isVoz
                              ? <span className="tag" style={{background:'var(--or2)',color:'var(--or)',fontSize:10}}>Vozvrat</span>
                              : <span className="tag" style={{background:'var(--bl3)',color:'var(--bl)',fontSize:10}}>Zakaz</span>
                            }
                          </td>
                          <td style={{maxWidth:175,color:'var(--t1)',fontWeight:row._type==='payment'?400:500}}>
                            <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:173}}>
                              {row.produkt||'—'}
                            </span>
                          </td>
                          <td style={{textAlign:'center',fontWeight:700,color:row.qty!=null?(row._isVoz?'var(--or)':row.qty>0?'var(--t1)':'var(--t3)'):'var(--t3)'}}>
                            {row.qty!=null?(row._isVoz?'-':'')+Math.abs(row.qty):'—'}
                          </td>
                          <td style={{textAlign:'right',fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>
                            {row.narx?fmt(Math.round(row.narx)):'—'}
                          </td>
                          <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700,color:row.summa?'var(--rd)':'var(--t4)'}}>
                            {row.summa?'-'+fmt(row.summa):'—'}
                          </td>
                          <td style={{textAlign:'center'}}>
                            <span className={row.currency==='USD'?'cur-usd':'cur-uzs'}>{row.currency||'UZS'}</span>
                          </td>
                          <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700,color:row.tolov?'var(--gr)':'var(--t4)'}}>
                            {row.tolov?'+'+fmt(row.tolov):'—'}
                          </td>
                          <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:800,color:row.balansUZS<0?'var(--rd)':row.balansUZS>0?'var(--gr)':'var(--t3)'}}>
                            {row.balansUZS<0?'-':row.balansUZS>0?'+':''}{fmt(Math.abs(row.balansUZS))}
                          </td>
                          <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700,color:row.balansUSD<0?'var(--rd)':row.balansUSD>0?'var(--gr)':'var(--t3)',fontSize:11}}>
                            {row.balansUSD<0?'-':row.balansUSD>0?'+':''}{fmt(Math.abs(row.balansUSD))}$
                          </td>
                          <td style={{fontSize:11,color:'var(--t3)'}}>{row.agent||'—'}</td>
                          <td style={{fontSize:11,color:'var(--t2)'}}>{row.driver||'—'}</td>
                          <td>
                            <span className="tag" style={{
                              background: row.status==='ДОСТАВЛЕН'?'var(--gr2)':row.status==='ПОЛУЧЕН НА СКЛАД'?'var(--gr2)':row.status==='ОТМЕНЕНО'?'var(--rd2)':'var(--s3)',
                              color:      row.status==='ДОСТАВЛЕН'?'var(--gr)':row.status==='ПОЛУЧЕН НА СКЛАД'?'var(--gr)':row.status==='ОТМЕНЕНО'?'var(--t4)':'var(--t3)',
                              fontSize:9.5,
                            }}>
                              {row.status||'—'}
                            </span>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
              {sverkaRows.length>0 && (
                <div style={{marginTop:10,fontSize:12,color:'var(--t3)',textAlign:'right'}}>
                  Jami {sverkaRows.length} ta yozuv ·{' '}
                  {sverkaRows.filter((r)=>r._type==='order'&&!r._isCancelled).length} ta zakaz ·{' '}
                  {sverkaRows.filter((r)=>r._type==='payment').length} ta to'lov
                  {sverkaRows.some((r)=>r._isCancelled) && (
                    <span style={{color:'var(--rd)',marginLeft:8}}>
                      · {sverkaRows.filter((r)=>r._isCancelled).length} ta otmena
                    </span>
                  )}
                </div>
              )}
            </div>
          )}

          {activeTab==='orders' && (
            <div style={{overflow:'auto',maxHeight:'60vh',borderRadius:9,border:'1px solid var(--b2)'}}>
              <table className="tbl">
                <thead>
                  <tr>
                    <th>Zakaz №</th><th>Sana</th><th>Mahsulot</th>
                    <th>Kol-vo</th><th>Narx</th><th>Summa</th>
                    <th>Valyuta</th><th>Agent</th><th>Dostavchik</th><th>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {myOrders.length===0
                    ? <tr><td colSpan={10} style={{textAlign:'center',padding:32,color:'var(--t3)'}}>Zakazlar yo'q</td></tr>
                    : myOrders.map((o,i) => (
                      <tr key={i}>
                        <td style={{fontFamily:'var(--mono)',fontSize:10.5,color:'var(--t3)'}}>{o.soNum}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(o.orderDate)}</td>
                        <td style={{maxWidth:180,fontWeight:500}}>
                          <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:178}}>{o.product}</span>
                        </td>
                        <td style={{textAlign:'center',fontWeight:700}}>{Math.abs(o.qty)}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{o.price?fmt(Math.round(o.price)):'—'}</td>
                        <td style={{fontFamily:'var(--mono)',fontWeight:700,color:'var(--gr)'}}>{o.sum?fmt(o.sum):'—'}</td>
                        <td><span className={o.currency==='USD'?'cur-usd':'cur-uzs'}>{o.currency||'UZS'}</span></td>
                        <td style={{fontSize:11}}>{o.agent||'—'}</td>
                        <td style={{fontSize:11,color:'var(--t2)'}}>{o.delivPerson||'—'}</td>
                        <td>
                          <span className="tag" style={{
                            background:o.status==='ДОСТАВЛЕН'||o.status==='ПОЛУЧЕН НА СКЛАД'?'var(--gr2)':o.status==='ОТМЕНЕНО'?'var(--rd2)':'var(--s3)',
                            color:o.status==='ДОСТАВЛЕН'||o.status==='ПОЛУЧЕН НА СКЛАД'?'var(--gr)':o.status==='ОТМЕНЕНО'?'var(--t4)':'var(--t3)',
                            fontSize:10,
                          }}>{o.status}</span>
                        </td>
                      </tr>
                    ))
                  }
                </tbody>
              </table>
            </div>
          )}

          {activeTab==='returns' && (
            <div style={{overflow:'auto',maxHeight:'60vh',borderRadius:9,border:'1px solid var(--b2)'}}>
              <table className="tbl">
                <thead>
                  <tr>
                    <th>Vozvrat №</th><th>Sana</th><th>Mahsulot</th>
                    <th>Kol-vo</th><th>Narx</th><th>Summa</th>
                    <th>Valyuta</th><th>Agent</th><th>Dostavchik</th><th>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {myReturns.length===0
                    ? <tr><td colSpan={10} style={{textAlign:'center',padding:32,color:'var(--t3)'}}>Vozvratlar yo'q</td></tr>
                    : myReturns.map((o,i) => (
                      <tr key={i}>
                        <td style={{fontFamily:'var(--mono)',fontSize:10.5,color:'var(--t3)'}}>{o.soNum}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(o.orderDate)}</td>
                        <td style={{maxWidth:180,fontWeight:500}}>
                          <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:178}}>{o.product}</span>
                        </td>
                        <td style={{textAlign:'center',fontWeight:700,color:'var(--or)'}}>-{Math.abs(o.qty)}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{o.price?fmt(Math.round(o.price)):'—'}</td>
                        <td style={{fontFamily:'var(--mono)',fontWeight:700,color:'var(--or)'}}>{o.sum?fmt(Math.abs(o.sum)):'—'}</td>
                        <td><span className={o.currency==='USD'?'cur-usd':'cur-uzs'}>{o.currency||'UZS'}</span></td>
                        <td style={{fontSize:11}}>{o.agent||'—'}</td>
                        <td style={{fontSize:11,color:'var(--t2)'}}>{o.delivPerson||'—'}</td>
                        <td>
                          <span className="tag" style={{
                            background:o.status==='ДОСТАВЛЕН'||o.status==='ПОЛУЧЕН НА СКЛАД'?'var(--gr2)':o.status==='ОТМЕНЕНО'?'var(--rd2)':'var(--s3)',
                            color:o.status==='ДОСТАВЛЕН'||o.status==='ПОЛУЧЕН НА СКЛАД'?'var(--gr)':o.status==='ОТМЕНЕНО'?'var(--t4)':'var(--t3)',
                            fontSize:10,
                          }}>{o.status}</span>
                        </td>
                      </tr>
                    ))
                  }
                </tbody>
              </table>
            </div>
          )}

          {activeTab==='pays' && (
            <div style={{overflow:'auto',maxHeight:'60vh',borderRadius:9,border:'1px solid var(--b2)'}}>
              <table className="tbl">
                <thead>
                  <tr><th>Op №</th><th>Sana</th><th>Summa</th><th>Valyuta</th><th>Kassa</th><th>Operator</th><th>Izoh</th></tr>
                </thead>
                <tbody>
                  {myPays.length===0
                    ? <tr><td colSpan={7} style={{textAlign:'center',padding:32,color:'var(--t3)'}}>To'lovlar yo'q</td></tr>
                    : myPays.map((p,i) => (
                      <tr key={i}>
                        <td style={{fontFamily:'var(--mono)',fontSize:10.5,color:'var(--t3)'}}>{p.opNum}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(p.sana)}</td>
                        <td style={{fontFamily:'var(--mono)',fontWeight:800,color:'var(--gr)'}}>+{fmt(p.amount)}</td>
                        <td><span className={p.currency==='USD'?'cur-usd':'cur-uzs'}>{p.currency||'UZS'}</span></td>
                        <td style={{fontSize:11,color:'var(--t3)'}}>{p.kassa||'—'}</td>
                        <td style={{fontSize:11}}>{p.operator||'—'}</td>
                        <td style={{fontSize:11,color:'var(--t3)'}}>{p.note||'—'}</td>
                      </tr>
                    ))
                  }
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* ═══════════════ STAT CARD ═══════════════ */
function StatCard({ l, v, s, c }) {
  return (
    <div className="stat" style={{padding:'13px 15px'}}>
      <div style={{fontSize:10,fontWeight:700,color:'var(--t3)',textTransform:'uppercase',letterSpacing:'.05em'}}>{l}</div>
      <div style={{fontSize:18,fontWeight:800,color:c||'var(--t1)',margin:'5px 0 3px'}}>{v}</div>
      {s && <div style={{fontSize:11,color:'var(--t3)'}}>{s}</div>}
      <div style={{position:'absolute',bottom:0,left:0,right:0,height:2,background:c||'var(--bl)',opacity:.4,borderRadius:'0 0 12px 12px'}}/>
    </div>
  );
}

/* ═══════════════ DASHBOARD ═══════════════ */
function Dashboard({ D }) {
  const { customers, cashbox, rawOrders } = D;
  const now = new Date();
  const curMonthKey = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const wDel = (rawOrders||[]).filter((o) => isWaterProduct(o.product) && isOrderDoc(o.docType) && isDeliveredStatus(o.status));
  const wDelThisMonth = wDel.filter((o) => monthKey(o.orderDate)===curMonthKey);
  const pays = (cashbox||[]).filter((c) => isPaymentFromCounterparty(c.opType));
  const debtors = customers.filter((c) => c.balanceUZS < 0);
  const debt = getDebtStats(customers);
  const stale2m = customers
    .filter((c) => c.hasOrders && c.daysAgo != null && c.daysAgo >= 60)
    .sort((a,b) => (b.daysAgo || 0) - (a.daysAgo || 0));

  const monthOptions = useMemo(() => {
    const ks = [...new Set(wDel.map((o) => monthKey(o.orderDate)).filter(Boolean))].sort();
    return ks;
  }, [wDel]);
  const [agentMonth, setAgentMonth] = useState(curMonthKey);
  useEffect(() => {
    if (!monthOptions.length) return;
    if (agentMonth !== 'all' && !monthOptions.includes(agentMonth)) {
      setAgentMonth(curMonthKey && monthOptions.includes(curMonthKey) ? curMonthKey : monthOptions[monthOptions.length - 1]);
    }
  }, [monthOptions, agentMonth, curMonthKey]);

  const monthly = useMemo(() => {
    const m = {};
    pays.forEach((c) => {
      const k = monthKey(c.sana); if (!k) return;
      if (!m[k]) m[k]={key:k,sum:0,cnt:0};
      m[k].sum+=c.amount; m[k].cnt++;
    });
    return Object.values(m).sort((a,b)=>a.key.localeCompare(b.key)).slice(-9);
  }, [pays]);

  const agents = useMemo(() => {
    const m = {};
    const source = agentMonth === 'all'
      ? wDel
      : wDel.filter((o) => monthKey(o.orderDate) === agentMonth);
    source.forEach((o) => {
      const a=o.agent||'—';
      if (!m[a]) m[a]={name:a,qty:0,sum:0};
      m[a].qty+=Math.abs(o.qty); m[a].sum+=o.sum;
    });
    return Object.values(m).sort((a,b)=>b.sum-a.sum).slice(0,7);
  }, [wDel, agentMonth]);

  return (
    <div style={{display:'flex',flexDirection:'column',gap:14}} className="ani">
      <div className="g5">
        <StatCard l="JAMI MIJOZLAR" v={customers.length} s={customers.filter((c)=>c.hasOrders).length+' aktiv'} c="var(--bl)"/>
        <StatCard l="QARZ UZS"
          v={fmt(debt.uzsSum)+" so'm"}
          s={debt.uzsCount+' ta qarzdor'} c="var(--rd)"/>
        <StatCard l="JAMI IDISH" v={customers.reduce((s,c)=>s+c.tara,0)+' ta'} s="barcha mijozlarda" c="var(--pu)"/>
        <StatCard l="2 OY OLMAGAN" v={stale2m.length+' ta'} s="aktiv mijozlar ichidan" c="var(--yl)"/>
        <StatCard
          l={`SUV (${String(now.getMonth()+1).padStart(2,'0')}.${now.getFullYear()})`}
          v={fmt(wDelThisMonth.reduce((s,o)=>s+Math.abs(o.qty),0))+' ta'}
          s={fmt(wDelThisMonth.reduce((s,o)=>s+o.sum,0))+" so'm"} c="var(--gr)"/>
      </div>
      <div className="g2">
        <div className="card" style={{padding:16}}>
          <div style={{fontWeight:700,fontSize:13,marginBottom:12}}>💰 Oylik to'lovlar (so'm)</div>
          {monthly.length===0
            ? <div style={{color:'var(--t3)',textAlign:'center',padding:24,fontSize:13}}>To'lov ma'lumoti yo'q</div>
            : <ResponsiveContainer width="100%" height={200}>
                <AreaChart data={monthly}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#21262d"/>
                  <XAxis dataKey="key" tick={{fontSize:10,fill:'#656d76'}}/>
                  <YAxis tick={{fontSize:10,fill:'#656d76'}} tickFormatter={fmtM}/>
                  <Tooltip formatter={(v)=>fmt(v)+" so'm"} contentStyle={{background:'#161b22',border:'1px solid #30363d',borderRadius:8}}/>
                  <Area type="monotone" dataKey="sum" fill="var(--gr2)" stroke="var(--gr)" name="Kirim"/>
                </AreaChart>
              </ResponsiveContainer>
          }
        </div>
        <div className="card" style={{padding:16}}>
          <div style={{display:'flex',justifyContent:'space-between',gap:8,alignItems:'center',marginBottom:12}}>
            <div style={{fontWeight:700,fontSize:13}}>🏆 Agentlar reytingi</div>
            <select className="select" value={agentMonth} onChange={(e)=>setAgentMonth(e.target.value)} style={{padding:'4px 8px',fontSize:12}}>
              <option value="all">Hammasi</option>
              {monthOptions.slice().reverse().map((m)=><option key={m} value={m}>{m}</option>)}
            </select>
          </div>
          {agents.map((a,i) => (
            <div key={i} style={{display:'flex',justifyContent:'space-between',padding:'7px 0',borderBottom:'1px solid var(--b2)'}}>
              <div style={{display:'flex',alignItems:'center',gap:8}}>
                <span style={{fontWeight:800,fontSize:11,color:i===0?'var(--yl)':i===1?'var(--t2)':i===2?'var(--or)':'var(--t3)',minWidth:18}}>#{i+1}</span>
                <span style={{fontWeight:600,fontSize:13}}>{a.name}</span>
              </div>
              <div style={{textAlign:'right'}}>
                <div style={{fontWeight:700,fontSize:13,color:'var(--gr)'}}>{fmt(a.sum)} so'm</div>
                <div style={{fontSize:11,color:'var(--t3)'}}>{a.qty} ta</div>
              </div>
            </div>
          ))}
        </div>
      </div>
      <div className="g2">
        <div className="card" style={{padding:16}}>
          <div style={{fontWeight:700,fontSize:13,marginBottom:12}}>⚠️ Top qarzdorlar</div>
          {debtors.length===0
            ? <div style={{color:'var(--t3)',textAlign:'center',padding:20,fontSize:13}}>Qarzdor yo'q ✓</div>
            : debtors.sort((a,b)=>a.balanceUZS-b.balanceUZS).slice(0,8).map((c,i) => (
              <div key={i} style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',gap:10,padding:'6px 0',borderBottom:'1px solid var(--b2)'}}>
                <div style={{minWidth:0,flex:1}}>
                  <div style={{fontWeight:600,fontSize:12.5,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{c.name}</div>
                  <div style={{fontSize:11,color:'var(--t3)'}}>{c.district} · <a href={`tel:${c.phone}`} style={{color:'var(--bl)',textDecoration:'none'}}>{c.phone}</a></div>
                </div>
                <div style={{textAlign:'right',flexShrink:0}}>
                  {c.balanceUZS<0 && <div><span className="tag" style={{background:'var(--rd2)',color:'var(--rd)'}}>-{fmt(Math.abs(c.balanceUZS))} so'm</span></div>}
                  {c.balanceUSD<0 && <div style={{marginTop:2}}><span className="tag" style={{background:'var(--yl2)',color:'var(--yl)'}}>-{fmt(Math.abs(c.balanceUSD))} $</span></div>}
                </div>
              </div>
            ))
          }
        </div>
        <div className="card" style={{padding:16}}>
          <div style={{fontWeight:700,fontSize:13,marginBottom:12}}>🕒 2+ oy suv olmaganlar</div>
          {stale2m.slice(0,8).map((c,i) => (
            <div key={i} style={{display:'flex',justifyContent:'space-between',padding:'6px 0',borderBottom:'1px solid var(--b2)'}}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontWeight:600,fontSize:12.5,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{c.name}</div>
                <div style={{fontSize:11,color:'var(--t3)'}}>{c.district}</div>
              </div>
              <div style={{textAlign:'right',flexShrink:0,marginLeft:8}}>
                <div><span className="tag" style={{background:'var(--yl2)',color:'var(--yl)'}}>{c.daysAgo} kun</span></div>
                <a href={`tel:${c.phone}`} style={{fontSize:11,color:'var(--bl)',textDecoration:'none'}}>{c.phone}</a>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ═══════════════ MIJOZLAR ═══════════════ */
function Customers({ D, currentUser='Admin' }) {
  const { customers } = D;
  const [tab, setTab] = useState(currentUser === 'Admin' ? 'all' : 'own');
  const [search,setS]   = useState('');
  const [sort,setSort]  = useState({ col:'name', dir:'asc' });
  const [det,setDet]    = useState(null);
  const [showAdv, setShowAdv] = useState(false);
  const [adv, setAdv] = useState({
    districts: [], sources: [], agents: [],
    uzsFrom: '', uzsTo: '', usdFrom: '', usdTo: '',
    taraFrom: '', taraTo: '', daysFrom: '', daysTo: '',
    lastFrom: '', lastTo: '',
  });

  const dists = [...new Set(customers.map((c)=>c.district).filter(Boolean))].sort();
  const sources = [...new Set(customers.map((c)=>c.source).filter(Boolean))].sort();
  const agents = [...new Set(customers.map((c)=>c.lastAgent).filter(Boolean))].sort();
  const ownIds = useMemo(
    () => new Set(Object.entries(D.assignmentById || {}).filter(([,op]) => op === currentUser).map(([id]) => id)),
    [D.assignmentById, currentUser]
  );
  useEffect(() => {
    if (currentUser === 'Admin') setTab('all');
  }, [currentUser]);
  const baseCustomers = useMemo(
    () => (tab === 'own' ? customers.filter((c) => ownIds.has(c.id)) : customers),
    [tab, customers, ownIds]
  );

  const toRange = (v, from, to) => {
    if (from !== '' && Number(v) < Number(from)) return false;
    if (to !== '' && Number(v) > Number(to)) return false;
    return true;
  };

  const list = useMemo(() => {
    let r = baseCustomers;
    if (search) {
      const q = search.toLowerCase();
      r = r.filter((c) => c.name.toLowerCase().includes(q)||c.phone.includes(q)||c.id.includes(q)||(c.district||'').toLowerCase().includes(q));
    }

    if (adv.districts.length) r = r.filter((c) => adv.districts.includes(c.district));
    if (adv.sources.length) r = r.filter((c) => adv.sources.includes(c.source));
    if (adv.agents.length) r = r.filter((c) => adv.agents.includes(c.lastAgent));
    r = r.filter((c) => toRange(c.balanceUZS, adv.uzsFrom, adv.uzsTo));
    r = r.filter((c) => toRange(c.balanceUSD, adv.usdFrom, adv.usdTo));
    r = r.filter((c) => toRange(c.tara, adv.taraFrom, adv.taraTo));
    r = r.filter((c) => toRange(c.daysAgo ?? 0, adv.daysFrom, adv.daysTo));
    if (adv.lastFrom) r = r.filter((c) => {
      const d = toDate(c.lastOrderDate); const f = toDate(adv.lastFrom);
      return d && f ? d >= f : false;
    });
    if (adv.lastTo) r = r.filter((c) => {
      const d = toDate(c.lastOrderDate); const t = toDate(adv.lastTo);
      return d && t ? d <= t : false;
    });

    return [...r].sort((a,b) => {
      let av=a[sort.col]??0, bv=b[sort.col]??0;
      if (typeof av==='string') av=av.toLowerCase();
      if (typeof bv==='string') bv=bv.toLowerCase();
      return sort.dir==='asc'?(av>bv?1:-1):av<bv?1:-1;
    });
  }, [baseCustomers,search,sort,adv]);

  const activeFilters = useMemo(() => {
    let count = 0;
    if (adv.districts.length) count++;
    if (adv.sources.length) count++;
    if (adv.agents.length) count++;
    if (adv.uzsFrom !== '' || adv.uzsTo !== '') count++;
    if (adv.usdFrom !== '' || adv.usdTo !== '') count++;
    if (adv.taraFrom !== '' || adv.taraTo !== '') count++;
    if (adv.daysFrom !== '' || adv.daysTo !== '') count++;
    if (adv.lastFrom !== '' || adv.lastTo !== '') count++;
    return count;
  }, [adv]);

  const tog = (col) => setSort((s) => s.col===col?{col,dir:s.dir==='asc'?'desc':'asc'}:{col,dir:'asc'});
  const SI = ({c:col}) => sort.col===col
    ? <span style={{marginLeft:3}}>{sort.dir==='asc'?'↑':'↓'}</span>
    : <span style={{marginLeft:3,opacity:.2}}>↕</span>;
  const balColor = (v) => v<0?'var(--rd)':v>0?'var(--gr)':'var(--t3)';
  const debt = getDebtStats(baseCustomers);

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:12,height:'100%'}}>
      <div className="tabs" style={{display:'inline-flex'}}>
        <button className={`tab${tab==='own'?' on':''}`} onClick={()=>setTab('own')}>👤 O'z mijozlarim</button>
        <button className={`tab${tab==='all'?' on':''}`} onClick={()=>setTab('all')}>👥 Barcha mijozlar</button>
      </div>
      <div className="g4">
        <StatCard l="JAMI MIJOZLAR" v={baseCustomers.length} s={baseCustomers.filter((c)=>c.hasOrders).length+' aktiv'} c="var(--bl)"/>
        <StatCard l="QARZDORLAR" v={debt.uzsCount+' ta'} s={fmt(debt.uzsSum)+" so'm"} c="var(--rd)"/>
        <StatCard l="JAMI IDISH" v={baseCustomers.reduce((s,c)=>s+c.tara,0)+' ta'} s="barcha mijozlarda" c="var(--pu)"/>
        <StatCard l="FILTRLANGAN" v={list.length+' ta'} c="var(--gr)"/>
      </div>
      <div style={{display:'flex',gap:8,flexWrap:'wrap',alignItems:'center',position:'relative'}}>
        <div className="sb" style={{flex:2,minWidth:200}}>
          <span style={{color:'var(--t3)'}}>🔍</span>
          <input placeholder="Ism, telefon, ID bo'yicha..." value={search} onChange={(e)=>setS(e.target.value)}/>
        </div>
        <button className="btn btn-gh btn-sm" onClick={()=>setShowAdv((v)=>!v)}>🧰 Filtr ({activeFilters})</button>
        <button className="btn btn-gr btn-sm" onClick={()=>exportAllReport(baseCustomers)}>⬇ Excel hisobot</button>

        {showAdv && (
          <div className="card" style={{position:'absolute',top:40,right:0,zIndex:30,width:520,padding:14,boxShadow:'0 24px 60px rgba(0,0,0,.55)',backdropFilter:'blur(8px)'}}>
            <div className="g2" style={{marginBottom:8}}>
              <div>
                <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Rayon</div>
                <div style={{maxHeight:90,overflow:'auto',border:'1px solid var(--b1)',borderRadius:8,padding:6}}>
                  {dists.map((d)=><label key={d} style={{display:'flex',gap:6,fontSize:12}}><input type="checkbox" checked={adv.districts.includes(d)} onChange={(e)=>setAdv((p)=>({...p,districts:e.target.checked?[...p.districts,d]:p.districts.filter((x)=>x!==d)}))}/>{d}</label>)}
                </div>
              </div>
              <div>
                <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Manba</div>
                <div style={{maxHeight:90,overflow:'auto',border:'1px solid var(--b1)',borderRadius:8,padding:6}}>
                  {sources.map((d)=><label key={d} style={{display:'flex',gap:6,fontSize:12}}><input type="checkbox" checked={adv.sources.includes(d)} onChange={(e)=>setAdv((p)=>({...p,sources:e.target.checked?[...p.sources,d]:p.sources.filter((x)=>x!==d)}))}/>{d}</label>)}
                </div>
              </div>
              <div>
                <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Agent</div>
                <div style={{maxHeight:90,overflow:'auto',border:'1px solid var(--b1)',borderRadius:8,padding:6}}>
                  {agents.map((d)=><label key={d} style={{display:'flex',gap:6,fontSize:12}}><input type="checkbox" checked={adv.agents.includes(d)} onChange={(e)=>setAdv((p)=>({...p,agents:e.target.checked?[...p.agents,d]:p.agents.filter((x)=>x!==d)}))}/>{d}</label>)}
                </div>
              </div>
              <div style={{display:'grid',gap:6}}>
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
                  <input className="input" placeholder="Balans UZS dan" value={adv.uzsFrom} onChange={(e)=>setAdv((p)=>({...p,uzsFrom:e.target.value}))}/>
                  <input className="input" placeholder="Balans UZS gacha" value={adv.uzsTo} onChange={(e)=>setAdv((p)=>({...p,uzsTo:e.target.value}))}/>
                </div>
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
                  <input className="input" placeholder="Balans USD dan" value={adv.usdFrom} onChange={(e)=>setAdv((p)=>({...p,usdFrom:e.target.value}))}/>
                  <input className="input" placeholder="Balans USD gacha" value={adv.usdTo} onChange={(e)=>setAdv((p)=>({...p,usdTo:e.target.value}))}/>
                </div>
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
                  <input className="input" placeholder="Idish dan" value={adv.taraFrom} onChange={(e)=>setAdv((p)=>({...p,taraFrom:e.target.value}))}/>
                  <input className="input" placeholder="Idish gacha" value={adv.taraTo} onChange={(e)=>setAdv((p)=>({...p,taraTo:e.target.value}))}/>
                </div>
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
                  <input className="input" placeholder="Kun dan" value={adv.daysFrom} onChange={(e)=>setAdv((p)=>({...p,daysFrom:e.target.value}))}/>
                  <input className="input" placeholder="Kun gacha" value={adv.daysTo} onChange={(e)=>setAdv((p)=>({...p,daysTo:e.target.value}))}/>
                </div>
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
                  <input className="input" type="date" value={adv.lastFrom} onChange={(e)=>setAdv((p)=>({...p,lastFrom:e.target.value}))}/>
                  <input className="input" type="date" value={adv.lastTo} onChange={(e)=>setAdv((p)=>({...p,lastTo:e.target.value}))}/>
                </div>
              </div>
            </div>
            <div style={{display:'flex',justifyContent:'space-between'}}>
              <button className="btn btn-gh btn-sm" onClick={()=>setAdv({ districts:[], sources:[], agents:[], uzsFrom:'', uzsTo:'', usdFrom:'', usdTo:'', taraFrom:'', taraTo:'', daysFrom:'', daysTo:'', lastFrom:'', lastTo:'' })}>
                Tozalash
              </button>
              <button className="btn btn-bl btn-sm" onClick={()=>setShowAdv(false)}>Qo'llash</button>
            </div>
          </div>
        )}
      </div>
      <div className="card" style={{overflow:'hidden',flex:1}}>
        <div style={{overflow:'auto',maxHeight:'100%'}}>
          <table className="tbl">
            <thead>
              <tr>
                <th onClick={()=>tog('id')} style={{minWidth:75}}>ID <SI c="id"/></th>
                <th onClick={()=>tog('name')} style={{minWidth:210}}>Kontragent <SI c="name"/></th>
                <th style={{minWidth:115}}>Telefon</th>
                <th onClick={()=>tog('district')} style={{minWidth:95}}>Rayon <SI c="district"/></th>
                <th onClick={()=>tog('balanceUZS')} style={{minWidth:110}}>Balans UZS <SI c="balanceUZS"/></th>
                <th onClick={()=>tog('balanceUSD')} style={{minWidth:95}}>Balans USD <SI c="balanceUSD"/></th>
                <th onClick={()=>tog('tara')} style={{minWidth:65}}>Idish <SI c="tara"/></th>
                <th onClick={()=>tog('kulers')} style={{minWidth:55}}>Kuler <SI c="kulers"/></th>
                <th onClick={()=>tog('lastOrderDate')} style={{minWidth:100}}>Oxirgi zakaz <SI c="lastOrderDate"/></th>
                <th onClick={()=>tog('daysAgo')} style={{minWidth:55}}>Kun <SI c="daysAgo"/></th>
                <th onClick={()=>tog('lastQty')} style={{minWidth:55}}>Dona <SI c="lastQty"/></th>
                <th style={{minWidth:80}}>Agent</th>
              </tr>
            </thead>
            <tbody>
              {list.length===0
                ? <tr><td colSpan={12} style={{textAlign:'center',padding:40,color:'var(--t3)'}}>Topilmadi</td></tr>
                : list.map((c,i) => (
                  <tr key={c.id||i} onClick={()=>setDet(c)}>
                    <td style={{fontFamily:'var(--mono)',fontSize:10.5,color:'var(--t3)'}}>{c.id}</td>
                    <td>
                      <div style={{fontWeight:600,fontSize:12.5,maxWidth:205,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{c.name}</div>
                      {c.address && <div style={{fontSize:10.5,color:'var(--t3)',maxWidth:205,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{c.address}</div>}
                    </td>
                    <td><a href={`tel:${c.phone}`} onClick={(e)=>e.stopPropagation()} style={{color:'var(--bl)',textDecoration:'none',fontFamily:'var(--mono)',fontSize:11.5}}>{c.phone}</a></td>
                    <td style={{fontSize:12}}>{c.district||'—'}</td>
                    <td style={{fontFamily:'var(--mono)',fontSize:11.5,fontWeight:700,color:balColor(c.balanceUZS)}}>{c.balanceUZS<0?'-':c.balanceUZS>0?'+':''}{fmt(Math.abs(c.balanceUZS))}</td>
                    <td style={{fontFamily:'var(--mono)',fontSize:11,fontWeight:700,color:balColor(c.balanceUSD)}}>{c.balanceUSD!==0?<>{c.balanceUSD<0?'-':c.balanceUSD>0?'+':''}{fmt(Math.abs(c.balanceUSD))}$</>:'—'}</td>
                    <td style={{textAlign:'center',fontWeight:700,color:c.tara<0?'var(--rd)':'var(--bl)'}}>{c.tara!==0?(c.tara<0?'-':'')+Math.abs(c.tara):'—'}</td>
                    <td style={{textAlign:'center'}}>{c.kulers>0?<span className="tag" style={{background:'var(--yl2)',color:'var(--yl)'}}>{c.kulers}</span>:'—'}</td>
                    <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(c.lastOrderDate)}</td>
                    <td>{c.daysAgo!=null&&c.daysAgo>=0 ? <span className="tag" style={{background:c.daysAgo>30?'var(--rd2)':c.daysAgo>14?'var(--yl2)':c.daysAgo>7?'var(--or2)':'var(--s3)',color:c.daysAgo>30?'var(--rd)':c.daysAgo>14?'var(--yl)':c.daysAgo>7?'var(--or)':'var(--t3)'}}>{c.daysAgo}k</span> : '—'}</td>
                    <td style={{textAlign:'center'}}>{c.lastQty||'—'}</td>
                    <td style={{fontSize:12}}>{c.lastAgent||'—'}</td>
                  </tr>
                ))
              }
            </tbody>
          </table>
        </div>
      </div>
      {det && <CustomerDetail c={det} D={D} onClose={()=>setDet(null)}/>}
    </div>
  );
}

/* ═══════════════ SO DETAIL MODAL ═══════════════ */
function SoDetailModal({ soGroup, onClose }) {
  return (
    <div className="modal-ov fade" onClick={(e)=>e.target===e.currentTarget&&onClose()}>
      <div className="modal ani" style={{maxWidth:680}}>
        <div className="mhdr">
          <div>
            <div style={{fontWeight:700,fontSize:14,display:'flex',alignItems:'center',gap:10}}>
              <span className="tag" style={{background:'var(--bl3)',color:'var(--bl)'}}>{soGroup.soNum}</span>
              <span style={{fontSize:13,color:'var(--t2)',fontWeight:500}}>{soGroup.contName}</span>
            </div>
            <div style={{fontSize:11,color:'var(--t3)',marginTop:4,fontFamily:'var(--mono)'}}>
              {fmtD(soGroup.orderDate)} · {soGroup.agent||'—'} · {soGroup.items.length} ta mahsulot
              {soGroup.delivPerson && <span> · 🚚 {soGroup.delivPerson}</span>}
            </div>
          </div>
          <button className="btn btn-gh btn-sm" onClick={onClose}>✕</button>
        </div>
        <div className="mbdy">
          <div style={{overflow:'auto',borderRadius:8,border:'1px solid var(--b2)'}}>
            <table className="tbl">
              <thead>
                <tr>
                  <th>Mahsulot</th>
                  <th style={{textAlign:'center'}}>Kol-vo</th>
                  <th style={{textAlign:'right'}}>Narx</th>
                  <th style={{textAlign:'right'}}>Summa</th>
                  <th style={{textAlign:'center'}}>Valyuta</th>
                </tr>
              </thead>
              <tbody>
                {soGroup.items.map((item,i) => (
                  <tr key={i}>
                    <td style={{fontWeight:500,color:'var(--t1)'}}>{item.product}</td>
                    <td style={{textAlign:'center',fontWeight:700,color:item.qty<0?'var(--or)':'var(--t1)'}}>{item.qty}</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{item.price?fmt(Math.round(item.price)):'—'}</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700,color:'var(--gr)'}}>{item.sum?fmt(item.sum):'—'}</td>
                    <td style={{textAlign:'center'}}><span className={item.currency==='USD'?'cur-usd':'cur-uzs'}>{item.currency||'UZS'}</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{marginTop:12,display:'flex',justifyContent:'space-between',alignItems:'center'}}>
            <span className="tag" style={{
              background:isDeliveredStatus(soGroup.status)?'var(--gr2)':soGroup.status==='ОТМЕНЕНО'?'var(--rd2)':'var(--s3)',
              color:isDeliveredStatus(soGroup.status)?'var(--gr)':soGroup.status==='ОТМЕНЕНО'?'var(--t4)':'var(--t3)',
            }}>{soGroup.status||'—'}</span>
            <div style={{display:'flex',gap:12,fontWeight:700,fontSize:13}}>
              {soGroup.totalSumUZS>0 && <span style={{color:'var(--gr)'}}>{fmt(soGroup.totalSumUZS)} so'm</span>}
              {soGroup.totalSumUSD>0 && <span style={{color:'var(--yl)'}}>{fmt(soGroup.totalSumUSD)} $</span>}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════ ZAKAZLAR ═══════════════ */
function Orders({ D }) {
  const { rawOrders=[] } = D;
  const [search,setS]  = useState('');
  const [fType,setT]   = useState('zakaz');
  const [fAgent,setA]  = useState('all');
  const [fStat,setSt]  = useState('all');
  const [fCur,setC]    = useState('all');
  const [selectedSO,setSelectedSO] = useState(null);

  const agents = [...new Set(rawOrders.map((o)=>o.agent).filter(Boolean))].sort();

  const soGroups = useMemo(() => {
    const groups = {};
    rawOrders.forEach((o) => {
      if (!o.soNum) return;
      if (!groups[o.soNum]) {
        groups[o.soNum] = {
          soNum: o.soNum, contName: o.contName, mId: o.mId,
          orderDate: o.orderDate, delivPerson: o.delivPerson,
          agent: o.agent, docType: o.docType, status: o.status,
          currency: o.currency||'UZS', items: [],
          totalQty: 0, totalSumUZS: 0, totalSumUSD: 0,
        };
      }
      const g = groups[o.soNum];
      g.items.push(o);
      g.totalQty += Math.abs(o.qty||0);
      if (o.currency==='USD') g.totalSumUSD+=o.sum||0;
      else g.totalSumUZS+=o.sum||0;
    });
    return Object.values(groups);
  }, [rawOrders]);

  const allZakaz   = soGroups.filter((g)=>isOrderDoc(g.docType));
  const allVozvrat = soGroups.filter((g)=>isReturnDoc(g.docType));
  const base = fType==='zakaz'?allZakaz:fType==='vozvrat'?allVozvrat:soGroups;

  const list = useMemo(() => {
    let r = base;
    const q = search.toLowerCase();
    if (q) r=r.filter((g)=>(g.contName||'').toLowerCase().includes(q)||(g.soNum||'').toLowerCase().includes(q));
    if (fAgent!=='all') r=r.filter((g)=>g.agent===fAgent);
    if (fStat!=='all')  r=r.filter((g)=>g.status===fStat);
    if (fCur!=='all')   r=r.filter((g)=>g.currency===fCur);
    return [...r].sort((a,b) => {
      const da=toDate(a.orderDate), db=toDate(b.orderDate);
      return da&&db?db-da:0;
    });
  }, [base,search,fAgent,fStat,fCur]);

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:12,height:'100%'}}>
      <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
        <div className="sb" style={{flex:1}}>
          <span style={{color:'var(--t3)'}}>🔍</span>
          <input placeholder="Mijoz, zakaz №..." value={search} onChange={(e)=>setS(e.target.value)}/>
        </div>
        <div className="tabs">
          {[['zakaz','📦 Zakaz'],['vozvrat','↩ Vozvrat'],['all','Barchasi']].map(([t,l]) => (
            <button key={t} className={`tab${fType===t?' on':''}`} onClick={()=>setT(t)}>{l}</button>
          ))}
        </div>
        <select className="select" value={fStat} onChange={(e)=>setSt(e.target.value)}>
          <option value="all">Barcha status</option>
          <option value="ДОСТАВЛЕН">ДОСТАВЛЕН</option>
          <option value="ПОЛУЧЕН НА СКЛАД">ПОЛУЧЕН НА СКЛАД</option>
          <option value="НОВЫЙ">НОВЫЙ</option>
          <option value="В ПУТИ">В ПУТИ</option>
          <option value="ОТМЕНЕНО">ОТМЕНЕНО</option>
        </select>
        <select className="select" value={fCur} onChange={(e)=>setC(e.target.value)}>
          <option value="all">Barcha valyuta</option>
          <option value="UZS">UZS</option>
          <option value="USD">USD</option>
        </select>
        <select className="select" value={fAgent} onChange={(e)=>setA(e.target.value)}>
          <option value="all">Barcha agent</option>
          {agents.map((a)=><option key={a}>{a}</option>)}
        </select>
      </div>
      <div className="card" style={{overflow:'hidden',flex:1}}>
        <div style={{overflow:'auto',maxHeight:'100%'}}>
          <table className="tbl">
            <thead>
              <tr>
                <th>Zakaz №</th><th>Kontragent</th><th>Sana</th>
                <th>Dostavchik</th><th style={{textAlign:'center'}}>Dona</th>
                <th style={{textAlign:'right'}}>Summa UZS</th>
                <th style={{textAlign:'right'}}>Summa USD</th>
                <th style={{textAlign:'center'}}>Mahsulotlar</th>
                <th>Tur</th><th>Agent</th><th>Status</th>
              </tr>
            </thead>
            <tbody>
              {list.length===0
                ? <tr><td colSpan={11} style={{textAlign:'center',padding:40,color:'var(--t3)'}}>Topilmadi</td></tr>
                : list.slice(0,600).map((g,i) => (
                  <tr key={i} onClick={()=>setSelectedSO(g)}>
                    <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{g.soNum}</td>
                    <td style={{maxWidth:180,fontWeight:600}}>
                      <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:178,fontSize:12}}>{g.contName}</span>
                    </td>
                    <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(g.orderDate)}</td>
                    <td style={{fontSize:11,color:'var(--t2)'}}>{g.delivPerson||'—'}</td>
                    <td style={{textAlign:'center',fontWeight:700,color:'var(--t1)'}}>{g.totalQty}</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontSize:11.5,fontWeight:700,color:g.totalSumUZS?'var(--gr)':'var(--t4)'}}>
                      {g.totalSumUZS?fmt(g.totalSumUZS):'—'}
                    </td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontSize:11,fontWeight:700,color:g.totalSumUSD?'var(--yl)':'var(--t4)'}}>
                      {g.totalSumUSD?fmt(g.totalSumUSD)+' $':'—'}
                    </td>
                    <td style={{textAlign:'center'}}>
                      <span className="tag" style={{background:'var(--s3)',color:'var(--t3)',cursor:'pointer'}}>{g.items.length} ta</span>
                    </td>
                    <td>
                      <span className="tag" style={{background:isOrderDoc(g.docType)?'var(--bl3)':'var(--or2)',color:isOrderDoc(g.docType)?'var(--bl)':'var(--or)',fontSize:10}}>
                        {g.docType}
                      </span>
                    </td>
                    <td style={{fontSize:12}}>{g.agent||'—'}</td>
                    <td>
                      <span className="tag" style={{
                        background:isDeliveredStatus(g.status)?'var(--gr2)':g.status==='ОТМЕНЕНО'?'var(--rd2)':'var(--s3)',
                        color:isDeliveredStatus(g.status)?'var(--gr)':g.status==='ОТМЕНЕНО'?'var(--t4)':'var(--t3)',
                        fontSize:10,
                      }}>{g.status}</span>
                    </td>
                  </tr>
                ))
              }
            </tbody>
          </table>
        </div>
      </div>
      {selectedSO && <SoDetailModal soGroup={selectedSO} onClose={()=>setSelectedSO(null)}/>}
    </div>
  );
}

/* ═══════════════ KASSA ═══════════════ */
function KassaFilterBtn({ active, label, color, dot, onClick }) {
  return (
    <button onClick={onClick} style={{
      cursor:'pointer',
      border:`1px solid ${active&&color?color:'var(--b1)'}`,
      borderRadius:20,padding:'5px 14px',fontSize:12.5,fontWeight:600,
      background:active&&color?(color==='var(--gr)'?'var(--gr2)':color==='var(--rd)'?'var(--rd2)':'var(--s3)'):'transparent',
      color:active&&color?color:'var(--t2)',
      display:'inline-flex',alignItems:'center',gap:7,transition:'all .13s',
    }}>
      {dot && <span className="dot" style={{background:color,opacity:active?1:0.4}}/>}
      {label}
    </button>
  );
}

function Kassa({ D }) {
  const { cashbox=[] } = D;
  const [fType,setT]  = useState('all');
  const [search,setS] = useState('');
  const now = new Date();
  const curMonthKey = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const pays   = cashbox.filter((c)=>isPaymentFromCounterparty(c.opType));
  const spends = cashbox.filter((c)=>isPaymentToCounterparty(c.opType));
  const paysMonth = pays.filter((c)=>monthKey(c.sana)===curMonthKey);
  const spendsMonth = spends.filter((c)=>monthKey(c.sana)===curMonthKey);
  const cashMonth = cashbox.filter((c)=>monthKey(c.sana)===curMonthKey);
  const totalIn  = paysMonth.reduce((s,c)=>s+c.amount,0);
  const totalOut = spendsMonth.reduce((s,c)=>s+c.amount,0);
  const base = fType==='in'?pays:fType==='out'?spends:cashbox;

  const list = useMemo(() => {
    let r = base;
    if (search) {
      const q = search.toLowerCase();
      r=r.filter((c)=>(c.contName||'').toLowerCase().includes(q)||(c.opNum||'').toLowerCase().includes(q));
    }
    return [...r].sort((a,b) => {
      const da=toDate(a.sana), db=toDate(b.sana);
      return da&&db?db-da:0;
    });
  }, [base,search]);

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:12,height:'100%'}}>
      <div className="g4">
        <StatCard l={`KIRIM (${curMonthKey})`}  v={fmt(totalIn)+" so'm"}  s={paysMonth.length+' ta operatsiya'}   c="var(--gr)"/>
        <StatCard l={`CHIQIM (${curMonthKey})`} v={fmt(totalOut)+" so'm"} s={spendsMonth.length+' ta operatsiya'} c="var(--rd)"/>
        <StatCard l="SALDO (OYLIK)" v={fmt(totalIn-totalOut)+" so'm"} s="joriy oy: kirim − chiqim" c={totalIn-totalOut>=0?'var(--gr)':'var(--rd)'}/>
        <StatCard l="OYLIK OPR."   v={cashMonth.length+' ta'}  s="joriy oy yozuvlari" c="var(--bl)"/>
      </div>
      <div style={{display:'flex',gap:8,alignItems:'center'}}>
        <div className="sb" style={{flex:1}}>
          <span style={{color:'var(--t3)'}}>🔍</span>
          <input placeholder="Mijoz, operatsiya №..." value={search} onChange={(e)=>setS(e.target.value)}/>
        </div>
        <div style={{display:'flex',gap:6}}>
          <KassaFilterBtn active={fType==='all'} label="Barchasi" onClick={()=>setT('all')}/>
          <KassaFilterBtn active={fType==='in'}  label="Kirimlar"  color="var(--gr)" dot onClick={()=>setT('in')}/>
          <KassaFilterBtn active={fType==='out'} label="Chiqimlar" color="var(--rd)" dot onClick={()=>setT('out')}/>
        </div>
      </div>
      <div className="card" style={{overflow:'hidden',flex:1}}>
        <div style={{overflow:'auto',maxHeight:'100%'}}>
          <table className="tbl">
            <thead>
              <tr>
                <th>Sana</th><th>Op №</th><th>Tur</th>
                <th>Kassa</th><th>Kontragent</th><th>Summa</th><th>Operator</th>
              </tr>
            </thead>
            <tbody>
              {list.length===0
                ? <tr><td colSpan={7} style={{textAlign:'center',padding:40,color:'var(--t3)'}}>Topilmadi</td></tr>
                : list.slice(0,600).map((c,i) => {
                  const isIn = isPaymentFromCounterparty(c.opType);
                  return (
                    <tr key={i}>
                      <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(c.sana)}</td>
                      <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{c.opNum}</td>
                      <td style={{fontSize:11.5,maxWidth:140}}>
                        <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:138,
                          color:isIn?'var(--gr)':isPaymentToCounterparty(c.opType)?'var(--rd)':'var(--t3)'}}>
                          {c.opType}
                        </span>
                      </td>
                      <td style={{fontSize:11.5,color:'var(--t3)',maxWidth:110}}>
                        <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:108}}>{c.kassa||'—'}</span>
                      </td>
                      <td style={{maxWidth:220}}>
                        <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:218,fontSize:12}}>{c.contName||'—'}</span>
                      </td>
                      <td style={{fontFamily:'var(--mono)',fontWeight:700,color:isIn?'var(--gr)':isPaymentToCounterparty(c.opType)?'var(--rd)':'var(--t2)'}}>
                        {isIn?'+':isPaymentToCounterparty(c.opType)?'-':''}{fmt(c.amount)}
                      </td>
                      <td style={{fontSize:11.5,color:'var(--t3)'}}>{c.operator||'—'}</td>
                    </tr>
                  );
                })
              }
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════ OBZVON ═══════════════ */
function Obzvon({ D, allRows=[], onAppendAllRow, onReloadAll, webhookUrl='', currentUser='Admin', records=[], setRecords=()=>{} }) {
  const { customers, rawOrders=[] } = D;
  const [tab, setTab] = useState('main');
  const [searchAll, setSearchAll] = useState('');
  const [pickQuery, setPickQuery] = useState('');
  const [opFilter, setOpFilter] = useState('Barchasi');
  const [topicFilter, setTopicFilter] = useState('Barchasi');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  const [allLimit, setAllLimit] = useState(500);
  const [opLimit, setOpLimit] = useState(500);
  const [pickTargetIdx, setPickTargetIdx] = useState(null);
  const [opSearch, setOpSearch] = useState('');
  const [opPickMode, setOpPickMode] = useState(false);
  const [opSelectedIds, setOpSelectedIds] = useState({});
  const deferredSearchAll = useDeferredValue(searchAll);
  const deferredOpSearch = useDeferredValue(opSearch);
  const saveRecords = (next) => {
    setRecords(next);
    S.set('aq-obzvon-records', next);
  };
  const appendRows = (count=1) => {
    const rows = Array.from({ length:count }, () => ({
      id: '',
      customer: '',
      callDate: new Date().toISOString().slice(0,10),
      topic: 'Buyurtma olish',
      note: '',
      nextDate: '',
      orderCount: '',
      orderDate: '',
      operator: currentUser,
    }));
    saveRecords([...(records||[]), ...rows]);
  };
  const addRecord = async (customer, topic='Buyurtma olish') => {
    const row = {
      id: customer.id,
      customer: customer.name,
      callDate: new Date().toISOString().slice(0,10),
      topic,
      note: '',
      nextDate: '',
      orderCount: '',
      orderDate: '',
      operator: currentUser,
    };
    const next = [row, ...records];
    saveRecords(next);
    if (onAppendAllRow) onAppendAllRow(row);
    if (webhookUrl) {
      try { await fetch(webhookUrl, { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(row) }); } catch {}
    }
  };

  const mergedAllRows = useMemo(() => {
    const localRows = records
      .filter((r) => r.id || r.customer)
      .map((r, i) => ({
        no: String(i + 1),
        customer: r.customer,
        callDate: r.callDate,
        topic: r.topic,
        note: r.note,
        nextDate: r.nextDate,
        orderCount: r.orderCount,
        operator: r.operator,
        customerId: r.id,
        orderDate: r.orderDate,
      }));
    let base = [...localRows, ...allRows];
    const seen = new Set();
    base = base.filter((r) => {
      const k = `${r.customerId || ''}|${r.customer || ''}|${r.callDate || ''}|${r.topic || ''}|${r.note || ''}|${r.operator || ''}`;
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });
    base.sort((a, b) => {
      const da = toDate(a.callDate);
      const db = toDate(b.callDate);
      return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
    });
    return base;
  }, [allRows, records]);

  const allList = useMemo(() => {
    const q = deferredSearchAll.toLowerCase();
    let base = mergedAllRows;
    if (opFilter !== 'Barchasi') base = base.filter((r) => (r.operator || '') === opFilter);
    if (topicFilter !== 'Barchasi') base = base.filter((r) => (r.topic || '') === topicFilter);
    if (dateFrom) {
      const f = toDate(dateFrom);
      if (f) base = base.filter((r) => {
        const d = toDate(r.callDate);
        return d ? d >= f : false;
      });
    }
    if (dateTo) {
      const t = toDate(dateTo);
      if (t) base = base.filter((r) => {
        const d = toDate(r.callDate);
        return d ? d <= t : false;
      });
    }
    if (!q) return base;
    return base.filter((r) =>
      (r.customer || '').toLowerCase().includes(q) ||
      (r.customerId || '').toLowerCase().includes(q) ||
      (r.operator || '').toLowerCase().includes(q) ||
      (r.note || '').toLowerCase().includes(q)
    );
  }, [mergedAllRows, deferredSearchAll, opFilter, topicFilter, dateFrom, dateTo]);

  useEffect(() => {
    setAllLimit(500);
  }, [searchAll, opFilter, topicFilter, dateFrom, dateTo]);

  const visibleAllRows = useMemo(
    () => allList.slice(0, allLimit),
    [allList, allLimit]
  );
  const operators = useMemo(() => {
    const raw = mergedAllRows.map((r) => ({ operator: r.operator }));
    return ['Barchasi', ...new Set(raw.map((r) => r.operator).filter(Boolean))];
  }, [mergedAllRows]);
  const topics = useMemo(() => {
    const raw = mergedAllRows.map((r) => ({ topic: r.topic }));
    return ['Barchasi', ...new Set(raw.map((r) => r.topic).filter(Boolean))];
  }, [mergedAllRows]);

  const suggestions = useMemo(() => {
    const q = pickQuery.toLowerCase().trim();
    if (!q) return [];
    return customers.filter((c) =>
      (c.name || '').toLowerCase().includes(q) ||
      String(c.id || '').includes(q) ||
      String(c.phone || '').includes(q) ||
      String(c.district || '').toLowerCase().includes(q)
    ).slice(0,40);
  }, [pickQuery, customers]);

  const dueCandidates = useMemo(() => {
    const grouped = {};
    rawOrders.forEach((o) => {
      if (!isOrderDoc(o.docType)) return;
      if (!isDeliveredStatus(o.status)) return;
      if (!isWaterProduct(o.product)) return;
      if ((o.delivPerson || '').toLowerCase().includes('vertual')) return;
      if (!o.mId) return;
      if (!grouped[o.mId]) grouped[o.mId] = [];
      grouped[o.mId].push(o);
    });
    const out = [];
    Object.entries(grouped).forEach(([mid, ords]) => {
      const sorted = ords
        .filter((o) => toDate(o.orderDate))
        .sort((a,b)=>(toDate(b.orderDate)-toDate(a.orderDate)));
      if (sorted.length < 3) return;
      const last3 = sorted.slice(0,3);
      const newest = toDate(last3[0].orderDate);
      const oldest = toDate(last3[2].orderDate);
      let spanDays = Math.round((newest - oldest) / 864e5);
      if (spanDays < 2) spanDays = 7;
      const totalQty = last3.reduce((s,o)=>s+Math.abs(o.qty||0),0);
      const daily = totalQty / spanDays;
      if (daily <= 0) return;
      const lastQty = Math.abs(last3[0].qty || 0);
      const shouldIn = Math.max(1, Math.floor(lastQty / daily) - 1);
      const c = customers.find((x)=>x.id===mid);
      if (!c) return;
      const nm = (c.name || '').trim().toUpperCase();
      if (nm.startsWith('Я TUGATILDI') || nm.startsWith('Я ESKI')) return;
      const passed = c.daysAgo ?? 0;
      if (passed >= shouldIn) {
        out.push({
          ...c,
          shouldIn,
          passed,
          last3Info: `${Math.abs(last3[0].qty)} / ${Math.abs(last3[1].qty)} / ${Math.abs(last3[2].qty)} ta`,
        });
      }
    });
    return out.sort((a,b)=>(b.passed-a.passed));
  }, [rawOrders, customers]);
  const latestCallByCustomer = useMemo(() => {
    const m = {};
    for (const r of mergedAllRows) {
      const cid = String(r.customerId || r.id || '').trim();
      if (!cid) continue;
      const d = toDate(r.callDate);
      const t = d ? d.getTime() : 0;
      const prev = m[cid];
      const prevT = prev ? ((toDate(prev.callDate)?.getTime()) || 0) : -1;
      if (!prev || t >= prevT) m[cid] = r;
    }
    return m;
  }, [mergedAllRows]);

  const latestOrdersByCustomer = useMemo(() => {
    const out = {};
    const src = D.ordersByMId || {};
    Object.keys(src).forEach((cid) => {
      const ords = (src[cid] || [])
        .filter((o) => isOrderDoc(o.docType))
        .sort((a,b)=>(toDate(b.orderDate)-toDate(a.orderDate)));
      out[cid] = {
        ord1: ords[0]?.orderDate || '',
        ord2: ords[1]?.orderDate || '',
        lastQty: ords[0] ? Math.abs(ords[0].qty || 0) : 0,
      };
    });
    return out;
  }, [D.ordersByMId]);

  const operatorTableRows = useMemo(() => {
    let rows = (D.customers || []).map((c) => {
      const lastCall = latestCallByCustomer[c.id];
      const ord = latestOrdersByCustomer[c.id] || { ord1:'', ord2:'', lastQty:0 };
      return {
        id: c.id,
        name: c.name,
        tara: c.tara,
        balance: c.balanceUZS,
        ord1: ord.ord1,
        ord2: ord.ord2,
        lastQty: ord.lastQty,
        lastCallDate: lastCall?.callDate || '',
        lastNote: lastCall?.note || '',
        nextDate: lastCall?.nextDate || '',
        operator: D.assignmentById?.[c.id] || '—',
      };
    });
    if (currentUser !== 'Admin') rows = rows.filter((r) => r.operator === currentUser);
    if (opFilter !== 'Barchasi') rows = rows.filter((r) => r.operator === opFilter);
    if (deferredOpSearch) {
      const q = deferredOpSearch.toLowerCase();
      rows = rows.filter((r) =>
        String(r.id || '').includes(q) ||
        String(r.name || '').toLowerCase().includes(q) ||
        String(r.operator || '').toLowerCase().includes(q) ||
        String(r.lastNote || '').toLowerCase().includes(q)
      );
    }
    return rows;
  }, [D, latestCallByCustomer, latestOrdersByCustomer, currentUser, opFilter, deferredOpSearch]);

  useEffect(() => {
    setOpLimit(500);
  }, [opSearch, opFilter]);

  const visibleOperatorRows = useMemo(
    () => operatorTableRows.slice(0, opLimit),
    [operatorTableRows, opLimit]
  );

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:12}}>
      <div className="tabs" style={{display:'inline-flex'}}>
        <button className={`tab${tab==='main'?' on':''}`} onClick={()=>setTab('main')}>Obzvon</button>
        <button className={`tab${tab==='all'?' on':''}`} onClick={()=>setTab('all')}>Obzvon ВСЕ</button>
        <button className={`tab${tab==='due'?' on':''}`} onClick={()=>setTab('due')}>Vaqti Kelgan</button>
        <button className={`tab${tab==='op'?' on':''}`} onClick={()=>setTab('op')}>Operator jadvali</button>
      </div>

      {tab==='main' && (
        <>
          <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
            <div className="sb" style={{minWidth:340,flex:1}}>
              <span style={{color:'var(--t3)'}}>🔍</span>
              <input placeholder={pickTargetIdx!=null ? "Qator uchun mijoz qidiring..." : "Yangi mijoz qo'shish uchun qidiring..."} value={pickQuery} onChange={(e)=>setPickQuery(e.target.value)} />
            </div>
            <span className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>{records.length} ta yozuv</span>
            <button className="btn btn-gh btn-sm" onClick={()=>appendRows(1)}>+ Qator</button>
          </div>
          {suggestions.length>0 && (
            <div className="card" style={{padding:10,maxHeight:160,overflow:'auto'}}>
              {suggestions.map((s) => (
                <div key={s.id} className="nav-i" style={{padding:'6px 8px'}} onClick={()=>{
                  if (pickTargetIdx != null) {
                    saveRecords(records.map((x,j)=>j===pickTargetIdx?{...x,id:s.id,customer:s.name}:x));
                    setPickTargetIdx(null);
                  } else {
                    addRecord(s);
                  }
                  setPickQuery('');
                }}>
                  <span style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{s.id}</span>
                  <span style={{overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{s.name}</span>
                </div>
              ))}
            </div>
          )}
          <div className="card" style={{overflow:'hidden'}}>
            <div style={{overflow:'auto',maxHeight:'58vh'}}>
              <table className="tbl">
                <thead><tr><th>ID</th><th>Mijoz</th><th>Sana</th><th>Maqsad</th><th>Izoh</th><th>Keyingi sana</th><th>Zakaz soni</th><th>Zakaz sana</th><th>Operator</th><th></th></tr></thead>
                <tbody>
                  {records.length===0 ? <tr><td colSpan={10} style={{textAlign:'center',padding:28,color:'var(--t3)'}}>Obzvon yozuvlari yo'q</td></tr> :
                    records.map((r,i)=>(
                      <tr key={i}>
                        <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{r.id || '—'}</td>
                        <td style={{minWidth:250}}>
                          <div style={{display:'flex',gap:6,alignItems:'center'}}>
                            <input
                              className="input"
                              value={r.customer || ''}
                              placeholder="Mijoz"
                              onChange={(e)=>saveRecords(records.map((x,j)=>j===i?{...x,customer:e.target.value,id:''}:x))}
                            />
                            <button className="btn btn-gh btn-sm" onClick={()=>{ setPickTargetIdx(i); setPickQuery(r.customer || ''); }}>Tanlash</button>
                          </div>
                        </td>
                        <td><input className="input" type="date" value={String(r.callDate||'').slice(0,10)} onChange={(e)=>saveRecords(records.map((x,j)=>j===i?{...x,callDate:e.target.value}:x))} /></td>
                        <td>
                          <select className="select" value={r.topic||'Buyurtma olish'} onChange={(e)=>saveRecords(records.map((x,j)=>j===i?{...x,topic:e.target.value}:x))}>
                            <option>Buyurtma olish</option>
                            <option>Qarzdorlik</option>
                            <option>Tara togrlash</option>
                          </select>
                        </td>
                        <td><input className="input" value={r.note||''} onChange={(e)=>saveRecords(records.map((x,j)=>j===i?{...x,note:e.target.value}:x))} /></td>
                        <td><input className="input" type="date" value={String(r.nextDate||'').slice(0,10)} onChange={(e)=>saveRecords(records.map((x,j)=>j===i?{...x,nextDate:e.target.value}:x))} /></td>
                        <td><input className="input" value={r.orderCount||''} onChange={(e)=>saveRecords(records.map((x,j)=>j===i?{...x,orderCount:e.target.value}:x))} style={{maxWidth:90}} /></td>
                        <td><input className="input" type="date" value={String(r.orderDate||'').slice(0,10)} onChange={(e)=>saveRecords(records.map((x,j)=>j===i?{...x,orderDate:e.target.value}:x))} /></td>
                        <td>{r.operator || currentUser}</td>
                        <td><button className="btn btn-gh btn-sm" onClick={()=>saveRecords(records.filter((_,j)=>j!==i))}>✕</button></td>
                      </tr>
                    ))
                  }
                </tbody>
              </table>
            </div>
          </div>
          <div style={{display:'flex',justifyContent:'flex-end'}}>
            <button className="btn btn-bl btn-sm" onClick={()=>appendRows(1000)}>+ 1000 ta qator qo'shish</button>
          </div>
        </>
      )}

      {tab==='all' && (
        <>
          <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
            <div className="sb" style={{maxWidth:420,flex:1}}>
              <span style={{color:'var(--t3)'}}>🔍</span>
              <input placeholder="Mijoz / ID / operator..." value={searchAll} onChange={(e)=>setSearchAll(e.target.value)} />
            </div>
            <select className="select" value={opFilter} onChange={(e)=>setOpFilter(e.target.value)}>
              {operators.map((o)=><option key={o}>{o}</option>)}
            </select>
            <select className="select" value={topicFilter} onChange={(e)=>setTopicFilter(e.target.value)}>
              {topics.map((t)=><option key={t}>{t}</option>)}
            </select>
            <input className="input" type="date" style={{maxWidth:140}} value={dateFrom} onChange={(e)=>setDateFrom(e.target.value)} />
            <input className="input" type="date" style={{maxWidth:140}} value={dateTo} onChange={(e)=>setDateTo(e.target.value)} />
            <button className="btn btn-gh btn-sm" onClick={() => { setOpFilter('Barchasi'); setTopicFilter('Barchasi'); setDateFrom(''); setDateTo(''); setSearchAll(''); }}>
              Tozalash
            </button>
            <button className="btn btn-bl btn-sm" onClick={() => onReloadAll && onReloadAll()}>
              🔄 Yangilash
            </button>
            <span className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>{visibleAllRows.length} / {allList.length} ta yozuv</span>
          </div>
          <div className="card" style={{overflow:'hidden',flex:1}}>
            <div style={{overflow:'auto',maxHeight:'calc(100vh - 230px)'}}>
              <table className="tbl">
                <thead><tr><th>№</th><th>ID</th><th>Kontragent</th><th>Sana</th><th>Mavzu</th><th>Izoh</th><th>Keyingi sana</th><th>Zakaz</th><th>Zakaz sanasi</th><th>Operator</th></tr></thead>
                <tbody>
                  {allList.length===0 ? <tr><td colSpan={10} style={{textAlign:'center',padding:32,color:'var(--t3)'}}>Ma'lumot topilmadi</td></tr> :
                    visibleAllRows.map((r,i)=>(
                      <tr key={i}>
                        <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{r.no || i+1}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{r.customerId || '—'}</td>
                        <td style={{maxWidth:360}}><span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.customer || `ID: ${r.customerId}`}</span></td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.callDate)}</td>
                        <td>{r.topic || '—'}</td>
                        <td style={{maxWidth:300}}><span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.note || '—'}</span></td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.nextDate)}</td>
                        <td>{r.orderCount || '—'}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.orderDate)}</td>
                        <td>{r.operator || '—'}</td>
                      </tr>
                    ))
                  }
                </tbody>
              </table>
            </div>
          </div>
          {visibleAllRows.length < allList.length && (
            <div style={{display:'flex',justifyContent:'center'}}>
              <button className="btn btn-gh btn-sm" onClick={() => setAllLimit((p) => p + 500)}>
                Yana 500 ta yuklash
              </button>
            </div>
          )}
        </>
      )}

      {tab==='due' && (
        <div className="card" style={{overflow:'hidden'}}>
          <div style={{overflow:'auto',maxHeight:'calc(100vh - 220px)'}}>
            <table className="tbl">
              <thead><tr><th>Mijoz</th><th>ID</th><th>Oxirgi 3 zakaz</th><th>O'tgan kun</th><th>Qo'ng'iroq me'yori</th><th>Amal</th></tr></thead>
              <tbody>
                {dueCandidates.length===0 ? <tr><td colSpan={6} style={{textAlign:'center',padding:28,color:'var(--t3)'}}>Vaqti kelgan mijoz yo'q</td></tr> :
                  dueCandidates.map((c,i)=>(
                    <tr key={i}>
                      <td style={{maxWidth:340}}><span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{c.name}</span></td>
                      <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{c.id}</td>
                      <td>{c.last3Info}</td>
                      <td>{c.passed} kun</td>
                      <td>{c.shouldIn} kunda</td>
                      <td><button className="btn btn-bl btn-sm" onClick={()=>{ addRecord(c,'Buyurtma olish'); setTab('main'); }}>Obzvonga qo'shish</button></td>
                    </tr>
                  ))
                }
              </tbody>
            </table>
          </div>
        </div>
      )}

      {tab==='op' && (
        <>
          <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
            <div className="sb" style={{maxWidth:420,flex:1}}>
              <span style={{color:'var(--t3)'}}>🔍</span>
              <input placeholder="ID, mijoz, izoh..." value={opSearch} onChange={(e)=>setOpSearch(e.target.value)} />
            </div>
            <button className={`btn ${opPickMode?'btn-gr':'btn-gh'} btn-sm`} onClick={()=>setOpPickMode((v)=>!v)}>
              + Obzvonga qo'shish
            </button>
          </div>
          <div className="card" style={{overflow:'hidden'}}>
            <div style={{overflow:'auto',maxHeight:'calc(100vh - 260px)'}}>
              <table className="tbl">
                <thead><tr><th>ID</th><th>Mijoz</th><th style={{textAlign:'right'}}>Balans</th><th>Oxirgi zakaz</th><th>Oxirgi qo'ng'iroq</th><th>Keyingi sana</th><th>Operator</th></tr></thead>
                <tbody>
                  {operatorTableRows.length===0 ? <tr><td colSpan={7} style={{textAlign:'center',padding:26,color:'var(--t3)'}}>Ma'lumot topilmadi</td></tr> :
                    visibleOperatorRows.map((r, i) => (
                      <tr
                        key={i}
                        onClick={()=>{
                          if (!opPickMode) return;
                          setOpSelectedIds((p)=>({ ...p, [r.id]: !p[r.id] }));
                        }}
                        style={opPickMode && opSelectedIds[r.id] ? { background:'rgba(88,166,255,.11)' } : undefined}
                      >
                        <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{r.id}</td>
                        <td style={{maxWidth:360}}><span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.name}</span></td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)',color:r.balance<0?'var(--rd)':r.balance>0?'var(--gr)':'var(--t3)'}}>{fmt(r.balance)}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.ord1)} {r.lastQty?`· ${r.lastQty} ta`:''}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.lastCallDate)}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.nextDate)}</td>
                        <td>{r.operator}</td>
                      </tr>
                    ))
                  }
                </tbody>
              </table>
            </div>
          </div>
          {visibleOperatorRows.length < operatorTableRows.length && (
            <div style={{display:'flex',justifyContent:'center'}}>
              <button className="btn btn-gh btn-sm" onClick={() => setOpLimit((p) => p + 500)}>
                Operatorlar: yana 500 ta
              </button>
            </div>
          )}
          {opPickMode && (
            <div style={{display:'flex',justifyContent:'flex-end'}}>
              <button
                className="btn btn-bl btn-sm"
                disabled={Object.values(opSelectedIds).filter(Boolean).length===0}
                onClick={() => {
                  const selected = operatorTableRows.filter((r) => opSelectedIds[r.id]);
                  if (!selected.length) return;
                  const rows = selected.map((r) => ({
                    id: r.id,
                    customer: r.name,
                    callDate: new Date().toISOString().slice(0,10),
                    topic: 'Buyurtma olish',
                    note: '',
                    nextDate: '',
                    orderCount: '',
                    orderDate: '',
                    operator: currentUser,
                  }));
                  saveRecords([...rows, ...records]);
                  rows.forEach((x)=>onAppendAllRow?.(x));
                  setOpSelectedIds({});
                  setOpPickMode(false);
                  setTab('main');
                }}
              >
                Tanlanganlarni qo'shish ({Object.values(opSelectedIds).filter(Boolean).length} ta)
              </button>
            </div>
          )}
        </>
      )}
    </div>
  );
}

/* ═══════════════ DOLJNIKI ═══════════════ */
function DoljnikModal({ row, D, onClose }) {
  const customer = (D.customers || []).find((c) => c.id === row.id);
  const rows = customer ? buildSverka(customer, D.ordersByMId || {}, D.cashByMId || {}) : [];
  return (
    <div className="modal-ov fade" onClick={(e)=>e.target===e.currentTarget&&onClose()}>
      <div className="modal ani" style={{maxWidth:980}}>
        <div className="mhdr">
          <div>
            <div style={{fontWeight:800,fontSize:14}}>{row.name}</div>
            <div style={{fontSize:11,color:'var(--t3)'}}>ID: {row.id} · Qarz zakaz: {row.orderNo || '—'}</div>
          </div>
          <button className="btn btn-gh btn-sm" onClick={onClose}>✕</button>
        </div>
        <div className="mbdy">
          <div style={{overflow:'auto',maxHeight:'62vh',border:'1px solid var(--b2)',borderRadius:8}}>
            <table className="tbl">
              <thead><tr><th>Kod</th><th>Sana</th><th>Hujjat</th><th>Mahsulot</th><th>Summa</th><th>To'lov</th><th>Balans UZS</th></tr></thead>
              <tbody>
                {rows.length===0 ? <tr><td colSpan={7} style={{textAlign:'center',padding:28,color:'var(--t3)'}}>Ma'lumot yo'q</td></tr> :
                  rows.map((r,i)=>(
                    <tr key={i} style={{background:r.kod===row.orderNo?'rgba(88,166,255,.08)':'transparent'}}>
                      <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{r.kod}</td>
                      <td style={{fontFamily:'var(--mono)',fontSize:11}}>{r.sana}</td>
                      <td>{r.dokument}</td>
                      <td style={{maxWidth:280,overflow:'hidden',textOverflow:'ellipsis'}}>{r.produkt || '—'}</td>
                      <td style={{fontFamily:'var(--mono)',textAlign:'right',color:r.summa?'var(--rd)':'var(--t4)'}}>{r.summa?'-'+fmt(r.summa):'—'}</td>
                      <td style={{fontFamily:'var(--mono)',textAlign:'right',color:r.tolov?'var(--gr)':'var(--t4)'}}>{r.tolov?'+'+fmt(r.tolov):'—'}</td>
                      <td style={{fontFamily:'var(--mono)',textAlign:'right',fontWeight:700,color:r.balansUZS<0?'var(--rd)':r.balansUZS>0?'var(--gr)':'var(--t3)'}}>{r.balansUZS<0?'-':r.balansUZS>0?'+':''}{fmt(Math.abs(r.balansUZS))}</td>
                    </tr>
                  ))
                }
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

function KulerModal({ row, D, onClose, onSaveMonths }) {
  const [tab, setTab] = useState('payments');
  const [months, setMonths] = useState(row.months || 6);
  useEffect(() => { setMonths(row.months || 6); }, [row.months]);

  const customer = useMemo(
    () => (D.customers || []).find((c) => c.id === row.customerId) || null,
    [D.customers, row.customerId]
  );
  const sverkaRows = useMemo(
    () => (customer ? buildSverka(customer, D.ordersByMId || {}, D.cashByMId || {}) : []),
    [customer, D.ordersByMId, D.cashByMId]
  );
  const kulerPayments = row.payments || [];

  return (
    <div className="modal-ov fade" onClick={(e)=>e.target===e.currentTarget&&onClose()}>
      <div className="modal ani" style={{maxWidth:980}}>
        <div className="mhdr">
          <div>
            <div style={{fontWeight:800,fontSize:14}}>{row.customerName}</div>
            <div style={{fontSize:11,color:'var(--t3)'}}>
              ID: {row.customerId} · Zakaz: {row.orderNo} · Olingan: {fmtD(row.purchaseDate)}
            </div>
          </div>
          <div style={{display:'flex',gap:8}}>
            <button className="btn btn-gr btn-sm" onClick={() => customer && exportSverkaExcel(customer, sverkaRows)}>⬇ Excel</button>
            <button className="btn btn-gh btn-sm" onClick={onClose}>✕</button>
          </div>
        </div>
        <div className="mbdy">
          <div className="g3" style={{marginBottom:12}}>
            <StatCard l="ASOSIY SUMMA" v={fmt(row.principal)+" so'm"} c="var(--bl)"/>
            <StatCard l="OYLIK TO'LOV" v={fmt(row.monthly)+" so'm"} c="var(--yl)"/>
            <StatCard l="TO'LANGAN" v={fmt(row.paid)+" so'm"} c="var(--gr)"/>
            <StatCard l="QOLDIQ" v={fmt(row.remaining)+" so'm"} c="var(--rd)"/>
            <StatCard l="MUDDATI KELGAN" v={row.overdueAmount>0?fmt(row.overdueAmount)+" so'm":'0 so\'m'} c={row.overdueAmount>0?'var(--rd)':'var(--gr)'}/>
            <StatCard l="QOLGAN OY" v={row.monthsLeft+' oy'} c="var(--or)"/>
          </div>

          <div className="tabs" style={{marginBottom:12,display:'inline-flex'}}>
            <button className={`tab${tab==='payments'?' on':''}`} onClick={()=>setTab('payments')}>💰 To'lovlar</button>
            <button className={`tab${tab==='settings'?' on':''}`} onClick={()=>setTab('settings')}>⚙️ Nastroyka</button>
          </div>

          {tab==='payments' && (
            <div className="card" style={{overflow:'auto',maxHeight:'52vh'}}>
              <table className="tbl">
                <thead><tr><th>Sana</th><th>Op №</th><th>Operator</th><th>Kassa</th><th>To'lov</th></tr></thead>
                <tbody>
                  {kulerPayments.length===0 ? (
                    <tr><td colSpan={5} style={{textAlign:'center',padding:24,color:'var(--t3)'}}>To'lovlar topilmadi</td></tr>
                  ) : kulerPayments.map((p, i) => (
                    <tr key={i}>
                      <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(p.sana)}</td>
                      <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{p.opNum||'—'}</td>
                      <td>{p.operator||'—'}</td>
                      <td>{p.kassa||'—'}</td>
                      <td style={{fontFamily:'var(--mono)',color:'var(--gr)',fontWeight:700}}>+{fmt(p.amount)} so'm</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {tab==='settings' && (
            <div className="card" style={{padding:14}}>
              <div style={{fontSize:12,color:'var(--t3)',marginBottom:8}}>Rassrochka muddatini o'zgartiring (oy):</div>
              <div style={{display:'flex',gap:8,alignItems:'center'}}>
                <input className="input" type="number" min={1} max={60} value={months} onChange={(e)=>setMonths(e.target.value)} style={{maxWidth:120}} />
                <button className="btn btn-bl btn-sm" onClick={()=>onSaveMonths(months)}>Saqlash</button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function Doljniki({ rows, D, kulerRows, onAddToObzvon, currentUser }) {
  const [tab, setTab] = useState('qarz');
  const [search, setSearch] = useState('');
  const [fCats, setCats] = useState([]);
  const [showCat, setShowCat] = useState(false);
  const [catQuery, setCatQuery] = useState('');
  const [fDayFrom, setDayFrom] = useState('');
  const [fDayTo, setDayTo] = useState('');
  const [fNote, setNote] = useState('');
  const [selected, setSelected] = useState(null);
  const [selectedKuler, setSelectedKuler] = useState(null);
  const [pickMode, setPickMode] = useState(false);
  const [selectedIds, setSelectedIds] = useState({});
  const [kulerMonthsCfg, setKulerMonthsCfg] = useState(() => S.get('aq-kuler-months', {}));
  const categories = [...new Set(rows.map((r) => r.category).filter(Boolean))].sort();
  const shownCats = categories.filter((c) => c.toLowerCase().includes(catQuery.toLowerCase()));

  const list = useMemo(() => {
    let r = rows;
    const q = search.toLowerCase();
    if (q) r = r.filter((x)=> (x.name||'').toLowerCase().includes(q) || String(x.id||'').toLowerCase().includes(q) || (x.orderNo||'').toLowerCase().includes(q));
    if (fCats.length) r = r.filter((x) => fCats.includes(x.category));
    if (fDayFrom !== '') r = r.filter((x) => (x.days ?? 0) >= Number(fDayFrom));
    if (fDayTo !== '') r = r.filter((x) => (x.days ?? 0) <= Number(fDayTo));
    if (fNote) r = r.filter((x) => (x.note || '').toLowerCase().includes(fNote.toLowerCase()));
    return r;
  }, [rows, search, fCats, fDayFrom, fDayTo, fNote]);

  const debtSum = list.reduce((s, r) => s + Math.abs(r.debtUZS), 0);
  const d15 = list.filter((r) => (r.days ?? 0) > 15);
  const d15Sum = d15.reduce((s, r) => s + Math.abs(r.debtUZS), 0);
  const kulerComputed = useMemo(() => {
    return (kulerRows || []).map((k) => {
      const key = `${k.customerId}__${k.orderNo}`;
      const months = kulerMonthsCfg[key] ?? k.months ?? 6;
      return recalcInstallment(k, months);
    });
  }, [kulerRows, kulerMonthsCfg]);
  const kulerActive = kulerComputed.filter((k) => k.remaining > 0);
  const kulerRemain = kulerActive.reduce((s, k) => s + k.remaining, 0);
  const kulerDue = kulerActive.filter((k) => k.overdueAmount > 0).length;
  const saveKulerMonths = (row, monthsValue) => {
    const m = Math.max(1, Math.min(60, Number(monthsValue) || 6));
    const key = `${row.customerId}__${row.orderNo}`;
    setKulerMonthsCfg((prev) => {
      const next = { ...prev, [key]: m };
      S.set('aq-kuler-months', next);
      return next;
    });
    setSelectedKuler((prev) => (prev ? recalcInstallment(prev, m) : prev));
  };
  const selectedCount = Object.values(selectedIds).filter(Boolean).length;
  const addSelectedToObzvon = () => {
    const rowsToAdd = list.filter((r) => selectedIds[r.id]);
    if (!rowsToAdd.length) return;
    onAddToObzvon?.(rowsToAdd.map((r) => ({
      id: r.id,
      customer: r.name,
      callDate: new Date().toISOString().slice(0,10),
      topic: 'Qarzdorlik',
      note: '',
      nextDate: '',
      orderCount: '',
      orderDate: '',
      operator: currentUser || 'Admin',
    })));
    setSelectedIds({});
    setPickMode(false);
  };

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:12,height:'100%'}}>
      <div className="tabs" style={{display:'inline-flex'}}>
        <button className={`tab${tab==='qarz'?' on':''}`} onClick={()=>setTab('qarz')}>Doljniki</button>
        <button className={`tab${tab==='kuler'?' on':''}`} onClick={()=>setTab('kuler')}>Kuler Nasiya</button>
      </div>
      {tab==='qarz' ? (
        <>
          <div className="g4">
            <StatCard l="QARZDORLAR" v={list.length+' ta'} s="UZS bo'yicha" c="var(--rd)"/>
            <StatCard l="JAMI QARZ" v={fmt(debtSum)+" so'm"} s="faqat UZS" c="var(--or)"/>
            <StatCard l="15+ KUN QARZDOR" v={d15.length+' ta'} s={fmt(d15Sum)+" so'm"} c="var(--yl)"/>
            <StatCard l="KULER TO'LOVCHILAR" v={kulerActive.length+' ta'} s={fmt(kulerRemain)+" so'm"} c="var(--bl)"/>
          </div>

          <div style={{display:'flex',gap:8,flexWrap:'wrap',alignItems:'center'}}>
            <div className="sb" style={{flex:2,minWidth:220}}>
              <span style={{color:'var(--t3)'}}>🔍</span>
              <input placeholder="Ism, ID, zakaz bo'yicha..." value={search} onChange={(e)=>setSearch(e.target.value)} />
            </div>
            <input className="input" style={{maxWidth:120}} placeholder="ot: kun" value={fDayFrom} onChange={(e)=>setDayFrom(e.target.value)} />
            <input className="input" style={{maxWidth:120}} placeholder="do: kun" value={fDayTo} onChange={(e)=>setDayTo(e.target.value)} />
            <div style={{position:'relative'}}>
              <button className="btn btn-gh btn-sm" onClick={()=>setShowCat((v)=>!v)}>
                Kategoriya ({fCats.length || 'all'})
              </button>
              {showCat && (
                <div className="card" style={{position:'absolute',top:34,right:0,zIndex:20,width:280,padding:10}}>
                  <input className="input" placeholder="Kategoriya qidirish..." value={catQuery} onChange={(e)=>setCatQuery(e.target.value)} />
                  <div style={{maxHeight:180,overflow:'auto',marginTop:8,paddingRight:4}}>
                    {shownCats.map((c) => (
                      <label key={c} style={{display:'flex',alignItems:'center',gap:8,padding:'4px 0',fontSize:12}}>
                        <input
                          type="checkbox"
                          checked={fCats.includes(c)}
                          onChange={(e)=>setCats((prev)=>e.target.checked?[...prev,c]:prev.filter((x)=>x!==c))}
                        />
                        <span>{c}</span>
                      </label>
                    ))}
                  </div>
                  <div style={{display:'flex',justifyContent:'space-between',marginTop:8}}>
                    <button className="btn btn-gh btn-sm" onClick={()=>setCats([])}>Tozalash</button>
                    <button className="btn btn-bl btn-sm" onClick={()=>setShowCat(false)}>Yopish</button>
                  </div>
                </div>
              )}
            </div>
            <input className="input" style={{maxWidth:220}} placeholder="Izoh filter (T ustun)" value={fNote} onChange={(e)=>setNote(e.target.value)} />
            <button className={`btn ${pickMode ? 'btn-gr' : 'btn-gh'} btn-sm`} onClick={()=>setPickMode((v)=>!v)}>
              + Obzvonga qo'shish
            </button>
          </div>

          <div className="card" style={{overflow:'hidden',flex:1}}>
            <div style={{overflow:'auto',maxHeight:'100%'}}>
              <table className="tbl">
                <thead>
                  <tr>
                    <th>ID</th>
                    <th>Kontragent</th>
                    <th>Kategoriya</th>
                    <th style={{textAlign:'right'}}>Qarz (UZS)</th>
                    <th>Sana</th>
                    <th style={{textAlign:'center'}}>Kun</th>
                    <th>Izoh (T)</th>
                  </tr>
                </thead>
                <tbody>
                  {list.length===0 ? (
                    <tr><td colSpan={7} style={{textAlign:'center',padding:36,color:'var(--t3)'}}>Ma'lumot topilmadi</td></tr>
                  ) : list.map((r, i) => (
                    <tr
                      key={i}
                      onClick={()=>{
                        if (pickMode) {
                          setSelectedIds((p)=>({ ...p, [r.id]: !p[r.id] }));
                          return;
                        }
                        setSelected(r);
                      }}
                      style={pickMode && selectedIds[r.id] ? { background:'rgba(88,166,255,.11)' } : undefined}
                    >
                      <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{r.id}</td>
                      <td style={{maxWidth:280}}>
                        <div style={{fontWeight:700,fontSize:12.5,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.name}</div>
                        <div style={{fontSize:10.5,color:'var(--t3)'}}>Qarz zakaz: {r.orderNo || '—'} {r.lastOrderProduct ? `· ${r.lastOrderProduct}` : ''}</div>
                      </td>
                      <td style={{fontSize:11.5,color:'var(--t2)'}}>{r.category}</td>
                      <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:800,color:'var(--rd)'}}>-{fmt(Math.abs(r.debtUZS))}</td>
                      <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.lastOrderDate)}</td>
                      <td style={{textAlign:'center'}}>
                        {r.days != null ? <span className="tag" style={{background:r.days>30?'var(--rd2)':r.days>15?'var(--yl2)':'var(--s3)',color:r.days>30?'var(--rd)':r.days>15?'var(--yl)':'var(--t3)'}}>{r.days}k</span> : '—'}
                      </td>
                      <td style={{maxWidth:260}}>
                        <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',fontSize:11,color:'var(--t3)'}}>{r.note || '—'}</span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
          {pickMode && (
            <div style={{display:'flex',justifyContent:'flex-end'}}>
              <button className="btn btn-bl" disabled={selectedCount===0} onClick={addSelectedToObzvon}>
                Tanlanganlarni qo'shish ({selectedCount} ta)
              </button>
            </div>
          )}
          {selected && <DoljnikModal row={selected} D={D} onClose={()=>setSelected(null)} />}
        </>
      ) : (
        <>
          <div className="g4">
            <StatCard l="KULER TO'LOVCHILAR" v={kulerActive.length+' ta'} s="nasiya mijozlar" c="var(--bl)"/>
            <StatCard l="KULER QOLDIQ" v={fmt(kulerRemain)+" so'm"} s="umumiy qoldiq" c="var(--or)"/>
            <StatCard l="MUDDATI KELGAN" v={kulerDue+' ta'} s="to'lov vaqti kelgan" c="var(--rd)"/>
            <StatCard l="KULER QARZDORLAR" v={kulerActive.filter((k)=>k.remaining>0).length+' ta'} s="qoldiq bor mijoz" c="var(--yl)"/>
          </div>
          <div className="card" style={{overflow:'hidden',flex:1}}>
            <div style={{overflow:'auto',maxHeight:'100%'}}>
              <table className="tbl">
                <thead><tr><th>Mijoz</th><th>Zakaz</th><th>Olingan sana</th><th>Oylar</th><th>To'langan</th><th>Qoldiq</th><th>Muddati kelgan</th></tr></thead>
                <tbody>
                  {kulerActive.length===0 ? <tr><td colSpan={7} style={{textAlign:'center',padding:30,color:'var(--t3)'}}>Kuler nasiya topilmadi</td></tr> :
                    kulerActive.map((k,i)=>(
                      <tr key={i} onClick={()=>setSelectedKuler(k)}>
                        <td style={{width:'66%',textAlign:'left'}}>
                          <div style={{fontWeight:700,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{k.customerName}</div>
                          <div style={{fontSize:11,color:'var(--t3)'}}>{k.product || 'Kuler'}</div>
                        </td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{k.orderNo}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(k.purchaseDate)}</td>
                        <td>{k.months} oy</td>
                        <td style={{fontFamily:'var(--mono)',color:'var(--gr)'}}>{fmt(k.paid)} so'm</td>
                        <td style={{fontFamily:'var(--mono)',color:'var(--rd)'}}>{fmt(k.remaining)} so'm</td>
                        <td>{k.overdueAmount>0 ? <span className="tag" style={{background:'var(--rd2)',color:'var(--rd)'}}>{fmt(k.overdueAmount)} so'm</span> : <span className="tag" style={{background:'var(--gr2)',color:'var(--gr)'}}>OK</span>}</td>
                      </tr>
                    ))
                  }
                </tbody>
              </table>
            </div>
          </div>
          {selectedKuler && <KulerModal row={selectedKuler} D={D} onClose={()=>setSelectedKuler(null)} onSaveMonths={(m)=>saveKulerMonths(selectedKuler, m)} />}
        </>
      )}
    </div>
  );
}

/* ═══════════════ HISOBOTLAR ═══════════════ */
function Reports({ D }) {
  const { customers, rawOrders=[] } = D;
  const now = new Date();
  const curMonthKey = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const wDel = rawOrders.filter((o) => isWaterProduct(o.product) && isOrderDoc(o.docType) && isDeliveredStatus(o.status));
  const monthOptions = useMemo(
    () => [...new Set(wDel.map((o) => monthKey(o.orderDate)).filter(Boolean))].sort(),
    [wDel]
  );
  const [repMonth, setRepMonth] = useState('all');
  useEffect(() => {
    if (repMonth === 'all') return;
    if (monthOptions.length && !monthOptions.includes(repMonth)) {
      setRepMonth(monthOptions.includes(curMonthKey) ? curMonthKey : 'all');
    }
  }, [monthOptions, repMonth, curMonthKey]);
  const reportOrders = useMemo(
    () => repMonth === 'all' ? wDel : wDel.filter((o) => monthKey(o.orderDate) === repMonth),
    [wDel, repMonth]
  );

  const byDist = useMemo(() => {
    const m = {};
    reportOrders.forEach((o) => {
      const cust = customers.find((c)=>c.id===o.mId);
      const d = cust?.district||"Noma'lum";
      if (!m[d]) m[d]={name:d,qty:0,sum:0,custs:new Set()};
      m[d].qty+=Math.abs(o.qty); m[d].sum+=o.sum; m[d].custs.add(o.mId);
    });
    return Object.values(m).map((d)=>({...d,custs:d.custs.size})).sort((a,b)=>b.sum-a.sum);
  }, [reportOrders,customers]);

  const byMonth = useMemo(() => {
    const m = {};
    wDel.forEach((o) => {
      const k = monthKey(o.orderDate); if (!k) return;
      if (!m[k]) m[k]={key:k,qty:0,sumUZS:0,sumUSD:0};
      m[k].qty+=Math.abs(o.qty);
      if (o.currency==='USD') m[k].sumUSD+=o.sum;
      else m[k].sumUZS+=o.sum;
    });
    return Object.values(m).sort((a,b)=>a.key.localeCompare(b.key)).slice(-12);
  }, [wDel]);

  const byAgent = useMemo(() => {
    const m = {};
    reportOrders.forEach((o) => {
      const a=o.agent||'—';
      if (!m[a]) m[a]={name:a,qty:0,sum:0,custs:new Set()};
      m[a].qty+=Math.abs(o.qty); m[a].sum+=o.sum; m[a].custs.add(o.mId);
    });
    return Object.values(m).map((d)=>({...d,custs:d.custs.size})).sort((a,b)=>b.sum-a.sum);
  }, [reportOrders]);

  const COLORS = ['#58a6ff','#3fb950','#f85149','#d29922','#bc8cff','#f0883e','#79c0ff','#56d364','#ffa657'];

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:14}}>
      <div style={{display:'flex',justifyContent:'space-between',gap:8,alignItems:'center',flexWrap:'wrap'}}>
        <div className="tabs" style={{display:'inline-flex'}}>
          <button className={`tab${repMonth==='all'?' on':''}`} onClick={()=>setRepMonth('all')}>Hammasi</button>
          {monthOptions.slice().reverse().slice(0,8).map((m)=>(
            <button key={m} className={`tab${repMonth===m?' on':''}`} onClick={()=>setRepMonth(m)}>{m}</button>
          ))}
        </div>
        <select className="select" value={repMonth} onChange={(e)=>setRepMonth(e.target.value)}>
          <option value="all">Hammasi</option>
          {monthOptions.slice().reverse().map((m)=><option key={m} value={m}>{m}</option>)}
        </select>
      </div>
      <div className="g4">
        <StatCard l="YETKAZILGAN JAMI" v={fmt(reportOrders.reduce((s,o)=>s+Math.abs(o.qty),0))+' ta'} c="var(--bl)"/>
        <StatCard l="JAMI SUMMA (UZS)" v={fmt(reportOrders.filter((o)=>o.currency!=='USD').reduce((s,o)=>s+o.sum,0))+" so'm"} c="var(--gr)"/>
        <StatCard l="AKTIV RAYONLAR"   v={byDist.length+' ta'} c="var(--yl)"/>
        <StatCard l="AKTIV AGENTLAR"   v={byAgent.length+' ta'} c="var(--pu)"/>
      </div>
      {byMonth.length>0 && (
        <div className="card" style={{padding:16}}>
          <div style={{fontWeight:700,fontSize:13,marginBottom:12}}>📦 Oylik yetkazish (UZS)</div>
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={byMonth}>
              <CartesianGrid strokeDasharray="3 3" stroke="#21262d"/>
              <XAxis dataKey="key" tick={{fontSize:10,fill:'#656d76'}}/>
              <YAxis yAxisId="q" tick={{fontSize:10,fill:'#656d76'}} orientation="left"/>
              <YAxis yAxisId="s" tick={{fontSize:10,fill:'#656d76'}} orientation="right" tickFormatter={fmtM}/>
              <Tooltip formatter={(v,n)=>n==='Summa UZS'?fmt(v)+" so'm":v+' ta'} contentStyle={{background:'#161b22',border:'1px solid #30363d',borderRadius:8}}/>
              <Bar yAxisId="q" dataKey="qty"    fill="var(--bl)" radius={[4,4,0,0]} name="Dona"/>
              <Bar yAxisId="s" dataKey="sumUZS" fill="var(--gr)" radius={[4,4,0,0]} name="Summa UZS"/>
            </BarChart>
          </ResponsiveContainer>
        </div>
      )}
      <div className="g2">
        <div className="card" style={{padding:16}}>
          <div style={{fontWeight:700,fontSize:13,marginBottom:14}}>🗺️ Rayonlar bo'yicha</div>
          {byDist.slice(0,12).map((d,i) => (
            <div key={i} style={{marginBottom:10}}>
              <div style={{display:'flex',justifyContent:'space-between',fontSize:12.5,marginBottom:3}}>
                <span style={{fontWeight:600}}>{d.name}</span>
                <div style={{display:'flex',gap:12,color:'var(--t3)'}}>
                  <span>{d.custs} mijoz</span>
                  <span style={{color:'var(--gr)',fontWeight:700}}>{fmtM(d.sum)}</span>
                </div>
              </div>
              <div style={{height:4,background:'var(--s3)',borderRadius:2}}>
                <div style={{height:'100%',width:`${(d.sum/(byDist[0]?.sum||1))*100}%`,background:COLORS[i%COLORS.length],borderRadius:2,transition:'width .5s'}}/>
              </div>
            </div>
          ))}
        </div>
        <div className="card" style={{padding:16}}>
          <div style={{fontWeight:700,fontSize:13,marginBottom:12}}>🏆 Agentlar natijalari</div>
          <table className="tbl">
            <thead><tr><th>#</th><th>Agent</th><th>Mijoz</th><th>Dona</th><th>Summa</th></tr></thead>
            <tbody>
              {byAgent.map((a,i) => (
                <tr key={i}>
                  <td style={{fontWeight:800,color:i===0?'var(--yl)':i===1?'var(--t2)':i===2?'var(--or)':'var(--t3)'}}>{i+1}</td>
                  <td style={{fontWeight:600}}>{a.name}</td>
                  <td style={{color:'var(--t3)'}}>{a.custs}</td>
                  <td style={{fontFamily:'var(--mono)'}}>{a.qty}</td>
                  <td style={{fontFamily:'var(--mono)',color:'var(--gr)',fontWeight:700}}>{fmtM(a.sum)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function SettingsPanel({
  users, setUsers, access, setAccess, currentUser, setCurrentUser,
  webhookUrl, setWebhookUrl, userCreds, setUserCreds, onSwitchUser, isAdminSession=false,
}) {
  const [tab, setTab] = useState('staff');
  const [sel, setSel] = useState(currentUser || users[0] || 'Admin');
  const [editUser, setEditUser] = useState('');
  const [editLogin, setEditLogin] = useState('');
  const [editPassword, setEditPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const conf = access[sel] || { scope: 'all', visible: {} };
  const pages = ['dash','cust','orders','kassa','obzvon','doljniki','reports','settings'];
  const defaultVisible = { dash:true, cust:true, orders:true, kassa:true, obzvon:true, doljniki:true, reports:true, settings:false };

  const addUser = () => {
    const name = prompt("Yangi login nomi:");
    if (!name || !name.trim()) return;
    const u = name.trim();
    if (users.includes(u)) return;
    const nextUsers = [...users, u];
    setUsers(nextUsers);
    S.set('aq-users', nextUsers);
    setAccess((prev) => {
      const next = { ...prev, [u]: { scope: 'own', visible: defaultVisible } };
      S.set('aq-access', next);
      return next;
    });
    setUserCreds((prev) => {
      const next = { ...(prev || {}), [u]: u };
      S.set('aq-user-creds', next);
      return next;
    });
    setSel(u);
  };

  const openCredEditor = (u) => {
    setSel(u);
    setEditUser(u);
    setEditLogin(u);
    setEditPassword(String(userCreds?.[u] ?? u));
    setShowPassword(false);
  };

  const closeCredEditor = () => {
    setEditUser('');
    setEditLogin('');
    setEditPassword('');
    setShowPassword(false);
  };

  const saveCredEditor = () => {
    if (!isAdminSession || !editUser) return;
    const oldLogin = editUser;
    const newLogin = String(editLogin || '').trim();
    const newPassword = String(editPassword || '').trim();
    if (!newLogin) { alert("Login bo'sh bo'lishi mumkin emas"); return; }
    if (newLogin !== oldLogin && users.includes(newLogin)) {
      alert("Bunday login allaqachon mavjud");
      return;
    }

    const nextUsers = users.map((u) => (u === oldLogin ? newLogin : u));
    setUsers(nextUsers);
    S.set('aq-users', nextUsers);

    setAccess((prev) => {
      const prevConf = prev[oldLogin] || { scope:'own', visible: defaultVisible };
      const next = { ...prev };
      if (newLogin !== oldLogin) delete next[oldLogin];
      next[newLogin] = prevConf;
      S.set('aq-access', next);
      return next;
    });

    setUserCreds((prev) => {
      const source = prev || {};
      const next = { ...source };
      const oldPass = String(source[oldLogin] ?? oldLogin).trim();
      if (newLogin !== oldLogin) delete next[oldLogin];
      next[newLogin] = newPassword || oldPass || newLogin;
      S.set('aq-user-creds', next);
      return next;
    });

    if (currentUser === oldLogin) {
      if (onSwitchUser) onSwitchUser(newLogin);
      else setCurrentUser(newLogin);
    }
    setSel(newLogin);
    closeCredEditor();
  };

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:12}}>
      <div className="tabs" style={{display:'inline-flex'}}>
        <button className={`tab${tab==='staff'?' on':''}`} onClick={()=>setTab('staff')}>?? Hodimlar ruhsatlari</button>
        <button className={`tab${tab==='app'?' on':''}`} onClick={()=>setTab('app')}>?? Ilova sozlamalari</button>
      </div>

      {tab==='staff' && (
        <>
          <div className="g4">
            <StatCard l="LOGINS" v={users.length+' ta'} c="var(--bl)"/>
            <StatCard l="AKTIV LOGIN" v={currentUser} c="var(--gr)"/>
            <StatCard l="TANLANGAN" v={sel} c="var(--yl)"/>
            <StatCard l="SCOPE" v={conf.scope==='all'?'Barcha':'O`ziga'} c="var(--pu)"/>
          </div>
          <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
            {isAdminSession ? (
              <select className="select" value={currentUser} onChange={(e)=>onSwitchUser ? onSwitchUser(e.target.value) : setCurrentUser(e.target.value)}>
                {users.map((u)=><option key={u}>{u}</option>)}
              </select>
            ) : (
              <span className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>{currentUser}</span>
            )}
            <button className="btn btn-gh btn-sm" onClick={addUser}>+ Login qo'shish</button>
          </div>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(280px,1fr))',gap:10}}>
            {users.map((u) => {
              const uc = access[u] || { scope: 'own', visible: {} };
              return (
                <div key={u} className="card" style={{padding:12,border:u===sel?'1px solid var(--bl)':'1px solid var(--b2)'}}>
                  <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:10}}>
                    <button
                      type="button"
                      onClick={() => isAdminSession && openCredEditor(u)}
                      style={{
                        display:'flex',alignItems:'center',gap:8,border:'none',background:'transparent',
                        color:'inherit',cursor:isAdminSession?'pointer':'default',padding:0
                      }}
                    >
                      <div style={{width:28,height:28,borderRadius:14,background:'var(--s3)',display:'flex',alignItems:'center',justifyContent:'center',fontWeight:800}}>
                        {u.slice(0,1).toUpperCase()}
                      </div>
                      <div style={{textAlign:'left'}}>
                        <div style={{fontWeight:700}}>{u}</div>
                        <div style={{fontSize:10.5,color:'var(--t3)'}}>Ismga bosing: login/parol</div>
                      </div>
                    </button>
                    <button className="btn btn-gh btn-sm" onClick={()=>setSel(u)}>Sozlash</button>
                  </div>
                  {editUser === u && (
                    <div style={{background:'var(--s3)',border:'1px solid var(--b1)',borderRadius:8,padding:8,marginBottom:8}}>
                      <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Login</div>
                      <input
                        className="input"
                        value={editLogin}
                        onChange={(e)=>setEditLogin(e.target.value)}
                        disabled={!isAdminSession}
                        style={{padding:'6px 8px',fontSize:12,marginBottom:8}}
                      />
                      <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Parol</div>
                      <div style={{display:'flex',gap:6}}>
                        <input
                          className="input"
                          type={showPassword ? 'text' : 'password'}
                          value={editPassword}
                          onChange={(e)=>setEditPassword(e.target.value)}
                          disabled={!isAdminSession}
                          style={{padding:'6px 8px',fontSize:12,flex:1}}
                        />
                        <button className="btn btn-gh btn-sm" onClick={()=>setShowPassword((v)=>!v)} type="button">
                          {showPassword ? "🙈 Yop" : "👁 Ko'r"}
                        </button>
                        <button className="btn btn-gh btn-sm" onClick={()=>setEditPassword('')} type="button" disabled={!isAdminSession}>
                          Tozalash
                        </button>
                      </div>
                      <div style={{display:'flex',justifyContent:'flex-end',gap:6,marginTop:8}}>
                        <button className="btn btn-gh btn-sm" onClick={closeCredEditor} type="button">Bekor</button>
                        <button className="btn btn-bl btn-sm" onClick={saveCredEditor} type="button" disabled={!isAdminSession}>Saqlash</button>
                      </div>
                    </div>
                  )}
                  <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:8}}>
                    <span style={{fontSize:12,color:'var(--t3)'}}>Faqat o'z mijozlari</span>
                    <button className={`toggle${uc.scope==='own'?' on':''}`} onClick={() => {
                      setAccess((prev) => {
                        const next = { ...prev, [u]: { ...(prev[u] || uc), scope: (prev[u]?.scope || uc.scope) === 'own' ? 'all' : 'own' } };
                        S.set('aq-access', next);
                        return next;
                      });
                    }}><span className="knob" /></button>
                  </div>
                  <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
                    {pages.map((p) => (
                      <div key={p} style={{display:'flex',alignItems:'center',justifyContent:'space-between',background:'var(--s3)',borderRadius:8,padding:'6px 8px'}}>
                        <span style={{fontSize:11,color:'var(--t2)'}}>{p}</span>
                        <button className={`toggle${(uc.visible||{})[p]?' on':''}`} onClick={() => {
                          setAccess((prev) => {
                            const confU = prev[u] || uc;
                            const next = { ...prev, [u]: { ...confU, visible: { ...(confU.visible || {}), [p]: !(confU.visible || {})[p] } } };
                            S.set('aq-access', next);
                            return next;
                          });
                        }}><span className="knob" /></button>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        </>
      )}

      {tab==='app' && (
        <div className="card" style={{padding:14}}>
          <div style={{fontWeight:700,marginBottom:8}}>Obzvon webhook URL</div>
          <input className="input" placeholder="https://script.google.com/macros/s/.../exec" value={webhookUrl} onChange={(e)=>setWebhookUrl(e.target.value)} />
          <div style={{fontSize:11,color:'var(--t3)',marginTop:8}}>Bu URL ga Tab 1 obzvon yozuvlari yuboriladi (real-time va 07:00 sinxron).</div>
          <div style={{marginTop:14,padding:10,border:'1px dashed var(--b1)',borderRadius:8,color:'var(--t3)',fontSize:12}}>Kelajak sozlamalari uchun joy.</div>
        </div>
      )}
    </div>
  );
}

function LoginScreen({ users=[], onLogin, onResetCreds }) {
  const [user, setUser] = useState(users[0] || 'Admin');
  const [pass, setPass] = useState('');
  const [err, setErr] = useState('');

  useEffect(() => {
    if (!users.includes(user)) setUser(users[0] || 'Admin');
  }, [users, user]);

  const submit = () => {
    const ok = onLogin?.(user, pass);
    if (!ok) {
      setErr('Login yoki parol xato');
      return;
    }
    setErr('');
    setPass('');
  };

  return (
    <div className="modal-ov" style={{background:'var(--bg)'}}>
      <div className="modal ani" style={{maxWidth:420}}>
        <div className="mhdr">
          <div>
            <div style={{fontWeight:800,fontSize:16}}>AquaBiz Pro</div>
            <div style={{fontSize:12,color:'var(--t3)',marginTop:3}}>Login orqali kirish</div>
          </div>
        </div>
        <div className="mbdy" style={{display:'flex',flexDirection:'column',gap:10}}>
          <div>
            <div style={{fontSize:11,color:'var(--t3)',marginBottom:5}}>Foydalanuvchi</div>
            <select className="select" style={{width:'100%'}} value={user} onChange={(e)=>setUser(e.target.value)}>
              {users.map((u)=><option key={u} value={u}>{u}</option>)}
            </select>
          </div>
          <div>
            <div style={{fontSize:11,color:'var(--t3)',marginBottom:5}}>Parol</div>
            <input
              className="input"
              type="password"
              value={pass}
              onChange={(e)=>setPass(e.target.value)}
              onKeyDown={(e)=>e.key==='Enter'&&submit()}
              placeholder="Parolni kiriting"
            />
            <div style={{fontSize:11,color:'var(--t3)',marginTop:5}}>
              Standart: Dildora/Dilfuza = login nomi, Admin = 12345
            </div>
          </div>
          {err && <div style={{fontSize:12,color:'var(--rd)'}}>{err}</div>}
          <button className="btn btn-gh" onClick={onResetCreds} style={{justifyContent:'center'}}>Parolni tiklash</button>
          <button className="btn btn-bl" onClick={submit} style={{justifyContent:'center'}}>Kirish</button>
        </div>
      </div>
    </div>
  );
}
/* ═══════════════ ROOT APP ═══════════════ */
const NAV = [
  { id:'dash',    label:'Dashboard',  icon:'🏠' },
  { id:'cust',    label:'Mijozlar',   icon:'👥', badge:'d' },
  { id:'orders',  label:'Zakazlar',   icon:'📦' },
  { id:'kassa',   label:'Kassa',      icon:'💰' },
  { id:'obzvon',  label:'Obzvon',     icon:'📞', badge:'o' },
  { id:'doljniki',label:'Doljniki',   icon:'🧾', badge:'dz' },
  { id:'reports', label:'Hisobotlar', icon:'📊' },
];

export default function App() {
  const [page,setPage]     = useState('dash');
  const [data,setData]     = useState(null);
  const [obzvonAllRows,setObzvonAllRows] = useState(() => S.get('aq-obzvon-all-cache', []) || []);
  const [obzvonAllInstalled,setObzvonAllInstalled] = useState(() => !!S.get(OBZVON_ALL_INSTALLED_KEY, false));
  const [obzvonRecords,setObzvonRecords] = useState(() => S.get('aq-obzvon-records', []));
  const [users,setUsers] = useState(() => S.get('aq-users', DEFAULT_USERS));
  const [access,setAccess] = useState(() => S.get('aq-access', DEFAULT_ACCESS));
  const [userCreds,setUserCreds] = useState(() => {
    const saved = S.get('aq-user-creds', null);
    if (saved && typeof saved === 'object' && Object.keys(saved).length > 0) return saved;
    return DEFAULT_USER_CREDS;
  });
  const [currentUser,setCurrentUser] = useState(() => S.get('aq-current-user', 'Admin'));
  const [sessionUser,setSessionUser] = useState('');
  const [isLoggedIn,setIsLoggedIn] = useState(false);
  const [mainSheetUrl, setMainSheetUrl] = useState(() => SHEET_CONFIG.url || '');
  const [obzvonSheetUrl, setObzvonSheetUrl] = useState(() => OBZVON_ALL_SHEET_URL || '');
  const [obzvonWebhook,setObzvonWebhook] = useState(() => S.get('aq-obzvon-webhook', OBZVON_WEBHOOK_DEFAULT));
  const [showUp,setUp]     = useState(false);
  const [notif,setNotif]   = useState(null);
  const [side,setSide]     = useState(true);
  const [autoLoad,setAutoLoad] = useState({ loading:false, progress:'', error:'' });
  const buildDefaultCreds = useCallback((baseUsers) => {
    const m = {};
    (baseUsers || []).forEach((u) => { m[u] = u; });
    return m;
  }, []);
  const getEffectiveCreds = useCallback(() => {
    const stored = S.get('aq-user-creds', {});
    const merged = {
      ...buildDefaultCreds(users),
      ...(stored && typeof stored === 'object' ? stored : {}),
      ...(userCreds || {}),
    };
    users.forEach((u) => {
      const v = String(merged[u] ?? '').trim();
      if (!v) merged[u] = (DEFAULT_USER_CREDS[u] || u);
    });
    return merged;
  }, [users, userCreds, buildDefaultCreds]);

  const notify = (msg, type='ok') => {
    setNotif({ msg, type });
    setTimeout(() => setNotif(null), 5000);
  };
  useEffect(() => { S.set('aq-current-user', currentUser); }, [currentUser]);
  useEffect(() => { S.set('aq-obzvon-webhook', obzvonWebhook); }, [obzvonWebhook]);
  useEffect(() => { S.set('aq-main-url', mainSheetUrl || ''); }, [mainSheetUrl]);
  useEffect(() => { S.set('aq-obzvon-url', obzvonSheetUrl || ''); }, [obzvonSheetUrl]);
  useEffect(() => {
    const fixed = SHEET_CONFIG.url || '';
    if (mainSheetUrl !== fixed) setMainSheetUrl(fixed);
  }, [mainSheetUrl]);
  useEffect(() => {
    const fixed = OBZVON_ALL_SHEET_URL || '';
    if (obzvonSheetUrl !== fixed) setObzvonSheetUrl(fixed);
  }, [obzvonSheetUrl]);
  useEffect(() => { S.set('aq-obzvon-records', obzvonRecords || []); }, [obzvonRecords]);
  useEffect(() => { S.set('aq-user-creds', userCreds || {}); }, [userCreds]);

  useEffect(() => {
    if (users.length === 0) return;
    if (!users.includes(currentUser)) setCurrentUser(users[0]);
    setUserCreds((prev) => {
      const next = { ...(prev || {}) };
      let changed = false;
      users.forEach((u) => {
        const v = String(next[u] ?? '').trim();
        if (!v) { next[u] = u; changed = true; }
      });
      return changed ? next : prev;
    });
    if (sessionUser && !users.includes(sessionUser)) {
      setSessionUser('');
      setIsLoggedIn(false);
    }
  }, [users, currentUser, sessionUser]);

  const authenticate = useCallback((loginName, password) => {
    const user = String(loginName || '').trim();
    if (!user || !users.includes(user)) return false;
    const typed = String(password || '').trim();
    if (!typed) return false;
    const t = typed.toLowerCase();
    const effective = getEffectiveCreds();
    const stored = S.get('aq-user-creds', {});
    const passCandidates = [
      String((effective || {})[user] ?? '').trim(),
      String((stored || {})[user] ?? '').trim(),
      String(DEFAULT_USER_CREDS[user] ?? '').trim(),
      user,
    ]
      .map((x) => x.toLowerCase())
      .filter(Boolean);
    if (!passCandidates.includes(t)) return false;
    setSessionUser(user);
    setCurrentUser(user);
    setIsLoggedIn(true);
    return true;
  }, [users, getEffectiveCreds]);

  const resetLoginCreds = useCallback(() => {
    const next = buildDefaultCreds(users);
    users.forEach((u) => {
      if (DEFAULT_USER_CREDS[u]) next[u] = DEFAULT_USER_CREDS[u];
    });
    setUserCreds(next);
    S.set('aq-user-creds', next);
    notify("Parollar tiklandi (Admin=12345, qolgani=login)", 'ok');
  }, [users, buildDefaultCreds]);

  const logout = useCallback(() => {
    setIsLoggedIn(false);
    setSessionUser('');
  }, []);

  const switchUser = (nextUser) => {
    if (nextUser === currentUser) return;
    if (sessionUser !== 'Admin') return;
    setCurrentUser(nextUser);
  };

  const handleLoad = (result) => {
    setData(result);
    setUp(false);
    notify(`✅ ${result.customers.length} mijoz · ${result.rawOrders.length} zakaz · ${result.cashbox.length} kassa`);
  };

  const loadFromConfig = useCallback(async () => {
    const configuredMainUrl = mainSheetUrl || SHEET_CONFIG.url || '';
    const sheetId = extractSheetId(configuredMainUrl);
    if (!sheetId) {
      setAutoLoad({ loading:false, progress:'', error:"Google Sheets URL topilmadi" });
      return;
    }
    setAutoLoad({ loading:true, progress:'Ulanilmoqda...', error:'' });
    try {
      const raw = await loadFromGoogleSheets(sheetId, SHEET_CONFIG.gids, (msg)=>setAutoLoad((p)=>({...p,progress:msg})), 'named');
      const d = processAll(raw);
      setData(d);
      setAutoLoad({ loading:false, progress:'', error:'' });
      notify(`✅ ${d.customers.length} mijoz · ${d.rawOrders.length} zakaz`);
    } catch (e) {
      setAutoLoad({ loading:false, progress:'', error:e.message });
      notify('❌ '+e.message, 'err');
    }
  }, [mainSheetUrl]);

  const loadObzvonAll = useCallback(async (sheetUrl, opts = {}) => {
    const full = !!opts.full;
    const force = !!opts.force;
    const cached = S.get('aq-obzvon-all-cache', []) || [];
    if (full && !force && cached.length > 0) {
      setObzvonAllRows(cached);
      setObzvonAllInstalled(true);
      return;
    }
    try {
      const configuredObzvonUrl = sheetUrl || obzvonSheetUrl || OBZVON_ALL_SHEET_URL || '';
      const sid = extractSheetId(configuredObzvonUrl);
      const gid = extractGid(configuredObzvonUrl);
      if (!sid) return;

      let rows = null;
      const nameCandidates = ['Обзвон ВСЕ', 'Обзвон BSE', 'Obzvon ВСЕ', 'Обзвон'];
      for (const nm of nameCandidates) {
        try {
          rows = await fetchSheetCsvByName(sid, nm, nm);
          if (rows?.length > 1) break;
        } catch {}
      }
      if ((!rows || rows.length < 2) && gid) {
        try {
          rows = await fetchSheetCsv(sid, gid, `gid:${gid}`);
        } catch {}
      }
      if (!rows || rows.length < 2) {
        for (const nm of nameCandidates) {
          try {
            rows = await fetchSheetOpenSheet(sid, nm, nm);
            if (rows?.length > 1) break;
          } catch {}
        }
      }

      const parsedAll = parseObzvonAllRows(rows || []);
      const parsed = full ? parsedAll : parsedAll.slice(0, OBZVON_ALL_LIMIT);
      if (parsedAll.length > 0) {
        if (full) {
          setObzvonAllRows(parsedAll);
          S.set('aq-obzvon-all-cache', parsedAll);
          S.set(OBZVON_ALL_INSTALLED_KEY, true);
          setObzvonAllInstalled(true);
        } else {
          setObzvonAllRows(parsed);
          if (!obzvonAllInstalled) S.set('aq-obzvon-all-cache', parsed);
        }
      } else if (cached.length) {
        setObzvonAllRows(cached);
        setObzvonAllInstalled(true);
      }
    } catch (e) {
      const cached = S.get('aq-obzvon-all-cache', []);
      if (cached.length) {
        setObzvonAllRows(cached);
        setObzvonAllInstalled(true);
      }
      notify(`Obzvon ВСЕ yuklanmadi: ${e?.message || 'xato'}`, 'err');
    }
  }, [obzvonSheetUrl, obzvonAllInstalled]);
  const appendObzvonAllRow = useCallback((row) => {
    setObzvonAllRows((prev) => [{
      no: '',
      customer: row.customer,
      callDate: row.callDate,
      topic: row.topic,
      note: row.note,
      nextDate: row.nextDate,
      orderCount: row.orderCount,
      orderDate: row.orderDate,
      operator: row.operator,
      customerId: row.id,
    }, ...prev]);
  }, []);
  const addObzvonRows = useCallback((rows) => {
    if (!rows?.length) return;
    setObzvonRecords((prev) => [...rows, ...prev]);
    rows.forEach((r) => appendObzvonAllRow(r));
  }, [appendObzvonAllRow]);
  const sendToWebhook = useCallback(async () => {
    if (!obzvonWebhook || !obzvonRecords.length) return;
    const payload = obzvonRecords.map((r) => ({
      mijozId: r.id || '',
      mijozIsmi: r.customer || '',
      sana: r.callDate || '',
      maqsad: r.topic || '',
      izoh: r.note || '',
      keyingiSana: r.nextDate || '',
      zakazSoni: r.orderCount || '',
      zakazSanasi: r.orderDate || '',
      operator: r.operator || currentUser || '',
    }));
    try {
      await fetch(obzvonWebhook, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
    } catch {}
  }, [obzvonWebhook, obzvonRecords, currentUser]);

  useEffect(() => {
    loadFromConfig();
    if (obzvonAllInstalled && obzvonAllRows.length > 0) return;
    loadObzvonAll(obzvonSheetUrl || OBZVON_ALL_SHEET_URL, { full: true, force: true });
  }, [loadFromConfig, loadObzvonAll, obzvonAllInstalled, obzvonAllRows.length, obzvonSheetUrl]);
  useEffect(() => {
    const checkAndSync = () => {
      const now = new Date();
      const lastSync = S.get('aq-last-sync', '');
      const todayKey = now.toISOString().slice(0,10);
      if (now.getHours() >= 7 && lastSync !== todayKey) {
        sendToWebhook();
        S.set('aq-last-sync', todayKey);
      }
    };
    const interval = setInterval(checkAndSync, 60000);
    checkAndSync();
    return () => clearInterval(interval);
  }, [sendToWebhook]);

  const rawD = data || { customers:[], orders:[], cashbox:[], contacts:[], rawOrders:[], rawCash:[], ordersByMId:{}, cashByMId:{}, kulerInstallments:[], assignmentById:{} };
  const currentAccess = access[currentUser] || DEFAULT_ACCESS.Admin;
  const scopeOwn = currentAccess.scope === 'own';
  const ownIds = new Set(Object.entries(rawD.assignmentById || {}).filter(([,op]) => op === currentUser).map(([id]) => id));
  const D = useMemo(() => {
    if (!scopeOwn) return rawD;
    const customers = (rawD.customers || []).filter((c) => ownIds.has(c.id));
    const idSet = new Set(customers.map((c)=>c.id));
    const rawOrders = (rawD.rawOrders || []).filter((o) => idSet.has(o.mId));
    const rawCash = (rawD.rawCash || []).filter((c) => idSet.has(c.mId));
    const ordersByMId = {};
    rawOrders.forEach((o)=>{ if(!ordersByMId[o.mId]) ordersByMId[o.mId]=[]; ordersByMId[o.mId].push(o); });
    const cashByMId = {};
    rawCash.forEach((c)=>{ if(!cashByMId[c.mId]) cashByMId[c.mId]=[]; cashByMId[c.mId].push(c); });
    return {
      ...rawD,
      customers,
      rawOrders,
      rawCash,
      ordersByMId,
      cashByMId,
      kulerInstallments: (rawD.kulerInstallments || []).filter((k)=>idSet.has(k.customerId)),
    };
  }, [rawD, scopeOwn, ownIds]);
  const doljniki = useMemo(() => {
    return (D.customers || [])
      .filter((c) => c.balanceUZS < 0)
      .map((c) => ({
        lastOrderProduct: ((D.ordersByMId?.[c.id] || []).find((o) => o.soNum === c.lastDocNum)?.product) || '',
        id: c.id || '—',
        category: c.source || '—',
        name: c.name || '—',
        debtUZS: c.balanceUZS,
        lastOrderDate: c.lastOrderDate || '',
        days: c.daysAgo,
        orderNo: c.lastDocNum || '—',
        qty: c.lastQty || 0,
        lastSum: c.lastSum || 0,
        agent: c.lastAgent || '—',
        note: c.merchantNote || '—',
      }))
      .sort((a, b) => a.debtUZS - b.debtUZS);
  }, [D.customers, D.ordersByMId]);
  const obzvonCnt  = D.customers.filter((c)=>c.daysAgo!=null&&c.daysAgo>14).length;
  const debtorCnt  = D.customers.filter((c)=>c.balanceUZS<0).length;
  const doljnikiCnt = doljniki.length;
  const visibleNav = NAV.filter((n) => (currentAccess.visible?.[n.id] ?? true));
  const pageMeta = NAV.find((n)=>n.id===page) || { id:'settings', icon:'⚙️', label:'Nastroyka' };
  useEffect(() => {
    if (page === 'settings') return;
    if (!visibleNav.find((n) => n.id === page)) setPage(visibleNav[0]?.id || 'dash');
  }, [visibleNav, page]);

  if (!isLoggedIn) {
    return (
      <>
        <style>{CSS}</style>
        <LoginScreen users={users} onLogin={authenticate} onResetCreds={resetLoginCreds} />
      </>
    );
  }

  return (
    <>
      <style>{CSS}</style>
      <div style={{height:'100vh',display:'flex',overflow:'hidden',background:'var(--bg)'}}>
        {/* SIDEBAR */}
        <div style={{width:side?215:56,background:'var(--s1)',borderRight:'1px solid var(--b2)',display:'flex',flexDirection:'column',transition:'width .2s',flexShrink:0,overflow:'hidden'}}>
          <div style={{padding:'13px 11px',borderBottom:'1px solid var(--b2)',display:'flex',alignItems:'center',gap:10}}>
            <div style={{width:30,height:30,background:'linear-gradient(135deg,var(--bl),#1d4ed8)',borderRadius:8,display:'flex',alignItems:'center',justifyContent:'center',fontSize:16,flexShrink:0}}>💧</div>
            {side && <span style={{fontWeight:800,fontSize:14,whiteSpace:'nowrap'}}>AquaBiz Pro</span>}
          </div>
          <div style={{flex:1,padding:'8px 6px',overflowY:'auto',overflowX:'hidden'}}>
            {visibleNav.map((n) => {
              const bc = n.badge==='o'?obzvonCnt:n.badge==='d'?debtorCnt:n.badge==='dz'?doljnikiCnt:0;
              return (
                <div key={n.id} className={`nav-i${page===n.id?' on':''}`} onClick={()=>setPage(n.id)}>
                  <span style={{fontSize:17,flexShrink:0}}>{n.icon}</span>
                  {side && <span style={{flex:1,whiteSpace:'nowrap'}}>{n.label}</span>}
                  {side&&bc>0 && (
                    <span style={{minWidth:17,height:17,borderRadius:9,background:'var(--rd)',color:'#fff',fontSize:9,fontWeight:800,display:'inline-flex',alignItems:'center',justifyContent:'center',padding:'0 4px'}}>{bc}</span>
                  )}
                  {!side&&bc>0 && (
                    <div style={{position:'absolute',top:6,right:6,width:7,height:7,background:'var(--rd)',borderRadius:'50%'}}/>
                  )}
                </div>
              );
            })}
          </div>
          <div style={{padding:6,borderTop:'1px solid var(--b2)'}}>
            <div className="nav-i" onClick={()=>setUp(true)} style={{ opacity: autoLoad.loading ? 0.6 : 1 }}>
              <span style={{fontSize:17,flexShrink:0}}>🔄</span>
              {side && <span>Yangilash</span>}
            </div>
            {(currentAccess.visible?.settings ?? true) && (
              <div className="nav-i" onClick={()=>setPage('settings')}>
                <span style={{fontSize:17,flexShrink:0}}>⚙️</span>
                {side && <span>Nastroyka</span>}
              </div>
            )}
          </div>
        </div>

        {/* MAIN */}
        <div style={{flex:1,display:'flex',flexDirection:'column',overflow:'hidden'}}>
          <div style={{background:'var(--s1)',borderBottom:'1px solid var(--b2)',padding:'0 18px',height:48,display:'flex',alignItems:'center',justifyContent:'space-between',flexShrink:0}}>
            <div style={{display:'flex',alignItems:'center',gap:10}}>
              <button className="nav-i" onClick={()=>setSide(!side)} style={{padding:'4px 8px',border:'none',background:'transparent',cursor:'pointer'}}>
                <span style={{fontSize:17,display:'inline-block',transform:side?'none':'rotate(180deg)',transition:'transform .2s'}}>◀</span>
              </button>
              <span style={{fontWeight:700,fontSize:14}}>
                {pageMeta.icon} {pageMeta.label}
              </span>
              {data && (
                <>
                  <span className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>{D.customers.length} mijoz</span>
                  <span className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>{(D.rawOrders||[]).length} zakaz</span>
                </>
              )}
            </div>
            <div style={{display:'flex',gap:6}}>
              {sessionUser === 'Admin' ? (
                <select className="select" value={currentUser} onChange={(e)=>switchUser(e.target.value)}>
                  {users.map((u)=><option key={u}>{u}</option>)}
                </select>
              ) : (
                <span className="tag" style={{background:'var(--s3)',color:'var(--t2)'}}>{currentUser}</span>
              )}
              <button className="btn btn-gh btn-sm" onClick={logout}>Chiqish</button>
              {obzvonCnt>0 && (
                <button className="btn btn-sm" style={{background:'var(--rd2)',color:'var(--rd)',border:'1px solid var(--rd2)'}} onClick={()=>setPage('obzvon')}>
                  📞 {obzvonCnt} obzvon
                </button>
              )}
            </div>
          </div>

          <div style={{flex:1,overflow:'auto',padding:16}}>
            {!data && page!=='doljniki' ? (
              <div style={{height:'100%',display:'flex',alignItems:'center',justifyContent:'center'}}>
                <div style={{textAlign:'center',maxWidth:440}}>
                  {autoLoad.loading ? (
                    <>
                      <div style={{fontSize:56,marginBottom:16,animation:'spin 1.2s linear infinite',display:'inline-block'}}>⟳</div>
                      <div style={{fontSize:18,fontWeight:700,marginBottom:12}}>Ma'lumot yuklanmoqda</div>
                      <div style={{color:'var(--bl)',fontSize:14,fontFamily:'var(--mono)',marginBottom:20}}>{autoLoad.progress}</div>
                    </>
                  ) : autoLoad.error ? (
                    <>
                      <div style={{fontSize:56,marginBottom:16}}>⚠️</div>
                      <div style={{fontSize:18,fontWeight:700,marginBottom:12,color:'var(--rd)'}}>Yuklash xatosi</div>
                      <div style={{color:'var(--t3)',fontSize:13,marginBottom:8,background:'var(--rd2)',border:'1px solid var(--rd)',borderRadius:9,padding:'12px 16px',textAlign:'left'}}>{autoLoad.error}</div>
                      <div style={{display:'flex',gap:10,justifyContent:'center',marginTop:16}}>
                        <button className="btn btn-bl" onClick={()=>loadFromConfig()}>🔄 Qayta urinish</button>
                        <button className="btn btn-gh" onClick={()=>setUp(true)}>📂 Excel yuklash</button>
                      </div>
                    </>
                  ) : (
                    <>
                      <div style={{fontSize:56,marginBottom:16}}>📊</div>
                      <div style={{fontSize:18,fontWeight:700,marginBottom:12}}>Ma'lumot yuklanmagan</div>
                      <button className="btn btn-gh" onClick={()=>setUp(true)}>📂 Excel fayl yuklash</button>
                    </>
                  )}
                </div>
              </div>
            ) : (
              <>
                {page==='dash'    && <Dashboard D={D}/>}
                {page==='cust'    && <Customers D={rawD} currentUser={currentUser}/>}
                {page==='orders'  && <Orders    D={D}/>}
                {page==='kassa'   && <Kassa     D={D}/>}
                {page==='obzvon'  && <Obzvon    D={D} allRows={obzvonAllRows} onAppendAllRow={appendObzvonAllRow} onReloadAll={()=>loadObzvonAll(obzvonSheetUrl, { full:true, force:true })} webhookUrl={obzvonWebhook} currentUser={currentUser} records={obzvonRecords} setRecords={setObzvonRecords} />}
                {page==='doljniki'&& <Doljniki rows={doljniki} D={D} kulerRows={D.kulerInstallments || []} onAddToObzvon={addObzvonRows} currentUser={currentUser} />}
                {page==='reports' && <Reports   D={D}/>}
                {page==='settings'&& <SettingsPanel users={users} setUsers={setUsers} access={access} setAccess={setAccess} currentUser={currentUser} setCurrentUser={setCurrentUser} webhookUrl={obzvonWebhook} setWebhookUrl={setObzvonWebhook} userCreds={userCreds} setUserCreds={setUserCreds} onSwitchUser={switchUser} isAdminSession={sessionUser==='Admin'} />}
              </>
            )}
          </div>
        </div>
      </div>

      {showUp && (
        <UploadModal
          onLoad={handleLoad}
          onLoadObzvonAll={(url)=>loadObzvonAll(url, { full:true, force:true })}
          hasData={!!data}
          onClose={data?()=>setUp(false):null}
          mainSheetUrl={mainSheetUrl}
          setMainSheetUrl={setMainSheetUrl}
          obzvonSheetUrl={obzvonSheetUrl}
          setObzvonSheetUrl={setObzvonSheetUrl}
        />
      )}
      {notif && (
        <div className="notif" style={{background:notif.type==='ok'?'var(--gr2)':'var(--rd2)',border:`1px solid ${notif.type==='ok'?'var(--gr)':'var(--rd)'}`,color:notif.type==='ok'?'var(--gr)':'var(--rd)'}}>
          {notif.msg}
        </div>
      )}
    </>
  );
}

