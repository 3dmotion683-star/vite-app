import { Fragment, useState, useMemo, useCallback, useEffect, useDeferredValue, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, AreaChart, Area,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
} from 'recharts';

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў EXCLUDED MERCHANT KEYWORDS Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
const EXCLUDED_KEYWORDS = [
  'Dividend','UOS','Rasxod','Rasxodnik','Postavshik',
  'Personal','Dolg','VIP','KULER','Ofis',
];
const isExcludedMerchant = (name) => {
  const nl = (name || '').toLowerCase();
  return EXCLUDED_KEYWORDS.some((k) => nl.includes(k.toLowerCase()));
};
const ALLOWED_AA_KEYWORD = 'murodbaxsh';
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
];
const isAllowedAA = (v) => {
  const s = String(v || '').trim().toLowerCase();
  return s === '' || s.includes(ALLOWED_AA_KEYWORD);
};
const isExcludedZCategory = (v) => {
  const s = String(v || '').trim().toLowerCase();
  return EXCLUDED_Z_CATEGORIES.some((k) => s === k);
};
const normCurrency = (v) => {
  const s = String(v || '').trim().toUpperCase();
  if (['USD', '$'].includes(s)) return 'USD';
  return 'UZS';
};
const isCommonDistrictLabel = (v) => {
  const s = String(v || '').trim().toLowerCase().replace(/\s+/g, ' ');
  if (!s) return false;
  return (
    s.includes('общая территория') ||
    s === 'общая' ||
    s.includes('РѕР±С‰Р°СЏ С‚РµСЂСЂРёС‚РѕСЂРёСЏ') ||
    s === 'РѕР±С‰Р°СЏ' ||
    s.includes('territoriya') ||
    s.includes('territoria') ||
    s.includes('obshaya') ||
    s.includes('obshchaya') ||
    s.includes('common territory')
  );
};
const pickPreferredDistrict = (raw) => {
  const parts = String(raw || '')
    .split(';')
    .map((s) => s.trim())
    .filter(Boolean);
  if (!parts.length) return '';
  const preferred = parts.find((p) => !isCommonDistrictLabel(p));
  return preferred || '';
};
const normalizeAccessApiUrl = (v) => {
  const s = String(v || '').trim();
  if (!s) return OBZVON_WEBHOOK_DEFAULT;
  return /workers\.dev/i.test(s) ? s : OBZVON_WEBHOOK_DEFAULT;
};
function getDebtStats(customers = []) {
  const debtorsUZS = customers.filter((c) => c.balanceUZS < 0);
  return {
    anyCount: debtorsUZS.length,
    uzsCount: debtorsUZS.length,
    uzsSum: debtorsUZS.reduce((s, c) => s + Math.abs(c.balanceUZS), 0),
  };
}
function addMonthsByAnchorDay(date, monthsToAdd) {
  const src = new Date(date);
  if (Number.isNaN(src.getTime())) return null;
  const targetMonthIndex = src.getMonth() + Number(monthsToAdd || 0);
  const targetYear = src.getFullYear() + Math.floor(targetMonthIndex / 12);
  const targetMonth = ((targetMonthIndex % 12) + 12) % 12;
  const lastDay = new Date(targetYear, targetMonth + 1, 0).getDate();
  const day = Math.min(src.getDate(), lastDay);
  const out = new Date(targetYear, targetMonth, day);
  out.setHours(0, 0, 0, 0);
  return out;
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
      now.setHours(0, 0, 0, 0);
      let shouldPayMonths = 0;
      for (let i = 1; i <= months; i++) {
        const dueDate = addMonthsByAnchorDay(start, i);
        if (!dueDate) break;
        if (now >= dueDate) shouldPayMonths = i;
        else break;
      }
      const shouldPaid = Math.min(principal, shouldPayMonths * monthly);
      overdueAmount = Math.max(0, Math.min(principal, shouldPaid) - paid);
      dueCount = overdueAmount > 0 ? Math.max(1, shouldPayMonths - paidMonths) : 0;
    }
  }
  return { ...row, months, monthly, remaining, paidMonths, monthsLeft, dueCount, overdueAmount };
}

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў CONFIG Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
const SHEET_CONFIG = {
  url: 'https://docs.google.com/spreadsheets/d/1RND6D5JIWh6vnk8i_FO1mx1if4hx182VNaPzGQtKuQM/edit?gid=1272423090#gid=1272423090',
  gids: { merchants: '', balans: '', orders: '', cashbox: '', integration: '', mijozlar: '', warehouseTransfer: '' },
};
const OBZVON_ALL_LOCAL_XLSX = '/obzvon-vse.xlsx';
const OBZVON_ALL_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1lfjqNFaD2Gy-tyKrmWVX4ja2DvsyLcACaGlW2ZJrkf8/edit?pli=1&gid=0#gid=0';
const OBZVON_NEW_EXPORT_SHEET_ID = '1lfjqNFaD2Gy-tyKrmWVX4ja2DvsyLcACaGlW2ZJrkf8';
const OBZVON_NEW_EXPORT_SHEET_NAME = 'Barcha_obzvon_yangi';
const OBZVON_NEW_EXPORT_SHEET_GID = '979057575';
const OBZVON_NEW_EXPORT_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1lfjqNFaD2Gy-tyKrmWVX4ja2DvsyLcACaGlW2ZJrkf8/edit?pli=1&gid=979057575#gid=979057575';
const OBZVON_NEW_EXPORT_HOUR = 3;
const OBZVON_NEW_EXPORTED_MAP_KEY = 'aq-obzvon-new-exported-map';
const OBZVON_NEW_LAST_AUTO_DATE_KEY = 'aq-obzvon-new-last-auto-date';
const OBZVON_NEW_HEADERS_WRITTEN_KEY = 'aq-obzvon-new-sheet-headers-written';
const OBZVON_NEW_STABLE_ID_MAP_KEY = 'aq-obzvon-new-stable-id-map';
const OBZVON_NEW_STABLE_ID_LAST_KEY = 'aq-obzvon-new-stable-id-last';
const ACCESS_SYNC_URL_KEY = 'aq-access-api-url';
const ACCESS_UPDATED_AT_KEY = 'aq-access-updated-at';
// Access konfiguratsiya sinxroni uchun Cloudflare Worker API URL:
const OBZVON_WEBHOOK_DEFAULT = 'https://solitary-brook-4889aquabiz-api.3dmotion683.workers.dev';
// Google Sheetga yozish uchun fallback Apps Script URL
const OBZVON_EXPORT_APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyi1OP_a_5C7-TBujLuo9dDast0RLVelhQsTiO6dlN_mefi55vnTHm_ZRXpDFFTTXb7qA/exec';

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў HELPERS Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
const fmt  = (n) => new Intl.NumberFormat('uz-UZ').format(Math.round(n || 0));
const fmtM = (n) =>
  n >= 1e6 ? (n / 1e6).toFixed(1) + 'M' :
  n >= 1e3 ? (n / 1e3).toFixed(0) + 'K' : fmt(n);
const E = {
  upload: '\u{1F4C2}',
  excel: '\u{1F4CA}',
  sheets: '\u{1F517}',
  ok: '\u2705',
  find: '\u{1F50D}',
  refresh: '\u21BB',
  pay: '\u{1F4B0}',
  top: '\u{1F3C6}',
  late: '\u{1F55B}',
  order: '\u{1F4E6}',
  area: '\u{1F5FA}',
  my: '\u{1F464}',
  all: '\u{1F465}',
  driver: '\u{1F69A}',
  settings: '\u2699\uFE0F',
  ret: '\u21A9',
  home: '\u{1F3E0}',
  users: '\u{1F465}',
  phone: '\u{1F4DE}',
  doc: '\u{1F4C4}',
  report: '\u{1F4CA}',
  plan: '📆',
  water: '\u{1F4A7}',
  uzs: '\u{1F4B4}',
  usd: '\u{1F4B5}',
  tara: '\u{1F9F4}',
  sv: '\u{1F4CB}',
};
const FILTER_CHECK_LABEL_STYLE = {
  display: 'grid',
  gridTemplateColumns: '14px minmax(0,1fr)',
  gap: 6,
  fontSize: 12,
  alignItems: 'start',
  minWidth: 0,
  padding: '3px 0',
};
const FILTER_CHECK_TEXT_STYLE = {
  display: 'block',
  minWidth: 0,
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
};
const UNIVERSAL_FILTER_MODES = [
  { value: 'show', label: "Ko'rsatsin" },
  { value: 'hide', label: "Ko'rsatmasin" },
  { value: 'to', label: 'Gacha' },
  { value: 'from', label: 'Keyin' },
];
const makeUniversalColState = () => ({
  mode: 'show',
  value: '',
  selected: [],
  optionsOpen: false,
  optionQuery: '',
});
const ensureUniversalFilterState = (columns = [], prev = {}) => {
  const out = { ...(prev || {}) };
  (columns || []).forEach((c) => {
    if (!out[c.key]) out[c.key] = makeUniversalColState();
    else {
      out[c.key] = {
        ...makeUniversalColState(),
        ...out[c.key],
        selected: Array.isArray(out[c.key].selected) ? out[c.key].selected : [],
      };
    }
  });
  return out;
};
const countUniversalFilters = (state = {}) =>
  Object.values(state || {}).reduce((s, st) => {
    const hasSel = Array.isArray(st?.selected) && st.selected.length > 0;
    const hasVal = String(st?.value ?? '').trim() !== '';
    return s + (hasSel || hasVal ? 1 : 0);
  }, 0);
const getUniversalColValue = (row, col) => {
  if (!col) return '';
  if (typeof col.getValue === 'function') return col.getValue(row);
  return row?.[col.key];
};
const normalizeUniversalOptionValue = (value, type = 'text') => {
  if (type === 'number') {
    const n = Number(value || 0);
    return Number.isFinite(n) ? String(n) : '';
  }
  if (type === 'date') {
    const d = toDate(value);
    return d ? d.toISOString().slice(0, 10) : '';
  }
  return String(value ?? '').trim();
};
const universalRowPasses = (row, columns = [], state = {}) => {
  for (const col of (columns || [])) {
    const st = state?.[col.key];
    if (!st) continue;
    const mode = st.mode || 'show';
    const selected = Array.isArray(st.selected) ? st.selected : [];
    const rawInput = String(st.value ?? '').trim();
    const hasSelected = selected.length > 0;
    const hasInput = rawInput !== '';
    if (!hasSelected && !hasInput) continue;

    const type = col.type || 'text';
    const raw = getUniversalColValue(row, col);

    if (type === 'number') {
      const n = Number(raw || 0);
      const selectedNums = selected.map((x) => Number(x)).filter((x) => Number.isFinite(x));
      if (hasSelected) {
        const inSet = selectedNums.includes(n);
        if (mode === 'hide' ? inSet : !inSet) return false;
        continue;
      }
      const target = Number(rawInput);
      if (!Number.isFinite(target)) continue;
      if (mode === 'from' && !(n >= target)) return false;
      else if (mode === 'to' && !(n <= target)) return false;
      else if (mode === 'hide' && n === target) return false;
      else if (mode === 'show' && n !== target) return false;
      continue;
    }

    if (type === 'date') {
      const d = toDate(raw);
      if (!d) return false;
      const dKey = d.toISOString().slice(0, 10);
      if (hasSelected) {
        const inSet = selected.includes(dKey);
        if (mode === 'hide' ? inSet : !inSet) return false;
        continue;
      }
      const t = toDate(rawInput);
      if (!t) continue;
      const tMs = t.getTime();
      const dMs = d.getTime();
      if (mode === 'from' && !(dMs >= tMs)) return false;
      else if (mode === 'to' && !(dMs <= tMs)) return false;
      else if (mode === 'hide' && dKey === t.toISOString().slice(0,10)) return false;
      else if (mode === 'show' && dKey !== t.toISOString().slice(0,10)) return false;
      continue;
    }

    const txt = String(raw ?? '').trim().toLowerCase();
    if (hasSelected) {
      const normalizedSelected = selected.map((x) => String(x ?? '').trim().toLowerCase());
      const inSet = normalizedSelected.includes(txt);
      if (mode === 'hide' ? inSet : !inSet) return false;
      continue;
    }
    const q = rawInput.toLowerCase();
    if (mode === 'hide' && txt.includes(q)) return false;
    else if (mode === 'from' && !(txt >= q)) return false;
    else if (mode === 'to' && !(txt <= q)) return false;
    else if (mode === 'show' && !txt.includes(q)) return false;
  }
  return true;
};
const applyUniversalFilters = (rows = [], columns = [], state = {}) =>
  (rows || []).filter((r) => universalRowPasses(r, columns, state));

function UniversalFilterPanel({
  open,
  title = 'Universal filtr',
  columns = [],
  rows = [],
  state = {},
  setState = () => {},
  onClose = () => {},
  width = 560,
}) {
  const optionsByKey = useMemo(() => {
    const out = {};
    (columns || []).forEach((col) => {
      const set = new Set();
      (rows || []).forEach((row) => {
        const raw = getUniversalColValue(row, col);
        const val = normalizeUniversalOptionValue(raw, col.type || 'text');
        if (val !== '') set.add(val);
      });
      const arr = Array.from(set);
      if ((col.type || 'text') === 'number') arr.sort((a, b) => Number(a) - Number(b));
      else arr.sort((a, b) => a.localeCompare(b));
      out[col.key] = arr;
    });
    return out;
  }, [columns, rows]);
  const panelRef = useRef(null);
  const [placement, setPlacement] = useState(() => ({
    side: 'right', // right => panel opens to right side of button (left:0)
    panelWidth: width,
    maxHeight: '56vh',
  }));
  const recalcPlacement = useCallback(() => {
    if (!open) return;
    const panelEl = panelRef.current;
    const anchorEl = panelEl?.parentElement;
    if (!anchorEl) return;
    const anchorRect = anchorEl.getBoundingClientRect();
    const vw = window.innerWidth || document.documentElement.clientWidth || 1280;
    const vh = window.innerHeight || document.documentElement.clientHeight || 800;
    const margin = 8;
    let boundary = {
      left: margin,
      right: vw - margin,
      top: margin,
      bottom: vh - margin,
    };
    const boundaryEl = anchorEl.closest('[data-filter-boundary="1"]');
    if (boundaryEl) {
      const b = boundaryEl.getBoundingClientRect();
      boundary = {
        left: Math.max(margin, b.left + 4),
        right: Math.min(vw - margin, b.right - 4),
        top: Math.max(margin, b.top + 4),
        bottom: Math.min(vh - margin, b.bottom - 4),
      };
    }
    const desiredWidth = Number(width) || 560;
    const maxAllowedWidth = Math.max(220, Math.floor(boundary.right - boundary.left));
    let panelWidth = Math.min(desiredWidth, maxAllowedWidth);
    const spaceRight = Math.floor(boundary.right - anchorRect.left);
    const spaceLeft = Math.floor(anchorRect.right - boundary.left);
    let side = 'right';
    if (panelWidth <= spaceRight) {
      side = 'right';
    } else if (panelWidth <= spaceLeft) {
      side = 'left';
    } else if (spaceLeft > spaceRight) {
      side = 'left';
      panelWidth = Math.max(220, Math.min(spaceLeft, maxAllowedWidth));
    } else {
      side = 'right';
      panelWidth = Math.max(220, Math.min(spaceRight, maxAllowedWidth));
    }
    const maxHeightPx = Math.max(260, Math.floor(boundary.bottom - anchorRect.top - 60));
    setPlacement({
      side,
      panelWidth,
      maxHeight: `${maxHeightPx}px`,
    });
  }, [open, width]);
  useEffect(() => {
    if (!open) return undefined;
    const run = () => recalcPlacement();
    const raf = requestAnimationFrame(run);
    window.addEventListener('resize', run);
    window.addEventListener('scroll', run, true);
    return () => {
      cancelAnimationFrame(raf);
      window.removeEventListener('resize', run);
      window.removeEventListener('scroll', run, true);
    };
  }, [open, recalcPlacement]);
  useEffect(() => {
    if (!open) return undefined;
    const onEsc = (e) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onEsc);
    return () => window.removeEventListener('keydown', onEsc);
  }, [open, onClose]);
  useEffect(() => {
    if (!open) return undefined;
    const panelEl = panelRef.current;
    const anchorEl = panelEl?.parentElement;
    const boundaryEl = anchorEl?.closest('[data-filter-boundary="1"]');
    const prevHtmlOverflow = document.documentElement.style.overflow;
    const prevBodyOverflow = document.body.style.overflow;
    const prevBodyOverscroll = document.body.style.overscrollBehavior;
    const prevBoundaryOverflow = boundaryEl ? boundaryEl.style.overflow : '';
    const prevBoundaryOverscroll = boundaryEl ? boundaryEl.style.overscrollBehavior : '';

    document.documentElement.style.overflow = 'hidden';
    document.body.style.overflow = 'hidden';
    document.body.style.overscrollBehavior = 'none';
    if (boundaryEl) {
      boundaryEl.style.overflow = 'hidden';
      boundaryEl.style.overscrollBehavior = 'none';
    }

    return () => {
      document.documentElement.style.overflow = prevHtmlOverflow;
      document.body.style.overflow = prevBodyOverflow;
      document.body.style.overscrollBehavior = prevBodyOverscroll;
      if (boundaryEl) {
        boundaryEl.style.overflow = prevBoundaryOverflow;
        boundaryEl.style.overscrollBehavior = prevBoundaryOverscroll;
      }
    };
  }, [open]);
  useEffect(() => {
    if (!open) return undefined;
    const handleDown = (e) => {
      const panelEl = panelRef.current;
      if (panelEl && !panelEl.contains(e.target)) onClose();
    };
    document.addEventListener('mousedown', handleDown, true);
    document.addEventListener('touchstart', handleDown, true);
    return () => {
      document.removeEventListener('mousedown', handleDown, true);
      document.removeEventListener('touchstart', handleDown, true);
    };
  }, [open, onClose]);

  if (!open) return null;
  const safeState = ensureUniversalFilterState(columns, state);
  const upd = (key, patch) => {
    setState((prev) => ({
      ...ensureUniversalFilterState(columns, prev),
      [key]: { ...(ensureUniversalFilterState(columns, prev)[key] || makeUniversalColState()), ...patch },
    }));
  };
  const setAll = (mode = 'select') => {
    setState((prev) => {
      const next = ensureUniversalFilterState(columns, prev);
      columns.forEach((c) => {
        next[c.key] = {
          ...(next[c.key] || makeUniversalColState()),
          selected: mode === 'select' ? [...(optionsByKey[c.key] || [])] : [],
          value: mode === 'select' ? '' : '',
        };
      });
      return { ...next };
    });
  };

  return (
    <>
      <div
        style={{position:'fixed',inset:0,zIndex:79,background:'transparent',touchAction:'none'}}
        onClick={onClose}
        onMouseDown={onClose}
        onTouchStart={onClose}
        onWheel={(e)=>e.preventDefault()}
      />
      <div
        ref={panelRef}
        className="card"
        style={{
          position:'absolute',
          top:40,
          ...(placement.side === 'left' ? { right:0, left:'auto' } : { left:0, right:'auto' }),
          zIndex:80,
          width:placement.panelWidth,
          maxWidth:'calc(100vw - 16px)',
          padding:12,
          boxShadow:'0 24px 60px rgba(0,0,0,.6)',
          backdropFilter:'blur(10px)',
          overflow:'hidden',
        }}
        onMouseDown={(e)=>e.stopPropagation()}
        onClick={(e)=>e.stopPropagation()}
        onTouchStart={(e)=>e.stopPropagation()}
        onWheelCapture={(e)=>e.stopPropagation()}
      >
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:8}}>
          <div style={{fontWeight:700,fontSize:13}}>{title}</div>
          <div style={{display:'flex',gap:6}}>
            <button className="btn btn-gh btn-sm" onClick={()=>setAll('select')}>Hammasini belgilash</button>
            <button className="btn btn-gh btn-sm" onClick={()=>setAll('clear')}>Hammasini tozalash</button>
          </div>
        </div>
        <div
          style={{
            maxHeight:placement.maxHeight,
            overflow:'auto',
            paddingRight:2,
            overscrollBehavior:'contain',
            overscrollBehaviorY:'contain',
            WebkitOverflowScrolling:'touch',
          }}
          onWheel={(e)=>e.stopPropagation()}
          onTouchMove={(e)=>e.stopPropagation()}
        >
          {columns.map((col) => {
            const st = safeState[col.key] || makeUniversalColState();
            const opts = optionsByKey[col.key] || [];
            const shownOpts = opts.filter((o) => o.toLowerCase().includes(String(st.optionQuery || '').toLowerCase()));
            return (
              <div key={col.key} className="card" style={{padding:8,marginBottom:8,background:'var(--s3)',border:'1px solid var(--b1)'}}>
                <div style={{fontSize:10.5,color:'var(--t3)',fontWeight:700,marginBottom:6}}>{col.label}</div>
                <div style={{display:'grid',gridTemplateColumns:'120px 1fr auto',gap:6,marginBottom:6}}>
                  <select className="select" value={st.mode || 'show'} onChange={(e)=>upd(col.key, { mode:e.target.value })}>
                    {UNIVERSAL_FILTER_MODES.map((m)=><option key={m.value} value={m.value}>{m.label}</option>)}
                  </select>
                  {col.type === 'date' ? (
                    <ModernDateInput
                      value={st.value || ''}
                      placeholder="sanani tanlang"
                      onChange={(e)=>upd(col.key, { value:e.target.value })}
                    />
                  ) : (
                    <input
                      className="input"
                      type={col.type === 'number' ? 'number' : 'text'}
                      value={st.value || ''}
                      placeholder={col.type === 'number' ? 'son kiriting' : 'qiymat yozing'}
                      onChange={(e)=>upd(col.key, { value:e.target.value })}
                    />
                  )}
                  <button className="btn btn-gh btn-sm" onClick={()=>upd(col.key, { optionsOpen: !st.optionsOpen })}>{st.optionsOpen ? 'Yop' : "Ro'yxat"}</button>
                </div>
                {st.optionsOpen && (
                  <div className="card" style={{padding:8,border:'1px solid var(--b1)'}}>
                    <div style={{display:'grid',gridTemplateColumns:'1fr auto auto',gap:6,marginBottom:6}}>
                      <input className="input" placeholder="Qidirish..." value={st.optionQuery || ''} onChange={(e)=>upd(col.key, { optionQuery: e.target.value })} />
                      <button className="btn btn-gh btn-sm" onClick={()=>upd(col.key, { selected:[...opts] })}>Hammasi</button>
                      <button className="btn btn-gh btn-sm" onClick={()=>upd(col.key, { selected:[] })}>Tozalash</button>
                    </div>
                    <div style={{maxHeight:120,overflow:'auto',overscrollBehavior:'contain'}}>
                      {shownOpts.map((opt) => {
                        const checked = (st.selected || []).includes(opt);
                        return (
                          <label key={opt} style={FILTER_CHECK_LABEL_STYLE}>
                            <input
                              type="checkbox"
                              checked={checked}
                              onChange={(e)=>upd(col.key, {
                                selected: e.target.checked
                                  ? [...(st.selected || []), opt]
                                  : (st.selected || []).filter((x)=>x!==opt),
                              })}
                            />
                            <span style={{...FILTER_CHECK_TEXT_STYLE,color:checked?'var(--bl)':'var(--t2)'}}>{opt}</span>
                          </label>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
        <div style={{display:'flex',justifyContent:'flex-end',gap:6}}>
          <button className="btn btn-gh btn-sm" onClick={onClose}>Yopish</button>
          <button className="btn btn-bl btn-sm" onClick={onClose}>Qo'llash</button>
        </div>
      </div>
    </>
  );
}
const createRowRid = () => `r_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`;
const hash32 = (input) => {
  const s = String(input || '');
  let h = 0;
  for (let i = 0; i < s.length; i += 1) {
    h = ((h << 5) - h + s.charCodeAt(i)) | 0;
  }
  return Math.abs(h).toString(36);
};
const buildFallbackObzvonRid = (row, i = 0) => {
  const src = row || {};
  const seed = [
    src.customerId || src.id || '',
    src.customer || '',
    src.callDate || '',
    src.topic || '',
    src.note || '',
    src.nextDate || '',
    src.orderCount || '',
    src.orderDate || '',
    src.operator || '',
    src.company || '',
    src.updatedAt || '',
  ].map((x) => String(x || '').trim()).join('|');
  if (!seed) return `nr_${i}`;
  return `nr_${hash32(seed)}`;
};
const isInvalidObzvonCustomerText = (v) => {
  const s = String(v ?? '').trim();
  if (!s) return true;
  const up = s.toUpperCase();
  return (
    up.startsWith('#') ||
    up === '#REF!' ||
    up === '#REF' ||
    up === '#N/A' ||
    up === 'N/A' ||
    up === 'NA' ||
    up === '#NA' ||
    up === 'NULL' ||
    up === 'UNDEFINED'
  );
};
const sanitizeObzvonCell = (v) => {
  const s = String(v ?? '').trim();
  return isInvalidObzvonCustomerText(s) ? '' : s;
};
const OBZVON_DELETED_NOTE = '__AQ_DELETED__';
const isObzvonDeletedRow = (row) => {
  const note = String(row?.note || '').trim();
  return note === OBZVON_DELETED_NOTE || String(row?.deleted || '') === '1';
};
const pickObzvonCustomerName = (row) => {
  const fromJ = sanitizeObzvonCell(row?.[9]);
  const fromB = sanitizeObzvonCell(row?.[1]);
  return fromJ || fromB;
};
const isNameExcludedForActiveStats = (name) => {
  const n = String(name || '').trim().toLowerCase();
  if (!n) return true;
  if (isNameInactiveByPrefix(n)) return true;
  if (/^(k|к)\b/.test(n)) return true;
  return false;
};
const isNameInactiveByPrefix = (name) => {
  const n = String(name || '').trim().toLowerCase().replace(/\s+/g, ' ');
  if (!n) return false;
  return /^(я|ya|ря|рї)\s+(tugatildi|eski)\b/.test(n);
};
const isAhmadteaTag = (v) => {
  const s = normText(v).replace(/[\s._-]+/g, '');
  return s.includes('ahmadtea');
};
const isAhmadteaCustomer = (c) => {
  if (!c) return false;
  // Faqat AA kategoriyasi asosida ajratamiz.
  return isAhmadteaTag(c.aaTag);
};
const isMurodbaxshCustomer = (c) => {
  return !isAhmadteaCustomer(c);
};
const COMPANY_OPTIONS = [
  { key: 'murodbaxsh', label: 'Murodbaxsh' },
  { key: 'ahmadtea', label: 'Ahmadtea' },
];
const normalizeCompanyKey = (v) => (String(v || '').trim().toLowerCase() === 'ahmadtea' ? 'ahmadtea' : 'murodbaxsh');
const normalizeStoredCompanyKey = (v) => {
  const s = String(v || '').trim();
  return s ? normalizeCompanyKey(s) : '';
};
const companyLabelByKey = (v) => (normalizeCompanyKey(v) === 'ahmadtea' ? 'Ahmadtea' : 'Murodbaxsh');
const normalizeIdKey = (v) => String(v ?? '').trim();
const filterDataByCustomerIds = (baseData, idSetInput) => {
  const source = baseData || {};
  const ids = new Set(
    Array.from(idSetInput || [])
      .map((v) => normalizeIdKey(v))
      .filter(Boolean)
  );
  const hasId = (v) => ids.has(normalizeIdKey(v));
  const customers = (source.customers || []).filter((c) => hasId(c?.id));
  const contacts = (source.contacts || []).filter((c) => hasId(c?.id));
  const rawOrders = (source.rawOrders || source.orders || []).filter((o) => hasId(o?.mId));
  const rawCash = (source.rawCash || source.cashbox || []).filter((c) => hasId(c?.mId));
  const ordersByMId = {};
  rawOrders.forEach((o) => {
    const key = normalizeIdKey(o?.mId);
    if (!key) return;
    if (!ordersByMId[key]) ordersByMId[key] = [];
    ordersByMId[key].push(o);
  });
  const cashByMId = {};
  rawCash.forEach((c) => {
    const key = normalizeIdKey(c?.mId);
    if (!key) return;
    if (!cashByMId[key]) cashByMId[key] = [];
    cashByMId[key].push(c);
  });
  const assignmentById = {};
  Object.entries(source.assignmentById || {}).forEach(([id, op]) => {
    if (hasId(id)) assignmentById[id] = op;
  });

  return {
    ...source,
    customers,
    contacts,
    orders: rawOrders,
    cashbox: rawCash,
    rawOrders,
    rawCash,
    ordersByMId,
    cashByMId,
    assignmentById,
    kulerInstallments: (source.kulerInstallments || []).filter((k) => hasId(k?.customerId)),
    debtorsByBalance: (source.debtorsByBalance || []).filter((c) => hasId(c?.id)),
    otherDebtorsByBalance: (source.otherDebtorsByBalance || []).filter((c) => hasId(c?.id)),
  };
};
const filterDataByCompany = (baseData, company = 'murodbaxsh') => {
  const source = baseData || {};
  const useAhmadtea = company === 'ahmadtea';
  const idSet = new Set(
    (source.customers || [])
      .filter((c) => (useAhmadtea ? isAhmadteaCustomer(c) : isMurodbaxshCustomer(c)))
      .map((c) => normalizeIdKey(c?.id))
      .filter(Boolean)
  );
  return filterDataByCustomerIds(source, idSet);
};
const isActiveCustomerName = (name) => {
  const n = String(name || '').trim();
  if (!n) return false;
  if (isNameExcludedForActiveStats(n)) return false;
  return /^\d/.test(n);
};

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
const fmtD = (v) => { const d = toDate(v); return d ? d.toLocaleDateString('ru-RU') : '-'; };
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
const normalizeMonthKey = (v) => {
  const m = String(v || '').match(/^(\d{4})-(\d{2})$/);
  return m ? `${m[1]}-${m[2]}` : '';
};
const buildDefaultPlanRows = (baseDate = new Date()) => {
  const start = new Date(baseDate.getFullYear(), baseDate.getMonth(), 1);
  const rows = [];
  for (let i = 0; i < 12; i++) {
    const d = new Date(start.getFullYear(), start.getMonth() + i, 1);
    rows.push({
      month: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`,
      waterPlan: 0,
      coolerPlan: 0,
      workDays: 26,
    });
  }
  return rows;
};
const normalizePlanRows = (rows = [], baseDate = new Date()) => {
  const byMonth = {};
  (rows || []).forEach((r) => {
    const mk = normalizeMonthKey(r?.month);
    if (!mk) return;
    byMonth[mk] = {
      month: mk,
      waterPlan: Math.max(0, Math.round(Number(r?.waterPlan || 0))),
      coolerPlan: Math.max(0, Math.round(Number(r?.coolerPlan || 0))),
      workDays: Math.max(1, Math.round(Number(r?.workDays || 26))),
    };
  });
  return buildDefaultPlanRows(baseDate).map((r) => byMonth[r.month] || r);
};
const WEEKDAY_OPTIONS = [
  { key: 1, label: 'Du' },
  { key: 2, label: 'Se' },
  { key: 3, label: 'Cho' },
  { key: 4, label: 'Pa' },
  { key: 5, label: 'Ju' },
  { key: 6, label: 'Sha' },
  { key: 0, label: 'Yak' },
];
const normalizeOffDaysConfig = (days = []) => {
  const rawWeekdays = Array.isArray(days)
    ? days
    : Array.isArray(days?.weekdays)
    ? days.weekdays
    : [];
  const rawDates = Array.isArray(days) ? [] : (Array.isArray(days?.dates) ? days.dates : []);

  const weekdaySet = new Set(
    rawWeekdays
      .map((d) => Number(d))
      .filter((d) => Number.isInteger(d) && d >= 0 && d <= 6)
  );
  const dateSet = new Set(
    rawDates
      .map((d) => String(d || '').trim())
      .filter((d) => /^\d{4}-\d{2}-\d{2}$/.test(d))
  );

  return {
    weekdays: WEEKDAY_OPTIONS.map((x) => x.key).filter((k) => weekdaySet.has(k)),
    dates: Array.from(dateSet).sort((a, b) => a.localeCompare(b)),
  };
};
const normalizeOffDays = (days = []) => normalizeOffDaysConfig(days).weekdays;
const toOffDaysStorage = (days = []) => {
  const cfg = normalizeOffDaysConfig(days);
  return { weekdays: cfg.weekdays, dates: cfg.dates };
};
const isWorkDay = (date, offDaysSet = new Set(), offDatesSet = new Set()) => {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  const iso = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  return !offDaysSet.has(d.getDay()) && !offDatesSet.has(iso);
};
const countWorkDays = (fromDate, toDate, offDaysSet = new Set(), offDatesSet = new Set()) => {
  const from = new Date(fromDate);
  const to = new Date(toDate);
  from.setHours(0, 0, 0, 0);
  to.setHours(0, 0, 0, 0);
  if (Number.isNaN(from.getTime()) || Number.isNaN(to.getTime())) return 0;
  if (from > to) return 0;
  let count = 0;
  const cur = new Date(from);
  while (cur <= to) {
    if (isWorkDay(cur, offDaysSet, offDatesSet)) count += 1;
    cur.setDate(cur.getDate() + 1);
  }
  return count;
};
const findNextWorkDate = (fromDate, toDate, offDaysSet = new Set(), offDatesSet = new Set()) => {
  const from = new Date(fromDate);
  const to = new Date(toDate);
  from.setHours(0, 0, 0, 0);
  to.setHours(0, 0, 0, 0);
  if (Number.isNaN(from.getTime()) || Number.isNaN(to.getTime())) return null;
  if (from > to) return null;
  const cur = new Date(from);
  while (cur <= to) {
    if (isWorkDay(cur, offDaysSet, offDatesSet)) return new Date(cur);
    cur.setDate(cur.getDate() + 1);
  }
  return null;
};
const toIntFromMixed = (v) => {
  const s = String(v ?? '').trim();
  if (!s) return 0;
  const parts = s.match(/\d+/g);
  if (!parts || !parts.length) return 0;
  const joined = parts.join('');
  if (!joined) return 0;
  const n = Number(joined);
  if (!Number.isFinite(n)) return 0;
  return s.includes('-') ? -n : n;
};
const toIsoDate = (v) => {
  const d = toDate(v);
  return d ? `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}` : '';
};
const shiftMonthKey = (monthKeyValue, delta = 0) => {
  const m = String(monthKeyValue || '').match(/^(\d{4})-(\d{2})$/);
  const base = m ? new Date(Number(m[1]), Number(m[2]) - 1 + Number(delta || 0), 1) : new Date();
  return `${base.getFullYear()}-${String(base.getMonth() + 1).padStart(2, '0')}`;
};
const buildCalendarMonth = (monthKeyValue) => {
  const m = String(monthKeyValue || '').match(/^(\d{4})-(\d{2})$/);
  const base = m ? new Date(Number(m[1]), Number(m[2]) - 1, 1) : new Date();
  base.setHours(0, 0, 0, 0);
  const start = new Date(base);
  const mondayIdx = (start.getDay() + 6) % 7;
  start.setDate(start.getDate() - mondayIdx);
  const cells = [];
  for (let i = 0; i < 42; i += 1) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    cells.push({
      iso: toIsoDate(d),
      day: d.getDate(),
      inMonth: d.getMonth() === base.getMonth() && d.getFullYear() === base.getFullYear(),
      weekday: d.getDay(),
    });
  }
  return {
    monthLabel: base.toLocaleDateString('uz-UZ', { month: 'long', year: 'numeric' }),
    cells,
  };
};
const isDateInRange = (iso, min, max) => {
  if (!iso) return false;
  if (min && iso < String(min)) return false;
  if (max && iso > String(max)) return false;
  return true;
};
function ModernDateInput({
  value = '',
  onChange = () => {},
  placeholder = 'Sanani tanlang',
  disabled = false,
  style,
  min = '',
  max = '',
}) {
  const wrapRef = useRef(null);
  const selectedIso = toIsoDate(value);
  const [open, setOpen] = useState(false);
  const [monthKeyValue, setMonthKeyValue] = useState(() => (selectedIso ? selectedIso.slice(0, 7) : toIsoDate(new Date()).slice(0, 7)));
  useEffect(() => {
    if (!selectedIso) return;
    setMonthKeyValue(selectedIso.slice(0, 7));
  }, [selectedIso]);
  useEffect(() => {
    if (!open) return undefined;
    const onDown = (e) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) setOpen(false);
    };
    const onEsc = (e) => {
      if (e.key === 'Escape') setOpen(false);
    };
    document.addEventListener('mousedown', onDown, true);
    document.addEventListener('touchstart', onDown, true);
    window.addEventListener('keydown', onEsc);
    return () => {
      document.removeEventListener('mousedown', onDown, true);
      document.removeEventListener('touchstart', onDown, true);
      window.removeEventListener('keydown', onEsc);
    };
  }, [open]);
  const calendar = useMemo(() => buildCalendarMonth(monthKeyValue), [monthKeyValue]);
  const todayIso = toIsoDate(new Date());
  const emit = (nextIso) => {
    onChange?.({ target: { value: String(nextIso || '') } });
  };
  return (
    <div ref={wrapRef} className="modern-date-wrap" style={style} onMouseDown={(e) => e.stopPropagation()} onClick={(e) => e.stopPropagation()}>
      <button
        type="button"
        className="input modern-date-btn"
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
      >
        <span className={`modern-date-label${selectedIso ? '' : ' ph'}`}>{selectedIso ? fmtD(selectedIso) : placeholder}</span>
        <span className="modern-date-icon" aria-hidden="true">📅</span>
      </button>
      {open && (
        <div className="modern-date-pop card">
          <div className="modern-date-head">
            <button type="button" className="btn btn-gh btn-sm" onClick={() => setMonthKeyValue((prev) => shiftMonthKey(prev, -1))}>{'<'}</button>
            <div className="modern-date-month">{calendar.monthLabel}</div>
            <button type="button" className="btn btn-gh btn-sm" onClick={() => setMonthKeyValue((prev) => shiftMonthKey(prev, 1))}>{'>'}</button>
          </div>
          <div className="modern-date-week">
            {WEEKDAY_OPTIONS.map((d) => <span key={`wk_${d.key}`}>{d.label}</span>)}
          </div>
          <div className="modern-date-grid">
            {calendar.cells.map((c) => {
              const enabled = !disabled && isDateInRange(c.iso, min, max);
              const isSel = selectedIso === c.iso;
              const isToday = c.iso === todayIso;
              return (
                <button
                  key={`cd_${c.iso}`}
                  type="button"
                  disabled={!enabled}
                  className={`modern-date-day${c.inMonth ? '' : ' out'}${isSel ? ' on' : ''}${isToday ? ' today' : ''}`}
                  onClick={() => {
                    emit(c.iso);
                    setOpen(false);
                  }}
                  title={c.iso}
                >
                  {c.day}
                </button>
              );
            })}
          </div>
          <div className="modern-date-actions">
            <button type="button" className="btn btn-gh btn-sm" onClick={() => emit('')}>Tozalash</button>
            <button type="button" className="btn btn-bl btn-sm" onClick={() => setOpen(false)}>Yopish</button>
          </div>
        </div>
      )}
    </div>
  );
}
const normId = (v) => {
  const s = String(v ?? '').trim();
  if (!s) return '';
  if (/^\d+$/.test(s)) return s;
  const compact = s.replace(/\s/g, '');
  const m = compact.match(/^(\d+)[.,]0+$/);
  if (m) return m[1];
  // Agar ID aralash formatda bo'lsa ham tushirib yubormaymiz.
  // 1) uzun raqam ketma-ketligi bo'lsa uni olamiz
  const seq = compact.match(/(\d{4,})/);
  if (seq) return seq[1];
  // 2) umuman raqam bo'lmasa ham original qiymatni qaytaramiz
  //    (public.view_merchants dagi qator yo'qolib ketmasligi uchun)
  return s;
};
const normText = (v) => String(v || '').trim().toLowerCase();
const cleanDistrictDisplay = (v) => {
  const raw = String(v || '').trim();
  if (!raw) return '';
  const parts = raw.split(';').map((x) => String(x || '').trim()).filter(Boolean);
  const filtered = parts.filter((p) => !isCommonDistrictLabel(p));
  return filtered[0] || '';
};
const normalizeMatchText = (v) =>
  String(v || '')
    .trim()
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[^a-z0-9а-я\s.$-]/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
const normalizeWarehouseKey = (v) => normalizeMatchText(v).replace(/\s+/g, ' ').trim();
const isMainWarehouseLabel = (v) => {
  const s = normalizeWarehouseKey(v).replace(/\s+/g, '');
  if (!s) return false;
  return (
    s.includes('основнойсклад') ||
    s.includes('основной') ||
    s.includes('главныйсклад') ||
    s.includes('mainwarehouse') ||
    s.includes('osnovnoysklad') ||
    s.includes('osnovnoy') ||
    s.includes('asosnoysklad') ||
    s.includes('asosiysklad')
  );
};
const textHasAny = (text, needles = []) =>
  needles.some((n) => text === n || text.includes(n));
const pickFirstBy = (candidates = [], predicate) => {
  for (const v of candidates) {
    if (predicate(v)) return String(v || '').trim();
  }
  return '';
};
const isOrderDoc = (v) => {
  const t = normalizeMatchText(v);
  if (!t) return false;
  if (textHasAny(t, ['заказ', 'zakaz'])) return true;
  return t.includes('зак') || t.includes('zak');
};
const isReturnDoc = (v) => {
  const t = normalizeMatchText(v);
  if (!t) return false;
  if (textHasAny(t, ['возврат', 'vozvrat'])) return true;
  return t.includes('возв') || t.includes('vozvr');
};
const isPaymentFromCounterparty = (v) => {
  const t = normalizeMatchText(v);
  if (!t) return false;
  if (textHasAny(t, [
    'оплата от контрагента',
    'oplata ot kontragenta',
    'payment from counterparty',
  ])) return true;
  return (
    (t.includes('оплата') && t.includes('контрагент') && (t.includes(' от ') || t.endsWith(' от') || t.startsWith('от '))) ||
    (t.includes('oplata') && t.includes('kontragent') && (t.includes(' ot ') || t.includes(' from ')))
  );
};
const isPaymentToCounterparty = (v) => {
  const t = normalizeMatchText(v);
  if (!t) return false;
  if (textHasAny(t, [
    'оплата контрагенту',
    'oplata kontragentu',
    'payment to counterparty',
  ])) return true;
  return (
    (t.includes('оплата') && t.includes('контрагент') && !isPaymentFromCounterparty(v)) ||
    (t.includes('oplata') && t.includes('kontragent') && t.includes(' to '))
  );
};
const isLikelyDocType = (v) => {
  const s = String(v || '').trim();
  if (!s) return false;
  return isOrderDoc(s) || isReturnDoc(s);
};
const isLikelyWarehouseName = (v) => {
  const t = normalizeMatchText(v);
  if (!t) return false;
  return (
    t.includes('склад') ||
    t.includes('sklad') ||
    t.includes('warehouse') ||
    t.includes('ombor')
  );
};
const isLikelyStatus = (v) => {
  const st = normalizeMatchText(v);
  if (!st) return false;
  return textHasAny(st, [
    'доставлен',
    'получен на склад',
    'отменено',
    'новый',
    'в пути',
    'new',
    'cancelled',
    'in transit',
  ]);
};
function parseObzvonAllRows(rows = []) {
  if (!Array.isArray(rows) || rows.length < 2) return [];
  const headerIdx = rows.findIndex((r) => {
    const c0 = String((r || [])[0] || '').trim().toLowerCase();
    const c1 = String((r || [])[1] || '').trim().toLowerCase();
    return c0 === 'Р Р†РІР‚С›РІР‚вЂњ' || c0 === 'Р В Р вЂ Р Р†Р вЂљРЎвЂєР Р†Р вЂљРІР‚Сљ' || c0 === 'no' || c1.includes('Р В РЎвЂќР В РЎвЂўР В Р вЂ¦Р РЋРІР‚С™Р РЋР вЂљР В Р’В°Р В РЎвЂ“Р В Р’ВµР В Р вЂ¦Р РЋРІР‚С™') || c1.includes('Р РЋР вЂљР РЋРІР‚СњР РЋР вЂљР РЋРІР‚СћР РЋР вЂљР РЋРІР‚СћР РЋР С“Р Р†Р вЂљРЎв„ўР РЋР С“Р РЋРІР‚в„ўР РЋР вЂљР вЂ™Р’В°Р РЋР вЂљР РЋРІР‚вЂњР РЋР вЂљР вЂ™Р’ВµР РЋР вЂљР РЋРІР‚СћР РЋР С“Р Р†Р вЂљРЎв„ў') || c1.includes('kontragent');
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

const ACCESS_SYNC_TOPIC = 'AQ_ACCESS';
const ACCESS_SYNC_CUSTOMER = '__AQCFG__';
const ACCESS_SYNC_OPERATOR = 'AQ_SYNC';

function encodeAccessConfig(payload) {
  try {
    const raw = JSON.stringify(payload || {});
    return btoa(unescape(encodeURIComponent(raw)));
  } catch {
    return '';
  }
}
function decodeAccessConfig(encoded) {
  try {
    if (!encoded) return null;
    const json = decodeURIComponent(escape(atob(String(encoded))));
    const obj = JSON.parse(json);
    return obj && typeof obj === 'object' ? obj : null;
  } catch {
    return null;
  }
}
function isAccessSyncRow(r) {
  if (!r) return false;
  const topic = String(r.topic || '').trim();
  const customer = String(r.customer || '').trim();
  const op = String(r.operator || '').trim();
  return topic === ACCESS_SYNC_TOPIC || customer === ACCESS_SYNC_CUSTOMER || op === ACCESS_SYNC_OPERATOR;
}
function extractAccessConfigFromRows(rows = []) {
  if (!Array.isArray(rows) || !rows.length) return null;
  const candidates = rows.filter((r) => isAccessSyncRow(r));
  if (!candidates.length) return null;
  const latest = candidates[candidates.length - 1];
  const encoded = String(latest.note || latest.callDate || latest.customerId || '').trim();
  const decoded = decodeAccessConfig(encoded);
  if (!decoded || typeof decoded !== 'object') return null;
  if (!Array.isArray(decoded.users) || !decoded.users.length) return null;
  return decoded;
}

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў STATUS HELPERS Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
// Р В Р’В Р Р†Р вЂљРЎСљР В Р’В Р РЋРІР‚С”Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™ va Р В Р’В Р РЋРЎСџР В Р’В Р РЋРІР‚С”Р В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р В РІвЂљВ¬Р В Р’В Р вЂ™Р’В§Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™ Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚в„ў Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРЎСљ - ikkisi ham yetkazilgan hisoblanadi
const isDeliveredStatus = (s) => {
  const st = normalizeMatchText(s);
  if (!st) return false;
  return textHasAny(st, [
    'доставлен',
    'получен на склад',
    'dostavlen',
    'poluchen na sklad',
  ]);
};
// Bekor qilingan
const isCancelledStatus = (s) => {
  const st = normalizeMatchText(s);
  return textHasAny(st, ['отменено', 'otmeneno', 'cancelled']);
};

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў MAHSULOT ANIQLASH Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
const normProduct = (p) => normalizeMatchText(p).replace(/\s+/g, ' ');
const WATER_PRODUCT_NAMES = new Set([
  'murodbaxsh 18.9l',
  'бонус murodbaxsh 18.9l',
  'bonus murodbaxsh 18.9l',
]);
// Tara ga ta'sir qiladigan mahsulotlar (zakaz=+, vozvrat=-)
const isTaraAffectingProduct = (p) => {
  const pl = normProduct(p);
  if (isWaterProduct(pl)) return true;
  if (!pl.includes('18.9l')) return false;
  return (
    pl.includes('tara') ||
    pl.includes('тара') ||
    pl.includes('пустой') ||
    pl.includes('pustoy') ||
    pl.includes('empty')
  );
};
// Faqat suv mahsulotlari (daromad hisoblash uchun)
const isWaterProduct = (p) => {
  const pl = normProduct(p);
  if (WATER_PRODUCT_NAMES.has(pl)) return true;
  // Nom variatsiyalari ko'p bo'lgani uchun kengroq tekshiruv.
  if (pl.includes('murodbaxsh') && pl.includes('18.9')) return true;
  if (pl.includes('муродбахш') && pl.includes('18.9')) return true;
  return false;
};

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў EXCEL READER Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
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

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў GOOGLE SHEETS Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
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
      { key:'public.view_merchants', label:'Mijozlar', sheet:'public.view_merchants', required:true },
      { key:'public.view_current_merchant_event_balance', label:'Balans', sheet:'public.view_current_merchant_event_balance', required:false },
      { key:'public.view_item_basket', label:'Zakazlar', sheet:'public.view_item_basket', required:false },
      { key:'public.view_cashbox_documents', label:'Kassa', sheet:'public.view_cashbox_documents', required:false },
      { key:'public.view_warehouse_transfer', label:'Permesheniya', sheet:'public.view_warehouse_transfer', required:false },
      { key:'intigratsiya', label:'Integratsiya', sheet:'intigratsiya', required:false },
      { key:'mijozlar', label:'Mijoz biriktiruv', sheet:'mijozlar', required:false },
    ];
    const sheets = {};
    for (const s of byName) {
      onProgress(`${s.label} yuklanmoqda...`);
      try {
        sheets[s.key] = await fetchSheetCsvByName(sheetId, s.sheet, s.label);
      } catch (e1) {
        try {
          sheets[s.key] = await fetchSheetOpenSheet(sheetId, s.sheet, s.label);
        } catch (e2) {
          if (s.required) throw e2;
          sheets[s.key] = [];
        }
      }
    }
    if (!sheets['public.view_merchants']) throw new Error('Mijozlar varaqi topilmadi!');
    return { sheets };
  }

  const sheetNames = [
    { key: 'public.view_merchants',           gid: gids.merchants,   label: 'Mijozlar' },
    { key: 'Balans',                           gid: gids.balans,      label: 'Balans' },
    { key: 'public.view_item_basket',          gid: gids.orders,      label: 'Zakazlar' },
    { key: 'public.view_cashbox_documents',    gid: gids.cashbox,     label: 'Kassa' },
    { key: 'public.view_warehouse_transfer',   gid: gids.warehouseTransfer, label: 'Permesheniya' },
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

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў DATA PROCESSING Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
function processAll(mainData) {
  if (!mainData) return { customers:[], orders:[], cashbox:[], contacts:[] };

  /* Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™ 1. MERCHANTS Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™
     A(r[0])=ID  C(r[2])=Nom  D(r[3])=Tel  E(r[4])=Kontakt
     O(r[14])=Manzil  V(r[21])=Rayon  Y(r[24])=EskiID  Z(r[25])=Manba */
  const merchantSheet = mainData.sheets['public.view_merchants'];
  const contacts = [];
  const allMerchants = [];
  if (merchantSheet) {
    merchantSheet.slice(1).forEach((r) => {
      const id = normId(r[0]);
      if (!id) return;
      const name = String(r[2] || '').trim();
      const source = String(r[25] || '').trim(); // Z ustun
      const merchNote = String(r[19] || '').trim(); // T ustun = primechaniya
      const aaValue = String(r[26] || '').trim(); // AA ustun
      allMerchants.push({ id, name, source, merchNote, aaTag: aaValue });
      contacts.push({
        id,
        name,
        phone:    String(r[3]  || '').trim().replace(/[^+\d]/g,'').slice(0,13),
        contact:  String(r[4]  || '').trim(),
        address:  String(r[14] || '').trim(),
        district: pickPreferredDistrict(r[21]),
        source,
        aaTag: aaValue,
        merchantNote: merchNote,
        eskiId:   String(r[24] || '').trim(), // Y ustun - eski ID (integratsiya uchun)
      });
    });
  }
  const contactById = {};
  contacts.forEach((c) => { contactById[c.id] = c; });

  /* Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™ 2. BALANS Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™
     8-qatordan boshlanadi (slice(7))
     G(r[6])=MijozID  D(r[3])=BalansUZS  C(r[2])=BalansUSD */
  const balMapUZS = {};
  const balMapUSD = {};
  const balMapUZSByName = {};
  const balMapUSDByName = {};
  const balSheet =
    mainData.sheets['public.view_current_merchant_event_balance'] ||
    mainData.sheets['Balans'];
  if (balSheet) {
    balSheet.forEach((r) => {
      const mid = normId(r[6]) || normId(r[1]); // G ustun (fallback: B ustun)
      const nameKey = normText(r[1] || r[10] || '');
      const uzs = toNum(r[3]);
      const usd = toNum(r[2]);
      if (mid) {
        balMapUZS[mid] = (balMapUZS[mid] || 0) + uzs; // D ustun = UZS balans (SUMIFS kabi yig'iladi)
        balMapUSD[mid] = (balMapUSD[mid] || 0) + usd; // C ustun = USD balans
      }
      if (nameKey && !/^\d+$/.test(nameKey)) {
        balMapUZSByName[nameKey] = (balMapUZSByName[nameKey] || 0) + uzs;
        balMapUSDByName[nameKey] = (balMapUSDByName[nameKey] || 0) + usd;
      }
    });
  }

  /* Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™ 3. INTIGRATSIYA Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™
     2-qatordan boshlanadi (slice(1))
     A(r[0])=EskiID  H(r[7])=TaraOstatka
     Logika: merchants.eskiId === intigratsiya.eskiId Р В Р вЂ Р Р†Р вЂљР’В Р Р†Р вЂљРІвЂћСћ boshlang'ich tara */
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

  /* Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™ 4. ZAKAZLAR Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™
     A(r[0])=ZakazNom  D(r[3])=Mahsulot  H(r[7])=Miqdor  I(r[8])=Summa
     J(r[9])=Kat  K(r[10])=KontragentNom  L(r[11])=Valyuta
     M(r[12])=HujjatTuri  N(r[13])=Status  R(r[17])=UniqueID
     V(r[21])=Agent  W(r[22])=Yetib borgan sana  X(r[23])=DostavchikIsmi  AD(r[29])=MijozID */
  const orderSheet = mainData.sheets['public.view_item_basket'];
  const rawOrders = [];
  if (orderSheet) {
    orderSheet.slice(1).forEach((r) => {
      const soNum = pickFirstBy([r[0], r[1]], (v) => String(v || '').trim().length > 0);
      const product = pickFirstBy([r[3], r[4], r[5]], (v) => String(v || '').trim().length > 0);
      const qty = toNum(pickFirstBy([r[7], r[8], r[6]], (v) => String(v ?? '').trim() !== ''));
      const sum = toNum(pickFirstBy([r[8], r[9], r[7]], (v) => String(v ?? '').trim() !== ''));
      const cat = pickFirstBy([r[9], r[10]], (v) => String(v || '').trim().length > 0);
      const contName = pickFirstBy([r[10], r[11], r[2]], (v) => String(v || '').trim().length > 0);
      const currency = normCurrency(pickFirstBy([r[11], r[12], r[10]], (v) => String(v || '').trim().length > 0));

      // Ustun siljigan fayllar uchun himoya: docType/status/date ni bir nechta ustundan qidiramiz.
      const detectedDocType = pickFirstBy([r[12], r[13], r[11], r[14]], (v) => isLikelyDocType(v));
      let docType = detectedDocType || String(r[12] || '').trim();
      if (!isLikelyDocType(docType)) {
        docType = qty < 0 || sum < 0 ? 'Возврат' : 'Заказ';
      }

      const detectedStatus = pickFirstBy([r[13], r[14], r[12], r[15]], (v) => isLikelyStatus(v));
      let status = detectedStatus || String(r[13] || '').trim();

      const uniqueId = String(r[17] || r[18] || '').trim();
      const agent = pickFirstBy([r[21], r[20], r[18]], (v) => String(v || '').trim().length > 0);
      const delivPerson = pickFirstBy([r[23], r[24], r[22]], (v) => String(v || '').trim().length > 0); // X
      const orderDate = pickFirstBy([r[22], r[24], r[23], r[21], r[20]], (v) => !!toDate(v)); // W
      const warehouse = pickFirstBy([r[19], r[20], r[18], r[16], r[25]], (v) => isLikelyWarehouseName(v));
      const mId = normId(pickFirstBy([r[29], r[30], r[28], r[27]], (v) => String(v || '').trim().length > 0));
      const note = pickFirstBy(
        [r[14], r[15], r[16], r[26], r[27], r[31], r[32], r[33], r[34]],
        (v) => {
          const s = String(v || '').trim();
          if (!s) return false;
          if (isLikelyDocType(s) || isLikelyStatus(s) || isLikelyWarehouseName(s) || toDate(s)) return false;
          return !/^\d+$/.test(s);
        }
      );

      if (!isLikelyStatus(status) && orderDate) {
        const od = toDate(orderDate);
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        if (od && od <= today) status = 'ДОСТАВЛЕН';
      }

      const price = qty && qty !== 0 ? Math.abs(sum / qty) : 0;

      if (!soNum || soNum === 'id') return;
      rawOrders.push({
        soNum, product, qty, sum, cat, contName, currency,
        docType, status, uniqueId, agent,
        delivPerson, // dostavchik ismi
        orderDate,   // zakaz/vozvrat sanasi
        mId, price, warehouse, note,
      });
    });
  }

  /* Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™ 5. KASSA Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™
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
      const mId      = normId(r[23]);
      const note     = String(r[13] || '').trim();

      if (!opNum || opNum === 'Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚СћР В Р’В Р РЋР’ВР В Р’В Р вЂ™Р’ВµР В Р Р‹Р В РІР‚С™ Р В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚вЂќР В Р’В Р вЂ™Р’ВµР В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р Р‹Р Р†Р вЂљР’В Р В Р’В Р РЋРІР‚ВР В Р’В Р РЋРІР‚В') return;
      rawCash.push({ opStatus, opNum, sana, amount, currency, opType, operator, contName, kassa, mId, note });
    });
  }

  /* -- 5.1. PERMESHENIYA --
     R(r[17])=OperationID  S(r[18])=SoSklada  T(r[19])=NaSklad
     U(r[20])=Sana         V(r[21])=Mahsulot  W(r[22])=Soni
     X(r[23])=Kod          Y(r[24])=UniqueID */
  const transferSheet = mainData.sheets['public.view_warehouse_transfer'];
  const warehouseTransfers = [];
  if (transferSheet) {
    transferSheet.slice(1).forEach((r) => {
      const movementId = String(r[17] || '').trim();
      const fromWarehouse = String(r[18] || '').trim();
      const toWarehouse = String(r[19] || '').trim();
      const moveDate = String(r[20] || '').trim();
      const product = String(r[21] || '').trim();
      const qty = Math.abs(toNum(r[22]));
      const changeCode = String(r[23] || '').trim();
      const uniqueId = String(r[24] || '').trim();
      if (!movementId && !fromWarehouse && !toWarehouse && !moveDate && !product && !qty && !uniqueId) return;
      warehouseTransfers.push({
        movementId,
        fromWarehouse,
        toWarehouse,
        moveDate,
        product,
        qty,
        changeCode,
        uniqueId,
      });
    });
  }

  /* Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™ 6. GURUHLASH Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™ */
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
      const id = normId(r[1]);
      if (!id) return;
      let operator = '';
      if (code === 'Op3' || code === '5') operator = 'Dildora';
      if (code === 'Op2' || code === '4') operator = 'Dilfuza';
      if (operator) assignmentById[id] = operator;
    });
  }

  /* Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™ 7. MIJOZLAR YARATISH Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™Р В Р вЂ Р Р†Р вЂљРЎСљР В РІР‚С™ */
  const customers = contacts.map((c) => {
    const myOrders = ordersByMId[c.id] || [];
    const myCash   = cashByMId[c.id]   || [];

    // Suv yetkazilgan (daromad, sotish statistikasi uchun)
    const waterDelivered = myOrders.filter(
      (o) => isWaterProduct(o.product) && isOrderDoc(o.docType) && isDeliveredStatus(o.status)
    );

    const deliveredTaraRows = myOrders.filter(
      (o) => isTaraAffectingProduct(o.product) && isDeliveredStatus(o.status)
    );
    // Boshlang'ich tara: merchants.eskiId Р В Р вЂ Р Р†Р вЂљР’В Р Р†Р вЂљРІвЂћСћ intigratsiya.A Р В Р вЂ Р Р†Р вЂљР’В Р Р†Р вЂљРІвЂћСћ intigratsiya.H
    const integTara  = c.eskiId ? (integTaraByEskiId[c.eskiId] || 0) : 0;
    const taraDelta = deliveredTaraRows.reduce((s, o) => {
      const q = Math.abs(toNum(o.qty));
      if (!q) return s;
      if (isReturnDoc(o.docType)) return s - q;
      if (isOrderDoc(o.docType)) return s + (toNum(o.qty) < 0 ? -q : q);
      return s;
    }, 0);
    const taraMinusQ = deliveredTaraRows
      .filter((o) => isReturnDoc(o.docType))
      .reduce((s, o) => s + Math.abs(toNum(o.qty)), 0);
    const tara       = integTara + taraDelta;

    const totalWaterQ     = waterDelivered.reduce((s, o) => s + Math.abs(o.qty), 0);
    const totalWaterS_uzs = waterDelivered.filter((o) => o.currency !== 'USD').reduce((s, o) => s + o.sum, 0);
    const totalWaterS_usd = waterDelivered.filter((o) => o.currency === 'USD').reduce((s, o) => s + o.sum, 0);

    const kulerOrds = myOrders.filter(
      (o) => (o.product || '').toLowerCase().includes('kuler') ||
              (o.product || '').toLowerCase().includes('Р В Р’В Р РЋРІР‚СњР В Р Р‹Р РЋРІР‚СљР В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’ВµР В Р Р‹Р В РІР‚С™')
    );
    const kulers = kulerOrds.filter((o) => isOrderDoc(o.docType)).length;

    const sortedWater = [...waterDelivered].sort((a, b) => {
      const da = toDate(a.orderDate), db = toDate(b.orderDate);
      return da && db ? db - da : 0;
    });
    const last = sortedWater[0];
    const days = last ? daysAgo(last.orderDate) : null;

    const nameKey = normText(c.name);
    const balanceUZS = (balMapUZS[c.id] ?? balMapUZSByName[nameKey] ?? 0);
    const balanceUSD = (balMapUSD[c.id] ?? balMapUSDByName[nameKey] ?? 0);

    const totalPaidUZS = myCash
      .filter((p) => isPaymentFromCounterparty(p.opType) && p.currency !== 'USD')
      .reduce((s, p) => s + p.amount, 0);
    const totalPaidUSD = myCash
      .filter((p) => isPaymentFromCounterparty(p.opType) && p.currency === 'USD')
      .reduce((s, p) => s + p.amount, 0);

    return {
      id: c.id,
      name: c.name,
      phone: c.phone || '-',
      district: c.district,
      source: c.source,
      aaTag: c.aaTag || '',
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
    return pl.includes('kuler') || pl.includes('Р В Р’В Р РЋРІР‚СњР В Р Р‹Р РЋРІР‚СљР В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’ВµР В Р Р‹Р В РІР‚С™');
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
      orderNo: firstOrder?.soNum || '-',
      operator: firstOrder?.agent || '-',
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
  const debtorsByBalance = customers.filter((c) => !isExcludedZCategory(c.source) && (c.balanceUZS ?? 0) < 0);
  const otherDebtorsByBalance = customers.filter((c) => isExcludedZCategory(c.source) && (c.balanceUZS ?? 0) < 0);

  return {
    customers, orders: rawOrders, cashbox: rawCash, contacts,
    rawOrders, rawCash, ordersByMId, cashByMId, warehouseTransfers, kulerInstallments, assignmentById, debtorsByBalance, otherDebtorsByBalance,
  };
}

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў SVERKA Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
function buildSverka(customer, ordersByMId, cashByMId) {
  const cid = customer.id;
  const myOrders = (ordersByMId[cid] || []).map((o) => ({
    ...o,
    _type: 'order',
    _date: toDate(o.orderDate) || new Date(0),
    _dateStr: o.orderDate,
  }));
  const myCash = (cashByMId[cid] || [])
    .filter((p) => isPaymentFromCounterparty(p.opType) || isPaymentToCounterparty(p.opType))
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
      // Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРЎв„ўР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚С” bo'lmagan va Р В Р’В Р Р†Р вЂљРЎСљР В Р’В Р РЋРІР‚С”Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™/Р В Р’В Р РЋРЎСџР В Р’В Р РЋРІР‚С”Р В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р В РІвЂљВ¬Р В Р’В Р вЂ™Р’В§Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™ Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚в„ў Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРЎСљ bo'lgan Р В Р вЂ Р Р†Р вЂљР’В Р Р†Р вЂљРІвЂћСћ balansga ta'sir qiladi
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
        dokument: row.docType || 'Р В Р’В -Р В Р’В Р вЂ™Р’В°Р В Р’В Р РЋРІР‚СњР В Р’В Р вЂ™Р’В°Р В Р’В  | ',
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
        driver: row.delivPerson, // X ustun - dostavchik ismi
        status: row.status,
        _isVoz: isVoz,
        _isTara: isTara,
        _isWater: isWater,
        _isCancelled: isCancelled,
        _type: 'order',
      };
    } else {
      const signedAmount = isPaymentToCounterparty(row.opType) ? -Math.abs(row.amount || 0) : Math.abs(row.amount || 0);
      if (row.currency === 'USD') runBalUSD += signedAmount;
      else runBalUZS += signedAmount;
      return {
        kod: row.opNum,
        sana: fmtD(row._dateStr),
        dokument: row.opType || "Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРІР‚вЂќР В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’В°Р В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В° Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎв„ў Р В Р’В Р РЋРІР‚СњР В Р’В Р РЋРІР‚СћР В Р’В Р В РІР‚В¦Р В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р’В Р РЋРІР‚вЂњР В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦Р В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В°",
        produkt: '',
        qty: null,
        narx: null,
        summa: 0,
        tolov: signedAmount,
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

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў EXCEL EXPORT Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
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
  const detailRows = rows.map((r) => [
    r.kod, r.sana, r.dokument, r.produkt,
    r.qty != null ? r.qty : '',
    r.narx ? Math.round(r.narx) : '',
    r.summa || '',
    r._type === 'payment' && r.currency !== 'USD' ? r.tolov : '',
    r._type === 'payment' && r.currency === 'USD'  ? r.tolov : '',
    r.balansUZS, r.balansUSD, r.currency,
    r.agent, r.driver, r.status,
  ]);
  const ws = XLSX.utils.aoa_to_sheet(title, { cellDates:true, raw:true });
  const { typedRows, colTypes } = buildTypedAoa(headers, detailRows);
  XLSX.utils.sheet_add_aoa(ws, [headers, ...typedRows], { origin: `A${title.length + 1}`, cellDates:true, raw:true });
  colTypes.forEach((type, colIndex) => {
    for (let rowIndex = title.length + 2; rowIndex <= title.length + typedRows.length + 1; rowIndex++) {
      const addr = XLSX.utils.encode_cell({ r: rowIndex - 1, c: colIndex });
      const cell = ws[addr];
      if (!cell) continue;
      if (type === 'number' && cell.t === 'n') cell.z = EXCEL_NUM_FMT;
      if (type === 'date' && (cell.t === 'd' || cell.t === 'n')) cell.z = EXCEL_DATE_FMT;
    }
  });
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
  const ws = buildTypedExcelSheet(
    headers,
    customers.map((c) => [
      c.id, c.name, c.phone, c.district, c.balanceUZS, c.balanceUSD,
      c.tara, c.kulers, c.totalWaterQ, c.totalWaterS_uzs, c.totalWaterS_usd,
      fmtD(c.lastOrderDate), c.daysAgo != null ? c.daysAgo : '', c.lastAgent, c.source,
    ])
  );
  ws['!cols'] = [10,35,14,16,14,12,10,7,12,16,12,14,8,12,14].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws, 'Mijozlar');

  const debtors = customers.filter((c) => c.balanceUZS < 0 || c.balanceUSD < 0)
    .sort((a,b) => a.balanceUZS - b.balanceUZS);
  const dh = ['ID','Mijoz','Telefon','Rayon','Qarz UZS','Qarz USD','Idish','Oxirgi zakaz','Kun','Agent'];
  const ws2 = buildTypedExcelSheet(
    dh,
    debtors.map((c) => [
      c.id, c.name, c.phone, c.district, c.balanceUZS, c.balanceUSD,
      c.tara, fmtD(c.lastOrderDate), c.daysAgo != null ? c.daysAgo : '', c.lastAgent,
    ])
  );
  ws2['!cols'] = [10,35,14,16,14,12,8,12,6,12].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws2, 'Qarzdorlar');
  XLSX.writeFile(wb, `AquaBiz_Hisobot_${new Date().toISOString().slice(0,10)}.xlsx`);
}
const safeExcelName = (v) => String(v || 'Report').replace(/[/\\:*?"<>|]/g, '_').slice(0, 120);
const buildColsFromAoa = (headers, rows) => {
  const sample = (rows || []).slice(0, 300);
  return headers.map((h, i) => {
    let w = String(h || '').length + 2;
    sample.forEach((r) => {
      const len = String((r || [])[i] ?? '').length + 2;
      if (len > w) w = len;
    });
    return { wch: Math.min(48, Math.max(8, w)) };
  });
};
const EXCEL_NUM_FMT = '#,##0';
const EXCEL_DATE_FMT = 'dd.mm.yyyy';
const isDateLikeHeader = (h) => /(sana|date|oxirgi|keyingi|olingan|muddat)/i.test(String(h || '').trim());
const isNumberLikeHeader = (h) => /(summa|balans|qarz|dona|soni|kol-vo|qty|oylar|qoldiq|to['`Р Р†Р вЂљРІвЂћСћ]?langan|idish|jami|miqdor|suv soni|plan|kunlik|itog|operatorga|count|(^|\s)(no|Р Р†РІР‚С›РІР‚вЂњ)(\s|$))/i.test(String(h || '').trim());
const isTextForcedHeader = (h) => /(^|\s)(id|kod|telefon|phone|operator|agent|kontragent|mijoz|rayon|status|mavzu|izoh|mahsulot|hujjat|tur|oy)(\s|$)/i.test(String(h || '').trim());
const parseExcelNumber = (v) => {
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  const s0 = String(v ?? '').trim();
  if (!s0 || s0 === '-' || s0 === '-') return null;
  let s = s0
    .replace(/\s+/g, '')
    .replace(/[^\d,.\-]/g, '');
  if (!s || s === '-' || s === '.' || s === ',') return null;
  const hasDot = s.includes('.');
  const hasComma = s.includes(',');
  if (hasDot && hasComma) {
    if (s.lastIndexOf('.') > s.lastIndexOf(',')) s = s.replace(/,/g, '');
    else s = s.replace(/\./g, '').replace(',', '.');
  } else if (hasComma && !hasDot) {
    const parts = s.split(',');
    if (parts.length > 2) s = parts.join('');
    else {
      const decLen = (parts[1] || '').length;
      s = decLen === 3 ? parts.join('') : `${parts[0]}.${parts[1] || ''}`;
    }
  } else if (hasDot && !hasComma) {
    const parts = s.split('.');
    if (parts.length > 2) s = parts.join('');
    else {
      const decLen = (parts[1] || '').length;
      if (decLen === 3 && /^\d+$/.test(parts[0] || '') && /^\d+$/.test(parts[1] || '')) s = parts.join('');
    }
  }
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
};
const isLikelyDateInputForExcel = (v) => {
  if (v instanceof Date) return true;
  if (typeof v === 'number') {
    return Number.isFinite(v) && v >= 40000 && v <= 60000;
  }
  const s = String(v ?? '').trim();
  if (!s) return false;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return true;
  if (/^\d{2}[./-]\d{2}[./-]\d{4}/.test(s)) return true;
  if (/^\d{4}[./-]\d{2}[./-]\d{2}/.test(s)) return true;
  return false;
};
const normalizeDateForExcel = (v) => {
  if (!isLikelyDateInputForExcel(v)) return null;
  const d = toDate(v);
  if (!d) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
};
const detectExcelColumnType = (header, rows, colIndex) => {
  if (isTextForcedHeader(header)) return 'text';
  if (isDateLikeHeader(header)) return 'date';
  if (isNumberLikeHeader(header)) return 'number';
  const sample = (rows || []).slice(0, 300).map((r) => (r || [])[colIndex]).filter((v) => String(v ?? '').trim() !== '');
  if (!sample.length) return 'text';
  let numHits = 0;
  let dateHits = 0;
  sample.forEach((v) => {
    if (parseExcelNumber(v) != null) numHits++;
    if (normalizeDateForExcel(v)) dateHits++;
  });
  if (dateHits / sample.length >= 0.7) return 'date';
  if (numHits / sample.length >= 0.7) return 'number';
  return 'text';
};
const normalizeExcelColumnType = (v) => {
  const t = String(v || '').trim().toLowerCase();
  if (t === 'date' || t === 'number' || t === 'text') return t;
  return '';
};
const buildTypedAoa = (headers = [], rows = [], forcedTypes = []) => {
  const colTypes = headers.map((h, idx) => normalizeExcelColumnType(forcedTypes[idx]) || detectExcelColumnType(h, rows, idx));
  const typedRows = (rows || []).map((r) => headers.map((_, colIndex) => {
    const type = colTypes[colIndex];
    const raw = (r || [])[colIndex];
    if (type === 'date') return normalizeDateForExcel(raw) || (raw ?? '');
    if (type === 'number') {
      const n = parseExcelNumber(raw);
      return n == null ? (raw ?? '') : n;
    }
    return raw ?? '';
  }));
  return { typedRows, colTypes };
};
const buildTypedExcelSheet = (headers = [], rows = [], forcedTypes = []) => {
  const { typedRows, colTypes } = buildTypedAoa(headers, rows, forcedTypes);
  const ws = XLSX.utils.aoa_to_sheet([headers, ...typedRows], { cellDates:true, raw:true });
  ws['!cols'] = buildColsFromAoa(headers, typedRows);
  headers.forEach((_, colIndex) => {
    const type = colTypes[colIndex];
    for (let rowIndex = 2; rowIndex <= typedRows.length + 1; rowIndex++) {
      const addr = XLSX.utils.encode_cell({ r: rowIndex - 1, c: colIndex });
      const cell = ws[addr];
      if (!cell) continue;
      if (type === 'number' && cell.t === 'n') cell.z = EXCEL_NUM_FMT;
      if (type === 'date' && (cell.t === 'd' || cell.t === 'n')) cell.z = EXCEL_DATE_FMT;
      if (type === 'text') {
        cell.t = 's';
        cell.v = cell.v == null ? '' : String(cell.v);
        cell.z = '@';
      }
    }
  });
  return ws;
};
function exportAoaExcel({ fileName, sheetName = 'Hisobot', headers = [], rows = [], columnTypes = [] }) {
  const wb = XLSX.utils.book_new();
  const ws = buildTypedExcelSheet(headers, rows, columnTypes);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, safeExcelName(fileName));
}

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў CSS Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
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
  padding:7px 13px;transition:transform .12s ease, box-shadow .18s ease, background .18s ease, color .18s ease;
  display:inline-flex;align-items:center;gap:6px;white-space:nowrap;position:relative;overflow:hidden}
.btn::after{content:'';position:absolute;inset:0;pointer-events:none;background:radial-gradient(circle at 20% 20%, rgba(255,255,255,.22), transparent 55%);opacity:0;transition:opacity .2s ease}
.btn:hover::after,.btn:focus-visible::after{opacity:1}
.btn:active{transform:translateY(1px) scale(.985)}
.btn:focus-visible{outline:2px solid var(--bl);outline-offset:2px}
.btn-bl{background:var(--bl);color:#000;box-shadow:0 4px 14px rgba(88,166,255,.18)}
.btn-bl:hover{filter:brightness(1.03)}
.btn-gr{background:var(--gr);color:#000;box-shadow:0 4px 14px rgba(63,185,80,.16)}
.btn-gr:hover{filter:brightness(1.03)}
.btn-gh{background:transparent;color:var(--t2);border:1px solid var(--b1)}
.btn-gh:hover{background:var(--s3);color:var(--t1)}
.btn-sm{padding:5px 10px;font-size:11.5px}
.input{width:100%;padding:8px 11px;border:1px solid var(--b1);border-radius:var(--r);
  font-size:13px;color:var(--t1);background:var(--s3);background:linear-gradient(180deg,var(--s3),color-mix(in oklab, var(--s3) 82%, var(--bg)));
  outline:none;transition:border-color .16s ease, box-shadow .16s ease, background .16s ease}
.input:focus{border-color:var(--bl);box-shadow:0 0 0 3px rgba(88,166,255,.12)}
.select{padding:7px 10px;border:1px solid var(--b1);border-radius:var(--r);
  font-size:12.5px;color:var(--t1);background:var(--s3);background:linear-gradient(180deg,var(--s3),color-mix(in oklab, var(--s3) 82%, var(--bg)));cursor:pointer;outline:none;transition:border-color .16s ease, box-shadow .16s ease}
.select:focus{border-color:var(--bl);box-shadow:0 0 0 3px rgba(88,166,255,.12)}
.select, select{color-scheme:dark}
.select option, select option{background:var(--s1);color:var(--t1)}
.modern-date-wrap{position:relative;width:100%}
.modern-date-btn{display:flex;align-items:center;justify-content:space-between;gap:8px;text-align:left}
.modern-date-label{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.modern-date-label.ph{color:var(--t3)}
.modern-date-icon{font-size:14px;opacity:.85}
.modern-date-pop{position:absolute;top:calc(100% + 6px);left:0;z-index:120;min-width:280px;padding:10px;box-shadow:0 16px 40px rgba(0,0,0,.58)}
.modern-date-head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:8px}
.modern-date-month{font-size:12.5px;font-weight:700;text-transform:capitalize;color:var(--t1)}
.modern-date-week{display:grid;grid-template-columns:repeat(7,1fr);gap:4px;margin-bottom:4px}
.modern-date-week span{font-size:10px;color:var(--t3);text-align:center}
.modern-date-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:4px}
.modern-date-day{border:1px solid var(--b1);border-radius:8px;background:var(--s3);min-height:30px;font-size:12px;color:var(--t2);cursor:pointer;transition:all .14s}
.modern-date-day:hover{border-color:var(--bl);color:var(--t1)}
.modern-date-day.out{opacity:.5}
.modern-date-day.on{background:var(--bl2);border-color:var(--bl);color:var(--bl);font-weight:700}
.modern-date-day.today:not(.on){border-style:dashed}
.modern-date-day:disabled{opacity:.35;cursor:not-allowed}
.modern-date-actions{display:flex;justify-content:space-between;gap:8px;margin-top:8px}
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
.tab:focus-visible{outline:2px solid var(--bl);outline-offset:1px}
.tab.on{background:var(--s1);color:var(--t1);box-shadow:0 1px 4px rgba(0,0,0,.4)}
.tab:active{transform:translateY(1px)}
.nav-i{display:flex;align-items:center;gap:9px;padding:8px 10px;border-radius:var(--r);cursor:pointer;color:var(--t2);font-size:13px;font-weight:500;transition:all .13s;user-select:none;position:relative}
.nav-i:hover{background:var(--s2);color:var(--t1)}
.nav-i.on{background:var(--bl2);color:var(--bl)}
.nav-i:active{transform:translateY(1px)}
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
.staff-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:10px}
.staff-user-card{background:var(--s1);border:1px solid var(--b2);border-radius:12px;padding:12px;display:grid;gap:8px;cursor:pointer;transition:border-color .16s ease, transform .12s ease, box-shadow .16s ease}
.staff-user-card:hover{border-color:var(--bl);box-shadow:0 8px 20px rgba(88,166,255,.12)}
.staff-user-card:active{transform:translateY(1px)}
.staff-user-card.on{border-color:var(--bl);box-shadow:0 10px 24px rgba(88,166,255,.16)}
.perm-section{border:1px solid var(--b2);border-radius:10px;background:var(--s2);overflow:hidden}
.perm-head{width:100%;display:flex;align-items:center;justify-content:space-between;gap:8px;padding:9px 10px;background:transparent;border:none;color:inherit;cursor:pointer}
.perm-title{display:flex;align-items:center;gap:8px;font-size:12.5px;font-weight:700;color:var(--t1)}
.perm-body{padding:8px 10px 10px;border-top:1px solid var(--b2);display:grid;gap:6px;background:var(--s1)}
.perm-row{display:flex;align-items:center;justify-content:space-between;gap:8px;background:var(--s3);border:1px solid var(--b2);border-radius:8px;padding:7px 8px}
.perm-row span{font-size:11.5px;color:var(--t2)}
.hol-week{display:grid;grid-template-columns:repeat(7,1fr);gap:6px}
.hol-week span{font-size:10.5px;color:var(--t3);text-align:center}
.hol-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:6px}
.hol-day{border:1px solid var(--b1);border-radius:8px;background:var(--s3);min-height:34px;font-size:12px;color:var(--t2);cursor:pointer;transition:all .14s}
.hol-day:hover{border-color:var(--bl);color:var(--t1)}
.hol-day.on{background:var(--bl2);color:var(--bl);border-color:var(--bl);font-weight:700}
.hol-day.out{opacity:.55}
.hol-day.weekly:not(.on){border-style:dashed}
.hol-day:disabled{cursor:not-allowed;opacity:.6}
`;
const LIGHT_THEME_VARS = {
  '--bg': '#f4f7fb',
  '--s1': '#ffffff',
  '--s2': '#edf2f8',
  '--s3': '#e6ecf5',
  '--b1': '#c9d3e1',
  '--b2': '#d9e1ec',
  '--t1': '#0f1b2d',
  '--t2': '#33445d',
  '--t3': '#5f7088',
  '--t4': '#8a96a9',
  '--bl': '#1f6feb',
  '--bl2': '#dbeafe',
  '--bl3': '#e8f1ff',
  '--gr2': '#dff7e6',
  '--gr3': '#effcf3',
  '--rd2': '#ffe4e4',
  '--yl2': '#fff2d9',
  '--or2': '#ffe8d8',
  '--pu2': '#efe6ff',
};

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў LOCAL STORAGE Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
const S = {
  get: (k,d) => { try { const v=localStorage.getItem(k); return v?JSON.parse(v):d; } catch { return d; } },
  set: (k,v) => { try { localStorage.setItem(k,JSON.stringify(v)); } catch {} },
};
const DEFAULT_USERS = ['Dildora', 'Dilfuza', 'Admin'];
const DEFAULT_USER_CREDS = { Dildora:'Dildora', Dilfuza:'Dilfuza', Admin:'12345' };
const DEFAULT_CUSTOMER_TABS = {
  all: true,
  ahmadtea: true,
  murodbaxsh: true,
  activeAll: true,
  activeOwn: true,
  inactive: true,
  other: true,
};
const DEFAULT_ACCESS = {
  Dildora: {
    scope: 'own',
    activeScope: 'own',
    role: 'operator',
    customerTabs: { ...DEFAULT_CUSTOMER_TABS },
    visible: { dash:true, cust:true, orders:true, kassa:true, obzvon:true, doljniki:true, nazorat:true, reports:true, plan:true, refresh:true, settings:false, settings_staff:false, settings_app:false, settings_ui:false, obzvon_new_edit:false, obzvon_new_delete:false, obzvon_new_publish:false },
    ui: { theme:'dark' },
    company: { canSwitch:false, default:'murodbaxsh' },
  },
  Dilfuza: {
    scope: 'own',
    activeScope: 'own',
    role: 'operator',
    customerTabs: { ...DEFAULT_CUSTOMER_TABS },
    visible: { dash:true, cust:true, orders:true, kassa:true, obzvon:true, doljniki:true, nazorat:true, reports:true, plan:true, refresh:true, settings:false, settings_staff:false, settings_app:false, settings_ui:false, obzvon_new_edit:false, obzvon_new_delete:false, obzvon_new_publish:false },
    ui: { theme:'dark' },
    company: { canSwitch:false, default:'murodbaxsh' },
  },
  Admin:   {
    scope: 'all',
    activeScope: 'all',
    role: 'admin',
    customerTabs: { ...DEFAULT_CUSTOMER_TABS },
    visible: { dash:true, cust:true, orders:true, kassa:true, obzvon:true, doljniki:true, nazorat:true, reports:true, plan:true, refresh:true, settings:true, settings_staff:true, settings_app:true, settings_ui:true, obzvon_new_edit:true, obzvon_new_delete:true, obzvon_new_publish:true },
    ui: { theme:'dark' },
    company: { canSwitch:true, default:'murodbaxsh' },
  },
};
const DEFAULT_VISIBLE_PAGES = { ...DEFAULT_ACCESS.Admin.visible };
const DEFAULT_OWN_CUSTOMER_TABS = { ...DEFAULT_CUSTOMER_TABS };
function normalizeAccessConfig(user, cfg) {
  const isAdmin = user === 'Admin';
  const src = cfg || {};
  const base = isAdmin
    ? {
        scope:'all',
        activeScope:'all',
        role:'admin',
        customerTabs:{ ...DEFAULT_CUSTOMER_TABS },
        visible:{ ...DEFAULT_VISIBLE_PAGES, settings:true, settings_staff:true, settings_app:true, settings_ui:true, obzvon_new_edit:true, obzvon_new_delete:true, obzvon_new_publish:true },
        ui:{ theme:'dark' },
        company:{ canSwitch:true, default:'murodbaxsh' },
      }
    : {
        scope:'own',
        activeScope:'own',
        role:'operator',
        customerTabs:{ ...DEFAULT_OWN_CUSTOMER_TABS },
        visible:{ ...DEFAULT_VISIBLE_PAGES, settings:false, settings_staff:false, settings_app:false, settings_ui:false, obzvon_new_edit:false, obzvon_new_delete:false, obzvon_new_publish:false },
        ui:{ theme:'dark' },
        company:{ canSwitch:false, default:'murodbaxsh' },
      };
  const scope = src.scope || base.scope;
  const srcCompany = (src.company && typeof src.company === 'object') ? src.company : {};
  const normalized = {
    scope,
    activeScope: scope === 'own' ? 'own' : 'all',
    role: src.role || base.role,
    visible: { ...base.visible, ...(src.visible || {}) },
    customerTabs: { ...base.customerTabs, ...(src.customerTabs || {}) },
    ui: { ...base.ui, ...((src.ui && typeof src.ui === 'object') ? src.ui : {}) },
    company: {
      canSwitch: Boolean(srcCompany.canSwitch ?? base.company.canSwitch),
      default: normalizeCompanyKey(srcCompany.default || base.company.default),
    },
  };
  // Admin always keeps Settings access to avoid accidental lockout.
  if (isAdmin) {
    normalized.visible.settings = true;
    normalized.visible.settings_staff = true;
    normalized.visible.settings_app = true;
    normalized.visible.settings_ui = true;
    normalized.visible.obzvon_new_edit = true;
    normalized.visible.obzvon_new_delete = true;
    normalized.visible.obzvon_new_publish = true;
  }
  return normalized;
}

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў UPLOAD MODAL Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
function UploadModal({
  onLoad,
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
    merchants:'', balans:'', orders:'', cashbox:'', integration:'', mijozlar:'', warehouseTransfer:'',
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
    const usedSheetId = extractSheetId(usedMainUrl);
    if (!usedSheetId) { alert("Google Sheets URL kodda sozlanmagan!"); return; }
    setLoad(true);
    try {
      if (setMainSheetUrl) setMainSheetUrl(usedMainUrl);
      S.set('aq-main-url', usedMainUrl);
      const raw = await loadFromGoogleSheets(usedSheetId, gids, setProg, 'named');
      setProg('Hisoblanmoqda...');
      setTimeout(() => {
        try {
          const data = processAll(raw);
          onLoad(data);
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
            <div style={{fontWeight:800,fontSize:16}}>{E.upload} Ma'lumot yuklash</div>
            <div style={{fontSize:12,color:'var(--t3)',marginTop:3}}>Excel yoki Google Sheets dan</div>
          </div>
          {hasData && onClose && <button className="btn btn-gh btn-sm" onClick={onClose}>X</button>}
        </div>
        <div className="mbdy">
          <div className="tabs" style={{marginBottom:18}}>
            {[['excel',`${E.excel} Excel fayl`],['sheets',`${E.sheets} Google Sheets`]].map(([t,l]) => (
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
                  <div style={{fontSize:28,marginBottom:6}}>{files.main?E.ok:E.excel}</div>
                  {files.main
                    ? <div style={{fontWeight:700,color:'var(--gr)',fontSize:13}}>{files.main.name}</div>
                    : <div>
                        <div style={{fontWeight:700,fontSize:13,marginBottom:4}}>Asosiy Excel faylni tanlang</div>
                        <div style={{fontSize:11.5,color:'var(--t3)'}}>public.view_merchants, Balans, Zakazlar, Kassa, public.view_warehouse_transfer varaqlari</div>
                      </div>
                  }
                </div>
              </div>
              <div style={{marginBottom:14}}>
                <div style={{fontSize:11.5,fontWeight:700,color:'var(--t3)',textTransform:'uppercase',letterSpacing:'.04em',marginBottom:8}}>
                  Integratsiya fayli (ixtiyoriy - eski baza ostatkasi)
                </div>
                <div className={`drop-z${files.integration?' done':''}`}
                  onClick={()=>!loading&&pickFile('integration')} style={{padding:'16px',textAlign:'center'}}>
                  <div style={{fontSize:22,marginBottom:4}}>{files.integration?E.ok:E.sheets}</div>
                  {files.integration
                    ? <div style={{fontWeight:700,color:'var(--gr)',fontSize:13}}>{files.integration.name}</div>
                    : <div>
                        <div style={{fontWeight:700,fontSize:13,marginBottom:3}}>Integratsiya faylini tanlang</div>
                        <div style={{fontSize:11,color:'var(--t3)'}}>A=eski ID   |   H=tara ostatkasi   |   Varaq nomi: "intigratsiya"</div>
                      </div>
                  }
                </div>
              </div>
              {loading && (
                <div style={{background:'var(--bl3)',border:'1px solid var(--bl2)',borderRadius:8,padding:'9px 14px',marginBottom:12,display:'flex',alignItems:'center',gap:10,color:'var(--bl)'}}>
                  <span style={{display:'inline-block',animation:'spin .7s linear infinite'}}>O</span>
                  <span style={{fontSize:13,fontWeight:600}}>{progress}</span>
                </div>
              )}
              <div style={{display:'flex',justifyContent:'flex-end',gap:8}}>
                {hasData && onClose && <button className="btn btn-gh" onClick={onClose}>Bekor</button>}
                <button className="btn btn-bl" onClick={doLoadExcel} disabled={loading||!files.main} style={{minWidth:140}}>
                  {loading?'O Yuklanmoqda...':`${E.ok} Yuklash`}
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
                {sheetId && <div style={{fontSize:11,color:'var(--gr)',marginTop:5,fontFamily:'var(--mono)'}}>{E.ok} ID: {sheetId}</div>}
              </div>
              <div style={{background:'var(--s3)',border:'1px solid var(--b1)',borderRadius:8,padding:'10px 12px',marginBottom:12,fontSize:12,color:'var(--t2)'}}>
                Kerakli listlar:
                <div style={{marginTop:6,fontFamily:'var(--mono)',fontSize:11}}>
                  public.view_merchants, public.view_current_merchant_event_balance, public.view_item_basket, public.view_cashbox_documents, public.view_warehouse_transfer, intigratsiya, mijozlar
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
                  "Р В РЎвЂєР В Р’В±Р В Р’В·Р В Р вЂ Р В РЎвЂўР В Р вЂ¦ Р В РІР‚в„ўР В Р Р‹Р В РІР‚Сћ" listi shu faylda bo'lishi kerak
                </div>
              </div>
              <div style={{background:'var(--yl2)',border:'1px solid var(--yl)',borderRadius:8,padding:'9px 14px',marginBottom:14,fontSize:12,color:'var(--yl)'}}>
                Eslatma: Fayl <strong>Fayl &gt; Veb-da nashr qilish</strong> qilingan bo'lishi kerak
              </div>
              {loading && (
                <div style={{background:'var(--bl3)',border:'1px solid var(--bl2)',borderRadius:8,padding:'9px 14px',marginBottom:12,display:'flex',alignItems:'center',gap:10,color:'var(--bl)'}}>
                  <span style={{display:'inline-block',animation:'spin .7s linear infinite'}}>O</span>
                  <span style={{fontSize:13,fontWeight:600}}>{progress}</span>
                </div>
              )}
              <div style={{display:'flex',justifyContent:'flex-end',gap:8}}>
                {hasData && onClose && <button className="btn btn-gh" onClick={onClose}>Bekor</button>}
                <button className="btn btn-bl" onClick={doLoadSheets} disabled={loading||!sheetId} style={{minWidth:160}}>
                  {loading?'O Ulanmoqda...':`${E.sheets} Ulash`}
                </button>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў CUSTOMER DETAIL Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
function CustomerDetail({ c, D, onClose }) {
  const { ordersByMId, cashByMId } = D;
  const [activeTab, setTab] = useState('sverka');
  const sverkaWrapRef = useRef(null);

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
    (p) => isPaymentFromCounterparty(p.opType) || isPaymentToCounterparty(p.opType)
  );

  const balColor = (v) => v < 0 ? 'var(--rd)' : v > 0 ? 'var(--gr)' : 'var(--t3)';
  const taraColor = c.tara < 0 ? 'var(--rd)' : 'var(--bl)';
  useEffect(() => {
    if (activeTab !== 'sverka') return;
    const el = sverkaWrapRef.current;
    if (!el) return;
    el.scrollTop = el.scrollHeight;
  }, [activeTab, sverkaRows.length, c.id]);

  return (
    <div className="modal-ov fade" onClick={(e)=>e.target===e.currentTarget&&onClose()}>
      <div className="modal ani" style={{maxWidth:1000}}>
        <div className="mhdr">
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontWeight:700,fontSize:14,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
              {c.name}
            </div>
            <div style={{fontSize:11,color:'var(--t3)',fontFamily:'var(--mono)',marginTop:2}}>
              ID: {c.id}   |   {c.district}   |   {c.phone}
            </div>
          </div>
          <div style={{display:'flex',gap:8,flexShrink:0,marginLeft:12}}>
            <button className="btn btn-gr btn-sm" onClick={()=>exportSverkaExcel(c,sverkaRows)}>Excel</button>
            <button className="btn btn-gh btn-sm" onClick={onClose}>X</button>
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
            {[['sverka',`${E.sv} Sverka`],['orders',`${E.order} Zakazlar`],['returns',`${E.ret} Vozvratlar`],['pays',`${E.pay} To'lovlar`]].map(([t,l]) => (
              <button key={t} className={`tab${activeTab===t?' on':''}`} onClick={()=>setTab(t)}>{l}</button>
            ))}
          </div>

          {activeTab==='sverka' && (
            <div>
              <div style={{background:'var(--s3)',border:'1px solid var(--b1)',borderRadius:8,padding:'8px 14px',marginBottom:12,fontSize:12,color:'var(--t3)',display:'flex',gap:16,flexWrap:'wrap'}}>
                <span>{E.order} Yetkazilgan: <strong style={{color:'var(--bl)'}}>{c.totalWaterQ} ta</strong></span>
                <span>{E.ret} Qaytarilgan: <strong style={{color:'var(--or)'}}>{c.vozvratQ} ta</strong></span>
                <span>{E.tara} Mijozda: <strong style={{color:c.tara<0?'var(--rd)':'var(--gr)'}}>{c.tara} ta</strong></span>
                <span>{E.uzs} Balans UZS: <strong style={{color:balColor(finalBalUZS)}}>{fmt(Math.abs(finalBalUZS))} so'm</strong></span>
                <span>{E.usd} Balans USD: <strong style={{color:balColor(finalBalUSD)}}>{fmt(Math.abs(finalBalUSD))} $</strong></span>
              </div>
              <div ref={sverkaWrapRef} style={{overflow:'auto',maxHeight:'52vh',borderRadius:9,border:'1px solid var(--b2)'}}>
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
                              {row.produkt||'-'}
                            </span>
                          </td>
                          <td style={{textAlign:'center',fontWeight:700,color:row.qty!=null?(row._isVoz?'var(--or)':row.qty>0?'var(--t1)':'var(--t3)'):'var(--t3)'}}>
                            {row.qty!=null?(row._isVoz?'-':'')+Math.abs(row.qty):'-'}
                          </td>
                          <td style={{textAlign:'right',fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>
                            {row.narx?fmt(Math.round(row.narx)):'-'}
                          </td>
                          <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700,color:row.summa?'var(--rd)':'var(--t4)'}}>
                            {row.summa?'-'+fmt(row.summa):'-'}
                          </td>
                          <td style={{textAlign:'center'}}>
                            <span className={row.currency==='USD'?'cur-usd':'cur-uzs'}>{row.currency||'UZS'}</span>
                          </td>
                          <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700,color:row.tolov>0?'var(--gr)':row.tolov<0?'var(--rd)':'var(--t4)'}}>
                            {row.tolov ? `${row.tolov > 0 ? '+' : '-'}${fmt(Math.abs(row.tolov))}` : '-'}
                          </td>
                          <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:800,color:row.balansUZS<0?'var(--rd)':row.balansUZS>0?'var(--gr)':'var(--t3)'}}>
                            {row.balansUZS<0?'-':row.balansUZS>0?'+':''}{fmt(Math.abs(row.balansUZS))}
                          </td>
                          <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700,color:row.balansUSD<0?'var(--rd)':row.balansUSD>0?'var(--gr)':'var(--t3)',fontSize:11}}>
                            {row.balansUSD<0?'-':row.balansUSD>0?'+':''}{fmt(Math.abs(row.balansUSD))}$
                          </td>
                          <td style={{fontSize:11,color:'var(--t3)'}}>{row.agent||'-'}</td>
                          <td style={{fontSize:11,color:'var(--t2)'}}>{row.driver||'-'}</td>
                          <td>
                            <span className="tag" style={{
                              background: row.status==='Р В Р’В Р Р†Р вЂљРЎСљР В Р’В Р РЋРІР‚С”Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™'?'var(--gr2)':row.status==='Р В Р’В Р РЋРЎСџР В Р’В Р РЋРІР‚С”Р В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р В РІвЂљВ¬Р В Р’В Р вЂ™Р’В§Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™ Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚в„ў Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРЎСљ'?'var(--gr2)':row.status==='Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРЎв„ўР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚С”'?'var(--rd2)':'var(--s3)',
                              color:      row.status==='Р В Р’В Р Р†Р вЂљРЎСљР В Р’В Р РЋРІР‚С”Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™'?'var(--gr)':row.status==='Р В Р’В Р РЋРЎСџР В Р’В Р РЋРІР‚С”Р В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р В РІвЂљВ¬Р В Р’В Р вЂ™Р’В§Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™ Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚в„ў Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРЎСљ'?'var(--gr)':row.status==='Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРЎв„ўР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚С”'?'var(--t4)':'var(--t3)',
                              fontSize:9.5,
                            }}>
                              {row.status||'-'}
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
                  Jami {sverkaRows.length} ta yozuv   |  {' '}
                  {sverkaRows.filter((r)=>r._type==='order'&&!r._isCancelled).length} ta zakaz   |  {' '}
                  {sverkaRows.filter((r)=>r._type==='payment').length} ta to'lov
                  {sverkaRows.some((r)=>r._isCancelled) && (
                    <span style={{color:'var(--rd)',marginLeft:8}}>
                        |   {sverkaRows.filter((r)=>r._isCancelled).length} ta otmena
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
                    <th>Zakaz No</th><th>Sana</th><th>Mahsulot</th>
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
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{o.price?fmt(Math.round(o.price)):'-'}</td>
                        <td style={{fontFamily:'var(--mono)',fontWeight:700,color:'var(--gr)'}}>{o.sum?fmt(o.sum):'-'}</td>
                        <td><span className={o.currency==='USD'?'cur-usd':'cur-uzs'}>{o.currency||'UZS'}</span></td>
                        <td style={{fontSize:11}}>{o.agent||'-'}</td>
                        <td style={{fontSize:11,color:'var(--t2)'}}>{o.delivPerson||'-'}</td>
                        <td>
                          <span className="tag" style={{
                            background:o.status==='Р В Р’В Р Р†Р вЂљРЎСљР В Р’В Р РЋРІР‚С”Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™'||o.status==='Р В Р’В Р РЋРЎСџР В Р’В Р РЋРІР‚С”Р В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р В РІвЂљВ¬Р В Р’В Р вЂ™Р’В§Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™ Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚в„ў Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРЎСљ'?'var(--gr2)':o.status==='Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРЎв„ўР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚С”'?'var(--rd2)':'var(--s3)',
                            color:o.status==='Р В Р’В Р Р†Р вЂљРЎСљР В Р’В Р РЋРІР‚С”Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™'||o.status==='Р В Р’В Р РЋРЎСџР В Р’В Р РЋРІР‚С”Р В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р В РІвЂљВ¬Р В Р’В Р вЂ™Р’В§Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™ Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚в„ў Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРЎСљ'?'var(--gr)':o.status==='Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРЎв„ўР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚С”'?'var(--t4)':'var(--t3)',
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
                    <th>Vozvrat No</th><th>Sana</th><th>Mahsulot</th>
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
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{o.price?fmt(Math.round(o.price)):'-'}</td>
                        <td style={{fontFamily:'var(--mono)',fontWeight:700,color:'var(--or)'}}>{o.sum?fmt(Math.abs(o.sum)):'-'}</td>
                        <td><span className={o.currency==='USD'?'cur-usd':'cur-uzs'}>{o.currency||'UZS'}</span></td>
                        <td style={{fontSize:11}}>{o.agent||'-'}</td>
                        <td style={{fontSize:11,color:'var(--t2)'}}>{o.delivPerson||'-'}</td>
                        <td>
                          <span className="tag" style={{
                            background:o.status==='Р В Р’В Р Р†Р вЂљРЎСљР В Р’В Р РЋРІР‚С”Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™'||o.status==='Р В Р’В Р РЋРЎСџР В Р’В Р РЋРІР‚С”Р В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р В РІвЂљВ¬Р В Р’В Р вЂ™Р’В§Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™ Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚в„ў Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРЎСљ'?'var(--gr2)':o.status==='Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРЎв„ўР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚С”'?'var(--rd2)':'var(--s3)',
                            color:o.status==='Р В Р’В Р Р†Р вЂљРЎСљР В Р’В Р РЋРІР‚С”Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™'||o.status==='Р В Р’В Р РЋРЎСџР В Р’В Р РЋРІР‚С”Р В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р В РІвЂљВ¬Р В Р’В Р вЂ™Р’В§Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™ Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚в„ў Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІвЂћСћР В Р’В Р Р†Р вЂљРЎвЂќР В Р’В Р РЋРІР‚в„ўР В Р’В Р Р†Р вЂљРЎСљ'?'var(--gr)':o.status==='Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРЎв„ўР В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р Р†Р вЂљРЎС›Р В Р’В Р РЋРЎС™Р В Р’В Р РЋРІР‚С”'?'var(--t4)':'var(--t3)',
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
                  <tr><th>Op No</th><th>Sana</th><th>Summa</th><th>Valyuta</th><th>Kassa</th><th>Operator</th><th>Izoh</th></tr>
                </thead>
                <tbody>
                  {myPays.length===0
                    ? <tr><td colSpan={7} style={{textAlign:'center',padding:32,color:'var(--t3)'}}>To'lovlar yo'q</td></tr>
                    : myPays.map((p,i) => {
                      const signed = isPaymentToCounterparty(p.opType) ? -Math.abs(p.amount || 0) : Math.abs(p.amount || 0);
                      return (
                      <tr key={i}>
                        <td style={{fontFamily:'var(--mono)',fontSize:10.5,color:'var(--t3)'}}>{p.opNum}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(p.sana)}</td>
                        <td style={{fontFamily:'var(--mono)',fontWeight:800,color:signed>=0?'var(--gr)':'var(--rd)'}}>
                          {signed>=0?'+':'-'}{fmt(Math.abs(signed))}
                        </td>
                        <td><span className={p.currency==='USD'?'cur-usd':'cur-uzs'}>{p.currency||'UZS'}</span></td>
                        <td style={{fontSize:11,color:'var(--t3)'}}>{p.kassa||'-'}</td>
                        <td style={{fontSize:11}}>{p.operator||'-'}</td>
                        <td style={{fontSize:11,color:'var(--t3)'}}>{p.note||'-'}</td>
                      </tr>
                    )})
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

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў STAT CARD Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
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

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў DASHBOARD Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
function Dashboard({ D }) {
  const { customers, cashbox, rawOrders, debtorsByBalance=[] } = D;
  const now = new Date();
  const curMonthKey = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const wDel = (rawOrders||[]).filter((o) => isWaterProduct(o.product) && isOrderDoc(o.docType) && isDeliveredStatus(o.status));
  const wDelThisMonth = wDel.filter((o) => monthKey(o.orderDate)===curMonthKey);
  const pays = (cashbox||[]).filter((c) => isPaymentFromCounterparty(c.opType));
  const debtors = debtorsByBalance.length ? debtorsByBalance : customers.filter((c) => c.balanceUZS < 0);
  const stale2m = customers
    .filter((c) => c.hasOrders && c.daysAgo != null && c.daysAgo >= 60 && isActiveCustomerName(c.name))
    .sort((a,b) => (b.daysAgo || 0) - (a.daysAgo || 0));

  const monthOptions = useMemo(
    () => [...new Set(wDel.map((o) => monthKey(o.orderDate)).filter(Boolean))].sort(),
    [wDel]
  );
  const [agentMonth, setAgentMonth] = useState(curMonthKey);
  const [rayonMonth, setRayonMonth] = useState(curMonthKey);
  useEffect(() => {
    if (!monthOptions.length) return;
    if (agentMonth !== 'all' && !monthOptions.includes(agentMonth)) {
      setAgentMonth(monthOptions.includes(curMonthKey) ? curMonthKey : monthOptions[monthOptions.length - 1]);
    }
    if (rayonMonth !== 'all' && !monthOptions.includes(rayonMonth)) {
      setRayonMonth(monthOptions.includes(curMonthKey) ? curMonthKey : monthOptions[monthOptions.length - 1]);
    }
  }, [monthOptions, agentMonth, rayonMonth, curMonthKey]);

  const monthlyPayments = useMemo(() => {
    const m = {};
    pays.forEach((c) => {
      const k = monthKey(c.sana); if (!k) return;
      if (!m[k]) m[k]={key:k,sum:0,cnt:0};
      m[k].sum+=c.amount; m[k].cnt++;
    });
    return Object.values(m).sort((a,b)=>a.key.localeCompare(b.key)).slice(-9);
  }, [pays]);
  const byMonth = useMemo(() => {
    const m = {};
    wDel.forEach((o) => {
      const k = monthKey(o.orderDate); if (!k) return;
      if (!m[k]) m[k]={key:k,qty:0,sumUZS:0};
      m[k].qty+=Math.abs(o.qty);
      if (o.currency !== 'USD') m[k].sumUZS += o.sum;
    });
    return Object.values(m).sort((a,b)=>a.key.localeCompare(b.key)).slice(-12);
  }, [wDel]);
  const byDist = useMemo(() => {
    const source = rayonMonth === 'all' ? wDel : wDel.filter((o) => monthKey(o.orderDate) === rayonMonth);
    const m = {};
    source.forEach((o) => {
      const c = customers.find((x)=>x.id===o.mId);
      const d = c?.district || "Noma'lum";
      if (!m[d]) m[d] = { name:d, qty:0, sum:0, custs:new Set() };
      m[d].qty += Math.abs(o.qty || 0);
      m[d].sum += (o.currency==='USD' ? 0 : (o.sum || 0));
      if (o.mId) m[d].custs.add(o.mId);
    });
    return Object.values(m)
      .map((x)=>({ ...x, custs:x.custs.size }))
      .sort((a,b)=>b.sum-a.sum);
  }, [wDel, customers, rayonMonth]);
  const agents = useMemo(() => {
    const source = agentMonth === 'all'
      ? wDel
      : wDel.filter((o) => monthKey(o.orderDate) === agentMonth);
    const m = {};
    source.forEach((o) => {
      const a=o.agent||'-';
      if (!m[a]) m[a]={name:a,qty:0,sum:0};
      m[a].qty+=Math.abs(o.qty); m[a].sum+=o.sum;
    });
    return Object.values(m).sort((a,b)=>b.sum-a.sum).slice(0,7);
  }, [wDel, agentMonth]);

  const debtSum = debtors.reduce((s, c) => s + Math.abs(c.balanceUZS || 0), 0);
  const COLORS = ['#58a6ff','#3fb950','#f85149','#d29922','#bc8cff','#f0883e','#79c0ff','#56d364','#ffa657'];

  return (
    <div style={{display:'flex',flexDirection:'column',gap:14}} className="ani">
      <div className="g5">
        <StatCard l="JAMI MIJOZLAR" v={customers.length} s={customers.filter((c)=>c.hasOrders).length+' aktiv'} c="var(--bl)"/>
        <StatCard l="QARZ UZS" v={fmt(debtSum)+" so'm"} s={debtors.length+' ta qarzdor'} c="var(--rd)"/>
        <StatCard l="JAMI IDISH" v={customers.reduce((s,c)=>s+c.tara,0)+' ta'} s="barcha mijozlarda" c="var(--pu)"/>
        <StatCard l="2 OY OLMAGAN" v={stale2m.length+' ta'} s="aktiv mijozlar ichidan" c="var(--yl)"/>
        <StatCard
          l={`SUV (${String(now.getMonth()+1).padStart(2,'0')}.${now.getFullYear()})`}
          v={fmt(wDelThisMonth.reduce((s,o)=>s+Math.abs(o.qty),0))+' ta'}
          s={fmt(wDelThisMonth.reduce((s,o)=>s+o.sum,0))+" so'm"} c="var(--gr)"/>
      </div>

      <div className="g2">
        <div className="card" style={{padding:16}}>
          <div style={{fontWeight:700,fontSize:13,marginBottom:12}}>{E.pay} Oylik to'lovlar (so'm)</div>
          {monthlyPayments.length===0
            ? <div style={{color:'var(--t3)',textAlign:'center',padding:24,fontSize:13}}>To'lov ma'lumoti yo'q</div>
            : <ResponsiveContainer width="100%" height={200}>
                <AreaChart data={monthlyPayments}>
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
            <div style={{fontWeight:700,fontSize:13}}>{E.top} Agentlar reytingi</div>
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

      <div className="card" style={{padding:16}}>
        <div style={{fontWeight:700,fontSize:13,marginBottom:12}}>{E.order} Oylik yetkazish (UZS)</div>
        {byMonth.length===0 ? (
          <div style={{color:'var(--t3)',textAlign:'center',padding:24,fontSize:13}}>Yetkazish ma'lumoti yo'q</div>
        ) : (
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={byMonth}>
              <CartesianGrid strokeDasharray="3 3" stroke="#21262d"/>
              <XAxis dataKey="key" tick={{fontSize:10,fill:'#656d76'}}/>
              <YAxis yAxisId="q" tick={{fontSize:10,fill:'#656d76'}} orientation="left"/>
              <YAxis yAxisId="s" tick={{fontSize:10,fill:'#656d76'}} orientation="right" tickFormatter={fmtM}/>
              <Tooltip formatter={(v,n)=>n==='Summa UZS'?fmt(v)+" so'm":v+' ta'} contentStyle={{background:'#161b22',border:'1px solid #30363d',borderRadius:8}}/>
              <Bar yAxisId="q" dataKey="qty" fill="var(--bl)" radius={[4,4,0,0]} name="Dona"/>
              <Bar yAxisId="s" dataKey="sumUZS" fill="var(--gr)" radius={[4,4,0,0]} name="Summa UZS"/>
            </BarChart>
          </ResponsiveContainer>
        )}
      </div>

      <div className="g2">
        <div className="card" style={{padding:16}}>
          <div style={{display:'flex',justifyContent:'space-between',gap:8,alignItems:'center',marginBottom:14}}>
            <div style={{fontWeight:700,fontSize:13}}>{E.area} Rayonlar bo'yicha</div>
            <select className="select" value={rayonMonth} onChange={(e)=>setRayonMonth(e.target.value)} style={{padding:'4px 8px',fontSize:12}}>
              <option value="all">Hammasi</option>
              {monthOptions.slice().reverse().map((m)=><option key={m} value={m}>{m}</option>)}
            </select>
          </div>
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
          <div style={{fontWeight:700,fontSize:13,marginBottom:12}}>{E.late} 2+ oy suv olmaganlar</div>
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

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў MIJOZLAR Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
function Customers({ D, currentUser='Admin', currentAccess=null, assignmentById={}, company='murodbaxsh' }) {
  const { customers } = D;
  const [segment, setSegment] = useState('all');
  const [search,setS]   = useState('');
  const [sort,setSort]  = useState({ col:'name', dir:'asc' });
  const [det,setDet]    = useState(null);
  const [showAdv, setShowAdv] = useState(false);
  const [uFilterOpen, setUFilterOpen] = useState(false);
  const [uFilterState, setUFilterState] = useState({});
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
    () => new Set(Object.entries(assignmentById || {}).filter(([, op]) => op === currentUser).map(([id]) => id)),
    [assignmentById, currentUser]
  );
  const tabVisible = {
    ...DEFAULT_CUSTOMER_TABS,
    ...((currentAccess && currentAccess.customerTabs) || {}),
  };
  const allowCompanyTabs = company === 'all';
  const activeScope = (currentAccess && currentAccess.activeScope) || ((currentAccess && currentAccess.scope) === 'own' ? 'own' : 'all');
  const availableTabs = useMemo(() => {
    const items = [];
    if (tabVisible.all) items.push({ id:'all', label:`${E.all} Hamma mijozlar` });
    if (allowCompanyTabs && tabVisible.ahmadtea) items.push({ id:'aa_ahmadtea', label:'Ahmadtea' });
    if (allowCompanyTabs && tabVisible.murodbaxsh) items.push({ id:'aa_other', label:'Murodbaxsh' });
    if (tabVisible.activeAll) items.push({ id:'active_all', label:'Aktiv (hammasi)' });
    if (tabVisible.activeOwn) items.push({ id:'active_own', label:"Aktiv (o'zimniki)" });
    if (tabVisible.inactive) items.push({ id:'inactive', label:'Nofaol mijozlar' });
    if (tabVisible.other) items.push({ id:'other_customers', label:'Boshqa mijozlar' });
    return items.length ? items : [{ id:'all', label:`${E.all} Hamma mijozlar` }];
  }, [tabVisible, allowCompanyTabs]);
  useEffect(() => {
    if (!availableTabs.some((t) => t.id === segment)) setSegment(availableTabs[0].id);
  }, [availableTabs, segment]);

  const nonAhmadtea = useMemo(
    () => customers.filter((c) => !isAhmadteaCustomer(c)),
    [customers]
  );
  const nonAhmadteaNonZ = useMemo(
    () => nonAhmadtea.filter((c) => !isExcludedZCategory(c.source)),
    [nonAhmadtea]
  );
  const activeBase = useMemo(
    () => nonAhmadteaNonZ.filter((c) => !isNameInactiveByPrefix(c.name)),
    [nonAhmadteaNonZ]
  );
  const segmentCustomers = useMemo(() => {
    if (segment === 'aa_ahmadtea') {
      return customers.filter((c) => isAhmadteaCustomer(c));
    }
    if (segment === 'aa_other') {
      return customers.filter((c) => isMurodbaxshCustomer(c));
    }
    if (segment === 'active_all') {
      return activeBase;
    }
    if (segment === 'active_own') {
      if (activeScope === 'all') return activeBase;
      return activeBase.filter((c) => ownIds.has(c.id));
    }
    if (segment === 'inactive') {
      return nonAhmadteaNonZ.filter((c) => isNameInactiveByPrefix(c.name));
    }
    if (segment === 'other_customers') {
      return nonAhmadtea.filter((c) => isExcludedZCategory(c.source) && !isNameInactiveByPrefix(c.name));
    }
    // Hamma mijozlar: public.view_merchants dagi hamma mijoz.
    return customers;
  }, [customers, segment, activeBase, ownIds, activeScope, nonAhmadtea, nonAhmadteaNonZ]);

  const toRange = (v, from, to) => {
    if (from !== '' && Number(v) < Number(from)) return false;
    if (to !== '' && Number(v) > Number(to)) return false;
    return true;
  };

  const customerFilterColumns = useMemo(() => ([
    { key:'id', label:'ID', type:'text' },
    { key:'name', label:'Kontragent', type:'text' },
    { key:'phone', label:'Telefon', type:'text' },
    { key:'district', label:'Rayon', type:'text' },
    { key:'source', label:'Manba', type:'text' },
    { key:'balanceUZS', label:'Balans UZS', type:'number' },
    { key:'balanceUSD', label:'Balans USD', type:'number' },
    { key:'tara', label:'Idish', type:'number' },
    { key:'kulers', label:'Kuler', type:'number' },
    { key:'lastOrderDate', label:'Oxirgi zakaz', type:'date' },
    { key:'daysAgo', label:'Kun', type:'number' },
    { key:'lastQty', label:'Dona', type:'number' },
    { key:'lastAgent', label:'Agent', type:'text' },
  ]), []);
  useEffect(() => {
    setUFilterState((prev) => ensureUniversalFilterState(customerFilterColumns, prev));
  }, [customerFilterColumns]);

  const list = useMemo(() => {
    let r = segmentCustomers;
    if (search) {
      const q = search.toLowerCase();
      r = r.filter((c) => c.name.toLowerCase().includes(q)||c.phone.includes(q)||c.id.includes(q)||(c.district||'').toLowerCase().includes(q));
    }
    // Hamma mijozlar bo'limi har doim to'liq ro'yxat bo'ladi.
    if (segment !== 'all') {
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
    }
    r = applyUniversalFilters(r, customerFilterColumns, uFilterState);

    return [...r].sort((a,b) => {
      let av=a[sort.col]??0, bv=b[sort.col]??0;
      if (typeof av==='string') av=av.toLowerCase();
      if (typeof bv==='string') bv=bv.toLowerCase();
      return sort.dir==='asc'?(av>bv?1:-1):av<bv?1:-1;
    });
  }, [segmentCustomers,search,sort,adv,segment,customerFilterColumns,uFilterState]);

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
  const universalFilterCount = useMemo(() => countUniversalFilters(uFilterState), [uFilterState]);

  const tog = (col) => setSort((s) => s.col===col?{col,dir:s.dir==='asc'?'desc':'asc'}:{col,dir:'asc'});
  const SI = ({c:col}) => sort.col===col
    ? <span style={{marginLeft:3}}>{sort.dir==='asc'?'^':'v'}</span>
    : <span style={{marginLeft:3,opacity:.2}}>↕</span>;
  const balColor = (v) => v<0?'var(--rd)':v>0?'var(--gr)':'var(--t3)';
  const debt = getDebtStats(segmentCustomers);

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:12,height:'100%'}}>
      <div className="tabs" style={{display:'inline-flex'}}>
        {availableTabs.map((t) => (
          <button key={t.id} className={`tab${segment===t.id?' on':''}`} onClick={()=>setSegment(t.id)}>{t.label}</button>
        ))}
      </div>
      <div className="g4">
        <StatCard l="JAMI MIJOZLAR" v={segmentCustomers.length} s={segmentCustomers.filter((c)=>c.hasOrders).length+' aktiv'} c="var(--bl)"/>
        <StatCard l="QARZDORLAR" v={debt.uzsCount+' ta'} s={fmt(debt.uzsSum)+" so'm"} c="var(--rd)"/>
        <StatCard l="JAMI IDISH" v={segmentCustomers.reduce((s,c)=>s+c.tara,0)+' ta'} s="barcha mijozlarda" c="var(--pu)"/>
        <StatCard l="FILTRLANGAN" v={list.length+' ta'} c="var(--gr)"/>
      </div>
      <div style={{display:'flex',gap:8,flexWrap:'wrap',alignItems:'center',position:'relative'}}>
        <div className="sb" style={{flex:2,minWidth:200}}>
          <span style={{color:'var(--t3)'}}>{E.find}</span>
          <input placeholder="Ism, telefon, ID bo'yicha..." value={search} onChange={(e)=>setS(e.target.value)}/>
        </div>
        <button className="btn btn-gh btn-sm" onClick={()=>setUFilterOpen((v)=>!v)}>
          <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
            <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z"/>
          </svg>
          Filtr ({universalFilterCount})
        </button>
        <button className="btn btn-gr btn-sm" onClick={()=>exportAllReport(segmentCustomers)}>Excel hisobot</button>
        <UniversalFilterPanel
          open={uFilterOpen}
          title="Mijozlar filtri"
          columns={customerFilterColumns}
          rows={segmentCustomers}
          state={uFilterState}
          setState={setUFilterState}
          onClose={()=>setUFilterOpen(false)}
          width={620}
        />

        {showAdv && (
            <div className="card" style={{position:'absolute',top:40,left:0,right:'auto',zIndex:50,width:'min(520px, calc(100vw - 24px))',padding:14,boxShadow:'0 24px 60px rgba(0,0,0,.55)',backdropFilter:'blur(8px)',overflow:'hidden'}}>
            <div className="g2" style={{marginBottom:8}}>
              <div>
                <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Rayon</div>
                <div style={{maxHeight:90,overflow:'auto',border:'1px solid var(--b1)',borderRadius:8,padding:6}}>
                  {dists.map((d)=><label key={d} style={FILTER_CHECK_LABEL_STYLE}><input type="checkbox" checked={adv.districts.includes(d)} onChange={(e)=>setAdv((p)=>({...p,districts:e.target.checked?[...p.districts,d]:p.districts.filter((x)=>x!==d)}))}/><span style={FILTER_CHECK_TEXT_STYLE}>{d}</span></label>)}
                </div>
              </div>
              <div>
                <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Manba</div>
                <div style={{maxHeight:90,overflow:'auto',border:'1px solid var(--b1)',borderRadius:8,padding:6}}>
                  {sources.map((d)=><label key={d} style={FILTER_CHECK_LABEL_STYLE}><input type="checkbox" checked={adv.sources.includes(d)} onChange={(e)=>setAdv((p)=>({...p,sources:e.target.checked?[...p.sources,d]:p.sources.filter((x)=>x!==d)}))}/><span style={FILTER_CHECK_TEXT_STYLE}>{d}</span></label>)}
                </div>
              </div>
              <div>
                <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Agent</div>
                <div style={{maxHeight:90,overflow:'auto',border:'1px solid var(--b1)',borderRadius:8,padding:6}}>
                  {agents.map((d)=><label key={d} style={FILTER_CHECK_LABEL_STYLE}><input type="checkbox" checked={adv.agents.includes(d)} onChange={(e)=>setAdv((p)=>({...p,agents:e.target.checked?[...p.agents,d]:p.agents.filter((x)=>x!==d)}))}/><span style={FILTER_CHECK_TEXT_STYLE}>{d}</span></label>)}
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
                  <ModernDateInput value={adv.lastFrom} onChange={(e)=>setAdv((p)=>({...p,lastFrom:e.target.value}))} />
                  <ModernDateInput value={adv.lastTo} onChange={(e)=>setAdv((p)=>({...p,lastTo:e.target.value}))} />
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
                    <td style={{fontSize:12}}>{c.district||'-'}</td>
                    <td style={{fontFamily:'var(--mono)',fontSize:11.5,fontWeight:700,color:balColor(c.balanceUZS)}}>{c.balanceUZS<0?'-':c.balanceUZS>0?'+':''}{fmt(Math.abs(c.balanceUZS))}</td>
                    <td style={{fontFamily:'var(--mono)',fontSize:11,fontWeight:700,color:balColor(c.balanceUSD)}}>{c.balanceUSD!==0?<>{c.balanceUSD<0?'-':c.balanceUSD>0?'+':''}{fmt(Math.abs(c.balanceUSD))}$</>:'-'}</td>
                    <td style={{textAlign:'center',fontWeight:700,color:c.tara<0?'var(--rd)':'var(--bl)'}}>{c.tara!==0?(c.tara<0?'-':'')+Math.abs(c.tara):'-'}</td>
                    <td style={{textAlign:'center'}}>{c.kulers>0?<span className="tag" style={{background:'var(--yl2)',color:'var(--yl)'}}>{c.kulers}</span>:'-'}</td>
                    <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(c.lastOrderDate)}</td>
                    <td>{c.daysAgo!=null&&c.daysAgo>=0 ? <span className="tag" style={{background:c.daysAgo>30?'var(--rd2)':c.daysAgo>14?'var(--yl2)':c.daysAgo>7?'var(--or2)':'var(--s3)',color:c.daysAgo>30?'var(--rd)':c.daysAgo>14?'var(--yl)':c.daysAgo>7?'var(--or)':'var(--t3)'}}>{c.daysAgo}k</span> : '-'}</td>
                    <td style={{textAlign:'center'}}>{c.lastQty||'-'}</td>
                    <td style={{fontSize:12}}>{c.lastAgent||'-'}</td>
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

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў SO DETAIL MODAL Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
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
              {fmtD(soGroup.orderDate)}   |   {soGroup.agent||'-'}   |   {soGroup.items.length} ta mahsulot
              {soGroup.delivPerson && <span>   |   {E.driver} {soGroup.delivPerson}</span>}
            </div>
          </div>
          <button className="btn btn-gh btn-sm" onClick={onClose}>X</button>
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
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{item.price?fmt(Math.round(item.price)):'-'}</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700,color:'var(--gr)'}}>{item.sum?fmt(item.sum):'-'}</td>
                    <td style={{textAlign:'center'}}><span className={item.currency==='USD'?'cur-usd':'cur-uzs'}>{item.currency||'UZS'}</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{marginTop:12,display:'flex',justifyContent:'space-between',alignItems:'center'}}>
            <span className="tag" style={{
              background:isDeliveredStatus(soGroup.status)?'var(--gr2)':isCancelledStatus(soGroup.status)?'var(--rd2)':'var(--s3)',
              color:isDeliveredStatus(soGroup.status)?'var(--gr)':isCancelledStatus(soGroup.status)?'var(--t4)':'var(--t3)',
            }}>{soGroup.status||'-'}</span>
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

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў ZAKAZLAR Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
function Orders({ D }) {
  const { rawOrders=[], customers=[] } = D;
  const [search,setS]  = useState('');
  const [fType,setT]   = useState('zakaz');
  const [showFilter, setShowFilter] = useState(false);
  const [uFilterOpen, setUFilterOpen] = useState(false);
  const [uFilterState, setUFilterState] = useState({});
  const [fAgents, setAgents] = useState([]);
  const [fStatuses, setStatuses] = useState([]);
  const [fCurrencies, setCurrencies] = useState([]);
  const [fDateFrom, setDateFrom] = useState('');
  const [fDateTo, setDateTo] = useState('');
  const [fQtyFrom, setQtyFrom] = useState('');
  const [fQtyTo, setQtyTo] = useState('');
  const [fSumFrom, setSumFrom] = useState('');
  const [fSumTo, setSumTo] = useState('');
  const [selectedSO,setSelectedSO] = useState(null);
  const districtById = useMemo(() => {
    const m = {};
    (customers || []).forEach((c) => { m[String(c.id || '').trim()] = c.district || ''; });
    return m;
  }, [customers]);

  const soGroups = useMemo(() => {
    const groups = {};
    rawOrders.forEach((o) => {
      if (!o.soNum) return;
      if (!groups[o.soNum]) {
        groups[o.soNum] = {
          soNum: o.soNum, contName: o.contName, mId: o.mId,
          district: districtById[String(o.mId || '').trim()] || '',
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
  }, [rawOrders, districtById]);

  const allZakaz   = soGroups.filter((g)=>isOrderDoc(g.docType));
  const allVozvrat = soGroups.filter((g)=>isReturnDoc(g.docType));
  const baseByType = fType==='zakaz'?allZakaz:fType==='vozvrat'?allVozvrat:soGroups;
  const base = baseByType;
  const options = useMemo(() => ({
    agents: [...new Set(soGroups.map((g)=>g.agent).filter(Boolean))].sort(),
    statuses: [...new Set(soGroups.map((g)=>g.status).filter(Boolean))].sort(),
    currencies: [...new Set(soGroups.map((g)=>g.currency || 'UZS').filter(Boolean))].sort(),
  }), [soGroups]);
  const inRange = (v, from, to) => {
    const n = Number(v || 0);
    if (from !== '' && n < Number(from)) return false;
    if (to !== '' && n > Number(to)) return false;
    return true;
  };
  const toggleIn = (setter, val) => setter((prev) => prev.includes(val) ? prev.filter((x)=>x!==val) : [...prev, val]);
  const orderFilterColumns = useMemo(() => ([
    { key:'soNum', label:'Zakaz No', type:'text' },
    { key:'contName', label:'Kontragent', type:'text' },
    { key:'mId', label:'Mijoz ID', type:'text' },
    { key:'district', label:'Rayon', type:'text' },
    { key:'orderDate', label:'Sana', type:'date' },
    { key:'delivPerson', label:'Dostavchik', type:'text' },
    { key:'totalQty', label:'Dona', type:'number' },
    { key:'totalSumUZS', label:'Summa UZS', type:'number' },
    { key:'totalSumUSD', label:'Summa USD', type:'number' },
    { key:'docType', label:'Tur', type:'text' },
    { key:'status', label:'Status', type:'text' },
    { key:'agent', label:'Agent', type:'text' },
    { key:'currency', label:'Valyuta', type:'text' },
  ]), []);
  useEffect(() => {
    setUFilterState((prev) => ensureUniversalFilterState(orderFilterColumns, prev));
  }, [orderFilterColumns]);

  const list = useMemo(() => {
    let r = base;
    const q = search.toLowerCase();
    if (q) r=r.filter((g)=>(g.contName||'').toLowerCase().includes(q)||(g.soNum||'').toLowerCase().includes(q));
    if (fAgents.length) r = r.filter((g)=>fAgents.includes(g.agent || ''));
    if (fStatuses.length) r = r.filter((g)=>fStatuses.includes(g.status || ''));
    if (fCurrencies.length) r = r.filter((g)=>fCurrencies.includes(g.currency || 'UZS'));
    if (fDateFrom) {
      const df = toDate(fDateFrom);
      if (df) r = r.filter((g) => {
        const d = toDate(g.orderDate);
        return d ? d >= df : false;
      });
    }
    if (fDateTo) {
      const dt = toDate(fDateTo);
      if (dt) r = r.filter((g) => {
        const d = toDate(g.orderDate);
        return d ? d <= dt : false;
      });
    }
    r = r.filter((g)=>inRange(g.totalQty, fQtyFrom, fQtyTo));
    r = r.filter((g)=>inRange(g.totalSumUZS, fSumFrom, fSumTo));
    r = applyUniversalFilters(r, orderFilterColumns, uFilterState);
    return [...r].sort((a,b) => {
      const da=toDate(a.orderDate), db=toDate(b.orderDate);
      return da&&db?db-da:0;
    });
  }, [base,search,fAgents,fStatuses,fCurrencies,fDateFrom,fDateTo,fQtyFrom,fQtyTo,fSumFrom,fSumTo,orderFilterColumns,uFilterState]);

  const activeFilterCount = useMemo(() => {
    let c = 0;
    if (fAgents.length) c++;
    if (fStatuses.length) c++;
    if (fCurrencies.length) c++;
    if (fDateFrom || fDateTo) c++;
    if (fQtyFrom !== '' || fQtyTo !== '') c++;
    if (fSumFrom !== '' || fSumTo !== '') c++;
    return c;
  }, [fAgents, fStatuses, fCurrencies, fDateFrom, fDateTo, fQtyFrom, fQtyTo, fSumFrom, fSumTo]);
  const universalFilterCount = useMemo(() => countUniversalFilters(uFilterState), [uFilterState]);

  const exportOrders = () => {
    exportAoaExcel({
      fileName: `Zakazlar_${new Date().toISOString().slice(0,10)}.xlsx`,
      sheetName: 'Zakazlar',
      headers: ['Zakaz', 'Mijoz', 'ID', 'Sana', 'Tur', 'Status', 'Valyuta', 'Dona', 'Summa UZS', 'Summa USD', 'Agent', 'Dostavchik', 'Mahsulot soni'],
      columnTypes: ['text', 'text', 'text', 'date', 'text', 'text', 'text', 'number', 'number', 'number', 'text', 'text', 'number'],
      rows: list.map((g) => [
        g.soNum,
        g.contName,
        g.mId,
        fmtD(g.orderDate),
        g.docType,
        g.status,
        g.currency || 'UZS',
        g.totalQty,
        g.totalSumUZS || 0,
        g.totalSumUSD || 0,
        g.agent || '',
        g.delivPerson || '',
        g.items.length,
      ]),
    });
  };

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:12,height:'100%'}}>
      <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
        <div className="sb" style={{flex:1}}>
          <span style={{color:'var(--t3)'}}>Qidiruv</span>
          <input placeholder="Mijoz, zakaz no..." value={search} onChange={(e)=>setS(e.target.value)}/>
        </div>
        <div className="tabs">
          {[['zakaz','Zakaz'],['vozvrat','Vozvrat'],['all','Barchasi']].map(([t,l]) => (
            <button key={t} className={`tab${fType===t?' on':''}`} onClick={()=>setT(t)}>{l}</button>
          ))}
        </div>
        <div style={{position:'relative'}}>
          <button className="btn btn-gh btn-sm" onClick={()=>setUFilterOpen((v)=>!v)}>
            <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
              <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z"/>
            </svg>
            Filtr{universalFilterCount>0?` (${universalFilterCount})`:''}
          </button>
          <UniversalFilterPanel
            open={uFilterOpen}
            title="Zakazlar filtri"
            columns={orderFilterColumns}
            rows={base}
            state={uFilterState}
            setState={setUFilterState}
            onClose={()=>setUFilterOpen(false)}
            width={620}
          />
          {showFilter && (
            <div className="card" style={{position:'absolute',top:34,left:0,right:'auto',zIndex:50,width:'min(380px, calc(100vw - 24px))',padding:10,overflow:'hidden'}}>
              <div className="g2" style={{gap:10}}>
                <div>
                  <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Agent</div>
                  <div style={{maxHeight:90,overflow:'auto',border:'1px solid var(--b1)',borderRadius:8,padding:6}}>
                    {options.agents.map((x)=><label key={x} style={FILTER_CHECK_LABEL_STYLE}><input type="checkbox" checked={fAgents.includes(x)} onChange={()=>toggleIn(setAgents,x)}/><span style={FILTER_CHECK_TEXT_STYLE}>{x}</span></label>)}
                  </div>
                </div>
                <div>
                  <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Status</div>
                  <div style={{maxHeight:90,overflow:'auto',border:'1px solid var(--b1)',borderRadius:8,padding:6}}>
                    {options.statuses.map((x)=><label key={x} style={FILTER_CHECK_LABEL_STYLE}><input type="checkbox" checked={fStatuses.includes(x)} onChange={()=>toggleIn(setStatuses,x)}/><span style={FILTER_CHECK_TEXT_STYLE}>{x}</span></label>)}
                  </div>
                </div>
                <div>
                  <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Valyuta</div>
                  <div style={{maxHeight:90,overflow:'auto',border:'1px solid var(--b1)',borderRadius:8,padding:6}}>
                    {options.currencies.map((x)=><label key={x} style={FILTER_CHECK_LABEL_STYLE}><input type="checkbox" checked={fCurrencies.includes(x)} onChange={()=>toggleIn(setCurrencies,x)}/><span style={FILTER_CHECK_TEXT_STYLE}>{x}</span></label>)}
                  </div>
                </div>
                <div style={{display:'grid',gap:6}}>
                  <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
                    <ModernDateInput value={fDateFrom} onChange={(e)=>setDateFrom(e.target.value)} />
                    <ModernDateInput value={fDateTo} onChange={(e)=>setDateTo(e.target.value)} />
                  </div>
                  <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
                    <input className="input" placeholder="Dona dan" value={fQtyFrom} onChange={(e)=>setQtyFrom(e.target.value)} />
                    <input className="input" placeholder="Dona gacha" value={fQtyTo} onChange={(e)=>setQtyTo(e.target.value)} />
                  </div>
                  <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
                    <input className="input" placeholder="Summa dan (UZS)" value={fSumFrom} onChange={(e)=>setSumFrom(e.target.value)} />
                    <input className="input" placeholder="Summa gacha (UZS)" value={fSumTo} onChange={(e)=>setSumTo(e.target.value)} />
                  </div>
                </div>
              </div>
              <div style={{display:'flex',justifyContent:'space-between',marginTop:8}}>
                <button className="btn btn-gh btn-sm" onClick={()=>{
                  setAgents([]); setStatuses([]); setCurrencies([]);
                  setDateFrom(''); setDateTo('');
                  setQtyFrom(''); setQtyTo('');
                  setSumFrom(''); setSumTo('');
                }}>Tozalash</button>
                <button className="btn btn-bl btn-sm" onClick={()=>setShowFilter(false)}>Qo'llash</button>
              </div>
            </div>
          )}
        </div>
        <button className="btn btn-gr btn-sm" onClick={exportOrders}>? Excel</button>
      </div>
      <div className="card" style={{overflow:'hidden',flex:1}}>
        <div style={{overflow:'auto',maxHeight:'100%'}}>
          <table className="tbl">
            <thead>
              <tr>
                <th>Zakaz No</th><th>Kontragent</th><th>Rayon</th><th>Sana</th>
                <th>Dostavchik</th><th style={{textAlign:'center'}}>Dona</th>
                <th style={{textAlign:'right'}}>Summa UZS</th>
                <th style={{textAlign:'right'}}>Summa USD</th>
                <th style={{textAlign:'center'}}>Mahsulotlar</th>
                <th>Tur</th><th>Agent</th><th>Status</th>
              </tr>
            </thead>
            <tbody>
              {list.length===0
                ? <tr><td colSpan={12} style={{textAlign:'center',padding:40,color:'var(--t3)'}}>Topilmadi</td></tr>
                : list.slice(0,600).map((g,i) => (
                  <tr key={i} onClick={()=>setSelectedSO(g)}>
                    <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{g.soNum}</td>
                    <td style={{maxWidth:180,fontWeight:600}}>
                      <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:178,fontSize:12}}>{g.contName}</span>
                    </td>
                    <td style={{fontSize:11,color:'var(--t3)'}}>{g.district || '-'}</td>
                    <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(g.orderDate)}</td>
                    <td style={{fontSize:11,color:'var(--t2)'}}>{g.delivPerson||'-'}</td>
                    <td style={{textAlign:'center',fontWeight:700,color:'var(--t1)'}}>{g.totalQty}</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontSize:11.5,fontWeight:700,color:g.totalSumUZS?'var(--gr)':'var(--t4)'}}>
                      {g.totalSumUZS?fmt(g.totalSumUZS):'-'}
                    </td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontSize:11,fontWeight:700,color:g.totalSumUSD?'var(--yl)':'var(--t4)'}}>
                      {g.totalSumUSD?fmt(g.totalSumUSD)+' $':'-'}
                    </td>
                    <td style={{textAlign:'center'}}>
                      <span className="tag" style={{background:'var(--s3)',color:'var(--t3)',cursor:'pointer'}}>{g.items.length} ta</span>
                    </td>
                    <td>
                      <span className="tag" style={{background:isOrderDoc(g.docType)?'var(--bl3)':'var(--or2)',color:isOrderDoc(g.docType)?'var(--bl)':'var(--or)',fontSize:10}}>
                        {g.docType}
                      </span>
                    </td>
                    <td style={{fontSize:12}}>{g.agent||'-'}</td>
                    <td>
                      <span className="tag" style={{
                        background:isDeliveredStatus(g.status)?'var(--gr2)':isCancelledStatus(g.status)?'var(--rd2)':'var(--s3)',
                        color:isDeliveredStatus(g.status)?'var(--gr)':isCancelledStatus(g.status)?'var(--t4)':'var(--t3)',
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
  const { cashbox=[], customers=[] } = D;
  const [fType,setT]  = useState('all');
  const [search,setS] = useState('');
  const [showFilter,setShowFilter] = useState(false);
  const [uFilterOpen,setUFilterOpen] = useState(false);
  const [uFilterState,setUFilterState] = useState({});
  const [fOperators,setOperators] = useState([]);
  const [fOpTypes,setOpTypes] = useState([]);
  const [fKassas,setKassas] = useState([]);
  const [fDateFrom,setDateFrom] = useState('');
  const [fDateTo,setDateTo] = useState('');
  const [fAmountFrom,setAmountFrom] = useState('');
  const [fAmountTo,setAmountTo] = useState('');
  const districtById = useMemo(() => {
    const m = {};
    (customers || []).forEach((c) => { m[String(c.id || '').trim()] = c.district || ''; });
    return m;
  }, [customers]);
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
  const options = useMemo(() => ({
    operators: [...new Set(cashbox.map((x)=>x.operator).filter(Boolean))].sort(),
    opTypes: [...new Set(cashbox.map((x)=>x.opType).filter(Boolean))].sort(),
    kassas: [...new Set(cashbox.map((x)=>x.kassa).filter(Boolean))].sort(),
  }), [cashbox]);
  const toggleIn = (setter, val) => setter((prev) => prev.includes(val) ? prev.filter((x)=>x!==val) : [...prev, val]);
  const kassaFilterColumns = useMemo(() => ([
    { key:'sana', label:'Sana', type:'date' },
    { key:'opNum', label:'Op No', type:'text' },
    { key:'opType', label:'Tur', type:'text' },
    { key:'kassa', label:'Kassa', type:'text' },
    { key:'contName', label:'Kontragent', type:'text' },
    { key:'mId', label:'Mijoz ID', type:'text' },
    { key:'district', label:'Rayon', type:'text', getValue:(r)=>districtById[String(r.mId || '').trim()] || '' },
    { key:'amount', label:'Summa', type:'number' },
    { key:'operator', label:'Operator', type:'text' },
    { key:'opStatus', label:'Status', type:'text' },
    { key:'currency', label:'Valyuta', type:'text' },
  ]), [districtById]);
  useEffect(() => {
    setUFilterState((prev) => ensureUniversalFilterState(kassaFilterColumns, prev));
  }, [kassaFilterColumns]);
  const inRange = (v, from, to) => {
    const n = Number(v || 0);
    if (from !== '' && n < Number(from)) return false;
    if (to !== '' && n > Number(to)) return false;
    return true;
  };

  const list = useMemo(() => {
    let r = base;
    if (search) {
      const q = search.toLowerCase();
      r=r.filter((c)=>(c.contName||'').toLowerCase().includes(q)||(c.opNum||'').toLowerCase().includes(q));
    }
    if (fOperators.length) r = r.filter((c)=>fOperators.includes(c.operator || ''));
    if (fOpTypes.length) r = r.filter((c)=>fOpTypes.includes(c.opType || ''));
    if (fKassas.length) r = r.filter((c)=>fKassas.includes(c.kassa || ''));
    if (fDateFrom) {
      const df = toDate(fDateFrom);
      if (df) r = r.filter((c)=> {
        const d = toDate(c.sana);
        return d ? d >= df : false;
      });
    }
    if (fDateTo) {
      const dt = toDate(fDateTo);
      if (dt) r = r.filter((c)=> {
        const d = toDate(c.sana);
        return d ? d <= dt : false;
      });
    }
    r = r.filter((c)=>inRange(c.amount, fAmountFrom, fAmountTo));
    r = applyUniversalFilters(r, kassaFilterColumns, uFilterState);
    return [...r].sort((a,b) => {
      const da=toDate(a.sana), db=toDate(b.sana);
      return da&&db?db-da:0;
    });
  }, [base,search,fOperators,fOpTypes,fKassas,fDateFrom,fDateTo,fAmountFrom,fAmountTo,kassaFilterColumns,uFilterState]);
  const activeFilterCount = useMemo(() => {
    let c = 0;
    if (fOperators.length) c++;
    if (fOpTypes.length) c++;
    if (fKassas.length) c++;
    if (fDateFrom || fDateTo) c++;
    if (fAmountFrom !== '' || fAmountTo !== '') c++;
    return c;
  }, [fOperators, fOpTypes, fKassas, fDateFrom, fDateTo, fAmountFrom, fAmountTo]);
  const universalFilterCount = useMemo(() => countUniversalFilters(uFilterState), [uFilterState]);
  const exportKassa = () => {
    exportAoaExcel({
      fileName: `Kassa_${new Date().toISOString().slice(0,10)}.xlsx`,
      sheetName: 'Kassa',
      headers: ['Sana', 'Op No', 'Tur', 'Kassa', 'Kontragent', 'Summa', 'Operator', 'Status', 'Valyuta', 'Mijoz ID', 'Izoh'],
      columnTypes: ['date', 'text', 'text', 'text', 'text', 'number', 'text', 'text', 'text', 'text', 'text'],
      rows: list.map((c) => [
        fmtD(c.sana),
        c.opNum || '',
        c.opType || '',
        c.kassa || '',
        c.contName || '',
        c.amount || 0,
        c.operator || '',
        c.opStatus || '',
        c.currency || 'UZS',
        c.mId || '',
        c.note || '',
      ]),
    });
  };

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:12,height:'100%'}}>
      <div className="g4">
        <StatCard l={`KIRIM (${curMonthKey})`}  v={fmt(totalIn)+" so'm"}  s={paysMonth.length+' ta operatsiya'}   c="var(--gr)"/>
        <StatCard l={`CHIQIM (${curMonthKey})`} v={fmt(totalOut)+" so'm"} s={spendsMonth.length+' ta operatsiya'} c="var(--rd)"/>
        <StatCard l="SALDO (OYLIK)" v={fmt(totalIn-totalOut)+" so'm"} s="joriy oy: kirim minus chiqim" c={totalIn-totalOut>=0?'var(--gr)':'var(--rd)'}/>
        <StatCard l="OYLIK OPR."   v={cashMonth.length+' ta'}  s="joriy oy yozuvlari" c="var(--bl)"/>
      </div>
      <div style={{display:'flex',gap:8,alignItems:'center'}}>
        <div className="sb" style={{flex:1}}>
          <span style={{color:'var(--t3)'}}>{E.find}</span>
          <input placeholder="Mijoz, operatsiya no..." value={search} onChange={(e)=>setS(e.target.value)}/>
        </div>
        <div style={{position:'relative'}}>
          <button className="btn btn-gh btn-sm" onClick={()=>setUFilterOpen((v)=>!v)}>
            <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
              <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z"/>
            </svg>
            Filtr{universalFilterCount>0?` (${universalFilterCount})`:''}
          </button>
          <UniversalFilterPanel
            open={uFilterOpen}
            title="Kassa filtri"
            columns={kassaFilterColumns}
            rows={base}
            state={uFilterState}
            setState={setUFilterState}
            onClose={()=>setUFilterOpen(false)}
            width={620}
          />
          {showFilter && (
            <div className="card" style={{position:'absolute',top:34,left:0,right:'auto',zIndex:50,width:'min(380px, calc(100vw - 24px))',padding:10,overflow:'hidden'}}>
              <div className="g2" style={{gap:10}}>
                <div>
                  <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Operator</div>
                  <div style={{maxHeight:90,overflow:'auto',border:'1px solid var(--b1)',borderRadius:8,padding:6}}>
                    {options.operators.map((x)=><label key={x} style={FILTER_CHECK_LABEL_STYLE}><input type="checkbox" checked={fOperators.includes(x)} onChange={()=>toggleIn(setOperators,x)}/><span style={FILTER_CHECK_TEXT_STYLE}>{x}</span></label>)}
                  </div>
                </div>
                <div>
                  <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Tur</div>
                  <div style={{maxHeight:90,overflow:'auto',border:'1px solid var(--b1)',borderRadius:8,padding:6}}>
                    {options.opTypes.map((x)=><label key={x} style={FILTER_CHECK_LABEL_STYLE}><input type="checkbox" checked={fOpTypes.includes(x)} onChange={()=>toggleIn(setOpTypes,x)}/><span style={FILTER_CHECK_TEXT_STYLE}>{x}</span></label>)}
                  </div>
                </div>
                <div>
                  <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Kassa</div>
                  <div style={{maxHeight:90,overflow:'auto',border:'1px solid var(--b1)',borderRadius:8,padding:6}}>
                    {options.kassas.map((x)=><label key={x} style={FILTER_CHECK_LABEL_STYLE}><input type="checkbox" checked={fKassas.includes(x)} onChange={()=>toggleIn(setKassas,x)}/><span style={FILTER_CHECK_TEXT_STYLE}>{x}</span></label>)}
                  </div>
                </div>
                <div style={{display:'grid',gap:6}}>
                  <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
                    <ModernDateInput value={fDateFrom} onChange={(e)=>setDateFrom(e.target.value)} />
                    <ModernDateInput value={fDateTo} onChange={(e)=>setDateTo(e.target.value)} />
                  </div>
                  <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:6}}>
                    <input className="input" placeholder="Summa dan" value={fAmountFrom} onChange={(e)=>setAmountFrom(e.target.value)} />
                    <input className="input" placeholder="Summa gacha" value={fAmountTo} onChange={(e)=>setAmountTo(e.target.value)} />
                  </div>
                </div>
              </div>
              <div style={{display:'flex',justifyContent:'space-between',marginTop:8}}>
                <button className="btn btn-gh btn-sm" onClick={()=>{
                  setOperators([]); setOpTypes([]); setKassas([]);
                  setDateFrom(''); setDateTo('');
                  setAmountFrom(''); setAmountTo('');
                }}>Tozalash</button>
                <button className="btn btn-bl btn-sm" onClick={()=>setShowFilter(false)}>Qo'llash</button>
              </div>
            </div>
          )}
        </div>
        <button className="btn btn-gr btn-sm" onClick={exportKassa}>? Excel</button>
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
                <th>Sana</th><th>Op No</th><th>Tur</th>
                <th>Kassa</th><th>Kontragent</th><th>Rayon</th><th>Summa</th><th>Operator</th>
              </tr>
            </thead>
            <tbody>
              {list.length===0
                ? <tr><td colSpan={8} style={{textAlign:'center',padding:40,color:'var(--t3)'}}>Topilmadi</td></tr>
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
                        <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:108}}>{c.kassa||'-'}</span>
                      </td>
                      <td style={{maxWidth:220}}>
                        <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',maxWidth:218,fontSize:12}}>{c.contName||'-'}</span>
                      </td>
                      <td style={{fontSize:11.5,color:'var(--t3)'}}>{districtById[String(c.mId || '').trim()] || '-'}</td>
                      <td style={{fontFamily:'var(--mono)',fontWeight:700,color:isIn?'var(--gr)':isPaymentToCounterparty(c.opType)?'var(--rd)':'var(--t2)'}}>
                        {isIn?'+':isPaymentToCounterparty(c.opType)?'-':''}{fmt(c.amount)}
                      </td>
                      <td style={{fontSize:11.5,color:'var(--t3)'}}>{c.operator||'-'}</td>
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

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў OBZVON Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
function Obzvon({
  D,
  company='murodbaxsh',
  currentUser='Admin',
  canEditAllNew=false,
  canDeleteAllNew=false,
  canPublishAllNew=false,
  records=[],
  setRecords=()=>{},
  allRows=[],
  newRows=[],
  onReloadAll=()=>{},
  onReloadAllNew=()=>{},
  onAppendAllRows=()=>{},
  onUpsertNewRows=()=>{},
  onPublishAllNew=()=>{},
}) {
  const { customers, rawOrders=[] } = D;
  const [tab, setTab] = useState('main');
  const [searchAll, setSearchAll] = useState('');
  const [searchAllNew, setSearchAllNew] = useState('');
  const [pickQuery, setPickQuery] = useState('');
  const [allShowMode, setAllShowMode] = useState('smart');
  const [mainFilterOpen, setMainFilterOpen] = useState(false);
  const [mainFilterState, setMainFilterState] = useState({});
  const [allUniFilterOpen, setAllUniFilterOpen] = useState(false);
  const [allUniFilterState, setAllUniFilterState] = useState({});
  const [allNewUniFilterOpen, setAllNewUniFilterOpen] = useState(false);
  const [allNewUniFilterState, setAllNewUniFilterState] = useState({});
  const [allNewEditRid, setAllNewEditRid] = useState('');
  const [allNewEditDraft, setAllNewEditDraft] = useState(null);
  const [dueUniFilterOpen, setDueUniFilterOpen] = useState(false);
  const [dueUniFilterState, setDueUniFilterState] = useState({});
  const [opUniFilterOpen, setOpUniFilterOpen] = useState(false);
  const [opUniFilterState, setOpUniFilterState] = useState({});
  const [opLimit, setOpLimit] = useState(500);
  const [pickTargetIdx, setPickTargetIdx] = useState(null);
  const [opSearch, setOpSearch] = useState('');
  const [opPickMode, setOpPickMode] = useState(false);
  const [opSelectedIds, setOpSelectedIds] = useState({});
  const [duePickMode, setDuePickMode] = useState(false);
  const [dueSelectedIds, setDueSelectedIds] = useState({});
  const [latePickMode, setLatePickMode] = useState(false);
  const [lateSelectedIds, setLateSelectedIds] = useState({});
  const deferredSearchAll = useDeferredValue(searchAll);
  const deferredSearchAllNew = useDeferredValue(searchAllNew);
  const deferredOpSearch = useDeferredValue(opSearch);
  const todayIso = () => new Date().toISOString().slice(0,10);
  const customerNameById = useMemo(() => {
    const m = {};
    (customers || []).forEach((c) => { m[String(c.id || '').trim()] = c.name || ''; });
    return m;
  }, [customers]);
  const customerBalanceById = useMemo(() => {
    const m = {};
    (customers || []).forEach((c) => {
      const cid = String(c.id || '').trim();
      if (!cid) return;
      m[cid] = Number(c.balanceUZS || 0);
    });
    return m;
  }, [customers]);
  const mainFilterColumns = useMemo(() => ([
    { key:'id', label:'ID', type:'text' },
    { key:'customer', label:'Mijoz', type:'text' },
    { key:'callDate', label:'Sana', type:'date' },
    { key:'topic', label:'Maqsad', type:'text' },
    { key:'note', label:'Izoh', type:'text' },
    { key:'nextDate', label:'Keyingi sana', type:'date' },
    { key:'orderCount', label:'Z.soni / summa', type:'number' },
    { key:'orderDate', label:'Zakaz sanasi', type:'date' },
    { key:'operator', label:'Operator', type:'text' },
  ]), []);
  const allFilterColumns = useMemo(() => ([
    { key:'no', label:'No', type:'number' },
    { key:'customerId', label:'ID', type:'text' },
    { key:'customer', label:'Kontragent', type:'text' },
    { key:'callDate', label:'Sana', type:'date' },
    { key:'topic', label:'Mavzu', type:'text' },
    { key:'note', label:'Izoh', type:'text' },
    { key:'nextDate', label:'Keyingi sana', type:'date' },
    { key:'orderCount', label:'Z.soni / summa', type:'number' },
    { key:'orderDate', label:'Zakaz sanasi', type:'date' },
    { key:'operator', label:'Operator', type:'text' },
  ]), []);
  const dueFilterColumns = useMemo(() => ([
    { key:'id', label:'ID', type:'text' },
    { key:'name', label:'Mijoz', type:'text' },
    { key:'district', label:'Rayon', type:'text' },
    { key:'phone', label:'Telefon', type:'text' },
    { key:'last3Info', label:'Oxirgi zakazlar', type:'text' },
    { key:'passed', label:"O'tgan kun", type:'number' },
    { key:'shouldIn', label:"Qo'ng'iroq me'yori", type:'number' },
  ]), []);
  const opFilterColumns = useMemo(() => ([
    { key:'id', label:'ID', type:'text' },
    { key:'name', label:'Mijoz', type:'text' },
    { key:'district', label:'Rayon', type:'text' },
    { key:'balance', label:'Balans', type:'number' },
    { key:'ord1', label:'Oxirgi zakaz', type:'date' },
    { key:'ord2', label:'Oxirgidan oldingi zakaz', type:'date' },
    { key:'lastQty', label:'Zakaz soni', type:'number' },
    { key:'lastCallDate', label:"Oxirgi qo'ng'iroq", type:'date' },
    { key:'nextDate', label:'Keyingi sana', type:'date' },
    { key:'operator', label:'Operator', type:'text' },
    { key:'lastNote', label:'Oxirgi izoh', type:'text' },
  ]), []);
  useEffect(() => {
    setMainFilterState((prev) => ensureUniversalFilterState(mainFilterColumns, prev));
    setAllUniFilterState((prev) => ensureUniversalFilterState(allFilterColumns, prev));
    setAllNewUniFilterState((prev) => ensureUniversalFilterState(allFilterColumns, prev));
    setDueUniFilterState((prev) => ensureUniversalFilterState(dueFilterColumns, prev));
    setOpUniFilterState((prev) => ensureUniversalFilterState(opFilterColumns, prev));
  }, [mainFilterColumns, allFilterColumns, dueFilterColumns, opFilterColumns]);
  const getDebtValueText = useCallback((row) => {
    if ((row?.topic || '') !== 'Qarzdorlik') return String(row?.orderCount || '');
    const cid = String(row?.id || '').trim();
    if (!cid) return String(row?.orderCount || '');
    const bal = Number(customerBalanceById[cid] || 0);
    return String(Math.round(bal));
  }, [customerBalanceById]);
  // Barcha Obzvon Yangi ga tushish sharti: mijoz + izoh to'ldirilgan bo'lishi kerak.
  const hasObzvonPayload = useCallback((row) => {
    return Boolean(
      String(row?.customer || '').trim() &&
      String(row?.note || '').trim()
    );
  }, []);
  const saveRecords = (next) => {
    const prepared = (next || []).map((r) => ({
      ...r,
      _rid: r?._rid || createRowRid(),
      orderCount: getDebtValueText(r),
      company: normalizeCompanyKey(r?.company || company),
    }));
    setRecords(prepared);
    const freshRows = prepared
      .filter((r) => (r.id || r.customer) && hasObzvonPayload(r))
      .map((r, i) => ({
        rid: r._rid,
        no: String(r.no || i + 1),
        customer: r.customer || '',
        callDate: r.callDate || '',
        topic: r.topic || '',
        note: r.note || '',
        nextDate: r.nextDate || '',
        orderCount: r.orderCount || '',
        operator: r.operator || currentUser,
        customerId: r.id || '',
        orderDate: r.orderDate || '',
        company: normalizeCompanyKey(r.company || company),
      }));
    if (freshRows.length) onUpsertNewRows(freshRows);
  };
  const appendRows = (count=1) => {
    const rows = Array.from({ length:count }, () => ({
      id: '',
      customer: '',
      callDate: todayIso(),
      topic: 'Buyurtma olish',
      note: '',
      nextDate: '',
      orderCount: '',
      orderDate: '',
      operator: currentUser,
      company: normalizeCompanyKey(company),
    }));
    saveRecords([...(records||[]), ...rows]);
  };
  const addRecord = async (customer, topic='Buyurtma olish') => {
    const row = {
      id: customer.id,
      customer: customer.name,
      callDate: todayIso(),
      topic,
      note: '',
      nextDate: '',
      orderCount: '',
      orderDate: '',
      operator: currentUser,
      company: normalizeCompanyKey(company),
    };
    const next = [row, ...records];
    saveRecords(next);
  };
  const archiveCurrentUserRows = () => {
    const mine = (records || []).filter((r) => (r.operator || currentUser) === currentUser);
    const archived = mine.filter((r) => (r.customer || r.id) && hasObzvonPayload(r));
    if (!archived.length) return;
    const normalized = archived.map((r, i) => ({
      rid: r._rid || createRowRid(),
      no: String(i + 1),
      customer: r.customer || '',
      callDate: r.callDate || '',
      topic: r.topic || '',
      note: r.note || '',
      nextDate: r.nextDate || '',
      orderCount: r.orderCount || '',
      operator: r.operator || currentUser,
      customerId: r.id || '',
      orderDate: r.orderDate || '',
      company: normalizeCompanyKey(r.company || company),
    }));
    onUpsertNewRows(normalized);
    const left = (records || []).filter((r) => {
      if ((r.operator || currentUser) !== currentUser) return true;
      return !(r.customer || r.id) || !hasObzvonPayload(r);
    });
    saveRecords(left);
  };

  const historicalAllRows = useMemo(() => {
    const rows = (allRows || [])
      .map((r, i) => ({
        no: String(r.no || i + 1),
        customer: r.customer || '',
        callDate: r.callDate || '',
        topic: r.topic || '',
        note: r.note || '',
        nextDate: r.nextDate || '',
        orderCount: r.orderCount || '',
        operator: r.operator || '',
        customerId: r.customerId || r.id || '',
        orderDate: r.orderDate || '',
        company: normalizeCompanyKey(r.company || company),
      }))
      .filter((r) => !isAccessSyncRow(r))
      .map((r) => {
        const cid = String(r.customerId || r.id || '').trim();
        const liveName = cid ? (customerNameById[cid] || '') : '';
        return {
          ...r,
          customerId: cid,
          customer: liveName || r.customer || (cid ? `ID: ${cid}` : ''),
        };
      });
    rows.sort((a, b) => {
      const da = toDate(a.callDate);
      const db = toDate(b.callDate);
      return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
    });
    return rows;
  }, [allRows, customerNameById, company]);

  const appAllNewRows = useMemo(() => {
    const fromNew = (newRows || []).map((r, i) => ({
      rid: String(r.rid || r._rid || '').trim() || buildFallbackObzvonRid(r, i),
      no: String(r.no || i + 1),
      customer: r.customer || '',
      callDate: r.callDate || '',
      topic: r.topic || '',
      note: r.note || '',
      nextDate: r.nextDate || '',
      orderCount: r.orderCount || '',
      operator: r.operator || '',
      customerId: r.customerId || r.id || '',
      orderDate: r.orderDate || '',
      company: normalizeCompanyKey(r.company || company),
    }));
    const fromCurrent = (records || [])
      .filter((r) => (r.id || r.customer) && hasObzvonPayload(r))
      .map((r, i) => ({
        rid: String(r._rid || '').trim() || buildFallbackObzvonRid(r, i),
        no: String(r.no || i + 1),
        customer: r.customer || '',
        callDate: r.callDate || '',
        topic: r.topic || '',
        note: r.note || '',
        nextDate: r.nextDate || '',
        orderCount: r.orderCount || '',
        operator: r.operator || '',
        customerId: r.id || '',
        orderDate: r.orderDate || '',
        company: normalizeCompanyKey(r.company || company),
      }));
    const m = new Map();
    [...fromNew, ...fromCurrent].forEach((r) => {
      m.set(r.rid, r);
    });
    const rows = Array.from(m.values()).map((r) => {
      const cid = String(r.customerId || '').trim();
      const liveName = cid ? (customerNameById[cid] || '') : '';
      return {
        ...r,
        customerId: cid,
        customer: liveName || r.customer || (cid ? `ID: ${cid}` : ''),
        company: normalizeCompanyKey(r.company || company),
      };
    }).filter((r) => !isObzvonDeletedRow(r));
    rows.sort((a, b) => {
      const da = toDate(a.callDate);
      const db = toDate(b.callDate);
      return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
    });
    return rows;
  }, [newRows, records, customerNameById, hasObzvonPayload, company]);

  const mainRows = useMemo(() => {
    const base = (records || []).map((r) => ({
      ...r,
      _rid: r?._rid || createRowRid(),
      orderCount: getDebtValueText(r),
    }));
    return applyUniversalFilters(base, mainFilterColumns, mainFilterState);
  }, [records, getDebtValueText, mainFilterColumns, mainFilterState]);
  const allList = useMemo(() => {
    const q = String(deferredSearchAll || '').toLowerCase().trim();
    let base = [...(historicalAllRows || [])];
    if (q) {
      base = base.filter((r) =>
        String(r.customer || '').toLowerCase().includes(q) ||
        String(r.customerId || '').toLowerCase().includes(q) ||
        String(r.operator || '').toLowerCase().includes(q) ||
        String(r.note || '').toLowerCase().includes(q)
      );
    }
    return applyUniversalFilters(base, allFilterColumns, allUniFilterState);
  }, [historicalAllRows, deferredSearchAll, allFilterColumns, allUniFilterState]);
  const allNewList = useMemo(() => {
    const q = String(deferredSearchAllNew || '').toLowerCase().trim();
    let base = [...(appAllNewRows || [])];
    if (q) {
      base = base.filter((r) =>
        String(r.customer || '').toLowerCase().includes(q) ||
        String(r.customerId || '').toLowerCase().includes(q) ||
        String(r.operator || '').toLowerCase().includes(q) ||
        String(r.note || '').toLowerCase().includes(q)
      );
    }
    const filtered = applyUniversalFilters(base, allFilterColumns, allNewUniFilterState);
    // Barcha qurilmalarda bir xil ko'rinishi uchun No ni har safar qayta ketma-ketlaymiz.
    return filtered.map((r, idx) => ({ ...r, no: String(idx + 1) }));
  }, [appAllNewRows, deferredSearchAllNew, allFilterColumns, allNewUniFilterState]);

  const visibleAllRows = useMemo(
    () => allShowMode === 'all' ? allList : allList.slice(0, 500),
    [allList, allShowMode]
  );
  const visibleAllNewRows = useMemo(
    () => allShowMode === 'all' ? allNewList : allNewList.slice(0, 500),
    [allNewList, allShowMode]
  );
  const startAllNewEdit = (row) => {
    setAllNewEditRid(String(row?.rid || ''));
    setAllNewEditDraft({
      ...row,
      callDate: String(row?.callDate || '').slice(0,10),
      nextDate: String(row?.nextDate || '').slice(0,10),
      orderDate: String(row?.orderDate || '').slice(0,10),
    });
  };
  const cancelAllNewEdit = () => {
    setAllNewEditRid('');
    setAllNewEditDraft(null);
  };
  const saveAllNewEdit = () => {
    if (!allNewEditDraft || !allNewEditRid) return;
    if (!String(allNewEditDraft.customer || allNewEditDraft.customerId || '').trim()) {
      alert("Mijoz bo'sh bo'lishi mumkin emas");
      return;
    }
    if (!String(allNewEditDraft.note || '').trim()) {
      alert("Izoh bo'sh bo'lishi mumkin emas");
      return;
    }
    const updated = {
      ...allNewEditDraft,
      rid: allNewEditRid,
      _rid: allNewEditRid,
      updatedAt: new Date().toISOString(),
      deleted: '',
      company: normalizeCompanyKey(allNewEditDraft.company || company),
    };
    onUpsertNewRows([updated]);
    cancelAllNewEdit();
  };
  const removeAllNewRow = (row) => {
    if (!row) return;
    const rid = row.rid || row._rid || createRowRid();
    const tomb = {
      ...row,
      rid,
      _rid: rid,
      deleted: '1',
      note: OBZVON_DELETED_NOTE,
      updatedAt: new Date().toISOString(),
      company: normalizeCompanyKey(row.company || company),
    };
    onUpsertNewRows([tomb]);
    if (String(row.rid || '') === String(allNewEditRid || '')) cancelAllNewEdit();
  };
  const mainFilterCount = useMemo(() => countUniversalFilters(mainFilterState), [mainFilterState]);
  const allFilterCount = useMemo(() => countUniversalFilters(allUniFilterState), [allUniFilterState]);
  const allNewFilterCount = useMemo(() => countUniversalFilters(allNewUniFilterState), [allNewUniFilterState]);
  const dueSelectedCount = useMemo(
    () => Object.values(dueSelectedIds).filter(Boolean).length,
    [dueSelectedIds]
  );
  const lateSelectedCount = useMemo(
    () => Object.values(lateSelectedIds).filter(Boolean).length,
    [lateSelectedIds]
  );
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

  const callCandidates = useMemo(() => {
    const assignment = D.assignmentById || {};
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
      if (sorted.length < 2) return;

      const newestDate = toDate(sorted[0].orderDate);
      const monthlyCount = sorted.filter((o) => {
        const od = toDate(o.orderDate);
        if (!od || !newestDate) return false;
        return ((newestDate - od) / 864e5) <= 30;
      }).length;
      const useCount = (monthlyCount >= 3 && sorted.length >= 3) ? 3 : 2;
      const used = sorted.slice(0, useCount);

      const newest = toDate(used[0].orderDate);
      const oldest = toDate(used[used.length - 1].orderDate);
      let spanDays = Math.round((newest - oldest) / 864e5);
      if (spanDays < 1) spanDays = 1;

      const totalQty = used.reduce((s,o)=>s+Math.abs(o.qty||0),0);
      const daily = totalQty / spanDays;
      if (daily <= 0) return;
      const lastQty = Math.abs(used[0].qty || 0);
      const shouldIn = Math.max(1, Math.ceil(lastQty / daily) - 1);
      const c = customers.find((x)=>x.id===mid);
      if (!c) return;
      if (!isActiveCustomerName(c.name)) return;
      if (currentUser !== 'Admin') {
        const owner = assignment[mid] || '';
        if (owner && owner !== currentUser) return;
      }
      const passed = daysAgo(used[0].orderDate) ?? 0;
      out.push({
        ...c,
        shouldIn,
        passed,
        sampleSize: useCount,
        last3Info: used
          .slice(0, 2)
          .map((o) => fmtD(o.orderDate))
          .join(' / '),
      });
    });
    return out.sort((a,b)=>(b.passed-a.passed));
  }, [rawOrders, customers, currentUser, D.assignmentById]);
  const dueCandidates = useMemo(
    () => callCandidates.filter((c) => c.passed < 60 && c.passed >= c.shouldIn),
    [callCandidates]
  );
  const staleCandidates = useMemo(
    () => callCandidates.filter((c) => c.passed >= 60),
    [callCandidates]
  );
  const filteredDueCandidates = useMemo(
    () => applyUniversalFilters(dueCandidates, dueFilterColumns, dueUniFilterState),
    [dueCandidates, dueFilterColumns, dueUniFilterState]
  );
  const filteredStaleCandidates = useMemo(
    () => applyUniversalFilters(staleCandidates, dueFilterColumns, dueUniFilterState),
    [staleCandidates, dueFilterColumns, dueUniFilterState]
  );
  const dueFilterCount = useMemo(() => countUniversalFilters(dueUniFilterState), [dueUniFilterState]);
  const latestCallByCustomer = useMemo(() => {
    const isBuyurtma = (x) => String(x?.topic || '').trim() === 'Buyurtma olish';
    const pickLatest = (rows) => {
      const m = {};
      for (const r of (rows || [])) {
        if (!isBuyurtma(r)) continue;
        const cid = String(r.customerId || r.id || '').trim();
        if (!cid) continue;
        const d = toDate(r.callDate);
        const t = d ? d.getTime() : 0;
        const prev = m[cid];
        const prevT = prev ? ((toDate(prev.callDate)?.getTime()) || 0) : -1;
        if (!prev || t >= prevT) m[cid] = r;
      }
      return m;
    };
    const newMap = pickLatest(appAllNewRows);
    const histMap = pickLatest(historicalAllRows);
    const m = {};
    Object.keys(histMap).forEach((cid) => { m[cid] = histMap[cid]; });
    Object.keys(newMap).forEach((cid) => { m[cid] = newMap[cid]; }); // Yangi birinchi prioritet
    return m;
  }, [historicalAllRows, appAllNewRows]);

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

  const operatorTableBaseRows = useMemo(() => {
    let rows = (D.customers || []).map((c) => {
      const lastCall = latestCallByCustomer[c.id];
      const ord = latestOrdersByCustomer[c.id] || { ord1:'', ord2:'', lastQty:0 };
      return {
        id: c.id,
        name: c.name,
        district: pickPreferredDistrict(c.district || ''),
        tara: c.tara,
        balance: c.balanceUZS,
        ord1: ord.ord1,
        ord2: ord.ord2,
        lastQty: ord.lastQty,
        lastCallDate: lastCall?.callDate || '',
        lastNote: lastCall?.note || '',
        nextDate: lastCall?.nextDate || '',
        operator: D.assignmentById?.[c.id] || '-',
      };
    });
    if (currentUser !== 'Admin') rows = rows.filter((r) => r.operator === currentUser);
    return rows;
  }, [D.customers, D.assignmentById, latestCallByCustomer, latestOrdersByCustomer, currentUser]);
  const operatorTableRows = useMemo(() => {
    let rows = [...operatorTableBaseRows];
    if (deferredOpSearch) {
      const q = deferredOpSearch.toLowerCase();
      rows = rows.filter((r) =>
        String(r.id || '').includes(q) ||
        String(r.name || '').toLowerCase().includes(q) ||
        String(r.district || '').toLowerCase().includes(q) ||
        String(r.ord1 || '').toLowerCase().includes(q) ||
        String(r.ord2 || '').toLowerCase().includes(q) ||
        String(r.operator || '').toLowerCase().includes(q) ||
        String(r.lastNote || '').toLowerCase().includes(q)
      );
    }
    return applyUniversalFilters(rows, opFilterColumns, opUniFilterState);
  }, [
    operatorTableBaseRows,
    deferredOpSearch,
    opFilterColumns,
    opUniFilterState,
  ]);
  const opUniversalFilterCount = useMemo(() => countUniversalFilters(opUniFilterState), [opUniFilterState]);
  useEffect(() => {
    if (tab === 'all_new') onReloadAllNew && onReloadAllNew();
  }, [tab, onReloadAllNew]);

  useEffect(() => {
    setOpLimit(500);
  }, [opSearch]);

  const visibleOperatorRows = useMemo(
    () => operatorTableRows.slice(0, opLimit),
    [operatorTableRows, opLimit]
  );
  const exportCurrentTab = () => {
    if (tab === 'main') {
      exportAoaExcel({
        fileName: `Obzvon_${new Date().toISOString().slice(0,10)}.xlsx`,
        sheetName: 'Obzvon',
        headers: ['ID', 'Mijoz', 'Sana', 'Maqsad', 'Izoh', 'Keyingi sana', 'Z.soni / summa', 'Zakaz sanasi', 'Operator'],
        columnTypes: ['text', 'text', 'date', 'text', 'text', 'date', 'number', 'date', 'text'],
        rows: (records || []).map((r) => [r.id || '', r.customer || '', fmtD(r.callDate), r.topic || '', r.note || '', fmtD(r.nextDate), r.orderCount || '', fmtD(r.orderDate), r.operator || '']),
      });
      return;
    }
    if (tab === 'all') {
      exportAoaExcel({
        fileName: `Barcha_obzvon_${new Date().toISOString().slice(0,10)}.xlsx`,
        sheetName: 'BarchaObzvon',
        headers: ['No', 'ID', 'Mijoz', 'Sana', 'Mavzu', 'Izoh', 'Keyingi sana', 'Z.soni / summa', 'Zakaz sanasi', 'Operator'],
        columnTypes: ['number', 'text', 'text', 'date', 'text', 'text', 'date', 'number', 'date', 'text'],
        rows: allList.map((r, i) => [r.no || i + 1, r.customerId || '', r.customer || '', fmtD(r.callDate), r.topic || '', r.note || '', fmtD(r.nextDate), r.orderCount || '', fmtD(r.orderDate), r.operator || '']),
      });
      return;
    }
    if (tab === 'all_new') {
      exportAoaExcel({
        fileName: `Barcha_obzvon_yangi_${new Date().toISOString().slice(0,10)}.xlsx`,
        sheetName: 'BarchaObzvonYangi',
        headers: ['No', 'ID', 'Mijoz', 'Sana', 'Mavzu', 'Izoh', 'Keyingi sana', 'Z.soni / summa', 'Zakaz sanasi', 'Operator'],
        columnTypes: ['number', 'text', 'text', 'date', 'text', 'text', 'date', 'number', 'date', 'text'],
        rows: allNewList.map((r, i) => [r.no || i + 1, r.customerId || '', r.customer || '', fmtD(r.callDate), r.topic || '', r.note || '', fmtD(r.nextDate), r.orderCount || '', fmtD(r.orderDate), r.operator || '']),
      });
      return;
    }
    if (tab === 'due') {
      exportAoaExcel({
        fileName: `Vaqti_kelgan_${new Date().toISOString().slice(0,10)}.xlsx`,
        sheetName: 'VaqtiKelgan',
        headers: ['ID', 'Mijoz', 'Telefon', 'Oxirgi zakazlar', 'Otgan kun', "Qongiroq meyori"],
        columnTypes: ['text', 'text', 'text', 'text', 'number', 'number'],
        rows: dueCandidates.map((r) => [r.id, r.name, r.phone || '', r.last3Info || '', r.passed ?? '', r.shouldIn ?? '']),
      });
      return;
    }
    if (tab === 'late2m') {
      exportAoaExcel({
        fileName: `Ikki_oydan_otgan_${new Date().toISOString().slice(0,10)}.xlsx`,
        sheetName: '2OydanOtgan',
        headers: ['ID', 'Mijoz', 'Telefon', 'Oxirgi zakazlar', 'Otgan kun', "Qongiroq meyori"],
        columnTypes: ['text', 'text', 'text', 'text', 'number', 'number'],
        rows: staleCandidates.map((r) => [r.id, r.name, r.phone || '', r.last3Info || '', r.passed ?? '', r.shouldIn ?? '']),
      });
      return;
    }
    if (tab === 'op') {
      exportAoaExcel({
        fileName: `Operator_jadvali_${new Date().toISOString().slice(0,10)}.xlsx`,
        sheetName: 'OperatorJadvali',
        headers: ['ID', 'Mijoz', 'Rayon', 'Balans', 'Oxirgi zakaz sana', 'Oldingi zakaz sana', 'Zakaz soni', "Oxirgi qongiroq", 'Keyingi sana', 'Operator', 'Oxirgi izoh'],
        columnTypes: ['text', 'text', 'text', 'number', 'date', 'date', 'number', 'date', 'date', 'text', 'text'],
        rows: operatorTableRows.map((r) => [r.id, r.name, r.district || '', r.balance || 0, fmtD(r.ord1), fmtD(r.ord2), r.lastQty || 0, fmtD(r.lastCallDate), fmtD(r.nextDate), r.operator || '', r.lastNote || '']),
      });
    }
  };

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:10}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:10,flexWrap:'wrap'}}>
        <div className="tabs" style={{display:'inline-flex'}}>
          <button className={`tab${tab==='main'?' on':''}`} onClick={()=>setTab('main')}>Obzvon</button>
          <button className={`tab${tab==='all'?' on':''}`} onClick={()=>setTab('all')}>Barcha Obzvon</button>
          <button className={`tab${tab==='all_new'?' on':''}`} onClick={()=>setTab('all_new')}>Barcha Obzvon Yangi</button>
          <button className={`tab${tab==='due'?' on':''}`} onClick={()=>setTab('due')}>Vaqti Kelgan</button>
          <button className={`tab${tab==='op'?' on':''}`} onClick={()=>setTab('op')}>Operator jadvali</button>
          <button className={`tab${tab==='late2m'?' on':''}`} onClick={()=>setTab('late2m')}>2 oydan o'tgan mijozlar</button>
        </div>
        <button className="btn btn-gr btn-sm" onClick={exportCurrentTab}>? Excel</button>
      </div>

      {tab==='main' && (
        <>
          <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
            <div style={{position:'relative'}}>
              <button className="btn btn-gh btn-sm" onClick={()=>setMainFilterOpen((v)=>!v)}>
                <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                  <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z"/>
                </svg>
                Filtr ({mainFilterCount})
              </button>
              <UniversalFilterPanel
                open={mainFilterOpen}
                title="Obzvon filtri"
                columns={mainFilterColumns}
                rows={records || []}
                state={mainFilterState}
                setState={setMainFilterState}
                onClose={()=>setMainFilterOpen(false)}
                width={620}
              />
            </div>
            <span className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>{mainRows.length} / {records.length} ta yozuv</span>
            <button className="btn btn-gh btn-sm" onClick={()=>appendRows(1)}>+ Qator</button>
            <button className="btn btn-gh btn-sm" onClick={archiveCurrentUserRows}>
              Kun yakunlash
            </button>
          </div>
          <div className="card" style={{overflow:'hidden'}}>
            <div style={{overflow:'auto',maxHeight:'58vh'}}>
              <table className="tbl">
                <thead><tr><th>ID</th><th>Mijoz</th><th>Sana</th><th>Maqsad</th><th>Izoh</th><th>Keyingi sana</th><th>Z.soni / summa</th><th>Zakaz sana</th><th>Operator</th><th></th></tr></thead>
                <tbody>
                  {mainRows.length===0 ? <tr><td colSpan={10} style={{textAlign:'center',padding:28,color:'var(--t3)'}}>Obzvon yozuvlari yo'q</td></tr> :
                    mainRows.map((r)=>{
                      const i = (records || []).findIndex((x) => (x?._rid || '') === (r?._rid || ''));
                      if (i < 0) return null;
                      return (
                      <tr key={r._rid || i}>
                        <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{r.id || '-'}</td>
                        <td style={{minWidth:420}}>
                          <div style={{position:'relative'}}>
                            <input
                              className="input"
                              value={r.customer || ''}
                              placeholder="Mijoz (yozing va tanlang)"
                              onFocus={()=>{ setPickTargetIdx(i); setPickQuery(r.customer || ''); }}
                              onBlur={()=>setTimeout(()=>{ setPickTargetIdx(null); setPickQuery(''); }, 120)}
                              onChange={(e)=>{
                                const v = e.target.value || '';
                                saveRecords(records.map((x,j)=>j===i?{
                                  ...x,
                                  customer:v,
                                  id:x.id || '',
                                  callDate:(x.callDate || (String(v).trim() ? todayIso() : x.callDate)),
                                }:x));
                                setPickTargetIdx(i);
                                setPickQuery(v);
                              }}
                              onKeyDown={(e)=>{
                                if (e.key === 'Escape') {
                                  setPickTargetIdx(null);
                                  setPickQuery('');
                                }
                              }}
                            />
                            {pickTargetIdx===i && suggestions.length>0 && (
                              <div
                                className="card"
                                style={{
                                  position:'absolute',
                                  left:0,
                                  top:'calc(100% + 4px)',
                                  zIndex:40,
                                  width:'max-content',
                                  minWidth:'100%',
                                  maxWidth:'min(960px, calc(100vw - 90px))',
                                  maxHeight:260,
                                  overflow:'auto',
                                  padding:6,
                                }}
                              >
                                {suggestions.map((s) => (
                                  <div
                                    key={`${i}_${s.id}`}
                                    className="nav-i"
                                    style={{padding:'6px 8px'}}
                                    onMouseDown={(ev)=>{
                                      ev.preventDefault();
                                      saveRecords(records.map((x,j)=>j===i?{
                                        ...x,
                                        id:s.id,
                                        customer:s.name,
                                        callDate:(x.callDate || todayIso()),
                                      }:x));
                                      setPickTargetIdx(null);
                                      setPickQuery('');
                                    }}
                                  >
                                    <span style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{s.id}</span>
                                    <span style={{whiteSpace:'normal',wordBreak:'break-word',lineHeight:1.25}}>{s.name}</span>
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>
                        </td>
                        <td><ModernDateInput value={String(r.callDate||'').slice(0,10)} onChange={(e)=>saveRecords(records.map((x,j)=>j===i?{...x,callDate:e.target.value}:x))} /></td>
                        <td>
                          <select
                            className="select"
                            value={r.topic||'Buyurtma olish'}
                            onChange={(e)=>saveRecords(records.map((x,j)=>j===i?{
                              ...x,
                              topic:e.target.value,
                              orderCount:e.target.value === 'Qarzdorlik'
                                ? String(Math.round(Number(customerBalanceById[String(x.id || '').trim()] || 0)))
                                : x.orderCount,
                            }:x))}
                          >
                            <option>Buyurtma olish</option>
                            <option>Qarzdorlik</option>
                            <option>Tara togrlash</option>
                          </select>
                        </td>
                        <td><input className="input" value={r.note||''} onChange={(e)=>saveRecords(records.map((x,j)=>j===i?{...x,note:e.target.value}:x))} /></td>
                        <td><ModernDateInput value={String(r.nextDate||'').slice(0,10)} onChange={(e)=>saveRecords(records.map((x,j)=>j===i?{...x,nextDate:e.target.value}:x))} /></td>
                        <td>
                          <input
                            className="input"
                            value={r.orderCount||''}
                            readOnly={(r.topic || '') === 'Qarzdorlik'}
                            onChange={(e)=>saveRecords(records.map((x,j)=>j===i?{...x,orderCount:e.target.value}:x))}
                            style={{maxWidth:120}}
                          />
                        </td>
                        <td><ModernDateInput value={String(r.orderDate||'').slice(0,10)} onChange={(e)=>saveRecords(records.map((x,j)=>j===i?{...x,orderDate:e.target.value}:x))} /></td>
                        <td>{r.operator || currentUser}</td>
                        <td><button className="btn btn-gh btn-sm" onClick={()=>saveRecords(records.filter((_,j)=>j!==i))}>X</button></td>
                      </tr>
                    )})
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
              <span style={{color:'var(--t3)'}}>{E.find}</span>
              <input placeholder="Mijoz / ID / operator..." value={searchAll} onChange={(e)=>setSearchAll(e.target.value)} />
            </div>
            <div style={{position:'relative'}}>
              <button className="btn btn-gh btn-sm" onClick={()=>setAllUniFilterOpen((v)=>!v)}>
                <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                  <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z"/>
                </svg>
                Filtr ({allFilterCount})
              </button>
              <UniversalFilterPanel
                open={allUniFilterOpen}
                title="Barcha obzvon filtri"
                columns={allFilterColumns}
                rows={historicalAllRows}
                state={allUniFilterState}
                setState={setAllUniFilterState}
                onClose={()=>setAllUniFilterOpen(false)}
                width={620}
              />
            </div>
            <button className="btn btn-gh btn-sm" onClick={() => setSearchAll('')}>
              Qidiruvni tozalash
            </button>
            <button className="btn btn-bl btn-sm" onClick={() => onReloadAll && onReloadAll()}>
              {E.refresh} Yangilash
            </button>
            <span className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>{visibleAllRows.length} / {allList.length} ta yozuv</span>
            {allList.length > 500 && (
              <button className="btn btn-gh btn-sm" onClick={()=>setAllShowMode((m)=>m==='all'?'smart':'all')}>
                {allShowMode==='all' ? "Faqat 500 ta ko'rsatish" : "Hammasini ko'rsatish"}
              </button>
            )}
          </div>
          <div className="card" style={{overflow:'hidden',flex:1}}>
            <div style={{overflow:'auto',maxHeight:'calc(100vh - 230px)'}}>
              <table className="tbl">
                <thead><tr><th>No</th><th>ID</th><th>Kontragent</th><th>Sana</th><th>Mavzu</th><th>Izoh</th><th>Keyingi sana</th><th>Z.soni / summa</th><th>Zakaz sanasi</th><th>Operator</th></tr></thead>
                <tbody>
                  {allList.length===0 ? <tr><td colSpan={10} style={{textAlign:'center',padding:32,color:'var(--t3)'}}>Ma'lumot topilmadi</td></tr> :
                    visibleAllRows.map((r,i)=>(
                      <tr key={i}>
                        <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{r.no || i+1}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{r.customerId || '-'}</td>
                        <td style={{maxWidth:360}}><span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.customer || (r.customerId ? `ID: ${r.customerId}` : '-')}</span></td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.callDate)}</td>
                        <td>{r.topic || '-'}</td>
                        <td style={{maxWidth:300}}><span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.note || '-'}</span></td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.nextDate)}</td>
                        <td>{r.orderCount || '-'}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.orderDate)}</td>
                        <td>{r.operator || '-'}</td>
                      </tr>
                    ))
                  }
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
      {tab==='all_new' && (
        <>
          <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
            <div className="sb" style={{maxWidth:420,flex:1}}>
              <span style={{color:'var(--t3)'}}>{E.find}</span>
              <input placeholder="Mijoz / ID / operator..." value={searchAllNew} onChange={(e)=>setSearchAllNew(e.target.value)} />
            </div>
            <div style={{position:'relative'}}>
              <button className="btn btn-gh btn-sm" onClick={()=>setAllNewUniFilterOpen((v)=>!v)}>
                <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                  <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z"/>
                </svg>
                Filtr ({allNewFilterCount})
              </button>
              <UniversalFilterPanel
                open={allNewUniFilterOpen}
                title="Barcha obzvon yangi filtri"
                columns={allFilterColumns}
                rows={appAllNewRows}
                state={allNewUniFilterState}
                setState={setAllNewUniFilterState}
                onClose={()=>setAllNewUniFilterOpen(false)}
                width={620}
              />
            </div>
            <button className="btn btn-gh btn-sm" onClick={() => setSearchAllNew('')}>
              Qidiruvni tozalash
            </button>
            <button className="btn btn-bl btn-sm" onClick={() => onReloadAllNew && onReloadAllNew()}>
              {E.refresh} Yangilash
            </button>
            {canPublishAllNew && (
              <div style={{display:'flex',gap:8}}>
                <button
                  className="btn btn-gr btn-sm"
                  onClick={async () => {
                    const ok = await Promise.resolve(onPublishAllNew({ manual:true, mode:'append', rows: appAllNewRows }));
                    if (!ok) alert("Google Sheetga yangi ma'lumotlarni joylash muvaffaqiyatsiz bo'ldi");
                  }}
                >
                  Sheetga faqat yangi
                </button>
                <button
                  className="btn btn-gh btn-sm"
                  onClick={async () => {
                    const yes = confirm("Google Sheetdagi 'Barcha_obzvon_yangi' listini to'liq almashtiraymi?");
                    if (!yes) return;
                    const ok = await Promise.resolve(onPublishAllNew({ manual:true, mode:'replace', rows: appAllNewRows }));
                    if (!ok) alert("Google Sheetni to'liq almashtirish muvaffaqiyatsiz bo'ldi");
                  }}
                >
                  Sheetni to'liq almashtirish
                </button>
              </div>
            )}
            <span className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>{visibleAllNewRows.length} / {allNewList.length} ta yozuv</span>
            {allNewList.length > 500 && (
              <button className="btn btn-gh btn-sm" onClick={()=>setAllShowMode((m)=>m==='all'?'smart':'all')}>
                {allShowMode==='all' ? "Faqat 500 ta ko'rsatish" : "Hammasini ko'rsatish"}
              </button>
            )}
          </div>
          <div className="card" style={{overflow:'hidden',flex:1}}>
            <div style={{overflow:'auto',maxHeight:'calc(100vh - 230px)'}}>
              <table className="tbl">
                <thead><tr><th>No</th><th>ID</th><th>Kontragent</th><th>Sana</th><th>Mavzu</th><th>Izoh</th><th>Keyingi sana</th><th>Z.soni / summa</th><th>Zakaz sanasi</th><th>Operator</th>{(canEditAllNew || canDeleteAllNew) && <th>Amal</th>}</tr></thead>
                <tbody>
                  {allNewList.length===0 ? <tr><td colSpan={(canEditAllNew || canDeleteAllNew) ? 11 : 10} style={{textAlign:'center',padding:32,color:'var(--t3)'}}>Yangi obzvon yozuvlari yo'q</td></tr> :
                    visibleAllNewRows.map((r,i)=>{
                      const isEdit = String(allNewEditRid || '') === String(r.rid || '');
                      const d = isEdit ? (allNewEditDraft || r) : r;
                      return (
                      <tr key={r.rid || i}>
                        <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{r.no || i+1}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{r.customerId || '-'}</td>
                        <td style={{maxWidth:360}}>
                          {isEdit ? (
                            <input className="input" value={d.customer || ''} onChange={(e)=>setAllNewEditDraft((p)=>({...(p||{}), customer:e.target.value}))} />
                          ) : (
                            <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.customer || (r.customerId ? `ID: ${r.customerId}` : '-')}</span>
                          )}
                        </td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>
                          {isEdit ? <ModernDateInput value={String(d.callDate || '').slice(0,10)} onChange={(e)=>setAllNewEditDraft((p)=>({...(p||{}), callDate:e.target.value}))} /> : fmtD(r.callDate)}
                        </td>
                        <td>
                          {isEdit ? (
                            <select className="select" value={d.topic || 'Buyurtma olish'} onChange={(e)=>setAllNewEditDraft((p)=>({...(p||{}), topic:e.target.value}))}>
                              <option>Buyurtma olish</option>
                              <option>Qarzdorlik</option>
                              <option>Tara togrlash</option>
                              <option>Baza to'g'irlash</option>
                              <option>Qarzini so'rash</option>
                            </select>
                          ) : (r.topic || '-')}
                        </td>
                        <td style={{maxWidth:300}}>
                          {isEdit ? (
                            <input className="input" value={d.note || ''} onChange={(e)=>setAllNewEditDraft((p)=>({...(p||{}), note:e.target.value}))} />
                          ) : (
                            <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.note || '-'}</span>
                          )}
                        </td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>
                          {isEdit ? <ModernDateInput value={String(d.nextDate || '').slice(0,10)} onChange={(e)=>setAllNewEditDraft((p)=>({...(p||{}), nextDate:e.target.value}))} /> : fmtD(r.nextDate)}
                        </td>
                        <td>
                          {isEdit ? <input className="input" value={d.orderCount || ''} onChange={(e)=>setAllNewEditDraft((p)=>({...(p||{}), orderCount:e.target.value}))} /> : (r.orderCount || '-')}
                        </td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>
                          {isEdit ? <ModernDateInput value={String(d.orderDate || '').slice(0,10)} onChange={(e)=>setAllNewEditDraft((p)=>({...(p||{}), orderDate:e.target.value}))} /> : fmtD(r.orderDate)}
                        </td>
                        <td>{r.operator || '-'}</td>
                        {(canEditAllNew || canDeleteAllNew) && (
                          <td>
                            <div style={{display:'flex',gap:6}}>
                              {canEditAllNew && !isEdit && <button className="btn btn-gh btn-sm" onClick={()=>startAllNewEdit(r)}>Edit</button>}
                              {canEditAllNew && isEdit && <button className="btn btn-bl btn-sm" onClick={saveAllNewEdit}>Saqlash</button>}
                              {canEditAllNew && isEdit && <button className="btn btn-gh btn-sm" onClick={cancelAllNewEdit}>Bekor</button>}
                              {canDeleteAllNew && <button className="btn btn-gh btn-sm" onClick={()=>removeAllNewRow(r)}>O'chir</button>}
                            </div>
                          </td>
                        )}
                      </tr>
                    )})
                  }
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
      {tab==='due' && (
        <>
          <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
            <div style={{position:'relative'}}>
              <button className="btn btn-gh btn-sm" onClick={()=>setDueUniFilterOpen((v)=>!v)}>
                <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                  <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z"/>
                </svg>
                Filtr ({dueFilterCount})
              </button>
              <UniversalFilterPanel
                open={dueUniFilterOpen}
                title="Vaqti kelganlar filtri"
                columns={dueFilterColumns}
                rows={dueCandidates}
                state={dueUniFilterState}
                setState={setDueUniFilterState}
                onClose={()=>setDueUniFilterOpen(false)}
                width={620}
              />
            </div>
            <button className={`btn ${duePickMode?'btn-gr':'btn-gh'} btn-sm`} onClick={()=>setDuePickMode((v)=>!v)}>
              + Obzvonga qo'shish
            </button>
          </div>
          <div className="card" style={{overflow:'hidden'}}>
            <div style={{overflow:'auto',maxHeight:'calc(100vh - 260px)'}}>
              <table className="tbl">
                <thead><tr><th>Mijoz</th><th>ID</th><th>Oxirgi zakazlar</th><th>O'tgan kun</th><th>Qo'ng'iroq me'yori</th><th>Amal</th></tr></thead>
                <tbody>
                  {filteredDueCandidates.length===0 ? <tr><td colSpan={6} style={{textAlign:'center',padding:28,color:'var(--t3)'}}>Vaqti kelgan mijoz yo'q</td></tr> :
                    filteredDueCandidates.map((c,i)=>(
                      <tr
                        key={i}
                        onClick={()=>{
                          if (!duePickMode) return;
                          setDueSelectedIds((p)=>({ ...p, [c.id]: !p[c.id] }));
                        }}
                        style={duePickMode && dueSelectedIds[c.id] ? { background:'rgba(88,166,255,.11)' } : undefined}
                      >
                        <td style={{maxWidth:340}}><span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{c.name}</span></td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{c.id}</td>
                        <td>{c.last3Info}</td>
                        <td>{c.passed} kun</td>
                        <td>{c.shouldIn} kunda</td>
                        <td>
                          {!duePickMode ? (
                            <button className="btn btn-bl btn-sm" onClick={()=>{ addRecord(c,'Buyurtma olish'); setTab('main'); }}>Obzvonga qo'shish</button>
                          ) : (
                            <span className="tag" style={{background:dueSelectedIds[c.id]?'var(--gr2)':'var(--s3)',color:dueSelectedIds[c.id]?'var(--gr)':'var(--t3)'}}>
                              {dueSelectedIds[c.id] ? 'Tanlangan' : 'Tanlash'}
                            </span>
                          )}
                        </td>
                      </tr>
                    ))
                  }
                </tbody>
              </table>
            </div>
          </div>
          {duePickMode && (
            <div style={{display:'flex',justifyContent:'flex-end'}}>
              <button
                className="btn btn-bl btn-sm"
                disabled={dueSelectedCount===0}
                onClick={() => {
                  const selected = filteredDueCandidates.filter((c) => dueSelectedIds[c.id]);
                  if (!selected.length) return;
                  const rows = selected.map((c) => ({
                    id: c.id,
                    customer: c.name,
                    callDate: new Date().toISOString().slice(0,10),
                    topic: 'Buyurtma olish',
                    note: '',
                    nextDate: '',
                    orderCount: '',
                    orderDate: '',
                    operator: currentUser,
                  }));
                  saveRecords([...rows, ...records]);
                  setDueSelectedIds({});
                  setDuePickMode(false);
                  setTab('main');
                }}
              >
                Tanlanganlarni qo'shish ({dueSelectedCount} ta)
              </button>
            </div>
          )}
        </>
      )}
      {tab==='late2m' && (
        <>
          <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
            <div style={{position:'relative'}}>
              <button className="btn btn-gh btn-sm" onClick={()=>setDueUniFilterOpen((v)=>!v)}>
                <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                  <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z"/>
                </svg>
                Filtr ({dueFilterCount})
              </button>
              <UniversalFilterPanel
                open={dueUniFilterOpen}
                title="2 oydan o'tganlar filtri"
                columns={dueFilterColumns}
                rows={staleCandidates}
                state={dueUniFilterState}
                setState={setDueUniFilterState}
                onClose={()=>setDueUniFilterOpen(false)}
                width={620}
              />
            </div>
            <button className={`btn ${latePickMode?'btn-gr':'btn-gh'} btn-sm`} onClick={()=>setLatePickMode((v)=>!v)}>
              + Obzvonga qo'shish
            </button>
          </div>
          <div className="card" style={{overflow:'hidden'}}>
            <div style={{overflow:'auto',maxHeight:'calc(100vh - 260px)'}}>
              <table className="tbl">
                <thead><tr><th>Mijoz</th><th>ID</th><th>Oxirgi zakazlar</th><th>O'tgan kun</th><th>Qo'ng'iroq me'yori</th><th>Amal</th></tr></thead>
                <tbody>
                  {filteredStaleCandidates.length===0 ? <tr><td colSpan={6} style={{textAlign:'center',padding:28,color:'var(--t3)'}}>2 oydan o'tgan mijoz yo'q</td></tr> :
                    filteredStaleCandidates.map((c,i)=>(
                      <tr
                        key={i}
                        onClick={()=>{
                          if (!latePickMode) return;
                          setLateSelectedIds((p)=>({ ...p, [c.id]: !p[c.id] }));
                        }}
                        style={latePickMode && lateSelectedIds[c.id] ? { background:'rgba(88,166,255,.11)' } : undefined}
                      >
                        <td style={{maxWidth:340}}><span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{c.name}</span></td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{c.id}</td>
                        <td>{c.last3Info}</td>
                        <td>{c.passed} kun</td>
                        <td>{c.shouldIn} kunda</td>
                        <td>
                          {!latePickMode ? (
                            <button className="btn btn-bl btn-sm" onClick={()=>{ addRecord(c,'Buyurtma olish'); setTab('main'); }}>Obzvonga qo'shish</button>
                          ) : (
                            <span className="tag" style={{background:lateSelectedIds[c.id]?'var(--gr2)':'var(--s3)',color:lateSelectedIds[c.id]?'var(--gr)':'var(--t3)'}}>
                              {lateSelectedIds[c.id] ? 'Tanlangan' : 'Tanlash'}
                            </span>
                          )}
                        </td>
                      </tr>
                    ))
                  }
                </tbody>
              </table>
            </div>
          </div>
          {latePickMode && (
            <div style={{display:'flex',justifyContent:'flex-end'}}>
              <button
                className="btn btn-bl btn-sm"
                disabled={lateSelectedCount===0}
                onClick={() => {
                  const selected = filteredStaleCandidates.filter((c) => lateSelectedIds[c.id]);
                  if (!selected.length) return;
                  const rows = selected.map((c) => ({
                    id: c.id,
                    customer: c.name,
                    callDate: new Date().toISOString().slice(0,10),
                    topic: 'Buyurtma olish',
                    note: '',
                    nextDate: '',
                    orderCount: '',
                    orderDate: '',
                    operator: currentUser,
                  }));
                  saveRecords([...rows, ...records]);
                  setLateSelectedIds({});
                  setLatePickMode(false);
                  setTab('main');
                }}
              >
                Tanlanganlarni qo'shish ({lateSelectedCount} ta)
              </button>
            </div>
          )}
        </>
      )}
      {tab==='op' && (
        <>
          <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
            <div className="sb" style={{maxWidth:420,flex:1}}>
              <span style={{color:'var(--t3)'}}>{E.find}</span>
              <input placeholder="ID, mijoz, izoh..." value={opSearch} onChange={(e)=>setOpSearch(e.target.value)} />
            </div>
            <div style={{position:'relative'}}>
              <button className="btn btn-gh btn-sm" onClick={()=>setOpUniFilterOpen((v)=>!v)}>
                <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                  <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z"/>
                </svg>
                Filtr ({opUniversalFilterCount})
              </button>
              <UniversalFilterPanel
                open={opUniFilterOpen}
                title="Operator jadvali filtri"
                columns={opFilterColumns}
                rows={operatorTableBaseRows}
                state={opUniFilterState}
                setState={setOpUniFilterState}
                onClose={()=>setOpUniFilterOpen(false)}
                width={620}
              />
            </div>
            <button className={`btn ${opPickMode?'btn-gr':'btn-gh'} btn-sm`} onClick={()=>setOpPickMode((v)=>!v)}>
              + Obzvonga qo'shish
            </button>
          </div>
          <div className="card" style={{overflow:'hidden'}}>
            <div style={{overflow:'auto',maxHeight:'calc(100vh - 260px)'}}>
              <table className="tbl">
                <thead><tr><th>ID</th><th>Mijoz</th><th>Rayon</th><th style={{textAlign:'right'}}>Balans</th><th>Oxirgi zakaz</th><th>Oxirgidan oldingi zakaz</th><th style={{textAlign:'right'}}>Zakaz soni</th><th>Oxirgi qo'ng'iroq</th><th>Keyingi sana</th><th>Operator</th><th>Oxirgi izoh</th></tr></thead>
                <tbody>
                  {operatorTableRows.length===0 ? <tr><td colSpan={11} style={{textAlign:'center',padding:26,color:'var(--t3)'}}>Ma'lumot topilmadi</td></tr> :
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
                        <td style={{maxWidth:140}}><span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.district || '-'}</span></td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)',color:r.balance<0?'var(--rd)':r.balance>0?'var(--gr)':'var(--t3)'}}>{fmt(r.balance)}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.ord1)}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.ord2)}</td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)',fontSize:11,color:'var(--bl)'}}>{r.lastQty ? `${r.lastQty}` : '0'}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.lastCallDate)}</td>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.nextDate)}</td>
                        <td>{r.operator}</td>
                        <td style={{maxWidth:260}}>
                          <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
                            {r.lastNote || '-'}
                          </span>
                        </td>
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

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў DOLJNIKI Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
function DoljnikModal({ row, D, onClose }) {
  const customer = (D.customers || []).find((c) => c.id === row.id);
  const rows = customer ? buildSverka(customer, D.ordersByMId || {}, D.cashByMId || {}) : [];
  return (
    <div className="modal-ov fade" onClick={(e)=>e.target===e.currentTarget&&onClose()}>
      <div className="modal ani" style={{maxWidth:980}}>
        <div className="mhdr">
          <div>
            <div style={{fontWeight:800,fontSize:14}}>{row.name}</div>
            <div style={{fontSize:11,color:'var(--t3)'}}>ID: {row.id}   |   Qarz zakaz: {row.orderNo || '-'}</div>
          </div>
          <button className="btn btn-gh btn-sm" onClick={onClose}>X</button>
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
                      <td style={{maxWidth:280,overflow:'hidden',textOverflow:'ellipsis'}}>{r.produkt || '-'}</td>
                      <td style={{fontFamily:'var(--mono)',textAlign:'right',color:r.summa?'var(--rd)':'var(--t4)'}}>{r.summa?'-'+fmt(r.summa):'-'}</td>
                      <td style={{fontFamily:'var(--mono)',textAlign:'right',color:r.tolov?'var(--gr)':'var(--t4)'}}>{r.tolov?'+'+fmt(r.tolov):'-'}</td>
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
              ID: {row.customerId}   |   Zakaz: {row.orderNo}   |   Olingan: {fmtD(row.purchaseDate)}
            </div>
          </div>
          <div style={{display:'flex',gap:8}}>
            <button className="btn btn-gr btn-sm" onClick={() => customer && exportSverkaExcel(customer, sverkaRows)}>Excel</button>
            <button className="btn btn-gh btn-sm" onClick={onClose}>X</button>
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
            <button className={`tab${tab==='payments'?' on':''}`} onClick={()=>setTab('payments')}>{E.pay} To'lovlar</button>
            <button className={`tab${tab==='settings'?' on':''}`} onClick={()=>setTab('settings')}>{E.settings} Nastroyka</button>
          </div>

          {tab==='payments' && (
            <div className="card" style={{overflow:'auto',maxHeight:'52vh'}}>
              <table className="tbl">
                <thead><tr><th>Sana</th><th>Op No</th><th>Operator</th><th>Kassa</th><th>To'lov</th></tr></thead>
                <tbody>
                  {kulerPayments.length===0 ? (
                    <tr><td colSpan={5} style={{textAlign:'center',padding:24,color:'var(--t3)'}}>To'lovlar topilmadi</td></tr>
                  ) : kulerPayments.map((p, i) => (
                    <tr key={i}>
                      <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(p.sana)}</td>
                      <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{p.opNum||'-'}</td>
                      <td>{p.operator||'-'}</td>
                      <td>{p.kassa||'-'}</td>
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

function Doljniki({ rows, otherRows = [], D, kulerRows, onAddToObzvon, currentUser, company='murodbaxsh' }) {
  const [tab, setTab] = useState('qarz');
  const [search, setSearch] = useState('');
  const [debtFilterOpen, setDebtFilterOpen] = useState(false);
  const [debtFilterState, setDebtFilterState] = useState({});
  const [kulerSearch, setKulerSearch] = useState('');
  const [kulerFilterOpen, setKulerFilterOpen] = useState(false);
  const [kulerFilterState, setKulerFilterState] = useState({});
  const [selected, setSelected] = useState(null);
  const [selectedKuler, setSelectedKuler] = useState(null);
  const [pickMode, setPickMode] = useState(false);
  const [selectedIds, setSelectedIds] = useState({});
  const [kulerMonthsCfg, setKulerMonthsCfg] = useState(() => S.get('aq-kuler-months', {}));
  const allowKulerTab = company !== 'ahmadtea';
  useEffect(() => {
    if (allowKulerTab) return;
    if (tab === 'kuler') setTab('qarz');
  }, [allowKulerTab, tab]);
  const activeDebtRows = useMemo(() => (tab === 'other_qarz' ? (otherRows || []) : (rows || [])), [tab, rows, otherRows]);
  const debtFilterColumns = useMemo(() => ([
    { key:'id', label:'ID', type:'text' },
    { key:'name', label:'Kontragent', type:'text' },
    { key:'category', label:'Kategoriya', type:'text' },
    { key:'debtUZS', label:'Qarz UZS', type:'number' },
    { key:'orderNo', label:'Qarz zakaz', type:'text' },
    { key:'lastOrderProduct', label:'Mahsulot', type:'text' },
    { key:'lastOrderDate', label:'Sana', type:'date' },
    { key:'days', label:'Kun', type:'number' },
    { key:'note', label:'Izoh (T)', type:'text' },
    { key:'agent', label:'Agent', type:'text' },
  ]), []);
  const kulerFilterColumns = useMemo(() => ([
    { key:'customerId', label:'Mijoz ID', type:'text' },
    { key:'customerName', label:'Mijoz', type:'text' },
    { key:'product', label:'Mahsulot', type:'text' },
    { key:'orderNo', label:'Zakaz', type:'text' },
    { key:'purchaseDate', label:'Olingan sana', type:'date' },
    { key:'months', label:'Oylar', type:'number' },
    { key:'principal', label:'Asosiy summa', type:'number' },
    { key:'paid', label:"To'langan", type:'number' },
    { key:'remaining', label:'Qoldiq', type:'number' },
    { key:'monthly', label:"Oylik to'lov", type:'number' },
    { key:'overdueAmount', label:'Muddati kelgan', type:'number' },
  ]), []);
  useEffect(() => {
    setDebtFilterState((prev) => ensureUniversalFilterState(debtFilterColumns, prev));
    setKulerFilterState((prev) => ensureUniversalFilterState(kulerFilterColumns, prev));
  }, [debtFilterColumns, kulerFilterColumns]);

  const list = useMemo(() => {
    let r = activeDebtRows;
    const q = search.toLowerCase().trim();
    if (q) {
      r = r.filter((x) =>
        (x.name || '').toLowerCase().includes(q) ||
        String(x.id || '').toLowerCase().includes(q) ||
        String(x.orderNo || '').toLowerCase().includes(q) ||
        String(x.category || '').toLowerCase().includes(q) ||
        String(x.note || '').toLowerCase().includes(q)
      );
    }
    return applyUniversalFilters(r, debtFilterColumns, debtFilterState);
  }, [activeDebtRows, search, debtFilterColumns, debtFilterState]);

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
  const kulerList = useMemo(() => {
    let r = kulerActive;
    const q = kulerSearch.toLowerCase().trim();
    if (q) {
      r = r.filter((k) =>
        String(k.customerId || '').toLowerCase().includes(q) ||
        String(k.customerName || '').toLowerCase().includes(q) ||
        String(k.orderNo || '').toLowerCase().includes(q) ||
        String(k.product || '').toLowerCase().includes(q)
      );
    }
    return applyUniversalFilters(r, kulerFilterColumns, kulerFilterState);
  }, [kulerActive, kulerSearch, kulerFilterColumns, kulerFilterState]);
  const kulerDebtTotal = useMemo(
    () => kulerActive.reduce((s, k) => s + (Number(k.overdueAmount || 0) > 0 ? Number(k.overdueAmount || 0) : 0), 0),
    [kulerActive]
  );
  const kulerDebtorCount = useMemo(() => {
    const ids = new Set(
      kulerActive
        .filter((k) => Number(k.overdueAmount || 0) > 0)
        .map((k) => String(k.customerId || k.customerName || '').trim())
        .filter(Boolean)
    );
    return ids.size;
  }, [kulerActive]);
  const kulerRemain = kulerActive.reduce((s, k) => s + k.remaining, 0);
  const kulerDue = kulerActive.filter((k) => k.overdueAmount > 0).length;
  const kulerDueAmount = kulerActive.reduce((s, k) => s + (k.overdueAmount > 0 ? k.overdueAmount : 0), 0);
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
  const debtFilterCount = useMemo(() => countUniversalFilters(debtFilterState), [debtFilterState]);
  const kulerFilterCount = useMemo(() => countUniversalFilters(kulerFilterState), [kulerFilterState]);
  const exportDoljniki = () => {
    exportAoaExcel({
      fileName: `Doljniki_${new Date().toISOString().slice(0,10)}.xlsx`,
      sheetName: 'Doljniki',
      headers: ['ID', 'Mijoz', 'Kategoriya', 'Qarz UZS', 'Qarz zakaz', 'Sana', 'Kun', 'Izoh (T)', 'Agent'],
      columnTypes: ['text', 'text', 'text', 'number', 'text', 'date', 'number', 'text', 'text'],
      rows: list.map((r) => [
        r.id,
        r.name,
        r.category || '',
        r.debtUZS || 0,
        r.orderNo || '',
        fmtD(r.lastOrderDate),
        r.days ?? '',
        r.note || '',
        r.agent || '',
      ]),
    });
  };
  const exportKuler = () => {
    exportAoaExcel({
      fileName: `Kuler_nasiya_${new Date().toISOString().slice(0,10)}.xlsx`,
      sheetName: 'KulerNasiya',
      headers: ['Mijoz ID', 'Mijoz', 'Zakaz', 'Mahsulot', 'Olingan sana', 'Oylar', 'Asosiy summa', 'Tolangan', 'Qoldiq', 'Oylik tolov', 'Muddati kelgan'],
      columnTypes: ['text', 'text', 'text', 'text', 'date', 'number', 'number', 'number', 'number', 'number', 'number'],
      rows: kulerList.map((k) => [
        k.customerId,
        k.customerName,
        k.orderNo,
        k.product || 'Kuler',
        fmtD(k.purchaseDate),
        k.months,
        k.principal || 0,
        k.paid || 0,
        k.remaining || 0,
        k.monthly || 0,
        k.overdueAmount || 0,
      ]),
    });
  };
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
      orderCount: String(Math.round(Number(r.debtUZS || 0))),
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
        <button className={`tab${tab==='other_qarz'?' on':''}`} onClick={()=>setTab('other_qarz')}>Boshqa qarzdorli</button>
        {allowKulerTab && <button className={`tab${tab==='kuler'?' on':''}`} onClick={()=>setTab('kuler')}>Kuler Nasiya</button>}
      </div>
      {tab!=='kuler' ? (
        <>
          <div className="g4">
                <StatCard l={tab==='other_qarz' ? "BOSHQA QARZDORLAR" : "QARZDORLAR"} v={list.length+' ta'} s="UZS bo'yicha" c="var(--rd)"/>
                <StatCard l="JAMI QARZ" v={fmt(debtSum)+" so'm"} s="faqat UZS" c="var(--or)"/>
                <StatCard l="15+ KUN QARZDOR" v={d15.length+' ta'} s={fmt(d15Sum)+" so'm"} c="var(--yl)"/>
                <StatCard l="KULER QARZDORLIK" v={fmt(kulerDebtTotal)+" so'm"} s={`${kulerDebtorCount} ta mijoz qarzdor`} c="var(--bl)"/>
              </div>

              <div style={{display:'flex',gap:8,flexWrap:'wrap',alignItems:'center'}}>
                <div className="sb" style={{flex:2,minWidth:220}}>
                  <span style={{color:'var(--t3)'}}>{E.find}</span>
                  <input placeholder="Ism, ID, zakaz bo'yicha..." value={search} onChange={(e)=>setSearch(e.target.value)} />
                </div>
                <div style={{position:'relative'}}>
                  <button className="btn btn-gh btn-sm" onClick={()=>setDebtFilterOpen((v)=>!v)}>
                    <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                      <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z"/>
                    </svg>
                    Filtr ({debtFilterCount})
                  </button>
                  <UniversalFilterPanel
                    open={debtFilterOpen}
                    title={tab==='other_qarz' ? "Boshqa qarzdorli filtri" : "Doljniki filtri"}
                    columns={debtFilterColumns}
                    rows={activeDebtRows}
                    state={debtFilterState}
                    setState={setDebtFilterState}
                    onClose={()=>setDebtFilterOpen(false)}
                    width={620}
                  />
                </div>
                <button className="btn btn-gr btn-sm" onClick={exportDoljniki}>? Excel</button>
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
                            <div style={{fontSize:10.5,color:'var(--t3)'}}>Qarz zakaz: {r.orderNo || '-'} {r.lastOrderProduct ? `  |   ${r.lastOrderProduct}` : ''}</div>
                          </td>
                          <td style={{fontSize:11.5,color:'var(--t2)'}}>{r.category}</td>
                          <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:800,color:'var(--rd)'}}>-{fmt(Math.abs(r.debtUZS))}</td>
                          <td style={{fontFamily:'var(--mono)',fontSize:11}}>{fmtD(r.lastOrderDate)}</td>
                          <td style={{textAlign:'center'}}>
                            {r.days != null ? <span className="tag" style={{background:r.days>30?'var(--rd2)':r.days>15?'var(--yl2)':'var(--s3)',color:r.days>30?'var(--rd)':r.days>15?'var(--yl)':'var(--t3)'}}>{r.days}k</span> : '-'}
                          </td>
                          <td style={{maxWidth:260}}>
                            <span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap',fontSize:11,color:'var(--t3)'}}>{r.note || '-'}</span>
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
          <div className="g5">
            <StatCard l="KULER TO'LOVCHILAR" v={kulerActive.length+' ta'} s="nasiya mijozlar" c="var(--bl)"/>
            <StatCard l="KULER QOLDIQ" v={fmt(kulerRemain)+" so'm"} s="umumiy qoldiq" c="var(--or)"/>
            <StatCard l="MUDDATI KELGAN" v={kulerDue+' ta'} s="to'lov vaqti kelgan" c="var(--rd)"/>
            <StatCard l="TO'LASHI KERAK" v={fmt(kulerDueAmount)+" so'm"} s={`${kulerDue} ta mijoz`} c="var(--yl)"/>
            <StatCard l="KULER QARZDORLAR" v={kulerActive.filter((k)=>k.remaining>0).length+' ta'} s="qoldiq bor mijoz" c="var(--yl)"/>
          </div>
          <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
            <div className="sb" style={{flex:1,minWidth:220,maxWidth:460}}>
              <span style={{color:'var(--t3)'}}>{E.find}</span>
              <input placeholder="Mijoz, ID, zakaz..." value={kulerSearch} onChange={(e)=>setKulerSearch(e.target.value)} />
            </div>
            <div style={{position:'relative'}}>
              <button className="btn btn-gh btn-sm" onClick={()=>setKulerFilterOpen((v)=>!v)}>
                <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                  <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z"/>
                </svg>
                Filtr ({kulerFilterCount})
              </button>
              <UniversalFilterPanel
                open={kulerFilterOpen}
                title="Kuler nasiya filtri"
                columns={kulerFilterColumns}
                rows={kulerActive}
                state={kulerFilterState}
                setState={setKulerFilterState}
                onClose={()=>setKulerFilterOpen(false)}
                width={620}
              />
            </div>
            <button className="btn btn-gr btn-sm" onClick={exportKuler}>? Excel</button>
          </div>
          <div className="card" style={{overflow:'hidden',flex:1}}>
            <div style={{overflow:'auto',maxHeight:'100%'}}>
              <table className="tbl">
                <thead><tr><th>Mijoz</th><th>Zakaz</th><th>Olingan sana</th><th>Oylar</th><th>To'langan</th><th>Qoldiq</th><th>Muddati kelgan</th></tr></thead>
                <tbody>
                  {kulerList.length===0 ? <tr><td colSpan={7} style={{textAlign:'center',padding:30,color:'var(--t3)'}}>Kuler nasiya topilmadi</td></tr> :
                    kulerList.map((k,i)=>(
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

/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў HISOBOTLAR / NAZORAT Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
function Reports({
  D,
  company='murodbaxsh',
  currentUser='Admin',
  currentAccess=null,
  users=[],
  access={},
  planRows=[],
  planOffDays=[],
  obzvonNewRows=[],
  mode='reports',
}) {
  const { rawOrders=[], warehouseTransfers=[] } = D;
  const todayIso = toIsoDate(new Date());
  const currentMonth = normalizeMonthKey(todayIso.slice(0,7));
  const canSeeAll = (currentAccess?.scope || 'all') === 'all';
  const showReport = mode !== 'nazorat';
  const showNazorat = mode !== 'reports';

  const planByMonth = useMemo(() => {
    const m = {};
    (planRows || []).forEach((r) => { if (r?.month) m[r.month] = r; });
    return m;
  }, [planRows]);
  const selectedPlan = planByMonth[currentMonth] || { month: currentMonth, waterPlan:0, coolerPlan:0, workDays:26 };
  const offConfig = useMemo(() => normalizeOffDaysConfig(planOffDays || []), [planOffDays]);
  const offDaysSet = useMemo(() => new Set(offConfig.weekdays || []), [offConfig]);
  const offDatesSet = useMemo(() => new Set(offConfig.dates || []), [offConfig]);

  const visibleOperators = useMemo(() => {
    const src = (users || []).filter((u) => {
      const conf = normalizeAccessConfig(u, access?.[u]);
      if (!(conf.visible?.reports ?? true)) return false;
      if (u === 'Admin') return false;
      if (conf.role === 'admin') return false;
      return true;
    });
    return src.length ? src : [currentUser || 'Admin'];
  }, [users, access, currentUser]);

  const visibleRows = useMemo(() => {
    const base = (obzvonNewRows || []).filter((r) => !isObzvonDeletedRow(r));
    if (canSeeAll) return base;
    return base.filter((r) => String(r.operator || '').trim() === String(currentUser || '').trim());
  }, [obzvonNewRows, canSeeAll, currentUser]);

  const rowsToday = useMemo(
    () => visibleRows.filter((r) => toIsoDate(r.callDate) === todayIso),
    [visibleRows, todayIso]
  );

  const monthWaterOrders = useMemo(() => {
    return (rawOrders || []).filter((o) =>
      isWaterProduct(o.product) &&
      isOrderDoc(o.docType) &&
      isDeliveredStatus(o.status) &&
      monthKey(o.orderDate) === currentMonth
    );
  }, [rawOrders, currentMonth]);
  const monthWaterQty = monthWaterOrders.reduce((s,o)=>s + Math.abs(o.qty || 0), 0);
  const monthWaterSum = monthWaterOrders.reduce((s,o)=>s + (o.currency === 'USD' ? 0 : (o.sum || 0)), 0);

  const todayDate = useMemo(() => {
    const d = toDate(todayIso);
    if (!d) return new Date();
    d.setHours(0, 0, 0, 0);
    return d;
  }, [todayIso]);
  const monthEndDate = useMemo(() => {
    const [y, m] = String(currentMonth || '').split('-').map((x) => Number(x));
    if (!y || !m) return new Date(todayDate.getFullYear(), todayDate.getMonth() + 1, 0);
    return new Date(y, m, 0);
  }, [currentMonth, todayDate]);
  const soldUntilTodayQty = useMemo(() => {
    return (rawOrders || []).filter((o) => {
      if (!isWaterProduct(o.product)) return false;
      if (!isOrderDoc(o.docType)) return false;
      if (isCancelledStatus(o.status)) return false;
      if (monthKey(o.orderDate) !== currentMonth) return false;
      const od = toDate(o.orderDate);
      if (!od) return false;
      od.setHours(0, 0, 0, 0);
      return od <= todayDate;
    }).reduce((s, o) => s + Math.abs(o.qty || 0), 0);
  }, [rawOrders, currentMonth, todayDate]);
  const remainingMonthPlanQty = Math.max(0, Number(selectedPlan.waterPlan || 0) - soldUntilTodayQty);
  const tomorrowDate = useMemo(() => {
    const d = new Date(todayDate);
    d.setDate(d.getDate() + 1);
    return d;
  }, [todayDate]);
  const nextWorkDate = useMemo(
    () => findNextWorkDate(tomorrowDate, monthEndDate, offDaysSet, offDatesSet),
    [tomorrowDate, monthEndDate, offDaysSet, offDatesSet]
  );
  const remainingWorkDays = useMemo(
    () => nextWorkDate ? countWorkDays(nextWorkDate, monthEndDate, offDaysSet, offDatesSet) : 0,
    [nextWorkDate, monthEndDate, offDaysSet, offDatesSet]
  );
  const nextWorkDayPlan = remainingWorkDays > 0 ? Math.ceil(remainingMonthPlanQty / remainingWorkDays) : 0;
  const nextWorkDayPlanPerOperator = Math.ceil(nextWorkDayPlan / Math.max(1, canSeeAll ? visibleOperators.length : 1));
  const currentPlanTarget = canSeeAll ? nextWorkDayPlan : nextWorkDayPlanPerOperator;

  const debtRowsToday = rowsToday.filter((r) => String(r.topic || '').toLowerCase().includes('qarz'));
  const callsToday = rowsToday.length;
  const debtWorkedCount = new Set(
    debtRowsToday.map((r) => String(r.customerId || r.id || r.customer || '').trim()).filter(Boolean)
  ).size;
  const debtWorkedSum = debtRowsToday.reduce((s,r)=>s + Math.abs(toIntFromMixed(r.orderCount || 0)), 0);

  const dailyByOperator = useMemo(() => {
    const ops = new Map();
    rowsToday.forEach((r) => {
      const op = String(r.operator || '-').trim() || '-';
      if (!ops.has(op)) ops.set(op, []);
      ops.get(op).push(r);
    });
    return Array.from(ops.entries()).map(([op, rows]) => {
      const buyRows = rows.filter((x) => String(x.topic || '').toLowerCase().includes('buyurtma'));
      const qarzRows = rows.filter((x) => String(x.topic || '').toLowerCase().includes('qarz'));
      const qty = buyRows.reduce((s,x)=>s + Math.max(0, toIntFromMixed(x.orderCount || 0)), 0);
      const debtSum = qarzRows.reduce((s,x)=>s + Math.abs(toIntFromMixed(x.orderCount || 0)), 0);
      const byOrderDate = {};
      buyRows.forEach((x) => {
        const k = toIsoDate(x.orderDate) || '-';
        byOrderDate[k] = (byOrderDate[k] || 0) + Math.max(0, toIntFromMixed(x.orderCount || 0));
      });
      const dateEntries = Object.entries(byOrderDate)
        .sort((a,b)=>a[0].localeCompare(b[0]))
        .map(([k,v]) => ({ date: k, qty: v }));
      return {
        operator: op,
        calls: rows.length,
        buyurtmaCount: buyRows.length,
        buyurtmaQty: qty,
        debtCount: qarzRows.length,
        debtSum,
        dateEntries,
      };
    }).sort((a,b)=>b.calls-a.calls);
  }, [rowsToday]);

  const dailyRows = useMemo(() => {
    if (canSeeAll) return dailyByOperator;
    return dailyByOperator.filter((x)=>x.operator===currentUser);
  }, [canSeeAll, dailyByOperator, currentUser]);
  const dailyTotals = useMemo(() => ({
    calls: dailyRows.reduce((s, r) => s + (r.calls || 0), 0),
    buyurtmaCount: dailyRows.reduce((s, r) => s + (r.buyurtmaCount || 0), 0),
    buyurtmaQty: dailyRows.reduce((s, r) => s + (r.buyurtmaQty || 0), 0),
    debtCount: dailyRows.reduce((s, r) => s + (r.debtCount || 0), 0),
    debtSum: dailyRows.reduce((s, r) => s + (r.debtSum || 0), 0),
  }), [dailyRows]);

  const mapStorageKey = useMemo(
    () => `aq-driver-warehouse-map-${normalizeCompanyKey(company)}`,
    [company]
  );
  const [driverWarehouseMap, setDriverWarehouseMap] = useState(() => S.get(mapStorageKey, {}));
  useEffect(() => {
    setDriverWarehouseMap(S.get(mapStorageKey, {}));
  }, [mapStorageKey]);
  const [nazoratSection, setNazoratSection] = useState('orders');
  const [nazoratView, setNazoratView] = useState('errors');
  const [nazoratFilterOpen, setNazoratFilterOpen] = useState(false);
  const [nazoratFilterState, setNazoratFilterState] = useState({});

  const controlOrderRows = useMemo(() => {
    let rows = (rawOrders || []).filter((o) => {
      if (!isWaterProduct(o.product)) return false;
      if (!isOrderDoc(o.docType)) return false;
      if (isCancelledStatus(o.status)) return false;
      if (!isMainWarehouseLabel(o.warehouse)) return false;
      const dateKey = toIsoDate(o.orderDate);
      if (!dateKey) return false;
      const driver = String(o.delivPerson || o.agent || '').trim();
      if (!driver) return false;
      return true;
    }).map((o) => ({
      date: toIsoDate(o.orderDate),
      driver: String(o.delivPerson || o.agent || '').trim(),
      qty: Math.abs(toNum(o.qty)),
    }));
    if (!canSeeAll) rows = rows.filter((r) => r.driver === currentUser);
    return rows;
  }, [rawOrders, canSeeAll, currentUser]);

  const transferRows = useMemo(() => {
    return (warehouseTransfers || []).filter((r) => {
      if (!isWaterProduct(r.product)) return false;
      if (!isMainWarehouseLabel(r.toWarehouse)) return false;
      const d = toIsoDate(r.moveDate);
      if (!d) return false;
      const from = String(r.fromWarehouse || '').trim();
      if (!from) return false;
      return true;
    }).map((r) => ({
      date: toIsoDate(r.moveDate),
      fromWarehouse: String(r.fromWarehouse || '').trim(),
      qty: Math.abs(toNum(r.qty)),
    }));
  }, [warehouseTransfers]);

  const { orderQtyByDriverDate, orderDatesByDriver } = useMemo(() => {
    const qtyMap = new Map();
    const datesByDriver = new Map();
    controlOrderRows.forEach((r) => {
      const key = `${r.driver}__${r.date}`;
      qtyMap.set(key, (qtyMap.get(key) || 0) + Math.abs(toNum(r.qty)));
      if (!datesByDriver.has(r.driver)) datesByDriver.set(r.driver, new Set());
      datesByDriver.get(r.driver).add(r.date);
    });
    return { orderQtyByDriverDate: qtyMap, orderDatesByDriver: datesByDriver };
  }, [controlOrderRows]);

  const { transferQtyByWarehouseDate, transferDatesByWarehouse } = useMemo(() => {
    const qtyMap = new Map();
    const datesByWarehouse = new Map();
    transferRows.forEach((r) => {
      const warehouseKey = normalizeWarehouseKey(r.fromWarehouse);
      if (!warehouseKey) return;
      const key = `${warehouseKey}__${r.date}`;
      qtyMap.set(key, (qtyMap.get(key) || 0) + Math.abs(toNum(r.qty)));
      if (!datesByWarehouse.has(warehouseKey)) datesByWarehouse.set(warehouseKey, new Set());
      datesByWarehouse.get(warehouseKey).add(r.date);
    });
    return { transferQtyByWarehouseDate: qtyMap, transferDatesByWarehouse: datesByWarehouse };
  }, [transferRows]);

  const allDrivers = useMemo(() => {
    const set = new Set([
      ...Array.from(orderDatesByDriver.keys()),
      ...Object.keys(driverWarehouseMap || {}),
    ]);
    let rows = Array.from(set).filter(Boolean).sort((a, b) => a.localeCompare(b, 'ru'));
    if (!canSeeAll) rows = rows.filter((x) => x === currentUser);
    return rows;
  }, [orderDatesByDriver, driverWarehouseMap, canSeeAll, currentUser]);

  const dailyControlRows = useMemo(() => {
    const out = [];
    allDrivers.forEach((driver) => {
      const warehouse = String(driverWarehouseMap?.[driver] || '').trim();
      const warehouseKey = normalizeWarehouseKey(warehouse);
      const orderDates = orderDatesByDriver.get(driver) || new Set();
      const transferDates = warehouseKey ? (transferDatesByWarehouse.get(warehouseKey) || new Set()) : new Set();
      const allDates = Array.from(new Set([...orderDates, ...transferDates])).sort((a, b) => a.localeCompare(b));
      allDates.forEach((date) => {
        const orderQty = Number(orderQtyByDriverDate.get(`${driver}__${date}`) || 0);
        const transferQty = warehouseKey ? Number(transferQtyByWarehouseDate.get(`${warehouseKey}__${date}`) || 0) : 0;
        if (!warehouse && !orderQty && !transferQty) return;
        if (warehouse && !orderQty && !transferQty) return;
        out.push({
          date,
          driver,
          warehouse: warehouse || '-',
          orderQty,
          transferQty,
          diff: orderQty - transferQty,
          mappingMissing: !warehouse,
        });
      });
    });
    return out.sort((a, b) => (a.date === b.date ? a.driver.localeCompare(b.driver, 'ru') : a.date.localeCompare(b.date)));
  }, [allDrivers, driverWarehouseMap, orderDatesByDriver, transferDatesByWarehouse, orderQtyByDriverDate, transferQtyByWarehouseDate]);

  const orderControlRows = useMemo(
    () => dailyControlRows.map((r) => ({
      ...r,
      status: r.mappingMissing ? 'Sklad biriktirilmagan' : (r.diff === 0 ? 'OK' : 'Mos emas'),
    })),
    [dailyControlRows]
  );
  const orderMismatchRows = useMemo(
    () => orderControlRows.filter((r) => r.mappingMissing || r.diff !== 0),
    [orderControlRows]
  );
  const returnControlRows = useMemo(() => {
    const scopedRows = (rawOrders || []).filter((o) => {
      if (!isOrderDoc(o.docType) && !isReturnDoc(o.docType)) return false;
      if (isCancelledStatus(o.status)) return false;
      const date = toIsoDate(o.orderDate);
      if (!date) return false;
      const customerId = String(o.mId || '').trim();
      if (!customerId) return false;
      if (!canSeeAll) {
        const person = String(o.delivPerson || o.agent || '').trim();
        if (person !== String(currentUser || '').trim()) return false;
      }
      return true;
    });
    const orderSet = new Set(
      scopedRows
        .filter((o) => isOrderDoc(o.docType))
        .map((o) => `${String(o.mId || '').trim()}__${toIsoDate(o.orderDate)}`)
    );
    return scopedRows
      .filter((o) => isReturnDoc(o.docType))
      .map((o) => {
        const customerId = String(o.mId || '').trim();
        const date = toIsoDate(o.orderDate);
        const hasOrder = orderSet.has(`${customerId}__${date}`);
        return {
          date,
          customerId,
          customer: String(o.contName || '').trim() || `ID ${customerId}`,
          returnNo: String(o.soNum || '').trim(),
          qty: Math.abs(toNum(o.qty)),
          driver: String(o.delivPerson || o.agent || '').trim() || '-',
          note: String(o.note || o.cat || '').trim() || '-',
          hasOrder,
          status: hasOrder ? 'Zakazi bor' : "Zakazi yo'q",
        };
      })
      .sort((a, b) => (a.date === b.date ? a.customer.localeCompare(b.customer, 'ru') : a.date.localeCompare(b.date)));
  }, [rawOrders, canSeeAll, currentUser]);
  const returnMismatchRows = useMemo(
    () => returnControlRows.filter((r) => !r.hasOrder),
    [returnControlRows]
  );
  const nazoratOrderColumns = useMemo(() => ([
    { key:'date', label:'Sana', type:'date' },
    { key:'driver', label:'Dostavchik', type:'text' },
    { key:'warehouse', label:'Sklad', type:'text' },
    { key:'orderQty', label:'Zakaz suv soni', type:'number' },
    { key:'transferQty', label:'Permesheniya (obshiyga)', type:'number' },
    { key:'diff', label:'Farq', type:'number' },
    { key:'status', label:'Holat', type:'text' },
  ]), []);
  const nazoratReturnColumns = useMemo(() => ([
    { key:'date', label:'Sana', type:'date' },
    { key:'customerId', label:'Mijoz ID', type:'text' },
    { key:'customer', label:'Mijoz', type:'text' },
    { key:'returnNo', label:'Vozvrat zakaz', type:'text' },
    { key:'qty', label:'Vozvrat suv soni', type:'number' },
    { key:'driver', label:'Dostavchik', type:'text' },
    { key:'note', label:'Izoh', type:'text' },
    { key:'status', label:'Status', type:'text' },
  ]), []);
  const activeNazoratColumns = nazoratSection === 'orders' ? nazoratOrderColumns : nazoratReturnColumns;
  const activeNazoratBaseRows = useMemo(() => {
    if (nazoratSection === 'orders') {
      return nazoratView === 'errors' ? orderMismatchRows : orderControlRows;
    }
    return nazoratView === 'errors' ? returnMismatchRows : returnControlRows;
  }, [nazoratSection, nazoratView, orderMismatchRows, orderControlRows, returnMismatchRows, returnControlRows]);
  useEffect(() => {
    setNazoratFilterState((prev) => ensureUniversalFilterState(activeNazoratColumns, prev));
  }, [activeNazoratColumns]);
  const nazoratFilteredRows = useMemo(
    () => applyUniversalFilters(activeNazoratBaseRows, activeNazoratColumns, nazoratFilterState),
    [activeNazoratBaseRows, activeNazoratColumns, nazoratFilterState]
  );
  const nazoratFilterCount = useMemo(
    () => countUniversalFilters(nazoratFilterState),
    [nazoratFilterState]
  );
  const orderDisplayTotals = useMemo(() => ({
    orderQty: nazoratFilteredRows.reduce((s, r) => s + Number(r.orderQty || 0), 0),
    transferQty: nazoratFilteredRows.reduce((s, r) => s + Number(r.transferQty || 0), 0),
  }), [nazoratFilteredRows]);

  const exportReports = () => {
    exportAoaExcel({
      fileName: `Hisobot_${todayIso}.xlsx`,
      sheetName: 'Hisobot',
      headers: ['Operator', 'Bugun obzvon', 'Buyurtma soni', 'Suv soni', 'Qarz mijoz', 'Qarz summa'],
      columnTypes: ['text', 'number', 'number', 'number', 'number', 'number'],
      rows: dailyRows.map((r) => [r.operator, r.calls, r.buyurtmaCount, r.buyurtmaQty, r.debtCount, r.debtSum]),
    });
  };

  const exportNazorat = () => {
    if (nazoratSection === 'orders') {
      exportAoaExcel({
        fileName: `Zakaz_nazorati_${todayIso}.xlsx`,
        sheetName: 'Zakaz_nazorati',
        headers: ['Sana', 'Dostavchik', 'Sklad', 'Zakaz suv soni', 'Permesheniya (obshiyga)', 'Farq', 'Holat'],
        columnTypes: ['date', 'text', 'text', 'number', 'number', 'number', 'text'],
        rows: nazoratFilteredRows.map((r) => [
          r.date,
          r.driver,
          r.warehouse,
          r.orderQty,
          r.transferQty,
          r.diff,
          r.status,
        ]),
      });
      return;
    }
    exportAoaExcel({
      fileName: `Vozvrat_nazorati_${todayIso}.xlsx`,
      sheetName: 'Vozvrat_nazorati',
      headers: ['Sana', 'Mijoz ID', 'Mijoz', 'Vozvrat zakaz', 'Vozvrat suv soni', 'Dostavchik', 'Izoh', 'Status'],
      columnTypes: ['date', 'text', 'text', 'text', 'number', 'text', 'text', 'text'],
      rows: nazoratFilteredRows.map((r) => [
        r.date,
        r.customerId,
        r.customer,
        r.returnNo,
        r.qty,
        r.driver,
        r.note,
        r.status,
      ]),
    });
  };

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:14}}>
      <div style={{display:'flex',justifyContent:'flex-end',gap:8}}>
        {showReport && <button className="btn btn-gr btn-sm" onClick={exportReports}>Hisobot Excel</button>}
        {showNazorat && <button className="btn btn-bl btn-sm" onClick={exportNazorat}>Nazorat Excel</button>}
      </div>
      {showReport && (
      <div className="g4">
        <StatCard
          l={`SUV (${currentMonth})  |  ${companyLabelByKey(company)}`}
          v={`${fmt(monthWaterQty)} ta`}
          s={`${fmt(monthWaterSum)} so'm`}
          c="var(--bl)"
        />
        <StatCard
          l="BUGUNGI PLAN QOLDIQ"
          v={`${fmt(nextWorkDayPlan)} ta`}
          s={`Keyingi ish kuni (${nextWorkDate ? toIsoDate(nextWorkDate) : '-'}) rejasi: ${fmt(nextWorkDayPlan)}  |  oy qoldiq: ${fmt(remainingMonthPlanQty)}  |  ish kun: ${fmt(remainingWorkDays)}`}
          c="var(--gr)"
        />
        <StatCard
          l="QARZDORLIK ISHLOVI"
          v={`${debtWorkedCount} ta`}
          s={`${fmt(debtWorkedSum)} so'm`}
          c="var(--or)"
        />
        <StatCard
          l="BUGUNGI OBZVON"
          v={`${callsToday} ta`}
          s={canSeeAll ? 'Hamma operator' : `${currentUser}`}
          c="var(--yl)"
        />
      </div>
      )}

      {showReport && (
      <div className="card" style={{padding:16}}>
        <div style={{fontWeight:700,fontSize:13,marginBottom:10}}>Kunlik natija</div>
        <div style={{overflow:'auto',maxHeight:'58vh'}}>
          <table className="tbl">
            <thead>
              <tr>
                <th>Operator</th>
                <th style={{textAlign:'right'}}>Obzvon</th>
                <th style={{textAlign:'right'}}>Buyurtma</th>
                <th style={{textAlign:'right'}}>Suv soni</th>
                <th style={{textAlign:'right'}}>Qarz mijoz</th>
                <th style={{textAlign:'right'}}>Qarz summa</th>
                <th>Qaysi sanaga zakaz</th>
                <th style={{textAlign:'right'}}>Kunlik plan</th>
              </tr>
            </thead>
            <tbody>
              {dailyRows.length===0 ? (
                <tr><td colSpan={8} style={{textAlign:'center',padding:30,color:'var(--t3)'}}>Bugungi natija yo'q</td></tr>
              ) : (
                <>
                  {dailyRows.map((r, i) => (
                    <Fragment key={`rep_${i}_${r.operator}`}>
                      <tr>
                        <td style={{fontWeight:700}}>{r.operator}</td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)'}}>{fmt(r.calls)}</td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)'}}>{fmt(r.buyurtmaCount)}</td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)',color:'var(--gr)'}}>{fmt(r.buyurtmaQty)}</td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)'}}>{fmt(r.debtCount)}</td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)',color:'var(--rd)'}}>{fmt(r.debtSum)}</td>
                        <td style={{fontSize:11,color:'var(--t2)'}}>-</td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)',color:'var(--yl)'}}>{fmt(currentPlanTarget)}</td>
                      </tr>
                      {(r.dateEntries || []).map((d, j) => (
                        <tr key={`rep_${i}_${j}_${d.date}`} style={{background:'rgba(88,166,255,.05)'}}>
                          <td style={{color:'var(--t3)',fontSize:11}}>sana</td>
                          <td style={{textAlign:'right',color:'var(--t4)'}}>-</td>
                          <td style={{textAlign:'right',color:'var(--t4)'}}>-</td>
                          <td style={{textAlign:'right',fontFamily:'var(--mono)',color:'var(--gr)'}}>{fmt(d.qty)}</td>
                          <td style={{textAlign:'right',color:'var(--t4)'}}>-</td>
                          <td style={{textAlign:'right',color:'var(--t4)'}}>-</td>
                          <td style={{fontFamily:'var(--mono)',fontSize:11}}>{d.date}</td>
                          <td style={{textAlign:'right',color:'var(--t4)'}}>-</td>
                        </tr>
                      ))}
                    </Fragment>
                  ))}
                  <tr style={{background:'var(--s2)'}}>
                    <td style={{fontWeight:800}}>ITOG</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700}}>{fmt(dailyTotals.calls)}</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700}}>{fmt(dailyTotals.buyurtmaCount)}</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700,color:'var(--gr)'}}>{fmt(dailyTotals.buyurtmaQty)}</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700}}>{fmt(dailyTotals.debtCount)}</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700,color:'var(--rd)'}}>{fmt(dailyTotals.debtSum)}</td>
                    <td style={{fontWeight:700,color:'var(--t3)'}}>-</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:700,color:'var(--yl)'}}>{fmt(currentPlanTarget)}</td>
                  </tr>
                </>
              )}
            </tbody>
          </table>
        </div>
      </div>
      )}

      {showNazorat && (
      <div style={{display:'grid',gap:10,minHeight:0}}>
        <div style={{display:'flex',justifyContent:'space-between',gap:8,alignItems:'center',flexWrap:'wrap',marginBottom:10}}>
          <div className="tabs" style={{display:'inline-flex'}}>
            <button className={`tab${nazoratSection==='orders'?' on':''}`} onClick={()=>setNazoratSection('orders')}>Zakaz nazorati</button>
            <button className={`tab${nazoratSection==='returns'?' on':''}`} onClick={()=>setNazoratSection('returns')}>Vozvrat nazorati</button>
          </div>
          <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
            <div style={{position:'relative'}}>
              <button className="btn btn-gh btn-sm" onClick={()=>setNazoratFilterOpen((v)=>!v)}>
                <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                  <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z"/>
                </svg>
                Filtr ({nazoratFilterCount})
              </button>
              <UniversalFilterPanel
                open={nazoratFilterOpen}
                title={nazoratSection === 'orders' ? 'Zakaz nazorati filtri' : 'Vozvrat nazorati filtri'}
                columns={activeNazoratColumns}
                rows={activeNazoratBaseRows}
                state={nazoratFilterState}
                setState={setNazoratFilterState}
                onClose={()=>setNazoratFilterOpen(false)}
                width={620}
              />
            </div>
            <div className="tabs" style={{display:'inline-flex'}}>
              <button className={`tab${nazoratView==='errors'?' on':''}`} onClick={()=>setNazoratView('errors')}>Hatolilar</button>
              <button className={`tab${nazoratView==='all'?' on':''}`} onClick={()=>setNazoratView('all')}>Barchasi</button>
            </div>
          </div>
        </div>

        {nazoratSection === 'orders' && (
          <div style={{fontSize:11,color:'var(--t3)',marginBottom:10}}>
            Faqat "Основной склад"dan urilgan zakazlar hisobga olinadi. Dostavchik-sklad biriktirish: Nastroyka &gt; Biriktirish bo'limida.
          </div>
        )}
        {nazoratSection === 'returns' && (
          <div style={{fontSize:11,color:'var(--t3)',marginBottom:10}}>
            Vozvrat tekshiruvi: shu sana va shu mijoz bo'yicha zakaz bo'lsa "Zakazi bor", bo'lmasa "Zakazi yo'q".
          </div>
        )}

        <div className="card" style={{overflow:'hidden',minHeight:'calc(100vh - 250px)'}}>
          <div style={{overflow:'auto',maxHeight:'calc(100vh - 250px)'}}>
          {nazoratSection === 'orders' ? (
            <table className="tbl">
              <thead>
                <tr>
                  <th>Sana</th>
                  <th>Dostavchik</th>
                  <th>Sklad</th>
                  <th style={{textAlign:'right'}}>Zakaz suv soni</th>
                  <th style={{textAlign:'right'}}>Permesheniya (obshiyga)</th>
                  <th style={{textAlign:'right'}}>Farq</th>
                  <th>Holat</th>
                </tr>
              </thead>
              <tbody>
                {nazoratFilteredRows.length === 0 ? (
                  <tr><td colSpan={7} style={{textAlign:'center',padding:24,color:'var(--t3)'}}>{nazoratView==='errors' ? 'Hatolik topilmadi' : "Ma'lumot topilmadi"}</td></tr>
                ) : (
                  <>
                    {nazoratFilteredRows.map((r, i) => (
                      <tr key={`nord_${i}_${r.driver}_${r.date}`} style={r.status === 'OK' ? undefined : {background:'rgba(248,81,73,.07)'}}>
                        <td style={{fontFamily:'var(--mono)',fontSize:11}}>{r.date || '-'}</td>
                        <td>{r.driver}</td>
                        <td>{r.warehouse}</td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)',color:'var(--gr)'}}>{fmt(r.orderQty)}</td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)',color:'var(--bl)'}}>{fmt(r.transferQty)}</td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)',color:r.diff===0?'var(--gr)':'var(--rd)'}}>{fmt(r.diff)}</td>
                        <td style={{fontSize:11,color:r.status === 'OK' ? 'var(--gr)' : 'var(--rd)'}}>{r.status}</td>
                      </tr>
                    ))}
                    {nazoratView === 'all' && (
                      <tr style={{background:'var(--s2)'}}>
                        <td colSpan={3} style={{fontWeight:800}}>ITOG</td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:800,color:'var(--gr)'}}>{fmt(orderDisplayTotals.orderQty)}</td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:800,color:'var(--bl)'}}>{fmt(orderDisplayTotals.transferQty)}</td>
                        <td style={{textAlign:'right',fontFamily:'var(--mono)',fontWeight:800,color:orderDisplayTotals.orderQty-orderDisplayTotals.transferQty===0?'var(--gr)':'var(--rd)'}}>
                          {fmt(orderDisplayTotals.orderQty - orderDisplayTotals.transferQty)}
                        </td>
                        <td />
                      </tr>
                    )}
                  </>
                )}
              </tbody>
            </table>
          ) : (
            <table className="tbl">
              <thead>
                <tr>
                  <th>Sana</th>
                  <th>Mijoz ID</th>
                  <th>Mijoz</th>
                  <th>Vozvrat zakaz</th>
                  <th style={{textAlign:'right'}}>Vozvrat suv soni</th>
                  <th>Dostavchik</th>
                  <th>Izoh</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {nazoratFilteredRows.length === 0 ? (
                  <tr><td colSpan={8} style={{textAlign:'center',padding:24,color:'var(--t3)'}}>{nazoratView==='errors' ? 'Hatolik topilmadi' : "Ma'lumot topilmadi"}</td></tr>
                ) : nazoratFilteredRows.map((r, i) => (
                  <tr key={`nret_${i}_${r.customerId}_${r.date}`} style={r.hasOrder ? undefined : {background:'rgba(248,81,73,.08)'}}>
                    <td style={{fontFamily:'var(--mono)',fontSize:11}}>{r.date || '-'}</td>
                    <td style={{fontFamily:'var(--mono)',fontSize:11,color:'var(--t3)'}}>{r.customerId}</td>
                    <td>{r.customer}</td>
                    <td style={{fontFamily:'var(--mono)',fontSize:11}}>{r.returnNo || '-'}</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',color:'var(--or)'}}>{fmt(r.qty)}</td>
                    <td>{r.driver}</td>
                    <td style={{maxWidth:280}}><span style={{display:'block',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{r.note || '-'}</span></td>
                    <td style={{fontSize:11,color:r.hasOrder ? 'var(--gr)' : 'var(--rd)'}}>{r.status}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
          </div>
        </div>
      </div>
      )}
    </div>
  );
}

function PlanPage({
  company='murodbaxsh',
  users=[],
  access={},
  currentUser='Admin',
  planRows=[],
  setPlanRows=()=>{},
  planOffDays=[],
  setPlanOffDays=()=>{},
}) {
  const [rows, setRows] = useState(() => normalizePlanRows(planRows));
  useEffect(() => { setRows(normalizePlanRows(planRows)); }, [planRows]);
  const [offDayConfig, setOffDayConfig] = useState(() => normalizeOffDaysConfig(planOffDays));
  useEffect(() => { setOffDayConfig(normalizeOffDaysConfig(planOffDays)); }, [planOffDays]);
  const [showHolidayPicker, setShowHolidayPicker] = useState(false);
  const [holidayDateInput, setHolidayDateInput] = useState('');
  const [holidayMonth, setHolidayMonth] = useState(() => {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  });
  const offDays = offDayConfig.weekdays || [];
  const extraOffDates = offDayConfig.dates || [];

  const operators = useMemo(() => {
    const list = (users || []).filter((u) => u !== 'Admin');
    return list.length ? list : ['Operator'];
  }, [users]);
  const opCount = Math.max(1, operators.length);
  const canEdit = currentUser === 'Admin';

  const updateCell = (month, key, value) => {
    const n = Math.max(0, Math.round(Number(value || 0)));
    setRows((prev) => prev.map((r) => r.month === month ? { ...r, [key]: key === 'workDays' ? Math.max(1, n) : n } : r));
  };
  const toggleExtraOffDay = (dateIso) => {
    const d = String(dateIso || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return;
    setOffDayConfig((prev) => {
      const cfg = normalizeOffDaysConfig(prev);
      if (cfg.dates.includes(d)) {
        return normalizeOffDaysConfig({ ...cfg, dates: cfg.dates.filter((x) => x !== d) });
      }
      return normalizeOffDaysConfig({ ...cfg, dates: [...cfg.dates, d] });
    });
  };
  const addExtraOffDay = () => {
    if (!holidayDateInput) return;
    toggleExtraOffDay(holidayDateInput);
    setHolidayDateInput('');
  };
  const removeExtraOffDay = (dateIso) => {
    setOffDayConfig((prev) => {
      const cfg = normalizeOffDaysConfig(prev);
      return normalizeOffDaysConfig({ ...cfg, dates: cfg.dates.filter((d) => d !== dateIso) });
    });
  };
  const holidayCalendar = useMemo(() => {
    const m = String(holidayMonth || '').match(/^(\d{4})-(\d{2})$/);
    const base = m ? new Date(Number(m[1]), Number(m[2]) - 1, 1) : new Date();
    base.setHours(0, 0, 0, 0);
    const start = new Date(base);
    const mondayIdx = (start.getDay() + 6) % 7;
    start.setDate(start.getDate() - mondayIdx);
    const cells = [];
    for (let i = 0; i < 42; i++) {
      const d = new Date(start);
      d.setDate(start.getDate() + i);
      const iso = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
      cells.push({
        iso,
        day: d.getDate(),
        inMonth: d.getMonth() === base.getMonth() && d.getFullYear() === base.getFullYear(),
        weekday: d.getDay(),
      });
    }
    return {
      monthLabel: base.toLocaleDateString('uz-UZ', { month: 'long', year: 'numeric' }),
      cells,
    };
  }, [holidayMonth]);
  const savePlan = () => {
    const next = normalizePlanRows(rows);
    setRows(next);
    setPlanRows(next);
    S.set('aq-plan-rows', next);
    const normalizedOff = normalizeOffDaysConfig(offDayConfig);
    const storedOffDays = toOffDaysStorage(normalizedOff);
    setOffDayConfig(normalizedOff);
    setPlanOffDays(storedOffDays);
    S.set('aq-plan-off-days', storedOffDays);
  };
  const exportPlan = () => {
    exportAoaExcel({
      fileName: `Plan_${new Date().toISOString().slice(0,10)}.xlsx`,
      sheetName: 'Plan',
      headers: ['Oy', 'Suv plan', 'Kuler plan', 'Ish kuni', 'Kunlik suv', 'Kunlik kuler', 'Operatorga suv', 'Operatorga kuler'],
      columnTypes: ['text', 'number', 'number', 'number', 'number', 'number', 'number', 'number'],
      rows: rows.map((r) => {
        const wd = Math.max(1, Number(r.workDays || 1));
        const waterDay = Math.ceil(Number(r.waterPlan || 0) / wd);
        const coolerDay = Math.ceil(Number(r.coolerPlan || 0) / wd);
        return [r.month, r.waterPlan, r.coolerPlan, r.workDays, waterDay, coolerDay, Math.ceil(waterDay / opCount), Math.ceil(coolerDay / opCount)];
      }),
    });
  };

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:12}}>
      <div style={{display:'flex',justifyContent:'space-between',gap:8,alignItems:'center',flexWrap:'wrap'}}>
        <div className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>
          Operatorlar: {operators.join(', ')} | Kompaniya: {companyLabelByKey(company)}
        </div>
        <div style={{display:'flex',gap:8}}>
          <button className="btn btn-gr btn-sm" onClick={exportPlan}>Excel</button>
          <button className="btn btn-bl btn-sm" onClick={savePlan} disabled={!canEdit}>Saqlash</button>
        </div>
      </div>
      <div className="card" style={{padding:12}}>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:8,marginBottom:8,flexWrap:'wrap'}}>
          <div style={{fontWeight:700,fontSize:12.5}}>Dam kunlari (haftalik)</div>
          <button className="btn btn-gh btn-sm" disabled={!canEdit} onClick={()=>setShowHolidayPicker((v)=>!v)}>
            <svg viewBox="0 0 24 24" width="12" height="12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
              <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
              <line x1="16" y1="2" x2="16" y2="6"></line>
              <line x1="8" y1="2" x2="8" y2="6"></line>
              <line x1="3" y1="10" x2="21" y2="10"></line>
            </svg>
            Kalendar
          </button>
        </div>
        <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
          {WEEKDAY_OPTIONS.map((d) => {
            const on = offDays.includes(d.key);
            return (
              <button
                key={d.key}
                className={`btn btn-sm ${on ? 'btn-bl' : 'btn-gh'}`}
                disabled={!canEdit}
                onClick={() => {
                  setOffDayConfig((prev) => {
                    const cfg = normalizeOffDaysConfig(prev);
                    const set = new Set(cfg.weekdays);
                    if (set.has(d.key)) set.delete(d.key);
                    else set.add(d.key);
                    return normalizeOffDaysConfig({ ...cfg, weekdays: Array.from(set) });
                  });
                }}
              >
                {d.label}
              </button>
            );
          })}
        </div>
        {showHolidayPicker && (
          <div style={{marginTop:10,padding:10,border:'1px dashed var(--b1)',borderRadius:8,display:'grid',gap:8}}>
            <div style={{fontSize:11,color:'var(--t3)'}}>Qo'shimcha dam kunlari (bayram va boshqa kunlar)</div>
            <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:8,flexWrap:'wrap'}}>
              <div style={{display:'flex',alignItems:'center',gap:8}}>
                <button
                  className="btn btn-gh btn-sm"
                  disabled={!canEdit}
                  onClick={() => {
                    setHolidayMonth((prev) => {
                      const m = String(prev || '').match(/^(\d{4})-(\d{2})$/);
                      const d = m ? new Date(Number(m[1]), Number(m[2]) - 2, 1) : new Date();
                      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
                    });
                  }}
                >
                  {'<'}
                </button>
                <div style={{fontWeight:700,fontSize:12.5,textTransform:'capitalize'}}>{holidayCalendar.monthLabel}</div>
                <button
                  className="btn btn-gh btn-sm"
                  disabled={!canEdit}
                  onClick={() => {
                    setHolidayMonth((prev) => {
                      const m = String(prev || '').match(/^(\d{4})-(\d{2})$/);
                      const d = m ? new Date(Number(m[1]), Number(m[2]), 1) : new Date();
                      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
                    });
                  }}
                >
                  {'>'}
                </button>
              </div>
              <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
                <ModernDateInput
                  value={holidayDateInput}
                  disabled={!canEdit}
                  onChange={(e)=>setHolidayDateInput(e.target.value)}
                  style={{maxWidth:180}}
                />
                <button className="btn btn-bl btn-sm" disabled={!canEdit || !holidayDateInput} onClick={addExtraOffDay}>
                  Qo'shish
                </button>
              </div>
            </div>
            <div className="hol-week">
              {['Du','Se','Cho','Pa','Ju','Sha','Yak'].map((x) => <span key={x}>{x}</span>)}
            </div>
            <div className="hol-grid">
              {holidayCalendar.cells.map((c) => {
                const on = extraOffDates.includes(c.iso);
                const weeklyOff = offDays.includes(c.weekday);
                return (
                  <button
                    key={c.iso}
                    type="button"
                    className={`hol-day${on ? ' on' : ''}${c.inMonth ? '' : ' out'}${weeklyOff ? ' weekly' : ''}`}
                    disabled={!canEdit}
                    onClick={() => toggleExtraOffDay(c.iso)}
                    title={c.iso}
                  >
                    {c.day}
                  </button>
                );
              })}
            </div>
            <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
              {extraOffDates.length === 0 ? (
                <span style={{fontSize:11,color:'var(--t3)'}}>Hozircha qo'shimcha dam kuni tanlanmagan</span>
              ) : extraOffDates.map((d) => (
                <span key={d} className="tag" style={{background:'var(--s3)',color:'var(--t2)',display:'inline-flex',alignItems:'center',gap:6}}>
                  {d}
                  <button
                    type="button"
                    disabled={!canEdit}
                    onClick={()=>removeExtraOffDay(d)}
                    style={{border:'none',background:'transparent',color:'inherit',cursor:canEdit?'pointer':'default',padding:0}}
                  >
                    x
                  </button>
                </span>
              ))}
            </div>
          </div>
        )}
        <div style={{fontSize:11,color:'var(--t3)',marginTop:8}}>
          Bugungi plan qoldiq hisobi haftalik dam kunlari va kalendardan qo'shilgan dam kunlariga qarab keyingi ish kuniga taqsimlanadi.
        </div>
      </div>
      <div className="card" style={{overflow:'hidden'}}>
        <div style={{overflow:'auto',maxHeight:'68vh'}}>
          <table className="tbl">
            <thead>
              <tr>
                <th>Oy</th>
                <th style={{textAlign:'right'}}>Suv plan</th>
                <th style={{textAlign:'right'}}>Kuler plan</th>
                <th style={{textAlign:'right'}}>Ish kuni</th>
                <th style={{textAlign:'right'}}>Kunlik suv</th>
                <th style={{textAlign:'right'}}>Kunlik kuler</th>
                <th style={{textAlign:'right'}}>1 operatorga suv</th>
                <th style={{textAlign:'right'}}>1 operatorga kuler</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r, i) => {
                const wd = Math.max(1, Number(r.workDays || 1));
                const waterDay = Math.ceil(Number(r.waterPlan || 0) / wd);
                const coolerDay = Math.ceil(Number(r.coolerPlan || 0) / wd);
                const perOpWater = Math.ceil(waterDay / opCount);
                const perOpCooler = Math.ceil(coolerDay / opCount);
                return (
                  <tr key={r.month || i}>
                    <td style={{fontFamily:'var(--mono)',fontSize:11}}>{r.month}</td>
                    <td style={{textAlign:'right'}}>
                      <input className="input" type="number" value={r.waterPlan} disabled={!canEdit} onChange={(e)=>updateCell(r.month,'waterPlan',e.target.value)} />
                    </td>
                    <td style={{textAlign:'right'}}>
                      <input className="input" type="number" value={r.coolerPlan} disabled={!canEdit} onChange={(e)=>updateCell(r.month,'coolerPlan',e.target.value)} />
                    </td>
                    <td style={{textAlign:'right'}}>
                      <input className="input" type="number" value={r.workDays} disabled={!canEdit} onChange={(e)=>updateCell(r.month,'workDays',e.target.value)} />
                    </td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)'}}>{fmt(waterDay)}</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)'}}>{fmt(coolerDay)}</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',color:'var(--gr)'}}>{fmt(perOpWater)}</td>
                    <td style={{textAlign:'right',fontFamily:'var(--mono)',color:'var(--yl)'}}>{fmt(perOpCooler)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function SettingsPanel({
  users, setUsers, access, setAccess, currentUser, setCurrentUser,
  webhookUrl, setWebhookUrl, userCreds, setUserCreds, onSwitchUser, isAdminSession=false, viewerAccess=null,
  D=null, company='murodbaxsh',
}) {
  const [tab, setTab] = useState('staff');
  const [sel, setSel] = useState(currentUser || users[0] || 'Admin');
  const [staffModalUser, setStaffModalUser] = useState('');
  const [openPermSection, setOpenPermSection] = useState('cust');
  const [editUser, setEditUser] = useState('');
  const [editLogin, setEditLogin] = useState('');
  const [editPassword, setEditPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const pages = ['dash','cust','orders','kassa','obzvon','doljniki','nazorat','reports','plan','refresh','settings'];
  const settingsSections = [
    { key:'settings_staff', label:"Hodimlar ruhsatlari" },
    { key:'settings_app', label:'Ilova sozlamalari' },
    { key:'settings_ui', label:'Interfeys nastroykasi' },
  ];
  const obzvonNewPermSections = [
    { key:'obzvon_new_edit', label:'Barcha Obzvon Yangi: tahrirlash' },
    { key:'obzvon_new_delete', label:"Barcha Obzvon Yangi: o'chirish" },
    { key:'obzvon_new_publish', label:"Barcha Obzvon Yangi: Google Sheetga joylash" },
  ];
  const defaultVisible = { ...DEFAULT_VISIBLE_PAGES };
  const customerTabDefs = [
    { key:'all', label:'Hamma mijozlar' },
    { key:'ahmadtea', label:'Ahmadtea' },
    { key:'murodbaxsh', label:'Murodbaxsh' },
    { key:'activeAll', label:'Aktiv (hammasi)' },
    { key:'activeOwn', label:"Aktiv (o'zimniki)" },
    { key:'inactive', label:'Nofaol' },
    { key:'other', label:'Boshqa mijozlar' },
  ];
  const viewerConf = normalizeAccessConfig(currentUser, viewerAccess || access[currentUser]);
  const uiConf = normalizeAccessConfig(currentUser, access[currentUser]);
  const bindingStorageKey = useMemo(
    () => `aq-driver-warehouse-map-${normalizeCompanyKey(company)}`,
    [company]
  );
  const [driverWarehouseMap, setDriverWarehouseMap] = useState(() => S.get(bindingStorageKey, {}));
  useEffect(() => {
    setDriverWarehouseMap(S.get(bindingStorageKey, {}));
  }, [bindingStorageKey]);
  useEffect(() => {
    S.set(bindingStorageKey, driverWarehouseMap || {});
  }, [bindingStorageKey, driverWarehouseMap]);
  useEffect(() => {
    if (users.includes(sel)) return;
    setSel(users[0] || 'Admin');
  }, [users, sel]);
  const bindingDrivers = useMemo(() => {
    const src = (D?.rawOrders || []).filter((o) =>
      isWaterProduct(o.product) &&
      isOrderDoc(o.docType) &&
      !isCancelledStatus(o.status)
    ).map((o) => String(o.delivPerson || o.agent || '').trim()).filter(Boolean);
    const uniq = Array.from(new Set(src)).sort((a, b) => a.localeCompare(b, 'ru'));
    if (isAdminSession) return uniq;
    return uniq.filter((d) => d === currentUser);
  }, [D, isAdminSession, currentUser]);
  const bindingWarehouses = useMemo(() => {
    const allTransfers = Array.isArray(D?.warehouseTransfers) ? D.warehouseTransfers : [];
    const strict = allTransfers
      .filter((r) => isWaterProduct(r.product) && isMainWarehouseLabel(r.toWarehouse))
      .map((r) => String(r.fromWarehouse || '').trim())
      .filter(Boolean);
    const strictUnique = Array.from(new Set(strict)).sort((a, b) => a.localeCompare(b, 'ru'));
    if (strictUnique.length) return strictUnique;

    const fallback = allTransfers
      .map((r) => String(r.fromWarehouse || '').trim())
      .filter(Boolean);
    return Array.from(new Set(fallback)).sort((a, b) => a.localeCompare(b, 'ru'));
  }, [D]);
  const canEditBinding = !!isAdminSession;
  const allowedSettingsTabs = useMemo(() => {
    const out = [];
    out.push('binding');
    if ((viewerConf.visible?.settings_staff ?? true)) out.push('staff');
    if ((viewerConf.visible?.settings_app ?? true)) out.push('app');
    if ((viewerConf.visible?.settings_ui ?? true)) out.push('ui');
    return out;
  }, [viewerConf]);
  useEffect(() => {
    if (!allowedSettingsTabs.length) return;
    if (!allowedSettingsTabs.includes(tab)) setTab(allowedSettingsTabs[0]);
  }, [allowedSettingsTabs, tab]);

  const addUser = () => {
    if (!isAdminSession) {
      alert("Yangi login qo'shish faqat Admin uchun ochiq");
      return;
    }
    const name = prompt("Yangi login nomi:");
    if (!name || !name.trim()) return;
    const u = name.trim();
    if (users.includes(u)) return;
    const nextUsers = [...users, u];
    setUsers(nextUsers);
    S.set('aq-users', nextUsers);
    setAccess((prev) => {
      const next = { ...prev, [u]: normalizeAccessConfig(u, { scope:'own', activeScope:'own', visible: defaultVisible }) };
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
      const prevConf = normalizeAccessConfig(oldLogin, prev[oldLogin] || { scope:'own', activeScope:'own', visible: defaultVisible });
      const next = { ...prev };
      if (newLogin !== oldLogin) delete next[oldLogin];
      next[newLogin] = normalizeAccessConfig(newLogin, prevConf);
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
  const canManageStaff = Boolean(isAdminSession);
  const openStaffModal = (u) => {
    setSel(u);
    setStaffModalUser(u);
    setOpenPermSection('cust');
  };
  const closeStaffModal = () => {
    setStaffModalUser('');
    closeCredEditor();
  };
  const updateUserAccess = (u, updater) => {
    if (!u) return;
    setAccess((prev) => {
      const current = normalizeAccessConfig(u, prev[u] || access[u]);
      const updated = normalizeAccessConfig(u, updater(current) || current);
      const next = { ...prev, [u]: updated };
      S.set('aq-access', next);
      return next;
    });
  };
  const activeStaffConf = staffModalUser ? normalizeAccessConfig(staffModalUser, access[staffModalUser]) : null;
  const permissionGroupDefs = [
    { key: 'dash', label: 'Dashboard', fallback: true, children: [] },
    {
      key: 'cust',
      label: "Mijozlar",
      fallback: true,
      children: customerTabDefs.map((t) => ({ key: t.key, label: t.label, kind: 'customerTabs', fallback: true })),
    },
    { key: 'orders', label: 'Zakazlar', fallback: true, children: [] },
    { key: 'kassa', label: 'Kassa', fallback: true, children: [] },
    {
      key: 'obzvon',
      label: 'Obzvon',
      fallback: true,
      children: obzvonNewPermSections.map((s) => ({ key: s.key, label: s.label, kind: 'visible', fallback: false })),
    },
    { key: 'doljniki', label: 'Doljniki', fallback: true, children: [] },
    { key: 'nazorat', label: 'Nazorat', fallback: true, children: [] },
    { key: 'reports', label: 'Hisobotlar', fallback: true, children: [] },
    { key: 'plan', label: 'Plan', fallback: true, children: [] },
    { key: 'refresh', label: 'Yangilash', fallback: true, children: [] },
    {
      key: 'settings',
      label: 'Nastroyka',
      fallback: false,
      children: settingsSections.map((s) => ({ key: s.key, label: s.label, kind: 'visible', fallback: true })),
    },
  ];

  return (
    <div className="ani" style={{display:'flex',flexDirection:'column',gap:12}}>
      <div className="tabs" style={{display:'inline-flex'}}>
        <button className={`tab${tab==='binding'?' on':''}`} onClick={()=>setTab('binding')}>Biriktirish</button>
        {(viewerConf.visible?.settings_staff ?? true) && (
          <button className={`tab${tab==='staff'?' on':''}`} onClick={()=>setTab('staff')}>{E.users} Hodimlar ruhsatlari</button>
        )}
        {(viewerConf.visible?.settings_app ?? true) && (
          <button className={`tab${tab==='app'?' on':''}`} onClick={()=>setTab('app')}>{E.settings} Ilova sozlamalari</button>
        )}
        {(viewerConf.visible?.settings_ui ?? true) && (
          <button className={`tab${tab==='ui'?' on':''}`} onClick={()=>setTab('ui')}>Interfeys nastroykasi</button>
        )}
      </div>
      {!allowedSettingsTabs.length && (
        <div className="card" style={{padding:14,color:'var(--t3)'}}>Siz uchun nastroyka bo'limlari yopilgan.</div>
      )}

      {tab==='binding' && (
        <div className="card" style={{padding:14}}>
          <div style={{fontWeight:700,marginBottom:8}}>Dostavchik va sklad biriktirish</div>
          <div style={{fontSize:11,color:'var(--t3)',marginBottom:10}}>
            Kompaniya: <strong style={{color:'var(--t1)'}}>{companyLabelByKey(company)}</strong>. Bu sozlama Nazorat bo'limi solishtiruvida ishlatiladi.
          </div>
          <div style={{overflow:'auto',maxHeight:'56vh'}}>
            <table className="tbl">
              <thead>
                <tr>
                  <th>Dostavchik</th>
                  <th>Sklad</th>
                </tr>
              </thead>
              <tbody>
                {bindingDrivers.length === 0 ? (
                  <tr><td colSpan={2} style={{textAlign:'center',padding:24,color:'var(--t3)'}}>Dostavchik topilmadi</td></tr>
                ) : bindingDrivers.map((driver) => (
                  <tr key={driver}>
                    <td style={{fontWeight:600}}>{driver}</td>
                    <td>
                      <select
                        className="select"
                        style={{minWidth:280}}
                        value={String(driverWarehouseMap?.[driver] || '')}
                        disabled={!canEditBinding}
                        onChange={(e) => {
                          const val = String(e.target.value || '').trim();
                          setDriverWarehouseMap((prev) => ({ ...(prev || {}), [driver]: val }));
                        }}
                      >
                        <option value="">Sklad tanlanmagan</option>
                        {bindingWarehouses.map((w) => <option key={w} value={w}>{w}</option>)}
                      </select>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {!canEditBinding && (
            <div style={{fontSize:11,color:'var(--yl)',marginTop:10}}>
              Biriktirishni o'zgartirish faqat Admin sessiyada ochiq.
            </div>
          )}
        </div>
      )}

      {tab==='staff' && (viewerConf.visible?.settings_staff ?? true) && (
        <div style={{display:'grid',gap:12}}>
          <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
            <span className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>Foydalanuvchilar: {users.length}</span>
            {isAdminSession ? (
              <select className="select" value={currentUser} onChange={(e)=>onSwitchUser ? onSwitchUser(e.target.value) : setCurrentUser(e.target.value)}>
                {users.map((u)=><option key={u}>{u}</option>)}
              </select>
            ) : (
              <span className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>Aktiv: {currentUser}</span>
            )}
            <button className="btn btn-gh btn-sm" onClick={addUser} disabled={!isAdminSession}>+ Login qo'shish</button>
          </div>
          <div className="staff-grid">
            {users.map((u) => {
              const uc = normalizeAccessConfig(u, access[u]);
              const visibleCount = pages.filter((p) => ((uc.visible || {})[p] ?? defaultVisible[p] ?? true)).length;
              const customerCount = customerTabDefs.filter((t) => ((uc.customerTabs || {})[t.key] ?? true)).length;
              return (
                <button
                  key={u}
                  className={`staff-user-card${sel === u ? ' on' : ''}`}
                  onClick={() => openStaffModal(u)}
                  type="button"
                >
                  <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:8}}>
                    <div style={{display:'flex',alignItems:'center',gap:8}}>
                      <div style={{width:30,height:30,borderRadius:15,background:'var(--s3)',display:'flex',alignItems:'center',justifyContent:'center',fontWeight:800,color:'var(--t1)'}}>
                        {u.slice(0,1).toUpperCase()}
                      </div>
                      <div style={{textAlign:'left'}}>
                        <div style={{fontWeight:700,color:'var(--t1)'}}>{u}</div>
                        <div style={{fontSize:10.5,color:'var(--t3)'}}>{uc.role || 'operator'}</div>
                      </div>
                    </div>
                    <span className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>{uc.scope === 'all' ? 'Barcha' : "O'ziga"}</span>
                  </div>
                  <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
                    <span className="tag" style={{background:'var(--s2)',color:'var(--t2)'}}>Bo'lim: {visibleCount}</span>
                    <span className="tag" style={{background:'var(--s2)',color:'var(--t2)'}}>Mijoz tab: {customerCount}</span>
                    <span className="tag" style={{background:'var(--s2)',color:'var(--t2)'}}>{companyLabelByKey(uc.company?.default || 'murodbaxsh')}</span>
                  </div>
                </button>
              );
            })}
          </div>
          {staffModalUser && activeStaffConf && (
            <div className="modal-ov fade" onClick={(e)=>e.target===e.currentTarget&&closeStaffModal()}>
              <div className="modal ani" style={{maxWidth:980}}>
                <div className="mhdr">
                  <div>
                    <div style={{fontWeight:800,fontSize:16}}>{staffModalUser} ruxsatlari</div>
                    <div style={{fontSize:12,color:'var(--t3)',marginTop:2}}>
                      Bo'lim nomini bossangiz ichki ruxsatlar ochiladi
                    </div>
                  </div>
                  <button className="btn btn-gh btn-sm" type="button" onClick={closeStaffModal}>Yopish</button>
                </div>
                <div className="mbdy" style={{display:'grid',gridTemplateColumns:'minmax(260px,320px) 1fr',gap:12}}>
                  <div className="card" style={{padding:12,display:'grid',gap:8,alignSelf:'start'}}>
                    <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:8}}>
                      <span className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>Login va profil</span>
                      <span className="tag" style={{background:'var(--s3)',color:'var(--t3)'}}>{activeStaffConf.scope === 'all' ? 'Barcha' : "O'ziga"}</span>
                    </div>
                    {editUser === staffModalUser ? (
                      <div style={{background:'var(--s3)',border:'1px solid var(--b1)',borderRadius:8,padding:8}}>
                        <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Login</div>
                        <input
                          className="input"
                          value={editLogin}
                          onChange={(e)=>setEditLogin(e.target.value)}
                          disabled={!canManageStaff}
                          style={{padding:'6px 8px',fontSize:12,marginBottom:8}}
                        />
                        <div style={{fontSize:11,color:'var(--t3)',marginBottom:4}}>Parol</div>
                        <div style={{display:'flex',gap:6}}>
                          <input
                            className="input"
                            type={showPassword ? 'text' : 'password'}
                            value={editPassword}
                            onChange={(e)=>setEditPassword(e.target.value)}
                            disabled={!canManageStaff}
                            style={{padding:'6px 8px',fontSize:12,flex:1}}
                          />
                          <button className="btn btn-gh btn-sm" onClick={()=>setShowPassword((v)=>!v)} type="button">
                            {showPassword ? "Yop" : "Ko'r"}
                          </button>
                        </div>
                        <div style={{display:'flex',justifyContent:'flex-end',gap:6,marginTop:8}}>
                          <button className="btn btn-gh btn-sm" onClick={closeCredEditor} type="button">Bekor</button>
                          <button className="btn btn-bl btn-sm" onClick={saveCredEditor} type="button" disabled={!canManageStaff}>Saqlash</button>
                        </div>
                      </div>
                    ) : (
                      <button
                        className="btn btn-gh btn-sm"
                        type="button"
                        disabled={!canManageStaff}
                        onClick={() => openCredEditor(staffModalUser)}
                      >
                        Login/parolni tahrirlash
                      </button>
                    )}
                    <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:8}}>
                      <span style={{fontSize:12,color:'var(--t2)'}}>Faqat o'z mijozlari</span>
                      <button
                        className={`toggle${activeStaffConf.scope==='own'?' on':''}`}
                        disabled={!canManageStaff}
                        onClick={() => updateUserAccess(staffModalUser, (confU) => ({ ...confU, scope: confU.scope === 'own' ? 'all' : 'own' }))}
                      ><span className="knob" /></button>
                    </div>
                    <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:8}}>
                      <span style={{fontSize:12,color:'var(--t2)'}}>Lavozim</span>
                      <select
                        className="select"
                        value={activeStaffConf.role || (staffModalUser==='Admin' ? 'admin' : 'operator')}
                        disabled={!canManageStaff}
                        onChange={(e)=>updateUserAccess(staffModalUser, (confU) => ({ ...confU, role: e.target.value }))}
                      >
                        <option value="operator">Operator</option>
                        <option value="supervisor">Supervisor</option>
                        <option value="admin">Admin</option>
                      </select>
                    </div>
                    <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:8}}>
                      <span style={{fontSize:12,color:'var(--t2)'}}>Kompaniya almashtirish tugmasi</span>
                      <button
                        className={`toggle${(activeStaffConf.company?.canSwitch ?? false)?' on':''}`}
                        disabled={!canManageStaff}
                        onClick={() => updateUserAccess(staffModalUser, (confU) => ({
                          ...confU,
                          company: {
                            ...(confU.company || {}),
                            canSwitch: !(confU.company?.canSwitch ?? false),
                            default: normalizeCompanyKey(confU.company?.default || 'murodbaxsh'),
                          },
                        }))}
                      ><span className="knob" /></button>
                    </div>
                    <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:8}}>
                      <span style={{fontSize:12,color:'var(--t2)'}}>
                        {(activeStaffConf.company?.canSwitch ?? false) ? "Boshlang'ich kompaniya" : "Biriktirilgan kompaniya"}
                      </span>
                      <select
                        className="select"
                        value={normalizeCompanyKey(activeStaffConf.company?.default || 'murodbaxsh')}
                        disabled={!canManageStaff}
                        onChange={(e)=>updateUserAccess(staffModalUser, (confU) => ({
                          ...confU,
                          company: { ...(confU.company || {}), default: normalizeCompanyKey(e.target.value) },
                        }))}
                      >
                        {COMPANY_OPTIONS.map((x)=><option key={x.key} value={x.key}>{x.label}</option>)}
                      </select>
                    </div>
                  </div>
                  <div style={{display:'grid',gap:8,alignContent:'start'}}>
                    {permissionGroupDefs.map((grp) => {
                      const topOn = ((activeStaffConf.visible || {})[grp.key] ?? defaultVisible[grp.key] ?? grp.fallback);
                      const hasChildren = Array.isArray(grp.children) && grp.children.length > 0;
                      const isOpen = openPermSection === grp.key;
                      return (
                        <div key={grp.key} className="perm-section">
                          <div
                            className="perm-head"
                            role={hasChildren ? 'button' : undefined}
                            tabIndex={hasChildren ? 0 : -1}
                            onClick={() => hasChildren && setOpenPermSection((p) => p === grp.key ? '' : grp.key)}
                            onKeyDown={(e) => {
                              if (!hasChildren) return;
                              if (e.key === 'Enter' || e.key === ' ') {
                                e.preventDefault();
                                setOpenPermSection((p) => p === grp.key ? '' : grp.key);
                              }
                            }}
                          >
                            <div className="perm-title">
                              <span style={{fontFamily:'var(--mono)',color:'var(--t3)',minWidth:14}}>
                                {hasChildren ? (isOpen ? 'v' : '>') : ' '}
                              </span>
                              <span>{grp.label}</span>
                            </div>
                            <button
                              className={`toggle${topOn?' on':''}`}
                              disabled={!canManageStaff}
                              onClick={(e) => {
                                e.stopPropagation();
                                updateUserAccess(staffModalUser, (confU) => {
                                  const currentOn = ((confU.visible || {})[grp.key] ?? defaultVisible[grp.key] ?? grp.fallback);
                                  return {
                                    ...confU,
                                    visible: { ...(confU.visible || {}), [grp.key]: !currentOn },
                                  };
                                });
                              }}
                            ><span className="knob" /></button>
                          </div>
                          {hasChildren && isOpen && (
                            <div className="perm-body">
                              {grp.children.map((child) => {
                                const isCustomerTab = child.kind === 'customerTabs';
                                const currentOn = isCustomerTab
                                  ? ((activeStaffConf.customerTabs || {})[child.key] ?? child.fallback ?? true)
                                  : ((activeStaffConf.visible || {})[child.key] ?? child.fallback ?? true);
                                return (
                                  <div key={child.key} className="perm-row">
                                    <span>{child.label}</span>
                                    <button
                                      className={`toggle${currentOn ? ' on' : ''}`}
                                      disabled={!canManageStaff}
                                      onClick={() => {
                                        if (isCustomerTab) {
                                          updateUserAccess(staffModalUser, (confU) => ({
                                            ...confU,
                                            customerTabs: { ...(confU.customerTabs || {}), [child.key]: !((confU.customerTabs || {})[child.key] ?? child.fallback ?? true) },
                                          }));
                                          return;
                                        }
                                        updateUserAccess(staffModalUser, (confU) => ({
                                          ...confU,
                                          visible: { ...(confU.visible || {}), [child.key]: !((confU.visible || {})[child.key] ?? child.fallback ?? true) },
                                        }));
                                      }}
                                    ><span className="knob" /></button>
                                  </div>
                                );
                              })}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>
            </div>
          )}
          </div>
          )}
      {tab==='app' && (viewerConf.visible?.settings_app ?? true) && (
        <div className="card" style={{padding:14}}>
          <div style={{fontWeight:700,marginBottom:8}}>Access API URL</div>
          <input className="input" placeholder="https://your-access-api.workers.dev" value={webhookUrl} onChange={(e)=>setWebhookUrl(e.target.value)} />
          <div style={{fontSize:11,color:'var(--t3)',marginTop:8}}>Bu URL faqat login, parol va ruxsatlarni barcha kompyuterlarda bir xil saqlash uchun ishlatiladi.</div>
          <div style={{marginTop:14,padding:10,border:'1px dashed var(--b1)',borderRadius:8,color:'var(--t3)',fontSize:12}}>Kelajak sozlamalari uchun joy.</div>
        </div>
      )}

      {tab==='ui' && (viewerConf.visible?.settings_ui ?? true) && (
        <div className="card" style={{padding:14}}>
          <div style={{fontWeight:700,marginBottom:10}}>Interfeys nastroykasi</div>
          <div style={{fontSize:12,color:'var(--t3)',marginBottom:10}}>
            Rang rejimi foydalanuvchi bo'yicha saqlanadi. Har kim o'zi uchun alohida tanlaydi.
          </div>
          <div style={{display:'flex',gap:10,flexWrap:'wrap'}}>
            <button
              className={`btn btn-sm ${uiConf.ui?.theme === 'dark' ? 'btn-bl' : 'btn-gh'}`}
              onClick={() => {
                S.set(`aq-ui-theme-${currentUser}`, 'dark');
                setAccess((prev) => {
                  const confU = normalizeAccessConfig(currentUser, prev[currentUser] || uiConf);
                  const next = {
                    ...prev,
                    [currentUser]: normalizeAccessConfig(currentUser, {
                      ...confU,
                      ui: { ...(confU.ui || {}), theme: 'dark' },
                    }),
                  };
                  S.set('aq-access', next);
                  return next;
                });
              }}
            >
              Qora tema
            </button>
            <button
              className={`btn btn-sm ${uiConf.ui?.theme === 'light' ? 'btn-bl' : 'btn-gh'}`}
              onClick={() => {
                S.set(`aq-ui-theme-${currentUser}`, 'light');
                setAccess((prev) => {
                  const confU = normalizeAccessConfig(currentUser, prev[currentUser] || uiConf);
                  const next = {
                    ...prev,
                    [currentUser]: normalizeAccessConfig(currentUser, {
                      ...confU,
                      ui: { ...(confU.ui || {}), theme: 'light' },
                    }),
                  };
                  S.set('aq-access', next);
                  return next;
                });
              }}
            >
              Oq tema
            </button>
          </div>
          <div style={{marginTop:12,fontSize:12,color:'var(--t3)'}}>
            Tanlangan foydalanuvchi: <strong style={{color:'var(--t1)'}}>{currentUser}</strong>  |  Joriy tema: <strong style={{color:'var(--bl)'}}>{(uiConf.ui?.theme || 'dark') === 'light' ? 'Oq' : 'Qora'}</strong>
          </div>
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
              Parol admin tomonidan beriladi
            </div>
          </div>
          {err && <div style={{fontSize:12,color:'var(--rd)'}}>{err}</div>}
          <button className="btn btn-bl" onClick={submit} style={{justifyContent:'center'}}>Kirish</button>
        </div>
      </div>
    </div>
  );
}
/* Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў ROOT APP Р В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ўР В Р вЂ Р Р†Р вЂљРЎС›Р РЋРІР‚в„ў */
const NAV = [
  { id:'dash',    label:'Dashboard',  icon:E.home },
  { id:'cust',    label:'Mijozlar',   icon:E.users, badge:'d' },
  { id:'orders',  label:'Zakazlar',   icon:E.order },
  { id:'kassa',   label:'Kassa',      icon:E.pay },
  { id:'obzvon',  label:'Obzvon',     icon:E.phone, badge:'o' },
  { id:'doljniki',label:'Doljniki',   icon:E.doc, badge:'dz' },
  { id:'nazorat', label:'Nazorat',    icon:E.report },
  { id:'reports', label:'Hisobotlar', icon:E.report },
  { id:'plan',    label:'Plan',       icon:E.plan },
];

export default function App() {
  const [page,setPage]     = useState('dash');
  const [data,setData]     = useState(null);
  const [obzvonRecords,setObzvonRecords] = useState(() => S.get('aq-obzvon-records', []));
  const [obzvonAllRows,setObzvonAllRows] = useState(() => S.get('aq-obzvon-all-rows', []));
  const [obzvonAllNewRows,setObzvonAllNewRows] = useState(() => S.get('aq-obzvon-all-new-rows', []));
  const [planRows, setPlanRows] = useState(() => normalizePlanRows(S.get('aq-plan-rows', [])));
  const [planOffDays, setPlanOffDays] = useState(() => toOffDaysStorage(S.get('aq-plan-off-days', { weekdays:[0], dates:[] })));
  const [users,setUsers] = useState(() => S.get('aq-users', DEFAULT_USERS));
  const [access,setAccess] = useState(() => S.get('aq-access', DEFAULT_ACCESS));
  const [userCreds,setUserCreds] = useState(() => {
    const saved = S.get('aq-user-creds', null);
    if (saved && typeof saved === 'object' && Object.keys(saved).length > 0) return saved;
    return DEFAULT_USER_CREDS;
  });
  const [currentUser,setCurrentUser] = useState(() => S.get('aq-current-user', 'Admin'));
  const [sessionUser,setSessionUser] = useState(() => S.get('aq-session-user', ''));
  const [isLoggedIn,setIsLoggedIn] = useState(() => !!S.get('aq-session-user', ''));
  const [mainSheetUrl, setMainSheetUrl] = useState(() => SHEET_CONFIG.url || '');
  const [obzvonSheetUrl, setObzvonSheetUrl] = useState(() => OBZVON_ALL_SHEET_URL || '');
  // Muhim: barcha qurilmalarda bir xil endpoint ishlashi uchun URL saqlanadi va normallashtiriladi.
  const [obzvonWebhook,setObzvonWebhook] = useState(() =>
    normalizeAccessApiUrl(S.get(ACCESS_SYNC_URL_KEY, OBZVON_WEBHOOK_DEFAULT))
  );
  const accessApiUrl = useMemo(
    () => normalizeAccessApiUrl(obzvonWebhook),
    [obzvonWebhook]
  );
  const [showUp,setUp]     = useState(false);
  const [notif,setNotif]   = useState(null);
  const [side,setSide]     = useState(true);
  const [companyFilter, setCompanyFilter] = useState(() =>
    normalizeCompanyKey(S.get('aq-company-filter', 'murodbaxsh'))
  );
  const [autoLoad,setAutoLoad] = useState({ loading:false, progress:'', error:'' });
  const [remoteAccessLoaded, setRemoteAccessLoaded] = useState(false);
  const [obzvonAllLoaded, setObzvonAllLoaded] = useState(false);
  const skipCloudPushRef = useRef(false);
  const remoteAccessConfigRef = useRef(null);
  const obzvonAllNewRowsRef = useRef(obzvonAllNewRows || []);
  const pendingObzvonNewRowsRef = useRef([]);
  const obzvonExportBusyRef = useRef(false);
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
  const applyRemoteAccessConfig = useCallback((payload) => {
    let cfg = payload;
    if (typeof cfg === 'string') {
      try { cfg = JSON.parse(cfg); } catch { return false; }
    }
    if (!cfg || typeof cfg !== 'object') return false;
    remoteAccessConfigRef.current = cfg;
    const srcUsers = Array.isArray(cfg.users)
      ? cfg.users.map((u) => String(u || '').trim()).filter(Boolean)
      : [];
    const localUsers = Array.isArray(users) ? users.map((u) => String(u || '').trim()).filter(Boolean) : [];
    const nextUsers = [...new Set([...(localUsers || []), ...(srcUsers || [])])];
    if (!nextUsers.length) return false;

    const rawAccess = cfg.access && typeof cfg.access === 'object' ? cfg.access : {};
    const rawCreds = cfg.userCreds && typeof cfg.userCreds === 'object' ? cfg.userCreds : {};
    const remoteUpdatedAt = toDate(cfg.updatedAt)?.getTime() || 0;
    const localUpdatedAt = Number(S.get(ACCESS_UPDATED_AT_KEY, 0)) || 0;
    // Juda eski remote payload lokalni bosib yubormasin.
    if (remoteUpdatedAt && localUpdatedAt && remoteUpdatedAt < localUpdatedAt - 1000) {
      return true;
    }

    const nextAccess = {};
    const nextCreds = {};
    nextUsers.forEach((u) => {
      const localConf = normalizeAccessConfig(u, access[u] || DEFAULT_ACCESS[u]);
      const remoteRaw = (rawAccess[u] && typeof rawAccess[u] === 'object') ? rawAccess[u] : null;
      const mergedConf = remoteRaw
        ? {
            scope: remoteRaw.scope ?? localConf.scope,
            activeScope: remoteRaw.activeScope ?? localConf.activeScope,
            visible: { ...(localConf.visible || {}), ...((remoteRaw.visible && typeof remoteRaw.visible === 'object') ? remoteRaw.visible : {}) },
            customerTabs: { ...(localConf.customerTabs || {}), ...((remoteRaw.customerTabs && typeof remoteRaw.customerTabs === 'object') ? remoteRaw.customerTabs : {}) },
            ui: { ...(localConf.ui || {}), ...((remoteRaw.ui && typeof remoteRaw.ui === 'object') ? remoteRaw.ui : {}) },
            company: {
              ...(localConf.company || {}),
              ...((remoteRaw.company && typeof remoteRaw.company === 'object') ? remoteRaw.company : {}),
            },
          }
        : localConf;
      nextAccess[u] = normalizeAccessConfig(u, mergedConf);
      const fromRemote = String(rawCreds[u] ?? '').trim();
      const fromLocal = String((userCreds || {})[u] ?? '').trim();
      nextCreds[u] = fromRemote || fromLocal || DEFAULT_USER_CREDS[u] || u;
    });

    const sameUsers = JSON.stringify(nextUsers) === JSON.stringify(users);
    const sameAccess = JSON.stringify(nextAccess) === JSON.stringify(access || {});
    const sameCreds = JSON.stringify(nextCreds) === JSON.stringify(userCreds || {});
    if (sameUsers && sameAccess && sameCreds) return true;

    skipCloudPushRef.current = true;
    setUsers(nextUsers);
    setAccess(nextAccess);
    setUserCreds(nextCreds);
    S.set('aq-users', nextUsers);
    S.set('aq-access', nextAccess);
    S.set('aq-user-creds', nextCreds);
    S.set(ACCESS_UPDATED_AT_KEY, remoteUpdatedAt || Date.now());
    return true;
  }, [users, access, userCreds]);

  const loadRemoteAccessConfig = useCallback(async () => {
    if (!accessApiUrl) {
      setRemoteAccessLoaded(true);
      return;
    }
    try {
      const sep = accessApiUrl.includes('?') ? '&' : '?';
      const url = `${accessApiUrl}${sep}action=access_get&_=${Date.now()}`;
      const r = await fetch(url);
      if (r.ok) {
        const j = await r.json();
        if (j?.ok && j?.accessConfig) {
          applyRemoteAccessConfig(j.accessConfig);
          setRemoteAccessLoaded(true);
          return;
        }
      }
    } catch {}
    setRemoteAccessLoaded(true);
  }, [accessApiUrl, applyRemoteAccessConfig]);

  const pushRemoteAccessConfig = useCallback(async (payload) => {
    if (!accessApiUrl) return;
    try {
      const baseRemote = (remoteAccessConfigRef.current && typeof remoteAccessConfigRef.current === 'object')
        ? remoteAccessConfigRef.current
        : {};
      const mergedPayload = {
        ...baseRemote,
        ...(payload && typeof payload === 'object' ? payload : {}),
      };
      const r = await fetch(accessApiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'access_set',
          accessConfig: mergedPayload,
        }),
      });
      const j = await r.json().catch(() => ({}));
      if (j?.ok && (j?.saved || j?.accessConfig || j?.mode === 'access_set')) {
        if (j?.accessConfig && typeof j.accessConfig === 'object') {
          remoteAccessConfigRef.current = j.accessConfig;
        } else {
          remoteAccessConfigRef.current = mergedPayload;
        }
        return;
      }
    } catch {}
  }, [accessApiUrl]);

  const normalizeObzvonAllRow = useCallback((r, i = 0) => ({
    no: sanitizeObzvonCell(r?.no) || String(i + 1),
    customer: sanitizeObzvonCell(r?.customer),
    callDate: sanitizeObzvonCell(r?.callDate),
    topic: sanitizeObzvonCell(r?.topic),
    note: sanitizeObzvonCell(r?.note),
    nextDate: sanitizeObzvonCell(r?.nextDate),
    orderCount: sanitizeObzvonCell(r?.orderCount),
    operator: sanitizeObzvonCell(r?.operator),
    customerId: sanitizeObzvonCell(r?.customerId || r?.id),
    orderDate: sanitizeObzvonCell(r?.orderDate),
    company: normalizeStoredCompanyKey(sanitizeObzvonCell(r?.company)),
  }), []);
  const normalizeObzvonNewRow = useCallback((r, i = 0) => ({
    rid: sanitizeObzvonCell(r?.rid || r?._rid) || buildFallbackObzvonRid(r, i),
    no: sanitizeObzvonCell(r?.no) || String(i + 1),
    customer: sanitizeObzvonCell(r?.customer),
    callDate: sanitizeObzvonCell(r?.callDate),
    topic: sanitizeObzvonCell(r?.topic),
    note: sanitizeObzvonCell(r?.note),
    nextDate: sanitizeObzvonCell(r?.nextDate),
    orderCount: sanitizeObzvonCell(r?.orderCount),
    operator: sanitizeObzvonCell(r?.operator),
    customerId: sanitizeObzvonCell(r?.customerId || r?.id),
    orderDate: sanitizeObzvonCell(r?.orderDate),
    updatedAt: sanitizeObzvonCell(r?.updatedAt) || new Date().toISOString(),
    company: normalizeStoredCompanyKey(sanitizeObzvonCell(r?.company)),
  }), []);
  const obzvonRowUpdatedTs = useCallback((row) => {
    const d = toDate(row?.updatedAt || row?.callDate || '');
    return d ? d.getTime() : 0;
  }, []);
  const mergeObzvonNewRows = useCallback((prevRows = [], incomingRows = []) => {
    const m = new Map((prevRows || []).map((r) => [String(r.rid || ''), normalizeObzvonNewRow(r)]));
    (incomingRows || []).forEach((raw, i) => {
      const r = normalizeObzvonNewRow(raw, i);
      const key = String(r.rid || '');
      if (!key) return;
      const prev = m.get(key);
      if (!prev) {
        m.set(key, r);
        return;
      }
      const prevTs = obzvonRowUpdatedTs(prev);
      const nextTs = obzvonRowUpdatedTs(r);
      if (nextTs >= prevTs) {
        const merged = { ...prev, ...r };
        const prevId = String(prev?.customerId || '').trim();
        const nextId = String(r?.customerId || '').trim();
        if (!nextId && prevId) merged.customerId = prevId;
        if (!String(merged?.customer || '').trim() && String(prev?.customer || '').trim()) {
          merged.customer = prev.customer;
        }
        m.set(key, merged);
      }
    });
    return Array.from(m.values()).sort((a, b) => obzvonRowUpdatedTs(b) - obzvonRowUpdatedTs(a));
  }, [normalizeObzvonNewRow, obzvonRowUpdatedTs]);
  // Global sync uchun ham bir xil qoidani saqlaymiz: mijoz + izoh bo'lsa yozuv tayyor.
  const hasObzvonRowPayload = useCallback((r) => {
    return Boolean(
      String(r?.customer || '').trim() &&
      String(r?.note || '').trim()
    );
  }, []);

  const loadObzvonAllRemote = useCallback(async () => {
    try {
      const r = await fetch(OBZVON_ALL_LOCAL_XLSX, { cache:'no-store' });
      if (!r.ok) return false;
      const arr = await r.arrayBuffer();
      const wb = XLSX.read(arr, { type:'array', cellDates:false, raw:true });
      const firstSheet = wb.SheetNames?.[0];
      if (!firstSheet) return false;
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[firstSheet], { header:1, defval:'', raw:true });
      const body = rows
        .slice(1)
        .map((row, i) => normalizeObzvonAllRow({
          no: row?.[0] ?? (i + 1),                 // A
          customer: pickObzvonCustomerName(row),    // J (prioritet), xato/bo'sh bo'lsa B
          callDate: row?.[2] || '',                // C
          topic: row?.[3] || '',                   // D
          note: row?.[4] || '',                    // E
          nextDate: row?.[5] || '',                // F
          orderCount: row?.[6] || '',              // G
          operator: row?.[7] || '',                // H
          customerId: row?.[8] || '',              // I
          orderDate: '',                            // Bu faylda yo'q
        }, i))
        .filter((x) => x.customer || x.customerId);
      setObzvonAllRows(body || []);
      setObzvonAllLoaded(true);
      S.set('aq-obzvon-all-rows', body || []);
      return true;
    } catch {}
    return false;
  }, [normalizeObzvonAllRow]);

  const appendObzvonAllRemote = useCallback(async (rows = []) => {
    const normalizedRows = (rows || [])
      .map((row, i) => normalizeObzvonAllRow(row, i))
      .filter((r) => (r.customerId || r.customer) && hasObzvonRowPayload(r));
    if (!normalizedRows.length) return;
    setObzvonAllRows((prev) => {
      const next = [...normalizedRows, ...(prev || [])];
      S.set('aq-obzvon-all-rows', next);
      return next;
    });
  }, [normalizeObzvonAllRow, hasObzvonRowPayload]);

  const pullRemoteObzvonNewRows = useCallback(async () => {
    if (!accessApiUrl) return false;
    const applyIncomingRows = (rows = []) => {
      const incoming = (rows || [])
        .map((x, i) => normalizeObzvonNewRow(x, i))
        .filter((x) => (x.customerId || x.customer) && hasObzvonRowPayload(x));
      setObzvonAllNewRows((prev) => {
        const next = mergeObzvonNewRows(prev || [], incoming);
        S.set('aq-obzvon-all-new-rows', next);
        return next;
      });
      return true;
    };
    try {
      const sep = accessApiUrl.includes('?') ? '&' : '?';
      const url = `${accessApiUrl}${sep}action=obzvon_new_get&_=${Date.now()}`;
      const r = await fetch(url, { cache:'no-store' });
      if (r.ok) {
        const j = await r.json().catch(() => ({}));
        if (j?.ok && Array.isArray(j?.rows)) return applyIncomingRows(j.rows);
      }
    } catch {}
    // Fallback: obzvon_new_* actionlari yo'q bo'lsa access_get ichidagi obzvonNewRows dan olamiz.
    try {
      const sep = accessApiUrl.includes('?') ? '&' : '?';
      const url = `${accessApiUrl}${sep}action=access_get&_=${Date.now()}`;
      const r = await fetch(url, { cache:'no-store' });
      if (!r.ok) return false;
      const j = await r.json().catch(() => ({}));
      const cfg = (j?.ok && j?.accessConfig && typeof j.accessConfig === 'object') ? j.accessConfig : null;
      if (!cfg) return false;
      remoteAccessConfigRef.current = cfg;
      const rows = Array.isArray(cfg.obzvonNewRows) ? cfg.obzvonNewRows : [];
      return applyIncomingRows(rows);
    } catch {}
    return false;
  }, [accessApiUrl, normalizeObzvonNewRow, hasObzvonRowPayload, mergeObzvonNewRows]);

  const pushRemoteObzvonNewRows = useCallback(async (rows = []) => {
    if (!accessApiUrl) return false;
    const payloadRows = (rows || [])
      .map((x, i) => normalizeObzvonNewRow(x, i))
      .filter((x) => (x.customerId || x.customer) && hasObzvonRowPayload(x));
    if (!payloadRows.length) return false;
    try {
      const r = await fetch(accessApiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'obzvon_new_upsert',
          rows: payloadRows,
          by: sessionUser || currentUser || 'unknown',
        }),
      });
      const j = await r.json().catch(() => ({}));
      if (j?.ok) return true;
    } catch {}
    // Fallback: access_set ichiga obzvonNewRows ni yozib, barcha qurilmalarda bitta umumiy holatga keltiramiz.
    // MUHIM: avval remote holatni aniq o'qib olmasdan yozmaymiz (aks holda boshqa operator yozuvlari bosilib ketishi mumkin).
    try {
      const sep = accessApiUrl.includes('?') ? '&' : '?';
      const getUrl = `${accessApiUrl}${sep}action=access_get&_=${Date.now()}`;
      const getResp = await fetch(getUrl, { cache:'no-store' });
      if (!getResp.ok) return false;
      const getJson = await getResp.json().catch(() => ({}));
      const remoteCfg = (getJson?.ok && getJson?.accessConfig && typeof getJson.accessConfig === 'object')
        ? getJson.accessConfig
        : null;
      if (!remoteCfg) return false;
      const remoteRows = Array.isArray(remoteCfg.obzvonNewRows) ? remoteCfg.obzvonNewRows : [];
      const localKnownRows = Array.isArray(obzvonAllNewRowsRef.current) ? obzvonAllNewRowsRef.current : [];
      const mergedRows = mergeObzvonNewRows(
        mergeObzvonNewRows(remoteRows, localKnownRows),
        payloadRows
      );
      const fallbackPayload = {
        ...remoteCfg,
        users: Array.isArray(remoteCfg.users) ? remoteCfg.users : users,
        access: (remoteCfg.access && typeof remoteCfg.access === 'object') ? remoteCfg.access : access,
        userCreds: (remoteCfg.userCreds && typeof remoteCfg.userCreds === 'object') ? remoteCfg.userCreds : userCreds,
        obzvonNewRows: mergedRows,
        updatedAt: new Date().toISOString(),
        updatedBy: sessionUser || currentUser || 'unknown',
      };
      const setResp = await fetch(accessApiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'access_set',
          accessConfig: fallbackPayload,
        }),
      });
      const setJson = await setResp.json().catch(() => ({}));
      if (setJson?.ok) {
        remoteAccessConfigRef.current = (setJson?.accessConfig && typeof setJson.accessConfig === 'object')
          ? setJson.accessConfig
          : fallbackPayload;
        return true;
      }
    } catch {}
    return false;
  }, [
    accessApiUrl,
    normalizeObzvonNewRow,
    hasObzvonRowPayload,
    sessionUser,
    currentUser,
    mergeObzvonNewRows,
    users,
    access,
    userCreds,
  ]);

  const flushPendingObzvonNewRows = useCallback(async () => {
    const queue = Array.isArray(pendingObzvonNewRowsRef.current) ? pendingObzvonNewRowsRef.current : [];
    if (!queue.length) return true;
    const ok = await pushRemoteObzvonNewRows(queue);
    if (ok) {
      pendingObzvonNewRowsRef.current = [];
      setTimeout(() => { pullRemoteObzvonNewRows(); }, 350);
    }
    return ok;
  }, [pushRemoteObzvonNewRows, pullRemoteObzvonNewRows]);

  const upsertObzvonAllNewRows = useCallback((rows = []) => {
    const normalizedRows = (rows || [])
      .map((row, i) => normalizeObzvonNewRow(row, i))
      .filter((r) => (r.customerId || r.customer) && hasObzvonRowPayload(r));
    if (!normalizedRows.length) return;
    setObzvonAllNewRows((prev) => {
      const next = mergeObzvonNewRows(prev || [], normalizedRows);
      S.set('aq-obzvon-all-new-rows', next);
      return next;
    });
    pendingObzvonNewRowsRef.current = mergeObzvonNewRows(
      pendingObzvonNewRowsRef.current || [],
      normalizedRows
    );
    Promise.resolve(flushPendingObzvonNewRows());
  }, [normalizeObzvonNewRow, hasObzvonRowPayload, mergeObzvonNewRows, flushPendingObzvonNewRows]);
  useEffect(() => {
    if (!Array.isArray(obzvonAllRows) || !obzvonAllRows.length) return;
    const cleaned = obzvonAllRows
      .map((r, i) => normalizeObzvonAllRow(r, i))
      .filter((r) => r.customer || r.customerId);
    if (JSON.stringify(cleaned) !== JSON.stringify(obzvonAllRows)) {
      setObzvonAllRows(cleaned);
    }
  }, [obzvonAllRows, normalizeObzvonAllRow]);
  useEffect(() => {
    if (!Array.isArray(obzvonAllNewRows) || !obzvonAllNewRows.length) return;
    let changed = false;
    const next = obzvonAllNewRows.map((r, i) => {
      const rid = String(r?.rid || r?._rid || '').trim();
      if (rid) return r;
      changed = true;
      const newRid = buildFallbackObzvonRid(r, i) || createRowRid();
      return {
        ...r,
        rid: newRid,
        _rid: newRid,
        updatedAt: String(r?.updatedAt || '').trim() || new Date().toISOString(),
      };
    });
    if (!changed) return;
    setObzvonAllNewRows(next);
    S.set('aq-obzvon-all-new-rows', next);
  }, [obzvonAllNewRows]);

  const notify = (msg, type='ok') => {
    setNotif({ msg, type });
    setTimeout(() => setNotif(null), 5000);
  };
  const publishObzvonNewRowsToGoogleSheet = useCallback(async ({ manual=false, mode='append', rows: externalRows=null } = {}) => {
    if (obzvonExportBusyRef.current) return false;
    obzvonExportBusyRef.current = true;
    try {
      const exportedMap = S.get(OBZVON_NEW_EXPORTED_MAP_KEY, {}) || {};
      const stableIdMap = S.get(OBZVON_NEW_STABLE_ID_MAP_KEY, {}) || {};
      let stableIdLast = Number(S.get(OBZVON_NEW_STABLE_ID_LAST_KEY, 0)) || 0;
      const ensureStableId = (rid) => {
        const key = String(rid || '').trim();
        if (!key) return 0;
        const ex = Number(stableIdMap[key] || 0);
        if (ex > 0) return ex;
        stableIdLast += 1;
        stableIdMap[key] = stableIdLast;
        return stableIdLast;
      };
      const rows = (Array.isArray(externalRows) ? externalRows : (obzvonAllNewRowsRef.current || []))
        .map((x, i) => normalizeObzvonNewRow(x, i))
        .filter((x) => (x.customerId || x.customer || x.note || x.orderCount))
        .filter((x) => !isObzvonDeletedRow(x));
      const byRid = new Map();
      rows.forEach((r, i) => {
        const rid = String(r.rid || r._rid || '').trim();
        if (!rid) return;
        const prev = byRid.get(rid);
        if (!prev) {
          byRid.set(rid, r);
          return;
        }
        const prevTs = toDate(prev.updatedAt || prev.callDate)?.getTime() || 0;
        const nextTs = toDate(r.updatedAt || r.callDate)?.getTime() || 0;
        if (nextTs >= prevTs) byRid.set(rid, r);
      });
      const dedupedRows = Array.from(byRid.values());
      if (!dedupedRows.length) {
        if (manual) notify("Google Sheetga yuboriladigan ma'lumot yo'q", 'ok');
        return true;
      }

      const rowsWithStableId = dedupedRows.map((r) => {
        const rid = String(r.rid || r._rid || '').trim();
        return {
          ...r,
          stableId: ensureStableId(rid),
        };
      });
      const urlsToTry = Array.from(new Set([OBZVON_EXPORT_APPS_SCRIPT_URL, accessApiUrl].filter(Boolean)));
      if (!urlsToTry.length) {
        if (manual) notify("Eksport URL topilmadi", 'err');
        return false;
      }
      const parseResponse = (text) => {
        const t = String(text || '').trim();
        if (!t) return {};
        try { return JSON.parse(t); } catch (_) { return { raw: t }; }
      };
      const fetchSheetLastInfo = async () => {
        for (const url of urlsToTry) {
          try {
            const isAppsScript = /script\.google\.com/i.test(String(url || ''));
            const reqInit = isAppsScript
              ? {
                  method: 'POST',
                  mode: 'cors',
                  cache: 'no-store',
                  headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
                  body: JSON.stringify({
                    action: 'obzvon_new_sheet_last_id',
                    sheetId: OBZVON_NEW_EXPORT_SHEET_ID,
                    sheetName: OBZVON_NEW_EXPORT_SHEET_NAME,
                  }),
                }
              : {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({
                    action: 'obzvon_new_sheet_last_id',
                    sheetId: OBZVON_NEW_EXPORT_SHEET_ID,
                    sheetName: OBZVON_NEW_EXPORT_SHEET_NAME,
                  }),
                };
            const resp = await fetch(url, reqInit);
            if (!resp.ok) continue;
            const rawText = await resp.text().catch(() => '');
            const js = parseResponse(rawText);
            if (js?.ok) {
              return {
                lastRid: String(js?.lastRid || '').trim(),
                lastStableId: Number(js?.lastStableId || 0) || 0,
              };
            }
          } catch {}
        }
        return { lastRid: '', lastStableId: 0 };
      };
      let rowsToSend = rowsWithStableId;
      if (mode !== 'replace') {
        const lastInfo = await fetchSheetLastInfo();
        const lastRid = String(lastInfo.lastRid || '').trim();
        if (lastRid) {
          const idx = rowsWithStableId.findIndex((r) => String(r.rid || r._rid || '').trim() === lastRid);
          if (idx > 0) rowsToSend = rowsWithStableId.slice(0, idx);
          else if (idx === 0) rowsToSend = [];
        } else {
          rowsToSend = rowsWithStableId.filter((r) => {
            const rid = String(r.rid || r._rid || '').trim();
            const ts = String(r.updatedAt || '').trim() || '0';
            return String(exportedMap[rid] || '') !== ts;
          });
        }
      }
      if (!rowsToSend.length) {
        if (manual) notify("Yangi yuboriladigan ma'lumot yo'q", 'ok');
        S.set(OBZVON_NEW_STABLE_ID_MAP_KEY, stableIdMap);
        S.set(OBZVON_NEW_STABLE_ID_LAST_KEY, stableIdLast);
        return true;
      }
      const payloadRows = rowsToSend.map((r) => ({
        // Google Sheetda A ustun (No) sifatida ketma-ket SyncID saqlanadi
        no: String(r.stableId || ensureStableId(r.rid || r._rid)),
        customerId: String(r.customerId || r.id || '').trim(),
        customer: String(r.customer || '').trim(),
        callDate: String(r.callDate || '').trim(),
        topic: String(r.topic || '').trim(),
        note: String(r.note || '').trim(),
        nextDate: String(r.nextDate || '').trim(),
        orderCount: String(r.orderCount || '').trim(),
        orderDate: String(r.orderDate || '').trim(),
        operator: String(r.operator || '').trim(),
        company: String(r.company || '').trim(),
        rid: String(r.rid || r._rid || '').trim(),
        updatedAt: String(r.updatedAt || '').trim(),
      }));

      const withHeaders = mode === 'replace' ? true : !S.get(OBZVON_NEW_HEADERS_WRITTEN_KEY, false);
      const postBody = {
        sheetId: OBZVON_NEW_EXPORT_SHEET_ID,
        sheetName: OBZVON_NEW_EXPORT_SHEET_NAME,
        sheetGid: OBZVON_NEW_EXPORT_SHEET_GID,
        gid: OBZVON_NEW_EXPORT_SHEET_GID,
        sheetUrl: OBZVON_NEW_EXPORT_SHEET_URL,
        withHeaders,
        headers: ['No', 'ID', 'Mijoz', 'Sana', 'Mavzu', 'Izoh', 'Keyingi sana', 'Z.soni / summa', 'Zakaz sanasi', 'Operator', 'Kompaniya', 'RID', 'UpdatedAt'],
        rows: payloadRows,
        by: sessionUser || currentUser || 'unknown',
      };
      let ok = false;
      const actionsToTry = mode === 'replace'
        ? ['obzvon_new_sheet_replace']
        : ['obzvon_new_sheet_append', 'obzvon_new_export'];
      let lastErr = '';
      for (const url of urlsToTry) {
        if (ok) break;
        const isAppsScript = /script\.google\.com/i.test(String(url || ''));
        for (const action of actionsToTry) {
          if (ok) break;
          const actionBody = { ...postBody, action };
          let actionOk = false;
          try {
            const reqInit = isAppsScript
              ? {
                  method: 'POST',
                  mode: 'cors',
                  cache: 'no-store',
                  headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
                  body: JSON.stringify(actionBody),
                }
              : {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify(actionBody),
                };
            const resp = await fetch(url, {
              ...reqInit,
            });
            const rawText = await resp.text().catch(() => '');
            const js = parseResponse(rawText);
            if (resp.ok && (js?.ok || String(js?.status || '').toLowerCase() === 'ok')) {
              actionOk = true;
            } else {
              lastErr = js?.error || js?.message || `HTTP ${resp.status} (${action})`;
            }
          } catch (e) {
            lastErr = String(e?.message || e || 'unknown');
          }
          if (actionOk) {
            ok = true;
            break;
          }
        }
      }
      if (!ok) {
        if (manual) {
          notify(
            `Google Sheetga joylashda xato bo'ldi${lastErr ? `: ${lastErr}` : ''}`,
            'err'
          );
        }
        return false;
      }

      const nextMap = { ...(exportedMap || {}) };
      rowsToSend.forEach((r) => {
        const rid = String(r.rid || r._rid || '').trim();
        if (!rid) return;
        nextMap[rid] = String(r.updatedAt || '').trim() || new Date().toISOString();
      });
      S.set(OBZVON_NEW_EXPORTED_MAP_KEY, nextMap);
      S.set(OBZVON_NEW_STABLE_ID_MAP_KEY, stableIdMap);
      S.set(OBZVON_NEW_STABLE_ID_LAST_KEY, stableIdLast);
      if (withHeaders) S.set(OBZVON_NEW_HEADERS_WRITTEN_KEY, true);
      if (manual) {
        if (mode === 'replace') notify(`Google Sheet to'liq yangilandi (${rowsToSend.length} ta)`, 'ok');
        else notify(`${rowsToSend.length} ta yangi yozuv Google Sheetga joylandi`, 'ok');
      }
      return true;
    } finally {
      obzvonExportBusyRef.current = false;
    }
  }, [accessApiUrl, normalizeObzvonNewRow, sessionUser, currentUser]);
  useEffect(() => { S.set('aq-current-user', currentUser); }, [currentUser]);
  useEffect(() => { S.set('aq-session-user', sessionUser || ''); }, [sessionUser]);
  useEffect(() => { S.set('aq-company-filter', companyFilter); }, [companyFilter]);
  useEffect(() => {
    const normalized = normalizeAccessApiUrl(obzvonWebhook);
    if (normalized !== String(obzvonWebhook || '').trim()) {
      setObzvonWebhook(normalized);
      return;
    }
    S.set(ACCESS_SYNC_URL_KEY, normalized || '');
    S.set('aq-obzvon-webhook', normalized || '');
  }, [obzvonWebhook]);
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
  useEffect(() => { S.set('aq-obzvon-all-rows', obzvonAllRows || []); }, [obzvonAllRows]);
  useEffect(() => { S.set('aq-obzvon-all-new-rows', obzvonAllNewRows || []); }, [obzvonAllNewRows]);
  useEffect(() => {
    obzvonAllNewRowsRef.current = obzvonAllNewRows || [];
  }, [obzvonAllNewRows]);
  useEffect(() => { S.set('aq-plan-rows', normalizePlanRows(planRows || [])); }, [planRows]);
  useEffect(() => { S.set('aq-plan-off-days', toOffDaysStorage(planOffDays || {})); }, [planOffDays]);
  useEffect(() => { S.set('aq-user-creds', userCreds || {}); }, [userCreds]);
  useEffect(() => {
    if (!Array.isArray(obzvonRecords) || !obzvonRecords.length) return;
    let changed = false;
    const next = obzvonRecords.map((r) => {
      if (r?._rid) return r;
      changed = true;
      return { ...r, _rid: createRowRid() };
    });
    if (changed) setObzvonRecords(next);
  }, [obzvonRecords]);
  useEffect(() => {
    if (!isLoggedIn) return;
    setRemoteAccessLoaded(false);
    loadRemoteAccessConfig();
  }, [isLoggedIn, sessionUser, loadRemoteAccessConfig]);
  useEffect(() => {
    if (!isLoggedIn || !accessApiUrl) return;
    const t = setInterval(() => {
      loadRemoteAccessConfig();
    }, 60000);
    return () => clearInterval(t);
  }, [isLoggedIn, accessApiUrl, loadRemoteAccessConfig]);
  useEffect(() => {
    if (!isLoggedIn || sessionUser !== 'Admin' || !remoteAccessLoaded) return;
    if (skipCloudPushRef.current) {
      skipCloudPushRef.current = false;
      return;
    }
    const nowTs = Date.now();
    S.set(ACCESS_UPDATED_AT_KEY, nowTs);
    const payload = {
      users,
      access,
      userCreds,
      updatedAt: new Date(nowTs).toISOString(),
      updatedBy: sessionUser,
    };
    const t = setTimeout(() => {
      pushRemoteAccessConfig(payload);
    }, 350);
    return () => clearTimeout(t);
  }, [users, access, userCreds, isLoggedIn, sessionUser, remoteAccessLoaded, pushRemoteAccessConfig]);

  useEffect(() => {
    if (!isLoggedIn) return;
    if (obzvonAllLoaded) return;
    loadObzvonAllRemote().then((ok) => {
      if (!ok) setObzvonAllLoaded(true);
    });
  }, [isLoggedIn, obzvonAllLoaded, loadObzvonAllRemote]);

  useEffect(() => {
    if (!isLoggedIn || !accessApiUrl) return;
    pullRemoteObzvonNewRows();
    const t = setInterval(() => {
      pullRemoteObzvonNewRows();
    }, 15000);
    return () => clearInterval(t);
  }, [isLoggedIn, accessApiUrl, pullRemoteObzvonNewRows]);
  useEffect(() => {
    if (!isLoggedIn || !accessApiUrl) return;
    const t = setInterval(() => {
      flushPendingObzvonNewRows();
    }, 7000);
    return () => clearInterval(t);
  }, [isLoggedIn, accessApiUrl, flushPendingObzvonNewRows]);
  useEffect(() => {
    if (!isLoggedIn || !accessApiUrl) return;
    const runAuto = async () => {
      const now = new Date();
      const dateKey = toIsoDate(now);
      const hh = now.getHours();
      if (!dateKey || hh < OBZVON_NEW_EXPORT_HOUR) return;
      const lastDone = String(S.get(OBZVON_NEW_LAST_AUTO_DATE_KEY, '') || '');
      if (lastDone === dateKey) return;
      const ok = await publishObzvonNewRowsToGoogleSheet({ manual:false });
      if (ok) S.set(OBZVON_NEW_LAST_AUTO_DATE_KEY, dateKey);
    };
    runAuto();
    const t = setInterval(runAuto, 5 * 60 * 1000);
    return () => clearInterval(t);
  }, [isLoggedIn, accessApiUrl, publishObzvonNewRowsToGoogleSheet]);

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
      S.set('aq-session-user', '');
      return;
    }
    if (sessionUser && users.includes(sessionUser)) {
      if (!isLoggedIn) setIsLoggedIn(true);
      // Non-admin login must always work with its own profile/ruxsatlar.
      if (sessionUser !== 'Admin' && currentUser !== sessionUser) {
        setCurrentUser(sessionUser);
      } else if (!currentUser || !users.includes(currentUser)) {
        setCurrentUser(sessionUser);
      }
    } else if (isLoggedIn) {
      setIsLoggedIn(false);
    }
  }, [users, currentUser, sessionUser, isLoggedIn]);

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
    S.set('aq-session-user', user);
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
    S.set('aq-session-user', '');
  }, []);

  const switchUser = (nextUser) => {
    if (nextUser === currentUser) return;
    if (sessionUser !== 'Admin') return;
    setCurrentUser(nextUser);
  };

  const handleLoad = (result) => {
    setData(result);
    setUp(false);
    notify(`OK: ${result.customers.length} mijoz  |  ${result.rawOrders.length} zakaz  |  ${result.cashbox.length} kassa`);
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
      notify(`OK: ${d.customers.length} mijoz  |  ${d.rawOrders.length} zakaz`);
    } catch (e) {
      setAutoLoad({ loading:false, progress:'', error:e.message });
      notify('Xato: '+e.message, 'err');
    }
  }, [mainSheetUrl]);

  const addObzvonRows = useCallback((rows) => {
    if (!rows?.length) return;
    const prepared = rows.map((r) => ({ ...r, _rid: r?._rid || createRowRid() }));
    setObzvonRecords((prev) => [...prepared, ...(prev || [])]);
    upsertObzvonAllNewRows(
      prepared.map((r, i) => ({
        rid: r._rid,
        no: String(i + 1),
        customer: r.customer || '',
        callDate: r.callDate || '',
        topic: r.topic || '',
        note: r.note || '',
        nextDate: r.nextDate || '',
        orderCount: r.orderCount || '',
        operator: r.operator || currentUser || 'Admin',
        customerId: r.id || '',
        orderDate: r.orderDate || '',
        updatedAt: new Date().toISOString(),
        company: normalizeCompanyKey(companyFilter),
      }))
    );
  }, [upsertObzvonAllNewRows, currentUser, companyFilter]);

  useEffect(() => {
    loadFromConfig();
  }, [loadFromConfig]);

  const rawD = data || { customers:[], orders:[], cashbox:[], contacts:[], rawOrders:[], rawCash:[], ordersByMId:{}, cashByMId:{}, warehouseTransfers:[], kulerInstallments:[], assignmentById:{}, debtorsByBalance:[], otherDebtorsByBalance:[] };
  // Effective profile: Admin can switch users, others are locked to their own login.
  const effectiveUser = sessionUser === 'Admin'
    ? currentUser
    : (sessionUser || currentUser || 'Admin');
  const currentAccess = normalizeAccessConfig(
    effectiveUser,
    access[effectiveUser] || DEFAULT_ACCESS[effectiveUser] || DEFAULT_ACCESS.Admin
  );
  const companyAccess = currentAccess.company || {};
  const canSwitchCompany = !!companyAccess.canSwitch;
  const lockedCompany = normalizeCompanyKey(companyAccess.default || 'murodbaxsh');
  const activeCompany = canSwitchCompany ? normalizeCompanyKey(companyFilter) : lockedCompany;
  const scopeOwn = currentAccess.scope === 'own';
  const ownIds = useMemo(
    () => new Set(Object.entries(rawD.assignmentById || {}).filter(([, op]) => op === effectiveUser).map(([id]) => id)),
    [rawD.assignmentById, effectiveUser]
  );
  const scopedD = useMemo(() => {
    if (!scopeOwn) return rawD;
    return filterDataByCustomerIds(rawD, ownIds);
  }, [rawD, scopeOwn, ownIds]);
  useEffect(() => {
    if (canSwitchCompany) return;
    if (normalizeCompanyKey(companyFilter) !== lockedCompany) {
      setCompanyFilter(lockedCompany);
    }
  }, [canSwitchCompany, companyFilter, lockedCompany]);
  const D = useMemo(() => filterDataByCompany(scopedD, activeCompany), [scopedD, activeCompany]);
  const companyCustomerIds = useMemo(
    () => new Set((D.customers || []).map((c) => normalizeIdKey(c?.id)).filter(Boolean)),
    [D.customers]
  );
  const rowBelongsToCompany = useCallback((r, idSet, companyKey) => {
    const cid = normalizeIdKey(r?.id || r?.customerId);
    if (cid) return idSet.has(cid);
    const rowCompany = normalizeStoredCompanyKey(r?.company);
    if (!rowCompany) return false;
    return rowCompany === normalizeCompanyKey(companyKey);
  }, []);
  const companyObzvonAllRows = useMemo(
    () => (obzvonAllRows || []).filter((r) => rowBelongsToCompany(r, companyCustomerIds, activeCompany)),
    [obzvonAllRows, rowBelongsToCompany, companyCustomerIds, activeCompany]
  );
  const companyObzvonAllNewRows = useMemo(
    () => (obzvonAllNewRows || []).filter((r) => rowBelongsToCompany(r, companyCustomerIds, activeCompany)),
    [obzvonAllNewRows, rowBelongsToCompany, companyCustomerIds, activeCompany]
  );
  const companyObzvonRecords = useMemo(
    () => (obzvonRecords || []).filter((r) => rowBelongsToCompany(r, companyCustomerIds, activeCompany)),
    [obzvonRecords, rowBelongsToCompany, companyCustomerIds, activeCompany]
  );
  const setCompanyScopedObzvonRecords = useCallback((valueOrUpdater) => {
    setObzvonRecords((prev) => {
      const source = Array.isArray(prev) ? prev : [];
      const visible = source.filter((r) => rowBelongsToCompany(r, companyCustomerIds, activeCompany));
      const hidden = source.filter((r) => !rowBelongsToCompany(r, companyCustomerIds, activeCompany));
      const nextVisibleRaw = typeof valueOrUpdater === 'function' ? valueOrUpdater(visible) : valueOrUpdater;
      const nextVisible = (Array.isArray(nextVisibleRaw) ? nextVisibleRaw : []).map((r) => ({
        ...(r || {}),
        company: normalizeCompanyKey((r || {}).company || activeCompany),
      }));
      return [...nextVisible, ...hidden];
    });
  }, [companyCustomerIds, activeCompany, rowBelongsToCompany]);
  const debtorsByBalance = (D.debtorsByBalance && D.debtorsByBalance.length)
    ? D.debtorsByBalance
    : (D.customers || []).filter((c) => c.balanceUZS < 0);
  const otherDebtorsByBalance = D.otherDebtorsByBalance || [];
  const doljniki = useMemo(() => {
    return debtorsByBalance
      .map((c) => {
        const related = (D.ordersByMId?.[c.id] || []);
        const debtOrder = related.find((o) => o.soNum === c.lastDocNum);
        const debtOrderDate = debtOrder?.orderDate || c.lastOrderDate || '';
        return {
          lastOrderProduct: debtOrder?.product || '',
          id: c.id || '-',
          category: c.source || '-',
          name: c.name || '-',
          debtUZS: c.balanceUZS,
          lastOrderDate: debtOrderDate,
          days: debtOrderDate ? daysAgo(debtOrderDate) : c.daysAgo,
          orderNo: c.lastDocNum || '-',
          qty: c.lastQty || 0,
          lastSum: c.lastSum || 0,
          agent: c.lastAgent || '-',
          note: c.merchantNote || '-',
        };
      })
      .sort((a, b) => a.debtUZS - b.debtUZS);
  }, [debtorsByBalance, D.ordersByMId]);
  const otherDoljniki = useMemo(() => {
    return otherDebtorsByBalance
      .map((c) => {
        const related = (D.ordersByMId?.[c.id] || []);
        const debtOrder = related.find((o) => o.soNum === c.lastDocNum);
        const debtOrderDate = debtOrder?.orderDate || c.lastOrderDate || '';
        return {
          lastOrderProduct: debtOrder?.product || '',
          id: c.id || '-',
          category: c.source || '-',
          name: c.name || '-',
          debtUZS: c.balanceUZS,
          lastOrderDate: debtOrderDate,
          days: debtOrderDate ? daysAgo(debtOrderDate) : c.daysAgo,
          orderNo: c.lastDocNum || '-',
          qty: c.lastQty || 0,
          lastSum: c.lastSum || 0,
          agent: c.lastAgent || '-',
          note: c.merchantNote || '-',
        };
      })
      .sort((a, b) => a.debtUZS - b.debtUZS);
  }, [otherDebtorsByBalance, D.ordersByMId]);
  const obzvonCnt  = D.customers.filter((c)=>c.daysAgo!=null&&c.daysAgo>14).length;
  const debtorCnt  = debtorsByBalance.length;
  const doljnikiCnt = doljniki.length;
  const companyBlockedPages = useMemo(() => (
    activeCompany === 'ahmadtea'
      ? { obzvon: true, reports: true, nazorat: true }
      : {}
  ), [activeCompany]);
  const canViewPage = useCallback((id) => {
    if (companyBlockedPages[id]) return false;
    return (currentAccess.visible?.[id] ?? true);
  }, [currentAccess, companyBlockedPages]);
  const visibleNav = NAV.filter((n) => canViewPage(n.id));
  const pageMeta = NAV.find((n)=>n.id===page) || { id:'settings', icon:E.settings, label:'Nastroyka' };
  const localThemeOverride = String(S.get(`aq-ui-theme-${effectiveUser}`, '') || '').toLowerCase();
  const currentTheme = localThemeOverride === 'light'
    ? 'light'
    : localThemeOverride === 'dark'
    ? 'dark'
    : (currentAccess.ui?.theme === 'light' ? 'light' : 'dark');
  const themeVars = currentTheme === 'light' ? LIGHT_THEME_VARS : {};
  useEffect(() => {
    if (canViewPage(page)) return;
    const fallback = visibleNav[0]?.id || (canViewPage('settings') ? 'settings' : 'dash');
    setPage(fallback);
  }, [visibleNav, page, canViewPage]);

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
      <div style={{height:'100vh',display:'flex',overflow:'hidden',background:'var(--bg)', ...themeVars}}>
        {/* SIDEBAR */}
        <div style={{width:side?215:56,background:'var(--s1)',borderRight:'1px solid var(--b2)',display:'flex',flexDirection:'column',transition:'width .2s',flexShrink:0,overflow:'hidden'}}>
          <div style={{padding:'13px 11px',borderBottom:'1px solid var(--b2)',display:'flex',alignItems:'center',gap:10}}>
            <img
              src="/New.png"
              alt="AquaBiz logo"
              style={{width:30,height:30,borderRadius:8,objectFit:'cover',flexShrink:0,display:'block'}}
            />
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
            {(currentAccess.visible?.refresh ?? true) && (
              <div className="nav-i" onClick={()=>setUp(true)} style={{ opacity: autoLoad.loading ? 0.6 : 1 }}>
                <span style={{fontSize:17,flexShrink:0}}>{E.refresh}</span>
                {side && <span>Yangilash</span>}
              </div>
            )}
            {(currentAccess.visible?.settings ?? true) && (
              <div className="nav-i" onClick={()=>setPage('settings')}>
                <span style={{fontSize:17,flexShrink:0}}>{E.settings}</span>
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
                <span style={{fontSize:17,display:'inline-block',transform:side?'none':'rotate(180deg)',transition:'transform .2s'}}>{'<'}</span>
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
              {canSwitchCompany ? (
                <select
                  className="select"
                  value={activeCompany}
                  onChange={(e) => setCompanyFilter(normalizeCompanyKey(e.target.value))}
                >
                  {COMPANY_OPTIONS.map((x) => <option key={x.key} value={x.key}>{x.label}</option>)}
                </select>
              ) : (
                <span className="tag" style={{background:'var(--s3)',color:'var(--t2)'}}>
                  {companyLabelByKey(activeCompany)}
                </span>
              )}
              {sessionUser === 'Admin' ? (
                <select className="select" value={currentUser} onChange={(e)=>switchUser(e.target.value)}>
                  {users.map((u)=><option key={u}>{u}</option>)}
                </select>
              ) : (
                <span className="tag" style={{background:'var(--s3)',color:'var(--t2)'}}>{effectiveUser}</span>
              )}
              <button className="btn btn-gh btn-sm" onClick={logout}>Chiqish</button>
              {canViewPage('obzvon') && obzvonCnt>0 && (
                <button className="btn btn-sm" style={{background:'var(--rd2)',color:'var(--rd)',border:'1px solid var(--rd2)'}} onClick={()=>setPage('obzvon')}>
                  Obzvon {obzvonCnt}
                </button>
              )}
            </div>
          </div>

          <div data-filter-boundary="1" style={{flex:1,overflow:'auto',padding:16}}>
            {!data && page!=='doljniki' ? (
              <div style={{height:'100%',display:'flex',alignItems:'center',justifyContent:'center'}}>
                <div style={{textAlign:'center',maxWidth:440}}>
                  {autoLoad.loading ? (
                    <>
                      <div style={{fontSize:22,marginBottom:16,animation:'spin 1.2s linear infinite',display:'inline-block'}}>O</div>
                      <div style={{fontSize:18,fontWeight:700,marginBottom:12}}>Ma'lumot yuklanmoqda</div>
                      <div style={{color:'var(--bl)',fontSize:14,fontFamily:'var(--mono)',marginBottom:20}}>{autoLoad.progress}</div>
                    </>
                  ) : autoLoad.error ? (
                    <>
                      <div style={{fontSize:56,marginBottom:16}}>!</div>
                      <div style={{fontSize:18,fontWeight:700,marginBottom:12,color:'var(--rd)'}}>Yuklash xatosi</div>
                      <div style={{color:'var(--t3)',fontSize:13,marginBottom:8,background:'var(--rd2)',border:'1px solid var(--rd)',borderRadius:9,padding:'12px 16px',textAlign:'left'}}>{autoLoad.error}</div>
                      <div style={{display:'flex',gap:10,justifyContent:'center',marginTop:16}}>
                        <button className="btn btn-bl" onClick={()=>loadFromConfig()}>Qayta urinish</button>
                        <button className="btn btn-gh" onClick={()=>setUp(true)}>Excel yuklash</button>
                      </div>
                    </>
                  ) : (
                    <>
                      <div style={{fontSize:56,marginBottom:16}}>i</div>
                      <div style={{fontSize:18,fontWeight:700,marginBottom:12}}>Ma'lumot yuklanmagan</div>
                      <button className="btn btn-gh" onClick={()=>setUp(true)}>Excel fayl yuklash</button>
                    </>
                  )}
                </div>
              </div>
            ) : (
              <>
                {page==='dash'    && canViewPage('dash') && <Dashboard D={D}/>}
                {page==='cust'    && canViewPage('cust') && <Customers D={D} currentUser={effectiveUser} currentAccess={currentAccess} assignmentById={D.assignmentById || {}} company={activeCompany}/>}
                {page==='orders'  && canViewPage('orders') && <Orders    D={D}/>}
                {page==='kassa'   && canViewPage('kassa') && <Kassa     D={D}/>}
                {page==='obzvon'  && canViewPage('obzvon') && (
                  <Obzvon
                  D={D}
                  company={activeCompany}
                  currentUser={effectiveUser}
                  canEditAllNew={!!(currentAccess.visible?.obzvon_new_edit ?? (effectiveUser === 'Admin'))}
                  canDeleteAllNew={!!(currentAccess.visible?.obzvon_new_delete ?? (effectiveUser === 'Admin'))}
                  canPublishAllNew={!!(currentAccess.visible?.obzvon_new_publish ?? (effectiveUser === 'Admin'))}
                  records={companyObzvonRecords}
                  setRecords={setCompanyScopedObzvonRecords}
                  allRows={companyObzvonAllRows}
                  newRows={companyObzvonAllNewRows}
                  onReloadAll={loadObzvonAllRemote}
                  onReloadAllNew={pullRemoteObzvonNewRows}
                  onAppendAllRows={appendObzvonAllRemote}
                  onUpsertNewRows={upsertObzvonAllNewRows}
                  onPublishAllNew={publishObzvonNewRowsToGoogleSheet}
                />
              )}
                {page==='doljniki'&& canViewPage('doljniki') && (
                  <Doljniki
                    rows={doljniki}
                    otherRows={otherDoljniki}
                    D={D}
                    kulerRows={D.kulerInstallments || []}
                    onAddToObzvon={addObzvonRows}
                    currentUser={effectiveUser}
                    company={activeCompany}
                  />
                )}
                {page==='reports' && canViewPage('reports') && (
                  <Reports
                    D={D}
                    company={activeCompany}
                    currentUser={effectiveUser}
                    currentAccess={currentAccess}
                    users={users}
                    access={access}
                    planRows={planRows}
                    planOffDays={planOffDays}
                    obzvonNewRows={companyObzvonAllNewRows}
                    mode="reports"
                  />
                )}
                {page==='nazorat' && canViewPage('nazorat') && (
                  <Reports
                    D={D}
                    company={activeCompany}
                    currentUser={effectiveUser}
                    currentAccess={currentAccess}
                    users={users}
                    access={access}
                    planRows={planRows}
                    planOffDays={planOffDays}
                    obzvonNewRows={companyObzvonAllNewRows}
                    mode="nazorat"
                  />
                )}
                {page==='plan' && canViewPage('plan') && (
                  <PlanPage
                    company={activeCompany}
                    users={users}
                    access={access}
                    currentUser={effectiveUser}
                    planRows={planRows}
                    setPlanRows={setPlanRows}
                    planOffDays={planOffDays}
                    setPlanOffDays={setPlanOffDays}
                  />
                )}
                {page==='settings' && canViewPage('settings') && <SettingsPanel users={users} setUsers={setUsers} access={access} setAccess={setAccess} currentUser={currentUser} setCurrentUser={setCurrentUser} webhookUrl={obzvonWebhook} setWebhookUrl={setObzvonWebhook} userCreds={userCreds} setUserCreds={setUserCreds} onSwitchUser={switchUser} isAdminSession={sessionUser==='Admin'} viewerAccess={currentAccess} D={D} company={activeCompany} />}
              </>
            )}
          </div>
        </div>
      </div>

      {showUp && (
        <UploadModal
          onLoad={handleLoad}
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












