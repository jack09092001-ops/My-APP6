'use strict';

const CONFIG_KEY = 'storeDashboard.config.v1';
const DATA_CACHE_KEY = 'storeDashboard.data.v1';
const GEOCODE_CACHE_KEY = 'storeDashboard.geocodeCache.v1';
const LOCATE_RADIUS_DEFAULT_METERS = 5000;

// 跟「月營收」工作表對應門市的共用欄位名稱，兩份工作表都要用同一個字串找欄位。
const STORE_NAME_HEADER = '門市名稱';

// Excel 表頭文字 <-> 內部欄位名稱的對應表。地圖 InfoWindow、地圖下方表格、搜尋詳情卡都從
// 這裡取欄位定義，避免同一個數字在多個畫面各自寫一份、之後改欄位名稱要改好幾處。
const FIELD_DEFS = [
  { header: '門市代號', key: 'customerId', type: 'text' },
  { header: STORE_NAME_HEADER, key: 'name', type: 'text' },
  { header: '品牌', key: 'brand', type: 'text' },
  { header: '督導名稱', key: 'supervisor', type: 'text' },
  { header: '開幕日期', key: 'openDate', type: 'date' },
  { header: '門市地址', key: 'address', type: 'text' },
  { header: '緯度', key: 'lat', type: 'number' },
  { header: '經度', key: 'lng', type: 'number' },
  { header: '店面租金', key: 'rent', type: 'currency' },
  { header: '叫貨金額', key: 'purchaseAmount', type: 'currency' },
  { header: '營業天數(平)', key: 'weekdayDays', type: 'int' },
  { header: '日均營業額(平)', key: 'weekdayAvgRevenue', type: 'currency' },
  { header: '日均來客數(平)', key: 'weekdayVisitors', type: 'int' },
  { header: '營業天數(假)', key: 'holidayDays', type: 'int' },
  { header: '日均營業額(假)', key: 'holidayAvgRevenue', type: 'currency' },
  { header: '日均來客數(假)', key: 'holidayVisitors', type: 'int' },
  { header: 'POS實收金額', key: 'posActual', type: 'currency' },
  { header: 'POS當月營收(預估)', key: 'posMonthlyEstimate', type: 'currency' },
  { header: '當月預估損益', key: 'estimatedProfit', type: 'currency-signed' },
  { header: '內用佔比', key: 'dineInPct', type: 'percent' },
  { header: '外帶佔比', key: 'takeoutPct', type: 'percent' },
  { header: '外送佔比', key: 'deliveryPct', type: 'percent' },
  { header: '掃碼佔比', key: 'qrPct', type: 'percent' },
  { header: 'APP佔比', key: 'appPct', type: 'percent' },
];
const FIELD_DEF_BY_KEY = Object.fromEntries(FIELD_DEFS.map((d) => [d.key, d]));
const MAP_INFO_KEYS = ['rent', 'posMonthlyEstimate'];
const NEARBY_TABLE_KEYS = ['rent', 'posMonthlyEstimate'];
const DETAIL_KEYS = FIELD_DEFS.filter((d) => !['name', 'lat', 'lng'].includes(d.key)).map((d) => d.key);

// 品牌別 -> 地圖標記顏色。沒對到這兩個值（空白、打錯字等）一律灰點，不會整個標記消失。
const BRAND_MARKER_COLORS = { MWD: '#2fa84f', CM: '#f2994a' };
const BRAND_MARKER_FALLBACK_COLOR = '#9aa5b1';

const state = {
  rows: [],
  monthlyRevenueByStore: new Map(),
  monthlyCostRateByStore: new Map(),
  monthlyKeyItemsByStore: new Map(),
  salesAnomalyByStore: new Map(),
  lastFetchedAt: null,
  modifiedTime: null,
  map: null,
  infoWindow: null,
  userMarker: null,
  userCircle: null,
  storeMarkers: [],
  userPos: null,
  locateRadiusMeters: LOCATE_RADIUS_DEFAULT_METERS,
  mapsReady: false,
  currentDetailName: null,
};

document.addEventListener('DOMContentLoaded', init);

async function init() {
  setupTabs();
  setupSettingsForm();
  setupSearch();
  setupRadiusSlider();

  const config = loadConfig();
  populateSettingsForm(config);

  const cached = loadCachedData();
  if (cached) {
    applyData(cached);
    showFreshness(cached, true);
  }
  renderDataStatus();

  if (!isConfigComplete(config)) {
    document.getElementById('setup-banner').classList.remove('hidden');
    switchTab('settings-tab');
    return;
  }
  document.getElementById('setup-banner').classList.add('hidden');

  loadGoogleMapsScript(config.mapsApiKey)
    .then(() => {
      initMapView();
      if (state.rows.length) {
        plotStores(state.rows);
        renderNearbyTable();
      }
    })
    .catch((err) => {
      setGeoStatus('地圖載入失敗：' + err.message);
    });

  await refreshData(config, Boolean(cached));
}

async function refreshData(config, hadCache) {
  try {
    const fresh = await fetchExcelData(config);
    saveCachedData(fresh);
    applyData(fresh);
    showFreshness(fresh, false);
  } catch (err) {
    if (!hadCache) {
      setGeoStatus('資料載入失敗：' + err.message);
      showFreshnessError(err);
    } else {
      showFreshnessError(err, true);
    }
  }
}

async function manualRefresh() {
  const config = loadConfig();
  if (!isConfigComplete(config)) return;
  setFreshnessText('正在重新整理…');
  await refreshData(config, state.rows.length > 0);
}

function applyData(data) {
  state.rows = data.rows;
  state.monthlyRevenueByStore = buildMonthlySeriesIndex(data.monthlyRevenue || []);
  state.monthlyCostRateByStore = buildMonthlySeriesIndex(data.monthlyCostRate || []);
  state.monthlyKeyItemsByStore = buildMonthlySeriesIndex(data.monthlyKeyItems || []);
  state.salesAnomalyByStore = buildSalesAnomalyIndex(data.salesAnomaly || []);
  state.lastFetchedAt = data.fetchedAt;
  state.modifiedTime = data.modifiedTime;
  renderDatalist(state.rows);
  renderDataStatus();
  if (state.mapsReady) {
    plotStores(state.rows);
    renderNearbyTable();
  }
  if (state.currentDetailName) renderDetailByName(state.currentDetailName);
}

// ---------- Tabs ----------

function setupTabs() {
  document.getElementById('tabbar').addEventListener('click', (e) => {
    const btn = e.target.closest('.tab-btn');
    if (!btn) return;
    switchTab(btn.dataset.tab);
  });
}

function switchTab(tabId) {
  document.querySelectorAll('.tab-panel').forEach((el) => el.classList.toggle('hidden', el.id !== tabId));
  document.querySelectorAll('.tab-btn').forEach((btn) => btn.classList.toggle('active', btn.dataset.tab === tabId));
}

// ---------- Settings ----------

function loadConfig() {
  const stored = (() => {
    try {
      const raw = localStorage.getItem(CONFIG_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch {
      return {};
    }
  })();
  const defaults = window.APP_CONFIG || {};
  return {
    mapsApiKey: stored.mapsApiKey || defaults.mapsApiKey || '',
    driveApiKey: stored.driveApiKey || defaults.driveApiKey || '',
    fileId: stored.fileId || defaults.fileId || '',
  };
}

function saveConfig(config) {
  try {
    localStorage.setItem(CONFIG_KEY, JSON.stringify(config));
  } catch {
    // 私密瀏覽模式等情境下 localStorage 可能無法寫入，安靜略過即可，不影響本次執行期使用。
  }
}

function isConfigComplete(c) {
  return Boolean(c && c.mapsApiKey && c.driveApiKey && c.fileId);
}

function populateSettingsForm(config) {
  document.getElementById('maps-api-key').value = config.mapsApiKey || '';
  document.getElementById('drive-api-key').value = config.driveApiKey || '';
  document.getElementById('drive-file-id').value = config.fileId || '';
}

function setupSettingsForm() {
  document.getElementById('settings-form').addEventListener('submit', (e) => {
    e.preventDefault();
    const config = {
      mapsApiKey: document.getElementById('maps-api-key').value.trim(),
      driveApiKey: document.getElementById('drive-api-key').value.trim(),
      fileId: document.getElementById('drive-file-id').value.trim(),
    };
    saveConfig(config);
    location.reload();
  });
  document.getElementById('settings-refresh-btn').addEventListener('click', manualRefresh);
  document.getElementById('refresh-btn').addEventListener('click', manualRefresh);
}

function renderDataStatus() {
  const el = document.getElementById('data-status-content');
  if (!state.rows.length) {
    el.textContent = '尚未載入門市資料。';
    return;
  }
  const modified = state.modifiedTime ? new Date(state.modifiedTime).toLocaleString('zh-Hant-TW') : '未知';
  const fetched = state.lastFetchedAt ? new Date(state.lastFetchedAt).toLocaleString('zh-Hant-TW') : '未知';
  el.textContent = `已載入 ${state.rows.length} 筆門市資料。Excel 檔案最後修改於 ${modified}；App 最後抓取於 ${fetched}。`;
}

function showFreshness(data, isStale) {
  const bar = document.getElementById('freshness-bar');
  bar.classList.remove('hidden');
  const fetched = data.fetchedAt ? new Date(data.fetchedAt).toLocaleString('zh-Hant-TW') : '';
  const modified = data.modifiedTime ? new Date(data.modifiedTime).toLocaleString('zh-Hant-TW') : null;
  setFreshnessText(
    isStale ? `離線快取資料（抓取於 ${fetched}）` : modified ? `資料更新於 ${modified}` : `已於 ${fetched} 重新整理`
  );
}

function showFreshnessError(err, hasStaleData) {
  const bar = document.getElementById('freshness-bar');
  bar.classList.remove('hidden');
  setFreshnessText((hasStaleData ? '重新整理失敗，顯示舊資料。' : '資料載入失敗。') + '（' + err.message + '）');
}

function setFreshnessText(text) {
  document.getElementById('freshness-text').textContent = text;
}

function setGeoStatus(text) {
  document.getElementById('geo-status').textContent = text;
}

// ---------- Data fetch (Google Drive API) ----------

async function fetchExcelData(config) {
  const base = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(config.fileId)}`;
  const [fileRes, metaRes] = await Promise.all([
    fetch(`${base}?alt=media&key=${encodeURIComponent(config.driveApiKey)}`),
    fetch(`${base}?fields=modifiedTime&key=${encodeURIComponent(config.driveApiKey)}`),
  ]);
  if (!fileRes.ok) throw new Error(await driveErrorMessage(fileRes));
  const buffer = await fileRes.arrayBuffer();
  const { rows, monthlyRevenue, monthlyCostRate, monthlyKeyItems, salesAnomaly } = await parseWorkbook(buffer);
  await geocodeMissingCoordinates(rows, config.mapsApiKey);
  let modifiedTime = null;
  if (metaRes.ok) {
    const meta = await metaRes.json();
    modifiedTime = meta.modifiedTime || null;
  }
  return { rows, monthlyRevenue, monthlyCostRate, monthlyKeyItems, salesAnomaly, fetchedAt: new Date().toISOString(), modifiedTime };
}

// ---------- 地址轉座標（門市 Excel 只有地址、沒有經緯度時）----------
// 結果永久快取在 localStorage（門市地址不太會變），同一個地址只會呼叫 Geocoding API 一次，
// 之後每次重新整理資料都直接吃快取，不會每次都重打一輪。

const GEOCODE_CONCURRENCY = 8;

function loadGeocodeCache() {
  try {
    const raw = localStorage.getItem(GEOCODE_CACHE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveGeocodeCache(cache) {
  try {
    localStorage.setItem(GEOCODE_CACHE_KEY, JSON.stringify(cache));
  } catch {
    // 快取寫入失敗頂多下次重新查一次，不影響當次顯示。
  }
}

async function geocodeOneAddress(address, apiKey) {
  try {
    const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${encodeURIComponent(apiKey)}`;
    const res = await fetch(url);
    const data = await res.json();
    if (data.status === 'OK' && data.results && data.results[0]) {
      const loc = data.results[0].geometry.location;
      return { lat: loc.lat, lng: loc.lng };
    }
  } catch {
    // 忽略單筆失敗，該店這次就不會出現在地圖上，不影響其他店。
  }
  return null;
}

async function geocodeMissingCoordinates(rows, apiKey) {
  const needsGeocode = rows.filter((r) => !hasValidCoords(r) && r.address);
  if (!needsGeocode.length) return rows;

  const cache = loadGeocodeCache();
  const uniqueAddresses = [...new Set(needsGeocode.map((r) => r.address))];
  const toFetch = uniqueAddresses.filter((addr) => !(addr in cache));

  for (let i = 0; i < toFetch.length; i += GEOCODE_CONCURRENCY) {
    const batch = toFetch.slice(i, i + GEOCODE_CONCURRENCY);
    await Promise.all(
      batch.map(async (addr) => {
        cache[addr] = await geocodeOneAddress(addr, apiKey);
      })
    );
    saveGeocodeCache(cache); // 每批就存一次，中途失敗也不會前功盡棄
    setFreshnessText(`正在將門市地址轉換為地圖座標…（${Math.min(i + GEOCODE_CONCURRENCY, toFetch.length)}/${toFetch.length}）`);
  }

  needsGeocode.forEach((r) => {
    const coords = cache[r.address];
    if (coords) {
      r.lat = coords.lat;
      r.lng = coords.lng;
    }
  });
  return rows;
}

async function driveErrorMessage(res) {
  try {
    const data = await res.json();
    return (data && data.error && data.error.message) || `Drive API 錯誤（${res.status}）`;
  } catch {
    return `Drive API 錯誤（${res.status}）`;
  }
}

function loadCachedData() {
  try {
    const raw = localStorage.getItem(DATA_CACHE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function saveCachedData(data) {
  try {
    localStorage.setItem(DATA_CACHE_KEY, JSON.stringify(data));
  } catch {
    // 快取寫入失敗不影響當次畫面顯示，安靜略過。
  }
}

// ---------- Excel parsing ----------
// 主要門市資料放第一個工作表（欄位見 FIELD_DEFS）；月度趨勢資料（營收／成本率／重點商品
// 銷貨項次）各自放在獨立工作表，寬格式：門市名稱 + 每個月一欄。

async function parseWorkbook(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return {
    rows: parseStoreRows(wb),
    monthlyRevenue: parseWideMonthlySeries(wb, MONTHLY_SERIES_DEFS.revenue.sheet, MONTHLY_SERIES_DEFS.revenue.pattern),
    monthlyCostRate: parseWideMonthlySeries(wb, MONTHLY_SERIES_DEFS.costRate.sheet, MONTHLY_SERIES_DEFS.costRate.pattern),
    monthlyKeyItems: parseWideMonthlySeries(wb, MONTHLY_SERIES_DEFS.keyItems.sheet, MONTHLY_SERIES_DEFS.keyItems.pattern),
    salesAnomaly: parseSalesAnomalyRows(wb),
  };
}

function parseStoreRows(wb) {
  const ws = wb.worksheets[0];
  if (!ws) return [];

  const colByKey = {};
  ws.getRow(1).eachCell((cell, colNumber) => {
    const text = String(cell.value ?? '').trim();
    const def = FIELD_DEFS.find((d) => d.header === text);
    if (def) colByKey[def.key] = colNumber;
  });

  const rows = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const nameCol = colByKey.name;
    const nameVal = nameCol ? row.getCell(nameCol).value : null;
    if (!nameVal) continue;

    const obj = {};
    for (const def of FIELD_DEFS) {
      const col = colByKey[def.key];
      obj[def.key] = col ? normalizeCellValue(row.getCell(col).value, def.type) : null;
    }
    rows.push(obj);
  }
  return rows;
}

// 三份月度趨勢工作表格式都一樣：門市名稱 + 每個月一欄，差別只在工作表名稱跟欄名的月份寫法。
// 「月營收」欄名開頭是「2026年1月」（後面接什麼字都算，例如「2026年1月POS營收」）；
// 「叫貨金額」工作表裡的成本率欄、「重點商品銷貨項次」工作表都是「202601成本率」/「202601銷貨項次」
// 這種六碼年月＋固定字尾的格式。每月新增資料時，對應工作表往右加一欄即可，不用改程式。
const MONTHLY_SERIES_DEFS = {
  revenue: { sheet: '月營收', pattern: /^(\d{4})年(\d{1,2})月/ },
  costRate: { sheet: '叫貨金額', pattern: /^(\d{4})(\d{2})成本率$/ },
  keyItems: { sheet: '重點商品銷貨項次', pattern: /^(\d{4})(\d{2})銷貨項次$/ },
};

function parseWideMonthlySeries(wb, sheetName, monthHeaderPattern) {
  const ws = wb.getWorksheet(sheetName);
  if (!ws) return [];

  let nameCol = null;
  const monthCols = [];
  ws.getRow(1).eachCell((cell, colNumber) => {
    const text = String(cell.value ?? '').trim();
    if (text === STORE_NAME_HEADER) {
      nameCol = colNumber;
      return;
    }
    const m = text.match(monthHeaderPattern);
    if (m) monthCols.push({ col: colNumber, year: Number(m[1]), month: Number(m[2]) });
  });
  if (!nameCol || !monthCols.length) return [];

  const rows = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const nameVal = normalizeCellValue(row.getCell(nameCol).value, 'text');
    if (!nameVal) continue;
    monthCols.forEach(({ col, year, month }) => {
      const value = normalizeCellValue(row.getCell(col).value, 'number');
      if (value === null) return; // 該店該月還沒填資料，不畫這個點
      rows.push({ name: nameVal, year, month, value });
    });
  }
  return rows;
}

function buildMonthlySeriesIndex(rows) {
  const map = new Map();
  rows.forEach((r) => {
    if (!r.name) return;
    const point = { year: r.year, month: r.month, value: Number(r.value), sortTime: new Date(r.year, r.month - 1, 1).getTime() };
    if (!map.has(r.name)) map.set(r.name, []);
    map.get(r.name).push(point);
  });
  map.forEach((points) => points.sort((a, b) => a.sortTime - b.sortTime));
  return map;
}

// 「銷貨異常清單」是長格式（一列＝一家店、一個月、一項異常商品），跟月度趨勢表的寬格式不同，
// 因為一家店一個月可能對到 0~51 個商品，塞進同一欄位不好維護也不好解析。
const SALES_ANOMALY_SHEET_NAME = '銷貨異常清單';
const SALES_ANOMALY_HEADERS = {
  月份: 'yearMonth',
  [STORE_NAME_HEADER]: 'name',
  異常商品: 'product',
  近三個月叫貨: 'recentOrder',
  '近三個月標準(50%)': 'standard',
  差距: 'gap',
};

function parseSalesAnomalyRows(wb) {
  const ws = wb.getWorksheet(SALES_ANOMALY_SHEET_NAME);
  if (!ws) return [];

  const colByKey = {};
  ws.getRow(1).eachCell((cell, colNumber) => {
    const text = String(cell.value ?? '').trim();
    const key = SALES_ANOMALY_HEADERS[text];
    if (key) colByKey[key] = colNumber;
  });
  if (!colByKey.name || !colByKey.yearMonth || !colByKey.product) return [];

  const rows = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const nameVal = normalizeCellValue(row.getCell(colByKey.name).value, 'text');
    if (!nameVal) continue;
    const yearMonthVal = normalizeCellValue(row.getCell(colByKey.yearMonth).value, 'number');
    if (yearMonthVal === null) continue;
    rows.push({
      name: nameVal,
      yearMonth: yearMonthVal,
      product: colByKey.product ? normalizeCellValue(row.getCell(colByKey.product).value, 'text') : null,
      recentOrder: colByKey.recentOrder ? normalizeCellValue(row.getCell(colByKey.recentOrder).value, 'number') : null,
      standard: colByKey.standard ? normalizeCellValue(row.getCell(colByKey.standard).value, 'number') : null,
      gap: colByKey.gap ? normalizeCellValue(row.getCell(colByKey.gap).value, 'number') : null,
    });
  }
  return rows;
}

function buildSalesAnomalyIndex(rows) {
  const map = new Map();
  rows.forEach((r) => {
    if (!r.name) return;
    if (!map.has(r.name)) map.set(r.name, new Map());
    const byMonth = map.get(r.name);
    if (!byMonth.has(r.yearMonth)) byMonth.set(r.yearMonth, []);
    byMonth.get(r.yearMonth).push(r);
  });
  return map;
}

// 不能直接用 isFinite(row.lat)：isFinite(null) 會回傳 true（null 被硬轉成 0），
// 導致「還沒有座標」的門市被誤判成「座標是 0,0」，混進地圖跟地址轉換的篩選條件裡。
function hasValidCoords(row) {
  return typeof row.lat === 'number' && typeof row.lng === 'number' && isFinite(row.lat) && isFinite(row.lng);
}

function normalizeCellValue(raw, type) {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === 'object' && 'result' in raw) return normalizeCellValue(raw.result, type);
  if (typeof raw === 'object' && raw.richText) return raw.richText.map((t) => t.text).join('');
  if (typeof raw === 'object' && 'text' in raw) return raw.text;

  if (type === 'date') {
    if (raw instanceof Date) return raw;
    const d = new Date(raw);
    return isNaN(d.getTime()) ? raw : d;
  }
  if (type === 'text') return String(raw);

  const n = Number(raw);
  return isNaN(n) ? null : n;
}

// Excel 儲存格若用「百分比」格式，ExcelJS 讀到的是 0~1 的小數（0.35 = 35%）；
// 若使用者直接打數字（35），這裡當作已經是百分比數值，兩種輸入方式都認得。
function toPercentNumber(value) {
  const n = Number(value);
  return Math.abs(n) <= 1 ? n * 100 : n;
}

function formatValue(value, type) {
  if (value === null || value === undefined || value === '') return '—';
  switch (type) {
    case 'date': {
      const d = value instanceof Date ? value : new Date(value);
      return isNaN(d.getTime()) ? String(value) : d.toLocaleDateString('zh-Hant-TW');
    }
    case 'currency':
      return Number(value).toLocaleString('zh-Hant-TW') + ' 元';
    case 'currency-signed': {
      const n = Number(value);
      return (n > 0 ? '+' : '') + n.toLocaleString('zh-Hant-TW') + ' 元';
    }
    case 'int':
    case 'number':
      return Number(value).toLocaleString('zh-Hant-TW');
    case 'percent':
      return toPercentNumber(value).toLocaleString('zh-Hant-TW', { maximumFractionDigits: 1 }) + '%';
    default:
      return String(value);
  }
}

// ---------- Google Maps ----------

function loadGoogleMapsScript(apiKey) {
  if (window.google && window.google.maps) return Promise.resolve();
  if (window.__gmapsLoadingPromise) return window.__gmapsLoadingPromise;
  window.__gmapsLoadingPromise = new Promise((resolve, reject) => {
    const cbName = '__initGoogleMapsCallback';
    window[cbName] = () => resolve();
    const script = document.createElement('script');
    script.src = `https://maps.googleapis.com/maps/api/js?key=${encodeURIComponent(apiKey)}&callback=${cbName}`;
    script.async = true;
    script.onerror = () => reject(new Error('無法載入 Google Maps（請確認金鑰是否正確）'));
    document.head.appendChild(script);
  });
  return window.__gmapsLoadingPromise;
}

function initMapView() {
  const mapEl = document.getElementById('map');
  state.map = new google.maps.Map(mapEl, { center: { lat: 23.9, lng: 121.0 }, zoom: 8 });
  state.infoWindow = new google.maps.InfoWindow();
  state.mapsReady = true;
  renderMapLegend();
  addLocateControl();
  setGeoStatus(`顯示全部門市，可調整下方範圍後按右下角「📍 定位」顯示我附近門市`);
}

function setupRadiusSlider() {
  const slider = document.getElementById('radius-slider');
  const valueLabel = document.getElementById('radius-value');
  slider.addEventListener('input', () => {
    valueLabel.textContent = slider.value;
    state.locateRadiusMeters = Math.round(Number(slider.value) * 1000);
    if (state.userCircle) state.userCircle.setRadius(state.locateRadiusMeters);
    if (state.userPos) renderNearbyTable();
  });
}

function renderMapLegend() {
  const entries = [...Object.entries(BRAND_MARKER_COLORS), ['其他', BRAND_MARKER_FALLBACK_COLOR]];
  document.getElementById('map-legend').innerHTML = entries
    .map(([label, color]) => `<span class="legend-item"><span class="legend-dot" style="background:${color}"></span>${escapeHtml(label)}</span>`)
    .join('');
}

function addLocateControl() {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'map-locate-btn';
  btn.title = '定位我的位置';
  btn.textContent = '📍 定位';
  btn.addEventListener('click', handleLocateClick);
  state.map.controls[google.maps.ControlPosition.RIGHT_BOTTOM].push(btn);
}

function handleLocateClick() {
  if (!navigator.geolocation) {
    setGeoStatus('此裝置不支援定位。');
    return;
  }
  setGeoStatus('正在取得目前位置…');
  navigator.geolocation.getCurrentPosition(
    (pos) => {
      state.userPos = { lat: pos.coords.latitude, lng: pos.coords.longitude };
      setGeoStatus(`已定位，下方清單顯示 ${state.locateRadiusMeters / 1000}km 內門市`);
      placeUserMarker();
      state.map.setCenter(state.userPos);
      state.map.setZoom(13);
      renderNearbyTable();
    },
    (err) => {
      setGeoStatus(`無法取得定位（${geoErrorMessage(err)}）`);
    },
    { enableHighAccuracy: true, timeout: 10000, maximumAge: 60000 }
  );
}

function geoErrorMessage(err) {
  switch (err.code) {
    case 1:
      return '定位權限被拒絕';
    case 2:
      return '無法取得定位資訊';
    case 3:
      return '定位逾時';
    default:
      return err.message || '未知錯誤';
  }
}

function placeUserMarker() {
  if (!state.map) return;
  if (state.userMarker) state.userMarker.setMap(null);
  state.userMarker = new google.maps.Marker({
    position: state.userPos,
    map: state.map,
    title: '目前位置',
    icon: {
      path: google.maps.SymbolPath.CIRCLE,
      scale: 8,
      fillColor: '#2f6fed',
      fillOpacity: 1,
      strokeColor: '#ffffff',
      strokeWeight: 2,
    },
  });
  if (state.userCircle) {
    // 拉桿之後重新定位：更新既有圓圈，不要疊一個新的上去。
    state.userCircle.setCenter(state.userPos);
    state.userCircle.setRadius(state.locateRadiusMeters);
  } else {
    state.userCircle = new google.maps.Circle({
      center: state.userPos,
      radius: state.locateRadiusMeters,
      map: state.map,
      fillColor: '#2f6fed',
      fillOpacity: 0.08,
      strokeColor: '#2f6fed',
      strokeOpacity: 0.4,
      strokeWeight: 1,
    });
  }
  state.map.setCenter(state.userPos);
}

function clearStoreMarkers() {
  state.storeMarkers.forEach((m) => m.setMap(null));
  state.storeMarkers = [];
}

function markerIconForBrand(brand) {
  const color = BRAND_MARKER_COLORS[String(brand || '').trim()] || BRAND_MARKER_FALLBACK_COLOR;
  return {
    path: google.maps.SymbolPath.CIRCLE,
    scale: 7,
    fillColor: color,
    fillOpacity: 1,
    strokeColor: '#ffffff',
    strokeWeight: 1.5,
  };
}

function plotStores(rows) {
  if (!state.map) return;
  clearStoreMarkers();
  const bounds = new google.maps.LatLngBounds();
  let any = false;
  rows.forEach((row) => {
    if (!hasValidCoords(row)) return;
    any = true;
    const marker = new google.maps.Marker({
      position: { lat: row.lat, lng: row.lng },
      map: state.map,
      title: row.name,
      icon: markerIconForBrand(row.brand),
    });
    marker.addListener('click', () => {
      state.infoWindow.setContent(buildInfoWindowHtml(row));
      state.infoWindow.open(state.map, marker);
    });
    state.storeMarkers.push(marker);
    bounds.extend(marker.getPosition());
  });
  // 定位過之後畫面已經是使用者自己選的 5km 檢視範圍，重新整理資料時不要把地圖搶回去
  // 重新 fit 成全台灣範圍，打斷使用者原本在看的畫面。
  if (any && !state.userPos) {
    state.map.fitBounds(bounds);
  }
}

function buildInfoWindowHtml(row) {
  const lines = MAP_INFO_KEYS.map((key) => {
    const def = FIELD_DEF_BY_KEY[key];
    return `<div>${escapeHtml(def.header)}：${escapeHtml(formatValue(row[key], def.type))}</div>`;
  }).join('');
  return `<div class="gm-info-content"><strong>${escapeHtml(row.name)}</strong>${lines}</div>`;
}

// ---------- Nearby list (map tab) ----------

function haversineMeters(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = (d) => (d * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a =
    Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// 一列＝一家門市（門市名稱／距離／租金／POS當月營收(預估)），直式往下呈現。
// 這份清單跟著「定位」按鈕＋下方拉桿的結果走：還沒定位過就顯示提示，定位後篩選範圍內、依距離排序。
function renderNearbyTable() {
  const card = document.getElementById('nearby-list-card');
  const container = document.getElementById('nearby-list');
  card.classList.remove('hidden');
  container.innerHTML = '';

  const radiusKm = state.locateRadiusMeters / 1000;

  if (!state.userPos) {
    container.innerHTML = `<p class="empty-state">按地圖右下角「📍 定位」，顯示範圍內門市資料</p>`;
    return;
  }

  const rows = state.rows
    .filter((r) => hasValidCoords(r))
    .map((r) => ({ ...r, distance: haversineMeters(state.userPos.lat, state.userPos.lng, r.lat, r.lng) }))
    .filter((r) => r.distance <= state.locateRadiusMeters)
    .sort((a, b) => a.distance - b.distance);

  if (!rows.length) {
    container.innerHTML = `<p class="empty-state">${radiusKm}km 內沒有門市資料</p>`;
    return;
  }

  const headerHtml = ['門市名稱', '距離', ...NEARBY_TABLE_KEYS.map((key) => FIELD_DEF_BY_KEY[key].header)]
    .map((label) => `<th>${escapeHtml(label)}</th>`)
    .join('');
  const bodyRows = rows
    .map((row) => {
      const metricCells = NEARBY_TABLE_KEYS.map((key) => `<td>${escapeHtml(formatValue(row[key], FIELD_DEF_BY_KEY[key].type))}</td>`).join('');
      return `<tr>
        <td class="store-name-cell"><button type="button" class="store-col-btn" data-name="${escapeHtml(row.name)}">${escapeHtml(row.name)}</button></td>
        <td>${Math.round(row.distance).toLocaleString('zh-Hant-TW')} 公尺</td>
        ${metricCells}
      </tr>`;
    })
    .join('');

  const scrollWrap = document.createElement('div');
  scrollWrap.className = 'table-scroll';
  scrollWrap.innerHTML = `
    <table class="nearby-table">
      <thead><tr>${headerHtml}</tr></thead>
      <tbody>${bodyRows}</tbody>
    </table>
  `;
  container.appendChild(scrollWrap);
  container.querySelectorAll('.store-col-btn').forEach((btn) => {
    btn.addEventListener('click', () => openStoreDetail(btn.dataset.name));
  });
}

// ---------- Search tab ----------

function renderDatalist(rows) {
  const dl = document.getElementById('store-names');
  dl.innerHTML = '';
  rows
    .map((r) => r.name)
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, 'zh-Hant'))
    .forEach((name) => {
      const opt = document.createElement('option');
      opt.value = name;
      dl.appendChild(opt);
    });
}

function setupSearch() {
  document.getElementById('store-search').addEventListener('input', (e) => {
    renderDetailByName(e.target.value.trim());
  });
}

function openStoreDetail(name) {
  switchTab('search-tab');
  document.getElementById('store-search').value = name;
  renderDetailByName(name);
}

function renderDetailByName(name) {
  const card = document.getElementById('detail-card');
  const row = state.rows.find((r) => r.name === name);
  if (!row) {
    card.classList.add('hidden');
    state.currentDetailName = null;
    return;
  }
  state.currentDetailName = name;
  card.classList.remove('hidden');
  document.getElementById('detail-store-name').textContent = row.name;

  const grid = document.createElement('div');
  grid.className = 'detail-grid';
  DETAIL_KEYS.forEach((key) => {
    const def = FIELD_DEF_BY_KEY[key];
    const value = row[key];
    const cls = def.type === 'currency-signed' ? (Number(value) > 0 ? 'positive' : Number(value) < 0 ? 'negative' : '') : '';
    const wrap = document.createElement('div');
    wrap.className = 'detail-row';
    wrap.innerHTML = `<span class="detail-label">${escapeHtml(def.header)}</span><span class="detail-value ${cls}">${escapeHtml(formatValue(value, def.type))}</span>`;
    grid.appendChild(wrap);
  });

  const content = document.getElementById('detail-content');
  content.innerHTML = '';
  content.appendChild(grid);

  const chartSection = document.createElement('div');
  chartSection.innerHTML =
    renderMonthlyTrendSection('POS 營收趨勢', buildWindowedSeries(state.monthlyRevenueByStore, row.name), (v) => formatValue(v, 'number'), formatMonthLabelSlash) +
    renderMonthlyTrendSection('成本率趨勢', buildWindowedSeries(state.monthlyCostRateByStore, row.name), (v) => formatValue(v, 'percent'), formatMonthLabelCompact) +
    renderMonthlyTrendSection('重點商品銷貨項次趨勢', buildWindowedSeries(state.monthlyKeyItemsByStore, row.name), (v) => formatValue(v, 'number') + ' 項', formatMonthLabelCompact) +
    renderSalesAnomalySectionShell();
  content.appendChild(chartSection);
  setupSalesAnomalySection(row.name);
}

// 三張趨勢圖固定顯示「上個月往前推 12 個月」的滾動區間（例如現在是 2026/8，就顯示
// 2025/8~2026/7），不是「資料裡有多少個月就顯示多少個月」，這樣月份軸每次都一樣寬、
// 也不會因為某個月忘了填資料而讓圖看起來莫名其妙變窄。缺資料的月份仍然佔一個 X 軸位置，
// 只是不畫點、不連線（見 buildMonthlyTrendSvg）。
function getRollingTwelveMonths(referenceDate = new Date()) {
  const lastMonth = new Date(referenceDate.getFullYear(), referenceDate.getMonth() - 1, 1);
  const months = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date(lastMonth.getFullYear(), lastMonth.getMonth() - i, 1);
    months.push({ year: d.getFullYear(), month: d.getMonth() + 1 });
  }
  return months;
}

function buildWindowedSeries(seriesMap, storeName) {
  const points = seriesMap.get(storeName) || [];
  const byKey = new Map(points.map((p) => [`${p.year}-${p.month}`, p.value]));
  return getRollingTwelveMonths().map(({ year, month }) => ({
    year,
    month,
    value: byKey.has(`${year}-${month}`) ? byKey.get(`${year}-${month}`) : null,
  }));
}

// ---------- 月度趨勢折線圖（POS 營收／成本率／重點商品銷貨項次共用同一套繪圖邏輯）----------

const TREND_CHART_POINT_GAP = 70;
const TREND_CHART_PAD_X = 30;
const TREND_CHART_PAD_TOP = 26;
const TREND_CHART_PLOT_H = 150;
const TREND_CHART_XLABEL_Y = TREND_CHART_PAD_TOP + TREND_CHART_PLOT_H + 22;
const TREND_CHART_H = TREND_CHART_XLABEL_Y + 10;
const TREND_CHART_FONT = "'Segoe UI','PingFang TC','Microsoft JhengHei',sans-serif";

function formatMonthLabelSlash(year, month) {
  return `${year}/${month}`;
}

function formatMonthLabelCompact(year, month) {
  return `${year}${String(month).padStart(2, '0')}`;
}

function renderMonthlyTrendSection(heading, points, formatPointValue, formatMonthLabel) {
  const headingHtml = `<h3 class="subsection-title">${escapeHtml(heading)}</h3>`;
  const hasAnyValue = points.some((p) => typeof p.value === 'number' && isFinite(p.value));
  if (!hasAnyValue) {
    return headingHtml + `<p class="empty-state">目前沒有${escapeHtml(heading)}資料。</p>`;
  }
  return headingHtml + `<div class="chart-scroll">${buildMonthlyTrendSvg(points, formatPointValue, formatMonthLabel, heading)}</div>`;
}

function buildMonthlyTrendSvg(points, formatPointValue, formatMonthLabel, ariaLabel) {
  const n = points.length;
  const w = Math.max(280, TREND_CHART_PAD_X * 2 + TREND_CHART_POINT_GAP * Math.max(n - 1, 0));
  const xFor = (i) => (n === 1 ? w / 2 : TREND_CHART_PAD_X + i * TREND_CHART_POINT_GAP);

  // 缺資料的月份還是要佔一個 X 軸位置（月份標籤照樣顯示），只是不畫點、不連線——
  // 所以「有值的點」跟「X 軸月份標籤」分開兩批處理，不能共用同一份 coords。
  const validPoints = points.map((p, i) => ({ ...p, x: xFor(i) })).filter((p) => typeof p.value === 'number' && isFinite(p.value));

  const values = validPoints.map((p) => p.value);
  const min = values.length ? Math.min(0, ...values) : 0;
  const max = values.length ? Math.max(...values) : 1;
  const range = max - min || 1;
  const yFor = (v) => TREND_CHART_PAD_TOP + TREND_CHART_PLOT_H - ((v - min) / range) * TREND_CHART_PLOT_H;

  const grid = [0, 0.5, 1]
    .map((f) => {
      const y = (TREND_CHART_PAD_TOP + TREND_CHART_PLOT_H - f * TREND_CHART_PLOT_H).toFixed(1);
      return `<line x1="${TREND_CHART_PAD_X}" y1="${y}" x2="${w - TREND_CHART_PAD_X}" y2="${y}" stroke="#e0e4e8" stroke-width="1"/>`;
    })
    .join('');

  const coords = validPoints.map((p) => ({ ...p, y: yFor(p.value) }));
  const line =
    coords.length > 1
      ? `<polyline points="${coords.map((c) => `${c.x.toFixed(1)},${c.y.toFixed(1)}`).join(' ')}" fill="none" stroke="#2f6fed" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`
      : '';
  const dots = coords
    .map(
      (c) => `
      <circle cx="${c.x.toFixed(1)}" cy="${c.y.toFixed(1)}" r="3.5" fill="#2f6fed" stroke="#ffffff" stroke-width="1"/>
      <text x="${c.x.toFixed(1)}" y="${(c.y - 10).toFixed(1)}" text-anchor="middle" font-family="${TREND_CHART_FONT}" font-size="10" fill="#1f2933">${escapeHtml(formatPointValue(c.value))}</text>`
    )
    .join('');
  const xLabels = points
    .map((p, i) => `<text x="${xFor(i).toFixed(1)}" y="${TREND_CHART_XLABEL_Y}" text-anchor="middle" font-family="${TREND_CHART_FONT}" font-size="11" fill="#6b7684">${escapeHtml(formatMonthLabel(p.year, p.month))}</text>`)
    .join('');

  return (
    `<svg class="revenue-chart-svg" xmlns="http://www.w3.org/2000/svg" width="${w}" height="${TREND_CHART_H}" viewBox="0 0 ${w} ${TREND_CHART_H}" role="img" aria-label="${escapeHtml(ariaLabel)}">` +
    grid +
    line +
    dots +
    xLabels +
    `</svg>`
  );
}

// ---------- 銷貨異常清單（重點商品銷貨項次趨勢圖下方）----------
// 月份選單固定用跟三張趨勢圖一樣的近 12 個月滾動區間，不管這家店這個月有沒有異常紀錄都能選，
// 選到沒有異常的月份就顯示「本月無銷貨異常」，而不是選單裡少一個月讓人誤會資料還沒更新。

function renderSalesAnomalySectionShell() {
  const months = getRollingTwelveMonths();
  const latest = months[months.length - 1];
  const latestValue = latest.year * 100 + latest.month;
  const options = months
    .map(({ year, month }) => {
      const value = year * 100 + month;
      return `<option value="${value}"${value === latestValue ? ' selected' : ''}>${formatMonthLabelCompact(year, month)}</option>`;
    })
    .join('');
  return `
    <h3 class="subsection-title">銷貨異常清單</h3>
    <div class="field">
      <label for="anomaly-month-select">選擇月份</label>
      <select id="anomaly-month-select">${options}</select>
    </div>
    <div id="anomaly-list-container"></div>
  `;
}

function setupSalesAnomalySection(storeName) {
  const select = document.getElementById('anomaly-month-select');
  if (!select) return;
  const render = () => renderAnomalyList(storeName, Number(select.value));
  select.addEventListener('change', render);
  render();
}

function renderAnomalyList(storeName, yearMonth) {
  const container = document.getElementById('anomaly-list-container');
  if (!container) return;

  const entries = (state.salesAnomalyByStore.get(storeName) || new Map()).get(yearMonth) || [];
  if (!entries.length) {
    container.innerHTML = '<p class="empty-state">本月無銷貨異常紀錄。</p>';
    return;
  }

  const rowsHtml = entries
    .map((e) => {
      const gap = Number(e.gap);
      const gapCls = gap > 0 ? 'value-positive' : gap < 0 ? 'value-negative' : '';
      const gapText = (gap > 0 ? '+' : '') + gap.toLocaleString('zh-Hant-TW');
      return `<tr>
        <td class="store-name-cell">${escapeHtml(e.product ?? '—')}</td>
        <td>${Number(e.recentOrder ?? 0).toLocaleString('zh-Hant-TW')}</td>
        <td>${Number(e.standard ?? 0).toLocaleString('zh-Hant-TW')}</td>
        <td class="${gapCls}">${gapText}</td>
      </tr>`;
    })
    .join('');

  container.innerHTML = `
    <div class="table-scroll">
      <table class="nearby-table">
        <thead><tr><th>異常商品</th><th>近三個月叫貨</th><th>近三個月標準(50%)</th><th>差距</th></tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>
  `;
}

// ---------- Utils ----------

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, (c) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;',
  }[c]));
}

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('service-worker.js').catch(() => {});
  });
}
