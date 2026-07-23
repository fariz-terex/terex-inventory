/* =============================================================
   TEREX INVENTORY SYSTEM — app.js
   Multi-customer version with Google Sheets live data
   Depends on: customers.js (loaded before this file)
============================================================= */

'use strict';

/* ===================== STATE ===================== */
const state = {
  customers: {},        // { IPT: { inboundRaw:[], outboundRaw:[], fileName:'', updatedAt:null, source:'excel'|'sheets'|'none' }, ... }
  activeCustomer: null,
  logoDataUrl: null,
};

/* Per-page UI state */
const inboundState  = { search:'', region:'', status:'', sortKey:null, sortDir:'asc', page:1, pageSize:25, lastFiltered:[] };
const outboundState = { search:'', returnFilter:'', sortKey:null, sortDir:'asc', page:1, pageSize:25, lastFiltered:[] };
const stockState    = { search:'', expandedPIC:null };

const chartInstances = { pie:null, topUsed:null, readyMat:null, faultyMat:null };

/* ===================== AUTO-REFRESH ===================== */
const REFRESH_INTERVAL_MS = 5 * 60 * 1000; // 5 menit
let   refreshTimer = null;
const refreshState = {
  isRefreshing: false,
  lastRefresh:  {},   // { IPT: Date, MSG: Date, ... }
};

/* ===================== STORAGE ===================== */
const STORAGE_PREFIX = 'terex_inv_v2_';

function saveCustomerData(id){
  try{
    const d = state.customers[id];
    if(!d) return;
    localStorage.setItem(STORAGE_PREFIX + id + '_inbound',  JSON.stringify(d.inboundRaw));
    localStorage.setItem(STORAGE_PREFIX + id + '_outbound', JSON.stringify(d.outboundRaw));
    localStorage.setItem(STORAGE_PREFIX + id + '_filename', d.fileName || '');
    localStorage.setItem(STORAGE_PREFIX + id + '_updated',  d.updatedAt ? d.updatedAt.toISOString() : '');
  }catch(e){ console.warn('Storage save failed:', e); }
}

function restoreAllCustomers(){
  try{
    CUSTOMER_LIST.forEach(id => {
      const ib  = localStorage.getItem(STORAGE_PREFIX + id + '_inbound');
      const ob  = localStorage.getItem(STORAGE_PREFIX + id + '_outbound');
      const fn  = localStorage.getItem(STORAGE_PREFIX + id + '_filename');
      const upd = localStorage.getItem(STORAGE_PREFIX + id + '_updated');
      state.customers[id] = {
        inboundRaw:  ib  ? JSON.parse(ib)  : [],
        outboundRaw: ob  ? JSON.parse(ob)  : [],
        fileName:    fn  || '',
        updatedAt:   upd ? new Date(upd) : null,
      };
    });
    const lastActive = localStorage.getItem(STORAGE_PREFIX + 'activeCustomer');
    state.activeCustomer = (lastActive && CUSTOMER_LIST.includes(lastActive)) ? lastActive : CUSTOMER_LIST[0];
    const logo = localStorage.getItem(STORAGE_PREFIX + 'logo');
    if(logo) state.logoDataUrl = logo;
  }catch(e){ console.warn('Storage restore failed:', e); }
}

function saveLogoToStorage(){
  try{ if(state.logoDataUrl) localStorage.setItem(STORAGE_PREFIX + 'logo', state.logoDataUrl); }catch(e){}
}

function resetCustomerData(id){
  state.customers[id] = { inboundRaw:[], outboundRaw:[], fileName:'', updatedAt:null };
  try{
    ['inbound','outbound','filename','updated'].forEach(k =>
      localStorage.removeItem(STORAGE_PREFIX + id + '_' + k));
  }catch(e){}
}

/* ===================== GOOGLE SHEETS INTEGRATION ===================== */

/* Jika ada nama kolom yang duplikat (mis. 2x "SOH"), beri suffix _2, _3, dst
   supaya tidak saling menimpa saat dibentuk jadi object. */
function dedupeHeaders(headers){
  const seen = {};
  return headers.map(h => {
    if(!h) return h;
    if(!(h in seen)){ seen[h] = 1; return h; }
    seen[h]++;
    return `${h}_${seen[h]}`;
  });
}

/* Parse CSV text dari Google Sheets publish URL.
   Returns array of objects dengan lowercase keys (konsisten dengan sheetToObjects).
   Auto-detects header row: skip title rows and empty rows. */
function parseCSV(csvText){
  const lines = csvText.split('\n').map(l => l.replace(/\r$/, ''));
  if(lines.length < 2) return [];

  function splitCSVLine(line){
    const result = []; let cur = ''; let inQuote = false;
    for(let i = 0; i < line.length; i++){
      const ch = line[i];
      if(ch === '"'){
        if(inQuote && line[i+1] === '"'){ cur += '"'; i++; }
        else inQuote = !inQuote;
      } else if(ch === ',' && !inQuote){
        result.push(cur); cur = '';
      } else { cur += ch; }
    }
    result.push(cur);
    return result;
  }

  /* Auto-detect header row:
     - Skip baris yang semua selnya kosong
     - Skip baris judul (hanya 1 sel terisi dari banyak kolom)
     - Skip baris TOTAL
     - Ambil baris pertama yang punya >= 3 sel terisi sebagai header */
  let headerIdx = -1;
  let maxCols = 0;

  /* First pass: find max columns to determine "real" rows vs title rows */
  for(let i = 0; i < Math.min(lines.length, 20); i++){
    const cells = splitCSVLine(lines[i]);
    if(cells.length > maxCols) maxCols = cells.length;
  }

  for(let i = 0; i < Math.min(lines.length, 20); i++){
    const cells = splitCSVLine(lines[i]);
    const filled = cells.filter(c => c.trim());
    if(filled.length === 0) continue;                          // baris kosong
    const first = filled[0].trim().toUpperCase();
    if(first === 'TOTAL') continue;                            // baris total
    /* Baris judul: hanya 1-2 sel terisi padahal ada banyak kolom */
    if(filled.length <= 2 && maxCols > 5) continue;
    /* Baris marker (mis. "WAJIB DIISI" diulang di beberapa sel): semua sel
       terisi punya nilai yang sama persis -> bukan header sungguhan */
    const uniqueFilled = new Set(filled.map(c => c.trim().toUpperCase()));
    if(uniqueFilled.size === 1 && filled.length > 1) continue;
    /* Ini header row */
    headerIdx = i;
    break;
  }

  if(headerIdx === -1) return [];

  const headers = dedupeHeaders(splitCSVLine(lines[headerIdx]).map(h => h.trim().toLowerCase()));
  const out = [];
  for(let i = headerIdx + 1; i < lines.length; i++){
    if(!lines[i].trim()) continue;
    const vals = splitCSVLine(lines[i]);
    if(vals.every(v => !v.trim())) continue;
    const obj = {};
    headers.forEach((h, idx) => { if(h) obj[h] = (vals[idx] ?? '').trim(); });
    out.push(obj);
  }
  return out;
}

/* Fetch satu sheet CSV dari Google Sheets */
async function fetchSheetCSV(url){
  const res = await fetch(url, { cache: 'no-store' });
  if(!res.ok) throw new Error(`HTTP ${res.status}`);
  const text = await res.text();
  /* Google Sheets kadang return HTML error page jika URL salah */
  if(text.trim().startsWith('<!')) throw new Error('URL tidak valid atau sheet belum di-publish');

  /* DEBUG — log raw CSV ke console untuk troubleshooting */
  const lines = text.split('\n').filter(l => l.trim());
  console.log('[TEREX Debug] CSV fetch OK:', url.slice(-30));
  console.log('[TEREX Debug] Total lines:', lines.length);
  console.log('[TEREX Debug] Row 1 (header):', lines[0]);
  console.log('[TEREX Debug] Row 2:', lines[1] || '(kosong)');
  console.log('[TEREX Debug] Row 3:', lines[2] || '(kosong)');

  const parsed = parseCSV(text);
  console.log('[TEREX Debug] Parsed rows:', parsed.length);
  if(parsed.length > 0){
    console.log('[TEREX Debug] ALL column keys:', JSON.stringify(Object.keys(parsed[0])));
    console.log('[TEREX Debug] First row values:', JSON.stringify(parsed[0]));
  }
  return parsed;
}

/* Fetch inbound + outbound untuk satu customer dari Google Sheets.
   Returns { inbound: [], outbound: [], ok: true } atau { ok: false, error: '' } */
async function fetchCustomerSheets(id){
  const sc = CUSTOMERS[id];
  if(!sc || !sc.sheets) return { ok:false, error:'Tidak ada konfigurasi sheets' };

  const { inbound: urlIn, outbound: urlOut } = sc.sheets;
  const hasIn  = urlIn  && urlIn.trim();
  const hasOut = urlOut && urlOut.trim();

  if(!hasIn && !hasOut) return { ok:false, error:'URL Google Sheets belum dikonfigurasi' };

  const results = { inbound:null, outbound:null, ok:true, errors:[] };

  await Promise.allSettled([
    hasIn  ? fetchSheetCSV(urlIn).then(d  => { results.inbound  = d; })
                                  .catch(e => { results.errors.push('Inbound: ' + e.message); })
           : Promise.resolve(),
    hasOut ? fetchSheetCSV(urlOut).then(d => { results.outbound = d; })
                                   .catch(e => { results.errors.push('Outbound: ' + e.message); })
           : Promise.resolve(),
  ]);

  if(results.errors.length > 0 && !results.inbound && !results.outbound){
    return { ok:false, error: results.errors.join('; ') };
  }
  return results;
}

/* Refresh satu customer dari Google Sheets, update state, re-render jika aktif */
async function refreshCustomer(id, silent = false){
  const sc = CUSTOMERS[id];
  if(!sc?.sheets?.inbound && !sc?.sheets?.outbound) return;

  const d = state.customers[id];
  if(!silent) setRefreshIndicator('loading');

  const result = await fetchCustomerSheets(id);

  if(!result.ok){
    if(!silent) showToast(`[${id}] Gagal fetch dari Google Sheets: ${result.error}`, 'error');
    setRefreshIndicator('error');
    return;
  }

  let changed = false;
  if(result.inbound  !== null){ d.inboundRaw  = sc.filterInboundRows  ? sc.filterInboundRows(result.inbound)   : result.inbound;  changed = true; }
  if(result.outbound !== null){ d.outboundRaw = sc.filterOutboundRows ? sc.filterOutboundRows(result.outbound) : result.outbound; changed = true; }

  if(changed){
    d.updatedAt = new Date();
    d.source    = 'sheets';
    d.fileName  = 'Google Sheets';
    refreshState.lastRefresh[id] = new Date();
    /* Jangan overwrite localStorage dengan data sheets — biarkan Excel backup tetap ada */
    if(id === state.activeCustomer){
      populateInboundFilters();
      renderAll();
    }
    updateSourceStatus();
    if(!silent) showToast(`[${id}] Data diperbarui dari Google Sheets ✓`);
  }

  setRefreshIndicator('ok');
}

/* Refresh semua customer yang punya URL sheets terkonfigurasi */
async function refreshAllSheets(silent = false){
  if(refreshState.isRefreshing) return;
  refreshState.isRefreshing = true;
  setRefreshIndicator('loading');

  const tasks = CUSTOMER_LIST
    .filter(id => {
      const sc = CUSTOMERS[id];
      return sc?.sheets?.inbound || sc?.sheets?.outbound;
    })
    .map(id => refreshCustomer(id, true));

  await Promise.allSettled(tasks);
  refreshState.isRefreshing = false;
  setRefreshIndicator('ok');
  updateLastRefreshLabel();
  if(!silent) showToast('Semua data berhasil diperbarui dari Google Sheets ✓');
}

/* ---- Auto-refresh timer ---- */
function startAutoRefresh(){
  stopAutoRefresh();
  refreshAllSheets(true); // immediate first fetch
  refreshTimer = setInterval(() => refreshAllSheets(true), REFRESH_INTERVAL_MS);
  updateCountdown();
}

function stopAutoRefresh(){
  if(refreshTimer){ clearInterval(refreshTimer); refreshTimer = null; }
}

let countdownTimer = null;
function updateCountdown(){
  if(countdownTimer) clearInterval(countdownTimer);
  let remaining = REFRESH_INTERVAL_MS / 1000;
  const el = document.getElementById('refreshCountdown');
  countdownTimer = setInterval(() => {
    remaining--;
    if(remaining <= 0){ remaining = REFRESH_INTERVAL_MS / 1000; }
    if(el) el.textContent = formatCountdown(remaining);
  }, 1000);
}

function formatCountdown(sec){
  const m = Math.floor(sec / 60);
  const s = sec % 60;
  return `${m}:${String(s).padStart(2,'0')}`;
}

function updateLastRefreshLabel(){
  const el = document.getElementById('lastRefreshTime');
  if(!el) return;
  const id = state.activeCustomer;
  const t  = refreshState.lastRefresh[id];
  el.textContent = t ? t.toLocaleTimeString('id-ID',{hour:'2-digit',minute:'2-digit',second:'2-digit'}) : '—';
}

function setRefreshIndicator(status){
  const btn = document.getElementById('manualRefreshBtn');
  const dot = document.getElementById('sheetsDot');
  if(!btn) return;
  if(status === 'loading'){
    btn.disabled = true;
    btn.classList.add('refreshing');
    if(dot){ dot.className = 'sheets-dot loading'; dot.title = 'Mengambil data...'; }
  } else if(status === 'error'){
    btn.disabled = false;
    btn.classList.remove('refreshing');
    if(dot){ dot.className = 'sheets-dot error'; dot.title = 'Gagal fetch Google Sheets'; }
  } else {
    btn.disabled = false;
    btn.classList.remove('refreshing');
    if(dot){ dot.className = 'sheets-dot ok'; dot.title = 'Terhubung ke Google Sheets'; }
  }
  updateLastRefreshLabel();
}

/* Check apakah customer aktif sudah punya Google Sheets terkonfigurasi */
function activeCustomerHasSheets(){
  const sc = schema();
  return !!(sc?.sheets?.inbound || sc?.sheets?.outbound);
}

/* ===================== HELPERS ===================== */
function escapeHtml(v){
  if(v === null || v === undefined) return '';
  return String(v)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function parseQty(v){
  if(v === null || v === undefined || v === '') return 0;
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g,''));
  return isNaN(n) ? 0 : n;
}

function formatNumber(n){ return (n||0).toLocaleString('id-ID'); }

function formatDateVal(v){
  if(!v) return '';
  if(v instanceof Date) return v.toLocaleDateString('id-ID');
  if(typeof v === 'number'){
    // Excel serial date
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return d.toLocaleDateString('id-ID');
  }
  const s = String(v);
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : d.toLocaleDateString('id-ID');
}

function debounce(fn, ms){
  let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
}

function schema(){ return CUSTOMERS[state.activeCustomer]; }
function cdata(){ return state.customers[state.activeCustomer]; }

/* ===================== EXCEL PARSING ===================== */
function normalizeHeaders(row){
  return row.map(h => String(h ?? '').trim().toLowerCase());
}

function sheetToObjects(ws){
  const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
  if(!rows || rows.length === 0) return [];

  /* Auto-detect header row: skip blank rows, formula rows, title/metadata rows */
  let headerIdx = 0;
  for(let i = 0; i < Math.min(rows.length, 12); i++){
    const row = rows[i];
    const cells = row.filter(c => c !== '' && c !== null && c !== undefined);
    if(cells.length === 0) continue;
    const first = String(cells[0]).trim();
    if(first.startsWith('=')) continue;
    if(first.toUpperCase() === 'TOTAL') continue;
    /* Skip rows that look like title/info rows (only 1-2 filled cells out of many cols) */
    if(cells.length <= 2 && rows[i].length > 5) continue;
    /* Skip marker rows where every filled cell repeats the same value (e.g. "WAJIB DIISI") */
    const uniqueCells = new Set(cells.map(c => String(c).trim().toUpperCase()));
    if(uniqueCells.size === 1 && cells.length > 1) continue;
    headerIdx = i;
    break;
  }

  const rawHeaders = rows[headerIdx];
  const headers = dedupeHeaders(normalizeHeaders(rawHeaders));
  const out = [];
  for(let i = headerIdx + 1; i < rows.length; i++){
    const r = rows[i];
    if(!r || r.every(c => c === '' || c === null || c === undefined)) continue;
    const obj = {};
    headers.forEach((h, idx) => { if(h) obj[h] = r[idx] !== undefined ? r[idx] : ''; });
    out.push(obj);
  }
  return out;
}

function findSheet(workbook, nameTarget){
  const target = nameTarget.toLowerCase().trim();
  const found = workbook.SheetNames.find(n => n.toLowerCase().trim() === target);
  return found ? workbook.Sheets[found] : null;
}

function handleExcelFile(file){
  const reader = new FileReader();
  reader.onload = e => {
    try{
      const sc = schema();
      const workbook = XLSX.read(e.target.result, { type:'array', cellDates:true });
      const inboundSheet  = findSheet(workbook, sc.sheetInbound);
      const outboundSheet = findSheet(workbook, sc.sheetOutbound);

      if(!inboundSheet && !outboundSheet){
        const available = workbook.SheetNames.join('", "');
        alert(`Sheet "${sc.sheetInbound}" / "${sc.sheetOutbound}" tidak ditemukan.\n\nSheet tersedia: "${available}"`);
        return;
      }

      const d = cdata();
      const imported = [];
      if(inboundSheet){
        let rows = sheetToObjects(inboundSheet);
        if(sc.filterInboundRows) rows = sc.filterInboundRows(rows);
        d.inboundRaw = rows; imported.push('Inbound');
      }
      if(outboundSheet){
        let rows = sheetToObjects(outboundSheet);
        if(sc.filterOutboundRows) rows = sc.filterOutboundRows(rows);
        d.outboundRaw = rows; imported.push('Outbound');
      }
      d.fileName  = file.name;
      d.updatedAt = new Date();

      saveCustomerData(state.activeCustomer);
      renderAll();
      showToast(`[${state.activeCustomer}] ${file.name} — ${imported.join(' & ')} ter-update.`);
    }catch(err){
      console.error(err);
      alert('Gagal membaca file Excel: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ===================== STATUS BADGE ===================== */
function statusBadgeClass(sc, status){
  if(sc.statusReady(status))  return 'badge badge-ready';
  if(sc.statusFaulty(status)) return 'badge badge-faulty';
  return 'badge badge-pending';
}

function statusBucketFromSchema(sc, status){
  if(sc.statusReady(status))  return 'ready';
  if(sc.statusFaulty(status)) return 'faulty';
  return 'pending';
}

/* ===================== DATA PROCESSING ===================== */
function processData(rows, { searchTerm, searchKeys, sortKey, sortDir }){
  let data = [...rows];
  if(searchTerm){
    const t = searchTerm.toLowerCase();
    data = data.filter(r => searchKeys.some(k => String(r[k]||'').toLowerCase().includes(t)));
  }
  if(sortKey){
    data.sort((a,b) => {
      const av = String(a[sortKey]||''); const bv = String(b[sortKey]||'');
      const n = parseFloat(av) - parseFloat(bv);
      const cmp = isNaN(n) ? av.localeCompare(bv,'id') : n;
      return sortDir === 'asc' ? cmp : -cmp;
    });
  }
  return data;
}

function paginate(data, page, size){
  const total = Math.max(1, Math.ceil(data.length / size));
  const safe  = Math.min(Math.max(1, page), total);
  return { pageData: data.slice((safe-1)*size, safe*size), totalPages: total, safePage: safe };
}

/* ===================== TABLE RENDERING ===================== */
function cellHtml(col, row){
  const raw = row[col.key];
  if(col.badge){
    const sc = schema();
    const cls = statusBadgeClass(sc, raw);
    return `<td><span class="${cls}">${escapeHtml(raw)}</span></td>`;
  }
  if(col.date) return `<td>${escapeHtml(formatDateVal(raw))}</td>`;
  if(col.num)  return `<td class="num">${formatNumber(parseQty(raw))}</td>`;
  if(col.mono) return `<td class="mono">${escapeHtml(raw)}</td>`;
  return `<td>${escapeHtml(raw)}</td>`;
}

function renderTable(tableId, bodyId, columns, rows){
  /* Headers */
  const table = document.getElementById(tableId);
  const thead = table.querySelector('thead tr');
  thead.innerHTML = columns.map(c =>
    `<th data-key="${c.key}"${c.num?' class="num"':''}>${escapeHtml(c.label)}<span class="sort-arrow">↕</span></th>`
  ).join('');

  /* Body */
  const tbody = document.getElementById(bodyId);
  if(rows.length === 0){
    tbody.innerHTML = `<tr><td colspan="${columns.length}" class="empty-row">Tidak ada data.</td></tr>`;
    return;
  }
  tbody.innerHTML = rows.map(row => `<tr>${columns.map(c => cellHtml(c, row)).join('')}</tr>`).join('');
}

function renderPagination(containerId, currentPage, totalPages, onPage){
  const el = document.getElementById(containerId);
  if(!el) return;
  if(totalPages <= 1){ el.innerHTML = ''; return; }
  const pages = [];
  const addBtn = (p, label, active, disabled) =>
    `<button ${disabled?'disabled':''} ${active?'class="active"':''} data-p="${p}">${label}</button>`;
  pages.push(addBtn(currentPage-1,'‹', false, currentPage===1));
  const range = new Set([1, totalPages, currentPage-1, currentPage, currentPage+1].filter(p=>p>=1&&p<=totalPages));
  let prev = 0;
  Array.from(range).sort((a,b)=>a-b).forEach(p=>{
    if(prev && p-prev>1) pages.push(`<span class="dots">…</span>`);
    pages.push(addBtn(p, p, p===currentPage, false));
    prev = p;
  });
  pages.push(addBtn(currentPage+1,'›', false, currentPage===totalPages));
  el.innerHTML = pages.join('');
  el.querySelectorAll('button:not([disabled])').forEach(btn =>
    btn.addEventListener('click', () => onPage(parseInt(btn.dataset.p)))
  );
}

function setupSortableTable(tableId, uiState, renderFn){
  const table = document.getElementById(tableId);
  if(!table) return;
  table.querySelector('thead').addEventListener('click', e => {
    const th = e.target.closest('th[data-key]');
    if(!th) return;
    const key = th.dataset.key;
    if(uiState.sortKey === key) uiState.sortDir = uiState.sortDir==='asc'?'desc':'asc';
    else { uiState.sortKey = key; uiState.sortDir = 'asc'; }
    uiState.page = 1;
    renderFn();
  });
}

/* ===================== INBOUND PAGE ===================== */
function renderInboundTable(){
  const sc = schema();
  const d  = cdata();
  const cols = sc.inboundColumns;

  let filtered = processData(d.inboundRaw, {
    searchTerm:  inboundState.search,
    searchKeys:  sc.searchKeysInbound,
    sortKey:     inboundState.sortKey,
    sortDir:     inboundState.sortDir,
  });

  if(inboundState.region && sc.inboundFilterRegionField){
    filtered = filtered.filter(r => String(r[sc.inboundFilterRegionField]||'') === inboundState.region);
  }
  if(inboundState.status){
    const statusKey = sc.map.status || 'status';
    filtered = filtered.filter(r => String(r[statusKey]||'') === inboundState.status);
  }

  inboundState.lastFiltered = filtered;
  const { pageData, totalPages, safePage } = paginate(filtered, inboundState.page, inboundState.pageSize);
  inboundState.page = safePage;

  renderTable('inboundTable', 'inboundBody', cols, pageData);
  renderPagination('inboundPagination', safePage, totalPages, p => { inboundState.page = p; renderInboundTable(); });
  updateTableInfo('inboundTableInfo', filtered.length, d.inboundRaw.length);
}

function populateInboundFilters(){
  const sc = schema();
  const d  = cdata();

  /* Region filter */
  const regionEl = document.getElementById('inboundFilterRegion');
  if(regionEl){
    if(sc.inboundFilterRegionField){
      regionEl.style.display = '';
      const vals = [...new Set(d.inboundRaw.map(r => String(r[sc.inboundFilterRegionField]||'').trim()).filter(Boolean))].sort();
      regionEl.innerHTML = `<option value="">Semua Region</option>` + vals.map(v=>`<option>${escapeHtml(v)}</option>`).join('');
    } else {
      regionEl.style.display = 'none';
    }
  }

  /* Status filter */
  const statusEl = document.getElementById('inboundFilterStatus');
  if(statusEl){
    const statusKey = sc.map.status || 'status';
    const vals = [...new Set(d.inboundRaw.map(r => String(r[statusKey]||'').trim()).filter(Boolean))].sort();
    statusEl.innerHTML = `<option value="">Semua Status</option>` + vals.map(v=>`<option>${escapeHtml(v)}</option>`).join('');
  }
}

/* ===================== OUTBOUND PAGE ===================== */
function renderOutboundTable(){
  const sc = schema();
  const d  = cdata();
  const cols = sc.outboundColumns;

  let filtered = processData(d.outboundRaw, {
    searchTerm:  outboundState.search,
    searchKeys:  sc.searchKeysOutbound,
    sortKey:     outboundState.sortKey,
    sortDir:     outboundState.sortDir,
  });

  if(outboundState.returnFilter === 'belum'){
    filtered = filtered.filter(r => sc.isUnreturned(r));
  } else if(outboundState.returnFilter === 'sudah'){
    filtered = filtered.filter(r => !sc.isUnreturned(r));
  }

  outboundState.lastFiltered = filtered;
  const { pageData, totalPages, safePage } = paginate(filtered, outboundState.page, outboundState.pageSize);
  outboundState.page = safePage;

  renderTable('outboundTable', 'outboundBody', cols, pageData);
  renderPagination('outboundPagination', safePage, totalPages, p => { outboundState.page = p; renderOutboundTable(); });
  updateTableInfo('outboundTableInfo', filtered.length, d.outboundRaw.length);
}

/* ===================== STOCK TEKNISI / STOCK PER SITE ===================== */
function getStockSummary(){
  const sc = schema();
  if(sc.stockMode === 'site') return getStockSummarySite();
  return getStockSummaryTechnician();
}

function getStockSummaryTechnician(){
  const sc = schema();
  const d  = cdata();
  const map = new Map();

  const getEntry = name => {
    const key = name.toLowerCase();
    for(const [k,v] of map.entries()){ if(k.toLowerCase()===key) return v; }
    const e = { pic:name, ready:0, faulty:0, pending:0, total:0, materials:new Map() };
    map.set(name, e); return e;
  };

  /* Build whitelist of technician names from Outbound PIC */
  const techSet = new Set();
  d.outboundRaw.forEach(row => {
    const pic = String(row[sc.map.pic || 'pic'] || '').trim();
    if(sc.isTechnicianPIC(pic)) techSet.add(pic.toLowerCase());
  });

  /* For MSG: PIC EOS in Inbound = technician holding the item */
  const inboundTechField = sc.inboundTechLocField;

  /* READY from Inbound: rows where the tech-location field matches a known tech */
  d.inboundRaw.forEach(row => {
    const lokasi = String(row[inboundTechField] || '').trim();
    if(!lokasi || !techSet.has(lokasi.toLowerCase())) return;
    const statusKey = sc.map.status || 'status';
    if(!sc.statusReady(row[statusKey])) return;

    const qty = parseQty(row[sc.map.qty || 'qty']);
    const entry = getEntry(lokasi);
    entry.ready += qty;
    entry.total += qty;
    const mat = String(row[sc.map.materialName || 'material name'] || '').trim() || '(Tanpa Nama)';
    entry.materials.set(mat, (entry.materials.get(mat)||0) + qty);
  });

  /* FAULTY/PENDING from Outbound: unreturned rows per technician */
  d.outboundRaw.forEach(row => {
    const picKey = sc.map.pic || 'pic';
    const picRaw = String(row[picKey] || '').trim();
    if(!sc.isTechnicianPIC(picRaw)) return;
    if(!sc.isUnreturned(row)) return;

    const qty = parseQty(row[sc.map.qty || 'qty']);
    const statusKey = (sc.map.statusOut || sc.map.status || 'status');
    const bucket = statusBucketFromSchema(sc, row[statusKey]);
    const entry = getEntry(picRaw);
    if(bucket === 'faulty') entry.faulty += qty;
    else entry.pending += qty;
    entry.total += qty;
    const mat = String(row[sc.map.materialName || 'material name'] || '').trim() || '(Tanpa Nama)';
    entry.materials.set(mat, (entry.materials.get(mat)||0) + qty);
  });

  return Array.from(map.values())
    .filter(e => e.total > 0)
    .map(e => ({
      pic:e.pic, ready:e.ready, faulty:e.faulty, pending:e.pending, total:e.total,
      materials: Array.from(e.materials.entries()).map(([name,qty])=>({name,qty})).sort((a,b)=>b.qty-a.qty)
    }));
}

/* Stock per Site/Cluster (mis. PIM) — tidak ada konsep PIC/teknisi.
   Dikelompokkan berdasarkan sc.stockGroupField, dihitung dari Outbound
   (unit yang sudah terpasang / Replacement), dengan status approval BAK/BAS. */
function getStockSummarySite(){
  const sc = schema();
  const d  = cdata();
  const groupField = sc.stockGroupField || 'cluster';
  const qtyKey = sc.map.qty || 'qty';
  const map = new Map();

  const getEntry = name => {
    if(!map.has(name)) map.set(name, { group:name, total:0, bakDone:0, bakPending:0, basDone:0, basPending:0, rows:[] });
    return map.get(name);
  };

  const isBakDone = v => String(v||'').trim().toUpperCase() === 'CLOSE';
  const isBasDone = v => String(v||'').trim().toUpperCase() === 'CLOSE';

  d.outboundRaw.forEach(row => {
    const groupVal = String(row[groupField] || '').trim() || '(Tanpa Cluster)';
    const qty = parseQty(row[qtyKey]) || 1;
    const entry = getEntry(groupVal);
    entry.total += qty;
    if(isBakDone(row['bak status'])) entry.bakDone += qty; else entry.bakPending += qty;
    if(isBasDone(row['bas status'])) entry.basDone += qty; else entry.basPending += qty;
    entry.rows.push(row);
  });

  return Array.from(map.values());
}

function getOnSearchMaterials(){
  const sc = schema();
  const d  = cdata();
  /* "On Search" = nilai khusus di kolom PIC yang menandakan material tidak diketahui lokasi */
  const ON_SEARCH_VALUES = ['on search'];
  const picKey = sc.map.pic || 'pic';
  const map = new Map();
  d.outboundRaw.forEach(row => {
    const pic = String(row[picKey]||'').trim().toLowerCase();
    if(!ON_SEARCH_VALUES.includes(pic)) return;
    const nameKey = sc.map.materialName || 'material name';
    const codeKey = sc.map.itemCode || 'item code';
    const name = String(row[nameKey]||'').trim() || '(Tanpa Nama)';
    const code = String(row[codeKey]||'').trim();
    const key  = (code||name).toLowerCase();
    const qty  = parseQty(row[sc.map.qty||'qty']);
    if(!map.has(key)) map.set(key, {name, code, qty:0, count:0});
    const e = map.get(key);
    e.qty += qty; e.count++;
  });
  return Array.from(map.values()).sort((a,b)=>b.qty-a.qty);
}

function renderStockSummary(){
  const sc = schema();
  if(sc.stockMode === 'site') return renderStockSummarySite();
  return renderStockSummaryTechnician();
}

function renderStockDetail(){
  const sc = schema();
  if(sc.stockMode === 'site') return renderStockDetailSite();
  return renderStockDetailTechnician();
}

function renderStockSummaryTechnician(){
  let summary = getStockSummaryTechnician();
  if(stockState.search){
    const t = stockState.search.toLowerCase();
    summary = summary.filter(s => s.pic.toLowerCase().includes(t));
  }
  summary.sort((a,b) => a.pic.localeCompare(b.pic,'id'));

  setStockTableHead(['PIC','Ready','Faulty','Pending','Total'], [false,true,true,true,true]);

  const body = document.getElementById('stockSummaryBody');
  const d = cdata();
  if(d.outboundRaw.length === 0){
    body.innerHTML = `<tr><td colspan="5" class="empty-row">Belum ada data.</td></tr>`; return;
  }
  if(summary.length === 0){
    body.innerHTML = `<tr><td colspan="5" class="empty-row">Tidak ada teknisi yang cocok.</td></tr>`; return;
  }
  body.innerHTML = summary.map(s => `
    <tr data-key="${escapeHtml(s.pic)}" class="${stockState.expandedPIC===s.pic?'active-row':''}">
      <td>${escapeHtml(s.pic)}</td>
      <td class="num">${formatNumber(s.ready)}</td>
      <td class="num">${formatNumber(s.faulty)}</td>
      <td class="num">${formatNumber(s.pending)}</td>
      <td class="num">${formatNumber(s.total)}</td>
    </tr>`).join('');
  body.querySelectorAll('tr').forEach(tr =>
    tr.addEventListener('click', () => {
      stockState.expandedPIC = stockState.expandedPIC===tr.dataset.key ? null : tr.dataset.key;
      renderStockSummary(); renderStockDetail();
    })
  );
}

function renderStockDetailTechnician(){
  const panel = document.getElementById('stockDetailPanel');
  if(!stockState.expandedPIC){ panel.style.display='none'; return; }
  panel.style.display='block';
  document.getElementById('stockDetailName').textContent = stockState.expandedPIC;

  const sc = schema(); const d = cdata();
  const picKey = sc.map.pic || 'pic';
  const rows = d.outboundRaw.filter(r => String(r[picKey]||'').trim() === stockState.expandedPIC);

  /* Material breakdown */
  const matMap = new Map();
  rows.forEach(r => {
    const name = String(r[sc.map.materialName||'material name']||'').trim() || '(Tanpa Nama)';
    matMap.set(name, (matMap.get(name)||0) + parseQty(r[sc.map.qty||'qty']));
  });
  const matList = Array.from(matMap.entries()).sort((a,b)=>b[1]-a[1]);
  const mbody = document.getElementById('stockMaterialBreakdownBody');
  mbody.innerHTML = matList.length
    ? matList.map(([name,qty]) => `<tr><td>${escapeHtml(name)}</td><td class="num">${formatNumber(qty)}</td></tr>`).join('')
    : `<tr><td colspan="2" class="empty-row">Tidak ada material.</td></tr>`;

  /* Detail rows */
  setStockDetailTableHead([
    {label:'Material Name'}, {label:'Item Code'}, {label:'Serial Number'}, {label:'Qty', num:true}, {label:'Status'}
  ]);
  const nameKey   = sc.map.materialName || 'material name';
  const codeKey   = sc.map.itemCode || 'item code';
  const serialKey = sc.map.serialNumber || 'serial number';
  const qtyKey    = sc.map.qty || 'qty';
  const statusKey = sc.map.statusOut || sc.map.status || 'status';
  document.getElementById('stockDetailBody').innerHTML = rows.map(r => `<tr>
    <td>${escapeHtml(r[nameKey])}</td>
    <td class="mono">${escapeHtml(r[codeKey])}</td>
    <td class="mono">${escapeHtml(r[serialKey])}</td>
    <td class="num">${formatNumber(parseQty(r[qtyKey]))}</td>
    <td><span class="${statusBadgeClass(sc,r[statusKey])}">${escapeHtml(r[statusKey])}</span></td>
  </tr>`).join('');
}

/* ---- Stock per Site/Cluster (PIM) ---- */
function renderStockSummarySite(){
  const sc = schema();
  let summary = getStockSummarySite();
  if(stockState.search){
    const t = stockState.search.toLowerCase();
    summary = summary.filter(s => s.group.toLowerCase().includes(t));
  }
  summary.sort((a,b) => a.group.localeCompare(b.group,'id'));

  setStockTableHead(
    [sc.stockGroupLabel || 'Cluster', 'Total Terpasang', 'BAK Selesai', 'BAK Pending', 'BAS Selesai', 'BAS Pending'],
    [false, true, true, true, true, true]
  );

  const body = document.getElementById('stockSummaryBody');
  const d = cdata();
  if(d.outboundRaw.length === 0){
    body.innerHTML = `<tr><td colspan="6" class="empty-row">Belum ada data.</td></tr>`; return;
  }
  if(summary.length === 0){
    body.innerHTML = `<tr><td colspan="6" class="empty-row">Tidak ada cluster yang cocok.</td></tr>`; return;
  }
  body.innerHTML = summary.map(s => `
    <tr data-key="${escapeHtml(s.group)}" class="${stockState.expandedPIC===s.group?'active-row':''}">
      <td>${escapeHtml(s.group)}</td>
      <td class="num">${formatNumber(s.total)}</td>
      <td class="num">${formatNumber(s.bakDone)}</td>
      <td class="num">${formatNumber(s.bakPending)}</td>
      <td class="num">${formatNumber(s.basDone)}</td>
      <td class="num">${formatNumber(s.basPending)}</td>
    </tr>`).join('');
  body.querySelectorAll('tr').forEach(tr =>
    tr.addEventListener('click', () => {
      stockState.expandedPIC = stockState.expandedPIC===tr.dataset.key ? null : tr.dataset.key;
      renderStockSummary(); renderStockDetail();
    })
  );
}

function renderStockDetailSite(){
  const panel = document.getElementById('stockDetailPanel');
  if(!stockState.expandedPIC){ panel.style.display='none'; return; }
  panel.style.display='block';
  document.getElementById('stockDetailName').textContent = stockState.expandedPIC;

  const sc = schema();
  const groupField = sc.stockGroupField || 'cluster';
  const summary = getStockSummarySite();
  const entry = summary.find(s => s.group === stockState.expandedPIC);
  const rows = entry ? entry.rows : [];

  /* Material breakdown */
  const matMap = new Map();
  rows.forEach(r => {
    const name = String(r[sc.map.materialName||'material'] || '').trim() || '(Tanpa Nama)';
    matMap.set(name, (matMap.get(name)||0) + (parseQty(r[sc.map.qty||'qty shipment']) || 1));
  });
  const matList = Array.from(matMap.entries()).sort((a,b)=>b[1]-a[1]);
  const mbody = document.getElementById('stockMaterialBreakdownBody');
  mbody.innerHTML = matList.length
    ? matList.map(([name,qty]) => `<tr><td>${escapeHtml(name)}</td><td class="num">${formatNumber(qty)}</td></tr>`).join('')
    : `<tr><td colspan="2" class="empty-row">Tidak ada material.</td></tr>`;

  /* Detail rows — pakai stockDetailColumns dari schema kalau ada */
  const cols = sc.stockDetailColumns || [
    {key:sc.map.materialName||'material', label:'Material'},
    {key:sc.map.itemCode||'item code', label:'Item Code', mono:true},
    {key:sc.map.serialNumber||'s/n after', label:'Serial Number', mono:true},
    {key:sc.map.qty||'qty shipment', label:'Qty', num:true},
  ];
  setStockDetailTableHead(cols);
  document.getElementById('stockDetailBody').innerHTML = rows.map(r => `<tr>${
    cols.map(c => {
      const raw = r[c.key];
      if(c.badge) return `<td><span class="${statusBadgeClass(sc, raw)}">${escapeHtml(raw)}</span></td>`;
      if(c.date)  return `<td>${escapeHtml(formatDateVal(raw))}</td>`;
      if(c.num)   return `<td class="num">${formatNumber(parseQty(raw))}</td>`;
      if(c.mono)  return `<td class="mono">${escapeHtml(raw)}</td>`;
      return `<td>${escapeHtml(raw)}</td>`;
    }).join('')
  }</tr>`).join('');
}

/* ---- Helper: render thead dinamis utk tabel Stock Summary & Stock Detail ---- */
function setStockTableHead(labels, numFlags){
  const thead = document.querySelector('#stockSummaryTable thead tr');
  if(!thead) return;
  thead.innerHTML = labels.map((l,i) => `<th${numFlags[i] ? ' class="num"' : ''}>${escapeHtml(l)}</th>`).join('');
}

function setStockDetailTableHead(cols){
  const thead = document.querySelector('#stockDetailTable thead tr');
  if(!thead) return;
  thead.innerHTML = cols.map(c => `<th${c.num ? ' class="num"' : ''}>${escapeHtml(c.label)}</th>`).join('');
}

/* ===================== DASHBOARD ===================== */
function renderDashboard(){
  const sc = schema(); const d = cdata();
  const inbound = d.inboundRaw; const outbound = d.outboundRaw;

  const nameKey   = sc.map.materialName || 'material name';
  const qtyKey    = sc.map.qty || 'qty';
  const statusKey = sc.map.status || 'status';
  const picKey    = sc.map.pic || 'pic';

  /* KPIs */
  const matSet = new Set();
  let totalQtyIn = 0;
  inbound.forEach(r => {
    const name = String(r[nameKey]||'').trim();
    if(name) matSet.add(name.toLowerCase());
    totalQtyIn += parseQty(r[qtyKey]);
  });

  let totalQtyOut = 0, stockReady = 0;
  inbound.forEach(r => {
    const qty = parseQty(r[qtyKey]);
    if(sc.statusReady(r[statusKey])) stockReady += qty;
    totalQtyOut; // outbound qty counted separately
  });
  outbound.forEach(r => totalQtyOut += parseQty(r[qtyKey]));

  /* Belum return = unreturned outbound */
  let stockFaulty = 0;
  outbound.forEach(r => { if(sc.isUnreturned(r)) stockFaulty += parseQty(r[qtyKey]); });

  let teknisiCount;
  if(sc.computeKpiTeknisi){
    teknisiCount = sc.computeKpiTeknisi(outbound);
  } else {
    const techSet = new Set();
    outbound.forEach(r => { const p = String(r[picKey]||'').trim(); if(sc.isTechnicianPIC(p)) techSet.add(p); });
    teknisiCount = techSet.size;
  }

  document.getElementById('kpiTotalMaterial').textContent = formatNumber(matSet.size);
  document.getElementById('kpiQtyInbound').textContent    = formatNumber(totalQtyIn);
  document.getElementById('kpiQtyOutbound').textContent   = formatNumber(totalQtyOut);
  document.getElementById('kpiStockReady').textContent    = formatNumber(stockReady);
  document.getElementById('kpiStockFaulty').textContent   = formatNumber(stockFaulty);
  document.getElementById('kpiTeknisi').textContent       = formatNumber(teknisiCount);

  renderCharts(inbound, outbound, stockReady, stockFaulty);
  renderReadyMatCard(inbound);
  renderFaultyMatCard(outbound);
}

function renderCharts(inbound, outbound, stockReady, stockFaulty){
  if(typeof Chart === 'undefined') return;
  const sc = schema();
  const qtyKey = sc.map.qty || 'qty';
  const nameKey = sc.map.materialName || 'material name';

  /* Pie: Ready vs Belum Return */
  const pieCtx = document.getElementById('chartReadyFaulty');
  if(chartInstances.pie) chartInstances.pie.destroy();
  chartInstances.pie = new Chart(pieCtx, {
    type:'doughnut',
    data:{ labels:['Ready (WH)','Belum Return'], datasets:[{ data:[stockReady, stockFaulty], backgroundColor:['#1E8E5A','#E31E24'], borderWidth:0 }] },
    options:{ responsive:true, maintainAspectRatio:false, cutout:'62%', plugins:{ legend:{ position:'bottom', labels:{ usePointStyle:true, boxWidth:8, font:{size:11.5} } } } }
  });

  /* Bar: Top 10 most used (outbound qty) */
  const usedMap = new Map();
  outbound.forEach(r => {
    const name = String(r[nameKey]||'').trim() || '(Tanpa Nama)';
    usedMap.set(name, (usedMap.get(name)||0) + parseQty(r[qtyKey]));
  });
  const top10 = Array.from(usedMap.entries()).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const topCtx = document.getElementById('chartTopUsed');
  if(chartInstances.topUsed) chartInstances.topUsed.destroy();
  chartInstances.topUsed = new Chart(topCtx, {
    type:'bar',
    data:{ labels:top10.map(d=>d[0]), datasets:[{ label:'Qty Outbound', data:top10.map(d=>d[1]), backgroundColor: schema().color || '#E31E24', borderRadius:4, maxBarThickness:22 }] },
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}}, scales:{ x:{beginAtZero:true, grid:{color:'#EFEFEF'}, ticks:{precision:0}}, y:{grid:{display:false}} } }
  });
}

/* ---------- HELPER: horizontal bar chart (shared config) ---------- */
function makeHBarChart(ctx, labels, data, color, instance){
  if(instance) instance.destroy();
  return new Chart(ctx, {
    type:'bar',
    data:{ labels, datasets:[{ data, backgroundColor: color, borderRadius:4, maxBarThickness:20 }] },
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false} },
      scales:{
        x:{ beginAtZero:true, grid:{color:'#EFEFEF'}, ticks:{precision:0} },
        y:{ grid:{display:false}, ticks:{ font:{size:11}, color:'#444' } }
      }
    }
  });
}

/* ---------- MATERIAL READY CARD ---------- */
function renderReadyMatCard(inbound){
  const sc = schema();
  const nameKey   = sc.map.materialName || 'material name';
  const qtyKey    = sc.map.qty || 'qty';
  const statusKey = sc.map.status || 'status';

  const map = new Map();
  inbound.forEach(r => {
    if(!sc.statusReady(r[statusKey])) return;
    const name = String(r[nameKey]||'').trim() || '(Tanpa Nama)';
    map.set(name, (map.get(name)||0) + parseQty(r[qtyKey]));
  });

  const sorted = Array.from(map.entries()).sort((a,b)=>b[1]-a[1]);
  const totalQty = sorted.reduce((s,[,q])=>s+q, 0);
  const top = sorted.slice(0,15);

  /* Badge */
  const badge = document.getElementById('readyTotalBadge');
  if(badge) badge.textContent = formatNumber(totalQty) + ' unit';

  /* Chart */
  const ctx = document.getElementById('chartReadyMat');
  if(ctx){
    chartInstances.readyMat = makeHBarChart(
      ctx,
      top.map(d=>d[0]),
      top.map(d=>d[1]),
      '#1E8E5A',
      chartInstances.readyMat
    );
  }

  /* Table */
  const tbody = document.getElementById('readyMatBody');
  if(!tbody) return;
  if(sorted.length === 0){
    tbody.innerHTML = '<tr><td colspan="2" class="empty-row">Tidak ada material Ready.</td></tr>'; return;
  }
  tbody.innerHTML = sorted.map(([name,qty]) =>
    `<tr><td>${escapeHtml(name)}</td><td class="num">${formatNumber(qty)}</td></tr>`
  ).join('');
}

/* ---------- MATERIAL FAULTY / PERLU RETURN CARD ---------- */
function renderFaultyMatCard(outbound){
  const sc = schema();
  const nameKey = sc.map.materialName || 'material name';
  const qtyKey  = sc.map.qty || 'qty';
  const picKey  = sc.map.pic || 'pic';

  /* Aggregate by material — only unreturned rows */
  const matMap = new Map();  // name -> { qty, pics: Set }
  outbound.forEach(r => {
    if(!sc.isUnreturned(r)) return;
    const name = String(r[nameKey]||'').trim() || '(Tanpa Nama)';
    const pic  = String(r[picKey]||'').trim();
    const qty  = parseQty(r[qtyKey]);
    if(!matMap.has(name)) matMap.set(name, {qty:0, pics:new Set()});
    const e = matMap.get(name);
    e.qty += qty;
    if(pic) e.pics.add(pic);
  });

  const sorted = Array.from(matMap.entries()).sort((a,b)=>b[1].qty-a[1].qty);
  const totalQty = sorted.reduce((s,[,e])=>s+e.qty, 0);
  const top = sorted.slice(0,15);

  /* Badge */
  const badge = document.getElementById('faultyTotalBadge');
  if(badge) badge.textContent = formatNumber(totalQty) + ' unit';

  /* Chart */
  const ctx = document.getElementById('chartFaultyMat');
  if(ctx){
    chartInstances.faultyMat = makeHBarChart(
      ctx,
      top.map(d=>d[0]),
      top.map(d=>d[1].qty),
      '#E31E24',
      chartInstances.faultyMat
    );
  }

  /* Table */
  const tbody = document.getElementById('faultyMatBody');
  if(!tbody) return;
  if(sorted.length === 0){
    tbody.innerHTML = '<tr><td colspan="3" class="empty-row">Tidak ada material yang perlu di-return.</td></tr>'; return;
  }
  tbody.innerHTML = sorted.map(([name,e]) =>
    `<tr>
      <td>${escapeHtml(name)}</td>
      <td class="num">${formatNumber(e.qty)}</td>
      <td class="mat-pic">${escapeHtml([...e.pics].join(', '))}</td>
    </tr>`
  ).join('');
}

/* ===================== SOURCE STATUS ===================== */
function updateSourceStatus(){
  const d    = cdata();
  const sc   = schema();
  const dot  = document.querySelector('#sourceStatus .status-dot');
  const text = document.getElementById('sourceStatusText');
  const meta = document.getElementById('sourceMeta');
  const hasData = d.inboundRaw.length > 0 || d.outboundRaw.length > 0;

  dot.classList.toggle('on',  hasData);
  dot.classList.toggle('off', !hasData);
  text.textContent = hasData
    ? `Inbound: ${d.inboundRaw.length} baris · Outbound: ${d.outboundRaw.length} baris`
    : 'Belum ada data';

  /* Source badge: Live (Google Sheets) atau Excel */
  const sourceBadge = document.getElementById('sourceBadge');
  if(sourceBadge){
    const hasSheets = sc?.sheets?.inbound || sc?.sheets?.outbound;
    if(hasSheets && d.source === 'sheets'){
      sourceBadge.textContent = '● Live';
      sourceBadge.className = 'source-badge live';
    } else if(hasSheets){
      sourceBadge.textContent = '○ Belum fetch';
      sourceBadge.className = 'source-badge pending';
    } else {
      sourceBadge.textContent = 'Excel';
      sourceBadge.className = 'source-badge excel';
    }
  }

  if(hasData && d.fileName){
    const dateStr = d.updatedAt
      ? d.updatedAt.toLocaleDateString('id-ID') + ' ' + d.updatedAt.toLocaleTimeString('id-ID',{hour:'2-digit',minute:'2-digit'})
      : '';
    meta.textContent = `${d.fileName}${dateStr?' • '+dateStr:''}`;
  } else { meta.textContent = ''; }

  updateLastRefreshLabel();
}

function updateTableInfo(id, filtered, total){
  const el = document.getElementById(id);
  if(el) el.textContent = filtered < total ? `${formatNumber(filtered)} dari ${formatNumber(total)} baris` : `${formatNumber(total)} baris`;
}

/* ===================== CUSTOMER SWITCHER ===================== */
function switchCustomer(id){
  if(!CUSTOMERS[id]) return;
  state.activeCustomer = id;
  localStorage.setItem(STORAGE_PREFIX + 'activeCustomer', id);

  /* Reset UI state */
  inboundState.search=''; inboundState.region=''; inboundState.status=''; inboundState.page=1; inboundState.sortKey=null;
  outboundState.search=''; outboundState.returnFilter=''; outboundState.page=1; outboundState.sortKey=null;
  stockState.search=''; stockState.expandedPIC=null;

  /* Update customer pill/badge in topbar */
  updateCustomerUI();
  populateInboundFilters();

  /* Reset search inputs */
  ['inboundSearch','outboundSearch','stockSearch'].forEach(id => {
    const el = document.getElementById(id); if(el) el.value='';
  });
  const returnEl = document.getElementById('outboundFilterReturn');
  if(returnEl) returnEl.value='';

  renderAll();
}

function updateCustomerUI(){
  const sc = schema();

  /* Sidebar customer badge */
  const badge = document.getElementById('customerBadge');
  if(badge){
    badge.textContent = sc.name;
    badge.style.background = sc.color;
  }

  /* Dropdown active */
  document.querySelectorAll('.customer-option').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.cid === state.activeCustomer);
  });

  /* Page subtitle */
  document.getElementById('pageSubtitle').textContent =
    `Ringkasan inventory material ${sc.name} — TEREX`;

  /* KPI label for "Belum Return" might vary per customer */
  const faultyLabel = document.getElementById('kpiFaultyLabel');
  if(faultyLabel) faultyLabel.textContent = 'Belum Return ke WH';

  /* KPI label "Total Teknisi" bisa berubah jadi "Total Cluster" (mis. PIM) */
  const teknisiLabel = document.getElementById('kpiTeknisiLabel');
  if(teknisiLabel) teknisiLabel.textContent = sc.kpiTeknisiLabel || 'Total Teknisi';

  /* Placeholder search di halaman Stock (per teknisi vs per cluster) */
  const stockSearchEl = document.getElementById('stockSearch');
  if(stockSearchEl) stockSearchEl.placeholder = sc.stockSearchPlaceholder || 'Cari nama PIC / teknisi...';
}

/* ===================== EXPORT ===================== */
function exportToExcel(rows, columns, filename){
  if(!rows.length){ alert('Tidak ada data untuk diexport.'); return; }
  const headers = columns.map(c => c.label);
  const data = rows.map(r => columns.map(c => {
    const v = r[c.key];
    return c.date ? formatDateVal(v) : (c.num ? parseQty(v) : String(v??''));
  }));
  const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Data');
  XLSX.writeFile(wb, filename);
}

/* ===================== TOAST ===================== */
let toastTimer = null;
function showToast(msg, type = 'info'){
  let el = document.getElementById('appToast');
  if(!el){ el=document.createElement('div'); el.id='appToast'; el.className='app-toast'; document.body.appendChild(el); }
  el.textContent = msg;
  el.className = 'app-toast show' + (type === 'error' ? ' app-toast-error' : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), type === 'error' ? 5000 : 3800);
}

/* ===================== PAGE NAVIGATION ===================== */
let currentPage = 'dashboard';
function showPage(page){
  currentPage = page;
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const pageEl = document.getElementById('page-' + page);
  if(pageEl) pageEl.classList.add('active');
  const navEl = document.querySelector(`.nav-item[data-page="${page}"]`);
  if(navEl) navEl.classList.add('active');

  const sc = schema();
  const titles = { dashboard:'Dashboard', inbound:'Data Inbound', outbound:'Data Outbound', stock: sc.stockPageTitle || 'Stock On Hand Teknisi' };
  document.getElementById('pageTitle').textContent = titles[page] || page;

  /* Lazy render on page switch */
  if(page==='inbound')   { populateInboundFilters(); renderInboundTable(); }
  if(page==='outbound')  renderOutboundTable();
  if(page==='stock')     { renderStockSummary(); renderStockDetail(); }
  if(page==='dashboard') renderDashboard();

  /* Mobile: close sidebar */
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('sidebarOverlay').classList.remove('show');
}

/* ===================== RENDER ALL ===================== */
function renderAll(){
  renderDashboard();
  if(currentPage==='inbound')  { populateInboundFilters(); renderInboundTable(); }
  if(currentPage==='outbound') renderOutboundTable();
  if(currentPage==='stock')    { renderStockSummary(); renderStockDetail(); }
  updateSourceStatus();
}

/* ===================== INIT ===================== */
function init(){
  /* Init customer data buckets */
  CUSTOMER_LIST.forEach(id => {
    if(!state.customers[id]) state.customers[id] = { inboundRaw:[], outboundRaw:[], fileName:'', updatedAt:null };
  });

  restoreAllCustomers();

  /* Build customer dropdown */
  const dropdownMenu = document.getElementById('customerDropdownMenu');
  if(dropdownMenu){
    dropdownMenu.innerHTML = CUSTOMER_LIST.map(id => {
      const sc = CUSTOMERS[id];
      return `<button class="customer-option" data-cid="${id}" style="--ccolor:${sc.color}">
        <span class="co-dot" style="background:${sc.color}"></span>${sc.name}
      </button>`;
    }).join('');
    dropdownMenu.querySelectorAll('.customer-option').forEach(btn =>
      btn.addEventListener('click', () => {
        switchCustomer(btn.dataset.cid);
        document.getElementById('customerDropdown').classList.remove('open');
      })
    );
  }

  /* Customer dropdown toggle */
  const dropdownToggle = document.getElementById('customerDropdownToggle');
  if(dropdownToggle){
    dropdownToggle.addEventListener('click', e => {
      e.stopPropagation();
      document.getElementById('customerDropdown').classList.toggle('open');
    });
    document.addEventListener('click', () => document.getElementById('customerDropdown').classList.remove('open'));
  }

  /* Logo upload */
  document.getElementById('logoFile').addEventListener('change', e => {
    const file = e.target.files[0]; if(!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      state.logoDataUrl = ev.target.result;
      document.getElementById('logoImg').src = state.logoDataUrl;
      document.getElementById('logoImg').style.display = 'block';
      document.getElementById('logoFallback').style.display = 'none';
      saveLogoToStorage();
    };
    reader.readAsDataURL(file);
  });
  if(state.logoDataUrl){
    document.getElementById('logoImg').src = state.logoDataUrl;
    document.getElementById('logoImg').style.display = 'block';
    document.getElementById('logoFallback').style.display = 'none';
  }

  /* Excel upload */
  document.getElementById('excelFile').addEventListener('change', e => {
    const file = e.target.files[0]; if(!file) return;
    handleExcelFile(file);
    e.target.value = '';
  });

  /* Reset button */
  document.getElementById('resetDataBtn').addEventListener('click', () => {
    if(!confirm(`Reset semua data ${state.activeCustomer}?`)) return;
    resetCustomerData(state.activeCustomer);
    stockState.expandedPIC=null; stockState.search='';
    ['inboundSearch','outboundSearch','stockSearch'].forEach(id => {
      const el=document.getElementById(id); if(el) el.value='';
    });
    const returnEl=document.getElementById('outboundFilterReturn');
    if(returnEl) returnEl.value='';
    renderAll();
    updateSourceStatus();
  });

  /* Nav */
  document.querySelectorAll('.nav-item[data-page]').forEach(btn =>
    btn.addEventListener('click', () => showPage(btn.dataset.page))
  );

  /* Sidebar mobile */
  document.getElementById('sidebarToggle').addEventListener('click', () => {
    document.getElementById('sidebar').classList.toggle('open');
    document.getElementById('sidebarOverlay').classList.toggle('show');
  });
  document.getElementById('sidebarOverlay').addEventListener('click', () => {
    document.getElementById('sidebar').classList.remove('open');
    document.getElementById('sidebarOverlay').classList.remove('show');
  });

  /* Inbound controls */
  document.getElementById('inboundSearch').addEventListener('input', debounce(e => {
    inboundState.search=e.target.value.trim(); inboundState.page=1; renderInboundTable();
  },200));
  document.getElementById('inboundFilterRegion').addEventListener('change', e => {
    inboundState.region=e.target.value; inboundState.page=1; renderInboundTable();
  });
  document.getElementById('inboundFilterStatus').addEventListener('change', e => {
    inboundState.status=e.target.value; inboundState.page=1; renderInboundTable();
  });
  document.getElementById('inboundPageSize').addEventListener('change', e => {
    inboundState.pageSize=parseInt(e.target.value); inboundState.page=1; renderInboundTable();
  });
  document.getElementById('inboundExportBtn').addEventListener('click', () =>
    exportToExcel(inboundState.lastFiltered, schema().inboundColumns, `Inbound_${state.activeCustomer}.xlsx`)
  );
  setupSortableTable('inboundTable', inboundState, renderInboundTable);

  /* Outbound controls */
  document.getElementById('outboundSearch').addEventListener('input', debounce(e => {
    outboundState.search=e.target.value.trim(); outboundState.page=1; renderOutboundTable();
  },200));
  document.getElementById('outboundFilterReturn').addEventListener('change', e => {
    outboundState.returnFilter=e.target.value; outboundState.page=1; renderOutboundTable();
  });
  document.getElementById('outboundPageSize').addEventListener('change', e => {
    outboundState.pageSize=parseInt(e.target.value); outboundState.page=1; renderOutboundTable();
  });
  document.getElementById('outboundExportBtn').addEventListener('click', () =>
    exportToExcel(outboundState.lastFiltered, schema().outboundColumns, `Outbound_${state.activeCustomer}.xlsx`)
  );
  setupSortableTable('outboundTable', outboundState, renderOutboundTable);

  /* Stock controls */
  document.getElementById('stockSearch').addEventListener('input', debounce(e => {
    stockState.search=e.target.value.trim(); renderStockSummary();
  },200));
  document.getElementById('stockDetailClose').addEventListener('click', () => {
    stockState.expandedPIC=null; renderStockSummary(); renderStockDetail();
  });

  updateCustomerUI();
  populateInboundFilters();
  showPage('dashboard');

  /* ---- Google Sheets auto-refresh ---- */
  const hasAnySheets = CUSTOMER_LIST.some(id => {
    const sc = CUSTOMERS[id];
    return sc?.sheets?.inbound || sc?.sheets?.outbound;
  });
  if(hasAnySheets) startAutoRefresh();

  /* Manual refresh button */
  const manualBtn = document.getElementById('manualRefreshBtn');
  if(manualBtn){
    manualBtn.addEventListener('click', async () => {
      await refreshAllSheets(false);
      updateCountdown();
    });
  }
}

document.addEventListener('DOMContentLoaded', init);
