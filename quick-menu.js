// ==UserScript==
// @name         HRIS KTI — Quick Menu
// @namespace    https://hris.kti.co.id/
// @version      1.3.0
// @description  Floating menu for background utility tasks on KTI HRIS (Full Automation)
// @match        https://hris.kti.co.id/*
// @icon         https://hris.kti.co.id/favicon.ico
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        GM_addStyle
// @grant        GM_getValue
// @grant        GM_setValue
// @run-at       document-idle
// @noframes
// ==/UserScript==

(function () {
  'use strict';


  /* ============================================================
   * 0. CONSTANTS
   * ============================================================ */

  const APP_CONFIG = Object.freeze({
    BATCH_POOL_SIZE: 10,
    BATCH_MAX_LIMIT: 100,
  });

  const SHIFT_RULES = {
    shift1: { jamTebakMulai: 5.0, jamTebakAkhir: 9.99, batasMasukLembur: 5.0, batasKeluarLembur: 14.5, batasTerlambatMasuk: 6.25, batasPulangAwal: 14.0 },
    shift2: { jamTebakMulai: 9.99, jamTebakAkhir: 16.0, batasMasukLembur: 13.0, batasKeluarLembur: 22.5, batasTerlambatMasuk: 14.25, batasPulangAwal: 22.0 },
    shift3: { batasMasukLemburAwal: 12.0, batasMasukLemburAkhir: 20.0, batasKeluarLembur: 30.5, batasAwalMasuk: 15.0, batasTerlambatMasuk: 22.25, batasPulangAwal: 30.0 }
  };

  const SPKL_RULES = [
    // Definisi Aturan Perbaikan SPKL (Shift, OT Saat Ini -> OT Target)
    { shift: '3', currentOtType: '1', targetOtType: '4', desc: 'Shift III + Biasa OT (1) -> OT Awal (4)' }
  ];

  const HALFDAY_RULES = {
    shift1: { jamMulai: 5.0, jamAkhir: 11.0, batasPulangAwal: 11.0 },
    shift2: { jamMulai: 10.0, jamAkhir: 16.0, altJamMulai: 14.0, altJamAkhir: 19.0, batasPulangAwal: 16.0, altBatasPulangAwal: 19.0 },
    shift3: { jamMulai: 22.0, jamAkhir: 28.0, batasPulangAwal: 28.0 } // 04:00 next day = 28.0
  };

  const COL = Object.freeze({
    TGL: 0,       // Tanggal
    MSK: 1,       // Jam Masuk
    KLR: 2,       // Jam Keluar
    KET: 3,       // Keterangan
    HARI_KERJA: 4, // Hari Kerja
    SHIFT1: 5,     // Shift I
    SHIFT2: 6,     // Shift II
    SHIFT3: 7,     // Shift III
    OTB: 8,        // Overtime Begin
    OTL: 9,        // Overtime Leave
    OTP: 10,       // Overtime Present
  });

  const COL_LABELS = Object.freeze({
    [COL.TGL]: 'Tanggal',
    [COL.MSK]: 'MSK',
    [COL.KLR]: 'KLR',
    [COL.KET]: 'Ket.',
    [COL.HARI_KERJA]: 'Hari Kerja',
    [COL.SHIFT1]: 'Shift I',
    [COL.SHIFT2]: 'Shift II',
    [COL.SHIFT3]: 'Shift III',
    [COL.OTB]: 'OTB',
    [COL.OTL]: 'OTL',
    [COL.OTP]: 'OTP',
  });

  /** Named thresholds — replaces magic numbers throughout the script. */
  const THRESHOLDS = Object.freeze({
    MIN_LIBUR_5_HARI_KERJA: 8,    // >= N libur/bulan → 5-day work week
    ADJ_KLR_SHIFT2_BATAS: 12.0,  // KLR <= 12 on shift2 → add 24 (next day)
    ADJ_MSK_SHIFT3_BATAS: 12.0,  // MSK <= 12 on shift3 → add 24 (next day)
    ADJ_KLR_SHIFT3_BATAS: 15.0,  // KLR <= 15 on shift3 → add 24 (next day)
    SHIFT1_MSK_UPPER_BATAS: 12.0, // MSK > 12 on shift1 → Cek Distribusi
    SHIFT2_MSK_LOWER_BATAS: 10.0, // MSK < 10 on shift2 → Cek Distribusi
    SHIFT2_MSK_UPPER_BATAS: 20.0, // MSK > 20 on shift2 → Cek Distribusi
    SHIFT3_MSK_UPPER_BATAS: 26.0, // MSK > 26 on shift3 → Cek Distribusi
    OT_BATAS_WAJAR: 24,           // OT value > N = unreasonable
    SHIFT1_KLR_LEMBUR_5HR: 17.0, // 5-day work week shift1 OT exit threshold
    SHIFT1_TERLAMBAT_5HR: 7.75,  // 5-day work week shift1 late threshold (07:45)
    SHIFT1_PULANG_AWAL_5HR: 16.5,// 5-day work week shift1 early leave threshold (16:30)
    SHIFT2_KLR_LEMBUR_5HR: 23.5, // 5-day work week shift2 OT exit threshold
  });

  const LIBUR_NASIONAL_2026 = Object.freeze({
    1: [1],
    2: [18, 19],
    3: [19, 21, 22],
    4: [3],
    5: [1, 14, 24, 27],
    6: [1],
    7: [17],
    8: [17],
    9: [27],
    10: [],
    11: [],
    12: [25]
  });

  /** Timing delays (ms) — replaces magic numbers throughout the script. */
  const TIMING = Object.freeze({
    RESULT_AUTO_HIDE: 3500,
    DEBOUNCE_RENDER: 150,
    RETRY_DELAY_BASE: 500,
    MAX_RETRY_COUNT: 3,
    DOM_POLL_TIMEOUT: 5000,
    SPKL_HIGHLIGHT_DELAY: 500,
    AUTO_FILL_DELAY: 500,
    BARCODE_AUTO_CLICK_DELAY: 500,
    BARCODE_AUTOFILL_DELAY: 600,
    DISTRIBUSI_START_DELAY: 500,
    DISTRIBUSI_SUBMIT_DELAY: 800,
    SPKL_REDIRECT_DELAY: 1000,
    SPKL_INPUT_DELAY: 450,
    SPKL_CLICK_DELAY: 850,
    SPKL_SUBMIT_DELAY: 500,
    SPKL_BATCH_NRP_RESUME_DELAY: 1500,
    SPKL_BATCH_NRP_CLICK_DELAY: 1200,
    SESSION_SAVE_DELAY: 200,
    SESSION_SAVE_DELAY2: 250,
    SESSION_SAVE_DELAY3: 350,
    MODAL_AUTO_CLICK_DELAY: 500,
    PANEL_TRANSITION_DELAY: 300,
    BACKGROUND_HEARTBEAT_INTERVAL: 5000,
    FADE_SHORT: 10,
    FADE_MEDIUM: 250,
  });

  /** sessionStorage schema version for cache invalidation. */
  const STORAGE_SCHEMA_VERSION = 1;

  /** sessionStorage key constants. */
  const STORAGE = Object.freeze({
    SCHEMA_VERSION: 'qm_schema_version',
    EMP_JK: 'qm_jk_',
    EMP_NAMA: 'qm_nama_',
    EMP_BAG: 'qm_bag_',
    EMP_SEK: 'qm_sek_',
    EMP_GRP: 'qm_grp_',
    AUTO_NRP: 'qm_auto_nrp',
    AUTO_BULAN: 'qm_auto_bulan',
    AUTO_NRP_FILL: 'qm_auto_nrp_fill',
    AUTO_DATE_FILL: 'qm_auto_date_fill',
    AUTO_ADD_DATA: 'qm_auto_add_data',
    AUTO_BARCODE_SEARCH: 'qm_auto_barcode_search',
    HIGHLIGHT_SPKL: 'qm_highlight_spkl_date',
    SPKL_SAVED: 'spkl_saved_data',
    SPKL_BATCH: 'hris_spkl_ot_runner_v1',
    INPUT_HADIR: 'qm_auto_input_hadir_data',
    HADIR_BATCH: 'qm_auto_hadir_batch_v1',
    HISTORY: 'qm_history',
    AUTO_FLOW: 'qm_auto_flow_v1',
    RETURN_URL: 'qm_return_url',
    AUTO_FINISHED: 'qm_auto_finished',
    SPKL_FIX_PENDING: 'qm_auto_spkl_redirect_pending',
    SPKL_QUEUE: 'qm_spkl_fix_queue',
    SPKL_CURRENT_INDEX: 'qm_spkl_fix_index',
    EMP_KK: 'qm_KK_'
  });

  const SELECTORS = Object.freeze({
    SPKL_TABLE_ROWS: 'table tbody tr',
    SPKL_EDIT_BTN: 'button, a, .btn', // Used inside row.querySelectorAll
    SPKL_MODAL_MSK: 'input[name*="jam_masuk"], input[name*="msk"], #jam_masuk_edit, #msk_edit',
    SPKL_MODAL_KLR: 'input[name*="jam_keluar"], input[name*="klr"], #jam_keluar_edit, #klr_edit',
    SPKL_MODAL_OT_TYPE: 'select[name*="jenis_ot"], select[name*="jenis_lembur"], select[name*="ot_type"], select[name*="kode_lembur"], #jenis_ot_edit, #type_ot_edit, #ot_type_edit',
    SPKL_MODAL_JAM_OT: 'input[name*="jam_lembur"], input[name*="jam_ot"], #jam_ot_edit',
    SPKL_MODAL_SUBMIT: 'button[type="submit"], .btn-primary, #btn-save'
  });

  /** URL route constants. */
  const ROUTES = Object.freeze({
    BASE: 'https://hris.kti.co.id',
    KARYAWAN_SEARCH: (nrp) => `https://hris.kti.co.id/karyawan?kode_bagian=&kode_seksi=&kode_group=&status_karyawan=A&s=${nrp}`,
    KARYAWANOS_SEARCH: (nrp) => `https://hris.kti.co.id/karyawanoutsource?kode_bagian=&kode_seksi=&kode_group=&status_karyawan=A&s=${nrp}`,
    KARYAWAN_GENERAL: (id) => `https://hris.kti.co.id/karyawan/general/${id}`,
    KARYAWAN_PROFILE: (id) => `https://hris.kti.co.id/karyawan/profile/${id}`,
    KARYAWAN_EDIT: (id) => `https://hris.kti.co.id/karyawan/editgeneral/${id}`,
    KARYAWANOS_GENERAL: (id) => `https://hris.kti.co.id/karyawanoutsource/general/${id}`,
    KARYAWANOS_PROFILE: (id) => `https://hris.kti.co.id/karyawanoutsource/profile/${id}`,
    KARYAWANOS_EDIT: (id) => `https://hris.kti.co.id/karyawanoutsource/editgeneral/${id}`,
    TABEL_HADIR: (bulan, tahun, nrp) => `https://hris.kti.co.id/tabelkehadiran?bulan=${bulan}&tahun=${tahun}&nrp=${nrp}`,
    TABEL_HADIR_OS: (bulan, tahun, nrp) => `https://hris.kti.co.id/tabelkehadirankkwt?bulan=${bulan}&tahun=${tahun}&nrp=${nrp}`,
    DISTRIBUSI: (nrp) => `https://hris.kti.co.id/distribusijamkerja`,
    DISTRIBUSI_OS: (nrp) => `https://hris.kti.co.id/distribusijamkerjaos`,
    DISTRIBUSI_KK: 'https://hris.kti.co.id/distribusikalenderkerja',
    ABSEN_BARCODE: (tahun, bulan, nrp) => `https://hris.kti.co.id/absenbarcode?tahun=${tahun}&bulan=${bulan}&kode_bagian=&kode_seksi=&kode_group=&nrp=${nrp}`,
    ABSEN_BARCODE_OS: (tahun, bulan, nrp) => `https://hris.kti.co.id/absenbarcodeos?tahun=${tahun}&bulan=${bulan}&kode_bagian=&kode_seksi=&kode_group=&nrp=${nrp}`,
    ABSEN_BARCODE_PAGE: 'https://hris.kti.co.id/absenbarcode',
    ABSEN_BARCODE_OS_PAGE: 'https://hris.kti.co.id/absenbarcodeos',
    ABSEN_BARCODE_CREATE: 'https://hris.kti.co.id/absenbarcode/create',
    ABSEN_BARCODE_OS_CREATE: 'https://hris.kti.co.id/absenbarcodeos/create',
    ABSEN_BARCODE_ADD: 'https://hris.kti.co.id/absenbarcode/add',
    ABSEN_BARCODE_OS_ADD: 'https://hris.kti.co.id/absenbarcodeos/add',
    SPKL_BASE: 'https://hris.kti.co.id/spkl',
    SPKL_OS_BASE: 'https://hris.kti.co.id/spkloutsource',
    SPKL_ONLINE: (tahun, bulan, minDate, maxDate, nrp) => `https://hris.kti.co.id/spklonline?tanggal_awal=${tahun}-${bulan}-${minDate}&tanggal_akhir=${tahun}-${bulan}-${maxDate}&kode_bagian=&kode_seksi=&kode_group=&approval_status=&status_shift=&s=${nrp}`,
    SPKL_ONLINE_SINGLE: (fullDate, nrp) => `https://hris.kti.co.id/spklonline?tanggal_awal=${fullDate}&tanggal_akhir=${fullDate}&kode_bagian=&kode_seksi=&kode_group=&approval_status=&status_shift=&s=${nrp}`,
    SPKL_CREATE: 'https://hris.kti.co.id/spkl/create',
    SPKL_OS_CREATE: 'https://hris.kti.co.id/spkloutsource/create',
    SPKL_ADD: 'https://hris.kti.co.id/spkl/add',
    SPKL_OS_ADD: 'https://hris.kti.co.id/spkloutsource/add',
  });

  const PROFILE_CONFIG = Object.freeze({
    MEDIAN_SAMPLE_SIZE: 3,
    HOT_SYNC_MS: 50,
    HOT_BATCH_RENDER_TOTAL_MS: 500,
    SAMPLE_HISTORY_LIMIT: 15,
  });

  /** Common success indicators in HRIS response text. */
  const SUCCESS_KEYWORDS = ['Berhasil', 'Selesai', 'sukses', 'successfully', 'Distribution Process Completed', 'alert-success'];

  /* ============================================================
   * 1. CORE UTILITIES
   * ============================================================ */

  /** Parse "HH:MM" or "HH.MM" into decimal hours. */
  function parseTimeToDecimal(t) {
    if (!t) return null;
    const parts = t.replace(':', '.').split('.');
    if (parts.length >= 2) {
      const h = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10);
      if (isNaN(h) || isNaN(m)) return null;
      return h + m / 60;
    }
    return null;
  }


  /** Simple HTML escape helper. */
  function escapeHtml(str) {
    if (!str) return '';
    return String(str).replace(/[&<>"']/g, m => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[m]));
  }


  /**
   * Enhanced logger that respects state.debug.
   * @param {string} msg
   * @param {any} data
   */
  // log(msg, data) — removed, replaced by Logger

  /**
   * Robust date parser for HRIS formats:
   * - DD/MM/YYYY
   * - DD-MM-YYYY
   * - DD-MMM-YYYY (e.g. 21-Apr-2026)
   * @param {string} str
   * @returns {Date|null}
   */
  function parseHrisDate(str) {
    if (!str) return null;
    const clean = str.trim();

    // DD/MM/YYYY or DD-MM-YYYY
    const dmyMatch = clean.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
    if (dmyMatch) {
      return new Date(dmyMatch[3], dmyMatch[2] - 1, dmyMatch[1]);
    }

    // DD-MMM-YYYY
    const dMmmYMatch = clean.match(/^(\d{1,2})-(\w{3})-(\d{4})$/);
    if (dMmmYMatch) {
      const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const monthNamesId = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
      const mStr = dMmmYMatch[2].charAt(0).toUpperCase() + dMmmYMatch[2].slice(1).toLowerCase();

      let mIdx = monthNames.indexOf(mStr);
      if (mIdx === -1) mIdx = monthNamesId.indexOf(mStr);

      if (mIdx !== -1) {
        return new Date(dMmmYMatch[3], mIdx, dMmmYMatch[1]);
      }
    }

    // YYYY-MM-DD
    const ymdMatch = clean.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
    if (ymdMatch) {
      return new Date(ymdMatch[1], ymdMatch[2] - 1, ymdMatch[3]);
    }

    // Fallback: browser default
    const d = new Date(clean);
    // Adjust ISO parsing to local time to prevent mismatch
    if (clean.match(/^\d{4}-\d{2}-\d{2}$/)) {
      return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
    }
    return isNaN(d.getTime()) ? null : d;
  }

  function isCurrentPageOutsource() {
    const path = getCurrentPath();
    return path.includes('kkwt') || path.includes('outsource') || path.includes('os');
  }

  function isOutsourceNrp(nrp) {
    if (!nrp) return false;
    const cached = sessionStorage.getItem('qm_is_os_' + nrp);
    if (cached !== null) return cached === 'true';

    // Fallback: check if active page context matches
    const ctx = getPageContext();
    if (ctx.nrp === nrp) {
      const isOS = isCurrentPageOutsource();
      sessionStorage.setItem('qm_is_os_' + nrp, isOS ? 'true' : 'false');
      return isOS;
    }

    // Default fallback
    return String(nrp || '').length === 8;
  }

  function getCurrentQueryParams() {
    return new URLSearchParams(window.location.search);
  }

  function getCurrentPath() {
    return window.location.pathname || '';
  }

  function isAttendancePagePath(path = getCurrentPath()) {
    return (path.includes('/tabelkehadiran') || path.includes('/tabelkehadirankkwt')) && !path.includes('/rekap');
  }

  function isBarcodePagePath(path = getCurrentPath()) {
    return path.includes('/absenbarcode');
  }

  function isBarcodeListPagePath(path = getCurrentPath()) {
    return isBarcodePagePath(path) && !isBarcodeCreatePagePath(path) && !isBarcodeAddPagePath(path);
  }

  function isBarcodeCreatePagePath(path = getCurrentPath()) {
    return path.includes('/absenbarcode/create') || path.includes('/absenbarcodeos/create');
  }

  function isBarcodeAddPagePath(path = getCurrentPath()) {
    return path.includes('/absenbarcode/add') || path.includes('/absenbarcodeos/add');
  }

  function isSpklPagePath(path = getCurrentPath()) {
    return path.includes('/spkl') && !path.includes('/spklonline');
  }

  function isDistribusiKalenderPagePath(path = getCurrentPath()) {
    return path.includes('/distribusikalenderkerja');
  }

  function spklAddPageKind(path = getCurrentPath()) {
    if (path === '/spkl/add') return 'internal';
    if (path === '/spkloutsource/add') return 'outsource';
    return null;
  }


  function absenCreatePageKind(path = getCurrentPath()) {
    if (path === '/absenbarcode/create') return 'internal';
    if (path === '/absenbarcodeos/create') return 'outsource';
    return null;
  }


  function isHrisSuccess(text) {
    return SUCCESS_KEYWORDS.some(kw => text.includes(kw));
  }

  function createAutomationFlowId(type) {
    return `${type}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  }

  function getAutomationFlow() {
    const raw = sessionStorage.getItem(STORAGE.AUTO_FLOW);
    if (!raw) return null;
    try {
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== 'object' || !parsed.id || !parsed.type) {
        sessionStorage.removeItem(STORAGE.AUTO_FLOW);
        return null;
      }
      return parsed;
    } catch (e) {
      sessionStorage.removeItem(STORAGE.AUTO_FLOW);
      Logger.warn('AUTO_FLOW payload tidak valid.', e);
      return null;
    }
  }

  function createAutomationFlow(type, returnUrl, meta = {}) {
    const flow = {
      id: createAutomationFlowId(type),
      type,
      returnUrl: returnUrl || window.location.href,
      sourceUrl: window.location.href,
      createdAt: Date.now(),
      finished: false,
      meta
    };
    sessionStorage.setItem(STORAGE.AUTO_FLOW, JSON.stringify(flow));
    sessionStorage.setItem(STORAGE.RETURN_URL, flow.returnUrl);
    sessionStorage.removeItem(STORAGE.AUTO_FINISHED);
    return flow;
  }

  function markAutomationFlowFinished(flowId) {
    const flow = getAutomationFlow();
    if (!flow || flow.id !== flowId) return false;
    flow.finished = true;
    flow.finishedAt = Date.now();
    sessionStorage.setItem(STORAGE.AUTO_FLOW, JSON.stringify(flow));
    sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
    return true;
  }

  function clearAutomationFlow(flowId) {
    const flow = getAutomationFlow();
    if (flow && flowId && flow.id !== flowId) return false;
    sessionStorage.removeItem(STORAGE.AUTO_FLOW);
    sessionStorage.removeItem(STORAGE.AUTO_FINISHED);
    if (!flow || !flow.returnUrl || sessionStorage.getItem(STORAGE.RETURN_URL) === flow.returnUrl) {
      sessionStorage.removeItem(STORAGE.RETURN_URL);
    }
    return !!flow;
  }

  function finishAutomationFlow(flowId) {
    const flow = getAutomationFlow();
    if (!flow || flow.id !== flowId) return;

    markAutomationFlowFinished(flowId);
    const returnUrl = flow.returnUrl;
    clearAutomationFlow(flowId);

    if (returnUrl && returnUrl !== window.location.href) {
      window.location.href = returnUrl;
    } else {
      window.location.reload();
    }
  }

  function isActiveAutomationFlow(flowId, type) {
    const flow = getAutomationFlow();
    if (!flow) return false;
    if (flowId && flow.id !== flowId) return false;
    if (type && flow.type !== type) return false;
    return true;
  }

  /** Parse CSS class/style/bgcolor attributes to detect Libur/HalfDay status.
   *  Used by both countHolidays() and getAttendanceRowStatus() to avoid duplicated regex. */
  function getRowFlags(tr) {
    const trStr = ((tr.getAttribute('class') || '') + (tr.getAttribute('style') || '') + (tr.getAttribute('bgcolor') || '')).toLowerCase();
    const tds = tr.querySelectorAll('td');
    const td0Str = tds.length > 0 ? ((tds[0].getAttribute('class') || '') + (tds[0].getAttribute('style') || '') + (tds[0].getAttribute('bgcolor') || '')).toLowerCase() : '';
    const combined = trStr + td0Str;
    return {
      isLiburColor: /danger|red|#ff0000/.test(combined),
      isHalfDayColor: /warning|yellow|blue|#ffff00|#0000ff/.test(combined)
    };
  }


  function median(values) {
    if (!values.length) return 0;
    const sorted = [...values].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  }

  function perfNow() {
    return (typeof performance !== 'undefined' && typeof performance.now === 'function')
      ? performance.now()
      : Date.now();
  }

  /* ============================================================
   * 2. DOM HELPERS
   * ============================================================ */

  /** Unified DOM-based sanitization and parsing. */
  function parseHTML(html, fullDoc = false) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const dangerous = doc.querySelectorAll('script,style,link,iframe,object,embed,math');
    dangerous.forEach(el => el.remove());
    doc.querySelectorAll('*').forEach(el => {
      Array.from(el.attributes).forEach(attr => {
        if (attr.name.startsWith('on') || attr.value.toLowerCase().includes('javascript:')) {
          el.removeAttribute(attr.name);
        }
      });
    });
    return doc;
  }


  /** Safely set innerHTML using DOM-based sanitization via parseHTML. */
  function renderSafe(el, html) {
    if (!el) return;
    if (!html && html !== '') { el.innerHTML = html; return; }
    const doc = parseHTML(`<div>${html}</div>`);
    const container = doc.querySelector('div');
    el.innerHTML = '';
    while (container.firstChild) {
      el.appendChild(container.firstChild);
    }
  }

  /** Set a form field value and dispatch the given events.
   *  Replaces the repeated pattern: el.value = val; el.dispatchEvent(new Event(...)) */
  function setField(el, value, events = ['change']) {
    if (!el) return;
    el.value = value;
    events.forEach(e => el.dispatchEvent(new Event(e, { bubbles: true })));
    // Auto-refresh SelectPicker if applicable
    if (el.tagName === 'SELECT' && el.classList.contains('selectpicker')) {
      refreshPicker(el);
    }
  }

  /** Find an option in a select element matching a predicate, select it, and dispatch change.
   *  Returns true if a match was found. */
  function pickOption(select, matchFn) {
    if (!select) return false;
    const options = select.querySelectorAll('option');
    for (const opt of options) {
      if (matchFn(opt)) {
        select.value = opt.value;
        select.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
      }
    }
    return false;
  }


  /** Refresh a bootstrap-selectpicker if jQuery is available. */
  function refreshPicker(select) {
    if (typeof window.$ !== 'undefined' && window.$(select).selectpicker) {
      window.$(select).selectpicker('refresh');
    }
  }


  /** Custom Event Delegation helper. */
  function on(eventName, selector, handler) {
    document.addEventListener(eventName, function (e) {
      let target = e.target;
      while (target && target !== document) {
        if (target.matches(selector)) {
          handler.call(target, e);
          return;
        }
        target = target.parentNode;
      }
    });
  }


  /** Wait for element using MutationObserver (replaces setInterval polling). */
  function waitFor(selector, timeout = 5000) {
    return new Promise((resolve, reject) => {
      const el = document.querySelector(selector);
      if (el) return resolve(el);

      const observer = new MutationObserver(() => {
        const target = document.querySelector(selector);
        if (target) {
          observer.disconnect();
          resolve(target);
        }
      });

      observer.observe(document.body, { childList: true, subtree: true });
      setTimeout(() => {
        observer.disconnect();
        reject(new Error(`Timeout waiting for ${selector}`));
      }, timeout);
    });
  }

  /** Extract common page params used by anomaly + SPKL + barcode checks. */
  function getPageContext() {
    const params = getCurrentQueryParams();
    let nrp = params.get('nrp');
    if (!nrp) {
      const input = document.querySelector('input[name="nrp"]');
      if (input) nrp = input.value;
    }
    if (!nrp) {
      const match = document.body.textContent.match(/NRP\s*:\s*(\d{4,8})/i);
      if (match) nrp = match[1];
    }

    if (nrp) {
      const isOS = isCurrentPageOutsource();
      sessionStorage.setItem('qm_is_os_' + nrp, isOS ? 'true' : 'false');
    }

    const selTahun = document.querySelector('select[name="tahun"]');
    const selBulan = document.querySelector('#bulan');
    const selBagian = document.querySelector('select[name="kode_bagian"]');
    const selSeksi = document.querySelector('select[name="kode_seksi"]');

    return {
      tahun: params.get('tahun') || (selTahun ? selTahun.value : null) || new Date().getFullYear(),
      bulan: params.get('bulan') || (selBulan ? selBulan.value : null) || (new Date().getMonth() + 1),
      nrp,
      bagian: (selBagian ? selBagian.value : null) || params.get('kode_bagian') || '',
      seksi: (selSeksi ? selSeksi.value : null) || params.get('kode_seksi') || '',
    };
  }


  /* ============================================================
   * 3. SESSION STORAGE
   * ============================================================ */


  /** Read cached employee data from sessionStorage. Returns null if not fully cached. */
  function isValidEmployeeId(id) {
    if (!id) return false;
    const cleanId = String(id).toLowerCase().trim();
    const invalidKeywords = ['rekap', 'laporan', 'pembayaran', 'detail', 'pembayarandetail'];
    for (const kw of invalidKeywords) {
      if (cleanId.includes(kw)) return false;
    }
    return /^[a-z0-9]+$/i.test(cleanId);
  }

  function readEmployeeCache(nrp) {
    const jk = sessionStorage.getItem('qm_jk_' + nrp);
    const nama = sessionStorage.getItem('qm_nama_' + nrp);
    const id = sessionStorage.getItem('qm_id_' + nrp);
    const kk = sessionStorage.getItem('qm_KK_' + nrp);
    const bag = sessionStorage.getItem('qm_bag_' + nrp);
    if (!jk || !nama || !id || !kk || !bag) return null;

    if (!isValidEmployeeId(id)) {
      clearEmployeeCache(nrp);
      return null;
    }

    // Auto-invalidate stale edit URLs that were mistakenly cached as "/edit/"
    sessionStorage.removeItem('qm_edit_url_' + nrp);
    const editUrl = employeeEditUrl(nrp, id);
    if (false) {
      sessionStorage.removeItem('qm_edit_url_' + nrp);
      return null; // Force fresh fetch
    }

    return {
      jk, KK: kk, nama, id,
      editUrl,
      bagian: bag || '',
      seksi: sessionStorage.getItem('qm_sek_' + nrp) || '',
      group: sessionStorage.getItem('qm_grp_' + nrp) || ''
    };
  }

  function writeEmployeeCache(nrp, emp) {
    if (emp.id) sessionStorage.setItem('qm_id_' + nrp, emp.id);
    if (emp.editUrl) sessionStorage.setItem('qm_edit_url_' + nrp, emp.editUrl);
    if (emp.jk) sessionStorage.setItem('qm_jk_' + nrp, emp.jk);
    if (emp.KK) sessionStorage.setItem('qm_KK_' + nrp, emp.KK);
    if (emp.nama) sessionStorage.setItem('qm_nama_' + nrp, emp.nama);
    if (emp.bagian) sessionStorage.setItem('qm_bag_' + nrp, emp.bagian);
    if (emp.seksi) sessionStorage.setItem('qm_sek_' + nrp, emp.seksi);
    if (emp.group) sessionStorage.setItem('qm_grp_' + nrp, emp.group);
  }

  function createEmptyKaryawanEditor() {
    return {
      key: '',
      nrp: '',
      loading: false,
      saving: false,
      error: '',
      notice: '',
      emp: null,
      jkOptions: [],
      kkOptions: []
    };
  }

  function createEmptyKaryawanDetail() {
    return {
      key: '',
      nrp: '',
      loading: false,
      error: '',
      profile: null
    };
  }

  function createEmptyAttendanceCheck() {
    return {
      loading: false,
      error: '',
      requestKey: '',
      summary: null
    };
  }

  function createEmptySpklCheck() {
    return {
      loading: false,
      error: '',
      requestKey: '',
      summary: null
    };
  }

  function resetKaryawanEditor() {
    state.karyawanEditor = createEmptyKaryawanEditor();
  }

  function resetKaryawanDetail() {
    state.karyawanDetail = createEmptyKaryawanDetail();
  }

  function resetKaryawanPanels() {
    state.karyawanActivePanel = { key: '', mode: '' };
    resetKaryawanEditor();
    resetKaryawanDetail();
  }

  function resetSpklCheck() {
    state.spklCheck = createEmptySpklCheck();
  }

  function clearEmployeeCache(nrp) {
    if (!nrp) return;
    [
      'qm_id_' + nrp,
      'qm_edit_url_' + nrp,
      'qm_jk_' + nrp,
      'qm_KK_' + nrp,
      'qm_nama_' + nrp,
      'qm_bag_' + nrp,
      'qm_sek_' + nrp,
      'qm_grp_' + nrp,
      'qm_jk_options_' + nrp,
      'qm_KK_options_' + nrp
    ].forEach(key => sessionStorage.removeItem(key));

    if (state._lastEditNrp === nrp) {
      state._lastEditNrp = null;
      cachedEditHtml = null;
    }
  }

  /**
   * Validates sessionStorage schema version. Clears all qm_* keys if
   * the version is missing or does not match the current STORAGE_SCHEMA_VERSION.
   * @returns {boolean} true if schema was already valid, false if it was cleared.
   */
  function validateStorageSchema() {
    const stored = sessionStorage.getItem(STORAGE.SCHEMA_VERSION);
    if (stored === String(STORAGE_SCHEMA_VERSION)) return true;

    const keysToRemove = [];
    for (let i = 0; i < sessionStorage.length; i++) {
      const key = sessionStorage.key(i);
      if (key && key.startsWith('qm_')) keysToRemove.push(key);
    }
    keysToRemove.forEach(k => sessionStorage.removeItem(k));
    sessionStorage.setItem(STORAGE.SCHEMA_VERSION, String(STORAGE_SCHEMA_VERSION));
    Logger.info(`Schema versi lama (${stored || 'kosong'}) dihapus. Versi baru: ${STORAGE_SCHEMA_VERSION}`);
    return false;
  }

  /* ============================================================
   * 4. ROUTE BUILDERS
   * ============================================================ */

  function employeeRoutesBySource(isOS) {
    return isOS
      ? {
        search: ROUTES.KARYAWANOS_SEARCH,
        general: ROUTES.KARYAWANOS_GENERAL,
        profile: ROUTES.KARYAWANOS_PROFILE,
        edit: ROUTES.KARYAWANOS_EDIT
      }
      : {
        search: ROUTES.KARYAWAN_SEARCH,
        general: ROUTES.KARYAWAN_GENERAL,
        profile: ROUTES.KARYAWAN_PROFILE,
        edit: ROUTES.KARYAWAN_EDIT
      };
  }

  function employeeRoutes(nrp) {
    const routes = employeeRoutesBySource(isOutsourceNrp(nrp));
    return {
      search: routes.search(nrp),
      general: routes.general,
      profile: routes.profile,
      edit: routes.edit
    };
  }


  /** Shorthand: select internal or OS route based on NRP type. */
  function routeByNrp(nrp, internalRoute, osRoute) {
    return isOutsourceNrp(nrp) ? osRoute : internalRoute;
  }

  function attendanceUrl(bulan, tahun, nrp) {
    return routeByNrp(nrp, ROUTES.TABEL_HADIR, ROUTES.TABEL_HADIR_OS)(bulan, tahun, nrp);
  }

  function distribusiUrl(nrp) {
    return routeByNrp(nrp, ROUTES.DISTRIBUSI, ROUTES.DISTRIBUSI_OS)(nrp);
  }

  /** Build auto-distribusi link for a given date + shift. */
  function buildDistribusiLink(ctx, tglAwal, shiftVal, tglAkhir = null) {
    if (!ctx.nrp) return '';
    const base = distribusiUrl(ctx.nrp);

    // If tglAwal is already a full date (YYYY-MM-DD), use it directly
    const dAwal = (tglAwal && tglAwal.includes('-')) ? tglAwal : `${ctx.tahun}-${String(ctx.bulan).padStart(2, '0')}-${String(tglAwal).padStart(2, '0')}`;
    const dAkhir = tglAkhir ? (tglAkhir.includes('-') ? tglAkhir : `${ctx.tahun}-${String(ctx.bulan).padStart(2, '0')}-${String(tglAkhir).padStart(2, '0')}`) : dAwal;

    return `${base}?qm_auto_distribusi=1&nrp=${encodeURIComponent(ctx.nrp)}&tanggal_awal=${encodeURIComponent(dAwal)}&tanggal_akhir=${encodeURIComponent(dAkhir)}&bagian=${encodeURIComponent(ctx.bagian)}&seksi=${encodeURIComponent(ctx.seksi)}&shift=${encodeURIComponent(shiftVal)}`;
  }


  /** Build Kehadiran (Barcode) link for a given context. */
  function buildKehadiranLink(ctx) {
    if (!ctx.nrp) return '';
    return (isOutsourceNrp(ctx.nrp) ? ROUTES.ABSEN_BARCODE_OS : ROUTES.ABSEN_BARCODE)(ctx.tahun, String(ctx.bulan).padStart(2, '0'), ctx.nrp);
  }

  function barcodePageUrl(nrp) {
    return isOutsourceNrp(nrp) ? ROUTES.ABSEN_BARCODE_OS_PAGE : ROUTES.ABSEN_BARCODE_PAGE;
  }


  function buildSpklOnlineUrl(ctx, minDate, maxDate) {
    const bulan = String(ctx.bulan).padStart(2, '0');
    const min = String(minDate).padStart(2, '0');
    const max = String(maxDate).padStart(2, '0');
    return ROUTES.SPKL_ONLINE(ctx.tahun, bulan, min, max, ctx.nrp);
  }

  function employeeEditUrl(nrp, id) {
    return employeeRoutes(nrp).edit(id);
  }

  function employeeSearchUrl(query, isOS) {
    return employeeRoutesBySource(isOS).search(encodeURIComponent(String(query || '').trim()));
  }

  function toAbsoluteHrisUrl(url) {
    if (!url) return '';
    if (/^https?:\/\//i.test(url)) return url;
    return `${ROUTES.BASE}${url.startsWith('/') ? '' : '/'}${url}`;
  }

  function extractEmployeeIdFromUrl(url) {
    if (!url) return '';
    try {
      const cleanUrl = String(url).trim();
      const absoluteUrl = cleanUrl.startsWith('http') ? cleanUrl : 'https://hris.kti.co.id/' + cleanUrl.replace(/^\//, '');
      const urlObj = new URL(absoluteUrl);
      const queryId = urlObj.searchParams.get('id');
      if (queryId) return queryId;
    } catch (e) { }

    const clean = String(url || '');
    const routeMatch = clean.match(/\/(?:general|profile|editgeneral|edit)\/([^/?#]+)/i);
    if (routeMatch) return routeMatch[1];
    return clean.split('?')[0].split('/').filter(Boolean).pop() || '';
  }

  function spklBaseUrl(nrp) {
    return routeByNrp(nrp, ROUTES.SPKL_BASE, ROUTES.SPKL_OS_BASE);
  }

  function spklListUrl(nrp, bulan, tahun) {
    const bulanStr = String(parseInt(bulan, 10)).padStart(2, '0');
    return `${spklBaseUrl(nrp)}?tahun=${encodeURIComponent(tahun)}&bulan=${encodeURIComponent(bulanStr)}&kode_bagian=&kode_seksi=&kode_group=&nrp=${encodeURIComponent(nrp)}`;
  }

  function spklCreateUrl(nrp) {
    return routeByNrp(nrp, ROUTES.SPKL_CREATE, ROUTES.SPKL_OS_CREATE);
  }

  function spklAddUrl(kind) {
    return kind === 'outsource' ? ROUTES.SPKL_OS_ADD : ROUTES.SPKL_ADD;
  }

  function absenCreateUrl(nrp) {
    return routeByNrp(nrp, ROUTES.ABSEN_BARCODE_CREATE, ROUTES.ABSEN_BARCODE_OS_CREATE);
  }

  function getAbsenCreateUrlByKind(kind) {
    return kind === 'outsource' ? ROUTES.ABSEN_BARCODE_OS_CREATE : ROUTES.ABSEN_BARCODE_CREATE;
  }

  function absenAddUrl(nrp) {
    return routeByNrp(nrp, ROUTES.ABSEN_BARCODE_ADD, ROUTES.ABSEN_BARCODE_OS_ADD);
  }

  /* ============================================================
   * 5. API LAYER
   * ============================================================ */


  /** Fetch wrapper with timeout. Returns response text. */
  async function hrisFetch(url, timeout = 15000) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), timeout);
    try {
      const res = await fetch(url, {
        signal: ctrl.signal,
        credentials: 'include'
      });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return await res.text();
    } finally {
      clearTimeout(timer);
    }
  }


  /** Unified fetch with AbortController timeout. Returns Response. */
  async function fetchWithTimeout(url, options = {}, timeoutMs = 15000) {
    const controller = new AbortController();
    const externalSignal = options.signal;
    const onExternalAbort = () => controller.abort(externalSignal.reason);
    const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
    try {
      if (externalSignal) {
        if (externalSignal.aborted) controller.abort(externalSignal.reason);
        else externalSignal.addEventListener('abort', onExternalAbort, { once: true });
      }
      const { signal: _ignoredSignal, ...restOptions } = options;
      const response = await fetch(url, {
        ...restOptions,
        signal: controller.signal,
        credentials: 'include'
      });
      return response;
    } finally {
      clearTimeout(timeoutId);
      if (externalSignal) externalSignal.removeEventListener('abort', onExternalAbort);
    }
  }

  function isAbortError(err) {
    return !!err && (err.name === 'AbortError' || /abort/i.test(String(err.message || '')));
  }

  function throwIfCancelled() {
    if (state.cancelRequested) throw new DOMException('User cancelled operation', 'AbortError');
  }

  function formatElapsedMs(ms) {
    const totalSeconds = Math.max(0, Math.floor(ms / 1000));
    const minutes = Math.floor(totalSeconds / 60);
    const seconds = totalSeconds % 60;
    return minutes > 0
      ? `${minutes}m ${String(seconds).padStart(2, '0')}s`
      : `${seconds}s`;
  }

  function stopBackgroundHeartbeat() {
    if (state.backgroundHeartbeatTimer) {
      clearInterval(state.backgroundHeartbeatTimer);
      state.backgroundHeartbeatTimer = null;
    }
  }

  function startBackgroundHeartbeat() {
    stopBackgroundHeartbeat();
    const startedAt = Date.now();
    state.backgroundHeartbeatTimer = setInterval(() => {
      if (state.cancelRequested) return;
      const elapsed = Date.now() - startedAt;
      const progress = Math.min(88, 70 + Math.floor(elapsed / TIMING.BACKGROUND_HEARTBEAT_INTERVAL));
      UI.setGlobalProgress(progress, `Menunggu respon distribusi... ${formatElapsedMs(elapsed)} berlalu. Server HRIS masih memproses.`, true);
    }, TIMING.BACKGROUND_HEARTBEAT_INTERVAL);
  }

  function beginCancelableDistributionFlow() {
    state.cancelRequested = false;
    state.activeCancelableFlow = true;
    state.activeAbortController = null;
    stopBackgroundHeartbeat();
  }

  function clearCancelableDistributionFlow() {
    stopBackgroundHeartbeat();
    state.activeAbortController = null;
    state.activeCancelableFlow = false;
    state.cancelRequested = false;
  }

  function extractForm(doc, overrides = {}) {
    const form = doc.querySelector('form');
    if (!form) throw new Error('Form tidak ditemukan.');
    const params = new URLSearchParams();
    const headers = { 'Content-Type': 'application/x-www-form-urlencoded', 'X-Requested-With': 'XMLHttpRequest' };
    const csrfMeta = doc.querySelector('meta[name="csrf-token"], meta[name="csrf-test-name"]');
    if (csrfMeta) {
      const tokenValue = csrfMeta.getAttribute('content');
      if (!tokenValue) throw new Error('Token CSRF tidak ditemukan. Struktur halaman mungkin berubah.');
      headers['X-CSRF-TOKEN'] = tokenValue;
    } else {
      const csrfHidden = form.querySelector('input[name="_token"], input[name="csrf_token"], input[name="csrf-token"]');
      if (!csrfHidden || !csrfHidden.value) throw new Error('Token CSRF tidak ditemukan. Struktur halaman mungkin berubah.');
      headers['X-CSRF-TOKEN'] = csrfHidden.value;
    }
    form.querySelectorAll('input, select, textarea').forEach(el => {
      if (!el.name || el.type === 'submit' || el.type === 'button') return;
      if ((el.type === 'checkbox' || el.type === 'radio') && !el.checked) return;
      const val = overrides.hasOwnProperty(el.name) ? overrides[el.name] : el.value;
      params.append(el.name, val);
    });
    for (const [key, val] of Object.entries(overrides)) {
      if (!form.querySelector(`[name="${key}"]`)) params.append(key, val);
    }
    const action = form.getAttribute('action') || window.location.href;
    const url = action.startsWith('http') ? action : `https://hris.kti.co.id${action}`;
    return { url, headers, body: params.toString() };
  }

  function runWithProgress(title, initialMsg, steps) {
    UI.showGlobalLoader(title, initialMsg);
    const execute = async () => {
      for (const step of steps) {
        UI.setGlobalProgress(step.progress, step.message);
        if (step.action) await step.action();
      }
      UI.setGlobalProgress(100, 'Selesai');
    };
    return execute().catch(e => {
      UI.showResult('danger', 'Gagal', e.message);
      throw e;
    }).finally(() => UI.hideGlobalLoader());
  }

  function employeeUrlSet(nrp) {
    const routeSet = employeeRoutes(nrp);
    return {
      isOS: isOutsourceNrp(nrp),
      searchUrl: routeSet.search,
      buildGeneralUrl: function (encodedNrp) {
        return routeSet.general(encodedNrp);
      },
      buildProfileUrl: function (encodedNrp) {
        return routeSet.profile(encodedNrp);
      }
    };
  }

  function findEmployeeDetailLink(doc, nrp) {
    let detailUrl = '';
    const rows = doc.querySelectorAll('table tbody tr');
    rows.forEach(function (row) {
      if (row.textContent.includes(nrp)) {
        const links = Array.from(row.querySelectorAll('a'));
        // Prioritize links that look like detail buttons or contain 'Detail'
        const btn = links.find(a => {
          const txt = a.textContent.trim();
          const cls = a.className;
          const href = a.getAttribute('href') || '';
          const isGenLink = href.includes('karyawan/general/') || href.includes('karyawanoutsource/general/');
          if (!isGenLink) return false;
          const id = extractEmployeeIdFromUrl(href);
          return isValidEmployeeId(id) && (txt.includes('Detail') || cls.includes('btn-info') || cls.includes('btn-primary'));
        });
        if (btn && btn.getAttribute('href')) detailUrl = btn.getAttribute('href');
      }
    });
    if (!detailUrl) {
      const allLinks = Array.from(doc.querySelectorAll('a'));
      const fallback = allLinks.find(a => {
        const href = a.getAttribute('href') || '';
        const isGenLink = href.includes('karyawan/general/') || href.includes('karyawanoutsource/general/');
        const isReport = href.includes('rekap') || href.includes('laporan') || href.includes('pembayaran');
        if (!isGenLink || isReport) return false;
        const id = extractEmployeeIdFromUrl(href);
        return isValidEmployeeId(id) && (a.textContent.includes('Detail') || a.classList.contains('btn-info'));
      });
      if (fallback) detailUrl = fallback.getAttribute('href');
    }
    return detailUrl;
  }


  /** Unified field value extractor (input or select). */
  function extractFieldValue(doc, fieldName) {
    const el = doc.querySelector(`[name="${fieldName}"]`);
    if (!el) return '';
    if (el.tagName === 'SELECT') {
      const opt = el.querySelector('option[selected]') || el.options[el.selectedIndex];
      return (opt ? opt.value || opt.textContent : el.value).trim();
    }
    return el.value.trim();
  }

  function extractEmployeeName(doc) {
    const selectors = ['input[name="name"]', 'input[name="nama"]', 'input[name="nama_karyawan"]'];
    for (const sel of selectors) {
      const el = doc.querySelector(sel);
      if (el) {
        const val = el.value || el.getAttribute('value');
        if (val && val.trim()) return val.trim();
      }
    }
    const nodes = doc.querySelectorAll('th, td, label, span');
    for (let i = 0; i < nodes.length; i++) {
      const text = nodes[i].textContent.trim().toLowerCase();
      // Match exact label or common variants
      if (text === 'nama' || text === 'nama karyawan' || text === 'name' || text === 'nama lengkap') {
        const next = nodes[i].nextElementSibling;
        if (next && next.textContent.trim()) return next.textContent.trim();
        const parentNext = nodes[i].parentElement?.nextElementSibling;
        if (parentNext && parentNext.textContent.trim()) return parentNext.textContent.trim();
      }
    }
    return '';
  }

  function normalizeProfileLookup(text) {
    return String(text || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  }

  function readProfileElementValue(el) {
    if (!el) return '';
    if (el.tagName === 'SELECT') {
      const opt = el.querySelector('option[selected]') || el.options[el.selectedIndex];
      return String(opt ? opt.textContent || opt.value : el.value).trim();
    }
    if (el.tagName === 'TEXTAREA') return String(el.value || el.textContent || '').trim();
    return String(el.value || el.getAttribute('value') || el.textContent || '').trim();
  }

  function extractProfileFieldByNames(doc, names) {
    for (const name of names) {
      const selector = `[name="${name}"], [name*="${name}"], #${name}, [id*="${name}"]`;
      const el = doc.querySelector(selector);
      const value = readProfileElementValue(el);
      if (value) return value;
    }
    return '';
  }

  function findProfileRowValue(doc, labels) {
    const normalizedLabels = labels.map(normalizeProfileLookup);

    const tableRows = Array.from(doc.querySelectorAll('tr'));
    for (const row of tableRows) {
      const cells = Array.from(row.querySelectorAll('th, td'));
      if (cells.length < 2) continue;
      const label = normalizeProfileLookup(cells[0].textContent);
      if (!label) continue;
      if (normalizedLabels.some(candidate => label.includes(candidate))) {
        const value = cells.slice(1).map(cell => normalizeEmployeeCellText(cell)).join(' ').trim();
        if (value) return value;
      }
    }

    const labelNodes = Array.from(doc.querySelectorAll('label, .control-label, .col-form-label'));
    for (const labelNode of labelNodes) {
      const label = normalizeProfileLookup(labelNode.textContent);
      if (!label) continue;
      if (!normalizedLabels.some(candidate => label.includes(candidate))) continue;

      const forId = labelNode.getAttribute('for');
      if (forId) {
        const value = readProfileElementValue(doc.getElementById(forId));
        if (value) return value;
      }

      const parent = labelNode.closest('.form-group, .form-row, .row, .col, .form-line') || labelNode.parentElement;
      if (parent) {
        const valueNode = parent.querySelector('input, textarea, select, .form-control-plaintext, .form-control-static, p, span');
        const value = readProfileElementValue(valueNode);
        if (value && normalizeProfileLookup(value) !== label) return value;
      }

      const siblingValue = readProfileElementValue(labelNode.nextElementSibling);
      if (siblingValue) return siblingValue;
    }

    const genericNodes = Array.from(doc.querySelectorAll('th, td, span'));
    for (const node of genericNodes) {
      const label = normalizeProfileLookup(node.textContent);
      if (!label) continue;
      if (!normalizedLabels.some(candidate => label === candidate || label.startsWith(candidate + ' '))) continue;
      const siblingValue = readProfileElementValue(node.nextElementSibling);
      if (siblingValue) return siblingValue;
      const parentNext = readProfileElementValue(node.parentElement?.nextElementSibling);
      if (parentNext) return parentNext;
    }

    return '';
  }

  function extractProfileField(doc, config) {
    const byName = extractProfileFieldByNames(doc, config.names || []);
    if (byName) return byName;
    const byLabel = findProfileRowValue(doc, config.labels || []);
    return byLabel || '';
  }

  function parseEmployeeProfile(doc) {
    return {
      telepon: extractProfileField(doc, { names: ['telepon', 'no_telp', 'nomor_telepon', 'hp', 'phone'], labels: ['telepon', 'nomor telepon', 'no telepon', 'hp', 'phone'] }),
      tanggalLahir: extractProfileField(doc, { names: ['tanggal_lahir', 'tgl_lahir', 'birth_date'], labels: ['tanggal lahir', 'tgl lahir', 'birth date'] }),
      nik: extractProfileField(doc, { names: ['nik', 'no_ktp', 'nomor_ktp'], labels: ['nik', 'no ktp', 'nomor ktp'] }),
      noKk: extractProfileField(doc, { names: ['no_kk', 'nomor_kk', 'kartu_keluarga'], labels: ['no kk', 'nomor kk', 'kartu keluarga'] }),
      alamatLengkap: extractProfileField(doc, { names: ['alamat', 'alamat_lengkap', 'address'], labels: ['alamat lengkap', 'alamat', 'address'] }),
      rtRw: extractProfileField(doc, { names: ['rt_rw', 'rtrw', 'rt'], labels: ['rt rw', 'rt/rw'] }),
      kelurahan: extractProfileField(doc, { names: ['kelurahan', 'desa'], labels: ['kelurahan', 'desa'] }),
      kecamatan: extractProfileField(doc, { names: ['kecamatan'], labels: ['kecamatan'] }),
      kotaTinggal: extractProfileField(doc, { names: ['kota_tinggal', 'kota', 'kabupaten'], labels: ['kota tinggal', 'kota', 'kabupaten'] }),
      statusMarital: extractProfileField(doc, { names: ['status_marital', 'status_perkawinan', 'marital_status'], labels: ['status marital', 'status perkawinan', 'marital'] }),
      agama: extractProfileField(doc, { names: ['agama', 'religion'], labels: ['agama', 'religion'] }),
      pendidikan: extractProfileField(doc, { names: ['pendidikan', 'pendidikan_terakhir', 'education'], labels: ['pendidikan', 'pendidikan terakhir', 'education'] })
    };
  }

  /** Unified employee data fetcher with sessionStorage cache. */
  async function fetchEmployee(nrp, knownId = null, knownIsOS = null) {
    const cached = readEmployeeCache(nrp);
    if (cached) {
      if (!knownId || cached.id === knownId) {
        return { found: true, ...cached };
      } else {
        clearEmployeeCache(nrp);
      }
    }
    
    let id = knownId;
    let isOS = knownIsOS;

    if (!id) {
      const urls = employeeUrlSet(nrp);
      isOS = urls.isOS;
      let searchDoc, detailUrl;

      // Retry logic for NRP search (Bug Fix 1)
      for (let attempt = 1; attempt <= 3; attempt++) {
        try {
          const html = await hrisFetch(urls.searchUrl);
          searchDoc = parseHTML(html);
          detailUrl = findEmployeeDetailLink(searchDoc, nrp);
          if (detailUrl) break;
        } catch (e) {
          if (attempt === 3) throw e;
        }
        if (attempt < 3) await new Promise(r => setTimeout(r, 500 * attempt));
      }

      if (!detailUrl) return { found: false };
      id = extractEmployeeIdFromUrl(detailUrl);
    } else {
      // If we provided an ID but not isOS, try to guess
      if (isOS === null) {
        isOS = isOutsourceNrp(nrp);
      }
    }

    const routeSet = employeeRoutesBySource(isOS);
    const [genHtml, profHtml] = await Promise.all([
      hrisFetch(routeSet.general(id)),
      hrisFetch(routeSet.profile(id))
    ]);
    const doc = parseHTML(genHtml);
    const profDoc = parseHTML(profHtml);

    // Find Edit General/Edit URL from the page buttons
    const editUrl = routeSet.edit(id);

    const emp = {
      found: true,
      id: id,
      editUrl: editUrl,
      jk: extractFieldValue(doc, 'kode_jam_kerja').split('-')[0].trim(),
      KK: extractFieldValue(doc, 'kode_kalender_kerja'),
      nama: extractEmployeeName(profDoc) || extractEmployeeName(doc),
      bagian: extractFieldValue(doc, 'kode_bagian') || '',
      seksi: extractFieldValue(doc, 'kode_seksi') || '',
      group: extractFieldValue(doc, 'kode_group') || ''
    };
    writeEmployeeCache(nrp, emp);
    return emp;
  }

  function normalizeEmployeeHeader(text) {
    return String(text || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  }

  function normalizeEmployeeCellText(cell) {
    return String(cell?.textContent || '').replace(/\s+/g, ' ').trim();
  }

  function firstMatchedHeaderIndex(headers, keywords) {
    return headers.findIndex(header => keywords.some(keyword => header.includes(keyword)));
  }

  function getEmployeeCellValue(cells, headers, keywords) {
    const index = firstMatchedHeaderIndex(headers, keywords);
    if (index === -1 || !cells[index]) return '';
    return normalizeEmployeeCellText(cells[index]);
  }

  function getFirstEmployeeNrp(text) {
    const match = String(text || '').match(/\b\d{8}\b|\b\d{4}\b/);
    return match ? match[0] : '';
  }

  function isValidNrp(value) {
    return /^\d{4}$|^\d{8}$/.test(String(value || '').trim());
  }

  function updateKaryawanResultCache(key, emp) {
    state.karyawanResults = state.karyawanResults.map(item => {
      if (item.key !== key) return item;
      return {
        ...item,
        nama: emp.nama || item.nama,
        bagian: emp.bagian || item.bagian,
        seksi: emp.seksi || item.seksi,
        group: emp.group || item.group,
        jk: emp.jk || item.jk,
        KK: emp.KK || item.KK,
        editUrl: emp.editUrl || item.editUrl
      };
    });
  }

  function findKaryawanResult(key) {
    return state.karyawanResults.find(item => item.key === key) || null;
  }

  function buildKaryawanResultKey(source, id, nrp, fallbackName) {
    if (id) return `${source}:${id}`;
    if (nrp) return `${source}:${nrp}`;
    return `${source}:${String(fallbackName || '').toLowerCase()}`;
  }

  function parseEmployeeSearchRow(row, headers, isOS) {
    const cells = Array.from(row.querySelectorAll('td'));
    if (cells.length === 0) return null;

    const links = Array.from(row.querySelectorAll('a[href]'));
    const detailLink = links.find(link => {
      const href = link.getAttribute('href') || '';
      const isGenLink = href.includes('karyawan/general/') || href.includes('karyawanoutsource/general/');
      if (!isGenLink) return false;
      const id = extractEmployeeIdFromUrl(href);
      return isValidEmployeeId(id);
    }) || links.find(link => {
      const href = link.getAttribute('href') || '';
      const id = extractEmployeeIdFromUrl(href);
      return isValidEmployeeId(id) && /detail|profile|edit/i.test(normalizeEmployeeCellText(link));
    });
    if (!detailLink) return null;

    const href = toAbsoluteHrisUrl(detailLink.getAttribute('href') || '');
    const id = extractEmployeeIdFromUrl(href);
    if (!id || !isValidEmployeeId(id)) return null;

    const nrpByHeader = getEmployeeCellValue(cells, headers, ['nrp']);
    const nrp = nrpByHeader || getFirstEmployeeNrp(row.textContent);
    if (nrp) {
      sessionStorage.setItem('qm_is_os_' + nrp, isOS ? 'true' : 'false');
    }
    const nama = getEmployeeCellValue(cells, headers, ['nama', 'name'])
      || normalizeEmployeeCellText(cells[Math.max(0, firstMatchedHeaderIndex(headers, ['nrp']) + 1)]);
    const bagian = getEmployeeCellValue(cells, headers, ['bagian', 'dept', 'departemen']);
    const seksi = getEmployeeCellValue(cells, headers, ['seksi', 'section']);
    const group = getEmployeeCellValue(cells, headers, ['group', 'grup']);
    const routeSet = employeeRoutesBySource(isOS);
    const source = isOS ? 'outsource' : 'internal';

    return {
      key: buildKaryawanResultKey(source, id, nrp, nama),
      source,
      sourceLabel: isOS ? 'Outsource' : 'Internal',
      nrp,
      nama,
      bagian,
      seksi,
      group,
      id,
      generalUrl: routeSet.general(id),
      profileUrl: routeSet.profile(id),
      editUrl: routeSet.edit(id)
    };
  }

  function parseEmployeeSearchResults(doc, isOS) {
    const headerCells = Array.from(doc.querySelectorAll('table thead th'));
    const headers = headerCells.map(cell => normalizeEmployeeHeader(cell.textContent));
    const rows = Array.from(doc.querySelectorAll('table tbody tr'));
    return rows.map(row => parseEmployeeSearchRow(row, headers, isOS)).filter(Boolean);
  }

  async function searchEmployees(query) {
    const trimmed = String(query || '').trim();
    if (!trimmed) return [];

    let isRegex = false;
    let regex = null;

    // Check if the input looks like a regex pattern (has special regex characters)
    const hasRegexChars = /[\.\*\+\?\^\$\{\}\(\)\|\[\]\\]/.test(trimmed);
    if (hasRegexChars) {
      try {
        regex = new RegExp(trimmed, 'i');
        isRegex = true;
      } catch (e) {
        // Invalid regex, will fallback to plain text matching
      }
    }

    // Clean server query: strip non-alphanumeric and special symbols for the HRIS server search
    // but keep space to allow searching name/terms
    let serverQuery = trimmed.replace(/[\.\*\+\?\^\$\{\}\(\)\|\[\]\\_\-\/]/g, ' ').replace(/\s+/g, ' ').trim();
    if (!serverQuery) {
      serverQuery = trimmed; // fallback if stripped completely
    }

    // Always query both internal and outsource databases to support partial matches for any query!
    const targets = [false, true];

    const responses = await Promise.all(targets.map(async isOS => {
      try {
        const html = await hrisFetch(employeeSearchUrl(serverQuery, isOS));
        return parseEmployeeSearchResults(parseHTML(html), isOS);
      } catch (e) {
        Logger.warn(`Gagal mencari karyawan (${isOS ? 'Outsource' : 'Internal'}): ${e.message}`);
        return [];
      }
    }));

    const map = new Map();
    responses.flat().forEach(item => {
      if (!map.has(item.key)) map.set(item.key, item);
    });

    let results = Array.from(map.values());

    // Local Regex / Normal Match Filtering to ensure extreme accuracy
    if (isRegex && regex) {
      results = results.filter(item => {
        const targetString = `${item.nrp} ${item.nama} ${item.divisi} ${item.bagian} ${item.group}`.toLowerCase();
        return regex.test(targetString) || regex.test(item.nrp) || regex.test(item.nama);
      });
    } else {
      // Alphanumeric plain text fallback check: if query has non-alphanumeric removed,
      // let's ensure we only show items matching the clean terms to match perfectly.
      const terms = serverQuery.toLowerCase().split(/\s+/).filter(Boolean);
      if (terms.length > 0) {
        results = results.filter(item => {
          const targetString = `${item.nrp} ${item.nama} ${item.divisi} ${item.bagian} ${item.group}`.toLowerCase();
          return terms.every(term => targetString.includes(term));
        });
      }
    }

    return results;
  }

  function getKaryawanPanelMode(key) {
    return state.karyawanActivePanel?.key === key ? state.karyawanActivePanel.mode : '';
  }

  function renderKaryawanDetail(result) {
    const detail = state.karyawanDetail;
    if (!detail || detail.key !== result.key || getKaryawanPanelMode(result.key) !== 'detail') return '';
    if (detail.loading) {
      return '<div class="qm-karyawan-panel"><div class="qm-flex qm-items-center qm-gap-s"><span class="qm-spinner"></span><span>Memuat data profil karyawan...</span></div></div>';
    }
    if (detail.error) {
      return `<div class="qm-karyawan-panel"><div class="qm-text-danger">${escapeHtml(detail.error)}</div></div>`;
    }

    const profile = detail.profile || {};
    const fields = [
      ['Telepon', profile.telepon],
      ['Tanggal Lahir', profile.tanggalLahir],
      ['NIK', profile.nik],
      ['No. KK', profile.noKk],
      ['Alamat Lengkap', profile.alamatLengkap],
      ['RT/RW', profile.rtRw],
      ['Kelurahan', profile.kelurahan],
      ['Kecamatan', profile.kecamatan],
      ['Kota Tinggal', profile.kotaTinggal],
      ['Status Marital', profile.statusMarital],
      ['Agama', profile.agama],
      ['Pendidikan', profile.pendidikan]
    ];

    return `
      <div class="qm-karyawan-panel" style="padding: 0 8px 0 4px; border-radius: 4px;">
        <div style="display: grid; grid-template-columns: 110px 1fr; gap: 6px 8px; font-size: 0.85rem; line-height: 1.4;">
          ${fields.map(([label, value]) => `
            <div style="color: var(--color-text-muted);">${escapeHtml(label)}</div>
            <div style="font-weight: 600; color: var(--color-text); word-break: break-word;">${escapeHtml(value || '-')}</div>
          `).join('')}
        </div>
      </div>
    `;
  }

  function renderKaryawanEditor(result) {
    const editor = state.karyawanEditor;
    if (!editor || editor.key !== result.key || getKaryawanPanelMode(result.key) !== 'edit') return '';
    if (editor.loading) {
      return '<div class="qm-karyawan-panel"><div class="qm-flex qm-items-center qm-gap-s"><span class="qm-spinner"></span><span>Memuat opsi edit data...</span></div></div>';
    }
    if (editor.error) {
      return `<div class="qm-karyawan-panel"><div class="qm-text-danger">${escapeHtml(editor.error)}</div></div>`;
    }

    const jkOptions = editor.jkOptions.map(opt => `<option value="${escapeHtml(opt.val)}" ${opt.selected ? 'selected' : ''}>${escapeHtml(opt.txt)}</option>`).join('');
    const kkOptions = editor.kkOptions.map(opt => `<option value="${escapeHtml(opt.val)}" ${opt.selected ? 'selected' : ''}>${escapeHtml(opt.txt)}</option>`).join('');
    const saveDisabled = editor.saving || (editor.jkOptions.length === 0 && editor.kkOptions.length === 0) ? 'disabled' : '';

    return `
      <div class="qm-karyawan-panel" data-key="${escapeHtml(result.key)}" style="padding: 0 4px;">
        <div style="display: flex; gap: 12px; margin-bottom: 16px;">
          <div style="flex:1; background: var(--color-surface-soft); padding: 10px 12px; border-radius: 8px; font-size: 0.8rem; border: 1px solid var(--color-border);">
            <span style="color: var(--color-text-muted); display:block; margin-bottom: 2px;">JK Saat Ini</span>
            <strong style="font-size: 0.95rem; color: var(--color-text);">${escapeHtml(editor.emp?.jk || '-')}</strong>
          </div>
          <div style="flex:1; background: var(--color-surface-soft); padding: 10px 12px; border-radius: 8px; font-size: 0.8rem; border: 1px solid var(--color-border);">
            <span style="color: var(--color-text-muted); display:block; margin-bottom: 2px;">KK Saat Ini</span>
            <strong style="font-size: 0.95rem; color: var(--color-text);">${escapeHtml(editor.emp?.KK || '-')}</strong>
          </div>
        </div>
        ${editor.notice ? `<div style="font-size: 0.85rem; color: #d32f2f; margin-bottom: 12px; padding: 8px; background: #ffebee; border-radius: 6px;">${escapeHtml(editor.notice)}</div>` : ''}
        <div style="display: flex; flex-direction: column; gap: 12px;">
          <div class="qm-form-group" style="margin-bottom: 0;">
            <label class="qm-form-label">Jam Kerja Baru</label>
            <select class="qm-premium-input qm-premium-select qm-karyawan-jk-select" data-key="${escapeHtml(result.key)}" ${editor.jkOptions.length === 0 ? 'disabled' : ''}>
              ${jkOptions || '<option value="">Tidak ada opsi JK</option>'}
            </select>
          </div>
          <div class="qm-form-group" style="margin-bottom: 0;">
            <label class="qm-form-label">Kalender Kerja Baru</label>
            <select class="qm-premium-input qm-premium-select qm-karyawan-kk-select" data-key="${escapeHtml(result.key)}" ${editor.kkOptions.length === 0 ? 'disabled' : ''}>
              ${kkOptions || '<option value="">Tidak ada opsi KK</option>'}
            </select>
          </div>
        </div>
        <button type="button" class="qm-btn-premium-primary qm-karyawan-save-btn" data-key="${escapeHtml(result.key)}" ${saveDisabled} style="width: 100%; justify-content: center; margin-top: 16px; padding: 10px 24px !important; font-size: 0.95rem !important;">${editor.saving ? 'Menyimpan...' : 'Simpan Perubahan'}</button>
      </div>
    `;
  }

  function renderKaryawanExpandedPanel(result) {
    const mode = getKaryawanPanelMode(result.key);
    if (mode === 'detail') return renderKaryawanDetail(result);
    if (mode === 'edit') return renderKaryawanEditor(result);
    return '';
  }

  function getInitials(name) {
    if (!name) return '??';
    const parts = name.split(' ').filter(Boolean);
    if (parts.length > 1 && parts[0][0] && parts[1][0]) {
      return (parts[0][0] + parts[1][0]).toUpperCase();
    }
    return name.slice(0, 2).toUpperCase();
  }

  function renderKaryawanResults(onlyLeftPanel = false) {
    const wrap = uiAdapter.get('karyawanResults');
    const btn = uiAdapter.get('karyawanSearchButton');
    if (!wrap) return;

    if (btn) {
      btn.disabled = state.karyawanLoading;
      btn.textContent = state.karyawanLoading ? 'Mencari...' : 'Cari Karyawan';
    }

    const previewContainer = document.getElementById('qm-karyawan-live-preview');
    const directoryTitle = document.getElementById('qm-karyawan-directory-title');

    if (!onlyLeftPanel) {
      // 1. Loading State
      if (state.karyawanLoading) {
        uiAdapter.html('karyawanResults', '<div class="qm-card qm-karyawan-empty" style="padding: 20px; font-size: 0.85rem; text-align: center; color: var(--color-text-soft); display: flex; align-items: center; justify-content: center; gap: 8px; width: 100%;"><span class="qm-spinner"></span><span>Mencari data karyawan...</span></div>');
        return;
      }

      // 2. Error State
      if (state.karyawanError) {
        uiAdapter.html('karyawanResults', `<div class="qm-card qm-karyawan-empty qm-text-danger" style="padding: 20px; font-size: 0.85rem; text-align: center; width: 100%;">${escapeHtml(state.karyawanError)}</div>`);
        return;
      }

      // 3. Empty Search / No Query State
      if (!state.karyawanQuery) {
        if (directoryTitle) directoryTitle.textContent = 'Hasil Pencarian';
        uiAdapter.html('karyawanResults', '<div class="qm-card qm-karyawan-empty" style="padding: 20px; font-size: 0.85rem; text-align: center; color: var(--color-text-soft); width: 100%;">Gunakan form pencarian untuk menemukan data karyawan.</div>');

        const leftPanel = document.getElementById('qm-karyawan-left-panel');
        if (leftPanel) {
          leftPanel.innerHTML = `
            <div style="font-size: 0.85rem; color: var(--color-text-muted); line-height: 1.5; padding: 16px; background: #faf8f5; border-radius: 8px; border: 1px dashed var(--color-border);">
              Pencarian NRP akan memeriksa database internal & outsource, lalu menampilkan profil detail & quick edit JK/KK.
            </div>
          `;
        }
        return;
      }

      // 4. No Results Found State
      if (state.karyawanResults.length === 0) {
        if (directoryTitle) directoryTitle.textContent = 'Hasil Pencarian';
        uiAdapter.html('karyawanResults', `<div class="qm-card qm-karyawan-empty" style="padding: 20px; font-size: 0.85rem; text-align: center; color: var(--color-text-soft); width: 100%;">Tidak ada data karyawan yang cocok untuk <strong>${escapeHtml(state.karyawanQuery)}</strong>.</div>`);

        const leftPanel = document.getElementById('qm-karyawan-left-panel');
        if (leftPanel) {
          leftPanel.innerHTML = `
            <div style="font-size: 0.85rem; color: var(--color-text-muted); line-height: 1.5; padding: 16px; background: #faf8f5; border-radius: 8px; border: 1px dashed var(--color-border);">
              Pencarian NRP akan memeriksa database internal & outsource, lalu menampilkan profil detail & quick edit JK/KK.
            </div>
          `;
        }
        return;
      }

      // 5. Results Loaded State
      if (directoryTitle) {
        directoryTitle.textContent = `Hasil Pencarian (${state.karyawanResults.length})`;
      }
    }

    // Determine active employee (default to none)
    if (!state.karyawanActivePanel || !state.karyawanResults.some(r => r.key === state.karyawanActivePanel.key)) {
      state.karyawanActivePanel = null;
    }

    const activeKey = state.karyawanActivePanel?.key;
    const activeResult = activeKey ? findKaryawanResult(activeKey) : null;
    const activeMode = state.karyawanActivePanel?.mode || 'detail';

    if (onlyLeftPanel) {
      // Just update active styling on card elements in the DOM without full re-rendering
      wrap.querySelectorAll('.qm-directory-item').forEach(card => {
        const key = card.dataset.key;
        if (key === activeKey) {
          card.classList.add('active');
          card.style.borderColor = 'var(--color-accent-strong)';
          card.style.background = '#f0f8ff';
          card.style.boxShadow = '0 4px 12px rgba(0,0,0,0.05)';
        } else {
          card.classList.remove('active');
          card.style.borderColor = '';
          card.style.background = '';
          card.style.boxShadow = '';
        }
      });
    } else {
      // B. Group and Render Matching Results List in Directory
      const internalResults = state.karyawanResults.filter(r => r.source !== 'outsource');
      const outsourceResults = state.karyawanResults.filter(r => r.source === 'outsource');

      if (!state.karyawanGroupsExpanded) {
        state.karyawanGroupsExpanded = { internal: false, outsource: false };
      }

      function renderEmployeeCardHtml(result) {
        const initials = getInitials(result.nama);
        const isOS = result.source === 'outsource';
        const sourceBadge = `<span class="qm-preview-tag ${isOS ? 'tag-red' : 'tag-blue'}" style="font-size: 0.7rem; padding: 2px 8px; border-radius: 12px;">${escapeHtml(result.sourceLabel)}</span>`;
        const avatarBg = isOS ? '#fdeceb' : '#e1f5fe';
        const avatarColor = isOS ? '#c62828' : '#0288d1';
        const isActive = result.key === activeKey;
        const activeStyle = isActive ? 'border-color: var(--color-accent-strong); background: #f0f8ff; box-shadow: 0 4px 12px rgba(0,0,0,0.05);' : '';

        return `
          <div class="qm-premium-preview-card qm-directory-item ${isActive ? 'active' : ''}" data-key="${escapeHtml(result.key)}" style="cursor: pointer; width: 100%; transition: all 0.2s; padding: 12px 16px; ${activeStyle}">
            <div class="qm-preview-avatar" style="background: ${avatarBg}; color: ${avatarColor}; flex-shrink: 0;">${escapeHtml(initials)}</div>
            <div style="display: flex; flex-direction: column; flex: 1; font-size: 0.85rem; line-height: 1.5; min-width: 0;">
              <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 6px;">
                <span class="qm-preview-name" style="font-size: 1rem; color: var(--color-text); line-height: 1.2;">${escapeHtml(result.nama || 'Nama belum terdeteksi')}</span>
                ${sourceBadge}
              </div>
              <div style="color: var(--color-text-muted); display: grid; grid-template-columns: 50px 1fr; gap: 2px;">
                <span>NRP</span><strong style="color: var(--color-text);">${escapeHtml(result.nrp || '-')}</strong>
                <span>Divisi</span><strong style="color: var(--color-text);">${escapeHtml(result.bagian || '-')}</strong>
                <span>Bagian</span><strong style="color: var(--color-text);">${escapeHtml(result.seksi || '-')}</strong>
                <span>Grup</span><strong style="color: var(--color-text);">${escapeHtml(result.group || '-')}</strong>
              </div>
            </div>
          </div>
        `;
      }

      let directoryHtml = '';

      // Render Internal Group
      if (internalResults.length > 0) {
        const isExpanded = state.karyawanGroupsExpanded.internal;
        const chevron = isExpanded ? '▼' : '▶';
        directoryHtml += `
          <div class="qm-preview-cards qm-group-header-card" data-group="internal">
            <span>Internal (${internalResults.length})</span>
            <span style="font-size: 0.8rem; color: var(--color-text-muted);">${chevron}</span>
          </div>
          <div class="qm-group-items-container" style="display: ${isExpanded ? 'flex' : 'none'}; flex-direction: column; gap: 8px; margin-bottom: 16px; width: 100%;">
            ${internalResults.map(result => renderEmployeeCardHtml(result)).join('')}
          </div>
        `;
      }

      // Render Outsource Group
      if (outsourceResults.length > 0) {
        const isExpanded = state.karyawanGroupsExpanded.outsource;
        const chevron = isExpanded ? '▼' : '▶';
        directoryHtml += `
          <div class="qm-preview-cards qm-group-header-card" data-group="outsource">
            <span>Outsource (${outsourceResults.length})</span>
            <span style="font-size: 0.8rem; color: var(--color-text-muted);">${chevron}</span>
          </div>
          <div class="qm-group-items-container" style="display: ${isExpanded ? 'flex' : 'none'}; flex-direction: column; gap: 8px; margin-bottom: 16px; width: 100%;">
            ${outsourceResults.map(result => renderEmployeeCardHtml(result)).join('')}
          </div>
        `;
      }

      uiAdapter.html('karyawanResults', directoryHtml);
    }

    // A. Render Selected Live Preview Data in Left Panel
    const leftPanel = document.getElementById('qm-karyawan-left-panel');
    if (leftPanel) {
      if (activeResult) {
        leftPanel.innerHTML = `
          <!-- Mode Toggle Tabs -->
          <div style="display: flex; gap: 24px; border-bottom: 1px solid var(--color-border); margin-bottom: 16px; padding: 0 4px;">
            <button type="button" class="qm-tab-btn ${activeMode === 'detail' ? 'active' : ''} qm-karyawan-detail-btn" data-key="${escapeHtml(activeResult.key)}" style="background: none; border: none; font-size: 0.95rem; font-weight: 600; cursor: pointer; padding: 8px 4px; border-bottom: 2px solid ${activeMode === 'detail' ? 'var(--color-text)' : 'transparent'}; color: ${activeMode === 'detail' ? 'var(--color-text)' : 'var(--color-text-muted)'}; transition: all 0.2s;">Detail Profil</button>
            <button type="button" class="qm-tab-btn ${activeMode === 'edit' ? 'active' : ''} qm-karyawan-edit-btn" data-key="${escapeHtml(activeResult.key)}" style="background: none; border: none; font-size: 0.95rem; font-weight: 600; cursor: pointer; padding: 8px 4px; border-bottom: 2px solid ${activeMode === 'edit' ? 'var(--color-text)' : 'transparent'}; color: ${activeMode === 'edit' ? 'var(--color-text)' : 'var(--color-text-muted)'}; transition: all 0.2s;">Edit HRIS</button>
          </div>

          <!-- Detail/Editor Expanded Area -->
          <div id="qm-karyawan-expanded-panel-container" style="width: 100%;">
            ${renderKaryawanExpandedPanel(activeResult)}
          </div>
        `;
      } else {
        leftPanel.innerHTML = `
          <div style="font-size: 0.85rem; color: var(--color-text-muted); line-height: 1.5; padding: 16px; background: #faf8f5; border-radius: 8px; border: 1px dashed var(--color-border); text-align: center;">
            Klik salah satu karyawan di hasil pencarian untuk melihat Detail Profil atau Edit HRIS.
          </div>
        `;
      }
    }
  }

  async function loadKaryawanDetail(key, nrp, profileUrl) {
    state.karyawanActivePanel = { key, mode: 'detail' };
    resetKaryawanEditor();
    state.karyawanDetail = {
      ...createEmptyKaryawanDetail(),
      key,
      nrp,
      loading: true
    };

    const isOS = key.startsWith('outsource:');
    sessionStorage.setItem('qm_is_os_' + nrp, isOS ? 'true' : 'false');

    renderKaryawanResults(true);

    try {
      if (!profileUrl) throw new Error('Link profile tidak tersedia.');
      const html = await hrisFetch(profileUrl);
      const profile = parseEmployeeProfile(parseHTML(html));
      state.karyawanDetail = {
        ...createEmptyKaryawanDetail(),
        key,
        nrp,
        profile
      };
    } catch (e) {
      state.karyawanDetail = {
        ...createEmptyKaryawanDetail(),
        key,
        nrp,
        error: e.message || 'Gagal memuat data profil.'
      };
    }

    renderKaryawanResults(true);
  }

  async function loadKaryawanEditor(key, nrp, knownId = null, knownIsOS = null) {
    state.karyawanActivePanel = { key, mode: 'edit' };
    resetKaryawanDetail();
    state.karyawanEditor = {
      ...createEmptyKaryawanEditor(),
      key,
      nrp,
      loading: true
    };

    const isOS = knownIsOS !== null ? knownIsOS : key.startsWith('outsource:');
    sessionStorage.setItem('qm_is_os_' + nrp, isOS ? 'true' : 'false');

    renderKaryawanResults(true);

    try {
      const emp = await fetchEmployee(nrp, knownId, isOS);
      if (!emp.found) throw new Error('Data karyawan tidak ditemukan.');

      const [jkResult, kkResult] = await Promise.allSettled([
        fetchJkOptions(nrp),
        fetchKKOptions(nrp)
      ]);

      const notices = [];
      const jkOptions = jkResult.status === 'fulfilled' ? jkResult.value : [];
      const kkOptions = kkResult.status === 'fulfilled' ? kkResult.value : [];

      if (jkResult.status === 'rejected' && kkResult.status === 'rejected') {
        clearEmployeeCache(nrp);
        notices.push("Opsi edit tidak tersedia (Mungkin karyawan outsource atau Anda tidak memiliki akses edit).");
      } else {
        if (jkResult.status === 'rejected') notices.push("Gagal memuat opsi Jam Kerja.");
        if (kkResult.status === 'rejected') notices.push("Gagal memuat opsi Kalender Kerja.");
      }

      updateKaryawanResultCache(key, emp);
      state.karyawanEditor = {
        ...createEmptyKaryawanEditor(),
        key,
        nrp,
        emp,
        jkOptions,
        kkOptions,
        notice: notices.join(' ')
      };
    } catch (e) {
      clearEmployeeCache(nrp);
      state.karyawanEditor = {
        ...createEmptyKaryawanEditor(),
        key,
        nrp,
        error: e.message || 'Gagal memuat data edit.'
      };
    }

    renderKaryawanResults(true);
  }

  function toggleKaryawanDetail(key) {
    const result = findKaryawanResult(key);
    if (!result || !result.nrp) {
      UI.showResult('warning', 'NRP Tidak Ditemukan', 'Data NRP tidak tersedia untuk melihat detail.');
      return;
    }

    loadKaryawanDetail(key, result.nrp, result.profileUrl);
  }

  function toggleKaryawanEditor(key) {
    const result = findKaryawanResult(key);
    if (!result || !result.nrp) {
      UI.showResult('warning', 'NRP Tidak Ditemukan', 'Data NRP tidak tersedia untuk quick edit.');
      return;
    }

    loadKaryawanEditor(key, result.nrp, result.id, result.source === 'outsource');
  }

  async function refreshKaryawanEditorAfterSave(key, nrp) {
    clearEmployeeCache(nrp);
    const result = findKaryawanResult(key);
    if (result) {
      await loadKaryawanEditor(key, nrp, result.id, result.source === 'outsource');
    } else {
      await loadKaryawanEditor(key, nrp);
    }
  }

  async function handleKaryawanSaveEdit(input) {
    const payload = (input && input.target && input.currentTarget) ? null : input;
    const key = payload?.key || this?.dataset?.key || '';
    const editor = state.karyawanEditor;
    if (!editor || editor.key !== key) return;

    const selectValues = payload || panelReaders.karyawanSave(key);
    const nextJk = selectValues.nextJk || '';
    const nextKk = selectValues.nextKk || '';

    if (!nextJk && !nextKk) {
      UI.showResult('warning', 'Data Belum Lengkap', 'Pilih JK atau KK yang ingin disimpan.');
      return;
    }

    const currentJk = String(editor.emp?.jk || '').trim();
    const currentKk = String(editor.emp?.KK || '').trim();
    const jkChanged = !!nextJk && nextJk !== currentJk;
    const kkChanged = !!nextKk && nextKk !== currentKk;

    if (!jkChanged && !kkChanged) {
      UI.showResult('warning', 'Tidak Ada Perubahan', 'JK dan KK masih sama dengan data saat ini.');
      return;
    }

    state.karyawanEditor = { ...editor, saving: true };
    renderKaryawanResults(true);

    try {
      if (jkChanged) await saveJkMaster(editor.nrp, nextJk);
      if (kkChanged) await saveKKMaster(editor.nrp, nextKk);
      await refreshKaryawanEditorAfterSave(key, editor.nrp);
      const parts = [];
      if (jkChanged) parts.push('JK');
      if (kkChanged) parts.push('KK');
      UI.showResult('success', 'Data Diperbarui', `${parts.join(' & ')} NRP ${editor.nrp} berhasil diperbarui.`);
    } catch (e) {
      state.karyawanEditor = { ...editor, saving: false };
      renderKaryawanResults(true);
      UI.showResult('danger', 'Gagal Menyimpan', e.message || 'Perubahan JK/KK gagal disimpan.');
    }
  }


  /** Fetch attendance table and return anomalies array. */
  async function fetchAttendance(nrp, bulan, tahun, bagian, seksi) {
    const prof = startProfile('fetchAttendance', { nrp, bulan, tahun });
    try {
      const html = await hrisFetch(attendanceUrl(bulan, tahun, nrp));
      const doc = parseHTML(html);
      return scanAttendance(doc, { tahun, bulan, nrp, bagian, seksi });
    } finally {
      finishProfile(prof, { nrp });
    }
  }

  /* ============================================================
   * 6. STATE & LOGGER
   * ============================================================ */

  const state = {
    isOpen: false,
    loading: false,
    history: [],
    maxHistory: 8,
    anomalyRunId: 0,
    anomalies: [],
    pendingChecks: 0,
    shortcut: GM_getValue('qm_shortcut', 'Ctrl+Shift+Q'),
    alwaysCollapse: GM_getValue('qm_always_collapse', false),
    theme: GM_getValue('qm_theme', 'light'),
    debug: GM_getValue('qm_debug', false),
    batchQueue: [],
    batchResults: [],
    batchLogs: [],
    batchBulan: 0,
    batchTahun: 0,
    batchRunId: 0,
    batchActiveWorkers: 0,
    batchTotal: 0,
    batchAborted: false,
    batchProfile: { renderTotal: 0, itemDurations: [] },
    profileStats: {},
    profileFlags: {},
    expandedAnomalyGroups: new Set(),
    activeAbortController: null,
    activeCancelableFlow: false,
    cancelRequested: false,
    backgroundHeartbeatTimer: null,
    karyawanQuery: '',
    karyawanResults: [],
    karyawanGroupsExpanded: { internal: false, outsource: false },
    karyawanLoading: false,
    karyawanError: '',
    refreshRunId: 0,
    karyawanActivePanel: { key: '', mode: '' },
    karyawanEditor: null,
    karyawanDetail: null,
    attendanceCheck: createEmptyAttendanceCheck(),
    spklCheck: createEmptySpklCheck(),
    spklEditCurrentIndex: -1,
    panelPos: JSON.parse(GM_getValue('qm_panel_pos', 'null')), // {top, left}
  };
  let shortcutKey = GM_getValue('qm_shortcut', 'Ctrl+Q');
  let alwaysCollapseMenu = GM_getValue('qm_always_collapse', false);
  let isRecordingShortcut = false;
  let cachedEditHtml = null; // Cache for editgeneral HTML to speed up save

  /**
   * Centralized Logging System
   */
  const Logger = {
    _formatTime: () => {
      const now = new Date();
      return `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}:${now.getSeconds().toString().padStart(2, '0')}`;
    },

    _logToUI: (level, msg) => {
      if (typeof pushLog === 'function') {
        pushLog(msg, level);
      }
    },

    debug: (msg, data) => {
      if (!state.debug) return;
      const time = Logger._formatTime();
      const prefix = `%c[${time}] [HRIS-QM] [DEBUG]`;
      const style = 'color: #888; font-weight: normal;';
      if (data !== undefined) console.log(prefix, style, msg, data);
      else console.log(prefix, style, msg);
      Logger._logToUI('debug', msg);
    },

    info: (msg, data) => {
      Logger._logToUI('info', msg);
      if (!state.debug) return;
      const time = Logger._formatTime();
      const prefix = `%c[${time}] [HRIS-QM] [INFO]`;
      const style = 'color: #00bcd4; font-weight: bold;';
      if (data !== undefined) console.log(prefix, style, msg, data);
      else console.log(prefix, style, msg);
    },

    success: (msg, data) => {
      Logger._logToUI('success', msg);
      if (!state.debug) return;
      const time = Logger._formatTime();
      const prefix = `%c[${time}] [HRIS-QM] [SUCCESS]`;
      const style = 'color: #4caf50; font-weight: bold;';
      if (data !== undefined) console.log(prefix, style, msg, data);
      else console.log(prefix, style, msg);
    },

    warn: (msg, data) => {
      const time = Logger._formatTime();
      const prefix = `%c[${time}] [HRIS-QM] [WARN]`;
      const style = 'color: #ff9800; font-weight: bold;';
      if (data !== undefined) console.warn(prefix, style, msg, data);
      else console.warn(prefix, style, msg);
      Logger._logToUI('warn', msg);
    },

    error: (msg, data, err) => {
      const time = Logger._formatTime();
      const prefix = `%c[${time}] [HRIS-QM] [ERROR]`;
      const style = 'color: #f44336; font-weight: bold;';
      if (err) console.error(prefix, style, msg, data, err);
      else if (data !== undefined) console.error(prefix, style, msg, data);
      else console.error(prefix, style, msg);

      let uiMsg = msg;
      if (data !== undefined) {
        let dataStr = '';
        if (data instanceof Error) dataStr = data.message;
        else if (typeof data === 'object') {
          try { dataStr = JSON.stringify(data); } catch (e) { dataStr = '[Object]'; }
        } else dataStr = data;
        uiMsg += ': ' + dataStr;
      }
      Logger._logToUI('danger', uiMsg);
    }
  };

  function startProfile(label, meta = {}) {
    if (!state.debug) return null;
    return { label, meta, startedAt: perfNow() };
  }


  function finishProfile(token, meta = {}) {
    if (!token) return 0;

    const duration = perfNow() - token.startedAt;
    if (!state.profileStats) state.profileStats = {};
    if (!state.profileFlags) state.profileFlags = {};

    const stats = state.profileStats[token.label] || {
      samples: [],
      count: 0,
      total: 0,
      max: 0
    };

    stats.count += 1;
    stats.total += duration;
    stats.max = Math.max(stats.max, duration);
    stats.samples.push(duration);
    if (stats.samples.length > PROFILE_CONFIG.SAMPLE_HISTORY_LIMIT) stats.samples.shift();
    state.profileStats[token.label] = stats;

    const recent = stats.samples.slice(-PROFILE_CONFIG.MEDIAN_SAMPLE_SIZE);
    const medianVal = recent.length >= PROFILE_CONFIG.MEDIAN_SAMPLE_SIZE ? median(recent) : null;
    const payload = {
      ms: Number(duration.toFixed(2)),
      avgMs: Number((stats.total / stats.count).toFixed(2)),
      maxMs: Number(stats.max.toFixed(2)),
      medianMs: medianVal === null ? null : Number(medianVal.toFixed(2)),
      samples: stats.count,
      ...token.meta,
      ...meta
    };

    Logger.debug(`[PROFILE] ${token.label}`, payload);

    if (medianVal !== null && medianVal >= PROFILE_CONFIG.HOT_SYNC_MS) {
      const warnKey = `${token.label}:sync`;
      if (!state.profileFlags[warnKey]) {
        state.profileFlags[warnKey] = true;
        Logger.warn(`Profiling: ${token.label} median ${medianVal.toFixed(2)}ms melewati ambang ${PROFILE_CONFIG.HOT_SYNC_MS}ms.`, payload);
      }
    }

    return duration;
  }


  function resetBatchProfile() {
    state.batchProfile = { renderTotal: 0, itemDurations: [] };
  }


  /* ============================================================
   * 7. ANOMALY DETECTION
   * ============================================================ */

  function hasShiftMark(td) {
    if (!td) return false;

    // 1. Check if there is an input element (like a checkbox)
    const input = td.querySelector('input');
    if (input) {
      return input.checked || td.querySelector('input:checked') !== null;
    }

    // 2. Check text content for checkmark symbols/text
    const text = td.textContent.trim().toLowerCase();
    if (text !== '') {
      if (text.includes('✓') || text.includes('☑') || text.includes('ok') || text === 'v' || text === 'check' || text === 'checked') {
        return true;
      }
    }

    // 3. Check HTML content for FontAwesome checkmark classes (avoid raw 'check' which matches 'checkbox')
    const html = td.innerHTML.toLowerCase();
    if (html.includes('fa-check') || html.includes('fa-square-check') || html.includes('fa-check-square') || html.includes('fa-check-circle')) {
      return true;
    }

    return false;
  }

  function guessActualShift(waktuMsk, rules) {
    if (waktuMsk === null) return '1';
    const jamMentah = waktuMsk >= 24.0 ? waktuMsk - 24.0 : waktuMsk;
    if (jamMentah >= rules.shift1.jamTebakMulai && jamMentah <= rules.shift1.jamTebakAkhir) return '1';
    if (jamMentah > rules.shift2.jamTebakMulai && jamMentah <= rules.shift2.jamTebakAkhir) return '2';
    return '3';
  }

  /** Count rows marked as libur by CSS classes/colors. */
  function countHolidays(docContext) {
    const root = docContext || document;
    const trs = Array.from(root.querySelectorAll('table tbody tr')).filter(tr => {
      return !tr.closest('#qm-panel') && !tr.closest('.command-menu');
    });
    let total = 0;

    trs.forEach(tr => {
      const tds = tr.querySelectorAll('td');
      if (tds.length < 12 || tr.textContent.toLowerCase().includes('total')) return;
      const { isLiburColor } = getRowFlags(tr);
      if (isLiburColor) total++;
    });
    return total;
  }


  /** Unified row status detector (Libur, HalfDay, Normal). */
  function getAttendanceRowStatus(tr) {
    const tds = tr.querySelectorAll('td');
    if (tds.length === 0) return { isLibur: false, isHalfDay: false, ketText: '' };

    const { isLiburColor, isHalfDayColor } = getRowFlags(tr);

    let ketText = '';
    if (tds.length > COL.KET) ketText = tds[COL.KET].textContent.trim().toUpperCase();

    const isLibur = isLiburColor || ['L', 'LB', 'LH'].includes(ketText);
    const isHalfDay = !isLibur && (isHalfDayColor || ['S', 'CH', 'HD'].includes(ketText));

    return { isLibur, isHalfDay, ketText };
  }

  /** Push an anomaly record. Pure — no DOM side effects. */
  function flagCell(anomalies, tglText, colIndex, title, customLink, cekSpklCells, fullDate, shiftVal) {
    anomalies.push({ tgl: tglText, fullDate, colIndex, msg: title, link: customLink || '', shift: shiftVal });
    if (title.includes('Cek SPKL') && cekSpklCells) {
      if (!cekSpklCells.some(c => c.tgl === tglText && c.colIndex === colIndex)) {
        cekSpklCells.push({ tgl: tglText, fullDate, colIndex, shift: shiftVal });
      }
    }
  }


  function pushAnomaly(tgl, col, msg, link = '') {
    state.anomalies.push({ tgl, col, msg, link });
  }

  /** Apply visual anomaly marks to the live DOM table. Called only on the attendance page. */
  function applyMark(docContext, anomalies) {
    const prof = startProfile('applyMark', { count: anomalies.length });
    const lookup = {};
    anomalies.forEach(a => {
      const key = a.tgl;
      if (!lookup[key]) lookup[key] = [];
      lookup[key].push(a);
    });

    const root = docContext || document;
    const trs = Array.from(root.querySelectorAll('table tbody tr')).filter(tr => {
      return !tr.closest('#qm-panel') && !tr.closest('.command-menu');
    });

    trs.forEach(tr => {
      const tds = tr.querySelectorAll('td');
      if (tds.length < 12 || tr.textContent.toLowerCase().includes('total')) return;
      const tgl = tds[COL.TGL].textContent.trim();
      const rowAnomalies = lookup[tgl];
      if (!rowAnomalies) return;

      rowAnomalies.forEach(a => {
        const colIdx = a.colIndex;
        if (colIdx === undefined) return;
        const td = tds[colIdx];
        if (!td) return;

        const existingTitle = td.getAttribute('title') || '';
        const newTitle = existingTitle && !existingTitle.includes(a.msg) ? existingTitle + ' | ' + a.msg : a.msg;
        td.classList.add('qm-anomaly-cell');
        td.setAttribute('title', newTitle);

        if (!td.querySelector('.qm-fix-dot')) {
          const linkStr = a.link ? `data-fix-link="${escapeHtml(a.link)}" data-fix-date="${escapeHtml(String(a.tgl))}" data-full-date="${escapeHtml(String(a.fullDate || ''))}"` : '';

          let titleStr = 'Buka Halaman Kehadiran';
          if (a.msg === 'Buka Halaman Kehadiran') titleStr = 'Buka Halaman Kehadiran';
          else if (a.msg && (a.msg.includes('Duplikasi') || a.msg.includes('Double entry'))) titleStr = 'Lihat Duplikasi Shift';
          else if (a.msg && a.msg.includes('Pulang awal')) titleStr = 'Cek Kehadiran (Pulang Awal)';
          else if (a.msg === 'Cek Distribusi' || (a.msg && a.msg.startsWith('Shift Kosong')) || a.msg === 'Shift muncul') titleStr = 'Perbaiki Distribusi Jam Kerja';
          else if (a.msg === 'Cek SPKL') titleStr = 'Cek SPKL Online';
          else if (!a.link) titleStr = 'Perbaiki Anomali (Segera Hadir)';

          td.insertAdjacentHTML('beforeend', `<button class="qm-fix-dot" title="${titleStr}" ${linkStr}></button>`);
        } else if (a.link) {
          const btn = td.querySelector('.qm-fix-dot');
          btn.setAttribute('data-fix-link', a.link);
          btn.setAttribute('data-fix-date', String(a.tgl));
          if (a.fullDate) btn.setAttribute('data-full-date', String(a.fullDate));

          let titleStr = 'Buka Halaman Kehadiran';
          if (a.msg === 'Buka Halaman Kehadiran') titleStr = 'Buka Halaman Kehadiran';
          else if (a.msg && (a.msg.includes('Duplikasi') || a.msg.includes('Double entry'))) titleStr = 'Lihat Duplikasi Shift';
          else if (a.msg && a.msg.includes('Pulang awal')) titleStr = 'Cek Kehadiran';
          else if (a.msg === 'Cek Distribusi' || (a.msg && a.msg.startsWith('Shift Kosong')) || a.msg === 'Shift muncul') titleStr = 'Perbaiki Distribusi Jam Kerja';
          else if (a.msg === 'Cek SPKL') titleStr = 'Cek SPKL Online';

          btn.setAttribute('title', titleStr);
        }
      });
    });
    finishProfile(prof, { count: anomalies.length });
  }

  function validateShiftRow(tds, tglText, mskText, klrText, rules, ctx, cekSpklCells, anomalies, isLibur, isHalfDay, fullDate) {
    const ketText = tds[COL.KET].textContent.trim().toUpperCase();
    let shift1 = hasShiftMark(tds[COL.SHIFT1]);
    let shift2 = hasShiftMark(tds[COL.SHIFT2]);
    let shift3 = hasShiftMark(tds[COL.SHIFT3]);

    // No shift checked but has clock data or is Mangkir → guess shift
    if (!shift1 && !shift2 && !shift3 && (mskText || klrText || ketText === 'A')) {
      const guessed = guessActualShift(parseTimeToDecimal(mskText), rules);
      const link = buildDistribusiLink(ctx, tglText, guessed);
      const msg = 'Shift Kosong | Keterangan Kosong';
      flagCell(anomalies, tglText, COL.SHIFT1, msg, link, cekSpklCells, fullDate);
      flagCell(anomalies, tglText, COL.SHIFT2, msg, link, cekSpklCells, fullDate);
      flagCell(anomalies, tglText, COL.SHIFT3, msg, link, cekSpklCells, fullDate);
    }

    // Derive active shift from clock-in if still unknown
    const mskTime = parseTimeToDecimal(mskText);
    if (!shift1 && !shift2 && !shift3) {
      if (mskTime !== null) {
        if (mskTime >= rules.shift1.jamTebakMulai && mskTime <= rules.shift1.jamTebakAkhir) shift1 = true;
        else if (mskTime > rules.shift2.jamTebakMulai && mskTime <= rules.shift2.jamTebakAkhir) shift2 = true;
        else shift3 = true;
      }
    }

    const activeShift = shift1 ? '1' : (shift2 ? '2' : (shift3 ? '3' : null));
    return { shift1, shift2, shift3, activeShift, mskTime };
  }


  /** Determine half-day Pulang Awal threshold based on actual clock-in time. */
  function halfDayLeaveThreshold(shift1, shift2, shift3, mskTime) {
    if (shift1) return HALFDAY_RULES.shift1.batasPulangAwal;
    if (shift3) return HALFDAY_RULES.shift3.batasPulangAwal;
    if (shift2) {
      if (mskTime !== null && mskTime >= HALFDAY_RULES.shift2.altJamMulai) {
        return HALFDAY_RULES.shift2.altBatasPulangAwal;
      }
      return HALFDAY_RULES.shift2.batasPulangAwal;
    }
    return null;
  }

  function validateOvertime(tds, tglText, mskTime, klrTime, shift1, shift2, shift3, isLibur, isHalfDay, rules, ctx, cekSpklCells, anomalies, is5HariKerja, fullDate) {
    let mskLembur = false, klrLembur = false;
    let adjMsk = mskTime, adjKlr = klrTime;

    if (shift2) { if (adjKlr !== null && adjKlr <= THRESHOLDS.ADJ_KLR_SHIFT2_BATAS) adjKlr += 24.0; }
    else if (shift3) {
      if (adjMsk !== null && adjMsk <= THRESHOLDS.ADJ_MSK_SHIFT3_BATAS) adjMsk += 24.0;
      if (adjKlr !== null && adjKlr <= THRESHOLDS.ADJ_KLR_SHIFT3_BATAS) adjKlr += 24.0;
    }

    if (adjMsk === null && adjKlr === null) return { mskLembur, klrLembur };

    if (isLibur) {
      if (adjMsk !== null) mskLembur = true;
      if (adjKlr !== null) klrLembur = true;
    }

    // Pulang Awal threshold: half-day uses shorter threshold
    const paThreshold = isHalfDay
      ? halfDayLeaveThreshold(shift1, shift2, shift3, mskTime)
      : null;

    // Dynamic Shift 1 Models
    let s1MasukLembur = rules.shift1.batasMasukLembur;
    let s1KeluarLembur = rules.shift1.batasKeluarLembur;
    let s1Terlambat = rules.shift1.batasTerlambatMasuk;
    let s1PulangAwal = rules.shift1.batasPulangAwal;

    if (is5HariKerja) {
      // Differentiate between 06.00-15.00 and 07.30-16.30 based on clock-in time
      if ((adjMsk !== null && adjMsk < 6.5) || (adjMsk === null && adjKlr !== null && adjKlr < 16.0)) {
        s1KeluarLembur = 15.5;
        s1Terlambat = 6.25;
        s1PulangAwal = 15.0;

        // Support 06.00 - 14.00 model (e.g., NRP 2869)
        if (adjMsk !== null && adjMsk <= 6.0 && ctx?.nrp === '2869') s1PulangAwal = 14.0;
      } else {
        s1KeluarLembur = THRESHOLDS.SHIFT1_KLR_LEMBUR_5HR;
        s1Terlambat = THRESHOLDS.SHIFT1_TERLAMBAT_5HR;
        s1PulangAwal = THRESHOLDS.SHIFT1_PULANG_AWAL_5HR;
      }
    }

    if (shift1) {
      if (adjMsk !== null && adjMsk < s1MasukLembur) mskLembur = true;
      if (adjKlr !== null && adjKlr > s1KeluarLembur) klrLembur = true;
      if (adjMsk !== null && adjMsk > s1Terlambat) {
        if (adjMsk > THRESHOLDS.SHIFT1_MSK_UPPER_BATAS) flagCell(anomalies, tglText, COL.SHIFT1, 'Cek Distribusi', buildDistribusiLink(ctx, tglText, guessActualShift(adjMsk, rules)), cekSpklCells, fullDate);
        else if (adjMsk !== adjKlr) flagCell(anomalies, tglText, COL.MSK, 'Terlambat Shift I', '', cekSpklCells, fullDate);
      }
      const paShift1 = isHalfDay && paThreshold !== null ? paThreshold : s1PulangAwal;
      if (adjKlr !== null && adjKlr < paShift1 && adjMsk !== adjKlr) flagCell(anomalies, tglText, COL.KLR, 'Pulang awal Shift I', buildKehadiranLink(ctx), cekSpklCells, fullDate, '1');
    } else if (shift2) {
      if (adjMsk !== null && adjMsk < rules.shift2.batasMasukLembur) mskLembur = true;
      if (adjKlr !== null && adjKlr > rules.shift2.batasKeluarLembur) klrLembur = true;
      if (adjMsk !== null) {
        if (adjMsk < THRESHOLDS.SHIFT2_MSK_LOWER_BATAS || adjMsk > THRESHOLDS.SHIFT2_MSK_UPPER_BATAS) flagCell(anomalies, tglText, COL.SHIFT2, 'Cek Distribusi', buildDistribusiLink(ctx, tglText, guessActualShift(adjMsk, rules)), cekSpklCells, fullDate, '2');
        else if (adjMsk > rules.shift2.batasTerlambatMasuk && adjMsk !== adjKlr) flagCell(anomalies, tglText, COL.MSK, 'Terlambat Shift II', '', cekSpklCells, fullDate, '2');
      }
      const paShift2 = isHalfDay && paThreshold !== null ? paThreshold : rules.shift2.batasPulangAwal;
      if (adjKlr !== null && adjKlr < paShift2 && adjMsk !== adjKlr) flagCell(anomalies, tglText, COL.KLR, 'Pulang awal Shift II', buildKehadiranLink(ctx), cekSpklCells, fullDate, '2');
    } else if (shift3) {
      if (adjMsk !== null && adjMsk > rules.shift3.batasMasukLemburAwal && adjMsk < rules.shift3.batasMasukLemburAkhir) mskLembur = true;
      if (adjKlr !== null && adjKlr > rules.shift3.batasKeluarLembur) klrLembur = true;
      if (adjMsk !== null) {
        if (adjMsk < rules.shift3.batasAwalMasuk || adjMsk > THRESHOLDS.SHIFT3_MSK_UPPER_BATAS) flagCell(anomalies, tglText, COL.SHIFT3, 'Cek Distribusi', buildDistribusiLink(ctx, tglText, guessActualShift(adjMsk, rules)), cekSpklCells, fullDate, '3');
        else if (adjMsk > rules.shift3.batasTerlambatMasuk && adjMsk !== adjKlr) flagCell(anomalies, tglText, COL.MSK, 'Terlambat Shift III', '', cekSpklCells, fullDate, '3');
      }
      const paShift3 = isHalfDay && paThreshold !== null ? paThreshold : rules.shift3.batasPulangAwal;
      if (adjKlr !== null && adjKlr < paShift3 && adjMsk !== adjKlr) flagCell(anomalies, tglText, COL.KLR, 'Pulang awal Shift III', buildKehadiranLink(ctx), cekSpklCells, fullDate, '3');
    }
    return { mskLembur, klrLembur };
  }

  function scanAttendance(doc, ctx) {
    const prof = startProfile('scanAttendance', { nrp: ctx?.nrp, bulan: ctx?.bulan, tahun: ctx?.tahun });
    const anomalies = [];
    const absentDates = [];
    const cekSpklCells = [];

    const rekaps = {
      hariKerja: 0,
      otb: 0,
      otl: 0,
      ota: 0,
      otp: 0,
      keterangan: {}
    };

    const totalLibur = countHolidays(doc);
    const is5HariKerja = totalLibur >= THRESHOLDS.MIN_LIBUR_5_HARI_KERJA;
    const rules = structuredClone(SHIFT_RULES);
    if (is5HariKerja) {
      rules.shift2.batasKeluarLembur = THRESHOLDS.SHIFT2_KLR_LEMBUR_5HR;
    }

    const root = doc || document;
    const trs = Array.from(root.querySelectorAll('table tbody tr')).filter(tr => {
      return !tr.closest('#qm-panel') && !tr.closest('.command-menu');
    });

    // Pre-calculate shift counts per date and shift number to detect duplicates
    const dateShiftCounts = {}; // Key: `${tgl}_${shiftNum}`
    trs.forEach(tr => {
      const tds = tr.querySelectorAll('td');
      if (tds.length < 12 || tr.textContent.toLowerCase().includes('total')) return;
      const tgl = tds[COL.TGL].textContent.trim();
      if (!tgl) return;
      if (hasShiftMark(tds[COL.SHIFT1])) {
        const key = `${tgl}_1`;
        dateShiftCounts[key] = (dateShiftCounts[key] || 0) + 1;
      }
      if (hasShiftMark(tds[COL.SHIFT2])) {
        const key = `${tgl}_2`;
        dateShiftCounts[key] = (dateShiftCounts[key] || 0) + 1;
      }
      if (hasShiftMark(tds[COL.SHIFT3])) {
        const key = `${tgl}_3`;
        dateShiftCounts[key] = (dateShiftCounts[key] || 0) + 1;
      }
    });

    trs.forEach(tr => {
      const tds = tr.querySelectorAll('td');
      if (tds.length < 12 || tr.textContent.toLowerCase().includes('total')) return;

      const tglText = tds[COL.TGL].textContent.trim();
      const fullDate = `${ctx.tahun}-${String(ctx.bulan).padStart(2, '0')}-${String(tglText).padStart(2, '0')}`;

      const mskText = tds[COL.MSK].textContent.trim();
      const klrText = tds[COL.KLR].textContent.trim();
      const { isLibur, isHalfDay, ketText } = getAttendanceRowStatus(tr);

      const shiftInfo = validateShiftRow(tds, tglText, mskText, klrText, rules, ctx, cekSpklCells, anomalies, isLibur, isHalfDay, fullDate);
      const { shift1, shift2, shift3, activeShift, mskTime } = shiftInfo;
      const klrTime = parseTimeToDecimal(klrText);

      // Detect multiple checked shifts inside the same row
      let checkedShiftsCount = 0;
      if (shift1) checkedShiftsCount++;
      if (shift2) checkedShiftsCount++;
      if (shift3) checkedShiftsCount++;
      if (checkedShiftsCount > 1) {
        const msg = 'Multi-shift Checked pada baris yang sama';
        const link = buildKehadiranLink(ctx);
        flagCell(anomalies, tglText, COL.SHIFT1, msg, link, cekSpklCells, fullDate);
        flagCell(anomalies, tglText, COL.SHIFT2, msg, link, cekSpklCells, fullDate);
        flagCell(anomalies, tglText, COL.SHIFT3, msg, link, cekSpklCells, fullDate);
      }

      // Detect multiple rows for same date with checked shifts (Barcode overlap/error)
      const isSaturday = new Date(ctx.tahun, ctx.bulan - 1, parseInt(tglText)).getDay() === 6;

      let hasDuplicateActiveShift = false;
      if (activeShift) {
        const key = `${tglText}_${activeShift}`;
        if (dateShiftCounts[key] > 1) {
          hasDuplicateActiveShift = true;
        }
      }

      if (hasDuplicateActiveShift && !isHalfDay && !isSaturday) {
        const barcodeLink = buildKehadiranLink(ctx);
        let msg = 'Duplikasi Shift pada tanggal yang sama';

        let adjMsk = mskTime;
        let adjKlr = klrTime;
        // Apply basic shift-based adjustments for accurate comparison
        if (shift2 && adjKlr !== null && adjKlr <= THRESHOLDS.ADJ_KLR_SHIFT2_BATAS) adjKlr += 24.0;
        else if (shift3) {
          if (adjMsk !== null && adjMsk <= THRESHOLDS.ADJ_MSK_SHIFT3_BATAS) adjMsk += 24.0;
          if (adjKlr !== null && adjKlr <= THRESHOLDS.ADJ_KLR_SHIFT3_BATAS) adjKlr += 24.0;
        }

        if (adjMsk !== null && adjKlr !== null && adjMsk >= adjKlr) {
          msg = 'Double entry / Error Barcode (MSK >= KLR)';
        }

        flagCell(anomalies, tglText, COL.TGL, msg, barcodeLink, cekSpklCells, fullDate);
        if (shift1) flagCell(anomalies, tglText, COL.SHIFT1, msg, barcodeLink, cekSpklCells, fullDate);
        if (shift2) flagCell(anomalies, tglText, COL.SHIFT2, msg, barcodeLink, cekSpklCells, fullDate);
        if (shift3) flagCell(anomalies, tglText, COL.SHIFT3, msg, barcodeLink, cekSpklCells, fullDate);
      }

      if (activeShift && mskTime !== null && !isLibur && !isHalfDay) {
        const guessed = guessActualShift(mskTime, rules);
        if (guessed !== activeShift) {
          const msg = 'Jam MSK tidak cocok dengan Shift ' + activeShift + ' (Terdeteksi Shift ' + guessed + ')';
          const link = buildDistribusiLink(ctx, tglText, guessed);
          flagCell(anomalies, tglText, COL.MSK, msg, link, cekSpklCells, fullDate);
          if (shift1) flagCell(anomalies, tglText, COL.SHIFT1, msg, '', cekSpklCells, fullDate);
          if (shift2) flagCell(anomalies, tglText, COL.SHIFT2, msg, '', cekSpklCells, fullDate);
          if (shift3) flagCell(anomalies, tglText, COL.SHIFT3, msg, '', cekSpklCells, fullDate);
        }
      }

      let isAbsent = ketText === 'A';
      for (let i = 11; i < tds.length; i++) {
        if (tds[i].textContent.trim() === 'A') isAbsent = true;
      }
      if (isAbsent) {
        absentDates.push({ date: tglText, tr: tr });
      }

      const barcodeLink = buildKehadiranLink(ctx);

      const isSunday = new Date(ctx.tahun, ctx.bulan - 1, parseInt(tglText)).getDay() === 0;
      const isWeekendRest = (isSaturday || isSunday) && !activeShift && !mskText && !klrText && !ketText;

      if (isWeekendRest) {
        // Weekend rest day: do not flag Buka Halaman Kehadiran
      } else if (!isLibur) {
        if (!mskText && !ketText) flagCell(anomalies, tglText, COL.MSK, 'Buka Halaman Kehadiran', barcodeLink, cekSpklCells, fullDate);
        if (!klrText && !ketText) flagCell(anomalies, tglText, COL.KLR, 'Buka Halaman Kehadiran', barcodeLink, cekSpklCells, fullDate);
      } else {
        if (mskText || klrText) {
          if (!mskText) flagCell(anomalies, tglText, COL.MSK, 'Buka Halaman Kehadiran', barcodeLink, cekSpklCells, fullDate);
          if (!klrText) flagCell(anomalies, tglText, COL.KLR, 'Buka Halaman Kehadiran', barcodeLink, cekSpklCells, fullDate);
        }
      }


      const ot = validateOvertime(tds, tglText, mskTime, klrTime, shift1, shift2, shift3, isLibur, isHalfDay, rules, ctx, cekSpklCells, anomalies, is5HariKerja, fullDate);

      const otbText = tds[COL.OTB].textContent.trim();
      const otlText = tds[COL.OTL].textContent.trim();
      const otpText = tds[COL.OTP].textContent.trim();

      const valOtb = parseFloat(otbText) || 0;
      const valOtl = parseFloat(otlText) || 0;
      const valOtp = parseFloat(otpText) || 0;

      const hkText = tds[COL.HARI_KERJA] ? tds[COL.HARI_KERJA].textContent.trim() : '0';
      rekaps.hariKerja += parseFloat(hkText) || 0;

      if (ketText && ketText !== '-') {
        rekaps.keterangan[ketText] = (rekaps.keterangan[ketText] || 0) + 1;
      }

      rekaps.otb += valOtb;
      rekaps.otl += valOtl;
      rekaps.otp += valOtp;
      rekaps.ota += (valOtb + valOtl);

      if (valOtb > THRESHOLDS.OT_BATAS_WAJAR) flagCell(anomalies, tglText, COL.OTB, 'Angka OTB tidak wajar', '', cekSpklCells, fullDate);
      if (valOtl > THRESHOLDS.OT_BATAS_WAJAR) flagCell(anomalies, tglText, COL.OTL, 'Angka OTL tidak wajar', '', cekSpklCells, fullDate);
      if (valOtp > THRESHOLDS.OT_BATAS_WAJAR) flagCell(anomalies, tglText, COL.OTP, 'Angka OTP tidak wajar', '', cekSpklCells, fullDate);

      const hasAnyOT = parseFloat(otbText) > 0 || parseFloat(otlText) > 0 || parseFloat(otpText) > 0;
      if ((ot.mskLembur || ot.klrLembur) && !hasAnyOT) {
        const d = String(tglText).padStart(2, '0');
        const spklUrl = ROUTES.SPKL_ONLINE(ctx.tahun, String(ctx.bulan).padStart(2, '0'), d, d, ctx.nrp);
        const sVal = activeShift || '';
        flagCell(anomalies, tglText, COL.OTB, 'Cek SPKL', spklUrl, cekSpklCells, fullDate, sVal);
        flagCell(anomalies, tglText, COL.OTL, 'Cek SPKL', spklUrl, cekSpklCells, fullDate, sVal);
        flagCell(anomalies, tglText, COL.OTP, 'Cek SPKL', spklUrl, cekSpklCells, fullDate, sVal);
        if (ot.mskLembur) flagCell(anomalies, tglText, COL.MSK, 'Cek SPKL', spklUrl, cekSpklCells, fullDate, sVal);
        if (ot.klrLembur) flagCell(anomalies, tglText, COL.KLR, 'Cek SPKL', spklUrl, cekSpklCells, fullDate, sVal);
      }
    });

    const result = { anomalies, absentDates, cekSpklCells, rekaps };
    finishProfile(prof, {
      nrp: ctx?.nrp,
      anomalyCount: anomalies.length,
      absentCount: absentDates.length,
      spklCheckCount: cekSpklCells.length
    });
    return result;
  }

  function detectAnomalies() {
    if (!isAttendancePagePath()) return;

    const prof = startProfile('detectAnomalies');
    const runId = ++state.anomalyRunId;

    try {
      state.anomalies = [];
      state.pendingChecks = 0;

      const anomaliTab = document.querySelector('[data-pane="anomali"]');
      if (anomaliTab) anomaliTab.classList.remove('qm-tab-loading');

      const ctx = getPageContext();
      const result = scanAttendance(document, ctx);
      state.anomalies = result.anomalies;

      applyMark(document, result.anomalies);
      renderAnomalies();

      if (result.absentDates.length > 0) { state.pendingChecks++; checkBarcodeMangkir(result.absentDates, runId); }
      if (result.cekSpklCells.length > 0) { state.pendingChecks++; checkSPKLOnline(result.cekSpklCells, runId); }
      if (state.pendingChecks > 0 && anomaliTab) anomaliTab.classList.add('qm-tab-loading');
    } finally {
      finishProfile(prof, { anomalyCount: state.anomalies.length });
    }
  }

  /* ============================================================
   * 8. BATCH PROCESSING
   * ============================================================ */

  function parseNrpList(text) {
    return text.split(/[\n,]+/).map(s => s.trim()).filter(s => /^\d{4}$|^\d{8}$/.test(s));
  }

  function startBatchAnomalyCheck() {
    const inputMulti = uiAdapter.get('#qm-input-multi-nrp');
    const inputBulan = uiAdapter.get('globalMonth');
    const inputTahun = uiAdapter.get('globalYear');
    const btnCheck = uiAdapter.get('batchCheckButton');
    const btnExport = uiAdapter.get('batchExportButton');
    const progress = uiAdapter.get('batchProgress');
    const results = uiAdapter.get('batchResults');

    const nrps = parseNrpList(inputMulti ? inputMulti.value : '');
    if (nrps.length === 0) { UI.showResult('warning', 'Data Tidak Valid', 'Masukkan NRP yang valid (4 atau 8 digit).'); return; }
    if (nrps.length > APP_CONFIG.BATCH_MAX_LIMIT) { UI.showResult('warning', 'Terlalu Banyak', `Maksimal ${APP_CONFIG.BATCH_MAX_LIMIT} NRP per batch.`); return; }

    const localBulan = parseInt(inputBulan ? inputBulan.value : '') || (new Date().getMonth() + 1);
    const localTahun = parseInt(inputTahun ? inputTahun.value : '') || new Date().getFullYear();
    state.batchBulan = Math.min(12, Math.max(1, localBulan));
    state.batchTahun = Math.min(2035, Math.max(2020, localTahun));
    state.batchTotal = nrps.length;
    state.batchAborted = false;
    state.batchRunId++;
    const batchRunId = state.batchRunId;
    const prof = startProfile('startBatchAnomalyCheck:init');

    state.batchQueue = nrps.map(nrp => ({ nrp, status: 'pending', msg: '' }));
    state.batchResults = nrps.map(nrp => ({
      nrp,
      found: true,
      nama: 'Memproses...',
      bagian: 'Sedang Memproses',
      seksi: 'Sedang Memproses',
      group: '-',
      anomalies: [],
      msg: 'Memproses...',
      processing: true
    }));
    localStorage.removeItem('qm-batch-results');
    localStorage.removeItem('qm-batch-bulan');
    localStorage.removeItem('qm-batch-tahun');
    if (results) results.classList.remove('is-visible');
    state.batchLogs = [];
    resetBatchProfile();
    const logBody = uiAdapter.get('logBody');
    if (logBody) renderSafe(logBody, '');
    pushLog(`Memulai batch check untuk ${nrps.length} NRP...`);

    if (btnCheck) { btnCheck.dataset.running = 'true'; btnCheck.textContent = 'Memproses...'; }
    if (progress) progress.classList.remove('is-hidden');

    // Render the initial pending/processing skeleton immediately
    _renderBatchResultsImmediate();

    if (results) {
      const firstCell = results.querySelector('.qm-batch-nrp-link') || results.querySelector('.qm-batch-cell');
      if (firstCell) firstCell.focus();
      else results.focus();
      const scrollContainer = results.closest('.panel-state');
      if (scrollContainer) {
        const targetScrollTop = results.offsetTop - (scrollContainer.clientHeight / 2) + (results.clientHeight / 2);
        scrollContainer.scrollTo({
          top: Math.max(0, targetScrollTop),
          behavior: 'smooth'
        });
      }
    }
    if (btnExport) { btnExport.classList.remove('qm-hidden'); btnExport.disabled = true; btnExport.style.opacity = '0.5'; }

    const poolSize = Math.min(APP_CONFIG.BATCH_POOL_SIZE, state.batchQueue.length);
    const workers = [];
    for (let i = 0; i < poolSize; i++) workers.push(processBatchWorker(batchRunId));
    Promise.all(workers).then(() => finishBatch(batchRunId));
    finishProfile(prof, { totalNrp: nrps.length, poolSize });
  }

  function handleBatchCancel() {
    state.batchAborted = true;
    state.batchRunId++;
    state.batchQueue = [];
    const btnCheck = uiAdapter.get('batchCheckButton');
    if (btnCheck) btnCheck.textContent = 'Membatalkan...';
    finishBatch(state.batchRunId);
  }

  function handleBatchClear() {
    if (uiAdapter.get('batchCheckButton')?.dataset.running) {
      uiAdapter.alert('Harap batalkan proses batch yang sedang berjalan terlebih dahulu.');
      return;
    }
    if (state.batchResults.length === 0) {
      uiAdapter.alert('Tidak ada data batch untuk dihapus.');
      return;
    }
    if (!uiAdapter.confirm('Apakah Anda yakin ingin menghapus semua hasil pemeriksaan batch ini?')) return;

    state.batchResults = [];
    state.batchQueue = [];
    state.batchTotal = 0;
    state.batchBulan = 0;
    state.batchTahun = 0;
    state.batchAborted = false;

    localStorage.removeItem('qm-batch-results');
    localStorage.removeItem('qm-batch-bulan');
    localStorage.removeItem('qm-batch-tahun');

    const inputMulti = uiAdapter.get('#qm-input-multi-nrp');
    if (inputMulti) inputMulti.value = '';

    renderBatchResults();

    const progress = uiAdapter.get('batchProgress');
    if (progress) progress.classList.add('is-hidden');

    const btnExport = uiAdapter.get('batchExportButton');
    if (btnExport) {
      btnExport.classList.add('qm-hidden');
      btnExport.style.opacity = '0';
      btnExport.disabled = true;
    }

    const statusBar = uiAdapter.get('batchStatus');
    if (statusBar) statusBar.textContent = '';

    const progressBar = uiAdapter.get('batchProgressBar');
    if (progressBar) progressBar.style.width = '0%';

    const container = uiAdapter.get('batchResults');
    if (container) {
      container.classList.remove('is-visible');
    }
  }

  async function processBatchWorker(batchRunId) {
    while (state.batchQueue.length > 0 && !state.batchAborted && batchRunId === state.batchRunId) {
      const item = state.batchQueue.shift();
      if (!item) return;
      const prof = startProfile('processBatchWorker:item', { nrp: item?.nrp });
      try {
        pushLog(`Memproses NRP ${item.nrp}...`);
        const emp = await fetchEmployee(item.nrp);
        if (batchRunId !== state.batchRunId || state.batchAborted) {
          finishProfile(prof, { nrp: item.nrp, stale: true });
          return;
        }
        item.found = emp.found;
        item.jk = emp.jk || '-';
        item.nama = emp.nama || '-';
        item.bagian = emp.bagian || '-';
        item.seksi = emp.seksi || '-';
        item.group = emp.group || '-';
        if (!emp.found) {
          item.anomalies = [];
          item.msg = '';
        } else {
          try {
            const scanRes = await fetchAttendance(item.nrp, state.batchBulan, state.batchTahun, item.bagian, item.seksi);
            if (batchRunId !== state.batchRunId || state.batchAborted) {
              finishProfile(prof, { nrp: item.nrp, stale: true });
              return;
            }
            item.anomalies = scanRes.anomalies || [];
            item.rekaps = scanRes.rekaps || null;
            item.msg = item.anomalies.length + ' anomali ditemukan';
          } catch (e) {
            item.msg = 'Gagal ambil data kehadiran';
            item.anomalies = [];
          }
        }
      } catch (e) {
        if (batchRunId !== state.batchRunId || state.batchAborted) {
          finishProfile(prof, { nrp: item.nrp, stale: true });
          return;
        }
        item.found = false;
        item.msg = 'Gagal akses HRIS';
        item.anomalies = [];
        pushLog(`Gagal memproses NRP ${item.nrp}: ${e.message}`, 'error');
      }
      if (batchRunId !== state.batchRunId || state.batchAborted) {
        finishProfile(prof, { nrp: item.nrp, stale: true });
        return;
      }
      pushBatchResult(item);
      if (item.found) pushLog(`Selesai memproses ${item.nama} (${item.nrp}).`, 'success');
      else pushLog(`NRP ${item.nrp} tidak ditemukan.`, 'error');
      const duration = finishProfile(prof, {
        nrp: item.nrp,
        found: !!item.found,
        anomalyCount: (item.anomalies || []).length
      });
      if (state.batchProfile && duration) {
        state.batchProfile.itemDurations.push(duration);
        if (state.batchProfile.itemDurations.length > PROFILE_CONFIG.SAMPLE_HISTORY_LIMIT) {
          state.batchProfile.itemDurations.shift();
        }
      }
    }
  }

  function finishBatch(batchRunId) {
    if (batchRunId !== state.batchRunId) return;
    const btnCheck = uiAdapter.get('batchCheckButton');
    const btnExport = uiAdapter.get('batchExportButton');
    const progress = uiAdapter.get('batchProgress');
    const statusBar = uiAdapter.get('batchStatus');
    const progressBar = uiAdapter.get('batchProgressBar');

    if (btnCheck) {
      btnCheck.textContent = 'Proses Batch';
      delete btnCheck.dataset.running;
    }
    const finishedCount = state.batchResults.filter(r => !r.processing).length;
    // Don't hide the progress bar immediately if we have export button
    if (btnExport && finishedCount > 0) {
      btnExport.classList.remove('qm-hidden');
      btnExport.disabled = false;
      btnExport.style.opacity = '1';
    } else {
      if (progress) progress.classList.add('is-hidden');
      if (btnExport) btnExport.classList.add('qm-hidden');
    }

    if (statusBar) statusBar.textContent = 'Selesai: ' + finishedCount + '/' + state.batchTotal + ' NRP';
    if (progressBar) progressBar.style.width = '100%';

    if (state.batchAborted) {
      UI.showResult('warning', 'Dibatalkan', 'Proses pemeriksaan batch dihentikan oleh pengguna.');
      pushLog('Proses batch dibatalkan oleh pengguna.', 'error');
    } else {
      pushLog(`Batch check selesai. Total ${finishedCount} NRP diproses.`);
      const container = uiAdapter.get('batchResults');
      const scrollContainer = container ? container.closest('.panel-state') : null;
      if (container && scrollContainer) {
        container.classList.remove('is-visible');
        void container.offsetWidth;
        container.classList.add('is-visible');
        const firstCell = container.querySelector('.qm-batch-nrp-link') || container.querySelector('.qm-batch-cell');
        if (firstCell) firstCell.focus();
        else container.focus();

        const targetScrollTop = container.offsetTop - (scrollContainer.clientHeight / 2) + (container.clientHeight / 2);
        scrollContainer.scrollTo({
          top: Math.max(0, targetScrollTop),
          behavior: 'smooth'
        });
      }
    }

    if (state.debug && state.batchProfile) {
      const renderTotal = Number((state.batchProfile.renderTotal || 0).toFixed(2));
      Logger.debug('[PROFILE] batchSummary', {
        processed: state.batchResults.length,
        renderTotalMs: renderTotal,
        recentWorkerMedianMs: Number(median((state.batchProfile.itemDurations || []).slice(-PROFILE_CONFIG.MEDIAN_SAMPLE_SIZE)).toFixed(2))
      });
      if (renderTotal >= PROFILE_CONFIG.HOT_BATCH_RENDER_TOTAL_MS) {
        Logger.warn(`Profiling: total render batch ${renderTotal.toFixed(2)}ms melewati ambang ${PROFILE_CONFIG.HOT_BATCH_RENDER_TOTAL_MS}ms.`);
      }
    }
  }

  function pushBatchResult(item) {
    const idx = state.batchResults.findIndex(r => r.nrp === item.nrp);
    if (idx !== -1) {
      state.batchResults[idx] = item;
    } else {
      state.batchResults.push(item);
    }
    localStorage.setItem('qm-batch-results', JSON.stringify(state.batchResults));
    localStorage.setItem('qm-batch-bulan', state.batchBulan);
    localStorage.setItem('qm-batch-tahun', state.batchTahun);
    renderBatchResults();
    updateBatchProgress();
  }

  function updateBatchProgress() {
    const finishedCount = state.batchResults.filter(r => !r.processing).length;
    const pct = state.batchTotal > 0 ? Math.round((finishedCount / state.batchTotal) * 100) : 0;
    const barEl = uiAdapter.get('batchProgressBar');
    const statusEl = uiAdapter.get('batchStatus');
    if (barEl) barEl.style.width = pct + '%';
    if (statusEl) statusEl.textContent = 'Memproses... ' + finishedCount + '/' + state.batchTotal;
  }

  let renderBatchTimeout;
  function renderBatchResults() {
    clearTimeout(renderBatchTimeout);
    renderBatchTimeout = setTimeout(() => {
      _renderBatchResultsImmediate();
    }, 150);
  }

  function _renderBatchResultsImmediate() {
    const prof = startProfile('renderBatchResults', { items: state.batchResults.length });
    const container = uiAdapter.get('batchResults');
    if (!container) {
      finishProfile(prof, { skipped: true });
      return;
    }
    if (state.batchResults.length === 0) {
      renderSafe(container, '');
      finishProfile(prof, { items: 0 });
      return;
    }

    let sumOtb = 0, sumOtl = 0, sumOta = 0, sumOtp = 0, sumHariKerja = 0;
    const ketMap = {};

    const sortedResults = [...state.batchResults].sort((a, b) => {
      const bagA = (a.bagian || '').toLowerCase();
      const bagB = (b.bagian || '').toLowerCase();
      if (bagA !== bagB) return bagA.localeCompare(bagB);

      const sekA = (a.seksi || '').toLowerCase();
      const sekB = (b.seksi || '').toLowerCase();
      if (sekA !== sekB) return sekA.localeCompare(sekB);

      return (parseInt(a.nrp) || 0) - (parseInt(b.nrp) || 0);
    });

    const tree = {};
    sortedResults.forEach(item => {
      if (item.rekaps) {
        sumOtb += item.rekaps.otb; sumOtl += item.rekaps.otl;
        sumOta += item.rekaps.ota; sumOtp += item.rekaps.otp;
        sumHariKerja += item.rekaps.hariKerja;
        for (const [k, v] of Object.entries(item.rekaps.keterangan)) { ketMap[k] = (ketMap[k] || 0) + v; }
      }

      const bag = item.bagian || 'Tanpa Bagian';
      const sek = item.seksi || 'Tanpa Seksi';
      if (!tree[bag]) tree[bag] = {};
      if (!tree[bag][sek]) tree[bag][sek] = [];
      tree[bag][sek].push(item);
    });

    var html = '<table class="qm-batch-table"><tbody>';

    let bagIdx = 0;
    for (const bag in tree) {
      const bagSafeId = 'bag-' + (bagIdx++);
      html += `<tr class="qm-batch-group-header qm-batch-header-bg expanded" data-target=".${bagSafeId}">`;
      html += `<td colspan="5" class="qm-batch-bagian-cell"><div class="qm-flex qm-items-center qm-justify-between qm-w-full"><span>${escapeHtml(bag)}</span><span class="qm-chevron qm-accordion-chevron">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>
              </span></div></td></tr>`;

      html += `<tr class="qm-batch-group-row ${bagSafeId} qm-table-header qm-batch-sub-header-bg qm-table-row">`;
      html += '<td class="qm-batch-cell-header qm-batch-col-nrp">NRP</td>';
      html += '<td class="qm-batch-cell-header qm-batch-col-nama">Nama</td>';
      html += '<td class="qm-batch-cell-header qm-batch-col-ot">Lembur</td>';
      html += '<td class="qm-batch-cell-header qm-batch-col-ket">Keterangan</td>';
      html += '<td class="qm-batch-cell-header qm-batch-col-msg">Masalah / Anomali</td>';
      html += '</tr>';

      let sekIdx = 0;
      for (const sek in tree[bag]) {
        const sekSafeId = bagSafeId + '-sek-' + (sekIdx++);
        html += `<tr class="qm-batch-group-row ${bagSafeId} qm-batch-seksi-header qm-batch-sub-header-bg qm-table-row expanded" data-target=".${sekSafeId}">`;
        html += `<td colspan="5" class="qm-batch-seksi-cell" style="padding-left: 32px;"><div class="qm-flex qm-items-center qm-justify-between qm-w-full"><span>${escapeHtml(sek)} <span class="qm-batch-seksi-count">(${tree[bag][sek].length})</span></span><span class="qm-chevron qm-accordion-chevron">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>
              </span></div></td></tr>`;

        tree[bag][sek].forEach(item => {
          var nrpLink = item.nrp
            ? '<a href="#" class="qm-batch-nrp-link" data-nrp="' + escapeHtml(item.nrp) + '" style="padding-left: 64px;">' + escapeHtml(item.nrp) + '</a>'
            : '-';
          var nama = item.processing ? '<span class="qm-text-muted" style="opacity: 0.6;">Memproses...</span>' : (item.found ? escapeHtml(item.nama || '-') : '<span class="qm-batch-not-found">Tidak Ditemukan</span>');
          var masalahHtml = '';

          if (item.processing) {
            masalahHtml = '<span class="status-pill status-pending" style="display: inline-flex; align-items: center; gap: 6px;"><span class="qm-spinner qm-spinner-xs" style="width: 12px; height: 12px; border-width: 1.5px;"></span>Memproses...</span>';
          } else if (item.anomalies && item.anomalies.length > 0) {
            const byTgl = {};
            item.anomalies.forEach(a => {
              const t = String(a.tgl);
              if (!byTgl[t]) byTgl[t] = [];
              byTgl[t].push(a);
            });

            let listItems = '';
            const sortedDates = Object.keys(byTgl).sort((a, b) => parseInt(a) - parseInt(b));
            for (const tgl of sortedDates) {
              const anomaliesTgl = byTgl[tgl];
              const firstAnomWithLink = anomaliesTgl.find(a => a.link);
              const firstLink = firstAnomWithLink?.link;
              const firstFullDate = firstAnomWithLink?.fullDate || '';

              let fixBtn = '';
              if (firstLink) {
                let finalLink = firstLink;
                let titleStr = firstAnomWithLink.msg || 'Fix Anomali';

                if (titleStr.includes('SPKL')) {
                  const isOS = item.nrp && item.nrp.length === 8;
                  const base = spklBaseUrl(item.nrp);
                  const bulanStr = String(state.batchBulan).padStart(2, '0');
                  const tglStr = String(tgl).padStart(2, '0');
                  const shiftVal = firstAnomWithLink.shift || '';
                  const shiftParam = shiftVal ? `&shift=${shiftVal}` : '';
                  const fDate = firstAnomWithLink.fullDate || `${state.batchTahun}-${bulanStr}-${tglStr}`;

                  const ev = v => (v === '-' ? '' : encodeURIComponent(v));
                  finalLink = `${base}?tahun=${state.batchTahun}&bulan=${bulanStr}&kode_bagian=${ev(item.bagian)}&kode_seksi=${ev(item.seksi)}&kode_group=${ev(item.group)}&nrp=${item.nrp}&qm_auto_spkl_fix=1&full_date=${fDate}${shiftParam}`;
                  titleStr = 'Cek Halaman SPKL';
                }

                const tglPad = String(tgl).padStart(2, '0');
                const fDate = firstAnomWithLink.fullDate || `${state.batchTahun}-${String(state.batchBulan).padStart(2, '0')}-${tglPad}`;

                fixBtn = `<button class="qm-batch-fix-btn" title="${escapeHtml(titleStr)}" data-fix-link="${escapeHtml(finalLink)}" data-fix-date="${escapeHtml(tglPad)}" data-full-date="${escapeHtml(fDate)}">Fix</button>`;
              }

              listItems += '<div class="qm-batch-date-row">';
              listItems += '<div class="qm-batch-date-header qm-flex qm-items-center qm-justify-between"><span><b>Tgl ' + escapeHtml(tgl) + '</b>' + fixBtn + '</span><span class="qm-chevron qm-accordion-chevron"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg></span></div>';
              listItems += '<div class="qm-batch-date-content qm-hidden">';
              // Deduplicate anomalies by message for the same date
              const msgMap = new Map();
              anomaliesTgl.forEach(a => {
                if (!msgMap.has(a.msg)) msgMap.set(a.msg, []);
                msgMap.get(a.msg).push(a.col);
              });

              msgMap.forEach((cols, msg) => {
                const uniqueCols = [...new Set(cols)].filter(Boolean);
                const colNames = uniqueCols.map(c => escapeHtml(c)).join(', ');
                const colPrefix = colNames ? colNames + ': ' : '';
                listItems += '<div class="qm-batch-anomaly-detail" style="font-size: 12px">• ' + colPrefix + escapeHtml(msg) + '</div>';
              });
              listItems += '</div></div>';
            }
            masalahHtml = '<div class="qm-batch-masalah-scroll">' + listItems + '</div>';
          } else if (item.found) {
            masalahHtml = '<span class="status-pill status-active">Aktif</span>';
          } else {
            masalahHtml = `<span class="status-pill status-error">${escapeHtml(item.msg || 'Tidak Ditemukan')}</span>`;
          }

          let lemburHtml = '';
          let ketHtml = '';
          if (item.processing) {
            lemburHtml = '<span style="color: var(--color-text-muted); opacity: 0.5;">-</span>';
            ketHtml = '<span style="color: var(--color-text-muted); opacity: 0.5;">-</span>';
          } else {
            const rk = item.rekaps || { otb: 0, otl: 0, ota: 0, otp: 0, keterangan: {} };
            const otValues = `B: ${rk.otb.toFixed(1)}<br>L: ${rk.otl.toFixed(1)}<br>A: ${rk.ota.toFixed(1)}<br>P: ${rk.otp.toFixed(1)}`;

            const ketKeys = ['CT', 'CH', 'SD', 'I', 'IS', 'IA', 'A'];
            const ketStr = ketKeys.map(k => `${k}: ${rk.keterangan[k] || 0}`).join('<br>');

            lemburHtml = `<div class="qm-text-xs qm-font-mono" style="color: var(--qm-olive);">${otValues}</div>`;
            ketHtml = `<div class="qm-text-xs qm-font-mono" style="color: var(--qm-stone); opacity: 0.8;">${ketStr}</div>`;
          }

          html += `<tr class="qm-batch-group-row ${bagSafeId} ${sekSafeId} qm-batch-item-row qm-table-row">`;
          html += '<td class="qm-batch-cell">' + nrpLink + '</td>';
          html += '<td class="qm-batch-cell qm-batch-nama">' + nama + '</td>';
          html += '<td class="qm-batch-cell qm-batch-col-lembur">' + lemburHtml + '</td>';
          html += '<td class="qm-batch-cell qm-batch-col-ket">' + ketHtml + '</td>';
          html += '<td class="qm-batch-cell qm-batch-col-masalah">' + masalahHtml + '</td>';
          html += '</tr>';
        });
      }
    }

    html += '</tbody></table>';

    renderSafe(container, html);
    container.classList.remove('is-visible');
    void container.offsetWidth; // Force DOM reflow
    container.classList.add('is-visible');
    const scrollContainer = container.closest('.panel-state');
    if (scrollContainer && state.batchResults.length === 1) {
      const firstCell = container.querySelector('.qm-batch-nrp-link') || container.querySelector('.qm-batch-cell');
      if (firstCell) firstCell.focus();
      else container.focus();
      const targetScrollTop = container.offsetTop - (scrollContainer.clientHeight / 2) + (container.clientHeight / 2);
      scrollContainer.scrollTo({
        top: Math.max(0, targetScrollTop),
        behavior: 'smooth'
      });
    }
    const duration = finishProfile(prof, { items: state.batchResults.length });
    if (state.batchProfile) state.batchProfile.renderTotal += duration;
  }

  function exportBatchResults() {
    if (state.batchResults.length === 0) { uiAdapter.alert('Tidak ada hasil untuk diekspor.'); return; }
    if (typeof XLSX !== 'undefined') {
      const ketHeaders = ['CT', 'CH', 'SD', 'I', 'IS', 'IA', 'A'];
      const headers = ['Bagian', 'Seksi', 'NRP', 'Nama', 'JK', 'OTB', 'OTL', 'OTA', 'OTP', 'Hari Kerja', ...ketHeaders, 'Jml Anomali', 'Detail Anomali'];
      var wsData = [headers];

      state.batchResults.forEach(function (r) {
        var detailAnomali = (r.anomalies || []).map(function (a) {
          return 'Tgl ' + a.tgl + ' ' + a.col + ': ' + a.msg;
        }).join('; ');

        const rk = r.rekaps || { otb: 0, otl: 0, ota: 0, otp: 0, hariKerja: 0, keterangan: {} };
        const row = [
          r.bagian || '-',
          r.seksi || '-',
          r.nrp,
          r.nama || '-',
          r.jk || '-',
          rk.otb,
          rk.otl,
          rk.ota,
          rk.otp,
          rk.hariKerja
        ];

        ketHeaders.forEach(k => row.push(rk.keterangan[k] || 0));
        row.push(r.anomalies ? r.anomalies.length : 0);
        row.push(detailAnomali);
        wsData.push(row);
      });
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Batch Check');
      XLSX.writeFile(wb, `Batch_Check_NRP_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } else {
      uiAdapter.alert('Library XLSX gagal dimuat.');
    }
  }

  function renderAnomalies() {
    const prof = startProfile('renderAnomalies', { count: state.anomalies.length });
    const badge = uiAdapter.get('anomalyBadge');
    const list = uiAdapter.get('anomalyList');
    if (!badge || !list) {
      finishProfile(prof, { skipped: true });
      return;
    }

    if (state.anomalies.length > 0) {
      badge.textContent = state.anomalies.length;
      badge.classList.remove('qm-hidden');
      badge.classList.add('qm-visible-inline-flex');

      const grouped = {};
      state.anomalies.forEach(a => {
        if (!grouped[a.tgl]) grouped[a.tgl] = [];
        grouped[a.tgl].push(a);
      });

      const sortedKeys = Object.keys(grouped).sort((a, b) => parseInt(a) - parseInt(b));
      let html = '';

      for (const tgl of sortedKeys) {
        const items = grouped[tgl];
        const spklStatusItem = items.find(a => a.col === 'SPKL Status' && a.link);
        const itemWithLink = spklStatusItem || items.find(a => a.link);
        let btnText = 'Perbaikan';

        if (itemWithLink) {
          if (itemWithLink.msg === 'Buka Halaman Kehadiran') btnText = 'Buka Halaman Kehadiran';
          else if (itemWithLink.msg && itemWithLink.msg.includes('Pulang awal')) btnText = 'Cek Kehadiran';
          else if (itemWithLink.msg === 'Cek Distribusi' || (itemWithLink.msg && itemWithLink.msg.startsWith('Shift Kosong')) || itemWithLink.msg === 'Shift muncul') btnText = 'Fix Distribusi';
          else if (itemWithLink.msg === 'SPKL Disetujui (Cek Jenis OT)') btnText = 'Cek Halaman SPKL';
          else if (itemWithLink.msg === 'SPKL Belum Disetujui') btnText = 'Cek SPKL Online';
          else if (itemWithLink.msg === 'SPKL Ditolak') btnText = 'Cek SPKL (Ditolak)';
          else if (itemWithLink.msg === 'SPKL Online tidak ada') btnText = 'Input SPKL Online';
          else if (itemWithLink.msg === 'SPKL Online tidak sesuai') btnText = 'Cek SPKL Online';
        }

        const fixBtn = itemWithLink
          ? `<button class="qm-btn-fix-pill" data-fix-link="${escapeHtml(itemWithLink.link)}" data-fix-date="${escapeHtml(tgl)}" data-full-date="${escapeHtml(itemWithLink.fullDate || '')}" title="${escapeHtml(btnText)}">${btnText}</button>`
          : '';

        const detailsHtml = items.map(a => {
          const type = escapeHtml(COL_LABELS[a.colIndex] || a.col || ('Kolom ' + a.colIndex));
          return `
            <div class="qm-anomaly-card">
              <div class="qm-anomaly-card-type">${type}</div>
              <div class="qm-anomaly-card-msg">${escapeHtml(a.msg)}</div>
            </div>
          `;
        }).join('');

        html += `
          <div class="qm-anomaly-item">
            <div class="qm-anomaly-left">
              <span class="qm-anomaly-date">Tgl ${tgl}</span>
              <span class="qm-badge qm-text-muted" style="background:var(--qm-sand); border-radius:4px; padding:2px 6px; font-size:10px;">${items.length}</span>
            </div>
            <div class="qm-anomaly-content">
              ${detailsHtml}
            </div>
            <div class="qm-anomaly-actions">
              ${fixBtn}
            </div>
          </div>
        `;
      }
      renderSafe(list, html);
    } else {
      badge.classList.add('qm-hidden');
      badge.classList.remove('qm-visible-inline-flex');
      renderSafe(list, '<div class="qm-anomaly-empty-state qm-text-center qm-text-muted qm-mt-xl">Tidak ada anomali ditemukan.</div>');
    }
    finishProfile(prof, { count: state.anomalies.length });
  }

  /* ============================================================
   * 9. SPKL & BARCODE CHECKS
   * ============================================================ */

  function isSameCalendarDay(a, b) {
    return !!a && !!b
      && a.getFullYear() === b.getFullYear()
      && a.getMonth() === b.getMonth()
      && a.getDate() === b.getDate();
  }

  function formatDisplayDate(dateStr) {
    const parsed = parseHrisDate(dateStr);
    if (!parsed) return dateStr;
    return parsed.toLocaleDateString('id-ID', {
      day: '2-digit',
      month: 'long',
      year: 'numeric'
    });
  }

  function extractDateTextCandidate(text) {
    const clean = String(text || '').trim();
    if (!clean) return '';

    const matchers = [
      /\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b/,
      /\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b/,
      /\b\d{1,2}-[A-Za-z]{3}-\d{4}\b/
    ];

    for (const matcher of matchers) {
      const match = clean.match(matcher);
      if (match) return match[0];
    }
    return '';
  }

  function parseBarcodeRowDate(cellTexts, dateColIdx, fallbackDate) {
    const candidates = [];
    if (dateColIdx !== -1 && cellTexts[dateColIdx]) candidates.push(cellTexts[dateColIdx]);
    cellTexts.forEach(text => {
      if (text && !candidates.includes(text)) candidates.push(text);
    });

    for (const candidate of candidates) {
      const extracted = extractDateTextCandidate(candidate);
      if (extracted) {
        const parsed = parseHrisDate(extracted);
        if (parsed) return { date: parsed, raw: extracted };
      }

      if (/^\d{1,2}$/.test(candidate) && fallbackDate) {
        const day = parseInt(candidate, 10);
        if (day >= 1 && day <= 31) {
          return {
            date: new Date(fallbackDate.getFullYear(), fallbackDate.getMonth(), day),
            raw: candidate.padStart(2, '0')
          };
        }
      }
    }

    return null;
  }

  function normalizeBarcodeStatusText(text) {
    const clean = String(text || '').trim();
    const lower = clean.toLowerCase();
    if (lower.includes('masuk')) return 'Masuk';
    if (lower.includes('keluar')) return 'Keluar';
    return clean || '-';
  }

  function extractBarcodeRowStatus(cellTexts, statusColIdx) {
    const candidates = [];
    if (statusColIdx !== -1 && cellTexts[statusColIdx]) candidates.push(cellTexts[statusColIdx]);
    cellTexts.forEach(text => {
      if (text && !candidates.includes(text)) candidates.push(text);
    });

    const matched = candidates.find(text => /masuk|keluar/i.test(text));
    return normalizeBarcodeStatusText(matched || candidates[0] || '');
  }

  function extractBarcodeRowTime(cellTexts, timeColIdx) {
    const candidates = [];
    if (timeColIdx !== -1 && cellTexts[timeColIdx]) candidates.push(cellTexts[timeColIdx]);
    cellTexts.forEach(text => {
      if (text && !candidates.includes(text)) candidates.push(text);
    });

    for (const candidate of candidates) {
      const match = String(candidate).match(/\b\d{1,2}[:.]\d{2}(?::\d{2})?\b/);
      if (match) return match[0];
    }
    return '-';
  }

  function parseBarcodeAttendanceSummary(doc, nrp, startDateStr, endDateStr) {
    const start = parseHrisDate(startDateStr);
    const end = parseHrisDate(endDateStr || startDateStr);
    if (!start || !end) throw new Error('Tanggal tidak valid.');

    const headerCells = Array.from(doc.querySelectorAll('table thead th'));
    const headers = headerCells.map(cell => normalizeEmployeeHeader(cell.textContent));
    const dateColIdx = firstMatchedHeaderIndex(headers, ['tanggal', 'tgl', 'date']);
    const statusColIdx = firstMatchedHeaderIndex(headers, ['status']);
    const timeColIdx = firstMatchedHeaderIndex(headers, ['jam', 'waktu', 'time', 'scan']);

    const rows = Array.from(doc.querySelectorAll('table tbody tr'));
    const entries = [];

    rows.forEach((row, index) => {
      const cells = Array.from(row.querySelectorAll('td'));
      if (cells.length === 0) return;

      const cellTexts = cells.map(cell => normalizeEmployeeCellText(cell)).filter(Boolean);
      if (cellTexts.length === 0) return;

      // Use start date as fallback for year/month if missing in row
      const parsedDate = parseBarcodeRowDate(cellTexts, dateColIdx, start);
      if (!parsedDate) return;

      const d = parsedDate.date;
      // Filter by range
      if (d.getTime() < start.getTime() || d.getTime() > end.getTime()) return;

      const status = extractBarcodeRowStatus(cellTexts, statusColIdx);
      const time = extractBarcodeRowTime(cellTexts, timeColIdx);
      const rawSummary = cellTexts.join(' | ');

      const actions = [];
      row.querySelectorAll('a, button').forEach(el => {
        const text = el.textContent.trim();
        if (/edit|ubah|hapus|delete/i.test(text)) {
          actions.push({
            text: text,
            href: el.getAttribute('href'),
            onclick: el.getAttribute('onclick'),
            className: el.className,
            dataTarget: el.getAttribute('data-target')
          });
        }
      });

      entries.push({
        key: `${nrp}-${parsedDate.raw || index}-${index}`,
        date: d,
        dateText: parsedDate.raw || formatDisplayDate(d),
        status,
        time,
        rawSummary,
        actions
      });
    });

    // Sort by date then time
    entries.sort((a, b) => a.date - b.date || (a.time || '').localeCompare(b.time || ''));

    const masukEntries = entries.filter(entry => entry.status === 'Masuk');
    const keluarEntries = entries.filter(entry => entry.status === 'Keluar');

    let dateLabel = formatDisplayDate(startDateStr);
    if (endDateStr && endDateStr !== startDateStr) {
      dateLabel += ` - ${formatDisplayDate(endDateStr)}`;
    }

    return {
      nrp,
      dateLabel,
      entries,
      hasMasuk: masukEntries.length > 0,
      hasKeluar: keluarEntries.length > 0
    };
  }

  async function fetchBarcodeAttendanceSummary(nrp, startDate, endDate) {
    const start = parseHrisDate(startDate);
    if (!start) throw new Error('Tanggal tidak valid.');

    // Fetch month of start date. Range across months is not supported yet for simplicity.
    const bulan = String(start.getMonth() + 1).padStart(2, '0');
    const tahun = start.getFullYear();
    const url = routeByNrp(nrp, ROUTES.ABSEN_BARCODE, ROUTES.ABSEN_BARCODE_OS)(tahun, bulan, nrp);
    const html = await hrisFetch(url);
    return parseBarcodeAttendanceSummary(parseHTML(html), nrp, startDate, endDate);
  }

  async function fetchSpklSummary(nrp, startDate, endDate) {
    const start = parseHrisDate(startDate);
    if (!start) throw new Error('Start Date tidak valid.');

    const bulan = String(start.getMonth() + 1).padStart(2, '0');
    const tahun = start.getFullYear();
    const url = spklListUrl(nrp, bulan, tahun);
    const html = await hrisFetch(url);
    const summary = parseSpklSummary(parseHTML(html), nrp, startDate, endDate);
    summary.bulan = bulan;
    summary.tahun = tahun;
    return summary;
  }

  function parseSpklSummary(doc, nrp, startDate, endDate) {
    const entries = [];
    const rows = doc.querySelectorAll('table tbody tr');

    // Parse range for filtering
    const startObj = startDate ? parseHrisDate(startDate) : null;
    const endObj = endDate ? parseHrisDate(endDate) : null;

    // Detect header index
    let tglIdx = -1, otIdx = -1, statusIdx = -1, shiftIdx = -1, mskIdx = -1, klrIdx = -1, jamOtIdx = -1;
    doc.querySelectorAll('table th').forEach((th, i) => {
      const txt = th.textContent.trim().toLowerCase();
      if (txt.includes('tanggal') || txt === 'tgl') tglIdx = i;
      if (txt.includes('kode') || txt.includes('jenis') || txt.includes('ot')) otIdx = i;
      if (txt.includes('status')) statusIdx = i;
      if (txt.includes('shift')) shiftIdx = i;
      if (txt.includes('jam masuk') || txt.includes('jam awal') || txt.includes('mulai')) mskIdx = i;
      if (txt.includes('jam pulang') || txt.includes('jam akhir') || txt.includes('selesai')) klrIdx = i;
      if (txt.includes('jam ot') || txt.includes('jam lembur') || txt.includes('tambahan')) jamOtIdx = i;
    });

    rows.forEach(tr => {
      const tds = tr.querySelectorAll('td');
      if (tds.length < 3) return;

      const dateText = tglIdx !== -1 ? tds[tglIdx].textContent.trim() : '';
      if (!dateText) return;

      // Filter by date range if provided
      if (startObj && endObj) {
        const entryDate = parseHrisDate(dateText);
        if (entryDate) {
          if (entryDate < startObj || entryDate > endObj) return;
        }
      }

      const rawOt = otIdx !== -1 ? tds[otIdx].textContent.trim().replace(/\s+/g, ' ').toUpperCase() : '-';
      const otMap = {
        'OT BIASA': '1',
        'LONG SHIFT': '2',
        'NON STOP': '3',
        'OT AWAL': '4',
        'NO REST (AWAL)': '5A',
        'NO REST (TENGAH)': '5B',
        'NO REST (AKHIR)': '5C',
        'STANDBY': '6',
        'LAIN-LAIN': '7',
        'OVERTIME': 'OT'
      };

      const rawShift = shiftIdx !== -1 ? tds[shiftIdx].textContent.trim().replace(/\s+/g, ' ').toUpperCase() : '';
      const shiftMap = {
        'SHIFT I': '1',
        'SHIFT II': '2',
        'SHIFT III': '3',
        'LONG SHIFT I': '4',
        'LONG SHIFT II': '5'
      };

      const actions = [];
      tr.querySelectorAll('a, button').forEach(el => {
        const text = el.textContent.trim();
        let href = el.getAttribute('href');
        const dataTarget = el.getAttribute('data-target');

        // If it's a modal button without href, try to construct edit URL
        if (!href && dataTarget && dataTarget.startsWith('#editData')) {
          const id = dataTarget.replace('#editData', '');
          href = `/spkl/edit/${id}`;
        }

        if (text || href || dataTarget) {
          actions.push({
            text: text || 'Edit',
            href: href,
            modalTarget: dataTarget,
            onclick: el.getAttribute('onclick')
          });
        }
      });

      entries.push({
        dateText,
        otCode: otMap[rawOt] || rawOt,
        status: statusIdx !== -1 ? tds[statusIdx].textContent.trim() : '-',
        shift: shiftMap[rawShift] || rawShift,
        jamAwal: mskIdx !== -1 ? tds[mskIdx].textContent.trim() : '-',
        jamAkhir: klrIdx !== -1 ? tds[klrIdx].textContent.trim() : '-',
        jamOt: jamOtIdx !== -1 ? tds[jamOtIdx].textContent.trim() : '-',
        actions
      });
    });

    return { nrp, entries, startDate, endDate };
  }

  function renderAttendanceCheckResult() {
    const wrap = uiAdapter.get('attendanceCheckResult');
    const btn = uiAdapter.get('attendanceCheckButton');
    if (!wrap) return;

    const current = state.attendanceCheck || createEmptyAttendanceCheck();
    if (btn) {
      btn.disabled = current.loading;
      btn.textContent = current.loading ? 'Checking...' : 'Check';
    }

    if (current.loading) {
      uiAdapter.html('attendanceCheckResult', '<div class="qm-hadir-check-card"><div class="qm-flex qm-items-center qm-gap-s"><span class="qm-spinner"></span><span>Memuat ringkasan kehadiran...</span></div></div>');
      return;
    }

    if (current.error) {
      uiAdapter.html('attendanceCheckResult', `<div class="qm-hadir-check-card qm-hadir-check-error">${escapeHtml(current.error)}</div>`);
      return;
    }

    if (!current.summary) {
      uiAdapter.html('attendanceCheckResult', '<div class="qm-hadir-check-card qm-hadir-check-empty">Pilih NRP dan tanggal untuk melihat ringkasan barcode harian.</div>');
      return;
    }

    const summary = current.summary;
    const emp = summary.employee || {};
    const totalEntries = summary.entries.length;
    const statusLabel = totalEntries === 0
      ? 'Tidak ada data'
      : (summary.hasMasuk && summary.hasKeluar ? 'LENGKAP' : 'PERLU CEK');
    const statusClass = totalEntries === 0 || !summary.hasMasuk || !summary.hasKeluar ? 'err' : 'ok';

    // Group entries by date
    const grouped = {};
    summary.entries.forEach((entry, idx) => {
      if (!grouped[entry.dateText]) {
        grouped[entry.dateText] = {
          dateText: entry.dateText,
          masuk: null,
          keluar: null,
          indices: []
        };
      }
      if (entry.status === 'Masuk') grouped[entry.dateText].masuk = entry;
      else if (entry.status === 'Keluar') grouped[entry.dateText].keluar = entry;
      grouped[entry.dateText].indices.push(idx);
    });

    const sortedDates = Object.keys(grouped).sort((a, b) => {
      const [da, ma, ya] = a.split('-');
      const [db, mb, yb] = b.split('-');
      return new Date(ya, ma - 1, da) - new Date(yb, mb - 1, db);
    });

    const detailHtml = sortedDates.length === 0
      ? '<div class="qm-hadir-check-detail-empty">Tidak ada baris barcode pada rentang tanggal ini.</div>'
      : `
        <div class="qm-hadir-table-container">
          <table class="qm-hadir-table">
            <thead>
              <tr>
                <th style="width: 30px;"></th>
                <th style="width: 80px;">Tanggal</th>
                <th>Masuk</th>
                <th>Keluar</th>
              </tr>
            </thead>
            <tbody>
              ${sortedDates.map(dateStr => {
        const group = grouped[dateStr];
        const m = group.masuk;
        const k = group.keluar;

        const renderTime = (entry) => {
          if (!entry) return '<span class="qm-text-muted qm-font-xs">-</span>';
          const isDelete = (entry.actions || []).some(a => /hapus|delete/i.test(a.text));
          const deleteBtn = isDelete
            ? `<button type="button" class="qm-btn-text qm-text-danger qm-btn-barcode-delete" data-url="${toAbsoluteHrisUrl((entry.actions.find(a => /hapus|delete/i.test(a.text))).href)}" style="margin-left: 4px; font-size: 10px;">&times;</button>`
            : '';
          return `<span class="qm-font-mono">${escapeHtml(entry.time || '-')}</span>${deleteBtn}`;
        };

        return `
                  <tr>
                    <td><input type="checkbox" class="qm-hadir-batch-cb" data-indices="${group.indices.join(',')}"></td>
                    <td><span style="font-size: var(--qm-font-xs);">${escapeHtml(dateStr)}</span></td>
                    <td>${renderTime(m)}</td>
                    <td>${renderTime(k)}</td>
                  </tr>
                `;
      }).join('')}
            </tbody>
          </table>
        </div>
      `;

    uiAdapter.html('attendanceCheckResult', `
      <div class="qm-hadir-check-card">
        <div class="qm-hadir-check-summary">
          <div class="qm-hadir-check-head">
            <div>
              <div class="qm-hadir-check-title" style="font-size: var(--qm-font-m);">
                ${escapeHtml(emp.nama || summary.nrp)}
                <span class="qm-text-muted qm-ml-s qm-font-normal" style="font-size: var(--qm-font-s);">${escapeHtml(summary.nrp)}</span>
                <span class="qm-text-muted qm-ml-m qm-font-normal" style="font-size: var(--qm-font-xs); border-left: 1px solid rgba(0,0,0,0.1); padding-left: 10px;">
                  ${escapeHtml(emp.bagian || '-')} • ${escapeHtml(emp.seksi || '-')} • ${escapeHtml(emp.group || '-')}
                </span>
              </div>
            </div>
            <span class="qm-badge ${statusClass}">${escapeHtml(statusLabel)}</span>
          </div>
        </div>
        <div class="qm-hadir-check-detail">
          <div class="qm-flex qm-justify-between qm-items-center qm-mb-s" style="padding-bottom: 4px; border-bottom: 1px solid var(--qm-border-warm);">
            <div class="qm-flex qm-items-center">
              <input type="checkbox" id="qm-hadir-batch-cb-all" class="qm-mr-s">
              <label for="qm-hadir-batch-cb-all" class="qm-hadir-check-detail-title qm-m-0 qm-cursor-pointer">PILIH SEMUA</label>
            </div>
            <button type="button" class="qm-btn qm-btn-primary qm-btn-sm" id="qm-btn-hadir-batch-edit" style="padding: 2px 8px; font-size: 10px;">Edit Terpilih</button>
          </div>
          ${detailHtml}
        </div>
      </div>
    `);
  }

  function openHadirBatchEdit() {
    const checkboxes = document.querySelectorAll('.qm-hadir-batch-cb:checked');
    if (checkboxes.length === 0) {
      uiAdapter.alert('Pilih setidaknya satu baris untuk diedit.');
      return;
    }

    const indices = [];
    checkboxes.forEach(cb => {
      const idxs = cb.dataset.indices.split(',').map(n => parseInt(n, 10));
      indices.push(...idxs);
    });
    state.activeHadirBatchIndices = indices;

    // Clear previous inputs
    uiAdapter.value('hadirEditJamMasuk', '');
    uiAdapter.value('hadirEditJamKeluar', '');
    uiAdapter.value('hadirEditStatus', '');

    const modal = uiAdapter.get('hadirEditModal');
    if (modal) {
      modal.classList.remove('qm-hidden');
      setTimeout(() => uiAdapter.focus('hadirEditJamMasuk'), 100);
    }
  }

  function closeHadirBatchEdit() {
    const modal = uiAdapter.get('hadirEditModal');
    if (modal) modal.classList.add('qm-hidden');
    state.activeHadirBatchIndices = [];
  }

  function toggleHadirRowInputs(checkbox) {
    const row = checkbox.closest('tr');
    if (!row) return;

    const cells = row.querySelectorAll('td');
    const masukCell = cells[2];
    const keluarCell = cells[3];

    if (checkbox.checked) {
      if (!row.dataset.originalMasuk) row.dataset.originalMasuk = masukCell.innerHTML;
      if (!row.dataset.originalKeluar) row.dataset.originalKeluar = keluarCell.innerHTML;

      const mTime = masukCell.querySelector('.qm-font-mono')?.textContent.trim() || '';
      const kTime = keluarCell.querySelector('.qm-font-mono')?.textContent.trim() || '';

      masukCell.innerHTML = `<input type="time" class="qm-hadir-inline-input qm-input" value="${mTime.includes(':') ? mTime.substring(0, 5) : ''}" data-prev="${mTime}" style="width: 85px; padding: 2px; height: 24px; font-size: 11px;">`;
      keluarCell.innerHTML = `<input type="time" class="qm-hadir-inline-input qm-input" value="${kTime.includes(':') ? kTime.substring(0, 5) : ''}" data-prev="${kTime}" style="width: 85px; padding: 2px; height: 24px; font-size: 11px;">`;
    } else {
      if (row.dataset.originalMasuk) masukCell.innerHTML = row.dataset.originalMasuk;
      if (row.dataset.originalKeluar) keluarCell.innerHTML = row.dataset.originalKeluar;
    }
  }

  function startHadirPageLoop() {
    const checkboxes = document.querySelectorAll('.qm-hadir-batch-cb:checked');
    if (checkboxes.length === 0) {
      uiAdapter.alert('Pilih setidaknya satu baris.');
      return;
    }

    const tasks = [];
    const current = state.attendanceCheck;
    if (!current || !current.summary) return;

    checkboxes.forEach(cb => {
      const row = cb.closest('tr');
      const inputs = row.querySelectorAll('.qm-hadir-inline-input');
      const mVal = inputs[0]?.value;
      const kVal = inputs[1]?.value;
      const mPrev = inputs[0]?.dataset.prev || '';
      const kPrev = inputs[1]?.dataset.prev || '';

      if (!cb.dataset.indices) return;
      const indices = cb.dataset.indices.split(',').map(n => parseInt(n, 10));

      indices.forEach(idx => {
        if (isNaN(idx)) return;
        const entry = current.summary.entries[idx];
        if (!entry) return;
        const editAction = (entry.actions || []).find(act => /edit|ubah/i.test(act.text));
        if (!editAction) return;

        const isMasuk = entry.status === 'Masuk';
        const targetTime = isMasuk ? mVal : kVal;
        const prevTime = (isMasuk ? mPrev : kPrev).substring(0, 5);

        // Logic: If target is empty but prev was not -> DELETE
        // If target is different from prev -> EDIT
        const isDelete = prevTime && !targetTime;
        const isEdit = targetTime && targetTime !== prevTime;

        if (isEdit || isDelete) {
          // Format date for matching (e.g., 01-Apr-2026)
          const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
          const d = entry.date;
          const dateMatch = `${String(d.getDate()).padStart(2, '0')}-${months[d.getMonth()]}-${d.getFullYear()}`;

          tasks.push({
            index: idx,
            type: isDelete ? 'delete' : 'edit',
            dataTarget: editAction.dataTarget,
            dateMatch: dateMatch,
            timeMatch: entry.time ? entry.time.substring(0, 5) : '',
            statusMatch: entry.status,
            time: targetTime,
            status: isMasuk ? '1' : '0'
          });
        }
      });
    });

    if (tasks.length === 0) {
      uiAdapter.alert('Tidak ada perubahan waktu yang terdeteksi.');
      return;
    }

    const nrp = current.summary.nrp;
    const startObj = parseHrisDate(uiAdapter.getValue('attendanceCheckStartDate'));
    const tahun = startObj ? startObj.getFullYear() : new Date().getFullYear();
    const bulan = startObj ? String(startObj.getMonth() + 1).padStart(2, '0') : String(new Date().getMonth() + 1).padStart(2, '0');
    const baseUrl = routeByNrp(nrp, ROUTES.ABSEN_BARCODE, ROUTES.ABSEN_BARCODE_OS)(tahun, bulan, nrp);

    createAutomationFlow('attendance-batch', window.location.href, {
      tasks,
      currentIndex: 0,
      baseUrl
    });

    UI.showResult('info', 'Otomasi Dimulai', `Memproses ${tasks.length} item...`);
    setTimeout(() => {
      window.location.href = baseUrl;
    }, 800);
  }

  async function resumeHadirPageLoop() {
    const flow = getAutomationFlow();
    if (!flow || flow.type !== 'attendance-batch') return;

    const tasks = flow.meta.tasks;
    const idx = flow.meta.currentIndex;
    const task = tasks[idx];

    if (!task) {
      finishAutomationFlow(flow.id);
      return;
    }

    if (!isBarcodePagePath()) {
      window.location.href = flow.meta.baseUrl;
      return;
    }

    UI.showGlobalLoader('Otomasi', `Mengolah baris ${idx + 1}/${tasks.length}...`);

    // Find modal - be more flexible with selectors
    const modal = document.querySelector('.modal.show, .modal.in, [role="dialog"].show') ||
      document.querySelector('.modal-content');
    const isVisible = modal && (modal.offsetWidth > 0 || modal.offsetHeight > 0);

    if (isVisible) {
      // Broad selectors for the date/time input
      const timeInput = modal.querySelector('input[name*="tanggal"], input[id*="tanggal"], .tanggal_edit, #tanggal_edit');
      const statusInput = modal.querySelector('select[name*="status"], select[id*="status"], #status_edit, .status_edit');

      if (timeInput) {
        const val = timeInput.value;
        // User's example: 2026-04-01 07:00:00.000
        // Screenshot shows: 01/04/2026 , 17:00
        // We try to keep the date prefix and only change the HH:mm
        let newVal = '';
        if (val.includes(',')) {
          // Format like "01/04/2026 , 17:00"
          newVal = val.split(',')[0] + ', ' + task.time;
        } else if (val.length >= 10) {
          // Format like "2026-04-01 ..." or "01-04-2026 ..."
          newVal = val.substring(0, 11) + task.time + (val.includes(':') && val.length > 16 ? val.substring(16) : ':00');
        } else {
          newVal = task.time; // Fallback
        }

        timeInput.value = newVal;
        // Trigger change event just in case
        timeInput.dispatchEvent(new Event('change', { bubbles: true }));

        if (statusInput && task.status) {
          statusInput.value = task.status;
          statusInput.dispatchEvent(new Event('change', { bubbles: true }));
        }

        setTimeout(() => {
          // Find Submit button by type, class, or text
          const buttons = Array.from(modal.querySelectorAll('button, input[type="submit"]'));
          const submitBtn = buttons.find(b =>
            b.type === 'submit' ||
            b.classList.contains('btn-primary') ||
            /submit|simpan|save/i.test(b.textContent)
          );

          if (submitBtn) {
            flow.meta.currentIndex++;
            sessionStorage.setItem(STORAGE.AUTO_FLOW, JSON.stringify(flow));
            submitBtn.click();
          } else {
            Logger.error('Tombol Submit tidak ditemukan di modal.');
            // Skip this one if we can't save
            flow.meta.currentIndex++;
            sessionStorage.setItem(STORAGE.AUTO_FLOW, JSON.stringify(flow));
            window.location.reload();
          }
        }, 400); // Speed up: 800ms -> 400ms
      } else {
        Logger.error('Input waktu tidak ditemukan di modal.');
        flow.meta.currentIndex++;
        sessionStorage.setItem(STORAGE.AUTO_FLOW, JSON.stringify(flow));
        window.location.reload();
      }
    } else {
      // Find the specific row by matching Date and Status
      let targetRow = null;

      const headers = Array.from(document.querySelectorAll('table thead th')).map(th => th.textContent.trim().toLowerCase());
      const tanggalIdx = headers.indexOf('tanggal');
      const statusIdx = headers.indexOf('status');

      if (tanggalIdx !== -1 && statusIdx !== -1) {
        const rows = document.querySelectorAll('table tbody tr');
        for (const row of rows) {
          const cells = row.querySelectorAll('td');
          if (cells.length <= Math.max(tanggalIdx, statusIdx)) continue;

          const rowTanggal = cells[tanggalIdx].textContent.trim();
          const rowStatus = cells[statusIdx].textContent.trim();

          if (rowTanggal.includes(task.dateMatch) && rowStatus === task.statusMatch) {
            targetRow = row;
            break;
          }
        }
      }

      if (!targetRow) {
        const rows = document.querySelectorAll('table tbody tr');
        targetRow = rows[task.index];
      }

      if (targetRow) {
        if (task.type === 'delete') {
          // Find delete button: btn-danger or text "Hapus/Delete"
          const buttons = Array.from(targetRow.querySelectorAll('button, a.btn'));
          const delBtn = buttons.find(b =>
            b.classList.contains('btn-danger') ||
            /hapus|delete|remove/i.test(b.textContent) ||
            b.querySelector('.fa-trash, .fa-times')
          );

          if (delBtn) {
            const oldConfirm = window.confirm;
            window.confirm = () => true;
            try {
              delBtn.click();
              // After delete, the page usually reloads or the row disappears.
              // We'll update the flow and wait for the next iteration.
              flow.meta.currentIndex++;
              sessionStorage.setItem(STORAGE.AUTO_FLOW, JSON.stringify(flow));
              // Small delay to let the click process before any potential reload
              setTimeout(() => {
                if (getAutomationFlow()) window.location.reload();
              }, 1000);
            } finally {
              setTimeout(() => { window.confirm = oldConfirm; }, 500);
            }
          } else {
            Logger.error('Tombol Delete tidak ditemukan di baris.');
            flow.meta.currentIndex++;
            sessionStorage.setItem(STORAGE.AUTO_FLOW, JSON.stringify(flow));
            window.location.reload();
          }
        } else {
          // EDIT ACTION
          let editBtn = targetRow.querySelector('.btn-warning, button[data-toggle="modal"], a[href*="edit"]');
          if (!editBtn && task.dataTarget) {
            editBtn = document.querySelector(`button[data-target="${task.dataTarget}"]`);
          }

          if (editBtn) {
            editBtn.click();
            setTimeout(resumeHadirPageLoop, 800);
          } else {
            Logger.error(`Gagal menemukan tombol edit untuk ${task.dateMatch}`);
            flow.meta.currentIndex++;
            sessionStorage.setItem(STORAGE.AUTO_FLOW, JSON.stringify(flow));
            window.location.reload();
          }
        }
      } else {
        Logger.error(`Gagal menemukan baris matching untuk ${task.dateMatch}`);
        flow.meta.currentIndex++;
        sessionStorage.setItem(STORAGE.AUTO_FLOW, JSON.stringify(flow));
        window.location.reload();
      }
    }
  }


  async function processHadirBatchEdit(input) {
    const indices = state.activeHadirBatchIndices || [];
    if (indices.length === 0) return;

    const current = state.attendanceCheck;
    if (!current || !current.summary) return;

    if (!input.jamMasuk && !input.jamKeluar && !input.status) {
      uiAdapter.alert('Tidak ada data yang diubah.');
      return;
    }

    const saveBtn = uiAdapter.get('#qm-btn-hadir-edit-save');
    if (saveBtn) saveBtn.disabled = true;
    closeHadirBatchEdit();

    // Start global progress bar
    UI.startProgress('Memproses Batch Edit Kehadiran...', indices.length);

    let successCount = 0;
    let failCount = 0;

    for (let i = 0; i < indices.length; i++) {
      const idx = indices[i];
      const entry = current.summary.entries[idx];
      if (!entry) { failCount++; continue; }

      const editAction = (entry.actions || []).find(act => /edit|ubah/i.test(act.text));
      if (!editAction || !editAction.href) { failCount++; continue; }

      try {
        const editUrl = toAbsoluteHrisUrl(editAction.href);
        const html = await hrisFetch(editUrl);
        const doc = parseHTML(html);
        const form = doc.querySelector('form');
        if (!form) throw new Error('Form edit tidak ditemukan.');

        const basePostData = new URLSearchParams();
        form.querySelectorAll('input, select, textarea').forEach(el => {
          if (!el.name || el.type === 'submit' || el.type === 'button') return;
          if ((el.type === 'checkbox' || el.type === 'radio') && !el.checked) return;
          basePostData.append(el.name, el.value);
        });

        // The date format is DD-MM-YYYY in entry.dateText
        const [d, m, y] = entry.dateText.split('-');
        const datePrefix = `${y}-${m}-${d}`;

        const submissions = [];

        // Intelligent selection of what to update based on the entry's current status
        if (entry.status === 'Masuk' && input.jamMasuk) {
          const pd = new URLSearchParams(basePostData.toString());
          pd.set('tanggal', `${datePrefix} ${input.jamMasuk}:00`);
          pd.set('status', input.status || '1');
          submissions.push(pd);
        } else if (entry.status === 'Keluar' && input.jamKeluar) {
          const pd = new URLSearchParams(basePostData.toString());
          pd.set('tanggal', `${datePrefix} ${input.jamKeluar}:00`);
          pd.set('status', input.status || '0');
          submissions.push(pd);
        } else if (input.status && !input.jamMasuk && !input.jamKeluar) {
          // If only status is changed, update current record
          const pd = new URLSearchParams(basePostData.toString());
          const timeStr = entry.time || '00:00:00';
          pd.set('tanggal', `${datePrefix} ${timeStr}`);
          pd.set('status', input.status);
          submissions.push(pd);
        }

        // Execute submissions for this date sequentially
        for (const postData of submissions) {
          const response = await fetch(form.action || editUrl, {
            method: 'POST',
            body: postData
          });
          if (!response.ok) throw new Error('HTTP ' + response.status);
        }

        successCount++;
      } catch (e) {
        Logger.error(`Batch edit failed for index ${idx}`, e);
        failCount++;
      }

      UI.updateProgress(i + 1, indices.length, `Selesai: ${i + 1}/${indices.length}`);
      // Small delay to be polite to the server
      await new Promise(r => setTimeout(r, 200));
    }

    UI.endProgress();
    if (saveBtn) saveBtn.disabled = false;

    uiAdapter.alert(`Proses selesai. Berhasil: ${successCount}, Gagal: ${failCount}`);

    // Refresh table
    KEHADIRAN.checkByRange(panelReaders.attendanceCheck());
  }

  function renderSpklCheckResult() {
    const wrap = uiAdapter.get('spklCheckResult');
    const btn = uiAdapter.get('spklCheckButton');
    if (!wrap) return;

    const current = state.spklCheck || createEmptySpklCheck();
    if (btn) {
      btn.disabled = current.loading;
      const originalText = btn.innerHTML;
      if (current.loading) {
        if (!btn.dataset.originalHtml) btn.dataset.originalHtml = originalText;
        btn.innerHTML = '<span class="qm-spinner qm-spinner-xs"></span> Checking...';
      } else if (btn.dataset.originalHtml) {
        btn.innerHTML = btn.dataset.originalHtml;
      }
    }

    if (current.loading) {
      uiAdapter.html('spklCheckResult', '<div class="qm-hadir-check-card"><div class="qm-flex qm-items-center qm-gap-s"><span class="qm-spinner"></span><span>Memuat data SPKL...</span></div></div>');
      return;
    }

    if (current.error) {
      uiAdapter.html('spklCheckResult', `<div class="qm-hadir-check-card qm-hadir-check-error">${escapeHtml(current.error)}</div>`);
      return;
    }

    if (!current.summary) {
      uiAdapter.html('spklCheckResult', '');
      return;
    }

    const summary = current.summary;
    const entries = summary.entries;
    const total = entries.length;

    const listHtml = total === 0
      ? '<div class="qm-hadir-check-detail-empty">Tidak ada data SPKL untuk bulan ini.</div>'
      : entries.map((entry, idx) => {
        const actionsHtml = (entry.actions || []).map(act => {
          const isEdit = /edit/i.test(act.text);
          const isDelete = /hapus|delete/i.test(act.text);
          const href = act.href ? toAbsoluteHrisUrl(act.href) : '#';

          if (isEdit) {
            // Edit handled by batch
            return '';
          }
          if (isDelete) {
            return `<button type="button" class="qm-btn-text qm-font-xs qm-text-danger qm-spkl-inline-delete-btn" data-url="${href}" style="margin-left: 8px;">Delete</button>`;
          }
          return `<a href="${href}" target="_blank" class="qm-btn-text qm-font-xs qm-text-primary" style="margin-left: 8px;">${escapeHtml(act.text)}</a>`;
        }).join('');

        const day = entry.dateText.split(/[-/]/)[0];

        return `
          <div class="qm-hadir-check-detail-item">
            <div class="qm-hadir-check-detail-top">
              <input type="checkbox" class="qm-spkl-batch-cb qm-mr-s" data-index="${idx}">
              <div class="qm-flex qm-items-center" style="min-width: 30px;">
                <span class="qm-font-semibold" style="font-size: var(--qm-font-s);">${escapeHtml(day)}</span>
              </div>
              <div class="qm-flex qm-flex-col qm-ml-s" style="min-width: 90px;">
                <span class="qm-font-mono" style="font-size: var(--qm-font-s); font-weight: 600;">${escapeHtml(entry.jamAwal)} - ${escapeHtml(entry.jamAkhir)}</span>
                <span class="qm-text-muted" style="font-size: 10px;">Jenis: ${escapeHtml(entry.otCode)}</span>
              </div>
              <div class="qm-flex qm-items-center qm-ml-m">
                <span class="qm-badge qm-badge-outline" style="font-size: 11px; padding: 1px 6px;">+${escapeHtml(entry.jamOt)} jam</span>
              </div>
              <div class="qm-flex-1 qm-ml-m">
                <span class="qm-badge ${entry.status.toLowerCase().includes('acc') || entry.status.toLowerCase().includes('ok') || entry.status.toLowerCase().includes('finish') ? 'ok' : 'err'}" style="font-size: 10px; padding: 1px 4px;">${escapeHtml(entry.status)}</span>
              </div>
              <div class="qm-hadir-check-actions">${actionsHtml}</div>
            </div>
            <div class="qm-spkl-inline-inputs qm-hidden" style="margin-top: 8px; padding-top: 8px; border-top: 1px dashed var(--qm-border-warm);">
              <div class="qm-flex qm-gap-s qm-mb-s">
                <div class="qm-flex-1">
                  <label class="qm-field-label" style="font-size: 9px; margin-bottom: 2px;">Jenis OT</label>
                  <select class="qm-select qm-spkl-inline-jenis" style="height: 24px; padding: 0 4px; font-size: 10px;">
                    <option value="" ${!entry.otCode ? 'selected' : ''}>Pilih Jenis</option>
                    <option value="1" ${entry.otCode === '1' ? 'selected' : ''}>OT BIASA</option>
                    <option value="2" ${entry.otCode === '2' ? 'selected' : ''}>LONG SHIFT</option>
                    <option value="3" ${entry.otCode === '3' ? 'selected' : ''}>NON STOP</option>
                    <option value="4" ${entry.otCode === '4' ? 'selected' : ''}>OT AWAL</option>
                    <option value="5A" ${entry.otCode === '5A' ? 'selected' : ''}>NO REST (AWAL)</option>
                    <option value="5B" ${entry.otCode === '5B' ? 'selected' : ''}>NO REST (TENGAH)</option>
                    <option value="5C" ${entry.otCode === '5C' ? 'selected' : ''}>NO REST (AKHIR)</option>
                    <option value="6" ${entry.otCode === '6' ? 'selected' : ''}>STANDBY</option>
                    <option value="7" ${entry.otCode === '7' ? 'selected' : ''}>LAIN-LAIN</option>
                    <option value="OT" ${entry.otCode === 'OT' ? 'selected' : ''}>OVERTIME</option>
                  </select>
                </div>
                <div class="qm-flex-1">
                  <label class="qm-field-label" style="font-size: 9px; margin-bottom: 2px;">Shift</label>
                  <select class="qm-select qm-spkl-inline-shift" style="height: 24px; padding: 0 4px; font-size: 10px;">
                    <option value="" ${!entry.shift ? 'selected' : ''}>Pilih Shift</option>
                    <option value="1" ${entry.shift === '1' ? 'selected' : ''}>SHIFT I</option>
                    <option value="2" ${entry.shift === '2' ? 'selected' : ''}>SHIFT II</option>
                    <option value="3" ${entry.shift === '3' ? 'selected' : ''}>SHIFT III</option>
                    <option value="4" ${entry.shift === '4' ? 'selected' : ''}>LONG SHIFT I</option>
                    <option value="5" ${entry.shift === '5' ? 'selected' : ''}>LONG SHIFT II</option>
                  </select>
                </div>
              </div>
              <div class="qm-flex qm-gap-s">
                <div class="qm-flex-1">
                  <label class="qm-field-label" style="font-size: 9px; margin-bottom: 2px;">Jam Awal</label>
                  <input type="time" class="qm-input qm-spkl-inline-awal" value="${entry.jamAwal.includes(':') ? entry.jamAwal.substring(0, 5) : ''}" style="height: 24px; padding: 2px 4px; font-size: 10px;">
                </div>
                <div class="qm-flex-1">
                  <label class="qm-field-label" style="font-size: 9px; margin-bottom: 2px;">Jam Akhir</label>
                  <input type="time" class="qm-input qm-spkl-inline-akhir" value="${entry.jamAkhir.includes(':') ? entry.jamAkhir.substring(0, 5) : ''}" style="height: 24px; padding: 2px 4px; font-size: 10px;">
                </div>
                <div class="qm-flex-1">
                  <label class="qm-field-label" style="font-size: 9px; margin-bottom: 2px;">Tambahan</label>
                  <input type="number" step="0.01" class="qm-input qm-spkl-inline-tambahan" value="${entry.jamOt}" style="height: 24px; padding: 2px 4px; font-size: 10px;">
                </div>
              </div>
            </div>
          </div>
        `;
      }).join('');

    uiAdapter.html('spklCheckResult', `
      <div class="qm-hadir-check-card" style="margin-top: 8px; padding: 10px 14px;">
        <div class="qm-hadir-check-detail">
          <div class="qm-flex qm-justify-between qm-items-center qm-mb-s" style="padding-bottom: 4px; border-bottom: 1px solid var(--qm-border-warm);">
            <div class="qm-flex qm-items-center">
              <input type="checkbox" id="qm-spkl-batch-cb-all" class="qm-mr-s">
              <label for="qm-spkl-batch-cb-all" class="qm-hadir-check-detail-title qm-m-0 qm-cursor-pointer" style="font-size: 11px; letter-spacing: 0.05em;">PILIH SEMUA SPKL</label>
            </div>
            <div class="qm-flex qm-items-center qm-gap-s">
              <button type="button" class="qm-btn qm-btn-primary qm-btn-sm" id="qm-btn-spkl-batch-edit" style="padding: 2px 8px; font-size: 10px;">Edit Terpilih</button>
              <a href="${spklListUrl(summary.nrp, uiAdapter.getValue('spklPageMonth') || (new Date().getMonth() + 1), new Date().getFullYear())}" target="_blank" class="qm-text-primary qm-font-xs" style="text-decoration: underline;">Buka Halaman Penuh</a>
            </div>
          </div>
          ${listHtml}
        </div>
      </div>
    `);
  }

  async function checkSPKLOnline(cells, runId) {
    const prof = startProfile('checkSPKLOnline', { cellCount: cells.length });
    const ctx = getPageContext();
    const bulan = String(ctx.bulan).padStart(2, '0');
    if (!ctx.nrp) {
      finishProfile(prof, { skipped: true });
      return;
    }

    const dates = cells.map(c => parseInt(c.tgl)).filter(n => !isNaN(n));
    if (!dates.length) {
      finishProfile(prof, { skipped: true });
      return;
    }
    const minDate = Math.min(...dates).toString().padStart(2, '0');
    const maxDate = Math.max(...dates).toString().padStart(2, '0');

    const spklUrl = buildSpklOnlineUrl(ctx, minDate, maxDate);

    try {
      const data = await hrisFetch(spklUrl);
      if (runId !== state.anomalyRunId) return;
      const doc = parseHTML(data);

      // Deteksi data nyata: ada tidaknya baris dengan minimal 3 kolom di tbody
      // (lebih andal dari pengecekan teks — "Belum ada data" bisa muncul di nav/notif)
      const allTableRows = doc.querySelectorAll('table tbody tr');
      let isNoData = true;
      allTableRows.forEach(tr => {
        if (tr.querySelectorAll('td').length >= 3) isNoData = false;
      });

      // Dynamic Header Detection for SPKL Online
      let tglColIdx = -1;
      const ths = doc.querySelectorAll('th');
      ths.forEach((th, i) => {
        const text = th.textContent.trim().toLowerCase();
        if (text.includes('tanggal') || text.includes('tgl')) tglColIdx = i;
      });

      const spklMap = {};
      Logger.debug('checkSPKLOnline: isNoData=', isNoData, 'tglColIdx=', tglColIdx, 'html preview:', data.slice(0, 500));
      if (!isNoData) {
        const rows = doc.querySelectorAll('table tbody tr');
        rows.forEach(tr => {
          const tds = tr.querySelectorAll('td');
          if (tds.length < 3) return;

          let rowTgl = '';
          if (tglColIdx !== -1 && tds[tglColIdx]) {
            rowTgl = tds[tglColIdx].textContent.trim();
          }
          if (!rowTgl) {
            // Fallback heuristics for date
            tds.forEach(td => {
              const val = td.textContent.trim();
              if (/^\d{4}[-\/]\d{2}[-\/]\d{2}/.test(val) || /^\d{2}[-\/]\d{2}[-\/]\d{4}/.test(val) || /^\d{2}[-\s]+[a-zA-Z]+[-\s]+\d{4}/.test(val)) {
                rowTgl = val;
              }
            });
          }

          let rowStatus = '';
          const badge = tr.querySelector('.badge');
          if (badge) {
            rowStatus = badge.textContent.trim().toUpperCase();
          } else {
            // Fallback status
            tds.forEach(td => {
              const val = td.textContent.trim().toUpperCase();
              if (val === 'APPROVED' || val === 'DISETUJUI' || val.includes('APPROVE') || val.includes('SETUJU') || val.includes('REJECT') || val.includes('TOLAK') || val.includes('DRAFT') || val.includes('MENUNGGU') || val.includes('ASK FOR APPROVAL')) {
                rowStatus = val;
              }
            });
          }

          if (rowTgl) {
            let dayNum = NaN;
            // Coba format ISO: YYYY-MM-DD atau YYYY/MM/DD (dengan atau tanpa leading zero)
            const isoM = rowTgl.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
            if (isoM) {
              dayNum = parseInt(isoM[3], 10);
            } else {
              // Coba format DD[-/]MM[-/]YYYY atau DD[-/]MM[-/]YY (hari di posisi pertama, 1-2 digit)
              const dmyM = rowTgl.match(/^(\d{1,2})[-\/](\d{1,2})[-\/]\d{2,4}/);
              if (dmyM) {
                dayNum = parseInt(dmyM[1], 10);
              } else {
                // Fallback: cari angka 1-2 digit yang berdiri sendiri (bukan bagian dari angka lebih panjang)
                const numM = rowTgl.match(/\b(\d{1,2})\b/);
                if (numM) dayNum = parseInt(numM[1], 10);
              }
            }
            if (!isNaN(dayNum) && dayNum >= 1 && dayNum <= 31) {
              Logger.debug(`row rowTgl: ${rowTgl} | dayNum: ${dayNum} | rowStatus: ${rowStatus}`);
              spklMap[dayNum] = rowStatus || 'APPROVED';
            }
          } else if (minDate === maxDate) {
            spklMap[parseInt(minDate)] = rowStatus || 'APPROVED';
          }
        });
      }
      Logger.debug('spklMap final:', JSON.stringify(spklMap));

      const processedDays = {};

      // Build DOM row index once (O(n)) — used for cell title/dot updates
      const domRowMap = {};
      const domRows = document.querySelectorAll('table tbody tr');
      for (let i = 0; i < domRows.length; i++) {
        const tds = domRows[i].querySelectorAll('td');
        if (tds.length >= 12) {
          const rowTgl = tds[COL.TGL]?.textContent.trim();
          const dNum = parseInt(rowTgl);
          if (!isNaN(dNum)) domRowMap[dNum] = tds;
        }
      }

      // Note: we can skip domRowMap entirely for dots because we have data-fix-date
      cells.forEach(item => {
        const dayNum = parseInt(item.tgl);
        let newMsg = '';
        const statusRaw = (spklMap[dayNum] || '').toUpperCase();
        Logger.debug(`cell tgl=${item.tgl} | dayNum=${dayNum} | statusRaw=${statusRaw}`);

        if (statusRaw) {
          if (statusRaw.includes('ASK FOR APPROVAL') || statusRaw.includes('MENUNGGU') || statusRaw.includes('DRAFT')) {
            newMsg = 'SPKL Belum Disetujui';
          } else if (statusRaw.includes('REJECTED') || statusRaw.includes('DITOLAK') || statusRaw.includes('DISAPPROVED')) {
            newMsg = 'SPKL Ditolak';
          } else if (statusRaw.includes('APPROVED') || statusRaw.includes('DISETUJUI')) {
            newMsg = 'SPKL Disetujui (Cek Jenis OT)';
          } else {
            // Status tidak dikenali tapi baris ditemukan → asumsikan approved, cek jenis OT
            newMsg = 'SPKL Disetujui (Cek Jenis OT)';
          }
        } else {
          newMsg = 'SPKL Online tidak ada';
        }

        let link = '';
        const isEntryFix = newMsg === 'SPKL Online tidak sesuai' || newMsg === 'SPKL Online tidak ada' || newMsg === 'SPKL Belum Disetujui';

        if (isEntryFix) {
          link = ROUTES.SPKL_ONLINE_SINGLE(item.fullDate, ctx.nrp);
        } else {
          const base = spklBaseUrl(ctx.nrp);
          const shiftParam = item.shift ? `&shift=${item.shift}` : '';
          link = `${base}?tahun=${ctx.tahun}&bulan=${bulan}&kode_bagian=&kode_seksi=&kode_group=&nrp=${ctx.nrp}&qm_auto_spkl_fix=1&full_date=${item.fullDate}${shiftParam}`;
        }

        if (newMsg && !processedDays[dayNum]) {
          state.anomalies.push({ tgl: item.tgl, fullDate: item.fullDate, col: 'SPKL Status', msg: newMsg, link: link });
          processedDays[dayNum] = true;
        }

        // Direct DOM update: find all dots for this date that are related to SPKL
        const dots = document.querySelectorAll(`.qm-fix-dot[data-fix-date="${item.tgl}"]`);
        dots.forEach(fixDot => {
          const currentTitle = fixDot.getAttribute('title') || '';
          if (currentTitle.includes('SPKL')) {
            let titleDot = 'Buka SPKL Online';
            if (newMsg === 'SPKL Belum Disetujui') titleDot = 'Buka SPKL Online (Menunggu Persetujuan)';
            else if (newMsg === 'SPKL Ditolak') titleDot = 'SPKL Ditolak, Cek Halaman SPKL';
            else if (newMsg === 'SPKL Disetujui (Cek Jenis OT)') titleDot = 'Cek Halaman SPKL';
            else if (isEntryFix) titleDot = 'Cek SPKL Online';
            else titleDot = 'Buka Halaman SPKL';

            fixDot.setAttribute('title', titleDot);
            fixDot.setAttribute('data-fix-link', link);

            // Also update the parent cell title if possible
            const td = fixDot.closest('td');
            if (td) {
              const cellTitle = td.getAttribute('title') || '';
              if (!cellTitle.includes(newMsg) && newMsg !== '') {
                td.setAttribute('title', cellTitle + (cellTitle ? ' | ' : '') + newMsg);
              }
            }
          }
        });
      });
      if (runId === state.anomalyRunId) renderAnomalies();
    } catch (e) {
      Logger.warn('Gagal mengambil data SPKL Online.', e);
    } finally {
      finishProfile(prof, { cellCount: cells.length, anomalyCount: state.anomalies.length });
      decrementPendingChecks(runId);
    }
  }

  async function checkBarcodeMangkir(absentDates, runId) {
    const prof = startProfile('checkBarcodeMangkir', { absentCount: absentDates.length });
    const ctx = getPageContext();
    if (!ctx.nrp) {
      finishProfile(prof, { skipped: true });
      return;
    }

    const barcodeUrl = buildKehadiranLink(ctx);

    try {
      const data = await hrisFetch(barcodeUrl);
      if (runId !== state.anomalyRunId) return;
      const doc = parseHTML(data);
      let dateColIdx = -1, statusColIdx = -1;

      const ths = doc.querySelectorAll('th');
      ths.forEach(function (th, i) {
        const text = th.textContent.trim().toLowerCase();
        if (text.includes('tanggal')) dateColIdx = i;
        if (text.includes('status')) statusColIdx = i;
      });

      const barcodeData = {};

      const rows = doc.querySelectorAll('table tbody tr');
      rows.forEach(function (row) {
        const tds = row.querySelectorAll('td');
        if (tds.length < 2) return;
        let tglText = '', statusText = '';

        if (dateColIdx !== -1 && statusColIdx !== -1) {
          tglText = tds[dateColIdx] ? tds[dateColIdx].textContent.trim() : '';
          statusText = tds[statusColIdx] ? tds[statusColIdx].textContent.trim() : '';
        } else {
          tds.forEach(function (td) {
            const val = td.textContent.trim();
            if (/masuk|keluar/i.test(val)) statusText = val;
            if (/^\d{2}[-\/]\d{2}[-\/]\d{4}/.test(val) || /^\d{2}$/.test(val)) tglText = val;
          });
        }

        if (tglText && statusText) {
          const tglStr = tglText.split(/\s+/)[0];
          let day = '';
          const m1 = tglStr.match(/^(\d{2})[-\/]/);
          const m2 = tglStr.match(/[-\/](\d{2})$/);
          const m3 = tglStr.match(/^(\d{1,2})$/);
          if (m1) day = m1[1]; else if (m2) day = m2[1]; else if (m3) day = m3[1].padStart(2, '0');
          if (day) {
            if (!barcodeData[day]) barcodeData[day] = [];
            barcodeData[day].push(statusText.toLowerCase());
          }
        }
      });

      absentDates.forEach(item => {
        const day = item.date.toString().padStart(2, '0');
        const statuses = barcodeData[day] || [];
        const hasMasuk = statuses.some(s => s.includes('masuk'));
        const hasKeluar = statuses.some(s => s.includes('keluar'));
        let errMessage = '';
        if (!hasMasuk && !hasKeluar) errMessage = 'Kedua jam masuk dan keluar kosong';
        else if (!hasMasuk) errMessage = 'Jam Masuk Kosong';
        else if (!hasKeluar) errMessage = 'Jam Keluar Kosong';

        if (errMessage) {
          const tds = item.tr.querySelectorAll('td');
          if (tds.length <= 3) return;
          const ketTd = tds[3];

          const existingTitle = ketTd.getAttribute('title') || '';
          const barcodeTitle = `Validasi Barcode: ${errMessage}`;
          ketTd.classList.add('qm-anomaly-cell');
          ketTd.setAttribute('title', existingTitle ? existingTitle + ' | ' + barcodeTitle : barcodeTitle);
          state.anomalies.push({ tgl: item.date, colIndex: 3, msg: 'Validasi Barcode: ' + errMessage, link: barcodeUrl });
        }
      });
      if (runId === state.anomalyRunId) {
        applyMark(document, state.anomalies);
        renderAnomalies();
      }
    } catch (e) {
      Logger.warn('Gagal mengambil data barcode.');
    } finally {
      finishProfile(prof, { absentCount: absentDates.length, anomalyCount: state.anomalies.length });
      decrementPendingChecks(runId);
    }
  }

  function spklHighlight() {
    if (!isSpklPagePath()) return;
    const hlDate = sessionStorage.getItem('qm_highlight_spkl_date');
    if (!hlDate) return;
    sessionStorage.removeItem('qm_highlight_spkl_date');
    setTimeout(() => {
      let hasScrolled = false;
      const rows = document.querySelectorAll('table tbody tr');
      rows.forEach(function (row) {
        let matched = false;
        const tds = row.querySelectorAll('td');
        tds.forEach(function (td) {
          const val = td.textContent.trim();
          const m = val.match(/^(\d{2})[-\/]\d{2}[-\/]\d{4}/) || val.match(/^(\d{4})[-\/]\d{2}[-\/](\d{2})/);
          if (m) {
            const d = m[1].length === 2 ? m[1] : m[2];
            if (d === hlDate.padStart(2, '0')) matched = true;
          }
        });
        if (matched) {
          row.classList.add('qm-row-highlight');
          tds.forEach(td => td.classList.add('qm-row-highlight'));
          if (!hasScrolled) {
            const y = row.getBoundingClientRect().top + window.scrollY - 150;
            window.scrollTo({ top: y, behavior: 'smooth' });
            hasScrolled = true;
          }
        }
      });
    }, 500);
  }

  /* ============================================================
   * 10. FORM AUTOMATION
   * ============================================================ */

  /** Unified Fix/Perbaikan click handler — shared by handleFixDotClick and handleBatchFixClick. */
  function handleFixClick(link, date, title, fullDate) {
    if (link) {
      const currentUrl = window.location.href;
      Logger.info('handleFixClick triggered', { link, date, title, currentUrl });
      // Only set return URL if we're not already in an auto-fix flow
      if (!currentUrl.includes('qm_auto_spkl_fix=1') && !currentUrl.includes('qm_auto_distribusi=1')) {
        Logger.info('Setting RETURN_URL to', currentUrl);
        sessionStorage.setItem(STORAGE.RETURN_URL, currentUrl);
      }
    }
    if (title === 'Buka Halaman Kehadiran' || title === 'Cek Kehadiran' || title === 'Lihat Duplikasi Shift') {
      const ctx = getPageContext();
      if (ctx.nrp && date) {
        sessionStorage.setItem(STORAGE.AUTO_NRP_FILL, ctx.nrp);
        sessionStorage.setItem(STORAGE.AUTO_DATE_FILL, date);
      }
      if (isBarcodePagePath() && title !== 'Lihat Duplikasi Shift') {
        const btn = document.querySelector('[data-target="#addData"]');
        if (btn) { btn.click(); return; }
      }
      if (link) {
        if (title !== 'Lihat Duplikasi Shift') sessionStorage.setItem(STORAGE.AUTO_ADD_DATA, 'true');
        window.open(link, '_blank');
      }
      return;
    }

    if (link) {
      if (link.includes('absenbarcode') && title && !title.includes('Pulang awal') && !title.includes('Duplikasi')) {
        sessionStorage.setItem(STORAGE.AUTO_ADD_DATA, 'true');
      }
      if (date) sessionStorage.setItem(STORAGE.HIGHLIGHT_SPKL, date);

      // Multi-step SPKL Fix logic
      if (link.includes('qm_auto_spkl_fix=1')) {
        // Collect all unique dates with SPKL fix links
        const spklFixAnomalies = state.anomalies.filter(a => a.link && a.link.includes('qm_auto_spkl_fix=1'));

        const queue = [];
        const seenDates = new Set();

        spklFixAnomalies.forEach(a => {
          // Extract full_date from link to ensure uniqueness across month/year if possible,
          // but usually tgl (day) is enough within a month
          const fullDateMatch = a.link.match(/full_date=([^&]+)/);
          const dateKey = fullDateMatch ? fullDateMatch[1] : a.tgl;

          if (!seenDates.has(dateKey)) {
            seenDates.add(dateKey);
            queue.push({
              link: a.link,
              date: a.tgl,
              fullDate: fullDateMatch ? fullDateMatch[1] : (a.fullDate || a.tgl),
              title: a.msg
            });
          }
        });

        if (queue.length > 0) {
          Logger.info('Starting background queue', queue);
          runSpklBackgroundQueue(queue).catch(e => Logger.error('SPKL queue error', e));
        }
        return;
      } else {
        window.open(link, '_blank');
      }
    } else {
      alert('Fitur perbaikan otomatis segera hadir!');
    }
  }


  /** Unified finally-block helper for pending checks state. */
  function decrementPendingChecks(runId) {
    if (runId !== state.anomalyRunId) return;
    state.pendingChecks--;
    if (state.pendingChecks <= 0) {
      const tab = document.querySelector('[data-pane="anomali"]');
      if (tab) tab.classList.remove('qm-tab-loading');
    }
  }

  /** Select dropdown value with MutationObserver support. */
  /** Select dropdown value with robust option polling (waits for AJAX options to load). */
  async function awaitDropdown(selector, nilaiTarget, callback, timeout = 5000) {
    if (!nilaiTarget) return callback();
    const startTime = Date.now();
    const target = nilaiTarget.toLowerCase();

    const poll = async () => {
      const select = document.querySelector(selector);
      if (select) {
        const options = select.querySelectorAll('option');
        let foundOpt = null;

        for (const opt of options) {
          if (opt.value.toLowerCase() === target || opt.textContent.trim().toLowerCase() === target) {
            foundOpt = opt;
            break;
          }
        }

        if (foundOpt) {
          select.value = foundOpt.value;
          if (typeof window.$ !== 'undefined' && window.$(select).selectpicker) {
            window.$(select).selectpicker('render').selectpicker('refresh');
          }
          select.dispatchEvent(new Event('change', { bubbles: true }));
          select.dispatchEvent(new Event('input', { bubbles: true }));
          setTimeout(callback, 400); // Small buffer for server-side triggers
          return;
        }
      }

      if (Date.now() - startTime < timeout) {
        setTimeout(poll, 150);
      } else {
        Logger.warn('Timeout waiting for option "' + nilaiTarget + '" in ' + selector);
        callback();
      }
    };

    poll();
  }

  /** Strips unsafe tags and attributes from HTML strings. */
  /** Fill the distribusi form and submit. */
  function fillDistribusiForm(dataKaryawan, nrp, tanggal, shift) {
    setTimeout(() => {
      UI.setGlobalProgress(20, 'Mengisi Jam Kerja...');

      const jkSelect = document.querySelector('select[name*="jam_kerja"], select[name*="jk"]');
      if (jkSelect && dataKaryawan.jk) {
        pickOption(jkSelect, opt => {
          const optVal = opt.value.trim();
          const textKode = opt.textContent.trim().split('-')[0].trim();
          return optVal === dataKaryawan.jk || textKode === dataKaryawan.jk;
        });
        refreshPicker(jkSelect);
      }

      UI.setGlobalProgress(35, 'Mengisi Periode...');
      const dateAwal = document.querySelectorAll('input[name*="tanggal_awal"], input[name*="periode_awal"], input[name*="tgl_awal"], input[name*="tgl_dari"], input[id*="tanggal_awal"], input[id*="tgl_awal"], input[name="start_date"]');
      const dateAkhir = document.querySelectorAll('input[name*="tanggal_akhir"], input[name*="periode_akhir"], input[name*="tgl_akhir"], input[name*="tgl_sampai"], input[name*="tgl_ke"], input[id*="tanggal_akhir"], input[id*="tgl_akhir"], input[name="end_date"]');

      const tglAwal = Array.isArray(tanggal) ? tanggal[0] : tanggal;
      const tglAkhir = Array.isArray(tanggal) ? (tanggal[1] || tglAwal) : tglAwal;

      dateAwal.forEach(input => setField(input, tglAwal));
      dateAkhir.forEach(input => setField(input, tglAkhir));

      UI.setGlobalProgress(50, 'Menyesuaikan Bagian...');
      awaitDropdown('select[name*="bagian"]', dataKaryawan.bag, () => {
        UI.setGlobalProgress(65, 'Menyesuaikan Seksi...');
        awaitDropdown('select[name*="seksi"]', dataKaryawan.sek, () => {
          UI.setGlobalProgress(75, 'Menyesuaikan Group & NRP...');
          const grpSelect = document.querySelector('select[name="kode_group"]');
          if (grpSelect && dataKaryawan.grp) {
            setField(grpSelect, dataKaryawan.grp);
            refreshPicker(grpSelect);
          }

          if (typeof nrp === 'object' && nrp.awal !== undefined) {
            const nrpIn1 = document.querySelector('input[list="nrp_awal"], input[name*="nrp_awal"], input[id*="nrp_initial"], input[name*="nrp_initial"], input[name*="nrp1"], input[name*="nrp_initial_text"]');
            const nrpIn2 = document.querySelector('input[list="nrp_akhir"], input[name*="nrp_akhir"], input[id*="nrp_final"], input[name*="nrp_final"], input[name*="nrp2"], input[name*="nrp_final_text"]');
            if (nrpIn1) setField(nrpIn1, nrp.awal, ['input', 'change', 'blur']);
            if (nrpIn2) setField(nrpIn2, nrp.akhir || nrp.awal, ['input', 'change', 'blur']);
          } else {
            let nrpInputs = document.querySelectorAll('input[list="nrp_awal"], input[list="nrp_akhir"], input[name*="nrp_awal"], input[name*="nrp_akhir"], input[name*="nrp1"], input[name*="nrp2"], input[name*="nrp_1"], input[name*="nrp_2"]');
            if (nrpInputs.length === 0) nrpInputs = document.querySelectorAll('input[name*="nrp"]');
            nrpInputs.forEach(input => setField(input, nrp, ['input', 'change', 'blur']));
          }

          if (shift) {
            const targetShiftRoman = shift === '1' ? 'I' : (shift === '2' ? 'II' : 'III');
            const expectedText = `${targetShiftRoman} - SHIFT ${targetShiftRoman}`;
            const shiftSelect = document.querySelector('select[name="kode_shift"], select[name="shift"]');
            if (shiftSelect) {
              pickOption(shiftSelect, opt => {
                const optText = opt.textContent.trim().toUpperCase();
                return opt.value === shift || optText === expectedText || optText === targetShiftRoman || optText === 'SHIFT ' + targetShiftRoman;
              });
              refreshPicker(shiftSelect);
            }
          }

          UI.setGlobalProgress(90, 'Validasi Form...');
          setTimeout(() => {
            // Final Validation: Ensure critical fields are NOT empty
            const requiredFields = [
              { sel: 'select[name*="jam_kerja"], select[name*="jk"]', label: 'Jam Kerja' },
              { sel: 'input[name*="tanggal_awal"], input[name*="periode_awal"], input[id*="tanggal_awal"]', label: 'Tanggal Awal' },
              { sel: 'select[name*="bagian"]', label: 'Bagian' },
              { sel: 'select[name*="seksi"]', label: 'Seksi' }
            ];

            const missing = requiredFields.filter(f => {
              const el = document.querySelector(f.sel);
              return el && (!el.value || el.value === '' || el.value === '0');
            });

            const nrpIn = document.querySelector('input[name*="nrp"], #nrp_initial_text');
            if (nrpIn && !nrpIn.value && typeof nrp !== 'object') missing.push({ label: 'NRP' });

            if (missing.length > 0) {
              const msg = 'Gagal: Field [' + missing.map(m => m.label).join(', ') + '] masih kosong.';
              UI.setGlobalProgress(100, msg);
              UI.showResult('danger', 'Validasi Gagal', msg);
              // Do not hide loader automatically so user can see what's missing
              return;
            }

            UI.setGlobalProgress(95, 'Mengirim permintaan...');
            const activeFlow = getAutomationFlow();
            if (activeFlow && (activeFlow.type === 'distribusi-jk' || activeFlow.type === 'distribusi-subsi')) {
              markAutomationFlowFinished(activeFlow.id);
            } else {
              sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
            }
            const buttons = document.querySelectorAll('button, input[type="submit"], .btn-primary');
            let submitBtn = null;
            for (const btn of buttons) {
              if ((btn.textContent && btn.textContent.includes('Start Distribusi')) || btn.value === 'Start Distribusi') {
                submitBtn = btn; break;
              }
            }
            if (submitBtn) submitBtn.click();
            else { const form = document.querySelector('form'); if (form) form.submit(); }
          }, 1200);
        });
      }, 6000); // 6s timeout for Bagian
    }, 500);
  }

  async function autoDistribusi() {
    const urlParams = getCurrentQueryParams();
    if (!urlParams.get('qm_auto_distribusi')) {
      const activeFlow = getAutomationFlow();
      const scopedDistribusiFlow = activeFlow && (activeFlow.type === 'distribusi-jk' || activeFlow.type === 'distribusi-subsi') ? activeFlow : null;
      // Check if we are on the result page of a distribution (no qm_auto param but AUTO_FINISHED is true)
      if ((scopedDistribusiFlow && scopedDistribusiFlow.finished) || (!activeFlow && sessionStorage.getItem(STORAGE.AUTO_FINISHED) === 'true')) {
        const pageText = document.body.textContent;
        const successAlert = document.querySelector('.alert-success, .alert-info');

        if (successAlert || pageText.includes('Distribution Process Completed')) {
          UI.showResult('success', 'Distribusi Selesai', 'Distribution Process Completed');
          const returnUrl = scopedDistribusiFlow?.returnUrl || sessionStorage.getItem(STORAGE.RETURN_URL);
          if (returnUrl) {
            setTimeout(() => {
              window.location.href = returnUrl;
            }, 1500);
          }
        }
      }
      return;
    }

    const cleanUrl = new URL(window.location);
    cleanUrl.searchParams.delete('qm_auto_distribusi');
    window.history.replaceState({}, document.title, cleanUrl.toString());

    const nrp = urlParams.get('nrp');
    const tanggalAwal = urlParams.get('tanggal') || urlParams.get('tanggal_awal');
    const tanggalAkhir = urlParams.get('tanggal_akhir') || tanggalAwal;
    const shift = urlParams.get('shift');
    const urlJk = urlParams.get('jk');
    Logger.info('Auto Distribusi started', { nrp, tanggalAwal, tanggalAkhir, shift, urlJk });
    if (!nrp || !tanggalAwal) return;

    UI.showGlobalLoader('Auto Distribusi', 'Mengambil Data...');
    try {
      const emp = await fetchEmployee(nrp);
      if (!emp.found) {
        UI.setGlobalProgress(100, 'Data karyawan tidak ditemukan.');
        UI.hideGlobalLoader(3000);
        return;
      }
      fillDistribusiForm({ jk: urlJk || emp.jk, bag: emp.bagian, sek: emp.seksi, grp: emp.group }, nrp, [tanggalAwal, tanggalAkhir], shift);
    } catch (e) {
      UI.setGlobalProgress(100, 'Gagal mengakses data karyawan.');
      UI.hideGlobalLoader(3000);
    }
  }

  async function autoDistribusiSubsi() {
    const urlParams = getCurrentQueryParams();
    if (!urlParams.get('qm_auto_distribusi_subsi')) return;

    const cleanUrl = new URL(window.location);
    cleanUrl.searchParams.delete('qm_auto_distribusi_subsi');
    window.history.replaceState({}, document.title, cleanUrl.toString());

    const jk = urlParams.get('jk');
    const tglAwal = urlParams.get('tglAwal');
    const tglAkhir = urlParams.get('tglAkhir');
    const shift = urlParams.get('shift');
    const bagian = urlParams.get('bagian');
    const seksi = urlParams.get('seksi');
    const grup = urlParams.get('grup');
    const nrp = urlParams.get('nrp');

    if (!jk || !tglAwal || !tglAkhir) return;

    UI.showGlobalLoader('Auto Distribusi Subsi', 'Mengisi data...');

    try {
      const isOS = nrp && nrp.length === 8;
      const nrpAwal = isOS ? '00000000' : '0000';
      const nrpAkhir = isOS ? '99999999' : '9999';

      fillDistribusiForm({ jk, bag: bagian, sek: seksi, grp: grup }, { awal: nrpAwal, akhir: nrpAkhir }, [tglAwal, tglAkhir], shift);
    } catch (e) {
      Logger.error('autoDistribusiSubsi error', e);
      UI.setGlobalProgress(100, 'Gagal: ' + e.message);
      UI.hideGlobalLoader(3000);
    }
  }

  function autoFillTargetPage() {
    if (!isAttendancePagePath()) return;
    const autoNrp = sessionStorage.getItem(STORAGE.AUTO_NRP);
    const autoBulan = sessionStorage.getItem(STORAGE.AUTO_BULAN);
    if (autoNrp && autoBulan) {
      setTimeout(() => {
        setField(document.querySelector('#bulan'), autoBulan);
        setField(document.querySelector('input[name="nrp"]'), autoNrp, ['input']);

        sessionStorage.removeItem(STORAGE.AUTO_NRP);
        sessionStorage.removeItem(STORAGE.AUTO_BULAN);
      }, TIMING.AUTO_FILL_DELAY);
    }
  }

  function autoFillSpklEdit() {
    if (!isActiveAutomationFlow(null, 'spkl-edit')) return;
    const data = JSON.parse(sessionStorage.getItem('qm_spkl_edit_pending'));
    if (!data) return;

    Logger.info('autoFillSpklEdit: Filling form', data);

    // Find fields using SELECTORS
    const selOt = document.querySelector(SELECTORS.SPKL_MODAL_OT_TYPE);
    const inMsk = document.querySelector(SELECTORS.SPKL_MODAL_MSK);
    const inKlr = document.querySelector(SELECTORS.SPKL_MODAL_KLR);
    const btnSave = document.querySelector(SELECTORS.SPKL_MODAL_SUBMIT);

    if (selOt && inMsk && inKlr) {
      // Small delay to ensure any internal scripts are ready
      setTimeout(() => {
        setField(selOt, data.ot);
        setField(inMsk, data.jamAwal);
        setField(inKlr, data.jamAkhir);

        const inJamOt = document.querySelector(SELECTORS.SPKL_MODAL_JAM_OT);
        if (inJamOt && data.jamOt) setField(inJamOt, data.jamOt);

        setTimeout(() => {
          if (btnSave) {
            Logger.info('autoFillSpklEdit: Clicking Save');
            markAutomationFlowFinished(getAutomationFlow().id);
            sessionStorage.removeItem('qm_spkl_edit_pending');
            btnSave.click();
          }
        }, 1000);
      }, 500);
    } else {
      Logger.warn('autoFillSpklEdit: Fields not found, retrying...');
      setTimeout(autoFillSpklEdit, 1000);
    }
  }

  function findBarcodeSearchTrigger(form) {
    const scope = form || document;
    const candidates = Array.from(scope.querySelectorAll('button, input[type="submit"], input[type="button"]'));
    return candidates.find(el => {
      const text = ((el.textContent || el.value || '') + '').trim().toLowerCase();
      const type = (el.getAttribute('type') || '').toLowerCase();
      return type === 'submit' || text === 'search' || text.includes('search') || text === 'cari' || text.includes('cari');
    }) || null;
  }

  function autoBarcodeSearchPage() {
    if (!isBarcodeListPagePath()) return;

    const saved = sessionStorage.getItem(STORAGE.AUTO_BARCODE_SEARCH);
    if (!saved) return;

    let payload = null;
    try {
      payload = JSON.parse(saved);
    } catch (e) {
      sessionStorage.removeItem(STORAGE.AUTO_BARCODE_SEARCH);
      Logger.warn('AUTO_BARCODE_SEARCH payload tidak valid.');
      return;
    }

    if (!payload?.nrp || !payload?.date) {
      sessionStorage.removeItem(STORAGE.AUTO_BARCODE_SEARCH);
      return;
    }

    const parts = String(payload.date).split('-');
    if (parts.length < 2) {
      sessionStorage.removeItem(STORAGE.AUTO_BARCODE_SEARCH);
      return;
    }

    const tahun = parts[0];
    const bulan = String(parseInt(parts[1], 10));
    let attempts = 0;
    const maxAttempts = 20;

    const timer = setInterval(() => {
      attempts++;

      const nrpInput = document.querySelector('input[name="nrp"], input[id*="nrp"], input[type="text"][maxlength="8"]');
      const monthSelect = document.querySelector('#bulan, select[name="bulan"]');
      const yearSelect = document.querySelector('select[name="tahun"], #tahun, input[name="tahun"]');
      const form = nrpInput?.closest('form') || document.querySelector('form[action*="absenbarcode"]') || document.querySelector('form');

      if (nrpInput && monthSelect) {
        if (monthSelect.tagName === 'SELECT') {
          pickOption(monthSelect, opt => opt.value === bulan || parseInt(opt.value, 10) === parseInt(bulan, 10));
          refreshPicker(monthSelect);
        } else {
          setField(monthSelect, bulan, ['change', 'input']);
        }
        if (yearSelect) {
          if (yearSelect.tagName === 'SELECT') {
            pickOption(yearSelect, opt => opt.value === tahun || opt.textContent.trim() === tahun);
            refreshPicker(yearSelect);
          } else {
            setField(yearSelect, tahun, ['change', 'input']);
          }
        }
        setField(nrpInput, payload.nrp, ['input', 'change']);

        const trigger = findBarcodeSearchTrigger(form);
        sessionStorage.removeItem(STORAGE.AUTO_BARCODE_SEARCH);
        clearInterval(timer);

        setTimeout(() => {
          if (trigger) trigger.click();
          else if (form && typeof form.requestSubmit === 'function') form.requestSubmit();
          else if (form) form.submit();
        }, 350);
        return;
      }

      if (attempts >= maxAttempts) {
        clearInterval(timer);
        sessionStorage.removeItem(STORAGE.AUTO_BARCODE_SEARCH);
        Logger.warn('Gagal menemukan elemen pencarian barcode otomatis.');
      }
    }, 250);
  }

  function autoClickAddData() {
    if (!isBarcodePagePath()) return;
    if (sessionStorage.getItem('qm_auto_add_data') === 'true') {
      sessionStorage.removeItem('qm_auto_add_data');
      setTimeout(() => {
        const btn = document.querySelector('[data-target="#addData"]');
        if (btn) btn.click();

        // Fill NRP and Date if available
        const nrpFill = sessionStorage.getItem(STORAGE.AUTO_NRP_FILL);
        const dateFill = sessionStorage.getItem(STORAGE.AUTO_DATE_FILL);
        if (nrpFill || dateFill) {
          setTimeout(() => {
            if (nrpFill) {
              const nrpInput = document.getElementById('nrp_input');
              if (nrpInput) {
                setField(nrpInput, nrpFill, ['input', 'change']);
              }
              sessionStorage.removeItem(STORAGE.AUTO_NRP_FILL);
            }
            if (dateFill) {
              const dateInput = document.getElementById('tanggal');
              if (dateInput) {
                const ctx = getPageContext();
                const fullDate = `${ctx.tahun}-${String(ctx.bulan).padStart(2, '0')}-${String(dateFill).padStart(2, '0')}`;
                setField(dateInput, fullDate, ['change']);
              }
              sessionStorage.removeItem(STORAGE.AUTO_DATE_FILL);
            }
          }, 600);
        }
      }, 500);
    }
  }

  function autoInputHadir() {
    if (!isBarcodeCreatePagePath()) return;

    const saved = sessionStorage.getItem(STORAGE.INPUT_HADIR);
    if (!saved) return;

    const data = JSON.parse(saved);
    sessionStorage.removeItem(STORAGE.INPUT_HADIR);

    UI.showGlobalLoader('Auto Input Kehadiran', 'Mengisi data...');

    setTimeout(async () => {
      try {
        const elTgl = document.getElementById('tanggal');
        const elNrp = document.getElementById('nrp_input');
        const elJam = document.getElementById('jam');
        const elStatus = document.getElementById('status');
        const btnTambah = document.getElementById('btnTambah');
        const btnSubmit = document.getElementById('submit');

        if (elTgl) setField(elTgl, data.tgl, ['change']);
        if (elNrp) setField(elNrp, data.nrp, ['input', 'change']);
        if (elJam) setField(elJam, data.jam, ['change']);
        if (elStatus) setField(elStatus, data.status, ['change']);

        await new Promise(r => setTimeout(r, 600));
        if (btnTambah) btnTambah.click();

        await new Promise(r => setTimeout(r, 800));
        if (btnSubmit) {
          const activeFlow = getAutomationFlow();
          if (activeFlow && activeFlow.type === 'hadir-single') {
            markAutomationFlowFinished(activeFlow.id);
          } else {
            sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
          }
          btnSubmit.click();
        }

        UI.hideGlobalLoader();
      } catch (e) {
        Logger.error('Auto Input Kehadiran Error', e);
        UI.hideGlobalLoader();
      }
    }, 1000);
  }


  async function processSpklBackgroundSingle(item) {
    Logger.info(`Fetching ${item.link}`);
    const html = await hrisFetch(item.link);
    const doc = parseHTML(html);

    const dateObj = parseHrisDate(item.fullDate);
    const dateStrId = dateObj ? dateObj.toLocaleDateString('id-ID') : item.fullDate;

    let tglColIdx = 0;
    const ths = doc.querySelectorAll('table th');
    ths.forEach((th, i) => {
      const text = th.textContent.trim().toLowerCase();
      if (text.includes('tanggal') || text === 'tgl') tglColIdx = i;
    });

    let editBtn = null;
    const rows = doc.querySelectorAll(SELECTORS.SPKL_TABLE_ROWS);
    for (const row of rows) {
      const tds = row.querySelectorAll('td');
      if (tds.length === 0) continue;

      const dateTd = tds[tglColIdx] || tds[0];
      const rowDateText = dateTd.textContent.trim();
      const rowDateObj = parseHrisDate(rowDateText);

      let isMatch = false;
      if (rowDateObj && dateObj) {
        isMatch = rowDateObj.getTime() === dateObj.getTime();
      } else {
        isMatch = Array.from(tds).some(td => {
          const v = td.textContent.trim();
          return v === rowDateText && parseHrisDate(v)?.getTime() === dateObj?.getTime();
        }) || rowDateText.includes(item.fullDate);
      }

      if (isMatch) {
        const lastTd = tds[tds.length - 1];
        const actionCandidates = Array.from(lastTd.querySelectorAll('button, a, .btn'));
        const allCandidates = Array.from(row.querySelectorAll(SELECTORS.SPKL_EDIT_BTN));

        const findEdit = (els) => els.find(el => {
          const txt = el.textContent.trim().toLowerCase();
          const target = (el.getAttribute('data-target') || '').toLowerCase();
          const href = (el.getAttribute('href') || '').toLowerCase();
          return txt.includes('edit') || target.includes('edit') || href.includes('edit');
        });

        editBtn = findEdit(actionCandidates) || findEdit(allCandidates);
        if (editBtn) break;
      }
    }

    if (!editBtn) {
      Logger.warn(`Baris tanggal ${dateStrId} tidak ditemukan.`);
      return false;
    }

    const modalId = editBtn.getAttribute('data-target');
    if (!modalId) return false;

    const modal = doc.getElementById(modalId.replace('#', ''));
    if (!modal) return false;

    const form = modal.querySelector('form');
    if (!form) return false;

    const mskInput = form.querySelector(SELECTORS.SPKL_MODAL_MSK);
    const otTypeSelect = form.querySelector(SELECTORS.SPKL_MODAL_OT_TYPE);
    if (!otTypeSelect) return false;

    const mskValue = mskInput ? mskInput.value : '';
    const mskTime = parseTimeToDecimal(mskValue);
    const urlParams = new URLSearchParams(item.link.split('?')[1] || '');
    const shift = urlParams.get('shift') || guessActualShift(mskTime, SHIFT_RULES);
    const currentOtType = otTypeSelect.value;

    let matchedRule = null;
    for (const rule of SPKL_RULES) {
      if (shift === rule.shift && currentOtType === rule.currentOtType) {
        matchedRule = rule;
        break;
      }
    }

    if (!matchedRule) {
      Logger.info(`No fix needed for ${dateStrId}: shift ${shift}, type ${currentOtType}`);
      return true; // Already correct
    }

    const params = new URLSearchParams();
    const headers = {
      'Content-Type': 'application/x-www-form-urlencoded',
      'X-Requested-With': 'XMLHttpRequest'
    };

    const csrfMeta = doc.querySelector('meta[name="csrf-token"], meta[name="csrf-test-name"], meta[name="_token"]');
    if (csrfMeta) {
      const token = csrfMeta.getAttribute('content');
      if (!token) {
        throw new Error('Token CSRF tidak ditemukan. Struktur halaman mungkin berubah.');
      }
      headers['X-CSRF-TOKEN'] = token;
      headers['X-XSRF-TOKEN'] = token;
    } else {
      const csrfHidden = form.querySelector('input[name="_token"], input[name="csrf_token"], input[name="csrf-token"]');
      if (csrfHidden && csrfHidden.value) {
        headers['X-CSRF-TOKEN'] = csrfHidden.value;
        headers['X-XSRF-TOKEN'] = csrfHidden.value;
      } else {
        throw new Error('Token CSRF tidak ditemukan. Struktur halaman mungkin berubah.');
      }
    }

    form.querySelectorAll('input, select, textarea').forEach(el => {
      if (!el.name || el.type === 'submit' || el.type === 'button') return;
      if ((el.type === 'checkbox' || el.type === 'radio') && !el.checked) return;

      let val = el.value;
      if (el === otTypeSelect) val = matchedRule.targetOtType;

      params.append(el.name, val);
    });

    const actionUrl = form.action.startsWith('http') ? form.action : ROUTES.BASE + (form.action.startsWith('/') ? '' : '/') + form.action;

    Logger.info(`Sending POST to ${actionUrl} for ${dateStrId}`);
    const res = await fetchWithTimeout(actionUrl, {
      method: 'POST',
      headers,
      body: params.toString()
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return true;
  }

  async function runSpklBackgroundQueue(queue) {
    UI.showGlobalLoader('Auto Fix SPKL', `Memulai perbaikan ${queue.length} data...`, true);
    let successCount = 0;

    for (let i = 0; i < queue.length; i++) {
      if (!uiAdapter.get('globalLoader')) {
        Logger.info('User cancelled background queue');
        return;
      }

      const item = queue[i];
      UI.setGlobalProgress(Math.round((i / queue.length) * 100), `Antrean ${i + 1}/${queue.length}: ${item.fullDate}...`);

      try {
        const result = await processSpklBackgroundSingle(item);
        if (result) successCount++;
        await new Promise(r => setTimeout(r, 800)); // Be nice to server
      } catch (e) {
        Logger.error(`Error processing ${item.fullDate}`, e);
      }
    }

    UI.setGlobalProgress(100, 'Selesai!');
    setTimeout(() => {
      UI.hideGlobalLoader();
      UI.showResult('success', 'Perbaikan SPKL', `Berhasil memproses ${successCount} dari ${queue.length} antrean.`);
      setTimeout(() => window.location.reload(), 1500); // Reload attendance table to update anomalies
    }, 1000);
  }


  async function runSpklBatchProcess() {
    const elNrp = document.getElementById("qm-fix-spkl-nrp");
    const elBulan = document.getElementById("qm-fix-spkl-bulan");
    const elTahun = document.getElementById("qm-fix-spkl-tahun");
    const elData = document.getElementById("qm-fix-spkl-data");

    const nrp = elNrp ? elNrp.value.trim() : "";
    const bulan = elBulan ? elBulan.value : "";
    const tahun = elTahun ? elTahun.value : "";
    const batchData = elData ? elData.value.trim() : "";

    if (!nrp || !bulan || !tahun || !batchData) {
      UI.showResult('warning', 'Data Belum Lengkap', 'Silakan isi NRP, Periode, dan Data Batch.');
      return;
    }

    createAutomationFlow('spkl-batch-single', window.location.href, { nrp, bulan, tahun });

    if (!/^\d{4}$|^\d{8}$/.test(nrp)) {
      UI.showResult('warning', 'NRP Tidak Valid', 'Gunakan 4 digit (Reguler) atau 8 digit (OS).');
      return;
    }

    let targetUrl = spklCreateUrl(nrp);

    // Save OT 7 details if any
    let jamAwal = "", jamAkhir = "", shiftVal = "";
    const elAwal = document.getElementById("qm-fix-spkl-jam-awal");
    const elAkhir = document.getElementById("qm-fix-spkl-jam-akhir");
    const elShift = document.getElementById("qm-fix-spkl-shift");
    jamAwal = elAwal ? elAwal.value : "";
    jamAkhir = elAkhir ? elAkhir.value : "";
    shiftVal = elShift ? elShift.value : "";

    if (!window.location.href.includes(targetUrl)) {
      sessionStorage.setItem("spkl_saved_data", JSON.stringify({ nrp, tahun, bulan, batchData, jamAwal, jamAkhir, shiftVal }));
      UI.showResult('success', 'Mengalihkan...', 'Halaman akan berpindah. Proses dilanjutkan otomatis.');
      setTimeout(() => { window.location.href = targetUrl; }, 1000);
      return;
    }

    _continueSpklBatch(nrp, tahun, bulan, batchData, jamAwal, jamAkhir, shiftVal);
  }

  async function _continueSpklBatch(nrp, tahun, bulan, batchData, jamAwal, jamAkhir, shiftVal) {
    const taskList = [];
    const items = batchData.split(',');
    for (let item of items) {
      item = item.trim();
      if (!item) continue;
      const parts = item.split(/[-:=]/);
      const hari = parts[0].trim();
      const jenisOt = parts.length > 1 ? parts[1].trim().toUpperCase() : "1";
      if (!isNaN(hari) && hari !== "") {
        taskList.push({ hari, jenisOt });
      }
    }

    if (taskList.length === 0) {
      UI.showResult('danger', 'Format Salah', 'Tidak ada data valid yang ditemukan.');
      return;
    }

    UI.showGlobalLoader('Proses SPKL Batch', 'Memulai...');

    // Hoist DOM lookups outside loop (constant per batch run)
    const nrpInput = document.getElementById("nrp_input");
    const tanggalInput = document.getElementById("tanggal");
    const jenisOtSelect = document.getElementById("jenis_ot");
    const btnTambah = document.getElementById("btnTambah");
    const jamAwalEl = document.querySelector("#jam_awal_ot");
    const jamAkhirEl = document.querySelector("#jam_akhir_ot");
    const shiftEl = document.querySelector("#shift");

    for (let i = 0; i < taskList.length; i++) {
      const task = taskList[i];
      const formatBulan = String(bulan).padStart(2, '0');
      const formatHari = String(task.hari).padStart(2, '0');
      const fullDate = `${tahun}-${formatBulan}-${formatHari}`;

      UI.setGlobalProgress((i / taskList.length) * 100, `Memproses Tgl ${task.hari}...`);

      setField(nrpInput, nrp, ['input', 'change']);
      setField(tanggalInput, fullDate, ['input', 'change']);
      if (jenisOtSelect) {
        setField(jenisOtSelect, task.jenisOt);
        if (window.jQuery && window.jQuery(jenisOtSelect).selectpicker) {
          window.jQuery(jenisOtSelect).selectpicker('refresh');
        }
      }

      if (task.jenisOt === "7") {
        setField(jamAwalEl, jamAwal);
        setField(jamAkhirEl, jamAkhir);
        setField(shiftEl, shiftVal);
      }

      await new Promise(r => setTimeout(r, TIMING.SPKL_INPUT_DELAY));
      if (btnTambah) btnTambah.click();
      await new Promise(r => setTimeout(r, TIMING.SPKL_CLICK_DELAY));
    }

    UI.setGlobalProgress(95, 'Menyimpan...');
    const btnSubmit = document.getElementById("submit");
    if (btnSubmit) {
      const activeFlow = getAutomationFlow();
      if (activeFlow && activeFlow.type === 'spkl-batch-single') {
        markAutomationFlowFinished(activeFlow.id);
      } else {
        sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
      }
      await new Promise(r => setTimeout(r, TIMING.SPKL_SUBMIT_DELAY));
      btnSubmit.click();
      UI.setGlobalProgress(100, 'Selesai, Master!');
      UI.hideGlobalLoader(2000);
    }
  }

  function checkSpklBatchResume() {
    // 1. Resume Per NRP (Batch Tanggal)
    const sessionData = sessionStorage.getItem("spkl_saved_data");
    if (sessionData) {
      const parsed = JSON.parse(sessionData);
      sessionStorage.removeItem("spkl_saved_data");

      setTimeout(() => {
        const elNrp = document.getElementById("qm-fix-spkl-nrp");
        const elBulan = document.getElementById("qm-fix-spkl-bulan");
        const elTahun = document.getElementById("qm-fix-spkl-tahun");
        const elData = document.getElementById("qm-fix-spkl-data");

        if (elNrp) elNrp.value = parsed.nrp;
        if (elBulan) elBulan.value = parsed.bulan;
        if (elTahun) elTahun.value = parsed.tahun;
        if (elData) elData.value = parsed.batchData;
        if (parsed.jamAwal) {
          const elA = document.getElementById("qm-fix-spkl-jam-awal");
          const elK = document.getElementById("qm-fix-spkl-jam-akhir");
          const elS = document.getElementById("qm-fix-spkl-shift");
          if (elA) elA.value = parsed.jamAwal;
          if (elK) elK.value = parsed.jamAkhir;
          if (elS) elS.value = parsed.shiftVal;
          const box = document.getElementById("qm-fix-spkl-ot7-box");
          if (box) box.classList.remove('qm-hidden');
        }

        // openPanel();
        // const tabFix = document.querySelector('[data-pane="fix"]');
        // if (tabFix) tabFix.click();

        UI.showResult('success', 'Melanjutkan...', 'Memulai proses batch otomatis.');
        _continueSpklBatch(parsed.nrp, parsed.tahun, parsed.bulan, parsed.batchData);
      }, 1200);
      return;
    }

    // 2. Resume Banyak NRP (Batch NRP)
    const MANY_NRP_KEY = "hris_spkl_ot_runner_v1";
    const st = JSON.parse(sessionStorage.getItem(MANY_NRP_KEY) || "null");
    if (st) {
      const pRoute = (function (s) {
        const current = spklAddPageKind();
        if (current && s.indexes[current] < s[current].length) return current;
        if (s.indexes.internal < s.internal.length) return "internal";
        if (s.indexes.outsource < s.outsource.length) return "outsource";
        return null;
      })(st);

      if (pRoute) {
        const current = spklAddPageKind();
        if (current === pRoute) {
          setTimeout(() => {
            // Populate fields for visibility
            const elNrps = document.getElementById("qm-fix-many-nrps");
            const elDate = document.getElementById("qm-fix-many-date");
            const elOt = document.getElementById("qm-fix-many-ot");
            if (elNrps) elNrps.value = [...st.internal, ...st.outsource].join(", ");
            if (elDate) elDate.value = st.date;
            if (elOt) {
              elOt.value = st.jenisOt;
              elOt.dispatchEvent(new Event('change', { bubbles: true }));
            }
            if (st.jenisOt === "7") {
              const elAwal = document.getElementById("qm-fix-many-jam-awal");
              const elAkhir = document.getElementById("qm-fix-many-jam-akhir");
              const elShift = document.getElementById("qm-fix-many-shift");
              if (elAwal) elAwal.value = st.jamAwal;
              if (elAkhir) elAkhir.value = st.jamAkhir;
              if (elShift) elShift.value = st.shiftVal;
            }

            // openPanel();
            // const tabFix = document.querySelector('[data-pane="fix"]');
            // if (tabFix) tabFix.click();

            UI.showGlobalLoader('Batch NRP', 'Melanjutkan...');
            _processManyNrpPage(pRoute, st);
          }, 1500);
        } else {
          Logger.info('Batch NRP pending redirect to ' + pRoute);
        }
      }
    }
  }

  async function runSpklManyNrpBatch() {
    const STORAGE_KEY = "hris_spkl_ot_runner_v1";

    const elNrps = document.getElementById("qm-fix-many-nrps");
    const elDate = document.getElementById("qm-fix-many-date");
    const elOt = document.getElementById("qm-fix-many-ot");

    const raw = elNrps ? elNrps.value.trim() : "";
    const dateVal = elDate ? elDate.value : "";
    const jO = elOt ? elOt.value : "";

    if (!raw || !dateVal || !jO) {
      UI.showResult('warning', 'Data Belum Lengkap', 'Silakan isi daftar NRP, Tanggal, dan Jenis OT.');
      return;
    }

    createAutomationFlow('spkl-batch-many', window.location.href, { date: dateVal, jenisOt: jO });

    let jamAwal = "", jamAkhir = "", shiftVal = "";
    if (jO === "7") {
      const elAwal = document.getElementById("qm-fix-many-jam-awal");
      const elAkhir = document.getElementById("qm-fix-many-jam-akhir");
      const elShift = document.getElementById("qm-fix-many-shift");
      jamAwal = elAwal ? elAwal.value : "";
      jamAkhir = elAkhir ? elAkhir.value : "";
      shiftVal = elShift ? elShift.value : "";
      if (!jamAwal || !jamAkhir || !shiftVal) {
        UI.showResult('warning', 'Detail OT 7 Kosong', 'Silakan isi Jam Awal, Akhir, dan Shift.');
        return;
      }
    }

    const nrps = raw.split(/[,\n\s]+/).map(v => v.trim()).filter(Boolean);
    const int = [], out = [], inv = [];
    for (const nrp of nrps) {
      if (/^\d{4}$/.test(nrp)) int.push(nrp);
      else if (/^\d{8}$/.test(nrp)) out.push(nrp);
      else inv.push(nrp);
    }

    if (inv.length) {
      UI.showResult('danger', 'NRP Tidak Valid', 'Ditemukan NRP salah: ' + inv.slice(0, 3).join(', ') + (inv.length > 3 ? '...' : ''));
      return;
    }

    const state = {
      date: dateVal,
      jenisOt: jO, jamAwal, jamAkhir, shiftVal,
      internal: int, outsource: out,
      indexes: { internal: 0, outsource: 0 }
    };

    sessionStorage.setItem(STORAGE_KEY, JSON.stringify(state));

    const current = spklAddPageKind();
    const target = (int.length > 0 ? "internal" : "outsource");

    if (current !== target) {
      UI.showResult('success', 'Mengalihkan...', 'Pindah ke halaman input ' + target + '.');
      setTimeout(() => { window.location.href = spklAddUrl(target); }, 1000);
    } else {
      _processManyNrpPage(target, state);
    }
  }

  async function _processManyNrpPage(route, s) {
    const STORAGE_KEY = "hris_spkl_ot_runner_v1";
    const tEl = document.querySelector("#tanggal");
    const nEl = document.querySelector("#nrp_input");
    const jEl = document.querySelector("#jenis_ot");
    const tbBtn = document.querySelector("#btnTambah");
    const sbBtn = document.querySelector("#submit");

    if (!tEl || !nEl || !jEl || !tbBtn || !sbBtn) {
      alert("Elemen form tidak lengkap. Pastikan berada di halaman input yang benar.");
      return;
    }

    const list = s[route];
    let idx = s.indexes[route];

    for (; idx < list.length; idx++) {
      const nrp = list[idx];
      UI.setGlobalProgress((idx / list.length) * 100, `Batch NRP: ${nrp}`);

      const setVal = (el, val) => setField(el, val);

      setVal(tEl, s.date);
      await new Promise(r => setTimeout(r, TIMING.SESSION_SAVE_DELAY));
      setVal(nEl, nrp);
      await new Promise(r => setTimeout(r, TIMING.SESSION_SAVE_DELAY2));
      setVal(jEl, s.jenisOt);
      await new Promise(r => setTimeout(r, TIMING.SESSION_SAVE_DELAY3));

      if (s.jenisOt === "7") {
        setVal(document.querySelector("#jam_awal_ot"), s.jamAwal);
        setVal(document.querySelector("#jam_akhir_ot"), s.jamAkhir);
        setVal(document.querySelector("#shift"), s.shiftVal);
        await new Promise(r => setTimeout(r, TIMING.SESSION_SAVE_DELAY2));
      }

      tbBtn.click();
      s.indexes[route] = idx + 1;
      sessionStorage.setItem(STORAGE_KEY, JSON.stringify(s));
      await new Promise(r => setTimeout(r, TIMING.SPKL_BATCH_NRP_CLICK_DELAY));
    }

    const nextR = (function (st) {
      if (st.indexes.internal < st.internal.length) return "internal";
      if (st.indexes.outsource < st.outsource.length) return "outsource";
      return null;
    })(s);

    if (!nextR) {
      sessionStorage.removeItem(STORAGE_KEY);
      UI.setGlobalProgress(100, 'Selesai!');
      const activeFlow = getAutomationFlow();
      if (activeFlow && activeFlow.type === 'spkl-batch-many') {
        markAutomationFlowFinished(activeFlow.id);
      } else {
        sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
      }
    } else {
      sessionStorage.setItem(STORAGE_KEY, JSON.stringify(s));
    }
    sbBtn.click();
  }

  async function runHadirManyNrpBatch() {
    const elNrps = document.getElementById("qm-input-hadir-many-nrps");
    const elDate = document.getElementById("qm-input-hadir-many-tanggal");
    const elJam = document.getElementById("qm-input-hadir-many-jam");
    const elStatus = document.getElementById("qm-input-hadir-many-status");

    const raw = elNrps ? elNrps.value.trim() : "";
    const dateVal = elDate ? elDate.value : "";
    const jamVal = elJam ? elJam.value : "";
    const statusVal = elStatus ? elStatus.value : "";

    if (!raw || !dateVal || !jamVal || statusVal === "") {
      UI.showResult('warning', 'Data Belum Lengkap', 'Silakan isi daftar NRP, Tanggal, Jam, dan Status.');
      return;
    }

    createAutomationFlow('hadir-batch-many', window.location.href, { date: dateVal, status: statusVal });

    const nrps = raw.split(/[,\n\s]+/).map(v => v.trim()).filter(Boolean);
    const int = [], out = [], inv = [];
    for (const nrp of nrps) {
      if (/^\d{4}$/.test(nrp)) int.push(nrp);
      else if (/^\d{8}$/.test(nrp)) out.push(nrp);
      else inv.push(nrp);
    }

    if (inv.length) {
      UI.showResult('danger', 'NRP Tidak Valid', 'Ditemukan NRP salah: ' + inv.slice(0, 3).join(', ') + (inv.length > 3 ? '...' : ''));
      return;
    }

    const batchState = {
      date: dateVal,
      jam: jamVal,
      status: statusVal,
      internal: int,
      outsource: out,
      indexes: { internal: 0, outsource: 0 }
    };

    sessionStorage.setItem(STORAGE.HADIR_BATCH, JSON.stringify(batchState));

    const current = absenCreatePageKind();
    const target = (int.length > 0 ? "internal" : "outsource");

    if (current !== target) {
      UI.showResult('success', 'Mengalihkan...', 'Pindah ke halaman input ' + target + '.');
      const targetUrl = getAbsenCreateUrlByKind(target);
      setTimeout(() => { window.location.href = targetUrl; }, 1000);
    } else {
      _processHadirManyNrpPage(target, batchState);
    }
  }

  async function _processHadirManyNrpPage(route, s) {
    const tEl = document.querySelector("#tanggal");
    const nEl = document.querySelector("#nrp_input");
    const jEl = document.querySelector("#jam");
    const sEl = document.querySelector("#status");
    const tbBtn = document.querySelector("#btnTambah");
    const sbBtn = document.querySelector("#submit");

    if (!tEl || !nEl || !jEl || !sEl || !tbBtn || !sbBtn) {
      alert("Elemen form tidak lengkap. Pastikan berada di halaman input yang benar.");
      return;
    }

    const list = s[route];
    let idx = s.indexes[route];

    UI.showGlobalLoader('Batch Kehadiran', 'Memulai...');

    for (; idx < list.length; idx++) {
      const nrp = list[idx];
      UI.setGlobalProgress((idx / list.length) * 100, `Batch Kehadiran: ${nrp}`);

      setField(tEl, s.date, ['change']);
      setField(nEl, nrp, ['input', 'change']);
      setField(jEl, s.jam, ['change']);
      setField(sEl, s.status, ['change']);

      await new Promise(r => setTimeout(r, 600));
      tbBtn.click();

      s.indexes[route] = idx + 1;
      sessionStorage.setItem(STORAGE.HADIR_BATCH, JSON.stringify(s));
      await new Promise(r => setTimeout(r, 1200));
    }

    const nextR = (function (st) {
      if (st.indexes.internal < st.internal.length) return "internal";
      if (st.indexes.outsource < st.outsource.length) return "outsource";
      return null;
    })(s);

    if (!nextR) {
      sessionStorage.removeItem(STORAGE.HADIR_BATCH);
      UI.setGlobalProgress(100, 'Selesai!');
      const activeFlow = getAutomationFlow();
      if (activeFlow && activeFlow.type === 'hadir-batch-many') {
        markAutomationFlowFinished(activeFlow.id);
      } else {
        sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
      }
    } else {
      sessionStorage.setItem(STORAGE.HADIR_BATCH, JSON.stringify(s));
      const targetUrl = getAbsenCreateUrlByKind(nextR);
      window.location.href = targetUrl;
      return;
    }
    sbBtn.click();
  }

  function checkHadirBatchResume() {
    const st = JSON.parse(sessionStorage.getItem(STORAGE.HADIR_BATCH) || "null");
    if (st) {
      const current = absenCreatePageKind();

      const pRoute = (function (s) {
        if (current && s.indexes[current] < s[current].length) return current;
        if (s.indexes.internal < s.internal.length) return "internal";
        if (s.indexes.outsource < s.outsource.length) return "outsource";
        return null;
      })(st);

      if (pRoute) {
        if (current === pRoute) {
          setTimeout(() => {
            _processHadirManyNrpPage(pRoute, st);
          }, 1500);
        } else {
          const targetUrl = getAbsenCreateUrlByKind(pRoute);
          window.location.href = targetUrl;
        }
      }
    }
  }

  async function runHadirBulanBatch() {
    const elNrp = document.getElementById("qm-input-hadir-bulan-nrp");
    const elBulan = document.getElementById("qm-input-hadir-bulan-bln");
    const elTahun = document.getElementById("qm-input-hadir-bulan-thn");
    const elHari = document.getElementById("qm-input-hadir-bulan-hari");
    const elMasuk = document.getElementById("qm-input-hadir-bulan-masuk");
    const elKeluar = document.getElementById("qm-input-hadir-bulan-keluar");

    const NRP = elNrp ? elNrp.value.trim() : "";
    const BULAN = elBulan ? parseInt(elBulan.value, 10) : 0;
    const TAHUN = elTahun ? parseInt(elTahun.value, 10) : 2026;
    const HARI_KERJA = elHari ? parseInt(elHari.value, 10) : 5;
    const jamMasuk = elMasuk ? elMasuk.value : "07:00";
    const jamKeluar = elKeluar ? elKeluar.value : "15:00";

    if (!NRP || !/^\d{4}$|^\d{8}$/.test(NRP)) {
      UI.showResult('warning', 'NRP Tidak Valid', 'Gunakan 4 digit (Reguler) atau 8 digit (OS).');
      return;
    }

    if (!BULAN || BULAN < 1 || BULAN > 12) {
      UI.showResult('warning', 'Bulan Tidak Valid', 'Silakan pilih bulan.');
      return;
    }

    UI.showGlobalLoader('Kalender Disiapkan', 'Menghitung hari kerja...');

    const liburBulanIni = LIBUR_NASIONAL_2026[BULAN] || [];
    const hariValid = [];
    const jumlahHariSeBulan = new Date(TAHUN, BULAN, 0).getDate();

    for (let tanggal = 1; tanggal <= jumlahHariSeBulan; tanggal++) {
      const dateObj = new Date(TAHUN, BULAN - 1, tanggal);
      const hari = dateObj.getDay(); // 0: Sunday, 6: Saturday
      const isWeekend = (HARI_KERJA === 5) ? (hari === 0 || hari === 6) : (hari === 0);

      if (!isWeekend && !liburBulanIni.includes(tanggal)) {
        hariValid.push(tanggal);
      }
    }

    if (hariValid.length === 0) {
      UI.showResult('warning', 'Tidak Ada Hari Kerja', 'Bulan ini tidak memiliki hari kerja valid.');
      UI.hideGlobalLoader();
      return;
    }

    const antrean = [];
    const bulanStr = String(BULAN).padStart(2, '0');
    hariValid.forEach(tgl => {
      const tglStr = String(tgl).padStart(2, '0');
      antrean.push({ waktu: `${TAHUN}-${bulanStr}-${tglStr}T${jamMasuk}`, status: "1", label: "Masuk" });
      antrean.push({ waktu: `${TAHUN}-${bulanStr}-${tglStr}T${jamKeluar}`, status: "0", label: "Keluar" });
    });

    sessionStorage.setItem('qm_auto_hadir_bulan_active', 'true');
    sessionStorage.setItem('qm_auto_hadir_bulan_nrp', NRP);
    sessionStorage.setItem('qm_auto_hadir_bulan_antrean', JSON.stringify(antrean));
    createAutomationFlow('hadir-bulan', window.location.href, { nrp: NRP, bulan: BULAN, tahun: TAHUN });

    const targetURL = absenAddUrl(NRP);

    UI.setGlobalProgress(100, 'Berhasil! Mengalihkan...');
    setTimeout(() => {
      window.location.href = targetURL;
    }, 1200);
  }

  async function checkHadirBulanResume() {
    if (sessionStorage.getItem('qm_auto_hadir_bulan_active') !== 'true') return;

    const NRP = sessionStorage.getItem('qm_auto_hadir_bulan_nrp');
    const antrean = JSON.parse(sessionStorage.getItem('qm_auto_hadir_bulan_antrean') || "[]");

    if (!NRP || antrean.length === 0) {
      sessionStorage.removeItem('qm_auto_hadir_bulan_active');
      return;
    }

    // Check if we are on the correct page
    const isAddPage = isBarcodeAddPagePath();
    if (!isAddPage) return;

    UI.showGlobalLoader('Automasi Berjalan', `Sisa data: ${antrean.length / 2} hari`);

    const inputNrp = document.getElementById('nrp_input');
    const inputTanggal = document.getElementById('tanggal');
    const inputStatus = document.getElementById('status');
    const btnTambah = document.getElementById('btnTambah');

    if (!inputNrp || !inputTanggal || !inputStatus || !btnTambah) {
      Logger.warn('Elemen form tidak ditemukan untuk resume automasi.');
      return;
    }

    // Start processing the queue
    for (let i = 0; i < antrean.length; i++) {
      const aksi = antrean[i];
      UI.setGlobalProgress((i / antrean.length) * 100, `Menyuntikkan: ${aksi.label} | ${aksi.waktu.split('T')[0]}`);

      setField(inputNrp, NRP, ['input', 'change']);
      setField(inputTanggal, aksi.waktu, ['input', 'change']);
      setField(inputStatus, aksi.status, ['change']);

      if (window.jQuery && window.jQuery(inputStatus).selectpicker) {
        window.jQuery(inputStatus).selectpicker('refresh');
      }

      await new Promise(r => setTimeout(r, 500));
      btnTambah.click();
      await new Promise(r => setTimeout(r, 1000));
    }

    sessionStorage.removeItem('qm_auto_hadir_bulan_active');
    sessionStorage.removeItem('qm_auto_hadir_bulan_nrp');
    sessionStorage.removeItem('qm_auto_hadir_bulan_antrean');

    UI.setGlobalProgress(100, 'Selesai! Menyimpan...');
    const btnSubmit = document.getElementById('submit');
    if (btnSubmit) {
      const activeFlow = getAutomationFlow();
      if (activeFlow && activeFlow.type === 'hadir-bulan') {
        markAutomationFlowFinished(activeFlow.id);
      } else {
        sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
      }
      setTimeout(() => btnSubmit.click(), 500);
    } else {
      UI.hideGlobalLoader(1500);
      alert(`[Quick Menu] Automasi Selesai!\n- NRP: ${NRP}\n- Total dieksekusi: ${antrean.length} baris.\nSilakan klik Simpan secara manual.`);
    }
  }

  /* ============================================================
   * 11. UI LAYER
   * ============================================================ */

  GM_addStyle(`

:root {
  --font-sans: "Plus Jakarta Sans", "Segoe UI", sans-serif;
  --color-page-start: #f3f0ea;
  --color-page-end: #ebdfcf;
  --color-surface: rgba(255, 255, 255, 0.95);
  --color-surface-strong: #ffffff;
  --color-surface-soft: #faf7f2;
  --color-border: rgba(75, 63, 47, 0.14);
  --color-border-strong: rgba(75, 63, 47, 0.2);
  --color-text: #3f3a33;
  --color-text-muted: #6f675d;
  --color-text-soft: #a59d93;
  --color-accent: #dfeef7;
  --color-accent-strong: #b7d7eb;
  --color-tag-red-bg: #fdeceb;
  --color-tag-red-text: #d16258;
  --color-tag-orange-bg: #fff1e6;
  --color-tag-orange-text: #dc8d42;
  --color-tag-violet-bg: #efedff;
  --color-tag-violet-text: #7567d8;
  --shadow-panel: 0 24px 64px rgba(52, 43, 33, 0.18), 0 4px 12px rgba(52, 43, 33, 0.08);
  --shadow-card: 0 1px 0 rgba(255, 255, 255, 0.85) inset, 0 0 0 1px rgba(75, 63, 47, 0.12);
  --radius-panel: 22px;
  --radius-card: 12px;
  --radius-pill: 10px;
  --color-tag-blue-bg: #e7f4ff;
  --color-tag-blue-text: #4a8dc2;
  --space-2: 8px;
  --space-3: 12px;
  --space-4: 16px;
  --space-5: 20px;
  --space-6: 24px;
  --space-7: 28px;
  --space-8: 32px;
  --space-10: 40px;
  --space-12: 48px;
  --panel-max-width: 820px;
  --transition-fast: 180ms ease;
}

#qa-shell {
  all: initial; /* Reset everything inside shell */
}

#qa-shell * {
  box-sizing: border-box;
  font-family: var(--font-sans);
}

#qa-shell button,
#qa-shell input {
  font: inherit;
  margin: 0;
}

#qa-shell button {
  border: 0;
  background: none;
  color: inherit;
  cursor: pointer;
  padding: 0;
}

.qa-wrapper {
  position: fixed;
  top: 0;
  left: 0;
  width: 100vw;
  height: 100vh;
  z-index: 999999;
  pointer-events: none;
  display: flex;
  align-items: center;
  justify-content: center;
}

.qa-wrapper.is-hidden {
  display: none !important;
}

.command-menu {
  width: min(90%, var(--panel-max-width));
  height: min(544px, calc(100vh - 80px));
  display: flex;
  flex-direction: column;
  background: var(--color-surface);
  backdrop-filter: blur(24px);
  border: 1px solid rgba(255, 255, 255, 0.8);
  border-radius: var(--radius-panel);
  box-shadow: var(--shadow-panel);
  overflow: hidden;
  transform-origin: center;
  pointer-events: auto;
  will-change: transform, opacity;
}

.search-bar,
.menu-footer {
  flex-shrink: 0;
  min-height: 62px;
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: 12px 20px;
}

.search-bar {
  border-bottom: 1px solid var(--color-border);
}

.search-input-wrap {
  flex: 1;
  min-width: 0;
  display: flex;
  align-items: center;
  gap: var(--space-3);
}

.search-icon,
.icon-button svg,
.result-icon svg,
.quick-action-icon svg,
.empty-state-icon svg,
.tab-chip svg {
  width: 22px;
  height: 22px;
  display: block;
}

.search-icon {
  color: var(--color-text);
  opacity: 0.8;
}

.search-input {
  width: 100%;
  padding: 0;
  border: 0 !important;
  outline: 0 !important;
  background: transparent !important;
  font-size: 1.1rem;
  line-height: 1.4;
  color: var(--color-text) !important;
  box-shadow: none !important;
}

.search-input::placeholder {
  color: var(--color-text-soft);
}

.icon-button {
  width: 36px;
  height: 36px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  color: var(--color-text-muted);
  border-radius: 10px;
  border: 1px solid transparent;
  transition: all var(--transition-fast);
}

.icon-button:hover {
  color: var(--color-text);
  background: rgba(255, 255, 255, 0.7);
}

.is-hidden {
  display: none !important;
}

.qm-hidden {
  display: none !important;
}

.menu-body {
  flex: 1;
  display: flex;
  min-height: 0;
  position: relative;
}

.panel-state {
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  overflow-y: auto;
  padding: 0;
  background: var(--color-surface);
  transition: opacity 200ms ease, transform 200ms ease;
}

@keyframes pop-out {
  from { opacity: 0; transform: scale(0.95); }
  to { opacity: 1; transform: scale(1); }
}

@keyframes pop-in {
  from { opacity: 1; transform: scale(1); }
  to { opacity: 0; transform: scale(0.95); }
}

.panel-state.is-entering { animation: pop-out 260ms cubic-bezier(0.2, 0, 0, 1) both; }
.panel-state.is-exiting { animation: pop-in 200ms cubic-bezier(0.4, 0, 1, 1) both; }

.detail-header {
  min-height: 62px;
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: 12px 20px;
  border-bottom: 1px solid var(--color-border);
}

.back-button {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  font-weight: 600;
  color: var(--color-text-muted);
  background: transparent;
  border: none;
  padding: 0;
  cursor: pointer;
}

.back-button svg {
  width: 18px;
  height: 18px;
}

.back-button:hover { color: var(--color-text); }
.detail-title {
  margin: 0;
  font-size: 1.05rem;
  font-weight: 600;
  color: var(--color-text);
  padding-left: 12px;
}

.empty-state {
  height: 100%;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  padding: 40px;
}

.empty-state-icon { margin-bottom: 20px; color: var(--color-text-muted); opacity: 0.6; }
.empty-state-title { margin: 0; font-size: 1.8rem; font-weight: 700; text-align: center; color: var(--color-text); }
.empty-state-subtitle { margin: 12px 0 0; font-size: 1rem; color: var(--color-text-muted); text-align: center; }

.quick-action-grid {
  width: 100%;
  max-width: 720px;
  margin-top: 32px;
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 12px;
}

.quick-action {
  min-height: 88px;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 10px;
  padding: 16px;
  background: var(--color-surface-strong);
  border-radius: var(--radius-card);
  box-shadow: var(--shadow-card);
  transition: all var(--transition-fast);
}

.quick-action:hover {
  transform: translateY(-2px);
  box-shadow: 0 12px 24px rgba(52, 43, 33, 0.12);
}

.quick-action-label { font-size: 0.95rem; font-weight: 600; color: var(--color-text); }

.results-panel { display: flex; flex-direction: column; padding: 20px; }
.results-group-title { margin: 0 0 12px; font-size: 0.9rem; font-weight: 700; color: var(--color-text-soft); text-transform: uppercase; letter-spacing: 0.05em; }
.results-list { display: flex; flex-direction: column; gap: 6px; }

.result-item {
  min-height: 52px;
  display: flex;
  align-items: center;
  gap: 14px;
  padding: 10px 16px;
  border-radius: 12px;
  background: transparent;
  transition: all var(--transition-fast);
}

.result-item:hover { background: rgba(255, 255, 255, 0.6); transform: translateX(4px); }
.result-content { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.result-label { font-size: 1rem; font-weight: 600; color: var(--color-text); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.result-tag { padding: 4px 10px; border-radius: 8px; font-size: 0.75rem; font-weight: 700; }

.tag-red { background: var(--color-tag-red-bg); color: var(--color-tag-red-text); }
.tag-orange { background: var(--color-tag-orange-bg); color: var(--color-tag-orange-text); }
.tag-violet { background: var(--color-tag-violet-bg); color: var(--color-tag-violet-text); }
.tag-blue { background: var(--color-tag-blue-bg); color: var(--color-tag-blue-text); }

.menu-footer {
  justify-content: space-between;
  border-top: 1px solid var(--color-border);
  background: rgba(255, 255, 255, 0.4);
}

.progress-bar-container {
  width: 100%;
  height: 44px;
  position: relative;
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 0 16px;
}

.progress-bar-track { 
  flex: 1;
  height: 28px;
  position: relative;
  background: var(--color-surface-soft);
  border-radius: 8px;
  border: 1px solid var(--color-border);
  overflow: hidden;
}

.progress-bar-fill { 
  height: 100%;
  background: var(--color-accent-strong);
  width: 0%;
  transition: width 0.2s ease-out;
}

.progress-status-text {
  position: absolute;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  font-size: 0.85rem;
  font-weight: 700;
  color: var(--color-text);
  white-space: nowrap;
  pointer-events: none;
  text-shadow: 0 0 4px rgba(255,255,255,0.4);
}

.qm-batch-table {
  width: 100%;
  border-collapse: separate;
  border-spacing: 0;
  border: 1px solid var(--color-border);
  border-radius: 12px;
  background: var(--color-surface);
  overflow: hidden;
  box-shadow: 0 1px 3px rgba(0,0,0,0.02);
  margin-top: 12px;
  transition: box-shadow 0.2s ease, border-color 0.2s ease;
}

.qm-batch-table:focus-within, #qm-batch-results:focus-within .qm-batch-table {
  outline: none;
}

.qm-batch-nrp-link {
  color: #1976D2;
  text-decoration: none;
  font-weight: 600;
  transition: color 0.2s ease;
}

.qm-batch-nrp-link:hover {
  text-decoration: underline;
  color: #1565C0;
}

.qm-batch-nrp-link:focus {
  outline: none !important;
  box-shadow: none !important;
}

#qm-batch-results:focus {
  outline: none;
}

#qm-batch-results {
  opacity: 0;
  transform: translateY(12px);
  will-change: transform, opacity;
}

#qm-batch-results.is-visible {
  animation: qm-fade-in-slide-up 0.5s cubic-bezier(0.16, 1, 0.3, 1) forwards !important;
}

@keyframes qm-fade-in-slide-up {
  from {
    opacity: 0;
    transform: translateY(16px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}

.qm-batch-fix-btn {
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  background: var(--color-tag-red-bg, #fdeceb) !important;
  color: var(--color-tag-red-text, #d16258) !important;
  border: 1px solid var(--color-tag-red-text, #d16258) !important;
  padding: 3px 10px !important;
  font-size: 11px !important;
  font-weight: 700 !important;
  border-radius: 12px !important;
  cursor: pointer !important;
  transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
  margin-left: 10px !important;
  vertical-align: middle !important;
  outline: none !important;
  box-shadow: 0 1px 2px rgba(0,0,0,0.02) !important;
  position: static !important;
}

.qm-batch-fix-btn:hover {
  background: var(--color-tag-red-text, #d16258) !important;
  color: #ffffff !important;
  box-shadow: 0 4px 8px rgba(209, 98, 88, 0.25) !important;
  transform: translateY(-1px) !important;
}

.qm-batch-fix-btn:active {
  transform: translateY(0px) !important;
}

.qm-table-header td {
  background: #F9F8F6 !important;
  text-transform: uppercase;
  font-size: 0.65rem !important;
  font-weight: 700 !important;
  letter-spacing: 0.08em;
  color: var(--color-text-muted);
  padding: 10px 16px !important;
  border-bottom: 1px solid var(--color-border);
}

.qm-batch-cell {
  padding: 12px 16px !important;
  border-bottom: 1px solid var(--color-border);
  vertical-align: middle;
}

.qm-batch-table tr:last-child td {
  border-bottom: none;
}

.qm-batch-group-header td, .qm-batch-seksi-header td {
  font-weight: 700;
  font-size: 0.9rem;
  background: var(--color-surface-soft);
  cursor: pointer;
  padding: 12px 16px !important;
}

/* Status Pills */
.status-pill {
  display: inline-flex;
  align-items: center;
  padding: 2px 10px;
  border-radius: 12px;
  font-size: 0.7rem;
  font-weight: 700;
  white-space: nowrap;
  text-transform: capitalize;
}

.status-active { background: #E3F2FD; color: #1976D2; }
.status-warning { background: #FFF3E0; color: #F57C00; }
.status-error { background: #FFEBEE; color: #D32F2F; }
.status-neutral { background: #F5F5F5; color: #616161; }

.qm-batch-item-row:hover td {
  background: var(--color-surface-soft);
}

.qm-table-row { display: table-row !important; }

.qm-accordion-chevron {
  transition: transform 0.2s ease;
  display: inline-flex;
  align-items: center;
  justify-content: center;
}
.expanded .qm-accordion-chevron { 
  transform: rotate(180deg); 
}

.qm-batch-group-header, .qm-batch-seksi-header {
  cursor: pointer;
  user-select: none;
}

.qm-batch-group-header:hover td, .qm-batch-seksi-header:hover td {
  background: rgba(122, 165, 194, 0.1) !important;
}

#qm-btn-batch-clear:hover {
  background: #FFEBEB !important;
  border-color: #EF9A9A !important;
}

#qm-btn-batch-clear:active {
  background: #FFCDD2 !important;
}

#qm-btn-batch-check:hover {
  background: #EEEEEE !important;
  opacity: 0.9;
}

.qm-batch-date-header {
  cursor: pointer;
  padding: 4px 0;
  user-select: none;
}
.qm-batch-date-header:hover {
  opacity: 0.8;
}
.qm-batch-date-content {
  padding-left: 12px;
  margin-top: 6px;
  border-left: 2px solid var(--color-border);
}

.qm-batch-col-lembur, .qm-batch-col-ket {
  padding-left: 8px !important;
  padding-right: 8px !important;
  width: 1%;
  white-space: pre-wrap;
}

.qm-batch-col-masalah {
  font-size: 12px !important;
}

.floating-menu-btn {
  width: 52px;
  height: 52px;
  border-radius: 50%;
  position: fixed;
  z-index: 999999;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--color-surface-strong);
  border: 1px solid var(--color-border);
  box-shadow: var(--shadow-panel);
  transition: all var(--transition-fast);
  padding: 0;
}

.floating-menu-btn:hover { transform: translateY(-2px) scale(1.05); box-shadow: 0 12px 32px rgba(52, 43, 33, 0.2); }
.floating-menu-btn svg { width: 26px; height: 26px; color: var(--color-text-muted); }

.button-is-hidden { opacity: 0; visibility: hidden; pointer-events: none; transform: scale(0.8); }

@keyframes menu-open-from-button {
  from { opacity: 0; transform: translate(var(--menu-origin-x), var(--menu-origin-y)) scale(0.2); }
  to { opacity: 1; transform: translate(0, 0) scale(1); }
}

@keyframes menu-close-to-button {
  from { opacity: 1; transform: translate(0, 0) scale(1); }
  to { opacity: 0; transform: translate(var(--menu-origin-x), var(--menu-origin-y)) scale(0.2); }
}

.is-opening { animation: menu-open-from-button 280ms cubic-bezier(0.2, 0.9, 0.2, 1) both; }
.is-closing { animation: menu-close-to-button 280ms cubic-bezier(0.4, 0, 0.2, 1) both; }

@media (max-width: 640px) {
  .quick-action-grid { grid-template-columns: repeat(2, 1fr); }
}

/* --- Special Features / Anomali --- */
.qm-anomaly-cell {
  border: 2px solid var(--color-tag-red-text, #d16258) !important;
  background-color: var(--color-tag-red-bg, #fdeceb) !important;
  position: relative !important;
}
.qm-fix-dot {
  position: absolute;
  top: 2px;
  right: 2px;
  width: 14px;
  height: 14px;
  background: linear-gradient(135deg, var(--color-tag-red-text, #d16258), #b34a41);
  border-radius: 50%;
  cursor: pointer;
  box-shadow: 0 2px 4px rgba(0,0,0,0.15), 0 0 0 1px rgba(0,0,0,0.05);
  border: 1.5px solid #ffffff;
  padding: 0;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
  z-index: 10;
}
.qm-fix-dot:hover {
  transform: scale(1.3) rotate(15deg);
  background: linear-gradient(135deg, var(--color-text, #3f3a33), #1a1612);
  box-shadow: 0 3px 6px rgba(0,0,0,0.25), 0 0 0 1.5px rgba(0,0,0,0.1);
}
.qm-row-highlight {
  background-color: var(--color-accent, #dfeef7) !important;
}

/* Premium Cari Karyawan Split Layout */
.qm-karyawan-split-layout {
  display: grid;
  grid-template-columns: 1.25fr 1fr;
  gap: 24px;
  width: 100%;
  align-items: stretch;
  height: 100%;
  min-height: 0;
  overflow: hidden;
}

@media (max-width: 768px) {
  .qm-karyawan-split-layout {
    grid-template-columns: 1fr;
    gap: 24px;
  }
}

/* Premium Forms & Input Elements */
.qm-form-group {
  display: flex;
  flex-direction: column;
  gap: 8px;
  margin-bottom: 20px;
}

.qm-form-label {
  font-size: 0.85rem;
  font-weight: 600;
  color: var(--color-text-muted);
}

.qm-premium-input {
  width: 100%;
  padding: 10px 14px;
  font-size: 0.95rem;
  background: var(--color-surface-strong, #ffffff) !important;
  border: 1px solid var(--color-border);
  border-radius: 10px;
  color: var(--color-text);
  outline: none;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.02);
  transition: border-color 0.2s, box-shadow 0.2s;
}

.qm-premium-input:focus {
  border-color: var(--color-accent-strong);
  box-shadow: 0 0 0 3px rgba(183, 215, 235, 0.3);
}

.qm-premium-select {
  appearance: none;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' fill='none' viewBox='0 0 24 24' stroke='%236f675d' stroke-width='2'%3E%3Cpath stroke-linecap='round' stroke-linejoin='round' d='M19 9l-7 7-7-7'%3E%3C/path%3E%3C/svg%3E");
  background-repeat: no-repeat;
  background-position: right 14px center;
  background-size: 16px;
  padding-right: 40px !important;
}

/* Premium Buttons & Button Rows */
.qm-button-row {
  display: flex;
  gap: 12px;
  margin-top: 24px;
}

.qm-premium-btn {
  padding: 10px 20px;
  font-size: 0.95rem;
  font-weight: 600;
  border-radius: 10px;
  cursor: pointer;
  border: none;
  transition: transform 0.15s, opacity 0.15s, box-shadow 0.15s;
  display: inline-flex;
  align-items: center;
  justify-content: center;
}

.qm-premium-btn:active {
  transform: scale(0.97);
}

.qm-premium-btn-primary {
  background: var(--color-accent-strong);
  color: #1a5276;
  border: 1px solid rgba(0, 0, 0, 0.04);
}

.qm-premium-btn-primary:hover {
  opacity: 0.92;
  box-shadow: 0 4px 12px rgba(183, 215, 235, 0.3);
}

.qm-premium-btn-secondary {
  background: #faf8f5;
  color: var(--color-text-muted);
  border: 1px solid var(--color-border);
}

.qm-premium-btn-secondary:hover {
  background: #f5f2ec;
}

/* Reusable Premium Buttons */
.qm-btn-premium-primary {
  padding: 10px 24px !important;
  background: var(--color-accent-strong, #b2d1e5) !important;
  color: #1a5276 !important;
  border: none !important;
  border-radius: 8px !important;
  font-weight: 600 !important;
  font-size: 0.95rem !important;
  cursor: pointer !important;
  transition: opacity 0.2s !important;
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  text-decoration: none !important;
}
.qm-btn-premium-primary:hover {
  opacity: 0.9 !important;
}
.qm-btn-premium-primary:active {
  opacity: 0.8 !important;
}

.qm-btn-premium-secondary {
  padding: 10px 24px !important;
  background: var(--color-surface-soft, #f9f8f6) !important;
  color: var(--color-text-muted, #555) !important;
  border: 1px solid var(--color-border) !important;
  border-radius: 8px !important;
  font-weight: 600 !important;
  font-size: 0.95rem !important;
  cursor: pointer !important;
  transition: opacity 0.2s !important;
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  text-decoration: none !important;
}
.qm-btn-premium-secondary:hover {
  background: #eee !important;
  opacity: 0.9 !important;
}
.qm-btn-premium-secondary:active {
  opacity: 0.8 !important;
}

.qm-btn-premium-danger {
  padding: 10px 24px !important;
  background: #FFF5F5 !important;
  color: #E53935 !important;
  border: 1px solid #FFCDD2 !important;
  border-radius: 8px !important;
  font-weight: 600 !important;
  font-size: 0.95rem !important;
  cursor: pointer !important;
  transition: all 0.2s ease !important;
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  text-decoration: none !important;
}
.qm-btn-premium-danger:hover {
  background: #FFEBEB !important;
  border-color: #EF9A9A !important;
}
.qm-btn-premium-danger:active {
  background: #FFCDD2 !important;
}

/* Premium Collapsible Group Headers */
.qm-preview-cards {
  background: var(--color-surface-soft, #f9f8f6) !important;
  border: 1px solid var(--color-border) !important;
  border-radius: 8px !important;
  padding: 12px 16px !important;
  font-weight: 600 !important;
  font-size: 0.95rem !important;
  color: var(--color-text) !important;
  cursor: pointer !important;
  display: flex !important;
  justify-content: space-between !important;
  align-items: center !important;
  margin-bottom: 8px !important;
  transition: background 0.2s, box-shadow 0.2s !important;
  box-shadow: 0 2px 4px rgba(0,0,0,0.02) !important;
}
.qm-preview-cards:hover {
  background: #f0ede8 !important;
  box-shadow: 0 4px 8px rgba(0,0,0,0.05) !important;
}

/* Scrollable Detail & Editor Panels */
.qm-karyawan-panel {
  
  padding-right: 8px !important;
}
/* Premium Live Preview Card */
.qm-premium-preview-card {
  background: linear-gradient(135deg, #ffffff 0%, #faf8f5 100%);
  border: 1px solid var(--color-border);
  border-radius: var(--radius-card);
  padding: 12px 16px;
  box-shadow: var(--shadow-card);
  display: flex;
  align-items: center;
  gap: 12px;
  position: relative;
  overflow: hidden;
}

.qm-preview-avatar {
  width: 48px;
  height: 48px;
  border-radius: 50%;
  background: #e1f5fe;
  color: #0288d1;
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: 700;
  font-size: 1.1rem;
  flex-shrink: 0;
}

.qm-preview-name {
  font-size: 1.05rem;
  font-weight: 600;
  color: var(--color-text);
  line-height: 1.3;
}

.qm-preview-subtitle {
  font-size: 0.85rem;
  color: var(--color-text-muted);
  margin-top: 3px;
}

.qm-preview-meta {
  font-size: 0.85rem;
  color: var(--color-text-soft);
  margin-top: 2px;
}

.qm-preview-tag {
  font-size: 0.75rem;
  padding: 2px 8px;
  border-radius: 12px;
  line-height: 1.2;
  font-weight: 600;
}

/* Premium Directory Items (Existing Team Members Style) */
.qm-directory-list {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.qm-directory-item {
  background: #ffffff;
  border: 1px solid var(--color-border);
  border-radius: var(--radius-card);
  padding: 12px 16px;
  display: flex;
  align-items: center;
  gap: 12px;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.01);
  cursor: pointer;
  transition: border-color 0.2s, background-color 0.2s;
}

.qm-directory-item:hover {
  border-color: var(--color-border-strong);
  background: #faf8f5;
}

.qm-directory-item.active {
  border-color: var(--color-accent-strong);
  background: #e7f4ff;
}

.qm-directory-avatar {
  width: 32px;
  height: 32px;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: 600;
  font-size: 0.85rem;
  flex-shrink: 0;
}

.qm-directory-name {
  font-weight: 600;
  color: var(--color-text);
  font-size: 0.9rem;
}

.qm-directory-subtitle {
  font-size: 0.8rem;
  color: var(--color-text-muted);
  margin-top: 1px;
}

.qm-section-header {
  font-size: 0.8rem;
  font-weight: 700;
  color: var(--color-text-muted);
  text-transform: uppercase;
  letter-spacing: 0.06em;
  margin: 0 0 12px 0;
}

`);

  /** Claude-style Spike Mark SVG (4-spoke radial). */
  const SPIKE_SVG = `<svg class="qm-spike-mark" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="4" stroke-linecap="round"><line x1="12" y1="4" x2="12" y2="20"/><line x1="4" y1="12" x2="20" y2="12"/><line x1="6.34" y1="6.34" x2="17.66" y2="17.66"/><line x1="6.34" y1="17.66" x2="17.66" y2="6.34"/></svg>`;

  const VIEW_CEK_NRP = `<div id="qm-pane-cek-nrp" class="qm-pane" style="padding: 24px 32px; display: flex; flex-direction: column; gap: 32px;">
    <!-- Single Check Group -->
    <div>
      <h3 style="margin: 0 0 16px 0; font-size: 1.05rem; font-weight: 700; color: var(--color-text);">Pencarian Tunggal</h3>
      
      <div style="display: flex; gap: 16px; margin-bottom: 16px;">
        <div style="flex: 1;">
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--color-text-muted); margin-bottom:8px;">Bulan</label>
          <select id="qm-input-bulan" style="width:100%; padding:10px 14px; border:1px solid var(--color-border); border-radius:8px; font-size:1rem; background:transparent; color:var(--color-text); outline:none; transition: border-color 0.2s;"></select>
        </div>
        <div style="flex: 1;">
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--color-text-muted); margin-bottom:8px;">Tahun</label>
          <select id="qm-input-tahun" style="width:100%; padding:10px 14px; border:1px solid var(--color-border); border-radius:8px; font-size:1rem; background:transparent; color:var(--color-text); outline:none; transition: border-color 0.2s;"></select>
        </div>
      </div>

      <div style="margin-bottom: 24px;">
        <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--color-text-muted); margin-bottom:8px;">NRP</label>
        <input id="qm-input-nrp" type="text" placeholder="Masukkan 4/8 digit NRP" maxlength="8" autocomplete="off" spellcheck="false" style="width:100%; padding:10px 14px; border:1px solid var(--color-border); border-radius:8px; font-size:1rem; background:transparent; color:var(--color-text); outline:none; transition: border-color 0.2s;">
      </div>

      <div style="display: flex; gap: 12px;">
        <button id="qm-btn-check" type="button" class="qm-btn-premium-primary">Cek NRP</button>
      </div>

      <div id="qm-result" style="display: none; margin-top: 16px;">
        <div id="qm-result-title"></div>
        <div id="qm-result-body"></div>
      </div>
      <div id="qm-history" style="margin-top: 16px;"></div>
    </div>

    <hr style="border: 0; border-top: 1px solid var(--color-border); margin: 0;">

    <!-- Batch Check Group -->
    <div>
      <h3 style="margin: 0 0 16px 0; font-size: 1.05rem; font-weight: 700; color: var(--color-text);">Pencarian Massal</h3>
      <div style="margin-bottom: 24px;">
        <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--color-text-muted); margin-bottom:8px;">Daftar NRP</label>
        <textarea id="qm-input-multi-nrp" placeholder="Pisahkan dengan enter atau koma" rows="4" style="width:100%; padding:10px 14px; border:1px solid var(--color-border); border-radius:8px; font-size:1rem; background:transparent; color:var(--color-text); outline:none; transition: border-color 0.2s; resize: vertical;"></textarea>
      </div>

      <div style="display: flex; gap: 12px;">
        <button id="qm-btn-batch-check" type="button" class="qm-btn-premium-secondary">Proses Batch</button>
        <button id="qm-btn-batch-clear" type="button" class="qm-btn-premium-danger">Hapus Batch</button>
      </div>

      <div id="qm-batch-results" style="margin-top: 16px; outline: none;" tabindex="-1"></div>
    </div>
  </div>`;
  const VIEW_CEK_KARY = `<div id="qm-pane-cek-kary" class="qm-pane" style="padding: 16px 24px 24px; display: flex; flex-direction: column; height: 100%; overflow: hidden; box-sizing: border-box;">
    <div class="qm-karyawan-split-layout">
      <!-- Left Column: Search Form & Details -->
      <div style="display: flex; flex-direction: column; gap: 12px; height: 100%; overflow-y: auto; padding-right: 8px; box-sizing: border-box; min-height: 0;">
        <div class="qm-card" style="padding: 16px 20px; display: flex; flex-direction: column; gap: 12px; flex-shrink: 0;">
          <h3 style="margin: 0; font-size: 1.1rem; font-weight: 600; color: var(--color-text);">Cari Data Personal</h3>
          
          <div class="qm-form-group" style="margin-bottom: 0;">
            <label class="qm-form-label" style="font-size: 0.8rem; margin-bottom: 4px;">NRP atau Nama Karyawan</label>
            <input id="qm-input-karyawan-search" type="text" class="qm-premium-input" placeholder="e.g. 80001234 atau Jane Cooper" autocomplete="off" spellcheck="false" style="padding: 8px 12px; font-size: 0.9rem;">
          </div>
          
          <div class="qm-button-row" style="gap: 8px;">
            <button id="qm-btn-karyawan-search" type="button" class="qm-btn-premium-primary" style="flex:1; padding: 8px 16px !important; font-size: 0.9rem !important;">Cari Karyawan</button>
            <button id="qm-btn-karyawan-reset" type="button" class="qm-btn-premium-secondary" style="flex:1; padding: 8px 16px !important; font-size: 0.9rem !important;">Reset</button>
          </div>
        </div>

        <div id="qm-karyawan-left-panel" style="padding: 0 4px;">
          <div style="font-size: 0.85rem; color: var(--color-text-muted); line-height: 1.5; padding: 16px; background: #faf8f5; border-radius: 8px; border: 1px dashed var(--color-border);">
            Pencarian NRP akan memeriksa database internal & outsource, lalu menampilkan profil detail & quick edit JK/KK.
          </div>
        </div>
      </div>
      
      <!-- Right Column: Results Directory -->
      <div style="display: flex; flex-direction: column; gap: 12px; height: 100%; overflow-y: auto; padding-right: 8px; box-sizing: border-box; min-height: 0;">
        <h3 id="qm-karyawan-directory-title" class="qm-section-header" style="margin: 0; font-size: 1.1rem; font-weight: 600; color: var(--color-text); padding-bottom: 4px; text-transform: none; letter-spacing: normal;">Hasil Pencarian</h3>
        <div id="qm-karyawan-results" class="qm-directory-list" style="display: flex; flex-direction: column; gap: 8px;"></div>
      </div>
    </div>
  </div>`;
  const VIEW_SPKL = `<div id="qm-pane-spkl" class="qm-pane">
        <div class="qm-pane-header">
          ${SPIKE_SVG}
          <h6 class="qm-serif">SPKL</h6>
        </div>
        <div class="qm-w-full">
          <div class="qm-card qm-mb-m">
            <h6 class="qm-section-title qm-mb-m qm-text-s">${SPIKE_SVG} Cek SPKL</h6>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">NRP</label>
                <input id="qm-spkl-page-nrp" type="text" placeholder="NRP" maxlength="8" autocomplete="off" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Start Date</label>
                <input id="qm-spkl-page-start-date" type="date" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">End Date</label>
                <input id="qm-spkl-page-end-date" type="date" class="qm-input">
              </div>
            </div>
            <button id="qm-btn-spkl-page-cek" type="button" class="qm-btn qm-btn-primary">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              Cek SPKL
            </button>
            <div id="qm-spkl-result"></div>
          </div>

          <div class="qm-card qm-mb-m">
            <h6 class="qm-section-title qm-mb-m qm-text-s">${SPIKE_SVG} SPKL Online</h6>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">NRP</label>
                <input id="qm-spkl-online-nrp" type="text" placeholder="NRP" maxlength="8" autocomplete="off" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Tanggal</label>
                <input id="qm-spkl-online-date" type="date" class="qm-input">
              </div>
            </div>
            <button id="qm-btn-spkl-online-cek" type="button" class="qm-btn qm-btn-primary">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              Cek SPKL Online
            </button>
          </div>

          <div class="qm-card qm-mb-m qm-info-box">
            <h6 class="qm-section-title qm-mb-s qm-text-s">${SPIKE_SVG} Referensi Kode OT</h6>
            <div style="font-size:11px; line-height:1.6">
              <b>1:</b> BIASA | <b>2:</b> LONG | <b>3:</b> NONSTOP | <b>4:</b> AWAL<br> <b>5A/B/C:</b> NOREST | <b>6:</b> STANDBY | <b>7:</b> LAIN | <b>OT:</b> OVERTIME
            </div>
          </div>

          <div class="qm-card qm-mb-m">
            <h6 class="qm-section-title qm-mb-m qm-text-s">${SPIKE_SVG} Per NRP</h6>
            <div class="qm-mb-m">
              <label class="qm-field-label">NRP</label>
              <input id="qm-fix-spkl-nrp" type="text" placeholder="NRP (4 atau 8 digit)" maxlength="8" autocomplete="off" class="qm-input">
            </div>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">Bulan</label>
                <select id="qm-fix-spkl-bulan" class="qm-select qm-font-semibold"></select>
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Tahun</label>
                <select id="qm-fix-spkl-tahun" class="qm-select qm-font-semibold"></select>
              </div>
            </div>
            <div class="qm-mb-m">
              <label class="qm-field-label">Data Tanggal-KodeOT</label>
              <textarea id="qm-fix-spkl-data" placeholder="Contoh: 2-1, 5-OT, 10-3" rows="2" class="qm-textarea qm-textarea-mono"></textarea>
            </div>
            <div id="qm-fix-spkl-ot7-box" class="qm-mb-m qm-ot7-box qm-hidden">
              <div class="qm-flex qm-gap-m qm-mb-m">
                <div class="qm-flex-1">
                  <label class="qm-field-label qm-field-label-normal">Jam Awal</label>
                  <input id="qm-fix-spkl-jam-awal" type="time" class="qm-input qm-field-time">
                </div>
                <div class="qm-flex-1">
                  <label class="qm-field-label qm-field-label-normal">Jam Akhir</label>
                  <input id="qm-fix-spkl-jam-akhir" type="time" class="qm-input qm-field-time">
                </div>
              </div>
              <div>
                <label class="qm-field-label qm-field-label-normal">Shift</label>
                <select id="qm-fix-spkl-shift" class="qm-select qm-field-shift">
                  <option value="1">SHIFT I</option>
                  <option value="2">SHIFT II</option>
                  <option value="3">SHIFT III</option>
                  <option value="4">LONG SHIFT I</option>
                  <option value="5">LONG SHIFT II</option>
                </select>
              </div>
            </div>
            <button id="qm-btn-spkl-batch" type="button" class="qm-btn qm-btn-primary">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/></svg>
              Proses Per NRP
            </button>
          </div>

          <div class="qm-card">
            <h6 class="qm-section-title qm-mb-m qm-text-s">${SPIKE_SVG} Banyak NRP</h6>
            <div class="qm-mb-m">
              <label class="qm-field-label">Daftar NRP</label>
              <textarea id="qm-fix-many-nrps" placeholder="Daftar NRP (pisahkan koma atau baris)" rows="2" class="qm-textarea qm-textarea-mono"></textarea>
            </div>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">Tanggal</label>
                <input id="qm-fix-many-date" type="date" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Jenis OT</label>
                <select id="qm-fix-many-ot" class="qm-select qm-font-semibold">
                  <option value="1">OT BIASA</option>
                  <option value="2">LONG SHIFT</option>
                  <option value="3">NON STOP</option>
                  <option value="4">OT AWAL</option>
                  <option value="5A">NO REST (AWAL)</option>
                  <option value="5B">NO REST (TENGAH)</option>
                  <option value="5C">NO REST (AKHIR)</option>
                  <option value="6">STANDBY</option>
                  <option value="7">LAIN-LAIN</option>
                  <option value="OT">OVERTIME</option>
                </select>
              </div>
            </div>
            <div id="qm-fix-many-ot7-box" class="qm-mb-m qm-ot7-box qm-hidden">
              <div class="qm-flex qm-gap-m qm-mb-m">
                <div class="qm-flex-1">
                  <label class="qm-field-label qm-field-label-normal">Jam Awal</label>
                  <input id="qm-fix-many-jam-awal" type="time" class="qm-input qm-field-time">
                </div>
                <div class="qm-flex-1">
                  <label class="qm-field-label qm-field-label-normal">Jam Akhir</label>
                  <input id="qm-fix-many-jam-akhir" type="time" class="qm-input qm-field-time">
                </div>
              </div>
              <div>
                <label class="qm-field-label qm-field-label-normal">Shift</label>
                <select id="qm-fix-many-shift" class="qm-select qm-field-shift">
                  <option value="1">SHIFT I</option>
                  <option value="2">SHIFT II</option>
                  <option value="3">SHIFT III</option>
                  <option value="4">LONG SHIFT I</option>
                  <option value="5">LONG SHIFT II</option>
                </select>
              </div>
            </div>
            <button id="qm-btn-spkl-many-nrp" type="button" class="qm-btn qm-btn-primary">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>
              Proses Banyak NRP
            </button>
          </div>
        </div>
      </div>`;
  const VIEW_KEHADIRAN = `<div id="qm-pane-kehadiran" class="qm-pane">
        <div class="qm-pane-header">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
          <h6>Kehadiran</h6>
        </div>
        <div class="qm-w-full">
          <div class="qm-card qm-mb-m">
            <h6 class="qm-section-title qm-mb-m qm-text-s">Check NRP</h6>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">NRP</label>
                <input id="qm-input-hadir-check-nrp" type="text" placeholder="NRP" maxlength="8" autocomplete="off" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Start date</label>
                <input id="qm-input-hadir-check-start-date" type="date" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">End date</label>
                <input id="qm-input-hadir-check-end-date" type="date" class="qm-input">
              </div>
            </div>
            <button id="qm-btn-hadir-check" type="button" class="qm-btn qm-btn-primary">Check</button>
            <div id="qm-hadir-check-result" class="qm-mt-m"></div>
          </div>

          <div class="qm-card qm-mb-m">
            <h6 class="qm-section-title qm-mb-m qm-text-s">👥 Per NRP</h6>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">NRP</label>
                <input id="qm-input-hadir-nrp" type="text" placeholder="NRP" maxlength="8" autocomplete="off" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Tanggal</label>
                <input id="qm-input-hadir-tanggal" type="date" class="qm-input">
              </div>
            </div>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">Jam</label>
                <input id="qm-input-hadir-jam" type="time" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Status</label>
                <select id="qm-input-hadir-status" class="qm-select">
                  <option value="">Pilih</option>
                  <option value="1">Masuk</option>
                  <option value="0">Keluar</option>
                </select>
              </div>
            </div>
            <button id="qm-btn-hadir-proses" type="button" class="qm-btn qm-btn-primary">Proses</button>
          </div>

          <!-- SECTION: Per NRP Satu Bulan -->
          <div class="qm-card qm-mb-m">
            <h6 class="qm-section-title qm-mb-m qm-text-s">📅 Per NRP satu bulan</h6>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">NRP</label>
                <input id="qm-input-hadir-bulan-nrp" type="text" placeholder="NRP" maxlength="8" autocomplete="off" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Bulan</label>
                <select id="qm-input-hadir-bulan-bln" class="qm-select">
                  <option value="1">Januari</option>
                  <option value="2">Februari</option>
                  <option value="3">Maret</option>
                  <option value="4">April</option>
                  <option value="5">Mei</option>
                  <option value="6">Juni</option>
                  <option value="7">Juli</option>
                  <option value="8">Agustus</option>
                  <option value="9">September</option>
                  <option value="10">Oktober</option>
                  <option value="11">November</option>
                  <option value="12">Desember</option>
                </select>
              </div>
            </div>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">Hari Kerja</label>
                <select id="qm-input-hadir-bulan-hari" class="qm-select">
                  <option value="5">5 Hari Kerja</option>
                  <option value="6">6 Hari Kerja</option>
                </select>
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Tahun</label>
                <input id="qm-input-hadir-bulan-thn" type="text" value="2026" class="qm-input" readonly>
              </div>
            </div>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">Jam Masuk</label>
                <input id="qm-input-hadir-bulan-masuk" type="time" value="07:00" class="qm-input qm-field-time">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Jam Keluar</label>
                <input id="qm-input-hadir-bulan-keluar" type="time" value="15:00" class="qm-input qm-field-time">
              </div>
            </div>
            <button id="qm-btn-hadir-bulan-proses" type="button" class="qm-btn qm-btn-primary">Mulai Automasi</button>
          </div>

          <div class="qm-card">
            <h6 class="qm-section-title qm-mb-m qm-text-s">👥 Banyak NRP</h6>
            <!-- NRP List -->
            <div class="qm-mb-m">
              <textarea id="qm-input-hadir-many-nrps" placeholder="Daftar NRP (pisahkan koma atau baris)" rows="2" class="qm-textarea qm-textarea-mono"></textarea>
            </div>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">Tanggal</label>
                <input id="qm-input-hadir-many-tanggal" type="date" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Jam</label>
                <input id="qm-input-hadir-many-jam" type="time" class="qm-input">
              </div>
            </div>
            <div class="qm-mb-m">
              <label class="qm-field-label">Status</label>
              <select id="qm-input-hadir-many-status" class="qm-select">
                <option value="">Pilih</option>
                <option value="1">Masuk</option>
                <option value="0">Keluar</option>
              </select>
            </div>
            <button id="qm-btn-hadir-many-proses" type="button" class="qm-btn qm-btn-primary">Proses Banyak NRP</button>
          </div>
        </div>
      </div>`;
  const VIEW_DISTRIBUSI = `<div id="qm-pane-distribusi" class="qm-pane">
        <div class="qm-pane-header">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="17 1 21 5 17 9"></polyline><path d="M3 11V9a4 4 0 0 1 4-4h14"></path><polyline points="7 23 3 19 7 15"></polyline><path d="M21 13v2a4 4 0 0 1-4 4H3"></path></svg>
          <h6>Distribusi</h6>
        </div>
        <div class="qm-w-full">
          <div class="qm-card qm-mb-m">
            <h6 class="qm-section-title qm-mb-m qm-text-s">🏢 Per Subsi</h6>

            <div id="qm-dist-subsi-jk-container" class="qm-mb-m">
              <div class="qm-flex qm-items-center qm-gap-s qm-text-muted qm-text-xs">
                <span class="qm-spinner qm-spinner-dark" style="width:12px;height:12px"></span>
                <span>Memuat opsi Jam Kerja...</span>
              </div>
            </div>

            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">Tanggal Awal</label>
                <input id="qm-input-distribusi-subsi-tgl-awal" type="date" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Tanggal Akhir</label>
                <input id="qm-input-distribusi-subsi-tgl-akhir" type="date" class="qm-input">
              </div>
            </div>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">Bagian</label>
                <select id="qm-input-distribusi-subsi-bagian" class="qm-select">
                   <option value="">Pilih Bagian</option>
                </select>
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Seksi</label>
                <select id="qm-input-distribusi-subsi-seksi" class="qm-select">
                   <option value="">Pilih Seksi</option>
                </select>
              </div>
            </div>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">Grup</label>
                <select id="qm-input-distribusi-subsi-grup" class="qm-select">
                   <option value="">Pilih Grup</option>
                </select>
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">SHIFT</label>
                <select id="qm-input-distribusi-subsi-shift" class="qm-select">
                  <option value="">Pilih Shift</option>
                </select>
              </div>
            </div>

            <div class="qm-flex qm-items-center qm-gap-s qm-mb-m">
              <input type="checkbox" id="qm-dist-subsi-use-distribusi" class="qm-config-checkbox" checked>
              <label for="qm-dist-subsi-use-distribusi" class="qm-field-label qm-m-0 qm-cursor-pointer qm-select-none qm-field-label-normal" style="font-size:12px; font-weight:600">On Background</label>
            </div>

            <button id="qm-btn-distribusi-subsi-proses" type="button" class="qm-btn qm-btn-primary">Proses Per Subsi</button>
          </div>

          <div class="qm-card qm-mb-m">
            <h6 class="qm-section-title qm-mb-m qm-text-s">🛠️ Per NRP</h6>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">NRP</label>
                <input id="qm-input-distribusi-nrp" type="text" placeholder="NRP" maxlength="8" autocomplete="off" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Jam Kerja</label>
                <div id="qm-dist-jk-options-container">
                  <div class="qm-flex qm-items-center qm-gap-s qm-text-muted qm-text-xs">
                    <span class="qm-spinner qm-spinner-dark" style="width:12px;height:12px"></span>
                    <span>Memuat opsi...</span>
                  </div>
                </div>
              </div>
            </div>

            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">Tanggal Awal</label>
                <input type="date" id="qm-dist-jk-target-date" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Tanggal Akhir</label>
                <input type="date" id="qm-dist-jk-target-date-end" class="qm-input">
              </div>
            </div>

            <div class="qm-mb-m">
              <label class="qm-field-label">Pilih Shift</label>
              <select id="qm-dist-jk-target-shift" class="qm-select">
                <option value="1">Shift 1 (Pagi)</option>
                <option value="2">Shift 2 (Siang)</option>
                <option value="3">Shift 3 (Malam)</option>
              </select>
            </div>

            <div class="qm-flex qm-items-center qm-gap-s qm-mb-m">
              <input type="checkbox" id="qm-dist-jk-use-distribusi" class="qm-config-checkbox" checked>
              <label for="qm-dist-jk-use-distribusi" class="qm-field-label qm-m-0 qm-cursor-pointer qm-select-none qm-field-label-normal" style="font-size:12px; font-weight:600">On Background</label>
            </div>

            <button id="qm-btn-distribusi-proses" type="button" class="qm-btn qm-btn-primary">Mulai Proses</button>
          </div>

          <div class="qm-card qm-mb-m">
            <h6 class="qm-section-title qm-mb-m qm-text-s">📅 Kalender Kerja</h6>
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">NRP</label>
                <input id="qm-dist-KK-nrp" type="text" placeholder="NRP" maxlength="8" autocomplete="off" class="qm-input">
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Periode</label>
                <input id="qm-dist-KK-date" type="month" class="qm-input">
              </div>
            </div>

            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">Bagian</label>
                <select id="qm-dist-KK-bagian" class="qm-select">
                  <option value="">Pilih Bagian</option>
                </select>
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Seksi</label>
                <select id="qm-dist-KK-seksi" class="qm-select">
                  <option value="">Pilih Seksi</option>
                </select>
              </div>
            </div>

            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <label class="qm-field-label">Grup</label>
                <select id="qm-dist-KK-grup" class="qm-select">
                  <option value="">Pilih Grup</option>
                </select>
              </div>
              <div class="qm-flex-1">
                <label class="qm-field-label">Kode Kalender Kerja</label>
                <div id="qm-dist-KK-options-container">
                  <div class="qm-flex qm-items-center qm-gap-s qm-text-muted qm-text-xs">
                    <span class="qm-spinner qm-spinner-dark" style="width:12px;height:12px"></span>
                    <span>Memuat opsi...</span>
                  </div>
                </div>
              </div>
            </div>

            <button id="qm-btn-KK-update" type="button" class="qm-btn qm-btn-primary">Update Kalender</button>
          </div>
        </div>
      </div>`;
  const VIEW_SETTINGS = `<div id="qm-pane-config" class="qm-pane">
        <div class="qm-mb-l">
          <label class="qm-field-label">Shortcut Akses Cepat</label>
          <div class="qm-flex qm-gap-m">
            <input id="qm-input-shortcut" type="text" value="Ctrl+Z" readonly class="qm-input qm-flex-1 qm-text-center qm-font-semibold qm-input-shortcut" />
            <button id="qm-btn-record-shortcut" class="qm-btn qm-btn-primary qm-btn-record">Ubah</button>
          </div>
          <small class="qm-text-muted qm-mt-s qm-block qm-lh-1-3">Klik Ubah, lalu tekan kombinasi tombol keyboard baru untuk menyimpannya.</small>
        </div>

        <div class="qm-mb-l">
          <label class="qm-field-label">Tema Tampilan</label>
          <div class="qm-flex qm-gap-m">
            <button id="qm-btn-theme-light" class="qm-theme-btn light">☀️ Terang</button>
            <button id="qm-btn-theme-dark" class="qm-theme-btn dark">🌙 Gelap</button>
          </div>
        </div>

        <div class="qm-flex qm-items-center qm-gap-s qm-mb-m">
          <input type="checkbox" id="qm-config-collapse-menu" class="qm-config-checkbox" ${state.alwaysCollapse ? 'checked' : ''}>
          <label for="qm-config-collapse-menu" class="qm-field-label qm-m-0 qm-cursor-pointer qm-select-none">Always collapsed sidebar menu</label>
        </div>

        <div class="qm-flex qm-items-center qm-gap-s qm-mb-l">
          <input type="checkbox" id="qm-config-debug-mode" class="qm-config-checkbox" ${state.debug ? 'checked' : ''}>
          <label for="qm-config-debug-mode" class="qm-field-label qm-m-0 qm-cursor-pointer qm-select-none">Debug Mode (Console Logs)</label>
        </div>

        <button id="qm-btn-show-logs" class="qm-btn qm-w-full">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="margin-right:8px"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
          <span>Lihat Log Aktivitas</span>
        </button>
        <div id="qm-log-container" class="qm-hidden qm-mt-m">
          <div class="qm-log-actions qm-flex qm-justify-between qm-mb-xs">
            <span class="qm-caption qm-text-muted qm-font-xs qm-uppercase" style="letter-spacing:1px">Activity Stream</span>
            <div class="qm-flex qm-gap-s">
              <button id="qm-btn-export-logs" class="qm-btn-text qm-font-xs" title="Export as Markdown (.md)">Export .md</button>
              <button id="qm-btn-clear-logs" class="qm-btn-text qm-font-xs qm-text-danger" title="Clear All Logs">Clear</button>
            </div>
          </div>
          <div id="qm-log-body" class="qm-log-body-inline">
            <div class="qm-text-muted qm-text-center qm-mt-xl">Belum ada log aktivitas.</div>
          </div>
        </div>
      </div>
      </div> <!-- End qm-content-area -->
      </div> <!-- End qm-panel-body -->`;

  const HTML = `
<div id="qa-shell">
  <button id="qa-floating-btn" class="floating-menu-btn button-is-hidden" aria-label="Open menu"></button>
  <div id="qa-menu-container" class="qa-wrapper is-hidden">
    <main class="command-menu" aria-label="Quick actions">
      <header class="search-bar">
        <div class="search-input-wrap">
          <span class="search-icon"></span>
          <input id="qa-search" class="search-input" type="search" placeholder="Search actions..." autocomplete="off" spellcheck="false" />
        </div>
        <button id="qa-search-clear" class="icon-button is-hidden" type="button" aria-label="Clear search">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M6 6L18 18M18 6L6 18"></path></svg>
        </button>
      </header>

      <header class="detail-header is-hidden">
        <button id="qa-back-btn" class="back-button" type="button">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><line x1="19" y1="12" x2="5" y2="12"></line><polyline points="12 19 5 12 12 5"></polyline></svg>
          <span>Back</span>
        </button>
        <h2 id="qa-detail-title" class="detail-title">Action</h2>
      </header>

      <section class="menu-body">
        <div id="qa-empty-state" class="panel-state"></div>
        <div id="qa-results-state" class="panel-state is-hidden"></div>
        <div id="qa-detail-state" class="panel-state is-hidden"></div>
      </section>

      <footer id="qa-footer" class="menu-footer is-hidden">
        <div class="progress-bar-container">
          <div class="progress-bar-track">
            <div id="qa-progress-fill" class="progress-bar-fill"></div>
            <div id="qa-progress-status" class="progress-status-text">Working...</div>
          </div>
          <button id="qm-btn-export-batch" type="button" class="qm-hidden" style="padding: 8px 16px; background: var(--color-accent); color: #1a5276; border: none; border-radius: 8px; font-weight: 700; font-size: 0.85rem; cursor: pointer; white-space: nowrap; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">Export (.xlsx)</button>
        </div>
      </footer>
    </main>
  </div>
    <div id="qa-views-pool" style="display:none;">
      ${VIEW_CEK_KARY}
      ${VIEW_SPKL}
      ${VIEW_KEHADIRAN}
      ${VIEW_DISTRIBUSI}
      ${VIEW_SETTINGS}
    </div>
  </div>
`;

  // --- QA ENGINE ---
  // 3. ICONS
  const ICONS = {
    search: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="6.5"></circle><path d="M16 16L21 21"></path></svg>`,
    grid: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><circle cx="6.2" cy="6.2" r="1.4"></circle><circle cx="12" cy="6.2" r="1.4"></circle><circle cx="17.8" cy="6.2" r="1.4"></circle><circle cx="6.2" cy="12" r="1.4"></circle><circle cx="12" cy="12" r="1.4"></circle><circle cx="17.8" cy="12" r="1.4"></circle><circle cx="6.2" cy="17.8" r="1.4"></circle><circle cx="12" cy="17.8" r="1.4"></circle><circle cx="17.8" cy="17.8" r="1.4"></circle></svg>`,
    lightbulb: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18H15M10 21H14M8.5 14.6C7.12 13.43 6.25 11.69 6.25 9.75C6.25 6.3 8.97 3.5 12.33 3.5C15.7 3.5 18.42 6.3 18.42 9.75C18.42 11.71 17.53 13.46 16.13 14.62C15.34 15.28 15 15.77 15 16.5V17.25H9.67V16.5C9.67 15.77 9.33 15.28 8.5 14.6Z"></path></svg>`,
    user: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path><circle cx="12" cy="7" r="4"></circle></svg>`,
    file: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>`,
    clock: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 16 14"></polyline></svg>`,
    map: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="1 6 1 22 8 18 16 22 23 18 23 2 16 6 8 2 1 6"></polygon><line x1="8" y1="2" x2="8" y2="18"></line><line x1="16" y1="6" x2="16" y2="22"></line></svg>`,
    settings: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path></svg>`,
    userCheck: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="8.5" cy="7" r="4"></circle><polyline points="17 11 19 13 23 9"></polyline></svg>`,
  };

  // 4. CORE ENGINE
  const qaState = {
    minimized: true,
    activePage: null,
    actions: [],
    results: []
  };

  let els = {};



  // Initialization
  function initQaEngine() {
    els = {
      floatingBtn: document.querySelector('#qa-floating-btn'),
      menuContainer: document.querySelector('#qa-menu-container'),
      menu: document.querySelector('#qa-menu-container .command-menu'),
      search: document.querySelector('#qa-search'),
      searchClear: document.querySelector('#qa-search-clear'),
      emptyState: document.querySelector('#qa-empty-state'),
      resultsState: document.querySelector('#qa-results-state'),
      detailState: document.querySelector('#qa-detail-state'),
      detailHeader: document.querySelector('.detail-header'),
      searchBar: document.querySelector('.search-bar'),
      detailTitle: document.querySelector('#qa-detail-title'),
      backBtn: document.querySelector('#qa-back-btn'),
      footer: document.querySelector('#qa-footer'),
      progressFill: document.querySelector('#qa-progress-fill'),
      progressStatus: document.querySelector('#qa-progress-status'),
      searchIcon: document.querySelector('.search-icon')
    };

    // Init icons
    if (els.searchIcon) els.searchIcon.innerHTML = ICONS.search;
    if (els.floatingBtn) els.floatingBtn.innerHTML = ICONS.grid;

    const savedPos = JSON.parse(localStorage.getItem('qa-pos') || '{"top":200, "edge":"right"}');
    if (els.floatingBtn) {
      els.floatingBtn.style.top = savedPos.top + 'px';
      if (savedPos.edge === 'right') els.floatingBtn.style.right = '20px';
      else els.floatingBtn.style.left = '20px';
      els.floatingBtn.classList.remove('button-is-hidden');
    }

    // Explicitly guarantee hidden menu state on initialization
    qaState.minimized = true;
    if (els.menuContainer) {
      els.menuContainer.classList.add('is-hidden');
    }

    renderEmptyState();
    setupEvents();
  }

  function setupEvents() {
    els.floatingBtn.onclick = toggleMenu;
    els.search.oninput = (e) => syncView(e.target.value);
    els.searchClear.onclick = () => { els.search.value = ''; syncView(''); els.search.focus(); };
    els.backBtn.onclick = () => goBack();

    // Event delegation for actions
    els.menu.addEventListener('click', (e) => {
      const btn = e.target.closest('.quick-action, .result-item');
      if (btn && btn.dataset.action) {
        window.QA.onActionClick(btn.dataset.action);
      }
    });

    document.addEventListener('keydown', (e) => {
      // Escape to close
      if (e.key === 'Escape' && !qaState.minimized) {
        if (qaState.activePage) goBack(); else toggleMenu();
      }

      // Ctrl+Shift+Z to toggle
      if (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        toggleMenu();
      }
    });

    // Click outside to close
    document.addEventListener('pointerdown', (e) => {
      if (!qaState.minimized && !els.menu.contains(e.target) && !els.floatingBtn.contains(e.target)) {
        toggleMenu();
      }
    });

    // Drag logic for floating button
    let isDragging = false, startY, startTop;
    els.floatingBtn.onpointerdown = (e) => {
      isDragging = false; startY = e.clientY; startTop = parseInt(els.floatingBtn.style.top);
      els.floatingBtn.setPointerCapture(e.pointerId);
      els.floatingBtn.onpointermove = (em) => {
        const dy = em.clientY - startY;
        if (Math.abs(dy) > 5) isDragging = true;
        els.floatingBtn.style.top = Math.max(0, Math.min(window.innerHeight - 50, startTop + dy)) + 'px';
      };
    };
    els.floatingBtn.onpointerup = (e) => {
      els.floatingBtn.onpointermove = null;
      els.floatingBtn.releasePointerCapture(e.pointerId);
      if (isDragging) {
        localStorage.setItem('qa-pos', JSON.stringify({ top: parseInt(els.floatingBtn.style.top), edge: els.floatingBtn.style.right ? 'right' : 'left' }));
      }
    };
  }

  function showHomePage() {
    goBack(true);
    if (els.search) els.search.value = '';
    syncView('');
  }

  function toggleMenu() {
    qaState.minimized = !qaState.minimized;
    if (!qaState.minimized) {
      showHomePage();
      els.menuContainer.classList.remove('is-hidden');
      els.menu.classList.add('is-opening');
      els.floatingBtn.classList.add('button-is-hidden');

      setTimeout(() => {
        els.menu.classList.remove('is-opening');
        els.search.focus();
      }, 280);
    } else {
      els.menu.classList.add('is-closing');
      setTimeout(() => {
        els.menuContainer.classList.add('is-hidden');
        els.menu.classList.remove('is-closing');
        els.floatingBtn.classList.remove('button-is-hidden');
      }, 280);
    }
  }

  function syncView(query) {
    const hasQuery = query.trim().length > 0;
    els.searchClear.classList.toggle('is-hidden', !hasQuery);

    // If typing while in a page, go back to search view
    if (hasQuery && qaState.activePage) {
      goBack(true);
    }

    if (qaState.activePage) return;

    const filtered = filterActions(query);
    const hasResults = hasQuery && filtered.length > 0;

    renderEmptyState(query, hasResults);
    if (hasResults) renderResults(filtered);

    els.emptyState.classList.toggle('is-hidden', hasResults);
    els.resultsState.classList.toggle('is-hidden', !hasResults);
  }

  function filterActions(query) {
    const q = normalizeSearchText(query);
    if (!q) return [];

    const seen = new Set();
    return [...qaState.actions, ...qaState.results]
      .filter(item => {
        if (!item || seen.has(item.label) || !item._searchText.includes(q)) return false;
        seen.add(item.label);
        return true;
      })
      .sort((a, b) => getSearchRank(a, q) - getSearchRank(b, q) || a.label.localeCompare(b.label));
  }

  function normalizeSearchText(value) {
    return String(value || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]+/g, ' ')
      .trim();
  }

  function buildSearchText(item) {
    return normalizeSearchText([
      item.label,
      item.description,
      ...(Array.isArray(item.keywords) ? item.keywords : [])
    ].join(' '));
  }

  function indexSearchItems(items) {
    return (Array.isArray(items) ? items : []).map(item => ({
      ...item,
      _searchLabel: normalizeSearchText(item && item.label),
      _searchText: buildSearchText(item || {})
    }));
  }

  function getSearchRank(item, query) {
    if (item._searchLabel === query) return 0;
    if (item._searchLabel.startsWith(query)) return 1;
    if (item._searchLabel.includes(query)) return 2;
    return 3;
  }

  function renderEmptyState(query = "", hasResults = false) {
    const title = query && !hasResults ? "No results" : "Quick Actions";
    const subtitle = query && !hasResults ? "Try a different search" : "Select an action below or start typing";

    els.emptyState.innerHTML = `
            <div class="empty-state">
                <div class="empty-state-icon">${ICONS.lightbulb}</div>
                <h1 class="empty-state-title">${title}</h1>
                <p class="empty-state-subtitle">${subtitle}</p>
                <div class="quick-action-grid">
                    ${qaState.actions.slice(0, 6).map(a => `
                        <button class="quick-action" data-action="${a.label}">
                            <span class="quick-action-icon">${a.iconMarkup || ICONS.grid}</span>
                            <span class="quick-action-label">${a.label}</span>
                        </button>
                    `).join('')}
                </div>
            </div>
        `;
  }

  function renderResults(list) {
    els.resultsState.innerHTML = `
            <div class="results-panel">
                <h2 class="results-group-title">Matches</h2>
                <div class="results-list">
                    ${list.map(item => `
                        <button class="result-item" data-action="${item.label}">
                            <div class="result-content">
                                <span class="result-label">${item.label}</span>
                            </div>
                            <span class="result-tag tag-${item.tagTone || 'blue'}">${item.tag || 'Action'}</span>
                        </button>
                    `).join('')}
                </div>
            </div>
        `;
  }

  function goBack(fast = false) {
    if (els.footer) els.footer.classList.add('is-hidden');
    if (fast) {
      qaState.activePage = null;
      els.detailState.classList.add('is-hidden');
      els.detailHeader.classList.add('is-hidden');
      els.searchBar.classList.remove('is-hidden');
      return;
    }

    // Immediately update states to show the home page behind the fading-out detail panel
    qaState.activePage = null;
    els.detailHeader.classList.add('is-hidden');
    els.searchBar.classList.remove('is-hidden');
    syncView(els.search.value);

    els.detailState.classList.remove('is-hidden'); // Ensure detailState remains visible during transition
    els.detailState.classList.add('is-exiting');

    setTimeout(() => {
      els.detailState.classList.add('is-hidden');
      els.detailState.classList.remove('is-exiting');
    }, 200);
  }

  // 5. PUBLIC API (Method 2: window.QA)
  window.QA = {
    registerActions: (actions) => { qaState.actions = indexSearchItems(actions); renderEmptyState(); },
    registerResults: (results) => { qaState.results = indexSearchItems(results); },
    onActionClick: (label) => {
      const action = qaState.actions.find(a => a.label === label) || qaState.results.find(r => r.label === label);
      if (action && action.onClick) {
        action.onClick(action);
      } else {
        window.QA.renderDetail(label, `<div style="padding: 40px; text-align:center;"><h3 style="color:var(--color-text);">${label}</h3><p style="color:var(--color-text-muted);">This action is ready for implementation logic.</p></div>`);
      }
    },
    renderDetail: (label, htmlContent) => {
      qaState.activePage = label;
      els.searchBar.classList.add('is-hidden');
      els.detailHeader.classList.remove('is-hidden');
      els.detailTitle.textContent = label;
      els.emptyState.classList.add('is-hidden');
      els.resultsState.classList.add('is-hidden');
      els.detailState.classList.remove('is-hidden');
      els.detailState.classList.add('is-entering');
      setTimeout(() => {
        els.detailState.classList.remove('is-entering');
      }, 260);

      const pool = document.getElementById('qa-views-pool') || document.body;

      // Parse the HTML content to find/create the pane element
      const temp = document.createElement('div');
      temp.innerHTML = htmlContent.trim();
      let paneEl = temp.firstElementChild;

      if (!paneEl) {
        paneEl = document.createElement('div');
        paneEl.innerHTML = htmlContent;
      }

      // Ensure the pane has class 'qm-pane' and an ID
      paneEl.classList.add('qm-pane');
      if (!paneEl.id) {
        paneEl.id = 'qm-pane-dynamic-' + label.toLowerCase().replace(/[^a-z0-9]+/g, '-');
      }

      const paneId = paneEl.id;

      // Check if this pane already exists in the DOM
      let targetEl = document.getElementById(paneId);
      if (!targetEl) {
        targetEl = paneEl;
        pool.appendChild(targetEl);
      }

      // Move all current children of detailState back to pool
      const activePanes = els.detailState.querySelectorAll('.qm-pane');
      activePanes.forEach(p => {
        p.style.display = 'none';
        p.classList.remove('active');
        pool.appendChild(p);
      });

      // Show and append the target pane to detailState
      targetEl.style.display = 'block';
      targetEl.classList.add('active');
      els.detailState.appendChild(targetEl);

      // Ensure Month/Year selects are populated if they exist in the new content
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();
      const monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
      const monthOptions = monthNames.map((name, i) => `<option value="${i + 1}" ${i + 1 === currentMonth ? 'selected' : ''}>${name}</option>`).join('');
      let yearOptions = '';
      for (let y = currentYear - 2; y <= currentYear + 1; y++) {
        yearOptions += `<option value="${y}" ${y === currentYear ? 'selected' : ''}>${y}</option>`;
      }

      const mSelect = targetEl.querySelector('#qm-input-bulan');
      const ySelect = targetEl.querySelector('#qm-input-tahun');
      if (mSelect && !mSelect.children.length) mSelect.innerHTML = monthOptions;
      if (ySelect && !ySelect.children.length) ySelect.innerHTML = yearOptions;
    },
    startProgress: (label, duration = 3000) => {
      els.footer.classList.remove('is-hidden');
      if (els.progressStatus) els.progressStatus.textContent = label + '...';
      let p = 0;
      const step = 100 / (duration / 100);
      const iv = setInterval(() => {
        p += step;
        if (p > 100) p = 100;
        if (els.progressFill) els.progressFill.style.width = p + '%';
        if (p >= 100) {
          clearInterval(iv);
          setTimeout(() => els.footer.classList.add('is-hidden'), 600);
        }
      }, 100);
    }
  };



  const PANEL_ELEMENTS = Object.freeze({
    panelShell: '#qm-panel, #qm-fab',
    allTabs: '.qm-tab',
    allPanes: '.qm-pane',
    activePane: '.qm-pane.active',
    accordionHeaders: '.qm-accordion-header',
    fab: '#qa-floating-btn',
    panel: '.command-menu',
    header: '#qm-header',
    result: '#qm-result',
    resultTitle: '#qm-result-title',
    resultBody: '#qm-result-body',
    history: '#qm-history',
    checkNrpPane: '#qm-pane-check-nrp',
    karyawanResults: '#qm-karyawan-results',
    karyawanSearchButton: '#qm-btn-karyawan-search',
    karyawanSearchInput: '#qm-input-karyawan-search',
    lookupButton: '#qm-btn-check',
    lookupNrp: '#qm-input-nrp',
    globalMonth: '#qm-input-bulan',
    globalYear: '#qm-input-tahun',
    attendanceCheckButton: '#qm-btn-hadir-check',
    attendanceCheckResult: '#qm-hadir-check-result',
    attendanceCheckNrp: '#qm-input-hadir-check-nrp',
    attendanceCheckStartDate: '#qm-input-hadir-check-start-date',
    attendanceCheckEndDate: '#qm-input-hadir-check-end-date',
    attendanceInputNrp: '#qm-input-hadir-nrp',
    attendanceInputDate: '#qm-input-hadir-tanggal',
    attendanceInputTime: '#qm-input-hadir-jam',
    attendanceInputStatus: '#qm-input-hadir-status',
    spklCheckButton: '#qm-btn-spkl-page-cek',
    spklCheckResult: '#qm-spkl-result',
    spklPageNrp: '#qm-spkl-page-nrp',
    spklPageMonth: '#qm-spkl-page-bulan',
    spklPageStartDate: '#qm-spkl-page-start-date',
    spklPageEndDate: '#qm-spkl-page-end-date',
    spklOnlineNrp: '#qm-spkl-online-nrp',
    spklOnlineDate: '#qm-spkl-online-date',
    spklEditModal: '#qm-modal-spkl-edit',
    spklEditDate: '#qm-edit-spkl-tgl',
    spklEditOt: '#qm-edit-spkl-ot',
    spklEditJamAwal: '#qm-edit-spkl-jam-awal',
    spklEditJamAkhir: '#qm-edit-spkl-jam-akhir',
    spklEditJamOt: '#qm-edit-spkl-jam-ot',
    hadirEditModal: '#qm-modal-hadir-edit',
    hadirEditNrp: '#qm-edit-hadir-nrp',
    hadirEditNama: '#qm-edit-hadir-nama',
    hadirEditBagian: '#qm-edit-hadir-bagian',
    hadirEditSeksi: '#qm-edit-hadir-seksi',
    hadirEditGroup: '#qm-edit-hadir-group',
    hadirEditTanggal: '#qm-edit-hadir-tanggal',
    hadirEditStatus: '#qm-edit-hadir-status',
    distribusiNrp: '#qm-input-distribusi-nrp',
    distribusiJkUseDistribusi: '#qm-dist-jk-use-distribusi',
    distribusiJkSelect: '#qm-dist-jk-select-input',
    distribusiJkTargetDate: '#qm-dist-jk-target-date',
    distribusiJkTargetDateEnd: '#qm-dist-jk-target-date-end',
    distribusiJkTargetShift: '#qm-dist-jk-target-shift',
    distribusiJkValue: '#qm-jk-value',
    distribusiSubsiUseDistribusi: '#qm-dist-subsi-use-distribusi',
    distribusiSubsiJkSelect: '#qm-dist-subsi-jk-select-input',
    distribusiSubsiDateStart: '#qm-input-distribusi-subsi-tgl-awal',
    distribusiSubsiDateEnd: '#qm-input-distribusi-subsi-tgl-akhir',
    distribusiSubsiShift: '#qm-input-distribusi-subsi-shift',
    distribusiSubsiBagian: '#qm-input-distribusi-subsi-bagian',
    distribusiSubsiSeksi: '#qm-input-distribusi-subsi-seksi',
    distribusiSubsiGrup: '#qm-input-distribusi-subsi-grup',
    distribusiKkNrp: '#qm-dist-KK-nrp',
    distribusiKkDate: '#qm-dist-KK-date',
    distribusiKkSelect: '#qm-dist-KK-select-input',
    distribusiKkBagian: '#qm-dist-KK-bagian',
    distribusiKkSeksi: '#qm-dist-KK-seksi',
    distribusiKkGrup: '#qm-dist-KK-grup',
    distribusiJkContainers: '#qm-dist-jk-options-container, #qm-dist-subsi-jk-container',
    distribusiJkPrimaryContainer: '#qm-dist-jk-options-container',
    distribusiSubsiJkContainer: '#qm-dist-subsi-jk-container',
    distribusiKkContainer: '#qm-dist-KK-options-container',
    anomalyBadge: '#qm-badge-anomali',
    anomalyList: '#qm-anomali-list',
    batchProgressBar: '#qa-progress-fill',
    batchStatus: '#qa-progress-status',
    batchResults: '#qm-batch-results',
    batchProgress: '#qa-footer',
    batchCheckButton: '#qm-btn-batch-check',
    batchClearButton: '#qm-btn-batch-clear',
    batchExportButton: '#qm-btn-export-batch',
    configCollapseMenu: '#qm-config-collapse-menu',
    themeLightButton: '#qm-btn-theme-light',
    themeDarkButton: '#qm-btn-theme-dark',
    showLogsButton: '#qm-btn-show-logs',
    logContainer: '#qm-log-container',
    logBody: '#qm-log-body',
    shortcutInput: '#qm-input-shortcut',
    shortcutRecordButton: '#qm-btn-record-shortcut',
    globalLoader: '#qm-global-loader',
    globalLoaderText: '#qm-global-loader-text',
    globalLoaderBar: '#qm-global-loader-bar',
    fixSpklNrp: '#qm-fix-spkl-nrp',
    fixSpklMonth: '#qm-fix-spkl-bulan',
    fixSpklYear: '#qm-fix-spkl-tahun',
    hadirBulanNrp: '#qm-input-hadir-bulan-nrp',
    hadirBulanMonth: '#qm-input-hadir-bulan-bln',
    hadirBulanYear: '#qm-input-hadir-bulan-thn',
    fixManyDate: '#qm-fix-many-date',
    distribusiDate: '#qm-input-distribusi-tanggal'
  });

  const uiAdapter = {
    resolve(target) {
      return PANEL_ELEMENTS[target] || target;
    },

    get(target, root = document) {
      return root.querySelector(this.resolve(target));
    },

    all(target, root = document) {
      return Array.from(root.querySelectorAll(this.resolve(target)));
    },

    html(target, html) {
      const el = this.get(target);
      if (el) renderSafe(el, html);
      return el;
    },

    text(target, value) {
      const el = this.get(target);
      if (el) el.textContent = value;
      return el;
    },

    value(target, value) {
      const el = this.get(target);
      if (!el || !('value' in el)) return el || null;
      if (value === undefined) return el.value;
      el.value = value;
      return el;
    },

    getValue(target, options = {}) {
      const raw = this.value(target);
      if (typeof raw !== 'string') return '';
      return options.trim ? raw.trim() : raw;
    },

    checked(target, value) {
      const el = this.get(target);
      if (!el || typeof el.checked !== 'boolean') return false;
      if (value === undefined) return el.checked;
      el.checked = !!value;
      return el.checked;
    },

    disabled(target, disabled) {
      const el = this.get(target);
      if (el && 'disabled' in el) el.disabled = !!disabled;
      return el;
    },

    setDataset(target, key, value) {
      const el = this.get(target);
      if (el) el.dataset[key] = value;
      return el;
    },

    addClass(target, className) {
      this.all(target).forEach(el => el.classList.add(className));
    },

    removeClass(target, className) {
      this.all(target).forEach(el => el.classList.remove(className));
    },

    toggleClass(target, className, force) {
      this.all(target).forEach(el => el.classList.toggle(className, force));
    },

    focus(target) {
      const el = this.get(target);
      if (el && typeof el.focus === 'function') el.focus();
      return el;
    },

    requestPrimaryFocus() {
      const input = document.querySelector('.qm-pane.active #qm-input-karyawan-search, .qm-pane.active #qm-input-nrp, #qm-input-nrp');
      if (input && typeof input.focus === 'function') input.focus();
      return input;
    },

    activePaneId() {
      return this.get('activePane')?.id || '';
    },

    activePaneKey() {
      return this.activePaneId().replace(/^qm-pane-/, '');
    },

    activatePane(paneKey) {
      this.all('allTabs').forEach(el => el.classList.toggle('active', el.dataset.pane === paneKey));
      this.all('allPanes').forEach(el => el.classList.toggle('active', el.id === `qm-pane-${paneKey}`));
    },

    confirm(message) {
      return window.confirm(message);
    },

    alert(message) {
      window.alert(message);
    },

    openUrl(url, target = '_blank') {
      if (target === '_self') window.location.href = url;
      else window.open(url, target);
    },

    reload(delay = 0) {
      setTimeout(() => window.location.reload(), delay);
    },

    findKaryawanSelect(key, type) {
      const selector = type === 'jk' ? '.qm-karyawan-jk-select' : '.qm-karyawan-kk-select';
      return this.all(selector).find(el => el.dataset.key === key) || null;
    }
  };

  const UI = {
    resultTimeout: null,

    showResult(type, title, bodyHtml) {
      clearTimeout(this.resultTimeout);
      const resultEl = uiAdapter.get('result');
      if (resultEl) {
        resultEl.classList.remove('success', 'danger', 'warning', 'qm-hidden', 'qm-fade-in');
        resultEl.classList.add(type, 'qm-visible-block');

        // Trigger animation in next frame
        requestAnimationFrame(() => {
          resultEl.classList.add('qm-fade-in');
        });

        uiAdapter.text('resultTitle', title);
        uiAdapter.html('resultBody', bodyHtml);
        this.resultTimeout = setTimeout(() => this.hideResult(), 3500);
      }
    },

    hideResult() {
      clearTimeout(this.resultTimeout);
      const resultEl = uiAdapter.get('result');
      if (resultEl) {
        resultEl.classList.remove('qm-fade-in');
        setTimeout(() => {
          resultEl.classList.add('qm-hidden');
          resultEl.classList.remove('qm-visible-block');
        }, 300);
      }
    },

    setLoading(on) {
      state.loading = on;
      const btn = uiAdapter.get('lookupButton');
      if (btn) {
        if (on) {
          btn.disabled = true;
          renderSafe(btn, '<span class="qm-spinner"></span> Mengarahkan...');
        } else {
          btn.disabled = false;
          renderSafe(btn, `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg> Cek`);
        }
      }
    },

    applyTheme(theme) {
      state.theme = theme;
      GM_setValue('qm_theme', theme);
      const panel = uiAdapter.get('panel');
      const fab = uiAdapter.get('fab');
      const loader = uiAdapter.get('globalLoader');

      const isDark = theme === 'dark';
      [panel, fab, loader].forEach(el => {
        if (el) el.classList.toggle('qm-dark', isDark);
      });

      // Update buttons in config
      const btnLight = uiAdapter.get('themeLightButton');
      const btnDark = uiAdapter.get('themeDarkButton');
      if (btnLight) btnLight.classList.toggle('active', !isDark);
      if (btnDark) btnDark.classList.toggle('active', isDark);
    },

    showGlobalLoader(title, initialMsg, allowCancel = false) {
      const existing = uiAdapter.get('globalLoader');
      if (existing) existing.remove();

      state.batchLogs = []; // Clear logs on new process
      const logBody = uiAdapter.get('logBody');
      if (logBody) renderSafe(logBody, '');

      const cancelBtnHtml = allowCancel
        ? `<div class="qm-loader-footer">
             <button id="qm-global-cancel-btn" class="qm-btn" style="width: 100%; padding: 6px 12px; font-size: 11px; height: auto;">Batalkan Proses</button>
           </div>`
        : '';

      document.body.insertAdjacentHTML('beforeend', `
        <div id="qm-global-loader" class="${state.theme === 'dark' ? 'qm-dark' : ''}">
          <div class="qm-loader-header">
            <div class="qm-spinner qm-spinner-dark qm-loader-spinner-size" style="width: 24px; height: 24px; border-width: 3px;"></div>
            <div class="qm-loader-body">
              <div class="qm-loader-title">${escapeHtml(title)}</div>
              <div id="qm-global-loader-text" class="qm-loader-text">${escapeHtml(initialMsg)}</div>
            </div>
          </div>
          <div class="qm-progress-container" title="Klik untuk lihat detail log">
            <div id="qm-global-loader-bar" class="qm-progress-bar"></div>
          </div>
          ${cancelBtnHtml}
        </div>
      `);

      pushLog(`Memulai proses: ${title} - ${initialMsg}`);
    },

    setGlobalProgress(pct, msg, silent = false) {
      if (msg) {
        const textEl = uiAdapter.get('globalLoaderText');
        if (textEl) textEl.textContent = msg;
        if (!silent) pushLog(msg);
      }
      const barEl = uiAdapter.get('globalLoaderBar');
      if (barEl) barEl.style.width = pct + '%';
    },

    hideGlobalLoader(delay = 800) {
      setTimeout(() => {
        const loader = uiAdapter.get('globalLoader');
        if (loader) {
          loader.classList.add('qm-loader-hiding');
          setTimeout(() => loader.remove(), 300);
        }
      }, delay);
    },

    startProgress(label, total = 0) {
      if (els.footer) {
        els.footer.classList.remove('is-hidden');
      }
      if (els.progressStatus) {
        els.progressStatus.textContent = label;
      }
      if (els.progressFill) {
        els.progressFill.style.width = '0%';
      }

      if (total > 0 && typeof total === 'number' && total < 100) {
        // total is item count, we will use updateProgress to manually drive it
      } else {
        // total is duration in ms, auto-progress
        const duration = typeof total === 'number' && total > 100 ? total : 3000;
        let p = 0;
        const step = 100 / (duration / 100);
        if (this._progressInterval) clearInterval(this._progressInterval);
        this._progressInterval = setInterval(() => {
          p += step;
          if (p > 100) p = 100;
          if (els.progressFill) els.progressFill.style.width = p + '%';
          if (p >= 100) {
            clearInterval(this._progressInterval);
            setTimeout(() => {
              if (els.footer) els.footer.classList.add('is-hidden');
            }, 600);
          }
        }, 100);
      }
    },

    updateProgress(current, total, statusText) {
      if (this._progressInterval) clearInterval(this._progressInterval);
      const pct = total > 0 ? Math.round((current / total) * 100) : 0;
      if (els.progressFill) {
        els.progressFill.style.width = pct + '%';
      }
      if (els.progressStatus) {
        els.progressStatus.textContent = statusText || `Memproses... ${current}/${total}`;
      }
    },

    endProgress() {
      if (this._progressInterval) clearInterval(this._progressInterval);
      if (els.progressFill) {
        els.progressFill.style.width = '100%';
      }
      setTimeout(() => {
        if (els.footer) els.footer.classList.add('is-hidden');
      }, 600);
    },

    renderHistory() {
      const histEl = uiAdapter.get('history');
      if (!histEl) return;
      if (!state.history.length) { renderSafe(histEl, ''); return; }
      const items = state.history.map(h => `<div class="qm-history-item"><span class="qm-badge ${h.ok ? 'ok' : 'err'}">${h.ok ? '✓' : '✗'}</span><span class="qm-history-nrp">${escapeHtml(h.nrp)}</span><span class="qm-history-label">${escapeHtml(h.label)}</span><span class="qm-history-time">${escapeHtml(h.time)}</span></div>`).join('');
      renderSafe(histEl, items);
    },

    pushHistory(nrp, ok, label) {
      state.history.unshift({ nrp, ok, label, time: now() });
      if (state.history.length > state.maxHistory) state.history.pop();
      this.renderHistory();
    }
  };


  function now() { return new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }); }
  function pushLog(msg, type = 'info') {
    const time = Logger._formatTime();
    state.batchLogs.push({ time, msg, level: type });
    if (state.batchLogs.length > 500) state.batchLogs.shift();

    const logBody = uiAdapter.get('logBody');
    if (logBody) {
      const div = document.createElement('div');
      div.className = 'qm-log-item';
      renderSafe(div, `<span class="qm-log-time">[${time}]</span><span class="qm-log-msg ${type}">${escapeHtml(msg)}</span>`);
      logBody.appendChild(div);
      logBody.scrollTop = logBody.scrollHeight;
    }
  }


  function initDraggable() {
    const panel = uiAdapter.get('panel');
    const header = uiAdapter.get('header');
    if (!panel || !header) return;

    let isDragging = false;
    let offset = { x: 0, y: 0 };

    if (state.panelPos) {
      panel.style.left = state.panelPos.left;
      panel.style.top = state.panelPos.top;
      panel.style.transform = 'none';
      panel.style.margin = '0';
    }

    header.style.cursor = 'move';
    header.addEventListener('mousedown', (e) => {
      if (e.target.closest('.qm-header-actions')) return;
      isDragging = true;
      const rect = panel.getBoundingClientRect();
      offset = {
        x: e.clientX - rect.left,
        y: e.clientY - rect.top
      };
      panel.style.transition = 'none';
      panel.style.transform = 'none';
      panel.style.margin = '0';
    });

    document.addEventListener('mousemove', (e) => {
      if (!isDragging) return;
      const x = e.clientX - offset.x;
      const y = e.clientY - offset.y;

      const left = x + 'px';
      const top = Math.max(0, y) + 'px';

      panel.style.left = left;
      panel.style.top = top;

      state.panelPos = { left, top };
    });

    document.addEventListener('mouseup', () => {
      if (isDragging) {
        isDragging = false;
        panel.style.transition = '';
        GM_setValue('qm_panel_pos', JSON.stringify(state.panelPos));
      }
    });
  }

  function renderLogs() {
    const logBody = uiAdapter.get('logBody');
    if (!logBody) return;

    if (state.batchLogs.length === 0) {
      renderSafe(logBody, '<div class="qm-text-muted qm-text-center qm-mt-xl">Belum ada log aktivitas.</div>');
      return;
    }

    renderSafe(logBody, state.batchLogs.map(log => `
      <div class="qm-log-item">
        <span class="qm-log-time">[${log.time}]</span>
        <span class="qm-log-msg ${log.level || log.type || 'info'}">${escapeHtml(log.msg)}</span>
      </div>
    `).join(''));
    logBody.scrollTop = logBody.scrollHeight;
  }

  /* ============================================================
   * 12. JK/KK CHANGE LOGIC
   * ============================================================ */

  function initJkChangeEvents() {
    on('click', '#qm-btn-KK-update', function () {
      handleUpdateKKMaster(panelReaders.distribusiKk());
    });

    let globalDebounce;
    on('input', '#qm-input-distribusi-nrp, #qm-dist-KK-nrp, #qm-input-distribusi-subsi-nrp, #qm-input-nrp', function () {
      const nrp = this.value.trim();
      clearTimeout(globalDebounce);
      globalDebounce = setTimeout(() => {
        refreshGlobalData(nrp);
      }, 600);
    });
  }

  /**
   * Unified Data Fetching & UI Synchronization
   */
  async function refreshGlobalData(nrp = '', bulan = '', tahun = '', sourceId = '') {
    const context = resolveSyncContext(sourceId);

    // 1. Detection
    if (!nrp) {
      nrp = firstFilledValue(context.nrpIds, true) || getPageContext().nrp || '';
    }
    if (!bulan) {
      bulan = firstFilledValue(context.bulanIds, false) || (new Date().getMonth() + 1);
    }
    if (!tahun) {
      tahun = firstFilledValue(context.tahunIds, false) || new Date().getFullYear();
    }

    if (!nrp || nrp.length < 4) return;
    const runId = ++state.refreshRunId;

    // 2. Sync all inputs across tabs
    syncGlobalInputs(nrp, bulan, tahun);

    // 3. Show loading in summary area
    const resBody = uiAdapter.get('resultBody');
    if (resBody && state.activeTab === 'check-nrp') {
      uiAdapter.html('resultBody', '<div class="qm-flex qm-items-center qm-gap-s qm-p-m"><span class="qm-spinner"></span> <span>Memuat data karyawan...</span></div>');
    }

    try {
      // 4. Single fetchEmployee call (cached internally)
      const emp = await fetchEmployee(nrp);
      if (runId !== state.refreshRunId) return;
      if (!emp.found) {
        if (resBody) uiAdapter.html('resultBody', '<div class="qm-text-danger qm-p-m">NRP tidak ditemukan.</div>');
        return;
      }

      // 5. Update Summary Info
      if (resBody) {
        uiAdapter.html('resultBody', `
          <div class="qm-p-m qm-bg-parchment qm-rounded-m qm-border">
            <div class="qm-flex qm-justify-between qm-mb-s">
              <span class="qm-text-muted qm-text-xs">Nama:</span>
              <span class="qm-font-bold qm-text-s">${escapeHtml(emp.nama)}</span>
            </div>
            <div class="qm-flex qm-justify-between qm-mb-s">
              <span class="qm-text-muted qm-text-xs">Bagian:</span>
              <span class="qm-text-s">${escapeHtml(emp.bagian || '-')}</span>
            </div>
            <div class="qm-flex qm-justify-between qm-mb-s">
              <span class="qm-text-muted qm-text-xs">Jam Kerja:</span>
              <span class="qm-text-s qm-font-mono qm-text-blue">${escapeHtml(emp.jk || '-')}</span>
            </div>
            <div class="qm-flex qm-justify-between">
              <span class="qm-text-muted qm-text-xs">Kalender:</span>
              <span class="qm-text-s qm-font-mono qm-text-teal">${escapeHtml(emp.KK || '-')}</span>
            </div>
          </div>
        `);
      }

      // 6. Refresh Active Pane Data
      const activePane = uiAdapter.activePaneId();
      if (activePane === 'qm-pane-distribusi') {
        await refreshDistribusiOptions(nrp, emp, runId);
      } else if (activePane === 'qm-pane-spkl') {
        // Additional SPKL specific refresh if needed
      }

    } catch (e) {
      Logger.error('refreshGlobalData error', e);
      if (runId === state.refreshRunId && resBody) uiAdapter.html('resultBody', `<div class="qm-text-danger qm-p-m">Error: ${e.message}</div>`);
    }
  }

  function activePaneId() {
    return uiAdapter.activePaneId();
  }

  function uniqueIds(ids) {
    return [...new Set(ids.filter(Boolean))];
  }

  function resolveSyncContext(sourceId = '') {
    const activePane = activePaneId();
    let key = 'global';

    if (sourceId === 'spkl' || sourceId.startsWith('qm-fix-spkl') || sourceId.startsWith('qm-spkl-page') || activePane === 'qm-pane-spkl') key = 'spkl';
    else if (sourceId === 'kehadiran' || sourceId.startsWith('qm-input-hadir')) key = 'kehadiran';
    else if (sourceId === 'check-nrp' || sourceId === 'qm-input-bulan' || sourceId === 'qm-input-tahun' || sourceId === 'qm-input-nrp' || activePane === 'qm-pane-check-nrp') key = 'check-nrp';

    const contextMap = {
      'check-nrp': {
        nrpIds: ['qm-input-nrp', 'qm-fix-spkl-nrp', 'qm-spkl-page-nrp', 'qm-spkl-online-nrp', 'qm-input-hadir-check-nrp', 'qm-input-hadir-nrp', 'qm-input-hadir-bulan-nrp', 'qm-input-distribusi-nrp', 'qm-dist-KK-nrp'],
        bulanIds: ['qm-input-bulan', 'qm-fix-spkl-bulan', 'qm-spkl-page-bulan', 'qm-input-hadir-bulan-bln'],
        tahunIds: ['qm-input-tahun', 'qm-fix-spkl-tahun', 'qm-input-hadir-bulan-thn']
      },
      spkl: {
        nrpIds: ['qm-fix-spkl-nrp', 'qm-spkl-page-nrp', 'qm-spkl-online-nrp', 'qm-input-nrp', 'qm-input-hadir-check-nrp', 'qm-input-hadir-nrp', 'qm-input-hadir-bulan-nrp', 'qm-input-distribusi-nrp', 'qm-dist-KK-nrp'],
        bulanIds: ['qm-fix-spkl-bulan', 'qm-spkl-page-bulan', 'qm-input-bulan', 'qm-input-hadir-bulan-bln'],
        tahunIds: ['qm-fix-spkl-tahun', 'qm-input-tahun', 'qm-input-hadir-bulan-thn']
      },
      kehadiran: {
        nrpIds: ['qm-input-hadir-check-nrp', 'qm-input-hadir-bulan-nrp', 'qm-input-hadir-nrp', 'qm-input-nrp', 'qm-fix-spkl-nrp', 'qm-spkl-page-nrp', 'qm-spkl-online-nrp', 'qm-input-distribusi-nrp', 'qm-dist-KK-nrp'],
        bulanIds: ['qm-input-hadir-bulan-bln', 'qm-input-bulan', 'qm-fix-spkl-bulan', 'qm-spkl-page-bulan'],
        tahunIds: ['qm-input-hadir-bulan-thn', 'qm-input-tahun', 'qm-fix-spkl-tahun']
      },
      global: {
        nrpIds: ['qm-input-nrp', 'qm-fix-spkl-nrp', 'qm-spkl-page-nrp', 'qm-spkl-online-nrp', 'qm-input-hadir-check-nrp', 'qm-input-hadir-nrp', 'qm-input-hadir-bulan-nrp', 'qm-input-distribusi-nrp', 'qm-dist-KK-nrp'],
        bulanIds: ['qm-input-bulan', 'qm-fix-spkl-bulan', 'qm-spkl-page-bulan', 'qm-input-hadir-bulan-bln'],
        tahunIds: ['qm-input-tahun', 'qm-fix-spkl-tahun', 'qm-input-hadir-bulan-thn']
      }
    };

    const selected = contextMap[key] || contextMap.global;
    return {
      nrpIds: uniqueIds(selected.nrpIds),
      bulanIds: uniqueIds(selected.bulanIds),
      tahunIds: uniqueIds(selected.tahunIds)
    };
  }

  function firstFilledValue(ids, trim = false) {
    for (const id of ids) {
      const rawValue = uiAdapter.getValue(`#${id}`, { trim });
      if (rawValue !== '' && rawValue !== null && rawValue !== undefined) return rawValue;
    }
    return '';
  }

  function syncGlobalInputs(nrp, bulan, tahun) {
    // NRP inputs
    const nrpIds = ['qm-input-nrp', 'qm-spkl-page-nrp', 'qm-spkl-online-nrp', 'qm-fix-spkl-nrp', 'qm-input-hadir-check-nrp', 'qm-input-hadir-nrp', 'qm-input-hadir-bulan-nrp', 'qm-input-distribusi-nrp', 'qm-dist-KK-nrp'];
    nrpIds.forEach(id => {
      const el = uiAdapter.get(`#${id}`);
      if (el && el.value !== nrp) el.value = nrp;
    });

    // Month selects
    const bulanIds = ['qm-input-bulan', 'qm-fix-spkl-bulan', 'qm-spkl-page-bulan', 'qm-input-hadir-bulan-bln'];
    bulanIds.forEach(id => {
      const el = uiAdapter.get(`#${id}`);
      if (el && el.value != bulan) el.value = bulan;
    });

    // Year selects
    const tahunIds = ['qm-input-tahun', 'qm-fix-spkl-tahun', 'qm-input-hadir-bulan-thn'];
    tahunIds.forEach(id => {
      const el = uiAdapter.get(`#${id}`);
      if (el && el.value != tahun) el.value = tahun;
    });

    // Date/Month picker sync
    const kkDate = uiAdapter.get('distribusiKkDate');
    if (kkDate) {
      const mStr = String(bulan).padStart(2, '0');
      kkDate.value = `${tahun}-${mStr}`;
    }
  }

  async function refreshDistribusiOptions(nrp, emp, refreshRunId) {
    const jkContainers = [
      uiAdapter.get('distribusiJkPrimaryContainer'),
      uiAdapter.get('distribusiSubsiJkContainer')
    ];
    const kkContainer = uiAdapter.get('distribusiKkContainer');

    jkContainers.forEach(c => { if (c) renderSafe(c, '<span class="qm-spinner qm-spinner-xs"></span>'); });
    if (kkContainer) renderSafe(kkContainer, '<span class="qm-spinner qm-spinner-xs"></span>');

    try {
      // 1. Fetch JK & KK options (using cache internally if available)
      const [jkOptions, kkOptions] = await Promise.all([
        fetchJkOptions(nrp),
        fetchKKOptions(nrp)
      ]);
      if (refreshRunId !== state.refreshRunId) return;

      jkContainers.forEach(c => {
        if (c && jkOptions.length) {
          const id = c.id === 'qm-dist-jk-options-container' ? 'qm-dist-jk-select-input' : 'qm-dist-subsi-jk-select-input';
          renderSafe(c, `<select id="${id}" class="qm-select qm-text-s">${jkOptions.map(o => `<option value="${escapeHtml(o.val)}" ${o.selected ? 'selected' : ''}>${escapeHtml(o.txt)}</option>`).join('')}</select>`);
        }
      });

      if (kkContainer && kkOptions.length) {
        renderSafe(kkContainer, `<select id="qm-dist-KK-select-input" class="qm-select qm-text-s">${kkOptions.map(o => `<option value="${escapeHtml(o.val)}" ${o.selected ? 'selected' : ''}>${escapeHtml(o.txt)}</option>`).join('')}</select>`);
      }

      // 2. Fetch and populate Bagian/Seksi/Grup/Shift
      // We don't await this fully for the first render if we want it "instant"
      // but we need the options to be there. updateDistribusiDropdowns now handles its own internal cache.
      await updateDistribusiDropdowns(nrp, '', '', refreshRunId);
      if (refreshRunId !== state.refreshRunId) return;

      // 3. Set values from emp data
      const setVal = (id, val) => {
        const el = uiAdapter.get(`#${id}`);
        if (el && val) {
          // If value not present in options, add a temporary one to show it immediately
          if (!Array.from(el.options).some(opt => opt.value === val)) {
            const tempOpt = document.createElement('option');
            tempOpt.value = val;
            tempOpt.textContent = val + ' (Memuat...)';
            el.appendChild(tempOpt);
          }
          el.value = val;
          el.dispatchEvent(new Event('change', { bubbles: true }));
        }
      };

      setVal('qm-input-distribusi-subsi-bagian', emp.bagian);
      setVal('qm-dist-KK-bagian', emp.bagian);
      setVal('qm-input-distribusi-subsi-seksi', emp.seksi);
      setVal('qm-dist-KK-seksi', emp.seksi);
      setVal('qm-input-distribusi-subsi-grup', emp.group);
      setVal('qm-dist-KK-grup', emp.group);

      attachDistribusiListeners(nrp);
    } catch (e) {
      if (refreshRunId === state.refreshRunId) Logger.error('refreshDistribusiOptions error', e);
    }
  }

  function attachDistribusiListeners(nrp) {
    const pairs = [
      { bag: 'qm-input-distribusi-subsi-bagian', sek: 'qm-input-distribusi-subsi-seksi' },
      { bag: 'qm-dist-KK-bagian', sek: 'qm-dist-KK-seksi' }
    ];

    pairs.forEach(p => {
      const elBag = uiAdapter.get(`#${p.bag}`);
      const elSek = uiAdapter.get(`#${p.sek}`);
      if (elBag) elBag.dataset.nrp = nrp;
      if (elSek) elSek.dataset.nrp = nrp;

      if (elBag && !elBag.dataset.hasListener) {
        elBag.dataset.hasListener = 'true';
        elBag.addEventListener('change', () => updateDistribusiDropdowns(elBag.dataset.nrp || nrp, elBag.value, '', state.refreshRunId));
      }
      if (elSek && !elSek.dataset.hasListener) {
        elSek.dataset.hasListener = 'true';
        elSek.addEventListener('change', () => updateDistribusiDropdowns(elSek.dataset.nrp || nrp, elBag?.value || '', elSek.value, state.refreshRunId));
      }
    });
  }


  /**
   * Fetches the distribution page and synchronizes all dropdowns (Per Subsi & Kalender Kerja).
   */
  async function updateDistribusiDropdowns(nrp, bag = '', sek = '', refreshRunId = state.refreshRunId) {
    const isOS = nrp.length === 8;
    const cacheKey = `qm_dist_html_${isOS ? 'os' : 'reg'}_${bag}_${sek}`;

    // Attempt to use cache for immediate UI update
    const cachedHtml = sessionStorage.getItem(cacheKey);
    if (cachedHtml) {
      applyDistDropdowns(cachedHtml, refreshRunId);
    }

    let url = distribusiUrl(nrp);
    if (bag || sek) {
      url += `?kode_bagian=${encodeURIComponent(bag)}&kode_seksi=${encodeURIComponent(sek)}`;
    }

    try {
      const html = await hrisFetch(url);
      if (refreshRunId !== state.refreshRunId) return;
      sessionStorage.setItem(cacheKey, html);
      applyDistDropdowns(html, refreshRunId);
    } catch (e) {
      if (refreshRunId === state.refreshRunId) Logger.error('updateDistribusiDropdowns fetch error', e);
    }
  }

  function applyDistDropdowns(html, refreshRunId = state.refreshRunId) {
    if (refreshRunId !== state.refreshRunId) return;
    const doc = parseHTML(html);
    const sync = (targetIds, selector) => {
      const source = doc.querySelector(selector);
      if (!source) return;
      targetIds.forEach(id => {
        if (refreshRunId !== state.refreshRunId) return;
        const el = document.getElementById(id);
        if (el) {
          const currentVal = el.value;
          const firstOpt = el.options[0];
          renderSafe(el, source.innerHTML);
          if (firstOpt && !firstOpt.value && el.options.length > 0 && el.options[0].value) {
            el.insertBefore(firstOpt, el.firstChild);
          }
          if (currentVal && Array.from(el.options).some(o => o.value === currentVal)) {
            el.value = currentVal;
          }
        }
      });
    };

    sync(['qm-input-distribusi-subsi-bagian', 'qm-dist-KK-bagian'], 'select[name="kode_bagian"]');
    sync(['qm-input-distribusi-subsi-seksi', 'qm-dist-KK-seksi'], 'select[name="kode_seksi"]');
    sync(['qm-input-distribusi-subsi-grup', 'qm-dist-KK-grup'], 'select[name="kode_group"]');
    sync(['qm-input-distribusi-subsi-shift'], 'select[name="status_shift"], select[name="shift"], select[name="kode_shift"], select[id*="shift"], select[name*="shift"]');

    // Fallback if Shift dropdown is still empty (only has "Pilih Shift")
    const shiftEl = uiAdapter.get('distribusiSubsiShift');
    if (refreshRunId === state.refreshRunId && shiftEl && shiftEl.options.length <= 1) {
      shiftEl.innerHTML += `
        <option value="1">Shift 1 (Pagi)</option>
        <option value="2">Shift 2 (Siang)</option>
        <option value="3">Shift 3 (Malam)</option>
      `;
    }
  }


  async function fetchJkOptions(nrp) {
    const cached = sessionStorage.getItem('qm_jk_options_' + nrp);
    if (cached) return JSON.parse(cached);

    const emp = await fetchEmployee(nrp);
    if (!emp.found || !emp.id) throw new Error('Data karyawan tidak lengkap.');

    let editUrl = emp.editUrl;
    if (!editUrl) {
      // Fallback guess if not found in detail page
      const isOS = nrp.length === 8;
      editUrl = employeeEditUrl(nrp, emp.id);
    }

    let html = '', select = null;
    try {
      html = await hrisFetch(editUrl);
      cachedEditHtml = html; // Cache HTML for later saveJkMaster use
      const doc = parseHTML(html);

      // --- BROWSER DIAGNOSTIC LOGS ---
      const title = doc.querySelector('title')?.textContent.trim() || '';
      const h1 = doc.querySelector('h1')?.textContent.trim() || '';
      console.log(`[QM DIAGNOSTIC] fetchJkOptions: URL=${editUrl}, Size=${html.length} chars, Title="${title}", H1="${h1}"`);

      const allSelects = Array.from(doc.querySelectorAll('select')).map(s => s.name || s.id || 'no-name');
      console.log(`[QM DIAGNOSTIC] Found select fields:`, allSelects);

      if (doc.querySelector('input[type="password"]') || title.toLowerCase().includes('login')) {
        throw new Error('Sesi login kedaluwarsa. Silakan login kembali di tab HRIS.');
      }
      // -------------------------------

      // Try multiple possible selectors
      const possibleNames = ['kerja_hour_code', 'kode_jam_kerja', 'jam_kerja', 'kerja_hour'];
      for (const name of possibleNames) {
        select = doc.querySelector(`select[name="${name}"], [name="${name}"]`);
        if (select) break;
      }
      if (false) {
        let altUrl = '';
        if (editUrl.includes('/editgeneral/')) {
          altUrl = editUrl.replace('/editgeneral/', '/edit/');
        } else if (editUrl.includes('/edit/')) {
          altUrl = editUrl.replace('/edit/', '/editgeneral/');
        }

        if (altUrl) {
          console.log(`[QM INFO] JK select not found in ${editUrl}. Trying alternative URL: ${altUrl}`);
          try {
            const altHtml = await hrisFetch(altUrl);
            const altDoc = parseHTML(altHtml);
            for (const name of possibleNames) {
              const altSelect = altDoc.querySelector(`select[name="${name}"], [name="${name}"]`);
              if (altSelect) {
                select = altSelect;
                editUrl = altUrl; // Update to the successful URL
                cachedEditHtml = altHtml;
                break;
              }
            }
          } catch (altErr) {
            console.error(`[QM ERROR] fetchJkOptions fallback error:`, altErr);
          }
        }
      }
    } catch (e) {
      console.error(`[QM ERROR] fetchJkOptions error:`, e);

    }

    if (!select) {
      throw new Error(`Elemen pilihan jam kerja tidak ditemukan (Mungkin karyawan outsource / tanpa akses edit).`);
    }

    const options = Array.from(select.querySelectorAll('option')).map(opt => ({
      val: opt.value,
      txt: opt.textContent.trim(),
      selected: opt.hasAttribute('selected') || opt.selected
    }));

    sessionStorage.setItem('qm_jk_options_' + nrp, JSON.stringify(options));
    return options;
  }

  async function handleSaveJkChange(input = panelReaders.distribusiJk()) {
    const nrp = String(input?.nrp || '').trim();
    if (!nrp) { uiAdapter.alert('Harap isi NRP.'); return; }

    const useDistribusi = !!input?.useDistribusi;
    const jk = String(input?.jk || '').trim();
    const date = String(input?.date || '').trim();
    const dateEnd = String(input?.dateEnd || '').trim();
    const shift = String(input?.shift || '').trim();
    const oldJk = sessionStorage.getItem('qm_jk_' + nrp);

    const emp = await fetchEmployee(nrp);
    if (!emp.found) {
      uiAdapter.alert('Data karyawan tidak ditemukan untuk NRP ' + nrp);
      return;
    }

    Logger.info(`Starting handleSaveJkChange for ${nrp}. New JK: ${jk}, Use Distribusi: ${useDistribusi}`);
    if (useDistribusi) beginCancelableDistributionFlow();
    UI.showGlobalLoader('Processing JK', 'Updating Master Data...', !!useDistribusi);

    try {
      UI.setGlobalProgress(10, 'Memeriksa data lama...');
      // If we are in distribusi mode, save the old JK to restore later
      if (useDistribusi && oldJk && oldJk !== jk) {
        Logger.info(`Saving old JK ${oldJk} for later restoration.`);
        sessionStorage.setItem('qm_jk_to_restore_' + nrp, oldJk);
      }

      // Step 1: Update Master Data if JK changed
      if (oldJk !== jk) {
        UI.setGlobalProgress(30, 'Memperbarui Master Data (editgeneral)...');
        await saveJkMaster(nrp, jk);
        throwIfCancelled();

        // Update UI in-place for immediate feedback
        const jkLabel = uiAdapter.get('distribusiJkValue');
        if (jkLabel) jkLabel.textContent = jk;
        sessionStorage.setItem('qm_jk_' + nrp, jk);
        Logger.success(`Master Data updated to ${jk} for ${nrp}.`);
      } else {
        Logger.info(`Kode JK sama (${jk}), skip update master.`);
        UI.setGlobalProgress(30, 'Master sudah sesuai, lanjut distribusi...');
      }

      // Step 2: Distribution (Background or Redirect)
      const ctx = getPageContext();
      ctx.nrp = nrp;

      if (useDistribusi) {
        throwIfCancelled();
        const progressMsg = oldJk === jk ? 'Master sesuai. Memproses Distribusi...' : 'Master terupdate. Memproses Distribusi...';
        UI.setGlobalProgress(80, progressMsg);
        try {
          const success = await executeBackgroundDistribusi({
            nrp: nrp,
            jk: jk,
            tglAwal: date,
            tglAkhir: dateEnd || date,
            shift: shift,
            bagian: emp.bagian,
            seksi: emp.seksi,
            grup: emp.group
          });

          if (success) {
            UI.setGlobalProgress(100, 'Distribusi Selesai!');
            UI.showResult('success', 'Berhasil', 'Master data & Distribusi selesai.');
          } else {
            UI.showResult('success', 'Selesai', 'Request terkirim, harap cek hasil di tabel.');
          }
          setTimeout(() => window.location.reload(), 1500);
        } catch (distErr) {
          Logger.error('Background distribution error', distErr);
          if (state.cancelRequested || isAbortError(distErr)) {
            UI.showResult('info', 'Dibatalkan', 'Proses distribusi dihentikan oleh pengguna.');
            UI.hideGlobalLoader(0);
          } else {
            UI.showResult('warning', 'Master OK, Distribusi Gagal', distErr.message);
            UI.hideGlobalLoader(5000);
          }
        }
      } else {
        // Redirect mode
        Logger.info(`Master Data updated for ${nrp}. Proceeding to distribution via redirect...`);
        UI.setGlobalProgress(80, 'Master terupdate. Mengalihkan ke halaman Distribusi...');

        // Store return URL for auto-return after redirect completion
        createAutomationFlow('distribusi-jk', window.location.href, { nrp, date, dateEnd, shift });
        const redirectUrl = buildDistribusiLink(ctx, date, shift, dateEnd);

        // Small delay to ensure sessionStorage/Master Update is committed
        setTimeout(() => {
          window.location.href = redirectUrl;
        }, 800);
      }
    } catch (e) {
      Logger.error(`Error in handleSaveJkChange: ${e.message}`, e);
      if (state.cancelRequested || isAbortError(e)) {
        UI.showResult('info', 'Dibatalkan', 'Proses distribusi dihentikan oleh pengguna.');
        UI.hideGlobalLoader(0);
      } else {
        UI.showResult('danger', 'Gagal', 'Terjadi kesalahan: ' + e.message);
        UI.hideGlobalLoader();
      }
    } finally {
      if (useDistribusi) clearCancelableDistributionFlow();
    }
  }

  /**
   * Unified background distribution logic used by both Per NRP and Per Subsi.
   */
  async function executeBackgroundDistribusi(params) {
    const { nrp, jk, tglAwal, tglAkhir, shift, bagian, seksi, grup } = params;
    const distUrl = distribusiUrl(nrp);

    throwIfCancelled();
    UI.setGlobalProgress(10, 'Mengambil form distribusi...');
    const html = await hrisFetch(distUrl);
    if (html.includes('id="login-form"') || html.includes('login_form')) throw new Error('Sesi berakhir. Silakan login kembali.');

    throwIfCancelled();
    UI.setGlobalProgress(30, 'Menganalisa form & CSRF token...');
    const doc = parseHTML(html);
    let form = doc.querySelector('form[action*="distribusijamkerja"]');
    if (!form) form = doc.querySelector('form');
    if (!form) throw new Error('Form distribusi tidak ditemukan.');

    UI.setGlobalProgress(50, 'Menyusun parameter request...');
    const postData = new FormData();
    const headers = {}; // Content-Type will be set automatically by the browser for FormData

    const csrfMeta = doc.querySelector('meta[name="csrf-token"], meta[name="csrf-test-name"], meta[name="_token"]');
    if (csrfMeta) {
      const token = csrfMeta.getAttribute('content');
      headers['X-CSRF-TOKEN'] = token;
      headers['X-XSRF-TOKEN'] = token;
    }

    // Collect all form fields initially
    form.querySelectorAll('input, select, textarea').forEach(el => {
      if (!el.name || el.type === 'submit' || el.type === 'button') return;
      if ((el.type === 'checkbox' || el.type === 'radio') && !el.checked) return;

      let val = el.value;
      if (el.tagName === 'SELECT') {
        const opt = el.querySelector('option[selected]') || el.options[el.selectedIndex] || el.options[0];
        val = opt ? opt.value : el.value;
      }
      postData.append(el.name, val);
    });

    const setVal = (selectors, val, nameFallback) => {
      let found = false;
      for (const sel of selectors) {
        const el = doc.querySelector(sel);
        if (el && el.name) {
          postData.set(el.name, val);
          found = true;
          break;
        }
      }
      if (!found && nameFallback) postData.set(nameFallback, val);
    };

    setVal(['select[name="kode_jam_kerja"]', 'select[name*="jk"]', 'input[name="kode_jam_kerja"]'], jk, 'kode_jam_kerja');
    setVal(['input[name="start_date"]', 'input[name="tanggal_awal"]', 'input[id*="tanggal_awal"]', 'input[name*="tgl_awal"]'], tglAwal, 'start_date');
    setVal(['input[name="end_date"]', 'input[name="tanggal_akhir"]', 'input[id*="tanggal_akhir"]', 'input[name*="tgl_akhir"]'], tglAkhir, 'end_date');
    setVal(['select[name="status_shift"]', 'select[name="shift"]', 'select[name*="shift"]', 'select[name="kode_shift"]'], shift, 'status_shift');
    setVal(['select[name="kode_bagian"]', 'select[name*="bagian"]'], bagian, 'kode_bagian');
    setVal(['select[name="kode_seksi"]', 'select[name*="seksi"]'], seksi, 'kode_seksi');
    setVal(['select[name="kode_group"]', 'select[name*="group"]'], grup, 'kode_group');

    // Handle NRP range (single NRP or range)
    if (typeof nrp === 'object') {
      setVal(['input[id*="nrp_initial"]', 'input[name*="nrp_initial"]', 'input[name*="nrp_awal"]', 'input[name="nrp_initial_text"]'], nrp.awal, 'nrp_initial_text');
      setVal(['input[id*="nrp_final"]', 'input[name*="nrp_final"]', 'input[name*="nrp_akhir"]', 'input[name="nrp_final_text"]'], nrp.akhir || nrp.awal, 'nrp_final_text');
    } else {
      setVal(['input[id*="nrp_initial"]', 'input[name*="nrp_initial"]', 'input[name*="nrp_awal"]', 'input[name="nrp_initial_text"]'], nrp, 'nrp_initial_text');
      setVal(['input[id*="nrp_final"]', 'input[name*="nrp_final"]', 'input[name*="nrp_akhir"]', 'input[name="nrp_final_text"]'], nrp, 'nrp_final_text');
    }

    // Robust CSRF Detection from form fields if meta fails
    form.querySelectorAll('input[type="hidden"]').forEach(el => {
      if (el.name && (el.name.toLowerCase().includes('csrf') || el.name.toLowerCase().includes('token'))) {
        if (!headers['X-CSRF-TOKEN']) {
          headers['X-CSRF-TOKEN'] = el.value;
          headers['X-XSRF-TOKEN'] = el.value;
          Logger.info('CSRF token found in hidden form field.');
        }
      }
    });

    if (!headers['X-CSRF-TOKEN']) {
      throw new Error('Token CSRF tidak ditemukan. Struktur halaman mungkin berubah.');
    }

    const submitBtn = form.querySelector('button[type="submit"], input[type="submit"]');
    if (submitBtn && submitBtn.name) {
      postData.set(submitBtn.name, submitBtn.value || 'Start Distribusi');
    } else {
      postData.set('btnSubmit', 'Start Distribusi');
    }

    UI.setGlobalProgress(70, 'Mengirim data (POST)... Harap tunggu, proses ini lama.');
    const action = form.getAttribute('action') || distUrl;
    const targetUrl = action.startsWith('http') ? action : `https://hris.kti.co.id${action}`;
    const abortController = new AbortController();
    state.activeAbortController = abortController;

    Logger.info(`Sending background distribution POST to ${targetUrl}...`);
    Logger.info('POST Keys:', Array.from(postData.keys()).join(', '));

    try {
      startBackgroundHeartbeat();
      const response = await fetchWithTimeout(targetUrl, {
        method: 'POST',
        headers: headers,
        body: postData,
        referrer: distUrl,
        referrerPolicy: 'origin-when-cross-origin',
        signal: abortController.signal
      }, 300000);

      if (!response.ok) throw new Error('Distribusi gagal (HTTP ' + response.status + ')');

      stopBackgroundHeartbeat();
      UI.setGlobalProgress(90, 'Membaca respon server...');
      const resText = await response.text();
      Logger.info('Response received from distribution server.');

      if (isHrisSuccess(resText)) {
        UI.setGlobalProgress(100, 'Selesai!');
        return true;
      } else {
        Logger.warn('Respon Distribusi (Bukan sukses)', resText.substring(0, 500));
        return false;
      }
    } catch (fetchErr) {
      stopBackgroundHeartbeat();
      if (fetchErr.name === 'AbortError') {
        if (state.cancelRequested) {
          throw new Error('Distribusi dibatalkan oleh pengguna.');
        }
        throw new Error('Distribusi timeout (300 detik). Harap cek secara manual.');
      }
      throw fetchErr;
    } finally {
      state.activeAbortController = null;
    }
  }

  // Removed in favor of unified refreshDistribusiPane
  // async function refreshSubsiJkOptionsInPane() { ... }

  // Removed in favor of unified updateDistribusiDropdowns
  // async function updateSubsiDropdowns(nrp, bag = '', sek = '') { ... }

  async function handleDistribusiSubsi(input = panelReaders.distribusiSubsi()) {
    const nrp = String(input?.nrp || '').trim();
    if (!nrp) { uiAdapter.alert('Gagal mendeteksi NRP. Pastikan Anda berada di halaman profile atau tabel kehadiran.'); return; }

    const useDistribusi = !!input?.useDistribusi;
    const jk = String(input?.jk || '').trim();
    const tglAwal = String(input?.tglAwal || '').trim();
    const tglAkhir = String(input?.tglAkhir || '').trim();
    const shift = String(input?.shift || '').trim();
    const bagian = String(input?.bagian || '').trim();
    const seksi = String(input?.seksi || '').trim();
    const grup = String(input?.grup || '').trim();

    if (!jk) { uiAdapter.alert('Tunggu opsi Jam Kerja termuat terlebih dahulu.'); return; }
    if (!tglAwal || !tglAkhir) { uiAdapter.alert('Harap isi Tanggal Awal dan Akhir.'); return; }

    Logger.info(`Starting handleDistribusiSubsi. JK: ${jk}, Bagian: ${bagian}, Seksi: ${seksi}, Grup: ${grup}`);

    if (useDistribusi) {
      beginCancelableDistributionFlow();
      UI.showGlobalLoader('Processing Subsi', 'Menyiapkan data...', true);
      try {
        const isOS = nrp.length === 8;
        const nrpAwal = isOS ? '00000000' : '0000';
        const nrpAkhir = isOS ? '99999999' : '9999';

        throwIfCancelled();
        const success = await executeBackgroundDistribusi({
          nrp: { awal: nrpAwal, akhir: nrpAkhir },
          jk,
          tglAwal,
          tglAkhir,
          shift,
          bagian,
          seksi,
          grup
        });

        if (success) {
          UI.showResult('success', 'Berhasil', 'Proses Distribusi Subsi Selesai.');
        } else {
          UI.showResult('success', 'Selesai', 'Request terkirim, harap cek hasil di tabel.');
        }
        uiAdapter.reload(1500);
      } catch (e) {
        Logger.error('handleDistribusiSubsi error', e);
        if (state.cancelRequested || isAbortError(e)) {
          UI.showResult('info', 'Dibatalkan', 'Proses distribusi dihentikan oleh pengguna.');
          UI.hideGlobalLoader(0);
        } else {
          UI.showResult('danger', 'Gagal', 'Error: ' + e.message);
          UI.hideGlobalLoader();
        }
      } finally {
        clearCancelableDistributionFlow();
      }
    } else {
      const base = distribusiUrl(nrp);
      const url = `${base}?qm_auto_distribusi_subsi=1&jk=${encodeURIComponent(jk)}&tglAwal=${encodeURIComponent(tglAwal)}&tglAkhir=${encodeURIComponent(tglAkhir)}&bagian=${encodeURIComponent(bagian)}&seksi=${encodeURIComponent(seksi)}&grup=${encodeURIComponent(grup)}&shift=${encodeURIComponent(shift)}&nrp=${encodeURIComponent(nrp)}`;

      // Store return URL for auto-return after redirect completion
      createAutomationFlow('distribusi-subsi', window.location.href, { nrp, tglAwal, tglAkhir, shift });
      uiAdapter.openUrl(url, '_self');
    }
  }

  async function saveJkMaster(nrp, jk) {
    Logger.info(`saveJkMaster started for NRP ${nrp} with JK ${jk}`);
    const emp = await fetchEmployee(nrp);
    if (!emp.found || !emp.id) {
      Logger.error(`Data karyawan tidak lengkap untuk NRP ${nrp}`);
      throw new Error('Data karyawan tidak lengkap.');
    }

    let editUrl = emp.editUrl;
    if (!editUrl) {
      editUrl = employeeEditUrl(nrp, emp.id);
    }

    Logger.info(`Fetching edit form from ${editUrl}...`);
    // NRP-aware caching for cachedEditHtml
    if (state._lastEditNrp !== nrp) {
      cachedEditHtml = null;
      state._lastEditNrp = nrp;
    }

    const html = cachedEditHtml || await hrisFetch(editUrl);
    cachedEditHtml = html;

    const doc = parseHTML(html);
    let form = null;
    const possibleNames = ['kerja_hour_code', 'kode_jam_kerja', 'jam_kerja', 'kerja_hour'];

    for (const name of possibleNames) {
      const el = doc.querySelector(`[name="${name}"]`);
      if (el) {
        form = el.closest('form');
        if (form) {
          Logger.info(`Found form via field name: ${name}`);
          break;
        }
      }
    }

    if (false) {
      let altUrl = '';
      if (editUrl.includes('/editgeneral/')) {
        altUrl = editUrl.replace('/editgeneral/', '/edit/');
      } else if (editUrl.includes('/edit/')) {
        altUrl = editUrl.replace('/edit/', '/editgeneral/');
      }

      if (altUrl) {
        Logger.info(`Form not found in ${editUrl}. Trying alternative URL: ${altUrl}`);
        try {
          const altHtml = await hrisFetch(altUrl);
          const altDoc = parseHTML(altHtml);
          for (const name of possibleNames) {
            const el = altDoc.querySelector(`[name="${name}"]`);
            if (el) {
              const altForm = el.closest('form');
              if (altForm) {
                form = altForm;
                editUrl = altUrl;
                cachedEditHtml = altHtml;
                break;
              }
            }
          }
        } catch (altErr) {
          Logger.error(`[QM ERROR] saveJkMaster fallback error:`, altErr);
        }
      }
    }

    if (!form) {
      Logger.error(`Gagal menemukan form edit master data di ${editUrl}`);
      throw new Error('Form edit data tidak ditemukan.');
    }

    const params = new URLSearchParams();
    const headers = {
      'Content-Type': 'application/x-www-form-urlencoded',
      'X-Requested-With': 'XMLHttpRequest'
    };

    // Robust CSRF Detection
    const csrfMeta = doc.querySelector('meta[name="csrf-token"], meta[name="csrf-test-name"], meta[name="_token"]');
    if (csrfMeta) {
      const token = csrfMeta.getAttribute('content');
      headers['X-CSRF-TOKEN'] = token;
      headers['X-XSRF-TOKEN'] = token;
      Logger.info('CSRF token found in meta tags.');
    }

    Logger.info('Building POST parameters...');
    form.querySelectorAll('input, select, textarea').forEach(el => {
      if (!el.name || el.type === 'submit' || el.type === 'button') return;
      if ((el.type === 'checkbox' || el.type === 'radio') && !el.checked) return;

      let val = el.value;
      // For SELECT elements in a parsed doc, el.value might be empty if no option is marked 'selected'
      if (el.tagName === 'SELECT') {
        const opt = el.querySelector('option[selected]') || el.options[0];
        val = opt ? opt.value : el.value;
      }

      // If it's a CSRF hidden field and we don't have a token yet
      if (el.name.toLowerCase().includes('csrf') || el.name.toLowerCase().includes('token')) {
        if (!headers['X-CSRF-TOKEN']) headers['X-CSRF-TOKEN'] = val;
      }

      if (possibleNames.includes(el.name)) {
        params.append(el.name, jk);
        Logger.info(`Field matched: ${el.name} = ${jk}`);
      } else {
        params.append(el.name, val);
      }
    });

    if (!headers['X-CSRF-TOKEN']) {
      throw new Error('Token CSRF tidak ditemukan. Struktur halaman mungkin berubah.');
    }

    const action = form.getAttribute('action') || editUrl;
    const targetUrl = action.startsWith('http') ? action : `https://hris.kti.co.id${action}`;

    Logger.info(`Sending POST to ${targetUrl}...`);

    try {
      const response = await fetchWithTimeout(targetUrl, {
        method: 'POST',
        headers: headers,
        body: params.toString(),
        redirect: 'follow'
      }, 15000);

      if (response.url.includes('/login') || response.url.includes('/auth')) {
        Logger.error('Sesi berakhir saat mencoba menyimpan.');
        throw new Error('Sesi berakhir atau gagal otentikasi. Silakan refresh halaman login.');
      }

      if (!response.ok) {
        const errorText = await response.text();
        Logger.error(`Update Master failed (HTTP ${response.status}). Response snippet: ${errorText.substring(0, 200)}`);
        throw new Error('Gagal menyimpan ke server (HTTP ' + response.status + ')');
      }

      Logger.success('Update Master Data successful.');
    } catch (err) {
      if (err.name === 'AbortError') {
        Logger.error('Request timeout (15s). Server HRIS tidak merespon.');
        throw new Error('Server HRIS lambat merespon. Silakan coba lagi.');
      }
      throw err;
    }
  }

  async function checkJkRestoration() {
    const nrp = getPageContext().nrp;
    if (!nrp) return;

    const oldJk = sessionStorage.getItem('qm_jk_to_restore_' + nrp);
    const activeFlow = getAutomationFlow();
    const isScopedRestoreFlow = activeFlow && activeFlow.type === 'distribusi-jk' && activeFlow.meta?.nrp === nrp;
    const finished = isScopedRestoreFlow ? activeFlow.finished : sessionStorage.getItem(STORAGE.AUTO_FINISHED) === 'true';

    if (oldJk && finished) {
      UI.showGlobalLoader('Cleaning up', 'Restoring Master JK...');
      try {
        await saveJkMaster(nrp, oldJk);
        sessionStorage.removeItem('qm_jk_to_restore_' + nrp);
        sessionStorage.setItem('qm_jk_' + nrp, oldJk);
        UI.showResult('success', 'Selesai', 'Kode Jam Kerja master telah dikembalikan.');
        setTimeout(() => window.location.reload(), 1500);
      } catch (e) {
        UI.showResult('danger', 'Gagal Restore', 'Gagal mengembalikan JK: ' + e.message);
        UI.hideGlobalLoader();
      }
    }
  }


  // Removed in favor of unified refreshDistribusiPane
  // async function refreshKKOptionsInPane() { ... }

  // Removed in favor of unified refreshDistribusiPane
  // async function updateKKDropdowns(nrp, bag = '', sek = '') { ... }

  async function fetchKKOptions(nrp) {
    const cached = sessionStorage.getItem('qm_KK_options_' + nrp);
    if (cached) return JSON.parse(cached);

    const emp = await fetchEmployee(nrp);
    if (!emp.found || !emp.id) throw new Error('Data karyawan tidak lengkap.');

    const editUrl = emp.editUrl || employeeEditUrl(nrp, emp.id);

    const html = await hrisFetch(editUrl);
    const doc = parseHTML(html);

    // --- DOM Validation ---
    const title = doc.querySelector('title')?.textContent.trim() || '';
    if (doc.querySelector('input[type="password"]') || title.toLowerCase().includes('login')) {
      throw new Error('Sesi login kedaluwarsa. Silakan login kembali di tab HRIS.');
    }
    // ----------------------
    let select = doc.querySelector('select[name="kode_kalender_kerja"]');

    if (false) {
      let altUrl = '';
      if (editUrl.includes('/editgeneral/')) {
        altUrl = editUrl.replace('/editgeneral/', '/edit/');
      } else if (editUrl.includes('/edit/')) {
        altUrl = editUrl.replace('/edit/', '/editgeneral/');
      }

      if (altUrl) {
        console.log(`[QM INFO] KK select not found in ${editUrl}. Trying alternative URL: ${altUrl}`);
        try {
          const altHtml = await hrisFetch(altUrl);
          const altDoc = parseHTML(altHtml);
          const altSelect = altDoc.querySelector('select[name="kode_kalender_kerja"]');
          if (altSelect) {
            select = altSelect;
          }
        } catch (altErr) {
          console.error(`[QM ERROR] fetchKKOptions fallback error:`, altErr);
        }
      }
    }

    if (!select) {
      const allSelects = Array.from(doc.querySelectorAll('select')).map(s => s.name || s.id || 'no-name');
      console.warn(`[QM WARNING] Kalender Kerja select not found. Selects present:`, allSelects);
      throw new Error('Elemen kode_kalender_kerja tidak ditemukan (Mungkin karyawan outsource / tanpa akses edit).');
    }

    const options = Array.from(select.querySelectorAll('option')).map(opt => ({
      val: opt.value,
      txt: opt.textContent.trim(),
      selected: opt.hasAttribute('selected') || opt.selected
    }));

    sessionStorage.setItem('qm_KK_options_' + nrp, JSON.stringify(options));
    return options;
  }

  async function handleUpdateKKMaster(input = panelReaders.distribusiKk()) {
    const nrp = String(input?.nrp || '').trim();
    const dateVal = String(input?.dateVal || '').trim(); // "YYYY-MM"
    const KK = String(input?.KK || '').trim();

    if (!nrp) { uiAdapter.alert('Harap isi NRP.'); return; }
    if (!KK) { uiAdapter.alert('Pilih Kalender Kerja.'); return; }
    if (!dateVal) { uiAdapter.alert('Pilih periode (bulan/tahun).'); return; }

    const [tahun, bulan] = dateVal.split('-');
    const bagian = input?.bagian;
    const seksi = input?.seksi;
    const grup = input?.grup;

    UI.showGlobalLoader('Processing KK', 'Checking current data...');
    try {
      const emp = await fetchEmployee(nrp);
      if (!emp.found) throw new Error('Data karyawan tidak ditemukan.');

      // Prepare data for distribution (override with UI values if present)
      const distData = {
        ...emp,
        bagian: bagian || emp.bagian,
        seksi: seksi || emp.seksi,
        group: grup || emp.group
      };

      if (emp.KK === KK) {
        Logger.info(`Kalender Kerja sudah sesuai (${KK}), melewati update Master.`);
        UI.setGlobalProgress(40, 'KK match. Proceeding to Distribution...');
      } else {
        UI.setGlobalProgress(30, 'Updating Master Data...');
        await saveKKMaster(nrp, KK);
        sessionStorage.setItem('qm_KK_' + nrp, KK);
        UI.setGlobalProgress(60, 'Master updated. Distributing Calendar...');
      }

      await distributeKkBackground(nrp, bulan, tahun, distData);

      UI.setGlobalProgress(100, 'All done!');
      UI.showResult('success', 'Berhasil', 'Proses KK Selesai.');
      uiAdapter.reload(1500);
    } catch (e) {
      UI.showResult('danger', 'Gagal', 'Error: ' + e.message);
      Logger.error('handleUpdateKKMaster error', e);
    } finally {
      UI.hideGlobalLoader(3000);
    }
  }

  async function saveKKMaster(nrp, KK) {
    const emp = await fetchEmployee(nrp);
    let editUrl = emp.editUrl || employeeEditUrl(nrp, emp.id);

    let html = await hrisFetch(editUrl);
    if (html.includes('id="login-form"') || html.includes('name="login_form"') || html.includes('login-box')) {
      throw new Error('Sesi berakhir. Silakan login kembali.');
    }
    let doc = parseHTML(html);
    let select = doc.querySelector('select[name="kode_kalender_kerja"]');

    if (false) {
      let altUrl = '';
      if (editUrl.includes('/editgeneral/')) {
        altUrl = editUrl.replace('/editgeneral/', '/edit/');
      } else if (editUrl.includes('/edit/')) {
        altUrl = editUrl.replace('/edit/', '/editgeneral/');
      }

      if (altUrl) {
        Logger.info(`Kalender Kerja select not found in ${editUrl}. Trying alternative URL: ${altUrl}`);
        try {
          const altHtml = await hrisFetch(altUrl);
          const altDoc = parseHTML(altHtml);
          const altSelect = altDoc.querySelector('select[name="kode_kalender_kerja"]');
          if (altSelect) {
            editUrl = altUrl;
            html = altHtml;
            doc = altDoc;
            select = altSelect;
          }
        } catch (altErr) {
          Logger.error(`[QM ERROR] saveKKMaster fallback error:`, altErr);
        }
      }
    }
    const form = select?.closest('form') || doc.querySelector('form');

    if (!form) throw new Error('Form edit tidak ditemukan.');

    const params = new URLSearchParams();
    const headers = { 'Content-Type': 'application/x-www-form-urlencoded', 'X-Requested-With': 'XMLHttpRequest' };

    const csrfMeta = doc.querySelector('meta[name="csrf-token"], meta[name="csrf-test-name"]');
    if (csrfMeta) headers['X-CSRF-TOKEN'] = csrfMeta.getAttribute('content');

    form.querySelectorAll('input, select, textarea').forEach(el => {
      if (!el.name || el.type === 'submit' || el.type === 'button') return;
      if ((el.type === 'checkbox' || el.type === 'radio') && !el.checked) return;

      let val = el.value;
      if (el.name === 'kode_kalender_kerja') val = KK;
      params.append(el.name, val);
    });

    const action = form.getAttribute('action') || editUrl;
    const targetUrl = action.startsWith('http') ? action : `https://hris.kti.co.id${action}`;

    const response = await fetch(targetUrl, {
      method: 'POST',
      headers: headers,
      body: params.toString(),
      redirect: 'follow'
    });

    if (!response.ok) throw new Error('Gagal menyimpan (HTTP ' + response.status + ')');
  }

  async function distributeKkBackground(nrp, month, year, emp) {
    const distUrl = ROUTES.DISTRIBUSI_KK;
    const html = await hrisFetch(distUrl);
    if (html.includes('id="login-form"') || html.includes('name="login_form"') || html.includes('login-box')) {
      throw new Error('Sesi berakhir. Silakan login kembali.');
    }
    const doc = parseHTML(html);

    // Improved form detection: Look for a form specifically related to "kalender"
    let form = null;
    const forms = doc.querySelectorAll('form');
    for (const f of forms) {
      if (f.textContent.toLowerCase().includes('kalender') || (f.getAttribute('action') || '').toLowerCase().includes('kalender')) {
        form = f;
        break;
      }
    }
    if (!form) form = doc.querySelector('form'); // Fallback to first form

    if (!form) throw new Error('Form distribusi tidak ditemukan di ' + distUrl);

    const params = new URLSearchParams();
    const headers = { 'Content-Type': 'application/x-www-form-urlencoded', 'X-Requested-With': 'XMLHttpRequest' };

    const csrfMeta = doc.querySelector('meta[name="csrf-token"], meta[name="csrf-test-name"]');
    if (csrfMeta) headers['X-CSRF-TOKEN'] = csrfMeta.getAttribute('content');

    // Fill basic fields from current form values to preserve other defaults
    form.querySelectorAll('input, select, textarea').forEach(el => {
      if (!el.name || el.type === 'submit' || el.type === 'button') return;
      if ((el.type === 'checkbox' || el.type === 'radio') && !el.checked) return;
      params.append(el.name, el.value);
    });

    // Override with our values dynamically based on IDs/Names
    const setVal = (sel, val, fallbackName) => {
      const el = doc.querySelector(sel);
      if (el && el.name) params.set(el.name, val);
      else if (fallbackName) params.set(fallbackName, val);
    };

    setVal('select[name="month"]', month, 'month');
    setVal('select[name="year"]', year, 'year');
    setVal('#kode_bagian', emp.bagian, 'kode_bagian');
    setVal('#kode_seksi', emp.seksi, 'kode_seksi');
    setVal('#kode_group', emp.group, 'kode_group');
    setVal('#nrp_initial_text', nrp, 'nrp_initial_text');
    setVal('#nrp_final_text', nrp, 'nrp_final_text');

    const action = form.getAttribute('action') || distUrl;
    const targetUrl = action.startsWith('http') ? action : `https://hris.kti.co.id${action}`;

    Logger.info('Sending background distribution POST...', { targetUrl, nrp, month });

    const response = await fetch(targetUrl, {
      method: 'POST',
      headers: headers,
      body: params.toString()
    });

    if (!response.ok) throw new Error('Distribusi gagal (HTTP ' + response.status + ')');
    const resText = await response.text();

    // Check for common success indicators
    if (isHrisSuccess(resText)) {
      return true;
    }

    // If not obviously success, log a bit of the response for debugging
    // Clean up text but keep some structure
    const cleanText = resText.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
      .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    const snippet = cleanText.substring(0, 1000);
    Logger.error('Distribusi failed. Response preview', snippet);
    throw new Error('Distribusi gagal: Respon tidak menunjukkan sukses.');
  }

  async function autoDistKK() {
    const urlParams = getCurrentQueryParams();
    if (!urlParams.get('qm_auto_dist_KK')) {
      const activeFlow = getAutomationFlow();
      const scopedKkFlow = activeFlow && activeFlow.type === 'dist-kk' ? activeFlow : null;
      // Result page check
      if (((scopedKkFlow && scopedKkFlow.finished) || (!activeFlow && sessionStorage.getItem(STORAGE.AUTO_FINISHED) === 'true')) && isDistribusiKalenderPagePath()) {
        const pageText = document.body.textContent;
        if (pageText.includes('Distribution Process Completed') || document.querySelector('.alert-success')) {
          UI.showResult('success', 'Selesai', 'Distribusi Kalender Kerja Selesai.');
          const returnUrl = scopedKkFlow?.returnUrl || sessionStorage.getItem(STORAGE.RETURN_URL);
          if (returnUrl) {
            setTimeout(() => {
              window.location.href = returnUrl;
            }, 1500);
          }
        }
      }
      return;
    }

    const cleanUrl = new URL(window.location);
    cleanUrl.searchParams.delete('qm_auto_dist_KK');
    window.history.replaceState({}, document.title, cleanUrl.toString());

    const nrp = urlParams.get('nrp');
    const month = urlParams.get('month');

    if (!nrp || !month) return;

    UI.showGlobalLoader('Auto Dist KK', 'Memulai Distribusi Kalender...');

    try {
      const emp = await fetchEmployee(nrp);
      if (!emp.found) throw new Error('Data karyawan tidak ditemukan.');

      // 1. Month
      const selMonth = document.querySelector('select[name="month"]');
      if (selMonth) setField(selMonth, month);

      // 2. Sequential Dropdowns: Bagian -> Seksi
      UI.setGlobalProgress(40, 'Menyesuaikan Bagian...');
      awaitDropdown('#kode_bagian', emp.bagian, () => {
        UI.setGlobalProgress(70, 'Menyesuaikan Seksi...');
        awaitDropdown('#kode_seksi', emp.seksi, () => {
          UI.setGlobalProgress(85, 'Mengisi NRP...');

          // 3. NRP Initial & Final
          const nrpInit = document.getElementById('nrp_initial_text');
          const nrpFinal = document.getElementById('nrp_final_text');
          if (nrpInit) setField(nrpInit, nrp);
          if (nrpFinal) setField(nrpFinal, nrp);

          // 4. Submit
          const btnSubmit = document.getElementById('btnSubmit');
          if (btnSubmit) {
            UI.setGlobalProgress(95, 'Submitting...');
            const activeFlow = getAutomationFlow();
            if (activeFlow && activeFlow.type === 'dist-kk') {
              markAutomationFlowFinished(activeFlow.id);
            } else {
              sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
            }
            setTimeout(() => btnSubmit.click(), 1000);
          }
        });
      });
    } catch (e) {
      UI.showResult('danger', 'Auto KK Gagal', e.message);
      UI.hideGlobalLoader(3000);
    }
  }

  /* ============================================================
   * 13. EVENT HANDLERS & PANEL
   * ============================================================ */

  const panelReaders = Object.freeze({
    lookup() {
      return {
        nrp: uiAdapter.getValue('lookupNrp', { trim: true }),
        bulan: uiAdapter.getValue('globalMonth', { trim: true }),
        tahun: uiAdapter.getValue('globalYear', { trim: true })
      };
    },

    karyawanSearch() {
      return {
        query: uiAdapter.getValue('karyawanSearchInput', { trim: true }),
        bulan: uiAdapter.getValue('globalMonth', { trim: true }) || (new Date().getMonth() + 1),
        tahun: uiAdapter.getValue('globalYear', { trim: true }) || new Date().getFullYear()
      };
    },

    attendanceCheck() {
      return {
        nrp: uiAdapter.getValue('attendanceCheckNrp', { trim: true }),
        startDate: uiAdapter.getValue('attendanceCheckStartDate', { trim: true }),
        endDate: uiAdapter.getValue('attendanceCheckEndDate', { trim: true })
      };
    },

    spklOnline() {
      return {
        nrp: uiAdapter.getValue('spklOnlineNrp', { trim: true }),
        dateInput: uiAdapter.getValue('spklOnlineDate', { trim: true })
      };
    },

    spklPageLink() {
      return {
        nrp: uiAdapter.getValue('spklPageNrp', { trim: true }),
        bulan: uiAdapter.getValue('spklPageMonth', { trim: true }) || uiAdapter.getValue('globalMonth', { trim: true }),
        tahun: uiAdapter.getValue('globalYear', { trim: true }) || new Date().getFullYear()
      };
    },

    spklCheck() {
      return {
        nrp: uiAdapter.getValue('spklPageNrp', { trim: true }),
        startDate: uiAdapter.getValue('spklPageStartDate', { trim: true }),
        endDate: uiAdapter.getValue('spklPageEndDate', { trim: true })
      };
    },

    spklInlineEdit() {
      return {
        ot: uiAdapter.getValue('spklEditOt', { trim: true }),
        jamAwal: uiAdapter.getValue('spklEditJamAwal', { trim: true }),
        jamAkhir: uiAdapter.getValue('spklEditJamAkhir', { trim: true }),
        jamOt: uiAdapter.getValue('spklEditJamOt', { trim: true })
      };
    },

    attendanceInput() {
      return {
        nrp: uiAdapter.getValue('attendanceInputNrp', { trim: true }),
        tgl: uiAdapter.getValue('attendanceInputDate', { trim: true }),
        jam: uiAdapter.getValue('attendanceInputTime', { trim: true }),
        status: uiAdapter.getValue('attendanceInputStatus', { trim: true })
      };
    },

    distribusiJk() {
      return {
        nrp: uiAdapter.getValue('distribusiNrp', { trim: true }),
        useDistribusi: uiAdapter.checked('distribusiJkUseDistribusi'),
        jk: uiAdapter.getValue('distribusiJkSelect', { trim: true }),
        date: uiAdapter.getValue('distribusiJkTargetDate', { trim: true }),
        dateEnd: uiAdapter.getValue('distribusiJkTargetDateEnd', { trim: true }),
        shift: uiAdapter.getValue('distribusiJkTargetShift', { trim: true })
      };
    },

    distribusiSubsi() {
      return {
        nrp: getPageContext().nrp,
        useDistribusi: uiAdapter.checked('distribusiSubsiUseDistribusi'),
        jk: uiAdapter.getValue('distribusiSubsiJkSelect', { trim: true }),
        tglAwal: uiAdapter.getValue('distribusiSubsiDateStart', { trim: true }),
        tglAkhir: uiAdapter.getValue('distribusiSubsiDateEnd', { trim: true }),
        shift: uiAdapter.getValue('distribusiSubsiShift', { trim: true }),
        bagian: uiAdapter.getValue('distribusiSubsiBagian', { trim: true }),
        seksi: uiAdapter.getValue('distribusiSubsiSeksi', { trim: true }),
        grup: uiAdapter.getValue('distribusiSubsiGrup', { trim: true })
      };
    },

    distribusiKk() {
      return {
        nrp: uiAdapter.getValue('distribusiKkNrp', { trim: true }),
        dateVal: uiAdapter.getValue('distribusiKkDate', { trim: true }),
        KK: uiAdapter.getValue('distribusiKkSelect', { trim: true }),
        bagian: uiAdapter.getValue('distribusiKkBagian', { trim: true }),
        seksi: uiAdapter.getValue('distribusiKkSeksi', { trim: true }),
        grup: uiAdapter.getValue('distribusiKkGrup', { trim: true })
      };
    },

    karyawanSave(key) {
      return {
        key,
        nextJk: uiAdapter.findKaryawanSelect(key, 'jk')?.value || '',
        nextKk: uiAdapter.findKaryawanSelect(key, 'kk')?.value || ''
      };
    },

    hadirInlineEdit() {
      return {
        jamMasuk: uiAdapter.getValue('hadirEditJamMasuk', { trim: true }),
        jamKeluar: uiAdapter.getValue('hadirEditJamKeluar', { trim: true }),
        status: uiAdapter.getValue('hadirEditStatus', { trim: true })
      };
    }
  });

  function openPanel() {
    state.isOpen = true;
    document.body.classList.add('qm-no-scroll');
    uiAdapter.addClass('panelShell', 'qm-open');
    setTimeout(() => {
      uiAdapter.requestPrimaryFocus();
    }, 250);
  }

  function closePanel() {
    state.isOpen = false;
    document.body.classList.remove('qm-no-scroll');
    uiAdapter.removeClass('panelShell', 'qm-open');
  }

  function togglePanel() {
    state.isOpen ? closePanel() : openPanel();
  }

  function getKaryawanSearchInput() {
    return uiAdapter.get('karyawanSearchInput');
  }

  function prefillKaryawanSearch(autoSearch = false) {
    const input = getKaryawanSearchInput();
    if (!input) return;

    const ctxNrp = getPageContext().nrp || uiAdapter.getValue('lookupNrp', { trim: true }) || '';
    if (!ctxNrp) return;

    if (!input.value.trim()) input.value = ctxNrp;
    if (autoSearch && !state.karyawanLoading && !state.karyawanQuery && isValidNrp(input.value.trim())) {
      CEK_NRP.searchByQuery(panelReaders.karyawanSearch());
    }
  }

  async function handleKaryawanSearch(input = panelReaders.karyawanSearch()) {
    const query = String(input?.query || '').trim();
    if (!query) {
      UI.showResult('warning', 'Data Tidak Lengkap', 'Masukkan NRP atau nama karyawan terlebih dahulu.');
      return;
    }

    state.karyawanQuery = query;
    state.karyawanLoading = true;
    state.karyawanError = '';
    state.karyawanResults = [];
    resetKaryawanPanels();
    renderKaryawanResults();

    if (isValidNrp(query)) {
      const bulan = input?.bulan || (new Date().getMonth() + 1);
      const tahun = input?.tahun || new Date().getFullYear();
      syncGlobalInputs(query, bulan, tahun);
    }

    // Start progress bar in footer
    UI.startProgress('Mencari Karyawan...', 1200);

    try {
      state.karyawanResults = await searchEmployees(query);
    } catch (e) {
      state.karyawanError = e.message || 'Gagal mencari data karyawan.';
    } finally {
      state.karyawanLoading = false;
      UI.endProgress();
      renderKaryawanResults();
    }
  }

  function handleNrpLookup(input = panelReaders.lookup()) {
    const nrp = String(input?.nrp || '').trim();
    const bulan = String(input?.bulan || '').trim();
    if (!nrp) { uiAdapter.alert('NRP is empty'); return; }
    if (!bulan) { UI.showResult('warning', 'Data Tidak Lengkap', 'Silakan masukkan Bulan terlebih dahulu.'); return; }
    if (!/^\d+$/.test(nrp) || (nrp.length !== 4 && nrp.length !== 8)) { UI.showResult('warning', 'Format Tidak Valid', 'Hanya menerima 4 dan 8 angka NRP'); return; }
    UI.setLoading(true);
    sessionStorage.setItem(STORAGE.AUTO_NRP, nrp);
    sessionStorage.setItem(STORAGE.AUTO_BULAN, bulan);
    const year = input?.tahun || new Date().getFullYear();
    uiAdapter.openUrl(attendanceUrl(bulan, year, nrp), '_self');
  }

  let manualSidebarOverride = false;
  function enforceSidebar() {
    if (!alwaysCollapseMenu || manualSidebarOverride) return;
    if (!document.body.classList.contains('enlarged')) {
      document.body.classList.add('enlarged');
      document.body.classList.remove('sidebar-enable');
    }
  }

  function handleExportAnomali() {
    if (state.anomalies.length === 0) { uiAdapter.alert('Tidak ada anomali untuk diekspor.'); return; }
    const wsData = [['Tanggal', 'Kolom', 'Pesan Anomali']];
    const sorted = [...state.anomalies].sort((a, b) => parseInt(a.tgl) - parseInt(b.tgl));
    sorted.forEach(a => wsData.push([a.tgl, a.col, a.msg]));
    if (typeof XLSX !== 'undefined') {
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Anomali');
      const month = uiAdapter.getValue('globalMonth', { trim: true }) || (new Date().getMonth() + 1);
      XLSX.writeFile(wb, `Anomali_Bulan_${month}.xlsx`);
    } else {
      uiAdapter.alert('Library XLSX gagal dimuat. Harap periksa koneksi atau header script.');
    }
  }

  function handleToggleAnomalyGroup(e) {
    if (e.target.closest('.qm-fix-dot') || e.target.closest('.qm-batch-fix-btn')) return;

    const content = this.nextElementSibling;
    if (content && content.classList.contains('qm-anomaly-group-content')) {
      const expanded = this.classList.toggle('expanded');
      // Gunakan class 'qm-content-open' pada content agar tidak bentrok dengan
      // selector CSS '.expanded' yang bisa bocor ke elemen lain
      content.classList.toggle('qm-content-open', expanded);

      const tgl = this.dataset.tgl;
      if (tgl) {
        if (expanded) state.expandedAnomalyGroups.add(String(tgl));
        else state.expandedAnomalyGroups.delete(String(tgl));
      }
    }
  }

  function handleSpklOnlineCheck(input = panelReaders.spklOnline()) {
    const nrp = String(input?.nrp || '').trim();
    const dateInput = String(input?.dateInput || '').trim();
    if (!nrp || !dateInput) { uiAdapter.alert('Harap isi NRP dan Tanggal.'); return; }

    const d = new Date(dateInput);
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();

    const url = ROUTES.SPKL_ONLINE(year, month, day, day, nrp);
    uiAdapter.openUrl(url, '_blank');
  }

  function handleSpklPageCheck(input = panelReaders.spklPageLink()) {
    const nrp = String(input?.nrp || '').trim();
    const bulan = String(input?.bulan || '').trim();

    if (!nrp || !bulan) {
      UI.showResult('warning', 'Data Belum Lengkap', 'Silakan isi NRP dan Bulan terlebih dahulu.');
      return;
    }
    if (!isValidNrp(nrp)) {
      UI.showResult('warning', 'NRP Tidak Valid', 'Gunakan 4 digit (Reguler) atau 8 digit (OS).');
      return;
    }

    const bulanInt = parseInt(bulan, 10);
    if (isNaN(bulanInt) || bulanInt < 1 || bulanInt > 12) {
      UI.showResult('warning', 'Bulan Tidak Valid', 'Silakan pilih bulan.');
      return;
    }

    const tahun = input?.tahun || new Date().getFullYear();
    uiAdapter.openUrl(spklListUrl(nrp, bulanInt, tahun), '_blank');
  }

  async function handleAttendanceCheckNrp(input = panelReaders.attendanceCheck()) {
    const nrp = String(input?.nrp || '').trim();
    const startDate = String(input?.startDate || '').trim();
    const endDate = String(input?.endDate || '').trim();

    if (!nrp || !startDate || !endDate) {
      uiAdapter.alert('Harap isi NRP, Start date, dan End date.');
      return;
    }
    if (!/^\d+$/.test(nrp) || (nrp.length !== 4 && nrp.length !== 8)) {
      uiAdapter.alert('NRP harus 4 atau 8 digit angka.');
      return;
    }
    if (parseHrisDate(startDate)?.getTime() > parseHrisDate(endDate)?.getTime()) {
      uiAdapter.alert('Start date tidak boleh lebih besar dari End date.');
      return;
    }

    const requestKey = `${nrp}|${startDate}|${endDate}|${Date.now()}`;
    state.attendanceCheck = {
      ...createEmptyAttendanceCheck(),
      loading: true,
      requestKey
    };
    renderAttendanceCheckResult();

    try {
      const [summary, emp] = await Promise.all([
        fetchBarcodeAttendanceSummary(nrp, startDate, endDate),
        fetchEmployee(nrp).catch(() => ({ found: false }))
      ]);

      if (state.attendanceCheck.requestKey !== requestKey) return;

      state.attendanceCheck = {
        ...createEmptyAttendanceCheck(),
        summary: {
          ...summary,
          employee: emp?.found ? emp : null
        }
      };
    } catch (e) {
      if (state.attendanceCheck.requestKey !== requestKey) return;
      state.attendanceCheck = {
        ...createEmptyAttendanceCheck(),
        error: e.message || 'Gagal memuat ringkasan kehadiran.'
      };
    }

    renderAttendanceCheckResult();
  }

  async function handleSpklCheckNrp(input = panelReaders.spklCheck()) {
    const nrp = String(input?.nrp || '').trim();
    const startDate = String(input?.startDate || '').trim();
    const endDate = String(input?.endDate || '').trim();

    if (!nrp || !startDate || !endDate) {
      UI.showResult('warning', 'Data Belum Lengkap', 'Silakan isi NRP dan Rentang Tanggal.');
      return;
    }
    if (!isValidNrp(nrp)) {
      UI.showResult('warning', 'NRP Tidak Valid', 'Gunakan 4 digit (Reguler) atau 8 digit (OS).');
      return;
    }

    const requestKey = `${nrp}|${startDate}|${endDate}|${Date.now()}`;
    state.spklCheck = {
      ...createEmptySpklCheck(),
      loading: true,
      requestKey
    };
    renderSpklCheckResult();

    try {
      const summary = await fetchSpklSummary(nrp, startDate, endDate);
      if (state.spklCheck.requestKey !== requestKey) return;

      state.spklCheck = {
        ...createEmptySpklCheck(),
        summary
      };
    } catch (e) {
      if (state.spklCheck.requestKey !== requestKey) return;
      state.spklCheck = {
        ...createEmptySpklCheck(),
        error: e.message || 'Gagal memuat data SPKL.'
      };
    }

    renderSpklCheckResult();
  }

  function toggleSpklRowInputs(checkbox) {
    const row = checkbox.closest('.qm-hadir-check-detail-item');
    if (!row) return;
    const inlineInputs = row.querySelector('.qm-spkl-inline-inputs');
    if (inlineInputs) {
      if (checkbox.checked) {
        inlineInputs.classList.remove('qm-hidden');
      } else {
        inlineInputs.classList.add('qm-hidden');
      }
    }
  }

  function startSpklPageLoop() {
    const checkboxes = document.querySelectorAll('.qm-spkl-batch-cb:checked');
    if (checkboxes.length === 0) {
      uiAdapter.alert('Pilih setidaknya satu baris SPKL untuk diedit.');
      return;
    }

    const tasks = [];
    const current = state.spklCheck;
    if (!current || !current.summary) return;

    checkboxes.forEach(cb => {
      const row = cb.closest('.qm-hadir-check-detail-item');
      if (!row) return;

      const idx = parseInt(cb.dataset.index, 10);
      if (isNaN(idx)) return;

      const entry = current.summary.entries[idx];
      if (!entry) return;

      const editAction = (entry.actions || []).find(act => /edit/i.test(act.text));
      if (!editAction) return;

      const inlineInputs = row.querySelector('.qm-spkl-inline-inputs');
      if (!inlineInputs) return;

      const jenis = inlineInputs.querySelector('.qm-spkl-inline-jenis')?.value || '';
      const shift = inlineInputs.querySelector('.qm-spkl-inline-shift')?.value || '';
      const awal = inlineInputs.querySelector('.qm-spkl-inline-awal')?.value || '';
      const akhir = inlineInputs.querySelector('.qm-spkl-inline-akhir')?.value || '';
      const tambahan = inlineInputs.querySelector('.qm-spkl-inline-tambahan')?.value || '';

      tasks.push({
        index: idx,
        modalTarget: editAction.modalTarget,
        href: toAbsoluteHrisUrl(editAction.href),
        jenis,
        shift,
        awal,
        akhir,
        tambahan
      });
    });

    if (tasks.length === 0) {
      uiAdapter.alert('Pilih setidaknya satu baris valid.');
      return;
    }

    const nrp = current.summary.nrp;
    const bulan = current.summary.bulan || uiAdapter.getValue('spklPageMonth') || (new Date().getMonth() + 1);
    const tahun = current.summary.tahun || new Date().getFullYear();
    const baseUrl = spklListUrl(nrp, bulan, tahun);

    createAutomationFlow('spkl-batch', window.location.href, {
      tasks,
      currentIndex: 0,
      baseUrl
    });

    UI.showResult('info', 'Otomasi Dimulai', `Memproses ${tasks.length} item...`);
    setTimeout(() => {
      window.location.href = baseUrl;
    }, 300);
  }

  async function resumeSpklPageLoop() {
    const flow = getAutomationFlow();
    if (!flow || flow.type !== 'spkl-batch') return;

    const tasks = flow.meta.tasks;
    const idx = flow.meta.currentIndex;
    const task = tasks[idx];

    if (!task) {
      finishAutomationFlow(flow.id);
      return;
    }

    const currentUrl = window.location.href.split('?')[0];
    const targetBase = (flow.meta.baseUrl || '').split('?')[0];

    if (currentUrl !== targetBase && !window.location.href.includes('/spkl')) {
      window.location.href = flow.meta.baseUrl;
      return;
    }

    UI.showGlobalLoader('Otomasi', `Mengolah baris ${idx + 1}/${tasks.length}...`);

    const modal = document.querySelector('.modal.show, .modal.in, [role="dialog"].show') ||
      document.querySelector('.modal-content');
    const isVisible = modal && (modal.offsetWidth > 0 || modal.offsetHeight > 0);

    if (isVisible) {
      const inputs = {
        jenis_ot: modal.querySelector('select[name*="jenis_ot"], select[id*="jenis_ot"], #jenis_ot_edit'),
        shift: modal.querySelector('select[name*="shift"], select[id*="shift"], #shift'),
        awal: modal.querySelector('input[name*="jam_awal_ot"], input[id*="jam_awal_ot"], input[type="time"]'),
        akhir: modal.querySelector('input[name*="jam_akhir_ot"], input[id*="jam_akhir_ot"], input[type="time"]:nth-of-type(2)'),
        tambahan: modal.querySelector('input[name*="tambahan_jam_ot"], input[id*="tambahan_jam_ot"]')
      };

      if (inputs.awal && task.awal) { inputs.awal.value = task.awal; inputs.awal.dispatchEvent(new Event('change', { bubbles: true })); }
      if (!inputs.akhir) {
        const timeInputs = modal.querySelectorAll('input[type="time"]');
        if (timeInputs.length >= 2) inputs.akhir = timeInputs[1];
      }
      if (inputs.akhir && task.akhir) { inputs.akhir.value = task.akhir; inputs.akhir.dispatchEvent(new Event('change', { bubbles: true })); }

      if (inputs.jenis_ot && task.jenis) { inputs.jenis_ot.value = task.jenis; inputs.jenis_ot.dispatchEvent(new Event('change', { bubbles: true })); }
      if (inputs.shift && task.shift) { inputs.shift.value = task.shift; inputs.shift.dispatchEvent(new Event('change', { bubbles: true })); }
      if (inputs.tambahan && task.tambahan) { inputs.tambahan.value = task.tambahan; inputs.tambahan.dispatchEvent(new Event('change', { bubbles: true })); }

      setTimeout(() => {
        const buttons = Array.from(modal.querySelectorAll('button, input[type="submit"]'));
        const submitBtn = buttons.find(b =>
          b.type === 'submit' ||
          b.classList.contains('btn-primary') ||
          /submit|simpan|save/i.test(b.textContent)
        );

        flow.meta.currentIndex++;
        sessionStorage.setItem(STORAGE.AUTO_FLOW, JSON.stringify(flow));

        if (submitBtn) {
          submitBtn.click();
        } else {
          Logger.error('Tombol Submit tidak ditemukan di modal SPKL.');
          window.location.reload();
        }
      }, 200);
    } else {
      let editBtn = null;
      if (task.modalTarget) {
        editBtn = document.querySelector(`button[data-target="${task.modalTarget}"], [data-target="${task.modalTarget}"]`);
      }

      if (editBtn) {
        editBtn.click();
        setTimeout(resumeSpklPageLoop, 300);
      } else {
        Logger.error('Tombol edit tidak ditemukan di halaman SPKL.');
        flow.meta.currentIndex++;
        sessionStorage.setItem(STORAGE.AUTO_FLOW, JSON.stringify(flow));
        window.location.reload();
      }
    }
  }

  function handleInputHadir(input = panelReaders.attendanceInput()) {
    const nrp = String(input?.nrp || '').trim();
    const tgl = String(input?.tgl || '').trim();
    const jam = String(input?.jam || '').trim();
    const status = String(input?.status ?? '').trim();

    if (!nrp || !tgl || !jam || status === "") {
      uiAdapter.alert('Harap isi semua field (NRP, Tanggal, Jam, dan Status).');
      return;
    }

    createAutomationFlow('hadir-single', window.location.href, { nrp, tgl, status });

    const data = { nrp, tgl, jam, status };
    sessionStorage.setItem(STORAGE.INPUT_HADIR, JSON.stringify(data));

    const targetUrl = absenCreateUrl(nrp);
    uiAdapter.openUrl(targetUrl, '_blank');
  }

  function handleDistribusi(input = panelReaders.distribusiJk()) {
    handleSaveJkChange(input);
  }

  function handleKeydownAnomalyGroup(e) {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.click(); }
  }

  function handleFixDotClick(e) {
    e.stopPropagation();
    handleFixClick(
      this.getAttribute('data-fix-link'),
      this.getAttribute('data-fix-date'),
      this.getAttribute('title'),
      this.getAttribute('data-full-date')
    );
  }

  function handleTabClick() {
    activatePane(this.getAttribute('data-pane'));
  }

  function handleInputBulan() {
    const val = parseInt(this.value);
    if (isNaN(val)) return;
    if (val < 1) this.value = 1;
    else if (val > 12) this.value = 12;
  }

  function handleDocumentClick(e) {
    if (state.isOpen && !e.target.closest('#qm-panel, #qm-fab')) closePanel();
  }

  function handleRecordShortcut() {
    isRecordingShortcut = true;
    this.textContent = 'Tunggu...';
    uiAdapter.value('shortcutInput', 'Tekan tombol...');
  }

  function handleBatchNrpClick(e) {
    e.preventDefault();
    var nrp = this.getAttribute('data-nrp');
    if (!nrp) return;
    var url = attendanceUrl(state.batchBulan, state.batchTahun, nrp);
    window.location.href = url;
  }

  function handleBatchFixClick(e) {
    e.stopPropagation();
    handleFixClick(
      this.getAttribute('data-fix-link'),
      this.getAttribute('data-fix-date'),
      this.getAttribute('title'),
      this.getAttribute('data-full-date')
    );
  }

  function handleKeydownDocument(e) {
    if (isRecordingShortcut) {
      if (e.key === 'Escape') {
        e.preventDefault();
        isRecordingShortcut = false;
        const btn = uiAdapter.get('shortcutRecordButton');
        if (btn) btn.textContent = 'Ubah';
        const input = uiAdapter.get('shortcutInput');
        if (input) input.value = shortcutKey;
        return;
      }
      e.preventDefault();
      const keys = [];
      if (e.ctrlKey) keys.push('Ctrl');
      if (e.altKey) keys.push('Alt');
      if (e.shiftKey) keys.push('Shift');
      if (['Control', 'Alt', 'Shift', 'Meta'].includes(e.key)) return;
      keys.push(e.key.length === 1 ? e.key.toUpperCase() : e.key);
      shortcutKey = keys.join('+');
      const input = uiAdapter.get('shortcutInput');
      if (input) input.value = shortcutKey;
      const btn = uiAdapter.get('shortcutRecordButton');
      if (btn) btn.textContent = 'Ubah';
      GM_setValue('qm_shortcut', shortcutKey);
      isRecordingShortcut = false;
      return;
    }

    if (e.key === 'Escape' && state.isOpen) {
      closePanel();
      return;
    }

    if (shortcutKey) {
      const parts = shortcutKey.split('+');
      const key = parts[parts.length - 1].toLowerCase();
      if (e.key.toLowerCase() === key && e.ctrlKey === parts.includes('Ctrl') && e.shiftKey === parts.includes('Shift') && e.altKey === parts.includes('Alt')) {
        const targetTag = e.target.tagName.toLowerCase();
        if ((targetTag === 'input' || targetTag === 'textarea') && parts.length === 1) return;
        e.preventDefault();
        togglePanel();
      }
    }
  }

  function handleDebugToggle() {
    state.debug = this.checked;
    GM_setValue('qm_debug', state.debug);
    Logger.info(`Debug mode ${state.debug ? 'diaktifkan' : 'dimatikan'}`);
  }

  function handleShowLogs() {
    const container = uiAdapter.get('logContainer');
    const btn = uiAdapter.get('showLogsButton');
    if (container && btn) {
      const isHidden = container.classList.toggle('qm-hidden');
      if (!isHidden) {
        btn.classList.add('qm-active');
        renderLogs();
        btn.querySelector('span').textContent = 'Sembunyikan Log';
      } else {
        btn.classList.remove('qm-active');
        btn.querySelector('span').textContent = 'Lihat Log Aktivitas';
      }
    }
  }

  function handleClearLogs() {
    if (uiAdapter.confirm('Bersihkan semua riwayat log aktivitas?')) {
      state.batchLogs = [];
      renderLogs();
    }
  }

  function handleExportLogs() {
    if (state.batchLogs.length === 0) {
      uiAdapter.alert('Tidak ada log untuk diekspor.');
      return;
    }
    const content = state.batchLogs.map(l => {
      const level = (l.level || l.type || 'info').toUpperCase();
      return `[${l.time}] [${level}] ${l.msg}`;
    }).join('\n');

    const header = `# HRIS Quick Menu Activity Log\nExported: ${new Date().toLocaleString()}\n\n`;
    const blob = new Blob([header + '```\n' + content + '\n```'], { type: 'text/markdown' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `hris_log_${new Date().toISOString().slice(0, 10)}.md`;
    a.click();
    URL.revokeObjectURL(url);
    Logger.info('Log diekspor ke format Markdown (.md).');
  }

  function getPrimaryCardButton(card) {
    const buttons = Array.from(card.querySelectorAll('button.qm-btn'))
      .filter(btn => !btn.disabled && !btn.classList.contains('qm-hidden'));
    if (buttons.length === 0) return null;

    return buttons.find(btn => btn.classList.contains('qm-btn-primary') || btn.classList.contains('qm-btn-success')) || buttons[0];
  }

  function initKeyboardNavigation() {
    // Enter key logic for various inputs
    on('keydown', '.qm-input, .qm-select, .qm-textarea', function (e) {
      if (e.key !== 'Enter') return;
      if (this.tagName === 'TEXTAREA' && !e.ctrlKey) return; // Allow newlines in textarea unless Ctrl+Enter

      if (!this.closest('.qm-pane')) return;
      const card = this.closest('.qm-card');
      if (!card) return;

      // Prefer semantic action buttons, then fall back to the first visible button in the card.
      const primaryBtn = getPrimaryCardButton(card);
      if (!primaryBtn || primaryBtn.disabled) return;

      // Validate required fields in this card
      const inputs = card.querySelectorAll('.qm-input:not([readonly]), .qm-select');
      let allFilled = true;
      inputs.forEach(input => {
        if (!input.value.trim() && !input.classList.contains('qm-input-optional')) {
          allFilled = false;
        }
      });

      if (allFilled) {
        e.preventDefault();
        primaryBtn.click();
      }
    });

    // Space/Enter for accordions
    on('keydown', '.qm-accordion-header', function (e) {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        this.click();
      }
    });
  }

  /* ============================================================
   * 14. DOMAIN MODULES
   * ============================================================ */

  const CEK_NRP = Object.freeze({
    runLookup: handleNrpLookup,
    searchByQuery: handleKaryawanSearch,
    lookup: handleNrpLookup,
    search: handleKaryawanSearch,
    prefillSearch: prefillKaryawanSearch,
    renderResults: renderKaryawanResults,
    renderDetail: renderKaryawanDetail,
    renderEditor: renderKaryawanEditor,
    renderExpandedPanel: renderKaryawanExpandedPanel,
    toggleDetail: toggleKaryawanDetail,
    toggleEditor: toggleKaryawanEditor,
    saveEditor: handleKaryawanSaveEdit,
    loadDetail: loadKaryawanDetail,
    loadEditor: loadKaryawanEditor,
    searchEmployees,
  });

  const SPKL = Object.freeze({
    highlightPendingDate: spklHighlight,
    autoFillTargetPage,
    autoFillEditPage: autoFillSpklEdit,
    renderCheckResult: renderSpklCheckResult,
    fetchSummary: fetchSpklSummary,
    parseSummary: parseSpklSummary,
    openOnlineForDate: handleSpklOnlineCheck,
    openPageForMonth: handleSpklPageCheck,
    checkByDateRange: handleSpklCheckNrp,
    openOnlineCheck: handleSpklOnlineCheck,
    openPageCheck: handleSpklPageCheck,
    checkByNrp: handleSpklCheckNrp,
    toggleRowInputs: toggleSpklRowInputs,
    startPageLoop: startSpklPageLoop,
    resumePageLoop: resumeSpklPageLoop,
    startBatchProcess: runSpklBatchProcess,
    resumeBatch: checkSpklBatchResume,
    startManyNrpBatch: runSpklManyNrpBatch,
    processBackgroundSingle: processSpklBackgroundSingle,
    continueBatch: _continueSpklBatch,
    processManyNrpPage: _processManyNrpPage,
    checkOnlineStatuses: checkSPKLOnline,
    runBackgroundQueue: runSpklBackgroundQueue,
  });

  const KEHADIRAN = Object.freeze({
    autoSearchPage: autoBarcodeSearchPage,
    autoClickAddData,
    autoInput: autoInputHadir,
    checkByRange: handleAttendanceCheckNrp,
    submitSingle: handleInputHadir,
    renderCheckResult: renderAttendanceCheckResult,
    checkByNrp: handleAttendanceCheckNrp,
    handleInput: handleInputHadir,
    fetchSpklSummary,
    startManyNrpBatch: runHadirManyNrpBatch,
    resumeManyNrpBatch: checkHadirBatchResume,
    processManyNrpPage: _processHadirManyNrpPage,
    startBulanBatch: runHadirBulanBatch,
    resumeBulanBatch: checkHadirBulanResume,
    fetchBarcodeSummary: fetchBarcodeAttendanceSummary,
    parseBarcodeSummary: parseBarcodeAttendanceSummary,
    openBatchEdit: openHadirBatchEdit,
    closeBatchEdit: closeHadirBatchEdit,
    toggleRowInputs: toggleHadirRowInputs,
    startPageLoop: startHadirPageLoop,
    resumePageLoop: resumeHadirPageLoop,
  });

  const DISTRIBUSI = Object.freeze({
    initChangeEvents: initJkChangeEvents,
    resolveSyncContext,
    syncGlobalInputs,
    refreshGlobalData,
    refreshOptions: refreshDistribusiOptions,
    submitJkChange: handleSaveJkChange,
    submitSubsi: handleDistribusiSubsi,
    submitKk: handleUpdateKKMaster,
    handleDistribusi,
    handleDistribusiSubsi,
    saveJkChange: handleSaveJkChange,
    executeBackground: executeBackgroundDistribusi,
    autoDistribusi,
    autoDistribusiSubsi,
    attachListeners: attachDistribusiListeners,
    updateDropdowns: updateDistribusiDropdowns,
    applyDropdowns: applyDistDropdowns,
    checkJkRestoration,
    fetchJkOptions,
    saveJkMaster,
    handleUpdateKKMaster,
    fetchKKOptions,
    saveKKMaster,
    distributeKkBackground,
    autoDistKK,
  });

  const ANOMALI = Object.freeze({
    detect: detectAnomalies,
    scanAttendance,
    validateShiftRow,
    validateOvertime,
    applyMark,
    render: renderAnomalies,
    exportCurrent: handleExportAnomali,
    startBatchCheck: startBatchAnomalyCheck,
    cancelBatchCheck: handleBatchCancel,
    clearBatchResults: handleBatchClear,
    processBatchWorker,
    finishBatch,
    pushBatchResult,
    updateBatchProgress,
    renderBatchResults,
    exportBatchResults,
    handleBatchNrpClick,
    handleBatchFixClick,
    handleFixClick,
    toggleGroup: handleToggleAnomalyGroup,
    checkBarcodeMangkir,
  });

  /* ============================================================
   * 15. INITIALIZATION
   * ============================================================ */

  function validateDomStructure(options) {
    const silent = options && options.silent;
    const missing = [];

    if (isAttendancePagePath()) {
      if (!document.querySelector('table tbody tr')) {
        missing.push('table tbody tr (tabel kehadiran)');
      } else {
        const headers = Array.from(document.querySelectorAll('table th'));
        const headerTexts = headers.map(th => th.textContent.trim().toLowerCase());
        const requiredHeaders = ['tanggal', 'msk', 'klr'];
        const missingHeaders = requiredHeaders.filter(h => !headerTexts.some(t => t.includes(h)));
        if (missingHeaders.length > 0) {
          missing.push('kolom header: ' + missingHeaders.join(', ') + ' (tabel kehadiran)');
        }
      }
    }
    if (isBarcodeCreatePagePath()) {
      if (!document.querySelector('form')) missing.push('form (barcode create)');
    }
    if (isDistribusiKalenderPagePath()) {
      if (!document.querySelector('form')) missing.push('form (distribusi kalender)');
    }
    if (isSpklPagePath()) {
      if (!document.querySelector('table')) {
        missing.push('table (SPKL)');
      } else {
        const headers = Array.from(document.querySelectorAll('table th'));
        const headerTexts = headers.map(th => th.textContent.trim().toLowerCase());
        const hasDateCol = headerTexts.some(t => t.includes('tanggal') || t === 'tgl');
        if (!hasDateCol) {
          missing.push('kolom tanggal (tabel SPKL)');
        }
      }
    }

    const valid = missing.length === 0;
    if (!valid) {
      Logger.warn('Validasi DOM gagal. Elemen tidak ditemukan: ' + missing.join(', '));
      if (!silent) {
        UI.showResult('warning', 'Struktur Halaman Berubah', 'Elemen kritis tidak ditemukan: ' + missing.join(', ') + '. Beberapa fitur mungkin tidak berfungsi.');
      }
    }
    return { valid, missing };
  }

  const PANE_REGISTRY = Object.freeze({
    'check-nrp': {
      onActivate() {
        CEK_NRP.prefillSearch(true);
        CEK_NRP.renderResults();
        refreshGlobalData('', '', '', 'check-nrp');
      }
    },
    karyawan: {
      onActivate() {
        CEK_NRP.prefillSearch(true);
        CEK_NRP.renderResults();
      }
    },
    spkl: {
      onActivate() {
        refreshGlobalData('', '', '', 'spkl');
      }
    },
    kehadiran: {
      onActivate() {
        refreshGlobalData('', '', '', 'kehadiran');
      }
    },
    distribusi: {
      onActivate() {
        refreshGlobalData('', '', '', 'distribusi');
      }
    },
    anomali: { onActivate() { } },
    config: { onActivate() { } }
  });

  function activatePane(pane) {
    if (!window.QA) return;

    if (pane === 'check-nrp') {
      window.QA.onActionClick('Cek NRP');
      return;
    }

    if (pane === 'karyawan' || pane === 'cek-kary') {
      window.QA.onActionClick('Cari Karyawan');
      return;
    }

    // Map old pane ids to action labels for the header
    const labels = {
      'karyawan': 'Cari Karyawan',
      'spkl': 'SPKL Tools',
      'kehadiran': 'Automasi Kehadiran',
      'distribusi': 'Distribusi Jam Kerja',
      'config': 'Settings'
    };

    const targetLabel = labels[pane] || 'Action';
    qaState.activePage = targetLabel;

    const pool = document.getElementById('qa-views-pool');
    const detailState = els.detailState;
    if (!pool || !detailState) return;

    // First, move all existing children of detailState back to pool
    const activePanes = detailState.querySelectorAll('.qm-pane');
    activePanes.forEach(p => {
      p.style.display = 'none';
      p.classList.remove('active');
      pool.appendChild(p);
    });

    // Find the requested pane
    const targetId = 'qm-pane-' + pane;
    const targetEl = pool.querySelector('#' + targetId);
    if (targetEl) {
      targetEl.style.display = 'block';
      targetEl.classList.add('active');
      detailState.appendChild(targetEl);
    }

    // QA UI transitions
    els.searchBar.classList.add('is-hidden');
    els.detailHeader.classList.remove('is-hidden');
    els.detailTitle.textContent = targetLabel;
    els.emptyState.classList.add('is-hidden');
    els.resultsState.classList.add('is-hidden');
    els.detailState.classList.remove('is-hidden');
    els.detailState.classList.add('is-entering');
    setTimeout(() => {
      els.detailState.classList.remove('is-entering');
    }, 260);

    state.activeTab = pane;
  }

  const UI_EVENT_BINDINGS = Object.freeze([
    { event: 'click', selector: '#qm-fab', handler: togglePanel },
    { event: 'click', selector: '#qm-btn-close-header', handler: closePanel },
    { event: 'click', selector: '#qm-btn-check', handler() { CEK_NRP.runLookup(panelReaders.lookup()); } },
    { event: 'click', selector: '#qm-btn-karyawan-search', handler() { CEK_NRP.searchByQuery(panelReaders.karyawanSearch()); } },
    { event: 'click', selector: '#qm-btn-spkl-batch', handler: SPKL.startBatchProcess },
    { event: 'click', selector: '#qm-btn-spkl-many-nrp', handler: SPKL.startManyNrpBatch },
    { event: 'click', selector: '#qm-btn-spkl-page-cek', handler() { SPKL.checkByDateRange(panelReaders.spklCheck()); } },
    { event: 'click', selector: '#qm-btn-spkl-online-cek', handler() { SPKL.openOnlineForDate(panelReaders.spklOnline()); } },
    { event: 'click', selector: '#qm-btn-hadir-check', handler() { KEHADIRAN.checkByRange(panelReaders.attendanceCheck()); } },
    { event: 'click', selector: '#qm-btn-hadir-proses', handler() { KEHADIRAN.submitSingle(panelReaders.attendanceInput()); } },
    { event: 'click', selector: '#qm-btn-hadir-bulan-proses', handler: KEHADIRAN.startBulanBatch },
    { event: 'click', selector: '#qm-btn-hadir-many-proses', handler: KEHADIRAN.startManyNrpBatch },
    { event: 'click', selector: '#qm-btn-distribusi-proses', handler() { DISTRIBUSI.submitJkChange(panelReaders.distribusiJk()); } },
    { event: 'click', selector: '#qm-btn-distribusi-subsi-proses', handler() { DISTRIBUSI.submitSubsi(panelReaders.distribusiSubsi()); } },
    { event: 'click', selector: '#qm-btn-show-logs', handler: handleShowLogs },
    { event: 'click', selector: '#qm-btn-clear-logs', handler: handleClearLogs },
    { event: 'click', selector: '#qm-btn-export-logs', handler: handleExportLogs },
    { event: 'change', selector: '#qm-config-debug-mode', handler: handleDebugToggle },
    {
      event: 'click', selector: '#qm-global-cancel-btn', handler() {
        Logger.info('User cancelled automation');
        state.cancelRequested = true;
        sessionStorage.removeItem(STORAGE.SPKL_QUEUE);
        sessionStorage.removeItem(STORAGE.SPKL_CURRENT_INDEX);
        sessionStorage.removeItem(STORAGE.SPKL_FIX_PENDING);
        const activeFlow = getAutomationFlow();
        if (activeFlow) clearAutomationFlow(activeFlow.id);
        else sessionStorage.removeItem(STORAGE.AUTO_FINISHED);
        if (state.activeCancelableFlow) {
          UI.setGlobalProgress(5, 'Membatalkan proses...', true);
          if (state.activeAbortController) state.activeAbortController.abort();
          return;
        }
        UI.hideGlobalLoader(0);
        UI.showResult('info', 'Dibatalkan', 'Proses otomasi dihentikan oleh pengguna.');
      }
    },
    {
      event: 'click', selector: '.qm-progress-container', handler() {
        activatePane('config');
        const container = uiAdapter.get('logContainer');
        const btn = uiAdapter.get('showLogsButton');
        if (container && btn) {
          container.classList.remove('qm-hidden');
          btn.classList.add('qm-active');
          renderLogs();
          btn.querySelector('span').textContent = 'Sembunyikan Log';
        }
      }
    },
    {
      event: 'keydown', selector: '#qm-spkl-online-nrp', handler(e) {
        if (e.key === 'Enter') {
          e.preventDefault();
          SPKL.openOnlineForDate(panelReaders.spklOnline());
        }
      }
    },
    {
      event: 'keydown', selector: '#qm-spkl-page-nrp', handler(e) {
        if (e.key === 'Enter') {
          e.preventDefault();
          SPKL.checkByDateRange(panelReaders.spklCheck());
        }
      }
    },
    {
      event: 'change', selector: '#qm-fix-many-ot', handler() {
        const box = uiAdapter.get('#qm-fix-many-ot7-box');
        if (!box) return;
        box.classList.toggle('qm-hidden', this.value !== '7');
      }
    },
    {
      event: 'input', selector: '#qm-fix-spkl-data', handler() {
        const box = uiAdapter.get('#qm-fix-spkl-ot7-box');
        if (!box) return;
        const has7 = this.value.split(/[,\n]+/).some(item => {
          const parts = item.trim().split(/[-:=]/);
          return parts.length > 1 && parts[1].trim() === '7';
        });
        box.classList.toggle('qm-hidden', !has7);
      }
    },
    {
      event: 'keydown', selector: '#qm-input-nrp, #qm-input-bulan, #qm-input-tahun', handler(e) {
        if (e.key !== 'Enter') return;
        e.preventDefault();
        CEK_NRP.runLookup(panelReaders.lookup());
      }
    },
    {
      event: 'input', selector: '#qm-spkl-page-nrp, #qm-spkl-online-nrp, #qm-fix-spkl-nrp, #qm-input-hadir-check-nrp, #qm-input-hadir-nrp, #qm-input-hadir-bulan-nrp', handler() {
        const nrp = this.value.trim();
        if (nrp.length >= 4) refreshGlobalData(nrp);
      }
    },
    {
      event: 'input', selector: '#qm-input-hadir-check-nrp', handler() {
        state.attendanceCheck = createEmptyAttendanceCheck();
        KEHADIRAN.renderCheckResult();
      }
    },
    {
      event: 'input', selector: '#qm-spkl-page-nrp', handler() {
        state.spklCheck = createEmptySpklCheck();
        SPKL.renderCheckResult();
      }
    },
    {
      event: 'change', selector: '#qm-spkl-page-start-date, #qm-spkl-page-end-date', handler() {
        state.spklCheck = createEmptySpklCheck();
        SPKL.renderCheckResult();
      }
    },
    {
      event: 'click', selector: '#qm-btn-spkl-batch-edit', handler() {
        SPKL.startPageLoop();
      }
    },
    {
      event: 'click', selector: '#qm-btn-hadir-batch-edit', handler() {
        KEHADIRAN.startPageLoop();
      }
    },
    {
      event: 'change', selector: '.qm-hadir-batch-cb', handler() {
        KEHADIRAN.toggleRowInputs(this);
      }
    },
    {
      event: 'click', selector: '.qm-modal-close-btn, .qm-modal-cancel-btn', handler() {
        SPKL.closeBatchEdit();
        KEHADIRAN.closeBatchEdit();
      }
    },
    { event: 'click', selector: '#qm-btn-spkl-edit-save', handler() { SPKL.processBatchEdit(panelReaders.spklInlineEdit()); } },
    { event: 'click', selector: '#qm-btn-hadir-edit-save', handler() { KEHADIRAN.processBatchEdit(panelReaders.hadirInlineEdit()); } },
    {
      event: 'change', selector: '#qm-hadir-batch-cb-all', handler() {
        const isChecked = this.checked;
        document.querySelectorAll('.qm-hadir-batch-cb').forEach(cb => {
          cb.checked = isChecked;
          KEHADIRAN.toggleRowInputs(cb);
        });
      }
    },
    {
      event: 'change', selector: '.qm-spkl-batch-cb', handler() {
        SPKL.toggleRowInputs(this);
      }
    },
    {
      event: 'change', selector: '#qm-spkl-batch-cb-all', handler() {
        const isChecked = this.checked;
        document.querySelectorAll('.qm-spkl-batch-cb').forEach(cb => {
          cb.checked = isChecked;
          SPKL.toggleRowInputs(cb);
        });
      }
    },
    {
      event: 'change', selector: '#qm-input-hadir-check-start-date, #qm-input-hadir-check-end-date', handler() {
        state.attendanceCheck = createEmptyAttendanceCheck();
        KEHADIRAN.renderCheckResult();
      }
    },
    {
      event: 'input', selector: '#qm-input-karyawan-search', handler() {
        const { query, bulan, tahun } = panelReaders.karyawanSearch();
        if (isValidNrp(query)) syncGlobalInputs(query, bulan, tahun);
      }
    },
    {
      event: 'keydown', selector: '#qm-input-karyawan-search', handler(e) {
        if (e.key === 'Enter') {
          e.preventDefault();
          CEK_NRP.searchByQuery(panelReaders.karyawanSearch());
        }
      }
    },
    {
      event: 'change', selector: '#qm-input-bulan, #qm-input-tahun, #qm-fix-spkl-bulan, #qm-fix-spkl-tahun, #qm-input-hadir-bulan-bln', handler() {
        const isMonthField = this.id === 'qm-input-bulan' || this.id === 'qm-fix-spkl-bulan' || this.id === 'qm-input-hadir-bulan-bln';
        const isYearField = this.id === 'qm-input-tahun' || this.id === 'qm-fix-spkl-tahun';
        refreshGlobalData('', isMonthField ? this.value : '', isYearField ? this.value : '', this.id);
      }
    },
    { event: 'click', selector: '.qm-tab', handler: handleTabClick },
    { event: 'click', selector: '.qm-karyawan-detail-btn', handler() { CEK_NRP.toggleDetail(this.dataset.key || ''); } },
    { event: 'click', selector: '.qm-karyawan-edit-btn', handler() { CEK_NRP.toggleEditor(this.dataset.key || ''); } },
    { event: 'click', selector: '.qm-karyawan-save-btn', handler() { CEK_NRP.saveEditor(panelReaders.karyawanSave(this.dataset.key || '')); } },
    {
      event: 'click', selector: '.qm-group-header-card', handler() {
        const group = this.dataset.group;
        if (!state.karyawanGroupsExpanded) {
          state.karyawanGroupsExpanded = { internal: false, outsource: false };
        }
        state.karyawanGroupsExpanded[group] = !state.karyawanGroupsExpanded[group];
        renderKaryawanResults();
      }
    },
    {
      event: 'click', selector: '#qm-btn-karyawan-reset', handler() {
        state.karyawanResults = [];
        state.karyawanQuery = '';
        state.karyawanActivePanel = null;
        state.karyawanError = null;
        const input = document.getElementById('qm-input-karyawan-search');
        if (input) input.value = '';
        renderKaryawanResults();
      }
    },
    {
      event: 'click', selector: '.qm-directory-item', handler(e) {
        if (e.target.closest('button') || e.target.closest('.qm-preview-tag')) return;
        const key = this.dataset.key || '';
        if (!key) return;
        const result = state.karyawanResults.find(r => r.key === key);
        if (result && result.nrp) {
          clearEmployeeCache(result.nrp);
          const activeMode = state.karyawanActivePanel?.mode || 'detail';
          state.karyawanActivePanel = { key, mode: activeMode };
          if (activeMode === 'edit') {
            loadKaryawanEditor(key, result.nrp, result.id, result.source === 'outsource');
          } else {
            loadKaryawanDetail(key, result.nrp, result.profileUrl);
          }
          setTimeout(() => {
            const leftPanel = document.getElementById('qm-karyawan-left-panel');
            if (leftPanel) leftPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
          }, 50);
        }
      }
    },
    { event: 'click', selector: '.qm-fix-dot', handler: handleFixDotClick },
    { event: 'click', selector: '.qm-btn-fix-pill', handler: handleFixDotClick },
    {
      event: 'click', selector: '#qm-btn-batch-check', handler() {
        if (this.dataset.running) ANOMALI.cancelBatchCheck();
        else ANOMALI.startBatchCheck();
      }
    },
    {
      event: 'mouseover', selector: '#qm-btn-batch-check', handler() {
        if (!this.dataset.running) return;
        this.classList.add('qm-btn-danger');
        this.textContent = 'Batal';
      }
    },
    {
      event: 'mouseout', selector: '#qm-btn-batch-check', handler() {
        if (!this.dataset.running) return;
        this.classList.remove('qm-btn-danger');
        this.textContent = 'Memproses...';
      }
    },
    { event: 'click', selector: '#qm-btn-batch-clear', handler: ANOMALI.clearBatchResults },
    { event: 'click', selector: '#qm-btn-export-batch', handler: ANOMALI.exportBatchResults },
    { event: 'click', selector: '.qm-batch-nrp-link', handler: ANOMALI.handleBatchNrpClick },
    { event: 'click', selector: '.qm-batch-fix-btn', handler: ANOMALI.handleBatchFixClick },
    {
      event: 'click', selector: '.qm-btn-barcode-delete', handler: async function () {
        const url = this.dataset.url;
        if (!url || !uiAdapter.confirm('Hapus data barcode ini?')) return;

        const btn = this;
        const originalText = btn.textContent;
        btn.disabled = true;
        btn.textContent = '...';

        try {
          await hrisFetch(url);
          KEHADIRAN.checkByRange(panelReaders.attendanceCheck());
        } catch (e) {
          uiAdapter.alert('Gagal menghapus data: ' + e.message);
          btn.disabled = false;
          btn.textContent = originalText;
        }
      }
    },
    {
      event: 'click', selector: '.qm-batch-group-header', handler() {
        const table = this.closest('table');
        if (!table) return;
        const rows = table.querySelectorAll(this.dataset.target);
        const isExpanded = this.classList.contains('expanded');
        this.classList.toggle('expanded');
        rows.forEach(r => {
          if (isExpanded) {
            r.classList.add('qm-hidden');
            r.classList.remove('qm-table-row', 'expanded');
            if (r.classList.contains('qm-batch-seksi-header')) {
              const childRows = table.querySelectorAll(r.dataset.target);
              childRows.forEach(cr => {
                cr.classList.add('qm-hidden');
                cr.classList.remove('qm-table-row');
              });
            }
          } else {
            if (r.classList.contains('qm-table-header') || r.classList.contains('qm-batch-seksi-header')) {
              r.classList.remove('qm-hidden');
              r.classList.add('qm-table-row');
            }
          }
        });
      }
    },
    {
      event: 'click', selector: '.qm-batch-seksi-header', handler() {
        const table = this.closest('table');
        if (!table) return;
        const rows = table.querySelectorAll(this.dataset.target);
        const isExpanded = this.classList.contains('expanded');
        this.classList.toggle('expanded');
        rows.forEach(r => {
          if (isExpanded) {
            r.classList.add('qm-hidden');
            r.classList.remove('qm-table-row');
          } else {
            r.classList.remove('qm-hidden');
            r.classList.add('qm-table-row');
          }
        });
      }
    },
    {
      event: 'click', selector: '.qm-batch-date-header', handler(e) {
        if (e.target.closest('.qm-fix-dot') || e.target.closest('.qm-batch-fix-btn')) return;
        const content = this.nextElementSibling;
        const isExpanded = this.classList.contains('expanded');
        this.classList.toggle('expanded');
        if (content) {
          if (isExpanded) {
            content.classList.add('qm-hidden');
            content.classList.remove('qm-visible-block');
          } else {
            content.classList.remove('qm-hidden');
            content.classList.add('qm-visible-block');
          }
        }
      }
    },
    {
      event: 'click', selector: '.qm-accordion-header', handler() {
        this.classList.toggle('expanded');
        const content = this.nextElementSibling;
        if (content) content.classList.toggle('qm-content-open');
      }
    },
    {
      event: 'change', selector: '#qm-config-collapse-menu', handler() {
        alwaysCollapseMenu = this.checked;
        GM_setValue('qm_always_collapse', alwaysCollapseMenu);
        if (alwaysCollapseMenu) enforceSidebar();
        else document.body.classList.remove('enlarged');
      }
    },
    { event: 'click', selector: '#qm-btn-theme-light', handler() { UI.applyTheme('light'); } },
    { event: 'click', selector: '#qm-btn-theme-dark', handler() { UI.applyTheme('dark'); } },
    { event: 'click', selector: '#qm-btn-record-shortcut', handler: handleRecordShortcut },
    { event: 'click', selector: '.button-menu-mobile, .open-left, #sidebar-menu', handler() { manualSidebarOverride = true; } }
  ]);

  function bindDeclarativeEvents(bindings) {
    bindings.forEach(binding => on(binding.event, binding.selector, binding.handler));
  }

  function bootstrapDomainState() {
    if (!state.karyawanEditor) resetKaryawanEditor();
    if (!state.karyawanDetail) resetKaryawanDetail();
  }

  function runStartupAutomations() {
    SPKL.highlightPendingDate();
    SPKL.autoFillTargetPage();
    KEHADIRAN.autoSearchPage();
    KEHADIRAN.autoClickAddData();
    KEHADIRAN.autoInput();
    DISTRIBUSI.autoDistribusi();
    DISTRIBUSI.autoDistribusiSubsi();
    DISTRIBUSI.autoDistKK();
    SPKL.resumeBatch();
    SPKL.resumePageLoop();
    KEHADIRAN.resumeManyNrpBatch();
    KEHADIRAN.resumeBulanBatch();
    KEHADIRAN.resumePageLoop();
    DISTRIBUSI.initChangeEvents();
    DISTRIBUSI.checkJkRestoration();
    SPKL.autoFillEditPage();
  }

  function handleAutomationReturnState() {
    const activeFlow = getAutomationFlow();
    const returnUrl = activeFlow?.returnUrl || sessionStorage.getItem(STORAGE.RETURN_URL);
    const isFinished = activeFlow ? !!activeFlow.finished : sessionStorage.getItem(STORAGE.AUTO_FINISHED) === 'true';
    const hasRestorationPending = sessionStorage.getItem('qm_jk_to_restore_' + getPageContext().nrp);

    if (!isFinished || !returnUrl) return;

    const currentUrl = window.location.href.split('?')[0];
    const targetUrlBase = returnUrl.split('?')[0];
    const isReturnPage = currentUrl === targetUrlBase || window.location.href.includes(returnUrl);

    if (isReturnPage) {
      if (!hasRestorationPending) {
        if (activeFlow) clearAutomationFlow(activeFlow.id);
        else {
          sessionStorage.removeItem(STORAGE.AUTO_FINISHED);
          sessionStorage.removeItem(STORAGE.RETURN_URL);
        }
        UI.showResult('success', 'Selesai', 'Tugas latar belakang telah diselesaikan.');

        if (isAttendancePagePath()) {
          setTimeout(() => {
            document.querySelector('[data-pane="anomali"]')?.click();
          }, 500);
        }
      }
      return;
    }

    if (!window.location.search.includes('qm_auto')) {
      UI.showGlobalLoader('Selesai', 'Kembali ke halaman awal...');
      setTimeout(() => uiAdapter.openUrl(returnUrl, '_self'), 1500);
    }
  }

  function mountPanel() {
    if (uiAdapter.get('fab')) return false;
    document.body.insertAdjacentHTML('beforeend', HTML);
    return true;
  }

  function initializePanelDefaults() {
    const ctx = getPageContext();
    const now = new Date();
    const currentMonth = now.getMonth() + 1;
    const currentYear = now.getFullYear();
    const defaultCheckDate = `${currentYear}-${String(currentMonth).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
    const monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    const monthOptions = monthNames.map((name, i) => `<option value="${i + 1}" ${i + 1 === currentMonth ? 'selected' : ''}>${name}</option>`).join('');

    ['fixSpklNrp', 'hadirBulanNrp', 'attendanceCheckNrp', 'distribusiNrp', 'distribusiKkNrp'].forEach(key => {
      if (ctx.nrp) uiAdapter.value(key, ctx.nrp);
    });

    ['globalMonth', 'fixSpklMonth', 'spklPageMonth', 'hadirBulanMonth'].forEach(key => {
      const el = uiAdapter.get(key);
      if (el) renderSafe(el, monthOptions);
    });

    let years = '';
    for (let y = currentYear - 2; y <= currentYear + 1; y++) {
      years += `<option value="${y}" ${y === currentYear ? 'selected' : ''}>${y}</option>`;
    }
    ['globalYear', 'fixSpklYear'].forEach(key => {
      const el = uiAdapter.get(key);
      if (el) renderSafe(el, years);
    });

    uiAdapter.value('fixManyDate', defaultCheckDate);
    uiAdapter.value('attendanceCheckStartDate', defaultCheckDate);
    uiAdapter.value('attendanceCheckEndDate', defaultCheckDate);
    uiAdapter.value('spklPageStartDate', defaultCheckDate);
    uiAdapter.value('spklPageEndDate', defaultCheckDate);
    uiAdapter.value('distribusiKkDate', `${currentYear}-${String(currentMonth).padStart(2, '0')}`);
    uiAdapter.value('distribusiDate', defaultCheckDate);
    uiAdapter.value('shortcutInput', shortcutKey);

    CEK_NRP.renderResults();
    KEHADIRAN.renderCheckResult();
    SPKL.renderCheckResult();
    UI.applyTheme(state.theme);

    const collapseCheckbox = uiAdapter.get('configCollapseMenu');
    if (collapseCheckbox) collapseCheckbox.checked = alwaysCollapseMenu;

    initDraggable();

    // Prefill batch results if they exist in localStorage
    try {
      const savedResults = localStorage.getItem('qm-batch-results');
      if (savedResults) {
        state.batchResults = JSON.parse(savedResults);
        state.batchBulan = parseInt(localStorage.getItem('qm-batch-bulan')) || 0;
        state.batchTahun = parseInt(localStorage.getItem('qm-batch-tahun')) || 0;

        if (state.batchBulan) {
          const bulEl = uiAdapter.get('globalMonth');
          if (bulEl) bulEl.value = state.batchBulan;
        }
        if (state.batchTahun) {
          const tahEl = uiAdapter.get('globalYear');
          if (tahEl) tahEl.value = state.batchTahun;
        }

        renderBatchResults();
        const container = uiAdapter.get('batchResults');
        if (container) container.classList.add('is-visible');
      }
    } catch (e) {
      console.error('Failed to restore batch results:', e);
    }

    const lastTab = localStorage.getItem('qm_last_tab') || 'check-nrp';
    activatePane(lastTab);
  }

  function applyAccessibilityBindings() {
    uiAdapter.all('accordionHeaders').forEach(el => {
      if (!el.hasAttribute('tabindex')) el.setAttribute('tabindex', '0');
    });
  }

  function startDomValidationPolling() {
    if (!(isAttendancePagePath() || isBarcodeCreatePagePath() || isDistribusiKalenderPagePath() || isSpklPagePath())) return;
    let domCheckAttempts = 0;
    const maxDomCheckAttempts = 10;
    const domCheckInterval = setInterval(function () {
      domCheckAttempts++;
      const isFinalAttempt = domCheckAttempts >= maxDomCheckAttempts;
      const result = validateDomStructure({ silent: !isFinalAttempt });
      if (result.valid || isFinalAttempt) clearInterval(domCheckInterval);
    }, 200);
  }

  function bindUiEvents() {
    document.addEventListener('keydown', handleKeydownDocument);
    document.addEventListener('click', handleDocumentClick);
    bindDeclarativeEvents(UI_EVENT_BINDINGS);
    initKeyboardNavigation();
  }

  function init() {
    bootstrapDomainState();
    const schemaValid = validateStorageSchema();
    runStartupAutomations();
    handleAutomationReturnState();

    if (mountPanel()) {
      initQaEngine();

      // Register QA Actions mapping legacy panes
      if (window.QA) {
        window.QA.registerActions([
          {
            label: 'Cek NRP',
            iconMarkup: ICONS.userCheck,
            description: 'Check employee attendance anomalies by NRP, month, and year.',
            keywords: ['nrp', 'cek nrp', 'absensi', 'anomalies', 'kehadiran', 'batch', 'barcode', 'employee check'],
            onClick: () => {
              window.QA.renderDetail('Cek NRP', VIEW_CEK_NRP);
              localStorage.setItem('qm_last_tab', 'check-nrp');
              if (typeof CEK_NRP !== 'undefined') {
                CEK_NRP.prefillSearch(true);
                CEK_NRP.renderResults();
              }
              if (typeof refreshGlobalData === 'function') {
                refreshGlobalData('', '', '', 'check-nrp');
              }
              if (state.batchBulan) {
                const bulEl = uiAdapter.get('globalMonth');
                if (bulEl) bulEl.value = state.batchBulan;
              }
              if (state.batchTahun) {
                const tahEl = uiAdapter.get('globalYear');
                if (tahEl) tahEl.value = state.batchTahun;
              }
              if (typeof renderBatchResults === 'function') {
                renderBatchResults();
              }
            }
          },
          {
            label: 'Cari Karyawan',
            iconMarkup: ICONS.user,
            description: 'Find employee records by NRP or employee name.',
            keywords: ['karyawan', 'employee', 'pegawai', 'nama', 'nrp', 'profile', 'data karyawan'],
            onClick: () => {
              window.QA.renderDetail('Cari Karyawan', VIEW_CEK_KARY);
              localStorage.setItem('qm_last_tab', 'cek-kary');
              if (typeof CEK_NRP !== 'undefined') {
                CEK_NRP.prefillSearch(true);
                CEK_NRP.renderResults();
              }
              if (typeof refreshGlobalData === 'function') {
                refreshGlobalData('', '', '', 'cek-kary');
              }
            }
          },
          {
            label: 'SPKL Tools',
            iconMarkup: ICONS.file,
            description: 'Review and fix SPKL overtime entries and approval status.',
            keywords: ['spkl', 'overtime', 'ot', 'lembur', 'approval', 'fix spkl', 'status spkl'],
            onClick: () => activatePane('spkl')
          },
          {
            label: 'Automasi Kehadiran',
            iconMarkup: ICONS.clock,
            description: 'Automate attendance corrections, barcode entries, and schedule checks.',
            keywords: ['kehadiran', 'attendance', 'absen', 'barcode', 'masuk', 'pulang', 'shift', 'auto hadir'],
            onClick: () => activatePane('kehadiran')
          },
          {
            label: 'Distribusi Jam Kerja',
            iconMarkup: ICONS.map,
            description: 'Run work-hour distribution and calendar assignment automation.',
            keywords: ['distribusi', 'jam kerja', 'kalender kerja', 'kk', 'schedule', 'calendar', 'assign shift'],
            onClick: () => activatePane('distribusi')
          },
          {
            label: 'Settings',
            iconMarkup: ICONS.settings,
            description: 'Configure quick menu behavior, preferences, and saved options.',
            keywords: ['settings', 'config', 'configuration', 'pengaturan', 'preferensi', 'options'],
            onClick: () => activatePane('config')
          }
        ]);
      }

      initializePanelDefaults();
    }


    ANOMALI.detect();

    if (!schemaValid) {
      UI.showResult('warning', 'Data Direset', 'Versi data sesi tidak cocok. Data lama telah dibersihkan.');
    }
    startDomValidationPolling();
    bindUiEvents();
    enforceSidebar();
    const _sidebarObserver = new MutationObserver(enforceSidebar);
    _sidebarObserver.observe(document.body, { attributes: true, attributeFilter: ['class'] });
    applyAccessibilityBindings();
  }


  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
