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
   * 0. CONFIGURATION
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
    FADE_SHORT: 10,
    FADE_MEDIUM: 250,
  });

  /** sessionStorage key constants. */
  const STORAGE = Object.freeze({
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
    HIGHLIGHT_SPKL: 'qm_highlight_spkl_date',
    SPKL_SAVED: 'spkl_saved_data',
    SPKL_BATCH: 'hris_spkl_ot_runner_v1',
    INPUT_HADIR: 'qm_auto_input_hadir_data',
    HADIR_BATCH: 'qm_auto_hadir_batch_v1',
    HISTORY: 'qm_history',
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

  function getCurrentQueryParams() {
    return new URLSearchParams(window.location.search);
  }

  function getCurrentPath() {
    return window.location.pathname || '';
  }

  function isOutsourceNrp(nrp) {
    return String(nrp || '').length === 8;
  }

  function isAttendancePagePath(path = getCurrentPath()) {
    return (path.includes('/tabelkehadiran') || path.includes('/tabelkehadirankkwt')) && !path.includes('/rekap');
  }

  function isBarcodePagePath(path = getCurrentPath()) {
    return path.includes('/absenbarcode');
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

  function getSpklAddPageKind(path = getCurrentPath()) {
    if (path === '/spkl/add') return 'internal';
    if (path === '/spkloutsource/add') return 'outsource';
    return null;
  }

  function getAbsenCreatePageKind(path = getCurrentPath()) {
    if (path === '/absenbarcode/create') return 'internal';
    if (path === '/absenbarcodeos/create') return 'outsource';
    return null;
  }

  function getEmployeeRouteSet(nrp) {
    return isOutsourceNrp(nrp)
      ? {
        search: ROUTES.KARYAWANOS_SEARCH(nrp),
        general: ROUTES.KARYAWANOS_GENERAL,
        profile: ROUTES.KARYAWANOS_PROFILE,
        edit: ROUTES.KARYAWANOS_EDIT
      }
      : {
        search: ROUTES.KARYAWAN_SEARCH(nrp),
        general: ROUTES.KARYAWAN_GENERAL,
        profile: ROUTES.KARYAWAN_PROFILE,
        edit: ROUTES.KARYAWAN_EDIT
      };
  }

  /** Shorthand: select internal or OS route based on NRP type. */
  function getRoute(nrp, internalRoute, osRoute) {
    return isOutsourceNrp(nrp) ? osRoute : internalRoute;
  }

  function getAttendanceUrl(bulan, tahun, nrp) {
    return getRoute(nrp, ROUTES.TABEL_HADIR, ROUTES.TABEL_HADIR_OS)(bulan, tahun, nrp);
  }

  function getDistribusiBaseUrl(nrp) {
    return getRoute(nrp, ROUTES.DISTRIBUSI, ROUTES.DISTRIBUSI_OS)(nrp);
  }

  function getEmployeeEditUrl(nrp, id) {
    return getEmployeeRouteSet(nrp).edit(id);
  }

  function getSpklBaseUrl(nrp) {
    return getRoute(nrp, ROUTES.SPKL_BASE, ROUTES.SPKL_OS_BASE);
  }

  function getSpklCreateUrl(nrp) {
    return getRoute(nrp, ROUTES.SPKL_CREATE, ROUTES.SPKL_OS_CREATE);
  }

  function getSpklAddUrlByKind(kind) {
    return kind === 'outsource' ? ROUTES.SPKL_OS_ADD : ROUTES.SPKL_ADD;
  }

  function getAbsenCreateUrl(nrp) {
    return getRoute(nrp, ROUTES.ABSEN_BARCODE_CREATE, ROUTES.ABSEN_BARCODE_OS_CREATE);
  }

  function getAbsenCreateUrlByKind(kind) {
    return kind === 'outsource' ? ROUTES.ABSEN_BARCODE_OS_CREATE : ROUTES.ABSEN_BARCODE_CREATE;
  }

  function getAbsenAddUrl(nrp) {
    return getRoute(nrp, ROUTES.ABSEN_BARCODE_ADD, ROUTES.ABSEN_BARCODE_OS_ADD);
  }

  function buildSpklOnlineUrl(ctx, minDate, maxDate) {
    const bulan = String(ctx.bulan).padStart(2, '0');
    const min = String(minDate).padStart(2, '0');
    const max = String(maxDate).padStart(2, '0');
    return ROUTES.SPKL_ONLINE(ctx.tahun, bulan, min, max, ctx.nrp);
  }

  /* ============================================================
   * 1. SHARED HELPERS
   * ============================================================ */

  /** Unified DOM-based sanitization and parsing. */
  function parseHTML(html, fullDoc = false) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const dangerous = doc.querySelectorAll('script,style,link,iframe,object,embed,svg,math');
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

  /** Simple HTML escape helper. */
  function escHtml(str) {
    if (!str) return '';
    return String(str).replace(/[&<>"']/g, m => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[m]));
  }

  /** Safely set innerHTML using DOM-based sanitization via parseHTML. */
  function setInnerHTML(el, html) {
    if (!el) return;
    if (!html && html !== '') { el.innerHTML = html; return; }
    const doc = parseHTML(`<div>${html}</div>`);
    const container = doc.querySelector('div');
    el.innerHTML = '';
    while (container.firstChild) {
      el.appendChild(container.firstChild);
    }
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

  /** Build auto-distribusi link for a given date + shift. */
  function getDistribusiLink(ctx, tglAwal, shiftVal, tglAkhir = null) {
    if (!ctx.nrp) return '';
    const base = getDistribusiBaseUrl(ctx.nrp);

    // If tglAwal is already a full date (YYYY-MM-DD), use it directly
    const dAwal = (tglAwal && tglAwal.includes('-')) ? tglAwal : `${ctx.tahun}-${String(ctx.bulan).padStart(2, '0')}-${String(tglAwal).padStart(2, '0')}`;
    const dAkhir = tglAkhir ? (tglAkhir.includes('-') ? tglAkhir : `${ctx.tahun}-${String(ctx.bulan).padStart(2, '0')}-${String(tglAkhir).padStart(2, '0')}`) : dAwal;

    return `${base}?qm_auto_distribusi=1&nrp=${encodeURIComponent(ctx.nrp)}&tanggal_awal=${encodeURIComponent(dAwal)}&tanggal_akhir=${encodeURIComponent(dAkhir)}&bagian=${encodeURIComponent(ctx.bagian)}&seksi=${encodeURIComponent(ctx.seksi)}&shift=${encodeURIComponent(shiftVal)}`;
  }

  /** Build Kehadiran (Barcode) link for a given context. */
  function getKehadiranLink(ctx) {
    if (!ctx.nrp) return '';
    return (isOutsourceNrp(ctx.nrp) ? ROUTES.ABSEN_BARCODE_OS : ROUTES.ABSEN_BARCODE)(ctx.tahun, String(ctx.bulan).padStart(2, '0'), ctx.nrp);
  }

  /** Parse "HH:MM" or "HH.MM" into decimal hours. */
  function parseTime(t) {
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

  // ── Utility Functions (Phase 0 deduplication) ───────────────────────────

  /** Set a form field value and dispatch the given events.
   *  Replaces the repeated pattern: el.value = val; el.dispatchEvent(new Event(...)) */
  function setFieldValue(el, value, events = ['change']) {
    if (!el) return;
    el.value = value;
    events.forEach(e => el.dispatchEvent(new Event(e, { bubbles: true })));
    // Auto-refresh SelectPicker if applicable
    if (el.tagName === 'SELECT' && el.classList.contains('selectpicker')) {
      selectPickerRefresh(el);
    }
  }

  /** Find an option in a select element matching a predicate, select it, and dispatch change.
   *  Returns true if a match was found. */
  function selectOption(select, matchFn) {
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
  function selectPickerRefresh(select) {
    if (typeof window.$ !== 'undefined' && window.$(select).selectpicker) {
      window.$(select).selectpicker('refresh');
    }
  }

  /** Custom Event Delegation helper. */
  function delegate(eventName, selector, handler) {
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

  /** Parse CSS class/style/bgcolor attributes to detect Libur/HalfDay status.
   *  Used by both countLibur() and getRowStatus() to avoid duplicated regex. */
  function parseRowCssFlags(tr) {
    const trStr = ((tr.getAttribute('class') || '') + (tr.getAttribute('style') || '') + (tr.getAttribute('bgcolor') || '')).toLowerCase();
    const tds = tr.querySelectorAll('td');
    const td0Str = tds.length > 0 ? ((tds[0].getAttribute('class') || '') + (tds[0].getAttribute('style') || '') + (tds[0].getAttribute('bgcolor') || '')).toLowerCase() : '';
    const combined = trStr + td0Str;
    return {
      isLiburColor: /danger|red|#ff0000/.test(combined),
      isHalfDayColor: /warning|yellow|blue|#ffff00|#0000ff/.test(combined)
    };
  }

  /** Unified Fix/Perbaikan click handler — shared by onFixDotClick and onBatchFixClick. */
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
  function decrementPendingChecks() {
    state.pendingChecks--;
    if (state.pendingChecks <= 0) {
      const tab = document.querySelector('[data-pane="anomali"]');
      if (tab) tab.classList.remove('qm-tab-loading');
    }
  }

  /** Common success indicators in HRIS response text. */
  const SUCCESS_KEYWORDS = ['Berhasil', 'Selesai', 'sukses', 'successfully', 'Distribution Process Completed', 'alert-success'];

  function isSuccessResponse(text) {
    return SUCCESS_KEYWORDS.some(kw => text.includes(kw));
  }

  function buildFormPayload(doc, overrides = {}) {
    const form = doc.querySelector('form');
    if (!form) throw new Error('Form tidak ditemukan.');
    const params = new URLSearchParams();
    const headers = { 'Content-Type': 'application/x-www-form-urlencoded', 'X-Requested-With': 'XMLHttpRequest' };
    const csrfMeta = doc.querySelector('meta[name="csrf-token"], meta[name="csrf-test-name"]');
    if (csrfMeta) headers['X-CSRF-TOKEN'] = csrfMeta.getAttribute('content');
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

  /* ============================================================
   * 2. SHARED EMPLOYEE FETCH HELPERS
   * ============================================================ */

  function buildEmployeeUrls(nrp) {
    const routeSet = getEmployeeRouteSet(nrp);
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

  function findDetailUrl(doc, nrp) {
    let detailUrl = '';
    const rows = doc.querySelectorAll('table tbody tr');
    rows.forEach(function (row) {
      if (row.textContent.includes(nrp)) {
        const links = Array.from(row.querySelectorAll('a'));
        // Prioritize links that look like detail buttons or contain 'Detail'
        const btn = links.find(a => {
          const txt = a.textContent.trim();
          const cls = a.className;
          return txt.includes('Detail') || cls.includes('btn-info') || cls.includes('btn-primary') || a.getAttribute('href')?.includes('/general/');
        });
        if (btn && btn.getAttribute('href')) detailUrl = btn.getAttribute('href');
      }
    });
    if (!detailUrl) {
      const allLinks = Array.from(doc.querySelectorAll('a'));
      const fallback = allLinks.find(a => (a.textContent.includes('Detail') || a.classList.contains('btn-info')) && a.getAttribute('href'));
      if (fallback) detailUrl = fallback.getAttribute('href');
    }
    return detailUrl;
  }

  /** Unified field value extractor (input or select). */
  function getFieldValue(doc, fieldName) {
    const el = doc.querySelector(`[name="${fieldName}"]`);
    if (!el) return '';
    if (el.tagName === 'SELECT') {
      const opt = el.querySelector('option[selected]') || el.options[el.selectedIndex];
      return (opt ? opt.value || opt.textContent : el.value).trim();
    }
    return el.value.trim();
  }

  function getNamaKaryawan(doc) {
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

  /** Wait for element using MutationObserver (replaces setInterval polling). */
  function waitForElement(selector, timeout = 5000) {
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

  // fetchEmployeePage — removed, replaced by getEmp()

  /** Fetch wrapper with timeout. Returns response text. */
  async function req(url, timeout = 15000) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), timeout);
    try {
      const res = await fetch(url, { signal: ctrl.signal });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return await res.text();
    } finally {
      clearTimeout(timer);
    }
  }

  /** Unified fetch with AbortController timeout. Returns Response. */
  async function fetchWithTimeout(url, options = {}, timeoutMs = 15000) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const response = await fetch(url, { ...options, signal: controller.signal });
      return response;
    } finally {
      clearTimeout(timeoutId);
    }
  }

  /** Read cached employee data from sessionStorage. Returns null if not fully cached. */
  function readEmployeeCache(nrp) {
    const jk = sessionStorage.getItem('qm_jk_' + nrp);
    const nama = sessionStorage.getItem('qm_nama_' + nrp);
    const id = sessionStorage.getItem('qm_id_' + nrp);
    const kk = sessionStorage.getItem('qm_KK_' + nrp);
    const bag = sessionStorage.getItem('qm_bag_' + nrp);
    if (!jk || !nama || !id || !kk || !bag) return null;
    return {
      jk, KK: kk, nama, id,
      editUrl: sessionStorage.getItem('qm_edit_url_' + nrp) || '',
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

  /** Unified employee data fetcher with sessionStorage cache. */
  async function getEmp(nrp) {
    const cached = readEmployeeCache(nrp);
    if (cached) return { found: true, ...cached };
    const urls = buildEmployeeUrls(nrp);
    let searchDoc, detailUrl;

    // Retry logic for NRP search (Bug Fix 1)
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        const html = await req(urls.searchUrl);
        searchDoc = parseHTML(html);
        detailUrl = findDetailUrl(searchDoc, nrp);
        if (detailUrl) break;
      } catch (e) {
        if (attempt === 3) throw e;
      }
      if (attempt < 3) await new Promise(r => setTimeout(r, 500 * attempt));
    }

    if (!detailUrl) return { found: false };
    const id = detailUrl.split('?')[0].split('/').filter(Boolean).pop();
    const [genHtml, profHtml] = await Promise.all([
      req(urls.buildGeneralUrl(id)),
      req(urls.buildProfileUrl(id))
    ]);
    const doc = parseHTML(genHtml);
    const profDoc = parseHTML(profHtml);

    // Find Edit General URL from the page buttons
    let editUrl = '';
    const editBtn = doc.querySelector('a[href*="editgeneral"]');
    if (editBtn) {
      editUrl = editBtn.getAttribute('href');
      if (!editUrl.startsWith('http')) editUrl = 'https://hris.kti.co.id' + (editUrl.startsWith('/') ? '' : '/') + editUrl;
    }

    const emp = {
      found: true,
      id: id,
      editUrl: editUrl,
      jk: getFieldValue(doc, 'kode_jam_kerja').split('-')[0].trim(),
      KK: getFieldValue(doc, 'kode_kalender_kerja'),
      nama: getNamaKaryawan(profDoc) || getNamaKaryawan(doc),
      bagian: getFieldValue(doc, 'kode_bagian') || '',
      seksi: getFieldValue(doc, 'kode_seksi') || '',
      group: getFieldValue(doc, 'kode_group') || ''
    };
    writeEmployeeCache(nrp, emp);
    return emp;
  }

  /** Fetch attendance table and return anomalies array. */
  async function fetchAttendance(nrp, bulan, tahun, bagian, seksi) {
    const prof = startProfile('fetchAttendance', { nrp, bulan, tahun });
    try {
      const html = await req(getAttendanceUrl(bulan, tahun, nrp));
      const doc = parseHTML(html);
      return scanAttendanceTable(doc, { tahun, bulan, nrp, bagian, seksi });
    } finally {
      finishProfile(prof, { nrp });
    }
  }

  /* ============================================================
   * 3. DEBUGGER
   * ============================================================ */

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

  function getPerfNow() {
    return (typeof performance !== 'undefined' && typeof performance.now === 'function')
      ? performance.now()
      : Date.now();
  }

  function getMedian(values) {
    if (!values.length) return 0;
    const sorted = [...values].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
  }

  function startProfile(label, meta = {}) {
    if (!state.debug) return null;
    return { label, meta, startedAt: getPerfNow() };
  }

  function finishProfile(token, meta = {}) {
    if (!token) return 0;

    const duration = getPerfNow() - token.startedAt;
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
    const median = recent.length >= PROFILE_CONFIG.MEDIAN_SAMPLE_SIZE ? getMedian(recent) : null;
    const payload = {
      ms: Number(duration.toFixed(2)),
      avgMs: Number((stats.total / stats.count).toFixed(2)),
      maxMs: Number(stats.max.toFixed(2)),
      medianMs: median === null ? null : Number(median.toFixed(2)),
      samples: stats.count,
      ...token.meta,
      ...meta
    };

    Logger.debug(`[PROFILE] ${token.label}`, payload);

    if (median !== null && median >= PROFILE_CONFIG.HOT_SYNC_MS) {
      const warnKey = `${token.label}:sync`;
      if (!state.profileFlags[warnKey]) {
        state.profileFlags[warnKey] = true;
        Logger.warn(`Profiling: ${token.label} median ${median.toFixed(2)}ms melewati ambang ${PROFILE_CONFIG.HOT_SYNC_MS}ms.`, payload);
      }
    }

    return duration;
  }

  function resetBatchProfile() {
    state.batchProfile = { renderTotal: 0, itemDurations: [] };
  }


  /* ============================================================
   * 4. STYLES
   * ============================================================ */
  GM_addStyle(`
    :root {
      /* --- Design Tokens --- */
      
      /* Colors: Primary (p), Success, Danger, Warning */
      /* Prefix p = Anthropic/Claude brand colors */
      --qm-p-500: #cc785c; /* Coral Brand */
      --qm-p-600: #a9583e; /* Coral Active */
      --qm-p-100: #faf9f5; /* Cream Canvas */
      
      --qm-success: #5db872;
      --qm-success-bg: #e8f5e9;
      --qm-danger: #c64545; /* Error Crimson */
      --qm-danger-bg: #fdf2f2;
      --qm-warning: #e8a55a; /* Amber Warning */
      --qm-warning-bg: #fff8e1;

      /* Neutrals: Exclusively Warm-toned */
      --qm-white: #ffffff;
      --qm-parchment: #faf9f5;   /* Canvas floor */
      --qm-ivory: #efe9de;       /* Surface Card */
      --qm-sand: #e6dfd8;        /* Hairline/Disabled */
      --qm-cream: #ebe6df;       /* Hairline Soft */
      --qm-near-black: #141413;  /* Ink */
      --qm-dark-surface: #181715;
      --qm-charcoal: #252523;    /* Body Strong */
      --qm-olive: #3d3d3a;       /* Body */
      --qm-stone: #6c6a64;       /* Muted */
      --qm-warm-silver: #8e8b82; /* Muted Soft */

      /* Spacing (s): 8px base editorial scale */
      --qm-s-xs: 4px;
      --qm-s-s: 8px;
      --qm-s-m: 12px;
      --qm-s-l: 16px;
      --qm-s-xl: 24px;

      /* Radius (r): Claude's rounded scale */
      --qm-r-s: 4px;
      --qm-r-m: 8px;
      --qm-r-l: 12px;
      --qm-r-xl: 16px;
      --qm-r-full: 9999px;

      /* Fonts: Serif for Authority, Sans for Utility */
      --qm-font-serif: "Anthropic Serif", Georgia, "Times New Roman", serif;
      --qm-font-sans: "Anthropic Sans", Inter, -apple-system, BlinkMacSystemFont, Arial, sans-serif;
      --qm-font-mono: "Anthropic Mono", "SF Mono", "Roboto Mono", monospace;

      --qm-font-xs: 11px;
      --qm-font-s: 13px;
      --qm-font-m: 15px;
      --qm-font-l: 17px;
      --qm-font-xl: 20px;

      /* Transitions: Quiet & Deliberate */
      --qm-t-fast: 0.15s ease;
      --qm-t-med: 0.25s cubic-bezier(0.4, 0, 0.2, 1);
      --qm-t-panel: 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.15);

      /* --- Semantic Variables (Light Default) --- */
      --qm-bg: var(--qm-parchment);
      --qm-bg-alt: var(--qm-ivory);
      --qm-surface: var(--qm-ivory);
      --qm-pane-bg: var(--qm-parchment);
      --qm-text: var(--qm-near-black);
      --qm-text-muted: var(--qm-olive);
      --qm-border: var(--qm-cream);        /* Border Cream #f0eee6 */
      --qm-border-warm: var(--qm-sand);   /* Border Warm #e8e6dc */
      --qm-ring: #d1cfc5;                 /* Ring Warm — DESIGN.md */
      --qm-ring-subtle: #e8e6dc;          /* Ring Subtle */
      --qm-ring-deep: #c2c0b6;            /* Ring Deep — active/pressed */
      --qm-shadow: rgba(20,20,19,0.05);   /* Whisper shadow */
      --qm-shadow-panel: rgba(20,20,19,0.15);
      --qm-overlay: rgba(20,20,19,0.3);
      --qm-input-border: var(--qm-sand);
      --qm-btn-bg: var(--qm-sand);
      --qm-btn-text: var(--qm-charcoal);
    }

    /* --- Dark Mode Semantic Overrides (Near Black) --- */
    #qm-panel.qm-dark, #qm-fab.qm-dark, #qm-global-loader.qm-dark {
      --qm-bg: #0a0a0a;
      --qm-bg-alt: #141413;
      --qm-surface: #1c1c1b;
      --qm-pane-bg: #0a0a0a;
      --qm-text: #ffffff;
      --qm-text-muted: #a09e95;
      --qm-border: #262624;
      --qm-border-warm: #30302e;
      --qm-ring: #30302e;
      --qm-ring-subtle: #262624;
      --qm-ring-deep: #404040;
      --qm-shadow: rgba(0,0,0,0.8);
      --qm-shadow-panel: rgba(0,0,0,0.9);
      --qm-danger-bg: #4a1414;
      --qm-input-border: #30302e;
      --qm-input-bg: #141413;
      --qm-btn-bg: #1c1c1b;
      --qm-btn-text: #ffffff;
      
      /* Force remap for all legacy/utility uses */
      --qm-ivory: #0a0a0a;
      --qm-parchment: #0a0a0a;
      --qm-white: #141413;
      --qm-near-black: #ffffff;
      --qm-charcoal: #f0eee6;
      --qm-olive: #b0aea5;
      --qm-stone: #87867f;
      --qm-cream: #1c1c1b;
      --qm-sand: #262624;
    }

    #qm-panel.qm-dark {
      color: #ffffff;
      
      .qm-pane, .qm-card, .qm-input, .qm-select, .qm-textarea, .qm-section-title, .qm-field-label {
        color: #ffffff !important;
      }
      
      .qm-input::placeholder, .qm-textarea::placeholder {
        color: rgba(255, 255, 255, 0.4) !important;
      }
      
      .qm-select option {
        background: #141413;
        color: #ffffff;
      }

      #qm-header h6 { color: #ffffff !important; }
    }

    /* --- Base & Utilities --- */
    .qm-flex { display: flex; }
    .qm-flex-col { display: flex; flex-direction: column; }
    .qm-items-center { align-items: center; }
    .qm-items-end { align-items: flex-end; }
    .qm-justify-center { justify-content: center; }
    .qm-justify-between { justify-content: space-between; }
    .qm-flex-1 { flex: 1; }
    .qm-flex-1-5 { flex: 1.5; }
    .qm-w-full { width: 100%; }
    .qm-w-140 { width: 140px !important; }
    .qm-gap-xs { gap: var(--qm-s-xs); }
    .qm-gap-s { gap: var(--qm-s-s); }
    .qm-gap-m { gap: var(--qm-s-m); }
    .qm-gap-l { gap: var(--qm-s-l); }
    .qm-mb-0 { margin-bottom: 0 !important; }
    .qm-mb-s { margin-bottom: var(--qm-s-s); }
    .qm-mb-m { margin-bottom: var(--qm-s-m); }
    .qm-mb-l { margin-bottom: var(--qm-s-l); }
    .qm-mb-xl { margin-bottom: var(--qm-s-xl); }
    .qm-mt-s { margin-top: var(--qm-s-s); }
    .qm-mt-m { margin-top: var(--qm-s-m); }
    .qm-ml-s { margin-left: var(--qm-s-s); }
    .qm-m-0 { margin: 0 !important; }
    .qm-block { display: block !important; }
    .qm-text-center { text-align: center; }
    .qm-font-bold { font-weight: 700; }
    .qm-font-semibold { font-weight: 600; }
    .qm-lh-1-3 { line-height: 1.3; }
    .qm-cursor-pointer { cursor: pointer !important; }
    .qm-select-none { user-select: none !important; }
    .qm-hidden { display: none !important; }
    .qm-visible-block { display: block !important; }
    .qm-visible-flex { display: flex !important; }
    .qm-visible-inline-flex { display: inline-flex !important; }
    .qm-no-scroll { overflow: hidden !important; }

    /* --- Typographic Foundation --- */
    #qm-panel, #qm-fab, #qm-global-loader {
      font-family: var(--qm-font-sans);
      -webkit-font-smoothing: antialiased;
      
      h1, h2, h3, h4, h5, h6, .qm-section-title { 
        font-family: var(--qm-font-serif); 
        font-weight: 500; 
        letter-spacing: -0.01em;
        line-height: 1.2;
      }
    }

    /* --- Text Styling --- */
    .qm-text-muted { color: var(--qm-text-muted); }
    .qm-text-primary { color: var(--qm-p-500) !important; }
    .qm-text-success { color: var(--qm-success) !important; }
    .qm-text-danger { color: var(--qm-danger) !important; }
    .qm-text-warning { color: var(--qm-warning) !important; }
    .qm-text-s { font-size: var(--qm-font-s) !important; }

    .qm-card {
      background: var(--qm-white);
      border-radius: var(--qm-r-l);
      padding: var(--qm-s-m) var(--qm-s-l);
      box-shadow: 0px 2px 8px var(--qm-shadow);
      border: 1px solid var(--qm-border-warm);
      &.qm-card-bordered { border-style: solid; }
      &.qm-card-dashed { border-style: dashed; }
    }

    .qm-section-title {
      font-family: var(--qm-font-serif);
      font-size: var(--qm-font-m);
      font-weight: 600;
      color: var(--qm-near-black);
      margin: 0 0 12px 0;
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .qm-pane-header {
      display: flex;
      align-items: center;
      gap: 12px;
      margin-bottom: 20px;
      padding-bottom: 12px;
      border-bottom: 1px solid var(--qm-border-warm);
      h6 {
        margin: 0;
        font-family: var(--qm-font-serif);
        font-size: 18px;
        font-weight: 600;
        color: var(--qm-near-black);
      }
      svg { color: var(--qm-p-500); }
    }

    .qm-badge {
      font-family: var(--qm-font-sans);
      font-size: var(--qm-font-xs);
      padding: 2px 8px;
      border-radius: var(--qm-r-full);
      font-weight: 600;
      white-space: nowrap;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      
      &.ok { background: var(--qm-success-bg); color: var(--qm-success); }
      &.err { background: var(--qm-danger-bg); color: var(--qm-danger); }
      &#qm-badge-anomali {
        display: inline-flex; align-items:center; justify-content:center; 
        min-width:16px; height:16px; background:var(--qm-danger); 
        color:#fff; padding:2px; border-radius:50%; 
        font-size:9px; font-weight:bold; line-height:1;
        position: absolute; top: 2px; right: 2px;
      }
    }

    /* --- Form Elements --- */
    .qm-field-label {
      font-size: 10px;
      font-weight: 700;
      color: var(--qm-stone);
      margin-bottom: var(--qm-s-xs);
      display: block;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      &.qm-field-label-normal { font-weight: 500; text-transform: none; letter-spacing: normal; font-size: var(--qm-font-m); }
    }
    
    .qm-input, .qm-select, .qm-textarea {
      font-size: var(--qm-font-m);
      border: 1px solid var(--qm-input-border);
      border-radius: var(--qm-r-l);
      padding: 0 14px;
      width: 100%;
      height: 42px; /* Fixed editorial height */
      box-sizing: border-box;
      background: var(--qm-input-bg, var(--qm-white));
      color: var(--qm-text);
      transition: border-color var(--qm-t-fast), box-shadow var(--qm-t-fast), background-color var(--qm-t-fast);
      outline: none;
      font-family: var(--qm-font-sans);

      &:focus {
        border-color: var(--qm-p-500);
        box-shadow: 0 0 0 3px rgba(201, 100, 66, 0.15);
      }
    }
    .qm-textarea { height: 80px; resize: vertical; padding: 12px 14px; line-height: 1.6; &.qm-textarea-mono { font-family: var(--qm-font-mono); } }
    .qm-select.qm-text-center { text-align-last: center; }
    .qm-input-shortcut { cursor: pointer; text-align: center; font-weight: 600; }
    .qm-config-checkbox { width: 18px; height: 18px; cursor: pointer; accent-color: var(--qm-p-500); }
    .qm-field-time, .qm-field-shift { padding: 0 12px; }

    /* --- Buttons --- */
    .qm-btn {
      border: none;
      border-radius: var(--qm-r-l);
      font-size: var(--qm-font-m);
      font-weight: 500;
      padding: 0 20px;
      height: 42px;
      cursor: pointer;
      transition: background-color var(--qm-t-fast), transform var(--qm-t-fast), box-shadow var(--qm-t-fast), color var(--qm-t-fast);
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      width: 100%;
      color: var(--qm-btn-text);
      background: var(--qm-btn-bg);
      box-sizing: border-box;
      box-shadow: 0px 0px 0px 1px var(--qm-ring);

      &:hover:not(:disabled) { 
        background: var(--qm-white); 
        box-shadow: 0px 0px 0px 1px var(--qm-ring-subtle), 0px 4px 12px var(--qm-shadow);
      }
      &:focus-visible {
        outline: none;
        box-shadow: 0 0 0 3px rgba(201, 100, 66, 0.4);
      }
      &:disabled { opacity: 0.5; cursor: not-allowed; }
      
      &.qm-btn-primary { 
        background: var(--qm-p-500); 
        color: #ffffff !important; 
        box-shadow: 0px 1px 2px rgba(201, 100, 66, 0.2);
        &:hover { 
          background: var(--qm-p-600); 
          box-shadow: 0px 4px 12px rgba(201, 100, 66, 0.2);
        }
      }
      &.qm-btn-dark { 
        background: var(--qm-dark-surface); 
        color: var(--qm-warm-silver); 
        box-shadow: 0px 0px 0px 1px var(--qm-dark-surface);
      }
      &.qm-btn-secondary { 
        background: var(--qm-white); 
        color: var(--qm-near-black); 
        box-shadow: 0px 0px 0px 1px var(--qm-ring); 
        &:hover:not(:disabled) { background: var(--qm-sand); } 
      }
      &.qm-btn-outline { 
        background: transparent; 
        color: var(--qm-charcoal); 
        box-shadow: 0px 0px 0px 1px var(--qm-ring); 
        &:hover:not(:disabled) { background: var(--qm-sand); } 
      }
      &.qm-btn-success { background: var(--qm-success); color: #fff; box-shadow: 0px 0px 0px 1px var(--qm-success); }
      &.qm-btn-record { width: auto; min-width: 80px; }
      &.qm-btn-mini {
        height: 24px; width: 24px; padding: 0;
        background: var(--qm-bg-alt);
        color: var(--qm-stone);
        border-radius: 50%;
        box-shadow: 0px 0px 0px 1px var(--qm-border-warm);
        display: inline-flex; align-items: center; justify-content: center;
        &:hover {
          background: var(--qm-sand);
          color: var(--qm-near-black);
          box-shadow: 0px 0px 0px 1px var(--qm-ring);
        }
        svg { width: 12px; height: 12px; }
      }
    }

    /* --- FAB & Backdrop --- */
    #qm-fab {
      position: fixed; bottom: 32px; right: 32px; z-index: 99999;
      width: 58px; height: 58px; background: var(--qm-p-500);
      border-radius: 50%; color: #fff; cursor: pointer; border: none;
      box-shadow: 0 4px 20px rgba(201, 100, 66, 0.3);
      display: flex; align-items: center; justify-content: center;
      transition: transform var(--qm-t-med), background-color var(--qm-t-med), box-shadow var(--qm-t-med);

      &:hover { transform: scale(1.05); background: var(--qm-p-600); box-shadow: 0 6px 24px rgba(201, 100, 66, 0.4); }
      &:active { transform: scale(0.95); }
      &.qm-open {
        .qm-icon-menu { opacity: 0; transform: rotate(45deg) scale(0.5); }
        .qm-icon-close { opacity: 1; transform: rotate(0deg) scale(1); }
      }

      .qm-icon-menu, .qm-icon-close {
        position: absolute; display: flex; align-items: center; justify-content: center;
        transition: transform var(--qm-t-med), opacity var(--qm-t-med);
      }
      .qm-icon-menu { opacity: 1; transform: rotate(0deg) scale(1); }
      .qm-icon-close { opacity: 0; transform: rotate(-45deg) scale(0.5); }
    }

    #qm-backdrop {
      position: fixed; top: 0; left: 0; right: 0; bottom: 0;
      z-index: 99997; opacity: 0; pointer-events: none; transition: opacity var(--qm-t-med);
      /* Use hardware acceleration for filter if kept, but blur is heavy */
      -webkit-backdrop-filter: blur(4px);
      &.qm-open { opacity: 1; pointer-events: auto; }
    }

    /* --- Main Panel --- */
    #qm-panel {
      position: fixed; top: 60px; left: 50%; z-index: 99998;
      width: 90%; max-width: 760px; min-width: 320px;
      height: auto; max-height: calc(100vh - 100px);
      display: flex; flex-direction: column;
      background: var(--qm-bg); border-radius: var(--qm-r-xl); box-shadow: 0 20px 80px var(--qm-shadow-panel);
      overflow: hidden; color: var(--qm-text);
      transform: translateX(-50%) translateY(10px); opacity: 0; pointer-events: none;
      transition: transform var(--qm-t-panel), opacity var(--qm-t-panel);
      will-change: transform, opacity;
      
      &.qm-open { transform: translateX(-50%) translateY(0); opacity: 1; pointer-events: auto; }

      @media (min-width: 1200px) {
        max-width: 1200px;
        #qm-sidebar { width: 72px; .qm-tab { width: 52px; height: 52px; svg { width: 24px; height: 24px; } } }
        .qm-pane { padding: 24px 32px; }
      }

      @media (max-width: 768px) {
        top: 20px;
        width: 95%;
        max-height: calc(100vh - 40px);
        #qm-header { padding: 12px 16px; h6 { font-size: 16px; } }
        #qm-sidebar { width: 50px; .qm-tab { width: 36px; height: 36px; svg { width: 18px; height: 18px; } } }
        .qm-pane { padding: 14px 16px; }
      }

      #qm-header {
        background: var(--qm-ivory); border-bottom: 1px solid var(--qm-border-warm);
        padding: 16px 20px 16px 24px; display: flex; align-items: center; gap: 12px;
        flex-shrink: 0;

        /* Brand mark — spike icon area */
        .qm-header-brand { display: flex; align-items: center; gap: 10px; flex: 1; min-width: 0; }
        .qm-header-mark {
          width: 32px; height: 32px; background: var(--qm-p-500);
          border-radius: var(--qm-r-m);
          display: flex; align-items: center; justify-content: center; flex-shrink: 0;
          svg { color: #fff; }
        }
        .qm-header-text { min-width: 0; }
        h6 { margin: 0; color: var(--qm-text); font-size: 18px; font-weight: 500; letter-spacing: -0.01em; white-space: nowrap; }
        small { color: var(--qm-text-muted); font-size: var(--qm-font-xs); display: block; margin-top: 2px; text-transform: uppercase; letter-spacing: 0.12em; }

        /* Right-side action cluster */
        .qm-header-actions { display: flex; align-items: center; gap: 8px; flex-shrink: 0; }
        .qm-header-icon-btn {
          width: 32px; height: 32px; border-radius: 50%;
          border: 1px solid var(--qm-border-warm); background: transparent;
          display: flex; align-items: center; justify-content: center;
          cursor: pointer; color: var(--qm-stone); transition: all var(--qm-t-fast);
          &:hover { background: var(--qm-sand); color: var(--qm-near-black); }
        }
      }

      #qm-panel-body {
        display: flex; flex: 1 1 auto; overflow: hidden;
      }

      /* --- Sidebar Nav --- */
      #qm-sidebar {
        display: flex; flex-direction: column; align-items: center;
        background: var(--qm-ivory); padding: 12px 0; gap: 8px;
        border-right: 1px solid var(--qm-border-warm);
        flex-shrink: 0; width: 64px;
        overflow-y: auto; scrollbar-width: none;
        &::-webkit-scrollbar { display: none; }

        .qm-tab {
          display: flex; align-items: center; justify-content: center;
          width: 44px; height: 44px;
          color: var(--qm-stone); cursor: pointer; border: none; background: transparent;
          border-radius: var(--qm-r-m);
          transition: background-color var(--qm-t-fast), color var(--qm-t-fast); position: relative;

          &.active {
            background: var(--qm-sand);
            color: var(--qm-p-500);
          }
          &:focus-visible {
            outline: none;
            background: var(--qm-sand);
            box-shadow: inset 0 0 0 2px var(--qm-p-500);
          }
          &:hover:not(.active) {
            background: var(--qm-cream);
            color: var(--qm-charcoal);
          }

          /* Loading pulse — bottom strip inside the pill */
          &.qm-tab-loading::after {
            content: ''; position: absolute; bottom: 6px; left: 10px; right: 10px;
            height: 2px; background: var(--qm-p-500);
            animation: qm-pulse-width 1.2s ease-in-out infinite; border-radius: 2px;
          }

          svg { flex-shrink: 0; opacity: 0.8; width: 20px; height: 20px; }
          &.active svg, &:hover svg { opacity: 1; }

          /* Tooltip */
          &:hover::before {
            content: attr(title);
            position: absolute; left: 54px;
            background: var(--qm-near-black); color: var(--qm-white);
            padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: 500;
            white-space: nowrap; z-index: 100;
            pointer-events: none;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
          }
        }
      }

      #qm-content-area {
        flex: 1 1 auto; position: relative; overflow-y: auto; background: var(--qm-bg);
        &::-webkit-scrollbar { width: 8px; }
        &::-webkit-scrollbar-track { background: var(--qm-bg); }
        &::-webkit-scrollbar-thumb { background: var(--qm-sand); border-radius: 10px; border: 2px solid var(--qm-bg); }
        &::-webkit-scrollbar-thumb:hover { background: var(--qm-stone); }
      }

      .qm-pane {
        display: none; padding: 16px 20px;
        &.active { display: block; }
      }
    }

    /* --- Interactive Elements --- */
    .qm-accordion-header {
      transition: background 0.2s;
      outline: none;
      
      &:focus-visible { background: var(--qm-sand); box-shadow: inset 0 0 0 2px var(--qm-p-500); }
      &:hover { background: var(--qm-sand); }
      &.expanded { border-bottom-color: transparent; }
      &.expanded .qm-accordion-chevron { transform: rotate(180deg); }
    }

    .qm-accordion-chevron {
      transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1); 
      display: flex; align-items: center; justify-content: center; 
      width: 28px; height: 28px; background: var(--qm-white); 
      border-radius: 50%; color: var(--qm-stone); transform: rotate(0deg);
      box-shadow: 0px 0px 0px 1px var(--qm-cream);
      &.qm-chevron { width: 18px; height: 18px; background: none !important; box-shadow: none !important; }
    }

    .qm-anomaly-group-content {
      max-height: 0; overflow: hidden; opacity: 0; padding: 0 16px;
      transition: max-height 0.4s cubic-bezier(0.4, 0, 0.2, 1), opacity 0.4s ease;
      background: var(--qm-bg);
      &.qm-content-open {
        max-height: 1000px; opacity: 1; padding: 16px;
        border-bottom: 1px solid var(--qm-border-warm);
      }
    }

    /* --- Global Loader Toast --- */
    #qm-global-loader {
      position: fixed; bottom: 40px; left: 50%; transform: translateX(-50%); z-index: 99999;
      background: var(--qm-surface); border-radius: var(--qm-r-xl); 
      padding: var(--qm-s-l) var(--qm-s-xl);
      box-shadow: 0 20px 80px rgba(0,0,0,0.3);
      display: flex; flex-direction: column; gap: var(--qm-s-m);
      border: 1px solid var(--qm-border-warm);
      width: 90%; max-width: 500px;
      animation: qm-toast-in-bottom 0.5s cubic-bezier(0.175, 0.885, 0.32, 1.1);
      transition: opacity 0.3s, transform 0.3s;
      
      &.qm-loader-hiding { opacity: 0; transform: translateX(-50%) translateY(20px) scale(0.95); }
      
      .qm-loader-header { display: flex; align-items: center; gap: var(--qm-s-xl); }
      .qm-loader-body { display: flex; flex-direction: column; flex: 1; }
      .qm-loader-title { font-weight: 700; font-size: 16px; color: var(--qm-near-black); line-height: 1.2; margin-bottom: 4px; font-family: var(--qm-font-serif); }
      .qm-loader-text { font-size: 11px; font-weight: 600; color: var(--qm-stone); line-height: 1.2; text-transform: uppercase; letter-spacing: 0.08em; }
      
      .qm-loader-footer { 
        display: flex; justify-content: center; margin-top: var(--qm-s-xs);
        border-top: 1px solid var(--qm-border); padding-top: var(--qm-s-s);
      }
    }

    .qm-progress-container {
      position: relative; height: 12px; background: var(--qm-sand); overflow: hidden;
      border-radius: var(--qm-r-full); margin: 12px 0; cursor: pointer;
      box-shadow: inset 0 1px 2px rgba(0,0,0,0.1);
      .qm-progress-bar { 
        height: 100%; width: 0%; background: linear-gradient(90deg, var(--qm-p-500), var(--qm-p-600)); 
        transition: width 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        box-shadow: 0 0 10px rgba(201, 100, 66, 0.3);
      }
      &:hover { filter: brightness(0.98); transform: translateY(-1px); }
      &:active { transform: translateY(0); }
    }

    /* --- Batch Results Table --- */
    .qm-batch-table {
      width: 100%; border-collapse: collapse; font-size: var(--qm-font-m); table-layout: fixed;
      .qm-batch-cell { padding: 12px 16px; vertical-align: top; border-bottom: 1px solid var(--qm-cream); }
      .qm-batch-cell-header { padding: 8px 16px; font-weight: 600; color: var(--qm-stone); font-size: var(--qm-font-xs); text-transform: uppercase; letter-spacing: 0.1em; border-bottom: 2px solid var(--qm-sand); }
      .qm-batch-bagian-cell { padding: 16px; font-weight: 500; font-family: var(--qm-font-serif); border-bottom: 1.5px solid var(--qm-sand); color: var(--qm-near-black); font-size: 20px; }
      .qm-batch-seksi-cell { padding: 12px 16px 12px 32px; font-weight: 600; background: var(--qm-ivory); border-bottom: 1px solid var(--qm-cream); color: var(--qm-olive); font-size: var(--qm-font-m); }
      .qm-batch-nrp-link { color: var(--qm-p-500); font-weight: 700; text-decoration: none; border-bottom: 1px solid transparent; &:hover { border-bottom-color: var(--qm-p-500); } }
      .qm-batch-nama { font-weight: 600; font-size: var(--qm-font-s); }
      .qm-batch-fix-btn { background: var(--qm-p-500); color: #fff; border: none; border-radius: var(--qm-r-m); padding: 4px 10px; font-size: 11px; cursor: pointer; white-space: nowrap; margin-left: 8px; font-weight: 600; }
      .qm-batch-date-content { display: none; padding-left: 16px; margin-bottom: 8px; border-left: 2px solid var(--qm-p-500); }
      .qm-batch-anomaly-detail { color: var(--qm-olive); font-size: var(--qm-font-s); padding: 2px 0; }
    }

    #qm-result {
      margin-top: 24px; border-radius: var(--qm-r-m); padding: 16px 20px;
      font-size: var(--qm-font-m); display: none; animation: qm-fadein .25s ease;
      background: var(--qm-ivory); border: 1px solid var(--qm-sand);
      &.success { border-left: 4px solid var(--qm-success); color: var(--qm-success); }
      &.danger  { border-left: 4px solid var(--qm-danger); color: var(--qm-danger); }
      &.warning { border-left: 4px solid var(--qm-warning); color: #e65100; }
    }

    /* --- History --- */
    #qm-history {
      max-height: 180px; overflow-y: auto; margin-top: 24px; font-size: var(--qm-font-s);
      .qm-history-item { 
        display: flex; align-items: center; gap: 10px; padding: 10px 0; border-bottom: 1px solid var(--qm-cream);
        &:last-child { border-bottom: none; }
        &:hover .qm-history-label { color: var(--qm-near-black); }
      }
      .qm-history-label { color: var(--qm-stone); flex: 1; transition: color 0.2s; }
    }

    /* --- Special Features --- */
    .qm-anomaly-cell { border: 2px solid var(--qm-danger) !important; background-color: var(--qm-danger-bg) !important; position: relative !important; }
    .qm-fix-dot { position: absolute; top: 2px; right: 2px; width: 8px; height: 12px; background-color: var(--qm-p-500); border-radius: 100%; cursor: pointer; box-shadow: 0 0 4px var(--qm-shadow); }
    .qm-info-box { font-size: 11px; line-height: 1.6; padding: 16px; background: var(--qm-white); border: 1px solid var(--qm-cream); border-radius: var(--qm-r-m); color: var(--qm-stone); }
    .qm-ot7-box { padding: 20px; background: var(--qm-white); border-radius: var(--qm-r-l); border: 1px solid var(--qm-cream); box-shadow: 0px 0px 0px 1px var(--qm-ring); margin: 12px 0; animation: qm-fadein 0.3s ease; }
    
    #qm-modal-jk {
      position: absolute; top: 100%; left: 0; z-index: 1000;
      width: 400px; background: var(--qm-white); border: 1px solid var(--qm-cream);
      border-radius: var(--qm-r-l); box-shadow: 0 10px 40px var(--qm-shadow-panel);
      padding: 24px; margin-top: 8px; animation: qm-fadein 0.2s ease;
    }

    /* --- Animations --- */
    .qm-skeleton {
      background: linear-gradient(90deg, var(--qm-sand) 25%, var(--qm-cream) 50%, var(--qm-sand) 75%);
      background-size: 200% 100%; animation: qm-shimmer 2s infinite;
      border-radius: 4px; display: inline-block; min-width: 60px; height: 14px;
    }
    @keyframes qm-toast-in { from { opacity: 0; transform: translate3d(40px, 0, 0) scale(0.95); } to { opacity: 1; transform: translate3d(0, 0, 0) scale(1); } }
    @keyframes qm-toast-in-bottom { from { opacity: 0; transform: translate3d(-50%, 40px, 0) scale(0.9); } to { opacity: 1; transform: translate3d(-50%, 0, 0) scale(1); } }
    @keyframes qm-fadein { from { opacity:0; transform:translate3d(0, 8px, 0); } to { opacity:1; transform:translate3d(0, 0, 0); } }
    @keyframes qm-shimmer { 0% { background-position: 200% 0; } 100% { background-position: -200% 0; } }
    @keyframes qm-pulse-width { 0%, 100% { transform: scaleX(0.2); opacity: 0.3; } 50% { transform: scaleX(1); opacity: 1; } }
    @keyframes qm-spin { to { transform: rotate(360deg); } }

    .qm-spinner {
      width: 18px; height: 18px;
      border: 2px solid var(--qm-sand); border-top-color: var(--qm-p-500);
      border-radius: 50%; animation: qm-spin .8s linear infinite;
      display: inline-block; vertical-align: middle;
      &.qm-spinner-white { border-color: rgba(255,255,255,0.2); border-top-color: #fff; }
    }

    .qm-anomaly-item {
      display: flex; justify-content: space-between; align-items: center;
      padding: 12px 20px; background: var(--qm-bg); border-bottom: 1px solid var(--qm-border);
      gap: 16px; animation: qm-fadein 0.3s ease;
      &:last-child { border-bottom: none; }
    }
    .qm-anomaly-left { display: flex; align-items: center; gap: 8px; flex-shrink: 0; min-width: 70px; }
    .qm-anomaly-date { font-weight: 700; font-size: 15px; color: var(--qm-near-black); font-family: var(--qm-font-serif); }
    .qm-anomaly-content { display: flex; flex-wrap: wrap; gap: 8px; flex-grow: 1; align-items: center; }
    .qm-anomaly-actions { flex-shrink: 0; }
    .qm-btn-fix-pill {
      height: 28px; padding: 0 14px; border-radius: 14px;
      border: 1px solid var(--qm-ring); background: var(--qm-white);
      color: var(--qm-p-500); font-size: 11px; font-weight: 600;
      cursor: pointer; transition: all 0.2s; white-space: nowrap;
      &:hover { background: var(--qm-p-500); color: #fff; border-color: var(--qm-p-500); box-shadow: 0 2px 8px var(--qm-shadow); }
    }
    .qm-anomaly-card {
      background: var(--qm-ivory); padding: 6px 10px; border-radius: 4px;
      width: auto; min-width: 120px; text-align: left; border: 1px solid var(--qm-border-warm);
      box-shadow: 0 1px 2px rgba(0,0,0,0.01);
    }
    .qm-anomaly-card-type { font-weight: 700; font-size: 12px; color: var(--qm-near-black); margin-bottom: 1px; }
    .qm-anomaly-card-msg { font-size: 10px; color: var(--qm-stone); line-height: 1.1; }

    .qm-theme-btn {
      flex: 1; height: 40px; font-weight: 600; font-size: 13px;
      border: 1px solid var(--qm-border-warm); border-radius: var(--qm-r-m);
      display: flex; align-items: center; justify-content: center; gap: 8px;
      cursor: pointer; transition: all 0.2s;
      &.light { background: #fdfaf3; color: #141413; }
      &.dark { background: #141413; color: #efe9de; border-color: #30302e; }
      &.active { box-shadow: 0 0 0 2px var(--qm-p-500) !important; border-color: var(--qm-p-500) !important; }
    }

    /* --- Log Activity (Terminal Style) --- */
    .qm-log-body-inline {
      margin-top: 8px;
      padding: 16px;
      background: #181715; /* Surface Dark */
      border: 1px solid #30302e; /* Hairline for dark */
      border-radius: var(--qm-r-l);
      max-height: 350px;
      overflow-y: auto;
      font-family: var(--qm-font-mono);
      font-size: 12px;
      line-height: 1.6;
      color: #faf9f5; /* On Dark */
      box-shadow: inset 0 2px 8px rgba(0,0,0,0.4);
      scrollbar-width: thin;
      scrollbar-color: #3d3d3a transparent;
    }
    .qm-log-body-inline::-webkit-scrollbar { width: 6px; }
    .qm-log-body-inline::-webkit-scrollbar-thumb { background: #3d3d3a; border-radius: 3px; }
    
    .qm-log-item {
      margin-bottom: 8px;
      padding-bottom: 8px;
      border-bottom: 1px solid rgba(255,255,255,0.05);
      display: flex;
      gap: 12px;
      align-items: flex-start;
    }
    .qm-log-time {
      color: #8e8b82; /* Muted Soft */
      flex-shrink: 0;
      user-select: none;
      min-width: 65px;
    }
    .qm-log-msg {
      word-break: break-word;
      color: #faf9f5;
    }
    .qm-log-msg.error, .qm-log-msg.danger { color: #c64545; }
    .qm-log-msg.success { color: #5db872; }
    .qm-log-msg.warn { color: #e8a55a; /* Accent Amber */ }
    .qm-log-msg.debug { color: #6c6a64; }

    .qm-log-actions {
      padding: 0 4px;
      align-items: flex-end;
    }
    .qm-btn-text {
      background: none; border: none; padding: 2px 6px;
      color: var(--qm-text-muted); cursor: pointer;
      border-radius: var(--qm-r-s);
      transition: all 0.2s;
    }
    .qm-btn-text:hover { background: var(--qm-cream); color: var(--qm-near-black); }
    .qm-btn-text.qm-text-danger:hover { background: #fdf2f2; color: #b53333; }

    #qm-btn-show-logs.qm-active {
      background: #181715 !important;
      color: #faf9f5 !important;
      border-color: #181715 !important;
    }

    /* --- Grid & Layout Utilities --- */
    .qm-grid-spkl {
      display: grid;
      grid-template-columns: 1.6fr 1fr;
      gap: var(--qm-s-l);
    }
    @media (max-width: 1024px) {
      .qm-grid-spkl { grid-template-columns: 1fr; }
    }

    .qm-spike-mark {
      color: var(--qm-p-500);
      flex-shrink: 0;
      width: 14px; height: 14px;
      margin-right: 10px;
      display: inline-block;
      vertical-align: middle;
    }

    .qm-serif {
      font-family: var(--qm-font-serif) !important;
      letter-spacing: -0.02em !important;
    }

    .qm-feature-card {
      background: var(--qm-ivory);
      border: 1px solid var(--qm-sand);
      border-radius: var(--qm-r-l);
      padding: var(--qm-s-xl);
    }

    .qm-btn-coral {
      background: var(--qm-p-500);
      color: #fff;
      border: none;
      transition: background var(--qm-t-fast);
      &:hover { background: var(--qm-p-600); }
    }

    .qm-textarea-dark {
      background: var(--qm-dark-surface);
      color: var(--qm-p-100);
      border: 1px solid var(--qm-border-warm);
      &:focus { border-color: var(--qm-p-500); }
    }
  `);

  /* ============================================================
   * 5. STATE
   * ============================================================ */
  const state = {
    isOpen: false,
    loading: false,
    history: [],
    maxHistory: 8,
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
    batchActiveWorkers: 0,
    batchTotal: 0,
    batchAborted: false,
    batchProfile: { renderTotal: 0, itemDurations: [] },
    profileStats: {},
    profileFlags: {},
    expandedAnomalyGroups: new Set(),
    panelPos: JSON.parse(GM_getValue('qm_panel_pos', 'null')), // {top, left}
  };
  let shortcutKey = GM_getValue('qm_shortcut', 'Ctrl+Q');
  let alwaysCollapseMenu = GM_getValue('qm_always_collapse', false);
  let isRecordingShortcut = false;
  let cachedEditHtml = null; // Cache for editgeneral HTML to speed up save

  /* ============================================================
   * 6. HTML TEMPLATE
   * ============================================================ */
  /** Claude-style Spike Mark SVG (4-spoke radial). */
  const SPIKE_SVG = `<svg class="qm-spike-mark" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="4" stroke-linecap="round"><line x1="12" y1="4" x2="12" y2="20"/><line x1="4" y1="12" x2="20" y2="12"/><line x1="6.34" y1="6.34" x2="17.66" y2="17.66"/><line x1="6.34" y1="17.66" x2="17.66" y2="6.34"/></svg>`;

  const HTML = `
    <button id="qm-fab" title="Quick Menu" aria-label="Quick Menu">
      <div class="qm-icon-menu">
        <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><circle cx="17.5" cy="17.5" r="3.5"/><line x1="17.5" y1="15.5" x2="17.5" y2="19.5"/><line x1="15.5" y1="17.5" x2="19.5" y2="17.5"/></svg>
      </div>
      <div class="qm-icon-close">
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
      </div>
    </button>

    <div id="qm-backdrop" role="presentation"></div>
    <div id="qm-panel" role="dialog" aria-label="Quick Menu Panel">
      <div id="qm-header">
        <div class="qm-header-brand">
          <div class="qm-header-mark">
            ${SPIKE_SVG}
          </div>
          <div class="qm-header-text">
            <h6 class="qm-serif">Quick Menu</h6>
            <small>HRIS KTI</small>
          </div>
        </div>
        <div class="qm-header-actions">
          <button class="qm-header-icon-btn" id="qm-btn-close-header" title="Tutup panel" aria-label="Tutup panel">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
          </button>
        </div>
      </div>

      <div id="qm-panel-body">
        <div id="qm-sidebar" role="tablist">
          <button class="qm-tab active" data-pane="check-nrp" role="tab" title="Cek NRP">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
          </button>
          <button class="qm-tab" data-pane="spkl" role="tab" title="SPKL">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
          </button>
          <button class="qm-tab" data-pane="kehadiran" role="tab" title="Kehadiran">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
          </button>
          <button class="qm-tab" data-pane="distribusi" role="tab" title="Distribusi">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="17 1 21 5 17 9"></polyline><path d="M3 11V9a4 4 0 0 1 4-4h14"></path><polyline points="7 23 3 19 7 15"></polyline><path d="M21 13v2a4 4 0 0 1-4 4H3"></path></svg>
          </button>
          <button class="qm-tab" data-pane="anomali" role="tab" title="Anomali">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
            <span id="qm-badge-anomali"></span>
          </button>
          <button class="qm-tab" data-pane="config" role="tab" title="Pengaturan">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>
          </button>
        </div>
        <div id="qm-content-area">
          <div id="qm-pane-check-nrp" class="qm-pane active">
            <div class="qm-pane-header">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <h6 class="qm-serif">Cek NRP</h6>
            </div>
        <!-- 1. Header Bulan/Tahun -->
        <div class="qm-card qm-flex qm-items-center qm-justify-center qm-gap-m qm-mb-m">
          <div class="qm-flex-1">
            <select id="qm-input-bulan" class="qm-select qm-text-center qm-font-semibold"></select>
          </div>
          <div class="qm-flex-1">
            <select id="qm-input-tahun" class="qm-select qm-text-center qm-font-semibold"></select>
          </div>
        </div>

        <!-- 2. Seksi Per NRP -->
        <div class="qm-card qm-mb-m qm-card-bordered">
          <h6 class="qm-section-title qm-mb-s qm-serif">${SPIKE_SVG} Per NRP</h6>
          <div class="qm-flex qm-items-end qm-gap-m">
            <div class="qm-w-140">
              <input id="qm-input-nrp" type="text" placeholder="NRP" maxlength="8" autocomplete="off" spellcheck="false" class="qm-input">
            </div>
            <div class="qm-flex-1">
               <button id="qm-btn-check" type="button" class="qm-btn qm-btn-coral"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg> Cek NRP</button>
            </div>
          </div>
          <div id="qm-result" class="qm-mt-s">
            <div id="qm-result-title"></div>
            <div id="qm-result-body"></div>
          </div>
          <div id="qm-history" class="qm-mt-s"></div>
        </div>

        <!-- 3. Seksi Banyak NRP -->
        <div class="qm-card qm-mb-m qm-card-dashed">
          <h6 class="qm-section-title qm-mb-s qm-serif">${SPIKE_SVG} Banyak NRP</h6>
          <div class="qm-mb-s">
            <textarea id="qm-input-multi-nrp" placeholder="Daftar NRP (pisah dengan enter/koma)" rows="4" class="qm-textarea"></textarea>
          </div>
          <button id="qm-btn-batch-check" type="button" class="qm-btn qm-btn-coral">Proses Batch</button>

          <div id="qm-batch-progress" class="qm-mt-m qm-hidden">
            <div class="qm-batch-progress-container">
              <div id="qm-batch-progress-bar" class="qm-progress-bar"></div>
            </div>
            <div id="qm-batch-status">Memproses... 0/0</div>
          </div>

          <div id="qm-batch-results" class="qm-mt-m qm-batch-results-container"></div>
          <button id="qm-btn-export-batch" type="button" class="qm-btn qm-btn-success qm-mt-s qm-batch-export-btn qm-hidden">Export Hasil (.xlsx)</button>
        </div>
      </div>

      <div id="qm-pane-spkl" class="qm-pane">
        <div class="qm-pane-header">
          ${SPIKE_SVG}
          <h6 class="qm-serif">SPKL</h6>
        </div>
        <div class="qm-w-full">
          <div class="qm-grid-spkl qm-mb-m">
            <!-- SECTION 0: SPKL Online -->
            <div class="qm-feature-card">
              <h6 class="qm-serif qm-mb-m qm-text-s">${SPIKE_SVG} SPKL Online</h6>
              <div class="qm-flex qm-gap-m qm-mb-m">
                <div class="qm-flex-1">
                  <input id="qm-spkl-online-nrp" type="text" placeholder="NRP" maxlength="8" autocomplete="off" class="qm-input">
                </div>
                <div class="qm-flex-1-5">
                  <input id="qm-spkl-online-date" type="date" class="qm-input">
                </div>
              </div>
              <button id="qm-btn-spkl-online-cek" type="button" class="qm-btn qm-btn-coral qm-w-full">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
                Cek SPKL Online
              </button>
            </div>

            <!-- INFO BOX -->
            <div class="qm-feature-card qm-text-muted qm-info-box">
              <h6 class="qm-serif qm-mb-s qm-text-s" style="color:var(--qm-stone)">${SPIKE_SVG} INFO KODE OT</h6>
              <div style="font-size:11px; line-height:1.6">
                <b>1:</b> BIASA | <b>2:</b> LONG | <b>3:</b> NONSTOP | <b>4:</b> AWAL<br>
                <b>5A/B/C:</b> NOREST | <b>6:</b> STANDBY | <b>7:</b> LAIN | <b>OT:</b> OVERTIME
              </div>
            </div>
          </div>

          <!-- SECTION 1: Per NRP -->
          <div class="qm-feature-card qm-mb-m">
            <h6 class="qm-serif qm-mb-m qm-text-s">${SPIKE_SVG} Per NRP</h6>

            <div class="qm-mb-m">
              <input id="qm-fix-spkl-nrp" type="text" placeholder="NRP (4 atau 8 digit)" maxlength="8" autocomplete="off" class="qm-input">
            </div>

            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <select id="qm-fix-spkl-bulan" class="qm-select qm-font-semibold"></select>
              </div>
              <div class="qm-flex-1">
                <select id="qm-fix-spkl-tahun" class="qm-select qm-font-semibold"></select>
              </div>
            </div>

            <div class="qm-mb-m">
              <label class="qm-field-label">Data Tanggal-KodeOT (dipisahkan koma):</label>
              <textarea id="qm-fix-spkl-data" placeholder="Contoh: 2-1, 5-OT, 10-3" rows="2" class="qm-textarea qm-textarea-mono"></textarea>
            </div>

            <!-- Extra Fields for OT 7 in Per NRP -->
            <div id="qm-fix-spkl-ot7-box" class="qm-mb-m qm-ot7-box qm-hidden">
              <div class="qm-flex qm-gap-m qm-mb-m">
                <div class="qm-flex-1">
                  <label class="qm-field-label qm-field-label-normal">Jam Awal:</label>
                  <input id="qm-fix-spkl-jam-awal" type="time" class="qm-input qm-field-time">
                </div>
                <div class="qm-flex-1">
                  <label class="qm-field-label qm-field-label-normal">Jam Akhir:</label>
                  <input id="qm-fix-spkl-jam-akhir" type="time" class="qm-input qm-field-time">
                </div>
              </div>
              <div>
                <label class="qm-field-label qm-field-label-normal">Shift:</label>
                <select id="qm-fix-spkl-shift" class="qm-select qm-field-shift">
                  <option value="1">SHIFT I</option>
                  <option value="2">SHIFT II</option>
                  <option value="3">SHIFT III</option>
                  <option value="4">LONG SHIFT I</option>
                  <option value="5">LONG SHIFT II</option>
                </select>
              </div>
            </div>

            <button id="qm-btn-spkl-batch" type="button" class="qm-btn qm-btn-coral qm-w-full">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/></svg>
              Proses Per NRP
            </button>
          </div>

          <!-- SECTION 2: Banyak NRP -->
          <div class="qm-feature-card">
            <h6 class="qm-serif qm-mb-m qm-text-s">${SPIKE_SVG} Banyak NRP</h6>
            
            <!-- NRP List -->
            <div class="qm-mb-m">
              <textarea id="qm-fix-many-nrps" placeholder="Daftar NRP (pisahkan koma atau baris)" rows="2" class="qm-textarea qm-textarea-mono qm-textarea-dark"></textarea>
            </div>

            <!-- Date & OT Row -->
            <div class="qm-flex qm-gap-m qm-mb-m">
              <div class="qm-flex-1">
                <input id="qm-fix-many-date" type="date" class="qm-input">
              </div>
              <div class="qm-flex-1">
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

            <!-- Extra Fields for OT 7 (Lain-lain) -->
            <div id="qm-fix-many-ot7-box" class="qm-mb-m qm-ot7-box qm-hidden">
              <div class="qm-flex qm-gap-m qm-mb-m">
                <div class="qm-flex-1">
                  <label class="qm-field-label qm-field-label-normal">Jam Awal:</label>
                  <input id="qm-fix-many-jam-awal" type="time" class="qm-input qm-field-time">
                </div>
                <div class="qm-flex-1">
                  <label class="qm-field-label qm-field-label-normal">Jam Akhir:</label>
                  <input id="qm-fix-many-jam-akhir" type="time" class="qm-input qm-field-time">
                </div>
              </div>
              <div>
                <label class="qm-field-label qm-field-label-normal">Shift:</label>
                <select id="qm-fix-many-shift" class="qm-select qm-field-shift">
                  <option value="1">SHIFT I</option>
                  <option value="2">SHIFT II</option>
                  <option value="3">SHIFT III</option>
                  <option value="4">LONG SHIFT I</option>
                  <option value="5">LONG SHIFT II</option>
                </select>
              </div>
            </div>

            <button id="qm-btn-spkl-many-nrp" type="button" class="qm-btn qm-btn-coral qm-w-full">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>
              Proses Banyak NRP
            </button>
          </div>
        </div>
      </div>

      <div id="qm-pane-kehadiran" class="qm-pane">
        <div class="qm-pane-header">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
          <h6>Kehadiran</h6>
        </div>
        <div class="qm-w-full">
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
            <button id="qm-btn-hadir-many-proses" type="button" class="qm-btn qm-btn-teal">Proses Banyak NRP</button>
          </div>
        </div>
      </div>

      <div id="qm-pane-distribusi" class="qm-pane">
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
      </div>
      <div id="qm-pane-anomali" class="qm-pane">
        <div class="qm-pane-header">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
          <h6>Anomali</h6>
        </div>
        <div id="qm-anomali-list">
          <div class="qm-text-center qm-text-muted qm-anomali-empty-state">Tabel kehadiran belum dimuat atau tidak ada anomali.</div>
        </div>
      </div>

      <div id="qm-pane-config" class="qm-pane">
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

        <button id="qm-btn-show-logs" class="qm-btn qm-btn-secondary qm-w-full">
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
      </div> <!-- End qm-panel-body -->
    </div>

    </div>
    
  `;


  /* ============================================================
   * 7. UI HELPERS
   * ============================================================ */
  function now() { return new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }); }
  function pushLog(msg, type = 'info') {
    const time = Logger._formatTime();
    state.batchLogs.push({ time, msg, level: type });
    if (state.batchLogs.length > 500) state.batchLogs.shift();

    const logBody = document.getElementById('qm-log-body');
    if (logBody) {
      const div = document.createElement('div');
      div.className = 'qm-log-item';
      setInnerHTML(div, `<span class="qm-log-time">[${time}]</span><span class="qm-log-msg ${type}">${escHtml(msg)}</span>`);
      logBody.appendChild(div);
      logBody.scrollTop = logBody.scrollHeight;
    }
  }


  function initDraggable() {
    const panel = document.getElementById('qm-panel');
    const header = document.getElementById('qm-header');
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

  /* ============================================================
   * 8. UI SERVICE
   * ============================================================ */
  const UI = {
    resultTimeout: null,

    showResult(type, title, bodyHtml) {
      clearTimeout(this.resultTimeout);
      const resultEl = document.getElementById('qm-result');
      if (resultEl) {
        resultEl.classList.remove('success', 'danger', 'warning', 'qm-hidden', 'qm-fade-in');
        resultEl.classList.add(type, 'qm-visible-block');

        // Trigger animation in next frame
        requestAnimationFrame(() => {
          resultEl.classList.add('qm-fade-in');
        });

        document.getElementById('qm-result-title').textContent = title;
        setInnerHTML(document.getElementById('qm-result-body'), bodyHtml);
        this.resultTimeout = setTimeout(() => this.hideResult(), 3500);
      }
    },

    hideResult() {
      clearTimeout(this.resultTimeout);
      const resultEl = document.getElementById('qm-result');
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
      const btn = document.getElementById('qm-btn-check');
      if (btn) {
        if (on) {
          btn.disabled = true;
          setInnerHTML(btn, '<span class="qm-spinner"></span> Mengarahkan...');
        } else {
          btn.disabled = false;
          setInnerHTML(btn, `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg> Cek`);
        }
      }
    },

    applyTheme(theme) {
      state.theme = theme;
      GM_setValue('qm_theme', theme);
      const panel = document.getElementById('qm-panel');
      const fab = document.getElementById('qm-fab');
      const loader = document.getElementById('qm-global-loader');

      const isDark = theme === 'dark';
      [panel, fab, loader].forEach(el => {
        if (el) el.classList.toggle('qm-dark', isDark);
      });

      // Update buttons in config
      const btnLight = document.getElementById('qm-btn-theme-light');
      const btnDark = document.getElementById('qm-btn-theme-dark');
      if (btnLight) btnLight.classList.toggle('active', !isDark);
      if (btnDark) btnDark.classList.toggle('active', isDark);
    },

    showGlobalLoader(title, initialMsg, allowCancel = false) {
      const existing = document.getElementById('qm-global-loader');
      if (existing) existing.remove();

      state.batchLogs = []; // Clear logs on new process
      const logBody = document.getElementById('qm-log-body');
      if (logBody) setInnerHTML(logBody, '');

      const cancelBtnHtml = allowCancel
        ? `<div class="qm-loader-footer">
             <button id="qm-global-cancel-btn" class="qm-btn qm-btn-secondary" style="width: 100%; padding: 6px 12px; font-size: 11px; height: auto;">Batalkan Proses</button>
           </div>`
        : '';

      document.body.insertAdjacentHTML('beforeend', `
        <div id="qm-global-loader" class="${state.theme === 'dark' ? 'qm-dark' : ''}">
          <div class="qm-loader-header">
            <div class="qm-spinner qm-spinner-dark qm-loader-spinner-size" style="width: 24px; height: 24px; border-width: 3px;"></div>
            <div class="qm-loader-body">
              <div class="qm-loader-title">${escHtml(title)}</div>
              <div id="qm-global-loader-text" class="qm-loader-text">${escHtml(initialMsg)}</div>
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

    setGlobalProgress(pct, msg) {
      if (msg) {
        const textEl = document.getElementById('qm-global-loader-text');
        if (textEl) textEl.textContent = msg;
        pushLog(msg);
      }
      const barEl = document.getElementById('qm-global-loader-bar');
      if (barEl) barEl.style.width = pct + '%';
    },

    hideGlobalLoader(delay = 800) {
      setTimeout(() => {
        const loader = document.getElementById('qm-global-loader');
        if (loader) {
          loader.classList.add('qm-loader-hiding');
          setTimeout(() => loader.remove(), 300);
        }
      }, delay);
    },

    renderHistory() {
      const histEl = document.getElementById('qm-history');
      if (!histEl) return;
      if (!state.history.length) { setInnerHTML(histEl, ''); return; }
      const items = state.history.map(h => `<div class="qm-history-item"><span class="qm-badge ${h.ok ? 'ok' : 'err'}">${h.ok ? '✓' : '✗'}</span><span class="qm-history-nrp">${escHtml(h.nrp)}</span><span class="qm-history-label">${escHtml(h.label)}</span><span class="qm-history-time">${escHtml(h.time)}</span></div>`).join('');
      setInnerHTML(histEl, items);
    },

    pushHistory(nrp, ok, label) {
      state.history.unshift({ nrp, ok, label, time: now() });
      if (state.history.length > state.maxHistory) state.history.pop();
      this.renderHistory();
    }
  };

  /* ============================================================
   * 9. ANOMALY HELPERS
   * ============================================================ */
  function addAnomaly(tgl, col, msg, link = '') {
    state.anomalies.push({ tgl, col, msg, link });
  }

  /* pushLog and toggleLogModal moved/deduplicated to UI Helpers */

  /* ============================================================
   * 10. BATCH CHECK NRPs
   * ============================================================ */
  function parseBatchNrps(text) {
    return text.split(/[\n,]+/).map(s => s.trim()).filter(s => /^\d{4}$|^\d{8}$/.test(s));
  }

  function runBatchCheck() {
    const inputMulti = document.getElementById('qm-input-multi-nrp');
    const inputBulan = document.getElementById('qm-input-bulan');
    const inputTahun = document.getElementById('qm-input-tahun');
    const btnCheck = document.getElementById('qm-btn-batch-check');
    const btnExport = document.getElementById('qm-btn-export-batch');
    const progress = document.getElementById('qm-batch-progress');
    const results = document.getElementById('qm-batch-results');

    const nrps = parseBatchNrps(inputMulti ? inputMulti.value : '');
    if (nrps.length === 0) { UI.showResult('warning', 'Data Tidak Valid', 'Masukkan NRP yang valid (4 atau 8 digit).'); return; }
    if (nrps.length > APP_CONFIG.BATCH_MAX_LIMIT) { UI.showResult('warning', 'Terlalu Banyak', `Maksimal ${APP_CONFIG.BATCH_MAX_LIMIT} NRP per batch.`); return; }

    const localBulan = parseInt(inputBulan ? inputBulan.value : '') || (new Date().getMonth() + 1);
    const localTahun = parseInt(inputTahun ? inputTahun.value : '') || new Date().getFullYear();
    state.batchBulan = Math.min(12, Math.max(1, localBulan));
    state.batchTahun = Math.min(2035, Math.max(2020, localTahun));
    state.batchTotal = nrps.length;
    state.batchAborted = false;
    const prof = startProfile('runBatchCheck:init');

    state.batchQueue = nrps.map(nrp => ({ nrp, status: 'pending', msg: '' }));
    state.batchResults = [];
    state.batchLogs = [];
    resetBatchProfile();
    const logBody = document.getElementById('qm-log-body');
    if (logBody) setInnerHTML(logBody, '');
    pushLog(`Memulai batch check untuk ${nrps.length} NRP...`);

    if (btnCheck) { btnCheck.dataset.running = 'true'; btnCheck.textContent = 'Memproses...'; }
    if (progress) progress.classList.remove('qm-hidden');
    if (results) setInnerHTML(results, '');
    if (btnExport) btnExport.classList.add('qm-hidden');

    const poolSize = Math.min(APP_CONFIG.BATCH_POOL_SIZE, state.batchQueue.length);
    const workers = [];
    for (let i = 0; i < poolSize; i++) workers.push(processBatchWorker());
    Promise.all(workers).then(finishBatch);
    finishProfile(prof, { totalNrp: nrps.length, poolSize });
  }

  function onBatchCancel() {
    state.batchAborted = true;
    state.batchQueue = [];
    const btnCheck = document.getElementById('qm-btn-batch-check');
    if (btnCheck) btnCheck.textContent = 'Membatalkan...';
  }

  async function processBatchWorker() {
    while (state.batchQueue.length > 0 && !state.batchAborted) {
      const item = state.batchQueue.shift();
      const prof = startProfile('processBatchWorker:item', { nrp: item?.nrp });
      try {
        pushLog(`Memproses NRP ${item.nrp}...`);
        const emp = await getEmp(item.nrp);
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
            item.anomalies = scanRes.anomalies || [];
            item.rekaps = scanRes.rekaps || null;
            item.msg = item.anomalies.length + ' anomali ditemukan';
          } catch (e) {
            item.msg = 'Gagal ambil data kehadiran';
            item.anomalies = [];
          }
        }
      } catch (e) {
        item.found = false;
        item.msg = 'Gagal akses HRIS';
        item.anomalies = [];
        pushLog(`Gagal memproses NRP ${item.nrp}: ${e.message}`, 'error');
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

  function finishBatch() {
    const btnCheck = document.getElementById('qm-btn-batch-check');
    const btnExport = document.getElementById('qm-btn-export-batch');
    const progress = document.getElementById('qm-batch-progress');
    const statusBar = document.getElementById('qm-batch-status');
    const progressBar = document.getElementById('qm-batch-progress-bar');

    if (btnCheck) {
      btnCheck.textContent = 'Proses Batch';
      delete btnCheck.dataset.running;
    }
    if (progress) progress.classList.add('qm-hidden');
    if (btnExport && state.batchResults.length > 0) btnExport.classList.remove('qm-hidden');

    if (statusBar) statusBar.textContent = 'Selesai: ' + state.batchResults.length + '/' + state.batchTotal + ' NRP';
    if (progressBar) progressBar.style.width = '100%';

    if (state.batchAborted) {
      UI.showResult('warning', 'Dibatalkan', 'Proses pemeriksaan batch dihentikan oleh pengguna.');
      pushLog('Proses batch dibatalkan oleh pengguna.', 'error');
    } else {
      pushLog(`Batch check selesai. Total ${state.batchResults.length} NRP diproses.`);
    }

    if (state.debug && state.batchProfile) {
      const renderTotal = Number((state.batchProfile.renderTotal || 0).toFixed(2));
      Logger.debug('[PROFILE] batchSummary', {
        processed: state.batchResults.length,
        renderTotalMs: renderTotal,
        recentWorkerMedianMs: Number(getMedian((state.batchProfile.itemDurations || []).slice(-PROFILE_CONFIG.MEDIAN_SAMPLE_SIZE)).toFixed(2))
      });
      if (renderTotal >= PROFILE_CONFIG.HOT_BATCH_RENDER_TOTAL_MS) {
        Logger.warn(`Profiling: total render batch ${renderTotal.toFixed(2)}ms melewati ambang ${PROFILE_CONFIG.HOT_BATCH_RENDER_TOTAL_MS}ms.`);
      }
    }
  }

  function pushBatchResult(item) {
    state.batchResults.push(item);
    renderBatchResults();
    updateBatchProgress();
  }

  function updateBatchProgress() {
    const pct = state.batchTotal > 0 ? Math.round((state.batchResults.length / state.batchTotal) * 100) : 0;
    const barEl = document.getElementById('qm-batch-progress-bar');
    const statusEl = document.getElementById('qm-batch-status');
    if (barEl) barEl.style.width = pct + '%';
    if (statusEl) statusEl.textContent = 'Memproses... ' + state.batchResults.length + '/' + state.batchTotal;
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
    const container = document.getElementById('qm-batch-results');
    if (!container) {
      finishProfile(prof, { skipped: true });
      return;
    }
    if (state.batchResults.length === 0) {
      setInnerHTML(container, '');
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
      html += `<tr class="qm-batch-group-header qm-batch-header-bg" data-target=".${bagSafeId}">`;
      html += `<td colspan="5" class="qm-batch-bagian-cell"><div class="qm-flex qm-items-center qm-justify-between qm-w-full"><span>${escHtml(bag)}</span><span class="qm-chevron qm-accordion-chevron">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>
              </span></div></td></tr>`;

      html += `<tr class="qm-batch-group-row ${bagSafeId} qm-table-header qm-batch-sub-header-bg qm-hidden">`;
      html += '<td class="qm-batch-cell-header qm-batch-col-nrp">NRP</td>';
      html += '<td class="qm-batch-cell-header qm-batch-col-nama">Nama</td>';
      html += '<td class="qm-batch-cell-header qm-batch-col-ot">Lembur</td>';
      html += '<td class="qm-batch-cell-header qm-batch-col-ket">Keterangan</td>';
      html += '<td class="qm-batch-cell-header qm-batch-col-msg">Masalah / Anomali</td>';
      html += '</tr>';

      let sekIdx = 0;
      for (const sek in tree[bag]) {
        const sekSafeId = bagSafeId + '-sek-' + (sekIdx++);
        html += `<tr class="qm-batch-group-row ${bagSafeId} qm-batch-seksi-header qm-batch-sub-header-bg qm-hidden" data-target=".${sekSafeId}">`;
        html += `<td colspan="5" class="qm-batch-seksi-cell" style="padding-left: 32px;"><div class="qm-flex qm-items-center qm-justify-between qm-w-full"><span>${escHtml(sek)} <span class="qm-batch-seksi-count">(${tree[bag][sek].length})</span></span><span class="qm-chevron qm-accordion-chevron">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg>
              </span></div></td></tr>`;

        tree[bag][sek].forEach(item => {
          var nrpLink = item.nrp
            ? '<a href="#" class="qm-batch-nrp-link" data-nrp="' + escHtml(item.nrp) + '" style="padding-left: 64px;">' + escHtml(item.nrp) + '</a>'
            : '-';
          var nama = item.found ? escHtml(item.nama || '-') : '<span class="qm-batch-not-found">Tidak Ditemukan</span>';
          var masalahHtml = '';

          if (item.anomalies && item.anomalies.length > 0) {
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
                  const base = getSpklBaseUrl(item.nrp);
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

                fixBtn = `<button class="qm-fix-dot" title="${escHtml(titleStr)}" data-fix-link="${escHtml(finalLink)}" data-fix-date="${escHtml(tglPad)}" data-full-date="${escHtml(fDate)}"></button>`;
              }

              listItems += '<div class="qm-batch-date-row">';
              listItems += '<div class="qm-batch-date-header qm-flex qm-items-center qm-justify-between"><span style="position: relative; padding-right: 15px;"><b>Tgl ' + escHtml(tgl) + '</b>' + fixBtn + '</span><span class="qm-chevron qm-accordion-chevron"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"></polyline></svg></span></div>';
              listItems += '<div class="qm-batch-date-content qm-hidden">';
              // Deduplicate anomalies by message for the same date
              const msgMap = new Map();
              anomaliesTgl.forEach(a => {
                if (!msgMap.has(a.msg)) msgMap.set(a.msg, []);
                msgMap.get(a.msg).push(a.col);
              });

              msgMap.forEach((cols, msg) => {
                const uniqueCols = [...new Set(cols)].filter(Boolean);
                const colNames = uniqueCols.map(c => escHtml(c)).join(', ');
                const colPrefix = colNames ? colNames + ': ' : '';
                listItems += '<div class="qm-batch-anomaly-detail">• ' + colPrefix + escHtml(msg) + '</div>';
              });
              listItems += '</div></div>';
            }
            masalahHtml = '<div class="qm-batch-masalah-scroll">' + listItems + '</div>';
          } else if (item.found) {
            masalahHtml = '<span class="qm-batch-no-anomaly">Tidak ada anomali</span>';
          } else {
            masalahHtml = '<span class="qm-batch-not-found">' + escHtml(item.msg || 'Error') + '</span>';
          }

          const rk = item.rekaps || { otb: 0, otl: 0, ota: 0, otp: 0, keterangan: {} };
          const otValues = `B:${rk.otb.toFixed(1)} L:${rk.otl.toFixed(1)} A:${rk.ota.toFixed(1)} P:${rk.otp.toFixed(1)}`;

          const ketKeys = ['CT', 'CH', 'SD', 'I', 'IS', 'IA', 'A'];
          const ketStr = ketKeys.map(k => `${k}:${rk.keterangan[k] || 0}`).join(' | ');

          const lemburHtml = `<div class="qm-text-xs qm-font-mono" style="color: var(--qm-olive);">${otValues}</div>`;
          const ketHtml = `<div class="qm-text-xs qm-font-mono" style="color: var(--qm-stone); opacity: 0.8;">${ketStr}</div>`;

          html += `<tr class="qm-batch-group-row ${bagSafeId} ${sekSafeId} qm-batch-item-row qm-hidden">`;
          html += '<td class="qm-batch-cell">' + nrpLink + '</td>';
          html += '<td class="qm-batch-cell qm-batch-nama">' + nama + '</td>';
          html += '<td class="qm-batch-cell">' + lemburHtml + '</td>';
          html += '<td class="qm-batch-cell">' + ketHtml + '</td>';
          html += '<td class="qm-batch-cell">' + masalahHtml + '</td>';
          html += '</tr>';
        });
      }
    }

    html += '</tbody></table>';

    setInnerHTML(container, html);
    const duration = finishProfile(prof, { items: state.batchResults.length });
    if (state.batchProfile) state.batchProfile.renderTotal += duration;
  }

  function exportBatchResults() {
    if (state.batchResults.length === 0) { alert('Tidak ada hasil untuk diekspor.'); return; }
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
      alert('Library XLSX gagal dimuat.');
    }
  }

  function renderAnomalies() {
    const prof = startProfile('renderAnomalies', { count: state.anomalies.length });
    const badge = document.getElementById('qm-badge-anomali');
    const list = document.getElementById('qm-anomali-list');
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
          else if (itemWithLink.msg === 'Cek Distribusi' || itemWithLink.msg === 'Shift kosong' || itemWithLink.msg === 'Shift muncul') btnText = 'Fix Distribusi';
          else if (itemWithLink.msg === 'SPKL Disetujui (Cek Jenis OT)') btnText = 'Cek Halaman SPKL';
          else if (itemWithLink.msg === 'SPKL Belum Disetujui') btnText = 'Cek SPKL Online';
          else if (itemWithLink.msg === 'SPKL Ditolak') btnText = 'Cek SPKL (Ditolak)';
          else if (itemWithLink.msg === 'SPKL Online tidak ada') btnText = 'Input SPKL Online';
          else if (itemWithLink.msg === 'SPKL Online tidak sesuai') btnText = 'Cek SPKL Online';
        }

        const fixBtn = itemWithLink
          ? `<button class="qm-btn-fix-pill" data-fix-link="${escHtml(itemWithLink.link)}" data-fix-date="${escHtml(tgl)}" data-full-date="${escHtml(itemWithLink.fullDate || '')}" title="${escHtml(btnText)}">${btnText}</button>`
          : '';

        const detailsHtml = items.map(a => {
          const type = escHtml(COL_LABELS[a.colIndex] || a.col || ('Kolom ' + a.colIndex));
          return `
            <div class="qm-anomaly-card">
              <div class="qm-anomaly-card-type">${type}</div>
              <div class="qm-anomaly-card-msg">${escHtml(a.msg)}</div>
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
      setInnerHTML(list, html);
    } else {
      badge.classList.add('qm-hidden');
      badge.classList.remove('qm-visible-inline-flex');
      setInnerHTML(list, '<div class="qm-anomaly-empty-state qm-text-center qm-text-muted qm-mt-xl">Tidak ada anomali ditemukan.</div>');
    }
    finishProfile(prof, { count: state.anomalies.length });
  }

  /* ============================================================
   * 11. SPKL HIGHLIGHT
   * ============================================================ */
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
   * 12. NRP CHECK & AUTOFILL
   * ============================================================ */
  function checkNrp() {
    const inputNrp = document.getElementById('qm-input-nrp');
    const inputBulan = document.getElementById('qm-input-bulan');
    const nrp = inputNrp ? inputNrp.value.trim() : '';
    const bulan = inputBulan ? inputBulan.value.trim() : '';
    if (!nrp || !bulan) { UI.showResult('warning', 'Data Tidak Lengkap', 'Silakan masukkan NRP dan Bulan terlebih dahulu.'); return; }
    if (!/^\d+$/.test(nrp) || (nrp.length !== 4 && nrp.length !== 8)) { UI.showResult('warning', 'Format Tidak Valid', 'Hanya menerima 4 dan 8 angka NRP'); return; }
    UI.setLoading(true);
    sessionStorage.setItem(STORAGE.AUTO_NRP, nrp);
    sessionStorage.setItem(STORAGE.AUTO_BULAN, bulan);
    const year = new Date().getFullYear();
    window.location.href = getAttendanceUrl(bulan, year, nrp);
  }

  function autoFillTargetPage() {
    if (!isAttendancePagePath()) return;
    const autoNrp = sessionStorage.getItem(STORAGE.AUTO_NRP);
    const autoBulan = sessionStorage.getItem(STORAGE.AUTO_BULAN);
    if (autoNrp && autoBulan) {
      setTimeout(() => {
        setFieldValue(document.querySelector('#bulan'), autoBulan);
        setFieldValue(document.querySelector('input[name="nrp"]'), autoNrp, ['input']);

        sessionStorage.removeItem(STORAGE.AUTO_NRP);
        sessionStorage.removeItem(STORAGE.AUTO_BULAN);
      }, TIMING.AUTO_FILL_DELAY);
    }
  }

  /* ============================================================
   * 13. PANEL TOGGLE
   * ============================================================ */
  function openPanel() {
    state.isOpen = true;
    document.body.classList.add('qm-no-scroll');
    document.querySelectorAll('#qm-panel, #qm-fab, #qm-backdrop').forEach(el => el.classList.add('qm-open'));
    setTimeout(() => {
      const input = document.querySelector('#qm-input-nrp');
      if (input) input.focus();
    }, 250);
  }

  function closePanel() {
    state.isOpen = false;
    document.body.classList.remove('qm-no-scroll');
    document.querySelectorAll('#qm-panel, #qm-fab, #qm-backdrop').forEach(el => el.classList.remove('qm-open'));
  }

  function togglePanel() {
    state.isOpen ? closePanel() : openPanel();
  }

  /* ============================================================
   * 14. ANOMALY DETECTION
   * ============================================================ */

  function isShiftChecked(td) {
    if (!td) return false;
    const text = td.textContent.trim().toLowerCase();
    if (text !== '' && (text.includes('check') || text.includes('ok') || text.includes('✓') || text.includes('☑') || text.includes('v'))) return true;
    if (td.querySelector('input:checked')) return true;
    const html = td.innerHTML.toLowerCase();
    if (html.includes('check') || html.includes('fa-check') || html.includes('fa-square-check')) return true;
    return false;
  }

  function tebakShiftSebenarnya(waktuMsk, rules) {
    if (waktuMsk === null) return '1';
    const jamMentah = waktuMsk >= 24.0 ? waktuMsk - 24.0 : waktuMsk;
    if (jamMentah >= rules.shift1.jamTebakMulai && jamMentah <= rules.shift1.jamTebakAkhir) return '1';
    if (jamMentah > rules.shift2.jamTebakMulai && jamMentah <= rules.shift2.jamTebakAkhir) return '2';
    return '3';
  }

  /** Count rows marked as libur by CSS classes/colors. */
  function countLibur(docContext) {
    const root = docContext || document;
    const trs = root.querySelectorAll('table tbody tr');
    let total = 0;

    trs.forEach(tr => {
      const tds = tr.querySelectorAll('td');
      if (tds.length < 12 || tr.textContent.toLowerCase().includes('total')) return;
      const { isLiburColor } = parseRowCssFlags(tr);
      if (isLiburColor) total++;
    });
    return total;
  }

  /** Push an anomaly record. Pure — no DOM side effects. */
  function markAnomalyCell(anomalies, tglText, colIndex, title, customLink, cekSpklCells, fullDate, shiftVal) {
    anomalies.push({ tgl: tglText, fullDate, colIndex, msg: title, link: customLink || '', shift: shiftVal });
    if (title.includes('Cek SPKL') && cekSpklCells) {
      if (!cekSpklCells.some(c => c.tgl === tglText && c.colIndex === colIndex)) {
        cekSpklCells.push({ tgl: tglText, fullDate, colIndex, shift: shiftVal });
      }
    }
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
    const trs = root.querySelectorAll('table tbody tr');

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
          const linkStr = a.link ? `data-fix-link="${escHtml(a.link)}" data-fix-date="${escHtml(String(a.tgl))}" data-full-date="${escHtml(String(a.fullDate || ''))}"` : '';

          let titleStr = 'Buka Halaman Kehadiran';
          if (a.msg === 'Buka Halaman Kehadiran') titleStr = 'Buka Halaman Kehadiran';
          else if (a.msg && (a.msg.includes('Duplikasi') || a.msg.includes('Double entry'))) titleStr = 'Lihat Duplikasi Shift';
          else if (a.msg && a.msg.includes('Pulang awal')) titleStr = 'Cek Kehadiran (Pulang Awal)';
          else if (a.msg === 'Cek Distribusi' || a.msg === 'Shift kosong' || a.msg === 'Shift muncul') titleStr = 'Perbaiki Distribusi Jam Kerja';
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
          else if (a.msg === 'Cek Distribusi' || a.msg === 'Shift kosong' || a.msg === 'Shift muncul') titleStr = 'Perbaiki Distribusi Jam Kerja';
          else if (a.msg === 'Cek SPKL') titleStr = 'Cek SPKL Online';

          btn.setAttribute('title', titleStr);
        }
      });
    });
    finishProfile(prof, { count: anomalies.length });
  }

  function validateShiftRow(tds, tglText, mskText, klrText, rules, ctx, cekSpklCells, anomalies, isLibur, isHalfDay, fullDate) {
    const ketText = tds[COL.KET].textContent.trim().toUpperCase();
    let shift1 = isShiftChecked(tds[COL.SHIFT1]);
    let shift2 = isShiftChecked(tds[COL.SHIFT2]);
    let shift3 = isShiftChecked(tds[COL.SHIFT3]);

    // No shift checked but has clock data or is Mangkir → guess shift
    if (!shift1 && !shift2 && !shift3 && (mskText || klrText || ketText === 'A')) {
      const guessed = tebakShiftSebenarnya(parseTime(mskText), rules);
      const link = getDistribusiLink(ctx, tglText, guessed);
      const msg = 'Shift kosong';
      markAnomalyCell(anomalies, tglText, COL.SHIFT1, msg, link, cekSpklCells, fullDate);
      markAnomalyCell(anomalies, tglText, COL.SHIFT2, msg, link, cekSpklCells, fullDate);
      markAnomalyCell(anomalies, tglText, COL.SHIFT3, msg, link, cekSpklCells, fullDate);
    }

    // Derive active shift from clock-in if still unknown
    const mskTime = parseTime(mskText);
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
  function getHalfDayPulangAwalThreshold(shift1, shift2, shift3, mskTime) {
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
      ? getHalfDayPulangAwalThreshold(shift1, shift2, shift3, mskTime)
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
        if (adjMsk !== null && adjMsk <= 6.0) s1PulangAwal = 14.0;
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
        if (adjMsk > THRESHOLDS.SHIFT1_MSK_UPPER_BATAS) markAnomalyCell(anomalies, tglText, COL.SHIFT1, 'Cek Distribusi', getDistribusiLink(ctx, tglText, tebakShiftSebenarnya(adjMsk, rules)), cekSpklCells, fullDate);
        else if (adjMsk !== adjKlr) markAnomalyCell(anomalies, tglText, COL.MSK, 'Terlambat Shift I', '', cekSpklCells, fullDate);
      }
      const paShift1 = isHalfDay && paThreshold !== null ? paThreshold : s1PulangAwal;
      if (adjKlr !== null && adjKlr < paShift1 && adjMsk !== adjKlr) markAnomalyCell(anomalies, tglText, COL.KLR, 'Pulang awal Shift I', getKehadiranLink(ctx), cekSpklCells, fullDate, '1');
    } else if (shift2) {
      if (adjMsk !== null && adjMsk < rules.shift2.batasMasukLembur) mskLembur = true;
      if (adjKlr !== null && adjKlr > rules.shift2.batasKeluarLembur) klrLembur = true;
      if (adjMsk !== null) {
        if (adjMsk < THRESHOLDS.SHIFT2_MSK_LOWER_BATAS || adjMsk > THRESHOLDS.SHIFT2_MSK_UPPER_BATAS) markAnomalyCell(anomalies, tglText, COL.SHIFT2, 'Cek Distribusi', getDistribusiLink(ctx, tglText, tebakShiftSebenarnya(adjMsk, rules)), cekSpklCells, fullDate, '2');
        else if (adjMsk > rules.shift2.batasTerlambatMasuk && adjMsk !== adjKlr) markAnomalyCell(anomalies, tglText, COL.MSK, 'Terlambat Shift II', '', cekSpklCells, fullDate, '2');
      }
      const paShift2 = isHalfDay && paThreshold !== null ? paThreshold : rules.shift2.batasPulangAwal;
      if (adjKlr !== null && adjKlr < paShift2 && adjMsk !== adjKlr) markAnomalyCell(anomalies, tglText, COL.KLR, 'Pulang awal Shift II', getKehadiranLink(ctx), cekSpklCells, fullDate, '2');
    } else if (shift3) {
      if (adjMsk !== null && adjMsk > rules.shift3.batasMasukLemburAwal && adjMsk < rules.shift3.batasMasukLemburAkhir) mskLembur = true;
      if (adjKlr !== null && adjKlr > rules.shift3.batasKeluarLembur) klrLembur = true;
      if (adjMsk !== null) {
        if (adjMsk < rules.shift3.batasAwalMasuk || adjMsk > THRESHOLDS.SHIFT3_MSK_UPPER_BATAS) markAnomalyCell(anomalies, tglText, COL.SHIFT3, 'Cek Distribusi', getDistribusiLink(ctx, tglText, tebakShiftSebenarnya(adjMsk, rules)), cekSpklCells, fullDate, '3');
        else if (adjMsk > rules.shift3.batasTerlambatMasuk && adjMsk !== adjKlr) markAnomalyCell(anomalies, tglText, COL.MSK, 'Terlambat Shift III', '', cekSpklCells, fullDate, '3');
      }
      const paShift3 = isHalfDay && paThreshold !== null ? paThreshold : rules.shift3.batasPulangAwal;
      if (adjKlr !== null && adjKlr < paShift3 && adjMsk !== adjKlr) markAnomalyCell(anomalies, tglText, COL.KLR, 'Pulang awal Shift III', getKehadiranLink(ctx), cekSpklCells, fullDate, '3');
    }
    return { mskLembur, klrLembur };
  }

  /** Unified row status detector (Libur, HalfDay, Normal). */
  function getRowStatus(tr) {
    const tds = tr.querySelectorAll('td');
    if (tds.length === 0) return { isLibur: false, isHalfDay: false, ketText: '' };

    const { isLiburColor, isHalfDayColor } = parseRowCssFlags(tr);

    let ketText = '';
    if (tds.length > COL.KET) ketText = tds[COL.KET].textContent.trim().toUpperCase();

    const isLibur = isLiburColor || ['L', 'LB', 'LH'].includes(ketText);
    const isHalfDay = !isLibur && (isHalfDayColor || ['S', 'CH', 'HD'].includes(ketText));

    return { isLibur, isHalfDay, ketText };
  }

  function scanAttendanceTable(doc, ctx) {
    const prof = startProfile('scanAttendanceTable', { nrp: ctx?.nrp, bulan: ctx?.bulan, tahun: ctx?.tahun });
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

    const totalLibur = countLibur(doc);
    const is5HariKerja = totalLibur >= THRESHOLDS.MIN_LIBUR_5_HARI_KERJA;
    const rules = structuredClone(SHIFT_RULES);
    if (is5HariKerja) {
      rules.shift2.batasKeluarLembur = THRESHOLDS.SHIFT2_KLR_LEMBUR_5HR;
    }

    const root = doc || document;
    const trs = root.querySelectorAll('table tbody tr');

    // Pre-calculate shift counts per date to detect duplicates
    const dateShiftCounts = {};
    trs.forEach(tr => {
      const tds = tr.querySelectorAll('td');
      if (tds.length < 12 || tr.textContent.toLowerCase().includes('total')) return;
      const tgl = tds[COL.TGL].textContent.trim();
      const hasChecked = isShiftChecked(tds[COL.SHIFT1]) || isShiftChecked(tds[COL.SHIFT2]) || isShiftChecked(tds[COL.SHIFT3]);
      if (tgl && hasChecked) dateShiftCounts[tgl] = (dateShiftCounts[tgl] || 0) + 1;
    });

    trs.forEach(tr => {
      const tds = tr.querySelectorAll('td');
      if (tds.length < 12 || tr.textContent.toLowerCase().includes('total')) return;

      const tglText = tds[COL.TGL].textContent.trim();
      const fullDate = `${ctx.tahun}-${String(ctx.bulan).padStart(2, '0')}-${String(tglText).padStart(2, '0')}`;

      const mskText = tds[COL.MSK].textContent.trim();
      const klrText = tds[COL.KLR].textContent.trim();
      const { isLibur, isHalfDay, ketText } = getRowStatus(tr);

      const shiftInfo = validateShiftRow(tds, tglText, mskText, klrText, rules, ctx, cekSpklCells, anomalies, isLibur, isHalfDay, fullDate);
      const { shift1, shift2, shift3, activeShift, mskTime } = shiftInfo;
      const klrTime = parseTime(klrText);

      // Detect multiple rows for same date with checked shifts (Barcode overlap/error)
      const isSaturday = new Date(ctx.tahun, ctx.bulan - 1, parseInt(tglText)).getDay() === 6;
      if (activeShift && dateShiftCounts[tglText] > 1 && !isHalfDay && !isSaturday) {
        const barcodeLink = getKehadiranLink(ctx);
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

        markAnomalyCell(anomalies, tglText, COL.TGL, msg, barcodeLink, cekSpklCells, fullDate);
        if (shift1) markAnomalyCell(anomalies, tglText, COL.SHIFT1, msg, barcodeLink, cekSpklCells, fullDate);
        if (shift2) markAnomalyCell(anomalies, tglText, COL.SHIFT2, msg, barcodeLink, cekSpklCells, fullDate);
        if (shift3) markAnomalyCell(anomalies, tglText, COL.SHIFT3, msg, barcodeLink, cekSpklCells, fullDate);
      }

      if (activeShift && mskTime !== null && !isLibur && !isHalfDay) {
        const guessed = tebakShiftSebenarnya(mskTime, rules);
        if (guessed !== activeShift) {
          const msg = 'Jam MSK tidak cocok dengan Shift ' + activeShift + ' (Terdeteksi Shift ' + guessed + ')';
          const link = getDistribusiLink(ctx, tglText, guessed);
          markAnomalyCell(anomalies, tglText, COL.MSK, msg, link, cekSpklCells, fullDate);
          if (shift1) markAnomalyCell(anomalies, tglText, COL.SHIFT1, msg, '', cekSpklCells, fullDate);
          if (shift2) markAnomalyCell(anomalies, tglText, COL.SHIFT2, msg, '', cekSpklCells, fullDate);
          if (shift3) markAnomalyCell(anomalies, tglText, COL.SHIFT3, msg, '', cekSpklCells, fullDate);
        }
      }

      let isAbsent = ketText === 'A';
      for (let i = 11; i < tds.length; i++) {
        if (tds[i].textContent.trim() === 'A') isAbsent = true;
      }
      if (isAbsent) {
        absentDates.push({ date: tglText, tr: tr });
      }

      const barcodeLink = getKehadiranLink(ctx);

      if (!isLibur) {
        if (!mskText && !ketText) markAnomalyCell(anomalies, tglText, COL.MSK, 'Buka Halaman Kehadiran', barcodeLink, cekSpklCells, fullDate);
        if (!klrText && !ketText) markAnomalyCell(anomalies, tglText, COL.KLR, 'Buka Halaman Kehadiran', barcodeLink, cekSpklCells, fullDate);
      } else {
        if (mskText || klrText) {
          if (!mskText) markAnomalyCell(anomalies, tglText, COL.MSK, 'Buka Halaman Kehadiran', barcodeLink, cekSpklCells, fullDate);
          if (!klrText) markAnomalyCell(anomalies, tglText, COL.KLR, 'Buka Halaman Kehadiran', barcodeLink, cekSpklCells, fullDate);
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

      if (valOtb > THRESHOLDS.OT_BATAS_WAJAR) markAnomalyCell(anomalies, tglText, COL.OTB, 'Angka OTB tidak wajar', '', cekSpklCells, fullDate);
      if (valOtl > THRESHOLDS.OT_BATAS_WAJAR) markAnomalyCell(anomalies, tglText, COL.OTL, 'Angka OTL tidak wajar', '', cekSpklCells, fullDate);
      if (valOtp > THRESHOLDS.OT_BATAS_WAJAR) markAnomalyCell(anomalies, tglText, COL.OTP, 'Angka OTP tidak wajar', '', cekSpklCells, fullDate);

      const hasAnyOT = parseFloat(otbText) > 0 || parseFloat(otlText) > 0 || parseFloat(otpText) > 0;
      if ((ot.mskLembur || ot.klrLembur) && !hasAnyOT) {
        const d = String(tglText).padStart(2, '0');
        const spklUrl = ROUTES.SPKL_ONLINE(ctx.tahun, String(ctx.bulan).padStart(2, '0'), d, d, ctx.nrp);
        const sVal = activeShift || '';
        markAnomalyCell(anomalies, tglText, COL.OTB, 'Cek SPKL', spklUrl, cekSpklCells, fullDate, sVal);
        markAnomalyCell(anomalies, tglText, COL.OTL, 'Cek SPKL', spklUrl, cekSpklCells, fullDate, sVal);
        markAnomalyCell(anomalies, tglText, COL.OTP, 'Cek SPKL', spklUrl, cekSpklCells, fullDate, sVal);
        if (ot.mskLembur) markAnomalyCell(anomalies, tglText, COL.MSK, 'Cek SPKL', spklUrl, cekSpklCells, fullDate, sVal);
        if (ot.klrLembur) markAnomalyCell(anomalies, tglText, COL.KLR, 'Cek SPKL', spklUrl, cekSpklCells, fullDate, sVal);
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

    try {
      state.anomalies = [];
      state.pendingChecks = 0;

      const anomaliTab = document.querySelector('[data-pane="anomali"]');
      if (anomaliTab) anomaliTab.classList.remove('qm-tab-loading');

      const ctx = getPageContext();
      const result = scanAttendanceTable(document, ctx);
      state.anomalies = result.anomalies;

      applyMark(document, result.anomalies);
      renderAnomalies();

      if (result.absentDates.length > 0) { state.pendingChecks++; checkBarcodeMangkir(result.absentDates); }
      if (result.cekSpklCells.length > 0) { state.pendingChecks++; checkSPKLOnline(result.cekSpklCells); }
      if (state.pendingChecks > 0 && anomaliTab) anomaliTab.classList.add('qm-tab-loading');
    } finally {
      finishProfile(prof, { anomalyCount: state.anomalies.length });
    }
  }

  /* ============================================================
   * 15. SPKL ONLINE CHECK
   * ============================================================ */
  async function checkSPKLOnline(cells) {
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
      const data = await req(spklUrl);
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
          const base = getSpklBaseUrl(ctx.nrp);
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
      renderAnomalies();
    } catch (e) {
      Logger.warn('Gagal mengambil data SPKL Online.', e);
    } finally {
      finishProfile(prof, { cellCount: cells.length, anomalyCount: state.anomalies.length });
      decrementPendingChecks();
    }
  }

  /* ============================================================
   * 16. BARCODE MANGKIR CHECK
   * ============================================================ */
  async function checkBarcodeMangkir(absentDates) {
    const prof = startProfile('checkBarcodeMangkir', { absentCount: absentDates.length });
    const ctx = getPageContext();
    if (!ctx.nrp) {
      finishProfile(prof, { skipped: true });
      return;
    }

    const barcodeUrl = getKehadiranLink(ctx);

    try {
      const data = await req(barcodeUrl);
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
      applyMark(document, state.anomalies);
      renderAnomalies();
    } catch (e) {
      Logger.warn('Gagal mengambil data barcode.');
    } finally {
      finishProfile(prof, { absentCount: absentDates.length, anomalyCount: state.anomalies.length });
      decrementPendingChecks();
    }
  }

  /* ============================================================
   * 17. AUTO FIX SPKL TYPE
   * ============================================================ */
  async function processSpklBackgroundSingle(item) {
    Logger.info(`Fetching ${item.link}`);
    const html = await req(item.link);
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
    const mskTime = parseTime(mskValue);
    const urlParams = new URLSearchParams(item.link.split('?')[1] || '');
    const shift = urlParams.get('shift') || tebakShiftSebenarnya(mskTime, SHIFT_RULES);
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
      headers['X-CSRF-TOKEN'] = token;
      headers['X-XSRF-TOKEN'] = token;
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
      if (!document.getElementById('qm-global-loader')) {
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

  /* ============================================================
   * 18. AUTO DISTRIBUSI JAM KERJA
   * ============================================================ */

  /** Select dropdown value with MutationObserver support. */
  /** Select dropdown value with robust option polling (waits for AJAX options to load). */
  async function pilihDropdownDinamis(selector, nilaiTarget, callback, timeout = 5000) {
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

  // fetchEmployeeData — removed, replaced by getEmp()

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

  /** Strips unsafe tags and attributes from HTML strings. */
  /** Fill the distribusi form and submit. */
  function fillDistribusiForm(dataKaryawan, nrp, tanggal, shift) {
    setTimeout(() => {
      UI.setGlobalProgress(20, 'Mengisi Jam Kerja...');

      const jkSelect = document.querySelector('select[name*="jam_kerja"], select[name*="jk"]');
      if (jkSelect && dataKaryawan.jk) {
        selectOption(jkSelect, opt => {
          const optVal = opt.value.trim();
          const textKode = opt.textContent.trim().split('-')[0].trim();
          return optVal === dataKaryawan.jk || textKode === dataKaryawan.jk;
        });
        selectPickerRefresh(jkSelect);
      }

      UI.setGlobalProgress(35, 'Mengisi Periode...');
      const dateAwal = document.querySelectorAll('input[name*="tanggal_awal"], input[name*="periode_awal"], input[name*="tgl_awal"], input[name*="tgl_dari"], input[id*="tanggal_awal"], input[id*="tgl_awal"], input[name="start_date"]');
      const dateAkhir = document.querySelectorAll('input[name*="tanggal_akhir"], input[name*="periode_akhir"], input[name*="tgl_akhir"], input[name*="tgl_sampai"], input[name*="tgl_ke"], input[id*="tanggal_akhir"], input[id*="tgl_akhir"], input[name="end_date"]');

      const tglAwal = Array.isArray(tanggal) ? tanggal[0] : tanggal;
      const tglAkhir = Array.isArray(tanggal) ? (tanggal[1] || tglAwal) : tglAwal;

      dateAwal.forEach(input => setFieldValue(input, tglAwal));
      dateAkhir.forEach(input => setFieldValue(input, tglAkhir));

      UI.setGlobalProgress(50, 'Menyesuaikan Bagian...');
      pilihDropdownDinamis('select[name*="bagian"]', dataKaryawan.bag, () => {
        UI.setGlobalProgress(65, 'Menyesuaikan Seksi...');
        pilihDropdownDinamis('select[name*="seksi"]', dataKaryawan.sek, () => {
          UI.setGlobalProgress(75, 'Menyesuaikan Group & NRP...');
          const grpSelect = document.querySelector('select[name="kode_group"]');
          if (grpSelect && dataKaryawan.grp) {
            setFieldValue(grpSelect, dataKaryawan.grp);
            selectPickerRefresh(grpSelect);
          }

          if (typeof nrp === 'object' && nrp.awal !== undefined) {
            const nrpIn1 = document.querySelector('input[list="nrp_awal"], input[name*="nrp_awal"], input[id*="nrp_initial"], input[name*="nrp_initial"], input[name*="nrp1"], input[name*="nrp_initial_text"]');
            const nrpIn2 = document.querySelector('input[list="nrp_akhir"], input[name*="nrp_akhir"], input[id*="nrp_final"], input[name*="nrp_final"], input[name*="nrp2"], input[name*="nrp_final_text"]');
            if (nrpIn1) setFieldValue(nrpIn1, nrp.awal, ['input', 'change', 'blur']);
            if (nrpIn2) setFieldValue(nrpIn2, nrp.akhir || nrp.awal, ['input', 'change', 'blur']);
          } else {
            let nrpInputs = document.querySelectorAll('input[list="nrp_awal"], input[list="nrp_akhir"], input[name*="nrp_awal"], input[name*="nrp_akhir"], input[name*="nrp1"], input[name*="nrp2"], input[name*="nrp_1"], input[name*="nrp_2"]');
            if (nrpInputs.length === 0) nrpInputs = document.querySelectorAll('input[name*="nrp"]');
            nrpInputs.forEach(input => setFieldValue(input, nrp, ['input', 'change', 'blur']));
          }

          if (shift) {
            const targetShiftRoman = shift === '1' ? 'I' : (shift === '2' ? 'II' : 'III');
            const expectedText = `${targetShiftRoman} - SHIFT ${targetShiftRoman}`;
            const shiftSelect = document.querySelector('select[name="kode_shift"], select[name="shift"]');
            if (shiftSelect) {
              selectOption(shiftSelect, opt => {
                const optText = opt.textContent.trim().toUpperCase();
                return opt.value === shift || optText === expectedText || optText === targetShiftRoman || optText === 'SHIFT ' + targetShiftRoman;
              });
              selectPickerRefresh(shiftSelect);
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
            sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
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
      // Check if we are on the result page of a distribution (no qm_auto param but AUTO_FINISHED is true)
      if (sessionStorage.getItem(STORAGE.AUTO_FINISHED) === 'true') {
        const pageText = document.body.textContent;
        const successAlert = document.querySelector('.alert-success, .alert-info');

        if (successAlert || pageText.includes('Distribution Process Completed')) {
          UI.showResult('success', 'Distribusi Selesai', 'Distribution Process Completed');
          const returnUrl = sessionStorage.getItem(STORAGE.RETURN_URL);
          if (returnUrl) {
            setTimeout(() => {
              sessionStorage.removeItem(STORAGE.RETURN_URL);
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
      const emp = await getEmp(nrp);
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

  /* ============================================================
   * 19. BATCH SPKL INPUT
   * ============================================================ */
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

    sessionStorage.setItem(STORAGE.RETURN_URL, window.location.href);

    if (!/^\d{4}$|^\d{8}$/.test(nrp)) {
      UI.showResult('warning', 'NRP Tidak Valid', 'Gunakan 4 digit (Reguler) atau 8 digit (OS).');
      return;
    }

    let targetUrl = getSpklCreateUrl(nrp);

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

      setFieldValue(nrpInput, nrp, ['input', 'change']);
      setFieldValue(tanggalInput, fullDate, ['input', 'change']);
      if (jenisOtSelect) {
        setFieldValue(jenisOtSelect, task.jenisOt);
        if (window.jQuery && window.jQuery(jenisOtSelect).selectpicker) {
          window.jQuery(jenisOtSelect).selectpicker('refresh');
        }
      }

      if (task.jenisOt === "7") {
        setFieldValue(jamAwalEl, jamAwal);
        setFieldValue(jamAkhirEl, jamAkhir);
        setFieldValue(shiftEl, shiftVal);
      }

      await new Promise(r => setTimeout(r, TIMING.SPKL_INPUT_DELAY));
      if (btnTambah) btnTambah.click();
      await new Promise(r => setTimeout(r, TIMING.SPKL_CLICK_DELAY));
    }

    UI.setGlobalProgress(95, 'Menyimpan...');
    const btnSubmit = document.getElementById("submit");
    if (btnSubmit) {
      sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
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
        const current = getSpklAddPageKind();
        if (current && s.indexes[current] < s[current].length) return current;
        if (s.indexes.internal < s.internal.length) return "internal";
        if (s.indexes.outsource < s.outsource.length) return "outsource";
        return null;
      })(st);

      if (pRoute) {
        const current = getSpklAddPageKind();
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

  /* ============================================================
   * 20. BATCH BANYAK NRP
   * ============================================================ */
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

    sessionStorage.setItem(STORAGE.RETURN_URL, window.location.href);

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

    const current = getSpklAddPageKind();
    const target = (int.length > 0 ? "internal" : "outsource");

    if (current !== target) {
      UI.showResult('success', 'Mengalihkan...', 'Pindah ke halaman input ' + target + '.');
      setTimeout(() => { window.location.href = getSpklAddUrlByKind(target); }, 1000);
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

      const setVal = (el, val) => setFieldValue(el, val);

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
      sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
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

    sessionStorage.setItem(STORAGE.RETURN_URL, window.location.href);

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

    const current = getAbsenCreatePageKind();
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

      setFieldValue(tEl, s.date, ['change']);
      setFieldValue(nEl, nrp, ['input', 'change']);
      setFieldValue(jEl, s.jam, ['change']);
      setFieldValue(sEl, s.status, ['change']);

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
      sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
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
      const current = getAbsenCreatePageKind();

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

  /* ============================================================
   * 21. AUTOMASI HADIR BULANAN
   * ============================================================ */
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
    sessionStorage.setItem(STORAGE.RETURN_URL, window.location.href);

    const targetURL = getAbsenAddUrl(NRP);

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

      setFieldValue(inputNrp, NRP, ['input', 'change']);
      setFieldValue(inputTanggal, aksi.waktu, ['input', 'change']);
      setFieldValue(inputStatus, aksi.status, ['change']);

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
      sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
      setTimeout(() => btnSubmit.click(), 500);
    } else {
      UI.hideGlobalLoader(1500);
      alert(`[Quick Menu] Automasi Selesai!\n- NRP: ${NRP}\n- Total dieksekusi: ${antrean.length} baris.\nSilakan klik Simpan secara manual.`);
    }
  }

  /* ============================================================
   * 22. EVENT HANDLERS
   * ============================================================ */
  function onExportAnomali() {
    if (state.anomalies.length === 0) { alert('Tidak ada anomali untuk diekspor.'); return; }
    const wsData = [['Tanggal', 'Kolom', 'Pesan Anomali']];
    const sorted = [...state.anomalies].sort((a, b) => parseInt(a.tgl) - parseInt(b.tgl));
    sorted.forEach(a => wsData.push([a.tgl, a.col, a.msg]));
    if (typeof XLSX !== 'undefined') {
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Anomali');
      const inputBulan = document.getElementById('qm-input-bulan');
      const month = (inputBulan && inputBulan.value) ? inputBulan.value : (new Date().getMonth() + 1);
      XLSX.writeFile(wb, `Anomali_Bulan_${month}.xlsx`);
    } else {
      alert('Library XLSX gagal dimuat. Harap periksa koneksi atau header script.');
    }
  }

  function onToggleAnomalyGroup(e) {
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

  function onCekSpklOnline() {
    const nrp = document.getElementById('qm-spkl-online-nrp')?.value.trim();
    const dateInput = document.getElementById('qm-spkl-online-date')?.value;
    if (!nrp || !dateInput) { alert('Harap isi NRP dan Tanggal.'); return; }

    const d = new Date(dateInput);
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();

    const url = ROUTES.SPKL_ONLINE(year, month, day, day, nrp);
    window.open(url, '_blank');
  }

  function onProsesInputHadir() {
    const nrp = document.getElementById('qm-input-hadir-nrp')?.value.trim();
    const tgl = document.getElementById('qm-input-hadir-tanggal')?.value;
    const jam = document.getElementById('qm-input-hadir-jam')?.value;
    const status = document.getElementById('qm-input-hadir-status')?.value;

    if (!nrp || !tgl || !jam || status === "") {
      alert('Harap isi semua field (NRP, Tanggal, Jam, dan Status).');
      return;
    }

    sessionStorage.setItem(STORAGE.RETURN_URL, window.location.href);

    const data = { nrp, tgl, jam, status };
    sessionStorage.setItem(STORAGE.INPUT_HADIR, JSON.stringify(data));

    const targetUrl = getAbsenCreateUrl(nrp);
    window.open(targetUrl, '_blank');
  }

  function onProsesDistribusi() {
    onSaveJkChange();
  }

  function onKeydownAnomalyGroup(e) {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.click(); }
  }

  function onFixDotClick(e) {
    e.stopPropagation();
    handleFixClick(
      this.getAttribute('data-fix-link'),
      this.getAttribute('data-fix-date'),
      this.getAttribute('title'),
      this.getAttribute('data-full-date')
    );
  }

  function onTabClick() {
    const pane = this.getAttribute('data-pane');
    document.querySelectorAll('.qm-tab').forEach(el => el.classList.remove('active'));
    this.classList.add('active');
    document.querySelectorAll('.qm-pane').forEach(el => el.classList.remove('active'));
    const targetPane = document.getElementById(`qm-pane-${pane}`);
    if (targetPane) targetPane.classList.add('active');
    localStorage.setItem('qm_last_tab', pane);

    if (pane === 'distribusi' || pane === 'check-nrp' || pane === 'spkl' || pane === 'kehadiran') {
      refreshGlobalData();
    }
  }

  function onInputBulan() {
    const val = parseInt(this.value);
    if (isNaN(val)) return;
    if (val < 1) this.value = 1;
    else if (val > 12) this.value = 12;
  }

  function onDocumentClick(e) {
    if (state.isOpen && !e.target.closest('#qm-panel, #qm-fab, #qm-backdrop')) closePanel();
  }

  function onRecordShortcut() {
    isRecordingShortcut = true;
    this.textContent = 'Tunggu...';
    document.getElementById('qm-input-shortcut').value = 'Tekan tombol...';
  }

  function onBatchNrpClick(e) {
    e.preventDefault();
    var nrp = this.getAttribute('data-nrp');
    if (!nrp) return;
    var url = getAttendanceUrl(state.batchBulan, state.batchTahun, nrp);
    window.open(url, '_blank');
  }

  function onBatchFixClick(e) {
    e.stopPropagation();
    handleFixClick(
      this.getAttribute('data-fix-link'),
      this.getAttribute('data-fix-date'),
      this.getAttribute('title'),
      this.getAttribute('data-full-date')
    );
  }

  function onKeydownDocument(e) {
    if (isRecordingShortcut) {
      if (e.key === 'Escape') {
        e.preventDefault();
        isRecordingShortcut = false;
        const btn = document.getElementById('qm-btn-record-shortcut');
        if (btn) btn.textContent = 'Ubah';
        const input = document.getElementById('qm-input-shortcut');
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
      const input = document.getElementById('qm-input-shortcut');
      if (input) input.value = shortcutKey;
      const btn = document.getElementById('qm-btn-record-shortcut');
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

  function onToggleDebugMode() {
    state.debug = this.checked;
    GM_setValue('qm_debug', state.debug);
    Logger.info(`Debug mode ${state.debug ? 'diaktifkan' : 'dimatikan'}`);
  }

  function onShowLogs() {
    const container = document.getElementById('qm-log-container');
    const btn = document.getElementById('qm-btn-show-logs');
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

  function onClearLogs() {
    if (confirm('Bersihkan semua riwayat log aktivitas?')) {
      state.batchLogs = [];
      renderLogs();
    }
  }

  function onExportLogs() {
    if (state.batchLogs.length === 0) {
      alert('Tidak ada log untuk diekspor.');
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

  function renderLogs() {
    const logBody = document.getElementById('qm-log-body');
    if (!logBody) return;

    if (state.batchLogs.length === 0) {
      setInnerHTML(logBody, '<div class="qm-text-muted qm-text-center qm-mt-xl">Belum ada log aktivitas.</div>');
      return;
    }

    setInnerHTML(logBody, state.batchLogs.map(log => `
      <div class="qm-log-item">
        <span class="qm-log-time">[${log.time}]</span>
        <span class="qm-log-msg ${log.level || log.type || 'info'}">${escHtml(log.msg)}</span>
      </div>
    `).join(''));
    logBody.scrollTop = logBody.scrollHeight;
  }

  let manualSidebarOverride = false;
  function enforceSidebar() {
    if (!alwaysCollapseMenu || manualSidebarOverride) return;
    if (!document.body.classList.contains('enlarged')) {
      document.body.classList.add('enlarged');
      document.body.classList.remove('sidebar-enable');
    }
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
                setFieldValue(nrpInput, nrpFill, ['input', 'change']);
              }
              sessionStorage.removeItem(STORAGE.AUTO_NRP_FILL);
            }
            if (dateFill) {
              const dateInput = document.getElementById('tanggal');
              if (dateInput) {
                const ctx = getPageContext();
                const fullDate = `${ctx.tahun}-${String(ctx.bulan).padStart(2, '0')}-${String(dateFill).padStart(2, '0')}`;
                setFieldValue(dateInput, fullDate, ['change']);
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

        if (elTgl) setFieldValue(elTgl, data.tgl, ['change']);
        if (elNrp) setFieldValue(elNrp, data.nrp, ['input', 'change']);
        if (elJam) setFieldValue(elJam, data.jam, ['change']);
        if (elStatus) setFieldValue(elStatus, data.status, ['change']);

        await new Promise(r => setTimeout(r, 600));
        if (btnTambah) btnTambah.click();

        await new Promise(r => setTimeout(r, 800));
        if (btnSubmit) {
          sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
          btnSubmit.click();
        }

        UI.hideGlobalLoader();
      } catch (e) {
        Logger.error('Auto Input Kehadiran Error', e);
        UI.hideGlobalLoader();
      }
    }, 1000);
  }

  /* ============================================================
   * 23. INIT
   * ============================================================ */
  function init() {

    spklHighlight();
    autoFillTargetPage();
    autoClickAddData();
    autoInputHadir();
    autoDistribusi();
    autoDistribusiSubsi();
    autoDistKK();
    checkSpklBatchResume();
    checkHadirBatchResume();
    checkHadirBulanResume();
    initJkChangeEvents();
    checkJkRestoration();



    // Return to source page after process finish (autoDistribusi result page)
    const returnUrl = sessionStorage.getItem(STORAGE.RETURN_URL);
    const isFinished = sessionStorage.getItem(STORAGE.AUTO_FINISHED) === 'true';
    const hasRestorationPending = sessionStorage.getItem('qm_jk_to_restore_' + getPageContext().nrp);

    if (isFinished && returnUrl) {
      const currentUrl = window.location.href.split('?')[0];
      const targetUrlBase = returnUrl.split('?')[0];
      const isReturnPage = currentUrl === targetUrlBase || window.location.href.includes(returnUrl);

      if (isReturnPage) {
        // We are back at the start page. Cleanup if no restoration is pending.
        if (!hasRestorationPending) {
          sessionStorage.removeItem(STORAGE.AUTO_FINISHED);
          sessionStorage.removeItem(STORAGE.RETURN_URL);
          UI.showResult('success', 'Selesai', 'Tugas latar belakang telah diselesaikan.');

          // If on attendance table, trigger anomaly refresh to show the fix
          if (isAttendancePagePath()) {
            setTimeout(() => {
              const btn = document.querySelector('[data-pane="anomali"]');
              if (btn) btn.click();
            }, 500);
          }
        }
      } else if (!window.location.search.includes('qm_auto')) {
        // We are on a result page, redirect back
        UI.showGlobalLoader('Selesai', 'Kembali ke halaman awal...');
        setTimeout(() => {
          window.location.href = returnUrl;
        }, 1500);
      }
    }

    if (!document.getElementById('qm-fab')) {
      document.body.insertAdjacentHTML('beforeend', HTML);

      // Default NRP for FIX Panel
      const ctx = getPageContext();
      const elFixNrp = document.getElementById('qm-fix-spkl-nrp');
      if (elFixNrp && ctx.nrp) elFixNrp.value = ctx.nrp;

      const elBulanNrp = document.getElementById('qm-input-hadir-bulan-nrp');
      if (elBulanNrp && ctx.nrp) elBulanNrp.value = ctx.nrp;

      const elDistNrp = document.getElementById('qm-input-distribusi-nrp');
      if (elDistNrp && ctx.nrp) elDistNrp.value = ctx.nrp;

      const elDistKKNrp = document.getElementById('qm-dist-KK-nrp');
      if (elDistKKNrp && ctx.nrp) elDistKKNrp.value = ctx.nrp;

      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();

      // Populate Month Select
      const monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
      const monthOptions = monthNames.map((name, i) => `<option value="${i + 1}" ${i + 1 === currentMonth ? 'selected' : ''}>${name}</option>`).join('');

      const inputBulan = document.getElementById('qm-input-bulan');
      if (inputBulan) setInnerHTML(inputBulan, monthOptions);

      const fixBulan = document.getElementById('qm-fix-spkl-bulan');
      if (fixBulan) setInnerHTML(fixBulan, monthOptions);

      const bulanBulan = document.getElementById('qm-input-hadir-bulan-bln');
      if (bulanBulan) setInnerHTML(bulanBulan, monthOptions);

      // Populate Year Select
      let years = '';
      for (let y = currentYear - 2; y <= currentYear + 1; y++) {
        years += `<option value="${y}" ${y === currentYear ? 'selected' : ''}>${y}</option>`;
      }

      const inputTahun = document.getElementById('qm-input-tahun');
      if (inputTahun) setInnerHTML(inputTahun, years);

      const fixTahun = document.getElementById('qm-fix-spkl-tahun');
      if (fixTahun) setInnerHTML(fixTahun, years);

      // Default date for Banyak NRP
      const elManyDate = document.getElementById('qm-fix-many-date');
      if (elManyDate) {
        elManyDate.value = `${currentYear}-${String(currentMonth).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
      }

      const elDistKKDate = document.getElementById('qm-dist-KK-date');
      if (elDistKKDate) {
        elDistKKDate.value = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;
      }

      const elDistDate = document.getElementById('qm-input-distribusi-tanggal');
      if (elDistDate) {
        elDistDate.value = `${currentYear}-${String(currentMonth).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
      }

      // Apply Theme
      UI.applyTheme(state.theme);

      // Restore Last Tab
      const lastTab = localStorage.getItem('qm_last_tab');
      if (lastTab) {
        const tabBtn = document.querySelector(`.qm-tab[data-pane="${lastTab}"]`);
        if (tabBtn) tabBtn.click();
      }
      initDraggable();
      delegate('change', '#qm-config-debug-mode', onToggleDebugMode);
      delegate('click', '#qm-btn-show-logs', onShowLogs);
      delegate('click', '#qm-btn-clear-logs', onClearLogs);
      delegate('click', '#qm-btn-export-logs', onExportLogs);
    }
    detectAnomalies();

    document.addEventListener('keydown', onKeydownDocument);
    document.addEventListener('click', onDocumentClick);

    delegate('click', '#qm-fab', togglePanel);
    delegate('click', '#qm-backdrop', closePanel);
    delegate('click', '#qm-btn-close-header', closePanel);
    delegate('click', '#qm-btn-check', checkNrp);
    delegate('click', '#qm-btn-spkl-batch', runSpklBatchProcess);
    delegate('click', '#qm-btn-spkl-many-nrp', runSpklManyNrpBatch);
    delegate('click', '#qm-btn-spkl-online-cek', onCekSpklOnline);
    delegate('click', '#qm-btn-hadir-proses', onProsesInputHadir);
    delegate('click', '#qm-btn-hadir-bulan-proses', runHadirBulanBatch);
    delegate('click', '#qm-btn-hadir-many-proses', runHadirManyNrpBatch);
    delegate('click', '#qm-btn-distribusi-proses', onProsesDistribusi);
    delegate('click', '#qm-btn-distribusi-subsi-proses', onProsesDistribusiSubsi);
    delegate('click', '#qm-global-cancel-btn', function () {
      Logger.info('User cancelled automation');
      sessionStorage.removeItem(STORAGE.SPKL_QUEUE);
      sessionStorage.removeItem(STORAGE.SPKL_CURRENT_INDEX);
      sessionStorage.removeItem(STORAGE.SPKL_FIX_PENDING);
      sessionStorage.removeItem(STORAGE.AUTO_FINISHED);
      UI.hideGlobalLoader(0);
      UI.showResult('info', 'Dibatalkan', 'Proses otomasi dihentikan oleh pengguna.');
    });

    delegate('click', '.qm-progress-container', () => {
      const tabBtn = document.querySelector('.qm-tab[data-pane="config"]');
      if (tabBtn) tabBtn.click();
      const btn = document.getElementById('qm-btn-show-logs');
      if (container && btn) {
        container.classList.remove('qm-hidden');
        btn.classList.add('qm-active');
        renderLogs();
        btn.querySelector('span').textContent = 'Sembunyikan Log';
      }
    });

    delegate('keydown', '#qm-spkl-online-nrp', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); document.getElementById('qm-btn-spkl-online-cek')?.click(); }
    });

    delegate('change', '#qm-fix-many-ot', function () {
      const box = document.getElementById('qm-fix-many-ot7-box');
      if (box) {
        if (this.value === '7') box.classList.remove('qm-hidden');
        else box.classList.add('qm-hidden');
      }
    });

    delegate('input', '#qm-fix-spkl-data', function () {
      const box = document.getElementById('qm-fix-spkl-ot7-box');
      if (box) {
        const has7 = this.value.split(/[,\n]+/).some(item => {
          const parts = item.trim().split(/[-:=]/);
          return parts.length > 1 && parts[1].trim() === '7';
        });
        if (has7) box.classList.remove('qm-hidden');
        else box.classList.add('qm-hidden');
      }
    });

    delegate('keydown', '#qm-input-nrp', function (e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        const nrp = this.value.trim();
        if (nrp.length >= 4) {
          refreshGlobalData(nrp);
        }
      }
    });

    // Sync other NRP inputs too
    delegate('input', '#qm-spkl-online-nrp, #qm-fix-spkl-nrp, #qm-input-hadir-nrp, #qm-input-hadir-bulan-nrp', function () {
      const nrp = this.value.trim();
      if (nrp.length >= 4) {
        refreshGlobalData(nrp);
      }
    });

    // Sync Month/Year changes
    delegate('change', '#qm-input-bulan, #qm-input-tahun, #qm-fix-spkl-bulan, #qm-fix-spkl-tahun, #qm-input-hadir-bulan-bln', function () {
      refreshGlobalData();
    });

    // General Keyboard Navigation for forms
    initKeyboardNavigation();

    delegate('click', '.qm-tab', onTabClick);
    delegate('click', '.qm-fix-dot', onFixDotClick);
    delegate('click', '.qm-btn-fix-pill', onFixDotClick);

    delegate('click', '#qm-btn-batch-check', function (e) {
      if (this.dataset.running) {
        onBatchCancel();
      } else {
        runBatchCheck();
      }
    });
    delegate('mouseover', '#qm-btn-batch-check', function (e) {
      if (this.dataset.running) {
        this.classList.add('qm-btn-danger');
        this.textContent = 'Batal';
      }
    });
    delegate('mouseout', '#qm-btn-batch-check', function (e) {
      if (this.dataset.running) {
        this.classList.remove('qm-btn-danger');
        this.textContent = 'Memproses...';
      }
    });
    delegate('click', '#qm-btn-export-batch', exportBatchResults);
    delegate('click', '.qm-batch-nrp-link', onBatchNrpClick);
    delegate('click', '.qm-batch-fix-btn', onBatchFixClick);

    // Batch Grouping (Bagian)
    delegate('click', '.qm-batch-group-header', function () {
      const targetSelector = this.dataset.target;
      const rows = document.querySelectorAll(targetSelector);
      const isExpanded = this.classList.contains('expanded');
      this.classList.toggle('expanded');
      rows.forEach(r => {
        if (isExpanded) {
          r.classList.add('qm-hidden');
          r.classList.remove('qm-table-row', 'expanded'); // Collapse sub-groups too
        } else {
          if (r.classList.contains('qm-batch-seksi-header')) {
            r.classList.remove('qm-hidden');
            r.classList.add('qm-table-row');
          } else {
            r.classList.add('qm-hidden');
            r.classList.remove('qm-table-row');
          }
        }
      });
    });

    // Batch Grouping (Seksi)
    delegate('click', '.qm-batch-seksi-header', function () {
      const targetSelector = this.dataset.target;
      const rows = document.querySelectorAll(targetSelector);
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
    });

    delegate('click', '.qm-batch-date-header', function (e) {
      if (e.target.closest('.qm-batch-fix-btn')) return;
      const content = this.nextElementSibling;
      this.classList.toggle('expanded');
      if (content) {
        content.classList.toggle('qm-hidden');
        content.classList.toggle('qm-visible-block');
      }
    });

    // 14. Accordion Toggle for FIX Panel
    delegate('click', '.qm-accordion-header', function () {
      this.classList.toggle('expanded');
      const content = this.nextElementSibling;
      if (content) {
        content.classList.toggle('qm-content-open');
      }
    });

    const collapseCheckbox = document.getElementById('qm-config-collapse-menu');
    if (collapseCheckbox) {
      collapseCheckbox.checked = alwaysCollapseMenu;
      delegate('change', '#qm-config-collapse-menu', function () {
        alwaysCollapseMenu = this.checked;
        GM_setValue('qm_always_collapse', alwaysCollapseMenu);
        if (alwaysCollapseMenu) {
          enforceSidebar();
        } else {
          document.body.classList.remove('enlarged');
        }
      });
    }

    // Theme Switcher
    delegate('click', '#qm-btn-theme-light', () => UI.applyTheme('light'));
    delegate('click', '#qm-btn-theme-dark', () => UI.applyTheme('dark'));

    enforceSidebar();
    const _sidebarObserver = new MutationObserver(enforceSidebar);
    _sidebarObserver.observe(document.body, { attributes: true, attributeFilter: ['class'] });

    const shortcutInput = document.getElementById('qm-input-shortcut');
    if (shortcutInput) shortcutInput.value = shortcutKey;
    delegate('click', '#qm-btn-record-shortcut', onRecordShortcut);

    // Sidebar Manual Override
    delegate('click', '.button-menu-mobile, .open-left, #sidebar-menu', function () {
      manualSidebarOverride = true;
    });

    // Make accordions keyboard accessible
    document.querySelectorAll('.qm-accordion-header').forEach(el => {
      if (!el.hasAttribute('tabindex')) el.setAttribute('tabindex', '0');
    });
  }

  /**
   * 16B. KEYBOARD NAVIGATION HELPERS
   */
  function initKeyboardNavigation() {
    // Enter key logic for various inputs
    delegate('keydown', '.qm-input, .qm-select, .qm-textarea', function (e) {
      if (e.key !== 'Enter') return;
      if (this.tagName === 'TEXTAREA' && !e.ctrlKey) return; // Allow newlines in textarea unless Ctrl+Enter

      const pane = this.closest('.qm-pane');
      if (!pane) return;

      const card = this.closest('.qm-card');
      if (!card) return;

      // Find the primary button in the same card
      const primaryBtn = card.querySelector('.qm-btn-primary') || card.querySelector('.qm-btn');
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
    delegate('keydown', '.qm-accordion-header', function (e) {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        this.click();
      }
    });
  }

  /* ============================================================
   * 24. JK CHANGE LOGIC
   * ============================================================ */
  function initJkChangeEvents() {
    delegate('click', '#qm-btn-KK-update', onUpdateKKMaster);

    let globalDebounce;
    delegate('input', '#qm-input-distribusi-nrp, #qm-dist-KK-nrp, #qm-input-distribusi-subsi-nrp, #qm-input-nrp', function () {
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
  async function refreshGlobalData(nrp = '', bulan = '', tahun = '') {
    // 1. Detection
    if (!nrp) {
      nrp = document.getElementById('qm-input-nrp')?.value.trim() ||
        document.getElementById('qm-input-distribusi-nrp')?.value.trim() ||
        document.getElementById('qm-fix-spkl-nrp')?.value.trim() ||
        getPageContext().nrp || '';
    }
    if (!bulan) {
      bulan = document.getElementById('qm-input-bulan')?.value ||
        document.getElementById('qm-fix-spkl-bulan')?.value ||
        new Date().getMonth() + 1;
    }
    if (!tahun) {
      tahun = document.getElementById('qm-input-tahun')?.value ||
        document.getElementById('qm-fix-spkl-tahun')?.value ||
        new Date().getFullYear();
    }

    if (!nrp || nrp.length < 4) return;

    // 2. Sync all inputs across tabs
    syncGlobalInputs(nrp, bulan, tahun);

    // 3. Show loading in summary area
    const resBody = document.getElementById('qm-result-body');
    if (resBody && document.getElementById('qm-pane-check-nrp').classList.contains('active')) {
      setInnerHTML(resBody, '<div class="qm-flex qm-items-center qm-gap-s qm-p-m"><span class="qm-spinner"></span> <span>Memuat data karyawan...</span></div>');
    }

    try {
      // 4. Single getEmp call (cached internally)
      const emp = await getEmp(nrp);
      if (!emp.found) {
        if (resBody) setInnerHTML(resBody, '<div class="qm-text-danger qm-p-m">NRP tidak ditemukan.</div>');
        return;
      }

      // 5. Update Summary Info
      if (resBody) {
        setInnerHTML(resBody, `
          <div class="qm-p-m qm-bg-parchment qm-rounded-m qm-border">
            <div class="qm-flex qm-justify-between qm-mb-s">
              <span class="qm-text-muted qm-text-xs">Nama:</span>
              <span class="qm-font-bold qm-text-s">${escHtml(emp.nama)}</span>
            </div>
            <div class="qm-flex qm-justify-between qm-mb-s">
              <span class="qm-text-muted qm-text-xs">Bagian:</span>
              <span class="qm-text-s">${escHtml(emp.bagian || '-')}</span>
            </div>
            <div class="qm-flex qm-justify-between qm-mb-s">
              <span class="qm-text-muted qm-text-xs">Jam Kerja:</span>
              <span class="qm-text-s qm-font-mono qm-text-blue">${escHtml(emp.jk || '-')}</span>
            </div>
            <div class="qm-flex qm-justify-between">
              <span class="qm-text-muted qm-text-xs">Kalender:</span>
              <span class="qm-text-s qm-font-mono qm-text-teal">${escHtml(emp.KK || '-')}</span>
            </div>
          </div>
        `);
      }

      // 6. Refresh Active Pane Data
      const activePane = document.querySelector('.qm-pane.active')?.id;
      if (activePane === 'qm-pane-distribusi') {
        await refreshDistribusiOptions(nrp, emp);
      } else if (activePane === 'qm-pane-spkl') {
        // Additional SPKL specific refresh if needed
      }

    } catch (e) {
      Logger.error('refreshGlobalData error', e);
      if (resBody) setInnerHTML(resBody, `<div class="qm-text-danger qm-p-m">Error: ${e.message}</div>`);
    }
  }

  function syncGlobalInputs(nrp, bulan, tahun) {
    // NRP inputs
    const nrpIds = ['qm-input-nrp', 'qm-spkl-online-nrp', 'qm-fix-spkl-nrp', 'qm-input-hadir-nrp', 'qm-input-hadir-bulan-nrp', 'qm-input-distribusi-nrp', 'qm-dist-KK-nrp'];
    nrpIds.forEach(id => {
      const el = document.getElementById(id);
      if (el && el.value !== nrp) el.value = nrp;
    });

    // Month selects
    const bulanIds = ['qm-input-bulan', 'qm-fix-spkl-bulan', 'qm-input-hadir-bulan-bln'];
    bulanIds.forEach(id => {
      const el = document.getElementById(id);
      if (el && el.value != bulan) el.value = bulan;
    });

    // Year selects
    const tahunIds = ['qm-input-tahun', 'qm-fix-spkl-tahun', 'qm-input-hadir-bulan-thn'];
    tahunIds.forEach(id => {
      const el = document.getElementById(id);
      if (el && el.value != tahun) el.value = tahun;
    });

    // Date/Month picker sync
    const kkDate = document.getElementById('qm-dist-KK-date');
    if (kkDate) {
      const mStr = String(bulan).padStart(2, '0');
      kkDate.value = `${tahun}-${mStr}`;
    }
  }

  async function refreshDistribusiOptions(nrp, emp) {
    const jkContainers = [
      document.getElementById('qm-dist-jk-options-container'),
      document.getElementById('qm-dist-subsi-jk-container')
    ];
    const kkContainer = document.getElementById('qm-dist-KK-options-container');

    jkContainers.forEach(c => { if (c) setInnerHTML(c, '<span class="qm-spinner qm-spinner-xs"></span>'); });
    if (kkContainer) setInnerHTML(kkContainer, '<span class="qm-spinner qm-spinner-xs"></span>');

    try {
      // 1. Fetch JK & KK options (using cache internally if available)
      const [jkOptions, kkOptions] = await Promise.all([
        fetchJkOptions(nrp),
        fetchKKOptions(nrp)
      ]);

      jkContainers.forEach(c => {
        if (c && jkOptions.length) {
          const id = c.id === 'qm-dist-jk-options-container' ? 'qm-dist-jk-select-input' : 'qm-dist-subsi-jk-select-input';
          setInnerHTML(c, `<select id="${id}" class="qm-select qm-text-s">${jkOptions.map(o => `<option value="${escHtml(o.val)}" ${o.selected ? 'selected' : ''}>${escHtml(o.txt)}</option>`).join('')}</select>`);
        }
      });

      if (kkContainer && kkOptions.length) {
        setInnerHTML(kkContainer, `<select id="qm-dist-KK-select-input" class="qm-select qm-text-s">${kkOptions.map(o => `<option value="${escHtml(o.val)}" ${o.selected ? 'selected' : ''}>${escHtml(o.txt)}</option>`).join('')}</select>`);
      }

      // 2. Fetch and populate Bagian/Seksi/Grup/Shift
      // We don't await this fully for the first render if we want it "instant"
      // but we need the options to be there. updateDistribusiDropdowns now handles its own internal cache.
      await updateDistribusiDropdowns(nrp);

      // 3. Set values from emp data
      const setVal = (id, val) => {
        const el = document.getElementById(id);
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
      Logger.error('refreshDistribusiOptions error', e);
    }
  }

  function attachDistribusiListeners(nrp) {
    const pairs = [
      { bag: 'qm-input-distribusi-subsi-bagian', sek: 'qm-input-distribusi-subsi-seksi' },
      { bag: 'qm-dist-KK-bagian', sek: 'qm-dist-KK-seksi' }
    ];

    pairs.forEach(p => {
      const elBag = document.getElementById(p.bag);
      const elSek = document.getElementById(p.sek);

      if (elBag && !elBag.dataset.hasListener) {
        elBag.dataset.hasListener = 'true';
        elBag.addEventListener('change', () => updateDistribusiDropdowns(nrp, elBag.value, ''));
      }
      if (elSek && !elSek.dataset.hasListener) {
        elSek.dataset.hasListener = 'true';
        elSek.addEventListener('change', () => updateDistribusiDropdowns(nrp, elBag?.value || '', elSek.value));
      }
    });
  }

  /**
   * Fetches the distribution page and synchronizes all dropdowns (Per Subsi & Kalender Kerja).
   */
  async function updateDistribusiDropdowns(nrp, bag = '', sek = '') {
    const isOS = nrp.length === 8;
    const cacheKey = `qm_dist_html_${isOS ? 'os' : 'reg'}_${bag}_${sek}`;

    // Attempt to use cache for immediate UI update
    const cachedHtml = sessionStorage.getItem(cacheKey);
    if (cachedHtml) {
      applyDistDropdowns(cachedHtml);
    }

    let url = getDistribusiBaseUrl(nrp);
    if (bag || sek) {
      url += `?kode_bagian=${encodeURIComponent(bag)}&kode_seksi=${encodeURIComponent(sek)}`;
    }

    try {
      const html = await req(url);
      sessionStorage.setItem(cacheKey, html);
      applyDistDropdowns(html);
    } catch (e) {
      Logger.error('updateDistribusiDropdowns fetch error', e);
    }
  }

  function applyDistDropdowns(html) {
    const doc = parseHTML(html);
    const sync = (targetIds, selector) => {
      const source = doc.querySelector(selector);
      if (!source) return;
      targetIds.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
          const currentVal = el.value;
          const firstOpt = el.options[0];
          setInnerHTML(el, source.innerHTML);
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
    const shiftEl = document.getElementById('qm-input-distribusi-subsi-shift');
    if (shiftEl && shiftEl.options.length <= 1) {
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

    const emp = await getEmp(nrp);
    if (!emp.found || !emp.id) throw new Error('Data karyawan tidak lengkap.');

    let editUrl = emp.editUrl;
    if (!editUrl) {
      // Fallback guess if not found in detail page
      const isOS = nrp.length === 8;
      editUrl = getEmployeeEditUrl(nrp, emp.id);
    }

    let html = '', select = null;
    try {
      html = await req(editUrl);
      cachedEditHtml = html; // Cache HTML for later saveJkMaster use
      const doc = parseHTML(html);

      // Try multiple possible selectors
      const possibleNames = ['kerja_hour_code', 'kode_jam_kerja', 'jam_kerja', 'kerja_hour'];
      for (const name of possibleNames) {
        select = doc.querySelector(`select[name="${name}"], [name="${name}"]`);
        if (select) break;
      }

      // Diagnostics if still not found
      if (!select) {
        const title = doc.title || 'Tanpa Judul';
        const h1 = doc.querySelector('h1, h2, h3')?.textContent.trim() || 'Tanpa Header';
        Logger.warn(`Gagal menemukan JK di ${editUrl}. Judul Page: ${title}, Header: ${h1}`);
      }
    } catch (e) {
      Logger.warn(`Gagal mengambil JK dari ${editUrl}`, e);
    }

    if (!select) {
      throw new Error(`Elemen pilihan jam kerja tidak ditemukan di ${editUrl}. Pastikan Anda memiliki akses edit.`);
    }

    const options = Array.from(select.querySelectorAll('option')).map(opt => ({
      val: opt.value,
      txt: opt.textContent.trim(),
      selected: opt.hasAttribute('selected') || opt.selected
    }));

    sessionStorage.setItem('qm_jk_options_' + nrp, JSON.stringify(options));
    return options;
  }

  async function onSaveJkChange() {
    const nrp = document.getElementById('qm-input-distribusi-nrp')?.value.trim();
    if (!nrp) { alert('Harap isi NRP.'); return; }

    const useDistribusi = document.getElementById('qm-dist-jk-use-distribusi')?.checked;
    const jk = document.getElementById('qm-dist-jk-select-input')?.value;
    const date = document.getElementById('qm-dist-jk-target-date')?.value;
    const dateEnd = document.getElementById('qm-dist-jk-target-date-end')?.value;
    const shift = document.getElementById('qm-dist-jk-target-shift')?.value;
    const oldJk = sessionStorage.getItem('qm_jk_' + nrp);

    const emp = await getEmp(nrp);
    if (!emp.found) {
      alert('Data karyawan tidak ditemukan untuk NRP ' + nrp);
      return;
    }

    Logger.info(`Starting onSaveJkChange for ${nrp}. New JK: ${jk}, Use Distribusi: ${useDistribusi}`);
    UI.showGlobalLoader('Processing JK', 'Updating Master Data...');

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

        // Update UI in-place for immediate feedback
        const jkLabel = document.getElementById('qm-jk-value');
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
          UI.showResult('warning', 'Master OK, Distribusi Gagal', distErr.message);
          UI.hideGlobalLoader(5000);
        }
      } else {
        // Redirect mode
        Logger.info(`Master Data updated for ${nrp}. Proceeding to distribution via redirect...`);
        UI.setGlobalProgress(80, 'Master terupdate. Mengalihkan ke halaman Distribusi...');

        // Store return URL for auto-return after redirect completion
        sessionStorage.setItem(STORAGE.RETURN_URL, window.location.href);
        const redirectUrl = getDistribusiLink(ctx, date, shift, dateEnd);

        // Small delay to ensure sessionStorage/Master Update is committed
        setTimeout(() => {
          window.location.href = redirectUrl;
        }, 800);
      }
    } catch (e) {
      Logger.error(`Error in onSaveJkChange: ${e.message}`, e);
      UI.showResult('danger', 'Gagal', 'Terjadi kesalahan: ' + e.message);
      UI.hideGlobalLoader();
    }
  }

  /**
   * Unified background distribution logic used by both Per NRP and Per Subsi.
   */
  async function executeBackgroundDistribusi(params) {
    const { nrp, jk, tglAwal, tglAkhir, shift, bagian, seksi, grup } = params;
    const isOS = nrp && nrp.length === 8;
    const distUrl = getDistribusiBaseUrl(nrp);

    UI.setGlobalProgress(10, 'Mengambil form distribusi...');
    const html = await req(distUrl);
    if (html.includes('id="login-form"') || html.includes('login_form')) throw new Error('Sesi berakhir. Silakan login kembali.');

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

    const submitBtn = form.querySelector('button[type="submit"], input[type="submit"]');
    if (submitBtn && submitBtn.name) {
      postData.set(submitBtn.name, submitBtn.value || 'Start Distribusi');
    } else {
      postData.set('btnSubmit', 'Start Distribusi');
    }

    UI.setGlobalProgress(70, 'Mengirim data (POST)... Harap tunggu, proses ini lama.');
    const action = form.getAttribute('action') || distUrl;
    const targetUrl = action.startsWith('http') ? action : `https://hris.kti.co.id${action}`;

    Logger.info(`Sending background distribution POST to ${targetUrl}...`);
    Logger.info('POST Keys:', Array.from(postData.keys()).join(', '));

    try {
      const response = await fetchWithTimeout(targetUrl, {
        method: 'POST',
        headers: headers,
        body: postData,
        referrer: distUrl,
        referrerPolicy: 'origin-when-cross-origin'
      }, 300000);

      if (!response.ok) throw new Error('Distribusi gagal (HTTP ' + response.status + ')');

      UI.setGlobalProgress(90, 'Membaca respon server...');
      const resText = await response.text();
      Logger.info('Response received from distribution server.');

      if (isSuccessResponse(resText)) {
        UI.setGlobalProgress(100, 'Selesai!');
        return true;
      } else {
        Logger.warn('Respon Distribusi (Bukan sukses)', resText.substring(0, 500));
        return false;
      }
    } catch (fetchErr) {
      if (fetchErr.name === 'AbortError') {
        throw new Error('Distribusi timeout (300 detik). Harap cek secara manual.');
      }
      throw fetchErr;
    }
  }

  // Removed in favor of unified refreshDistribusiPane
  // async function refreshSubsiJkOptionsInPane() { ... }

  // Removed in favor of unified updateDistribusiDropdowns
  // async function updateSubsiDropdowns(nrp, bag = '', sek = '') { ... }

  async function onProsesDistribusiSubsi() {
    const nrp = getPageContext().nrp;
    if (!nrp) { alert('Gagal mendeteksi NRP. Pastikan Anda berada di halaman profile atau tabel kehadiran.'); return; }

    const useDistribusi = document.getElementById('qm-dist-subsi-use-distribusi')?.checked;
    const jk = document.getElementById('qm-dist-subsi-jk-select-input')?.value;
    const tglAwal = document.getElementById('qm-input-distribusi-subsi-tgl-awal')?.value;
    const tglAkhir = document.getElementById('qm-input-distribusi-subsi-tgl-akhir')?.value;
    const shift = document.getElementById('qm-input-distribusi-subsi-shift')?.value;
    const bagian = document.getElementById('qm-input-distribusi-subsi-bagian')?.value || '';
    const seksi = document.getElementById('qm-input-distribusi-subsi-seksi')?.value || '';
    const grup = document.getElementById('qm-input-distribusi-subsi-grup')?.value || '';

    if (!jk) { alert('Tunggu opsi Jam Kerja termuat terlebih dahulu.'); return; }
    if (!tglAwal || !tglAkhir) { alert('Harap isi Tanggal Awal dan Akhir.'); return; }

    Logger.info(`Starting onProsesDistribusiSubsi. JK: ${jk}, Bagian: ${bagian}, Seksi: ${seksi}, Grup: ${grup}`);

    if (useDistribusi) {
      UI.showGlobalLoader('Processing Subsi', 'Menyiapkan data...');
      try {
        const isOS = nrp.length === 8;
        const nrpAwal = isOS ? '00000000' : '0000';
        const nrpAkhir = isOS ? '99999999' : '9999';

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
        setTimeout(() => window.location.reload(), 1500);
      } catch (e) {
        UI.showResult('danger', 'Gagal', 'Error: ' + e.message);
        Logger.error('onProsesDistribusiSubsi error', e);
        UI.hideGlobalLoader();
      }
    } else {
      const base = getDistribusiBaseUrl(nrp);
      const url = `${base}?qm_auto_distribusi_subsi=1&jk=${encodeURIComponent(jk)}&tglAwal=${encodeURIComponent(tglAwal)}&tglAkhir=${encodeURIComponent(tglAkhir)}&bagian=${encodeURIComponent(bagian)}&seksi=${encodeURIComponent(seksi)}&grup=${encodeURIComponent(grup)}&shift=${encodeURIComponent(shift)}&nrp=${encodeURIComponent(nrp)}`;

      // Store return URL for auto-return after redirect completion
      sessionStorage.setItem(STORAGE.RETURN_URL, window.location.href);
      window.location.href = url;
    }
  }

  async function saveJkMaster(nrp, jk) {
    Logger.info(`saveJkMaster started for NRP ${nrp} with JK ${jk}`);
    const emp = await getEmp(nrp);
    if (!emp.found || !emp.id) {
      Logger.error(`Data karyawan tidak lengkap untuk NRP ${nrp}`);
      throw new Error('Data karyawan tidak lengkap.');
    }

    let editUrl = emp.editUrl;
    if (!editUrl) {
      editUrl = getEmployeeEditUrl(nrp, emp.id);
    }

    Logger.info(`Fetching edit form from ${editUrl}...`);
    // NRP-aware caching for cachedEditHtml
    if (state._lastEditNrp !== nrp) {
      cachedEditHtml = null;
      state._lastEditNrp = nrp;
    }

    const html = cachedEditHtml || await req(editUrl);
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
    const finished = sessionStorage.getItem(STORAGE.AUTO_FINISHED);

    if (oldJk && finished === 'true') {
      sessionStorage.removeItem('qm_jk_to_restore_' + nrp);
      sessionStorage.removeItem(STORAGE.AUTO_FINISHED);

      UI.showGlobalLoader('Cleaning up', 'Restoring Master JK...');
      try {
        await saveJkMaster(nrp, oldJk);
        sessionStorage.setItem('qm_jk_' + nrp, oldJk);
        UI.showResult('success', 'Selesai', 'Kode Jam Kerja master telah dikembalikan.');
        setTimeout(() => window.location.reload(), 1500);
      } catch (e) {
        UI.showResult('danger', 'Gagal Restore', 'Gagal mengembalikan JK: ' + e.message);
        UI.hideGlobalLoader();
      }
    }
  }

  /* ============================================================
   * 25. KK CHANGE LOGIC
   * ============================================================ */
  // Removed in favor of unified refreshDistribusiPane
  // async function refreshKKOptionsInPane() { ... }

  // Removed in favor of unified refreshDistribusiPane
  // async function updateKKDropdowns(nrp, bag = '', sek = '') { ... }

  async function fetchKKOptions(nrp) {
    const cached = sessionStorage.getItem('qm_KK_options_' + nrp);
    if (cached) return JSON.parse(cached);

    const emp = await getEmp(nrp);
    if (!emp.found || !emp.id) throw new Error('Data karyawan tidak lengkap.');

    const editUrl = emp.editUrl || getEmployeeEditUrl(nrp, emp.id);

    const html = await req(editUrl);
    const doc = parseHTML(html);
    const select = doc.querySelector('select[name="kode_kalender_kerja"]');

    if (!select) throw new Error('Elemen kode_kalender_kerja tidak ditemukan di ' + editUrl);

    const options = Array.from(select.querySelectorAll('option')).map(opt => ({
      val: opt.value,
      txt: opt.textContent.trim(),
      selected: opt.hasAttribute('selected') || opt.selected
    }));

    sessionStorage.setItem('qm_KK_options_' + nrp, JSON.stringify(options));
    return options;
  }

  async function onUpdateKKMaster() {
    const nrpInput = document.getElementById('qm-dist-KK-nrp');
    const dateInput = document.getElementById('qm-dist-KK-date');
    const nrp = nrpInput?.value.trim();
    const dateVal = dateInput?.value; // "YYYY-MM"
    const KK = document.getElementById('qm-dist-KK-select-input')?.value;

    if (!nrp) { alert('Harap isi NRP.'); return; }
    if (!KK) { alert('Pilih Kalender Kerja.'); return; }
    if (!dateVal) { alert('Pilih periode (bulan/tahun).'); return; }

    const [tahun, bulan] = dateVal.split('-');
    const bagian = document.getElementById('qm-dist-KK-bagian')?.value;
    const seksi = document.getElementById('qm-dist-KK-seksi')?.value;
    const grup = document.getElementById('qm-dist-KK-grup')?.value;

    UI.showGlobalLoader('Processing KK', 'Checking current data...');
    try {
      const emp = await getEmp(nrp);
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
      setTimeout(() => window.location.reload(), 1500);
    } catch (e) {
      UI.showResult('danger', 'Gagal', 'Error: ' + e.message);
      Logger.error('onUpdateKKMaster error', e);
    } finally {
      UI.hideGlobalLoader(3000);
    }
  }

  async function saveKKMaster(nrp, KK) {
    const emp = await getEmp(nrp);
    const editUrl = emp.editUrl || getEmployeeEditUrl(nrp, emp.id);

    const html = await req(editUrl);
    if (html.includes('id="login-form"') || html.includes('name="login_form"') || html.includes('login-box')) {
      throw new Error('Sesi berakhir. Silakan login kembali.');
    }
    const doc = parseHTML(html);
    const select = doc.querySelector('select[name="kode_kalender_kerja"]');
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
    const html = await req(distUrl);
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
    if (isSuccessResponse(resText)) {
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
      // Result page check
      if (sessionStorage.getItem(STORAGE.AUTO_FINISHED) === 'true' && isDistribusiKalenderPagePath()) {
        const pageText = document.body.textContent;
        if (pageText.includes('Distribution Process Completed') || document.querySelector('.alert-success')) {
          UI.showResult('success', 'Selesai', 'Distribusi Kalender Kerja Selesai.');
          const returnUrl = sessionStorage.getItem(STORAGE.RETURN_URL);
          if (returnUrl) {
            setTimeout(() => {
              sessionStorage.removeItem(STORAGE.RETURN_URL);
              sessionStorage.removeItem(STORAGE.AUTO_FINISHED);
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
      const emp = await getEmp(nrp);
      if (!emp.found) throw new Error('Data karyawan tidak ditemukan.');

      // 1. Month
      const selMonth = document.querySelector('select[name="month"]');
      if (selMonth) setFieldValue(selMonth, month);

      // 2. Sequential Dropdowns: Bagian -> Seksi
      UI.setGlobalProgress(40, 'Menyesuaikan Bagian...');
      pilihDropdownDinamis('#kode_bagian', emp.bagian, () => {
        UI.setGlobalProgress(70, 'Menyesuaikan Seksi...');
        pilihDropdownDinamis('#kode_seksi', emp.seksi, () => {
          UI.setGlobalProgress(85, 'Mengisi NRP...');

          // 3. NRP Initial & Final
          const nrpInit = document.getElementById('nrp_initial_text');
          const nrpFinal = document.getElementById('nrp_final_text');
          if (nrpInit) setFieldValue(nrpInit, nrp);
          if (nrpFinal) setFieldValue(nrpFinal, nrp);

          // 4. Submit
          const btnSubmit = document.getElementById('btnSubmit');
          if (btnSubmit) {
            UI.setGlobalProgress(95, 'Submitting...');
            sessionStorage.setItem(STORAGE.AUTO_FINISHED, 'true');
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
   * 26. BOOT
   * ============================================================ */
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
