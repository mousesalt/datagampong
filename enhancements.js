/*
 * Data Gampong - Fitur tambahan
 * - Ekspor semua dusun dalam satu workbook Excel
 * - Setiap dusun mempunyai sheet sendiri
 * - Membaca nama dusun otomatis dari kolom Excel DUSUN/LORONG/NAMA DUSUN
 * - Menyertakan MASTER_DUSUN agar daftar dusun kosong ikut saat workbook dipindah perangkat
 *
 * Letakkan file ini di folder yang sama dengan index.html dan app.js.
 * Muat setelah app.js: <script src="enhancements.js"></script>
 */
(function () {
  'use strict';

  const STORAGE_KEY = 'datagampong.masterDusun.v1';
  // Bersihkan daftar lama yang sebelumnya tercampur nilai dari sheet selain DATAPENDUDUK.
  try { localStorage.removeItem(STORAGE_KEY); } catch (_) {}

  // Kenali hanya judul kolom nama dusun yang memang digunakan pada sheet DATAPENDUDUK.
  function normalizeHeaderName(value) {
    return String(value == null ? '' : value)
      .toLocaleUpperCase('id-ID')
      .replace(/[^A-Z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  const DUSUN_HEADERS = new Set([
    'LORONG', 'DUSUN', 'NAMA LORONG', 'NAMA DUSUN',
    'DUSUN LORONG', 'LORONG DUSUN', 'DUSUN NAMA', 'LORONG NAMA',
    'DUSUN LORONG NAMA', 'LORONG DUSUN NAMA', 'NAMA DUSUN LORONG', 'NAMA LORONG DUSUN'
  ]);

  function isDusunHeader(value) {
    return DUSUN_HEADERS.has(normalizeHeaderName(value));
  }

  const originalColumnKey = key;
  key = function (header) {
    if (isDusunHeader(header)) return 'lorong';
    return originalColumnKey(header);
  };

  const originalValueGetter = val;
  val = function (row, field) {
    if (field === 'lorong' && row && typeof row === 'object') {
      const priority = ['LORONG', 'DUSUN', 'NAMA LORONG', 'NAMA DUSUN', 'DUSUN LORONG', 'LORONG DUSUN'];
      const headers = Object.keys(row).filter(header => isDusunHeader(header));
      headers.sort((a, b) => {
        const ai = priority.indexOf(normalizeHeaderName(a));
        const bi = priority.indexOf(normalizeHeaderName(b));
        return (ai < 0 ? 100 : ai) - (bi < 0 ? 100 : bi);
      });
      for (const header of headers) {
        const value = clean(row[header]);
        if (value) return row[header];
      }
      return '';
    }
    return originalValueGetter(row, field);
  };

  function normalizeDusunValue(value) {
    return String(value == null ? '' : value).trim().replace(/\s+/g, ' ').toLocaleUpperCase('id-ID');
  }

  function sameDusun(left, right) {
    return normalizeDusunValue(left) === normalizeDusunValue(right);
  }

  function uniqueNames(list) {
    const result = new Map();
    list.forEach(value => {
      const name = String(value == null ? '' : value).trim().replace(/\s+/g, ' ');
      // Nilai angka murni yang sempat masuk dari sheet rekap lama bukan nama dusun.
      if (!name || /^\d+$/.test(name) || /^[-–—]+$/.test(name)) return;
      const normalized = normalizeDusunValue(name);
      if (!result.has(normalized)) result.set(normalized, name);
    });
    return Array.from(result.values()).sort((a, b) => a.localeCompare(b, 'id'));
  }

  function isResidentRecord(row) {
    // Baris yang hanya berisi nama dusun tanpa identitas warga bukan data penduduk.
    return Boolean(clean(val(row, 'nama')) || clean(val(row, 'nik')) || clean(val(row, 'no_kk')));
  }

  function allDusunNames() {
    // Nama dusun bersumber HANYA dari baris penduduk yang benar-benar berisi identitas.
    return uniqueNames(rows.filter(isResidentRecord).map(row => clean(val(row, 'lorong'))));
  }

  // Filter dusun mengikuti nama yang benar-benar muncul pada data penduduk.
  const originalMakeSelect = makeSelect;
  makeSelect = function (id, label, key) {
    if (key !== 'lorong') return originalMakeSelect(id, label, key);
    const values = allDusunNames();
    if (!values.length) return '';
    return `<select id="${id}" data-key="${key}">
      <option value="">${label}</option>
      ${values.map(name => `<option value="${esc(name)}">${esc(name)}</option>`).join('')}
    </select>`;
  };

  // Rekap dusun dibentuk hanya dari warga pada sheet DATAPENDUDUK aktif.
  renderDusunReport = function () {
    const statsByDusun = new Map();
    rows.forEach(row => {
      if (!isResidentRecord(row)) return;
      const rawName = clean(val(row, 'lorong'));
      if (!rawName || /^\d+$/.test(rawName)) return;
      const normalized = normalizeDusunValue(rawName);
      if (!statsByDusun.has(normalized)) statsByDusun.set(normalized, { name: rawName, rows: [], l: 0, p: 0, kk: 0 });
      const item = statsByDusun.get(normalized);
      item.rows.push(row);
      if (norm(val(row, 'jk')) === 'L') item.l++;
      if (norm(val(row, 'jk')) === 'P') item.p++;
      if (isHead(row)) item.kk++;
    });
    const entries = Array.from(statsByDusun.values()).sort((a, b) => a.name.localeCompare(b.name, 'id'));
    $('dusun').innerHTML = '<thead><tr><th>DUSUN/LORONG</th><th>L</th><th>P</th><th>TOTAL</th><th>KK</th><th>AKSI</th></tr></thead><tbody>' +
      (entries.length ? entries.map(item => `<tr><td><b>${esc(item.name)}</b></td><td>${item.l}</td><td>${item.p}</td><td><b>${item.rows.length}</b></td><td>${item.kk}</td><td>${reportButton(item.name)}</td></tr>`).join('') : '<tr><td colspan="6" class="empty">Belum ada data dusun pada sheet DATAPENDUDUK.</td></tr>') + '</tbody>';
    document.querySelectorAll('#dusun .viewBtn').forEach(button => {
      const name = button.dataset.report;
      button.onclick = () => applyReportFilter(`Dusun/Lorong: ${name}`, row => sameDusun(val(row, 'lorong'), name));
    });
  };

  function headsForDusun(name) {
    return rows
      .filter(row => isResidentRecord(row) && isHead(row) && sameDusun(val(row, 'lorong'), name))
      .sort((a, b) => clean(val(a, 'nama')).localeCompare(clean(val(b, 'nama')), 'id'));
  }

  // Sheet untuk satu dusun: tetap dapat diekspor meskipun belum ada kepala keluarga.
  function makeDusunSheet(name) {
    const list = headsForDusun(name);
    const data = [
      ['DAFTAR KEPALA KELUARGA PER DUSUN'],
      ['DUSUN / LORONG', name],
      [],
      ['NO', 'NAMA', 'JENIS KELAMIN', 'PEKERJAAN', 'TANDA TANGAN'],
      ...list.map((row, index) => [index + 1, clean(val(row, 'nama')), clean(val(row, 'jk')), clean(val(row, 'pekerjaan')), ''])
    ];
    const sheet = XLSX.utils.aoa_to_sheet(data);
    sheet['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
      { s: { r: 1, c: 1 }, e: { r: 1, c: 4 } }
    ];
    sheet['!cols'] = [{ wch: 7 }, { wch: 32 }, { wch: 18 }, { wch: 30 }, { wch: 32 }];
    sheet['!rows'] = [];
    sheet['!rows'][0] = { hpt: 24 };
    sheet['!rows'][1] = { hpt: 20 };
    sheet['!rows'][3] = { hpt: 24 };
    for (let row = 4; row < data.length; row++) sheet['!rows'][row] = { hpt: 42 };
    ['A1', 'A2'].forEach(address => {
      if (sheet[address]) sheet[address].s = { font: { bold: true }, alignment: { horizontal: 'center', vertical: 'center' } };
    });
    ['A4', 'B4', 'C4', 'D4', 'E4'].forEach(address => {
      if (sheet[address]) sheet[address].s = { font: { bold: true }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } };
    });
    return sheet;
  }

  function safeFilePart(value) {
    return String(value).replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, '-').replace(/-+/g, '-').replace(/^-|-$/g, '') || 'Dusun';
  }

  function safeSheetName(value, workbook) {
    const base = String(value || 'Dusun').replace(/[\\/?:*\[\]]/g, ' ').trim().slice(0, 31) || 'Dusun';
    const used = new Set(workbook.SheetNames.map(name => name.toLocaleLowerCase('id-ID')));
    if (!used.has(base.toLocaleLowerCase('id-ID'))) return base;
    let number = 2;
    while (true) {
      const suffix = ` (${number})`;
      const candidate = base.slice(0, 31 - suffix.length) + suffix;
      if (!used.has(candidate.toLocaleLowerCase('id-ID'))) return candidate;
      number++;
    }
  }

  function appendUniqueSheet(workbook, name, sheet) {
    XLSX.utils.book_append_sheet(workbook, sheet, safeSheetName(name, workbook));
  }

  function exportDusun(name) {
    const workbook = XLSX.utils.book_new();
    appendUniqueSheet(workbook, name, makeDusunSheet(name));
    XLSX.writeFile(workbook, `Daftar-Kepala-Keluarga-${safeFilePart(name)}.xlsx`);
  }

  renderKKReport = function () {
    const groups = new Map();
    rows.forEach(row => {
      if (!isHead(row)) return;
      const name = clean(val(row, 'lorong'));
      if (!name || /^\d+$/.test(name)) return;
      const normalized = normalizeDusunValue(name);
      if (!groups.has(normalized)) groups.set(normalized, { name, rows: [] });
      groups.get(normalized).rows.push(row);
    });
    const entries = Array.from(groups.values()).sort((a, b) => a.name.localeCompare(b.name, 'id'));
    const total = entries.reduce((sum, item) => sum + item.rows.length, 0);
    $('kkTotal').textContent = `${total.toLocaleString('id-ID')} kepala keluarga · ${entries.length.toLocaleString('id-ID')} dusun`;
    $('kkReport').innerHTML = entries.map(item => {
      const list = item.rows.sort((a, b) => clean(val(a, 'nama')).localeCompare(clean(val(b, 'nama')), 'id'));
      const safeName = esc(item.name).replaceAll('"', '&quot;');
      const table = list.map((row, index) => `<tr><td>${index + 1}</td><td>${esc(val(row, 'nama'))}</td><td>${esc(val(row, 'jk'))}</td><td>${esc(val(row, 'pekerjaan'))}</td><td class="signature">&nbsp;</td></tr>`).join('');
      return `<div class="kkGroup"><div class="kkTitle"><h3>${esc(item.name)} <small>(${list.length.toLocaleString('id-ID')} KK)</small></h3><button class="excelBtn" data-export-dusun="${safeName}" type="button">📊 Excel</button></div><div class="table"><table><thead><tr><th>NO</th><th>NAMA</th><th>JENIS KELAMIN</th><th>PEKERJAAN</th><th>TANDA TANGAN</th></tr></thead><tbody>${table}</tbody></table></div></div>`;
    }).join('') || '<div class="empty">Belum ada kepala keluarga yang terdata pada dusun di sheet DATAPENDUDUK.</div>';
    document.querySelectorAll('#kkReport .excelBtn').forEach(button => {
      button.onclick = () => exportDusun(button.dataset.exportDusun);
    });
  };


  function exportAllDusun() {
    if (!Array.isArray(rows) || rows.length === 0) {
      alert('Muat file DATA-GAMPONG.xlsx terlebih dahulu.');
      return;
    }

    // Ekspor hanya daftar kepala keluarga. Setiap dusun menjadi satu sheet;
    // jangan sertakan seluruh database penduduk, master, atau rekap tambahan.
    const names = allDusunNames().filter(name => headsForDusun(name).length > 0);
    if (!names.length) {
      alert('Belum ada kepala keluarga dengan nama dusun yang terisi.');
      return;
    }

    const workbook = XLSX.utils.book_new();
    names.forEach(name => appendUniqueSheet(workbook, name, makeDusunSheet(name)));
    XLSX.writeFile(workbook, 'Daftar-Kepala-Keluarga-Semua-Dusun.xlsx');
  }

  function addControlPanel() {
    const report = document.getElementById('kkReport');
    if (!report || document.getElementById('kkExtraActions')) return;

    const wrapper = document.createElement('div');
    wrapper.id = 'kkExtraActions';
    wrapper.innerHTML = `
      <div class="kkExtraButtons">
        <button type="button" id="exportAllDusun" class="excelBtn">📥 Export Semua Dusun</button>
      </div>
      <p class="kkExtraHelp">Daftar ini hanya menampilkan dusun yang memiliki kepala keluarga pada sheet DATAPENDUDUK. Nama dusun dibaca dari kolom LORONG/DUSUN/NAMA LORONG/NAMA DUSUN. Untuk menambahkan dusun, isi nama dusun pada baris penduduk di Excel, lalu muat ulang file. Export Semua Dusun membuat satu file Excel dengan satu sheet untuk setiap dusun yang memiliki kepala keluarga. Hanya daftar kepala keluarga yang diekspor; data penduduk lengkap tidak disertakan.</p>
    `;
    report.parentElement.insertBefore(wrapper, report);
    document.getElementById('exportAllDusun').addEventListener('click', exportAllDusun);

    const style = document.createElement('style');
    style.textContent = `
      #kkExtraActions { margin: 12px 0 18px; }
      #kkExtraActions .kkExtraButtons { display:flex; flex-wrap:wrap; gap:8px; }
      #kkExtraActions .kkExtraHelp { margin:9px 0 0; color:#64748b; font-size:13px; line-height:1.55; }
      #kkReport .kkTitle { display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap; }
      #kkReport .kkTitle h3 { margin:0; }
      #kkReport .kkTitle h3 small { font-weight:500; font-size:12px; color:#64748b; }
      #kkReport .empty { text-align:center; color:#64748b; padding:12px; }
      @media (max-width:600px) { #kkExtraActions .kkExtraButtons button { flex:1 1 100%; } }
    `;
    document.head.appendChild(style);
  }

  addControlPanel();

  // ---------------------------------------------------------------------------
  // Tab navigation: only one feature pane is visible at a time.
  // ---------------------------------------------------------------------------
  function setupFeatureTabs() {
    const nav = document.getElementById('dgTabNav');
    if (!nav) return;
    const buttons = Array.from(nav.querySelectorAll('[data-tab]'));
    const panes = Array.from(document.querySelectorAll('[data-pane]'));

    function activate(name, persist = true) {
      const exists = panes.some(pane => pane.dataset.pane === name);
      if (!exists) name = 'overview';
      buttons.forEach(button => {
        const active = button.dataset.tab === name;
        button.setAttribute('aria-selected', String(active));
        button.tabIndex = active ? 0 : -1;
      });
      panes.forEach(pane => { pane.hidden = pane.dataset.pane !== name; });
      if (persist) {
        try { localStorage.setItem('datagampong.activeTab.v1', name); } catch (_) {}
      }
    }

    buttons.forEach((button, index) => {
      button.addEventListener('click', () => activate(button.dataset.tab));
      button.addEventListener('keydown', event => {
        if (!['ArrowRight', 'ArrowLeft', 'Home', 'End'].includes(event.key)) return;
        event.preventDefault();
        let next = index;
        if (event.key === 'ArrowRight') next = (index + 1) % buttons.length;
        if (event.key === 'ArrowLeft') next = (index - 1 + buttons.length) % buttons.length;
        if (event.key === 'Home') next = 0;
        if (event.key === 'End') next = buttons.length - 1;
        buttons[next].focus();
        activate(buttons[next].dataset.tab);
      });
    });

    let saved = 'overview';
    try { saved = localStorage.getItem('datagampong.activeTab.v1') || 'overview'; } catch (_) {}
    activate(saved, false);
  }


  // Semua tombol "Lihat Data" dari rekap harus membuka tab Data Penduduk lebih dulu.
  // Wrapper ini mempertahankan filter/predicate asli dari tiap jenis rekap.
  const originalApplyReportFilter = applyReportFilter;
  applyReportFilter = function (label, predicate) {
    const dataTab = document.querySelector('[data-tab="data"]');
    if (dataTab) dataTab.click();
    originalApplyReportFilter(label, predicate);
    const dataPanel = document.getElementById('dataPanel');
    if (dataPanel) dataPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  };

  function setupRecapTabs() {
    const tabs = Array.from(document.querySelectorAll('[data-recap-tab]'));
    const panels = Array.from(document.querySelectorAll('[data-recap-view]'));
    if (!tabs.length || !panels.length) return;
    function activate(name) {
      tabs.forEach(tab => tab.setAttribute('aria-selected', String(tab.dataset.recapTab === name)));
      panels.forEach(panel => { panel.hidden = panel.dataset.recapView !== name; });
      try { localStorage.setItem('datagampong.activeRecapTab.v1', name); } catch (_) {}
    }
    tabs.forEach((tab, index) => {
      tab.addEventListener('click', () => activate(tab.dataset.recapTab));
      tab.addEventListener('keydown', event => {
        if (!['ArrowRight', 'ArrowLeft', 'Home', 'End'].includes(event.key)) return;
        event.preventDefault();
        let next = index;
        if (event.key === 'ArrowRight') next = (index + 1) % tabs.length;
        if (event.key === 'ArrowLeft') next = (index - 1 + tabs.length) % tabs.length;
        if (event.key === 'Home') next = 0;
        if (event.key === 'End') next = tabs.length - 1;
        tabs[next].focus(); activate(tabs[next].dataset.recapTab);
      });
    });
    let saved = 'dusun';
    try { saved = localStorage.getItem('datagampong.activeRecapTab.v1') || 'dusun'; } catch (_) {}
    activate(panels.some(panel => panel.dataset.recapView === saved) ? saved : 'dusun');
  }

  function updateOverview() {
    const bars = document.getElementById('dgOverviewDusunBars');
    const complete = document.getElementById('dgOverviewCompleteness');
    if (!bars || !complete) return;
    const hasData = Array.isArray(rows) && rows.length > 0;
    const setText = (id, value) => { const el = document.getElementById(id); if (el) el.textContent = String(value); };
    if (!hasData) {
      setText('dgOverviewDusunCount', '—'); setText('dgOverviewEducationCount', '—'); setText('dgOverviewHeadCount', '—');
      bars.innerHTML = '<div class="dg-empty-state">Pilih file Excel untuk melihat jumlah penduduk per dusun.</div>';
      complete.innerHTML = '<div class="dg-empty-state">Ringkasan kelengkapan akan muncul setelah data dimuat.</div>';
      return;
    }
    const dusunCounts = new Map();
    let educated = 0, heads = 0;
    let nik = 0, noKK = 0, jk = 0, lorong = 0, pekerjaan = 0;
    rows.forEach(row => {
      if (!isResidentRecord(row)) return;
      const name = clean(val(row, 'lorong'));
      if (name && !/^\d+$/.test(name)) {
        const normalized = normalizeDusunValue(name);
        const existing = dusunCounts.get(normalized);
        if (existing) existing.count++;
        else dusunCounts.set(normalized, { name, count: 1 });
        lorong++;
      }
      if (clean(val(row, 'pendidikan'))) educated++;
      if (isHead(row)) heads++;
      if (identityNormalizeNik(val(row, 'nik')).length === 16) nik++;
      if (clean(val(row, 'no_kk'))) noKK++;
      if (clean(val(row, 'jk'))) jk++;
      if (clean(val(row, 'pekerjaan'))) pekerjaan++;
    });
    const entries = Array.from(dusunCounts.values()).sort((a, b) => b.count - a.count || a.name.localeCompare(b.name, 'id'));
    setText('dgOverviewDusunCount', entries.length.toLocaleString('id-ID'));
    setText('dgOverviewEducationCount', educated.toLocaleString('id-ID'));
    setText('dgOverviewHeadCount', heads.toLocaleString('id-ID'));
    const max = entries.length ? entries[0].count : 0;
    bars.innerHTML = entries.length ? entries.slice(0, 5).map(item => {
      const width = max ? Math.max(4, item.count / max * 100) : 0;
      return `<div class="dg-bar-row"><span class="dg-bar-name" title="${identityEscape(item.name)}">${identityEscape(item.name)}</span><div class="dg-bar-track"><div class="dg-bar-fill" style="width:${width}%"></div></div><strong class="dg-bar-value">${item.count.toLocaleString('id-ID')}</strong></div>`;
    }).join('') : '<div class="dg-empty-state">Belum ada nama dusun dalam kolom DUSUN/LORONG.</div>';
    const completenessItems = [
      ['NIK 16 digit', nik], ['Nomor KK', noKK], ['Jenis kelamin', jk], ['Dusun/Lorong', lorong], ['Pekerjaan', pekerjaan], ['Pendidikan', educated]
    ];
    complete.innerHTML = completenessItems.map(([label, count]) => {
      const pct = Math.round(count / rows.length * 100);
      return `<div class="dg-completeness-row"><span>${label}</span><strong>${count.toLocaleString('id-ID')} / ${rows.length.toLocaleString('id-ID')} (${pct}%)</strong><div class="dg-completeness-track"><i style="width:${pct}%"></i></div></div>`;
    }).join('');
  }

  function setupQuickActions() {
    document.querySelectorAll('[data-go-tab]').forEach(button => button.addEventListener('click', () => {
      const tab = document.querySelector(`[data-tab="${button.dataset.goTab}"]`);
      if (tab) tab.click();
    }));
  }

  // ---------------------------------------------------------------------------
  // NIK lookup and KTP-inspired local preview. No photograph and not an official ID.
  // The form requires every field before a preview can be rendered.
  // ---------------------------------------------------------------------------
  const IDENTITY_STORAGE_KEY = 'datagampong.identitySupplement.v1';
  let selectedIdentityRow = null;

  const identityFieldDefs = [
    { key: 'nik', label: 'NIK', aliases: ['nik'], readonly: true },
    { key: 'nama', label: 'Nama', aliases: ['nama'] },
    { key: 'tmpt_lhr', label: 'Tempat Lahir', aliases: ['tmpt_lhr', 'tempat_lahir'] },
    { key: 'tgl_lhr', label: 'Tanggal Lahir', aliases: ['tgl_lhr', 'tanggal_lahir'] },
    { key: 'jk', label: 'Jenis Kelamin', aliases: ['jk'], type: 'select', options: ['LAKI-LAKI', 'PEREMPUAN'] },
    { key: 'gdr', label: 'Golongan Darah', aliases: ['gdr', 'golongan_darah'] },
    { key: 'alamat', label: 'Alamat', aliases: ['alamat', 'alamat_lengkap'] },
    { key: 'rt_rw', label: 'RT/RW', aliases: ['rt_rw', 'rtrw'], defaultValue: '-/-', readonly: true },
    { key: 'kel_desa', label: 'Kel/Desa atau Gampong', aliases: ['kel_desa', 'kelurahan', 'desa', 'gampong'] },
    { key: 'kecamatan', label: 'Kecamatan', aliases: ['kecamatan'] },
    { key: 'kabupaten_kota', label: 'Kabupaten/Kota', aliases: ['kabupaten_kota', 'kabupaten', 'kota', 'kab_kota'] },
    { key: 'provinsi', label: 'Provinsi', aliases: ['provinsi', 'province'] },
    { key: 'agama', label: 'Agama', aliases: ['agama'] },
    { key: 'status_perkawinan', label: 'Status Perkawinan', aliases: ['status_perkawinan', 'status'] },
    { key: 'pekerjaan', label: 'Pekerjaan', aliases: ['pekerjaan'] },
    { key: 'kewarganegaraan', label: 'Kewarganegaraan', aliases: ['kewarganegaraan'], defaultValue: 'WNI' },
    { key: 'berlaku_hingga', label: 'Berlaku Hingga', aliases: ['berlaku_hingga'], defaultValue: 'SEUMUR HIDUP' }
  ];

  function identityEscape(value) {
    return String(value == null ? '' : value).replace(/[&<>"']/g, ch => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[ch]));
  }

  function identityStorageRead() {
    try {
      const data = JSON.parse(localStorage.getItem(IDENTITY_STORAGE_KEY) || '{}');
      return data && typeof data === 'object' ? data : {};
    } catch (_) { return {}; }
  }

  function identityStorageSave(data) {
    try {
      localStorage.setItem(IDENTITY_STORAGE_KEY, JSON.stringify(data));
      return true;
    } catch (_) {
      const status = document.getElementById('dgIdentityStatus');
      if (status) {
        status.className = 'dg-identity-status error';
        status.textContent = 'Browser tidak mengizinkan penyimpanan lokal. Data kartu tetap dapat ditampilkan saat sesi ini terbuka, tetapi pelengkap tidak disimpan.';
      }
      return false;
    }
  }

  function identityReadRow(row, aliases) {
    for (const key of aliases) {
      try {
        const value = clean(val(row, key));
        if (value) return value;
      } catch (_) {}
    }
    return '';
  }

  function identityNormalizeNik(value) {
    return String(value == null ? '' : value).replace(/\D/g, '');
  }

  function identityNormalizeGender(value) {
    const valueNorm = norm(value);
    if (['L', 'LAKI-LAKI', 'LAKI LAKI', 'PRIA'].includes(valueNorm)) return 'LAKI-LAKI';
    if (['P', 'PEREMPUAN', 'WANITA'].includes(valueNorm)) return 'PEREMPUAN';
    return clean(value);
  }

  function identityFormatDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return `${String(value.getDate()).padStart(2, '0')}-${String(value.getMonth() + 1).padStart(2, '0')}-${value.getFullYear()}`;
    }
    const text = clean(value);
    if (!text) return '';
    let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:T.*)?$/);
    if (match) return `${match[3].padStart(2, '0')}-${match[2].padStart(2, '0')}-${match[1]}`;
    match = text.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})$/);
    if (match) return `${match[1].padStart(2, '0')}-${match[2].padStart(2, '0')}-${match[3]}`;
    return text;
  }

  function identitySetStatus(message, kind = '') {
    const status = document.getElementById('dgIdentityStatus');
    if (!status) return;
    status.className = 'dg-identity-status' + (kind ? ` ${kind}` : '');
    status.textContent = message;
  }

  function identityBuildForm(row) {
    selectedIdentityRow = row;
    const form = document.getElementById('dgIdentityForm');
    const fields = document.getElementById('dgIdentityFields');
    const nik = identityNormalizeNik(val(row, 'nik'));
    const savedAll = identityStorageRead();
    const saved = savedAll[nik] && typeof savedAll[nik] === 'object' ? savedAll[nik] : {};

    fields.innerHTML = identityFieldDefs.map(def => {
      let value = saved[def.key] || identityReadRow(row, def.aliases);
      if (def.key === 'nik') value = nik;
      if (def.key === 'rt_rw') value = '-/-';
      if (def.key === 'jk') value = identityNormalizeGender(value);
      if (def.key === 'tgl_lhr') value = identityFormatDate(value);
      if (!value && def.defaultValue) value = def.defaultValue;
      const required = 'required';
      const inputId = `dgIdentityField-${def.key}`;
      const control = def.type === 'select'
        ? `<select id="${inputId}" name="${def.key}" ${required}><option value="">Pilih jenis kelamin</option>${def.options.map(option => `<option value="${identityEscape(option)}" ${value === option ? 'selected' : ''}>${identityEscape(option)}</option>`).join('')}</select>`
        : `<input id="${inputId}" name="${def.key}" type="text" value="${identityEscape(value)}" ${def.readonly ? 'readonly' : ''} ${required} autocomplete="off">`;
      return `<div class="dg-identity-field"><label for="${inputId}">${identityEscape(def.label)} <span aria-hidden="true">*</span></label>${control}</div>`;
    }).join('');

    form.hidden = false;
    document.getElementById('dgIdentityCardResult').innerHTML = '';
    identitySetStatus(`Data ditemukan untuk NIK ${nik}. Periksa dan lengkapi setiap kolom bertanda * sebelum menampilkan kartu.`, 'success');
  }

  function identityChooseRow(row) {
    const nik = identityNormalizeNik(val(row, 'nik'));
    if (!/^\d{16}$/.test(nik)) {
      identitySetStatus(`Data yang dipilih memiliki NIK ${nik.length} digit. NIK harus terdiri dari 16 digit. Periksa format NIK di Excel.`, 'error');
      return;
    }
    identityBuildForm(row);
    document.getElementById('dgIdentityResults').innerHTML = '';
  }

  function identitySearch() {
    const input = document.getElementById('dgIdentityNik');
    const query = identityNormalizeNik(input.value);
    const results = document.getElementById('dgIdentityResults');
    const form = document.getElementById('dgIdentityForm');
    const card = document.getElementById('dgIdentityCardResult');
    results.innerHTML = '';
    form.hidden = true;
    card.innerHTML = '';
    selectedIdentityRow = null;

    if (!Array.isArray(rows) || rows.length === 0) {
      identitySetStatus('Muat file DATA-GAMPONG.xlsx terlebih dahulu.', 'error');
      return;
    }
    if (!query) {
      identitySetStatus('Masukkan NIK atau sebagian angka NIK untuk memulai pencarian.', 'error');
      input.focus();
      return;
    }

    const matches = rows.filter(row => identityNormalizeNik(val(row, 'nik')).includes(query));
    if (!matches.length) {
      identitySetStatus('NIK tidak ditemukan. Periksa kembali angka NIK atau muat file Excel yang benar.', 'error');
      return;
    }

    if (matches.length === 1 && identityNormalizeNik(val(matches[0], 'nik')) === query) {
      identityChooseRow(matches[0]);
      return;
    }

    if (matches.length === 1) {
      identitySetStatus('Satu data cocok dengan pencarian. Pilih data di bawah untuk memeriksa kelengkapannya.', 'success');
    } else {
      identitySetStatus(`Ditemukan ${matches.length.toLocaleString('id-ID')} data yang cocok. Pilih nama dan NIK yang benar.`, 'success');
    }
    results.innerHTML = matches.slice(0, 50).map((row, index) => {
      const nik = identityNormalizeNik(val(row, 'nik'));
      const name = clean(val(row, 'nama')) || '(Nama belum diisi)';
      const dusun = clean(val(row, 'lorong'));
      return `<button type="button" class="dg-candidate" data-identity-index="${index}"><span><strong>${identityEscape(nik)}</strong><br>${identityEscape(name)}${dusun ? `<br><small>${identityEscape(dusun)}</small>` : ''}</span><span aria-hidden="true">Pilih →</span></button>`;
    }).join('');
    const visibleMatches = matches.slice(0, 50);
    results.querySelectorAll('[data-identity-index]').forEach(button => {
      button.addEventListener('click', () => identityChooseRow(visibleMatches[Number(button.dataset.identityIndex)]));
    });
    if (matches.length > 50) {
      results.insertAdjacentHTML('beforeend', '<p class="dg-identity-help">Hanya 50 hasil pertama ditampilkan. Tambahkan angka NIK agar hasil lebih spesifik.</p>');
    }
  }

  function identityCollectForm() {
    const values = {};
    const missing = [];
    identityFieldDefs.forEach(def => {
      const control = document.querySelector(`#dgIdentityFields [name="${def.key}"]`);
      const value = clean(control ? control.value : '');
      values[def.key] = value;
      if (!value) missing.push(def.label);
    });
    if (!/^\d{16}$/.test(identityNormalizeNik(values.nik))) missing.unshift('NIK 16 digit');
    return { values, missing };
  }

  function identityCardRow(label, value, extraClass = '') {
    return `<div class="dg-id-row ${extraClass}"><span class="key">${identityEscape(label)}</span><span>:</span><span class="value">${identityEscape(value)}</span></div>`;
  }

  function identityRenderCard(data) {
    const top = `
      <div class="dg-identity-top">
        <div class="province">PROVINSI ${identityEscape(data.provinsi)}</div>
        <div class="city">${identityEscape(data.kabupaten_kota)}</div>
        <div class="dg-id-banner">PRATINJAU DATA PENDUDUK • BUKAN DOKUMEN RESMI</div>
      </div>`;
    const leftBody = [
      identityCardRow('NIK', data.nik, 'dg-id-nik'),
      identityCardRow('Nama', data.nama),
      identityCardRow('Tempat/Tgl Lahir', `${data.tmpt_lhr} / ${data.tgl_lhr}`),
      `<div class="dg-id-grid">${identityCardRow('Jenis Kelamin', data.jk)}${identityCardRow('Gol. Darah', data.gdr)}</div>`,
      identityCardRow('Alamat', data.alamat)
    ].join('');
    const rightBody = [
      `<div class="dg-id-grid">${identityCardRow('RT/RW', '-/-')}${identityCardRow('Kel/Desa', data.kel_desa)}</div>`,
      `<div class="dg-id-grid">${identityCardRow('Kecamatan', data.kecamatan)}${identityCardRow('Agama', data.agama)}</div>`,
      `<div class="dg-id-grid">${identityCardRow('Status', data.status_perkawinan)}${identityCardRow('Pekerjaan', data.pekerjaan)}</div>`,
      `<div class="dg-id-grid">${identityCardRow('Warga Negara', data.kewarganegaraan)}${identityCardRow('Berlaku Hingga', data.berlaku_hingga)}</div>`
    ].join('');
    const body = `<div class="dg-id-column dg-id-column-left">${leftBody}</div><div class="dg-id-column dg-id-column-right">${rightBody}</div>`;
    const card = `
      <div class="dg-identity-card" aria-label="Pratinjau kartu identitas penduduk tanpa foto">
        <div class="dg-id-watermark">PRATINJAU • BUKAN DOKUMEN RESMI</div>
        ${top}<div class="dg-identity-main">${body}</div>
        <div class="dg-identity-footer"><span>Data lokal gampong</span><span>Pratinjau internal — tidak berlaku sebagai identitas resmi</span></div>
      </div>
      <div class="dg-identity-actions"><button type="button" id="dgIdentityPrint">🖨 Cetak Pratinjau</button></div>`;
    document.getElementById('dgIdentityCardResult').innerHTML = card;
    document.getElementById('dgIdentityPrint').addEventListener('click', () => {
      document.body.classList.add('dg-print-identity');
      window.print();
      window.setTimeout(() => document.body.classList.remove('dg-print-identity'), 1200);
    });
  }

  function setupIdentityFeature() {
    const input = document.getElementById('dgIdentityNik');
    const searchButton = document.getElementById('dgIdentitySearch');
    const resetButton = document.getElementById('dgIdentityResetSearch');
    const form = document.getElementById('dgIdentityForm');
    if (!input || !searchButton || !form) return;

    searchButton.addEventListener('click', identitySearch);
    input.addEventListener('keydown', event => {
      if (event.key === 'Enter') { event.preventDefault(); identitySearch(); }
    });
    input.addEventListener('input', () => {
      // Keep the NIK filter numeric and within the 16-digit limit.
      const cleanValue = identityNormalizeNik(input.value).slice(0, 16);
      if (input.value !== cleanValue) input.value = cleanValue;
    });
    resetButton.addEventListener('click', () => {
      input.value = '';
      selectedIdentityRow = null;
      document.getElementById('dgIdentityResults').innerHTML = '';
      document.getElementById('dgIdentityCardResult').innerHTML = '';
      form.hidden = true;
      identitySetStatus('Pencarian dikosongkan. Masukkan NIK untuk memulai pencarian.');
      input.focus();
    });
    document.getElementById('dgIdentityClearCard').addEventListener('click', () => {
      document.getElementById('dgIdentityCardResult').innerHTML = '';
      identitySetStatus('Pratinjau kartu dibersihkan. Anda masih dapat mengubah data sebelum menampilkannya lagi.');
    });

    form.addEventListener('submit', event => {
      event.preventDefault();
      if (!selectedIdentityRow) {
        identitySetStatus('Cari dan pilih satu data penduduk terlebih dahulu.', 'error');
        return;
      }
      const { values, missing } = identityCollectForm();
      if (missing.length) {
        identitySetStatus(`Kartu belum dapat ditampilkan. Lengkapi kolom wajib berikut: ${[...new Set(missing)].join(', ')}.`, 'error');
        const firstMissing = identityFieldDefs.find(def => missing.includes(def.label));
        if (firstMissing) document.querySelector(`#dgIdentityFields [name="${firstMissing.key}"]`)?.focus();
        return;
      }
      const allSaved = identityStorageRead();
      allSaved[values.nik] = values;
      identityStorageSave(allSaved);
      identityRenderCard(values);
      identitySetStatus('Berhasil. Semua kolom wajib telah terisi. Pratinjau tanpa foto ditampilkan di bawah.', 'success');
    });
  }

  setupFeatureTabs();
  setupRecapTabs();
  setupQuickActions();
  setupIdentityFeature();
  const originalRender = render;
  render = function () {
    originalRender();
    updateOverview();
  };
  updateOverview();
})();
