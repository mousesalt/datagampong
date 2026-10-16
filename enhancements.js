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

  // Kenali berbagai judul kolom Excel untuk nama dusun. Banyak file menggunakan
  // NAMA DUSUN atau NAMA LORONG, bukan hanya DUSUN/LORONG.
  function normalizeHeaderName(value) {
    return String(value == null ? '' : value)
      .toLocaleUpperCase('id-ID')
      .replace(/[^A-Z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function isDusunHeader(value) {
    const header = normalizeHeaderName(value);
    const exactHeaders = new Set([
      'DUSUN', 'LORONG', 'NAMA DUSUN', 'NAMA LORONG',
      'DUSUN LORONG', 'LORONG DUSUN', 'DUSUN NAMA', 'LORONG NAMA',
      'NAMA DUSUN LORONG', 'NAMA LORONG DUSUN',
      'DUSUN LORONG NAMA', 'LORONG DUSUN NAMA'
    ]);
    return exactHeaders.has(header) || /^NAMA (DUSUN|LORONG)( |$)/.test(header);
  }

  // App utama memakai key()/val() untuk membaca isi Excel. Perluas pemetaan
  // kolom agar alias nama dusun tetap terbaca oleh filter dan laporan lama.
  const originalColumnKey = key;
  key = function (header) {
    if (isDusunHeader(header)) return 'lorong';
    return originalColumnKey(header);
  };

  // Bila Excel memiliki lebih dari satu kolom alias dusun, pakai nilai pertama
  // yang benar-benar terisi agar kolom kosong tidak menutupi nama yang tersedia.
  const originalValueGetter = val;
  val = function (row, field) {
    if (field === 'lorong' && row && typeof row === 'object') {
      const candidates = Object.keys(row).filter(header => key(header) === 'lorong');
      for (const header of candidates) {
        const value = clean(row[header]);
        if (value) return row[header];
      }
      if (candidates.length) return row[candidates[0]];
    }
    return originalValueGetter(row, field);
  };

  function normalizeDusunValue(value) {
    return String(value == null ? '' : value).trim().replace(/\s+/g, ' ').toLocaleUpperCase('id-ID');
  }

  function sameDusun(left, right) {
    return normalizeDusunValue(left) === normalizeDusunValue(right);
  }

  // Baca nama dusun dari semua sheet Excel, termasuk sheet terpisah seperti
  // MASTER_DUSUN/DAFTAR DUSUN. Dengan demikian tidak wajib ditambahkan manual.
  function collectDusunNamesFromWorkbook(workbook) {
    const collected = [];
    (workbook.SheetNames || []).forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) return;
      let matrix = [];
      try { matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }); } catch (_) { return; }
      // Cari baris judul kolom di seluruh sheet, bukan hanya baris pertama.
      // Ini mendukung sheet daftar dusun dengan judul laporan di atasnya.
      matrix.forEach((headerRow, headerIndex) => {
        headerRow.forEach((headerCell, columnIndex) => {
          if (!isDusunHeader(headerCell)) return;
          for (let rowIndex = headerIndex + 1; rowIndex < matrix.length; rowIndex++) {
            const raw = matrix[rowIndex] && matrix[rowIndex][columnIndex];
            const value = String(raw == null ? '' : raw).trim();
            if (!value || isDusunHeader(value)) continue;
            collected.push(value);
          }
        });
      });
    });
    return uniqueNames(collected);
  }

  function readSavedDusun() {
    try {
      const value = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
      return Array.isArray(value) ? value.map(x => String(x || '').trim()).filter(Boolean) : [];
    } catch (_) {
      return [];
    }
  }

  function uniqueNames(list) {
    const result = new Map();
    list.forEach(value => {
      const name = String(value == null ? '' : value).trim();
      if (!name) return;
      const normalized = name.toLocaleUpperCase('id-ID').replace(/\s+/g, ' ');
      if (!result.has(normalized)) result.set(normalized, name);
    });
    return Array.from(result.values()).sort((a, b) => a.localeCompare(b, 'id'));
  }

  function saveDusun(list) {
    const names = uniqueNames(list);
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(names));
    } catch (error) {
      alert('Nama dusun tidak dapat disimpan di browser ini. Periksa pengaturan penyimpanan browser.');
    }
    return names;
  }

  function allDusunNames() {
    // Nama dari data penduduk selalu dibaca langsung dari baris Excel aktif.
    // MASTER_DUSUN menyertakan dusun yang memang belum memiliki warga.
    const fromRows = rows.map(row => clean(val(row, 'lorong')));
    return uniqueNames([...fromRows, ...readSavedDusun()]);
  }

  // Baca daftar dusun dari workbook yang dipilih sebelum impor data utama.
  const originalLoadExcelFile = loadExcelFile;
  loadExcelFile = function (file) {
    if (!file) return originalLoadExcelFile(file);

    const reader = new FileReader();
    reader.onload = function () {
      try {
        const workbook = XLSX.read(new Uint8Array(reader.result), { type: 'array' });
        // Ganti daftar tersimpan dengan daftar yang benar-benar ditemukan pada
        // Excel yang baru dipilih. Ini menghindari dusun lama yang tidak ada lagi
        // di file tetap muncul sebagai dusun palsu berisi 0 KK.
        const names = collectDusunNamesFromWorkbook(workbook);
        saveDusun(names);
      } catch (error) {
        console.warn('Sheet MASTER_DUSUN tidak dapat dibaca; melanjutkan impor biasa.', error);
      }
      originalLoadExcelFile(file);
    };
    reader.onerror = function () { originalLoadExcelFile(file); };
    reader.readAsArrayBuffer(file);
  };

  // Perluas pilihan filter dusun agar nama dusun kosong yang baru ditambahkan juga terlihat.
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

  // Tampilkan dusun baru yang belum memiliki penduduk pada rekap dusun.
  renderDusunReport = function () {
    const statsByDusun = {};
    allDusunNames().forEach(name => {
      statsByDusun[name] = { rows: [], l: 0, p: 0, kk: 0 };
    });

    const canonicalNames = new Map(Object.keys(statsByDusun).map(name => [normalizeDusunValue(name), name]));
    rows.forEach(row => {
      const rawName = clean(val(row, 'lorong'));
      if (!rawName) return;
      const name = canonicalNames.get(normalizeDusunValue(rawName)) || rawName;
      if (!statsByDusun[name]) statsByDusun[name] = { rows: [], l: 0, p: 0, kk: 0 };
      canonicalNames.set(normalizeDusunValue(rawName), name);
      const item = statsByDusun[name];
      item.rows.push(row);
      if (norm(val(row, 'jk')) === 'L') item.l++;
      if (norm(val(row, 'jk')) === 'P') item.p++;
      if (isHead(row)) item.kk++;
    });

    const names = Object.keys(statsByDusun).sort((a, b) => a.localeCompare(b, 'id'));
    $('dusun').innerHTML =
      '<thead><tr><th>DUSUN/LORONG</th><th>L</th><th>P</th><th>TOTAL</th><th>KK</th><th>AKSI</th></tr></thead><tbody>' +
      names.map(name => {
        const item = statsByDusun[name];
        return `<tr><td><b>${esc(name)}</b></td><td>${item.l}</td><td>${item.p}</td><td><b>${item.rows.length}</b></td><td>${item.kk}</td><td>${reportButton(name)}</td></tr>`;
      }).join('') + '</tbody>';

    document.querySelectorAll('#dusun .viewBtn').forEach(button => {
      const name = button.dataset.report;
      if (statsByDusun[name]) {
        button.onclick = () => applyReportFilter(`Dusun/Lorong: ${name}`, row => sameDusun(val(row, 'lorong'), name));
      }
    });
  };

  function headsForDusun(name) {
    return rows
      .filter(row => isHead(row) && sameDusun(val(row, 'lorong'), name))
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
    const names = allDusunNames();
    $('kkTotal').textContent = `${names.reduce((total, name) => total + headsForDusun(name).length, 0).toLocaleString('id-ID')} kepala keluarga`;
    $('kkReport').innerHTML = names.map(name => {
      const count = headsForDusun(name).length;
      const safeName = esc(name).replaceAll('"', '&quot;');
      const list = headsForDusun(name);
      const table = list.length
        ? list.map((row, index) => `<tr><td>${index + 1}</td><td>${esc(val(row, 'nama'))}</td><td>${esc(val(row, 'jk'))}</td><td>${esc(val(row, 'pekerjaan'))}</td><td class="signature">&nbsp;</td></tr>`).join('')
        : '<tr><td colspan="5" class="empty">Belum ada kepala keluarga pada dusun ini.</td></tr>';
      return `<div class="kkGroup">
        <div class="kkTitle"><h3>${esc(name)} <small>(${count.toLocaleString('id-ID')} KK)</small></h3><button class="excelBtn" data-export-dusun="${safeName}" type="button">📊 Excel</button></div>
        <div class="table"><table><thead><tr><th>NO</th><th>NAMA</th><th>JENIS KELAMIN</th><th>PEKERJAAN</th><th>TANDA TANGAN</th></tr></thead><tbody>${table}</tbody></table></div>
      </div>`;
    }).join('') || '<div class="empty">Belum ada dusun. Tambahkan nama dusun atau muat data Excel.</div>';

    document.querySelectorAll('#kkReport .excelBtn').forEach(button => {
      button.onclick = () => exportDusun(button.dataset.exportDusun);
    });
  };

  function exportAllDusun() {
    if (!Array.isArray(rows) || rows.length === 0) {
      alert('Muat file DATA-GAMPONG.xlsx terlebih dahulu.');
      return;
    }

    const names = allDusunNames();
    const workbook = XLSX.utils.book_new();

    // Database sumber tetap ikut agar file ekspor dapat diunggah kembali ke aplikasi.
    const rawSheet = XLSX.utils.json_to_sheet(rows.map(row => ({ ...row })));
    rawSheet['!cols'] = Object.keys(rows[0] || {}).map(() => ({ wch: 20 }));
    appendUniqueSheet(workbook, 'DATAPENDUDUK', rawSheet);

    // Master dusun memungkinkan daftar nama dusun dibawa ke perangkat/browser lain.
    const masterSheet = XLSX.utils.aoa_to_sheet([['DUSUN'], ...names.map(name => [name])]);
    masterSheet['!cols'] = [{ wch: 32 }];
    appendUniqueSheet(workbook, 'MASTER_DUSUN', masterSheet);

    const recaps = [['REKAP JUMLAH KEPALA KELUARGA PER DUSUN'], [], ['DUSUN / LORONG', 'JUMLAH KK'], ...names.map(name => [name, headsForDusun(name).length])];
    const recapSheet = XLSX.utils.aoa_to_sheet(recaps);
    recapSheet['!cols'] = [{ wch: 34 }, { wch: 18 }];
    appendUniqueSheet(workbook, 'REKAP KK', recapSheet);

    const allHeads = [];
    names.forEach(name => {
      headsForDusun(name).forEach(row => allHeads.push({ name, row }));
    });
    const combinedData = [
      ['DAFTAR KEPALA KELUARGA - SEMUA DUSUN'],
      [],
      ['NO', 'DUSUN / LORONG', 'NAMA', 'JENIS KELAMIN', 'PEKERJAAN', 'TANDA TANGAN'],
      ...allHeads.map((item, index) => [index + 1, item.name, clean(val(item.row, 'nama')), clean(val(item.row, 'jk')), clean(val(item.row, 'pekerjaan')), ''])
    ];
    const combinedSheet = XLSX.utils.aoa_to_sheet(combinedData);
    combinedSheet['!cols'] = [{ wch: 7 }, { wch: 28 }, { wch: 32 }, { wch: 18 }, { wch: 30 }, { wch: 32 }];
    appendUniqueSheet(workbook, 'SEMUA DUSUN', combinedSheet);

    // Tiap dusun menjadi satu sheet tersendiri, termasuk dusun baru yang masih kosong.
    names.forEach(name => appendUniqueSheet(workbook, name, makeDusunSheet(name)));

    XLSX.writeFile(workbook, 'Data-Gampong-Ekspor-Semua-Dusun.xlsx');
  }

  function addDusun() {
    const input = prompt('Masukkan nama dusun/lorong baru:');
    if (input === null) return;
    const name = input.trim().replace(/\s+/g, ' ');
    if (!name) {
      alert('Nama dusun tidak boleh kosong.');
      return;
    }
    const existing = allDusunNames();
    if (existing.some(item => item.toLocaleUpperCase('id-ID') === name.toLocaleUpperCase('id-ID'))) {
      alert('Nama dusun tersebut sudah ada.');
      return;
    }
    saveDusun([...readSavedDusun(), name]);
    buildFilters();
    renderDusunReport();
    renderKKReport();
    alert(`Dusun "${name}" ditambahkan untuk penggunaan sementara di browser ini. Agar menjadi bagian dari sumber data, tambahkan namanya pada kolom DUSUN/LORONG/NAMA DUSUN di Excel, atau masukkan dalam sheet tersendiri dengan judul kolom DUSUN, lalu unggah ulang file.`);
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
      <p class="kkExtraHelp">Nama dusun dibaca otomatis dari Excel. Untuk dusun yang sudah memiliki warga, isi kolom DUSUN/LORONG/NAMA DUSUN/NAMA LORONG pada data penduduk. Untuk dusun yang belum memiliki warga, buat sheet Excel khusus dengan judul kolom DUSUN dan tuliskan satu nama dusun per baris, lalu unggah ulang file. Ekspor menghasilkan rekap, daftar gabungan, dan satu sheet untuk setiap dusun.</p>
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
    { key: 'rt_rw', label: 'RT/RW', aliases: ['rt_rw', 'rtrw'] },
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
    const body = [
      identityCardRow('NIK', data.nik, 'dg-id-nik'),
      identityCardRow('Nama', data.nama),
      identityCardRow('Tempat/Tgl Lahir', `${data.tmpt_lhr} / ${data.tgl_lhr}`),
      `<div class="dg-id-grid">${identityCardRow('Jenis Kelamin', data.jk)}${identityCardRow('Gol. Darah', data.gdr)}</div>`,
      identityCardRow('Alamat', data.alamat),
      `<div class="dg-id-grid">${identityCardRow('RT/RW', data.rt_rw)}${identityCardRow('Kel/Desa', data.kel_desa)}</div>`,
      `<div class="dg-id-grid">${identityCardRow('Kecamatan', data.kecamatan)}${identityCardRow('Agama', data.agama)}</div>`,
      `<div class="dg-id-grid">${identityCardRow('Status Perkawinan', data.status_perkawinan)}${identityCardRow('Pekerjaan', data.pekerjaan)}</div>`,
      `<div class="dg-id-grid">${identityCardRow('Kewarganegaraan', data.kewarganegaraan)}${identityCardRow('Berlaku Hingga', data.berlaku_hingga)}</div>`
    ].join('');
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
  setupIdentityFeature();
})();
