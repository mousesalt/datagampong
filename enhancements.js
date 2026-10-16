/*
 * Data Gampong - Fitur tambahan
 * - Ekspor semua dusun dalam satu workbook Excel
 * - Setiap dusun mempunyai sheet sendiri
 * - Menambah nama dusun baru dan menyimpannya di browser
 * - Menyertakan MASTER_DUSUN agar daftar dusun ikut saat workbook dipindah perangkat
 *
 * Letakkan file ini di folder yang sama dengan index.html dan app.js.
 * Muat setelah app.js: <script src="enhancements.js"></script>
 */
(function () {
  'use strict';

  const STORAGE_KEY = 'datagampong.masterDusun.v1';

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
    const fromRows = rows.map(row => clean(val(row, 'lorong')));
    return uniqueNames([...fromRows, ...readSavedDusun()]);
  }

  // Baca sheet MASTER_DUSUN jika workbook sebelumnya pernah diekspor oleh fitur ini.
  // File Excel biasa tetap dapat dibuka seperti sebelumnya.
  const originalLoadExcelFile = loadExcelFile;
  loadExcelFile = function (file) {
    if (!file) return originalLoadExcelFile(file);

    const reader = new FileReader();
    reader.onload = function () {
      try {
        const workbook = XLSX.read(new Uint8Array(reader.result), { type: 'array' });
        const masterSheet = workbook.Sheets.MASTER_DUSUN;
        if (masterSheet) {
          const values = XLSX.utils.sheet_to_json(masterSheet, { header: 1, defval: '' });
          const names = values.slice(1).map(row => row[0]);
          saveDusun([...readSavedDusun(), ...names]);
        }
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

    rows.forEach(row => {
      const name = clean(val(row, 'lorong'));
      if (!name) return;
      if (!statsByDusun[name]) statsByDusun[name] = { rows: [], l: 0, p: 0, kk: 0 };
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
        button.onclick = () => applyReportFilter(`Dusun/Lorong: ${name}`, row => clean(val(row, 'lorong')) === name);
      }
    });
  };

  function headsForDusun(name) {
    return rows
      .filter(row => isHead(row) && clean(val(row, 'lorong')) === name)
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
    alert(`Dusun "${name}" ditambahkan. Nama akan ikut dalam Excel saat Anda memilih “Export Semua Dusun”.\n\nUntuk memasukkan warga ke dusun tersebut, ubah kolom LORONG/DUSUN pada sheet DATAPENDUDUK, lalu unggah kembali file Excel itu.`);
  }

  function addControlPanel() {
    const report = document.getElementById('kkReport');
    if (!report || document.getElementById('kkExtraActions')) return;

    const wrapper = document.createElement('div');
    wrapper.id = 'kkExtraActions';
    wrapper.innerHTML = `
      <div class="kkExtraButtons">
        <button type="button" id="exportAllDusun" class="excelBtn">📥 Export Semua Dusun</button>
        <button type="button" id="addDusun" class="excelBtn">➕ Tambah Dusun</button>
      </div>
      <p class="kkExtraHelp">Ekspor menghasilkan satu file Excel berisi data sumber, master dusun, rekap KK, daftar gabungan, dan satu sheet untuk setiap dusun. Nama dusun baru tersimpan di browser ini serta ikut terbawa dalam file hasil ekspor.</p>
    `;
    report.parentElement.insertBefore(wrapper, report);
    document.getElementById('exportAllDusun').addEventListener('click', exportAllDusun);
    document.getElementById('addDusun').addEventListener('click', addDusun);

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
})();
