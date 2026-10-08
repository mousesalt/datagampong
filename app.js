let rows=[],filtered=[],reportFilter=null;
const $=id=>document.getElementById(id);

const aliases={
  no:['NO','No.'],
  no_kk:['NO.KK','NO KK','NO_KK'],
  nik:['NIK'],
  nama:['NAMA'],
  jk:['JK','JENIS KELAMIN'],
  tmpt_lhr:['TMPT LHR','TEMPAT LAHIR'],
  tgl_lhr:['TGL LHR','TANGGAL LAHIR'],
  gdr:['GDR'],
  agama:['AGAMA'],
  status:['STATUS'],
  shdk:['SHDK'],
  pendidikan:['PENDIDIKAN'],
  pekerjaan:['PEKERJAAN'],
  nama_ayah:['NAMA AYAH'],
  nama_ibu:['NAMA IBU'],
  lorong:['LORONG','DUSUN'],
  status_umur:['STATUS UMUR'],
  umur:['UMUR'],
  umur_bulan:['UMUR BULAN']
};

const norm=x=>String(x??'').trim().toUpperCase().replace(/\s+/g,' ');
function key(h){
  let n=norm(h);
  for(const [k,a] of Object.entries(aliases)) if(a.includes(n)) return k;
  return n.toLowerCase().replace(/[^a-z0-9]+/g,'_');
}
function val(r,k){
  const h=Object.keys(r).find(h=>key(h)===k);
  if(h===undefined)return '';
  const v=r[h];
  return (k==='no_kk'||k==='nik') ? expandScientific(v) : v;
}
function clean(x){return String(x??'').trim();}
function expandScientific(x){
  const s=String(x??'').trim();
  const m=s.match(/^([+-]?)(\d+)(?:\.(\d+))?[eE]([+-]?\d+)$/);
  if(!m)return s;
  const sign=m[1]||'';
  const intPart=m[2], frac=m[3]||'', exp=Number(m[4]);
  const digits=(intPart+frac).replace(/^0+(?=\d)/,'');
  const pos=intPart.length+exp;
  if(pos<=0)return sign+'0.'+'0'.repeat(-pos)+digits;
  if(pos>=digits.length)return sign+digits+'0'.repeat(pos-digits.length);
  return sign+digits.slice(0,pos)+'.'+digits.slice(pos);
}
function num(x){
  const n=Number(String(x??'').replace(',','.'));
  return Number.isFinite(n)?n:null;
}
function esc(x){
  return String(x??'').replace(/[&<>"']/g,c=>({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[c]));
}
function isHead(r){
  const s=norm(val(r,'shdk')).replace(/[^A-Z0-9]+/g,'');
  return s.includes('KEPALAKELUARGA') || s==='KEPALA' || s==='KK';
}
function vals(k){
  return [...new Set(rows.map(r=>clean(val(r,k))).filter(Boolean))]
    .sort((a,b)=>a.localeCompare(b,'id'));
}
function makeSelect(id,label,k){
  const v=vals(k);
  if(!v.length)return '';
  return `<select id="${id}" data-key="${k}">
    <option value="">${label}</option>
    ${v.map(x=>`<option value="${esc(x)}">${esc(x)}</option>`).join('')}
  </select>`;
}

function buildFilters(){
  $('filters').innerHTML=[
    makeSelect('lorong','Semua Dusun/Lorong','lorong'),
    makeSelect('jk','Semua JK','jk'),
    makeSelect('status','Semua Status','status'),
    makeSelect('shdk','Semua SHDK','shdk'),
    makeSelect('pendidikan','Semua Pendidikan','pendidikan'),
    makeSelect('pekerjaan','Semua Pekerjaan','pekerjaan'),
    `<select id="umur">
      <option value="">Semua Umur</option>
      <option value="lansia">Lansia ≥65</option>
      <option value="dewasa">Dewasa 17–64</option>
      <option value="remaja">Remaja 10–16</option>
      <option value="anak">Anak &lt;10</option>
    </select>`
  ].filter(Boolean).join('');
  document.querySelectorAll('#filters select').forEach(x=>x.onchange=render);
}

function match(r){
  const q=clean($('q').value).toLowerCase();
  if(q && !['nama','nik','no_kk','nama_ayah','nama_ibu','pekerjaan']
    .some(k=>clean(val(r,k)).toLowerCase().includes(q))) return false;

  for(const k of ['lorong','jk','status','shdk','pendidikan','pekerjaan']){
    const el=$(k);
    if(el && el.value && clean(val(r,k))!==el.value) return false;
  }

  const age=num(val(r,'umur')), u=$('umur').value;
  if(age===null && u) return false;
  if(u==='lansia' && age<65) return false;
  if(u==='dewasa' && (age<17 || age>=65)) return false;
  if(u==='remaja' && (age<10 || age>=17)) return false;
  if(u==='anak' && age>=10) return false;
  return true;
}

function clearTopFilters(){
  $('q').value='';
  document.querySelectorAll('#filters select').forEach(s=>s.value='');
}

function applyReportFilter(label,predicate){
  reportFilter={label,predicate};
  clearTopFilters();
  render();
  $('dataPanel').scrollIntoView({behavior:'smooth',block:'start'});
}

function clearReportFilter(){
  reportFilter=null;
  render();
}

function reportButton(label){
  return `<button class="viewBtn" type="button" data-report="${esc(label)}">👥 Lihat Data</button>`;
}

function render(){
  filtered=rows.filter(r=>match(r) && (!reportFilter || reportFilter.predicate(r)));
  $('total').textContent=`${filtered.length.toLocaleString('id-ID')} data`;

  $('reportFilterBar').innerHTML=reportFilter
    ? `<div class="activeReport">
        <span>🔎 <b>Filter Rekap:</b> ${esc(reportFilter.label)} • ${filtered.length.toLocaleString('id-ID')} orang</span>
        <button type="button" id="clearReport">✕ Hapus Filter Rekap</button>
       </div>`
    : '';
  if($('clearReport')) $('clearReport').onclick=clearReportFilter;

  const cols=['no','no_kk','nik','nama','jk','tmpt_lhr','tgl_lhr','gdr','agama','status','shdk','pendidikan','pekerjaan','nama_ayah','nama_ibu','lorong','status_umur','umur','umur_bulan'];

  $('data').innerHTML=
    '<thead><tr>'+cols.map(x=>`<th>${x.replaceAll('_',' ').toUpperCase()}</th>`).join('')+'</tr></thead>'+
    '<tbody>'+filtered.slice(0,2000).map(r=>'<tr>'+
      cols.map(k=>`<td>${esc(val(r,k))}</td>`).join('')+
    '</tr>').join('')+'</tbody>';

  stats();
  renderRecaps();
}

function stats(){
  let total=rows.length,l=0,p=0,kk=0;
  rows.forEach(r=>{
    if(norm(val(r,'jk'))==='L')l++;
    if(norm(val(r,'jk'))==='P')p++;
    if(isHead(r))kk++;
  });
  $('stats').innerHTML=[
    ['Total Penduduk',total],
    ['Laki-laki',l],
    ['Perempuan',p],
    ['Kepala Keluarga',kk]
  ].map(x=>`<div class="card">${x[0]}<b>${x[1].toLocaleString('id-ID')}</b></div>`).join('');
}

function splitJK(list){
  const l=list.filter(r=>norm(val(r,'jk'))==='L').length;
  const p=list.filter(r=>norm(val(r,'jk'))==='P').length;
  return {l,p,total:l+p};
}

function renderRecaps(){
  renderAgeReport();
  renderDusunReport();
  renderJobReport();
  renderEducationReport();
  renderKKReport();
}

function renderAgeReport(){
  const ages=[
    ['LANSIA (>=65)',r=>{const a=num(val(r,'umur'));return a!==null&&a>=65;}],
    ['DEWASA (>=17)',r=>{const a=num(val(r,'umur'));return a!==null&&a>=17;}],
    ['REMAJA (>=10)',r=>{const a=num(val(r,'umur'));return a!==null&&a>=10;}],
    ['ANAK-ANAK (0–<17)',r=>{const a=num(val(r,'umur'));return a!==null&&a>=0&&a<17;}],
    ['17 TAHUN KE BAWAH',r=>{const a=num(val(r,'umur'));return a!==null&&a<=17;}],
    ['12 TAHUN KE BAWAH',r=>{const a=num(val(r,'umur'));return a!==null&&a<=12;}],
    ['6 TAHUN KE BAWAH',r=>{const a=num(val(r,'umur'));return a!==null&&a<=6;}],
    ['3 TAHUN KE BAWAH',r=>{const a=num(val(r,'umur'));return a!==null&&a<=3;}],
    ['11 BULAN KE BAWAH',r=>{const a=num(val(r,'umur_bulan'));return a!==null&&a<=11;}]
  ];

  $('ageReport').innerHTML=
    '<thead><tr><th>KELOMPOK UMUR</th><th>L</th><th>P</th><th>TOTAL</th><th>%</th><th>AKSI</th></tr></thead><tbody>'+
    ages.map(([label,test])=>{
      const list=rows.filter(test),s=splitJK(list),pct=rows.length?(s.total/rows.length*100):0;
      return `<tr><td><b>${label}</b></td><td>${s.l}</td><td>${s.p}</td><td><b>${s.total}</b></td><td>${pct.toFixed(1)}%</td><td>${reportButton(label)}</td></tr>`;
    }).join('')+'</tbody>';

  document.querySelectorAll('#ageReport .viewBtn').forEach(btn=>{
    const label=btn.dataset.report;
    const item=ages.find(x=>x[0]===label);
    if(item) btn.onclick=()=>applyReportFilter(label,item[1]);
  });
}

function renderDusunReport(){
  const d={};
  rows.forEach(r=>{
    const x=clean(val(r,'lorong'));
    if(!x)return;
    if(!d[x])d[x]={rows:[],l:0,p:0,kk:0};
    d[x].rows.push(r);
    if(norm(val(r,'jk'))==='L')d[x].l++;
    if(norm(val(r,'jk'))==='P')d[x].p++;
    if(isHead(r))d[x].kk++;
  });

  $('dusun').innerHTML=
    '<thead><tr><th>DUSUN/LORONG</th><th>L</th><th>P</th><th>TOTAL</th><th>KK</th><th>AKSI</th></tr></thead><tbody>'+
    Object.entries(d).sort((a,b)=>a[0].localeCompare(b[0],'id')).map(([name,v])=>{
      return `<tr><td><b>${esc(name)}</b></td><td>${v.l}</td><td>${v.p}</td><td><b>${v.rows.length}</b></td><td>${v.kk}</td><td>${reportButton(name)}</td></tr>`;
    }).join('')+'</tbody>';

  document.querySelectorAll('#dusun .viewBtn').forEach(btn=>{
    const name=btn.dataset.report;
    if(d[name]) btn.onclick=()=>applyReportFilter(`Dusun/Lorong: ${name}`,r=>clean(val(r,'lorong'))===name);
  });
}

function renderJobReport(){
  const jm={};
  rows.forEach(r=>{
    const x=clean(val(r,'pekerjaan'));
    if(!x)return;
    if(!jm[x])jm[x]=[];
    jm[x].push(r);
  });

  const entries=Object.entries(jm).sort((a,b)=>b[1].length-a[1].length||a[0].localeCompare(b[0],'id'));
  const grand=entries.reduce((n,[,list])=>n+list.length,0);

  $('job').innerHTML=
    '<thead><tr><th>PEKERJAAN</th><th>L</th><th>P</th><th>TOTAL</th><th>%</th><th>AKSI</th></tr></thead><tbody>'+
    entries.map(([name,list])=>{
      const s=splitJK(list),pct=grand?(s.total/grand*100):0;
      return `<tr><td><b>${esc(name)}</b></td><td>${s.l}</td><td>${s.p}</td><td><b>${s.total}</b></td><td>${pct.toFixed(1)}%</td><td>${reportButton(name)}</td></tr>`;
    }).join('')+'</tbody>';

  document.querySelectorAll('#job .viewBtn').forEach(btn=>{
    const name=btn.dataset.report;
    if(jm[name]) btn.onclick=()=>applyReportFilter(`Pekerjaan: ${name}`,r=>clean(val(r,'pekerjaan'))===name);
  });
}

function renderEducationReport(){
  const ed={};
  rows.forEach(r=>{
    const x=clean(val(r,'pendidikan'));
    if(!x)return;
    if(!ed[x])ed[x]=[];
    ed[x].push(r);
  });

  const entries=Object.entries(ed).sort((a,b)=>b[1].length-a[1].length||a[0].localeCompare(b[0],'id'));
  const grand=entries.reduce((n,[,list])=>n+list.length,0);

  $('eduTotal').textContent=`${grand.toLocaleString('id-ID')} penduduk berisi data pendidikan`;
  $('eduReport').innerHTML=
    '<thead><tr><th>PENDIDIKAN</th><th>L</th><th>P</th><th>TOTAL</th><th>%</th><th>AKSI</th></tr></thead><tbody>'+
    entries.map(([name,list])=>{
      const s=splitJK(list),pct=grand?(s.total/grand*100):0;
      return `<tr><td><b>${esc(name)}</b></td><td>${s.l}</td><td>${s.p}</td><td><b>${s.total}</b></td><td>${pct.toFixed(1)}%</td><td>${reportButton(name)}</td></tr>`;
    }).join('')+
    `<tr class="totalRow"><th>TOTAL DATA PENDIDIKAN</th><th>${entries.reduce((a,[,v])=>a+splitJK(v).l,0)}</th><th>${entries.reduce((a,[,v])=>a+splitJK(v).p,0)}</th><th>${grand}</th><th>100%</th><th></th></tr></tbody>`;

  document.querySelectorAll('#eduReport .viewBtn').forEach(btn=>{
    const name=btn.dataset.report;
    if(ed[name]) btn.onclick=()=>applyReportFilter(`Pendidikan: ${name}`,r=>clean(val(r,'pendidikan'))===name);
  });
}

function renderKKReport(){
  const groups={};
  rows.forEach(r=>{
    if(!isHead(r))return;
    const dus=clean(val(r,'lorong'));
    if(!dus)return;
    if(!groups[dus])groups[dus]=[];
    groups[dus].push(r);
  });

  const dusuns=Object.keys(groups).sort((a,b)=>a.localeCompare(b,'id'));
  const totalKK=dusuns.reduce((n,k)=>n+groups[k].length,0);
  $('kkTotal').textContent=`${totalKK.toLocaleString('id-ID')} kepala keluarga`;

  $('kkReport').innerHTML=dusuns.map(dus=>{
    const list=groups[dus].sort((a,b)=>clean(val(a,'nama')).localeCompare(clean(val(b,'nama')),'id'));
    const safeDus=esc(dus).replaceAll('"','&quot;');
    return `<div class="kkGroup">
      <div class="kkTitle"><h3>${esc(dus)}</h3><button class="excelBtn" data-dusun="${safeDus}" type="button">📊 Excel</button></div>
      <div class="table"><table><thead><tr><th>NO</th><th>NAMA</th><th>JENIS KELAMIN</th><th>PEKERJAAN</th><th>TANDA TANGAN</th></tr></thead><tbody>`+
      list.map((r,i)=>`<tr><td>${i+1}</td><td>${esc(val(r,'nama'))}</td><td>${esc(val(r,'jk'))}</td><td>${esc(val(r,'pekerjaan'))}</td><td class="signature">&nbsp;</td></tr>`).join('')+
      `</tbody></table></div></div>`;
  }).join('')||'<div class="empty">Tidak ada kepala keluarga dengan dusun terisi.</div>';

  document.querySelectorAll('.excelBtn').forEach(btn=>btn.onclick=()=>excelDusun(btn.dataset.dusun));
}

function excelDusun(dusun){
  const list=rows.filter(r=>isHead(r) && clean(val(r,'lorong'))===dusun)
    .sort((a,b)=>clean(val(a,'nama')).localeCompare(clean(val(b,'nama')),'id'));
  if(!list.length){alert('Tidak ada kepala keluarga untuk dusun ini.');return;}

  const data=[
    ['DAFTAR KEPALA KELUARGA PER DUSUN'],
    ['DUSUN / LORONG',dusun],
    [],
    ['NO','NAMA','JENIS KELAMIN','PEKERJAAN','TANDA TANGAN'],
    ...list.map((r,i)=>[i+1,clean(val(r,'nama')),clean(val(r,'jk')),clean(val(r,'pekerjaan')),''])
  ];
  const ws=XLSX.utils.aoa_to_sheet(data);
  ws['!merges']=[
    {s:{r:0,c:0},e:{r:0,c:4}},
    {s:{r:1,c:1},e:{r:1,c:4}}
  ];
  ws['!cols']=[{wch:7},{wch:32},{wch:18},{wch:30},{wch:32}];
  ws['!rows']=[];
  ws['!rows'][0]={hpt:24};
  ws['!rows'][1]={hpt:20};
  ws['!rows'][3]={hpt:24};
  for(let r=4;r<data.length;r++)ws['!rows'][r]={hpt:42};

  ['A1','A2'].forEach(addr=>{
    if(ws[addr])ws[addr].s={font:{bold:true},alignment:{horizontal:'center',vertical:'center'}};
  });
  ['A4','B4','C4','D4','E4'].forEach(addr=>{
    if(ws[addr])ws[addr].s={font:{bold:true},alignment:{horizontal:'center',vertical:'center',wrapText:true}};
  });

  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'KEPALA KELUARGA');
  const filename='Daftar-Kepala-Keluarga-'+dusun.replace(/[\\/:*?"<>|]+/g,'-').replace(/\s+/g,'-')+'.xlsx';
  XLSX.writeFile(wb,filename);
}

function csv(){
  const cols=['no','no_kk','nik','nama','jk','tmpt_lhr','tgl_lhr','gdr','agama','status','shdk','pendidikan','pekerjaan','nama_ayah','nama_ibu','lorong','status_umur','umur','umur_bulan'];
  const data=[cols.map(x=>x.toUpperCase()),...filtered.map(r=>cols.map(k=>String(val(r,k)??'')))];
  const ws=XLSX.utils.aoa_to_sheet(data);

  // Pastikan NO.KK dan NIK selalu disimpan sebagai TEXT, bukan NUMBER.
  const idCols=[1,2];
  for(let row=1;row<data.length;row++){
    for(const c of idCols){
      const addr=XLSX.utils.encode_cell({r:row,c});
      if(ws[addr]){
        ws[addr].t='s';
        ws[addr].z='@';
        ws[addr].v=String(data[row][c]??'');
      }
    }
  }

  ws['!cols']=[
    {wch:6},{wch:20},{wch:20},{wch:28},{wch:5},{wch:18},{wch:15},{wch:8},{wch:16},
    {wch:18},{wch:22},{wch:24},{wch:28},{wch:28},{wch:28},{wch:18},{wch:18},{wch:10},{wch:12}
  ];

  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'DATA PENDUDUK');
  XLSX.writeFile(wb,'data-penduduk-hasil-filter.xlsx');
}

function loadExcelFile(f){
  if(!f)return;
  const rd=new FileReader();
  rd.onload=ev=>{
    try{
      const wb=XLSX.read(new Uint8Array(ev.target.result),{type:'array',cellDates:true});
      const sh=wb.Sheets['DATAPENDUDUK'];
      if(!sh){alert('Sheet DATAPENDUDUK tidak ditemukan.');return;}
      rows=XLSX.utils.sheet_to_json(sh,{defval:'',raw:false});
      rows=rows.filter(r=>Object.values(r).some(v=>clean(v)!==''));
      reportFilter=null;
      buildFilters();
      render();
      const now=new Date().toLocaleString('id-ID');
      $('info').innerHTML=`<b>${esc(f.name)}</b> • Sheet <b>DATAPENDUDUK</b> • Total <b>${rows.length.toLocaleString('id-ID')} penduduk</b> • Dimuat ${now}`;
    }catch(err){
      alert('Gagal membaca Excel: '+err.message);
    }
  };
  rd.readAsArrayBuffer(f);
}

$('file').onchange=e=>loadExcelFile(e.target.files[0]);

$('reloadExcel').onclick=()=>{
  $('file').value='';
  $('file').click();
};

$('q').oninput=render;
$('reset').onclick=()=>{
  document.querySelectorAll('#filters select').forEach(s=>s.value='');
  $('q').value='';
  reportFilter=null;
  render();
};
$('csv').onclick=csv;
$('print').onclick=()=>window.print();
