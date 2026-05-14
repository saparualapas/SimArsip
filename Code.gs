// ============================================================
//  SIMARSIP — Google Apps Script Backend  v3.0
//
//  ARSITEKTUR:
//  ┌─ Spreadsheet MASTER (spreadsheet ini) ─────────────────┐
//  │  • Config      : key | value                           │
//  │  • KodeWilayah : key | value                           │
//  │  • Users       : ID | Role | Email | Password | Nama   │
//  │  • JenisArsip  : ID | Kode | Nama | Aktif              │
//  │  • Klasifikasi : ID | Nama | Kode | Keterangan         │
//  └────────────────────────────────────────────────────────┘
//
//  ┌─ Spreadsheet DATA (terpisah, per tahun) ───────────────┐
//  │  • Sheet "Data_Surat" : No | No_Surat | Alamat |       │
//  │                         Tanggal_Surat | Perihal        │
//  └────────────────────────────────────────────────────────┘
//
//  Config keys:
//    aktifSpreadsheetId : ID spreadsheet data aktif
//    aktifSpreadsheetNama : Nama/tahun spreadsheet data aktif
//    CTR_XXXX : counter per kode klasifikasi
//
//  Cara setup awal:
//  1. Jalankan fungsi setupAwal() sekali dari Apps Script editor
//  2. Spreadsheet data 2026 akan dibuat otomatis
// ============================================================

const SS_MASTER = SpreadsheetApp.getActiveSpreadsheet();

// ════════════════════════════════════════════════════════════
//  CONFIG HELPERS
// ════════════════════════════════════════════════════════════
function getConfig(key) {
  const sh = SS_MASTER.getSheetByName('Config');
  if (!sh) return null;
  const rows = sh.getDataRange().getValues();
  const row  = rows.find(r => String(r[0]).trim() === key);
  return row ? String(row[1]).trim() : null;
}

function setConfig(key, value) {
  let sh = SS_MASTER.getSheetByName('Config');
  if (!sh) { sh = SS_MASTER.insertSheet('Config'); sh.appendRow(['Key','Value']); }
  const rows = sh.getDataRange().getValues();
  let found = false;
  rows.forEach((r, i) => {
    if (i === 0) return;
    if (String(r[0]).trim() === key) { sh.getRange(i+1, 2).setValue(value); found = true; }
  });
  if (!found) sh.appendRow([key, value]);
}

// ════════════════════════════════════════════════════════════
//  SPREADSHEET DATA (terpisah)
// ════════════════════════════════════════════════════════════

// Buka spreadsheet data aktif
function getDataSS() {
  const id = getConfig('aktifSpreadsheetId');
  if (!id) return null;
  try {
    return SpreadsheetApp.openById(id);
  } catch(e) {
    return null;
  }
}

// Buka sheet "Data_Surat" di spreadsheet data aktif
function getDataSheet(ss) {
  if (!ss) return null;
  let sh = ss.getSheetByName('Data_Surat');
  if (!sh) {
    sh = ss.insertSheet('Data_Surat');
    sh.appendRow(['No','No_Surat','Alamat_Penerima','Tanggal_Surat','Perihal']);
    sh.setColumnWidth(1, 60);
    sh.setColumnWidth(2, 230);
    sh.setColumnWidth(3, 320);
    sh.setColumnWidth(4, 130);
    sh.setColumnWidth(5, 280);
    sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
  }
  return sh;
}

// Buat spreadsheet data baru
function buatSpreadsheetData(namaLabel) {
  const judul  = 'SIMARSIP Data Surat - ' + namaLabel;
  const newSS  = SpreadsheetApp.create(judul);
  // Pindah ke folder yang sama dengan spreadsheet master
  try {
    const masterFile = DriveApp.getFileById(SS_MASTER.getId());
    const folder     = masterFile.getParents().next();
    DriveApp.getFileById(newSS.getId()).moveTo(folder);
  } catch(e) { /* tidak apa-apa jika gagal pindah folder */ }
  // Buat sheet Data_Surat
  const sh = newSS.getActiveSheet();
  sh.setName('Data_Surat');
  sh.appendRow(['No','No_Surat','Alamat_Penerima','Tanggal_Surat','Perihal']);
  sh.setColumnWidth(1, 60);
  sh.setColumnWidth(2, 230);
  sh.setColumnWidth(3, 320);
  sh.setColumnWidth(4, 130);
  sh.setColumnWidth(5, 280);
  sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
  return newSS;
}

// ════════════════════════════════════════════════════════════
//  SETUP AWAL — jalankan sekali dari editor Apps Script
// ════════════════════════════════════════════════════════════
function setupAwal() {
  const yr = new Date().getFullYear();
  const label = String(yr);

  // Cek apakah sudah ada spreadsheet aktif
  const existing = getConfig('aktifSpreadsheetId');
  if (existing) {
    try {
      const ss = SpreadsheetApp.openById(existing);
      Logger.log('Spreadsheet aktif sudah ada: ' + ss.getName());
      return 'Sudah setup, ID: ' + existing;
    } catch(e) { /* lanjut buat baru */ }
  }

  // Buat spreadsheet data baru
  const newSS = buatSpreadsheetData(label);
  setConfig('aktifSpreadsheetId',   newSS.getId());
  setConfig('aktifSpreadsheetNama', label);

  // Pastikan Config sheet ada
  if (!SS_MASTER.getSheetByName('KodeWilayah')) {
    const shW = SS_MASTER.insertSheet('KodeWilayah');
    shW.appendRow(['key','value']);
    shW.appendRow(['prefixWil','WP.28.PAS']);
    shW.appendRow(['kodeWil','8']);
  }

  Logger.log('Setup selesai. Spreadsheet data: ' + newSS.getId());
  Logger.log('URL: ' + newSS.getUrl());
  return 'OK - ' + newSS.getUrl();
}

// ════════════════════════════════════════════════════════════
//  HEADER HELPER
// ════════════════════════════════════════════════════════════
function ci(H, name) {
  const n = name.toLowerCase().replace(/[_\s]/g,'');
  let i = H.indexOf(n);
  if (i >= 0) return i;
  return H.findIndex(h => h.startsWith(n) || n.startsWith(h));
}

function hdrs(sh) {
  return sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0]
    .map(h => String(h).trim().toLowerCase().replace(/[_\s]/g,''));
}

// ════════════════════════════════════════════════════════════
//  doPost ROUTER
// ════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    let result;
    switch(body.action) {
      case 'login':              result = login(body); break;
      case 'getJenisArsip':     result = getJenisArsip(); break;
      case 'getKlasifikasi':    result = getKlasifikasi(body); break;
      case 'getKodeWilayah':    result = getKodeWilayah(); break;
      case 'submitSurat':       result = submitSurat(body); break;
      case 'getSurat':          result = getSurat(body); break;
      case 'deleteSurat':       result = deleteSurat(body); break;
      case 'updateKodeWilayah': result = updateKodeWilayah(body); break;
      case 'addJenisArsip':     result = addJenisArsip(body); break;
      case 'editJenisArsip':    result = editJenisArsip(body); break;
      case 'deleteJenisArsip':  result = deleteJenisArsip(body); break;
      case 'addKlasifikasi':    result = addKlasifikasi(body); break;
      case 'editKlasifikasi':   result = editKlasifikasi(body); break;
      case 'deleteKlasifikasi': result = deleteKlasifikasi(body); break;
      case 'getUsers':          result = getUsers(body); break;
      case 'addUser':           result = addUser(body); break;
      case 'editUser':          result = editUser(body); break;
      case 'deleteUser':        result = deleteUser(body); break;
      // Sheet management
      case 'getSheetList':      result = getSheetList(body); break;
      case 'addDataSheet':      result = addDataSheet(body); break;
      case 'setActiveSheet':    result = setActiveSheet(body); break;
      case 'deleteDataSheet':   result = deleteDataSheet(body); break;
      default: result = {success:false, message:'Action tidak dikenali: '+body.action};
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      success:false, message:'Server error: '+err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({
    status:'SIMARSIP v3.0 OK', time:new Date().toISOString(),
    aktif: getConfig('aktifSpreadsheetNama') || '-'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  AUTH
// ════════════════════════════════════════════════════════════
function login({email, password}) {
  const sh = SS_MASTER.getSheetByName('Users');
  if (!sh) return {success:false, message:'Sheet Users tidak ditemukan.'};
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return {success:false, message:'Belum ada data pengguna.'};
  const H    = rows[0].map(h => String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iId  = ci(H,'id'), iRole=ci(H,'role'), iMail=ci(H,'email'), iPw=ci(H,'password'), iNama=ci(H,'nama');
  if (iMail<0||iPw<0) return {success:false, message:'Kolom Email/Password tidak ada.'};
  const found = rows.slice(1).find(r =>
    String(r[iMail]).trim().toLowerCase() === String(email).trim().toLowerCase() &&
    String(r[iPw]).trim() === String(password).trim()
  );
  if (!found) return {success:false, message:'Email atau password salah.'};
  return {success:true, user:{
    id:    iId>=0  ? String(found[iId]).trim()  : String(email).trim(),
    email: String(found[iMail]).trim(),
    role:  iRole>=0? String(found[iRole]).trim().toLowerCase() : 'user',
    nama:  iNama>=0? String(found[iNama]).trim() : String(email),
  }};
}

// ════════════════════════════════════════════════════════════
//  KODE WILAYAH
// ════════════════════════════════════════════════════════════
function getKodeWilayah() {
  let sh = SS_MASTER.getSheetByName('KodeWilayah');
  if (!sh) return {success:true, data:{prefixWil:'WP.28.PAS', kodeWil:'8'}};
  const map = {prefixWil:'WP.28.PAS', kodeWil:'8'};
  sh.getDataRange().getValues().slice(1).forEach(r => {
    if (r[0]) map[String(r[0]).trim()] = String(r[1]).trim();
  });
  return {success:true, data:map};
}

function updateKodeWilayah({user, prefixWil, kodeWil}) {
  if (!isAdmin(user)) return noAuth();
  let sh = SS_MASTER.getSheetByName('KodeWilayah');
  if (!sh) { sh = SS_MASTER.insertSheet('KodeWilayah'); sh.appendRow(['key','value']); }
  const rows = sh.getDataRange().getValues();
  let fp=false, fk=false;
  rows.forEach((r,i) => {
    if (!i) return;
    if (String(r[0]).trim()==='prefixWil'){sh.getRange(i+1,2).setValue(prefixWil);fp=true;}
    if (String(r[0]).trim()==='kodeWil')  {sh.getRange(i+1,2).setValue(kodeWil);fk=true;}
  });
  if (!fp) sh.appendRow(['prefixWil',prefixWil]);
  if (!fk) sh.appendRow(['kodeWil',kodeWil]);
  return {success:true};
}

// ════════════════════════════════════════════════════════════
//  JENIS ARSIP (dari master)
// ════════════════════════════════════════════════════════════
function getJenisArsip() {
  const sh = SS_MASTER.getSheetByName('JenisArsip');
  if (sh && sh.getLastRow()>=2) {
    const rows = sh.getDataRange().getValues();
    const H = rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
    const data = rows.slice(1).filter(r=>r.some(c=>c!=='')).map(r=>({
      id:   ci(H,'id')>=0  ? String(r[ci(H,'id')]).trim()   : String(r[0]).trim(),
      kode: ci(H,'kode')>=0? String(r[ci(H,'kode')]).trim() : String(r[0]).trim(),
      nama: ci(H,'nama')>=0? String(r[ci(H,'nama')]).trim() : '',
    })).filter(j => {
      const iA = ci(H,'aktif');
      if (iA<0) return true;
      const v = String(rows.slice(1).find(r=>r[ci(H,'id')||0]===j.id)?.[iA]||'true').toUpperCase();
      return v!=='FALSE'&&v!=='0';
    });
    if (data.length) return {success:true, data};
  }
  return _deriveJenisArsip();
}

function _deriveJenisArsip() {
  const sh = SS_MASTER.getSheetByName('Klasifikasi');
  if (!sh||sh.getLastRow()<2) return {success:true,data:[]};
  const rows = sh.getDataRange().getValues();
  const H = rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iKode=ci(H,'kode'), iKet=ci(H,'keterangan');
  const map = {};
  rows.slice(1).filter(r=>r.some(c=>c!=='')).forEach(r=>{
    const kode = iKode>=0?String(r[iKode]).trim():'';
    if (!kode) return;
    const prefix = kode.split('.')[0].replace(/[^A-Za-z]/g,'').toUpperCase();
    if (!prefix||map[prefix]) return;
    let nama = prefix;
    if (iKet>=0){ const ket=String(r[iKet]).trim(); const b=ket.split('–')[0].split(' - ')[0].trim(); if(b&&b.length>1&&b.length<50) nama=b; }
    map[prefix]={id:prefix,kode:prefix,nama};
  });
  return {success:true,data:Object.values(map)};
}

function addJenisArsip({user,kode,nama}) {
  if(!isAdmin(user)) return noAuth();
  let sh = SS_MASTER.getSheetByName('JenisArsip');
  if (!sh) { sh = SS_MASTER.insertSheet('JenisArsip'); sh.appendRow(['ID','Kode','Nama','Aktif']); }
  sh.appendRow([kode,kode,nama,true]);
  return {success:true};
}

function editJenisArsip({user,id,kode,nama,aktif}) {
  if(!isAdmin(user)) return noAuth();
  const sh=SS_MASTER.getSheetByName('JenisArsip'); if(!sh) return {success:false};
  const rows=sh.getDataRange().getValues(); const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  rows.forEach((r,i)=>{ if(!i) return; if(String(r[ci(H,'id')||0]).trim()===id){ if(ci(H,'kode')>=0) sh.getRange(i+1,ci(H,'kode')+1).setValue(kode); if(ci(H,'nama')>=0) sh.getRange(i+1,ci(H,'nama')+1).setValue(nama); if(ci(H,'aktif')>=0) sh.getRange(i+1,ci(H,'aktif')+1).setValue(aktif); } });
  return {success:true};
}

function deleteJenisArsip({user,id}) {
  if(!isAdmin(user)) return noAuth();
  return _deleteByField(SS_MASTER.getSheetByName('JenisArsip'),'id',id);
}

// ════════════════════════════════════════════════════════════
//  KLASIFIKASI (dari master)
// ════════════════════════════════════════════════════════════
function getKlasifikasi({jenisArsipId}) {
  const sh=SS_MASTER.getSheetByName('Klasifikasi');
  if(!sh||sh.getLastRow()<2) return {success:true,data:[]};
  const rows=sh.getDataRange().getValues();
  const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iId=ci(H,'id'),iNama=ci(H,'nama'),iKode=ci(H,'kode'),iKet=ci(H,'keterangan'),iAktif=ci(H,'aktif');
  let data=rows.slice(1).filter(r=>r.some(c=>c!=='')).map(r=>{
    const kode=iKode>=0?String(r[iKode]).trim():'';
    const prefix=kode.split('.')[0].replace(/[^A-Za-z]/g,'').toUpperCase();
    return {id:iId>=0?String(r[iId]).trim():kode, nama:iNama>=0?String(r[iNama]).trim():'', kode, keterangan:iKet>=0?String(r[iKet]).trim():'', jenisArsipId:prefix, aktif:iAktif>=0?r[iAktif]:true};
  }).filter(k=>{const v=String(k.aktif).toUpperCase();return v!=='FALSE'&&v!=='0'&&v!=='TIDAK';});
  if(jenisArsipId) data=data.filter(k=>k.jenisArsipId===jenisArsipId);
  return {success:true,data};
}

function addKlasifikasi({user,kode,nama,keterangan}) {
  if(!isAdmin(user)) return noAuth();
  const sh=SS_MASTER.getSheetByName('Klasifikasi'); if(!sh) return {success:false,message:'Sheet Klasifikasi tidak ditemukan.'};
  sh.appendRow(['KLS_'+kode.replace(/\./g,''), nama||kode, kode, keterangan||'', true]);
  return {success:true};
}

function editKlasifikasi({user,id,kode,nama,keterangan,aktif}) {
  if(!isAdmin(user)) return noAuth();
  const sh=SS_MASTER.getSheetByName('Klasifikasi'); if(!sh) return {success:false};
  const rows=sh.getDataRange().getValues(); const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  rows.forEach((r,i)=>{ if(!i) return; if(String(r[ci(H,'id')||0]).trim()===id){ if(ci(H,'nama')>=0) sh.getRange(i+1,ci(H,'nama')+1).setValue(nama||kode); if(ci(H,'kode')>=0) sh.getRange(i+1,ci(H,'kode')+1).setValue(kode); if(ci(H,'keterangan')>=0) sh.getRange(i+1,ci(H,'keterangan')+1).setValue(keterangan||''); if(ci(H,'aktif')>=0) sh.getRange(i+1,ci(H,'aktif')+1).setValue(aktif); } });
  return {success:true};
}

function deleteKlasifikasi({user,id}) {
  if(!isAdmin(user)) return noAuth();
  return _deleteByField(SS_MASTER.getSheetByName('Klasifikasi'),'id',id);
}

// ════════════════════════════════════════════════════════════
//  SURAT — baca/tulis ke spreadsheet data terpisah
// ════════════════════════════════════════════════════════════
function submitSurat({kodeKlasifikasi, keterangan, alamat, perihal, tglSurat}) {
  if (!kodeKlasifikasi) return {success:false, message:'Kode klasifikasi tidak dikirim.'};

  const ss = getDataSS();
  if (!ss) return {success:false, message:'Spreadsheet data belum dikonfigurasi. Hubungi admin untuk jalankan setupAwal().'};

  const sh  = getDataSheet(ss);
  const wil = getKodeWilayah().data;

  // Nomor urut = lastRow (header=1, data mulai row 2, jadi No = lastRow)
  const nomorUrut = sh.getLastRow();
  const noSurat   = wil.prefixWil + '.' + wil.kodeWil + '.' + kodeKlasifikasi + '-' + nomorUrut;

  // Tulis ke sheet data
  const H = hdrs(sh);
  const newRow = new Array(Math.max(H.length, 5)).fill('');
  const map = {
    'no':nomorUrut, 'nosurat':noSurat, 'no_surat':noSurat,
    'alamatpenerima':alamat||'', 'tanggalsurat':tglSurat||'',
    'tanggalsura':tglSurat||'', 'perihal':perihal||'',
  };
  H.forEach((h,i) => { if(map[h]!==undefined) newRow[i]=map[h]; });
  sh.appendRow(newRow);

  return {success:true, noSurat, nomorUrut};
}

function getSurat({user, limit}) {
  if (!isAdmin(user)) return noAuth();
  const ss = getDataSS();
  if (!ss) return {success:true, data:[], info:'Spreadsheet data belum dikonfigurasi.'};

  const sh = getDataSheet(ss);
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return {success:true, data:[]};

  const H = rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iNo=ci(H,'no'), iNoS=ci(H,'nosurat'), iAl=ci(H,'alamatpenerima'), iTgl=ci(H,'tanggalsurat'), iPer=ci(H,'perihal');

  let data = rows.slice(1).filter(r=>r.some(c=>c!=='')).map((r,idx)=>({
    rowIndex: idx+2,
    no:       iNo>=0 ? r[iNo] : idx+1,
    noSurat:  iNoS>=0? String(r[iNoS]).trim() : '',
    alamat:   iAl>=0 ? String(r[iAl]).trim()  : '',
    tglSurat: iTgl>=0? String(r[iTgl]).trim()  : '',
    perihal:  iPer>=0? String(r[iPer]).trim()  : '',
  }));

  data.sort((a,b) => b.rowIndex - a.rowIndex);
  if (limit) data = data.slice(0, limit);
  return {success:true, data};
}

function deleteSurat({user, rowIndex}) {
  if (!isAdmin(user)) return noAuth();
  const ss = getDataSS();
  if (!ss) return {success:false, message:'Spreadsheet data tidak ditemukan.'};
  const sh = getDataSheet(ss);
  if (!rowIndex || rowIndex < 2) return {success:false, message:'Row tidak valid.'};
  sh.deleteRow(rowIndex);
  return {success:true};
}

// ════════════════════════════════════════════════════════════
//  SHEET (SPREADSHEET) MANAGEMENT
// ════════════════════════════════════════════════════════════
function getSheetList({user}) {
  if (!isAdmin(user)) return noAuth();

  const aktifId   = getConfig('aktifSpreadsheetId')   || '';
  const aktifNama = getConfig('aktifSpreadsheetNama') || '-';

  // Cari semua spreadsheet data SIMARSIP dari Config
  // Simpan daftar sebagai JSON di Config key 'spreadsheetList'
  let list = [];
  try {
    const raw = getConfig('spreadsheetList');
    if (raw) list = JSON.parse(raw);
  } catch(e) { list = []; }

  // Pastikan spreadsheet aktif ada di list
  if (aktifId && !list.find(x => x.id === aktifId)) {
    list.push({id: aktifId, nama: aktifNama});
    setConfig('spreadsheetList', JSON.stringify(list));
  }

  // Tambah info jumlah data dan status aktif
  const result = list.map(item => {
    let rows = 0;
    try {
      const ss = SpreadsheetApp.openById(item.id);
      const sh = ss.getSheetByName('Data_Surat');
      if (sh) rows = Math.max(0, sh.getLastRow() - 1);
    } catch(e) { rows = -1; } // -1 = tidak bisa diakses
    return {
      id:    item.id,
      nama:  item.nama,
      rows,
      aktif: item.id === aktifId,
    };
  });

  return {success:true, data: result, aktifNama};
}

function addDataSheet({user, sheetName}) {
  if (!isAdmin(user)) return noAuth();
  const label = String(sheetName).trim();
  if (!label) return {success:false, message:'Nama / tahun wajib diisi.'};

  // Buat spreadsheet baru
  const newSS = buatSpreadsheetData(label);
  const newId = newSS.getId();

  // Tambah ke list
  let list = [];
  try { const raw = getConfig('spreadsheetList'); if(raw) list = JSON.parse(raw); } catch(e){}
  if (list.find(x => x.id === newId)) return {success:false, message:'Spreadsheet sudah ada.'};
  list.push({id: newId, nama: label});
  setConfig('spreadsheetList', JSON.stringify(list));

  // Aktifkan otomatis
  setConfig('aktifSpreadsheetId',   newId);
  setConfig('aktifSpreadsheetNama', label);

  return {success:true, id: newId, nama: label, url: newSS.getUrl()};
}

function setActiveSheet({user, id}) {
  if (!isAdmin(user)) return noAuth();
  let list = [];
  try { const raw = getConfig('spreadsheetList'); if(raw) list = JSON.parse(raw); } catch(e){}
  const item = list.find(x => x.id === id);
  if (!item) return {success:false, message:'Spreadsheet tidak ditemukan di daftar.'};
  setConfig('aktifSpreadsheetId',   item.id);
  setConfig('aktifSpreadsheetNama', item.nama);
  return {success:true};
}

function deleteDataSheet({user, id}) {
  if (!isAdmin(user)) return noAuth();
  let list = [];
  try { const raw = getConfig('spreadsheetList'); if(raw) list = JSON.parse(raw); } catch(e){}
  if (list.length <= 1) return {success:false, message:'Minimal harus ada 1 spreadsheet data.'};
  const idx = list.findIndex(x => x.id === id);
  if (idx < 0) return {success:false, message:'Spreadsheet tidak ditemukan.'};

  // Jika ini yang aktif, ganti ke yang lain dulu
  const aktifId = getConfig('aktifSpreadsheetId');
  if (aktifId === id) {
    const other = list.find(x => x.id !== id);
    setConfig('aktifSpreadsheetId',   other.id);
    setConfig('aktifSpreadsheetNama', other.nama);
  }

  // Hapus dari list (tidak hapus file Google Drive — data tetap aman)
  list.splice(idx, 1);
  setConfig('spreadsheetList', JSON.stringify(list));
  return {success:true, note:'Spreadsheet dihapus dari daftar. File Google Drive tidak dihapus untuk keamanan data.'};
}

// ════════════════════════════════════════════════════════════
//  USERS (dari master)
// ════════════════════════════════════════════════════════════
function getUsers({user}) {
  if(!isAdmin(user)) return noAuth();
  const sh=SS_MASTER.getSheetByName('Users'); if(!sh) return {success:false,message:'Sheet Users tidak ditemukan.'};
  const rows=sh.getDataRange().getValues(); const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  return {success:true, data:rows.slice(1).filter(r=>r.some(c=>c!=='')).map(r=>({
    id:   ci(H,'id')>=0  ? String(r[ci(H,'id')]).trim()   : '',
    role: ci(H,'role')>=0? String(r[ci(H,'role')]).trim()  : 'user',
    email:ci(H,'email')>=0? String(r[ci(H,'email')]).trim(): '',
    password:'***',
    nama: ci(H,'nama')>=0? String(r[ci(H,'nama')]).trim()  : '',
  }))};
}

function addUser({user,email,password,role,nama}) {
  if(!isAdmin(user)) return noAuth();
  const sh=SS_MASTER.getSheetByName('Users'); if(!sh) return {success:false};
  const H=sh.getDataRange().getValues()[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const newRow=new Array(H.length).fill('');
  const id='USR_'+String(Date.now()).slice(-6);
  if(ci(H,'id')>=0) newRow[ci(H,'id')]=id; if(ci(H,'role')>=0) newRow[ci(H,'role')]=role||'user';
  if(ci(H,'email')>=0) newRow[ci(H,'email')]=email; if(ci(H,'password')>=0) newRow[ci(H,'password')]=password;
  if(ci(H,'nama')>=0) newRow[ci(H,'nama')]=nama;
  sh.appendRow(newRow); return {success:true};
}

function editUser({user,id,email,password,role,nama}) {
  if(!isAdmin(user)) return noAuth();
  const sh=SS_MASTER.getSheetByName('Users'); if(!sh) return {success:false};
  const rows=sh.getDataRange().getValues(); const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  rows.forEach((r,i)=>{ if(!i) return; if(String(r[ci(H,'id')||0]).trim()===id){ if(ci(H,'email')>=0) sh.getRange(i+1,ci(H,'email')+1).setValue(email); if(ci(H,'role')>=0) sh.getRange(i+1,ci(H,'role')+1).setValue(role); if(ci(H,'nama')>=0) sh.getRange(i+1,ci(H,'nama')+1).setValue(nama); if(password&&password!=='***'&&ci(H,'password')>=0) sh.getRange(i+1,ci(H,'password')+1).setValue(password); } });
  return {success:true};
}

function deleteUser({user,id}) {
  if(!isAdmin(user)) return noAuth();
  return _deleteByField(SS_MASTER.getSheetByName('Users'),'id',id);
}

// ════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════
function isAdmin(user) {
  if(!user||!user.email) return false;
  const sh=SS_MASTER.getSheetByName('Users'); if(!sh) return false;
  const rows=sh.getDataRange().getValues();
  const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iM=ci(H,'email'), iR=ci(H,'role'); if(iM<0||iR<0) return false;
  return rows.slice(1).some(r=>
    String(r[iM]).trim().toLowerCase()===String(user.email).trim().toLowerCase()&&
    String(r[iR]).trim().toLowerCase()==='admin'
  );
}

function noAuth() { return {success:false, message:'Akses ditolak.'}; }

function _deleteByField(sh, field, value) {
  if(!sh) return {success:false, message:'Sheet tidak ditemukan.'};
  const rows=sh.getDataRange().getValues();
  const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iF=ci(H,field.toLowerCase().replace(/[_\s]/g,''));
  for(let i=rows.length-1;i>=1;i--){
    if(String(rows[i][iF||0]).trim()===value){sh.deleteRow(i+1);return {success:true};}
  }
  return {success:false,message:'Data tidak ditemukan.'};
}
