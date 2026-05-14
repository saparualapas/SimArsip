// ============================================================
//  SIMARSIP — Google Apps Script Backend  v2.3
//
//  PERUBAHAN UTAMA v2.3:
//  • submitSurat menerima {kodeKlasifikasi, keterangan} langsung
//    dari frontend — tidak perlu cari ke sheet Klasifikasi
//  • Sheet Data_Surat_YYYY format BERSIH:
//    No | No_Surat | Alamat_Penerima | Tanggal_Surat | Perihal
//  • Counter nomor urut pakai kolom No di sheet (auto-increment)
//  • getSurat membaca sheet Data_Surat dengan toleransi kolom
//
//  Sheet lain yang tetap dipakai:
//  • Users       : ID | Role | Email | Password | Nama
//  • KodeWilayah : key | value  (prefixWil, kodeWil)
//  • Config      : key | value  (cadangan)
//  • JenisArsip  : ID | Kode | Nama | Aktif (opsional, untuk admin)
//  • Klasifikasi : ID | Nama | Kode | Keterangan (hanya data parent)
// ============================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ── Sheet Data_Surat aktif (dari Config) ─────────────────────
function getSuratSheetName() {
  // Cek config aktifSheet
  let sh = SS.getSheetByName('Config');
  if (sh) {
    const rows = sh.getDataRange().getValues();
    const row  = rows.find(r => String(r[0]).trim() === 'aktifSheet');
    if (row && row[1]) {
      const name = String(row[1]).trim();
      if (SS.getSheetByName(name)) return name;
    }
  }
  // Fallback: cari Data_Surat_tahunIni atau Data_Surat
  const yr = new Date().getFullYear();
  if (SS.getSheetByName('Data_Surat_' + yr)) return 'Data_Surat_' + yr;
  if (SS.getSheetByName('Data_Surat'))       return 'Data_Surat';
  return 'Data_Surat_' + yr;
}

// ── Manajemen Sheet Data Surat ────────────────────────────────
function getSheetList({user}) {
  if (!isAdmin(user)) return noAuth();
  const aktif = getSuratSheetName();
  const sheets = SS.getSheets()
    .filter(sh => sh.getName().startsWith('Data_Surat'))
    .map(sh => ({
      name:   sh.getName(),
      rows:   Math.max(0, sh.getLastRow() - 1), // jumlah data (minus header)
      aktif:  sh.getName() === aktif,
    }));
  return {success:true, data: sheets, aktif};
}

function addDataSheet({user, sheetName}) {
  if (!isAdmin(user)) return noAuth();
  if (!sheetName) return {success:false, message:'Nama sheet wajib diisi.'};
  // Pastikan format: Data_Surat_YYYY atau Data_Surat_custom
  const safeName = sheetName.startsWith('Data_Surat') ? sheetName : 'Data_Surat_' + sheetName;
  if (SS.getSheetByName(safeName)) return {success:false, message:'Sheet "'+safeName+'" sudah ada.'};
  const sh = SS.insertSheet(safeName);
  sh.appendRow(['No','No_Surat','Alamat_Penerima','Tanggal_Surat','Perihal']);
  sh.setColumnWidth(1, 50);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 300);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 250);
  sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
  // Aktifkan otomatis sheet baru
  _setAktifSheet(safeName);
  return {success:true, name: safeName};
}

function setActiveSheet({user, sheetName}) {
  if (!isAdmin(user)) return noAuth();
  if (!SS.getSheetByName(sheetName)) return {success:false, message:'Sheet tidak ditemukan.'};
  _setAktifSheet(sheetName);
  return {success:true};
}

function deleteDataSheet({user, sheetName}) {
  if (!isAdmin(user)) return noAuth();
  if (!sheetName.startsWith('Data_Surat')) return {success:false, message:'Hanya sheet Data_Surat yang boleh dihapus.'};
  const sh = SS.getSheetByName(sheetName);
  if (!sh) return {success:false, message:'Sheet tidak ditemukan.'};
  // Jangan hapus jika cuma satu sheet Data_Surat
  const allDataSheets = SS.getSheets().filter(s => s.getName().startsWith('Data_Surat'));
  if (allDataSheets.length <= 1) return {success:false, message:'Minimal harus ada 1 sheet data surat.'};
  // Jika sheet ini aktif, pindahkan ke sheet lain dulu
  if (getSuratSheetName() === sheetName) {
    const other = allDataSheets.find(s => s.getName() !== sheetName);
    if (other) _setAktifSheet(other.getName());
  }
  SS.deleteSheet(sh);
  return {success:true};
}

function _setAktifSheet(name) {
  const sh = getSheet('Config');
  const rows = sh.getDataRange().getValues();
  let found = false;
  rows.forEach((r, i) => {
    if (i === 0) return;
    if (String(r[0]).trim() === 'aktifSheet') {
      sh.getRange(i+1, 2).setValue(name);
      found = true;
    }
  });
  if (!found) sh.appendRow(['aktifSheet', name]);
}

// ── Sheet getter + auto-create ───────────────────────────────
function getSheet(name) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    if (name.startsWith('Data_Surat')) {
      sh.appendRow(['No','No_Surat','Alamat_Penerima','Tanggal_Surat','Perihal']);
      // Format kolom
      sh.setColumnWidth(1, 50);
      sh.setColumnWidth(2, 220);
      sh.setColumnWidth(3, 300);
      sh.setColumnWidth(4, 120);
      sh.setColumnWidth(5, 250);
      // Bold header
      sh.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
    } else if (name === 'KodeWilayah') {
      sh.appendRow(['key','value']);
      sh.appendRow(['prefixWil','WP.28.PAS']);
      sh.appendRow(['kodeWil','8']);
    } else if (name === 'Config') {
      sh.appendRow(['key','value']);
    }
  }
  return sh;
}

// ── Header → kolom index (case-insensitive, toleran truncation) ─
function ci(H, name) {
  const n = name.toLowerCase().replace(/[_\s]/g,'');
  let i = H.indexOf(n);
  if (i >= 0) return i;
  i = H.findIndex(h => h.startsWith(n) || n.startsWith(h));
  return i;
}

function genId() { return Utilities.getUuid().replace(/-/g,'').substring(0,16); }

// ── doPost router ────────────────────────────────────────────
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
      case 'getSheetList':    result = getSheetList(body); break;
      case 'addDataSheet':    result = addDataSheet(body); break;
      case 'setActiveSheet':  result = setActiveSheet(body); break;
      case 'deleteDataSheet': result = deleteDataSheet(body); break;
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
    status:'SIMARSIP v2.3 OK', time:new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  AUTH
// ════════════════════════════════════════════════════════════
function login({email, password}) {
  const sh = SS.getSheetByName('Users');
  if (!sh) return {success:false, message:'Sheet Users tidak ditemukan.'};
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return {success:false, message:'Belum ada data pengguna.'};
  const H = rows[0].map(h => String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iId=ci(H,'id'), iRole=ci(H,'role'), iMail=ci(H,'email'), iPw=ci(H,'password'), iNama=ci(H,'nama');
  if (iMail<0||iPw<0) return {success:false, message:'Kolom Email/Password tidak ada di sheet Users.'};
  const emailIn = String(email).trim().toLowerCase();
  const pwIn    = String(password).trim();
  const found = rows.slice(1).find(r =>
    String(r[iMail]).trim().toLowerCase() === emailIn &&
    String(r[iPw]).trim() === pwIn
  );
  if (!found) return {success:false, message:'Email atau password salah.'};
  return {success:true, user:{
    id:    iId>=0   ? String(found[iId]).trim()   : emailIn,
    email: String(found[iMail]).trim(),
    role:  iRole>=0 ? String(found[iRole]).trim().toLowerCase() : 'user',
    nama:  iNama>=0 ? String(found[iNama]).trim() : email,
  }};
}

// ════════════════════════════════════════════════════════════
//  KODE WILAYAH
// ════════════════════════════════════════════════════════════
function getKodeWilayah() {
  let sh = SS.getSheetByName('KodeWilayah') || SS.getSheetByName('Config');
  if (!sh) sh = getSheet('KodeWilayah');
  const map = {prefixWil:'WP.28.PAS', kodeWil:'8'};
  sh.getDataRange().getValues().slice(1).forEach(r => {
    if (r[0]) map[String(r[0]).trim()] = String(r[1]).trim();
  });
  return {success:true, data:map};
}

function updateKodeWilayah({user, prefixWil, kodeWil}) {
  if (!isAdmin(user)) return noAuth();
  let sh = SS.getSheetByName('KodeWilayah');
  if (!sh) sh = getSheet('KodeWilayah');
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
//  JENIS ARSIP — hanya untuk panel admin, tidak dipakai di submit
// ════════════════════════════════════════════════════════════
function getJenisArsip() {
  const shJA = SS.getSheetByName('JenisArsip');
  if (shJA && shJA.getLastRow()>=2) {
    const rows = shJA.getDataRange().getValues();
    const H = rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
    const iId=ci(H,'id'), iKode=ci(H,'kode'), iNama=ci(H,'nama'), iAktif=ci(H,'aktif');
    const data = rows.slice(1).filter(r=>r.some(c=>c!=='')).map(r=>({
      id:   iId>=0   ? String(r[iId]).trim()  : String(r[iKode]||r[0]).trim(),
      kode: iKode>=0 ? String(r[iKode]).trim() : String(r[0]).trim(),
      nama: iNama>=0 ? String(r[iNama]).trim() : '',
    })).filter(j => {
      if (iAktif<0) return true;
      const v = String(rows.slice(1).find(r=>(iId>=0?r[iId]:r[iKode||0])===j.id||r[iId]===j.id)?.[iAktif]||'true').toUpperCase();
      return v!=='FALSE'&&v!=='0'&&v!=='TIDAK';
    });
    if (data.length) return {success:true, data};
  }
  // Fallback: derive dari Klasifikasi
  return _deriveJenisArsip();
}

function _deriveJenisArsip() {
  const sh = SS.getSheetByName('Klasifikasi');
  if (!sh || sh.getLastRow()<2) return {success:true, data:[]};
  const rows = sh.getDataRange().getValues();
  const H = rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iKode=ci(H,'kode'), iKet=ci(H,'keterangan');
  const map = {};
  rows.slice(1).filter(r=>r.some(c=>c!=='')).forEach(r=>{
    const kode = iKode>=0 ? String(r[iKode]).trim() : '';
    if (!kode) return;
    const prefix = kode.split('.')[0].replace(/[^A-Za-z]/g,'').toUpperCase();
    if (!prefix||map[prefix]) return;
    let nama = prefix;
    if (iKet>=0){
      const ket=String(r[iKet]).trim();
      const bag=ket.split('–')[0].split(' - ')[0].trim();
      if(bag&&bag.length>1&&bag.length<50) nama=bag;
    }
    map[prefix]={id:prefix, kode:prefix, nama};
  });
  return {success:true, data:Object.values(map)};
}

function addJenisArsip({user,kode,nama}) {
  if(!isAdmin(user)) return noAuth();
  let sh = SS.getSheetByName('JenisArsip');
  if (!sh) { sh = SS.insertSheet('JenisArsip'); sh.appendRow(['ID','Kode','Nama','Aktif']); }
  sh.appendRow([kode,kode,nama,true]);
  return {success:true};
}

function editJenisArsip({user,id,kode,nama,aktif}) {
  if(!isAdmin(user)) return noAuth();
  const sh=SS.getSheetByName('JenisArsip');
  if(!sh) return {success:false};
  const rows=sh.getDataRange().getValues();
  const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  rows.forEach((r,i)=>{
    if(!i) return;
    if(String(r[ci(H,'id')||0]).trim()===id){
      if(ci(H,'kode')>=0) sh.getRange(i+1,ci(H,'kode')+1).setValue(kode);
      if(ci(H,'nama')>=0) sh.getRange(i+1,ci(H,'nama')+1).setValue(nama);
      if(ci(H,'aktif')>=0) sh.getRange(i+1,ci(H,'aktif')+1).setValue(aktif);
    }
  });
  return {success:true};
}

function deleteJenisArsip({user,id}) {
  if(!isAdmin(user)) return noAuth();
  return _deleteByField(SS.getSheetByName('JenisArsip'),'id',id);
}

// ════════════════════════════════════════════════════════════
//  KLASIFIKASI (hanya untuk admin panel)
// ════════════════════════════════════════════════════════════
function getKlasifikasi({jenisArsipId}) {
  const sh=SS.getSheetByName('Klasifikasi');
  if(!sh||sh.getLastRow()<2) return {success:true,data:[]};
  const rows=sh.getDataRange().getValues();
  const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iId=ci(H,'id'), iNama=ci(H,'nama'), iKode=ci(H,'kode'), iKet=ci(H,'keterangan'), iAktif=ci(H,'aktif');
  let data=rows.slice(1).filter(r=>r.some(c=>c!=='')).map(r=>{
    const kode=iKode>=0?String(r[iKode]).trim():'';
    const prefix=kode.split('.')[0].replace(/[^A-Za-z]/g,'').toUpperCase();
    return {
      id:           iId>=0   ? String(r[iId]).trim()  : kode,
      nama:         iNama>=0 ? String(r[iNama]).trim() : '',
      kode:         kode,
      keterangan:   iKet>=0  ? String(r[iKet]).trim()  : '',
      jenisArsipId: prefix,
      aktif:        iAktif>=0 ? r[iAktif] : true,
    };
  }).filter(k=>{ const v=String(k.aktif).toUpperCase(); return v!=='FALSE'&&v!=='0'&&v!=='TIDAK'; });
  if(jenisArsipId) data=data.filter(k=>k.jenisArsipId===jenisArsipId);
  return {success:true, data};
}

function addKlasifikasi({user,jenisArsipId,kode,nama,keterangan}) {
  if(!isAdmin(user)) return noAuth();
  const sh=SS.getSheetByName('Klasifikasi');
  if(!sh) return {success:false,message:'Sheet Klasifikasi tidak ditemukan.'};
  const id='KLS_'+kode.replace(/\./g,'');
  sh.appendRow([id,nama||kode,kode,keterangan||'',true]);
  return {success:true};
}

function editKlasifikasi({user,id,kode,nama,keterangan,aktif}) {
  if(!isAdmin(user)) return noAuth();
  const sh=SS.getSheetByName('Klasifikasi');
  if(!sh) return {success:false};
  const rows=sh.getDataRange().getValues();
  const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  rows.forEach((r,i)=>{
    if(!i) return;
    if(String(r[ci(H,'id')||0]).trim()===id){
      if(ci(H,'nama')>=0)       sh.getRange(i+1,ci(H,'nama')+1).setValue(nama||kode);
      if(ci(H,'kode')>=0)       sh.getRange(i+1,ci(H,'kode')+1).setValue(kode);
      if(ci(H,'keterangan')>=0) sh.getRange(i+1,ci(H,'keterangan')+1).setValue(keterangan||'');
      if(ci(H,'aktif')>=0)      sh.getRange(i+1,ci(H,'aktif')+1).setValue(aktif);
    }
  });
  return {success:true};
}

function deleteKlasifikasi({user,id}) {
  if(!isAdmin(user)) return noAuth();
  return _deleteByField(SS.getSheetByName('Klasifikasi'),'id',id);
}

// ════════════════════════════════════════════════════════════
//  SURAT
//
//  submitSurat menerima:
//    kodeKlasifikasi  : kode lengkap dari frontend, e.g. "UM.01.01"
//    keterangan       : nama/keterangan klasifikasi dari frontend
//    alamat           : alamat penerima
//    perihal          : perihal surat
//    tglSurat         : tanggal surat (DD/MM/YYYY)
//
//  Sheet output: No | No_Surat | Alamat_Penerima | Tanggal_Surat | Perihal
//  Nomor urut  : baca lastRow sheet (termasuk header) → No = lastRow
//  No_Surat    : prefixWil.kodeWil.kodeKlasifikasi-No
// ════════════════════════════════════════════════════════════
function submitSurat({kodeKlasifikasi, keterangan, alamat, perihal, tglSurat}) {
  if (!kodeKlasifikasi) return {success:false, message:'Kode klasifikasi tidak dikirim.'};

  const wil = getKodeWilayah().data;
  const shSurat = getSheet(getSuratSheetName());

  // Nomor urut = jumlah baris termasuk header = lastRow (baris baru akan jadi lastRow+1)
  // Tapi No yang kita tulis = lastRow (karena header row=1, data mulai row=2)
  const nomorUrut = shSurat.getLastRow(); // = jumlah baris saat ini = No surat berikutnya

  // Format: WP.28.PAS.8.UM.01.01-187
  const noSurat = `${wil.prefixWil}.${wil.kodeWil}.${kodeKlasifikasi}-${nomorUrut}`;

  // Baca header untuk tentukan urutan kolom
  const hRow = shSurat.getRange(1,1,1,shSurat.getLastColumn()).getValues()[0];
  const H = hRow.map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));

  const newRow = new Array(H.length).fill('');

  // Map field → value
  const map = {
    'no':              nomorUrut,
    'nosurat':         noSurat,
    'no_surat':        noSurat,
    'alamatpenerima':  alamat || '',
    'alamat_penerima': alamat || '',
    'tanggalsurat':    tglSurat || '',
    'tanggal_surat':   tglSurat || '',
    'tanggalsura':     tglSurat || '',   // kolom lama terpotong
    'perihal':         perihal || '',
  };

  H.forEach((h,i) => { if (map[h] !== undefined) newRow[i] = map[h]; });

  shSurat.appendRow(newRow);
  return {success:true, noSurat, nomorUrut};
}

// ── Ambil data surat untuk admin ─────────────────────────────
function getSurat({user, limit}) {
  if (!isAdmin(user)) return noAuth();
  let allData = [];

  SS.getSheets().filter(sh=>sh.getName().startsWith('Data_Surat')).forEach(sh=>{
    const rows=sh.getDataRange().getValues();
    if(rows.length<2) return;
    const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));

    const iNo      = ci(H,'no');
    const iNoSurat = ci(H,'nosurat');
    const iAlamat  = ci(H,'alamatpenerima');
    const iTgl     = ci(H,'tanggalsurat');
    const iPerihal = ci(H,'perihal');

    rows.slice(1).filter(r=>r.some(c=>c!=='')).forEach((r,idx)=>{
      const noSurat = iNoSurat>=0 ? String(r[iNoSurat]).trim() : '';
      allData.push({
        rowIndex:  idx+2,
        no:        iNo>=0 ? r[iNo] : idx+1,
        noSurat,
        alamat:    iAlamat>=0  ? String(r[iAlamat]).trim()  : '',
        perihal:   iPerihal>=0 ? String(r[iPerihal]).trim() : '',
        tglSurat:  iTgl>=0     ? String(r[iTgl]).trim()     : '',
        sheetName: sh.getName(),
      });
    });
  });

  // Sort: yang ada noSurat dulu, lalu by rowIndex descending
  allData.sort((a,b) => {
    if (a.noSurat && !b.noSurat) return -1;
    if (!a.noSurat && b.noSurat) return 1;
    return b.rowIndex - a.rowIndex;
  });

  if(limit) allData=allData.slice(0,limit);
  return {success:true, data:allData};
}

function deleteSurat({user, rowIndex, sheetName}) {
  if (!isAdmin(user)) return noAuth();
  const shName = sheetName || getSuratSheetName();
  const sh = SS.getSheetByName(shName);
  if(!sh) return {success:false, message:'Sheet tidak ditemukan.'};
  if(!rowIndex||rowIndex<2) return {success:false, message:'Row tidak valid.'};
  sh.deleteRow(rowIndex);
  return {success:true};
}

// ════════════════════════════════════════════════════════════
//  USERS
// ════════════════════════════════════════════════════════════
function getUsers({user}) {
  if(!isAdmin(user)) return noAuth();
  const sh=SS.getSheetByName('Users');
  if(!sh) return {success:false,message:'Sheet Users tidak ditemukan.'};
  const rows=sh.getDataRange().getValues();
  const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iId=ci(H,'id'),iRole=ci(H,'role'),iMail=ci(H,'email'),iNama=ci(H,'nama');
  return {success:true, data:rows.slice(1).filter(r=>r.some(c=>c!=='')).map(r=>({
    id:    iId>=0  ? String(r[iId]).trim()  : '',
    role:  iRole>=0? String(r[iRole]).trim() : 'user',
    email: iMail>=0? String(r[iMail]).trim() : '',
    password:'***',
    nama:  iNama>=0? String(r[iNama]).trim() : '',
  }))};
}

function addUser({user,email,password,role,nama}) {
  if(!isAdmin(user)) return noAuth();
  const sh=SS.getSheetByName('Users');
  if(!sh) return {success:false};
  const H=sh.getDataRange().getValues()[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const newRow=new Array(H.length).fill('');
  const id='USR_'+String(Date.now()).slice(-6);
  if(ci(H,'id')>=0)       newRow[ci(H,'id')]       = id;
  if(ci(H,'role')>=0)     newRow[ci(H,'role')]     = role||'user';
  if(ci(H,'email')>=0)    newRow[ci(H,'email')]    = email;
  if(ci(H,'password')>=0) newRow[ci(H,'password')] = password;
  if(ci(H,'nama')>=0)     newRow[ci(H,'nama')]     = nama;
  sh.appendRow(newRow);
  return {success:true};
}

function editUser({user,id,email,password,role,nama}) {
  if(!isAdmin(user)) return noAuth();
  const sh=SS.getSheetByName('Users');
  if(!sh) return {success:false};
  const rows=sh.getDataRange().getValues();
  const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  rows.forEach((r,i)=>{
    if(!i) return;
    if(String(r[ci(H,'id')||0]).trim()===id){
      if(ci(H,'email')>=0)    sh.getRange(i+1,ci(H,'email')+1).setValue(email);
      if(ci(H,'role')>=0)     sh.getRange(i+1,ci(H,'role')+1).setValue(role);
      if(ci(H,'nama')>=0)     sh.getRange(i+1,ci(H,'nama')+1).setValue(nama);
      if(password&&password!=='***'&&ci(H,'password')>=0)
        sh.getRange(i+1,ci(H,'password')+1).setValue(password);
    }
  });
  return {success:true};
}

function deleteUser({user,id}) {
  if(!isAdmin(user)) return noAuth();
  return _deleteByField(SS.getSheetByName('Users'),'id',id);
}

// ════════════════════════════════════════════════════════════
//  SETUP — jalankan sekali untuk merapikan sheet existing
// ════════════════════════════════════════════════════════════
function setupSheets() {
  // Buat/periksa KodeWilayah
  getSheet('KodeWilayah');

  // Periksa Data_Surat_2026 — jika header lama, tambahkan kolom No di depan
  const shName = getSuratSheetName();
  let sh = SS.getSheetByName(shName);
  if (sh) {
    const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]
               .map(x=>String(x).trim().toLowerCase().replace(/[_\s]/g,''));
    if (!h.includes('no') && h.includes('nosurat')) {
      // Insert kolom No di posisi 1
      sh.insertColumnBefore(1);
      sh.getRange(1,1).setValue('No');
      // Isi No untuk baris yang sudah ada
      const lastR = sh.getLastRow();
      for (let r=2; r<=lastR; r++) sh.getRange(r,1).setValue(r-1);
      Logger.log('Setup: Kolom No ditambahkan ke '+shName);
    }
  }

  Logger.log('Setup selesai.');
  return 'OK';
}

// ════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════
function isAdmin(user) {
  if(!user||!user.email) return false;
  const sh=SS.getSheetByName('Users');
  if(!sh) return false;
  const rows=sh.getDataRange().getValues();
  const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iMail=ci(H,'email'), iRole=ci(H,'role');
  if(iMail<0||iRole<0) return false;
  return rows.slice(1).some(r=>
    String(r[iMail]).trim().toLowerCase()===String(user.email).trim().toLowerCase()&&
    String(r[iRole]).trim().toLowerCase()==='admin'
  );
}

function noAuth() { return {success:false, message:'Akses ditolak. Silakan login ulang.'}; }

function _deleteByField(sh, field, value) {
  if(!sh) return {success:false, message:'Sheet tidak ditemukan.'};
  const rows=sh.getDataRange().getValues();
  const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iF=ci(H,field.toLowerCase());
  for(let i=rows.length-1;i>=1;i--){
    if(String(rows[i][iF||0]).trim()===value){sh.deleteRow(i+1);return {success:true};}
  }
  return {success:false,message:'Data tidak ditemukan.'};
}
