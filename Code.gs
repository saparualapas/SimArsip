// ============================================================
//  SIMARSIP — Google Apps Script Backend  v2.2
//
//  Struktur sheet Data_Surat_YYYY yang sudah ada:
//    No_Surat | Alamat_Penerima | Tanggal_Sura | Perihal |
//    Timestamp_Inpu | User_Emai | User_Nama
//    + kolom No (nomor urut) dibaca dari baris terakhir
//
//  Sheet Klasifikasi: ID | Nama | Kode | Keterangan
//    Kode contoh: UM.01.01, PR.01.01 → prefix huruf = jenis arsip
//    Nama = label jenis arsip (Program dan Anggaran, dst)
//
//  Nomor surat format: WP.28.PAS.8.UM.01.01-N
//    prefixWil.kodeWil.kodeKlasifikasi-nomorUrut
// ============================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ── Cari sheet Data_Surat tahun ini ──────────────────────────
function getSuratSheetName() {
  const yr = new Date().getFullYear();
  if (SS.getSheetByName('Data_Surat_' + yr)) return 'Data_Surat_' + yr;
  if (SS.getSheetByName('Data_Surat'))       return 'Data_Surat';
  return 'Data_Surat_' + yr;
}

// ── Sheet getter + auto-create ───────────────────────────────
function getSheet(name) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    if (name.startsWith('Data_Surat')) {
      // Header sesuai format sheet yang sudah ada
      sh.appendRow(['No_Surat','Alamat_Penerima','Tanggal_Sura','Perihal','Timestamp_Inpu','User_Emai','User_Nama']);
    } else if (name === 'KodeWilayah') {
      sh.appendRow(['key','value']);
      sh.appendRow(['prefixWil','WP.28.PAS']);
      sh.appendRow(['kodeWil','8']);
    } else if (name === 'Config') {
      sh.appendRow(['key','value']);
    } else if (name === 'JenisArsip') {
      sh.appendRow(['ID','Kode','Nama','Aktif']);
      _seedJenisFromKlasifikasi(sh);
    }
  }
  return sh;
}

// ── Header helper ─────────────────────────────────────────────
function hdrs(sh) {
  return sh.getDataRange().getValues()[0].map(h => String(h).trim().toLowerCase().replace(/[_\s]/g,''));
}
function ci(H, name) {
  // Cari kolom by nama, toleran terhadap underscore/spasi/truncation
  const n = name.toLowerCase().replace(/[_\s]/g,'');
  // Exact match dulu
  let i = H.indexOf(n);
  if (i >= 0) return i;
  // Partial: header starts with name (untuk kolom yang terpotong seperti "tanggal_sura")
  i = H.findIndex(h => h.startsWith(n) || n.startsWith(h));
  return i;
}

// ── Sheet → array objects ────────────────────────────────────
function sheetRows(sh) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const H = data[0].map(h => String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  return data.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null && c !== undefined))
    .map((r, idx) => {
      const obj = {_rowIndex: idx + 2}; // 1-based, +1 header, +1 for slice
      H.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
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
  return ContentService.createTextOutput(JSON.stringify({status:'SIMARSIP v2.2 OK', time:new Date().toISOString()}))
    .setMimeType(ContentService.MimeType.JSON);
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
  if (iMail<0||iPw<0) return {success:false, message:'Kolom Email/Password tidak ditemukan di sheet Users.'};
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
//  JENIS ARSIP
//  Derive dari prefix kode Klasifikasi: UM.01.01 → prefix "UM"
//  Nama jenis = bagian sebelum " – " di Keterangan, atau kolom Nama
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
    })).filter(j=>{
      if (iAktif<0) return true;
      const found = rows.slice(1).find(r=>(iId>=0?r[iId]:r[iKode||0])===j.id||r[iId]===j.id);
      if (!found) return true;
      const v = String(found[iAktif]).toUpperCase();
      return v!=='FALSE'&&v!=='0'&&v!=='TIDAK';
    });
    if (data.length) return {success:true, data};
  }
  return _deriveJenisArsip();
}

function _deriveJenisArsip() {
  const sh = SS.getSheetByName('Klasifikasi');
  if (!sh || sh.getLastRow()<2) return {success:true, data:[]};
  const rows = sh.getDataRange().getValues();
  const H = rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iKode=ci(H,'kode'), iNama=ci(H,'nama'), iKet=ci(H,'keterangan');
  const map = {};
  rows.slice(1).filter(r=>r.some(c=>c!=='')).forEach(r=>{
    const kode = iKode>=0 ? String(r[iKode]).trim() : '';
    if (!kode) return;
    // Prefix = huruf di awal sebelum titik pertama: UM.01.01 → UM
    const prefix = kode.split('.')[0].replace(/[^A-Za-z]/g,'').toUpperCase();
    if (!prefix||map[prefix]) return;
    // Nama = bagian sebelum " – " di keterangan
    let nama = prefix;
    if (iKet>=0){
      const ket=String(r[iKet]).trim();
      const bag=ket.split('–')[0].split(' - ')[0].trim();
      if(bag&&bag.length>1&&bag.length<50) nama=bag;
    } else if (iNama>=0) {
      nama = String(r[iNama]).trim() || prefix;
    }
    map[prefix]={id:prefix, kode:prefix, nama};
  });
  return {success:true, data:Object.values(map)};
}

function _seedJenisFromKlasifikasi(sh) {
  const r=_deriveJenisArsip();
  if(r.success) r.data.forEach(j=>sh.appendRow([j.id,j.kode,j.nama,true]));
}

function addJenisArsip({user,kode,nama}) {
  if(!isAdmin(user)) return noAuth();
  getSheet('JenisArsip').appendRow([kode,kode,nama,true]);
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
//  KLASIFIKASI
//  Sheet: ID | Nama | Kode | Keterangan
//  Kode contoh: UM.01.01, PR.01.02 — PAKAI KODE PENUH di nomor surat
//  jenisArsipId = prefix huruf dari kode (UM, PR, KU, ...)
// ════════════════════════════════════════════════════════════
function getKlasifikasi({jenisArsipId}) {
  const sh=SS.getSheetByName('Klasifikasi');
  if(!sh||sh.getLastRow()<2) return {success:true,data:[]};
  const rows=sh.getDataRange().getValues();
  const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
  const iId=ci(H,'id'), iNama=ci(H,'nama'), iKode=ci(H,'kode'), iKet=ci(H,'keterangan'), iAktif=ci(H,'aktif');

  let data=rows.slice(1).filter(r=>r.some(c=>c!=='')).map(r=>{
    const kode=iKode>=0?String(r[iKode]).trim():'';
    // Prefix = huruf pertama sebelum titik
    const prefix=kode.split('.')[0].replace(/[^A-Za-z]/g,'').toUpperCase();
    return {
      id:           iId>=0   ? String(r[iId]).trim()  : kode,
      nama:         iNama>=0 ? String(r[iNama]).trim() : '',
      kode:         kode,           // KODE PENUH: UM.01.01
      keterangan:   iKet>=0  ? String(r[iKet]).trim()  : '',
      jenisArsipId: prefix,
      aktif:        iAktif>=0 ? r[iAktif] : true,
    };
  }).filter(k=>{
    const v=String(k.aktif).toUpperCase();
    return v!=='FALSE'&&v!=='0'&&v!=='TIDAK';
  });

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
//  SURAT — menulis ke sheet Data_Surat_YYYY dengan format asli
//
//  Kolom output: No_Surat | Alamat_Penerima | Tanggal_Sura |
//                Perihal | Timestamp_Inpu | User_Emai | User_Nama
//
//  Nomor urut (No) dibaca dari jumlah baris data existing
//  Nomor surat: prefixWil.kodeWil.kodeKlasifikasiPenuh-nomorUrut
//  Contoh:      WP.28.PAS.8.UM.01.01-187
// ════════════════════════════════════════════════════════════
function submitSurat({jenisArsipId, klasifikasiId, alamat, perihal, tglSurat, userEmail, userNama}) {
  // Ambil data klasifikasi
  const klData = getKlasifikasi({}).data;
  const kl = klData.find(k=>k.id===klasifikasiId);
  if (!kl) return {success:false, message:'Klasifikasi tidak ditemukan: '+klasifikasiId};

  // Kode wilayah
  const wil = getKodeWilayah().data;

  // Nomor urut — baca dari jumlah baris di sheet + 1
  const shSurat = getSheet(getSuratSheetName());
  const lastRow = shSurat.getLastRow(); // termasuk header
  // lastRow - 1 = jumlah data (header = row 1)
  const nomorUrut = lastRow; // row berikutnya = lastRow+1, tapi nomorUrut = lastRow (karena header=1)

  // Format nomor surat LENGKAP: WP.28.PAS.8.UM.01.01-187
  // kl.kode sudah berisi kode penuh misal "UM.01.01"
  const noSurat = `${wil.prefixWil}.${wil.kodeWil}.${kl.kode}-${nomorUrut}`;

  // Cek header sheet untuk tahu urutan kolom yang benar
  const hRow = shSurat.getRange(1,1,1,shSurat.getLastColumn()).getValues()[0];
  const H = hRow.map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));

  // Buat baris baru sesuai urutan kolom sheet asli
  const newRow = new Array(H.length).fill('');
  const now = new Date().toISOString();

  // Map nilai ke kolom yang tepat
  const mapping = {
    'nosurat':          noSurat,
    'no_surat':         noSurat,
    'alamatpenerima':   alamat || '',
    'alamat_penerima':  alamat || '',
    'tanggalsura':      tglSurat || '',
    'tanggal_sura':     tglSurat || '',
    'tanggalsurat':     tglSurat || '',
    'tanggal_surat':    tglSurat || '',
    'perihal':          perihal || '',
    'timestampinpu':    now,
    'timestamp_inpu':   now,
    'timestamp':        now,
    'useremai':         userEmail || '',
    'user_emai':        userEmail || '',
    'useremail':        userEmail || '',
    'usernama':         userNama  || '',
    'user_nama':        userNama  || '',
  };

  H.forEach((h,i) => {
    if (mapping[h] !== undefined) newRow[i] = mapping[h];
  });

  shSurat.appendRow(newRow);
  return {success:true, noSurat, nomorUrut};
}

// ── Ambil data surat untuk admin ────────────────────────────
function getSurat({user, limit}) {
  if (!isAdmin(user)) return noAuth();
  const klAll = getKlasifikasi({}).data;
  let allData = [];

  SS.getSheets().filter(sh=>sh.getName().startsWith('Data_Surat')).forEach(sh=>{
    const rows=sh.getDataRange().getValues();
    if(rows.length<2) return;
    const H=rows[0].map(h=>String(h).trim().toLowerCase().replace(/[_\s]/g,''));
    const iNo    = ci(H,'nosurat');
    const iAlamat= ci(H,'alamatpenerima');
    const iTgl   = ci(H,'tanggalsura');      // kolom terpotong
    const iPerihal=ci(H,'perihal');
    const iTglIn = ci(H,'timestampinpu');
    const iUEmail= ci(H,'useremai');
    const iUNama = ci(H,'usernama');

    rows.slice(1).filter(r=>r.some(c=>c!=='')).forEach((r,idx)=>{
      const noSurat = iNo>=0 ? String(r[iNo]).trim() : '';
      // Skip baris lama yang tidak punya noSurat valid (data sebelum pakai sistem ini)
      allData.push({
        no:       idx+1,
        noSurat:  noSurat,
        alamat:   iAlamat>=0 ? String(r[iAlamat]).trim() : '',
        perihal:  iPerihal>=0 ? String(r[iPerihal]).trim() : '',
        tglSurat: iTgl>=0 ? r[iTgl] : '',
        tglInput: iTglIn>=0 ? r[iTglIn] : '',
        userEmail:iUEmail>=0 ? String(r[iUEmail]).trim() : '',
        userNama: iUNama>=0 ? String(r[iUNama]).trim() : '',
        sheetName:sh.getName(),
        rowIndex: idx+2, // 1-based, +1 header
      });
    });
  });

  // Sort terbaru dulu
  allData.sort((a,b)=>String(b.tglInput).localeCompare(String(a.tglInput)));
  if(limit) allData=allData.slice(0,limit);
  return {success:true, data:allData};
}

// ── Hapus surat ─────────────────────────────────────────────
function deleteSurat({user, rowIndex, sheetName}) {
  if(!isAdmin(user)) return noAuth();
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
