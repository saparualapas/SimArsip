// ============================================================
//  SIMARSIP — Google Apps Script Backend  v2.0
//  Sheets yang dibutuhkan (buat otomatis jika belum ada):
//    • Users         : email | password | role | nama
//    • Surat         : id | noSurat | jenisArsipId | klasifikasiId | perihal | pengirim | tglSurat | tglInput | userId | status
//    • JenisArsip    : id | kode | nama | aktif
//    • Klasifikasi   : id | jenisArsipId | kode | keterangan | aktif
//    • KodeWilayah   : key | value        (row 1: kodeWil | WP.28.PAS)
//    • Counter       : key | count        (row per kombinasi jenisArsipId+klasifikasiId)
// ============================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ─── Sheet helpers ───────────────────────────────────────────
function getSheet(name) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    // seed header
    const headers = {
      Users:       ['id','email','password','role','nama'],
      Surat:       ['id','noSurat','jenisArsipId','klasifikasiId','perihal','pengirim','tglSurat','tglInput','userId','status'],
      JenisArsip:  ['id','kode','nama','aktif'],
      Klasifikasi: ['id','jenisArsipId','kode','keterangan','aktif'],
      KodeWilayah: ['key','value'],
      Counter:     ['key','count'],
    };
    if (headers[name]) sh.appendRow(headers[name]);
    // seed default data
    if (name === 'KodeWilayah') {
      sh.appendRow(['prefixWil','WP.28.PAS']);
      sh.appendRow(['kodeWil','8']);
    }
    if (name === 'Users') {
      sh.appendRow([genId(),'admin@simarsip.go.id','admin123','admin','Administrator']);
    }
  }
  return sh;
}

function sheetToObjects(sh) {
  const [headers, ...rows] = sh.getDataRange().getValues();
  return rows.map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

function genId() {
  return Utilities.getUuid().replace(/-/g,'').substring(0,16);
}

// ─── doPost router ───────────────────────────────────────────
function doPost(e) {
  const cors = ContentService.createTextOutput();
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    switch(action) {
      case 'login':           result = login(body); break;
      case 'getJenisArsip':  result = getJenisArsip(); break;
      case 'getKlasifikasi': result = getKlasifikasi(body); break;
      case 'getKodeWilayah': result = getKodeWilayah(); break;
      case 'submitSurat':    result = submitSurat(body); break;
      case 'getSurat':       result = getSurat(body); break;
      case 'deleteSurat':    result = deleteSurat(body); break;

      // Admin: Kode Wilayah
      case 'updateKodeWilayah': result = updateKodeWilayah(body); break;

      // Admin: Jenis Arsip CRUD
      case 'addJenisArsip':    result = addJenisArsip(body); break;
      case 'editJenisArsip':   result = editJenisArsip(body); break;
      case 'deleteJenisArsip': result = deleteJenisArsip(body); break;

      // Admin: Klasifikasi CRUD
      case 'addKlasifikasi':    result = addKlasifikasi(body); break;
      case 'editKlasifikasi':   result = editKlasifikasi(body); break;
      case 'deleteKlasifikasi': result = deleteKlasifikasi(body); break;

      // Admin: Users
      case 'getUsers':   result = getUsers(body); break;
      case 'addUser':    result = addUser(body); break;
      case 'editUser':   result = editUser(body); break;
      case 'deleteUser': result = deleteUser(body); break;

      default: result = {success:false, message:'Action tidak dikenali'};
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success:false, message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status:'SIMARSIP GAS OK'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── AUTH ─────────────────────────────────────────────────────
function login({email, password}) {
  const users = sheetToObjects(getSheet('Users'));
  const u = users.find(x => x.email === email && x.password === password);
  if (!u) return {success:false, message:'Email atau password salah.'};
  return {success:true, user:{id:u.id, email:u.email, role:u.role, nama:u.nama}};
}

// ─── KODE WILAYAH ─────────────────────────────────────────────
function getKodeWilayah() {
  const sh = getSheet('KodeWilayah');
  const rows = sh.getDataRange().getValues();
  const map = {};
  rows.slice(1).forEach(r => map[r[0]] = r[1]);
  return {success:true, data: map};
}

function updateKodeWilayah({user, prefixWil, kodeWil}) {
  if (!isAdmin(user)) return noAuth();
  const sh = getSheet('KodeWilayah');
  const rows = sh.getDataRange().getValues();
  rows.forEach((r, i) => {
    if (i === 0) return;
    if (r[0] === 'prefixWil') sh.getRange(i+1, 2).setValue(prefixWil);
    if (r[0] === 'kodeWil')   sh.getRange(i+1, 2).setValue(kodeWil);
  });
  return {success:true};
}

// ─── JENIS ARSIP ──────────────────────────────────────────────
function getJenisArsip() {
  const data = sheetToObjects(getSheet('JenisArsip')).filter(x => x.aktif !== false && x.aktif !== 'FALSE');
  return {success:true, data};
}

function addJenisArsip({user, kode, nama}) {
  if (!isAdmin(user)) return noAuth();
  getSheet('JenisArsip').appendRow([genId(), kode, nama, true]);
  return {success:true};
}

function editJenisArsip({user, id, kode, nama, aktif}) {
  if (!isAdmin(user)) return noAuth();
  const sh = getSheet('JenisArsip');
  const rows = sh.getDataRange().getValues();
  rows.forEach((r, i) => {
    if (i === 0) return;
    if (r[0] === id) {
      sh.getRange(i+1, 2).setValue(kode);
      sh.getRange(i+1, 3).setValue(nama);
      sh.getRange(i+1, 4).setValue(aktif);
    }
  });
  return {success:true};
}

function deleteJenisArsip({user, id}) {
  if (!isAdmin(user)) return noAuth();
  return deleteRowById('JenisArsip', id);
}

// ─── KLASIFIKASI ──────────────────────────────────────────────
function getKlasifikasi({jenisArsipId}) {
  let data = sheetToObjects(getSheet('Klasifikasi')).filter(x => x.aktif !== false && x.aktif !== 'FALSE');
  if (jenisArsipId) data = data.filter(x => x.jenisArsipId === jenisArsipId);
  return {success:true, data};
}

function addKlasifikasi({user, jenisArsipId, kode, keterangan}) {
  if (!isAdmin(user)) return noAuth();
  getSheet('Klasifikasi').appendRow([genId(), jenisArsipId, kode, keterangan, true]);
  return {success:true};
}

function editKlasifikasi({user, id, jenisArsipId, kode, keterangan, aktif}) {
  if (!isAdmin(user)) return noAuth();
  const sh = getSheet('Klasifikasi');
  const rows = sh.getDataRange().getValues();
  rows.forEach((r, i) => {
    if (i === 0) return;
    if (r[0] === id) {
      sh.getRange(i+1, 2).setValue(jenisArsipId);
      sh.getRange(i+1, 3).setValue(kode);
      sh.getRange(i+1, 4).setValue(keterangan);
      sh.getRange(i+1, 5).setValue(aktif);
    }
  });
  return {success:true};
}

function deleteKlasifikasi({user, id}) {
  if (!isAdmin(user)) return noAuth();
  return deleteRowById('Klasifikasi', id);
}

// ─── SURAT ────────────────────────────────────────────────────
function submitSurat({jenisArsipId, klasifikasiId, perihal, pengirim, tglSurat, userId}) {
  // Ambil data pendukung
  const jaList = sheetToObjects(getSheet('JenisArsip'));
  const klList = sheetToObjects(getSheet('Klasifikasi'));
  const wil    = getKodeWilayah().data;

  const ja = jaList.find(x => x.id === jenisArsipId);
  const kl = klList.find(x => x.id === klasifikasiId);
  if (!ja || !kl) return {success:false, message:'Jenis arsip atau klasifikasi tidak ditemukan.'};

  // Counter key
  const cKey = `${jenisArsipId}_${klasifikasiId}`;
  const cSh  = getSheet('Counter');
  const cRows = cSh.getDataRange().getValues();
  let cRow = -1, cVal = 0;
  cRows.forEach((r, i) => {
    if (i === 0) return;
    if (r[0] === cKey) { cRow = i+1; cVal = Number(r[1]); }
  });
  cVal++;
  if (cRow > 0) cSh.getRange(cRow, 2).setValue(cVal);
  else cSh.appendRow([cKey, cVal]);

  // Format: WP.28.PAS.8.UM.01.01-194
  // prefixWil.kodeWil.kodeJenis.kodeKlasifikasi-counter
  const noSurat = `${wil.prefixWil}.${wil.kodeWil}.${ja.kode}.${kl.kode}-${cVal}`;
  const id = genId();
  const tglInput = new Date().toISOString();

  getSheet('Surat').appendRow([id, noSurat, jenisArsipId, klasifikasiId, perihal, pengirim, tglSurat, tglInput, userId || '', 'aktif']);
  return {success:true, noSurat, id};
}

function getSurat({user, limit}) {
  if (!isAdmin(user)) return noAuth();
  let data = sheetToObjects(getSheet('Surat'));
  // enrich
  const ja = sheetToObjects(getSheet('JenisArsip'));
  const kl = sheetToObjects(getSheet('Klasifikasi'));
  data = data.map(s => ({
    ...s,
    namaJenis: (ja.find(x => x.id === s.jenisArsipId)||{}).nama || '',
    namaKlasifikasi: (kl.find(x => x.id === s.klasifikasiId)||{}).keterangan || '',
  }));
  if (limit) data = data.slice(-limit).reverse();
  return {success:true, data};
}

function deleteSurat({user, id}) {
  if (!isAdmin(user)) return noAuth();
  return deleteRowById('Surat', id);
}

// ─── USERS (admin) ────────────────────────────────────────────
function getUsers({user}) {
  if (!isAdmin(user)) return noAuth();
  const data = sheetToObjects(getSheet('Users')).map(u => ({...u, password:'***'}));
  return {success:true, data};
}

function addUser({user, email, password, role, nama}) {
  if (!isAdmin(user)) return noAuth();
  getSheet('Users').appendRow([genId(), email, password, role||'user', nama]);
  return {success:true};
}

function editUser({user, id, email, password, role, nama}) {
  if (!isAdmin(user)) return noAuth();
  const sh = getSheet('Users');
  const rows = sh.getDataRange().getValues();
  rows.forEach((r, i) => {
    if (i === 0) return;
    if (r[0] === id) {
      sh.getRange(i+1, 2).setValue(email);
      if (password && password !== '***') sh.getRange(i+1, 3).setValue(password);
      sh.getRange(i+1, 4).setValue(role);
      sh.getRange(i+1, 5).setValue(nama);
    }
  });
  return {success:true};
}

function deleteUser({user, id}) {
  if (!isAdmin(user)) return noAuth();
  return deleteRowById('Users', id);
}

// ─── Helpers ──────────────────────────────────────────────────
function isAdmin(user) {
  if (!user) return false;
  // verify di sheet
  const users = sheetToObjects(getSheet('Users'));
  const u = users.find(x => x.id === user.id && x.email === user.email && x.role === 'admin');
  return !!u;
}

function noAuth() { return {success:false, message:'Tidak diizinkan.'}; }

function deleteRowById(sheetName, id) {
  const sh = getSheet(sheetName);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][0] === id) { sh.deleteRow(i+1); return {success:true}; }
  }
  return {success:false, message:'Data tidak ditemukan.'};
}
