// ============================================================
//  SIMARSIP — Google Apps Script Backend  v2.1
//  Kompatibel dengan struktur sheet yang sudah ada:
//
//  • Users       : ID | Role | Email | Password | Nama
//  • Klasifikasi : ID | Nama | Kode | Keterangan
//  • KodeWilayah : key | value
//  • Config      : key | value  (counter & konfigurasi)
//  • Data_Surat_YYYY : id | noSurat | klasifikasiId | jenisArsipId | perihal | pengirim | tglSurat | tglInput | userId
//  • JenisArsip  : ID | Kode | Nama | Aktif  (opsional, bisa di-derive dari Klasifikasi)
// ============================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ─── Data_Surat sheet name (per tahun) ───────────────────────
function getSuratSheetName() {
  const yr = new Date().getFullYear();
  const byYear = SS.getSheetByName('Data_Surat_' + yr);
  if (byYear) return 'Data_Surat_' + yr;
  const generic = SS.getSheetByName('Data_Surat');
  if (generic) return 'Data_Surat';
  return 'Data_Surat_' + yr;
}

// ─── Sheet getter + auto-create ──────────────────────────────
function getSheet(name) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    if (name.startsWith('Data_Surat')) {
      sh.appendRow(['id','noSurat','klasifikasiId','jenisArsipId','perihal','pengirim','tglSurat','tglInput','userId']);
    } else if (name === 'KodeWilayah') {
      sh.appendRow(['key','value']);
      sh.appendRow(['prefixWil','WP.28.PAS']);
      sh.appendRow(['kodeWil','8']);
    } else if (name === 'Config') {
      sh.appendRow(['key','value']);
    } else if (name === 'JenisArsip') {
      sh.appendRow(['ID','Kode','Nama','Aktif']);
      seedJenisArsipFromKlasifikasi(sh);
    }
  }
  return sh;
}

// ─── Sheet → array of objects (header case-insensitive) ──────
function sheetToObjects(sh) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  return data.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null && c !== undefined))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
}

function genId() {
  return Utilities.getUuid().replace(/-/g,'').substring(0,16);
}

// ─── Get column index safely ─────────────────────────────────
function col(headers, name) { return headers.indexOf(name.toLowerCase()); }

// ─── doPost router ───────────────────────────────────────────
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
      default: result = {success:false, message:'Action tidak dikenali: ' + body.action};
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false, message: 'Server error: ' + err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'SIMARSIP GAS v2.1 OK', time: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════
//  AUTH — membaca kolom dinamis (ID|Role|Email|Password|Nama)
// ══════════════════════════════════════════════════════════════
function login({email, password}) {
  const sh = SS.getSheetByName('Users');
  if (!sh) return {success:false, message:'Sheet Users tidak ditemukan.'};

  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return {success:false, message:'Belum ada data pengguna.'};

  const H = rows[0].map(h => String(h).trim().toLowerCase());
  const iId   = col(H,'id');
  const iRole = col(H,'role');
  const iMail = col(H,'email');
  const iPw   = col(H,'password');
  const iNama = col(H,'nama');

  if (iMail < 0 || iPw < 0) return {success:false, message:'Kolom Email/Password tidak ada di sheet Users.'};

  const emailIn = String(email).trim().toLowerCase();
  const pwIn    = String(password).trim();

  const found = rows.slice(1).find(r =>
    String(r[iMail]).trim().toLowerCase() === emailIn &&
    String(r[iPw]).trim() === pwIn
  );
  if (!found) return {success:false, message:'Email atau password salah.'};

  return {success:true, user:{
    id:    iId   >= 0 ? String(found[iId]).trim()   : emailIn,
    email: String(found[iMail]).trim(),
    role:  iRole >= 0 ? String(found[iRole]).trim().toLowerCase() : 'user',
    nama:  iNama >= 0 ? String(found[iNama]).trim()  : email,
  }};
}

// ══════════════════════════════════════════════════════════════
//  KODE WILAYAH
// ══════════════════════════════════════════════════════════════
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

// ══════════════════════════════════════════════════════════════
//  JENIS ARSIP
//  Derive otomatis dari prefix kode Klasifikasi (PR, KU, OT, SA...)
//  Nama diambil dari bagian sebelum " – " di kolom Keterangan
// ══════════════════════════════════════════════════════════════
function getJenisArsip() {
  // Jika ada sheet JenisArsip manual → pakai itu
  const shJA = SS.getSheetByName('JenisArsip');
  if (shJA && shJA.getLastRow() >= 2) {
    const rows = shJA.getDataRange().getValues();
    const H = rows[0].map(h=>String(h).trim().toLowerCase());
    const iId=col(H,'id'), iKode=col(H,'kode'), iNama=col(H,'nama'), iAktif=col(H,'aktif');
    const data = rows.slice(1).filter(r=>r.some(c=>c!=='')).map(r=>({
      id:   iId>=0   ? String(r[iId]).trim()   : String(r[iKode]||r[0]).trim(),
      kode: iKode>=0 ? String(r[iKode]).trim()  : String(r[0]).trim(),
      nama: iNama>=0 ? String(r[iNama]).trim()  : '',
    })).filter(j=>{
      const row = rows.slice(1).find(r=> (iId>=0?String(r[iId]):String(r[iKode||0])).trim()===j.id);
      if (!row||iAktif<0) return true;
      const v = String(row[iAktif]).toUpperCase();
      return v!=='FALSE'&&v!=='0'&&v!=='TIDAK';
    });
    if (data.length) return {success:true, data};
  }
  // Derive dari Klasifikasi
  return deriveJenisArsip();
}

function deriveJenisArsip() {
  const sh = SS.getSheetByName('Klasifikasi');
  if (!sh || sh.getLastRow()<2) return {success:true, data:[]};
  const rows = sh.getDataRange().getValues();
  const H = rows[0].map(h=>String(h).trim().toLowerCase());
  const iKode=col(H,'kode'), iKet=col(H,'keterangan');
  const map = {};
  rows.slice(1).filter(r=>r.some(c=>c!=='')).forEach(r=>{
    const kode = iKode>=0 ? String(r[iKode]).trim() : '';
    if (!kode) return;
    const prefix = kode.replace(/[^A-Za-z]/g,'').toUpperCase().slice(0,2);
    if (!prefix||map[prefix]) return;
    let nama = prefix;
    if (iKet>=0){
      const ket = String(r[iKet]).trim();
      const bag = ket.split('–')[0].split('-')[0].trim();
      if (bag && bag.length>1 && bag.length<50) nama = bag;
    }
    map[prefix] = {id:prefix, kode:prefix, nama};
  });
  return {success:true, data:Object.values(map)};
}

function seedJenisArsipFromKlasifikasi(sh) {
  const r = deriveJenisArsip();
  if (r.success) r.data.forEach(j=>sh.appendRow([j.id,j.kode,j.nama,true]));
}

function addJenisArsip({user,kode,nama}) {
  if (!isAdmin(user)) return noAuth();
  const sh = getSheet('JenisArsip');
  sh.appendRow([kode,kode,nama,true]);
  return {success:true};
}

function editJenisArsip({user,id,kode,nama,aktif}) {
  if (!isAdmin(user)) return noAuth();
  const sh = SS.getSheetByName('JenisArsip');
  if (!sh) return {success:false};
  const rows = sh.getDataRange().getValues();
  const H = rows[0].map(h=>String(h).trim().toLowerCase());
  rows.forEach((r,i)=>{
    if (!i) return;
    if (String(r[col(H,'id')||0]).trim()===id){
      if (col(H,'kode')>=0) sh.getRange(i+1,col(H,'kode')+1).setValue(kode);
      if (col(H,'nama')>=0) sh.getRange(i+1,col(H,'nama')+1).setValue(nama);
      if (col(H,'aktif')>=0) sh.getRange(i+1,col(H,'aktif')+1).setValue(aktif);
    }
  });
  return {success:true};
}

function deleteJenisArsip({user,id}) {
  if (!isAdmin(user)) return noAuth();
  return deleteRowByField(SS.getSheetByName('JenisArsip'),'id',id);
}

// ══════════════════════════════════════════════════════════════
//  KLASIFIKASI
//  Sheet: ID | Nama | Kode | Keterangan
//  jenisArsipId diderive dari prefix kode huruf
// ══════════════════════════════════════════════════════════════
function getKlasifikasi({jenisArsipId}) {
  const sh = SS.getSheetByName('Klasifikasi');
  if (!sh || sh.getLastRow()<2) return {success:true, data:[]};
  const rows = sh.getDataRange().getValues();
  const H = rows[0].map(h=>String(h).trim().toLowerCase());
  const iId=col(H,'id'), iNama=col(H,'nama'), iKode=col(H,'kode'), iKet=col(H,'keterangan'), iAktif=col(H,'aktif');

  let data = rows.slice(1).filter(r=>r.some(c=>c!=='')).map(r=>{
    const kode = iKode>=0 ? String(r[iKode]).trim() : '';
    const prefix = kode.replace(/[^A-Za-z]/g,'').toUpperCase().slice(0,2);
    return {
      id:           iId>=0   ? String(r[iId]).trim()   : kode,
      nama:         iNama>=0 ? String(r[iNama]).trim()  : '',
      kode:         kode,
      keterangan:   iKet>=0  ? String(r[iKet]).trim()   : '',
      jenisArsipId: prefix,
      aktif:        iAktif>=0 ? r[iAktif] : true,
    };
  }).filter(k=>{
    const v = String(k.aktif).toUpperCase();
    return v!=='FALSE'&&v!=='0'&&v!=='TIDAK';
  });

  if (jenisArsipId) data = data.filter(k=>k.jenisArsipId===jenisArsipId);
  return {success:true, data};
}

function addKlasifikasi({user,jenisArsipId,kode,nama,keterangan}) {
  if (!isAdmin(user)) return noAuth();
  const sh = SS.getSheetByName('Klasifikasi');
  if (!sh) return {success:false, message:'Sheet Klasifikasi tidak ditemukan.'};
  const id = 'KLS_'+kode.replace(/\./g,'');
  // Urutan kolom: ID | Nama | Kode | Keterangan
  sh.appendRow([id, nama||kode, kode, keterangan||'', true]);
  return {success:true};
}

function editKlasifikasi({user,id,kode,nama,keterangan,aktif}) {
  if (!isAdmin(user)) return noAuth();
  const sh = SS.getSheetByName('Klasifikasi');
  if (!sh) return {success:false};
  const rows = sh.getDataRange().getValues();
  const H = rows[0].map(h=>String(h).trim().toLowerCase());
  rows.forEach((r,i)=>{
    if (!i) return;
    if (String(r[col(H,'id')||0]).trim()===id){
      if (col(H,'nama')>=0)        sh.getRange(i+1,col(H,'nama')+1).setValue(nama||kode);
      if (col(H,'kode')>=0)        sh.getRange(i+1,col(H,'kode')+1).setValue(kode);
      if (col(H,'keterangan')>=0)  sh.getRange(i+1,col(H,'keterangan')+1).setValue(keterangan||'');
      if (col(H,'aktif')>=0)       sh.getRange(i+1,col(H,'aktif')+1).setValue(aktif);
    }
  });
  return {success:true};
}

function deleteKlasifikasi({user,id}) {
  if (!isAdmin(user)) return noAuth();
  return deleteRowByField(SS.getSheetByName('Klasifikasi'),'id',id);
}

// ══════════════════════════════════════════════════════════════
//  SURAT
// ══════════════════════════════════════════════════════════════
function submitSurat({jenisArsipId,klasifikasiId,perihal,pengirim,tglSurat,userId}) {
  const klData = getKlasifikasi({}).data;
  const kl = klData.find(k=>k.id===klasifikasiId);
  if (!kl) return {success:false, message:'Klasifikasi tidak ditemukan: '+klasifikasiId};

  const jaKode = jenisArsipId || kl.jenisArsipId;
  const wil    = getKodeWilayah().data;

  // Counter per kode klasifikasi, simpan di sheet Config
  const cKey = 'CTR_'+kl.kode.replace(/\./g,'');
  const shCfg = getSheet('Config');
  const cfgRows = shCfg.getDataRange().getValues();
  let cRow=-1, cVal=0;
  cfgRows.forEach((r,i)=>{ if(!i)return; if(String(r[0]).trim()===cKey){cRow=i+1;cVal=Number(r[1])||0;} });
  cVal++;
  if (cRow>0) shCfg.getRange(cRow,2).setValue(cVal);
  else shCfg.appendRow([cKey,cVal]);

  // Nomor surat: WP.28.PAS.8.PR.01-5
  const noSurat = `${wil.prefixWil}.${wil.kodeWil}.${kl.kode}-${cVal}`;

  const shSurat = getSheet(getSuratSheetName());
  const id = genId();
  shSurat.appendRow([id, noSurat, klasifikasiId, jaKode, perihal, pengirim, tglSurat, new Date().toISOString(), userId||'']);
  return {success:true, noSurat, id};
}

function getSurat({user,limit}) {
  if (!isAdmin(user)) return noAuth();
  const klAll = getKlasifikasi({}).data;
  let allData = [];

  SS.getSheets().filter(sh=>sh.getName().startsWith('Data_Surat')).forEach(sh=>{
    const rows = sh.getDataRange().getValues();
    if (rows.length<2) return;
    const H = rows[0].map(h=>String(h).trim().toLowerCase());
    rows.slice(1).filter(r=>r.some(c=>c!=='')).forEach(r=>{
      const klId = col(H,'klasifikasiid')>=0 ? String(r[col(H,'klasifikasiid')]).trim() : '';
      const klObj = klAll.find(k=>k.id===klId)||{};
      allData.push({
        id:             col(H,'id')>=0        ? r[col(H,'id')]       : '',
        noSurat:        col(H,'nosurat')>=0   ? r[col(H,'nosurat')]  : '',
        klasifikasiId:  klId,
        jenisArsipId:   col(H,'jenisarsipid')>=0 ? r[col(H,'jenisarsipid')] : klObj.jenisArsipId||'',
        perihal:        col(H,'perihal')>=0   ? r[col(H,'perihal')]  : '',
        pengirim:       col(H,'pengirim')>=0  ? r[col(H,'pengirim')] : '',
        tglSurat:       col(H,'tglsurat')>=0  ? r[col(H,'tglsurat')] : '',
        tglInput:       col(H,'tglinput')>=0  ? r[col(H,'tglinput')] : '',
        namaKlasifikasi: klObj.nama||klObj.kode||'',
        namaJenis:      klObj.jenisArsipId||'',
        sheetName:      sh.getName(),
      });
    });
  });

  allData.sort((a,b)=>String(b.tglInput).localeCompare(String(a.tglInput)));
  if (limit) allData = allData.slice(0,limit);
  return {success:true, data:allData};
}

function deleteSurat({user,id,sheetName}) {
  if (!isAdmin(user)) return noAuth();
  const sheets = sheetName
    ? [SS.getSheetByName(sheetName)].filter(Boolean)
    : SS.getSheets().filter(sh=>sh.getName().startsWith('Data_Surat'));
  for (const sh of sheets) {
    const r = deleteRowByField(sh,'id',id);
    if (r.success) return r;
  }
  return {success:false, message:'Data tidak ditemukan.'};
}

// ══════════════════════════════════════════════════════════════
//  USERS
// ══════════════════════════════════════════════════════════════
function getUsers({user}) {
  if (!isAdmin(user)) return noAuth();
  const sh = SS.getSheetByName('Users');
  if (!sh) return {success:false,message:'Sheet Users tidak ditemukan.'};
  const rows = sh.getDataRange().getValues();
  const H = rows[0].map(h=>String(h).trim().toLowerCase());
  return {success:true, data: rows.slice(1).filter(r=>r.some(c=>c!=='')).map(r=>({
    id:    col(H,'id')>=0    ? String(r[col(H,'id')]).trim()    : '',
    role:  col(H,'role')>=0  ? String(r[col(H,'role')]).trim()  : 'user',
    email: col(H,'email')>=0 ? String(r[col(H,'email')]).trim() : '',
    password:'***',
    nama:  col(H,'nama')>=0  ? String(r[col(H,'nama')]).trim()  : '',
  }))};
}

function addUser({user,email,password,role,nama}) {
  if (!isAdmin(user)) return noAuth();
  const sh = SS.getSheetByName('Users');
  if (!sh) return {success:false};
  const H = sh.getDataRange().getValues()[0].map(h=>String(h).trim().toLowerCase());
  const newRow = new Array(H.length).fill('');
  const id = 'USR_'+String(Date.now()).slice(-6);
  if(col(H,'id')>=0)       newRow[col(H,'id')]       = id;
  if(col(H,'role')>=0)     newRow[col(H,'role')]     = role||'user';
  if(col(H,'email')>=0)    newRow[col(H,'email')]    = email;
  if(col(H,'password')>=0) newRow[col(H,'password')] = password;
  if(col(H,'nama')>=0)     newRow[col(H,'nama')]     = nama;
  sh.appendRow(newRow);
  return {success:true};
}

function editUser({user,id,email,password,role,nama}) {
  if (!isAdmin(user)) return noAuth();
  const sh = SS.getSheetByName('Users');
  if (!sh) return {success:false};
  const rows = sh.getDataRange().getValues();
  const H = rows[0].map(h=>String(h).trim().toLowerCase());
  rows.forEach((r,i)=>{
    if (!i) return;
    if (String(r[col(H,'id')||0]).trim()===id){
      if(col(H,'email')>=0)    sh.getRange(i+1,col(H,'email')+1).setValue(email);
      if(col(H,'role')>=0)     sh.getRange(i+1,col(H,'role')+1).setValue(role);
      if(col(H,'nama')>=0)     sh.getRange(i+1,col(H,'nama')+1).setValue(nama);
      if(password&&password!=='***'&&col(H,'password')>=0)
        sh.getRange(i+1,col(H,'password')+1).setValue(password);
    }
  });
  return {success:true};
}

function deleteUser({user,id}) {
  if (!isAdmin(user)) return noAuth();
  return deleteRowByField(SS.getSheetByName('Users'),'id',id);
}

// ══════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════
function isAdmin(user) {
  if (!user||!user.email) return false;
  const sh = SS.getSheetByName('Users');
  if (!sh) return false;
  const rows = sh.getDataRange().getValues();
  const H = rows[0].map(h=>String(h).trim().toLowerCase());
  const iMail=col(H,'email'), iRole=col(H,'role');
  if (iMail<0||iRole<0) return false;
  return rows.slice(1).some(r=>
    String(r[iMail]).trim().toLowerCase()===String(user.email).trim().toLowerCase()&&
    String(r[iRole]).trim().toLowerCase()==='admin'
  );
}

function noAuth() { return {success:false, message:'Akses ditolak. Silakan login ulang.'}; }

function deleteRowByField(sh, fieldName, value) {
  if (!sh) return {success:false, message:'Sheet tidak ditemukan.'};
  const rows = sh.getDataRange().getValues();
  const H = rows[0].map(h=>String(h).trim().toLowerCase());
  const iF = col(H, fieldName.toLowerCase());
  for (let i=rows.length-1; i>=1; i--) {
    if (String(rows[i][iF||0]).trim()===value) { sh.deleteRow(i+1); return {success:true}; }
  }
  return {success:false, message:'Data tidak ditemukan.'};
}
