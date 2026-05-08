/**
 * SIMARSIP Backend v2.4
 * Struktur kolom: No Surat | Alamat Penerima | Tanggal Surat | Perihal
 * Fitur baru: ganti Spreadsheet ID dari panel admin (action: updateSpreadsheetId)
 */

// ── KONFIGURASI ──
// Jika MASTER_SPREADSHEET_ID diisi, semua data dibaca dari spreadsheet itu.
// Jika kosong, pakai Active Spreadsheet (spreadsheet tempat script ini dijalankan).
// Admin bisa mengubah ID ini dari dashboard tanpa perlu deploy ulang.
var MASTER_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';

const TAB_CONFIG      = 'Config';
const TAB_USERS       = 'Users';
const TAB_KLASIFIKASI = 'Klasifikasi';

// Kolom data surat (urutan di spreadsheet)
// A: No_Surat | B: Alamat_Penerima | C: Tanggal_Surat | D: Perihal | E: Timestamp_Input | F: User_Email | G: User_Nama
const COL_HEADERS = ['No_Surat', 'Alamat_Penerima', 'Tanggal_Surat', 'Perihal', 'Timestamp_Input', 'User_Email', 'User_Nama'];

// ── HELPERS ──
function makeRes(d) {
  return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON);
}
function ok(d)  { return makeRes({ success: true,  ...d }); }
function err(m) { return makeRes({ success: false, message: m }); }
function doGet() { return makeRes({ status: 'SIMARSIP API v2.4', spreadsheet_id: getSS().getId() }); }

function doPost(e) {
  try {
    const b = JSON.parse((e.postData && e.postData.contents) || '{}');
    switch (b.action) {
      case 'login':                return handleLogin(b);
      case 'getConfig':            return handleGetConfig();
      case 'updateConfig':         return handleUpdateConfig(b);
      case 'updateSpreadsheetId':  return handleUpdateSpreadsheetId(b);
      case 'getSurat':             return handleGetSurat(b);
      case 'generateNomor':        return handleGenerateNomor(b);
      case 'editSurat':            return handleEditSurat(b);
      case 'deleteSurat':          return handleDeleteSurat(b);
      case 'deleteSheet':          return handleDeleteSheet(b);
      case 'getUsers':             return handleGetUsers();
      case 'addUser':              return handleAddUser(b);
      case 'updateUser':           return handleUpdateUser(b);
      case 'deleteUser':           return handleDeleteUser(b);
      case 'getKlasifikasi':       return handleGetKlasifikasi();
      case 'addKlasifikasi':       return handleAddKlasifikasi(b);
      case 'updateKlasifikasi':    return handleUpdateKlasifikasi(b);
      case 'deleteKlasifikasi':    return handleDeleteKlasifikasi(b);
      case 'updateKodeWilayah':    return handleUpdateKodeWilayah(b);
      default:                     return err('Action tidak dikenal: ' + b.action);
    }
  } catch(e) { return err('Server error: ' + e.toString()); }
}

// ── SPREADSHEET ──
function getSS() {
  // Cek ScriptProperties dulu (bisa diubah dari admin)
  const savedId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (savedId && savedId.trim()) {
    try { return SpreadsheetApp.openById(savedId.trim()); } catch(e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getConfig() {
  const sh = getSS().getSheetByName(TAB_CONFIG);
  if (!sh) return {};
  const c = {};
  sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) c[String(r[0])] = r[1]; });
  return c;
}

function tabName(y) { return 'Data_Surat_' + y; }

function ensureSheet(ss, y) {
  const name = tabName(y);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    // Header dengan urutan baru: No Surat | Alamat Penerima | Tanggal Surat | Perihal | ...
    sh.appendRow(COL_HEADERS);
    const hdr = sh.getRange(1, 1, 1, COL_HEADERS.length);
    hdr.setFontWeight('bold');
    hdr.setBackground('#1a3a5c');
    hdr.setFontColor('#ffffff');
    hdr.setFontSize(11);
    // Lebar kolom
    sh.setColumnWidth(1, 90);   // No_Surat
    sh.setColumnWidth(2, 200);  // Alamat_Penerima
    sh.setColumnWidth(3, 180);  // Tanggal_Surat
    sh.setColumnWidth(4, 350);  // Perihal
    sh.setColumnWidth(5, 160);  // Timestamp_Input
    sh.setColumnWidth(6, 160);  // User_Email
    sh.setColumnWidth(7, 140);  // User_Nama
    sh.setFrozenRows(1);
    Logger.log('Sheet baru dibuat: ' + name);
  }
  return sh;
}

// Format tanggal dari ISO ke Indonesia
function formatTanggal(tanggal) {
  if (!tanggal) return '';
  const bln = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  try {
    // Input: "2026-04-27" atau "2026-04-27T14:30"
    const dt = new Date(tanggal.includes('T') ? tanggal : tanggal + 'T00:00:00+09:00');
    const tglStr = dt.getDate() + ' ' + bln[dt.getMonth()] + ' ' + dt.getFullYear();
    return tglStr;
  } catch(_) { return tanggal; }
}

// ── 1. LOGIN ──
function handleLogin(b) {
  const { email, password } = b;
  if (!email || !password) return err('Email dan password wajib diisi');
  const sh = getSS().getSheetByName(TAB_USERS);
  if (!sh) return err('Tab Users tidak ditemukan. Jalankan setupSpreadsheet() terlebih dahulu.');
  for (const r of sh.getDataRange().getValues().slice(1)) {
    if ((r[2]||'').toString().trim().toLowerCase() === email.trim().toLowerCase() &&
        (r[3]||'').toString().trim() === password.trim()) {
      return ok({ user: { id: r[0], role: r[1], email: r[2], nama: r[4] } });
    }
  }
  return err('Email atau password salah');
}

// ── 2. GET CONFIG ──
function handleGetConfig() {
  const ss = getSS();
  const config = getConfig();
  const available_years = ss.getSheets()
    .map(s => s.getName())
    .filter(n => n.startsWith('Data_Surat_'))
    .map(n => n.replace('Data_Surat_', ''))
    .sort().reverse();
  if (!available_years.length) available_years.push(String(config.Tahun_Aktif || new Date().getFullYear()));
  const uSh = ss.getSheetByName(TAB_USERS);
  const currentId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || ss.getId();
  // Sertakan daftar klasifikasi dan kode wilayah
  const klasResult = handleGetKlasifikasi();
  const klasData   = JSON.parse(klasResult.getContent());
  return ok({
    config,
    available_years,
    userCount: uSh ? Math.max(0, uSh.getLastRow()-1) : 0,
    spreadsheet_id: currentId,
    klasifikasi: klasData.success ? klasData.data : [],
    kode_wilayah: config.Kode_Wilayah || ''
  });
}

// ── 3. UPDATE CONFIG (tahun aktif / tambah sheet baru) ──
function handleUpdateConfig(b) {
  const ss = getSS();
  let cfgSh = ss.getSheetByName(TAB_CONFIG);
  if (!cfgSh) { cfgSh = ss.insertSheet(TAB_CONFIG); cfgSh.appendRow(['Key', 'Value']); }

  if (b.Tahun_Aktif) {
    const rows = cfgSh.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === 'Tahun_Aktif') { cfgSh.getRange(i+1, 2).setValue(b.Tahun_Aktif.toString()); found = true; break; }
    }
    if (!found) cfgSh.appendRow(['Tahun_Aktif', b.Tahun_Aktif.toString()]);
    ensureSheet(ss, b.Tahun_Aktif.toString());
  }
  if (b.newYear) ensureSheet(ss, b.newYear.toString());
  return ok({ message: 'Config diperbarui' });
}

// ── 4. UPDATE SPREADSHEET ID (dari admin dashboard) ──
function handleUpdateSpreadsheetId(b) {
  const { spreadsheet_id } = b;
  if (!spreadsheet_id || !spreadsheet_id.trim()) return err('Spreadsheet ID tidak boleh kosong');
  // Validasi: coba buka spreadsheet dengan ID tersebut
  let targetSS;
  try {
    targetSS = SpreadsheetApp.openById(spreadsheet_id.trim());
  } catch(e) {
    return err('Spreadsheet ID tidak valid atau tidak bisa diakses. Pastikan spreadsheet sudah dishare ke akun script ini.');
  }
  // Simpan ke ScriptProperties
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheet_id.trim());
  // Setup spreadsheet baru jika belum ada tab Config & Users
  const tahun = new Date().getFullYear().toString();
  let cfgSh = targetSS.getSheetByName(TAB_CONFIG);
  if (!cfgSh) {
    cfgSh = targetSS.insertSheet(TAB_CONFIG);
    cfgSh.appendRow(['Key', 'Value']);
    cfgSh.appendRow(['Tahun_Aktif', tahun]);
  }
  let usrSh = targetSS.getSheetByName(TAB_USERS);
  if (!usrSh) {
    usrSh = targetSS.insertSheet(TAB_USERS);
    usrSh.appendRow(['ID', 'Role', 'Email', 'Password', 'Nama']);
    // Copy users dari spreadsheet lama jika ada
    try {
      const oldSS = SpreadsheetApp.getActiveSpreadsheet();
      const oldUsr = oldSS.getSheetByName(TAB_USERS);
      if (oldUsr && oldUsr.getLastRow() > 1) {
        const oldUsers = oldUsr.getDataRange().getValues().slice(1);
        oldUsers.forEach(u => usrSh.appendRow(u));
      } else {
        usrSh.appendRow(['USR_001', 'admin', 'admin@instansi.go.id', 'admin123', 'Administrator']);
        usrSh.appendRow(['USR_002', 'user', 'user@instansi.go.id', 'user123', 'Staf Umum']);
      }
    } catch(_) {
      usrSh.appendRow(['USR_001', 'admin', 'admin@instansi.go.id', 'admin123', 'Administrator']);
    }
  }
  ensureSheet(targetSS, tahun);
  return ok({ message: 'Spreadsheet berhasil diganti ke: ' + targetSS.getName(), spreadsheet_name: targetSS.getName(), spreadsheet_id: spreadsheet_id.trim() });
}

// ── 5. DELETE SHEET ──
function handleDeleteSheet(b) {
  const { tahun } = b;
  if (!tahun) return err('Tahun wajib diisi');
  const ss = getSS();
  if (String(tahun) === String(getConfig().Tahun_Aktif)) return err('Tidak bisa menghapus sheet tahun aktif');
  const sh = ss.getSheetByName(tabName(tahun));
  if (!sh) return err('Sheet tidak ditemukan');
  if (ss.getSheets().length <= 1) return err('Tidak bisa menghapus satu-satunya sheet');
  ss.deleteSheet(sh);
  return ok({ message: 'Sheet ' + tahun + ' dihapus' });
}

// ── 6. GET SURAT ──
function handleGetSurat(b) {
  const y = String(b.tahun || new Date().getFullYear());
  const sh = getSS().getSheetByName(tabName(y));
  if (!sh || sh.getLastRow() <= 1) return ok({ surat: [], total: 0 });
  const data = sh.getDataRange().getValues();
  const hdrs = data[0];
  const surat = data.slice(1)
    .filter(r => r[0]) // harus ada No_Surat
    .map(r => {
      const obj = {};
      hdrs.forEach((h, i) => { obj[h] = r[i] !== undefined ? r[i].toString() : ''; });
      // Alias untuk kompatibilitas frontend
      obj['No']             = obj['No_Surat']        || '';
      obj['Perihal']        = obj['Perihal']          || '';
      obj['Alamat_Penerima']= obj['Alamat_Penerima']  || '';
      obj['Tanggal_Surat']  = obj['Tanggal_Surat']    || '';
      obj['Timestamp_Input']= obj['Timestamp_Input']  || '';
      return obj;
    });
  return ok({ surat, total: surat.length });
}

// ── 7. GENERATE NOMOR ──
function handleGenerateNomor(b) {
  const { perihal, alamat, tanggal, tahun, klasifikasi_id } = b;
  if (!perihal) return err('Perihal wajib diisi');
  const y = String(tahun || getConfig().Tahun_Aktif || new Date().getFullYear());
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e) { return err('Sistem sibuk, coba lagi'); }
  try {
    const ss = getSS(), sh = ensureSheet(ss, y);
    const lastRow = sh.getLastRow();
    let nextNo = 1;
    if (lastRow > 1) {
      const col = sh.getRange(2, 1, lastRow-1, 1).getValues();
      // Cari nomor urut terbesar dari kolom No_Surat (bisa format penuh atau angka)
      let mx = 0;
      col.forEach(r => {
        const val = r[0].toString();
        // Ambil angka di bagian akhir (setelah tanda "-" terakhir)
        const parts = val.split('-');
        const n = parseInt(parts[parts.length - 1]);
        if (!isNaN(n) && n > mx) mx = n;
      });
      nextNo = mx + 1;
    }

    // Bangun nomor surat lengkap jika ada klasifikasi
    let nomorSurat = String(nextNo);
    let kodeKlasifikasi = '';
    let namaJenis = '';
    if (klasifikasi_id) {
      const klas = getKlasifikasiById(klasifikasi_id);
      if (klas) {
        kodeKlasifikasi = klas.kode;
        namaJenis       = klas.nama;
        const cfg = getConfig();
        const kodeWilayah = cfg.Kode_Wilayah || '';
        // Format: [KodeWilayah]-[KodeKlasifikasi]-[NoUrut]
        // Contoh: PAS.1-UM.01.01-274
        nomorSurat = (kodeWilayah ? kodeWilayah + '-' : '') + kodeKlasifikasi + '-' + nextNo;
      }
    }

    const tglFmt = formatTanggal(tanggal);
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sh.appendRow([nomorSurat, alamat || '', tglFmt, perihal, ts, b.user_email || '', b.user_nama || '']);
    try { sh.autoResizeColumns(1, COL_HEADERS.length); } catch(_) {}
    return ok({
      nomor_surat: nomorSurat,
      no_urut: nextNo,
      tanggal: tglFmt,
      timestamp: ts,
      kode_klasifikasi: kodeKlasifikasi,
      nama_jenis: namaJenis
    });
  } finally { lock.releaseLock(); }
}

// ── 8. EDIT SURAT ──
function handleEditSurat(b) {
  const { no, perihal, alamat, tanggal, tahun } = b;
  if (!no) return err('No surat wajib ada');
  if (!perihal) return err('Perihal wajib diisi');
  const y = String(tahun || getConfig().Tahun_Aktif || new Date().getFullYear());
  const sh = getSS().getSheetByName(tabName(y));
  if (!sh) return err('Sheet tidak ditemukan');
  const rows = sh.getDataRange().getValues();
  const hdrs = rows[0];
  const iNo     = hdrs.indexOf('No_Surat');
  const iPerihal= hdrs.indexOf('Perihal');
  const iAlamat = hdrs.indexOf('Alamat_Penerima');
  const iTgl    = hdrs.indexOf('Tanggal_Surat');
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][iNo].toString() === no.toString()) {
      if (iPerihal >= 0) sh.getRange(i+1, iPerihal+1).setValue(perihal);
      if (iAlamat  >= 0) sh.getRange(i+1, iAlamat+1).setValue(alamat || '');
      if (iTgl     >= 0 && tanggal) sh.getRange(i+1, iTgl+1).setValue(formatTanggal(tanggal));
      return ok({ message: 'Surat berhasil diperbarui' });
    }
  }
  return err('Surat tidak ditemukan');
}

// ── 9. DELETE SURAT ──
function handleDeleteSurat(b) {
  const { no, tahun } = b;
  if (!no) return err('No wajib ada');
  const y = String(tahun || getConfig().Tahun_Aktif || new Date().getFullYear());
  const sh = getSS().getSheetByName(tabName(y));
  if (!sh) return err('Sheet tidak ditemukan');
  const rows = sh.getDataRange().getValues();
  const iNo = rows[0].indexOf('No_Surat');
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][iNo].toString() === no.toString()) { sh.deleteRow(i+1); return ok({ message: 'Surat dihapus' }); }
  }
  return err('Surat tidak ditemukan');
}

// ── KLASIFIKASI ──
function ensureKlasifikasiSheet(ss) {
  let sh = ss.getSheetByName(TAB_KLASIFIKASI);
  if (!sh) {
    sh = ss.insertSheet(TAB_KLASIFIKASI);
    sh.appendRow(['ID', 'Nama', 'Kode', 'Keterangan']);
    const hdr = sh.getRange(1, 1, 1, 4);
    hdr.setFontWeight('bold');
    hdr.setBackground('#1a3a5c');
    hdr.setFontColor('#ffffff');
    sh.setColumnWidth(1, 120);
    sh.setColumnWidth(2, 200);
    sh.setColumnWidth(3, 150);
    sh.setColumnWidth(4, 300);
    sh.setFrozenRows(1);
    // Seed contoh data klasifikasi umum
    sh.appendRow(['KLS_001', 'Umum', 'UM.01.01', 'Surat umum/undangan']);
    sh.appendRow(['KLS_002', 'Kepegawaian', 'KP.01.01', 'Surat bidang kepegawaian']);
    sh.appendRow(['KLS_003', 'Keuangan', 'KU.00.01', 'Surat bidang keuangan']);
    sh.appendRow(['KLS_004', 'Pembinaan Napi', 'PK.01.01', 'Surat pembinaan narapidana']);
    sh.appendRow(['KLS_005', 'Keamanan', 'KM.00.01', 'Surat bidang keamanan & ketertiban']);
  }
  return sh;
}

function handleGetKlasifikasi() {
  const sh = ensureKlasifikasiSheet(getSS());
  if (sh.getLastRow() <= 1) return makeRes({ success: true, data: [] });
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0]);
  const data = rows.map(r => ({ id: r[0], nama: r[1], kode: r[2], keterangan: r[3] || '' }));
  return makeRes({ success: true, data });
}

function getKlasifikasiById(id) {
  const sh = getSS().getSheetByName(TAB_KLASIFIKASI);
  if (!sh) return null;
  const rows = sh.getDataRange().getValues().slice(1);
  for (const r of rows) {
    if (r[0].toString() === id.toString()) return { id: r[0], nama: r[1], kode: r[2], keterangan: r[3] || '' };
  }
  return null;
}

function handleAddKlasifikasi(b) {
  const { nama, kode, keterangan } = b;
  if (!nama || !kode) return err('Nama dan kode wajib diisi');
  const sh = ensureKlasifikasiSheet(getSS());
  const rows = sh.getDataRange().getValues().slice(1);
  if (rows.some(r => r[2].toString().trim().toLowerCase() === kode.trim().toLowerCase()))
    return err('Kode klasifikasi sudah ada');
  sh.appendRow(['KLS_' + Date.now(), nama.trim(), kode.trim().toUpperCase(), keterangan || '']);
  return ok({ message: 'Klasifikasi ditambahkan' });
}

function handleUpdateKlasifikasi(b) {
  const { id, nama, kode, keterangan } = b;
  if (!id) return err('ID wajib ada');
  const sh = getSS().getSheetByName(TAB_KLASIFIKASI);
  if (!sh) return err('Sheet Klasifikasi tidak ditemukan');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === id.toString()) {
      // Cek duplikat kode (kecuali row sendiri)
      if (kode) {
        const dup = rows.slice(1).some((r, idx) => idx !== i-1 && r[2].toString().trim().toLowerCase() === kode.trim().toLowerCase());
        if (dup) return err('Kode sudah digunakan klasifikasi lain');
        sh.getRange(i+1, 3).setValue(kode.trim().toUpperCase());
      }
      if (nama) sh.getRange(i+1, 2).setValue(nama.trim());
      sh.getRange(i+1, 4).setValue(keterangan || '');
      return ok({ message: 'Klasifikasi diperbarui' });
    }
  }
  return err('Klasifikasi tidak ditemukan');
}

function handleDeleteKlasifikasi(b) {
  const { id } = b;
  if (!id) return err('ID wajib ada');
  const sh = getSS().getSheetByName(TAB_KLASIFIKASI);
  if (!sh) return err('Sheet Klasifikasi tidak ditemukan');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === id.toString()) { sh.deleteRow(i+1); return ok({ message: 'Klasifikasi dihapus' }); }
  }
  return err('Klasifikasi tidak ditemukan');
}

function handleUpdateKodeWilayah(b) {
  const { kode_wilayah } = b;
  if (kode_wilayah === undefined) return err('kode_wilayah wajib ada');
  const ss = getSS();
  let cfgSh = ss.getSheetByName(TAB_CONFIG);
  if (!cfgSh) { cfgSh = ss.insertSheet(TAB_CONFIG); cfgSh.appendRow(['Key', 'Value']); }
  const rows = cfgSh.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'Kode_Wilayah') { cfgSh.getRange(i+1, 2).setValue(kode_wilayah.toString().trim()); found = true; break; }
  }
  if (!found) cfgSh.appendRow(['Kode_Wilayah', kode_wilayah.toString().trim()]);
  return ok({ message: 'Kode wilayah diperbarui', kode_wilayah: kode_wilayah.toString().trim() });
}

// ── 10. USERS ──
function handleGetUsers() {
  const sh = getSS().getSheetByName(TAB_USERS);
  if (!sh) return ok({ users: [] });
  return ok({ users: sh.getDataRange().getValues().slice(1).filter(r => r[0])
    .map(r => ({ ID: r[0], Role: r[1], Email: r[2], Nama: r[4] })) });
}

function handleAddUser(b) {
  const { nama, email, password, role } = b;
  if (!nama || !email || !password) return err('Semua field wajib diisi');
  const ss = getSS();
  let sh = ss.getSheetByName(TAB_USERS);
  if (!sh) { sh = ss.insertSheet(TAB_USERS); sh.appendRow(['ID','Role','Email','Password','Nama']); }
  if (sh.getDataRange().getValues().slice(1).some(r => (r[2]||'').toLowerCase() === email.toLowerCase()))
    return err('Email sudah terdaftar');
  sh.appendRow(['USR_' + Date.now(), role || 'user', email, password, nama]);
  return ok({ message: 'User ditambahkan' });
}

function handleUpdateUser(b) {
  const { id, nama, email, password, role } = b;
  if (!id) return err('ID wajib ada');
  const sh = getSS().getSheetByName(TAB_USERS);
  if (!sh) return err('Tab Users tidak ditemukan');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === id.toString()) {
      if (email) {
        const dup = rows.slice(1).some((r, idx) => idx !== i-1 && (r[2]||'').toLowerCase() === email.toLowerCase());
        if (dup) return err('Email sudah digunakan akun lain');
        sh.getRange(i+1, 3).setValue(email);
      }
      if (nama    && nama.trim())     sh.getRange(i+1, 5).setValue(nama);
      if (password && password.trim()) sh.getRange(i+1, 4).setValue(password);
      if (role)                        sh.getRange(i+1, 2).setValue(role);
      return ok({ message: 'User berhasil diperbarui' });
    }
  }
  return err('User tidak ditemukan');
}

function handleDeleteUser(b) {
  const { id } = b;
  if (!id) return err('ID wajib ada');
  const sh = getSS().getSheetByName(TAB_USERS);
  if (!sh) return err('Tab Users tidak ditemukan');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === id.toString()) { sh.deleteRow(i+1); return ok({ message: 'User dihapus' }); }
  }
  return err('User tidak ditemukan');
}

// ── SETUP AWAL (jalankan sekali saja) ──
function setupSpreadsheet() {
  const ss = getSS();
  const y  = new Date().getFullYear().toString();
  let cfg = ss.getSheetByName(TAB_CONFIG);
  if (!cfg) { cfg = ss.insertSheet(TAB_CONFIG); cfg.appendRow(['Key','Value']); cfg.appendRow(['Tahun_Aktif', y]); }
  let usr = ss.getSheetByName(TAB_USERS);
  if (!usr) {
    usr = ss.insertSheet(TAB_USERS);
    usr.appendRow(['ID','Role','Email','Password','Nama']);
    usr.appendRow(['USR_001','admin','admin@instansi.go.id','admin123','Administrator']);
    usr.appendRow(['USR_002','user','user@instansi.go.id','user123','Staf Umum']);
  }
  ensureSheet(ss, y);
  ensureKlasifikasiSheet(ss);
  Logger.log('✅ Setup selesai v2.5');
  Logger.log('Spreadsheet ID: ' + ss.getId());
  Logger.log('Admin: admin@instansi.go.id / admin123');
  Logger.log('User : user@instansi.go.id / user123');
}
