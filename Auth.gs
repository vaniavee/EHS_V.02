/**
 * Modul Authentication & Authorization (Auth) bertanggung jawab untuk mengelola proses autentikasi (authentication) dan otorisasi (authorization) pada aplikasi 
 * Menggunakan Capability-based acces control: setiap user memiliki sekumpulan capabilities yang menentukan fitur apa saja yang dapat diakses
 */

// Mengambil profil lengkap pengguna beserta capability yang dimiliki
function getUserProfile_(nik) {
  const id = normalizeNik_(nik);
  const ss = getSpreadsheet_();
  ensureSheet_(ss, EHS.sheets.users, EHS_SCHEMA[EHS.sheets.users]);
  const users = readObjects_(ss.getSheetByName(EHS.sheets.users));
  const row = users.find(function(u) { return normalizeNik_(u.NIK) === id; });

  if (!row) return { nik: id, active: false, found: false, isRegistered: false };

  const isAdmin = isTrue_(row.IsAdmin);
  const isSupervisor = isTrue_(row.IsSupervisor);
  const isRegistered = String(row.IsRegistered).toLowerCase() === 'yes';
  const divisiDiawasi = clean_(row.DivisiDiawasi)
    .split(',').map(function(d) { return d.trim(); }).filter(Boolean);
  const programPreferences = clean_(row.ProgramPreferences)
    .split(',').map(function(p) { return p.trim(); }).filter(Boolean);

  let profileInterests = { goals: [], activities: [], topics: [] };
  const profileInterestsRaw = clean_(row.ProfileInterests || row.ProfileInterest);
  if (profileInterestsRaw) {
    try { profileInterests = JSON.parse(profileInterestsRaw); } catch (e) {}
  }

  return {
    nik: id,
    nama: clean_(row.Nama),
    divisi: clean_(row.Divisi),
    active: clean_(row.Active).toLowerCase() !== 'no',
    found: true,
    isRegistered: isRegistered,
    isAdmin: isAdmin,
    isSupervisor: isSupervisor,
    persona: isAdmin ? 'Admin' : (isSupervisor ? 'Leader' : 'Non-leader / Volunteer'),
    programPreferences: programPreferences,
    profileInterests: profileInterests,
    divisiDiawasi: divisiDiawasi,
    canSubmitHealthTask: true,
    canSubmitEnergyTask: true,
    //canSubmitSafetyReport: isSupervisor,
    canSubmitSafetyReport: true,
    canApproveOwnDivisionReports: isSupervisor,
    canManageMasterData: isAdmin,
    canApproveAllReports: isAdmin,
    canViewAdminDashboard: isAdmin
  };
}

// Mengubah representasi teks menjadi nilai Boolean
function isTrue_(v) {
  const s = clean_(v).toLowerCase();
  return s === 'true' || s === 'yes' || s === '1';
}

/**
 * Guard generik. Contoh pakai:
 *   assertCapability_(nik, 'canManageMasterData');
 *   assertCapability_(nik, 'canSubmitSafetyReport');
 */

// Memastikan pengguna memiliki hak akses terhadap suatu fitur
function assertCapability_(nik, capability) {
  const user = getUserProfile_(nik);
  if (!user.active) throw new Error('User tidak aktif atau tidak ditemukan.');
  if (!user[capability]) {
    throw new Error('Akses ditolak. Anda tidak memiliki izin: ' + capability);
  }
  return user;
}

/**
 * Khusus supervisor: cek apakah dia berwenang atas divisi tertentu
 * (misal saat approve laporan ST/SW milik karyawan di divisinya).
 */
// Memastikan supervisor berwenang terhadap divisi tertentu
function assertSupervisesDivision_(nik, targetDivisi) {
  const user = assertCapability_(nik, 'canSubmitSafetyReport');
  if (user.isAdmin) return user; // admin bebas semua divisi
  if (user.divisiDiawasi.indexOf(clean_(targetDivisi)) === -1) {
    throw new Error('Anda tidak berwenang atas divisi: ' + targetDivisi);
  }
  return user;
}

// Mengambil daftar divisi yang dapat diakses pengguna
function getDivisiDiawasiForUser(nik) {
  const user = getUserProfile_(nik);
  if (user.isAdmin) return ['Production', 'Engineering', 'Warehouse', 'Quality Control', 'Quality Assurance', 'HSE', 'Demand Planning', 'Office/Admin', 'Other'];
  //if (!user.isSupervisor) throw new Error('Hanya Kabag/Admin yang bisa mengakses form ini.');
  if (user.isSupervisor) return user.divisiDiawasi;
  return user.divisi ? [user.divisi] : [];
  //return user.divisiDiawasi;
}

/**
 * Login sederhana berbasis NIK, tanpa password — konsisten dengan pola
 * mentor (login.employeeId). Mengembalikan profile lengkap + capability
 * flags supaya frontend bisa langsung render UI yang sesuai role.
 */
// Melakukan autentifikasi pengguna menggunakan NIK
function loginUser(payload) {
  validateRequired_(payload, ['nik']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);

  if (!user.found) {
    throw new Error('NIK ' + nik + ' belum terdaftar di database perusahaan. Hubungi Admin EHS.');
  }
  if (!user.active) {
    throw new Error('Akun Anda tidak aktif. Hubungi Admin EHS.');
  }
  if (!user.isRegistered) {
    throw new Error('Akun Anda belum diaktivasi. Silakan lakukan Registrasi terlebih dahulu.');
  }

  logLoginEvent_(user);
  return user;
}

// Mencatat aktivitas login sebagai audit sistem
function logLoginEvent_(user) {
  // Log ringan untuk audit — sheet ini opsional, dibuat lazy kalau belum ada.
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName('15_DB_LoginLog');
  if (!sh) {
    sh = ss.insertSheet('15_DB_LoginLog');
    sh.getRange(1, 1, 1, 4).setValues([['Timestamp', 'NIK', 'Nama', 'Role']]);
    sh.setFrozenRows(1);
    sh.getRange(2, 2, 999, 1).setNumberFormat('@'); // kolom NIK
  }
  sh.appendRow([now_(), user.nik, user.nama, user.isAdmin ? 'Admin' : (user.isSupervisor ? 'Supervisor' : 'Participant')]);
}

/**
 * Data ringkasan untuk Dashboard User — dipanggil sekali saat dashboard
 * dibuka, menggabungkan profile + poin + rank supaya cuma 1 round-trip.
 */
// Menghasilkan ringkasan data dashboard pengguna
function getDashboardSummary(payload) {
  validateRequired_(payload, ['nik']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  if (!user.found) throw new Error('User tidak ditemukan.');

  const seasonId = normalizeSeasonId_(payload.seasonId);
  const totalPoints = getTotalPointsForUser_(nik, seasonId);
  const leaderboard = getLeaderboard_(seasonId, null, 9999);
  const myRank = leaderboard.find(function(r) { return r.nik === nik; });

  const result = {
    user: user,
    seasonId: seasonId,
    totalPoints: totalPoints,
    rank: myRank ? myRank.rank : '-',
    badge: resolveBadgeTier_(totalPoints)
  };

  if (!user.isAdmin) {
    result.domains = getDomainOverview({ nik: nik, seasonId: seasonId });
  }

  return result;
}

/**
 * Registrasi mandiri: user mencocokkan NIK + Nama dengan data yang
 * sudah di-preload Admin. Kalau cocok, akun diaktivasi (IsRegistered=Yes).
 * Pencocokan nama case-insensitive & abaikan spasi berlebih, supaya
 * tidak terlalu ketat (misal "budi santoso" vs "Budi Santoso").
 */
function registerUser(payload) {
  validateRequired_(payload, ['nik', 'nama']);
  const nik = normalizeNik_(payload.nik);

  const sh = getSpreadsheet_().getSheetByName(EHS.sheets.users);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  const rowIdx = data.findIndex(function(r, i) { return i > 0 && normalizeNikLenient_(r[col.NIK]) === nik; });
  if (rowIdx === -1) throw new Error('NIK tidak ditemukan di database perusahaan. Hubungi Admin EHS.');

  const row = data[rowIdx];
  const namaTerdaftar = clean_(row[col.Nama]).toLowerCase().replace(/\s+/g, ' ');
  const namaInput = clean_(payload.nama).toLowerCase().replace(/\s+/g, ' ');
  if (namaTerdaftar !== namaInput) {
    throw new Error('NIK dan Nama tidak cocok dengan data perusahaan. Periksa kembali ejaan nama Anda, atau hubungi Admin EHS.');
  }
  if (String(row[col.IsRegistered]).toLowerCase() === 'yes') {
    throw new Error('Akun ini sudah teregistrasi. Silakan langsung Login.');
  }

  sh.getRange(rowIdx + 1, col.IsRegistered + 1).setValue('Yes');
  if (payload.noWa) sh.getRange(rowIdx + 1, col.No_WA + 1).setValue(payload.noWa);

  const isPriority = isPriorityDivision_(clean_(row[col.Divisi]));
  return {
    ok: true,
    needsPreference: !isPriority, // volunteer -> lanjut ke step pilih program
    nik: nik,
    message: isPriority
      ? 'Registrasi berhasil! Divisi Anda memiliki program wajib otomatis.'
      : 'Registrasi berhasil! Mari sesuaikan program yang ingin Anda ikuti.'
  };
}

/**
 * Badge tier sesuai desain gamifikasi yang sudah disepakati sebelumnya.
 */
// Menentukan badge berdasarkan total poin pengguna
function resolveBadgeTier_(totalPoints) {
  if (totalPoints >= 1500) return { name: 'Role Model', icon: '👑', next: null };
  if (totalPoints >= 800) return { name: 'Juara EHS', icon: '🏆', next: 1500 };
  if (totalPoints >= 400) return { name: 'Penggerak', icon: '⚡', next: 800 };
  if (totalPoints >= 150) return { name: 'Konsisten', icon: '💧', next: 400 };
  return { name: 'Perintis', icon: '🌱', next: 150 };
}

function saveProgramPreference(payload) {
  validateRequired_(payload, ['nik', 'preferences']);
  const nik = normalizeNik_(payload.nik);
  const ss = getSpreadsheet_();
  ensureSheet_(ss, EHS.sheets.users, EHS_SCHEMA[EHS.sheets.users]);
  const sh = ss.getSheetByName(EHS.sheets.users);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const programPrefCol = getHeaderIndex_(headers, ['ProgramPreferences', 'ProgramPreference']);
  const nikCol = getHeaderIndex_(headers, ['NIK']);

  const rowIdx = data.findIndex(function(r, i) { return i > 0 && normalizeNikLenient_(r[nikCol]) === nik; });
  if (rowIdx === -1) throw new Error('User tidak ditemukan.');

  const prefsStr = Array.isArray(payload.preferences) ? payload.preferences.join(',') : clean_(payload.preferences);
  sh.getRange(rowIdx + 1, programPrefCol + 1).setValue(prefsStr);

  return { ok: true, message: 'Preferensi program tersimpan.' };
}

function saveProfileInterests(payload) {
  validateRequired_(payload, ['nik']);
  const nik = normalizeNik_(payload.nik);
  const ss = getSpreadsheet_();
  ensureSheet_(ss, EHS.sheets.users, EHS_SCHEMA[EHS.sheets.users]);
  const sh = ss.getSheetByName(EHS.sheets.users);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const nikCol = getHeaderIndex_(headers, ['NIK']);
  const profileInterestsCol = getHeaderIndex_(headers, ['ProfileInterests', 'ProfileInterest']);

  const rowIdx = data.findIndex(function(r, i) { return i > 0 && normalizeNikLenient_(r[nikCol]) === nik; });
  if (rowIdx === -1) throw new Error('User tidak ditemukan.');

  const interests = {
    goals: payload.goals || [],
    activities: payload.activities || [],
    topics: payload.topics || []
  };
  sh.getRange(rowIdx + 1, profileInterestsCol + 1).setValue(JSON.stringify(interests));

  return { ok: true, message: 'Preferensi tersimpan.' };
}