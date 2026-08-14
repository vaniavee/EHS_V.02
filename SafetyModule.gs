/**
 * Safety Module — Safety Talk (ST), Safety Walk (SW), Laporan Bahaya (Hazard)
 * Hanya bisa diakses Kabag/Supervisor (atau Admin) untuk divisi yang diawasi.
 * Alur validasi:
 *   ST      -> auto approve, poin langsung masuk (frekuensi tinggi, harian)
 *   SW      -> perlu approval Admin EHS sebelum poin dicatat
 *   Hazard  -> perlu approval Admin EHS, poin final ditentukan dari Severity
 */

// Menentukan jumlah poin yang diberikan untuk setiap jenis laporan keselamatan
const SAFETY_POINTS = {
  ST: 5,
  SW: 5,
  HAZARD_BASE: { Low: 5, Medium: 7, High: 10 }
};

// Menentukan jumlah maksimum laporan yang dapat dibuat dalam satu hari untuk setiap divisi
const SAFETY_DAILY_LIMIT = { ST: 1, SW: 1, Hazard: 5 }; // Hazard boleh lebih dari 1x/hari

/**
 * Submit laporan Safety Talk / Safety Walk / Hazard.
 * payload: { nik, jenisLaporan, divisiDilaporkan, deskripsi, buktiUrl?, severity? }
 */
// Memproses pengiriman laporan Safety Talk, Safety Walk, maupun Hazard
function submitSafetyReport(payload) {
  validateRequired_(payload, ['nik', 'jenisLaporan', 'divisiDilaporkan', 'deskripsi']);
  const jenis = clean_(payload.jenisLaporan);
  if (['ST', 'SW', 'Hazard'].indexOf(jenis) === -1) {
    throw new Error('JenisLaporan harus salah satu: ST, SW, Hazard.');
  }

  const user = assertCanReportDivision_(payload.nik, payload.divisiDilaporkan);
  const nik = user.nik;
  const seasonId = normalizeSeasonId_(payload.seasonId);
  const periodKey = getCurrentPeriodKey_('daily');
  const divisi = clean_(payload.divisiDilaporkan);

  const usedToday = countSafetyReportsInPeriod_(nik, jenis, seasonId, periodKey);
  const limit = SAFETY_DAILY_LIMIT[jenis];
  if (usedToday >= limit) {
    throw new Error(
      'Laporan ' + jenis + ' Anda sudah mencapai batas hari ini (' + limit + 'x).'
    );
  }

  const referenceId = [seasonId, 'SAFETY', jenis, nik, periodKey, usedToday + 1].join(':');

  // --- Tentukan status & poin awal berdasarkan jenis laporan ---
  let status, points, severity;
  if (jenis === 'ST') {
    status = 'Approved';         // auto — sesuai desain governance (semi-auto/auto utk ST)
    points = SAFETY_POINTS.ST;
    severity = '';
  } else if (jenis === 'SW') {
    status = 'Pending';          // butuh verifier
    points = 0;                  // poin baru dicatat saat di-approve admin
    severity = '';
  } else { // Hazard
    status = 'Pending';
    severity = clean_(payload.severity || 'Low');
    if (['Low', 'Medium', 'High'].indexOf(severity) === -1) severity = 'Low';
    points = 0;                  // poin final ditentukan admin saat approve (lihat approveSafetyReport)
  }
  // Penyimpanan bukti foto, sistem akan mengunggahnya ke Google Drive
  let buktiUrlFinal = clean_(payload.buktiUrl);
  if (payload.buktiBase64) {
    buktiUrlFinal = uploadPhotoToDrive_(payload.buktiBase64, payload.buktiMime, payload.buktiFileName);
  }
  appendObjectRow_(EHS.sheets.safetyReports, {
    Timestamp: now_(),
    SeasonId: seasonId,
    ReferenceId: referenceId,
    SupervisorNik: user.nik,
    SupervisorNama: user.nama,
    DivisiDilaporkan: divisi,
    JenisLaporan: jenis,
    Deskripsi: clean_(payload.deskripsi),
    BuktiUrl: buktiUrlFinal, //clean_(payload.buktiUrl),
    Severity: severity,
    Status: status,
    AdminFeedback: '',
    Points: points,
    PeriodKey: periodKey,
    ReviewedBy: '',
    ReviewedAt: ''
  });

  // ST langsung masuk Points Ledger karena auto-approve.
  if (jenis === 'ST') {
    awardPoints_(
      { nik: user.nik, nama: user.nama, divisi: user.divisi },
      'Safety', 'SAFETY_' + jenis, referenceId, points,
      'Safety Talk - ' + divisi, '', seasonId
    );
  }
  if (jenis !== 'ST') {
  notifyAdmins_('Laporan Safety Baru', jenis + ' dari ' + divisi + ' oleh ' + user.nama + ' menunggu review.', 'safety-review', referenceId);
  }
  return {
    ok: true,
    status: status,
    message: jenis === 'ST'
      ? 'Safety Talk tercatat, +' + points + ' poin.'
      : 'Laporan ' + jenis + ' terkirim, menunggu review Admin EHS.'
  };
}

// Menghitung jumlah laporan pada divisi dan periode tertentu
function countSafetyReportsInPeriod_(nik, jenis, seasonId, periodKey) {
  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.safetyReports));
  return rows.filter(function(r) {
    return normalizeNikLenient_(r.SupervisorNik) === nik &&
           clean_(r.JenisLaporan) === jenis &&
           clean_(r.SeasonId) === seasonId &&
           clean_(r.PeriodKey) === periodKey &&
           clean_(r.Status) !== 'Revise';
  }).length;
}

/**
 * Admin EHS approve laporan SW / Hazard -> poin baru masuk ke ledger di sini
 * (bukan saat submit), karena butuh verifikasi dulu.
 * payload: { adminNik, referenceId, finalSeverity? (opsional, admin bisa koreksi) }
 */
// Menyetujui laporan Safety Walk/Hazard serta memberikan poin
function approveSafetyReport(payload) {
  validateRequired_(payload, ['referenceId']);

  const sh = getSpreadsheet_().getSheetByName(EHS.sheets.safetyReports);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const refCol = headers.indexOf('ReferenceId');
  const statusCol = headers.indexOf('Status');
  const pointsCol = headers.indexOf('Points');
  const severityCol = headers.indexOf('Severity');
  const reviewedByCol = headers.indexOf('ReviewedBy');
  const reviewedAtCol = headers.indexOf('ReviewedAt');
  const jenisCol = headers.indexOf('JenisLaporan');
  const nikCol = headers.indexOf('SupervisorNik');
  const namaCol = headers.indexOf('SupervisorNama');
  const divisiCol = headers.indexOf('DivisiDilaporkan');
  const seasonCol = headers.indexOf('SeasonId');

  const rowIndex = data.findIndex(function(r) { return clean_(r[refCol]) === clean_(payload.referenceId); });
  if (rowIndex === -1) throw new Error('Laporan tidak ditemukan: ' + payload.referenceId);
  if (rowIndex === 0) throw new Error('Referensi tidak valid.');

  const admin = assertCanApprove_(payload.adminNik, clean_(row[divisiCol]))
  const row = data[rowIndex];
  if (clean_(row[statusCol]) !== 'Pending') {
    throw new Error('Laporan ini sudah direview sebelumnya (status: ' + row[statusCol] + ').');
  }

  const jenis = clean_(row[jenisCol]);
  const severity = clean_(payload.finalSeverity || row[severityCol] || 'Low');
  const points = jenis === 'SW'
    ? SAFETY_POINTS.SW
    : (SAFETY_POINTS.HAZARD_BASE[severity] || SAFETY_POINTS.HAZARD_BASE.Low);

  sh.getRange(rowIndex + 1, statusCol + 1).setValue('Approved');
  sh.getRange(rowIndex + 1, pointsCol + 1).setValue(points);
  sh.getRange(rowIndex + 1, severityCol + 1).setValue(severity);
  sh.getRange(rowIndex + 1, reviewedByCol + 1).setValue(admin.nama);
  sh.getRange(rowIndex + 1, reviewedAtCol + 1).setValue(now_());

  awardPoints_(
    { nik: row[nikCol], nama: row[namaCol], divisi: row[divisiCol] },
    'Safety', 'SAFETY_' + jenis, clean_(row[refCol]), points,
    jenis + ' disetujui - ' + row[divisiCol], '', clean_(row[seasonCol])
  );

  return { ok: true, points: points, message: jenis + ' disetujui, +' + points + ' poin.' };
}

/**
 * Admin menolak/minta revisi laporan SW/Hazard — tidak ada poin dicatat,
 * kuota harian divisi tersebut ikut terbuka lagi (lihat filter di
 * countSafetyReportsInPeriod_ yang mengecualikan status Revise).
 */

// Mengembalikan laporan kepada supervisor untuk diperbaiki
function reviseSafetyReport(payload) {
  const admin = assertCapability_(payload.adminNik, 'canApproveAllReports');
  validateRequired_(payload, ['referenceId', 'feedback']);

  const sh = getSpreadsheet_().getSheetByName(EHS.sheets.safetyReports);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const refCol = headers.indexOf('ReferenceId');
  const statusCol = headers.indexOf('Status');
  const feedbackCol = headers.indexOf('AdminFeedback');
  const reviewedByCol = headers.indexOf('ReviewedBy');
  const reviewedAtCol = headers.indexOf('ReviewedAt');

  const rowIndex = data.findIndex(function(r) { return clean_(r[refCol]) === clean_(payload.referenceId); });
  if (rowIndex === -1) throw new Error('Laporan tidak ditemukan: ' + payload.referenceId);

  sh.getRange(rowIndex + 1, statusCol + 1).setValue('Revise');
  sh.getRange(rowIndex + 1, feedbackCol + 1).setValue(clean_(payload.feedback));
  sh.getRange(rowIndex + 1, reviewedByCol + 1).setValue(admin.nama);
  sh.getRange(rowIndex + 1, reviewedAtCol + 1).setValue(now_());

  return { ok: true, message: 'Laporan dikembalikan untuk revisi.' };
}

/**
 * Admin memberi komentar/feedback TANPA mengubah status approval —
 * ini memenuhi requirement awal Vee: "Admin dapat memberikan
 * feedback/komentar terhadap laporan" (termasuk laporan yang sudah Approved,
 * misal ST harian yang auto-approve tapi tetap perlu dikomentari admin).
 */

// Menambahkan komentar administrator tanpa mengubah status laporan
function giveAdminFeedback(payload) {
  const admin = assertCapability_(payload.adminNik, 'canApproveAllReports');
  validateRequired_(payload, ['referenceId', 'feedback']);

  const sh = getSpreadsheet_().getSheetByName(EHS.sheets.safetyReports);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const refCol = headers.indexOf('ReferenceId');
  const feedbackCol = headers.indexOf('AdminFeedback');

  const rowIndex = data.findIndex(function(r) { return clean_(r[refCol]) === clean_(payload.referenceId); });
  if (rowIndex === -1) throw new Error('Laporan tidak ditemukan: ' + payload.referenceId);

  const existing = clean_(data[rowIndex][feedbackCol]);
  const combined = existing ? existing + ' | ' + admin.nama + ': ' + clean_(payload.feedback)
                             : admin.nama + ': ' + clean_(payload.feedback);
  sh.getRange(rowIndex + 1, feedbackCol + 1).setValue(combined);

  return { ok: true, message: 'Feedback tersimpan.' };
}

/**
 * Antrian review untuk Admin (SW & Hazard yang masih Pending).
 * Bisa juga dipakai Supervisor untuk lihat status laporan divisinya sendiri.
 */

// Mengambil daftar laporan yang dapat ditinjau sesuai hak akses pengguna
function listSafetyReportsForReview(payload) {
  const user = getUserProfile_(payload.nik);
  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.safetyReports));

  return rows.filter(function(r) {
    if (payload.statusFilter && clean_(r.Status) !== payload.statusFilter) return false;
    if (user.isAdmin) return true; // admin lihat semua divisi
    // supervisor hanya lihat laporan divisi yang dia awasi (atau miliknya sendiri)
    return user.divisiDiawasi.indexOf(clean_(r.DivisiDilaporkan)) !== -1 ||
           clean_(r.SupervisorNik) === user.nik;
  });
}

// Menggunggah foto bukti ke google drive
function uploadPhotoToDrive_(base64Data, mimeType, fileName) {
  const folder = getOrCreateEhsFolder_();
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName || 'bukti.jpg');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // Format ini bisa langsung dipakai sebagai src <img>, beda dari file.getUrl()
  // yang formatnya /view (halaman preview, bukan gambar mentah).
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

// Mengambil atau membuat folder penyimpanan bukti foto
function getOrCreateEhsFolder_() {
  const folderName = 'EHS App - Bukti Foto';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

// Mengambil riwayat laporan safety milik pengguna
function getMySafetyReports(payload) {
  validateRequired_(payload, ['nik']);
  const nik = normalizeNik_(payload.nik);
  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.safetyReports))
    .filter(function(r) { return normalizeNikLenient_(r.SupervisorNik) === nik; });
  rows.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  return rows.slice(0, 20);
}

function listAllSafetyReports(payload) {
  assertCapability_(payload.nik, 'canApproveAllReports');
  let rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.safetyReports));
  if (payload.status) rows = rows.filter(function(r) { return clean_(r.Status) === payload.status; });
  if (payload.jenis) rows = rows.filter(function(r) { return clean_(r.JenisLaporan) === payload.jenis; });
  rows.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  return rows;
}

function assertCanReportDivision_(nik, targetDivisi) {
  const user = getUserProfile_(nik);
  if (!user.active) throw new Error('User tidak aktif atau tidak ditemukan.');
  if (user.isAdmin) return user;
  if (user.isSupervisor) {
    if (user.divisiDiawasi.indexOf(clean_(targetDivisi)) === -1) {
      throw new Error('Anda tidak berwenang melaporkan untuk divisi: ' + targetDivisi);
    }
    return user;
  }
  if (clean_(user.divisi) !== clean_(targetDivisi)) {
    throw new Error('Anda hanya bisa melaporkan untuk divisi Anda sendiri: ' + user.divisi);
  }
  return user;
}

function resolveSafetyObligationLevel_(user, jenis) {
  // Leader/Admin: ST & SW tetap wajib, tidak terpengaruh preferensi.
  if (user.isAdmin || user.isSupervisor) {
    return (jenis === 'ST' || jenis === 'SW') ? 'Required' : 'Optional';
  }

  // Hazard selalu opsional untuk semua non-leader.
  if (jenis !== 'ST' && jenis !== 'SW') return 'Optional';

  // Divisi prioritas: ST/SW tetap wajib apapun preferensinya.
  if (isPriorityDivision_(user.divisi)) return 'Required';

  // Karyawan biasa yang pilih minat Safety Participation saat register -> naik jadi Recommended.
  const preferences = Array.isArray(user.programPreferences) ? user.programPreferences : [];
  const normalizedPreferences = preferences.map(normalizeText_);
  if (normalizedPreferences.indexOf('safetyparticipation') !== -1) return 'Recommended';

  return 'Optional';
}

function getSafetyMissionsForUser(payload) {
  validateRequired_(payload, ['nik']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  if (!user.active) throw new Error('NIK tidak terdaftar atau tidak aktif.');

  const seasonId = normalizeSeasonId_(payload.seasonId);
  const periodKey = getCurrentPeriodKey_('daily');

  const jenisList = ['ST', 'SW', 'Hazard'];
  return jenisList.map(function(jenis) {
    const used = countSafetyReportsInPeriod_(nik, jenis, seasonId, periodKey);
    const limit = SAFETY_DAILY_LIMIT[jenis];
    const available = used < limit;
    return {
      taskId: 'SAFETY_' + jenis,
      title: jenis === 'ST' ? 'Safety Talk' : (jenis === 'SW' ? 'Safety Walk' : 'Lapor Bahaya'),
      description: jenis === 'Hazard'
        ? 'Laporkan kondisi/tindakan tidak aman yang Anda temukan.'
        : 'Lakukan dan laporkan ' + (jenis === 'ST' ? 'Safety Talk' : 'Safety Walk') + ' harian.',
      points: jenis === 'Hazard' ? SAFETY_POINTS.HAZARD_BASE.Low : SAFETY_POINTS[jenis],
      obligationLevel: resolveSafetyObligationLevel_(user, jenis),
      used: used, limit: limit, available: available,
      reason: available ? '' : 'Kuota harian Anda untuk ' + jenis + ' sudah tercapai.'
    };
  });
}