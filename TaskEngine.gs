/**
 * Catat poin ke ledger secara idempotent (anti double-count) dan aman
 * dari race condition (LockService) saat banyak user submit bersamaan.
 * Mengembalikan jumlah poin yang benar-benar dicatat (0 jika duplikat).
 */

// Pengelolaan task, validasi penyelesaian task, pemberian poin, serta pencatatan transaksi poin ke dalam Points Ledger.

// Fungsi untuk menambahkan poin pengguna ke tabel PointsLedger
function awardPoints_(user, pillar, taskId, referenceId, points, note, campaignId, seasonId) {
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(10000); // tunggu maks 10 detik
  if (!gotLock) {
    throw new Error('Sistem sedang sibuk, coba submit ulang beberapa detik lagi.');
  }
  try {
    const sh = getSpreadsheet_().getSheetByName(EHS.sheets.pointsLedger);
    const sid = clean_(seasonId || getActiveSeason_().SeasonId);
    const rows = readObjects_(sh).filter(function(r) { return clean_(r.SeasonId) === sid; });
    const exists = rows.some(function(r) { return clean_(r.ReferenceId) === clean_(referenceId); });
    if (exists) return 0;

    // Menambahkan satu baris baru ke Spreadsheet
    appendObjectRow_(EHS.sheets.pointsLedger, {
      Timestamp: now_(),
      SeasonId: sid,
      Pillar: pillar,
      TaskId: taskId || '',
      CampaignId: campaignId || '',
      NIK: user.nik,
      Nama: user.nama,
      Divisi: user.divisi,
      ReferenceId: referenceId,
      Points: Number(points || 0),
      Note: note || ''
    });
    return Number(points || 0);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Ambil daftar task untuk 1 pilar (Health/Energy), lengkap dengan status
 * ketersediaan (sudah diklaim berapa kali, kuota, kapan reset) — supaya
 * frontend tidak perlu panggil getTaskAvailability_ satu-satu per task.
 */
function getTasksForUser(payload) {
  validateRequired_(payload, ['nik', 'pillar']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  if (!user.active) throw new Error('NIK tidak terdaftar atau tidak aktif.');

  const seasonId = normalizeSeasonId_(payload.seasonId);
  const pillar = clean_(payload.pillar);

  const tasks = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.tasks))
    .filter(function(t) {
      return clean_(t.Pillar) === pillar &&
             clean_(t.Status).toLowerCase() === 'active' &&
             (!clean_(t.SeasonId) || clean_(t.SeasonId) === seasonId);
    });

  return tasks.map(function(t) {
    const availability = getTaskAvailability_(t, nik, seasonId);
    return {
      taskId: t.TaskId,
      campaignId: clean_(t.CampaignId), // <- ditambahkan
      title: t.Title,
      description: t.Description,
      points: Number(t.Points),
      frequencyType: availability.frequencyType,
      validation: t.Validation,
      available: availability.available,
      used: availability.used,
      limit: availability.limit,
      reason: availability.reason
    };
  });
}

/**
 * Submit task Health/Energy dengan data detail (jenis aktivitas, rute,
 * checklist, foto, dst). Selain Konimex Move (auto), semua masuk status
 * Pending menunggu verifikasi Admin — poin baru dicatat saat disetujui.
 */
function submitHealthTaskDetailed(payload) {
  validateRequired_(payload, ['nik', 'taskId']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  if (!user.active) throw new Error('NIK tidak terdaftar atau tidak aktif.');

  const seasonId = normalizeSeasonId_(payload.seasonId);
  const task = getTaskById_(payload.taskId, seasonId);
  if (!task) throw new Error('Task tidak ditemukan atau tidak aktif: ' + payload.taskId);

  const availability = getTaskAvailability_(task, nik, seasonId);
  if (!availability.available) throw new Error(availability.reason);

  const claimNumber = availability.used + 1;
  const referenceId = [seasonId, task.TaskId, nik, availability.periodKey, claimNumber].join(':');

  let buktiUrl = '';
  if (payload.buktiBase64) {
    buktiUrl = uploadPhotoToDrive_(payload.buktiBase64, payload.buktiMime, payload.buktiFileName);
  }

  // Poin dinamis khusus Konimex Walk Challenge: Points di master = poin per putaran.
  const detail = payload.detail || {};
  let points = Number(task.Points || 0);
  if (task.TaskId === 'H04' && detail.jumlahPutaran) {
    points = Number(task.Points) * Number(detail.jumlahPutaran);
  }

  const isHealth = clean_(task.Pillar).toLowerCase() === 'health';
  const hasPhoto = Boolean(payload.buktiBase64);
  const photoTakenAt = payload.photoTakenAt ? new Date(payload.photoTakenAt) : null;
  const enforcePhotoDate = Boolean(payload.verifyPhotoDate);
  const photoDateValid = enforcePhotoDate
    ? (photoTakenAt instanceof Date && !isNaN(photoTakenAt) && isSameDay_(photoTakenAt, new Date()))
    : true;
  const isAuto = isHealth && (clean_(task.Validation).toLowerCase() === 'auto' || !hasPhoto || photoDateValid);
  const status = isAuto ? 'Approved' : 'Pending';

  if (payload.gpsLat) detail.gpsLat = payload.gpsLat;
  if (payload.gpsLng) detail.gpsLng = payload.gpsLng;
  if (payload.photoTakenAt) detail.photoTakenAt = payload.photoTakenAt;

  appendObjectRow_(EHS.sheets.taskClaims, {
    Timestamp: now_(),
    SeasonId: seasonId,
    Pillar: task.Pillar,
    NIK: nik,
    Nama: user.nama,
    Divisi: user.divisi,
    TaskId: task.TaskId,
    ReferenceId: referenceId,
    Status: status,
    Points: points,
    Note: task.Title,
    PeriodKey: availability.periodKey,
    FrequencyType: availability.frequencyType,
    NextAvailableAt: getNextAvailableLabel_(availability.frequencyType),
    BuktiUrl: buktiUrl,
    Detail: JSON.stringify(detail)
  });

  if (isAuto) {
    awardPoints_(user, task.Pillar, task.TaskId, referenceId, points, task.Title, task.CampaignId, seasonId);
  }

  return {
    ok: true,
    status: status,
    points: points,
    message: isAuto
      ? 'Task berhasil diklaim (+' + points + ' poin).'
      : 'Data terkirim, menunggu verifikasi Admin EHS. Poin (+' + points + ') masuk setelah disetujui. Pastikan tanggal foto diambil hari ini.'
  };
}

function isSameDay_(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
    date1.getMonth() === date2.getMonth() &&
    date1.getDate() === date2.getDate();
}

/**
 * Admin approve klaim Health/Energy yang statusnya Pending -> poin masuk ledger.
 */
function approveHealthClaim(payload) {
  assertCapability_(payload.adminNik, 'canApproveAllReports');
  validateRequired_(payload, ['referenceId']);

  const sh = getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  const rowIndex = data.findIndex(function(r) { return clean_(r[col.ReferenceId]) === clean_(payload.referenceId); });
  if (rowIndex === -1) throw new Error('Klaim tidak ditemukan.');
  const row = data[rowIndex];
  if (clean_(row[col.Status]) !== 'Pending') throw new Error('Klaim sudah direview sebelumnya.');

  sh.getRange(rowIndex + 1, col.Status + 1).setValue('Approved');

  awardPoints_(
    { nik: row[col.NIK], nama: row[col.Nama], divisi: row[col.Divisi] },
    row[col.Pillar], row[col.TaskId], clean_(row[col.ReferenceId]), Number(row[col.Points]),
    row[col.Note], '', clean_(row[col.SeasonId])
  );

  return { ok: true, message: 'Klaim disetujui, +' + row[col.Points] + ' poin dicatat.' };
}

function reviseHealthClaim(payload) {
  assertCapability_(payload.adminNik, 'canApproveAllReports');
  validateRequired_(payload, ['referenceId', 'feedback']);

  const sh = getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  const rowIndex = data.findIndex(function(r) { return clean_(r[col.ReferenceId]) === clean_(payload.referenceId); });
  if (rowIndex === -1) throw new Error('Klaim tidak ditemukan.');

  sh.getRange(rowIndex + 1, col.Status + 1).setValue('Revise');
  const existingNote = clean_(data[rowIndex][col.Note]);
  sh.getRange(rowIndex + 1, col.Note + 1).setValue(existingNote + ' | Revisi Admin: ' + payload.feedback);

  return { ok: true, message: 'Klaim dikembalikan untuk revisi.' };
}

/**
 * Antrian klaim Health/Energy yang Pending, untuk Admin review.
 */
function listHealthClaimsForReview(payload) {
  assertCapability_(payload.nik, 'canApproveAllReports');
  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims))
    .filter(function(r) { return clean_(r.Status) === 'Pending'; });
  rows.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  return rows;
}

function getMyTaskClaims(payload) {
  validateRequired_(payload, ['nik', 'pillar']);
  const nik = normalizeNik_(payload.nik);
  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims))
    .filter(function(r) { return normalizeNikLenient_(r.NIK) === nik && clean_(r.Pillar) === payload.pillar; });
  rows.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  return rows.slice(0, 20);
}