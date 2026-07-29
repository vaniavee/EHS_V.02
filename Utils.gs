// Membersihkan nilai input sebelum diproses lebih lanjut
function clean_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}
// Menghasilkan waktu saat ini berdasarkan server Google Apps Script
function now_() {
  return new Date();
}

/**
 * Normalisasi + validasi NIK.
 * Format yang diterima: 5-8 digit angka, opsional diawali 1-3 huruf kapital
 * (mis. "00180", "MJ00313", "MPI01252").
 * Menolak input yang jelas bukan NIK (mis. nama orang tertulis di sini) —
 */


// Melakukan normalisasi sekaligus validasi Nomor Induk Karyawan (NIK)
function normalizeNik_(raw) {
  const v = clean_(raw).toUpperCase().replace(/\s+/g, '');
  const pattern = /^[A-Z]{0,3}\d{5,8}$/;
  if (!pattern.test(v)) {
    throw new Error(
      'NIK "' + raw + '" tidak valid. Format yang benar: 5-8 digit angka, ' +
      'boleh diawali 1-3 huruf (contoh: 00180, MJ00313).'
    );
  }
  // Pad angka murni ke 5 digit, konsisten dengan data existing (00180, bukan 180)
  if (/^\d+$/.test(v) && v.length < 5) {
    return v.padStart(5, '0');
  }
  return v;
}

// Dipakai hanya untuk membaca & membandingkan data yang sudah ada di sheet (bukan validasi input baru).
function normalizeNikLenient_(raw) {
  const v = clean_(raw).toUpperCase().replace(/\s+/g, '');
  if (/^\d+$/.test(v) && v.length > 0 && v.length < 5) {
    return v.padStart(5, '0');
  }
  return v;
}

// Menentukan Season ID yang akan digunakan
function normalizeSeasonId_(v) {
  return clean_(v || getActiveSeason_().SeasonId);
}

function getHeaderIndex_(headers, candidates) {
  const names = Array.isArray(candidates) ? candidates : [candidates];
  for (let i = 0; i < names.length; i++) {
    const idx = headers.indexOf(names[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

// Membaca worksheet menjadi array objek
function readObjects_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  const headers = values.shift().map(clean_);
  return values
    .filter(function(row) { return row.some(function(v) { return clean_(v) !== ''; }); })
    .map(function(row) {
      const obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });
}

// Menambahkan satu baris data berdasarkan nama header
function appendObjectRow_(sheetName, valuesByHeader) {
  const sh = getSpreadsheet_().getSheetByName(sheetName);
  const headers = EHS_SCHEMA[sheetName];
  const targetRow = sh.getLastRow() + 1;

  headers.forEach(function(h, i) {
    if (h.toUpperCase().indexOf('NIK') !== -1) {
      sh.getRange(targetRow, i + 1).setNumberFormat('@');
    }
  });

  const row = headers.map(function(h) {
    return valuesByHeader.hasOwnProperty(h) ? valuesByHeader[h] : '';
  });

  sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
}

// Mengambil season yang sedang aktif.
function getActiveSeason_() {
  const seasons = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.seasons));
  const active = seasons.find(function(s) { return clean_(s.Status).toLowerCase() === 'active'; });
  if (!active) throw new Error('Tidak ada season aktif. Cek sheet ' + EHS.sheets.seasons);
  return active;
}

// Mengambil data task berdasarkan Task ID
function getTaskById_(taskId, seasonId) {
  const tasks = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.tasks));
  const sid = clean_(seasonId);
  return tasks.find(function(t) {
    return clean_(t.TaskId) === clean_(taskId) &&
           clean_(t.Status).toLowerCase() === 'active' &&
           (!clean_(t.SeasonId) || clean_(t.SeasonId) === sid);
  });
}

// Menghitung jumlah klaim pada suatu periode
function countClaimsInPeriod_(nik, seasonId, taskId, periodKey) {
  const sh = getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims);
  const rows = readObjects_(sh);
  return rows.filter(function(r) {
    return normalizeNikLenient_(r.NIK) === nik &&   // <- ganti dari normalizeNik_
           clean_(r.SeasonId) === seasonId &&
           clean_(r.TaskId) === taskId &&
           clean_(r.PeriodKey) === periodKey &&
           clean_(r.Status).toLowerCase() !== 'duplicate';
  }).length;
}

// Memeriksa apakah task masih dapat diklaim
function getTaskAvailability_(task, nik, seasonId) {
  const frequencyType = clean_(task.FrequencyType || 'season_once').toLowerCase();
  const limit = Number(task.FrequencyLimit || 1);
  const periodKey = getCurrentPeriodKey_(frequencyType);
  const used = countClaimsInPeriod_(nik, seasonId, task.TaskId, periodKey);

  return {
    available: used < limit,
    used: used,
    limit: limit,
    periodKey: periodKey,
    frequencyType: frequencyType,
    reason: used >= limit
      ? 'Sudah mencapai batas klaim periode ini. Tersedia lagi ' + getNextAvailableLabel_(frequencyType) + '.'
      : ''
  };
}

/**
 * Endpoint utama untuk klaim task Health & Energy.
 * payload: { nik, taskId, seasonId?, buktiUrl?, note? }
 */
// Memproses klaim aktivitas dan pemberian poin
function claimTask(payload) {
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
  const note = clean_(payload.note || task.Title);

  const points = awardPoints_(
    user, task.Pillar, task.TaskId, referenceId,
    Number(task.Points || 0), 'Klaim ' + task.Title, task.CampaignId, seasonId,
    Number(task.DomainXP || 0), Number(task.CoinReward || 0)
  );

  appendObjectRow_(EHS.sheets.taskClaims, {
    Timestamp: now_(),
    SeasonId: seasonId,
    Pillar: task.Pillar,
    NIK: nik,
    Nama: user.nama,
    Divisi: user.divisi,
    TaskId: task.TaskId,
    Status: points > 0 ? 'Claimed' : 'Duplicate',
    Points: points,
    Note: note,
    PeriodKey: availability.periodKey,
    FrequencyType: availability.frequencyType,
    NextAvailableAt: getNextAvailableLabel_(availability.frequencyType),
    BuktiUrl: clean_(payload.buktiUrl)
  });

  return {
    ok: true,
    message: points > 0 ? 'Task berhasil diklaim (+' + points + ' poin).' : 'Task sudah diklaim periode ini.',
    points: points, domainXp: Number(task.DomainXP || 0), coin: Number(task.CoinReward || 0)
  };
}

// Memastikan seluruh field wajib telah diisi
function validateRequired_(obj, fields) {
  fields.forEach(function(f) {
    if (!obj || clean_(obj[f]) === '') throw new Error('Field wajib: ' + f);
  });
}