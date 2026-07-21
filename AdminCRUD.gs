/**
 * Admin CRUD Module — Kelola Master Data generik + Akun + Klaim Poin.
 * Semua operasi di sini butuh capability 'canManageMasterData' (Admin saja).
 */

/**
 * Definisi field per master type — tambahkan entri baru di sini
 * untuk membuka CRUD ke tabel master lain, tanpa nulis fungsi baru.
 */
function getMasterDefinition_(masterType) {
  const defs = {
    task: {
      sheet: EHS.sheets.tasks,
      idField: 'TaskId',
      fields: ['TaskId', 'Pillar', 'Category', 'Level', 'Title', 'Points',
               'Description', 'FrequencyType', 'FrequencyLimit', 'Validation',
               'Status', 'SeasonId', 'CampaignId']
    },
    quiz: {
      sheet: '05_Master_QuizBank',
      idField: 'QuizId',
      fields: ['QuizId', 'TaskId', 'Question', 'OptionA', 'OptionB', 'OptionC', 'OptionD',
               'CorrectOption', 'Explanation', 'Status']
    },
    campaign: {
      sheet: '06_Master_Campaigns',
      idField: 'CampaignId',
      fields: ['CampaignId', 'Title', 'MediaType', 'MediaUrl', 'Tagline',
               'Description', 'SubmissionMode', 'MinExposureSeconds',
               'SurveyPoints', 'SelfEvalPoints', 'Status', 'SeasonId']
    },
    faq: {
      sheet: '08_Master_FAQ',
      idField: 'FaqId',
      fields: ['FaqId', 'Category', 'Question', 'Answer', 'Status']
    }
  };
  const def = defs[masterType];
  if (!def) throw new Error('Master type tidak dikenal: ' + masterType);
  return def;
}

// --- READ (generik, dipakai tab Task & Campaign) ---
function getAdminMasterData(payload) {
  assertCapability_(payload.nik, 'canManageMasterData');
  const def = getMasterDefinition_(payload.masterType);
  return readObjects_(getSpreadsheet_().getSheetByName(def.sheet));
}

// --- CREATE / UPDATE (generik) ---
function saveMasterRecord(payload) {
  assertCapability_(payload.nik, 'canManageMasterData');
  const def = getMasterDefinition_(payload.masterType);
  const sh = getSpreadsheet_().getSheetByName(def.sheet);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf(def.idField);
  const idValue = clean_(payload.record[def.idField]);

  if (!idValue) throw new Error(def.idField + ' wajib diisi.');

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (clean_(data[i][idCol]) === idValue) { rowIndex = i; break; }
  }

  const rowValues = def.fields.map(function(f) {
    return payload.record.hasOwnProperty(f) ? payload.record[f] : '';
  });

  if (rowIndex === -1) {
    sh.appendRow(rowValues);
    return { ok: true, action: 'created', id: idValue };
  } else {
    sh.getRange(rowIndex + 1, 1, 1, rowValues.length).setValues([rowValues]);
    return { ok: true, action: 'updated', id: idValue };
  }
}

// --- DELETE (generik) ---
function deleteMasterRecord(payload) {
  assertCapability_(payload.nik, 'canManageMasterData');
  const def = getMasterDefinition_(payload.masterType);
  const sh = getSpreadsheet_().getSheetByName(def.sheet);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf(def.idField);
  const idValue = clean_(payload.id);

  for (let i = 1; i < data.length; i++) {
    if (clean_(data[i][idCol]) === idValue) {
      sh.deleteRow(i + 1);
      return { ok: true, action: 'deleted', id: idValue };
    }
  }
  throw new Error('Record tidak ditemukan: ' + idValue);
}

/**
 * Admin menambahkan/mempromosikan seseorang jadi Kabag/Supervisor.
 * Kalau NIK sudah terdaftar sebagai karyawan biasa -> hanya update flag
 * (TIDAK membuat identitas baru, Kabag tetap partisipan yang sama).
 * Kalau NIK belum terdaftar -> buat baru sekaligus dengan flag Supervisor.
 */
function addOrPromoteSupervisor(payload) {
  const admin = assertCapability_(payload.adminNik, 'canManageMasterData');
  validateRequired_(payload, ['nik', 'nama', 'divisi', 'divisiDiawasi']);

  const nik = normalizeNik_(payload.nik);
  const sh = getSpreadsheet_().getSheetByName(EHS.sheets.users);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const nikCol = headers.indexOf('NIK');

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (normalizeNikLenient_(data[i][nikCol]) === nik) { rowIndex = i; break; }
  }

  const record = {
    NIK: nik,
    Nama: clean_(payload.nama),
    Divisi: clean_(payload.divisi),
    No_WA: clean_(payload.noWa || (rowIndex >= 0 ? data[rowIndex][headers.indexOf('No_WA')] : '')),
    IsAdmin: rowIndex >= 0 ? data[rowIndex][headers.indexOf('IsAdmin')] : 'FALSE',
    IsSupervisor: 'TRUE',
    DivisiDiawasi: Array.isArray(payload.divisiDiawasi)
      ? payload.divisiDiawasi.join(',') : clean_(payload.divisiDiawasi),
    Active: 'Yes',
    TanggalDitambahkan: rowIndex >= 0 ? data[rowIndex][headers.indexOf('TanggalDitambahkan')] : now_(),
    DitambahkanOleh: admin.nama
  };

  const rowValues = headers.map(function(h) { return record[h]; });

  if (rowIndex === -1) {
    sh.appendRow(rowValues);
    return { ok: true, action: 'created_as_supervisor', nik: nik };
  } else {
    sh.getRange(rowIndex + 1, 1, 1, rowValues.length).setValues([rowValues]);
    return { ok: true, action: 'promoted_to_supervisor', nik: nik };
  }
}

/**
 * List semua user (dipakai tab Akun & Kabag di Admin Dashboard).
 */
function getAllUsers(payload) {
  assertCapability_(payload.nik, 'canManageMasterData');
  return readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.users));
}

/**
 * Tambah/edit user secara umum (bukan cuma Kabag) — dipakai tab Akun.
 * Bisa ubah field apapun termasuk IsAdmin/IsSupervisor/Active.
 */
function saveUserRecord(payload) {
  assertCapability_(payload.nik, 'canManageMasterData');
  validateRequired_(payload.record, ['NIK', 'Nama', 'Divisi']);

  const sh = getSpreadsheet_().getSheetByName(EHS.sheets.users);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const nikCol = headers.indexOf('NIK');
  const targetNik = normalizeNik_(payload.record.NIK);

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (normalizeNikLenient_(data[i][nikCol]) === targetNik) { rowIndex = i; break; }
  }

  const record = Object.assign({}, payload.record, { NIK: targetNik });

  // Jangan biarkan frontend menimpa IsRegistered saat edit — pertahankan nilai lama.
  if (rowIndex >= 0) {
    record.IsRegistered = data[rowIndex][headers.indexOf('IsRegistered')];
  } else if (!record.IsRegistered) {
    record.IsRegistered = 'No'; // default untuk akun baru
  }

  const rowValues = headers.map(function(h) {
    if (record.hasOwnProperty(h)) return record[h];
    return rowIndex >= 0 ? data[rowIndex][headers.indexOf(h)] : '';
  });

  if (rowIndex === -1) {
    sh.getRange(2, nikCol + 1, 999, 1).setNumberFormat('@');
    if (!record.TanggalDitambahkan) {
      const tglCol = headers.indexOf('TanggalDitambahkan');
      if (tglCol !== -1) rowValues[tglCol] = now_();
    }
    sh.appendRow(rowValues);
    return { ok: true, action: 'created' };
  }
  sh.getRange(rowIndex + 1, 1, 1, rowValues.length).setValues([rowValues]);
  return { ok: true, action: 'updated' };
}

/**
 * Hapus user dari 02_Master_Users. Tidak menghapus riwayat poin/klaim
 * lama milik user itu (biar audit trail tetap utuh) — cuma menghapus
 * baris identitasnya, sehingga user itu tidak bisa login lagi.
 */
function deleteUserRecord(payload) {
  assertCapability_(payload.nik, 'canManageMasterData');
  validateRequired_(payload, ['targetNik']);

  const sh = getSpreadsheet_().getSheetByName(EHS.sheets.users);
  const data = sh.getDataRange().getValues();
  const nikCol = data[0].indexOf('NIK');
  const targetNik = normalizeNik_(payload.targetNik);

  for (let i = 1; i < data.length; i++) {
    if (normalizeNikLenient_(data[i][nikCol]) === targetNik) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  throw new Error('User tidak ditemukan: ' + targetNik);
}

/**
 * List klaim task terbaru (Health & Energy) untuk keperluan koreksi manual
 * oleh Admin — dipakai tab Klaim. Dibatasi 100 baris terbaru supaya
 * tidak berat kalau datanya sudah banyak.
 */
function listAllTaskClaimsForAdmin(payload) {
  assertCapability_(payload.nik, 'canManageMasterData');
  let rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims));

  if (payload.nikFilter) {
    const target = normalizeNikLenient_(payload.nikFilter);
    rows = rows.filter(function(r) { return normalizeNikLenient_(r.NIK) === target; });
  }

  rows.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  return rows.slice(0, 100);
}

/**
 * Admin edit/koreksi poin SATU baris klaim (misal salah input).
 * Selisih poin dicatat sebagai entri baru di ledger (bukan edit angka lama),
 * supaya audit trail (siapa mengoreksi, kapan, kenapa) tetap utuh.
 */
function adjustTaskClaim(payload) {
  const admin = assertCapability_(payload.adminNik, 'canManageMasterData');
  validateRequired_(payload, ['referenceId', 'newPoints']);

  const claimSh = getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims);
  const claimData = claimSh.getDataRange().getValues();
  const claimHeaders = claimData[0];
  const cc = {};
  claimHeaders.forEach(function(h, i) { cc[h] = i; });

  const claimRowIdx = claimData.findIndex(function(r) { return clean_(r[cc.ReferenceId]) === clean_(payload.referenceId); });
  if (claimRowIdx === -1) throw new Error('Klaim tidak ditemukan.');

  const oldPoints = Number(claimData[claimRowIdx][cc.Points]);
  const newPoints = Number(payload.newPoints);

  claimSh.getRange(claimRowIdx + 1, cc.Points + 1).setValue(newPoints);

  const delta = newPoints - oldPoints;
  if (delta !== 0) {
    const row = claimData[claimRowIdx];
    appendObjectRow_(EHS.sheets.pointsLedger, {
      Timestamp: now_(),
      SeasonId: clean_(row[cc.SeasonId]),
      Pillar: clean_(row[cc.Pillar]),
      TaskId: clean_(row[cc.TaskId]),
      CampaignId: '',
      NIK: clean_(row[cc.NIK]),
      Nama: clean_(row[cc.Nama]),
      Divisi: clean_(row[cc.Divisi]),
      ReferenceId: payload.referenceId + ':KOREKSI:' + now_().getTime(),
      Points: delta,
      Note: 'Koreksi oleh ' + admin.nama + (payload.reason ? ' — ' + payload.reason : '')
    });
  }
  return { ok: true, message: 'Poin dikoreksi dari ' + oldPoints + ' menjadi ' + newPoints + '.' };
}

/**
 * Admin hapus SATU baris klaim (misal submit ganda/keliru total).
 * Poin yang sudah masuk otomatis "dibatalkan" lewat entri koreksi negatif
 * di ledger, bukan menghapus baris ledger lama.
 */
function deleteTaskClaim(payload) {
  const admin = assertCapability_(payload.adminNik, 'canManageMasterData');
  validateRequired_(payload, ['referenceId']);

  const sh = getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims);
  const data = sh.getDataRange().getValues();
  const col = {};
  data[0].forEach(function(h, i) { col[h] = i; });

  const rowIdx = data.findIndex(function(r) { return clean_(r[col.ReferenceId]) === clean_(payload.referenceId); });
  if (rowIdx === -1) throw new Error('Klaim tidak ditemukan.');
  const row = data[rowIdx];
  const points = Number(row[col.Points]);

  sh.deleteRow(rowIdx + 1);

  if (points > 0) {
    appendObjectRow_(EHS.sheets.pointsLedger, {
      Timestamp: now_(),
      SeasonId: clean_(row[col.SeasonId]),
      Pillar: clean_(row[col.Pillar]),
      TaskId: clean_(row[col.TaskId]),
      CampaignId: '',
      NIK: clean_(row[col.NIK]),
      Nama: clean_(row[col.Nama]),
      Divisi: clean_(row[col.Divisi]),
      ReferenceId: payload.referenceId + ':HAPUS:' + now_().getTime(),
      Points: -points,
      Note: 'Klaim dihapus oleh ' + admin.nama
    });
  }
  return { ok: true, message: 'Klaim dihapus, poin ' + points + ' dibatalkan dari ledger.' };
}

function getCampaignOptions(payload) {
  assertCapability_(payload.nik, 'canManageMasterData');
  return readObjects_(getSpreadsheet_().getSheetByName('06_Master_Campaigns'))
    .map(function(c) { return { id: c.CampaignId, title: c.Title }; });
}