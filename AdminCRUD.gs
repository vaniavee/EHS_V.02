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
      fields: ['TaskId', 'Pillar', 'Category', 'Level', 'Title', 'Points', 'DomainXP', 'CoinReward',
               'Description', 'FrequencyType', 'FrequencyLimit', 'Validation', 'ApprovedBy', 'FormType',
               'ObligationLevel', 'Status', 'SeasonId', 'CampaignId', 'GroupTag']
    },
    quiz: {
      sheet: '05_Master_QuizBank',
      idField: 'QuizId',
      fields: EHS_SCHEMA['05_Master_QuizBank']
    },
    campaign: {
      sheet: '06_Master_Campaigns',
      idField: 'CampaignId',
      fields: EHS_SCHEMA['06_Master_Campaigns']
    },
    faq: {
      sheet: '08_Master_FAQ',
      idField: 'FaqId',
      fields: ['FaqId', 'Category', 'Question', 'Answer', 'Status']
    },
    department: {
      sheet: '18_Master_Departments',
      idField: 'DepartmentId',
      fields: ['DepartmentId', 'Name', 'Note']
    },
    survey: {
      sheet: '07_Master_SurveyQuestions',
      idField: 'QuestionId',
      fields: ['QuestionId', 'Dimension', 'QuestionText', 'Scale', 'EvaluationLevel', 'Order', 'Status']
    },
    obskelompok: {
      sheet: '09_Master_ObsKelompok',
      idField: 'QuestionId',
      fields: ['QuestionId', 'Dimension', 'QuestionText', 'Scale', 'EvaluationLevel', 'Order', 'Status']
    },
    obsindividu: {
      sheet: '19_Master_ObsIndividu',
      idField: 'QuestionId',
      fields: ['QuestionId', 'Dimension', 'QuestionText', 'Scale', 'EvaluationLevel', 'Order', 'Status']
    },
    awareness: { 
      sheet: '20_Master_AwarenessContent', 
      idField: 'ContentId', 
      fields: EHS_SCHEMA['20_Master_AwarenessContent'] 
    },
    missiontemplate: { 
      sheet: '21_Master_MissionTemplates', 
      idField: 'MissionId', 
      fields: EHS_SCHEMA['21_Master_MissionTemplates'] 
    },
    pointmapping: {
      sheet: '22_Master_PointMapping', 
      idField: 'MappingId', 
      fields: EHS_SCHEMA['22_Master_PointMapping'] 
    },
    rewardcatalog: { 
      sheet: '23_Master_RewardCatalog', 
      idField: 'RewardId', 
      fields: EHS_SCHEMA['23_Master_RewardCatalog'] 
      },
    club: { 
      sheet: '24_Master_Clubs', 
      idField: 'ClubId', 
      fields: EHS_SCHEMA['24_Master_Clubs'] 
    },
    challenge: { 
      sheet: '25_Master_Challenges', 
      idField: 'ChallengeId', 
      fields: EHS_SCHEMA['25_Master_Challenges'] 
    },
    journeyrule: { 
      sheet: '26_Master_JourneyRules', 
      idField: 'RuleId', 
      fields: EHS_SCHEMA['26_Master_JourneyRules'] 
    },
    season: { 
      sheet: '01_Master_Seasons', 
      idField: 'SeasonId', 
      fields: EHS_SCHEMA['01_Master_Seasons'] 
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

  // Auto-timestamp UpdatedAt kalau kolom ini ada di sheet — abaikan apapun yang dikirim frontend.
  if (headers.indexOf('UpdatedAt') !== -1) {
    payload.record.UpdatedAt = now_();
  }

  // PENTING: nulis berdasarkan HEADER ASLI SHEET (headers), bukan def.fields.
  // def.fields cuma dipakai untuk hal lain (mis. validasi/dokumentasi) — kalau
  // isinya gak persis sama urutan kolom fisik sheet, nulis pakai def.fields
  // bakal nggeser semua kolom setelah titik yang beda. Nulis berdasarkan
  // `headers` (dibaca langsung dari sheet di atas) selalu align, apapun
  // urutan kolomnya, dan otomatis ngikut kalau ada kolom baru ditambah di sheet.
  const rowValues = headers.map(function(h) {
    return payload.record.hasOwnProperty(h) ? payload.record[h] : '';
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

function bulkSaveMasterRecords(payload) {
  assertCapability_(payload.nik, 'canManageMasterData');
  const def = getMasterDefinition_(payload.masterType);
  const sh = getSpreadsheet_().getSheetByName(def.sheet);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf(def.idField);

  let countCreated = 0;
  let countUpdated = 0;

  const existingIndex = {};
  for (let i = 1; i < data.length; i++) {
    const key = clean_(data[i][idCol]);
    if (key) existingIndex[key] = i;
  }

  payload.records.forEach(function(record) {
    const idValue = clean_(record[def.idField]);
    if (!idValue) return;
    const rowValues = headers.map(function(h) {
      return record.hasOwnProperty(h) ? record[h] : '';
    });

    if (existingIndex[idValue] === undefined) {
      sh.appendRow(rowValues);
      countCreated++;
    } else {
      const rowIndex = existingIndex[idValue];
      sh.getRange(rowIndex + 1, 1, 1, rowValues.length).setValues([rowValues]);
      countUpdated++;
    }
  });

  return { ok: true, countCreated: countCreated, countUpdated: countUpdated };
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
  if (nikCol === -1) throw new Error('Kolom NIK tidak ditemukan di sheet ' + EHS.sheets.users + '.');
  const targetNik = normalizeNik_(payload.record.NIK);

  sh.getRange(2, nikCol + 1, Math.max(sh.getLastRow() - 1, 1), 1).setNumberFormat('@');

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
// function bulkImportQuizFromText(payload) {
//   assertCapability_(payload.nik, 'canManageMasterData');
//   validateRequired_(payload, ['format', 'content']);

//   let records;
//   if (payload.format === 'json') {
//     records = JSON.parse(payload.content);
//   } else if (payload.format === 'csv') {
//     records = parseDelimitedQuiz_(payload.content, ',');
//   } else if (payload.format === 'txt') {
//     records = parseDelimitedQuiz_(payload.content, '\t');
//   } else {
//     throw new Error('Format tidak dikenal: ' + payload.format);
//   }

//   if (!Array.isArray(records) || !records.length) throw new Error('Tidak ada data soal yang valid untuk diimport.');

//   records.forEach(function(r, i) {
//     if (!clean_(r.QuizId) || !clean_(r.CampaignId) || !clean_(r.QuestionText)) {
//       throw new Error('Baris ' + (i + 2) + ': QuizId, CampaignId, dan QuestionText wajib diisi.');
//     }
//   });

//   return bulkSaveMasterRecords({ nik: payload.nik, masterType: 'quiz', records: records });
// }

// Header wajib (urutan bebas, harus persis nama field schema):
// SeasonId,QuizId,CampaignId,QuizScope,PeriodType,Difficulty,QuestionText,OptionA,OptionB,OptionC,OptionD,CorrectOption,Explanation,PointMappingId,Points,SortOrder,ShuffleOptions,Status,Source
function parseDelimitedQuiz_(content, delimiter) {
  const lines = content.split(/\r?\n/).filter(function(l) { return l.trim() !== ''; });
  if (lines.length < 2) return [];
  const headers = lines[0].split(delimiter).map(function(h) { return h.trim(); });
  return lines.slice(1).map(function(line) {
    const cols = line.split(delimiter);
    const obj = {};
    headers.forEach(function(h, i) { obj[h] = clean_(cols[i]); });
    return obj;
  });
}

// // Excel (.xlsx) — dikonversi via Google Drive API (Advanced Service).
// // WAJIB aktifkan dulu: Apps Script editor -> Services (ikon +) -> cari "Drive API" -> Add.
// function bulkImportQuizFromExcel(payload) {
//   assertCapability_(payload.nik, 'canManageMasterData');
//   validateRequired_(payload, ['base64Data']);

//   const blob = Utilities.newBlob(Utilities.base64Decode(payload.base64Data), MimeType.MICROSOFT_EXCEL, 'quiz_import.xlsx');
//   const tempFile = DriveApp.createFile(blob);

//   let convertedFile;
//   try {
//     convertedFile = Drive.Files.copy({ title: 'quiz_import_temp_' + Date.now() }, tempFile.getId(), { convert: true });
//   } catch (e) {
//     tempFile.setTrashed(true);
//     throw new Error('Gagal konversi Excel. Pastikan "Drive API" sudah diaktifkan di Services (Apps Script editor). Detail: ' + e.message);
//   }

//   const ss = SpreadsheetApp.openById(convertedFile.id);
//   const sh = ss.getSheets()[0];
//   const data = sh.getDataRange().getValues();
//   const headers = data[0].map(function(h) { return clean_(h); });

//   const records = data.slice(1)
//     .filter(function(row) { return row.some(function(v) { return clean_(v) !== ''; }); })
//     .map(function(row) {
//       const obj = {};
//       headers.forEach(function(h, i) { obj[h] = clean_(row[i]); });
//       return obj;
//     });

//   DriveApp.getFileById(convertedFile.id).setTrashed(true);
//   tempFile.setTrashed(true);

//   if (!records.length) throw new Error('Tidak ada baris data ditemukan di file Excel.');
//   records.forEach(function(r, i) {
//     if (!clean_(r.QuizId) || !clean_(r.CampaignId) || !clean_(r.QuestionText)) {
//       throw new Error('Baris ' + (i + 2) + ': QuizId, CampaignId, dan QuestionText wajib diisi.');
//     }
//   });

//   return bulkSaveMasterRecords({ nik: payload.nik, masterType: 'quiz', records: records });
// }