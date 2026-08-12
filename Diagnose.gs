function TEST_isolasi_total() {
  try {
    const ss = SpreadsheetApp.openById('1LYgQ7AUDuUAcrnMtpoMTzENoGQCbnr_2qvqNnSILl4Q');
    Logger.log('BERHASIL. Nama: ' + ss.getName());
  } catch (e) {
    Logger.log('ERROR TERTANGKAP: ' + e.message);
  }
}

function TEST_reset_users() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName('02_Master_Users');
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  }
  Logger.log('Sheet 02_Master_Users dikosongkan, siap di-seed ulang.');
}

function TEST_claim() {
  const result = claimTask({ nik: '00180', taskId: 'H01' });
  Logger.log(JSON.stringify(result, null, 2));
}

function TEST_reset_claims_and_ledger() {
  const ss = getSpreadsheet_();
  ['11_DB_TaskClaims', '10_DB_PointsLedger'].forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    }
  });
  Logger.log('11_DB_TaskClaims dan 10_DB_PointsLedger dikosongkan, siap dites ulang dari nol.');
}

function TEST_reset_bmi() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName('12_DB_BMI_Records');
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  }
  Logger.log('12_DB_BMI_Records dikosongkan.');
}

function TEST_bmi_flow() {
  // Pengukuran pertama -> baseline, 0 poin (belum ada pembanding).
  const r1 = submitBmiRecord({ nik: '00180', tinggiCm: 165, beratKg: 80, lingkarPinggangCm: 95 });
  Logger.log('Pengukuran 1 (baseline): ' + JSON.stringify(r1));
  // -> kategori kemungkinan "Obesitas I" atau sejenis dari BMI 80/1.65^2 ≈ 29.4 -> Obesitas I
}

//  Bikin 1 user Kabag/Supervisor untuk testing
function TEST_tambah_kabag_demo() {
  const result = addOrPromoteSupervisor({
    adminNik: '99999',
    nik: '50001',
    nama: 'Kabag Demo',
    divisi: 'Production',
    divisiDiawasi: ['Production']
  });
  Logger.log(JSON.stringify(result));
}

function TEST_migrate_taskclaims_schema() {
  const sh = getSpreadsheet_().getSheetByName('11_DB_TaskClaims');
  const headers = ['Timestamp', 'SeasonId', 'Pillar', 'NIK', 'Nama', 'Divisi', 'TaskId',
    'ReferenceId', 'Status', 'Points', 'Note', 'PeriodKey', 'FrequencyType',
    'NextAvailableAt', 'BuktiUrl', 'Detail'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  Logger.log('Header 11_DB_TaskClaims diperbarui.');
}


function TEST_reset_group_survey() {
  const sh = getSpreadsheet_().getSheetByName('17_DB_GroupSurveySubmissions');
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  }
  Logger.log('17_DB_GroupSurveySubmissions dikosongkan.');
}

function TEST_add_column_isregistered() {
  const sh = getSpreadsheet_().getSheetByName('02_Master_Users');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  if (headers.indexOf('IsRegistered') !== -1) {
    Logger.log('Kolom IsRegistered sudah ada, tidak perlu ditambah lagi.');
    return;
  }

  // Sisipkan kolom baru tepat setelah "DivisiDiawasi", sebelum "Active" —
  // supaya urutan kolom persis sama dengan EHS_SCHEMA yang baru.
  const activeCol = headers.indexOf('Active') + 1; // 1-indexed
  sh.insertColumnBefore(activeCol);
  sh.getRange(1, activeCol).setValue('IsRegistered');

  Logger.log('Kolom IsRegistered berhasil ditambahkan di posisi kolom ' + activeCol);
}

function TEST_migrate_isregistered() {
  const sh = getSpreadsheet_().getSheetByName('02_Master_Users');
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const col = headers.indexOf('IsRegistered');
  for (let i = 1; i < data.length; i++) {
    sh.getRange(i + 1, col + 1).setValue('Yes'); // semua akun existing dianggap sudah teregistrasi
  }
  Logger.log('Migrasi IsRegistered selesai.');
}

function TEST_add_column_obligation() {
  const sh = getSpreadsheet_().getSheetByName('03_Master_Task');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('ObligationLevel') !== -1) { Logger.log('Sudah ada.'); return; }
  const insertBefore = headers.indexOf('Status') + 1;
  sh.insertColumnBefore(insertBefore);
  sh.getRange(1, insertBefore).setValue('ObligationLevel');
  // Default semua task lama jadi "Recommended" supaya tidak ada yang kosong
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertBefore, lastRow - 1, 1).setValue('Recommended');
  Logger.log('Kolom ObligationLevel ditambahkan.');
}

function TEST_add_column_profileinterests() {
  const sh = getSpreadsheet_().getSheetByName('02_Master_Users');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('ProfileInterests') !== -1) { Logger.log('Sudah ada.'); return; }
  const insertBefore = headers.indexOf('IsRegistered') + 1;
  sh.insertColumnBefore(insertBefore);
  sh.getRange(1, insertBefore).setValue('ProfileInterests');
  Logger.log('Kolom ProfileInterests ditambahkan.');
}

function TEST_migrate_task_reward_columns() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, '03_Master_Task', EHS_SCHEMA['03_Master_Task']);
  ensureSheet_(ss, '10_DB_PointsLedger', EHS_SCHEMA['10_DB_PointsLedger']);

  // Default kolom baru jadi 0 untuk baris yang sudah ada, biar tidak kosong/NaN
  ['03_Master_Task'].forEach(function(name) {
    const sh = ss.getSheetByName(name);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    ['DomainXP', 'CoinReward'].forEach(function(col) {
      const idx = headers.indexOf(col) + 1;
      const lastRow = sh.getLastRow();
      if (idx > 0 && lastRow > 1) {
        const range = sh.getRange(2, idx, lastRow - 1, 1);
        const values = range.getValues();
        const filled = values.map(function(r) { return [r[0] === '' ? 0 : r[0]]; });
        range.setValues(filled);
      }
    });
  });

  const ledgerSh = ss.getSheetByName('10_DB_PointsLedger');
  const ledgerHeaders = ledgerSh.getRange(1, 1, 1, ledgerSh.getLastColumn()).getValues()[0];
  ['DomainXP', 'Coin'].forEach(function(col) {
    const idx = ledgerHeaders.indexOf(col) + 1;
    const lastRow = ledgerSh.getLastRow();
    if (idx > 0 && lastRow > 1) {
      const range = ledgerSh.getRange(2, idx, lastRow - 1, 1);
      const values = range.getValues();
      range.setValues(values.map(function(r) { return [r[0] === '' ? 0 : r[0]]; }));
    }
  });

  Logger.log('Migrasi kolom reward selesai.');
}

function TEST_migrate_question_sheets() {
  const ss = getSpreadsheet_();
  ['07_Master_SurveyQuestions', '09_Master_ObsKelompok', '19_Master_ObsIndividu'].forEach(function(name) {
    ensureSheet_(ss, name, EHS_SCHEMA[name]);
  });
  Logger.log('Sheet pertanyaan (Survey/ObsKelompok/ObsIndividu) siap.');
}

function TEST_migrate_master_data_batch2() {
  const ss = getSpreadsheet_();
  ['20_Master_AwarenessContent', '21_Master_MissionTemplates', '22_Master_PointMapping',
   '23_Master_RewardCatalog', '24_Master_Clubs', '25_Master_Challenges',
   '26_Master_JourneyRules'].forEach(function(name) {
    ensureSheet_(ss, name, EHS_SCHEMA[name]);
  });
  ensureSheet_(ss, '01_Master_Seasons', EHS_SCHEMA['01_Master_Seasons']);
  ensureSheet_(ss, '06_Master_Campaigns', EHS_SCHEMA['06_Master_Campaigns']);
  ensureSheet_(ss, '05_Master_QuizBank', EHS_SCHEMA['05_Master_QuizBank']);
  Logger.log('Migrasi batch 2 (versi field lengkap) selesai.');
}
function TEST_migrate_reward_redemptions() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, '27_DB_RewardRedemptions', EHS_SCHEMA['27_DB_RewardRedemptions']);
  Logger.log('Sheet Reward Redemptions siap.');
}
function TEST_migrate_notifications() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, '28_DB_Notifications', EHS_SCHEMA['28_DB_Notifications']);
  Logger.log('Sheet Notifications siap.');
}