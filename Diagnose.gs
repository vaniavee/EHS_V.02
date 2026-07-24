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