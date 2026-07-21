/**
 * EHS App — Environment, Health & Safety Gamification
 * Core Engine v1
 * Melakukan proses inisialisasi sistem, pembuatan struktur database pada Google Spreadsheet, serta pengisian data master awal sehingga aplikasi dapat langsung digunakan
 */


// Berfungsi menyimpan ID Google Spreadsheet yang digunakan sebagai database utama aplikas
const EHS_SPREADSHEET_ID = '1LYgQ7AUDuUAcrnMtpoMTzENoGQCbnr_2qvqNnSILl4Q';

// Membuka dan mengembalikan objek Google Spreadsheet berdasarkan ID yang telah didefinisikan.
function getSpreadsheet_() {
  return SpreadsheetApp.openById(EHS_SPREADSHEET_ID);
}

// Mendefinisikan ID Spreadsheet tujuan, nama-nama sheet yang digunakan, konfigurasi zona waktu default (Asia/Jakarta)
const EHS = {
  timezone: 'Asia/Jakarta',
  adminEmployeeId: '99999',
  sheets: {
    config: '00_Config',
    seasons: '01_Master_Seasons',
    users: '02_Master_Users',
    tasks: '03_Master_Task',
    bmiRules: '04_Master_BMI_ScoringRule',
    pointsLedger: '10_DB_PointsLedger',
    taskClaims: '11_DB_TaskClaims',
    bmiRecords: '12_DB_BMI_Records',
    safetyReports: '13_DB_SafetyReports',
    helpRequests: '14_DB_HelpRequests'
  }
};

//  Mendefinisikan skema susunan kolom (headers) untuk setiap sheet database agar terstruktur dengan konsisten.
const EHS_SCHEMA = {
  '00_Config': ['Key', 'Value', 'Note'],
  '01_Master_Seasons': ['SeasonId', 'SeasonName', 'StartDate', 'EndDate', 'Status'],
  '02_Master_Users': [
    'NIK', 'Nama', 'Divisi', 'No_WA',
    'IsAdmin', 'IsSupervisor', 'DivisiDiawasi',
    'IsRegistered','Active', 'TanggalDitambahkan', 'DitambahkanOleh'
  ],
  '03_Master_Task': [
    'TaskId', 'Pillar', 'Category', 'Level', 'Title', 'Points',
    'Description', 'FrequencyType', 'FrequencyLimit', 'Validation',
    'Status', 'SeasonId', 'CampaignId'
  ],
  '04_Master_BMI_ScoringRule': ['KategoriLama', 'KategoriBaru', 'KriteriaPoin', 'Poin'],
  '10_DB_PointsLedger': [
    'Timestamp', 'SeasonId', 'Pillar', 'TaskId', 'CampaignId',
    'NIK', 'Nama', 'Divisi', 'ReferenceId', 'Points', 'Note'
  ],
  '05_Master_QuizBank': [
  'QuizId', 'TaskId', 'Question', 'OptionA', 'OptionB', 'OptionC', 'OptionD',
  'CorrectOption', 'Explanation', 'Status'
  ],
  '06_Master_Campaigns': [
    'CampaignId', 'Title', 'MediaType', 'MediaUrl', 'Tagline', 'Description',
    'SubmissionMode', 'MinExposureSeconds', 'SurveyPoints', 'SelfEvalPoints',
    'Status', 'SeasonId'
  ],
  '07_Master_SurveyQuestions': ['QuestionId', 'QuestionText', 'Dimension', 'Order', 'Status'
  ],
  '08_Master_FAQ': ['FaqId', 'Category', 'Question', 'Answer', 'Status'
  ],
  '11_DB_TaskClaims': [
    'Timestamp', 'SeasonId', 'Pillar', 'NIK', 'Nama', 'Divisi', 'TaskId',
    'ReferenceId', 'Status', 'Points', 'Note', 'PeriodKey', 'FrequencyType',
    'NextAvailableAt', 'BuktiUrl', 'Detail'
  ],
  '12_DB_BMI_Records': [
    'Timestamp', 'NIK', 'Nama', 'Divisi', 'Tanggal', 'Tinggi_cm', 'Berat_kg',
    'Lingkar_Pinggang_cm', 'BMI', 'Kategori_BMI', 'Kategori_BMI_Sebelumnya',
    'Kriteria_Poin', 'Points', 'PeriodeKey'
  ],
  '13_DB_SafetyReports': [
    'Timestamp', 'SeasonId', 'ReferenceId', 'SupervisorNik', 'SupervisorNama',
    'DivisiDilaporkan', 'JenisLaporan', 'Deskripsi', 'BuktiUrl', 'Severity',
    'Status', 'AdminFeedback', 'Points', 'PeriodKey', 'ReviewedBy', 'ReviewedAt'
  ],
  '14_DB_HelpRequests': [
    'Timestamp', 'RequestId', 'NIK', 'Nama', 'Question', 'Status',
    'Answer', 'AnsweredAt', 'AnsweredBy'
  ],
  // SubmissionMode: Individu | Kelompok | Keduanya
  '16_DB_MiniProjects': [
    'Timestamp', 'ReferenceId', 'SeasonId', 'NIK', 'Nama', 'Divisi',
    'JudulProject', 'AreaKerja', 'DeskripsiMasalah', 'TindakanPerbaikan',
    'FotoBeforeUrl', 'FotoAfterUrl', 'EstimasiDampak', 'AnggotaTim',
    'Status', 'Points', 'AdminFeedback', 'PeriodKey'
  ],
  '17_DB_GroupSurveySubmissions': [
    'Timestamp', 'ReferenceId', 'SeasonId', 'CampaignId', 'NIK_Pelapor',
    'Nama_Pelapor', 'AnggotaKelompok', 'Divisi', 'Jawaban', 'Points', 'Status'
  ]
};

// Fungsi utama untuk menginisialisasiseluruh sistem ketika aplikasi pertama kali dipasang
function setup() {
  const ss = getSpreadsheet_();
  Object.keys(EHS_SCHEMA).forEach(function(name) {
    ensureSheet_(ss, name, EHS_SCHEMA[name]);
  });
  seedInitialMasterData_(ss);
}

// Fungsi untuk memastikan worksheet telah tersedia dengan struktur kolom yang benar.
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);

    // Cari kolom NIK otomatis, paksa formatnya jadi Teks sejak awal.
  const nikColIndex = headers.indexOf('NIK');
  if (nikColIndex !== -1) {
    sh.getRange(2, nikColIndex + 1, 999, 1).setNumberFormat('@');
  }
}
  return sh;
}

// Fungsi untuk melakukan pengujian apakah Google Apps Script berhasil terhubung dengan spreadsheet.
function TEST_koneksi() {
  const ss = getSpreadsheet_();
  if (ss === null || ss === undefined) {
    Logger.log('GAGAL: getSpreadsheet_() mengembalikan null/undefined.');
    return;
  }
  Logger.log('BERHASIL terhubung ke: ' + ss.getName());
  Logger.log('URL: ' + ss.getUrl());
}


/**
 * Isi data awal supaya EHS App langsung bisa dites tanpa isi manual.
 * Idempotent — cek dulu apakah sheet sudah ada isinya sebelum menambah,
 * supaya run setup() berkali-kali tidak bikin baris dobel.
 */

// Menjalankan seluruh proses pengisian data master awal
// Jika data sudah ada maka proses tidak akan menambahkan data yang sama kembali.
function seedInitialMasterData_(ss) {
  seedConfig_(ss);
  seedSeason_(ss);
  seedAdminUser_(ss);
  seedSampleTasks_(ss);
  seedBmiScoringRules_(ss);
  seedSurveyQuestions_(ss);
  seedFaqData_(ss);
}

// Mengisi konfigurasi aplikasi pada worksheet 00_Config
function seedConfig_(ss) {
  const sh = ss.getSheetByName('00_Config');
  if (sh.getLastRow() > 1) return; // sudah pernah diisi

  const rows = [
    ['AppName', 'EHS App', 'Nama program: Environment, Health & Safety'],
    ['CurrentSeasonId', 'S01', 'Season aktif default'],
    ['AdminEmployeeId', '99999', 'NIK khusus admin'],
    ['Timezone', 'Asia/Jakarta', 'Zona waktu']
  ];
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

// Membuat season pertama
function seedSeason_(ss) {
  const sh = ss.getSheetByName('01_Master_Seasons');
  if (sh.getLastRow() > 1) return;

  sh.getRange(2, 1, 1, 5).setValues([[
    'S01', 'Season 01', new Date('2026-01-01'), new Date('2026-12-31'), 'Active'
  ]]);
}

// Menambahkan akun administrator dan pengguna contoh.
function seedAdminUser_(ss) {
  const sh = ss.getSheetByName('02_Master_Users');
  if (sh.getLastRow() > 1) return;

  // Paksa kolom A (NIK) selalu format Teks, supaya "00180" tidak
  // pernah dikonversi jadi angka 180 oleh Google Sheets.
  sh.getRange('A2:A1000').setNumberFormat('@');

  const rows = [
    ['99999', 'Admin EHS', 'HSE', '', 'TRUE', 'FALSE', '', 'Yes', 'Yes', new Date(), 'System'],
    ['00180', 'Karyawan Demo', 'Production', '', 'FALSE', 'FALSE', '', 'Yes', 'Yes', new Date(), 'System']
  ];
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

// Mengisi contoh aktivitas (task) yang dapat diklaim peserta
function seedSampleTasks_(ss) {
  const sh = ss.getSheetByName('03_Master_Task');
  if (sh.getLastRow() > 1) return;

  // Kolom: TaskId, Pillar, Category, Level, Title, Points, Description,
  //        FrequencyType, FrequencyLimit, Validation, Status, SeasonId, CampaignId
  const rows = [
    ['H01', 'Health', 'Konimex Move', 'Easy', 'Konimex Move Harian', 5,
     'Submit status hidrasi & nutrisi harian', 'daily', 1, 'auto', 'Active', 'S01', ''],

    ['H02', 'Health', 'TPP Exercise', 'Medium', 'TPP Exercise', 20,
     'Latihan fisik terjadwal perusahaan dengan bukti', 'weekly', 3, 'semiauto', 'Active', 'S01', ''],

    ['H03', 'Health', 'Exercise Mandiri', 'Medium', 'Exercise Mandiri', 20,
     'Olahraga bebas di luar jam kerja', 'weekly', 3, 'semiauto', 'Active', 'S01', ''],

    ['H04', 'Health', 'Konimex Walk Challenge', 'Easy', 'Konimex Walk Challenge', 15,
     'Poin dikali jumlah putaran saat klaim (lihat catatan)', 'daily', 5, 'semiauto', 'Active', 'S01', ''],

    ['E01', 'Energy', 'Awareness', 'Easy', 'Akses Media Awareness dan Isi Survey', 2,
     'Buka media awareness aktif dan isi survey singkat', 'campaign_once', 1, 'auto', 'Active', 'S01', ''],

    ['E02', 'Energy', 'Action', 'Medium', 'Submit Potensi Inefisiensi Energi', 5,
     'Laporkan potensi pemborosan energi di area kerja', 'weekly', 2, 'verifier', 'Active', 'S01', ''],

    ['E03', 'Energy', 'Improvement', 'Hard', 'Mini Project Improvement Area', 15,
     'Perbaikan kecil dengan bukti before-after', 'monthly', 1, 'panel', 'Active', 'S01', '']
  ];
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

// Mengisi aturan pemberian poin berdasarkan perubahan kategori BMI
function seedBmiScoringRules_(ss) {
  const sh = ss.getSheetByName('04_Master_BMI_ScoringRule');
  if (sh.getLastRow() > 1) return;

  // Kolom: KategoriLama, KategoriBaru, KriteriaPoin, Poin
  // Sesuai requirement Vee:
  //  - Turun 1 level BMI                          -> 300 poin
  //  - Bertahan di level Normal (tetap Normal)     -> 350 poin
  //  - Bertahan di level baru yang lebih sehat
  //    (sudah turun level, lalu tidak naik lagi)   -> 300 poin
  const rows = [
    ['Obesitas II', 'Obesitas I', 'turun_1_level', 300],
    ['Obesitas I', 'Overweight', 'turun_1_level', 300],
    ['Overweight', 'Normal', 'turun_1_level', 300],
    ['Underweight', 'Normal', 'turun_1_level', 300],

    ['Normal', 'Normal', 'bertahan_normal', 350],

    // "Bertahan di level baru yang lebih sehat" — kategori BMI dua kali
    // pengukuran terakhir SAMA, dan kategori itu BUKAN kategori awal
    // sebelum penurunan. Ini butuh histori 2 langkah, ditangani khusus
    // di kode (lihat resolveBmiKriteria_ di bawah), bukan lookup 1 baris.
    ['Obesitas I', 'Obesitas I', 'bertahan_level_baru', 300],
    ['Overweight', 'Overweight', 'bertahan_level_baru', 300],
    ['Underweight', 'Underweight', 'bertahan_level_baru', 300]
  ];
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedSurveyQuestions_(ss) {
  const sh = ss.getSheetByName('07_Master_SurveyQuestions');
  if (sh.getLastRow() > 1) return;
  const rows = [
    ['S1', 'Media awareness ini menarik perhatian saya.', 'Atensi', 1, 'Active'],
    ['S2', 'Isi atau pesan media ini mudah saya pahami.', 'Kemudahan Pemahaman', 2, 'Active'],
    ['S3', 'Media ini relevan dengan pekerjaan atau area kerja saya.', 'Relevansi Kerja', 3, 'Active'],
    ['S4', 'Setelah melihat media ini, saya tahu tindakan kecil yang bisa dilakukan.', 'Kesiapan Bertindak', 4, 'Active'],
    ['S5', 'Media ini membuat saya lebih terdorong memperhatikan topik ini.', 'Motivasi', 5, 'Active']
  ];
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedFaqData_(ss) {
  const sh = ss.getSheetByName('08_Master_FAQ');
  if (sh.getLastRow() > 1) return;
  const rows = [
    ['FAQ01', 'Login', 'Bagaimana cara login ke EHS App?', 'Masukkan NIK Anda di halaman login, lalu tekan Masuk. Tidak perlu password. Kalau NIK belum terdaftar, hubungi Admin EHS untuk didaftarkan terlebih dahulu.', 'Active'],
    ['FAQ02', 'Login', 'NIK saya tidak dikenali, apa yang harus dilakukan?', 'Pastikan format NIK sudah benar (contoh: 00180 atau MJ00313). Kalau masih gagal, kemungkinan NIK Anda belum didaftarkan Admin EHS — hubungi Admin lewat WhatsApp di halaman login.', 'Active'],
    ['FAQ03', 'Task', 'Kenapa tombol klaim task tertulis "Kuota Habis"?', 'Setiap task punya batas klaim per periode (harian/mingguan/bulanan). Kalau kuota sudah habis, tombol akan menampilkan kapan kuota tersedia lagi.', 'Active'],
    ['FAQ04', 'Task', 'Kenapa poin task saya belum masuk setelah submit?', 'Beberapa task (misalnya TPP Exercise, Safety Walk, Mini Project) butuh verifikasi Admin EHS dulu sebelum poinnya dicatat. Status "Pending" berarti sedang menunggu review.', 'Active'],
    ['FAQ05', 'Survey', 'Apa bedanya isi survey Individu dan Kelompok?', 'Individu berarti hanya Anda yang mendapat poin. Kelompok berarti Anda memasukkan NIK anggota lain (tanpa NIK Anda sendiri), dan semua anggota kelompok ikut mendapat poin yang sama.', 'Active'],
    ['FAQ06', 'Poin', 'Bagaimana cara menghitung tier/badge saya?', 'Tier ditentukan dari total poin akumulasi Anda di season aktif: Perintis (0-149), Konsisten (150-399), Penggerak (400-799), Juara EHS (800-1499), Role Model (1500+).', 'Active'],
    ['FAQ07', 'Poin', 'Kenapa peringkat saya di leaderboard tidak berubah?', 'Leaderboard dihitung dari total poin yang sudah disetujui (Approved). Kalau ada klaim yang masih Pending, poinnya belum ikut dihitung sampai disetujui Admin.', 'Active'],
    ['FAQ08', 'BMI', 'Kapan saya bisa isi Measure Your Body lagi?', 'Pengukuran BMI dibatasi 1x per kuartal (3 bulan). Kalau sudah pernah isi di kuartal ini, form akan menolak submisi baru sampai kuartal berikutnya.', 'Active']
  ];
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}


