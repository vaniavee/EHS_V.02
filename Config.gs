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
  '01_Master_Seasons': ['SeasonId', 'SeasonName', 'StartDate', 'EndDate', 'Status', 'Notes'],
  '02_Master_Users': [
    'NIK', 'Nama', 'Divisi', 'No_WA',
    'IsAdmin', 'IsSupervisor', 'DivisiDiawasi', 'ProgramPreferences',
    'ProfileInterests','IsRegistered','Active', 'TanggalDitambahkan', 'DitambahkanOleh'
  ],
  '03_Master_Task': [
    'TaskId', 'Pillar', 'Category', 'Level', 'Title', 'Points','DomainXP', 'CoinReward',
    'Description', 'FrequencyType', 'FrequencyLimit', 'Validation',
    'ObligationLevel', 'Status', 'SeasonId', 'CampaignId'
  ],
  '04_Master_BMI_ScoringRule': ['KategoriLama', 'KategoriBaru', 'KriteriaPoin', 'Poin'
  ],
  '10_DB_PointsLedger': [
    'Timestamp', 'SeasonId', 'Pillar', 'TaskId', 'CampaignId',
    'NIK', 'Nama', 'Divisi', 'ReferenceId', 'Points', 'DomainXP', 'Coin', 'Note'
  ],
  '05_Master_QuizBank': [
  'QuizId', 'TaskId', 'Question', 'OptionA', 'OptionB', 'OptionC', 'OptionD',
  'CorrectOption', 'Explanation', 'Status', 'SeasonId', 'CampaignId', 'QuizScope', 'PeriodType',
  'Difficulty', 'SortOrder', 'ShuffleOptions', 'Source'
  ],
  '18_Master_Departments': ['DepartmentId', 'Name', 'Note'
  ],
  '06_Master_Campaigns': [
    'CampaignId', 'Title', 'MediaType', 'MediaUrl', 'Tagline', 'Description',
    'SubmissionMode', 'MinExposureSeconds', 'SurveyPoints', 'SelfEvalPoints',
    'Status', 'SeasonId','CampaignId', 'PeriodType','ExpectedAction', 'SurveyPointMappingId',
    'QuizPointMappingId','SortOrder','UpdatedAt'
  ],
  '07_Master_SurveyQuestions': ['QuestionId', 'Dimension', 'QuestionText', 'Scale', 'EvaluationLevel', 'Order', 'Status'
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
  ],
  '09_Master_ObsKelompok':    ['QuestionId', 'Dimension', 'QuestionText', 'Scale', 'EvaluationLevel', 'Order', 'Status'
  ],
  '19_Master_ObsIndividu':    ['QuestionId', 'Dimension', 'QuestionText', 'Scale', 'EvaluationLevel', 'Order', 'Status'
  ],
  '20_Master_AwarenessContent': [
    'SeasonId', 'ContentId', 'LegacyCampaignId', 'Title', 'Summary', 'DomainTags', 'TargetPersonas',
    'JourneyModes', 'AudienceSites', 'AudienceDepartments', 'RequirementLabel', 'ContentType',
    'PublishDate', 'ExpiryDate', 'ThumbnailUrl', 'MediaUrl', 'EstimatedMinutes', 'KnowledgeTags',
    'RelatedMissionIds', 'LightXP', 'CTAType', 'PrivacyClass', 'SortOrder', 'Status', 'UpdatedAt'
  ],

  '21_Master_MissionTemplates': [
    'SeasonId', 'MissionId', 'Title', 'Description', 'DomainTags', 'JourneyModes', 'TargetPersonas',
    'AssignmentRuleJson', 'RequirementLabel', 'FrequencyType', 'FrequencyLimit', 'StartDate', 'EndDate',
    'EstimatedMinutes', 'LocationRequirement', 'EvidenceRequirement', 'ValidationMethod', 'BaseXP',
    'DomainXP', 'SustainabilityContribution', 'IntegratedScore', 'RewardCoin', 'DepartmentContribution',
    'DailyCap', 'SeasonalCap', 'PrerequisiteMissionIds', 'RelatedContentIds', 'RelatedChallengeId',
    'SafetyWarning', 'PrivacyClass', 'SortOrder', 'Status', 'UpdatedAt'
  ],

  '22_Master_PointMapping': [
    'SeasonId', 'MappingId', 'MappingName', 'ActivityType', 'Difficulty', 'PeriodType',
    'BasePoints', 'CompletionBonus', 'MaxPerPeriod', 'Status', 'Notes'
  ],

  '23_Master_RewardCatalog': [
    'RewardId', 'Title', 'Description', 'Category', 'CoinCost', 'Stock', 'EligibilityRuleJson',
    'Partner', 'VoucherType', 'ExpiryDays', 'Status', 'UpdatedAt'
  ],

  '24_Master_Clubs': [
    'ClubId', 'Title', 'Description', 'DomainTags', 'Department', 'Site', 'OwnerEmployeeId',
    'MemberLimit', 'Visibility', 'Status', 'UpdatedAt'
  ],

  '25_Master_Challenges': [
    'SeasonId', 'ChallengeId', 'Title', 'Description', 'DomainTags', 'ChallengeType', 'StartDate',
    'EndDate', 'MissionIds', 'TeamSize', 'RewardCoin', 'BonusIntegratedScore', 'PrivacyClass',
    'Status', 'UpdatedAt'
  ],

  '26_Master_JourneyRules': [
    'RuleId', 'RolePattern', 'UserRolePattern', 'Persona', 'ForceJourneyMode', 'AllowedJourneyModes',
    'RequiredDomains', 'RecommendedDomains', 'HealthTrack', 'RequiredMissionTags', 'Priority',
    'Status', 'UpdatedAt'
  ],
  '27_DB_RewardRedemptions': [
  'Timestamp', 'ReferenceId', 'SeasonId', 'NIK', 'Nama', 'Divisi',
  'RewardId', 'RewardTitle', 'CoinCost', 'Status', 'Notes', 'FulfilledBy', 'FulfilledAt'
  ],
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
  } else {
    const existingHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const existingNormalized = existingHeaders.map(function(h) { return clean_(h); });
    const missingHeaders = headers.filter(function(h) {
      return existingNormalized.indexOf(clean_(h)) === -1;
    });
    if (missingHeaders.length) {
      sh.getRange(1, existingHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    }
  }

  const nikColIndex = headers.indexOf('NIK');
  if (nikColIndex !== -1) {
    sh.getRange(2, nikColIndex + 1, 999, 1).setNumberFormat('@');
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
  seedDepartments_(ss);
  seedSampleTasks_(ss);
  seedBmiScoringRules_(ss);
  seedSurveyQuestions_(ss);
  seedFaqData_(ss);
}

function seedDepartments_(ss) {
  const sh = ss.getSheetByName('18_Master_Departments');
  if (sh.getLastRow() > 1) return;

  const rows = getSeedDepartmentRows_();
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
}

function getSeedDepartmentRows_() {
  const departmentNames = `
LOGISTICS
LOGISTICS/ GBB & KEMASAN A (Pharmaceutical I)
LOGISTICS/ GBB & KEMASAN B (Pharma. II & Diagnostic)
LOGISTICS/ GBB & KEMASAN D (Food I)
LOGISTICS/ GBB & KEMASAN E (Food II)
LOGISTICS/ GBB-KEMASAN C & GBJ (Natpro & Extraction)
LOGISTICS/ GBJ (Candy)
LOGISTICS/ GBJ (Pharmaceutical)
LOGISTICS/ GBJ (Sobisco & Diagnostic)
LOGISTICS/ GUDANG MATERIAL PROMOSI
PLANT FOOD
PRE PROCESS
PRE PROCESS/ AT FOOD I
PRE PROCESS/ AT FOOD II
PRODUCTION (FOOD I)
PRODUCTION (FOOD I) / PROSES A
PRODUCTION (FOOD I) / PROSES B
PRODUCTION (FOOD I) / PROSES C
PRODUCTION (FOOD I) / PROSES D
PRODUCTION (FOOD I) / PROSES E
PRODUCTION (FOOD I) / PROSES F
PRODUCTION (FOOD I) / VERPAK A
PRODUCTION (FOOD I) / VERPAK B
PRODUCTION (FOOD I) / VERPAK C
PRODUCTION (FOOD I) / VERPAK D
PRODUCTION (FOOD II)
PRODUCTION (FOOD II) / PROSES A
PRODUCTION (FOOD II) / PROSES B
PRODUCTION (FOOD II) / PROSES C
PRODUCTION (FOOD II) / PROSES D
PRODUCTION (FOOD II) / PROSES E
PRODUCTION (FOOD II) / PROSES F
PRODUCTION (FOOD II) / VERPAK
PLANT FARMASI
PRODUCTION (NATPRO & EXTRACTION)/ PROSES A/ OL
PRODUCTION (NATPRO & EXTRACTION)/ PROSES B/ OD
PRODUCTION (NATPRO & EXTRACTION)/ PROSES C
PRODUCTION (NATPRO & EXTRACTION)/ VERPAK
PRODUCTION (PHARMACEUTICAL I)
PRODUCTION (PHARMACEUTICAL I) / PROSES A
PRODUCTION (PHARMACEUTICAL I) / PROSES B
PRODUCTION (PHARMACEUTICAL I) / PROSES C
PRODUCTION (PHARMACEUTICAL I) / PROSES D
PRODUCTION (PHARMACEUTICAL I) / PROSES E
PRODUCTION (PHARMACEUTICAL I) / PROSES F (AT)
PRODUCTION (PHARMACEUTICAL I) / PROSES G (CETAK)
PRODUCTION (PHARMACEUTICAL I) / PROSES H (GRANUL)
PRODUCTION (PHARMACEUTICAL I) / PROSES I (STRIP)
PRODUCTION (PHARMACEUTICAL I) / PROSES J (EFF/KX 9)
PRODUCTION (PHARMACEUTICAL I) / VERPAK (BORONG)
PRODUCTION (PHARMACEUTICAL I) / VERPAK (EFFERVESCENT)
PRODUCTION (PHARMACEUTICAL I) / VERPAK (TABLET BORONG A)
PRODUCTION (PHARMACEUTICAL I) / VERPAK (TABLET BORONG B)
PRODUCTION (PHARMACEUTICAL I) / VERPAK (TABLET)
PRODUCTION (PHARMACEUTICAL I) / VERPAK A
PRODUCTION (PHARMACEUTICAL II)
PRODUCTION (PHARMACEUTICAL II) / PROSES A (SACHET)
PRODUCTION (PHARMACEUTICAL II) / PROSES B (KOSM/ ALKES)
PRODUCTION (PHARMACEUTICAL II) / PROSES C (AT)
PRODUCTION (PHARMACEUTICAL II) / PROSES D (SIRUP)
PRODUCTION (PHARMACEUTICAL II) / VERPAK A (SIRUP)
PRODUCTION (PHARMACEUTICAL II) / VERPAK B (SIRUP)
PRODUCTION (PHARMACEUTICAL II) / VERPAK C (KOSMETIK)
PRODUCTION (PHARMACEUTICAL II) / VERPAK D (KOSMETIK)
PRODUCTION (PHARMACEUTICAL II) / VERPAK E (KOSMETIK)
SUB DIV OPERATION SUPPORT/ PROJECT BUILDING
SUB DIVISI OPERATION SUPPORT
TECHNICAL SERVICE (FOOD)
TECHNICAL SERVICE (FOOD/ GUDANG SPARE PART)
TECHNICAL SERVICE (FOOD/ PROD & LAB)
TECHNICAL SERVICE (FOOD/ PROD & LAB/ FOOD 1)
TECHNICAL SERVICE (FOOD/ PROD & LAB/ FOOD 2)
TECHNICAL SERVICE (FOOD/ UTILITY)
TECHNICAL SERVICE (FOOD/ UTILITY/ A)
TECHNICAL SERVICE (FOOD/ UTILITY/ ELEKTRO)
TECHNICAL SERVICE (FOOD/ WORKSHOP M/E)
TECHNICAL SERVICE (PHARMA)/ GUDANG SPAREPART
TECHNICAL SERVICE (PHARMA)/ NATPRO & EXTRACT
TECHNICAL SERVICE (PHARMA)/ PROD & LAB
TECHNICAL SERVICE (PHARMA)/ PROD & LAB/ TMP 1
TECHNICAL SERVICE (PHARMA)/ PROD & LAB/ TMP 2
TECHNICAL SERVICE (PHARMA)/ TMP 3/ NATPRO & EXTRACT
TECHNICAL SERVICE (PHARMA)/ UT 2/ NATPRO & EXTRACT
TECHNICAL SERVICE (PHARMA)/ UTILITY
TECHNICAL SERVICE (PHARMA)/ UTILITY 1
TECHNICAL SERVICE (PHARMA)/ WORKSHOP M/E
TECHNICAL SERVICE (PHARMACEUTICAL)
DIRECTOR
DEPUTY DIRECTOR EXTERNAL RELATION (PHARMA & FOOD)
PURCHASING
DIVISI CORP STRATEGIC PLANNING & BUSINESS
CORP. STRATEGIC PLANNING & PERFORMANCE MANAGEMENT
CORP. MGMT INFORMATION SYSTEM
DIVISI FINANCE
ACCOUNTING
ACCOUNTING/ AP & PPH
ACCOUNTING/ GL & PPN
TREASURY
MANAGEMENT AUDIT
FINANCIAL ADVISOR
ETHICAL
ETHICAL / ADMINISTRASI
ETHICAL / MARKETING / PRODUCT
ETHICAL / SALES / JAKARTA I
ETHICAL / SALES / JAKARTA II
ETHICAL / SALES / JAWA BARAT 1
ETHICAL / SALES / JAWA TIMUR I
ETHICAL / SALES / KPP ACEH
ETHICAL / SALES / KPP BDS
ETHICAL / SALES / KPP BDU
ETHICAL / SALES / KPP BENGKULU
ETHICAL / SALES / KPP BUKITTINGGI
ETHICAL / SALES / KPP CIANJUR
ETHICAL / SALES / KPP CIMAHI
ETHICAL / SALES / KPP CIREBON
ETHICAL / SALES / KPP DENPASAR
ETHICAL / SALES / KPP GARUT
ETHICAL / SALES / KPP JAKARTA II
ETHICAL / SALES / KPP MEDAN
ETHICAL / SALES / KPP PATI
ETHICAL / SALES / KPP SIDOARJO
ETHICAL / SALES / KPP SUKABUMI
ETHICAL / SALES / KPP SURABAYA
ETHICAL / SALES / KPP TASIKMALAYA
ETHICAL / SALES / KPP TUBAN
ETHICAL / SALES / MEDAN
ETHICAL / SALES / SUMATERA SELATAN
ETHICAL/ SALES/ KPP BATAM
ETHICAL/ SALES/ KPP GORONTALO
ETHICAL/ SALES/ KPP JAMBI
ETHICAL/ SALES/ KPP JEMBER
ETHICAL/ SALES/ KPP KUDUS
ETHICAL/ SALES/ KPP MADIUN
ETHICAL/ SALES/ KPP MAGELANG
ETHICAL/ SALES/ KPP MALANG
ETHICAL/ SALES/ KPP PADANG
ETHICAL/ SALES/ KPP PALEMBANG
ETHICAL/ SALES/ KPP PEKANBARU
ETHICAL/ SALES/ KPP SEMARANG
ETHICAL/ SALES/ KPP SOLO
ETHICAL/ SALES/ KPP TULUNGAGUNG
ETHICAL/SALES I
ETHICAL/SALES/ KPP BREBES
ETHICAL/SALES/ KPP KEBUMEN
ETHICAL/SALES/ KPP PURBALINGGA
ETHICAL/SALES/ KPP PURWOKERTO
ETHICAL/SALES/ KPP TEGAL
ETHICAL/SALES/ KPP YOGYAKARTA
ETHICAL/SALES/JAWA TENGAH II
ETHICAL/SALES/JAWA TENGAH III
ETHICAL/SALES/KPP BANDUNG
ETHICAL/SALES/KPP BEKASI
ETHICAL/SALES/KPP BOGOR
ETHICAL/SALES/KPP CIKARANG
ETHICAL/SALES/KPP GRESIK
ETHICAL/SALES/KPP JAKARTA I
ETHICAL/SALES/KPP LAMPUNG
ETHICAL/SALES/KPP MOJOKERTO
ETHICAL/SALES/KPP SERANG
ETHICAL/SALES/KPP TANGERANG
DEPUTY DIRECTOR BUSS. DEV. (HEALTHCARE & EXTRACTION)
EXTRACTION / PENGEM. METODE
EXTRACTION / PENGEM. PROSES
EXTRACTION RESEARCH
KONIMEX DIAGNOSTIC CENTER
KONIMEX DIAGNOSTIC CENTER/ RESEARCH
DIVISI MARKETING (HEALTHCARE & EXTRACTION)
GROUP BRAND II
GROUP BRAND III
DIVISI MARKETING FARMASI
GROUP BRAND I
GROUP BRAND IV
MARKETING RESEARCH
MARKETING SUPPORT
MARKETING SUPPORT / ADM
GROUP BRAND V
GROUP BRAND VI
DIVISI HUMAN RESOURCES & ORGANIZATION
HUMAN RESOURCES DEVELOPMENT
GENERAL SERVICE
GENERAL SERVICE/ REP OFFICE JKT
EHS & INDUSTRIAL RELATIONS
PERSONNEL
ORGANIZATION DEVELOPMENT & RECRUITMENT
DIVISI INTERNATIONAL BUSINESS
INTERNATIONAL BUSINESS (FOOD)
INTERNATIONAL BUSINESS (PHARMACEUTICAL)
INTERNATIONAL BUSINESS/ EXPORT SERVICE
DIVISI MODERN TRADE MARKETING
MTM / KEY ACCOUNT
MTM / KEY ACCOUNT & TRADE SUPPORT
MTM / KEY ACCOUNT & TRADE SUPPORT/ ADMINISTRASI
DIVISI QUALITY
DOCUMENT CONTROL
VALIDATION
QUALITY ASSURANCE
QUALITY CONTROL
QUALITY CONTROL / FOOD
QUALITY CONTROL / IMI
QUALITY CONTROL / IPC 1
QUALITY CONTROL / IPC 2
QUALITY CONTROL / MIKRO
DIVISI R & D
PACKAGING DEVELOPMENT
PACKAGING DEVELOPMENT/ FOOD A
PACKAGING DEVELOPMENT/ FOOD B
PACKAGING DEVELOPMENT/ PHARMACEUTICAL A
PACKAGING DEVELOPMENT/ PHARMACEUTICAL B
ANALYTICAL DEVELOPMENT
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ A
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ A-1
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ A-2
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ A-3
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ B
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ C
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ C-1
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ C-2
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ C-3
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ D
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ D-1
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ D-2
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ E
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ E-1
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ E-2
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ E-3
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ F
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ F-1
ANALYTICAL DEVELOPMENT/ COMPLIANCE/ F-2
ANALYTICAL DEVELOPMENT/ NPD/ A
ANALYTICAL DEVELOPMENT/ NPD/ A-1
ANALYTICAL DEVELOPMENT/ NPD/ A-2
ANALYTICAL DEVELOPMENT/ NPD/ B
ANALYTICAL DEVELOPMENT/ NPD/ B-1
ANALYTICAL DEVELOPMENT/ NPD/ B-2
ANALYTICAL DEVELOPMENT/ NPD/ C
ANALYTICAL DEVELOPMENT/ NPD/ C-1
ANALYTICAL DEVELOPMENT/ NPD/ C-2
ANALYTICAL DEVELOPMENT/ NPD/ D
ANALYTICAL DEVELOPMENT/ NPD/ D-1
ANALYTICAL DEVELOPMENT/ NPD/ D-2
ANALYTICAL DEVELOPMENT/ NPD/ E
ANALYTICAL DEVELOPMENT/ NPD/ E-1
ANALYTICAL DEVELOPMENT/ NPD/ E-2
ANALYTICAL DEVELOPMENT/ NPD/ E-3
ANALYTICAL DEVELOPMENT/ NPD/ F
ANALYTICAL DEVELOPMENT/ NPD/ F-1
ANALYTICAL DEVELOPMENT/ NPD/ F-2
REGULATORY AFFAIRS
RPD (PHARMACEUTICAL II)
RPD (PHARMACEUTICAL II)/ PROCESS DEVELOPMENT
RPD (PHARMACEUTICAL II)/ PRODUCT DEV. A
RPD (PHARMACEUTICAL II)/ PRODUCT DEV. B
RPD (CONFECTIONERY)
RPD (FOOD)
RPD (FOOD)/ OPTIMASI
RPD (FOOD)/ PENGEMBANGAN
RPD (PHARMACEUTICAL I)
RPD (PHARMACEUTICAL I)/ PROCESS DEV. A
RPD (PHARMACEUTICAL I)/ PROCESS DEV. B
RPD (PHARMACEUTICAL I)/ PROCESS DEV. C
RPD (PHARMACEUTICAL I)/ PRODUCT DEV. A
RPD (PHARMACEUTICAL I)/ PRODUCT DEV. C
RPD (PHARMACEUTICAL I)/ PRODUCT DEV. D
RPD (PHARMACEUTICAL I)/ PRODUCT DEV. E
RPD (PHARMACEUTICAL I)/ PRODUCT DEV. F
SEC. OF BOD/EXTERNAL RELATION
SEC.OF BOD/EXECUTIVE SECRETARY
SECRETARY OF BOD
`;

  return departmentNames.trim().split(/\r?\n/).map(function(name) {
    name = name.trim();
    return [name, name];
  });
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


