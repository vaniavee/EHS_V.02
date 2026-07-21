/**
 * Help Module — FAQ statis + Help Request (tanya ke Admin).
 */

// const EHS_FAQ = [
//   {
//     category: 'Login',
//     question: 'Bagaimana cara login ke EHS App?',
//     answer: 'Masukkan NIK Anda di halaman login, lalu tekan Masuk. Tidak perlu password. Kalau NIK belum terdaftar, hubungi Admin EHS untuk didaftarkan terlebih dahulu.'
//   },
//   {
//     category: 'Login',
//     question: 'NIK saya tidak dikenali, apa yang harus dilakukan?',
//     answer: 'Pastikan format NIK sudah benar (contoh: 00180 atau MJ00313). Kalau masih gagal, kemungkinan NIK Anda belum didaftarkan Admin EHS — hubungi Admin lewat WhatsApp di halaman login.'
//   },
//   {
//     category: 'Task',
//     question: 'Kenapa tombol klaim task tertulis "Kuota Habis"?',
//     answer: 'Setiap task punya batas klaim per periode (harian/mingguan/bulanan). Kalau kuota sudah habis, tombol akan menampilkan kapan kuota tersedia lagi.'
//   },
//   {
//     category: 'Task',
//     question: 'Kenapa poin task saya belum masuk setelah submit?',
//     answer: 'Beberapa task (misalnya TPP Exercise, Safety Walk, Mini Project) butuh verifikasi Admin EHS dulu sebelum poinnya dicatat. Status "Pending" berarti sedang menunggu review.'
//   },
//   {
//     category: 'Survey',
//     question: 'Apa bedanya isi survey Individu dan Kelompok?',
//     answer: 'Individu berarti hanya Anda yang mendapat poin. Kelompok berarti Anda memasukkan NIK anggota lain (tanpa NIK Anda sendiri), dan semua anggota kelompok ikut mendapat poin yang sama.'
//   },
//   {
//     category: 'Poin',
//     question: 'Bagaimana cara menghitung tier/badge saya?',
//     answer: 'Tier ditentukan dari total poin akumulasi Anda di season aktif: Perintis (0-149), Konsisten (150-399), Penggerak (400-799), Juara EHS (800-1499), Role Model (1500+).'
//   },
//   {
//     category: 'Poin',
//     question: 'Kenapa peringkat saya di leaderboard tidak berubah?',
//     answer: 'Leaderboard dihitung dari total poin yang sudah disetujui (Approved). Kalau ada klaim yang masih Pending, poinnya belum ikut dihitung sampai disetujui Admin.'
//   },
//   {
//     category: 'BMI',
//     question: 'Kapan saya bisa isi Measure Your Body lagi?',
//     answer: 'Pengukuran BMI dibatasi 1x per kuartal (3 bulan). Kalau sudah pernah isi di kuartal ini, form akan menolak submisi baru sampai kuartal berikutnya.'
//   }
// ];

// /**
//  * Ambil daftar FAQ statis — dikirim ke frontend sebagai referensi awal
//  * sebelum user mengetik pertanyaan sendiri.
//  */
// function getFaqList() {
//   return EHS_FAQ;
// }
function getFaqList() {
  return readObjects_(getSpreadsheet_().getSheetByName('08_Master_FAQ'))
    .filter(function(f) { return clean_(f.Status) === 'Active'; });
}

/**
 * User mengirim pertanyaan baru ke Admin (kalau tidak terjawab oleh FAQ).
 */
function submitHelpRequest(payload) {
  validateRequired_(payload, ['nik', 'question']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  if (!user.active) throw new Error('NIK tidak terdaftar atau tidak aktif.');

  const requestId = 'HELP:' + nik + ':' + now_().getTime();

  appendObjectRow_(EHS.sheets.helpRequests, {
    Timestamp: now_(),
    RequestId: requestId,
    NIK: nik,
    Nama: user.nama,
    Question: payload.question,
    Status: 'Pending',
    Answer: '',
    AnsweredAt: '',
    AnsweredBy: ''
  });

  return { ok: true, message: 'Pertanyaan terkirim. Admin EHS akan menjawab secepatnya.' };
}

/**
 * Riwayat pertanyaan milik user sendiri (ditampilkan di halaman Bantuan).
 */
function getMyHelpRequests(payload) {
  validateRequired_(payload, ['nik']);
  const nik = normalizeNik_(payload.nik);
  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.helpRequests))
    .filter(function(r) { return normalizeNikLenient_(r.NIK) === nik; });
  rows.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  return rows;
}

/**
 * Admin: lihat semua pertanyaan (bisa filter status Pending saja).
 */
function listAllHelpRequests(payload) {
  assertCapability_(payload.nik, 'canManageMasterData');
  let rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.helpRequests));
  if (payload.statusFilter) rows = rows.filter(function(r) { return clean_(r.Status) === payload.statusFilter; });
  rows.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  return rows;
}

/**
 * Admin menjawab satu pertanyaan.
 */
function answerHelpRequest(payload) {
  const admin = assertCapability_(payload.adminNik, 'canManageMasterData');
  validateRequired_(payload, ['requestId', 'answer']);

  const sh = getSpreadsheet_().getSheetByName(EHS.sheets.helpRequests);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  const rowIdx = data.findIndex(function(r) { return clean_(r[col.RequestId]) === clean_(payload.requestId); });
  if (rowIdx === -1) throw new Error('Pertanyaan tidak ditemukan.');

  sh.getRange(rowIdx + 1, col.Status + 1).setValue('Answered');
  sh.getRange(rowIdx + 1, col.Answer + 1).setValue(payload.answer);
  sh.getRange(rowIdx + 1, col.AnsweredAt + 1).setValue(now_());
  sh.getRange(rowIdx + 1, col.AnsweredBy + 1).setValue(admin.nama);

  return { ok: true, message: 'Jawaban tersimpan.' };
}