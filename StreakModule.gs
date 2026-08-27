// Setiap kali user berhasil klaim/submit apa saja yang menghasilkan reward (lewat 'awardPoints_'), update streaknya
// Kalau LastActiveDate = kemarin → StreakCount + 1. 
// Kalau = hari ini → tidak berubah (sudah dihitung). 
// Kalau lebih dari 1 hari lalu → reset ke 1

function updateUserStreak_(nik) {
  const sh = getSpreadsheet_().getSheetByName(EHS.sheets.users);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function(h, i) { col[h] = i; });
  if (col.StreakCount === undefined || col.LastActiveDate === undefined) return;

  const rowIdx = data.findIndex(function(r, i) { return i > 0 && normalizeNikLenient_(r[col.NIK]) === nik; });
  if (rowIdx === -1) return;

  const today = Utilities.formatDate(new Date(), EHS.timezone, 'yyyy-MM-dd');
  const lastActive = clean_(data[rowIdx][col.LastActiveDate]);
  const currentStreak = Number(data[rowIdx][col.StreakCount] || 0);

  if (lastActive === today) return; // sudah dihitung hari ini

  const yesterday = Utilities.formatDate(new Date(Date.now() - 86400000), EHS.timezone, 'yyyy-MM-dd');
  const newStreak = lastActive === yesterday ? currentStreak + 1 : 1;

  sh.getRange(rowIdx + 1, col.StreakCount + 1).setValue(newStreak);
  sh.getRange(rowIdx + 1, col.LastActiveDate + 1).setValue(today);
}

function getUserStreak_(nik) {
  const users = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.users));
  const user = users.find(function(u) { return normalizeNikLenient_(u.NIK) === nik; });
  if (!user) return 0;

  const storedStreak = Number(user.StreakCount || 0);
  const lastActive = clean_(user.LastActiveDate);
  if (!lastActive) return 0;

  const today = Utilities.formatDate(new Date(), EHS.timezone, 'yyyy-MM-dd');
  const yesterday = Utilities.formatDate(new Date(Date.now() - 86400000), EHS.timezone, 'yyyy-MM-dd');

  // Streak masih hidup kalau terakhir aktif hari ini atau kemarin (belum lewat 1 hari penuh
  // tanpa aktivitas). Kalau lebih lama dari itu, streak-nya putus — tampilkan 0 SEKARANG JUGA,
  // jangan nunggu sampai user submit sesuatu lagi baru ke-reset di database.
  if (lastActive === today || lastActive === yesterday) return storedStreak;
  return 0;
}