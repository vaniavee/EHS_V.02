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
  return user ? Number(user.StreakCount || 0) : 0;
}