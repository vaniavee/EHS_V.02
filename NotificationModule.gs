/**
 * Notification Module — notifikasi personal (user) & broadcast (semua admin).
 */

function createNotification_(nik, title, message, relatedPage, relatedRefId) {
  appendObjectRow_('28_DB_Notifications', {
    NotificationId: 'NOTIF:' + nik + ':' + Date.now() + ':' + Math.floor(Math.random() * 1000),
    NIK: nik, Title: title, Message: message,
    RelatedPage: relatedPage || '', RelatedRefId: relatedRefId || '',
    IsRead: 'No', CreatedAt: now_()
  });
}

// Kirim notifikasi ke SEMUA admin sekaligus — dipakai saat ada item baru butuh review.
function notifyAdmins_(title, message, relatedPage, relatedRefId) {
  const admins = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.users))
    .filter(function(u) { return isTrue_(u.IsAdmin) && clean_(u.Active).toLowerCase() !== 'no'; });
  admins.forEach(function(a) {
    createNotification_(normalizeNikLenient_(a.NIK), title, message, relatedPage, relatedRefId);
  });
}

function getUnreadNotificationSummary(payload) {
  try {
    validateRequired_(payload, ['nik']);
    const nik = normalizeNik_(payload.nik);
    const rows = readObjects_(getSpreadsheet_().getSheetByName('28_DB_Notifications'))
      .filter(function(r) { return normalizeNikLenient_(r.NIK) === nik && clean_(r.IsRead) !== 'Yes'; });

    const byPage = {};
    rows.forEach(function(r) {
      const page = clean_(r.RelatedPage) || 'general';
      byPage[page] = (byPage[page] || 0) + 1;
    });
    return { totalUnread: rows.length, byPage: byPage };
  } catch (e) {
    Logger.log('getUnreadNotificationSummary error: ' + (e && e.stack || e));
    return { totalUnread: 0, byPage: {} };
  }
}

function getMyNotifications(payload) {
  try {
    validateRequired_(payload, ['nik']);
    const nik = normalizeNik_(payload.nik);
    const rows = readObjects_(getSpreadsheet_().getSheetByName('28_DB_Notifications'))
      .filter(function(r) { return normalizeNikLenient_(r.NIK) === nik; });
    rows.sort(function(a, b) { return new Date(b.CreatedAt) - new Date(a.CreatedAt); });
    return rows.slice(0, 30);
  } catch (e) {
    Logger.log('getMyNotifications error: ' + (e && e.stack || e));
    return [];
  }
}

function markNotificationRead(payload) {
  validateRequired_(payload, ['nik', 'notificationId']);
  const sh = getSpreadsheet_().getSheetByName('28_DB_Notifications');
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  const rowIdx = data.findIndex(function(r) { return clean_(r[col.NotificationId]) === clean_(payload.notificationId); });
  if (rowIdx === -1) return { ok: false };
  sh.getRange(rowIdx + 1, col.IsRead + 1).setValue('Yes');
  return { ok: true };
}

function markAllNotificationsRead(payload) {
  validateRequired_(payload, ['nik']);
  const nik = normalizeNik_(payload.nik);
  const sh = getSpreadsheet_().getSheetByName('28_DB_Notifications');
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  for (let i = 1; i < data.length; i++) {
    if (normalizeNikLenient_(data[i][col.NIK]) === nik && clean_(data[i][col.IsRead]) !== 'Yes') {
      sh.getRange(i + 1, col.IsRead + 1).setValue('Yes');
    }
  }
  return { ok: true };
}