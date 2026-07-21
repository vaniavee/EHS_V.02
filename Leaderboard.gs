// Leaderboard dihitung LIVE dari ledger, bukan counter tersimpan —
// pola ini dipertahankan dari mentor karena mencegah data leaderboard "drift".
function getLeaderboard_(seasonId, pillar, limit) {
  const sid = normalizeSeasonId_(seasonId);
  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.pointsLedger))
    .filter(function(r) {
      if (clean_(r.SeasonId) !== sid) return false;
      if (pillar && clean_(r.Pillar) !== pillar) return false; // pillar=null -> gabungan 3 pilar
      return true;
    });

  const map = {};
  rows.forEach(function(r) {
    const key = clean_(r.NIK);
    if (!key) return;
    if (!map[key]) map[key] = { nik: key, nama: clean_(r.Nama), divisi: clean_(r.Divisi), points: 0 };
    map[key].points += Number(r.Points || 0);
  });

  return Object.values(map)
    .sort(function(a, b) { return b.points - a.points; })
    .slice(0, limit || 10)
    .map(function(r, i) { return Object.assign({ rank: i + 1 }, r); });
}

/**
 * Total poin akumulasi user dari ledger (lintas pilar Health/Safety/Energy).
 * Dipakai oleh Dashboard, BMI, dan modul lain yang butuh cek total poin.
 */
function getTotalPointsForUser_(nik, seasonId) {
  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.pointsLedger))
    .filter(function(r) {
      return normalizeNikLenient_(r.NIK) === nik && clean_(r.SeasonId) === seasonId;
    });
  return rows.reduce(function(sum, r) { return sum + Number(r.Points || 0); }, 0);
}

function getMonitorData(payload) {
  assertCapability_(payload.nik, 'canManageMasterData');
  const type = payload.type;

  if (type === 'bmi') {
    return readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.bmiRecords))
      .sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  }

  // health / energy -> baca dari task claims, filter by pillar
  const pillar = type === 'health' ? 'Health' : 'Energy';
  return readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims))
    .filter(function(r) { return clean_(r.Pillar) === pillar; })
    .sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
}

function getMonitorDataFiltered(payload) {
  assertCapability_(payload.nik, 'canManageMasterData');
  const type = payload.type;
  let rows;

  if (type === 'bmi') {
    rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.bmiRecords));
  } else if (type === 'energy') {
    const claimRows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims))
      .filter(function(r) { return clean_(r.Pillar) === 'Energy'; });

    // Gabungkan submisi Kelompok (sheet terpisah), map ke bentuk kolom yang sama
    // supaya bisa ditampilkan di tabel yang sama seperti data individu.
    const groupRows = readObjects_(getSpreadsheet_().getSheetByName('17_DB_GroupSurveySubmissions'))
      .map(function(g) {
        return {
          Timestamp: g.Timestamp, NIK: g.NIK_Pelapor, Nama: g.Nama_Pelapor + ' (+kelompok)',
          Divisi: g.Divisi, TaskId: g.CampaignId + ' (Kelompok)', Status: g.Status, Points: g.Points
        };
      });

    rows = claimRows.concat(groupRows);
  } else {
    rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims))
      .filter(function(r) { return clean_(r.Pillar) === 'Health'; });
  }

  if (payload.divisi) rows = rows.filter(function(r) { return clean_(r.Divisi) === payload.divisi; });
  if (payload.status) rows = rows.filter(function(r) { return clean_(r.Status) === payload.status; });
  if (payload.dateFrom) rows = rows.filter(function(r) { return new Date(r.Timestamp) >= new Date(payload.dateFrom); });
  if (payload.dateTo) rows = rows.filter(function(r) { return new Date(r.Timestamp) <= new Date(payload.dateTo + 'T23:59:59'); });

  rows.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  return rows;
}

function getDistinctDivisions() {
  const users = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.users));
  const set = {};
  users.forEach(function(u) { if (clean_(u.Divisi)) set[clean_(u.Divisi)] = true; });
  return Object.keys(set).sort();
}


/**
 * Leaderboard dengan filter pilar + divisi (opsional keduanya).
 * pillar: null = gabungan 3 pilar, atau 'Health'/'Safety'/'Energy'.
 * divisi: '' = semua divisi, atau nama divisi spesifik.
 */
function getLeaderboardFiltered(payload) {
  const seasonId = normalizeSeasonId_(payload.seasonId);
  const pillar = clean_(payload.pillar); // '' berarti gabungan
  const divisi = clean_(payload.divisi);
  const limit = Number(payload.limit || 50);

  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.pointsLedger))
    .filter(function(r) {
      if (clean_(r.SeasonId) !== seasonId) return false;
      if (pillar && clean_(r.Pillar) !== pillar) return false;
      if (divisi && clean_(r.Divisi) !== divisi) return false;
      return true;
    });

  const map = {};
  rows.forEach(function(r) {
    const key = clean_(r.NIK);
    if (!key) return;
    if (!map[key]) map[key] = { nik: key, nama: clean_(r.Nama), divisi: clean_(r.Divisi), points: 0 };
    map[key].points += Number(r.Points || 0);
  });

  return Object.values(map)
    .filter(function(u) { return u.points > 0; }) // sembunyikan yang 0 poin (belum kontribusi apa-apa di filter ini)
    .sort(function(a, b) { return b.points - a.points; })
    .slice(0, limit)
    .map(function(r, i) { return Object.assign({ rank: i + 1 }, r, { badge: resolveBadgeTier_(r.points) }); });
}

/**
 * Daftar divisi unik untuk dropdown filter — dipakai Leaderboard & Admin Monitor.
 * Tanpa capability check karena ini cuma daftar nama divisi, bukan data sensitif.
 */
function getDistinctDivisions() {
  const users = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.users));
  const set = {};
  users.forEach(function(u) { if (clean_(u.Divisi)) set[clean_(u.Divisi)] = true; });
  return Object.keys(set).sort();
}