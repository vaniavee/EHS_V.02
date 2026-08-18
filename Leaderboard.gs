// Leaderboard dihitung LIVE dari ledger, bukan counter tersimpan —
// pola ini dipertahankan dari mentor karena mencegah data leaderboard "drift".
function getLeaderboard_(seasonId, pillar, limit) {
  const sid = normalizeSeasonId_(seasonId);
  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.pointsLedger))
    .filter(function(r) {
      if (clean_(r.SeasonId) !== sid) return false;
      if (clean_(r.TaskId) === 'BMI') return false; // privasi: poin BMI tidak tampil di leaderboard publik
      if (pillar && clean_(r.Pillar) !== pillar) return false;
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
  return rows.reduce(function(sum, r) { const n = Number(r.Points); return sum + (isNaN(n) ? 0 : n); }, 0);
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
  if (payload.kategori) rows = rows.filter(function(r) { return clean_(r.Kategori_BMI) === payload.kategori; });
  if (payload.dateFrom) rows = rows.filter(function(r) { return new Date(r.Timestamp) >= new Date(payload.dateFrom); });
  if (payload.dateTo) rows = rows.filter(function(r) { return new Date(r.Timestamp) <= new Date(payload.dateTo + 'T23:59:59'); });

  rows.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  return rows;
}

/**
 * Daftar divisi unik untuk dropdown filter — dipakai Leaderboard & Admin Monitor.
 * Sumber: sheet master 18_Master_Departments (bukan distinct dari Users), supaya
 * semua divisi yang terdaftar perusahaan muncul di filter — termasuk yang belum
 * punya user terdaftar sama sekali.
 * Tanpa capability check karena ini cuma daftar nama divisi, bukan data sensitif.
 */
function getDistinctDivisions() {
  const rows = readObjects_(getSpreadsheet_().getSheetByName('18_Master_Departments'));
  const set = {};
  rows.forEach(function(d) { if (clean_(d.Name)) set[clean_(d.Name)] = true; });
  return Object.keys(set).sort();
}
/**
 * pillar: null = gabungan 3 pilar, atau 'Health'/'Safety'/'Energy'.
 * divisi: '' = semua divisi, atau nama divisi spesifik.
 */
function getLeaderboardFiltered(payload) {
  const seasonId = normalizeSeasonId_(payload.seasonId);
  const pillar = clean_(payload.pillar);
  const divisi = clean_(payload.divisi);
  const limit = Number(payload.limit || 50);

  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.pointsLedger))
    .filter(function(r) {
      if (clean_(r.SeasonId) !== seasonId) return false;
      if (clean_(r.TaskId) === 'BMI') return false; // privasi: poin BMI tidak tampil di leaderboard publik
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

function getTotalDomainXpForUser_(nik, seasonId, pillar) {
  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.pointsLedger))
    .filter(function(r) {
      return normalizeNikLenient_(r.NIK) === nik && clean_(r.SeasonId) === seasonId &&
             (!pillar || clean_(r.Pillar) === pillar);
    });
  return rows.reduce(function(sum, r) { return sum + Number(r.DomainXP || 0); }, 0);
}

function getTotalCoinsForUser_(nik, seasonId) {
  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.pointsLedger))
    .filter(function(r) { return normalizeNikLenient_(r.NIK) === nik && clean_(r.SeasonId) === seasonId; });
  return rows.reduce(function(sum, r) { const n = Number(r.Coin); return sum + (isNaN(n) ? 0 : n); }, 0);
}

const EHS_DOMAIN_META = {
  Energy: { icon: '⚡', title: 'Energy', tagline: 'Learn efficiently, improve responsibly.', cssClass: 'domain-energy' },
  Safety: { icon: '🛡️', title: 'Safety', tagline: 'Observe clearly. Act positively.', cssClass: 'domain-safety' },
  Health: { icon: '❤️', title: 'Healthy Lifestyle', tagline: 'Move consistently. Recover wisely.', cssClass: 'domain-health' }
};

function getDomainOverview(payload) {
  validateRequired_(payload, ['nik']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  const seasonId = normalizeSeasonId_(payload.seasonId);

  const energyStats = getDomainStats_(nik, seasonId, 'Energy');
  const healthStats = getDomainStats_(nik, seasonId, 'Health');
  const safetyStats = getSafetyDomainStats_(user, seasonId);

  const domains = [
    Object.assign({ pillar: 'Energy' }, EHS_DOMAIN_META.Energy, energyStats),
    Object.assign({ pillar: 'Safety' }, EHS_DOMAIN_META.Safety, safetyStats),
    Object.assign({ pillar: 'Health' }, EHS_DOMAIN_META.Health, healthStats)
  ];

  const totalMissions = domains.reduce(function(s, d) { return s + d.totalMissions; }, 0);
  const totalCompleted = domains.reduce(function(s, d) { return s + d.completedMissions; }, 0);
  const totalOpen = domains.reduce(function(s, d) { return s + d.openMissions; }, 0);
  const sustainabilityPct = totalMissions ? Math.round((totalCompleted / totalMissions) * 100) : 0;

  domains.push({
    pillar: 'Sustainability', icon: '🌱', title: 'Sustainability',
    tagline: 'Safe. Healthy. Efficient.', cssClass: 'domain-sustainability',
    openMissions: totalOpen, totalMissions: totalMissions, progressPct: sustainabilityPct,
    clickable: false, comingSoon: true
  });

  return domains;
}

function getDomainStats_(nik, seasonId, pillar) {
  const tasks = getTasksForUser({ nik: nik, pillar: pillar, seasonId: seasonId });
  const total = tasks.length;
  const openMissions = tasks.filter(function(t) { return t.available; }).length;
  const completedMissions = tasks.filter(function(t) { return t.used >= t.limit && t.limit > 0; }).length;
  const pct = total ? Math.round((completedMissions / total) * 100) : 0;
  return { totalMissions: total, openMissions: openMissions, completedMissions: completedMissions, progressPct: pct, clickable: true };
}

function getSafetyDomainStats_(user, seasonId) {
  const missions = getSafetyMissionsForUser({ nik: user.nik, seasonId: seasonId });
  const total = missions.length;
  const completedMissions = missions.filter(function(m) { return m.used >= m.limit && m.limit > 0; }).length;
  const openMissions = missions.filter(function(m) { return m.available; }).length;
  const pct = total ? Math.round((completedMissions / total) * 100) : 0;
  return { totalMissions: total, openMissions: openMissions, completedMissions: completedMissions, progressPct: pct, clickable: true };
}