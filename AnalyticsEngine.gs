/**
 * Analytics Engine — KPI cards + chart datasets untuk Executive HSE Dashboard.
 * Lihat Analytics_Blueprint.md untuk desain lengkap & justifikasi tiap KPI/chart.
 * Semua fungsi di sini scoped ke SeasonId aktif kecuali disebutkan lain (trend butuh multi-bulan).
 */

const ANALYTICS_SEVERITY_WEIGHT = { 'Low': 1, 'Medium': 3, 'High': 5 };

function pctRound_(n, d) { return d > 0 ? Math.round((n / d) * 1000) / 10 : 0; }

// ---------- KPI CARD ROW (role-aware) ----------

function getExecutiveKpiCards(payload) {
  validateRequired_(payload, ['nik']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  if (!user.found) throw new Error('User tidak ditemukan: ' + nik);
  const seasonId = getActiveSeason_().SeasonId;

  const allUsers = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.users));
  const claims = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims))
    .filter(function(r) { return clean_(r.SeasonId) === seasonId && clean_(r.Status) === 'Approved'; });
  const safety = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.safetyReports))
    .filter(function(r) { return clean_(r.SeasonId) === seasonId && clean_(r.Status) === 'Approved'; });
  const miniProjects = readObjects_(getSpreadsheet_().getSheetByName('16_DB_MiniProjects'))
    .filter(function(r) { return clean_(r.SeasonId) === seasonId && clean_(r.Status) === 'Approved'; });

  if (user.isAdmin) {
    return buildKpiSet_('admin', { claims: claims, safety: safety, miniProjects: miniProjects, users: allUsers }, null);
  }
  if (user.isSupervisor) {
    const viewScope = payload.viewScope === 'self' ? 'self' : 'team';
    if (viewScope === 'self') return buildKpiSet_('karyawan', { claims: claims, safety: safety, miniProjects: miniProjects, users: allUsers }, nik, seasonId);
    return buildKpiSet_('supervisor', { claims: claims, safety: safety, miniProjects: miniProjects, users: allUsers }, user.divisiDiawasi || []);
  }
  return buildKpiSet_('karyawan', { claims: claims, safety: safety, miniProjects: miniProjects, users: allUsers }, nik, seasonId);
}

function buildKpiSet_(role, data, scope, seasonId) {
  let scopedUsers, scopedClaims, scopedSafety, scopedMini;

  if (role === 'admin') {
    scopedUsers = data.users.filter(function(u) { return truthy_(u.Active); });
    scopedClaims = data.claims; scopedSafety = data.safety; scopedMini = data.miniProjects;
  } else if (role === 'supervisor') {
    const divisiSet = scope || [];
    scopedUsers = data.users.filter(function(u) { return truthy_(u.Active) && divisiSet.indexOf(clean_(u.Divisi)) !== -1; });
    const nikSet = {}; scopedUsers.forEach(function(u) { nikSet[normalizeNikLenient_(u.NIK)] = true; });
    scopedClaims = data.claims.filter(function(r) { return nikSet[normalizeNikLenient_(r.NIK)]; });
    scopedSafety = data.safety.filter(function(r) { return divisiSet.indexOf(clean_(r.DivisiDilaporkan)) !== -1; });
    scopedMini = data.miniProjects.filter(function(r) { return nikSet[normalizeNikLenient_(r.NIK)]; });
  } else { // karyawan (atau supervisor viewScope=self)
    const targetNik = scope;
    scopedUsers = data.users.filter(function(u) { return normalizeNikLenient_(u.NIK) === targetNik; });
    scopedClaims = data.claims.filter(function(r) { return normalizeNikLenient_(r.NIK) === targetNik; });
    scopedSafety = data.safety.filter(function(r) { return false; }); // safety report tidak per-nik (dilaporkan oleh supervisor), skip untuk kartu individu
    scopedMini = data.miniProjects.filter(function(r) { return normalizeNikLenient_(r.NIK) === targetNik; });
  }

  const healthCount = scopedClaims.filter(function(r) { return clean_(r.Pillar) === 'Health'; }).length;
  const energyCount = scopedClaims.filter(function(r) { return clean_(r.Pillar) === 'Energy'; }).length;
  const safetyCount = scopedSafety.length;
  const sustainCount = scopedMini.length;

  if (role === 'admin') {
    const activeParticipants = uniqueNikCount_(scopedClaims.concat(scopedMini));
    return {
      role: 'admin',
      cards: [
        { label: 'Total Karyawan Terdaftar', value: scopedUsers.length, icon: '👥' },
        { label: 'Total Energy Submissions', value: energyCount, icon: '⚡' },
        { label: 'Total Safety Observations', value: safetyCount, icon: '🛡️' },
        { label: 'Total Health Submissions', value: healthCount, icon: '❤️' },
        { label: 'Total Sustainability Submissions', value: sustainCount, icon: '🌱' },
        { label: 'Average Participation Rate', value: pctRound_(activeParticipants, scopedUsers.length) + '%', icon: '📊' }
      ]
    };
  }

  if (role === 'supervisor') {
    const teamPoints = sumPoints_(scopedClaims) + sumPoints_(scopedSafety) + sumPoints_(scopedMini);
    const activeParticipants = uniqueNikCount_(scopedClaims.concat(scopedMini));
    return {
      role: 'supervisor',
      cards: [
        { label: 'Total Poin Departemen', value: teamPoints, icon: '🏆' },
        { label: 'Tingkat Partisipasi Tim', value: pctRound_(activeParticipants, scopedUsers.length) + '%', icon: '📊' },
        { label: 'Jumlah Anggota Tim', value: scopedUsers.length, icon: '👥' },
        { label: 'Total Energy Submissions', value: energyCount, icon: '⚡' },
        { label: 'Total Safety Observations', value: safetyCount, icon: '🛡️' },
        { label: 'Total Health Submissions', value: healthCount, icon: '❤️' },
        { label: 'Total Sustainability Submissions', value: sustainCount, icon: '🌱' }
      ]
    };
  }

  // karyawan
  const seasonIdActive = seasonId || getActiveSeason_().SeasonId;
  const totalPoints = getTotalPointsForUser_(scope, seasonIdActive);
  const badge = resolveBadgeTier_(totalPoints);
  const streak = getUserStreak_(scope);
  const lb = getLeaderboard_(seasonIdActive, null, 9999);
  const myRank = lb.find(function(r) { return normalizeNikLenient_(r.NIK) === scope; });

  return {
    role: 'karyawan',
    cards: [
      { label: 'Total Energy Submissions', value: energyCount, icon: '⚡' },
      { label: 'Total Safety Observations', value: safetyCount, icon: '🛡️' },
      { label: 'Total Health Submissions', value: healthCount, icon: '❤️' },
      { label: 'Total Sustainability Submissions', value: sustainCount, icon: '🌱' },
      { label: 'Peringkat', value: myRank ? '#' + myRank.rank : '-', icon: '📈' },
      { label: 'Active Streak', value: (streak || 0) + ' hari', icon: '🔥' },
      { label: 'Lencana Diperoleh', value: badge.name, icon: badge.icon || '🎖️' },
      { label: 'Total Poin', value: totalPoints, icon: '⭐' }
    ]
  };
}

function truthy_(v) { return v === true || clean_(v).toUpperCase() === 'TRUE' || clean_(v) === '1'; }
function uniqueNikCount_(rows) { const s = {}; rows.forEach(function(r) { s[normalizeNikLenient_(r.NIK)] = true; }); return Object.keys(s).length; }
function sumPoints_(rows) { return rows.reduce(function(sum, r) { const n = Number(r.Points); return sum + (isNaN(n) ? 0 : n); }, 0); }

// ---------- CHART 1: Trend & Velocity (monthly, per pilar) ----------

function getTrendAnalysis(payload) {
  validateRequired_(payload, ['nik']);
  assertCapability_(payload.nik, 'canViewAdminDashboard');
  const monthsBack = payload.monthsBack || 6;

  const claims = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims))
    .filter(function(r) { return clean_(r.Status) === 'Approved'; });
  const safety = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.safetyReports))
    .filter(function(r) { return clean_(r.Status) === 'Approved'; });

  const months = lastNMonthKeys_(monthsBack);
  const series = { Health: [], Energy: [], Safety: [], SafetySeverityIndex: [] };

  months.forEach(function(mKey) {
    const claimsInMonth = claims.filter(function(r) { return monthKeyOf_(r.Timestamp) === mKey; });
    const safetyInMonth = safety.filter(function(r) { return monthKeyOf_(r.Timestamp) === mKey; });
    series.Health.push(claimsInMonth.filter(function(r) { return clean_(r.Pillar) === 'Health'; }).length);
    series.Energy.push(claimsInMonth.filter(function(r) { return clean_(r.Pillar) === 'Energy'; }).length);
    series.Safety.push(safetyInMonth.length);
    const weighted = safetyInMonth.reduce(function(sum, r) { return sum + (ANALYTICS_SEVERITY_WEIGHT[clean_(r.Severity)] || 1); }, 0);
    series.SafetySeverityIndex.push(weighted);
  });

  const insight = trendInsight_(series, months);
  return { months: months, series: series, insight: insight };
}

function trendInsight_(series, months) {
  const n = series.Safety.length;
  if (n < 2) return { level: 'info', text: 'Data belum cukup untuk analisis tren (minimal 2 bulan).' };

  const lastSafety = series.Safety[n - 1], prevSafety = series.Safety[n - 2];
  const lastSeverity = series.SafetySeverityIndex[n - 1], prevSeverity = series.SafetySeverityIndex[n - 2];
  const safetyDropPct = prevSafety > 0 ? ((prevSafety - lastSafety) / prevSafety) * 100 : 0;

  if (safetyDropPct > 20 && lastSeverity >= prevSeverity) {
    return { level: 'danger', text: 'Under-reporting Hazard Anomaly: jumlah laporan Safety turun ' + Math.round(safetyDropPct) + '% bulan ini, tapi tingkat keparahan (Severity-Weighted Index) tidak ikut turun. Ini bisa berarti karyawan segan melapor, bukan area kerja jadi lebih aman. Rekomendasi: cek langsung ke lapangan, jangan cuma andalkan angka laporan.' };
  }
  const last3 = series.SafetySeverityIndex.slice(-3);
  if (last3.length === 3 && last3.every(function(v) { return v <= last3[0]; }) && last3[2] < last3[0]) {
    return { level: 'success', text: 'Positive Safety Trend Stabilized: Severity-Weighted Index menurun konsisten 3 bulan terakhir. Program pencegahan tampak efektif — pertahankan ritme kampanye & reward saat ini.' };
  }
  return { level: 'info', text: 'Tren belum menunjukkan pola signifikan (naik/turun tajam) dalam ' + n + ' bulan terakhir. Terus pantau tiap bulan.' };
}

function lastNMonthKeys_(n) {
  const out = [];
  const now = new Date();
  for (let i = n - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    out.push(Utilities.formatDate(d, Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM'));
  }
  return out;
}
function monthKeyOf_(ts) {
  try { return Utilities.formatDate(new Date(ts), Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM'); }
  catch (e) { return ''; }
}

// ---------- CHART 4: Domain Distribution ----------

function getDomainDistribution(payload) {
  validateRequired_(payload, ['nik']);
  assertCapability_(payload.nik, 'canViewAdminDashboard');
  const seasonId = getActiveSeason_().SeasonId;

  const claims = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims))
    .filter(function(r) { return clean_(r.SeasonId) === seasonId && clean_(r.Status) === 'Approved'; });
  const safetyCount = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.safetyReports))
    .filter(function(r) { return clean_(r.SeasonId) === seasonId && clean_(r.Status) === 'Approved'; }).length;
  const sustainCount = readObjects_(getSpreadsheet_().getSheetByName('16_DB_MiniProjects'))
    .filter(function(r) { return clean_(r.SeasonId) === seasonId && clean_(r.Status) === 'Approved'; }).length;

  const dist = {
    Health: claims.filter(function(r) { return clean_(r.Pillar) === 'Health'; }).length,
    Energy: claims.filter(function(r) { return clean_(r.Pillar) === 'Energy'; }).length,
    Safety: safetyCount,
    Sustainability: sustainCount
  };
  const total = dist.Health + dist.Energy + dist.Safety + dist.Sustainability;

  let gapPillar = null;
  Object.keys(dist).forEach(function(k) {
    if (total > 0 && (dist[k] / total) < 0.10) gapPillar = k;
  });
  const insight = total === 0
    ? { level: 'info', text: 'Belum ada submission di season ini.' }
    : gapPillar
      ? { level: 'warning', text: 'Pillar Engagement Gap di ' + gapPillar + ': kontribusinya di bawah 10% dari total submission. Pertimbangkan campaign atau reward booster khusus pilar ini.' }
      : { level: 'success', text: 'Distribusi submission relatif seimbang di keempat pilar.' };

  return { distribution: dist, total: total, insight: insight };
}

// ---------- CHART 3: Department Comparison ----------

function getDepartmentComparison(payload) {
  validateRequired_(payload, ['nik']);
  assertCapability_(payload.nik, 'canViewAdminDashboard');
  const seasonId = getActiveSeason_().SeasonId;

  const users = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.users)).filter(function(u) { return truthy_(u.Active); });
  const ledger = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.pointsLedger))
    .filter(function(r) { return clean_(r.SeasonId) === seasonId; });
  const claims = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims))
    .filter(function(r) { return clean_(r.SeasonId) === seasonId && clean_(r.Status) === 'Approved'; });

  const divisiList = {};
  users.forEach(function(u) { const d = clean_(u.Divisi); if (d) { divisiList[d] = divisiList[d] || { total: 0, active: {} }; divisiList[d].total++; } });
  claims.forEach(function(r) { const d = clean_(r.Divisi); if (divisiList[d]) divisiList[d].active[normalizeNikLenient_(r.NIK)] = true; });

  const rows = Object.keys(divisiList).map(function(d) {
    const points = sumPoints_(ledger.filter(function(r) { return clean_(r.Divisi) === d; }));
    const activeCount = Object.keys(divisiList[d].active).length;
    const rate = pctRound_(activeCount, divisiList[d].total);
    return { divisi: d, totalPoints: points, participationRate: rate, headcount: divisiList[d].total };
  }).sort(function(a, b) { return b.totalPoints - a.totalPoints; }).slice(0, 10);

  let concentrationRisk = null;
  rows.forEach(function(r) { if (r.totalPoints > 0 && r.participationRate < 30) concentrationRisk = r.divisi; });
  const insight = concentrationRisk
    ? { level: 'warning', text: 'Points Concentration Risk di ' + concentrationRisk + ': total poin tinggi tapi participation rate di bawah 30% — kemungkinan poin disumbang segelintir orang saja, partisipasi divisi belum merata.' }
    : { level: 'info', text: 'Belum ada indikasi konsentrasi poin ekstrem pada satu-dua orang di divisi manapun.' };

  return { rows: rows, insight: insight };
}