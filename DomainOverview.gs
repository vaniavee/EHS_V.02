/**
 * DomainOverview.gs — ringkasan 4 kartu di dashboard karyawan.
 *
 * Angka "open missions" dan progress bar dihitung dari sumber yang SAMA dengan
 * daftar misi di Mission Hub (getAllMissionsForPillar_ = task lama + mission
 * template), jadi jumlahnya selalu cocok dengan yang tampil di daftar.
 */

const EHS_DOMAIN_META = {
  Energy: { icon: '⚡', title: 'Energy', tagline: 'Learn efficiently, improve responsibly.', cssClass: 'domain-energy' },
  Safety: { icon: '🛡️', title: 'Safety', tagline: 'Observe clearly. Act positively.', cssClass: 'domain-safety' },
  Health: { icon: '❤️', title: 'Healthy Lifestyle', tagline: 'Move consistently. Recover wisely.', cssClass: 'domain-health' },
  Sustainability: { icon: '🌱', title: 'Sustainability', tagline: 'Safe. Healthy. Efficient.', cssClass: 'domain-sustainability' }
};

function getDomainOverview(payload) {
  validateRequired_(payload, ['nik']);
  const nik = normalizeNik_(payload.nik);
  const seasonId = normalizeSeasonId_(payload.seasonId);

  return ['Energy', 'Safety', 'Health', 'Sustainability'].map(function(pillar) {
    return Object.assign({ pillar: pillar }, EHS_DOMAIN_META[pillar], getDomainStats_(nik, seasonId, pillar));
  });
}

function getDomainStats_(nik, seasonId, pillar) {
  const missions = getAllMissionsForPillar_(nik, pillar, seasonId);
  // Misi dengan limit rusak (NaN / 0, biasanya karena baris sheet bergeser) dikeluarkan
  const valid = missions.filter(function(m) { return isFinite(m.limit) && m.limit > 0; });
  const total = valid.length;
  const completed = valid.filter(function(m) { return m.used >= m.limit; }).length;
  const open = valid.filter(function(m) { return m.available; }).length;
  return {
    totalMissions: total,
    openMissions: open,
    completedMissions: completed,
    progressPct: total ? Math.round((completed / total) * 100) : 0,
    clickable: true
  };
}