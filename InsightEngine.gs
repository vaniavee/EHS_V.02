/**
 * Insight Engine — decision matrix personal berdasarkan partisipasi Health/Energy + kategori BMI terakhir.
 */

const INSIGHT_RISKY_BMI = ['Obesitas I', 'Obesitas II',' Obesitas III', 'Underweight'];
const INSIGHT_ACTIVE_THRESHOLD = 3; // total klaim approved (Health+Energy) minimal, dianggap "aktif"

function generateUserInsight(nik) {
  const normalizedNik = normalizeNik_(nik);
  const user = getUserProfile_(normalizedNik);
  if (!user.found) throw new Error('User tidak ditemukan: ' + normalizedNik);

  const seasonId = getActiveSeason_().SeasonId;

  const claims = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims))
    .filter(function(r) { return normalizeNikLenient_(r.NIK) === normalizedNik && clean_(r.SeasonId) === seasonId; });

  const healthCount = claims.filter(function(r) { return clean_(r.Pillar) === 'Health' && clean_(r.Status) === 'Approved'; }).length;
  const energyCount = claims.filter(function(r) { return clean_(r.Pillar) === 'Energy' && clean_(r.Status) === 'Approved'; }).length;

  const bmiRecords = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.bmiRecords))
    .filter(function(r) { return normalizeNikLenient_(r.NIK) === normalizedNik; })
    .sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  const latestBmiCategory = bmiRecords.length ? clean_(bmiRecords[0].Kategori_BMI) : null;

  return resolveInsight_(user, healthCount, energyCount, latestBmiCategory);
}

function resolveInsight_(user, healthCount, energyCount, bmiCategory) {
  const totalActivity = healthCount + energyCount;
  const isActive = totalActivity >= INSIGHT_ACTIVE_THRESHOLD;
  const isRiskyBmi = bmiCategory && INSIGHT_RISKY_BMI.indexOf(bmiCategory) !== -1;

  if (isActive && isRiskyBmi) {
    return {
      statusTipe: 'Aktif — Fisik Berisiko', level: 'warning',
      insightDampak: 'Motivasi bergerak tinggi, tapi kondisi fisik (' + bmiCategory + ') masih berisiko. Kombinasi ini berpotensi memicu cedera sendi atau kelelahan berlebih kalau intensitas tidak terukur.',
      rekomendasiAdmin: 'Perlu pendampingan edukasi nutrisi (diet kalori) dan pemantauan intensitas latihan agar tidak over-training.',
      rekomendasiUser: 'Semangatmu luar biasa! Namun untuk menjaga sendimu, yuk imbangi dengan konsultasi menu nutrisi di klinik dan sesuaikan target langkah harianmu.'
    };
  }
  if (totalActivity === 0) {
    return {
      statusTipe: 'Abai / Pasif', level: 'danger',
      insightDampak: 'Nol partisipasi di Health maupun Energy sepanjang season berjalan. Tipe "Disengaged" seperti ini umumnya berkorelasi dengan situational awareness rendah di area kerja.',
      rekomendasiAdmin: 'Peringatan dini — pertimbangkan inspeksi mendadak (sidak) perilaku keselamatan di area kerjanya, masukkan ke daftar intervensi personal.',
      rekomendasiUser: 'Keamanan dan kesehatanmu adalah prioritas keluarga di rumah. Yuk, mulai dengan ambil 1 tugas kesehatan ringan minggu ini!'
    };
  }
  if (isActive && !isRiskyBmi) {
    return {
      statusTipe: 'Aktif & Ideal', level: 'success',
      insightDampak: 'Partisipasi konsisten dengan kondisi fisik dalam rentang aman. Pola ini bagus untuk dipertahankan.',
      rekomendasiAdmin: 'Tidak perlu intervensi khusus — pertimbangkan sebagai kandidat role model/champion EHS di divisinya.',
      rekomendasiUser: 'Kerja bagus! Konsistensimu di Health & Energy sudah on track. Terus pertahankan ritme ini.'
    };
  }
  return {
    statusTipe: 'Partisipasi Rendah', level: 'warning',
    insightDampak: 'Partisipasi masih di bawah target konsistensi. Belum ada indikasi risiko fisik langsung, tapi keterlibatan rendah tetap perlu didorong.',
    rekomendasiAdmin: 'Berikan reminder/ajakan ringan untuk meningkatkan partisipasi, tanpa perlu eskalasi khusus.',
    rekomendasiUser: 'Yuk mulai lagi sedikit demi sedikit — coba 1 task Energy atau Health minggu ini untuk menjaga ritme.'
  };
}

function getMyInsight(payload) {
  validateRequired_(payload, ['nik']);
  return generateUserInsight(payload.nik);
}

function getUserInsightForAdmin(payload) {
  assertCapability_(payload.adminNik, 'canManageMasterData');
  validateRequired_(payload, ['targetNik']);
  return generateUserInsight(payload.targetNik);
}