/**
 * MissionEngine.gs — jembatan antara sheet master baru
 * (21_Master_MissionTemplates, 25_Master_Challenges, 20_Master_AwarenessContent)
 * dan halaman MissionHub yang dilihat karyawan.
 *
 * Kenapa dibuat terpisah dari TaskEngine.gs: 03_Master_Task (lama) dan
 * 21_Master_MissionTemplates (baru) adalah dua sumber task yang HIDUP
 * BERDAMPINGAN untuk saat ini — H01-H04 dkk di Master_Task belum dimigrasikan.
 * Klaim untuk MissionId dicatat di sheet klaim (EHS.sheets.taskClaims) yang
 * sama seperti task lama — kolom TaskId di sana cuma teks biasa, jadi aman
 * dipakai bareng tanpa migrasi data.
 */

function tagsInclude_(tagsField, pillar) {
  return clean_(tagsField).split(',').map(function(s) { return s.trim(); }).indexOf(pillar) !== -1;
}

/**
 * Sustainability bukan pilar "berdiri sendiri" — dia adalah gabungan lintas
 * pilar (lihat SDD/PRD). Jadi untuk pillar==='Sustainability', item apapun
 * yang tag-nya menyentuh ≥2 pilar otomatis dianggap masuk Sustainability,
 * TANPA perlu literally di-tag "Sustainability" di sheet. Pillar lain
 * (Energy/Safety/Health) tetap exact-match seperti biasa.
 */
function matchesPillarFilter_(tagsField, pillar) {
  const tags = clean_(tagsField).split(',').map(function(s) { return s.trim(); }).filter(Boolean);
  if (pillar === 'Sustainability') {
    return tags.length >= 2 || tags.indexOf('Sustainability') !== -1;
  }
  return tags.indexOf(pillar) !== -1;
}

function isContentLive_(row) {
  const now = new Date();
  if (row.PublishDate && new Date(row.PublishDate) > now) return false;
  if (row.ExpiryDate && new Date(row.ExpiryDate) < now) return false;
  return true;
}

// ---------- Mission Templates ----------

function getMissionsForPillar_(nik, pillar, seasonId) {
  const rows = readObjects_(getSpreadsheet_().getSheetByName('21_Master_MissionTemplates'))
    .filter(function(m) {
      return clean_(m.Status).toLowerCase() === 'active' &&
             clean_(m.SeasonId) === seasonId &&
             matchesPillarFilter_(m.DomainTags, pillar);
    });

  return rows.map(function(m) {
    const pseudoTask = { TaskId: m.MissionId, FrequencyType: m.FrequencyType, FrequencyLimit: m.FrequencyLimit };
    const availability = getTaskAvailability_(pseudoTask, nik, seasonId);
    return {
      source: 'mission', missionId: m.MissionId, title: m.Title, description: m.Description,
      obligationLevel: mapRequirementLabel_(m.RequirementLabel),
      formType: clean_(m.FormType || 'SubmitMission'),
      points: Number(m.BaseXP || 0), domainXp: Number(m.DomainXP || 0), coin: Number(m.RewardCoin || 0),
      estimatedMinutes: Number(m.EstimatedMinutes || 0),
      validationMethod: clean_(m.ValidationMethod || 'Auto'),
      evidenceRequirement: clean_(m.EvidenceRequirement),
      locationRequirement: clean_(m.LocationRequirement),
      safetyWarning: clean_(m.SafetyWarning),
      relatedContentIds: clean_(m.RelatedContentIds),
      relatedChallengeId: clean_(m.RelatedChallengeId),
      available: availability.available, used: availability.used, limit: availability.limit,
      frequencyType: availability.frequencyType, reason: availability.reason
    };
  });
}

/**
 * Task lama (03_Master_Task) dipetakan ke bentuk yang sama dengan Mission,
 * supaya "Recommended" bisa pilih dari GABUNGAN dua sumber (Task lama +
 * Mission baru) — bukan cuma salah satu.
 */
function getLegacyTasksForPillar_(nik, pillar, seasonId) {
  const tasks = getTasksForUser({ nik: nik, pillar: pillar, seasonId: seasonId });
  return tasks.map(function(t) {
    const rawTask = getTaskById_(t.taskId, seasonId) || {};
    return {
      source: 'task', missionId: t.taskId, title: t.title, description: t.description,
      obligationLevel: t.obligationLevel, formType: clean_(rawTask.FormType || 'Legacy'),
      points: t.points, domainXp: Number(rawTask.DomainXP || 0), coin: Number(rawTask.CoinReward || 0),
      estimatedMinutes: 0, validationMethod: clean_(rawTask.Validation || 'auto'),
      evidenceRequirement: '', locationRequirement: '', safetyWarning: '',
      relatedContentIds: '', relatedChallengeId: '',
      available: t.available, used: t.used, limit: t.limit,
      frequencyType: t.frequencyType, reason: t.reason
    };
  });
}

function getAllMissionsForPillar_(nik, pillar, seasonId) {
  return getLegacyTasksForPillar_(nik, pillar, seasonId).concat(getMissionsForPillar_(nik, pillar, seasonId));
}

function mapRequirementLabel_(label) {
  const l = clean_(label);
  if (l === 'Wajib' || l === 'Required') return 'Required';
  if (l === 'Optional') return 'Optional';
  return 'Recommended';
}

function pickRecommendedMission_(missions) {
  const priority = { Required: 0, Recommended: 1, Optional: 2 };
  const candidates = missions.filter(function(m) { return m.available; });
  candidates.sort(function(a, b) { return priority[a.obligationLevel] - priority[b.obligationLevel]; });
  return candidates[0] || null;
}

/**
 * Klaim mission generik (untuk mission yang belum punya form detail khusus —
 * validasi Auto/langsung disetujui). Mission yang butuh EvidenceRequirement
 * kompleks (foto/lokasi/before-after) tetap perlu form & handler submit
 * khusus terpisah — ini baseline dulu.
 */
function submitMissionClaim(payload) {
  validateRequired_(payload, ['nik', 'missionId']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  if (!user.active) throw new Error('NIK tidak terdaftar atau tidak aktif.');
  const seasonId = normalizeSeasonId_(payload.seasonId);

  const mission = readObjects_(getSpreadsheet_().getSheetByName('21_Master_MissionTemplates'))
    .find(function(m) { return clean_(m.MissionId) === clean_(payload.missionId) && clean_(m.SeasonId) === seasonId && clean_(m.Status).toLowerCase() === 'active'; });
  if (!mission) throw new Error('Mission tidak ditemukan atau tidak aktif: ' + payload.missionId);

  const pseudoTask = { TaskId: mission.MissionId, FrequencyType: mission.FrequencyType, FrequencyLimit: mission.FrequencyLimit };
  const availability = getTaskAvailability_(pseudoTask, nik, seasonId);
  if (!availability.available) throw new Error(availability.reason);

  const pillar = clean_(mission.DomainTags).split(',')[0].trim();
  const claimNumber = availability.used + 1;
  const referenceId = [seasonId, mission.MissionId, nik, availability.periodKey, claimNumber].join(':');

  const points = awardPoints_(
    user, pillar, mission.MissionId, referenceId,
    Number(mission.BaseXP || 0), 'Klaim ' + mission.Title, '', seasonId,
    Number(mission.DomainXP || 0), Number(mission.RewardCoin || 0)
  );

  appendObjectRow_(EHS.sheets.taskClaims, {
    Timestamp: now_(), SeasonId: seasonId, Pillar: pillar, NIK: nik, Nama: user.nama, Divisi: user.divisi,
    TaskId: mission.MissionId, Status: points > 0 ? 'Claimed' : 'Duplicate', Points: points,
    Note: clean_(payload.note || mission.Title), PeriodKey: availability.periodKey,
    FrequencyType: availability.frequencyType, NextAvailableAt: getNextAvailableLabel_(availability.frequencyType)
  });

  updateUserStreak_(nik);
  return { ok: true, message: 'Mission "' + mission.Title + '" berhasil dikirim, +' + points + ' poin.', points: points };
}

// ---------- Challenges ----------

function getChallengesForPillar_(nik, pillar, seasonId) {
  const rows = readObjects_(getSpreadsheet_().getSheetByName('25_Master_Challenges'))
    .filter(function(c) {
      return clean_(c.Status).toLowerCase() === 'active' &&
             clean_(c.SeasonId) === seasonId &&
             matchesPillarFilter_(c.DomainTags, pillar);
    });

  const claimedMissionIds = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims))
    .filter(function(r) { return normalizeNikLenient_(r.NIK) === nik && clean_(r.SeasonId) === seasonId && clean_(r.Status) !== 'Duplicate'; })
    .map(function(r) { return clean_(r.TaskId); });

  return rows.map(function(c) {
    const missionIds = clean_(c.MissionIds).split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    const doneCount = missionIds.filter(function(id) { return claimedMissionIds.indexOf(id) !== -1; }).length;
    return {
      challengeId: c.ChallengeId, title: c.Title, description: c.Description,
      challengeType: clean_(c.ChallengeType || 'Individual'), missionIds: missionIds,
      progressPct: missionIds.length ? Math.round((doneCount / missionIds.length) * 100) : 0,
      doneCount: doneCount, totalCount: missionIds.length,
      isComplete: missionIds.length > 0 && doneCount === missionIds.length,
      rewardCoin: Number(c.RewardCoin || 0), bonusScore: Number(c.BonusIntegratedScore || 0)
    };
  });
}

// Ambil 1 challenge yang belum lengkap (buat card ringkasan di Mission Hub).
// Kalau semua udah lengkap, kasih yang terakhir (biar user tetap lihat sesuatu, dengan status "Selesai").
function pickNextChallenge_(challenges) {
  if (!challenges.length) return null;
  return challenges.find(function(c) { return !c.isComplete; }) || challenges[challenges.length - 1];
}

// ---------- Awareness Content ----------

function getAwarenessForPillar_(nik, pillar, seasonId) {
  const doneContentIds = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.pointsLedger))
    .filter(function(r) { return normalizeNikLenient_(r.NIK) === nik; })
    .map(function(r) { return clean_(r.ReferenceId).split(':')[1]; }); // format ReferenceId: seasonId:contentId:nik

  return readObjects_(getSpreadsheet_().getSheetByName('20_Master_AwarenessContent'))
    .filter(function(a) {
      return clean_(a.Status).toLowerCase() === 'active' &&
             clean_(a.SeasonId) === seasonId &&
             matchesPillarFilter_(a.DomainTags, pillar) &&
             isContentLive_(a);
    })
    .map(function(a) {
      return {
        contentId: a.ContentId, title: a.Title, summary: a.Summary,
        contentType: clean_(a.ContentType), mediaUrl: clean_(a.MediaUrl), thumbnailUrl: clean_(a.ThumbnailUrl),
        estimatedMinutes: Number(a.EstimatedMinutes || 0), lightXp: Number(a.LightXP || 0),
        requirementLabel: clean_(a.RequirementLabel), sortOrder: Number(a.SortOrder || 0),
        isComplete: doneContentIds.indexOf(clean_(a.ContentId)) !== -1
      };
    })
    .sort(function(a, b) { return a.sortOrder - b.sortOrder; });
}

// Ambil 1 konten awareness yang belum diselesaikan (buat card ringkasan di Mission Hub).
function pickNextAwareness_(items) {
  if (!items.length) return null;
  return items.find(function(a) { return !a.isComplete; }) || items[items.length - 1];
}

function markAwarenessComplete(payload) {
  validateRequired_(payload, ['nik', 'contentId']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  if (!user.active) throw new Error('NIK tidak terdaftar atau tidak aktif.');
  const seasonId = normalizeSeasonId_(payload.seasonId);

  const content = readObjects_(getSpreadsheet_().getSheetByName('20_Master_AwarenessContent'))
    .find(function(a) { return clean_(a.ContentId) === clean_(payload.contentId); });
  if (!content) throw new Error('Konten tidak ditemukan: ' + payload.contentId);

  const referenceId = [seasonId, content.ContentId, nik].join(':');
  const already = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.pointsLedger))
    .some(function(r) { return clean_(r.ReferenceId) === referenceId; });
  if (already) return { ok: true, message: 'Konten ini sudah pernah kamu selesaikan.', points: 0 };

  const pillar = clean_(content.DomainTags).split(',')[0].trim();
  const lightXp = Number(content.LightXP || 0);
  if (lightXp > 0) {
    awardPoints_(user, pillar, content.ContentId, referenceId, 0, 'Selesai: ' + content.Title, '', seasonId, lightXp, 0);
  }
  return { ok: true, message: 'Konten "' + content.Title + '" ditandai selesai.', points: lightXp };
}