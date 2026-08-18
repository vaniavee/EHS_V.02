/**
 * Energy Module — Quiz, Survey Media Awareness (Individu/Kelompok), Mini Project.
 */

// --- QUIZ ---
function getQuizForTask(payload) {
  validateRequired_(payload, ['taskId']);
  const seasonId = normalizeSeasonId_(payload.seasonId);
  const task = getTaskById_(payload.taskId, seasonId);
  if (!task) throw new Error('Task tidak ditemukan: ' + payload.taskId);

  const campaignId = clean_(task.CampaignId);
  if (!campaignId) throw new Error('Task ini belum ditautkan ke Campaign manapun. Set CampaignId di Task terlebih dahulu.');

  const questions = readObjects_(getSpreadsheet_().getSheetByName('05_Master_QuizBank'))
    .filter(function(q) {
      return clean_(q.CampaignId) === campaignId && clean_(q.Status) === 'Active';
    })
    .sort(function(a, b) { return Number(a.SortOrder || 0) - Number(b.SortOrder || 0); });

  return questions.map(function(q) {
    let keys = ['A', 'B', 'C', 'D'];
    if (clean_(q.ShuffleOptions).toLowerCase() === 'yes') {
      keys = keys.map(function(k) { return { k: k, sort: Math.random() }; })
        .sort(function(a, b) { return a.sort - b.sort; }).map(function(x) { return x.k; });
    }
    const optionsMap = { A: q.OptionA, B: q.OptionB, C: q.OptionC, D: q.OptionD };
    const options = {};
    keys.forEach(function(originalKey, idx) {
      const displayKey = ['A', 'B', 'C', 'D'][idx];
      if (optionsMap[originalKey]) options[displayKey] = optionsMap[originalKey];
    });
    return { quizId: q.QuizId, question: q.QuestionText, options: options, keyMap: keys };
  });
}

function submitQuizAnswers(payload) {
  validateRequired_(payload, ['nik', 'taskId', 'answers']); // answers: { quizId: displayKey misal 'A' }
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  const seasonId = normalizeSeasonId_(payload.seasonId);
  const task = getTaskById_(payload.taskId, seasonId);
  if (!task) throw new Error('Task tidak ditemukan.');

  const availability = getTaskAvailability_(task, nik, seasonId);
  if (!availability.available) throw new Error(availability.reason);

  const campaignId = clean_(task.CampaignId);
  const questions = readObjects_(getSpreadsheet_().getSheetByName('05_Master_QuizBank'))
    .filter(function(q) { return clean_(q.CampaignId) === campaignId; });

  // keyMap dikirim balik oleh frontend (karena shuffle dilakukan per-sesi di getQuizForTask),
  // supaya displayKey yang dipilih user bisa dipetakan balik ke originalKey yang benar.
  const keyMaps = payload.keyMaps || {}; // { quizId: ['A','C','B','D'] } — urutan originalKey sesuai displayKey A,B,C,D

  let correctCount = 0;
  questions.forEach(function(q) {
    const displayAnswer = clean_(payload.answers[q.QuizId]);
    if (!displayAnswer) return;
    const keyMap = keyMaps[q.QuizId];
    const originalAnswer = keyMap ? keyMap[['A', 'B', 'C', 'D'].indexOf(displayAnswer)] : displayAnswer;
    if (originalAnswer === clean_(q.CorrectOption)) correctCount++;
  });

  const scorePct = questions.length ? correctCount / questions.length : 0;
  const points = Math.round(Number(task.Points || 0) * scorePct);

  const claimNumber = availability.used + 1;
  const referenceId = [seasonId, task.TaskId, nik, availability.periodKey, claimNumber].join(':');

  appendObjectRow_(EHS.sheets.taskClaims, {
    Timestamp: now_(), SeasonId: seasonId, Pillar: 'Energy', NIK: nik, Nama: user.nama, Divisi: user.divisi,
    TaskId: task.TaskId, ReferenceId: referenceId, Status: 'Approved', Points: points,
    Note: correctCount + '/' + questions.length + ' benar', PeriodKey: availability.periodKey,
    FrequencyType: availability.frequencyType, NextAvailableAt: getNextAvailableLabel_(availability.frequencyType),
    BuktiUrl: '', Detail: JSON.stringify(payload.answers)
  });

  awardPoints_(user, 'Energy', task.TaskId, referenceId, points, 'Quiz: ' + correctCount + '/' + questions.length, '', seasonId);
  return { ok: true, correctCount: correctCount, total: questions.length, points: points, message: 'Quiz selesai! ' + correctCount + '/' + questions.length + ' benar, +' + points + ' poin.' };
}

// --- SURVEY MEDIA AWARENESS (Individu & Kelompok) ---
function getActiveCampaigns() {
  return readObjects_(getSpreadsheet_().getSheetByName('06_Master_Campaigns'))
    .filter(function(c) { return clean_(c.Status) === 'Active'; });
}

function submitAwarenessSurvey(payload) {
  validateRequired_(payload, ['nik', 'taskId', 'mode', 'jawaban']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  const seasonId = normalizeSeasonId_(payload.seasonId);

  const task = getTaskById_(payload.taskId, seasonId);
  if (!task) throw new Error('Task tidak ditemukan: ' + payload.taskId);

  const availability = getTaskAvailability_(task, nik, seasonId);
  if (!availability.available) throw new Error(availability.reason);

  // // --- Validasi minimal exposure media ---
  // const campaignId = clean_(payload.campaignId || task.CampaignId);
  // if (campaignId) {
  //   const campaign = readObjects_(getSpreadsheet_().getSheetByName('06_Master_Campaigns'))
  //     .find(function(c) { return clean_(c.CampaignId) === campaignId; });
  //   const minSeconds = campaign ? Number(campaign.MinExposureSeconds || 0) : 0;
  //   if (minSeconds > 0) {
  //     const exposedSeconds = Number(payload.exposedSeconds || 0);
  //     if (exposedSeconds < minSeconds) {
  //       throw new Error('Anda harus membuka/menonton media minimal ' + minSeconds + ' detik sebelum submit. Baru ' + exposedSeconds + ' detik.');
  //     }
  //   }
  // }
  const points = Number(task.Points);
  const title = task.Title;
  const campaignId = clean_(payload.campaignId || task.CampaignId);
  const claimNumber = availability.used + 1;
  const referenceId = [seasonId, task.TaskId, nik, availability.periodKey, claimNumber].join(':');

  if (payload.mode === 'Kelompok') {
    const anggotaNiks = (payload.anggotaKelompok || []).map(function(n) { return normalizeNikLenient_(n); });

    appendObjectRow_('17_DB_GroupSurveySubmissions', {
      Timestamp: now_(), ReferenceId: referenceId, SeasonId: seasonId, CampaignId: campaignId,
      NIK_Pelapor: nik, Nama_Pelapor: user.nama, AnggotaKelompok: anggotaNiks.join(', '),
      Divisi: user.divisi, Jawaban: JSON.stringify(payload.jawaban), Points: points, Status: 'Approved'
    });

    awardPoints_(user, 'Energy', task.TaskId, referenceId, points, 'Survey Kelompok: ' + title, campaignId, seasonId);

    anggotaNiks.forEach(function(anggotaNik, idx) {
      if (!anggotaNik || anggotaNik === nik) return;
      const anggotaUser = getUserProfile_(anggotaNik);
      if (!anggotaUser.found || !anggotaUser.active) return;
      awardPoints_(anggotaUser, 'Energy', task.TaskId, referenceId + ':MEMBER:' + idx, points, 'Survey Kelompok (anggota): ' + title, campaignId, seasonId);
    });
  } else {
    appendObjectRow_(EHS.sheets.taskClaims, {
      Timestamp: now_(), SeasonId: seasonId, Pillar: 'Energy', NIK: nik, Nama: user.nama, Divisi: user.divisi,
      TaskId: task.TaskId, ReferenceId: referenceId, Status: 'Approved', Points: points,
      Note: title, PeriodKey: availability.periodKey, FrequencyType: availability.frequencyType,
      NextAvailableAt: getNextAvailableLabel_(availability.frequencyType), BuktiUrl: '',
      Detail: JSON.stringify({ campaignId: campaignId, jawaban: payload.jawaban })
    });
    awardPoints_(user, 'Energy', task.TaskId, referenceId, points, 'Survey: ' + title, campaignId, seasonId);
  }

  return { ok: true, points: points, message: 'Survey terkirim, +' + points + ' poin untuk semua anggota.' };
}

function getSurveyQuestions() {
  return readObjects_(getSpreadsheet_().getSheetByName('07_Master_SurveyQuestions'))
    .filter(function(q) { return clean_(q.Status) === 'Active'; })
    .sort(function(a, b) { return Number(a.Order) - Number(b.Order); });
}

/**
 * Report singkat Energy (misal E02 - Gemba/Potensi Inefisiensi).
 * Semi-auto: langsung dapat poin tapi tercatat untuk diaudit sampling
 * (konsisten dengan desain validasi awal: Energy Medium = 'verifier').
 */
function submitEnergyActionReport(payload) {
  validateRequired_(payload, ['nik', 'taskId', 'deskripsi']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  const seasonId = normalizeSeasonId_(payload.seasonId);
  const task = getTaskById_(payload.taskId, seasonId);
  if (!task) throw new Error('Task tidak ditemukan.');

  const availability = getTaskAvailability_(task, nik, seasonId);
  if (!availability.available) throw new Error(availability.reason);

  let buktiUrl = '';
  if (payload.buktiBase64) buktiUrl = uploadPhotoToDrive_(payload.buktiBase64, payload.buktiMime, payload.buktiFileName);

  const claimNumber = availability.used + 1;
  const referenceId = [seasonId, task.TaskId, nik, availability.periodKey, claimNumber].join(':');
  const isAuto = clean_(task.Validation).toLowerCase() === 'auto';
  const status = isAuto ? 'Approved' : 'Pending';
  const points = Number(task.Points || 0);

  appendObjectRow_(EHS.sheets.taskClaims, {
    Timestamp: now_(), SeasonId: seasonId, Pillar: 'Energy', NIK: nik, Nama: user.nama, Divisi: user.divisi,
    TaskId: task.TaskId, ReferenceId: referenceId, Status: status, Points: points, Note: payload.deskripsi,
    PeriodKey: availability.periodKey, FrequencyType: availability.frequencyType,
    NextAvailableAt: getNextAvailableLabel_(availability.frequencyType), BuktiUrl: buktiUrl, Detail: ''
  });

  if (isAuto) awardPoints_(user, 'Energy', task.TaskId, referenceId, points, task.Title, task.CampaignId, seasonId);

  return {
    ok: true, points: points,
    message: isAuto ? 'Laporan tersimpan (+' + points + ' poin).' : 'Laporan terkirim, menunggu review Admin EHS.'
  };
}

// --- MINI PROJECT (Hard, Panel Review) ---
function submitMiniProject(payload) {
  validateRequired_(payload, ['nik', 'judulProject', 'areaKerja', 'deskripsiMasalah', 'tindakanPerbaikan']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  const seasonId = normalizeSeasonId_(payload.seasonId);
  const task = getTaskById_('E03', seasonId); // E03 = Mini Project Improvement Area
  if (!task) throw new Error('Task Mini Project tidak ditemukan.');

  const availability = getTaskAvailability_(task, nik, seasonId);
  if (!availability.available) throw new Error(availability.reason);

  let fotoBeforeUrl = '', fotoAfterUrl = '';
  if (payload.fotoBeforeBase64) fotoBeforeUrl = uploadPhotoToDrive_(payload.fotoBeforeBase64, payload.fotoBeforeMime, 'before_' + nik + '.jpg');
  if (payload.fotoAfterBase64) fotoAfterUrl = uploadPhotoToDrive_(payload.fotoAfterBase64, payload.fotoAfterMime, 'after_' + nik + '.jpg');

  const referenceId = [seasonId, 'E03', nik, availability.periodKey].join(':');

  appendObjectRow_('16_DB_MiniProjects', {
    Timestamp: now_(), ReferenceId: referenceId, SeasonId: seasonId, NIK: nik, Nama: user.nama, Divisi: user.divisi,
    JudulProject: payload.judulProject, AreaKerja: payload.areaKerja, DeskripsiMasalah: payload.deskripsiMasalah,
    TindakanPerbaikan: payload.tindakanPerbaikan, FotoBeforeUrl: fotoBeforeUrl, FotoAfterUrl: fotoAfterUrl,
    EstimasiDampak: payload.estimasiDampak || '', AnggotaTim: (payload.anggotaTim || []).join(', '),
    Status: 'Pending', Points: Number(task.Points), AdminFeedback: '', PeriodKey: availability.periodKey
  });
  
  notifyAdmins_('Mini Project Baru', user.nama + ' mengirim Mini Project: ' + payload.judulProject, 'admin', referenceId);

  return { ok: true, message: 'Mini Project terkirim, menunggu review Panel EHS.' };
}

function listMiniProjectsForReview(payload) {
  assertCapability_(payload.nik, 'canApproveAllReports');
  return readObjects_(getSpreadsheet_().getSheetByName('16_DB_MiniProjects'))
    .filter(function(r) { return clean_(r.Status) === 'Pending'; })
    .sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
}

function approveMiniProject(payload) {
  const sh = getSpreadsheet_().getSheetByName('16_DB_MiniProjects');
  const data = sh.getDataRange().getValues();
  const headers = data[0]; const col = {};
  headers.forEach(function(h, i) { col[h] = i; });
  const rowIndex = data.findIndex(function(r) { return clean_(r[col.ReferenceId]) === clean_(payload.referenceId); });
  if (rowIndex === -1) throw new Error('Mini Project tidak ditemukan.');
  const row = data[rowIndex];
  assertCanApprove_(payload.adminNik, clean_(row[col.Divisi]));

  sh.getRange(rowIndex + 1, col.Status + 1).setValue('Approved');
  awardPoints_({ nik: row[col.NIK], nama: row[col.Nama], divisi: row[col.Divisi] },
    'Energy', 'E03', clean_(row[col.ReferenceId]), Number(row[col.Points]), row[col.JudulProject], '', clean_(row[col.SeasonId]));
  return { ok: true, message: 'Mini Project disetujui, +' + row[col.Points] + ' poin.' };
}

function getCampaignById(payload) {
  validateRequired_(payload, ['campaignId']);
  const campaign = readObjects_(getSpreadsheet_().getSheetByName('06_Master_Campaigns'))
    .find(function(c) { return clean_(c.CampaignId) === clean_(payload.campaignId); });
  if (!campaign) throw new Error('Campaign tidak ditemukan.');
  return campaign;
}

function reviseMiniProject(payload) {
  validateRequired_(payload, ['referenceId', 'feedback']);
  const sh = getSpreadsheet_().getSheetByName('16_DB_MiniProjects');
  const data = sh.getDataRange().getValues();
  const headers = data[0]; const col = {};
  headers.forEach(function(h, i) { col[h] = i; });
  const rowIndex = data.findIndex(function(r) { return clean_(r[col.ReferenceId]) === clean_(payload.referenceId); });
  if (rowIndex === -1) throw new Error('Mini Project tidak ditemukan.');
  assertCanApprove_(payload.adminNik, clean_(data[rowIndex][col.Divisi]));

  sh.getRange(rowIndex + 1, col.Status + 1).setValue('Revise');
  const existing = clean_(data[rowIndex][col.AdminFeedback]);
  sh.getRange(rowIndex + 1, col.AdminFeedback + 1).setValue(existing + ' | Revisi: ' + payload.feedback);
  return { ok: true, message: 'Mini Project dikembalikan untuk revisi.' };
}