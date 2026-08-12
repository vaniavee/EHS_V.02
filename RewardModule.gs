/**
 * Reward Module — Katalog reward, redeem Coin, dan antrian fulfillment Admin.
 */

// Eligibility sederhana: { "minTier": "Konsisten" } — bandingkan urutan tier badge.
const REWARD_TIER_ORDER = ['Perintis', 'Konsisten', 'Penggerak', 'Juara EHS', 'Role Model'];

function evaluateRewardEligibility_(user, totalPoints, eligibilityRuleJson) {
  const raw = clean_(eligibilityRuleJson);
  if (!raw) return { eligible: true };

  let rule;
  try { rule = JSON.parse(raw); } catch (e) { return { eligible: true }; } // JSON rusak -> jangan blokir user

  if (rule.minTier) {
    const currentTier = resolveBadgeTier_(totalPoints).name;
    const currentIdx = REWARD_TIER_ORDER.indexOf(currentTier);
    const requiredIdx = REWARD_TIER_ORDER.indexOf(rule.minTier);
    if (requiredIdx !== -1 && currentIdx < requiredIdx) {
      return { eligible: false, reason: 'Reward ini butuh minimal tier ' + rule.minTier + '.' };
    }
  }
  return { eligible: true };
}

/**
 * Katalog reward untuk user — sudah dilengkapi status eligible & saldo Coin.
 */
function getRewardCatalogForUser(payload) {
  validateRequired_(payload, ['nik']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  if (!user.active) throw new Error('NIK tidak terdaftar atau tidak aktif.');

  const seasonId = normalizeSeasonId_(payload.seasonId);
  const totalCoins = getTotalCoinsForUser_(nik, seasonId);
  const totalPoints = getTotalPointsForUser_(nik, seasonId);

  const rewards = readObjects_(getSpreadsheet_().getSheetByName('23_Master_RewardCatalog'))
    .filter(function(r) { return clean_(r.Status) === 'Active'; });

  return {
    myCoins: totalCoins,
    rewards: rewards.map(function(r) {
      const eligibility = evaluateRewardEligibility_(user, totalPoints, r.EligibilityRuleJson);
      const stock = Number(r.Stock || 0);
      const canAfford = totalCoins >= Number(r.CoinCost || 0);
      return {
        rewardId: r.RewardId, title: r.Title, description: r.Description, category: r.Category,
        coinCost: Number(r.CoinCost || 0), stock: stock, partner: r.Partner, voucherType: r.VoucherType,
        eligible: eligibility.eligible && stock > 0 && canAfford,
        reason: !eligibility.eligible ? eligibility.reason
          : (stock <= 0 ? 'Stok habis.' : (!canAfford ? 'Coin Anda tidak cukup.' : ''))
      };
    })
  };
}

/**
 * Proses redeem reward — aman dari race condition pakai LockService
 * (mirip pola awardPoints_), supaya stok tidak terjual lebih dari yang ada.
 */
function redeemReward(payload) {
  validateRequired_(payload, ['nik', 'rewardId']);
  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  if (!user.active) throw new Error('NIK tidak terdaftar atau tidak aktif.');

  const seasonId = normalizeSeasonId_(payload.seasonId);

  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(10000);
  if (!gotLock) throw new Error('Sistem sedang sibuk, coba lagi beberapa detik.');

  try {
    const sh = getSpreadsheet_().getSheetByName('23_Master_RewardCatalog');
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const col = {};
    headers.forEach(function(h, i) { col[h] = i; });

    const rowIdx = data.findIndex(function(r, i) { return i > 0 && clean_(r[col.RewardId]) === clean_(payload.rewardId); });
    if (rowIdx === -1) throw new Error('Reward tidak ditemukan.');
    const row = data[rowIdx];

    if (clean_(row[col.Status]) !== 'Active') throw new Error('Reward ini sudah tidak aktif.');

    const stock = Number(row[col.Stock] || 0);
    if (stock <= 0) throw new Error('Stok reward ini sudah habis.');

    const coinCost = Number(row[col.CoinCost] || 0);
    const totalCoins = getTotalCoinsForUser_(nik, seasonId);
    if (totalCoins < coinCost) throw new Error('Coin Anda tidak cukup. Dibutuhkan ' + coinCost + ', Anda punya ' + totalCoins + '.');

    const totalPoints = getTotalPointsForUser_(nik, seasonId);
    const eligibility = evaluateRewardEligibility_(user, totalPoints, row[col.EligibilityRuleJson]);
    if (!eligibility.eligible) throw new Error(eligibility.reason);

    // Kurangi stok
    sh.getRange(rowIdx + 1, col.Stock + 1).setValue(stock - 1);

    // Catat pengurangan Coin sebagai entri negatif di ledger (audit trail konsisten dengan pola awardPoints_)
    const referenceId = ['REDEEM', seasonId, payload.rewardId, nik, Date.now()].join(':');
    appendObjectRow_(EHS.sheets.pointsLedger, {
      Timestamp: now_(), SeasonId: seasonId, Pillar: 'Reward', TaskId: 'REDEEM_' + payload.rewardId,
      CampaignId: '', NIK: nik, Nama: user.nama, Divisi: user.divisi, ReferenceId: referenceId,
      Points: 0, DomainXP: 0, Coin: -coinCost, Note: 'Redeem: ' + row[col.Title]
    });

    // Catat sebagai permintaan redemption -> menunggu fulfillment Admin
    appendObjectRow_('27_DB_RewardRedemptions', {
      Timestamp: now_(), ReferenceId: referenceId, SeasonId: seasonId, NIK: nik, Nama: user.nama, Divisi: user.divisi,
      RewardId: payload.rewardId, RewardTitle: row[col.Title], CoinCost: coinCost, Status: 'Pending',
      Notes: '', FulfilledBy: '', FulfilledAt: ''
    });
    createNotification_(nik, 'Redeem Berhasil!', 'Anda berhasil redeem "' + row[col.Title] + '" seharga ' + coinCost + ' coin. Menunggu diproses Admin.', 'reward', referenceId);
    notifyAdmins_('Redemption Baru', user.nama + ' redeem "' + row[col.Title] + '", menunggu diproses.', 'rewardqueue', referenceId);

    return { ok: true, message: 'Reward "' + row[col.Title] + '" berhasil di-redeem, -' + coinCost + ' coin. Menunggu diproses Admin EHS.' };
  } finally {
    lock.releaseLock();
  }
}

function getMyRewardRedemptions(payload) {
  validateRequired_(payload, ['nik']);
  const nik = normalizeNik_(payload.nik);
  const rows = readObjects_(getSpreadsheet_().getSheetByName('27_DB_RewardRedemptions'))
    .filter(function(r) { return normalizeNikLenient_(r.NIK) === nik; });
  rows.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  return rows;
}

/**
 * Admin — antrian fulfillment reward yang di-redeem user.
 */
function listRewardRedemptions(payload) {
  assertCapability_(payload.nik, 'canManageMasterData');
  let rows = readObjects_(getSpreadsheet_().getSheetByName('27_DB_RewardRedemptions'));
  if (payload.statusFilter) rows = rows.filter(function(r) { return clean_(r.Status) === payload.statusFilter; });
  rows.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  return rows;
}

function fulfillRewardRedemption(payload) {
  const admin = assertCapability_(payload.adminNik, 'canManageMasterData');
  validateRequired_(payload, ['referenceId']);

  const sh = getSpreadsheet_().getSheetByName('27_DB_RewardRedemptions');
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  const rowIdx = data.findIndex(function(r) { return clean_(r[col.ReferenceId]) === clean_(payload.referenceId); });
  if (rowIdx === -1) throw new Error('Redemption tidak ditemukan.');
  if (clean_(data[rowIdx][col.Status]) !== 'Pending') throw new Error('Redemption ini sudah diproses sebelumnya.');

  sh.getRange(rowIdx + 1, col.Status + 1).setValue('Fulfilled');
  sh.getRange(rowIdx + 1, col.Notes + 1).setValue(clean_(payload.notes));
  sh.getRange(rowIdx + 1, col.FulfilledBy + 1).setValue(admin.nama);
  sh.getRange(rowIdx + 1, col.FulfilledAt + 1).setValue(now_());
  
  createNotification_(clean_(data[rowIdx][col.NIK]), 'Reward Siap Diambil!', 'Redemption "' + clean_(data[rowIdx][col.RewardTitle]) + '" sudah diproses Admin.', 'reward', payload.referenceId);
  
  return { ok: true, message: 'Redemption ditandai selesai diproses.' };
}

/**
 * Admin batalkan redemption (misal stok fisik ternyata tidak ada) -> Coin & Stok dikembalikan.
 */
function cancelRewardRedemption(payload) {
  const admin = assertCapability_(payload.adminNik, 'canManageMasterData');
  validateRequired_(payload, ['referenceId', 'reason']);

  const sh = getSpreadsheet_().getSheetByName('27_DB_RewardRedemptions');
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  const rowIdx = data.findIndex(function(r) { return clean_(r[col.ReferenceId]) === clean_(payload.referenceId); });
  if (rowIdx === -1) throw new Error('Redemption tidak ditemukan.');
  const row = data[rowIdx];
  if (clean_(row[col.Status]) !== 'Pending') throw new Error('Redemption ini sudah diproses sebelumnya.');

  sh.getRange(rowIdx + 1, col.Status + 1).setValue('Cancelled');
  sh.getRange(rowIdx + 1, col.Notes + 1).setValue(clean_(payload.reason));
  sh.getRange(rowIdx + 1, col.FulfilledBy + 1).setValue(admin.nama);
  sh.getRange(rowIdx + 1, col.FulfilledAt + 1).setValue(now_());

  // Kembalikan Coin (entri positif baru, bukan hapus entri lama, biar audit trail utuh)
  appendObjectRow_(EHS.sheets.pointsLedger, {
    Timestamp: now_(), SeasonId: clean_(row[col.SeasonId]), Pillar: 'Reward', TaskId: 'REFUND_' + row[col.RewardId],
    CampaignId: '', NIK: clean_(row[col.NIK]), Nama: clean_(row[col.Nama]), Divisi: clean_(row[col.Divisi]),
    ReferenceId: payload.referenceId + ':REFUND:' + now_().getTime(),
    Points: 0, DomainXP: 0, Coin: Number(row[col.CoinCost]), Note: 'Refund dibatalkan Admin — ' + payload.reason
  });

  // Kembalikan stok
  const rewardSh = getSpreadsheet_().getSheetByName('23_Master_RewardCatalog');
  const rewardData = rewardSh.getDataRange().getValues();
  const rewardHeaders = rewardData[0];
  const rewardIdCol = rewardHeaders.indexOf('RewardId');
  const stockCol = rewardHeaders.indexOf('Stock');
  const rewardRowIdx = rewardData.findIndex(function(r, i) { return i > 0 && clean_(r[rewardIdCol]) === clean_(row[col.RewardId]); });
  if (rewardRowIdx !== -1) {
    const currentStock = Number(rewardData[rewardRowIdx][stockCol] || 0);
    rewardSh.getRange(rewardRowIdx + 1, stockCol + 1).setValue(currentStock + 1);
  }

  return { ok: true, message: 'Redemption dibatalkan, Coin dan stok dikembalikan.' };
}