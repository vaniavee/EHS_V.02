/**
 * Unified Review Queue — gabungan antrian approval dari 3 sumber:
 * Task Claims (Health/Energy), Mini Project, Safety Report (SW/Hazard).
 */

function getUnifiedReviewQueue(payload) {
  validateRequired_(payload, ['nik']);
  const user = getUserProfile_(payload.nik);
  if (!user.active) throw new Error('User tidak aktif.');
  if (!user.isAdmin && !user.isSupervisor) throw new Error('Akses ditolak.');

  // statusFilter: '' (semua), 'Pending' (Menunggu Persetujuan), 'Approved' (Disetujui), 'Revise' (Ditolak)
  const statusFilter = clean_(payload.statusFilter || 'Pending');
  const matchStatus = function(s) { return !statusFilter || clean_(s) === statusFilter; };

  const scopeDivisions = user.isAdmin ? null : user.divisiDiawasi;
  const items = [];

  readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.taskClaims))
    .filter(function(r) { return matchStatus(r.Status); })
    .forEach(function(r) {
      if (scopeDivisions && scopeDivisions.indexOf(clean_(r.Divisi)) === -1) return;
      items.push({
        source: 'taskclaim', referenceId: r.ReferenceId, pillar: r.Pillar, typeLabel: r.TaskId,
        nik: r.NIK, nama: r.Nama, divisi: r.Divisi, title: r.Note || r.TaskId,
        timestamp: r.Timestamp, points: r.Points, detail: r.Detail, buktiUrl: r.BuktiUrl,
        status: r.Status
      });
    });

  readObjects_(getSpreadsheet_().getSheetByName('16_DB_MiniProjects'))
    .filter(function(r) { return matchStatus(r.Status); })
    .forEach(function(r) {
      if (scopeDivisions && scopeDivisions.indexOf(clean_(r.Divisi)) === -1) return;
      items.push({
        source: 'miniproject', referenceId: r.ReferenceId, pillar: 'Sustainability', typeLabel: 'Mini Project',
        nik: r.NIK, nama: r.Nama, divisi: r.Divisi, title: r.JudulProject,
        timestamp: r.Timestamp, points: r.Points,
        detail: JSON.stringify({ areaKerja: r.AreaKerja, masalah: r.DeskripsiMasalah, tindakan: r.TindakanPerbaikan, dampak: r.EstimasiDampak }),
        buktiUrl: r.FotoAfterUrl || r.FotoBeforeUrl,
        status: r.Status
      });
    });

  readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.safetyReports))
    .filter(function(r) { return matchStatus(r.Status); })
    .forEach(function(r) {
      if (scopeDivisions && scopeDivisions.indexOf(clean_(r.DivisiDilaporkan)) === -1) return;
      items.push({
        source: 'safety', referenceId: r.ReferenceId, pillar: 'Safety', typeLabel: r.JenisLaporan,
        nik: r.SupervisorNik, nama: r.SupervisorNama, divisi: r.DivisiDilaporkan, title: r.Deskripsi,
        timestamp: r.Timestamp, points: r.Points || 0, detail: JSON.stringify({ severity: r.Severity }), buktiUrl: r.BuktiUrl,
        status: r.Status
      });
    });

  items.sort(function(a, b) { return new Date(b.timestamp) - new Date(a.timestamp); });
  return items;
}

function approveUnifiedReviewItem(payload) {
  validateRequired_(payload, ['adminNik', 'source', 'referenceId']);
  if (payload.source === 'taskclaim') return approveHealthClaim({ adminNik: payload.adminNik, referenceId: payload.referenceId });
  if (payload.source === 'miniproject') return approveMiniProject({ adminNik: payload.adminNik, referenceId: payload.referenceId });
  if (payload.source === 'safety') return approveSafetyReport({ adminNik: payload.adminNik, referenceId: payload.referenceId });
  throw new Error('Source tidak dikenal: ' + payload.source);
}

function reviseUnifiedReviewItem(payload) {
  validateRequired_(payload, ['adminNik', 'source', 'referenceId', 'feedback']);
  if (payload.source === 'taskclaim') return reviseHealthClaim({ adminNik: payload.adminNik, referenceId: payload.referenceId, feedback: payload.feedback });
  if (payload.source === 'miniproject') return reviseMiniProject({ adminNik: payload.adminNik, referenceId: payload.referenceId, feedback: payload.feedback });
  if (payload.source === 'safety') return reviseSafetyReport({ adminNik: payload.adminNik, referenceId: payload.referenceId, feedback: payload.feedback });
  throw new Error('Source tidak dikenal: ' + payload.source);
}