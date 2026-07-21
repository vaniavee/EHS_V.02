/**
 * BMI Module — Measure Your Body
 * Modul BMI (Measure Your Body) mengelola proses pencatatan pengukuran indeks massa tubuh (Body Mass Index/BMI)
 * Kuartalan. TIDAK ADA syarat akses berbasis poin — semua user aktif
 * bisa submit kapan saja (dibatasi 1x per kuartal).
 * Poin BESAR (300-350) diberikan berdasarkan perbandingan kategori BMI
 * lama vs baru, dibaca dari 04_Master_BMI_ScoringRule.
 */

// Menghitung nilai BMI dan menentukan kategori BMI pengguna
function calculateBmiCategory_(tinggiCm, beratKg) {
  const tinggiM = Number(tinggiCm) / 100;
  const bmi = Number(beratKg) / (tinggiM * tinggiM);
  const rounded = Math.round(bmi * 10) / 10;

  let kategori;
  if (rounded < 18.5) kategori = 'Underweight';
  else if (rounded < 25) kategori = 'Normal';
  else if (rounded < 30) kategori = 'Overweight';
  else if (rounded < 35) kategori = 'Obesitas I';
  else kategori = 'Obesitas II';

  return { bmi: rounded, kategori: kategori };
}

// Mengambil kategori BMI terakhir yang dimiliki pengguna
function getLastBmiCategory_(nik) {
  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.bmiRecords))
    .filter(function(r) { return normalizeNikLenient_(r.NIK) === nik; });
  if (!rows.length) return null;

  rows.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  return clean_(rows[0].Kategori_BMI);
}

/**
 * Tentukan kriteria & poin berdasarkan kategori lama vs baru.
 * Logika disederhanakan langsung di sini (bukan cuma lookup tabel),
 * supaya "bertahan Normal" vs "bertahan level baru non-Normal" tidak
 * perlu didaftar berulang per kategori di master sheet.
 */

// Menentukan kriteria penilaian dan jumlah poin berdasarkan perubahan kategori BMI
function resolveBmiScoring_(kategoriLama, kategoriBaru) {
  if (!kategoriLama) {
    return { kriteriaPoin: 'baseline', points: 0, note: 'Pengukuran BMI pertama (baseline), belum ada poin pembanding.' };
  }

  const urutan = ['Underweight', 'Normal', 'Overweight', 'Obesitas I', 'Obesitas II'];
  const idxLama = urutan.indexOf(kategoriLama);
  const idxBaru = urutan.indexOf(kategoriBaru);

  // Kategori tidak berubah -> bertahan.
  if (kategoriLama === kategoriBaru) {
    if (kategoriBaru === 'Normal') {
      return { kriteriaPoin: 'bertahan_normal', points: 350, note: 'Bertahan di level Normal.' };
    }
    return { kriteriaPoin: 'bertahan_level_baru', points: 300, note: 'Bertahan di level ' + kategoriBaru + ' (tidak naik kembali).' };
  }

  // Kategori membaik (index mendekati Normal, dari kedua arah Underweight/Overweight).
  const membaik =
    (idxLama > urutan.indexOf('Normal') && idxBaru < idxLama && idxBaru >= urutan.indexOf('Normal')) || // turun dari Overweight+ menuju Normal
    (idxLama < urutan.indexOf('Normal') && idxBaru > idxLama); // naik dari Underweight menuju Normal

  if (membaik) {
    return { kriteriaPoin: 'turun_1_level', points: 300, note: kategoriLama + ' -> ' + kategoriBaru + ' (membaik).' };
  }

  // Kategori memburuk -> tidak memenuhi kriteria poin apapun, tetap tercatat 0 poin.
  return { kriteriaPoin: 'tidak_memenuhi_kriteria', points: 0, note: kategoriLama + ' -> ' + kategoriBaru + ' (memburuk, tidak ada poin).' };
}

/**
 * Submit pengukuran BMI (Measure Your Body).
 * payload: { nik, seasonId?, tinggiCm, beratKg, lingkarPinggangCm }
 * Dibatasi 1x per kuartal, TANPA syarat poin minimal untuk mengakses.
 */
// Memproses pengukuran BMI, menyimpan data, dan memberikan poin sesuai hasil evaluasi
function submitBmiRecord(payload) {
  validateRequired_(payload, ['nik', 'tinggiCm', 'beratKg', 'lingkarPinggangCm']);

  const nik = normalizeNik_(payload.nik);
  const user = getUserProfile_(nik);
  if (!user.active) throw new Error('NIK tidak terdaftar atau tidak aktif.');

  const seasonId = normalizeSeasonId_(payload.seasonId);
  const periodKey = getCurrentPeriodKey_('quarterly');

  const existingThisQuarter = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.bmiRecords))
    .some(function(r) {
      return normalizeNikLenient_(r.NIK) === nik && clean_(r.PeriodeKey) === periodKey;
    });
  if (existingThisQuarter) {
    throw new Error('Anda sudah submit Measure Your Body untuk periode ' + periodKey + '.');
  }

  const result = calculateBmiCategory_(payload.tinggiCm, payload.beratKg);
  const kategoriLama = getLastBmiCategory_(nik);
  const scoring = resolveBmiScoring_(kategoriLama, result.kategori);

  appendObjectRow_(EHS.sheets.bmiRecords, {
    Timestamp: now_(),
    NIK: nik,
    Nama: user.nama,
    Divisi: user.divisi,
    Tanggal: now_(),
    Tinggi_cm: Number(payload.tinggiCm),
    Berat_kg: Number(payload.beratKg),
    Lingkar_Pinggang_cm: Number(payload.lingkarPinggangCm),
    BMI: result.bmi,
    Kategori_BMI: result.kategori,
    Kategori_BMI_Sebelumnya: kategoriLama || '',
    Kriteria_Poin: scoring.kriteriaPoin,
    Points: scoring.points,
    PeriodeKey: periodKey
  });

  if (scoring.points > 0) {
    const referenceId = ['BMI', seasonId, nik, periodKey].join(':');
    awardPoints_(user, 'Health', 'BMI', referenceId, scoring.points, scoring.note, '', seasonId);
  }

  return {
    ok: true,
    bmi: result.bmi,
    kategori: result.kategori,
    kategoriSebelumnya: kategoriLama || null,
    points: scoring.points,
    message: scoring.points > 0
      ? 'Pengukuran tersimpan. ' + scoring.note + ' +' + scoring.points + ' poin.'
      : 'Pengukuran tersimpan. ' + scoring.note
  };
}

function getMyBmiRecords(payload) {
  validateRequired_(payload, ['nik']);
  const nik = normalizeNik_(payload.nik);
  const rows = readObjects_(getSpreadsheet_().getSheetByName(EHS.sheets.bmiRecords))
    .filter(function(r) { return normalizeNikLenient_(r.NIK) === nik; });
  rows.sort(function(a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
  return rows;
}