// Period Key adalah identitas unik yang merepresentasikan suatu rentang waktu berdasarkan frekuensi pelaksanaan aktivitas (task), seperti harian, mingguan, bulanan, kuartalan, maupun sekali dalam satu campaign atau season.
// Period Key digunakan sebagai dasar untuk membatasi jumlah klaim aktivitas sesuai aturan frekuensi yang telah ditetapkan

// Menghasilkan Period Key berdasarkan tanggal saat ini.
function getCurrentPeriodKey_(frequencyType) {
  return getPeriodKeyForDate_(frequencyType, new Date());
}

// Menentukan Period Key sesuai jenis frekuensi aktivitas.
function getPeriodKeyForDate_(frequencyType, date) {
  const ft = clean_(frequencyType || 'season_once').toLowerCase();
  const d = date || new Date();
  if (ft === 'daily') return Utilities.formatDate(d, EHS.timezone, 'yyyyMMdd');
  if (ft === 'weekly') return getWeeklyPeriodKey_(d);
  if (ft === 'monthly') return Utilities.formatDate(d, EHS.timezone, 'yyyyMM');
  if (ft === 'quarterly') return getQuarterlyPeriodKey_(d);
  if (ft === 'campaign_once') return 'CAMPAIGN_ONCE';
  return 'SEASON_ONCE';
}

// Menghasilkan Period Key mingguan berdasarkan hari Senin sebagai awal minggu.
function getWeeklyPeriodKey_(date) {
  const y = Number(Utilities.formatDate(date, EHS.timezone, 'yyyy'));
  const m = Number(Utilities.formatDate(date, EHS.timezone, 'MM'));
  const d = Number(Utilities.formatDate(date, EHS.timezone, 'dd'));
  const day = Number(Utilities.formatDate(date, EHS.timezone, 'u')); // 1=Senin .. 7=Minggu
  const monday = new Date(y, m - 1, d, 12, 0, 0);
  monday.setDate(monday.getDate() - (day - 1));
  return 'WEEK_' + Utilities.formatDate(monday, EHS.timezone, 'yyyyMMdd');
}

/**
 * Kuartal kalender: Q1=Jan-Mar, Q2=Apr-Jun, Q3=Jul-Sep, Q4=Okt-Des.
 * Dipakai khusus untuk Measure Your Body (BMI), sesuai siklus 3 bulanan
 * yang disebutkan Vee di requirement awal.
 */

// Menghasilkan Period Key kuartalan berdasarkan kalender tahunan.
function getQuarterlyPeriodKey_(date) {
  const y = Number(Utilities.formatDate(date, EHS.timezone, 'yyyy'));
  const m = Number(Utilities.formatDate(date, EHS.timezone, 'MM'));
  const q = Math.ceil(m / 3);
  return y + '-Q' + q;
}

// Menghasilkan informasi kapan aktivitas dapat diklaim kembali sesuai frekuensinya
function getNextAvailableLabel_(frequencyType) {
  const ft = clean_(frequencyType || 'season_once').toLowerCase();
  const map = {
    daily: 'setelah 23.59 WIB hari ini',
    weekly: 'setelah Minggu 23.59 WIB',
    monthly: 'setelah akhir bulan 23.59 WIB',
    quarterly: 'setelah kuartal berjalan berakhir',
    campaign_once: 'campaign berikutnya',
    season_once: 'season berikutnya'
  };
  return map[ft] || 'periode berikutnya';
}