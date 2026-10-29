/**
 * SISTEM ABSENSI (GOOGLE APPS SCRIPT)
 * Lokasi: Jl. Progo No. 12, Bandung
 */

const LOKASI = { lat: -6.9046, lng: 107.6184 };
const RADIUS_M = 100;
const TZ = 'Asia/Jakarta';
const SPREADSHEET_ID = '1BNZnczmcVutERAM0z_G9NNvU1eGt9yzkEpYF4OjbaYY';

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Presensi Jonas Photo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Inisialisasi sheet Freelancer, Absensi, dan Rekap
 */
function setup() {
  const ss = getSS();
  ss.setSpreadsheetTimeZone(TZ);
  ss.setSpreadsheetLocale('en_US');

  // 1. Sheet Database
  let f = ss.getSheetByName('Database') || ss.getSheetByName('Freelancer');
  if (!f) {
    f = ss.insertSheet('Database');
  }
  f.getRange('A1:I1').setValues([['Nama', 'ID', 'PIN', 'Status', 'Cabang', 'Departemen', 'Tarif', 'No Rekening', 'Bank']]).setFontWeight('bold');
  f.getRange('B:B').setNumberFormat('@'); // Format teks ID
  f.getRange('C:C').setNumberFormat('@'); // Format teks PIN
  f.getRange('G:G').setNumberFormat('#,##0'); // Format Tarif
  f.getRange('H:H').setNumberFormat('@'); // Format teks No Rekening
  if (f.getLastRow() < 2) {
    f.getRange('A2:I2').setValues([['Mikhael Yobel Saputra', '@mikhaelyobels', '1111', 'AI Prompter', 'Pusat', 'IT', 75000, '3790944231', 'BCA']]);
  }

  // 2. Sheet Absensi
  let a = ss.getSheetByName('Absensi');
  if (!a) {
    a = ss.insertSheet('Absensi');
  }
  a.getRange('A1:E1').setValues([['Tanggal', 'Nama', 'Masuk', 'Pulang', 'Jam Kerja']]).setFontWeight('bold');
  a.getRange('C:D').setNumberFormat('HH.mm');
  a.getRange('E:E').setNumberFormat('0.00');

  // Konversi tanggal string jika ada (misal diinput manual sebagai DD/MM/YYYY) ke Date objek
  if (a.getLastRow() >= 2) {
    const rangeTgl = a.getRange(2, 1, a.getLastRow() - 1, 1);
    const vals = rangeTgl.getValues();
    let modified = false;
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i][0];
      if (typeof v === 'string' && v.includes('/')) {
        const p = v.split('/');
        if (p.length === 3) {
          vals[i][0] = new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
          modified = true;
        }
      } else if (typeof v === 'string' && v.includes('-')) {
        const p = v.split('-');
        if (p.length === 3) {
          vals[i][0] = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
          modified = true;
        }
      }
    }
    if (modified) {
      rangeTgl.setValues(vals);
    }
    a.getRange('A2:A').setNumberFormat('dd/mm/yyyy');

    // Pastikan rumus Jam Kerja di kolom E terpasang untuk semua baris
    const numRows = a.getLastRow() - 1;
    const formulasE = [];
    for (let i = 2; i <= a.getLastRow(); i++) {
      formulasE.push([`=IF(OR(C${i}="", D${i}=""), "", ROUND(IF(D${i} < C${i}, D${i} + 1 - C${i}, D${i} - C${i}) * 24, 2))`]);
    }
    a.getRange(2, 5, numRows, 1).setFormulas(formulasE);
    a.getRange('E2:E').setNumberFormat('0.00').setHorizontalAlignment('center');
  }

  // 3. Sheet Rekap (Filter Nama, Bulan, & Perhitungan Siklus Tanggal 21 - 20)
  let r = ss.getSheetByName('Rekap');
  if (!r) {
    r = ss.insertSheet('Rekap');
  }

  // Baris 1: Filter Controls
  r.getRange('A1').setValue('Filter Nama:').setFontWeight('bold');
  if (!r.getRange('B1').getValue()) {
    r.getRange('B1').setValue('Semua');
  }

  r.getRange('C1').setValue('Bulan:').setFontWeight('bold');
  const bulanList = [
    '1 (Januari)', '2 (Februari)', '3 (Maret)', '4 (April)',
    '5 (Mei)', '6 (Juni)', '7 (Juli)', '8 (Agustus)',
    '9 (September)', '10 (Oktober)', '11 (November)', '12 (Desember)'
  ];
  const ruleBulan = SpreadsheetApp.newDataValidation()
    .requireValueInList(bulanList, true)
    .build();

  const now = new Date();
  let defaultMonthIdx = now.getMonth(); // 0..11
  let defaultYear = now.getFullYear();
  if (now.getDate() >= 21) {
    if (defaultMonthIdx === 11) {
      defaultMonthIdx = 0;
      defaultYear += 1;
    } else {
      defaultMonthIdx += 1;
    }
  }

  if (!r.getRange('D1').getValue()) {
    r.getRange('D1').setDataValidation(ruleBulan).setValue(bulanList[defaultMonthIdx]);
  } else {
    r.getRange('D1').setDataValidation(ruleBulan);
  }

  r.getRange('E1').setValue('Tahun:').setFontWeight('bold');
  if (!r.getRange('F1').getValue()) {
    r.getRange('F1').setValue(defaultYear);
  }

  // Baris 2: Keterangan Periode Cut-Off (21 Bulan Lalu s/d 20 Bulan Ini)
  r.getRange('A2').setValue('Periode Cut-Off:').setFontWeight('bold');
  r.getRange('B2').setFormula(
    `=TEXT(DATE(F1, VALUE(LEFT(D1, FIND(" ", D1 & " ") - 1)) - 1, 21), "dd/mm/yyyy") & " s/d " & TEXT(DATE(F1, VALUE(LEFT(D1, FIND(" ", D1 & " ") - 1)), 20), "dd/mm/yyyy")`
  ).setFontWeight('bold').setFontColor('#15803d');

  // Sel bantu Tanggal Mulai (G2) dan Tanggal Selesai (H2) format yyyy-mm-dd
  r.getRange('G2').setFormula(`=TEXT(DATE(F1, VALUE(LEFT(D1, FIND(" ", D1 & " ") - 1)) - 1, 21), "yyyy-mm-dd")`);
  r.getRange('H2').setFormula(`=TEXT(DATE(F1, VALUE(LEFT(D1, FIND(" ", D1 & " ") - 1)), 20), "yyyy-mm-dd")`);
  r.getRange('G2:H2').setFontColor('#ffffff'); // Sembunyikan warna teks sel bantu

  // Kolom K: Dynamic list nama yang hadir pada periode cut-off terpilih
  r.getRange('K1').setValue('Semua');
  r.getRange('K2').setFormula(
    `=IFERROR(UNIQUE(QUERY(Absensi!A2:B, "SELECT B WHERE B IS NOT NULL AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' ORDER BY B LABEL B ''", 0)), "")`
  );
  r.getRange('K:K').setFontColor('#ffffff');
  try {
    r.hideColumns(11); // Sembunyikan kolom K agar sheet rapi
  } catch (err) {}

  // Update Data Validation Filter Nama (B1) mengacu ke Kolom K (hanya nama yang aktif di bulan tsb)
  const ruleNama = SpreadsheetApp.newDataValidation()
    .requireValueInRange(r.getRange('K1:K'))
    .setAllowInvalid(true)
    .build();
  r.getRange('B1').setDataValidation(ruleNama);

  // Tabel 1 (Kolom A-C): Ringkasan Total
  r.getRange('A4:C4').setValues([['Nama', 'Total Hari Kerja', 'Total Jam Kerja']])
    .setFontWeight('bold')
    .setBackground('#e2e8f0');
  r.getRange('A5').setFormula(
    `=IFERROR(IF(OR(B1="Semua", B1=""), QUERY(Absensi!A2:E, "SELECT B, COUNT(A), SUM(E) WHERE B IS NOT NULL AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' GROUP BY B LABEL COUNT(A) '', SUM(E) ''", 0), QUERY(Absensi!A2:E, "SELECT B, COUNT(A), SUM(E) WHERE B = '" & B1 & "' AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' GROUP BY B LABEL COUNT(A) '', SUM(E) ''", 0)), "")`
  );
  r.getRange('C5:C').setNumberFormat('0.00');

  // Tabel 2 (Kolom E-I): Rincian Absensi Harian (Terbaru di paling atas)
  r.getRange('E4:I4').setValues([['Tanggal', 'Nama', 'Masuk', 'Pulang', 'Jam Kerja']])
    .setFontWeight('bold')
    .setBackground('#e2e8f0');
  r.getRange('E5').setFormula(
    `=IFERROR(IF(OR(B1="Semua", B1=""), QUERY(Absensi!A2:E, "SELECT A, B, C, D, E WHERE B IS NOT NULL AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' ORDER BY A DESC, B ASC", 0), QUERY(Absensi!A2:E, "SELECT A, B, C, D, E WHERE B = '" & B1 & "' AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' ORDER BY A DESC", 0)), "")`
  );
  r.getRange('E5:E').setNumberFormat('dd/mm/yyyy');
  r.getRange('G5:H').setNumberFormat('HH.mm');
  r.getRange('I5:I').setNumberFormat('0.00');

  // 4. Sheet Fee (Konfigurasi Dinamis 1 Orang / Semua Orang)
  const fee = ss.getSheetByName('Fee');
  if (fee) {
    renderFeeSheet(fee);
  }
}

/**
 * Render Sheet Fee secara dinamis:
 * Mendukung filter 1 orang maupun "Semua" orang.
 * Posisi baris "Jumlah" dan blok TTD (beserta gambar tanda tangan)
 * otomatis bergeser turun di bawah baris data orang terakhir.
 */
function renderFeeSheet(fee) {
  if (!fee) {
    const ss = getSS();
    fee = ss.getSheetByName('Fee');
  }
  if (!fee) return;

  // 1. Pastikan kolom cukup (minimal 44 kolom: A..AP)
  if (fee.getMaxColumns() < 44) {
    fee.insertColumnsAfter(fee.getMaxColumns(), 44 - fee.getMaxColumns());
  }

  // 2. Bulan & Tahun di F1:G1 & F2:G2
  const defaultYear = 2026;
  const defaultMonthIdx = 10; // November (11)
  const bulanList = [
    '1 (Januari)', '2 (Februari)', '3 (Maret)', '4 (April)',
    '5 (Mei)', '6 (Juni)', '7 (Juli)', '8 (Agustus)',
    '9 (September)', '10 (Oktober)', '11 (November)', '12 (Desember)'
  ];
  const ruleBulan = SpreadsheetApp.newDataValidation()
    .requireValueInList(bulanList, true)
    .build();

  fee.getRange('F1').setValue('Bulan:').setFontWeight('bold');
  if (!fee.getRange('G1').getValue()) {
    fee.getRange('G1').setValue(bulanList[defaultMonthIdx]);
  }
  fee.getRange('G1').setDataValidation(ruleBulan).setHorizontalAlignment('center');

  fee.getRange('F2').setValue('Tahun:').setFontWeight('bold');
  if (!fee.getRange('G2').getValue()) {
    fee.getRange('G2').setValue(defaultYear);
  }
  fee.getRange('G2').setHorizontalAlignment('center');

  // Judul di A1:B1 & Formula Periode Cut-Off di A2:C2
  try {
    fee.getRange('A1:B1').merge();
  } catch (err) {}
  fee.getRange('A1').setValue('Perhitungan Fee').setFontWeight('bold');

  try {
    fee.getRange('A2:C2').merge();
  } catch (err) {}
  const fA2 = '="Periode : 21 " & CHOOSE(IF(VALUE(LEFT(G1, FIND(" ", G1 & " ") - 1))=1, 12, VALUE(LEFT(G1, FIND(" ", G1 & " ") - 1)) - 1), "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember") & " - 20 " & CHOOSE(VALUE(LEFT(G1, FIND(" ", G1 & " ") - 1)), "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember") & " " & G2';
  fee.getRange('A2').setFormula(fA2);

  // 3. Header Tabel di Baris 4
  const headerTitles = ['No', 'Name', 'Status', 'Cabang', 'Tarif', 'Total Fee', 'Total Fee Manual', 'Days', 'No Rekening', 'Bank', 'Keterangan'];
  fee.getRange('A4:K4').setValues([headerTitles])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(10);

  // Tanggal cut-off L4:AP4 (31 kolom tanggal)
  const dateRowFormulas = [];
  dateRowFormulas.push('=DATE($G$2, VALUE(LEFT($G$1, FIND(" ", $G$1 & " ") - 1)) - 1, 21)');
  for (let c = 13; c <= 42; c++) {
    let prevLetter = '';
    let temp = c - 1;
    while (temp > 0) {
      let rem = (temp - 1) % 26;
      prevLetter = String.fromCharCode(65 + rem) + prevLetter;
      temp = Math.floor((temp - rem) / 26);
    }
    dateRowFormulas.push('=IF(' + prevLetter + '4="", "", IF(' + prevLetter + '4 < DATE($G$2, VALUE(LEFT($G$1, FIND(" ", $G$1 & " ") - 1)), 20), ' + prevLetter + '4 + 1, ""))');
  }
  fee.getRange('L4:AP4').setFormulas([dateRowFormulas])
    .setNumberFormat('dd/mm')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#334155')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(9);

  // 4. Deteksi Nama-Nama yang ada di Baris 5 ke bawah
  const dbSheet = fee.getSheetByName ? (fee.getParent().getSheetByName('Database') || fee.getParent().getSheetByName('Freelancer')) : null;
  const dbName = dbSheet ? dbSheet.getName() : 'Database';

  let names = [];
  const maxSearchRows = Math.min(fee.getLastRow(), 100);
  if (maxSearchRows >= 5) {
    const rawColB = fee.getRange(5, 2, maxSearchRows - 4, 1).getValues();
    const rawColA = fee.getRange(5, 1, maxSearchRows - 4, 1).getValues();
    for (let i = 0; i < rawColB.length; i++) {
      const valA = String(rawColA[i][0] || '').trim();
      const valB = String(rawColB[i][0] || '').trim();
      if (valA.toLowerCase() === 'jumlah' || valB.toLowerCase() === 'jumlah') {
        break;
      }
      if (valB && !valB.startsWith('=')) {
        names.push(valB);
      }
    }
  }
  if (names.length === 0) {
    names = ['Mikhael Yobel Saputra'];
  }

  const numPeople = names.length;
  const lastDataRow = 4 + numPeople;
  const jumlahRow = lastDataRow + 1;

  // 5. Isi Formula untuk Setiap Orang (Baris 5 s/d lastDataRow)
  for (let idx = 0; idx < numPeople; idx++) {
    const r = 5 + idx;
    const name = names[idx];
    const rowBg = (idx % 2 === 0) ? '#FFFFFF' : '#F8FAFC';
    fee.getRange(r, 1, 1, 42).setBackground(rowBg).setFontFamily('Arial').setFontSize(9);

    fee.getRange('A' + r).setValue(idx + 1).setHorizontalAlignment('center');
    fee.getRange('B' + r).setValue(name).setHorizontalAlignment('center').setFontWeight('bold');
    fee.getRange('C' + r).setFormula(`=IF(B${r}="", "", IFERROR(VLOOKUP(B${r}, '${dbName}'!$A:$I, 4, FALSE), "-"))`).setHorizontalAlignment('center');
    fee.getRange('D' + r).setFormula(`=IF(B${r}="", "", IFERROR(VLOOKUP(B${r}, '${dbName}'!$A:$I, 5, FALSE), "Pusat"))`).setHorizontalAlignment('center');
    fee.getRange('E' + r).setFormula(`=IF(B${r}="", "", IFERROR(IF(VLOOKUP(B${r}, '${dbName}'!$A:$I, 7, FALSE)="-", 0, VLOOKUP(B${r}, '${dbName}'!$A:$I, 7, FALSE)), 0))`).setNumberFormat('Rp #,##0').setHorizontalAlignment('center').setFontWeight('bold');
    fee.getRange('F' + r).setFormula(`=IF(OR(B${r}="", NOT(ISNUMBER(E${r}))), 0, E${r} * H${r})`).setNumberFormat('Rp #,##0').setHorizontalAlignment('center').setFontWeight('bold');
    fee.getRange('G' + r).setFormula(`=IF(OR(B${r}="", NOT(ISNUMBER(E${r}))), 0, E${r} * H${r})`).setNumberFormat('Rp #,##0').setHorizontalAlignment('center').setFontWeight('bold');
    fee.getRange('H' + r).setFormula(`=IF(B${r}="", "", COUNTIF(L${r}:AP${r}, "Hadir"))`).setNumberFormat('0').setHorizontalAlignment('center').setFontWeight('bold');
    fee.getRange('I' + r).setFormula(`=IF(B${r}="", "", IFERROR(VLOOKUP(B${r}, '${dbName}'!$A:$I, 8, FALSE), "-"))`).setHorizontalAlignment('center');
    fee.getRange('J' + r).setFormula(`=IF(B${r}="", "", IFERROR(VLOOKUP(B${r}, '${dbName}'!$A:$I, 9, FALSE), "-"))`).setHorizontalAlignment('center').setFontWeight('bold');
    fee.getRange('K' + r).setValue('');

    // Attendance L..AP
    const attFormulas = [];
    for (let c = 12; c <= 42; c++) {
      let colLet = '';
      let t = c;
      while (t > 0) {
        let rem = (t - 1) % 26;
        colLet = String.fromCharCode(65 + rem) + colLet;
        t = Math.floor((t - rem) / 26);
      }
      attFormulas.push(`=IF(OR($B${r}="", ${colLet}$4=""), "", IF(COUNTIFS(Absensi!$A:$A, ${colLet}$4, Absensi!$B:$B, $B${r}) > 0, "Hadir", ""))`);
    }
    fee.getRange(r, 12, 1, 31).setFormulas([attFormulas]).setHorizontalAlignment('center');
  }

  // 6. Baris Jumlah
  fee.getRange('A' + jumlahRow).setValue('Jumlah').setFontWeight('bold').setHorizontalAlignment('center');
  try {
    fee.getRange(`A${jumlahRow}:E${jumlahRow}`).merge();
  } catch (err) {}
  fee.getRange(`A${jumlahRow}:AP${jumlahRow}`).setBackground('#F1F5F9').setFontFamily('Arial').setFontSize(9).setFontWeight('bold');
  fee.getRange('F' + jumlahRow).setFormula(`=SUM(F5:F${lastDataRow})`).setNumberFormat('Rp #,##0').setHorizontalAlignment('center');
  fee.getRange('G' + jumlahRow).setFormula(`=SUM(G5:G${lastDataRow})`).setNumberFormat('Rp #,##0').setHorizontalAlignment('center');

  // Kolom Tanggal pada Baris Jumlah
  const sumAttFormulas = [];
  for (let c = 12; c <= 42; c++) {
    let colLet = '';
    let t = c;
    while (t > 0) {
      let rem = (t - 1) % 26;
      colLet = String.fromCharCode(65 + rem) + colLet;
      t = Math.floor((t - rem) / 26);
    }
    if (numPeople === 1) {
      sumAttFormulas.push(`=${colLet}5`);
    } else {
      sumAttFormulas.push(`=COUNTIF(${colLet}5:${colLet}${lastDataRow}, "Hadir")`);
    }
  }
  fee.getRange(jumlahRow, 12, 1, 31).setFormulas([sumAttFormulas]).setHorizontalAlignment('center');

  if (numPeople === 1) {
    try {
      fee.getRange(`H5:H6`).merge();
      fee.getRange(`I5:I6`).merge();
      fee.getRange(`J5:J6`).merge();
      fee.getRange(`K5:K6`).merge();
    } catch (err) {}
  }

  // Border solid rapi A4:AP{jumlahRow}
  fee.getRange(`A4:AP${jumlahRow}`).setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
  fee.getRange(`A4:AP4`).setBorder(true, true, true, true, true, true, '#1E293B', SpreadsheetApp.BorderStyle.SOLID);
  fee.getRange(`A${jumlahRow}:AP${jumlahRow}`).setBorder(true, true, true, true, true, true, '#1E293B', SpreadsheetApp.BorderStyle.SOLID);

  // 7. Tanggal Bandung
  const dateRow = jumlahRow + 2;
  const fDate = '="Bandung, 21 " & CHOOSE(VALUE(LEFT(G1, FIND(" ", G1 & " ") - 1)), "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember") & " " & G2';
  fee.getRange('A' + dateRow).setFormula(fDate).setHorizontalAlignment('left').setFontStyle('italic').setFontColor('#475569');

  // 8. Blok Tanda Tangan
  const ttdTitleRow = dateRow + 2;
  fee.getRange(`B${ttdTitleRow}:F${ttdTitleRow}`).setValues([['Di buat oleh', '', 'Mengetahui', '', 'Menyetujui']]).setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#1E293B');

  const imgRow = ttdTitleRow + 1;
  fee.getRange('B' + imgRow).setFormula('=IMAGE("https://lh3.googleusercontent.com/d/1h2OmvURnGv6EPKhy-VnqAug_Kd71gUAa", 1)');
  try {
    fee.getRange(`B${imgRow}:B${imgRow + 2}`).merge();
  } catch (err) {}

  const nameRow = imgRow + 3;
  fee.getRange(`B${nameRow}:F${nameRow}`).setValues([['Surya Tio Pratama', '', 'Indra Gunawan', '', 'Benard Silalahi']]).setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#0F172A');

  // Bersihkan baris sisa di bawah blok TTD jika ada bekas sebelumnya
  try {
    const trailingStart = nameRow + 2;
    if (fee.getLastRow() >= trailingStart) {
      fee.getRange(`A${trailingStart}:AP${fee.getLastRow()}`).clearContent().clearFormat();
    }
  } catch (err) {}
}

/**
 * Trigger otomatis saat cell di Spreadsheet diedit
 */
function onEdit(e) {
  if (!e || !e.range) return;
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== 'Fee') return;
    const row = e.range.getRow();
    const col = e.range.getColumn();
    // G1 (Bulan), G2 (Tahun)
    if ((row === 1 || row === 2) && col === 7) {
      renderFeeSheet(sheet);
    }
  } catch (err) {
    console.error(err);
  }
}

/**
 * Menu kustom di Google Sheets
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Presensi & Fee')
    .addItem('Perbarui Sheet Fee', 'menuRenderFee')
    .addToUi();
}

function menuRenderFee() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fee = ss.getSheetByName('Fee');
  if (fee) {
    renderFeeSheet(fee);
    SpreadsheetApp.getActiveSpreadsheet().toast('Sheet Fee berhasil diperbarui!', 'Presensi', 3);
  }

  // Hapus sheet default bawaan (Sheet1 / Sheet 1) jika ada
  const defaultSheets = ['Sheet1', 'Sheet 1', 'Sheet'];
  for (const name of defaultSheets) {
    const s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch (err) {}
    }
  }

  // Aktifkan sheet Rekap
  ss.setActiveSheet(r);
}

/**
 * Fungsi proses absensi (Masuk / Pulang)
 * Mendukung autentikasi ID dan PIN
 */
function absen(idUser, pin, aksi, lat, lng) {
  // Dukungan kompatibilitas jika dipanggil dengan 4 parameter lama: absen(pin, aksi, lat, lng)
  if (lat === undefined && lng === undefined && (aksi === undefined || typeof aksi === 'number')) {
    lng = lat;
    lat = aksi;
    aksi = pin;
    pin = idUser;
  }

  idUser = String(idUser || '').trim();
  pin = String(pin || '').trim();
  aksi = String(aksi || '').toLowerCase().trim();

  if (!idUser) {
    throw new Error('ID harus diisi.');
  }
  if (!pin) {
    throw new Error('PIN harus diisi.');
  }
  if (aksi !== 'masuk' && aksi !== 'pulang') {
    throw new Error('Aksi tidak valid.');
  }

  // 1. Cek radius lokasi dengan rumus Haversine
  const jarak = jarakMeter(LOKASI.lat, LOKASI.lng, Number(lat), Number(lng));
  if (isNaN(jarak) || jarak > RADIUS_M) {
    const jarakStr = isNaN(jarak) ? '-' : Math.round(jarak);
    throw new Error(`Anda berada di luar lokasi kerja (jarak: ${jarakStr} meter). Maksimal radius ${RADIUS_M} meter.`);
  }

  // 2. Cek user di Sheet "Database" (atau fallback "Freelancer")
  const ss = getSS();
  const shFreelancer = ss.getSheetByName('Database') || ss.getSheetByName('Freelancer');
  if (!shFreelancer) {
    throw new Error('Sheet "Database" belum dibuat. Buka ulang web app atau jalankan setup().');
  }

  const dataFreelancer = shFreelancer.getDataRange().getDisplayValues().slice(1);
  const q = idUser.toLowerCase();
  const orang = dataFreelancer.find(row => {
    const rowNama = (row[0] || '').trim().toLowerCase();
    const rowId = (row[1] || '').trim().toLowerCase();
    return q === rowId || q === rowNama;
  });

  if (!orang || !orang[0].trim()) {
    throw new Error(`ID "${idUser}" tidak terdaftar.`);
  }

  // Cek PIN: kolom C (index 2), fallback ke kolom I (index 8) atau kolom B (index 1)
  const expectedPin = (orang[2] && orang[2].trim()) ? orang[2].trim() : ((orang[8] && orang[8].trim()) ? orang[8].trim() : orang[1].trim());
  if (pin !== expectedPin) {
    throw new Error(`PIN salah untuk ${orang[0].trim()}. Silakan periksa kembali.`);
  }

  const nama = orang[0].trim();

  // 3. LockService agar tidak bentrok saat banyak orang absen bersamaan
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const shAbsensi = ss.getSheetByName('Absensi');
    if (!shAbsensi) {
      throw new Error('Sheet "Absensi" belum dibuat. Buka ulang web app atau jalankan setup().');
    }

    const rows = shAbsensi.getDataRange().getDisplayValues();
    const now = new Date();
    const hariDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const hariIni1 = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    const hariIni2 = Utilities.formatDate(now, TZ, 'dd/MM/yyyy');
    const jamTeks = Utilities.formatDate(now, TZ, 'HH.mm');

    // Cari riwayat baris absensi terbaru untuk freelancer ini (baris teratas = terbaru)
    let idx = -1;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][1] === nama) {
        idx = i;
        break;
      }
    }
    const terakhir = idx > 0 ? rows[idx] : null;
    const isHariIni = terakhir && (
      terakhir[0] === hariIni1 ||
      terakhir[0] === hariIni2 ||
      terakhir[0].startsWith(hariIni1) ||
      terakhir[0].startsWith(hariIni2)
    );

    if (aksi === 'masuk') {
      if (isHariIni) {
        throw new Error('Sudah absen masuk hari ini.');
      }
      // Sisipkan di baris 2 (paling atas di bawah header agar data terbaru selalu di atas)
      shAbsensi.insertRowAfter(1);
      shAbsensi.getRange(2, 1, 1, 4).setValues([[hariDate, nama, now, '']]);
      shAbsensi.getRange('E2').setFormula('=IF(OR(C2="", D2=""), "", ROUND(IF(D2 < C2, D2 + 1 - C2, D2 - C2) * 24, 2))');
      shAbsensi.getRange('A2').setNumberFormat('dd/mm/yyyy');
      shAbsensi.getRange('C2:D2').setNumberFormat('HH.mm');
      shAbsensi.getRange('E2').setNumberFormat('0.00');
      return {
        success: true,
        message: `Halo ${nama}, absensi MASUK berhasil dicatat pukul ${jamTeks} WIB.`,
        nama: nama,
        jam: jamTeks,
        tipe: 'masuk'
      };
    }

    // Aksi pulang
    if (!terakhir || !isHariIni) {
      throw new Error('Belum absen masuk hari ini.');
    }
    if (terakhir[3]) {
      throw new Error('Sudah absen pulang hari ini.');
    }

    // Hitung jam kerja otomatis akurat
    const rowNum = idx + 1;
    const masukVal = shAbsensi.getRange(rowNum, 3).getValue();
    let masukJam = 0;
    if (masukVal instanceof Date) {
      masukJam = masukVal.getHours() + masukVal.getMinutes() / 60 + masukVal.getSeconds() / 3600;
    } else if (typeof masukVal === 'number') {
      masukJam = masukVal * 24;
    } else {
      const parts = String(masukVal || '').replace('.', ':').split(':');
      masukJam = (parseInt(parts[0], 10) || 0) + (parseInt(parts[1], 10) || 0) / 60;
    }

    const nowJam = now.getHours() + now.getMinutes() / 60 + now.getSeconds() / 3600;
    let selisih = nowJam - masukJam;
    if (selisih < 0) selisih += 24;
    const jamKerja = Math.max(0, Math.round(selisih * 100) / 100);

    // Set waktu pulang di kolom D dan rumus akurat di kolom E
    shAbsensi.getRange(rowNum, 4).setValue(now).setNumberFormat('HH.mm');
    shAbsensi.getRange(rowNum, 5).setFormula(`=IF(OR(C${rowNum}="", D${rowNum}=""), "", ROUND(IF(D${rowNum} < C${rowNum}, D${rowNum} + 1 - C${rowNum}, D${rowNum} - C${rowNum}) * 24, 2))`).setNumberFormat('0.00');
    return {
      success: true,
      message: `Halo ${nama}, absensi PULANG berhasil dicatat pukul ${jamTeks} WIB. Total jam kerja: ${jamKerja} jam.`,
      nama: nama,
      jam: jamTeks,
      jamKerja: jamKerja,
      tipe: 'pulang'
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Rumus Haversine: menghitung jarak antara 2 koordinat (hasil dalam meter)
 */
function jarakMeter(lat1, lng1, lat2, lng2) {
  const rad = d => d * Math.PI / 180;
  const a = Math.sin(rad(lat2 - lat1) / 2) ** 2 +
    Math.cos(rad(lat1)) * Math.cos(rad(lat2)) * Math.sin(rad(lng2 - lng1) / 2) ** 2;
  return 2 * 6371000 * Math.asin(Math.sqrt(a));
}

