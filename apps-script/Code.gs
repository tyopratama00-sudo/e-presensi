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

function doGet(e) {
  if (e && e.parameter && e.parameter.setup === '1') {
    const ss = getSS();
    const a = ss.getSheetByName('Absensi');
    if (a) {
      if (e.parameter.month) {
        a.getRange('B1').setValue(e.parameter.month);
      }
      if (e.parameter.year) {
        a.getRange('D1').setValue(Number(e.parameter.year));
      }
    }
    setup();
    return ContentService.createTextOutput('SETUP_HORIZONTAL_SUCCESS');
  }

  if (e && e.parameter && e.parameter.inspect === '1') {
    const ss = getSS();
    const sheets = ss.getSheets();
    const result = {
      sheetNames: sheets.map(s => s.getName()),
      sheetsData: {}
    };
    sheets.forEach(s => {
      const name = s.getName();
      const numRows = Math.min(s.getLastRow(), 20);
      const numCols = Math.min(s.getLastColumn(), 36);
      if (numRows > 0 && numCols > 0) {
        result.sheetsData[name] = {
          lastRow: s.getLastRow(),
          lastCol: s.getLastColumn(),
          values: s.getRange(1, 1, numRows, numCols).getDisplayValues(),
          formulas: s.getRange(1, 1, Math.min(numRows, 6), Math.min(numCols, 15)).getFormulas()
        };
      } else {
        result.sheetsData[name] = { lastRow: 0, lastCol: 0, values: [] };
      }
    });
    return ContentService.createTextOutput(JSON.stringify(result, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Presensi Jonas Photo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
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

  // 2. Sheet Absensi (Format Horizontal Ke Samping)
  setupAbsensiHorizontal(ss);

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
  let f1Val = r.getRange('F1').getValue();
  if (!f1Val || isNaN(Number(f1Val)) || Number(f1Val) < 2000) {
    r.getRange('F1').setValue(defaultYear);
  }
  r.getRange('F1').setNumberFormat('0').setFontWeight('bold');
  r.getRange('G1:J1').clearContent().clearFormat();

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
    `=IFERROR(UNIQUE(QUERY(Absensi_Log_Backup!A2:B, "SELECT B WHERE B IS NOT NULL AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' ORDER BY B LABEL B ''", 0)), "")`
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
    `=IFERROR(IF(OR(B1="Semua", B1=""), QUERY(Absensi_Log_Backup!A2:E, "SELECT B, COUNT(A), SUM(E) WHERE B IS NOT NULL AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' GROUP BY B LABEL COUNT(A) '', SUM(E) ''", 0), QUERY(Absensi_Log_Backup!A2:E, "SELECT B, COUNT(A), SUM(E) WHERE B = '" & B1 & "' AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' GROUP BY B LABEL COUNT(A) '', SUM(E) ''", 0)), "")`
  );
  r.getRange('C5:C').setNumberFormat('0.00');

  // Tabel 2 (Kolom E-I): Rincian Absensi Harian (Terbaru di paling atas)
  r.getRange('E4:I4').setValues([['Tanggal', 'Nama', 'Masuk', 'Pulang', 'Jam Kerja']])
    .setFontWeight('bold')
    .setBackground('#e2e8f0');
  r.getRange('E5').setFormula(
    `=IFERROR(IF(OR(B1="Semua", B1=""), QUERY(Absensi_Log_Backup!A2:E, "SELECT A, B, C, D, E WHERE B IS NOT NULL AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' ORDER BY A DESC, B ASC", 0), QUERY(Absensi_Log_Backup!A2:E, "SELECT A, B, C, D, E WHERE B = '" & B1 & "' AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' ORDER BY A DESC", 0)), "")`
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
 * Inisialisasi dan format Sheet Absensi Horizontal (Ke Samping):
 * - Baris 1: Filter Bulan & Tahun
 * - Baris 2: Keterangan Periode Cut-Off (21 bulan lalu s/d 20 bulan ini)
 * - Baris 4: Header Kolom No, Nama, Tanggal 21..20 (C4:AG4, 31 kolom), Total Hadir (AH4)
 * - Baris 5+: Data karyawan (Nama di kolom B, kotak absen di kolom C..AG)
 * - Format jam di kotak absen: "08.00" (Masuk) atau "08.00 - 17.00" (Lengkap)
 */
/**
 * Konversi index kolom (1-indexed) ke huruf kolom (A, B, ..., Z, AA, AB, ...)
 */
function getColLetter(colIndex) {
  let temp = colIndex;
  let letter = '';
  while (temp > 0) {
    let rem = (temp - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    temp = Math.floor((temp - rem) / 26);
  }
  return letter;
}

function setupAbsensiHorizontal(ss) {
  if (!ss) ss = getSS();
  let a = ss.getSheetByName('Absensi');
  if (!a) {
    a = ss.insertSheet('Absensi');
  }

  // Pastikan sheet backup log ada
  let backup = ss.getSheetByName('Absensi_Log_Backup');
  if (!backup) {
    backup = ss.insertSheet('Absensi_Log_Backup');
    backup.getRange('A1:E1').setValues([['Tanggal', 'Nama', 'Masuk', 'Pulang', 'Jam Kerja']]).setFontWeight('bold');
  }

  // Backup data lama jika masih format kebawah di Absensi
  if (a.getLastRow() >= 1 && a.getRange('A1').getValue() === 'Tanggal') {
    if (a.getLastRow() >= 2) {
      const oldVals = a.getRange(2, 1, a.getLastRow() - 1, 5).getValues();
      const bLast = backup.getLastRow();
      backup.getRange(bLast + 1, 1, oldVals.length, 5).setValues(oldVals);
    }
    a.clear();
  }

  // Pastikan kolom cukup (minimal 36 kolom: A..AJ)
  if (a.getMaxColumns() < 36) {
    a.insertColumnsAfter(a.getMaxColumns(), 36 - a.getMaxColumns());
  }

  // 1. Reset format dasar baris 1-2
  a.getRange('A:A').setNumberFormat('0').setHorizontalAlignment('center');
  a.getRange('C1:D1').setNumberFormat('0').setHorizontalAlignment('center');

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

  // Header Baris 1: Filter Bulan & Tahun
  a.getRange('A1').setValue('Bulan:').setFontWeight('bold').setHorizontalAlignment('right').setFontColor('#334155');
  let currentBulanVal = String(a.getRange('B1').getValue() || '').trim();
  if (!currentBulanVal || !bulanList.includes(currentBulanVal)) {
    a.getRange('B1').setDataValidation(ruleBulan).setValue(bulanList[defaultMonthIdx]);
    currentBulanVal = bulanList[defaultMonthIdx];
  } else {
    a.getRange('B1').setDataValidation(ruleBulan);
  }
  a.getRange('B1').setFontWeight('bold').setHorizontalAlignment('left').setBackground('#EFF6FF').setFontColor('#1E40AF');

  a.getRange('C1').setValue('Tahun:').setFontWeight('bold').setHorizontalAlignment('right').setFontColor('#334155');
  let currentTahunVal = a.getRange('D1').getValue();
  if (!currentTahunVal || isNaN(Number(currentTahunVal)) || Number(currentTahunVal) < 2000) {
    a.getRange('D1').setValue(defaultYear);
    currentTahunVal = defaultYear;
  }
  a.getRange('D1').setNumberFormat('0').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#EFF6FF').setFontColor('#1E40AF');

  try {
    a.getRange('E1:G1').merge().setValue('PRESENSI DIGITAL JONAS PHOTO')
      .setFontWeight('bold').setFontColor('#0F172A').setFontSize(11).setHorizontalAlignment('left');
  } catch (err) {}

  // Baris 2: Keterangan Periode Cut-Off
  a.getRange('A2').setValue('Periode:').setFontWeight('bold').setHorizontalAlignment('right').setFontColor('#334155');
  try {
    a.getRange('B2:G2').merge().setFormula(
      '=TEXT(DATE(D1, VALUE(LEFT(B1, FIND(" ", B1 & " ") - 1)) - 1, 21), "dd/mm/yyyy") & " s/d " & TEXT(DATE(D1, VALUE(LEFT(B1, FIND(" ", B1 & " ") - 1)), 20), "dd/mm/yyyy")'
    ).setFontWeight('bold').setFontColor('#15803D').setHorizontalAlignment('left');
  } catch (err) {}

  // Baris 3: Nama Hari (Sen, Sel, Rab, Kam, Jum, Sab, Min)
  a.getRange('A3:B3').setValues([['', '']]).setBackground('#F8FAFC');
  const dayNameFormulas = [];
  for (let c = 3; c <= 33; c++) {
    let colLet = getColLetter(c);
    dayNameFormulas.push(`=IF(${colLet}4="", "", CHOOSE(WEEKDAY(${colLet}4, 2), "Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"))`);
  }
  a.getRange('C3:AG3').setFormulas([dayNameFormulas])
    .setHorizontalAlignment('center')
    .setFontColor('#64748B')
    .setFontSize(8)
    .setFontWeight('bold')
    .setBackground('#F8FAFC');
  a.getRange('AH3').setValue('').setBackground('#F8FAFC');

  // Baris 4: Header Kolom
  a.getRange('A4').setValue('No').setHorizontalAlignment('center');
  a.getRange('B4').setValue('Nama Karyawan').setHorizontalAlignment('left');
  a.getRange('A4:B4').setBackground('#1E293B').setFontColor('#FFFFFF').setFontWeight('bold').setFontFamily('Arial').setFontSize(9.5);

  // Tanggal C4:AG4 (31 kolom)
  const dateFormulas = [];
  dateFormulas.push('=DATE($D$1, VALUE(LEFT($B$1, FIND(" ", $B$1 & " ") - 1)) - 1, 21)');
  for (let c = 4; c <= 33; c++) {
    let prevLetter = getColLetter(c - 1);
    dateFormulas.push('=IF(' + prevLetter + '4="", "", IF(' + prevLetter + '4 < DATE($D$1, VALUE(LEFT($B$1, FIND(" ", $B$1 & " ") - 1)), 20), ' + prevLetter + '4 + 1, ""))');
  }
  a.getRange('C4:AG4').setFormulas([dateFormulas])
    .setNumberFormat('dd/mm')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#1E293B')
    .setFontColor('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(9);

  a.getRange('AH4').setValue('Total Hadir').setHorizontalAlignment('center')
    .setBackground('#0F766E').setFontColor('#FFFFFF').setFontWeight('bold').setFontFamily('Arial').setFontSize(9.5);

  // 2. Daftar Nama Karyawan dari Database
  const shDb = ss.getSheetByName('Database') || ss.getSheetByName('Freelancer');
  let names = [];
  if (shDb && shDb.getLastRow() >= 2) {
    const rawNames = shDb.getRange(2, 1, shDb.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < rawNames.length; i++) {
      const n = String(rawNames[i][0] || '').trim();
      if (n && !names.includes(n)) {
        names.push(n);
      }
    }
  }
  if (names.length === 0) {
    names = ['Mikhael Yobel Saputra'];
  }

  // Tulis baris karyawan mulai baris 5
  const numEmp = names.length;
  for (let i = 0; i < numEmp; i++) {
    const r = 5 + i;
    a.getRange('A' + r).setValue(i + 1).setNumberFormat('0').setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#64748B');
    a.getRange('B' + r).setValue(names[i]).setFontWeight('bold').setFontColor('#0F172A').setHorizontalAlignment('left');
    a.getRange('AH' + r).setFormula(`=COUNTIF(C${r}:AG${r}, "<>")`).setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#0F766E');
    a.getRange(r, 1, 1, 34).setFontFamily('Arial').setFontSize(9);
  }

  // Bersihkan baris berlebih jika ada
  const maxRow = a.getLastRow();
  if (maxRow > 4 + numEmp) {
    try {
      a.getRange(5 + numEmp, 1, maxRow - (4 + numEmp), 34).clearContent().clearFormat();
    } catch (err) {}
  }

  // 3. ISI KOTAK PRESENSI C5:AG DARI MASTER LOG (Absensi_Log_Backup) SESUAI BULAN TERPILIH
  const selectedMonthNum = parseInt(currentBulanVal.split(' ')[0], 10) || (defaultMonthIdx + 1);
  const selectedYearNum = parseInt(currentTahunVal, 10) || defaultYear;

  const startDate = new Date(selectedYearNum, selectedMonthNum - 2, 21);
  const endDate = new Date(selectedYearNum, selectedMonthNum - 1, 20, 23, 59, 59);

  const cutOffDates = [];
  let cur = new Date(startDate.getTime());
  while (cur <= endDate && cutOffDates.length < 31) {
    cutOffDates.push(new Date(cur.getTime()));
    cur.setDate(cur.getDate() + 1);
  }

  const targetGrid = a.getRange(5, 3, numEmp, 31);
  const gridValues = [];
  for (let r = 0; r < numEmp; r++) {
    const rowCells = [];
    for (let c = 0; c < 31; c++) {
      rowCells.push('');
    }
    gridValues.push(rowCells);
  }

  if (backup && backup.getLastRow() >= 2) {
    const logData = backup.getRange(2, 1, backup.getLastRow() - 1, 4).getValues();
    for (let i = 0; i < logData.length; i++) {
      const rec = logData[i];
      let tgl = rec[0];
      const nm = String(rec[1] || '').trim();
      if (!tgl || !nm) continue;

      let tglDate = null;
      if (tgl instanceof Date) {
        tglDate = tgl;
      } else if (typeof tgl === 'string') {
        const parts = tgl.includes('/') ? tgl.split('/') : tgl.split('-');
        if (parts.length === 3) {
          if (parts[0].length === 4) {
            tglDate = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
          } else {
            tglDate = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
          }
        }
      }
      if (!tglDate || isNaN(tglDate.getTime())) continue;

      if (tglDate < startDate || tglDate > endDate) continue;

      const empIdx = names.indexOf(nm);
      if (empIdx === -1) continue;

      for (let c = 0; c < cutOffDates.length; c++) {
        const cod = cutOffDates[c];
        if (cod.getFullYear() === tglDate.getFullYear() &&
            cod.getMonth() === tglDate.getMonth() &&
            cod.getDate() === tglDate.getDate()) {

          let msk = '';
          if (rec[2] instanceof Date) {
            msk = Utilities.formatDate(rec[2], TZ, 'HH.mm');
          } else if (rec[2]) {
            msk = String(rec[2]).trim().replace(':', '.');
          }

          let plg = '';
          if (rec[3] instanceof Date) {
            plg = Utilities.formatDate(rec[3], TZ, 'HH.mm');
          } else if (rec[3]) {
            plg = String(rec[3]).trim().replace(':', '.');
          }

          const cellText = plg ? `${msk} - ${plg}` : msk;
          gridValues[empIdx][c] = cellText;
          break;
        }
      }
    }
  }

  targetGrid.setValues(gridValues).setNumberFormat('@').setHorizontalAlignment('center');

  // 4. BORDERS & STYLING
  a.getRange(4, 1, numEmp + 1, 34).setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
  a.getRange('A4:AH4').setBorder(true, true, true, true, null, null, '#0F172A', SpreadsheetApp.BorderStyle.MEDIUM);

  // Lebar kolom
  a.setColumnWidth(1, 45);   // No
  a.setColumnWidth(2, 210);  // Nama
  for (let c = 3; c <= 33; c++) {
    a.setColumnWidth(c, 100); // 31 Tanggal
  }
  a.setColumnWidth(34, 95);  // Total Hadir

  // Freeze 2 kolom dan 4 baris header
  try {
    a.setFrozenColumns(2);
    a.setFrozenRows(4);
  } catch (err) {}

  // 5. FILTER CORONG (Native Google Sheets Filter)
  try {
    if (a.getFilter()) {
      a.getFilter().remove();
    }
    a.getRange(4, 1, numEmp + 1, 34).createFilter();
  } catch (err) {}

  // 6. CONDITIONAL FORMATTING
  a.clearConditionalFormatRules();
  const ruleComplete = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains(' - ')
    .setBackground('#DCFCE7')
    .setFontColor('#166534')
    .setRanges([a.getRange(5, 3, numEmp, 31)])
    .build();

  const ruleInProgress = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(C5<>"", ISERROR(FIND(" - ", C5)))')
    .setBackground('#FEF3C7')
    .setFontColor('#92400E')
    .setRanges([a.getRange(5, 3, numEmp, 31)])
    .build();

  a.setConditionalFormatRules([ruleComplete, ruleInProgress]);
}
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
      attFormulas.push(`=IF(OR($B${r}="", ${colLet}$4=""), "", IF(IFERROR(INDEX(Absensi!$C$5:$AG$100, MATCH($B${r}, Absensi!$B$5:$B$100, 0), MATCH(${colLet}$4, Absensi!$C$4:$AG$4, 0)))<>"", "Hadir", ""))`);
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
    const sheetName = sheet.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();

    // Sheet Absensi: B1 (Bulan) atau D1 (Tahun)
    if (sheetName === 'Absensi' && row === 1 && (col === 2 || col === 4)) {
      setupAbsensiHorizontal(sheet.getParent());
      const fee = sheet.getParent().getSheetByName('Fee');
      if (fee) renderFeeSheet(fee);
    }

    // Sheet Fee: G1 (Bulan) atau G2 (Tahun)
    if (sheetName === 'Fee' && (row === 1 || row === 2) && col === 7) {
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
    .addItem('Perbarui Format Absensi Horizontal', 'menuSetupAbsensi')
    .addItem('Perbarui Sheet Fee', 'menuRenderFee')
    .addToUi();
}

function menuSetupAbsensi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupAbsensiHorizontal(ss);
  ss.toast('Sheet Absensi (Horizontal) berhasil diperbarui!', 'Presensi', 3);
}

function menuRenderFee() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fee = ss.getSheetByName('Fee');
  if (fee) {
    renderFeeSheet(fee);
    ss.toast('Sheet Fee berhasil diperbarui!', 'Presensi', 3);
    ss.setActiveSheet(fee);
  }

  // Hapus sheet default bawaan (Sheet1 / Sheet 1) jika ada
  const defaultSheets = ['Sheet1', 'Sheet 1', 'Sheet'];
  for (const name of defaultSheets) {
    const s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch (err) {}
    }
  }
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
    let shAbsensi = ss.getSheetByName('Absensi');
    if (!shAbsensi) {
      setupAbsensiHorizontal(ss);
      shAbsensi = ss.getSheetByName('Absensi');
    }

    const now = new Date();
    const jamTeks = Utilities.formatDate(now, TZ, 'HH.mm');
    const todayDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    // 4. Cari kolom tanggal hari ini di Baris 4 (Kolom C s/d AG, 31 tanggal cut-off)
    const dateRowValues = shAbsensi.getRange(4, 3, 1, 31).getValues()[0];
    let targetCol = -1;

    for (let c = 0; c < dateRowValues.length; c++) {
      const d = dateRowValues[c];
      if (d instanceof Date) {
        if (d.getFullYear() === todayDate.getFullYear() &&
            d.getMonth() === todayDate.getMonth() &&
            d.getDate() === todayDate.getDate()) {
          targetCol = 3 + c;
          break;
        }
      } else if (d) {
        const dObj = new Date(d);
        if (!isNaN(dObj.getTime()) &&
            dObj.getFullYear() === todayDate.getFullYear() &&
            dObj.getMonth() === todayDate.getMonth() &&
            dObj.getDate() === todayDate.getDate()) {
          targetCol = 3 + c;
          break;
        }
      }
    }

    if (targetCol === -1) {
      const tglStr = Utilities.formatDate(now, TZ, 'dd/MM/yyyy');
      throw new Error(`Tanggal hari ini (${tglStr}) belum ada pada rentang tanggal cut-off sheet Absensi.`);
    }

    // 5. Cari baris karyawan di Kolom B (mulai baris 5)
    const lastRow = Math.max(shAbsensi.getLastRow(), 5);
    const names = shAbsensi.getRange(5, 2, lastRow - 4, 1).getValues();
    let targetRow = -1;

    for (let r = 0; r < names.length; r++) {
      if (String(names[r][0] || '').trim().toLowerCase() === nama.toLowerCase()) {
        targetRow = 5 + r;
        break;
      }
    }

    // Jika nama belum terdaftar di tabel, tambahkan baris baru
    if (targetRow === -1) {
      targetRow = lastRow + 1;
      const noUrut = targetRow - 4;
      shAbsensi.getRange(targetRow, 1).setValue(noUrut).setHorizontalAlignment('center');
      shAbsensi.getRange(targetRow, 2).setValue(nama).setFontWeight('bold');
      shAbsensi.getRange(targetRow, 34).setFormula(`=COUNTIF(C${targetRow}:AG${targetRow}, "<>")`).setHorizontalAlignment('center').setFontWeight('bold');
    }

    const targetCell = shAbsensi.getRange(targetRow, targetCol);
    const currentVal = String(targetCell.getValue() || '').trim();

    if (aksi === 'masuk') {
      if (currentVal) {
        throw new Error(`Sudah absen masuk hari ini (${currentVal}).`);
      }
      targetCell.setValue(jamTeks).setNumberFormat('@').setHorizontalAlignment('center');

      // Simpan log ke master log Absensi_Log_Backup
      const shBackup = ss.getSheetByName('Absensi_Log_Backup');
      if (shBackup) {
        shBackup.insertRowAfter(1);
        shBackup.getRange(2, 1, 1, 4).setValues([[todayDate, nama, now, '']]);
        shBackup.getRange('A2').setNumberFormat('dd/mm/yyyy');
        shBackup.getRange('C2:D2').setNumberFormat('HH.mm');
        shBackup.getRange('E2').setFormula('=IF(OR(C2="", D2=""), "", ROUND(IF(D2 < C2, D2 + 1 - C2, D2 - C2) * 24, 2))').setNumberFormat('0.00');
      }

      return {
        success: true,
        message: `Halo ${nama}, absensi MASUK berhasil dicatat pukul ${jamTeks} WIB.`,
        nama: nama,
        jam: jamTeks,
        tipe: 'masuk'
      };
    }

    // Aksi pulang
    if (!currentVal) {
      throw new Error('Belum absen masuk hari ini.');
    }
    if (currentVal.includes(' - ')) {
      throw new Error(`Sudah absen pulang hari ini (${currentVal}).`);
    }

    // Hitung jam kerja
    const jamMasukStr = currentVal.replace('.', ':');
    const parts = jamMasukStr.split(':');
    const masukHours = (parseInt(parts[0], 10) || 0) + (parseInt(parts[1], 10) || 0) / 60;
    const nowHours = now.getHours() + now.getMinutes() / 60 + now.getSeconds() / 3600;
    let selisih = nowHours - masukHours;
    if (selisih < 0) selisih += 24;
    const jamKerja = Math.max(0, Math.round(selisih * 100) / 100);

    const valBaru = `${currentVal} - ${jamTeks}`;
    targetCell.setValue(valBaru).setNumberFormat('@').setHorizontalAlignment('center');

    // Update waktu pulang di master log Absensi_Log_Backup
    const shBackup = ss.getSheetByName('Absensi_Log_Backup');
    if (shBackup && shBackup.getLastRow() >= 2) {
      const maxScan = Math.min(shBackup.getLastRow() - 1, 100);
      const rows = shBackup.getRange(2, 1, maxScan, 2).getValues();
      const hariIniStr = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
      for (let i = 0; i < rows.length; i++) {
        const tgl = rows[i][0];
        let tglStr = '';
        if (tgl instanceof Date) {
          tglStr = Utilities.formatDate(tgl, TZ, 'yyyy-MM-dd');
        } else if (typeof tgl === 'string') {
          tglStr = tgl;
        }
        if (rows[i][1] === nama && (tglStr === hariIniStr || tglStr.startsWith(hariIniStr))) {
          shBackup.getRange(2 + i, 4).setValue(now).setNumberFormat('HH.mm');
          break;
        }
      }
    }

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

