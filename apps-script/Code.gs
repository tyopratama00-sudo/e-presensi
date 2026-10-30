/**
 * SISTEM ABSENSI (GOOGLE APPS SCRIPT)
 * Lokasi: Jl. Progo No. 12, Bandung
 */

/**
 * Formula: total hours in Absensi row r, parsed from "HH.mm - HH.mm" cells (C:AG).
 * Clock-in-only cells count 0. Overnight shifts wrap via MOD.
 */
function hoursFormula(r) {
  const c = `C${r}:AG${r}`;
  const n = (re) => `VALUE(REGEXEXTRACT(${c}, "${re}"))`;
  const sh = n('^([0-9]+)[.]'), sm = n('[.]([0-9]+) '), eh = n('- ([0-9]+)[.]'), em = n('[.]([0-9]+)$');
  return `=ROUND(SUM(ARRAYFORMULA(IFERROR(MOD((${eh}*60+${em}-${sh}*60-${sm})/60, 24), 0))), 2)`;
}

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
        a.getRange('D1').setValue(e.parameter.month);
      }
      if (e.parameter.year) {
        a.getRange('G1').setValue(Number(e.parameter.year));
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
    .setTitle('Jonas Photo Attendance')
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
  f.getRange('A1:I1').setValues([['Name', 'ID', 'PIN', 'Status', 'Branch', 'Department', 'Rate', 'Account No.', 'Bank']]).setFontWeight('bold');
  f.getRange('B:B').setNumberFormat('@'); // Format teks ID
  f.getRange('C:C').setNumberFormat('@'); // Format teks PIN
  f.getRange('G:G').setNumberFormat('#,##0'); // Format Tarif
  f.getRange('H:H').setNumberFormat('@'); // Format teks No Rekening
  if (f.getLastRow() < 2) {
    f.getRange('A2:I2').setValues([['Mikhael Yobel Saputra', '@mikhaelyobels', '1111', 'AI Prompter', 'Pusat', 'IT', 75000, '3790944231', 'BCA']]);
  }

  // 2. Sheet Absensi (Format Horizontal Ke Samping)
  setupAbsensiHorizontal(ss);

  // 3. Sheet Rekap (Hanya jika sheet Rekap ada)
  let r = ss.getSheetByName('Rekap');
  if (r) {

  // Baris 1: Filter Controls
  r.getRange('A1').setValue('Name filter:').setFontWeight('bold');
  if (!r.getRange('B1').getValue()) {
    r.getRange('B1').setValue('All');
  }

  r.getRange('C1').setValue('Month:').setFontWeight('bold');
  const bulanList = [
    '1 (January)', '2 (February)', '3 (March)', '4 (April)',
    '5 (May)', '6 (June)', '7 (July)', '8 (August)',
    '9 (September)', '10 (October)', '11 (November)', '12 (December)'
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
    r.getRange('D1').setDataValidation(ruleBulan).setValue(bulanList[(parseInt(r.getRange('D1').getValue(), 10) || (defaultMonthIdx + 1)) - 1]);
  }

  r.getRange('E1').setValue('Year:').setFontWeight('bold');
  let f1Val = r.getRange('F1').getValue();
  if (!f1Val || isNaN(Number(f1Val)) || Number(f1Val) < 2000) {
    r.getRange('F1').setValue(defaultYear);
  }
  r.getRange('F1').setNumberFormat('0').setFontWeight('bold');
  r.getRange('G1:J1').clearContent().clearFormat();

  // Baris 2: Keterangan Periode Cut-Off (21 Bulan Lalu s/d 20 Bulan Ini)
  r.getRange('A2').setValue('Cut-off period:').setFontWeight('bold');
  r.getRange('B2').setFormula(
    `=TEXT(DATE(F1, VALUE(LEFT(D1, FIND(" ", D1 & " ") - 1)) - 1, 21), "dd/mm/yyyy") & " to " & TEXT(DATE(F1, VALUE(LEFT(D1, FIND(" ", D1 & " ") - 1)), 20), "dd/mm/yyyy")`
  ).setFontWeight('bold').setFontColor('#15803d');

  // Sel bantu Tanggal Mulai (G2) dan Tanggal Selesai (H2) format yyyy-mm-dd
  r.getRange('G2').setFormula(`=TEXT(DATE(F1, VALUE(LEFT(D1, FIND(" ", D1 & " ") - 1)) - 1, 21), "yyyy-mm-dd")`);
  r.getRange('H2').setFormula(`=TEXT(DATE(F1, VALUE(LEFT(D1, FIND(" ", D1 & " ") - 1)), 20), "yyyy-mm-dd")`);
  r.getRange('G2:H2').setFontColor('#ffffff'); // Sembunyikan warna teks sel bantu

  // Kolom K: Dynamic list nama yang hadir pada periode cut-off terpilih
  r.getRange('K1').setValue('All');
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
  r.getRange('A4:C4').setValues([['Name', 'Total Work Days', 'Total Work Hours']])
    .setFontWeight('bold')
    .setBackground('#e2e8f0');
  r.getRange('A5').setFormula(
    `=IFERROR(IF(OR(B1="All", B1=""), QUERY(Absensi_Log_Backup!A2:E, "SELECT B, COUNT(A), SUM(E) WHERE B IS NOT NULL AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' GROUP BY B LABEL COUNT(A) '', SUM(E) ''", 0), QUERY(Absensi_Log_Backup!A2:E, "SELECT B, COUNT(A), SUM(E) WHERE B = '" & B1 & "' AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' GROUP BY B LABEL COUNT(A) '', SUM(E) ''", 0)), "")`
  );
  r.getRange('C5:C').setNumberFormat('0.00');

  // Tabel 2 (Kolom E-I): Rincian Absensi Harian (Terbaru di paling atas)
  r.getRange('E4:I4').setValues([['Date', 'Name', 'Clock In', 'Clock Out', 'Work Hours']])
    .setFontWeight('bold')
    .setBackground('#e2e8f0');
  r.getRange('E5').setFormula(
    `=IFERROR(IF(OR(B1="All", B1=""), QUERY(Absensi_Log_Backup!A2:E, "SELECT A, B, C, D, E WHERE B IS NOT NULL AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' ORDER BY A DESC, B ASC", 0), QUERY(Absensi_Log_Backup!A2:E, "SELECT A, B, C, D, E WHERE B = '" & B1 & "' AND A >= date '" & G2 & "' AND A <= date '" & H2 & "' ORDER BY A DESC", 0)), "")`
  );
  r.getRange('E5:E').setNumberFormat('dd/mm/yyyy');
  r.getRange('G5:H').setNumberFormat('HH.mm');
  r.getRange('I5:I').setNumberFormat('0.00');
  }

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
 * - Baris 4: Header Kolom No, Nama, Tanggal 21..20 (C4:AG4, 31 kolom), Total Days (AH4), Total Hours (AI4)
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
    backup.getRange('A1:E1').setValues([['Date', 'Name', 'Clock In', 'Clock Out', 'Work Hours']]).setFontWeight('bold');
  }

  // Backup data lama jika masih format kebawah di Absensi
  if (a.getLastRow() >= 1 && (a.getRange('A1').getValue() === 'Tanggal' || a.getRange('A1').getValue() === 'Date')) {
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

  const bulanList = [
    '1 (January)', '2 (February)', '3 (March)', '4 (April)',
    '5 (May)', '6 (June)', '7 (July)', '8 (August)',
    '9 (September)', '10 (October)', '11 (November)', '12 (December)'
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

  // Baca Month (D1) dan Year (G1) sesuai layout user
  let currentBulanVal = String(a.getRange('D1').getValue() || '').trim();
  const _m = parseInt(currentBulanVal, 10);
  if (_m >= 1 && _m <= 12) currentBulanVal = bulanList[_m - 1];
  if (!currentBulanVal || !bulanList.includes(currentBulanVal)) {
    currentBulanVal = bulanList[defaultMonthIdx];
  }

  let currentTahunVal = a.getRange('G1').getValue();
  if (!currentTahunVal || isNaN(Number(currentTahunVal)) || Number(currentTahunVal) < 2000) {
    currentTahunVal = defaultYear;
  }

  // JANGAN timpa baris 1-4 jika header sudah ada (melindungi perubahan tata letak manual user)
  const isHeaderSet = String(a.getRange('A1').getValue() || '').includes('JONAS PHOTO') && a.getRange('A4').getValue() === 'No';
  if (!isHeaderSet) {
    // 1. Bersihkan hanya jika sheet masih kosong/baru
    try {
      a.getRange(1, 1, 4, a.getMaxColumns()).breakApart();
      a.getRange(1, 1, 4, a.getMaxColumns()).clearContent();
    } catch (err) {}

    a.getRange('A:A').setNumberFormat('0').setHorizontalAlignment('center');

    // Header Baris 1-3: Title Box A1:B3
    try {
      a.getRange('A1:B3').merge()
        .setValue('JONAS PHOTO DIGITAL ATTENDANCE')
        .setFontWeight('bold')
        .setFontColor('#0F172A')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBackground('#EFF6FF');
    } catch (err) {}

    // Month di C1, D1:E1 merged value
    a.getRange('C1').setValue('Month:').setFontWeight('bold').setHorizontalAlignment('center').setFontColor('#334155').setBackground('#EFF6FF');
    try { a.getRange('D1:E1').merge(); } catch (err) {}
    a.getRange('D1').setDataValidation(ruleBulan).setValue(currentBulanVal)
      .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#EFF6FF').setFontColor('#1E40AF');

    // Year di F1, G1 value
    a.getRange('F1').setValue('Year:').setFontWeight('bold').setHorizontalAlignment('center').setFontColor('#334155').setBackground('#EFF6FF');
    a.getRange('G1').setNumberFormat('0').setValue(currentTahunVal)
      .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#EFF6FF').setFontColor('#1E40AF');

    // Period di C2, D2:E2 merged formula
    a.getRange('C2').setValue('Period:').setFontWeight('bold').setHorizontalAlignment('center').setFontColor('#334155').setBackground('#EFF6FF');
    try { a.getRange('D2:E2').merge(); } catch (err) {}
    a.getRange('D2').setFormula(
      '=TEXT(DATE(G1, VALUE(LEFT(D1, FIND(" ", D1 & " ") - 1)) - 1, 21), "dd/mm/yyyy") & " to " & TEXT(DATE(G1, VALUE(LEFT(D1, FIND(" ", D1 & " ") - 1)), 20), "dd/mm/yyyy")'
    ).setFontWeight('bold').setFontColor('#15803D').setHorizontalAlignment('center').setBackground('#EFF6FF');

    // Baris 3: Nama Hari (Mon..Sun)
    const dayNameFormulas = [];
    for (let c = 3; c <= 33; c++) {
      let colLet = getColLetter(c);
      dayNameFormulas.push(`=IF(${colLet}4="", "", CHOOSE(WEEKDAY(${colLet}4, 2), "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"))`);
    }
    a.getRange('C3:AG3').setFormulas([dayNameFormulas])
      .setHorizontalAlignment('center')
      .setFontColor('#64748B')
      .setFontSize(8)
      .setFontWeight('bold')
      .setBackground('#F8FAFC');
    a.getRange('AH3:AI3').setValue('').setBackground('#F8FAFC');

    // Baris 4: Header Kolom
    a.getRange('A4').setValue('No').setHorizontalAlignment('center');
    a.getRange('B4').setValue('Employee Name').setHorizontalAlignment('left');
    a.getRange('A4:B4').setBackground('#1E293B').setFontColor('#FFFFFF').setFontWeight('bold').setFontFamily('Arial').setFontSize(9.5);

    // Tanggal C4:AG4 (31 kolom) mengacu ke $G$1 (Year) dan $D$1 (Month)
    const dateFormulas = [];
    dateFormulas.push('=DATE($G$1, VALUE(LEFT($D$1, FIND(" ", $D$1 & " ") - 1)) - 1, 21)');
    for (let c = 4; c <= 33; c++) {
      let prevLetter = getColLetter(c - 1);
      dateFormulas.push('=IF(' + prevLetter + '4="", "", IF(' + prevLetter + '4 < DATE($G$1, VALUE(LEFT($D$1, FIND(" ", $D$1 & " ") - 1)), 20), ' + prevLetter + '4 + 1, ""))');
    }
    a.getRange('C4:AG4').setFormulas([dateFormulas])
      .setNumberFormat('dd/mm')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#1E293B')
      .setFontColor('#FFFFFF')
      .setFontFamily('Arial')
      .setFontSize(9);

    a.getRange('AH4').setValue('Total Days').setHorizontalAlignment('center')
      .setBackground('#0F766E').setFontColor('#FFFFFF').setFontWeight('bold').setFontFamily('Arial').setFontSize(9.5);
    a.getRange('AI4').setValue('Total Hours').setHorizontalAlignment('center')
      .setBackground('#0F766E').setFontColor('#FFFFFF').setFontWeight('bold').setFontFamily('Arial').setFontSize(9.5);
  }

  // 2. HITUNG TANGGAL CUT-OFF UNTUK FILTER BULAN
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

  // 3. IDENTIFIKASI HANYA KARYAWAN YANG HADIR PADA PERIODE BULAN TERPILIH
  const shDb = ss.getSheetByName('Database') || ss.getSheetByName('Freelancer');
  let dbNames = [];
  if (shDb && shDb.getLastRow() >= 2) {
    const rawNames = shDb.getRange(2, 1, shDb.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < rawNames.length; i++) {
      const n = String(rawNames[i][0] || '').trim();
      if (n && !dbNames.includes(n)) dbNames.push(n);
    }
  }

  const presentSet = new Set();
  const attendanceMap = {};

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

      if (tglDate >= startDate && tglDate <= endDate) {
        presentSet.add(nm);

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
        const dateKey = Utilities.formatDate(tglDate, TZ, 'yyyy-MM-dd');

        if (!attendanceMap[nm]) attendanceMap[nm] = {};
        attendanceMap[nm][dateKey] = cellText;
      }
    }
  }

  // Filter: HANYA orang yang hadir pada bulan ini yang masuk ke tabel
  let names = [];
  for (let i = 0; i < dbNames.length; i++) {
    if (presentSet.has(dbNames[i])) {
      names.push(dbNames[i]);
    }
  }
  presentSet.forEach(nm => {
    if (!names.includes(nm)) {
      names.push(nm);
    }
  });

  const numEmp = names.length;
  const previousLastRow = a.getLastRow();

  // 4. TULIS BARIS KARYAWAN AKTIF & KOTAK PRESENSI
  if (numEmp > 0) {
    for (let i = 0; i < numEmp; i++) {
      const r = 5 + i;
      const empName = names[i];
      a.getRange('A' + r).setValue(i + 1).setNumberFormat('0').setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#64748B');
      a.getRange('B' + r).setValue(empName).setFontWeight('bold').setFontColor('#0F172A').setHorizontalAlignment('left');
      a.getRange('AH' + r).setFormula(`=COUNTIFS(C${r}:AG${r}, "<>", C${r}:AG${r}, "<>OFF", C${r}:AG${r}, "<>Sakit", C${r}:AG${r}, "<>Izin")`).setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#0F766E');
      a.getRange('AI' + r).setFormula(hoursFormula(r)).setNumberFormat('0.00').setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#0F766E');
      a.getRange(r, 1, 1, 35).setFontFamily('Arial').setFontSize(9);
    }

    const todayDateStr = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    const currentHour = parseInt(Utilities.formatDate(now, TZ, 'HH'), 10);
    const isPast3PM = currentHour >= 15;

    // Baca status manual (Sakit/Izin/OFF) yang sudah ada di sheet agar tidak terhapus
    const existingManualStatus = {};
    if (a && a.getLastRow() >= 5) {
      const existingLastRow = Math.min(a.getLastRow(), 4 + numEmp);
      if (existingLastRow >= 5) {
        const rawGrid = a.getRange(5, 2, existingLastRow - 4, 32).getValues();
        for (let r = 0; r < rawGrid.length; r++) {
          const nm = String(rawGrid[r][0] || '').trim();
          if (!nm) continue;
          existingManualStatus[nm] = {};
          for (let c = 0; c < 31; c++) {
            const val = String(rawGrid[r][c + 1] || '').trim();
            if (val === 'Sakit' || val === 'Izin' || val === 'OFF') {
              if (c < cutOffDates.length) {
                const dk = Utilities.formatDate(cutOffDates[c], TZ, 'yyyy-MM-dd');
                existingManualStatus[nm][dk] = val;
              }
            }
          }
        }
      }
    }

    const gridValues = [];
    for (let r = 0; r < numEmp; r++) {
      const empName = names[r];
      const rowCells = [];
      const empAttendance = attendanceMap[empName] || {};
      const empManual = existingManualStatus[empName] || {};

      for (let c = 0; c < 31; c++) {
        if (c < cutOffDates.length) {
          const dateKey = Utilities.formatDate(cutOffDates[c], TZ, 'yyyy-MM-dd');
          const attended = empAttendance[dateKey];
          const manual = empManual[dateKey];

          if (attended) {
            rowCells.push(attended);
          } else if (manual && (manual === 'Sakit' || manual === 'Izin')) {
            rowCells.push(manual);
          } else if (dateKey < todayDateStr) {
            // Tanggal sudah lewat: jika tidak hadir maka OFF
            rowCells.push(manual || 'OFF');
          } else if (dateKey === todayDateStr) {
            // Tanggal hari ini: jika sudah jam 3 sore (15:00) ke atas dan belum hadir maka OFF, sebelum jam 3 biarkan kosong
            rowCells.push(manual || (isPast3PM ? 'OFF' : ''));
          } else {
            // Tanggal belum lewat (masa depan): kosongkan (jangan ada tanda OFF)
            rowCells.push((manual === 'Sakit' || manual === 'Izin') ? manual : '');
          }
        } else {
          rowCells.push('');
        }
      }
      gridValues.push(rowCells);
    }

    a.getRange(5, 3, numEmp, 31).setValues(gridValues).setNumberFormat('@').setHorizontalAlignment('center');

    // Data validation dropdown HANYA untuk sel yang sudah lewat dan berstatus OFF, Sakit, atau Izin (kosong/jam kerja tidak diberi dropdown)
    const ruleStatus = SpreadsheetApp.newDataValidation()
      .requireValueInList(['OFF', 'Sakit', 'Izin'], true)
      .setAllowInvalid(true)
      .build();
    const validationMatrix = [];
    for (let r = 0; r < numEmp; r++) {
      const rowValRules = [];
      for (let c = 0; c < 31; c++) {
        const v = String(gridValues[r][c] || '').trim();
        const hasStatus = v === 'OFF' || v === 'Sakit' || v === 'Izin';
        rowValRules.push(hasStatus ? ruleStatus : null);
      }
      validationMatrix.push(rowValRules);
    }
    a.getRange(5, 3, numEmp, 31).setDataValidations(validationMatrix);

    a.getRange(4, 1, numEmp + 1, 35).setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
    a.getRange('A4:AI4').setBorder(true, true, true, true, null, null, '#0F172A', SpreadsheetApp.BorderStyle.MEDIUM);
  }

  // 5. BERSIHKAN SEMUA BARIS DI BAWAH TABEL (agar tabel menciut/menyesuaikan persis jumlah orang yang hadir)
  const activeTableEnd = 4 + numEmp;
  if (previousLastRow > activeTableEnd) {
    const rowsToClear = previousLastRow - activeTableEnd;
    try {
      a.getRange(activeTableEnd + 1, 1, rowsToClear, a.getMaxColumns()).clearContent().clearFormat();
    } catch (err) {}
  }

  // Lebar kolom
  a.setColumnWidth(1, 45);   // No
  a.setColumnWidth(2, 210);  // Nama
  for (let c = 3; c <= 33; c++) {
    a.setColumnWidth(c, 100); // 31 Tanggal
  }
  a.setColumnWidth(34, 95);  // Total Days
  a.setColumnWidth(35, 95);  // Total Hours

  // Polish: no gridlines, comfortable row heights
  a.setHiddenGridlines(true);
  a.setRowHeight(4, 30);
  if (numEmp > 0) {
    a.setRowHeights(5, numEmp, 28);
    a.getRange(5, 1, numEmp, 35).setVerticalAlignment('middle');
  }
  a.getRange(1, 1, 4, 35).setVerticalAlignment('middle');

  // Freeze 2 kolom dan 4 baris header
  try {
    a.setFrozenColumns(2);
    a.setFrozenRows(4);
  } catch (err) {}

  // 6. FILTER CORONG (Native Google Sheets Filter)
  try {
    if (a.getFilter()) {
      a.getFilter().remove();
    }
    if (numEmp > 0) {
      a.getRange(4, 1, numEmp + 1, 35).createFilter();
    }
  } catch (err) {}

  // 7. CONDITIONAL FORMATTING
  a.clearConditionalFormatRules();
  if (numEmp > 0) {
    const rangeData = a.getRange(5, 3, numEmp, 31);

    const ruleOff = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OFF')
      .setBackground('#E2E8F0')
      .setFontColor('#475569')
      .setBold(true)
      .setRanges([rangeData])
      .build();

    const ruleSakit = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Sakit')
      .setBackground('#FEE2E2')
      .setFontColor('#991B1B')
      .setBold(true)
      .setRanges([rangeData])
      .build();

    const ruleIzin = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Izin')
      .setBackground('#FEF3C7')
      .setFontColor('#92400E')
      .setBold(true)
      .setRanges([rangeData])
      .build();

    const ruleComplete = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains(' - ')
      .setBackground('#DCFCE7')
      .setFontColor('#166534')
      .setBold(true)
      .setRanges([rangeData])
      .build();

    const ruleInProgress = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(C5<>"", C5<>"OFF", C5<>"Sakit", C5<>"Izin", ISERROR(FIND(" - ", C5)))')
      .setBackground('#FFFBEB')
      .setFontColor('#B45309')
      .setBold(true)
      .setRanges([rangeData])
      .build();

    // Sel belum terlewat (kosong) di akhir pekan (Sat/Sun) -> #E8EDF2
    const ruleEmptyWeekend = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(C5="", C$4<>"", OR(C$3="Sat", C$3="Sun"))')
      .setBackground('#E8EDF2')
      .setRanges([rangeData])
      .build();

    // Sel belum terlewat (kosong) di hari kerja (Mon-Fri) -> #F1F5F9 (jangan putih)
    const ruleEmptyWeekday = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(C5="", C$4<>"")')
      .setBackground('#F1F5F9')
      .setRanges([rangeData])
      .build();

    a.setConditionalFormatRules([ruleOff, ruleSakit, ruleIzin, ruleComplete, ruleInProgress, ruleEmptyWeekend, ruleEmptyWeekday]);
  }
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
    '1 (January)', '2 (February)', '3 (March)', '4 (April)',
    '5 (May)', '6 (June)', '7 (July)', '8 (August)',
    '9 (September)', '10 (October)', '11 (November)', '12 (December)'
  ];
  const ruleBulan = SpreadsheetApp.newDataValidation()
    .requireValueInList(bulanList, true)
    .build();

  // Bersihkan cell sekitar F1/F2 agar tidak ada label Month/Year duplikat
  try {
    fee.getRange('E1:E2').clearContent().clearFormat();
    fee.getRange('H1:H2').clearContent().clearFormat();
  } catch (err) {}

  fee.getRange('F1').setValue('Month:').setFontWeight('bold').setHorizontalAlignment('right');
  if (!fee.getRange('G1').getValue()) {
    fee.getRange('G1').setValue(bulanList[defaultMonthIdx]);
  }
  fee.getRange('G1').setDataValidation(ruleBulan).setValue(bulanList[(parseInt(fee.getRange('G1').getValue(), 10) || (defaultMonthIdx + 1)) - 1]).setHorizontalAlignment('center').setFontWeight('bold');

  fee.getRange('F2').setValue('Year:').setFontWeight('bold').setHorizontalAlignment('right');
  if (!fee.getRange('G2').getValue()) {
    fee.getRange('G2').setValue(defaultYear);
  }
  fee.getRange('G2').setNumberFormat('0').setHorizontalAlignment('center').setFontWeight('bold');

  // Judul di A1:B1 & Formula Periode Cut-Off di A2:C2
  try {
    fee.getRange('A1:B1').merge();
  } catch (err) {}
  fee.getRange('A1').setValue('Fee Calculation').setFontWeight('bold').setFontSize(14).setFontColor('#0F172A');

  try {
    fee.getRange('A2:C2').merge();
  } catch (err) {}
  const fA2 = '="Period: 21 " & CHOOSE(IF(VALUE(LEFT(G1, FIND(" ", G1 & " ") - 1))=1, 12, VALUE(LEFT(G1, FIND(" ", G1 & " ") - 1)) - 1), "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December") & " - 20 " & CHOOSE(VALUE(LEFT(G1, FIND(" ", G1 & " ") - 1)), "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December") & " " & G2';
  fee.getRange('A2').setFormula(fA2).setFontStyle('italic').setFontSize(10).setFontColor('#475569');

  // Baris 3: Kosongkan baris spasi
  try {
    fee.getRange('A3:AP3').breakApart().clearContent().clearFormat();
  } catch (err) {}

  // 3. Header Tabel di Baris 4
  const headerTitles = ['No', 'Name', 'Status', 'Branch', 'Rate', 'Total Fee', 'Manual Total Fee', 'Days', 'Account No.', 'Bank', 'Notes'];
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
    const prevLetter = getColLetter(c - 1);
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

  // Bersihkan merge dan baris lama dari baris 5 ke bawah agar tidak ada tabrakan merge & data lama
  try {
    const prevFeeLast = fee.getLastRow();
    if (prevFeeLast >= 5) {
      fee.getRange(5, 1, prevFeeLast - 4, fee.getMaxColumns()).breakApart();
      fee.getRange(5, 1, prevFeeLast - 4, fee.getMaxColumns()).clearContent().clearFormat();
    }
  } catch (err) {}

  // 4. Deteksi Nama-Nama yang ada di Baris 5 ke bawah
  const dbSheet = fee.getSheetByName ? (fee.getParent().getSheetByName('Database') || fee.getParent().getSheetByName('Freelancer')) : null;
  const dbName = dbSheet ? dbSheet.getName() : 'Database';

  // Ambil nama-nama karyawan yang hadir di bulan terpilih dari sheet Absensi
  const shAbsensi = fee.getParent().getSheetByName('Absensi');
  let names = [];
  if (shAbsensi && shAbsensi.getLastRow() >= 5) {
    const rawAbsensi = shAbsensi.getRange(5, 2, shAbsensi.getLastRow() - 4, 1).getValues();
    for (let i = 0; i < rawAbsensi.length; i++) {
      const n = String(rawAbsensi[i][0] || '').trim();
      if (n && !names.includes(n)) names.push(n);
    }
  }
  // Fallback ke Database jika belum ada yang hadir
  if (names.length === 0 && dbSheet && dbSheet.getLastRow() >= 2) {
    const rawDb = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < rawDb.length; i++) {
      const n = String(rawDb[i][0] || '').trim();
      if (n && !names.includes(n)) names.push(n);
    }
  }

  const numPeople = names.length;
  const lastDataRow = 4 + numPeople;
  const jumlahRow = lastDataRow + 1;

  // 5. Isi Formula untuk Setiap Orang (Baris 5 s/d lastDataRow)
  for (let idx = 0; idx < numPeople; idx++) {
    const r = 5 + idx;
    const name = names[idx];
    const rowBg = (idx % 2 === 0) ? '#FFFFFF' : '#F8FAFC';
    fee.getRange(r, 1, 1, 42).setBackground(rowBg).setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle');

    fee.getRange('A' + r).setValue(idx + 1).setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#64748B');
    fee.getRange('B' + r).setValue(name).setHorizontalAlignment('left').setFontWeight('bold').setFontColor('#0F172A');
    fee.getRange('C' + r).setFormula(`=IF(B${r}="", "", IFERROR(VLOOKUP(B${r}, '${dbName}'!$A:$I, 4, FALSE), "-"))`).setHorizontalAlignment('center');
    fee.getRange('D' + r).setFormula(`=IF(B${r}="", "", IFERROR(VLOOKUP(B${r}, '${dbName}'!$A:$I, 5, FALSE), "Pusat"))`).setHorizontalAlignment('center');
    fee.getRange('E' + r).setFormula(`=IF(B${r}="", "", IFERROR(IF(VLOOKUP(B${r}, '${dbName}'!$A:$I, 7, FALSE)="-", 0, VLOOKUP(B${r}, '${dbName}'!$A:$I, 7, FALSE)), 0))`).setNumberFormat('Rp #,##0').setHorizontalAlignment('right').setFontWeight('bold');
    fee.getRange('F' + r).setFormula(`=IF(OR(B${r}="", NOT(ISNUMBER(E${r}))), 0, E${r} * H${r})`).setNumberFormat('Rp #,##0').setHorizontalAlignment('right').setFontWeight('bold');
    fee.getRange('G' + r).setFormula(`=IF(OR(B${r}="", NOT(ISNUMBER(E${r}))), 0, E${r} * H${r})`).setNumberFormat('Rp #,##0').setHorizontalAlignment('right').setFontWeight('bold');
    fee.getRange('H' + r).setFormula(`=IF(B${r}="", "", COUNT(L${r}:AP${r}))`).setNumberFormat('0').setHorizontalAlignment('center').setFontWeight('bold');
    fee.getRange('I' + r).setFormula(`=IF(B${r}="", "", IFERROR(VLOOKUP(B${r}, '${dbName}'!$A:$I, 8, FALSE), "-"))`).setHorizontalAlignment('center');
    fee.getRange('J' + r).setFormula(`=IF(B${r}="", "", IFERROR(VLOOKUP(B${r}, '${dbName}'!$A:$I, 9, FALSE), "-"))`).setHorizontalAlignment('center').setFontWeight('bold');
    fee.getRange('K' + r).setValue('');

    // Attendance L..AP (Rate nominal harian per karyawan, kosong jika tidak hadir)
    const attFormulas = [];
    for (let c = 12; c <= 42; c++) {
      const colLet = getColLetter(c);
      attFormulas.push(`=IF(OR($B${r}="", ${colLet}$4=""), "", LET(v, IFERROR(INDEX(Absensi!$C$5:$AG$100, MATCH($B${r}, Absensi!$B$5:$B$100, 0), MATCH(${colLet}$4, Absensi!$C$4:$AG$4, 0)), ""), IF(OR(v="", v="OFF", v="Sakit", v="Izin"), "", $E${r})))`);
    }
    fee.getRange(r, 12, 1, 31).setFormulas([attFormulas]).setNumberFormat('#,##0').setHorizontalAlignment('center');
  }

  // 6. Baris Jumlah
  fee.getRange('A' + jumlahRow).setValue('Total').setFontWeight('bold').setHorizontalAlignment('center');
  try {
    fee.getRange(`A${jumlahRow}:E${jumlahRow}`).merge();
  } catch (err) {}
  fee.getRange(`A${jumlahRow}:AP${jumlahRow}`).setBackground('#F1F5F9').setFontFamily('Arial').setFontSize(9.5).setFontWeight('bold').setVerticalAlignment('middle');
  fee.getRange('F' + jumlahRow).setFormula(`=SUM(F5:F${lastDataRow})`).setNumberFormat('Rp #,##0').setHorizontalAlignment('right');
  fee.getRange('G' + jumlahRow).setFormula(`=SUM(G5:G${lastDataRow})`).setNumberFormat('Rp #,##0').setHorizontalAlignment('right');

  // Kolom Tanggal pada Baris Jumlah (Menampilkan total fee harian, jika kosong/0 maka kosongkan)
  const sumAttFormulas = [];
  for (let c = 12; c <= 42; c++) {
    const colLet = getColLetter(c);
    sumAttFormulas.push(`=IF(SUM(${colLet}5:${colLet}${lastDataRow})=0, "", SUM(${colLet}5:${colLet}${lastDataRow}))`);
  }
  fee.getRange(jumlahRow, 12, 1, 31).setFormulas([sumAttFormulas]).setNumberFormat('#,##0').setHorizontalAlignment('center');

  // Lebar Kolom
  fee.setColumnWidth(1, 45);
  fee.setColumnWidth(2, 220);
  fee.setColumnWidth(3, 110);
  fee.setColumnWidth(4, 90);
  fee.setColumnWidth(5, 100);
  fee.setColumnWidth(6, 115);
  fee.setColumnWidth(7, 125);
  fee.setColumnWidth(8, 60);
  fee.setColumnWidth(9, 120);
  fee.setColumnWidth(10, 85);
  fee.setColumnWidth(11, 100);
  for (let c = 12; c <= 42; c++) {
    fee.setColumnWidth(c, 70);
  }

  // Tinggi Baris
  fee.setRowHeight(4, 32);
  if (numPeople > 0) {
    fee.setRowHeights(5, numPeople, 28);
  }
  fee.setRowHeight(jumlahRow, 30);

  // Border solid rapi A4:AP{jumlahRow}
  fee.setHiddenGridlines(true);
  fee.getRange(`A4:AP${jumlahRow}`).setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
  fee.getRange(`A4:AP4`).setBorder(true, true, true, true, true, true, '#1E293B', SpreadsheetApp.BorderStyle.SOLID);
  fee.getRange(`A${jumlahRow}:AP${jumlahRow}`).setBorder(true, true, true, true, true, true, '#1E293B', SpreadsheetApp.BorderStyle.SOLID);

  // 7. Tanggal Bandung
  const dateRow = jumlahRow + 2;
  const fDate = '="Bandung, 21 " & CHOOSE(VALUE(LEFT(G1, FIND(" ", G1 & " ") - 1)), "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December") & " " & G2';
  fee.getRange('A' + dateRow).setFormula(fDate).setHorizontalAlignment('left').setFontStyle('italic').setFontColor('#475569');

  // 8. Blok Tanda Tangan
  const ttdTitleRow = dateRow + 2;
  fee.getRange(`B${ttdTitleRow}:F${ttdTitleRow}`).setValues([['Prepared by', '', 'Acknowledged by', '', 'Approved by']]).setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#1E293B');

  const imgRow = ttdTitleRow + 1;
  fee.getRange('B' + imgRow).setFormula('=IMAGE("https://lh3.googleusercontent.com/d/1h2OmvURnGv6EPKhy-VnqAug_Kd71gUAa", 1)');
  try {
    fee.getRange(`B${imgRow}:B${imgRow + 2}`).merge();
  } catch (err) {}

  const nameRow = imgRow + 3;
  fee.getRange(`B${nameRow}:F${nameRow}`).setValues([['Surya Tio Pratama', '', 'Indra Gunawan', '', 'Benard Silalahi']]).setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#0F172A').setFontLine('underline');

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

    // Sheet Absensi: D1 (Month, col 4) atau G1 (Year, col 7)
    if (sheetName === 'Absensi' && row === 1 && (col === 4 || col === 7)) {
      setupAbsensiHorizontal(sheet.getParent());
      const fee = sheet.getParent().getSheetByName('Fee');
      if (fee) renderFeeSheet(fee);
    }

    // Sheet Absensi: Jika kotak kehadiran (C5:AG) diedit, hanya beri dropdown jika status OFF, Sakit, atau Izin (jika jam kerja atau kosong, hapus dropdown)
    if (sheetName === 'Absensi' && row >= 5 && col >= 3 && col <= 33) {
      const val = String(e.range.getValue() || '').trim();
      const hasStatus = val === 'OFF' || val === 'Sakit' || val === 'Izin';
      if (hasStatus) {
        const ruleStatus = SpreadsheetApp.newDataValidation()
          .requireValueInList(['OFF', 'Sakit', 'Izin'], true)
          .setAllowInvalid(true)
          .build();
        e.range.setDataValidation(ruleStatus);
      } else {
        e.range.clearDataValidations();
      }
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
    .createMenu('Attendance & Fee')
    .addItem('Refresh Attendance Sheet', 'menuSetupAbsensi')
    .addItem('Refresh Fee Sheet', 'menuRenderFee')
    .addItem('Check & Mark OFF (Past 3 PM)', 'autoCheckAndMarkOff')
    .addToUi();

  try {
    autoCheckAndMarkOff();
  } catch (err) {}
}

/**
 * Otomatis mengecek dan mengisi status OFF setelah jam 3 sore (15:00 WIB)
 * Dipanggil saat spreadsheet dibuka atau via Time-driven Trigger
 */
function autoCheckAndMarkOff() {
  const ss = getSS();
  const shAbsensi = ss.getSheetByName('Absensi');
  if (!shAbsensi || shAbsensi.getLastRow() < 5) return;

  const now = new Date();
  const currentHour = parseInt(Utilities.formatDate(now, TZ, 'HH'), 10);
  if (currentHour < 15) return; // Belum jam 3 sore

  const todayStr = Utilities.formatDate(now, TZ, 'dd/MM');
  const dateRowValues = shAbsensi.getRange(4, 3, 1, 31).getValues()[0];
  let todayCol = -1;
  for (let c = 0; c < dateRowValues.length; c++) {
    const dVal = dateRowValues[c];
    let colDateStr = '';
    if (dVal instanceof Date) {
      colDateStr = Utilities.formatDate(dVal, TZ, 'dd/MM');
    } else if (typeof dVal === 'string') {
      colDateStr = dVal.trim();
    }
    if (colDateStr === todayStr) {
      todayCol = 3 + c;
      break;
    }
  }

  if (todayCol === -1) return;

  const numRows = shAbsensi.getLastRow() - 4;
  const colVals = shAbsensi.getRange(5, todayCol, numRows, 1).getValues();
  const ruleStatus = SpreadsheetApp.newDataValidation()
    .requireValueInList(['OFF', 'Sakit', 'Izin'], true)
    .setAllowInvalid(true)
    .build();

  for (let i = 0; i < numRows; i++) {
    const val = String(colVals[i][0] || '').trim();
    if (!val) {
      const cell = shAbsensi.getRange(5 + i, todayCol);
      cell.setValue('OFF').setNumberFormat('@').setHorizontalAlignment('center');
      cell.setDataValidation(ruleStatus);
    }
  }
}

function menuSetupAbsensi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupAbsensiHorizontal(ss);
  ss.toast('Attendance sheet refreshed!', 'Attendance', 3);
}

function menuRenderFee() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fee = ss.getSheetByName('Fee');
  if (fee) {
    renderFeeSheet(fee);
    ss.toast('Fee sheet refreshed!', 'Attendance', 3);
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
    throw new Error('ID is required.');
  }
  if (!pin) {
    throw new Error('PIN is required.');
  }
  if (aksi !== 'masuk' && aksi !== 'pulang') {
    throw new Error('Invalid action.');
  }

  // 1. Cek radius lokasi dengan rumus Haversine
  const jarak = jarakMeter(LOKASI.lat, LOKASI.lng, Number(lat), Number(lng));
  if (isNaN(jarak) || jarak > RADIUS_M) {
    const jarakStr = isNaN(jarak) ? '-' : Math.round(jarak);
    throw new Error(`You are outside the work location (distance: ${jarakStr} m). Maximum radius is ${RADIUS_M} m.`);
  }

  // 2. Cek user di Sheet "Database" (atau fallback "Freelancer")
  const ss = getSS();
  const shFreelancer = ss.getSheetByName('Database') || ss.getSheetByName('Freelancer');
  if (!shFreelancer) {
    throw new Error('Sheet "Database" does not exist yet. Reopen the web app or run setup().');
  }

  const dataFreelancer = shFreelancer.getDataRange().getDisplayValues().slice(1);
  const q = idUser.toLowerCase();
  const orang = dataFreelancer.find(row => {
    const rowNama = (row[0] || '').trim().toLowerCase();
    const rowId = (row[1] || '').trim().toLowerCase();
    return q === rowId || q === rowNama;
  });

  if (!orang || !orang[0].trim()) {
    throw new Error(`ID "${idUser}" is not registered.`);
  }

  // Cek PIN: kolom C (index 2), fallback ke kolom I (index 8) atau kolom B (index 1)
  const expectedPin = (orang[2] && orang[2].trim()) ? orang[2].trim() : ((orang[8] && orang[8].trim()) ? orang[8].trim() : orang[1].trim());
  if (pin !== expectedPin) {
    throw new Error(`Wrong PIN for ${orang[0].trim()}. Please check and try again.`);
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
      throw new Error(`Today's date (${tglStr}) is not in the cut-off date range of the Absensi sheet.`);
    }

    // 5. Cari baris karyawan di Kolom B (mulai baris 5)
    const currentTableEnd = shAbsensi.getLastRow();
    let targetRow = -1;

    if (currentTableEnd >= 5) {
      const names = shAbsensi.getRange(5, 2, currentTableEnd - 4, 1).getValues();
      for (let r = 0; r < names.length; r++) {
        if (String(names[r][0] || '').trim().toLowerCase() === nama.toLowerCase()) {
          targetRow = 5 + r;
          break;
        }
      }
    }

    // Jika nama belum ada di tabel bulan ini, tambahkan baris baru untuk karyawan ini
    if (targetRow === -1) {
      targetRow = (currentTableEnd >= 5 ? currentTableEnd : 4) + 1;
      const noUrut = targetRow - 4;
      shAbsensi.getRange(targetRow, 1).setValue(noUrut).setNumberFormat('0').setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#64748B');
      shAbsensi.getRange(targetRow, 2).setValue(nama).setFontWeight('bold').setFontColor('#0F172A').setHorizontalAlignment('left');
      shAbsensi.getRange(targetRow, 34).setFormula(`=COUNTIFS(C${targetRow}:AG${targetRow}, "<>", C${targetRow}:AG${targetRow}, "<>OFF", C${targetRow}:AG${targetRow}, "<>Sakit", C${targetRow}:AG${targetRow}, "<>Izin")`).setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#0F766E');
      shAbsensi.getRange(targetRow, 35).setFormula(hoursFormula(targetRow)).setNumberFormat('0.00').setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#0F766E');
      shAbsensi.getRange(targetRow, 1, 1, 35).setFontFamily('Arial').setFontSize(9);
      shAbsensi.getRange(targetRow, 1, 1, 35).setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);

      // Inisialisasi sel C..AG dengan 'OFF' dan validation dropdown
      const defaultCells = Array(31).fill('OFF');
      shAbsensi.getRange(targetRow, 3, 1, 31).setValues([defaultCells]).setNumberFormat('@').setHorizontalAlignment('center');
      const ruleStatus = SpreadsheetApp.newDataValidation()
        .requireValueInList(['OFF', 'Sakit', 'Izin'], true)
        .setAllowInvalid(true)
        .build();
      shAbsensi.getRange(targetRow, 3, 1, 31).setDataValidation(ruleStatus);

      try {
        if (shAbsensi.getFilter()) shAbsensi.getFilter().remove();
        shAbsensi.getRange(4, 1, targetRow - 3, 35).createFilter();
      } catch (err) {}
    }

    const targetCell = shAbsensi.getRange(targetRow, targetCol);
    const currentVal = String(targetCell.getValue() || '').trim();
    const isNotClockedIn = !currentVal || currentVal === 'OFF' || currentVal === 'Sakit' || currentVal === 'Izin';

    if (aksi === 'masuk') {
      if (!isNotClockedIn) {
        throw new Error(`You have already clocked in today (${currentVal}).`);
      }
      targetCell.setValue(jamTeks).setNumberFormat('@').setHorizontalAlignment('center');
      targetCell.clearDataValidations();

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
        message: `Hi ${nama}, your CLOCK IN was recorded at ${jamTeks} WIB.`,
        nama: nama,
        jam: jamTeks,
        tipe: 'masuk'
      };
    }

    // Aksi pulang
    if (isNotClockedIn) {
      throw new Error('You have not clocked in today.');
    }
    if (currentVal.includes(' - ')) {
      throw new Error(`You have already clocked out today (${currentVal}).`);
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
      message: `Hi ${nama}, your CLOCK OUT was recorded at ${jamTeks} WIB. Total hours worked: ${jamKerja} h.`,
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


// ============================================================
// WEBAUTHN BIOMETRIC SUPPORT
// ============================================================

/**
 * Ambil/buat sheet WebAuthn untuk menyimpan credential terdaftar.
 */
function getWebAuthnSheet_(ss) {
  let sh = ss.getSheetByName('WebAuthn');
  if (!sh) {
    sh = ss.insertSheet('WebAuthn');
    sh.getRange('A1:E1').setValues([['credentialId', 'idUser', 'nama', 'registeredAt', 'origin']]);
    sh.getRange('A1:E1').setFontWeight('bold');
    sh.setColumnWidth(1, 300);
    sh.setColumnWidth(2, 160);
    sh.setColumnWidth(3, 200);
    sh.setColumnWidth(4, 160);
    sh.setColumnWidth(5, 260);
  }
  return sh;
}

/**
 * Hasilkan random challenge (base64url 32 byte) dan simpan ke ScriptProperties
 * dengan TTL 5 menit.
 */
function getWebAuthnChallenge() {
  const bytes = [];
  for (let i = 0; i < 32; i++) bytes.push(Math.floor(Math.random() * 256));
  const challenge = Utilities.base64EncodeWebSafe(bytes).replace(/=/g, '');
  const props = PropertiesService.getScriptProperties();
  const key = 'wac_' + challenge.substring(0, 10);
  props.setProperty(key, JSON.stringify({ challenge: challenge, ts: Date.now() }));
  return challenge;
}

/**
 * Validasi challenge dari clientDataJSON (base64url-encoded).
 * Menghapus challenge setelah berhasil divalidasi (one-time use).
 */
function verifyWebAuthnChallenge_(clientDataJSON) {
  var clientData;
  try {
    var decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(clientDataJSON)).getDataAsString();
    clientData = JSON.parse(decoded);
  } catch (e) {
    throw new Error('clientDataJSON tidak valid.');
  }

  var challenge = clientData.challenge;
  if (!challenge) throw new Error('Challenge tidak ditemukan di clientData.');

  var props = PropertiesService.getScriptProperties();
  var key = 'wac_' + challenge.substring(0, 10);
  var stored = props.getProperty(key);
  if (!stored) throw new Error('Challenge sudah kedaluwarsa atau tidak valid. Silakan coba lagi.');

  var storedData = JSON.parse(stored);
  if (storedData.challenge !== challenge) throw new Error('Challenge tidak cocok.');
  if (Date.now() - storedData.ts > 5 * 60 * 1000) {
    props.deleteProperty(key);
    throw new Error('Challenge kedaluwarsa. Silakan coba lagi.');
  }

  props.deleteProperty(key); // One-time use
  return clientData.origin || '';
}

/**
 * Daftarkan credential biometrik (WebAuthn) untuk user yang sudah terverifikasi PIN.
 * Dipanggil setelah absen PIN berhasil, sebagai step opt-in registrasi biometrik.
 */
function registerWebAuthnCredential(idUser, pin, credentialId, clientDataJSON) {
  idUser = String(idUser || '').trim();
  pin = String(pin || '').trim();
  credentialId = String(credentialId || '').trim();

  if (!idUser || !pin || !credentialId || !clientDataJSON) {
    throw new Error('Parameter tidak lengkap untuk registrasi biometrik.');
  }

  var ss = getSS();
  var dbSheet = ss.getSheetByName('Database') || ss.getSheetByName('Freelancer');
  if (!dbSheet) throw new Error('Sheet Database tidak ditemukan.');

  var rows = dbSheet.getDataRange().getDisplayValues().slice(1);
  var q = idUser.toLowerCase();
  var orang = null;
  for (var i = 0; i < rows.length; i++) {
    var rId = (rows[i][1] || '').trim().toLowerCase();
    var rNama = (rows[i][0] || '').trim().toLowerCase();
    if (q === rId || q === rNama) { orang = rows[i]; break; }
  }

  if (!orang || !orang[0].trim()) throw new Error('ID tidak terdaftar.');
  var expectedPin = (orang[2] && orang[2].trim()) ? orang[2].trim() : orang[1].trim();
  if (pin !== expectedPin) throw new Error('PIN salah. Registrasi biometrik ditolak.');

  var nama = orang[0].trim();

  // Verifikasi challenge
  verifyWebAuthnChallenge_(clientDataJSON);

  // Simpan credential (update jika sudah ada)
  var waSh = getWebAuthnSheet_(ss);
  var waData = waSh.getDataRange().getValues();
  var existingRow = -1;
  for (var j = 1; j < waData.length; j++) {
    if (String(waData[j][0]) === credentialId) { existingRow = j + 1; break; }
  }

  var now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  if (existingRow > 0) {
    waSh.getRange(existingRow, 1, 1, 5).setValues([[credentialId, idUser, nama, now, 'updated']]);
  } else {
    waSh.appendRow([credentialId, idUser, nama, now, 'registered']);
  }

  return { success: true, message: 'Biometrik berhasil didaftarkan untuk ' + nama + '.', nama: nama };
}

/**
 * Absen menggunakan WebAuthn tanpa input ID/PIN.
 * credentialId dicari di sheet WebAuthn → resolve idUser → rekam absensi.
 */
function absenWithWebAuthn(credentialId, clientDataJSON, aksi, lat, lng) {
  credentialId = String(credentialId || '').trim();
  aksi = String(aksi || '').toLowerCase().trim();

  if (!credentialId || !clientDataJSON) throw new Error('Data biometrik tidak lengkap.');
  if (aksi !== 'masuk' && aksi !== 'pulang') throw new Error('Aksi tidak valid.');

  // 1. Verifikasi challenge
  verifyWebAuthnChallenge_(clientDataJSON);

  // 2. Cari credential di sheet WebAuthn
  var ss = getSS();
  var waSh = getWebAuthnSheet_(ss);
  var waData = waSh.getDataRange().getValues();

  var idUser = null;
  var nama = null;
  for (var i = 1; i < waData.length; i++) {
    if (String(waData[i][0]) === credentialId) {
      idUser = String(waData[i][1]);
      nama = String(waData[i][2]);
      break;
    }
  }

  if (!idUser) throw new Error('Perangkat ini belum terdaftar biometrik. Gunakan ID + PIN terlebih dahulu, lalu aktifkan biometrik.');

  // 3. Verifikasi lokasi
  var jarak = jarakMeter(LOKASI.lat, LOKASI.lng, Number(lat), Number(lng));
  if (isNaN(jarak) || jarak > RADIUS_M) {
    var jarakStr = isNaN(jarak) ? '-' : Math.round(jarak);
    throw new Error('Anda berada di luar lokasi kerja (jarak: ' + jarakStr + ' m). Maksimal radius ' + RADIUS_M + ' m.');
  }

  // 4. Rekam absensi (tanpa cek PIN)
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var shAbsensi = ss.getSheetByName('Absensi');
    if (!shAbsensi) {
      setupAbsensiHorizontal(ss);
      shAbsensi = ss.getSheetByName('Absensi');
    }

    var now = new Date();
    var jamTeks = Utilities.formatDate(now, TZ, 'HH.mm');
    var todayDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    var dateRowValues = shAbsensi.getRange(4, 3, 1, 31).getValues()[0];
    var targetCol = -1;
    for (var c = 0; c < dateRowValues.length; c++) {
      var d = dateRowValues[c];
      var dObj = (d instanceof Date) ? d : new Date(d);
      if (!isNaN(dObj.getTime()) &&
          dObj.getFullYear() === todayDate.getFullYear() &&
          dObj.getMonth() === todayDate.getMonth() &&
          dObj.getDate() === todayDate.getDate()) {
        targetCol = 3 + c; break;
      }
    }
    if (targetCol === -1) {
      var tglStr = Utilities.formatDate(now, TZ, 'dd/MM/yyyy');
      throw new Error('Tanggal hari ini (' + tglStr + ') tidak ada di sheet Absensi.');
    }

    var currentTableEnd = shAbsensi.getLastRow();
    var targetRow = -1;
    if (currentTableEnd >= 5) {
      var names = shAbsensi.getRange(5, 2, currentTableEnd - 4, 1).getValues();
      for (var r = 0; r < names.length; r++) {
        if (String(names[r][0] || '').trim().toLowerCase() === nama.toLowerCase()) {
          targetRow = 5 + r; break;
        }
      }
    }

    if (targetRow === -1) {
      targetRow = (currentTableEnd >= 5 ? currentTableEnd : 4) + 1;
      var noUrut = targetRow - 4;
      shAbsensi.getRange(targetRow, 1).setValue(noUrut).setNumberFormat('0').setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#64748B');
      shAbsensi.getRange(targetRow, 2).setValue(nama).setFontWeight('bold').setFontColor('#0F172A').setHorizontalAlignment('left');
      shAbsensi.getRange(targetRow, 34).setFormula('=COUNTIFS(C' + targetRow + ':AG' + targetRow + ', "<>", C' + targetRow + ':AG' + targetRow + ', "<>OFF", C' + targetRow + ':AG' + targetRow + ', "<>Sakit", C' + targetRow + ':AG' + targetRow + ', "<>Izin")').setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#0F766E');
      shAbsensi.getRange(targetRow, 35).setFormula(hoursFormula(targetRow)).setNumberFormat('0.00').setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#0F766E');
      shAbsensi.getRange(targetRow, 1, 1, 35).setFontFamily('Arial').setFontSize(9);
      shAbsensi.getRange(targetRow, 1, 1, 35).setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
      var defaultCells = [];
      for (var dc = 0; dc < 31; dc++) defaultCells.push('OFF');
      shAbsensi.getRange(targetRow, 3, 1, 31).setValues([defaultCells]).setNumberFormat('@').setHorizontalAlignment('center');
      var ruleStatus = SpreadsheetApp.newDataValidation().requireValueInList(['OFF', 'Sakit', 'Izin'], true).setAllowInvalid(true).build();
      shAbsensi.getRange(targetRow, 3, 1, 31).setDataValidation(ruleStatus);
      try { if (shAbsensi.getFilter()) shAbsensi.getFilter().remove(); shAbsensi.getRange(4, 1, targetRow - 3, 35).createFilter(); } catch (ex) {}
    }

    var targetCell = shAbsensi.getRange(targetRow, targetCol);
    var currentVal = String(targetCell.getValue() || '').trim();
    var isNotClockedIn = !currentVal || currentVal === 'OFF' || currentVal === 'Sakit' || currentVal === 'Izin';

    if (aksi === 'masuk') {
      if (!isNotClockedIn) throw new Error('Anda sudah Clock In hari ini (' + currentVal + ').');
      targetCell.setValue(jamTeks).setNumberFormat('@').setHorizontalAlignment('center');
      targetCell.clearDataValidations();
      var shBackup = ss.getSheetByName('Absensi_Log_Backup');
      if (shBackup) {
        shBackup.insertRowAfter(1);
        shBackup.getRange(2, 1, 1, 4).setValues([[todayDate, nama, now, '']]);
        shBackup.getRange('A2').setNumberFormat('dd/mm/yyyy');
        shBackup.getRange('C2:D2').setNumberFormat('HH.mm');
        shBackup.getRange('E2').setFormula('=IF(OR(C2="", D2=""), "", ROUND(IF(D2 < C2, D2 + 1 - C2, D2 - C2) * 24, 2))').setNumberFormat('0.00');
      }
      return { success: true, message: 'Hi ' + nama + ', CLOCK IN tercatat pukul ' + jamTeks + ' WIB. [🔒 Biometrik ✓]', nama: nama, jam: jamTeks, tipe: 'masuk', biometric: true };
    }

    // Aksi pulang
    if (isNotClockedIn) throw new Error('Anda belum Clock In hari ini.');
    if (currentVal.includes(' - ')) throw new Error('Anda sudah Clock Out hari ini (' + currentVal + ').');

    var jamMasukParts = currentVal.replace('.', ':').split(':');
    var masukHours = (parseInt(jamMasukParts[0], 10) || 0) + (parseInt(jamMasukParts[1], 10) || 0) / 60;
    var nowHours = now.getHours() + now.getMinutes() / 60 + now.getSeconds() / 3600;
    var selisih = nowHours - masukHours;
    if (selisih < 0) selisih += 24;
    var jamKerja = Math.max(0, Math.round(selisih * 100) / 100);

    var valBaru = currentVal + ' - ' + jamTeks;
    targetCell.setValue(valBaru).setNumberFormat('@').setHorizontalAlignment('center');

    var shBackupOut = ss.getSheetByName('Absensi_Log_Backup');
    if (shBackupOut && shBackupOut.getLastRow() >= 2) {
      var maxScan = Math.min(shBackupOut.getLastRow() - 1, 100);
      var backupRows = shBackupOut.getRange(2, 1, maxScan, 2).getValues();
      var hariIniStr = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
      for (var bi = 0; bi < backupRows.length; bi++) {
        var tgl = backupRows[bi][0];
        var tglStr2 = (tgl instanceof Date) ? Utilities.formatDate(tgl, TZ, 'yyyy-MM-dd') : String(tgl);
        if (backupRows[bi][1] === nama && tglStr2.startsWith(hariIniStr)) {
          shBackupOut.getRange(2 + bi, 4).setValue(now).setNumberFormat('HH.mm');
          break;
        }
      }
    }

    return { success: true, message: 'Hi ' + nama + ', CLOCK OUT tercatat pukul ' + jamTeks + ' WIB. Total: ' + jamKerja + ' jam. [🔒 Biometrik ✓]', nama: nama, jam: jamTeks, jamKerja: jamKerja, tipe: 'pulang', biometric: true };

  } finally {
    lock.releaseLock();
  }
}


/**
 * Endpoint POST untuk melayani permintaan API dari Vercel / Client Eksternal
 */
function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      try {
        body = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        body = e.parameter || {};
      }
    } else if (e && e.parameter) {
      body = e.parameter;
    }

    var action = String(body.action || '').trim();
    var res;

    if (action === 'getWebAuthnChallenge') {
      res = { success: true, challenge: getWebAuthnChallenge() };
    } else if (action === 'registerWebAuthnCredential') {
      res = registerWebAuthnCredential(body.idUser, body.pin, body.credentialId, body.clientDataJSON);
    } else if (action === 'absenWithWebAuthn') {
      res = absenWithWebAuthn(body.credentialId, body.clientDataJSON, body.aksi, body.lat, body.lng);
    } else if (action === 'absen') {
      res = absen(body.idUser, body.pin, body.aksi, body.lat, body.lng);
    } else {
      throw new Error('Aksi API tidak valid: ' + action);
    }

    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message || String(err)
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
