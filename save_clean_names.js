const fs = require('fs');

const TOKEN_PATH = 'C:\\Users\\Tio\\.gemini\\antigravity\\plugin_data\\gsheets\\token.json';
const SPREADSHEET_ID = '1BNZnczmcVutERAM0z_G9NNvU1eGt9yzkEpYF4OjbaYY';

const CLEAN_FREELANCERS = [
  { nama: 'Mikhael Yobel Saputra', pin: '1111' },
  { nama: 'Agung Rizki Pangestu', pin: '2002' },
  { nama: 'Azka Azkia', pin: '2003' },
  { nama: 'Claudio David Julian Kurnia', pin: '2004' },
  { nama: 'Faris Zaidan Sentosa', pin: '2005' },
  { nama: 'Grace Caldera Situmorang', pin: '2006' },
  { nama: 'Ilham Ramadhan', pin: '2008' },
  { nama: 'Kevin Gultom', pin: '2010' },
  { nama: 'Luthfiandi Rizki Kurnia', pin: '2011' },
  { nama: 'Muhammad Ichsan Khairullah', pin: '2007' },
  { nama: 'Muhammad Rifqi Yusupi', pin: '2015' },
  { nama: 'Muhammad Zuhdi Adam', pin: '2001' },
  { nama: 'Nizar Abdul Malik', pin: '2012' },
  { nama: 'Prisca Manurung', pin: '2013' },
  { nama: 'Rafli Handi Rojabi', pin: '2014' },
  { nama: 'Simorangkir Jonathan Syahloom', pin: '2009' },
  { nama: 'Stephanie Sussany', pin: '2016' },
  { nama: 'Syakira Karmelia', pin: '2017' }
];

async function main() {
  const tokenData = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));
  const accessToken = tokenData.token.access_token;

  // 1. Save local backup file
  fs.writeFileSync('freelancers.json', JSON.stringify(CLEAN_FREELANCERS, null, 2), 'utf8');
  console.log('Saved local backup to freelancers.json');

  // 2. Clear Freelancer!A2:B30
  console.log('Clearing old Freelancer!A2:B30...');
  await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/Freelancer!A2:B30:clear`, {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + accessToken,
      'Content-Type': 'application/json'
    }
  });

  // 3. Write clean names and PINs to Freelancer!A2:B
  const rows = CLEAN_FREELANCERS.map(f => [f.nama, f.pin]);
  console.log(`Writing ${rows.length} clean freelancer names to Freelancer sheet...`);
  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/Freelancer!A2:B${rows.length + 1}?valueInputOption=USER_ENTERED`,
    {
      method: 'PUT',
      headers: {
        Authorization: 'Bearer ' + accessToken,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        values: rows
      })
    }
  );
  const jsonRes = await res.json();
  console.log('Write completed:', jsonRes.updatedRows, 'rows updated.');
}

main().catch(console.error);
