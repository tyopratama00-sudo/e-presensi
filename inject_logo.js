const fs = require('fs');
const b64 = fs.readFileSync('c:/Users/Tio/Project/Absensi/apps-script/logo.png').toString('base64');
let html = fs.readFileSync('c:/Users/Tio/Project/Absensi/apps-script/Index.html', 'utf8');

// Add .logo-wrapper and .logo-img css before h1 style
if (!html.includes('.logo-wrapper')) {
  const cssRule = `    .logo-wrapper {
      margin-bottom: 14px;
      display: flex;
      justify-content: center;
    }
    .logo-img {
      max-width: 190px;
      height: auto;
      display: block;
    }
    h1 {`;

  html = html.replace('    h1 {', cssRule);
}

// Replace card opening with logo + h1
const cardHeader = `  <div class="card">
    <div class="logo-wrapper">
      <img src="data:image/png;base64,${b64}" alt="JONAS PHOTO" class="logo-img">
    </div>
    <h1>Absensi</h1>`;

html = html.replace(/<div class="card">\s*<h1>Absensi<\/h1>/, cardHeader);

fs.writeFileSync('c:/Users/Tio/Project/Absensi/apps-script/Index.html', html, 'utf8');
console.log('Successfully injected Jonas Photo logo into Index.html!');
