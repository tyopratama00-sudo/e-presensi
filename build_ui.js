const fs = require('fs');

const current = fs.readFileSync('apps-script/Index.html', 'utf8');
const logoMatch = current.match(/data:image\/png;base64,[A-Za-z0-9+/=]+/);
if (!logoMatch) {
  console.error('Logo not found!');
  process.exit(1);
}
const logoBase64 = logoMatch[0];

const newHtml = `<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
  <title>Presensi Jonas Photo</title>
  <base target="_top">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    :root {
      --font-main: "Plus Jakarta Sans", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      --color-primary: #2563eb;
      --color-primary-dark: #1d4ed8;
      --color-masuk: #059669;
      --color-masuk-hover: #047857;
      --color-pulang: #e11d48;
      --color-pulang-hover: #be123c;
      --color-bg: #f8fafc;
      --color-surface: #ffffff;
      --color-text-main: #0f172a;
      --color-text-muted: #64748b;
      --color-border: #e2e8f0;
      --radius-card: 24px;
      --radius-element: 14px;
    }

    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      -webkit-tap-highlight-color: transparent;
    }

    body {
      font-family: var(--font-main);
      background-color: var(--color-bg);
      background-image:
        radial-gradient(at 10% 20%, rgba(37, 99, 235, 0.07) 0px, transparent 40%),
        radial-gradient(at 90% 80%, rgba(5, 150, 105, 0.07) 0px, transparent 40%),
        radial-gradient(at 50% 50%, rgba(225, 29, 72, 0.04) 0px, transparent 50%);
      color: var(--color-text-main);
      display: flex;
      justify-content: center;
      align-items: center;
      min-height: 100vh;
      padding: 20px 16px;
      position: relative;
      overflow-x: hidden;
    }

    /* Ambient Decorative Glows */
    .ambient-glow-1,
    .ambient-glow-2 {
      position: fixed;
      width: 320px;
      height: 320px;
      border-radius: 50%;
      filter: blur(80px);
      pointer-events: none;
      z-index: 0;
      opacity: 0.6;
    }
    .ambient-glow-1 {
      top: -80px;
      left: -80px;
      background: #38bdf8;
    }
    .ambient-glow-2 {
      bottom: -80px;
      right: -80px;
      background: #34d399;
    }

    .card {
      position: relative;
      z-index: 1;
      background: var(--color-surface);
      width: 100%;
      max-width: 390px;
      padding: 32px 26px;
      border-radius: var(--radius-card);
      box-shadow:
        0 20px 40px -15px rgba(15, 23, 42, 0.08),
        0 0 0 1px rgba(226, 232, 240, 0.85);
      text-align: center;
      animation: cardEntrance 0.5s cubic-bezier(0.16, 1, 0.3, 1) forwards;
    }

    @keyframes cardEntrance {
      from {
        opacity: 0;
        transform: translateY(16px) scale(0.98);
      }
      to {
        opacity: 1;
        transform: translateY(0) scale(1);
      }
    }

    /* Logo Header */
    .brand-header {
      display: flex;
      flex-direction: column;
      align-items: center;
      margin-bottom: 22px;
    }
    .logo-container {
      display: flex;
      justify-content: center;
      align-items: center;
      padding: 4px 12px;
      margin-bottom: 8px;
      transition: transform 0.2s ease;
    }
    .logo-container:hover {
      transform: scale(1.02);
    }
    .logo-img {
      max-width: 185px;
      height: auto;
      display: block;
    }
    .badge-sub {
      display: inline-block;
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 2px;
      text-transform: uppercase;
      color: #64748b;
      background: #f1f5f9;
      padding: 4px 10px;
      border-radius: 999px;
      margin-top: 4px;
    }

    /* Clock & Date Widget */
    .time-widget {
      background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
      border: 1px solid var(--color-border);
      border-radius: var(--radius-element);
      padding: 16px 14px;
      margin-bottom: 24px;
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 4px;
    }
    .date-text {
      font-size: 12px;
      font-weight: 600;
      color: var(--color-text-muted);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .time-row {
      display: flex;
      align-items: baseline;
      gap: 8px;
    }
    .clock-display {
      font-size: 32px;
      font-weight: 800;
      letter-spacing: -0.5px;
      color: var(--color-text-main);
      font-variant-numeric: tabular-nums;
      line-height: 1.1;
    }
    .tz-badge {
      background: #e0e7ff;
      color: #3730a3;
      font-size: 11px;
      font-weight: 700;
      padding: 2px 7px;
      border-radius: 6px;
    }

    /* Input Form Group */
    .form-group {
      text-align: left;
      margin-bottom: 20px;
    }
    .label-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 8px;
    }
    label {
      font-size: 13px;
      font-weight: 700;
      color: #334155;
    }
    .label-hint {
      font-size: 11px;
      color: var(--color-text-muted);
    }
    .pin-input-wrapper {
      position: relative;
      display: flex;
      align-items: center;
    }
    input[type="password"],
    input[type="text"] {
      width: 100%;
      height: 54px;
      padding: 0 48px 0 16px;
      font-family: var(--font-main);
      font-size: 24px;
      font-weight: 700;
      letter-spacing: 6px;
      text-align: center;
      color: var(--color-text-main);
      background: #ffffff;
      border: 2px solid var(--color-border);
      border-radius: var(--radius-element);
      outline: none;
      transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1);
    }
    input:focus {
      border-color: var(--color-primary);
      box-shadow: 0 0 0 4px rgba(37, 99, 235, 0.12);
    }
    input::placeholder {
      color: #cbd5e1;
      letter-spacing: 4px;
      font-weight: 400;
    }
    .btn-toggle-eye {
      position: absolute;
      right: 12px;
      width: 32px;
      height: 32px;
      display: flex;
      align-items: center;
      justify-content: center;
      background: transparent;
      border: none;
      border-radius: 8px;
      color: #94a3b8;
      cursor: pointer;
      transition: all 0.2s;
    }
    .btn-toggle-eye:hover {
      color: var(--color-text-main);
      background: #f1f5f9;
    }
    .btn-toggle-eye svg {
      width: 18px;
      height: 18px;
    }

    /* Action Buttons */
    .btn-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 12px;
      margin-bottom: 20px;
    }
    .btn-action {
      height: 52px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      font-family: var(--font-main);
      font-size: 15px;
      font-weight: 700;
      color: #ffffff;
      border: none;
      border-radius: var(--radius-element);
      cursor: pointer;
      transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1);
      position: relative;
      overflow: hidden;
    }
    .btn-action svg {
      width: 18px;
      height: 18px;
      flex-shrink: 0;
    }
    .btn-masuk {
      background: linear-gradient(135deg, #10b981 0%, #059669 100%);
      box-shadow: 0 8px 18px -4px rgba(16, 185, 129, 0.4);
    }
    .btn-masuk:hover:not(:disabled) {
      background: linear-gradient(135deg, #059669 0%, #047857 100%);
      transform: translateY(-2px);
      box-shadow: 0 12px 22px -4px rgba(16, 185, 129, 0.48);
    }
    .btn-pulang {
      background: linear-gradient(135deg, #f43f5e 0%, #e11d48 100%);
      box-shadow: 0 8px 18px -4px rgba(244, 63, 94, 0.38);
    }
    .btn-pulang:hover:not(:disabled) {
      background: linear-gradient(135deg, #e11d48 0%, #be123c 100%);
      transform: translateY(-2px);
      box-shadow: 0 12px 22px -4px rgba(244, 63, 94, 0.46);
    }
    .btn-action:active:not(:disabled) {
      transform: scale(0.97);
    }
    .btn-action:disabled {
      opacity: 0.55;
      cursor: not-allowed;
      transform: none;
      box-shadow: none;
    }

    /* Status Notification Box */
    .status-box {
      display: none;
      padding: 14px 16px;
      border-radius: var(--radius-element);
      font-size: 13px;
      line-height: 1.5;
      font-weight: 500;
      text-align: left;
      margin-bottom: 20px;
      animation: statusFade 0.3s ease-out forwards;
    }
    @keyframes statusFade {
      from { opacity: 0; transform: translateY(-6px); }
      to { opacity: 1; transform: translateY(0); }
    }
    .status-box.show {
      display: flex;
      align-items: flex-start;
      gap: 10px;
    }
    .status-icon {
      width: 20px;
      height: 20px;
      flex-shrink: 0;
      margin-top: 1px;
    }
    .status-box.info {
      background: #eff6ff;
      color: #1e40af;
      border: 1px solid #bfdbfe;
    }
    .status-box.success {
      background: #f0fdf4;
      color: #166534;
      border: 1px solid #bbf7d0;
    }
    .status-box.error {
      background: #fef2f2;
      color: #991b1b;
      border: 1px solid #fecaca;
    }

    /* Spinner Animation */
    .spinner {
      animation: rotate 1s linear infinite;
    }
    @keyframes rotate {
      100% { transform: rotate(360deg); }
    }

    /* Location Verification Badge */
    .geo-badge {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 6px 14px;
      background: #f8fafc;
      border: 1px solid #e2e8f0;
      border-radius: 999px;
      font-size: 11.5px;
      font-weight: 600;
      color: #475569;
    }
    .pulse-dot-wrapper {
      position: relative;
      width: 8px;
      height: 8px;
    }
    .pulse-dot {
      width: 8px;
      height: 8px;
      background-color: #10b981;
      border-radius: 50%;
    }
    .pulse-ring {
      position: absolute;
      top: 0;
      left: 0;
      width: 8px;
      height: 8px;
      background-color: #10b981;
      border-radius: 50%;
      opacity: 0.75;
      animation: pingRing 1.8s cubic-bezier(0, 0, 0.2, 1) infinite;
    }
    @keyframes pingRing {
      75%, 100% {
        transform: scale(2.5);
        opacity: 0;
      }
    }
  </style>
</head>
<body>
  <div class="ambient-glow-1"></div>
  <div class="ambient-glow-2"></div>

  <div class="card">
    <!-- Brand Header -->
    <header class="brand-header">
      <div class="logo-container">
        <img src="${logoBase64}" alt="JONAS PHOTO" class="logo-img">
      </div>
      <span class="badge-sub">PRESENSI DIGITAL</span>
    </header>

    <!-- Time & Date Widget -->
    <div class="time-widget">
      <div class="date-text" id="dateDisplay">--</div>
      <div class="time-row">
        <div class="clock-display" id="clockDisplay">--:--:--</div>
        <span class="tz-badge">WIB</span>
      </div>
    </div>

    <!-- PIN Input -->
    <div class="form-group">
      <div class="label-row">
        <label for="pinInput">PIN Absensi</label>
        <span class="label-hint">4–8 Digit Angka</span>
      </div>
      <div class="pin-input-wrapper">
        <input
          id="pinInput"
          type="password"
          inputmode="numeric"
          pattern="[0-9]*"
          maxlength="8"
          placeholder="••••"
          autocomplete="one-time-code"
        >
        <button type="button" class="btn-toggle-eye" id="btnToggleEye" onclick="togglePinVisibility()" aria-label="Lihat PIN">
          <!-- Eye Icon (visible by default) -->
          <svg id="eyeIconShow" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
          </svg>
          <!-- Eye Slash Icon (hidden by default) -->
          <svg id="eyeIconHide" style="display:none;" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858.908a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242M9.88 9.88l-3.29-3.29m7.532 7.532l3.29 3.29M3 3l18 18" />
          </svg>
        </button>
      </div>
    </div>

    <!-- Action Buttons -->
    <div class="btn-grid">
      <button type="button" class="btn-action btn-masuk" id="btnMasuk" onclick="lakukanAbsen('masuk')">
        <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M11 16l-4-4m0 0l4-4m-4 4h14m-5 4v1a3 3 0 01-3 3H6a3 3 0 01-3-3V7a3 3 0 013-3h7a3 3 0 013 3v1" />
        </svg>
        <span>Masuk</span>
      </button>
      <button type="button" class="btn-action btn-pulang" id="btnPulang" onclick="lakukanAbsen('pulang')">
        <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M17 16l4-4m0 0l-4-4m4 4H7m6 4v1a3 3 0 01-3 3H6a3 3 0 01-3-3V7a3 3 0 013-3h4a3 3 0 013 3v1" />
        </svg>
        <span>Pulang</span>
      </button>
    </div>

    <!-- Status Message Alert -->
    <div id="statusBox" class="status-box" role="alert">
      <div id="statusIcon" class="status-icon"></div>
      <div id="statusText"></div>
    </div>

    <!-- GPS Location Indicator -->
    <div class="geo-badge">
      <div class="pulse-dot-wrapper">
        <div class="pulse-ring"></div>
        <div class="pulse-dot"></div>
      </div>
      <span>Validasi Lokasi GPS Aktif (Radius 100m)</span>
    </div>
  </div>

  <script>
    // Realtime Clock & Indonesian Date
    function updateClock() {
      const now = new Date();
      const opsiJam = { timeZone: "Asia/Jakarta", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false };
      const jamStr = new Intl.DateTimeFormat("id-ID", opsiJam).format(now);
      document.getElementById("clockDisplay").textContent = jamStr;

      const opsiTanggal = { timeZone: "Asia/Jakarta", weekday: "long", year: "numeric", month: "long", day: "numeric" };
      const tglStr = new Intl.DateTimeFormat("id-ID", opsiTanggal).format(now);
      document.getElementById("dateDisplay").textContent = tglStr;
    }
    setInterval(updateClock, 1000);
    updateClock();

    // PIN Visibility Toggle
    function togglePinVisibility() {
      const input = document.getElementById("pinInput");
      const iconShow = document.getElementById("eyeIconShow");
      const iconHide = document.getElementById("eyeIconHide");
      if (input.type === "password") {
        input.type = "text";
        iconShow.style.display = "none";
        iconHide.style.display = "block";
      } else {
        input.type = "password";
        iconShow.style.display = "block";
        iconHide.style.display = "none";
      }
    }

    // Status Alert Helper
    const statusBox = document.getElementById("statusBox");
    const statusIcon = document.getElementById("statusIcon");
    const statusText = document.getElementById("statusText");
    const btnMasuk = document.getElementById("btnMasuk");
    const btnPulang = document.getElementById("btnPulang");

    function showStatus(pesan, tipe) {
      statusBox.className = "status-box show " + tipe;
      statusText.innerHTML = pesan;

      if (tipe === "info") {
        statusIcon.innerHTML = '<svg class="spinner" fill="none" stroke="currentColor" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke-width="3" stroke-dasharray="30 60" stroke-linecap="round" /></svg>';
      } else if (tipe === "success") {
        statusIcon.innerHTML = '<svg fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>';
      } else if (tipe === "error") {
        statusIcon.innerHTML = '<svg fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>';
      }
    }

    function setButtonsBusy(isBusy) {
      btnMasuk.disabled = isBusy;
      btnPulang.disabled = isBusy;
    }

    // Absen Submission with GPS Validation
    function lakukanAbsen(aksi) {
      const pinInput = document.getElementById("pinInput");
      const pin = pinInput.value.trim();

      if (!/^\\d{4,8}$/.test(pin)) {
        showStatus("<strong>PIN Tidak Valid:</strong> Masukkan 4 sampai 8 digit angka.", "error");
        pinInput.focus();
        return;
      }

      if (!navigator.geolocation) {
        showStatus("<strong>GPS Tidak Didukung:</strong> Browser Anda tidak mendukung deteksi lokasi.", "error");
        return;
      }

      setButtonsBusy(true);
      showStatus("Menghubungkan ke satelit GPS & memindai lokasi Anda...", "info");

      navigator.geolocation.getCurrentPosition(
        function (pos) {
          showStatus("Memvalidasi radius kerja & mencatat presensi...", "info");

          google.script.run
            .withSuccessHandler(function (pesan) {
              showStatus("<strong>Berhasil!</strong> " + pesan, "success");
              pinInput.value = "";
              setButtonsBusy(false);
            })
            .withFailureHandler(function (err) {
              showStatus("<strong>Gagal:</strong> " + (err.message || "Terjadi kendala pada sistem."), "error");
              setButtonsBusy(false);
            })
            .absen(pin, aksi, pos.coords.latitude, pos.coords.longitude);
        },
        function (err) {
          setButtonsBusy(false);
          let pesan = "Gagal mendeteksi lokasi GPS Anda.";
          if (err.code === 1) {
            pesan = "<strong>Izin Lokasi Ditolak:</strong> Silakan aktifkan izin lokasi di browser Anda untuk melanjutkan absensi.";
          } else if (err.code === 2) {
            pesan = "<strong>Sinyal GPS Lemah:</strong> Pastikan GPS pada perangkat Anda aktif dan memiliki sinyal yang memadai.";
          } else if (err.code === 3) {
            pesan = "<strong>Waktu Habis:</strong> Pencarian sinyal GPS membutuhkan waktu terlalu lama. Silakan coba lagi.";
          }
          showStatus(pesan, "error");
        },
        { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 }
      );
    }
  </script>
</body>
</html>
`;

fs.writeFileSync('apps-script/Index.html', newHtml, 'utf8');
console.log('Successfully updated apps-script/Index.html!');
