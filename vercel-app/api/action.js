import crypto from 'node:crypto';

// Env Vercel: GAS_URL (URL /exec Apps Script), API_SECRET (sama dengan Script Properties).
// APP_ORIGIN opsional (mis. https://e-presensi.vercel.app); default dari header Host.

const b64u = (s) => Buffer.from(String(s || ''), 'base64url');
const sha256 = (buf) => crypto.createHash('sha256').update(buf).digest();

function parseClientData(clientDataJSON, type, origin) {
  let cd;
  try { cd = JSON.parse(b64u(clientDataJSON).toString('utf8')); } catch { throw new Error('Data biometrik tidak valid.'); }
  if (cd.type !== type) throw new Error('Jenis data biometrik salah.');
  if (cd.origin !== origin) throw new Error('Asal halaman tidak sah.');
  if (!cd.challenge) throw new Error('Challenge tidak ada.');
  return cd;
}

function publicKeyFrom(spkiB64u) {
  const key = crypto.createPublicKey({ key: b64u(spkiB64u), format: 'der', type: 'spki' });
  const ok = key.asymmetricKeyType === 'rsa' || (key.asymmetricKeyType === 'ec' && key.asymmetricKeyDetails?.namedCurve === 'prime256v1');
  if (!ok) throw new Error('Jenis kunci biometrik tidak didukung.');
  return key;
}

/**
 * Verifikasi assertion WebAuthn (ES256 / RS256).
 * Return { challenge, signCount } jika sah; throw jika tidak.
 */
export function verifyAssertion({ clientDataJSON, authenticatorData, signature, publicKey, storedSignCount, origin }) {
  const cd = parseClientData(clientDataJSON, 'webauthn.get', origin);
  const auth = b64u(authenticatorData);
  if (auth.length < 37) throw new Error('Data autentikator tidak valid.');
  if (!auth.subarray(0, 32).equals(sha256(new URL(origin).hostname))) throw new Error('Domain biometrik tidak cocok.');
  const flags = auth[32];
  if (!(flags & 0x01) || !(flags & 0x04)) throw new Error('Face ID / sidik jari wajib dipakai.');

  const data = Buffer.concat([auth, sha256(b64u(clientDataJSON))]);
  if (!crypto.verify('sha256', data, publicKeyFrom(publicKey), b64u(signature))) {
    throw new Error('Tanda tangan biometrik tidak sah.');
  }

  const signCount = auth.readUInt32BE(33);
  if ((signCount || storedSignCount) && signCount <= storedSignCount) {
    throw new Error('Biometrik terdeteksi digandakan. Hubungi admin.');
  }
  return { challenge: cd.challenge, signCount };
}

async function gas(payload) {
  const r = await fetch(process.env.GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ...payload, secret: process.env.API_SECRET }),
    redirect: 'follow',
    signal: AbortSignal.timeout(25000)
  });
  const text = await r.text();
  let data;
  try { data = JSON.parse(text); } catch { throw new Error('Server absensi tidak merespons dengan benar.'); }
  if (!data.success) throw new Error(data.error || 'Terjadi kesalahan pada server.');
  return data;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ success: false, error: 'Method not allowed' });
  if (!process.env.GAS_URL || !process.env.API_SECRET) {
    return res.status(500).json({ success: false, error: 'GAS_URL / API_SECRET belum diatur di Vercel.' });
  }

  const b = req.body || {};
  const origin = process.env.APP_ORIGIN || `https://${req.headers.host}`;
  const lokasi = { aksi: b.aksi, lat: b.lat, lng: b.lng };

  try {
    let data;
    switch (b.action) {
      case 'challenge':
        data = await gas({ action: 'challenge' });
        break;

      case 'absen':
        data = await gas({ action: 'absen', idUser: b.idUser, pin: b.pin, ...lokasi });
        break;

      case 'register': {
        const cd = parseClientData(b.clientDataJSON, 'webauthn.create', origin);
        publicKeyFrom(b.publicKey);
        data = await gas({ action: 'register', idUser: b.idUser, pin: b.pin, credentialId: b.credentialId, publicKey: b.publicKey, challenge: cd.challenge });
        break;
      }

      case 'absenBio': {
        const cred = await gas({ action: 'credential', credentialId: b.credentialId });
        const v = verifyAssertion({ ...b, publicKey: cred.publicKey, storedSignCount: cred.signCount, origin });
        data = await gas({ action: 'absenBio', credentialId: b.credentialId, challenge: v.challenge, signCount: v.signCount, ...lokasi });
        break;
      }

      default:
        throw new Error('Aksi tidak dikenal.');
    }
    return res.status(200).json(data);
  } catch (err) {
    const timeout = err.name === 'TimeoutError';
    return res.status(timeout ? 504 : 400).json({ success: false, error: timeout ? 'Server absensi terlalu lama merespons. Coba lagi.' : err.message });
  }
}
