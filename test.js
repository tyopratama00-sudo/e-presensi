// node test.js — cek radius lokasi & verifikasi tanda tangan biometrik
const fs = require('fs'), vm = require('vm'), assert = require('assert'), crypto = require('crypto');

const ctx = {};
vm.runInNewContext(fs.readFileSync(__dirname + '/apps-script/Code.gs', 'utf8') + ';this.jarakMeter=jarakMeter;this.LOKASI=LOKASI', ctx);
const { jarakMeter, LOKASI } = ctx;
assert.strictEqual(jarakMeter(LOKASI.lat, LOKASI.lng, LOKASI.lat, LOKASI.lng), 0);
const d = jarakMeter(LOKASI.lat, LOKASI.lng, LOKASI.lat + 0.0045, LOKASI.lng); // ~500 m ke utara
assert(d > 490 && d < 510, d);

import('./vercel-app/api/action.js').then(({ verifyAssertion }) => {
  const origin = 'https://e-presensi.vercel.app';
  const { publicKey, privateKey } = crypto.generateKeyPairSync('ec', { namedCurve: 'prime256v1' });
  const b64u = (b) => Buffer.from(b).toString('base64url');
  const sha = (b) => crypto.createHash('sha256').update(b).digest();

  const assertion = (count, o = origin, flags = 0x05) => {
    const cd = Buffer.from(JSON.stringify({ type: 'webauthn.get', challenge: 'abc', origin: o }));
    const auth = Buffer.alloc(37);
    sha(new URL(origin).hostname).copy(auth);
    auth[32] = flags;
    auth.writeUInt32BE(count, 33);
    const sig = crypto.sign('sha256', Buffer.concat([auth, sha(cd)]), privateKey);
    return { clientDataJSON: b64u(cd), authenticatorData: b64u(auth), signature: b64u(sig) };
  };
  const pk = b64u(publicKey.export({ format: 'der', type: 'spki' }));
  const v = (a, stored = 0) => verifyAssertion({ ...a, publicKey: pk, storedSignCount: stored, origin });

  assert.deepStrictEqual(v(assertion(5)), { challenge: 'abc', signCount: 5 });
  assert.deepStrictEqual(v(assertion(0)), { challenge: 'abc', signCount: 0 }); // passkey tanpa counter
  assert.throws(() => v(assertion(5), 5), /digandakan/);
  assert.throws(() => v(assertion(1, 'https://evil.com')), /Asal halaman/);
  assert.throws(() => v(assertion(1, origin, 0x01)), /wajib/);
  const palsu = assertion(1);
  palsu.signature = assertion(2).signature;
  assert.throws(() => v(palsu), /tidak sah/);
  console.log('ok');
});
