// node test.js — cek hitung jarak radius
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const ctx = {};
vm.runInNewContext(fs.readFileSync(__dirname + '/apps-script/Code.gs', 'utf8') + ';this.jarakMeter=jarakMeter;this.LOKASI=LOKASI', ctx);
const { jarakMeter, LOKASI } = ctx;

assert.strictEqual(jarakMeter(LOKASI.lat, LOKASI.lng, LOKASI.lat, LOKASI.lng), 0);
const d = jarakMeter(LOKASI.lat, LOKASI.lng, LOKASI.lat + 0.0045, LOKASI.lng); // ~500 m ke utara
assert(d > 490 && d < 510, d);
assert(jarakMeter(LOKASI.lat, LOKASI.lng, -6.9147, 107.6098) > 500); // Alun-alun Bandung
console.log('ok');
