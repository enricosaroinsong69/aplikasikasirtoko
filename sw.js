/* Service worker aplikasi Gudang Toko Bangunan. VERSI berganti otomatis setiap kali aplikasi dirakit ulang. */
const VERSI = 'gudang-e5201cb02e';
const AWALAN = 'gudang-';
const INTI = ["./", "index.html", "manifest.webmanifest", "demo.js", "ikon-180.png", "ikon-192.png", "ikon-512.png", "ikon-maskable-512.png", "fonts/barlow-condensed-latin-500-normal.woff2", "fonts/barlow-condensed-latin-600-normal.woff2", "fonts/barlow-condensed-latin-700-normal.woff2", "fonts/barlow-semi-condensed-latin-400-normal.woff2", "fonts/barlow-semi-condensed-latin-500-normal.woff2", "fonts/barlow-semi-condensed-latin-600-normal.woff2", "fonts/barlow-semi-condensed-latin-700-normal.woff2"];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSI).then(c => c.addAll(INTI)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  // hanya hapus cache lama milik aplikasi ini; aplikasi lain di alamat yang sama punya cache sendiri
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k.startsWith(AWALAN) && k !== VERSI).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
function jaringanDulu(req) {
  return new Promise(resolve => {
    let selesai = false;
    const dariCache = () => caches.open(VERSI).then(c => c.match(req, { ignoreSearch: true }).then(r => r || c.match('./index.html')));
    const waktu = setTimeout(() => { if (!selesai) { selesai = true; dariCache().then(resolve); } }, 4000);
    fetch(req).then(res => {
      if (res && res.ok) { const salin = res.clone(); caches.open(VERSI).then(c => c.put(req, salin)); }
      if (!selesai) { selesai = true; clearTimeout(waktu); resolve(res); }
    }).catch(() => { if (!selesai) { selesai = true; clearTimeout(waktu); dariCache().then(resolve); } });
  });
}
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  const lingkup = new URL(self.registration.scope).pathname;
  if (url.origin !== location.origin || !url.pathname.startsWith(lingkup)) return;
  // berkas di subfolder lain (mis. aplikasi Gudang di folder gudang/) diurus service worker-nya sendiri
  const rel = url.pathname.slice(lingkup.length);
  if (rel.includes('/') && !rel.startsWith('fonts/')) return;
  if (req.mode === 'navigate' || /\.(html|webmanifest)$/.test(url.pathname) || url.pathname.endsWith('/')) { e.respondWith(jaringanDulu(req)); return; }
  e.respondWith(caches.open(VERSI).then(c => c.match(req).then(r => r || fetch(req).then(res => { if (res && res.ok) c.put(req, res.clone()); return res; }))));
});
