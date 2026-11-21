/* Invoice Summariser - service worker
   On every release: change VERSION here AND APP_VERSION in index.html.
   A changed VERSION makes installed copies show the "Update" button. */
const VERSION = '1.1.0';
const CACHE = 'invsum-' + VERSION;
const RUNTIME = 'invsum-runtime';           // fonts + logo, kept across versions
const SHELL = ['./', './index.html', './manifest.json', './images/icon-192.png', './images/icon-512.png'];

self.addEventListener('install', (e) => {
  // fetch fresh copies (bypass the HTTP cache) so a new version never caches an old page
  e.waitUntil(caches.open(CACHE).then((c) => Promise.all(
    SHELL.map((u) => fetch(u, { cache: 'reload' }).then((r) => { if (r.ok) return c.put(u, r); }).catch(() => {}))
  )));
  // no skipWaiting here: the page asks the user first (Update button)
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k.startsWith('invsum-') && k !== CACHE && k !== RUNTIME).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('message', (e) => {
  if (e.data === 'SKIP_WAITING') self.skipWaiting();
  if (e.data === 'GET_VERSION' && e.source) e.source.postMessage({ type: 'VERSION', version: VERSION });
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // the app page: cached copy first (instant, works offline), network as fallback
  if (req.mode === 'navigate') {
    e.respondWith(caches.open(CACHE).then((c) =>
      c.match('./index.html').then((hit) => hit || c.match('./')).then((hit) => hit || fetch(req))
    ).catch(() => fetch(req)));
    return;
  }

  // own files (manifest, icons)
  if (url.origin === location.origin) {
    e.respondWith(caches.match(req, { ignoreSearch: true }).then((hit) => hit || fetch(req)));
    return;
  }

  // Google Fonts + Weave logo: serve cached, refresh in background
  if (/fonts\.(googleapis|gstatic)\.com$/.test(url.hostname) || url.hostname.endsWith('weavesolution.com')) {
    e.respondWith(caches.open(RUNTIME).then((c) => c.match(req).then((hit) => {
      const net = fetch(req).then((r) => { if (r.ok || r.type === 'opaque') c.put(req, r.clone()); return r; }).catch(() => hit);
      return hit || net;
    })));
  }
});
