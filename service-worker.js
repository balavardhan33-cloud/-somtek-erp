/**
 * TEKFLOW / Somtek ERP — Service Worker
 * ============================================================
 * Upload this file to the EXACT SAME folder as your ERP HTML file
 * on your web host (must be served over https:// — or localhost
 * for local testing; service workers do not run from a file://
 * path or from Google Drive/Dropbox-style file links).
 *
 * What this gives you:
 *  - "Install app" / "Add to Home Screen" becomes available in the
 *    browser (Chrome/Edge show an install icon in the address bar;
 *    Android shows an "Install app" prompt; iOS Safari uses
 *    Share -> "Add to Home Screen", which already works via the
 *    manifest/meta tags in the HTML file regardless of this file).
 *  - The app shell (HTML/CSS/JS) loads instantly from cache on
 *    repeat visits and still opens if you're briefly offline.
 *
 * What this does NOT do:
 *  - It does not make your data available offline. Every login,
 *    save, and data load still talks to Firestore over the network
 *    as normal — this only caches the app itself, never your
 *    business data, so you're always seeing live information the
 *    moment you have a connection.
 * ============================================================
 */

const CACHE_NAME = 'tekflow-erp-v2';

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  // Never cache live data calls (Firestore / any Google API) - always
  // go straight to the network so the app never shows stale business data.
  if (req.url.includes('googleapis.com')) return;

  // NETWORK-FIRST for the app page itself: always try to fetch the
  // latest index.html when online so every deploy shows up on the very
  // next reload. Fall back to the cached copy only when offline.
  if (req.mode === 'navigate' || req.destination === 'document') {
    event.respondWith(
      fetch(req)
        .then((response) => {
          if (response && response.status === 200) {
            const copy = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
          }
          return response;
        })
        .catch(() => caches.match(req))
    );
    return;
  }

  // Stale-while-revalidate for static assets (fonts, the xlsx library,
  // icons, etc.) - instant load from cache, then quietly refreshes the
  // cache in the background for next time.
  event.respondWith(
    caches.match(req).then((cached) => {
      const network = fetch(req)
        .then((response) => {
          if (response && response.status === 200) {
            const copy = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
          }
          return response;
        })
        .catch(() => cached);
      return cached || network;
    })
  );
});
