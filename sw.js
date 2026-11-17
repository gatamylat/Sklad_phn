const scope = self.registration.scope;
const prefix = `sklad-phone:${new URL(scope).pathname}:`;
const cacheName = `${prefix}v1`;
const assets = ['./', './index.html', './app.js', './stock.js', './manifest.webmanifest',
  './assets/xlsx.min.js', './assets/icon-192.png', './assets/icon-512.png'].map(path => new URL(path, scope).href);

self.addEventListener('install', event => {
  event.waitUntil(caches.open(cacheName).then(cache => cache.addAll(assets)));
});

self.addEventListener('activate', event => {
  event.waitUntil(Promise.all([self.clients.claim(), caches.keys().then(keys => Promise.all(keys
    .filter(name => name.startsWith(prefix) && name !== cacheName).map(name => caches.delete(name))))]));
});

self.addEventListener('fetch', event => {
  const {request} = event;
  if (request.method !== 'GET' || !request.url.startsWith(scope)) return;
  if (request.mode === 'navigate' && [scope, `${scope}index.html`].includes(request.url.split('?')[0])) {
    event.respondWith(fetch(request).catch(() => caches.open(cacheName).then(cache => cache.match(scope))));
  } else if (assets.includes(request.url)) {
    event.respondWith(caches.open(cacheName).then(async cache => (await cache.match(request)) ?? fetch(request)));
  }
});
