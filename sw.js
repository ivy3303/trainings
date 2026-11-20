// Minimal service worker: it only makes the page installable. Nothing is cached, so every update shows up at once.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (event) => event.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => {});
