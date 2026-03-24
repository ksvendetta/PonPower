const CACHE_NAME = 'fiber-gen-v3';
const BASE = self.registration.scope;
const STATIC_ASSETS = [
  BASE,
  BASE + 'index.html',
  BASE + 'manifest.json',
  BASE + 'favicon.png',
  BASE + 'icon-192.png',
  BASE + 'icon-512.png'
];

// Install event: Cache core static assets immediately
self.addEventListener('install', (event) => {
  self.skipWaiting(); // Force waiting service worker to become active
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('Opened cache');
      return cache.addAll(STATIC_ASSETS);
    })
  );
});

// Activate event: Clean up old caches
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            console.log('Deleting old cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => self.clients.claim()) // Take control of all clients immediately
  );
});

// Fetch event: Network-first for HTML, Cache-first for assets
// Actually, for a robust offline experience with hashed assets we don't know about yet,
// we'll use Stale-While-Revalidate or Cache-First with dynamic caching.
// Given this is a SPA, we want to cache everything that loads.

self.addEventListener('fetch', (event) => {
  // Skip non-http requests (like chrome-extension://)
  if (!event.request.url.startsWith('http')) return;

  // Strategy: Cache First, falling back to Network, then caching the network response
  // This ensures assets load fast.
  // For navigation requests (HTML), we might want Network First to get updates, 
  // but for now let's stick to a simple dynamic cache.
  
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        // Return cached response immediately
        return cachedResponse;
      }

      return fetch(event.request).then((response) => {
        // Check if we received a valid response
        if (!response || response.status !== 200) {
          return response;
        }

        // Clone the response because it's a stream and can only be consumed once
        const responseToCache = response.clone();

        caches.open(CACHE_NAME).then((cache) => {
          // Cache the fetched response for future use
          cache.put(event.request, responseToCache);
        });

        return response;
      }).catch(() => {
        // If fetch fails (offline), and not in cache
        // We could return a fallback offline page here if we had one
        // For now, we hope the main index.html is cached
        if (event.request.mode === 'navigate') {
            return caches.match(BASE + 'index.html');
        }
      });
    })
  );
});
