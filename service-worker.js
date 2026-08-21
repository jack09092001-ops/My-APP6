'use strict';

// 更新任何 app 檔案後，記得把版本號 +1，否則已安裝的使用者會繼續吃到舊快取。
const CACHE_NAME = 'store-dashboard-v10';
const APP_SHELL = [
  './',
  './index.html',
  './style.css',
  './app.js',
  './config.js',
  './manifest.json',
  './icon.svg',
  './vendor/exceljs.min.js',
];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = event.request.url;
  // 門市資料（Google Drive API）與地圖圖磚一律直接連網，不進 Service Worker 快取，
  // 資料新鮮度由 app.js 的 localStorage 快取機制負責，這裡只顧 App 殼。
  if (url.includes('googleapis.com') || url.includes('google.com/maps') || url.includes('gstatic.com')) {
    return;
  }
  event.respondWith(caches.match(event.request).then((cached) => cached || fetch(event.request)));
});
