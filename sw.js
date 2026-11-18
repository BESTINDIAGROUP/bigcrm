/* Best India Group CRM - shows alerts (also when the CRM is closed) and opens the app when an alert is tapped */
importScripts('https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.12.2/firebase-messaging-compat.js');
try {
  firebase.initializeApp({
    apiKey: "AIzaSyCHVzY3_9mGKBYaIAfx5sCJ-Wo7YP5eKxY",
    authDomain: "bigcrm-8bc50.firebaseapp.com",
    projectId: "bigcrm-8bc50",
    storageBucket: "bigcrm-8bc50.firebasestorage.app",
    messagingSenderId: "866789460849",
    appId: "1:866789460849:web:003c674626d1e96c03c4d0"
  });
  firebase.messaging(); // shows the alerts that arrive while the CRM is closed
} catch (e) {}

self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('notificationclick', function (e) {
  e.notification.close();
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (list) {
    for (var i = 0; i < list.length; i++) { if ('focus' in list[i]) return list[i].focus(); }
    if (self.clients.openWindow) return self.clients.openWindow('./');
  }));
});
