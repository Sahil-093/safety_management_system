self.addEventListener("install", function (event) {
  event.waitUntil(self.skipWaiting());
});

self.addEventListener("activate", function (event) {
  event.waitUntil(self.clients.claim());
});

self.addEventListener("fetch", function (event) {
  if (event.request.method !== "GET") return;
  event.respondWith(fetch(event.request));
});

self.addEventListener("notificationclick", function (event) {
  event.notification.close();
  const username = event.notification.data && event.notification.data.username;
  event.waitUntil(self.clients.matchAll({ type: "window", includeUncontrolled: true }).then(function (list) {
    const client = list[0];
    if (client) {
      client.postMessage({ type: "open-chat", username: username || "" });
      return client.focus();
    }
    return self.clients.openWindow("./");
  }));
});
