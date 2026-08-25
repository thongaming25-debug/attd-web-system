// Workforce Suite — service worker
//
// This is the ONE file that genuinely has to live outside App.jsx:
// browsers require a service worker to be a real script served from a
// stable same-origin URL, so it can't be inlined into the app bundle
// or injected at runtime the way the manifest/icon are (see
// usePwaSetup in App.jsx). Everything else it needs — including the
// offline fallback page — is written directly below, so this is the
// only extra file to deploy.
//
// Scope: caches the app SHELL only (the HTML page + JS/CSS bundle) so
// the app installs as a PWA and opens instantly on repeat visits. It
// deliberately does NOT cache or queue any Supabase request —
// attendance check-in/out, payroll, QR verification, etc. all still
// require a live network connection. This is a UI-loading
// optimization, not offline data entry.
//
// Also handles incoming Web Push messages (see usePushSubscription in
// App.jsx, which registers this service worker's push subscription)
// and routes a tap on the resulting notification back into the app.
//
// IMPORTANT: bump CACHE_VERSION on every deploy (or wire it to your
// build's commit hash / package version at build time) so returning
// visitors pick up the new build instead of a stale cached shell.
const CACHE_VERSION = "v4";
const CACHE_NAME = `workforce-shell-${CACHE_VERSION}`;

const OFFLINE_HTML = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Offline — Workforce Suite</title>
<style>
body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;
background:#0a0f1a;color:#fff;font-family:-apple-system,"Segoe UI",Roboto,sans-serif;
text-align:center;padding:32px}
.box{max-width:340px}
h1{font-size:18px;margin:16px 0 8px}
p{font-size:13px;color:#a9b4c7;line-height:1.5;margin:0 0 20px}
button{background:#1fa26b;color:#fff;border:none;border-radius:8px;padding:10px 20px;
font-size:13px;font-weight:600;cursor:pointer}
</style></head>
<body><div class="box">
<h1>No connection</h1>
<p>Workforce Suite needs the internet to check in/out, load payroll, and sync data. Please reconnect and try again.</p>
<button onclick="location.reload()">Retry</button>
</div></body></html>`;

const offlineResponse = () =>
  new Response(OFFLINE_HTML, {
    headers: { "Content-Type": "text/html; charset=utf-8" },
  });

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches
      .open(CACHE_NAME)
      .then((cache) => cache.add("/"))
      .then(() => self.skipWaiting()),
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(
          keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)),
        ),
      )
      .then(() => self.clients.claim()),
  );
});

self.addEventListener("fetch", (event) => {
  const { request } = event;
  if (request.method !== "GET") return;

  // Never intercept Supabase (or any cross-origin API) traffic — those
  // must always hit the network live.
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  // Page loads / hash-route navigations: network-first, so people
  // always get the latest build when online. Only fall back to the
  // cached shell (or the built-in offline page above) when the
  // network is unreachable.
  if (request.mode === "navigate") {
    event.respondWith(
      fetch(request)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put("/", copy));
          return res;
        })
        .catch(() => caches.match("/").then((res) => res || offlineResponse())),
    );
    return;
  }

  // Static assets (JS/CSS/images/fonts): serve from cache instantly if
  // we have it, and refresh the cache from the network in the
  // background either way.
  event.respondWith(
    caches.match(request).then((cached) => {
      const networkFetch = fetch(request)
        .then((res) => {
          if (res && res.status === 200 && res.type === "basic") {
            const copy = res.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
          }
          return res;
        })
        // On network failure, fall back to cache if we have it, and
        // otherwise to a plain error Response — never resolve to
        // undefined, since respondWith() requires an actual Response.
        .catch(
          () =>
            cached ||
            new Response(null, { status: 504, statusText: "Gateway Timeout" }),
        );
      return cached || networkFetch;
    }),
  );
});

/* ---------------------------------------------------------------
   Web Push
   The payload is whatever JSON the sender (the push_notify Edge
   Function) puts in the push message body:
     { title, body, page, portal, icon?, badge?, tag? }
   "page" is one of App.jsx's internal page ids (e.g. "leave", "ot",
   "attcorr") and "portal" is "admin" or "employee" — together they
   match the app's hash router (#/admin/leave, #/employee/ot, ...) so
   the click handler can deep-link there. Every field is optional
   except title — missing ones fall back to sane defaults so a
   malformed/older payload still shows something instead of failing
   silently.
----------------------------------------------------------------*/
self.addEventListener("push", (event) => {
  let data = {};
  try {
    data = event.data ? event.data.json() : {};
  } catch {
    data = { title: "Workforce Suite", body: event.data?.text() || "" };
  }

  const title = data.title || "Workforce Suite";
  const options = {
    body: data.body || "",
    icon: data.icon || "/icon-192.png",
    badge: data.badge || "/icon-192.png",
    tag: data.tag || undefined,
    // Re-notify (vibrate/sound again) even if a notification with the
    // same tag is already showing — e.g. two separate leave requests
    // shouldn't silently collapse into one alert.
    renotify: !!data.tag,
    data: { page: data.page || null, portal: data.portal || "admin" },
  };

  event.waitUntil(self.registration.showNotification(title, options));
});

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const targetPage = event.notification.data?.page;
  const targetPortal = event.notification.data?.portal || "admin";
  const targetUrl = `/#/${targetPortal}${targetPage ? `/${targetPage}` : ""}`;

  event.waitUntil(
    self.clients
      .matchAll({ type: "window", includeUncontrolled: true })
      .then((clientsList) => {
        // Reuse an already-open tab when there is one — focus it and
        // tell the page (via postMessage) which portal/page to switch
        // to, rather than doing a full navigation that would lose state.
        for (const client of clientsList) {
          if ("focus" in client) {
            client.focus();
            client.postMessage({
              type: "push-navigate",
              page: targetPage,
              portal: targetPortal,
            });
            return;
          }
        }
        if (self.clients.openWindow) {
          return self.clients.openWindow(targetUrl);
        }
      }),
  );
});
