const CACHE_NAME = "blo-forms-tracker-v8";
const APP_SHELL = [
	"/",
	"/index.html",
	"/styles.css",
	"/script.js",
	"/manifest.webmanifest",
	"/sw.js"
	// Note: icons are omitted here to avoid failing install when missing locally
];

self.addEventListener("install", (event) => {
	// Cache app shell, but don't fail install if some resources are missing (common in dev)
	event.waitUntil(
		(async () => {
			try {
				const cache = await caches.open(CACHE_NAME);
				await Promise.allSettled(
					APP_SHELL.map(async (path) => {
						try {
							const req = new Request(path, { cache: "reload" });
							const resp = await fetch(req);
							if (resp && resp.ok) {
								await cache.put(req, resp.clone());
							}
						} catch {
							// ignore missing/failed resources
						}
					})
				);
				await self.skipWaiting();
			} catch (error) {
				console.error("Service worker installation encountered an error:", error);
			}
		})()
	);
});

self.addEventListener("activate", (event) => {
	event.waitUntil(
		caches
			.keys()
			.then((keys) =>
				Promise.all(
					keys.map((key) => {
						if (key !== CACHE_NAME) {
							return caches.delete(key);
						}
						return null;
					})
				)
			)
			.then(() => self.clients.claim())
	);
});

self.addEventListener("fetch", (event) => {
	const { request } = event;
	if (request.method !== "GET") return;

	const url = new URL(request.url);

	// Never intercept non-HTTP(S)
	if (url.protocol !== "http:" && url.protocol !== "https:") return;

	// For top-level navigations: network-first with cache fallback (prevents hanging)
	if (request.mode === "navigate") {
		event.respondWith(
			(async () => {
				try {
					const network = await fetch(request);
					// Optionally cache successful navigations for offline
					if (network && network.ok) {
						try {
							const cache = await caches.open(CACHE_NAME);
							cache.put("/", network.clone()).catch(() => {});
							cache.put("/index.html", network.clone()).catch(() => {});
						} catch {}
					}
					return network;
				} catch {
					const cached = (await caches.match("/index.html")) || (await caches.match("/"));
					return cached || Response.error();
				}
			})()
		);
		return;
	}

	// Same-origin assets: cache-first, then network; cache successful responses
	if (url.origin === self.location.origin) {
		event.respondWith(
			(async () => {
				const cached = await caches.match(request);
				if (cached) return cached;
				try {
					const network = await fetch(request);
					if (network && network.ok && network.type === "basic") {
						try {
							const cache = await caches.open(CACHE_NAME);
							cache.put(request, network.clone()).catch(() => {});
						} catch {}
					}
					return network;
				} catch {
					// If request was for a shell asset, try fallback to index
					if (APP_SHELL.includes(url.pathname)) {
						const fallback = (await caches.match("/index.html")) || (await caches.match("/"));
						if (fallback) return fallback;
					}
					return Response.error();
				}
			})()
		);
		return;
	}

	// Cross-origin resources: just pass through to network (don't cache to avoid CORS/opaque issues)
	event.respondWith(fetch(request));
});

