const CACHE_NAME = "blo-forms-tracker-v6";
const APP_SHELL = [
	"/",
	"/index.html",
	"/styles.css",
	"/script.js",
	"/manifest.webmanifest",
	"/sw.js",
	"/icons/icon-192.png",
	"/icons/icon-512.png"
];

self.addEventListener("install", (event) => {
	event.waitUntil(
		caches
			.open(CACHE_NAME)
			.then((cache) => cache.addAll(APP_SHELL))
			.then(() => self.skipWaiting())
			.catch((error) => {
				console.error("Service worker installation failed:", error);
			})
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
	if (request.method !== "GET") {
		return;
	}

	event.respondWith(
		caches.match(request).then((cached) => {
			if (cached) {
				return cached;
			}
			return fetch(request)
				.then((response) => {
					if (!response || response.status !== 200 || response.type !== "basic") {
						return response;
					}
					const responseClone = response.clone();
					caches
						.open(CACHE_NAME)
						.then((cache) => {
							cache.put(request, responseClone);
						})
						.catch((error) => {
							console.error("Failed to cache response:", error);
						});
					return response;
				})
				.catch((error) => {
					if (request.mode === "navigate") {
						return caches.match("/index.html");
					}
					throw error;
				});
		})
	);
});

