const ASSET_VERSION = "12";
const CACHE_NAME = `blo-forms-tracker-v${ASSET_VERSION}`;
const VERSION_QUERY = `?v=${ASSET_VERSION}`;
const APP_SHELL = [
	"/",
	"/index.html",
	`/styles.css${VERSION_QUERY}`,
	`/script.js${VERSION_QUERY}`,
	"/manifest.webmanifest"
	// Note: icons are omitted to avoid failing installs when running without assets locally
];
const APP_SHELL_PATHS = new Set(APP_SHELL.map((entry) => new URL(entry, self.location.origin).pathname));

async function putInCache(request, response) {
	if (!response || !response.ok) return;
	if (response.type !== "basic" && response.type !== "default") return;
	try {
		const cache = await caches.open(CACHE_NAME);
		await cache.put(request, response.clone());
	} catch (error) {
		// eslint-disable-next-line no-console
		console.warn("Service worker cache put failed:", error);
	}
}

async function networkFirst(request) {
	try {
		const response = await fetch(request);
		if (response && response.ok) {
			await putInCache(request, response);
		}
		return response;
	} catch {
		const fallback = await caches.match(request, { ignoreSearch: true });
		return fallback || Response.error();
	}
}

async function cacheFirst(request) {
	const cached = await caches.match(request);
	if (cached) return cached;
	try {
		const response = await fetch(request);
		if (response && response.ok) {
			await putInCache(request, response);
		}
		return response;
	} catch {
		const fallback = await caches.match(request, { ignoreSearch: true });
		if (fallback) return fallback;
		const url = new URL(request.url);
		if (APP_SHELL_PATHS.has(url.pathname)) {
			const landing = (await caches.match("/index.html")) || (await caches.match("/"));
			if (landing) return landing;
		}
		return Response.error();
	}
}

self.addEventListener("install", (event) => {
	event.waitUntil(
		(async () => {
			try {
				const cache = await caches.open(CACHE_NAME);
				await Promise.allSettled(
					APP_SHELL.map(async (path) => {
						try {
							const request = new Request(path, { cache: "reload" });
							const response = await fetch(request);
							if (response && response.ok) {
								await cache.put(request, response.clone());
							}
						} catch {
							// Ignore missing resources so install still succeeds (useful in dev)
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
		(async () => {
			const keys = await caches.keys();
			await Promise.all(
				keys.map((key) => {
					if (key !== CACHE_NAME) {
						return caches.delete(key);
					}
					return null;
				})
			);
			await self.clients.claim();
		})()
	);
});

self.addEventListener("fetch", (event) => {
	const { request } = event;
	if (request.method !== "GET") return;

	const url = new URL(request.url);
	if (url.protocol !== "http:" && url.protocol !== "https:") return;

	if (request.mode === "navigate") {
		event.respondWith(networkFirst(request));
		return;
	}

	const isScriptOrStyle =
		request.destination === "script" ||
		request.destination === "style" ||
		url.pathname.endsWith(".js") ||
		url.pathname.endsWith(".css");

	if (isScriptOrStyle) {
		event.respondWith(networkFirst(request));
		return;
	}

	if (url.origin === self.location.origin) {
		event.respondWith(cacheFirst(request));
		return;
	}

	event.respondWith(fetch(request));
});
