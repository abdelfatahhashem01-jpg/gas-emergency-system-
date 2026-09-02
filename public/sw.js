// Service Worker بسيط لتفعيل العمل بدون إنترنت (PWA):
// - يخزّن نسخة من ملفات الواجهة الثابتة (HTML/JS/CSS/الصور) بعد أول زيارة
// - عند انقطاع الإنترنت، يقدّم النسخة المخزَّنة بدل ما يفشل تحميل الصفحة تماماً
// - لا يتدخل إطلاقاً في طلبات Firebase/سيرفر الشركة الخارجية (بيانات فعلية حيّة) —
//   دي بيتعامل معاها التطبيق نفسه عبر النسخة الاحتياطية المحلية (localStorage)
//   الموجودة بالفعل في طبقة التخزين (loadJSON/saveJSON)
const CACHE_NAME = "gas-emergency-shell-v1";

self.addEventListener("install", () => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const { request } = event;
  if (request.method !== "GET") return;

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return; // اترك طلبات الخوادم الخارجية (Firebase وغيرها) بدون تدخل

  event.respondWith(
    caches.open(CACHE_NAME).then(async (cache) => {
      try {
        const fresh = await fetch(request);
        cache.put(request, fresh.clone());
        return fresh;
      } catch (e) {
        const cached = await cache.match(request);
        return cached || cache.match("/index.html") || Response.error();
      }
    })
  );
});
