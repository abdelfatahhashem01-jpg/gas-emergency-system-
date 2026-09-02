import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

// تسجيل Service Worker لتفعيل العمل بدون إنترنت وإمكانية تثبيت البرنامج
// كتطبيق على الهاتف (PWA) — لا يمنع أي وظيفة لو فشل التسجيل لأي سبب
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("/sw.js").catch(() => {});
  });
}
