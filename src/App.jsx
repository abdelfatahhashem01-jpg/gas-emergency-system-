import React, { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid } from "recharts";

/* ============================== ثوابت ============================== */

const ROLES = {
  super_admin: { label: "سوبر أدمن", icon: "👑" },
  manager: { label: "مدير", icon: "🛡️" },
  response: { label: "فريق الاستجابة", icon: "🚒" },
  observer: { label: "مراقب", icon: "👁️" },
  reporter: { label: "مُبلِّغ", icon: "📢" },
};

const INCIDENT_TYPES = [
  "تسرب غاز", "رائحة غاز قوية", "غاز ضعيف", "خط غاز مكسور", "انفجار غاز",
  "حريق ناتج عن غاز", "اشتباه في تسرب", "عطل عداد مسبق الدفع", "عبث عميل",
  "تأمين", "تسريب شبكة أرضية", "ماس كهربائي", "حريق", "عطل عداد ميكانيكي",
  "إعادة تشغيل", "أخرى",
];

const SEVERITIES = [
  { id: "1", label: "أولى", color: "#D93025", bg: "#FDECEA", emoji: "🔴" },
  { id: "2", label: "ثانية", color: "#B76B00", bg: "#FFF3DE", emoji: "🟠" },
  { id: "3", label: "ثالثة", color: "#188038", bg: "#E6F4EA", emoji: "🟢" },
];

const STATUSES = [
  { id: "new", label: "جديد", color: "#1A73E8", bg: "#E8F0FE" },
  { id: "processing", label: "قيد المعالجة", color: "#B76B00", bg: "#FFF3DE" },
  { id: "done", label: "مكتمل", color: "#188038", bg: "#E6F4EA" },
];

const sevMeta = (id) => SEVERITIES.find((s) => s.id === id) || SEVERITIES[2];
const statusMeta = (id) => STATUSES.find((s) => s.id === id) || STATUSES[0];

function genCode(len = 6) {
  const chars = "0123456789";
  let s = "";
  for (let i = 0; i < len; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}
function fmtTime(ts) {
  try {
    return new Date(ts).toLocaleString("ar-EG", {
      day: "2-digit", month: "2-digit", year: "numeric",
      hour: "2-digit", minute: "2-digit",
    });
  } catch { return ""; }
}
function fmtDuration(ms) {
  if (ms == null || ms < 0) return "—";
  const mins = Math.round(ms / 60000);
  if (mins < 1) return "أقل من دقيقة";
  if (mins < 60) return `${mins} دقيقة`;
  const hrs = Math.floor(mins / 60);
  const rem = mins % 60;
  return rem ? `${hrs} س ${rem} د` : `${hrs} ساعة`;
}
// ترتيب أولوية تلقائي: درجة الخطورة (أولى فأولى) ثم حالة المعالجة (الأحدث احتياجاً أولاً) ثم الأحدث زمنياً
const SEVERITY_RANK = { "1": 0, "2": 1, "3": 2 };
const STATUS_RANK = { new: 0, processing: 1, done: 2 };
function sortByPriority(list) {
  return [...list].sort((a, b) => {
    const sv = (SEVERITY_RANK[a.severity] ?? 9) - (SEVERITY_RANK[b.severity] ?? 9);
    if (sv !== 0) return sv;
    const stt = (STATUS_RANK[a.status] ?? 9) - (STATUS_RANK[b.status] ?? 9);
    if (stt !== 0) return stt;
    return b.createdAt - a.createdAt;
  });
}

/* ============================== تخزين ============================== */

// طبقة تخزين مستقلة بالكامل: تعتمد على localStorage الخاص بالمتصفح فقط،
// بدون أي استدعاء لأدوات Claude الخارجية. هذا يجعل التطبيق يعمل بشكل مستقل
// تماماً عند رفعه على أي استضافة خاصة بك (Vercel أو غيرها) دون أي اعتماد خارجي.
// ملاحظة: داخل معاينة Claude.ai نفسها قد يمنع الوضع الآمن (sandbox) الوصول لـ
// localStorage، فيتحول التطبيق تلقائياً لتخزين مؤقت بالذاكرة (لجلسة العرض هذه
// فقط) حتى لا يتعطل — لكن بعد رفع الملف على سيرفرك الخاص سيعمل التخزين الدائم
// بشكل طبيعي وكامل.
const memoryStore = {};
const LOCAL_PREFIX = "gas_emergency_";

function hasLocalStorage() {
  try {
    const t = "__mg_test__";
    window.localStorage.setItem(t, "1");
    window.localStorage.removeItem(t);
    return true;
  } catch (e) {
    return false;
  }
}

async function loadJSON(key, fallback) {
  try {
    if (hasLocalStorage()) {
      const raw = window.localStorage.getItem(LOCAL_PREFIX + key);
      if (raw !== null) {
        const parsed = JSON.parse(raw);
        memoryStore[key] = parsed;
        return parsed;
      }
    }
  } catch (e) { /* تجاهل واستخدم الذاكرة المؤقتة */ }
  return memoryStore[key] !== undefined ? memoryStore[key] : fallback;
}

// يعيد { persisted, error }. لا يمنع التطبيق من الاستمرار حتى لو تعذّر التخزين
// الدائم — البيانات تبقى متاحة محلياً خلال الجلسة الحالية عبر memoryStore.
async function saveJSON(key, value) {
  memoryStore[key] = value;
  try {
    if (hasLocalStorage()) {
      window.localStorage.setItem(LOCAL_PREFIX + key, JSON.stringify(value));
      return { persisted: true, error: null };
    }
    return { persisted: false, error: "localStorage غير متاح في هذه البيئة (وضع المعاينة)" };
  } catch (e) {
    return { persisted: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/* ============================== أنماط عامة ============================== */

// روابط خرائط جوجل — إما بإحداثيات دقيقة (موقع GPS) أو ببحث نصي (عنوان/منطقة)
function mapsLinkForCoords(lat, lng) {
  return `https://www.google.com/maps?q=${lat},${lng}`;
}
function mapsLinkForQuery(q) {
  return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(q)}`;
}
// يطلب من المتصفح تحديد الموقع الحالي (GPS)، ويُرجع الإحداثيات إن وافق المستخدم
function captureLocation() {
  return new Promise((resolve, reject) => {
    if (!("geolocation" in navigator)) {
      return reject(new Error("المتصفح أو الجهاز لا يدعم تحديد الموقع"));
    }
    navigator.geolocation.getCurrentPosition(
      (pos) => resolve({
        lat: pos.coords.latitude, lng: pos.coords.longitude,
        accuracy: pos.coords.accuracy, ts: Date.now(),
      }),
      (err) => {
        const messages = {
          1: "تم رفض إذن الوصول للموقع",
          2: "تعذّر تحديد الموقع الحالي",
          3: "انتهت مهلة تحديد الموقع",
        };
        reject(new Error(messages[err.code] || "تعذّر تحديد الموقع"));
      },
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 60000 }
    );
  });
}

function GlobalStyle() {
  return (
    <style>{`
      * { box-sizing: border-box; }
      .mg-root {
        direction: rtl;
        font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
        background: #FBF7F6;
        color: #241614;
        min-height: 100vh;
      }
      .mg-btn {
        border: none; border-radius: 10px; padding: 10px 16px;
        font-size: 14px; font-weight: 600; cursor: pointer;
        display: inline-flex; align-items: center; gap: 6px;
        transition: transform .08s ease, opacity .15s ease;
      }
      .mg-btn:active { transform: scale(0.97); }
      .mg-btn:disabled { opacity: .5; cursor: not-allowed; }
      .mg-btn-primary { background: #D93025; color: #fff; }
      .mg-btn-primary:hover { background: #B0271B; }
      .mg-btn-outline { background: #fff; color: #D93025; border: 1.5px solid #D93025; }
      .mg-btn-ghost { background: #F3E9E8; color: #6B4B47; }
      .mg-btn-sm { padding: 6px 10px; font-size: 12.5px; border-radius: 8px; }
      .mg-input, .mg-select, .mg-textarea {
        width: 100%; border: 1.5px solid #EBD9D6; border-radius: 10px;
        padding: 10px 12px; font-size: 14.5px; background: #fff; color: #241614;
        font-family: inherit;
      }
      .mg-input:focus, .mg-select:focus, .mg-textarea:focus {
        outline: none; border-color: #D93025;
      }
      .mg-label { font-size: 13px; font-weight: 600; color: #6B4B47; margin-bottom: 6px; display: block; }
      .mg-card {
        background: #fff; border-radius: 16px; padding: 18px;
        box-shadow: 0 1px 3px rgba(46,20,17,.08), 0 1px 2px rgba(46,20,17,.06);
        border: 1px solid #F3E9E8;
      }
      .mg-header {
        background: linear-gradient(135deg, #D93025, #B0271B);
        color: #fff; padding: 14px 18px; position: sticky; top: 0; z-index: 20;
        box-shadow: 0 2px 8px rgba(0,0,0,.15);
      }
      .mg-tabs {
        display: flex; gap: 6px; overflow-x: auto; padding: 8px 14px;
        background: #fff; border-bottom: 1px solid #F3E9E8; position: sticky; top: 62px; z-index: 15;
      }
      .mg-tab {
        border: none; background: transparent; padding: 8px 14px; border-radius: 20px;
        font-size: 13.5px; font-weight: 600; color: #8A6C68; white-space: nowrap; cursor: pointer;
      }
      .mg-tab.active { background: #D93025; color: #fff; }
      .mg-badge {
        display: inline-flex; align-items: center; gap: 4px; padding: 3px 9px;
        border-radius: 20px; font-size: 12px; font-weight: 700;
      }
      .mg-modal-overlay {
        position: fixed; inset: 0; background: rgba(36,22,20,.55);
        display: flex; align-items: flex-end; justify-content: center; z-index: 100;
      }
      @media (min-width: 640px) {
        .mg-modal-overlay { align-items: center; }
      }
      .mg-modal {
        background: #fff; border-radius: 18px 18px 0 0; width: 100%; max-width: 480px;
        max-height: 88vh; overflow-y: auto; padding: 20px;
      }
      @media (min-width: 640px) {
        .mg-modal { border-radius: 18px; }
      }
      .mg-stat {
        background: #fff; border-radius: 14px; padding: 14px; text-align: center;
        border: 1px solid #F3E9E8;
      }
      .mg-stat .num { font-size: 26px; font-weight: 800; color: #D93025; }
      .mg-stat .lbl { font-size: 12px; color: #8A6C68; margin-top: 2px; }
      .mg-report-row {
        background: #fff; border-radius: 14px; padding: 14px; border: 1px solid #F3E9E8;
        margin-bottom: 10px; cursor: pointer;
      }
      .mg-report-row:active { background: #FFF8F7; }
      .mg-grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
      ::placeholder { color: #B79A96; }
      .mg-fab {
        position: fixed; bottom: 20px; left: 20px; z-index: 30;
        border-radius: 999px; padding: 14px 20px; box-shadow: 0 4px 14px rgba(217,48,37,.4);
      }
    `}</style>
  );
}

/* ============================== مكوّن نافذة عامة ============================== */

function Modal({ title, onClose, children }) {
  return (
    <div className="mg-modal-overlay" onClick={onClose}>
      <div className="mg-modal" onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <h3 style={{ margin: 0, fontSize: 17, color: "#241614" }}>{title}</h3>
          <button className="mg-btn mg-btn-ghost mg-btn-sm" onClick={onClose}>✕</button>
        </div>
        {children}
      </div>
    </div>
  );
}

/* ============================== شاشة الإعداد الأولي ============================== */

function SetupScreen({ onDone }) {
  const [name, setName] = useState("");
  const [code, setCode] = useState(genCode());
  const [confirm, setConfirm] = useState("");
  const [err, setErr] = useState("");
  const [saving, setSaving] = useState(false);
  const [created, setCreated] = useState(null);
  const [warn, setWarn] = useState("");

  async function handleCreate() {
    setErr(""); setWarn("");
    if (!name.trim()) return setErr("أدخل اسم المسؤول");
    if (code.trim().length < 4) return setErr("كود الدخول يجب ألا يقل عن 4 خانات");
    if (code !== confirm) return setErr("الكود وتأكيد الكود غير متطابقين");
    setSaving(true);
    const user = { id: uid(), name: name.trim(), role: "super_admin", code: code.trim(), region: "" };
    const r1 = await saveJSON("gas_users", [user]);
    await saveJSON("gas_regions", []);
    await saveJSON("gas_reports", []);
    await saveJSON("gas_alerts", [{ id: uid(), text: `تم إنشاء حساب السوبر أدمن: ${user.name}`, ts: Date.now() }]);
    setSaving(false);
    // لا نمنع الدخول حتى لو فشلت المزامنة الدائمة — الحساب يعمل محلياً في هذه الجلسة
    if (!r1.persisted) {
      setWarn(`تنبيه: لم يتم حفظ البيانات بشكل دائم (${r1.error || "سبب غير معروف"}). سيعمل الحساب في هذه الجلسة فقط ولن يتزامن مع أجهزة أخرى حتى تتم إعادة المحاولة.`);
    }
    setCreated(user);
  }

  if (created) {
    return (
      <div className="mg-root" style={{ display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
        <GlobalStyle />
        <div className="mg-card" style={{ maxWidth: 420, width: "100%", textAlign: "center" }}>
          <div style={{ fontSize: 40 }}>✅</div>
          <h2 style={{ color: "#188038", margin: "8px 0" }}>تم إنشاء حساب السوبر أدمن</h2>
          <p style={{ color: "#6B4B47", fontSize: 14 }}>احتفظ بهذه البيانات في مكان آمن — ستحتاجها لتسجيل الدخول</p>
          {warn && <div style={{ background: "#FFF3DE", color: "#B76B00", border: "1px solid #F2C67A", borderRadius: 10, padding: 10, fontSize: 12.5, marginBottom: 12, textAlign: "right" }}>⚠️ {warn}</div>}
          <div style={{ background: "#FFF8F7", border: "1.5px dashed #D93025", borderRadius: 12, padding: 16, margin: "16px 0", textAlign: "right" }}>
            <div style={{ fontSize: 13, color: "#8A6C68" }}>الاسم</div>
            <div style={{ fontSize: 17, fontWeight: 700, marginBottom: 10 }}>{created.name}</div>
            <div style={{ fontSize: 13, color: "#8A6C68" }}>كود الدخول</div>
            <div style={{ fontSize: 24, fontWeight: 800, letterSpacing: 3, color: "#D93025", fontFamily: "monospace" }}>{created.code}</div>
          </div>
          <button className="mg-btn mg-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={() => onDone(created)}>
            متابعة إلى النظام ←
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="mg-root" style={{ display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <GlobalStyle />
      <div className="mg-card" style={{ maxWidth: 420, width: "100%" }}>
        <div style={{ textAlign: "center", marginBottom: 18 }}>
          <div style={{
            width: 56, height: 56, borderRadius: 14, background: "linear-gradient(135deg,#D93025,#B0271B)",
            display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 800,
            fontSize: 18, margin: "0 auto 10px",
          }}>MG</div>
          <h2 style={{ margin: "4px 0 2px" }}>الشركة الحديثة للغاز الطبيعي</h2>
          <div style={{ color: "#8A6C68", fontSize: 13.5 }}>إعداد أول تشغيل — إنشاء حساب السوبر أدمن</div>
        </div>

        <label className="mg-label">اسم المسؤول *</label>
        <input className="mg-input" style={{ marginBottom: 12 }} value={name} onChange={(e) => setName(e.target.value)} placeholder="مثال: محمد عبد المقصود" />

        <label className="mg-label">كود الدخول *</label>
        <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
          <input className="mg-input" value={code} onChange={(e) => setCode(e.target.value)} placeholder="كود الدخول" />
          <button className="mg-btn mg-btn-outline mg-btn-sm" type="button" onClick={() => setCode(genCode())}>🎲 توليد</button>
        </div>

        <label className="mg-label">تأكيد كود الدخول *</label>
        <input className="mg-input" style={{ marginBottom: 14 }} value={confirm} onChange={(e) => setConfirm(e.target.value)} placeholder="أعد كتابة الكود" />

        {err && <div style={{ color: "#D93025", fontSize: 13, marginBottom: 10 }}>⚠️ {err}</div>}

        <button className="mg-btn mg-btn-primary" style={{ width: "100%", justifyContent: "center" }} disabled={saving} onClick={handleCreate}>
          {saving ? "جارٍ الإنشاء…" : "✅ إنشاء حساب السوبر أدمن"}
        </button>
      </div>
    </div>
  );
}

/* ============================== شاشة تسجيل الدخول ============================== */

function LoginScreen({ onLogin }) {
  const [code, setCode] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);

  async function handleLogin() {
    setErr("");
    if (!code.trim()) return setErr("أدخل كود الدخول");
    setBusy(true);
    const users = await loadJSON("gas_users", []);
    setBusy(false);
    const found = users.find((u) => u.code === code.trim());
    if (!found) return setErr("كود الدخول غير صحيح");
    onLogin(found);
  }

  return (
    <div className="mg-root" style={{ display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <GlobalStyle />
      <div className="mg-card" style={{ maxWidth: 380, width: "100%", textAlign: "center" }}>
        <div style={{
          width: 64, height: 64, borderRadius: 16, background: "linear-gradient(135deg,#D93025,#B0271B)",
          display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 800,
          fontSize: 20, margin: "0 auto 12px",
        }}>MG</div>
        <h2 style={{ margin: "2px 0" }}>نظام طوارئ الغاز</h2>
        <div style={{ color: "#8A6C68", fontSize: 13.5, marginBottom: 4 }}>الشركة الحديثة للغاز الطبيعي</div>
        <div style={{ color: "#B79A96", fontSize: 12, marginBottom: 20 }}>إدارة الطوارئ والعمليات</div>

        <div style={{ textAlign: "right" }}>
          <label className="mg-label">🔑 كود الدخول</label>
          <input
            className="mg-input" style={{ marginBottom: 12, textAlign: "center", letterSpacing: 3, fontSize: 18 }}
            value={code} onChange={(e) => setCode(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleLogin()}
            placeholder="••••••" inputMode="numeric"
          />
        </div>
        {err && <div style={{ color: "#D93025", fontSize: 13, marginBottom: 10 }}>⚠️ {err}</div>}
        <button className="mg-btn mg-btn-primary" style={{ width: "100%", justifyContent: "center" }} disabled={busy} onClick={handleLogin}>
          {busy ? "جارٍ التحقق…" : "دخول ←"}
        </button>
      </div>
    </div>
  );
}

/* ============================== نموذج بلاغ جديد ============================== */

function IncidentForm({ currentUser, regions, onSubmitted }) {
  const empty = {
    reporterName: "", phone: "", subscriberNumber: "", incidentType: "",
    severity: "2", details: "", region: regions[0]?.id || "",
    address: "", location: null,
  };
  const [form, setForm] = useState(empty);
  const [err, setErr] = useState("");
  const [ok, setOk] = useState(false);
  const [saving, setSaving] = useState(false);
  const [locStatus, setLocStatus] = useState("idle"); // idle | loading | done | error
  const [locErr, setLocErr] = useState("");
  const [lastCrn, setLastCrn] = useState("");

  function set(k, v) { setForm((f) => ({ ...f, [k]: v })); }

  async function attachLocation() {
    setLocStatus("loading"); setLocErr("");
    try {
      const loc = await captureLocation();
      set("location", loc);
      setLocStatus("done");
    } catch (e) {
      setLocStatus("error");
      setLocErr(e.message || "تعذّر تحديد الموقع");
    }
  }

  async function submit() {
    setErr(""); setOk(false);
    if (!form.reporterName.trim()) return setErr("اسم المُبلِّغ مطلوب");
    if (!/^\d{11}$/.test(form.phone.trim())) return setErr("رقم الهاتف يجب أن يكون 11 رقماً");
    if (!form.subscriberNumber.trim()) return setErr("رقم المشترك مطلوب");
    if (!form.incidentType) return setErr("اختر نوع الحادثة");
    setSaving(true);
    const reports = await loadJSON("gas_reports", []);
    // رقم مرجعي فريد للبلاغ (CRN) يسهّل البحث والمتابعة لاحقاً
    const crn = `MG-${new Date().getFullYear()}-${String(reports.length + 1).padStart(6, "0")}`;
    const report = {
      id: uid(), crn, ...form, status: "new", assignee: "", comments: "",
      createdBy: currentUser.name, createdAt: Date.now(),
      startedAt: null, completedAt: null,
    };
    const next = [report, ...reports];
    await saveJSON("gas_reports", next);
    const alerts = await loadJSON("gas_alerts", []);
    const locNote = form.location ? " 📍 مع الموقع الجغرافي" : "";
    const urgentPrefix = form.severity === "1" ? "🚨 عاجل — " : "";
    await saveJSON("gas_alerts", [{ id: uid(), text: `${urgentPrefix}بلاغ جديد (${sevMeta(form.severity).label}) رقم ${crn} من ${form.reporterName}${locNote}`, ts: Date.now() }, ...alerts].slice(0, 100));
    setSaving(false);
    setForm(empty);
    setLocStatus("idle"); setLocErr("");
    setLastCrn(crn);
    setOk(true);
    onSubmitted && onSubmitted(next);
    setTimeout(() => setOk(false), 5000);
  }

  return (
    <div className="mg-card">
      <h3 style={{ marginTop: 0 }}>📋 تسجيل بلاغ جديد</h3>

      <label className="mg-label">اسم المُبلِّغ *</label>
      <input className="mg-input" style={{ marginBottom: 12 }} value={form.reporterName} onChange={(e) => set("reporterName", e.target.value)} />

      <label className="mg-label">رقم الهاتف * (11 رقم)</label>
      <input className="mg-input" style={{ marginBottom: 4 }} value={form.phone} onChange={(e) => set("phone", e.target.value.replace(/\D/g, ""))} inputMode="numeric" maxLength={11} />
      {form.phone && form.phone.length !== 11 && (
        <div style={{ color: "#B76B00", fontSize: 12, marginBottom: 8 }}>⚠️ رقم الهاتف ناقص — يجب إدخال 11 رقماً كاملاً</div>
      )}
      <div style={{ marginBottom: 12 }} />

      <label className="mg-label">رقم المشترك *</label>
      <input className="mg-input" style={{ marginBottom: 12 }} value={form.subscriberNumber} onChange={(e) => set("subscriberNumber", e.target.value)} placeholder="**-**-**-**" />

      <label className="mg-label">العنوان البريدي</label>
      <input className="mg-input" style={{ marginBottom: 8 }} value={form.address} onChange={(e) => set("address", e.target.value)} placeholder="مثال: شارع..، حي..، المدينة" />
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
        <button type="button" className="mg-btn mg-btn-outline mg-btn-sm" disabled={locStatus === "loading"} onClick={attachLocation}>
          {locStatus === "loading" ? "⏳ جارٍ تحديد الموقع…" : "📍 إرفاق موقعي الحالي"}
        </button>
        {locStatus === "done" && form.location && (
          <span style={{ fontSize: 12.5, color: "#188038" }}>
            ✅ تم إرفاق الموقع —{" "}
            <a href={mapsLinkForCoords(form.location.lat, form.location.lng)} target="_blank" rel="noreferrer" style={{ color: "#188038" }}>عرض على الخريطة</a>
          </span>
        )}
        {locStatus === "error" && <span style={{ fontSize: 12.5, color: "#D93025" }}>⚠️ {locErr}</span>}
      </div>

      <label className="mg-label">المنطقة</label>
      <select className="mg-select" style={{ marginBottom: 12 }} value={form.region} onChange={(e) => set("region", e.target.value)}>
        <option value="">— بدون تحديد —</option>
        {regions.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
      </select>

      <label className="mg-label">نوع الحادثة *</label>
      <select className="mg-select" style={{ marginBottom: 12 }} value={form.incidentType} onChange={(e) => set("incidentType", e.target.value)}>
        <option value="">— اختر —</option>
        {INCIDENT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
      </select>

      <label className="mg-label">درجة الخطورة *</label>
      <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
        {SEVERITIES.map((s) => (
          <button
            key={s.id} type="button"
            onClick={() => set("severity", s.id)}
            className="mg-btn mg-btn-sm"
            style={{
              flex: 1, justifyContent: "center",
              background: form.severity === s.id ? s.color : s.bg,
              color: form.severity === s.id ? "#fff" : s.color,
              border: `1.5px solid ${s.color}`,
            }}
          >{s.emoji} {s.label}</button>
        ))}
      </div>

      <label className="mg-label">تفاصيل</label>
      <textarea className="mg-textarea" style={{ marginBottom: 14, minHeight: 80 }} value={form.details} onChange={(e) => set("details", e.target.value)} />

      {err && <div style={{ color: "#D93025", fontSize: 13, marginBottom: 10 }}>⚠️ {err}</div>}
      {ok && <div style={{ color: "#188038", fontSize: 13, marginBottom: 10 }}>✅ تم إرسال البلاغ بنجاح — رقم المرجع: <strong style={{ fontFamily: "monospace" }}>{lastCrn}</strong></div>}

      <button className="mg-btn mg-btn-primary" style={{ width: "100%", justifyContent: "center" }} disabled={saving} onClick={submit}>
        {saving ? "جارٍ الإرسال…" : "📤 إرسال البلاغ"}
      </button>
    </div>
  );
}

/* ============================== قائمة البلاغات ============================== */

// أفضل رابط خريطة متاح للبلاغ: إحداثيات GPS دقيقة، فالعنوان اليدوي + اسم المنطقة،
// فإحداثيات مركز المنطقة، فاسم المنطقة فقط كبحث نصي
function reportMapLink(r, regions) {
  if (r.location) return mapsLinkForCoords(r.location.lat, r.location.lng);
  const region = regions.find((rg) => rg.id === r.region);
  if (r.address && r.address.trim()) {
    const q = region ? `${r.address} ${region.name}` : r.address;
    return mapsLinkForQuery(q);
  }
  if (region && region.lat != null && region.lng != null) return mapsLinkForCoords(region.lat, region.lng);
  if (region) return mapsLinkForQuery(region.name);
  return null;
}

function ReportsList({ reports, regions, currentUser, canEdit, onChanged }) {
  const [filter, setFilter] = useState("all");
  const [q, setQ] = useState("");
  const [selected, setSelected] = useState(null);
  const [completing, setCompleting] = useState(null);
  const [exporting, setExporting] = useState(false);

  const visible = sortByPriority(
    reports.filter((r) => {
      if (currentUser.role === "reporter" && r.createdBy !== currentUser.name) return false;
      if (filter !== "all" && filter !== "sev1" && r.status !== filter) return false;
      if (filter === "sev1" && r.severity !== "1") return false;
      if (q.trim()) {
        const s = q.trim().toLowerCase();
        const haystack = [
          r.crn, r.reporterName, r.phone, r.subscriberNumber, r.address,
          r.incidentType, r.assignee, r.comments, r.details,
        ].filter(Boolean).join(" ").toLowerCase();
        if (!haystack.includes(s)) return false;
      }
      return true;
    })
  );

  async function logAlert(text) {
    const alerts = await loadJSON("gas_alerts", []);
    await saveJSON("gas_alerts", [{ id: uid(), text, ts: Date.now() }, ...alerts].slice(0, 100));
  }

  async function persist(next) {
    await saveJSON("gas_reports", next);
    onChanged(next);
  }

  // بدء المعالجة: يحوّل الحالة تلقائياً إلى "قيد المعالجة" ويسجّل وقت البدء
  // لحساب سرعة الاستجابة لاحقاً، ويربط البلاغ بالمستخدم الذي بدأ العمل عليه
  async function startProcessing(r) {
    const updated = { ...r, status: "processing", assignee: currentUser.name, startedAt: Date.now() };
    const next = reports.map((x) => (x.id === r.id ? updated : x));
    await persist(next);
    await logAlert(`⚙️ بدأ ${currentUser.name} معالجة بلاغ ${r.crn || r.reporterName}`);
  }

  // إكمال الحالة: يسجّل وقت الإنجاز ويحفظ ملخص التقرير النهائي كتعليق،
  // وتُستخدم الفروق الزمنية (استلام←بدء←إنجاز) في لوحة تقييم الأداء
  async function completeReport(r, note) {
    const updated = {
      ...r, status: "done", completedAt: Date.now(),
      comments: note ? (r.comments ? r.comments + "\n" + note : note) : r.comments,
    };
    const next = reports.map((x) => (x.id === r.id ? updated : x));
    await persist(next);
    await logAlert(`✅ أنجز ${currentUser.name} بلاغ ${r.crn || r.reporterName}`);
    setCompleting(null);
  }

  return (
    <div>
      <div className="mg-card" style={{ marginBottom: 12 }}>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <input className="mg-input" placeholder="🔍 بحث برقم البلاغ (CRN) أو الاسم أو الهاتف أو العنوان..." value={q} onChange={(e) => setQ(e.target.value)} />
          <button className="mg-btn mg-btn-outline mg-btn-sm" onClick={() => setExporting(true)}>📊 تصدير Excel</button>
        </div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {[
            { id: "all", label: "📋 الكل" }, { id: "new", label: "🆕 جديد" },
            { id: "processing", label: "⚙️ معالجة" }, { id: "done", label: "✅ مكتمل" },
            { id: "sev1", label: "🔴 أولى" },
          ].map((f) => (
            <button
              key={f.id} className="mg-tab" style={{ background: filter === f.id ? "#D93025" : "#F3E9E8", color: filter === f.id ? "#fff" : "#6B4B47" }}
              onClick={() => setFilter(f.id)}
            >{f.label}</button>
          ))}
        </div>
        <div style={{ fontSize: 11.5, color: "#B79A96", marginTop: 8 }}>الترتيب تلقائي حسب درجة الخطورة ثم حالة المعالجة الأكثر إلحاحاً</div>
      </div>

      {visible.length === 0 && (
        <div className="mg-card" style={{ textAlign: "center", color: "#8A6C68" }}>لا توجد بلاغات مطابقة</div>
      )}

      {visible.map((r) => {
        const sv = sevMeta(r.severity), st = statusMeta(r.status);
        const urgent = r.severity === "1" && r.status !== "done";
        return (
          <div
            key={r.id} className="mg-report-row" onClick={() => setSelected(r)}
            style={urgent ? { borderRight: "4px solid #D93025", background: "#FFFBFA" } : undefined}
          >
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div>
                <div style={{ fontWeight: 700 }}>
                  {urgent && <span style={{ marginLeft: 4 }}>🚨</span>}
                  {r.reporterName}
                </div>
                <div style={{ fontSize: 11.5, color: "#B79A96", fontFamily: "monospace" }}>{r.crn}</div>
                <div style={{ fontSize: 12.5, color: "#8A6C68" }}>{r.incidentType} · {r.phone}</div>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                <span className="mg-badge" style={{ background: sv.bg, color: sv.color }}>{sv.emoji} {sv.label}</span>
                <span className="mg-badge" style={{ background: st.bg, color: st.color }}>{st.label}</span>
              </div>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8 }}>
              <div style={{ fontSize: 11.5, color: "#B79A96" }}>
                {fmtTime(r.createdAt)}
                {r.assignee && <span> · 👤 {r.assignee}</span>}
                {r.startedAt && <span> · ⏱️ استجابة: {fmtDuration(r.startedAt - r.createdAt)}</span>}
                {r.completedAt && <span> · إنجاز: {fmtDuration(r.completedAt - r.startedAt)}</span>}
                {reportMapLink(r, regions) && (
                  <>
                    {" · "}
                    <a href={reportMapLink(r, regions)} target="_blank" rel="noreferrer" onClick={(e) => e.stopPropagation()} style={{ color: "#1A73E8" }}>
                      🗺️ الموقع
                    </a>
                  </>
                )}
              </div>
              {canEdit && r.status === "new" && (
                <button
                  className="mg-btn mg-btn-sm" style={{ background: "#1A73E8", color: "#fff" }}
                  onClick={(e) => { e.stopPropagation(); startProcessing(r); }}
                >🚀 بدء المعالجة</button>
              )}
              {canEdit && r.status === "processing" && (
                <button
                  className="mg-btn mg-btn-sm" style={{ background: "#188038", color: "#fff" }}
                  onClick={(e) => { e.stopPropagation(); setCompleting(r); }}
                >✅ مكتمل</button>
              )}
            </div>
          </div>
        );
      })}

      {selected && (
        <ReportDetailModal
          report={selected} regions={regions} canEdit={canEdit}
          onClose={() => setSelected(null)}
          onSave={async (updated) => {
            const next = reports.map((r) => (r.id === updated.id ? updated : r));
            await persist(next);
            setSelected(null);
          }}
        />
      )}

      {completing && (
        <Modal title="✅ إكمال الحالة وإرسال التقرير" onClose={() => setCompleting(null)}>
          <CompleteReportForm report={completing} onCancel={() => setCompleting(null)} onConfirm={(note) => completeReport(completing, note)} />
        </Modal>
      )}

      {exporting && (
        <ExportExcelModal reports={reports} regions={regions} onClose={() => setExporting(false)} />
      )}
    </div>
  );
}

// نافذة تصدير Excel — تسمح باختيار الفترة (يوم محدد / شهر محدد / سنة محددة / الكل)
// قبل توليد ملف .xlsx فعلي بكل بيانات البلاغات المطابقة للفترة
function ExportExcelModal({ reports, regions, onClose }) {
  const now = new Date();
  const [periodType, setPeriodType] = useState("all"); // all | day | month | year
  const [dayVal, setDayVal] = useState(now.toISOString().slice(0, 10));
  const [monthVal, setMonthVal] = useState(now.toISOString().slice(0, 7));
  const [yearVal, setYearVal] = useState(String(now.getFullYear()));

  function inPeriod(r) {
    if (periodType === "all") return true;
    const d = new Date(r.createdAt);
    if (periodType === "day") return d.toISOString().slice(0, 10) === dayVal;
    if (periodType === "month") return d.toISOString().slice(0, 7) === monthVal;
    if (periodType === "year") return String(d.getFullYear()) === yearVal;
    return true;
  }

  function doExport() {
    const rows = reports.filter(inPeriod).map((r) => ({
      "رقم البلاغ": r.crn || "",
      "الاسم": r.reporterName,
      "الهاتف": r.phone,
      "رقم المشترك": r.subscriberNumber,
      "العنوان": r.address || "",
      "المنطقة": regions.find((rg) => rg.id === r.region)?.name || "",
      "نوع الحادثة": r.incidentType,
      "الخطورة": sevMeta(r.severity).label,
      "الحالة": statusMeta(r.status).label,
      "المسؤول": r.assignee || "",
      "تاريخ البلاغ": fmtTime(r.createdAt),
      "وقت الاستجابة": r.startedAt ? fmtDuration(r.startedAt - r.createdAt) : "",
      "وقت الإنجاز": r.completedAt && r.startedAt ? fmtDuration(r.completedAt - r.startedAt) : "",
      "رابط الموقع": reportMapLink(r, regions) || "",
      "التفاصيل": r.details || "",
      "التعليقات/التقرير النهائي": r.comments || "",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "البلاغات");
    const label = periodType === "day" ? dayVal : periodType === "month" ? monthVal : periodType === "year" ? yearVal : "الكل";
    XLSX.writeFile(wb, `تقرير-البلاغات-${label}.xlsx`);
    onClose();
  }

  return (
    <Modal title="📊 تصدير تقرير Excel" onClose={onClose}>
      <label className="mg-label">الفترة</label>
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 14 }}>
        {[
          { id: "all", label: "الكل" }, { id: "day", label: "يوم محدد" },
          { id: "month", label: "شهر محدد" }, { id: "year", label: "سنة محددة" },
        ].map((p) => (
          <button
            key={p.id} type="button" className="mg-btn mg-btn-sm"
            style={{ background: periodType === p.id ? "#D93025" : "#F3E9E8", color: periodType === p.id ? "#fff" : "#6B4B47" }}
            onClick={() => setPeriodType(p.id)}
          >{p.label}</button>
        ))}
      </div>

      {periodType === "day" && (
        <>
          <label className="mg-label">اختر اليوم</label>
          <input type="date" className="mg-input" style={{ marginBottom: 14 }} value={dayVal} onChange={(e) => setDayVal(e.target.value)} />
        </>
      )}
      {periodType === "month" && (
        <>
          <label className="mg-label">اختر الشهر</label>
          <input type="month" className="mg-input" style={{ marginBottom: 14 }} value={monthVal} onChange={(e) => setMonthVal(e.target.value)} />
        </>
      )}
      {periodType === "year" && (
        <>
          <label className="mg-label">اختر السنة</label>
          <input type="number" className="mg-input" style={{ marginBottom: 14 }} value={yearVal} onChange={(e) => setYearVal(e.target.value)} />
        </>
      )}

      <button className="mg-btn mg-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={doExport}>📥 تنزيل ملف Excel</button>
    </Modal>
  );
}

function CompleteReportForm({ report, onCancel, onConfirm }) {
  const [note, setNote] = useState("");
  return (
    <div>
      <div style={{ fontSize: 13.5, color: "#6B4B47", marginBottom: 10 }}>
        بلاغ <strong>{report.reporterName}</strong> ({report.crn}) — اكتب ملخص التقرير النهائي قبل تحويل الحالة إلى مكتمل.
      </div>
      <label className="mg-label">ملخص التقرير النهائي</label>
      <textarea className="mg-textarea" style={{ marginBottom: 14, minHeight: 90 }} value={note} onChange={(e) => setNote(e.target.value)} placeholder="ما الذي تم تنفيذه لإغلاق البلاغ؟" />
      <div style={{ display: "flex", gap: 8 }}>
        <button className="mg-btn mg-btn-ghost" style={{ flex: 1, justifyContent: "center" }} onClick={onCancel}>إلغاء</button>
        <button className="mg-btn mg-btn-primary" style={{ flex: 1, justifyContent: "center" }} onClick={() => onConfirm(note)}>✅ إرسال التقرير وإكمال الحالة</button>
      </div>
    </div>
  );
}

function ReportDetailModal({ report, regions, canEdit, onClose, onSave }) {
  const [form, setForm] = useState({ ...report });
  function set(k, v) { setForm((f) => ({ ...f, [k]: v })); }
  return (
    <Modal title={`✏️ تفاصيل البلاغ — ${report.crn || ""}`} onClose={onClose}>
      <div className="mg-grid2" style={{ marginBottom: 10 }}>
        <div>
          <label className="mg-label">اسم المُبلِّغ</label>
          <input className="mg-input" disabled={!canEdit} value={form.reporterName} onChange={(e) => set("reporterName", e.target.value)} />
        </div>
        <div>
          <label className="mg-label">رقم الهاتف</label>
          <input className="mg-input" disabled={!canEdit} value={form.phone} onChange={(e) => set("phone", e.target.value)} />
        </div>
      </div>
      <label className="mg-label">رقم المشترك</label>
      <input className="mg-input" style={{ marginBottom: 10 }} disabled={!canEdit} value={form.subscriberNumber} onChange={(e) => set("subscriberNumber", e.target.value)} />

      <label className="mg-label">نوع الحادثة</label>
      <select className="mg-select" style={{ marginBottom: 10 }} disabled={!canEdit} value={form.incidentType} onChange={(e) => set("incidentType", e.target.value)}>
        {INCIDENT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
      </select>

      <div className="mg-grid2" style={{ marginBottom: 10 }}>
        <div>
          <label className="mg-label">الخطورة</label>
          <select className="mg-select" disabled={!canEdit} value={form.severity} onChange={(e) => set("severity", e.target.value)}>
            {SEVERITIES.map((s) => <option key={s.id} value={s.id}>{s.emoji} {s.label}</option>)}
          </select>
        </div>
        <div>
          <label className="mg-label">الحالة</label>
          <select className="mg-select" disabled={!canEdit} value={form.status} onChange={(e) => set("status", e.target.value)}>
            {STATUSES.map((s) => <option key={s.id} value={s.id}>{s.label}</option>)}
          </select>
        </div>
      </div>

      <label className="mg-label">المنطقة</label>
      <select className="mg-select" style={{ marginBottom: 10 }} disabled={!canEdit} value={form.region || ""} onChange={(e) => set("region", e.target.value)}>
        <option value="">— بدون تحديد —</option>
        {regions.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
      </select>

      <label className="mg-label">العنوان البريدي</label>
      <input className="mg-input" style={{ marginBottom: 8 }} disabled={!canEdit} value={form.address || ""} onChange={(e) => set("address", e.target.value)} />
      {reportMapLink(form, regions) && (
        <div style={{ marginBottom: 10 }}>
          <a href={reportMapLink(form, regions)} target="_blank" rel="noreferrer" style={{ color: "#1A73E8", fontSize: 13 }}>
            🗺️ {form.location ? "فتح الموقع الجغرافي على الخريطة" : "فتح العنوان/المنطقة على الخريطة"}
          </a>
          {form.location && <span style={{ fontSize: 11.5, color: "#8A6C68" }}> (بدقة ~{Math.round(form.location.accuracy || 0)} متر)</span>}
        </div>
      )}

      <label className="mg-label">تفاصيل</label>
      <textarea className="mg-textarea" style={{ marginBottom: 10, minHeight: 70 }} disabled={!canEdit} value={form.details} onChange={(e) => set("details", e.target.value)} />

      <label className="mg-label">التعليقات</label>
      <textarea className="mg-textarea" style={{ marginBottom: 14, minHeight: 60 }} disabled={!canEdit} value={form.comments} onChange={(e) => set("comments", e.target.value)} />

      <div style={{ fontSize: 12, color: "#B79A96", marginBottom: 10 }}>
        سجّله: {form.createdBy} · {fmtTime(form.createdAt)}
        {form.assignee && <><br />المسؤول: {form.assignee}</>}
        {form.startedAt && <><br />⏱️ وقت الاستجابة: {fmtDuration(form.startedAt - form.createdAt)}</>}
        {form.completedAt && <><br />✅ وقت الإنجاز: {fmtDuration(form.completedAt - form.startedAt)}</>}
      </div>

      {canEdit ? (
        <button className="mg-btn mg-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={() => onSave(form)}>💾 حفظ التعديلات</button>
      ) : (
        <div style={{ fontSize: 12.5, color: "#8A6C68", textAlign: "center" }}>عرض فقط — ليست لديك صلاحية التعديل</div>
      )}
    </Modal>
  );
}

/* ============================== لوحة الإحصائيات ============================== */

// توزيع عدد البلاغات على مدار ساعات اليوم (00–23) — يساعد في تحديد أوقات الذروة
function computeHourlyDistribution(reports) {
  const buckets = Array.from({ length: 24 }, (_, h) => ({ hour: `${String(h).padStart(2, "0")}:00`, عدد: 0 }));
  reports.forEach((r) => {
    const h = new Date(r.createdAt).getHours();
    buckets[h].عدد += 1;
  });
  return buckets;
}

function HourlyStatsChart({ reports }) {
  const data = computeHourlyDistribution(reports);
  if (reports.length === 0) {
    return <div style={{ color: "#8A6C68", fontSize: 14 }}>لا توجد بيانات كافية بعد</div>;
  }
  return (
    <div style={{ width: "100%", height: 220 }}>
      <ResponsiveContainer>
        <BarChart data={data} margin={{ top: 6, right: 6, left: -20, bottom: 0 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#F3E9E8" />
          <XAxis dataKey="hour" tick={{ fontSize: 10 }} interval={2} />
          <YAxis allowDecimals={false} tick={{ fontSize: 10 }} />
          <Tooltip />
          <Bar dataKey="عدد" fill="#D93025" radius={[4, 4, 0, 0]} />
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

function StatsPanel({ reports }) {
  const total = reports.length;
  const byStatus = (id) => reports.filter((r) => r.status === id).length;
  const bySev = (id) => reports.filter((r) => r.severity === id).length;
  const done = byStatus("done");
  const pct = total ? Math.round((done / total) * 100) : 0;

  const cards = [
    { num: total, lbl: "إجمالي" },
    { num: bySev("1"), lbl: "درجة أولى" },
    { num: byStatus("processing"), lbl: "قيد المعالجة" },
    { num: done, lbl: "مكتمل" },
    { num: byStatus("new"), lbl: "جديد" },
    { num: bySev("2"), lbl: "درجة ثانية" },
    { num: bySev("3"), lbl: "درجة ثالثة" },
    { num: `${pct}%`, lbl: "نسبة الإنجاز" },
  ];

  return (
    <div className="mg-card">
      <h3 style={{ marginTop: 0 }}>📊 الإحصائيات</h3>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10 }}>
        {cards.map((c, i) => (
          <div key={i} className="mg-stat">
            <div className="num">{c.num}</div>
            <div className="lbl">{c.lbl}</div>
          </div>
        ))}
      </div>
      <h3 style={{ marginTop: 18, marginBottom: 6 }}>📈 عدد البلاغات حسب ساعة اليوم</h3>
      <HourlyStatsChart reports={reports} />
    </div>
  );
}

/* ============================== أداء المستخدمين (سرعة الاستجابة والإنجاز) ============================== */

// يُحسب لكل مستخدم أُسندت له بلاغات (بمجرد ضغطه "بدء المعالجة"):
// عدد البلاغات، نسبة الإنجاز، متوسط سرعة الاستجابة (من وقت البلاغ لبدء المعالجة)
// ومتوسط وقت الإنجاز (من بدء المعالجة للإكمال) — ومنها تقييم أداء مبسّط.
// عتبات التقييم (15 دقيقة / ساعة) قابلة للتعديل حسب سياسة الشركة.
function computeUserPerformance(reports) {
  const map = {};
  reports.forEach((r) => {
    if (!r.assignee) return;
    if (!map[r.assignee]) {
      map[r.assignee] = { name: r.assignee, assigned: 0, completed: 0, responseSum: 0, responseCount: 0, resolutionSum: 0, resolutionCount: 0 };
    }
    const m = map[r.assignee];
    m.assigned += 1;
    if (r.startedAt) { m.responseSum += (r.startedAt - r.createdAt); m.responseCount += 1; }
    if (r.status === "done") m.completed += 1;
    if (r.completedAt && r.startedAt) { m.resolutionSum += (r.completedAt - r.startedAt); m.resolutionCount += 1; }
  });
  return Object.values(map)
    .map((m) => {
      const avgResponse = m.responseCount ? m.responseSum / m.responseCount : null;
      const avgResolution = m.resolutionCount ? m.resolutionSum / m.resolutionCount : null;
      const completionRate = m.assigned ? Math.round((m.completed / m.assigned) * 100) : 0;
      let rating = "لا توجد بيانات كافية", ratingColor = "#8A6C68";
      if (avgResponse != null) {
        if (avgResponse <= 15 * 60000) { rating = "⭐ ممتاز"; ratingColor = "#188038"; }
        else if (avgResponse <= 60 * 60000) { rating = "👍 جيد"; ratingColor = "#B76B00"; }
        else { rating = "⚠️ يحتاج تحسين"; ratingColor = "#D93025"; }
      }
      return { ...m, avgResponse, avgResolution, completionRate, rating, ratingColor };
    })
    .sort((a, b) => b.assigned - a.assigned);
}

function UserPerformancePanel({ reports }) {
  const rows = computeUserPerformance(reports);
  return (
    <div className="mg-card" style={{ marginTop: 12 }}>
      <h3 style={{ marginTop: 0 }}>👥 ملخص أداء المستخدمين</h3>
      {rows.length === 0 && <div style={{ color: "#8A6C68" }}>لا توجد بلاغات مُسندة بعد — الأداء يُحسب بمجرد أن يضغط أحد المستخدمين "بدء المعالجة"</div>}
      {rows.map((u) => (
        <div key={u.name} style={{ padding: "12px 0", borderBottom: "1px solid #F3E9E8" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
            <div style={{ fontWeight: 700 }}>{u.name}</div>
            <span className="mg-badge" style={{ background: "#F3E9E8", color: u.ratingColor }}>{u.rating}</span>
          </div>
          <div style={{ fontSize: 12.5, color: "#6B4B47", display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4 }}>
            <div>البلاغات المُسندة: {u.assigned}</div>
            <div>المكتملة: {u.completed} ({u.completionRate}%)</div>
            <div>متوسط سرعة الاستجابة: {u.avgResponse != null ? fmtDuration(u.avgResponse) : "—"}</div>
            <div>متوسط وقت الإنجاز: {u.avgResolution != null ? fmtDuration(u.avgResolution) : "—"}</div>
          </div>
        </div>
      ))}
    </div>
  );
}

/* ============================== إدارة المناطق ============================== */

function RegionsManager({ regions, onChanged }) {
  const [open, setOpen] = useState(false);
  const [name, setName] = useState("");
  const [code, setCode] = useState("");
  const [lat, setLat] = useState("");
  const [lng, setLng] = useState("");
  const [err, setErr] = useState("");

  async function create() {
    setErr("");
    if (!name.trim() || !code.trim()) return setErr("أدخل اسم المنطقة والكود");
    if ((lat && !lng) || (!lat && lng)) return setErr("أدخل خط الطول والعرض معاً أو اتركهما فارغين");
    const region = {
      id: uid(), name: name.trim(), code: code.trim().toUpperCase(),
      lat: lat ? parseFloat(lat) : null, lng: lng ? parseFloat(lng) : null,
    };
    const next = [...regions, region];
    await saveJSON("gas_regions", next);
    onChanged(next);
    setName(""); setCode(""); setLat(""); setLng(""); setOpen(false);
  }

  return (
    <div className="mg-card">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <h3 style={{ margin: 0 }}>🗺️ إدارة المناطق</h3>
        <button className="mg-btn mg-btn-primary mg-btn-sm" onClick={() => setOpen(true)}>➕ منطقة جديدة</button>
      </div>
      {regions.length === 0 && <div style={{ color: "#8A6C68", fontSize: 14 }}>لا توجد مناطق مضافة بعد</div>}
      {regions.map((r) => (
        <div key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: "1px solid #F3E9E8" }}>
          <span>{r.name}</span>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <a
              href={r.lat != null && r.lng != null ? mapsLinkForCoords(r.lat, r.lng) : mapsLinkForQuery(r.name)}
              target="_blank" rel="noreferrer" style={{ fontSize: 12.5, color: "#1A73E8" }}
            >🗺️ خريطة المنطقة</a>
            <span style={{ color: "#8A6C68", fontFamily: "monospace" }}>{r.code}</span>
          </div>
        </div>
      ))}
      {open && (
        <Modal title="🗺️ إضافة منطقة جديدة" onClose={() => setOpen(false)}>
          <label className="mg-label">اسم المنطقة *</label>
          <input className="mg-input" style={{ marginBottom: 12 }} value={name} onChange={(e) => setName(e.target.value)} />
          <label className="mg-label">كود المنطقة (بالإنجليزية) *</label>
          <input className="mg-input" style={{ marginBottom: 12 }} value={code} onChange={(e) => setCode(e.target.value)} />
          <label className="mg-label">إحداثيات مركز المنطقة (اختياري)</label>
          <div className="mg-grid2" style={{ marginBottom: 8 }}>
            <input className="mg-input" placeholder="خط العرض Lat" value={lat} onChange={(e) => setLat(e.target.value)} inputMode="decimal" />
            <input className="mg-input" placeholder="خط الطول Lng" value={lng} onChange={(e) => setLng(e.target.value)} inputMode="decimal" />
          </div>
          <div style={{ fontSize: 11.5, color: "#8A6C68", marginBottom: 12 }}>لو تُركت فارغة، سيتم البحث عن المنطقة بالاسم مباشرة على الخريطة</div>
          {err && <div style={{ color: "#D93025", fontSize: 13, marginBottom: 10 }}>⚠️ {err}</div>}
          <button className="mg-btn mg-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={create}>✅ إنشاء المنطقة</button>
        </Modal>
      )}
    </div>
  );
}

/* ============================== إدارة المستخدمين ============================== */

function UsersManager({ users, regions, onChanged }) {
  const [open, setOpen] = useState(false);
  const [name, setName] = useState("");
  const [role, setRole] = useState("response");
  const [code, setCode] = useState(genCode());
  const [region, setRegion] = useState("");
  const [err, setErr] = useState("");

  async function create() {
    setErr("");
    if (!name.trim()) return setErr("أدخل اسم المستخدم");
    if (code.trim().length < 4) return setErr("كود الدخول يجب ألا يقل عن 4 خانات");
    if (users.some((u) => u.code === code.trim())) return setErr("هذا الكود مستخدم بالفعل، جرّب كوداً آخر");
    const next = [...users, { id: uid(), name: name.trim(), role, code: code.trim(), region }];
    await saveJSON("gas_users", next);
    onChanged(next);
    setName(""); setCode(genCode()); setRole("response"); setRegion(""); setOpen(false);
  }

  return (
    <div className="mg-card">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <h3 style={{ margin: 0 }}>إدارة المستخدمين</h3>
        <button className="mg-btn mg-btn-primary mg-btn-sm" onClick={() => setOpen(true)}>➕ إضافة</button>
      </div>
      {users.map((u) => (
        <div key={u.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: "1px solid #F3E9E8" }}>
          <div>
            <div style={{ fontWeight: 600 }}>{ROLES[u.role]?.icon} {u.name}</div>
            <div style={{ fontSize: 12, color: "#8A6C68" }}>{ROLES[u.role]?.label}{u.region ? " · " + (regions.find((r) => r.id === u.region)?.name || "") : ""}</div>
          </div>
          <span style={{ fontFamily: "monospace", color: "#D93025", fontWeight: 700 }}>{u.code}</span>
        </div>
      ))}
      {open && (
        <Modal title="➕ إضافة مستخدم" onClose={() => setOpen(false)}>
          <label className="mg-label">الاسم *</label>
          <input className="mg-input" style={{ marginBottom: 12 }} value={name} onChange={(e) => setName(e.target.value)} />

          <label className="mg-label">الدور *</label>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 12 }}>
            {Object.entries(ROLES).map(([id, r]) => (
              <button key={id} type="button" className="mg-btn mg-btn-sm"
                style={{ background: role === id ? "#D93025" : "#F3E9E8", color: role === id ? "#fff" : "#6B4B47" }}
                onClick={() => setRole(id)}>
                {r.icon} {r.label}
              </button>
            ))}
          </div>

          <label className="mg-label">كود الدخول — احفظه وأعطه للمستخدم</label>
          <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
            <input className="mg-input" value={code} onChange={(e) => setCode(e.target.value)} />
            <button className="mg-btn mg-btn-outline mg-btn-sm" type="button" onClick={() => setCode(genCode())}>🎲</button>
          </div>

          <label className="mg-label">المنطقة</label>
          <select className="mg-select" style={{ marginBottom: 12 }} value={region} onChange={(e) => setRegion(e.target.value)}>
            <option value="">— بدون تحديد —</option>
            {regions.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
          </select>

          {err && <div style={{ color: "#D93025", fontSize: 13, marginBottom: 10 }}>⚠️ {err}</div>}
          <button className="mg-btn mg-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={create}>✅ إنشاء المستخدم</button>
        </Modal>
      )}
    </div>
  );
}

/* ============================== تغيير كود الدخول ============================== */

function ChangeCodeModal({ currentUser, users, onClose, onChanged }) {
  const [cur, setCur] = useState("");
  const [next, setNext] = useState("");
  const [confirm, setConfirm] = useState("");
  const [err, setErr] = useState("");
  const [ok, setOk] = useState(false);

  async function submit() {
    setErr("");
    if (cur !== currentUser.code) return setErr("الكود الحالي غير صحيح");
    if (next.trim().length < 4) return setErr("الكود الجديد يجب ألا يقل عن 4 أحرف أو أرقام");
    if (next !== confirm) return setErr("الكود الجديد وتأكيده غير متطابقين");
    const updatedUsers = users.map((u) => (u.id === currentUser.id ? { ...u, code: next.trim() } : u));
    await saveJSON("gas_users", updatedUsers);
    onChanged(updatedUsers, { ...currentUser, code: next.trim() });
    setOk(true);
    setTimeout(onClose, 1200);
  }

  return (
    <Modal title="🔑 تغيير كود الدخول" onClose={onClose}>
      <label className="mg-label">الكود الحالي *</label>
      <input className="mg-input" style={{ marginBottom: 12 }} value={cur} onChange={(e) => setCur(e.target.value)} />
      <label className="mg-label">الكود الجديد *</label>
      <input className="mg-input" style={{ marginBottom: 12 }} value={next} onChange={(e) => setNext(e.target.value)} />
      <label className="mg-label">تأكيد الكود الجديد *</label>
      <input className="mg-input" style={{ marginBottom: 8 }} value={confirm} onChange={(e) => setConfirm(e.target.value)} />
      <div style={{ fontSize: 12, color: "#B76B00", marginBottom: 12 }}>⚠️ احفظه جيداً — لن تتمكن من الدخول بدونه.</div>
      {err && <div style={{ color: "#D93025", fontSize: 13, marginBottom: 10 }}>⚠️ {err}</div>}
      {ok && <div style={{ color: "#188038", fontSize: 13, marginBottom: 10 }}>✅ تم تغيير الكود</div>}
      <button className="mg-btn mg-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={submit}>✅ تغيير الكود</button>
    </Modal>
  );
}

/* ============================== سجل التنبيهات ============================== */

function AlertsPanel({ alerts, onClear }) {
  return (
    <div className="mg-card">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <h3 style={{ margin: 0 }}>🔔 سجل التنبيهات</h3>
        <button className="mg-btn mg-btn-ghost mg-btn-sm" onClick={onClear}>🗑 مسح الكل</button>
      </div>
      {alerts.length === 0 && <div style={{ color: "#8A6C68" }}>لا توجد تنبيهات</div>}
      {alerts.map((a) => (
        <div key={a.id} style={{ padding: "10px 0", borderBottom: "1px solid #F3E9E8" }}>
          <div style={{ fontSize: 14 }}>{a.text}</div>
          <div style={{ fontSize: 11.5, color: "#B79A96" }}>{fmtTime(a.ts)}</div>
        </div>
      ))}
    </div>
  );
}

/* ============================== الهيدر ============================== */

function Header({ currentUser, onLogout, onChangeCode, onRefresh, tab, setTab }) {
  const tabsByRole = {
    super_admin: ["report", "list", "stats", "regions", "users", "alerts"],
    manager: ["report", "list", "stats", "regions", "alerts"],
    response: ["report", "list", "stats", "alerts"],
    observer: ["list", "stats", "alerts"],
    reporter: ["report", "list", "alerts"],
  };
  const tabLabels = {
    report: "📋 بلاغ جديد", list: "📄 البلاغات", stats: "📊 إحصائيات",
    regions: "🗺️ المناطق", users: "👥 المستخدمين", alerts: "🔔 التنبيهات",
  };
  const tabs = tabsByRole[currentUser.role] || ["list"];

  return (
    <>
      <div className="mg-header">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontWeight: 800, fontSize: 16 }}>MG · نظام طوارئ الغاز</div>
            <div style={{ fontSize: 12, opacity: 0.9 }}>{ROLES[currentUser.role]?.icon} {currentUser.name} · {ROLES[currentUser.role]?.label}</div>
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            <button className="mg-btn mg-btn-sm" style={{ background: "rgba(255,255,255,.2)", color: "#fff" }} onClick={onRefresh}>🔄</button>
            <button className="mg-btn mg-btn-sm" style={{ background: "rgba(255,255,255,.2)", color: "#fff" }} onClick={onChangeCode}>🔑</button>
            <button className="mg-btn mg-btn-sm" style={{ background: "rgba(255,255,255,.2)", color: "#fff" }} onClick={onLogout}>⏻</button>
          </div>
        </div>
      </div>
      <div className="mg-tabs">
        {tabs.map((t) => (
          <button key={t} className={`mg-tab ${tab === t ? "active" : ""}`} onClick={() => setTab(t)}>{tabLabels[t]}</button>
        ))}
      </div>
    </>
  );
}

/* ============================== التطبيق الرئيسي ============================== */

export default function App() {
  const [phase, setPhase] = useState("loading"); // loading | setup | login | app
  const [currentUser, setCurrentUser] = useState(null);
  const [users, setUsers] = useState([]);
  const [regions, setRegions] = useState([]);
  const [reports, setReports] = useState([]);
  const [alerts, setAlerts] = useState([]);
  const [tab, setTab] = useState("report");
  const [showChangeCode, setShowChangeCode] = useState(false);
  const pollRef = useRef(null);

  const loadAll = useCallback(async () => {
    const [u, r, rep, al] = await Promise.all([
      loadJSON("gas_users", []), loadJSON("gas_regions", []),
      loadJSON("gas_reports", []), loadJSON("gas_alerts", []),
    ]);
    setUsers(u); setRegions(r); setReports(rep); setAlerts(al);
    return u;
  }, []);

  useEffect(() => {
    (async () => {
      const u = await loadAll();
      setPhase(u.length === 0 ? "setup" : "login");
    })();
  }, [loadAll]);

  useEffect(() => {
    if (phase === "app") {
      pollRef.current = setInterval(() => { loadAll(); }, 20000);
      return () => clearInterval(pollRef.current);
    }
  }, [phase, loadAll]);

  if (phase === "loading") {
    return (
      <div className="mg-root" style={{ display: "flex", alignItems: "center", justifyContent: "center", minHeight: "100vh" }}>
        <GlobalStyle />
        <div style={{ color: "#8A6C68" }}>جاري الاتصال…</div>
      </div>
    );
  }

  if (phase === "setup") {
    return <SetupScreen onDone={(u) => { setCurrentUser(u); loadAll(); setPhase("app"); setTab("report"); }} />;
  }

  if (phase === "login") {
    return <LoginScreen onLogin={(u) => { setCurrentUser(u); loadAll(); setPhase("app"); setTab(u.role === "observer" ? "list" : "report"); }} />;
  }

  const canEdit = ["super_admin", "manager", "response"].includes(currentUser.role);

  return (
    <div className="mg-root">
      <GlobalStyle />
      <Header
        currentUser={currentUser}
        tab={tab} setTab={setTab}
        onRefresh={loadAll}
        onChangeCode={() => setShowChangeCode(true)}
        onLogout={() => { setCurrentUser(null); setPhase("login"); }}
      />
      <div style={{ padding: 14, maxWidth: 720, margin: "0 auto" }}>
        {tab === "report" && <IncidentForm currentUser={currentUser} regions={regions} onSubmitted={(rep) => { setReports(rep); loadAll(); }} />}
        {tab === "list" && (
          <ReportsList
            reports={reports} regions={regions} currentUser={currentUser}
            canEdit={canEdit}
            onChanged={(rep) => setReports(rep)}
          />
        )}
        {tab === "stats" && (
          <>
            <StatsPanel reports={reports} />
            <UserPerformancePanel reports={reports} />
          </>
        )}
        {tab === "regions" && <RegionsManager regions={regions} onChanged={setRegions} />}
        {tab === "users" && currentUser.role === "super_admin" && <UsersManager users={users} regions={regions} onChanged={setUsers} />}
        {tab === "alerts" && (
          <AlertsPanel
            alerts={alerts}
            onClear={async () => { await saveJSON("gas_alerts", []); setAlerts([]); }}
          />
        )}
      </div>

      {showChangeCode && (
        <ChangeCodeModal
          currentUser={currentUser} users={users}
          onClose={() => setShowChangeCode(false)}
          onChanged={(u, updatedMe) => { setUsers(u); setCurrentUser(updatedMe); }}
        />
      )}
    </div>
  );
}
