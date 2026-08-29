import React, { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid } from "recharts";
import { initializeApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc } from "firebase/firestore";

/* ============================== Firebase (قاعدة بيانات الشركة) ============================== */
// بيانات مشروع Firebase الخاص بالشركة — كل المستخدمين على أي جهاز/متصفح
// بيقروا ويكتبوا في نفس القاعدة دي، فمفيش داعي بعد كده لـ localStorage كمصدر
// أساسي للبيانات (بيفضل موجود بس كنسخة احتياطية عند انقطاع الإنترنت).
const firebaseConfig = {
  apiKey: "AIzaSyCBuOOxedbybLg9t1YnjnQsB4WMsWWvfHs",
  authDomain: "moderngas-emergency.firebaseapp.com",
  projectId: "moderngas-emergency",
  storageBucket: "moderngas-emergency.firebasestorage.app",
  messagingSenderId: "1081793028931",
  appId: "1:1081793028931:web:394e93f8fb2fe074ba45f6",
  measurementId: "G-B5H4439YMB",
};
const fbApp = initializeApp(firebaseConfig);
const db = getFirestore(fbApp);
const FS_COLLECTION = "gas_emergency_data";

// دعم التحوّل لسيرفر الشركة الخاص بدل Firebase — بمجرد ضبط هذين المتغيرين
// عند بناء المشروع (Environment Variables)، يتحوّل التطبيق تلقائياً لاستخدام
// سيرفر الشركة بدل Firebase من غير أي تعديل إضافي في الكود. راجع
// gas-emergency-backend/README.md للتفاصيل الكاملة الموجّهة لفريق الـ IT.
const SELF_API_BASE = (typeof import.meta !== "undefined" && import.meta.env && import.meta.env.VITE_API_BASE_URL) || "";
const SELF_API_KEY = (typeof import.meta !== "undefined" && import.meta.env && import.meta.env.VITE_API_KEY) || "";

async function selfApiLoad(key) {
  const res = await fetch(`${SELF_API_BASE}/api/data/${encodeURIComponent(key)}`, {
    headers: SELF_API_KEY ? { "x-api-key": SELF_API_KEY } : {},
  });
  if (res.status === 404) return undefined;
  if (!res.ok) throw new Error("self-host API error " + res.status);
  const json = await res.json();
  return json.value;
}
async function selfApiSave(key, value) {
  const res = await fetch(`${SELF_API_BASE}/api/data/${encodeURIComponent(key)}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json", ...(SELF_API_KEY ? { "x-api-key": SELF_API_KEY } : {}) },
    body: JSON.stringify({ value }),
  });
  if (!res.ok) throw new Error("self-host API error " + res.status);
}

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
  "إعادة تشغيل", "انهيار عقار", "انفجار خط غاز 7 بار", "حريق منظم", "حادثة محطة غاز", "أخرى",
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

const sevMeta = (id) => {
  const base = SEVERITIES.find((s) => s.id === id) || SEVERITIES[2];
  const override = dynamicSettings.severityIcons[id];
  return override ? { ...base, emoji: override } : base;
};
const statusMeta = (id) => STATUSES.find((s) => s.id === id) || STATUSES[0];

// إعدادات قابلة للتخصيص من صفحة "⚙️ الإعدادات" (اسم البرنامج، أيقونات الأدوار
// ودرجات الخطورة) — تُخزَّن في localStorage وتُقرأ من هنا في كل مكان بالتطبيق
// بدل تمريرها كخاصية (prop) عبر كل مكوّن على حدة.
// جهات تنبيهات الطوارئ عبر واتساب — قائمة قابلة للتعديل بالكامل (إضافة/حذف/
// تعديل)، وكل جهة لها: مسمّى (وظيفة أو اسم)، رقم واتساب، ووصف اختياري.
// "الرقم المختصر" لكل جهة هو موضعها في القائمة (١، ٢، ٣...) ويُحسب تلقائياً
// من ترتيب القائمة، فيُستخدم في نص رسالة واتساب بدل كتابة المسمى بالكامل.
const NOTIFY_ROLES = [
  "رئيس الشركة",
  "مساعد رئيس الشركة",
  "مدير عام مراقبة التشغيل",
  "مدير المنطقة",
  "مهندس العمليات",
];
function defaultWhatsappNumbers() {
  return NOTIFY_ROLES.map((role) => ({ role, phone: "", description: "" }));
}
function defaultAppSettings() {
  return { appName: "نظام طوارئ الغاز", roleIcons: {}, severityIcons: {}, whatsappNumbers: defaultWhatsappNumbers(), zoomLink: "" };
}
let dynamicSettings = defaultAppSettings();
// يهاجر أي شكل قديم (بحقل "name" بدل "role"، أو بدون "description") للشكل
// الحالي، مع الحفاظ على كل البيانات الموجودة فعلاً
function migrateWhatsappNumbers(list) {
  if (!Array.isArray(list) || list.length === 0) return defaultWhatsappNumbers();
  return list.map((n) => ({
    role: (n && (n.role || n.name)) || "",
    phone: (n && n.phone) || "",
    description: (n && n.description) || "",
  }));
}
function applyAppSettings(s) {
  dynamicSettings = {
    appName: (s && s.appName) || "نظام طوارئ الغاز",
    roleIcons: (s && s.roleIcons) || {},
    severityIcons: (s && s.severityIcons) || {},
    whatsappNumbers: migrateWhatsappNumbers(s && s.whatsappNumbers),
    zoomLink: (s && s.zoomLink) || "",
  };
}
function roleIcon(role) { return dynamicSettings.roleIcons[role] || ROLES[role]?.icon || ""; }
function currentAppName() { return dynamicSettings.appName; }

// يبني نص رسالة التنبيه لحادثة شديدة الخطورة — يتضمن رابط موقع الحادثة على
// الخريطة (لو متاح) ورابط اجتماع الطوارئ الثابت على Zoom (لو مضبوط في الإعدادات)
// وسطر "تم إبلاغ" بأكواد الوظائف المسجَّل لها رقم هاتف (بدل كتابة كل مسمى
// وظيفي بالكامل، للحفاظ على اختصار الرسالة)
function buildWhatsAppAlertMessage(report, regions) {
  const mapLink = regions ? reportMapLink(report, regions) : null;
  const notifiedCodes = (dynamicSettings.whatsappNumbers || [])
    .map((n, i) => ({ ...n, code: i + 1 }))
    .filter((n) => n.phone)
    .map((n) => n.code);
  let msg =
    `🚨 حالة طوارئ شديدة الخطورة\n` +
    `نوع الحادثة: ${report.incidentType}\n` +
    `المُبلِّغ: ${report.reporterName} - ${report.phone}\n` +
    `العنوان: ${report.address || "—"}\n` +
    `رقم البلاغ: ${report.crn || "—"}\n`;
  if (mapLink) msg += `📍 موقع الحادثة: ${mapLink}\n`;
  if (dynamicSettings.zoomLink) msg += `🎥 اجتماع الطوارئ (Zoom): ${dynamicSettings.zoomLink}\n`;
  if (notifiedCodes.length > 0) msg += `✅ تم إبلاغ: ${notifiedCodes.join("، ")}\n`;
  msg += `يرجى المتابعة الفورية.`;
  return msg;
}

// يحاول الإرسال التلقائي الكامل عبر سيرفر الشركة (لو مزوَّد بنقطة /api/whatsapp
// متصلة بـ Meta WhatsApp Business API). يُرجع true لو نجح الإرسال التلقائي،
// و false لو غير متاح أو فشل — وقتها يظهر للمستخدم اختيار الأرقام يدوياً بدلاً
// من محاولة فتح كل الروابط تلقائياً (المتصفحات بتمنع فتح أكتر من نافذة واحدة
// تلقائياً في نفس الوقت، فالأفضل يختار المستخدم ويضغط "إرسال" بنفسه لكل رقم).
async function trySelfHostedWhatsApp(report, regions) {
  if (!SELF_API_BASE) return false;
  const numbers = (dynamicSettings.whatsappNumbers || []).filter((n) => n.phone);
  if (numbers.length === 0) return true; // لا يوجد أرقام مسجّلة أصلاً
  try {
    await fetch(`${SELF_API_BASE}/api/whatsapp`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...(SELF_API_KEY ? { "x-api-key": SELF_API_KEY } : {}) },
      body: JSON.stringify({ numbers: numbers.map((n) => n.phone), message: buildWhatsAppAlertMessage(report, regions) }),
    });
    return true;
  } catch (e) {
    return false;
  }
}

// تحديد درجة الخطورة تلقائياً حسب نوع الحادثة — بدلاً من إدخالها يدوياً من
// المُبلِّغ. القيم اجتهادية حسب خطورة كل نوع ويمكن تعديلها لاحقاً حسب سياسة الشركة.
const INCIDENT_SEVERITY_MAP = {
  "تسرب غاز": "1", "رائحة غاز قوية": "1", "غاز ضعيف": "2", "خط غاز مكسور": "1",
  "انفجار غاز": "1", "حريق ناتج عن غاز": "1", "اشتباه في تسرب": "2",
  "عطل عداد مسبق الدفع": "3", "عبث عميل": "2", "تأمين": "3",
  "تسريب شبكة أرضية": "1", "ماس كهربائي": "2", "حريق": "1",
  "عطل عداد ميكانيكي": "3", "إعادة تشغيل": "3", "انهيار عقار": "1",
  "انفجار خط غاز 7 بار": "1", "حريق منظم": "1", "حادثة محطة غاز": "1", "أخرى": "2",
};
function autoSeverity(type) { return INCIDENT_SEVERITY_MAP[type] || "2"; }

// أنواع الحوادث اللي تستوجب تنبيه فوري عبر واتساب للسوبر أدمن ومجموعة
// المهندسين، بالإضافة للتصنيف التلقائي كدرجة "أولى"
const WHATSAPP_ALERT_TYPES = [
  "انفجار خط غاز 7 بار", "خط غاز مكسور", "انفجار غاز",
  "حريق", "حريق ناتج عن غاز", "حريق منظم", "حادثة محطة غاز",
];

// إجراءات خاصة بحالات معينة (بالإضافة للإجراءات القياسية العامة) — بتتحدد
// حسب نوع الحادثة نفسه، وبتظهر لفريق الاستجابة أول ما يبدأ معالجة البلاغ
const INCIDENT_SPECIFIC_PROCEDURES = {
  "انهيار عقار": [
    "الوصول الفوري لمكان البلاغ وعمل كردون أمني لعدم حدوث إصابات ولسهولة وسرعة التعامل مع الحالة",
    "ربط سكويز لعدم سريان الغاز حسب مسارات خطوط الغاز",
    "عمل حصر بحالات الوفيات والإصابات إن وجدت",
    "عمل حصر بأعداد العملاء الذين تم فصل الغاز عنهم",
  ],
  "حريق": [
    "فصل الغاز عن العقار من المحبس الرئيسي (الشاكوش) إذا استدعى الأمر ذلك",
    "في حالة امتداد الحريق للمبنى بالكامل، يتم تأمين المباني المجاورة",
    "بعد تأمين الحالة من قبل الدفاع المدني وانتهاء الحريق، يتم التحري قدر الإمكان عن أسباب الحريق وحصر خسائر الشركة مبدئياً (المواسير والخامات)",
    "عمل حصر بحالات الوفيات والإصابات إن وجدت",
    "عمل حصر بأعداد العملاء الذين تم فصل الغاز عنهم",
    "بعد إخماد الحريق من قبل الحماية المدنية، تتم معاينة العقار للتأكد من سلامة الوصلات وإعادة تشغيل الغاز للحالات التي لا يوجد بها أضرار بوصلات الغاز",
  ],
  "انفجار غاز": [
    "الوصول الفوري للموقع وعمل كردون أمني واسع طبقاً لمستوى الخطورة",
    "إيقاف الحركة المرورية بالموقع بالتنسيق مع الجهات المختصة",
    "قياس نسب الغاز وتحديد حدود منطقة الخطر",
    "إخلاء المواطنين والمنشآت الواقعة داخل نطاق التأثير",
    "منع جميع مصادر الاشتعال والأعمال بالموقع",
    "التنسيق مع غرفة العمليات لعزل الخط المتضرر",
  ],
};
// "حريق ناتج عن غاز" و"خط غاز مكسور" يشتركان في نفس إجراءات "حريق" و"انفجار غاز" على الترتيب
INCIDENT_SPECIFIC_PROCEDURES["حريق ناتج عن غاز"] = INCIDENT_SPECIFIC_PROCEDURES["حريق"];
INCIDENT_SPECIFIC_PROCEDURES["خط غاز مكسور"] = INCIDENT_SPECIFIC_PROCEDURES["انفجار غاز"];
INCIDENT_SPECIFIC_PROCEDURES["انفجار خط غاز 7 بار"] = INCIDENT_SPECIFIC_PROCEDURES["انفجار غاز"];
INCIDENT_SPECIFIC_PROCEDURES["حريق منظم"] = INCIDENT_SPECIFIC_PROCEDURES["حريق"];
INCIDENT_SPECIFIC_PROCEDURES["حادثة محطة غاز"] = INCIDENT_SPECIFIC_PROCEDURES["انفجار غاز"];

// الإجراءات القياسية ببطاقة غرفة العمليات (نموذج FP-36-01)
const OPS_ACTIONS = [
  { id: "confirmTenantPresence", label: "التأكد من العميل بأهمية التواجد بالشقة" },
  { id: "thankReporter", label: "توجيه الشكر للمُبلِّغ مع التأكد من اتخاذ اللازم" },
  { id: "confirmRepairAction", label: "التأكد بأنه سيتم اتخاذ اللازم لإصلاح الجهاز" },
  { id: "directToAuthorities", label: "توجيه المُبلِّغ للاتصال بالجهات المختصة" },
  { id: "closeApplianceValve", label: "إغلاق محبس الأجهزة" },
  { id: "closeMeterValve", label: "إغلاق محبس العداد" },
  { id: "openVents", label: "فتح منافذ التهوية" },
  { id: "noElectricSwitch", label: "لا تستخدم مفتاح كهرباء (فتح/غلق) وأطفئ مصادر اللهب" },
];
function emptyOpsCard() {
  return {
    workOrderNo: "", receivedBy: "", alertRecipient: "",
    actions: {}, arrivalTime: "", endTime: "",
    workDetails: "", performedBy: "",
    review: { actionsMatch: "", priorityMatch: "", arrivalStandard: "", resolutionStandard: "", followUp: "" },
    shiftLeader: "", supervisor: "", engineer: "",
  };
}

function genCode(len = 6) {
  const chars = "0123456789";
  let s = "";
  for (let i = 0; i < len; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}
function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
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
// يحوّل توقيت (ms) إلى صيغة تصلح لحقل <input type="datetime-local">
function toDatetimeLocal(ms) {
  try {
    const d = new Date(ms);
    const pad = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  } catch { return ""; }
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
  // أولاً: لو سيرفر الشركة الخاص متضبط (VITE_API_BASE_URL)، نستخدمه كمصدر رئيسي
  if (SELF_API_BASE) {
    try {
      const val = await selfApiLoad(key);
      if (val !== undefined) {
        memoryStore[key] = val;
        try { if (hasLocalStorage()) window.localStorage.setItem(LOCAL_PREFIX + key, JSON.stringify(val)); } catch (e) {}
        return val;
      }
    } catch (e) {
      // تعذّر الاتصال بسيرفر الشركة — ننتقل للنسخة المحلية كخطة بديلة
    }
  } else {
    // وإلا نستخدم Firestore — المصدر المشترك الافتراضي بين كل المستخدمين والأجهزة
    try {
      const snap = await getDoc(doc(db, FS_COLLECTION, key));
      if (snap.exists()) {
        const val = snap.data().value;
        memoryStore[key] = val;
        try { if (hasLocalStorage()) window.localStorage.setItem(LOCAL_PREFIX + key, JSON.stringify(val)); } catch (e) {}
        return val;
      }
    } catch (e) {
      // لا يوجد إنترنت أو تعذّر الاتصال بـ Firestore — ننتقل للنسخة المحلية كخطة بديلة
    }
  }
  // نسخة محلية احتياطية (localStorage / ذاكرة الجلسة) لو مفيش اتصال بالإنترنت
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

// يعيد { persisted, error }. يكتب في سيرفر الشركة (لو متضبط) أو Firestore، وفي
// localStorage كنسخة احتياطية محلية في نفس الوقت. لو تعذّر الاتصال بالإنترنت،
// يحفظ محلياً فقط ولا يمنع التطبيق من الاستمرار.
async function saveJSON(key, value) {
  memoryStore[key] = value;
  let cloudOk = false, cloudError = null;
  try {
    if (SELF_API_BASE) {
      await selfApiSave(key, value);
    } else {
      await setDoc(doc(db, FS_COLLECTION, key), { value, updatedAt: Date.now() });
    }
    cloudOk = true;
  } catch (e) {
    cloudError = (e && e.message) ? e.message : String(e);
  }
  try {
    if (hasLocalStorage()) window.localStorage.setItem(LOCAL_PREFIX + key, JSON.stringify(value));
  } catch (e) { /* تجاهل */ }
  if (cloudOk) return { persisted: true, error: null };
  return { persisted: false, error: cloudError || "تعذّر الاتصال بقاعدة البيانات — تم الحفظ محلياً فقط مؤقتاً" };
}

/* ============================== أنماط عامة ============================== */

// روابط خرائط جوجل — إما بإحداثيات دقيقة (موقع GPS) أو ببحث نصي (عنوان/منطقة)
function mapsLinkForCoords(lat, lng) {
  return `https://www.google.com/maps?q=${lat},${lng}`;
}
function mapsLinkForQuery(q) {
  return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(q)}`;
}
// روابط خريطة مُضمَّنة (iframe) داخل البرنامج نفسه بدل الاكتفاء برابط خارجي —
// بهذا يشوف فريق الاستجابة موقع الحالة على الخريطة مباشرة من داخل تفاصيل البلاغ
function mapsEmbedForCoords(lat, lng) {
  return `https://www.google.com/maps?q=${lat},${lng}&z=16&output=embed`;
}
function mapsEmbedForQuery(q) {
  return `https://www.google.com/maps?q=${encodeURIComponent(q)}&output=embed`;
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
          <img src="/logo.png" alt="MODERNGAS" style={{ height: 52, margin: "0 auto 10px", display: "block" }} />
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
        <img src="/logo.png" alt="MODERNGAS" style={{ height: 60, margin: "0 auto 12px", display: "block" }} />
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
    severity: "", details: "", region: regions[0]?.id || "",
    address: "", location: null,
  };
  const [form, setForm] = useState(empty);
  const [err, setErr] = useState("");
  const [ok, setOk] = useState(false);
  const [saving, setSaving] = useState(false);
  const [locStatus, setLocStatus] = useState("idle"); // idle | loading | done | error
  const [locErr, setLocErr] = useState("");
  const [lastCrn, setLastCrn] = useState("");
  const [waAlertReport, setWaAlertReport] = useState(null); // بلاغ حرِج يحتاج اختيار أرقام واتساب يدوياً

  function set(k, v) { setForm((f) => ({ ...f, [k]: v })); }
  // اختيار نوع الحادثة يحدد درجة الخطورة تلقائياً — لا تدخل يدوي من المُبلِّغ
  function setIncidentType(t) { setForm((f) => ({ ...f, incidentType: t, severity: autoSeverity(t) })); }

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
      startedAt: null, completedAt: null, opsCard: emptyOpsCard(),
    };
    const next = [report, ...reports];
    await saveJSON("gas_reports", next);
    const alerts = await loadJSON("gas_alerts", []);
    const locNote = form.location ? " 📍 مع الموقع الجغرافي" : "";
    const urgentPrefix = form.severity === "1" ? "🚨 عاجل — " : "";
    await saveJSON("gas_alerts", [{ id: uid(), text: `${urgentPrefix}بلاغ جديد (${sevMeta(form.severity).label}) رقم ${crn} من ${form.reporterName}${locNote}`, ts: Date.now() }, ...alerts].slice(0, 100));
    // حوادث شديدة الخطورة (انفجار/كسر خط غاز/حريق/محطة غاز) تستوجب تنبيه فوري
    // عبر واتساب للسوبر أدمن ومجموعة المهندسين، بجانب التنبيه الداخلي بالأعلى.
    // نحاول أولاً الإرسال التلقائي الكامل عبر سيرفر الشركة (لو متاح)، وإلا
    // نعرض للمستخدم قائمة الأرقام ليختار ويرسل بنفسه (لتفادي منع المتصفح
    // لفتح عدة نوافذ واتساب تلقائياً في نفس الوقت).
    if (WHATSAPP_ALERT_TYPES.includes(form.incidentType)) {
      const sentAutomatically = await trySelfHostedWhatsApp(report, regions);
      if (!sentAutomatically) setWaAlertReport(report);
    }
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
      <select className="mg-select" style={{ marginBottom: 12 }} value={form.incidentType} onChange={(e) => setIncidentType(e.target.value)}>
        <option value="">— اختر —</option>
        {INCIDENT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
      </select>

      {form.incidentType && (
        <div style={{ marginBottom: 12 }}>
          <label className="mg-label">درجة الخطورة (تُحدَّد تلقائياً حسب نوع الحادثة)</label>
          <span className="mg-badge" style={{ background: sevMeta(form.severity).bg, color: sevMeta(form.severity).color, fontSize: 14, padding: "8px 14px" }}>
            {sevMeta(form.severity).emoji} {sevMeta(form.severity).label}
          </span>
        </div>
      )}

      <label className="mg-label">تفاصيل</label>
      <textarea className="mg-textarea" style={{ marginBottom: 14, minHeight: 80 }} value={form.details} onChange={(e) => set("details", e.target.value)} />

      {err && <div style={{ color: "#D93025", fontSize: 13, marginBottom: 10 }}>⚠️ {err}</div>}
      {ok && <div style={{ color: "#188038", fontSize: 13, marginBottom: 10 }}>✅ تم إرسال البلاغ بنجاح — رقم المرجع: <strong style={{ fontFamily: "monospace" }}>{lastCrn}</strong></div>}

      <button className="mg-btn mg-btn-primary" style={{ width: "100%", justifyContent: "center" }} disabled={saving} onClick={submit}>
        {saving ? "جارٍ الإرسال…" : "📤 إرسال البلاغ"}
      </button>

      {waAlertReport && (
        <WhatsAppSendModal report={waAlertReport} regions={regions} onClose={() => setWaAlertReport(null)} />
      )}
    </div>
  );
}

/* ============================== نافذة إرسال تنبيه واتساب يدوياً ============================== */

// تعرض قائمة أرقام السوبر أدمن/المهندسين المسجّلة في الإعدادات، مع خانة اختيار
// لكل رقم وزر إرسال مستقل بجانبه — كل ضغطة تفتح واتساب لنفس الرقم فوراً كنتيجة
// مباشرة لتفاعل المستخدم، فلا يمنعها المتصفح (بعكس فتح عدة نوافذ تلقائياً دفعة واحدة)
function WhatsAppSendModal({ report, regions, onClose }) {
  const numbers = (dynamicSettings.whatsappNumbers || [])
    .map((n, i) => ({ ...n, code: i + 1 }))
    .filter((n) => n.phone);
  const [selected, setSelected] = useState(() => Object.fromEntries(numbers.map((n) => [n.code, true])));
  const [sentTo, setSentTo] = useState({});
  // نص الرسالة قابل للتعديل والإضافة قبل الإرسال — يبدأ بالنص المقترح تلقائياً
  const [message, setMessage] = useState(() => buildWhatsAppAlertMessage(report, regions));

  function toggle(code) { setSelected((s) => ({ ...s, [code]: !s[code] })); }
  function selectAll(v) { setSelected(Object.fromEntries(numbers.map((n) => [n.code, v]))); }
  function sendTo(n) {
    const digits = (n.phone || "").replace(/[^\d]/g, "");
    if (!digits) return;
    window.open(`https://wa.me/${digits}?text=${encodeURIComponent(message)}`, "_blank");
    setSentTo((s) => ({ ...s, [n.code]: true }));
  }

  return (
    <Modal title="📱 إرسال تنبيه واتساب — حالة شديدة الخطورة" onClose={onClose}>
      <div style={{ fontSize: 13, color: "#D93025", fontWeight: 700, marginBottom: 10 }}>
        بلاغ {report.crn} — {report.incidentType}. اختر الأرقام المطلوب تنبيهها وادوس "إرسال" بجانب كل رقم.
      </div>

      <label className="mg-label">نص الرسالة (تقدر تعدّل أو تضيف أي شيء قبل الإرسال)</label>
      <textarea
        className="mg-textarea" style={{ minHeight: 130, marginBottom: 14, fontFamily: "inherit" }}
        value={message} onChange={(e) => setMessage(e.target.value)}
      />

      {dynamicSettings.zoomLink && (
        <button
          type="button" className="mg-btn mg-btn-sm"
          style={{ width: "100%", justifyContent: "center", background: "#2D8CFF", color: "#fff", marginBottom: 14 }}
          onClick={() => window.open(dynamicSettings.zoomLink, "_blank")}
        >
          🎥 بدء اجتماع الطوارئ على Zoom
        </button>
      )}

      {numbers.length === 0 ? (
        <div style={{ fontSize: 13, color: "#8A6C68" }}>
          لا يوجد أرقام واتساب مسجّلة. أضِفها من ⚙️ الإعدادات → أرقام واتساب تنبيهات الطوارئ.
        </div>
      ) : (
        <>
          <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
            <button type="button" className="mg-btn mg-btn-ghost mg-btn-sm" onClick={() => selectAll(true)}>تحديد الكل</button>
            <button type="button" className="mg-btn mg-btn-ghost mg-btn-sm" onClick={() => selectAll(false)}>إلغاء التحديد</button>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 14 }}>
            {numbers.map((n) => (
              <div key={n.code} style={{ display: "flex", alignItems: "center", gap: 10, background: "#FBF0EF", borderRadius: 8, padding: "8px 10px" }}>
                <input type="checkbox" checked={!!selected[n.code]} onChange={() => toggle(n.code)} />
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 700 }}>
                    <span style={{ color: "#D93025", fontFamily: "monospace" }}>({n.code})</span> {n.role || "—"}
                  </div>
                  <div style={{ fontSize: 12, color: "#8A6C68", fontFamily: "monospace" }}>{n.phone}</div>
                  {n.description && <div style={{ fontSize: 11.5, color: "#8A6C68", marginTop: 2 }}>{n.description}</div>}
                </div>
                <button
                  type="button" className="mg-btn mg-btn-sm"
                  style={{ background: sentTo[n.code] ? "#188038" : "#25D366", color: "#fff" }}
                  disabled={!selected[n.code]}
                  onClick={() => sendTo(n)}
                >
                  {sentTo[n.code] ? "✅ أُرسل" : "📤 إرسال"}
                </button>
              </div>
            ))}
          </div>
        </>
      )}

      <button className="mg-btn mg-btn-outline" style={{ width: "100%", justifyContent: "center" }} onClick={onClose}>إغلاق</button>
    </Modal>
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
// نفس منطق reportMapLink لكن يُرجع رابط خريطة "مُضمَّنة" (embed) لعرضها كإطار داخل التطبيق
function reportMapEmbedUrl(r, regions) {
  if (r.location) return mapsEmbedForCoords(r.location.lat, r.location.lng);
  const region = regions.find((rg) => rg.id === r.region);
  if (r.address && r.address.trim()) {
    const q = region ? `${r.address} ${region.name}` : r.address;
    return mapsEmbedForQuery(q);
  }
  if (region && region.lat != null && region.lng != null) return mapsEmbedForCoords(region.lat, region.lng);
  if (region) return mapsEmbedForQuery(region.name);
  return null;
}

function ReportsList({ reports, regions, currentUser, canEdit, onChanged }) {
  const [filter, setFilter] = useState("all");
  const [q, setQ] = useState("");
  const [selected, setSelected] = useState(null);
  const [completing, setCompleting] = useState(null);
  const [exporting, setExporting] = useState(false);
  const [safetyFor, setSafetyFor] = useState(null); // بلاغ "أولى" تم بدء معالجته حالاً — نعرض له إجراءات السلامة فوراً

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
    // بلاغات درجة "أولى" (شديدة الخطورة): تُعرض إجراءات السلامة القياسية
    // فوراً أمام فريق الاستجابة عند بدء المعالجة، بدون الحاجة لفتح بطاقة العمليات
    if (r.severity === "1") setSafetyFor(updated);
  }

  // إكمال الحالة: يسجّل وقت الإنجاز، ويحفظ ملخص التقرير النهائي كتعليق،
  // وأيضاً يسجّله تلقائياً داخل بطاقة غرفة العمليات (وصف الأعمال + وقت
  // الوصول/الانتهاء) بدل ما يتطلب من فريق الاستجابة تعبئتها يدوياً مرة أخرى
  async function completeReport(r, note) {
    const now = Date.now();
    const card = r.opsCard || emptyOpsCard();
    const updatedCard = {
      ...card,
      workDetails: note ? (card.workDetails ? card.workDetails + "\n" + note : note) : card.workDetails,
      arrivalTime: card.arrivalTime || (r.startedAt ? toDatetimeLocal(r.startedAt) : card.arrivalTime),
      endTime: toDatetimeLocal(now),
    };
    const updated = {
      ...r, status: "done", completedAt: now,
      comments: note ? (r.comments ? r.comments + "\n" + note : note) : r.comments,
      opsCard: updatedCard,
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
                <div style={{ fontWeight: 700, display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                  {urgent && <span>🚨</span>}
                  <span>{r.reporterName}</span>
                  <span style={{ fontSize: 11, fontWeight: 700, color: "#D93025", background: "#FDECEA", borderRadius: 8, padding: "1px 7px", fontFamily: "monospace" }}>{r.crn}</span>
                </div>
                <div style={{ fontSize: 12.5, color: "#8A6C68", marginTop: 2 }}>{r.incidentType} · {r.phone}</div>
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

      {safetyFor && (
        <Modal title="🚨 إجراءات التعامل مع الحالات شديدة الخطورة" onClose={() => setSafetyFor(null)}>
          <div style={{ fontSize: 13.5, color: "#D93025", fontWeight: 700, marginBottom: 12 }}>
            بلاغ <strong>{safetyFor.reporterName}</strong> ({safetyFor.crn}) — {safetyFor.incidentType} — درجة الخطورة: 🔴 أولى.
            يجب اتباع الإجراءات التالية فوراً قبل وأثناء التعامل مع الحالة:
          </div>
          {INCIDENT_SPECIFIC_PROCEDURES[safetyFor.incidentType] ? (
            <ol style={{ margin: "0 0 16px", paddingRight: 20, display: "flex", flexDirection: "column", gap: 8 }}>
              {INCIDENT_SPECIFIC_PROCEDURES[safetyFor.incidentType].map((step, i) => (
                <li key={i} style={{ fontSize: 13, color: "#6B4B47", background: "#FBF0EF", padding: "9px 12px", borderRadius: 8, listStyle: "none", display: "flex", gap: 8 }}>
                  <span style={{ fontWeight: 700, color: "#D93025" }}>{i + 1}.</span>
                  <span>{step}</span>
                </li>
              ))}
            </ol>
          ) : (
            <div style={{ display: "grid", gap: 8, marginBottom: 16 }}>
              {OPS_ACTIONS.map((a) => (
                <div key={a.id} style={{ display: "flex", gap: 8, alignItems: "flex-start", fontSize: 13, color: "#6B4B47", background: "#FBF0EF", padding: "9px 12px", borderRadius: 8 }}>
                  <span>⚠️</span><span>{a.label}</span>
                </div>
              ))}
            </div>
          )}
          <button className="mg-btn mg-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={() => setSafetyFor(null)}>
            ✅ فهمت — بدء التنفيذ
          </button>
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
      "رقم أمر شغل الطوارئ": r.opsCard?.workOrderNo || "",
      "متلقي البلاغ": r.opsCard?.receivedBy || "",
      "القائم بالعمل": r.opsCard?.performedBy || "",
      "وصف الأعمال (بطاقة العمليات)": r.opsCard?.workDetails || "",
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
  const [form, setForm] = useState({ ...report, opsCard: report.opsCard || emptyOpsCard() });
  const [opsOpen, setOpsOpen] = useState(false);
  function set(k, v) { setForm((f) => ({ ...f, [k]: v })); }
  // اختيار نوع الحادثة يحدد درجة الخطورة تلقائياً من البرنامج نفسه — بدون تدخل يدوي
  function setIncidentType(t) { setForm((f) => ({ ...f, incidentType: t, severity: autoSeverity(t) })); }
  return (
    <Modal title={`✏️ تفاصيل البلاغ — ${report.crn || ""}`} onClose={onClose}>
      <div className="mg-grid2" style={{ marginBottom: 10 }}>
        <div>
          <label className="mg-label">
            اسم المُبلِّغ{" "}
            <span style={{ fontSize: 11, fontWeight: 700, color: "#D93025", background: "#FDECEA", borderRadius: 8, padding: "1px 7px", fontFamily: "monospace" }}>{form.crn}</span>
          </label>
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
      <select className="mg-select" style={{ marginBottom: 10 }} disabled={!canEdit} value={form.incidentType} onChange={(e) => setIncidentType(e.target.value)}>
        {INCIDENT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
      </select>

      <div className="mg-grid2" style={{ marginBottom: 10 }}>
        <div>
          <label className="mg-label">درجة الأولوية (تلقائي)</label>
          <div>
            <span className="mg-badge" style={{ background: sevMeta(form.severity).bg, color: sevMeta(form.severity).color, fontSize: 13.5, padding: "7px 12px" }}>
              {sevMeta(form.severity).emoji} {sevMeta(form.severity).label}
            </span>
          </div>
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
      {reportMapEmbedUrl(form, regions) && (
        <div style={{ marginBottom: 6, borderRadius: 12, overflow: "hidden", border: "1px solid #F3E9E8" }}>
          <iframe
            title="موقع الحالة" width="100%" height="170" style={{ border: 0, display: "block" }}
            loading="lazy" src={reportMapEmbedUrl(form, regions)}
          />
        </div>
      )}
      {reportMapLink(form, regions) && (
        <div style={{ marginBottom: 10 }}>
          <a href={reportMapLink(form, regions)} target="_blank" rel="noreferrer" style={{ color: "#1A73E8", fontSize: 13 }}>
            🗺️ فتح الموقع في تطبيق خرائط جوجل
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

      <button
        type="button" className="mg-btn mg-btn-ghost mg-btn-sm"
        style={{ width: "100%", justifyContent: "space-between", marginTop: 4, marginBottom: opsOpen ? 8 : 14 }}
        onClick={() => setOpsOpen((o) => !o)}
      >
        <span>🗂️ بطاقة غرفة عمليات (نموذج FP-36-01)</span><span>{opsOpen ? "▲" : "▼"}</span>
      </button>
      {opsOpen && (
        <div style={{ marginBottom: 14, paddingTop: 4, borderTop: "1px dashed #EBD9D6" }}>
          <button
            type="button" className="mg-btn mg-btn-outline mg-btn-sm"
            style={{ width: "100%", justifyContent: "center", marginBottom: 12 }}
            onClick={() => printOpsCard(form, form.opsCard, regions)}
          >🖨️ طباعة البطاقة</button>
          <OpsCardFields card={form.opsCard} canEdit={canEdit} onChange={(c) => set("opsCard", c)} />
        </div>
      )}

      {canEdit ? (
        <button className="mg-btn mg-btn-primary" style={{ width: "100%", justifyContent: "center" }} onClick={() => onSave(form)}>💾 حفظ التعديلات (والبطاقة)</button>
      ) : (
        <div style={{ fontSize: 12.5, color: "#8A6C68", textAlign: "center" }}>عرض فقط — ليست لديك صلاحية التعديل</div>
      )}
    </Modal>
  );
}

/* ============================== بطاقة غرفة عمليات (FP-36-01) ============================== */

/* ============================== طباعة بطاقة غرفة العمليات ============================== */

// يبني صفحة HTML مستقلة بتنسيق قريب من نموذج FP-36-01 الورقي الأصلي، ويفتحها
// في نافذة جديدة ثم يستدعي أمر الطباعة مباشرة — لا حاجة لأي مكتبة خارجية.
function printOpsCard(report, card, regions) {
  const region = regions.find((r) => r.id === report.region);
  const sv = sevMeta(report.severity);
  const yn = (v) => (v ? escapeHtml(v) : "");
  const created = new Date(report.createdAt);
  const dayStr = created.toLocaleDateString("ar-EG", { weekday: "long" });
  const dateStr = created.toLocaleDateString("ar-EG");

  // ترتيب الإجراءات كما في النموذج الورقي الأصلي: عمود يمين لإجراءات
  // التواصل مع المُبلِّغ، وعمود يسار لإجراءات السلامة الفنية (المحابس والتهوية)
  const rightActions = OPS_ACTIONS.slice(0, 4);
  const leftActions = OPS_ACTIONS.slice(4, 8);
  const actionsRowsHtml = rightActions.map((ra, i) => {
    const la = leftActions[i];
    return `<tr>
      <td class="chkcell">${card.actions[ra.id] ? "☑" : "☐"}</td>
      <td class="chklabel">${escapeHtml(ra.label)}</td>
      <td class="chkcell">${card.actions[la.id] ? "☑" : "☐"}</td>
      <td class="chklabel">${escapeHtml(la.label)}</td>
    </tr>`;
  }).join("");

  const sevBox = (id, label) =>
    `<div class="sevbox ${report.severity === id ? "sevbox-on" : ""}">${report.severity === id ? "✓" : ""}<span>${label}</span></div>`;

  const html = `<!doctype html>
<html dir="rtl" lang="ar">
<head>
<meta charset="utf-8" />
<title>بطاقة غرفة عمليات — ${escapeHtml(report.crn || "")}</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; padding: 22px; color: #111; direction: rtl; font-size: 12.5px; }
  .top { display: grid; grid-template-columns: 1fr 2fr 1fr; align-items: center; border-bottom: 2px solid #000; padding-bottom: 8px; margin-bottom: 10px; }
  .top img { height: 46px; }
  .top .formid { font-size: 11px; color: #333; text-align: left; }
  .top h1 { font-size: 15px; margin: 0; text-align: center; }
  .top h2 { font-size: 11.5px; margin: 2px 0 0; text-align: center; font-weight: normal; }
  .region-line { text-align: right; font-size: 12px; margin-bottom: 8px; font-weight: bold; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 8px; }
  td, th { border: 1px solid #333; padding: 5px 8px; font-size: 12px; text-align: right; vertical-align: middle; }
  th { background: #f0f0f0; white-space: nowrap; font-weight: bold; width: 16%; }
  .full { width: 100%; }
  .grid-top { display: grid; grid-template-columns: 2fr 1fr; gap: 0; margin-bottom: 8px; }
  .grid-top table { margin-bottom: 0; }
  .sevwrap { border: 1px solid #333; border-right: none; padding: 6px; display: flex; flex-direction: column; gap: 6px; justify-content: center; }
  .sevbox { display: flex; align-items: center; justify-content: space-between; border: 1px solid #999; border-radius: 3px; padding: 3px 8px; font-size: 11.5px; }
  .sevbox-on { background: #ffe3e0; border-color: #D93025; font-weight: bold; }
  .section-title { background: #333; color: #fff; font-weight: bold; padding: 5px 10px; margin: 10px 0 6px; font-size: 12.5px; }
  .chkcell { width: 26px; text-align: center; font-size: 14px; }
  .chklabel { width: 40%; }
  .sig { display: flex; justify-content: space-between; margin-top: 24px; border-top: 1px solid #333; padding-top: 10px; }
  .sig div { text-align: center; font-size: 11.5px; width: 32%; }
  .sig .line { border-top: 1px solid #333; margin-top: 30px; padding-top: 4px; }
  .footer { display: flex; justify-content: space-between; font-size: 10.5px; color: #444; margin-top: 14px; border-top: 1px dashed #999; padding-top: 6px; }
  .no-print { text-align: center; margin-bottom: 14px; }
  .no-print button { padding: 10px 22px; font-size: 14px; background: #D93025; color: #fff; border: none; border-radius: 8px; cursor: pointer; }
  @media print { .no-print { display: none; } body { padding: 8px; } }
</style>
</head>
<body>
  <div class="no-print"><button onclick="window.print()">🖨️ طباعة</button></div>

  <div class="top">
    <div class="formid">FP-36-01</div>
    <div>
      <h1>الشركة الحديثة للغاز الطبيعي</h1>
      <h2>الإدارة العامة للعمليات والطوارئ — بطاقة غرفة عمليات</h2>
    </div>
    <img src="/logo.png" alt="MODERNGAS" />
  </div>

  <div class="region-line">منطقة عمليات: ${escapeHtml(region?.name || "—")}</div>

  <table>
    <tr><th>اسم المُبلِّغ</th><td>${escapeHtml(report.reporterName)}</td><th>العنوان</th><td>${escapeHtml(report.address || "")}</td></tr>
    <tr><th>اسم العميل</th><td>${escapeHtml(report.subscriberNumber)}</td><th>الرقم الكودي</th><td>${escapeHtml(report.crn || "")}</td></tr>
    <tr><th>تليفون</th><td>${escapeHtml(report.phone)}</td><th>رقم البلاغ</th><td>${escapeHtml(report.crn || "")}</td></tr>
    <tr><th>موضوع البلاغ</th><td colspan="3">${escapeHtml(report.incidentType)}${report.details ? " — " + escapeHtml(report.details) : ""}</td></tr>
  </table>

  <div class="grid-top">
    <table>
      <tr><th>رقم أمر شغل الطوارئ</th><td>${yn(card.workOrderNo)}</td><th>وقت استلام البلاغ</th><td>${escapeHtml(fmtTime(report.createdAt))}</td></tr>
      <tr><th>متلقي البلاغ</th><td>${yn(card.receivedBy)}</td><th>وقت إرسال البلاغ</th><td></td></tr>
      <tr><th>مستلم الإخطار</th><td>${yn(card.alertRecipient)}</td><th>الزمن المستغرق للوصول</th><td>${report.startedAt ? escapeHtml(fmtDuration(report.startedAt - report.createdAt)) : ""}</td></tr>
      <tr><th>اليوم</th><td>${escapeHtml(dayStr)}</td><th>التاريخ</th><td>${escapeHtml(dateStr)}</td></tr>
    </table>
    <div class="sevwrap">
      ${sevBox("1", "أولى")}
      ${sevBox("2", "ثانية")}
      ${sevBox("3", "ثالثة")}
    </div>
  </div>

  <div class="section-title">ضع علامة (✓) على ما تم إنجازه</div>
  <table>${actionsRowsHtml}</table>

  <div class="section-title">تقرير الأعمال</div>
  <table>
    <tr><th>وقت الوصول</th><td>${yn(card.arrivalTime)}</td><th>وقت انتهاء الحالة</th><td>${yn(card.endTime)}</td></tr>
    <tr><th>تفاصيل الأعمال</th><td colspan="3" style="min-height:50px;">${escapeHtml(card.workDetails || "").replace(/\n/g, "<br/>")}</td></tr>
    <tr><th>القائم بالعمل / رقم الأداء</th><td colspan="3">${yn(card.performedBy)}</td></tr>
  </table>

  <div class="section-title">مراجعة الأعمال</div>
  <table>
    <tr><th>الإجراءات المتخذة للحالة</th><td>${yn(card.review.actionsMatch)}</td><th>درجة أولوية الحالة مطابقة</th><td>${yn(card.review.priorityMatch)}</td></tr>
    <tr><th>الوقت المستغرق للوصول قياسي</th><td>${yn(card.review.arrivalStandard)}</td><th>الوقت المستغرق لعلاج الحالة قياسي</th><td>${yn(card.review.resolutionStandard)}</td></tr>
    <tr><th>تعقيب</th><td colspan="3">${escapeHtml(card.review.followUp || "")}</td></tr>
  </table>

  <div class="sig">
    <div>الاسم: ${yn(card.shiftLeader)}<div class="line">رئيس الوردية</div></div>
    <div>الاسم: ${yn(card.supervisor)}<div class="line">مشرف الطوارئ</div></div>
    <div>الاسم: ${yn(card.engineer)}<div class="line">مهندس الطوارئ</div></div>
  </div>

  <div class="footer">
    <span>رقم الإصدار: 1</span>
    <span>تاريخ الإصدار: 2022/8</span>
  </div>
</body>
</html>`;

  const win = window.open("", "_blank", "width=850,height=950");
  if (!win) return alert("المتصفح منع فتح نافذة الطباعة — فعّل النوافذ المنبثقة لهذا الموقع وحاول مرة أخرى");
  win.document.open();
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(() => { try { win.print(); } catch (e) {} }, 400);
}

function OpsCardFields({ card, canEdit, onChange }) {
  function setField(k, v) { onChange({ ...card, [k]: v }); }
  function toggleAction(id) { onChange({ ...card, actions: { ...card.actions, [id]: !card.actions[id] } }); }
  function setReview(k, v) { onChange({ ...card, review: { ...card.review, [k]: v } }); }

  return (
    <div>
      <div className="mg-grid2" style={{ marginBottom: 10 }}>
        <div>
          <label className="mg-label">رقم أمر شغل الطوارئ</label>
          <input className="mg-input" disabled={!canEdit} value={card.workOrderNo} onChange={(e) => setField("workOrderNo", e.target.value)} />
        </div>
        <div>
          <label className="mg-label">متلقي البلاغ</label>
          <input className="mg-input" disabled={!canEdit} value={card.receivedBy} onChange={(e) => setField("receivedBy", e.target.value)} />
        </div>
      </div>

      <label className="mg-label">مستلم الإخطار</label>
      <input className="mg-input" style={{ marginBottom: 12 }} disabled={!canEdit} value={card.alertRecipient} onChange={(e) => setField("alertRecipient", e.target.value)} />

      <label className="mg-label" style={{ marginBottom: 8 }}>الإجراءات المتخذة قبل/عند الوصول</label>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6, marginBottom: 14 }}>
        {OPS_ACTIONS.map((a) => (
          <label key={a.id} style={{ display: "flex", alignItems: "flex-start", gap: 6, fontSize: 12, color: "#6B4B47" }}>
            <input type="checkbox" style={{ marginTop: 2 }} disabled={!canEdit} checked={!!card.actions[a.id]} onChange={() => toggleAction(a.id)} />
            <span>{a.label}</span>
          </label>
        ))}
      </div>

      <div style={{ fontWeight: 700, marginBottom: 8 }}>تقرير الأعمال</div>
      <div className="mg-grid2" style={{ marginBottom: 10 }}>
        <div>
          <label className="mg-label">وقت الوصول</label>
          <input type="datetime-local" className="mg-input" disabled={!canEdit} value={card.arrivalTime} onChange={(e) => setField("arrivalTime", e.target.value)} />
        </div>
        <div>
          <label className="mg-label">وقت انتهاء الحالة</label>
          <input type="datetime-local" className="mg-input" disabled={!canEdit} value={card.endTime} onChange={(e) => setField("endTime", e.target.value)} />
        </div>
      </div>

      <label className="mg-label">وصف الأعمال المنفذة</label>
      <textarea className="mg-textarea" style={{ marginBottom: 10, minHeight: 70 }} disabled={!canEdit} value={card.workDetails} onChange={(e) => setField("workDetails", e.target.value)} />

      <label className="mg-label">القائم بالعمل / رقم الأداء</label>
      <input className="mg-input" style={{ marginBottom: 16 }} disabled={!canEdit} value={card.performedBy} onChange={(e) => setField("performedBy", e.target.value)} />

      <div style={{ fontWeight: 700, marginBottom: 8 }}>مراجعة الأعمال</div>
      {[
        { k: "actionsMatch", label: "الإجراءات المتخذة مطابقة للحالة" },
        { k: "priorityMatch", label: "درجة أولوية الحالة مطابقة" },
        { k: "arrivalStandard", label: "الوقت المستغرق للوصول قياسي" },
        { k: "resolutionStandard", label: "الوقت المستغرق لعلاج الحالة قياسي" },
      ].map((q) => (
        <div key={q.k} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, gap: 8 }}>
          <span style={{ fontSize: 12.5, color: "#6B4B47" }}>{q.label}</span>
          <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
            {["نعم", "لا"].map((v) => (
              <button
                key={v} type="button" disabled={!canEdit} className="mg-btn mg-btn-sm"
                style={{ background: card.review[q.k] === v ? "#D93025" : "#F3E9E8", color: card.review[q.k] === v ? "#fff" : "#6B4B47" }}
                onClick={() => setReview(q.k, v)}
              >{v}</button>
            ))}
          </div>
        </div>
      ))}
      <label className="mg-label">تعقيب</label>
      <textarea className="mg-textarea" style={{ marginBottom: 14, minHeight: 50 }} disabled={!canEdit} value={card.review.followUp} onChange={(e) => setReview("followUp", e.target.value)} />

      <div className="mg-grid2" style={{ marginBottom: 10 }}>
        <div>
          <label className="mg-label">رئيس الوردية</label>
          <input className="mg-input" disabled={!canEdit} value={card.shiftLeader} onChange={(e) => setField("shiftLeader", e.target.value)} />
        </div>
        <div>
          <label className="mg-label">مشرف الطوارئ</label>
          <input className="mg-input" disabled={!canEdit} value={card.supervisor} onChange={(e) => setField("supervisor", e.target.value)} />
        </div>
      </div>
      <label className="mg-label">مهندس الطوارئ</label>
      <input className="mg-input" disabled={!canEdit} value={card.engineer} onChange={(e) => setField("engineer", e.target.value)} />
    </div>
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

function Header({ currentUser, onLogout, onChangeCode, onRefresh, tab, setTab, appName }) {
  const tabsByRole = {
    super_admin: ["report", "list", "stats", "regions", "users", "settings", "alerts"],
    manager: ["report", "list", "stats", "regions", "alerts"],
    response: ["report", "list"],
    observer: ["list", "stats", "alerts"],
    reporter: ["report", "list", "alerts"],
  };
  const tabLabels = {
    report: "📋 بلاغ جديد", list: "📄 البلاغات", stats: "📊 إحصائيات",
    regions: "🗺️ المناطق", users: "👥 المستخدمين", settings: "⚙️ الإعدادات", alerts: "🔔 التنبيهات",
  };
  const tabs = tabsByRole[currentUser.role] || ["list"];

  return (
    <>
      <div className="mg-header">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <img src="/logo.png" alt="MODERNGAS" style={{ height: 32, background: "#fff", borderRadius: 6, padding: "2px 6px" }} />
            <div>
              <div style={{ fontWeight: 800, fontSize: 15 }}>{appName}</div>
              <div style={{ fontSize: 12, opacity: 0.9 }}>{roleIcon(currentUser.role)} {currentUser.name} · {ROLES[currentUser.role]?.label}</div>
            </div>
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

/* ============================== لوحة الإعدادات (اسم البرنامج + الأيقونات) ============================== */

function SettingsPanel({ settings, onSaved }) {
  const [form, setForm] = useState({
    appName: settings.appName || "نظام طوارئ الغاز",
    roleIcons: { ...settings.roleIcons },
    severityIcons: { ...settings.severityIcons },
    whatsappNumbers: migrateWhatsappNumbers(settings.whatsappNumbers),
    zoomLink: settings.zoomLink || "",
  });
  const [saved, setSaved] = useState(false);
  const [exporting, setExporting] = useState(false);

  // يجمع كل بيانات النظام من قاعدة بيانات الشركة (Firestore) في ملف JSON واحد
  // يقدر المسؤول يحمّله وينقله لأي سيرفر أو قاعدة بيانات تانية تخص الشركة لاحقاً.
  async function exportAllData() {
    setExporting(true);
    try {
      const [users, regions, reports, alerts, settingsData] = await Promise.all([
        loadJSON("gas_users", []), loadJSON("gas_regions", []),
        loadJSON("gas_reports", []), loadJSON("gas_alerts", []),
        loadJSON("gas_settings", defaultAppSettings()),
      ]);
      const payload = {
        exportedAt: new Date().toISOString(),
        source: "moderngas-emergency (Firebase)",
        users, regions, reports, alerts, settings: settingsData,
      };
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `gas-emergency-backup-${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      alert("تعذّر تصدير البيانات: " + ((e && e.message) || String(e)));
    }
    setExporting(false);
  }

  function setRoleIcon(role, val) { setForm((f) => ({ ...f, roleIcons: { ...f.roleIcons, [role]: val } })); }
  function setSevIcon(id, val) { setForm((f) => ({ ...f, severityIcons: { ...f.severityIcons, [id]: val } })); }

  // قائمة أرقام واتساب: قابلة للتعديل بالكامل — إضافة، حذف، وتعديل أي حقل
  function setWhatsappField(i, key, val) {
    setForm((f) => ({ ...f, whatsappNumbers: f.whatsappNumbers.map((n, idx) => (idx === i ? { ...n, [key]: val } : n)) }));
  }
  function addWhatsappEntry() {
    setForm((f) => ({ ...f, whatsappNumbers: [...f.whatsappNumbers, { role: "", phone: "", description: "" }] }));
  }
  function removeWhatsappEntry(i) {
    setForm((f) => ({ ...f, whatsappNumbers: f.whatsappNumbers.filter((_, idx) => idx !== i) }));
  }

  async function save() {
    applyAppSettings(form);
    await saveJSON("gas_settings", form);
    onSaved(form);
    setSaved(true);
    setTimeout(() => setSaved(false), 2500);
  }
  function resetDefaults() { setForm(defaultAppSettings()); }

  return (
    <div className="mg-card">
      <h3 style={{ marginTop: 0 }}>⚙️ إعدادات البرنامج</h3>

      <label className="mg-label">اسم البرنامج (يظهر أعلى الشاشة)</label>
      <input
        className="mg-input" style={{ marginBottom: 18 }}
        value={form.appName} onChange={(e) => setForm((f) => ({ ...f, appName: e.target.value }))}
      />

      <div style={{ fontWeight: 700, marginBottom: 10 }}>أيقونات الأدوار (الرموز التعبيرية)</div>
      {Object.entries(ROLES).map(([key, r]) => (
        <div key={key} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
          <span style={{ fontSize: 13, color: "#6B4B47" }}>{r.label}</span>
          <input
            className="mg-input" style={{ width: 70, textAlign: "center" }}
            value={form.roleIcons[key] ?? r.icon}
            onChange={(e) => setRoleIcon(key, e.target.value)}
          />
        </div>
      ))}

      <div style={{ fontWeight: 700, margin: "18px 0 10px" }}>أيقونات درجات الخطورة</div>
      {SEVERITIES.map((s) => (
        <div key={s.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
          <span style={{ fontSize: 13, color: "#6B4B47" }}>{s.label}</span>
          <input
            className="mg-input" style={{ width: 70, textAlign: "center" }}
            value={form.severityIcons[s.id] ?? s.emoji}
            onChange={(e) => setSevIcon(s.id, e.target.value)}
          />
        </div>
      ))}

      <div style={{ marginTop: 20, paddingTop: 14, borderTop: "1px dashed #EBD9D6" }}>
        <div style={{ fontWeight: 700, marginBottom: 6 }}>📱 أرقام واتساب تنبيهات الطوارئ</div>
        <p style={{ fontSize: 12.5, color: "#8A6C68", marginTop: 0, marginBottom: 10 }}>
          تُرسل لهم رسالة واتساب فورية عند تسجيل بلاغ شديد الخطورة (انفجار خط غاز 7 بار، كسر خط غاز،
          حريق، حريق منظم، حادثة محطة غاز). الرقم المختصر بجانب كل جهة يعتمد على ترتيبها في القائمة،
          ويُستخدم في نص الرسالة بدل كتابة المسمى كاملاً — مثال: "✅ تم إبلاغ: 1، 3، 5".
        </p>
        {form.whatsappNumbers.map((n, i) => (
          <div key={i} style={{ display: "flex", gap: 8, alignItems: "flex-start", marginBottom: 10, background: "#FBF0EF", borderRadius: 10, padding: 10 }}>
            <span style={{ fontWeight: 700, color: "#D93025", background: "#FDECEA", borderRadius: 8, padding: "6px 10px", fontFamily: "monospace", flexShrink: 0 }}>{i + 1}</span>
            <div style={{ flex: 1 }}>
              <div className="mg-grid2" style={{ marginBottom: 6 }}>
                <input className="mg-input" placeholder="المسمّى (مثال: مدير المنطقة)" value={n.role} onChange={(e) => setWhatsappField(i, "role", e.target.value)} />
                <input className="mg-input" placeholder="رقم واتساب مع كود الدولة (2010xxxxxxxx)" value={n.phone} onChange={(e) => setWhatsappField(i, "phone", e.target.value)} />
              </div>
              <input className="mg-input" placeholder="وصف اختياري (مثال: متاح على مدار الساعة)" value={n.description} onChange={(e) => setWhatsappField(i, "description", e.target.value)} />
            </div>
            <button type="button" className="mg-btn mg-btn-outline mg-btn-sm" onClick={() => removeWhatsappEntry(i)}>🗑️</button>
          </div>
        ))}
        <button type="button" className="mg-btn mg-btn-ghost mg-btn-sm" style={{ width: "100%", justifyContent: "center" }} onClick={addWhatsappEntry}>
          ➕ إضافة جهة اتصال
        </button>
      </div>

      <div style={{ marginTop: 18, paddingTop: 14, borderTop: "1px dashed #EBD9D6" }}>
        <div style={{ fontWeight: 700, marginBottom: 6 }}>🎥 رابط اجتماع الطوارئ (Zoom)</div>
        <p style={{ fontSize: 12.5, color: "#8A6C68", marginTop: 0, marginBottom: 10 }}>
          رابط ثابت لغرفة اجتماعات Zoom الخاصة بالسوبر أدمن والمهندسين — يُرفق تلقائياً في رسائل
          تنبيه الحالات شديدة الخطورة، ويظهر كزر مباشر لبدء الاجتماع عند إرسال التنبيه.
        </p>
        <input
          className="mg-input" placeholder="https://zoom.us/j/xxxxxxxxxx?pwd=..."
          value={form.zoomLink} onChange={(e) => setForm((f) => ({ ...f, zoomLink: e.target.value }))}
        />
      </div>

      <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
        <button className="mg-btn mg-btn-primary" style={{ flex: 1, justifyContent: "center" }} onClick={save}>💾 حفظ الإعدادات</button>
        <button className="mg-btn mg-btn-outline" onClick={resetDefaults}>↩️ استعادة الافتراضي</button>
      </div>
      {saved && <div style={{ color: "#188038", fontSize: 13, marginTop: 10, textAlign: "center" }}>✅ تم الحفظ — التغييرات ظاهرة الآن في كل الشاشات</div>}

      <div style={{ marginTop: 22, paddingTop: 14, borderTop: "1px dashed #EBD9D6" }}>
        <div style={{ fontWeight: 700, marginBottom: 8 }}>📦 نسخة احتياطية / نقل البيانات</div>
        <p style={{ fontSize: 12.5, color: "#8A6C68", marginTop: 0, marginBottom: 10 }}>
          تحميل نسخة كاملة من كل البيانات (المستخدمين، البلاغات، المناطق، التنبيهات) في ملف JSON واحد،
          لحفظها أو نقلها لأي سيرفر أو قاعدة بيانات خاصة بالشركة لاحقاً.
        </p>
        <button className="mg-btn mg-btn-outline" style={{ width: "100%", justifyContent: "center" }} onClick={exportAllData} disabled={exporting}>
          {exporting ? "⏳ جاري التصدير…" : "⬇️ تحميل نسخة كاملة من البيانات (JSON)"}
        </button>
      </div>

      <div style={{ marginTop: 20, paddingTop: 14, borderTop: "1px dashed #EBD9D6", fontSize: 12.5, color: "#8A6C68" }}>
        🔑 لتغيير <b>كود الدخول الخاص بك</b>، استخدم زر المفتاح أعلى الصفحة بجانب زر تسجيل الخروج.
      </div>
    </div>
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
  const [settings, setSettings] = useState(defaultAppSettings());
  const [tab, setTab] = useState("report");
  const [showChangeCode, setShowChangeCode] = useState(false);
  const pollRef = useRef(null);

  const loadAll = useCallback(async () => {
    const [u, r, rep, al, st] = await Promise.all([
      loadJSON("gas_users", []), loadJSON("gas_regions", []),
      loadJSON("gas_reports", []), loadJSON("gas_alerts", []),
      loadJSON("gas_settings", defaultAppSettings()),
    ]);
    setUsers(u); setRegions(r); setReports(rep); setAlerts(al);
    applyAppSettings(st); setSettings(st);
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
      pollRef.current = setInterval(() => { loadAll(); }, 8000);
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
        appName={settings.appName || currentAppName()}
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
        {tab === "settings" && currentUser.role === "super_admin" && (
          <SettingsPanel settings={settings} onSaved={(s) => setSettings(s)} />
        )}
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
