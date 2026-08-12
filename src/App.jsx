import React, {
  useState,
  useEffect,
  useCallback,
  useRef,
  useMemo,
} from "react";
import { supabase } from "./lib/supabaseClient";
import {
  LayoutDashboard,
  Users,
  Building2,
  Clock,
  Wallet,
  LogIn,
  LogOut,
  Plus,
  Pencil,
  Trash2,
  X,
  CalendarDays,
  Search,
  AlertCircle,
  CheckCircle2,
  Receipt,
  BadgeCheck,
  Loader2,
  Menu,
  Lock,
  ShieldCheck,
  KeyRound,
  UserCircle2,
  ArrowLeft,
  Copy,
  QrCode,
  Store,
  Watch,
  MapPin,
  Bell,
} from "lucide-react";

/* ---------------------------------------------------------------
   Tokens
----------------------------------------------------------------*/
const T = {
  ink: "#12203D",
  inkDark: "#0B1730",
  paper: "#F5F2EA",
  card: "#FFFFFF",
  forest: "#2E6F4E",
  forestDark: "#245A3F",
  forestSoft: "#E8F2EC",
  forestText: "#215D3F",
  clay: "#B5502F",
  gold: "#C08A2E",
  goldSoft: "#FBF1DF",
  goldText: "#8A5E14",
  rose: "#A93E4C",
  roseDark: "#8C3140",
  roseSoft: "#F3E9E9",
  blue: "#3E5C8A",
  line: "#E7E2D6",
  lineSoft: "#EEE9DC",
  muted: "#8A8577",
  mutedLight: "#B0AA98",
  text: "#12203D",
  textSoft: "#4A4638",
};
const PALETTE = [
  "#2E6F4E",
  "#C08A2E",
  "#3E5C8A",
  "#A93E4C",
  "#6B5B95",
  "#B5502F",
];
// Legacy single admin password — kept only as a fallback reference; login now
// checks against the ADMINS list below, which supports multiple accounts
// with different permission levels.
const ADMIN_PASSWORD = "admin123";
const ADMIN_ROLE_LABEL = {
  superadmin: "អ្នកគ្រប់គ្រងជាន់ខ្ពស់",
  manager: "អ្នកគ្រប់គ្រង HR",
};
// Standard working days used to derive a daily rate from monthly salary
// when calculating unpaid-absence deductions in payroll.
const WORKING_DAYS_PER_MONTH = 26;

/* ---------------------------------------------------------------
   Global stylesheet + fonts
----------------------------------------------------------------*/
const STYLE_ID = "wf-suite-style";
const CSS = `
html,body,#root{height:100%;}
.wf-root{display:flex;height:100vh;height:100dvh;min-height:640px;max-height:100vh;max-height:100dvh;background:${T.paper};font-family:'Inter',sans-serif;color:${T.text};position:relative;overflow:hidden;border-radius:16px;box-shadow:0 1px 2px rgba(18,32,61,0.05),0 20px 48px -16px rgba(18,32,61,0.22);}
.wf-sidebar{background:linear-gradient(175deg,${T.ink} 0%,${T.inkDark} 100%);color:#fff;width:250px;flex-shrink:0;display:flex;flex-direction:column;transition:transform .25s cubic-bezier(.4,0,.2,1);}
.wf-sidebar-inner{display:flex;flex-direction:column;height:100%;}
.wf-logo-badge{width:34px;height:34px;border-radius:10px;background:linear-gradient(145deg,${T.forest},${T.forestDark});display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;font-family:'Space Grotesk',sans-serif;flex-shrink:0;box-shadow:0 2px 8px rgba(46,111,78,0.4),inset 0 1px 0 rgba(255,255,255,0.18);}
.wf-nav-item{position:relative;width:100%;display:flex;align-items:center;gap:11px;padding:10px 14px;border-radius:10px;font-size:13.5px;font-weight:500;background:transparent;color:#AEB8CC;border:none;cursor:pointer;text-align:left;transition:background .15s ease,color .15s ease;}
.wf-nav-item:hover{background:rgba(255,255,255,0.07);color:#fff;}
.wf-nav-item.active{background:rgba(255,255,255,0.09);color:#fff;font-weight:600;}
.wf-nav-item.active::before{content:"";position:absolute;left:-10px;top:8px;bottom:8px;width:3px;border-radius:0 3px 3px 0;background:${T.gold};}
.wf-main{flex:1;min-width:0;min-height:0;display:flex;flex-direction:column;background:${T.paper};}
.wf-header{background:rgba(255,255,255,0.92);backdrop-filter:blur(8px);border-bottom:1px solid ${T.lineSoft};padding:14px 22px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:20;}
.wf-content{flex:1;overflow-y:auto;padding:22px;}
.wf-card{background:${T.card};border-radius:14px;border:1px solid ${T.line};box-shadow:0 1px 2px rgba(18,32,61,0.04);transition:box-shadow .15s ease;}
.wf-btn{display:inline-flex;align-items:center;gap:6px;font-weight:600;border-radius:10px;font-size:13px;padding:9px 15px;border:1px solid transparent;cursor:pointer;transition:background .15s ease,transform .1s ease,box-shadow .15s ease;}
.wf-btn:active:not(:disabled){transform:scale(.97);}
.wf-btn:disabled{opacity:.5;cursor:not-allowed;}
.wf-btn-sm{padding:6px 10px;font-size:12px;}
.wf-btn-primary{background:${T.ink};color:#fff;box-shadow:0 1px 2px rgba(18,32,61,0.18);}
.wf-btn-primary:hover:not(:disabled){background:${T.inkDark};}
.wf-btn-accent{background:${T.forest};color:#fff;box-shadow:0 1px 2px rgba(46,111,78,0.22);}
.wf-btn-accent:hover:not(:disabled){background:${T.forestDark};}
.wf-btn-ghost{background:transparent;color:${T.ink};border-color:${T.line};}
.wf-btn-ghost:hover:not(:disabled){background:${T.paper};}
.wf-btn-danger{background:transparent;color:${T.rose};border-color:#E4C7CB;}
.wf-btn-danger:hover:not(:disabled){background:#F8ECEE;}
.wf-btn-danger-solid{background:${T.rose};color:#fff;}
.wf-btn-danger-solid:hover:not(:disabled){background:${T.roseDark};}
.wf-input{width:100%;padding:9px 12px;border-radius:10px;border:1px solid #D8D2C2;font-size:13px;background:#FDFCF9;color:${T.text};outline:none;font-family:inherit;transition:border-color .15s ease,box-shadow .15s ease;}
.wf-input:focus{border-color:${T.forest};box-shadow:0 0 0 3px rgba(46,111,78,0.15);}
.wf-field-label{display:block;font-size:11px;font-weight:700;color:#6B6455;margin-bottom:6px;text-transform:uppercase;letter-spacing:.04em;}
.wf-modal-overlay{position:fixed;inset:0;z-index:50;display:flex;align-items:center;justify-content:center;padding:16px;background:rgba(18,32,61,0.5);backdrop-filter:blur(2px);animation:wf-fade .15s ease;}
.wf-modal{background:#fff;border-radius:16px;box-shadow:0 24px 64px rgba(18,32,61,0.35);width:100%;max-height:90vh;overflow-y:auto;animation:wf-pop .18s cubic-bezier(.2,.9,.3,1.2);}
.wf-modal-head{display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid ${T.lineSoft};position:sticky;top:0;background:#fff;border-radius:16px 16px 0 0;}
.wf-avatar{border-radius:999px;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:600;flex-shrink:0;}
.wf-badge{display:inline-block;padding:4px 10px;border-radius:8px;font-size:11px;font-weight:700;letter-spacing:.02em;white-space:nowrap;}
.wf-table{width:100%;font-size:13px;border-collapse:collapse;}
.wf-table th{text-align:left;font-size:11px;color:${T.muted};text-transform:uppercase;padding:11px 16px;background:#FAF8F2;border-bottom:1px solid ${T.lineSoft};font-weight:700;letter-spacing:.03em;}
.wf-table td{padding:10px 16px;border-bottom:1px solid #F5F2EA;}
.wf-table tr:last-child td{border-bottom:none;}
.wf-table tbody tr{transition:background .12s ease;}
.wf-table tbody tr:hover{background:#FAF8F2;}
.wf-grid{display:grid;gap:16px;}
.wf-punch-clock{font-size:34px;font-weight:700;font-family:'JetBrains Mono',monospace;color:${T.ink};font-variant-numeric:tabular-nums;}
.wf-menu-btn{display:none;background:none;border:none;color:${T.ink};cursor:pointer;padding:4px;}
.wf-overlay-scrim{display:none;}
.wf-content::-webkit-scrollbar,.wf-sidebar nav::-webkit-scrollbar,.wf-modal::-webkit-scrollbar{width:8px;}
.wf-content::-webkit-scrollbar-thumb,.wf-modal::-webkit-scrollbar-thumb{background:${T.mutedLight};border-radius:8px;}
.wf-content::-webkit-scrollbar-track,.wf-modal::-webkit-scrollbar-track{background:transparent;}
.wf-sidebar nav::-webkit-scrollbar-thumb{background:rgba(255,255,255,0.15);border-radius:8px;}
@keyframes wf-fade{from{opacity:0}to{opacity:1}}
@keyframes wf-pop{from{opacity:0;transform:scale(.96) translateY(6px)}to{opacity:1;transform:scale(1) translateY(0)}}
@media (max-width: 820px){
  .wf-sidebar{position:absolute;inset:0 auto 0 0;z-index:40;transform:translateX(-100%);height:100%;box-shadow:8px 0 24px rgba(0,0,0,0.25);}
  .wf-sidebar.open{transform:translateX(0);}
  .wf-menu-btn{display:inline-flex;}
  .wf-overlay-scrim.open{display:block;position:absolute;inset:0;background:rgba(18,32,61,0.45);z-index:35;backdrop-filter:blur(1px);}
  .wf-header{padding:12px 16px;}
  .wf-content{padding:16px;}
}
`;
function useGlobalStyle() {
  useEffect(() => {
    if (!document.getElementById(STYLE_ID)) {
      const tag = document.createElement("style");
      tag.id = STYLE_ID;
      tag.innerHTML = CSS;
      document.head.appendChild(tag);
    }
    if (!document.getElementById("wf-fonts")) {
      const link = document.createElement("link");
      link.id = "wf-fonts";
      link.rel = "stylesheet";
      link.href =
        "https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap";
      document.head.appendChild(link);
    }
  }, []);
}

/* ---------------------------------------------------------------
   Data helpers
----------------------------------------------------------------*/
const K = {
  DEPARTMENTS: "hrsuite:departments",
  EMPLOYEES: "hrsuite:employees",
  SHIFTS: "hrsuite:shifts",
  ATTENDANCE: "hrsuite:attendance",
  PAYROLL: "hrsuite:payroll",
  LEAVE_REQUESTS: "hrsuite:leaverequests",
  ADMINS: "hrsuite:admins",
  OFFICE_LOCATION: "hrsuite:officelocation",
  SESSION_ADMIN: "hrsuite:session:admin",
  SESSION_EMPLOYEE: "hrsuite:session:employee",
};
const DEFAULT_DEPARTMENTS = [
  {
    id: "d1",
    name: "ធនធានមនុស្ស",
    code: "HR",
    desc: "គ្រប់គ្រងបុគ្គលិក និងគោលនយោបាយក្រុមហ៊ុន",
  },
  {
    id: "d2",
    name: "បច្ចេកវិទ្យា",
    code: "IT",
    desc: "ប្រព័ន្ធព័ត៌មាន និងហេដ្ឋារចនាសម្ព័ន្ធ",
  },
  { id: "d3", name: "ហិរញ្ញវត្ថុ", code: "FIN", desc: "គណនេយ្យ និងថវិកា" },
  { id: "d4", name: "ទីផ្សារ", code: "MKT", desc: "ទីផ្សារ និងលក់" },
];
// Two permission levels:
// - superadmin: everything, including deleting records and managing
//   other admin accounts.
// - manager: day-to-day HR work (add/edit employees, approve leave,
//   run payroll) but cannot delete records or touch admin accounts.
const DEFAULT_ADMINS = [
  {
    id: "ad1",
    username: "admin",
    password: "admin123",
    name: "អ្នកគ្រប់គ្រងប្រព័ន្ធ",
    role: "superadmin",
  },
  {
    id: "ad2",
    username: "manager",
    password: "manager123",
    name: "អ្នកគ្រប់គ្រង HR",
    role: "manager",
  },
];
const DEFAULT_SHIFTS = [
  { id: "s1", name: "វេនព្រឹក", start: "06:00", end: "14:00" },
  { id: "s2", name: "វេនថ្ងៃ", start: "14:00", end: "22:00" },
  { id: "s3", name: "វេនយប់", start: "22:00", end: "06:00" },
];
const DEFAULT_EMPLOYEES = [
  {
    id: "e1",
    code: "EMP-001",
    pin: "1001",
    name: "លោក សុវណ្ណ ដារា",
    deptId: "d2",
    shiftId: "s1",
    role: "IT Manager",
    email: "dara.sovann@company.com",
    phone: "012 345 678",
    salary: 950,
    status: "active",
    joined: "2024-03-11",
  },
  {
    id: "e2",
    code: "EMP-002",
    pin: "1002",
    name: "កញ្ញា ចាន់ សុភា",
    deptId: "d1",
    shiftId: "s1",
    role: "HR Officer",
    email: "sophea.chan@company.com",
    phone: "096 234 551",
    salary: 620,
    status: "active",
    joined: "2024-07-02",
  },
  {
    id: "e3",
    code: "EMP-003",
    pin: "1003",
    name: "លោក ហេង សុខា",
    deptId: "d3",
    shiftId: "s2",
    role: "Accountant",
    email: "sokha.heng@company.com",
    phone: "070 998 213",
    salary: 700,
    status: "active",
    joined: "2025-01-20",
  },
];
function colorFor(seed) {
  let h = 0;
  for (let i = 0; i < seed.length; i++) h = seed.charCodeAt(i) + ((h << 5) - h);
  return PALETTE[Math.abs(h) % PALETTE.length];
}
function initials(name) {
  const parts = (name || "?").trim().split(/\s+/);
  return (parts[0]?.[0] || "").concat(parts[1]?.[0] || "").toUpperCase();
}
function todayStr() {
  return new Date().toISOString().slice(0, 10);
}
function monthKey(d = new Date()) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}
// Human-readable label for a "YYYY-MM" key, e.g. "2026-01" -> "មករា 2026".
function monthLabel(mk) {
  const [y, m] = mk.split("-").map(Number);
  try {
    return new Intl.DateTimeFormat("km-KH", {
      year: "numeric",
      month: "long",
    }).format(new Date(y, m - 1, 1));
  } catch {
    return mk;
  }
}
function fmtMoney(n) {
  return (
    "$" +
    Number(n || 0).toLocaleString(undefined, {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })
  );
}
function timeNow() {
  const d = new Date();
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}
function shiftLabel(shift) {
  if (!shift) return "—";
  return `${shift.name} · ${shift.start}–${shift.end}`;
}
function isOvernightShift(shift) {
  return !!shift && shift.end <= shift.start;
}
// Determines whether a check-in time counts as late for a given shift.
// Overnight shifts (e.g. 22:00–06:00) wrap past midnight, so a check-in is
// only compared against the shift's start time, not the end.
function isLateForShift(checkInTime, shift) {
  if (!shift) return checkInTime > "09:00"; // fallback when no shift assigned
  return checkInTime > shift.start;
}
function uid(p) {
  return p + Math.random().toString(36).slice(2, 9);
}
// Minutes of grace after a shift's start time before a missing check-in
// is surfaced as a "missed clock-in" notification.
const LATE_GRACE_MINUTES = 15;
function addMinutesToClock(hhmm, mins) {
  const [h, m] = hhmm.split(":").map(Number);
  const total = (((h * 60 + m + mins) % 1440) + 1440) % 1440;
  return `${String(Math.floor(total / 60)).padStart(2, "0")}:${String(total % 60).padStart(2, "0")}`;
}
function timeAgoLabel(iso) {
  if (!iso) return "";
  const mins = Math.round((Date.now() - new Date(iso).getTime()) / 60000);
  if (mins < 1) return "ឥឡូវនេះ";
  if (mins < 60) return `${mins} នាទីមុន`;
  const hrs = Math.round(mins / 60);
  if (hrs < 24) return `${hrs} ម៉ោងមុន`;
  return `${Math.round(hrs / 24)} ថ្ងៃមុន`;
}
// Notifications are derived live from existing data rather than stored
// separately: admins/managers see new pending leave requests and active
// employees who haven't clocked in today past their shift's start time;
// employees see decisions on their own leave requests.
function buildNotifications({
  role,
  currentEmp,
  employees,
  shifts,
  attendance,
  leaveRequests,
}) {
  const list = [];
  if (role === "admin") {
    leaveRequests
      .filter((r) => r.status === "pending")
      .forEach((r) => {
        const emp = employees.find((e) => e.id === r.employeeId);
        list.push({
          id: `lr-pending-${r.id}`,
          page: "leave",
          tone: "gold",
          title: "សំណើសុំច្បាប់ថ្មី",
          message: `${emp?.name || "?"} បានស្នើសុំ${LEAVE_TYPE_LABEL[r.type] || "ច្បាប់"}`,
          time: r.createdAt,
        });
      });
    const today = todayStr();
    const now = timeNow();
    employees
      .filter((e) => e.status === "active" && e.shiftId)
      .forEach((e) => {
        const shift = shifts.find((s) => s.id === e.shiftId);
        if (!shift || isOvernightShift(shift)) return;
        const alreadyLogged = attendance.some(
          (a) => a.employeeId === e.id && a.date === today,
        );
        if (alreadyLogged) return;
        if (now <= addMinutesToClock(shift.start, LATE_GRACE_MINUTES)) return;
        list.push({
          id: `mc-${e.id}-${today}`,
          page: "attendance",
          tone: "rose",
          title: "មិនទាន់ចុះឈ្មោះចូលធ្វើការ",
          message: `${e.name} មិនទាន់ចុះឈ្មោះចូលធ្វើការទេ (${shift.name} ${shift.start})`,
          time: `${today}T${now}:00`,
        });
      });
  } else if (currentEmp) {
    leaveRequests
      .filter(
        (r) =>
          r.employeeId === currentEmp.id &&
          (r.status === "approved" || r.status === "rejected") &&
          r.reviewedAt,
      )
      .forEach((r) => {
        list.push({
          id: `lr-decided-${r.id}`,
          page: "leave",
          tone: r.status === "approved" ? "forest" : "rose",
          title:
            r.status === "approved"
              ? "សំណើសុំច្បាប់របស់អ្នកត្រូវបានអនុម័ត"
              : "សំណើសុំច្បាប់របស់អ្នកត្រូវបានបដិសេធ",
          message: `${LEAVE_TYPE_LABEL[r.type] || "ច្បាប់"} (${r.startDate} – ${r.endDate})`,
          time: r.reviewedAt,
        });
      });
  }
  return list.sort((a, b) => (b.time || "").localeCompare(a.time || ""));
}
function randomPin() {
  return String(Math.floor(1000 + Math.random() * 9000));
}
// Great-circle distance between two lat/lng points, in meters (Haversine
// formula). Used to check an employee's punch-in location against the
// configured office location.
function distanceMeters(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = (d) => (d * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}
// Promise wrapper around the browser geolocation API.
function getCurrentPosition(options) {
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) {
      reject(new Error("no-geo"));
      return;
    }
    navigator.geolocation.getCurrentPosition(
      (pos) => resolve(pos.coords),
      (err) => reject(err),
      { enableHighAccuracy: true, timeout: 10000, ...options },
    );
  });
}
// Summarizes an employee's attendance for a given "YYYY-MM" month key.
// - absentDays: unauthorized absences → unpaid, reduces salary
// - leaveDays: approved leave → paid, does not reduce salary (informational only)
function monthAttendanceStats(attendance, employeeId, mk) {
  let absentDays = 0;
  let leaveDays = 0;
  let lateDays = 0;
  for (const a of attendance) {
    if (a.employeeId !== employeeId || !a.date.startsWith(mk)) continue;
    if (a.status === "absent") absentDays++;
    else if (a.status === "leave") leaveDays++;
    else if (a.status === "late") lateDays++;
  }
  return { absentDays, leaveDays, lateDays };
}
// Computes payroll figures for one employee for a given month, factoring in
// unpaid absences recorded in attendance. Leave is paid and does not deduct.
function computePayroll(emp, attendance, mk) {
  const { absentDays, leaveDays, lateDays } = monthAttendanceStats(
    attendance,
    emp.id,
    mk,
  );
  const dailyRate = emp.salary / WORKING_DAYS_PER_MONTH;
  const absenceDeduction = Math.min(emp.salary, absentDays * dailyRate);
  const adjustedBase = emp.salary - absenceDeduction;
  const tax = adjustedBase * 0.05;
  const insurance = adjustedBase * 0.02;
  const net = adjustedBase - tax - insurance;
  return {
    absentDays,
    leaveDays,
    lateDays,
    dailyRate,
    absenceDeduction,
    adjustedBase,
    tax,
    insurance,
    net,
  };
}

/* ---------------------------------------------------------------
   Tiny hash router — keeps the employee portal and admin portal
   as two separate URLs (#/employee and #/admin) inside one app.
----------------------------------------------------------------*/
function normalizeHash(h) {
  const clean = (h || "").replace(/^#\/?/, "");
  return clean === "employee" ? "employee" : "admin";
}
function usePortalRoute() {
  const [route, setRoute] = useState(() => normalizeHash(window.location.hash));
  useEffect(() => {
    const onHash = () => setRoute(normalizeHash(window.location.hash));
    window.addEventListener("hashchange", onHash);
    return () => window.removeEventListener("hashchange", onHash);
  }, []);
  const go = useCallback((next) => {
    window.location.hash = next ? `/${next}` : "";
  }, []);
  return [route, go];
}

/* ---------------------------------------------------------------
   Supabase-backed persistence hooks
   These replace the old Claude-Artifact-only window.storage hook.
   Each one keeps the same [value, setValue, ready] shape the rest
   of the app already expects, so components below don't change.
----------------------------------------------------------------*/

// Generic hook for a Supabase table holding an array of rows keyed by `id`.
// setValue is called elsewhere in the app with the FULL next array (never
// an updater function), so on every call we diff against the previous
// array to figure out which rows to upsert and which to delete.
function useSupabaseArray(table, { fromDb, toDb, orderBy } = {}) {
  const [value, setValueState] = useState([]);
  const [ready, setReady] = useState(false);
  const prevRef = useRef([]);
  const mapFromDb = fromDb || ((r) => r);
  const mapToDb = toDb || ((r) => r);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      let query = supabase.from(table).select("*");
      if (orderBy) query = query.order(orderBy);
      const { data, error } = await query;
      if (cancelled) return;
      if (error) {
        console.error(`[supabase] failed to load ${table}:`, error.message);
        prevRef.current = [];
        setValueState([]);
      } else {
        const mapped = (data || []).map(mapFromDb);
        prevRef.current = mapped;
        setValueState(mapped);
      }
      setReady(true);
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [table]);

  const setValue = useCallback(
    (next) => {
      const prev = prevRef.current;
      prevRef.current = next;
      setValueState(next);

      const nextIds = new Set(next.map((r) => r.id));
      const toDelete = prev.filter((r) => !nextIds.has(r.id)).map((r) => r.id);
      const toUpsert = next.filter((r) => {
        const old = prev.find((p) => p.id === r.id);
        return !old || JSON.stringify(old) !== JSON.stringify(r);
      });

      (async () => {
        if (toDelete.length) {
          const { error } = await supabase
            .from(table)
            .delete()
            .in("id", toDelete);
          if (error)
            console.error(
              `[supabase] delete failed on ${table}:`,
              error.message,
            );
        }
        if (toUpsert.length) {
          const { error } = await supabase
            .from(table)
            .upsert(toUpsert.map(mapToDb));
          if (error)
            console.error(
              `[supabase] upsert failed on ${table}:`,
              error.message,
            );
        }
      })();
    },
    [table, mapToDb],
  );

  return [value, setValue, ready];
}

// payroll_paid is stored as one row per (employee, month) but the app
// works with it as a flat map: { "<employeeId>-<YYYY-MM>": true }.
// Employee ids never contain "-", so splitting on the first "-" safely
// recovers employeeId and month from a key.
function splitPayrollKey(key) {
  const i = key.indexOf("-");
  return [key.slice(0, i), key.slice(i + 1)];
}
function usePayrollPaid() {
  const [value, setValueState] = useState({});
  const [ready, setReady] = useState(false);
  const prevRef = useRef({});

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const { data, error } = await supabase.from("payroll_paid").select("*");
      if (cancelled) return;
      if (error) {
        console.error("[supabase] failed to load payroll_paid:", error.message);
        prevRef.current = {};
        setValueState({});
      } else {
        const map = {};
        (data || []).forEach((r) => {
          map[`${r.employee_id}-${r.month}`] = r.paid;
        });
        prevRef.current = map;
        setValueState(map);
      }
      setReady(true);
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const setValue = useCallback((next) => {
    const prev = prevRef.current;
    prevRef.current = next;
    setValueState(next);

    const changed = Object.keys(next).filter((k) => next[k] !== prev[k]);
    const removed = Object.keys(prev).filter((k) => !(k in next));

    (async () => {
      for (const key of changed) {
        const [employeeId, month] = splitPayrollKey(key);
        const { error } = await supabase
          .from("payroll_paid")
          .upsert({ employee_id: employeeId, month, paid: next[key] });
        if (error)
          console.error(
            "[supabase] upsert failed on payroll_paid:",
            error.message,
          );
      }
      for (const key of removed) {
        const [employeeId, month] = splitPayrollKey(key);
        const { error } = await supabase
          .from("payroll_paid")
          .delete()
          .eq("employee_id", employeeId)
          .eq("month", month);
        if (error)
          console.error(
            "[supabase] delete failed on payroll_paid:",
            error.message,
          );
      }
    })();
  }, []);

  return [value, setValue, ready];
}

// office_location is a single settings row (id = 1). null means "not set".
function useOfficeLocation() {
  const [value, setValueState] = useState(null);
  const [ready, setReady] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const { data, error } = await supabase
        .from("office_location")
        .select("*")
        .eq("id", 1)
        .maybeSingle();
      if (cancelled) return;
      if (error) {
        console.error(
          "[supabase] failed to load office_location:",
          error.message,
        );
        setValueState(null);
      } else if (data && data.lat != null) {
        setValueState({ lat: data.lat, lng: data.lng, radius: data.radius });
      } else {
        setValueState(null);
      }
      setReady(true);
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const setValue = useCallback((next) => {
    setValueState(next);
    (async () => {
      const row = next
        ? { id: 1, lat: next.lat, lng: next.lng, radius: next.radius }
        : { id: 1, lat: null, lng: null, radius: null };
      const { error } = await supabase.from("office_location").upsert(row);
      if (error)
        console.error(
          "[supabase] save failed on office_location:",
          error.message,
        );
    })();
  }, []);

  return [value, setValue, ready];
}

// Login sessions are intentionally per-device, not shared data, so they
// stay in the browser's own localStorage instead of Supabase.
function useLocalStorage(key, fallback) {
  const [value, setValueState] = useState(() => {
    try {
      const raw = window.localStorage.getItem(key);
      return raw != null ? JSON.parse(raw) : fallback;
    } catch {
      return fallback;
    }
  });
  const setValue = useCallback(
    (next) => {
      setValueState(next);
      try {
        window.localStorage.setItem(key, JSON.stringify(next));
      } catch {}
    },
    [key],
  );
  return [value, setValue, true];
}

/* ---------------------------------------------------------------
   Small UI atoms
----------------------------------------------------------------*/
function Avatar({ name, size = 40, photo }) {
  if (photo) {
    return (
      <img
        src={photo}
        alt={name || "avatar"}
        className="wf-avatar"
        style={{
          width: size,
          height: size,
          objectFit: "cover",
          flexShrink: 0,
        }}
      />
    );
  }
  return (
    <div
      className="wf-avatar"
      style={{
        width: size,
        height: size,
        background: colorFor(name || "?"),
        fontSize: size * 0.36,
      }}
    >
      {initials(name)}
    </div>
  );
}
const STATUS_MAP = {
  active: { bg: T.forestSoft, fg: T.forestText, label: "សកម្ម" },
  inactive: { bg: T.roseSoft, fg: T.roseDark, label: "អសកម្ម" },
  present: { bg: T.forestSoft, fg: T.forestText, label: "មកធ្វើការ" },
  late: { bg: T.goldSoft, fg: T.goldText, label: "មកយឺត" },
  absent: { bg: T.roseSoft, fg: T.roseDark, label: "អវត្តមាន" },
  leave: { bg: "#E7ECF6", fg: T.blue, label: "ឈប់សម្រាក" },
  pending: { bg: T.goldSoft, fg: T.goldText, label: "រង់ចាំបង់" },
  paid: { bg: T.forestSoft, fg: T.forestText, label: "បង់រួច" },
  approved: { bg: T.forestSoft, fg: T.forestText, label: "អនុម័តហើយ" },
  rejected: { bg: T.roseSoft, fg: T.roseDark, label: "បដិសេធ" },
};
const LEAVE_TYPE_LABEL = {
  annual: "ច្បាប់ប្រចាំឆ្នាំ",
  sick: "ច្បាប់ឈឺ",
  other: "ផ្សេងៗ",
};
// Returns an array of "YYYY-MM-DD" strings for every day from start to end,
// inclusive. Used to expand an approved leave request into attendance rows.
function dateRange(start, end) {
  const out = [];
  let cur = new Date(start + "T00:00:00");
  const last = new Date(end + "T00:00:00");
  if (isNaN(cur) || isNaN(last) || cur > last) return out;
  while (cur <= last) {
    out.push(cur.toISOString().slice(0, 10));
    cur.setDate(cur.getDate() + 1);
  }
  return out;
}
function StatusPill({ status }) {
  const s = STATUS_MAP[status] || { bg: "#EEE", fg: "#555", label: status };
  return (
    <span className="wf-badge" style={{ background: s.bg, color: s.fg }}>
      {s.label}
    </span>
  );
}
function Card({ children, style, accent, ...rest }) {
  return (
    <div
      className="wf-card"
      style={{
        ...(accent ? { borderLeft: `4px solid ${accent}` } : {}),
        ...style,
      }}
      {...rest}
    >
      {children}
    </div>
  );
}
function Button({ children, variant = "primary", size, style, ...props }) {
  const cls = {
    primary: "wf-btn-primary",
    accent: "wf-btn-accent",
    ghost: "wf-btn-ghost",
    danger: "wf-btn-danger",
    "danger-solid": "wf-btn-danger-solid",
  }[variant];
  return (
    <button
      className={`wf-btn ${cls} ${size === "sm" ? "wf-btn-sm" : ""}`}
      style={style}
      {...props}
    >
      {children}
    </button>
  );
}
const NOTIF_TONE = {
  gold: "#C08A2E",
  rose: "#A93E4C",
  forest: "#2E6F4E",
  blue: "#3E5C8A",
};
function NotificationBell({
  role,
  currentAdmin,
  currentEmp,
  employees,
  shifts,
  attendance,
  leaveRequests,
  setPage,
}) {
  const [open, setOpen] = useState(false);
  const boxRef = useRef(null);
  const userId = role === "admin" ? currentAdmin?.id : currentEmp?.id;
  const [readIds, setReadIds] = useLocalStorage(
    `hrsuite:notifications:read:${userId || "guest"}`,
    [],
  );

  const notifications = useMemo(
    () =>
      buildNotifications({
        role,
        currentEmp,
        employees,
        shifts,
        attendance,
        leaveRequests,
      }),
    [role, currentEmp, employees, shifts, attendance, leaveRequests],
  );
  const unread = notifications.filter((n) => !readIds.includes(n.id));

  useEffect(() => {
    function onDocClick(e) {
      if (boxRef.current && !boxRef.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, []);

  const markRead = (ids) =>
    setReadIds(Array.from(new Set([...readIds, ...ids])));

  return (
    <div ref={boxRef} style={{ position: "relative" }}>
      <button
        onClick={() => setOpen((o) => !o)}
        aria-label="ការជូនដំណឹង"
        style={{
          position: "relative",
          background: "none",
          border: "none",
          cursor: "pointer",
          color: T.ink,
          padding: 6,
          display: "flex",
        }}
      >
        <Bell size={19} />
        {unread.length > 0 && (
          <span
            style={{
              position: "absolute",
              top: 1,
              right: 1,
              minWidth: 16,
              height: 16,
              borderRadius: 999,
              background: T.rose,
              color: "#fff",
              fontSize: 10,
              fontWeight: 700,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              padding: "0 3px",
              border: `2px solid ${T.paper}`,
            }}
          >
            {unread.length > 9 ? "9+" : unread.length}
          </span>
        )}
      </button>
      {open && (
        <div
          className="wf-card"
          style={{
            position: "absolute",
            right: 0,
            top: "calc(100% + 8px)",
            width: 320,
            maxWidth: "calc(100vw - 32px)",
            maxHeight: 400,
            overflowY: "auto",
            zIndex: 30,
            boxShadow: "0 12px 32px rgba(18,32,61,0.18)",
          }}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              padding: "12px 14px",
              borderBottom: `1px solid ${T.lineSoft}`,
            }}
          >
            <span style={{ fontWeight: 700, fontSize: 13, color: T.ink }}>
              ការជូនដំណឹង
            </span>
            {notifications.length > 0 && (
              <button
                onClick={() => markRead(notifications.map((n) => n.id))}
                style={{
                  background: "none",
                  border: "none",
                  cursor: "pointer",
                  color: T.forest,
                  fontSize: 11.5,
                  fontWeight: 600,
                }}
              >
                កំណត់ថាបានអានទាំងអស់
              </button>
            )}
          </div>
          {notifications.length === 0 ? (
            <div
              style={{
                padding: "28px 14px",
                textAlign: "center",
                color: T.muted,
                fontSize: 12.5,
              }}
            >
              មិនមានការជូនដំណឹងទេ
            </div>
          ) : (
            notifications.map((n) => {
              const isUnread = !readIds.includes(n.id);
              return (
                <button
                  key={n.id}
                  onClick={() => {
                    markRead([n.id]);
                    setPage(n.page);
                    setOpen(false);
                  }}
                  style={{
                    display: "flex",
                    gap: 10,
                    width: "100%",
                    textAlign: "left",
                    padding: "11px 14px",
                    background: isUnread ? T.paper : "transparent",
                    border: "none",
                    borderBottom: `1px solid ${T.lineSoft}`,
                    cursor: "pointer",
                  }}
                >
                  <span
                    style={{
                      width: 7,
                      height: 7,
                      borderRadius: 999,
                      background: NOTIF_TONE[n.tone] || T.muted,
                      marginTop: 5,
                      flexShrink: 0,
                    }}
                  />
                  <span style={{ minWidth: 0 }}>
                    <div
                      style={{ fontSize: 12.5, fontWeight: 600, color: T.ink }}
                    >
                      {n.title}
                    </div>
                    <div
                      style={{
                        fontSize: 12,
                        color: T.textSoft,
                        marginTop: 2,
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        display: "-webkit-box",
                        WebkitLineClamp: 2,
                        WebkitBoxOrient: "vertical",
                      }}
                    >
                      {n.message}
                    </div>
                    <div
                      style={{
                        fontSize: 10.5,
                        color: T.mutedLight,
                        marginTop: 3,
                      }}
                    >
                      {timeAgoLabel(n.time)}
                    </div>
                  </span>
                </button>
              );
            })
          )}
        </div>
      )}
    </div>
  );
}
function Modal({ title, onClose, children, width = 480 }) {
  return (
    <div
      className="wf-modal-overlay"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="wf-modal" style={{ maxWidth: width }}>
        <div className="wf-modal-head">
          <h3
            style={{
              fontFamily: "'Space Grotesk',sans-serif",
              fontWeight: 600,
              color: T.ink,
              flex: 1,
            }}
          >
            {title}
          </h3>
          <button
            onClick={onClose}
            style={{
              background: "none",
              border: "none",
              cursor: "pointer",
              color: T.muted,
            }}
          >
            <X size={18} />
          </button>
        </div>
        <div style={{ padding: 20 }}>{children}</div>
      </div>
    </div>
  );
}
function Field({ label, children }) {
  return (
    <label style={{ display: "block", marginBottom: 14 }}>
      <span className="wf-field-label">{label}</span>
      {children}
    </label>
  );
}
function Input(props) {
  return <input className="wf-input" {...props} />;
}
function Select(props) {
  return <select className="wf-input" {...props} />;
}
function ConfirmDialog({ text, onCancel, onConfirm }) {
  return (
    <Modal title="បញ្ជាក់ការលុប" onClose={onCancel} width={380}>
      <p style={{ fontSize: 14, color: T.textSoft, marginBottom: 20 }}>
        {text}
      </p>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Button variant="ghost" onClick={onCancel}>
          បោះបង់
        </Button>
        <Button variant="danger-solid" onClick={onConfirm}>
          <Trash2 size={14} /> លុប
        </Button>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------
   Login screens — two fully separate pages/URLs:
   #/employee  → employee portal, no admin option visible at all
   #/admin     → admin portal, no employee option visible at all
----------------------------------------------------------------*/
function AuthShell({ children }) {
  return (
    <div
      className="wf-root"
      style={{
        alignItems: "center",
        justifyContent: "center",
        minHeight: 640,
        background: T.ink,
      }}
    >
      <div
        style={{
          width: "100%",
          maxWidth: 380,
          padding: "28px 26px",
          background: "#fff",
          borderRadius: 18,
          boxShadow: "0 20px 60px rgba(0,0,0,0.35)",
        }}
      >
        {children}
      </div>
    </div>
  );
}
function BackToChooser({ go }) {
  return (
    <button
      onClick={() => go("employee")}
      style={{
        display: "flex",
        alignItems: "center",
        gap: 6,
        background: "none",
        border: "none",
        cursor: "pointer",
        color: T.muted,
        fontSize: 12,
        fontWeight: 600,
        marginBottom: 16,
        padding: 0,
      }}
    >
      <ArrowLeft size={14} /> ត្រឡប់ក្រោយ
    </button>
  );
}

function EmployeeLoginScreen({ employees, onLogin }) {
  const [code, setCode] = useState("");
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");

  const submit = (e) => {
    e.preventDefault();
    const emp = employees.find(
      (x) => x.code.trim().toLowerCase() === code.trim().toLowerCase(),
    );
    if (!emp) {
      setError("រកមិនឃើញលេខសម្គាល់បុគ្គលិកនេះទេ");
      return;
    }
    if (emp.status !== "active") {
      setError("គណនីនេះមិនទាន់សកម្មទេ សូមទាក់ទង Admin");
      return;
    }
    if ((emp.pin || "") !== pin.trim()) {
      setError("កូដសម្ងាត់ (PIN) មិនត្រឹមត្រូវទេ");
      return;
    }
    setError("");
    onLogin(emp.id);
  };

  return (
    <AuthShell>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          marginBottom: 22,
          justifyContent: "center",
        }}
      >
        <div className="wf-logo-badge">WS</div>
        <span
          style={{
            fontWeight: 700,
            fontSize: 17,
            color: T.ink,
            fontFamily: "'Space Grotesk',sans-serif",
          }}
        >
          Workforce Suite
        </span>
      </div>
      <form onSubmit={submit}>
        <Field label="លេខសម្គាល់បុគ្គលិក (Employee ID)">
          <Input
            value={code}
            onChange={(e) => setCode(e.target.value)}
            placeholder="EMP-001"
            autoFocus
          />
        </Field>
        <Field label="កូដសម្ងាត់ (PIN)">
          <Input
            value={pin}
            onChange={(e) => setPin(e.target.value)}
            placeholder="••••"
            type="password"
            inputMode="numeric"
            maxLength={6}
          />
        </Field>
        {error && (
          <p style={{ fontSize: 12.5, color: T.rose, marginBottom: 10 }}>
            {error}
          </p>
        )}
        <Button
          variant="accent"
          type="submit"
          style={{
            width: "100%",
            justifyContent: "center",
            padding: "11px 0",
            fontSize: 14,
          }}
        >
          <LogIn size={16} /> ចូលប្រើប្រព័ន្ធ
        </Button>
      </form>
      <div
        style={{
          marginTop: 18,
          padding: "10px 12px",
          background: T.paper,
          borderRadius: 10,
          fontSize: 11,
          color: T.muted,
          lineHeight: 1.6,
        }}
      >
        <b style={{ color: T.textSoft }}>សម្រាប់សាកល្បង៖</b>
        <br />
        {employees
          .slice(0, 3)
          .map((e) => `${e.code}/${e.pin}`)
          .join("  ·  ")}
      </div>
    </AuthShell>
  );
}

function AdminLoginScreen({ admins, onLogin, go }) {
  const [username, setUsername] = useState("");
  const [adminPass, setAdminPass] = useState("");
  const [error, setError] = useState("");

  const submit = (e) => {
    e.preventDefault();
    const acct = admins.find(
      (a) => a.username.trim().toLowerCase() === username.trim().toLowerCase(),
    );
    if (!acct) {
      setError("រកមិនឃើញគណនីអ្នកគ្រប់គ្រងនេះទេ");
      return;
    }
    if (acct.password !== adminPass) {
      setError("ពាក្យសម្ងាត់មិនត្រឹមត្រូវទេ");
      return;
    }
    setError("");
    onLogin(acct.id);
  };

  return (
    <AuthShell>
      <BackToChooser go={go} />
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          marginBottom: 20,
          justifyContent: "center",
        }}
      >
        <div
          style={{
            width: 32,
            height: 32,
            borderRadius: 8,
            background: "#EDE7F6",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            flexShrink: 0,
          }}
        >
          <ShieldCheck size={17} color="#4A3B7A" />
        </div>
        <span
          style={{
            fontWeight: 700,
            fontSize: 16,
            color: T.ink,
            fontFamily: "'Space Grotesk',sans-serif",
          }}
        >
          ចូលប្រើសម្រាប់អ្នកគ្រប់គ្រង
        </span>
      </div>
      <form onSubmit={submit}>
        <Field label="ឈ្មោះគណនី (Username)">
          <Input
            value={username}
            onChange={(e) => setUsername(e.target.value)}
            placeholder="admin"
            autoFocus
          />
        </Field>
        <Field label="ពាក្យសម្ងាត់អ្នកគ្រប់គ្រង">
          <Input
            value={adminPass}
            onChange={(e) => setAdminPass(e.target.value)}
            placeholder="••••••••"
            type="password"
          />
        </Field>
        {error && (
          <p style={{ fontSize: 12.5, color: T.rose, marginBottom: 10 }}>
            {error}
          </p>
        )}
        <Button
          variant="primary"
          type="submit"
          style={{
            width: "100%",
            justifyContent: "center",
            padding: "11px 0",
            fontSize: 14,
          }}
        >
          <ShieldCheck size={16} /> ចូលជាអ្នកគ្រប់គ្រង
        </Button>
      </form>
      <div
        style={{
          marginTop: 18,
          padding: "10px 12px",
          background: T.paper,
          borderRadius: 10,
          fontSize: 11,
          color: T.muted,
          lineHeight: 1.6,
        }}
      >
        <b style={{ color: T.textSoft }}>សម្រាប់សាកល្បង៖</b>
        <br />
        {admins.map((a) => `${a.username}/${a.password}`).join("  ·  ")}
      </div>
    </AuthShell>
  );
}

/* ---------------------------------------------------------------
   Dashboard
----------------------------------------------------------------*/
/* ---------------------------------------------------------------
   Employee portal link — admin can copy/QR the link that goes
   straight to the employee login (no admin option shown there).
----------------------------------------------------------------*/
function employeePortalUrl() {
  const { origin, pathname } = window.location;
  return `${origin}${pathname}#/employee`;
}
function QrModal({ url, onClose }) {
  const qrSrc = `https://api.qrserver.com/v1/create-qr-code/?size=240x240&data=${encodeURIComponent(url)}`;
  return (
    <Modal title="QR Code សម្រាប់បុគ្គលិក" onClose={onClose} width={340}>
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          gap: 14,
        }}
      >
        <div
          style={{
            padding: 12,
            background: "#fff",
            border: `1px solid ${T.lineSoft}`,
            borderRadius: 12,
          }}
        >
          <img
            src={qrSrc}
            alt="Employee portal QR code"
            width={220}
            height={220}
          />
        </div>
        <p style={{ fontSize: 12, color: T.muted, textAlign: "center" }}>
          ស្កេនដើម្បីបើកទំព័រចូលប្រើសម្រាប់បុគ្គលិកដោយផ្ទាល់
        </p>
      </div>
    </Modal>
  );
}
function EmployeeLinkCard() {
  const [copied, setCopied] = useState(false);
  const [showQr, setShowQr] = useState(false);
  const url = employeePortalUrl();

  const copyLink = async () => {
    try {
      await navigator.clipboard.writeText(url);
    } catch {
      // Clipboard API unavailable (e.g. insecure context) — fall back silently.
    }
    setCopied(true);
    setTimeout(() => setCopied(false), 1800);
  };

  return (
    <Card
      style={{
        padding: "16px 18px",
        marginBottom: 22,
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        gap: 16,
        flexWrap: "wrap",
      }}
    >
      <div
        style={{ display: "flex", alignItems: "center", gap: 12, minWidth: 0 }}
      >
        <div
          style={{
            width: 38,
            height: 38,
            borderRadius: 10,
            background: T.forestSoft,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            flexShrink: 0,
          }}
        >
          <Store size={18} color={T.forestText} />
        </div>
        <div style={{ minWidth: 0 }}>
          <div style={{ fontSize: 12, color: T.muted }}>
            តំណភ្ជាប់សម្រាប់បុគ្គលិក
          </div>
          <div
            style={{
              fontSize: 13,
              fontWeight: 600,
              color: T.ink,
              fontFamily: "'JetBrains Mono',monospace",
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
            }}
          >
            {url}
          </div>
        </div>
      </div>
      <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
        <Button variant="ghost" onClick={() => setShowQr(true)}>
          <QrCode size={14} /> បង្ហាញ QR Code
        </Button>
        <Button variant="accent" onClick={copyLink}>
          {copied ? <CheckCircle2 size={14} /> : <Copy size={14} />}{" "}
          {copied ? "បានចម្លង!" : "ចម្លងតំណ"}
        </Button>
      </div>
      {showQr && <QrModal url={url} onClose={() => setShowQr(false)} />}
    </Card>
  );
}

function Dashboard({
  employees,
  departments,
  attendance,
  payrollPaid,
  role,
  currentEmp,
}) {
  const today = todayStr();
  const activeEmployees = employees.filter((e) => e.status === "active");
  const presentToday = attendance.filter((a) => a.date === today).length;
  const mk = monthKey();
  const paidCount = activeEmployees.filter(
    (e) => payrollPaid[`${e.id}-${mk}`],
  ).length;
  const pendingPayroll = activeEmployees.length - paidCount;
  const rate = activeEmployees.length
    ? Math.round((presentToday / activeEmployees.length) * 100)
    : 0;
  const stats = [
    {
      label: "បុគ្គលិកសរុប",
      value: employees.length,
      sub: `${activeEmployees.length} សកម្ម`,
      icon: Users,
      accent: T.forest,
    },
    {
      label: "នាយកដ្ឋាន",
      value: departments.length,
      sub: "នាយកដ្ឋានសរុប",
      icon: Building2,
      accent: T.blue,
    },
    {
      label: "មកធ្វើការថ្ងៃនេះ",
      value: presentToday,
      sub: `${rate}% អត្រាមកធ្វើការ`,
      icon: Clock,
      accent: T.gold,
    },
    {
      label: "រង់ចាំបើកប្រាក់ខែ",
      value: pendingPayroll,
      sub: "ខែនេះ",
      icon: Wallet,
      accent: T.rose,
    },
  ];
  const recent = [...attendance]
    .filter((a) => a.date === today)
    .slice(-5)
    .reverse();

  return (
    <div>
      <Card
        style={{
          padding: 20,
          marginBottom: 22,
          background: T.ink,
          color: "#fff",
          position: "relative",
          overflow: "hidden",
        }}
      >
        <p style={{ color: "#A9B4C7", fontSize: 13 }}>សូមអញ្ជើញ</p>
        <h2
          style={{
            fontFamily: "'Space Grotesk',sans-serif",
            fontSize: 24,
            fontWeight: 600,
            marginTop: 2,
          }}
        >
          {role === "admin" ? "អ្នកគ្រប់គ្រងប្រព័ន្ធ" : currentEmp?.name}
        </h2>
        <p
          style={{
            color: "#A9B4C7",
            fontSize: 12,
            marginTop: 6,
            fontFamily: "'JetBrains Mono',monospace",
          }}
        >
          {new Date().toLocaleDateString("km-KH", {
            weekday: "long",
            year: "numeric",
            month: "long",
            day: "numeric",
          })}
        </p>
      </Card>

      {role === "admin" && <EmployeeLinkCard />}

      <div
        className="wf-grid"
        style={{
          gridTemplateColumns: "repeat(auto-fit, minmax(190px,1fr))",
          marginBottom: 22,
        }}
      >
        {stats.map((s) => (
          <Card key={s.label} accent={s.accent} style={{ padding: 16 }}>
            <div
              style={{
                width: 36,
                height: 36,
                borderRadius: 10,
                background: s.accent + "1A",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                marginBottom: 12,
              }}
            >
              <s.icon size={18} color={s.accent} />
            </div>
            <div
              style={{
                fontSize: 26,
                fontWeight: 700,
                fontFamily: "'JetBrains Mono',monospace",
                color: T.ink,
              }}
            >
              {s.value}
            </div>
            <div
              style={{
                fontSize: 13,
                fontWeight: 600,
                color: T.textSoft,
                marginTop: 2,
              }}
            >
              {s.label}
            </div>
            <div style={{ fontSize: 11, color: T.muted, marginTop: 2 }}>
              {s.sub}
            </div>
          </Card>
        ))}
      </div>

      <Card style={{ padding: 18 }}>
        <h3
          style={{
            fontFamily: "'Space Grotesk',sans-serif",
            fontWeight: 600,
            color: T.ink,
            marginBottom: 12,
            fontSize: 14,
          }}
        >
          ការចូលធ្វើការថ្ងៃនេះ
        </h3>
        {recent.length === 0 ? (
          <p
            style={{
              fontSize: 13,
              color: T.muted,
              textAlign: "center",
              padding: "28px 0",
            }}
          >
            មិនទាន់មានការចុះឈ្មោះចូលធ្វើការថ្ងៃនេះទេ
          </p>
        ) : (
          <div>
            {recent.map((a) => {
              const emp = employees.find((e) => e.id === a.employeeId);
              return (
                <div
                  key={a.id}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 12,
                    padding: "10px 0",
                    borderTop: `1px solid #F0EDE2`,
                  }}
                >
                  <Avatar
                    name={emp?.name || "?"}
                    photo={emp?.photo}
                    size={32}
                  />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div
                      style={{ fontSize: 13, fontWeight: 500, color: T.ink }}
                    >
                      {emp?.name}
                    </div>
                    <div
                      style={{
                        fontSize: 11,
                        color: T.muted,
                        fontFamily: "'JetBrains Mono',monospace",
                      }}
                    >
                      ចូល {a.checkIn || "—"}{" "}
                      {a.checkOut ? `· ចេញ ${a.checkOut}` : ""}
                    </div>
                  </div>
                  <StatusPill status={a.status} />
                </div>
              );
            })}
          </div>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------
   Employees
----------------------------------------------------------------*/
function EmployeeForm({ initial, departments, shifts, onSave, onCancel }) {
  const [f, setF] = useState(
    initial || {
      code: "",
      pin: randomPin(),
      name: "",
      deptId: departments[0]?.id || "",
      shiftId: shifts[0]?.id || "",
      role: "",
      email: "",
      phone: "",
      salary: "",
      status: "active",
      joined: todayStr(),
    },
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <div>
      <Field label="ឈ្មោះពេញ">
        <Input
          value={f.name}
          onChange={set("name")}
          placeholder="ឧ. លោក សុវណ្ណ ដារា"
        />
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="លេខសម្គាល់ (ប្រើសម្រាប់ Login)">
          <Input value={f.code} onChange={set("code")} placeholder="EMP-004" />
        </Field>
        <Field label="តួនាទី">
          <Input
            value={f.role}
            onChange={set("role")}
            placeholder="ឧ. Accountant"
          />
        </Field>
      </div>
      <Field label="កូដសម្ងាត់ PIN (សម្រាប់ Login)">
        <div style={{ display: "flex", gap: 8 }}>
          <Input
            value={f.pin}
            onChange={set("pin")}
            placeholder="1234"
            maxLength={6}
            style={{ flex: 1 }}
          />
          <Button
            type="button"
            variant="ghost"
            size="sm"
            onClick={() => setF({ ...f, pin: randomPin() })}
          >
            <KeyRound size={13} /> បង្កើតថ្មី
          </Button>
        </div>
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="នាយកដ្ឋាន">
          <Select value={f.deptId} onChange={set("deptId")}>
            {departments.map((d) => (
              <option key={d.id} value={d.id}>
                {d.name}
              </option>
            ))}
          </Select>
        </Field>
        <Field label="វេនធ្វើការ">
          <Select value={f.shiftId} onChange={set("shiftId")}>
            {shifts.map((s) => (
              <option key={s.id} value={s.id}>
                {shiftLabel(s)}
              </option>
            ))}
          </Select>
        </Field>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="អ៊ីមែល">
          <Input
            value={f.email}
            onChange={set("email")}
            placeholder="name@company.com"
          />
        </Field>
        <Field label="លេខទូរស័ព្ទ">
          <Input
            value={f.phone}
            onChange={set("phone")}
            placeholder="012 345 678"
          />
        </Field>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="ប្រាក់ខែ (USD)">
          <Input
            type="number"
            value={f.salary}
            onChange={set("salary")}
            placeholder="600"
          />
        </Field>
        <Field label="ស្ថានភាព">
          <Select value={f.status} onChange={set("status")}>
            <option value="active">សកម្ម</option>
            <option value="inactive">អសកម្ម</option>
          </Select>
        </Field>
      </div>
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          gap: 8,
          marginTop: 16,
          paddingTop: 14,
          borderTop: `1px solid ${T.lineSoft}`,
        }}
      >
        <Button variant="ghost" onClick={onCancel}>
          បោះបង់
        </Button>
        <Button
          variant="accent"
          onClick={() => onSave(f)}
          disabled={!f.name || !f.code || !f.pin}
        >
          រក្សាទុក
        </Button>
      </div>
    </div>
  );
}

function Employees({
  employees,
  departments,
  shifts,
  setEmployees,
  isSuperAdmin,
}) {
  const [modal, setModal] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const [query, setQuery] = useState("");
  const filtered = employees.filter((e) =>
    (e.name + e.code + e.role).toLowerCase().includes(query.toLowerCase()),
  );
  const deptName = (id) => departments.find((d) => d.id === id)?.name || "—";
  const shiftOf = (id) => shifts.find((s) => s.id === id);

  const save = (data) => {
    if (modal.mode === "add")
      setEmployees([
        ...employees,
        { ...data, id: uid("e"), salary: Number(data.salary) || 0 },
      ]);
    else
      setEmployees(
        employees.map((e) =>
          e.id === data.id ? { ...data, salary: Number(data.salary) || 0 } : e,
        ),
      );
    setModal(null);
  };

  return (
    <div>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: 12,
          marginBottom: 16,
          flexWrap: "wrap",
        }}
      >
        <div
          style={{
            position: "relative",
            flex: 1,
            minWidth: 200,
            maxWidth: 320,
          }}
        >
          <Search
            size={15}
            style={{
              position: "absolute",
              left: 12,
              top: "50%",
              transform: "translateY(-50%)",
              color: T.muted,
            }}
          />
          <Input
            style={{ paddingLeft: 34 }}
            placeholder="ស្វែងរកបុគ្គលិក..."
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
        <Button variant="accent" onClick={() => setModal({ mode: "add" })}>
          <Plus size={15} /> បន្ថែមបុគ្គលិក
        </Button>
      </div>

      <div
        className="wf-grid"
        style={{ gridTemplateColumns: "repeat(auto-fill, minmax(260px,1fr))" }}
      >
        {filtered.map((e) => (
          <Card key={e.id} accent={colorFor(e.name)} style={{ padding: 16 }}>
            <div
              style={{
                display: "flex",
                alignItems: "flex-start",
                justifyContent: "space-between",
                marginBottom: 12,
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <Avatar name={e.name} photo={e.photo} />
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontWeight: 600, color: T.ink, fontSize: 13 }}>
                    {e.name}
                  </div>
                  <div
                    style={{
                      fontSize: 11,
                      color: T.muted,
                      fontFamily: "'JetBrains Mono',monospace",
                    }}
                  >
                    {e.code}
                  </div>
                </div>
              </div>
              <StatusPill status={e.status} />
            </div>
            <div
              style={{
                fontSize: 12,
                color: T.textSoft,
                marginBottom: 12,
                display: "flex",
                flexDirection: "column",
                gap: 4,
              }}
            >
              <div>
                🏢 {deptName(e.deptId)} · {e.role}
              </div>
              <div>⏰ {shiftLabel(shiftOf(e.shiftId))}</div>
              <div
                style={{
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                  whiteSpace: "nowrap",
                }}
              >
                ✉ {e.email}
              </div>
              <div>☎ {e.phone}</div>
              <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                <Lock size={11} color={T.muted} />
                <span
                  style={{
                    fontFamily: "'JetBrains Mono',monospace",
                    color: T.muted,
                  }}
                >
                  PIN: {e.pin || "—"}
                </span>
              </div>
              <div
                style={{
                  fontWeight: 700,
                  color: T.forestText,
                  fontFamily: "'JetBrains Mono',monospace",
                }}
              >
                {fmtMoney(e.salary)} / ខែ
              </div>
            </div>
            <div
              style={{
                display: "flex",
                gap: 8,
                paddingTop: 10,
                borderTop: `1px solid #F0EDE2`,
              }}
            >
              <Button
                size="sm"
                variant="ghost"
                style={{ flex: 1, justifyContent: "center" }}
                onClick={() => setModal({ mode: "edit", data: e })}
              >
                <Pencil size={13} /> កែសម្រួល
              </Button>
              {isSuperAdmin && (
                <Button
                  size="sm"
                  variant="danger"
                  style={{ flex: 1, justifyContent: "center" }}
                  onClick={() => setConfirmDel(e)}
                >
                  <Trash2 size={13} /> លុប
                </Button>
              )}
            </div>
          </Card>
        ))}
        {filtered.length === 0 && (
          <p
            style={{
              gridColumn: "1/-1",
              textAlign: "center",
              color: T.muted,
              fontSize: 13,
              padding: "40px 0",
            }}
          >
            មិនមានលទ្ធផលទេ
          </p>
        )}
      </div>

      {modal && (
        <Modal
          title={
            modal.mode === "add"
              ? "បន្ថែមបុគ្គលិកថ្មី"
              : "កែសម្រួលព័ត៌មានបុគ្គលិក"
          }
          onClose={() => setModal(null)}
        >
          <EmployeeForm
            initial={modal.data}
            departments={departments}
            shifts={shifts}
            onSave={save}
            onCancel={() => setModal(null)}
          />
        </Modal>
      )}
      {confirmDel && (
        <ConfirmDialog
          text={`តើអ្នកពិតជាចង់លុប "${confirmDel.name}" មែនទេ?`}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setEmployees(employees.filter((e) => e.id !== confirmDel.id));
            setConfirmDel(null);
          }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Departments
----------------------------------------------------------------*/
function DeptForm({ initial, onSave, onCancel }) {
  const [f, setF] = useState(initial || { name: "", code: "", desc: "" });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <div>
      <Field label="ឈ្មោះនាយកដ្ឋាន">
        <Input
          value={f.name}
          onChange={set("name")}
          placeholder="ឧ. ធនធានមនុស្ស"
        />
      </Field>
      <Field label="កូដ">
        <Input value={f.code} onChange={set("code")} placeholder="ឧ. HR" />
      </Field>
      <Field label="ការពិពណ៌នា">
        <textarea
          className="wf-input"
          rows={2}
          value={f.desc}
          onChange={set("desc")}
        />
      </Field>
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          gap: 8,
          marginTop: 16,
          paddingTop: 14,
          borderTop: `1px solid ${T.lineSoft}`,
        }}
      >
        <Button variant="ghost" onClick={onCancel}>
          បោះបង់
        </Button>
        <Button variant="accent" onClick={() => onSave(f)} disabled={!f.name}>
          រក្សាទុក
        </Button>
      </div>
    </div>
  );
}

function Departments({ departments, setDepartments, employees, isSuperAdmin }) {
  const [modal, setModal] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const countIn = (id) => employees.filter((e) => e.deptId === id).length;
  const save = (data) => {
    if (modal.mode === "add")
      setDepartments([...departments, { ...data, id: uid("d") }]);
    else setDepartments(departments.map((d) => (d.id === data.id ? data : d)));
    setModal(null);
  };
  return (
    <div>
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          marginBottom: 16,
        }}
      >
        <Button variant="accent" onClick={() => setModal({ mode: "add" })}>
          <Plus size={15} /> បន្ថែមនាយកដ្ឋាន
        </Button>
      </div>
      <div
        className="wf-grid"
        style={{ gridTemplateColumns: "repeat(auto-fill, minmax(280px,1fr))" }}
      >
        {departments.map((d) => (
          <Card key={d.id} accent={colorFor(d.name)} style={{ padding: 16 }}>
            <div
              style={{
                display: "flex",
                alignItems: "flex-start",
                justifyContent: "space-between",
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div
                  style={{
                    width: 38,
                    height: 38,
                    borderRadius: 10,
                    background: T.paper,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                  }}
                >
                  <Building2 size={18} color={T.ink} />
                </div>
                <div>
                  <div style={{ fontWeight: 600, color: T.ink, fontSize: 13 }}>
                    {d.name}
                  </div>
                  <div
                    style={{
                      fontSize: 11,
                      color: T.muted,
                      fontFamily: "'JetBrains Mono',monospace",
                    }}
                  >
                    {d.code}
                  </div>
                </div>
              </div>
              <span
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  color: T.textSoft,
                  background: T.paper,
                  padding: "5px 9px",
                  borderRadius: 8,
                }}
              >
                {countIn(d.id)} នាក់
              </span>
            </div>
            <p style={{ fontSize: 12, color: T.muted, marginTop: 12 }}>
              {d.desc}
            </p>
            <div
              style={{
                display: "flex",
                gap: 8,
                paddingTop: 12,
                marginTop: 12,
                borderTop: "1px solid #F0EDE2",
              }}
            >
              <Button
                size="sm"
                variant="ghost"
                style={{ flex: 1, justifyContent: "center" }}
                onClick={() => setModal({ mode: "edit", data: d })}
              >
                <Pencil size={13} /> កែសម្រួល
              </Button>
              {isSuperAdmin && (
                <Button
                  size="sm"
                  variant="danger"
                  style={{ flex: 1, justifyContent: "center" }}
                  onClick={() => setConfirmDel(d)}
                >
                  <Trash2 size={13} /> លុប
                </Button>
              )}
            </div>
          </Card>
        ))}
      </div>
      {modal && (
        <Modal
          title={
            modal.mode === "add" ? "បន្ថែមនាយកដ្ឋានថ្មី" : "កែសម្រួលនាយកដ្ឋាន"
          }
          onClose={() => setModal(null)}
        >
          <DeptForm
            initial={modal.data}
            onSave={save}
            onCancel={() => setModal(null)}
          />
        </Modal>
      )}
      {confirmDel && (
        <ConfirmDialog
          text={`តើអ្នកពិតជាចង់លុបនាយកដ្ឋាន "${confirmDel.name}" មែនទេ? (មានបុគ្គលិក ${countIn(confirmDel.id)} នាក់)`}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setDepartments(departments.filter((d) => d.id !== confirmDel.id));
            setConfirmDel(null);
          }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Shifts (admin-configurable work shifts, e.g. morning/afternoon/night)
----------------------------------------------------------------*/
function ShiftForm({ initial, onSave, onCancel }) {
  const [f, setF] = useState(
    initial || { name: "", start: "06:00", end: "14:00" },
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <div>
      <Field label="ឈ្មោះវេន">
        <Input
          value={f.name}
          onChange={set("name")}
          placeholder="ឧ. វេនព្រឹក"
        />
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="ម៉ោងចូល">
          <Input type="time" value={f.start} onChange={set("start")} />
        </Field>
        <Field label="ម៉ោងចេញ">
          <Input type="time" value={f.end} onChange={set("end")} />
        </Field>
      </div>
      {f.end <= f.start && f.start && f.end && (
        <p
          style={{
            fontSize: 11.5,
            color: T.muted,
            marginTop: -6,
            marginBottom: 14,
          }}
        >
          ⓘ វេននេះលើសពាក់កណ្តាលអធ្រាត្រ (ឧ. ចូលយប់ ចេញព្រឹក)
        </p>
      )}
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          gap: 8,
          marginTop: 16,
          paddingTop: 14,
          borderTop: `1px solid ${T.lineSoft}`,
        }}
      >
        <Button variant="ghost" onClick={onCancel}>
          បោះបង់
        </Button>
        <Button
          variant="accent"
          onClick={() => onSave(f)}
          disabled={!f.name || !f.start || !f.end}
        >
          រក្សាទុក
        </Button>
      </div>
    </div>
  );
}

function Shifts({ shifts, setShifts, employees, isSuperAdmin }) {
  const [modal, setModal] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const countIn = (id) => employees.filter((e) => e.shiftId === id).length;
  const save = (data) => {
    if (modal.mode === "add") setShifts([...shifts, { ...data, id: uid("s") }]);
    else setShifts(shifts.map((s) => (s.id === data.id ? data : s)));
    setModal(null);
  };
  return (
    <div>
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          marginBottom: 16,
        }}
      >
        <Button variant="accent" onClick={() => setModal({ mode: "add" })}>
          <Plus size={15} /> បន្ថែមវេន
        </Button>
      </div>
      <div
        className="wf-grid"
        style={{ gridTemplateColumns: "repeat(auto-fill, minmax(260px,1fr))" }}
      >
        {shifts.map((s) => (
          <Card key={s.id} accent={colorFor(s.name)} style={{ padding: 16 }}>
            <div
              style={{
                display: "flex",
                alignItems: "flex-start",
                justifyContent: "space-between",
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div
                  style={{
                    width: 38,
                    height: 38,
                    borderRadius: 10,
                    background: T.paper,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                  }}
                >
                  <Watch size={18} color={T.ink} />
                </div>
                <div>
                  <div style={{ fontWeight: 600, color: T.ink, fontSize: 13 }}>
                    {s.name}
                  </div>
                  <div
                    style={{
                      fontSize: 11,
                      color: T.muted,
                      fontFamily: "'JetBrains Mono',monospace",
                    }}
                  >
                    {s.start} – {s.end}
                    {isOvernightShift(s) ? " (ឆ្លងអធ្រាត្រ)" : ""}
                  </div>
                </div>
              </div>
              <span
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  color: T.textSoft,
                  background: T.paper,
                  padding: "5px 9px",
                  borderRadius: 8,
                }}
              >
                {countIn(s.id)} នាក់
              </span>
            </div>
            <div
              style={{
                display: "flex",
                gap: 8,
                paddingTop: 12,
                marginTop: 12,
                borderTop: "1px solid #F0EDE2",
              }}
            >
              <Button
                size="sm"
                variant="ghost"
                style={{ flex: 1, justifyContent: "center" }}
                onClick={() => setModal({ mode: "edit", data: s })}
              >
                <Pencil size={13} /> កែសម្រួល
              </Button>
              {isSuperAdmin && (
                <Button
                  size="sm"
                  variant="danger"
                  style={{ flex: 1, justifyContent: "center" }}
                  onClick={() => setConfirmDel(s)}
                >
                  <Trash2 size={13} /> លុប
                </Button>
              )}
            </div>
          </Card>
        ))}
      </div>
      {modal && (
        <Modal
          title={modal.mode === "add" ? "បន្ថែមវេនថ្មី" : "កែសម្រួលវេន"}
          onClose={() => setModal(null)}
        >
          <ShiftForm
            initial={modal.data}
            onSave={save}
            onCancel={() => setModal(null)}
          />
        </Modal>
      )}
      {confirmDel && (
        <ConfirmDialog
          text={`តើអ្នកពិតជាចង់លុបវេន "${confirmDel.name}" មែនទេ? (មានបុគ្គលិក ${countIn(confirmDel.id)} នាក់កំពុងប្រើវេននេះ)`}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setShifts(shifts.filter((s) => s.id !== confirmDel.id));
            setConfirmDel(null);
          }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Attendance
----------------------------------------------------------------*/
function SelfPunch({ emp, shift, attendance, setAttendance, officeLocation }) {
  const today = todayStr();
  const rec = attendance.find(
    (a) => a.employeeId === emp.id && a.date === today,
  );
  const [locBusy, setLocBusy] = useState(false);
  const [locError, setLocError] = useState("");

  // If an office location + radius is configured, require the employee's
  // current GPS position to fall within it before allowing a punch.
  // Returns the coords to attach to the attendance record, or null if the
  // punch should be blocked (locError is set in that case).
  const verifyLocation = async () => {
    if (!officeLocation) return null; // no geofence configured — skip check
    setLocError("");
    setLocBusy(true);
    try {
      const coords = await getCurrentPosition();
      const dist = distanceMeters(
        coords.latitude,
        coords.longitude,
        officeLocation.lat,
        officeLocation.lng,
      );
      if (dist > officeLocation.radius) {
        setLocError(
          `អ្នកនៅឆ្ងាយពីការិយាល័យ ${Math.round(dist)}m (កំណត់អនុញ្ញាត ${officeLocation.radius}m) — មិនអាចចុះឈ្មោះបានទេ`,
        );
        return null;
      }
      return {
        lat: coords.latitude,
        lng: coords.longitude,
        distance: Math.round(dist),
      };
    } catch {
      setLocError(
        "មិនអាចទាញយកទីតាំង GPS បានទេ សូមបើក Location សម្រាប់កម្មវិធីនេះ",
      );
      return null;
    } finally {
      setLocBusy(false);
    }
  };

  const punchIn = async () => {
    let loc = null;
    if (officeLocation) {
      loc = await verifyLocation();
      if (!loc) return;
    }
    const t = timeNow();
    const status = isLateForShift(t, shift) ? "late" : "present";
    setAttendance([
      ...attendance,
      {
        id: uid("a"),
        employeeId: emp.id,
        date: today,
        checkIn: t,
        checkOut: null,
        status,
        checkInLoc: loc,
      },
    ]);
  };
  const punchOut = async () => {
    let loc = null;
    if (officeLocation) {
      loc = await verifyLocation();
      if (!loc) return;
    }
    setAttendance(
      attendance.map((a) =>
        a.id === rec.id ? { ...a, checkOut: timeNow(), checkOutLoc: loc } : a,
      ),
    );
  };

  return (
    <Card
      accent={T.forest}
      style={{ padding: 26, marginBottom: 22, textAlign: "center" }}
    >
      <div
        style={{
          fontSize: 11,
          fontWeight: 700,
          letterSpacing: ".08em",
          color: T.muted,
          textTransform: "uppercase",
          marginBottom: 4,
        }}
      >
        ម៉ោងបច្ចុប្បន្ន
      </div>
      <div className="wf-punch-clock" style={{ marginBottom: 8 }}>
        {timeNow()}
      </div>
      {shift && (
        <div
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 5,
            fontSize: 11.5,
            color: T.muted,
            background: T.paper,
            padding: "4px 10px",
            borderRadius: 999,
            marginBottom: 16,
          }}
        >
          <Watch size={12} /> {shiftLabel(shift)}
        </div>
      )}
      {!shift && <div style={{ marginBottom: 16 }} />}
      {officeLocation && (
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            gap: 5,
            fontSize: 11,
            color: T.muted,
            marginBottom: 12,
          }}
        >
          <MapPin size={12} /> ត្រូវការទីតាំង GPS ក្នុងចម្ងាយ{" "}
          {officeLocation.radius}m ពីការិយាល័យ
        </div>
      )}
      {locError && (
        <p
          style={{
            fontSize: 12.5,
            color: T.rose,
            marginBottom: 12,
            maxWidth: 320,
            marginLeft: "auto",
            marginRight: "auto",
          }}
        >
          {locError}
        </p>
      )}
      {!rec && (
        <Button
          variant="accent"
          onClick={punchIn}
          disabled={locBusy}
          style={{ padding: "12px 26px", fontSize: 15 }}
        >
          {locBusy ? (
            <Loader2
              size={18}
              style={{ animation: "spin 1s linear infinite" }}
            />
          ) : (
            <LogIn size={18} />
          )}{" "}
          ចុចចូលធ្វើការ
        </Button>
      )}
      {rec && !rec.checkOut && (
        <div>
          <p
            style={{
              fontSize: 13,
              color: T.textSoft,
              marginBottom: 12,
              fontFamily: "'JetBrains Mono',monospace",
            }}
          >
            ចូលធ្វើការម៉ោង {rec.checkIn} · <StatusPill status={rec.status} />
          </p>
          <Button
            variant="danger-solid"
            onClick={punchOut}
            disabled={locBusy}
            style={{ padding: "12px 26px", fontSize: 15 }}
          >
            {locBusy ? (
              <Loader2
                size={18}
                style={{ animation: "spin 1s linear infinite" }}
              />
            ) : (
              <LogOut size={18} />
            )}{" "}
            ចុចចេញពីធ្វើការ
          </Button>
        </div>
      )}
      {rec && rec.checkOut && (
        <div
          style={{
            fontSize: 13,
            color: T.textSoft,
            fontFamily: "'JetBrains Mono',monospace",
          }}
        >
          <CheckCircle2
            size={22}
            color={T.forest}
            style={{ margin: "0 auto 8px" }}
          />
          ថ្ងៃនេះបានបញ្ចប់ការងារ · ចូល {rec.checkIn} · ចេញ {rec.checkOut}
        </div>
      )}
    </Card>
  );
}

/* ---------------------------------------------------------------
   Office location (geofence) settings — admin-only. When configured,
   employees must be within `radius` meters of (lat, lng) to punch in
   or out from the self-service clock.
----------------------------------------------------------------*/
function OfficeLocationSettings({ officeLocation, setOfficeLocation }) {
  const [open, setOpen] = useState(false);
  const [f, setF] = useState(
    officeLocation || { lat: "", lng: "", radius: 150 },
  );
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });

  const useCurrentLocation = async () => {
    setError("");
    setBusy(true);
    try {
      const coords = await getCurrentPosition();
      setF({
        ...f,
        lat: coords.latitude.toFixed(6),
        lng: coords.longitude.toFixed(6),
      });
    } catch {
      setError("មិនអាចទាញយកទីតាំង GPS បច្ចុប្បន្នបានទេ");
    } finally {
      setBusy(false);
    }
  };

  const save = () => {
    const lat = Number(f.lat);
    const lng = Number(f.lng);
    const radius = Number(f.radius);
    if (!lat || !lng || !radius) {
      setError("សូមបំពេញកូអរដោនេ និងកាំឲ្យត្រឹមត្រូវ");
      return;
    }
    setError("");
    setOfficeLocation({ lat, lng, radius });
    setOpen(false);
  };

  return (
    <Card style={{ padding: 16, marginBottom: 16 }}>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: 12,
          cursor: "pointer",
        }}
        onClick={() => setOpen(!open)}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <MapPin size={16} color={T.forest} />
          <span style={{ fontWeight: 600, fontSize: 13.5, color: T.ink }}>
            ការការពារទីតាំង GPS សម្រាប់ Check-in
          </span>
        </div>
        {officeLocation ? (
          <span
            style={{
              fontSize: 11.5,
              color: T.textSoft,
              fontFamily: "'JetBrains Mono',monospace",
            }}
          >
            {officeLocation.lat.toFixed(4)}, {officeLocation.lng.toFixed(4)} ·{" "}
            {officeLocation.radius}m
          </span>
        ) : (
          <span style={{ fontSize: 11.5, color: T.mutedLight }}>
            មិនទាន់កំណត់
          </span>
        )}
      </div>
      {open && (
        <div
          style={{
            marginTop: 16,
            paddingTop: 16,
            borderTop: `1px solid ${T.lineSoft}`,
          }}
        >
          <p style={{ fontSize: 12, color: T.muted, marginBottom: 14 }}>
            កំណត់ទីតាំងការិយាល័យ ដើម្បីតម្រូវឲ្យបុគ្គលិកនៅជិតកន្លែងធ្វើការ
            ពេលចុច check-in/check-out ដោយខ្លួនឯង។
          </p>
          <div
            style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}
          >
            <Field label="Latitude">
              <Input
                value={f.lat}
                onChange={set("lat")}
                placeholder="11.5564"
              />
            </Field>
            <Field label="Longitude">
              <Input
                value={f.lng}
                onChange={set("lng")}
                placeholder="104.9282"
              />
            </Field>
          </div>
          <Field label="កាំអនុញ្ញាត (ម៉ែត្រ)">
            <Input
              type="number"
              value={f.radius}
              onChange={set("radius")}
              placeholder="150"
            />
          </Field>
          {error && (
            <p style={{ fontSize: 12.5, color: T.rose, marginBottom: 10 }}>
              {error}
            </p>
          )}
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              gap: 8,
              marginTop: 6,
            }}
          >
            <div style={{ display: "flex", gap: 8 }}>
              <Button
                variant="ghost"
                size="sm"
                onClick={useCurrentLocation}
                disabled={busy}
              >
                <MapPin size={13} /> ប្រើទីតាំងបច្ចុប្បន្ន
              </Button>
              {officeLocation && (
                <Button
                  variant="danger"
                  size="sm"
                  onClick={() => {
                    setOfficeLocation(null);
                    setF({ lat: "", lng: "", radius: 150 });
                  }}
                >
                  បិទការការពារទីតាំង
                </Button>
              )}
            </div>
            <Button variant="accent" size="sm" onClick={save}>
              រក្សាទុក
            </Button>
          </div>
        </div>
      )}
    </Card>
  );
}

function ManualAttendanceForm({ employees, onSave, onCancel }) {
  const [f, setF] = useState({
    employeeId: employees[0]?.id || "",
    date: todayStr(),
    checkIn: "08:00",
    checkOut: "",
    status: "present",
  });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <div>
      <Field label="បុគ្គលិក">
        <Select value={f.employeeId} onChange={set("employeeId")}>
          {employees.map((e) => (
            <option key={e.id} value={e.id}>
              {e.name} ({e.code})
            </option>
          ))}
        </Select>
      </Field>
      <Field label="កាលបរិច្ឆេទ">
        <Input type="date" value={f.date} onChange={set("date")} />
      </Field>
      <Field label="ស្ថានភាព">
        <Select
          value={f.status}
          onChange={(e) => {
            const status = e.target.value;
            const noTime = status === "absent" || status === "leave";
            setF({
              ...f,
              status,
              checkIn: noTime ? "" : f.checkIn || "08:00",
              checkOut: noTime ? "" : f.checkOut,
            });
          }}
        >
          <option value="present">មកធ្វើការ</option>
          <option value="late">មកយឺត</option>
          <option value="absent">អវត្តមាន (មិនបានអនុញ្ញាត)</option>
          <option value="leave">ឈប់សម្រាក (មានប្រាក់ខែ)</option>
        </Select>
      </Field>
      {f.status !== "absent" && f.status !== "leave" && (
        <div
          style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}
        >
          <Field label="ម៉ោងចូល">
            <Input type="time" value={f.checkIn} onChange={set("checkIn")} />
          </Field>
          <Field label="ម៉ោងចេញ">
            <Input type="time" value={f.checkOut} onChange={set("checkOut")} />
          </Field>
        </div>
      )}
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          gap: 8,
          marginTop: 16,
          paddingTop: 14,
          borderTop: `1px solid ${T.lineSoft}`,
        }}
      >
        <Button variant="ghost" onClick={onCancel}>
          បោះបង់
        </Button>
        <Button
          variant="accent"
          onClick={() => onSave(f)}
          disabled={!f.employeeId}
        >
          រក្សាទុក
        </Button>
      </div>
    </div>
  );
}

function Attendance({
  role,
  currentEmp,
  employees,
  shifts,
  attendance,
  setAttendance,
  isSuperAdmin,
  officeLocation,
  setOfficeLocation,
}) {
  const [date, setDate] = useState(todayStr());
  const [modal, setModal] = useState(false);
  const [confirmDel, setConfirmDel] = useState(null);
  const dayRecords = attendance.filter((a) => a.date === date);
  const activeEmployees = employees.filter((e) => e.status === "active");
  const shiftOf = (id) => shifts.find((s) => s.id === id);
  const rows = activeEmployees.map((e) => ({
    emp: e,
    rec: dayRecords.find((a) => a.employeeId === e.id),
  }));

  const save = (f) => {
    const existing = attendance.find(
      (a) => a.employeeId === f.employeeId && a.date === f.date,
    );
    if (existing)
      setAttendance(
        attendance.map((a) =>
          a.id === existing.id
            ? { ...a, ...f, checkOut: f.checkOut || null }
            : a,
        ),
      );
    else
      setAttendance([
        ...attendance,
        { id: uid("a"), ...f, checkOut: f.checkOut || null },
      ]);
    setModal(false);
  };

  if (role !== "admin" && currentEmp) {
    const myHistory = attendance
      .filter((a) => a.employeeId === currentEmp.id)
      .sort((a, b) => b.date.localeCompare(a.date));
    return (
      <div>
        <SelfPunch
          emp={currentEmp}
          shift={shiftOf(currentEmp.shiftId)}
          attendance={attendance}
          setAttendance={setAttendance}
          officeLocation={officeLocation}
        />
        <Card style={{ padding: 16 }}>
          <h3
            style={{
              fontFamily: "'Space Grotesk',sans-serif",
              fontWeight: 600,
              color: T.ink,
              marginBottom: 12,
              fontSize: 14,
            }}
          >
            ប្រវត្តិវត្តមានរបស់ខ្ញុំ
          </h3>
          <div style={{ overflowX: "auto" }}>
            <table className="wf-table">
              <thead>
                <tr>
                  <th>កាលបរិច្ឆេទ</th>
                  <th>ចូល</th>
                  <th>ចេញ</th>
                  <th>ស្ថានភាព</th>
                </tr>
              </thead>
              <tbody>
                {myHistory.length === 0 && (
                  <tr>
                    <td
                      colSpan={4}
                      style={{
                        textAlign: "center",
                        color: T.muted,
                        padding: "24px 0",
                      }}
                    >
                      មិនទាន់មានទិន្នន័យទេ
                    </td>
                  </tr>
                )}
                {myHistory.map((a) => (
                  <tr key={a.id}>
                    <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                      {a.date}
                    </td>
                    <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                      {a.checkIn || "—"}
                    </td>
                    <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                      {a.checkOut || "—"}
                    </td>
                    <td>
                      <StatusPill status={a.status} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      </div>
    );
  }

  return (
    <div>
      <OfficeLocationSettings
        officeLocation={officeLocation}
        setOfficeLocation={setOfficeLocation}
      />
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: 12,
          marginBottom: 16,
          flexWrap: "wrap",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <CalendarDays size={16} color={T.muted} />
          <Input
            type="date"
            style={{ width: "auto" }}
            value={date}
            onChange={(e) => setDate(e.target.value)}
          />
        </div>
        <Button variant="accent" onClick={() => setModal(true)}>
          <Plus size={15} /> កត់ត្រាដោយដៃ
        </Button>
      </div>
      <Card style={{ overflowX: "auto" }}>
        <table className="wf-table">
          <thead>
            <tr>
              <th>បុគ្គលិក</th>
              <th>វេន</th>
              <th>ចូល</th>
              <th>ចេញ</th>
              <th>ស្ថានភាព</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {rows.map(({ emp, rec }) => (
              <tr key={emp.id}>
                <td>
                  <div
                    style={{ display: "flex", alignItems: "center", gap: 10 }}
                  >
                    <Avatar name={emp.name} photo={emp.photo} size={30} />
                    <div>
                      <div
                        style={{ fontWeight: 500, color: T.ink, fontSize: 13 }}
                      >
                        {emp.name}
                      </div>
                      <div
                        style={{
                          fontSize: 10.5,
                          color: T.muted,
                          fontFamily: "'JetBrains Mono',monospace",
                        }}
                      >
                        {emp.code}
                      </div>
                    </div>
                  </div>
                </td>
                <td
                  style={{
                    fontSize: 11.5,
                    color: T.muted,
                    whiteSpace: "nowrap",
                  }}
                >
                  {shiftLabel(shiftOf(emp.shiftId))}
                </td>
                <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                  {rec?.checkIn || "—"}
                </td>
                <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                  {rec?.checkOut || "—"}
                </td>
                <td>
                  {rec ? (
                    <StatusPill status={rec.status} />
                  ) : (
                    <span style={{ fontSize: 12, color: T.mutedLight }}>
                      គ្មានទិន្នន័យ
                    </span>
                  )}
                </td>
                <td style={{ textAlign: "right" }}>
                  {rec && isSuperAdmin && (
                    <button
                      onClick={() => setConfirmDel(rec)}
                      style={{
                        background: "none",
                        border: "none",
                        cursor: "pointer",
                        color: T.mutedLight,
                      }}
                    >
                      <Trash2 size={14} />
                    </button>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>
      {modal && (
        <Modal title="កត់ត្រាវត្តមានដោយដៃ" onClose={() => setModal(false)}>
          <ManualAttendanceForm
            employees={activeEmployees}
            onSave={save}
            onCancel={() => setModal(false)}
          />
        </Modal>
      )}
      {confirmDel && (
        <ConfirmDialog
          text="តើអ្នកពិតជាចង់លុបកំណត់ត្រាវត្តមាននេះមែនទេ?"
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setAttendance(attendance.filter((a) => a.id !== confirmDel.id));
            setConfirmDel(null);
          }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Leave requests
----------------------------------------------------------------*/
function LeaveRequestForm({ onSave, onCancel }) {
  const [f, setF] = useState({
    type: "annual",
    startDate: todayStr(),
    endDate: todayStr(),
    reason: "",
  });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const invalidRange = f.endDate < f.startDate;
  return (
    <div>
      <Field label="ប្រភេទច្បាប់">
        <Select value={f.type} onChange={set("type")}>
          <option value="annual">ច្បាប់ប្រចាំឆ្នាំ</option>
          <option value="sick">ច្បាប់ឈឺ</option>
          <option value="other">ផ្សេងៗ</option>
        </Select>
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="ចាប់ពីថ្ងៃ">
          <Input type="date" value={f.startDate} onChange={set("startDate")} />
        </Field>
        <Field label="ដល់ថ្ងៃ">
          <Input type="date" value={f.endDate} onChange={set("endDate")} />
        </Field>
      </div>
      {invalidRange && (
        <p
          style={{
            fontSize: 12.5,
            color: T.rose,
            marginTop: -8,
            marginBottom: 12,
          }}
        >
          ថ្ងៃបញ្ចប់ត្រូវតែក្រោយថ្ងៃចាប់ផ្តើម
        </p>
      )}
      <Field label="មូលហេតុ">
        <textarea
          className="wf-input"
          rows={3}
          style={{ resize: "vertical", fontFamily: "inherit" }}
          value={f.reason}
          onChange={set("reason")}
          placeholder="សរសេរមូលហេតុសង្ខេប..."
        />
      </Field>
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          gap: 8,
          marginTop: 16,
          paddingTop: 14,
          borderTop: `1px solid ${T.lineSoft}`,
        }}
      >
        <Button variant="ghost" onClick={onCancel}>
          បោះបង់
        </Button>
        <Button
          variant="accent"
          onClick={() => onSave(f)}
          disabled={invalidRange || !f.startDate || !f.endDate}
        >
          ដាក់ស្នើ
        </Button>
      </div>
    </div>
  );
}

function LeaveRequests({
  role,
  currentEmp,
  employees,
  leaveRequests,
  setLeaveRequests,
  attendance,
  setAttendance,
  isSuperAdmin,
}) {
  const [modal, setModal] = useState(false);
  const [confirmDel, setConfirmDel] = useState(null);
  const empOf = (id) => employees.find((e) => e.id === id);

  // Approving a request marks every day in its range as paid leave in
  // attendance, so payroll (which already treats "leave" as paid) picks
  // it up automatically — no separate payroll logic needed.
  const applyLeaveToAttendance = (req) => {
    const days = dateRange(req.startDate, req.endDate);
    let next = attendance;
    for (const d of days) {
      const existing = next.find(
        (a) => a.employeeId === req.employeeId && a.date === d,
      );
      if (existing) {
        next = next.map((a) =>
          a.id === existing.id
            ? { ...a, status: "leave", checkIn: null, checkOut: null }
            : a,
        );
      } else {
        next = [
          ...next,
          {
            id: uid("a"),
            employeeId: req.employeeId,
            date: d,
            status: "leave",
            checkIn: null,
            checkOut: null,
          },
        ];
      }
    }
    setAttendance(next);
  };

  const decide = (req, status) => {
    setLeaveRequests(
      leaveRequests.map((r) =>
        r.id === req.id
          ? { ...r, status, reviewedAt: new Date().toISOString() }
          : r,
      ),
    );
    if (status === "approved") applyLeaveToAttendance(req);
  };

  const submit = (f) => {
    if (!currentEmp) return;
    setLeaveRequests([
      ...leaveRequests,
      {
        id: uid("lr"),
        employeeId: currentEmp.id,
        type: f.type,
        startDate: f.startDate,
        endDate: f.endDate,
        reason: f.reason.trim(),
        status: "pending",
        createdAt: new Date().toISOString(),
      },
    ]);
    setModal(false);
  };

  if (role !== "admin" && currentEmp) {
    const mine = leaveRequests
      .filter((r) => r.employeeId === currentEmp.id)
      .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
    return (
      <div>
        <div
          style={{
            display: "flex",
            justifyContent: "flex-end",
            marginBottom: 16,
          }}
        >
          <Button variant="accent" onClick={() => setModal(true)}>
            <Plus size={15} /> ស្នើសុំច្បាប់ឈប់សម្រាក
          </Button>
        </div>
        <Card style={{ overflowX: "auto" }}>
          <table className="wf-table">
            <thead>
              <tr>
                <th>ប្រភេទ</th>
                <th>ចាប់ពី</th>
                <th>ដល់</th>
                <th>មូលហេតុ</th>
                <th>ស្ថានភាព</th>
              </tr>
            </thead>
            <tbody>
              {mine.length === 0 && (
                <tr>
                  <td
                    colSpan={5}
                    style={{
                      textAlign: "center",
                      color: T.muted,
                      padding: "24px 0",
                    }}
                  >
                    មិនទាន់មានសំណើទេ
                  </td>
                </tr>
              )}
              {mine.map((r) => (
                <tr key={r.id}>
                  <td>{LEAVE_TYPE_LABEL[r.type] || r.type}</td>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {r.startDate}
                  </td>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {r.endDate}
                  </td>
                  <td style={{ fontSize: 12.5, color: T.textSoft }}>
                    {r.reason || "—"}
                  </td>
                  <td>
                    <StatusPill status={r.status} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>
        {modal && (
          <Modal title="ស្នើសុំច្បាប់ឈប់សម្រាក" onClose={() => setModal(false)}>
            <LeaveRequestForm
              onSave={submit}
              onCancel={() => setModal(false)}
            />
          </Modal>
        )}
      </div>
    );
  }

  // Admin view — newest first, pending requests surfaced on top.
  const sorted = [...leaveRequests].sort((a, b) => {
    if (a.status === "pending" && b.status !== "pending") return -1;
    if (a.status !== "pending" && b.status === "pending") return 1;
    return b.createdAt.localeCompare(a.createdAt);
  });
  return (
    <div>
      <Card style={{ overflowX: "auto" }}>
        <table className="wf-table">
          <thead>
            <tr>
              <th>បុគ្គលិក</th>
              <th>ប្រភេទ</th>
              <th>ចាប់ពី</th>
              <th>ដល់</th>
              <th>មូលហេតុ</th>
              <th>ស្ថានភាព</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {sorted.length === 0 && (
              <tr>
                <td
                  colSpan={7}
                  style={{
                    textAlign: "center",
                    color: T.muted,
                    padding: "24px 0",
                  }}
                >
                  មិនទាន់មានសំណើទេ
                </td>
              </tr>
            )}
            {sorted.map((r) => {
              const emp = empOf(r.employeeId);
              return (
                <tr key={r.id}>
                  <td>
                    <div
                      style={{ display: "flex", alignItems: "center", gap: 10 }}
                    >
                      <Avatar
                        name={emp?.name || "?"}
                        photo={emp?.photo}
                        size={30}
                      />
                      <div>
                        <div
                          style={{
                            fontWeight: 500,
                            color: T.ink,
                            fontSize: 13,
                          }}
                        >
                          {emp?.name || "—"}
                        </div>
                        <div
                          style={{
                            fontSize: 10.5,
                            color: T.muted,
                            fontFamily: "'JetBrains Mono',monospace",
                          }}
                        >
                          {emp?.code}
                        </div>
                      </div>
                    </div>
                  </td>
                  <td>{LEAVE_TYPE_LABEL[r.type] || r.type}</td>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {r.startDate}
                  </td>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {r.endDate}
                  </td>
                  <td
                    style={{ fontSize: 12.5, color: T.textSoft, maxWidth: 200 }}
                  >
                    {r.reason || "—"}
                  </td>
                  <td>
                    <StatusPill status={r.status} />
                  </td>
                  <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                    {r.status === "pending" ? (
                      <div
                        style={{
                          display: "flex",
                          gap: 6,
                          justifyContent: "flex-end",
                        }}
                      >
                        <Button
                          size="sm"
                          variant="accent"
                          onClick={() => decide(r, "approved")}
                        >
                          អនុម័ត
                        </Button>
                        <Button
                          size="sm"
                          variant="danger"
                          onClick={() => decide(r, "rejected")}
                        >
                          បដិសេធ
                        </Button>
                      </div>
                    ) : (
                      isSuperAdmin && (
                        <button
                          onClick={() => setConfirmDel(r)}
                          style={{
                            background: "none",
                            border: "none",
                            cursor: "pointer",
                            color: T.mutedLight,
                          }}
                        >
                          <Trash2 size={14} />
                        </button>
                      )
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
      {confirmDel && (
        <ConfirmDialog
          text="តើអ្នកពិតជាចង់លុបសំណើនេះមែនទេ?"
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setLeaveRequests(
              leaveRequests.filter((r) => r.id !== confirmDel.id),
            );
            setConfirmDel(null);
          }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Admin accounts — superadmin only. Two permission levels:
   superadmin (full access, incl. deleting records and managing other
   admin accounts) and manager (day-to-day HR work, no deletes).
----------------------------------------------------------------*/
function AdminAccountForm({ initial, onSave, onCancel }) {
  const [f, setF] = useState(
    initial || { username: "", password: "", name: "", role: "manager" },
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <div>
      <Field label="ឈ្មោះពេញ">
        <Input
          value={f.name}
          onChange={set("name")}
          placeholder="ឧ. សុខ សម្បត្តិ"
        />
      </Field>
      <Field label="ឈ្មោះគណនី (Username)">
        <Input
          value={f.username}
          onChange={set("username")}
          placeholder="ឧ. manager2"
        />
      </Field>
      <Field label="ពាក្យសម្ងាត់">
        <Input
          value={f.password}
          onChange={set("password")}
          placeholder="••••••••"
        />
      </Field>
      <Field label="សិទ្ធិ">
        <Select value={f.role} onChange={set("role")}>
          <option value="manager">
            អ្នកគ្រប់គ្រង HR (មិនអាចលុបទិន្នន័យបាន)
          </option>
          <option value="superadmin">
            អ្នកគ្រប់គ្រងជាន់ខ្ពស់ (សិទ្ធិពេញលេញ)
          </option>
        </Select>
      </Field>
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          gap: 8,
          marginTop: 16,
          paddingTop: 14,
          borderTop: `1px solid ${T.lineSoft}`,
        }}
      >
        <Button variant="ghost" onClick={onCancel}>
          បោះបង់
        </Button>
        <Button
          variant="accent"
          onClick={() => onSave(f)}
          disabled={!f.name || !f.username || !f.password}
        >
          រក្សាទុក
        </Button>
      </div>
    </div>
  );
}

function AdminAccounts({ admins, setAdmins, currentAdminId }) {
  const [modal, setModal] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const [error, setError] = useState("");

  const superadminCount = admins.filter((a) => a.role === "superadmin").length;

  const save = (data) => {
    const dupe = admins.find(
      (a) =>
        a.username.trim().toLowerCase() ===
          data.username.trim().toLowerCase() && a.id !== modal.data?.id,
    );
    if (dupe) {
      setError("ឈ្មោះគណនីនេះមានរួចហើយ");
      return;
    }
    setError("");
    if (modal.mode === "add")
      setAdmins([...admins, { ...data, id: uid("ad") }]);
    else setAdmins(admins.map((a) => (a.id === data.id ? { ...data } : a)));
    setModal(null);
  };

  const askDelete = (a) => {
    if (a.id === currentAdminId) return; // can't delete yourself
    if (a.role === "superadmin" && superadminCount <= 1) return; // keep at least one
    setConfirmDel(a);
  };

  return (
    <div>
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          marginBottom: 16,
        }}
      >
        <Button
          variant="accent"
          onClick={() => {
            setError("");
            setModal({ mode: "add" });
          }}
        >
          <Plus size={15} /> បន្ថែមគណនី
        </Button>
      </div>
      <Card style={{ overflowX: "auto" }}>
        <table className="wf-table">
          <thead>
            <tr>
              <th>ឈ្មោះ</th>
              <th>Username</th>
              <th>សិទ្ធិ</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {admins.map((a) => {
              const isSelf = a.id === currentAdminId;
              const isLastSuperadmin =
                a.role === "superadmin" && superadminCount <= 1;
              return (
                <tr key={a.id}>
                  <td>
                    <div
                      style={{ display: "flex", alignItems: "center", gap: 10 }}
                    >
                      <Avatar name={a.name} size={30} />
                      <div>
                        <div
                          style={{
                            fontWeight: 500,
                            color: T.ink,
                            fontSize: 13,
                          }}
                        >
                          {a.name}{" "}
                          {isSelf && (
                            <span style={{ color: T.muted }}>(អ្នក)</span>
                          )}
                        </div>
                      </div>
                    </div>
                  </td>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {a.username}
                  </td>
                  <td>
                    <span
                      className="wf-badge"
                      style={
                        a.role === "superadmin"
                          ? { background: "#EDE7F6", color: "#4A3B7A" }
                          : { background: T.forestSoft, color: T.forestText }
                      }
                    >
                      {ADMIN_ROLE_LABEL[a.role]}
                    </span>
                  </td>
                  <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                    <div style={{ display: "inline-flex", gap: 10 }}>
                      <button
                        onClick={() => {
                          setError("");
                          setModal({ mode: "edit", data: a });
                        }}
                        style={{
                          background: "none",
                          border: "none",
                          cursor: "pointer",
                          color: T.muted,
                        }}
                      >
                        <Pencil size={14} />
                      </button>
                      <button
                        onClick={() => askDelete(a)}
                        disabled={isSelf || isLastSuperadmin}
                        title={
                          isSelf
                            ? "មិនអាចលុបគណនីខ្លួនឯងបានទេ"
                            : isLastSuperadmin
                              ? "ត្រូវមានអ្នកគ្រប់គ្រងជាន់ខ្ពស់យ៉ាងតិចម្នាក់"
                              : undefined
                        }
                        style={{
                          background: "none",
                          border: "none",
                          cursor:
                            isSelf || isLastSuperadmin
                              ? "not-allowed"
                              : "pointer",
                          color:
                            isSelf || isLastSuperadmin ? T.mutedLight : T.rose,
                          opacity: isSelf || isLastSuperadmin ? 0.4 : 1,
                        }}
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
      {modal && (
        <Modal
          title={
            modal.mode === "add" ? "បន្ថែមគណនីអ្នកគ្រប់គ្រង" : "កែសម្រួលគណនី"
          }
          onClose={() => setModal(null)}
        >
          {error && (
            <p style={{ fontSize: 12.5, color: T.rose, marginBottom: 12 }}>
              {error}
            </p>
          )}
          <AdminAccountForm
            initial={modal.data}
            onSave={save}
            onCancel={() => setModal(null)}
          />
        </Modal>
      )}
      {confirmDel && (
        <ConfirmDialog
          text={`តើអ្នកពិតជាចង់លុបគណនី "${confirmDel.name}" មែនទេ?`}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setAdmins(admins.filter((a) => a.id !== confirmDel.id));
            setConfirmDel(null);
          }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   My profile — employee self-service. Lets a logged-in employee
   update their own contact info and photo, and change their PIN,
   without needing an admin to do it for them.
----------------------------------------------------------------*/
const MAX_PHOTO_BYTES = 800 * 1024; // keep stored photos reasonably small

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function MyProfile({
  currentEmp,
  employees,
  setEmployees,
  departments,
  shifts,
}) {
  const [f, setF] = useState({
    phone: currentEmp.phone || "",
    email: currentEmp.email || "",
  });
  const [photoPreview, setPhotoPreview] = useState(currentEmp.photo || null);
  const [photoError, setPhotoError] = useState("");
  const [saved, setSaved] = useState(false);

  const [pinForm, setPinForm] = useState({
    current: "",
    next: "",
    confirm: "",
  });
  const [pinError, setPinError] = useState("");
  const [pinSaved, setPinSaved] = useState(false);

  const deptName =
    departments.find((d) => d.id === currentEmp.deptId)?.name || "—";
  const shift = shifts.find((s) => s.id === currentEmp.shiftId);

  const set = (k) => (e) => {
    setF({ ...f, [k]: e.target.value });
    setSaved(false);
  };

  const onPhotoChange = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      setPhotoError("សូមជ្រើសរើសឯកសាររូបភាព");
      return;
    }
    if (file.size > MAX_PHOTO_BYTES) {
      setPhotoError("រូបភាពធំពេក សូមជ្រើសរើសរូបតូចជាងនេះ (តិចជាង 800KB)");
      return;
    }
    setPhotoError("");
    const dataUrl = await fileToDataUrl(file);
    setPhotoPreview(dataUrl);
    setSaved(false);
  };

  const saveProfile = () => {
    setEmployees(
      employees.map((e) =>
        e.id === currentEmp.id
          ? { ...e, phone: f.phone, email: f.email, photo: photoPreview }
          : e,
      ),
    );
    setSaved(true);
  };

  const savePin = () => {
    setPinSaved(false);
    if ((currentEmp.pin || "") !== pinForm.current.trim()) {
      setPinError("កូដសម្ងាត់បច្ចុប្បន្នមិនត្រឹមត្រូវទេ");
      return;
    }
    if (!pinForm.next.trim() || pinForm.next.trim().length < 4) {
      setPinError("កូដសម្ងាត់ថ្មីត្រូវមានយ៉ាងតិច 4 ខ្ទង់");
      return;
    }
    if (pinForm.next.trim() !== pinForm.confirm.trim()) {
      setPinError("ការបញ្ជាក់កូដសម្ងាត់ថ្មីមិនត្រូវគ្នាទេ");
      return;
    }
    setPinError("");
    setEmployees(
      employees.map((e) =>
        e.id === currentEmp.id ? { ...e, pin: pinForm.next.trim() } : e,
      ),
    );
    setPinForm({ current: "", next: "", confirm: "" });
    setPinSaved(true);
  };

  return (
    <div
      className="wf-grid"
      style={{ gridTemplateColumns: "repeat(auto-fit, minmax(320px,1fr))" }}
    >
      <Card style={{ padding: 20 }}>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 14,
            marginBottom: 18,
          }}
        >
          <Avatar name={currentEmp.name} photo={photoPreview} size={56} />
          <div style={{ minWidth: 0 }}>
            <div style={{ fontWeight: 700, color: T.ink, fontSize: 15 }}>
              {currentEmp.name}
            </div>
            <div
              style={{
                fontSize: 11.5,
                color: T.muted,
                fontFamily: "'JetBrains Mono',monospace",
              }}
            >
              {currentEmp.code}
            </div>
          </div>
        </div>
        <div
          style={{
            fontSize: 12,
            color: T.textSoft,
            marginBottom: 18,
            display: "flex",
            flexDirection: "column",
            gap: 4,
          }}
        >
          <div>
            🏢 {deptName} · {currentEmp.role}
          </div>
          <div>⏰ {shiftLabel(shift)}</div>
          <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
            <BadgeCheck size={12} color={T.muted} /> ចូលបម្រើការតាំងពី{" "}
            {currentEmp.joined}
          </div>
        </div>
        <Field label="រូបភាពប្រវត្តិរូប">
          <input
            type="file"
            accept="image/*"
            onChange={onPhotoChange}
            style={{ fontSize: 12.5 }}
          />
        </Field>
        {photoError && (
          <p style={{ fontSize: 12.5, color: T.rose, marginBottom: 10 }}>
            {photoError}
          </p>
        )}
        <Field label="លេខទូរស័ព្ទ">
          <Input value={f.phone} onChange={set("phone")} />
        </Field>
        <Field label="អ៊ីមែល">
          <Input value={f.email} onChange={set("email")} type="email" />
        </Field>
        {saved && (
          <p
            style={{
              fontSize: 12.5,
              color: T.forestText,
              display: "flex",
              alignItems: "center",
              gap: 5,
              marginBottom: 10,
            }}
          >
            <CheckCircle2 size={14} /> បានរក្សាទុករួចរាល់
          </p>
        )}
        <Button variant="accent" onClick={saveProfile}>
          រក្សាទុកព័ត៌មាន
        </Button>
      </Card>

      <Card style={{ padding: 20 }}>
        <h3
          style={{
            fontFamily: "'Space Grotesk',sans-serif",
            fontWeight: 600,
            color: T.ink,
            marginBottom: 4,
            fontSize: 14,
            display: "flex",
            alignItems: "center",
            gap: 8,
          }}
        >
          <KeyRound size={16} /> ប្តូរកូដសម្ងាត់ (PIN)
        </h3>
        <p style={{ fontSize: 12, color: T.muted, marginBottom: 16 }}>
          កូដសម្ងាត់នេះប្រើសម្រាប់ចូលប្រើប្រព័ន្ធ។ សូមកុំប្រាប់អ្នកដទៃ។
        </p>
        <Field label="កូដសម្ងាត់បច្ចុប្បន្ន">
          <Input
            type="password"
            inputMode="numeric"
            maxLength={6}
            value={pinForm.current}
            onChange={(e) => {
              setPinForm({ ...pinForm, current: e.target.value });
              setPinSaved(false);
            }}
          />
        </Field>
        <Field label="កូដសម្ងាត់ថ្មី">
          <Input
            type="password"
            inputMode="numeric"
            maxLength={6}
            value={pinForm.next}
            onChange={(e) => {
              setPinForm({ ...pinForm, next: e.target.value });
              setPinSaved(false);
            }}
          />
        </Field>
        <Field label="បញ្ជាក់កូដសម្ងាត់ថ្មី">
          <Input
            type="password"
            inputMode="numeric"
            maxLength={6}
            value={pinForm.confirm}
            onChange={(e) => {
              setPinForm({ ...pinForm, confirm: e.target.value });
              setPinSaved(false);
            }}
          />
        </Field>
        {pinError && (
          <p style={{ fontSize: 12.5, color: T.rose, marginBottom: 10 }}>
            {pinError}
          </p>
        )}
        {pinSaved && (
          <p
            style={{
              fontSize: 12.5,
              color: T.forestText,
              display: "flex",
              alignItems: "center",
              gap: 5,
              marginBottom: 10,
            }}
          >
            <CheckCircle2 size={14} /> បានប្តូរកូដសម្ងាត់ដោយជោគជ័យ
          </p>
        )}
        <Button variant="accent" onClick={savePin}>
          ប្តូរកូដសម្ងាត់
        </Button>
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------
   Payroll
----------------------------------------------------------------*/
function Payslip({ emp, mk, attendance, onClose }) {
  const {
    absentDays,
    leaveDays,
    dailyRate,
    absenceDeduction,
    tax,
    insurance,
    net,
  } = computePayroll(emp, attendance, mk);
  return (
    <Modal title="" onClose={onClose} width={460}>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          marginBottom: 16,
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 8,
            color: T.forest,
            fontWeight: 700,
            fontFamily: "'Space Grotesk',sans-serif",
          }}
        >
          <Receipt size={20} /> សន្លឹកប្រាក់ខែ
        </div>
        <span
          style={{
            fontSize: 12,
            color: T.muted,
            fontFamily: "'JetBrains Mono',monospace",
          }}
        >
          {mk}
        </span>
      </div>
      <div
        style={{
          border: `1px solid ${T.lineSoft}`,
          borderRadius: 10,
          padding: 14,
          marginBottom: 16,
        }}
      >
        <div style={{ fontSize: 14, fontWeight: 600, color: T.ink }}>
          {emp.name}
        </div>
        <div
          style={{
            fontSize: 11,
            color: T.muted,
            fontFamily: "'JetBrains Mono',monospace",
          }}
        >
          {emp.code} · {emp.role}
        </div>
      </div>
      <div
        style={{
          fontSize: 13,
          display: "flex",
          flexDirection: "column",
          gap: 8,
          marginBottom: 12,
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between" }}>
          <span style={{ color: T.textSoft }}>ប្រាក់ខែមូលដ្ឋាន</span>
          <span style={{ fontFamily: "'JetBrains Mono',monospace" }}>
            {fmtMoney(emp.salary)}
          </span>
        </div>
        {absentDays > 0 && (
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              color: T.rose,
            }}
          >
            <span>
              ដកប្រាក់ដោយសារអវត្តមាន ({absentDays} ថ្ងៃ × {fmtMoney(dailyRate)})
            </span>
            <span style={{ fontFamily: "'JetBrains Mono',monospace" }}>
              -{fmtMoney(absenceDeduction)}
            </span>
          </div>
        )}
        {leaveDays > 0 && (
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              color: T.blue,
              fontSize: 12,
            }}
          >
            <span>ថ្ងៃឈប់សម្រាក (មានប្រាក់ខែ)</span>
            <span style={{ fontFamily: "'JetBrains Mono',monospace" }}>
              {leaveDays} ថ្ងៃ
            </span>
          </div>
        )}
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            color: T.rose,
          }}
        >
          <span>ពន្ធលើប្រាក់ខែ (5%)</span>
          <span style={{ fontFamily: "'JetBrains Mono',monospace" }}>
            -{fmtMoney(tax)}
          </span>
        </div>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            color: T.rose,
          }}
        >
          <span>ធានារ៉ាប់រង (2%)</span>
          <span style={{ fontFamily: "'JetBrains Mono',monospace" }}>
            -{fmtMoney(insurance)}
          </span>
        </div>
      </div>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          background: T.ink,
          color: "#fff",
          borderRadius: 10,
          padding: "12px 16px",
          marginTop: 16,
        }}
      >
        <span style={{ fontWeight: 600, fontSize: 13 }}>ប្រាក់ខែសុទ្ធ</span>
        <span
          style={{
            fontWeight: 700,
            fontSize: 18,
            fontFamily: "'JetBrains Mono',monospace",
          }}
        >
          {fmtMoney(net)}
        </span>
      </div>
    </Modal>
  );
}

function Payroll({
  role,
  currentEmp,
  employees,
  attendance,
  payrollPaid,
  setPayrollPaid,
}) {
  const currentMk = monthKey();
  const availableMonths = useMemo(() => {
    const set = new Set([currentMk]);
    attendance.forEach((a) => {
      if (a.date) set.add(a.date.slice(0, 7));
    });
    return Array.from(set).sort((a, b) => (a < b ? 1 : -1));
  }, [attendance, currentMk]);
  const [mk, setMk] = useState(currentMk);
  const [slipFor, setSlipFor] = useState(null);
  const isPastMonth = mk !== currentMk;
  const activeEmployees = employees.filter((e) => e.status === "active");
  const list =
    role === "admin"
      ? activeEmployees
      : activeEmployees.filter((e) => e.id === currentEmp?.id);
  const togglePaid = (empId) => {
    const key = `${empId}-${mk}`;
    setPayrollPaid({ ...payrollPaid, [key]: !payrollPaid[key] });
  };
  const totalNet = list.reduce(
    (sum, e) => sum + computePayroll(e, attendance, mk).net,
    0,
  );

  return (
    <div>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          marginBottom: 16,
          flexWrap: "wrap",
        }}
      >
        <span
          style={{
            fontSize: 11,
            fontWeight: 700,
            color: T.muted,
            textTransform: "uppercase",
            letterSpacing: ".03em",
          }}
        >
          ខែ
        </span>
        <Select
          value={mk}
          onChange={(e) => setMk(e.target.value)}
          style={{ width: "auto", minWidth: 170 }}
        >
          {availableMonths.map((m) => (
            <option key={m} value={m}>
              {monthLabel(m)}
              {m === currentMk ? " (ខែបច្ចុប្បន្ន)" : ""}
            </option>
          ))}
        </Select>
        {isPastMonth && (
          <span
            style={{
              fontSize: 11,
              fontWeight: 700,
              background: T.goldSoft,
              color: T.goldText,
              padding: "5px 10px",
              borderRadius: 8,
            }}
          >
            កំពុងមើលប្រវត្តិខែមុន
          </span>
        )}
      </div>
      {role === "admin" && (
        <Card
          accent={T.gold}
          style={{
            padding: 16,
            marginBottom: 16,
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
          }}
        >
          <div>
            <div
              style={{
                fontSize: 11,
                color: T.muted,
                textTransform: "uppercase",
                fontWeight: 700,
              }}
            >
              ចំណាយប្រាក់ខែសរុប · {monthLabel(mk)}
            </div>
            <div
              style={{
                fontSize: 24,
                fontWeight: 700,
                color: T.ink,
                fontFamily: "'JetBrains Mono',monospace",
              }}
            >
              {fmtMoney(totalNet)}
            </div>
          </div>
          <BadgeCheck size={28} color={T.gold} />
        </Card>
      )}
      <Card style={{ overflowX: "auto" }}>
        <table className="wf-table">
          <thead>
            <tr>
              <th>បុគ្គលិក</th>
              <th>ប្រាក់ខែមូលដ្ឋាន</th>
              <th>ប្រាក់ខែសុទ្ធ</th>
              <th>ស្ថានភាព</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {list.map((e) => {
              const paid = !!payrollPaid[`${e.id}-${mk}`];
              const { net, absentDays } = computePayroll(e, attendance, mk);
              return (
                <tr key={e.id}>
                  <td>
                    <div
                      style={{ display: "flex", alignItems: "center", gap: 10 }}
                    >
                      <Avatar name={e.name} photo={e.photo} size={30} />
                      <div>
                        <div
                          style={{
                            fontWeight: 500,
                            color: T.ink,
                            fontSize: 13,
                          }}
                        >
                          {e.name}
                        </div>
                        {absentDays > 0 && (
                          <div style={{ fontSize: 10.5, color: T.rose }}>
                            អវត្តមាន {absentDays} ថ្ងៃ
                          </div>
                        )}
                      </div>
                    </div>
                  </td>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {fmtMoney(e.salary)}
                  </td>
                  <td
                    style={{
                      fontWeight: 700,
                      color: T.forestText,
                      fontFamily: "'JetBrains Mono',monospace",
                    }}
                  >
                    {fmtMoney(net)}
                  </td>
                  <td>
                    <StatusPill status={paid ? "paid" : "pending"} />
                  </td>
                  <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                    <button
                      onClick={() => setSlipFor(e)}
                      style={{
                        background: "none",
                        border: "none",
                        cursor: "pointer",
                        fontSize: 12,
                        fontWeight: 700,
                        color: T.blue,
                        marginRight: 12,
                      }}
                    >
                      មើលសន្លឹក
                    </button>
                    {role === "admin" && (
                      <button
                        onClick={() => togglePaid(e.id)}
                        style={{
                          background: "none",
                          border: "none",
                          cursor: "pointer",
                          fontSize: 12,
                          fontWeight: 700,
                          color: T.forest,
                        }}
                      >
                        {paid ? "ដកសញ្ញាបង់" : "សម្គាល់ថាបានបង់"}
                      </button>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
      {slipFor && (
        <Payslip
          emp={slipFor}
          mk={mk}
          attendance={attendance}
          onClose={() => setSlipFor(null)}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   App shell
----------------------------------------------------------------*/
const NAV_ADMIN = [
  { id: "dashboard", label: "ផ្ទាំងគ្រប់គ្រង", icon: LayoutDashboard },
  { id: "employees", label: "បុគ្គលិក", icon: Users },
  { id: "departments", label: "នាយកដ្ឋាន", icon: Building2 },
  { id: "shifts", label: "វេន", icon: Watch },
  { id: "attendance", label: "វត្តមាន", icon: Clock },
  { id: "leave", label: "ច្បាប់ឈប់សម្រាក", icon: CalendarDays },
  { id: "payroll", label: "ប្រាក់ខែ", icon: Wallet },
  // Superadmin-only — filtered out for manager accounts below.
  {
    id: "admins",
    label: "គណនីអ្នកគ្រប់គ្រង",
    icon: ShieldCheck,
    superadminOnly: true,
  },
];
const NAV_EMPLOYEE = [
  { id: "dashboard", label: "ផ្ទាំងគ្រប់គ្រង", icon: LayoutDashboard },
  { id: "attendance", label: "វត្តមានរបស់ខ្ញុំ", icon: Clock },
  { id: "leave", label: "ច្បាប់ឈប់សម្រាករបស់ខ្ញុំ", icon: CalendarDays },
  { id: "payroll", label: "ប្រាក់ខែរបស់ខ្ញុំ", icon: Wallet },
  { id: "profile", label: "ប្រវត្តិរូបរបស់ខ្ញុំ", icon: UserCircle2 },
];

export default function App() {
  useGlobalStyle();
  const [departments, setDepartments, dReady] = useSupabaseArray(
    "departments",
    {
      fromDb: (r) => ({
        id: r.id,
        name: r.name,
        code: r.code,
        desc: r.description,
      }),
      toDb: (r) => ({
        id: r.id,
        name: r.name,
        code: r.code,
        description: r.desc,
      }),
    },
  );
  const [employees, setEmployees, eReady] = useSupabaseArray("employees", {
    fromDb: (r) => ({
      id: r.id,
      code: r.code,
      pin: r.pin,
      name: r.name,
      deptId: r.dept_id,
      shiftId: r.shift_id,
      role: r.role,
      email: r.email,
      phone: r.phone,
      salary: r.salary,
      status: r.status,
      joined: r.joined,
      photo: r.photo,
    }),
    toDb: (r) => ({
      id: r.id,
      code: r.code,
      pin: r.pin,
      name: r.name,
      dept_id: r.deptId,
      shift_id: r.shiftId,
      role: r.role,
      email: r.email,
      phone: r.phone,
      salary: r.salary,
      status: r.status,
      joined: r.joined,
      photo: r.photo,
    }),
  });
  const [shifts, setShifts, shReady] = useSupabaseArray("shifts");
  const [attendance, setAttendance, aReady] = useSupabaseArray("attendance", {
    fromDb: (r) => ({
      id: r.id,
      employeeId: r.employee_id,
      date: r.date,
      checkIn: r.check_in,
      checkOut: r.check_out,
      status: r.status,
      checkInLoc: r.check_in_loc,
      checkOutLoc: r.check_out_loc,
    }),
    toDb: (r) => ({
      id: r.id,
      employee_id: r.employeeId,
      date: r.date,
      check_in: r.checkIn,
      check_out: r.checkOut,
      status: r.status,
      check_in_loc: r.checkInLoc,
      check_out_loc: r.checkOutLoc,
    }),
  });
  const [payrollPaid, setPayrollPaid, pReady] = usePayrollPaid();
  const [leaveRequests, setLeaveRequests, lrReady] = useSupabaseArray(
    "leave_requests",
    {
      fromDb: (r) => ({
        id: r.id,
        employeeId: r.employee_id,
        type: r.type,
        startDate: r.start_date,
        endDate: r.end_date,
        reason: r.reason,
        status: r.status,
        createdAt: r.created_at,
      }),
      toDb: (r) => ({
        id: r.id,
        employee_id: r.employeeId,
        type: r.type,
        start_date: r.startDate,
        end_date: r.endDate,
        reason: r.reason,
        status: r.status,
        created_at: r.createdAt,
      }),
    },
  );
  const [admins, setAdmins, adminsReady] = useSupabaseArray("admins");
  const [officeLocation, setOfficeLocation, olReady] = useOfficeLocation();
  // Session state stays personal (per device) on purpose — each phone
  // remembers only who is logged in on THAT phone, so employees never
  // see each other's login state.
  const [sessionAdmin, setSessionAdmin, sAdminReady] = useLocalStorage(
    K.SESSION_ADMIN,
    null,
  );
  const [sessionEmployee, setSessionEmployee, sEmpReady] = useLocalStorage(
    K.SESSION_EMPLOYEE,
    null,
  );
  const [page, setPage] = useState("dashboard");
  const [navOpen, setNavOpen] = useState(false);
  const [portal, goPortal] = usePortalRoute();

  const ready =
    dReady &&
    eReady &&
    shReady &&
    aReady &&
    pReady &&
    lrReady &&
    adminsReady &&
    olReady &&
    sAdminReady &&
    sEmpReady;
  // Each portal only ever reads its own session — the admin URL never
  // shows an employee who is signed in elsewhere, and vice versa.
  // sessionAdmin now stores an admin account id (not a fixed "admin"
  // literal), so different admin accounts can carry different permissions.
  const currentAdmin =
    portal === "admin" && sessionAdmin
      ? admins.find((a) => a.id === sessionAdmin) || null
      : null;
  const isSuperAdmin = currentAdmin?.role === "superadmin";
  const role =
    portal === "admin"
      ? currentAdmin
        ? "admin"
        : null
      : sessionEmployee || null;
  const currentEmp =
    role && role !== "admin" ? employees.find((e) => e.id === role) : null;
  const loggedIn = role === "admin" || !!currentEmp;
  const nav =
    role === "admin"
      ? NAV_ADMIN.filter((n) => !n.superadminOnly || isSuperAdmin)
      : NAV_EMPLOYEE;

  useEffect(() => {
    if (loggedIn && !nav.find((n) => n.id === page)) setPage("dashboard");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [role]);

  if (!ready) {
    return (
      <div
        style={{
          minHeight: 500,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          background: T.paper,
          borderRadius: 12,
        }}
      >
        <Loader2
          size={28}
          color={T.forest}
          style={{ animation: "spin 1s linear infinite" }}
        />
        <style>
          {
            "@keyframes spin{from{transform:rotate(0)}to{transform:rotate(360deg)}}"
          }
        </style>
      </div>
    );
  }

  if (!loggedIn) {
    if (portal === "admin")
      return (
        <AdminLoginScreen
          admins={admins}
          onLogin={(id) => setSessionAdmin(id)}
          go={goPortal}
        />
      );
    return (
      <EmployeeLoginScreen
        employees={employees}
        onLogin={(id) => setSessionEmployee(id)}
      />
    );
  }

  return (
    <div className="wf-root">
      <div
        className={`wf-overlay-scrim ${navOpen ? "open" : ""}`}
        onClick={() => setNavOpen(false)}
      />
      <aside className={`wf-sidebar ${navOpen ? "open" : ""}`}>
        <div className="wf-sidebar-inner">
          <div
            style={{
              padding: "18px 18px",
              display: "flex",
              alignItems: "center",
              gap: 10,
              borderBottom: "1px solid rgba(255,255,255,0.1)",
            }}
          >
            <div className="wf-logo-badge">WS</div>
            <span
              style={{
                fontWeight: 600,
                fontSize: 15,
                fontFamily: "'Space Grotesk',sans-serif",
              }}
            >
              Workforce Suite
            </span>
            <button
              className="wf-menu-btn"
              style={{ marginLeft: "auto", color: "rgba(255,255,255,0.6)" }}
              onClick={() => setNavOpen(false)}
            >
              <X size={18} />
            </button>
          </div>
          <nav
            style={{
              flex: 1,
              minHeight: 0,
              overflowY: "auto",
              padding: "14px 10px",
              display: "flex",
              flexDirection: "column",
              gap: 4,
            }}
          >
            {nav.map((n) => (
              <button
                key={n.id}
                className={`wf-nav-item ${page === n.id ? "active" : ""}`}
                onClick={() => {
                  setPage(n.id);
                  setNavOpen(false);
                }}
              >
                <n.icon size={17} /> {n.label}
              </button>
            ))}
          </nav>
          <div
            style={{
              padding: 12,
              borderTop: "1px solid rgba(255,255,255,0.1)",
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                padding: "9px 10px",
                marginBottom: 8,
                background: "rgba(255,255,255,0.05)",
                border: "1px solid rgba(255,255,255,0.07)",
                borderRadius: 12,
              }}
            >
              <Avatar
                name={
                  role === "admin"
                    ? currentAdmin?.name || "?"
                    : currentEmp?.name || "?"
                }
                photo={role === "admin" ? null : currentEmp?.photo}
                size={34}
              />
              <div style={{ minWidth: 0 }}>
                <div
                  style={{
                    fontSize: 13,
                    fontWeight: 600,
                    color: "#fff",
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                    whiteSpace: "nowrap",
                  }}
                >
                  {role === "admin" ? currentAdmin?.name : currentEmp?.name}
                </div>
                <div
                  style={{
                    fontSize: 10.5,
                    color: "#7C8AA3",
                    fontFamily: "'JetBrains Mono',monospace",
                  }}
                >
                  {role === "admin"
                    ? ADMIN_ROLE_LABEL[currentAdmin?.role]
                    : currentEmp?.code}
                </div>
              </div>
            </div>
            <button
              className="wf-nav-item"
              style={{ color: "#E3B7BE" }}
              onClick={() => {
                if (role === "admin") setSessionAdmin(null);
                else setSessionEmployee(null);
              }}
            >
              <LogOut size={17} /> ចាកចេញ
            </button>
          </div>
        </div>
      </aside>

      <div className="wf-main">
        <header className="wf-header">
          <button className="wf-menu-btn" onClick={() => setNavOpen(true)}>
            <Menu size={20} />
          </button>
          <h1
            style={{
              fontFamily: "'Space Grotesk',sans-serif",
              fontWeight: 600,
              color: T.ink,
              fontSize: 16,
            }}
          >
            {nav.find((n) => n.id === page)?.label}
          </h1>
          <div
            style={{
              marginLeft: "auto",
              display: "flex",
              alignItems: "center",
              gap: 8,
            }}
          >
            {role !== "admin" && currentEmp && (
              <span
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  background: T.forestSoft,
                  color: T.forestText,
                  padding: "5px 10px",
                  borderRadius: 8,
                }}
              >
                បុគ្គលិក
              </span>
            )}
            {role === "admin" && (
              <span
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  background: "#EDE7F6",
                  color: "#4A3B7A",
                  padding: "5px 10px",
                  borderRadius: 8,
                }}
              >
                {ADMIN_ROLE_LABEL[currentAdmin?.role]}
              </span>
            )}
            <NotificationBell
              role={role}
              currentAdmin={currentAdmin}
              currentEmp={currentEmp}
              employees={employees}
              shifts={shifts}
              attendance={attendance}
              leaveRequests={leaveRequests}
              setPage={setPage}
            />
            <Avatar
              name={
                role === "admin"
                  ? currentAdmin?.name || "?"
                  : currentEmp?.name || "?"
              }
              photo={role === "admin" ? null : currentEmp?.photo}
              size={32}
            />
          </div>
        </header>

        <main className="wf-content">
          {employees.length === 0 &&
            role === "admin" &&
            page !== "employees" && (
              <div
                style={{
                  marginBottom: 16,
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  background: T.goldSoft,
                  color: T.goldText,
                  fontSize: 13,
                  padding: "10px 16px",
                  borderRadius: 10,
                }}
              >
                <AlertCircle size={16} /> មិនទាន់មានបុគ្គលិកទេ សូមចូលទៅផ្នែក
                "បុគ្គលិក" ដើម្បីបន្ថែម
              </div>
            )}
          {page === "dashboard" && (
            <Dashboard
              employees={employees}
              departments={departments}
              attendance={attendance}
              payrollPaid={payrollPaid}
              role={role}
              currentEmp={currentEmp}
            />
          )}
          {page === "employees" && role === "admin" && (
            <Employees
              employees={employees}
              departments={departments}
              shifts={shifts}
              setEmployees={setEmployees}
              isSuperAdmin={isSuperAdmin}
            />
          )}
          {page === "departments" && role === "admin" && (
            <Departments
              departments={departments}
              setDepartments={setDepartments}
              employees={employees}
              isSuperAdmin={isSuperAdmin}
            />
          )}
          {page === "shifts" && role === "admin" && (
            <Shifts
              shifts={shifts}
              setShifts={setShifts}
              employees={employees}
              isSuperAdmin={isSuperAdmin}
            />
          )}
          {page === "attendance" && (
            <Attendance
              role={role}
              currentEmp={currentEmp}
              employees={employees}
              shifts={shifts}
              attendance={attendance}
              setAttendance={setAttendance}
              isSuperAdmin={isSuperAdmin}
              officeLocation={officeLocation}
              setOfficeLocation={setOfficeLocation}
            />
          )}
          {page === "leave" && (
            <LeaveRequests
              role={role}
              currentEmp={currentEmp}
              employees={employees}
              leaveRequests={leaveRequests}
              setLeaveRequests={setLeaveRequests}
              attendance={attendance}
              setAttendance={setAttendance}
              isSuperAdmin={isSuperAdmin}
            />
          )}
          {page === "payroll" && (
            <Payroll
              role={role}
              currentEmp={currentEmp}
              employees={employees}
              attendance={attendance}
              payrollPaid={payrollPaid}
              setPayrollPaid={setPayrollPaid}
            />
          )}
          {page === "admins" && role === "admin" && isSuperAdmin && (
            <AdminAccounts
              admins={admins}
              setAdmins={setAdmins}
              currentAdminId={currentAdmin?.id}
            />
          )}
          {page === "profile" && role !== "admin" && currentEmp && (
            <MyProfile
              currentEmp={currentEmp}
              employees={employees}
              setEmployees={setEmployees}
              departments={departments}
              shifts={shifts}
            />
          )}
        </main>
      </div>
    </div>
  );
}
