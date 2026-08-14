import React, {
  useState,
  useEffect,
  useCallback,
  useRef,
  useMemo,
  createContext,
  useContext as useCtx,
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
  Timer,
  ThumbsUp,
  ThumbsDown,
  Sun,
  Moon,
  Settings2,
  Camera,
} from "lucide-react";

/* ---------------------------------------------------------------
   Language / i18n  — add new keys here as needed
----------------------------------------------------------------*/
const LANG = {
  km: {
    appName: "Workforce Suite",
    login: {
      employeeId: "លេខសម្គាល់បុគ្គលិក",
      employeeIdPlaceholder: "EMP-001",
      pin: "កូដសម្ងាត់ (PIN)",
      pinPlaceholder: "••••",
      submit: "ចូលប្រើប្រព័ន្ធ",
      adminSubmit: "ចូលជាអ្នកគ្រប់គ្រង",
      username: "ឈ្មោះគណនី (Username)",
      usernamePlaceholder: "admin",
      password: "ពាក្យសម្ងាត់អ្នកគ្រប់គ្រង",
      passwordPlaceholder: "••••••••",
      adminTitle: "ចូលប្រើសម្រាប់អ្នកគ្រប់គ្រង",
      employeePortal: "ចូលប្រើបុគ្គលិក",
      demoLabel: "សម្រាប់សាកល្បង៖",
      errNoEmp: "រកមិនឃើញលេខសម្គាល់បុគ្គលិកនេះទេ",
      errInactive: "គណនីនេះមិនទាន់សកម្មទេ សូមទាក់ទង Admin",
      errPin: "កូដសម្ងាត់ (PIN) មិនត្រឹមត្រូវទេ",
      errNoAdmin: "រកមិនឃើញគណនីអ្នកគ្រប់គ្រងនេះទេ",
      errPass: "ពាក្យសម្ងាត់មិនត្រឹមត្រូវទេ",
      back: "ត្រឡប់ក្រោយ",
      switchToAdmin: "អ្នកគ្រប់គ្រង? ចូលពីទីនេះ",
    },
    nav: {
      dashboard: "ផ្ទាំងគ្រប់គ្រង",
      employees: "បុគ្គលិក",
      departments: "នាយកដ្ឋាន",
      shifts: "វេនការងារ",
      attendance: "វត្តមាន",
      leave: "ច្បាប់ឈប់សម្រាក",
      overtime: "ការងារបន្ថែម (OT)",
      payroll: "ប្រាក់ខែ",
      admins: "គណនីអ្នកគ្រប់គ្រង",
      myAttendance: "វត្តមានរបស់ខ្ញុំ",
      myLeave: "ច្បាប់ឈប់សម្រាករបស់ខ្ញុំ",
      myOvertime: "ការងារបន្ថែម (OT) របស់ខ្ញុំ",
      myPayroll: "ប្រាក់ខែរបស់ខ្ញុំ",
      myProfile: "ប្រវត្តិរូបរបស់ខ្ញុំ",
      settings: "ការកំណត់",
    },
    logout: "ចាកចេញ",
    notifications: "ការជូនដំណឹង",
    markAllRead: "កំណត់ថាបានអានទាំងអស់",
    noNotif: "មិនមានការជូនដំណឹងទេ",
    confirmDelete: "បញ្ជាក់ការលុប",
    cancel: "បោះបង់",
    delete: "លុប",
    edit: "កែសម្រួល",
    save: "រក្សាទុក",
    add: "បន្ថែម",
    search: "ស្វែងរក...",
    employee: "បុគ្គលិក",
    status: "ស្ថានភាព",
    actions: "សកម្មភាព",
    noData: "មិនមានទិន្នន័យ",
    dash: {
      welcome: "សូមអញ្ជើញ",
      totalEmp: "បុគ្គលិកសរុប",
      active: "សកម្ម",
      totalDept: "នាយកដ្ឋាន",
      totalDeptSub: "នាយកដ្ឋានសរុប",
      presentToday: "មកធ្វើការថ្ងៃនេះ",
      attendRate: "អត្រាមកធ្វើការ",
      pendingPayroll: "រង់ចាំបើកប្រាក់ខែ",
      thisMonth: "ខែនេះ",
      recentAttend: "ការចូលធ្វើការថ្ងៃនេះ",
      noAttend: "មិនទាន់មានការចូលធ្វើការថ្ងៃនេះ",
      empPortalLink: "តំណភ្ជាប់សម្រាប់បុគ្គលិក",
      showQR: "បង្ហាញ QR Code",
      copyLink: "ចម្លងតំណ",
      copied: "បានចម្លងហើយ!",
      noEmpWarn: 'មិនទាន់មានបុគ្គលិកទេ សូមចូលទៅផ្នែក "បុគ្គលិក" ដើម្បីបន្ថែម',
      myDept: "នាយកដ្ឋានរបស់ខ្ញុំ",
      myShift: "វេនការងាររបស់ខ្ញុំ",
      todayStatus: "ស្ថានភាពថ្ងៃនេះ",
      payrollStatus: "ស្ថានភាពប្រាក់ខែ",
      notCheckedIn: "មិនទាន់ចូលធ្វើការ",
    },
    depts: {
      addBtn: "បន្ថែមនាយកដ្ឋាន",
      editTitle: "កែសម្រួលនាយកដ្ឋាន",
      addTitle: "បន្ថែមនាយកដ្ឋាន",
      nameLabel: "ឈ្មោះនាយកដ្ឋាន",
      codeLabel: "លេខកូដ (ខ្លី)",
      descLabel: "ការពិពណ៌នា",
      staffCount: "នាក់",
      noDesc: "គ្មានការពិពណ៌នា",
      confirmDel: "តើអ្នកប្រាកដទេថាចង់លុបនាយកដ្ឋាននេះ?",
    },
    emps: {
      addBtn: "បន្ថែមបុគ្គលិក",
      editTitle: "កែសម្រួលបុគ្គលិក",
      addTitle: "បន្ថែមបុគ្គលិក",
      name: "ឈ្មោះ",
      code: "លេខសម្គាល់",
      dept: "នាយកដ្ឋាន",
      shift: "វេន",
      role: "តួនាទី",
      salary: "ប្រាក់ខែ",
      phone: "លេខទូរស័ព្ទ",
      email: "អ៊ីម៉ែល",
      joined: "ថ្ងៃចូលធ្វើការ",
      pin: "PIN (លេខ)",
      photo: "URL រូបភាព",
      active: "សកម្ម",
      inactive: "អសកម្ម",
      confirmDel: "តើអ្នកប្រាកដទេថាចង់លុបបុគ្គលិកនេះ?",
      noEmp: "មិនទាន់មានបុគ្គលិកទេ",
    },
    sh: {
      addBtn: "បន្ថែមវេន",
      editTitle: "កែសម្រួលវេន",
      addTitle: "បន្ថែមវេន",
      nameLabel: "ឈ្មោះវេន",
      startLabel: "ម៉ោងចូល",
      endLabel: "ម៉ោងចេញ",
      graceLabel: "ផ្តល់ (នាទី)",
      confirmDel: "តើអ្នកប្រាកដទេថាចង់លុបវេននេះ?",
      noShift: "មិនទាន់មានវេនទេ",
      assignedEmp: "បុគ្គលិកប្រើវេននេះ",
    },
    att: {
      checkIn: "ចូលធ្វើការ",
      checkOut: "ចេញធ្វើការ",
      present: "មានវត្តមាន",
      absent: "អវត្តមាន",
      late: "មកយឺត",
      onLeave: "ច្បាប់",
      date: "កាលបរិច្ឆេទ",
      inTime: "ម៉ោងចូល",
      outTime: "ម៉ោងចេញ",
      setOffice: "កំណត់ទីតាំងការិយាល័យ",
      noRecord: "មិនទាន់មានកំណត់ត្រាទេ",
      absentDays: "ថ្ងៃអវត្តមាន",
    },
    lv: {
      addBtn: "សំណើច្បាប់ថ្មី",
      type: "ប្រភេទច្បាប់",
      startDate: "ថ្ងៃចាប់ផ្តើម",
      endDate: "ថ្ងៃបញ្ចប់",
      reason: "មូលហេតុ",
      pending: "កំពុងរង់ចាំ",
      approved: "បានអនុម័ត",
      rejected: "បានបដិសេធ",
      approve: "អនុម័ត",
      reject: "បដិសេធ",
      confirmDel: "តើអ្នកប្រាកដទេថាចង់លុបសំណើនេះ?",
      noRequest: "មិនទាន់មានសំណើទេ",
    },
    ot: {
      addBtn: "សុំ OT ថ្មី",
      date: "កាលបរិច្ឆេទ",
      hours: "ចំនួនម៉ោង OT",
      dayType: "ប្រភេទថ្ងៃ",
      dtNormal: "ថ្ងៃធម្មតា",
      dtWeekend: "ថ្ងៃឈប់សម្រាក/ចុងសប្តាហ៍",
      dtHoliday: "ថ្ងៃបុណ្យជាតិ",
      reason: "មូលហេតុស្នើសុំ",
      reasonPlaceholder: "សរសេរមូលហេតុសង្ខេប...",
      submit: "ដាក់ស្នើសំណើ",
      approve: "អនុម័ត",
      reject: "បដិសេធ",
      approvedBy: "អនុម័តដោយ",
      rejectedBy: "បដិសេធដោយ",
      rejectTitle: "បដិសេធសំណើ OT",
      rejectReason: "មូលហេតុបដិសេធ",
      rejectReasonPlaceholder: "សូមបញ្ជាក់មូលហេតុបដិសេធ...",
      rejectReasonRequired: "សូមបញ្ចូលមូលហេតុបដិសេធ",
      confirmDel: "តើអ្នកប្រាកដទេថាចង់លុបសំណើ OT នេះ?",
      noRequest: "មិនទាន់មានសំណើ OT ទេ",
      hoursShort: "ម៉ោង",
      policyTitle: "គោលការណ៍អត្រា OT",
      policyDesc:
        "កំណត់អត្រាគុណប្រាក់ OT សម្រាប់ថ្ងៃធម្មតា ថ្ងៃឈប់សម្រាក និងថ្ងៃបុណ្យជាតិ ព្រមទាំងចំនួនម៉ោងធ្វើការស្តង់ដារក្នុងមួយថ្ងៃ។",
      rateNormal: "អត្រាគុណ · ថ្ងៃធម្មតា",
      rateWeekend: "អត្រាគុណ · ថ្ងៃឈប់សម្រាក",
      rateHoliday: "អត្រាគុណ · ថ្ងៃបុណ្យជាតិ",
      hoursPerDay: "ម៉ោងធ្វើការស្តង់ដារ/ថ្ងៃ",
      totalOtHours: "ម៉ោង OT សរុប",
    },
    pay: {
      baseSalary: "ប្រាក់ខែមូលដ្ឋាន",
      netSalary: "ប្រាក់ខែសុទ្ធ",
      markPaid: "បើកប្រាក់ខែ",
      paid: "បានបើក",
      unpaid: "មិនទាន់បើក",
      totalPaid: "ចំណាយប្រាក់ខែសរុប",
      absentDed: "ថ្ងៃអវត្តមាន",
      otPay: "ប្រាក់ OT",
      viewSlip: "មើលសន្លឹកប្រាក់ខែ",
      unmarkPaid: "ដកសញ្ញាបានបើក",
      taxLabel: "ពន្ធលើប្រាក់ខែ",
      insuranceLabel: "ធានារ៉ាប់រង",
      policyTitle: "គោលការណ៍កាត់ប្រាក់ខែ",
      policyDesc:
        "កំណត់អត្រាភាគរយពន្ធលើប្រាក់ខែ និងធានារ៉ាប់រង ដែលនឹងកាត់ចេញពីប្រាក់ខែមូលដ្ឋានរបស់បុគ្គលិកទាំងអស់។",
      taxRateLabel: "អត្រាពន្ធលើប្រាក់ខែ (%)",
      insuranceRateLabel: "អត្រាធានារ៉ាប់រង (%)",
    },
    admAcc: {
      addBtn: "បន្ថែមអ្នកគ្រប់គ្រង",
      editTitle: "កែសម្រួលអ្នកគ្រប់គ្រង",
      addTitle: "បន្ថែមអ្នកគ្រប់គ្រង",
      roleLabel: "តួនាទី",
      confirmDel: "តើអ្នកប្រាកដទេថាចង់លុបគណនីនេះ?",
    },
    profile: {
      title: "ប្រវត្តិរូបរបស់ខ្ញុំ",
      editBtn: "កែប្រវត្តិរូប",
      changePin: "ផ្លាស់ប្តូរ PIN",
      oldPin: "PIN ចាស់",
      newPin: "PIN ថ្មី",
    },
    qr: {
      title: "QR Code សម្រាប់បុគ្គលិក",
      desc: "ស្កេនដើម្បីបើកទំព័រចូលប្រើសម្រាប់បុគ្គលិកដោយផ្ទាល់",
    },
    settings: {
      title: "ការកំណត់គណនី",
      photoLabel: "រូបភាពប្រវត្តិរូប",
      choosePhoto: "ជ្រើសរើសរូបភាព",
      nameLabel: "ឈ្មោះ",
      namePlaceholder: "បញ្ចូលឈ្មោះរបស់អ្នក",
      appearance: "រូបរាងទំព័រ",
      lightMode: "ពន្លឺ (Light)",
      darkMode: "ងងឹត (Dark)",
      appearanceDesc: "ប្តូររូបរាងទំព័រសម្រាប់ឧបករណ៍នេះ",
      saved: "បានរក្សាទុកដោយជោគជ័យ",
      nameRequired: "សូមបញ្ចូលឈ្មោះ",
      brandingTitle: "ម៉ាកយីហោក្រុមហ៊ុន",
      brandingDesc:
        "ដាក់ឈ្មោះក្រុមហ៊ុន និងឡូហ្គោផ្ទាល់ខ្លួន ដើម្បីជំនួសឈ្មោះ និងឡូហ្គោលំនាំដើម",
      companyNameLabel: "ឈ្មោះក្រុមហ៊ុន",
      companyNamePlaceholder: "Workforce Suite",
      companyLogoLabel: "ឡូហ្គោក្រុមហ៊ុន",
      chooseLogo: "ជ្រើសរើសឡូហ្គោ",
      removeLogo: "លុបឡូហ្គោ",
      brandingSaved: "បានរក្សាទុកម៉ាកយីហោដោយជោគជ័យ",
    },
  },
  en: {
    appName: "Workforce Suite",
    login: {
      employeeId: "Employee ID",
      employeeIdPlaceholder: "EMP-001",
      pin: "PIN Code",
      pinPlaceholder: "••••",
      submit: "Sign In",
      adminSubmit: "Sign In as Admin",
      username: "Username",
      usernamePlaceholder: "admin",
      password: "Admin Password",
      passwordPlaceholder: "••••••••",
      adminTitle: "Admin Login",
      employeePortal: "Employee Portal",
      demoLabel: "Demo credentials:",
      errNoEmp: "Employee ID not found",
      errInactive: "This account is inactive. Contact Admin.",
      errPin: "Incorrect PIN code",
      errNoAdmin: "Admin account not found",
      errPass: "Incorrect password",
      back: "Back",
      switchToAdmin: "Admin? Sign in here",
    },
    nav: {
      dashboard: "Dashboard",
      employees: "Employees",
      departments: "Departments",
      shifts: "Shifts",
      attendance: "Attendance",
      leave: "Leave Requests",
      overtime: "Overtime (OT)",
      payroll: "Payroll",
      admins: "Admin Accounts",
      myAttendance: "My Attendance",
      myLeave: "My Leave",
      myOvertime: "My Overtime (OT)",
      myPayroll: "My Payroll",
      myProfile: "My Profile",
      settings: "Settings",
    },
    logout: "Sign Out",
    notifications: "Notifications",
    markAllRead: "Mark all as read",
    noNotif: "No notifications",
    confirmDelete: "Confirm Delete",
    cancel: "Cancel",
    delete: "Delete",
    edit: "Edit",
    save: "Save",
    add: "Add",
    search: "Search...",
    employee: "Employee",
    status: "Status",
    actions: "Actions",
    noData: "No data",
    dash: {
      welcome: "Welcome",
      totalEmp: "Total Employees",
      active: "active",
      totalDept: "Departments",
      totalDeptSub: "Total departments",
      presentToday: "Present Today",
      attendRate: "Attendance rate",
      pendingPayroll: "Pending Payroll",
      thisMonth: "This month",
      recentAttend: "Today\'s Attendance",
      noAttend: "No attendance records today",
      empPortalLink: "Employee Portal Link",
      showQR: "Show QR Code",
      copyLink: "Copy Link",
      copied: "Copied!",
      noEmpWarn: 'No employees yet. Go to "Employees" to add one.',
      myDept: "My Department",
      myShift: "My Shift",
      todayStatus: "Today's Status",
      payrollStatus: "Payroll Status",
      notCheckedIn: "Not checked in",
    },
    depts: {
      addBtn: "Add Department",
      editTitle: "Edit Department",
      addTitle: "Add Department",
      nameLabel: "Department Name",
      codeLabel: "Short Code",
      descLabel: "Description",
      staffCount: "staff",
      noDesc: "No description",
      confirmDel: "Are you sure you want to delete this department?",
    },
    emps: {
      addBtn: "Add Employee",
      editTitle: "Edit Employee",
      addTitle: "Add Employee",
      name: "Name",
      code: "Employee ID",
      dept: "Department",
      shift: "Shift",
      role: "Role",
      salary: "Salary",
      phone: "Phone",
      email: "Email",
      joined: "Join Date",
      pin: "PIN (digits)",
      photo: "Photo URL",
      active: "Active",
      inactive: "Inactive",
      confirmDel: "Are you sure you want to delete this employee?",
      noEmp: "No employees yet",
    },
    sh: {
      addBtn: "Add Shift",
      editTitle: "Edit Shift",
      addTitle: "Add Shift",
      nameLabel: "Shift Name",
      startLabel: "Start Time",
      endLabel: "End Time",
      graceLabel: "Grace (minutes)",
      confirmDel: "Are you sure you want to delete this shift?",
      noShift: "No shifts yet",
      assignedEmp: "Employees using this shift",
    },
    att: {
      checkIn: "Check In",
      checkOut: "Check Out",
      present: "Present",
      absent: "Absent",
      late: "Late",
      onLeave: "On Leave",
      date: "Date",
      inTime: "In Time",
      outTime: "Out Time",
      setOffice: "Set Office Location",
      noRecord: "No records yet",
      absentDays: "Absent days",
    },
    lv: {
      addBtn: "New Leave Request",
      type: "Leave Type",
      startDate: "Start Date",
      endDate: "End Date",
      reason: "Reason",
      pending: "Pending",
      approved: "Approved",
      rejected: "Rejected",
      approve: "Approve",
      reject: "Reject",
      confirmDel: "Are you sure you want to delete this request?",
      noRequest: "No requests yet",
    },
    ot: {
      addBtn: "New OT Request",
      date: "Date",
      hours: "OT Hours",
      dayType: "Day Type",
      dtNormal: "Normal Working Day",
      dtWeekend: "Day Off / Weekend",
      dtHoliday: "Public Holiday",
      reason: "Reason",
      reasonPlaceholder: "Briefly explain the reason...",
      submit: "Submit Request",
      approve: "Approve",
      reject: "Reject",
      approvedBy: "Approved by",
      rejectedBy: "Rejected by",
      rejectTitle: "Reject OT Request",
      rejectReason: "Rejection Reason",
      rejectReasonPlaceholder: "Please state the reason for rejecting...",
      rejectReasonRequired: "Please enter a rejection reason",
      confirmDel: "Are you sure you want to delete this OT request?",
      noRequest: "No OT requests yet",
      hoursShort: "hrs",
      policyTitle: "Overtime Rate Policy",
      policyDesc:
        "Set the OT pay multiplier for normal working days, days off, and public holidays, plus standard working hours per day.",
      rateNormal: "Multiplier · Normal Day",
      rateWeekend: "Multiplier · Day Off",
      rateHoliday: "Multiplier · Public Holiday",
      hoursPerDay: "Standard Hours / Day",
      totalOtHours: "Total OT Hours",
    },
    pay: {
      baseSalary: "Base Salary",
      netSalary: "Net Salary",
      markPaid: "Mark as Paid",
      paid: "Paid",
      unpaid: "Unpaid",
      totalPaid: "Total Payroll",
      absentDed: "Absent days",
      otPay: "OT Pay",
      viewSlip: "View Payslip",
      unmarkPaid: "Unmark as Paid",
      taxLabel: "Income Tax",
      insuranceLabel: "Insurance",
      policyTitle: "Payroll Deduction Policy",
      policyDesc:
        "Set the tax and insurance percentage rates deducted from every employee's base salary.",
      taxRateLabel: "Tax Rate (%)",
      insuranceRateLabel: "Insurance Rate (%)",
    },
    admAcc: {
      addBtn: "Add Admin",
      editTitle: "Edit Admin",
      addTitle: "Add Admin",
      roleLabel: "Role",
      confirmDel: "Are you sure you want to delete this account?",
    },
    profile: {
      title: "My Profile",
      editBtn: "Edit Profile",
      changePin: "Change PIN",
      oldPin: "Old PIN",
      newPin: "New PIN",
    },
    qr: {
      title: "Employee Portal QR Code",
      desc: "Scan to open the employee login portal directly",
    },
    settings: {
      title: "Account Settings",
      photoLabel: "Profile Photo",
      choosePhoto: "Choose Photo",
      nameLabel: "Name",
      namePlaceholder: "Enter your name",
      appearance: "Appearance",
      lightMode: "Light",
      darkMode: "Dark",
      appearanceDesc: "Switch the theme for this device",
      saved: "Saved successfully",
      nameRequired: "Please enter a name",
      brandingTitle: "Company Branding",
      brandingDesc:
        "Set a custom company name and logo to replace the default name and logo",
      companyNameLabel: "Company Name",
      companyNamePlaceholder: "Workforce Suite",
      companyLogoLabel: "Company Logo",
      chooseLogo: "Choose Logo",
      removeLogo: "Remove Logo",
      brandingSaved: "Branding saved successfully",
    },
  },
};

const LangContext = createContext({
  lang: "km",
  t: LANG.km,
  setLang: () => {},
});
const useLang = () => useCtx(LangContext);

const ThemeContext = createContext({
  theme: "light",
  setTheme: () => {},
  toggleTheme: () => {},
});
const useTheme = () => useCtx(ThemeContext);

// Company branding (custom app name + logo), editable from Admin Settings
// by a super admin and shared across the login screens, sidebar, etc.
const BrandingContext = createContext({
  branding: { name: "", logo: null },
  setBranding: () => {},
});
const useBranding = () => useCtx(BrandingContext);

function getInitials(name) {
  const parts = (name || "").trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "WS";
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[1][0]).toUpperCase();
}

function ThemeToggle({ variant = "dark" }) {
  const { theme, toggleTheme } = useTheme();
  const isDark = theme === "dark";
  const chromeIsDark = variant === "dark";
  return (
    <button
      onClick={toggleTheme}
      title="Dark mode / របៀបងងឹត"
      aria-label="Toggle dark mode"
      style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        width: 30,
        height: 30,
        background: chromeIsDark ? "rgba(255,255,255,0.10)" : "var(--wf-card)",
        border: `1px solid ${chromeIsDark ? "rgba(255,255,255,0.18)" : "var(--wf-line)"}`,
        borderRadius: 8,
        color: chromeIsDark ? "#fff" : "var(--wf-ink)",
        cursor: "pointer",
        transition: "background .15s, border-color .15s",
        flexShrink: 0,
      }}
    >
      {isDark ? <Sun size={15} /> : <Moon size={15} />}
    </button>
  );
}

function LangToggle({ variant = "dark" }) {
  const { lang, setLang } = useLang();
  const isDark = variant === "dark";
  return (
    <button
      onClick={() => setLang(lang === "km" ? "en" : "km")}
      title="Switch language / ប្តូរភាសា"
      style={{
        display: "flex",
        alignItems: "center",
        gap: 5,
        background: isDark ? "rgba(255,255,255,0.10)" : T.paper,
        border: `1px solid ${isDark ? "rgba(255,255,255,0.18)" : T.line}`,
        borderRadius: 8,
        color: isDark ? "#fff" : T.ink,
        fontSize: 12,
        fontWeight: 600,
        padding: "5px 10px",
        cursor: "pointer",
        transition: "background .15s, border-color .15s",
        backdropFilter: "blur(4px)",
      }}
    >
      {lang === "km" ? "🇬🇧 EN" : "🇰🇭 KM"}
    </button>
  );
}

/* ---------------------------------------------------------------
   Tokens
----------------------------------------------------------------*/
// BRAND is the fixed navy used for "always-dark" chrome — the sidebar,
// primary buttons, and hero banners — which stays the same regardless of
// light/dark mode (white text sits on it either way).
const BRAND = {
  ink: "#12203D",
  inkDark: "#0B1730",
};
// T holds the tokens that DO change between light/dark mode. Each value is
// a CSS custom property reference — the actual light/dark hex values are
// defined once in the injected stylesheet (:root and .wf-dark), so every
// existing `T.xxx` usage across the app repaints automatically when the
// theme toggles, with no per-usage changes needed.
const T = {
  ink: "var(--wf-ink)",
  inkDark: "var(--wf-ink-dark)",
  paper: "var(--wf-paper)",
  card: "var(--wf-card)",
  forest: "#2E6F4E",
  forestDark: "#245A3F",
  forestSoft: "var(--wf-forest-soft)",
  forestText: "var(--wf-forest-text)",
  clay: "#B5502F",
  gold: "#C08A2E",
  goldSoft: "var(--wf-gold-soft)",
  goldText: "var(--wf-gold-text)",
  rose: "#A93E4C",
  roseDark: "var(--wf-rose-dark)",
  roseSoft: "var(--wf-rose-soft)",
  blue: "#3E5C8A",
  line: "var(--wf-line)",
  lineSoft: "var(--wf-line-soft)",
  muted: "var(--wf-muted)",
  mutedLight: "var(--wf-muted-light)",
  text: "var(--wf-ink)",
  textSoft: "var(--wf-text-soft)",
  inputBorder: "var(--wf-input-border)",
  inputBg: "var(--wf-input-bg)",
  fieldLabel: "var(--wf-field-label)",
  tableHeadBg: "var(--wf-table-head-bg)",
  divider: "var(--wf-divider)",
  dangerBorder: "var(--wf-danger-border)",
  dangerHoverBg: "var(--wf-danger-hover-bg)",
  headerBg: "var(--wf-header-bg)",
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
*,*::before,*::after{box-sizing:border-box;}
html,body,#root{height:100%;}
:root{
  --wf-ink:#12203D; --wf-ink-dark:#0B1730; --wf-paper:#F5F2EA; --wf-card:#FFFFFF;
  --wf-forest-soft:#E8F2EC; --wf-forest-text:#215D3F; --wf-gold-soft:#FBF1DF; --wf-gold-text:#8A5E14;
  --wf-rose-dark:#8C3140; --wf-rose-soft:#F3E9E9; --wf-line:#E7E2D6; --wf-line-soft:#EEE9DC;
  --wf-muted:#8A8577; --wf-muted-light:#B0AA98; --wf-text-soft:#4A4638;
  --wf-input-border:#D8D2C2; --wf-input-bg:#FDFCF9; --wf-field-label:#6B6455;
  --wf-table-head-bg:#FAF8F2; --wf-divider:#F0EDE2; --wf-danger-border:#E4C7CB;
  --wf-danger-hover-bg:#F8ECEE; --wf-header-bg:rgba(255,255,255,0.92);
}
.wf-dark{
  --wf-ink:#E9ECF4; --wf-ink-dark:#B7BFD4; --wf-paper:#0E1526; --wf-card:#161F35;
  --wf-forest-soft:#173226; --wf-forest-text:#7FD9A8; --wf-gold-soft:#332715; --wf-gold-text:#E8C067;
  --wf-rose-dark:#F4A6B0; --wf-rose-soft:#3A2126; --wf-line:#2A3350; --wf-line-soft:#233049;
  --wf-muted:#8D96B3; --wf-muted-light:#5B6486; --wf-text-soft:#C3C9DC;
  --wf-input-border:#2E3A5C; --wf-input-bg:#101A30; --wf-field-label:#98A1C0;
  --wf-table-head-bg:#131C32; --wf-divider:#232D4A; --wf-danger-border:#5C2C36;
  --wf-danger-hover-bg:#2C1820; --wf-header-bg:rgba(14,21,38,0.92);
}
.wf-root{display:flex;height:100vh;height:100dvh;min-height:640px;max-height:100vh;max-height:100dvh;background:${T.paper};font-family:'Inter',sans-serif;color:${T.text};position:relative;overflow:hidden;border-radius:16px;box-shadow:0 1px 2px rgba(18,32,61,0.05),0 20px 48px -16px rgba(18,32,61,0.22);transition:background .15s ease,color .15s ease;}
.wf-sidebar{background:linear-gradient(175deg,${BRAND.ink} 0%,${BRAND.inkDark} 100%);color:#fff;width:250px;flex-shrink:0;display:flex;flex-direction:column;transition:transform .25s cubic-bezier(.4,0,.2,1);}
.wf-sidebar-inner{display:flex;flex-direction:column;height:100%;}
.wf-logo-badge{width:34px;height:34px;border-radius:10px;background:linear-gradient(145deg,${T.forest},${T.forestDark});display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;font-family:'Space Grotesk',sans-serif;flex-shrink:0;box-shadow:0 2px 8px rgba(46,111,78,0.4),inset 0 1px 0 rgba(255,255,255,0.18);}
.wf-nav-item{position:relative;width:100%;display:flex;align-items:center;gap:11px;padding:10px 14px;border-radius:10px;font-size:13.5px;font-weight:500;background:transparent;color:#AEB8CC;border:none;cursor:pointer;text-align:left;transition:background .15s ease,color .15s ease;}
.wf-nav-item:hover{background:rgba(255,255,255,0.07);color:#fff;}
.wf-nav-item.active{background:rgba(255,255,255,0.09);color:#fff;font-weight:600;}
.wf-nav-item.active::before{content:"";position:absolute;left:-10px;top:8px;bottom:8px;width:3px;border-radius:0 3px 3px 0;background:${T.gold};}
.wf-main{flex:1;min-width:0;min-height:0;display:flex;flex-direction:column;background:${T.paper};}
.wf-header{background:${T.headerBg};backdrop-filter:blur(8px);border-bottom:1px solid ${T.lineSoft};padding:14px 22px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:20;transition:background .15s ease,border-color .15s ease;}
.wf-content{flex:1;overflow-y:auto;padding:22px;}
.wf-card{background:${T.card};border-radius:14px;border:1px solid ${T.line};box-shadow:0 1px 2px rgba(18,32,61,0.04);transition:box-shadow .15s ease,background .15s ease,border-color .15s ease;}
.wf-btn{display:inline-flex;align-items:center;gap:6px;font-weight:600;border-radius:10px;font-size:13px;padding:9px 15px;border:1px solid transparent;cursor:pointer;transition:background .15s ease,transform .1s ease,box-shadow .15s ease;}
.wf-btn:active:not(:disabled){transform:scale(.97);}
.wf-btn:disabled{opacity:.5;cursor:not-allowed;}
.wf-btn-sm{padding:6px 10px;font-size:12px;}
.wf-btn-primary{background:${BRAND.ink};color:#fff;box-shadow:0 1px 2px rgba(18,32,61,0.18);}
.wf-btn-primary:hover:not(:disabled){background:${BRAND.inkDark};}
.wf-btn-accent{background:${T.forest};color:#fff;box-shadow:0 1px 2px rgba(46,111,78,0.22);}
.wf-btn-accent:hover:not(:disabled){background:${T.forestDark};}
.wf-btn-ghost{background:transparent;color:${T.ink};border-color:${T.line};}
.wf-btn-ghost:hover:not(:disabled){background:${T.paper};}
.wf-btn-danger{background:transparent;color:${T.rose};border-color:${T.dangerBorder};}
.wf-btn-danger:hover:not(:disabled){background:${T.dangerHoverBg};}
.wf-btn-danger-solid{background:${T.rose};color:#fff;}
.wf-btn-danger-solid:hover:not(:disabled){background:${T.roseDark};}
.wf-input{width:100%;padding:9px 12px;border-radius:10px;border:1px solid ${T.inputBorder};font-size:13px;background:${T.inputBg};color:${T.text};outline:none;font-family:inherit;transition:border-color .15s ease,box-shadow .15s ease,background .15s ease;}
.wf-input:focus{border-color:${T.forest};box-shadow:0 0 0 3px rgba(46,111,78,0.15);}
.wf-field-label{display:block;font-size:11px;font-weight:700;color:${T.fieldLabel};margin-bottom:6px;text-transform:uppercase;letter-spacing:.04em;}
.wf-modal-overlay{position:fixed;inset:0;z-index:50;display:flex;align-items:center;justify-content:center;padding:16px;background:rgba(8,12,24,0.55);backdrop-filter:blur(2px);animation:wf-fade .15s ease;}
.wf-modal{background:${T.card};border-radius:16px;box-shadow:0 24px 64px rgba(18,32,61,0.35);width:100%;max-height:90vh;overflow-y:auto;animation:wf-pop .18s cubic-bezier(.2,.9,.3,1.2);}
.wf-modal-head{display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid ${T.lineSoft};position:sticky;top:0;background:${T.card};border-radius:16px 16px 0 0;}
.wf-avatar{border-radius:999px;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:600;flex-shrink:0;}
.wf-badge{display:inline-block;padding:4px 10px;border-radius:8px;font-size:11px;font-weight:700;letter-spacing:.02em;white-space:nowrap;}
.wf-table{width:100%;font-size:13px;border-collapse:collapse;}
.wf-table th{text-align:left;font-size:11px;color:${T.muted};text-transform:uppercase;padding:11px 16px;background:${T.tableHeadBg};border-bottom:1px solid ${T.lineSoft};font-weight:700;letter-spacing:.03em;}
.wf-table td{padding:10px 16px;border-bottom:1px solid ${T.paper};}
.wf-table tr:last-child td{border-bottom:none;}
.wf-table tbody tr{transition:background .12s ease;}
.wf-table tbody tr:hover{background:${T.tableHeadBg};}
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
// Default OT policy used until an admin saves one in Supabase (ot_policy
// table, single row id=1). Multipliers apply to the derived hourly rate.
const DEFAULT_OT_POLICY = {
  rateNormal: 1.5,
  rateWeekend: 2,
  rateHoliday: 3,
  hoursPerDay: 8,
};
// Maps an OT request's day type to the matching multiplier key in policy.
const OT_RATE_KEY = {
  normal: "rateNormal",
  weekend: "rateWeekend",
  holiday: "rateHoliday",
};
// Default payroll deduction policy used until an admin saves one in
// Supabase (payroll_policy table, single row id=1). Rates are percentages
// (e.g. 5 means 5%) applied to the adjusted base salary.
const DEFAULT_PAYROLL_POLICY = {
  taxRate: 5,
  insuranceRate: 2,
};
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
function hhmm(v) {
  return typeof v === "string" ? v.slice(0, 5) : v;
}
function shiftLabel(shift) {
  if (!shift) return "—";
  return `${shift.name} · ${hhmm(shift.start)}–${hhmm(shift.end)}`;
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
// Returns how many minutes past the shift's start time a check-in was, or
// 0 if the check-in was on time or early. Used to show "late by X" detail.
function lateMinutesForShift(checkInTime, shift) {
  if (!checkInTime) return 0;
  const start = shift ? shift.start : "09:00";
  if (checkInTime <= start) return 0;
  const [ch, cm] = checkInTime.split(":").map(Number);
  const [sh, sm] = start.split(":").map(Number);
  return ch * 60 + cm - (sh * 60 + sm);
}
function formatLateDuration(mins, lang) {
  if (!mins || mins <= 0) return "";
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  if (lang === "en") {
    const parts = [];
    if (h) parts.push(`${h}h`);
    if (m || !h) parts.push(`${m}m`);
    return `Late by ${parts.join(" ")}`;
  }
  const parts = [];
  if (h) parts.push(`${h} ម៉ោង`);
  if (m || !h) parts.push(`${m} នាទី`);
  return `មកយឺត ${parts.join(" ")}`;
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
  overtimeRequests = [],
  lang = "km",
}) {
  const LEAVE_TYPE_LABEL = getLeaveTypeLabel(lang);
  const en = lang === "en";
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
    overtimeRequests
      .filter((r) => r.status === "pending")
      .forEach((r) => {
        const emp = employees.find((e) => e.id === r.employeeId);
        list.push({
          id: `ot-pending-${r.id}`,
          page: "ot",
          tone: "gold",
          title: en ? "New OT Request" : "សំណើសុំ OT ថ្មី",
          message: en
            ? `${emp?.name || "?"} requested ${r.hours}h OT on ${r.date}`
            : `${emp?.name || "?"} បានស្នើសុំ OT ចំនួន ${r.hours} ម៉ោង នៅថ្ងៃទី ${r.date}`,
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
          message: `${e.name} មិនទាន់ចុះឈ្មោះចូលធ្វើការទេ (${shift.name} ${hhmm(shift.start)})`,
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
    overtimeRequests
      .filter(
        (r) =>
          r.employeeId === currentEmp.id &&
          (r.status === "approved" || r.status === "rejected") &&
          r.reviewedAt,
      )
      .forEach((r) => {
        list.push({
          id: `ot-decided-${r.id}`,
          page: "ot",
          tone: r.status === "approved" ? "forest" : "rose",
          title:
            r.status === "approved"
              ? en
                ? "Your OT request was approved"
                : "សំណើសុំ OT របស់អ្នកត្រូវបានអនុម័ត"
              : en
                ? "Your OT request was rejected"
                : "សំណើសុំ OT របស់អ្នកត្រូវបានបដិសេធ",
          message: `${r.date} · ${r.hours}${en ? "h" : " ម៉ោង"}`,
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
// Given a list of offices [{id, name, lat, lng, radius}] and a coordinate,
// returns the closest office whose radius actually contains the point
// (i.e. { office, distance } for a valid punch), or null if the point
// falls outside every configured office's geofence.
function findMatchingOffice(offices, lat, lng) {
  let best = null;
  for (const o of offices) {
    const dist = distanceMeters(lat, lng, o.lat, o.lng);
    if (dist <= o.radius && (!best || dist < best.distance)) {
      best = { office: o, distance: Math.round(dist) };
    }
  }
  return best;
}
// Like findMatchingOffice but ignores the radius — used only to build a
// helpful "you're Xm from Office Y" message when no office matched.
function nearestOffice(offices, lat, lng) {
  let best = null;
  for (const o of offices) {
    const dist = distanceMeters(lat, lng, o.lat, o.lng);
    if (!best || dist < best.distance)
      best = { office: o, distance: Math.round(dist) };
  }
  return best;
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
// Suggests a default OT day type from a "YYYY-MM-DD" date: Saturday/Sunday
// suggest "weekend", everything else suggests "normal". Employees can still
// override this in the request form (e.g. for public holidays).
function suggestDayType(dateStr) {
  if (!dateStr) return "normal";
  const day = new Date(dateStr + "T00:00:00").getDay();
  return day === 0 || day === 6 ? "weekend" : "normal";
}
function getDayTypeLabel(lang, t) {
  return {
    normal: t.ot.dtNormal,
    weekend: t.ot.dtWeekend,
    holiday: t.ot.dtHoliday,
  };
}
// Sums approved OT hours/pay for one employee within a "YYYY-MM" month.
// Hourly rate is derived from monthly salary using the policy's standard
// working hours per day, then multiplied by the per-day-type OT rate.
function computeOvertimeForMonth(emp, overtimeRequests, mk, otPolicy) {
  const policy = otPolicy || DEFAULT_OT_POLICY;
  const hourlyRate =
    emp.salary / (WORKING_DAYS_PER_MONTH * (policy.hoursPerDay || 8));
  let otHours = 0;
  let otPay = 0;
  overtimeRequests
    .filter(
      (r) =>
        r.employeeId === emp.id &&
        r.status === "approved" &&
        r.date &&
        r.date.startsWith(mk),
    )
    .forEach((r) => {
      const mult = policy[OT_RATE_KEY[r.dayType]] ?? 1.5;
      otHours += Number(r.hours) || 0;
      otPay += (Number(r.hours) || 0) * hourlyRate * mult;
    });
  return { otHours, otPay, hourlyRate };
}
// Computes payroll figures for one employee for a given month, factoring in
// unpaid absences recorded in attendance and approved OT requests. Leave is
// paid and does not deduct; OT pay is added on top of net salary.
function computePayroll(
  emp,
  attendance,
  mk,
  overtimeRequests = [],
  otPolicy,
  payrollPolicy,
) {
  const policy = payrollPolicy || DEFAULT_PAYROLL_POLICY;
  const { absentDays, leaveDays, lateDays } = monthAttendanceStats(
    attendance,
    emp.id,
    mk,
  );
  const dailyRate = emp.salary / WORKING_DAYS_PER_MONTH;
  const absenceDeduction = Math.min(emp.salary, absentDays * dailyRate);
  const adjustedBase = emp.salary - absenceDeduction;
  const taxRate = Number(policy.taxRate) || 0;
  const insuranceRate = Number(policy.insuranceRate) || 0;
  const tax = adjustedBase * (taxRate / 100);
  const insurance = adjustedBase * (insuranceRate / 100);
  const { otHours, otPay } = computeOvertimeForMonth(
    emp,
    overtimeRequests,
    mk,
    otPolicy,
  );
  const net = adjustedBase - tax - insurance + otPay;
  return {
    absentDays,
    leaveDays,
    lateDays,
    dailyRate,
    absenceDeduction,
    adjustedBase,
    tax,
    insurance,
    taxRate,
    insuranceRate,
    otHours,
    otPay,
    net,
  };
}

/* ---------------------------------------------------------------
   Tiny hash router — keeps the employee portal and admin portal
   as two separate URLs (#/employee and #/admin) inside one app.
----------------------------------------------------------------*/
function normalizeHash(h) {
  const clean = (h || "").replace(/^#\/?/, "");
  const [portalPart, ...rest] = clean.split("/");
  const portal = portalPart === "employee" ? "employee" : "admin";
  const page = rest.filter(Boolean).join("/") || null;
  return { portal, page };
}
function usePortalRoute() {
  const [route, setRoute] = useState(() => normalizeHash(window.location.hash));
  useEffect(() => {
    const onHash = () => setRoute(normalizeHash(window.location.hash));
    window.addEventListener("hashchange", onHash);
    return () => window.removeEventListener("hashchange", onHash);
  }, []);
  // Switching portals (admin <-> employee) is a real navigation, so it
  // goes through location.hash like before and adds a history entry.
  const goPortal = useCallback((next) => {
    window.location.hash = next ? `/${next}` : "";
  }, []);
  // Remembering which page we're on inside a portal uses replaceState
  // instead of location.hash, so clicking around the sidebar doesn't
  // spam the browser's back button — but the page still survives a
  // refresh, because it's baked into the URL either way.
  const setPage = useCallback((page) => {
    setRoute((prev) => {
      const next = { ...prev, page };
      const hash = `#/${next.portal}${page ? `/${page}` : ""}`;
      window.history.replaceState(null, "", hash);
      return next;
    });
  }, []);
  return [route.portal, route.page, goPortal, setPage];
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

// branding is a single settings row (id = 1) shared by everyone — the
// company name + logo shown on the login screens and sidebar for every
// admin and employee, not just the device that set it.
function useBrandingSettings() {
  const [value, setValueState] = useState({ name: "", logo: null });
  const [ready, setReady] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const { data, error } = await supabase
        .from("branding")
        .select("*")
        .eq("id", 1)
        .maybeSingle();
      if (cancelled) return;
      if (error) {
        console.error("[supabase] failed to load branding:", error.message);
        setValueState({ name: "", logo: null });
      } else if (data) {
        setValueState({ name: data.name || "", logo: data.logo || null });
      } else {
        setValueState({ name: "", logo: null });
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
      const { error } = await supabase.from("branding").upsert({
        id: 1,
        name: next.name || "",
        logo: next.logo || null,
      });
      if (error)
        console.error("[supabase] save failed on branding:", error.message);
    })();
  }, []);

  return [value, setValue, ready];
}

// Login sessions are intentionally per-device, not shared data, so they
// stay in the browser's own localStorage instead of Supabase.
// ot_policy is a single settings row (id = 1), same shape as
// DEFAULT_OT_POLICY. Falls back to the defaults until an admin saves one.
function useOtPolicy() {
  const [value, setValueState] = useState(DEFAULT_OT_POLICY);
  const [ready, setReady] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const { data, error } = await supabase
        .from("ot_policy")
        .select("*")
        .eq("id", 1)
        .maybeSingle();
      if (cancelled) return;
      if (error) {
        console.error("[supabase] failed to load ot_policy:", error.message);
        setValueState(DEFAULT_OT_POLICY);
      } else if (data) {
        setValueState({
          rateNormal: data.rate_normal ?? DEFAULT_OT_POLICY.rateNormal,
          rateWeekend: data.rate_weekend ?? DEFAULT_OT_POLICY.rateWeekend,
          rateHoliday: data.rate_holiday ?? DEFAULT_OT_POLICY.rateHoliday,
          hoursPerDay: data.hours_per_day ?? DEFAULT_OT_POLICY.hoursPerDay,
        });
      } else {
        setValueState(DEFAULT_OT_POLICY);
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
      const { error } = await supabase.from("ot_policy").upsert({
        id: 1,
        rate_normal: next.rateNormal,
        rate_weekend: next.rateWeekend,
        rate_holiday: next.rateHoliday,
        hours_per_day: next.hoursPerDay,
      });
      if (error)
        console.error("[supabase] save failed on ot_policy:", error.message);
    })();
  }, []);

  return [value, setValue, ready];
}

// payroll_policy is a single settings row (id = 1), same shape as
// DEFAULT_PAYROLL_POLICY. Falls back to the defaults until an admin saves one.
function usePayrollPolicy() {
  const [value, setValueState] = useState(DEFAULT_PAYROLL_POLICY);
  const [ready, setReady] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const { data, error } = await supabase
        .from("payroll_policy")
        .select("*")
        .eq("id", 1)
        .maybeSingle();
      if (cancelled) return;
      if (error) {
        console.error(
          "[supabase] failed to load payroll_policy:",
          error.message,
        );
        setValueState(DEFAULT_PAYROLL_POLICY);
      } else if (data) {
        setValueState({
          taxRate: data.tax_rate ?? DEFAULT_PAYROLL_POLICY.taxRate,
          insuranceRate:
            data.insurance_rate ?? DEFAULT_PAYROLL_POLICY.insuranceRate,
        });
      } else {
        setValueState(DEFAULT_PAYROLL_POLICY);
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
      const { error } = await supabase.from("payroll_policy").upsert({
        id: 1,
        tax_rate: next.taxRate,
        insurance_rate: next.insuranceRate,
      });
      if (error)
        console.error(
          "[supabase] save failed on payroll_policy:",
          error.message,
        );
    })();
  }, []);

  return [value, setValue, ready];
}
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
function getStatusMap(lang) {
  const en = lang === "en";
  return {
    active: {
      bg: T.forestSoft,
      fg: T.forestText,
      label: en ? "Active" : "សកម្ម",
    },
    inactive: {
      bg: T.roseSoft,
      fg: T.roseDark,
      label: en ? "Inactive" : "អសកម្ម",
    },
    present: {
      bg: T.forestSoft,
      fg: T.forestText,
      label: en ? "Present" : "មកធ្វើការ",
    },
    late: { bg: T.goldSoft, fg: T.goldText, label: en ? "Late" : "មកយឺត" },
    absent: {
      bg: T.roseSoft,
      fg: T.roseDark,
      label: en ? "Absent" : "អវត្តមាន",
    },
    leave: { bg: "#E7ECF6", fg: T.blue, label: en ? "On Leave" : "ឈប់សម្រាក" },
    pending: {
      bg: T.goldSoft,
      fg: T.goldText,
      label: en ? "Pending" : "រង់ចាំបង់",
    },
    paid: { bg: T.forestSoft, fg: T.forestText, label: en ? "Paid" : "បង់រួច" },
    approved: {
      bg: T.forestSoft,
      fg: T.forestText,
      label: en ? "Approved" : "អនុម័តហើយ",
    },
    rejected: {
      bg: T.roseSoft,
      fg: T.roseDark,
      label: en ? "Rejected" : "បដិសេធ",
    },
  };
}
function getLeaveTypeLabel(lang) {
  const en = lang === "en";
  return {
    annual: en ? "Annual Leave" : "ច្បាប់ប្រចាំឆ្នាំ",
    sick: en ? "Sick Leave" : "ច្បាប់ឈឺ",
    other: en ? "Other" : "ផ្សេងៗ",
  };
}
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
  const { lang } = useLang();
  const STATUS_MAP = getStatusMap(lang);
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
// notification_reads stores, per user (admin or employee id), which
// notification ids they've already seen. This used to live in
// localStorage (so "read" state was per-device); now it's a shared table
// so marking a notification read syncs across every device/browser that
// account signs into.
function useNotificationReadIds(userId) {
  const [ids, setIdsState] = useState([]);
  const [ready, setReady] = useState(false);

  useEffect(() => {
    if (!userId) {
      setIdsState([]);
      setReady(true);
      return;
    }
    let cancelled = false;
    setReady(false);
    (async () => {
      const { data, error } = await supabase
        .from("notification_reads")
        .select("*")
        .eq("user_id", userId)
        .maybeSingle();
      if (cancelled) return;
      if (error) {
        console.error(
          "[supabase] failed to load notification_reads:",
          error.message,
        );
        setIdsState([]);
      } else if (data) {
        setIdsState(data.read_ids || []);
      } else {
        setIdsState([]);
      }
      setReady(true);
    })();
    return () => {
      cancelled = true;
    };
  }, [userId]);

  const setIds = useCallback(
    (next) => {
      setIdsState(next);
      if (!userId) return;
      (async () => {
        const { error } = await supabase.from("notification_reads").upsert({
          user_id: userId,
          read_ids: next,
        });
        if (error)
          console.error(
            "[supabase] save failed on notification_reads:",
            error.message,
          );
      })();
    },
    [userId],
  );

  return [ids, setIds, ready];
}

function NotificationBell({
  role,
  currentAdmin,
  currentEmp,
  employees,
  shifts,
  attendance,
  leaveRequests,
  overtimeRequests,
  setPage,
}) {
  const { t, lang } = useLang();
  const [open, setOpen] = useState(false);
  const boxRef = useRef(null);
  // Admins share one notification mailbox — same as every other admin
  // table in this app (leave requests, employees, etc.), so if Admin A
  // marks something read (or it disappears because the underlying record
  // was actioned/deleted), Admin B sees the same state on any device.
  // Employees still get their own mailbox, since their notifications are
  // personal (their own leave/OT decisions) — but it now syncs across
  // that employee's own devices too instead of staying on one browser.
  const userId = role === "admin" ? "admin_shared" : currentEmp?.id;
  const [readIds, setReadIds] = useNotificationReadIds(userId);

  const notifications = useMemo(
    () =>
      buildNotifications({
        role,
        currentEmp,
        employees,
        shifts,
        attendance,
        leaveRequests,
        overtimeRequests,
        lang,
      }),
    [
      role,
      currentEmp,
      employees,
      shifts,
      attendance,
      leaveRequests,
      overtimeRequests,
    ],
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
        aria-label={t.notifications}
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
              {t.notifications}
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
                {t.markAllRead}
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
              {t.noNotif}
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
  const { t } = useLang();
  return (
    <Modal title={t.confirmDelete} onClose={onCancel} width={380}>
      <p style={{ fontSize: 14, color: T.textSoft, marginBottom: 20 }}>
        {text}
      </p>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Button variant="ghost" onClick={onCancel}>
          {t.cancel}
        </Button>
        <Button variant="danger-solid" onClick={onConfirm}>
          <Trash2 size={14} /> {t.delete}
        </Button>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------
   Login screens — redesigned with smooth animations + language toggle
----------------------------------------------------------------*/
const LOGIN_CSS_ID = "wf-login-style";
const LOGIN_CSS = `
@keyframes wf-float-up { from { opacity:0; transform:translateY(28px) scale(.96); } to { opacity:1; transform:translateY(0) scale(1); } }
@keyframes wf-bg-drift { 0%,100% { transform:translate(0,0) scale(1.05); } 50% { transform:translate(-20px, -14px) scale(1.08); } }
@keyframes wf-bg-hue { 0%,100% { filter:hue-rotate(0deg); } 50% { filter:hue-rotate(12deg); } }
@keyframes wf-pulse-ring { 0%,100% { transform:scale(1); opacity:.5; } 50% { transform:scale(1.12); opacity:.2; } }
@keyframes wf-orb-a { 0%,100% { transform:translate(0,0) scale(1); } 33% { transform:translate(60px,-70px) scale(1.15); } 66% { transform:translate(-40px,40px) scale(.9); } }
@keyframes wf-orb-b { 0%,100% { transform:translate(0,0) scale(1); } 50% { transform:translate(-70px,60px) scale(1.2); } }
@keyframes wf-orb-c { 0%,100% { transform:translate(0,0) scale(1); } 40% { transform:translate(70px,50px) scale(.88); } 75% { transform:translate(-50px,-30px) scale(1.1); } }
@keyframes wf-orb-d { 0%,100% { transform:translate(0,0) scale(1); } 45% { transform:translate(-55px,-50px) scale(1.12); } }
.wf-login-root {
  display:flex; align-items:center; justify-content:center;
  min-height:100vh; min-height:100dvh; position:relative; overflow:hidden;
  background: linear-gradient(145deg, #0B1730 0%, #12203D 45%, #1a2e50 100%);
}
.wf-login-bg {
  position:absolute; inset:-40px; z-index:0; pointer-events:none;
  background: radial-gradient(ellipse 80% 60% at 20% 30%, rgba(46,111,78,0.18) 0%, transparent 60%),
              radial-gradient(ellipse 60% 80% at 80% 70%, rgba(62,92,138,0.15) 0%, transparent 60%),
              radial-gradient(ellipse 50% 40% at 60% 10%, rgba(192,138,46,0.10) 0%, transparent 60%);
  animation: wf-bg-drift 14s ease-in-out infinite, wf-bg-hue 22s ease-in-out infinite;
  will-change: transform, filter;
}
.wf-login-orbs { position:absolute; inset:0; z-index:0; pointer-events:none; overflow:hidden; }
.wf-login-orb {
  position:absolute; border-radius:50%; filter:blur(50px); opacity:.55;
  will-change: transform;
}
.wf-login-orb-1 {
  width:260px; height:260px; top:8%; left:8%;
  background:radial-gradient(circle,rgba(46,111,78,0.55),transparent 70%);
  animation: wf-orb-a 16s ease-in-out infinite;
}
.wf-login-orb-2 {
  width:320px; height:320px; bottom:6%; right:6%;
  background:radial-gradient(circle,rgba(62,92,138,0.5),transparent 70%);
  animation: wf-orb-b 20s ease-in-out infinite;
}
.wf-login-orb-3 {
  width:200px; height:200px; top:55%; left:2%;
  background:radial-gradient(circle,rgba(192,138,46,0.4),transparent 70%);
  animation: wf-orb-c 18s ease-in-out infinite;
}
.wf-login-orb-4 {
  width:180px; height:180px; top:4%; right:16%;
  background:radial-gradient(circle,rgba(140,49,64,0.35),transparent 70%);
  animation: wf-orb-d 15s ease-in-out infinite;
}
.wf-login-card {
  position:relative; z-index:2;
  width:100%; max-width:420px; margin:16px;
  padding:36px 32px 28px;
  background:rgba(255,255,255,0.97);
  backdrop-filter:blur(24px);
  border-radius:24px;
  box-shadow: 0 4px 24px rgba(0,0,0,0.08), 0 32px 80px rgba(0,0,0,0.32), inset 0 1px 0 rgba(255,255,255,0.9);
  animation: wf-float-up .6s cubic-bezier(.16,.9,.28,1) both;
}
.wf-login-logo-ring {
  width:64px; height:64px; border-radius:20px;
  background:linear-gradient(145deg,${T.forest},${T.forestDark});
  display:flex; align-items:center; justify-content:center;
  font-weight:800; font-size:20px; color:#fff;
  font-family:'Space Grotesk',sans-serif;
  box-shadow:0 4px 20px rgba(46,111,78,0.45), inset 0 1px 0 rgba(255,255,255,0.2);
  position:relative;
}
.wf-login-logo-ring::before {
  content:''; position:absolute; inset:-8px; border-radius:28px;
  border:2px solid rgba(46,111,78,0.25);
  animation: wf-pulse-ring 3s ease-in-out infinite;
}
.wf-login-input {
  width:100%; padding:12px 14px; border-radius:12px;
  border:1.5px solid #E2DDD4; font-size:14px;
  background:#FDFCFA; color:${T.text}; outline:none;
  font-family:inherit; transition:border-color .25s ease, box-shadow .25s ease, background .25s ease, transform .15s ease;
  box-sizing:border-box;
}
.wf-login-input:focus {
  border-color:${T.forest}; background:#fff;
  box-shadow:0 0 0 4px rgba(46,111,78,0.12);
  transform:translateY(-1px);
}
.wf-login-btn {
  width:100%; padding:13px; border:none; border-radius:14px;
  font-size:15px; font-weight:700; cursor:pointer; display:flex;
  align-items:center; justify-content:center; gap:8px;
  font-family:inherit; transition:transform .2s cubic-bezier(.2,.9,.3,1), box-shadow .2s ease, filter .2s ease;
  position:relative; overflow:hidden;
}
.wf-login-btn:hover:not(:disabled) { transform:translateY(-1px); }
.wf-login-btn:active:not(:disabled) { transform:scale(.97) translateY(0); }
.wf-login-btn-emp {
  background:linear-gradient(135deg,${T.forest} 0%,${T.forestDark} 100%);
  color:#fff;
  box-shadow:0 4px 16px rgba(46,111,78,0.35);
}
.wf-login-btn-emp:hover { box-shadow:0 6px 24px rgba(46,111,78,0.45); filter:brightness(1.06); }
.wf-login-btn-adm {
  background:linear-gradient(135deg,${BRAND.ink} 0%,${BRAND.inkDark} 100%);
  color:#fff;
  box-shadow:0 4px 16px rgba(18,32,61,0.35);
}
.wf-login-btn-adm:hover { box-shadow:0 6px 24px rgba(18,32,61,0.45); filter:brightness(1.1); }
.wf-login-divider { display:flex; align-items:center; gap:10px; margin:18px 0; }
.wf-login-divider::before,.wf-login-divider::after { content:''; flex:1; height:1px; background:#ECEAE3; }
.wf-login-error {
  display:flex; align-items:center; gap:7px; font-size:12.5px;
  color:${T.rose}; background:${T.roseSoft}; border-radius:10px;
  padding:9px 12px; margin-bottom:14px; border:1px solid #F0D4D8;
}
.wf-login-demo {
  margin-top:20px; padding:11px 14px; background:#F7F5EE;
  border-radius:12px; font-size:11px; color:${T.muted};
  line-height:1.7; border:1px solid #EDE9DF;
}
`;
function useLoginStyle() {
  useEffect(() => {
    if (!document.getElementById(LOGIN_CSS_ID)) {
      const tag = document.createElement("style");
      tag.id = LOGIN_CSS_ID;
      tag.innerHTML = LOGIN_CSS;
      document.head.appendChild(tag);
    }
  }, []);
}

// Layered, slowly-drifting gradient blobs behind the login card. Purely
// decorative (CSS animation only, no JS ticking) so it's cheap even on
// low-end phones, but gives the screen a living, non-static feel.
function LoginBackground() {
  return (
    <>
      <div className="wf-login-bg" />
      <div className="wf-login-orbs">
        <div className="wf-login-orb wf-login-orb-1" />
        <div className="wf-login-orb wf-login-orb-2" />
        <div className="wf-login-orb wf-login-orb-3" />
        <div className="wf-login-orb wf-login-orb-4" />
      </div>
    </>
  );
}

function LoginField({ label, children }) {
  return (
    <div style={{ marginBottom: 16 }}>
      <div
        style={{
          fontSize: 11.5,
          fontWeight: 700,
          color: "#6B6455",
          marginBottom: 7,
          textTransform: "uppercase",
          letterSpacing: ".04em",
        }}
      >
        {label}
      </div>
      {children}
    </div>
  );
}

function EmployeeLoginScreen({ employees, onLogin, go }) {
  useLoginStyle();
  const { t } = useLang();
  const { branding } = useBranding();
  const displayName = branding.name?.trim() || t.appName;
  const L = t.login;
  const [code, setCode] = useState("");
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setLoading(true);
    await new Promise((r) => setTimeout(r, 320));
    const emp = employees.find(
      (x) => x.code.trim().toLowerCase() === code.trim().toLowerCase(),
    );
    if (!emp) {
      setError(L.errNoEmp);
      setLoading(false);
      return;
    }
    if (emp.status !== "active") {
      setError(L.errInactive);
      setLoading(false);
      return;
    }
    if ((emp.pin || "") !== pin.trim()) {
      setError(L.errPin);
      setLoading(false);
      return;
    }
    setError("");
    setLoading(false);
    onLogin(emp.id);
  };

  return (
    <div className="wf-login-root">
      <LoginBackground />
      <div style={{ position: "absolute", top: 16, right: 16, zIndex: 10 }}>
        <LangToggle />
      </div>
      <div className="wf-login-card">
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            marginBottom: 28,
          }}
        >
          <div className="wf-login-logo-ring">
            {branding.logo ? (
              <img
                src={branding.logo}
                alt={displayName}
                style={{
                  width: "100%",
                  height: "100%",
                  objectFit: "cover",
                  borderRadius: "inherit",
                }}
              />
            ) : (
              getInitials(displayName)
            )}
          </div>
          <div
            style={{
              marginTop: 14,
              fontFamily: "'Space Grotesk',sans-serif",
              fontWeight: 700,
              fontSize: 20,
              color: T.ink,
            }}
          >
            {displayName}
          </div>
          <div style={{ fontSize: 13, color: T.muted, marginTop: 4 }}>
            {L.employeePortal}
          </div>
        </div>
        <form onSubmit={submit}>
          <LoginField label={L.employeeId}>
            <input
              className="wf-login-input"
              value={code}
              onChange={(e) => {
                setCode(e.target.value);
                setError("");
              }}
              placeholder={L.employeeIdPlaceholder}
              autoFocus
            />
          </LoginField>
          <LoginField label={L.pin}>
            <input
              className="wf-login-input"
              value={pin}
              onChange={(e) => {
                setPin(e.target.value);
                setError("");
              }}
              placeholder={L.pinPlaceholder}
              type="password"
              inputMode="numeric"
              maxLength={6}
            />
          </LoginField>
          {error && (
            <div className="wf-login-error">
              <AlertCircle size={14} /> {error}
            </div>
          )}
          <button
            type="submit"
            className="wf-login-btn wf-login-btn-emp"
            disabled={loading}
          >
            {loading ? (
              <Loader2
                size={16}
                style={{ animation: "spin 1s linear infinite" }}
              />
            ) : (
              <LogIn size={16} />
            )}
            {loading ? "..." : L.submit}
          </button>
        </form>
      </div>
    </div>
  );
}

function AdminLoginScreen({ admins, onLogin, go }) {
  useLoginStyle();
  const { t } = useLang();
  const { branding } = useBranding();
  const displayName = branding.name?.trim() || t.appName;
  const L = t.login;
  const [username, setUsername] = useState("");
  const [adminPass, setAdminPass] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setLoading(true);
    await new Promise((r) => setTimeout(r, 320));
    const acct = admins.find(
      (a) => a.username.trim().toLowerCase() === username.trim().toLowerCase(),
    );
    if (!acct) {
      setError(L.errNoAdmin);
      setLoading(false);
      return;
    }
    if (acct.password !== adminPass) {
      setError(L.errPass);
      setLoading(false);
      return;
    }
    setError("");
    setLoading(false);
    onLogin(acct.id);
  };

  return (
    <div className="wf-login-root">
      <LoginBackground />
      <div style={{ position: "absolute", top: 16, right: 16, zIndex: 10 }}>
        <LangToggle />
      </div>
      <div className="wf-login-card">
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
            fontSize: 12.5,
            fontWeight: 600,
            marginBottom: 20,
            padding: 0,
          }}
        >
          <ArrowLeft size={14} /> {L.back}
        </button>
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            marginBottom: 28,
          }}
        >
          <div
            style={{
              width: 64,
              height: 64,
              borderRadius: 20,
              background: branding.logo
                ? "#fff"
                : "linear-gradient(145deg,#5C4B9E,#3D3070)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              boxShadow: "0 4px 20px rgba(92,75,158,0.4)",
              overflow: "hidden",
              border: branding.logo ? `1px solid ${T.line}` : "none",
            }}
          >
            {branding.logo ? (
              <img
                src={branding.logo}
                alt={displayName}
                style={{ width: "100%", height: "100%", objectFit: "cover" }}
              />
            ) : (
              <ShieldCheck size={28} color="#fff" />
            )}
          </div>
          <div
            style={{
              marginTop: 14,
              fontFamily: "'Space Grotesk',sans-serif",
              fontWeight: 700,
              fontSize: 20,
              color: T.ink,
            }}
          >
            {L.adminTitle}
          </div>
          <div style={{ fontSize: 13, color: T.muted, marginTop: 4 }}>
            {displayName}
          </div>
        </div>
        <form onSubmit={submit}>
          <LoginField label={L.username}>
            <input
              className="wf-login-input"
              value={username}
              onChange={(e) => {
                setUsername(e.target.value);
                setError("");
              }}
              placeholder={L.usernamePlaceholder}
              autoFocus
            />
          </LoginField>
          <LoginField label={L.password}>
            <input
              className="wf-login-input"
              value={adminPass}
              onChange={(e) => {
                setAdminPass(e.target.value);
                setError("");
              }}
              placeholder={L.passwordPlaceholder}
              type="password"
            />
          </LoginField>
          {error && (
            <div className="wf-login-error">
              <AlertCircle size={14} /> {error}
            </div>
          )}
          <button
            type="submit"
            className="wf-login-btn wf-login-btn-adm"
            disabled={loading}
          >
            {loading ? (
              <Loader2
                size={16}
                style={{ animation: "spin 1s linear infinite" }}
              />
            ) : (
              <ShieldCheck size={16} />
            )}
            {loading ? "..." : L.adminSubmit}
          </button>
        </form>
      </div>
    </div>
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
  const { t: t2 } = useLang();
  const qrSrc = `https://api.qrserver.com/v1/create-qr-code/?size=240x240&data=${encodeURIComponent(url)}`;
  return (
    <Modal title={t2.qr.title} onClose={onClose} width={340}>
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
          {t2.qr.desc}
        </p>
      </div>
    </Modal>
  );
}
function EmployeeLinkCard() {
  const { t } = useLang();
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
            {t.dash.empPortalLink}
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
          <QrCode size={14} /> {t.dash.showQR}
        </Button>
        <Button variant="accent" onClick={copyLink}>
          {copied ? <CheckCircle2 size={14} /> : <Copy size={14} />}{" "}
          {copied ? t.dash.copied : t.dash.copyLink}
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
  shifts,
}) {
  const { t, lang } = useLang();
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

  // Employees only ever see their own data on the dashboard — never
  // company-wide totals, which are admin/manager-only.
  const myDept = departments.find((d) => d.id === currentEmp?.deptId);
  const myShift = shifts?.find((s) => s.id === currentEmp?.shiftId);
  const myTodayRecord = attendance.find(
    (a) => a.date === today && a.employeeId === currentEmp?.id,
  );
  const STATUS_MAP = getStatusMap(lang);
  const myStatusLabel = myTodayRecord
    ? STATUS_MAP[myTodayRecord.status]?.label || myTodayRecord.status
    : t.dash.notCheckedIn;
  const myPayrollPaid = !!payrollPaid[`${currentEmp?.id}-${mk}`];

  const stats =
    role === "admin"
      ? [
          {
            label: t.dash.totalEmp,
            value: employees.length,
            sub: `${activeEmployees.length} ${t.dash.active}`,
            icon: Users,
            accent: T.forest,
          },
          {
            label: t.nav.departments,
            value: departments.length,
            sub: t.dash.totalDeptSub,
            icon: Building2,
            accent: T.blue,
          },
          {
            label: t.dash.presentToday,
            value: presentToday,
            sub: `${rate}% ${t.dash.attendRate}`,
            icon: Clock,
            accent: T.gold,
          },
          {
            label: t.dash.pendingPayroll,
            value: pendingPayroll,
            sub: t.dash.thisMonth,
            icon: Wallet,
            accent: T.rose,
          },
        ]
      : [
          {
            label: t.dash.myDept,
            value: myDept?.name || "—",
            sub: "",
            icon: Building2,
            accent: T.blue,
          },
          {
            label: t.dash.myShift,
            value: myShift?.name || "—",
            sub: myShift ? `${myShift.start}–${myShift.end}` : "",
            icon: Clock,
            accent: T.gold,
          },
          {
            label: t.dash.todayStatus,
            value: myStatusLabel,
            sub: "",
            icon: CheckCircle2,
            accent: T.forest,
          },
          {
            label: t.dash.payrollStatus,
            value: myPayrollPaid
              ? STATUS_MAP.paid.label
              : STATUS_MAP.pending.label,
            sub: t.dash.thisMonth,
            icon: Wallet,
            accent: T.rose,
          },
        ];
  const recent = [...attendance]
    .filter((a) => a.date === today)
    .filter((a) => role === "admin" || a.employeeId === currentEmp?.id)
    .slice(-5)
    .reverse();

  return (
    <div>
      <Card
        style={{
          padding: 20,
          marginBottom: 22,
          background: BRAND.ink,
          color: "#fff",
          position: "relative",
          overflow: "hidden",
        }}
      >
        <p style={{ color: "#A9B4C7", fontSize: 13 }}>{t.dash.welcome}</p>
        <h2
          style={{
            fontFamily: "'Space Grotesk',sans-serif",
            fontSize: 24,
            fontWeight: 600,
            marginTop: 2,
          }}
        >
          {role === "admin" ? t.nav.admins : currentEmp?.name}
        </h2>
        <p
          style={{
            color: "#A9B4C7",
            fontSize: 12,
            marginTop: 6,
            fontFamily: "'JetBrains Mono',monospace",
          }}
        >
          {new Date().toLocaleDateString(lang === "en" ? "en-US" : "km-KH", {
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
          {t.dash.recentAttend}
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
            {t.dash.noAttend}
          </p>
        ) : (
          <div>
            {recent.map((a) => {
              const emp = employees.find((e) => e.id === a.employeeId);
              const shift = shifts?.find((s) => s.id === emp?.shiftId);
              const lateMins =
                a.status === "late" ? lateMinutesForShift(a.checkIn, shift) : 0;
              return (
                <div
                  key={a.id}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 12,
                    padding: "10px 0",
                    borderTop: `1px solid ${T.divider}`,
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
                    {lateMins > 0 && (
                      <div
                        style={{
                          fontSize: 11,
                          color: T.goldText,
                          marginTop: 2,
                        }}
                      >
                        {formatLateDuration(lateMins, lang)}
                      </div>
                    )}
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
  const { t, lang } = useLang();
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
      <Field label={t.emps.name}>
        <Input
          value={f.name}
          onChange={set("name")}
          placeholder="ឧ. លោក សុវណ្ណ ដារា"
        />
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label={t.emps.code}>
          <Input value={f.code} onChange={set("code")} placeholder="EMP-004" />
        </Field>
        <Field label={t.emps.role}>
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
      <Field label={t.emps.joined}>
        <Input
          type="date"
          value={f.joined || todayStr()}
          onChange={set("joined")}
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
          {t.cancel}
        </Button>
        <Button
          variant="accent"
          onClick={() => onSave(f)}
          disabled={!f.name || !f.code || !f.pin}
        >
          {t.save}
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
  const { t, lang } = useLang();
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
          <Plus size={15} /> {t.emps.addBtn}
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
                borderTop: `1px solid ${T.divider}`,
              }}
            >
              <Button
                size="sm"
                variant="ghost"
                style={{ flex: 1, justifyContent: "center" }}
                onClick={() => setModal({ mode: "edit", data: e })}
              >
                <Pencil size={13} /> {t.edit}
              </Button>
              {isSuperAdmin && (
                <Button
                  size="sm"
                  variant="danger"
                  style={{ flex: 1, justifyContent: "center" }}
                  onClick={() => setConfirmDel(e)}
                >
                  <Trash2 size={13} /> {t.delete}
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
          title={modal.mode === "add" ? t.emps.addTitle : t.emps.editTitle}
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
  const { t, lang } = useLang();
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
          {t.cancel}
        </Button>
        <Button variant="accent" onClick={() => onSave(f)} disabled={!f.name}>
          {t.save}
        </Button>
      </div>
    </div>
  );
}

function Departments({ departments, setDepartments, employees, isSuperAdmin }) {
  const { t, lang } = useLang();
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
          <Plus size={15} /> {t.depts.addBtn}
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
                {countIn(d.id)} {t.depts.staffCount}
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
                borderTop: `1px solid ${T.divider}`,
              }}
            >
              <Button
                size="sm"
                variant="ghost"
                style={{ flex: 1, justifyContent: "center" }}
                onClick={() => setModal({ mode: "edit", data: d })}
              >
                <Pencil size={13} /> {t.edit}
              </Button>
              {isSuperAdmin && (
                <Button
                  size="sm"
                  variant="danger"
                  style={{ flex: 1, justifyContent: "center" }}
                  onClick={() => setConfirmDel(d)}
                >
                  <Trash2 size={13} /> {t.delete}
                </Button>
              )}
            </div>
          </Card>
        ))}
      </div>
      {modal && (
        <Modal
          title={modal.mode === "add" ? t.depts.addTitle : t.depts.editTitle}
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
  const { t, lang } = useLang();
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
          {t.cancel}
        </Button>
        <Button
          variant="accent"
          onClick={() => onSave(f)}
          disabled={!f.name || !f.start || !f.end}
        >
          {t.save}
        </Button>
      </div>
    </div>
  );
}

function Shifts({ shifts, setShifts, employees, isSuperAdmin }) {
  const { t, lang } = useLang();
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
          <Plus size={15} /> {t.sh.addBtn}
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
                    {hhmm(s.start)} – {hhmm(s.end)}
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
                {countIn(s.id)} {t.depts.staffCount}
              </span>
            </div>
            <div
              style={{
                display: "flex",
                gap: 8,
                paddingTop: 12,
                marginTop: 12,
                borderTop: `1px solid ${T.divider}`,
              }}
            >
              <Button
                size="sm"
                variant="ghost"
                style={{ flex: 1, justifyContent: "center" }}
                onClick={() => setModal({ mode: "edit", data: s })}
              >
                <Pencil size={13} /> {t.edit}
              </Button>
              {isSuperAdmin && (
                <Button
                  size="sm"
                  variant="danger"
                  style={{ flex: 1, justifyContent: "center" }}
                  onClick={() => setConfirmDel(s)}
                >
                  <Trash2 size={13} /> {t.delete}
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
function SelfPunch({ emp, shift, attendance, setAttendance, offices }) {
  const { t, lang } = useLang();
  const today = todayStr();
  const rec = attendance.find(
    (a) => a.employeeId === emp.id && a.date === today,
  );
  const [locBusy, setLocBusy] = useState(false);
  const [locError, setLocError] = useState("");
  const hasOffices = offices && offices.length > 0;

  // If one or more office branches (each with lat/lng + radius) are
  // configured, require the employee's current GPS position to fall
  // within at least one of them before allowing a punch. Returns
  // { lat, lng, distance, officeId, officeName } for the closest matching
  // branch to attach to the attendance record, or null if the punch
  // should be blocked (locError is set in that case).
  const verifyLocation = async () => {
    if (!hasOffices) return null; // no geofence configured — skip check
    setLocError("");
    setLocBusy(true);
    try {
      const coords = await getCurrentPosition();
      const match = findMatchingOffice(
        offices,
        coords.latitude,
        coords.longitude,
      );
      if (!match) {
        const nearest = nearestOffice(
          offices,
          coords.latitude,
          coords.longitude,
        );
        setLocError(
          nearest
            ? `អ្នកនៅឆ្ងាយពីការិយាល័យ "${nearest.office.name}" ${nearest.distance}m (កំណត់អនុញ្ញាត ${nearest.office.radius}m) — មិនអាចចុះឈ្មោះបានទេ`
            : "មិនអាចផ្ទៀងផ្ទាត់ទីតាំងបានទេ",
        );
        return null;
      }
      return {
        lat: coords.latitude,
        lng: coords.longitude,
        distance: match.distance,
        officeId: match.office.id,
        officeName: match.office.name,
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
    if (hasOffices) {
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
    if (hasOffices) {
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
      {hasOffices && (
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
          <MapPin size={12} /> ត្រូវការទីតាំង GPS នៅជិតសាខាមួយក្នុងចំណោម{" "}
          {offices.length} សាខា
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
          {t.att.checkIn}
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
            {t.att.checkIn} {rec.checkIn} · <StatusPill status={rec.status} />
            {rec.checkInLoc?.officeName && (
              <>
                {" "}
                ·{" "}
                <span
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 3,
                  }}
                >
                  <MapPin size={11} /> {rec.checkInLoc.officeName}
                </span>
              </>
            )}
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
          {t.att.checkOut} · {t.att.checkIn} {rec.checkIn} · {t.att.checkOut}{" "}
          {rec.checkOut}
          {rec.checkOutLoc?.officeName && (
            <>
              {" "}
              ·{" "}
              <span
                style={{ display: "inline-flex", alignItems: "center", gap: 3 }}
              >
                <MapPin size={11} /> {rec.checkOutLoc.officeName}
              </span>
            </>
          )}
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
// Add/edit form for a single office branch. Used inside a Modal by
// OfficeLocationSettings below.
function OfficeForm({ initial, onSave, onCancel }) {
  const { t } = useLang();
  const [f, setF] = useState(
    initial
      ? {
          name: initial.name,
          lat: initial.lat,
          lng: initial.lng,
          radius: initial.radius,
        }
      : { name: "", lat: "", lng: "", radius: 150 },
  );
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });

  const useCurrentLocation = async () => {
    setError("");
    setBusy(true);
    try {
      const coords = await getCurrentPosition();
      setF((prev) => ({
        ...prev,
        lat: coords.latitude.toFixed(6),
        lng: coords.longitude.toFixed(6),
      }));
    } catch {
      setError("មិនអាចទាញយកទីតាំង GPS បច្ចុប្បន្នបានទេ");
    } finally {
      setBusy(false);
    }
  };

  const submit = () => {
    const name = f.name.trim();
    const lat = Number(f.lat);
    const lng = Number(f.lng);
    const radius = Number(f.radius);
    if (!name) {
      setError("សូមបញ្ចូលឈ្មោះសាខា/ការិយាល័យ");
      return;
    }
    if (!lat || !lng || !radius) {
      setError("សូមបំពេញកូអរដោនេ និងកាំឲ្យត្រឹមត្រូវ");
      return;
    }
    setError("");
    onSave({ name, lat, lng, radius });
  };

  return (
    <div>
      <Field label="ឈ្មោះសាខា / ការិយាល័យ">
        <Input
          value={f.name}
          onChange={set("name")}
          placeholder="ការិយាល័យកណ្តាល, សាខាទួលគោក..."
        />
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="Latitude">
          <Input value={f.lat} onChange={set("lat")} placeholder="11.5564" />
        </Field>
        <Field label="Longitude">
          <Input value={f.lng} onChange={set("lng")} placeholder="104.9282" />
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
          paddingTop: 14,
          borderTop: `1px solid ${T.lineSoft}`,
        }}
      >
        <Button
          variant="ghost"
          size="sm"
          onClick={useCurrentLocation}
          disabled={busy}
        >
          <MapPin size={13} /> ប្រើទីតាំងបច្ចុប្បន្ន
        </Button>
        <div style={{ display: "flex", gap: 8 }}>
          <Button variant="ghost" size="sm" onClick={onCancel}>
            {t.cancel}
          </Button>
          <Button variant="accent" size="sm" onClick={submit}>
            {t.save}
          </Button>
        </div>
      </div>
    </div>
  );
}

// Manages the list of office branches used to geofence self-service
// check-in/check-out. Each branch has its own name, lat/lng, and radius —
// an employee's punch is accepted if they're within range of ANY branch,
// and the matched branch's name is stamped onto the attendance record.
function OfficeLocationSettings({ offices, setOffices }) {
  const { t } = useLang();
  const [open, setOpen] = useState(false);
  const [formOpen, setFormOpen] = useState(null); // null | "add" | <office being edited>
  const [confirmDel, setConfirmDel] = useState(null);

  const saveOffice = (data) => {
    if (formOpen === "add") {
      setOffices([...offices, { ...data, id: uid("off") }]);
    } else {
      setOffices(
        offices.map((o) => (o.id === formOpen.id ? { ...o, ...data } : o)),
      );
    }
    setFormOpen(null);
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
            ការការពារទីតាំង GPS សម្រាប់ Check-in (ច្រើនសាខា)
          </span>
        </div>
        <span
          style={{
            fontSize: 11.5,
            color: offices.length ? T.textSoft : T.mutedLight,
          }}
        >
          {offices.length > 0 ? `${offices.length} សាខា` : "មិនទាន់កំណត់"}
        </span>
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
            កំណត់ទីតាំងសាខានីមួយៗ ដើម្បីតម្រូវឲ្យបុគ្គលិកនៅជិតសាខាមួយណាមួយ
            ពេលចុច check-in/check-out ដោយខ្លួនឯង។ ឈ្មោះសាខាដែលបុគ្គលិកចូលជិត
            នឹងត្រូវបានកត់ត្រាទុកជាមួយកំណត់ត្រាវត្តមានរបស់គេ។
            បើមិនបន្ថែមសាខាណាមួយទេ ការការពារទីតាំងនឹងមិនដំណើរការទេ។
          </p>
          {offices.length === 0 && (
            <p style={{ fontSize: 12, color: T.mutedLight, marginBottom: 12 }}>
              មិនទាន់មានសាខាទេ
            </p>
          )}
          {offices.map((o) => (
            <div
              key={o.id}
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                gap: 8,
                padding: "9px 0",
                borderBottom: `1px solid ${T.lineSoft}`,
              }}
            >
              <div>
                <div style={{ fontWeight: 600, fontSize: 13, color: T.ink }}>
                  {o.name}
                </div>
                <div
                  style={{
                    fontSize: 11,
                    color: T.muted,
                    fontFamily: "'JetBrains Mono',monospace",
                  }}
                >
                  {o.lat.toFixed(4)}, {o.lng.toFixed(4)} · {o.radius}m
                </div>
              </div>
              <div style={{ display: "flex", gap: 10 }}>
                <button
                  onClick={() => setFormOpen(o)}
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
                  onClick={() => setConfirmDel(o)}
                  style={{
                    background: "none",
                    border: "none",
                    cursor: "pointer",
                    color: T.rose,
                  }}
                >
                  <Trash2 size={14} />
                </button>
              </div>
            </div>
          ))}
          <div style={{ marginTop: 12 }}>
            <Button
              variant="ghost"
              size="sm"
              onClick={() => setFormOpen("add")}
            >
              <Plus size={13} /> បន្ថែមសាខា
            </Button>
          </div>
        </div>
      )}
      {formOpen && (
        <Modal
          title={formOpen === "add" ? "បន្ថែមសាខា" : "កែសម្រួលសាខា"}
          onClose={() => setFormOpen(null)}
        >
          <OfficeForm
            initial={formOpen === "add" ? null : formOpen}
            onSave={saveOffice}
            onCancel={() => setFormOpen(null)}
          />
        </Modal>
      )}
      {confirmDel && (
        <ConfirmDialog
          text={`តើអ្នកពិតជាចង់លុបសាខា "${confirmDel.name}" មែនទេ?`}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setOffices(offices.filter((o) => o.id !== confirmDel.id));
            setConfirmDel(null);
          }}
        />
      )}
    </Card>
  );
}

function ManualAttendanceForm({ employees, onSave, onCancel }) {
  const { t, lang } = useLang();
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
          {t.cancel}
        </Button>
        <Button
          variant="accent"
          onClick={() => onSave(f)}
          disabled={!f.employeeId}
        >
          {t.save}
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
  offices,
  setOffices,
}) {
  const { t, lang } = useLang();
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
          offices={offices}
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
            {t.nav.myAttendance}
          </h3>
          <div style={{ overflowX: "auto" }}>
            <table className="wf-table">
              <thead>
                <tr>
                  <th>កាលបរិច្ឆេទ</th>
                  <th>ចូល</th>
                  <th>ចេញ</th>
                  <th>សាខា</th>
                  <th>{t.status}</th>
                </tr>
              </thead>
              <tbody>
                {myHistory.length === 0 && (
                  <tr>
                    <td
                      colSpan={5}
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
                    <td style={{ fontSize: 11.5, color: T.muted }}>
                      {a.checkInLoc?.officeName ||
                        a.checkOutLoc?.officeName ||
                        "—"}
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
      <OfficeLocationSettings offices={offices} setOffices={setOffices} />
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
              <th>សាខា</th>
              <th>{t.status}</th>
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
                <td style={{ fontSize: 11.5, color: T.muted }}>
                  {rec?.checkInLoc?.officeName ||
                    rec?.checkOutLoc?.officeName ||
                    "—"}
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
  const { t, lang } = useLang();
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
          {t.lv.endDate + " " + t.lv.startDate}
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
          {t.cancel}
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
  const { t, lang } = useLang();
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
            <Plus size={15} /> {t.lv.addBtn}
          </Button>
        </div>
        <Card style={{ overflowX: "auto" }}>
          <table className="wf-table">
            <thead>
              <tr>
                <th>{t.lv.type}</th>
                <th>ចាប់ពី</th>
                <th>ដល់</th>
                <th>{t.lv.reason}</th>
                <th>{t.status}</th>
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
                  <td>{getLeaveTypeLabel(lang)[r.type] || r.type}</td>
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
              <th>{t.lv.type}</th>
              <th>ចាប់ពី</th>
              <th>ដល់</th>
              <th>{t.lv.reason}</th>
              <th>{t.status}</th>
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
                  <td>{getLeaveTypeLabel(lang)[r.type] || r.type}</td>
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
                          {t.lv.approve}
                        </Button>
                        <Button
                          size="sm"
                          variant="danger"
                          onClick={() => decide(r, "rejected")}
                        >
                          {t.lv.reject}
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
   Overtime (OT) requests
   Employees request OT ahead of time (date, hours, day type, reason).
   Admins approve or reject — both decisions record who decided (name +
   role) and stay visible to the employee. Rejections require a reason.
   Approved requests feed OT pay into Payroll via computePayroll.
----------------------------------------------------------------*/
function OvertimeRequestForm({ onSave, onCancel }) {
  const { t } = useLang();
  const [f, setF] = useState({
    date: todayStr(),
    hours: "",
    dayType: suggestDayType(todayStr()),
    reason: "",
  });
  const set = (k) => (e) => {
    const val = e.target.value;
    if (k === "date") {
      setF({ ...f, date: val, dayType: suggestDayType(val) });
    } else {
      setF({ ...f, [k]: val });
    }
  };
  const hoursNum = Number(f.hours);
  const invalid = !f.date || !hoursNum || hoursNum <= 0 || hoursNum > 16;
  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label={t.ot.date}>
          <Input type="date" value={f.date} onChange={set("date")} />
        </Field>
        <Field label={t.ot.hours}>
          <Input
            type="number"
            min="0.5"
            max="16"
            step="0.5"
            value={f.hours}
            onChange={set("hours")}
            placeholder="2"
          />
        </Field>
      </div>
      <Field label={t.ot.dayType}>
        <Select value={f.dayType} onChange={set("dayType")}>
          <option value="normal">{t.ot.dtNormal}</option>
          <option value="weekend">{t.ot.dtWeekend}</option>
          <option value="holiday">{t.ot.dtHoliday}</option>
        </Select>
      </Field>
      <Field label={t.ot.reason}>
        <textarea
          className="wf-input"
          rows={3}
          style={{ resize: "vertical", fontFamily: "inherit" }}
          value={f.reason}
          onChange={set("reason")}
          placeholder={t.ot.reasonPlaceholder}
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
          {t.cancel}
        </Button>
        <Button
          variant="accent"
          onClick={() => onSave({ ...f, hours: hoursNum })}
          disabled={invalid}
        >
          {t.ot.submit}
        </Button>
      </div>
    </div>
  );
}

function OvertimeRejectModal({ onCancel, onConfirm }) {
  const { t } = useLang();
  const [reason, setReason] = useState("");
  const trimmed = reason.trim();
  return (
    <Modal title={t.ot.rejectTitle} onClose={onCancel} width={420}>
      <Field label={t.ot.rejectReason}>
        <textarea
          className="wf-input"
          rows={3}
          style={{ resize: "vertical", fontFamily: "inherit" }}
          value={reason}
          onChange={(e) => setReason(e.target.value)}
          placeholder={t.ot.rejectReasonPlaceholder}
        />
      </Field>
      {reason !== "" && !trimmed && (
        <p
          style={{
            fontSize: 12.5,
            color: T.rose,
            marginTop: -8,
            marginBottom: 12,
          }}
        >
          {t.ot.rejectReasonRequired}
        </p>
      )}
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Button variant="ghost" onClick={onCancel}>
          {t.cancel}
        </Button>
        <Button
          variant="danger-solid"
          disabled={!trimmed}
          onClick={() => onConfirm(trimmed)}
        >
          {t.ot.reject}
        </Button>
      </div>
    </Modal>
  );
}

function OvertimePolicySettings({ otPolicy, setOtPolicy }) {
  const { t } = useLang();
  const [open, setOpen] = useState(false);
  const [f, setF] = useState(otPolicy);
  useEffect(() => setF(otPolicy), [otPolicy]);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const save = () => {
    setOtPolicy({
      rateNormal: Number(f.rateNormal) || DEFAULT_OT_POLICY.rateNormal,
      rateWeekend: Number(f.rateWeekend) || DEFAULT_OT_POLICY.rateWeekend,
      rateHoliday: Number(f.rateHoliday) || DEFAULT_OT_POLICY.rateHoliday,
      hoursPerDay: Number(f.hoursPerDay) || DEFAULT_OT_POLICY.hoursPerDay,
    });
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
          <Timer size={16} color={T.forest} />
          <span style={{ fontWeight: 600, fontSize: 13.5, color: T.ink }}>
            {t.ot.policyTitle}
          </span>
        </div>
        <span
          style={{
            fontSize: 11.5,
            color: T.textSoft,
            fontFamily: "'JetBrains Mono',monospace",
          }}
        >
          {otPolicy.rateNormal}x / {otPolicy.rateWeekend}x /{" "}
          {otPolicy.rateHoliday}x
        </span>
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
            {t.ot.policyDesc}
          </p>
          <div
            style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}
          >
            <Field label={t.ot.rateNormal}>
              <Input
                type="number"
                step="0.1"
                min="1"
                value={f.rateNormal}
                onChange={set("rateNormal")}
              />
            </Field>
            <Field label={t.ot.rateWeekend}>
              <Input
                type="number"
                step="0.1"
                min="1"
                value={f.rateWeekend}
                onChange={set("rateWeekend")}
              />
            </Field>
            <Field label={t.ot.rateHoliday}>
              <Input
                type="number"
                step="0.1"
                min="1"
                value={f.rateHoliday}
                onChange={set("rateHoliday")}
              />
            </Field>
            <Field label={t.ot.hoursPerDay}>
              <Input
                type="number"
                step="0.5"
                min="1"
                value={f.hoursPerDay}
                onChange={set("hoursPerDay")}
              />
            </Field>
          </div>
          <div
            style={{
              display: "flex",
              justifyContent: "flex-end",
              marginTop: 6,
            }}
          >
            <Button variant="accent" size="sm" onClick={save}>
              {t.save}
            </Button>
          </div>
        </div>
      )}
    </Card>
  );
}

// Lets an admin configure the tax/insurance deduction percentages applied
// to every payslip, instead of the rates being hardcoded in computePayroll.
function PayrollPolicySettings({ payrollPolicy, setPayrollPolicy }) {
  const { t } = useLang();
  const [open, setOpen] = useState(false);
  const [f, setF] = useState(payrollPolicy);
  useEffect(() => setF(payrollPolicy), [payrollPolicy]);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const save = () => {
    setPayrollPolicy({
      taxRate: Number(f.taxRate) || 0,
      insuranceRate: Number(f.insuranceRate) || 0,
    });
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
          <Receipt size={16} color={T.forest} />
          <span style={{ fontWeight: 600, fontSize: 13.5, color: T.ink }}>
            {t.pay.policyTitle}
          </span>
        </div>
        <span
          style={{
            fontSize: 11.5,
            color: T.textSoft,
            fontFamily: "'JetBrains Mono',monospace",
          }}
        >
          {payrollPolicy.taxRate}% / {payrollPolicy.insuranceRate}%
        </span>
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
            {t.pay.policyDesc}
          </p>
          <div
            style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}
          >
            <Field label={t.pay.taxRateLabel}>
              <Input
                type="number"
                step="0.1"
                min="0"
                value={f.taxRate}
                onChange={set("taxRate")}
              />
            </Field>
            <Field label={t.pay.insuranceRateLabel}>
              <Input
                type="number"
                step="0.1"
                min="0"
                value={f.insuranceRate}
                onChange={set("insuranceRate")}
              />
            </Field>
          </div>
          <div
            style={{
              display: "flex",
              justifyContent: "flex-end",
              marginTop: 6,
            }}
          >
            <Button variant="accent" size="sm" onClick={save}>
              {t.save}
            </Button>
          </div>
        </div>
      )}
    </Card>
  );
}

// Small inline "decided by" line shown on both the admin and employee
// views, so everyone sees the same approve/reject attribution.
function OtDecisionNote({ r, admins }) {
  const { t } = useLang();
  if (r.status !== "approved" && r.status !== "rejected") return null;
  const decider = admins.find((a) => a.id === r.decidedById);
  const name = r.decidedByName || decider?.name || "—";
  const roleLabel = ADMIN_ROLE_LABEL[r.decidedByRole] || r.decidedByRole || "";
  return (
    <div style={{ fontSize: 11.5, color: T.textSoft, marginTop: 3 }}>
      {r.status === "approved" ? (
        <span>
          {t.ot.approvedBy} <strong>{name}</strong>
          {roleLabel ? ` · ${roleLabel}` : ""}
        </span>
      ) : (
        <span style={{ color: T.rose }}>
          {t.ot.rejectedBy} <strong>{name}</strong>
          {roleLabel ? ` · ${roleLabel}` : ""}
          {r.decisionReason ? ` — ${r.decisionReason}` : ""}
        </span>
      )}
    </div>
  );
}

function OvertimeRequests({
  role,
  currentAdmin,
  currentEmp,
  employees,
  admins,
  overtimeRequests,
  setOvertimeRequests,
  otPolicy,
  setOtPolicy,
  isSuperAdmin,
}) {
  const { t, lang } = useLang();
  const [modal, setModal] = useState(false);
  const [rejectFor, setRejectFor] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const empOf = (id) => employees.find((e) => e.id === id);
  const DAY_TYPE_LABEL = getDayTypeLabel(lang, t);

  const approve = (req) => {
    setOvertimeRequests(
      overtimeRequests.map((r) =>
        r.id === req.id
          ? {
              ...r,
              status: "approved",
              decidedById: currentAdmin?.id || null,
              decidedByName: currentAdmin?.name || "",
              decidedByRole: currentAdmin?.role || "",
              decisionReason: "",
              reviewedAt: new Date().toISOString(),
            }
          : r,
      ),
    );
  };
  const reject = (req, reason) => {
    setOvertimeRequests(
      overtimeRequests.map((r) =>
        r.id === req.id
          ? {
              ...r,
              status: "rejected",
              decidedById: currentAdmin?.id || null,
              decidedByName: currentAdmin?.name || "",
              decidedByRole: currentAdmin?.role || "",
              decisionReason: reason,
              reviewedAt: new Date().toISOString(),
            }
          : r,
      ),
    );
    setRejectFor(null);
  };

  const submit = (f) => {
    if (!currentEmp) return;
    setOvertimeRequests([
      ...overtimeRequests,
      {
        id: uid("ot"),
        employeeId: currentEmp.id,
        date: f.date,
        hours: f.hours,
        dayType: f.dayType,
        reason: f.reason.trim(),
        status: "pending",
        decidedById: null,
        decidedByName: "",
        decidedByRole: "",
        decisionReason: "",
        createdAt: new Date().toISOString(),
        reviewedAt: null,
      },
    ]);
    setModal(false);
  };

  if (role !== "admin" && currentEmp) {
    const mine = overtimeRequests
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
            <Plus size={15} /> {t.ot.addBtn}
          </Button>
        </div>
        <Card style={{ overflowX: "auto" }}>
          <table className="wf-table">
            <thead>
              <tr>
                <th>{t.ot.date}</th>
                <th>{t.ot.hours}</th>
                <th>{t.ot.dayType}</th>
                <th>{t.ot.reason}</th>
                <th>{t.status}</th>
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
                    {t.ot.noRequest}
                  </td>
                </tr>
              )}
              {mine.map((r) => (
                <tr key={r.id}>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {r.date}
                  </td>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {r.hours} {t.ot.hoursShort}
                  </td>
                  <td>{DAY_TYPE_LABEL[r.dayType] || r.dayType}</td>
                  <td
                    style={{ fontSize: 12.5, color: T.textSoft, maxWidth: 200 }}
                  >
                    {r.reason || "—"}
                  </td>
                  <td>
                    <StatusPill status={r.status} />
                    <OtDecisionNote r={r} admins={admins} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>
        {modal && (
          <Modal title={t.ot.addBtn} onClose={() => setModal(false)}>
            <OvertimeRequestForm
              onSave={submit}
              onCancel={() => setModal(false)}
            />
          </Modal>
        )}
      </div>
    );
  }

  // Admin view — pending requests surfaced on top, newest first.
  const sorted = [...overtimeRequests].sort((a, b) => {
    if (a.status === "pending" && b.status !== "pending") return -1;
    if (a.status !== "pending" && b.status === "pending") return 1;
    return b.createdAt.localeCompare(a.createdAt);
  });
  return (
    <div>
      <OvertimePolicySettings otPolicy={otPolicy} setOtPolicy={setOtPolicy} />
      <Card style={{ overflowX: "auto" }}>
        <table className="wf-table">
          <thead>
            <tr>
              <th>បុគ្គលិក</th>
              <th>{t.ot.date}</th>
              <th>{t.ot.hours}</th>
              <th>{t.ot.dayType}</th>
              <th>{t.ot.reason}</th>
              <th>{t.status}</th>
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
                  {t.ot.noRequest}
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
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {r.date}
                  </td>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {r.hours} {t.ot.hoursShort}
                  </td>
                  <td>{DAY_TYPE_LABEL[r.dayType] || r.dayType}</td>
                  <td
                    style={{ fontSize: 12.5, color: T.textSoft, maxWidth: 200 }}
                  >
                    {r.reason || "—"}
                  </td>
                  <td>
                    <StatusPill status={r.status} />
                    <OtDecisionNote r={r} admins={admins} />
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
                          onClick={() => approve(r)}
                        >
                          <ThumbsUp size={13} /> {t.ot.approve}
                        </Button>
                        <Button
                          size="sm"
                          variant="danger"
                          onClick={() => setRejectFor(r)}
                        >
                          <ThumbsDown size={13} /> {t.ot.reject}
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
      {rejectFor && (
        <OvertimeRejectModal
          onCancel={() => setRejectFor(null)}
          onConfirm={(reason) => reject(rejectFor, reason)}
        />
      )}
      {confirmDel && (
        <ConfirmDialog
          text={t.ot.confirmDel}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setOvertimeRequests(
              overtimeRequests.filter((r) => r.id !== confirmDel.id),
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
  const { t, lang } = useLang();
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
          {t.cancel}
        </Button>
        <Button
          variant="accent"
          onClick={() => onSave(f)}
          disabled={!f.name || !f.username || !f.password}
        >
          {t.save}
        </Button>
      </div>
    </div>
  );
}

function AdminAccounts({ admins, setAdmins, currentAdminId }) {
  const { t, lang } = useLang();
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
          <Plus size={15} /> {t.admAcc.addBtn}
        </Button>
      </div>
      <Card style={{ overflowX: "auto" }}>
        <table className="wf-table">
          <thead>
            <tr>
              <th>{t.emps.name}</th>
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
// Raw uploads can be several MB; we still cap the original file size to
// avoid hanging the browser while decoding huge images, but the value
// actually persisted to the database is always the small compressed
// square avatar produced below (a few KB), not the original file.
const MAX_PHOTO_BYTES = 5 * 1024 * 1024;
// Target size for the compressed square avatar we actually store.
const AVATAR_MAX_DIM = 160;
const AVATAR_JPEG_QUALITY = 0.72;

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

// Reads an image file, downsizes+crops it to a small square, and returns a
// compressed JPEG data URL. This keeps every stored employee photo to
// roughly 15-40KB regardless of how large the original upload was, which
// keeps `employees?select=*` (fetched on every page load) fast even as
// more employees add profile photos.
function fileToCompressedAvatarDataUrl(
  file,
  maxDim = AVATAR_MAX_DIM,
  quality = AVATAR_JPEG_QUALITY,
) {
  return new Promise((resolve, reject) => {
    const rawUrl = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      const side = Math.min(img.width, img.height);
      const sx = (img.width - side) / 2;
      const sy = (img.height - side) / 2;
      const canvas = document.createElement("canvas");
      canvas.width = maxDim;
      canvas.height = maxDim;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, sx, sy, side, side, 0, 0, maxDim, maxDim);
      URL.revokeObjectURL(rawUrl);
      resolve(canvas.toDataURL("image/jpeg", quality));
    };
    img.onerror = () => {
      URL.revokeObjectURL(rawUrl);
      reject(new Error("image-decode-failed"));
    };
    img.src = rawUrl;
  });
}

function MyProfile({
  currentEmp,
  employees,
  setEmployees,
  departments,
  shifts,
}) {
  const { t, lang } = useLang();
  const [f, setF] = useState({
    name: currentEmp.name || "",
    phone: currentEmp.phone || "",
    email: currentEmp.email || "",
  });
  const [nameError, setNameError] = useState("");
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
      setPhotoError("រូបភាពធំពេក សូមជ្រើសរើសរូបតូចជាងនេះ (តិចជាង 5MB)");
      return;
    }
    setPhotoError("");
    try {
      const dataUrl = await fileToCompressedAvatarDataUrl(file);
      setPhotoPreview(dataUrl);
      setSaved(false);
    } catch {
      setPhotoError("មិនអាចអានរូបភាពនេះបានទេ សូមសាកល្បងរូបភាពផ្សេង");
    }
  };

  const saveProfile = () => {
    if (!f.name.trim()) {
      setNameError(t.settings.nameRequired);
      return;
    }
    setNameError("");
    setEmployees(
      employees.map((e) =>
        e.id === currentEmp.id
          ? {
              ...e,
              name: f.name.trim(),
              phone: f.phone,
              email: f.email,
              photo: photoPreview,
            }
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
          <Avatar
            name={f.name || currentEmp.name}
            photo={photoPreview}
            size={56}
          />
          <div style={{ minWidth: 0 }}>
            <div style={{ fontWeight: 700, color: T.ink, fontSize: 15 }}>
              {f.name || currentEmp.name}
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
        <Field label={t.settings.nameLabel}>
          <Input
            value={f.name}
            onChange={(e) => {
              setF({ ...f, name: e.target.value });
              setSaved(false);
              setNameError("");
            }}
            placeholder={t.settings.namePlaceholder}
          />
        </Field>
        {nameError && (
          <p style={{ fontSize: 12.5, color: T.rose, marginBottom: 10 }}>
            {nameError}
          </p>
        )}
        <Field label="រូបភាពប្រវត្តិរូប">
          <label
            className="wf-btn wf-btn-ghost"
            style={{ cursor: "pointer", display: "inline-flex" }}
          >
            <input
              type="file"
              accept="image/*"
              onChange={onPhotoChange}
              style={{
                position: "absolute",
                width: 1,
                height: 1,
                opacity: 0,
                overflow: "hidden",
              }}
            />
            <UserCircle2 size={14} /> ជ្រើសរើសរូបភាព
          </label>
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
            <CheckCircle2 size={14} /> {t.save}
          </p>
        )}
        <Button variant="accent" onClick={saveProfile}>
          {t.save}
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

      <AppearanceCard />
    </div>
  );
}

function AppearanceCard() {
  const { t } = useLang();
  const { theme, setTheme } = useTheme();
  return (
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
        {theme === "dark" ? <Moon size={16} /> : <Sun size={16} />}{" "}
        {t.settings.appearance}
      </h3>
      <p style={{ fontSize: 12, color: T.muted, marginBottom: 16 }}>
        {t.settings.appearanceDesc}
      </p>
      <div style={{ display: "flex", gap: 10 }}>
        <button
          className="wf-btn"
          onClick={() => setTheme("light")}
          style={{
            flex: 1,
            justifyContent: "center",
            background: theme === "light" ? T.ink : "transparent",
            color: theme === "light" ? "#fff" : T.ink,
            border: `1px solid ${T.line}`,
          }}
        >
          <Sun size={14} /> {t.settings.lightMode}
        </button>
        <button
          className="wf-btn"
          onClick={() => setTheme("dark")}
          style={{
            flex: 1,
            justifyContent: "center",
            background: theme === "dark" ? T.ink : "transparent",
            color: theme === "dark" ? "#fff" : T.ink,
            border: `1px solid ${T.line}`,
          }}
        >
          <Moon size={14} /> {t.settings.darkMode}
        </button>
      </div>
    </Card>
  );
}

/* ---------------------------------------------------------------
   Admin — My Settings (name, photo, appearance)
----------------------------------------------------------------*/
function AdminSettings({ currentAdmin, admins, setAdmins, isSuperAdmin }) {
  const { t } = useLang();
  const { theme, setTheme } = useTheme();
  const { branding, setBranding } = useBranding();
  const [f, setF] = useState({ name: currentAdmin.name || "" });
  const [nameError, setNameError] = useState("");
  const [photoPreview, setPhotoPreview] = useState(currentAdmin.photo || null);
  const [photoError, setPhotoError] = useState("");
  const [saved, setSaved] = useState(false);

  const [brandForm, setBrandForm] = useState({
    name: branding.name || "",
    logo: branding.logo || null,
  });
  const [brandLogoError, setBrandLogoError] = useState("");
  const [brandSaved, setBrandSaved] = useState(false);

  const onBrandLogoChange = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      setBrandLogoError("សូមជ្រើសរើសឯកសាររូបភាព");
      return;
    }
    if (file.size > MAX_PHOTO_BYTES) {
      setBrandLogoError("រូបភាពធំពេក សូមជ្រើសរើសរូបតូចជាងនេះ (តិចជាង 5MB)");
      return;
    }
    setBrandLogoError("");
    try {
      const dataUrl = await fileToCompressedAvatarDataUrl(file, 200, 0.85);
      setBrandForm((prev) => ({ ...prev, logo: dataUrl }));
      setBrandSaved(false);
    } catch {
      setBrandLogoError("មិនអាចអានរូបភាពនេះបានទេ សូមសាកល្បងរូបភាពផ្សេង");
    }
  };

  const saveBranding = () => {
    setBranding({ name: brandForm.name.trim(), logo: brandForm.logo || null });
    setBrandSaved(true);
  };

  const onPhotoChange = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      setPhotoError("សូមជ្រើសរើសឯកសាររូបភាព");
      return;
    }
    if (file.size > MAX_PHOTO_BYTES) {
      setPhotoError("រូបភាពធំពេក សូមជ្រើសរើសរូបតូចជាងនេះ (តិចជាង 5MB)");
      return;
    }
    setPhotoError("");
    try {
      const dataUrl = await fileToCompressedAvatarDataUrl(file);
      setPhotoPreview(dataUrl);
      setSaved(false);
    } catch {
      setPhotoError("មិនអាចអានរូបភាពនេះបានទេ សូមសាកល្បងរូបភាពផ្សេង");
    }
  };

  const save = () => {
    if (!f.name.trim()) {
      setNameError(t.settings.nameRequired);
      return;
    }
    setNameError("");
    setAdmins(
      admins.map((a) =>
        a.id === currentAdmin.id
          ? { ...a, name: f.name.trim(), photo: photoPreview }
          : a,
      ),
    );
    setSaved(true);
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
          <Avatar
            name={f.name || currentAdmin.name}
            photo={photoPreview}
            size={56}
          />
          <div style={{ minWidth: 0 }}>
            <div style={{ fontWeight: 700, color: T.ink, fontSize: 15 }}>
              {f.name || currentAdmin.name}
            </div>
            <div
              style={{
                fontSize: 11.5,
                color: T.muted,
                fontFamily: "'JetBrains Mono',monospace",
              }}
            >
              {ADMIN_ROLE_LABEL[currentAdmin.role]}
            </div>
          </div>
        </div>
        <Field label={t.settings.nameLabel}>
          <Input
            value={f.name}
            onChange={(e) => {
              setF({ ...f, name: e.target.value });
              setSaved(false);
              setNameError("");
            }}
            placeholder={t.settings.namePlaceholder}
          />
        </Field>
        {nameError && (
          <p style={{ fontSize: 12.5, color: T.rose, marginBottom: 10 }}>
            {nameError}
          </p>
        )}
        <Field label={t.settings.photoLabel}>
          <label
            className="wf-btn wf-btn-ghost"
            style={{ cursor: "pointer", display: "inline-flex" }}
          >
            <input
              type="file"
              accept="image/*"
              onChange={onPhotoChange}
              style={{
                position: "absolute",
                width: 1,
                height: 1,
                opacity: 0,
                overflow: "hidden",
              }}
            />
            <Camera size={14} /> {t.settings.choosePhoto}
          </label>
        </Field>
        {photoError && (
          <p style={{ fontSize: 12.5, color: T.rose, marginBottom: 10 }}>
            {photoError}
          </p>
        )}
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
            <CheckCircle2 size={14} /> {t.settings.saved}
          </p>
        )}
        <Button variant="accent" onClick={save}>
          {t.save}
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
          {theme === "dark" ? <Moon size={16} /> : <Sun size={16} />}{" "}
          {t.settings.appearance}
        </h3>
        <p style={{ fontSize: 12, color: T.muted, marginBottom: 16 }}>
          {t.settings.appearanceDesc}
        </p>
        <div style={{ display: "flex", gap: 10 }}>
          <button
            className="wf-btn"
            onClick={() => setTheme("light")}
            style={{
              flex: 1,
              justifyContent: "center",
              background: theme === "light" ? T.ink : "transparent",
              color: theme === "light" ? "#fff" : T.ink,
              border: `1px solid ${T.line}`,
            }}
          >
            <Sun size={14} /> {t.settings.lightMode}
          </button>
          <button
            className="wf-btn"
            onClick={() => setTheme("dark")}
            style={{
              flex: 1,
              justifyContent: "center",
              background: theme === "dark" ? T.ink : "transparent",
              color: theme === "dark" ? "#fff" : T.ink,
              border: `1px solid ${T.line}`,
            }}
          >
            <Moon size={14} /> {t.settings.darkMode}
          </button>
        </div>
      </Card>

      {isSuperAdmin && (
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
            <Building2 size={16} /> {t.settings.brandingTitle}
          </h3>
          <p style={{ fontSize: 12, color: T.muted, marginBottom: 16 }}>
            {t.settings.brandingDesc}
          </p>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 14,
              marginBottom: 16,
            }}
          >
            <div
              style={{
                width: 56,
                height: 56,
                borderRadius: 14,
                background: brandForm.logo
                  ? "#fff"
                  : `linear-gradient(145deg,${T.forest},${T.forestDark})`,
                border: `1px solid ${T.line}`,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                overflow: "hidden",
                flexShrink: 0,
                color: "#fff",
                fontWeight: 700,
                fontFamily: "'Space Grotesk',sans-serif",
              }}
            >
              {brandForm.logo ? (
                <img
                  src={brandForm.logo}
                  alt={brandForm.name || t.appName}
                  style={{ width: "100%", height: "100%", objectFit: "cover" }}
                />
              ) : (
                getInitials(brandForm.name || t.appName)
              )}
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <label
                className="wf-btn wf-btn-ghost"
                style={{ cursor: "pointer", display: "inline-flex" }}
              >
                <input
                  type="file"
                  accept="image/*"
                  onChange={onBrandLogoChange}
                  style={{
                    position: "absolute",
                    width: 1,
                    height: 1,
                    opacity: 0,
                    overflow: "hidden",
                  }}
                />
                <Camera size={14} /> {t.settings.chooseLogo}
              </label>
              {brandForm.logo && (
                <button
                  type="button"
                  className="wf-btn wf-btn-ghost"
                  onClick={() => {
                    setBrandForm((prev) => ({ ...prev, logo: null }));
                    setBrandSaved(false);
                  }}
                >
                  <X size={14} /> {t.settings.removeLogo}
                </button>
              )}
            </div>
          </div>
          {brandLogoError && (
            <p style={{ fontSize: 12.5, color: T.rose, marginBottom: 10 }}>
              {brandLogoError}
            </p>
          )}
          <Field label={t.settings.companyNameLabel}>
            <Input
              value={brandForm.name}
              onChange={(e) => {
                setBrandForm((prev) => ({ ...prev, name: e.target.value }));
                setBrandSaved(false);
              }}
              placeholder={t.settings.companyNamePlaceholder}
            />
          </Field>
          {brandSaved && (
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
              <CheckCircle2 size={14} /> {t.settings.brandingSaved}
            </p>
          )}
          <Button variant="accent" onClick={saveBranding}>
            {t.save}
          </Button>
        </Card>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Payroll
----------------------------------------------------------------*/
function Payslip({
  emp,
  mk,
  attendance,
  overtimeRequests,
  otPolicy,
  payrollPolicy,
  onClose,
}) {
  const { t, lang } = useLang();
  const {
    absentDays,
    leaveDays,
    dailyRate,
    absenceDeduction,
    tax,
    insurance,
    taxRate,
    insuranceRate,
    otHours,
    otPay,
    net,
  } = computePayroll(
    emp,
    attendance,
    mk,
    overtimeRequests,
    otPolicy,
    payrollPolicy,
  );
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
          <Receipt size={20} /> {t.nav.myPayroll}
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
          <span style={{ color: T.textSoft }}>{t.pay.baseSalary}</span>
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
              {t.pay.absentDed} ({absentDays} × {fmtMoney(dailyRate)})
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
            <span>{t.lv.approved}</span>
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
          <span>
            {t.pay.taxLabel} ({taxRate}%)
          </span>
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
          <span>
            {t.pay.insuranceLabel} ({insuranceRate}%)
          </span>
          <span style={{ fontFamily: "'JetBrains Mono',monospace" }}>
            -{fmtMoney(insurance)}
          </span>
        </div>
        {otHours > 0 && (
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              color: T.forestText,
            }}
          >
            <span>
              {t.pay.otPay} ({otHours} {t.ot.hoursShort})
            </span>
            <span style={{ fontFamily: "'JetBrains Mono',monospace" }}>
              +{fmtMoney(otPay)}
            </span>
          </div>
        )}
      </div>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          background: BRAND.ink,
          color: "#fff",
          borderRadius: 10,
          padding: "12px 16px",
          marginTop: 16,
        }}
      >
        <span style={{ fontWeight: 600, fontSize: 13 }}>{t.pay.netSalary}</span>
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
  overtimeRequests,
  otPolicy,
  payrollPolicy,
  setPayrollPolicy,
}) {
  const { t, lang } = useLang();
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
    (sum, e) =>
      sum +
      computePayroll(
        e,
        attendance,
        mk,
        overtimeRequests,
        otPolicy,
        payrollPolicy,
      ).net,
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
          {lang === "en" ? "Month" : "ខែ"}
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
        <PayrollPolicySettings
          payrollPolicy={payrollPolicy}
          setPayrollPolicy={setPayrollPolicy}
        />
      )}
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
              {t.pay.totalPaid} · {monthLabel(mk)}
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
              <th>{t.employee}</th>
              <th>{t.pay.baseSalary}</th>
              <th>{t.pay.netSalary}</th>
              <th>{t.status}</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {list.map((e) => {
              const paid = !!payrollPaid[`${e.id}-${mk}`];
              const { net, absentDays, otHours } = computePayroll(
                e,
                attendance,
                mk,
                overtimeRequests,
                otPolicy,
                payrollPolicy,
              );
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
                            {t.att.absentDays} {absentDays}
                          </div>
                        )}
                        {otHours > 0 && (
                          <div style={{ fontSize: 10.5, color: T.forestText }}>
                            {t.ot.totalOtHours}: {otHours} {t.ot.hoursShort}
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
                      {t.pay.viewSlip}
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
                        {paid ? t.pay.unmarkPaid : t.pay.markPaid}
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
          overtimeRequests={overtimeRequests}
          otPolicy={otPolicy}
          payrollPolicy={payrollPolicy}
          onClose={() => setSlipFor(null)}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   App shell
----------------------------------------------------------------*/
function buildNavAdmin(n) {
  return [
    { id: "dashboard", label: n.dashboard, icon: LayoutDashboard },
    { id: "employees", label: n.employees, icon: Users },
    { id: "departments", label: n.departments, icon: Building2 },
    { id: "shifts", label: n.shifts, icon: Watch },
    { id: "attendance", label: n.attendance, icon: Clock },
    { id: "leave", label: n.leave, icon: CalendarDays },
    { id: "ot", label: n.overtime, icon: Timer },
    { id: "payroll", label: n.payroll, icon: Wallet },
    { id: "admins", label: n.admins, icon: ShieldCheck, superadminOnly: true },
    { id: "settings", label: n.settings, icon: Settings2 },
  ];
}
function buildNavEmployee(n) {
  return [
    { id: "dashboard", label: n.dashboard, icon: LayoutDashboard },
    { id: "attendance", label: n.myAttendance, icon: Clock },
    { id: "leave", label: n.myLeave, icon: CalendarDays },
    { id: "ot", label: n.myOvertime, icon: Timer },
    { id: "payroll", label: n.myPayroll, icon: Wallet },
    { id: "profile", label: n.myProfile, icon: UserCircle2 },
  ];
}

function AppInner() {
  useGlobalStyle();
  const { t } = useLang();
  const [branding, setBranding, brandingReady] = useBrandingSettings();
  const brandDisplayName = branding.name?.trim() || t.appName;
  const { theme } = useTheme();

  // The browser tab title/favicon live outside React (in the document
  // head), so setting `branding` state alone doesn't touch them — push
  // the custom name/logo into the tab explicitly whenever they change.
  useEffect(() => {
    if (!brandingReady) return;
    document.title = brandDisplayName;
  }, [brandingReady, brandDisplayName]);

  useEffect(() => {
    if (!brandingReady) return;
    let link = document.querySelector("link[rel~='icon']");
    if (!link) {
      link = document.createElement("link");
      link.rel = "icon";
      document.head.appendChild(link);
    }
    if (branding.logo) {
      link.href = branding.logo;
    }
    // If no custom logo is set, leave whatever favicon index.html already
    // defines (e.g. a default app icon) rather than clearing it.
  }, [brandingReady, branding.logo]);
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
  const [overtimeRequests, setOvertimeRequests, otrReady] = useSupabaseArray(
    "overtime_requests",
    {
      fromDb: (r) => ({
        id: r.id,
        employeeId: r.employee_id,
        date: r.date,
        hours: r.hours,
        dayType: r.day_type,
        reason: r.reason,
        status: r.status,
        decidedById: r.decided_by_id,
        decidedByName: r.decided_by_name,
        decidedByRole: r.decided_by_role,
        decisionReason: r.decision_reason,
        createdAt: r.created_at,
        reviewedAt: r.reviewed_at,
      }),
      toDb: (r) => ({
        id: r.id,
        employee_id: r.employeeId,
        date: r.date,
        hours: r.hours,
        day_type: r.dayType,
        reason: r.reason,
        status: r.status,
        decided_by_id: r.decidedById,
        decided_by_name: r.decidedByName,
        decided_by_role: r.decidedByRole,
        decision_reason: r.decisionReason,
        created_at: r.createdAt,
        reviewed_at: r.reviewedAt,
      }),
    },
  );
  const [admins, setAdmins, adminsReady] = useSupabaseArray("admins");
  const [offices, setOffices, officesReady] = useSupabaseArray("offices", {
    fromDb: (r) => ({
      id: r.id,
      name: r.name,
      lat: Number(r.lat),
      lng: Number(r.lng),
      radius: Number(r.radius),
    }),
    toDb: (r) => ({
      id: r.id,
      name: r.name,
      lat: r.lat,
      lng: r.lng,
      radius: r.radius,
    }),
    orderBy: "name",
  });
  const [otPolicy, setOtPolicy, otPolicyReady] = useOtPolicy();
  const [payrollPolicy, setPayrollPolicy, payrollPolicyReady] =
    usePayrollPolicy();
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
  const [navOpen, setNavOpen] = useState(false);
  const [portal, routedPage, goPortal, setPage] = usePortalRoute();
  // Falls back to "dashboard" only when the URL has no page segment yet
  // (e.g. a bare #/employee link) — otherwise refreshing always restores
  // whatever page was open, for both the admin and staff portals.
  const page = routedPage || "dashboard";

  const ready =
    dReady &&
    eReady &&
    shReady &&
    aReady &&
    pReady &&
    lrReady &&
    otrReady &&
    adminsReady &&
    officesReady &&
    otPolicyReady &&
    payrollPolicyReady &&
    brandingReady &&
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
      ? buildNavAdmin(t.nav).filter((n) => !n.superadminOnly || isSuperAdmin)
      : buildNavEmployee(t.nav);

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
        <BrandingContext.Provider value={{ branding, setBranding }}>
          <AdminLoginScreen
            admins={admins}
            onLogin={(id) => setSessionAdmin(id)}
            go={goPortal}
          />
        </BrandingContext.Provider>
      );
    return (
      <BrandingContext.Provider value={{ branding, setBranding }}>
        <EmployeeLoginScreen
          employees={employees}
          onLogin={(id) => setSessionEmployee(id)}
          go={goPortal}
        />
      </BrandingContext.Provider>
    );
  }

  return (
    <BrandingContext.Provider value={{ branding, setBranding }}>
      <div className={`wf-root ${theme === "dark" ? "wf-dark" : ""}`}>
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
              <div className="wf-logo-badge">
                {branding.logo ? (
                  <img
                    src={branding.logo}
                    alt={brandDisplayName}
                    style={{
                      width: "100%",
                      height: "100%",
                      objectFit: "cover",
                      borderRadius: "inherit",
                    }}
                  />
                ) : (
                  getInitials(brandDisplayName)
                )}
              </div>
              <span
                style={{
                  fontWeight: 600,
                  fontSize: 15,
                  fontFamily: "'Space Grotesk',sans-serif",
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                  whiteSpace: "nowrap",
                }}
              >
                {brandDisplayName}
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
                  photo={
                    role === "admin" ? currentAdmin?.photo : currentEmp?.photo
                  }
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
                <LogOut size={17} /> {t.logout}
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
                  {t.employee}
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
              <ThemeToggle variant="light" />
              <LangToggle variant="light" />
              <NotificationBell
                role={role}
                currentAdmin={currentAdmin}
                currentEmp={currentEmp}
                employees={employees}
                shifts={shifts}
                attendance={attendance}
                leaveRequests={leaveRequests}
                overtimeRequests={overtimeRequests}
                setPage={setPage}
              />
              <Avatar
                name={
                  role === "admin"
                    ? currentAdmin?.name || "?"
                    : currentEmp?.name || "?"
                }
                photo={
                  role === "admin" ? currentAdmin?.photo : currentEmp?.photo
                }
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
                  <AlertCircle size={16} /> {t.dash.noEmpWarn}
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
                shifts={shifts}
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
                offices={offices}
                setOffices={setOffices}
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
            {page === "ot" && (
              <OvertimeRequests
                role={role}
                currentAdmin={currentAdmin}
                currentEmp={currentEmp}
                employees={employees}
                admins={admins}
                overtimeRequests={overtimeRequests}
                setOvertimeRequests={setOvertimeRequests}
                otPolicy={otPolicy}
                setOtPolicy={setOtPolicy}
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
                overtimeRequests={overtimeRequests}
                otPolicy={otPolicy}
                payrollPolicy={payrollPolicy}
                setPayrollPolicy={setPayrollPolicy}
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
            {page === "settings" && role === "admin" && currentAdmin && (
              <AdminSettings
                currentAdmin={currentAdmin}
                admins={admins}
                setAdmins={setAdmins}
                isSuperAdmin={isSuperAdmin}
              />
            )}
          </main>
        </div>
      </div>
    </BrandingContext.Provider>
  );
}

export default function App() {
  const [lang, setLang] = useLocalStorage("hrsuite:lang", "km");
  const t = LANG[lang] || LANG.km;
  const [theme, setTheme] = useLocalStorage("hrsuite:theme", "light");
  const toggleTheme = () => setTheme(theme === "dark" ? "light" : "dark");
  return (
    <LangContext.Provider value={{ lang, t, setLang }}>
      <ThemeContext.Provider value={{ theme, setTheme, toggleTheme }}>
        <AppInner />
      </ThemeContext.Provider>
    </LangContext.Provider>
  );
}
