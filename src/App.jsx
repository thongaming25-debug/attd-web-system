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
import ExcelJS from "exceljs";
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
  ChevronLeft,
  ChevronRight,
  Download,
  Star,
  Megaphone,
  FileText,
  Paperclip,
  Upload,
  CalendarClock,
  History,
  RefreshCw,
  Smartphone,
  Tablet,
  Monitor,
  BarChart3,
  FileSpreadsheet,
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
      announcements: "សេចក្តីប្រកាស",
      employees: "បុគ្គលិក",
      departments: "នាយកដ្ឋាន",
      shifts: "វេនការងារ",
      attendance: "វត្តមាន",
      holidays: "ថ្ងៃឈប់សម្រាកជាតិ",
      leave: "ច្បាប់ឈប់សម្រាក",
      overtime: "ការងារបន្ថែម (OT)",
      payroll: "ប្រាក់ខែ",
      performance: "ការវាយតម្លៃការងារ",
      attCorrection: "សំណើកែតម្រូវវត្តមាន",
      admins: "គណនីអ្នកគ្រប់គ្រង",
      myAttendance: "វត្តមានរបស់ខ្ញុំ",
      myLeave: "ច្បាប់ឈប់សម្រាករបស់ខ្ញុំ",
      myOvertime: "ការងារបន្ថែម (OT) របស់ខ្ញុំ",
      myPayroll: "ប្រាក់ខែរបស់ខ្ញុំ",
      myPerformance: "ការវាយតម្លៃការងាររបស់ខ្ញុំ",
      myDocuments: "ឯកសាររបស់ខ្ញុំ",
      myAttCorrection: "ស្នើសុំកែតម្រូវវត្តមាន",
      myProfile: "ប្រវត្តិរូបរបស់ខ្ញុំ",
      settings: "ការកំណត់",
      audits: "កំណត់ត្រាសកម្មភាព",
      loginActivity: "សកម្មភាពចូលគណនី",
      analytics: "វិភាគទិន្នន័យ",
      rolePerms: "សិទ្ធិតួនាទី",
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
    exportCsv: "នាំចេញ CSV",
    exportExcel: "នាំចេញ Excel",
    pagination: { of: "នៃ" },
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
    analytics: {
      title: "វិភាគទិន្នន័យ",
      subtitle: "ទិដ្ឋភាពទូទៅនិន្នាការវត្តមាន ប្រាក់ខែ និង OT",
      attendTrend: "អត្រាមកធ្វើការ (៦ ខែចុងក្រោយ)",
      attendTrendSub: "ភាគរយវត្តមាន/មកយឺត ធៀបនឹងកំណត់ត្រាសរុបប្រចាំខែ",
      deptCost: "ចំណាយប្រាក់ខែសុទ្ធតាមនាយកដ្ឋាន",
      deptCostSub: "ប្រចាំខែបច្ចុប្បន្ន បុគ្គលិកសកម្មប៉ុណ្ណោះ",
      otTrend: "ម៉ោង OT ដែលបានអនុម័ត (៦ ខែចុងក្រោយ)",
      otTrendSub: "សរុបម៉ោង OT ដែលបានអនុម័តគ្រប់នាយកដ្ឋាន",
      noChartData: "មិនទាន់មានទិន្នន័យគ្រប់គ្រាន់ដើម្បីបង្ហាញទេ",
      totalOt: "OT សរុប",
      hours: "ម៉ោង",
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
      printBadge: "បោះពុម្ពកាតសម្គាល់",
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
      approvedBy: "អនុម័តដោយ",
      rejectedBy: "បដិសេធដោយ",
      rejectTitle: "បដិសេធសំណើច្បាប់",
      rejectReason: "មូលហេតុបដិសេធ",
      rejectReasonPlaceholder: "សូមបញ្ជាក់មូលហេតុបដិសេធ...",
      rejectReasonRequired: "សូមបញ្ចូលមូលហេតុបដិសេធ",
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
        "កំណត់អត្រាភាគរយពន្ធលើប្រាក់ខែ និងធានារ៉ាប់រង ព្រមទាំងកម្រិតប្រាក់ខែអប្បបរមា។ ការកាត់ប្រាក់នេះនឹងអនុវត្តលើបុគ្គលិកដែលមានប្រាក់ខែមូលដ្ឋានស្មើ ឬលើសពីកម្រិតកំណត់នេះប៉ុណ្ណោះ។",
      taxRateLabel: "អត្រាពន្ធលើប្រាក់ខែ (%)",
      insuranceRateLabel: "អត្រាធានារ៉ាប់រង (%)",
      minSalaryThresholdLabel: "ចាប់ផ្តើមគិតចាប់ពីប្រាក់ខែ ($)",
      minSalaryThresholdHint:
        "បុគ្គលិកដែលមានប្រាក់ខែមូលដ្ឋានតិចជាងចំនួននេះ នឹងមិនត្រូវកាត់ពន្ធ/ធានារ៉ាប់រងទេ",
      noThreshold: "អនុវត្តលើគ្រប់បុគ្គលិកទាំងអស់",
      downloadPdf: "ទាញយក PDF",
      payslipTitle: "សន្លឹកប្រាក់ខែ",
      payPeriod: "រយៈពេលបើកប្រាក់ខែ",
      generatedOn: "បង្កើតនៅថ្ងៃ",
      customRateToggle: "កំណត់អត្រាពន្ធ/ធានារ៉ាប់រងផ្ទាល់ខ្លួន",
      customRateHint:
        "នៅពេលបើក បុគ្គលិកនេះនឹងប្រើអត្រាដែលកំណត់ខាងក្រោម ជំនួសអត្រាទូទៅរបស់គោលការណ៍ក្រុមហ៊ុន (មិនគិតតាមកម្រិតប្រាក់ខែអប្បបរមាទេ)",
      customTaxRateLabel: "អត្រាពន្ធផ្ទាល់ខ្លួន (%)",
      customInsuranceRateLabel: "អត្រាធានារ៉ាប់រងផ្ទាល់ខ្លួន (%)",
      customRateBadge: "អត្រាផ្ទាល់ខ្លួន",
      historicalBtn: "របាយការណ៍ខែចាស់",
      historicalTitle: "របាយការណ៍ប្រាក់ខែខែចាស់",
      historicalDesc:
        "ជ្រើសខែ/ឆ្នាំណាមួយ ដើម្បីទាញយកទិន្នន័យវត្តមានពី database ដោយផ្ទាល់ (មិនកម្រិតត្រឹម ៦ខែថ្មីៗទេ)",
      historicalPick: "ខែ/ឆ្នាំ",
      historicalLoad: "ទាញយកទិន្នន័យ",
      historicalLoading: "កំពុងទាញយកទិន្នន័យ...",
      historicalEmpty: "គ្មានកំណត់ត្រាវត្តមានសម្រាប់ខែនេះទេ",
      historicalError: "មិនអាចទាញយកទិន្នន័យបានទេ សូមព្យាយាមម្តងទៀត",
    },
    pr: {
      addBtn: "បន្ថែមការវាយតម្លៃ",
      editTitle: "កែសម្រួលការវាយតម្លៃ",
      addTitle: "បន្ថែមការវាយតម្លៃការងារ",
      employee: "បុគ្គលិក",
      period: "រយៈពេលវាយតម្លៃ",
      periodPlaceholder: "ឧ. ត្រីមាសទី១ ឆ្នាំ២០២៦",
      rating: "ការវាយតម្លៃ (ផ្កាយ)",
      strengths: "ចំណុចខ្លាំង",
      strengthsPlaceholder: "សរសេរចំណុចខ្លាំងរបស់បុគ្គលិក...",
      improvements: "ចំណុចត្រូវកែលម្អ",
      improvementsPlaceholder: "សរសេរចំណុចដែលគួរកែលម្អ...",
      reviewedBy: "វាយតម្លៃដោយ",
      confirmDel: "តើអ្នកប្រាកដទេថាចង់លុបការវាយតម្លៃនេះ?",
      noReview: "មិនទាន់មានការវាយតម្លៃទេ",
      notifTitle: "អ្នកទទួលបានការវាយតម្លៃការងារថ្មី",
    },
    ann: {
      addBtn: "បង្កើតសេចក្តីប្រកាស",
      editTitle: "កែសម្រួលសេចក្តីប្រកាស",
      addTitle: "សេចក្តីប្រកាសថ្មី",
      titleLabel: "ចំណងជើង",
      titlePlaceholder: "ឧ. ថ្ងៃឈប់សម្រាកចុងឆ្នាំ",
      bodyLabel: "ខ្លឹមសារ",
      bodyPlaceholder: "សរសេរខ្លឹមសារសេចក្តីប្រកាស...",
      postedBy: "ប្រកាសដោយ",
      confirmDel: "តើអ្នកប្រាកដទេថាចង់លុបសេចក្តីប្រកាសនេះ?",
      noAnn: "មិនទាន់មានសេចក្តីប្រកាសទេ",
      notifTitle: "សេចក្តីប្រកាសថ្មីពីក្រុមហ៊ុន",
    },
    doc: {
      title: "ឯកសារបុគ្គលិក",
      addBtn: "ផ្ទុកឯកសារ",
      category: "ប្រភេទឯកសារ",
      catCv: "ប្រវត្តិរូបសង្ខេប (CV)",
      catContract: "កិច្ចសន្យា",
      catId: "អត្តសញ្ញាណប័ណ្ណ",
      catOther: "ផ្សេងៗ",
      chooseFile: "ជ្រើសរើសឯកសារ",
      uploadedBy: "ផ្ទុកឡើងដោយ",
      confirmDel: "តើអ្នកប្រាកដទេថាចង់លុបឯកសារនេះ?",
      noDocs: "មិនទាន់មានឯកសារទេ",
      tooLarge: "ឯកសារធំពេក សូមជ្រើសរើសឯកសារតូចជាងនេះ (តិចជាង 4MB)",
      view: "មើល/ទាញយក",
      myTitle: "ឯកសាររបស់ខ្ញុំ",
    },
    hol: {
      addBtn: "បន្ថែមថ្ងៃឈប់សម្រាក",
      editTitle: "កែសម្រួលថ្ងៃឈប់សម្រាកជាតិ",
      addTitle: "បន្ថែមថ្ងៃឈប់សម្រាកជាតិ",
      dateLabel: "កាលបរិច្ឆេទ",
      nameLabel: "ឈ្មោះថ្ងៃឈប់សម្រាក",
      namePlaceholder: "ឧ. ថ្ងៃចូលឆ្នាំថ្មីខ្មែរ",
      confirmDel: "តើអ្នកប្រាកដទេថាចង់លុបថ្ងៃឈប់សម្រាកនេះ?",
      noHolidays: "មិនទាន់មានកំណត់ថ្ងៃឈប់សម្រាកជាតិទេ",
      upcoming: "ថ្ងៃឈប់សម្រាកខាងមុខ",
      past: "ថ្ងៃឈប់សម្រាកកន្លងទៅ",
      todayBanner: "ថ្ងៃនេះជាថ្ងៃឈប់សម្រាកជាតិ",
    },
    ac: {
      addBtn: "ស្នើសុំកែតម្រូវវត្តមាន",
      date: "កាលបរិច្ឆេទ",
      requestedCheckIn: "ម៉ោងចូលដែលស្នើសុំ",
      requestedCheckOut: "ម៉ោងចេញដែលស្នើសុំ",
      reason: "មូលហេតុ",
      reasonPlaceholder: "ឧ. ភ្លេចចុចម៉ោងចូល/ចេញ...",
      approve: "អនុម័ត",
      reject: "បដិសេធ",
      approvedBy: "អនុម័តដោយ",
      rejectedBy: "បដិសេធដោយ",
      rejectTitle: "បដិសេធសំណើកែតម្រូវវត្តមាន",
      rejectReason: "មូលហេតុបដិសេធ",
      rejectReasonPlaceholder: "សូមបញ្ជាក់មូលហេតុបដិសេធ...",
      rejectReasonRequired: "សូមបញ្ចូលមូលហេតុបដិសេធ",
      confirmDel: "តើអ្នកប្រាកដទេថាចង់លុបសំណើនេះ?",
      noRequest: "មិនទាន់មានសំណើកែតម្រូវវត្តមានទេ",
      needOneField: "សូមបញ្ចូលម៉ោងចូល ឬម៉ោងចេញ យ៉ាងហោចណាស់មួយ",
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
      saveFailed: "រក្សាទុកមិនបានសម្រេច៖",
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
    audit: {
      title: "កំណត់ត្រាសកម្មភាព",
      subtitle: "តាមដានថាអ្នកណាបានកែប្រែអ្វី និងនៅពេលណា",
      refresh: "ផ្ទុកឡើងវិញ",
      time: "ពេលវេលា",
      actor: "អ្នកអនុវត្ត",
      action: "សកម្មភាព",
      entity: "ធាតុទិន្នន័យ",
      details: "ព័ត៌មានលម្អិត",
      noLogs: "មិនទាន់មានកំណត់ត្រាទេ",
      loading: "កំពុងផ្ទុក...",
      filterTable: "ត្រងតាមប្រភេទ",
      allTables: "ទាំងអស់",
      actionCreate: "បានបង្កើត",
      actionUpdate: "បានកែប្រែ",
      actionDelete: "បានលុប",
      actionLogin: "បានចូលប្រព័ន្ធ",
      actionLogout: "បានចាកចេញ",
      changedFrom: "ពី",
      changedTo: "ទៅ",
      unknownActor: "មិនស្គាល់",
      tables: {
        departments: "នាយកដ្ឋាន",
        employees: "បុគ្គលិក",
        shifts: "វេនការងារ",
        holidays: "ថ្ងៃឈប់សម្រាកជាតិ",
        leave_requests: "ច្បាប់ឈប់សម្រាក",
        overtime_requests: "ការងារបន្ថែម (OT)",
        performance_reviews: "ការវាយតម្លៃការងារ",
        announcements: "សេចក្តីប្រកាស",
        admins: "គណនីអ្នកគ្រប់គ្រង",
        offices: "ទីតាំងការិយាល័យ",
        attendance_corrections: "សំណើកែតម្រូវវត្តមាន",
        employee_documents: "ឯកសារបុគ្គលិក",
      },
    },
    loginAct: {
      title: "សកម្មភាពចូលគណនី",
      subtitle: "ប្រវត្តិចូល និងចេញពីគណនីរបស់អ្នក រួមទាំងឧបករណ៍ដែលបានប្រើ",
      subtitleAll: "ប្រវត្តិចូល និងចេញរបស់អ្នកប្រើប្រាស់ទាំងអស់",
      scopeMine: "របស់ខ្ញុំ",
      scopeAll: "អ្នកប្រើប្រាស់ទាំងអស់",
      refresh: "ផ្ទុកឡើងវិញ",
      time: "ពេលវេលា",
      actor: "អ្នកប្រើប្រាស់",
      action: "សកម្មភាព",
      device: "ឧបករណ៍",
      os: "ប្រព័ន្ធប្រតិបត្តិការ",
      model: "ម៉ូដែល (Android)",
      browser: "កម្មវិធីរុករក",
      noLogs: "មិនទាន់មានកំណត់ត្រាទេ",
      loading: "កំពុងផ្ទុក...",
      actionLogin: "បានចូលប្រព័ន្ធ",
      actionLogout: "បានចាកចេញ",
      unknownActor: "មិនស្គាល់",
      deviceMobile: "ទូរស័ព្ទ",
      deviceTablet: "Tablet",
      deviceDesktop: "កុំព្យូទ័រ",
      unknown: "មិនស្គាល់",
      active: "កំពុងចូល",
      currentDevice: "ឧបករណ៍នេះ",
      revoke: "មិនមែនខ្ញុំ — Sign out ភ្លាម",
      revokeConfirm:
        "តើអ្នកប្រាកដទេថាចង់បង្ខំ Sign out ឧបករណ៍នេះចេញពីគណនីភ្លាមៗ?",
      revoked: "បាន Sign out",
      revokedAlert:
        "គណនីរបស់អ្នកត្រូវបានបង្ខំឲ្យ Sign out ពីឧបករណ៍មួយផ្សេងទៀត។",
      clearOld: "លុបចាស់ (30+ ថ្ងៃ)",
      clearAll: "លុបទាំងអស់",
      clearOldConfirm:
        "លុបកំណត់ត្រាទាំងអស់ដែលចាស់ជាង 30 ថ្ងៃ? សកម្មភាពនេះមិនអាចត្រឡប់វិញបានទេ។",
      clearAllConfirm:
        "លុបកំណត់ត្រា Login Activity ទាំងអស់ដែលកំពុងបង្ហាញ? សកម្មភាពនេះមិនអាចត្រឡប់វិញបានទេ។",
      deleteConfirm: "លុបកំណត់ត្រានេះ?",
      deleteBlocked:
        "មិនអាចលុបបានទេ — ប្រព័ន្ធ database បដិសេធសំណើនេះ (RLS policy)។ សូមពិនិត្យមើលការកំណត់សិទ្ធិលុបសម្រាប់ Super Admin នៅលើតារាង login_activity។",
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
      announcements: "Announcements",
      employees: "Employees",
      departments: "Departments",
      shifts: "Shifts",
      attendance: "Attendance",
      holidays: "Public Holidays",
      leave: "Leave Requests",
      overtime: "Overtime (OT)",
      payroll: "Payroll",
      performance: "Performance Reviews",
      attCorrection: "Attendance Corrections",
      admins: "Admin Accounts",
      myAttendance: "My Attendance",
      myLeave: "My Leave",
      myOvertime: "My Overtime (OT)",
      myPayroll: "My Payroll",
      myPerformance: "My Performance Reviews",
      myDocuments: "My Documents",
      myAttCorrection: "Attendance Correction",
      myProfile: "My Profile",
      settings: "Settings",
      audits: "Audit Log",
      loginActivity: "Login Activity",
      analytics: "Analytics",
      rolePerms: "Roles & Permissions",
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
    exportCsv: "Export CSV",
    exportExcel: "Export Excel",
    pagination: { of: "of" },
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
    analytics: {
      title: "Analytics",
      subtitle: "Overview of attendance, payroll and OT trends",
      attendTrend: "Attendance Rate (last 6 months)",
      attendTrendSub: "Present/late records vs. total records logged per month",
      deptCost: "Net Payroll Cost by Department",
      deptCostSub: "Current month, active employees only",
      otTrend: "Approved OT Hours (last 6 months)",
      otTrendSub: "Total approved overtime hours across all departments",
      noChartData: "Not enough data to show yet",
      totalOt: "Total OT",
      hours: "hrs",
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
      printBadge: "Print ID Badge",
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
      approvedBy: "Approved by",
      rejectedBy: "Rejected by",
      rejectTitle: "Reject Leave Request",
      rejectReason: "Rejection Reason",
      rejectReasonPlaceholder: "Please state the rejection reason...",
      rejectReasonRequired: "Please enter a rejection reason",
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
        "Set the tax and insurance percentage rates, and the minimum base salary at which they start applying. Employees whose base salary is below the threshold are not deducted.",
      taxRateLabel: "Tax Rate (%)",
      insuranceRateLabel: "Insurance Rate (%)",
      minSalaryThresholdLabel: "Applies From Salary ($)",
      minSalaryThresholdHint:
        "Employees earning less than this base salary won't have tax/insurance deducted",
      noThreshold: "Applies to all employees",
      downloadPdf: "Download PDF",
      payslipTitle: "Payslip",
      payPeriod: "Pay Period",
      generatedOn: "Generated on",
      customRateToggle: "Custom tax/insurance rate for this employee",
      customRateHint:
        "When enabled, this employee uses the rates set below instead of the company-wide policy (the minimum salary threshold is ignored)",
      customTaxRateLabel: "Custom Tax Rate (%)",
      customInsuranceRateLabel: "Custom Insurance Rate (%)",
      customRateBadge: "Custom rate",
      historicalBtn: "Older Months Report",
      historicalTitle: "Historical Payroll Report",
      historicalDesc:
        "Pick any month/year to pull attendance directly from the database (not limited to the last 6 months kept live in the app)",
      historicalPick: "Month / Year",
      historicalLoad: "Load Data",
      historicalLoading: "Loading data...",
      historicalEmpty: "No attendance records found for this month",
      historicalError: "Couldn't load this data — please try again",
    },
    pr: {
      addBtn: "Add Review",
      editTitle: "Edit Review",
      addTitle: "Add Performance Review",
      employee: "Employee",
      period: "Review Period",
      periodPlaceholder: "e.g. Q1 2026",
      rating: "Rating (stars)",
      strengths: "Strengths",
      strengthsPlaceholder: "Note the employee's strengths...",
      improvements: "Areas to Improve",
      improvementsPlaceholder: "Note areas for improvement...",
      reviewedBy: "Reviewed by",
      confirmDel: "Are you sure you want to delete this review?",
      noReview: "No reviews yet",
      notifTitle: "You received a new performance review",
    },
    ann: {
      addBtn: "New Announcement",
      editTitle: "Edit Announcement",
      addTitle: "New Announcement",
      titleLabel: "Title",
      titlePlaceholder: "e.g. Year-end holiday schedule",
      bodyLabel: "Message",
      bodyPlaceholder: "Write the announcement...",
      postedBy: "Posted by",
      confirmDel: "Are you sure you want to delete this announcement?",
      noAnn: "No announcements yet",
      notifTitle: "New company announcement",
    },
    doc: {
      title: "Employee Documents",
      addBtn: "Upload Document",
      category: "Document Type",
      catCv: "CV / Resume",
      catContract: "Contract",
      catId: "ID Card",
      catOther: "Other",
      chooseFile: "Choose File",
      uploadedBy: "Uploaded by",
      confirmDel: "Are you sure you want to delete this document?",
      noDocs: "No documents yet",
      tooLarge: "File too large — please choose a file under 4MB",
      view: "View / Download",
      myTitle: "My Documents",
    },
    hol: {
      addBtn: "Add Holiday",
      editTitle: "Edit Public Holiday",
      addTitle: "Add Public Holiday",
      dateLabel: "Date",
      nameLabel: "Holiday Name",
      namePlaceholder: "e.g. Khmer New Year",
      confirmDel: "Are you sure you want to delete this holiday?",
      noHolidays: "No public holidays set yet",
      upcoming: "Upcoming Holidays",
      past: "Past Holidays",
      todayBanner: "Today is a public holiday",
    },
    ac: {
      addBtn: "Request Correction",
      date: "Date",
      requestedCheckIn: "Requested Check In",
      requestedCheckOut: "Requested Check Out",
      reason: "Reason",
      reasonPlaceholder: "e.g. Forgot to clock in/out...",
      approve: "Approve",
      reject: "Reject",
      approvedBy: "Approved by",
      rejectedBy: "Rejected by",
      rejectTitle: "Reject Correction Request",
      rejectReason: "Rejection Reason",
      rejectReasonPlaceholder: "Please state the rejection reason...",
      rejectReasonRequired: "Please enter a rejection reason",
      confirmDel: "Are you sure you want to delete this request?",
      noRequest: "No correction requests yet",
      needOneField: "Please enter at least a check-in or check-out time",
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
      saveFailed: "Save failed:",
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
    audit: {
      title: "Audit Log",
      subtitle: "Track who changed what, and when",
      refresh: "Refresh",
      time: "Time",
      actor: "Actor",
      action: "Action",
      entity: "Entity",
      details: "Details",
      noLogs: "No log entries yet",
      loading: "Loading...",
      filterTable: "Filter by type",
      allTables: "All",
      actionCreate: "Created",
      actionUpdate: "Updated",
      actionDelete: "Deleted",
      actionLogin: "Logged in",
      actionLogout: "Logged out",
      changedFrom: "from",
      changedTo: "to",
      unknownActor: "Unknown",
      tables: {
        departments: "Department",
        employees: "Employee",
        shifts: "Shift",
        holidays: "Holiday",
        leave_requests: "Leave Request",
        overtime_requests: "Overtime (OT)",
        performance_reviews: "Performance Review",
        announcements: "Announcement",
        admins: "Admin Account",
        offices: "Office Location",
        attendance_corrections: "Attendance Correction",
        employee_documents: "Employee Document",
      },
    },
    loginAct: {
      title: "Login Activity",
      subtitle: "Your sign-in and sign-out history, including the device used",
      subtitleAll: "Sign-in and sign-out history for all users",
      scopeMine: "Mine",
      scopeAll: "All users",
      refresh: "Refresh",
      time: "Time",
      actor: "User",
      action: "Action",
      device: "Device",
      os: "OS",
      model: "Model (Android)",
      browser: "Browser",
      noLogs: "No log entries yet",
      loading: "Loading...",
      actionLogin: "Logged in",
      actionLogout: "Logged out",
      unknownActor: "Unknown",
      deviceMobile: "Mobile",
      deviceTablet: "Tablet",
      deviceDesktop: "Desktop",
      unknown: "Unknown",
      active: "Active",
      currentDevice: "This device",
      revoke: "Not me — Sign out now",
      revokeConfirm: "Force-sign this device out of the account right now?",
      revoked: "Signed out",
      revokedAlert: "Your session was force-signed-out from another device.",
      clearOld: "Clear old (30+ days)",
      clearAll: "Clear all",
      clearOldConfirm:
        "Delete every log entry older than 30 days? This can't be undone.",
      clearAllConfirm:
        "Delete every Login Activity entry currently in view? This can't be undone.",
      deleteConfirm: "Delete this log entry?",
      deleteBlocked:
        "Delete was blocked by the database (RLS policy) — the entry is still there. Check that Super Admin has delete permission on the login_activity table.",
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
// Live ledger-style clock shown in the header — a small functional nod to
// this being a time & attendance system, not just decoration: admins and
// employees glance at it to sanity-check punch times against the system
// clock. Ticks every second; unmounts cleanly via the interval cleanup.
function HeaderClock() {
  const [now, setNow] = useState(() => new Date());
  useEffect(() => {
    const id = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(id);
  }, []);
  const hh = String(now.getHours()).padStart(2, "0");
  const mm = String(now.getMinutes()).padStart(2, "0");
  const ss = String(now.getSeconds()).padStart(2, "0");
  return (
    <div
      style={{
        display: "flex",
        alignItems: "baseline",
        gap: 5,
        fontFamily: "'JetBrains Mono',monospace",
        fontSize: 13,
        fontWeight: 600,
        color: T.text,
        padding: "5px 10px",
        borderRadius: 8,
        border: `1px solid ${T.line}`,
        background: T.tableHeadBg,
        letterSpacing: "-.01em",
      }}
    >
      <span style={{ width: 6, height: 6, borderRadius: 999, background: T.forest, flexShrink: 0 }} />
      {hh}:{mm}
      <span style={{ color: T.muted, fontSize: 11 }}>{ss}</span>
    </div>
  );
}

/* ---------------------------------------------------------------
   Tokens
----------------------------------------------------------------*/
// BRAND is the fixed navy used for "always-dark" chrome — the sidebar,
// primary buttons, and hero banners — which stays the same regardless of
// light/dark mode (white text sits on it either way).
const BRAND = {
  ink: "#0A0F1A",
  inkDark: "#050810",
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
  forest: "#1FA26B",
  forestDark: "#168053",
  forestSoft: "var(--wf-forest-soft)",
  forestText: "var(--wf-forest-text)",
  clay: "#D9622E",
  gold: "#F0A83B",
  goldSoft: "var(--wf-gold-soft)",
  goldText: "var(--wf-gold-text)",
  rose: "#E5637A",
  roseDark: "var(--wf-rose-dark)",
  roseSoft: "var(--wf-rose-soft)",
  blue: "#5B8DEF",
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
  "#1FA26B",
  "#F0A83B",
  "#5B8DEF",
  "#E5637A",
  "#35D0BA",
  "#D9622E",
];
// Legacy single admin password — kept only as a fallback reference; login now
// checks against the ADMINS list below, which supports multiple accounts
// with different permission levels.
const ADMIN_PASSWORD = "admin123";
// Seven configurable HR ranks below Superadmin, in ascending seniority
// order ("staff" is the lowest — read-only by default, meant for admin
// portal accounts that shouldn't act on anything until Superadmin grants
// specific modules). Superadmin sits above all of them, always has every
// permission, and is the only rank that can manage other admin accounts
// (kept hardcoded, not configurable, since granting "manage admins" to a
// lower rank would let that rank promote itself — a privilege-escalation
// hole).
const ADMIN_RANKS = [
  "staff",
  "officer",
  "senior",
  "supervisor",
  "manager",
  "seniorManager",
  "admin",
];
const ADMIN_ROLE_LABEL = {
  km: {
    staff: "បុគ្គលិក (Employee)",
    officer: "មន្ត្រី HR (Officer)",
    senior: "មន្ត្រីជាន់ខ្ពស់ (Senior)",
    supervisor: "ប្រធានក្រុម (Supervisor)",
    manager: "អ្នកគ្រប់គ្រង HR (Manager)",
    seniorManager: "អ្នកគ្រប់គ្រងជាន់ខ្ពស់ (Senior Manager)",
    admin: "អ្នកគ្រប់គ្រងប្រព័ន្ធ (Admin)",
    superadmin: "អ្នកគ្រប់គ្រងកំពូល (Superadmin)",
  },
  en: {
    staff: "Employee",
    officer: "HR Officer",
    senior: "Senior Officer",
    supervisor: "Supervisor",
    manager: "HR Manager",
    seniorManager: "Senior Manager",
    admin: "Admin",
    superadmin: "Superadmin",
  },
};
// The permission modules Superadmin can grant/revoke per rank. Each key
// corresponds to a real gate somewhere in the app (see PERMISSION_KEY
// usage below) — "manageAdmins" is deliberately absent here; that stays
// superadmin-only and isn't configurable, for the reason noted above.
const PERMISSION_MODULES = [
  "manageDepartments",
  "manageEmployees",
  "approveRequests",
  "managePayroll",
  "manageDocuments",
  "manageAnnouncements",
  "manageSettings",
  "viewAuditLog",
];
const PERMISSION_LABEL = {
  km: {
    manageDepartments: "គ្រប់គ្រងនាយកដ្ឋាន/វេន",
    manageEmployees: "គ្រប់គ្រងបុគ្គលិក",
    approveRequests: "អនុម័តច្បាប់/OT/កែតម្រូវវត្តមាន",
    managePayroll: "គ្រប់គ្រងប្រាក់ខែ",
    manageDocuments: "គ្រប់គ្រងឯកសារបុគ្គលិក",
    manageAnnouncements: "គ្រប់គ្រងសេចក្តីប្រកាស",
    manageSettings: "គ្រប់គ្រងការកំណត់ក្រុមហ៊ុន",
    viewAuditLog: "មើលកំណត់ត្រាសកម្មភាព",
  },
  en: {
    manageDepartments: "Manage Departments/Shifts",
    manageEmployees: "Manage Employees",
    approveRequests: "Approve Leave/OT/Corrections",
    managePayroll: "Manage Payroll",
    manageDocuments: "Manage Employee Documents",
    manageAnnouncements: "Manage Announcements",
    manageSettings: "Manage Company Settings",
    viewAuditLog: "View Audit Log",
  },
};
// Fixed id under which the Employee Self-Service module toggles are saved
// inside the same role_permissions table/array as the admin rank matrix
// (id column is free text, so this just piggybacks on the existing
// storage — no new table needed). This is NOT an admin rank; it's a
// single company-wide on/off switchboard that applies to every regular
// employee using the self-service portal (buildNavEmployee), since
// employees don't have seniority ranks the way admin accounts do.
const EMPLOYEE_MODULES_ID = "employeePortal";
// The self-service pages Superadmin can hide company-wide. "dashboard" is
// deliberately absent — it's the landing page every employee needs, so it
// always stays on and isn't configurable.
const EMPLOYEE_MODULES = [
  "announcements",
  "attendance",
  "leave",
  "ot",
  "payroll",
  "review",
  "attcorr",
  "documents",
  "loginActivity",
  "profile",
];
const EMPLOYEE_MODULE_LABEL = {
  km: {
    announcements: "សេចក្តីប្រកាស",
    attendance: "វត្តមានផ្ទាល់ខ្លួន",
    leave: "សំណើច្បាប់ឈប់សម្រាក",
    ot: "ម៉ោងបន្ថែម (OT)",
    payroll: "ប្រាក់ខែផ្ទាល់ខ្លួន",
    review: "ការវាយតម្លៃការងារ",
    attcorr: "សំណើកែតម្រូវវត្តមាន",
    documents: "ឯកសារផ្ទាល់ខ្លួន",
    loginActivity: "សកម្មភាពចូលប្រើ",
    profile: "ប្រវត្តិរូបផ្ទាល់ខ្លួន",
  },
  en: {
    announcements: "Announcements",
    attendance: "My Attendance",
    leave: "My Leave",
    ot: "My Overtime (OT)",
    payroll: "My Payroll",
    review: "My Performance Reviews",
    attcorr: "Attendance Correction",
    documents: "My Documents",
    loginActivity: "Login Activity",
    profile: "My Profile",
  },
};
function employeeModuleLabel(key, lang) {
  return (EMPLOYEE_MODULE_LABEL[lang] || EMPLOYEE_MODULE_LABEL.km)[key] || key;
}
// Every module defaults to visible — nothing disappears for existing
// companies until Superadmin explicitly unchecks something.
const DEFAULT_EMPLOYEE_MODULES = {
  // The employeePortal row lives in the same role_permissions table as
  // the admin rank rows, so it must also carry these manage_* columns
  // (all false — they're meaningless for this row, but the columns are
  // NOT NULL) or the upsert fails with a not-null constraint violation.
  manageDepartments: false,
  manageEmployees: false,
  approveRequests: false,
  managePayroll: false,
  manageDocuments: false,
  manageAnnouncements: false,
  manageSettings: false,
  viewAuditLog: false,
  announcements: true,
  attendance: true,
  leave: true,
  ot: true,
  payroll: true,
  review: true,
  attcorr: true,
  documents: true,
  loginActivity: true,
  profile: true,
};
// Sensible starting point covering the natural HR seniority ladder — an
// Officer can only view, each step up adds more, Admin gets everything
// short of managing other admin accounts. Superadmin edits this freely
// from the Roles & Permissions page; these are just the seed defaults
// used until Superadmin has saved a custom matrix (and as a fallback for
// any rank the matrix doesn't yet have a row for).
const DEFAULT_ROLE_PERMISSIONS = {
  staff: {
    manageDepartments: false,
    manageEmployees: false,
    approveRequests: false,
    managePayroll: false,
    manageDocuments: false,
    manageAnnouncements: false,
    manageSettings: false,
    viewAuditLog: false,
  },
  officer: {
    manageDepartments: false,
    manageEmployees: false,
    approveRequests: false,
    managePayroll: false,
    manageDocuments: false,
    manageAnnouncements: false,
    manageSettings: false,
    viewAuditLog: false,
  },
  senior: {
    manageDepartments: false,
    manageEmployees: true,
    approveRequests: false,
    managePayroll: false,
    manageDocuments: true,
    manageAnnouncements: false,
    manageSettings: false,
    viewAuditLog: false,
  },
  supervisor: {
    manageDepartments: false,
    manageEmployees: true,
    approveRequests: true,
    managePayroll: false,
    manageDocuments: true,
    manageAnnouncements: false,
    manageSettings: false,
    viewAuditLog: false,
  },
  manager: {
    manageDepartments: true,
    manageEmployees: true,
    approveRequests: true,
    managePayroll: true,
    manageDocuments: true,
    manageAnnouncements: true,
    manageSettings: false,
    viewAuditLog: false,
  },
  seniorManager: {
    manageDepartments: true,
    manageEmployees: true,
    approveRequests: true,
    managePayroll: true,
    manageDocuments: true,
    manageAnnouncements: true,
    manageSettings: true,
    viewAuditLog: true,
  },
  admin: {
    manageDepartments: true,
    manageEmployees: true,
    approveRequests: true,
    managePayroll: true,
    manageDocuments: true,
    manageAnnouncements: true,
    manageSettings: true,
    viewAuditLog: true,
  },
};
// Admin role names follow the app's km/en language toggle (like every
// other label) rather than being hardcoded to Khmer everywhere.
function adminRoleLabel(role, lang) {
  return (ADMIN_ROLE_LABEL[lang] || ADMIN_ROLE_LABEL.km)[role] || role || "";
}
function permissionLabel(key, lang) {
  return (PERMISSION_LABEL[lang] || PERMISSION_LABEL.km)[key] || key;
}
// Central "can this admin do X" check. Superadmin always passes every
// check (and is the only rank allowed to manage other admin accounts —
// that check bypasses the matrix entirely, see manageAdmins usage
// elsewhere). Everyone else is looked up in the Superadmin-editable
// rolePermissions matrix, falling back to the seed defaults above for
// any rank that isn't in it yet (e.g. right after this feature ships,
// before Superadmin has saved anything).
function canDo(admin, rolePermissions, key) {
  if (!admin) return false;
  if (admin.role === "superadmin") return true;
  if (key === "manageAdmins") return false; // superadmin-only, not configurable
  const row =
    rolePermissions?.[admin.role] || DEFAULT_ROLE_PERMISSIONS[admin.role];
  return !!row?.[key];
}
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
  --wf-ink:#10141C; --wf-ink-dark:#050810; --wf-paper:#F3F4F7; --wf-card:#FFFFFF;
  --wf-forest-soft:#E4F5EC; --wf-forest-text:#127449; --wf-gold-soft:#FCF0DC; --wf-gold-text:#9A6212;
  --wf-rose-dark:#B23752; --wf-rose-soft:#FBEAEE; --wf-line:#E2E5EB; --wf-line-soft:#EAECF1;
  --wf-muted:#767E8F; --wf-muted-light:#A7ADBB; --wf-text-soft:#3C4250;
  --wf-input-border:#D6DAE2; --wf-input-bg:#FAFBFC; --wf-field-label:#5B6274;
  --wf-table-head-bg:#F6F7F9; --wf-divider:#ECEEF2; --wf-danger-border:#F0C7D0;
  --wf-danger-hover-bg:#FBEEF1; --wf-header-bg:rgba(255,255,255,0.85);
}
.wf-dark{
  --wf-ink:#EEF1F6; --wf-ink-dark:#AEB6C7; --wf-paper:#080B12; --wf-card:#0E121B;
  --wf-forest-soft:#0E2A20; --wf-forest-text:#3FD996; --wf-gold-soft:#2E2211; --wf-gold-text:#F5BE5F;
  --wf-rose-dark:#F0879B; --wf-rose-soft:#2C151B; --wf-line:#1C2230; --wf-line-soft:#161B27;
  --wf-muted:#717A90; --wf-muted-light:#414A5E; --wf-text-soft:#B9C0D2;
  --wf-input-border:#1F2634; --wf-input-bg:#0B0F17; --wf-field-label:#8791A8;
  --wf-table-head-bg:#0B0F18; --wf-divider:#171D29; --wf-danger-border:#3D1D26;
  --wf-danger-hover-bg:#20121A; --wf-header-bg:rgba(8,11,18,0.82);
}
.wf-root{display:flex;height:100vh;height:100dvh;min-height:640px;max-height:100vh;max-height:100dvh;background:${T.paper};font-family:'Inter','Noto Sans Khmer',sans-serif;color:${T.text};position:relative;overflow:hidden;border-radius:10px;box-shadow:0 1px 0 rgba(0,0,0,0.02),0 16px 40px -18px rgba(5,8,16,0.35);border:1px solid ${T.line};transition:background .15s ease,color .15s ease;}
.wf-sidebar{background:linear-gradient(180deg,${BRAND.ink} 0%,${BRAND.inkDark} 100%);color:#fff;width:246px;flex-shrink:0;display:flex;flex-direction:column;border-right:1px solid rgba(255,255,255,0.06);transition:transform .25s cubic-bezier(.4,0,.2,1);}
.wf-sidebar-inner{display:flex;flex-direction:column;height:100%;}
.wf-logo-badge{width:32px;height:32px;border-radius:7px;background:${T.gold};display:flex;align-items:center;justify-content:center;font-weight:700;font-size:12px;color:#1A1300;font-family:'JetBrains Mono',monospace;flex-shrink:0;box-shadow:inset 0 0 0 1px rgba(255,255,255,0.25);}
.wf-nav-eyebrow{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:#4C5670;padding:4px 14px 8px;}
.wf-nav-item{position:relative;width:100%;display:flex;align-items:center;gap:11px;padding:9px 14px;border-radius:7px;font-size:13.5px;font-weight:500;background:transparent;color:#8891A6;border:none;cursor:pointer;text-align:left;transition:background .15s ease,color .15s ease;}
.wf-nav-item:hover{background:rgba(255,255,255,0.045);color:#fff;}
.wf-nav-item.active{background:rgba(240,168,59,0.09);color:#fff;font-weight:600;}
.wf-nav-item.active::before{content:"";position:absolute;left:-10px;top:6px;bottom:6px;width:2px;border-radius:0;background:${T.gold};}
.wf-main{flex:1;min-width:0;min-height:0;display:flex;flex-direction:column;background:${T.paper};}
.wf-header{background:${T.headerBg};backdrop-filter:blur(8px);border-bottom:1px solid ${T.lineSoft};padding:13px 22px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:20;transition:background .15s ease,border-color .15s ease;}
.wf-content{flex:1;overflow-y:auto;padding:22px;}
.wf-card{background:${T.card};border-radius:9px;border:1px solid ${T.line};box-shadow:none;transition:box-shadow .15s ease,background .15s ease,border-color .15s ease;}
.wf-btn{display:inline-flex;align-items:center;gap:6px;font-weight:600;border-radius:7px;font-size:13px;padding:9px 15px;border:1px solid transparent;cursor:pointer;transition:background .15s ease,transform .1s ease,box-shadow .15s ease;}
.wf-btn:active:not(:disabled){transform:scale(.97);}
.wf-btn:disabled{opacity:.5;cursor:not-allowed;}
.wf-btn-sm{padding:6px 10px;font-size:12px;}
.wf-btn-primary{background:${T.gold};color:#1A1300;box-shadow:0 1px 2px rgba(240,168,59,0.25);}
.wf-btn-primary:hover:not(:disabled){background:#D89430;}
.wf-btn-accent{background:${T.forest};color:#fff;box-shadow:0 1px 2px rgba(31,162,107,0.22);}
.wf-btn-accent:hover:not(:disabled){background:${T.forestDark};}
.wf-btn-ghost{background:transparent;color:${T.ink};border-color:${T.line};}
.wf-btn-ghost:hover:not(:disabled){background:${T.tableHeadBg};}
.wf-btn-danger{background:transparent;color:${T.rose};border-color:${T.dangerBorder};}
.wf-btn-danger:hover:not(:disabled){background:${T.dangerHoverBg};}
.wf-btn-danger-solid{background:${T.rose};color:#fff;}
.wf-btn-danger-solid:hover:not(:disabled){background:${T.roseDark};}
.wf-input{width:100%;padding:9px 12px;border-radius:7px;border:1px solid ${T.inputBorder};font-size:13px;background:${T.inputBg};color:${T.text};outline:none;font-family:inherit;transition:border-color .15s ease,box-shadow .15s ease,background .15s ease;}
.wf-input:focus{border-color:${T.gold};box-shadow:0 0 0 3px rgba(240,168,59,0.16);}
.wf-field-label{display:block;font-size:11px;font-weight:700;color:${T.fieldLabel};margin-bottom:6px;text-transform:uppercase;letter-spacing:.05em;}
.wf-dp-wrap{position:relative;display:inline-block;width:100%;}
.wf-dp-trigger{width:100%;padding:9px 12px;border-radius:7px;border:1px solid ${T.inputBorder};font-size:13px;background:${T.inputBg};color:${T.text};outline:none;font-family:inherit;display:flex;align-items:center;justify-content:space-between;gap:8px;cursor:pointer;transition:border-color .15s ease,box-shadow .15s ease,background .15s ease;text-align:left;}
.wf-dp-trigger:hover{border-color:${T.gold};}
.wf-dp-trigger.open{border-color:${T.gold};box-shadow:0 0 0 3px rgba(240,168,59,0.16);}
.wf-dp-placeholder{color:${T.muted};}
.wf-dp-pop{position:absolute;z-index:60;top:calc(100% + 6px);left:0;min-width:264px;background:${T.card};border:1px solid ${T.line};border-radius:10px;box-shadow:0 12px 32px rgba(5,8,16,0.28);padding:12px;animation:wf-pop .15s cubic-bezier(.2,.9,.3,1.2);}
.wf-dp-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;}
.wf-dp-nav{background:none;border:none;cursor:pointer;color:${T.muted};display:flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:6px;transition:background .12s ease,color .12s ease;flex-shrink:0;}
.wf-dp-nav:hover{background:${T.tableHeadBg};color:${T.ink};}
.wf-dp-title{font-size:13px;font-weight:700;color:${T.ink};font-family:'Sora','Noto Sans Khmer',sans-serif;}
.wf-dp-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;}
.wf-dp-dow{font-size:10px;font-weight:700;color:${T.muted};text-align:center;padding:4px 0;text-transform:uppercase;}
.wf-dp-day{font-size:12.5px;text-align:center;padding:7px 0;border-radius:6px;cursor:pointer;color:${T.text};background:none;border:1px solid transparent;transition:background .12s ease,color .12s ease;}
.wf-dp-day:hover{background:${T.tableHeadBg};}
.wf-dp-day.outside{color:${T.mutedLight};}
.wf-dp-day.today{border-color:${T.gold};font-weight:700;}
.wf-dp-day.selected{background:${T.gold};color:#1A1300;font-weight:700;}
.wf-dp-day.selected:hover{background:#D89430;}
.wf-dp-foot{display:flex;align-items:center;justify-content:space-between;margin-top:10px;padding-top:10px;border-top:1px solid ${T.lineSoft};}
.wf-dp-link{background:none;border:none;cursor:pointer;font-size:12px;font-weight:600;color:${T.forestText};padding:2px 4px;border-radius:6px;}
.wf-dp-link:hover{background:${T.forestSoft};}
.wf-tp-pop{position:absolute;z-index:60;top:calc(100% + 6px);left:0;background:${T.card};border:1px solid ${T.line};border-radius:10px;box-shadow:0 12px 32px rgba(5,8,16,0.28);padding:10px;display:flex;flex-direction:column;animation:wf-pop .15s cubic-bezier(.2,.9,.3,1.2);}
.wf-tp-cols{display:flex;align-items:center;gap:6px;}
.wf-tp-col{width:50px;height:168px;overflow-y:auto;scroll-snap-type:y mandatory;border-radius:8px;background:${T.tableHeadBg};scrollbar-width:none;}
.wf-tp-col::-webkit-scrollbar{width:0;height:0;}
.wf-tp-pad{height:68px;flex-shrink:0;scroll-snap-align:none;}
.wf-tp-item{scroll-snap-align:center;text-align:center;padding:8px 0;font-size:13px;font-family:'JetBrains Mono',monospace;color:${T.text};cursor:pointer;border-radius:6px;transition:background .12s ease,color .12s ease;}
.wf-tp-item:hover{background:${T.line};}
.wf-tp-item.selected{background:${T.gold};color:#1A1300;font-weight:700;}
.wf-tp-sep{font-weight:700;color:${T.muted};padding-bottom:2px;}
.wf-modal-overlay{position:fixed;inset:0;z-index:50;display:flex;align-items:center;justify-content:center;padding:16px;background:rgba(3,5,10,0.65);backdrop-filter:blur(2px);animation:wf-fade .15s ease;}
.wf-modal{background:${T.card};border-radius:11px;border:1px solid ${T.line};box-shadow:0 24px 64px rgba(3,5,10,0.5);width:100%;max-height:90vh;overflow-y:auto;animation:wf-pop .18s cubic-bezier(.2,.9,.3,1.2);}
.wf-modal-head{display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid ${T.lineSoft};position:sticky;top:0;background:${T.card};border-radius:11px 11px 0 0;}
.wf-avatar{border-radius:999px;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:600;flex-shrink:0;}
.wf-badge{display:inline-block;padding:4px 10px;border-radius:6px;font-size:11px;font-weight:700;letter-spacing:.02em;white-space:nowrap;}
.wf-table{width:100%;font-size:13px;border-collapse:collapse;}
.wf-table th{text-align:left;font-size:10.5px;color:${T.muted};text-transform:uppercase;padding:11px 16px;background:${T.tableHeadBg};border-bottom:1px solid ${T.lineSoft};font-weight:700;letter-spacing:.05em;font-family:'JetBrains Mono',monospace;}
.wf-table td{padding:10px 16px;border-bottom:1px solid ${T.lineSoft};}
.wf-table tr:last-child td{border-bottom:none;}
.wf-table tbody tr{transition:background .12s ease;}
.wf-table tbody tr:hover{background:${T.tableHeadBg};}
.wf-grid{display:grid;gap:16px;}
.wf-punch-clock{font-size:34px;font-weight:700;font-family:'JetBrains Mono',monospace;color:${T.ink};font-variant-numeric:tabular-nums;letter-spacing:-.01em;}
.wf-menu-btn{display:none;background:none;border:none;color:${T.ink};cursor:pointer;padding:4px;}
.wf-overlay-scrim{display:none;}
.wf-content::-webkit-scrollbar,.wf-sidebar nav::-webkit-scrollbar,.wf-modal::-webkit-scrollbar{width:8px;}
.wf-content::-webkit-scrollbar-thumb,.wf-modal::-webkit-scrollbar-thumb{background:${T.mutedLight};border-radius:8px;}
.wf-content::-webkit-scrollbar-track,.wf-modal::-webkit-scrollbar-track{background:transparent;}
.wf-sidebar nav::-webkit-scrollbar-thumb{background:rgba(255,255,255,0.15);border-radius:8px;}
@keyframes wf-fade{from{opacity:0}to{opacity:1}}
@keyframes wf-pop{from{opacity:0;transform:scale(.96) translateY(6px)}to{opacity:1;transform:scale(1) translateY(0)}}
@keyframes wf-app-in{from{opacity:0;}to{opacity:1;}}
@keyframes wf-page-in{from{opacity:0;}to{opacity:1;}}
.wf-app-enter{animation:wf-app-in .38s ease both;}
.wf-page-enter{animation:wf-page-in .24s ease both;}
.wf-nav-item,.wf-bottomnav-item{transition:background .15s ease,color .15s ease,transform .15s ease;}
.wf-bottomnav-item:active{transform:scale(.93);}
.wf-menu-btn,.wf-btn{transition:background .15s ease,transform .12s ease,box-shadow .15s ease,color .15s ease;}
.wf-bottomnav.wf-bottomnav-hidden{display:none !important;}
.wf-role-badge{white-space:nowrap;flex-shrink:0;}
.wf-bottomnav{display:none;position:absolute;left:0;right:0;bottom:0;z-index:45;background:${T.headerBg};backdrop-filter:blur(10px);border-top:1px solid ${T.lineSoft};align-items:stretch;justify-content:space-around;padding:5px 2px calc(5px + env(safe-area-inset-bottom));box-shadow:0 -2px 12px rgba(5,8,16,0.08);}
.wf-bottomnav-item{position:relative;flex:1;min-width:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:2px;background:none;border:none;cursor:pointer;color:${T.muted};padding:5px 2px 4px;border-radius:9px;font-size:10px;font-weight:600;transition:color .15s ease;}
.wf-bottomnav-item span{max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.wf-bottomnav-item.active{color:${T.goldText};}
.wf-bnav-icon-wrap{width:38px;height:24px;display:flex;align-items:center;justify-content:center;border-radius:8px;transition:background .15s ease;}
.wf-bottomnav-item.active .wf-bnav-icon-wrap{background:${T.goldSoft};}
.wf-bottomnav-item .wf-bnav-dot{position:absolute;top:2px;right:calc(50% - 20px);width:7px;height:7px;border-radius:999px;background:${T.rose};border:1.5px solid ${T.card};}
@media (max-width: 820px){
  .wf-sidebar{position:absolute;inset:0 auto 0 0;z-index:40;transform:translateX(-100%);height:100%;box-shadow:8px 0 24px rgba(0,0,0,0.4);}
  .wf-sidebar.open{transform:translateX(0);}
  .wf-menu-btn{display:inline-flex;}
  .wf-overlay-scrim.open{display:block;position:absolute;inset:0;background:rgba(18,32,61,0.45);z-index:35;backdrop-filter:blur(1px);}
  .wf-header{padding:12px 16px;}
  .wf-content{padding:16px;}
  .wf-role-badge{display:none;}
  .wf-bottomnav{display:flex;}
  .wf-content.wf-content-bnpad{padding-bottom:86px;}
  .wf-role-staff .wf-menu-btn{display:none;}
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
        "https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700;800&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&family=Noto+Sans+Khmer:wght@400;500;600;700&display=swap";
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
  ACTIVE_SESSION_ID: "hrsuite:session:activeId",
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
  // Minimum base salary at which tax/insurance start being deducted.
  // 0 means the deduction applies to every employee regardless of salary.
  minSalaryThreshold: 0,
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
const WEEKDAY_LABELS = ["អា", "ច", "អ", "ព", "ព្រ", "សុ", "ស"];
// True if `dateStr` (YYYY-MM-DD) is a day off for `emp`, either because it
// falls on one of their recurring weekly-off days or matches one of their
// specific one-off days-off dates.
function isDayOff(emp, dateStr) {
  if (!emp || !dateStr) return false;
  if (Array.isArray(emp.customDaysOff) && emp.customDaysOff.includes(dateStr))
    return true;
  if (Array.isArray(emp.weeklyOff) && emp.weeklyOff.length) {
    const [y, m, d] = dateStr.split("-").map(Number);
    if (y && m && d) {
      const dow = new Date(y, m - 1, d).getDay();
      if (emp.weeklyOff.includes(dow)) return true;
    }
  }
  return false;
}
const DEFAULT_ANNUAL_LEAVE_DAYS = 18;
const DEFAULT_SICK_LEAVE_DAYS = 7;
// Number of *approved* leave days of a given `type` ("annual" | "sick")
// an employee has used within `year` (defaults to the current year).
// Counts every calendar day in each approved request's range that falls
// inside that year, so a request spanning New Year's is split correctly
// across both years.
function usedLeaveDaysByType(employeeId, leaveRequests, type, year) {
  const y = year || new Date().getFullYear();
  return leaveRequests
    .filter(
      (r) =>
        r.employeeId === employeeId &&
        r.type === type &&
        r.status === "approved",
    )
    .reduce((sum, r) => {
      const days = dateRange(r.startDate, r.endDate).filter(
        (d) => d.slice(0, 4) === String(y),
      );
      return sum + days.length;
    }, 0);
}
function usedAnnualLeaveDays(employeeId, leaveRequests, year) {
  return usedLeaveDaysByType(employeeId, leaveRequests, "annual", year);
}
function usedSickLeaveDays(employeeId, leaveRequests, year) {
  return usedLeaveDaysByType(employeeId, leaveRequests, "sick", year);
}
function annualLeaveBalance(emp, leaveRequests, year) {
  const quota = Number.isFinite(emp?.annualLeaveDays)
    ? emp.annualLeaveDays
    : DEFAULT_ANNUAL_LEAVE_DAYS;
  const used = usedAnnualLeaveDays(emp?.id, leaveRequests, year);
  return { quota, used, remaining: quota - used };
}
function sickLeaveBalance(emp, leaveRequests, year) {
  const quota = Number.isFinite(emp?.sickLeaveDays)
    ? emp.sickLeaveDays
    : DEFAULT_SICK_LEAVE_DAYS;
  const used = usedSickLeaveDays(emp?.id, leaveRequests, year);
  return { quota, used, remaining: quota - used };
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
  performanceReviews = [],
  announcements = [],
  attendanceCorrections = [],
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
    attendanceCorrections
      .filter((r) => r.status === "pending")
      .forEach((r) => {
        const emp = employees.find((e) => e.id === r.employeeId);
        list.push({
          id: `ac-pending-${r.id}`,
          page: "attcorr",
          tone: "gold",
          title: en
            ? "New attendance correction request"
            : "សំណើកែតម្រូវវត្តមានថ្មី",
          message: en
            ? `${emp?.name || "?"} requested a correction on ${r.date}`
            : `${emp?.name || "?"} បានស្នើសុំកែតម្រូវវត្តមាននៅថ្ងៃទី ${r.date}`,
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
    performanceReviews
      .filter((r) => r.employeeId === currentEmp.id)
      .forEach((r) => {
        list.push({
          id: `pr-new-${r.id}`,
          page: "review",
          tone: "gold",
          title: en
            ? "You received a new performance review"
            : "អ្នកទទួលបានការវាយតម្លៃការងារថ្មី",
          message: r.period,
          time: r.createdAt,
        });
      });
    attendanceCorrections
      .filter(
        (r) =>
          r.employeeId === currentEmp.id &&
          (r.status === "approved" || r.status === "rejected") &&
          r.reviewedAt,
      )
      .forEach((r) => {
        list.push({
          id: `ac-decided-${r.id}`,
          page: "attcorr",
          tone: r.status === "approved" ? "forest" : "rose",
          title:
            r.status === "approved"
              ? en
                ? "Your attendance correction was approved"
                : "សំណើកែតម្រូវវត្តមានរបស់អ្នកត្រូវបានអនុម័ត"
              : en
                ? "Your attendance correction was rejected"
                : "សំណើកែតម្រូវវត្តមានរបស់អ្នកត្រូវបានបដិសេធ",
          message: r.date,
          time: r.reviewedAt,
        });
      });
    announcements.forEach((a) => {
      list.push({
        id: `ann-new-${a.id}`,
        page: "announcements",
        tone: "gold",
        title: en ? "New company announcement" : "សេចក្តីប្រកាសថ្មីពីក្រុមហ៊ុន",
        message: a.title,
        time: a.createdAt,
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
// Suggests a default OT day type from a "YYYY-MM-DD" date: a date on the
// company holiday calendar suggests "holiday"; Saturday/Sunday suggest
// "weekend"; everything else suggests "normal". Employees can still
// override this in the request form.
function suggestDayType(dateStr, holidays) {
  if (!dateStr) return "normal";
  if (isHoliday(dateStr, holidays)) return "holiday";
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
  // An employee can override the company-wide tax/insurance rates with
  // their own. When active, the override applies unconditionally (the
  // policy's minimum-salary threshold is only meant to gate the
  // *default* rates, not an explicit per-employee rate).
  const usesCustomRate = !!emp.useCustomRate;
  let minSalaryThreshold, deductionApplies, taxRate, insuranceRate;
  if (usesCustomRate) {
    minSalaryThreshold = 0;
    deductionApplies = true;
    taxRate = Number(emp.customTaxRate) || 0;
    insuranceRate = Number(emp.customInsuranceRate) || 0;
  } else {
    minSalaryThreshold = Number(policy.minSalaryThreshold) || 0;
    // Tax/insurance only kick in once the employee's base salary reaches the
    // configured threshold. Below it, no deduction is applied at all.
    deductionApplies = (Number(emp.salary) || 0) >= minSalaryThreshold;
    taxRate = deductionApplies ? Number(policy.taxRate) || 0 : 0;
    insuranceRate = deductionApplies ? Number(policy.insuranceRate) || 0 : 0;
  }
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
    minSalaryThreshold,
    deductionApplies,
    usesCustomRate,
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
   Audit log — records who changed what, and when.
   writeAuditLog() is a fire-and-forget insert into the "audit_logs"
   table; it never throws or blocks the UI, so a missing table or a
   flaky connection can't break the rest of the app.
----------------------------------------------------------------*/
// Fields that are large/binary (base64 photos, uploaded files) or too
// noisy to be useful in a diff — never include their raw values.
const AUDIT_SKIP_FIELDS = new Set([
  "photo",
  "logo",
  "dataUrl",
  "pin",
  "password",
]);
function auditValueToText(v) {
  if (v === null || v === undefined || v === "") return "—";
  if (Array.isArray(v)) v = v.join(", ");
  const s = typeof v === "string" ? v : JSON.stringify(v);
  return s.length > 80 ? s.slice(0, 80) + "…" : s;
}
// Shallow field-by-field diff between two rows, skipping large/sensitive
// fields — used to record exactly what changed on an "update" action.
function diffFields(oldRow, newRow) {
  const keys = new Set([
    ...Object.keys(oldRow || {}),
    ...Object.keys(newRow || {}),
  ]);
  const changes = {};
  keys.forEach((k) => {
    if (AUDIT_SKIP_FIELDS.has(k)) return;
    const a = oldRow ? oldRow[k] : undefined;
    const b = newRow ? newRow[k] : undefined;
    if (JSON.stringify(a) === JSON.stringify(b)) return;
    changes[k] = { from: auditValueToText(a), to: auditValueToText(b) };
  });
  return changes;
}
function writeAuditLog({ actor, action, table, entityId, label, changes }) {
  try {
    supabase
      .from("audit_logs")
      .insert({
        actor_type: actor?.type || null,
        actor_id: actor?.id != null ? String(actor.id) : null,
        actor_name: actor?.name || null,
        action,
        entity_table: table,
        entity_id: entityId != null ? String(entityId) : null,
        entity_label: label != null ? String(label).slice(0, 200) : null,
        changes: changes && Object.keys(changes).length ? changes : null,
        created_at: new Date().toISOString(),
      })
      .then(({ error }) => {
        if (error) console.error("[audit] insert failed:", error.message);
      });
  } catch (e) {
    console.error("[audit] insert threw:", e);
  }
}

/* ---------------------------------------------------------------
   Login activity — a dedicated, per-user record of sign-in/sign-out
   events plus a best-effort read of the device used. This is
   deliberately kept separate from audit_logs: audit_logs answers
   "who changed what", this answers "who signed in, from what, and
   when" — the question people actually ask when reviewing account
   security.

   Device detection is done purely client-side via navigator.userAgent.
   Browsers intentionally limit how much they reveal for privacy
   reasons, so treat this as a best-effort hint, not a precise
   fingerprint:
   - deviceType (mobile / tablet / desktop) and the OS/browser name+
     version are generally reliable.
   - The exact hardware model (e.g. "iPhone 15 Pro") is NOT reliably
     available from the browser — iOS Safari never exposes it, and
     Android only exposes it through the newer User-Agent Client
     Hints API, gated behind a permission-like call and not
     guaranteed to be granted. We don't attempt to parse a model out
     of the UA string because doing so silently produces wrong
     answers more often than right ones.
   - IP address / location is NOT available from client-side JS at
     all — that requires a server (e.g. a Supabase Edge Function
     reading the request's forwarded-for header). Add later if
     needed; out of scope here.

   Expected Supabase table (create this before using the page):
     create table login_activity (
       id bigint generated always as identity primary key,
       actor_type text,        -- 'admin' | 'employee'
       actor_id text,
       actor_name text,
       action text,            -- 'login' | 'logout'
       device_type text,       -- 'mobile' | 'tablet' | 'desktop'
       os text,
       browser text,
       device_model text,      -- Android-only, best-effort internal model
                                -- code (e.g. "SM-A536E") via User-Agent
                                -- Client Hints. Always null on iOS/desktop
                                -- and on non-Chromium Android browsers —
                                -- see getAndroidModel() above.
       user_agent text,
       session_id text,        -- random id generated at login, carried
                                -- through to the matching logout row —
                                -- lets a "login" row be revoked to force
                                -- that specific device to sign out
       revoked_at timestamptz, -- set when someone force-signs-out this
                                -- device from the Login Activity page
       created_at timestamptz default now()
     );
   If you already created this table before device_model existed, add
   it with: alter table login_activity add column device_model text;
   Remember to add a row-level-security policy that lets each user
   read only their own rows, plus a superadmin-only policy (or a
   view) for reading everyone's.
----------------------------------------------------------------*/
// A fresh id generated client-side at every login and carried in
// localStorage for the lifetime of that sign-in. It ties this browser's
// session to one login_activity row, so that row can later be flagged
// with revoked_at to force this exact device to sign out remotely.
function newSessionId() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return `sess_${Date.now()}_${Math.random().toString(36).slice(2)}`;
}
function getDeviceInfo() {
  const ua =
    typeof navigator !== "undefined" && navigator.userAgent
      ? navigator.userAgent
      : "";
  if (!ua) return { deviceType: "desktop", os: "", browser: "", userAgent: "" };

  let deviceType = "desktop";
  if (/iPad/i.test(ua) || (/Android/i.test(ua) && !/Mobile/i.test(ua))) {
    deviceType = "tablet";
  } else if (/Mobi|iPhone|iPod|Android/i.test(ua)) {
    deviceType = "mobile";
  }

  let os = "";
  let m;
  if ((m = ua.match(/Windows NT ([\d.]+)/))) {
    const known = {
      "10.0": "Windows 10/11",
      6.3: "Windows 8.1",
      6.1: "Windows 7",
    };
    os = known[m[1]] || `Windows (NT ${m[1]})`;
  } else if ((m = ua.match(/Mac OS X ([\d_]+)/))) {
    os = `macOS ${m[1].replace(/_/g, ".")}`;
  } else if ((m = ua.match(/Android ([\d.]+)/))) {
    os = `Android ${m[1]}`;
  } else if ((m = ua.match(/OS ([\d_]+) like Mac OS X/))) {
    os = `iOS ${m[1].replace(/_/g, ".")}`;
  } else if (/Linux/i.test(ua)) {
    os = "Linux";
  }

  let browser = "";
  if ((m = ua.match(/Edg\/([\d.]+)/))) browser = `Edge ${m[1]}`;
  else if ((m = ua.match(/OPR\/([\d.]+)/))) browser = `Opera ${m[1]}`;
  else if (
    /Chrome\//.test(ua) &&
    !/Chromium/.test(ua) &&
    (m = ua.match(/Chrome\/([\d.]+)/))
  )
    browser = `Chrome ${m[1]}`;
  else if ((m = ua.match(/Firefox\/([\d.]+)/))) browser = `Firefox ${m[1]}`;
  else if (
    /Safari\//.test(ua) &&
    !/Chrome/.test(ua) &&
    (m = ua.match(/Version\/([\d.]+)/))
  )
    browser = `Safari ${m[1]}`;

  return { deviceType, os, browser, userAgent: ua };
}
// Best-effort Android device model (e.g. "SM-A536E") via the User-Agent
// Client Hints API. Chromium-based browsers on Android only (Chrome,
// Edge, Samsung Internet) — not Firefox, and not available on iOS at
// all (Apple never exposes it to web content, by design). The value
// returned is the internal model code, not a friendly marketing name
// (e.g. "SM-A536E" rather than "Galaxy A54") — there's no public,
// always-current code-to-name mapping we can rely on, so we show the
// raw code as-is rather than guessing a name that might be wrong.
async function getAndroidModel(ua) {
  try {
    if (!/Android/i.test(ua)) return null;
    if (typeof navigator === "undefined" || !navigator.userAgentData)
      return null;
    const hints = await navigator.userAgentData.getHighEntropyValues(["model"]);
    return hints?.model || null;
  } catch {
    return null;
  }
}
function writeLoginActivity({ actor, action, sessionId }) {
  try {
    const { deviceType, os, browser, userAgent } = getDeviceInfo();
    getAndroidModel(userAgent).then((deviceModel) => {
      supabase
        .from("login_activity")
        .insert({
          actor_type: actor?.type || null,
          actor_id: actor?.id != null ? String(actor.id) : null,
          actor_name: actor?.name || null,
          action,
          device_type: deviceType,
          os: os || null,
          browser: browser || null,
          device_model: deviceModel || null,
          user_agent: userAgent || null,
          session_id: sessionId || null,
          created_at: new Date().toISOString(),
        })
        .then(({ error }) => {
          if (error)
            console.error("[login-activity] insert failed:", error.message);
        });
    });
  } catch (e) {
    console.error("[login-activity] insert threw:", e);
  }
}

/* ---------------------------------------------------------------
   Supabase-backed persistence hooks
   These replace the old Claude-Artifact-only window.storage hook.
   Each one keeps the same [value, setValue, ready] shape the rest
   of the app already expects, so components below don't change.
----------------------------------------------------------------*/

// ---------------------------------------------------------------
// Shared realtime bus: ONE Supabase Realtime channel per browser tab,
// instead of one private channel per table.
//
// Each useSupabaseArray(table) call used to open its own
// `supabase.channel("realtime:"+table)` — fine with a handful of tables,
// but this app calls useSupabaseArray for 13 different tables on every
// login (departments, employees, shifts, attendance, leave_requests,
// overtime_requests, performance_reviews, announcements,
// employee_documents, holidays, attendance_corrections, admins,
// offices), all mounted together in AppInner regardless of which page
// is showing. That's 13 open channels per signed-in user — at 1000+
// concurrent staff, ~13,000 channels, which can run into a Supabase
// project's realtime connection limits (plan-dependent) and slow down
// delivery for everyone.
//
// Supabase's realtime client supports chaining multiple
// `.on("postgres_changes", { table: ... }, handler)` calls on a SINGLE
// channel before calling `.subscribe()` once — the server multiplexes
// change events for all of them over that one connection. We register
// every known table up front (before subscribing) so hooks can attach
// and detach their own per-table callback at any time, in any order,
// without needing to re-subscribe the channel.
const REALTIME_TABLES = [
  "departments",
  "employees",
  "shifts",
  "attendance",
  "leave_requests",
  "overtime_requests",
  "performance_reviews",
  "announcements",
  "employee_documents",
  "holidays",
  "attendance_corrections",
  "admins",
  "offices",
  "payroll_paid",
  "role_permissions",
];
const realtimeHandlers = new Map(REALTIME_TABLES.map((t) => [t, new Set()]));
let realtimeChannel = null;

function ensureRealtimeChannel() {
  if (realtimeChannel) return realtimeChannel;
  try {
    const channel = supabase.channel("realtime:app-wide");
    REALTIME_TABLES.forEach((table) => {
      channel.on(
        "postgres_changes",
        { event: "*", schema: "public", table },
        (payload) => {
          realtimeHandlers.get(table)?.forEach((fn) => fn(payload));
        },
      );
    });
    channel.subscribe();
    realtimeChannel = channel;
  } catch (err) {
    console.error("[supabase] failed to open shared realtime channel:", err);
  }
  return realtimeChannel;
}
// Registers `handler` for postgres_changes events on `table` and returns
// an unsubscribe function. Tables outside REALTIME_TABLES still work —
// they just fall back to their own small channel — so adding a new
// useSupabaseArray table elsewhere doesn't silently lose realtime; it's
// just worth adding that table name to REALTIME_TABLES above too so it
// shares the one connection like everything else.
function subscribeTableChanges(table, handler) {
  if (!realtimeHandlers.has(table)) {
    realtimeHandlers.set(table, new Set());
    try {
      supabase
        .channel(`realtime:${table}`)
        .on(
          "postgres_changes",
          { event: "*", schema: "public", table },
          (payload) => {
            realtimeHandlers.get(table)?.forEach((fn) => fn(payload));
          },
        )
        .subscribe();
    } catch (err) {
      console.error(`[supabase] realtime subscribe failed for ${table}:`, err);
    }
  } else {
    ensureRealtimeChannel();
  }
  realtimeHandlers.get(table).add(handler);
  return () => {
    realtimeHandlers.get(table)?.delete(handler);
  };
}

// Generic hook for a Supabase table holding an array of rows keyed by `id`.
// setValue is called elsewhere in the app with the FULL next array (never
// an updater function), so on every call we diff against the previous
// array to figure out which rows to upsert and which to delete.
// Passing `audit: true` (plus an `actorRef`) additionally writes a create/
// update/delete entry to the audit log for every row that actually changed —
// no caller elsewhere in the app has to remember to log anything itself.
function useSupabaseArray(
  table,
  {
    fromDb,
    toDb,
    orderBy,
    audit,
    actorRef,
    entityLabel,
    // Optional: scope the initial load (and what a remote realtime event is
    // allowed to add) to rows where `dateField >= today - daysBack days`.
    // Use this for tables that grow without bound over time (attendance,
    // punch logs, etc.) so the app doesn't keep dragging years of history
    // into every session as the company grows. Leave both unset for tables
    // that don't grow that way (employees, departments, ...).
    dateField,
    daysBack,
  } = {},
) {
  const [value, setValueState] = useState([]);
  const [ready, setReady] = useState(false);
  // Surfaces the last save/delete failure so screens like Settings can
  // tell the user "this didn't actually save" instead of showing a
  // success checkmark while Supabase silently rejected the write (e.g. a
  // column that doesn't exist, or a row-level-security policy denial) —
  // previously such failures only went to console.error and the UI kept
  // showing optimistic local state until the next refresh reverted it.
  const [saveError, setSaveError] = useState(null);
  const prevRef = useRef([]);
  const mapFromDb = fromDb || ((r) => r);
  const mapToDb = toDb || ((r) => r);
  const labelOf =
    entityLabel || ((r) => r?.name || r?.title || r?.code || r?.id);
  const cutoffDate =
    dateField && daysBack
      ? new Date(Date.now() - daysBack * 86400000).toISOString().slice(0, 10)
      : null;

  useEffect(() => {
    let cancelled = false;
    (async () => {
      // Supabase/PostgREST caps any single response at its configured
      // "max rows" (1000 by default) and truncates silently past that —
      // no error, just a short result. A company with 1000+ staff can
      // blow past 1000 rows on tables like `attendance` within a day or
      // two, so we page through with .range() and keep fetching until a
      // page comes back shorter than the page size, rather than trusting
      // one request to return everything.
      const PAGE_SIZE = 1000;
      let all = [];
      let offset = 0;
      let pageError = null;
      for (;;) {
        let query = supabase.from(table).select("*");
        if (orderBy) query = query.order(orderBy);
        if (cutoffDate) query = query.gte(dateField, cutoffDate);
        query = query.range(offset, offset + PAGE_SIZE - 1);
        const { data, error } = await query;
        if (cancelled) return;
        if (error) {
          pageError = error;
          break;
        }
        all = all.concat(data || []);
        if (!data || data.length < PAGE_SIZE) break;
        offset += PAGE_SIZE;
      }
      if (pageError) {
        console.error(`[supabase] failed to load ${table}:`, pageError.message);
        prevRef.current = [];
        setValueState([]);
      } else {
        const mapped = all.map(mapFromDb);
        prevRef.current = mapped;
        setValueState(mapped);
      }
      setReady(true);
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [table, cutoffDate]);

  // Live sync: without this, admin and staff only ever see what was on the
  // table at the moment their tab loaded — a staff check-in, an admin's
  // leave-request decision, a new announcement, etc. would sit invisible on
  // every other open tab until someone manually reloaded the page. This
  // registers a handler on the shared realtime bus above (one WebSocket
  // channel per tab, not one per table) and folds each remote insert/
  // update/delete into local state as it happens, so every open admin and
  // staff screen stays live without polling or a manual refresh. Requires
  // realtime to be enabled for the table in Supabase (Database →
  // Replication, or `alter publication supabase_realtime add table <name>;`).
  useEffect(() => {
    if (!ready) return; // wait for the initial load so we don't race it
    const sortIfNeeded = (rows) => {
      if (!orderBy) return rows;
      const sorted = rows.slice().sort((a, b) => {
        const av = a?.[orderBy];
        const bv = b?.[orderBy];
        if (av == null && bv == null) return 0;
        if (av == null) return -1;
        if (bv == null) return 1;
        return av < bv ? -1 : av > bv ? 1 : 0;
      });
      return sorted;
    };
    const handler = (payload) => {
      setValueState((current) => {
        let next;
        if (payload.eventType === "DELETE") {
          const deletedId = payload.old?.id;
          next = current.filter((r) => r.id !== deletedId);
        } else {
          const row = mapFromDb(payload.new);
          // If this table is date-windowed, ignore remote rows that
          // fall outside the window rather than letting them sneak
          // into state (and then potentially get "upserted" back as
          // if they were new/local on the next setValue() call).
          if (
            cutoffDate &&
            dateField &&
            row?.[dateField] &&
            row[dateField] < cutoffDate
          ) {
            next = current;
          } else {
            const idx = current.findIndex((r) => r.id === row.id);
            next =
              idx === -1
                ? [...current, row]
                : current.map((r, i) => (i === idx ? row : r));
            next = sortIfNeeded(next);
          }
        }
        // Keep the local-edit baseline in sync so this remote-origin
        // change isn't mistaken for a pending local edit next time
        // setValue() runs its create/update/delete diff — otherwise
        // a change made on another device would look, to this tab,
        // like something *it* just created/edited and get pushed
        // straight back to Supabase (and logged to the audit trail
        // a second time, credited to the wrong device).
        prevRef.current = next;
        return next;
      });
    };
    const unsubscribe = subscribeTableChanges(table, handler);
    return unsubscribe;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [table, ready]);

  const setValue = useCallback(
    (next) => {
      const prev = prevRef.current;
      prevRef.current = next;
      setValueState(next);

      const nextIds = new Set(next.map((r) => r.id));
      const deletedRows = prev.filter((r) => !nextIds.has(r.id));
      const toDelete = deletedRows.map((r) => r.id);
      const createdRows = [];
      const updatedRows = []; // [{ row, old }]
      next.forEach((r) => {
        const old = prev.find((p) => p.id === r.id);
        if (!old) createdRows.push(r);
        else if (JSON.stringify(old) !== JSON.stringify(r))
          updatedRows.push({ row: r, old });
      });
      const toUpsert = [...createdRows, ...updatedRows.map((u) => u.row)];

      setSaveError(null);
      (async () => {
        if (toDelete.length) {
          const { error } = await supabase
            .from(table)
            .delete()
            .in("id", toDelete);
          if (error) {
            console.error(
              `[supabase] delete failed on ${table}:`,
              error.message,
            );
            setSaveError(error.message);
          }
        }
        if (toUpsert.length) {
          const { error } = await supabase
            .from(table)
            .upsert(toUpsert.map(mapToDb));
          if (error) {
            console.error(
              `[supabase] upsert failed on ${table}:`,
              error.message,
            );
            setSaveError(error.message);
          }
        }
      })();

      if (audit) {
        const actor = actorRef?.current || null;
        createdRows.forEach((r) =>
          writeAuditLog({
            actor,
            action: "create",
            table,
            entityId: r.id,
            label: labelOf(r),
          }),
        );
        updatedRows.forEach(({ row, old }) =>
          writeAuditLog({
            actor,
            action: "update",
            table,
            entityId: row.id,
            label: labelOf(row),
            changes: diffFields(old, row),
          }),
        );
        deletedRows.forEach((r) =>
          writeAuditLog({
            actor,
            action: "delete",
            table,
            entityId: r.id,
            label: labelOf(r),
          }),
        );
      }
    },
    [table, mapToDb, audit, actorRef, labelOf],
  );

  return [value, setValue, ready, saveError];
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

  // Same live-sync reasoning as useSupabaseArray above: without this, a
  // "mark as paid" done by one admin stays invisible to any other admin
  // tab until it's reloaded. Uses the same shared realtime bus as
  // useSupabaseArray (one channel per tab) instead of a private channel.
  useEffect(() => {
    if (!ready) return;
    const handler = (payload) => {
      setValueState((current) => {
        const row = payload.eventType === "DELETE" ? payload.old : payload.new;
        if (!row) return current;
        const key = `${row.employee_id}-${row.month}`;
        const next = { ...current };
        if (payload.eventType === "DELETE") delete next[key];
        else next[key] = row.paid;
        prevRef.current = next;
        return next;
      });
    };
    const unsubscribe = subscribeTableChanges("payroll_paid", handler);
    return unsubscribe;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ready]);

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
          minSalaryThreshold:
            data.min_salary_threshold ??
            DEFAULT_PAYROLL_POLICY.minSalaryThreshold,
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
        min_salary_threshold: next.minSalaryThreshold,
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
// Reusable pagination: give it the full (already filtered/sorted) array
// and a page size, get back just the slice for the current page plus
// everything needed to render controls. Resets to page 1 whenever the
// input list's length changes (e.g. a new search/filter narrows results)
// so users never land on a page that no longer exists.
function usePagination(items, pageSize = 20) {
  const [page, setPage] = useState(1);
  const total = items.length;
  const pageCount = Math.max(1, Math.ceil(total / pageSize));
  useEffect(() => {
    setPage(1);
  }, [total, pageSize]);
  const safePage = Math.min(page, pageCount);
  const pageItems = useMemo(() => {
    const start = (safePage - 1) * pageSize;
    return items.slice(start, start + pageSize);
  }, [items, safePage, pageSize]);
  return {
    page: safePage,
    setPage,
    pageCount,
    total,
    pageItems,
    rangeStart: total === 0 ? 0 : (safePage - 1) * pageSize + 1,
    rangeEnd: Math.min(safePage * pageSize, total),
  };
}
function Pagination({ page, pageCount, setPage, total, rangeStart, rangeEnd }) {
  const { t } = useLang();
  if (total === 0 || pageCount <= 1) return null;
  const go = (p) => setPage(Math.min(pageCount, Math.max(1, p)));
  return (
    <div
      style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        flexWrap: "wrap",
        gap: 10,
        marginTop: 14,
        fontSize: 12.5,
        color: T.muted,
      }}
    >
      <span>
        {rangeStart}–{rangeEnd} {t.pagination?.of || "នៃ"} {total}
      </span>
      <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
        <Button
          size="sm"
          variant="ghost"
          onClick={() => go(page - 1)}
          disabled={page <= 1}
          style={{ opacity: page <= 1 ? 0.4 : 1 }}
        >
          <ChevronLeft size={14} />
        </Button>
        <span
          style={{
            fontSize: 12.5,
            color: T.ink,
            padding: "0 6px",
            fontFamily: "'JetBrains Mono',monospace",
          }}
        >
          {page} / {pageCount}
        </span>
        <Button
          size="sm"
          variant="ghost"
          onClick={() => go(page + 1)}
          disabled={page >= pageCount}
          style={{ opacity: page >= pageCount ? 0.4 : 1 }}
        >
          <ChevronRight size={14} />
        </Button>
      </div>
    </div>
  );
}
const NOTIF_TONE = {
  gold: "#F0A83B",
  rose: "#E5637A",
  forest: "#1FA26B",
  blue: "#5B8DEF",
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
  performanceReviews,
  announcements,
  attendanceCorrections,
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
        performanceReviews,
        announcements,
        attendanceCorrections,
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
      performanceReviews,
      announcements,
      attendanceCorrections,
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
              fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
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
// Styled replacement for window.confirm(), used only by LoginActivityPage
// (named LoginAct* to avoid colliding with the app-wide ConfirmDialog
// used elsewhere for delete confirmations — that one hardcodes its
// confirm button to "Delete", which doesn't fit actions like "Revoke").
function LoginActConfirmDialog({
  title,
  message,
  danger,
  confirmLabel,
  cancelLabel,
  onConfirm,
  onCancel,
}) {
  const { t } = useLang();
  return (
    <Modal title={title || t.confirmDelete} onClose={onCancel} width={400}>
      <p
        style={{
          fontSize: 13.5,
          color: T.ink,
          margin: "0 0 20px",
          lineHeight: 1.5,
        }}
      >
        {message}
      </p>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Button variant="ghost" onClick={onCancel}>
          {cancelLabel || t.cancel}
        </Button>
        <Button
          variant={danger ? "danger-solid" : "primary"}
          onClick={onConfirm}
        >
          {confirmLabel || t.delete}
        </Button>
      </div>
    </Modal>
  );
}
// Styled replacement for window.alert(). Same visual language as
// LoginActConfirmDialog but with a single acknowledgement button.
function LoginActAlertDialog({ title, message, closeLabel, onClose }) {
  const { t } = useLang();
  return (
    <Modal title={title || t.appName} onClose={onClose} width={400}>
      <p
        style={{
          fontSize: 13.5,
          color: T.ink,
          margin: "0 0 20px",
          lineHeight: 1.5,
        }}
      >
        {message}
      </p>
      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <Button variant="primary" onClick={onClose}>
          {closeLabel || "OK"}
        </Button>
      </div>
    </Modal>
  );
}
function Input(props) {
  return <input className="wf-input" {...props} />;
}
function Select(props) {
  return <select className="wf-input" {...props} />;
}
/* ---------------------------------------------------------------
   Custom Date / Time pickers (replace native input[type=date/time])
----------------------------------------------------------------*/
function parseYMD(s) {
  if (!s) return null;
  const [y, m, d] = s.split("-").map(Number);
  if (!y || !m || !d) return null;
  return new Date(y, m - 1, d);
}
function fmtYMD(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function fmtDateDisplay(s) {
  const d = parseYMD(s);
  if (!d) return "";
  try {
    return new Intl.DateTimeFormat("km-KH", {
      year: "numeric",
      month: "short",
      day: "numeric",
    }).format(d);
  } catch {
    return s;
  }
}
const DP_DOW = ["អា", "ច", "អ", "ព", "ព្រ", "សុ", "ស"];

function useCloseOnOutside(ref, onClose) {
  useEffect(() => {
    function onDocClick(e) {
      if (ref.current && !ref.current.contains(e.target)) onClose();
    }
    function onEsc(e) {
      if (e.key === "Escape") onClose();
    }
    document.addEventListener("mousedown", onDocClick);
    document.addEventListener("keydown", onEsc);
    return () => {
      document.removeEventListener("mousedown", onDocClick);
      document.removeEventListener("keydown", onEsc);
    };
  }, [ref, onClose]);
}

function DatePicker({ value, onChange, placeholder, style, disabled }) {
  const [open, setOpen] = useState(false);
  const selected = parseYMD(value);
  const [cursor, setCursor] = useState(() => selected || new Date());
  const wrapRef = useRef(null);
  useCloseOnOutside(wrapRef, () => setOpen(false));

  useEffect(() => {
    if (open) setCursor(selected || new Date());
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const fire = (d) => onChange && onChange({ target: { value: fmtYMD(d) } });

  const year = cursor.getFullYear();
  const month = cursor.getMonth();
  const firstDow = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const daysInPrevMonth = new Date(year, month, 0).getDate();
  const totalCells = Math.ceil((firstDow + daysInMonth) / 7) * 7;
  const today = todayStr();

  const cells = [];
  for (let i = 0; i < firstDow; i++) {
    const day = daysInPrevMonth - firstDow + 1 + i;
    cells.push({ day, outside: true, date: new Date(year, month - 1, day) });
  }
  for (let d = 1; d <= daysInMonth; d++) {
    cells.push({ day: d, outside: false, date: new Date(year, month, d) });
  }
  let nextDay = 1;
  while (cells.length < totalCells) {
    cells.push({
      day: nextDay,
      outside: true,
      date: new Date(year, month + 1, nextDay),
    });
    nextDay++;
  }

  let monthTitle = `${year}-${month + 1}`;
  try {
    monthTitle = new Intl.DateTimeFormat("km-KH", {
      year: "numeric",
      month: "long",
    }).format(cursor);
  } catch {}

  return (
    <div className="wf-dp-wrap" ref={wrapRef} style={style}>
      <button
        type="button"
        disabled={disabled}
        className={`wf-dp-trigger${open ? " open" : ""}`}
        onClick={() => !disabled && setOpen((o) => !o)}
        style={disabled ? { opacity: 0.6, cursor: "not-allowed" } : undefined}
      >
        <span className={value ? "" : "wf-dp-placeholder"}>
          {value ? fmtDateDisplay(value) : placeholder || "ជ្រើសរើសកាលបរិច្ឆេទ"}
        </span>
        <CalendarDays size={15} style={{ opacity: 0.55, flexShrink: 0 }} />
      </button>
      {open && !disabled && (
        <div className="wf-dp-pop" onMouseDown={(e) => e.stopPropagation()}>
          <div className="wf-dp-head">
            <button
              type="button"
              className="wf-dp-nav"
              onClick={() => setCursor(new Date(year, month - 1, 1))}
            >
              <ChevronLeft size={15} />
            </button>
            <span className="wf-dp-title">{monthTitle}</span>
            <button
              type="button"
              className="wf-dp-nav"
              onClick={() => setCursor(new Date(year, month + 1, 1))}
            >
              <ChevronRight size={15} />
            </button>
          </div>
          <div className="wf-dp-grid">
            {DP_DOW.map((w, i) => (
              <div className="wf-dp-dow" key={i}>
                {w}
              </div>
            ))}
            {cells.map((c, i) => {
              const ymd = fmtYMD(c.date);
              const isSel = value && ymd === value;
              const isToday = ymd === today;
              return (
                <button
                  type="button"
                  key={i}
                  className={`wf-dp-day${c.outside ? " outside" : ""}${
                    isToday ? " today" : ""
                  }${isSel ? " selected" : ""}`}
                  onClick={() => {
                    fire(c.date);
                    setOpen(false);
                  }}
                >
                  {c.day}
                </button>
              );
            })}
          </div>
          <div className="wf-dp-foot">
            <button
              type="button"
              className="wf-dp-link"
              onClick={() => {
                onChange && onChange({ target: { value: "" } });
                setOpen(false);
              }}
            >
              សម្អាត
            </button>
            <button
              type="button"
              className="wf-dp-link"
              onClick={() => {
                fire(new Date());
                setOpen(false);
              }}
            >
              ថ្ងៃនេះ
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function parseHM(s) {
  if (!s) return null;
  const [h, m] = String(s).split(":").map(Number);
  if (Number.isNaN(h) || Number.isNaN(m)) return null;
  return { h, m };
}
function fmtHM(h, m) {
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}
const TP_HOURS = Array.from({ length: 24 }, (_, i) => i);
const TP_MINUTES = Array.from({ length: 60 }, (_, i) => i);

function TimePicker({ value, onChange, placeholder, style, disabled }) {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef(null);
  const hourColRef = useRef(null);
  const minColRef = useRef(null);
  const parsed = parseHM(value);
  useCloseOnOutside(wrapRef, () => setOpen(false));

  useEffect(() => {
    if (!open) return;
    const h = parsed ? parsed.h : 9;
    const m = parsed ? parsed.m : 0;
    requestAnimationFrame(() => {
      hourColRef.current
        ?.querySelector(`[data-v="${h}"]`)
        ?.scrollIntoView({ block: "center" });
      minColRef.current
        ?.querySelector(`[data-v="${m}"]`)
        ?.scrollIntoView({ block: "center" });
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const fire = (h, m) =>
    onChange && onChange({ target: { value: fmtHM(h, m) } });

  return (
    <div className="wf-dp-wrap" ref={wrapRef} style={style}>
      <button
        type="button"
        disabled={disabled}
        className={`wf-dp-trigger${open ? " open" : ""}`}
        onClick={() => !disabled && setOpen((o) => !o)}
        style={disabled ? { opacity: 0.6, cursor: "not-allowed" } : undefined}
      >
        <span className={value ? "" : "wf-dp-placeholder"}>
          {value || placeholder || "ម៉ោង"}
        </span>
        <Clock size={15} style={{ opacity: 0.55, flexShrink: 0 }} />
      </button>
      {open && !disabled && (
        <div className="wf-tp-pop" onMouseDown={(e) => e.stopPropagation()}>
          <div className="wf-tp-cols">
            <div className="wf-tp-col" ref={hourColRef}>
              <div className="wf-tp-pad" />
              {TP_HOURS.map((h) => (
                <div
                  key={h}
                  data-v={h}
                  className={`wf-tp-item${
                    parsed && parsed.h === h ? " selected" : ""
                  }`}
                  onClick={() => fire(h, parsed ? parsed.m : 0)}
                >
                  {String(h).padStart(2, "0")}
                </div>
              ))}
              <div className="wf-tp-pad" />
            </div>
            <div className="wf-tp-sep">:</div>
            <div className="wf-tp-col" ref={minColRef}>
              <div className="wf-tp-pad" />
              {TP_MINUTES.map((m) => (
                <div
                  key={m}
                  data-v={m}
                  className={`wf-tp-item${
                    parsed && parsed.m === m ? " selected" : ""
                  }`}
                  onClick={() => fire(parsed ? parsed.h : 9, m)}
                >
                  {String(m).padStart(2, "0")}
                </div>
              ))}
              <div className="wf-tp-pad" />
            </div>
          </div>
          <div className="wf-dp-foot">
            <button
              type="button"
              className="wf-dp-link"
              onClick={() => {
                onChange && onChange({ target: { value: "" } });
                setOpen(false);
              }}
            >
              សម្អាត
            </button>
            <button
              type="button"
              className="wf-dp-link"
              onClick={() => {
                const now = new Date();
                fire(now.getHours(), now.getMinutes());
                setOpen(false);
              }}
            >
              ឥឡូវ
            </button>
          </div>
        </div>
      )}
    </div>
  );
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
  background: linear-gradient(160deg, #050810 0%, #0A0F1A 55%, #0D1420 100%);
}
.wf-login-bg {
  position:absolute; inset:-40px; z-index:0; pointer-events:none;
  background: radial-gradient(ellipse 80% 60% at 20% 30%, rgba(240,168,59,0.10) 0%, transparent 60%),
              radial-gradient(ellipse 60% 80% at 80% 70%, rgba(91,141,239,0.12) 0%, transparent 60%),
              radial-gradient(ellipse 50% 40% at 60% 10%, rgba(31,162,107,0.08) 0%, transparent 60%);
  animation: wf-bg-drift 14s ease-in-out infinite, wf-bg-hue 22s ease-in-out infinite;
  will-change: transform, filter;
}
.wf-login-orbs { position:absolute; inset:0; z-index:0; pointer-events:none; overflow:hidden; }
.wf-login-orb {
  position:absolute; border-radius:50%; filter:blur(50px); opacity:.4;
  will-change: transform;
}
.wf-login-orb-1 {
  width:260px; height:260px; top:8%; left:8%;
  background:radial-gradient(circle,rgba(31,162,107,0.5),transparent 70%);
  animation: wf-orb-a 16s ease-in-out infinite;
}
.wf-login-orb-2 {
  width:320px; height:320px; bottom:6%; right:6%;
  background:radial-gradient(circle,rgba(91,141,239,0.45),transparent 70%);
  animation: wf-orb-b 20s ease-in-out infinite;
}
.wf-login-orb-3 {
  width:200px; height:200px; top:55%; left:2%;
  background:radial-gradient(circle,rgba(240,168,59,0.4),transparent 70%);
  animation: wf-orb-c 18s ease-in-out infinite;
}
.wf-login-orb-4 {
  width:180px; height:180px; top:4%; right:16%;
  background:radial-gradient(circle,rgba(229,99,122,0.3),transparent 70%);
  animation: wf-orb-d 15s ease-in-out infinite;
}
.wf-login-card {
  position:relative; z-index:2;
  width:100%; max-width:420px; margin:16px;
  padding:36px 32px 28px;
  background:rgba(14,18,27,0.86);
  backdrop-filter:blur(24px);
  border-radius:16px;
  border:1px solid rgba(255,255,255,0.08);
  box-shadow: 0 1px 0 rgba(255,255,255,0.04) inset, 0 32px 80px rgba(0,0,0,0.55);
  animation: wf-float-up .6s cubic-bezier(.16,.9,.28,1) both;
}
.wf-login-logo-ring {
  width:56px; height:56px; border-radius:12px;
  background:${T.gold};
  display:flex; align-items:center; justify-content:center;
  font-weight:800; font-size:18px; color:#1A1300;
  font-family:'JetBrains Mono',monospace;
  box-shadow:inset 0 0 0 1px rgba(255,255,255,0.25);
  position:relative;
}
.wf-login-logo-ring::before {
  content:''; position:absolute; inset:-8px; border-radius:18px;
  border:1px solid rgba(240,168,59,0.3);
  animation: wf-pulse-ring 3s ease-in-out infinite;
}
.wf-login-input {
  width:100%; padding:12px 14px; border-radius:8px;
  border:1.5px solid rgba(255,255,255,0.1); font-size:14px;
  background:rgba(255,255,255,0.04); color:#EEF1F6; outline:none;
  font-family:inherit; transition:border-color .25s ease, box-shadow .25s ease, background .25s ease, transform .15s ease;
  box-sizing:border-box;
}
.wf-login-input::placeholder { color:#5B6478; }
.wf-login-input:focus {
  border-color:${T.gold}; background:rgba(255,255,255,0.06);
  box-shadow:0 0 0 4px rgba(240,168,59,0.14);
  transform:translateY(-1px);
}
.wf-login-btn {
  width:100%; padding:13px; border:none; border-radius:9px;
  font-size:15px; font-weight:700; cursor:pointer; display:flex;
  align-items:center; justify-content:center; gap:8px;
  font-family:inherit; transition:transform .2s cubic-bezier(.2,.9,.3,1), box-shadow .2s ease, filter .2s ease;
  position:relative; overflow:hidden;
}
.wf-login-btn:hover:not(:disabled) { transform:translateY(-1px); }
.wf-login-btn:active:not(:disabled) { transform:scale(.97) translateY(0); }
.wf-login-btn-emp {
  background:${T.gold};
  color:#1A1300;
  box-shadow:0 4px 16px rgba(240,168,59,0.3);
}
.wf-login-btn-emp:hover { box-shadow:0 6px 24px rgba(240,168,59,0.4); filter:brightness(1.05); }
.wf-login-btn-adm {
  background:rgba(255,255,255,0.06);
  color:#fff;
  border:1px solid rgba(255,255,255,0.14);
  box-shadow:none;
}
.wf-login-btn-adm:hover { background:rgba(255,255,255,0.1); }
.wf-login-divider { display:flex; align-items:center; gap:10px; margin:18px 0; }
.wf-login-divider::before,.wf-login-divider::after { content:''; flex:1; height:1px; background:rgba(255,255,255,0.1); }
.wf-login-error {
  display:flex; align-items:center; gap:7px; font-size:12.5px;
  color:#F0879B; background:rgba(229,99,122,0.12); border-radius:8px;
  padding:9px 12px; margin-bottom:14px; border:1px solid rgba(229,99,122,0.25);
}
.wf-login-demo {
  margin-top:20px; padding:11px 14px; background:rgba(255,255,255,0.04);
  border-radius:8px; font-size:11px; color:#8891A6;
  line-height:1.7; border:1px solid rgba(255,255,255,0.07);
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
          color: "#8A93A8",
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
              fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
              fontWeight: 700,
              fontSize: 20,
              color: "#EEF1F6",
            }}
          >
            {displayName}
          </div>
          <div style={{ fontSize: 13, color: "#8891A6", marginTop: 4 }}>
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
            color: "#8891A6",
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
              width: 56,
              height: 56,
              borderRadius: 12,
              background: branding.logo ? "#fff" : T.blue,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              boxShadow: "inset 0 0 0 1px rgba(255,255,255,0.25)",
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
              <ShieldCheck size={26} color="#fff" />
            )}
          </div>
          <div
            style={{
              marginTop: 14,
              fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
              fontWeight: 700,
              fontSize: 20,
              color: "#EEF1F6",
            }}
          >
            {L.adminTitle}
          </div>
          <div style={{ fontSize: 13, color: "#8891A6", marginTop: 4 }}>
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
            fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
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
            fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
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
   Analytics — lightweight, dependency-free SVG charts built from
   data already loaded elsewhere in the app (no new package needed).
----------------------------------------------------------------*/
// Returns the last `count` "YYYY-MM" month keys, oldest first, ending
// at the current month.
function lastMonthKeys(count) {
  const out = [];
  const now = new Date();
  for (let i = count - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    out.push(monthKey(d));
  }
  return out;
}
// Short month label ("Jan", "ម.ក") instead of monthLabel's full
// "January 2026" — compact enough for chart axis ticks.
function shortMonthLabel(mk, lang) {
  const [y, m] = mk.split("-").map(Number);
  try {
    return new Intl.DateTimeFormat(lang === "en" ? "en-US" : "km-KH", {
      month: "short",
    }).format(new Date(y, m - 1, 1));
  } catch {
    return mk;
  }
}
// A minimal vertical bar chart. `data` is [{ label, value }]. Values are
// scaled to the tallest bar in the set; `formatValue` controls the label
// drawn above each bar and `suffix` is appended to it.
function MiniBarChart({ data, color, formatValue, height = 160 }) {
  const max = Math.max(1, ...data.map((d) => d.value));
  const barW = 100 / data.length;
  return (
    <div style={{ display: "flex", alignItems: "flex-end", gap: 4, height }}>
      {data.map((d, i) => {
        const h = Math.max(2, Math.round((d.value / max) * (height - 34)));
        return (
          <div
            key={i}
            style={{
              flex: `0 0 ${barW}%`,
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "flex-end",
              height: "100%",
            }}
          >
            <div
              style={{
                fontSize: 10,
                color: T.muted,
                fontFamily: "'JetBrains Mono',monospace",
                marginBottom: 4,
                whiteSpace: "nowrap",
              }}
            >
              {formatValue ? formatValue(d.value) : d.value}
            </div>
            <div
              style={{
                width: "62%",
                height: h,
                borderRadius: "5px 5px 2px 2px",
                background: color,
                transition: "height .25s ease",
              }}
              title={`${d.label}: ${formatValue ? formatValue(d.value) : d.value}`}
            />
            <div
              style={{
                fontSize: 10,
                color: T.muted,
                marginTop: 6,
                textAlign: "center",
              }}
            >
              {d.label}
            </div>
          </div>
        );
      })}
    </div>
  );
}
// A minimal horizontal bar chart, better suited to longer text labels
// (department names) than vertical bars.
function HorizontalBarChart({ data, color, formatValue }) {
  const max = Math.max(1, ...data.map((d) => d.value));
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      {data.map((d, i) => (
        <div key={i}>
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              fontSize: 12,
              marginBottom: 4,
            }}
          >
            <span style={{ color: T.ink, fontWeight: 500 }}>{d.label}</span>
            <span
              style={{
                color: T.textSoft,
                fontFamily: "'JetBrains Mono',monospace",
              }}
            >
              {formatValue ? formatValue(d.value) : d.value}
            </span>
          </div>
          <div
            style={{
              height: 8,
              borderRadius: 5,
              background: T.divider,
              overflow: "hidden",
            }}
          >
            <div
              style={{
                width: `${Math.max(2, Math.round((d.value / max) * 100))}%`,
                height: "100%",
                borderRadius: 5,
                background: color,
              }}
            />
          </div>
        </div>
      ))}
    </div>
  );
}
function ChartCard({ title, subtitle, children, noData, noDataLabel }) {
  return (
    <Card style={{ padding: 18 }}>
      <h3
        style={{
          fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
          fontWeight: 600,
          color: T.ink,
          fontSize: 14,
        }}
      >
        {title}
      </h3>
      {subtitle && (
        <p
          style={{
            fontSize: 11,
            color: T.muted,
            marginTop: 2,
            marginBottom: 14,
          }}
        >
          {subtitle}
        </p>
      )}
      {noData ? (
        <p
          style={{
            fontSize: 13,
            color: T.muted,
            textAlign: "center",
            padding: "28px 0",
          }}
        >
          {noDataLabel}
        </p>
      ) : (
        <div style={{ marginTop: subtitle ? 0 : 14 }}>{children}</div>
      )}
    </Card>
  );
}
function AnalyticsPage({
  employees,
  departments,
  attendance,
  overtimeRequests,
  otPolicy,
  payrollPolicy,
}) {
  const { t, lang } = useLang();
  const months = useMemo(() => lastMonthKeys(6), []);
  const activeEmployees = employees.filter((e) => e.status === "active");

  // Attendance rate per month: share of that month's attendance records
  // that were present/late (as opposed to absent), out of all records
  // logged in that month.
  const attendTrend = months.map((mk) => {
    const recs = attendance.filter((a) => a.date && a.date.startsWith(mk));
    const present = recs.filter(
      (a) => a.status === "present" || a.status === "late",
    ).length;
    const rate = recs.length ? Math.round((present / recs.length) * 100) : 0;
    return { label: shortMonthLabel(mk, lang), value: rate };
  });
  const hasAttendData = attendance.some((a) => a.date);

  // Net payroll cost by department for the current month.
  const mk = monthKey();
  const deptCost = departments
    .map((d) => {
      const total = activeEmployees
        .filter((e) => e.deptId === d.id)
        .reduce((sum, e) => {
          const p = computePayroll(
            e,
            attendance,
            mk,
            overtimeRequests,
            otPolicy,
            payrollPolicy,
          );
          return sum + p.net;
        }, 0);
      return { label: d.name, value: Math.round(total) };
    })
    .filter((d) => d.value > 0)
    .sort((a, b) => b.value - a.value);

  // Approved OT hours per month across all employees.
  const otTrend = months.map((mkm) => {
    const hours = overtimeRequests
      .filter(
        (r) => r.status === "approved" && r.date && r.date.startsWith(mkm),
      )
      .reduce((sum, r) => sum + (Number(r.hours) || 0), 0);
    return { label: shortMonthLabel(mkm, lang), value: Math.round(hours) };
  });
  const hasOtData = overtimeRequests.some((r) => r.status === "approved");
  const totalOtThisTrend = otTrend.reduce((s, d) => s + d.value, 0);

  return (
    <div>
      <Card
        style={{
          padding: 20,
          marginBottom: 22,
          background: BRAND.ink,
          color: "#fff",
        }}
      >
        <h2
          style={{
            fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
            fontSize: 20,
            fontWeight: 600,
            display: "flex",
            alignItems: "center",
            gap: 8,
          }}
        >
          <BarChart3 size={20} /> {t.analytics.title}
        </h2>
        <p style={{ color: "#A9B4C7", fontSize: 12, marginTop: 6 }}>
          {t.analytics.subtitle}
        </p>
      </Card>

      <div
        className="wf-grid"
        style={{
          gridTemplateColumns: "repeat(auto-fit, minmax(280px,1fr))",
          marginBottom: 16,
        }}
      >
        <ChartCard
          title={t.analytics.attendTrend}
          subtitle={t.analytics.attendTrendSub}
          noData={!hasAttendData}
          noDataLabel={t.analytics.noChartData}
        >
          <MiniBarChart
            data={attendTrend}
            color={T.forest}
            formatValue={(v) => `${v}%`}
          />
        </ChartCard>

        <ChartCard
          title={t.analytics.otTrend}
          subtitle={`${t.analytics.otTrendSub} · ${t.analytics.totalOt}: ${totalOtThisTrend} ${t.analytics.hours}`}
          noData={!hasOtData}
          noDataLabel={t.analytics.noChartData}
        >
          <MiniBarChart
            data={otTrend}
            color={T.gold}
            formatValue={(v) => `${v}h`}
          />
        </ChartCard>
      </div>

      <ChartCard
        title={t.analytics.deptCost}
        subtitle={t.analytics.deptCostSub}
        noData={deptCost.length === 0}
        noDataLabel={t.analytics.noChartData}
      >
        <HorizontalBarChart
          data={deptCost}
          color={T.blue}
          formatValue={fmtMoney}
        />
      </ChartCard>
    </div>
  );
}

/* ---------------------------------------------------------------
   Employees
----------------------------------------------------------------*/
function EmployeeForm({
  initial,
  departments,
  shifts,
  offices,
  onSave,
  onCancel,
}) {
  const { t, lang } = useLang();
  const [f, setF] = useState(
    initial || {
      code: "",
      pin: randomPin(),
      name: "",
      deptId: departments[0]?.id || "",
      shiftId: shifts[0]?.id || "",
      officeId: "",
      weeklyOff: [],
      customDaysOff: [],
      annualLeaveDays: DEFAULT_ANNUAL_LEAVE_DAYS,
      sickLeaveDays: DEFAULT_SICK_LEAVE_DAYS,
      role: "",
      email: "",
      phone: "",
      salary: "",
      status: "active",
      joined: todayStr(),
      useCustomRate: false,
      customTaxRate: "",
      customInsuranceRate: "",
    },
  );
  const [newOffDate, setNewOffDate] = useState("");
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const toggleWeeklyOff = (dow) => {
    const cur = f.weeklyOff || [];
    setF({
      ...f,
      weeklyOff: cur.includes(dow)
        ? cur.filter((d) => d !== dow)
        : [...cur, dow].sort(),
    });
  };
  const addCustomDayOff = () => {
    if (!newOffDate) return;
    const cur = f.customDaysOff || [];
    if (cur.includes(newOffDate)) return;
    setF({ ...f, customDaysOff: [...cur, newOffDate].sort() });
    setNewOffDate("");
  };
  const removeCustomDayOff = (dateStr) => {
    setF({
      ...f,
      customDaysOff: (f.customDaysOff || []).filter((d) => d !== dateStr),
    });
  };
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
      <div
        style={{
          border: `1px solid ${T.lineSoft}`,
          borderRadius: 10,
          padding: "10px 12px",
          marginBottom: 14,
        }}
      >
        <label
          style={{
            display: "flex",
            alignItems: "center",
            gap: 8,
            cursor: "pointer",
            fontSize: 13,
            fontWeight: 600,
            color: T.ink,
          }}
        >
          <input
            type="checkbox"
            checked={!!f.useCustomRate}
            onChange={(e) => setF({ ...f, useCustomRate: e.target.checked })}
            style={{ width: 15, height: 15, accentColor: T.forest }}
          />
          {t.pay.customRateToggle}
        </label>
        <p
          style={{
            fontSize: 11.5,
            color: T.muted,
            marginTop: 6,
            marginBottom: f.useCustomRate ? 12 : 0,
          }}
        >
          {t.pay.customRateHint}
        </p>
        {f.useCustomRate && (
          <div
            style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}
          >
            <Field label={t.pay.customTaxRateLabel}>
              <Input
                type="number"
                step="0.1"
                min="0"
                value={f.customTaxRate}
                onChange={set("customTaxRate")}
                placeholder="0"
              />
            </Field>
            <Field label={t.pay.customInsuranceRateLabel}>
              <Input
                type="number"
                step="0.1"
                min="0"
                value={f.customInsuranceRate}
                onChange={set("customInsuranceRate")}
                placeholder="0"
              />
            </Field>
          </div>
        )}
      </div>
      <Field label={t.emps.joined}>
        <DatePicker value={f.joined || todayStr()} onChange={set("joined")} />
      </Field>
      <Field label="សាខា (Branch)">
        <Select value={f.officeId || ""} onChange={set("officeId")}>
          <option value="">មិនទាន់កំណត់</option>
          {(offices || []).map((o) => (
            <option key={o.id} value={o.id}>
              {o.name}
            </option>
          ))}
        </Select>
      </Field>
      <Field label="ថ្ងៃច្បាប់ប្រចាំឆ្នាំ (ថ្ងៃ/ឆ្នាំ)">
        <Input
          type="number"
          min={0}
          value={f.annualLeaveDays ?? DEFAULT_ANNUAL_LEAVE_DAYS}
          onChange={set("annualLeaveDays")}
          placeholder={String(DEFAULT_ANNUAL_LEAVE_DAYS)}
        />
      </Field>
      <Field label="ថ្ងៃច្បាប់ឈឺ (ថ្ងៃ/ឆ្នាំ)">
        <Input
          type="number"
          min={0}
          value={f.sickLeaveDays ?? DEFAULT_SICK_LEAVE_DAYS}
          onChange={set("sickLeaveDays")}
          placeholder={String(DEFAULT_SICK_LEAVE_DAYS)}
        />
      </Field>
      <Field label="ថ្ងៃឈប់សម្រាកប្រចាំសប្តាហ៍">
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {WEEKDAY_LABELS.map((label, dow) => {
            const active = (f.weeklyOff || []).includes(dow);
            return (
              <button
                type="button"
                key={dow}
                onClick={() => toggleWeeklyOff(dow)}
                style={{
                  width: 38,
                  height: 38,
                  borderRadius: 10,
                  border: `1px solid ${active ? T.forest : T.line}`,
                  background: active ? T.forest : T.inputBg,
                  color: active ? "#fff" : T.text,
                  fontWeight: 700,
                  fontSize: 12.5,
                  cursor: "pointer",
                  transition: "background .12s ease,border-color .12s ease",
                }}
              >
                {label}
              </button>
            );
          })}
        </div>
      </Field>
      <Field label="ថ្ងៃឈប់សម្រាកពិសេស (កាលបរិច្ឆេទជាក់លាក់)">
        <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
          <DatePicker
            style={{ flex: 1 }}
            value={newOffDate}
            onChange={(e) => setNewOffDate(e.target.value)}
          />
          <Button type="button" variant="ghost" onClick={addCustomDayOff}>
            <Plus size={14} /> បន្ថែម
          </Button>
        </div>
        {(f.customDaysOff || []).length > 0 && (
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {(f.customDaysOff || []).map((d) => (
              <span
                key={d}
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 5,
                  fontSize: 12,
                  fontWeight: 600,
                  color: T.forestText,
                  background: T.forestSoft,
                  padding: "5px 8px 5px 10px",
                  borderRadius: 999,
                }}
              >
                {fmtDateDisplay(d)}
                <button
                  type="button"
                  onClick={() => removeCustomDayOff(d)}
                  style={{
                    background: "none",
                    border: "none",
                    cursor: "pointer",
                    color: "inherit",
                    display: "flex",
                    padding: 0,
                  }}
                >
                  <X size={12} />
                </button>
              </span>
            ))}
          </div>
        )}
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
  offices,
  setEmployees,
  isSuperAdmin,
  currentAdmin,
  documents,
  setDocuments,
}) {
  const { t, lang } = useLang();
  const { branding } = useBranding();
  const [modal, setModal] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const [docsFor, setDocsFor] = useState(null);
  const [query, setQuery] = useState("");
  const [branchFilter, setBranchFilter] = useState("");
  const filtered = useMemo(
    () =>
      employees.filter(
        (e) =>
          (e.name + e.code + e.role)
            .toLowerCase()
            .includes(query.toLowerCase()) &&
          (!branchFilter || e.officeId === branchFilter),
      ),
    [employees, query, branchFilter],
  );
  const pg = usePagination(filtered, 24);
  const deptName = (id) => departments.find((d) => d.id === id)?.name || "—";
  const shiftOf = (id) => shifts.find((s) => s.id === id);
  const officeName = (id) => offices.find((o) => o.id === id)?.name || null;

  const save = (data) => {
    const clean = {
      ...data,
      salary: Number(data.salary) || 0,
      annualLeaveDays: Number.isFinite(Number(data.annualLeaveDays))
        ? Number(data.annualLeaveDays)
        : DEFAULT_ANNUAL_LEAVE_DAYS,
      sickLeaveDays: Number.isFinite(Number(data.sickLeaveDays))
        ? Number(data.sickLeaveDays)
        : DEFAULT_SICK_LEAVE_DAYS,
      useCustomRate: !!data.useCustomRate,
      customTaxRate: data.useCustomRate
        ? Number(data.customTaxRate) || 0
        : null,
      customInsuranceRate: data.useCustomRate
        ? Number(data.customInsuranceRate) || 0
        : null,
    };
    if (modal.mode === "add")
      setEmployees([...employees, { ...clean, id: uid("e") }]);
    else setEmployees(employees.map((e) => (e.id === clean.id ? clean : e)));
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
        {offices && offices.length > 0 && (
          <Select
            style={{ maxWidth: 180 }}
            value={branchFilter}
            onChange={(e) => setBranchFilter(e.target.value)}
          >
            <option value="">គ្រប់សាខា</option>
            {offices.map((o) => (
              <option key={o.id} value={o.id}>
                {o.name}
              </option>
            ))}
          </Select>
        )}
        <Button
          variant="ghost"
          onClick={() =>
            exportCsv(
              `employees-${todayStr()}.csv`,
              [
                t.emps.name,
                t.emps.code,
                t.emps.dept,
                t.emps.role,
                t.emps.shift,
                t.emps.phone,
                t.emps.email,
                t.emps.salary,
                t.emps.joined,
                t.status,
              ],
              filtered.map((e) => [
                e.name,
                e.code,
                deptName(e.deptId),
                e.role,
                shiftLabel(shiftOf(e.shiftId)),
                e.phone,
                e.email,
                e.salary,
                e.joined,
                e.status === "active" ? t.emps.active : t.emps.inactive,
              ]),
            )
          }
        >
          <Download size={15} /> {t.exportCsv}
        </Button>
        <Button variant="accent" onClick={() => setModal({ mode: "add" })}>
          <Plus size={15} /> {t.emps.addBtn}
        </Button>
      </div>

      <div
        className="wf-grid"
        style={{ gridTemplateColumns: "repeat(auto-fill, minmax(260px,1fr))" }}
      >
        {pg.pageItems.map((e) => (
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
              {officeName(e.officeId) && <div>🏬 {officeName(e.officeId)}</div>}
              {((e.weeklyOff && e.weeklyOff.length > 0) ||
                (e.customDaysOff && e.customDaysOff.length > 0)) && (
                <div style={{ color: T.rose }}>
                  🛌{" "}
                  {(e.weeklyOff || []).map((d) => WEEKDAY_LABELS[d]).join(" ")}
                  {e.weeklyOff?.length && e.customDaysOff?.length ? " · " : ""}
                  {e.customDaysOff?.length
                    ? `${e.customDaysOff.length} ថ្ងៃពិសេស`
                    : ""}
                </div>
              )}
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
            <Button
              size="sm"
              variant="ghost"
              style={{ width: "100%", justifyContent: "center", marginTop: 8 }}
              onClick={() =>
                printEmployeeBadge({
                  t,
                  brandName: branding?.name?.trim() || t.appName,
                  brandLogo: branding?.logo || null,
                  emp: e,
                  deptLabel: deptName(e.deptId),
                  roleLabel: e.role,
                  shiftText: shiftLabel(shiftOf(e.shiftId)),
                  officeText: officeName(e.officeId),
                  statusLabel:
                    e.status === "active" ? t.emps.active : t.emps.inactive,
                })
              }
            >
              <BadgeCheck size={13} /> {t.emps.printBadge}
            </Button>
            <Button
              size="sm"
              variant="ghost"
              style={{ width: "100%", justifyContent: "center", marginTop: 8 }}
              onClick={() => setDocsFor(e)}
            >
              <FileText size={13} /> {t.doc.title}
            </Button>
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
      <Pagination {...pg} />

      {modal && (
        <Modal
          title={modal.mode === "add" ? t.emps.addTitle : t.emps.editTitle}
          onClose={() => setModal(null)}
        >
          <EmployeeForm
            initial={modal.data}
            departments={departments}
            shifts={shifts}
            offices={offices}
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
      {docsFor && (
        <EmployeeDocumentsModal
          emp={docsFor}
          currentAdmin={currentAdmin}
          documents={documents}
          setDocuments={setDocuments}
          isSuperAdmin={isSuperAdmin}
          onClose={() => setDocsFor(null)}
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
          <TimePicker value={f.start} onChange={set("start")} />
        </Field>
        <Field label="ម៉ោងចេញ">
          <TimePicker value={f.end} onChange={set("end")} />
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
  const [branchWarning, setBranchWarning] = useState("");
  const hasOffices = offices && offices.length > 0;
  const todayIsDayOff = isDayOff(emp, today);

  // If one or more office branches (each with lat/lng + radius) are
  // configured, require the employee's current GPS position to fall
  // within at least one of them before allowing a punch. Returns
  // { lat, lng, distance, officeId, officeName } for the closest matching
  // branch to attach to the attendance record, or null if the punch
  // should be blocked (locError is set in that case).
  const verifyLocation = async () => {
    if (!hasOffices) return null; // no geofence configured — skip check
    setLocError("");
    setBranchWarning("");
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
      if (emp.officeId && match.office.id !== emp.officeId) {
        setBranchWarning(
          `⚠️ អ្នកកំពុងចុះឈ្មោះនៅសាខា "${match.office.name}" ប៉ុន្តែអ្នកត្រូវបានកំណត់ឲ្យធ្វើការនៅសាខាផ្សេង`,
        );
      } else {
        setBranchWarning("");
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
    // Re-fetch the employee's *current* shift assignment right before
    // computing lateness. `shift` here is whatever was loaded into this
    // browser tab when the page opened — if an admin reassigned the
    // employee to a different shift after that (e.g. moments before this
    // punch), the in-memory value would be stale and give the wrong
    // late/present result. A fresh read avoids that.
    let effectiveShift = shift;
    try {
      const { data: freshEmp } = await supabase
        .from("employees")
        .select("shift_id")
        .eq("id", emp.id)
        .maybeSingle();
      if (freshEmp && freshEmp.shift_id !== emp.shiftId) {
        const { data: freshShift } = await supabase
          .from("shifts")
          .select("*")
          .eq("id", freshEmp.shift_id)
          .maybeSingle();
        if (freshShift) effectiveShift = freshShift;
      }
    } catch {
      // If the re-fetch fails (offline, etc.), fall back to the shift
      // already in memory rather than blocking the punch.
    }
    const status = isLateForShift(t, effectiveShift) ? "late" : "present";
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
      {todayIsDayOff && (
        <p
          style={{
            fontSize: 12.5,
            fontWeight: 600,
            color: T.goldText,
            background: T.goldSoft,
            padding: "8px 12px",
            borderRadius: 10,
            marginBottom: 12,
            maxWidth: 320,
            marginLeft: "auto",
            marginRight: "auto",
          }}
        >
          🛌 ថ្ងៃនេះជាថ្ងៃឈប់សម្រាករបស់អ្នក —
          អ្នកនៅតែអាចចុះឈ្មោះបានប្រសិនបើអ្នកមកធ្វើការ
        </p>
      )}
      {branchWarning && (
        <p
          style={{
            fontSize: 12.5,
            fontWeight: 600,
            color: T.goldText,
            background: T.goldSoft,
            padding: "8px 12px",
            borderRadius: 10,
            marginBottom: 12,
            maxWidth: 320,
            marginLeft: "auto",
            marginRight: "auto",
          }}
        >
          {branchWarning}
        </p>
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

function ManualAttendanceForm({ employees, initial, onSave, onCancel }) {
  const { t, lang } = useLang();
  const editing = !!initial;
  const [f, setF] = useState(
    initial
      ? {
          employeeId: initial.employeeId,
          date: initial.date,
          checkIn: initial.checkIn || "08:00",
          checkOut: initial.checkOut || "",
          status: initial.status,
        }
      : {
          employeeId: employees[0]?.id || "",
          date: todayStr(),
          checkIn: "08:00",
          checkOut: "",
          status: "present",
        },
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const emp = employees.find((e) => e.id === f.employeeId);
  return (
    <div>
      <Field label="បុគ្គលិក">
        {editing ? (
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 10,
              padding: "9px 12px",
              background: T.paper,
              borderRadius: 10,
            }}
          >
            <Avatar name={emp?.name} photo={emp?.photo} size={26} />
            <div>
              <div style={{ fontWeight: 600, fontSize: 13, color: T.ink }}>
                {emp?.name}
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
        ) : (
          <Select value={f.employeeId} onChange={set("employeeId")}>
            {employees.map((e) => (
              <option key={e.id} value={e.id}>
                {e.name} ({e.code})
              </option>
            ))}
          </Select>
        )}
      </Field>
      <Field label="កាលបរិច្ឆេទ">
        <DatePicker value={f.date} onChange={set("date")} disabled={editing} />
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
            <TimePicker value={f.checkIn} onChange={set("checkIn")} />
          </Field>
          <Field label="ម៉ោងចេញ">
            <TimePicker value={f.checkOut} onChange={set("checkOut")} />
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

/* ---------------------------------------------------------------
   Public holidays — admin maintains a company-wide holiday calendar.
   The dates feed into OT day-type auto-suggestion and show as an
   informational banner on the Attendance page.
----------------------------------------------------------------*/
function isHoliday(dateStr, holidays) {
  return (holidays || []).find((h) => h.date === dateStr) || null;
}

function HolidayForm({ initial, onSave, onCancel }) {
  const { t } = useLang();
  const [f, setF] = useState(initial || { date: todayStr(), name: "" });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const invalid = !f.date || !f.name.trim();
  return (
    <div>
      <Field label={t.hol.dateLabel}>
        <DatePicker value={f.date} onChange={set("date")} />
      </Field>
      <Field label={t.hol.nameLabel}>
        <Input
          value={f.name}
          onChange={set("name")}
          placeholder={t.hol.namePlaceholder}
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
          onClick={() => onSave({ date: f.date, name: f.name.trim() })}
          disabled={invalid}
        >
          {t.save}
        </Button>
      </div>
    </div>
  );
}

function Holidays({ holidays, setHolidays, isSuperAdmin }) {
  const { t } = useLang();
  const [modal, setModal] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const today = todayStr();

  const save = (f) => {
    if (modal.mode === "add") {
      setHolidays([...holidays, { id: uid("hol"), ...f }]);
    } else {
      setHolidays(
        holidays.map((h) => (h.id === modal.data.id ? { ...h, ...f } : h)),
      );
    }
    setModal(null);
  };

  const sorted = [...holidays].sort((a, b) => a.date.localeCompare(b.date));
  const upcoming = sorted.filter((h) => h.date >= today);
  const past = sorted.filter((h) => h.date < today).reverse();

  const renderRow = (h) => (
    <div
      key={h.id}
      style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        gap: 10,
        padding: "10px 12px",
        border: `1px solid ${T.lineSoft}`,
        borderRadius: 10,
      }}
    >
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <CalendarDays size={16} color={T.forestText} />
        <div>
          <div style={{ fontSize: 13, fontWeight: 600, color: T.ink }}>
            {h.name}
          </div>
          <div
            style={{
              fontSize: 11,
              color: T.muted,
              fontFamily: "'JetBrains Mono',monospace",
            }}
          >
            {h.date}
          </div>
        </div>
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <button
          onClick={() => setModal({ mode: "edit", data: h })}
          style={{
            background: "none",
            border: "none",
            cursor: "pointer",
            color: T.mutedLight,
          }}
        >
          <Pencil size={14} />
        </button>
        {isSuperAdmin && (
          <button
            onClick={() => setConfirmDel(h)}
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
      </div>
    </div>
  );

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
          <Plus size={15} /> {t.hol.addBtn}
        </Button>
      </div>
      {sorted.length === 0 && (
        <Card
          style={{
            textAlign: "center",
            padding: "32px 0",
            color: T.muted,
            fontSize: 13,
          }}
        >
          {t.hol.noHolidays}
        </Card>
      )}
      {upcoming.length > 0 && (
        <div style={{ marginBottom: 20 }}>
          <div
            style={{
              fontSize: 11,
              fontWeight: 700,
              color: T.muted,
              textTransform: "uppercase",
              letterSpacing: ".03em",
              marginBottom: 8,
            }}
          >
            {t.hol.upcoming}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {upcoming.map(renderRow)}
          </div>
        </div>
      )}
      {past.length > 0 && (
        <div>
          <div
            style={{
              fontSize: 11,
              fontWeight: 700,
              color: T.muted,
              textTransform: "uppercase",
              letterSpacing: ".03em",
              marginBottom: 8,
            }}
          >
            {t.hol.past}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {past.map(renderRow)}
          </div>
        </div>
      )}
      {modal && (
        <Modal
          title={modal.mode === "add" ? t.hol.addTitle : t.hol.editTitle}
          onClose={() => setModal(null)}
        >
          <HolidayForm
            initial={modal.data}
            onSave={save}
            onCancel={() => setModal(null)}
          />
        </Modal>
      )}
      {confirmDel && (
        <ConfirmDialog
          text={t.hol.confirmDel}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setHolidays(holidays.filter((h) => h.id !== confirmDel.id));
            setConfirmDel(null);
          }}
        />
      )}
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
  holidays,
}) {
  const { t, lang } = useLang();
  const [date, setDate] = useState(todayStr());
  const [modal, setModal] = useState(false);
  const [editRecord, setEditRecord] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const dayRecords = attendance.filter((a) => a.date === date);
  const activeEmployees = employees.filter((e) => e.status === "active");
  const shiftOf = (id) => shifts.find((s) => s.id === id);
  // Hooks must run unconditionally (not just in the employee-view branch
  // below), so this is computed here even though only that branch uses it.
  const myHistory = useMemo(
    () =>
      currentEmp
        ? attendance
            .filter((a) => a.employeeId === currentEmp.id)
            .sort((a, b) => b.date.localeCompare(a.date))
        : [],
    [attendance, currentEmp],
  );
  const myHistoryPg = usePagination(myHistory, 15);
  const rows = activeEmployees.map((e) => ({
    emp: e,
    rec: dayRecords.find((a) => a.employeeId === e.id),
  }));
  const holidayToday = isHoliday(date, holidays);

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
    setEditRecord(null);
  };

  if (role !== "admin" && currentEmp) {
    const pg = myHistoryPg;
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
              fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
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
                {pg.pageItems.map((a) => (
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
          <Pagination {...pg} />
        </Card>
      </div>
    );
  }

  return (
    <div>
      <OfficeLocationSettings offices={offices} setOffices={setOffices} />
      {holidayToday && (
        <Card
          accent={T.gold}
          style={{
            padding: "12px 16px",
            marginBottom: 16,
            display: "flex",
            alignItems: "center",
            gap: 10,
          }}
        >
          <CalendarDays size={18} color={T.goldText} />
          <span style={{ fontSize: 13, fontWeight: 600, color: T.ink }}>
            {t.hol.todayBanner}: {holidayToday.name}
          </span>
        </Card>
      )}
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
          <DatePicker
            style={{ width: 168 }}
            value={date}
            onChange={(e) => setDate(e.target.value)}
          />
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <Button
            variant="ghost"
            onClick={() =>
              exportCsv(
                `attendance-${date}.csv`,
                [t.employee, "Code", "Check In", "Check Out", t.status],
                rows.map(({ emp, rec }) => [
                  emp.name,
                  emp.code,
                  rec?.checkIn || "",
                  rec?.checkOut || "",
                  rec?.status || "absent",
                ]),
              )
            }
          >
            <Download size={15} /> {t.exportCsv}
          </Button>
          <Button
            variant="accent"
            onClick={() => {
              setEditRecord(null);
              setModal(true);
            }}
          >
            <Plus size={15} /> កត់ត្រាដោយដៃ
          </Button>
        </div>
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
                  <div
                    style={{
                      display: "flex",
                      gap: 10,
                      justifyContent: "flex-end",
                    }}
                  >
                    <button
                      onClick={() => {
                        setEditRecord(
                          rec || {
                            employeeId: emp.id,
                            date,
                            checkIn: "08:00",
                            checkOut: "",
                            status: "present",
                          },
                        );
                        setModal(true);
                      }}
                      style={{
                        background: "none",
                        border: "none",
                        cursor: "pointer",
                        color: T.mutedLight,
                      }}
                    >
                      <Pencil size={14} />
                    </button>
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
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>
      {modal && (
        <Modal
          title={editRecord?.id ? "កែសម្រួលវត្តមាន" : "កត់ត្រាវត្តមានដោយដៃ"}
          onClose={() => {
            setModal(false);
            setEditRecord(null);
          }}
        >
          <ManualAttendanceForm
            employees={activeEmployees}
            initial={editRecord}
            onSave={save}
            onCancel={() => {
              setModal(false);
              setEditRecord(null);
            }}
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
function LeaveRequestForm({ onSave, onCancel, remaining }) {
  const { t, lang } = useLang();
  const [f, setF] = useState({
    type: "annual",
    startDate: todayStr(),
    endDate: todayStr(),
    reason: "",
  });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const invalidRange = f.endDate < f.startDate;
  const requestedDays = invalidRange
    ? 0
    : dateRange(f.startDate, f.endDate).length;
  // remaining can be a plain number (legacy annual-only balance) or an
  // { annual, sick } object — normalize so both shapes keep working.
  const remainingForType =
    f.type === "annual"
      ? typeof remaining === "number"
        ? remaining
        : remaining?.annual
      : f.type === "sick"
        ? remaining?.sick
        : undefined;
  const overQuota =
    (f.type === "annual" || f.type === "sick") &&
    typeof remainingForType === "number" &&
    requestedDays > remainingForType;
  return (
    <div>
      <Field label="ប្រភេទច្បាប់">
        <Select value={f.type} onChange={set("type")}>
          <option value="annual">ច្បាប់ប្រចាំឆ្នាំ</option>
          <option value="sick">ច្បាប់ឈឺ</option>
          <option value="other">ផ្សេងៗ</option>
        </Select>
      </Field>
      {f.type === "annual" && typeof remainingForType === "number" && (
        <p
          style={{
            fontSize: 12,
            color: T.muted,
            marginTop: -8,
            marginBottom: 12,
          }}
        >
          នៅសល់ {remainingForType} ថ្ងៃច្បាប់ប្រចាំឆ្នាំ
        </p>
      )}
      {f.type === "sick" && typeof remainingForType === "number" && (
        <p
          style={{
            fontSize: 12,
            color: T.muted,
            marginTop: -8,
            marginBottom: 12,
          }}
        >
          នៅសល់ {remainingForType} ថ្ងៃច្បាប់ឈឺ
        </p>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label="ចាប់ពីថ្ងៃ">
          <DatePicker value={f.startDate} onChange={set("startDate")} />
        </Field>
        <Field label="ដល់ថ្ងៃ">
          <DatePicker value={f.endDate} onChange={set("endDate")} />
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
      {overQuota && (
        <p
          style={{
            fontSize: 12.5,
            fontWeight: 600,
            color: T.goldText,
            background: T.goldSoft,
            padding: "8px 12px",
            borderRadius: 10,
            marginBottom: 12,
          }}
        >
          ⚠️ សំណើនេះ ({requestedDays} ថ្ងៃ) លើសពីសមតុល្យ
          {f.type === "annual" ? "ច្បាប់ប្រចាំឆ្នាំ" : "ច្បាប់ឈឺ"}ដែលនៅសល់ (
          {remainingForType} ថ្ងៃ) — អ្នកនៅតែអាចដាក់ស្នើបាន
          តែសូមរង់ចាំការសម្រេចពី admin
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

// Small inline "decided by" line shown on both the admin and employee
// views, so everyone sees the same approve/reject attribution.
function LeaveDecisionNote({ r, admins }) {
  const { t, lang } = useLang();
  if (r.status !== "approved" && r.status !== "rejected") return null;
  const decider = admins.find((a) => a.id === r.decidedById);
  const name = r.decidedByName || decider?.name || "—";
  const roleLabel = adminRoleLabel(r.decidedByRole, lang);
  return (
    <div style={{ fontSize: 11.5, color: T.textSoft, marginTop: 3 }}>
      {r.status === "approved" ? (
        <span>
          {t.lv.approvedBy} <strong>{name}</strong>
          {roleLabel ? ` · ${roleLabel}` : ""}
        </span>
      ) : (
        <span style={{ color: T.rose }}>
          {t.lv.rejectedBy} <strong>{name}</strong>
          {roleLabel ? ` · ${roleLabel}` : ""}
          {r.decisionReason ? ` — ${r.decisionReason}` : ""}
        </span>
      )}
    </div>
  );
}

function LeaveRejectModal({ onCancel, onConfirm }) {
  const { t } = useLang();
  const [reason, setReason] = useState("");
  const trimmed = reason.trim();
  return (
    <Modal title={t.lv.rejectTitle} onClose={onCancel} width={420}>
      <Field label={t.lv.rejectReason}>
        <textarea
          className="wf-input"
          rows={3}
          style={{ resize: "vertical", fontFamily: "inherit" }}
          value={reason}
          onChange={(e) => setReason(e.target.value)}
          placeholder={t.lv.rejectReasonPlaceholder}
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
          {t.lv.rejectReasonRequired}
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
          {t.lv.reject}
        </Button>
      </div>
    </Modal>
  );
}

function LeaveRequests({
  role,
  currentAdmin,
  currentEmp,
  employees,
  admins,
  leaveRequests,
  setLeaveRequests,
  attendance,
  setAttendance,
  isSuperAdmin,
  canApprove,
}) {
  const { t, lang } = useLang();
  const [modal, setModal] = useState(false);
  const [rejectFor, setRejectFor] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const empOf = (id) => employees.find((e) => e.id === id);

  // Approving a request marks every day in its range as paid leave in
  // attendance, so payroll (which already treats "leave" as paid) picks
  // it up automatically — no separate payroll logic needed.
  // Days the employee has already actually checked into (e.g. a same-day
  // leave request approved after they clocked in that morning, or a
  // leave range that overlaps a day they already worked) are left alone
  // — approving leave elsewhere must never erase real attendance.
  const applyLeaveToAttendance = (req) => {
    const days = dateRange(req.startDate, req.endDate);
    let next = attendance;
    for (const d of days) {
      const existing = next.find(
        (a) => a.employeeId === req.employeeId && a.date === d,
      );
      if (existing && existing.checkIn) continue;
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

  const approve = (req) => {
    setLeaveRequests(
      leaveRequests.map((r) =>
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
    applyLeaveToAttendance(req);
  };
  const reject = (req, reason) => {
    setLeaveRequests(
      leaveRequests.map((r) =>
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
    const bal = annualLeaveBalance(currentEmp, leaveRequests);
    const sickBal = sickLeaveBalance(currentEmp, leaveRequests);
    return (
      <div>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            gap: 12,
            marginBottom: 16,
          }}
        >
          <Card
            accent={bal.remaining <= 0 ? T.rose : T.forest}
            style={{
              padding: "14px 18px",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              flexWrap: "wrap",
              gap: 8,
            }}
          >
            <div>
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  textTransform: "uppercase",
                  letterSpacing: ".04em",
                  color: T.muted,
                  marginBottom: 2,
                }}
              >
                ថ្ងៃច្បាប់ប្រចាំឆ្នាំ {new Date().getFullYear()}
              </div>
              <div style={{ fontSize: 13, color: T.textSoft }}>
                បានប្រើ {bal.used} ក្នុងចំណោម {bal.quota} ថ្ងៃ
              </div>
            </div>
            <div
              style={{
                fontSize: 22,
                fontWeight: 700,
                fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
                color: bal.remaining <= 0 ? T.rose : T.forestText,
              }}
            >
              នៅសល់ {bal.remaining} ថ្ងៃ
            </div>
          </Card>
          <Card
            accent={sickBal.remaining <= 0 ? T.rose : T.blue}
            style={{
              padding: "14px 18px",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              flexWrap: "wrap",
              gap: 8,
            }}
          >
            <div>
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  textTransform: "uppercase",
                  letterSpacing: ".04em",
                  color: T.muted,
                  marginBottom: 2,
                }}
              >
                ថ្ងៃច្បាប់ឈឺ {new Date().getFullYear()}
              </div>
              <div style={{ fontSize: 13, color: T.textSoft }}>
                បានប្រើ {sickBal.used} ក្នុងចំណោម {sickBal.quota} ថ្ងៃ
              </div>
            </div>
            <div
              style={{
                fontSize: 22,
                fontWeight: 700,
                fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
                color: sickBal.remaining <= 0 ? T.rose : T.blue,
              }}
            >
              នៅសល់ {sickBal.remaining} ថ្ងៃ
            </div>
          </Card>
        </div>
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
                    <LeaveDecisionNote r={r} admins={admins} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>
        {modal && (
          <Modal title="ស្នើសុំច្បាប់ឈប់សម្រាក" onClose={() => setModal(false)}>
            <LeaveRequestForm
              remaining={{ annual: bal.remaining, sick: sickBal.remaining }}
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
                  <td>
                    {getLeaveTypeLabel(lang)[r.type] || r.type}
                    {r.type === "annual" && emp && (
                      <div style={{ fontSize: 10.5, color: T.muted }}>
                        នៅសល់ {annualLeaveBalance(emp, leaveRequests).remaining}{" "}
                        ថ្ងៃ
                      </div>
                    )}
                    {r.type === "sick" && emp && (
                      <div style={{ fontSize: 10.5, color: T.muted }}>
                        នៅសល់ {sickLeaveBalance(emp, leaveRequests).remaining}{" "}
                        ថ្ងៃ
                      </div>
                    )}
                  </td>
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
                    <LeaveDecisionNote r={r} admins={admins} />
                  </td>
                  <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                    {r.status === "pending"
                      ? canApprove && (
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
                              {t.lv.approve}
                            </Button>
                            <Button
                              size="sm"
                              variant="danger"
                              onClick={() => setRejectFor(r)}
                            >
                              {t.lv.reject}
                            </Button>
                          </div>
                        )
                      : isSuperAdmin && (
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
                        )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
      {rejectFor && (
        <LeaveRejectModal
          onCancel={() => setRejectFor(null)}
          onConfirm={(reason) => reject(rejectFor, reason)}
        />
      )}
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
function OvertimeRequestForm({ onSave, onCancel, holidays }) {
  const { t } = useLang();
  const [f, setF] = useState({
    date: todayStr(),
    hours: "",
    dayType: suggestDayType(todayStr(), holidays),
    reason: "",
  });
  const set = (k) => (e) => {
    const val = e.target.value;
    if (k === "date") {
      setF({ ...f, date: val, dayType: suggestDayType(val, holidays) });
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
          <DatePicker value={f.date} onChange={set("date")} />
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
      minSalaryThreshold: Number(f.minSalaryThreshold) || 0,
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
          {Number(payrollPolicy.minSalaryThreshold) > 0
            ? ` · ≥ ${fmtMoney(payrollPolicy.minSalaryThreshold)}`
            : ""}
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
          <div style={{ marginTop: 12 }}>
            <Field label={t.pay.minSalaryThresholdLabel}>
              <Input
                type="number"
                step="1"
                min="0"
                placeholder="0"
                value={f.minSalaryThreshold}
                onChange={set("minSalaryThreshold")}
              />
            </Field>
            <p style={{ fontSize: 11.5, color: T.muted, marginTop: 6 }}>
              {Number(f.minSalaryThreshold) > 0
                ? t.pay.minSalaryThresholdHint
                : t.pay.noThreshold}
            </p>
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
  const { t, lang } = useLang();
  if (r.status !== "approved" && r.status !== "rejected") return null;
  const decider = admins.find((a) => a.id === r.decidedById);
  const name = r.decidedByName || decider?.name || "—";
  const roleLabel = adminRoleLabel(r.decidedByRole, lang);
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
  canApprove,
  holidays,
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
              holidays={holidays}
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
                    {r.status === "pending"
                      ? canApprove && (
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
                        )
                      : isSuperAdmin && (
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
   Attendance corrections
   Employees request a fix for a missed/incorrect clock-in or
   clock-out (date + requested check-in/out + reason). Admins
   approve or reject; approving upserts the attendance record for
   that employee/date, mirroring a manual admin edit.
----------------------------------------------------------------*/
function AcRejectModal({ onCancel, onConfirm }) {
  const { t } = useLang();
  const [reason, setReason] = useState("");
  const trimmed = reason.trim();
  return (
    <Modal title={t.ac.rejectTitle} onClose={onCancel} width={420}>
      <Field label={t.ac.rejectReason}>
        <textarea
          className="wf-input"
          rows={3}
          style={{ resize: "vertical", fontFamily: "inherit" }}
          value={reason}
          onChange={(e) => setReason(e.target.value)}
          placeholder={t.ac.rejectReasonPlaceholder}
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
          {t.ac.rejectReasonRequired}
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
          {t.ac.reject}
        </Button>
      </div>
    </Modal>
  );
}

function AcDecisionNote({ r, admins }) {
  const { t } = useLang();
  if (r.status !== "approved" && r.status !== "rejected") return null;
  const decider = admins.find((a) => a.id === r.decidedById);
  const name = r.decidedByName || decider?.name || "—";
  return (
    <div style={{ fontSize: 11.5, color: T.textSoft, marginTop: 3 }}>
      {r.status === "approved" ? (
        <span>
          {t.ac.approvedBy} <strong>{name}</strong>
        </span>
      ) : (
        <span style={{ color: T.rose }}>
          {t.ac.rejectedBy} <strong>{name}</strong>
          {r.decisionReason ? ` — ${r.decisionReason}` : ""}
        </span>
      )}
    </div>
  );
}

function AttendanceCorrectionForm({ onSave, onCancel }) {
  const { t } = useLang();
  const [f, setF] = useState({
    date: todayStr(),
    requestedCheckIn: "",
    requestedCheckOut: "",
    reason: "",
  });
  const set = (k) => (e) =>
    setF({ ...f, [k]: typeof e === "string" ? e : e.target.value });
  const hasTime = !!(f.requestedCheckIn || f.requestedCheckOut);
  const invalid = !f.date || !hasTime || !f.reason.trim();
  return (
    <div>
      <Field label={t.ac.date}>
        <DatePicker value={f.date} onChange={set("date")} />
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label={t.ac.requestedCheckIn}>
          <TimePicker
            value={f.requestedCheckIn}
            onChange={(e) => setF({ ...f, requestedCheckIn: e.target.value })}
          />
        </Field>
        <Field label={t.ac.requestedCheckOut}>
          <TimePicker
            value={f.requestedCheckOut}
            onChange={(e) => setF({ ...f, requestedCheckOut: e.target.value })}
          />
        </Field>
      </div>
      {!hasTime && (
        <p
          style={{
            fontSize: 12.5,
            color: T.rose,
            marginTop: -8,
            marginBottom: 12,
          }}
        >
          {t.ac.needOneField}
        </p>
      )}
      <Field label={t.ac.reason}>
        <textarea
          className="wf-input"
          rows={3}
          style={{ resize: "vertical", fontFamily: "inherit" }}
          value={f.reason}
          onChange={set("reason")}
          placeholder={t.ac.reasonPlaceholder}
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
          onClick={() =>
            onSave({
              date: f.date,
              requestedCheckIn: f.requestedCheckIn || null,
              requestedCheckOut: f.requestedCheckOut || null,
              reason: f.reason.trim(),
            })
          }
          disabled={invalid}
        >
          {t.save}
        </Button>
      </div>
    </div>
  );
}

function AttendanceCorrections({
  role,
  currentAdmin,
  currentEmp,
  employees,
  admins,
  attendanceCorrections,
  setAttendanceCorrections,
  attendance,
  setAttendance,
  isSuperAdmin,
  canApprove,
}) {
  const { t } = useLang();
  const [modal, setModal] = useState(false);
  const [rejectFor, setRejectFor] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const empOf = (id) => employees.find((e) => e.id === id);

  const approve = (req) => {
    setAttendanceCorrections(
      attendanceCorrections.map((r) =>
        r.id === req.id
          ? {
              ...r,
              status: "approved",
              decidedById: currentAdmin?.id || null,
              decidedByName: currentAdmin?.name || "",
              decisionReason: "",
              reviewedAt: new Date().toISOString(),
            }
          : r,
      ),
    );
    // Upsert the attendance record for that employee/date, same as a
    // manual admin edit would — the requested times become the record.
    const existing = attendance.find(
      (a) => a.employeeId === req.employeeId && a.date === req.date,
    );
    const patch = {
      checkIn: req.requestedCheckIn || existing?.checkIn || null,
      checkOut: req.requestedCheckOut || existing?.checkOut || null,
      status: "present",
    };
    if (existing) {
      setAttendance(
        attendance.map((a) => (a.id === existing.id ? { ...a, ...patch } : a)),
      );
    } else {
      setAttendance([
        ...attendance,
        {
          id: uid("a"),
          employeeId: req.employeeId,
          date: req.date,
          ...patch,
        },
      ]);
    }
  };
  const reject = (req, reason) => {
    setAttendanceCorrections(
      attendanceCorrections.map((r) =>
        r.id === req.id
          ? {
              ...r,
              status: "rejected",
              decidedById: currentAdmin?.id || null,
              decidedByName: currentAdmin?.name || "",
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
    setAttendanceCorrections([
      ...attendanceCorrections,
      {
        id: uid("ac"),
        employeeId: currentEmp.id,
        date: f.date,
        requestedCheckIn: f.requestedCheckIn,
        requestedCheckOut: f.requestedCheckOut,
        reason: f.reason,
        status: "pending",
        decidedById: null,
        decidedByName: "",
        decisionReason: "",
        createdAt: new Date().toISOString(),
        reviewedAt: null,
      },
    ]);
    setModal(false);
  };

  if (role !== "admin" && currentEmp) {
    const mine = attendanceCorrections
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
            <Plus size={15} /> {t.ac.addBtn}
          </Button>
        </div>
        <Card style={{ overflowX: "auto" }}>
          <table className="wf-table">
            <thead>
              <tr>
                <th>{t.ac.date}</th>
                <th>{t.ac.requestedCheckIn}</th>
                <th>{t.ac.requestedCheckOut}</th>
                <th>{t.ac.reason}</th>
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
                    {t.ac.noRequest}
                  </td>
                </tr>
              )}
              {mine.map((r) => (
                <tr key={r.id}>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {r.date}
                  </td>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {r.requestedCheckIn || "—"}
                  </td>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {r.requestedCheckOut || "—"}
                  </td>
                  <td
                    style={{ fontSize: 12.5, color: T.textSoft, maxWidth: 200 }}
                  >
                    {r.reason || "—"}
                  </td>
                  <td>
                    <StatusPill status={r.status} />
                    <AcDecisionNote r={r} admins={admins} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>
        {modal && (
          <Modal title={t.ac.addBtn} onClose={() => setModal(false)}>
            <AttendanceCorrectionForm
              onSave={submit}
              onCancel={() => setModal(false)}
            />
          </Modal>
        )}
      </div>
    );
  }

  // Admin view — pending requests surfaced on top, newest first.
  const sorted = [...attendanceCorrections].sort((a, b) => {
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
              <th>{t.employee}</th>
              <th>{t.ac.date}</th>
              <th>{t.ac.requestedCheckIn}</th>
              <th>{t.ac.requestedCheckOut}</th>
              <th>{t.ac.reason}</th>
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
                  {t.ac.noRequest}
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
                    {r.requestedCheckIn || "—"}
                  </td>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {r.requestedCheckOut || "—"}
                  </td>
                  <td
                    style={{ fontSize: 12.5, color: T.textSoft, maxWidth: 200 }}
                  >
                    {r.reason || "—"}
                  </td>
                  <td>
                    <StatusPill status={r.status} />
                    <AcDecisionNote r={r} admins={admins} />
                  </td>
                  <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                    {r.status === "pending"
                      ? canApprove && (
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
                              <ThumbsUp size={13} /> {t.ac.approve}
                            </Button>
                            <Button
                              size="sm"
                              variant="danger"
                              onClick={() => setRejectFor(r)}
                            >
                              <ThumbsDown size={13} /> {t.ac.reject}
                            </Button>
                          </div>
                        )
                      : isSuperAdmin && (
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
                        )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
      {rejectFor && (
        <AcRejectModal
          onCancel={() => setRejectFor(null)}
          onConfirm={(reason) => reject(rejectFor, reason)}
        />
      )}
      {confirmDel && (
        <ConfirmDialog
          text={t.ac.confirmDel}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setAttendanceCorrections(
              attendanceCorrections.filter((r) => r.id !== confirmDel.id),
            );
            setConfirmDel(null);
          }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Announcements — admin posts company-wide messages that every
   employee (and admin) sees in a shared feed, newest first.
----------------------------------------------------------------*/
function AnnouncementForm({ initial, onSave, onCancel }) {
  const { t } = useLang();
  const [f, setF] = useState(initial || { title: "", body: "" });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const invalid = !f.title.trim() || !f.body.trim();
  return (
    <div>
      <Field label={t.ann.titleLabel}>
        <Input
          value={f.title}
          onChange={set("title")}
          placeholder={t.ann.titlePlaceholder}
        />
      </Field>
      <Field label={t.ann.bodyLabel}>
        <textarea
          className="wf-input"
          rows={5}
          style={{ resize: "vertical", fontFamily: "inherit" }}
          value={f.body}
          onChange={set("body")}
          placeholder={t.ann.bodyPlaceholder}
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
          onClick={() => onSave({ title: f.title.trim(), body: f.body.trim() })}
          disabled={invalid}
        >
          {t.save}
        </Button>
      </div>
    </div>
  );
}

function Announcements({
  role,
  currentAdmin,
  announcements,
  setAnnouncements,
  isSuperAdmin,
}) {
  const { t, lang } = useLang();
  const [modal, setModal] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);

  const save = (f) => {
    if (modal.mode === "add") {
      setAnnouncements([
        ...announcements,
        {
          id: uid("ann"),
          title: f.title,
          body: f.body,
          createdById: currentAdmin?.id || null,
          createdByName: currentAdmin?.name || "",
          createdAt: new Date().toISOString(),
        },
      ]);
    } else {
      setAnnouncements(
        announcements.map((a) =>
          a.id === modal.data.id ? { ...a, title: f.title, body: f.body } : a,
        ),
      );
    }
    setModal(null);
  };

  const sorted = [...announcements].sort((a, b) =>
    b.createdAt.localeCompare(a.createdAt),
  );

  return (
    <div>
      {role === "admin" && (
        <div
          style={{
            display: "flex",
            justifyContent: "flex-end",
            marginBottom: 16,
          }}
        >
          <Button variant="accent" onClick={() => setModal({ mode: "add" })}>
            <Plus size={15} /> {t.ann.addBtn}
          </Button>
        </div>
      )}
      {sorted.length === 0 && (
        <Card
          style={{
            textAlign: "center",
            padding: "32px 0",
            color: T.muted,
            fontSize: 13,
          }}
        >
          {t.ann.noAnn}
        </Card>
      )}
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        {sorted.map((a) => (
          <Card key={a.id} style={{ padding: 16 }}>
            <div
              style={{
                display: "flex",
                alignItems: "flex-start",
                justifyContent: "space-between",
                gap: 10,
                marginBottom: 8,
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  fontWeight: 700,
                  color: T.ink,
                  fontSize: 14,
                }}
              >
                <Megaphone size={16} color={T.goldText} /> {a.title}
              </div>
              {role === "admin" && (
                <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                  <button
                    onClick={() => setModal({ mode: "edit", data: a })}
                    style={{
                      background: "none",
                      border: "none",
                      cursor: "pointer",
                      color: T.mutedLight,
                    }}
                  >
                    <Pencil size={14} />
                  </button>
                  {isSuperAdmin && (
                    <button
                      onClick={() => setConfirmDel(a)}
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
                </div>
              )}
            </div>
            <div
              style={{
                fontSize: 13,
                color: T.textSoft,
                whiteSpace: "pre-wrap",
                marginBottom: 10,
              }}
            >
              {a.body}
            </div>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                fontSize: 11,
                color: T.muted,
              }}
            >
              <span>
                {t.ann.postedBy} {a.createdByName || "—"}
              </span>
              <span style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                {timeAgoLabel(a.createdAt)}
              </span>
            </div>
          </Card>
        ))}
      </div>
      {modal && (
        <Modal
          title={modal.mode === "add" ? t.ann.addTitle : t.ann.editTitle}
          onClose={() => setModal(null)}
        >
          <AnnouncementForm
            initial={modal.data}
            onSave={save}
            onCancel={() => setModal(null)}
          />
        </Modal>
      )}
      {confirmDel && (
        <ConfirmDialog
          text={t.ann.confirmDel}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setAnnouncements(
              announcements.filter((a) => a.id !== confirmDel.id),
            );
            setConfirmDel(null);
          }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Performance reviews — admin writes periodic reviews (rating +
   notes) for each employee; employees can view their own history
   read-only.
----------------------------------------------------------------*/
function RatingStars({ value, onChange, size = 16 }) {
  const stars = [1, 2, 3, 4, 5];
  return (
    <div style={{ display: "flex", gap: 4 }}>
      {stars.map((n) => (
        <span
          key={n}
          onClick={onChange ? () => onChange(n) : undefined}
          style={{
            cursor: onChange ? "pointer" : "default",
            lineHeight: 0,
          }}
        >
          <Star
            size={size}
            fill={n <= value ? T.gold : "none"}
            color={n <= value ? T.gold : T.mutedLight}
            strokeWidth={1.75}
          />
        </span>
      ))}
    </div>
  );
}

function PerformanceReviewForm({ initial, employees, onSave, onCancel }) {
  const { t } = useLang();
  const [f, setF] = useState(
    initial || {
      employeeId: employees[0]?.id || "",
      period: "",
      rating: 5,
      strengths: "",
      improvements: "",
    },
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const invalid = !f.employeeId || !f.period.trim();
  return (
    <div>
      <Field label={t.pr.employee}>
        <Select
          value={f.employeeId}
          onChange={set("employeeId")}
          disabled={!!initial}
        >
          {employees.map((e) => (
            <option key={e.id} value={e.id}>
              {e.name} ({e.code})
            </option>
          ))}
        </Select>
      </Field>
      <Field label={t.pr.period}>
        <Input
          value={f.period}
          onChange={set("period")}
          placeholder={t.pr.periodPlaceholder}
        />
      </Field>
      <Field label={t.pr.rating}>
        <RatingStars
          value={f.rating}
          onChange={(n) => setF({ ...f, rating: n })}
          size={22}
        />
      </Field>
      <Field label={t.pr.strengths}>
        <textarea
          className="wf-input"
          rows={3}
          style={{ resize: "vertical", fontFamily: "inherit" }}
          value={f.strengths}
          onChange={set("strengths")}
          placeholder={t.pr.strengthsPlaceholder}
        />
      </Field>
      <Field label={t.pr.improvements}>
        <textarea
          className="wf-input"
          rows={3}
          style={{ resize: "vertical", fontFamily: "inherit" }}
          value={f.improvements}
          onChange={set("improvements")}
          placeholder={t.pr.improvementsPlaceholder}
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
          onClick={() =>
            onSave({
              ...f,
              period: f.period.trim(),
              strengths: f.strengths.trim(),
              improvements: f.improvements.trim(),
            })
          }
          disabled={invalid}
        >
          {t.save}
        </Button>
      </div>
    </div>
  );
}

function PerformanceReviews({
  role,
  currentAdmin,
  currentEmp,
  employees,
  performanceReviews,
  setPerformanceReviews,
  isSuperAdmin,
}) {
  const { t } = useLang();
  const [modal, setModal] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const empOf = (id) => employees.find((e) => e.id === id);

  const save = (f) => {
    if (modal.mode === "add") {
      setPerformanceReviews([
        ...performanceReviews,
        {
          id: uid("pr"),
          employeeId: f.employeeId,
          period: f.period,
          rating: f.rating,
          strengths: f.strengths,
          improvements: f.improvements,
          reviewedById: currentAdmin?.id || null,
          reviewedByName: currentAdmin?.name || "",
          createdAt: new Date().toISOString(),
        },
      ]);
    } else {
      setPerformanceReviews(
        performanceReviews.map((r) =>
          r.id === modal.data.id
            ? {
                ...r,
                period: f.period,
                rating: f.rating,
                strengths: f.strengths,
                improvements: f.improvements,
              }
            : r,
        ),
      );
    }
    setModal(null);
  };

  if (role !== "admin" && currentEmp) {
    const mine = performanceReviews
      .filter((r) => r.employeeId === currentEmp.id)
      .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
    return (
      <div>
        {mine.length === 0 && (
          <Card
            style={{
              textAlign: "center",
              padding: "32px 0",
              color: T.muted,
              fontSize: 13,
            }}
          >
            {t.pr.noReview}
          </Card>
        )}
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {mine.map((r) => (
            <Card key={r.id} style={{ padding: 16 }}>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  marginBottom: 8,
                }}
              >
                <span style={{ fontWeight: 600, color: T.ink, fontSize: 13 }}>
                  {r.period}
                </span>
                <RatingStars value={r.rating} />
              </div>
              {r.strengths && (
                <div
                  style={{ fontSize: 12.5, color: T.textSoft, marginBottom: 6 }}
                >
                  <strong style={{ color: T.forestText }}>
                    {t.pr.strengths}:
                  </strong>{" "}
                  {r.strengths}
                </div>
              )}
              {r.improvements && (
                <div
                  style={{ fontSize: 12.5, color: T.textSoft, marginBottom: 6 }}
                >
                  <strong style={{ color: T.goldText }}>
                    {t.pr.improvements}:
                  </strong>{" "}
                  {r.improvements}
                </div>
              )}
              <div style={{ fontSize: 11, color: T.muted, marginTop: 8 }}>
                {t.pr.reviewedBy} {r.reviewedByName || "—"}
              </div>
            </Card>
          ))}
        </div>
      </div>
    );
  }

  const sorted = [...performanceReviews].sort((a, b) =>
    b.createdAt.localeCompare(a.createdAt),
  );
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
          onClick={() => setModal({ mode: "add" })}
          disabled={employees.length === 0}
        >
          <Plus size={15} /> {t.pr.addBtn}
        </Button>
      </div>
      <Card style={{ overflowX: "auto" }}>
        <table className="wf-table">
          <thead>
            <tr>
              <th>{t.pr.employee}</th>
              <th>{t.pr.period}</th>
              <th>{t.pr.rating}</th>
              <th>{t.pr.reviewedBy}</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {sorted.length === 0 && (
              <tr>
                <td
                  colSpan={5}
                  style={{
                    textAlign: "center",
                    color: T.muted,
                    padding: "24px 0",
                  }}
                >
                  {t.pr.noReview}
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
                  <td style={{ fontSize: 12.5 }}>{r.period}</td>
                  <td>
                    <RatingStars value={r.rating} size={14} />
                  </td>
                  <td style={{ fontSize: 12, color: T.textSoft }}>
                    {r.reviewedByName || "—"}
                  </td>
                  <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                    <div
                      style={{
                        display: "flex",
                        gap: 6,
                        justifyContent: "flex-end",
                      }}
                    >
                      <button
                        onClick={() => setModal({ mode: "edit", data: r })}
                        style={{
                          background: "none",
                          border: "none",
                          cursor: "pointer",
                          color: T.mutedLight,
                        }}
                      >
                        <Pencil size={14} />
                      </button>
                      {isSuperAdmin && (
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
                      )}
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
          title={modal.mode === "add" ? t.pr.addTitle : t.pr.editTitle}
          onClose={() => setModal(null)}
        >
          <PerformanceReviewForm
            initial={modal.data}
            employees={employees}
            onSave={save}
            onCancel={() => setModal(null)}
          />
        </Modal>
      )}
      {confirmDel && (
        <ConfirmDialog
          text={t.pr.confirmDel}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setPerformanceReviews(
              performanceReviews.filter((r) => r.id !== confirmDel.id),
            );
            setConfirmDel(null);
          }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Employee documents — admin uploads/manages files (CV, contract,
   ID card, etc.) per employee. Files are read client-side and
   stored as base64 data URLs (no external storage bucket needed),
   so uploads are capped to keep row sizes reasonable.
----------------------------------------------------------------*/
const MAX_DOC_BYTES = 4 * 1024 * 1024;

function getDocCategoryLabel(t) {
  return {
    cv: t.doc.catCv,
    contract: t.doc.catContract,
    id: t.doc.catId,
    other: t.doc.catOther,
  };
}

function DocUploadRow({ emp, currentAdmin, documents, setDocuments, t }) {
  const [category, setCategory] = useState("cv");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const inputRef = useRef(null);

  const onFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    if (file.size > MAX_DOC_BYTES) {
      setError(t.doc.tooLarge);
      e.target.value = "";
      return;
    }
    setBusy(true);
    try {
      const dataUrl = await fileToDataUrl(file);
      setDocuments([
        ...documents,
        {
          id: uid("doc"),
          employeeId: emp.id,
          category,
          fileName: file.name,
          mimeType: file.type || "application/octet-stream",
          dataUrl,
          uploadedByName: currentAdmin?.name || "",
          createdAt: new Date().toISOString(),
        },
      ]);
    } finally {
      setBusy(false);
      if (inputRef.current) inputRef.current.value = "";
    }
  };

  return (
    <div
      style={{
        display: "flex",
        gap: 8,
        alignItems: "flex-end",
        marginBottom: 16,
        flexWrap: "wrap",
      }}
    >
      <Field label={t.doc.category}>
        <Select value={category} onChange={(e) => setCategory(e.target.value)}>
          <option value="cv">{t.doc.catCv}</option>
          <option value="contract">{t.doc.catContract}</option>
          <option value="id">{t.doc.catId}</option>
          <option value="other">{t.doc.catOther}</option>
        </Select>
      </Field>
      <div>
        <input
          ref={inputRef}
          type="file"
          id="doc-upload-input"
          style={{ display: "none" }}
          onChange={onFile}
        />
        <Button
          variant="ghost"
          disabled={busy}
          onClick={() => inputRef.current?.click()}
        >
          {busy ? (
            <Loader2
              size={15}
              style={{ animation: "spin 1s linear infinite" }}
            />
          ) : (
            <Upload size={15} />
          )}{" "}
          {t.doc.chooseFile}
        </Button>
      </div>
      {error && <span style={{ fontSize: 12, color: T.rose }}>{error}</span>}
    </div>
  );
}

function EmployeeDocumentsModal({
  emp,
  currentAdmin,
  documents,
  setDocuments,
  isSuperAdmin,
  onClose,
}) {
  const { t } = useLang();
  const [confirmDel, setConfirmDel] = useState(null);
  const CAT_LABEL = getDocCategoryLabel(t);
  const mine = documents
    .filter((d) => d.employeeId === emp.id)
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));

  return (
    <Modal title={`${t.doc.title} · ${emp.name}`} onClose={onClose} width={520}>
      <DocUploadRow
        emp={emp}
        currentAdmin={currentAdmin}
        documents={documents}
        setDocuments={setDocuments}
        t={t}
      />
      {mine.length === 0 && (
        <div
          style={{
            textAlign: "center",
            color: T.muted,
            fontSize: 13,
            padding: "20px 0",
          }}
        >
          {t.doc.noDocs}
        </div>
      )}
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {mine.map((d) => (
          <div
            key={d.id}
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              gap: 10,
              padding: "10px 12px",
              border: `1px solid ${T.lineSoft}`,
              borderRadius: 10,
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                minWidth: 0,
              }}
            >
              <FileText
                size={18}
                color={T.forestText}
                style={{ flexShrink: 0 }}
              />
              <div style={{ minWidth: 0 }}>
                <div
                  style={{
                    fontSize: 12.5,
                    fontWeight: 600,
                    color: T.ink,
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                    whiteSpace: "nowrap",
                  }}
                >
                  {d.fileName}
                </div>
                <div style={{ fontSize: 10.5, color: T.muted }}>
                  {CAT_LABEL[d.category] || d.category} · {t.doc.uploadedBy}{" "}
                  {d.uploadedByName || "—"}
                </div>
              </div>
            </div>
            <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
              <a
                href={d.dataUrl}
                download={d.fileName}
                target="_blank"
                rel="noreferrer"
                style={{ color: T.forestText }}
              >
                <Download size={16} />
              </a>
              {isSuperAdmin && (
                <button
                  onClick={() => setConfirmDel(d)}
                  style={{
                    background: "none",
                    border: "none",
                    cursor: "pointer",
                    color: T.mutedLight,
                  }}
                >
                  <Trash2 size={16} />
                </button>
              )}
            </div>
          </div>
        ))}
      </div>
      {confirmDel && (
        <ConfirmDialog
          text={t.doc.confirmDel}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setDocuments(documents.filter((d) => d.id !== confirmDel.id));
            setConfirmDel(null);
          }}
        />
      )}
    </Modal>
  );
}

function MyDocuments({ currentEmp, documents }) {
  const { t } = useLang();
  const CAT_LABEL = getDocCategoryLabel(t);
  const mine = documents
    .filter((d) => d.employeeId === currentEmp.id)
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  return (
    <div>
      {mine.length === 0 && (
        <Card
          style={{
            textAlign: "center",
            padding: "32px 0",
            color: T.muted,
            fontSize: 13,
          }}
        >
          {t.doc.noDocs}
        </Card>
      )}
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {mine.map((d) => (
          <Card
            key={d.id}
            style={{
              padding: "12px 16px",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              gap: 10,
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                minWidth: 0,
              }}
            >
              <FileText
                size={18}
                color={T.forestText}
                style={{ flexShrink: 0 }}
              />
              <div style={{ minWidth: 0 }}>
                <div
                  style={{
                    fontSize: 13,
                    fontWeight: 600,
                    color: T.ink,
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                    whiteSpace: "nowrap",
                  }}
                >
                  {d.fileName}
                </div>
                <div style={{ fontSize: 11, color: T.muted }}>
                  {CAT_LABEL[d.category] || d.category}
                </div>
              </div>
            </div>
            <a
              href={d.dataUrl}
              download={d.fileName}
              target="_blank"
              rel="noreferrer"
              style={{ color: T.forestText, flexShrink: 0 }}
            >
              <Download size={18} />
            </a>
          </Card>
        ))}
      </div>
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
          {ADMIN_RANKS.map((rank) => (
            <option key={rank} value={rank}>
              {adminRoleLabel(rank, lang)}
            </option>
          ))}
          <option value="superadmin">
            {adminRoleLabel("superadmin", lang)} — សិទ្ធិពេញលេញ
          </option>
        </Select>
      </Field>
      <p
        style={{
          fontSize: 11.5,
          color: T.muted,
          marginTop: -8,
          marginBottom: 14,
        }}
      >
        តើតួនាទីនីមួយៗអាចធ្វើអ្វីបាន? កំណត់នៅទំព័រ "សិទ្ធិតួនាទី" (Superadmin
        ប៉ុណ្ណោះ)
      </p>
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

function RolePermissionsPage({ rolePermissions, setRolePermissions }) {
  const { t, lang } = useLang();
  // Local editable draft so clicking a checkbox doesn't fire a save on
  // every click — one explicit "Save" commits the whole matrix at once.
  const buildDraft = () => {
    const map = {};
    ADMIN_RANKS.forEach((rank) => {
      const saved = rolePermissions.find((r) => r.id === rank);
      map[rank] = saved
        ? { ...DEFAULT_ROLE_PERMISSIONS[rank], ...saved }
        : { ...DEFAULT_ROLE_PERMISSIONS[rank] };
    });
    return map;
  };
  const buildEmpDraft = () => {
    const saved = rolePermissions.find((r) => r.id === EMPLOYEE_MODULES_ID);
    return saved
      ? { ...DEFAULT_EMPLOYEE_MODULES, ...saved }
      : { ...DEFAULT_EMPLOYEE_MODULES };
  };
  const [draft, setDraft] = useState(buildDraft);
  const [empDraft, setEmpDraft] = useState(buildEmpDraft);
  const [dirty, setDirty] = useState(false);
  // If another Superadmin saves changes elsewhere while this tab is open,
  // pull in the fresh matrix — but only while this tab has no unsaved
  // edits of its own, so we never silently overwrite someone mid-edit.
  useEffect(() => {
    if (!dirty) {
      setDraft(buildDraft());
      setEmpDraft(buildEmpDraft());
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rolePermissions]);

  const toggle = (rank, key) => {
    setDraft((d) => ({ ...d, [rank]: { ...d[rank], [key]: !d[rank][key] } }));
    setDirty(true);
  };
  const toggleEmp = (key) => {
    setEmpDraft((d) => ({ ...d, [key]: !d[key] }));
    setDirty(true);
  };

  const save = () => {
    const next = ADMIN_RANKS.map((rank) => ({ id: rank, ...draft[rank] }));
    next.push({ id: EMPLOYEE_MODULES_ID, ...empDraft });
    setRolePermissions(next);
    setDirty(false);
  };

  return (
    <div>
      <Card style={{ padding: 20, marginBottom: 18 }}>
        <h2
          style={{
            fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
            fontSize: 18,
            fontWeight: 600,
            color: T.ink,
          }}
        >
          {t.nav.rolePerms}
        </h2>
        <p style={{ fontSize: 12, color: T.muted, marginTop: 6 }}>
          {lang === "en"
            ? "Choose what each rank can access below. Superadmin always has full access (including managing admin accounts) and isn't shown here — that stays superadmin-exclusive."
            : "កំណត់ថាតួនាទីនីមួយៗអាចប្រើមុខងារអ្វីខ្លះខាងក្រោម។ Superadmin មានសិទ្ធិពេញលេញជានិច្ច (រួមទាំងគ្រប់គ្រងគណនីអ្នកគ្រប់គ្រង) ដូច្នេះមិនបង្ហាញនៅទីនេះទេ — សិទ្ធិនោះនៅតែសម្រាប់ Superadmin ប៉ុណ្ណោះ។"}
        </p>
      </Card>
      <Card style={{ overflowX: "auto", padding: 0, marginBottom: 24 }}>
        <table className="wf-table">
          <thead>
            <tr>
              <th style={{ whiteSpace: "nowrap" }}>
                {lang === "en" ? "Rank" : "តួនាទី"}
              </th>
              {PERMISSION_MODULES.map((key) => (
                <th
                  key={key}
                  style={{ textAlign: "center", whiteSpace: "nowrap" }}
                >
                  {permissionLabel(key, lang)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {ADMIN_RANKS.map((rank) => (
              <tr key={rank}>
                <td
                  style={{
                    fontWeight: 600,
                    color: T.ink,
                    whiteSpace: "nowrap",
                  }}
                >
                  {adminRoleLabel(rank, lang)}
                </td>
                {PERMISSION_MODULES.map((key) => (
                  <td key={key} style={{ textAlign: "center" }}>
                    <input
                      type="checkbox"
                      checked={!!draft[rank]?.[key]}
                      onChange={() => toggle(rank, key)}
                      style={{ width: 16, height: 16, cursor: "pointer" }}
                    />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </Card>

      <Card style={{ padding: 20, marginBottom: 18 }}>
        <h2
          style={{
            fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
            fontSize: 18,
            fontWeight: 600,
            color: T.ink,
          }}
        >
          {lang === "en"
            ? "Employee Self-Service Portal"
            : "កម្មវិធីសម្រាប់បុគ្គលិក (Self-Service)"}
        </h2>
        <p style={{ fontSize: 12, color: T.muted, marginTop: 6 }}>
          {lang === "en"
            ? "Turn features off here to hide them for every employee company-wide. This applies to all employees equally — they don't have ranks the way admin accounts do. Dashboard always stays on."
            : "បិទមុខងារនៅទីនេះ ដើម្បីលាក់វាចេញពីបុគ្គលិកទាំងអស់ក្នុងក្រុមហ៊ុនតែម្តង។ ការកំណត់នេះអនុវត្តដូចគ្នាចំពោះបុគ្គលិកទាំងអស់ ព្រោះពួកគេគ្មានតួនាទីខុសៗគ្នាដូច admin ទេ។ Dashboard នៅតែបើកជានិច្ច។"}
        </p>
      </Card>
      <Card style={{ overflowX: "auto", padding: 0 }}>
        <table className="wf-table">
          <thead>
            <tr>
              {EMPLOYEE_MODULES.map((key) => (
                <th
                  key={key}
                  style={{ textAlign: "center", whiteSpace: "nowrap" }}
                >
                  {employeeModuleLabel(key, lang)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            <tr>
              {EMPLOYEE_MODULES.map((key) => (
                <td key={key} style={{ textAlign: "center" }}>
                  <input
                    type="checkbox"
                    checked={empDraft[key] !== false}
                    onChange={() => toggleEmp(key)}
                    style={{ width: 16, height: 16, cursor: "pointer" }}
                  />
                </td>
              ))}
            </tr>
          </tbody>
        </table>
      </Card>

      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          alignItems: "center",
          gap: 10,
          marginTop: 16,
        }}
      >
        {dirty && (
          <span style={{ fontSize: 12, color: T.gold }}>
            {lang === "en" ? "Unsaved changes" : "មិនទាន់រក្សាទុក"}
          </span>
        )}
        <Button variant="accent" onClick={save} disabled={!dirty}>
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
                          ? { background: "rgba(91,141,239,0.14)", color: T.blue }
                          : { background: T.forestSoft, color: T.forestText }
                      }
                    >
                      {adminRoleLabel(a.role, lang)}
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
  saveError,
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
  // See the identical flag in AdminSettings: only treat a saveError as
  // "my save just failed" if this screen actually attempted one.
  const [attemptedSave, setAttemptedSave] = useState(false);

  useEffect(() => {
    if (attemptedSave && saveError) setSaved(false);
  }, [saveError, attemptedSave]);

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
    setAttemptedSave(false);
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
      setAttemptedSave(false);
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
    setAttemptedSave(true);
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
        {attemptedSave && saveError && (
          <p
            style={{
              fontSize: 12.5,
              color: T.rose,
              display: "flex",
              alignItems: "flex-start",
              gap: 5,
              marginBottom: 10,
            }}
          >
            <AlertCircle size={14} style={{ flexShrink: 0, marginTop: 1 }} />
            <span>
              {t.settings.saveFailed} {saveError}
            </span>
          </p>
        )}
        {saved && !saveError && (
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
        <Button variant="accent" onClick={saveProfile}>
          {t.save}
        </Button>
      </Card>

      <Card style={{ padding: 20 }}>
        <h3
          style={{
            fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
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
          fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
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
function AdminSettings({
  currentAdmin,
  admins,
  setAdmins,
  isSuperAdmin,
  saveError,
}) {
  const { t, lang } = useLang();
  const { theme, setTheme } = useTheme();
  const { branding, setBranding } = useBranding();
  const [f, setF] = useState({ name: currentAdmin.name || "" });
  const [nameError, setNameError] = useState("");
  const [photoPreview, setPhotoPreview] = useState(currentAdmin.photo || null);
  const [photoError, setPhotoError] = useState("");
  const [saved, setSaved] = useState(false);
  // True once this component has actually attempted a save — otherwise a
  // saveError left over from some unrelated earlier admins write (e.g.
  // another admin's edit) would incorrectly flag this screen as failed.
  const [attemptedSave, setAttemptedSave] = useState(false);

  useEffect(() => {
    if (attemptedSave && saveError) setSaved(false);
  }, [saveError, attemptedSave]);

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
      setAttemptedSave(false);
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
    setAttemptedSave(true);
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
              {adminRoleLabel(currentAdmin.role, lang)}
            </div>
          </div>
        </div>
        <Field label={t.settings.nameLabel}>
          <Input
            value={f.name}
            onChange={(e) => {
              setF({ ...f, name: e.target.value });
              setSaved(false);
              setAttemptedSave(false);
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
        {attemptedSave && saveError && (
          <p
            style={{
              fontSize: 12.5,
              color: T.rose,
              display: "flex",
              alignItems: "flex-start",
              gap: 5,
              marginBottom: 10,
            }}
          >
            <AlertCircle size={14} style={{ flexShrink: 0, marginTop: 1 }} />
            <span>
              {t.settings.saveFailed || "Save failed:"} {saveError}
            </span>
          </p>
        )}
        {saved && !saveError && (
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
            fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
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
              fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
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
                fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
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
   Audit Log page — read-only view of the audit_logs table written
   by writeAuditLog() (see the useSupabaseArray hook near the top of
   the file). Superadmin-only.
----------------------------------------------------------------*/
const AUDIT_ACTION_TONE = {
  create: { bg: T.forestSoft, fg: T.forestText },
  update: { bg: T.goldSoft, fg: T.goldText },
  delete: { bg: T.roseSoft, fg: T.rose },
  login: { bg: T.forestSoft, fg: T.forestText },
  logout: { bg: T.lineSoft, fg: T.muted },
};
function AuditActionBadge({ action, t }) {
  const tone = AUDIT_ACTION_TONE[action] || {
    bg: T.lineSoft,
    fg: T.muted,
  };
  const label =
    {
      create: t.audit.actionCreate,
      update: t.audit.actionUpdate,
      delete: t.audit.actionDelete,
      login: t.audit.actionLogin,
      logout: t.audit.actionLogout,
    }[action] || action;
  return (
    <span
      style={{
        fontSize: 11,
        fontWeight: 700,
        padding: "3px 9px",
        borderRadius: 999,
        background: tone.bg,
        color: tone.fg,
        whiteSpace: "nowrap",
      }}
    >
      {label}
    </span>
  );
}
function fmtAuditTime(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  try {
    return new Intl.DateTimeFormat("km-KH", {
      year: "numeric",
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    }).format(d);
  } catch {
    return iso;
  }
}
function AuditLogPage() {
  const { t } = useLang();
  const [logs, setLogs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [q, setQ] = useState("");
  const [tableFilter, setTableFilter] = useState("all");
  const [expanded, setExpanded] = useState(null);

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase
      .from("audit_logs")
      .select("*")
      .order("created_at", { ascending: false })
      .limit(300);
    if (error) {
      console.error("[audit] failed to load audit_logs:", error.message);
      setLogs([]);
    } else {
      setLogs(data || []);
    }
    setLoading(false);
  }, []);

  useEffect(() => {
    load();
  }, [load]);

  const tableNames = useMemo(() => {
    const seen = new Set(logs.map((l) => l.entity_table).filter(Boolean));
    return Array.from(seen).sort();
  }, [logs]);

  const filtered = useMemo(() => {
    const query = q.trim().toLowerCase();
    return logs.filter((l) => {
      if (tableFilter !== "all" && l.entity_table !== tableFilter) return false;
      if (!query) return true;
      const hay = [l.actor_name, l.entity_label, l.entity_table, l.action]
        .filter(Boolean)
        .join(" ")
        .toLowerCase();
      return hay.includes(query);
    });
  }, [logs, q, tableFilter]);
  const pg = usePagination(filtered, 25);

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
        <div>
          <h2
            style={{
              fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
              fontSize: 18,
              fontWeight: 600,
              color: T.ink,
              margin: 0,
            }}
          >
            {t.audit.title}
          </h2>
          <p style={{ fontSize: 12.5, color: T.muted, margin: "4px 0 0" }}>
            {t.audit.subtitle}
          </p>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <div style={{ position: "relative" }}>
            <Search
              size={14}
              color={T.muted}
              style={{
                position: "absolute",
                left: 10,
                top: "50%",
                transform: "translateY(-50%)",
              }}
            />
            <Input
              value={q}
              onChange={(e) => setQ(e.target.value)}
              placeholder={t.search}
              style={{ paddingLeft: 30, minWidth: 200 }}
            />
          </div>
          <Select
            value={tableFilter}
            onChange={(e) => setTableFilter(e.target.value)}
          >
            <option value="all">{t.audit.allTables}</option>
            {tableNames.map((tn) => (
              <option key={tn} value={tn}>
                {t.audit.tables[tn] || tn}
              </option>
            ))}
          </Select>
          <Button variant="ghost" size="sm" onClick={load}>
            <RefreshCw size={14} /> {t.audit.refresh}
          </Button>
        </div>
      </div>

      {loading && (
        <Card
          style={{
            textAlign: "center",
            padding: "32px 0",
            color: T.muted,
            fontSize: 13,
          }}
        >
          <Loader2
            size={20}
            color={T.forest}
            style={{ animation: "spin 1s linear infinite" }}
          />
          <div style={{ marginTop: 8 }}>{t.audit.loading}</div>
        </Card>
      )}

      {!loading && filtered.length === 0 && (
        <Card
          style={{
            textAlign: "center",
            padding: "32px 0",
            color: T.muted,
            fontSize: 13,
          }}
        >
          {t.audit.noLogs}
        </Card>
      )}

      {!loading && filtered.length > 0 && (
        <Card style={{ overflowX: "auto", padding: 0 }}>
          <table className="wf-table">
            <thead>
              <tr>
                <th>{t.audit.time}</th>
                <th>{t.audit.actor}</th>
                <th>{t.audit.action}</th>
                <th>{t.audit.entity}</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {pg.pageItems.map((l) => {
                const hasChanges = l.changes && Object.keys(l.changes).length;
                const isOpen = expanded === l.id;
                return (
                  <React.Fragment key={l.id}>
                    <tr
                      style={{ cursor: hasChanges ? "pointer" : "default" }}
                      onClick={() =>
                        hasChanges && setExpanded(isOpen ? null : l.id)
                      }
                    >
                      <td
                        style={{
                          fontFamily: "'JetBrains Mono',monospace",
                          fontSize: 11.5,
                          color: T.muted,
                          whiteSpace: "nowrap",
                        }}
                      >
                        {fmtAuditTime(l.created_at)}
                      </td>
                      <td>
                        <div
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: 8,
                          }}
                        >
                          <Avatar
                            name={l.actor_name || t.audit.unknownActor}
                            size={26}
                          />
                          <div>
                            <div
                              style={{
                                fontSize: 12.5,
                                fontWeight: 500,
                                color: T.ink,
                              }}
                            >
                              {l.actor_name || t.audit.unknownActor}
                            </div>
                            {l.actor_type && (
                              <div style={{ fontSize: 10.5, color: T.muted }}>
                                {l.actor_type === "admin"
                                  ? t.nav.admins
                                  : t.employee}
                              </div>
                            )}
                          </div>
                        </div>
                      </td>
                      <td>
                        <AuditActionBadge action={l.action} t={t} />
                      </td>
                      <td style={{ fontSize: 12.5 }}>
                        <div style={{ color: T.ink }}>
                          {l.entity_label || l.entity_id || "—"}
                        </div>
                        <div style={{ fontSize: 10.5, color: T.muted }}>
                          {t.audit.tables[l.entity_table] || l.entity_table}
                        </div>
                      </td>
                      <td style={{ textAlign: "right" }}>
                        {hasChanges &&
                          (isOpen ? (
                            <ChevronLeft size={14} color={T.muted} />
                          ) : (
                            <ChevronRight size={14} color={T.muted} />
                          ))}
                      </td>
                    </tr>
                    {isOpen && hasChanges && (
                      <tr>
                        <td
                          colSpan={5}
                          style={{ background: T.tableHeadBg, padding: 12 }}
                        >
                          <div
                            style={{
                              display: "flex",
                              flexDirection: "column",
                              gap: 6,
                            }}
                          >
                            {Object.entries(l.changes).map(([field, ch]) => (
                              <div
                                key={field}
                                style={{
                                  display: "flex",
                                  gap: 8,
                                  fontSize: 11.5,
                                  flexWrap: "wrap",
                                }}
                              >
                                <span
                                  style={{
                                    fontWeight: 700,
                                    color: T.ink,
                                    minWidth: 120,
                                  }}
                                >
                                  {field}
                                </span>
                                <span style={{ color: T.muted }}>
                                  {t.audit.changedFrom}
                                </span>
                                <span
                                  style={{
                                    color: T.rose,
                                    fontFamily: "'JetBrains Mono',monospace",
                                  }}
                                >
                                  {ch.from}
                                </span>
                                <span style={{ color: T.muted }}>
                                  {t.audit.changedTo}
                                </span>
                                <span
                                  style={{
                                    color: T.forestText,
                                    fontFamily: "'JetBrains Mono',monospace",
                                  }}
                                >
                                  {ch.to}
                                </span>
                              </div>
                            ))}
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
          <div style={{ padding: "0 16px 14px" }}>
            <Pagination {...pg} />
          </div>
        </Card>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Login Activity page — per-user sign-in/out history with a
   best-effort device summary, read from the login_activity table
   (see writeLoginActivity near the top of the file). Every signed-in
   user can see their own history by default; superadmins get an
   extra toggle to view everyone's.
----------------------------------------------------------------*/
function LoginActionBadge({ action, t }) {
  const isLogin = action === "login";
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 4,
        fontSize: 11,
        fontWeight: 700,
        padding: "3px 9px",
        borderRadius: 999,
        background: isLogin ? T.forestSoft : T.lineSoft,
        color: isLogin ? T.forestText : T.muted,
        whiteSpace: "nowrap",
      }}
    >
      {isLogin ? <LogIn size={11} /> : <LogOut size={11} />}
      {isLogin ? t.loginAct.actionLogin : t.loginAct.actionLogout}
    </span>
  );
}
function DeviceTypeIcon({ deviceType, size = 14, color }) {
  if (deviceType === "mobile") return <Smartphone size={size} color={color} />;
  if (deviceType === "tablet") return <Tablet size={size} color={color} />;
  return <Monitor size={size} color={color} />;
}
function deviceTypeLabel(deviceType, t) {
  if (deviceType === "mobile") return t.loginAct.deviceMobile;
  if (deviceType === "tablet") return t.loginAct.deviceTablet;
  return t.loginAct.deviceDesktop;
}
// Best-effort Android model code -> friendly marketing name, applied at
// display time (the raw code from getAndroidModel() is what's actually
// stored, so this table can be extended later without touching old
// rows). Keys are code *prefixes* with the trailing region/carrier
// letter dropped (Samsung in particular ships the same phone under many
// suffixes — "SM-A546E", "SM-A546B", "SM-A546W" are all a Galaxy A54),
// matched longest-prefix-first. Coverage favors brands common in this
// region (Samsung, Xiaomi/Redmi/POCO, OPPO, vivo, realme) — anything
// not in the table just falls back to showing the raw code, which is
// always correct even when we don't have a friendly name for it yet.
const ANDROID_MODEL_MAP = {
  // Samsung Galaxy S
  "SM-S911": "Galaxy S23",
  "SM-S916": "Galaxy S23+",
  "SM-S918": "Galaxy S23 Ultra",
  "SM-S921": "Galaxy S24",
  "SM-S926": "Galaxy S24+",
  "SM-S928": "Galaxy S24 Ultra",
  "SM-G991": "Galaxy S21",
  "SM-G996": "Galaxy S21+",
  "SM-G998": "Galaxy S21 Ultra",
  "SM-S901": "Galaxy S22",
  "SM-S906": "Galaxy S22+",
  "SM-S908": "Galaxy S22 Ultra",
  // Samsung Galaxy A
  "SM-A125": "Galaxy A12",
  "SM-A135": "Galaxy A13",
  "SM-A145": "Galaxy A14",
  "SM-A155": "Galaxy A15",
  "SM-A165": "Galaxy A16",
  "SM-A235": "Galaxy A23",
  "SM-A245": "Galaxy A24",
  "SM-A255": "Galaxy A25",
  "SM-A325": "Galaxy A32",
  "SM-A336": "Galaxy A33",
  "SM-A346": "Galaxy A34",
  "SM-A356": "Galaxy A35",
  "SM-A525": "Galaxy A52",
  "SM-A536": "Galaxy A53",
  "SM-A546": "Galaxy A54",
  "SM-A556": "Galaxy A55",
  "SM-A715": "Galaxy A71",
  "SM-A725": "Galaxy A72",
  // Samsung Galaxy M / Note / Z
  "SM-M115": "Galaxy M11",
  "SM-M127": "Galaxy M12",
  "SM-M135": "Galaxy M13",
  "SM-M146": "Galaxy M14",
  "SM-N970": "Galaxy Note10",
  "SM-N975": "Galaxy Note10+",
  "SM-N980": "Galaxy Note20",
  "SM-N985": "Galaxy Note20 Ultra",
  "SM-F711": "Galaxy Z Flip3",
  "SM-F721": "Galaxy Z Flip4",
  "SM-F731": "Galaxy Z Flip5",
  "SM-F926": "Galaxy Z Fold3",
  "SM-F936": "Galaxy Z Fold4",
  "SM-F946": "Galaxy Z Fold5",
  // Google Pixel
  "Pixel 6": "Pixel 6",
  "Pixel 7": "Pixel 7",
  "Pixel 8": "Pixel 8",
  "Pixel 9": "Pixel 9",
  // OnePlus
  CPH2449: "OnePlus 11",
  CPH2581: "OnePlus 12",
  CPH2609: "OnePlus Nord 3",
  // Xiaomi / Redmi / POCO — Xiaomi's own codes (e.g. "2201117TG") are
  // rarely human-readable at all, so only the few widely-known ones
  // are listed; most Xiaomi devices will just show the raw code.
  "2201117TG": "Redmi Note 11",
  "2201117SG": "Redmi Note 11",
  "23021RAAEG": "Redmi Note 12",
  "22111317I": "Redmi Note 12 Pro",
  "2210132G": "POCO X5",
  // OPPO
  CPH2477: "OPPO Reno8",
  CPH2413: "OPPO A96",
  CPH2481: "OPPO A78",
  // vivo
  V2145: "vivo Y33s",
  V2219: "vivo Y36",
  V2247: "vivo V29",
  // realme
  RMX3630: "realme 10",
  RMX3710: "realme 11",
  RMX3785: "realme 12",
};
function androidModelName(code) {
  if (!code) return null;
  const upper = code.toUpperCase();
  // Longest-prefix-first so e.g. "SM-A546" (Galaxy A54) doesn't get
  // shadowed by a shorter, coincidentally-matching key.
  const prefixes = Object.keys(ANDROID_MODEL_MAP).sort(
    (a, b) => b.length - a.length,
  );
  for (const prefix of prefixes) {
    if (upper.startsWith(prefix.toUpperCase()))
      return ANDROID_MODEL_MAP[prefix];
  }
  return null;
}
function LoginActivityPage({
  role,
  currentAdmin,
  currentEmp,
  isSuperAdmin,
  activeSessionId,
}) {
  const { t } = useLang();
  const selfType = role === "admin" ? "admin" : "employee";
  const selfId = role === "admin" ? currentAdmin?.id : currentEmp?.id;
  // "mine" | "all" — the "all" scope is only ever offered to superadmins;
  // everyone else is hard-locked to their own history below.
  const [scope, setScope] = useState("mine");
  const [logs, setLogs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [q, setQ] = useState("");
  const [revokingId, setRevokingId] = useState(null);
  const viewingAll = isSuperAdmin && scope === "all";

  const load = useCallback(async () => {
    setLoading(true);
    let query = supabase
      .from("login_activity")
      .select("*")
      .order("created_at", { ascending: false })
      .limit(300);
    if (!(isSuperAdmin && scope === "all")) {
      query = query.eq("actor_type", selfType).eq("actor_id", String(selfId));
    }
    const { data, error } = await query;
    if (error) {
      console.error("[login-activity] failed to load:", error.message);
      setLogs([]);
    } else {
      setLogs(data || []);
    }
    setLoading(false);
  }, [isSuperAdmin, scope, selfType, selfId]);

  useEffect(() => {
    load();
  }, [load]);

  // A "login" row counts as an active/live session when it has no
  // matching "logout" row (same session_id) among what we've loaded,
  // and hasn't been revoked yet — that's exactly what a "revoke"
  // action should be offered on.
  const loggedOutSessionIds = useMemo(() => {
    const s = new Set();
    logs.forEach((l) => {
      if (l.action === "logout" && l.session_id) s.add(l.session_id);
    });
    return s;
  }, [logs]);
  const isActiveSession = useCallback(
    (l) =>
      l.action === "login" &&
      !!l.session_id &&
      !l.revoked_at &&
      !loggedOutSessionIds.has(l.session_id),
    [loggedOutSessionIds],
  );

  // Drive ConfirmDialog / AlertDialog instead of window.confirm/alert —
  // confirmDialog holds {message, danger, onConfirm}; alertDialog holds
  // {message}. Rendered once at the bottom of this component's JSX.
  const [confirmDialog, setConfirmDialog] = useState(null);
  const [alertDialog, setAlertDialog] = useState(null);

  const doRevokeSession = useCallback(async (row) => {
    setRevokingId(row.id);
    const { error } = await supabase
      .from("login_activity")
      .update({ revoked_at: new Date().toISOString() })
      .eq("id", row.id);
    if (error) {
      console.error("[login-activity] revoke failed:", error.message);
    } else {
      setLogs((prev) =>
        prev.map((l) =>
          l.id === row.id ? { ...l, revoked_at: new Date().toISOString() } : l,
        ),
      );
    }
    setRevokingId(null);
  }, []);
  const revokeSession = useCallback(
    (row) => {
      setConfirmDialog({
        message: t.loginAct.revokeConfirm,
        danger: true,
        confirmLabel: t.loginAct.revoke,
        onConfirm: () => doRevokeSession(row),
      });
    },
    [t, doRevokeSession],
  );

  const [deletingId, setDeletingId] = useState(null);
  const doDeleteEntry = useCallback(
    async (row) => {
      setDeletingId(row.id);
      // .select() makes Postgres return the rows it actually deleted.
      // Without it, an RLS policy can silently block the delete (0 rows
      // affected) while Supabase still reports no error — which would
      // make the UI remove the row locally even though it's still in
      // the database. Checking the returned rows catches that case.
      const { data, error } = await supabase
        .from("login_activity")
        .delete()
        .eq("id", row.id)
        .select("id");
      if (error) {
        console.error("[login-activity] delete failed:", error.message);
        setAlertDialog({ message: t.loginAct.deleteBlocked });
      } else if (!data || data.length === 0) {
        console.error(
          "[login-activity] delete affected 0 rows — likely blocked by an RLS policy",
        );
        setAlertDialog({ message: t.loginAct.deleteBlocked });
      } else {
        setLogs((prev) => prev.filter((l) => l.id !== row.id));
      }
      setDeletingId(null);
    },
    [t],
  );
  const deleteEntry = useCallback(
    (row) => {
      setConfirmDialog({
        message: t.loginAct.deleteConfirm,
        danger: true,
        onConfirm: () => doDeleteEntry(row),
      });
    },
    [t, doDeleteEntry],
  );

  // Bulk cleanup — scoped to whatever's currently in view (own history,
  // or everyone's when a superadmin has switched to "all"), so people
  // can't accidentally wipe records outside what they can already see.
  const [clearing, setClearing] = useState(false);
  const doClearLogs = useCallback(
    async (olderThanDays) => {
      setClearing(true);
      let query = supabase.from("login_activity").delete();
      if (!(isSuperAdmin && scope === "all")) {
        query = query.eq("actor_type", selfType).eq("actor_id", String(selfId));
      }
      if (olderThanDays != null) {
        const cutoff = new Date(
          Date.now() - olderThanDays * 24 * 60 * 60 * 1000,
        ).toISOString();
        query = query.lt("created_at", cutoff);
      } else {
        // Supabase requires an explicit filter for delete — .gt on the
        // primary key matches every row without needing a date cutoff.
        query = query.gt("id", 0);
      }
      // Same RLS check as deleteEntry: confirm rows were actually
      // deleted rather than trusting the absence of an error.
      const { data, error } = await query.select("id");
      if (error) {
        console.error("[login-activity] clear failed:", error.message);
        setAlertDialog({ message: t.loginAct.deleteBlocked });
      } else if (!data || data.length === 0) {
        console.error(
          "[login-activity] clear affected 0 rows — likely blocked by an RLS policy",
        );
        setAlertDialog({ message: t.loginAct.deleteBlocked });
      } else {
        load();
      }
      setClearing(false);
    },
    [isSuperAdmin, scope, selfType, selfId, load, t],
  );
  const clearLogs = useCallback(
    (olderThanDays) => {
      setConfirmDialog({
        message:
          olderThanDays == null
            ? t.loginAct.clearAllConfirm
            : t.loginAct.clearOldConfirm,
        danger: true,
        onConfirm: () => doClearLogs(olderThanDays),
      });
    },
    [t, doClearLogs],
  );

  const filtered = useMemo(() => {
    const query = q.trim().toLowerCase();
    if (!query) return logs;
    return logs.filter((l) => {
      const hay = [l.actor_name, l.os, l.browser, l.device_type, l.action]
        .filter(Boolean)
        .join(" ")
        .toLowerCase();
      return hay.includes(query);
    });
  }, [logs, q]);
  const pg = usePagination(filtered, 25);

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
        <div>
          <h2
            style={{
              fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
              fontSize: 18,
              fontWeight: 600,
              color: T.ink,
              margin: 0,
            }}
          >
            {t.loginAct.title}
          </h2>
          <p style={{ fontSize: 12.5, color: T.muted, margin: "4px 0 0" }}>
            {viewingAll ? t.loginAct.subtitleAll : t.loginAct.subtitle}
          </p>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <div style={{ position: "relative" }}>
            <Search
              size={14}
              color={T.muted}
              style={{
                position: "absolute",
                left: 10,
                top: "50%",
                transform: "translateY(-50%)",
              }}
            />
            <Input
              value={q}
              onChange={(e) => setQ(e.target.value)}
              placeholder={t.search}
              style={{ paddingLeft: 30, minWidth: 200 }}
            />
          </div>
          {isSuperAdmin && (
            <Select value={scope} onChange={(e) => setScope(e.target.value)}>
              <option value="mine">{t.loginAct.scopeMine}</option>
              <option value="all">{t.loginAct.scopeAll}</option>
            </Select>
          )}
          <Button variant="ghost" size="sm" onClick={load}>
            <RefreshCw size={14} /> {t.loginAct.refresh}
          </Button>
          {isSuperAdmin && (
            <>
              <Button
                variant="ghost"
                size="sm"
                disabled={clearing}
                onClick={() => clearLogs(30)}
                style={{ color: T.rose }}
              >
                <Trash2 size={14} /> {t.loginAct.clearOld}
              </Button>
              <Button
                variant="ghost"
                size="sm"
                disabled={clearing}
                onClick={() => clearLogs(null)}
                style={{ color: T.rose }}
              >
                <Trash2 size={14} /> {t.loginAct.clearAll}
              </Button>
            </>
          )}
        </div>
      </div>

      {loading && (
        <Card
          style={{
            textAlign: "center",
            padding: "32px 0",
            color: T.muted,
            fontSize: 13,
          }}
        >
          <Loader2
            size={20}
            color={T.forest}
            style={{ animation: "spin 1s linear infinite" }}
          />
          <div style={{ marginTop: 8 }}>{t.loginAct.loading}</div>
        </Card>
      )}

      {!loading && filtered.length === 0 && (
        <Card
          style={{
            textAlign: "center",
            padding: "32px 0",
            color: T.muted,
            fontSize: 13,
          }}
        >
          {t.loginAct.noLogs}
        </Card>
      )}

      {!loading && filtered.length > 0 && (
        <Card style={{ overflowX: "auto", padding: 0 }}>
          <table className="wf-table">
            <thead>
              <tr>
                <th>{t.loginAct.time}</th>
                {viewingAll && <th>{t.loginAct.actor}</th>}
                <th>{t.loginAct.action}</th>
                <th>{t.loginAct.device}</th>
                <th>{t.loginAct.os}</th>
                <th>{t.loginAct.model}</th>
                <th>{t.loginAct.browser}</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {pg.pageItems.map((l) => (
                <tr key={l.id}>
                  <td
                    style={{
                      fontFamily: "'JetBrains Mono',monospace",
                      fontSize: 11.5,
                      color: T.muted,
                      whiteSpace: "nowrap",
                    }}
                  >
                    {fmtAuditTime(l.created_at)}
                  </td>
                  {viewingAll && (
                    <td>
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 8,
                        }}
                      >
                        <Avatar
                          name={l.actor_name || t.loginAct.unknownActor}
                          size={26}
                        />
                        <div
                          style={{
                            fontSize: 12.5,
                            fontWeight: 500,
                            color: T.ink,
                          }}
                        >
                          {l.actor_name || t.loginAct.unknownActor}
                        </div>
                      </div>
                    </td>
                  )}
                  <td>
                    <LoginActionBadge action={l.action} t={t} />
                  </td>
                  <td>
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                        fontSize: 12.5,
                        color: T.ink,
                      }}
                    >
                      <DeviceTypeIcon
                        deviceType={l.device_type}
                        color={T.muted}
                      />
                      {deviceTypeLabel(l.device_type, t)}
                    </div>
                  </td>
                  <td style={{ fontSize: 12.5, color: T.ink }}>
                    {l.os || t.loginAct.unknown}
                  </td>
                  <td style={{ fontSize: 12.5 }}>
                    {l.device_model ? (
                      <>
                        <div style={{ color: T.ink }}>
                          {androidModelName(l.device_model) || l.device_model}
                        </div>
                        {androidModelName(l.device_model) && (
                          <div
                            style={{
                              fontSize: 10.5,
                              color: T.muted,
                              fontFamily: "'JetBrains Mono',monospace",
                            }}
                          >
                            {l.device_model}
                          </div>
                        )}
                      </>
                    ) : (
                      <span style={{ color: T.muted }}>—</span>
                    )}
                  </td>
                  <td style={{ fontSize: 12.5, color: T.ink }}>
                    {l.browser || t.loginAct.unknown}
                  </td>
                  <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 8,
                        justifyContent: "flex-end",
                      }}
                    >
                      {l.session_id && l.session_id === activeSessionId ? (
                        <span style={{ fontSize: 10.5, color: T.muted }}>
                          {t.loginAct.currentDevice}
                        </span>
                      ) : isActiveSession(l) ? (
                        <>
                          <span
                            style={{
                              fontSize: 10.5,
                              fontWeight: 700,
                              color: T.forestText,
                              background: T.forestSoft,
                              padding: "2px 8px",
                              borderRadius: 999,
                            }}
                          >
                            {t.loginAct.active}
                          </span>
                          <Button
                            variant="ghost"
                            size="sm"
                            disabled={revokingId === l.id}
                            onClick={() => revokeSession(l)}
                            style={{ color: T.rose }}
                          >
                            <LogOut size={13} /> {t.loginAct.revoke}
                          </Button>
                        </>
                      ) : l.revoked_at ? (
                        <span style={{ fontSize: 10.5, color: T.muted }}>
                          {t.loginAct.revoked}
                        </span>
                      ) : null}
                      {isSuperAdmin && (
                        <button
                          onClick={() => deleteEntry(l)}
                          disabled={deletingId === l.id}
                          title={t.delete}
                          style={{
                            background: "none",
                            border: "none",
                            cursor: "pointer",
                            color: T.muted,
                            padding: 4,
                            display: "flex",
                            alignItems: "center",
                          }}
                        >
                          <Trash2 size={13} />
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{ padding: "0 16px 14px" }}>
            <Pagination {...pg} />
          </div>
        </Card>
      )}
      {confirmDialog && (
        <LoginActConfirmDialog
          message={confirmDialog.message}
          danger={confirmDialog.danger}
          confirmLabel={confirmDialog.confirmLabel}
          onCancel={() => setConfirmDialog(null)}
          onConfirm={() => {
            const fn = confirmDialog.onConfirm;
            setConfirmDialog(null);
            fn();
          }}
        />
      )}
      {alertDialog && (
        <LoginActAlertDialog
          message={alertDialog.message}
          onClose={() => setAlertDialog(null)}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Payroll
----------------------------------------------------------------*/
// Builds a CSV file from headers + row arrays and triggers a browser
// download. A UTF-8 BOM is prepended so Excel opens Khmer text
// correctly instead of mangling it as another encoding.
function exportCsv(filename, headers, rows) {
  const escapeCell = (v) => {
    const s = v === null || v === undefined ? "" : String(v);
    if (/[",\n\r]/.test(s)) {
      return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
  };
  const lines = [headers, ...rows].map((row) => row.map(escapeCell).join(","));
  const csv = "\uFEFF" + lines.join("\r\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Builds a formatted, bank/accountant-ready .xlsx payroll report (bold
// header band, merged title row, currency number formats, a totals row,
// and frozen header) and triggers a browser download. Uses ExcelJS
// (`npm install exceljs`) rather than the CSV path above because real
// payroll handoffs need actual formatting, not raw comma-separated text.
async function exportPayrollXlsx({
  filename,
  companyName,
  reportTitle,
  periodLabel,
  columns, // [{ header, width, currency? }]
  rows, // array of arrays, same order as columns
  totalLabel,
  totalValue,
  totalColIndex, // 0-based index into `columns` where the total goes
}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = companyName || "Workforce Suite";
  wb.created = new Date();
  const ws = wb.addWorksheet("Payroll", {
    views: [{ state: "frozen", ySplit: 4 }],
  });

  const colCount = columns.length;
  const lastColLetter = String.fromCharCode(64 + colCount); // supports up to 26 cols

  // Title band
  ws.mergeCells(`A1:${lastColLetter}1`);
  const titleCell = ws.getCell("A1");
  titleCell.value = companyName
    ? `${companyName} — ${reportTitle}`
    : reportTitle;
  titleCell.font = { bold: true, size: 14, color: { argb: "FF12203D" } };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(1).height = 26;

  ws.mergeCells(`A2:${lastColLetter}2`);
  const subCell = ws.getCell("A2");
  subCell.value = periodLabel;
  subCell.font = { italic: true, size: 11, color: { argb: "FF6B7280" } };
  subCell.alignment = { horizontal: "center" };

  ws.addRow([]); // spacer, row 3

  // Header row (row 4)
  const headerRow = ws.addRow(columns.map((c) => c.header));
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF12203D" },
    };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = {
      top: { style: "thin", color: { argb: "FF0B1730" } },
      bottom: { style: "thin", color: { argb: "FF0B1730" } },
    };
  });
  headerRow.height = 20;

  // Data rows
  rows.forEach((r, i) => {
    const row = ws.addRow(r);
    const zebra = i % 2 === 1;
    row.eachCell((cell, colNumber) => {
      const col = columns[colNumber - 1];
      if (col?.currency) cell.numFmt = '"$"#,##0.00';
      cell.alignment = { horizontal: col?.currency ? "right" : "left" };
      if (zebra) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF4F6F9" },
        };
      }
      cell.border = { bottom: { style: "hair", color: { argb: "FFE3E7EE" } } };
    });
  });

  // Totals row
  if (totalLabel !== undefined && totalColIndex !== undefined) {
    const totalRowValues = columns.map((_, i) =>
      i === 0 ? totalLabel : i === totalColIndex ? totalValue : "",
    );
    const totalRow = ws.addRow(totalRowValues);
    totalRow.eachCell((cell, colNumber) => {
      cell.font = { bold: true, color: { argb: "FF12203D" } };
      cell.border = { top: { style: "double", color: { argb: "FF12203D" } } };
      if (columns[colNumber - 1]?.currency) cell.numFmt = '"$"#,##0.00';
    });
  }

  columns.forEach((c, i) => {
    ws.getColumn(i + 1).width = c.width || 16;
  });

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function escapeHtml(s) {
  return String(s ?? "").replace(
    /[&<>"']/g,
    (c) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      })[c],
  );
}

// Builds a standalone, print-ready payslip document and opens the
// browser's print dialog on it. The user can choose "Save as PDF" as
// the destination — this avoids pulling in a PDF-generation library
// just for a single printable document.
function printPayslip({
  t,
  brandName,
  brandLogo,
  emp,
  mk,
  rows, // [{ label, value, tone: "neg"|"pos"|undefined }]
  net,
}) {
  const win = window.open("", "_blank", "width=480,height=720");
  if (!win) {
    alert("Please allow pop-ups to download the payslip.");
    return;
  }
  const rowsHtml = rows
    .map(
      (r) => `
      <div class="row${r.tone ? " " + r.tone : ""}">
        <span>${escapeHtml(r.label)}</span>
        <span class="mono">${r.tone === "neg" ? "-" : r.tone === "pos" ? "+" : ""}${escapeHtml(r.value)}</span>
      </div>`,
    )
    .join("");
  const generatedOn = new Date().toLocaleDateString();
  const html = `<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<title>${escapeHtml(t.pay.payslipTitle)} — ${escapeHtml(emp.name)} — ${escapeHtml(mk)}</title>
<style>
  * { box-sizing: border-box; }
  body {
    font-family: 'Noto Sans Khmer', 'Segoe UI', Arial, sans-serif;
    color: #16213a;
    margin: 0;
    padding: 32px;
    background: #fff;
  }
  .header {
    display: flex;
    align-items: center;
    gap: 10px;
    padding-bottom: 16px;
    border-bottom: 2px solid #16213a;
    margin-bottom: 20px;
  }
  .header img { width: 34px; height: 34px; border-radius: 8px; object-fit: cover; }
  .brand { font-size: 16px; font-weight: 700; }
  .title { font-size: 13px; color: #6b7280; margin-top: 2px; }
  .empbox {
    border: 1px solid #e5e7eb;
    border-radius: 10px;
    padding: 12px 14px;
    margin-bottom: 18px;
  }
  .empname { font-size: 14px; font-weight: 600; }
  .empsub { font-size: 11px; color: #6b7280; margin-top: 2px; }
  .row {
    display: flex;
    justify-content: space-between;
    font-size: 13px;
    padding: 6px 0;
  }
  .row.neg { color: #b3261e; }
  .row.pos { color: #1b6e4c; }
  .mono { font-family: 'Consolas', 'Courier New', monospace; }
  .net {
    display: flex;
    justify-content: space-between;
    align-items: center;
    background: #16213a;
    color: #fff;
    border-radius: 10px;
    padding: 14px 16px;
    margin-top: 16px;
  }
  .net .label { font-size: 13px; font-weight: 600; }
  .net .value { font-size: 19px; font-weight: 700; font-family: 'Consolas', 'Courier New', monospace; }
  .footer {
    margin-top: 28px;
    font-size: 10px;
    color: #9ca3af;
    text-align: center;
  }
  @media print {
    body { padding: 0 24px; }
  }
</style>
</head>
<body>
  <div class="header">
    ${brandLogo ? `<img src="${escapeHtml(brandLogo)}" />` : ""}
    <div>
      <div class="brand">${escapeHtml(brandName)}</div>
      <div class="title">${escapeHtml(t.pay.payslipTitle)} · ${escapeHtml(t.pay.payPeriod)}: ${escapeHtml(monthLabel(mk))}</div>
    </div>
  </div>
  <div class="empbox">
    <div class="empname">${escapeHtml(emp.name)}</div>
    <div class="empsub">${escapeHtml(emp.code)} · ${escapeHtml(emp.role || "")}</div>
  </div>
  <div>${rowsHtml}</div>
  <div class="net">
    <span class="label">${escapeHtml(t.pay.netSalary)}</span>
    <span class="value">${escapeHtml(net)}</span>
  </div>
  <div class="footer">${escapeHtml(t.pay.generatedOn)}: ${escapeHtml(generatedOn)}</div>
</body>
</html>`;
  win.document.open();
  win.document.write(html);
  win.document.close();
  win.focus();
  // Give the new document a moment to lay out (esp. any logo image)
  // before invoking print, so nothing is cut off blank.
  setTimeout(() => {
    win.print();
  }, 350);
}

// Builds a small, wearable-badge-sized printable ID card for one employee.
// Deliberately excludes sensitive fields (salary, PIN, phone, email) —
// this is meant to be worn/displayed, not a payroll or login document.
function printEmployeeBadge({
  t,
  brandName,
  brandLogo,
  emp,
  deptLabel,
  roleLabel,
  shiftText,
  officeText,
  statusLabel,
}) {
  const win = window.open("", "_blank", "width=420,height=620");
  if (!win) {
    alert("Please allow pop-ups to print the badge.");
    return;
  }
  const initials = getInitials(emp.name);
  const photoHtml = emp.photo
    ? `<img class="photo" src="${escapeHtml(emp.photo)}" />`
    : `<div class="photo initials">${escapeHtml(initials)}</div>`;
  const html = `<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<title>${escapeHtml(emp.name)} — ID Badge</title>
<style>
  * { box-sizing: border-box; }
  body {
    font-family: 'Noto Sans Khmer', 'Segoe UI', Arial, sans-serif;
    color: #16213a;
    margin: 0;
    padding: 28px;
    background: #f4f1ea;
    display: flex;
    justify-content: center;
  }
  .badge {
    width: 320px;
    background: #fff;
    border-radius: 18px;
    border: 1px solid #e5e7eb;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    overflow: hidden;
  }
  .badge-header {
    background: #16213a;
    color: #fff;
    padding: 14px 18px;
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .badge-header img { width: 22px; height: 22px; border-radius: 6px; object-fit: cover; }
  .badge-header .brand { font-size: 13px; font-weight: 700; }
  .badge-body { padding: 22px 18px; text-align: center; }
  .photo {
    width: 92px;
    height: 92px;
    border-radius: 50%;
    object-fit: cover;
    margin: 0 auto 14px;
    display: block;
    border: 3px solid #f4f1ea;
  }
  .photo.initials {
    display: flex;
    align-items: center;
    justify-content: center;
    background: #16213a;
    color: #fff;
    font-size: 30px;
    font-weight: 700;
  }
  .name { font-size: 17px; font-weight: 700; }
  .code {
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 12px;
    color: #6b7280;
    margin-top: 2px;
  }
  .status {
    display: inline-block;
    margin-top: 10px;
    font-size: 11px;
    font-weight: 700;
    color: #1b6e4c;
    background: #e4f3ea;
    padding: 4px 10px;
    border-radius: 999px;
  }
  .divider { border-top: 1px dashed #e5e7eb; margin: 16px 0; }
  .info { text-align: left; font-size: 12px; display: flex; flex-direction: column; gap: 7px; }
  .info .line { display: flex; gap: 6px; }
  .info .label { color: #9ca3af; min-width: 78px; }
  .info .value { font-weight: 600; }
  @media print {
    body { background: #fff; padding: 0; }
  }
</style>
</head>
<body>
  <div class="badge">
    <div class="badge-header">
      ${brandLogo ? `<img src="${escapeHtml(brandLogo)}" />` : ""}
      <span class="brand">${escapeHtml(brandName)}</span>
    </div>
    <div class="badge-body">
      ${photoHtml}
      <div class="name">${escapeHtml(emp.name)}</div>
      <div class="code">${escapeHtml(emp.code)}</div>
      <div class="status">${escapeHtml(statusLabel)}</div>
      <div class="divider"></div>
      <div class="info">
        <div class="line"><span class="label">${escapeHtml(t.emps.dept)}</span><span class="value">${escapeHtml(deptLabel)}</span></div>
        <div class="line"><span class="label">${escapeHtml(t.emps.role)}</span><span class="value">${escapeHtml(roleLabel)}</span></div>
        <div class="line"><span class="label">${escapeHtml(t.emps.shift)}</span><span class="value">${escapeHtml(shiftText)}</span></div>
        ${officeText ? `<div class="line"><span class="label">&#127974;</span><span class="value">${escapeHtml(officeText)}</span></div>` : ""}
      </div>
    </div>
  </div>
</body>
</html>`;
  win.document.open();
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(() => {
    win.print();
  }, 350);
}

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
  const { branding } = useBranding();
  const {
    absentDays,
    leaveDays,
    dailyRate,
    absenceDeduction,
    tax,
    insurance,
    taxRate,
    insuranceRate,
    deductionApplies,
    usesCustomRate,
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
  const handleDownload = () => {
    const rows = [
      { label: t.pay.baseSalary, value: fmtMoney(emp.salary) },
      ...(absentDays > 0
        ? [
            {
              label: `${t.pay.absentDed} (${absentDays} × ${fmtMoney(dailyRate)})`,
              value: fmtMoney(absenceDeduction),
              tone: "neg",
            },
          ]
        : []),
      ...(leaveDays > 0
        ? [{ label: t.lv.approved, value: `${leaveDays} ថ្ងៃ` }]
        : []),
      ...(deductionApplies
        ? [
            {
              label: `${t.pay.taxLabel} (${taxRate}%)`,
              value: fmtMoney(tax),
              tone: "neg",
            },
            {
              label: `${t.pay.insuranceLabel} (${insuranceRate}%)`,
              value: fmtMoney(insurance),
              tone: "neg",
            },
          ]
        : []),
      ...(otHours > 0
        ? [
            {
              label: `${t.pay.otPay} (${otHours} ${t.ot.hoursShort})`,
              value: fmtMoney(otPay),
              tone: "pos",
            },
          ]
        : []),
    ];
    printPayslip({
      t,
      brandName: branding?.name?.trim() || t.appName,
      brandLogo: branding?.logo || null,
      emp,
      mk,
      rows,
      net: fmtMoney(net),
    });
  };
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
            fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
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
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ fontSize: 14, fontWeight: 600, color: T.ink }}>
            {emp.name}
          </span>
          {usesCustomRate && (
            <span
              style={{
                fontSize: 9.5,
                fontWeight: 700,
                color: T.goldText,
                background: T.goldSoft,
                padding: "2px 6px",
                borderRadius: 6,
                whiteSpace: "nowrap",
              }}
            >
              {t.pay.customRateBadge}
            </span>
          )}
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
        {deductionApplies && (
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
        )}
        {deductionApplies && (
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
        )}
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
      <Button
        variant="ghost"
        style={{ width: "100%", marginTop: 12, justifyContent: "center" }}
        onClick={handleDownload}
      >
        <Download size={16} /> {t.pay.downloadPdf}
      </Button>
    </Modal>
  );
}

// Fetches attendance for exactly one "YYYY-MM" month straight from
// Supabase (paginated the same way as useSupabaseArray, so it's correct
// no matter how many rows the month has), without touching the app's
// windowed `attendance` state. Used by the historical payroll report to
// look at months older than the live 6-month window kept in memory.
async function fetchAttendanceForMonth(mk) {
  const [y, m] = mk.split("-").map(Number);
  const start = `${mk}-01`;
  const lastDay = new Date(y, m, 0).getDate();
  const end = `${mk}-${String(lastDay).padStart(2, "0")}`;
  const PAGE_SIZE = 1000;
  let all = [];
  let offset = 0;
  for (;;) {
    const { data, error } = await supabase
      .from("attendance")
      .select("*")
      .gte("date", start)
      .lte("date", end)
      .range(offset, offset + PAGE_SIZE - 1);
    if (error) throw error;
    all = all.concat(data || []);
    if (!data || data.length < PAGE_SIZE) break;
    offset += PAGE_SIZE;
  }
  return all.map((r) => ({
    id: r.id,
    employeeId: r.employee_id,
    date: r.date,
    checkIn: r.check_in,
    checkOut: r.check_out,
    status: r.status,
    checkInLoc: r.check_in_loc,
    checkOutLoc: r.check_out_loc,
  }));
}
function HistoricalPayrollModal({
  onClose,
  employees,
  overtimeRequests,
  otPolicy,
  payrollPolicy,
  branding,
}) {
  const { t, lang } = useLang();
  const [mk, setMk] = useState(() => {
    // Default to just before the live 6-month window so the picker
    // opens on a month that actually needs this on-demand fetch.
    const d = new Date();
    d.setMonth(d.getMonth() - 7);
    return monthKey(d);
  });
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [histAttendance, setHistAttendance] = useState(null); // null = not loaded yet
  const [xlsxExporting, setXlsxExporting] = useState(false);

  const activeEmployees = employees.filter((e) => e.status === "active");

  const load = async () => {
    setLoading(true);
    setError(null);
    setHistAttendance(null);
    try {
      const rows = await fetchAttendanceForMonth(mk);
      setHistAttendance(rows);
    } catch (err) {
      console.error("[historical payroll] load failed:", err);
      setError(t.pay.historicalError);
    } finally {
      setLoading(false);
    }
  };

  const rows = histAttendance
    ? activeEmployees.map((e) => ({
        emp: e,
        p: computePayroll(
          e,
          histAttendance,
          mk,
          overtimeRequests,
          otPolicy,
          payrollPolicy,
        ),
      }))
    : [];
  const totalNet = rows.reduce((sum, r) => sum + r.p.net, 0);

  return (
    <Modal title={t.pay.historicalTitle} onClose={onClose} width={760}>
      <p style={{ fontSize: 12, color: T.muted, marginBottom: 14 }}>
        {t.pay.historicalDesc}
      </p>
      <div
        style={{
          display: "flex",
          gap: 10,
          alignItems: "flex-end",
          marginBottom: 16,
          flexWrap: "wrap",
        }}
      >
        <Field label={t.pay.historicalPick}>
          <Input
            type="month"
            value={mk}
            max={monthKey()}
            onChange={(e) => e.target.value && setMk(e.target.value)}
          />
        </Field>
        <Button onClick={load} disabled={loading} style={{ marginBottom: 14 }}>
          {loading ? t.pay.historicalLoading : t.pay.historicalLoad}
        </Button>
      </div>

      {error && (
        <p style={{ color: T.rose, fontSize: 13, marginBottom: 12 }}>{error}</p>
      )}

      {loading && (
        <p
          style={{
            fontSize: 13,
            color: T.muted,
            textAlign: "center",
            padding: "24px 0",
          }}
        >
          {t.pay.historicalLoading}
        </p>
      )}

      {!loading && histAttendance && !error && (
        <>
          {histAttendance.length === 0 && (
            <p
              style={{
                fontSize: 13,
                color: T.muted,
                textAlign: "center",
                padding: "16px 0",
              }}
            >
              {t.pay.historicalEmpty}
            </p>
          )}
          <div style={{ maxHeight: 360, overflow: "auto" }}>
            <table className="wf-table">
              <thead>
                <tr>
                  <th>{t.employee}</th>
                  <th>{t.pay.baseSalary}</th>
                  <th>{t.pay.otPay}</th>
                  <th>{t.pay.taxLabel}</th>
                  <th>{t.pay.insuranceLabel}</th>
                  <th>{t.pay.netSalary}</th>
                </tr>
              </thead>
              <tbody>
                {rows.map(({ emp, p }) => (
                  <tr key={emp.id}>
                    <td>{emp.name}</td>
                    <td>{fmtMoney(emp.salary)}</td>
                    <td>{fmtMoney(p.otPay)}</td>
                    <td>{fmtMoney(p.tax)}</td>
                    <td>{fmtMoney(p.insurance)}</td>
                    <td style={{ fontWeight: 600 }}>{fmtMoney(p.net)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginTop: 16,
            }}
          >
            <span style={{ fontWeight: 600, color: T.ink }}>
              {t.pay.totalPaid}: {fmtMoney(totalNet)}
            </span>
            <Button
              variant="ghost"
              disabled={xlsxExporting || rows.length === 0}
              onClick={async () => {
                setXlsxExporting(true);
                try {
                  await exportPayrollXlsx({
                    filename: `payroll-${mk}.xlsx`,
                    companyName: branding?.name,
                    reportTitle:
                      lang === "en" ? "Payroll Report" : "របាយការណ៍ប្រាក់ខែ",
                    periodLabel: monthLabel(mk),
                    columns: [
                      { header: t.employee, width: 24 },
                      { header: t.pay.baseSalary, width: 16, currency: true },
                      { header: t.pay.otPay, width: 14, currency: true },
                      { header: t.pay.taxLabel, width: 14, currency: true },
                      {
                        header: t.pay.insuranceLabel,
                        width: 14,
                        currency: true,
                      },
                      { header: t.pay.netSalary, width: 16, currency: true },
                    ],
                    rows: rows.map(({ emp, p }) => [
                      emp.name,
                      emp.salary,
                      p.otPay,
                      p.tax,
                      p.insurance,
                      p.net,
                    ]),
                    totalLabel: lang === "en" ? "TOTAL" : "សរុប",
                    totalValue: totalNet,
                    totalColIndex: 5,
                  });
                } finally {
                  setXlsxExporting(false);
                }
              }}
            >
              <FileSpreadsheet size={15} />{" "}
              {xlsxExporting ? "…" : t.exportExcel}
            </Button>
          </div>
        </>
      )}
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
  const { branding } = useBranding();
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
  const [xlsxExporting, setXlsxExporting] = useState(false);
  const [showHistorical, setShowHistorical] = useState(false);
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
        {role === "admin" && (
          <div style={{ display: "flex", gap: 8, marginLeft: "auto" }}>
            <Button
              variant="ghost"
              onClick={() =>
                exportCsv(
                  `payroll-${mk}.csv`,
                  [
                    t.employee,
                    "Code",
                    t.pay.baseSalary,
                    t.pay.otPay,
                    t.pay.taxLabel,
                    t.pay.insuranceLabel,
                    t.pay.netSalary,
                    t.status,
                  ],
                  list.map((e) => {
                    const paid = !!payrollPaid[`${e.id}-${mk}`];
                    const { net, otPay, tax, insurance } = computePayroll(
                      e,
                      attendance,
                      mk,
                      overtimeRequests,
                      otPolicy,
                      payrollPolicy,
                    );
                    return [
                      e.name,
                      e.code,
                      e.salary,
                      otPay,
                      tax,
                      insurance,
                      net,
                      paid ? t.pay.paid : t.pay.unpaid,
                    ];
                  }),
                )
              }
            >
              <Download size={15} /> {t.exportCsv}
            </Button>
            <Button
              variant="ghost"
              disabled={xlsxExporting}
              onClick={async () => {
                setXlsxExporting(true);
                try {
                  const dataRows = list.map((e) => {
                    const paid = !!payrollPaid[`${e.id}-${mk}`];
                    const { net, otPay, tax, insurance } = computePayroll(
                      e,
                      attendance,
                      mk,
                      overtimeRequests,
                      otPolicy,
                      payrollPolicy,
                    );
                    return [
                      e.name,
                      e.code,
                      e.salary,
                      otPay,
                      tax,
                      insurance,
                      net,
                      paid ? t.pay.paid : t.pay.unpaid,
                    ];
                  });
                  await exportPayrollXlsx({
                    filename: `payroll-${mk}.xlsx`,
                    companyName: branding?.name,
                    reportTitle:
                      lang === "en" ? "Payroll Report" : "របាយការណ៍ប្រាក់ខែ",
                    periodLabel: monthLabel(mk),
                    columns: [
                      { header: t.employee, width: 24 },
                      { header: "Code", width: 12 },
                      { header: t.pay.baseSalary, width: 16, currency: true },
                      { header: t.pay.otPay, width: 14, currency: true },
                      { header: t.pay.taxLabel, width: 14, currency: true },
                      {
                        header: t.pay.insuranceLabel,
                        width: 14,
                        currency: true,
                      },
                      { header: t.pay.netSalary, width: 16, currency: true },
                      { header: t.status, width: 14 },
                    ],
                    rows: dataRows,
                    totalLabel: lang === "en" ? "TOTAL" : "សរុប",
                    totalValue: totalNet,
                    totalColIndex: 6,
                  });
                } finally {
                  setXlsxExporting(false);
                }
              }}
            >
              <FileSpreadsheet size={15} />{" "}
              {xlsxExporting ? "…" : t.exportExcel}
            </Button>
            <Button variant="ghost" onClick={() => setShowHistorical(true)}>
              <History size={15} /> {t.pay.historicalBtn}
            </Button>
          </div>
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
              const { net, absentDays, otHours, usesCustomRate } =
                computePayroll(
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
                            display: "flex",
                            alignItems: "center",
                            gap: 6,
                          }}
                        >
                          <span
                            style={{
                              fontWeight: 500,
                              color: T.ink,
                              fontSize: 13,
                            }}
                          >
                            {e.name}
                          </span>
                          {usesCustomRate && (
                            <span
                              style={{
                                fontSize: 9.5,
                                fontWeight: 700,
                                color: T.goldText,
                                background: T.goldSoft,
                                padding: "2px 6px",
                                borderRadius: 6,
                                whiteSpace: "nowrap",
                              }}
                            >
                              {t.pay.customRateBadge}
                            </span>
                          )}
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
      {showHistorical && (
        <HistoricalPayrollModal
          employees={employees}
          overtimeRequests={overtimeRequests}
          otPolicy={otPolicy}
          payrollPolicy={payrollPolicy}
          branding={branding}
          onClose={() => setShowHistorical(false)}
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
    { id: "analytics", label: n.analytics, icon: BarChart3 },
    { id: "announcements", label: n.announcements, icon: Megaphone },
    { id: "employees", label: n.employees, icon: Users },
    { id: "departments", label: n.departments, icon: Building2 },
    { id: "shifts", label: n.shifts, icon: Watch },
    { id: "attendance", label: n.attendance, icon: Clock },
    { id: "holidays", label: n.holidays, icon: CalendarDays },
    {
      id: "leave",
      label: n.leave,
      icon: CalendarDays,
      permission: "approveRequests",
    },
    {
      id: "ot",
      label: n.overtime,
      icon: Timer,
      permission: "approveRequests",
    },
    { id: "payroll", label: n.payroll, icon: Wallet },
    { id: "review", label: n.performance, icon: Star },
    {
      id: "attcorr",
      label: n.attCorrection,
      icon: CalendarClock,
      permission: "approveRequests",
    },
    {
      id: "admins",
      label: n.admins,
      icon: ShieldCheck,
      superadminOnly: true,
    },
    {
      id: "rolePerms",
      label: n.rolePerms,
      icon: KeyRound,
      superadminOnly: true,
    },
    {
      id: "audits",
      label: n.audits,
      icon: History,
      permission: "viewAuditLog",
    },
    { id: "loginActivity", label: n.loginActivity, icon: Smartphone },
    { id: "settings", label: n.settings, icon: Settings2 },
  ];
}
function buildNavEmployee(n, enabledModules) {
  const enabled = enabledModules || DEFAULT_EMPLOYEE_MODULES;
  return [
    { id: "dashboard", label: n.dashboard, icon: LayoutDashboard },
    { id: "announcements", label: n.announcements, icon: Megaphone },
    { id: "attendance", label: n.myAttendance, icon: Clock },
    { id: "leave", label: n.myLeave, icon: CalendarDays },
    { id: "ot", label: n.myOvertime, icon: Timer },
    { id: "payroll", label: n.myPayroll, icon: Wallet },
    { id: "review", label: n.myPerformance, icon: Star },
    { id: "attcorr", label: n.myAttCorrection, icon: CalendarClock },
    { id: "documents", label: n.myDocuments, icon: FileText },
    { id: "loginActivity", label: n.loginActivity, icon: Smartphone },
    { id: "profile", label: n.myProfile, icon: UserCircle2 },
  ].filter((item) => item.id === "dashboard" || enabled[item.id] !== false);
}

// Compact 5-item tab bar shown on phones for the employee/staff role,
// mirroring the "home / activity / requests / wallet / more" pattern of
// consumer apps (Grab, ABA...) so staff get one-tap access to the things
// they touch daily. Less-frequent pages stay reachable via "More", which
// opens the existing full side menu rather than duplicating it.
function buildBottomNavEmployee(n, enabledModules) {
  const enabled = enabledModules || DEFAULT_EMPLOYEE_MODULES;
  return [
    { id: "dashboard", label: n.dashboard, icon: LayoutDashboard },
    { id: "attendance", label: n.myAttendance, icon: Clock },
    { id: "leave", label: n.myLeave, icon: CalendarDays },
    { id: "payroll", label: n.myPayroll, icon: Wallet },
  ].filter((item) => item.id === "dashboard" || enabled[item.id] !== false);
}

function AppInner() {
  useGlobalStyle();
  const { t, lang } = useLang();
  const [branding, setBranding, brandingReady] = useBrandingSettings();
  const brandDisplayName = branding.name?.trim() || t.appName;
  const { theme } = useTheme();
  // Shows a styled AlertDialog (instead of window.alert) when this
  // device's session gets force-signed-out via the Login Activity page.
  // Rendered in the logged-out branches below, since checkRevoked always
  // clears the session before setting this.
  const [revokedNotice, setRevokedNotice] = useState(false);

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
  // Who is currently signed in, kept in a ref (not state) so the
  // useSupabaseArray hooks below can read the *latest* actor at the
  // moment a change is saved without needing to be re-created whenever
  // the logged-in person changes.
  const actorRef = useRef(null);
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
      audit: true,
      actorRef,
    },
  );
  const [employees, setEmployees, eReady, employeesSaveError] =
    useSupabaseArray("employees", {
      fromDb: (r) => ({
        id: r.id,
        code: r.code,
        pin: r.pin,
        name: r.name,
        deptId: r.dept_id,
        shiftId: r.shift_id,
        officeId: r.office_id,
        weeklyOff: r.weekly_off || [],
        customDaysOff: r.custom_days_off || [],
        annualLeaveDays: r.annual_leave_days ?? DEFAULT_ANNUAL_LEAVE_DAYS,
        sickLeaveDays: r.sick_leave_days ?? DEFAULT_SICK_LEAVE_DAYS,
        role: r.role,
        email: r.email,
        phone: r.phone,
        salary: r.salary,
        status: r.status,
        joined: r.joined,
        photo: r.photo,
        useCustomRate: !!r.use_custom_rate,
        customTaxRate: r.custom_tax_rate,
        customInsuranceRate: r.custom_insurance_rate,
      }),
      toDb: (r) => ({
        id: r.id,
        code: r.code,
        pin: r.pin,
        name: r.name,
        dept_id: r.deptId,
        shift_id: r.shiftId,
        office_id: r.officeId || null,
        weekly_off: r.weeklyOff || [],
        custom_days_off: r.customDaysOff || [],
        annual_leave_days: Number.isFinite(r.annualLeaveDays)
          ? r.annualLeaveDays
          : DEFAULT_ANNUAL_LEAVE_DAYS,
        sick_leave_days: Number.isFinite(r.sickLeaveDays)
          ? r.sickLeaveDays
          : DEFAULT_SICK_LEAVE_DAYS,
        role: r.role,
        email: r.email,
        phone: r.phone,
        salary: r.salary,
        status: r.status,
        joined: r.joined,
        photo: r.photo,
        use_custom_rate: !!r.useCustomRate,
        custom_tax_rate: r.useCustomRate ? Number(r.customTaxRate) || 0 : null,
        custom_insurance_rate: r.useCustomRate
          ? Number(r.customInsuranceRate) || 0
          : null,
      }),
      audit: true,
      actorRef,
      entityLabel: (r) => `${r.name || "?"} (${r.code || r.id})`,
    });
  const [shifts, setShifts, shReady] = useSupabaseArray("shifts", {
    audit: true,
    actorRef,
  });
  const [attendance, setAttendance, aReady] = useSupabaseArray("attendance", {
    // Attendance grows without bound (every check-in/out, forever), so we
    // only keep a rolling 6-month window live in the app — enough for the
    // Payroll month-picker's usual back-history and the Analytics trend
    // (which only ever looks at the last 6 months anyway). This keeps the
    // payload bounded even at 1000+ staff instead of dragging in years of
    // punches on every load. Combined with the pagination fix in
    // useSupabaseArray above, this also protects against Supabase/
    // PostgREST's default 1000-row response cap.
    dateField: "date",
    daysBack: 180,
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
        decidedById: r.decided_by_id,
        decidedByName: r.decided_by_name,
        decidedByRole: r.decided_by_role,
        decisionReason: r.decision_reason,
        reviewedAt: r.reviewed_at,
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
        decided_by_id: r.decidedById,
        decided_by_name: r.decidedByName,
        decided_by_role: r.decidedByRole,
        decision_reason: r.decisionReason,
        reviewed_at: r.reviewedAt,
      }),
      audit: true,
      actorRef,
      entityLabel: (r) => `${r.type || "leave"} · ${r.employeeId || "?"}`,
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
      audit: true,
      actorRef,
      entityLabel: (r) => `OT ${r.hours ?? "?"}h · ${r.employeeId || "?"}`,
    },
  );
  const [performanceReviews, setPerformanceReviews, prReady] = useSupabaseArray(
    "performance_reviews",
    {
      fromDb: (r) => ({
        id: r.id,
        employeeId: r.employee_id,
        period: r.period,
        rating: r.rating,
        strengths: r.strengths,
        improvements: r.improvements,
        reviewedById: r.reviewed_by_id,
        reviewedByName: r.reviewed_by_name,
        createdAt: r.created_at,
      }),
      toDb: (r) => ({
        id: r.id,
        employee_id: r.employeeId,
        period: r.period,
        rating: r.rating,
        strengths: r.strengths,
        improvements: r.improvements,
        reviewed_by_id: r.reviewedById,
        reviewed_by_name: r.reviewedByName,
        created_at: r.createdAt,
      }),
      audit: true,
      actorRef,
      entityLabel: (r) => `${r.period || "?"} · ${r.employeeId || "?"}`,
    },
  );
  const [announcements, setAnnouncements, annReady] = useSupabaseArray(
    "announcements",
    {
      fromDb: (r) => ({
        id: r.id,
        title: r.title,
        body: r.body,
        createdById: r.created_by_id,
        createdByName: r.created_by_name,
        createdAt: r.created_at,
      }),
      toDb: (r) => ({
        id: r.id,
        title: r.title,
        body: r.body,
        created_by_id: r.createdById,
        created_by_name: r.createdByName,
        created_at: r.createdAt,
      }),
      audit: true,
      actorRef,
      entityLabel: (r) => r.title || "?",
    },
  );
  const [documents, setDocuments, docsReady] = useSupabaseArray(
    "employee_documents",
    {
      fromDb: (r) => ({
        id: r.id,
        employeeId: r.employee_id,
        category: r.category,
        fileName: r.file_name,
        mimeType: r.mime_type,
        dataUrl: r.data_url,
        uploadedByName: r.uploaded_by_name,
        createdAt: r.created_at,
      }),
      toDb: (r) => ({
        id: r.id,
        employee_id: r.employeeId,
        category: r.category,
        file_name: r.fileName,
        mime_type: r.mimeType,
        data_url: r.dataUrl,
        uploaded_by_name: r.uploadedByName,
        created_at: r.createdAt,
      }),
      audit: true,
      actorRef,
      entityLabel: (r) => r.fileName || "?",
    },
  );
  const [holidays, setHolidays, holReady] = useSupabaseArray("holidays", {
    fromDb: (r) => ({ id: r.id, date: r.date, name: r.name }),
    toDb: (r) => ({ id: r.id, date: r.date, name: r.name }),
    audit: true,
    actorRef,
  });
  const [attendanceCorrections, setAttendanceCorrections, acReady] =
    useSupabaseArray("attendance_corrections", {
      fromDb: (r) => ({
        id: r.id,
        employeeId: r.employee_id,
        date: r.date,
        requestedCheckIn: r.requested_check_in,
        requestedCheckOut: r.requested_check_out,
        reason: r.reason,
        status: r.status,
        decidedById: r.decided_by_id,
        decidedByName: r.decided_by_name,
        decisionReason: r.decision_reason,
        createdAt: r.created_at,
        reviewedAt: r.reviewed_at,
      }),
      toDb: (r) => ({
        id: r.id,
        employee_id: r.employeeId,
        date: r.date,
        requested_check_in: r.requestedCheckIn,
        requested_check_out: r.requestedCheckOut,
        reason: r.reason,
        status: r.status,
        decided_by_id: r.decidedById,
        decided_by_name: r.decidedByName,
        decision_reason: r.decisionReason,
        created_at: r.createdAt,
        reviewed_at: r.reviewedAt,
      }),
      audit: true,
      actorRef,
      entityLabel: (r) => `${r.date || "?"} · ${r.employeeId || "?"}`,
    });
  const [admins, setAdmins, adminsReady, adminsSaveError] = useSupabaseArray(
    "admins",
    {
      audit: true,
      actorRef,
      entityLabel: (r) => r.name || r.username || r.id,
    },
  );
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
    audit: true,
    actorRef,
  });
  const [otPolicy, setOtPolicy, otPolicyReady] = useOtPolicy();
  const [payrollPolicy, setPayrollPolicy, payrollPolicyReady] =
    usePayrollPolicy();
  // Superadmin-editable permission matrix: one row per rank (Officer,
  // Senior, Supervisor, Manager, Senior Manager, Admin), each holding
  // which modules that rank can access. `id` IS the rank name (e.g.
  // "officer") — role_permissions has no separate numeric id, the rank
  // name itself is the primary key, which lets this reuse the same
  // generic upsert/delete-by-id logic as every other table above.
  const [rolePermissions, setRolePermissions, rolePermsReady] =
    useSupabaseArray("role_permissions", {
      fromDb: (r) => ({
        id: r.id,
        manageDepartments: !!r.manage_departments,
        manageEmployees: !!r.manage_employees,
        approveRequests: !!r.approve_requests,
        managePayroll: !!r.manage_payroll,
        manageDocuments: !!r.manage_documents,
        manageAnnouncements: !!r.manage_announcements,
        manageSettings: !!r.manage_settings,
        viewAuditLog: !!r.view_audit_log,
        // Employee self-service module toggles — only meaningful on the
        // single row where id === EMPLOYEE_MODULES_ID, but mapped for
        // every row since the columns live on the same table. Coerced
        // with ?? true rather than !! so a brand-new column (still null
        // for every existing row right after the migration runs) reads
        // as "on" instead of silently turning every module off.
        announcements: r.emp_announcements ?? true,
        attendance: r.emp_attendance ?? true,
        leave: r.emp_leave ?? true,
        ot: r.emp_ot ?? true,
        payroll: r.emp_payroll ?? true,
        review: r.emp_review ?? true,
        attcorr: r.emp_attcorr ?? true,
        documents: r.emp_documents ?? true,
        loginActivity: r.emp_login_activity ?? true,
        profile: r.emp_profile ?? true,
      }),
      toDb: (r) => ({
        id: r.id,
        manage_departments: r.manageDepartments,
        manage_employees: r.manageEmployees,
        approve_requests: r.approveRequests,
        manage_payroll: r.managePayroll,
        manage_documents: r.manageDocuments,
        manage_announcements: r.manageAnnouncements,
        manage_settings: r.manageSettings,
        view_audit_log: r.viewAuditLog,
        emp_announcements: r.announcements,
        emp_attendance: r.attendance,
        emp_leave: r.leave,
        emp_ot: r.ot,
        emp_payroll: r.payroll,
        emp_review: r.review,
        emp_attcorr: r.attcorr,
        emp_documents: r.documents,
        emp_login_activity: r.loginActivity,
        emp_profile: r.profile,
      }),
      audit: true,
      actorRef,
    });
  // { officer: {...}, senior: {...}, ... } — the saved matrix keyed by
  // rank, falling back to DEFAULT_ROLE_PERMISSIONS for any rank without
  // a saved row yet (fresh install, or a rank added after go-live).
  const rolePermissionsMap = useMemo(() => {
    const map = { ...DEFAULT_ROLE_PERMISSIONS };
    rolePermissions.forEach((r) => {
      map[r.id] = { ...DEFAULT_ROLE_PERMISSIONS[r.id], ...r };
    });
    return map;
  }, [rolePermissions]);
  // Company-wide switchboard for the employee self-service portal —
  // saved as one row (id === EMPLOYEE_MODULES_ID) inside the same
  // rolePermissions array as the admin rank matrix. Falls back to
  // "everything on" until Superadmin has saved a custom set.
  const employeeModules = useMemo(() => {
    const saved = rolePermissions.find((r) => r.id === EMPLOYEE_MODULES_ID);
    return saved
      ? { ...DEFAULT_EMPLOYEE_MODULES, ...saved }
      : { ...DEFAULT_EMPLOYEE_MODULES };
  }, [rolePermissions]);
  const moduleEnabled = useCallback(
    (key) => employeeModules[key] !== false,
    [employeeModules],
  );
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
  // Id of this browser's current sign-in, generated fresh at login and
  // carried on its login_activity row — lets that specific device be
  // force-signed-out remotely from the Login Activity page (see the
  // revoke-session effect below).
  const [activeSessionId, setActiveSessionId] = useLocalStorage(
    K.ACTIVE_SESSION_ID,
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
  // can("managePayroll") etc. — the granular permission check pages
  // should use going forward instead of isSuperAdmin, so a rank
  // Superadmin has granted a module to can actually use it.
  const can = useCallback(
    (key) => canDo(currentAdmin, rolePermissionsMap, key),
    [currentAdmin, rolePermissionsMap],
  );
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
      ? buildNavAdmin(t.nav).filter(
          (n) =>
            (!n.superadminOnly || isSuperAdmin) &&
            (!n.permission || can(n.permission)),
        )
      : buildNavEmployee(t.nav, employeeModules);
  const bottomNav =
    role !== "admin" ? buildBottomNavEmployee(t.nav, employeeModules) : null;

  // Keep actorRef in sync with whoever is signed in right now, so every
  // useSupabaseArray hook above always audits changes under the correct
  // name/id — even though the hooks themselves were created before login.
  useEffect(() => {
    actorRef.current =
      role === "admin"
        ? currentAdmin
          ? { type: "admin", id: currentAdmin.id, name: currentAdmin.name }
          : null
        : currentEmp
          ? { type: "employee", id: currentEmp.id, name: currentEmp.name }
          : null;
  }, [role, currentAdmin, currentEmp]);

  // Periodically checks whether THIS device's session has been revoked
  // from the Login Activity page (e.g. a superadmin marking a device
  // they don't recognize). If so, force-sign this device out right
  // away — this is what makes "revoke" actually kick the device out,
  // rather than just deleting a log row.
  useEffect(() => {
    if (!loggedIn || !activeSessionId) return;
    let cancelled = false;
    const checkRevoked = async () => {
      const { data, error } = await supabase
        .from("login_activity")
        .select("revoked_at")
        .eq("session_id", activeSessionId)
        .eq("action", "login")
        .not("revoked_at", "is", null)
        .limit(1);
      if (cancelled || error || !data || data.length === 0) return;
      const actor =
        role === "admin"
          ? { type: "admin", id: currentAdmin?.id, name: currentAdmin?.name }
          : { type: "employee", id: currentEmp?.id, name: currentEmp?.name };
      writeAuditLog({
        actor,
        action: "logout",
        table: role === "admin" ? "admins" : "employees",
        entityId: actor.id,
        label: actor.name,
      });
      writeLoginActivity({
        actor,
        action: "logout",
        sessionId: activeSessionId,
      });
      setActiveSessionId(null);
      if (role === "admin") setSessionAdmin(null);
      else setSessionEmployee(null);
      setRevokedNotice(true);
    };
    checkRevoked();
    const interval = setInterval(checkRevoked, 20000);
    window.addEventListener("focus", checkRevoked);
    return () => {
      cancelled = true;
      clearInterval(interval);
      window.removeEventListener("focus", checkRevoked);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [loggedIn, activeSessionId, role, currentAdmin, currentEmp]);

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
            onLogin={(id) => {
              setSessionAdmin(id);
              const acct = admins.find((a) => a.id === id);
              writeAuditLog({
                actor: { type: "admin", id, name: acct?.name },
                action: "login",
                table: "admins",
                entityId: id,
                label: acct?.name,
              });
              const sid = newSessionId();
              setActiveSessionId(sid);
              writeLoginActivity({
                actor: { type: "admin", id, name: acct?.name },
                action: "login",
                sessionId: sid,
              });
            }}
            go={goPortal}
          />
          {revokedNotice && (
            <LoginActAlertDialog
              message={t.loginAct.revokedAlert}
              onClose={() => setRevokedNotice(false)}
            />
          )}
        </BrandingContext.Provider>
      );
    return (
      <BrandingContext.Provider value={{ branding, setBranding }}>
        <EmployeeLoginScreen
          employees={employees}
          onLogin={(id) => {
            setSessionEmployee(id);
            const emp = employees.find((e) => e.id === id);
            writeAuditLog({
              actor: { type: "employee", id, name: emp?.name },
              action: "login",
              table: "employees",
              entityId: id,
              label: emp?.name,
            });
            const sid = newSessionId();
            setActiveSessionId(sid);
            writeLoginActivity({
              actor: { type: "employee", id, name: emp?.name },
              action: "login",
              sessionId: sid,
            });
          }}
          go={goPortal}
        />
        {revokedNotice && (
          <LoginActAlertDialog
            message={t.loginAct.revokedAlert}
            onClose={() => setRevokedNotice(false)}
          />
        )}
      </BrandingContext.Provider>
    );
  }

  return (
    <BrandingContext.Provider value={{ branding, setBranding }}>
      <div
        className={`wf-root wf-app-enter ${theme === "dark" ? "wf-dark" : ""} ${
          bottomNav ? "wf-role-staff" : ""
        }`}
      >
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
                  fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
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
                gap: 3,
              }}
            >
              <div className="wf-nav-eyebrow">
                {lang === "km" ? "ម៉ឺនុយ" : "Menu"}
              </div>
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
                      ? adminRoleLabel(currentAdmin?.role, lang)
                      : currentEmp?.code}
                  </div>
                </div>
              </div>
              <button
                className="wf-nav-item"
                style={{ color: "#E3B7BE" }}
                onClick={() => {
                  if (role === "admin") {
                    writeAuditLog({
                      actor: {
                        type: "admin",
                        id: currentAdmin?.id,
                        name: currentAdmin?.name,
                      },
                      action: "logout",
                      table: "admins",
                      entityId: currentAdmin?.id,
                      label: currentAdmin?.name,
                    });
                    writeLoginActivity({
                      actor: {
                        type: "admin",
                        id: currentAdmin?.id,
                        name: currentAdmin?.name,
                      },
                      action: "logout",
                      sessionId: activeSessionId,
                    });
                    setActiveSessionId(null);
                    setSessionAdmin(null);
                  } else {
                    writeAuditLog({
                      actor: {
                        type: "employee",
                        id: currentEmp?.id,
                        name: currentEmp?.name,
                      },
                      action: "logout",
                      table: "employees",
                      entityId: currentEmp?.id,
                      label: currentEmp?.name,
                    });
                    writeLoginActivity({
                      actor: {
                        type: "employee",
                        id: currentEmp?.id,
                        name: currentEmp?.name,
                      },
                      action: "logout",
                      sessionId: activeSessionId,
                    });
                    setActiveSessionId(null);
                    setSessionEmployee(null);
                  }
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
            <div>
              <div
                style={{
                  fontFamily: "'JetBrains Mono',monospace",
                  fontSize: 10.5,
                  fontWeight: 600,
                  letterSpacing: ".08em",
                  textTransform: "uppercase",
                  color: T.muted,
                  marginBottom: 2,
                }}
              >
                {brandDisplayName}
              </div>
              <h1
                style={{
                  fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
                  fontWeight: 700,
                  color: T.ink,
                  fontSize: 16,
                  letterSpacing: "-.01em",
                }}
              >
                {nav.find((n) => n.id === page)?.label}
              </h1>
            </div>
            <div
              style={{
                marginLeft: "auto",
                display: "flex",
                alignItems: "center",
                gap: 8,
              }}
            >
              <div className="wf-role-badge">
                <HeaderClock />
              </div>
              {role !== "admin" && currentEmp && (
                <span
                  className="wf-role-badge"
                  style={{
                    fontSize: 11,
                    fontWeight: 700,
                    background: T.forestSoft,
                    color: T.forestText,
                    padding: "5px 10px",
                    borderRadius: 8,
                    whiteSpace: "nowrap",
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                    maxWidth: 90,
                  }}
                >
                  {t.employee}
                </span>
              )}
              {role === "admin" && (
                <span
                  className="wf-role-badge"
                  style={{
                    fontSize: 11,
                    fontWeight: 700,
                    background: "rgba(91,141,239,0.14)",
                    color: T.blue,
                    padding: "5px 10px",
                    borderRadius: 6,
                    whiteSpace: "nowrap",
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                    maxWidth: 90,
                  }}
                  title={adminRoleLabel(currentAdmin?.role, lang)}
                >
                  {adminRoleLabel(currentAdmin?.role, lang)}
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
                performanceReviews={performanceReviews}
                announcements={announcements}
                attendanceCorrections={attendanceCorrections}
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

          <main className={`wf-content ${bottomNav ? "wf-content-bnpad" : ""}`}>
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
            <div key={page} className="wf-page-enter">
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
              {page === "analytics" && role === "admin" && (
                <AnalyticsPage
                  employees={employees}
                  departments={departments}
                  attendance={attendance}
                  overtimeRequests={overtimeRequests}
                  otPolicy={otPolicy}
                  payrollPolicy={payrollPolicy}
                />
              )}
              {page === "announcements" &&
                (role === "admin" || moduleEnabled("announcements")) && (
                  <Announcements
                    role={role}
                    currentAdmin={currentAdmin}
                    announcements={announcements}
                    setAnnouncements={setAnnouncements}
                    isSuperAdmin={isSuperAdmin || can("manageAnnouncements")}
                  />
                )}
              {page === "employees" && role === "admin" && (
                <Employees
                  employees={employees}
                  departments={departments}
                  shifts={shifts}
                  offices={offices}
                  setEmployees={setEmployees}
                  isSuperAdmin={isSuperAdmin || can("manageEmployees")}
                  currentAdmin={currentAdmin}
                  documents={documents}
                  setDocuments={setDocuments}
                />
              )}
              {page === "departments" && role === "admin" && (
                <Departments
                  departments={departments}
                  setDepartments={setDepartments}
                  employees={employees}
                  isSuperAdmin={isSuperAdmin || can("manageDepartments")}
                />
              )}
              {page === "shifts" && role === "admin" && (
                <Shifts
                  shifts={shifts}
                  setShifts={setShifts}
                  employees={employees}
                  isSuperAdmin={isSuperAdmin || can("manageDepartments")}
                />
              )}
              {page === "attendance" &&
                (role === "admin" || moduleEnabled("attendance")) && (
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
                    holidays={holidays}
                  />
                )}
              {page === "holidays" && role === "admin" && (
                <Holidays
                  holidays={holidays}
                  setHolidays={setHolidays}
                  isSuperAdmin={isSuperAdmin || can("manageDepartments")}
                />
              )}
              {page === "leave" &&
                (role === "admin" || moduleEnabled("leave")) && (
                  <LeaveRequests
                    role={role}
                    currentAdmin={currentAdmin}
                    currentEmp={currentEmp}
                    employees={employees}
                    admins={admins}
                    leaveRequests={leaveRequests}
                    setLeaveRequests={setLeaveRequests}
                    attendance={attendance}
                    setAttendance={setAttendance}
                    isSuperAdmin={isSuperAdmin || can("approveRequests")}
                    canApprove={isSuperAdmin || can("approveRequests")}
                  />
                )}
              {page === "ot" && (role === "admin" || moduleEnabled("ot")) && (
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
                  isSuperAdmin={isSuperAdmin || can("approveRequests")}
                  canApprove={isSuperAdmin || can("approveRequests")}
                  holidays={holidays}
                />
              )}
              {page === "payroll" &&
                (role === "admin" || moduleEnabled("payroll")) && (
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
              {page === "review" &&
                (role === "admin" || moduleEnabled("review")) && (
                  <PerformanceReviews
                    role={role}
                    currentAdmin={currentAdmin}
                    currentEmp={currentEmp}
                    employees={employees}
                    performanceReviews={performanceReviews}
                    setPerformanceReviews={setPerformanceReviews}
                    isSuperAdmin={isSuperAdmin}
                  />
                )}
              {page === "documents" &&
                role !== "admin" &&
                currentEmp &&
                moduleEnabled("documents") && (
                  <MyDocuments currentEmp={currentEmp} documents={documents} />
                )}
              {page === "attcorr" &&
                (role === "admin" || moduleEnabled("attcorr")) && (
                  <AttendanceCorrections
                    role={role}
                    currentAdmin={currentAdmin}
                    currentEmp={currentEmp}
                    employees={employees}
                    admins={admins}
                    attendanceCorrections={attendanceCorrections}
                    setAttendanceCorrections={setAttendanceCorrections}
                    attendance={attendance}
                    setAttendance={setAttendance}
                    isSuperAdmin={isSuperAdmin || can("approveRequests")}
                    canApprove={isSuperAdmin || can("approveRequests")}
                  />
                )}
              {page === "admins" && role === "admin" && isSuperAdmin && (
                <AdminAccounts
                  admins={admins}
                  setAdmins={setAdmins}
                  currentAdminId={currentAdmin?.id}
                />
              )}
              {page === "rolePerms" && role === "admin" && isSuperAdmin && (
                <RolePermissionsPage
                  rolePermissions={rolePermissions}
                  setRolePermissions={setRolePermissions}
                />
              )}
              {page === "profile" &&
                role !== "admin" &&
                currentEmp &&
                moduleEnabled("profile") && (
                  <MyProfile
                    currentEmp={currentEmp}
                    employees={employees}
                    setEmployees={setEmployees}
                    departments={departments}
                    shifts={shifts}
                    saveError={employeesSaveError}
                  />
                )}
              {page === "settings" && role === "admin" && currentAdmin && (
                <AdminSettings
                  currentAdmin={currentAdmin}
                  admins={admins}
                  setAdmins={setAdmins}
                  isSuperAdmin={isSuperAdmin}
                  saveError={adminsSaveError}
                />
              )}
              {page === "audits" &&
                role === "admin" &&
                (isSuperAdmin || can("viewAuditLog")) && <AuditLogPage />}
              {page === "loginActivity" &&
                (role === "admin" || moduleEnabled("loginActivity")) && (
                  <LoginActivityPage
                    role={role}
                    currentAdmin={currentAdmin}
                    currentEmp={currentEmp}
                    isSuperAdmin={isSuperAdmin}
                    activeSessionId={activeSessionId}
                  />
                )}
            </div>
          </main>

          {bottomNav && (
            <nav
              className={`wf-bottomnav ${navOpen ? "wf-bottomnav-hidden" : ""}`}
            >
              {bottomNav.map((n) => (
                <button
                  key={n.id}
                  className={`wf-bottomnav-item ${page === n.id ? "active" : ""}`}
                  onClick={() => setPage(n.id)}
                >
                  <span className="wf-bnav-icon-wrap">
                    <n.icon size={19} />
                  </span>
                  <span>{n.label}</span>
                </button>
              ))}
              <button
                className={`wf-bottomnav-item ${navOpen ? "active" : ""}`}
                onClick={() => setNavOpen(true)}
              >
                <span className="wf-bnav-icon-wrap">
                  <Menu size={19} />
                </span>
                <span>{lang === "km" ? "ផ្សេងទៀត" : "More"}</span>
              </button>
            </nav>
          )}
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