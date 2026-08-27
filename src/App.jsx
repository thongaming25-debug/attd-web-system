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
  Eye,
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
  Volume2,
  Package,
  Repeat,
  Award,
  Send,
  GraduationCap,
  Briefcase,
  ListChecks,
  MessageCircle,
  Mic,
  MicOff,
  Phone,
  PhoneOff,
  Square,
} from "lucide-react";

/* ---------------------------------------------------------------
   Language / i18n  — add new keys here as needed
----------------------------------------------------------------*/
const LANG_RAW = {
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
      shiftSwap: "សំណើដូរវេន",
      messages: "សារ",
      assets: "ទ្រព្យសម្បត្តិក្រុមហ៊ុន",
      training: "ការបណ្តុះបណ្តាល & សញ្ញាបត្រ",
      docExpiry: "កិច្ចសន្យា/ឯកសារជិតផុតកំណត់",
      recruitment: "ការជ្រើសរើសបុគ្គលិក",
      onboarding: "ចាប់ផ្តើម/បញ្ចប់ការងារ",
      admins: "គណនីអ្នកគ្រប់គ្រង",
      myAttendance: "វត្តមានរបស់ខ្ញុំ",
      myLeave: "ច្បាប់ឈប់សម្រាករបស់ខ្ញុំ",
      myOvertime: "ការងារបន្ថែម (OT) របស់ខ្ញុំ",
      myPayroll: "ប្រាក់ខែរបស់ខ្ញុំ",
      myPerformance: "ការវាយតម្លៃការងាររបស់ខ្ញុំ",
      myTraining: "ការបណ្តុះបណ្តាល & សញ្ញាបត្ររបស់ខ្ញុំ",
      myDocuments: "ឯកសាររបស់ខ្ញុំ",
      myAttCorrection: "ស្នើសុំកែតម្រូវវត្តមាន",
      myShiftSwap: "ស្នើសុំដូរវេន",
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
    clear: "សម្អាត",
    today: "ថ្ងៃនេះ",
    yesterday: "ម្សិលមិញ",
    now: "ឥឡូវ",
    timeLabel: "ម៉ោង",
    selectDate: "ជ្រើសរើសកាលបរិច្ឆេទ",
    noResults: "មិនមានលទ្ធផលទេ",
    pagination: { of: "នៃ" },
    popupBlockedTitle: "កម្មវិធីរុករករារ Pop-up",
    popupBlockedPayslip:
      "សូមអនុញ្ញាត pop-up សម្រាប់គេហទំព័រនេះ ដើម្បីទាញយកបញ្ជីប្រាក់ខែ",
    popupBlockedBadge:
      "សូមអនុញ្ញាត pop-up សម្រាប់គេហទំព័រនេះ ដើម្បីបោះពុម្ពកាតសម្គាល់",
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
      workingNow: "កំពុងធ្វើការឥឡូវនេះ",
      workingNowSub: (n) => `${n} នាក់កំពុងបើកកម្មវិធីធ្វើការ`,
      noOneWorkingNow: "មិនទាន់មានបុគ្គលិកចូលធ្វើការនៅឡើយទេ",
      unassignedBranch: "មិនកំណត់សាខា",
      sinceLabel: "តាំងពី",
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
      newHiresTrend: "និន្នាការជួលបុគ្គលិកថ្មី (៦ ខែចុងក្រោយ)",
      newHiresTrendSub: "ចំនួនបុគ្គលិកចូលថ្មីប្រចាំខែ (គិតតាមថ្ងៃចូលធ្វើការ)",
      absenteeism: "អត្រាអវត្តមានតាមនាយកដ្ឋាន",
      absenteeismSub:
        'ភាគរយកំណត់ត្រា "អវត្តមាន" ធៀបនឹងកំណត់ត្រាវត្តមានសរុបប្រចាំខែនេះ',
      statActiveHeadcount: "បុគ្គលិកសកម្មសរុប",
      statAvgTenure: "អាយុកាលធ្វើការជាមធ្យម",
      statInactiveRate: "អត្រាអសកម្មបច្ចុប្បន្ន",
      statInactiveRateNote:
        "ចំណាំ៖ នេះជាភាគរយស្ថិតិបច្ចុប្បន្ន មិនមែនអត្រាចេញ-ចូលក្នុងកំឡុងពេលជាក់លាក់ទេ (ប្រព័ន្ធមិនទាន់កត់ត្រាកាលបរិច្ឆេទចាកចេញ)",
      tenureFormat: (y, m) => (y > 0 ? `${y} ឆ្នាំ ${m} ខែ` : `${m} ខែ`),
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
      namePlaceholder: "ឧ. ធនធានមនុស្ស",
      codePlaceholder: "ឧ. HR",
      confirmDelWithCount: (name, n) =>
        `តើអ្នកពិតជាចង់លុបនាយកដ្ឋាន "${name}" មែនទេ? (មានបុគ្គលិក ${n} នាក់)`,
    },
    assets: {
      title: "ទ្រព្យសម្បត្តិក្រុមហ៊ុន",
      addBtn: "បន្ថែមទ្រព្យសម្បត្តិ",
      editTitle: "កែសម្រួលទ្រព្យសម្បត្តិ",
      addTitle: "បន្ថែមទ្រព្យសម្បត្តិ",
      name: "ឈ្មោះទ្រព្យសម្បត្តិ",
      category: "ប្រភេទ",
      serial: "លេខស៊េរី / លេខកូដ",
      assignedTo: "កាន់កាប់ដោយ",
      unassigned: "មិនទាន់ចាត់តាំង",
      status: "ស្ថានភាព",
      statusAvailable: "នៅសល់",
      statusAssigned: "កំពុងប្រើប្រាស់",
      statusMaintenance: "កំពុងជួសជុល",
      statusRetired: "ឈប់ប្រើ",
      purchaseDate: "ថ្ងៃទិញ",
      notes: "កំណត់ចំណាំ",
      namePlaceholder: "ឧ. កុំព្យូទ័រយួរដៃ Dell",
      categoryPlaceholder: "ឧ. កុំព្យូទ័រ, យានយន្ត, ទូរស័ព្ទ",
      serialPlaceholder: "ឧ. SN-00123",
      searchPlaceholder: "ស្វែងរកទ្រព្យសម្បត្តិ...",
      allStatus: "ស្ថានភាពទាំងអស់",
      noAssets: "មិនទាន់មានទ្រព្យសម្បត្តិទេ",
      confirmDelWithName: (name) => `តើអ្នកប្រាកដទេថាចង់លុប "${name}"?`,
    },
    training: {
      title: "ការបណ្តុះបណ្តាល & សញ្ញាបត្រ",
      subtitle: "តាមដានវគ្គបណ្តុះបណ្តាល និងសញ្ញាបត្រនៃបុគ្គលិកគ្រប់រូប",
      addBtn: "បន្ថែមកំណត់ត្រា",
      editTitle: "កែសម្រួលកំណត់ត្រាបណ្តុះបណ្តាល",
      addTitle: "បន្ថែមកំណត់ត្រាបណ្តុះបណ្តាល",
      employee: "បុគ្គលិក",
      courseName: "ឈ្មោះវគ្គ/សញ្ញាបត្រ",
      provider: "អ្នកផ្តល់វគ្គ/ស្ថាប័ន",
      category: "ប្រភេទ",
      startDate: "ថ្ងៃចាប់ផ្តើម",
      completionDate: "ថ្ងៃបញ្ចប់",
      certExpiry: "ថ្ងៃផុតកំណត់សញ្ញាបត្រ (បើមាន)",
      status: "ស្ថានភាព",
      statusPlanned: "គ្រោងទុក",
      statusOngoing: "កំពុងរៀន",
      statusCompleted: "បញ្ចប់ហើយ",
      statusCancelled: "បានលុបចោល",
      notes: "កំណត់ចំណាំ",
      courseNamePlaceholder: "ឧ. វគ្គបណ្តុះបណ្តាលសុវត្ថិភាពការងារ",
      providerPlaceholder: "ឧ. ក្រសួងការងារ, NGO មួយ, Coursera",
      categoryPlaceholder: "ឧ. សុវត្ថិភាព, បច្ចេកទេស, ជំនាញទន់",
      searchPlaceholder: "ស្វែងរកតាមឈ្មោះបុគ្គលិក ឬឈ្មោះវគ្គ...",
      allStatus: "ស្ថានភាពទាំងអស់",
      allEmployees: "បុគ្គលិកទាំងអស់",
      noTrainings: "មិនទាន់មានកំណត់ត្រាបណ្តុះបណ្តាលទេ",
      confirmDelWithName: (name) => `តើអ្នកប្រាកដទេថាចង់លុប "${name}"?`,
      expired: "សញ្ញាបត្រផុតកំណត់ហើយ",
      expiringSoon: (n) => `ជិតផុតកំណត់ក្នុងរយៈ ${n} ថ្ងៃទៀត`,
      expiresOn: "ផុតកំណត់នៅ",
      myTitle: "ការបណ្តុះបណ្តាល & សញ្ញាបត្ររបស់ខ្ញុំ",
      noMyTrainings: "អ្នកមិនទាន់មានកំណត់ត្រាបណ្តុះបណ្តាលទេ",
    },
    contracts: {
      title: "កិច្ចសន្យា និងឯកសារជិតផុតកំណត់",
      subtitle:
        "តាមដានឯកសារ (កិច្ចសន្យា, អត្តសញ្ញាណប័ណ្ណ ។ល។) ដែលមានកាលបរិច្ឆេទផុតកំណត់ គ្រប់បុគ្គលិកទាំងអស់",
      searchPlaceholder: "ស្វែងរកតាមឈ្មោះបុគ្គលិក ឬឈ្មោះឯកសារ...",
      allStatus: "ស្ថានភាពទាំងអស់",
      statusExpired: "ផុតកំណត់ហើយ",
      statusSoon: "ជិតផុតកំណត់",
      statusValid: "នៅមានសុពលភាព",
      noItems: "មិនទាន់មានឯកសារកំណត់ថ្ងៃផុតកំណត់ទេ",
      noItemsHint:
        'ដើម្បីតាមដាននៅទីនេះ សូមបញ្ចូល "ថ្ងៃផុតកំណត់" ពេលផ្ទុកឯកសារនៅផ្នែក "បុគ្គលិក" → ឯកសារ',
      colEmployee: "បុគ្គលិក",
      colDocument: "ឈ្មោះឯកសារ",
      colCategory: "ប្រភេទ",
      colExpiry: "ថ្ងៃផុតកំណត់",
      colStatus: "ស្ថានភាព",
      daysLeft: (n) => `${n} ថ្ងៃទៀត`,
      daysAgo: (n) => `ផុតកំណត់ ${n} ថ្ងៃមុន`,
    },
    recruit: {
      title: "ការជ្រើសរើសបុគ្គលិក",
      tabJobs: "កាលានុវត្តភាពការងារ",
      tabCandidates: "បេក្ខជន",
      addJobBtn: "បង្ហោះកាលានុវត្តភាព",
      editJobTitle: "កែសម្រួលកាលានុវត្តភាព",
      addJobTitle: "បង្ហោះកាលានុវត្តភាពថ្មី",
      jobTitle: "តំណែងការងារ",
      jobTitlePlaceholder: "ឧ. គណនេយ្យករ",
      jobDept: "នាយកដ្ឋាន",
      jobType: "ប្រភេទការងារ",
      jobTypeFullTime: "ពេញម៉ោង",
      jobTypePartTime: "ក្រៅម៉ោង",
      jobTypeContract: "កិច្ចសន្យា",
      jobTypeInternship: "កម្មសិក្សា",
      jobStatus: "ស្ថានភាព",
      jobStatusOpen: "កំពុងបើក",
      jobStatusClosed: "បិទហើយ",
      jobDesc: "ការពិពណ៌នាការងារ",
      jobPostedDate: "ថ្ងៃបង្ហោះ",
      jobClosingDate: "ថ្ងៃផុតកំណត់ដាក់ពាក្យ (បើមាន)",
      noJobs: "មិនទាន់មានកាលានុវត្តភាពការងារទេ",
      confirmDelJob: (title) => `តើអ្នកប្រាកដទេថាចង់លុប "${title}"?`,
      candidatesFor: (n) => `បេក្ខជន ${n} នាក់`,
      addCandidateBtn: "បន្ថែមបេក្ខជន",
      editCandidateTitle: "កែសម្រួលបេក្ខជន",
      addCandidateTitle: "បន្ថែមបេក្ខជនថ្មី",
      candidateName: "ឈ្មោះបេក្ខជន",
      candidateNamePlaceholder: "ឧ. កញ្ញា សុភា",
      candidatePhone: "លេខទូរស័ព្ទ",
      candidateEmail: "អ៊ីម៉ែល",
      appliedFor: "ដាក់ពាក្យសម្រាប់តំណែង",
      noJobSelected: "មិនទាន់ជ្រើសរើសកាលានុវត្តភាព",
      stage: "ដំណាក់កាល",
      stageApplied: "បានដាក់ពាក្យ",
      stageScreening: "កំពុងពិនិត្យ",
      stageInterview: "សម្ភាសន៍",
      stageOffer: "ផ្តល់ជូនការងារ",
      stageHired: "ជាប់ការងារ",
      stageRejected: "បដិសេធ",
      allStages: "ដំណាក់កាលទាំងអស់",
      notes: "កំណត់ចំណាំ",
      resume: "ប្រវត្តិរូបសង្ខេប (CV/Resume)",
      uploadResume: "ផ្ទុក CV",
      noResume: "មិនទាន់មានឯកសារ",
      appliedDate: "ថ្ងៃដាក់ពាក្យ",
      searchCandidates: "ស្វែងរកបេក្ខជន...",
      noCandidates: "មិនទាន់មានបេក្ខជនទេ",
      confirmDelCandidate: (name) => `តើអ្នកប្រាកដទេថាចង់លុប "${name}"?`,
      hireHint:
        'នៅពេលបេក្ខជនត្រូវបានជ្រើសរើស សូមបន្ថែមគាត់ដោយផ្ទាល់នៅផ្នែក "បុគ្គលិក"',
    },
    onboard: {
      title: "ចាប់ផ្តើម/បញ្ចប់ការងារ",
      subtitle:
        "តាមដានកិច្ចការចាំបាច់ពេលបុគ្គលិកថ្មីចូលធ្វើការ ឬពេលបុគ្គលិកចាស់ឈប់ធ្វើការ",
      selectEmployee: "ជ្រើសរើសបុគ្គលិក",
      chooseEmployee: "-- ជ្រើសរើសបុគ្គលិក --",
      typeOnboarding: "ចាប់ផ្តើមការងារ (Onboarding)",
      typeOffboarding: "បញ្ចប់ការងារ (Offboarding)",
      loadTemplate: "ផ្ទុកបញ្ជីស្តង់ដារ",
      addTask: "បន្ថែមកិច្ចការ",
      taskPlaceholder: "ឧ. ចេញកាតសម្គាល់បុគ្គលិក",
      dueDate: "កាលកំណត់ (បើមាន)",
      noTasks: "មិនទាន់មានកិច្ចការទេ",
      noTasksHint:
        'ចុច "ផ្ទុកបញ្ជីស្តង់ដារ" ដើម្បីចាប់ផ្តើម ឬបន្ថែមផ្ទាល់ខ្លួន',
      progress: (done, total) => `បានបញ្ចប់ ${done}/${total}`,
      confirmDelTask: "តើអ្នកប្រាកដទេថាចង់លុបកិច្ចការនេះ?",
      selectEmployeeHint: "សូមជ្រើសរើសបុគ្គលិកមុនសិន",
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
      namePlaceholder: "ឧ. លោក សុវណ្ណ ដារា",
      codePlaceholder: "EMP-004",
      rolePlaceholder: "ឧ. Accountant",
      pinLabel: "កូដសម្ងាត់ PIN (សម្រាប់ Login)",
      pinRegenerate: "បង្កើតថ្មី",
      status: "ស្ថានភាព",
      branch: "សាខា (Branch)",
      branchNotSet: "មិនទាន់កំណត់",
      annualLeaveDaysLabel: "ថ្ងៃច្បាប់ប្រចាំឆ្នាំ (ថ្ងៃ/ឆ្នាំ)",
      sickLeaveDaysLabel: "ថ្ងៃច្បាប់ឈឺ (ថ្ងៃ/ឆ្នាំ)",
      dependentsLabel: "ចំនួនអាស្រ័យ (កូន/ប្តី-ប្រពន្ធ)",
      dependentsHint:
        "ប្រើសម្រាប់គណនាការកាត់ពន្ធលើប្រាក់បៀវត្សរ៍ (១៥០,០០០ រៀល/នាក់) នៅពេលបើកមុខងារគណនាតាមច្បាប់កម្ពុជា",
      weeklyOffLabel: "ថ្ងៃឈប់សម្រាកប្រចាំសប្តាហ៍",
      customDaysOffLabel: "ថ្ងៃឈប់សម្រាកពិសេស (កាលបរិច្ឆេទជាក់លាក់)",
      addBtnShort: "បន្ថែម",
      searchPlaceholder: "ស្វែងរកបុគ្គលិក...",
      allBranches: "គ្រប់សាខា",
      confirmDelWithName: (name) => `តើអ្នកពិតជាចង់លុប "${name}" មែនទេ?`,
      importBtn: "នាំចូល",
      importTitle: "នាំចូលបុគ្គលិកជាបាច់",
      importDesc:
        "ទាញយកគំរូឯកសារខាងក្រោម បំពេញព័ត៌មានបុគ្គលិក រួចផ្ទុកឡើងវិញ។ ជួរឈរ៖ ឈ្មោះ, លេខសម្គាល់, នាយកដ្ឋាន, តួនាទី, វេន, លេខទូរស័ព្ទ, អ៊ីម៉ែល, ប្រាក់ខែ, ថ្ងៃចូលធ្វើការ, សាខា, ស្ថានភាព។ (នាយកដ្ឋាន/វេន/សាខា ត្រូវសរសេរឲ្យត្រូវនឹងឈ្មោះដែលមានស្រាប់ក្នុងប្រព័ន្ធ)",
      downloadTemplate: "ទាញយកគំរូ (Excel)",
      chooseFile: "ជ្រើសរើសឯកសារ (.xlsx / .csv)",
      changeFile: "ជ្រើសរើសឯកសារផ្សេង",
      parsing: "កំពុងអានឯកសារ...",
      previewTitle: "ត្រួតពិនិត្យទិន្នន័យមុននាំចូល",
      colRow: "ជួរ",
      colStatus: "ស្ថានភាព",
      rowValid: "ត្រឹមត្រូវ",
      rowError: "មានបញ្ហា",
      summaryText: (total, valid, errors) =>
        `សរុប ${total} ជួរ — ត្រឹមត្រូវ ${valid} ជួរ, មានបញ្ហា ${errors} ជួរ`,
      confirmImportBtn: (n) => `នាំចូល ${n} នាក់`,
      importSuccessToast: (n) => `បាននាំចូលបុគ្គលិកចំនួន ${n} នាក់ដោយជោគជ័យ`,
      noValidRows:
        "គ្មានជួរណាត្រឹមត្រូវសម្រាប់នាំចូលទេ សូមកែឯកសារ ហើយផ្ទុកម្តងទៀត",
      emptyFile: "ឯកសារនេះគ្មានទិន្នន័យទេ",
      backToUpload: "ត្រឡប់ក្រោយ",
      errMissingName: "ខ្វះឈ្មោះ",
      errMissingCode: "ខ្វះលេខសម្គាល់",
      errDuplicateCode: (code) => `លេខសម្គាល់ "${code}" ត្រូវបានប្រើរួចហើយ`,
      errDeptNotFound: (v) => `រកមិនឃើញនាយកដ្ឋាន "${v}" ទេ`,
      errShiftNotFound: (v) => `រកមិនឃើញវេន "${v}" ទេ`,
      errOfficeNotFound: (v) => `រកមិនឃើញសាខា "${v}" ទេ`,
      errInvalidSalary: "ប្រាក់ខែមិនត្រឹមត្រូវ",
      importFontNote:
        "ចំណាំ៖ ប្រសិនបើឈ្មោះនាយកដ្ឋាន/វេន បង្ហាញជាប្រអប់ ឬបន្ទាត់នៅក្នុង Excel នោះគ្រាន់តែជាបញ្ហាពុម្ពអក្សរប៉ុណ្ណោះ (ពុម្ពអក្សរលំនាំដើមរបស់ Excel មិនអាចបង្ហាញអក្សរខ្មែរបានទេ)។ ទិន្នន័យខាងក្នុងនៅតែត្រឹមត្រូវ អាចវាយសរសេរជាន់លើ ឬកែសម្រួលធម្មតា ហើយនៅតែនាំចូលចូលប្រព័ន្ធបានធម្មតា។ ដើម្បីឲ្យបង្ហាញច្បាស់ សូមជ្រើសរើសជួរឈរនោះ ហើយប្តូរពុម្ពអក្សរទៅ Khmer OS ឬ Khmer Sangam MN។",
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
      namePlaceholder: "ឧ. វេនព្រឹក",
      overnightHint: "ⓘ វេននេះលើសពាក់កណ្តាលអធ្រាត្រ (ឧ. ចូលយប់ ចេញព្រឹក)",
      overnightTag: "(ឆ្លងអធ្រាត្រ)",
      confirmDelWithCount: (name, n) =>
        `តើអ្នកពិតជាចង់លុបវេន "${name}" មែនទេ? (មានបុគ្គលិក ${n} នាក់កំពុងប្រើវេននេះ)`,
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
      branch: "សាខា",
      liveClockLabel: "ម៉ោងបច្ចុប្បន្ន",
      punchOutBtn: "ចុចចេញពីធ្វើការ",
      awayFromOffice: (name, dist, radius) =>
        `អ្នកនៅឆ្ងាយពីការិយាល័យ "${name}" ${dist}m (កំណត់អនុញ្ញាត ${radius}m) — មិនអាចចុះឈ្មោះបានទេ`,
      cannotVerifyLoc: "មិនអាចផ្ទៀងផ្ទាត់ទីតាំងបានទេ",
      wrongBranchWarning: (name) =>
        `⚠️ អ្នកកំពុងចុះឈ្មោះនៅសាខា "${name}" ប៉ុន្តែអ្នកត្រូវបានកំណត់ឲ្យធ្វើការនៅសាខាផ្សេង`,
      gpsFailed:
        "មិនអាចទាញយកទីតាំង GPS បានទេ សូមបើក Location សម្រាប់កម្មវិធីនេះ",
      gpsRequiredHint: (n) =>
        `ត្រូវការទីតាំង GPS នៅជិតសាខាមួយក្នុងចំណោម ${n} សាខា`,
      noData: "គ្មានទិន្នន័យ",
      manualEntry: "កត់ត្រាដោយដៃ",
      dayOffNote:
        "🛌 ថ្ងៃនេះជាថ្ងៃឈប់សម្រាករបស់អ្នក — អ្នកនៅតែអាចចុះឈ្មោះបានប្រសិនបើអ្នកមកធ្វើការ",
      officeGeofenceTitle: "ការការពារទីតាំង GPS សម្រាប់ Check-in (ច្រើនសាខា)",
      officeGeofenceDesc:
        "កំណត់ទីតាំងសាខានីមួយៗ ដើម្បីតម្រូវឲ្យបុគ្គលិកនៅជិតសាខាមួយណាមួយ ពេលចុច check-in/check-out ដោយខ្លួនឯង។ ឈ្មោះសាខាដែលបុគ្គលិកចូលជិត នឹងត្រូវបានកត់ត្រាទុកជាមួយកំណត់ត្រាវត្តមានរបស់គេ។ បើមិនបន្ថែមសាខាណាមួយទេ ការការពារទីតាំងនឹងមិនដំណើរការទេ។",
      officeCountLabel: (n) => `${n} សាខា`,
      officeNotSet: "មិនទាន់កំណត់",
      officeNoneYet: "មិនទាន់មានសាខាទេ",
      officeNameLabel: "ឈ្មោះសាខា / ការិយាល័យ",
      officeNamePlaceholder: "ការិយាល័យកណ្តាល, សាខាទួលគោក...",
      officeRadiusLabel: "កាំអនុញ្ញាត (ម៉ែត្រ)",
      officeUseCurrentLoc: "ប្រើទីតាំងបច្ចុប្បន្ន",
      officeGpsFailed: "មិនអាចទាញយកទីតាំង GPS បច្ចុប្បន្នបានទេ",
      officeNameRequired: "សូមបញ្ចូលឈ្មោះសាខា/ការិយាល័យ",
      officeCoordsRequired: "សូមបំពេញកូអរដោនេ និងកាំឲ្យត្រឹមត្រូវ",
      officeAddBtn: "បន្ថែមសាខា",
      officeEditTitle: "កែសម្រួលសាខា",
      officeConfirmDel: (name) => `តើអ្នកពិតជាចង់លុបសាខា "${name}" មែនទេ?`,
      statusPresent: "មកធ្វើការ",
      statusLate: "មកយឺត",
      statusAbsent: "អវត្តមាន (មិនបានអនុញ្ញាត)",
      statusLeavePaid: "ឈប់សម្រាក (មានប្រាក់ខែ)",
      statusUnpaid: "ច្បាប់គ្មានប្រាក់ខែ (UL)",
      scanQrBtn: "ស្កេន QR សាខា",
      scanQrOr: "ឬ",
      scanQrTitle: "ស្កេន QR សាខា",
      scanQrDesc: "ដាក់កាមេរ៉ាឲ្យត្រង់ QR code នៅសាខារបស់អ្នក",
      scanQrHint: "កំពុងស្វែងរក QR...",
      scanQrOpeningCamera: "កំពុងបើកកាមេរ៉ា...",
      qrNoMatch: "QR នេះមិនមែនជារបស់សាខាណាមួយឡើយ សូមសាកល្បងម្តងទៀត",
      qrExpired: "QR នេះផុតកំណត់ហើយ សូមស្កេន QR ថ្មីនៅលើអេក្រង់/ក្រដាសនៅសាខា",
      cameraDenied:
        "មិនអាចបើកកាមេរ៉ាបានទេ សូមអនុញ្ញាតការប្រើប្រាស់កាមេរ៉ាសម្រាប់កម្មវិធីនេះ",
      cameraNotFound: "រកមិនឃើញកាមេរ៉ានៅលើឧបករណ៍នេះទេ",
      viaQrLabel: "ស្កេន QR",
      officeQrBtn: "QR Code",
      officeQrTitle: (name) => `QR Code · ${name}`,
      officeQrDesc:
        "បង្ហាញអេក្រង់នេះនៅច្រកចូលសាខា (ឧ. លើថេប្លេត/អេក្រង់តាំងទុក) ដើម្បីឲ្យបុគ្គលិកស្កេន check-in/check-out។ កុំបោះពុម្ពដាក់ជាផ្ទាំង ព្រោះ QR នេះនឹងលែងដំណើរការក្រោយពេលវាផ្លាស់ប្តូរ",
      officeQrRefreshHint:
        "QR នេះនឹងផ្លាស់ប្តូរដោយស្វ័យប្រវត្តិរៀងរាល់ ២០ វិនាទី ដើម្បីសុវត្ថិភាព — សូមកុំថតទុករូបនេះសម្រាប់ប្រើក្រោយ",
      openKioskBtn: "បើកអេក្រង់បង្ហាញ QR",
      kioskScanHint: "ស្កេនកូដដើម្បីចូល ឬ ចេញការងារ",
      kioskRefreshHint: "កូដនេះប្តូរដោយស្វ័យប្រវត្តិរៀងរាល់ ២០ វិនាទី",
      kioskNotFoundTitle: "រកមិនឃើញសាខានេះទេ",
      kioskNotFoundDesc: "តំណនេះមិនត្រឹមត្រូវទេ ឬសាខានេះត្រូវបានលុបហើយ",
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
      typeAnnual: "ច្បាប់ប្រចាំឆ្នាំ",
      typeSick: "ច្បាប់ឈឺ",
      typeUnpaid: "ច្បាប់គ្មានប្រាក់ខែ (UL)",
      typeOther: "ផ្សេងៗ",
      unpaidHint:
        "ថ្ងៃច្បាប់ប្រភេទនេះនឹងមិនបានទទួលប្រាក់ខែទេ — គិតកាត់ដូចថ្ងៃអវត្តមាន",
      remainingAnnual: (n) => `នៅសល់ ${n} ថ្ងៃច្បាប់ប្រចាំឆ្នាំ`,
      remainingSick: (n) => `នៅសល់ ${n} ថ្ងៃច្បាប់ឈឺ`,
      overQuotaWarning: (days, typeLabel, remaining) =>
        `⚠️ សំណើនេះ (${days} ថ្ងៃ) លើសពីសមតុល្យ${typeLabel}ដែលនៅសល់ (${remaining} ថ្ងៃ) — អ្នកនៅតែអាចដាក់ស្នើបាន តែសូមរង់ចាំការសម្រេចពី admin`,
      reasonPlaceholder: "សរសេរមូលហេតុសង្ខេប...",
      submit: "ដាក់ស្នើ",
      annualLeaveYear: (y) => `ថ្ងៃច្បាប់ប្រចាំឆ្នាំ ${y}`,
      sickLeaveYear: (y) => `ថ្ងៃច្បាប់ឈឺ ${y}`,
      usedOfQuota: (used, quota) => `បានប្រើ ${used} ក្នុងចំណោម ${quota} ថ្ងៃ`,
      remainingDays: (n) => `នៅសល់ ${n} ថ្ងៃ`,
      modalTitle: "ស្នើសុំច្បាប់ឈប់សម្រាក",
      fromShort: "ចាប់ពី",
      toShort: "ដល់",
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
      unpaidLeaveDed: "ច្បាប់គ្មានប្រាក់ខែ (UL)",
      lateDed: "កាត់ប្រាក់មកយឺត",
      otPay: "ប្រាក់ OT",
      viewSlip: "មើលសន្លឹកប្រាក់ខែ",
      unmarkPaid: "ដកសញ្ញាបានបើក",
      monthLabel: "ខែ",
      currentMonthTag: "(ខែបច្ចុប្បន្ន)",
      viewingPastMonth: "កំពុងមើលប្រវត្តិខែមុន",
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
      latePolicyDesc:
        "កំណត់ចំនួនដងអនុញ្ញាតឲ្យមកយឺតដោយឥតកាត់ប្រាក់ក្នុងមួយខែ (grace count)។ ចាប់ពីលើសពីនេះទៅ ថ្ងៃមកយឺតនីមួយៗនឹងត្រូវកាត់ប្រាក់ខែ តាមអត្រាដែលកំណត់ខាងក្រោម។",
      lateGraceCountLabel: "ចំនួនដងអនុញ្ញាតឲ្យយឺត/ខែ",
      lateDeductionTypeLabel: "របៀបកាត់ប្រាក់",
      lateDeductionTypeFixed: "ចំនួនថេរ ($) ក្នុងមួយថ្ងៃ",
      lateDeductionTypePercent: "% នៃប្រាក់ថ្ងៃ",
      lateDeductionValueFixedLabel: "ចំនួនកាត់ ($/ថ្ងៃ)",
      lateDeductionValuePercentLabel: "ភាគរយកាត់ (%)",
      lateDeductionHint: (grace) =>
        `បុគ្គលិកអាចយឺតបាន ${grace} ដងដោយឥតកាត់ប្រាក់ ចាប់ពីដងទី ${grace + 1} ទៅ ថ្ងៃនីមួយៗនឹងត្រូវកាត់ប្រាក់`,
      lateDeductionDisabledHint:
        "កំណត់ចំនួនកាត់ឲ្យធំជាង ០ ដើម្បីបើកការកាត់ប្រាក់មកយឺត",
      lateBadgeShort: (grace) => `យឺត > ${grace}`,
      ulPolicyDesc:
        "កំណត់របៀបកាត់ប្រាក់សម្រាប់ច្បាប់គ្មានប្រាក់ខែ (Unpaid Leave/UL)។ លំនាំដើម កាត់ពេញមួយថ្ងៃ (ប្រាក់ខែ/26) ក្នុងមួយថ្ងៃ UL ប៉ុន្តែអាចប្តូរជាចំនួនថេរ ឬភាគរយបាន។",
      ulDeductionTypeLabel: "របៀបកាត់ប្រាក់ UL",
      ulDeductionTypeFullDay: "ពេញមួយថ្ងៃ (ប្រាក់ខែ/26)",
      ulDeductionTypeFixed: "ចំនួនថេរ ($) ក្នុងមួយថ្ងៃ",
      ulDeductionTypePercent: "% នៃប្រាក់ថ្ងៃ",
      ulDeductionValueFixedLabel: "ចំនួនកាត់ ($/ថ្ងៃ)",
      ulDeductionValuePercentLabel: "ភាគរយកាត់ (%)",
      ulDeductionHint:
        "រាល់ថ្ងៃ UL ដែលបានអនុម័ត នឹងត្រូវកាត់ប្រាក់តាមអត្រានេះ (ចំនួនថ្ងៃ × អត្រា)",
      ulBadgeShort: (type) =>
        type === "fixed"
          ? "UL · ថេរ"
          : type === "percentDaily"
            ? "UL · %"
            : "UL · ពេញថ្ងៃ",
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
      customLatePolicyToggle: "កំណត់គោលការណ៍កាត់ប្រាក់មកយឺតផ្ទាល់ខ្លួន",
      customLatePolicyHint:
        "នៅពេលបើក បុគ្គលិកនេះនឹងប្រើចំនួនដងអនុញ្ញាត និងអត្រាកាត់ប្រាក់ដែលកំណត់ខាងក្រោម ជំនួសគោលការណ៍ទូទៅរបស់ក្រុមហ៊ុន",
      customLatePolicyBadge: "គោលការណ៍យឺតផ្ទាល់ខ្លួន",
      customUlPolicyToggle: "កំណត់គោលការណ៍កាត់ប្រាក់ UL ផ្ទាល់ខ្លួន",
      customUlPolicyHint:
        "នៅពេលបើក បុគ្គលិកនេះនឹងប្រើរបៀបកាត់ប្រាក់ UL ដែលកំណត់ខាងក្រោម ជំនួសគោលការណ៍ទូទៅរបស់ក្រុមហ៊ុន",
      customUlPolicyBadge: "គោលការណ៍ UL ផ្ទាល់ខ្លួន",
      historicalBtn: "របាយការណ៍ខែចាស់",
      historicalTitle: "របាយការណ៍ប្រាក់ខែខែចាស់",
      historicalDesc:
        "ជ្រើសខែ/ឆ្នាំណាមួយ ដើម្បីទាញយកទិន្នន័យវត្តមានពី database ដោយផ្ទាល់ (មិនកម្រិតត្រឹម ៦ខែថ្មីៗទេ)",
      historicalPick: "ខែ/ឆ្នាំ",
      historicalLoad: "ទាញយកទិន្នន័យ",
      historicalLoading: "កំពុងទាញយកទិន្នន័យ...",
      historicalEmpty: "គ្មានកំណត់ត្រាវត្តមានសម្រាប់ខែនេះទេ",
      historicalError: "មិនអាចទាញយកទិន្នន័យបានទេ សូមព្យាយាមម្តងទៀត",
      taxModeLabel: "របៀបគណនាពន្ធ/ធានារ៉ាប់រង",
      taxModeFlat: "ភាគរយថេរ (កំណត់ដោយខ្លួនឯង)",
      taxModeKhmer:
        "ស្វ័យប្រវត្តិតាមច្បាប់កម្ពុជា (NSSF + ពន្ធលើប្រាក់បៀវត្សរ៍)",
      taxModeKhmerDesc:
        "គណនាដោយស្វ័យប្រវត្តិ តាមតារាងអត្រាពន្ធជាដុំៗ (Progressive) របស់អគ្គនាយកដ្ឋានពន្ធដារ និងអត្រា NSSF ផ្លូវការ។ អត្រាទាំងនេះអាចផ្លាស់ប្តូរតាមច្បាប់ សូមពិនិត្យ និងកែសម្រួលបើចាំបាច់។",
      exchangeRateLabel: "អត្រាប្តូរប្រាក់ (រៀល/ដុល្លារ)",
      exchangeRateHint:
        "ប្រើសម្រាប់បម្លែងប្រាក់ខែពីដុល្លារទៅរៀល មុននឹងគណនាតាមតារាងពន្ធ។ សូមធ្វើបច្ចុប្បន្នភាពតាមអត្រាធនាគារជាតិនៃកម្ពុជា",
      nssfWageCapLabel: "កំរិតកំណត់ប្រាក់ខែសម្រាប់គណនា NSSF (រៀល)",
      nssfWageCapHint:
        "ភាគរយ NSSF គណនាតែលើប្រាក់ខែរហូតដល់កម្រិតកំណត់នេះប៉ុណ្ណោះ ទោះបីជាប្រាក់ខែពិតប្រាកដខ្ពស់ជាងនេះក៏ដោយ",
      nssfPensionEmployeeLabel: "អត្រា NSSF សោធន (និយោជិត កាត់ពីប្រាក់ខែ) %",
      nssfPensionEmployerLabel: "អត្រា NSSF សោធន (និយោជក បង់បន្ថែម) %",
      nssfOrcLabel: "អត្រា NSSF ហានិភ័យការងារ (និយោជកទាំងអស់） %",
      nssfHealthLabel: "អត្រា NSSF សុខភាព (និយោជកទាំងអស់) %",
      nssfEmployerOnlyHint:
        "អត្រាទាំងនេះជាបន្ទុករបស់និយោជកទាំងស្រុង មិនកាត់ពីប្រាក់ខែបុគ្គលិកទេ បង្ហាញត្រឹមជាព័ត៌មានប៉ុណ្ណោះ",
      dependentsLabelShort: "អាស្រ័យ",
      effectiveRateNote: "(អត្រាមធ្យម)",
      employerNssfCostLabel: "បន្ទុក NSSF របស់ក្រុមហ៊ុន (មិនកាត់ពីប្រាក់ខែ)",
      taxableIncomeLabel: "ប្រាក់ចំណូលជាប់ពន្ធ",
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
      view: "មើល",
      download: "ទាញយក",
      myTitle: "ឯកសាររបស់ខ្ញុំ",
      expiryDate: "ថ្ងៃផុតកំណត់ (បើមាន)",
      expiryDateHint: "ទុកទទេ ប្រសិនបើឯកសារនេះមិនផុតកំណត់",
      expired: "ផុតកំណត់ហើយ",
      expiringSoon: (n) => `ជិតផុតកំណត់ (${n} ថ្ងៃទៀត)`,
      expiresOn: "ផុតកំណត់នៅ",
    },
    cert: {
      btn: "ចេញវិញ្ញាបនបត្រ",
      modalTitle: "ចេញវិញ្ញាបនបត្រ",
      type: "ប្រភេទវិញ្ញាបនបត្រ",
      typeEmployment: "លិខិតបញ្ជាក់ការងារ",
      typeSalary: "លិខិតបញ្ជាក់ប្រាក់ខែ",
      typeService: "លិខិតបញ្ជាក់រយៈពេលបម្រើការងារ",
      typeCustom: "អត្ថបទផ្ទាល់ខ្លួន",
      titleLabel: "ចំណងជើងវិញ្ញាបនបត្រ",
      refNo: "លេខយោង",
      issueDate: "កាលបរិច្ឆេទចេញ",
      endDate: "ដល់ថ្ងៃទី (ប្រសិនបើឈប់ការងារ)",
      endDateHint: "ទុកទទេ ប្រសិនបើនៅបម្រើការងារបន្ត",
      purpose: "គោលបំណង / ជូនចំពោះ",
      purposePlaceholder: "ឧ. សម្រាប់ដាក់ពាក្យធ្វើទិដ្ឋាការ",
      customBody: "អត្ថបទវិញ្ញាបនបត្រ",
      customBodyPlaceholder: "សរសេរខ្លឹមសារវិញ្ញាបនបត្រនៅទីនេះ...",
      signatoryName: "ឈ្មោះអ្នកចុះហត្ថលេខា",
      signatoryTitle: "តួនាទីអ្នកចុះហត្ថលេខា",
      generate: "បង្កើត និង បោះពុម្ព",
      toWhomItMayConcern: "ជូនចំពោះជាទីគោរព",
      signatureLine: "ហត្ថលេខា និង ត្រា",
      issuedOn: "ចេញឱ្យនៅថ្ងៃទី",
      toPresent: "រហូតដល់បច្ចុប្បន្ន",
      toDate: "រហូតដល់ថ្ងៃទី {date}",
      defaultPurpose: "តាមការស្នើសុំ",
      bodyEmployment:
        "សូមបញ្ជាក់ថា {name} (លេខសម្គាល់៖ {code}) បច្ចុប្បន្នកំពុងបម្រើការងារនៅ {company} ក្នុងតួនាទីជា {position} ចាប់ពីថ្ងៃទី {joinDate} រហូតដល់បច្ចុប្បន្ន។ លិខិតបញ្ជាក់នេះត្រូវបានចេញ {purpose}។",
      bodySalary:
        "សូមបញ្ជាក់ថា {name} (លេខសម្គាល់៖ {code}) កំពុងបម្រើការងារនៅ {company} ក្នុងតួនាទីជា {position} ដោយទទួលបានប្រាក់ខែប្រចាំខែចំនួន {salary}។ លិខិតបញ្ជាក់នេះត្រូវបានចេញ {purpose}។",
      bodyService:
        "សូមបញ្ជាក់ថា {name} (លេខសម្គាល់៖ {code}) បានបម្រើការងារនៅ {company} ក្នុងតួនាទីជា {position} ចាប់ពីថ្ងៃទី {joinDate} {endPart}។ លិខិតបញ្ជាក់នេះត្រូវបានចេញ {purpose}។",
      generatedOn: "បង្កើតនៅថ្ងៃទី",
      popupBlocked:
        "កម្មវិធីរុករករបស់អ្នកបានទប់ស្កាត់បង្អួចវិញ្ញាបនបត្រ។ សូមអនុញ្ញាត pop-up រួចសាកល្បងម្តងទៀត។",
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
    ss: {
      addBtn: "ស្នើសុំដូរវេន",
      date: "ថ្ងៃចង់ឱ្យប្តូរចូលជាធរមាន",
      myShift: "វេនរបស់អ្នក",
      fromShift: "វេនបច្ចុប្បន្ន",
      toShift: "ដូរទៅវេន",
      shiftChange: "ការផ្លាស់ប្តូរវេន",
      noOtherShift: "គ្មានវេនផ្សេងទៀតដើម្បីស្នើសុំដូរទេ",
      reason: "មូលហេតុ",
      reasonPlaceholder: "ឧ. មានកិច្ចការគ្រួសារនៅថ្ងៃនោះ...",
      approve: "អនុម័ត",
      reject: "បដិសេធ",
      approvedBy: "អនុម័តដោយ",
      rejectedBy: "បដិសេធដោយ",
      rejectTitle: "បដិសេធសំណើដូរវេន",
      rejectReason: "មូលហេតុបដិសេធ",
      rejectReasonPlaceholder: "សូមបញ្ជាក់មូលហេតុបដិសេធ...",
      rejectReasonRequired: "សូមបញ្ចូលមូលហេតុបដិសេធ",
      confirmDel: "តើអ្នកប្រាកដទេថាចង់លុបសំណើនេះ?",
      noRequest: "មិនទាន់មានសំណើដូរវេនទេ",
      needReason: "សូមបញ្ចូលមូលហេតុ",
      note: "ចំណាំ៖ ពេលអនុម័ត វេនរបស់បុគ្គលិកនឹងផ្លាស់ប្តូរដោយស្វ័យប្រវត្តិទៅតាមវេនដែលបានស្នើសុំភ្លាមៗ គ្មានចាំបាច់កែដោយដៃទៀតទេ។",
    },
    chat: {
      title: "សារ",
      searchEmployee: "ស្វែងរកបុគ្គលិក...",
      noConversations: "មិនទាន់មានការសន្ទនាទេ",
      noEmployeesFound: "រកមិនឃើញបុគ្គលិកទេ",
      selectEmployee: "ជ្រើសរើសបុគ្គលិកម្នាក់ ដើម្បីចាប់ផ្តើមសន្ទនា",
      placeholder: "សរសេរសារ...",
      send: "ផ្ញើ",
      noMessages: "មិនទាន់មានសារទេ សរសេរសារដើម្បីចាប់ផ្តើមសន្ទនា",
      adminLabel: "អ្នកគ្រប់គ្រង",
      youLabel: "អ្នក",
      newMessageFrom: "សារថ្មីពី",
      newMessageFromAdmin: "សារថ្មីពីអ្នកគ្រប់គ្រង",
      today: "ថ្ងៃនេះ",
      attach: "ភ្ជាប់ឯកសារ",
      removeAttachment: "លុបចេញ",
      fileTooLarge: "ឯកសារធំពេក សូមជ្រើសរើសឯកសារតូចជាងនេះ (តិចជាង 8MB)",
      fileReadError: "មិនអាចអានឯកសារនេះបានទេ សូមសាកល្បងម្តងទៀត",
      recordVoice: "ថតសារជាសំឡេង",
      recording: "កំពុងថត",
      stopRecording: "បញ្ចប់",
      cancelRecording: "បោះបង់ការថត",
      micDenied:
        "មិនអាចចូលប្រើមីក្រូហ្វូនបានទេ សូមអនុញ្ញាតសិទ្ធិប្រើមីក្រូហ្វូនសម្រាប់គេហទំព័រនេះ",
      micUnsupported: "ឧបករណ៍ ឬកម្មវិធីរុករកនេះមិនគាំទ្រការថតសំឡេងទេ",
      download: "ទាញយក",
      photoAttachment: "រូបភាព",
      edited: "បានកែសម្រួល",
      seen: "បានឃើញ",
      delivered: "បានផ្ញើ",
      messageDeleted: "សារនេះត្រូវបានលុប",
      editMessage: "កែសម្រួលសារ",
      deleteMessage: "លុបសារ",
      confirmDeleteMsg:
        "តើអ្នកប្រាកដទេថាចង់លុបសារនេះ? សកម្មភាពនេះមិនអាចត្រឡប់វិញបានទេ",
      deleteConversation: "លុបការសន្ទនាទាំងមូល",
      confirmDeleteConversation:
        "តើអ្នកប្រាកដទេថាចង់លុបការសន្ទនាទាំងមូលជាមួយបុគ្គលិកនេះ? សារទាំងអស់នឹងបាត់ជាអចិន្ត្រៃយ៍ សកម្មភាពនេះមិនអាចត្រឡប់វិញបានទេ",
      saveEdit: "រក្សាទុក",
      cancelEdit: "បោះបង់",
      call: "ហៅសំឡេង",
      callAdmin: "ហៅអ្នកគ្រប់គ្រង",
      callEmployee: "ហៅបុគ្គលិកនេះ",
      calling: "កំពុងហៅ...",
      incomingCall: "សំឡេងហៅចូល",
      accept: "ទទួល",
      decline: "បដិសេធ",
      hangup: "ព្យួរទូរស័ព្ទ",
      mute: "បិទសំឡេង",
      unmute: "បើកសំឡេង",
      callBusy: "ម្ខាងទៀតកំពុងជជែកខ្សែផ្សេង",
      callDeclined: "ការហៅត្រូវបានបដិសេធ",
      callNoAnswer: "គ្មានការឆ្លើយតប",
      callConnectionLost: "ការតភ្ជាប់ត្រូវបានផ្តាច់",
      callFailedToStart: "មិនអាចចាប់ផ្តើមការហៅបានទេ សូមសាកល្បងម្តងទៀត",
      callCancelled: "ការហៅត្រូវបានបោះបង់",
      inCall: "កំពុងហៅ",
    },
    admAcc: {
      addBtn: "បន្ថែមអ្នកគ្រប់គ្រង",
      editTitle: "កែសម្រួលអ្នកគ្រប់គ្រង",
      addTitle: "បន្ថែមអ្នកគ្រប់គ្រង",
      roleLabel: "តួនាទី",
      confirmDel: "តើអ្នកប្រាកដទេថាចង់លុបគណនីនេះ?",
      username: "ឈ្មោះគណនី (Username)",
      usernamePlaceholder: "ឧ. manager2",
      password: "ពាក្យសម្ងាត់",
      fullAccessSuffix: "— សិទ្ធិពេញលេញ",
      permsNote:
        'តើតួនាទីនីមួយៗអាចធ្វើអ្វីបាន? កំណត់នៅទំព័រ "សិទ្ធិតួនាទី" (Superadmin ប៉ុណ្ណោះ)',
      confirmDelWithName: (name) => `តើអ្នកពិតជាចង់លុបគណនី "${name}" មែនទេ?`,
      fullName: "ឈ្មោះពេញ",
      fullNamePlaceholder: "ឧ. សុខ សម្បត្តិ",
    },
    profile: {
      title: "ប្រវត្តិរូបរបស់ខ្ញុំ",
      editBtn: "កែប្រវត្តិរូប",
      changePin: "ប្តូរកូដសម្ងាត់ (PIN)",
      oldPin: "កូដសម្ងាត់បច្ចុប្បន្ន",
      newPin: "កូដសម្ងាត់ថ្មី",
      confirmPin: "បញ្ជាក់កូដសម្ងាត់ថ្មី",
      pinDesc: "កូដសម្ងាត់នេះប្រើសម្រាប់ចូលប្រើប្រព័ន្ធ។ សូមកុំប្រាប់អ្នកដទៃ។",
      pinChanged: "បានប្តូរកូដសម្ងាត់ដោយជោគជ័យ",
      pinWrongCurrent: "កូដសម្ងាត់បច្ចុប្បន្នមិនត្រឹមត្រូវទេ",
      pinTooShort: "កូដសម្ងាត់ថ្មីត្រូវមានយ៉ាងតិច 4 ខ្ទង់",
      pinMismatch: "ការបញ្ជាក់កូដសម្ងាត់ថ្មីមិនត្រូវគ្នាទេ",
      joinedSince: "ចូលបម្រើការតាំងពី",
      photoLabel: "រូបភាពប្រវត្តិរូប",
      choosePhoto: "ជ្រើសរើសរូបភាព",
      phone: "លេខទូរស័ព្ទ",
      email: "អ៊ីមែល",
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
      empPortalDesc:
        "ចែករំលែកតំណ ឬ QR Code នេះទៅបុគ្គលិក ដើម្បីឲ្យពួកគេចូលប្រើប្រព័ន្ធ",
      soundTitle: "សំឡេងពេលស្កេន QR",
      soundDesc: "ជ្រើសរើសសំឡេងដែលនឹងឮពេលបុគ្គលិកស្កេន QR ចូល/ចេញការងារជោគជ័យ",
      soundPresets: {
        chime: "សំឡេងកណ្តឹង (លំនាំដើម)",
        bell: "សំឡេងកណ្តឹងវត្ត",
        marimba: "សំឡេងម៉ារីមបា",
        pop: "សំឡេងខ្លីៗ",
        classic: "សំឡេងធម្មតា",
        silent: "គ្មានសំឡេង",
      },
      soundPreview: "ស្តាប់សាកល្បង",
      soundSaved: "បានរក្សាទុកសំឡេងដោយជោគជ័យ",
      pushTitle: "ការជូនដំណឹងលើឧបករណ៍ (Push)",
      pushDesc:
        "បើកដើម្បីទទួលការជូនដំណឹងភ្លាមៗនៅលើឧបករណ៍នេះ ទោះបីអ្នកមិនបានបើកកម្មវិធីនេះក៏ដោយ",
      pushEnable: "បើកការជូនដំណឹង",
      pushDisable: "បិទការជូនដំណឹងលើឧបករណ៍នេះ",
      pushEnabledMsg: "ការជូនដំណឹងលើឧបករណ៍នេះកំពុងបើក",
      pushBlocked:
        "ការជូនដំណឹងត្រូវបានទប់ស្កាត់សម្រាប់គេហទំព័រនេះ សូមកែសម្រួលការអនុញ្ញាតនៅក្នុងកម្មវិធីរុករក",
      pushError: "មិនអាចបើកការជូនដំណឹងបានទេ៖",
      telegramTitle: "ការជូនដំណឹងតាម Telegram",
      telegramDesc:
        "ភ្ជាប់ Telegram Bot ដើម្បីឲ្យក្រុមអ្នកគ្រប់គ្រងទទួលបានសារជូនដំណឹងភ្លាមៗ ពេលមានសំណើថ្មីរង់ចាំអនុម័ត (ច្បាប់ / OT / កែតម្រូវវត្តមាន / ដូរវេន) ចូលទៅក្នុងក្រុម Telegram របស់ក្រុមហ៊ុន",
      telegramEnable: "បើកការជូនដំណឹងតាម Telegram",
      telegramBotTokenLabel: "Bot Token",
      telegramBotTokenPlaceholder: "ទទួលបានពី @BotFather",
      telegramChatIdLabel: "Chat ID",
      telegramChatIdPlaceholder: "ឧ. -1001234567890",
      telegramChatIdHint:
        "បន្ថែម Bot របស់អ្នកទៅក្រុម Telegram រួចប្រើ @userinfobot ឬ getUpdates API ដើម្បីរក Chat ID",
      telegramCategoriesLabel: "ជូនដំណឹងសម្រាប់",
      telegramCatLeave: "សំណើសុំច្បាប់",
      telegramCatOt: "សំណើសុំ OT",
      telegramCatAttcorr: "សំណើកែតម្រូវវត្តមាន",
      telegramCatShiftswap: "សំណើដូរវេន",
      telegramCatPayroll: "ប្រាក់ខែបានបើកចប់",
      telegramCatLate: "បុគ្គលិកមកយឺត",
      telegramCatChat: "សារជជែក (Chat)",
      telegramTestBtn: "ផ្ញើសារសាកល្បង",
      telegramTestSending: "កំពុងផ្ញើ...",
      telegramTestSuccess:
        "បានផ្ញើសារសាកល្បងជោគជ័យ! សូមពិនិត្យក្រុម Telegram របស់អ្នក",
      telegramTestError:
        "មិនអាចផ្ញើសារបានទេ សូមពិនិត្យ Bot Token / Chat ID ឡើងវិញ",
      telegramSaved: "បានរក្សាទុកការកំណត់ Telegram",
      telegramNeedsSetup: "សូមបំពេញ Bot Token និង Chat ID ដើម្បីបើកមុខងារនេះ",
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
      shiftSwap: "Shift Swap Requests",
      messages: "Messages",
      assets: "Company Assets",
      training: "Training & Certifications",
      docExpiry: "Contracts & Expiring Docs",
      recruitment: "Recruitment",
      onboarding: "Onboarding / Offboarding",
      admins: "Admin Accounts",
      myAttendance: "My Attendance",
      myLeave: "My Leave",
      myOvertime: "My Overtime (OT)",
      myPayroll: "My Payroll",
      myPerformance: "My Performance Reviews",
      myTraining: "My Training & Certifications",
      myDocuments: "My Documents",
      myAttCorrection: "Attendance Correction",
      myShiftSwap: "Shift Swap Request",
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
    clear: "Clear",
    today: "Today",
    yesterday: "Yesterday",
    now: "Now",
    timeLabel: "Time",
    selectDate: "Select a date",
    noResults: "No results found",
    pagination: { of: "of" },
    popupBlockedTitle: "Pop-up Blocked",
    popupBlockedPayslip:
      "Please allow pop-ups for this site to download the payslip",
    popupBlockedBadge: "Please allow pop-ups for this site to print the badge",
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
      workingNow: "Working Now",
      workingNowSub: (n) => `${n} clocked in right now`,
      noOneWorkingNow: "No one is checked in yet",
      unassignedBranch: "Unassigned branch",
      sinceLabel: "Since",
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
      newHiresTrend: "New Hires (last 6 months)",
      newHiresTrendSub: "New employees added per month, based on join date",
      absenteeism: "Absenteeism Rate by Department",
      absenteeismSub:
        'Share of "absent" records vs. total attendance records this month',
      statActiveHeadcount: "Total Active Headcount",
      statAvgTenure: "Average Tenure",
      statInactiveRate: "Current Inactive Rate",
      statInactiveRateNote:
        "Note: this is a current snapshot, not a turnover rate over a period — the system doesn't yet track an exit date.",
      tenureFormat: (y, m) => (y > 0 ? `${y}y ${m}m` : `${m}m`),
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
      namePlaceholder: "e.g. Human Resources",
      codePlaceholder: "e.g. HR",
      confirmDelWithCount: (name, n) =>
        `Are you sure you want to delete "${name}"? (${n} staff assigned)`,
    },
    assets: {
      title: "Company Assets",
      addBtn: "Add Asset",
      editTitle: "Edit Asset",
      addTitle: "Add Asset",
      name: "Asset Name",
      category: "Category",
      serial: "Serial / Asset Code",
      assignedTo: "Assigned To",
      unassigned: "Unassigned",
      status: "Status",
      statusAvailable: "Available",
      statusAssigned: "In Use",
      statusMaintenance: "In Maintenance",
      statusRetired: "Retired",
      purchaseDate: "Purchase Date",
      notes: "Notes",
      namePlaceholder: "e.g. Dell Laptop",
      categoryPlaceholder: "e.g. Computer, Vehicle, Phone",
      serialPlaceholder: "e.g. SN-00123",
      searchPlaceholder: "Search assets...",
      allStatus: "All Status",
      noAssets: "No assets yet",
      confirmDelWithName: (name) =>
        `Are you sure you want to delete "${name}"?`,
    },
    training: {
      title: "Training & Certifications",
      subtitle: "Track training courses and certificates for every employee",
      addBtn: "Add Record",
      editTitle: "Edit Training Record",
      addTitle: "Add Training Record",
      employee: "Employee",
      courseName: "Course / Certificate Name",
      provider: "Provider / Institution",
      category: "Category",
      startDate: "Start Date",
      completionDate: "Completion Date",
      certExpiry: "Certificate Expiry (optional)",
      status: "Status",
      statusPlanned: "Planned",
      statusOngoing: "Ongoing",
      statusCompleted: "Completed",
      statusCancelled: "Cancelled",
      notes: "Notes",
      courseNamePlaceholder: "e.g. Workplace Safety Training",
      providerPlaceholder: "e.g. Ministry of Labour, an NGO, Coursera",
      categoryPlaceholder: "e.g. Safety, Technical, Soft Skills",
      searchPlaceholder: "Search by employee or course name...",
      allStatus: "All Status",
      allEmployees: "All Employees",
      noTrainings: "No training records yet",
      confirmDelWithName: (name) =>
        `Are you sure you want to delete "${name}"?`,
      expired: "Certificate expired",
      expiringSoon: (n) => `Expires in ${n} day(s)`,
      expiresOn: "Expires on",
      myTitle: "My Training & Certifications",
      noMyTrainings: "You don't have any training records yet",
    },
    contracts: {
      title: "Contracts & Expiring Documents",
      subtitle:
        "Track documents (contracts, ID cards, etc.) with an expiry date, across all staff",
      searchPlaceholder: "Search by employee or document name...",
      allStatus: "All Status",
      statusExpired: "Expired",
      statusSoon: "Expiring Soon",
      statusValid: "Valid",
      noItems: "No documents with an expiry date yet",
      noItemsHint:
        'To track something here, set an "Expiry Date" when uploading it under Employees → Documents',
      colEmployee: "Employee",
      colDocument: "Document",
      colCategory: "Category",
      colExpiry: "Expiry Date",
      colStatus: "Status",
      daysLeft: (n) => `${n}d left`,
      daysAgo: (n) => `Expired ${n}d ago`,
    },
    recruit: {
      title: "Recruitment",
      tabJobs: "Job Postings",
      tabCandidates: "Candidates",
      addJobBtn: "Post a Job",
      editJobTitle: "Edit Job Posting",
      addJobTitle: "Post a New Job",
      jobTitle: "Job Title",
      jobTitlePlaceholder: "e.g. Accountant",
      jobDept: "Department",
      jobType: "Employment Type",
      jobTypeFullTime: "Full-time",
      jobTypePartTime: "Part-time",
      jobTypeContract: "Contract",
      jobTypeInternship: "Internship",
      jobStatus: "Status",
      jobStatusOpen: "Open",
      jobStatusClosed: "Closed",
      jobDesc: "Job Description",
      jobPostedDate: "Posted Date",
      jobClosingDate: "Closing Date (optional)",
      noJobs: "No job postings yet",
      confirmDelJob: (title) => `Are you sure you want to delete "${title}"?`,
      candidatesFor: (n) => `${n} candidate(s)`,
      addCandidateBtn: "Add Candidate",
      editCandidateTitle: "Edit Candidate",
      addCandidateTitle: "Add New Candidate",
      candidateName: "Candidate Name",
      candidateNamePlaceholder: "e.g. Sophea Kong",
      candidatePhone: "Phone",
      candidateEmail: "Email",
      appliedFor: "Applied For",
      noJobSelected: "No job posting selected",
      stage: "Stage",
      stageApplied: "Applied",
      stageScreening: "Screening",
      stageInterview: "Interview",
      stageOffer: "Offer",
      stageHired: "Hired",
      stageRejected: "Rejected",
      allStages: "All Stages",
      notes: "Notes",
      resume: "Resume / CV",
      uploadResume: "Upload CV",
      noResume: "No file yet",
      appliedDate: "Applied Date",
      searchCandidates: "Search candidates...",
      noCandidates: "No candidates yet",
      confirmDelCandidate: (name) =>
        `Are you sure you want to delete "${name}"?`,
      hireHint:
        'Once a candidate is hired, add them directly under "Employees"',
    },
    onboard: {
      title: "Onboarding / Offboarding",
      subtitle:
        "Track the tasks needed when a new employee joins, or when one leaves",
      selectEmployee: "Select Employee",
      chooseEmployee: "-- Select an employee --",
      typeOnboarding: "Onboarding",
      typeOffboarding: "Offboarding",
      loadTemplate: "Load Default Checklist",
      addTask: "Add Task",
      taskPlaceholder: "e.g. Issue staff ID badge",
      dueDate: "Due Date (optional)",
      noTasks: "No tasks yet",
      noTasksHint:
        'Click "Load Default Checklist" to get started, or add your own',
      progress: (done, total) => `${done}/${total} done`,
      confirmDelTask: "Are you sure you want to delete this task?",
      selectEmployeeHint: "Please select an employee first",
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
      namePlaceholder: "e.g. John Doe",
      codePlaceholder: "EMP-004",
      rolePlaceholder: "e.g. Accountant",
      pinLabel: "PIN (used for Login)",
      pinRegenerate: "Regenerate",
      status: "Status",
      branch: "Branch",
      branchNotSet: "Not set",
      annualLeaveDaysLabel: "Annual Leave Days (per year)",
      sickLeaveDaysLabel: "Sick Leave Days (per year)",
      dependentsLabel: "Dependents (children / spouse)",
      dependentsHint:
        "Used to calculate the Tax on Salary dependent deduction (150,000 KHR each) when the Cambodia auto-calculation mode is enabled",
      weeklyOffLabel: "Weekly Day(s) Off",
      customDaysOffLabel: "Custom Days Off (specific dates)",
      addBtnShort: "Add",
      searchPlaceholder: "Search employees...",
      allBranches: "All Branches",
      confirmDelWithName: (name) =>
        `Are you sure you want to delete "${name}"?`,
      importBtn: "Import",
      importTitle: "Bulk Import Employees",
      importDesc:
        "Download the template below, fill in employee details, then upload it. Columns: Name, Employee ID, Department, Role, Shift, Phone, Email, Salary, Join Date, Branch, Status. (Department/Shift/Branch must match names already set up in the system)",
      downloadTemplate: "Download Template (Excel)",
      chooseFile: "Choose File (.xlsx / .csv)",
      changeFile: "Choose a Different File",
      parsing: "Reading file...",
      previewTitle: "Review Data Before Importing",
      colRow: "Row",
      colStatus: "Status",
      rowValid: "Valid",
      rowError: "Error",
      summaryText: (total, valid, errors) =>
        `${total} rows total — ${valid} valid, ${errors} with errors`,
      confirmImportBtn: (n) => `Import ${n} Employees`,
      importSuccessToast: (n) => `Successfully imported ${n} employees`,
      noValidRows: "No valid rows to import. Please fix the file and re-upload",
      emptyFile: "This file has no data",
      backToUpload: "Back",
      errMissingName: "Missing name",
      errMissingCode: "Missing employee ID",
      errDuplicateCode: (code) => `Employee ID "${code}" is already in use`,
      errDeptNotFound: (v) => `Department "${v}" not found`,
      errShiftNotFound: (v) => `Shift "${v}" not found`,
      errOfficeNotFound: (v) => `Branch "${v}" not found`,
      errInvalidSalary: "Invalid salary",
      importFontNote:
        "Note: if department/shift names show as boxes or blank marks in Excel, that's just a font display issue — Excel's default font doesn't include Khmer glyphs. The underlying data is still correct, and you can type/edit over it normally; it will still import fine. To display it properly, select that column and change its font to Khmer OS or Khmer Sangam MN.",
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
      namePlaceholder: "e.g. Morning Shift",
      overnightHint:
        "ⓘ This shift crosses midnight (e.g. starts at night, ends in the morning)",
      overnightTag: "(overnight)",
      confirmDelWithCount: (name, n) =>
        `Are you sure you want to delete "${name}"? (${n} staff use this shift)`,
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
      branch: "Branch",
      liveClockLabel: "Current Time",
      punchOutBtn: "Check Out",
      awayFromOffice: (name, dist, radius) =>
        `You're ${dist}m away from "${name}" (allowed radius ${radius}m) — cannot check in`,
      cannotVerifyLoc: "Couldn't verify your location",
      wrongBranchWarning: (name) =>
        `⚠️ You're checking in at "${name}" but you're assigned to a different branch`,
      gpsFailed:
        "Couldn't get your GPS location. Please enable Location for this app.",
      gpsRequiredHint: (n) => `GPS location required near one of ${n} branches`,
      noData: "No data",
      manualEntry: "Manual Entry",
      dayOffNote:
        "🛌 Today is your day off — you can still check in if you're coming to work",
      officeGeofenceTitle: "GPS Check-in Geofencing (Multiple Branches)",
      officeGeofenceDesc:
        "Set the location of each branch to require employees to be near a branch when they self check-in/check-out. The branch they're closest to is recorded on their attendance record. If no branch is added, geofencing is disabled.",
      officeCountLabel: (n) => `${n} branches`,
      officeNotSet: "Not set",
      officeNoneYet: "No branches yet",
      officeNameLabel: "Branch / Office Name",
      officeNamePlaceholder: "Main Office, Toul Kork Branch...",
      officeRadiusLabel: "Allowed Radius (meters)",
      officeUseCurrentLoc: "Use Current Location",
      officeGpsFailed: "Couldn't get your current GPS location",
      officeNameRequired: "Please enter a branch/office name",
      officeCoordsRequired: "Please fill in valid coordinates and radius",
      officeAddBtn: "Add Branch",
      officeEditTitle: "Edit Branch",
      officeConfirmDel: (name) =>
        `Are you sure you want to delete branch "${name}"?`,
      statusPresent: "Present",
      statusLate: "Late",
      statusAbsent: "Absent (Unexcused)",
      statusLeavePaid: "On Leave (Paid)",
      statusUnpaid: "Unpaid Leave (UL)",
      scanQrBtn: "Scan Branch QR",
      scanQrOr: "or",
      scanQrTitle: "Scan Branch QR",
      scanQrDesc: "Point your camera at the QR code at your branch",
      scanQrHint: "Looking for a QR code...",
      scanQrOpeningCamera: "Opening camera...",
      qrNoMatch: "That QR code doesn't match any branch — try again",
      qrExpired:
        "That QR code has expired — scan the current one at the branch",
      cameraDenied:
        "Couldn't access the camera. Please allow camera access for this app.",
      cameraNotFound: "No camera was found on this device",
      viaQrLabel: "QR scan",
      officeQrBtn: "QR Code",
      officeQrTitle: (name) => `QR Code · ${name}`,
      officeQrDesc:
        "Display this screen at the branch entrance (e.g. on a tablet or monitor) so employees can scan to check in/out. Don't print it as a static poster — this QR stops working once it rotates.",
      officeQrRefreshHint:
        "This QR refreshes automatically every 20 seconds for security — don't save a screenshot to reuse later.",
      openKioskBtn: "Open Kiosk Display",
      kioskScanHint: "Scan the code to check in or check out",
      kioskRefreshHint: "This code refreshes automatically every 20 seconds",
      kioskNotFoundTitle: "Branch not found",
      kioskNotFoundDesc: "This link is invalid, or the branch was deleted",
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
      typeAnnual: "Annual Leave",
      typeSick: "Sick Leave",
      typeUnpaid: "Unpaid Leave (UL)",
      typeOther: "Other",
      unpaidHint:
        "Days under this leave type are not paid — deducted like an absence.",
      remainingAnnual: (n) => `${n} annual leave days remaining`,
      remainingSick: (n) => `${n} sick leave days remaining`,
      overQuotaWarning: (days, typeLabel, remaining) =>
        `⚠️ This request (${days} days) exceeds your remaining ${typeLabel} balance (${remaining} days) — you can still submit it, but it'll need admin approval`,
      reasonPlaceholder: "Write a brief reason...",
      submit: "Submit",
      annualLeaveYear: (y) => `${y} Annual Leave`,
      sickLeaveYear: (y) => `${y} Sick Leave`,
      usedOfQuota: (used, quota) => `Used ${used} of ${quota} days`,
      modalTitle: "New Leave Request",
      fromShort: "From",
      toShort: "To",
      remainingDays: (n) => `${n} days remaining`,
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
      unpaidLeaveDed: "Unpaid Leave (UL)",
      lateDed: "Late-arrival deduction",
      otPay: "OT Pay",
      viewSlip: "View Payslip",
      unmarkPaid: "Unmark as Paid",
      monthLabel: "Month",
      currentMonthTag: "(Current Month)",
      viewingPastMonth: "Viewing past month's history",
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
      latePolicyDesc:
        "Set how many times an employee can arrive late each month before it's deducted (grace count). Every late day beyond that grace count is deducted at the rate set below.",
      lateGraceCountLabel: "Grace count (late days/month)",
      lateDeductionTypeLabel: "Deduction type",
      lateDeductionTypeFixed: "Fixed amount ($) per day",
      lateDeductionTypePercent: "% of daily rate",
      lateDeductionValueFixedLabel: "Deduction amount ($/day)",
      lateDeductionValuePercentLabel: "Deduction percentage (%)",
      lateDeductionHint: (grace) =>
        `Employees can be late ${grace} time(s) with no penalty. From the ${grace + 1}th time on, each late day is deducted.`,
      lateDeductionDisabledHint:
        "Set the deduction amount above 0 to enable late-arrival deductions",
      lateBadgeShort: (grace) => `late > ${grace}`,
      ulPolicyDesc:
        "Set how Unpaid Leave (UL) days are deducted. By default a full day's pay (salary/26) is docked per UL day, but this can be changed to a fixed amount or a percentage instead.",
      ulDeductionTypeLabel: "UL deduction type",
      ulDeductionTypeFullDay: "Full day (salary/26)",
      ulDeductionTypeFixed: "Fixed amount ($) per day",
      ulDeductionTypePercent: "% of daily rate",
      ulDeductionValueFixedLabel: "Deduction amount ($/day)",
      ulDeductionValuePercentLabel: "Deduction percentage (%)",
      ulDeductionHint:
        "Every approved UL day is deducted at this rate (days × rate)",
      ulBadgeShort: (type) =>
        type === "fixed"
          ? "UL · fixed"
          : type === "percentDaily"
            ? "UL · %"
            : "UL · full day",
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
      customLatePolicyToggle: "Custom late-deduction policy for this employee",
      customLatePolicyHint:
        "When enabled, this employee uses the grace count and deduction rate set below instead of the company-wide policy",
      customLatePolicyBadge: "Custom late policy",
      customUlPolicyToggle: "Custom UL-deduction policy for this employee",
      customUlPolicyHint:
        "When enabled, this employee uses the UL deduction rule set below instead of the company-wide policy",
      customUlPolicyBadge: "Custom UL policy",
      historicalBtn: "Older Months Report",
      historicalTitle: "Historical Payroll Report",
      historicalDesc:
        "Pick any month/year to pull attendance directly from the database (not limited to the last 6 months kept live in the app)",
      historicalPick: "Month / Year",
      historicalLoad: "Load Data",
      historicalLoading: "Loading data...",
      historicalEmpty: "No attendance records found for this month",
      historicalError: "Couldn't load this data — please try again",
      taxModeLabel: "Tax / Insurance Calculation Mode",
      taxModeFlat: "Flat percentage (set manually)",
      taxModeKhmer: "Automatic — Cambodia law (NSSF + Tax on Salary)",
      taxModeKhmerDesc:
        "Automatically computed using the General Department of Taxation's progressive salary-tax brackets and official NSSF rates. These rates can change by law — please review and adjust if needed.",
      exchangeRateLabel: "Exchange Rate (KHR per USD)",
      exchangeRateHint:
        "Used to convert base salary from USD to KHR before applying the tax brackets. Update this to match the current National Bank of Cambodia rate",
      nssfWageCapLabel: "NSSF Contributory Wage Cap (KHR)",
      nssfWageCapHint:
        "NSSF percentages are calculated only up to this wage cap, even if the employee's actual salary is higher",
      nssfPensionEmployeeLabel: "NSSF Pension Rate (employee, deducted) %",
      nssfPensionEmployerLabel: "NSSF Pension Rate (employer, extra) %",
      nssfOrcLabel: "NSSF Occupational Risk Rate (employer-only) %",
      nssfHealthLabel: "NSSF Healthcare Rate (employer-only) %",
      nssfEmployerOnlyHint:
        "These rates are fully paid by the employer and are not deducted from the employee's salary — shown for information only",
      dependentsLabelShort: "Dependents",
      effectiveRateNote: "(effective rate)",
      employerNssfCostLabel: "Employer NSSF cost (not deducted from salary)",
      taxableIncomeLabel: "Taxable Income",
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
      view: "View",
      download: "Download",
      myTitle: "My Documents",
      expiryDate: "Expiry Date (optional)",
      expiryDateHint: "Leave blank if this document doesn't expire",
      expired: "Expired",
      expiringSoon: (n) => `Expiring soon (${n}d left)`,
      expiresOn: "Expires",
    },
    cert: {
      btn: "Issue Certificate",
      modalTitle: "Issue Certificate",
      type: "Certificate Type",
      typeEmployment: "Employment Certificate",
      typeSalary: "Salary Certificate",
      typeService: "Certificate of Service",
      typeCustom: "Custom Text",
      titleLabel: "Certificate Title",
      refNo: "Reference No.",
      issueDate: "Issue Date",
      endDate: "End Date (if no longer employed)",
      endDateHint: "Leave blank if still employed",
      purpose: "Purpose / Addressed To",
      purposePlaceholder: "e.g. For visa application",
      customBody: "Certificate Body",
      customBodyPlaceholder: "Write the certificate content here...",
      signatoryName: "Signatory Name",
      signatoryTitle: "Signatory Title",
      generate: "Generate & Print",
      toWhomItMayConcern: "To Whom It May Concern",
      signatureLine: "Signature & Stamp",
      issuedOn: "Issued on",
      toPresent: "to present",
      toDate: "to {date}",
      defaultPurpose: "upon request",
      bodyEmployment:
        "This is to certify that {name} (ID: {code}) is currently employed at {company} as {position}, since {joinDate} to present. This certificate is issued {purpose}.",
      bodySalary:
        "This is to certify that {name} (ID: {code}) is currently employed at {company} as {position}, with a current monthly salary of {salary}. This certificate is issued {purpose}.",
      bodyService:
        "This is to certify that {name} (ID: {code}) has been employed at {company} as {position}, from {joinDate} {endPart}. This certificate is issued {purpose}.",
      generatedOn: "Generated on",
      popupBlocked:
        "Your browser blocked the certificate window. Please allow pop-ups and try again.",
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
    ss: {
      addBtn: "Request Shift Swap",
      date: "Requested effective date",
      myShift: "Your shift",
      fromShift: "Current shift",
      toShift: "Change to shift",
      shiftChange: "Shift change",
      noOtherShift: "No other shift available to request a swap to",
      reason: "Reason",
      reasonPlaceholder: "e.g. Family matter that day...",
      approve: "Approve",
      reject: "Reject",
      approvedBy: "Approved by",
      rejectedBy: "Rejected by",
      rejectTitle: "Reject Shift Swap Request",
      rejectReason: "Rejection Reason",
      rejectReasonPlaceholder: "Please state the rejection reason...",
      rejectReasonRequired: "Please enter a rejection reason",
      confirmDel: "Are you sure you want to delete this request?",
      noRequest: "No shift swap requests yet",
      needReason: "Please enter a reason",
      note: "Note: on approval, the employee's shift is updated automatically to the requested shift — no manual edit needed afterward.",
    },
    chat: {
      title: "Messages",
      searchEmployee: "Search employee...",
      noConversations: "No conversations yet",
      noEmployeesFound: "No employees found",
      selectEmployee: "Select an employee to start a conversation",
      placeholder: "Write a message...",
      send: "Send",
      noMessages: "No messages yet — write one to start the conversation",
      adminLabel: "Admin",
      youLabel: "You",
      newMessageFrom: "New message from",
      newMessageFromAdmin: "New message from admin",
      today: "Today",
      attach: "Attach file",
      removeAttachment: "Remove",
      fileTooLarge: "File is too large — please choose one under 8MB",
      fileReadError: "Couldn't read that file — please try again",
      recordVoice: "Record a voice message",
      recording: "Recording",
      stopRecording: "Stop",
      cancelRecording: "Cancel recording",
      micDenied:
        "Couldn't access the microphone — please allow microphone permission for this site.",
      micUnsupported: "This device or browser doesn't support voice recording.",
      download: "Download",
      photoAttachment: "Photo",
      edited: "edited",
      seen: "Seen",
      delivered: "Delivered",
      messageDeleted: "This message was deleted",
      editMessage: "Edit message",
      deleteMessage: "Delete message",
      confirmDeleteMsg:
        "Are you sure you want to delete this message? This can't be undone.",
      deleteConversation: "Delete entire conversation",
      confirmDeleteConversation:
        "Are you sure you want to delete this entire conversation with this employee? All messages will be permanently gone. This can't be undone.",
      saveEdit: "Save",
      cancelEdit: "Cancel",
      call: "Voice call",
      callAdmin: "Call admin",
      callEmployee: "Call this employee",
      calling: "Calling...",
      incomingCall: "Incoming call",
      accept: "Accept",
      decline: "Decline",
      hangup: "Hang up",
      mute: "Mute",
      unmute: "Unmute",
      callBusy: "They're on another call",
      callDeclined: "Call declined",
      callNoAnswer: "No answer",
      callConnectionLost: "Call connection lost",
      callFailedToStart: "Couldn't start the call — please try again",
      callCancelled: "Call cancelled",
      inCall: "In call",
    },
    admAcc: {
      addBtn: "Add Admin",
      editTitle: "Edit Admin",
      addTitle: "Add Admin",
      roleLabel: "Role",
      confirmDel: "Are you sure you want to delete this account?",
      username: "Username",
      usernamePlaceholder: "e.g. manager2",
      password: "Password",
      fullAccessSuffix: "— Full Access",
      permsNote:
        'What can each role do? Set it on the "Role Permissions" page (Superadmin only)',
      confirmDelWithName: (name) =>
        `Are you sure you want to delete account "${name}"?`,
      fullName: "Full Name",
      fullNamePlaceholder: "e.g. John Doe",
    },
    profile: {
      title: "My Profile",
      editBtn: "Edit Profile",
      changePin: "Change PIN",
      oldPin: "Current PIN",
      newPin: "New PIN",
      confirmPin: "Confirm New PIN",
      pinDesc:
        "This PIN is used to log into the system. Don't share it with anyone.",
      pinChanged: "PIN changed successfully",
      pinWrongCurrent: "Current PIN is incorrect",
      pinTooShort: "New PIN must be at least 4 digits",
      pinMismatch: "PIN confirmation doesn't match",
      joinedSince: "Joined since",
      photoLabel: "Profile Photo",
      choosePhoto: "Choose Photo",
      phone: "Phone",
      email: "Email",
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
      empPortalDesc:
        "Share this link or QR code with employees so they can access the portal",
      soundTitle: "QR Scan Sound",
      soundDesc:
        "Choose the sound that plays when an employee successfully scans a QR to check in/out",
      soundPresets: {
        chime: "Chime (default)",
        bell: "Bell",
        marimba: "Marimba",
        pop: "Pop",
        classic: "Classic beep",
        silent: "Silent",
      },
      soundPreview: "Preview",
      soundSaved: "Sound saved successfully",
      pushTitle: "Push Notifications",
      pushDesc:
        "Turn on to get instant notifications on this device, even when the app isn't open",
      pushEnable: "Enable notifications",
      pushDisable: "Disable on this device",
      pushEnabledMsg: "Push notifications are on for this device",
      pushBlocked:
        "Notifications are blocked for this site — check your browser's site permissions",
      pushError: "Couldn't enable notifications:",
      telegramTitle: "Telegram Notifications",
      telegramDesc:
        "Connect a Telegram Bot so your admin team gets instant messages in a company Telegram group whenever a new request needs approval (leave / OT / attendance correction / shift swap)",
      telegramEnable: "Enable Telegram notifications",
      telegramBotTokenLabel: "Bot Token",
      telegramBotTokenPlaceholder: "From @BotFather",
      telegramChatIdLabel: "Chat ID",
      telegramChatIdPlaceholder: "e.g. -1001234567890",
      telegramChatIdHint:
        "Add your bot to the Telegram group, then use @userinfobot or the getUpdates API to find the Chat ID",
      telegramCategoriesLabel: "Notify for",
      telegramCatLeave: "Leave requests",
      telegramCatOt: "OT requests",
      telegramCatAttcorr: "Attendance correction requests",
      telegramCatShiftswap: "Shift swap requests",
      telegramCatPayroll: "Payroll completed",
      telegramCatLate: "Employee checked in late",
      telegramCatChat: "Chat messages",
      telegramTestBtn: "Send test message",
      telegramTestSending: "Sending...",
      telegramTestSuccess: "Test message sent! Check your Telegram group",
      telegramTestError: "Couldn't send — check the Bot Token / Chat ID",
      telegramSaved: "Telegram settings saved",
      telegramNeedsSetup: "Fill in the Bot Token and Chat ID to enable this",
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
  // Chinese (Simplified) is a partial translation covering the most
  // visible screens (nav, dashboard, attendance/QR, settings, login).
  // Anything not listed here falls back to English automatically — see
  // mergeLangFallback / the LANG export below — rather than showing
  // undefined or crashing.
  zh: {
    appName: "Workforce Suite",
    login: {
      employeeId: "员工编号",
      employeeIdPlaceholder: "EMP-001",
      pin: "密码 (PIN)",
      pinPlaceholder: "••••",
      submit: "登录系统",
      adminSubmit: "管理员登录",
      username: "账号 (Username)",
      usernamePlaceholder: "admin",
      password: "管理员密码",
      passwordPlaceholder: "••••••••",
      adminTitle: "管理员登录",
      employeePortal: "员工登录",
      demoLabel: "测试账号：",
      errNoEmp: "找不到此员工编号",
      errInactive: "此账号尚未启用，请联系管理员",
      errPin: "PIN 密码不正确",
      errNoAdmin: "找不到此管理员账号",
      errPass: "密码不正确",
      back: "返回",
      switchToAdmin: "管理员？请从这里登录",
    },
    nav: {
      dashboard: "仪表板",
      announcements: "公告",
      employees: "员工",
      departments: "部门",
      shifts: "班次",
      attendance: "考勤",
      holidays: "法定假日",
      leave: "请假",
      overtime: "加班 (OT)",
      payroll: "工资",
      performance: "绩效考核",
      attCorrection: "考勤更正申请",
      shiftSwap: "换班申请",
      admins: "管理员账号",
      myAttendance: "我的考勤",
      myLeave: "我的请假",
      myOvertime: "我的加班 (OT)",
      myPayroll: "我的工资",
      myPerformance: "我的绩效考核",
      myDocuments: "我的文件",
      myAttCorrection: "申请考勤更正",
      myShiftSwap: "申请换班",
      myProfile: "我的个人资料",
      settings: "设置",
      audits: "操作日志",
      loginActivity: "登录记录",
      analytics: "数据分析",
      rolePerms: "角色权限",
    },
    logout: "登出",
    notifications: "通知",
    markAllRead: "标记全部已读",
    noNotif: "没有通知",
    confirmDelete: "确认删除",
    cancel: "取消",
    delete: "删除",
    edit: "编辑",
    save: "保存",
    add: "添加",
    search: "搜索...",
    employee: "员工",
    status: "状态",
    actions: "操作",
    noData: "没有数据",
    exportCsv: "导出 CSV",
    exportExcel: "导出 Excel",
    clear: "清除",
    today: "今天",
    yesterday: "昨天",
    now: "现在",
    timeLabel: "时间",
    selectDate: "选择日期",
    noResults: "没有找到结果",
    pagination: { of: "之" },
    popupBlockedTitle: "浏览器已阻止弹出窗口",
    popupBlockedPayslip: "请允许此网站的弹出窗口，以下载工资单",
    popupBlockedBadge: "请允许此网站的弹出窗口，以打印工作证",
    dash: {
      welcome: "欢迎",
      totalEmp: "员工总数",
      active: "在职",
      totalDept: "部门",
      totalDeptSub: "部门总数",
      presentToday: "今日出勤",
      attendRate: "出勤率",
      pendingPayroll: "待发工资",
      thisMonth: "本月",
      recentAttend: "今日考勤记录",
      noAttend: "今天还没有打卡记录",
      empPortalLink: "员工入口链接",
      showQR: "显示 QR Code",
      copyLink: "复制链接",
      copied: "已复制！",
      noEmpWarn: '还没有员工，请前往"员工"页面添加',
      myDept: "我的部门",
      myShift: "我的班次",
      todayStatus: "今日状态",
      payrollStatus: "工资状态",
      notCheckedIn: "尚未打卡",
      workingNow: "正在工作",
      workingNowSub: (n) => `目前有 ${n} 人正在打卡上班`,
      noOneWorkingNow: "目前还没有员工打卡上班",
      unassignedBranch: "未分配分店",
      sinceLabel: "自",
    },
    analytics: {
      title: "数据分析",
      subtitle: "考勤、工资和加班趋势总览",
      attendTrend: "出勤率（近 6 个月）",
      attendTrendSub: "出勤/迟到百分比，对比每月总记录",
      deptCost: "各部门工资净支出",
      deptCostSub: "当前月份，仅计算在职员工",
      otTrend: "已批准加班时数（近 6 个月）",
      otTrendSub: "所有部门已批准的加班时数总和",
      noChartData: "暂无足够数据可显示",
      totalOt: "加班总数",
      hours: "小时",
    },
    att: {
      checkIn: "上班打卡",
      checkOut: "下班打卡",
      present: "出勤",
      absent: "缺勤",
      late: "迟到",
      onLeave: "请假",
      date: "日期",
      inTime: "上班时间",
      outTime: "下班时间",
      setOffice: "设置办公室位置",
      noRecord: "还没有记录",
      absentDays: "缺勤天数",
      branch: "分店",
      liveClockLabel: "当前时间",
      punchOutBtn: "打卡下班",
      awayFromOffice: (name, dist, radius) =>
        `您距离办公室"${name}" ${dist}m（允许范围 ${radius}m）— 无法打卡`,
      cannotVerifyLoc: "无法验证位置",
      wrongBranchWarning: (name) =>
        `⚠️ 您正在分店"${name}"打卡，但您被分配在其他分店工作`,
      gpsFailed: "无法获取 GPS 位置，请为此应用开启定位服务",
      gpsRequiredHint: (n) => `需要在 ${n} 个分店之一的 GPS 范围内`,
      noData: "没有数据",
      manualEntry: "手动登记",
      dayOffNote: "🛌 今天是您的休息日 — 如果您来上班仍可以打卡",
      officeGeofenceTitle: "多分店 GPS 打卡范围设置",
      officeGeofenceDesc:
        "设置各分店的位置，要求员工自助打卡时必须在其中一个分店附近。员工所在的分店名称会与其考勤记录一起保存。若未添加任何分店，位置限制将不会生效。",
      officeCountLabel: (n) => `${n} 个分店`,
      officeNotSet: "尚未设置",
      officeNoneYet: "还没有分店",
      officeNameLabel: "分店 / 办公室名称",
      officeNamePlaceholder: "总部, 分店...",
      officeRadiusLabel: "允许半径（米）",
      officeUseCurrentLoc: "使用当前位置",
      officeGpsFailed: "无法获取当前 GPS 位置",
      officeNameRequired: "请输入分店/办公室名称",
      officeCoordsRequired: "请正确填写坐标和半径",
      officeAddBtn: "添加分店",
      officeEditTitle: "编辑分店",
      officeConfirmDel: (name) => `确定要删除分店"${name}"吗？`,
      statusPresent: "出勤",
      statusLate: "迟到",
      statusAbsent: "缺勤（未批准）",
      statusLeavePaid: "请假（带薪）",
      scanQrBtn: "扫描分店 QR",
      scanQrOr: "或",
      scanQrTitle: "扫描分店 QR",
      scanQrDesc: "将摄像头对准您分店的 QR code",
      scanQrHint: "正在寻找 QR...",
      scanQrOpeningCamera: "正在打开摄像头...",
      qrNoMatch: "此 QR 不属于任何分店，请再试一次",
      qrExpired: "此 QR 已过期，请扫描分店屏幕/纸张上的最新 QR",
      cameraDenied: "无法打开摄像头，请为此应用开启摄像头权限",
      cameraNotFound: "在此设备上找不到摄像头",
      viaQrLabel: "QR 扫描",
      officeQrBtn: "QR Code",
      officeQrTitle: (name) => `QR Code · ${name}`,
      officeQrDesc:
        "将此屏幕显示在分店入口（例如平板电脑/固定屏幕上），供员工扫描打卡上下班。请勿打印张贴，因为该 QR 刷新后将失效",
      officeQrRefreshHint:
        "为了安全，此 QR 每 20 秒会自动更换 — 请勿截图留存以后使用",
      openKioskBtn: "打开 QR 展示屏幕",
      kioskScanHint: "扫描此码以打卡上/下班",
      kioskRefreshHint: "此码每 20 秒会自动更换",
      kioskNotFoundTitle: "找不到此分店",
      kioskNotFoundDesc: "此链接无效，或该分店已被删除",
    },
    qr: {
      title: "员工 QR Code",
      desc: "扫描以直接打开员工登录页面",
    },
    settings: {
      title: "账号设置",
      photoLabel: "个人头像",
      choosePhoto: "选择照片",
      nameLabel: "姓名",
      namePlaceholder: "请输入您的姓名",
      appearance: "外观",
      lightMode: "浅色 (Light)",
      darkMode: "深色 (Dark)",
      appearanceDesc: "切换此设备的显示外观",
      saved: "保存成功",
      saveFailed: "保存失败：",
      nameRequired: "请输入姓名",
      brandingTitle: "公司品牌",
      brandingDesc: "设置自定义公司名称和标志，以替换默认名称和标志",
      companyNameLabel: "公司名称",
      companyNamePlaceholder: "Workforce Suite",
      companyLogoLabel: "公司标志",
      chooseLogo: "选择标志",
      removeLogo: "移除标志",
      brandingSaved: "品牌信息保存成功",
      empPortalDesc: "将此链接或 QR code 分享给员工，方便他们进入系统",
      soundTitle: "QR 扫描提示音",
      soundDesc: "选择员工成功扫描 QR 打卡上/下班时播放的提示音",
      soundPresets: {
        chime: "铃声（默认）",
        bell: "钟声",
        marimba: "马林巴琴声",
        pop: "短促音",
        classic: "经典提示音",
        silent: "静音",
      },
      soundPreview: "试听",
      soundSaved: "提示音保存成功",
      pushTitle: "推送通知",
      pushDesc: "开启后，即使未打开本应用，也可在此设备上收到即时通知",
      pushEnable: "开启通知",
      pushDisable: "关闭此设备的通知",
      pushEnabledMsg: "此设备的推送通知已开启",
      pushBlocked: "此网站的通知已被屏蔽，请在浏览器的网站权限中调整",
      pushError: "无法开启通知：",
    },
  },
};

// Fills in any key missing from `override` using the matching key from
// `base`, recursively — used so a partial translation (currently `zh`)
// never shows "undefined" or crashes: anything not yet translated just
// reads in English until someone adds it above. Functions/arrays/other
// non-plain-object leaves are taken whole from whichever side has them.
function mergeLangFallback(override, base) {
  if (
    typeof base !== "object" ||
    base === null ||
    Array.isArray(base) ||
    typeof override === "function"
  ) {
    return override === undefined ? base : override;
  }
  if (typeof override !== "object" || override === null) {
    return override === undefined ? base : override;
  }
  const result = {};
  for (const key of Object.keys(base)) {
    result[key] = mergeLangFallback(override[key], base[key]);
  }
  for (const key of Object.keys(override)) {
    if (!(key in result)) result[key] = override[key];
  }
  return result;
}

const LANG = {
  km: LANG_RAW.km,
  en: LANG_RAW.en,
  zh: mergeLangFallback(LANG_RAW.zh, LANG_RAW.en),
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

// Small inline SVG flags for LangToggle — plain emoji flags (🇬🇧🇰🇭🇨🇳)
// render as two-letter text codes on Windows (no color flag glyphs in
// the default font there, unlike macOS/iOS), so we draw them ourselves
// for a consistent look on every OS.
function FlagIcon({ code }) {
  const common = {
    width: 18,
    height: 13,
    style: { flexShrink: 0, borderRadius: 2 },
  };
  if (code === "gb") {
    return (
      <svg {...common} viewBox="0 0 60 40">
        <rect width="60" height="40" fill="#00247d" />
        <path d="M0,0 L60,40 M60,0 L0,40" stroke="#fff" strokeWidth="8" />
        <path d="M0,0 L60,40 M60,0 L0,40" stroke="#cf142b" strokeWidth="3" />
        <path d="M30,0 V40 M0,20 H60" stroke="#fff" strokeWidth="12" />
        <path d="M30,0 V40 M0,20 H60" stroke="#cf142b" strokeWidth="7" />
      </svg>
    );
  }
  if (code === "kh") {
    return (
      <svg {...common} viewBox="0 0 60 40">
        <rect width="60" height="40" fill="#032ea1" />
        <rect y="9" width="60" height="22" fill="#e00025" />
        <rect x="21" y="13" width="18" height="14" fill="#fff" />
      </svg>
    );
  }
  if (code === "cn") {
    return (
      <svg {...common} viewBox="0 0 60 40">
        <rect width="60" height="40" fill="#de2910" />
        <polygon
          points="11,7 13,13 19,13 14,17 16,23 11,19 6,23 8,17 3,13 9,13"
          fill="#ffde00"
        />
      </svg>
    );
  }
  return null;
}

// Language options for LangToggle, in display order. `flag` is the code
// FlagIcon draws — see above.
const LANG_OPTIONS = [
  { code: "en", flag: "gb", label: "English" },
  { code: "km", flag: "kh", label: "ខ្មែរ" },
  { code: "zh", flag: "cn", label: "中文" },
];

function LangToggle({ variant = "dark" }) {
  const { lang, setLang } = useLang();
  const isDark = variant === "dark";
  const [open, setOpen] = useState(false);
  const boxRef = useRef(null);
  const current = LANG_OPTIONS.find((o) => o.code === lang) || LANG_OPTIONS[0];

  useEffect(() => {
    function onDocClick(e) {
      if (boxRef.current && !boxRef.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, []);

  return (
    <div ref={boxRef} style={{ position: "relative" }}>
      <button
        onClick={() => setOpen((v) => !v)}
        title="Switch language / ប្តូរភាសា / 切换语言"
        style={{
          display: "flex",
          alignItems: "center",
          gap: 6,
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
        <FlagIcon code={current.flag} />
        <span>{current.label}</span>
        <ChevronLeft
          size={12}
          style={{
            transform: open ? "rotate(90deg)" : "rotate(-90deg)",
            transition: "transform .15s",
            opacity: 0.7,
          }}
        />
      </button>
      {open && (
        <div
          style={{
            position: "absolute",
            top: "calc(100% + 6px)",
            right: 0,
            minWidth: 130,
            background: T.paper,
            border: `1px solid ${T.line}`,
            borderRadius: 10,
            boxShadow: "0 12px 30px rgba(0,0,0,0.18)",
            padding: 4,
            zIndex: 60,
          }}
        >
          {LANG_OPTIONS.map((o) => (
            <button
              key={o.code}
              onClick={() => {
                setLang(o.code);
                setOpen(false);
              }}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                width: "100%",
                background: o.code === lang ? T.forestSoft : "transparent",
                border: "none",
                borderRadius: 7,
                color: T.ink,
                fontSize: 13,
                fontWeight: o.code === lang ? 700 : 500,
                padding: "7px 10px",
                cursor: "pointer",
                textAlign: "left",
              }}
            >
              <FlagIcon code={o.flag} />
              <span>{o.label}</span>
            </button>
          ))}
        </div>
      )}
    </div>
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
      <span
        style={{
          width: 6,
          height: 6,
          borderRadius: 999,
          background: T.forest,
          flexShrink: 0,
        }}
      />
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
  "manageAssets",
  "manageRecruitment",
  "canMessage",
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
    manageAssets: "គ្រប់គ្រងទ្រព្យសម្បត្តិក្រុមហ៊ុន",
    manageRecruitment: "គ្រប់គ្រងការជ្រើសរើសនិងចាប់ផ្តើមការងារ",
    canMessage: "សារ & ការហៅសំឡេងទៅបុគ្គលិក",
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
    manageAssets: "Manage Company Assets",
    manageRecruitment: "Manage Recruitment & Onboarding",
    canMessage: "Message & Voice-Call Employees",
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
  "training",
  "attcorr",
  "shiftswap",
  "documents",
  "loginActivity",
  "profile",
  "messages",
];
const EMPLOYEE_MODULE_LABEL = {
  km: {
    announcements: "សេចក្តីប្រកាស",
    attendance: "វត្តមានផ្ទាល់ខ្លួន",
    leave: "សំណើច្បាប់ឈប់សម្រាក",
    ot: "ម៉ោងបន្ថែម (OT)",
    payroll: "ប្រាក់ខែផ្ទាល់ខ្លួន",
    review: "ការវាយតម្លៃការងារ",
    training: "ការបណ្តុះបណ្តាល & សញ្ញាបត្រ",
    attcorr: "សំណើកែតម្រូវវត្តមាន",
    shiftswap: "សំណើដូរវេន",
    documents: "ឯកសារផ្ទាល់ខ្លួន",
    loginActivity: "សកម្មភាពចូលប្រើ",
    profile: "ប្រវត្តិរូបផ្ទាល់ខ្លួន",
    messages: "សារ & ការហៅសំឡេង",
  },
  en: {
    announcements: "Announcements",
    attendance: "My Attendance",
    leave: "My Leave",
    ot: "My Overtime (OT)",
    payroll: "My Payroll",
    review: "My Performance Reviews",
    training: "My Training & Certifications",
    attcorr: "Attendance Correction",
    shiftswap: "Shift Swap Request",
    documents: "My Documents",
    loginActivity: "Login Activity",
    profile: "My Profile",
    messages: "Messages & Voice Call",
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
  manageRecruitment: false,
  announcements: true,
  attendance: true,
  leave: true,
  ot: true,
  payroll: true,
  review: true,
  training: true,
  attcorr: true,
  shiftswap: true,
  documents: true,
  loginActivity: true,
  profile: true,
  messages: true,
};
// Sensible starting point covering the natural HR seniority ladder — an
// Officer can only view, each step up adds more, Admin gets everything
// short of managing other admin accounts. Superadmin edits this freely
// from the Roles & Permissions page; these are just the seed defaults
// used until Superadmin has saved a custom matrix (and as a fallback for
// any rank the matrix doesn't yet have a row for).
const DEFAULT_ROLE_PERMISSIONS = {
  // canMessage defaults to true on every rank (unlike the other keys
  // above, which default to false for the lower ranks) so that shipping
  // this permission doesn't silently take Messages/Call away from any
  // admin who already had it — same "starts on, Superadmin opts out"
  // philosophy as DEFAULT_EMPLOYEE_MODULES below.
  staff: {
    manageDepartments: false,
    manageEmployees: false,
    approveRequests: false,
    managePayroll: false,
    manageDocuments: false,
    manageAnnouncements: false,
    manageSettings: false,
    viewAuditLog: false,
    manageAssets: false,
    manageRecruitment: false,
    canMessage: true,
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
    manageAssets: false,
    manageRecruitment: false,
    canMessage: true,
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
    manageAssets: true,
    manageRecruitment: true,
    canMessage: true,
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
    manageAssets: true,
    manageRecruitment: true,
    canMessage: true,
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
    manageAssets: true,
    manageRecruitment: true,
    canMessage: true,
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
    manageAssets: true,
    manageRecruitment: true,
    canMessage: true,
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
    manageAssets: true,
    manageRecruitment: true,
    canMessage: true,
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
.wf-dp-day{position:relative;font-size:12.5px;text-align:center;padding:7px 0;border-radius:6px;cursor:pointer;color:${T.text};background:none;border:1px solid transparent;transition:background .12s ease,color .12s ease;}
.wf-dp-day:hover{background:${T.tableHeadBg};}
.wf-dp-day.outside{color:${T.mutedLight};}
.wf-dp-day.today{border-color:${T.gold};font-weight:700;}
.wf-dp-day.selected{background:${T.gold};color:#1A1300;font-weight:700;}
.wf-dp-day.selected:hover{background:#D89430;}
.wf-dp-foot{display:flex;align-items:center;justify-content:space-between;margin-top:10px;padding-top:10px;border-top:1px solid ${T.lineSoft};}
.wf-dp-link{background:none;border:none;cursor:pointer;font-size:12px;font-weight:600;color:${T.forestText};padding:2px 4px;border-radius:6px;}
.wf-dp-link:hover{background:${T.forestSoft};}
.wf-dr-triggers{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
.wf-dp-pop-wide{width:520px;max-width:calc(100vw - 64px);padding:14px;}
.wf-grid-2{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
@media (max-width:480px){.wf-grid-2{grid-template-columns:1fr;}}
.wf-dp-head-dual{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:10px;}
.wf-dp-dual-titles{display:flex;flex:1;justify-content:space-around;}
.wf-dp-dual{display:flex;gap:22px;}
.wf-dp-panel{flex:1;min-width:0;}
.wf-dp-grid-range{gap:0;row-gap:3px;}
.wf-dp-day-num{display:inline-flex;align-items:center;justify-content:center;width:24px;height:24px;border-radius:50%;position:relative;z-index:1;margin:0 auto;transition:background .12s ease,color .12s ease;}
.wf-dp-day.in-range::before,.wf-dp-day.range-start::before,.wf-dp-day.range-end::before{content:'';position:absolute;top:3px;bottom:3px;left:0;right:0;background:${T.forestSoft};z-index:0;}
.wf-dp-day.range-start::before{left:50%;}
.wf-dp-day.range-end::before{right:50%;}
.wf-dp-day.range-single::before{content:none;}
.wf-dp-day.range-start .wf-dp-day-num,.wf-dp-day.range-end .wf-dp-day-num,.wf-dp-day.range-single .wf-dp-day-num{background:${T.forest};color:#fff;font-weight:700;}
.wf-dp-day:hover .wf-dp-day-num{background:${T.tableHeadBg};}
.wf-dp-day.range-start:hover .wf-dp-day-num,.wf-dp-day.range-end:hover .wf-dp-day-num,.wf-dp-day.range-single:hover .wf-dp-day-num{background:${T.forestDark};}
@media (max-width:640px){
  .wf-dp-pop-wide{min-width:0;width:min(94vw,320px);padding:12px;}
  .wf-dp-dual{flex-direction:column;gap:14px;}
  .wf-dp-dual-titles{flex-direction:column;align-items:center;gap:2px;}
}
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
.wf-modal{background:${T.card};border-radius:11px;border:1px solid ${T.line};box-shadow:0 24px 64px rgba(3,5,10,0.5);width:100%;max-height:90vh;overflow-y:auto;overflow-x:hidden;animation:wf-pop .18s cubic-bezier(.2,.9,.3,1.2);}
.wf-modal-head{display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid ${T.lineSoft};position:sticky;top:0;z-index:5;background:${T.card};border-radius:11px 11px 0 0;}
.wf-avatar{border-radius:999px;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:600;flex-shrink:0;}
.wf-badge{display:inline-block;padding:4px 10px;border-radius:6px;font-size:11px;font-weight:700;letter-spacing:.02em;white-space:nowrap;}
.wf-table{width:100%;min-width:640px;font-size:13px;border-collapse:collapse;}
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
@keyframes wf-rec-pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.35;transform:scale(.75)}}
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
  .wf-chat-layout{height:calc(100vh - 190px) !important;}
  .wf-chat-list-pane{width:100% !important;border-right:none !important;}
  .wf-chat-layout.wf-chat-thread-open .wf-chat-list-pane{display:none;}
  .wf-chat-layout:not(.wf-chat-thread-open) .wf-chat-thread-pane{display:none;}
  .wf-chat-back-btn{display:inline-flex !important;}
}
.wf-chat-layout{display:flex;gap:0;height:calc(100vh - 230px);min-height:420px;border:1px solid ${T.lineSoft};border-radius:14px;overflow:hidden;background:${T.card};}
.wf-chat-list-pane{width:300px;flex-shrink:0;display:flex;flex-direction:column;border-right:1px solid ${T.lineSoft};overflow-y:auto;}
.wf-chat-thread-pane{flex:1;display:flex;flex-direction:column;min-width:0;background:${T.paper};}
.wf-chat-item{display:flex;gap:10px;width:100%;text-align:left;padding:12px 14px;background:transparent;border:none;border-bottom:1px solid ${T.lineSoft};cursor:pointer;align-items:center;}
.wf-chat-item:hover{background:${T.paper};}
.wf-chat-item.active{background:${T.forestSoft};}
.wf-chat-bubble-row{padding:2px 14px;}
.wf-chat-bubble{max-width:100%;padding:9px 13px;border-radius:14px;font-size:13px;line-height:1.5;overflow-wrap:break-word;word-break:normal;white-space:pre-wrap;}
.wf-chat-back-btn{display:none;background:none;border:none;cursor:pointer;color:${T.ink};padding:6px;align-items:center;}
.wf-callscreen{position:fixed;inset:0;z-index:9999;display:flex;flex-direction:column;align-items:center;justify-content:space-between;background:radial-gradient(circle at 50% 18%,#243158 0%,#0a0e1a 72%);color:#fff;padding:max(28px,env(safe-area-inset-top)) 24px max(32px,env(safe-area-inset-bottom));animation:wf-fade .22s ease;}
.wf-callscreen-ring{position:absolute;top:50%;left:50%;border-radius:50%;border:2px solid rgba(255,255,255,0.25);transform:translate(-50%,-50%);animation:wf-pulse-ring 2.4s ease-out infinite;}
.wf-callscreen-dots span{animation:wf-calldot 1.4s infinite ease-in-out;display:inline-block;}
.wf-callscreen-dots span:nth-child(2){animation-delay:.2s;}
.wf-callscreen-dots span:nth-child(3){animation-delay:.4s;}
@keyframes wf-calldot{0%,80%,100%{opacity:.25;}40%{opacity:1;}}
.wf-callbtn{display:flex;flex-direction:column;align-items:center;gap:9px;background:none;border:none;cursor:pointer;color:#fff;padding:0;}
.wf-callbtn-circle{width:62px;height:62px;border-radius:50%;display:flex;align-items:center;justify-content:center;transition:transform .12s ease,background .15s ease;box-shadow:0 8px 20px rgba(0,0,0,0.3);}
.wf-callbtn:active .wf-callbtn-circle{transform:scale(.92);}
.wf-callbtn-circle.wf-callbtn-accept{background:#1FA26B;}
.wf-callbtn-circle.wf-callbtn-reject{background:#e5484d;}
.wf-callbtn-circle.wf-callbtn-secondary{background:rgba(255,255,255,0.14);border:1px solid rgba(255,255,255,0.22);box-shadow:none;}
.wf-callbtn-circle.wf-callbtn-secondary.wf-callbtn-active{background:#fff;color:#0a0e1a;}
.wf-callbtn-label{font-size:12.5px;color:rgba(255,255,255,0.75);font-weight:600;}
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
  // Late-arrival deduction: an employee can be late up to `lateGraceCount`
  // times per month with no penalty. Every late day beyond that grace
  // count is deducted, either as a fixed $ amount per day ("fixed") or
  // as a percentage of that employee's daily rate ("percentDaily").
  lateGraceCount: 3,
  lateDeductionType: "fixed",
  lateDeductionValue: 0,
  // Unpaid Leave (UL) deduction: "fullDay" docks a full day's pay
  // (salary/WORKING_DAYS_PER_MONTH) per UL day — this is the original
  // behavior. "fixed" docks a flat $ amount per UL day; "percentDaily"
  // docks a percentage of the employee's daily rate per UL day.
  ulDeductionType: "fullDay",
  ulDeductionValue: 0,
  // Tax/insurance calculation mode: "flat" uses the taxRate/insuranceRate
  // percentages above (original behavior, applied to any currency).
  // "khmerProgressive" instead computes Cambodia's official Tax on
  // Salary (ToS) progressive brackets and NSSF contributions, converting
  // the USD base salary to KHR using exchangeRate. taxRate/insuranceRate
  // are ignored in this mode.
  taxMode: "flat",
  // KHR per 1 USD, used only in "khmerProgressive" mode. Update this to
  // track the official National Bank of Cambodia / GDT rate — it drifts
  // over time and isn't something this app can look up on its own.
  exchangeRate: 4100,
  // NSSF contributions are calculated on the employee's wage only up to
  // this cap (KHR), per NSSF's contributory-wage rules.
  nssfWageCapKHR: 1200000,
  // Pension scheme (Stage 1): shared 4% of contributory wage, split
  // evenly. The employee half is what actually reduces net pay; the
  // employer half is shown for cost transparency only.
  nssfPensionEmployeeRate: 2,
  nssfPensionEmployerRate: 2,
  // Occupational Risk and Healthcare contributions are, by current NSSF
  // rules, paid entirely by the employer — they never reduce an
  // employee's net pay, but are tracked so the employer's true cost is
  // visible on the payslip.
  nssfOrcRate: 0.8,
  nssfHealthRate: 2.6,
};
// Cambodia's monthly Tax on Salary (ToS) brackets for resident employees,
// per Sub-Decree No. 48 ANKr.BK (11 March 2024) / GDT Instruction 017.
// Amounts are in KHR. Tiered: each portion of taxable salary is taxed at
// the rate for that bracket only, not the whole salary at one rate.
// NOTE: this is current as of this app's release but Cambodian tax law
// changes by sub-decree from time to time — double check against the
// General Department of Taxation before relying on this for compliance,
// and adjust the brackets here if the law changes.
const KHMER_TOS_BRACKETS = [
  { upTo: 1500000, rate: 0 },
  { upTo: 2000000, rate: 0.05 },
  { upTo: 8500000, rate: 0.1 },
  { upTo: 12500000, rate: 0.15 },
  { upTo: Infinity, rate: 0.2 },
];
// Dependent deduction (child under 14 / in full-time study up to 25, or
// a non-earning spouse), per Sub-Decree 48 Article 3.2.
const KHMER_DEPENDENT_DEDUCTION_KHR = 150000;
// Computes Cambodia's progressive monthly Tax on Salary for a given
// taxable amount (already net of dependent deductions), in KHR.
function khmerSalaryTaxKHR(taxableKHR) {
  const t = Math.max(0, taxableKHR);
  let tax = 0;
  let lower = 0;
  for (const bracket of KHMER_TOS_BRACKETS) {
    if (t <= lower) break;
    const upper = Math.min(t, bracket.upTo);
    tax += Math.max(0, upper - lower) * bracket.rate;
    lower = bracket.upTo;
  }
  return tax;
}
// Computes NSSF contributions for one employee, in USD, given their base
// salary (USD), the policy's exchange rate/cap/rates, and their
// dependent count. Returns both the employee-deducted pension share and
// the employer-only shares (pension top-up, occupational risk, health).
function computeKhmerNssf(salaryUSD, policy) {
  const rate = Number(policy.exchangeRate) || 4100;
  const capKHR = Number(policy.nssfWageCapKHR) || 1200000;
  const wageKHR = Math.min((Number(salaryUSD) || 0) * rate, capKHR);
  const pensionEmployeeKHR =
    wageKHR * ((Number(policy.nssfPensionEmployeeRate) || 0) / 100);
  const pensionEmployerKHR =
    wageKHR * ((Number(policy.nssfPensionEmployerRate) || 0) / 100);
  const orcKHR = wageKHR * ((Number(policy.nssfOrcRate) || 0) / 100);
  const healthKHR = wageKHR * ((Number(policy.nssfHealthRate) || 0) / 100);
  return {
    pensionEmployee: pensionEmployeeKHR / rate,
    pensionEmployer: pensionEmployerKHR / rate,
    orc: orcKHR / rate,
    health: healthKHR / rate,
    employerTotal: (pensionEmployerKHR + orcKHR + healthKHR) / rate,
  };
}
// Named tone recipes for the QR check-in/out chime (see playScanBeep).
// Each preset gives a short tone sequence for "in" and a separate one
// for "out" so the two stay distinguishable by ear. Frequencies are in
// Hz; `type` is the oscillator waveform. "silent" is the mute option.
const SOUND_PRESETS = {
  chime: {
    in: [
      { freq: 880, start: 0, dur: 0.09, type: "sine" },
      { freq: 1318.5, start: 0.1, dur: 0.15, type: "sine" },
    ],
    out: [
      { freq: 987.77, start: 0, dur: 0.09, type: "sine" },
      { freq: 659.25, start: 0.1, dur: 0.16, type: "sine" },
    ],
  },
  bell: {
    in: [{ freq: 1567.98, start: 0, dur: 0.4, type: "triangle" }],
    out: [{ freq: 1046.5, start: 0, dur: 0.4, type: "triangle" }],
  },
  marimba: {
    in: [
      { freq: 659.25, start: 0, dur: 0.13, type: "sine" },
      { freq: 987.77, start: 0.09, dur: 0.2, type: "sine" },
    ],
    out: [
      { freq: 587.33, start: 0, dur: 0.13, type: "sine" },
      { freq: 440, start: 0.09, dur: 0.2, type: "sine" },
    ],
  },
  pop: {
    in: [{ freq: 1200, start: 0, dur: 0.06, type: "square" }],
    out: [{ freq: 500, start: 0, dur: 0.06, type: "square" }],
  },
  classic: {
    in: [{ freq: 1000, start: 0, dur: 0.13, type: "sine" }],
    out: [{ freq: 1000, start: 0, dur: 0.13, type: "sine" }],
  },
  silent: { in: [], out: [] },
};
const DEFAULT_SOUND_POLICY = { preset: "chime" };
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
// Chat date-separator label ("Today" / "Yesterday" / a localized date),
// used to break up a message thread by day the way most chat apps do.
// Compares calendar days (not 24h windows) so a message sent at 11:58pm
// and one sent at 12:02am the next day fall on either side of the line.
function chatDateSeparatorLabel(iso, lang, t) {
  if (!iso) return "";
  const d = new Date(iso);
  const now = new Date();
  const startOfDay = (date) =>
    new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const diffDays = Math.round((startOfDay(now) - startOfDay(d)) / 86400000);
  if (diffDays === 0) return t.today;
  if (diffDays === 1) return t.yesterday;
  const locale = lang === "km" ? "km-KH" : lang === "zh" ? "zh-CN" : "en-US";
  return d.toLocaleDateString(locale, {
    day: "numeric",
    month: "long",
    year: d.getFullYear() !== now.getFullYear() ? "numeric" : undefined,
  });
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
  shiftSwapRequests = [],
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
    shiftSwapRequests
      .filter((r) => r.status === "pending")
      .forEach((r) => {
        const emp = employees.find((e) => e.id === r.employeeId);
        list.push({
          id: `ss-pending-${r.id}`,
          page: "shiftswap",
          tone: "gold",
          title: en ? "New shift swap request" : "សំណើដូរវេនថ្មី",
          message: en
            ? `${emp?.name || "?"} requested coverage on ${r.date}`
            : `${emp?.name || "?"} បានស្នើសុំគ្របដណ្តប់វេននៅថ្ងៃទី ${r.date}`,
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
    shiftSwapRequests
      .filter(
        (r) =>
          r.employeeId === currentEmp.id &&
          (r.status === "approved" || r.status === "rejected") &&
          r.reviewedAt,
      )
      .forEach((r) => {
        list.push({
          id: `ss-decided-${r.id}`,
          page: "shiftswap",
          tone: r.status === "approved" ? "forest" : "rose",
          title:
            r.status === "approved"
              ? en
                ? "Your shift swap request was approved"
                : "សំណើដូរវេនរបស់អ្នកត្រូវបានអនុម័ត"
              : en
                ? "Your shift swap request was rejected"
                : "សំណើដូរវេនរបស់អ្នកត្រូវបានបដិសេធ",
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
// Payload encoded into an office's printable check-in QR code, and the
// matching verifier. The payload embeds a token that rotates every
// QR_ROTATE_MS derived from each office's own `qrSecret` (generated once
// per office, stored alongside it) — this means a screenshot of the QR
// only works for a short window after it's taken, closing off the easy
// "photograph the QR and text it to a friend who isn't on-site" version
// of buddy-punching that a static QR would allow. It is NOT meant to be
// unbreakable cryptography (the secret lives in the same client bundle
// as everything else in this app) — it raises the bar from "trivial" to
// "would need to be relayed within ~20-40s", which is the realistic goal
// for a self-service kiosk-style check-in.
const OFFICE_QR_PREFIX = "WFOFFICE:";
const QR_ROTATE_MS = 20000;
function fnv1aHash(str) {
  let h = 0x811c9dc5;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(36);
}
function officeQrWindow(windowMs = QR_ROTATE_MS) {
  return Math.floor(Date.now() / windowMs);
}
function officeQrToken(office, win) {
  return fnv1aHash(`${office.qrSecret || ""}:${office.id}:${win}`);
}
function officeQrPayload(office, windowMs = QR_ROTATE_MS) {
  const win = officeQrWindow(windowMs);
  return `${OFFICE_QR_PREFIX}${office.id}:${win}:${officeQrToken(office, win)}`;
}
// Verifies a scanned QR payload against the known offices list and
// reports *why* it failed, so the scanner can show a more useful
// message than a blanket "invalid" (expired vs. simply wrong/unknown).
// Accepts the current time window OR the previous one — a short grace
// period so a QR that rotated a split-second before the scan finished
// still works.
function evaluateOfficeQrPayload(offices, text, windowMs = QR_ROTATE_MS) {
  if (typeof text !== "string" || !text.startsWith(OFFICE_QR_PREFIX))
    return { status: "invalid" };
  const parts = text.slice(OFFICE_QR_PREFIX.length).split(":");
  if (parts.length !== 3) return { status: "invalid" };
  const [officeId, winStr, token] = parts;
  const win = Number(winStr);
  if (!officeId || !token || !Number.isFinite(win))
    return { status: "invalid" };
  const office = (offices || []).find((o) => o.id === officeId);
  if (!office || !office.qrSecret) return { status: "invalid" };
  if (officeQrToken(office, win) !== token) return { status: "invalid" };
  const currentWin = officeQrWindow(windowMs);
  if (win === currentWin || win === currentWin - 1)
    return { status: "ok", office };
  return { status: "expired" };
}
// Convenience wrapper for callers that only care whether it matched.
function verifyOfficeQrPayload(offices, text, windowMs = QR_ROTATE_MS) {
  const result = evaluateOfficeQrPayload(offices, text, windowMs);
  return result.status === "ok" ? result.office : null;
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
  let unpaidLeaveDays = 0;
  for (const a of attendance) {
    if (a.employeeId !== employeeId || !a.date.startsWith(mk)) continue;
    if (a.status === "absent") absentDays++;
    else if (a.status === "leave") leaveDays++;
    else if (a.status === "late") lateDays++;
    else if (a.status === "unpaid") unpaidLeaveDays++;
  }
  return { absentDays, leaveDays, lateDays, unpaidLeaveDays };
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
  const { absentDays, leaveDays, lateDays, unpaidLeaveDays } =
    monthAttendanceStats(attendance, emp.id, mk);
  const dailyRate = emp.salary / WORKING_DAYS_PER_MONTH;
  // Unauthorized absences always dock a full day's pay at the daily rate.
  const absenceDeduction = Math.min(emp.salary, absentDays * dailyRate);
  // Approved Unpaid Leave (UL) is docked per the UL policy in effect:
  // "fullDay" (default, same as before) docks a full day's pay per UL
  // day; "fixed" docks a flat $ amount per UL day; "percentDaily" docks
  // a percentage of the daily rate per UL day. An employee can override
  // the company-wide UL policy with their own, the same way the late
  // policy and tax/insurance rates can be overridden.
  const usesCustomUlPolicy = !!emp.useCustomUlPolicy;
  const ulDeductionType = usesCustomUlPolicy
    ? emp.customUlDeductionType || "fullDay"
    : policy.ulDeductionType || "fullDay";
  const ulDeductionValue = usesCustomUlPolicy
    ? Number(emp.customUlDeductionValue) || 0
    : Number(policy.ulDeductionValue) || 0;
  const ulDeductionPerDay =
    ulDeductionType === "percentDaily"
      ? dailyRate * (ulDeductionValue / 100)
      : ulDeductionType === "fixed"
        ? ulDeductionValue
        : dailyRate; // "fullDay"
  const unpaidLeaveDeduction = Math.min(
    emp.salary - absenceDeduction,
    unpaidLeaveDays * ulDeductionPerDay,
  );
  // Late-arrival deduction: the first `lateGraceCount` late days each month
  // are free; every late day beyond that is docked, either a fixed $
  // amount or a percentage of the daily rate. An employee can override
  // the company-wide late policy with their own grace count/rate, the
  // same way tax/insurance rates can be overridden below.
  const usesCustomLatePolicy = !!emp.useCustomLatePolicy;
  const lateGraceCount = usesCustomLatePolicy
    ? Number(emp.customLateGraceCount) || 0
    : Number(policy.lateGraceCount) || 0;
  const lateDeductionType = usesCustomLatePolicy
    ? emp.customLateDeductionType || "fixed"
    : policy.lateDeductionType || "fixed";
  const lateDeductionValue = usesCustomLatePolicy
    ? Number(emp.customLateDeductionValue) || 0
    : Number(policy.lateDeductionValue) || 0;
  const excessLateDays = Math.max(0, lateDays - lateGraceCount);
  const lateDeductionPerDay =
    lateDeductionType === "percentDaily"
      ? dailyRate * (lateDeductionValue / 100)
      : lateDeductionValue;
  const lateDeduction = Math.min(
    emp.salary - absenceDeduction - unpaidLeaveDeduction,
    excessLateDays * lateDeductionPerDay,
  );
  const adjustedBase = emp.salary - absenceDeduction - unpaidLeaveDeduction;
  // An employee can override the company-wide tax/insurance rates with
  // their own. When active, the override applies unconditionally (the
  // policy's minimum-salary threshold is only meant to gate the
  // *default* rates, not an explicit per-employee rate).
  const usesCustomRate = !!emp.useCustomRate;
  const usesKhmerMode =
    policy.taxMode === "khmerProgressive" && !usesCustomRate;
  let minSalaryThreshold,
    deductionApplies,
    taxRate,
    insuranceRate,
    tax,
    insurance,
    khmer = null;
  if (usesKhmerMode) {
    // Cambodia mode: NSSF pension (employee share) always applies, and
    // the Tax on Salary is computed from the official progressive
    // brackets — the minSalaryThreshold setting doesn't apply here since
    // the brackets already have their own tax-free floor (1.5M KHR).
    minSalaryThreshold = 0;
    deductionApplies = true;
    const nssf = computeKhmerNssf(adjustedBase, policy);
    const exchangeRate = Number(policy.exchangeRate) || 4100;
    const dependents = Math.max(0, Number(emp.dependents) || 0);
    const dependentDeductionKHR = dependents * KHMER_DEPENDENT_DEDUCTION_KHR;
    const grossKHR = adjustedBase * exchangeRate;
    // Employee pension contributions are deductible from taxable salary.
    const taxableKHR = Math.max(
      0,
      grossKHR - dependentDeductionKHR - nssf.pensionEmployee * exchangeRate,
    );
    const taxKHR = khmerSalaryTaxKHR(taxableKHR);
    tax = taxKHR / exchangeRate;
    insurance = nssf.pensionEmployee;
    taxRate =
      adjustedBase > 0 ? Math.round((tax / adjustedBase) * 1000) / 10 : 0;
    insuranceRate = Number(policy.nssfPensionEmployeeRate) || 0;
    khmer = {
      exchangeRate,
      dependents,
      dependentDeductionKHR,
      taxableKHR,
      nssfPensionEmployer: nssf.pensionEmployer,
      nssfOrc: nssf.orc,
      nssfHealth: nssf.health,
      employerNssfTotal: nssf.employerTotal,
    };
  } else if (usesCustomRate) {
    minSalaryThreshold = 0;
    deductionApplies = true;
    taxRate = Number(emp.customTaxRate) || 0;
    insuranceRate = Number(emp.customInsuranceRate) || 0;
    tax = adjustedBase * (taxRate / 100);
    insurance = adjustedBase * (insuranceRate / 100);
  } else {
    minSalaryThreshold = Number(policy.minSalaryThreshold) || 0;
    // Tax/insurance only kick in once the employee's base salary reaches the
    // configured threshold. Below it, no deduction is applied at all.
    deductionApplies = (Number(emp.salary) || 0) >= minSalaryThreshold;
    taxRate = deductionApplies ? Number(policy.taxRate) || 0 : 0;
    insuranceRate = deductionApplies ? Number(policy.insuranceRate) || 0 : 0;
    tax = adjustedBase * (taxRate / 100);
    insurance = adjustedBase * (insuranceRate / 100);
  }
  const { otHours, otPay } = computeOvertimeForMonth(
    emp,
    overtimeRequests,
    mk,
    otPolicy,
  );
  const net = adjustedBase - tax - insurance + otPay - lateDeduction;
  return {
    absentDays,
    leaveDays,
    lateDays,
    unpaidLeaveDays,
    dailyRate,
    absenceDeduction,
    unpaidLeaveDeduction,
    usesCustomUlPolicy,
    ulDeductionType,
    ulDeductionValue,
    ulDeductionPerDay,
    usesCustomLatePolicy,
    lateGraceCount,
    excessLateDays,
    lateDeductionType,
    lateDeductionValue,
    lateDeduction,
    adjustedBase,
    tax,
    insurance,
    taxRate,
    insuranceRate,
    minSalaryThreshold,
    deductionApplies,
    usesCustomRate,
    usesKhmerMode,
    khmer,
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
  const portal =
    portalPart === "employee"
      ? "employee"
      : portalPart === "kiosk"
        ? "kiosk"
        : "admin";
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
   Global toast notifications — a lightweight pub/sub that lives
   outside React state, so any hook or async closure (in particular
   useSupabaseArray's save/delete handlers below, which run far from
   the render tree and have no props to receive a dispatcher through)
   can surface a toast without prop-drilling. <ToastHost/> — mounted
   once near the app root — is the sole subscriber and owns the
   on-screen list. This is what replaces the old silent-failure
   pattern (a write rejected by Supabase used to only reach
   console.error while the UI kept showing the optimistic local
   change, e.g. the Training module toggle that looked saved but
   wasn't).
----------------------------------------------------------------*/
let toastSeq = 0;
const toastSubscribers = new Set();
function pushToast(message, type = "success") {
  if (!message) return;
  const toast = { id: ++toastSeq, message, type };
  toastSubscribers.forEach((fn) => fn(toast));
}
function ToastHost() {
  const [toasts, setToasts] = useState([]);
  useEffect(() => {
    const onToast = (toast) => {
      setToasts((cur) => [...cur, toast]);
      const life = toast.type === "error" ? 6000 : 3000;
      setTimeout(() => {
        setToasts((cur) => cur.filter((x) => x.id !== toast.id));
      }, life);
    };
    toastSubscribers.add(onToast);
    return () => toastSubscribers.delete(onToast);
  }, []);
  const dismiss = (id) => setToasts((cur) => cur.filter((x) => x.id !== id));
  if (!toasts.length) return null;
  const accent = (type) =>
    type === "error" ? T.rose : type === "info" ? T.blue : T.forest;
  return (
    <div
      style={{
        position: "fixed",
        top: 84,
        right: 20,
        zIndex: 9999,
        display: "flex",
        flexDirection: "column",
        gap: 10,
        alignItems: "flex-end",
        pointerEvents: "none",
        width: "min(92vw, 360px)",
      }}
    >
      {toasts.map((toast) => {
        const c = accent(toast.type);
        return (
          <div
            key={toast.id}
            role="status"
            style={{
              pointerEvents: "auto",
              display: "flex",
              alignItems: "flex-start",
              gap: 10,
              width: "100%",
              padding: "12px 14px",
              borderRadius: 12,
              fontSize: 13,
              fontWeight: 500,
              color: T.text,
              lineHeight: 1.45,
              background: T.card,
              border: `1px solid ${T.lineSoft}`,
              borderLeft: `3px solid ${c}`,
              boxShadow: "0 10px 30px rgba(0,0,0,0.14)",
              animation: "wfToastIn 0.24s ease-out",
            }}
          >
            <div
              style={{
                flexShrink: 0,
                width: 22,
                height: 22,
                borderRadius: "50%",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                background: `${c}1f`,
                color: c,
                marginTop: 1,
              }}
            >
              {toast.type === "error" ? (
                <AlertCircle size={13} />
              ) : (
                <CheckCircle2 size={13} />
              )}
            </div>
            <span style={{ flex: 1, paddingTop: 2 }}>{toast.message}</span>
            <button
              onClick={() => dismiss(toast.id)}
              aria-label="Dismiss"
              style={{
                flexShrink: 0,
                background: "none",
                border: "none",
                cursor: "pointer",
                color: T.muted,
                padding: 2,
                display: "flex",
                borderRadius: 6,
                marginTop: -1,
              }}
            >
              <X size={14} />
            </button>
          </div>
        );
      })}
      <style>{`
        @keyframes wfToastIn {
          from { opacity: 0; transform: translateX(16px); }
          to { opacity: 1; transform: translateX(0); }
        }
      `}</style>
    </div>
  );
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
// but this app calls useSupabaseArray for 14 different tables on every
// login (departments, employees, shifts, attendance, leave_requests,
// overtime_requests, performance_reviews, announcements,
// employee_documents, holidays, attendance_corrections,
// shift_swap_requests, admins, offices), all mounted together in
// AppInner regardless of which page is showing. That's 14 open
// channels per signed-in user — at 1000+
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
  "shift_swap_requests",
  "admins",
  "offices",
  "payroll_paid",
  "role_permissions",
  "assets",
  "trainings",
  "messages",
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
// ---------------------------------------------------------------
// Voice-call signaling — a single shared Supabase Realtime *broadcast*
// channel (separate from the postgres_changes bus above, since calls are
// never persisted). Every signaling message carries the employeeId it
// belongs to — same idea as messages.employee_id — so each client can
// tell whether a given offer/answer/ICE candidate is about its own
// thread. `broadcast: { self: false }` means a sender never receives its
// own message back, so handlers below don't need to filter those out.
//
// This is intentionally a flat, unauthenticated broadcast (like the rest
// of this app's realtime tables rely on table-level RLS rather than
// channel-level auth) — good enough for a first version, but a
// production deployment with many admins should look at Supabase
// Realtime Authorization (per-channel RLS policies) so an employee's
// browser can't technically see signaling metadata for other employees'
// calls even though it already ignores it.
let callSignalChannel = null;
const callSignalHandlers = new Set();
function ensureCallSignalChannel() {
  if (callSignalChannel) return callSignalChannel;
  try {
    const channel = supabase.channel("realtime:calls", {
      config: { broadcast: { self: false } },
    });
    channel.on("broadcast", { event: "signal" }, ({ payload }) => {
      callSignalHandlers.forEach((fn) => fn(payload));
    });
    channel.subscribe();
    callSignalChannel = channel;
  } catch (err) {
    console.error("[supabase] failed to open call signal channel:", err);
  }
  return callSignalChannel;
}
function subscribeCallSignal(handler) {
  ensureCallSignalChannel();
  callSignalHandlers.add(handler);
  return () => callSignalHandlers.delete(handler);
}
function sendCallSignal(payload) {
  const channel = ensureCallSignalChannel();
  if (!channel) return;
  channel.send({ type: "broadcast", event: "signal", payload });
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
    // Optional: (evt) => pushPayload | null, called once per created row
    // (evt = { type: "create", row }) and once per updated row (evt =
    // { type: "update", row, old }). Return a push_notify request body
    // (userType, userId?, title, body, page, portal, tag) to fire a real
    // Web Push for that change, or null/undefined to skip. Best-effort:
    // failures are logged, never surfaced to the user or thrown, since a
    // push failing shouldn't block the save that triggered it.
    notify,
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
  // Every table gets a toast for free (see pushToast/ToastHost above),
  // regardless of whether its screen also wires up the saveError banner
  // below — so a write that silently fails on a page with no bespoke
  // error UI (which, before this, was most pages) still tells the user.
  const { t } = useLang();
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
    (nextOrUpdater) => {
      const prev = prevRef.current;
      const next =
        typeof nextOrUpdater === "function"
          ? nextOrUpdater(prev)
          : nextOrUpdater;
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
        let hadError = false;
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
            pushToast(`${t.settings.saveFailed} ${error.message}`, "error");
            hadError = true;
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
            pushToast(`${t.settings.saveFailed} ${error.message}`, "error");
            hadError = true;
          }
        }
        // Confirms the optimistic update the UI already made actually
        // stuck server-side. Only fires when something real changed
        // (toDelete/toUpsert non-empty) — not on every re-render — so
        // it stays a genuine confirmation rather than background noise.
        if (!hadError && (toDelete.length || toUpsert.length)) {
          pushToast(t.settings.saved, "success");
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

      if (notify) {
        const sendPush = (body) => {
          if (!body) return;
          supabase.functions.invoke("push_notify", { body }).then(
            ({ error }) => {
              if (error)
                console.error(
                  `[push] send failed for ${table}:`,
                  error.message,
                );
            },
            (err) => console.error(`[push] send failed for ${table}:`, err),
          );
        };
        // Only the admin-facing case (a brand-new pending request that
        // needs someone to act on it) gets forwarded to Telegram — the
        // employee-facing "your request was approved/rejected" case
        // stays as an in-app/push notification, since there's no
        // per-employee Telegram chat to send it to. The edge function
        // re-checks the enabled flag and the per-category toggle itself
        // (using `category`, which matches `body.page`: "leave"/"ot"/
        // "attcorr"/"shiftswap"), so this call is safe to fire even when
        // Telegram isn't configured at all — it just no-ops server-side.
        // `body.skipTelegram` is an escape hatch for notify() callbacks
        // (e.g. new chat messages) that want push notifications to keep
        // working but need to opt OUT of Telegram specifically based on
        // a toggle the edge function doesn't know about — checked here
        // on the client rather than trusting the server side to filter
        // a category it may not recognize.
        const sendTelegram = (body) => {
          if (!body || body.userType !== "admin" || body.skipTelegram) return;
          supabase.functions
            .invoke("telegram_notify", {
              body: {
                text: `${body.title}\n${body.body}`,
                category: body.page,
                // entityId intentionally omitted for now — leaving it
                // out means telegram_notify won't attach the Approve/
                // Reject buttons, since tapping them needs the
                // telegram_webhook function (not set up yet). Add
                // `entityId: body.entityId,` back once that's deployed
                // and its webhook is registered with Telegram.
              },
            })
            .then(
              ({ error }) => {
                if (error)
                  console.error(
                    `[telegram] send failed for ${table}:`,
                    error.message,
                  );
              },
              (err) =>
                console.error(`[telegram] send failed for ${table}:`, err),
            );
        };
        createdRows.forEach((row) => {
          const body = notify({ type: "create", row });
          sendPush(body);
          sendTelegram(body);
        });
        updatedRows.forEach(({ row, old }) => {
          const body = notify({ type: "update", row, old });
          sendPush(body);
          sendTelegram(body);
        });
      }
    },
    [table, mapToDb, audit, actorRef, labelOf, notify, t],
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
  const { t } = useLang();

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

  const setValue = useCallback(
    (next) => {
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
          if (error) {
            console.error(
              "[supabase] upsert failed on payroll_paid:",
              error.message,
            );
            pushToast(`${t.settings.saveFailed} ${error.message}`, "error");
          }
        }
        for (const key of removed) {
          const [employeeId, month] = splitPayrollKey(key);
          const { error } = await supabase
            .from("payroll_paid")
            .delete()
            .eq("employee_id", employeeId)
            .eq("month", month);
          if (error) {
            console.error(
              "[supabase] delete failed on payroll_paid:",
              error.message,
            );
            pushToast(`${t.settings.saveFailed} ${error.message}`, "error");
          }
        }
      })();
    },
    [t],
  );

  return [value, setValue, ready];
}

// branding is a single settings row (id = 1) shared by everyone — the
// company name + logo shown on the login screens and sidebar for every
// admin and employee, not just the device that set it.
function useBrandingSettings() {
  const [value, setValueState] = useState({ name: "", logo: null });
  const [ready, setReady] = useState(false);
  const { t } = useLang();

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

  const setValue = useCallback(
    (next) => {
      setValueState(next);
      (async () => {
        const { error } = await supabase.from("branding").upsert({
          id: 1,
          name: next.name || "",
          logo: next.logo || null,
        });
        if (error) {
          console.error("[supabase] save failed on branding:", error.message);
          pushToast(`${t.settings.saveFailed} ${error.message}`, "error");
        } else {
          pushToast(t.settings.saved, "success");
        }
      })();
    },
    [t],
  );

  return [value, setValue, ready];
}

// Login sessions are intentionally per-device, not shared data, so they
// stay in the browser's own localStorage instead of Supabase.
// ot_policy is a single settings row (id = 1), same shape as
// DEFAULT_OT_POLICY. Falls back to the defaults until an admin saves one.
function useOtPolicy() {
  const [value, setValueState] = useState(DEFAULT_OT_POLICY);
  const [ready, setReady] = useState(false);
  const { t } = useLang();

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

  const setValue = useCallback(
    (next) => {
      setValueState(next);
      (async () => {
        const { error } = await supabase.from("ot_policy").upsert({
          id: 1,
          rate_normal: next.rateNormal,
          rate_weekend: next.rateWeekend,
          rate_holiday: next.rateHoliday,
          hours_per_day: next.hoursPerDay,
        });
        if (error) {
          console.error("[supabase] save failed on ot_policy:", error.message);
          pushToast(`${t.settings.saveFailed} ${error.message}`, "error");
        } else {
          pushToast(t.settings.saved, "success");
        }
      })();
    },
    [t],
  );

  return [value, setValue, ready];
}

// payroll_policy is a single settings row (id = 1), same shape as
// DEFAULT_PAYROLL_POLICY. Falls back to the defaults until an admin saves one.
function usePayrollPolicy() {
  const [value, setValueState] = useState(DEFAULT_PAYROLL_POLICY);
  const [ready, setReady] = useState(false);
  const { t } = useLang();

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
          lateGraceCount:
            data.late_grace_count ?? DEFAULT_PAYROLL_POLICY.lateGraceCount,
          lateDeductionType:
            data.late_deduction_type ??
            DEFAULT_PAYROLL_POLICY.lateDeductionType,
          lateDeductionValue:
            data.late_deduction_value ??
            DEFAULT_PAYROLL_POLICY.lateDeductionValue,
          ulDeductionType:
            data.ul_deduction_type ?? DEFAULT_PAYROLL_POLICY.ulDeductionType,
          ulDeductionValue:
            data.ul_deduction_value ?? DEFAULT_PAYROLL_POLICY.ulDeductionValue,
          taxMode: data.tax_mode ?? DEFAULT_PAYROLL_POLICY.taxMode,
          exchangeRate:
            data.exchange_rate ?? DEFAULT_PAYROLL_POLICY.exchangeRate,
          nssfWageCapKHR:
            data.nssf_wage_cap_khr ?? DEFAULT_PAYROLL_POLICY.nssfWageCapKHR,
          nssfPensionEmployeeRate:
            data.nssf_pension_employee_rate ??
            DEFAULT_PAYROLL_POLICY.nssfPensionEmployeeRate,
          nssfPensionEmployerRate:
            data.nssf_pension_employer_rate ??
            DEFAULT_PAYROLL_POLICY.nssfPensionEmployerRate,
          nssfOrcRate: data.nssf_orc_rate ?? DEFAULT_PAYROLL_POLICY.nssfOrcRate,
          nssfHealthRate:
            data.nssf_health_rate ?? DEFAULT_PAYROLL_POLICY.nssfHealthRate,
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

  const setValue = useCallback(
    (next) => {
      setValueState(next);
      (async () => {
        const { error } = await supabase.from("payroll_policy").upsert({
          id: 1,
          tax_rate: next.taxRate,
          insurance_rate: next.insuranceRate,
          min_salary_threshold: next.minSalaryThreshold,
          late_grace_count: next.lateGraceCount,
          late_deduction_type: next.lateDeductionType,
          late_deduction_value: next.lateDeductionValue,
          ul_deduction_type: next.ulDeductionType,
          ul_deduction_value: next.ulDeductionValue,
          tax_mode: next.taxMode || "flat",
          exchange_rate: Number(next.exchangeRate) || 4100,
          nssf_wage_cap_khr: Number(next.nssfWageCapKHR) || 1200000,
          nssf_pension_employee_rate: Number(next.nssfPensionEmployeeRate) || 0,
          nssf_pension_employer_rate: Number(next.nssfPensionEmployerRate) || 0,
          nssf_orc_rate: Number(next.nssfOrcRate) || 0,
          nssf_health_rate: Number(next.nssfHealthRate) || 0,
        });
        if (error) {
          console.error(
            "[supabase] save failed on payroll_policy:",
            error.message,
          );
          pushToast(`${t.settings.saveFailed} ${error.message}`, "error");
        } else {
          pushToast(t.settings.saved, "success");
        }
      })();
    },
    [t],
  );

  return [value, setValue, ready];
}
// sound_policy is a single settings row (id = 1) holding which chime
// preset (see SOUND_PRESETS) plays on a QR check-in/out scan. Falls back
// to DEFAULT_SOUND_POLICY until a superadmin saves one.
function useSoundPolicy() {
  const [value, setValueState] = useState(DEFAULT_SOUND_POLICY);
  const [ready, setReady] = useState(false);
  const { t } = useLang();

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const { data, error } = await supabase
        .from("sound_policy")
        .select("*")
        .eq("id", 1)
        .maybeSingle();
      if (cancelled) return;
      if (error) {
        console.error("[supabase] failed to load sound_policy:", error.message);
        setValueState(DEFAULT_SOUND_POLICY);
      } else if (data) {
        setValueState({
          preset:
            data.preset && SOUND_PRESETS[data.preset]
              ? data.preset
              : DEFAULT_SOUND_POLICY.preset,
        });
      } else {
        setValueState(DEFAULT_SOUND_POLICY);
      }
      setReady(true);
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const setValue = useCallback(
    (next) => {
      setValueState(next);
      (async () => {
        const { error } = await supabase.from("sound_policy").upsert({
          id: 1,
          preset: next.preset,
        });
        if (error) {
          console.error(
            "[supabase] save failed on sound_policy:",
            error.message,
          );
          pushToast(`${t.settings.saveFailed} ${error.message}`, "error");
        } else {
          pushToast(t.settings.saved, "success");
        }
      })();
    },
    [t],
  );

  return [value, setValue, ready];
}
// telegram_settings is a single settings row (id = 1) holding the bot
// credentials and which request categories should be forwarded to a
// Telegram group chat. The bot_token/chat_id never need to leave the
// server: the client only ever calls the `telegram_notify` edge
// function with a plain text message + category, and that function
// looks up these settings itself (via the service role) before
// deciding whether/where to send. This hook exists purely so the
// Settings screen can show/edit the row.
const DEFAULT_TELEGRAM_SETTINGS = {
  enabled: false,
  botToken: "",
  chatId: "",
  notifyLeave: true,
  notifyOt: true,
  notifyAttcorr: true,
  notifyShiftswap: true,
  notifyPayroll: true,
  notifyLate: true,
  notifyChat: true,
};
function useTelegramSettings() {
  const [value, setValueState] = useState(DEFAULT_TELEGRAM_SETTINGS);
  const [ready, setReady] = useState(false);
  const { t } = useLang();

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const { data, error } = await supabase
        .from("telegram_settings")
        .select("*")
        .eq("id", 1)
        .maybeSingle();
      if (cancelled) return;
      if (error) {
        console.error(
          "[supabase] failed to load telegram_settings:",
          error.message,
        );
        setValueState(DEFAULT_TELEGRAM_SETTINGS);
      } else if (data) {
        setValueState({
          enabled: !!data.enabled,
          botToken: data.bot_token || "",
          chatId: data.chat_id || "",
          notifyLeave: data.notify_leave ?? true,
          notifyOt: data.notify_ot ?? true,
          notifyAttcorr: data.notify_attcorr ?? true,
          notifyShiftswap: data.notify_shiftswap ?? true,
          notifyPayroll: data.notify_payroll ?? true,
          notifyLate: data.notify_late ?? true,
          notifyChat: data.notify_chat ?? true,
        });
      } else {
        setValueState(DEFAULT_TELEGRAM_SETTINGS);
      }
      setReady(true);
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const setValue = useCallback(
    (next) => {
      setValueState(next);
      (async () => {
        const { error } = await supabase.from("telegram_settings").upsert({
          id: 1,
          enabled: !!next.enabled,
          bot_token: next.botToken || "",
          chat_id: next.chatId || "",
          notify_leave: !!next.notifyLeave,
          notify_ot: !!next.notifyOt,
          notify_attcorr: !!next.notifyAttcorr,
          notify_shiftswap: !!next.notifyShiftswap,
          notify_payroll: !!next.notifyPayroll,
          notify_late: !!next.notifyLate,
          notify_chat: !!next.notifyChat,
        });
        if (error) {
          console.error(
            "[supabase] save failed on telegram_settings:",
            error.message,
          );
          pushToast(`${t.settings.saveFailed} ${error.message}`, "error");
        } else {
          pushToast(t.settings.saved, "success");
        }
      })();
    },
    [t],
  );

  return [value, setValue, ready];
}
// Public VAPID key for Web Push (base64url, generated once for this
// deployment via `npx web-push generate-vapid-keys`). Push subscribe UI
// hides itself entirely when this is blank, so leaving it empty is safe.
// The matching PRIVATE key never goes in client code — it lives only in
// the server-side function that actually sends push messages (see
// push_subscriptions table + Edge Function, deployed separately).
const PUSH_VAPID_PUBLIC_KEY =
  "BEUW05cxGKpj15XXpvr-2EmtTAoXll3zWKBHj8S7Lj-18APot7YlEjNI4MsxbbJLLkDS4TPyWiBYE80DSToaFis";
function urlBase64ToUint8Array(base64String) {
  const padding = "=".repeat((4 - (base64String.length % 4)) % 4);
  const base64 = (base64String + padding).replace(/-/g, "+").replace(/_/g, "/");
  const rawData = window.atob(base64);
  const outputArray = new Uint8Array(rawData.length);
  for (let i = 0; i < rawData.length; i++)
    outputArray[i] = rawData.charCodeAt(i);
  return outputArray;
}
// Manages this device's Web Push subscription for either an admin or an
// employee (userType/userId identify who to notify). The subscription
// itself — endpoint + keys — is stored per-device in Supabase so a
// server-side function can look up who to send a push to; this hook only
// handles the browser side (permission, subscribe/unsubscribe, syncing
// that one row). Silently reports unsupported when PUSH_VAPID_PUBLIC_KEY
// isn't configured yet, or the browser has no Push support (e.g. iOS
// Safari unless the app is installed to the home screen).
function usePushSubscription(userType, userId) {
  const supported =
    typeof window !== "undefined" &&
    "serviceWorker" in navigator &&
    "PushManager" in window &&
    "Notification" in window &&
    !!PUSH_VAPID_PUBLIC_KEY;
  const [permission, setPermission] = useState(
    supported ? Notification.permission : "unsupported",
  );
  const [subscribed, setSubscribed] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    if (!supported || !userId) return;
    let cancelled = false;
    (async () => {
      try {
        const reg = await navigator.serviceWorker.ready;
        const sub = await reg.pushManager.getSubscription();
        if (!cancelled) setSubscribed(!!sub);
      } catch {
        // Service worker not ready yet (e.g. first paint) — leave as-is,
        // the user can still tap Enable to trigger registration.
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [supported, userId]);

  const subscribe = useCallback(async () => {
    if (!supported || !userId) return;
    setBusy(true);
    setError("");
    try {
      const perm = await Notification.requestPermission();
      setPermission(perm);
      if (perm !== "granted") {
        setBusy(false);
        return;
      }
      const reg = await navigator.serviceWorker.ready;
      let sub = await reg.pushManager.getSubscription();
      if (!sub) {
        sub = await reg.pushManager.subscribe({
          userVisibleOnly: true,
          applicationServerKey: urlBase64ToUint8Array(PUSH_VAPID_PUBLIC_KEY),
        });
      }
      const json = sub.toJSON();
      const { error: dbError } = await supabase
        .from("push_subscriptions")
        .upsert(
          {
            endpoint: json.endpoint,
            user_type: userType,
            user_id: userId,
            p256dh: json.keys?.p256dh,
            auth: json.keys?.auth,
          },
          { onConflict: "endpoint" },
        );
      if (dbError) throw dbError;
      setSubscribed(true);
    } catch (err) {
      console.error("[push] subscribe failed:", err.message);
      setError(err.message);
    } finally {
      setBusy(false);
    }
  }, [supported, userId, userType]);

  const unsubscribe = useCallback(async () => {
    if (!supported) return;
    setBusy(true);
    setError("");
    try {
      const reg = await navigator.serviceWorker.ready;
      const sub = await reg.pushManager.getSubscription();
      if (sub) {
        const endpoint = sub.endpoint;
        await sub.unsubscribe();
        const { error: dbError } = await supabase
          .from("push_subscriptions")
          .delete()
          .eq("endpoint", endpoint);
        if (dbError) throw dbError;
      }
      setSubscribed(false);
    } catch (err) {
      console.error("[push] unsubscribe failed:", err.message);
      setError(err.message);
    } finally {
      setBusy(false);
    }
  }, [supported]);

  return {
    supported,
    permission,
    subscribed,
    busy,
    error,
    subscribe,
    unsubscribe,
  };
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
    unpaid: {
      bg: T.roseSoft,
      fg: T.roseDark,
      label: en ? "Unpaid Leave" : "ច្បាប់គ្មានប្រាក់ខែ",
    },
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
    unpaid: en ? "Unpaid Leave (UL)" : "ច្បាប់គ្មានប្រាក់ខែ (UL)",
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
  shiftSwapRequests,
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
        shiftSwapRequests,
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
      shiftSwapRequests,
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
// A single icon button (no dropdown, unlike NotificationBell) that jumps
// straight to the Messages page, badged with the unread count for
// whoever is looking at it: an admin sees unread messages sent by any
// employee across every thread (admins share one mailbox, same as the
// rest of this app's admin-facing tables); an employee only ever sees
// unread messages in their own thread.
function ChatQuickAccess({ role, currentEmp, messages, setPage }) {
  const { t } = useLang();
  const unread = useMemo(() => {
    if (role === "admin") {
      return messages.filter(
        (m) => m.senderRole === "employee" && !m.readByAdmin,
      ).length;
    }
    if (!currentEmp) return 0;
    return messages.filter(
      (m) =>
        m.employeeId === currentEmp.id &&
        m.senderRole === "admin" &&
        !m.readByEmployee,
    ).length;
  }, [role, currentEmp, messages]);

  // Chimes whenever unread count rises — i.e. a new incoming message just
  // landed for whoever is looking at this (admin or employee). Skipped on
  // first mount so opening the app with pre-existing unread messages
  // doesn't fire the sound, and skipped on drops (marking read).
  const prevUnreadRef = useRef(unread);
  const mountedRef = useRef(false);
  useEffect(() => {
    if (mountedRef.current && unread > prevUnreadRef.current) {
      playChatNotifySound();
    }
    prevUnreadRef.current = unread;
    mountedRef.current = true;
  }, [unread]);

  return (
    <button
      onClick={() => setPage("messages")}
      aria-label={t.chat.title}
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
      <MessageCircle size={19} />
      {unread > 0 && (
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
          {unread > 9 ? "9+" : unread}
        </span>
      )}
    </button>
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
  const { t } = useLang();
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
          {value ? fmtDateDisplay(value) : placeholder || t.selectDate}
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
              {t.clear}
            </button>
            <button
              type="button"
              className="wf-dp-link"
              onClick={() => {
                fire(new Date());
                setOpen(false);
              }}
            >
              {t.today}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// Range variant of DatePicker: Start/End sit side by side but share ONE
// two-month calendar popover (like common "date range" pickers), with the
// days between start and end highlighted as a continuous band across both
// months — click a day to set the start of a fresh range, then click a
// later day to set the end (the popover closes automatically once the end
// is picked). Clicking an earlier day while only the start is set moves
// the start instead.
function buildCalendarCells(year, month) {
  const firstDow = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const daysInPrevMonth = new Date(year, month, 0).getDate();
  const totalCells = Math.ceil((firstDow + daysInMonth) / 7) * 7;
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
  return cells;
}
function monthTitleFor(year, month) {
  let title = `${year}-${month + 1}`;
  try {
    title = new Intl.DateTimeFormat("km-KH", {
      year: "numeric",
      month: "long",
    }).format(new Date(year, month, 1));
  } catch {}
  return title;
}
function DateRangePicker({
  startValue,
  endValue,
  onChangeStart,
  onChangeEnd,
  startLabel,
  endLabel,
  placeholder,
  style,
}) {
  const { t } = useLang();
  const [open, setOpen] = useState(false);
  const startSel = parseYMD(startValue);
  const endSel = parseYMD(endValue);
  const [cursor, setCursor] = useState(() => startSel || new Date());
  const wrapRef = useRef(null);
  useCloseOnOutside(wrapRef, () => setOpen(false));

  useEffect(() => {
    if (open) setCursor(startSel || endSel || new Date());
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const today = todayStr();
  const leftYear = cursor.getFullYear();
  const leftMonth = cursor.getMonth();
  const rightRef = new Date(leftYear, leftMonth + 1, 1);
  const rightYear = rightRef.getFullYear();
  const rightMonth = rightRef.getMonth();
  const leftCells = buildCalendarCells(leftYear, leftMonth);
  const rightCells = buildCalendarCells(rightYear, rightMonth);
  const leftTitle = monthTitleFor(leftYear, leftMonth);
  const rightTitle = monthTitleFor(rightYear, rightMonth);

  const pick = (d) => {
    const ymd = fmtYMD(d);
    if (!startValue || (startValue && endValue)) {
      // Nothing picked yet, or a full range already exists — start fresh.
      onChangeStart && onChangeStart({ target: { value: ymd } });
      onChangeEnd && onChangeEnd({ target: { value: "" } });
    } else if (ymd < startValue) {
      // Picked something earlier than the current start — move the start.
      onChangeStart && onChangeStart({ target: { value: ymd } });
    } else {
      // Completes the range.
      onChangeEnd && onChangeEnd({ target: { value: ymd } });
      setOpen(false);
    }
  };

  const clear = () => {
    onChangeStart && onChangeStart({ target: { value: "" } });
    onChangeEnd && onChangeEnd({ target: { value: "" } });
  };

  const inRange = (ymd) =>
    startValue && endValue && ymd > startValue && ymd < endValue;

  const renderPanel = (cells, key) => (
    <div className="wf-dp-panel" key={key}>
      <div className="wf-dp-grid wf-dp-grid-range">
        {DP_DOW.map((w, i) => (
          <div className="wf-dp-dow" key={i}>
            {w}
          </div>
        ))}
        {cells.map((c, i) => {
          const ymd = fmtYMD(c.date);
          const isStart = startValue && ymd === startValue;
          const isEnd = endValue && ymd === endValue;
          const isSingle = isStart && isEnd;
          const isToday = ymd === today;
          const within = inRange(ymd);
          return (
            <button
              type="button"
              key={i}
              className={`wf-dp-day${c.outside ? " outside" : ""}${
                isToday ? " today" : ""
              }${within ? " in-range" : ""}${
                isSingle
                  ? " range-single"
                  : isStart
                    ? " range-start"
                    : isEnd
                      ? " range-end"
                      : ""
              }`}
              onClick={() => pick(c.date)}
            >
              <span className="wf-dp-day-num">{c.day}</span>
            </button>
          );
        })}
      </div>
    </div>
  );

  return (
    <div className="wf-dp-wrap" ref={wrapRef} style={style}>
      <div className="wf-dr-triggers">
        <div>
          <span className="wf-field-label">{startLabel}</span>
          <button
            type="button"
            className={`wf-dp-trigger${open ? " open" : ""}`}
            onClick={() => setOpen((o) => !o)}
          >
            <span className={startValue ? "" : "wf-dp-placeholder"}>
              {startValue
                ? fmtDateDisplay(startValue)
                : placeholder || t.selectDate}
            </span>
            <CalendarDays size={15} style={{ opacity: 0.55, flexShrink: 0 }} />
          </button>
        </div>
        <div>
          <span className="wf-field-label">{endLabel}</span>
          <button
            type="button"
            className={`wf-dp-trigger${open ? " open" : ""}`}
            onClick={() => setOpen((o) => !o)}
          >
            <span className={endValue ? "" : "wf-dp-placeholder"}>
              {endValue
                ? fmtDateDisplay(endValue)
                : placeholder || t.selectDate}
            </span>
            <CalendarDays size={15} style={{ opacity: 0.55, flexShrink: 0 }} />
          </button>
        </div>
      </div>
      {open && (
        <div
          className="wf-dp-pop wf-dp-pop-wide"
          onMouseDown={(e) => e.stopPropagation()}
        >
          <div className="wf-dp-head-dual">
            <button
              type="button"
              className="wf-dp-nav"
              onClick={() => setCursor(new Date(leftYear, leftMonth - 1, 1))}
            >
              <ChevronLeft size={15} />
            </button>
            <div className="wf-dp-dual-titles">
              <span className="wf-dp-title">{leftTitle}</span>
              <span className="wf-dp-title">{rightTitle}</span>
            </div>
            <button
              type="button"
              className="wf-dp-nav"
              onClick={() => setCursor(new Date(leftYear, leftMonth + 1, 1))}
            >
              <ChevronRight size={15} />
            </button>
          </div>
          <div className="wf-dp-dual">
            {renderPanel(leftCells, "left")}
            {renderPanel(rightCells, "right")}
          </div>
          <div className="wf-dp-foot">
            <button type="button" className="wf-dp-link" onClick={clear}>
              {t.clear}
            </button>
            <button
              type="button"
              className="wf-dp-link"
              onClick={() => pick(new Date())}
            >
              {t.today}
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
  const { t } = useLang();
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
          {value || placeholder || t.timeLabel}
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
              {t.clear}
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
              {t.now}
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
const APP_VERSION = "1.0.0";
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
  display:flex; flex-direction:column; align-items:center; justify-content:center;
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
.wf-login-credit {
  margin-top:32px; text-align:center; font-size:10.5px;
  letter-spacing:0.3px; color:#5B6478; opacity:0.6;
}
`;
function LoginCredit() {
  return (
    <div className="wf-login-credit">
      v{APP_VERSION} &middot; Developed by Ou SoThon
    </div>
  );
}
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
      <LoginCredit />
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
      <LoginCredit />
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
// Public, unauthenticated URL for one office's kiosk display (see
// KioskDisplay below) — meant to be left open on a tablet/monitor
// mounted at that branch's entrance.
function officeKioskUrl(officeId) {
  const { origin, pathname } = window.location;
  return `${origin}${pathname}#/kiosk/${officeId}`;
}
// Generic QR display modal. Pass `url` for the existing employee-portal
// link use case (kept for backwards compatibility), or `data` for any
// other raw payload to encode (e.g. an office check-in QR token). `title`
// and `desc` let callers override the default employee-portal copy.
function QrModal({ url, data, title, desc, onClose, footer }) {
  const { t: t2 } = useLang();
  const payload = data != null ? data : url;
  const qrSrc = `https://api.qrserver.com/v1/create-qr-code/?size=240x240&data=${encodeURIComponent(payload)}`;
  return (
    <Modal title={title || t2.qr.title} onClose={onClose} width={340}>
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
          <img src={qrSrc} alt="QR code" width={220} height={220} />
        </div>
        <p style={{ fontSize: 12, color: T.muted, textAlign: "center" }}>
          {desc || t2.qr.desc}
        </p>
        {footer}
      </div>
    </Modal>
  );
}
function EmployeeLinkCard({ variant = "card" }) {
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

  // "card" (default): a standalone Card, used wherever this needs to sit
  // on its own (e.g. previously on the Dashboard). "inline": just the
  // row content with no outer Card/margin, for embedding inside another
  // Card (e.g. the Settings page) that already provides its own padding.
  const Wrapper = variant === "inline" ? "div" : Card;
  const wrapperStyle =
    variant === "inline"
      ? {
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: 16,
          flexWrap: "wrap",
        }
      : {
          padding: "16px 18px",
          marginBottom: 22,
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: 16,
          flexWrap: "wrap",
        };

  return (
    <Wrapper style={wrapperStyle}>
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
    </Wrapper>
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
  offices,
}) {
  const { t, lang } = useLang();
  const today = todayStr();
  // "Working now" = checked in today, not checked out yet, and not on
  // leave/absent. Driven straight off the `attendance` array, which is
  // kept live via the shared Supabase realtime channel (see
  // REALTIME_TABLES), so this list updates on its own as people punch
  // in/out — no polling needed here.
  const workingNowByBranch = useMemo(() => {
    if (role !== "admin") return [];
    const active = attendance.filter(
      (a) =>
        a.date === today &&
        a.checkIn &&
        !a.checkOut &&
        (a.status === "present" || a.status === "late"),
    );
    const groups = new Map();
    active.forEach((a) => {
      const emp = employees.find((e) => e.id === a.employeeId);
      if (!emp) return;
      const officeId = a.officeId || emp.officeId || "";
      const office = (offices || []).find((o) => o.id === officeId);
      const key = office?.id || "unassigned";
      if (!groups.has(key)) {
        groups.set(key, { office, entries: [] });
      }
      groups.get(key).entries.push({ emp, checkIn: a.checkIn });
    });
    return Array.from(groups.values()).sort(
      (a, b) => b.entries.length - a.entries.length,
    );
  }, [role, attendance, employees, offices, today]);
  const workingNowTotal = workingNowByBranch.reduce(
    (sum, g) => sum + g.entries.length,
    0,
  );
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

      {role === "admin" && offices && offices.length > 0 && (
        <Card style={{ padding: 18, marginBottom: 22 }}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              marginBottom: 12,
            }}
          >
            <h3
              style={{
                fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
                fontWeight: 600,
                color: T.ink,
                fontSize: 14,
                display: "flex",
                alignItems: "center",
                gap: 8,
              }}
            >
              <span
                style={{
                  width: 8,
                  height: 8,
                  borderRadius: 999,
                  background: T.forest,
                  display: "inline-block",
                  boxShadow: `0 0 0 4px ${T.forestSoft}`,
                }}
              />
              {t.dash.workingNow}
            </h3>
            <span
              style={{
                fontSize: 11,
                color: T.muted,
                fontFamily: "'JetBrains Mono',monospace",
              }}
            >
              {t.dash.workingNowSub(workingNowTotal)}
            </span>
          </div>
          {workingNowByBranch.length === 0 ? (
            <p
              style={{
                fontSize: 13,
                color: T.muted,
                textAlign: "center",
                padding: "28px 0",
              }}
            >
              {t.dash.noOneWorkingNow}
            </p>
          ) : (
            <div
              className="wf-grid"
              style={{
                gridTemplateColumns: "repeat(auto-fit, minmax(220px,1fr))",
                gap: 12,
              }}
            >
              {workingNowByBranch.map((g) => (
                <div
                  key={g.office?.id || "unassigned"}
                  style={{
                    border: `1px solid ${T.divider}`,
                    borderRadius: 12,
                    padding: 12,
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 6,
                      fontSize: 12,
                      fontWeight: 600,
                      color: T.ink,
                      marginBottom: 8,
                    }}
                  >
                    <Store size={13} color={T.blue} />
                    {g.office?.name || t.dash.unassignedBranch}
                    <span
                      style={{
                        marginLeft: "auto",
                        fontFamily: "'JetBrains Mono',monospace",
                        color: T.muted,
                        fontWeight: 500,
                      }}
                    >
                      {g.entries.length}
                    </span>
                  </div>
                  <div
                    style={{
                      display: "flex",
                      flexDirection: "column",
                      gap: 6,
                    }}
                  >
                    {g.entries.map(({ emp, checkIn }) => (
                      <div
                        key={emp.id}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 8,
                        }}
                      >
                        <Avatar name={emp.name} photo={emp.photo} size={24} />
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div
                            style={{
                              fontSize: 12,
                              color: T.ink,
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                              whiteSpace: "nowrap",
                            }}
                          >
                            {emp.name}
                          </div>
                        </div>
                        <div
                          style={{
                            fontSize: 10,
                            color: T.muted,
                            fontFamily: "'JetBrains Mono',monospace",
                            whiteSpace: "nowrap",
                          }}
                        >
                          {t.dash.sinceLabel} {checkIn}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          )}
        </Card>
      )}

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
                      {t.att.checkIn} {a.checkIn || "—"}{" "}
                      {a.checkOut ? `· ${t.att.checkOut} ${a.checkOut}` : ""}
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

  // New hires per month, based on join date, across all employees
  // (active or not) — a hire event is a historical fact regardless of
  // the employee's current status.
  const newHiresTrend = months.map((mkm) => {
    const count = employees.filter(
      (e) => e.joined && e.joined.startsWith(mkm),
    ).length;
    return { label: shortMonthLabel(mkm, lang), value: count };
  });
  const hasHiresData = employees.some((e) => e.joined);

  // Absenteeism rate by department for the current month: share of
  // that department's attendance records this month marked "absent".
  const absenteeismByDept = departments
    .map((d) => {
      const deptEmpIds = new Set(
        employees.filter((e) => e.deptId === d.id).map((e) => e.id),
      );
      const recs = attendance.filter(
        (a) => deptEmpIds.has(a.employeeId) && a.date && a.date.startsWith(mk),
      );
      const absent = recs.filter((a) => a.status === "absent").length;
      const rate = recs.length ? Math.round((absent / recs.length) * 100) : 0;
      return { label: d.name, value: rate, recCount: recs.length };
    })
    .filter((d) => d.recCount > 0)
    .sort((a, b) => b.value - a.value);
  const hasAbsenceData = absenteeismByDept.length > 0;

  // Snapshot workforce stats: active headcount, average tenure of
  // active employees, and the current inactive share. The inactive
  // share is a point-in-time snapshot, not a turnover rate over a
  // period — the app tracks a status flag, not an exit date.
  const avgTenureDays = activeEmployees.length
    ? activeEmployees.reduce((sum, e) => {
        if (!e.joined) return sum;
        const days = Math.max(
          0,
          (Date.now() - new Date(e.joined + "T00:00:00").getTime()) / 86400000,
        );
        return sum + days;
      }, 0) / activeEmployees.length
    : 0;
  const tenureYears = Math.floor(avgTenureDays / 365);
  const tenureMonths = Math.floor((avgTenureDays % 365) / 30);
  const inactiveCount = employees.filter((e) => e.status === "inactive").length;
  const inactiveRate = employees.length
    ? Math.round((inactiveCount / employees.length) * 100)
    : 0;

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
          gridTemplateColumns: "repeat(auto-fit, minmax(190px,1fr))",
          marginBottom: 16,
        }}
      >
        <Card accent={T.forest} style={{ padding: 16 }}>
          <div
            style={{
              width: 36,
              height: 36,
              borderRadius: 10,
              background: T.forest + "1A",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              marginBottom: 12,
            }}
          >
            <Users size={18} color={T.forest} />
          </div>
          <div
            style={{
              fontSize: 26,
              fontWeight: 700,
              fontFamily: "'JetBrains Mono',monospace",
              color: T.ink,
            }}
          >
            {activeEmployees.length}
          </div>
          <div
            style={{
              fontSize: 13,
              fontWeight: 600,
              color: T.textSoft,
              marginTop: 2,
            }}
          >
            {t.analytics.statActiveHeadcount}
          </div>
        </Card>

        <Card accent={T.blue} style={{ padding: 16 }}>
          <div
            style={{
              width: 36,
              height: 36,
              borderRadius: 10,
              background: T.blue + "1A",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              marginBottom: 12,
            }}
          >
            <CalendarClock size={18} color={T.blue} />
          </div>
          <div
            style={{
              fontSize: 26,
              fontWeight: 700,
              fontFamily: "'JetBrains Mono',monospace",
              color: T.ink,
            }}
          >
            {t.analytics.tenureFormat(tenureYears, tenureMonths)}
          </div>
          <div
            style={{
              fontSize: 13,
              fontWeight: 600,
              color: T.textSoft,
              marginTop: 2,
            }}
          >
            {t.analytics.statAvgTenure}
          </div>
        </Card>

        <Card accent={T.rose} style={{ padding: 16 }}>
          <div
            style={{
              width: 36,
              height: 36,
              borderRadius: 10,
              background: T.rose + "1A",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              marginBottom: 12,
            }}
          >
            <AlertCircle size={18} color={T.rose} />
          </div>
          <div
            style={{
              fontSize: 26,
              fontWeight: 700,
              fontFamily: "'JetBrains Mono',monospace",
              color: T.ink,
            }}
          >
            {inactiveRate}%
          </div>
          <div
            style={{
              fontSize: 13,
              fontWeight: 600,
              color: T.textSoft,
              marginTop: 2,
            }}
          >
            {t.analytics.statInactiveRate}
          </div>
          <div
            style={{
              fontSize: 10.5,
              color: T.muted,
              marginTop: 4,
              lineHeight: 1.5,
            }}
          >
            {t.analytics.statInactiveRateNote}
          </div>
        </Card>
      </div>

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

        <ChartCard
          title={t.analytics.newHiresTrend}
          subtitle={t.analytics.newHiresTrendSub}
          noData={!hasHiresData}
          noDataLabel={t.analytics.noChartData}
        >
          <MiniBarChart
            data={newHiresTrend}
            color={T.blue}
            formatValue={(v) => `${v}`}
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

      <div style={{ marginTop: 16 }}>
        <ChartCard
          title={t.analytics.absenteeism}
          subtitle={t.analytics.absenteeismSub}
          noData={!hasAbsenceData}
          noDataLabel={t.analytics.noChartData}
        >
          <HorizontalBarChart
            data={absenteeismByDept}
            color={T.rose}
            formatValue={(v) => `${v}%`}
          />
        </ChartCard>
      </div>
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
      dependents: 0,
      status: "active",
      joined: todayStr(),
      useCustomRate: false,
      customTaxRate: "",
      customInsuranceRate: "",
      useCustomLatePolicy: false,
      customLateGraceCount: "",
      customLateDeductionType: "fixed",
      customLateDeductionValue: "",
      useCustomUlPolicy: false,
      customUlDeductionType: "fullDay",
      customUlDeductionValue: "",
      messagesDisabled: false,
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
          placeholder={t.emps.namePlaceholder}
        />
      </Field>
      <div className="wf-grid-2">
        <Field label={t.emps.code}>
          <Input
            value={f.code}
            onChange={set("code")}
            placeholder={t.emps.codePlaceholder}
          />
        </Field>
        <Field label={t.emps.role}>
          <Input
            value={f.role}
            onChange={set("role")}
            placeholder={t.emps.rolePlaceholder}
          />
        </Field>
      </div>
      <Field label={t.emps.pinLabel}>
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
            <KeyRound size={13} /> {t.emps.pinRegenerate}
          </Button>
        </div>
      </Field>
      <div className="wf-grid-2">
        <Field label={t.emps.dept}>
          <Select value={f.deptId} onChange={set("deptId")}>
            {departments.map((d) => (
              <option key={d.id} value={d.id}>
                {d.name}
              </option>
            ))}
          </Select>
        </Field>
        <Field label={t.emps.shift}>
          <Select value={f.shiftId} onChange={set("shiftId")}>
            {shifts.map((s) => (
              <option key={s.id} value={s.id}>
                {shiftLabel(s)}
              </option>
            ))}
          </Select>
        </Field>
      </div>
      <div className="wf-grid-2">
        <Field label={t.emps.email}>
          <Input
            value={f.email}
            onChange={set("email")}
            placeholder="name@company.com"
          />
        </Field>
        <Field label={t.emps.phone}>
          <Input
            value={f.phone}
            onChange={set("phone")}
            placeholder="012 345 678"
          />
        </Field>
      </div>
      <div className="wf-grid-2">
        <Field label={`${t.emps.salary} (USD)`}>
          <Input
            type="number"
            value={f.salary}
            onChange={set("salary")}
            placeholder="600"
          />
        </Field>
        <Field label={t.emps.status}>
          <Select value={f.status} onChange={set("status")}>
            <option value="active">{t.emps.active}</option>
            <option value="inactive">{t.emps.inactive}</option>
          </Select>
        </Field>
      </div>
      <div style={{ marginBottom: 14 }}>
        <Field label={t.emps.dependentsLabel}>
          <Input
            type="number"
            min="0"
            step="1"
            value={f.dependents}
            onChange={set("dependents")}
            placeholder="0"
          />
        </Field>
        <p style={{ fontSize: 11.5, color: T.muted, marginTop: 4 }}>
          {t.emps.dependentsHint}
        </p>
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
          <div className="wf-grid-2">
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
            checked={!!f.useCustomLatePolicy}
            onChange={(e) =>
              setF({ ...f, useCustomLatePolicy: e.target.checked })
            }
            style={{ width: 15, height: 15, accentColor: T.forest }}
          />
          {t.pay.customLatePolicyToggle}
        </label>
        <p
          style={{
            fontSize: 11.5,
            color: T.muted,
            marginTop: 6,
            marginBottom: f.useCustomLatePolicy ? 12 : 0,
          }}
        >
          {t.pay.customLatePolicyHint}
        </p>
        {f.useCustomLatePolicy && (
          <>
            <div className="wf-grid-2">
              <Field label={t.pay.lateGraceCountLabel}>
                <Input
                  type="number"
                  step="1"
                  min="0"
                  value={f.customLateGraceCount}
                  onChange={set("customLateGraceCount")}
                  placeholder="0"
                />
              </Field>
              <Field label={t.pay.lateDeductionTypeLabel}>
                <Select
                  value={f.customLateDeductionType || "fixed"}
                  onChange={set("customLateDeductionType")}
                >
                  <option value="fixed">{t.pay.lateDeductionTypeFixed}</option>
                  <option value="percentDaily">
                    {t.pay.lateDeductionTypePercent}
                  </option>
                </Select>
              </Field>
            </div>
            <div style={{ marginTop: 12 }}>
              <Field
                label={
                  f.customLateDeductionType === "percentDaily"
                    ? t.pay.lateDeductionValuePercentLabel
                    : t.pay.lateDeductionValueFixedLabel
                }
              >
                <Input
                  type="number"
                  step="0.1"
                  min="0"
                  value={f.customLateDeductionValue}
                  onChange={set("customLateDeductionValue")}
                  placeholder="0"
                />
              </Field>
            </div>
          </>
        )}
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
            checked={!!f.useCustomUlPolicy}
            onChange={(e) =>
              setF({ ...f, useCustomUlPolicy: e.target.checked })
            }
            style={{ width: 15, height: 15, accentColor: T.forest }}
          />
          {t.pay.customUlPolicyToggle}
        </label>
        <p
          style={{
            fontSize: 11.5,
            color: T.muted,
            marginTop: 6,
            marginBottom: f.useCustomUlPolicy ? 12 : 0,
          }}
        >
          {t.pay.customUlPolicyHint}
        </p>
        {f.useCustomUlPolicy && (
          <>
            <Field label={t.pay.ulDeductionTypeLabel}>
              <Select
                value={f.customUlDeductionType || "fullDay"}
                onChange={set("customUlDeductionType")}
              >
                <option value="fullDay">{t.pay.ulDeductionTypeFullDay}</option>
                <option value="fixed">{t.pay.ulDeductionTypeFixed}</option>
                <option value="percentDaily">
                  {t.pay.ulDeductionTypePercent}
                </option>
              </Select>
            </Field>
            {f.customUlDeductionType !== "fullDay" && (
              <div style={{ marginTop: 12 }}>
                <Field
                  label={
                    f.customUlDeductionType === "percentDaily"
                      ? t.pay.ulDeductionValuePercentLabel
                      : t.pay.ulDeductionValueFixedLabel
                  }
                >
                  <Input
                    type="number"
                    step="0.1"
                    min="0"
                    value={f.customUlDeductionValue}
                    onChange={set("customUlDeductionValue")}
                    placeholder="0"
                  />
                </Field>
              </div>
            )}
          </>
        )}
      </div>
      <Field label={t.emps.joined}>
        <DatePicker value={f.joined || todayStr()} onChange={set("joined")} />
      </Field>
      <Field label={t.emps.branch}>
        <Select value={f.officeId || ""} onChange={set("officeId")}>
          <option value="">{t.emps.branchNotSet}</option>
          {(offices || []).map((o) => (
            <option key={o.id} value={o.id}>
              {o.name}
            </option>
          ))}
        </Select>
      </Field>
      <Field label={t.emps.annualLeaveDaysLabel}>
        <Input
          type="number"
          min={0}
          value={f.annualLeaveDays ?? DEFAULT_ANNUAL_LEAVE_DAYS}
          onChange={set("annualLeaveDays")}
          placeholder={String(DEFAULT_ANNUAL_LEAVE_DAYS)}
        />
      </Field>
      <Field label={t.emps.sickLeaveDaysLabel}>
        <Input
          type="number"
          min={0}
          value={f.sickLeaveDays ?? DEFAULT_SICK_LEAVE_DAYS}
          onChange={set("sickLeaveDays")}
          placeholder={String(DEFAULT_SICK_LEAVE_DAYS)}
        />
      </Field>
      <Field label={t.emps.weeklyOffLabel}>
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
      <Field label={t.emps.customDaysOffLabel}>
        <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
          <DatePicker
            style={{ flex: 1 }}
            value={newOffDate}
            onChange={(e) => setNewOffDate(e.target.value)}
          />
          <Button type="button" variant="ghost" onClick={addCustomDayOff}>
            <Plus size={14} /> {t.emps.addBtnShort}
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
            checked={!!f.messagesDisabled}
            onChange={(e) => setF({ ...f, messagesDisabled: e.target.checked })}
            style={{ width: 15, height: 15, accentColor: T.forest }}
          />
          {lang === "en"
            ? "Disable Messages & Voice Call for this employee"
            : "បិទ សារ & ការហៅសំឡេង សម្រាប់បុគ្គលិកនេះ"}
        </label>
        <p style={{ fontSize: 11.5, color: T.muted, marginTop: 6 }}>
          {lang === "en"
            ? "Overrides the company-wide setting for this one person only. Has no effect if Messages & Voice Call is already off for everyone in Role Permissions."
            : "អនុវត្តតែចំពោះបុគ្គលិកនេះម្នាក់ប៉ុណ្ណោះ ដោយបដិសេធការកំណត់ទូទៅរបស់ក្រុមហ៊ុន។ ប្រសិនបើមុខងារនេះបានបិទសម្រាប់អ្នកគ្រប់គ្នារួចហើយនៅ សិទ្ធិតួនាទី វានឹងគ្មានប្រសិទ្ធភាពទេ។"}
        </p>
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

/* ---------------------------------------------------------------
   Bulk employee import (Excel/CSV) — parses a spreadsheet the admin
   uploads, validates each row against existing departments/shifts/
   offices and existing employee codes, and lets the admin confirm
   before writing anything. Reuses ExcelJS (already a dependency for
   payroll export) for .xlsx, and a small hand-rolled parser for
   .csv so no extra package is required.
----------------------------------------------------------------*/
const IMPORT_TEMPLATE_HEADERS = [
  "Name",
  "Employee ID",
  "Department",
  "Role",
  "Shift",
  "Phone",
  "Email",
  "Salary",
  "Join Date (YYYY-MM-DD)",
  "Branch",
  "Status (active/inactive)",
];

// Ships the template as a real .xlsx (rather than plain CSV) so we can
// set an explicit font on the sample row. Excel's default Calibri has
// no Khmer glyphs, so a department/shift name in Khmer script renders
// as empty boxes/underscores in a plain CSV opened in Excel — the
// underlying text is still correct and still imports fine, but it's
// unreadable while editing. A Khmer-capable font fixes the display;
// it doesn't change what data actually gets written or parsed.
async function downloadImportTemplate(departments, shifts) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Workforce Suite";
  wb.created = new Date();
  const ws = wb.addWorksheet("Employees");

  const headerRow = ws.addRow(IMPORT_TEMPLATE_HEADERS);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF12203D" },
    };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });

  const sampleRow = ws.addRow([
    "John Doe",
    "EMP-101",
    departments[0]?.name || "HR",
    "Staff",
    shifts[0] ? shiftLabel(shifts[0]) : "Morning Shift",
    "012 345 678",
    "john@company.com",
    600,
    todayStr(),
    "",
    "active",
  ]);
  // "Khmer OS" is the most commonly pre-installed Khmer font on
  // Cambodian business machines; falls back harmlessly (Excel just
  // keeps whatever it was already using) if it isn't present.
  sampleRow.font = { name: "Khmer OS Battambang", size: 11 };

  ws.columns = IMPORT_TEMPLATE_HEADERS.map((h) => ({
    width: Math.max(16, h.length + 4),
  }));

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "employee-import-template.xlsx";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Minimal RFC-4180-ish CSV parser: handles quoted fields, escaped
// quotes ("") inside quotes, and both \n and \r\n line endings.
function parseCsvText(text) {
  const clean = text.replace(/^\uFEFF/, "");
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  for (let i = 0; i < clean.length; i++) {
    const c = clean[i];
    if (inQuotes) {
      if (c === '"') {
        if (clean[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ",") {
      row.push(field);
      field = "";
    } else if (c === "\n" || c === "\r") {
      if (c === "\r" && clean[i + 1] === "\n") i++;
      row.push(field);
      if (!(row.length === 1 && row[0] === "")) rows.push(row);
      row = [];
      field = "";
    } else {
      field += c;
    }
  }
  if (field !== "" || row.length) {
    row.push(field);
    if (!(row.length === 1 && row[0] === "")) rows.push(row);
  }
  return rows;
}

async function parseImportFile(file) {
  if (/\.csv$/i.test(file.name)) {
    const text = await file.text();
    return parseCsvText(text);
  }
  const buf = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const ws = wb.worksheets[0];
  const rows = [];
  ws?.eachRow((wsRow) => {
    const values = wsRow.values.slice(1); // ExcelJS rows are 1-indexed
    const cells = values.map((v) => {
      if (v === null || v === undefined) return "";
      if (typeof v === "object" && v.text !== undefined) return v.text;
      if (typeof v === "object" && v.result !== undefined) return v.result;
      return String(v);
    });
    rows.push(cells);
  });
  return rows;
}

// Turns raw parsed rows (first row = headers, ignored — column order
// is assumed to match IMPORT_TEMPLATE_HEADERS) into validated
// employee objects, matching Department/Shift/Branch by name
// (case-insensitive) and flagging anything that can't be resolved
// or that collides with an existing/earlier-in-file employee code.
function validateImportRows(
  rows,
  { departments, shifts, offices, employees, t },
) {
  const dataRows = rows
    .slice(1)
    .filter((r) => r.some((c) => String(c ?? "").trim() !== ""));
  const existingCodes = new Set(
    employees.map((e) =>
      String(e.code || "")
        .trim()
        .toLowerCase(),
    ),
  );
  const seenCodesInFile = new Set();
  const findByName = (list, name) =>
    list.find(
      (x) =>
        x.name.trim().toLowerCase() ===
        String(name || "")
          .trim()
          .toLowerCase(),
    );

  return dataRows.map((r, idx) => {
    const [
      name,
      code,
      deptName,
      role,
      shiftName,
      phone,
      email,
      salaryRaw,
      joined,
      officeName,
      statusRaw,
    ] = r.map((c) => (c ?? "").toString().trim());

    const errors = [];
    if (!name) errors.push(t.emps.errMissingName);
    if (!code) errors.push(t.emps.errMissingCode);
    const codeKey = code.toLowerCase();
    if (code && (existingCodes.has(codeKey) || seenCodesInFile.has(codeKey))) {
      errors.push(t.emps.errDuplicateCode(code));
    }

    const dept = deptName ? findByName(departments, deptName) : departments[0];
    if (deptName && !dept) errors.push(t.emps.errDeptNotFound(deptName));

    const shift = shiftName ? findByName(shifts, shiftName) : shifts[0];
    if (shiftName && !shift) errors.push(t.emps.errShiftNotFound(shiftName));

    let office = null;
    if (officeName) {
      office = findByName(offices || [], officeName);
      if (!office) errors.push(t.emps.errOfficeNotFound(officeName));
    }

    const salary = Number(salaryRaw);
    if (salaryRaw && (!Number.isFinite(salary) || salary < 0)) {
      errors.push(t.emps.errInvalidSalary);
    }

    const statusNorm = statusRaw.toLowerCase();
    const status = statusNorm === "inactive" ? "inactive" : "active";

    if (code && errors.length === 0) seenCodesInFile.add(codeKey);

    const employee =
      errors.length === 0
        ? {
            id: uid("e"),
            code,
            pin: randomPin(),
            name,
            deptId: dept?.id || "",
            shiftId: shift?.id || "",
            officeId: office?.id || "",
            weeklyOff: [],
            customDaysOff: [],
            annualLeaveDays: DEFAULT_ANNUAL_LEAVE_DAYS,
            sickLeaveDays: DEFAULT_SICK_LEAVE_DAYS,
            role,
            email,
            phone,
            salary: salaryRaw ? salary : 0,
            dependents: 0,
            status,
            joined: joined || todayStr(),
            useCustomRate: false,
            customTaxRate: null,
            customInsuranceRate: null,
            useCustomLatePolicy: false,
            customLateGraceCount: null,
            customLateDeductionType: null,
            customLateDeductionValue: null,
            useCustomUlPolicy: false,
            customUlDeductionType: null,
            customUlDeductionValue: null,
          }
        : null;

    return { rowNumber: idx + 2, raw: { name, code }, errors, employee };
  });
}

function ImportEmployeesModal({
  departments,
  shifts,
  offices,
  employees,
  onImport,
  onClose,
}) {
  const { t } = useLang();
  const fileRef = useRef(null);
  const [fileName, setFileName] = useState("");
  const [parsing, setParsing] = useState(false);
  const [results, setResults] = useState(null); // array from validateImportRows
  const [parseError, setParseError] = useState("");

  const handleFile = async (file) => {
    if (!file) return;
    setFileName(file.name);
    setParsing(true);
    setParseError("");
    setResults(null);
    try {
      const rows = await parseImportFile(file);
      if (!rows.length) {
        setParseError(t.emps.emptyFile);
      } else {
        setResults(
          validateImportRows(rows, {
            departments,
            shifts,
            offices,
            employees,
            t,
          }),
        );
      }
    } catch (err) {
      setParseError(err?.message || String(err));
    } finally {
      setParsing(false);
    }
  };

  const validRows = (results || []).filter((r) => r.errors.length === 0);
  const errorRows = (results || []).filter((r) => r.errors.length > 0);

  return (
    <Modal title={t.emps.importTitle} onClose={onClose} width={680}>
      <p
        style={{
          fontSize: 12.5,
          color: T.textSoft,
          marginBottom: 14,
          lineHeight: 1.6,
        }}
      >
        {t.emps.importDesc}
      </p>
      <p
        style={{
          fontSize: 11.5,
          color: T.muted,
          marginBottom: 14,
          lineHeight: 1.6,
          background: T.tableHeadBg,
          borderRadius: 8,
          padding: "8px 10px",
        }}
      >
        ℹ️ {t.emps.importFontNote}
      </p>
      <div
        style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}
      >
        <Button
          variant="ghost"
          type="button"
          onClick={() =>
            downloadImportTemplate(departments, shifts).catch((err) =>
              pushToast(err?.message || String(err), "error"),
            )
          }
        >
          <FileSpreadsheet size={15} /> {t.emps.downloadTemplate}
        </Button>
        <Button
          variant="accent"
          type="button"
          onClick={() => fileRef.current?.click()}
        >
          <Upload size={15} />{" "}
          {fileName ? t.emps.changeFile : t.emps.chooseFile}
        </Button>
        <input
          ref={fileRef}
          type="file"
          accept=".xlsx,.xls,.csv"
          style={{ display: "none" }}
          onChange={(e) => handleFile(e.target.files?.[0])}
        />
      </div>

      {fileName && (
        <div style={{ fontSize: 12, color: T.muted, marginBottom: 10 }}>
          📄 {fileName}
        </div>
      )}
      {parsing && (
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 8,
            color: T.muted,
            fontSize: 13,
          }}
        >
          <Loader2 size={15} style={{ animation: "spin 1s linear infinite" }} />{" "}
          {t.emps.parsing}
        </div>
      )}
      {parseError && (
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 8,
            color: T.rose,
            fontSize: 13,
            marginBottom: 10,
          }}
        >
          <AlertCircle size={15} /> {parseError}
        </div>
      )}

      {results && (
        <div>
          <div
            style={{
              fontSize: 12.5,
              fontWeight: 600,
              color: errorRows.length ? T.rose : T.forestText,
              marginBottom: 10,
            }}
          >
            {t.emps.summaryText(
              results.length,
              validRows.length,
              errorRows.length,
            )}
          </div>
          <div
            style={{
              maxHeight: 260,
              overflowY: "auto",
              border: `1px solid ${T.lineSoft}`,
              borderRadius: 10,
              marginBottom: 16,
            }}
          >
            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
                fontSize: 12,
              }}
            >
              <thead>
                <tr style={{ background: T.tableHeadBg }}>
                  <th style={{ textAlign: "left", padding: "8px 10px" }}>
                    {t.emps.colRow}
                  </th>
                  <th style={{ textAlign: "left", padding: "8px 10px" }}>
                    {t.emps.name}
                  </th>
                  <th style={{ textAlign: "left", padding: "8px 10px" }}>
                    {t.emps.code}
                  </th>
                  <th style={{ textAlign: "left", padding: "8px 10px" }}>
                    {t.emps.colStatus}
                  </th>
                </tr>
              </thead>
              <tbody>
                {results.map((r) => (
                  <tr
                    key={r.rowNumber}
                    style={{ borderTop: `1px solid ${T.lineSoft}` }}
                  >
                    <td style={{ padding: "6px 10px", color: T.muted }}>
                      {r.rowNumber}
                    </td>
                    <td style={{ padding: "6px 10px" }}>{r.raw.name || "—"}</td>
                    <td
                      style={{
                        padding: "6px 10px",
                        fontFamily: "'JetBrains Mono',monospace",
                      }}
                    >
                      {r.raw.code || "—"}
                    </td>
                    <td style={{ padding: "6px 10px" }}>
                      {r.errors.length === 0 ? (
                        <span
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: 5,
                            color: T.forestText,
                          }}
                        >
                          <CheckCircle2 size={13} /> {t.emps.rowValid}
                        </span>
                      ) : (
                        <span
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: 5,
                            color: T.rose,
                          }}
                        >
                          <AlertCircle size={13} /> {r.errors.join(" · ")}
                        </span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
        <Button variant="ghost" type="button" onClick={onClose}>
          {t.cancel}
        </Button>
        {results && (
          <Button
            variant="accent"
            type="button"
            disabled={validRows.length === 0}
            onClick={() => onImport(validRows.map((r) => r.employee))}
          >
            <Upload size={15} />{" "}
            {validRows.length === 0
              ? t.emps.noValidRows
              : t.emps.confirmImportBtn(validRows.length)}
          </Button>
        )}
      </div>
    </Modal>
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
  const [certFor, setCertFor] = useState(null);
  const [query, setQuery] = useState("");
  const [branchFilter, setBranchFilter] = useState("");
  const [badgePopupBlocked, setBadgePopupBlocked] = useState(false);
  const [importOpen, setImportOpen] = useState(false);
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
      dependents: Math.max(0, Number(data.dependents) || 0),
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
      useCustomLatePolicy: !!data.useCustomLatePolicy,
      customLateGraceCount: data.useCustomLatePolicy
        ? Number(data.customLateGraceCount) || 0
        : null,
      customLateDeductionType: data.useCustomLatePolicy
        ? data.customLateDeductionType || "fixed"
        : null,
      customLateDeductionValue: data.useCustomLatePolicy
        ? Number(data.customLateDeductionValue) || 0
        : null,
      useCustomUlPolicy: !!data.useCustomUlPolicy,
      customUlDeductionType: data.useCustomUlPolicy
        ? data.customUlDeductionType || "fullDay"
        : null,
      customUlDeductionValue: data.useCustomUlPolicy
        ? Number(data.customUlDeductionValue) || 0
        : null,
      messagesDisabled: !!data.messagesDisabled,
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
            placeholder={t.emps.searchPlaceholder}
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
            <option value="">{t.emps.allBranches}</option>
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
        <Button variant="ghost" onClick={() => setImportOpen(true)}>
          <Upload size={15} /> {t.emps.importBtn}
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
                  onPopupBlocked: () => setBadgePopupBlocked(true),
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
            <Button
              size="sm"
              variant="ghost"
              style={{ width: "100%", justifyContent: "center", marginTop: 8 }}
              onClick={() => setCertFor(e)}
            >
              <Award size={13} /> {t.cert.btn}
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
            {t.noResults}
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
          text={t.emps.confirmDelWithName(confirmDel.name)}
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
      {certFor && (
        <CertificateModal
          emp={certFor}
          deptLabel={deptName(certFor.deptId)}
          onClose={() => setCertFor(null)}
        />
      )}
      {badgePopupBlocked && (
        <LoginActAlertDialog
          title={t.popupBlockedTitle}
          message={t.popupBlockedBadge}
          onClose={() => setBadgePopupBlocked(false)}
        />
      )}
      {importOpen && (
        <ImportEmployeesModal
          departments={departments}
          shifts={shifts}
          offices={offices}
          employees={employees}
          onClose={() => setImportOpen(false)}
          onImport={(newEmployees) => {
            setEmployees([...employees, ...newEmployees]);
            pushToast(
              t.emps.importSuccessToast(newEmployees.length),
              "success",
            );
            setImportOpen(false);
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
      <Field label={t.depts.nameLabel}>
        <Input
          value={f.name}
          onChange={set("name")}
          placeholder={t.depts.namePlaceholder}
        />
      </Field>
      <Field label={t.depts.codeLabel}>
        <Input
          value={f.code}
          onChange={set("code")}
          placeholder={t.depts.codePlaceholder}
        />
      </Field>
      <Field label={t.depts.descLabel}>
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
          text={t.depts.confirmDelWithCount(
            confirmDel.name,
            countIn(confirmDel.id),
          )}
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
   Assets (company equipment tracked and optionally assigned to a
   specific employee — laptops, phones, vehicles, tools, etc.)
----------------------------------------------------------------*/
const ASSET_STATUSES = ["available", "assigned", "maintenance", "retired"];
const ASSET_STATUS_COLOR = {
  available: { bg: "#E6F4EA", fg: "#1E7B3C" },
  assigned: { bg: "#E8F0FE", fg: "#1A56C4" },
  maintenance: { bg: "#FFF4E0", fg: "#B5720B" },
  retired: { bg: "#F1F1F1", fg: "#666" },
};
function AssetStatusPill({ status }) {
  const { t } = useLang();
  const labelKey = {
    available: "statusAvailable",
    assigned: "statusAssigned",
    maintenance: "statusMaintenance",
    retired: "statusRetired",
  }[status];
  const c = ASSET_STATUS_COLOR[status] || ASSET_STATUS_COLOR.available;
  return (
    <span className="wf-badge" style={{ background: c.bg, color: c.fg }}>
      {t.assets[labelKey] || status}
    </span>
  );
}
function AssetForm({ initial, employees, onSave, onCancel }) {
  const { t } = useLang();
  const [f, setF] = useState(
    initial || {
      name: "",
      category: "",
      serial: "",
      assignedTo: "",
      status: "available",
      purchaseDate: "",
      notes: "",
    },
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <div>
      <Field label={t.assets.name}>
        <Input
          value={f.name}
          onChange={set("name")}
          placeholder={t.assets.namePlaceholder}
        />
      </Field>
      <Field label={t.assets.category}>
        <Input
          value={f.category}
          onChange={set("category")}
          placeholder={t.assets.categoryPlaceholder}
        />
      </Field>
      <Field label={t.assets.serial}>
        <Input
          value={f.serial}
          onChange={set("serial")}
          placeholder={t.assets.serialPlaceholder}
        />
      </Field>
      <Field label={t.assets.status}>
        <Select value={f.status} onChange={set("status")}>
          {ASSET_STATUSES.map((s) => (
            <option key={s} value={s}>
              {
                t.assets[
                  {
                    available: "statusAvailable",
                    assigned: "statusAssigned",
                    maintenance: "statusMaintenance",
                    retired: "statusRetired",
                  }[s]
                ]
              }
            </option>
          ))}
        </Select>
      </Field>
      <Field label={t.assets.assignedTo}>
        <Select value={f.assignedTo || ""} onChange={set("assignedTo")}>
          <option value="">{t.assets.unassigned}</option>
          {employees.map((e) => (
            <option key={e.id} value={e.id}>
              {e.name}
            </option>
          ))}
        </Select>
      </Field>
      <Field label={t.assets.purchaseDate}>
        <DatePicker value={f.purchaseDate} onChange={set("purchaseDate")} />
      </Field>
      <Field label={t.assets.notes}>
        <textarea
          className="wf-input"
          rows={2}
          value={f.notes}
          onChange={set("notes")}
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
function Assets({ assets, setAssets, employees, isSuperAdmin }) {
  const { t } = useLang();
  const [modal, setModal] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const [query, setQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const empName = (id) => employees.find((e) => e.id === id)?.name || "";
  const filtered = useMemo(
    () =>
      assets.filter(
        (a) =>
          (a.name + a.category + a.serial + empName(a.assignedTo))
            .toLowerCase()
            .includes(query.toLowerCase()) &&
          (!statusFilter || a.status === statusFilter),
      ),
    [assets, employees, query, statusFilter],
  );
  const save = (data) => {
    if (modal.mode === "add")
      setAssets([...assets, { ...data, id: uid("as") }]);
    else setAssets(assets.map((a) => (a.id === data.id ? data : a)));
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
            placeholder={t.assets.searchPlaceholder}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
        <Select
          style={{ maxWidth: 180 }}
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
        >
          <option value="">{t.assets.allStatus}</option>
          {ASSET_STATUSES.map((s) => (
            <option key={s} value={s}>
              {
                t.assets[
                  {
                    available: "statusAvailable",
                    assigned: "statusAssigned",
                    maintenance: "statusMaintenance",
                    retired: "statusRetired",
                  }[s]
                ]
              }
            </option>
          ))}
        </Select>
        <Button
          variant="ghost"
          onClick={() =>
            exportCsv(
              `assets-${todayStr()}.csv`,
              [
                t.assets.name,
                t.assets.category,
                t.assets.serial,
                t.assets.assignedTo,
                t.assets.status,
                t.assets.purchaseDate,
              ],
              filtered.map((a) => [
                a.name,
                a.category,
                a.serial,
                a.assignedTo ? empName(a.assignedTo) : t.assets.unassigned,
                t.assets[
                  {
                    available: "statusAvailable",
                    assigned: "statusAssigned",
                    maintenance: "statusMaintenance",
                    retired: "statusRetired",
                  }[a.status]
                ],
                a.purchaseDate,
              ]),
            )
          }
        >
          <Download size={15} /> {t.exportCsv}
        </Button>
        <Button variant="accent" onClick={() => setModal({ mode: "add" })}>
          <Plus size={15} /> {t.assets.addBtn}
        </Button>
      </div>

      {filtered.length === 0 ? (
        <Card style={{ padding: 24, textAlign: "center", color: T.muted }}>
          {t.assets.noAssets}
        </Card>
      ) : (
        <div
          className="wf-grid"
          style={{
            gridTemplateColumns: "repeat(auto-fill, minmax(260px,1fr))",
          }}
        >
          {filtered.map((a) => (
            <Card key={a.id} accent={colorFor(a.name)} style={{ padding: 16 }}>
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
                    <Package size={18} color={T.ink} />
                  </div>
                  <div>
                    <div
                      style={{ fontWeight: 600, color: T.ink, fontSize: 13 }}
                    >
                      {a.name}
                    </div>
                    <div
                      style={{
                        fontSize: 11,
                        color: T.muted,
                        fontFamily: "'JetBrains Mono',monospace",
                      }}
                    >
                      {a.serial || "—"}
                    </div>
                  </div>
                </div>
                <AssetStatusPill status={a.status} />
              </div>
              <div
                style={{
                  fontSize: 12,
                  color: T.textSoft,
                  marginTop: 12,
                  display: "flex",
                  flexDirection: "column",
                  gap: 4,
                }}
              >
                <div>🏷️ {a.category || "—"}</div>
                <div>
                  👤{" "}
                  {a.assignedTo ? empName(a.assignedTo) : t.assets.unassigned}
                </div>
                {a.purchaseDate && <div>📅 {a.purchaseDate}</div>}
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
                  onClick={() => setModal({ mode: "edit", data: a })}
                >
                  <Pencil size={13} /> {t.edit}
                </Button>
                {isSuperAdmin && (
                  <Button
                    size="sm"
                    variant="danger"
                    style={{ flex: 1, justifyContent: "center" }}
                    onClick={() => setConfirmDel(a)}
                  >
                    <Trash2 size={13} /> {t.delete}
                  </Button>
                )}
              </div>
            </Card>
          ))}
        </div>
      )}

      {modal && (
        <Modal
          title={modal.mode === "add" ? t.assets.addTitle : t.assets.editTitle}
          onClose={() => setModal(null)}
        >
          <AssetForm
            initial={modal.data}
            employees={employees}
            onSave={save}
            onCancel={() => setModal(null)}
          />
        </Modal>
      )}
      {confirmDel && (
        <ConfirmDialog
          text={t.assets.confirmDelWithName(confirmDel.name)}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setAssets(assets.filter((a) => a.id !== confirmDel.id));
            setConfirmDel(null);
          }}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Training & Certifications
----------------------------------------------------------------*/
const TRAINING_STATUSES = ["planned", "ongoing", "completed", "cancelled"];
const TRAINING_STATUS_COLOR = {
  planned: { bg: "#EEF2FF", fg: "#4338CA" },
  ongoing: { bg: "#FFF7ED", fg: "#C2410C" },
  completed: { bg: "#ECFDF5", fg: "#047857" },
  cancelled: { bg: "#F3F4F6", fg: "#6B7280" },
};
const TRAINING_STATUS_LABEL_KEY = {
  planned: "statusPlanned",
  ongoing: "statusOngoing",
  completed: "statusCompleted",
  cancelled: "statusCancelled",
};
function TrainingStatusPill({ status }) {
  const { t } = useLang();
  const c = TRAINING_STATUS_COLOR[status] || TRAINING_STATUS_COLOR.planned;
  const labelKey = TRAINING_STATUS_LABEL_KEY[status] || "statusPlanned";
  return (
    <span className="wf-badge" style={{ background: c.bg, color: c.fg }}>
      {t.training[labelKey] || status}
    </span>
  );
}
// Reuses the same expiry-window logic/colors already built for document
// expiry tracking (getDocExpiryInfo / DOC_EXPIRY_COLOR) so a training
// certificate that's about to lapse (e.g. first-aid, food-safety) shows
// the same visual language as contract/document expiry elsewhere in the
// app, instead of inventing a second parallel "expiring soon" system.
function TrainingCertBadge({ expiryDate }) {
  const { t } = useLang();
  const info = getDocExpiryInfo(expiryDate);
  if (!info) return null;
  const c = DOC_EXPIRY_COLOR[info.status];
  const label =
    info.status === "expired"
      ? t.training.expired
      : info.status === "soon"
        ? t.training.expiringSoon(info.daysLeft)
        : `${t.training.expiresOn} ${fmtDateDisplay(expiryDate)}`;
  return (
    <span
      className="wf-badge"
      style={{ background: c.bg, color: c.fg, whiteSpace: "nowrap" }}
    >
      {label}
    </span>
  );
}
function TrainingForm({ initial, employees, onSave, onCancel }) {
  const { t } = useLang();
  const [f, setF] = useState(
    initial || {
      employeeId: employees[0]?.id || "",
      courseName: "",
      provider: "",
      category: "",
      startDate: "",
      completionDate: "",
      certExpiry: "",
      status: "planned",
      notes: "",
    },
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <div>
      <Field label={t.training.employee}>
        <Select value={f.employeeId} onChange={set("employeeId")}>
          {employees.map((e) => (
            <option key={e.id} value={e.id}>
              {e.name}
            </option>
          ))}
        </Select>
      </Field>
      <Field label={t.training.courseName}>
        <Input
          value={f.courseName}
          onChange={set("courseName")}
          placeholder={t.training.courseNamePlaceholder}
        />
      </Field>
      <Field label={t.training.provider}>
        <Input
          value={f.provider}
          onChange={set("provider")}
          placeholder={t.training.providerPlaceholder}
        />
      </Field>
      <Field label={t.training.category}>
        <Input
          value={f.category}
          onChange={set("category")}
          placeholder={t.training.categoryPlaceholder}
        />
      </Field>
      <Field label={t.training.status}>
        <Select value={f.status} onChange={set("status")}>
          {TRAINING_STATUSES.map((s) => (
            <option key={s} value={s}>
              {t.training[TRAINING_STATUS_LABEL_KEY[s]]}
            </option>
          ))}
        </Select>
      </Field>
      <Field label={t.training.startDate}>
        <DatePicker value={f.startDate} onChange={set("startDate")} />
      </Field>
      <Field label={t.training.completionDate}>
        <DatePicker value={f.completionDate} onChange={set("completionDate")} />
      </Field>
      <Field label={t.training.certExpiry}>
        <DatePicker value={f.certExpiry} onChange={set("certExpiry")} />
      </Field>
      <Field label={t.training.notes}>
        <textarea
          className="wf-input"
          rows={2}
          value={f.notes}
          onChange={set("notes")}
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
          disabled={!f.courseName || !f.employeeId}
        >
          {t.save}
        </Button>
      </div>
    </div>
  );
}
function Trainings({ trainings, setTrainings, employees, isSuperAdmin }) {
  const { t } = useLang();
  const [modal, setModal] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const [query, setQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [empFilter, setEmpFilter] = useState("");
  const empName = (id) => employees.find((e) => e.id === id)?.name || "";
  const filtered = useMemo(
    () =>
      trainings
        .filter(
          (r) =>
            (r.courseName + r.provider + empName(r.employeeId))
              .toLowerCase()
              .includes(query.toLowerCase()) &&
            (!statusFilter || r.status === statusFilter) &&
            (!empFilter || r.employeeId === empFilter),
        )
        .sort((a, b) => (b.startDate || "").localeCompare(a.startDate || "")),
    [trainings, employees, query, statusFilter, empFilter],
  );
  const save = (data) => {
    if (modal.mode === "add")
      setTrainings([...trainings, { ...data, id: uid("tr") }]);
    else setTrainings(trainings.map((r) => (r.id === data.id ? data : r)));
    setModal(null);
  };
  return (
    <div>
      <div style={{ marginBottom: 14 }}>
        <div style={{ fontSize: 13, color: T.muted }}>
          {t.training.subtitle}
        </div>
      </div>
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
            placeholder={t.training.searchPlaceholder}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
        <Select
          style={{ maxWidth: 180 }}
          value={empFilter}
          onChange={(e) => setEmpFilter(e.target.value)}
        >
          <option value="">{t.training.allEmployees}</option>
          {employees.map((e) => (
            <option key={e.id} value={e.id}>
              {e.name}
            </option>
          ))}
        </Select>
        <Select
          style={{ maxWidth: 160 }}
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
        >
          <option value="">{t.training.allStatus}</option>
          {TRAINING_STATUSES.map((s) => (
            <option key={s} value={s}>
              {t.training[TRAINING_STATUS_LABEL_KEY[s]]}
            </option>
          ))}
        </Select>
        <Button
          variant="ghost"
          onClick={() =>
            exportCsv(
              `training-${todayStr()}.csv`,
              [
                t.training.employee,
                t.training.courseName,
                t.training.provider,
                t.training.category,
                t.training.status,
                t.training.startDate,
                t.training.completionDate,
                t.training.certExpiry,
              ],
              filtered.map((r) => [
                empName(r.employeeId),
                r.courseName,
                r.provider,
                r.category,
                t.training[TRAINING_STATUS_LABEL_KEY[r.status]],
                r.startDate,
                r.completionDate,
                r.certExpiry,
              ]),
            )
          }
        >
          <Download size={15} /> {t.exportCsv}
        </Button>
        <Button variant="accent" onClick={() => setModal({ mode: "add" })}>
          <Plus size={15} /> {t.training.addBtn}
        </Button>
      </div>

      {filtered.length === 0 ? (
        <Card style={{ padding: 24, textAlign: "center", color: T.muted }}>
          {t.training.noTrainings}
        </Card>
      ) : (
        <div
          className="wf-grid"
          style={{
            gridTemplateColumns: "repeat(auto-fill, minmax(280px,1fr))",
          }}
        >
          {filtered.map((r) => (
            <Card
              key={r.id}
              accent={colorFor(r.courseName)}
              style={{ padding: 16 }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "space-between",
                  gap: 8,
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
                      flexShrink: 0,
                    }}
                  >
                    <GraduationCap size={18} color={T.ink} />
                  </div>
                  <div>
                    <div
                      style={{ fontWeight: 600, color: T.ink, fontSize: 13 }}
                    >
                      {r.courseName}
                    </div>
                    <div style={{ fontSize: 11, color: T.muted }}>
                      {empName(r.employeeId)}
                    </div>
                  </div>
                </div>
                <TrainingStatusPill status={r.status} />
              </div>
              <div
                style={{
                  fontSize: 12,
                  color: T.textSoft,
                  marginTop: 12,
                  display: "flex",
                  flexDirection: "column",
                  gap: 4,
                }}
              >
                {r.provider && <div>🏫 {r.provider}</div>}
                {r.category && <div>🏷️ {r.category}</div>}
                {(r.startDate || r.completionDate) && (
                  <div>
                    📅 {r.startDate || "—"} → {r.completionDate || "—"}
                  </div>
                )}
              </div>
              {r.certExpiry && (
                <div style={{ marginTop: 10 }}>
                  <TrainingCertBadge expiryDate={r.certExpiry} />
                </div>
              )}
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
                  onClick={() => setModal({ mode: "edit", data: r })}
                >
                  <Pencil size={13} /> {t.edit}
                </Button>
                {isSuperAdmin && (
                  <Button
                    size="sm"
                    variant="danger"
                    style={{ flex: 1, justifyContent: "center" }}
                    onClick={() => setConfirmDel(r)}
                  >
                    <Trash2 size={13} /> {t.delete}
                  </Button>
                )}
              </div>
            </Card>
          ))}
        </div>
      )}

      {modal && (
        <Modal
          title={
            modal.mode === "add" ? t.training.addTitle : t.training.editTitle
          }
          onClose={() => setModal(null)}
        >
          <TrainingForm
            initial={modal.data}
            employees={employees}
            onSave={save}
            onCancel={() => setModal(null)}
          />
        </Modal>
      )}
      {confirmDel && (
        <ConfirmDialog
          text={t.training.confirmDelWithName(confirmDel.courseName)}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setTrainings(trainings.filter((r) => r.id !== confirmDel.id));
            setConfirmDel(null);
          }}
        />
      )}
    </div>
  );
}
// Read-only self-service view: an employee sees only their own training
// and certification history, newest first, with the same expiry badge
// used on the admin side so an approaching certificate renewal is just
// as visible to the person who holds it as it is to HR.
function MyTrainings({ currentEmp, trainings }) {
  const { t } = useLang();
  const mine = useMemo(
    () =>
      trainings
        .filter((r) => r.employeeId === currentEmp?.id)
        .sort((a, b) => (b.startDate || "").localeCompare(a.startDate || "")),
    [trainings, currentEmp],
  );
  return (
    <div>
      {mine.length === 0 ? (
        <Card style={{ padding: 24, textAlign: "center", color: T.muted }}>
          {t.training.noMyTrainings}
        </Card>
      ) : (
        <div
          className="wf-grid"
          style={{
            gridTemplateColumns: "repeat(auto-fill, minmax(280px,1fr))",
          }}
        >
          {mine.map((r) => (
            <Card
              key={r.id}
              accent={colorFor(r.courseName)}
              style={{ padding: 16 }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "space-between",
                  gap: 8,
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
                      flexShrink: 0,
                    }}
                  >
                    <GraduationCap size={18} color={T.ink} />
                  </div>
                  <div style={{ fontWeight: 600, color: T.ink, fontSize: 13 }}>
                    {r.courseName}
                  </div>
                </div>
                <TrainingStatusPill status={r.status} />
              </div>
              <div
                style={{
                  fontSize: 12,
                  color: T.textSoft,
                  marginTop: 12,
                  display: "flex",
                  flexDirection: "column",
                  gap: 4,
                }}
              >
                {r.provider && <div>🏫 {r.provider}</div>}
                {r.category && <div>🏷️ {r.category}</div>}
                {(r.startDate || r.completionDate) && (
                  <div>
                    📅 {r.startDate || "—"} → {r.completionDate || "—"}
                  </div>
                )}
                {r.notes && <div>📝 {r.notes}</div>}
              </div>
              {r.certExpiry && (
                <div style={{ marginTop: 10 }}>
                  <TrainingCertBadge expiryDate={r.certExpiry} />
                </div>
              )}
            </Card>
          ))}
        </div>
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
      <Field label={t.sh.nameLabel}>
        <Input
          value={f.name}
          onChange={set("name")}
          placeholder={t.sh.namePlaceholder}
        />
      </Field>
      <div className="wf-grid-2">
        <Field label={t.sh.startLabel}>
          <TimePicker value={f.start} onChange={set("start")} />
        </Field>
        <Field label={t.sh.endLabel}>
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
          {t.sh.overnightHint}
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
                    {isOvernightShift(s) ? ` ${t.sh.overnightTag}` : ""}
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
          title={modal.mode === "add" ? t.sh.addTitle : t.sh.editTitle}
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
          text={t.sh.confirmDelWithCount(
            confirmDel.name,
            countIn(confirmDel.id),
          )}
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
/* ---------------------------------------------------------------
   QR scan modal — lets an employee check in/out by scanning a
   branch's printed QR code instead of relying on GPS. Prefers the
   browser-native BarcodeDetector API when available; falls back to
   loading the small "jsQR" library from a CDN at runtime (no extra
   build dependency required) for browsers that lack it (e.g. Safari).
----------------------------------------------------------------*/
let jsQRLoadPromise = null;
function loadJsQR() {
  if (window.jsQR) return Promise.resolve(window.jsQR);
  if (!jsQRLoadPromise) {
    jsQRLoadPromise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.js";
      script.async = true;
      script.onload = () => resolve(window.jsQR);
      script.onerror = () => reject(new Error("jsqr-load-failed"));
      document.head.appendChild(script);
    });
  }
  return jsQRLoadPromise;
}

// Short two-tone "success" chime played the instant a QR scan matches a
// branch (see handleDecoded below) — built with the Web Audio API rather
// than an audio file so it needs no extra asset/network request and
// Short "success" chime played the instant a QR scan matches a branch
// (see handleDecoded below) — built with the Web Audio API rather than
// an audio file so it needs no extra asset/network request and still
// works on an offline-for-a-moment kiosk tablet. Silently no-ops if
// A brand-new AudioContext always starts life "suspended" under every
// browser's autoplay policy, and only resumes once there has been real
// user activation (a click/tap/keydown) on the page. That's harmless for
// the QR scan beep and the chat "ding" — both fire as a direct result of
// something the user just did (scanning, opening Messages) — but it's
// exactly what breaks the *incoming call* ringtone: that one has to play
// the instant a Realtime broadcast/push arrives, with no click anywhere
// near it, so a context created fresh at that moment can stay silently
// suspended and the callee never hears it ring even though the on-screen
// "incoming call" state is correct.
//
// Fix: keep ONE AudioContext for the whole app's lifetime instead of a
// new one per sound, created lazily on first use, and explicitly
// `.resume()` it before every single playTones() call (resume() is a
// no-op if it's already running). A context that has already been
// resumed once via a real user gesture stays unlocked for the rest of
// the page's life, so as long as the person has tapped/clicked
// *anywhere* on the page since it loaded — which they always will have,
// just by logging in — later resume() calls made from a push/broadcast
// handler succeed even with no gesture directly attached to them.
let sharedAudioCtx = null;
function getAudioCtx() {
  const Ctx = window.AudioContext || window.webkitAudioContext;
  if (!Ctx) return null;
  if (!sharedAudioCtx || sharedAudioCtx.state === "closed") {
    sharedAudioCtx = new Ctx();
  }
  return sharedAudioCtx;
}
// Proactively unlock the shared context on the very first interaction
// with the page (well before any call can come in), so a later
// resume() from an unattended event — like an incoming-call ring — has
// the best chance of actually being allowed to play. Best-effort only:
// if this never fires (e.g. the push notification opens a brand-new,
// still-untouched tab), playTones()'s own resume() call below is still
// attempted and simply may not produce sound on that first cold open,
// per browser policy — nothing else here depends on it succeeding.
if (typeof document !== "undefined") {
  const unlockAudio = () => {
    const ctx = getAudioCtx();
    if (ctx && ctx.state === "suspended") ctx.resume().catch(() => {});
    document.removeEventListener("pointerdown", unlockAudio);
    document.removeEventListener("keydown", unlockAudio);
  };
  document.addEventListener("pointerdown", unlockAudio, { passive: true });
  document.addEventListener("keydown", unlockAudio);
}
// Shared WebAudio tone player used by the QR scan beep, the chat
// notification sound, and the call ringtone below. Silently no-ops if
// AudioContext is unavailable or blocked (e.g. autoplay policy) — the
// on-screen state already confirms the event on its own, so a missing
// beep is never the only signal.
function playTones(tones) {
  try {
    const ctx = getAudioCtx();
    if (!ctx || !tones || tones.length === 0) return;
    const schedule = () => {
      const now = ctx.currentTime;
      tones.forEach(({ freq, start, dur, type }) => {
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = type || "sine";
        osc.frequency.value = freq;
        gain.gain.setValueAtTime(0, now + start);
        gain.gain.linearRampToValueAtTime(0.35, now + start + 0.01);
        gain.gain.exponentialRampToValueAtTime(0.0001, now + start + dur);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start(now + start);
        osc.stop(now + start + dur + 0.02);
      });
    };
    // Always attempt resume() first — required for a context that's
    // still suspended, and a harmless no-op otherwise. This is what
    // actually plays the ringtone for an incoming call that arrived
    // with no user gesture attached to it, as long as the page had
    // *any* earlier interaction (see unlockAudio above).
    if (ctx.state === "suspended") {
      ctx
        .resume()
        .then(schedule)
        .catch(() => {});
    } else {
      schedule();
    }
  } catch {
    // Ignore — see comment above.
  }
}
// `mode` ("in"/"out") picks which half of the preset plays, so
// check-in and check-out stay distinguishable by ear. `preset` selects
// the tone recipe from SOUND_PRESETS — defaults to "chime" when unset.
function playScanBeep(mode, preset) {
  const recipe = SOUND_PRESETS[preset] || SOUND_PRESETS.chime;
  playTones((mode === "out" ? recipe.out : recipe.in) || []);
}
// Short two-note "ding" for an incoming chat message — same sound
// whichever side receives it (admin gets a message from an employee, or
// an employee gets one from admin), independent of the QR sound preset.
function playChatNotifySound() {
  playTones([
    { freq: 784, start: 0, dur: 0.08, type: "sine" },
    { freq: 1174.66, start: 0.09, dur: 0.16, type: "sine" },
  ]);
}
// Repeating two-tone ring for an incoming/outgoing voice call — same
// playTones() primitive as the chat "ding", just looped on an interval
// until stopped. Returns a stop function; safe to call multiple times.
function startRingtoneLoop() {
  const ring = () =>
    playTones([
      { freq: 523.25, start: 0, dur: 0.32, type: "sine" },
      { freq: 659.25, start: 0.36, dur: 0.32, type: "sine" },
    ]);
  ring();
  const intervalId = setInterval(ring, 1600);
  return () => clearInterval(intervalId);
}

function QrScanModal({ offices, mode, soundPreset, onMatch, onClose }) {
  const { t } = useLang();
  const videoRef = useRef(null);
  const canvasRef = useRef(null);
  const streamRef = useRef(null);
  const rafRef = useRef(null);
  const [error, setError] = useState("");
  const [ready, setReady] = useState(false); // true once the video has real frames to scan

  useEffect(() => {
    let cancelled = false;

    const stop = () => {
      if (rafRef.current) cancelAnimationFrame(rafRef.current);
      if (streamRef.current) {
        streamRef.current.getTracks().forEach((tr) => tr.stop());
        streamRef.current = null;
      }
    };

    const handleDecoded = (text) => {
      const result = evaluateOfficeQrPayload(offices, text);
      if (result.status === "ok") {
        stop();
        playScanBeep(mode, soundPreset);
        onMatch(result.office);
      } else if (result.status === "expired") {
        setError(t.att.qrExpired);
      } else {
        setError(t.att.qrNoMatch);
      }
    };

    const tickBarcodeDetector = async (bd) => {
      if (cancelled) return;
      try {
        const video = videoRef.current;
        if (video && video.readyState >= 2) {
          const codes = await bd.detect(video);
          if (codes && codes.length > 0) {
            handleDecoded(codes[0].rawValue);
            return; // handleDecoded stops the loop on a real match
          }
        }
      } catch {
        // transient decode errors between frames are expected — ignore
      }
      rafRef.current = requestAnimationFrame(() => tickBarcodeDetector(bd));
    };

    const tickJsQR = (jsQR) => {
      if (cancelled) return;
      const video = videoRef.current;
      const canvas = canvasRef.current;
      if (video && canvas && video.readyState >= 2) {
        canvas.width = video.videoWidth;
        canvas.height = video.videoHeight;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
        const imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const code = jsQR(imgData.data, imgData.width, imgData.height);
        if (code && code.data) {
          handleDecoded(code.data);
          return;
        }
      }
      rafRef.current = requestAnimationFrame(() => tickJsQR(jsQR));
    };

    // Tries the rear ("environment") camera first — the natural choice
    // for scanning something in front of you — but falls back to
    // whatever camera is available (e.g. front-only tablets/laptops)
    // instead of failing outright when the rear camera can't be
    // satisfied.
    const openCamera = async () => {
      try {
        return await navigator.mediaDevices.getUserMedia({
          video: { facingMode: { ideal: "environment" } },
        });
      } catch (err) {
        if (
          err &&
          (err.name === "OverconstrainedError" || err.name === "NotFoundError")
        ) {
          return navigator.mediaDevices.getUserMedia({ video: true });
        }
        throw err;
      }
    };

    const start = async () => {
      let stream;
      try {
        stream = await openCamera();
      } catch (err) {
        if (cancelled) return;
        setError(
          err && err.name === "NotFoundError"
            ? t.att.cameraNotFound
            : t.att.cameraDenied,
        );
        return;
      }
      if (cancelled) {
        stream.getTracks().forEach((tr) => tr.stop());
        return;
      }
      streamRef.current = stream;
      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        try {
          await videoRef.current.play();
        } catch {
          // Autoplay can be blocked until a user gesture on some
          // browsers — the video's onLoadedData handler still fires
          // once frames are actually available, so scanning starts
          // either way.
        }
      }
      try {
        if (typeof window.BarcodeDetector !== "undefined") {
          const bd = new window.BarcodeDetector({ formats: ["qr_code"] });
          tickBarcodeDetector(bd);
          return;
        }
      } catch {
        // unsupported format/config — fall through to jsQR
      }
      try {
        const jsQR = await loadJsQR();
        if (cancelled) return;
        tickJsQR(jsQR);
      } catch {
        if (cancelled) return;
        setError(t.att.cameraDenied);
      }
    };

    start();
    return () => {
      cancelled = true;
      stop();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <Modal title={t.att.scanQrTitle} onClose={onClose} width={360}>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <div
          style={{
            position: "relative",
            width: "100%",
            aspectRatio: "1 / 1",
            background: "#000",
            borderRadius: 12,
            overflow: "hidden",
          }}
        >
          <video
            ref={videoRef}
            muted
            playsInline
            onLoadedData={() => setReady(true)}
            style={{ width: "100%", height: "100%", objectFit: "cover" }}
          />
          {!ready && !error && (
            <div
              style={{
                position: "absolute",
                inset: 0,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                background: "rgba(0,0,0,0.35)",
              }}
            >
              <Loader2
                size={28}
                color="#fff"
                style={{ animation: "spin 1s linear infinite" }}
              />
            </div>
          )}
        </div>
        <canvas ref={canvasRef} style={{ display: "none" }} />
        {error ? (
          <p style={{ fontSize: 12.5, color: T.rose, textAlign: "center" }}>
            {error}
          </p>
        ) : (
          <p style={{ fontSize: 12.5, color: T.muted, textAlign: "center" }}>
            {ready ? t.att.scanQrHint : t.att.scanQrOpeningCamera}
          </p>
        )}
      </div>
    </Modal>
  );
}

function SelfPunch({
  emp,
  shift,
  attendance,
  setAttendance,
  offices,
  soundPreset,
}) {
  const { t, lang } = useLang();
  const today = todayStr();
  const rec = attendance.find(
    (a) => a.employeeId === emp.id && a.date === today,
  );
  const [locBusy, setLocBusy] = useState(false);
  const [locError, setLocError] = useState("");
  const [branchWarning, setBranchWarning] = useState("");
  const [scanOpen, setScanOpen] = useState(false);
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
            ? t.att.awayFromOffice(
                nearest.office.name,
                nearest.distance,
                nearest.office.radius,
              )
            : t.att.cannotVerifyLoc,
        );
        return null;
      }
      if (emp.officeId && match.office.id !== emp.officeId) {
        setBranchWarning(t.att.wrongBranchWarning(match.office.name));
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
      setLocError(t.att.gpsFailed);
      return null;
    } finally {
      setLocBusy(false);
    }
  };

  // `preLoc` lets a QR scan supply an already-verified branch match
  // (see QrScanModal below) so we skip the GPS geofence check. Any
  // branch-mismatch warning for a QR-sourced loc is set by the caller
  // (handleScanMatch), since translations aren't reachable in here —
  // `t` below is intentionally shadowed to mean "current time".
  const punchIn = async (preLoc) => {
    let loc = preLoc || null;
    if (!loc && hasOffices) {
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
  const punchOut = async (preLoc) => {
    let loc = preLoc || null;
    if (!loc && hasOffices) {
      loc = await verifyLocation();
      if (!loc) return;
    }
    setAttendance(
      attendance.map((a) =>
        a.id === rec.id ? { ...a, checkOut: timeNow(), checkOutLoc: loc } : a,
      ),
    );
  };

  // Called when QrScanModal decodes a QR code that matches one of the
  // configured office branches. Builds the same shaped loc object
  // verifyLocation would return (minus GPS distance, since none was
  // taken) and routes it into whichever punch is currently pending.
  const handleScanMatch = (office) => {
    setScanOpen(false);
    setLocError("");
    if (emp.officeId && office.id !== emp.officeId) {
      setBranchWarning(t.att.wrongBranchWarning(office.name));
    } else {
      setBranchWarning("");
    }
    const loc = {
      officeId: office.id,
      officeName: office.name,
      viaQr: true,
    };
    if (!rec) punchIn(loc);
    else if (!rec.checkOut) punchOut(loc);
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
        {t.att.liveClockLabel}
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
          <MapPin size={12} /> {t.att.gpsRequiredHint(offices.length)}
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
          {t.att.dayOffNote}
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
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            gap: 10,
          }}
        >
          <Button
            variant="accent"
            onClick={() => punchIn()}
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
          {hasOffices && (
            <>
              <div style={{ fontSize: 11, color: T.mutedLight }}>
                {t.att.scanQrOr}
              </div>
              <Button
                variant="ghost"
                size="sm"
                onClick={() => setScanOpen(true)}
                disabled={locBusy}
              >
                <QrCode size={13} /> {t.att.scanQrBtn}
              </Button>
            </>
          )}
        </div>
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
          <div
            style={{
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              gap: 10,
            }}
          >
            <Button
              variant="danger-solid"
              onClick={() => punchOut()}
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
              {t.att.punchOutBtn}
            </Button>
            {hasOffices && (
              <>
                <div style={{ fontSize: 11, color: T.mutedLight }}>
                  {t.att.scanQrOr}
                </div>
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={() => setScanOpen(true)}
                  disabled={locBusy}
                >
                  <QrCode size={13} /> {t.att.scanQrBtn}
                </Button>
              </>
            )}
          </div>
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
          {rec.checkInLoc?.officeName &&
          rec.checkOutLoc?.officeName &&
          rec.checkInLoc.officeName !== rec.checkOutLoc.officeName ? (
            <>
              {" "}
              ·{" "}
              <span
                style={{ display: "inline-flex", alignItems: "center", gap: 3 }}
              >
                <MapPin size={11} /> {rec.checkInLoc.officeName} →{" "}
                {rec.checkOutLoc.officeName}
              </span>
            </>
          ) : (
            (rec.checkOutLoc?.officeName || rec.checkInLoc?.officeName) && (
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
                  <MapPin size={11} />{" "}
                  {rec.checkOutLoc?.officeName || rec.checkInLoc?.officeName}
                </span>
              </>
            )
          )}
        </div>
      )}
      {scanOpen && (
        <QrScanModal
          offices={offices}
          mode={!rec ? "in" : "out"}
          soundPreset={soundPreset}
          onMatch={handleScanMatch}
          onClose={() => setScanOpen(false)}
        />
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
      setError(t.att.officeGpsFailed);
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
      setError(t.att.officeNameRequired);
      return;
    }
    if (!lat || !lng || !radius) {
      setError(t.att.officeCoordsRequired);
      return;
    }
    setError("");
    onSave({ name, lat, lng, radius });
  };

  return (
    <div>
      <Field label={t.att.officeNameLabel}>
        <Input
          value={f.name}
          onChange={set("name")}
          placeholder={t.att.officeNamePlaceholder}
        />
      </Field>
      <div className="wf-grid-2">
        <Field label="Latitude">
          <Input value={f.lat} onChange={set("lat")} placeholder="11.5564" />
        </Field>
        <Field label="Longitude">
          <Input value={f.lng} onChange={set("lng")} placeholder="104.9282" />
        </Field>
      </div>
      <Field label={t.att.officeRadiusLabel}>
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
          <MapPin size={13} /> {t.att.officeUseCurrentLoc}
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
// Fullscreen, unauthenticated display for one office's rotating check-in
// QR — the link (see officeKioskUrl above) is meant to be opened once on
// a tablet or monitor mounted at that branch's entrance and then left
// running, so it deliberately needs no admin/employee login. It reads
// `offices` (already loaded by AppInner for every portal) to find the
// matching branch and regenerates the QR image every QR_ROTATE_MS, the
// same rotation window QrScanModal verifies against.
function KioskDisplay({ officeId, offices, branding }) {
  const { t } = useLang();
  const [now, setNow] = useState(() => new Date());
  const [tick, setTick] = useState(0);
  useEffect(() => {
    const id = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(id);
  }, []);
  useEffect(() => {
    const iv = setInterval(() => setTick((n) => n + 1), QR_ROTATE_MS);
    return () => clearInterval(iv);
  }, []);
  const office = (offices || []).find((o) => o.id === officeId);
  const brandName = branding?.name?.trim() || t.appName;
  const hh = String(now.getHours()).padStart(2, "0");
  const mm = String(now.getMinutes()).padStart(2, "0");
  const ss = String(now.getSeconds()).padStart(2, "0");

  const shellStyle = {
    minHeight: "100vh",
    width: "100%",
    position: "relative",
    overflow: "hidden",
    background: BRAND.ink,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    padding: 32,
  };

  if (!office) {
    return (
      <div style={shellStyle}>
        <LoginBackground />
        <div style={{ position: "absolute", top: 20, right: 20, zIndex: 2 }}>
          <LangToggle variant="dark" />
        </div>
        <div
          style={{
            position: "relative",
            zIndex: 1,
            textAlign: "center",
            color: "#fff",
            maxWidth: 360,
          }}
        >
          <MapPin size={40} color="#8A93A8" style={{ marginBottom: 14 }} />
          <h1
            style={{
              fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
              fontSize: 19,
              fontWeight: 600,
              marginBottom: 8,
            }}
          >
            {t.att.kioskNotFoundTitle}
          </h1>
          <p style={{ fontSize: 13, color: "#A9B4C7" }}>
            {t.att.kioskNotFoundDesc}
          </p>
        </div>
      </div>
    );
  }

  const qrPayload = officeQrPayload(office);
  const qrSrc = `https://api.qrserver.com/v1/create-qr-code/?size=420x420&data=${encodeURIComponent(qrPayload)}`;

  return (
    <div style={shellStyle}>
      <LoginBackground />
      <div style={{ position: "absolute", top: 20, right: 20, zIndex: 2 }}>
        <LangToggle variant="dark" />
      </div>
      <div
        style={{
          position: "relative",
          zIndex: 1,
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          width: "100%",
          maxWidth: 520,
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 10,
            marginBottom: 26,
          }}
        >
          {branding?.logo ? (
            <img
              src={branding.logo}
              alt=""
              style={{
                width: 40,
                height: 40,
                borderRadius: 10,
                objectFit: "cover",
              }}
            />
          ) : (
            <div className="wf-logo-badge">{getInitials(brandName)}</div>
          )}
          <span
            style={{
              color: "#fff",
              fontWeight: 700,
              fontSize: 18,
              fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
            }}
          >
            {brandName}
          </span>
        </div>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 7,
            color: "#A9B4C7",
            fontSize: 13.5,
            fontWeight: 600,
            marginBottom: 8,
          }}
        >
          <Store size={15} />
          {office.name}
        </div>
        <h1
          style={{
            color: "#fff",
            fontFamily: "'Sora','Noto Sans Khmer',sans-serif",
            fontSize: 21,
            fontWeight: 600,
            marginBottom: 26,
            textAlign: "center",
          }}
        >
          {t.att.kioskScanHint}
        </h1>
        <div
          key={tick}
          style={{
            background: "#fff",
            padding: 18,
            borderRadius: 20,
            boxShadow: "0 24px 60px rgba(0,0,0,0.4)",
          }}
        >
          <img
            src={qrSrc}
            alt="QR"
            width={300}
            height={300}
            style={{ display: "block", width: 300, height: 300 }}
          />
        </div>
        <p
          style={{
            color: "#8A93A8",
            fontSize: 12.5,
            marginTop: 20,
            textAlign: "center",
          }}
        >
          {t.att.kioskRefreshHint}
        </p>
        <div
          style={{
            display: "flex",
            alignItems: "baseline",
            gap: 4,
            marginTop: 34,
            fontFamily: "'JetBrains Mono',monospace",
            color: "#fff",
            fontSize: 30,
            fontWeight: 700,
            fontVariantNumeric: "tabular-nums",
          }}
        >
          {hh}:{mm}
          <span style={{ fontSize: 16, color: "#8A93A8" }}>:{ss}</span>
        </div>
      </div>
    </div>
  );
}

function OfficeLocationSettings({ offices, setOffices }) {
  const { t } = useLang();
  const [open, setOpen] = useState(false);
  const [formOpen, setFormOpen] = useState(null); // null | "add" | <office being edited>
  const [confirmDel, setConfirmDel] = useState(null);
  const [qrOffice, setQrOffice] = useState(null); // office currently shown in the QR modal
  const [qrTick, setQrTick] = useState(0); // bumped on an interval to force the QR to re-render as it rotates

  const saveOffice = (data) => {
    if (formOpen === "add") {
      setOffices([
        ...offices,
        { ...data, id: uid("off"), qrSecret: uid("qs") },
      ]);
    } else {
      setOffices(
        offices.map((o) => (o.id === formOpen.id ? { ...o, ...data } : o)),
      );
    }
    setFormOpen(null);
  };

  // Ensures an office has a rotating-QR secret before it's used anywhere
  // (the in-app QR modal, or the kiosk display link below) — offices
  // created before this feature existed won't have one yet. Returns the
  // office (with a freshly-generated secret persisted to Supabase if it
  // was missing) so the caller always has a usable one synchronously.
  const ensureQrSecret = (office) => {
    if (office.qrSecret) return office;
    const withSecret = { ...office, qrSecret: uid("qs") };
    setOffices(offices.map((o) => (o.id === office.id ? withSecret : o)));
    return withSecret;
  };

  // Opens the QR modal for an office, generating its rotating-QR secret
  // on first use if it predates this feature (older offices won't have
  // one yet).
  const openQr = (office) => {
    setQrOffice(ensureQrSecret(office));
  };

  // Opens this office's public kiosk display link in a new tab — same
  // lazy secret-generation as openQr, so a branch that has never had its
  // in-app QR modal opened still gets a working kiosk on first use.
  const openKiosk = (office) => {
    const ready = ensureQrSecret(office);
    window.open(officeKioskUrl(ready.id), "_blank");
  };

  // While the QR modal is open, force a re-render every QR_ROTATE_MS so
  // the displayed image (recomputed from the current time window) rotates
  // in front of the admin's eyes rather than only on next open.
  useEffect(() => {
    if (!qrOffice) return;
    const iv = setInterval(() => setQrTick((n) => n + 1), QR_ROTATE_MS);
    return () => clearInterval(iv);
  }, [qrOffice]);

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
            {t.att.officeGeofenceTitle}
          </span>
        </div>
        <span
          style={{
            fontSize: 11.5,
            color: offices.length ? T.textSoft : T.mutedLight,
          }}
        >
          {offices.length > 0
            ? t.att.officeCountLabel(offices.length)
            : t.att.officeNotSet}
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
            {t.att.officeGeofenceDesc}
          </p>
          {offices.length === 0 && (
            <p style={{ fontSize: 12, color: T.mutedLight, marginBottom: 12 }}>
              {t.att.officeNoneYet}
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
              <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                <button
                  onClick={() => openQr(o)}
                  title={t.att.officeQrBtn}
                  style={{
                    background: "none",
                    border: "none",
                    cursor: "pointer",
                    color: T.muted,
                  }}
                >
                  <QrCode size={14} />
                </button>
                <button
                  onClick={() => openKiosk(o)}
                  title={t.att.openKioskBtn}
                  style={{
                    background: "none",
                    border: "none",
                    cursor: "pointer",
                    color: T.muted,
                  }}
                >
                  <Monitor size={14} />
                </button>
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
              <Plus size={13} /> {t.att.officeAddBtn}
            </Button>
          </div>
        </div>
      )}
      {formOpen && (
        <Modal
          title={
            formOpen === "add" ? t.att.officeAddBtn : t.att.officeEditTitle
          }
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
          text={t.att.officeConfirmDel(confirmDel.name)}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setOffices(offices.filter((o) => o.id !== confirmDel.id));
            setConfirmDel(null);
          }}
        />
      )}
      {qrOffice && (
        <QrModal
          key={qrTick}
          data={officeQrPayload(qrOffice)}
          title={t.att.officeQrTitle(qrOffice.name)}
          desc={t.att.officeQrDesc}
          onClose={() => setQrOffice(null)}
          footer={
            <p
              style={{
                fontSize: 11,
                color: T.mutedLight,
                textAlign: "center",
                margin: 0,
              }}
            >
              {t.att.officeQrRefreshHint}
            </p>
          }
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
      <Field label={t.employee}>
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
      <Field label={t.att.date}>
        <DatePicker value={f.date} onChange={set("date")} disabled={editing} />
      </Field>
      <Field label={t.status}>
        <Select
          value={f.status}
          onChange={(e) => {
            const status = e.target.value;
            const noTime =
              status === "absent" || status === "leave" || status === "unpaid";
            setF({
              ...f,
              status,
              checkIn: noTime ? "" : f.checkIn || "08:00",
              checkOut: noTime ? "" : f.checkOut,
            });
          }}
        >
          <option value="present">{t.att.statusPresent}</option>
          <option value="late">{t.att.statusLate}</option>
          <option value="absent">{t.att.statusAbsent}</option>
          <option value="leave">{t.att.statusLeavePaid}</option>
          <option value="unpaid">{t.att.statusUnpaid}</option>
        </Select>
      </Field>
      {f.status !== "absent" &&
        f.status !== "leave" &&
        f.status !== "unpaid" && (
          <div className="wf-grid-2">
            <Field label={t.att.inTime}>
              <TimePicker value={f.checkIn} onChange={set("checkIn")} />
            </Field>
            <Field label={t.att.outTime}>
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
  soundPreset,
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
          soundPreset={soundPreset}
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
                  <th>{t.att.date}</th>
                  <th>{t.att.inTime}</th>
                  <th>{t.att.outTime}</th>
                  <th>{t.att.branch}</th>
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
                      {t.att.noRecord}
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
                      {a.checkInLoc?.officeName &&
                      a.checkOutLoc?.officeName &&
                      a.checkInLoc.officeName !== a.checkOutLoc.officeName
                        ? `${a.checkInLoc.officeName} → ${a.checkOutLoc.officeName}`
                        : a.checkInLoc?.officeName ||
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
                [
                  t.employee,
                  t.emps.code,
                  t.att.inTime,
                  t.att.outTime,
                  t.status,
                ],
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
            <Plus size={15} /> {t.att.manualEntry}
          </Button>
        </div>
      </div>
      <Card style={{ overflowX: "auto" }}>
        <table className="wf-table">
          <thead>
            <tr>
              <th>{t.employee}</th>
              <th>{t.emps.shift}</th>
              <th>{t.att.inTime}</th>
              <th>{t.att.outTime}</th>
              <th>{t.att.branch}</th>
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
                  {rec?.checkInLoc?.officeName &&
                  rec?.checkOutLoc?.officeName &&
                  rec.checkInLoc.officeName !== rec.checkOutLoc.officeName
                    ? `${rec.checkInLoc.officeName} → ${rec.checkOutLoc.officeName}`
                    : rec?.checkInLoc?.officeName ||
                      rec?.checkOutLoc?.officeName ||
                      "—"}
                </td>
                <td>
                  {rec ? (
                    <StatusPill status={rec.status} />
                  ) : (
                    <span style={{ fontSize: 12, color: T.mutedLight }}>
                      {t.att.noData}
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
      <Field label={t.lv.type}>
        <Select value={f.type} onChange={set("type")}>
          <option value="annual">{t.lv.typeAnnual}</option>
          <option value="sick">{t.lv.typeSick}</option>
          <option value="unpaid">{t.lv.typeUnpaid}</option>
          <option value="other">{t.lv.typeOther}</option>
        </Select>
      </Field>
      {f.type === "unpaid" && (
        <p
          style={{
            fontSize: 12,
            color: T.muted,
            marginTop: -8,
            marginBottom: 12,
          }}
        >
          {t.lv.unpaidHint}
        </p>
      )}
      {f.type === "annual" && typeof remainingForType === "number" && (
        <p
          style={{
            fontSize: 12,
            color: T.muted,
            marginTop: -8,
            marginBottom: 12,
          }}
        >
          {t.lv.remainingAnnual(remainingForType)}
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
          {t.lv.remainingSick(remainingForType)}
        </p>
      )}
      <DateRangePicker
        startValue={f.startDate}
        endValue={f.endDate}
        onChangeStart={set("startDate")}
        onChangeEnd={set("endDate")}
        startLabel={t.lv.startDate}
        endLabel={t.lv.endDate}
        style={{ marginBottom: 14 }}
      />
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
          {t.lv.overQuotaWarning(
            requestedDays,
            f.type === "annual" ? t.lv.typeAnnual : t.lv.typeSick,
            remainingForType,
          )}
        </p>
      )}
      <Field label={t.lv.reason}>
        <textarea
          className="wf-input"
          rows={3}
          style={{ resize: "vertical", fontFamily: "inherit" }}
          value={f.reason}
          onChange={set("reason")}
          placeholder={t.lv.reasonPlaceholder}
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
          {t.lv.submit}
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

  // Approving a request marks every day in its range as leave in
  // attendance. Annual/sick/other leave is marked "leave" (paid, no
  // payroll deduction). "unpaid" (UL) requests are marked "unpaid" so
  // payroll docks a full day's pay for each of those days, same as an
  // unexcused absence.
  // Days the employee has already actually checked into (e.g. a same-day
  // leave request approved after they clocked in that morning, or a
  // leave range that overlaps a day they already worked) are left alone
  // — approving leave elsewhere must never erase real attendance.
  const applyLeaveToAttendance = (req) => {
    const days = dateRange(req.startDate, req.endDate);
    const status = req.type === "unpaid" ? "unpaid" : "leave";
    let next = attendance;
    for (const d of days) {
      const existing = next.find(
        (a) => a.employeeId === req.employeeId && a.date === d,
      );
      if (existing && existing.checkIn) continue;
      if (existing) {
        next = next.map((a) =>
          a.id === existing.id
            ? { ...a, status, checkIn: null, checkOut: null }
            : a,
        );
      } else {
        next = [
          ...next,
          {
            id: uid("a"),
            employeeId: req.employeeId,
            date: d,
            status,
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
        <div className="wf-grid-2" style={{ marginBottom: 16 }}>
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
                {t.lv.annualLeaveYear(new Date().getFullYear())}
              </div>
              <div style={{ fontSize: 13, color: T.textSoft }}>
                {t.lv.usedOfQuota(bal.used, bal.quota)}
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
              {t.lv.remainingDays(bal.remaining)}
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
                {t.lv.sickLeaveYear(new Date().getFullYear())}
              </div>
              <div style={{ fontSize: 13, color: T.textSoft }}>
                {t.lv.usedOfQuota(sickBal.used, sickBal.quota)}
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
              {t.lv.remainingDays(sickBal.remaining)}
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
                <th>{t.lv.fromShort}</th>
                <th>{t.lv.toShort}</th>
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
                    {t.lv.noRequest}
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
          <Modal
            title={t.lv.modalTitle}
            onClose={() => setModal(false)}
            width={620}
          >
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
              <th>{t.employee}</th>
              <th>{t.lv.type}</th>
              <th>{t.lv.fromShort}</th>
              <th>{t.lv.toShort}</th>
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
                  {t.lv.noRequest}
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
                        {t.lv.remainingDays(
                          annualLeaveBalance(emp, leaveRequests).remaining,
                        )}
                      </div>
                    )}
                    {r.type === "sick" && emp && (
                      <div style={{ fontSize: 10.5, color: T.muted }}>
                        {t.lv.remainingDays(
                          sickLeaveBalance(emp, leaveRequests).remaining,
                        )}
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
      <div className="wf-grid-2">
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
          <div className="wf-grid-2">
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
      lateGraceCount: Number(f.lateGraceCount) || 0,
      lateDeductionType: f.lateDeductionType || "fixed",
      lateDeductionValue: Number(f.lateDeductionValue) || 0,
      ulDeductionType: f.ulDeductionType || "fullDay",
      ulDeductionValue: Number(f.ulDeductionValue) || 0,
      taxMode: f.taxMode || "flat",
      exchangeRate: Number(f.exchangeRate) || 4100,
      nssfWageCapKHR: Number(f.nssfWageCapKHR) || 1200000,
      nssfPensionEmployeeRate: Number(f.nssfPensionEmployeeRate) || 0,
      nssfPensionEmployerRate: Number(f.nssfPensionEmployerRate) || 0,
      nssfOrcRate: Number(f.nssfOrcRate) || 0,
      nssfHealthRate: Number(f.nssfHealthRate) || 0,
    });
    setOpen(false);
  };
  const isKhmerMode = f.taxMode === "khmerProgressive";
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
          {payrollPolicy.taxMode === "khmerProgressive"
            ? t.pay.taxModeKhmer
            : `${payrollPolicy.taxRate}% / ${payrollPolicy.insuranceRate}%`}
          {payrollPolicy.taxMode !== "khmerProgressive" &&
          Number(payrollPolicy.minSalaryThreshold) > 0
            ? ` · ≥ ${fmtMoney(payrollPolicy.minSalaryThreshold)}`
            : ""}
          {Number(payrollPolicy.lateDeductionValue) > 0
            ? ` · ${t.pay.lateBadgeShort(Number(payrollPolicy.lateGraceCount) || 0)}`
            : ""}
          {` · ${t.pay.ulBadgeShort(payrollPolicy.ulDeductionType || "fullDay")}`}
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
          <div style={{ marginBottom: 14 }}>
            <Field label={t.pay.taxModeLabel}>
              <Select value={f.taxMode || "flat"} onChange={set("taxMode")}>
                <option value="flat">{t.pay.taxModeFlat}</option>
                <option value="khmerProgressive">{t.pay.taxModeKhmer}</option>
              </Select>
            </Field>
            {isKhmerMode && (
              <p style={{ fontSize: 11.5, color: T.muted, marginTop: 6 }}>
                {t.pay.taxModeKhmerDesc}
              </p>
            )}
          </div>
          {isKhmerMode ? (
            <div>
              <div className="wf-grid-2">
                <Field label={t.pay.exchangeRateLabel}>
                  <Input
                    type="number"
                    step="1"
                    min="1"
                    value={f.exchangeRate}
                    onChange={set("exchangeRate")}
                  />
                </Field>
                <Field label={t.pay.nssfWageCapLabel}>
                  <Input
                    type="number"
                    step="1000"
                    min="0"
                    value={f.nssfWageCapKHR}
                    onChange={set("nssfWageCapKHR")}
                  />
                </Field>
              </div>
              <p style={{ fontSize: 11.5, color: T.muted, marginTop: 6 }}>
                {t.pay.exchangeRateHint}
              </p>
              <div className="wf-grid-2" style={{ marginTop: 12 }}>
                <Field label={t.pay.nssfPensionEmployeeLabel}>
                  <Input
                    type="number"
                    step="0.1"
                    min="0"
                    value={f.nssfPensionEmployeeRate}
                    onChange={set("nssfPensionEmployeeRate")}
                  />
                </Field>
                <Field label={t.pay.nssfPensionEmployerLabel}>
                  <Input
                    type="number"
                    step="0.1"
                    min="0"
                    value={f.nssfPensionEmployerRate}
                    onChange={set("nssfPensionEmployerRate")}
                  />
                </Field>
                <Field label={t.pay.nssfOrcLabel}>
                  <Input
                    type="number"
                    step="0.1"
                    min="0"
                    value={f.nssfOrcRate}
                    onChange={set("nssfOrcRate")}
                  />
                </Field>
                <Field label={t.pay.nssfHealthLabel}>
                  <Input
                    type="number"
                    step="0.1"
                    min="0"
                    value={f.nssfHealthRate}
                    onChange={set("nssfHealthRate")}
                  />
                </Field>
              </div>
              <p style={{ fontSize: 11.5, color: T.muted, marginTop: 6 }}>
                {t.pay.nssfEmployerOnlyHint}
              </p>
            </div>
          ) : (
            <>
              <div className="wf-grid-2">
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
            </>
          )}
          <div
            style={{
              marginTop: 16,
              paddingTop: 16,
              borderTop: `1px solid ${T.lineSoft}`,
            }}
          >
            <p style={{ fontSize: 12, color: T.muted, marginBottom: 14 }}>
              {t.pay.latePolicyDesc}
            </p>
            <div className="wf-grid-2">
              <Field label={t.pay.lateGraceCountLabel}>
                <Input
                  type="number"
                  step="1"
                  min="0"
                  value={f.lateGraceCount}
                  onChange={set("lateGraceCount")}
                />
              </Field>
              <Field label={t.pay.lateDeductionTypeLabel}>
                <Select
                  value={f.lateDeductionType || "fixed"}
                  onChange={set("lateDeductionType")}
                >
                  <option value="fixed">{t.pay.lateDeductionTypeFixed}</option>
                  <option value="percentDaily">
                    {t.pay.lateDeductionTypePercent}
                  </option>
                </Select>
              </Field>
            </div>
            <div style={{ marginTop: 12 }}>
              <Field
                label={
                  f.lateDeductionType === "percentDaily"
                    ? t.pay.lateDeductionValuePercentLabel
                    : t.pay.lateDeductionValueFixedLabel
                }
              >
                <Input
                  type="number"
                  step="0.1"
                  min="0"
                  placeholder="0"
                  value={f.lateDeductionValue}
                  onChange={set("lateDeductionValue")}
                />
              </Field>
              <p style={{ fontSize: 11.5, color: T.muted, marginTop: 6 }}>
                {Number(f.lateDeductionValue) > 0
                  ? t.pay.lateDeductionHint(Number(f.lateGraceCount) || 0)
                  : t.pay.lateDeductionDisabledHint}
              </p>
            </div>
          </div>
          <div
            style={{
              marginTop: 16,
              paddingTop: 16,
              borderTop: `1px solid ${T.lineSoft}`,
            }}
          >
            <p style={{ fontSize: 12, color: T.muted, marginBottom: 14 }}>
              {t.pay.ulPolicyDesc}
            </p>
            <div className="wf-grid-2">
              <Field label={t.pay.ulDeductionTypeLabel}>
                <Select
                  value={f.ulDeductionType || "fullDay"}
                  onChange={set("ulDeductionType")}
                >
                  <option value="fullDay">
                    {t.pay.ulDeductionTypeFullDay}
                  </option>
                  <option value="fixed">{t.pay.ulDeductionTypeFixed}</option>
                  <option value="percentDaily">
                    {t.pay.ulDeductionTypePercent}
                  </option>
                </Select>
              </Field>
              {f.ulDeductionType !== "fullDay" && (
                <Field
                  label={
                    f.ulDeductionType === "percentDaily"
                      ? t.pay.ulDeductionValuePercentLabel
                      : t.pay.ulDeductionValueFixedLabel
                  }
                >
                  <Input
                    type="number"
                    step="0.1"
                    min="0"
                    placeholder="0"
                    value={f.ulDeductionValue}
                    onChange={set("ulDeductionValue")}
                  />
                </Field>
              )}
            </div>
            <p style={{ fontSize: 11.5, color: T.muted, marginTop: 6 }}>
              {t.pay.ulDeductionHint}
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
              <th>{t.employee}</th>
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
      <div className="wf-grid-2">
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
   Shift Swap Requests — an employee asks a named colleague to cover
   their shift on a specific date; admin approves or rejects (same
   single-decider flow as Attendance Corrections). Approval is a
   record/notification only — it does NOT auto-edit shiftId or the
   Shifts schedule, since shifts in this app are a static per-employee
   assignment rather than a per-date roster. Admin still updates the
   actual schedule by hand if the coverage needs to be reflected there.
----------------------------------------------------------------*/
function SsRejectModal({ onCancel, onConfirm }) {
  const { t } = useLang();
  const [reason, setReason] = useState("");
  const trimmed = reason.trim();
  return (
    <Modal title={t.ss.rejectTitle} onClose={onCancel} width={420}>
      <Field label={t.ss.rejectReason}>
        <textarea
          className="wf-input"
          rows={3}
          style={{ resize: "vertical", fontFamily: "inherit" }}
          value={reason}
          onChange={(e) => setReason(e.target.value)}
          placeholder={t.ss.rejectReasonPlaceholder}
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
          {t.ss.rejectReasonRequired}
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
          {t.ss.reject}
        </Button>
      </div>
    </Modal>
  );
}

function SsDecisionNote({ r, admins }) {
  const { t } = useLang();
  if (r.status !== "approved" && r.status !== "rejected") return null;
  const decider = admins.find((a) => a.id === r.decidedById);
  const name = r.decidedByName || decider?.name || "—";
  return (
    <div style={{ fontSize: 11.5, color: T.textSoft, marginTop: 3 }}>
      {r.status === "approved" ? (
        <span>
          {t.ss.approvedBy} <strong>{name}</strong>
        </span>
      ) : (
        <span style={{ color: T.rose }}>
          {t.ss.rejectedBy} <strong>{name}</strong>
          {r.decisionReason ? ` — ${r.decisionReason}` : ""}
        </span>
      )}
    </div>
  );
}

function ShiftSwapRequestForm({ currentEmp, shifts, onSave, onCancel }) {
  const { t } = useLang();
  const myShift = shifts?.find((s) => s.id === currentEmp?.shiftId);
  // Every shift except the one the employee is already on — swapping
  // "into" your own current shift isn't a real request.
  const otherShifts = (shifts || []).filter(
    (s) => s.id !== currentEmp?.shiftId,
  );
  const [f, setF] = useState({
    toShiftId: otherShifts[0]?.id || "",
    date: todayStr(),
    reason: "",
  });
  const set = (k) => (e) =>
    setF({ ...f, [k]: typeof e === "string" ? e : e.target.value });
  const invalid = !f.toShiftId || !f.date || !f.reason.trim();
  if (otherShifts.length === 0) {
    return (
      <div style={{ padding: "8px 0" }}>
        <p style={{ fontSize: 13, color: T.textSoft }}>{t.ss.noOtherShift}</p>
        <div
          style={{ display: "flex", justifyContent: "flex-end", marginTop: 12 }}
        >
          <Button variant="ghost" onClick={onCancel}>
            {t.cancel}
          </Button>
        </div>
      </div>
    );
  }
  return (
    <div>
      {myShift && (
        <Field label={t.ss.fromShift}>
          <div
            style={{
              fontSize: 13,
              color: T.ink,
              background: T.chipBg || "rgba(91,141,239,0.08)",
              borderRadius: 8,
              padding: "8px 12px",
            }}
          >
            {shiftLabel(myShift)}
          </div>
        </Field>
      )}
      <Field label={t.ss.toShift}>
        <Select value={f.toShiftId} onChange={set("toShiftId")}>
          {otherShifts.map((s) => (
            <option key={s.id} value={s.id}>
              {shiftLabel(s)}
            </option>
          ))}
        </Select>
      </Field>
      <Field label={t.ss.date}>
        <DatePicker value={f.date} onChange={set("date")} />
      </Field>
      <Field label={t.ss.reason}>
        <textarea
          className="wf-input"
          rows={3}
          style={{ resize: "vertical", fontFamily: "inherit" }}
          value={f.reason}
          onChange={set("reason")}
          placeholder={t.ss.reasonPlaceholder}
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
              fromShiftId: myShift?.id || null,
              toShiftId: f.toShiftId,
              date: f.date,
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

function ShiftSwapRequests({
  role,
  currentAdmin,
  currentEmp,
  employees,
  setEmployees,
  shifts,
  admins,
  shiftSwapRequests,
  setShiftSwapRequests,
  isSuperAdmin,
  canApprove,
}) {
  const { t } = useLang();
  const [modal, setModal] = useState(false);
  const [rejectFor, setRejectFor] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const empOf = (id) => employees.find((e) => e.id === id);
  const shiftOf = (id) => shifts.find((s) => s.id === id);

  // Approving doesn't just flip the request's status — it also applies
  // the requested shift straight onto the employee record, so admin
  // never has to go re-edit it by hand on the Employees/Shifts page
  // afterward.
  const approve = (req) => {
    setShiftSwapRequests(
      shiftSwapRequests.map((r) =>
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
    if (req.toShiftId) {
      setEmployees(
        employees.map((e) =>
          e.id === req.employeeId ? { ...e, shiftId: req.toShiftId } : e,
        ),
      );
    }
  };
  const reject = (req, reason) => {
    setShiftSwapRequests(
      shiftSwapRequests.map((r) =>
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
    setShiftSwapRequests([
      ...shiftSwapRequests,
      {
        id: uid("ss"),
        employeeId: currentEmp.id,
        fromShiftId: f.fromShiftId,
        toShiftId: f.toShiftId,
        date: f.date,
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

  const ShiftChangeCell = ({ r }) => (
    <div
      style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5 }}
    >
      <span style={{ color: T.textSoft }}>
        {shiftLabel(shiftOf(r.fromShiftId))}
      </span>
      <span style={{ color: T.mutedLight }}>→</span>
      <span style={{ color: T.ink, fontWeight: 500 }}>
        {shiftLabel(shiftOf(r.toShiftId))}
      </span>
    </div>
  );

  if (role !== "admin" && currentEmp) {
    const mine = shiftSwapRequests
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
            <Plus size={15} /> {t.ss.addBtn}
          </Button>
        </div>
        <Card style={{ overflowX: "auto" }}>
          <table className="wf-table">
            <thead>
              <tr>
                <th>{t.ss.date}</th>
                <th>{t.ss.shiftChange}</th>
                <th>{t.ss.reason}</th>
                <th>{t.status}</th>
              </tr>
            </thead>
            <tbody>
              {mine.length === 0 && (
                <tr>
                  <td
                    colSpan={4}
                    style={{
                      textAlign: "center",
                      color: T.muted,
                      padding: "24px 0",
                    }}
                  >
                    {t.ss.noRequest}
                  </td>
                </tr>
              )}
              {mine.map((r) => (
                <tr key={r.id}>
                  <td style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                    {r.date}
                  </td>
                  <td>
                    <ShiftChangeCell r={r} />
                  </td>
                  <td
                    style={{ fontSize: 12.5, color: T.textSoft, maxWidth: 200 }}
                  >
                    {r.reason || "—"}
                  </td>
                  <td>
                    <StatusPill status={r.status} />
                    <SsDecisionNote r={r} admins={admins} />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </Card>
        {modal && (
          <Modal title={t.ss.addBtn} onClose={() => setModal(false)}>
            <ShiftSwapRequestForm
              currentEmp={currentEmp}
              shifts={shifts}
              onSave={submit}
              onCancel={() => setModal(false)}
            />
          </Modal>
        )}
      </div>
    );
  }

  // Admin view — pending requests surfaced on top, newest first.
  const sorted = [...shiftSwapRequests].sort((a, b) => {
    if (a.status === "pending" && b.status !== "pending") return -1;
    if (a.status !== "pending" && b.status === "pending") return 1;
    return b.createdAt.localeCompare(a.createdAt);
  });
  return (
    <div>
      <p style={{ fontSize: 12.5, color: T.textSoft, marginBottom: 12 }}>
        {t.ss.note}
      </p>
      <Card style={{ overflowX: "auto" }}>
        <table className="wf-table">
          <thead>
            <tr>
              <th>{t.employee}</th>
              <th>{t.ss.date}</th>
              <th>{t.ss.shiftChange}</th>
              <th>{t.ss.reason}</th>
              <th>{t.status}</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {sorted.length === 0 && (
              <tr>
                <td
                  colSpan={6}
                  style={{
                    textAlign: "center",
                    color: T.muted,
                    padding: "24px 0",
                  }}
                >
                  {t.ss.noRequest}
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
                  <td>
                    <ShiftChangeCell r={r} />
                  </td>
                  <td
                    style={{ fontSize: 12.5, color: T.textSoft, maxWidth: 200 }}
                  >
                    {r.reason || "—"}
                  </td>
                  <td>
                    <StatusPill status={r.status} />
                    <SsDecisionNote r={r} admins={admins} />
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
                              <ThumbsUp size={13} /> {t.ss.approve}
                            </Button>
                            <Button
                              size="sm"
                              variant="danger"
                              onClick={() => setRejectFor(r)}
                            >
                              <ThumbsDown size={13} /> {t.ss.reject}
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
        <SsRejectModal
          onCancel={() => setRejectFor(null)}
          onConfirm={(reason) => reject(rejectFor, reason)}
        />
      )}
      {confirmDel && (
        <ConfirmDialog
          text={t.ss.confirmDel}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setShiftSwapRequests(
              shiftSwapRequests.filter((r) => r.id !== confirmDel.id),
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

// Days remaining until a document's expiry date (negative if already
// expired). A document with no expiryDate never shows a status.
const DOC_EXPIRY_WARN_DAYS = 30;
function getDocExpiryInfo(expiryDate) {
  if (!expiryDate) return null;
  const today = parseYMD(todayStr());
  const exp = parseYMD(expiryDate);
  if (!exp) return null;
  const daysLeft = Math.round((exp - today) / 86400000);
  if (daysLeft < 0) return { status: "expired", daysLeft };
  if (daysLeft <= DOC_EXPIRY_WARN_DAYS) return { status: "soon", daysLeft };
  return { status: "valid", daysLeft };
}
const DOC_EXPIRY_COLOR = {
  expired: { bg: "var(--wf-rose-soft)", fg: "#E5637A" },
  soon: { bg: "var(--wf-gold-soft)", fg: "#F0A83B" },
  valid: { bg: "var(--wf-forest-soft)", fg: "var(--wf-forest-text)" },
};
function DocExpiryBadge({ expiryDate }) {
  const { t } = useLang();
  const info = getDocExpiryInfo(expiryDate);
  if (!info) return null;
  const c = DOC_EXPIRY_COLOR[info.status];
  const label =
    info.status === "expired"
      ? t.doc.expired
      : info.status === "soon"
        ? t.doc.expiringSoon(info.daysLeft)
        : `${t.doc.expiresOn} ${fmtDateDisplay(expiryDate)}`;
  return (
    <span
      className="wf-badge"
      style={{ background: c.bg, color: c.fg, whiteSpace: "nowrap" }}
    >
      {label}
    </span>
  );
}

// Converts a "data:<mime>;base64,<...>" string (how documents are stored
// — see MAX_DOC_BYTES above) into a Blob, so we can hand the browser a
// blob: URL instead of navigating straight to the data: URL.
function dataUrlToBlob(dataUrl) {
  const commaIdx = dataUrl.indexOf(",");
  const header = dataUrl.slice(0, commaIdx);
  const base64 = dataUrl.slice(commaIdx + 1);
  const mimeMatch = header.match(/data:(.*?);base64/);
  const mime = mimeMatch ? mimeMatch[1] : "application/octet-stream";
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new Blob([bytes], { type: mime });
}

// Chrome (and some other browsers) can open a new tab on a `data:` URL
// but leave it blank/unpainted until something forces a repaint (e.g. the
// user manually reloads) — worse the larger the base64 string is, which
// is exactly the case for documents stored this way. Converting to a
// short-lived `blob:` URL first sidesteps that entirely and always
// renders immediately. The object URL is revoked a little while later,
// once the browser has had time to load it into the new tab/download.
function openDocBlobUrl(doc, { forceDownload } = {}) {
  let url;
  try {
    url = URL.createObjectURL(dataUrlToBlob(doc.dataUrl));
  } catch {
    // Fallback: something about the stored data URL couldn't be parsed —
    // fall back to the old (occasionally-blank-until-refresh) behavior
    // rather than doing nothing.
    url = doc.dataUrl;
  }
  if (forceDownload) {
    const a = document.createElement("a");
    a.href = url;
    a.download = doc.fileName || "file";
    document.body.appendChild(a);
    a.click();
    a.remove();
  } else {
    window.open(url, "_blank", "noopener,noreferrer");
  }
  if (url.startsWith("blob:")) {
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }
}

// Two distinct actions on a stored document, for both admin and staff:
// "View" opens the file in a new tab so the browser previews it inline
// whenever it can (PDF, images, plain text) instead of saving it to
// disk. "Download" always forces a save-to-disk regardless of file
// type. For formats the browser can't render on its own (docx, xlsx,
// ...), View still opens the tab — the browser then falls back to its
// own "can't preview, save instead" handling, same as any other website
// without a bundled document viewer.
function DocViewDownloadLinks({ doc, iconOnly }) {
  const { t } = useLang();
  const linkStyle = iconOnly
    ? {
        color: T.forestText,
        display: "flex",
        alignItems: "center",
        background: "none",
        border: "none",
        cursor: "pointer",
        padding: 0,
      }
    : {
        flex: 1,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: 6,
        fontSize: 12.5,
        fontWeight: 600,
        color: T.forestText,
        padding: "6px 0",
        background: "none",
        border: "none",
        cursor: "pointer",
      };
  return (
    <>
      <button
        type="button"
        onClick={() => openDocBlobUrl(doc)}
        style={linkStyle}
        title={t.doc.view}
      >
        <Eye size={iconOnly ? 16 : 14} />
        {!iconOnly && <span>{t.doc.view}</span>}
      </button>
      <button
        type="button"
        onClick={() => openDocBlobUrl(doc, { forceDownload: true })}
        style={linkStyle}
        title={t.doc.download}
      >
        <Download size={iconOnly ? 16 : 14} />
        {!iconOnly && <span>{t.doc.download}</span>}
      </button>
    </>
  );
}

function DocUploadRow({ emp, currentAdmin, documents, setDocuments, t }) {
  const [category, setCategory] = useState("cv");
  const [expiryDate, setExpiryDate] = useState("");
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
          expiryDate: expiryDate || "",
        },
      ]);
      setExpiryDate("");
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
      <Field label={t.doc.expiryDate}>
        <DatePicker
          value={expiryDate}
          onChange={(e) => setExpiryDate(e.target.value)}
          style={{ minWidth: 150 }}
        />
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
                {d.expiryDate && (
                  <div style={{ marginTop: 4 }}>
                    <DocExpiryBadge expiryDate={d.expiryDate} />
                  </div>
                )}
              </div>
            </div>
            <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
              <DocViewDownloadLinks doc={d} iconOnly />
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
                {d.expiryDate && (
                  <div style={{ marginTop: 4 }}>
                    <DocExpiryBadge expiryDate={d.expiryDate} />
                  </div>
                )}
              </div>
            </div>
            <div style={{ display: "flex", gap: 10, flexShrink: 0 }}>
              <DocViewDownloadLinks doc={d} iconOnly />
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Contracts & Expiring Documents — a company-wide, read-only view
   across every employee's uploaded documents that have an expiry
   date set (contracts, ID cards, work permits, etc.), so admins
   don't have to open each employee's Documents modal one by one to
   catch something about to lapse. Uploading/editing still happens
   from the per-employee Documents modal (Employees page); this page
   is a monitoring/reporting surface on top of the same `documents`
   data.
----------------------------------------------------------------*/
function DocumentExpiryPage({ documents, employees }) {
  const { t } = useLang();
  const [query, setQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const CAT_LABEL = getDocCategoryLabel(t);
  const empName = (id) => employees.find((e) => e.id === id)?.name || "";

  const rows = useMemo(() => {
    return documents
      .filter((d) => !!d.expiryDate)
      .map((d) => ({ ...d, _info: getDocExpiryInfo(d.expiryDate) }))
      .filter((d) => d._info)
      .sort((a, b) => a._info.daysLeft - b._info.daysLeft);
  }, [documents]);

  const filtered = useMemo(
    () =>
      rows.filter(
        (d) =>
          (empName(d.employeeId) + d.fileName)
            .toLowerCase()
            .includes(query.toLowerCase()) &&
          (!statusFilter || d._info.status === statusFilter),
      ),
    [rows, employees, query, statusFilter],
  );

  return (
    <div>
      <div style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13, color: T.textSoft }}>
          {t.contracts.subtitle}
        </div>
      </div>
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
            placeholder={t.contracts.searchPlaceholder}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
        <Select
          style={{ maxWidth: 200 }}
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
        >
          <option value="">{t.contracts.allStatus}</option>
          <option value="expired">{t.contracts.statusExpired}</option>
          <option value="soon">{t.contracts.statusSoon}</option>
          <option value="valid">{t.contracts.statusValid}</option>
        </Select>
        <Button
          variant="ghost"
          onClick={() =>
            exportCsv(
              `contracts-expiry-${todayStr()}.csv`,
              [
                t.contracts.colEmployee,
                t.contracts.colDocument,
                t.contracts.colCategory,
                t.contracts.colExpiry,
                t.contracts.colStatus,
              ],
              filtered.map((d) => [
                empName(d.employeeId),
                d.fileName,
                CAT_LABEL[d.category] || d.category,
                d.expiryDate,
                d._info.status === "expired"
                  ? t.contracts.statusExpired
                  : d._info.status === "soon"
                    ? t.contracts.statusSoon
                    : t.contracts.statusValid,
              ]),
            )
          }
        >
          <Download size={15} /> {t.exportCsv}
        </Button>
      </div>

      {filtered.length === 0 ? (
        <Card style={{ padding: 24, textAlign: "center", color: T.muted }}>
          <div>{t.contracts.noItems}</div>
          <div style={{ fontSize: 11.5, marginTop: 6 }}>
            {t.contracts.noItemsHint}
          </div>
        </Card>
      ) : (
        <div
          className="wf-grid"
          style={{
            gridTemplateColumns: "repeat(auto-fill, minmax(260px,1fr))",
          }}
        >
          {filtered.map((d) => (
            <Card
              key={d.id}
              accent={colorFor(empName(d.employeeId) || d.fileName)}
              style={{ padding: 16 }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "space-between",
                  gap: 8,
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
                      flexShrink: 0,
                    }}
                  >
                    <FileText size={18} color={T.ink} />
                  </div>
                  <div style={{ minWidth: 0 }}>
                    <div
                      style={{
                        fontWeight: 600,
                        color: T.ink,
                        fontSize: 13,
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {empName(d.employeeId) || "—"}
                    </div>
                    <div
                      style={{
                        fontSize: 11,
                        color: T.muted,
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {d.fileName}
                    </div>
                  </div>
                </div>
              </div>
              <div
                style={{
                  fontSize: 12,
                  color: T.textSoft,
                  marginTop: 12,
                  display: "flex",
                  flexDirection: "column",
                  gap: 6,
                }}
              >
                <div>🏷️ {CAT_LABEL[d.category] || d.category}</div>
                <div>
                  <DocExpiryBadge expiryDate={d.expiryDate} />
                </div>
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
                <DocViewDownloadLinks doc={d} />
              </div>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Recruitment — job postings + candidate pipeline. Two tabs: post
   and manage job openings, then track applicants per posting
   through the hiring stages (applied → screening → interview →
   offer → hired/rejected).
----------------------------------------------------------------*/
const JOB_TYPES = ["full_time", "part_time", "contract", "internship"];
function jobTypeLabel(type, t) {
  return (
    {
      full_time: t.recruit.jobTypeFullTime,
      part_time: t.recruit.jobTypePartTime,
      contract: t.recruit.jobTypeContract,
      internship: t.recruit.jobTypeInternship,
    }[type] || type
  );
}
const CANDIDATE_STAGES = [
  "applied",
  "screening",
  "interview",
  "offer",
  "hired",
  "rejected",
];
function stageLabel(stage, t) {
  return (
    {
      applied: t.recruit.stageApplied,
      screening: t.recruit.stageScreening,
      interview: t.recruit.stageInterview,
      offer: t.recruit.stageOffer,
      hired: t.recruit.stageHired,
      rejected: t.recruit.stageRejected,
    }[stage] || stage
  );
}
function stageColor(stage) {
  return (
    {
      applied: T.blue,
      screening: T.gold,
      interview: T.clay,
      offer: T.forest,
      hired: T.forestDark,
      rejected: T.rose,
    }[stage] || T.muted
  );
}
function StagePill({ stage, t }) {
  const c = stageColor(stage);
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        padding: "3px 10px",
        borderRadius: 999,
        fontSize: 11.5,
        fontWeight: 600,
        color: c,
        background: `${c}1a`,
        border: `1px solid ${c}33`,
        whiteSpace: "nowrap",
      }}
    >
      {stageLabel(stage, t)}
    </span>
  );
}

function JobPostingForm({ initial, departments, onSave, onCancel }) {
  const { t } = useLang();
  const [f, setF] = useState(
    initial || {
      title: "",
      departmentId: "",
      employmentType: "full_time",
      status: "open",
      description: "",
      postedDate: todayStr(),
      closingDate: "",
    },
  );
  const set = (k) => (e) =>
    setF((p) => ({
      ...p,
      [k]: e && e.target ? e.target.value : e,
    }));
  return (
    <div style={{ display: "grid", gap: 12 }}>
      <Field label={t.recruit.jobTitle}>
        <Input
          value={f.title}
          onChange={set("title")}
          placeholder={t.recruit.jobTitlePlaceholder}
        />
      </Field>
      <Field label={t.recruit.jobDept}>
        <Select value={f.departmentId || ""} onChange={set("departmentId")}>
          <option value="">—</option>
          {departments.map((d) => (
            <option key={d.id} value={d.id}>
              {d.name}
            </option>
          ))}
        </Select>
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label={t.recruit.jobType}>
          <Select value={f.employmentType} onChange={set("employmentType")}>
            {JOB_TYPES.map((jt) => (
              <option key={jt} value={jt}>
                {jobTypeLabel(jt, t)}
              </option>
            ))}
          </Select>
        </Field>
        <Field label={t.recruit.jobStatus}>
          <Select value={f.status} onChange={set("status")}>
            <option value="open">{t.recruit.jobStatusOpen}</option>
            <option value="closed">{t.recruit.jobStatusClosed}</option>
          </Select>
        </Field>
      </div>
      <Field label={t.recruit.jobDesc}>
        <textarea
          value={f.description}
          onChange={set("description")}
          rows={4}
          style={{
            width: "100%",
            borderRadius: 10,
            border: `1px solid ${T.divider}`,
            background: T.inputBg,
            color: T.text,
            padding: "10px 12px",
            fontSize: 13.5,
            fontFamily: "inherit",
            resize: "vertical",
          }}
        />
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label={t.recruit.jobPostedDate}>
          <DatePicker value={f.postedDate} onChange={set("postedDate")} />
        </Field>
        <Field label={t.recruit.jobClosingDate}>
          <DatePicker value={f.closingDate} onChange={set("closingDate")} />
        </Field>
      </div>
      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
        <Button variant="ghost" onClick={onCancel}>
          {t.cancel}
        </Button>
        <Button
          onClick={() => onSave({ ...f, id: initial?.id || uid("job") })}
          disabled={!f.title.trim()}
        >
          {t.save}
        </Button>
      </div>
    </div>
  );
}

function JobPostingsTab({
  jobPostings,
  setJobPostings,
  candidates,
  departments,
  isSuperAdmin,
  onSelectJob,
}) {
  const { t } = useLang();
  const [modal, setModal] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const deptName = (id) => departments.find((d) => d.id === id)?.name || "";
  const candidateCount = (jobId) =>
    candidates.filter((c) => c.jobPostingId === jobId).length;

  const save = (row) => {
    setJobPostings((prev) => {
      const exists = prev.some((j) => j.id === row.id);
      return exists
        ? prev.map((j) => (j.id === row.id ? row : j))
        : [...prev, row];
    });
    setModal(null);
  };

  return (
    <div>
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          marginBottom: 14,
        }}
      >
        {isSuperAdmin && (
          <Button onClick={() => setModal({ mode: "add" })}>
            <Plus size={15} /> {t.recruit.addJobBtn}
          </Button>
        )}
      </div>
      {jobPostings.length === 0 ? (
        <Card style={{ padding: 24, textAlign: "center", color: T.textSoft }}>
          {t.recruit.noJobs}
        </Card>
      ) : (
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))",
            gap: 12,
          }}
        >
          {jobPostings.map((j) => (
            <Card key={j.id} style={{ padding: 16, cursor: "pointer" }}>
              <div
                onClick={() => onSelectJob && onSelectJob(j.id)}
                style={{ marginBottom: 10 }}
              >
                <div style={{ fontWeight: 700, fontSize: 15 }}>{j.title}</div>
                <div style={{ fontSize: 12.5, color: T.textSoft }}>
                  {deptName(j.departmentId)} ·{" "}
                  {jobTypeLabel(j.employmentType, t)}
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  gap: 8,
                }}
              >
                <span
                  style={{
                    fontSize: 11.5,
                    fontWeight: 600,
                    padding: "3px 10px",
                    borderRadius: 999,
                    color: j.status === "open" ? T.forest : T.muted,
                    background:
                      j.status === "open" ? `${T.forest}1a` : `${T.muted}1a`,
                  }}
                >
                  {j.status === "open"
                    ? t.recruit.jobStatusOpen
                    : t.recruit.jobStatusClosed}
                </span>
                <span style={{ fontSize: 12, color: T.textSoft }}>
                  {t.recruit.candidatesFor(candidateCount(j.id))}
                </span>
              </div>
              {isSuperAdmin && (
                <div
                  style={{
                    display: "flex",
                    gap: 8,
                    marginTop: 12,
                    paddingTop: 12,
                    borderTop: `1px solid ${T.divider}`,
                  }}
                >
                  <Button
                    variant="ghost"
                    onClick={() => setModal({ mode: "edit", data: j })}
                  >
                    <Pencil size={14} />
                  </Button>
                  <Button variant="ghost" onClick={() => setConfirmDel(j)}>
                    <Trash2 size={14} color={T.rose} />
                  </Button>
                </div>
              )}
            </Card>
          ))}
        </div>
      )}
      {modal && (
        <Modal
          title={
            modal.mode === "add"
              ? t.recruit.addJobTitle
              : t.recruit.editJobTitle
          }
          onClose={() => setModal(null)}
        >
          <JobPostingForm
            initial={modal.data}
            departments={departments}
            onSave={save}
            onCancel={() => setModal(null)}
          />
        </Modal>
      )}
      {confirmDel && (
        <ConfirmDialog
          text={t.recruit.confirmDelJob(confirmDel.title)}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setJobPostings((prev) =>
              prev.filter((j) => j.id !== confirmDel.id),
            );
            setConfirmDel(null);
          }}
        />
      )}
    </div>
  );
}

function CandidateForm({
  initial,
  jobPostings,
  defaultJobId,
  onSave,
  onCancel,
}) {
  const { t } = useLang();
  const [f, setF] = useState(
    initial || {
      name: "",
      phone: "",
      email: "",
      jobPostingId: defaultJobId || "",
      stage: "applied",
      notes: "",
      appliedDate: todayStr(),
      resumeFileName: "",
      resumeMimeType: "",
      resumeDataUrl: "",
    },
  );
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState("");
  const resumeInputRef = useRef(null);
  const set = (k) => (e) =>
    setF((p) => ({
      ...p,
      [k]: e && e.target ? e.target.value : e,
    }));

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setError("");
    if (file.size > MAX_DOC_BYTES) {
      setError(t.doc.tooLarge);
      if (resumeInputRef.current) resumeInputRef.current.value = "";
      return;
    }
    setUploading(true);
    try {
      const dataUrl = await fileToDataUrl(file);
      setF((p) => ({
        ...p,
        resumeFileName: file.name,
        resumeMimeType: file.type || "application/octet-stream",
        resumeDataUrl: dataUrl,
      }));
    } finally {
      setUploading(false);
      if (resumeInputRef.current) resumeInputRef.current.value = "";
    }
  };

  return (
    <div style={{ display: "grid", gap: 12 }}>
      <Field label={t.recruit.candidateName}>
        <Input
          value={f.name}
          onChange={set("name")}
          placeholder={t.recruit.candidateNamePlaceholder}
        />
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label={t.recruit.candidatePhone}>
          <Input value={f.phone} onChange={set("phone")} />
        </Field>
        <Field label={t.recruit.candidateEmail}>
          <Input value={f.email} onChange={set("email")} />
        </Field>
      </div>
      <Field label={t.recruit.appliedFor}>
        <Select value={f.jobPostingId || ""} onChange={set("jobPostingId")}>
          <option value="">{t.recruit.noJobSelected}</option>
          {jobPostings.map((j) => (
            <option key={j.id} value={j.id}>
              {j.title}
            </option>
          ))}
        </Select>
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        <Field label={t.recruit.stage}>
          <Select value={f.stage} onChange={set("stage")}>
            {CANDIDATE_STAGES.map((s) => (
              <option key={s} value={s}>
                {stageLabel(s, t)}
              </option>
            ))}
          </Select>
        </Field>
        <Field label={t.recruit.appliedDate}>
          <DatePicker value={f.appliedDate} onChange={set("appliedDate")} />
        </Field>
      </div>
      <Field label={t.recruit.notes}>
        <textarea
          value={f.notes}
          onChange={set("notes")}
          rows={3}
          style={{
            width: "100%",
            borderRadius: 10,
            border: `1px solid ${T.divider}`,
            background: T.inputBg,
            color: T.text,
            padding: "10px 12px",
            fontSize: 13.5,
            fontFamily: "inherit",
            resize: "vertical",
          }}
        />
      </Field>
      <Field label={t.recruit.resume}>
        {f.resumeDataUrl ? (
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <DocViewDownloadLinks
              doc={{
                fileName: f.resumeFileName,
                mimeType: f.resumeMimeType,
                dataUrl: f.resumeDataUrl,
              }}
            />
            <Button
              variant="ghost"
              onClick={() =>
                setF((p) => ({
                  ...p,
                  resumeFileName: "",
                  resumeMimeType: "",
                  resumeDataUrl: "",
                }))
              }
            >
              <Trash2 size={14} color={T.rose} />
            </Button>
          </div>
        ) : (
          <div style={{ fontSize: 12.5, color: T.textSoft, marginBottom: 6 }}>
            {t.recruit.noResume}
          </div>
        )}
        <input
          ref={resumeInputRef}
          type="file"
          accept=".pdf,.doc,.docx,image/*"
          onChange={handleFile}
          style={{ display: "none" }}
        />
        <Button
          variant="ghost"
          disabled={uploading}
          onClick={() => resumeInputRef.current?.click()}
        >
          {uploading ? (
            <Loader2
              size={14}
              style={{ animation: "spin 1s linear infinite" }}
            />
          ) : (
            <Upload size={14} />
          )}{" "}
          {t.recruit.uploadResume}
        </Button>
        {error && (
          <div style={{ fontSize: 12, color: T.rose, marginTop: 6 }}>
            {error}
          </div>
        )}
      </Field>
      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
        <Button variant="ghost" onClick={onCancel}>
          {t.cancel}
        </Button>
        <Button
          onClick={() => onSave({ ...f, id: initial?.id || uid("cand") })}
          disabled={!f.name.trim()}
        >
          {t.save}
        </Button>
      </div>
    </div>
  );
}

function CandidatesTab({
  candidates,
  setCandidates,
  jobPostings,
  isSuperAdmin,
  selectedJobId,
}) {
  const { t } = useLang();
  const [modal, setModal] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const [query, setQuery] = useState("");
  const [stageFilter, setStageFilter] = useState("");
  const [jobFilter, setJobFilter] = useState(selectedJobId || "");

  useEffect(() => {
    if (selectedJobId) setJobFilter(selectedJobId);
  }, [selectedJobId]);

  const jobTitle = (id) => jobPostings.find((j) => j.id === id)?.title || "";

  const filtered = useMemo(
    () =>
      candidates.filter(
        (c) =>
          (c.name + c.email + c.phone)
            .toLowerCase()
            .includes(query.toLowerCase()) &&
          (!stageFilter || c.stage === stageFilter) &&
          (!jobFilter || c.jobPostingId === jobFilter),
      ),
    [candidates, query, stageFilter, jobFilter],
  );

  const save = (row) => {
    setCandidates((prev) => {
      const exists = prev.some((c) => c.id === row.id);
      return exists
        ? prev.map((c) => (c.id === row.id ? row : c))
        : [...prev, row];
    });
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
          marginBottom: 14,
          flexWrap: "wrap",
        }}
      >
        <div
          style={{
            display: "flex",
            gap: 8,
            flexWrap: "wrap",
            flex: 1,
          }}
        >
          <div style={{ position: "relative", minWidth: 200 }}>
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
              placeholder={t.recruit.searchCandidates}
              value={query}
              onChange={(e) => setQuery(e.target.value)}
            />
          </div>
          <Select
            style={{ maxWidth: 180 }}
            value={stageFilter}
            onChange={(e) => setStageFilter(e.target.value)}
          >
            <option value="">{t.recruit.allStages}</option>
            {CANDIDATE_STAGES.map((s) => (
              <option key={s} value={s}>
                {stageLabel(s, t)}
              </option>
            ))}
          </Select>
          <Select
            style={{ maxWidth: 200 }}
            value={jobFilter}
            onChange={(e) => setJobFilter(e.target.value)}
          >
            <option value="">{t.recruit.noJobSelected}</option>
            {jobPostings.map((j) => (
              <option key={j.id} value={j.id}>
                {j.title}
              </option>
            ))}
          </Select>
        </div>
        {isSuperAdmin && (
          <Button
            onClick={() =>
              setModal({ mode: "add", defaultJobId: jobFilter || "" })
            }
          >
            <Plus size={15} /> {t.recruit.addCandidateBtn}
          </Button>
        )}
      </div>
      {filtered.length === 0 ? (
        <Card style={{ padding: 24, textAlign: "center", color: T.textSoft }}>
          {t.recruit.noCandidates}
        </Card>
      ) : (
        <div style={{ display: "grid", gap: 10 }}>
          {filtered.map((c) => (
            <Card
              key={c.id}
              style={{
                padding: 14,
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                gap: 12,
                flexWrap: "wrap",
              }}
            >
              <div style={{ minWidth: 180 }}>
                <div style={{ fontWeight: 700, fontSize: 14 }}>{c.name}</div>
                <div style={{ fontSize: 12, color: T.textSoft }}>
                  {jobTitle(c.jobPostingId) || t.recruit.noJobSelected}
                </div>
              </div>
              <div style={{ fontSize: 12.5, color: T.textSoft }}>
                {c.phone} {c.email && `· ${c.email}`}
              </div>
              <StagePill stage={c.stage} t={t} />
              {isSuperAdmin && (
                <div style={{ display: "flex", gap: 6 }}>
                  <Button
                    variant="ghost"
                    onClick={() => setModal({ mode: "edit", data: c })}
                  >
                    <Pencil size={14} />
                  </Button>
                  <Button variant="ghost" onClick={() => setConfirmDel(c)}>
                    <Trash2 size={14} color={T.rose} />
                  </Button>
                </div>
              )}
            </Card>
          ))}
        </div>
      )}
      {t.recruit.hireHint && (
        <div style={{ fontSize: 11.5, color: T.textSoft, marginTop: 10 }}>
          {t.recruit.hireHint}
        </div>
      )}
      {modal && (
        <Modal
          title={
            modal.mode === "add"
              ? t.recruit.addCandidateTitle
              : t.recruit.editCandidateTitle
          }
          onClose={() => setModal(null)}
        >
          <CandidateForm
            initial={modal.data}
            defaultJobId={modal.defaultJobId}
            jobPostings={jobPostings}
            onSave={save}
            onCancel={() => setModal(null)}
          />
        </Modal>
      )}
      {confirmDel && (
        <ConfirmDialog
          text={t.recruit.confirmDelCandidate(confirmDel.name)}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setCandidates((prev) => prev.filter((c) => c.id !== confirmDel.id));
            setConfirmDel(null);
          }}
        />
      )}
    </div>
  );
}

function Recruitment({
  jobPostings,
  setJobPostings,
  candidates,
  setCandidates,
  departments,
  isSuperAdmin,
}) {
  const { t } = useLang();
  const [tab, setTab] = useState("jobs");
  const [selectedJobId, setSelectedJobId] = useState("");

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <Button
          variant={tab === "jobs" ? "accent" : "ghost"}
          onClick={() => setTab("jobs")}
        >
          {t.recruit.tabJobs}
        </Button>
        <Button
          variant={tab === "candidates" ? "accent" : "ghost"}
          onClick={() => setTab("candidates")}
        >
          {t.recruit.tabCandidates}
        </Button>
      </div>
      {tab === "jobs" ? (
        <JobPostingsTab
          jobPostings={jobPostings}
          setJobPostings={setJobPostings}
          candidates={candidates}
          departments={departments}
          isSuperAdmin={isSuperAdmin}
          onSelectJob={(id) => {
            setSelectedJobId(id);
            setTab("candidates");
          }}
        />
      ) : (
        <CandidatesTab
          candidates={candidates}
          setCandidates={setCandidates}
          jobPostings={jobPostings}
          isSuperAdmin={isSuperAdmin}
          selectedJobId={selectedJobId}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Onboarding / Offboarding — per-employee checklist for the tasks
   needed when someone joins or leaves. Standard checklists can be
   loaded as a starting point, then customized freely.
----------------------------------------------------------------*/
function onboardingTemplate(lang) {
  return lang === "en"
    ? [
        "Prepare workstation & equipment",
        "Issue staff ID badge",
        "Create company email / accounts",
        "Collect signed contract & ID copy",
        "Introduce to team & manager",
        "Explain company policies & benefits",
        "Set up payroll & bank details",
      ]
    : [
        "រៀបចំកន្លែងធ្វើការ និងឧបករណ៍",
        "ចេញកាតសម្គាល់បុគ្គលិក",
        "បង្កើតអ៊ីម៉ែល/គណនីក្រុមហ៊ុន",
        "ប្រមូលកិច្ចសន្យាដែលបានចុះហត្ថលេខា + ID",
        "ណែនាំជាមួយក្រុម និងអ្នកគ្រប់គ្រង",
        "ពន្យល់អំពីគោលការណ៍ក្រុមហ៊ុន និងអត្ថប្រយោជន៍",
        "រៀបចំប្រាក់ខែ និងព័ត៌មានធនាគារ",
      ];
}
function offboardingTemplate(lang) {
  return lang === "en"
    ? [
        "Collect company assets & ID badge",
        "Revoke system/account access",
        "Final payroll & unused leave settlement",
        "Exit interview",
        "Update team on handover",
        "Issue employment certificate (if requested)",
      ]
    : [
        "ប្រមូលទ្រព្យសម្បត្តិក្រុមហ៊ុន និងកាតសម្គាល់",
        "ដកសិទ្ធិចូលប្រើប្រព័ន្ធ/គណនី",
        "គណនាប្រាក់ខែចុងក្រោយ និងថ្ងៃឈប់សម្រាកនៅសល់",
        "សម្ភាសន៍មុនចាកចេញ",
        "ជូនដំណឹងក្រុមអំពីការប្រគល់ការងារបន្ត",
        "ចេញលិខិតបញ្ជាក់ការងារ (ប្រសិនបើស្នើសុំ)",
      ];
}

function OnboardingOffboarding({
  employees,
  onboardingTasks,
  setOnboardingTasks,
}) {
  const { t, lang } = useLang();
  const [employeeId, setEmployeeId] = useState("");
  const [type, setType] = useState("onboarding");
  const [newTask, setNewTask] = useState("");
  const [newDue, setNewDue] = useState("");
  const [confirmDel, setConfirmDel] = useState(null);

  const tasks = useMemo(
    () =>
      onboardingTasks
        .filter((tk) => tk.employeeId === employeeId && tk.type === type)
        .sort((a, b) => (a.orderIndex ?? 0) - (b.orderIndex ?? 0)),
    [onboardingTasks, employeeId, type],
  );

  const addTask = (title, dueDate = "") => {
    if (!employeeId || !title.trim()) return;
    setOnboardingTasks((prev) => [
      ...prev,
      {
        id: uid("obt"),
        employeeId,
        type,
        title: title.trim(),
        done: false,
        dueDate,
        orderIndex: tasks.length,
        createdAt: new Date().toISOString(),
      },
    ]);
  };

  const loadTemplate = () => {
    if (!employeeId) return;
    const items =
      type === "onboarding"
        ? onboardingTemplate(lang)
        : offboardingTemplate(lang);
    const existingTitles = new Set(tasks.map((tk) => tk.title));
    const toAdd = items
      .filter((title) => !existingTitles.has(title))
      .map((title, i) => ({
        id: uid("obt"),
        employeeId,
        type,
        title,
        done: false,
        dueDate: "",
        orderIndex: tasks.length + i,
        createdAt: new Date().toISOString(),
      }));
    if (toAdd.length) setOnboardingTasks((prev) => [...prev, ...toAdd]);
  };

  const toggleDone = (taskId) => {
    setOnboardingTasks((prev) =>
      prev.map((tk) => (tk.id === taskId ? { ...tk, done: !tk.done } : tk)),
    );
  };

  const doneCount = tasks.filter((tk) => tk.done).length;

  return (
    <div>
      <div style={{ fontSize: 13, color: T.textSoft, marginBottom: 16 }}>
        {t.onboard.subtitle}
      </div>
      <div
        style={{
          display: "flex",
          gap: 12,
          flexWrap: "wrap",
          marginBottom: 16,
        }}
      >
        <div style={{ minWidth: 220 }}>
          <Field label={t.onboard.selectEmployee}>
            <Select
              value={employeeId}
              onChange={(e) => setEmployeeId(e.target.value)}
            >
              <option value="">{t.onboard.chooseEmployee}</option>
              {employees.map((e) => (
                <option key={e.id} value={e.id}>
                  {e.name}
                </option>
              ))}
            </Select>
          </Field>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
          <Button
            variant={type === "onboarding" ? "accent" : "ghost"}
            onClick={() => setType("onboarding")}
          >
            {t.onboard.typeOnboarding}
          </Button>
          <Button
            variant={type === "offboarding" ? "accent" : "ghost"}
            onClick={() => setType("offboarding")}
          >
            {t.onboard.typeOffboarding}
          </Button>
        </div>
      </div>
      {!employeeId ? (
        <Card style={{ padding: 24, textAlign: "center", color: T.textSoft }}>
          {t.onboard.selectEmployeeHint}
        </Card>
      ) : (
        <Card style={{ padding: 16 }}>
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              marginBottom: 14,
            }}
          >
            <div style={{ fontSize: 13, color: T.textSoft }}>
              {t.onboard.progress(doneCount, tasks.length)}
            </div>
            {tasks.length === 0 && (
              <Button variant="ghost" onClick={loadTemplate}>
                {t.onboard.loadTemplate}
              </Button>
            )}
          </div>
          {tasks.length === 0 ? (
            <div
              style={{
                textAlign: "center",
                color: T.textSoft,
                fontSize: 13,
                padding: "16px 0",
              }}
            >
              <div>{t.onboard.noTasks}</div>
              <div style={{ fontSize: 12, marginTop: 4 }}>
                {t.onboard.noTasksHint}
              </div>
            </div>
          ) : (
            <div style={{ display: "grid", gap: 8, marginBottom: 14 }}>
              {tasks.map((tk) => (
                <div
                  key={tk.id}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 10,
                    padding: "8px 10px",
                    borderRadius: 8,
                    background: tk.done ? `${T.forest}0d` : "transparent",
                  }}
                >
                  <input
                    type="checkbox"
                    checked={tk.done}
                    onChange={() => toggleDone(tk.id)}
                    style={{ width: 16, height: 16 }}
                  />
                  <div
                    style={{
                      flex: 1,
                      fontSize: 13.5,
                      textDecoration: tk.done ? "line-through" : "none",
                      color: tk.done ? T.textSoft : T.text,
                    }}
                  >
                    {tk.title}
                  </div>
                  {tk.dueDate && (
                    <div style={{ fontSize: 11.5, color: T.textSoft }}>
                      {tk.dueDate}
                    </div>
                  )}
                  <Button variant="ghost" onClick={() => setConfirmDel(tk)}>
                    <Trash2 size={13} color={T.rose} />
                  </Button>
                </div>
              ))}
            </div>
          )}
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Input
              style={{ flex: 1, minWidth: 180 }}
              placeholder={t.onboard.taskPlaceholder}
              value={newTask}
              onChange={(e) => setNewTask(e.target.value)}
            />
            <DatePicker
              value={newDue}
              onChange={(e) => setNewDue(e && e.target ? e.target.value : e)}
            />
            <Button
              onClick={() => {
                addTask(newTask, newDue);
                setNewTask("");
                setNewDue("");
              }}
              disabled={!newTask.trim()}
            >
              <Plus size={15} /> {t.onboard.addTask}
            </Button>
          </div>
        </Card>
      )}
      {confirmDel && (
        <ConfirmDialog
          text={t.onboard.confirmDelTask}
          onCancel={() => setConfirmDel(null)}
          onConfirm={() => {
            setOnboardingTasks((prev) =>
              prev.filter((tk) => tk.id !== confirmDel.id),
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
      <Field label={t.admAcc.fullName}>
        <Input
          value={f.name}
          onChange={set("name")}
          placeholder={t.admAcc.fullNamePlaceholder}
        />
      </Field>
      <Field label={t.admAcc.username}>
        <Input
          value={f.username}
          onChange={set("username")}
          placeholder={t.admAcc.usernamePlaceholder}
        />
      </Field>
      <Field label={t.admAcc.password}>
        <Input
          value={f.password}
          onChange={set("password")}
          placeholder="••••••••"
        />
      </Field>
      <Field label={t.admAcc.roleLabel}>
        <Select value={f.role} onChange={set("role")}>
          {ADMIN_RANKS.map((rank) => (
            <option key={rank} value={rank}>
              {adminRoleLabel(rank, lang)}
            </option>
          ))}
          <option value="superadmin">
            {adminRoleLabel("superadmin", lang)} {t.admAcc.fullAccessSuffix}
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
        {t.admAcc.permsNote}
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

// Admin <-> Employee direct chat.
//
// Admin view: a searchable list of employees on the left (existing
// conversations sorted by latest activity, or a live search to start a
// new one) and the open thread on the right — collapses to a single
// pane on phones via the wf-chat-thread-open class (see the @media rule
// in CSS above), with a back button to return to the list.
//
// Employee view: just their own single thread with the admin team, full
// width — there's no one else to pick, so no list pane at all.
// Chat attachments: cap on the original upload (before any compression) so
// a single attachment can't bloat every client's sync payload, plus the
// target size for photo attachments sent in a thread. Larger than the tiny
// square profile-photo avatar since these are viewed full-size in the chat.
const MAX_CHAT_FILE_BYTES = 8 * 1024 * 1024;
const CHAT_IMAGE_MAX_DIM = 1280;
const CHAT_IMAGE_JPEG_QUALITY = 0.75;

// Like fileToCompressedAvatarDataUrl but keeps the original aspect ratio
// (no square crop) — a chat photo should look like what was actually sent.
function fileToCompressedChatImageDataUrl(
  file,
  maxDim = CHAT_IMAGE_MAX_DIM,
  quality = CHAT_IMAGE_JPEG_QUALITY,
) {
  return new Promise((resolve, reject) => {
    const rawUrl = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      const scale = Math.min(1, maxDim / Math.max(img.width, img.height));
      const w = Math.max(1, Math.round(img.width * scale));
      const h = Math.max(1, Math.round(img.height * scale));
      const canvas = document.createElement("canvas");
      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0, w, h);
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

function humanFileSize(bytes) {
  if (!bytes && bytes !== 0) return "";
  if (bytes < 1024) return `${bytes} B`;
  const units = ["KB", "MB", "GB"];
  let val = bytes / 1024;
  let i = 0;
  while (val >= 1024 && i < units.length - 1) {
    val /= 1024;
    i += 1;
  }
  return `${val.toFixed(val < 10 ? 1 : 0)} ${units[i]}`;
}

// Chrome (and other Chromium browsers) can't read correct duration
// metadata from a MediaRecorder-produced webm blob up front — it
// reports Infinity/NaN until the element has been scrubbed once, which
// is why a freshly recorded voice message shows "0:00" with no total
// time instead of "0:00 / 0:04" like a normal audio file. Seeking to a
// huge timestamp and back to 0 on load forces the browser to walk the
// whole file and compute the real duration, so the player shows it
// immediately instead of only after the user drags the seek bar once.
function fixAudioBlobDuration(e) {
  const audio = e.currentTarget;
  if (audio.duration !== Infinity && !Number.isNaN(audio.duration)) return;
  const onTimeUpdate = () => {
    audio.removeEventListener("timeupdate", onTimeUpdate);
    audio.currentTime = 0;
  };
  audio.addEventListener("timeupdate", onTimeUpdate);
  audio.currentTime = 1e101;
}

/* ---------------------------------------------------------------
   Voice calling — WebRTC audio calls over the shared call-signal
   broadcast channel above. One admin<->employee call at a time per
   employee thread, mirroring the shared-mailbox model messages already
   use: an employee "calls admin" and whichever admin is online and free
   picks it up; an admin calls one specific employee directly.

   This hook is meant to be instantiated ONCE, at the top of AppInner
   (not inside MessagesPage), so an incoming call still rings no matter
   which page the person currently has open — only the small
   <CallOverlay/> UI needs to live near the app root.

   Basic-version scope/known limits (see chat with the person who asked
   for this feature): uses only free public STUN servers, no TURN — call
   setup can fail on networks with strict NAT/firewalls (some corporate
   WiFi, some mobile carriers). Upgrading to add a TURN server later
   doesn't require touching this signaling logic, only ICE_SERVERS below.
---------------------------------------------------------------- */
const CALL_ICE_SERVERS = [
  {
    urls: ["stun:stun.l.google.com:19302", "stun:stun1.l.google.com:19302"],
  },
  // STUN alone only works when both sides can reach each other directly —
  // it fails whenever either side is behind a restrictive/symmetric NAT,
  // which is the normal case for mobile carrier networks (4G/5G) and many
  // corporate networks. Call then shows "Connected" (ICE/signaling found
  // each other) but no audio ever flows, because there's no direct path
  // and nothing to relay through. A TURN server is required as a relay
  // fallback for those cases.
  //
  // ⚠️ The credentials below are Open Relay Project's free public demo
  // TURN server — fine for testing, NOT for production (shared, rate
  // limited, no uptime guarantee). Before going live, replace this with
  // your own TURN server: either self-hosted (coturn is free/open-source)
  // or a paid provider (Twilio, Metered, Xirsys, etc). See
  // https://www.metered.ca/tools/openrelay/ for the free-tier signup that
  // gives you your own credentials.
  {
    urls: "turn:openrelay.metered.ca:80",
    username: "openrelayproject",
    credential: "openrelayproject",
  },
  {
    urls: "turn:openrelay.metered.ca:443",
    username: "openrelayproject",
    credential: "openrelayproject",
  },
  {
    urls: "turn:openrelay.metered.ca:443?transport=tcp",
    username: "openrelayproject",
    credential: "openrelayproject",
  },
];
const CALL_RING_TIMEOUT_MS = 45000; // auto-cancel an unanswered call

function useVoiceCall({
  role,
  currentAdmin,
  currentEmp,
  employees,
  admins,
  t,
  canUseMessages,
}) {
  const isAdmin = role === "admin";
  const selfId = isAdmin ? currentAdmin?.id || null : currentEmp?.id || null;
  const selfName = isAdmin
    ? currentAdmin?.name || t.chat.adminLabel
    : currentEmp?.name || "";
  // Read inside the signal handler/startCall closures below without
  // forcing those effects/callbacks to be rebuilt on every permission
  // recompute — same pattern as callRef.
  const canUseMessagesRef = useRef(canUseMessages);
  canUseMessagesRef.current = canUseMessages;

  // call: null | { status: 'outgoing'|'incoming'|'connected', employeeId,
  //                peerName, muted, startedAt }
  const [call, setCall] = useState(null);
  const [callError, setCallError] = useState("");

  const callRef = useRef(null);
  callRef.current = call;
  const pcRef = useRef(null);
  const localStreamRef = useRef(null);
  const remoteAudioRef = useRef(null);
  const pendingCandidatesRef = useRef([]);
  const ringTimeoutRef = useRef(null);
  const offerRetryRef = useRef(null);
  const stopRingtoneRef = useRef(null);

  useEffect(() => {
    // Hidden <audio> element the remote peer's stream plays through.
    // Created once here (not in JSX) so playback keeps going regardless
    // of which page/component is currently mounted.
    const audioEl = document.createElement("audio");
    audioEl.autoplay = true;
    remoteAudioRef.current = audioEl;
    return () => {
      remoteAudioRef.current = null;
    };
  }, []);

  const clearRingTimeout = () => {
    if (ringTimeoutRef.current) {
      clearTimeout(ringTimeoutRef.current);
      ringTimeoutRef.current = null;
    }
  };
  const clearOfferRetry = () => {
    if (offerRetryRef.current) {
      clearInterval(offerRetryRef.current);
      offerRetryRef.current = null;
    }
  };
  const stopRingtone = () => {
    if (stopRingtoneRef.current) {
      stopRingtoneRef.current();
      stopRingtoneRef.current = null;
    }
  };
  const teardownPeer = () => {
    stopRingtone();
    clearRingTimeout();
    clearOfferRetry();
    pendingCandidatesRef.current = [];
    if (pcRef.current) {
      pcRef.current.onicecandidate = null;
      pcRef.current.ontrack = null;
      pcRef.current.onconnectionstatechange = null;
      try {
        pcRef.current.close();
      } catch {
        // already closed
      }
      pcRef.current = null;
    }
    if (localStreamRef.current) {
      localStreamRef.current.getTracks().forEach((tr) => tr.stop());
      localStreamRef.current = null;
    }
    if (remoteAudioRef.current) remoteAudioRef.current.srcObject = null;
  };

  const endCall = useCallback(
    (opts = {}) => {
      const current = callRef.current;
      if (current && !opts.silent) {
        sendCallSignal({
          kind: "hangup",
          employeeId: current.employeeId,
          fromRole: isAdmin ? "admin" : "employee",
          fromId: selfId,
        });
      }
      teardownPeer();
      setCall(null);
    },
    [isAdmin, selfId],
  );

  const rejectCall = useCallback(() => {
    const incoming = callRef.current;
    if (!incoming) return;
    sendCallSignal({
      kind: "reject",
      employeeId: incoming.employeeId,
      fromRole: isAdmin ? "admin" : "employee",
      fromId: selfId,
    });
    teardownPeer();
    setCall(null);
  }, [isAdmin, selfId]);

  const createPeerConnection = (employeeId) => {
    const pc = new RTCPeerConnection({ iceServers: CALL_ICE_SERVERS });
    pc.onicecandidate = (e) => {
      if (e.candidate) {
        sendCallSignal({
          kind: "ice",
          employeeId,
          fromRole: isAdmin ? "admin" : "employee",
          fromId: selfId,
          candidate: e.candidate.toJSON(),
        });
      }
    };
    pc.ontrack = (e) => {
      if (remoteAudioRef.current)
        remoteAudioRef.current.srcObject = e.streams[0];
    };
    pc.onconnectionstatechange = () => {
      // "disconnected" can be a brief network blip and often recovers on
      // its own — only treat a genuine "failed" state as the call ending.
      if (pc.connectionState === "failed") {
        setCallError(t.chat.callConnectionLost);
        endCall({ silent: true });
      }
    };
    pcRef.current = pc;
    return pc;
  };

  const startLocalStream = async () => {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
    localStreamRef.current = stream;
    return stream;
  };

  const startCall = useCallback(
    async (employeeId, peerName, peerPhoto) => {
      if (callRef.current || !employeeId || !canUseMessagesRef.current) return;
      setCallError("");
      setCall({
        status: "outgoing",
        employeeId,
        peerName,
        // Admin calling a specific employee: real photo, identified.
        // Employee calling the shared admin mailbox: no specific admin
        // is chosen yet (whichever admin picks up), so this stays
        // generic — matches the incoming side, where that same call
        // only becomes "identified" once a specific admin's offer
        // arrives with their own id/name/photo.
        peerPhoto: isAdmin ? peerPhoto || null : null,
        peerKind: isAdmin ? "identified" : "generic",
        muted: false,
        startedAt: null,
      });
      try {
        const stream = await startLocalStream();
        const pc = createPeerConnection(employeeId);
        stream.getTracks().forEach((tr) => pc.addTrack(tr, stream));
        const offer = await pc.createOffer();
        await pc.setLocalDescription(offer);
        const broadcastOffer = () =>
          sendCallSignal({
            kind: "offer",
            employeeId,
            fromRole: isAdmin ? "admin" : "employee",
            fromId: selfId,
            fromName: selfName,
            sdp: offer,
          });
        broadcastOffer();
        // A single broadcast only reaches a receiver whose tab is already
        // open, foregrounded, and subscribed at that exact instant —
        // mobile browsers commonly suspend a backgrounded tab's websocket,
        // so the very first offer is easy to miss entirely (nothing
        // buffers/replays a Realtime broadcast). Re-send the identical
        // offer every few seconds for as long as we're still ringing, so
        // a receiver whose connection comes back mid-ring (e.g. they
        // switch back to this app, or the socket auto-reconnects) still
        // catches a later copy instead of the call just silently not
        // arriving. The receiving side (see the "offer" branch in the
        // signal handler below) treats a repeat offer for a call it's
        // already ringing on as a no-op refresh, not a new/duplicate call.
        offerRetryRef.current = setInterval(() => {
          if (callRef.current?.status === "outgoing") broadcastOffer();
        }, 3000);
        // Realtime broadcast (sendCallSignal above) only reaches a peer
        // whose tab is already open and subscribed — it does nothing if
        // the receiver's app is closed/backgrounded. Fire a real Web
        // Push too, same pattern as the `notify` callbacks on
        // useSupabaseArray (messages, leave, OT, ...), so an incoming
        // call actually rings the receiver's device even when the app
        // isn't in the foreground. Best-effort: a push failure must
        // never block or fail the call itself.
        supabase.functions
          .invoke("push_notify", {
            body: isAdmin
              ? {
                  userType: "employee",
                  userId: employeeId,
                  title: t.chat.incomingCall,
                  body: selfName || "",
                  page: "messages",
                  portal: "employee",
                  tag: `call-${employeeId}`,
                }
              : {
                  userType: "admin",
                  title: t.chat.incomingCall,
                  body: `${selfName || "?"} (${employeeId})`,
                  page: "messages",
                  portal: "admin",
                  tag: `call-${employeeId}`,
                },
          })
          .catch((err) => console.error("[push] call notify failed:", err));
        stopRingtoneRef.current = startRingtoneLoop();
        ringTimeoutRef.current = setTimeout(() => {
          setCallError(t.chat.callNoAnswer);
          endCall();
        }, CALL_RING_TIMEOUT_MS);
      } catch (err) {
        setCallError(
          err?.name === "NotAllowedError"
            ? t.chat.micDenied
            : t.chat.callFailedToStart,
        );
        teardownPeer();
        setCall(null);
      }
      // eslint-disable-next-line react-hooks/exhaustive-deps
    },
    [isAdmin, selfId, selfName, endCall, t],
  );

  const acceptCall = useCallback(async () => {
    const incoming = callRef.current;
    if (!incoming || incoming.status !== "incoming") return;
    stopRingtone();
    clearRingTimeout();
    setCallError("");
    try {
      const stream = await startLocalStream();
      const pc = pcRef.current;
      stream.getTracks().forEach((tr) => pc.addTrack(tr, stream));
      const answer = await pc.createAnswer();
      await pc.setLocalDescription(answer);
      sendCallSignal({
        kind: "answer",
        employeeId: incoming.employeeId,
        fromRole: isAdmin ? "admin" : "employee",
        fromId: selfId,
        sdp: answer,
      });
      if (isAdmin) {
        // Tell any other admin who might also be ringing that this call
        // has already been taken, so they stop ringing too.
        sendCallSignal({
          kind: "answered-elsewhere",
          employeeId: incoming.employeeId,
          fromRole: "admin",
          fromId: selfId,
        });
      }
      setCall((c) =>
        c ? { ...c, status: "connected", startedAt: Date.now() } : c,
      );
    } catch (err) {
      setCallError(
        err?.name === "NotAllowedError"
          ? t.chat.micDenied
          : t.chat.callFailedToStart,
      );
      sendCallSignal({
        kind: "reject",
        employeeId: incoming.employeeId,
        fromRole: isAdmin ? "admin" : "employee",
        fromId: selfId,
      });
      teardownPeer();
      setCall(null);
    }
  }, [isAdmin, selfId, t]);

  const toggleMute = useCallback(() => {
    setCall((c) => {
      if (!c || !localStreamRef.current) return c;
      const next = !c.muted;
      localStreamRef.current
        .getAudioTracks()
        .forEach((tr) => (tr.enabled = !next));
      return { ...c, muted: next };
    });
  }, []);

  // Incoming signaling — one handler processes every event on the shared
  // channel; each branch below filters down to what's relevant.
  useEffect(() => {
    if (!selfId) return undefined;
    const handler = async (msg) => {
      if (!msg) return;
      // Employees only ever care about their own thread; admins can see
      // signaling for any employee (shared-mailbox model, same as chat).
      if (!isAdmin && msg.employeeId !== selfId) return;

      if (msg.kind === "offer") {
        // Permission was revoked (or never granted) for this session —
        // decline silently rather than ringing a screen the user isn't
        // allowed to use. Uses "busy" (not "reject") so the caller sees
        // a generic can't-connect state instead of a misleading
        // "declined by the person" message.
        if (!canUseMessagesRef.current) {
          sendCallSignal({
            kind: "busy",
            employeeId: msg.employeeId,
            fromRole: isAdmin ? "admin" : "employee",
            fromId: selfId,
          });
          return;
        }
        // Already ringing/connected. If this is the *same* caller
        // re-sending the same offer (our retry loop above), just refresh
        // the ring timeout rather than rejecting our own retried call as
        // "busy" — that would cancel a call we're actively ringing for.
        // Anything else (a different caller, or already answered/on a
        // different call) really is busy.
        if (callRef.current) {
          if (
            callRef.current.status === "incoming" &&
            callRef.current.employeeId === msg.employeeId &&
            callRef.current.fromId === msg.fromId &&
            msg.fromRole !== (isAdmin ? "admin" : "employee")
          ) {
            clearRingTimeout();
            ringTimeoutRef.current = setTimeout(
              () => rejectCall(),
              CALL_RING_TIMEOUT_MS,
            );
            return;
          }
          sendCallSignal({
            kind: "busy",
            employeeId: msg.employeeId,
            fromRole: isAdmin ? "admin" : "employee",
            fromId: selfId,
          });
          return;
        }
        // An offer from "our own side" (another admin calling that same
        // employee) isn't something we should ring for.
        if (msg.fromRole === (isAdmin ? "admin" : "employee")) return;
        // Every signaling message already carries the real caller's id
        // (msg.fromId) and a name snapshot (msg.fromName) — look up the
        // live admin record too, in case their name/photo changed since
        // the call started, but fall back to the snapshot if that admin
        // account can't be found (e.g. deleted mid-call). Only the
        // reverse direction (employee calling the *shared* admin
        // mailbox, with no specific admin chosen yet) stays generic —
        // see startCall below.
        const callingAdmin = !isAdmin
          ? admins.find((a) => a.id === msg.fromId)
          : null;
        const peerName = isAdmin
          ? msg.fromName ||
            employees.find((e) => e.id === msg.employeeId)?.name ||
            "?"
          : callingAdmin?.name || msg.fromName || t.chat.adminLabel;
        const peerPhoto = isAdmin
          ? employees.find((e) => e.id === msg.employeeId)?.photo || null
          : callingAdmin?.photo || null;
        try {
          const pc = createPeerConnection(msg.employeeId);
          await pc.setRemoteDescription(new RTCSessionDescription(msg.sdp));
          setCall({
            status: "incoming",
            employeeId: msg.employeeId,
            fromId: msg.fromId,
            peerName,
            peerPhoto,
            peerKind: "identified",
            muted: false,
            startedAt: null,
          });
          stopRingtoneRef.current = startRingtoneLoop();
          ringTimeoutRef.current = setTimeout(
            () => rejectCall(),
            CALL_RING_TIMEOUT_MS,
          );
        } catch {
          teardownPeer();
        }
        return;
      }

      const current = callRef.current;
      if (!current || current.employeeId !== msg.employeeId) return;

      if (msg.kind === "answer") {
        if (current.status !== "outgoing" || !pcRef.current) return;
        clearRingTimeout();
        clearOfferRetry();
        stopRingtone();
        try {
          await pcRef.current.setRemoteDescription(
            new RTCSessionDescription(msg.sdp),
          );
          for (const c of pendingCandidatesRef.current) {
            await pcRef.current.addIceCandidate(c).catch(() => {});
          }
          pendingCandidatesRef.current = [];
          setCall((c) =>
            c ? { ...c, status: "connected", startedAt: Date.now() } : c,
          );
        } catch {
          setCallError(t.chat.callFailedToStart);
          endCall({ silent: true });
        }
        return;
      }

      if (msg.kind === "ice") {
        const candidate = new RTCIceCandidate(msg.candidate);
        if (pcRef.current?.remoteDescription) {
          pcRef.current.addIceCandidate(candidate).catch(() => {});
        } else {
          pendingCandidatesRef.current.push(candidate);
        }
        return;
      }

      if (msg.kind === "reject" || msg.kind === "busy") {
        setCallError(
          msg.kind === "busy" ? t.chat.callBusy : t.chat.callDeclined,
        );
        teardownPeer();
        setCall(null);
        return;
      }

      if (msg.kind === "hangup") {
        teardownPeer();
        setCall(null);
        return;
      }

      if (msg.kind === "answered-elsewhere") {
        if (isAdmin && current.status === "incoming") {
          teardownPeer();
          setCall(null);
        }
        return;
      }

      if (msg.kind === "cancel") {
        if (current.status === "incoming") {
          setCallError(t.chat.callCancelled);
          teardownPeer();
          setCall(null);
        }
        return;
      }
    };
    return subscribeCallSignal(handler);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selfId, isAdmin, employees, admins]);

  // If the whole app unmounts mid-call (sign-out, tab close), let the
  // other side know instead of just going silent.
  useEffect(() => {
    return () => {
      const current = callRef.current;
      if (current) {
        sendCallSignal({
          kind: current.status === "incoming" ? "reject" : "cancel",
          employeeId: current.employeeId,
          fromRole: isAdmin ? "admin" : "employee",
          fromId: selfId,
        });
      }
      teardownPeer();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return {
    call,
    callError,
    clearCallError: () => setCallError(""),
    startCall,
    acceptCall,
    rejectCall,
    endCall,
    toggleMute,
  };
}

// Small floating call UI — rendered once near the app root so it stays
// visible (and the ringtone keeps playing) no matter which page the
// person is currently looking at. Three states: incoming (ring/accept/
// decline), outgoing (calling.../cancel), connected (timer/mute/hangup).
function CallOverlay({
  call,
  callError,
  onAccept,
  onReject,
  onEnd,
  onToggleMute,
  onDismissError,
}) {
  const { t } = useLang();
  const [elapsed, setElapsed] = useState(0);

  useEffect(() => {
    if (!call || call.status !== "connected" || !call.startedAt) return;
    const id = setInterval(() => {
      setElapsed(Math.floor((Date.now() - call.startedAt) / 1000));
    }, 1000);
    return () => clearInterval(id);
  }, [call]);

  useEffect(() => {
    if (!callError) return;
    const id = setTimeout(onDismissError, 4000);
    return () => clearTimeout(id);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [callError]);

  if (!call && !callError) return null;

  const fmtElapsed = `${String(Math.floor(elapsed / 60)).padStart(2, "0")}:${String(elapsed % 60).padStart(2, "0")}`;

  // A dropped/failed call with no active `call` state (no-answer timeout,
  // mic permission denied, etc.) only needs a small dismissible toast —
  // the full call screen below is only for a call that's actually live.
  if (!call) {
    return (
      <div
        style={{
          position: "fixed",
          top: "max(16px, env(safe-area-inset-top))",
          left: "50%",
          transform: "translateX(-50%)",
          zIndex: 9999,
          maxWidth: "calc(100vw - 32px)",
          background: "#151b2b",
          color: "#fff",
          borderRadius: 12,
          boxShadow: "0 12px 32px rgba(0,0,0,0.35)",
          padding: "12px 14px",
          display: "flex",
          alignItems: "center",
          gap: 10,
          fontSize: 13,
        }}
      >
        <AlertCircle size={16} style={{ flexShrink: 0, color: "#ff8a80" }} />
        <span>{callError}</span>
        <button
          type="button"
          onClick={onDismissError}
          style={{
            border: "none",
            background: "transparent",
            color: "rgba(255,255,255,0.6)",
            cursor: "pointer",
            display: "flex",
            flexShrink: 0,
            marginLeft: 2,
          }}
        >
          <X size={14} />
        </button>
      </div>
    );
  }

  const isIncoming = call.status === "incoming";
  const isConnected = call.status === "connected";
  const statusLabel = isIncoming
    ? t.chat.incomingCall
    : isConnected
      ? fmtElapsed
      : t.chat.calling;

  return (
    <div className="wf-callscreen">
      <div
        style={{
          fontSize: 12.5,
          fontWeight: 700,
          color: "rgba(255,255,255,0.5)",
          letterSpacing: 0.5,
          textTransform: "uppercase",
        }}
      >
        {t.chat.call}
      </div>

      <div
        style={{
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          gap: 20,
        }}
      >
        <div
          style={{
            position: "relative",
            width: 136,
            height: 136,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          {!isConnected && (
            <>
              <span
                className="wf-callscreen-ring"
                style={{ width: 136, height: 136 }}
              />
              <span
                className="wf-callscreen-ring"
                style={{ width: 136, height: 136, animationDelay: "0.7s" }}
              />
            </>
          )}
          {call.peerKind === "generic" ? (
            // Employee calling out to the shared admin mailbox: no
            // specific admin has been chosen yet, so a generic badge
            // is honest here — this flips to a real photo/name the
            // moment a specific admin's offer arrives (see peerKind
            // "identified" below).
            <div
              style={{
                width: 112,
                height: 112,
                borderRadius: "50%",
                background: "rgba(255,255,255,0.12)",
                border: "1px solid rgba(255,255,255,0.2)",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              <ShieldCheck size={48} color="#fff" />
            </div>
          ) : (
            <Avatar name={call.peerName} photo={call.peerPhoto} size={112} />
          )}
        </div>
        <div style={{ textAlign: "center" }}>
          <div style={{ fontSize: 23, fontWeight: 700 }}>{call.peerName}</div>
          <div
            style={{
              marginTop: 7,
              fontSize: 14.5,
              color: "rgba(255,255,255,0.65)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              gap: 3,
              fontVariantNumeric: "tabular-nums",
              minHeight: 18,
            }}
          >
            {statusLabel}
            {!isConnected && (
              <span
                className="wf-callscreen-dots"
                style={{ display: "inline-flex", gap: 2 }}
              >
                <span>.</span>
                <span>.</span>
                <span>.</span>
              </span>
            )}
          </div>
        </div>
      </div>

      <div style={{ display: "flex", alignItems: "flex-start", gap: 44 }}>
        {isIncoming ? (
          <>
            <button type="button" className="wf-callbtn" onClick={onReject}>
              <span className="wf-callbtn-circle wf-callbtn-reject">
                <PhoneOff size={26} />
              </span>
              <span className="wf-callbtn-label">{t.chat.decline}</span>
            </button>
            <button type="button" className="wf-callbtn" onClick={onAccept}>
              <span className="wf-callbtn-circle wf-callbtn-accept">
                <Phone size={26} />
              </span>
              <span className="wf-callbtn-label">{t.chat.accept}</span>
            </button>
          </>
        ) : (
          <>
            {isConnected && (
              <button
                type="button"
                className="wf-callbtn"
                onClick={onToggleMute}
              >
                <span
                  className={`wf-callbtn-circle wf-callbtn-secondary${
                    call.muted ? " wf-callbtn-active" : ""
                  }`}
                >
                  {call.muted ? <MicOff size={22} /> : <Mic size={22} />}
                </span>
                <span className="wf-callbtn-label">
                  {call.muted ? t.chat.unmute : t.chat.mute}
                </span>
              </button>
            )}
            <button type="button" className="wf-callbtn" onClick={onEnd}>
              <span className="wf-callbtn-circle wf-callbtn-reject">
                <PhoneOff size={26} />
              </span>
              <span className="wf-callbtn-label">{t.chat.hangup}</span>
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function MessagesPage({
  role,
  currentAdmin,
  currentEmp,
  employees,
  messages,
  setMessages,
  activeCall,
  onStartCall,
}) {
  const { t, lang } = useLang();
  const isAdmin = role === "admin";
  const [selectedEmpId, setSelectedEmpId] = useState(
    isAdmin ? null : currentEmp?.id || null,
  );
  const [query, setQuery] = useState("");
  const [draft, setDraft] = useState("");
  const scrollRef = useRef(null);
  const fileInputRef = useRef(null);
  const [pendingAttachment, setPendingAttachment] = useState(null); // { dataUrl, name, type, size }
  const [attaching, setAttaching] = useState(false);
  const [attachError, setAttachError] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editDraft, setEditDraft] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [confirmDeleteThread, setConfirmDeleteThread] = useState(false);
  const [recording, setRecording] = useState(false);
  const [recordSeconds, setRecordSeconds] = useState(0);
  const [recordError, setRecordError] = useState("");
  const mediaRecorderRef = useRef(null);
  const recordChunksRef = useRef([]);
  const recordStreamRef = useRef(null);
  const recordTimerRef = useRef(null);

  const threadEmpId = isAdmin ? selectedEmpId : currentEmp?.id || null;

  const threadMessages = useMemo(
    () =>
      messages
        .filter((m) => m.employeeId === threadEmpId)
        .sort((a, b) => (a.createdAt || "").localeCompare(b.createdAt || "")),
    [messages, threadEmpId],
  );

  // Id of the viewer's own most recent (non-deleted) message in this
  // thread — only that bubble shows a Seen/Delivered receipt, matching
  // the usual chat-app convention instead of stamping every message.
  const lastMineMessageId = useMemo(() => {
    const mineRole = isAdmin ? "admin" : "employee";
    for (let i = threadMessages.length - 1; i >= 0; i--) {
      const m = threadMessages[i];
      if (m.senderRole === mineRole && !m.deleted) return m.id;
    }
    return null;
  }, [threadMessages, isAdmin]);

  // Flags the first message of each calendar day so the thread can show
  // a "Today / Yesterday / <date>" separator above it, chat-app style.
  // threadMessages is already sorted ascending by createdAt, so a single
  // pass comparing each message's date to the previous one is enough.
  const threadMessagesWithDaySeparators = useMemo(() => {
    let lastDateKey = null;
    return threadMessages.map((m) => {
      const dateKey = m.createdAt ? new Date(m.createdAt).toDateString() : null;
      const showDaySeparator = dateKey !== lastDateKey;
      lastDateKey = dateKey;
      return { message: m, showDaySeparator };
    });
  }, [threadMessages]);

  // One row per employee who has ever exchanged a message, each carrying
  // its most recent message and how many are unread — newest activity
  // first, same "pending things surface first" idea used elsewhere in
  // this app (e.g. AttendanceCorrections).
  const conversations = useMemo(() => {
    if (!isAdmin) return [];
    const byEmp = new Map();
    messages.forEach((m) => {
      const cur = byEmp.get(m.employeeId);
      if (!cur || (m.createdAt || "") > (cur.lastMessage.createdAt || "")) {
        byEmp.set(m.employeeId, { lastMessage: m, unread: 0 });
      }
    });
    messages.forEach((m) => {
      if (m.senderRole === "employee" && !m.readByAdmin) {
        const cur = byEmp.get(m.employeeId);
        if (cur) cur.unread += 1;
      }
    });
    return Array.from(byEmp.entries())
      .map(([employeeId, info]) => ({
        employeeId,
        employee: employees.find((e) => e.id === employeeId),
        ...info,
      }))
      .filter((c) => c.employee)
      .sort((a, b) =>
        (b.lastMessage.createdAt || "").localeCompare(
          a.lastMessage.createdAt || "",
        ),
      );
  }, [isAdmin, messages, employees]);

  // Every active employee, shown as a selectable row whether or not a
  // conversation exists yet — employees with an existing thread keep their
  // last-message preview and sort by recency; everyone else is listed
  // alphabetically underneath so admin can pick anyone and message first.
  const allChatItems = useMemo(() => {
    if (!isAdmin) return [];
    const q = query.trim().toLowerCase();
    const convByEmp = new Map(conversations.map((c) => [c.employeeId, c]));
    const items = employees
      .filter((e) => e.status === "active")
      .filter(
        (e) =>
          !q ||
          e.name?.toLowerCase().includes(q) ||
          e.code?.toLowerCase().includes(q),
      )
      .map((e) => {
        const conv = convByEmp.get(e.id);
        return conv
          ? { employeeId: e.id, employee: e, ...conv, hasConversation: true }
          : {
              employeeId: e.id,
              employee: e,
              lastMessage: null,
              unread: 0,
              hasConversation: false,
            };
      });
    return items.sort((a, b) => {
      if (a.hasConversation && b.hasConversation) {
        return (b.lastMessage.createdAt || "").localeCompare(
          a.lastMessage.createdAt || "",
        );
      }
      if (a.hasConversation !== b.hasConversation) {
        return a.hasConversation ? -1 : 1;
      }
      return (a.employee.name || "").localeCompare(b.employee.name || "");
    });
  }, [isAdmin, employees, conversations, query]);

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight });
  }, [threadMessages.length, threadEmpId]);

  // Marks the open thread's incoming messages read. Guarded on there
  // actually being unread rows before calling setMessages, so this
  // doesn't turn into an infinite loop re-firing on its own writes.
  useEffect(() => {
    if (!threadEmpId) return;
    const unreadIncoming = messages.filter(
      (m) =>
        m.employeeId === threadEmpId &&
        (isAdmin
          ? m.senderRole === "employee" && !m.readByAdmin
          : m.senderRole === "admin" && !m.readByEmployee),
    );
    if (unreadIncoming.length === 0) return;
    const ids = new Set(unreadIncoming.map((m) => m.id));
    setMessages(
      messages.map((m) =>
        ids.has(m.id)
          ? isAdmin
            ? { ...m, readByAdmin: true }
            : { ...m, readByEmployee: true }
          : m,
      ),
    );
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [threadEmpId, messages, isAdmin]);

  const send = () => {
    const content = draft.trim();
    if ((!content && !pendingAttachment) || !threadEmpId) return;
    setMessages([
      ...messages,
      {
        id: uid("msg"),
        employeeId: threadEmpId,
        senderRole: isAdmin ? "admin" : "employee",
        senderId: isAdmin ? currentAdmin?.id || null : currentEmp?.id || null,
        senderName: isAdmin
          ? currentAdmin?.name || t.chat.adminLabel
          : currentEmp?.name || "",
        content,
        attachmentData: pendingAttachment?.dataUrl || null,
        attachmentName: pendingAttachment?.name || null,
        attachmentType: pendingAttachment?.type || null,
        readByAdmin: isAdmin,
        readByEmployee: !isAdmin,
        createdAt: new Date().toISOString(),
      },
    ]);
    setDraft("");
    setPendingAttachment(null);
    setAttachError("");
  };

  const pickAttachment = () => fileInputRef.current?.click();

  const onAttachmentChange = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = ""; // allow re-picking the same file next time
    if (!file) return;
    if (file.size > MAX_CHAT_FILE_BYTES) {
      setAttachError(t.chat.fileTooLarge);
      return;
    }
    setAttachError("");
    setAttaching(true);
    try {
      const isImage = file.type.startsWith("image/");
      const dataUrl = isImage
        ? await fileToCompressedChatImageDataUrl(file)
        : await fileToDataUrl(file);
      setPendingAttachment({
        dataUrl,
        name: file.name,
        type: file.type || "application/octet-stream",
        size: file.size,
      });
    } catch {
      setAttachError(t.chat.fileReadError);
    } finally {
      setAttaching(false);
    }
  };

  const threadEmployee = isAdmin
    ? employees.find((e) => e.id === threadEmpId)
    : currentEmp;

  // Voice messages reuse the exact same pendingAttachment shape as file/
  // image attachments (dataUrl/name/type/size), so send() and the bubble
  // renderer below don't need a separate code path — only a branch on
  // attachmentType starting with "audio/" to show a player instead of a
  // generic download link or image.
  const stopRecordingTracks = () => {
    if (recordTimerRef.current) {
      clearInterval(recordTimerRef.current);
      recordTimerRef.current = null;
    }
    recordStreamRef.current?.getTracks().forEach((tr) => tr.stop());
    recordStreamRef.current = null;
  };

  const startRecording = async () => {
    setRecordError("");
    if (!navigator.mediaDevices?.getUserMedia || !window.MediaRecorder) {
      setRecordError(t.chat.micUnsupported);
      return;
    }
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        audio: true,
      });
      recordStreamRef.current = stream;
      const mimeType = MediaRecorder.isTypeSupported("audio/webm")
        ? "audio/webm"
        : MediaRecorder.isTypeSupported("audio/mp4")
          ? "audio/mp4"
          : "";
      const recorder = mimeType
        ? new MediaRecorder(stream, { mimeType })
        : new MediaRecorder(stream);
      recordChunksRef.current = [];
      recorder.ondataavailable = (e) => {
        if (e.data && e.data.size > 0) recordChunksRef.current.push(e.data);
      };
      recorder.onstop = async () => {
        stopRecordingTracks();
        const blob = new Blob(recordChunksRef.current, {
          type: recorder.mimeType || mimeType || "audio/webm",
        });
        recordChunksRef.current = [];
        if (blob.size === 0) return; // tapped stop almost instantly
        if (blob.size > MAX_CHAT_FILE_BYTES) {
          setAttachError(t.chat.fileTooLarge);
          return;
        }
        try {
          const dataUrl = await fileToDataUrl(blob);
          const ext = (blob.type || "").includes("mp4") ? "m4a" : "webm";
          setPendingAttachment({
            dataUrl,
            name: `voice-${Date.now()}.${ext}`,
            type: blob.type || "audio/webm",
            size: blob.size,
          });
        } catch {
          setAttachError(t.chat.fileReadError);
        }
      };
      mediaRecorderRef.current = recorder;
      recorder.start();
      setRecording(true);
      setRecordSeconds(0);
      recordTimerRef.current = setInterval(
        () => setRecordSeconds((s) => s + 1),
        1000,
      );
    } catch {
      setRecordError(t.chat.micDenied);
    }
  };

  const stopRecording = () => {
    setRecording(false);
    mediaRecorderRef.current?.stop();
    mediaRecorderRef.current = null;
  };

  const cancelRecording = () => {
    setRecording(false);
    if (mediaRecorderRef.current) {
      mediaRecorderRef.current.onstop = null;
      mediaRecorderRef.current.stop();
      mediaRecorderRef.current = null;
    }
    recordChunksRef.current = [];
    stopRecordingTracks();
  };

  useEffect(() => {
    // Stop any in-progress recording if the admin/employee navigates away
    // from Messages entirely, so the mic doesn't stay hot in the background.
    return () => {
      if (mediaRecorderRef.current) {
        mediaRecorderRef.current.onstop = null;
        try {
          mediaRecorderRef.current.stop();
        } catch {
          // already stopped
        }
      }
      stopRecordingTracks();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Edit/delete is sender-only: a message is "mine to modify" only if it
  // was actually sent by the currently logged-in identity, not just by
  // "an admin" — admins share one mailbox per employee, so a message
  // showing on the right (role match) may still belong to a different
  // admin.
  const canModify = (m) =>
    isAdmin
      ? m.senderRole === "admin" && m.senderId === (currentAdmin?.id || null)
      : m.senderRole === "employee" && m.senderId === (currentEmp?.id || null);

  const startEdit = (m) => {
    setEditingId(m.id);
    setEditDraft(m.content || "");
  };

  const cancelEdit = () => {
    setEditingId(null);
    setEditDraft("");
  };

  const saveEdit = () => {
    const msg = messages.find((m) => m.id === editingId);
    if (!msg) return;
    const content = editDraft.trim();
    if (!content && !msg.attachmentData) return; // nothing left to show
    setMessages(
      messages.map((m) =>
        m.id === editingId
          ? { ...m, content, editedAt: new Date().toISOString() }
          : m,
      ),
    );
    cancelEdit();
  };

  const deleteMessage = (id) => {
    setMessages(
      messages.map((m) =>
        m.id === id
          ? {
              ...m,
              content: "",
              attachmentData: null,
              attachmentName: null,
              attachmentType: null,
              deleted: true,
            }
          : m,
      ),
    );
    setConfirmDeleteId(null);
    if (editingId === id) cancelEdit();
  };

  // Superadmin-only: wipes the entire thread with one employee, not just
  // one message. Unlike deleteMessage above (soft delete — content is
  // cleared but the row stays as a "This message was deleted" tombstone),
  // this hard-removes every row for threadEmpId, since the point is to
  // clear the conversation out of the mailbox entirely.
  const deleteConversation = () => {
    if (!threadEmpId) return;
    setMessages(messages.filter((m) => m.employeeId !== threadEmpId));
    setConfirmDeleteThread(false);
    setSelectedEmpId(null);
  };

  const threadPane = (
    <>
      <div className="wf-chat-thread-pane">
        {threadEmpId ? (
          <>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                padding: "12px 16px",
                borderBottom: `1px solid ${T.lineSoft}`,
                background: T.card,
              }}
            >
              {isAdmin && (
                <button
                  className="wf-chat-back-btn"
                  onClick={() => setSelectedEmpId(null)}
                >
                  <ArrowLeft size={18} />
                </button>
              )}
              <Avatar
                name={isAdmin ? threadEmployee?.name : t.chat.adminLabel}
                photo={isAdmin ? threadEmployee?.photo : null}
                size={32}
              />
              <div style={{ minWidth: 0, flex: 1 }}>
                <div
                  style={{
                    fontWeight: 700,
                    fontSize: 13.5,
                    color: T.ink,
                    overflow: "hidden",
                    textOverflow: "ellipsis",
                    whiteSpace: "nowrap",
                  }}
                >
                  {isAdmin ? threadEmployee?.name || "?" : t.chat.adminLabel}
                </div>
                {isAdmin && (
                  <div style={{ fontSize: 11, color: T.muted }}>
                    {threadEmployee?.code}
                  </div>
                )}
              </div>
              {onStartCall && (
                <button
                  type="button"
                  onClick={() =>
                    onStartCall(
                      threadEmpId,
                      isAdmin ? threadEmployee?.name || "?" : t.chat.adminLabel,
                      isAdmin ? threadEmployee?.photo : null,
                    )
                  }
                  disabled={!!activeCall}
                  title={isAdmin ? t.chat.callEmployee : t.chat.callAdmin}
                  style={{
                    border: "none",
                    background: "transparent",
                    color: activeCall ? T.mutedLight : T.forest,
                    cursor: activeCall ? "default" : "pointer",
                    opacity: activeCall ? 0.5 : 1,
                    display: "flex",
                    padding: 6,
                    flexShrink: 0,
                  }}
                >
                  <Phone size={17} />
                </button>
              )}
              {isAdmin && currentAdmin?.role === "superadmin" && (
                <button
                  type="button"
                  onClick={() => setConfirmDeleteThread(true)}
                  title={t.chat.deleteConversation}
                  style={{
                    border: "none",
                    background: "transparent",
                    color: T.mutedLight,
                    cursor: "pointer",
                    display: "flex",
                    padding: 6,
                    flexShrink: 0,
                  }}
                >
                  <Trash2 size={16} />
                </button>
              )}
            </div>
            <div
              ref={scrollRef}
              style={{
                flex: 1,
                overflowY: "auto",
                padding: "12px 0",
                display: "flex",
                flexDirection: "column",
                gap: 6,
              }}
            >
              {threadMessages.length === 0 ? (
                <div
                  style={{
                    margin: "auto",
                    color: T.muted,
                    fontSize: 12.5,
                    textAlign: "center",
                    padding: "0 24px",
                  }}
                >
                  {t.chat.noMessages}
                </div>
              ) : (
                threadMessagesWithDaySeparators.map(
                  ({ message: m, showDaySeparator }) => {
                    const mine = isAdmin
                      ? m.senderRole === "admin"
                      : m.senderRole === "employee";
                    const isEditing = editingId === m.id;
                    const modifiable = canModify(m) && !m.deleted;
                    return (
                      <React.Fragment key={m.id}>
                        {showDaySeparator && (
                          <div
                            style={{
                              display: "flex",
                              alignItems: "center",
                              justifyContent: "center",
                              margin: "10px 0 4px",
                            }}
                          >
                            <span
                              style={{
                                fontSize: 10.5,
                                fontWeight: 600,
                                color: T.muted,
                                background: T.inputBg,
                                border: `1px solid ${T.lineSoft}`,
                                borderRadius: 999,
                                padding: "3px 12px",
                              }}
                            >
                              {chatDateSeparatorLabel(m.createdAt, lang, t)}
                            </span>
                          </div>
                        )}
                        <div
                          className="wf-chat-bubble-row"
                          style={{ textAlign: mine ? "right" : "left" }}
                        >
                          <div
                            style={{
                              display: "inline-block",
                              maxWidth: "74%",
                              textAlign: "left",
                            }}
                          >
                            {isAdmin &&
                              m.senderRole === "admin" &&
                              m.senderName && (
                                <div
                                  style={{
                                    fontSize: 10,
                                    color: T.muted,
                                    marginBottom: 2,
                                    textAlign: "right",
                                  }}
                                >
                                  {m.senderName}
                                </div>
                              )}
                            {m.deleted ? (
                              <div
                                className="wf-chat-bubble"
                                style={{
                                  background: "transparent",
                                  color: T.muted,
                                  border: `1px dashed ${T.lineSoft}`,
                                  fontStyle: "italic",
                                  marginLeft: mine ? "auto" : 0,
                                }}
                              >
                                {t.chat.messageDeleted}
                              </div>
                            ) : (
                              <>
                                {m.attachmentData &&
                                  (m.attachmentType || "").startsWith(
                                    "image/",
                                  ) && (
                                    <a
                                      href={m.attachmentData}
                                      target="_blank"
                                      rel="noreferrer"
                                      style={{
                                        display: "block",
                                        marginLeft: mine ? "auto" : 0,
                                        marginBottom: m.content ? 6 : 0,
                                        width: "fit-content",
                                      }}
                                    >
                                      <img
                                        src={m.attachmentData}
                                        alt={
                                          m.attachmentName ||
                                          t.chat.photoAttachment
                                        }
                                        style={{
                                          maxWidth: 220,
                                          maxHeight: 260,
                                          borderRadius: 14,
                                          display: "block",
                                          cursor: "zoom-in",
                                          objectFit: "cover",
                                        }}
                                      />
                                    </a>
                                  )}
                                {m.attachmentData &&
                                  (m.attachmentType || "").startsWith(
                                    "audio/",
                                  ) && (
                                    <audio
                                      controls
                                      src={m.attachmentData}
                                      onLoadedMetadata={fixAudioBlobDuration}
                                      style={{
                                        display: "block",
                                        marginLeft: mine ? "auto" : 0,
                                        marginBottom: m.content ? 6 : 0,
                                        height: 36,
                                        maxWidth: 240,
                                      }}
                                    />
                                  )}
                                {m.attachmentData &&
                                  !(m.attachmentType || "").startsWith(
                                    "image/",
                                  ) &&
                                  !(m.attachmentType || "").startsWith(
                                    "audio/",
                                  ) && (
                                    <a
                                      href={m.attachmentData}
                                      download={m.attachmentName || "file"}
                                      style={{
                                        display: "flex",
                                        alignItems: "center",
                                        gap: 8,
                                        padding: "9px 13px",
                                        borderRadius: 14,
                                        background: mine ? T.forest : T.card,
                                        color: mine ? "#fff" : T.ink,
                                        border: mine
                                          ? "none"
                                          : `1px solid ${T.lineSoft}`,
                                        marginLeft: mine ? "auto" : 0,
                                        marginBottom: m.content ? 6 : 0,
                                        textDecoration: "none",
                                        maxWidth: "100%",
                                        width: "fit-content",
                                      }}
                                    >
                                      <FileText size={16} />
                                      <div
                                        style={{
                                          minWidth: 0,
                                          fontSize: 12.5,
                                          fontWeight: 600,
                                          overflow: "hidden",
                                          textOverflow: "ellipsis",
                                          whiteSpace: "nowrap",
                                        }}
                                      >
                                        {m.attachmentName || t.chat.download}
                                      </div>
                                    </a>
                                  )}
                                {isEditing ? (
                                  <div
                                    style={{
                                      display: "flex",
                                      flexDirection: "column",
                                      gap: 6,
                                      minWidth: 200,
                                    }}
                                  >
                                    <Input
                                      value={editDraft}
                                      onChange={(e) =>
                                        setEditDraft(e.target.value)
                                      }
                                      onKeyDown={(e) => {
                                        if (e.key === "Enter" && !e.shiftKey) {
                                          e.preventDefault();
                                          saveEdit();
                                        } else if (e.key === "Escape") {
                                          cancelEdit();
                                        }
                                      }}
                                      autoFocus
                                    />
                                    <div
                                      style={{
                                        display: "flex",
                                        justifyContent: "flex-end",
                                        gap: 6,
                                      }}
                                    >
                                      <Button
                                        variant="ghost"
                                        onClick={cancelEdit}
                                      >
                                        {t.chat.cancelEdit}
                                      </Button>
                                      <Button
                                        variant="accent"
                                        onClick={saveEdit}
                                        disabled={
                                          !editDraft.trim() && !m.attachmentData
                                        }
                                      >
                                        {t.chat.saveEdit}
                                      </Button>
                                    </div>
                                  </div>
                                ) : (
                                  m.content && (
                                    <div
                                      className="wf-chat-bubble"
                                      style={{
                                        background: mine ? T.forest : T.card,
                                        color: mine ? "#fff" : T.ink,
                                        border: mine
                                          ? "none"
                                          : `1px solid ${T.lineSoft}`,
                                        marginLeft: mine ? "auto" : 0,
                                      }}
                                    >
                                      {m.content}
                                    </div>
                                  )
                                )}
                              </>
                            )}
                            <div
                              style={{
                                display: "flex",
                                alignItems: "center",
                                gap: 6,
                                justifyContent: mine
                                  ? "flex-end"
                                  : "flex-start",
                                marginTop: 2,
                              }}
                            >
                              {mine && modifiable && !isEditing && (
                                <>
                                  {m.content && (
                                    <button
                                      type="button"
                                      onClick={() => startEdit(m)}
                                      title={t.chat.editMessage}
                                      style={{
                                        border: "none",
                                        background: "transparent",
                                        color: T.mutedLight,
                                        cursor: "pointer",
                                        display: "flex",
                                        padding: 0,
                                      }}
                                    >
                                      <Pencil size={11} />
                                    </button>
                                  )}
                                  <button
                                    type="button"
                                    onClick={() => setConfirmDeleteId(m.id)}
                                    title={t.chat.deleteMessage}
                                    style={{
                                      border: "none",
                                      background: "transparent",
                                      color: T.mutedLight,
                                      cursor: "pointer",
                                      display: "flex",
                                      padding: 0,
                                    }}
                                  >
                                    <Trash2 size={11} />
                                  </button>
                                </>
                              )}
                              <div
                                style={{ fontSize: 10, color: T.mutedLight }}
                              >
                                {m.editedAt && !m.deleted
                                  ? `${timeAgoLabel(m.createdAt)} · ${t.chat.edited}`
                                  : timeAgoLabel(m.createdAt)}
                                {mine && m.id === lastMineMessageId && (
                                  <>
                                    {" · "}
                                    {(
                                      isAdmin ? m.readByEmployee : m.readByAdmin
                                    )
                                      ? t.chat.seen
                                      : t.chat.delivered}
                                  </>
                                )}
                              </div>
                            </div>
                          </div>
                        </div>
                      </React.Fragment>
                    );
                  },
                )
              )}
            </div>
            <div
              style={{
                borderTop: `1px solid ${T.lineSoft}`,
                background: T.card,
              }}
            >
              {(pendingAttachment || attaching) && (
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 8,
                    padding: "8px 12px 0 12px",
                  }}
                >
                  {attaching ? (
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                        fontSize: 12,
                        color: T.muted,
                      }}
                    >
                      <Loader2
                        size={14}
                        style={{ animation: "spin 1s linear infinite" }}
                      />
                    </div>
                  ) : (
                    <>
                      {(pendingAttachment.type || "").startsWith("image/") ? (
                        <img
                          src={pendingAttachment.dataUrl}
                          alt={pendingAttachment.name}
                          style={{
                            width: 44,
                            height: 44,
                            borderRadius: 8,
                            objectFit: "cover",
                            border: `1px solid ${T.lineSoft}`,
                          }}
                        />
                      ) : (pendingAttachment.type || "").startsWith(
                          "audio/",
                        ) ? (
                        <audio
                          controls
                          src={pendingAttachment.dataUrl}
                          onLoadedMetadata={fixAudioBlobDuration}
                          style={{ height: 34, maxWidth: 220 }}
                        />
                      ) : (
                        <div
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: 6,
                            padding: "6px 10px",
                            borderRadius: 8,
                            border: `1px solid ${T.lineSoft}`,
                            fontSize: 12,
                            color: T.ink,
                            maxWidth: 220,
                          }}
                        >
                          <FileText size={14} />
                          <span
                            style={{
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                              whiteSpace: "nowrap",
                            }}
                          >
                            {pendingAttachment.name}
                          </span>
                        </div>
                      )}
                      <button
                        type="button"
                        onClick={() => setPendingAttachment(null)}
                        style={{
                          border: "none",
                          background: "transparent",
                          cursor: "pointer",
                          color: T.muted,
                          display: "flex",
                          alignItems: "center",
                        }}
                        title={t.chat.removeAttachment}
                      >
                        <X size={16} />
                      </button>
                    </>
                  )}
                </div>
              )}
              {attachError && (
                <div
                  style={{
                    padding: "6px 12px 0 12px",
                    fontSize: 11.5,
                    color: T.rose,
                  }}
                >
                  {attachError}
                </div>
              )}
              {recordError && (
                <div
                  style={{
                    padding: "6px 12px 0 12px",
                    fontSize: 11.5,
                    color: T.rose,
                  }}
                >
                  {recordError}
                </div>
              )}
              {recording ? (
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 10,
                    padding: 12,
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 8,
                      flex: 1,
                      color: T.rose,
                      fontSize: 13,
                      fontWeight: 600,
                    }}
                  >
                    <span
                      style={{
                        width: 9,
                        height: 9,
                        borderRadius: 999,
                        background: T.rose,
                        animation: "wf-rec-pulse 1.2s ease-in-out infinite",
                        flexShrink: 0,
                      }}
                    />
                    {t.chat.recording}
                    <span style={{ color: T.muted, fontWeight: 500 }}>
                      {String(Math.floor(recordSeconds / 60)).padStart(2, "0")}:
                      {String(recordSeconds % 60).padStart(2, "0")}
                    </span>
                  </div>
                  <button
                    type="button"
                    onClick={cancelRecording}
                    title={t.chat.cancelRecording}
                    style={{
                      border: `1px solid ${T.lineSoft}`,
                      background: T.inputBg,
                      color: T.muted,
                      borderRadius: 10,
                      width: 40,
                      height: 40,
                      flexShrink: 0,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      cursor: "pointer",
                    }}
                  >
                    <X size={17} />
                  </button>
                  <Button variant="accent" onClick={stopRecording}>
                    <Square size={14} /> {t.chat.stopRecording}
                  </Button>
                </div>
              ) : (
                <div style={{ display: "flex", gap: 8, padding: 12 }}>
                  <input
                    ref={fileInputRef}
                    type="file"
                    onChange={onAttachmentChange}
                    style={{ display: "none" }}
                  />
                  <button
                    type="button"
                    onClick={pickAttachment}
                    disabled={attaching}
                    title={t.chat.attach}
                    style={{
                      border: `1px solid ${T.lineSoft}`,
                      background: T.inputBg,
                      color: T.muted,
                      borderRadius: 10,
                      width: 40,
                      flexShrink: 0,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      cursor: attaching ? "default" : "pointer",
                    }}
                  >
                    <Paperclip size={17} />
                  </button>
                  <button
                    type="button"
                    onClick={startRecording}
                    disabled={attaching || !!pendingAttachment}
                    title={t.chat.recordVoice}
                    style={{
                      border: `1px solid ${T.lineSoft}`,
                      background: T.inputBg,
                      color: T.muted,
                      borderRadius: 10,
                      width: 40,
                      flexShrink: 0,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      cursor:
                        attaching || pendingAttachment ? "default" : "pointer",
                      opacity: pendingAttachment ? 0.5 : 1,
                    }}
                  >
                    <Mic size={17} />
                  </button>
                  <Input
                    value={draft}
                    onChange={(e) => setDraft(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter" && !e.shiftKey) {
                        e.preventDefault();
                        send();
                      }
                    }}
                    placeholder={t.chat.placeholder}
                    style={{ flex: 1 }}
                  />
                  <Button
                    variant="accent"
                    onClick={send}
                    disabled={!draft.trim() && !pendingAttachment}
                  >
                    <Send size={15} /> {t.chat.send}
                  </Button>
                </div>
              )}
            </div>
          </>
        ) : (
          <div
            style={{
              margin: "auto",
              color: T.muted,
              fontSize: 13,
              textAlign: "center",
              padding: 24,
            }}
          >
            {t.chat.selectEmployee}
          </div>
        )}
      </div>
      {confirmDeleteId && (
        <ConfirmDialog
          text={t.chat.confirmDeleteMsg}
          onCancel={() => setConfirmDeleteId(null)}
          onConfirm={() => deleteMessage(confirmDeleteId)}
        />
      )}
      {confirmDeleteThread && (
        <ConfirmDialog
          text={t.chat.confirmDeleteConversation}
          onCancel={() => setConfirmDeleteThread(false)}
          onConfirm={deleteConversation}
        />
      )}
    </>
  );

  if (!isAdmin) {
    return (
      <div className="wf-chat-layout wf-chat-thread-open">{threadPane}</div>
    );
  }

  return (
    <div
      className={`wf-chat-layout ${selectedEmpId ? "wf-chat-thread-open" : ""}`}
    >
      <div className="wf-chat-list-pane">
        <div style={{ padding: 10, borderBottom: `1px solid ${T.lineSoft}` }}>
          <div style={{ position: "relative" }}>
            <Search
              size={14}
              style={{
                position: "absolute",
                left: 10,
                top: "50%",
                transform: "translateY(-50%)",
                color: T.muted,
              }}
            />
            <Input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder={t.chat.searchEmployee}
              style={{ paddingLeft: 30 }}
            />
          </div>
        </div>
        {allChatItems.length === 0 ? (
          <div
            style={{
              padding: "24px 14px",
              textAlign: "center",
              color: T.muted,
              fontSize: 12.5,
            }}
          >
            {query.trim() ? t.chat.noEmployeesFound : t.chat.noConversations}
          </div>
        ) : (
          allChatItems.map((c) => (
            <button
              key={c.employeeId}
              className={`wf-chat-item ${selectedEmpId === c.employeeId ? "active" : ""}`}
              onClick={() => setSelectedEmpId(c.employeeId)}
            >
              <Avatar
                name={c.employee.name}
                photo={c.employee.photo}
                size={36}
              />
              <div style={{ minWidth: 0, flex: 1 }}>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    gap: 6,
                  }}
                >
                  <div
                    style={{
                      fontWeight: 600,
                      fontSize: 12.5,
                      color: T.ink,
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {c.employee.name}
                  </div>
                  {c.hasConversation && (
                    <div
                      style={{
                        fontSize: 10,
                        color: T.mutedLight,
                        flexShrink: 0,
                      }}
                    >
                      {timeAgoLabel(c.lastMessage.createdAt)}
                    </div>
                  )}
                </div>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    gap: 6,
                    marginTop: 2,
                  }}
                >
                  <div
                    style={{
                      fontSize: 11.5,
                      color: T.textSoft,
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {c.hasConversation
                      ? `${c.lastMessage.senderRole === "admin" ? `${t.chat.youLabel}: ` : ""}${c.lastMessage.content}`
                      : c.employee.code}
                  </div>
                  {c.unread > 0 && (
                    <span
                      style={{
                        background: T.rose,
                        color: "#fff",
                        fontSize: 10,
                        fontWeight: 700,
                        borderRadius: 999,
                        minWidth: 16,
                        height: 16,
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        padding: "0 4px",
                        flexShrink: 0,
                      }}
                    >
                      {c.unread > 9 ? "9+" : c.unread}
                    </span>
                  )}
                </div>
              </div>
            </button>
          ))
        )}
      </div>
      {threadPane}
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
              <th>{t.admAcc.username}</th>
              <th>{t.admAcc.roleLabel}</th>
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
                          ? {
                              background: "rgba(91,141,239,0.14)",
                              color: T.blue,
                            }
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
          text={t.admAcc.confirmDelWithName(confirmDel.name)}
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
      setPinError(t.profile.pinWrongCurrent);
      return;
    }
    if (!pinForm.next.trim() || pinForm.next.trim().length < 4) {
      setPinError(t.profile.pinTooShort);
      return;
    }
    if (pinForm.next.trim() !== pinForm.confirm.trim()) {
      setPinError(t.profile.pinMismatch);
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
            <BadgeCheck size={12} color={T.muted} /> {t.profile.joinedSince}{" "}
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
        <Field label={t.profile.photoLabel}>
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
            <UserCircle2 size={14} /> {t.profile.choosePhoto}
          </label>
        </Field>
        {photoError && (
          <p style={{ fontSize: 12.5, color: T.rose, marginBottom: 10 }}>
            {photoError}
          </p>
        )}
        <Field label={t.profile.phone}>
          <Input value={f.phone} onChange={set("phone")} />
        </Field>
        <Field label={t.profile.email}>
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
          <KeyRound size={16} /> {t.profile.changePin}
        </h3>
        <p style={{ fontSize: 12, color: T.muted, marginBottom: 16 }}>
          {t.profile.pinDesc}
        </p>
        <Field label={t.profile.oldPin}>
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
        <Field label={t.profile.newPin}>
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
        <Field label={t.profile.confirmPin}>
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
            <CheckCircle2 size={14} /> {t.profile.pinChanged}
          </p>
        )}
        <Button variant="accent" onClick={savePin}>
          {t.profile.changePin}
        </Button>
      </Card>

      <AppearanceCard />
      <PushNotificationCard userType="employee" userId={currentEmp.id} />
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

// Per-device push-notification opt-in card, shared by AdminSettings and
// MyProfile. Renders nothing when push isn't supported/configured
// (see usePushSubscription), so it's safe to drop in unconditionally.
function PushNotificationCard({ userType, userId }) {
  const { t } = useLang();
  const {
    supported,
    permission,
    subscribed,
    busy,
    error,
    subscribe,
    unsubscribe,
  } = usePushSubscription(userType, userId);

  if (!supported) return null;

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
        <Bell size={16} /> {t.settings.pushTitle}
      </h3>
      <p style={{ fontSize: 12, color: T.muted, marginBottom: 16 }}>
        {t.settings.pushDesc}
      </p>
      {permission === "denied" ? (
        <p
          style={{
            fontSize: 12,
            color: T.rose,
            display: "flex",
            alignItems: "center",
            gap: 6,
          }}
        >
          <AlertCircle size={14} /> {t.settings.pushBlocked}
        </p>
      ) : (
        <Button
          variant={subscribed ? "ghost" : "accent"}
          onClick={subscribed ? unsubscribe : subscribe}
          disabled={busy}
        >
          {busy ? (
            <Loader2
              size={14}
              style={{ animation: "spin 1s linear infinite" }}
            />
          ) : (
            <Bell size={14} />
          )}{" "}
          {subscribed ? t.settings.pushDisable : t.settings.pushEnable}
        </Button>
      )}
      {subscribed && permission !== "denied" && (
        <p
          style={{
            fontSize: 12,
            color: T.forest,
            display: "flex",
            alignItems: "center",
            gap: 6,
            marginTop: 12,
          }}
        >
          <CheckCircle2 size={14} /> {t.settings.pushEnabledMsg}
        </p>
      )}
      {error && (
        <p style={{ fontSize: 12, color: T.rose, marginTop: 10 }}>
          {t.settings.pushError} {error}
        </p>
      )}
    </Card>
  );
}

/* ---------------------------------------------------------------
   Admin — My Settings (name, photo, appearance)
----------------------------------------------------------------*/
// Self-contained card for the Settings screen: reads/writes its own
// telegram_settings row via useTelegramSettings, and lets a superadmin
// send a one-off test message through the `telegram_notify` edge
// function (bypassing the per-category toggles, since a test message
// should always go through once enabled+configured) to confirm the bot
// is wired up correctly before relying on it for real approvals.
function TelegramSettingsCard() {
  const { t } = useLang();
  const [policy, setPolicy, ready] = useTelegramSettings();
  const [f, setF] = useState(policy);
  useEffect(() => setF(policy), [policy]);
  const [saved, setSaved] = useState(false);
  const [testState, setTestState] = useState("idle"); // idle|sending|ok|error

  if (!ready) return null;

  const set = (k) => (e) => setF((prev) => ({ ...prev, [k]: e.target.value }));
  const toggle = (k) => () => setF((prev) => ({ ...prev, [k]: !prev[k] }));

  const canEnable = !!(f.botToken?.trim() && f.chatId?.trim());

  const save = () => {
    setPolicy({ ...f, enabled: !!f.enabled && canEnable });
    setSaved(true);
  };

  const sendTest = async () => {
    setTestState("sending");
    try {
      const { error } = await supabase.functions.invoke("telegram_notify", {
        body: {
          text: "✅ Test message from Workforce Suite",
          category: "test",
        },
      });
      setTestState(error ? "error" : "ok");
    } catch {
      setTestState("error");
    }
  };

  const categories = [
    { key: "notifyLeave", label: t.settings.telegramCatLeave },
    { key: "notifyOt", label: t.settings.telegramCatOt },
    { key: "notifyAttcorr", label: t.settings.telegramCatAttcorr },
    { key: "notifyShiftswap", label: t.settings.telegramCatShiftswap },
    { key: "notifyPayroll", label: t.settings.telegramCatPayroll },
    { key: "notifyLate", label: t.settings.telegramCatLate },
    { key: "notifyChat", label: t.settings.telegramCatChat },
  ];

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
        <Send size={16} /> {t.settings.telegramTitle}
      </h3>
      <p style={{ fontSize: 12, color: T.muted, marginBottom: 16 }}>
        {t.settings.telegramDesc}
      </p>
      <label
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          marginBottom: 14,
          cursor: "pointer",
        }}
      >
        <input
          type="checkbox"
          checked={!!f.enabled}
          onChange={toggle("enabled")}
          style={{ accentColor: T.forest, width: 16, height: 16 }}
        />
        <span style={{ fontSize: 13, fontWeight: 600, color: T.ink }}>
          {t.settings.telegramEnable}
        </span>
      </label>
      <div className="wf-grid-2">
        <Field label={t.settings.telegramBotTokenLabel}>
          <Input
            type="password"
            value={f.botToken}
            onChange={set("botToken")}
            placeholder={t.settings.telegramBotTokenPlaceholder}
          />
        </Field>
        <Field label={t.settings.telegramChatIdLabel}>
          <Input
            value={f.chatId}
            onChange={set("chatId")}
            placeholder={t.settings.telegramChatIdPlaceholder}
          />
        </Field>
      </div>
      <p style={{ fontSize: 11.5, color: T.muted, marginTop: 6 }}>
        {t.settings.telegramChatIdHint}
      </p>
      <div style={{ marginTop: 16 }}>
        <span
          style={{
            fontSize: 12,
            fontWeight: 600,
            color: T.ink,
            display: "block",
            marginBottom: 8,
          }}
        >
          {t.settings.telegramCategoriesLabel}
        </span>
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {categories.map((c) => (
            <label
              key={c.key}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                cursor: "pointer",
              }}
            >
              <input
                type="checkbox"
                checked={!!f[c.key]}
                onChange={toggle(c.key)}
                style={{ accentColor: T.forest, width: 15, height: 15 }}
              />
              <span style={{ fontSize: 12.5, color: T.textSoft }}>
                {c.label}
              </span>
            </label>
          ))}
        </div>
      </div>
      {!canEnable && (
        <p style={{ fontSize: 11.5, color: T.rose, marginTop: 10 }}>
          {t.settings.telegramNeedsSetup}
        </p>
      )}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          marginTop: 16,
          flexWrap: "wrap",
        }}
      >
        <Button variant="accent" onClick={save}>
          {t.save}
        </Button>
        <Button
          variant="ghost"
          onClick={sendTest}
          disabled={!canEnable || testState === "sending"}
        >
          <Send size={14} />
          {testState === "sending"
            ? t.settings.telegramTestSending
            : t.settings.telegramTestBtn}
        </Button>
      </div>
      {saved && (
        <p
          style={{
            fontSize: 12,
            color: T.forest,
            display: "flex",
            alignItems: "center",
            gap: 6,
            marginTop: 10,
          }}
        >
          <CheckCircle2 size={14} /> {t.settings.telegramSaved}
        </p>
      )}
      {testState === "ok" && (
        <p
          style={{
            fontSize: 12,
            color: T.forest,
            display: "flex",
            alignItems: "center",
            gap: 6,
            marginTop: 10,
          }}
        >
          <CheckCircle2 size={14} /> {t.settings.telegramTestSuccess}
        </p>
      )}
      {testState === "error" && (
        <p
          style={{
            fontSize: 12,
            color: T.rose,
            display: "flex",
            alignItems: "center",
            gap: 6,
            marginTop: 10,
          }}
        >
          <AlertCircle size={14} /> {t.settings.telegramTestError}
        </p>
      )}
    </Card>
  );
}
function AdminSettings({
  currentAdmin,
  admins,
  setAdmins,
  isSuperAdmin,
  saveError,
  soundPolicy,
  setSoundPolicy,
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

  const [soundForm, setSoundForm] = useState(soundPolicy?.preset || "chime");
  const [soundSaved, setSoundSaved] = useState(false);
  const saveSound = () => {
    setSoundPolicy({ preset: soundForm });
    setSoundSaved(true);
  };
  const previewSound = (presetId, presetMode) => {
    playScanBeep(presetMode, presetId);
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

      <PushNotificationCard userType="admin" userId={currentAdmin.id} />

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
            <Bell size={16} /> {t.settings.soundTitle}
          </h3>
          <p style={{ fontSize: 12, color: T.muted, marginBottom: 16 }}>
            {t.settings.soundDesc}
          </p>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {Object.keys(SOUND_PRESETS).map((id) => (
              <label
                key={id}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 10,
                  padding: "10px 12px",
                  borderRadius: 10,
                  border: `1px solid ${
                    soundForm === id ? T.forest : T.divider
                  }`,
                  background: soundForm === id ? T.forestSoft : "transparent",
                  cursor: "pointer",
                }}
              >
                <input
                  type="radio"
                  name="soundPreset"
                  checked={soundForm === id}
                  onChange={() => {
                    setSoundForm(id);
                    setSoundSaved(false);
                  }}
                  style={{ accentColor: T.forest }}
                />
                <span
                  style={{
                    fontSize: 13,
                    fontWeight: 600,
                    color: T.ink,
                    flex: 1,
                  }}
                >
                  {t.settings.soundPresets[id]}
                </span>
                {id !== "silent" && (
                  <button
                    type="button"
                    onClick={(e) => {
                      e.preventDefault();
                      previewSound(id, "in");
                    }}
                    title={t.settings.soundPreview}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 4,
                      background: "none",
                      border: `1px solid ${T.divider}`,
                      borderRadius: 8,
                      padding: "4px 9px",
                      fontSize: 11.5,
                      color: T.muted,
                      cursor: "pointer",
                    }}
                  >
                    <Volume2 size={12} /> {t.settings.soundPreview}
                  </button>
                )}
              </label>
            ))}
          </div>
          {soundSaved && (
            <p
              style={{
                fontSize: 12,
                color: T.forest,
                display: "flex",
                alignItems: "center",
                gap: 6,
                marginTop: 14,
                marginBottom: 10,
              }}
            >
              <CheckCircle2 size={14} /> {t.settings.soundSaved}
            </p>
          )}
          <Button
            variant="accent"
            onClick={saveSound}
            style={{ marginTop: soundSaved ? 0 : 14 }}
          >
            {t.save}
          </Button>
        </Card>
      )}

      {isSuperAdmin && <TelegramSettingsCard />}

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
          <Store size={16} /> {t.dash.empPortalLink}
        </h3>
        <p style={{ fontSize: 12, color: T.muted, marginBottom: 16 }}>
          {t.settings.empPortalDesc}
        </p>
        <EmployeeLinkCard variant="inline" />
      </Card>
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
  onPopupBlocked,
}) {
  const win = window.open("", "_blank", "width=480,height=720");
  if (!win) {
    onPopupBlocked?.();
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
  onPopupBlocked,
}) {
  const win = window.open("", "_blank", "width=420,height=620");
  if (!win) {
    onPopupBlocked?.();
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

// Builds a formal, letterhead-style certificate (Employment / Salary /
// Service / free-form Custom text) for one employee and opens it in a
// print-ready window — mirrors printPayslip/printEmployeeBadge's
// window.open + write + print pattern. Intentionally leaves a blank
// signature line rather than embedding a signature image: these are
// meant to be printed and wet-ink signed/stamped by the signatory.
function printCertificate({
  brandName,
  brandLogo,
  titleText,
  refNo,
  issueDateText,
  bodyText,
  signatoryName,
  signatoryTitle,
  signatureLineLabel,
  generatedOnLabel,
  onPopupBlocked,
}) {
  const win = window.open("", "_blank", "width=560,height=760");
  if (!win) {
    onPopupBlocked?.();
    return;
  }
  const generatedOn = new Date().toLocaleDateString();
  const html = `<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<title>${escapeHtml(titleText)}</title>
<style>
  * { box-sizing: border-box; }
  body {
    font-family: 'Noto Sans Khmer', 'Segoe UI', Arial, sans-serif;
    color: #16213a;
    margin: 0;
    padding: 40px 48px;
    background: #fff;
  }
  .cert-frame {
    border: 2px solid #16213a;
    padding: 32px 36px;
    min-height: 620px;
    display: flex;
    flex-direction: column;
  }
  .header {
    display: flex;
    align-items: center;
    gap: 10px;
    padding-bottom: 14px;
    border-bottom: 1px solid #d8dbe3;
    margin-bottom: 6px;
  }
  .header img { width: 36px; height: 36px; border-radius: 8px; object-fit: cover; }
  .brand { font-size: 15px; font-weight: 700; }
  .meta {
    display: flex;
    justify-content: space-between;
    font-size: 11px;
    color: #6b7280;
    margin-top: 10px;
    margin-bottom: 22px;
  }
  .title {
    text-align: center;
    font-size: 22px;
    font-weight: 700;
    letter-spacing: .02em;
    margin: 8px 0 4px;
  }
  .title-underline {
    width: 90px;
    height: 3px;
    background: #16213a;
    margin: 0 auto 26px;
    border-radius: 2px;
  }
  .body-text {
    font-size: 14px;
    line-height: 2;
    text-align: justify;
    flex: 1;
  }
  .sign-block {
    margin-top: 48px;
    align-self: flex-end;
    text-align: center;
    min-width: 220px;
  }
  .sign-line {
    border-bottom: 1px solid #16213a;
    height: 56px;
    margin-bottom: 6px;
  }
  .sign-caption { font-size: 11px; color: #9ca3af; margin-bottom: 2px; }
  .sign-name { font-size: 13px; font-weight: 700; }
  .sign-title { font-size: 11.5px; color: #6b7280; margin-top: 1px; }
  .footer {
    margin-top: 24px;
    font-size: 9.5px;
    color: #9ca3af;
    text-align: center;
  }
  @media print {
    body { padding: 0; }
  }
</style>
</head>
<body>
  <div class="cert-frame">
    <div class="header">
      ${brandLogo ? `<img src="${escapeHtml(brandLogo)}" />` : ""}
      <div class="brand">${escapeHtml(brandName)}</div>
    </div>
    <div class="meta">
      <span>${refNo ? escapeHtml(refNo) : ""}</span>
      <span>${escapeHtml(issueDateText)}</span>
    </div>
    <div class="title">${escapeHtml(titleText)}</div>
    <div class="title-underline"></div>
    <div class="body-text">${escapeHtml(bodyText)}</div>
    <div class="sign-block">
      <div class="sign-line"></div>
      <div class="sign-caption">${escapeHtml(signatureLineLabel)}</div>
      ${signatoryName ? `<div class="sign-name">${escapeHtml(signatoryName)}</div>` : ""}
      ${signatoryTitle ? `<div class="sign-title">${escapeHtml(signatoryTitle)}</div>` : ""}
    </div>
  </div>
  <div class="footer">${escapeHtml(generatedOnLabel)}: ${escapeHtml(generatedOn)}</div>
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

// Fills a {token} style i18n template string with plain values. Missing
// tokens resolve to an empty string rather than leaving the literal
// "{token}" visible on a printed certificate.
function fillCertTemplate(str, vars) {
  return str.replace(/\{(\w+)\}/g, (_, k) => (vars[k] != null ? vars[k] : ""));
}

// Admin-facing form for issuing a certificate to one employee. Issuance
// is immediate (no request/approval step, per how this company hands
// out certificates) — filling the form and clicking Generate opens the
// print-ready window straight away via printCertificate(). Nothing here
// is persisted; the admin re-fills the form each time, matching how
// Payslip/Badge printing already works in this app.
function CertificateModal({ emp, deptLabel, onClose }) {
  const { t } = useLang();
  const { branding } = useBranding();
  const [type, setType] = useState("employment");
  const [refNo, setRefNo] = useState(
    () => `CERT-${todayStr().replace(/-/g, "")}-${emp.code || ""}`,
  );
  const [issueDate, setIssueDate] = useState(todayStr());
  const [endDate, setEndDate] = useState("");
  const [purpose, setPurpose] = useState("");
  const [customBody, setCustomBody] = useState("");
  const [signatoryName, setSignatoryName] = useState("");
  const [signatoryTitle, setSignatoryTitle] = useState("");
  const [popupBlocked, setPopupBlocked] = useState(false);

  const company = branding?.name?.trim() || t.appName;
  const typeLabel = {
    employment: t.cert.typeEmployment,
    salary: t.cert.typeSalary,
    service: t.cert.typeService,
    custom: t.cert.typeCustom,
  }[type];

  // The big heading printed on the certificate. Defaults to the
  // selected type's label, but the admin can overwrite it (e.g. rename
  // "Custom Text" to "Recommendation Letter"). Once edited by hand it
  // stops following the type dropdown, so switching type again won't
  // clobber a title the admin already typed.
  const [title, setTitle] = useState(typeLabel);
  const [titleTouched, setTitleTouched] = useState(false);
  useEffect(() => {
    if (!titleTouched) setTitle(typeLabel);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [type]);

  const handleGenerate = () => {
    const purposeText = purpose.trim() || t.cert.defaultPurpose;
    const vars = {
      name: emp.name,
      code: emp.code,
      company,
      position: emp.role || deptLabel || "",
      joinDate: fmtDateDisplay(emp.joined) || emp.joined || "",
      salary: fmtMoney(emp.salary),
      purpose: purposeText,
      endPart: endDate
        ? fillCertTemplate(t.cert.toDate, { date: fmtDateDisplay(endDate) })
        : t.cert.toPresent,
    };
    let bodyText;
    if (type === "employment")
      bodyText = fillCertTemplate(t.cert.bodyEmployment, vars);
    else if (type === "salary")
      bodyText = fillCertTemplate(t.cert.bodySalary, vars);
    else if (type === "service")
      bodyText = fillCertTemplate(t.cert.bodyService, vars);
    else bodyText = customBody.trim();

    printCertificate({
      brandName: company,
      brandLogo: branding?.logo || null,
      titleText: title.trim() || typeLabel,
      refNo: refNo.trim(),
      issueDateText: `${t.cert.issuedOn}: ${fmtDateDisplay(issueDate) || issueDate}`,
      bodyText,
      signatoryName: signatoryName.trim(),
      signatoryTitle: signatoryTitle.trim(),
      signatureLineLabel: t.cert.signatureLine,
      generatedOnLabel: t.cert.generatedOn,
      onPopupBlocked: () => setPopupBlocked(true),
    });
  };

  return (
    <Modal
      title={`${t.cert.modalTitle} · ${emp.name}`}
      onClose={onClose}
      width={480}
    >
      <Field label={t.cert.type}>
        <Select value={type} onChange={(e) => setType(e.target.value)}>
          <option value="employment">{t.cert.typeEmployment}</option>
          <option value="salary">{t.cert.typeSalary}</option>
          <option value="service">{t.cert.typeService}</option>
          <option value="custom">{t.cert.typeCustom}</option>
        </Select>
      </Field>
      <Field label={t.cert.titleLabel}>
        <Input
          value={title}
          onChange={(e) => {
            setTitle(e.target.value);
            setTitleTouched(true);
          }}
          placeholder={typeLabel}
        />
      </Field>
      <Field label={t.cert.refNo}>
        <Input value={refNo} onChange={(e) => setRefNo(e.target.value)} />
      </Field>
      <Field label={t.cert.issueDate}>
        <DatePicker
          value={issueDate}
          onChange={(e) => setIssueDate(e.target.value)}
        />
      </Field>
      {type === "service" && (
        <Field label={t.cert.endDate}>
          <DatePicker
            value={endDate}
            onChange={(e) => setEndDate(e.target.value)}
            placeholder={t.cert.endDateHint}
          />
          <span
            style={{
              fontSize: 11,
              color: T.muted,
              display: "block",
              marginTop: 4,
            }}
          >
            {t.cert.endDateHint}
          </span>
        </Field>
      )}
      {type !== "custom" ? (
        <Field label={t.cert.purpose}>
          <Input
            value={purpose}
            onChange={(e) => setPurpose(e.target.value)}
            placeholder={t.cert.purposePlaceholder}
          />
        </Field>
      ) : (
        <Field label={t.cert.customBody}>
          <textarea
            className="wf-input"
            rows={6}
            value={customBody}
            onChange={(e) => setCustomBody(e.target.value)}
            placeholder={t.cert.customBodyPlaceholder}
            style={{ resize: "vertical", lineHeight: 1.6 }}
          />
        </Field>
      )}
      <Field label={t.cert.signatoryName}>
        <Input
          value={signatoryName}
          onChange={(e) => setSignatoryName(e.target.value)}
        />
      </Field>
      <Field label={t.cert.signatoryTitle}>
        <Input
          value={signatoryTitle}
          onChange={(e) => setSignatoryTitle(e.target.value)}
        />
      </Field>
      {popupBlocked && (
        <p style={{ fontSize: 12.5, color: T.rose, marginBottom: 10 }}>
          {t.cert.popupBlocked}
        </p>
      )}
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          gap: 8,
          marginTop: 8,
        }}
      >
        <Button variant="ghost" onClick={onClose}>
          {t.cancel}
        </Button>
        <Button
          variant="accent"
          onClick={handleGenerate}
          disabled={type === "custom" && !customBody.trim()}
        >
          <Award size={14} /> {t.cert.generate}
        </Button>
      </div>
    </Modal>
  );
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
  const [popupBlocked, setPopupBlocked] = useState(false);
  const {
    absentDays,
    leaveDays,
    unpaidLeaveDays,
    unpaidLeaveDeduction,
    ulDeductionPerDay,
    usesCustomUlPolicy,
    usesCustomLatePolicy,
    excessLateDays,
    lateDeduction,
    dailyRate,
    absenceDeduction,
    tax,
    insurance,
    taxRate,
    insuranceRate,
    deductionApplies,
    usesCustomRate,
    usesKhmerMode,
    khmer,
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
      ...(unpaidLeaveDays > 0
        ? [
            {
              label: `${t.pay.unpaidLeaveDed} (${unpaidLeaveDays} × ${fmtMoney(ulDeductionPerDay)})`,
              value: fmtMoney(unpaidLeaveDeduction),
              tone: "neg",
            },
          ]
        : []),
      ...(excessLateDays > 0
        ? [
            {
              label: `${t.pay.lateDed} (${excessLateDays})`,
              value: fmtMoney(lateDeduction),
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
              label: `${t.pay.taxLabel} (${taxRate}%${usesKhmerMode ? ` ${t.pay.effectiveRateNote}` : ""})`,
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
      ...(usesKhmerMode && khmer
        ? [
            {
              label: t.pay.employerNssfCostLabel,
              value: fmtMoney(khmer.employerNssfTotal),
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
      onPopupBlocked: () => setPopupBlocked(true),
    });
  };
  return (
    <>
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
            {usesCustomLatePolicy && (
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
                {t.pay.customLatePolicyBadge}
              </span>
            )}
            {usesCustomUlPolicy && (
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
                {t.pay.customUlPolicyBadge}
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
          {unpaidLeaveDays > 0 && (
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                color: T.rose,
              }}
            >
              <span>
                {t.pay.unpaidLeaveDed} ({unpaidLeaveDays} ×{" "}
                {fmtMoney(ulDeductionPerDay)})
              </span>
              <span style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                -{fmtMoney(unpaidLeaveDeduction)}
              </span>
            </div>
          )}
          {excessLateDays > 0 && (
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                color: T.rose,
              }}
            >
              <span>
                {t.pay.lateDed} ({excessLateDays})
              </span>
              <span style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                -{fmtMoney(lateDeduction)}
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
                {t.pay.taxLabel} ({taxRate}%
                {usesKhmerMode ? ` ${t.pay.effectiveRateNote}` : ""})
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
          {usesKhmerMode && khmer && (
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                color: T.textSoft,
                fontSize: 11.5,
              }}
            >
              <span>{t.pay.employerNssfCostLabel}</span>
              <span style={{ fontFamily: "'JetBrains Mono',monospace" }}>
                {fmtMoney(khmer.employerNssfTotal)}
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
          <span style={{ fontWeight: 600, fontSize: 13 }}>
            {t.pay.netSalary}
          </span>
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
      {popupBlocked && (
        <LoginActAlertDialog
          title={t.popupBlockedTitle}
          message={t.popupBlockedPayslip}
          onClose={() => setPopupBlocked(false)}
        />
      )}
    </>
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
    const willBePaid = !payrollPaid[key];
    setPayrollPaid({ ...payrollPaid, [key]: willBePaid });

    // Only worth checking "is everyone done now?" on the transition
    // into paid — unmarking someone can't complete the batch.
    if (willBePaid) {
      const allNowPaid = activeEmployees.every((e) =>
        e.id === empId ? true : !!payrollPaid[`${e.id}-${mk}`],
      );
      if (allNowPaid && activeEmployees.length > 0) {
        supabase.functions
          .invoke("telegram_notify", {
            body: {
              text: `ប្រាក់ខែសម្រាប់ខែ ${mk} ត្រូវបានបើកផ្តល់ចប់សព្វគ្រប់ សម្រាប់និយោជិកទាំង ${activeEmployees.length} នាក់ សរុប $${totalNet.toFixed(2)}`,
              category: "payroll",
            },
          })
          .then(({ error }) => {
            if (error)
              console.error("[telegram] payroll notify failed:", error.message);
          });
      }
    }
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
          {t.pay.monthLabel}
        </span>
        <Select
          value={mk}
          onChange={(e) => setMk(e.target.value)}
          style={{ width: "auto", minWidth: 170 }}
        >
          {availableMonths.map((m) => (
            <option key={m} value={m}>
              {monthLabel(m)}
              {m === currentMk ? ` ${t.pay.currentMonthTag}` : ""}
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
            {t.pay.viewingPastMonth}
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
              const {
                net,
                absentDays,
                unpaidLeaveDays,
                excessLateDays,
                otHours,
                usesCustomRate,
                usesCustomLatePolicy,
                usesCustomUlPolicy,
              } = computePayroll(
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
                          {usesCustomLatePolicy && (
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
                              {t.pay.customLatePolicyBadge}
                            </span>
                          )}
                          {usesCustomUlPolicy && (
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
                              {t.pay.customUlPolicyBadge}
                            </span>
                          )}
                        </div>
                        {absentDays > 0 && (
                          <div style={{ fontSize: 10.5, color: T.rose }}>
                            {t.att.absentDays} {absentDays}
                          </div>
                        )}
                        {unpaidLeaveDays > 0 && (
                          <div style={{ fontSize: 10.5, color: T.rose }}>
                            {t.pay.unpaidLeaveDed}: {unpaidLeaveDays}
                          </div>
                        )}
                        {excessLateDays > 0 && (
                          <div style={{ fontSize: 10.5, color: T.rose }}>
                            {t.pay.lateDed}: {excessLateDays}
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
    { id: "messages", label: n.messages, icon: MessageCircle },
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
    { id: "training", label: n.training, icon: GraduationCap },
    {
      id: "attcorr",
      label: n.attCorrection,
      icon: CalendarClock,
      permission: "approveRequests",
    },
    {
      id: "shiftswap",
      label: n.shiftSwap,
      icon: Repeat,
      permission: "approveRequests",
    },
    {
      id: "assets",
      label: n.assets,
      icon: Package,
      permission: "manageAssets",
    },
    {
      id: "docExpiry",
      label: n.docExpiry,
      icon: AlertCircle,
      permission: "manageDocuments",
    },
    {
      id: "recruitment",
      label: n.recruitment,
      icon: Briefcase,
      permission: "manageRecruitment",
    },
    {
      id: "onboarding",
      label: n.onboarding,
      icon: ListChecks,
      permission: "manageRecruitment",
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
    { id: "messages", label: n.messages, icon: MessageCircle },
    { id: "leave", label: n.myLeave, icon: CalendarDays },
    { id: "ot", label: n.myOvertime, icon: Timer },
    { id: "payroll", label: n.myPayroll, icon: Wallet },
    { id: "review", label: n.myPerformance, icon: Star },
    { id: "training", label: n.myTraining, icon: GraduationCap },
    { id: "attcorr", label: n.myAttCorrection, icon: CalendarClock },
    { id: "shiftswap", label: n.myShiftSwap, icon: Repeat },
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
  const [assets, setAssets, assetsReady] = useSupabaseArray("assets", {
    fromDb: (r) => ({
      id: r.id,
      name: r.name,
      category: r.category,
      serial: r.serial_number,
      assignedTo: r.assigned_to,
      status: r.status,
      purchaseDate: r.purchase_date,
      notes: r.notes,
    }),
    toDb: (r) => ({
      id: r.id,
      name: r.name,
      category: r.category,
      serial_number: r.serial,
      assigned_to: r.assignedTo || null,
      status: r.status,
      purchase_date: r.purchaseDate || null,
      notes: r.notes,
    }),
    audit: true,
    actorRef,
  });
  const [jobPostings, setJobPostings, jobPostingsReady] = useSupabaseArray(
    "job_postings",
    {
      fromDb: (r) => ({
        id: r.id,
        title: r.title,
        departmentId: r.department_id,
        employmentType: r.employment_type,
        status: r.status,
        description: r.description,
        postedDate: r.posted_date,
        closingDate: r.closing_date || "",
      }),
      toDb: (r) => ({
        id: r.id,
        title: r.title,
        department_id: r.departmentId || null,
        employment_type: r.employmentType,
        status: r.status,
        description: r.description,
        posted_date: r.postedDate || null,
        closing_date: r.closingDate || null,
      }),
      audit: true,
      actorRef,
      entityLabel: (r) => r.title || "?",
    },
  );
  const [candidates, setCandidates, candidatesReady] = useSupabaseArray(
    "candidates",
    {
      fromDb: (r) => ({
        id: r.id,
        name: r.name,
        phone: r.phone,
        email: r.email,
        jobPostingId: r.job_posting_id,
        stage: r.stage,
        notes: r.notes,
        appliedDate: r.applied_date,
        resumeFileName: r.resume_file_name || "",
        resumeMimeType: r.resume_mime_type || "",
        resumeDataUrl: r.resume_data_url || "",
      }),
      toDb: (r) => ({
        id: r.id,
        name: r.name,
        phone: r.phone,
        email: r.email,
        job_posting_id: r.jobPostingId || null,
        stage: r.stage,
        notes: r.notes,
        applied_date: r.appliedDate || null,
        resume_file_name: r.resumeFileName || null,
        resume_mime_type: r.resumeMimeType || null,
        resume_data_url: r.resumeDataUrl || null,
      }),
      audit: true,
      actorRef,
      entityLabel: (r) => r.name || "?",
    },
  );
  const [onboardingTasks, setOnboardingTasks, onboardingTasksReady] =
    useSupabaseArray("onboarding_tasks", {
      fromDb: (r) => ({
        id: r.id,
        employeeId: r.employee_id,
        type: r.type,
        title: r.title,
        done: !!r.done,
        dueDate: r.due_date || "",
        orderIndex: r.order_index ?? 0,
        createdAt: r.created_at,
      }),
      toDb: (r) => ({
        id: r.id,
        employee_id: r.employeeId,
        type: r.type,
        title: r.title,
        done: r.done,
        due_date: r.dueDate || null,
        order_index: r.orderIndex ?? 0,
      }),
      audit: true,
      actorRef,
      entityLabel: (r) => r.title || "?",
    });
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
        useCustomLatePolicy: !!r.use_custom_late_policy,
        customLateGraceCount: r.custom_late_grace_count,
        customLateDeductionType: r.custom_late_deduction_type,
        customLateDeductionValue: r.custom_late_deduction_value,
        useCustomUlPolicy: !!r.use_custom_ul_policy,
        customUlDeductionType: r.custom_ul_deduction_type,
        customUlDeductionValue: r.custom_ul_deduction_value,
        dependents: r.dependents ?? 0,
        // Per-employee override that hides Messages & Voice Call for this
        // one person even when the company-wide toggle (role_permissions /
        // EMPLOYEE_MODULES) has it turned on. One-directional only: this
        // can never turn the module back ON for someone if the
        // company-wide switch is off — both gates must pass.
        messagesDisabled: !!r.messages_disabled,
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
        use_custom_late_policy: !!r.useCustomLatePolicy,
        custom_late_grace_count: r.useCustomLatePolicy
          ? Number(r.customLateGraceCount) || 0
          : null,
        custom_late_deduction_type: r.useCustomLatePolicy
          ? r.customLateDeductionType || "fixed"
          : null,
        custom_late_deduction_value: r.useCustomLatePolicy
          ? Number(r.customLateDeductionValue) || 0
          : null,
        use_custom_ul_policy: !!r.useCustomUlPolicy,
        custom_ul_deduction_type: r.useCustomUlPolicy
          ? r.customUlDeductionType || "fullDay"
          : null,
        custom_ul_deduction_value: r.useCustomUlPolicy
          ? Number(r.customUlDeductionValue) || 0
          : null,
        dependents: Number(r.dependents) || 0,
        messages_disabled: !!r.messagesDisabled,
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
      // check_in/check_out are Postgres `time` columns, which reject an
      // empty string ("") with a 400 — only null or a valid "HH:MM" is
      // accepted. Any upstream code path (manual entry, self-punch,
      // corrections, leave approval) that leaves these as "" instead of
      // null would otherwise fail the whole batch upsert silently
      // (console-only error), so normalize defensively right here.
      check_in: r.checkIn || null,
      check_out: r.checkOut || null,
      status: r.status,
      check_in_loc: r.checkInLoc,
      check_out_loc: r.checkOutLoc,
    }),
    notify: ({ type, row }) => {
      if (type === "create" && row.status === "late") {
        const emp = employees.find((e) => e.id === row.employeeId);
        const empShift = shifts.find((s) => s.id === emp?.shiftId);
        const mins = lateMinutesForShift(row.checkIn, empShift);
        const lateLabel = formatLateDuration(mins, "km"); // e.g. "មកយឺត 4 ម៉ោង 49 នាទី"
        return {
          userType: "admin",
          title: "បុគ្គលិកមកយឺត",
          body: `${emp?.name || "?"} (${emp?.code || row.employeeId}) បានចូលធ្វើការនៅម៉ោង ${row.checkIn}${lateLabel ? ` (${lateLabel})` : ""}`,
          page: "late",
          portal: "admin",
        };
      }
      return null;
    },
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
      notify: ({ type, row, old }) => {
        const LEAVE_TYPE_LABEL = getLeaveTypeLabel("km");
        if (type === "create" && row.status === "pending") {
          const emp = employees.find((e) => e.id === row.employeeId);
          return {
            userType: "admin",
            title: "សំណើសុំច្បាប់ថ្មី",
            body: `${emp?.name || "?"} (${emp?.code || row.employeeId}) បានស្នើសុំ${LEAVE_TYPE_LABEL[row.type] || "ច្បាប់"} (${row.startDate} – ${row.endDate})${row.reason ? `\nមូលហេតុ: ${row.reason}` : ""}`,
            page: "leave",
            portal: "admin",
            tag: `leave-req-${row.id}`,
            entityId: row.id,
          };
        }
        if (
          type === "update" &&
          old?.status !== row.status &&
          (row.status === "approved" || row.status === "rejected")
        ) {
          return {
            userType: "employee",
            userId: row.employeeId,
            title:
              row.status === "approved"
                ? "សំណើសុំច្បាប់របស់អ្នកត្រូវបានអនុម័ត"
                : "សំណើសុំច្បាប់របស់អ្នកត្រូវបានបដិសេធ",
            body: `${LEAVE_TYPE_LABEL[row.type] || "ច្បាប់"} (${row.startDate} – ${row.endDate})`,
            page: "leave",
            portal: "employee",
            tag: `leave-dec-${row.id}`,
          };
        }
        return null;
      },
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
      notify: ({ type, row, old }) => {
        if (type === "create" && row.status === "pending") {
          const emp = employees.find((e) => e.id === row.employeeId);
          return {
            userType: "admin",
            title: "សំណើសុំ OT ថ្មី",
            body: `${emp?.name || "?"} (${emp?.code || row.employeeId}) បានស្នើសុំ OT ចំនួន ${row.hours} ម៉ោង នៅថ្ងៃទី ${row.date}${row.reason ? `\nមូលហេតុ: ${row.reason}` : ""}`,
            page: "ot",
            portal: "admin",
            tag: `ot-req-${row.id}`,
            entityId: row.id,
          };
        }
        if (
          type === "update" &&
          old?.status !== row.status &&
          (row.status === "approved" || row.status === "rejected")
        ) {
          return {
            userType: "employee",
            userId: row.employeeId,
            title:
              row.status === "approved"
                ? "សំណើសុំ OT របស់អ្នកត្រូវបានអនុម័ត"
                : "សំណើសុំ OT របស់អ្នកត្រូវបានបដិសេធ",
            body: `${row.date} · ${row.hours} ម៉ោង`,
            page: "ot",
            portal: "employee",
            tag: `ot-dec-${row.id}`,
          };
        }
        return null;
      },
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
  const [trainings, setTrainings, trainingsReady] = useSupabaseArray(
    "trainings",
    {
      fromDb: (r) => ({
        id: r.id,
        employeeId: r.employee_id,
        courseName: r.course_name,
        provider: r.provider,
        category: r.category,
        startDate: r.start_date,
        completionDate: r.completion_date,
        certExpiry: r.cert_expiry,
        status: r.status,
        notes: r.notes,
      }),
      toDb: (r) => ({
        id: r.id,
        employee_id: r.employeeId,
        course_name: r.courseName,
        provider: r.provider,
        category: r.category,
        start_date: r.startDate || null,
        completion_date: r.completionDate || null,
        cert_expiry: r.certExpiry || null,
        status: r.status,
        notes: r.notes,
      }),
      audit: true,
      actorRef,
      entityLabel: (r) => r.courseName,
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
        expiryDate: r.expiry_date || "",
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
        expiry_date: r.expiryDate || null,
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
      notify: ({ type, row, old }) => {
        if (type === "create" && row.status === "pending") {
          const emp = employees.find((e) => e.id === row.employeeId);
          return {
            userType: "admin",
            title: "សំណើកែតម្រូវវត្តមានថ្មី",
            body: `${emp?.name || "?"} (${emp?.code || row.employeeId}) បានស្នើសុំកែតម្រូវវត្តមាននៅថ្ងៃទី ${row.date}${row.reason ? `\nមូលហេតុ: ${row.reason}` : ""}`,
            page: "attcorr",
            portal: "admin",
            tag: `ac-req-${row.id}`,
            entityId: row.id,
          };
        }
        if (
          type === "update" &&
          old?.status !== row.status &&
          (row.status === "approved" || row.status === "rejected")
        ) {
          return {
            userType: "employee",
            userId: row.employeeId,
            title:
              row.status === "approved"
                ? "សំណើកែតម្រូវវត្តមានរបស់អ្នកត្រូវបានអនុម័ត"
                : "សំណើកែតម្រូវវត្តមានរបស់អ្នកត្រូវបានបដិសេធ",
            body: row.date,
            page: "attcorr",
            portal: "employee",
            tag: `ac-dec-${row.id}`,
          };
        }
        return null;
      },
    });
  const [shiftSwapRequests, setShiftSwapRequests, ssReady] = useSupabaseArray(
    "shift_swap_requests",
    {
      fromDb: (r) => ({
        id: r.id,
        employeeId: r.employee_id,
        fromShiftId: r.from_shift_id,
        toShiftId: r.to_shift_id,
        date: r.date,
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
        from_shift_id: r.fromShiftId,
        to_shift_id: r.toShiftId,
        date: r.date,
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
      notify: ({ type, row, old }) => {
        if (type === "create" && row.status === "pending") {
          const emp = employees.find((e) => e.id === row.employeeId);
          return {
            userType: "admin",
            title: "សំណើដូរវេនថ្មី",
            body: `${emp?.name || "?"} (${emp?.code || row.employeeId}) បានស្នើសុំគ្របដណ្តប់វេននៅថ្ងៃទី ${row.date}${row.reason ? `\nមូលហេតុ: ${row.reason}` : ""}`,
            page: "shiftswap",
            portal: "admin",
            tag: `ss-req-${row.id}`,
            entityId: row.id,
          };
        }
        if (
          type === "update" &&
          old?.status !== row.status &&
          (row.status === "approved" || row.status === "rejected")
        ) {
          return {
            userType: "employee",
            userId: row.employeeId,
            title:
              row.status === "approved"
                ? "សំណើដូរវេនរបស់អ្នកត្រូវបានអនុម័ត"
                : "សំណើដូរវេនរបស់អ្នកត្រូវបានបដិសេធ",
            body: row.date,
            page: "shiftswap",
            portal: "employee",
            tag: `ss-dec-${row.id}`,
          };
        }
        return null;
      },
    },
  );
  // Admin <-> Employee direct messages. One thread per employee — admins
  // share a single mailbox per employee (same philosophy as the shared
  // admin notification mailbox above), so whichever admin replies, every
  // other admin sees the same conversation and read state. Employees only
  // ever see their own thread.
  //
  // Requires this table in Supabase (SQL editor):
  //   create table messages (
  //     id text primary key,
  //     employee_id text not null references employees(id) on delete cascade,
  //     sender_role text not null,       -- 'admin' | 'employee'
  //     sender_id text,                  -- admin id or employee id of the sender
  //     sender_name text,                -- snapshot of the sender's display name
  //     content text not null,
  //     attachment_data text,            -- base64 data URL of an attached photo/file, or null
  //     attachment_name text,            -- original file name, or null
  //     attachment_type text,            -- MIME type, or null
  //     edited_at timestamptz,           -- set when the sender edits their own message
  //     deleted boolean not null default false,
  //     read_by_admin boolean not null default false,
  //     read_by_employee boolean not null default false,
  //     created_at timestamptz not null default now()
  //   );
  //   alter publication supabase_realtime add table messages;
  // Existing installs: run
  //   alter table messages add column attachment_data text;
  //   alter table messages add column attachment_name text;
  //   alter table messages add column attachment_type text;
  //   alter table messages add column edited_at timestamptz;
  //   alter table messages add column deleted boolean not null default false;
  // Plus a row-level-security policy allowing admins to read/write every
  // row, and employees to read/write only rows where employee_id matches
  // their own id.
  //
  // Attachments are stored inline as base64 data URLs (same approach as
  // employee profile photos elsewhere in this file) rather than in a
  // Supabase Storage bucket — this keeps deployment to a single SQL
  // migration with no bucket/policy setup required. Images are
  // downscaled/compressed before upload (see fileToCompressedChatImageDataUrl);
  // other file types are capped at MAX_CHAT_FILE_BYTES so a single
  // attachment can't bloat every client's sync payload.
  //
  // Edit/delete is Messenger/Telegram-style and sender-only: deleting a
  // message keeps the row (for thread continuity/ordering) but clears its
  // content/attachment and sets deleted=true, rendered as a muted "message
  // was deleted" placeholder rather than actually removing the row.
  //
  // Like attendance, chat history grows without bound, so only the last
  // 90 days are kept live in the app — plenty for an ongoing conversation
  // without dragging years of messages into every session.
  // Read-only peek at the Telegram policy so the chat notify() callback
  // below can opt individual messages out of Telegram forwarding (see
  // `skipTelegram` in the generic sendTelegram helper) without owning
  // the settings row — TelegramSettingsCard still has its own
  // read/write instance of this hook for the actual Settings screen.
  const [telegramPolicy] = useTelegramSettings();
  const [messages, setMessages, msgReady] = useSupabaseArray("messages", {
    dateField: "created_at",
    daysBack: 90,
    fromDb: (r) => ({
      id: r.id,
      employeeId: r.employee_id,
      senderRole: r.sender_role,
      senderId: r.sender_id,
      senderName: r.sender_name,
      content: r.content,
      attachmentData: r.attachment_data || null,
      attachmentName: r.attachment_name || null,
      attachmentType: r.attachment_type || null,
      editedAt: r.edited_at || null,
      deleted: !!r.deleted,
      readByAdmin: !!r.read_by_admin,
      readByEmployee: !!r.read_by_employee,
      createdAt: r.created_at,
    }),
    toDb: (r) => ({
      id: r.id,
      employee_id: r.employeeId,
      sender_role: r.senderRole,
      sender_id: r.senderId,
      sender_name: r.senderName,
      content: r.content,
      attachment_data: r.attachmentData || null,
      attachment_name: r.attachmentName || null,
      attachment_type: r.attachmentType || null,
      edited_at: r.editedAt || null,
      deleted: !!r.deleted,
      read_by_admin: !!r.readByAdmin,
      read_by_employee: !!r.readByEmployee,
      created_at: r.createdAt,
    }),
    notify: ({ type, row }) => {
      if (type !== "create") return null;
      const emp = employees.find((e) => e.id === row.employeeId);
      const body =
        row.content || (row.attachmentData ? t.chat.photoAttachment : "");
      if (row.senderRole === "employee") {
        return {
          userType: "admin",
          title: `${t.chat.newMessageFrom} ${emp?.name || row.senderName || "?"}`,
          body,
          page: "messages",
          portal: "admin",
          tag: `msg-emp-${row.employeeId}`,
          entityId: row.employeeId,
          skipTelegram: !telegramPolicy.notifyChat,
        };
      }
      return {
        userType: "employee",
        userId: row.employeeId,
        title: t.chat.newMessageFromAdmin,
        body,
        page: "messages",
        portal: "employee",
        tag: `msg-admin-${row.id}`,
      };
    },
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
      qrSecret: r.qr_secret || "",
    }),
    toDb: (r) => ({
      id: r.id,
      name: r.name,
      lat: r.lat,
      lng: r.lng,
      radius: r.radius,
      qr_secret: r.qrSecret || null,
    }),
    orderBy: "name",
    audit: true,
    actorRef,
  });
  const [otPolicy, setOtPolicy, otPolicyReady] = useOtPolicy();
  const [payrollPolicy, setPayrollPolicy, payrollPolicyReady] =
    usePayrollPolicy();
  const [soundPolicy, setSoundPolicy, soundPolicyReady] = useSoundPolicy();
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
        manageRecruitment: !!r.manage_recruitment,
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
        training: r.emp_training ?? true,
        attcorr: r.emp_attcorr ?? true,
        shiftswap: r.emp_shiftswap ?? true,
        documents: r.emp_documents ?? true,
        loginActivity: r.emp_login_activity ?? true,
        profile: r.emp_profile ?? true,
        messages: r.emp_messages ?? true,
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
        manage_recruitment: r.manageRecruitment,
        emp_announcements: r.announcements,
        emp_attendance: r.attendance,
        emp_leave: r.leave,
        emp_ot: r.ot,
        emp_payroll: r.payroll,
        emp_review: r.review,
        emp_training: r.training,
        emp_attcorr: r.attcorr,
        emp_shiftswap: r.shiftswap,
        emp_documents: r.documents,
        emp_login_activity: r.loginActivity,
        emp_profile: r.profile,
        emp_messages: r.messages,
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

  // A tap on a Web Push notification (see sw.js "notificationclick") posts
  // this message to any already-open tab instead of doing a hard
  // navigation, so in-memory state (unsaved forms, scroll position) isn't
  // lost. Switches portal first when the notification was for the other
  // one (e.g. an admin push arriving while an employee tab is focused),
  // then jumps to the target page.
  useEffect(() => {
    if (!("serviceWorker" in navigator)) return;
    const onMessage = (event) => {
      if (event.data?.type !== "push-navigate") return;
      const { page: targetPage, portal: targetPortal } = event.data;
      if (targetPortal && targetPortal !== portal) {
        goPortal(targetPortal);
      }
      if (targetPage) setPage(targetPage);
    };
    navigator.serviceWorker.addEventListener("message", onMessage);
    return () =>
      navigator.serviceWorker.removeEventListener("message", onMessage);
  }, [portal, goPortal, setPage]);

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
    soundPolicyReady &&
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
  // The kiosk portal (see the `if (portal === "kiosk")` early-return
  // below) is a public, unauthenticated tablet display — it must never
  // pick up an authenticated identity. Without this guard, `role` fell
  // through to `sessionEmployee`, which is `useLocalStorage`-backed and
  // therefore shared by every tab in that browser: if the same device
  // was ever used to log in as a particular employee (e.g. testing the
  // employee portal before mounting the kiosk tablet), the leftover
  // localStorage value made this "passive QR display" tab silently
  // resolve to that employee's identity too. Hooks still run
  // unconditionally above the early-return, so useVoiceCall (and the
  // Messages/notification wiring) would then actually subscribe and
  // ring for that employee's calls on the kiosk tablet — a ghost
  // duplicate of the ring the real recipient already gets on their own
  // device. Forcing role to null here means selfId inside useVoiceCall
  // resolves to null and its subscribing effect never fires for kiosk.
  const role =
    portal === "kiosk"
      ? null
      : portal === "admin"
        ? currentAdmin
          ? "admin"
          : null
        : sessionEmployee || null;
  const currentEmp =
    role && role !== "admin" ? employees.find((e) => e.id === role) : null;
  // Single source of truth for "can this session use Messages/Call" —
  // admin side checks the rank permission matrix (superadmin always
  // passes), employee side checks the company-wide module toggle AND
  // this employee's own messagesDisabled override. Both gates must
  // pass — the per-employee flag can only take the module away, never
  // grant it back when the company-wide switch is off.
  const canUseMessages =
    role === "admin"
      ? isSuperAdmin || can("canMessage")
      : moduleEnabled("messages") && !currentEmp?.messagesDisabled;
  const loggedIn = role === "admin" || !!currentEmp;
  // Same merge feeds the side nav so the "Messages" item itself
  // disappears for an employee whose override is on, not just the page
  // route.
  const employeeModulesForNav =
    role !== "admin" && currentEmp?.messagesDisabled
      ? { ...employeeModules, messages: false }
      : employeeModules;
  const nav =
    role === "admin"
      ? buildNavAdmin(t.nav).filter(
          (n) =>
            (!n.superadminOnly || isSuperAdmin) &&
            (!n.permission || can(n.permission)),
        )
      : buildNavEmployee(t.nav, employeeModulesForNav);
  const bottomNav =
    role !== "admin" ? buildBottomNavEmployee(t.nav, employeeModules) : null;

  // Instantiated once here (not inside MessagesPage) so an incoming call
  // still rings no matter which page is currently open. selfId inside
  // the hook is null until currentAdmin/currentEmp resolve, so it's safe
  // to always call this even before login finishes.
  const voiceCall = useVoiceCall({
    role,
    currentAdmin,
    currentEmp,
    employees,
    admins,
    t,
    canUseMessages,
  });

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
    // Shape-of-the-app skeleton (sidebar rail + topbar + card grid)
    // instead of a bare spinner on a blank page — it gives the person
    // something to look at that already resembles where they're headed,
    // so the first load reads as "loading the dashboard" rather than
    // "nothing has happened yet". Pure CSS shimmer, no extra libs.
    const shimmer = {
      background:
        "linear-gradient(90deg, var(--wf-line-soft) 25%, var(--wf-line) 37%, var(--wf-line-soft) 63%)",
      backgroundSize: "400% 100%",
      animation: "wfShimmer 1.4s ease infinite",
      borderRadius: 8,
    };
    return (
      <div
        style={{
          minHeight: 500,
          display: "flex",
          background: T.paper,
          borderRadius: 12,
          overflow: "hidden",
        }}
      >
        <div
          style={{
            width: 220,
            padding: 18,
            display: "flex",
            flexDirection: "column",
            gap: 10,
            borderRight: `1px solid ${T.lineSoft}`,
          }}
        >
          <div
            style={{ ...shimmer, height: 28, width: "70%", marginBottom: 14 }}
          />
          {Array.from({ length: 7 }).map((_, i) => (
            <div
              key={i}
              style={{ ...shimmer, height: 16, width: `${85 - (i % 3) * 12}%` }}
            />
          ))}
        </div>
        <div style={{ flex: 1, padding: 22 }}>
          <div
            style={{ ...shimmer, height: 22, width: 200, marginBottom: 20 }}
          />
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))",
              gap: 14,
              marginBottom: 20,
            }}
          >
            {Array.from({ length: 4 }).map((_, i) => (
              <div key={i} style={{ ...shimmer, height: 70 }} />
            ))}
          </div>
          <div style={{ ...shimmer, height: 220 }} />
        </div>
        <style>
          {
            "@keyframes wfShimmer{0%{background-position:100% 50%}100%{background-position:0% 50%}}"
          }
        </style>
      </div>
    );
  }

  // The kiosk portal is public and unauthenticated by design (see
  // officeKioskUrl / KioskDisplay) — it renders standalone, before the
  // login-screen branch below, and never touches sessionAdmin/
  // sessionEmployee.
  if (portal === "kiosk") {
    return (
      <KioskDisplay
        officeId={routedPage}
        offices={offices}
        branding={branding}
      />
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
              {canUseMessages && (
                <ChatQuickAccess
                  role={role}
                  currentEmp={currentEmp}
                  messages={messages}
                  setPage={setPage}
                />
              )}
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
                shiftSwapRequests={shiftSwapRequests}
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
                  offices={offices}
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
              {page === "assets" &&
                role === "admin" &&
                (isSuperAdmin || can("manageAssets")) && (
                  <Assets
                    assets={assets}
                    setAssets={setAssets}
                    employees={employees}
                    isSuperAdmin={isSuperAdmin || can("manageAssets")}
                  />
                )}
              {page === "docExpiry" &&
                role === "admin" &&
                (isSuperAdmin || can("manageDocuments")) && (
                  <DocumentExpiryPage
                    documents={documents}
                    employees={employees}
                  />
                )}
              {page === "recruitment" &&
                role === "admin" &&
                (isSuperAdmin || can("manageRecruitment")) && (
                  <Recruitment
                    jobPostings={jobPostings}
                    setJobPostings={setJobPostings}
                    candidates={candidates}
                    setCandidates={setCandidates}
                    departments={departments}
                    isSuperAdmin={isSuperAdmin || can("manageRecruitment")}
                  />
                )}
              {page === "onboarding" &&
                role === "admin" &&
                (isSuperAdmin || can("manageRecruitment")) && (
                  <OnboardingOffboarding
                    employees={employees}
                    onboardingTasks={onboardingTasks}
                    setOnboardingTasks={setOnboardingTasks}
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
                    soundPreset={soundPolicy.preset}
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
              {page === "training" && role === "admin" && (
                <Trainings
                  trainings={trainings}
                  setTrainings={setTrainings}
                  employees={employees}
                  isSuperAdmin={isSuperAdmin}
                />
              )}
              {page === "training" &&
                role !== "admin" &&
                currentEmp &&
                moduleEnabled("training") && (
                  <MyTrainings currentEmp={currentEmp} trainings={trainings} />
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
              {page === "shiftswap" &&
                (role === "admin" || moduleEnabled("shiftswap")) && (
                  <ShiftSwapRequests
                    role={role}
                    currentAdmin={currentAdmin}
                    currentEmp={currentEmp}
                    employees={employees}
                    setEmployees={setEmployees}
                    shifts={shifts}
                    admins={admins}
                    shiftSwapRequests={shiftSwapRequests}
                    setShiftSwapRequests={setShiftSwapRequests}
                    isSuperAdmin={isSuperAdmin || can("approveRequests")}
                    canApprove={isSuperAdmin || can("approveRequests")}
                  />
                )}
              {page === "messages" && canUseMessages && (
                <MessagesPage
                  role={role}
                  currentAdmin={currentAdmin}
                  currentEmp={currentEmp}
                  employees={employees}
                  messages={messages}
                  setMessages={setMessages}
                  activeCall={voiceCall.call}
                  onStartCall={voiceCall.startCall}
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
                  soundPolicy={soundPolicy}
                  setSoundPolicy={setSoundPolicy}
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
      <CallOverlay
        call={voiceCall.call}
        callError={voiceCall.callError}
        onAccept={voiceCall.acceptCall}
        onReject={voiceCall.rejectCall}
        onEnd={voiceCall.endCall}
        onToggleMute={voiceCall.toggleMute}
        onDismissError={voiceCall.clearCallError}
      />
    </BrandingContext.Provider>
  );
}

// Installs the app as a PWA. Everything except the service worker is
// generated right here and injected into <head> at runtime — the
// manifest and the app icon are both built as data: URIs, so there's
// no manifest.json or icons/ folder to deploy separately. The one
// exception is the service worker itself (registered at /sw.js):
// browsers require it to be a real same-origin script file reachable
// at a stable URL — that's a platform security rule, not something we
// can inline into the JS bundle. Deliberately does NOT cache or queue
// Supabase writes — checking in/out, payroll, etc. still need a live
// connection; this only makes the shell load fast and work when the
// network blips.
const PWA_ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">' +
  '<rect width="100" height="100" rx="22" fill="#0A0F1A"/>' +
  '<circle cx="50" cy="50" r="30" fill="none" stroke="#1FA26B" stroke-width="4.5"/>' +
  '<path d="M32 52 L45 65 L70 35" fill="none" stroke="#F0A83B" ' +
  'stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/>' +
  "</svg>";
const PWA_ICON_URL = `data:image/svg+xml;base64,${btoa(PWA_ICON_SVG)}`;

function usePwaSetup() {
  useEffect(() => {
    const addTag = (tag, attrs) => {
      if (
        document.querySelector(
          `${tag}[${Object.keys(attrs)[0]}="${Object.values(attrs)[0]}"]`,
        )
      )
        return; // already present (e.g. hot-reload) — don't duplicate
      const el = document.createElement(tag);
      Object.entries(attrs).forEach(([k, v]) => el.setAttribute(k, v));
      document.head.appendChild(el);
    };

    const manifest = {
      name: "Workforce Suite",
      short_name: "Workforce",
      start_url: "/",
      scope: "/",
      display: "standalone",
      background_color: "#0A0F1A",
      theme_color: "#0A0F1A",
      icons: [
        {
          src: PWA_ICON_URL,
          sizes: "192x192",
          type: "image/svg+xml",
          purpose: "any",
        },
        {
          src: PWA_ICON_URL,
          sizes: "512x512",
          type: "image/svg+xml",
          purpose: "any",
        },
        {
          src: PWA_ICON_URL,
          sizes: "512x512",
          type: "image/svg+xml",
          purpose: "maskable",
        },
      ],
    };
    const manifestUrl = `data:application/manifest+json,${encodeURIComponent(
      JSON.stringify(manifest),
    )}`;

    addTag("link", { rel: "manifest", href: manifestUrl });
    addTag("meta", { name: "theme-color", content: "#0A0F1A" });
    // iOS Safari ignores the web manifest for its home-screen icon and
    // reads this tag instead — a data: URI works here too.
    addTag("link", { rel: "apple-touch-icon", href: PWA_ICON_URL });
    // iOS Safari also ignores the manifest's "display": "standalone" —
    // without this tag, tapping the Home Screen icon just opens the
    // URL inside a normal Safari tab (with the browser chrome), not a
    // true standalone app. That matters beyond cosmetics: iOS only
    // exposes window.PushManager (and therefore Web Push) to pages
    // running in that real standalone mode, so without this tag Push
    // Notifications silently stays unsupported on iOS even after the
    // user adds the app to their Home Screen.
    addTag("meta", { name: "apple-mobile-web-app-capable", content: "yes" });
    addTag("meta", {
      name: "apple-mobile-web-app-status-bar-style",
      content: "black-translucent",
    });
    addTag("meta", {
      name: "apple-mobile-web-app-title",
      content: "Workforce",
    });

    if ("serviceWorker" in navigator) {
      navigator.serviceWorker
        .register("/sw.js")
        .catch((err) =>
          console.error("[pwa] service worker registration failed:", err),
        );
    }
  }, []);
}

export default function App() {
  usePwaSetup();
  const [lang, setLang] = useLocalStorage("hrsuite:lang", "km");
  const t = LANG[lang] || LANG.km;
  const [theme, setTheme] = useLocalStorage("hrsuite:theme", "light");
  const toggleTheme = () => setTheme(theme === "dark" ? "light" : "dark");
  return (
    <LangContext.Provider value={{ lang, t, setLang }}>
      <ThemeContext.Provider value={{ theme, setTheme, toggleTheme }}>
        <AppInner />
        <ToastHost />
      </ThemeContext.Provider>
    </LangContext.Provider>
  );
}
