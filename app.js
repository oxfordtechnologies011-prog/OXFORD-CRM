const STORE_KEY = "msc_crm_roles_v1";
const SESSION_KEY = "msc_crm_session_v1";
const CLOUD_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxW3yqS-54pkBb5cVMPpfnzzl2492595HsFG47vvUi4Tlj2zwY016Ij7ksuH9NoKc1SyQ/exec";
const CLOUD_API_KEY = "crm_oxford_2026_key";

const $ = (id) => document.getElementById(id);

let state = loadState();
let currentUser = null;
let currentPage = "dashboard";
let selectedEngineerJobId = null;
let selectedAdminEditJobId = null;
let selectedUserEditId = null;
let engineerSelectedParts = [];
let selectedPartItem = null;
let engineerFocusMode = false;
let quickWarrantyFilter = "";
let quickOpenOnly = false;
let cloudSyncReady = false;
let cloudSaveTimer = null;
let cloudSaveInFlight = false;
let cloudDirty = false;
let cloudLastHash = "";
let cloudPullInFlight = false;
let cloudPollTimer = null;

const el = {
  topbar: document.querySelector(".topbar"),
  loginPanel: $("loginPanel"),
  loginCrmName: $("loginCrmName"),
  appPanel: $("appPanel"),
  kpiCards: $("kpiCards"),
  loginForm: $("loginForm"),
  resetLoginDataBtn: $("resetLoginDataBtn"),
  openForgotPasswordBtn: $("openForgotPasswordBtn"),
  forgotPasswordDialog: $("forgotPasswordDialog"),
  forgotPasswordForm: $("forgotPasswordForm"),
  closeForgotPasswordDialog: $("closeForgotPasswordDialog"),
  openProfileBtn: $("openProfileBtn"),
  logoutBtn: $("logoutBtn"),
  welcomeText: $("welcomeText"),
  appTitle: $("appTitle"),
  appSubtitle: $("appSubtitle"),
  dashboardHeading: $("dashboardHeading"),
  receptionHeading: $("receptionHeading"),
  engineerHeading: $("engineerHeading"),
  jobsheetHeading: $("jobsheetHeading"),
  mastersHeading: $("mastersHeading"),
  pageTabs: Array.from(document.querySelectorAll("#pageTabs .tab")),
  pages: Array.from(document.querySelectorAll(".page")),
  kpiTotalJobs: $("kpiTotalJobs"),
  kpiOpenJobs: $("kpiOpenJobs"),
  kpiIw: $("kpiIw"),
  kpiOow: $("kpiOow"),
  kpiDoa: $("kpiDoa"),
  kpiRefurbish: $("kpiRefurbish"),
  kpiRepacking: $("kpiRepacking"),
  kpiPendingRequests: $("kpiPendingRequests"),
  kpiReady: $("kpiReady"),
  kpiDelivered: $("kpiDelivered"),
  kpiDateFrom: $("kpiDateFrom"),
  kpiDateTo: $("kpiDateTo"),
  kpiApplyBtn: $("kpiApplyBtn"),
  kpiClearBtn: $("kpiClearBtn"),
  statusBoards: $("statusBoards"),
  jobForm: $("jobForm"),
  jobNumber: $("jobNumber"),
  jobDate: $("jobDate"),
  jobTime: $("jobTime"),
  warrantyCategory: $("warrantyCategory"),
  brand: $("brand"),
  customerType: $("customerType"),
  productCondition: $("productCondition"),
  invoiceMode: $("invoiceMode"),
  advanceAmount: $("advanceAmount"),
  totalAmount: $("totalAmount"),
  attachments: $("attachments"),
  duplicateHint: $("duplicateHint"),
  duplicateSidePanel: $("duplicateSidePanel"),
  duplicateSideList: $("duplicateSideList"),
  closeDuplicatePanel: $("closeDuplicatePanel"),
  jobsBody: $("jobsBody"),
  jobSearch: $("jobSearch"),
  jobDateFrom: $("jobDateFrom"),
  jobDateTo: $("jobDateTo"),
  jobStatusFilter: $("jobStatusFilter"),
  applyFilterBtn: $("applyFilterBtn"),
  clearFilterBtn: $("clearFilterBtn"),
  exportJobsBtn: $("exportJobsBtn"),
  priceSearch: $("priceSearch"),
  priceApplyBtn: $("priceApplyBtn"),
  priceClearBtn: $("priceClearBtn"),
  priceExportBtn: $("priceExportBtn"),
  priceAddForm: $("priceAddForm"),
  priceImportForm: $("priceImportForm"),
  priceImportFile: $("priceImportFile"),
  downloadPriceTemplate: $("downloadPriceTemplate"),
  jobSheetTemplateUploadForm: $("jobSheetTemplateUploadForm"),
  estimateTemplateUploadForm: $("estimateTemplateUploadForm"),
  invoiceTemplateUploadForm: $("invoiceTemplateUploadForm"),
  jobSheetTemplateFile: $("jobSheetTemplateFile"),
  estimateTemplateFile: $("estimateTemplateFile"),
  invoiceTemplateFile: $("invoiceTemplateFile"),
  downloadJobSheetTemplateSample: $("downloadJobSheetTemplateSample"),
  downloadEstimateTemplateSample: $("downloadEstimateTemplateSample"),
  downloadInvoiceTemplateSample: $("downloadInvoiceTemplateSample"),
  priceBody: $("priceBody"),
  attendanceStaffFilter: $("attendanceStaffFilter"),
  attendancePeriod: $("attendancePeriod"),
  attendanceMonth: $("attendanceMonth"),
  attendanceYear: $("attendanceYear"),
  attendanceFrom: $("attendanceFrom"),
  attendanceTo: $("attendanceTo"),
  attendanceSearch: $("attendanceSearch"),
  attendanceApplyBtn: $("attendanceApplyBtn"),
  attendanceClearBtn: $("attendanceClearBtn"),
  attendanceExportBtn: $("attendanceExportBtn"),
  attendanceCheckInBtn: $("attendanceCheckInBtn"),
  attendanceCheckOutBtn: $("attendanceCheckOutBtn"),
  attendanceActionRemark: $("attendanceActionRemark"),
  attendanceTodayText: $("attendanceTodayText"),
  attendancePolicyForm: $("attendancePolicyForm"),
  attendanceRulesText: $("attendanceRulesText"),
  attendanceAdminForm: $("attendanceAdminForm"),
  attendanceAdminUser: $("attendanceAdminUser"),
  attendanceBody: $("attendanceBody"),
  attendanceRequestTitle: $("attendanceRequestTitle"),
  attendanceRequestWrap: $("attendanceRequestWrap"),
  attendanceRequestBody: $("attendanceRequestBody"),
  ledgerPeriod: $("ledgerPeriod"),
  ledgerAnchorDate: $("ledgerAnchorDate"),
  ledgerFromDate: $("ledgerFromDate"),
  ledgerToDate: $("ledgerToDate"),
  ledgerSearch: $("ledgerSearch"),
  ledgerApplyBtn: $("ledgerApplyBtn"),
  ledgerClearBtn: $("ledgerClearBtn"),
  ledgerExportBtn: $("ledgerExportBtn"),
  ledgerExpenseForm: $("ledgerExpenseForm"),
  ledgerAdminCashForm: $("ledgerAdminCashForm"),
  ledgerBody: $("ledgerBody"),
  ledgerTotalIn: $("ledgerTotalIn"),
  ledgerTotalOut: $("ledgerTotalOut"),
  ledgerBalance: $("ledgerBalance"),
  accountsSearch: $("accountsSearch"),
  accountsFrom: $("accountsFrom"),
  accountsTo: $("accountsTo"),
  accountsApply: $("accountsApply"),
  accountsClear: $("accountsClear"),
  accountsExportBtn: $("accountsExportBtn"),
  accountsBody: $("accountsBody"),
  engineerJobSelect: $("engineerJobSelect"),
  engineerJobsBody: $("engineerJobsBody"),
  loadEngineerJob: $("loadEngineerJob"),
  engineerBackBtn: $("engineerBackBtn"),
  engineerJobMeta: $("engineerJobMeta"),
  engineerUpdateForm: $("engineerUpdateForm"),
  engineerName: $("engineerName"),
  engCustName: $("engCustName"),
  engCustAddress: $("engCustAddress"),
  engCustContact: $("engCustContact"),
  engCustAlt: $("engCustAlt"),
  engCustEmail: $("engCustEmail"),
  engCustModel: $("engCustModel"),
  engCustImei: $("engCustImei"),
  engCustComplaint: $("engCustComplaint"),
  repairStatus: $("repairStatus"),
  engineerWarranty: $("engineerWarranty"),
  serviceCharge: $("serviceCharge"),
  paymentAmount: $("paymentAmount"),
  discountAmount: $("discountAmount"),
  autoCollectionAmount: $("autoCollectionAmount"),
  closeCollectionMode: $("closeCollectionMode"),
  handoverRemark: $("handoverRemark"),
  printEstimateBtn: $("printEstimateBtn"),
  partSearch: $("partSearch"),
  partSuggestionBox: $("partSuggestionBox"),
  partQtyInput: $("partQtyInput"),
  partsDatalist: $("partsDatalist"),
  addPartBtn: $("addPartBtn"),
  selectedParts: $("selectedParts"),
  engineerPhotos: $("engineerPhotos"),
  engineerZip: $("engineerZip"),
  invoicePatternForm: $("invoicePatternForm"),
  invoicePattern: $("invoicePattern"),
  headingForm: $("headingForm"),
  jobsheetTemplateForm: $("jobsheetTemplateForm"),
  printStyleForm: $("printStyleForm"),
  printTextForm: $("printTextForm"),
  importJobsForm: $("importJobsForm"),
  importJobsFile: $("importJobsFile"),
  importMode: $("importMode"),
  downloadImportSample: $("downloadImportSample"),
  openTimelineBtn: $("openTimelineBtn"),
  openEngineerLogBtn: $("openEngineerLogBtn"),
  historyDialog: $("historyDialog"),
  historyTitle: $("historyTitle"),
  historyList: $("historyList"),
  closeHistoryDialog: $("closeHistoryDialog"),
  profileDialog: $("profileDialog"),
  closeProfileDialog: $("closeProfileDialog"),
  togglePasswordPanelBtn: $("togglePasswordPanelBtn"),
  savePasswordBtn: $("savePasswordBtn"),
  profileAvatar: $("profileAvatar"),
  profileNameHeading: $("profileNameHeading"),
  profileSubline: $("profileSubline"),
  profileName: $("profileName"),
  profileTitle: $("profileTitle"),
  profileUsername: $("profileUsername"),
  profileRole: $("profileRole"),
  profileJoining: $("profileJoining"),
  profileAge: $("profileAge"),
  profilePasswordForm: $("profilePasswordForm"),
  adminEditDialog: $("adminEditDialog"),
  adminEditForm: $("adminEditForm"),
  closeAdminEdit: $("closeAdminEdit"),
  adminWarranty: $("adminWarranty"),
  adminCustomerType: $("adminCustomerType"),
  usersBody: $("usersBody"),
  userForm: $("userForm"),
  userEditDialog: $("userEditDialog"),
  userEditForm: $("userEditForm"),
  closeUserEditDialog: $("closeUserEditDialog"),
  approvalsBody: $("approvalsBody"),
  masterForms: Array.from(document.querySelectorAll(".masterForm"))
};

function now() {
  return new Date();
}

function nowIso() {
  return now().toISOString();
}

function ymd(dateLike = null) {
  const d = dateLike ? new Date(dateLike) : now();
  return d.toISOString().slice(0, 10);
}

function hm(dateLike = null) {
  const d = dateLike ? new Date(dateLike) : now();
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

function money(value) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 2
  }).format(Number(value || 0));
}

function uid(prefix) {
  return `${prefix}_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
}

function touchJob(job, at = nowIso()) {
  if (!job || typeof job !== "object") return;
  job.updatedAt = at;
}

function defaultState() {
  const today = new Date().toISOString().slice(0, 10);
  return {
    users: [
      { id: uid("USR"), fullName: "Administrator", profileTitle: "System Admin", username: "admin", password: "admin123", role: "admin", age: "", joiningDate: today, active: true },
      { id: uid("USR"), fullName: "Reception Staff", profileTitle: "Front Desk", username: "receptionist", password: "recep123", role: "receptionist", age: "", joiningDate: today, active: true },
      { id: uid("USR"), fullName: "Engineer Staff", profileTitle: "Service Engineer", username: "engineer", password: "eng123", role: "engineer", age: "", joiningDate: today, active: true }
    ],
    masters: {
      invoicePattern: "JOB-{YYYY}-{####}",
      warrantyCategories: ["IW", "OOW", "DOA", "Repacking", "Refurbish"],
      brands: ["Realme", "Karbon"],
      customerTypes: ["Customer", "Dealer"],
      productConditions: ["Good", "Scratch", "Display Broken", "Body Damage", "Water Mark"],
      invoiceModes: ["WhatsApp", "Mail", "In Hand"],
      engineers: ["Engineer 1"],
      repairStatuses: ["Received", "Diagnosing", "Waiting Parts", "Repairing", "Ready Pickup", "Delivered", "Cancelled"],
      parts: ["Display", "Battery", "Charging Port"]
    },
    settings: {
      appTitle: "Mobile Service Center CRM",
      loginCrmName: "Mobile Service Center CRM",
      appSubtitle: "Admin controlled CRM for Reception and Engineer workflow.",
      dashboardHeading: "Status Boards",
      receptionHeading: "Reception Job Registration",
      engineerHeading: "Engineer Desk",
      jobsheetHeading: "Job Sheets",
      mastersHeading: "Admin Masters",
      jobSheetTemplate: {
        companyName: "Mobile Service Center",
        phone: "",
        email: "",
        address: "",
        headerNote: "",
        footerNote: "Thanks for choosing us."
      },
      printStyle: {
        fontFamily: "Arial",
        baseSize: 12,
        headingSize: 22
      },
      printText: {
        jobSheetTitle: "Service Center Handover Report",
        jobTermsHeading: "Please read the following terms carefully before signing. After signing, you accept all terms.",
        jobTerms: "",
        estimateTitle: "Estimate",
        estimateIntro: "This is with reference to Job no {JOB_NO}. The handset on inspection found with complaint: {COMPLAINT}",
        estimateNote: "Please note: Actual cost may vary at the time of repair.",
        estimateTerms: "",
        taxInvoiceTitle: "Tax Invoice",
        taxTerms: "Out-of-warranty parts replaced carry limited service warranty as applicable. This is a computer generated invoice.",
        invoiceTerms: "",
        signatureCustomerLabel: "Customer signature",
        signatureStaffLabel: "Signature by Service staff"
      },
      attendancePolicy: {
        checkInStart: "08:00",
        checkInEnd: "10:30",
        checkOutStart: "16:00",
        checkOutEnd: "23:30"
      }
    },
    jobCounter: 0,
    invoiceCounter: 0,
    priceList: [],
    attendance: [],
    attendanceRequests: [],
    cashLedger: [],
    jobs: [],
    requests: []
  };
}

function migrateShape(raw) {
  const base = defaultState();
  const merged = {
    ...base,
    ...raw,
    masters: {
      ...base.masters,
      ...(raw?.masters || {})
    },
    settings: {
      ...base.settings,
      ...(raw?.settings || {})
    }
  };
  if (!Array.isArray(merged.users)) merged.users = base.users;
  merged.users = merged.users.filter((u) => u && typeof u === "object" && String(u.username || "").trim());
  if (!merged.users.length) merged.users = base.users;
  if (!Array.isArray(merged.jobs)) merged.jobs = [];
  if (!Array.isArray(merged.requests)) merged.requests = [];
  if (!Array.isArray(merged.priceList)) merged.priceList = [];
  if (!Array.isArray(merged.attendance)) merged.attendance = [];
  if (!Array.isArray(merged.attendanceRequests)) merged.attendanceRequests = [];
  if (!Array.isArray(merged.cashLedger)) merged.cashLedger = [];

  merged.users = merged.users.map((u) => ({
    ...u,
    fullName: String(u.fullName || u.username || "").trim(),
    profileTitle: String(u.profileTitle || "").trim(),
    age: u.age === undefined || u.age === null ? "" : String(u.age),
    joiningDate: String(u.joiningDate || "").trim() || new Date().toISOString().slice(0, 10)
  }));

  merged.priceList = merged.priceList
    .filter((p) => p && typeof p === "object")
    .map((p) => ({
      id: p.id || uid("PRC"),
      itemCode: String(p.itemCode || "").trim(),
      partName: String(p.partName || p.name || "").trim(),
      brand: String(p.brand || "").trim(),
      model: String(p.model || "").trim(),
      price: Number(p.price || 0),
      updatedAt: p.updatedAt || nowIso(),
      updatedBy: p.updatedBy || "system"
    }))
    .filter((p) => p.partName);

  merged.cashLedger = merged.cashLedger
    .filter((r) => r && typeof r === "object")
    .map((r) => ({
      id: r.id || uid("LED"),
      at: r.at || nowIso(),
      type: r.type === "OUT" ? "OUT" : "IN",
      category: String(r.category || "General").trim(),
      amount: Number(r.amount || 0),
      mode: String(r.mode || "Cash").trim(),
      note: String(r.note || "").trim(),
      by: String(r.by || "system").trim(),
      jobId: r.jobId || null
    }))
    .filter((r) => Number.isFinite(r.amount) && r.amount >= 0);

  merged.attendance = merged.attendance
    .filter((a) => a && typeof a === "object")
    .map((a) => ({
      id: a.id || uid("ATT"),
      userId: String(a.userId || ""),
      date: String(a.date || "").slice(0, 10),
      checkIn: String(a.checkIn || ""),
      checkOut: String(a.checkOut || ""),
      status: String(a.status || "Present"),
      note: String(a.note || "").trim(),
      by: String(a.by || "system"),
      updatedAt: a.updatedAt || nowIso()
    }))
    .filter((a) => a.userId && a.date);

  merged.attendanceRequests = merged.attendanceRequests
    .filter((r) => r && typeof r === "object")
    .map((r) => ({
      id: r.id || uid("ATR"),
      userId: String(r.userId || ""),
      date: String(r.date || "").slice(0, 10),
      action: String(r.action || "checkIn"),
      requestedTime: String(r.requestedTime || ""),
      remark: String(r.remark || "").trim(),
      status: ["Pending", "Approved", "Rejected"].includes(String(r.status || "")) ? String(r.status) : "Pending",
      requestedAt: r.requestedAt || nowIso(),
      reviewedAt: r.reviewedAt || "",
      reviewedBy: String(r.reviewedBy || ""),
      adminRemark: String(r.adminRemark || "").trim()
    }))
    .filter((r) => r.userId && r.date && ["checkIn", "checkOut"].includes(r.action));

  const basePolicy = base.settings.attendancePolicy;
  merged.settings.attendancePolicy = {
    checkInStart: String(merged.settings?.attendancePolicy?.checkInStart || basePolicy.checkInStart),
    checkInEnd: String(merged.settings?.attendancePolicy?.checkInEnd || basePolicy.checkInEnd),
    checkOutStart: String(merged.settings?.attendancePolicy?.checkOutStart || basePolicy.checkOutStart),
    checkOutEnd: String(merged.settings?.attendancePolicy?.checkOutEnd || basePolicy.checkOutEnd)
  };

  const ensureUser = (username, password, role) => {
    const found = merged.users.find((u) => String(u.username || "").toLowerCase() === username.toLowerCase());
    if (!found) {
      merged.users.push({
        id: uid("USR"),
        fullName: username.charAt(0).toUpperCase() + username.slice(1),
        profileTitle: "",
        username,
        password,
        role,
        age: "",
        joiningDate: new Date().toISOString().slice(0, 10),
        active: true
      });
      return;
    }
    if (!found.password) found.password = password;
    if (!found.role) found.role = role;
    if (!found.fullName) found.fullName = username.charAt(0).toUpperCase() + username.slice(1);
    if (found.profileTitle === undefined || found.profileTitle === null) found.profileTitle = "";
    if (!found.joiningDate) found.joiningDate = new Date().toISOString().slice(0, 10);
    if (found.age === undefined || found.age === null) found.age = "";
    if (found.active === false && username === "admin") found.active = true;
  };

  ensureUser("admin", "admin123", "admin");
  ensureUser("receptionist", "recep123", "receptionist");
  ensureUser("engineer", "eng123", "engineer");

  merged.jobs = merged.jobs.map((j) => ({
    ...j,
    createdAt: j.createdAt || j.updatedAt || nowIso(),
    updatedAt: j.updatedAt || j.createdAt || nowIso(),
    customerType: String(j.customerType || "Customer"),
    serviceCharge: Number(j.serviceCharge || 0),
    discountAmount: Number(j.discountAmount || 0),
    attachments: Array.isArray(j.attachments) ? j.attachments : [],
    engineerUpdates: Array.isArray(j.engineerUpdates) ? j.engineerUpdates : [],
    payments: Array.isArray(j.payments) ? j.payments : [],
    timeline: Array.isArray(j.timeline) ? j.timeline : []
  }));

  merged.jobs.forEach((j) => {
    const amt = Number(j?.closeCollection?.amount || 0);
    if (!j.closeCollection || j.status !== "Delivered" || j.warrantyCategory !== "OOW" || amt <= 0) return;
    const existing = j.closeCollection.ledgerId ? merged.cashLedger.find((r) => r.id === j.closeCollection.ledgerId) : null;
    if (existing) return;
    const id = uid("LED");
    merged.cashLedger.push({
      id,
      at: j.closeCollection.at || j.createdAt || nowIso(),
      type: "IN",
      category: "OOW Collection",
      amount: amt,
      mode: String(j.closeCollection.mode || "Cash"),
      note: `Collection for ${j.jobNumber || "job"}`,
      by: String(j.closeCollection.by || "system"),
      jobId: j.id || null
    });
    j.closeCollection.ledgerId = id;
  });

  return merged;
}

function loadState() {
  const raw = localStorage.getItem(STORE_KEY);
  if (!raw) return defaultState();
  try {
    return migrateShape(JSON.parse(raw));
  } catch (_err) {
    return defaultState();
  }
}

function saveState() {
  localStorage.setItem(STORE_KEY, JSON.stringify(state));
}

function saveSession() {
  if (!currentUser?.id) {
    localStorage.removeItem(SESSION_KEY);
    return;
  }
  const u = getUserById(currentUser.id);
  const payload = {
    id: currentUser.id,
    username: String(currentUser.username || u?.username || "").trim().toLowerCase()
  };
  localStorage.setItem(SESSION_KEY, JSON.stringify(payload));
}

function loadSession() {
  const raw = localStorage.getItem(SESSION_KEY);
  if (!raw) return null;

  let savedId = "";
  let savedUsername = "";
  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === "object") {
      savedId = String(parsed.id || "").trim();
      savedUsername = String(parsed.username || "").trim().toLowerCase();
    }
  } catch (_err) {
    // Backward compatibility: old format stored plain user id.
    savedId = String(raw || "").trim();
  }

  let user = savedId ? state.users.find((u) => u.id === savedId && u.active) : null;
  if (!user && savedUsername) {
    user = state.users.find((u) => String(u.username || "").trim().toLowerCase() === savedUsername && u.active) || null;
  }
  return user ? { id: user.id, username: String(user.username || "").trim().toLowerCase() } : null;
}

function cloudSyncEnabled() {
  return Boolean(
    CLOUD_WEB_APP_URL &&
    CLOUD_API_KEY &&
    CLOUD_API_KEY !== "REPLACE_WITH_API_KEY" &&
    /^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/i.test(CLOUD_WEB_APP_URL)
  );
}

function buildCloudQuery(params) {
  const usp = new URLSearchParams();
  Object.keys(params || {}).forEach((k) => {
    const v = params[k];
    if (v === undefined || v === null) return;
    usp.set(k, String(v));
  });
  return `${CLOUD_WEB_APP_URL}?${usp.toString()}`;
}

function mergeUniqueStrings(...lists) {
  return Array.from(
    new Set(
      lists
        .flat()
        .map((v) => String(v || "").trim())
        .filter(Boolean)
    )
  );
}

function parseRecordTime(value) {
  if (!value) return 0;
  const time = new Date(value).getTime();
  return Number.isFinite(time) ? time : 0;
}

function getRecordVersion(item) {
  if (!item || typeof item !== "object") return 0;
  return Math.max(
    parseRecordTime(item.updatedAt),
    parseRecordTime(item.modifiedAt),
    parseRecordTime(item.lastUpdatedAt),
    parseRecordTime(item.createdAt),
    parseRecordTime(item.at),
    parseRecordTime(item.requestedAt),
    parseRecordTime(item.reviewedAt)
  );
}

function mergeRecordsByKey(baseArr, overlayArr, keyFn) {
  const map = new Map();
  const noKey = [];
  const noKeySeen = new Set();

  const pushNoKey = (item) => {
    const sig = JSON.stringify(item);
    if (noKeySeen.has(sig)) return;
    noKeySeen.add(sig);
    noKey.push(item);
  };

  const upsert = (item) => {
    if (!item || typeof item !== "object") return;
    const key = String(keyFn(item) || "").trim();
    if (!key) {
      pushNoKey(item);
      return;
    }
    const prev = map.get(key);
    if (!prev) {
      map.set(key, { ...item });
      return;
    }

    const prevVer = getRecordVersion(prev);
    const nextVer = getRecordVersion(item);
    if (nextVer > prevVer) {
      map.set(key, { ...prev, ...item });
    } else if (nextVer < prevVer) {
      map.set(key, { ...item, ...prev });
    } else {
      map.set(key, { ...prev, ...item });
    }
  };

  (baseArr || []).forEach(upsert);
  (overlayArr || []).forEach(upsert);
  return [...map.values(), ...noKey];
}

function mergeCloudStates(remoteRaw, localRaw) {
  const remote = migrateShape(remoteRaw || {});
  const local = migrateShape(localRaw || {});

  const merged = {
    ...remote,
    ...local,
    jobCounter: Math.max(Number(remote.jobCounter || 0), Number(local.jobCounter || 0)),
    invoiceCounter: Math.max(Number(remote.invoiceCounter || 0), Number(local.invoiceCounter || 0)),
    users: mergeRecordsByKey(remote.users, local.users, (u) => u.id || `u:${String(u.username || "").toLowerCase()}`),
    jobs: mergeRecordsByKey(remote.jobs, local.jobs, (j) => j.id || j.jobNumber),
    priceList: mergeRecordsByKey(remote.priceList, local.priceList, (p) => p.id || `${p.itemCode}|${p.partName}|${p.model}`),
    attendance: mergeRecordsByKey(remote.attendance, local.attendance, (a) => a.id || `${a.userId}|${a.date}`),
    attendanceRequests: mergeRecordsByKey(remote.attendanceRequests, local.attendanceRequests, (r) => r.id || `${r.userId}|${r.date}|${r.action}`),
    cashLedger: mergeRecordsByKey(remote.cashLedger, local.cashLedger, (r) => r.id || `${r.jobId}|${r.at}|${r.amount}`),
    requests: mergeRecordsByKey(remote.requests, local.requests, (r) => r.id || `${r.jobId}|${r.field}|${r.createdAt}`),
    masters: {
      ...remote.masters,
      ...local.masters,
      warrantyCategories: mergeUniqueStrings(remote.masters?.warrantyCategories || [], local.masters?.warrantyCategories || []),
      brands: mergeUniqueStrings(remote.masters?.brands || [], local.masters?.brands || []),
      customerTypes: mergeUniqueStrings(remote.masters?.customerTypes || [], local.masters?.customerTypes || []),
      productConditions: mergeUniqueStrings(remote.masters?.productConditions || [], local.masters?.productConditions || []),
      invoiceModes: mergeUniqueStrings(remote.masters?.invoiceModes || [], local.masters?.invoiceModes || []),
      engineers: mergeUniqueStrings(remote.masters?.engineers || [], local.masters?.engineers || []),
      repairStatuses: mergeUniqueStrings(remote.masters?.repairStatuses || [], local.masters?.repairStatuses || []),
      parts: mergeUniqueStrings(remote.masters?.parts || [], local.masters?.parts || [])
    },
    settings: {
      ...remote.settings,
      ...local.settings,
      attendancePolicy: {
        ...(remote.settings?.attendancePolicy || {}),
        ...(local.settings?.attendancePolicy || {})
      },
      jobSheetTemplate: {
        ...(remote.settings?.jobSheetTemplate || {}),
        ...(local.settings?.jobSheetTemplate || {})
      },
      printStyle: {
        ...(remote.settings?.printStyle || {}),
        ...(local.settings?.printStyle || {})
      },
      printText: {
        ...(remote.settings?.printText || {}),
        ...(local.settings?.printText || {})
      }
    }
  };

  return migrateShape(merged);
}

function cloudGetJsonp(action) {
  return new Promise((resolve, reject) => {
    const cb = `crmCloudCb_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement("script");
    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error("Cloud JSONP timeout"));
    }, 12000);

    function cleanup() {
      clearTimeout(timeout);
      if (script.parentNode) script.parentNode.removeChild(script);
      try {
        delete window[cb];
      } catch (_err) {
        window[cb] = undefined;
      }
    }

    window[cb] = (payload) => {
      cleanup();
      resolve(payload);
    };

    script.onerror = () => {
      cleanup();
      reject(new Error("Cloud JSONP request failed"));
    };

    script.src = buildCloudQuery({
      action,
      apiKey: CLOUD_API_KEY,
      callback: cb,
      _: Date.now()
    });
    document.head.appendChild(script);
  });
}

async function cloudFetchState() {
  let json;
  try {
    // Apps Script web apps often block cross-origin fetch JSON reads.
    // JSONP avoids CORS read restrictions for GET actions.
    json = await cloudGetJsonp("getState");
  } catch (_jsonpErr) {
    const url = buildCloudQuery({
      action: "getState",
      apiKey: CLOUD_API_KEY,
      _: Date.now()
    });
    const res = await fetch(url, { method: "GET", cache: "no-store" });
    if (!res.ok) throw new Error(`Cloud getState failed (${res.status})`);
    json = await res.json();
  }
  if (!json?.ok) throw new Error(String(json?.error || "Cloud getState error"));
  return json.state && typeof json.state === "object" ? json.state : {};
}

async function cloudSaveStateNow() {
  let mergedForSave = state;
  try {
    // Merge with latest remote first to prevent last-write-wins data loss
    // when multiple systems save around the same time.
    const remote = await cloudFetchState();
    mergedForSave = mergeCloudStates(remote, state);
  } catch (err) {
    console.warn("Cloud pre-merge read failed:", err);
  }

  const payload = {
    action: "saveState",
    apiKey: CLOUD_API_KEY,
    state: mergedForSave
  };
  try {
    const res = await fetch(CLOUD_WEB_APP_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload)
    });
    if (!res.ok) throw new Error(`Cloud saveState failed (${res.status})`);
    const json = await res.json();
    if (!json?.ok) throw new Error(String(json?.error || "Cloud saveState error"));
    state = migrateShape(mergedForSave);
    saveState();
    cloudLastHash = JSON.stringify(state);
    return;
  } catch (_corsErr) {
    // Fallback: send without requiring CORS-readable response.
    await fetch(CLOUD_WEB_APP_URL, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(payload)
    });
  }
  state = migrateShape(mergedForSave);
  saveState();
  cloudLastHash = JSON.stringify(state);
}

function scheduleCloudSave() {
  if (!cloudSyncReady || !cloudSyncEnabled()) return;
  const nextHash = JSON.stringify(state);
  if (nextHash === cloudLastHash) return;
  cloudDirty = true;
  if (cloudSaveTimer) clearTimeout(cloudSaveTimer);
  cloudSaveTimer = setTimeout(() => {
    flushCloudSave();
  }, 350);
}

async function flushCloudSave() {
  if (!cloudSyncReady || !cloudSyncEnabled()) return;
  if (cloudSaveInFlight || !cloudDirty) return;
  cloudSaveInFlight = true;
  cloudDirty = false;
  try {
    await cloudSaveStateNow();
  } catch (err) {
    cloudDirty = true;
    console.warn("Cloud save failed:", err);
  } finally {
    cloudSaveInFlight = false;
    if (cloudDirty) {
      if (cloudSaveTimer) clearTimeout(cloudSaveTimer);
      cloudSaveTimer = setTimeout(() => {
        flushCloudSave();
      }, 700);
    }
  }
}

async function initializeCloudSync() {
  if (!cloudSyncEnabled()) {
    console.warn("Cloud sync disabled. Set CLOUD_API_KEY in app.js to enable shared data.");
    return;
  }
  try {
    const remote = await cloudFetchState();
    if (remote && Object.keys(remote).length) {
      const remoteState = migrateShape(remote);
      const merged = mergeCloudStates(remoteState, state);
      state = merged;
      saveState();
      const remoteHash = JSON.stringify(remoteState);
      const mergedHash = JSON.stringify(merged);
      cloudLastHash = remoteHash;
      if (remoteHash !== mergedHash) scheduleCloudSave();
    } else {
      cloudLastHash = JSON.stringify(state);
      cloudSyncReady = true;
      await cloudSaveStateNow();
      return;
    }
    cloudSyncReady = true;
  } catch (err) {
    cloudSyncReady = true;
    console.warn("Cloud sync init failed:", err);
  }
}

function syncCurrentUserReference() {
  if (!currentUser) return;
  let user = getUserById(currentUser.id);
  if (user && user.active) {
    currentUser.username = String(user.username || "").trim().toLowerCase();
    saveSession();
    return;
  }

  const preferredUsername = String(
    currentUser.username ||
      (() => {
        try {
          const raw = localStorage.getItem(SESSION_KEY);
          if (!raw) return "";
          const parsed = JSON.parse(raw);
          return parsed?.username || "";
        } catch (_err) {
          return "";
        }
      })()
  )
    .trim()
    .toLowerCase();

  if (preferredUsername) {
    user = state.users.find((u) => String(u.username || "").trim().toLowerCase() === preferredUsername && u.active) || null;
    if (user) {
      currentUser = { id: user.id, username: preferredUsername };
      saveSession();
      return;
    }
  }

  currentUser = null;
  saveSession();
}

function renderUi(skipRegisterReset = false) {
  renderHeadings();
  populateMasterDropdowns();
  if (!skipRegisterReset) resetRegisterHeader();
  updateConditionalFields();
  renderKpis();
  renderStatusBoards();
  renderJobsTable();
  renderPriceList();
  renderAttendance();
  renderLedger();
  renderAccounts();
  renderEngineerJobSelect();
  renderEngineerJobsList();
  renderEngineerMeta();
  renderMasterLists();
  renderUsers();
  renderApprovals();
}

async function pullCloudUpdates(force = false) {
  if (!cloudSyncReady || !cloudSyncEnabled()) return;
  if (cloudPullInFlight || cloudSaveInFlight || cloudDirty) return;
  if (!force && !currentUser) return;

  cloudPullInFlight = true;
  try {
    const remote = await cloudFetchState();
    const remoteState = migrateShape(remote || {});
    const localState = state;
    const merged = mergeCloudStates(remoteState, localState);
    const mergedHash = JSON.stringify(merged);
    const localHash = JSON.stringify(localState);
    const remoteHash = JSON.stringify(remoteState);

    // Keep remote hash so unsynced local changes still trigger save.
    cloudLastHash = remoteHash;

    if (mergedHash === localHash) {
      if (remoteHash !== localHash) scheduleCloudSave();
      return;
    }

    state = merged;
    saveState();
    syncCurrentUserReference();
    renderUi(true);
    updateRoleUi();
    if (remoteHash !== mergedHash) scheduleCloudSave();
  } catch (err) {
    console.warn("Cloud pull failed:", err);
  } finally {
    cloudPullInFlight = false;
  }
}

function startCloudPolling() {
  if (cloudPollTimer) clearInterval(cloudPollTimer);
  if (!cloudSyncEnabled()) return;
  cloudPollTimer = setInterval(() => {
    if (document.hidden) return;
    pullCloudUpdates();
  }, 2000);
}

function formatJobNumber(counter) {
  const d = now();
  const tokenMap = {
    "{YYYY}": String(d.getFullYear()),
    "{MM}": String(d.getMonth() + 1).padStart(2, "0"),
    "{DD}": String(d.getDate()).padStart(2, "0"),
    "{####}": String(counter).padStart(4, "0")
  };

  let out = state.masters.invoicePattern || "JOB-{YYYY}-{####}";
  Object.keys(tokenMap).forEach((token) => {
    out = out.replaceAll(token, tokenMap[token]);
  });
  return out;
}

function nextJobNumber(seedCounter) {
  let c = Number(seedCounter || 0) + 1;
  let n = formatJobNumber(c);
  const existing = new Set((state.jobs || []).map((j) => j.jobNumber));
  while (existing.has(n)) {
    c += 1;
    n = formatJobNumber(c);
  }
  return { counter: c, jobNumber: n };
}

function resetRegisterHeader() {
  const d = now();
  el.jobNumber.value = nextJobNumber(state.jobCounter).jobNumber;
  el.jobDate.value = d.toLocaleDateString();
  el.jobTime.value = d.toLocaleTimeString();
}

function getUserById(id) {
  return state.users.find((u) => u.id === id);
}

function getCurrentUser() {
  return currentUser ? getUserById(currentUser.id) : null;
}

function findUserByUsername(username) {
  const key = String(username || "").trim().toLowerCase();
  if (!key) return null;
  return state.users.find((u) => String(u.username || "").trim().toLowerCase() === key) || null;
}

function renderProfileCard() {
  const user = getCurrentUser();
  if (!user) return;
  const initials = String(user.fullName || user.username || "U")
    .split(" ")
    .filter(Boolean)
    .slice(0, 2)
    .map((s) => s.charAt(0).toUpperCase())
    .join("") || "U";
  if (el.profileAvatar) el.profileAvatar.textContent = initials;
  if (el.profileNameHeading) el.profileNameHeading.textContent = user.fullName || user.username || "Staff Profile";
  if (el.profileSubline) el.profileSubline.textContent = `${user.profileTitle || "Staff Account"} | ${user.role || "-"}`;
  if (el.profileName) el.profileName.textContent = user.fullName || user.username || "-";
  if (el.profileTitle) el.profileTitle.textContent = user.profileTitle || "-";
  if (el.profileUsername) el.profileUsername.textContent = user.username || "-";
  if (el.profileRole) el.profileRole.textContent = user.role || "-";
  if (el.profileJoining) el.profileJoining.textContent = user.joiningDate || "-";
  if (el.profileAge) el.profileAge.textContent = user.age || "-";
}

function roleAllowed(tabRoleAttr, role) {
  if (!tabRoleAttr) return true;
  const list = tabRoleAttr.split(",").map((x) => x.trim());
  return list.includes(role);
}

function updateRoleUi() {
  const user = getCurrentUser();
  if (!user) {
    if (el.topbar) el.topbar.classList.add("hidden");
    el.loginPanel.classList.remove("hidden");
    el.appPanel.classList.add("hidden");
    el.kpiCards.classList.add("hidden");
    el.pages.forEach((p) => p.classList.add("hidden"));
    el.profileDialog?.close();
    return;
  }

  if (el.topbar) el.topbar.classList.remove("hidden");
  el.loginPanel.classList.add("hidden");
  el.appPanel.classList.remove("hidden");
  el.kpiCards.classList.remove("hidden");
  el.welcomeText.textContent = `${user.fullName || user.username} (${user.role})`;

  el.pageTabs.forEach((tab) => {
    const ok = roleAllowed(tab.dataset.role || "", user.role);
    tab.classList.toggle("hidden", !ok);
    if (!ok && tab.dataset.page === currentPage) currentPage = "dashboard";
  });
  switchPage(currentPage);
  applyEngineerFocusMode();
  renderProfileCard();
}

function switchPage(page) {
  if (page !== "engineer") engineerFocusMode = false;
  if (page !== "register") closeDuplicatePanel();
  currentPage = page;
  el.pageTabs.forEach((t) => t.classList.toggle("active", t.dataset.page === page));
  el.pages.forEach((p) => {
    const isTarget = p.id === `page-${page}`;
    p.classList.toggle("active", isTarget);
    p.classList.toggle("hidden", !isTarget);
  });
  applyEngineerFocusMode();
  if (page === "attendance") renderAttendance();
}

function applyEngineerFocusMode() {
  const inFocus = engineerFocusMode && currentPage === "engineer";
  el.appPanel.classList.toggle("hidden", inFocus);
  el.kpiCards.classList.toggle("hidden", inFocus);
  const engPage = document.getElementById("page-engineer");
  if (engPage) engPage.classList.toggle("focus-page", inFocus);
}

function populateMasterDropdowns() {
  fillSelect(el.warrantyCategory, state.masters.warrantyCategories);
  fillSelect(el.brand, state.masters.brands);
  fillSelect(el.customerType, state.masters.customerTypes);
  fillSelect(el.productCondition, state.masters.productConditions);
  fillSelectWithPlaceholder(el.invoiceMode, state.masters.invoiceModes, "Invoice Mode");
  fillSelect(el.engineerName, state.masters.engineers);
  fillSelect(el.repairStatus, state.masters.repairStatuses);
  fillSelect(el.engineerWarranty, state.masters.warrantyCategories);
  if (el.adminWarranty) fillSelect(el.adminWarranty, state.masters.warrantyCategories);
  if (el.adminCustomerType) fillSelect(el.adminCustomerType, state.masters.customerTypes);
  fillSelectWithPlaceholder(el.jobStatusFilter, state.masters.repairStatuses, "All Status");
  const partOptions = [];
  (state.masters.parts || []).forEach((p) => partOptions.push(p));
  (state.priceList || []).forEach((p) => {
    if (p.partName) partOptions.push(`${p.itemCode ? `${p.itemCode} - ` : ""}${p.partName}`);
    if (p.itemCode) partOptions.push(p.itemCode);
    if (p.model) partOptions.push(`${p.model} - ${p.partName}`);
  });
  const uniquePartOptions = Array.from(new Set(partOptions.filter(Boolean)));
  el.partsDatalist.innerHTML = uniquePartOptions.map((p) => `<option value="${escapeHtml(p)}"></option>`).join("");

  el.invoicePattern.value = state.masters.invoicePattern;
  if (el.headingForm) {
    el.headingForm.appTitle.value = state.settings.appTitle;
    if (el.headingForm.loginCrmName) {
      el.headingForm.loginCrmName.value = state.settings.loginCrmName || state.settings.appTitle || "Mobile Service Center CRM";
    }
    el.headingForm.appSubtitle.value = state.settings.appSubtitle;
    el.headingForm.dashboardHeading.value = state.settings.dashboardHeading;
    el.headingForm.receptionHeading.value = state.settings.receptionHeading;
    el.headingForm.engineerHeading.value = state.settings.engineerHeading;
    el.headingForm.jobsheetHeading.value = state.settings.jobsheetHeading;
    el.headingForm.mastersHeading.value = state.settings.mastersHeading;
  }
  if (el.jobsheetTemplateForm) {
    const t = getJobSheetTemplate();
    el.jobsheetTemplateForm.companyName.value = t.companyName;
    el.jobsheetTemplateForm.phone.value = t.phone;
    el.jobsheetTemplateForm.email.value = t.email;
    el.jobsheetTemplateForm.address.value = t.address;
    el.jobsheetTemplateForm.headerNote.value = t.headerNote;
    el.jobsheetTemplateForm.footerNote.value = t.footerNote;
  }
  if (el.printStyleForm) {
    const ps = getPrintStyle();
    el.printStyleForm.fontFamily.value = ps.fontFamily;
    el.printStyleForm.baseSize.value = ps.baseSize;
    el.printStyleForm.headingSize.value = ps.headingSize;
  }
  if (el.printTextForm) {
    const p = getPrintText();
    el.printTextForm.jobSheetTitle.value = p.jobSheetTitle;
    el.printTextForm.jobTermsHeading.value = p.jobTermsHeading;
    if (el.printTextForm.jobTerms) el.printTextForm.jobTerms.value = p.jobTerms;
    el.printTextForm.estimateTitle.value = p.estimateTitle;
    el.printTextForm.estimateIntro.value = p.estimateIntro;
    el.printTextForm.estimateNote.value = p.estimateNote;
    if (el.printTextForm.estimateTerms) el.printTextForm.estimateTerms.value = p.estimateTerms;
    el.printTextForm.taxInvoiceTitle.value = p.taxInvoiceTitle;
    el.printTextForm.taxTerms.value = p.taxTerms;
    if (el.printTextForm.invoiceTerms) el.printTextForm.invoiceTerms.value = p.invoiceTerms;
    el.printTextForm.signatureCustomerLabel.value = p.signatureCustomerLabel;
    el.printTextForm.signatureStaffLabel.value = p.signatureStaffLabel;
  }
}

function fillSelect(select, values) {
  select.innerHTML = values.map((v) => `<option>${escapeHtml(v)}</option>`).join("");
}

function fillSelectWithPlaceholder(select, values, placeholder) {
  select.innerHTML = `<option value="">${escapeHtml(placeholder)}</option>${values
    .map((v) => `<option>${escapeHtml(v)}</option>`)
    .join("")}`;
}

function renderHeadings() {
  const s = state.settings;
  el.appTitle.textContent = s.appTitle;
  if (el.loginCrmName) el.loginCrmName.textContent = s.loginCrmName || s.appTitle || "Mobile Service Center CRM";
  el.appSubtitle.textContent = s.appSubtitle;
  el.dashboardHeading.textContent = s.dashboardHeading;
  el.receptionHeading.textContent = s.receptionHeading;
  el.engineerHeading.textContent = s.engineerHeading;
  el.jobsheetHeading.textContent = s.jobsheetHeading;
  el.mastersHeading.textContent = s.mastersHeading;
}

function getJobSheetTemplate() {
  return {
    companyName: state.settings?.jobSheetTemplate?.companyName || "Mobile Service Center",
    phone: state.settings?.jobSheetTemplate?.phone || "",
    email: state.settings?.jobSheetTemplate?.email || "",
    address: state.settings?.jobSheetTemplate?.address || "",
    headerNote: state.settings?.jobSheetTemplate?.headerNote || "",
    footerNote: state.settings?.jobSheetTemplate?.footerNote || ""
  };
}

function getPrintStyle() {
  return {
    fontFamily: state.settings?.printStyle?.fontFamily || "Arial",
    baseSize: Number(state.settings?.printStyle?.baseSize || 12),
    headingSize: Number(state.settings?.printStyle?.headingSize || 22)
  };
}

function getPrintText() {
  const t = state.settings?.printText || {};
  return {
    jobSheetTitle: String(t.jobSheetTitle || "Service Center Handover Report"),
    jobTermsHeading: String(t.jobTermsHeading || "Please read the following terms carefully before signing. After signing, you accept all terms."),
    jobTerms: String(t.jobTerms || ""),
    estimateTitle: String(t.estimateTitle || "Estimate"),
    estimateIntro: String(t.estimateIntro || "This is with reference to Job no {JOB_NO}. The handset on inspection found with complaint: {COMPLAINT}"),
    estimateNote: String(t.estimateNote || "Please note: Actual cost may vary at the time of repair."),
    estimateTerms: String(t.estimateTerms || ""),
    taxInvoiceTitle: String(t.taxInvoiceTitle || "Tax Invoice"),
    taxTerms: String(t.taxTerms || "Out-of-warranty parts replaced carry limited service warranty as applicable. This is a computer generated invoice."),
    invoiceTerms: String(t.invoiceTerms || t.taxTerms || ""),
    signatureCustomerLabel: String(t.signatureCustomerLabel || "Customer signature"),
    signatureStaffLabel: String(t.signatureStaffLabel || "Signature by Service staff")
  };
}

function splitInclusiveGst(amount, rate = 18) {
  const gross = Number(amount || 0);
  const r = Number(rate || 0);
  if (!Number.isFinite(gross) || gross <= 0 || !Number.isFinite(r) || r <= 0) {
    return { gross: Math.max(0, gross || 0), taxable: Math.max(0, gross || 0), gst: 0 };
  }
  const taxable = gross * 100 / (100 + r);
  const gst = gross - taxable;
  return { gross, taxable, gst };
}

function formatParts(parts) {
  if (!Array.isArray(parts) || !parts.length) return "";
  return parts
    .map((p) => {
      if (typeof p === "string") return p;
      const code = p.partCode ? `[${p.partCode}] ` : "";
      const qty = Number(p.qty || 1);
      const price = Number(p.price || 0);
      return `${code}${p.name} x${qty} @ ${price.toFixed(2)}`;
    })
    .join(", ");
}

function escapeHtml(v) {
  return String(v || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function validateTenDigit(value) {
  return !value || /^\d{10}$/.test(value);
}

function findDuplicateByContact(number) {
  return state.jobs.filter((j) => j.contactNumber === number || j.altNumber === number).slice().reverse();
}

function closeDuplicatePanel() {
  if (!el.duplicateSidePanel || !el.duplicateSideList) return;
  el.duplicateSidePanel.classList.add("hidden");
  el.duplicateSideList.innerHTML = "";
}

function loadDuplicateJobToForm(job) {
  if (!job || !el.jobForm) return;
  el.jobForm.customerType.value = job.customerType || "Customer";
  el.jobForm.name.value = job.name || "";
  el.jobForm.address.value = job.address || "";
  el.jobForm.contactNumber.value = job.contactNumber || "";
  el.jobForm.altNumber.value = job.altNumber || "";
  el.jobForm.email.value = job.email || "";
  el.jobForm.warrantyCategory.value = job.warrantyCategory || "";
  el.jobForm.brand.value = job.brand || "";
  el.jobForm.model.value = job.model || "";
  el.jobForm.imei.value = job.imei || "";
  el.jobForm.productCondition.value = job.productCondition || "";
  el.jobForm.complaint.value = job.complaint || "";
  el.jobForm.advanceAmount.value = job.advanceAmount || "";
  el.jobForm.totalAmount.value = job.totalAmount || "";
  el.jobForm.invoiceMode.value = job.invoiceMode || "";
  updateConditionalFields();
}

function renderDuplicateSidePanel(matches) {
  if (!el.duplicateSidePanel || !el.duplicateSideList) return;
  if (!matches.length) {
    closeDuplicatePanel();
    return;
  }
  el.duplicateSidePanel.classList.remove("hidden");
  el.duplicateSideList.innerHTML = matches
    .map((m) => `
      <button type="button" class="duplicate-row" data-id="${m.id}">
        <strong>${escapeHtml(m.jobNumber || "-")}</strong><br/>
        ${escapeHtml(m.name || "-")} | ${escapeHtml(m.contactNumber || "-")}<br/>
        ${escapeHtml(m.brand || "")} ${escapeHtml(m.model || "")} | ${escapeHtml(m.status || "-")}<br/>
        ${new Date(m.createdAt).toLocaleString()}
      </button>
    `)
    .join("");

  Array.from(el.duplicateSideList.querySelectorAll(".duplicate-row")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const job = matches.find((m) => m.id === btn.dataset.id);
      if (!job) return;
      loadDuplicateJobToForm(job);
      closeDuplicatePanel();
    });
  });
}

function updateDuplicateHint() {
  if (!el.duplicateHint || !el.jobForm?.contactNumber) return;
  const number = (el.jobForm.contactNumber.value || "").trim();
  if (!/^\d{10}$/.test(number)) {
    el.duplicateHint.classList.add("hidden");
    el.duplicateHint.innerHTML = "";
    closeDuplicatePanel();
    return;
  }

  const matches = findDuplicateByContact(number);
  if (!matches.length) {
    el.duplicateHint.classList.add("hidden");
    el.duplicateHint.innerHTML = "";
    closeDuplicatePanel();
    return;
  }

  el.duplicateHint.classList.remove("hidden");
  el.duplicateHint.innerHTML = `Same number already registered (${matches.length}). Previous jobs shown on side panel.`;
  renderDuplicateSidePanel(matches);
}

function updateConditionalFields() {
  const w = el.warrantyCategory.value;
  const isOow = w === "OOW";
  el.advanceAmount.disabled = !isOow;
  el.totalAmount.disabled = !isOow;
  if (!isOow) {
    el.advanceAmount.value = "";
    el.totalAmount.value = "";
  }

  const needInvoiceMode = ["IW", "DOA"].includes(w);
  el.invoiceMode.disabled = !needInvoiceMode;
  if (!needInvoiceMode) el.invoiceMode.value = "";
}

function addLedgerEntry(entry) {
  const amount = Number(entry?.amount || 0);
  if (!Number.isFinite(amount) || amount < 0) return null;
  const row = {
    id: uid("LED"),
    at: entry?.at || nowIso(),
    type: entry?.type === "OUT" ? "OUT" : "IN",
    category: String(entry?.category || "General").trim(),
    amount,
    mode: String(entry?.mode || "Cash").trim(),
    note: String(entry?.note || "").trim(),
    by: String(entry?.by || (getCurrentUser()?.username || "system")).trim(),
    jobId: entry?.jobId || null
  };
  state.cashLedger.push(row);
  return row.id;
}

function upsertJobCollectionLedger(job, amount, mode, byUser) {
  if (!job) return;
  const touchAt = nowIso();
  const cleanAmount = Number(amount || 0);
  const cleanMode = String(mode || "Cash").trim() || "Cash";
  const by = String(byUser || (getCurrentUser()?.username || "system")).trim();

  const existingId = job.closeCollection?.ledgerId || null;
  if (existingId) {
    state.cashLedger = state.cashLedger.filter((r) => r.id !== existingId);
  }

  let ledgerId = null;
  if (Number.isFinite(cleanAmount) && cleanAmount > 0) {
    ledgerId = addLedgerEntry({
      type: "IN",
      category: "OOW Collection",
      amount: cleanAmount,
      mode: cleanMode,
      note: `Collection for ${job.jobNumber}`,
      by,
      jobId: job.id
    });
  }

  job.closeCollection = {
    amount: Number.isFinite(cleanAmount) ? cleanAmount : 0,
    mode: cleanMode,
    by,
    at: touchAt,
    ledgerId
  };
  touchJob(job, touchAt);
}

function removeJobCollectionLedger(job) {
  if (!job) return;
  const ledgerId = job?.closeCollection?.ledgerId;
  if (!ledgerId) return;
  state.cashLedger = state.cashLedger.filter((r) => r.id !== ledgerId);
  if (job.closeCollection) job.closeCollection.ledgerId = null;
  touchJob(job);
}

function normalizePriceRow(row) {
  const clean = {
    id: row?.id || uid("PRC"),
    itemCode: String(row?.itemCode || row?.partCode || row?.code || "").trim(),
    partName: String(row?.partName || row?.part || row?.name || "").trim(),
    brand: String(row?.brand || "").trim(),
    model: String(row?.model || "").trim(),
    price: Number(row?.price || row?.rate || row?.amount || 0),
    updatedAt: row?.updatedAt || nowIso(),
    updatedBy: row?.updatedBy || (getCurrentUser()?.username || "admin")
  };
  if (!Number.isFinite(clean.price) || clean.price < 0) clean.price = 0;
  return clean;
}

function getPriceIdentity(item) {
  const code = String(item.itemCode || "").toLowerCase();
  if (code) return `code:${code}`;
  return `name:${String(item.partName || "").toLowerCase()}|${String(item.brand || "").toLowerCase()}|${String(item.model || "").toLowerCase()}`;
}

function upsertPriceItem(rawItem) {
  const item = normalizePriceRow(rawItem);
  if (!item.partName) return false;
  const key = getPriceIdentity(item);
  const idx = state.priceList.findIndex((p) => getPriceIdentity(p) === key);
  if (idx >= 0) {
    state.priceList[idx] = {
      ...state.priceList[idx],
      ...item,
      updatedAt: nowIso(),
      updatedBy: getCurrentUser()?.username || "admin"
    };
  } else {
    state.priceList.push({
      ...item,
      updatedAt: nowIso(),
      updatedBy: getCurrentUser()?.username || "admin"
    });
  }
  return true;
}

function filteredPriceRows() {
  const q = String(el.priceSearch?.value || "").trim().toLowerCase();
  let rows = [...(state.priceList || [])];
  if (q) {
    rows = rows.filter((p) => `${p.itemCode} ${p.partName} ${p.brand} ${p.model} ${p.price}`.toLowerCase().includes(q));
  }
  return rows.sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime());
}

function renderPriceList() {
  if (!el.priceBody) return;
  const user = getCurrentUser();
  const isAdmin = user?.role === "admin";
  if (el.priceAddForm) el.priceAddForm.classList.toggle("hidden", !isAdmin);
  if (el.priceImportForm) el.priceImportForm.classList.toggle("hidden", !isAdmin);

  const rows = filteredPriceRows();
  el.priceBody.innerHTML = rows.length
    ? rows
        .map(
          (p) => `
      <tr>
        <td>${escapeHtml(p.itemCode || "-")}</td>
        <td>${escapeHtml(p.partName)}</td>
        <td>${escapeHtml(p.brand || "-")}</td>
        <td>${escapeHtml(p.model || "-")}</td>
        <td>${money(p.price || 0)}</td>
        <td>${new Date(p.updatedAt).toLocaleString()}</td>
        <td>${isAdmin ? `<button type="button" class="deletePriceBtn action-btn" data-id="${p.id}">Delete</button>` : "-"}</td>
      </tr>`
        )
        .join("")
    : '<tr><td colspan="7">No price list items found.</td></tr>';

  if (!isAdmin) return;
  Array.from(document.querySelectorAll(".deletePriceBtn")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const id = btn.dataset.id;
      state.priceList = state.priceList.filter((x) => x.id !== id);
      afterChange();
    });
  });
}

function parseExcelRowsWithXlsx(file) {
  return new Promise((resolve, reject) => {
    if (!window.XLSX) {
      reject(new Error("Excel parser unavailable."));
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = reader.result;
        const wb = window.XLSX.read(data, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = window.XLSX.utils.sheet_to_json(sheet, { defval: "" });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function findPriceForPart(partName) {
  const part = String(partName || "").trim().toLowerCase();
  if (!part) return null;
  return state.priceList.find((p) => String(p.partName || "").trim().toLowerCase() === part) || null;
}

function findPricePartByInput(text) {
  const q = String(text || "").trim().toLowerCase();
  if (!q) return null;
  const exact = state.priceList.find((p) => {
    const part = String(p.partName || "").trim().toLowerCase();
    const code = String(p.itemCode || "").trim().toLowerCase();
    return q === part || q === code || q === `${code} - ${part}`.trim().toLowerCase();
  });
  if (exact) return exact;
  const matches = state.priceList.filter((p) => {
    const part = String(p.partName || "").toLowerCase();
    const code = String(p.itemCode || "").toLowerCase();
    const model = String(p.model || "").toLowerCase();
    const brand = String(p.brand || "").toLowerCase();
    return part.includes(q) || code.includes(q) || model.includes(q) || brand.includes(q);
  });
  return matches.length === 1 ? matches[0] : null;
}

function getLatestParts(job) {
  const latest = (job?.engineerUpdates || []).slice().reverse().find((u) => Array.isArray(u.partsUsed) && u.partsUsed.length);
  return latest?.partsUsed || [];
}

function getPartsTotalFromJob(job) {
  const latestParts = getLatestParts(job);
  const partsTotal = latestParts.reduce((sum, p) => sum + Number(p.qty || 0) * Number(p.price || 0), 0);
  if (partsTotal > 0) return partsTotal;
  return Number(job?.totalAmount || job?.estimate?.amount || 0);
}

function getServiceChargeBase(job) {
  const base = Number(job?.serviceCharge || 0);
  if (!Number.isFinite(base) || base < 0) return 0;
  return base;
}

function getServiceChargeTax(baseOrJob) {
  const base = typeof baseOrJob === "number" ? Number(baseOrJob || 0) : getServiceChargeBase(baseOrJob);
  const clean = Number.isFinite(base) && base > 0 ? base : 0;
  return clean * 18 / 100;
}

function getServiceChargeGross(baseOrJob) {
  const base = typeof baseOrJob === "number" ? Number(baseOrJob || 0) : getServiceChargeBase(baseOrJob);
  const clean = Number.isFinite(base) && base > 0 ? base : 0;
  return clean + getServiceChargeTax(clean);
}

function getNetBillAmount(job, opts = {}) {
  const partsTotal = Number.isFinite(Number(opts.partsTotal)) ? Number(opts.partsTotal) : getPartsTotalFromJob(job);
  const warranty = String(opts.warrantyCategory || job?.warrantyCategory || "");
  const serviceBaseRaw = Number.isFinite(Number(opts.serviceChargeBase)) ? Number(opts.serviceChargeBase) : getServiceChargeBase(job);
  const serviceBase = warranty === "OOW" ? serviceBaseRaw : 0;
  const serviceGross = getServiceChargeGross(serviceBase);
  const advance = Number.isFinite(Number(opts.advanceAmount)) ? Number(opts.advanceAmount) : Number(job?.advanceAmount || 0);
  const discount = Number.isFinite(Number(opts.discountAmount)) ? Number(opts.discountAmount) : Number(job?.discountAmount || 0);
  return Math.max(0, partsTotal + serviceGross - advance - discount);
}

function clearPartSuggestions(resetSelected = false) {
  if (resetSelected) selectedPartItem = null;
  if (!el.partSuggestionBox) return;
  el.partSuggestionBox.classList.add("hidden");
  el.partSuggestionBox.innerHTML = "";
}

function renderPartSuggestions(query) {
  if (!el.partSuggestionBox) return;
  const q = String(query || "").trim().toLowerCase();
  if (!q) {
    clearPartSuggestions(true);
    return;
  }
  const matches = state.priceList
    .filter((p) => {
      const part = String(p.partName || "").toLowerCase();
      const code = String(p.itemCode || "").toLowerCase();
      const model = String(p.model || "").toLowerCase();
      const brand = String(p.brand || "").toLowerCase();
      return part.includes(q) || code.includes(q) || model.includes(q) || brand.includes(q);
    })
    .slice(0, 8);

  if (!matches.length) {
    clearPartSuggestions(true);
    return;
  }

  el.partSuggestionBox.classList.remove("hidden");
  el.partSuggestionBox.innerHTML = matches
    .map((m) => `<button type="button" class="part-suggestion-item" data-id="${m.id}">${escapeHtml(m.itemCode || "-")} | ${escapeHtml(m.partName)} | ${escapeHtml(m.model || m.brand || "-")} | ${money(m.price || 0)}</button>`)
    .join("");

  Array.from(el.partSuggestionBox.querySelectorAll(".part-suggestion-item")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const item = state.priceList.find((x) => x.id === btn.dataset.id);
      if (!item) return;
      selectedPartItem = item;
      el.partSearch.value = `${item.itemCode ? `${item.itemCode} - ` : ""}${item.partName}`;
      clearPartSuggestions();
    });
  });
}

function toggleLedgerDateInputs() {
  if (!el.ledgerPeriod || !el.ledgerFromDate || !el.ledgerToDate || !el.ledgerAnchorDate) return;
  const isCustom = el.ledgerPeriod.value === "custom";
  el.ledgerFromDate.classList.toggle("hidden", !isCustom);
  el.ledgerToDate.classList.toggle("hidden", !isCustom);
  el.ledgerAnchorDate.classList.toggle("hidden", isCustom);
}

function ledgerRangeByPeriod(period, anchorDate, fromDate, toDate) {
  if (period === "custom") {
    const from = fromDate ? new Date(`${fromDate}T00:00:00`).getTime() : null;
    const to = toDate ? new Date(`${toDate}T23:59:59`).getTime() : null;
    return { from, to };
  }

  const base = anchorDate ? new Date(`${anchorDate}T00:00:00`) : now();
  const d = new Date(base.getTime());
  let start = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
  let end = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999);

  if (period === "weekly") {
    const day = start.getDay();
    const diff = day === 0 ? 6 : day - 1;
    start.setDate(start.getDate() - diff);
    end = new Date(start.getTime());
    end.setDate(end.getDate() + 6);
    end.setHours(23, 59, 59, 999);
  } else if (period === "monthly") {
    start = new Date(d.getFullYear(), d.getMonth(), 1, 0, 0, 0, 0);
    end = new Date(d.getFullYear(), d.getMonth() + 1, 0, 23, 59, 59, 999);
  } else if (period === "yearly") {
    start = new Date(d.getFullYear(), 0, 1, 0, 0, 0, 0);
    end = new Date(d.getFullYear(), 11, 31, 23, 59, 59, 999);
  }

  return { from: start.getTime(), to: end.getTime() };
}

function filteredLedgerRows() {
  const q = String(el.ledgerSearch?.value || "").trim().toLowerCase();
  const period = el.ledgerPeriod?.value || "daily";
  const anchor = el.ledgerAnchorDate?.value || "";
  const fromDate = el.ledgerFromDate?.value || "";
  const toDate = el.ledgerToDate?.value || "";
  const range = ledgerRangeByPeriod(period, anchor, fromDate, toDate);

  let rows = [...(state.cashLedger || [])];
  if (range.from !== null) rows = rows.filter((r) => new Date(r.at).getTime() >= range.from);
  if (range.to !== null) rows = rows.filter((r) => new Date(r.at).getTime() <= range.to);

  if (q) {
    rows = rows.filter((r) => {
      const job = state.jobs.find((j) => j.id === r.jobId);
      const txt = `${r.type} ${r.category} ${r.mode} ${r.note} ${r.by} ${job?.jobNumber || ""}`.toLowerCase();
      return txt.includes(q);
    });
  }
  return rows.sort((a, b) => new Date(b.at).getTime() - new Date(a.at).getTime());
}

function renderLedger() {
  if (!el.ledgerBody) return;
  const user = getCurrentUser();
  const isAdmin = user?.role === "admin";
  const canAddExpense = user && ["admin", "receptionist"].includes(user.role);

  if (el.ledgerExpenseForm) el.ledgerExpenseForm.classList.toggle("hidden", !canAddExpense);
  if (el.ledgerAdminCashForm) el.ledgerAdminCashForm.classList.toggle("hidden", !isAdmin);

  toggleLedgerDateInputs();
  const rows = filteredLedgerRows();
  const totalIn = rows.reduce((s, r) => s + (r.type === "IN" ? Number(r.amount || 0) : 0), 0);
  const totalOut = rows.reduce((s, r) => s + (r.type === "OUT" ? Number(r.amount || 0) : 0), 0);
  const balance = totalIn - totalOut;

  if (el.ledgerTotalIn) el.ledgerTotalIn.textContent = money(totalIn);
  if (el.ledgerTotalOut) el.ledgerTotalOut.textContent = money(totalOut);
  if (el.ledgerBalance) el.ledgerBalance.textContent = money(balance);

  el.ledgerBody.innerHTML = rows.length
    ? rows
        .map((r) => {
          const job = state.jobs.find((j) => j.id === r.jobId);
          return `
      <tr>
        <td>${new Date(r.at).toLocaleString()}</td>
        <td>${escapeHtml(r.type)}</td>
        <td>${escapeHtml(r.category || "-")}</td>
        <td>${money(r.amount)}</td>
        <td>${escapeHtml(r.mode || "-")}</td>
        <td>${escapeHtml(job?.jobNumber || "-")}</td>
        <td>${escapeHtml(r.note || "-")}</td>
        <td>${escapeHtml(r.by || "-")}</td>
      </tr>`;
        })
        .join("")
    : '<tr><td colspan="8">No ledger entries for selected filter.</td></tr>';
}

function getKpiDateRange() {
  return {
    from: String(el.kpiDateFrom?.value || ""),
    to: String(el.kpiDateTo?.value || "")
  };
}

function getKpiJobs() {
  const { from, to } = getKpiDateRange();
  let rows = [...state.jobs];
  if (from) {
    const fromTs = new Date(`${from}T00:00:00`).getTime();
    rows = rows.filter((j) => new Date(j.createdAt).getTime() >= fromTs);
  }
  if (to) {
    const toTs = new Date(`${to}T23:59:59`).getTime();
    rows = rows.filter((j) => new Date(j.createdAt).getTime() <= toTs);
  }
  return rows;
}

function warrantyCount(rows, target) {
  const t = String(target || "").trim().toLowerCase();
  return rows.filter((j) => String(j.warrantyCategory || "").trim().toLowerCase() === t).length;
}

function renderKpis() {
  const jobs = getKpiJobs();
  const total = jobs.length;
  const open = jobs.filter((j) => !["Delivered", "Cancelled"].includes(j.status)).length;
  const ready = jobs.filter((j) => j.status === "Ready Pickup").length;
  const delivered = jobs.filter((j) => j.status === "Delivered").length;
  const iw = warrantyCount(jobs, "IW");
  const oow = warrantyCount(jobs, "OOW");
  const doa = warrantyCount(jobs, "DOA");
  const refurbish = warrantyCount(jobs, "Refurbish");
  const repacking = warrantyCount(jobs, "Repacking");

  if (el.kpiTotalJobs) el.kpiTotalJobs.textContent = total;
  if (el.kpiOpenJobs) el.kpiOpenJobs.textContent = open;
  if (el.kpiIw) el.kpiIw.textContent = iw;
  if (el.kpiOow) el.kpiOow.textContent = oow;
  if (el.kpiDoa) el.kpiDoa.textContent = doa;
  if (el.kpiRefurbish) el.kpiRefurbish.textContent = refurbish;
  if (el.kpiRepacking) el.kpiRepacking.textContent = repacking;
  if (el.kpiReady) el.kpiReady.textContent = ready;
  if (el.kpiDelivered) el.kpiDelivered.textContent = delivered;
}

function renderStatusBoards() {
  const hiddenQuickLinks = new Set(["open price list", "open cash ledger", "price list", "cash ledger", "register", "pending approvals"]);
  const statuses = (state.masters.repairStatuses || []).filter((s) => !hiddenQuickLinks.has(String(s || "").trim().toLowerCase()));
  if (!statuses.length) {
    el.statusBoards.innerHTML = '<div class="panel mini-panel"><h3>No status boards configured.</h3></div>';
    return;
  }
  el.statusBoards.innerHTML = statuses
    .map((status) => {
      const list = state.jobs.filter((j) => j.status === status).slice(-8);
      return `
      <div class="panel mini-panel status-card-click" data-status="${escapeHtml(status)}">
        <h3>${escapeHtml(status)} (${list.length})</h3>
        <ul>
          ${list.length ? list.map((j) => `<li>${escapeHtml(j.jobNumber)} - ${escapeHtml(j.name)}</li>`).join("") : "<li>No jobs</li>"}
        </ul>
      </div>`;
    })
    .join("");

  Array.from(el.statusBoards.querySelectorAll(".status-card-click")).forEach((card) => {
    card.addEventListener("click", () => {
      el.jobStatusFilter.value = card.dataset.status || "";
      switchPage("jobsheets");
      renderJobsTable();
    });
  });
}

function filteredJobsForView() {
  const q = (el.jobSearch.value || "").trim().toLowerCase();
  const sf = el.jobStatusFilter.value || "";
  const fromDate = el.jobDateFrom?.value || "";
  const toDate = el.jobDateTo?.value || "";
  let rows = [...state.jobs];

  if (q) {
    rows = rows.filter((j) => {
      const txt = `${j.jobNumber} ${j.customerType || ""} ${j.name} ${j.contactNumber} ${j.model} ${j.imei} ${j.status}`.toLowerCase();
      return txt.includes(q);
    });
  }
  if (sf) rows = rows.filter((j) => j.status === sf);
  if (fromDate) {
    const fromTs = new Date(`${fromDate}T00:00:00`).getTime();
    rows = rows.filter((j) => new Date(j.createdAt).getTime() >= fromTs);
  }
  if (toDate) {
    const toTs = new Date(`${toDate}T23:59:59`).getTime();
    rows = rows.filter((j) => new Date(j.createdAt).getTime() <= toTs);
  }
  if (quickWarrantyFilter) {
    rows = rows.filter((j) => j.warrantyCategory === quickWarrantyFilter);
  }
  if (quickOpenOnly) {
    rows = rows.filter((j) => !["Delivered", "Cancelled"].includes(j.status));
  }
  return rows;
}

function renderJobsTable() {
  const user = getCurrentUser();
  const rows = filteredJobsForView().slice().reverse();
  el.jobsBody.innerHTML = rows.length
    ? rows
        .map((j) => `
        <tr>
          <td>${escapeHtml(j.jobNumber)}</td>
          <td>${new Date(j.createdAt).toLocaleDateString()}</td>
          <td>${escapeHtml(j.name)}</td>
          <td>${escapeHtml(j.contactNumber)}</td>
          <td>${escapeHtml(j.brand)} ${escapeHtml(j.model)}</td>
          <td>${escapeHtml(j.warrantyCategory)}</td>
          <td>${escapeHtml(j.status)}</td>
          <td>${escapeHtml(j.assignedEngineer || "-")}</td>
          <td><div class="table-actions">
            <button class="openJobBtn action-btn" data-id="${j.id}" type="button">Open</button>
            <button class="printBtn action-btn" data-id="${j.id}" type="button">Print</button>
            ${(j.warrantyCategory === "OOW") ? `<button class="printInvBtn action-btn" data-id="${j.id}" type="button">Est.Inv</button>` : ""}
            ${(j.warrantyCategory === "OOW") ? `<button class="printTaxBtn action-btn" data-id="${j.id}" type="button">Tax Inv</button>` : ""}
            ${(j.attachments || []).length ? `<button class="filesBtn action-btn" data-id="${j.id}" type="button">Files</button>` : ""}
            ${
              (user?.role === "receptionist" || user?.role === "admin") && j.status === "Ready Pickup"
                ? `${
                    j.warrantyCategory === "OOW"
                      ? `<span class="action-note">Auto: ${money(getNetBillAmount(j))}</span>
                         <select class="closeModeInput action-select" data-id="${j.id}">
                           ${["Cash", "UPI", "Card", "Bank"]
                             .map((m) => `<option ${String(j.closeCollection?.mode || "Cash") === m ? "selected" : ""}>${m}</option>`)
                             .join("")}
                         </select>`
                      : ""
                  }
                  <button class="closeJobBtn secondary action-btn" data-id="${j.id}" type="button">Close</button>`
                : ""
            }
            ${user?.role === "admin" ? `<button class="adminEditBtn secondary action-btn" data-id="${j.id}" type="button">Edit</button>` : ""}
            ${user?.role === "admin" ? `<button class="deleteJobBtn secondary action-btn" data-id="${j.id}" type="button">Delete</button>` : ""}
          </div></td>
        </tr>`)
        .join("")
    : '<tr><td colspan="9">No jobs found.</td></tr>';

  Array.from(document.querySelectorAll(".printBtn")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const job = state.jobs.find((j) => j.id === btn.dataset.id);
      if (job) printJobSheet(job);
    });
  });

  Array.from(document.querySelectorAll(".printInvBtn")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const job = state.jobs.find((j) => j.id === btn.dataset.id);
      if (!job) return;
      const parts = (job.engineerUpdates || []).slice().reverse().find((u) => Array.isArray(u.partsUsed) && u.partsUsed.length)?.partsUsed || [];
      printEstimateInvoice(job, parts);
    });
  });

  Array.from(document.querySelectorAll(".printTaxBtn")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const job = state.jobs.find((j) => j.id === btn.dataset.id);
      if (!job) return;
      printTaxInvoice(job);
    });
  });

  Array.from(document.querySelectorAll(".openJobBtn")).forEach((btn) => {
    btn.addEventListener("click", () => {
      selectedEngineerJobId = btn.dataset.id;
      engineerFocusMode = true;
      switchPage("engineer");
      renderEngineerMeta();
    });
  });

  Array.from(document.querySelectorAll(".filesBtn")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const job = state.jobs.find((j) => j.id === btn.dataset.id);
      if (!job?.attachments?.length) return;
      job.attachments.forEach((att) => downloadAttachment(att));
    });
  });

  Array.from(document.querySelectorAll(".closeJobBtn")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const job = state.jobs.find((j) => j.id === btn.dataset.id);
      const who = getCurrentUser();
      if (!job || !who) return;
      if (job.warrantyCategory === "OOW") {
        const modeInput = document.querySelector(`.closeModeInput[data-id="${job.id}"]`);
        const value = getNetBillAmount(job);
        if (!Number.isFinite(value) || value < 0) return alert("Invalid amount.");
        const mode = String(modeInput?.value || "Cash");
        upsertJobCollectionLedger(job, value, mode, who.username);
        job.payments = (job.payments || []).filter((p) => p.source !== "CLOSE_COLLECTION");
        if (value > 0) {
          job.payments.push({ id: uid("PAY"), amount: value, method: `Reception ${mode}`, at: nowIso(), source: "CLOSE_COLLECTION" });
        }
      } else {
        removeJobCollectionLedger(job);
      }
      job.status = "Delivered";
      const touchAt = nowIso();
      job.timeline.push({
        at: touchAt,
        by: who.username,
        note: job.warrantyCategory === "OOW"
          ? `Job closed by reception after delivery. Auto Collected: ${money(getNetBillAmount(job))}`
          : "Job closed by reception after delivery."
      });
      touchJob(job, touchAt);
      afterChange();
    });
  });

  Array.from(document.querySelectorAll(".adminEditBtn")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const job = state.jobs.find((j) => j.id === btn.dataset.id);
      if (!job || !el.adminEditDialog || !el.adminEditForm) return;
      selectedAdminEditJobId = job.id;
      el.adminEditForm.name.value = job.name || "";
      if (el.adminEditForm.customerType) el.adminEditForm.customerType.value = job.customerType || "Customer";
      el.adminEditForm.contactNumber.value = job.contactNumber || "";
      el.adminEditForm.altNumber.value = job.altNumber || "";
      el.adminEditForm.warrantyCategory.value = job.warrantyCategory || "";
      el.adminEditForm.address.value = job.address || "";
      el.adminEditForm.email.value = job.email || "";
      el.adminEditForm.model.value = job.model || "";
      el.adminEditForm.imei.value = job.imei || "";
      el.adminEditDialog.showModal();
    });
  });

  Array.from(document.querySelectorAll(".deleteJobBtn")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const job = state.jobs.find((j) => j.id === btn.dataset.id);
      if (!job) return;
      const ok = confirm(`Delete job ${job.jobNumber}?`);
      if (!ok) return;
      state.jobs = state.jobs.filter((j) => j.id !== job.id);
      state.cashLedger = (state.cashLedger || []).filter((r) => r.jobId !== job.id);
      afterChange();
    });
  });
}

function printJobSheet(job) {
  const tpl = getJobSheetTemplate();
  const ps = getPrintStyle();
  const text = getPrintText();
  const createdDate = new Date(job.createdAt);
  const printAt = new Date();
  const imeiBar = (job.imei || "").replace(/\s+/g, "");
  const defaultTerms = [
    "1. realme respects and devotes to protect your personal privacy and handles personal data only for service usage.",
    "2. Before inspection, please log out account and back up important data.",
    "3. Service Center safe keeps device for up to 60 days from repair date.",
    "4. Receipt/Jobsheet is required for handover. Keep it safely.",
    "5. Customer agrees to receive SMS/call updates regarding repair status and estimate.",
    "6. Out-of-warranty replaced parts are chargeable; defective parts may not be returned."
  ];
  const termsFromPrintText = String(text.jobTerms || "")
    .split(/\r?\n/)
    .map((x) => x.trim())
    .filter(Boolean);
  const termsFromTemplate = String(tpl.footerNote || "")
    .split(/\r?\n/)
    .map((x) => x.trim())
    .filter(Boolean);
  const finalTerms = termsFromPrintText.length
    ? termsFromPrintText
    : (termsFromTemplate.length ? termsFromTemplate : defaultTerms);
  const html = `
  <html><head><title>${escapeHtml(job.jobNumber)} - Job Sheet</title>
  <style>
    body{font-family:${escapeHtml(ps.fontFamily)},sans-serif;font-size:${Math.max(10, Number(ps.baseSize || 12))}px;padding:10px;color:#111;}
    .sheet{max-width:980px;margin:0 auto;border:1px solid #111;}
    .top{display:grid;grid-template-columns:1fr 2fr 1fr;align-items:center;border-bottom:1px solid #111;padding:6px 8px;gap:8px;}
    .brand{font-size:${Math.max(13, Number(ps.baseSize || 12) + 5)}px;font-weight:700;text-transform:lowercase;}
    .title{font-size:${Math.max(18, Number(ps.headingSize || 22))}px;font-weight:700;text-align:center;}
    .barcode{text-align:right;font-family:monospace;}
    .barcode-box{display:inline-block;padding:2px 6px;border:1px solid #111;letter-spacing:2px;font-size:12px;}
    table{width:100%;border-collapse:collapse;}
    th,td{border:1px solid #111;padding:5px;vertical-align:top;}
    th{font-weight:700;background:#f4f4f4;}
    .section{background:#ececec;text-align:center;font-weight:700;}
    .terms{padding:6px 8px;line-height:1.45;border-top:1px solid #111;}
    .terms h4{margin:0 0 6px;font-size:13px;}
    .terms p{margin:0 0 5px;}
    .pin{display:grid;grid-template-columns:repeat(3,34px);gap:6px;justify-content:center;padding:8px;}
    .pin div{border:1px solid #06b6d4;border-radius:50%;width:34px;height:34px;display:flex;align-items:center;justify-content:center;font-weight:700;}
    .footer{font-size:11px;}
  </style></head><body>
    <div class="sheet">
      <div class="top">
        <div class="brand">${escapeHtml((tpl.companyName || "realme").split(" ")[0] || "realme")}</div>
        <div class="title">${escapeHtml(text.jobSheetTitle || "Service Center Handover Report")}</div>
        <div class="barcode"><span class="barcode-box">${escapeHtml(job.jobNumber)}</span><br/><small>${escapeHtml(imeiBar || job.jobNumber)}</small></div>
      </div>
      ${tpl.headerNote ? `<div style="padding:6px 8px;border-bottom:1px solid #111;">${escapeHtml(tpl.headerNote)}</div>` : ""}
      <table>
        <tr><th class="section" colspan="4">Customer information</th><th class="section" colspan="4">Dealer Information</th></tr>
        <tr>
          <th>Customer name</th><td>${escapeHtml(job.name || "-")}</td>
          <th>Type</th><td>${escapeHtml(job.customerType || "Customer")}</td>
          <th>Customer phone number</th><td>${escapeHtml(job.contactNumber || "-")}</td>
          <th>Dealer/Salesman Name</th><td>${escapeHtml(job.assignedEngineer || "-")}</td>
        </tr>
        <tr>
          <th>Customer address</th><td>${escapeHtml(job.address || "-")}</td>
          <th>Customer phone number</th><td>${escapeHtml(job.contactNumber || "-")}</td>
          <th>Dealer/Salesman Phone No.</th><td colspan="3">${escapeHtml(tpl.phone || "-")}</td>
        </tr>
        <tr><th class="section" colspan="8">Product Information</th></tr>
        <tr>
          <th>Model</th><td>${escapeHtml(`${job.brand || ""} ${job.model || ""}`.trim() || "-")}</td>
          <th>Colour</th><td>${escapeHtml(job.productCondition || "-")}</td>
          <th>Purchase Date</th><td>${createdDate.toLocaleDateString("en-CA")}</td>
          <th>Receiving Date</th><td>${createdDate.toLocaleDateString("en-CA")}</td>
        </tr>
        <tr>
          <th>IMEI</th><td colspan="3">${escapeHtml(job.imei || "-")}</td>
          <th>RAM+ROM</th><td colspan="3">${escapeHtml("-")}</td>
        </tr>
        <tr>
          <th>Appearance</th><td colspan="3">${escapeHtml(job.productCondition || "Used")}</td>
          <th>Standard accessories</th><td colspan="3">${escapeHtml(job.attachments?.length ? job.attachments.map((a) => a.name).join(", ") : "-")}</td>
        </tr>
        <tr>
          <th>Customer Fault Description</th><td colspan="3">${escapeHtml(job.complaint || "-")}</td>
          <th>Special Description</th><td colspan="3">${escapeHtml((job.engineerUpdates || []).slice(-1)[0]?.remark || job.complaint || "-")}</td>
        </tr>
        <tr><th class="section" colspan="8">Advance Payment against parts</th></tr>
        <tr><th>Amount</th><td colspan="7">${Number(job.advanceAmount || 0).toFixed(2)}</td></tr>
        <tr><th>Service Charge</th><td colspan="3">${getServiceChargeBase(job).toFixed(2)}</td><th>Service GST (18%)</th><td colspan="3">${getServiceChargeTax(job).toFixed(2)}</td></tr>
        <tr><th>Discount</th><td colspan="3">${Number(job.discountAmount || 0).toFixed(2)}</td><th>Net Bill</th><td colspan="3">${getNetBillAmount(job).toFixed(2)}</td></tr>
      </table>
      <div class="terms">
        <h4>${escapeHtml(text.jobTermsHeading)}</h4>
        ${finalTerms.map((t) => `<p>${escapeHtml(t)}</p>`).join("")}
      </div>
      <table>
        <tr>
          <td style="width:72%;">
            <div class="footer">
              <strong>Call Center:</strong> 1800-102-2777<br/>
              <strong>Website:</strong> www.realme.com<br/>
              <strong>E-mail:</strong> service@realme.com<br/>
              <strong>SC Address:</strong> ${escapeHtml(tpl.address || "-")}<br/>
              <strong>${escapeHtml(text.signatureCustomerLabel)}:</strong> ____________________<br/>
              <strong>${escapeHtml(text.signatureStaffLabel)}:</strong> ____________________
            </div>
          </td>
          <td>
            <div class="pin">
              <div>1</div><div>2</div><div>3</div>
              <div>4</div><div>5</div><div>6</div>
              <div>7</div><div>8</div><div>9</div>
            </div>
          </td>
        </tr>
        <tr><td colspan="2"><strong>Print Time:</strong> ${printAt.toLocaleString()}</td></tr>
      </table>
    </div>
  </body></html>`;

  const w = window.open("", "_blank");
  if (!w) return;
  w.document.write(html);
  w.document.close();
  w.focus();
  w.print();
}

function printEstimateInvoice(job, parts) {
  const txtTpl = getPrintText();
  const ps = getPrintStyle();
  const tpl = getJobSheetTemplate();
  const taxRate = 18;
  const serviceBase = getServiceChargeBase(job);
  const serviceTax = getServiceChargeTax(serviceBase);
  const serviceGross = getServiceChargeGross(serviceBase);
  const items = (parts || []).length
    ? parts.map((p) => (typeof p === "string" ? { name: p, qty: 1, price: 0 } : p))
    : [];
  if (!items.length && serviceGross <= 0) return alert("No estimate items available.");
  const introText = String(txtTpl.estimateIntro || "")
    .replaceAll("{JOB_NO}", job.jobNumber || "-")
    .replaceAll("{COMPLAINT}", job.complaint || "-")
    .replaceAll("{NAME}", job.name || "-")
    .replaceAll("{MODEL}", `${job.brand || ""} ${job.model || ""}`.trim() || "-");
  const partsRows = items.map((p) => {
    const priceItem = findPriceForPart(p.name);
    const qty = Number(p.qty || 1);
    const rateIncl = Number(p.price || priceItem?.price || 0);
    const gross = qty * rateIncl;
    const split = splitInclusiveGst(gross, taxRate);
    return {
      partCode: p.partCode || priceItem?.itemCode || "-",
      hsn: "85177990",
      desc: p.name,
      qty,
      rateIncl,
      taxable: split.taxable,
      gstRate: taxRate,
      gst: split.gst,
      gross
    };
  });
  const spareTotal = partsRows.reduce((s, r) => s + r.gross, 0);
  const spareTaxable = partsRows.reduce((s, r) => s + r.taxable, 0);
  const spareTax = partsRows.reduce((s, r) => s + r.gst, 0);
  const grand = spareTotal + serviceGross;
  const advance = Number(job.advanceAmount || 0);
  const discount = Number(job.discountAmount || 0);
  const netPayable = Math.max(0, grand - advance - discount);
  const estimateTerms = String(txtTpl.estimateTerms || "")
    .split(/\r?\n/)
    .map((x) => x.trim())
    .filter(Boolean);
  const html = `
  <html><head><title>Estimate - ${escapeHtml(job.jobNumber)}</title>
  <style>
    body{font-family:${escapeHtml(ps.fontFamily)},sans-serif;font-size:${Math.max(10, Number(ps.baseSize || 12))}px;padding:16px;color:#111;}
    h2{margin:0 0 8px;font-size:${Math.max(18, Number(ps.headingSize || 22))}px;}
    .box{border:1px solid #333;padding:10px;max-width:980px;margin:0 auto;}
    .row{display:flex;justify-content:space-between;gap:12px;}
    .row div{flex:1;}
    table{width:100%;border-collapse:collapse;margin-top:8px;font-size:12px;}
    th,td{border:1px solid #555;padding:6px;text-align:left;vertical-align:top;}
    th{background:#f1f5f9;}
    .right{text-align:right;}
    .note{margin-top:10px;font-size:12px;}
  </style></head><body>
    <div class="box">
      <h2>${escapeHtml(txtTpl.estimateTitle)}</h2>
      <div class="row">
        <div><strong>From:</strong> ${escapeHtml(tpl.companyName || "ASP")}<br/><strong>Job No:</strong> ${escapeHtml(job.jobNumber)}<br/><strong>Date:</strong> ${new Date(job.createdAt).toLocaleDateString("en-CA")}</div>
        <div><strong>To Customer:</strong> ${escapeHtml(job.name)}<br/><strong>Mobile:</strong> ${escapeHtml(job.contactNumber)}<br/><strong>Model:</strong> ${escapeHtml(`${job.brand || ""} ${job.model || ""}`.trim())}<br/><strong>IMEI:</strong> ${escapeHtml(job.imei || "-")}</div>
      </div>
      ${tpl.headerNote ? `<p><strong>Note:</strong> ${escapeHtml(tpl.headerNote)}</p>` : ""}
      <p>${escapeHtml(introText)}</p>
      <table>
        <tr>
          <th>Sno</th><th>Part Code</th><th>HSN Code</th><th>Part Description</th><th class="right">Qty</th><th class="right">Rate (Incl.GST)</th><th class="right">Taxable</th><th class="right">GST %</th><th class="right">GST Amt</th><th class="right">Total (Incl.GST)</th>
        </tr>
        ${partsRows.map((r, i) => `<tr><td>${i + 1}</td><td>${escapeHtml(r.partCode)}</td><td>${escapeHtml(r.hsn)}</td><td>${escapeHtml(r.desc)}</td><td class="right">${r.qty}</td><td class="right">${r.rateIncl.toFixed(2)}</td><td class="right">${r.taxable.toFixed(2)}</td><td class="right">${r.gstRate.toFixed(2)}</td><td class="right">${r.gst.toFixed(2)}</td><td class="right">${r.gross.toFixed(2)}</td></tr>`).join("")}
        <tr><th colspan="9" class="right">Total Spares Taxable</th><th class="right">${spareTaxable.toFixed(2)}</th></tr>
        <tr><th colspan="9" class="right">Total Spares GST</th><th class="right">${spareTax.toFixed(2)}</th></tr>
        <tr><th colspan="9" class="right">Total Spares (Incl. GST)</th><th class="right">${spareTotal.toFixed(2)}</th></tr>
        ${serviceBase > 0 ? `<tr><th colspan="9" class="right">Service Charge (Taxable)</th><th class="right">${serviceBase.toFixed(2)}</th></tr>` : ""}
        ${serviceTax > 0 ? `<tr><th colspan="9" class="right">Service GST (18%)</th><th class="right">${serviceTax.toFixed(2)}</th></tr>` : ""}
        ${serviceGross > 0 ? `<tr><th colspan="9" class="right">Service Total (Incl. GST)</th><th class="right">${serviceGross.toFixed(2)}</th></tr>` : ""}
        <tr><th colspan="9" class="right">Grand Total</th><th class="right">${grand.toFixed(2)}</th></tr>
        <tr><th colspan="9" class="right">Advance</th><th class="right">-${advance.toFixed(2)}</th></tr>
        <tr><th colspan="9" class="right">Discount</th><th class="right">-${discount.toFixed(2)}</th></tr>
        <tr><th colspan="9" class="right">Net Payable</th><th class="right">${netPayable.toFixed(2)}</th></tr>
      </table>
      <p class="note"><strong>Note:</strong> ${escapeHtml(txtTpl.estimateNote)}</p>
      ${estimateTerms.length ? `<p class="note"><strong>Terms and Conditions:</strong><br/>${estimateTerms.map((t) => escapeHtml(t)).join("<br/>")}</p>` : ""}
      <p class="note"><strong>Declaration Date:</strong> ${new Date().toLocaleDateString("en-CA")} | <strong>Signature:</strong> ____________________</p>
    </div>
  </body></html>`;
  const w = window.open("", "_blank");
  if (!w) return;
  w.document.write(html);
  w.document.close();
  w.focus();
  w.print();
}

function printTaxInvoice(job) {
  const txtTpl = getPrintText();
  const ps = getPrintStyle();
  const tpl = getJobSheetTemplate();
  const latest = (job.engineerUpdates || []).slice().reverse().find((u) => Array.isArray(u.partsUsed) && u.partsUsed.length);
  const parts = latest?.partsUsed || [];

  const taxRate = 18;
  const halfTax = taxRate / 2;
  const serviceBase = getServiceChargeBase(job);
  const serviceTax = getServiceChargeTax(serviceBase);
  const serviceGross = getServiceChargeGross(serviceBase);
  if (!parts.length && serviceGross <= 0) return alert("No items found for tax invoice.");
  const payMode = job.closeCollection?.mode || job.invoiceMode || "Cash";
  const rows = parts.map((p) => {
    const priceItem = findPriceForPart(typeof p === "string" ? p : p.name);
    const qty = Number((typeof p === "string" ? 1 : p.qty) || 1);
    const rateIncl = Number((typeof p === "string" ? 0 : p.price) || priceItem?.price || 0);
    const gross = qty * rateIncl;
    const split = splitInclusiveGst(gross, taxRate);
    const cgstAmt = split.gst / 2;
    const sgstAmt = split.gst / 2;
    return {
      desc: typeof p === "string" ? p : p.name,
      hsn: "85177990",
      qty,
      unit: "PCS",
      rate: rateIncl,
      taxable: split.taxable,
      cgstRate: halfTax,
      cgstAmt,
      sgstRate: halfTax,
      sgstAmt,
      total: gross
    };
  });
  if (serviceGross > 0) {
    rows.push({
      desc: "Service Charges",
      hsn: "998716",
      qty: 1,
      unit: "JOB",
      rate: serviceBase,
      taxable: serviceBase,
      cgstRate: halfTax,
      cgstAmt: serviceTax / 2,
      sgstRate: halfTax,
      sgstAmt: serviceTax / 2,
      total: serviceGross
    });
  }

  const taxableTotal = rows.reduce((s, r) => s + r.taxable, 0);
  const cgstTotal = rows.reduce((s, r) => s + r.cgstAmt, 0);
  const sgstTotal = rows.reduce((s, r) => s + r.sgstAmt, 0);
  const grand = rows.reduce((s, r) => s + r.total, 0);
  const advance = Number(job.advanceAmount || 0);
  const discount = Number(job.discountAmount || 0);
  const netPayable = Math.max(0, grand - advance - discount);
  const invNo = job.invoice?.invoiceNo || `INV-${new Date().getFullYear()}-${job.jobNumber}`;
  const invoiceTerms = String(txtTpl.invoiceTerms || txtTpl.taxTerms || "");

  const html = `
  <html><head><title>Tax Invoice - ${escapeHtml(invNo)}</title>
  <style>
    body{font-family:${escapeHtml(ps.fontFamily)},sans-serif;font-size:${Math.max(10, Number(ps.baseSize || 12))}px;padding:16px;color:#111;}
    .box{max-width:1100px;margin:0 auto;border:1px solid #222;padding:10px;}
    h2{margin:0 0 8px;font-size:${Math.max(18, Number(ps.headingSize || 22))}px;}
    .grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;font-size:12px;}
    table{width:100%;border-collapse:collapse;margin-top:8px;font-size:12px;}
    th,td{border:1px solid #555;padding:6px;vertical-align:top;}
    th{background:#f3f4f6;}
    .right{text-align:right;}
    .small{font-size:11px;}
  </style></head><body>
    <div class="box">
      <h2>${escapeHtml(txtTpl.taxInvoiceTitle)}</h2>
      <div class="grid">
        <div>
          <strong>Bill From:</strong> ${escapeHtml(tpl.companyName || "realme Authorized Service Center")}<br/>
          <strong>Address:</strong> ${escapeHtml(tpl.address || "-")}<br/>
          <strong>Phone:</strong> ${escapeHtml(tpl.phone || "-")}<br/>
          <strong>GSTIN:</strong> 32DIVPP6757F1XM
        </div>
        <div>
          <strong>Job No:</strong> ${escapeHtml(job.jobNumber)}<br/>
          <strong>Invoice No:</strong> ${escapeHtml(invNo)}<br/>
          <strong>Invoice Date:</strong> ${new Date().toLocaleDateString("en-CA")}<br/>
          <strong>Payment Mode:</strong> ${escapeHtml(payMode)}
        </div>
      </div>
      <div class="grid" style="margin-top:8px;">
        <div>
          <strong>Bill To:</strong> ${escapeHtml(job.name)}<br/>
          <strong>Contact:</strong> ${escapeHtml(job.contactNumber || "-")}<br/>
          <strong>Address:</strong> ${escapeHtml(job.address || "-")}
        </div>
        <div>
          <strong>Model:</strong> ${escapeHtml(`${job.brand || ""} ${job.model || ""}`.trim())}<br/>
          <strong>IMEI:</strong> ${escapeHtml(job.imei || "-")}<br/>
          <strong>Status:</strong> ${escapeHtml(job.status || "-")}
        </div>
      </div>
      <table>
        <tr>
          <th>Sno</th><th>Item Description</th><th>HSN/SAC</th><th class="right">Qty</th><th>Unit</th><th class="right">Rate Per Item</th><th class="right">Taxable Value</th><th class="right">CGST %</th><th class="right">CGST Amt</th><th class="right">SGST %</th><th class="right">SGST Amt</th><th class="right">Total (Incl.GST)</th>
        </tr>
        ${rows.map((r, i) => `<tr><td>${i + 1}</td><td>${escapeHtml(r.desc)}</td><td>${escapeHtml(r.hsn)}</td><td class="right">${r.qty}</td><td>${escapeHtml(r.unit)}</td><td class="right">${r.rate.toFixed(2)}</td><td class="right">${r.taxable.toFixed(2)}</td><td class="right">${r.cgstRate.toFixed(2)}</td><td class="right">${r.cgstAmt.toFixed(2)}</td><td class="right">${r.sgstRate.toFixed(2)}</td><td class="right">${r.sgstAmt.toFixed(2)}</td><td class="right">${r.total.toFixed(2)}</td></tr>`).join("")}
        <tr><th colspan="6" class="right">Total</th><th class="right">${taxableTotal.toFixed(2)}</th><th></th><th class="right">${cgstTotal.toFixed(2)}</th><th></th><th class="right">${sgstTotal.toFixed(2)}</th><th class="right">${grand.toFixed(2)}</th></tr>
        <tr><th colspan="11" class="right">Advance</th><th class="right">-${advance.toFixed(2)}</th></tr>
        <tr><th colspan="11" class="right">Discount</th><th class="right">-${discount.toFixed(2)}</th></tr>
        <tr><th colspan="11" class="right">Net Payable</th><th class="right">${netPayable.toFixed(2)}</th></tr>
      </table>
      <p class="small"><strong>Total Invoice Value (In Words):</strong> ${money(netPayable)}</p>
      <p class="small"><strong>Terms and Conditions:</strong> ${escapeHtml(invoiceTerms)}</p>
      <p class="small"><strong>${escapeHtml(txtTpl.signatureCustomerLabel)}:</strong> ____________________ &nbsp;&nbsp; <strong>${escapeHtml(txtTpl.signatureStaffLabel)}:</strong> ____________________</p>
    </div>
  </body></html>`;
  const w = window.open("", "_blank");
  if (!w) return;
  w.document.write(html);
  w.document.close();
  w.focus();
  w.print();
}

function renderEngineerJobSelect() {
  if (!el.engineerJobSelect) return;
  const options = state.jobs
    .slice()
    .reverse()
    .map((j) => `<option value="${j.id}">${escapeHtml(j.jobNumber)} - ${escapeHtml(j.name)} (${escapeHtml(j.status)})</option>`)
    .join("");

  el.engineerJobSelect.innerHTML = options || '<option value="">No jobs</option>';
}

function renderEngineerJobsList() {
  if (!el.engineerJobsBody) return;
  const rows = [...state.jobs].slice().reverse();
  el.engineerJobsBody.innerHTML = rows.length
    ? rows
        .map(
          (j) => `
      <tr>
        <td>${escapeHtml(j.jobNumber)}</td>
        <td>${escapeHtml(j.name)}</td>
        <td>${escapeHtml(j.contactNumber)}</td>
        <td>${escapeHtml(j.status)}</td>
        <td><button type="button" class="engOpenBtn" data-id="${j.id}">Open</button></td>
      </tr>`
        )
        .join("")
    : '<tr><td colspan="5">No jobs</td></tr>';

  Array.from(document.querySelectorAll(".engOpenBtn")).forEach((btn) => {
    btn.addEventListener("click", () => {
      selectedEngineerJobId = btn.dataset.id;
      engineerSelectedParts = [];
      renderSelectedParts();
      renderEngineerMeta();
    });
  });
}

function renderEngineerMeta() {
  const job = state.jobs.find((j) => j.id === selectedEngineerJobId);
  if (!job) {
    el.engineerJobMeta.innerHTML = "<p>Select a job from list to open and edit.</p>";
    return;
  }

  const updatesCount = (job.engineerUpdates || []).length;
  const timelineCount = (job.timeline || []).length;
  if (el.engCustName) el.engCustName.value = job.name || "";
  if (el.engCustAddress) el.engCustAddress.value = job.address || "";
  if (el.engCustContact) el.engCustContact.value = job.contactNumber || "";
  if (el.engCustAlt) el.engCustAlt.value = job.altNumber || "";
  if (el.engCustEmail) el.engCustEmail.value = job.email || "";
  if (el.engCustModel) el.engCustModel.value = job.model || "";
  if (el.engCustImei) el.engCustImei.value = job.imei || "";
  if (el.engCustComplaint) el.engCustComplaint.value = job.complaint || "";
  el.engineerWarranty.value = job.warrantyCategory || el.engineerWarranty.value;
  el.repairStatus.value = job.status || el.repairStatus.value;
  if (job.assignedEngineer) el.engineerName.value = job.assignedEngineer;
  if (el.serviceCharge) el.serviceCharge.value = Number(job.serviceCharge || 0) || "";
  if (el.discountAmount) el.discountAmount.value = Number(job.discountAmount || 0) || "";
  if (el.autoCollectionAmount) el.autoCollectionAmount.value = getNetBillAmount(job).toFixed(2);
  if (el.closeCollectionMode) el.closeCollectionMode.value = job.closeCollection?.mode || "Cash";
  toggleEngineerPaymentField();
  applyEngineerEditPermissions(job);
  refreshAutoCollectionPreview();
  const partsTotal = getPartsTotalFromJob(job);
  const serviceBase = getServiceChargeBase(job);
  const serviceTax = getServiceChargeTax(serviceBase);
  const serviceTotal = getServiceChargeGross(serviceBase);
  const netBill = getNetBillAmount(job);

  el.engineerJobMeta.innerHTML = `
    <p><strong>Job:</strong> ${escapeHtml(job.jobNumber)} | <strong>Name:</strong> ${escapeHtml(job.name)} | <strong>Warranty:</strong> ${escapeHtml(job.warrantyCategory)}</p>
    <p><strong>Device:</strong> ${escapeHtml(job.brand)} ${escapeHtml(job.model)} | <strong>IMEI/SN:</strong> ${escapeHtml(job.imei)}</p>
    <p><strong>Latest Status:</strong> ${escapeHtml(job.status)}</p>
    <p><strong>Parts Total:</strong> ${money(partsTotal)} | <strong>Service:</strong> ${job.warrantyCategory === "OOW" ? `${money(serviceBase)} + GST ${money(serviceTax)} = ${money(serviceTotal)}` : "-"}</p>
    <p><strong>Advance:</strong> ${money(job.advanceAmount || 0)} | <strong>Discount:</strong> ${money(job.discountAmount || 0)} | <strong>Net Bill:</strong> ${money(netBill)}</p>
    <p><strong>Estimate:</strong> ${job.estimate ? `${money(job.estimate.amount)} (${escapeHtml(job.estimate.mode)})` : "-"} | <strong>Invoice:</strong> ${job.invoice ? `${escapeHtml(job.invoice.invoiceNo)} (${money(job.invoice.amount)})` : "-"}</p>
    <p><strong>Handover Remark:</strong> ${escapeHtml(job.handoverRemark || "-")}</p>
    <p><strong>Timeline Count:</strong> ${timelineCount} | <strong>Engineer Log Count:</strong> ${updatesCount}</p>
  `;
}

function applyEngineerEditPermissions(job) {
  const user = getCurrentUser();
  const isReception = user?.role === "receptionist";
  const isAdmin = user?.role === "admin";
  const isEngineer = user?.role === "engineer";
  const lockClosed = job.status === "Delivered" && !isAdmin;

  const customerFields = [el.engCustName, el.engCustAddress, el.engCustContact, el.engCustAlt, el.engCustEmail, el.engCustModel, el.engCustImei, el.engCustComplaint];
  const engineerFields = [el.engineerName, el.engineerWarranty, el.paymentAmount, el.partSearch, el.engineerPhotos, el.engineerZip];
  customerFields.forEach((f) => { if (f) f.disabled = lockClosed || isReception; });
  engineerFields.forEach((f) => { if (f) f.disabled = lockClosed || isReception; });

  if (el.addPartBtn) el.addPartBtn.disabled = lockClosed || isReception;
  if (el.partSearch) el.partSearch.classList.toggle("hidden", isReception);
  if (el.partSuggestionBox) el.partSuggestionBox.classList.toggle("hidden", true);
  if (el.addPartBtn) el.addPartBtn.classList.toggle("hidden", isReception);
  if (el.selectedParts) el.selectedParts.classList.toggle("hidden", isReception);
  if (el.engineerPhotos) el.engineerPhotos.classList.toggle("hidden", isReception);
  if (el.engineerZip) el.engineerZip.classList.toggle("hidden", isReception);
  if (el.printEstimateBtn) {
    el.printEstimateBtn.classList.toggle("hidden", isReception);
    el.printEstimateBtn.disabled = lockClosed || isReception;
  }
  if (el.engineerName) el.engineerName.classList.toggle("hidden", isReception);
  if (el.engineerWarranty) el.engineerWarranty.classList.toggle("hidden", isReception);
  if (el.paymentAmount) {
    el.paymentAmount.classList.add("hidden");
    el.paymentAmount.disabled = true;
    el.paymentAmount.value = "";
  }
  if (el.discountAmount) {
    const show = isReception || isAdmin || isEngineer;
    el.discountAmount.classList.toggle("hidden", !show);
    el.discountAmount.disabled = lockClosed || !show;
  }
  if (el.serviceCharge) {
    const show = job.warrantyCategory === "OOW";
    const canEdit = (isEngineer || isAdmin) && !isReception;
    el.serviceCharge.classList.toggle("hidden", !show);
    el.serviceCharge.disabled = lockClosed || !show || !canEdit;
    if (!show) el.serviceCharge.value = "";
  }
  if (el.autoCollectionAmount) {
    const show = isReception || isAdmin;
    el.autoCollectionAmount.classList.toggle("hidden", !show);
    el.autoCollectionAmount.disabled = true;
    if (job.warrantyCategory !== "OOW") el.autoCollectionAmount.value = "";
  }
  if (el.closeCollectionMode) {
    const show = isReception || isAdmin;
    el.closeCollectionMode.classList.toggle("hidden", !show);
    el.closeCollectionMode.disabled = lockClosed || !show || job.warrantyCategory !== "OOW";
  }
  if (el.partQtyInput) {
    el.partQtyInput.classList.toggle("hidden", isReception);
    el.partQtyInput.disabled = lockClosed || isReception;
  }
  if (el.handoverRemark) el.handoverRemark.classList.toggle("hidden", !isReception);
  const remarkField = el.engineerUpdateForm?.querySelector('textarea[name="remark"]');
  if (remarkField) remarkField.classList.toggle("hidden", isReception);
  if (el.repairStatus) {
    if (isReception) {
      el.repairStatus.innerHTML = `<option>Delivered</option><option>Cancelled</option>`;
      if (!["Delivered", "Cancelled"].includes(el.repairStatus.value)) el.repairStatus.value = "Delivered";
      el.repairStatus.disabled = lockClosed;
    } else {
      fillSelect(el.repairStatus, state.masters.repairStatuses);
      el.repairStatus.value = job.status || el.repairStatus.value;
      el.repairStatus.disabled = lockClosed;
    }
  }

  const submitBtn = el.engineerUpdateForm?.querySelector('button[type="submit"]');
  if (submitBtn) {
    submitBtn.disabled = lockClosed;
    if (lockClosed) submitBtn.textContent = "Closed Job (Read Only)";
    else submitBtn.textContent = isReception ? "Update Status" : "Save Engineer Update";
  }

  if (el.printEstimateBtn) {
    const allowPrint = (isEngineer || isAdmin) && job.warrantyCategory === "OOW";
    el.printEstimateBtn.classList.toggle("hidden", !allowPrint);
  }
}

function openHistory(type) {
  const job = state.jobs.find((j) => j.id === selectedEngineerJobId);
  if (!job || !el.historyDialog || !el.historyList || !el.historyTitle) return;

  if (type === "timeline") {
    el.historyTitle.textContent = "Job Timeline History";
    const rows = (job.timeline || []).slice().reverse();
    el.historyList.innerHTML = rows.length
      ? rows.map((r) => `<li>${new Date(r.at).toLocaleString()} | ${escapeHtml(r.by || "-")} | ${escapeHtml(r.note || "")}</li>`).join("")
      : "<li>No timeline records.</li>";
  } else {
    el.historyTitle.textContent = "Engineer Action History";
    const rows = (job.engineerUpdates || []).slice().reverse();
    el.historyList.innerHTML = rows.length
      ? rows.map((u) => `<li>${new Date(u.createdAt).toLocaleString()} | ${escapeHtml(u.engineerName || "-")} | ${escapeHtml(u.repairStatus || "-")} | ${escapeHtml(u.remark || "")}${u.partsUsed?.length ? ` | Parts: ${escapeHtml(formatParts(u.partsUsed))}` : ""}</li>`).join("")
      : "<li>No engineer records.</li>";
  }
  el.historyDialog.showModal();
}

function toggleEngineerPaymentField() {
  const w = el.engineerWarranty.value;
  const isOow = w === "OOW";
  const user = getCurrentUser();
  if (el.paymentAmount) {
    el.paymentAmount.disabled = !isOow;
    if (!isOow) el.paymentAmount.value = "";
  }
  if (el.serviceCharge) {
    const canEdit = user && ["engineer", "admin"].includes(user.role);
    el.serviceCharge.disabled = !isOow || !canEdit;
    el.serviceCharge.classList.toggle("hidden", !isOow);
    if (!isOow) el.serviceCharge.value = "";
  }
  if (el.autoCollectionAmount) {
    el.autoCollectionAmount.disabled = true;
    if (!isOow) el.autoCollectionAmount.value = "";
  }
  if (el.closeCollectionMode) {
    const canEditCollection = user && ["receptionist", "admin"].includes(user.role);
    el.closeCollectionMode.disabled = !isOow || !canEditCollection;
  }
}

function refreshAutoCollectionPreview() {
  if (!el.autoCollectionAmount) return;
  const job = state.jobs.find((j) => j.id === selectedEngineerJobId);
  if (!job) return;
  const previewPartsTotal = engineerSelectedParts.length
    ? engineerSelectedParts.reduce((sum, p) => sum + Number(p.qty || 0) * Number(p.price || 0), 0)
    : getPartsTotalFromJob(job);
  const previewServiceBase = Number(el.serviceCharge?.value || job.serviceCharge || 0);
  const discount = Number(el.discountAmount?.value || job.discountAmount || 0);
  const net = getNetBillAmount(job, {
    partsTotal: previewPartsTotal,
    serviceChargeBase: Number.isFinite(previewServiceBase) ? previewServiceBase : 0,
    discountAmount: Number.isFinite(discount) ? discount : Number(job.discountAmount || 0),
    warrantyCategory: String(el.engineerWarranty?.value || job.warrantyCategory || "")
  });
  el.autoCollectionAmount.value = net.toFixed(2);
}

function renderSelectedParts() {
  el.selectedParts.innerHTML = engineerSelectedParts
    .map((p, i) => {
      const code = p.partCode ? `[${p.partCode}] ` : "";
      return `<span class="chip">${escapeHtml(code + p.name)} x${p.qty} @ ${money(p.price)} <button type="button" class="removePart" data-idx="${i}">x</button></span>`;
    })
    .join("");

  Array.from(document.querySelectorAll(".removePart")).forEach((b) => {
    b.addEventListener("click", () => {
      const idx = Number(b.dataset.idx);
      engineerSelectedParts = engineerSelectedParts.filter((_, i) => i !== idx);
      renderSelectedParts();
    });
  });
  refreshAutoCollectionPreview();
}

function addMasterValue(key, value) {
  const v = String(value || "").trim();
  if (!v) return;
  if (!Array.isArray(state.masters[key])) return;
  if (state.masters[key].some((x) => x.toLowerCase() === v.toLowerCase())) return;
  state.masters[key].push(v);
}

function renderMasterLists() {
  const keys = [
    "warrantyCategories",
    "brands",
    "customerTypes",
    "productConditions",
    "invoiceModes",
    "engineers",
    "repairStatuses",
    "parts"
  ];

  keys.forEach((key) => {
    const host = document.getElementById(`master-${key}`);
    if (!host) return;
    host.innerHTML = state.masters[key]
      .map((v) => `<li>${escapeHtml(v)} <button type="button" class="deleteMaster" data-key="${key}" data-value="${escapeHtml(v)}">Delete</button></li>`)
      .join("");
  });

  Array.from(document.querySelectorAll(".deleteMaster")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.dataset.key;
      const value = btn.dataset.value;
      state.masters[key] = state.masters[key].filter((x) => x !== value);
      afterChange();
    });
  });
}

function renderUsers() {
  el.usersBody.innerHTML = state.users
    .map((u) => `
      <tr>
        <td>${escapeHtml(u.fullName || "-")}</td>
        <td>${escapeHtml(u.profileTitle || "-")}</td>
        <td>${escapeHtml(u.username)}</td>
        <td>${escapeHtml(u.joiningDate || "-")}</td>
        <td>${escapeHtml(u.age || "-")}</td>
        <td>${escapeHtml(u.role)}</td>
        <td>${u.active ? "Active" : "Disabled"}</td>
        <td>
          <button type="button" class="editUserBtn action-btn" data-id="${u.id}">Edit</button>
          <button type="button" class="toggleUser" data-id="${u.id}">${u.active ? "Disable" : "Enable"}</button>
        </td>
      </tr>`)
    .join("");

  Array.from(document.querySelectorAll(".editUserBtn")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const user = state.users.find((u) => u.id === btn.dataset.id);
      if (!user || !el.userEditDialog || !el.userEditForm) return;
      selectedUserEditId = user.id;
      el.userEditForm.fullName.value = user.fullName || "";
      el.userEditForm.profileTitle.value = user.profileTitle || "";
      el.userEditForm.username.value = user.username || "";
      el.userEditForm.joiningDate.value = user.joiningDate || "";
      el.userEditForm.age.value = user.age || "";
      el.userEditForm.role.value = user.role || "receptionist";
      el.userEditForm.newPassword.value = "";
      el.userEditDialog.showModal();
    });
  });

  Array.from(document.querySelectorAll(".toggleUser")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const user = state.users.find((u) => u.id === btn.dataset.id);
      if (!user) return;
      if (user.username === "admin" && user.active) {
        alert("Default admin cannot be disabled.");
        return;
      }
      user.active = !user.active;
      afterChange();
    });
  });
}

function timeToMinutes(value) {
  const raw = String(value || "");
  if (!/^\d{2}:\d{2}$/.test(raw)) return null;
  const [h, m] = raw.split(":").map(Number);
  if (!Number.isFinite(h) || !Number.isFinite(m)) return null;
  if (h < 0 || h > 23 || m < 0 || m > 59) return null;
  return h * 60 + m;
}

function getAttendancePolicy() {
  const p = state.settings?.attendancePolicy || {};
  return {
    checkInStart: String(p.checkInStart || "08:00"),
    checkInEnd: String(p.checkInEnd || "10:30"),
    checkOutStart: String(p.checkOutStart || "16:00"),
    checkOutEnd: String(p.checkOutEnd || "23:30")
  };
}

function attendanceTimeAllowed(action, timeValue, policy = getAttendancePolicy()) {
  const value = timeToMinutes(timeValue);
  if (value === null) return false;
  if (action === "checkIn") {
    const start = timeToMinutes(policy.checkInStart);
    const end = timeToMinutes(policy.checkInEnd);
    if (start === null || end === null) return true;
    return value >= start && value <= end;
  }
  const start = timeToMinutes(policy.checkOutStart);
  const end = timeToMinutes(policy.checkOutEnd);
  if (start === null || end === null) return true;
  return value >= start && value <= end;
}

function mergeAttendanceNotes(...parts) {
  return parts
    .map((p) => String(p || "").trim())
    .filter(Boolean)
    .join(" | ");
}

function hasPendingAttendanceRequest(userId, date, action) {
  return (state.attendanceRequests || []).some((r) => r.userId === userId && r.date === date && r.action === action && r.status === "Pending");
}

function queueAttendanceRequest({ userId, date, action, requestedTime, remark, by }) {
  state.attendanceRequests = state.attendanceRequests || [];
  state.attendanceRequests.push({
    id: uid("ATR"),
    userId,
    date,
    action,
    requestedTime,
    remark: String(remark || "").trim(),
    status: "Pending",
    requestedAt: nowIso(),
    reviewedAt: "",
    reviewedBy: "",
    adminRemark: "",
    by: String(by || "")
  });
}

function applyAttendanceRequest(req, adminUser, adminRemark) {
  const existing = state.attendance.find((r) => r.userId === req.userId && r.date === req.date);
  const note = mergeAttendanceNotes(
    existing?.note || "",
    req.remark ? `${req.action === "checkIn" ? "Check-in" : "Check-out"} remark: ${req.remark}` : "",
    adminRemark ? `Admin remark: ${adminRemark}` : ""
  );

  if (req.action === "checkIn") {
    upsertAttendanceRow({
      userId: req.userId,
      date: req.date,
      checkIn: req.requestedTime || existing?.checkIn || "",
      checkOut: existing?.checkOut || "",
      status: existing?.status || "Present",
      note,
      by: adminUser
    });
    return;
  }

  upsertAttendanceRow({
    userId: req.userId,
    date: req.date,
    checkIn: existing?.checkIn || req.requestedTime || "",
    checkOut: req.requestedTime || existing?.checkOut || "",
    status: existing?.status || "Present",
    note,
    by: adminUser
  });
}

function getAttendanceRange() {
  const period = String(el.attendancePeriod?.value || "monthly");
  if (period === "custom") {
    return {
      from: el.attendanceFrom?.value || "",
      to: el.attendanceTo?.value || ""
    };
  }
  if (period === "yearly") {
    const y = Number(el.attendanceYear?.value || new Date().getFullYear());
    return { from: `${y}-01-01`, to: `${y}-12-31` };
  }
  const month = String(el.attendanceMonth?.value || ymd().slice(0, 7));
  return { from: `${month}-01`, to: `${month}-31` };
}

function workedHours(checkIn, checkOut) {
  if (!checkIn || !checkOut) return "-";
  const [ih, im] = String(checkIn).split(":").map(Number);
  const [oh, om] = String(checkOut).split(":").map(Number);
  if (![ih, im, oh, om].every((v) => Number.isFinite(v))) return "-";
  const start = ih * 60 + im;
  const end = oh * 60 + om;
  if (end <= start) return "-";
  const mins = end - start;
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  return `${h}h ${m}m`;
}

function upsertAttendanceRow(row) {
  const idx = state.attendance.findIndex((r) => r.userId === row.userId && r.date === row.date);
  if (idx >= 0) {
    state.attendance[idx] = { ...state.attendance[idx], ...row, updatedAt: nowIso() };
  } else {
    state.attendance.push({ id: uid("ATT"), ...row, updatedAt: nowIso() });
  }
}

function filteredAttendanceRows() {
  const user = getCurrentUser();
  if (!user) return [];
  const selectedUser = String(el.attendanceStaffFilter?.value || "");
  const q = String(el.attendanceSearch?.value || "").trim().toLowerCase();
  const range = getAttendanceRange();
  let rows = [...(state.attendance || [])];

  if (user.role !== "admin") rows = rows.filter((r) => r.userId === user.id);
  if (user.role === "admin" && selectedUser) rows = rows.filter((r) => r.userId === selectedUser);
  if (range.from) rows = rows.filter((r) => r.date >= range.from);
  if (range.to) rows = rows.filter((r) => r.date <= range.to);
  if (q) {
    rows = rows.filter((r) => {
      const u = state.users.find((x) => x.id === r.userId);
      const txt = `${u?.fullName || ""} ${u?.username || ""} ${u?.role || ""} ${r.note || ""} ${r.status || ""}`.toLowerCase();
      return txt.includes(q);
    });
  }

  return rows.sort((a, b) => {
    if (a.date === b.date) return (b.checkIn || "").localeCompare(a.checkIn || "");
    return b.date.localeCompare(a.date);
  });
}

function toggleAttendancePeriodInputs() {
  if (!el.attendancePeriod) return;
  const period = el.attendancePeriod.value;
  const isCustom = period === "custom";
  const isYear = period === "yearly";
  if (el.attendanceMonth) el.attendanceMonth.classList.toggle("hidden", isCustom || isYear);
  if (el.attendanceYear) el.attendanceYear.classList.toggle("hidden", isCustom || !isYear);
  if (el.attendanceFrom) el.attendanceFrom.classList.toggle("hidden", !isCustom);
  if (el.attendanceTo) el.attendanceTo.classList.toggle("hidden", !isCustom);
}

function renderAttendance() {
  if (!el.attendanceBody) return;
  const user = getCurrentUser();
  if (!user) return;

  if (el.attendanceYear && !el.attendanceYear.value) el.attendanceYear.value = String(new Date().getFullYear());
  if (el.attendanceMonth && !el.attendanceMonth.value) el.attendanceMonth.value = ymd().slice(0, 7);
  toggleAttendancePeriodInputs();

  const staff = state.users.filter((u) => u.active);
  if (el.attendanceStaffFilter) {
    const previous = el.attendanceStaffFilter.value;
    const options = [`<option value="">All Staff</option>`]
      .concat(staff.map((s) => `<option value="${s.id}">${escapeHtml(s.fullName || s.username)} (${escapeHtml(s.role)})</option>`));
    el.attendanceStaffFilter.innerHTML = options.join("");
    if (user.role !== "admin") {
      el.attendanceStaffFilter.value = user.id;
      el.attendanceStaffFilter.disabled = true;
    } else {
      el.attendanceStaffFilter.disabled = false;
      el.attendanceStaffFilter.value = previous || "";
    }
  }

  if (el.attendanceAdminUser) {
    el.attendanceAdminUser.innerHTML = staff.map((s) => `<option value="${s.id}">${escapeHtml(s.fullName || s.username)} (${escapeHtml(s.role)})</option>`).join("");
    if (!el.attendanceAdminUser.value) el.attendanceAdminUser.value = user.id;
  }

  const policy = getAttendancePolicy();
  if (el.attendancePolicyForm) {
    el.attendancePolicyForm.classList.toggle("hidden", user.role !== "admin");
    if (user.role === "admin") {
      el.attendancePolicyForm.checkInStart.value = policy.checkInStart;
      el.attendancePolicyForm.checkInEnd.value = policy.checkInEnd;
      el.attendancePolicyForm.checkOutStart.value = policy.checkOutStart;
      el.attendancePolicyForm.checkOutEnd.value = policy.checkOutEnd;
    }
  }
  if (el.attendanceRulesText) {
    el.attendanceRulesText.textContent = `Allowed Check-In: ${policy.checkInStart} to ${policy.checkInEnd} | Allowed Check-Out: ${policy.checkOutStart} to ${policy.checkOutEnd}`;
  }

  if (el.attendanceAdminForm) el.attendanceAdminForm.classList.toggle("hidden", user.role !== "admin");
  if (el.attendanceCheckInBtn) el.attendanceCheckInBtn.classList.toggle("hidden", user.role === "admin");
  if (el.attendanceCheckOutBtn) el.attendanceCheckOutBtn.classList.toggle("hidden", user.role === "admin");
  if (el.attendanceActionRemark) el.attendanceActionRemark.classList.toggle("hidden", user.role === "admin");

  const todayRow = state.attendance.find((r) => r.userId === user.id && r.date === ymd());
  const todayPendingReq = (state.attendanceRequests || []).filter((r) => r.userId === user.id && r.date === ymd() && r.status === "Pending").length;
  if (el.attendanceTodayText) {
    el.attendanceTodayText.textContent = todayRow
      ? `Today: ${todayRow.checkIn || "-"} - ${todayRow.checkOut || "-"} (${todayRow.status || "Present"})${todayPendingReq ? ` | Pending approval: ${todayPendingReq}` : ""}`
      : `Today: Not marked${todayPendingReq ? ` | Pending approval: ${todayPendingReq}` : ""}`;
  }

  const rows = filteredAttendanceRows();
  el.attendanceBody.innerHTML = rows.length
    ? rows
        .map((r) => {
          const u = state.users.find((x) => x.id === r.userId);
          return `<tr>
            <td>${escapeHtml(r.date)}</td>
            <td>${escapeHtml(u?.fullName || u?.username || "-")}</td>
            <td>${escapeHtml(u?.role || "-")}</td>
            <td>${escapeHtml(r.checkIn || "-")}</td>
            <td>${escapeHtml(r.checkOut || "-")}</td>
            <td>${escapeHtml(workedHours(r.checkIn, r.checkOut))}</td>
            <td>${escapeHtml(r.status || "-")}</td>
            <td>${escapeHtml(r.note || "-")}</td>
            <td>${escapeHtml(r.by || "-")}</td>
          </tr>`;
        })
        .join("")
    : '<tr><td colspan="9">No attendance records.</td></tr>';

  if (!el.attendanceRequestBody || !el.attendanceRequestWrap) return;
  const reqRows = (state.attendanceRequests || [])
    .filter((r) => (user.role === "admin" ? true : r.userId === user.id))
    .sort((a, b) => new Date(b.requestedAt).getTime() - new Date(a.requestedAt).getTime());
  if (el.attendanceRequestTitle) {
    el.attendanceRequestTitle.classList.toggle("hidden", user.role !== "admin" && !reqRows.length);
  }
  el.attendanceRequestWrap.classList.toggle("hidden", user.role !== "admin" && !reqRows.length);
  el.attendanceRequestBody.innerHTML = reqRows.length
    ? reqRows
        .map((r) => {
          const u = state.users.find((x) => x.id === r.userId);
          const pending = r.status === "Pending";
          return `<tr>
            <td>${new Date(r.requestedAt).toLocaleString()}</td>
            <td>${escapeHtml(r.date)}</td>
            <td>${escapeHtml(u?.fullName || u?.username || "-")}</td>
            <td>${escapeHtml(r.action === "checkIn" ? "Check In" : "Check Out")}</td>
            <td>${escapeHtml(r.requestedTime || "-")}</td>
            <td>${escapeHtml(r.remark || "-")}</td>
            <td>${escapeHtml(r.status)}</td>
            <td>${user.role === "admin" && pending ? `<input class="att-admin-remark action-input" data-id="${r.id}" placeholder="Admin remark" />` : escapeHtml(r.adminRemark || "-")}</td>
            <td>${
              user.role === "admin" && pending
                ? `<button type="button" class="approveAttendanceReq action-btn" data-id="${r.id}">Approve</button> <button type="button" class="rejectAttendanceReq secondary action-btn" data-id="${r.id}">Reject</button>`
                : "-"
            }</td>
          </tr>`;
        })
        .join("")
    : '<tr><td colspan="9">No attendance requests.</td></tr>';

  if (user.role === "admin") {
    Array.from(document.querySelectorAll(".approveAttendanceReq")).forEach((btn) => {
      btn.addEventListener("click", () => {
        const req = (state.attendanceRequests || []).find((r) => r.id === btn.dataset.id);
        if (!req || req.status !== "Pending") return;
        const remarkInput = document.querySelector(`.att-admin-remark[data-id="${req.id}"]`);
        const adminRemark = String(remarkInput?.value || "").trim();
        req.status = "Approved";
        req.reviewedAt = nowIso();
        req.reviewedBy = user.username;
        req.adminRemark = adminRemark;
        applyAttendanceRequest(req, user.username, adminRemark);
        afterChange();
      });
    });
    Array.from(document.querySelectorAll(".rejectAttendanceReq")).forEach((btn) => {
      btn.addEventListener("click", () => {
        const req = (state.attendanceRequests || []).find((r) => r.id === btn.dataset.id);
        if (!req || req.status !== "Pending") return;
        const remarkInput = document.querySelector(`.att-admin-remark[data-id="${req.id}"]`);
        const adminRemark = String(remarkInput?.value || "").trim();
        req.status = "Rejected";
        req.reviewedAt = nowIso();
        req.reviewedBy = user.username;
        req.adminRemark = adminRemark;
        afterChange();
      });
    });
  }
}

function exportAttendanceCsv() {
  const rows = filteredAttendanceRows();
  const header = ["Date", "Staff", "Username", "Role", "CheckIn", "CheckOut", "Worked", "Status", "Note", "UpdatedBy"];
  const lines = [header.join(",")];
  rows.forEach((r) => {
    const u = state.users.find((x) => x.id === r.userId);
    lines.push(
      [
        r.date || "",
        u?.fullName || "",
        u?.username || "",
        u?.role || "",
        r.checkIn || "",
        r.checkOut || "",
        workedHours(r.checkIn, r.checkOut),
        r.status || "",
        r.note || "",
        r.by || ""
      ]
        .map(csvEscape)
        .join(",")
    );
  });
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `attendance-${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

function getJobFieldValue(job, field) {
  return job?.[field] ?? "";
}

function applyApprovedRequest(req) {
  const job = state.jobs.find((j) => j.id === req.jobId);
  if (!job) return;
  if (req.field === "contactNumber" || req.field === "altNumber") {
    if (!/^\d{10}$/.test(req.newValue)) return;
  }
  job[req.field] = req.newValue;
  job.timeline = job.timeline || [];
  job.timeline.push({
    at: nowIso(),
    by: req.reviewedBy,
    note: `Admin approved change: ${req.field} -> ${req.newValue}`
  });
}

function renderApprovals() {
  if (!el.approvalsBody) return;
  const rows = state.requests.slice().reverse();
  el.approvalsBody.innerHTML = rows.length
    ? rows
        .map((r) => {
          const job = state.jobs.find((j) => j.id === r.jobId);
          const pending = r.status === "Pending";
          return `
            <tr>
              <td>${new Date(r.requestedAt).toLocaleString()}</td>
              <td>${escapeHtml(job?.jobNumber || "-")}</td>
              <td>${escapeHtml(r.requestedBy)}</td>
              <td>${escapeHtml(r.field)}</td>
              <td>${escapeHtml(r.currentValue)}</td>
              <td>${escapeHtml(r.newValue)}</td>
              <td>${escapeHtml(r.status)}</td>
              <td>
                ${pending ? `<button class="approveReq" data-id="${r.id}" type="button">Approve</button> <button class="rejectReq secondary" data-id="${r.id}" type="button">Reject</button>` : "-"}
              </td>
            </tr>`;
        })
        .join("")
    : '<tr><td colspan="8">No requests.</td></tr>';

  Array.from(document.querySelectorAll(".approveReq")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const req = state.requests.find((x) => x.id === btn.dataset.id);
      if (!req || req.status !== "Pending") return;
      req.status = "Approved";
      req.reviewedAt = nowIso();
      req.reviewedBy = getCurrentUser()?.username || "admin";
      applyApprovedRequest(req);
      afterChange();
    });
  });

  Array.from(document.querySelectorAll(".rejectReq")).forEach((btn) => {
    btn.addEventListener("click", () => {
      const req = state.requests.find((x) => x.id === btn.dataset.id);
      if (!req || req.status !== "Pending") return;
      req.status = "Rejected";
      req.reviewedAt = nowIso();
      req.reviewedBy = getCurrentUser()?.username || "admin";
      afterChange();
    });
  });
}

function collectFileMeta(fileList) {
  return Array.from(fileList || []).map((f) => ({
    name: f.name,
    type: f.type || "",
    size: f.size || 0
  }));
}

async function collectFilesWithData(fileList) {
  const files = Array.from(fileList || []);
  return Promise.all(
    files.map(
      (f) =>
        new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () =>
            resolve({
              id: uid("FILE"),
              name: f.name,
              type: f.type || "",
              size: f.size || 0,
              dataUrl: String(reader.result || "")
            });
          reader.onerror = reject;
          reader.readAsDataURL(f);
        })
    )
  );
}

function downloadAttachment(att) {
  if (!att?.dataUrl) {
    alert("Old file metadata only. Re-upload new file to enable download.");
    return;
  }
  const a = document.createElement("a");
  a.href = att.dataUrl;
  a.download = att.name || "file";
  a.click();
}

function csvEscape(v) {
  const s = String(v ?? "");
  if (s.includes(",") || s.includes("\n") || s.includes('"')) return `"${s.replaceAll('"', '""')}"`;
  return s;
}

function exportJobsCsv() {
  const rows = filteredJobsForView();
  const header = ["JobNo", "Date", "CustomerType", "Name", "Contact", "AltNumber", "Warranty", "Brand", "Model", "IMEI", "Status", "Engineer", "Advance", "ServiceCharge", "ServiceGST", "ServiceTotal", "Discount", "Total", "NetBill", "Paid"];
  const lines = [header.join(",")];
  rows.forEach((j) => {
    const paid = (j.payments || []).reduce((sum, p) => sum + Number(p.amount || 0), 0);
    lines.push(
      [
        j.jobNumber,
        new Date(j.createdAt).toLocaleString(),
        j.customerType || "Customer",
        j.name,
        j.contactNumber,
        j.altNumber || "",
        j.warrantyCategory,
        j.brand,
        j.model,
        j.imei,
        j.status,
        j.assignedEngineer || "",
        j.advanceAmount || 0,
        getServiceChargeBase(j),
        getServiceChargeTax(j),
        getServiceChargeGross(j),
        j.discountAmount || 0,
        j.totalAmount || 0,
        getNetBillAmount(j),
        paid
      ]
        .map(csvEscape)
        .join(",")
    );
  });
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `jobsheet-export-${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

function exportPriceCsv() {
  const rows = filteredPriceRows();
  const header = ["itemCode", "partName", "brand", "model", "price", "updatedAt", "updatedBy"];
  const lines = [header.join(",")];
  rows.forEach((p) => {
    lines.push(
      [
        p.itemCode || "",
        p.partName || "",
        p.brand || "",
        p.model || "",
        p.price || 0,
        p.updatedAt || "",
        p.updatedBy || ""
      ]
        .map(csvEscape)
        .join(",")
    );
  });
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `price-list-${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

function exportLedgerCsv() {
  const rows = filteredLedgerRows();
  const header = ["Date", "Type", "Category", "Amount", "Mode", "JobNo", "Note", "By"];
  const lines = [header.join(",")];
  rows.forEach((r) => {
    const job = state.jobs.find((j) => j.id === r.jobId);
    lines.push(
      [
        new Date(r.at).toLocaleString(),
        r.type || "",
        r.category || "",
        r.amount || 0,
        r.mode || "",
        job?.jobNumber || "",
        r.note || "",
        r.by || ""
      ]
        .map(csvEscape)
        .join(",")
    );
  });
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `cash-ledger-${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

function exportAccountsCsv() {
  const rows = filteredAccountsRows();
  const header = ["Date", "JobNo", "Name", "Contact", "Advance", "Discount", "NetBill", "Collected", "Mode", "CollectedBy"];
  const lines = [header.join(",")];
  rows.forEach((r) => {
    lines.push(
      [
        new Date(r.col.at).toLocaleString(),
        r.job.jobNumber || "",
        r.job.name || "",
        r.job.contactNumber || "",
        r.job.advanceAmount || 0,
        r.job.discountAmount || 0,
        getNetBillAmount(r.job),
        r.col.amount || 0,
        r.col.mode || "",
        r.col.by || ""
      ]
        .map(csvEscape)
        .join(",")
    );
  });
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `accounts-${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

function filteredAccountsRows() {
  const q = (el.accountsSearch?.value || "").trim().toLowerCase();
  const from = el.accountsFrom?.value || "";
  const to = el.accountsTo?.value || "";

  let rows = (state.jobs || [])
    .filter((j) => j.warrantyCategory === "OOW" && j.closeCollection && j.status === "Delivered")
    .map((j) => ({ job: j, col: j.closeCollection }));

  if (q) {
    rows = rows.filter((r) => `${r.job.jobNumber} ${r.job.name} ${r.job.contactNumber}`.toLowerCase().includes(q));
  }
  if (from) {
    const f = new Date(`${from}T00:00:00`).getTime();
    rows = rows.filter((r) => new Date(r.col.at).getTime() >= f);
  }
  if (to) {
    const t = new Date(`${to}T23:59:59`).getTime();
    rows = rows.filter((r) => new Date(r.col.at).getTime() <= t);
  }
  return rows.sort((a, b) => new Date(b.col.at).getTime() - new Date(a.col.at).getTime());
}

function renderAccounts() {
  if (!el.accountsBody) return;
  const rows = filteredAccountsRows();
  const total = rows.reduce((sum, r) => sum + Number(r.col.amount || 0), 0);
  el.accountsBody.innerHTML = rows.length
    ? rows
        .map(
          (r) => `
      <tr>
        <td>${new Date(r.col.at).toLocaleString()}</td>
        <td>${escapeHtml(r.job.jobNumber)}</td>
        <td>${escapeHtml(r.job.name)}</td>
        <td>${money(r.job.advanceAmount || 0)}</td>
        <td>${money(r.job.discountAmount || 0)}</td>
        <td>${money(getNetBillAmount(r.job))}</td>
        <td>${money(r.col.amount)}</td>
        <td>${escapeHtml(r.col.mode || "-")}</td>
        <td>${escapeHtml(r.col.by || "-")}</td>
      </tr>`
        )
        .join("") + `<tr><td colspan="6"><strong>Total Collected</strong></td><td><strong>${money(total)}</strong></td><td colspan="2"></td></tr>`
    : '<tr><td colspan="9">No OOW collections found.</td></tr>';
}

function bindKpiCards() {
  document.querySelectorAll(".kpi-click").forEach((card) => {
    card.addEventListener("click", () => {
      const page = card.getAttribute("data-open-page") || "jobsheets";
      const kpiFrom = String(el.kpiDateFrom?.value || "");
      const kpiTo = String(el.kpiDateTo?.value || "");
      quickWarrantyFilter = "";
      quickOpenOnly = false;
      if (el.jobStatusFilter) el.jobStatusFilter.value = "";
      if (el.jobSearch) el.jobSearch.value = "";
      if (el.jobDateFrom) el.jobDateFrom.value = kpiFrom;
      if (el.jobDateTo) el.jobDateTo.value = kpiTo;

      const status = card.getAttribute("data-open-status");
      const warranty = card.getAttribute("data-open-warranty");
      if (status === "!closed") {
        quickOpenOnly = true;
      } else if (status) {
        if (el.jobStatusFilter) el.jobStatusFilter.value = status;
      }
      if (warranty) quickWarrantyFilter = warranty;

      switchPage(page);
      if (page === "jobsheets") renderJobsTable();
    });
  });
}

function normalizeImportedJobs(rows) {
  return rows
    .map((j) => ({
      id: j.id || uid("JOB"),
      jobNumber: j.jobNumber || j.JobNo || j.job_no || formatJobNumber(state.jobCounter + 1),
      createdAt: j.createdAt || j.Date || nowIso(),
      name: j.name || j.Name || "",
      customerType: j.customerType || j.CustomerType || "Customer",
      address: j.address || j.Address || "",
      contactNumber: String(j.contactNumber || j.Contact || "").trim(),
      altNumber: String(j.altNumber || j.AltNumber || "").trim(),
      email: j.email || j.Email || "",
      warrantyCategory: j.warrantyCategory || j.Warranty || "IW",
      brand: j.brand || j.Brand || "",
      model: j.model || j.Model || "",
      imei: j.imei || j.IMEI || "",
      complaint: j.complaint || j.Complaint || "",
      advanceAmount: Number(j.advanceAmount || j.Advance || 0),
      serviceCharge: Number(j.serviceCharge || j.ServiceCharge || 0),
      discountAmount: Number(j.discountAmount || j.Discount || 0),
      totalAmount: Number(j.totalAmount || j.Total || 0),
      invoiceMode: j.invoiceMode || j.InvoiceMode || "",
      productCondition: j.productCondition || j.ProductCondition || "",
      attachments: Array.isArray(j.attachments) ? j.attachments : [],
      assignedEngineer: j.assignedEngineer || j.Engineer || "",
      status: j.status || j.Status || "Received",
      engineerUpdates: Array.isArray(j.engineerUpdates) ? j.engineerUpdates : [],
      payments: Array.isArray(j.payments) ? j.payments : [],
      closeCollection: j.closeCollection || null,
      timeline: Array.isArray(j.timeline) ? j.timeline : [],
      updatedAt: j.updatedAt || j.modifiedAt || j.createdAt || j.Date || nowIso()
    }))
    .filter((j) => j.name || j.contactNumber || j.jobNumber);
}

function parseCsvToObjects(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.trim().length);
  if (lines.length < 2) return [];
  const headers = lines[0].split(",").map((h) => h.trim());
  return lines.slice(1).map((line) => {
    const cols = line.split(",");
    const o = {};
    headers.forEach((h, i) => { o[h] = (cols[i] || "").trim(); });
    return o;
  });
}

function normalizeTemplateKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function getTemplateValue(map, keys, fallback = "") {
  for (const key of keys) {
    const normalized = normalizeTemplateKey(key);
    if (Object.prototype.hasOwnProperty.call(map, normalized)) {
      const value = String(map[normalized] ?? "").trim().replace(/\\n/g, "\n");
      if (value) return value;
    }
  }
  return fallback;
}

function parseTemplateRowsToMap(rows) {
  const map = {};
  (rows || []).forEach((row) => {
    if (!row || typeof row !== "object") return;
    const values = Object.values(row);
    const rawKey =
      row.key ?? row.Key ?? row.KEY ??
      row.field ?? row.Field ?? row.FIELD ??
      row.name ?? row.Name ?? row.NAME ??
      values[0] ?? "";
    const rawValue =
      row.value ?? row.Value ?? row.VALUE ??
      row.text ?? row.Text ??
      row.content ?? row.Content ??
      values[1] ?? "";
    const key = normalizeTemplateKey(rawKey);
    if (!key) return;
    map[key] = String(rawValue ?? "");
  });
  return map;
}

async function parseTemplateFileToMap(file) {
  const name = String(file?.name || "").toLowerCase();
  let rows = [];
  if (name.endsWith(".csv")) {
    const txt = await file.text();
    rows = parseCsvToObjects(txt);
  } else if (name.endsWith(".xls") || name.endsWith(".xlsx")) {
    rows = await parseExcelRowsWithXlsx(file);
  } else {
    throw new Error("Only .csv, .xls, .xlsx files supported.");
  }
  const map = parseTemplateRowsToMap(rows);
  if (!Object.keys(map).length) {
    throw new Error("No key/value rows found. Use sample template format.");
  }
  return map;
}

function applyTemplateUpload(templateType, map) {
  const prevTpl = getJobSheetTemplate();
  const prevText = getPrintText();
  const prevStyle = getPrintStyle();
  const nextTpl = { ...prevTpl };
  const nextText = { ...prevText };
  const nextStyle = { ...prevStyle };

  const companyName = getTemplateValue(map, ["companyName", "billFrom", "serviceCenterName"], nextTpl.companyName);
  const phone = getTemplateValue(map, ["phone", "phoneNo", "servicePhone", "contactNumber"], nextTpl.phone);
  const email = getTemplateValue(map, ["email", "mail", "serviceEmail"], nextTpl.email);
  const address = getTemplateValue(map, ["address", "serviceAddress"], nextTpl.address);
  const fontFamily = getTemplateValue(map, ["fontFamily", "font"], nextStyle.fontFamily);
  const baseSize = Number(getTemplateValue(map, ["baseSize", "baseFontSize"], String(nextStyle.baseSize)));
  const headingSize = Number(getTemplateValue(map, ["headingSize", "titleSize"], String(nextStyle.headingSize)));

  nextTpl.companyName = companyName;
  nextTpl.phone = phone;
  nextTpl.email = email;
  nextTpl.address = address;
  nextStyle.fontFamily = fontFamily || nextStyle.fontFamily;
  if (Number.isFinite(baseSize) && baseSize >= 10 && baseSize <= 22) nextStyle.baseSize = baseSize;
  if (Number.isFinite(headingSize) && headingSize >= 14 && headingSize <= 42) nextStyle.headingSize = headingSize;

  if (templateType === "jobsheet") {
    nextTpl.headerNote = getTemplateValue(map, ["headerNote", "note"], nextTpl.headerNote);
    nextTpl.footerNote = getTemplateValue(map, ["footerNote", "footer", "footerTerms"], nextTpl.footerNote);
    nextText.jobSheetTitle = getTemplateValue(map, ["jobSheetTitle", "title"], nextText.jobSheetTitle);
    nextText.jobTermsHeading = getTemplateValue(map, ["jobTermsHeading", "termsHeading"], nextText.jobTermsHeading);
    nextText.jobTerms = getTemplateValue(map, ["jobTerms", "terms", "termsAndConditions"], nextText.jobTerms);
    nextText.signatureCustomerLabel = getTemplateValue(map, ["signatureCustomerLabel", "customerSignatureLabel"], nextText.signatureCustomerLabel);
    nextText.signatureStaffLabel = getTemplateValue(map, ["signatureStaffLabel", "staffSignatureLabel"], nextText.signatureStaffLabel);
  }

  if (templateType === "estimate") {
    nextText.estimateTitle = getTemplateValue(map, ["estimateTitle", "title"], nextText.estimateTitle);
    nextText.estimateIntro = getTemplateValue(map, ["estimateIntro", "intro", "introText"], nextText.estimateIntro);
    nextText.estimateNote = getTemplateValue(map, ["estimateNote", "note"], nextText.estimateNote);
    nextText.estimateTerms = getTemplateValue(map, ["estimateTerms", "terms", "termsAndConditions"], nextText.estimateTerms);
    nextText.signatureCustomerLabel = getTemplateValue(map, ["signatureCustomerLabel", "customerSignatureLabel"], nextText.signatureCustomerLabel);
    nextText.signatureStaffLabel = getTemplateValue(map, ["signatureStaffLabel", "staffSignatureLabel"], nextText.signatureStaffLabel);
  }

  if (templateType === "invoice") {
    nextText.taxInvoiceTitle = getTemplateValue(map, ["taxInvoiceTitle", "invoiceTitle", "title"], nextText.taxInvoiceTitle);
    nextText.taxTerms = getTemplateValue(map, ["taxTerms", "taxTerm"], nextText.taxTerms);
    nextText.invoiceTerms = getTemplateValue(map, ["invoiceTerms", "terms", "termsAndConditions"], nextText.invoiceTerms);
    nextText.signatureCustomerLabel = getTemplateValue(map, ["signatureCustomerLabel", "customerSignatureLabel"], nextText.signatureCustomerLabel);
    nextText.signatureStaffLabel = getTemplateValue(map, ["signatureStaffLabel", "staffSignatureLabel"], nextText.signatureStaffLabel);
  }

  state.settings.jobSheetTemplate = nextTpl;
  state.settings.printText = nextText;
  state.settings.printStyle = nextStyle;
}

function csvCell(value) {
  const str = String(value ?? "");
  if (!/[",\n]/.test(str)) return str;
  return `"${str.replace(/"/g, "\"\"")}"`;
}

function createTemplateSampleCsv(templateType) {
  const rows = [["key", "value"]];
  if (templateType === "jobsheet") {
    rows.push(
      ["jobSheetTitle", "realme Service Center Handover Report"],
      ["companyName", "realme Service Center"],
      ["phone", "9070606190"],
      ["email", "service@realme.com"],
      ["address", "Manjeri, Kerala"],
      ["jobTermsHeading", "Please read the following terms carefully before signing."],
      ["jobTerms", "1. Data backup responsibility is customer.\\n2. Device must be collected within 60 days.\\n3. Jobsheet is mandatory for delivery."],
      ["signatureCustomerLabel", "Customer signature"],
      ["signatureStaffLabel", "Service staff signature"]
    );
  } else if (templateType === "estimate") {
    rows.push(
      ["estimateTitle", "Estimate"],
      ["estimateIntro", "This is with reference to Job no {JOB_NO}. Complaint: {COMPLAINT}."],
      ["estimateNote", "Actual cost may vary after final diagnosis."],
      ["estimateTerms", "Estimate is valid for 7 days.\\nParts once ordered cannot be returned."],
      ["companyName", "realme Service Center"],
      ["phone", "9070606190"],
      ["address", "Manjeri, Kerala"]
    );
  } else if (templateType === "invoice") {
    rows.push(
      ["taxInvoiceTitle", "Tax Invoice"],
      ["taxTerms", "GST included as per applicable law."],
      ["invoiceTerms", "No cash refund after delivery.\\nWarranty on service as per policy."],
      ["companyName", "realme Service Center"],
      ["phone", "9070606190"],
      ["address", "Manjeri, Kerala"]
    );
  }
  return `${rows.map((r) => `${csvCell(r[0])},${csvCell(r[1])}`).join("\n")}\n`;
}

function downloadTemplateSample(templateType, fileName) {
  const csv = createTemplateSampleCsv(templateType);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(a.href);
}

function afterChange(options = {}) {
  saveState();
  if (!options.skipCloudSave) scheduleCloudSave();
  syncCurrentUserReference();
  renderUi(Boolean(options.skipRegisterReset));
}

el.loginForm.addEventListener("submit", (e) => {
  e.preventDefault();
  const fd = new FormData(el.loginForm);
  const username = String(fd.get("username") || "").trim().toLowerCase();
  const password = String(fd.get("password") || "");

  const user = state.users.find((u) => String(u.username || "").trim().toLowerCase() === username && u.password === password && u.active);
  if (!user) {
    alert("Invalid username or password.");
    return;
  }

  currentUser = { id: user.id, username: String(user.username || "").trim().toLowerCase() };
  saveSession();
  el.loginForm.reset();
  updateRoleUi();
  afterChange();
  pullCloudUpdates(true);
});

if (el.openForgotPasswordBtn) {
  el.openForgotPasswordBtn.addEventListener("click", () => {
    el.forgotPasswordForm?.reset();
    el.forgotPasswordDialog?.showModal();
  });
}

if (el.closeForgotPasswordDialog) {
  el.closeForgotPasswordDialog.addEventListener("click", () => {
    el.forgotPasswordForm?.reset();
    el.forgotPasswordDialog?.close();
  });
}

if (el.forgotPasswordDialog) {
  el.forgotPasswordDialog.addEventListener("close", () => {
    el.forgotPasswordForm?.reset();
  });
}

if (el.forgotPasswordForm) {
  el.forgotPasswordForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const fd = new FormData(el.forgotPasswordForm);
    const username = String(fd.get("username") || "").trim();
    const newPassword = String(fd.get("newPassword") || "");
    const confirmPassword = String(fd.get("confirmPassword") || "");
    if (!username) return alert("Enter username.");
    const target = findUserByUsername(username);
    if (!target) return alert("Username not found.");
    if (newPassword.length < 4) return alert("Password must be at least 4 characters.");
    if (newPassword !== confirmPassword) return alert("Password confirm mismatch.");
    target.password = newPassword;
    afterChange();
    el.forgotPasswordDialog?.close();
    alert("Password updated. Please login.");
  });
}

if (el.resetLoginDataBtn) {
  el.resetLoginDataBtn.addEventListener("click", () => {
    const ok = confirm("Reset all local CRM data and restore default login?");
    if (!ok) return;
    localStorage.removeItem(STORE_KEY);
    localStorage.removeItem(SESSION_KEY);
    location.reload();
  });
}

if (el.openProfileBtn) {
  el.openProfileBtn.addEventListener("click", () => {
    renderProfileCard();
    el.profilePasswordForm?.classList.add("hidden");
    if (el.togglePasswordPanelBtn) el.togglePasswordPanelBtn.textContent = "Change Password";
    el.profilePasswordForm?.reset();
    el.profileDialog?.showModal();
  });
}

if (el.closeProfileDialog) {
  el.closeProfileDialog.addEventListener("click", () => {
    el.profilePasswordForm?.classList.add("hidden");
    if (el.togglePasswordPanelBtn) el.togglePasswordPanelBtn.textContent = "Change Password";
    el.profilePasswordForm?.reset();
    el.profileDialog?.close();
  });
}

if (el.profileDialog) {
  el.profileDialog.addEventListener("close", () => {
    el.profilePasswordForm?.classList.add("hidden");
    if (el.togglePasswordPanelBtn) el.togglePasswordPanelBtn.textContent = "Change Password";
    el.profilePasswordForm?.reset();
  });
}

if (el.togglePasswordPanelBtn) {
  el.togglePasswordPanelBtn.addEventListener("click", () => {
    const isHidden = el.profilePasswordForm?.classList.contains("hidden");
    if (!el.profilePasswordForm) return;
    if (isHidden) {
      el.profilePasswordForm.classList.remove("hidden");
      el.togglePasswordPanelBtn.textContent = "Hide Password Panel";
    } else {
      el.profilePasswordForm.classList.add("hidden");
      el.profilePasswordForm.reset();
      el.togglePasswordPanelBtn.textContent = "Change Password";
    }
  });
}

if (el.profilePasswordForm) {
  el.profilePasswordForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user) return;
    const fd = new FormData(el.profilePasswordForm);
    const currentPassword = String(fd.get("currentPassword") || "");
    const newPassword = String(fd.get("newPassword") || "");
    const confirmPassword = String(fd.get("confirmPassword") || "");
    if (currentPassword !== user.password) return alert("Current password is wrong.");
    if (newPassword.length < 4) return alert("New password must be at least 4 characters.");
    if (newPassword !== confirmPassword) return alert("Password confirm mismatch.");
    user.password = newPassword;
    el.profilePasswordForm.reset();
    el.profilePasswordForm.classList.add("hidden");
    if (el.togglePasswordPanelBtn) el.togglePasswordPanelBtn.textContent = "Change Password";
    afterChange();
    alert("Password updated.");
  });
}

if (el.logoutBtn) {
  el.logoutBtn.addEventListener("click", () => {
    currentUser = null;
    selectedEngineerJobId = null;
    selectedUserEditId = null;
    engineerSelectedParts = [];
    engineerFocusMode = false;
    saveSession();
    el.profilePasswordForm?.classList.add("hidden");
    if (el.togglePasswordPanelBtn) el.togglePasswordPanelBtn.textContent = "Change Password";
    el.profileDialog?.close();
    updateRoleUi();
  });
}

el.pageTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    if (tab.classList.contains("hidden")) return;
    switchPage(tab.dataset.page);
  });
});

el.warrantyCategory.addEventListener("change", updateConditionalFields);
if (el.jobForm?.contactNumber) {
  el.jobForm.contactNumber.addEventListener("input", updateDuplicateHint);
  el.jobForm.contactNumber.addEventListener("focus", updateDuplicateHint);
}
if (el.closeDuplicatePanel) {
  el.closeDuplicatePanel.addEventListener("click", closeDuplicatePanel);
}

el.jobForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const user = getCurrentUser();
  if (!user || !["receptionist", "admin"].includes(user.role)) {
    alert("Only Receptionist/Admin can register jobs.");
    return;
  }

  const fd = new FormData(el.jobForm);
  const contactNumber = String(fd.get("contactNumber") || "").trim();
  const altNumber = String(fd.get("altNumber") || "").trim();

  if (!validateTenDigit(contactNumber)) {
    alert("Contact Number must be exactly 10 digits.");
    return;
  }
  if (!validateTenDigit(altNumber)) {
    alert("Alternative Number must be exactly 10 digits.");
    return;
  }

  const warrantyCategory = String(fd.get("warrantyCategory") || "");
  const isOow = warrantyCategory === "OOW";

  const uploadedFiles = await collectFilesWithData(el.attachments.files);

  const next = nextJobNumber(state.jobCounter);
  state.jobCounter = next.counter;
  const createdAt = nowIso();
  const job = {
    id: uid("JOB"),
    jobNumber: next.jobNumber,
    createdAt,
    updatedAt: createdAt,
    customerType: String(fd.get("customerType") || "Customer"),
    name: String(fd.get("name") || "").trim(),
    address: String(fd.get("address") || "").trim(),
    contactNumber,
    altNumber,
    email: String(fd.get("email") || "").trim(),
    warrantyCategory,
    brand: String(fd.get("brand") || ""),
    model: String(fd.get("model") || "").trim(),
    imei: String(fd.get("imei") || "").trim(),
    complaint: String(fd.get("complaint") || "").trim(),
    advanceAmount: isOow ? Number(fd.get("advanceAmount") || 0) : 0,
    serviceCharge: 0,
    discountAmount: 0,
    totalAmount: isOow ? Number(fd.get("totalAmount") || 0) : 0,
    invoiceMode: ["IW", "DOA"].includes(warrantyCategory) ? String(fd.get("invoiceMode") || "") : "",
    productCondition: String(fd.get("productCondition") || ""),
    attachments: uploadedFiles,
    assignedEngineer: "",
    status: "Received",
    engineerUpdates: [],
    payments: [],
    timeline: [{ at: createdAt, by: user.username, note: "Job created by reception." }]
  };

  state.jobs.push(job);
  el.jobForm.reset();
  if (el.duplicateHint) {
    el.duplicateHint.classList.add("hidden");
    el.duplicateHint.innerHTML = "";
  }
  closeDuplicatePanel();
  resetRegisterHeader();
  updateConditionalFields();
  afterChange();
  switchPage("jobsheets");
});

if (el.applyFilterBtn) el.applyFilterBtn.addEventListener("click", renderJobsTable);
if (el.clearFilterBtn) {
  el.clearFilterBtn.addEventListener("click", () => {
    if (el.jobSearch) el.jobSearch.value = "";
    if (el.jobDateFrom) el.jobDateFrom.value = "";
    if (el.jobDateTo) el.jobDateTo.value = "";
    if (el.jobStatusFilter) el.jobStatusFilter.value = "";
    quickWarrantyFilter = "";
    quickOpenOnly = false;
    renderJobsTable();
  });
}
if (el.exportJobsBtn) el.exportJobsBtn.addEventListener("click", exportJobsCsv);
if (el.kpiApplyBtn) {
  el.kpiApplyBtn.addEventListener("click", () => {
    renderKpis();
  });
}
if (el.kpiClearBtn) {
  el.kpiClearBtn.addEventListener("click", () => {
    if (el.kpiDateFrom) el.kpiDateFrom.value = "";
    if (el.kpiDateTo) el.kpiDateTo.value = "";
    renderKpis();
  });
}
if (el.accountsApply) el.accountsApply.addEventListener("click", renderAccounts);
if (el.accountsExportBtn) el.accountsExportBtn.addEventListener("click", exportAccountsCsv);
if (el.accountsClear) {
  el.accountsClear.addEventListener("click", () => {
    if (el.accountsSearch) el.accountsSearch.value = "";
    if (el.accountsFrom) el.accountsFrom.value = "";
    if (el.accountsTo) el.accountsTo.value = "";
    renderAccounts();
  });
}
if (el.priceApplyBtn) el.priceApplyBtn.addEventListener("click", renderPriceList);
if (el.priceClearBtn) {
  el.priceClearBtn.addEventListener("click", () => {
    if (el.priceSearch) el.priceSearch.value = "";
    renderPriceList();
  });
}
if (el.priceExportBtn) el.priceExportBtn.addEventListener("click", exportPriceCsv);
if (el.attendancePeriod) {
  el.attendancePeriod.addEventListener("change", () => {
    toggleAttendancePeriodInputs();
    renderAttendance();
  });
}
if (el.attendanceApplyBtn) el.attendanceApplyBtn.addEventListener("click", renderAttendance);
if (el.attendanceClearBtn) {
  el.attendanceClearBtn.addEventListener("click", () => {
    if (el.attendanceSearch) el.attendanceSearch.value = "";
    if (el.attendancePeriod) el.attendancePeriod.value = "monthly";
    if (el.attendanceMonth) el.attendanceMonth.value = ymd().slice(0, 7);
    if (el.attendanceYear) el.attendanceYear.value = String(new Date().getFullYear());
    if (el.attendanceFrom) el.attendanceFrom.value = "";
    if (el.attendanceTo) el.attendanceTo.value = "";
    toggleAttendancePeriodInputs();
    renderAttendance();
  });
}
if (el.attendanceExportBtn) el.attendanceExportBtn.addEventListener("click", exportAttendanceCsv);
if (el.attendanceStaffFilter) el.attendanceStaffFilter.addEventListener("change", renderAttendance);
if (el.attendanceSearch) el.attendanceSearch.addEventListener("input", renderAttendance);
if (el.attendanceMonth) el.attendanceMonth.addEventListener("change", renderAttendance);
if (el.attendanceYear) el.attendanceYear.addEventListener("change", renderAttendance);
if (el.attendanceFrom) el.attendanceFrom.addEventListener("change", renderAttendance);
if (el.attendanceTo) el.attendanceTo.addEventListener("change", renderAttendance);

if (el.attendanceCheckInBtn) {
  el.attendanceCheckInBtn.addEventListener("click", () => {
    const user = getCurrentUser();
    if (!user) return;
    const date = ymd();
    const checkTime = hm();
    const existing = state.attendance.find((r) => r.userId === user.id && r.date === date);
    const remark = String(el.attendanceActionRemark?.value || "").trim();
    if (!attendanceTimeAllowed("checkIn", checkTime)) {
      if (!remark) return alert("Outside allowed check-in time. Please add remark and try again.");
      if (hasPendingAttendanceRequest(user.id, date, "checkIn")) return alert("Check-in request already pending for today.");
      queueAttendanceRequest({
        userId: user.id,
        date,
        action: "checkIn",
        requestedTime: checkTime,
        remark,
        by: user.username
      });
      if (el.attendanceActionRemark) el.attendanceActionRemark.value = "";
      afterChange();
      alert("Check-in request sent to admin approval.");
      return;
    }
    upsertAttendanceRow({
      userId: user.id,
      date,
      checkIn: checkTime,
      checkOut: existing?.checkOut || "",
      status: existing?.status || "Present",
      note: mergeAttendanceNotes(existing?.note || "", remark),
      by: user.username
    });
    if (el.attendanceActionRemark) el.attendanceActionRemark.value = "";
    afterChange();
  });
}

if (el.attendanceCheckOutBtn) {
  el.attendanceCheckOutBtn.addEventListener("click", () => {
    const user = getCurrentUser();
    if (!user) return;
    const date = ymd();
    const checkTime = hm();
    const existing = state.attendance.find((r) => r.userId === user.id && r.date === date);
    const remark = String(el.attendanceActionRemark?.value || "").trim();
    if (!attendanceTimeAllowed("checkOut", checkTime)) {
      if (!remark) return alert("Outside allowed check-out time. Please add remark and try again.");
      if (hasPendingAttendanceRequest(user.id, date, "checkOut")) return alert("Check-out request already pending for today.");
      queueAttendanceRequest({
        userId: user.id,
        date,
        action: "checkOut",
        requestedTime: checkTime,
        remark,
        by: user.username
      });
      if (el.attendanceActionRemark) el.attendanceActionRemark.value = "";
      afterChange();
      alert("Check-out request sent to admin approval.");
      return;
    }
    upsertAttendanceRow({
      userId: user.id,
      date,
      checkIn: existing?.checkIn || checkTime,
      checkOut: checkTime,
      status: existing?.status || "Present",
      note: mergeAttendanceNotes(existing?.note || "", remark),
      by: user.username
    });
    if (el.attendanceActionRemark) el.attendanceActionRemark.value = "";
    afterChange();
  });
}

if (el.attendancePolicyForm) {
  el.attendancePolicyForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const fd = new FormData(el.attendancePolicyForm);
    const next = {
      checkInStart: String(fd.get("checkInStart") || "").trim(),
      checkInEnd: String(fd.get("checkInEnd") || "").trim(),
      checkOutStart: String(fd.get("checkOutStart") || "").trim(),
      checkOutEnd: String(fd.get("checkOutEnd") || "").trim()
    };
    if ([next.checkInStart, next.checkInEnd, next.checkOutStart, next.checkOutEnd].some((v) => !/^\d{2}:\d{2}$/.test(v))) {
      return alert("Enter valid attendance policy times.");
    }
    state.settings.attendancePolicy = next;
    afterChange();
  });
}

if (el.attendanceAdminForm) {
  el.attendanceAdminForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const fd = new FormData(el.attendanceAdminForm);
    const userId = String(fd.get("userId") || "");
    const date = String(fd.get("date") || "").slice(0, 10);
    if (!userId || !date) return alert("Select staff and date.");
    upsertAttendanceRow({
      userId,
      date,
      checkIn: String(fd.get("checkIn") || ""),
      checkOut: String(fd.get("checkOut") || ""),
      status: String(fd.get("status") || "Present"),
      note: String(fd.get("note") || "").trim(),
      by: user.username
    });
    el.attendanceAdminForm.reset();
    afterChange();
  });
}

if (el.priceAddForm) {
  el.priceAddForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const fd = new FormData(el.priceAddForm);
    const ok = upsertPriceItem({
      itemCode: String(fd.get("itemCode") || "").trim(),
      partName: String(fd.get("partName") || "").trim(),
      brand: String(fd.get("brand") || "").trim(),
      model: String(fd.get("model") || "").trim(),
      price: Number(fd.get("price") || 0)
    });
    if (!ok) return alert("Part Name is required.");
    el.priceAddForm.reset();
    afterChange();
  });
}
if (el.priceImportForm) {
  el.priceImportForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const file = el.priceImportFile?.files?.[0];
    if (!file) return alert("Select an Excel/CSV file.");
    const name = file.name.toLowerCase();
    let rows = [];
    try {
      if (name.endsWith(".csv")) {
        const txt = await file.text();
        rows = parseCsvToObjects(txt);
      } else if (name.endsWith(".xls") || name.endsWith(".xlsx")) {
        rows = await parseExcelRowsWithXlsx(file);
      } else {
        return alert("Only .csv, .xls, .xlsx files supported.");
      }
    } catch (_err) {
      return alert("Failed to read price file.");
    }
    let count = 0;
    rows.forEach((r) => {
      const changed = upsertPriceItem({
        itemCode: r.itemCode || r.code || r.partCode || "",
        partName: r.partName || r.part || r.name || "",
        brand: r.brand || "",
        model: r.model || "",
        price: r.price || r.rate || r.amount || 0
      });
      if (changed) count += 1;
    });
    el.priceImportForm.reset();
    afterChange();
    alert(`Imported/updated ${count} price items.`);
  });
}
if (el.downloadPriceTemplate) {
  el.downloadPriceTemplate.addEventListener("click", () => {
    const header = ["itemCode", "partName", "brand", "model", "price"];
    const sample = ["DSP-001", "Display", "Realme", "Realme 13 Pro+", "2500"];
    const csv = `${header.join(",")}\n${sample.join(",")}`;
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "price-list-template.csv";
    a.click();
    URL.revokeObjectURL(a.href);
  });
}

if (el.jobSheetTemplateUploadForm) {
  el.jobSheetTemplateUploadForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const file = el.jobSheetTemplateFile?.files?.[0];
    if (!file) return alert("Select jobsheet template file.");
    try {
      const map = await parseTemplateFileToMap(file);
      applyTemplateUpload("jobsheet", map);
      el.jobSheetTemplateUploadForm.reset();
      afterChange();
      alert("Jobsheet template uploaded successfully.");
    } catch (err) {
      alert(err?.message || "Failed to upload jobsheet template.");
    }
  });
}

if (el.estimateTemplateUploadForm) {
  el.estimateTemplateUploadForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const file = el.estimateTemplateFile?.files?.[0];
    if (!file) return alert("Select estimate template file.");
    try {
      const map = await parseTemplateFileToMap(file);
      applyTemplateUpload("estimate", map);
      el.estimateTemplateUploadForm.reset();
      afterChange();
      alert("Estimate template uploaded successfully.");
    } catch (err) {
      alert(err?.message || "Failed to upload estimate template.");
    }
  });
}

if (el.invoiceTemplateUploadForm) {
  el.invoiceTemplateUploadForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const file = el.invoiceTemplateFile?.files?.[0];
    if (!file) return alert("Select invoice template file.");
    try {
      const map = await parseTemplateFileToMap(file);
      applyTemplateUpload("invoice", map);
      el.invoiceTemplateUploadForm.reset();
      afterChange();
      alert("Invoice template uploaded successfully.");
    } catch (err) {
      alert(err?.message || "Failed to upload invoice template.");
    }
  });
}

if (el.downloadJobSheetTemplateSample) {
  el.downloadJobSheetTemplateSample.addEventListener("click", () => {
    downloadTemplateSample("jobsheet", "jobsheet-template-upload-sample.csv");
  });
}
if (el.downloadEstimateTemplateSample) {
  el.downloadEstimateTemplateSample.addEventListener("click", () => {
    downloadTemplateSample("estimate", "estimate-template-upload-sample.csv");
  });
}
if (el.downloadInvoiceTemplateSample) {
  el.downloadInvoiceTemplateSample.addEventListener("click", () => {
    downloadTemplateSample("invoice", "invoice-template-upload-sample.csv");
  });
}

if (el.ledgerPeriod) {
  el.ledgerPeriod.addEventListener("change", () => {
    toggleLedgerDateInputs();
    renderLedger();
  });
}
if (el.ledgerApplyBtn) el.ledgerApplyBtn.addEventListener("click", renderLedger);
if (el.ledgerExportBtn) el.ledgerExportBtn.addEventListener("click", exportLedgerCsv);
if (el.ledgerClearBtn) {
  el.ledgerClearBtn.addEventListener("click", () => {
    if (el.ledgerSearch) el.ledgerSearch.value = "";
    if (el.ledgerPeriod) el.ledgerPeriod.value = "daily";
    if (el.ledgerAnchorDate) el.ledgerAnchorDate.value = new Date().toISOString().slice(0, 10);
    if (el.ledgerFromDate) el.ledgerFromDate.value = "";
    if (el.ledgerToDate) el.ledgerToDate.value = "";
    toggleLedgerDateInputs();
    renderLedger();
  });
}
if (el.ledgerExpenseForm) {
  el.ledgerExpenseForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || !["admin", "receptionist"].includes(user.role)) return;
    const fd = new FormData(el.ledgerExpenseForm);
    const amount = Number(fd.get("amount") || 0);
    if (!Number.isFinite(amount) || amount <= 0) return alert("Enter valid expense amount.");
    addLedgerEntry({
      type: "OUT",
      category: "Expense",
      amount,
      mode: String(fd.get("mode") || "Cash"),
      note: String(fd.get("note") || "").trim(),
      by: user.username
    });
    el.ledgerExpenseForm.reset();
    afterChange();
  });
}
if (el.ledgerAdminCashForm) {
  el.ledgerAdminCashForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const fd = new FormData(el.ledgerAdminCashForm);
    const amount = Number(fd.get("amount") || 0);
    if (!Number.isFinite(amount) || amount <= 0) return alert("Enter valid amount.");
    const type = String(fd.get("type") || "IN") === "OUT" ? "OUT" : "IN";
    addLedgerEntry({
      type,
      category: type === "IN" ? "Admin Cash Add" : "Admin Cash Take",
      amount,
      mode: String(fd.get("mode") || "Cash"),
      note: String(fd.get("note") || "").trim(),
      by: user.username
    });
    el.ledgerAdminCashForm.reset();
    afterChange();
  });
}

if (el.engineerBackBtn) {
  el.engineerBackBtn.addEventListener("click", () => {
    engineerFocusMode = false;
    switchPage("jobsheets");
  });
}
if (el.printEstimateBtn) {
  el.printEstimateBtn.addEventListener("click", () => {
    const job = state.jobs.find((j) => j.id === selectedEngineerJobId);
    if (!job) return alert("Open a job first.");
    const parts = engineerSelectedParts.length
      ? engineerSelectedParts
      : ((job.engineerUpdates || []).slice().reverse().find((u) => Array.isArray(u.partsUsed) && u.partsUsed.length)?.partsUsed || []);
    printEstimateInvoice(job, parts);
  });
}
if (el.openTimelineBtn) el.openTimelineBtn.addEventListener("click", () => openHistory("timeline"));
if (el.openEngineerLogBtn) el.openEngineerLogBtn.addEventListener("click", () => openHistory("engineer"));
if (el.closeHistoryDialog) el.closeHistoryDialog.addEventListener("click", () => el.historyDialog?.close());

if (el.loadEngineerJob && el.engineerJobSelect) {
  el.loadEngineerJob.addEventListener("click", () => {
    selectedEngineerJobId = el.engineerJobSelect.value || null;
    engineerSelectedParts = [];
    renderSelectedParts();
    renderEngineerMeta();
  });
}

el.engineerWarranty.addEventListener("change", () => {
  toggleEngineerPaymentField();
  refreshAutoCollectionPreview();
});

el.addPartBtn.addEventListener("click", () => {
  const val = String(el.partSearch.value || "").trim();
  if (!val) return;
  const picked = selectedPartItem || findPricePartByInput(val);
  if (!picked) {
    alert("Select a valid part from price list.");
    return;
  }
  const qty = Number(el.partQtyInput?.value || 1);
  if (!Number.isFinite(qty) || qty <= 0) return alert("Invalid quantity.");
  const price = Number(picked.price || 0);
  if (!Number.isFinite(price) || price < 0) return alert("Price not available in price list.");
  engineerSelectedParts.push({
    name: picked.partName,
    partCode: picked.itemCode || "",
    model: picked.model || "",
    brand: picked.brand || "",
    qty,
    price
  });
  el.partSearch.value = "";
  selectedPartItem = null;
  if (el.partQtyInput) el.partQtyInput.value = "1";
  clearPartSuggestions(true);
  renderSelectedParts();
  refreshAutoCollectionPreview();
});

if (el.partSearch) {
  el.partSearch.addEventListener("input", () => {
    selectedPartItem = null;
    renderPartSuggestions(el.partSearch.value);
  });
  el.partSearch.addEventListener("blur", () => {
    setTimeout(() => clearPartSuggestions(), 120);
  });
  el.partSearch.addEventListener("focus", () => {
    if (el.partSearch.value) renderPartSuggestions(el.partSearch.value);
  });
}

if (el.discountAmount) {
  el.discountAmount.addEventListener("input", refreshAutoCollectionPreview);
}
if (el.serviceCharge) {
  el.serviceCharge.addEventListener("input", refreshAutoCollectionPreview);
}

el.engineerUpdateForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const user = getCurrentUser();
  if (!user || !["engineer", "admin", "receptionist"].includes(user.role)) {
    alert("No access to update this section.");
    return;
  }

  const job = state.jobs.find((j) => j.id === selectedEngineerJobId);
  if (!job) {
    alert("Open a job first.");
    return;
  }
  if (job.status === "Delivered" && user.role !== "admin") {
    alert("Closed job cannot be edited. Admin only can edit.");
    return;
  }

  const fd = new FormData(el.engineerUpdateForm);
  const isReception = user.role === "receptionist";
  const isAdmin = user.role === "admin";
  const custName = String(fd.get("custName") || "").trim();
  const custAddress = String(fd.get("custAddress") || "").trim();
  const custContact = String(fd.get("custContact") || "").trim();
  const custAlt = String(fd.get("custAlt") || "").trim();
  const custEmail = String(fd.get("custEmail") || "").trim();
  const custModel = String(fd.get("custModel") || "").trim();
  const custImei = String(fd.get("custImei") || "").trim();
  const custComplaint = String(fd.get("custComplaint") || "").trim();
  const engineerName = String(fd.get("engineerName") || "");
  const repairStatus = String(fd.get("repairStatus") || "");
  const engineerWarranty = String(fd.get("engineerWarranty") || "");
  const serviceChargeInput = Number(fd.get("serviceCharge") || 0);
  const paymentAmount = 0;
  const discountAmount = Number(fd.get("discountAmount") || 0);
  const closeCollectionMode = String(fd.get("closeCollectionMode") || "Cash");
  const handoverRemark = String(fd.get("handoverRemark") || "").trim();
  const remark = String(fd.get("remark") || "").trim();

  if (!isReception && !validateTenDigit(custContact)) {
    alert("Contact Number must be exactly 10 digits.");
    return;
  }
  if (!isReception && !validateTenDigit(custAlt)) {
    alert("Alternative Number must be exactly 10 digits.");
    return;
  }
  if (isReception && !["Delivered", "Cancelled"].includes(repairStatus)) {
    alert("Reception can update only Delivered or Cancelled.");
    return;
  }
  if (!Number.isFinite(discountAmount) || discountAmount < 0) {
    alert("Invalid discount amount.");
    return;
  }
  if (!Number.isFinite(serviceChargeInput) || serviceChargeInput < 0) {
    alert("Invalid service charge amount.");
    return;
  }
  job.discountAmount = discountAmount;
  const serviceCharge = engineerWarranty === "OOW" ? serviceChargeInput : 0;
  if (!isReception) {
    job.serviceCharge = serviceCharge;
  }
  const canCollectionEdit = isReception || isAdmin;
  if (canCollectionEdit && repairStatus === "Delivered" && job.warrantyCategory === "OOW") {
    const autoAmount = getNetBillAmount(job);
    upsertJobCollectionLedger(job, autoAmount, closeCollectionMode, user.username);
    job.payments = (job.payments || []).filter((p) => p.source !== "CLOSE_COLLECTION");
    if (autoAmount > 0) {
      job.payments.push({
        id: uid("PAY"),
        amount: autoAmount,
        method: `Reception ${closeCollectionMode}`,
        at: nowIso(),
        source: "CLOSE_COLLECTION"
      });
    }
  }
  if (canCollectionEdit && repairStatus !== "Delivered") {
    removeJobCollectionLedger(job);
    job.payments = (job.payments || []).filter((p) => p.source !== "CLOSE_COLLECTION");
    if (job.closeCollection) {
      job.closeCollection.amount = 0;
      job.closeCollection.mode = closeCollectionMode || "Cash";
      job.closeCollection.at = nowIso();
    }
  }

  const photos = isReception ? [] : await collectFilesWithData(el.engineerPhotos.files);
  const zipFiles = isReception ? [] : await collectFilesWithData(el.engineerZip.files);

  const partsTotal = engineerSelectedParts.reduce((sum, p) => sum + Number(p.qty || 0) * Number(p.price || 0), 0);
  const update = {
    id: uid("UPD"),
    createdAt: nowIso(),
    engineerName: isReception ? (job.assignedEngineer || "") : engineerName,
    repairStatus,
    engineerWarranty: isReception ? job.warrantyCategory : engineerWarranty,
    paymentAmount: isReception ? 0 : (engineerWarranty === "OOW" ? paymentAmount : 0),
    serviceCharge,
    discountAmount: Number(job.discountAmount || 0),
    closeCollectionAmount: canCollectionEdit ? getNetBillAmount(job) : Number(job.closeCollection?.amount || 0),
    closeCollectionMode: canCollectionEdit ? closeCollectionMode : String(job.closeCollection?.mode || ""),
    handoverRemark,
    remark,
    partsUsed: isReception ? [] : [...engineerSelectedParts],
    photos,
    zipFiles
  };

  job.assignedEngineer = update.engineerName;
  job.status = repairStatus;
  job.warrantyCategory = update.engineerWarranty;
  if (!isReception) {
    job.name = custName || job.name;
    job.address = custAddress || job.address;
    job.contactNumber = custContact || job.contactNumber;
    job.altNumber = custAlt || job.altNumber;
    job.email = custEmail;
    job.model = custModel || job.model;
    job.imei = custImei || job.imei;
    job.complaint = custComplaint || job.complaint;
  }
  if (!isReception && update.engineerWarranty === "OOW" && (partsTotal > 0 || getServiceChargeBase(job) > 0)) {
    job.totalAmount = partsTotal;
    job.estimate = {
      amount: partsTotal,
      mode: "Part Estimate",
      by: user.username,
      at: nowIso()
    };
    job.timeline.push({
      at: nowIso(),
      by: user.username,
      note: `Estimate updated from parts: ${money(partsTotal)}, service: ${money(getServiceChargeBase(job))} + GST ${money(getServiceChargeTax(job))}`
    });
  }
  if (!isReception && update.engineerWarranty === "OOW" && ["Ready Pickup", "Delivered"].includes(repairStatus) && !job.invoice) {
    state.invoiceCounter = Number(state.invoiceCounter || 0) + 1;
    job.invoice = {
      invoiceNo: `INV-${new Date().getFullYear()}-${String(state.invoiceCounter).padStart(5, "0")}`,
      createdAt: nowIso(),
      amount: Number(getNetBillAmount(job))
    };
    job.timeline.push({ at: nowIso(), by: user.username, note: `Invoice generated: ${job.invoice.invoiceNo}` });
  }
  if (isReception && handoverRemark) {
    job.handoverRemark = handoverRemark;
  }
  job.engineerUpdates.push(update);
  const touchAt = nowIso();
  job.timeline.push({
    at: touchAt,
    by: user.username,
    note: isReception
      ? `Reception status update: ${repairStatus}${handoverRemark ? ` | Remark: ${handoverRemark}` : ""}${repairStatus === "Delivered" && job.warrantyCategory === "OOW" ? ` | Collected: ${money(getNetBillAmount(job))} (${closeCollectionMode})` : ""}`
      : `Engineer update: ${repairStatus}${update.engineerWarranty === "OOW" ? ` | Service: ${money(getServiceChargeBase(job))} + GST ${money(getServiceChargeTax(job))}` : ""}`
  });
  touchJob(job, touchAt);

  engineerSelectedParts = [];
  el.engineerUpdateForm.reset();
  toggleEngineerPaymentField();
  renderSelectedParts();
  afterChange();
});

if (el.adminEditForm) {
  el.adminEditForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const job = state.jobs.find((j) => j.id === selectedAdminEditJobId);
    if (!job) return;

    const fd = new FormData(el.adminEditForm);
    const contactNumber = String(fd.get("contactNumber") || "").trim();
    const altNumber = String(fd.get("altNumber") || "").trim();
    if (!validateTenDigit(contactNumber)) return alert("Contact Number must be exactly 10 digits.");
    if (!validateTenDigit(altNumber)) return alert("Alternative Number must be exactly 10 digits.");

    job.name = String(fd.get("name") || "").trim();
    job.customerType = String(fd.get("customerType") || "Customer");
    job.contactNumber = contactNumber;
    job.altNumber = altNumber;
    job.warrantyCategory = String(fd.get("warrantyCategory") || "");
    if (job.warrantyCategory !== "OOW") job.serviceCharge = 0;
    job.address = String(fd.get("address") || "").trim();
    job.email = String(fd.get("email") || "").trim();
    job.model = String(fd.get("model") || "").trim();
    job.imei = String(fd.get("imei") || "").trim();
    const touchAt = nowIso();
    job.timeline.push({ at: touchAt, by: user.username, note: "Admin edited job details directly." });
    touchJob(job, touchAt);
    el.adminEditDialog.close();
    selectedAdminEditJobId = null;
    afterChange();
  });
}

if (el.closeAdminEdit) el.closeAdminEdit.addEventListener("click", () => el.adminEditDialog.close());

el.invoicePatternForm.addEventListener("submit", (e) => {
  e.preventDefault();
  const user = getCurrentUser();
  if (!user || user.role !== "admin") return;
  const v = String(el.invoicePattern.value || "").trim();
  if (!v) return;
  state.masters.invoicePattern = v;
  afterChange();
});

el.headingForm.addEventListener("submit", (e) => {
  e.preventDefault();
  const user = getCurrentUser();
  if (!user || user.role !== "admin") return;
  const fd = new FormData(el.headingForm);
  state.settings.appTitle = String(fd.get("appTitle") || "").trim() || state.settings.appTitle;
  state.settings.loginCrmName = String(fd.get("loginCrmName") || "").trim() || state.settings.loginCrmName || state.settings.appTitle;
  state.settings.appSubtitle = String(fd.get("appSubtitle") || "").trim() || state.settings.appSubtitle;
  state.settings.dashboardHeading = String(fd.get("dashboardHeading") || "").trim() || state.settings.dashboardHeading;
  state.settings.receptionHeading = String(fd.get("receptionHeading") || "").trim() || state.settings.receptionHeading;
  state.settings.engineerHeading = String(fd.get("engineerHeading") || "").trim() || state.settings.engineerHeading;
  state.settings.jobsheetHeading = String(fd.get("jobsheetHeading") || "").trim() || state.settings.jobsheetHeading;
  state.settings.mastersHeading = String(fd.get("mastersHeading") || "").trim() || state.settings.mastersHeading;
  afterChange();
});

if (el.jobsheetTemplateForm) {
  el.jobsheetTemplateForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const fd = new FormData(el.jobsheetTemplateForm);
    state.settings.jobSheetTemplate = {
      companyName: String(fd.get("companyName") || "").trim() || "Mobile Service Center",
      phone: String(fd.get("phone") || "").trim(),
      email: String(fd.get("email") || "").trim(),
      address: String(fd.get("address") || "").trim(),
      headerNote: String(fd.get("headerNote") || "").trim(),
      footerNote: String(fd.get("footerNote") || "").trim()
    };
    afterChange();
  });
}

if (el.printStyleForm) {
  el.printStyleForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const fd = new FormData(el.printStyleForm);
    state.settings.printStyle = {
      fontFamily: String(fd.get("fontFamily") || "").trim() || "Arial",
      baseSize: Number(fd.get("baseSize") || 12),
      headingSize: Number(fd.get("headingSize") || 22)
    };
    afterChange();
  });
}

if (el.printTextForm) {
  el.printTextForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const fd = new FormData(el.printTextForm);
    state.settings.printText = {
      jobSheetTitle: String(fd.get("jobSheetTitle") || "").trim() || getPrintText().jobSheetTitle,
      jobTermsHeading: String(fd.get("jobTermsHeading") || "").trim() || getPrintText().jobTermsHeading,
      jobTerms: String(fd.get("jobTerms") || "").trim() || getPrintText().jobTerms,
      estimateTitle: String(fd.get("estimateTitle") || "").trim() || getPrintText().estimateTitle,
      estimateIntro: String(fd.get("estimateIntro") || "").trim() || getPrintText().estimateIntro,
      estimateNote: String(fd.get("estimateNote") || "").trim() || getPrintText().estimateNote,
      estimateTerms: String(fd.get("estimateTerms") || "").trim() || getPrintText().estimateTerms,
      taxInvoiceTitle: String(fd.get("taxInvoiceTitle") || "").trim() || getPrintText().taxInvoiceTitle,
      taxTerms: String(fd.get("taxTerms") || "").trim() || getPrintText().taxTerms,
      invoiceTerms: String(fd.get("invoiceTerms") || "").trim() || getPrintText().invoiceTerms,
      signatureCustomerLabel: String(fd.get("signatureCustomerLabel") || "").trim() || getPrintText().signatureCustomerLabel,
      signatureStaffLabel: String(fd.get("signatureStaffLabel") || "").trim() || getPrintText().signatureStaffLabel
    };
    afterChange();
  });
}

if (el.importJobsForm) {
  el.importJobsForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const file = el.importJobsFile?.files?.[0];
    if (!file) return alert("Select a file to import.");
    const mode = el.importMode?.value || "merge";
    const txt = await file.text();
    let importedRows = [];
    try {
      if (file.name.toLowerCase().endsWith(".json")) {
        const parsed = JSON.parse(txt);
        importedRows = Array.isArray(parsed) ? parsed : (Array.isArray(parsed.jobs) ? parsed.jobs : []);
      } else {
        importedRows = parseCsvToObjects(txt);
      }
    } catch (_err) {
      return alert("Import file invalid.");
    }

    const normalized = normalizeImportedJobs(importedRows);
    if (!normalized.length) return alert("No valid jobs found in import file.");
    if (mode === "replace") {
      state.jobs = normalized;
      state.jobCounter = normalized.length;
    } else {
      state.jobs = [...state.jobs, ...normalized];
      state.jobCounter = Math.max(state.jobCounter, state.jobs.length);
    }
    el.importJobsForm.reset();
    afterChange();
    alert(`Imported ${normalized.length} jobs successfully.`);
  });
}

if (el.downloadImportSample) {
  el.downloadImportSample.addEventListener("click", () => {
    const sample = [
      {
        jobNumber: "JOB-2026-0001",
        createdAt: "2026-04-07T10:00:00.000Z",
        name: "Sample Customer",
        customerType: "Customer",
        address: "Sample Address",
        contactNumber: "9999999999",
        altNumber: "",
        email: "sample@example.com",
        warrantyCategory: "OOW",
        brand: "Realme",
        model: "Narzo",
        imei: "123456789012345",
        complaint: "Display issue",
        advanceAmount: 0,
        serviceCharge: 0,
        totalAmount: 1200,
        invoiceMode: "",
        productCondition: "Used",
        status: "Received",
        assignedEngineer: ""
      }
    ];
    const blob = new Blob([JSON.stringify(sample, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "jobs-import-template.json";
    a.click();
    URL.revokeObjectURL(a.href);
  });
}

el.masterForms.forEach((form) => {
  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const fd = new FormData(form);
    const value = String(fd.get("value") || "").trim();
    const key = form.dataset.master;
    addMasterValue(key, value);
    form.reset();
    afterChange();
  });
});

el.userForm.addEventListener("submit", (e) => {
  e.preventDefault();
  const user = getCurrentUser();
  if (!user || user.role !== "admin") return;
  const fd = new FormData(el.userForm);
  const fullName = String(fd.get("fullName") || "").trim();
  const profileTitle = String(fd.get("profileTitle") || "").trim();
  const username = String(fd.get("username") || "").trim();
  const password = String(fd.get("password") || "").trim();
  const joiningDate = String(fd.get("joiningDate") || "").trim() || ymd();
  const age = String(fd.get("age") || "").trim();
  const role = String(fd.get("role") || "").trim();

  if (!fullName || !username || !password || !role) return;
  if (state.users.some((u) => u.username.toLowerCase() === username.toLowerCase())) {
    alert("Username already exists.");
    return;
  }

  state.users.push({ id: uid("USR"), fullName, profileTitle, username, password, joiningDate, age, role, active: true });
  el.userForm.reset();
  afterChange();
});

if (el.closeUserEditDialog) {
  el.closeUserEditDialog.addEventListener("click", () => {
    selectedUserEditId = null;
    el.userEditDialog?.close();
  });
}

if (el.userEditForm) {
  el.userEditForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const user = getCurrentUser();
    if (!user || user.role !== "admin") return;
    const target = state.users.find((u) => u.id === selectedUserEditId);
    if (!target) return;
    const fd = new FormData(el.userEditForm);
    const fullName = String(fd.get("fullName") || "").trim();
    const profileTitle = String(fd.get("profileTitle") || "").trim();
    const joiningDate = String(fd.get("joiningDate") || "").trim() || target.joiningDate || ymd();
    const age = String(fd.get("age") || "").trim();
    const role = String(fd.get("role") || target.role || "receptionist");
    const newPassword = String(fd.get("newPassword") || "").trim();
    if (!fullName) return alert("Full name required.");
    target.fullName = fullName;
    target.profileTitle = profileTitle;
    target.joiningDate = joiningDate;
    target.age = age;
    target.role = role;
    if (newPassword) target.password = newPassword;
    el.userEditDialog.close();
    selectedUserEditId = null;
    afterChange();
  });
}

async function boot() {
  bindKpiCards();
  if (el.ledgerAnchorDate && !el.ledgerAnchorDate.value) {
    el.ledgerAnchorDate.value = new Date().toISOString().slice(0, 10);
  }
  toggleLedgerDateInputs();
  await initializeCloudSync();
  currentUser = loadSession();
  afterChange({ skipCloudSave: true });
  toggleEngineerPaymentField();
  renderSelectedParts();
  updateRoleUi();
  startCloudPolling();
  pullCloudUpdates(true);
}

document.addEventListener("visibilitychange", () => {
  if (!document.hidden) pullCloudUpdates(true);
});

boot();




