const LOCAL_KEY = "tenant-rental-system-v3";
const GOOGLE_CLIENT_ID = window.RENTAL_GOOGLE_CLIENT_ID || localStorage.getItem("rental-google-client-id") || "";
const GOOGLE_SCOPES = "openid email profile https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets";
const DEFAULT_WORKSPACE = {
  mainFolderId: "1oU8FGtekTQD9G00UlIAeVZGTFuzNFNpd",
  mainFolderUrl: "https://drive.google.com/drive/folders/1oU8FGtekTQD9G00UlIAeVZGTFuzNFNpd",
  sheetId: "1Zfc0OvFh3rdHT4qywdcgvH8RpZDaFUKfp-MXkr3mc9k",
  sheetUrl: "https://docs.google.com/spreadsheets/d/1Zfc0OvFh3rdHT4qywdcgvH8RpZDaFUKfp-MXkr3mc9k/edit?usp=drivesdk",
  ownerEmail: "yu731108@gmail.com",
  folders: {
    backend: "16lepXDUrjYcVBX3kLFHqoOiIW10-2Kqb",
    leases: "1JQAFD5eyujxPEM9sq-IbmWLG5vjOF9I7",
    taipower: "12-_PQ7FlXtTYJu1rPQZ79_K1xzCXwAhh",
    meters: "10WajNa8b78ri6yo3rmn_MrR4oRJoOiuG",
    maintenance: "1JoIWcT3j2tSvwKWbL419bnUBn22GnCGq",
    receipts: "1O8ER92C17XeiiBED0A_JXBJV0xuZX0b7",
    tenants: "171J9m9AE9pmpGZZvz5K9Wry2dNCiX7JB",
  },
};
const WORKSPACE_BLUEPRINT = [
  { key: "excel", name: "01_Excel表" },
  { key: "leases", name: "02_租約照片" },
  { key: "taipower", name: "03_臺電帳單照片" },
  { key: "meters", name: "04_電表照片" },
  { key: "maintenance", name: "05_維修照片" },
  { key: "receipts", name: "06_收款憑證照片" },
  { key: "tenants", name: "07_租客照片" },
];

const MODULES = {
  dashboard: { id: "dashboard", title: "今日總覽", group: "overview" },
  rooms: { id: "rooms", title: "房間狀態", group: "overview" },
  tenants: { id: "tenants", title: "租客名單", group: "movein" },
  leases: { id: "leases", title: "租約管理", group: "movein" },
  deposits: { id: "deposits", title: "押金管理", group: "movein" },
  rent: { id: "rent", title: "房租收款", group: "finance" },
  bills: { id: "bills", title: "臺電帳單", group: "finance" },
  meters: { id: "meters", title: "電表抄表", group: "finance" },
  allocations: { id: "allocations", title: "電費分攤", group: "finance" },
  expenses: { id: "expenses", title: "支出管理", group: "finance" },
  ledger: { id: "ledger", title: "總收支明細", group: "finance" },
  maintenance: { id: "maintenance", title: "維修處理", group: "operations" },
  attachments: { id: "attachments", title: "附件中心", group: "operations" },
  reports: { id: "reports", title: "經營報表", group: "analysis" },
  settings: { id: "settings", title: "系統與儲存", group: "analysis" },
};

const NAV_GROUPS = [
  { id: "overview", label: "營運總覽", modules: ["dashboard", "rooms"] },
  { id: "movein", label: "入住與租約", modules: ["tenants", "leases", "deposits"] },
  { id: "finance", label: "收款與帳務", modules: ["rent", "bills", "meters", "allocations", "expenses", "ledger"] },
  { id: "operations", label: "維修與附件", modules: ["maintenance", "attachments"] },
  { id: "analysis", label: "報表與設定", modules: ["reports", "settings"] },
];

const roomIds = ["301", "302", "303", "304"];
const money = new Intl.NumberFormat("zh-TW", { style: "currency", currency: "TWD", maximumFractionDigits: 0 });
const percent = new Intl.NumberFormat("zh-TW", { style: "percent", maximumFractionDigits: 1 });
const defaultMonth = new Date().toISOString().slice(0, 7);

const primaryNav = document.querySelector("#primary-nav");
const secondaryNav = document.querySelector("#secondary-nav");
const appContent = document.querySelector("#app-content");
const monthSelector = document.querySelector("#month-selector");
const loginPanel = document.querySelector("#login-panel");
const recordModal = document.querySelector("#record-modal");
const recordForm = document.querySelector("#record-form");
const storedGoogleSession = loadGoogleSession();

let state = normalizeState(loadLocalState());
let googleAccessToken = storedGoogleSession.accessToken;
let googleAccessTokenExpiresAt = storedGoogleSession.expiresAt;
let googleProfile = storedGoogleSession.profile;
let googleTokenClient = null;
if (monthSelector) monthSelector.value = state.ui.month;

bootstrap();

function bootstrap() {
  bindFrameEvents();
  render();
}

function buildExcelPreset() {
  return {
    version: 3,
    ui: {
      month: "2026-07",
      activeGroup: "overview",
      activeModule: "dashboard",
    },
    connection: {
      provider: "google-drive",
      googleClientId: "",
      googleApiKey: "",
      googleFolderId: "",
      workspaceOwnerEmail: "",
      workspaceMainFolderId: "",
      workspaceMainFolderUrl: "",
      workspaceSheetId: "",
      workspaceSheetUrl: "",
      workspaceFolders: {},
      lastSyncedAt: "",
    },
    rooms: [
      { id: "301", name: "301 套房", status: "115/4/16起租", rent: 16000, deposit: 32000, rentDay: 15, equipment: "冷氣、熱水器、床架、衣櫃", note: "吳明峰（鴻海工程師）" },
      { id: "302", name: "302 套房", status: "115/4/25起租", rent: 15000, deposit: 30000, rentDay: 15, equipment: "冷氣、熱水器、床架、衣櫃", note: "黃書偉＋洪敬茹（餐飲業）" },
      { id: "303", name: "303 套房", status: "115/4/17起租", rent: 15500, deposit: 31000, rentDay: 15, equipment: "冷氣、熱水器、床架、衣櫃", note: "徐楷淵（南亞科工程師）" },
      { id: "304", name: "304 套房", status: "115/4/5起租", rent: 16000, deposit: 32000, rentDay: 5, equipment: "冷氣、熱水器、床架、衣櫃", note: "王弘軒（司法官實習生）" },
    ],
    tenants: [
      { id: "TEN-301", name: "吳明峰", phone: "", line: "", job: "鴻海工程師", emergency: "", roomId: "301", moveIn: "2026-04-16", moveOut: "", note: "Excel 預設資料匯入。" },
      { id: "TEN-302", name: "黃書偉＋洪敬茹", phone: "", line: "", job: "餐飲業", emergency: "", roomId: "302", moveIn: "2026-04-25", moveOut: "", note: "Excel 預設資料匯入。" },
      { id: "TEN-303", name: "徐楷淵", phone: "", line: "", job: "南亞科工程師", emergency: "", roomId: "303", moveIn: "2026-04-17", moveOut: "", note: "Excel 預設資料匯入。" },
      { id: "TEN-304", name: "王弘軒", phone: "", line: "", job: "司法官實習生", emergency: "", roomId: "304", moveIn: "2026-04-05", moveOut: "", note: "Excel 預設資料匯入。" },
    ],
    leases: [
      { id: "LE-301-202604", roomId: "301", tenantId: "TEN-301", startDate: "2026-04-16", endDate: "2027-04-15", rent: 16000, deposit: 32000, rentDay: 15, status: "生效中", note: "依 Excel 總表設定。" },
      { id: "LE-302-202604", roomId: "302", tenantId: "TEN-302", startDate: "2026-04-25", endDate: "2027-04-24", rent: 15000, deposit: 30000, rentDay: 15, status: "生效中", note: "依 Excel 總表設定。" },
      { id: "LE-303-202604", roomId: "303", tenantId: "TEN-303", startDate: "2026-04-17", endDate: "2027-04-16", rent: 15500, deposit: 31000, rentDay: 15, status: "生效中", note: "依 Excel 總表設定。" },
      { id: "LE-304-202604", roomId: "304", tenantId: "TEN-304", startDate: "2026-04-05", endDate: "2027-04-04", rent: 16000, deposit: 32000, rentDay: 5, status: "生效中", note: "依 Excel 總表設定。" },
    ],
    rentPayments: [
      { id: "RP-2026-07-304", month: "2026-07", roomId: "304", tenantId: "TEN-304", dueDate: "2026-07-05", amountDue: 16000, amountPaid: 16000, paymentDate: "2026-07-03", method: "現金", status: "已收", note: "Excel 每月記帳分表。" },
      { id: "RP-2026-07-302", month: "2026-07", roomId: "302", tenantId: "TEN-302", dueDate: "2026-07-15", amountDue: 15000, amountPaid: 15000, paymentDate: "2026-07-15", method: "轉帳", status: "已收", note: "Excel 每月記帳分表。" },
      { id: "RP-2026-07-303", month: "2026-07", roomId: "303", tenantId: "TEN-303", dueDate: "2026-07-18", amountDue: 15500, amountPaid: 15500, paymentDate: "2026-07-18", method: "轉帳", status: "已收", note: "Excel 每月記帳分表。" },
      { id: "RP-2026-07-301", month: "2026-07", roomId: "301", tenantId: "TEN-301", dueDate: "2026-07-18", amountDue: 16000, amountPaid: 16000, paymentDate: "2026-07-18", method: "轉帳", status: "已收", note: "Excel 每月記帳分表。" },
    ],
    bills: [
      { id: "BILL-2026-06", name: "2026 年 5-6 月臺電帳單", startDate: "2026-04-01", endDate: "2026-06-02", commonFee: 0, indoorFee: 3871, dueDate: "2026-07-18", paidToTaipower: true, payDate: "2026-07-01", note: "總表電費欄位。" },
    ],
    meterReadings: [
      { id: "MR-2026-06-301", billId: "BILL-2026-06", roomId: "301", previous: 1000, current: 1064, usage: 64, payer: "房客", note: "示意抄表。" },
      { id: "MR-2026-06-302", billId: "BILL-2026-06", roomId: "302", previous: 1000, current: 1060, usage: 60, payer: "房客", note: "示意抄表。" },
      { id: "MR-2026-06-303", billId: "BILL-2026-06", roomId: "303", previous: 1000, current: 1062, usage: 62, payer: "房客", note: "示意抄表。" },
      { id: "MR-2026-06-304", billId: "BILL-2026-06", roomId: "304", previous: 1000, current: 1065, usage: 65, payer: "房客", note: "示意抄表。" },
    ],
    expenses: [
      { id: "EX-001", date: "2026-07-10", month: "2026-07", type: "垃圾清潔費", scope: "公共", amount: 1200, payer: "房東", paid: "是", maintenanceId: "", note: "Excel 每月記帳分表。" },
      { id: "EX-002", date: "2026-04-08", month: "2026-04", type: "設備", scope: "公共", amount: 2184, payer: "房東", paid: "是", maintenanceId: "", note: "偵煙器 / momo購買。" },
      { id: "EX-003", date: "2026-04-12", month: "2026-04", type: "設備", scope: "公共", amount: 300, payer: "房東", paid: "是", maintenanceId: "", note: "滅火器 / 蝦皮購買。" },
      { id: "EX-004", date: "2026-05-31", month: "2026-05", type: "水費", scope: "公共", amount: 232, payer: "房東", paid: "是", maintenanceId: "", note: "2月27日～4月30日水費。" },
    ],
    maintenance: [],
    deposits: [
      { id: "DP-301", roomId: "301", tenantId: "TEN-301", leaseId: "LE-301-202604", expected: 32000, received: 32000, receiveDate: "2026-04-16", method: "轉帳", status: "保管中", offsetType: "", offsetAmount: 0, refundAmount: 0, refundDate: "", note: "" },
      { id: "DP-302", roomId: "302", tenantId: "TEN-302", leaseId: "LE-302-202604", expected: 30000, received: 30000, receiveDate: "2026-04-25", method: "轉帳", status: "保管中", offsetType: "", offsetAmount: 0, refundAmount: 0, refundDate: "", note: "" },
      { id: "DP-303", roomId: "303", tenantId: "TEN-303", leaseId: "LE-303-202604", expected: 31000, received: 31000, receiveDate: "2026-04-17", method: "轉帳", status: "保管中", offsetType: "", offsetAmount: 0, refundAmount: 0, refundDate: "", note: "" },
      { id: "DP-304", roomId: "304", tenantId: "TEN-304", leaseId: "LE-304-202604", expected: 32000, received: 32000, receiveDate: "2026-04-05", method: "現金", status: "保管中", offsetType: "", offsetAmount: 0, refundAmount: 0, refundDate: "", note: "" },
    ],
    attachments: [],
    ledgerEntries: [
      { id: "LG-001", date: "2026-04-05", month: "2026-04", roomId: "304", item: "4月房租", category: "房租", income: 16000, expense: 0, note: "現金" },
      { id: "LG-002", date: "2026-04-06", month: "2026-04", roomId: "304", item: "公證費", category: "租賃行政", income: 0, expense: 1500, note: "" },
      { id: "LG-003", date: "2026-04-08", month: "2026-04", roomId: "304", item: "戴先生仲介費（半個月）", category: "仲介費", income: 0, expense: 8000, note: "" },
      { id: "LG-004", date: "2026-04-08", month: "2026-04", roomId: "公共", item: "偵煙器", category: "公共設備", income: 0, expense: 2184, note: "momo購買" },
      { id: "LG-005", date: "2026-04-12", month: "2026-04", roomId: "公共", item: "滅火器", category: "公共設備", income: 0, expense: 300, note: "蝦皮購買" },
      { id: "LG-006", date: "2026-04-15", month: "2026-04", roomId: "301", item: "4月房租", category: "房租", income: 16000, expense: 0, note: "" },
      { id: "LG-007", date: "2026-04-15", month: "2026-04", roomId: "302", item: "4月房租", category: "房租", income: 15000, expense: 0, note: "" },
      { id: "LG-008", date: "2026-04-15", month: "2026-04", roomId: "303", item: "4月房租", category: "房租", income: 15500, expense: 0, note: "" },
      { id: "LG-009", date: "2026-05-01", month: "2026-05", roomId: "公共", item: "垃圾清潔費", category: "垃圾清潔費", income: 0, expense: 1200, note: "" },
      { id: "LG-010", date: "2026-05-31", month: "2026-05", roomId: "公共", item: "2月27日～4月30日水費", category: "水費", income: 0, expense: 232, note: "" },
      { id: "LG-011", date: "2026-07-03", month: "2026-07", roomId: "304", item: "7月房租", category: "房租", income: 16000, expense: 0, note: "" },
      { id: "LG-012", date: "2026-07-10", month: "2026-07", roomId: "公共", item: "垃圾清潔費", category: "垃圾清潔費", income: 0, expense: 1200, note: "" },
      { id: "LG-013", date: "2026-07-15", month: "2026-07", roomId: "302", item: "7月房租", category: "房租", income: 15000, expense: 0, note: "" },
      { id: "LG-014", date: "2026-07-18", month: "2026-07", roomId: "303", item: "7月房租", category: "房租", income: 15500, expense: 0, note: "" },
      { id: "LG-015", date: "2026-07-18", month: "2026-07", roomId: "301", item: "7月房租", category: "房租", income: 16000, expense: 0, note: "" },
    ],
    settings: {
      roomStatuses: ["出租中", "空房", "整修中", "即將退租"],
      expenseTypes: ["水費", "垃圾清潔費", "維修", "設備", "清潔", "耗材", "稅費", "仲介費", "代管費", "租賃行政", "其他"],
      repairStatuses: ["待處理", "已派工", "處理中", "已完成", "取消"],
      electricMode: "臺電帳單分攤模式",
      rounding: "四捨五入至整數元，尾差調整至用電最高房",
    },
  };
}

function buildBlankState() {
  const base = buildExcelPreset();
  return {
    ...base,
    ui: {
      month: defaultMonth,
      activeGroup: "overview",
      activeModule: "dashboard",
    },
    rooms: roomIds.map((roomId) => ({ id: roomId, name: `${roomId} 套房`, status: "空房", rent: 0, deposit: 0, rentDay: 5, equipment: "", note: "" })),
    tenants: [],
    leases: [],
    rentPayments: [],
    bills: [],
    meterReadings: [],
    expenses: [],
    maintenance: [],
    deposits: [],
    attachments: [],
    ledgerEntries: [],
  };
}

function normalizeState(raw) {
  const seed = buildExcelPreset();
  if (!raw) return sanitizeState(seed);
  return sanitizeState({
    ...seed,
    ...raw,
    ui: { ...seed.ui, ...(raw.ui || {}) },
    connection: { ...seed.connection, ...(raw.connection || {}) },
    settings: { ...seed.settings, ...(raw.settings || {}) },
    rooms: Array.isArray(raw.rooms) ? raw.rooms : seed.rooms,
    tenants: Array.isArray(raw.tenants) ? raw.tenants : seed.tenants,
    leases: Array.isArray(raw.leases) ? raw.leases : seed.leases,
    rentPayments: Array.isArray(raw.rentPayments) ? raw.rentPayments : seed.rentPayments,
    bills: Array.isArray(raw.bills) ? raw.bills : seed.bills,
    meterReadings: Array.isArray(raw.meterReadings) ? raw.meterReadings : seed.meterReadings,
    expenses: Array.isArray(raw.expenses) ? raw.expenses : seed.expenses,
    maintenance: Array.isArray(raw.maintenance) ? raw.maintenance : seed.maintenance,
    deposits: Array.isArray(raw.deposits) ? raw.deposits : seed.deposits,
    attachments: Array.isArray(raw.attachments) ? raw.attachments : seed.attachments,
    ledgerEntries: Array.isArray(raw.ledgerEntries) ? raw.ledgerEntries : seed.ledgerEntries,
  });
}

function loadLocalState() {
  const raw = localStorage.getItem(LOCAL_KEY);
  if (!raw) return buildExcelPreset();
  try {
    return sanitizeState(JSON.parse(raw));
  } catch {
    return buildExcelPreset();
  }
}

function saveLocalState() {
  localStorage.setItem(LOCAL_KEY, JSON.stringify(state));
}

function bindFrameEvents() {
  const resetDemo = document.querySelector("#reset-demo");
  if (resetDemo) {
    resetDemo.addEventListener("click", () => {
      state = buildExcelPreset();
      persistState();
      render();
    });
  }

  if (monthSelector) {
    monthSelector.addEventListener("change", (event) => {
      state.ui.month = event.target.value;
      persistState();
      render();
    });
  }

  recordModal.addEventListener("click", (event) => {
    if (event.target === recordModal) recordModal.close();
  });
}

function persistState() {
  state.connection.lastSyncedAt = new Date().toISOString();
  saveLocalState();
}

function render() {
  ensureDerivedRecords();
  applyTheme();
  renderLoginPanel();
  renderNav();
  renderModule();
}

function applyTheme() {
  const themeMap = {
    overview: ["#ff9b6a", "#ff5ca8", "#6e7bff"],
    movein: ["#ffb24f", "#ff7f50", "#ff4f9a"],
    finance: ["#00c389", "#06a8ff", "#6a67ff"],
    operations: ["#ff6b6b", "#ff9f43", "#ffd166"],
    analysis: ["#6f7cff", "#8d5bff", "#23c8c8"],
  };
  const [a, b, c] = themeMap[state.ui.activeGroup] || themeMap.overview;
  document.documentElement.style.setProperty("--theme-a", a);
  document.documentElement.style.setProperty("--theme-b", b);
  document.documentElement.style.setProperty("--theme-c", c);
}

function renderLoginPanel() {
  const workspace = getWorkspace();
  const signedIn = Boolean(googleProfile?.email);
  const activeToken = hasActiveGoogleToken();
  const usingPersonalWorkspace = state.connection.workspaceOwnerEmail === googleProfile?.email;
  const accessMismatch = signedIn && workspace.ownerEmail && googleProfile.email !== workspace.ownerEmail;
  const statusClass = signedIn ? "success" : GOOGLE_CLIENT_ID ? "" : "danger";
  const statusText = signedIn
    ? activeToken
      ? `已登入：${googleProfile.email}`
      : `已記住帳號：${googleProfile.email}，請重新授權`
    : GOOGLE_CLIENT_ID
      ? "尚未登入 Google"
      : "Google 登入尚未啟用";
  loginPanel.innerHTML = `
    <p class="overline">Google 雲端工作區</p>
    <h3>${signedIn ? escapeHtml(googleProfile.name || "Google 使用者") : "登入 Google 帳號讀取後臺資料"}</h3>
    <div class="sync-pill ${statusClass}">${escapeHtml(statusText)}</div>
    <div class="meta-line">流程：開啟系統 → Google 登入 → 讀取後臺主表 → 上傳附件到 Drive 子資料夾 → 自動把檔案連結回寫主表。</div>
    <div class="meta-line">目前作用中的儲存位置屬於 ${escapeHtml(workspace.ownerEmail || "尚未設定")}。${accessMismatch ? "你現在登入的帳號和資料夾擁有者不同，所以點進資料夾會要求權限。" : "若使用同一個擁有者帳號登入，資料會直接存到這個共用資料夾。"} </div>
    <div class="toolbar">
      <button class="cta-btn" id="google-login-button">${signedIn ? (activeToken ? "重新授權 Google" : "重新連線 Google") : "使用 Google 登入"}</button>
      ${signedIn ? `<button class="soft-btn" id="create-drive-workspace">${usingPersonalWorkspace ? "重建我的雲端工作區" : "建立我的雲端工作區"}</button>` : ""}
      <a class="soft-btn" href="${workspace.mainFolderUrl}" target="_blank" rel="noreferrer">雲端資料夾</a>
      <a class="soft-btn" href="${workspace.sheetUrl}" target="_blank" rel="noreferrer">後臺主表</a>
    </div>
  `;
}

function renderNav() {
  primaryNav.innerHTML = NAV_GROUPS.map((group) => {
    const active = group.id === state.ui.activeGroup ? "active" : "";
    return `<button class="nav-btn ${active}" data-group="${group.id}">${group.label}</button>`;
  }).join("");

  const activeGroup = NAV_GROUPS.find((item) => item.id === state.ui.activeGroup) || NAV_GROUPS[0];
  secondaryNav.innerHTML = activeGroup.modules.map((moduleId) => {
    const module = MODULES[moduleId];
    const active = module.id === state.ui.activeModule ? "active" : "";
    return `<button class="subnav-btn ${active}" data-module="${module.id}">${module.title}</button>`;
  }).join("");

  primaryNav.querySelectorAll("[data-group]").forEach((button) => {
    button.addEventListener("click", () => {
      const group = NAV_GROUPS.find((item) => item.id === button.dataset.group);
      state.ui.activeGroup = group.id;
      state.ui.activeModule = group.modules[0];
      persistState();
      render();
    });
  });

  secondaryNav.querySelectorAll("[data-module]").forEach((button) => {
    button.addEventListener("click", () => {
      state.ui.activeModule = button.dataset.module;
      persistState();
      render();
    });
  });
}

function renderModule() {
  const moduleMap = {
    dashboard: renderDashboard,
    rooms: renderRooms,
    tenants: renderTenants,
    leases: renderLeases,
    deposits: renderDeposits,
    rent: renderRent,
    bills: renderBills,
    meters: renderMeters,
    allocations: renderAllocations,
    expenses: renderExpenses,
    ledger: renderLedger,
    maintenance: renderMaintenance,
    attachments: renderAttachments,
    reports: renderReports,
    settings: renderSettings,
  };

  appContent.innerHTML = moduleMap[state.ui.activeModule]();
  bindDynamicActions();
}

function ensureDerivedRecords() {
  state.leases.forEach((lease) => {
    if (!leaseActiveInMonth(lease, state.ui.month)) return;
    const paymentId = `RP-${state.ui.month}-${lease.roomId}`;
    if (state.rentPayments.find((item) => item.id === paymentId)) return;
    state.rentPayments.push({
      id: paymentId,
      month: state.ui.month,
      roomId: lease.roomId,
      tenantId: lease.tenantId,
      dueDate: `${state.ui.month}-${String(lease.rentDay).padStart(2, "0")}`,
      amountDue: Number(lease.rent),
      amountPaid: 0,
      paymentDate: "",
      method: "",
      status: "未收",
      note: "系統依租約自動產生。",
    });
  });
  saveLocalState();
}

function leaseActiveInMonth(lease, month) {
  const current = new Date(`${month}-01`);
  const start = new Date(lease.startDate);
  const end = new Date(lease.endDate);
  return current >= new Date(start.getFullYear(), start.getMonth(), 1) && current <= new Date(end.getFullYear(), end.getMonth() + 1, 0);
}

function renderDashboard() {
  const summary = buildSummary();
  const progressCards = [
    { title: "入住建檔", value: `${state.tenants.length} 位租客`, note: `${state.leases.filter((item) => item.status === "生效中").length} 份有效租約` },
    { title: "本月收租", value: `${money.format(summary.rentPaid)} / ${money.format(summary.rentDue)}`, note: `${summary.unpaidRentRooms.length} 間待追款` },
    { title: "本月收支", value: money.format(summary.netProfit), note: `收入 ${money.format(summary.rentPaid)}，支出 ${money.format(summary.expenses)}` },
  ];

  return `
    <section class="hero-banner">
      <div class="stack">
        <div>
          <p class="overline">今日重點</p>
          <h2>把本月真正需要處理的房租、電費、退租與附件，集中在一個畫面先處理。</h2>
        </div>
        <p class="top-note">首頁只顯示目前真的存在的資料，不再預設塞入示意維修或示意附件。文字數字會優先保存在後臺主表，圖片與 PDF 另存到 Drive 子資料夾。</p>
        <div class="process-inline">
          <span class="flow-inline">1. 入住建檔</span>
          <span class="flow-inline">2. 每月收租與抄表</span>
          <span class="flow-inline">3. 上傳附件與結算</span>
        </div>
        <div class="hero-actions">
          ${actionButton("rent", "補登房租")}
          ${actionButton("leases", "新增租約")}
          ${actionButton("maintenance", "新增維修")}
        </div>
      </div>
      <div class="hero-side">
        ${progressCards.map((item) => `
          <article class="summary-card">
            <span class="overline">${item.title}</span>
            <strong>${item.value}</strong>
            <span class="muted">${item.note}</span>
          </article>
        `).join("")}
      </div>
    </section>
    <section class="metric-grid">
      ${metricCard("本月房租已收", money.format(summary.rentPaid), `尚未收 ${money.format(summary.rentDue - summary.rentPaid)}`)}
      ${metricCard("本月電費應收", money.format(summary.electricDue), `已回收 ${money.format(summary.electricPaid)}`)}
      ${metricCard("本月支出", money.format(summary.expenses), "房東負擔支出")}
      ${metricCard("本月淨利", money.format(summary.netProfit), "房租收入 - 房東支出")}
    </section>
    <section class="three-grid">
      <article class="flow-card">
        <p class="overline">待辦清單</p>
        <div class="stack">
          ${summary.reminders.length ? summary.reminders.map((text) => `<div class="alert-card">${text}</div>`).join("") : `<div class="alert-card">本月沒有待處理提醒。</div>`}
        </div>
      </article>
      <article class="flow-card">
        <p class="overline">即將到期</p>
        <div class="stack">
          ${summary.expiringLeases.length ? summary.expiringLeases.map((lease) => `<div class="timeline-card"><strong>${lease.roomId} ${tenantName(lease.tenantId)}</strong><div class="detail-list"><span>租約到期：${lease.endDate}</span><span>押金：${money.format(lease.deposit)}</span></div></div>`).join("") : `<div class="alert-card">目前沒有即將到期租約。</div>`}
        </div>
      </article>
      <article class="flow-card">
        <p class="overline">營運健康度</p>
        <div class="stack">${state.rooms.map((room) => roomHealthCard(room)).join("")}</div>
      </article>
    </section>
  `;
}

function renderRooms() {
  return `
    ${moduleBanner("房間狀態", "每間房目前住誰、租金多少、狀態如何，一次看清楚。")}
    <section class="four-grid">
      ${state.rooms.map((room) => `
        <article class="room-health">
          <div class="split">
            <div><p class="overline">${room.id}</p><h3>${room.name}</h3></div>
            ${statusTag(room.status)}
          </div>
          <div class="detail-list">
            <span>目前租客：${activeTenantByRoom(room.id)?.name || "無"}</span>
            <span>月租：${money.format(room.rent)}</span>
            <span>押金：${money.format(room.deposit)}</span>
            <span>設備：${room.equipment || "未填"}</span>
            <span>備註：${room.note || "無"}</span>
          </div>
        </article>
      `).join("")}
    </section>
  `;
}

function renderTenants() {
  return renderTableModule({
    title: "租客名單",
    subtitle: "租客資料獨立保存，未來再次入住時不需要重建。",
    addKey: "tenants",
    addLabel: "新增租客",
    headers: ["姓名", "聯絡方式", "房號", "入住 / 退租", "職業", "備註"],
    rows: state.tenants.map((tenant) => `
      <tr>
        <td><strong>${tenant.name}</strong></td>
        <td>${tenant.phone || "未填"}<br /><small>${tenant.line || "無其他聯絡方式"}</small></td>
        <td>${tenant.roomId || "未入住"}</td>
        <td>${tenant.moveIn || "未填"}<br /><small>${tenant.moveOut || "現住中"}</small></td>
        <td>${tenant.job || "未填"}</td>
        <td>${tenant.note || "無"}</td>
      </tr>
    `).join(""),
  });
}

function renderLeases() {
  const leaseAttachments = state.attachments.filter((item) => item.module === "租約");
  return `
    ${renderTableModule({
      title: "租約管理",
      subtitle: "租約是房租、押金、到期提醒與退租結算的基礎，也支援上傳合約檔案中繼資料。",
      addKey: "leases",
      addLabel: "新增租約",
      headers: ["租約編號", "房號", "租客", "租期", "月租 / 押金", "狀態", "備註"],
      rows: state.leases.map((lease) => `
        <tr>
          <td>${lease.id}</td>
          <td>${lease.roomId}</td>
          <td>${tenantName(lease.tenantId)}</td>
          <td>${lease.startDate}<br /><small>${lease.endDate}</small></td>
          <td>${money.format(lease.rent)}<br /><small>押金 ${money.format(lease.deposit)}</small></td>
          <td>${statusTag(lease.status)}</td>
          <td>${lease.note || "無"}</td>
        </tr>
      `).join(""),
    })}
    <section class="panel">
      <div class="panel-header">
        <div>
          <p class="overline">租約附件</p>
          <h3>照片檔、PDF、掃描檔可從這裡補登</h3>
        </div>
      </div>
      <div class="upload-box">
        <div class="form-grid">
          ${selectField("lease-upload-target", "對應租約", state.leases.map((item) => ({ value: item.id, label: `${item.roomId} / ${tenantName(item.tenantId)}` })))}
          ${inputField("lease-upload-note", "檔案說明", "text")}
          <div class="field full-span">
            <label for="lease-upload-file">選擇租約檔案</label>
            <input id="lease-upload-file" type="file" accept=".pdf,image/*" />
          </div>
        </div>
        <div class="toolbar wrap-top">
          <button class="cta-btn" id="save-lease-upload">新增租約附件</button>
        </div>
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>附件名稱</th><th>租約</th><th>房號</th><th>租客</th><th>日期</th><th>備註</th></tr></thead>
          <tbody>
            ${leaseAttachments.map((file) => `
              <tr>
                <td>${file.name}</td>
                <td>${file.recordId}</td>
                <td>${file.roomId || "無"}</td>
                <td>${file.tenant || "無"}</td>
                <td>${file.date}</td>
                <td>${file.note || "無"}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderDeposits() {
  return renderTableModule({
    title: "押金管理",
    subtitle: "押金不列入租金收入，保留扣抵與退還過程。",
    addKey: "deposits",
    addLabel: "新增押金紀錄",
    headers: ["房號", "租客", "應收 / 實收", "狀態", "扣抵", "退還", "備註"],
    rows: state.deposits.map((item) => `
      <tr>
        <td>${item.roomId}</td>
        <td>${tenantName(item.tenantId)}</td>
        <td>${money.format(item.expected)}<br /><small>實收 ${money.format(item.received)}</small></td>
        <td>${statusTag(item.status)}</td>
        <td>${item.offsetAmount ? `${item.offsetType} ${money.format(item.offsetAmount)}` : "無"}</td>
        <td>${item.refundAmount ? `${money.format(item.refundAmount)} / ${item.refundDate || "未填日期"}` : "未退還"}</td>
        <td>${item.note || "無"}</td>
      </tr>
    `).join(""),
  });
}

function renderRent() {
  const payments = monthPayments();
  return `
    ${moduleBanner("房租收款", `${state.ui.month} 的應收、已收與未收集中管理。`)}
    <section class="three-grid">
      ${metricCard("應收房租", money.format(sumBy(payments, "amountDue")), `${payments.length} 筆應收`)}
      ${metricCard("已收房租", money.format(sumBy(payments, "amountPaid")), `${payments.filter((item) => item.status === "已收").length} 間已收齊`)}
      ${metricCard("待追款", money.format(sumBy(payments, "amountDue") - sumBy(payments, "amountPaid")), `${payments.filter((item) => item.amountPaid < item.amountDue).length} 間未結清`)}
    </section>
    ${renderTableModule({
      title: "收租明細",
      subtitle: "收款日期與歸屬月份分開記錄，可追蹤延後繳與部分收款。",
      addKey: "rent",
      addLabel: "新增收款",
      headers: ["房號", "租客", "應收日期", "應收", "已收", "狀態", "收款日期", "備註"],
      rows: payments.map((payment) => `
        <tr>
          <td>${payment.roomId}</td>
          <td>${tenantName(payment.tenantId)}</td>
          <td>${payment.dueDate}</td>
          <td>${money.format(payment.amountDue)}</td>
          <td>${money.format(payment.amountPaid)}</td>
          <td>${statusTag(payment.status)}</td>
          <td>${payment.paymentDate || "未收"}</td>
          <td>${payment.note || "無"}</td>
        </tr>
      `).join(""),
    })}
  `;
}

function renderBills() {
  return renderTableModule({
    title: "臺電帳單",
    subtitle: "先輸入帳單，再接續抄表與電費分攤。",
    addKey: "bills",
    addLabel: "新增臺電帳單",
    headers: ["帳單名稱", "期間", "公用電費", "戶內電費", "總額", "繳費狀態", "備註"],
    rows: state.bills.map((bill) => `
      <tr>
        <td>${bill.name}</td>
        <td>${bill.startDate}<br /><small>${bill.endDate}</small></td>
        <td>${money.format(bill.commonFee)}</td>
        <td>${money.format(bill.indoorFee)}</td>
        <td>${money.format(Number(bill.commonFee) + Number(bill.indoorFee))}</td>
        <td>${bill.paidToTaipower ? "已繳" : "未繳"}<br /><small>${bill.payDate || "待填"}</small></td>
        <td>${bill.note || "無"}</td>
      </tr>
    `).join(""),
  });
}

function renderMeters() {
  return renderTableModule({
    title: "電表抄表",
    subtitle: "若遇換表、空房或中途入住，可用備註保留原因。",
    addKey: "meters",
    addLabel: "新增抄表紀錄",
    headers: ["帳單編號", "房號", "上期", "本期", "使用度數", "負擔者", "備註"],
    rows: state.meterReadings.map((reading) => `
      <tr>
        <td>${reading.billId}</td>
        <td>${reading.roomId}</td>
        <td>${reading.previous}</td>
        <td>${reading.current}</td>
        <td>${reading.usage}</td>
        <td>${reading.payer}</td>
        <td>${reading.note || "無"}</td>
      </tr>
    `).join(""),
  });
}

function renderAllocations() {
  const bill = state.bills.at(-1);
  const allocation = calculateAllocation(bill?.id);
  return `
    ${moduleBanner("電費分攤", "依臺電總額與各房用電比例分攤，電費不併入租金收入。")}
    <section class="three-grid">
      ${metricCard("本期臺電總額", bill ? money.format(Number(bill.commonFee) + Number(bill.indoorFee)) : money.format(0), bill ? bill.name : "尚無帳單")}
      ${metricCard("本期電費應收", money.format(sumBy(allocation, "finalDue")), `${allocation.length} 間分攤`)}
      ${metricCard("本期已回收", money.format(sumBy(allocation, "paid")), `${allocation.filter((item) => item.paid >= item.finalDue).length} 間已結清`)}
    </section>
    ${renderTableModule({
      title: "分攤明細",
      subtitle: "尾差自動調整至用電最高房，便於對房客說明。",
      addKey: "",
      addLabel: "",
      headers: ["房號", "使用度數", "比例", "戶內分攤", "公用分攤", "尾差", "最終應收", "已收", "狀態"],
      rows: allocation.map((item) => `
        <tr>
          <td>${item.roomId}</td>
          <td>${item.usage} 度</td>
          <td>${percent.format(item.ratio)}</td>
          <td>${money.format(item.indoor)}</td>
          <td>${money.format(item.common)}</td>
          <td>${money.format(item.tailAdjustment)}</td>
          <td>${money.format(item.finalDue)}</td>
          <td>${money.format(item.paid)}</td>
          <td>${statusTag(item.paid >= item.finalDue ? "已收" : item.paid > 0 ? "部分收" : "未收")}</td>
        </tr>
      `).join(""),
    })}
  `;
}

function renderExpenses() {
  const monthly = state.expenses.filter((item) => item.month === state.ui.month);
  return `
    ${moduleBanner("支出管理", "公共支出與房間支出分開記錄，方便後續計算整體與房間別損益。")}
    <section class="three-grid">
      ${expenseBreakdownCard("公共支出", monthly.filter((item) => item.scope === "公共"))}
      ${expenseBreakdownCard("房間支出", monthly.filter((item) => item.scope !== "公共"))}
      ${expenseBreakdownCard("未付款支出", monthly.filter((item) => item.paid === "否"))}
    </section>
    ${renderTableModule({
      title: "支出明細",
      subtitle: "水費與垃圾清潔費自動視為房東支出。",
      addKey: "expenses",
      addLabel: "新增支出",
      headers: ["日期", "類型", "歸屬", "金額", "負擔者", "已付款", "備註"],
      rows: state.expenses.map((expense) => `
        <tr>
          <td>${expense.date}</td>
          <td>${expense.type}</td>
          <td>${expense.scope}</td>
          <td>${money.format(expense.amount)}</td>
          <td>${expense.payer}</td>
          <td>${expense.paid}</td>
          <td>${expense.note || "無"}</td>
        </tr>
      `).join(""),
    })}
  `;
}

function renderLedger() {
  const entries = state.ledgerEntries.filter((item) => item.month === state.ui.month);
  const income = sumBy(entries, "income");
  const expense = sumBy(entries, "expense");
  const categories = aggregateLedgerCategories(entries);
  return `
    ${moduleBanner("總收支明細", "把收入與支出的大項分類、筆數與盈虧集中看，快速掌握經營狀況。")}
    <section class="three-grid">
      ${metricCard("本月總收入", money.format(income), `${entries.filter((item) => item.income > 0).length} 筆收入`)}
      ${metricCard("本月總支出", money.format(expense), `${entries.filter((item) => item.expense > 0).length} 筆支出`)}
      ${metricCard("本月收支差額", money.format(income - expense), "依 Excel 預設與系統逐筆資料彙整")}
    </section>
    <section class="two-grid">
      <article class="panel">
        <div class="panel-header"><div><p class="overline">分類彙總</p><h3>收入與支出大項</h3></div></div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>分類</th><th>收入</th><th>支出</th><th>筆數</th></tr></thead>
            <tbody>
              ${categories.map((item) => `
                <tr>
                  <td>${item.category}</td>
                  <td class="ledger-income">${money.format(item.income)}</td>
                  <td class="ledger-expense">${money.format(item.expense)}</td>
                  <td>${item.count}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      </article>
      <article class="panel">
        <div class="panel-header"><div><p class="overline">本月明細</p><h3>${state.ui.month}</h3></div></div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>日期</th><th>房號</th><th>項目</th><th>收入</th><th>支出</th><th>備註</th></tr></thead>
            <tbody>
              ${entries.map((item) => `
                <tr>
                  <td>${item.date}</td>
                  <td>${item.roomId}</td>
                  <td>${item.item}</td>
                  <td class="ledger-income">${item.income ? money.format(item.income) : "-"}</td>
                  <td class="ledger-expense">${item.expense ? money.format(item.expense) : "-"}</td>
                  <td>${item.note || "無"}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      </article>
    </section>
  `;
}

function renderMaintenance() {
  return `
    ${moduleBanner("維修處理", "維修不只記一筆費用，還要保留責任歸屬、處理狀態與退租爭議證據。")}
    <section class="three-grid">
      ${statusSummaryCard("待處理", state.maintenance.filter((item) => item.status === "待處理"))}
      ${statusSummaryCard("處理中", state.maintenance.filter((item) => ["已派工", "處理中"].includes(item.status)))}
      ${statusSummaryCard("已完成", state.maintenance.filter((item) => item.status === "已完成"))}
    </section>
    ${renderTableModule({
      title: "維修案件",
      subtitle: "完成後可再轉入支出、應收或押金扣抵，並搭配收據與照片附件。",
      addKey: "maintenance",
      addLabel: "新增維修案件",
      headers: ["案件編號", "房號", "類型", "回報人", "狀態", "金額", "負擔者", "備註"],
      rows: state.maintenance.map((item) => `
        <tr>
          <td>${item.id}</td>
          <td>${item.roomId}</td>
          <td>${item.type}</td>
          <td>${item.reporter}</td>
          <td>${statusTag(item.status)}</td>
          <td>${money.format(item.cost)}</td>
          <td>${item.owner}</td>
          <td>${item.note || "無"}</td>
        </tr>
      `).join(""),
    })}
  `;
}

function renderAttachments() {
  return `
    ${moduleBanner("附件中心", "租約、帳單、維修照片與收據都集中在這裡管理。")}
    <section class="three-grid">
      ${state.attachments.map((file) => `
        <article class="attachment-card">
          <div class="split">
            <div><p class="overline">${file.type}</p><h3>${file.name}</h3></div>
            <span class="chip">${file.module}</span>
          </div>
          <div class="detail-list">
            <span>關聯資料：${file.recordId}</span>
            <span>房號：${file.roomId || "無"}</span>
            <span>租客：${file.tenant || "無"}</span>
            <span>日期：${file.date}</span>
            <span>路徑：${file.url || "尚未上傳到雲端"}</span>
          </div>
        </article>
      `).join("")}
    </section>
    <section class="panel">
      <div class="panel-header">
        <div><p class="overline">新增附件</p><h3>建立附件紀錄</h3></div>
        <button class="cta-btn" data-add="attachments">新增附件</button>
      </div>
    </section>
  `;
}

function renderReports() {
  const summary = buildSummary();
  const roomRows = roomIds.map((roomId) => {
    const paidRent = monthPayments().find((item) => item.roomId === roomId)?.amountPaid || 0;
    const roomExpenses = state.expenses.filter((item) => item.month === state.ui.month && item.scope === roomId).reduce((sum, item) => sum + Number(item.amount), 0);
    return { roomId, paidRent, roomExpenses, publicShare: summary.publicExpenseShare, net: paidRent - roomExpenses - summary.publicExpenseShare };
  });

  return `
    ${moduleBanner("經營報表", "報表只由逐筆紀錄彙整，避免人工重做總表。")}
    <section class="metric-grid">
      ${metricCard("月報房租實收", money.format(summary.rentPaid), state.ui.month)}
      ${metricCard("年度化房租", money.format(summary.rentPaid * 12), "依目前月資料估算")}
      ${metricCard("年度化支出", money.format(summary.expenses * 12), "依目前月資料估算")}
      ${metricCard("年度化淨利", money.format(summary.netProfit * 12), "依目前月資料估算")}
    </section>
    ${renderTableModule({
      title: "房間別損益",
      subtitle: "公共支出平均分攤後，先看目前月的房間別經營表現。",
      addKey: "",
      addLabel: "",
      headers: ["房號", "房租收入", "房間支出", "公共支出分攤", "房間淨利"],
      rows: roomRows.map((item) => `
        <tr>
          <td>${item.roomId}</td>
          <td>${money.format(item.paidRent)}</td>
          <td>${money.format(item.roomExpenses)}</td>
          <td>${money.format(item.publicShare)}</td>
          <td>${money.format(item.net)}</td>
        </tr>
      `).join(""),
    })}
  `;
}

function renderSettings() {
  const workspace = getWorkspace();
  const folderNames = WORKSPACE_BLUEPRINT.map((item) => item.name).join("、");
  return `
    ${moduleBanner("系統與儲存", "這裡不再放技術設定，而是直接用你看得懂的資料入口與備份結構說明。")}
    <section class="two-grid">
      <article class="panel">
        <div class="panel-header">
          <div><p class="overline">Google 雲端入口</p><h3>主資料夾、主表與附件結構</h3></div>
        </div>
        <div class="sync-detail">
          <div class="alert-card">目前資料實際會存到這裡：主資料夾屬於 ${workspace.ownerEmail || "尚未設定"}，後臺主表放在 01_Excel表，照片與 PDF 會依序進入 ${folderNames} 這些子資料夾。</div>
          <div class="toolbar">
            <a class="cta-btn" href="${workspace.mainFolderUrl}" target="_blank" rel="noreferrer">開啟主資料夾</a>
            <a class="soft-btn" href="${workspace.sheetUrl}" target="_blank" rel="noreferrer">開啟後臺主表</a>
          </div>
        </div>
      </article>
      <article class="panel">
        <div class="panel-header">
          <div><p class="overline">預設資料</p><h3>以 Excel 工作簿作為系統母版</h3></div>
        </div>
        <p class="list-note">目前預設資料已依循 \`20260322-新埔八街-收租財報分頁整合.xlsx\` 的房號、租客、租金、押金與每月收支。若只是想清空回到空白專案，也可以一鍵重置。</p>
        <div class="toolbar wrap-top">
          <button class="cta-btn" id="apply-excel-preset">套用 Excel 預設</button>
          <button class="danger-btn" id="clear-all-data">清空所有資料</button>
        </div>
      </article>
    </section>
    <section class="three-grid">
      <article class="panel">
        <p class="overline">房間狀態</p>
        <div class="inline-pills">${state.settings.roomStatuses.map((item) => `<span class="chip">${item}</span>`).join("")}</div>
      </article>
      <article class="panel">
        <p class="overline">支出類型</p>
        <div class="inline-pills">${state.settings.expenseTypes.map((item) => `<span class="chip">${item}</span>`).join("")}</div>
      </article>
      <article class="panel">
        <p class="overline">最後更新</p>
        <div class="detail-list"><span>${formatDateTime(state.connection.lastSyncedAt)}</span><span>目前先保存於本機，待正式接上 Google Drive 後同步到共享資料夾。</span></div>
      </article>
    </section>
  `;
}

function handleGoogleLogin() {
  if (!GOOGLE_CLIENT_ID) {
    const clientId = window.prompt("請貼上 Google OAuth Client ID。這是公開識別碼，不是密碼。貼上後系統會存在這台瀏覽器，重新按一次登入即可授權。");
    if (clientId) {
      localStorage.setItem("rental-google-client-id", clientId.trim());
      window.alert("已儲存 Client ID，請重新按一次「使用 Google 登入」。");
      window.location.reload();
    }
    return;
  }

  if (!window.google?.accounts?.oauth2) {
    window.alert("Google 登入程式尚未載入完成，請稍後再試一次。");
    return;
  }

  if (!googleTokenClient) {
    googleTokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID,
      scope: GOOGLE_SCOPES,
      callback: async (tokenResponse) => {
        if (tokenResponse.error) {
          window.alert(`Google 授權失敗：${tokenResponse.error}`);
          return;
        }
        googleAccessToken = tokenResponse.access_token;
        googleAccessTokenExpiresAt = Date.now() + Number(tokenResponse.expires_in || 3600) * 1000;
        googleProfile = await fetchGoogleProfile();
        saveGoogleSession();
        render();
      },
    });
  }

  googleTokenClient.requestAccessToken({ prompt: hasActiveGoogleToken() ? "" : "consent" });
}

async function fetchGoogleProfile() {
  const response = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
    headers: { Authorization: `Bearer ${googleAccessToken}` },
  });
  if (!response.ok) return { email: "已授權 Google 帳號", name: "Google 使用者" };
  return response.json();
}

async function uploadAttachmentToDrive(file, meta) {
  if (!hasActiveGoogleToken()) {
    throw new Error("尚未登入 Google，附件先保留為本機待上傳紀錄。");
  }

  const folderId = resolveAttachmentFolder(meta);
  const metadata = {
    name: `${new Date().toISOString().slice(0, 10)}_${meta.roomId || "公共"}_${file.name}`,
    parents: [folderId],
  };
  const boundary = `rent_multipart_${Date.now()}`;
  const body = new Blob([
    `--${boundary}\r\n`,
    "Content-Type: application/json; charset=UTF-8\r\n\r\n",
    JSON.stringify(metadata),
    `\r\n--${boundary}\r\n`,
    `Content-Type: ${file.type || "application/octet-stream"}\r\n\r\n`,
    file,
    `\r\n--${boundary}--`,
  ], { type: `multipart/related; boundary=${boundary}` });

  const response = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink,parents", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${googleAccessToken}`,
      "Content-Type": `multipart/related; boundary=${boundary}`,
    },
    body,
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Drive 上傳失敗：${text}`);
  }

  return response.json();
}

function resolveAttachmentFolder(meta) {
  const workspace = getWorkspace();
  if (meta.module === "租約") return workspace.folders.leases;
  if (meta.module === "臺電帳單") return workspace.folders.taipower;
  if (meta.module === "電表抄表") return workspace.folders.meters;
  if (meta.module === "維修") return workspace.folders.maintenance;
  if (meta.module === "收款") return workspace.folders.receipts;
  if (meta.module === "租客") return workspace.folders.tenants;
  return workspace.folders.excel || workspace.mainFolderId;
}

async function appendAttachmentToSheet(attachment) {
  if (!hasActiveGoogleToken()) return false;
  await ensureAttachmentSheet();
  const range = encodeURIComponent("附件中心!A:J");
  const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${getWorkspace().sheetId}/values/${range}:append?valueInputOption=USER_ENTERED`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${googleAccessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      values: [[
        attachment.id,
        attachment.date,
        attachment.module,
        attachment.type,
        attachment.roomId,
        attachment.tenant,
        attachment.recordId,
        attachment.name,
        attachment.url,
        attachment.note,
      ]],
    }),
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`後臺主表回寫失敗：${text}`);
  }
  return true;
}

async function ensureAttachmentSheet() {
  const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${getWorkspace().sheetId}:batchUpdate`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${googleAccessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      requests: [
        { addSheet: { properties: { title: "附件中心" } } },
      ],
    }),
  });

  if (!response.ok && response.status !== 400) {
    const text = await response.text();
    throw new Error(`建立附件中心工作表失敗：${text}`);
  }

  const headerRange = encodeURIComponent("附件中心!A1:J1");
  await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${getWorkspace().sheetId}/values/${headerRange}?valueInputOption=USER_ENTERED`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${googleAccessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      values: [["附件編號", "日期", "模組", "類型", "房號", "租客", "關聯紀錄", "檔名", "Drive連結", "備註"]],
    }),
  });
}

async function createManagedWorkspace() {
  if (!hasActiveGoogleToken() || !googleProfile?.email) {
    window.alert("請先登入 Google 帳號，再建立你的雲端工作區。");
    return;
  }

  const createButton = document.querySelector("#create-drive-workspace");
  if (createButton) createButton.textContent = "建立中...";

  try {
    const mainFolder = await createDriveFolder("013_新埔八街出租管理系統");
    const folderMap = {};
    for (const item of WORKSPACE_BLUEPRINT) {
      const folder = await createDriveFolder(item.name, mainFolder.id);
      folderMap[item.key] = folder.id;
    }

    const spreadsheet = await createBackendSpreadsheet("新埔八街出租管理系統_後臺主表");
    await moveDriveFileToFolder(spreadsheet.spreadsheetId, folderMap.excel);
    await ensureBackendSheets(spreadsheet.spreadsheetId);

    state.connection.workspaceOwnerEmail = googleProfile.email;
    state.connection.workspaceMainFolderId = mainFolder.id;
    state.connection.workspaceMainFolderUrl = `https://drive.google.com/drive/folders/${mainFolder.id}`;
    state.connection.workspaceSheetId = spreadsheet.spreadsheetId;
    state.connection.workspaceSheetUrl = spreadsheet.spreadsheetUrl;
    state.connection.workspaceFolders = folderMap;
    persistState();
    render();
    window.alert("已在你目前登入的 Google 帳號底下建立新的雲端工作區。");
  } catch (error) {
    window.alert(`建立雲端工作區失敗：${error.message}`);
    render();
  }
}

async function createDriveFolder(name, parentId = "") {
  const metadata = {
    name,
    mimeType: "application/vnd.google-apps.folder",
  };
  if (parentId) metadata.parents = [parentId];
  const response = await fetch("https://www.googleapis.com/drive/v3/files?fields=id,name,webViewLink,parents", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${googleAccessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(metadata),
  });
  if (!response.ok) throw new Error(await response.text());
  return response.json();
}

async function createBackendSpreadsheet(title) {
  const response = await fetch("https://sheets.googleapis.com/v4/spreadsheets", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${googleAccessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      properties: { title },
      sheets: [{ properties: { title: "總覽設定" } }],
    }),
  });
  if (!response.ok) throw new Error(await response.text());
  return response.json();
}

async function moveDriveFileToFolder(fileId, parentId) {
  const metaResponse = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?fields=parents`, {
    headers: { Authorization: `Bearer ${googleAccessToken}` },
  });
  if (!metaResponse.ok) throw new Error(await metaResponse.text());
  const meta = await metaResponse.json();
  const removeParents = (meta.parents || []).join(",");
  const response = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?addParents=${encodeURIComponent(parentId)}&removeParents=${encodeURIComponent(removeParents)}&fields=id,parents`, {
    method: "PATCH",
    headers: { Authorization: `Bearer ${googleAccessToken}` },
  });
  if (!response.ok) throw new Error(await response.text());
}

async function ensureBackendSheets(spreadsheetId) {
  const requests = [
    { updateSpreadsheetProperties: { properties: { locale: "zh_TW", timeZone: "Asia/Taipei" }, fields: "locale,timeZone" } },
  ];
  const sheets = [
    ["房間資料", ["房號", "房名", "狀態", "月租金", "押金", "收租日", "設備", "備註", "更新時間"]],
    ["租客資料", ["租客編號", "姓名", "電話", "LINE", "職業", "緊急聯絡人", "房號", "入住日", "退租日", "備註", "更新時間"]],
    ["租約資料", ["租約編號", "房號", "租客編號", "起租日", "到期日", "月租金", "押金", "收租日", "狀態", "租約附件連結", "備註", "更新時間"]],
    ["收款紀錄", ["收款編號", "月份", "房號", "租客編號", "應收日", "應收金額", "實收金額", "收款日", "方式", "狀態", "憑證連結", "備註"]],
    ["臺電帳單", ["帳單編號", "帳單名稱", "起始日", "結束日", "公共電費", "室內電費", "繳費期限", "是否已繳臺電", "繳費日", "帳單附件連結", "備註"]],
    ["電表抄表", ["抄表編號", "帳單編號", "房號", "前期度數", "本期度數", "使用度數", "負擔者", "照片連結", "備註"]],
    ["支出紀錄", ["支出編號", "日期", "月份", "類型", "範圍", "金額", "付款者", "是否已付款", "關聯維修", "收據連結", "備註"]],
    ["維修紀錄", ["維修編號", "房號", "通報者", "通報日", "類型", "狀態", "費用", "負擔者", "廠商", "完成日", "附件連結", "備註"]],
    ["附件中心", ["附件編號", "日期", "模組", "類型", "房號", "租客", "關聯紀錄", "檔名", "Drive連結", "備註"]],
    ["總帳明細", ["明細編號", "日期", "月份", "房號", "項目", "分類", "收入", "支出", "附件連結", "備註"]],
  ];
  sheets.forEach((item) => requests.push({ addSheet: { properties: { title: item[0] } } }));

  const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${googleAccessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ requests }),
  });
  if (!response.ok) throw new Error(await response.text());

  for (const [sheetName, headers] of sheets) {
    const range = encodeURIComponent(`${sheetName}!A1:${String.fromCharCode(64 + headers.length)}1`);
    const headerResponse = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}?valueInputOption=USER_ENTERED`, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${googleAccessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ values: [headers] }),
    });
    if (!headerResponse.ok) throw new Error(await headerResponse.text());
  }
}

function bindDynamicActions() {
  const googleLoginButton = document.querySelector("#google-login-button");
  if (googleLoginButton) {
    googleLoginButton.addEventListener("click", handleGoogleLogin);
  }

  const createDriveWorkspaceButton = document.querySelector("#create-drive-workspace");
  if (createDriveWorkspaceButton) {
    createDriveWorkspaceButton.addEventListener("click", createManagedWorkspace);
  }

  document.querySelectorAll("[data-add]").forEach((button) => {
    if (!button.dataset.add) return;
    button.addEventListener("click", () => openModal(button.dataset.add));
  });

  document.querySelectorAll("[data-open-module]").forEach((button) => {
    button.addEventListener("click", () => {
      const moduleId = button.dataset.openModule;
      state.ui.activeGroup = MODULES[moduleId].group;
      state.ui.activeModule = moduleId;
      persistState();
      render();
    });
  });

  const applyPreset = document.querySelector("#apply-excel-preset");
  if (applyPreset) {
    applyPreset.addEventListener("click", () => {
      state = buildExcelPreset();
      persistState();
      render();
    });
  }

  const clearAll = document.querySelector("#clear-all-data");
  if (clearAll) {
    clearAll.addEventListener("click", () => {
      state = buildBlankState();
      persistState();
      render();
    });
  }

  const saveLeaseUpload = document.querySelector("#save-lease-upload");
  if (saveLeaseUpload) {
    saveLeaseUpload.addEventListener("click", async () => {
      const fileInput = document.querySelector("#lease-upload-file");
      const targetLeaseId = document.querySelector("#lease-upload-target")?.value;
      if (!fileInput?.files?.length || !targetLeaseId) {
        window.alert("請先選擇租約與檔案。");
        return;
      }
      const lease = state.leases.find((item) => item.id === targetLeaseId);
      const tenant = state.tenants.find((item) => item.id === lease?.tenantId);
      const file = fileInput.files[0];
      const attachment = {
        id: `AT-${Date.now()}`,
        name: file.name,
        type: file.type || "租約附件",
        module: "租約",
        recordId: targetLeaseId,
        roomId: lease?.roomId || "",
        tenant: tenant?.name || "",
        date: new Date().toISOString().slice(0, 10),
        url: `pending-upload://${file.name}`,
        note: document.querySelector("#lease-upload-note")?.value || "由租約管理頁面新增。",
      };

      try {
        const driveFile = await uploadAttachmentToDrive(file, attachment);
        attachment.url = driveFile.webViewLink || `https://drive.google.com/file/d/${driveFile.id}/view`;
        await appendAttachmentToSheet(attachment);
        attachment.note = `${attachment.note} 已上傳 Drive 並回寫後臺主表。`;
      } catch (error) {
        attachment.note = `${attachment.note} ${error.message}`;
      }

      state.attachments.push(attachment);
      persistState();
      render();
    });
  }
}

function openModal(moduleKey) {
  const configMap = {
    tenants: {
      title: "新增租客",
      fields: [
        inputField("id", "租客編號", "text"),
        inputField("name", "租客姓名", "text"),
        inputField("phone", "手機", "text"),
        inputField("line", "LINE", "text"),
        inputField("job", "職業", "text"),
        inputField("emergency", "緊急聯絡人", "text"),
        selectField("roomId", "目前房號", roomIds),
        inputField("moveIn", "入住日期", "date"),
        inputField("moveOut", "退租日期", "date"),
        inputField("note", "備註", "text", true),
      ],
      save(values) { state.tenants.push(values); },
    },
    leases: {
      title: "新增租約",
      fields: [
        inputField("id", "租約編號", "text"),
        selectField("roomId", "房號", roomIds),
        selectField("tenantId", "租客", state.tenants.map((item) => ({ value: item.id, label: item.name }))),
        inputField("startDate", "起租日", "date"),
        inputField("endDate", "到期日", "date"),
        inputField("rent", "月租金", "number"),
        inputField("deposit", "押金", "number"),
        inputField("rentDay", "收租日", "number"),
        selectField("status", "租約狀態", ["生效中", "即將到期", "已到期", "已退租"]),
        inputField("note", "備註", "text", true),
      ],
      save(values) { state.leases.push({ ...values, rent: Number(values.rent), deposit: Number(values.deposit), rentDay: Number(values.rentDay) }); },
    },
    deposits: {
      title: "新增押金紀錄",
      fields: [
        inputField("id", "押金編號", "text"),
        selectField("roomId", "房號", roomIds),
        selectField("tenantId", "租客", state.tenants.map((item) => ({ value: item.id, label: item.name }))),
        selectField("leaseId", "租約", state.leases.map((item) => ({ value: item.id, label: item.id }))),
        inputField("expected", "押金應收", "number"),
        inputField("received", "押金實收", "number"),
        inputField("receiveDate", "收取日期", "date"),
        selectField("method", "付款方式", ["轉帳", "現金", "其他"]),
        selectField("status", "押金狀態", ["保管中", "部分扣抵", "已退還", "全額扣抵"]),
        inputField("offsetType", "扣抵項目", "text"),
        inputField("offsetAmount", "扣抵金額", "number"),
        inputField("refundAmount", "退還金額", "number"),
        inputField("refundDate", "退還日期", "date"),
        inputField("note", "備註", "text", true),
      ],
      save(values) { state.deposits.push({ ...values, expected: Number(values.expected), received: Number(values.received), offsetAmount: Number(values.offsetAmount || 0), refundAmount: Number(values.refundAmount || 0) }); },
    },
    rent: {
      title: "新增收款",
      fields: [
        inputField("id", "收款編號", "text"),
        inputField("month", "歸屬月份", "month"),
        selectField("roomId", "房號", roomIds),
        selectField("tenantId", "租客", state.tenants.map((item) => ({ value: item.id, label: item.name }))),
        inputField("dueDate", "應收日期", "date"),
        inputField("amountDue", "應收金額", "number"),
        inputField("amountPaid", "實收金額", "number"),
        inputField("paymentDate", "收款日期", "date"),
        selectField("method", "付款方式", ["轉帳", "現金", "其他"]),
        selectField("status", "收款狀態", ["未收", "已收", "部分收", "逾期", "免收"]),
        inputField("note", "備註", "text", true),
      ],
      save(values) { state.rentPayments.push({ ...values, amountDue: Number(values.amountDue), amountPaid: Number(values.amountPaid) }); },
    },
    bills: {
      title: "新增臺電帳單",
      fields: [
        inputField("id", "帳單編號", "text"),
        inputField("name", "帳單名稱", "text"),
        inputField("startDate", "起始日", "date"),
        inputField("endDate", "結束日", "date"),
        inputField("commonFee", "公用電費", "number"),
        inputField("indoorFee", "戶內電費", "number"),
        inputField("dueDate", "繳費期限", "date"),
        selectField("paidToTaipower", "是否已繳臺電", [{ value: "true", label: "是" }, { value: "false", label: "否" }]),
        inputField("payDate", "繳費日期", "date"),
        inputField("note", "備註", "text", true),
      ],
      save(values) { state.bills.push({ ...values, commonFee: Number(values.commonFee), indoorFee: Number(values.indoorFee), paidToTaipower: values.paidToTaipower === "true" }); },
    },
    meters: {
      title: "新增抄表紀錄",
      fields: [
        selectField("billId", "帳單", state.bills.map((item) => ({ value: item.id, label: item.name }))),
        selectField("roomId", "房號", roomIds),
        inputField("id", "抄表編號", "text"),
        inputField("previous", "上期度數", "number"),
        inputField("current", "本期度數", "number"),
        selectField("payer", "負擔者", ["房客", "房東", "前租客", "新租客"]),
        inputField("note", "備註", "text", true),
      ],
      save(values) {
        const previous = Number(values.previous);
        const current = Number(values.current);
        state.meterReadings.push({ ...values, previous, current, usage: current - previous });
      },
    },
    expenses: {
      title: "新增支出",
      fields: [
        inputField("id", "支出編號", "text"),
        inputField("date", "支出日期", "date"),
        inputField("month", "歸屬月份", "month"),
        selectField("type", "支出類型", state.settings.expenseTypes),
        selectField("scope", "支出歸屬", ["公共", ...roomIds]),
        inputField("amount", "金額", "number"),
        selectField("payer", "負擔者", ["房東", "房客", "押金扣抵", "共同分攤"]),
        selectField("paid", "是否已付款", ["是", "否"]),
        inputField("maintenanceId", "關聯維修案件", "text"),
        inputField("note", "備註", "text", true),
      ],
      save(values) { state.expenses.push({ ...values, amount: Number(values.amount) }); },
    },
    maintenance: {
      title: "新增維修案件",
      fields: [
        inputField("id", "案件編號", "text"),
        selectField("roomId", "房號", [...roomIds, "公共區域"]),
        inputField("reporter", "回報人", "text"),
        inputField("requestDate", "申請日期", "date"),
        inputField("type", "維修類型", "text"),
        selectField("status", "處理狀態", state.settings.repairStatuses),
        inputField("cost", "維修金額", "number"),
        selectField("owner", "負擔者", ["房東", "房客", "押金扣抵", "共同分攤"]),
        inputField("vendor", "維修廠商", "text"),
        inputField("doneDate", "完成日期", "date"),
        inputField("note", "備註", "text", true),
      ],
      save(values) { state.maintenance.push({ ...values, cost: Number(values.cost) }); },
    },
    attachments: {
      title: "新增附件紀錄",
      fields: [
        inputField("id", "附件編號", "text"),
        inputField("name", "附件名稱", "text"),
        inputField("type", "附件類型", "text"),
        inputField("module", "關聯模組", "text"),
        inputField("recordId", "關聯資料編號", "text"),
        selectField("roomId", "房號", [{ value: "", label: "無" }, ...roomIds.map((room) => ({ value: room, label: room }))]),
        inputField("tenant", "租客", "text"),
        inputField("date", "日期", "date"),
        inputField("url", "檔案網址", "text"),
        inputField("note", "備註", "text", true),
      ],
      save(values) { state.attachments.push(values); },
    },
  };

  const config = configMap[moduleKey];
  if (!config) return;
  recordForm.innerHTML = `
    <div class="panel-header"><div><p class="overline">新增資料</p><h3>${config.title}</h3></div></div>
    <div class="form-grid">${config.fields.join("")}</div>
    <div class="modal-actions">
      <button class="ghost-btn" type="button" data-close-modal="true">取消</button>
      <button class="cta-btn" type="submit">儲存</button>
    </div>
  `;

  recordForm.querySelector("[data-close-modal]").addEventListener("click", () => recordModal.close());
  recordForm.onsubmit = (event) => {
    event.preventDefault();
    const values = Object.fromEntries(new FormData(recordForm).entries());
    config.save(values);
    persistState();
    recordModal.close();
    render();
  };
  recordModal.showModal();
}

function renderTableModule({ title, subtitle, addKey, addLabel, headers, rows }) {
  return `
    <section class="panel">
      <div class="panel-header">
        <div><p class="overline">${title}</p><h3>${subtitle}</h3></div>
        ${addKey ? `<button class="cta-btn" data-add="${addKey}">${addLabel}</button>` : ""}
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr>${headers.map((header) => `<th>${header}</th>`).join("")}</tr></thead>
          <tbody>${rows || ""}</tbody>
        </table>
      </div>
    </section>
  `;
}

function moduleBanner(title, description) {
  return `<section class="panel"><div class="section-title"><div><p class="overline">${title}</p><h2>${description}</h2></div></div></section>`;
}

function roomHealthCard(room) {
  const payment = monthPayments().find((item) => item.roomId === room.id);
  const paid = payment ? Math.min(payment.amountPaid / payment.amountDue, 1) : 0;
  const score = computeRoomScore(room, paid);
  return `
    <div class="summary-card score-ring">
      <div class="split">
        <strong>${room.id}</strong>
        ${statusTag(room.status)}
      </div>
      <div class="split">
        <span class="score-badge">${score}</span>
        <div class="detail-list">
          <span>健康分數：${score} / 100</span>
          <span>房租收款：${payment ? `${money.format(payment.amountPaid)} / ${money.format(payment.amountDue)}` : "尚無資料"}</span>
        </div>
      </div>
    </div>
  `;
}

function computeRoomScore(room, paidRatio) {
  let score = Math.round(paidRatio * 70);
  if (room.status.includes("起租") || room.status === "出租中") score += 25;
  if (room.status.includes("整修") || room.status.includes("空房")) score -= 20;
  return Math.max(35, Math.min(99, score));
}

function metricCard(label, value, note) {
  return `<article class="metric-card"><p class="overline">${label}</p><strong>${value}</strong><span class="muted">${note}</span></article>`;
}

function expenseBreakdownCard(title, items) {
  return `<article class="summary-card"><span class="overline">${title}</span><strong>${money.format(sumBy(items, "amount"))}</strong><span class="muted">${items.length} 筆紀錄</span></article>`;
}

function statusSummaryCard(label, items) {
  return `<article class="summary-card"><span class="overline">${label}</span><strong>${items.length} 件</strong><span class="muted">${items.slice(0, 2).map((item) => `${item.roomId} ${item.type}`).join("、") || "目前無案件"}</span></article>`;
}

function statusTag(status) {
  const tone = ["已收", "生效中", "出租中", "保管中", "已完成"].includes(status)
    ? "success"
    : ["部分收", "整修中", "處理中", "已派工", "即將到期"].includes(status)
      ? "warning"
      : ["未收", "逾期", "空房", "取消"].includes(status)
        ? "danger"
        : "warning";
  return `<span class="status-tag ${tone}">${status}</span>`;
}

function actionButton(moduleId, label) {
  return `<button class="soft-btn" data-open-module="${moduleId}">${label}</button>`;
}

function buildSummary() {
  const payments = monthPayments();
  const allocation = calculateAllocation(state.bills.at(-1)?.id);
  const rentDue = sumBy(payments, "amountDue");
  const rentPaid = sumBy(payments, "amountPaid");
  const expenses = state.expenses.filter((item) => item.month === state.ui.month && item.payer === "房東").reduce((sum, item) => sum + Number(item.amount), 0);
  const electricDue = sumBy(allocation, "finalDue");
  const electricPaid = sumBy(allocation, "paid");
  const publicExpenses = state.expenses.filter((item) => item.month === state.ui.month && item.scope === "公共").reduce((sum, item) => sum + Number(item.amount), 0);
  const expiringLeases = state.leases.filter((item) => item.status === "即將到期");
  const reminders = [];

  payments.filter((item) => item.amountPaid < item.amountDue).forEach((item) => reminders.push(`${item.roomId} 房租尚差 ${money.format(item.amountDue - item.amountPaid)}`));
  allocation.filter((item) => item.paid < item.finalDue).forEach((item) => reminders.push(`${item.roomId} 電費尚差 ${money.format(item.finalDue - item.paid)}`));
  state.maintenance.filter((item) => ["待處理", "已派工", "處理中"].includes(item.status)).forEach((item) => reminders.push(`${item.roomId} ${item.type} 維修目前為 ${item.status}`));

  return {
    rentDue,
    rentPaid,
    electricDue,
    electricPaid,
    expenses,
    netProfit: rentPaid - expenses,
    publicExpenseShare: publicExpenses / 4,
    reminders: reminders.slice(0, 8),
    expiringLeases,
    unpaidRentRooms: payments.filter((item) => item.amountPaid < item.amountDue),
    unpaidElectricRooms: allocation.filter((item) => item.paid < item.finalDue),
  };
}

function sanitizeState(rawState) {
  const next = { ...rawState };
  next.maintenance = (next.maintenance || []).filter((item) => !String(item.note || "").includes("示意"));
  next.attachments = (next.attachments || []).filter((item) => !String(item.url || "").startsWith("drive://"));
  return next;
}

function getWorkspace() {
  if (state.connection.workspaceMainFolderId && state.connection.workspaceSheetId) {
    return {
      mainFolderId: state.connection.workspaceMainFolderId,
      mainFolderUrl: state.connection.workspaceMainFolderUrl || `https://drive.google.com/drive/folders/${state.connection.workspaceMainFolderId}`,
      sheetId: state.connection.workspaceSheetId,
      sheetUrl: state.connection.workspaceSheetUrl || `https://docs.google.com/spreadsheets/d/${state.connection.workspaceSheetId}/edit`,
      ownerEmail: state.connection.workspaceOwnerEmail,
      folders: state.connection.workspaceFolders || {},
    };
  }
  return DEFAULT_WORKSPACE;
}

function calculateAllocation(billId) {
  if (!billId) return [];
  const bill = state.bills.find((item) => item.id === billId);
  const readings = state.meterReadings.filter((item) => item.billId === billId);
  if (!bill || !readings.length) return [];
  const totalUsage = readings.reduce((sum, item) => sum + Number(item.usage || 0), 0) || 1;
  const totalBill = Number(bill.commonFee) + Number(bill.indoorFee);
  const base = readings.map((reading) => {
    const ratio = Number(reading.usage) / totalUsage;
    const indoor = Math.round(Number(bill.indoorFee) * ratio);
    const common = Math.round(Number(bill.commonFee) * ratio);
    return { roomId: reading.roomId, usage: Number(reading.usage), ratio, indoor, common, due: indoor + common, paid: 0 };
  });
  const diff = totalBill - base.reduce((sum, item) => sum + item.due, 0);
  const maxUsageRoom = [...base].sort((a, b) => b.usage - a.usage)[0];
  return base.map((item) => ({
    ...item,
    tailAdjustment: item.roomId === maxUsageRoom.roomId ? diff : 0,
    finalDue: item.due + (item.roomId === maxUsageRoom.roomId ? diff : 0),
  }));
}

function aggregateLedgerCategories(entries) {
  const map = new Map();
  entries.forEach((item) => {
    const key = item.category || "未分類";
    const current = map.get(key) || { category: key, income: 0, expense: 0, count: 0 };
    current.income += Number(item.income || 0);
    current.expense += Number(item.expense || 0);
    current.count += 1;
    map.set(key, current);
  });
  return [...map.values()].sort((a, b) => (b.income + b.expense) - (a.income + a.expense));
}

function monthPayments() {
  return state.rentPayments.filter((item) => item.month === state.ui.month);
}

function activeTenantByRoom(roomId) {
  return state.tenants.find((item) => item.roomId === roomId && !item.moveOut);
}

function tenantName(tenantId) {
  return state.tenants.find((item) => item.id === tenantId)?.name || "未指定";
}

function sumBy(list, key) {
  return list.reduce((sum, item) => sum + Number(item[key] || 0), 0);
}

function inputField(name, label, type, fullSpan = false) {
  return `<div class="field ${fullSpan ? "full-span" : ""}"><label for="${name}">${label}</label><input id="${name}" name="${name}" type="${type}" /></div>`;
}

function selectField(name, label, options) {
  const normalized = options.map((option) => typeof option === "string" ? { value: option, label: option } : option);
  return `
    <div class="field">
      <label for="${name}">${label}</label>
      <select id="${name}" name="${name}">
        ${normalized.map((option) => `<option value="${escapeAttr(option.value)}">${option.label}</option>`).join("")}
      </select>
    </div>
  `;
}

function formatDateTime(value) {
  if (!value) return "尚未更新";
  try {
    return new Date(value).toLocaleString("zh-TW");
  } catch {
    return value;
  }
}

function loadGoogleSession() {
  try {
    const raw = JSON.parse(sessionStorage.getItem("rental-google-session") || "{}");
    return {
      accessToken: raw.accessToken || "",
      expiresAt: Number(raw.expiresAt || 0),
      profile: raw.profile || null,
    };
  } catch {
    return { accessToken: "", expiresAt: 0, profile: null };
  }
}

function saveGoogleSession() {
  sessionStorage.setItem("rental-google-session", JSON.stringify({
    accessToken: googleAccessToken,
    expiresAt: googleAccessTokenExpiresAt,
    profile: googleProfile,
  }));
}

function hasActiveGoogleToken() {
  return Boolean(googleAccessToken && googleAccessTokenExpiresAt > Date.now() + 60_000);
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#039;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeAttr(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("\"", "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}
