const LOCAL_KEY = "tenant-rental-system-v4";
const GOOGLE_CLIENT_ID = window.RENTAL_GOOGLE_CLIENT_ID || localStorage.getItem("rental-google-client-id") || "";
const GOOGLE_SCOPES = "openid email profile https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/spreadsheets";
const DEFAULT_WORKSPACE = {
  mainFolderId: "1EkzIQ4_0vVjruU0j8jED9FwV6hBkS43k",
  mainFolderUrl: "https://drive.google.com/drive/folders/1EkzIQ4_0vVjruU0j8jED9FwV6hBkS43k",
  sheetId: "1GNLdYv7inuU4U1jSlsG403XTqwbFbx2VfFmm1RYDu7w",
  sheetUrl: "https://docs.google.com/spreadsheets/d/1GNLdYv7inuU4U1jSlsG403XTqwbFbx2VfFmm1RYDu7w/edit?usp=drivesdk",
  ownerEmail: "craigpop.tw@gmail.com",
  folders: {
    excel: "1cOMrkpd8le8wgsJB3A6xGuXcb0pJDpv-",
    leases: "1__J1hdE2d-D9dgQ5q2yLRg2uL_EKZLF-",
    taipower: "12iKbeUDjiXPBGWwSN5V6sS_Mc0eo_ams",
    meters: "1XjahCEKLVQH2oix4tOzS8R_wp1N3qim4",
    maintenance: "1I8mKvwxCxCKFWw0D9ob0JG4ofW8xCu-t",
    receipts: "1UkOwtFphiDorG3UH1wnCefmH6jVC8W4c",
    tenants: "1uoLJOoxRuD5JVQzkYrpuww2vnFVfNaOC",
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
  managementFees: { id: "managementFees", title: "代租代管", group: "finance" },
  ledger: { id: "ledger", title: "總收支明細", group: "finance" },
  maintenance: { id: "maintenance", title: "維修處理", group: "operations" },
  incidents: { id: "incidents", title: "事件紀錄", group: "operations" },
  attachments: { id: "attachments", title: "附件中心", group: "operations" },
  lifecycle: { id: "lifecycle", title: "年度保存", group: "analysis" },
  reports: { id: "reports", title: "經營分析", group: "analysis" },
  settings: { id: "settings", title: "雲端與母版", group: "analysis" },
};

const NAV_GROUPS = [
  { id: "overview", label: "營運總覽", modules: ["dashboard", "rooms"] },
  { id: "movein", label: "入住與租約", modules: ["tenants", "leases", "deposits"] },
  { id: "finance", label: "收款與帳務", modules: ["rent", "bills", "meters", "allocations", "expenses", "managementFees", "ledger"] },
  { id: "operations", label: "維修與附件", modules: ["maintenance", "incidents", "attachments"] },
  { id: "analysis", label: "報表與儲存", modules: ["reports", "lifecycle", "settings"] },
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
let backendSyncTimer = null;
if (monthSelector) monthSelector.value = state.ui.month;

bootstrap();

function bootstrap() {
  bindFrameEvents();
  render();
}

function buildExcelPreset() {
  return {
    version: 4,
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
    managementFees: [
      { id: "MF-2026-04-304-A", date: "2026-04-08", month: "2026-04", year: "2026", roomId: "304", tenantId: "TEN-304", feeType: "代租媒合費", baseAmount: 16000, rate: 50, amount: 8000, vendor: "戴先生", status: "已支付", contractId: "LE-304-202604", note: "Excel 記帳分表：戴先生仲介費（半個月）。" },
      { id: "MF-2026-04-301-A", date: "2026-04-15", month: "2026-04", year: "2026", roomId: "301", tenantId: "TEN-301", feeType: "代租媒合費", baseAmount: 16000, rate: 50, amount: 8000, vendor: "戴先生", status: "已支付", contractId: "LE-301-202604", note: "Excel 記帳分表：戴先生仲介費（半個月）。" },
      { id: "MF-2026-04-303-A", date: "2026-04-17", month: "2026-04", year: "2026", roomId: "303", tenantId: "TEN-303", feeType: "代租媒合費", baseAmount: 15500, rate: 50, amount: 7750, vendor: "陳先生", status: "已支付", contractId: "LE-303-202604", note: "Excel 記帳分表：陳先生仲介費（半個月）。" },
      { id: "MF-2026-04-302-A", date: "2026-04-18", month: "2026-04", year: "2026", roomId: "302", tenantId: "TEN-302", feeType: "代租媒合費", baseAmount: 15000, rate: 50, amount: 7500, vendor: "呂小姐", status: "已支付", contractId: "LE-302-202604", note: "Excel 記帳分表：呂小姐仲介費（半個月）。" },
      { id: "MF-2026-04-301-M", date: "2026-04-18", month: "2026-04", year: "2026", roomId: "301", tenantId: "TEN-301", feeType: "代管月費", baseAmount: 16000, rate: 50, amount: 8000, vendor: "朱小姐", status: "已支付", contractId: "LE-301-202604", note: "Excel 記帳分表：朱小姐代管服務費（年繳）。" },
      { id: "MF-2026-04-302-M", date: "2026-04-18", month: "2026-04", year: "2026", roomId: "302", tenantId: "TEN-302", feeType: "代管月費", baseAmount: 15000, rate: 50, amount: 7500, vendor: "朱小姐", status: "已支付", contractId: "LE-302-202604", note: "Excel 記帳分表：朱小姐代管服務費（年繳）。" },
      { id: "MF-2026-04-303-M", date: "2026-04-18", month: "2026-04", year: "2026", roomId: "303", tenantId: "TEN-303", feeType: "代管月費", baseAmount: 15500, rate: 50, amount: 7750, vendor: "朱小姐", status: "已支付", contractId: "LE-303-202604", note: "Excel 記帳分表：朱小姐代管服務費（年繳）。" },
      { id: "MF-2026-04-304-M", date: "2026-04-18", month: "2026-04", year: "2026", roomId: "304", tenantId: "TEN-304", feeType: "代管月費", baseAmount: 16000, rate: 50, amount: 8000, vendor: "朱小姐", status: "已支付", contractId: "LE-304-202604", note: "Excel 記帳分表：朱小姐代管服務費（年繳）。" },
    ],
    maintenance: [],
    incidents: [],
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
      { id: "LG-011", date: "2026-06-04", month: "2026-06", roomId: "304", item: "6月房租", category: "房租", income: 16000, expense: 0, note: "依最新營運狀態補正。" },
      { id: "LG-012", date: "2026-06-15", month: "2026-06", roomId: "302", item: "6月房租", category: "房租", income: 15000, expense: 0, note: "依最新營運狀態補正。" },
      { id: "LG-013", date: "2026-06-18", month: "2026-06", roomId: "303", item: "6月房租", category: "房租", income: 15500, expense: 0, note: "依最新營運狀態補正。" },
      { id: "LG-014", date: "2026-06-18", month: "2026-06", roomId: "301", item: "6月房租", category: "房租", income: 16000, expense: 0, note: "依最新營運狀態補正。" },
      { id: "LG-015", date: "2026-07-01", month: "2026-06", roomId: "公共", item: "5-6月臺電帳單已繳", category: "電費", income: 0, expense: 3871, note: "總表載明已繳。" },
      { id: "LG-016", date: "2026-07-03", month: "2026-07", roomId: "304", item: "7月房租", category: "房租", income: 16000, expense: 0, note: "" },
      { id: "LG-017", date: "2026-07-10", month: "2026-07", roomId: "公共", item: "垃圾清潔費", category: "垃圾清潔費", income: 0, expense: 1200, note: "" },
      { id: "LG-018", date: "2026-07-15", month: "2026-07", roomId: "302", item: "7月房租", category: "房租", income: 15000, expense: 0, note: "" },
      { id: "LG-019", date: "2026-07-18", month: "2026-07", roomId: "303", item: "7月房租", category: "房租", income: 15500, expense: 0, note: "" },
      { id: "LG-020", date: "2026-07-18", month: "2026-07", roomId: "301", item: "7月房租", category: "房租", income: 16000, expense: 0, note: "" },
    ],
    settings: {
      roomStatuses: ["出租中", "空房", "整修中", "即將退租"],
      expenseTypes: ["水費", "垃圾清潔費", "維修", "設備", "清潔", "耗材", "稅費", "保險", "社區管理費", "銀行手續費", "公證費", "仲介費", "代租費", "代管費", "廣告刊登", "法律諮詢", "租賃行政", "其他"],
      repairStatuses: ["待處理", "已派工", "處理中", "已完成", "取消"],
      incidentTypes: ["鄰里投訴", "違約", "逾期未繳", "水電異常", "安全事件", "災損", "設備遺失", "行政通知", "其他"],
      managementFeeTypes: ["代租媒合費", "代管月費", "續約服務費", "廣告刊登費", "清潔交屋費", "點交服務費", "其他"],
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
    managementFees: [],
    maintenance: [],
    incidents: [],
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
    managementFees: Array.isArray(raw.managementFees) ? raw.managementFees : seed.managementFees,
    maintenance: Array.isArray(raw.maintenance) ? raw.maintenance : seed.maintenance,
    incidents: Array.isArray(raw.incidents) ? raw.incidents : seed.incidents,
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
  scheduleBackendSync();
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
  const onLocalFile = window.location.protocol === "file:";
  const oauthState = signedIn ? "已登入並可同步" : GOOGLE_CLIENT_ID ? "可登入，但第一次需完成 Google 授權" : "尚未設定 OAuth Client ID";
  const statusClass = signedIn ? "success" : GOOGLE_CLIENT_ID ? "" : "danger";
  const statusText = signedIn
    ? activeToken
      ? `已登入：${googleProfile.email}`
      : `已記住帳號：${googleProfile.email}，請重新授權`
    : GOOGLE_CLIENT_ID
      ? "尚未登入 Google"
      : "Google 登入尚未啟用";
  loginPanel.innerHTML = `
    <div class="login-hero">
      <div>
        <p class="overline">Google 雲端工作區</p>
        <h3>${signedIn ? escapeHtml(googleProfile.name || "Google 使用者") : "登入 Google 帳號接手資料與附件"}</h3>
      </div>
      <div class="sync-pill ${statusClass}">${escapeHtml(statusText)}</div>
    </div>
    <div class="status-strip">
      <div class="status-chip"><strong>登入狀態</strong><span>${oauthState}</span></div>
      <div class="status-chip"><strong>資料擁有者</strong><span>${escapeHtml(workspace.ownerEmail || "尚未指定")}</span></div>
      <div class="status-chip"><strong>後臺主表</strong><span>${workspace.sheetId ? "已建立" : "尚未建立"}</span></div>
    </div>
    <div class="meta-line">流程：開啟網站 → 使用 Google 登入 → 讀取後臺主表 → 在 Drive 子資料夾上傳照片或 PDF → 自動把連結回寫主表。</div>
    <div class="meta-line">${onLocalFile ? "你現在是用本機 file:// 開啟，Google OAuth 會失敗；請改用 GitHub Pages 網址登入。" : "請從正式網站網址登入，不要用本機 file:// 版本做 Google OAuth。"} ${accessMismatch ? "你現在登入的帳號和資料夾擁有者不同，所以打開資料夾時會跳權限申請。" : "若使用同一個擁有者帳號登入，資料會直接寫進同一組雲端資料夾。"} </div>
    <div class="meta-line">若 Google 顯示「這個應用程式未經 Google 驗證」，代表 OAuth consent screen 還在測試狀態，第一次授權需按「進階」再繼續。</div>
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
    managementFees: renderManagementFees,
    ledger: renderLedger,
    maintenance: renderMaintenance,
    incidents: renderIncidents,
    attachments: renderAttachments,
    lifecycle: renderLifecycle,
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
  const score = operatingScore(summary);
  const latestBill = state.bills.at(-1);
  const focusItems = [
    `本月房租 ${money.format(summary.rentPaid)} / ${money.format(summary.rentDue)}，${summary.unpaidRentRooms.length ? `仍有 ${summary.unpaidRentRooms.length} 間待追蹤。` : "四間房本月都已完成收款。"}`,
    latestBill ? `最近一期臺電帳單為 ${latestBill.name}，總額 ${money.format(Number(latestBill.commonFee) + Number(latestBill.indoorFee))}，${latestBill.paidToTaipower ? "已完成對台電付款。" : "尚未繳納。"}`
      : "目前尚未建立台電帳單資料。",
    summary.expenses ? `本月已記錄支出 ${money.format(summary.expenses)}，集中於 ${topExpenseLabel(state.ui.month)}。` : "本月尚未建立額外支出資料。",
  ];
  const progressCards = [
    { title: "營運總分", value: `${score} 分`, note: scoreLabel(score) },
    { title: "入住建檔", value: `${state.tenants.length} 位租客`, note: `${state.leases.filter((item) => item.status === "生效中").length} 份有效租約` },
    { title: "本月收租", value: `${money.format(summary.rentPaid)} / ${money.format(summary.rentDue)}`, note: summary.unpaidRentRooms.length ? `${summary.unpaidRentRooms.length} 間待追款` : "本月已全數收款" },
    { title: "同步時間", value: formatDateTime(state.connection.lastSyncedAt), note: workspaceStatusText() },
  ];

  return `
    <section class="hero-banner">
      <div class="hero-copy stack">
        <div class="hero-title-row">
          <div>
            <p class="overline">今日重點</p>
            <h2>用一個高辨識度的營運首頁，先看本月租金、台電、支出與附件是否都已經歸位。</h2>
          </div>
          <div class="score-badge ${scoreTone(score)}">${score}</div>
        </div>
        <p class="top-note">首頁只顯示目前真實存在的資料。文字與數字優先回寫後臺主表，照片、PDF、租約與收據則分類存入 Google Drive 子資料夾。</p>
        <div class="process-inline process-inline-compact">
          <span class="flow-inline"><strong>1.</strong> Google 登入並讀取主表</span>
          <span class="flow-inline"><strong>2.</strong> 每月收租、抄表、記支出</span>
          <span class="flow-inline"><strong>3.</strong> 上傳附件並自動回寫連結</span>
        </div>
        <div class="hero-actions">
          ${actionButton("rent", "補登房租")}
          ${actionButton("bills", "查看電費")}
          ${actionButton("maintenance", "新增維修")}
        </div>
      </div>
      <div class="hero-side hero-side-rich">
        ${progressCards.map((item) => `
          <article class="summary-card">
            <span class="overline">${item.title}</span>
            <strong>${item.value}</strong>
            <span class="muted">${item.note}</span>
          </article>
        `).join("")}
      </div>
    </section>
    <section class="panel spotlight-panel">
      <div class="panel-header"><div><p class="overline">本月總覽</p><h3>${state.ui.month} 你真正需要確認的 3 件事</h3></div></div>
      <div class="focus-grid">
        ${focusItems.map((item, index) => `<article class="focus-card"><span class="focus-index">0${index + 1}</span><p>${item}</p></article>`).join("")}
      </div>
    </section>
    <section class="metric-grid">
      ${metricCard("本月房租已收", money.format(summary.rentPaid), `尚未收 ${money.format(summary.rentDue - summary.rentPaid)}`)}
      ${metricCard("最近一期電費", money.format(summary.electricDue), latestBill?.paidToTaipower ? "已對台電付款" : "待繳台電")}
      ${metricCard("本月支出", money.format(summary.expenses), "房東負擔支出")}
      ${metricCard("本月淨利", money.format(summary.netProfit), summary.reminders.length ? `${summary.reminders.length} 則待辦提醒` : "目前沒有待辦")}
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
  const monthlyManagementFees = state.managementFees.filter((item) => item.month === state.ui.month);
  return `
    ${moduleBanner("支出管理", "公共支出與房間支出分開記錄，方便後續計算整體與房間別損益。")}
    <section class="three-grid">
      ${expenseBreakdownCard("公共支出", monthly.filter((item) => item.scope === "公共"))}
      ${expenseBreakdownCard("房間支出", monthly.filter((item) => item.scope !== "公共"))}
      ${metricCard("代租代管", money.format(sumBy(monthlyManagementFees, "amount")), `${monthlyManagementFees.length} 筆服務費`)}
    </section>
    ${renderTableModule({
      title: "支出明細",
      subtitle: "支出類型涵蓋水費、清潔、設備、稅費、保險、社區管理、法律、公證、代租與代管等長期經營項目。",
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

function renderManagementFees() {
  const monthly = state.managementFees.filter((item) => item.month === state.ui.month);
  const yearly = state.managementFees.filter((item) => item.year === selectedYear());
  const categoryRows = aggregateBy(monthly, "feeType", "amount");
  return `
    ${moduleBanner("代租代管", "把代租媒合、代管月費、續約服務、廣告刊登與點交費用獨立管理，避免混在一般維修支出。")}
    <section class="three-grid">
      ${metricCard("本月服務費", money.format(sumBy(monthly, "amount")), `${monthly.length} 筆紀錄`)}
      ${metricCard("本年度服務費", money.format(sumBy(yearly, "amount")), `${selectedYear()} 年累計`)}
      ${metricCard("關聯租約", `${new Set(yearly.map((item) => item.contractId).filter(Boolean)).size} 份`, "可回查租約與房號")}
    </section>
    <section class="two-grid">
      <article class="panel chart-panel">
        <div class="panel-header"><div><p class="overline">服務費結構</p><h3>${state.ui.month}</h3></div></div>
        ${pieChart(categoryRows.map((item, index) => ({ label: item.key, value: item.total, color: chartColor(index) })))}
      </article>
      ${renderTableModule({
        title: "代租代管明細",
        subtitle: "每筆費用都保留月份、年度、房號、租客、租約與業者，未來做年度比較時不會混帳。",
        addKey: "managementFees",
        addLabel: "新增服務費",
        headers: ["日期", "類型", "房號", "租客", "計算基礎", "金額", "業者", "狀態"],
        rows: state.managementFees.map((item) => `
          <tr>
            <td>${item.date}</td>
            <td>${item.feeType}</td>
            <td>${item.roomId || "公共"}</td>
            <td>${tenantName(item.tenantId)}</td>
            <td>${money.format(item.baseAmount)} × ${item.rate || 0}%</td>
            <td>${money.format(item.amount)}</td>
            <td>${item.vendor || "未填"}</td>
            <td>${statusTag(item.status)}</td>
          </tr>
        `).join(""),
      })}
    </section>
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

function renderIncidents() {
  const monthly = state.incidents.filter((item) => item.month === state.ui.month);
  const openItems = state.incidents.filter((item) => !["已結案", "取消"].includes(item.status));
  return `
    ${moduleBanner("事件紀錄", "把違約、鄰里投訴、災損、逾期、水電異常與行政通知獨立建檔，保留責任歸屬與佐證。")}
    <section class="three-grid">
      ${statusSummaryCard("本月事件", monthly)}
      ${statusSummaryCard("未結案", openItems)}
      ${metricCard("佐證附件", `${state.attachments.filter((item) => item.module === "事件").length} 件`, "可連到照片、PDF、通知截圖")}
    </section>
    ${renderTableModule({
      title: "事件清單",
      subtitle: "突發事件可以關聯房號、租客、租約、附件與後續維修或支出。",
      addKey: "incidents",
      addLabel: "新增事件",
      headers: ["事件編號", "日期", "類型", "房號", "租客", "嚴重度", "狀態", "處理摘要"],
      rows: state.incidents.map((item) => `
        <tr>
          <td>${item.id}</td>
          <td>${item.date}</td>
          <td>${item.type}</td>
          <td>${item.roomId || "公共"}</td>
          <td>${tenantName(item.tenantId)}</td>
          <td>${statusTag(item.severity)}</td>
          <td>${statusTag(item.status)}</td>
          <td>${item.action || item.note || "未填"}</td>
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

function renderLifecycle() {
  const year = selectedYear();
  const yearMonths = Array.from({ length: 12 }, (_, index) => `${year}-${String(index + 1).padStart(2, "0")}`);
  const yearLedger = state.ledgerEntries.filter((item) => item.month?.startsWith(`${year}-`));
  const yearFees = state.managementFees.filter((item) => item.year === year || item.month?.startsWith(`${year}-`));
  const relationRows = buildRelationRows();
  const archiveRules = [
    { title: "年 / 月層級", note: "所有記錄都用 YYYY-MM，跨年租約也能追蹤。" },
    { title: "租客獨立 ID", note: "同一人換房或二次入住，不會和舊資料混在一起。" },
    { title: "附件回寫主表", note: "照片與 PDF 存 Drive，主表只留連結與關聯欄位。" },
  ];
  return `
    ${moduleBanner("年度保存", "把年份、月份、租客、租約與附件的關係做成可長期保存的經營年表。")}
    <section class="metric-grid">
      ${metricCard("年度收入", money.format(sumBy(yearLedger, "income")), `${year} 年`)}
      ${metricCard("年度支出", money.format(sumBy(yearLedger, "expense") + sumBy(yearFees, "amount")), "含代租代管")}
      ${metricCard("資料關聯完整度", `${relationScore()} / 100`, "附件、租約、事件、帳務 ID 檢查")}
      ${metricCard("永久保存節點", `${state.attachments.length} 件`, "Drive 連結與後臺主表互相對應")}
    </section>
    <section class="two-grid">
      <article class="panel timeline-panel">
        <div class="panel-header"><div><p class="overline">年度月份軸</p><h3>${year} 經營資料密度</h3></div></div>
        <div class="month-timeline">
          ${yearMonths.map((month) => {
            const records = recordCountByMonth(month);
            return `<div class="month-node ${records ? "active" : ""}"><strong>${month.slice(5)}</strong><span>${records} 筆</span></div>`;
          }).join("")}
        </div>
      </article>
      <article class="panel">
        <div class="panel-header"><div><p class="overline">保存規則</p><h3>避免資料混淆的 3 個原則</h3></div></div>
        <div class="stack">
          ${archiveRules.map((item) => `
            <article class="rule-card compact">
              <strong>${item.title}</strong>
              <span>${item.note}</span>
            </article>
          `).join("")}
        </div>
      </article>
    </section>
    <section class="panel">
      <div class="panel-header"><div><p class="overline">最近關聯資料</p><h3>近期建立或更新的記錄</h3></div></div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>資料</th><th>月份</th><th>房號</th><th>租客</th><th>附件</th></tr></thead>
          <tbody>
            ${relationRows.map((item) => `
              <tr>
                <td>${item.label}</td>
                <td>${item.month}</td>
                <td>${item.roomId || "公共"}</td>
                <td>${item.tenant}</td>
                <td>${item.attachments} 件</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderReports() {
  const summary = buildSummary();
  const score = operatingScore(summary);
  const roomRows = roomIds.map((roomId) => {
    const paidRent = monthPayments().find((item) => item.roomId === roomId)?.amountPaid || 0;
    const roomExpenses = state.expenses.filter((item) => item.month === state.ui.month && item.scope === roomId).reduce((sum, item) => sum + Number(item.amount), 0);
    return { roomId, paidRent, roomExpenses, publicShare: summary.publicExpenseShare, net: paidRent - roomExpenses - summary.publicExpenseShare };
  });
  const ledgerEntries = state.ledgerEntries.filter((item) => item.month === state.ui.month);
  const expenseCategories = aggregateLedgerCategories(ledgerEntries.filter((item) => item.expense > 0));
  const scoreRows = buildScoreRows(summary);
  const monthTrend = buildMonthlyTrend(selectedYear());

  return `
    ${moduleBanner("經營分析", "用更直觀的分數、趨勢與分類圖來看懂本月狀態，不再用難理解的雷達圖。")}
    <section class="metric-grid">
      ${metricCard("月報房租實收", money.format(summary.rentPaid), state.ui.month)}
      ${metricCard("年度化房租", money.format(summary.rentPaid * 12), "依目前月資料估算")}
      ${metricCard("年度化支出", money.format(summary.expenses * 12), "依目前月資料估算")}
      ${metricCard("年度化淨利", money.format(summary.netProfit * 12), "依目前月資料估算")}
    </section>
    <section class="two-grid">
      <article class="panel chart-panel">
        <div class="panel-header"><div><p class="overline">支出圓餅圖</p><h3>${state.ui.month} 支出結構</h3></div></div>
        ${pieChart(expenseCategories.map((item, index) => ({ label: item.category, value: item.expense, color: chartColor(index) })))}
      </article>
      <article class="panel score-panel">
        <div class="panel-header"><div><p class="overline">營運分數</p><h3>本月健康檢查</h3></div></div>
        <div class="score-head">
          <div class="score-badge ${scoreTone(score)}">${score}</div>
          <div class="stack">
            <strong class="score-label">${scoreLabel(score)}</strong>
            <span class="muted">分數綜合收租完成率、房間滿租度、資料完整度與待處理風險。</span>
          </div>
        </div>
        <div class="score-list">
          ${scoreRows.map((item) => `
            <div class="score-row">
              <div class="split"><strong>${item.label}</strong><span>${item.value} / 100</span></div>
              <div class="progress-track"><div class="progress-fill" style="width:${item.value}%;"></div></div>
            </div>
          `).join("")}
        </div>
      </article>
    </section>
    <section class="panel">
      <div class="panel-header"><div><p class="overline">年度趨勢</p><h3>${selectedYear()} 月別收入與支出</h3></div></div>
      <div class="trend-list">
        ${monthTrend.map((item) => `
          <div class="trend-row">
            <strong>${item.month}</strong>
            <div class="trend-bars">
              <span class="trend-bar income" style="width:${item.incomeWidth}%;"></span>
              <span class="trend-bar expense" style="width:${item.expenseWidth}%;"></span>
            </div>
            <span class="trend-meta">收 ${money.format(item.income)} / 支 ${money.format(item.expense)}</span>
          </div>
        `).join("")}
      </div>
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
    ${moduleBanner("雲端與母版", "把登入、主表、照片資料夾與預設母版整理成看得懂的雲端工作區。")}
    <section class="two-grid">
      <article class="panel">
        <div class="panel-header">
          <div><p class="overline">Google 雲端入口</p><h3>主資料夾、主表與附件結構</h3></div>
        </div>
        <div class="sync-detail">
          <div class="alert-card">目前資料應存到這裡：主資料夾屬於 ${workspace.ownerEmail || "尚未設定"}，後臺主表放在 01_Excel表，照片與 PDF 依序進入 ${folderNames} 這些子資料夾。</div>
          <div class="toolbar">
            <a class="cta-btn" href="${workspace.mainFolderUrl}" target="_blank" rel="noreferrer">開啟主資料夾</a>
            <a class="soft-btn" href="${workspace.sheetUrl}" target="_blank" rel="noreferrer">開啟後臺主表</a>
          </div>
          <div class="storage-map">
            ${WORKSPACE_BLUEPRINT.map((item, index) => `<div class="storage-node"><span>${String(index + 1).padStart(2, "0")}</span><strong>${item.name}</strong></div>`).join("")}
          </div>
        </div>
      </article>
      <article class="panel">
        <div class="panel-header">
          <div><p class="overline">預設資料</p><h3>以 Excel 工作簿作為系統母版</h3></div>
        </div>
        <p class="list-note">目前預設資料已依循 \`20260322-新埔八街-收租財報分頁整合.xlsx\` 的房號、租客、租金、押金、仲介費、代管費與每月收支。若只是想清空回到空白專案，也可以一鍵重置。</p>
        <div class="toolbar wrap-top">
          <button class="cta-btn" id="apply-excel-preset">套用 Excel 預設</button>
          <button class="danger-btn" id="clear-all-data">清空所有資料</button>
        </div>
      </article>
    </section>
    <section class="three-grid">
      <article class="panel">
        <p class="overline">共享與權限</p>
        <div class="detail-list"><span>建議房東、管理者、代管人員都使用同一組共享 Drive 資料夾。</span><span>若登入帳號不是擁有者，請先由擁有者在 Drive 內直接分享資料夾與主表。</span></div>
      </article>
      <article class="panel">
        <p class="overline">資料更新時間</p>
        <div class="detail-list"><span>${formatDateTime(state.connection.lastSyncedAt)}</span><span>${workspaceStatusText()}</span></div>
      </article>
      <article class="panel">
        <p class="overline">目前母版內容</p>
        <div class="detail-list"><span>${state.rooms.length} 間房、${state.tenants.length} 位租客、${state.ledgerEntries.length} 筆總帳。</span><span>若主表看起來空白，通常代表 Google 授權尚未完成或尚未同步回寫。</span></div>
      </article>
    </section>
    <section class="two-grid">
      <article class="panel">
        <p class="overline">管理費分類</p>
        <div class="inline-pills">${state.settings.managementFeeTypes.map((item) => `<span class="chip">${item}</span>`).join("")}</div>
      </article>
      <article class="panel">
        <p class="overline">主要支出分類</p>
        <div class="inline-pills">${state.settings.expenseTypes.slice(0, 10).map((item) => `<span class="chip">${item}</span>`).join("")}</div>
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
        await initializeActiveWorkspaceSheets({ silent: true });
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

async function initializeActiveWorkspaceSheets({ silent = false } = {}) {
  if (!hasActiveGoogleToken()) return false;
  const workspace = getWorkspace();
  if (!workspace.sheetId) return false;
  if (workspace.ownerEmail && googleProfile?.email && workspace.ownerEmail !== googleProfile.email) return false;

  try {
    await ensureBackendSheets(workspace.sheetId);
    state.connection.lastSyncedAt = new Date().toISOString();
    persistState();
    return true;
  } catch (error) {
    if (!silent) window.alert(`初始化後臺主表失敗：${error.message}`);
    return false;
  }
}

function scheduleBackendSync() {
  if (!hasActiveGoogleToken() || !googleProfile?.email) return;
  const workspace = getWorkspace();
  if (!workspace.sheetId) return;
  if (workspace.ownerEmail && workspace.ownerEmail !== googleProfile.email) return;
  window.clearTimeout(backendSyncTimer);
  backendSyncTimer = window.setTimeout(() => {
    syncStateToBackend().catch((error) => {
      console.warn("後臺主表同步失敗", error);
    });
  }, 1200);
}

async function syncStateToBackend() {
  const workspace = getWorkspace();
  if (!hasActiveGoogleToken() || !workspace.sheetId) return false;
  await ensureBackendSheets(workspace.sheetId);

  const tables = buildBackendTables();
  for (const table of tables) {
    const clearRange = encodeURIComponent(`${table.sheet}!A2:Z`);
    await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${workspace.sheetId}/values/${clearRange}:clear`, {
      method: "POST",
      headers: { Authorization: `Bearer ${googleAccessToken}` },
    });

    if (!table.rows.length) continue;
    const updateRange = encodeURIComponent(`${table.sheet}!A2`);
    const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${workspace.sheetId}/values/${updateRange}?valueInputOption=USER_ENTERED`, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${googleAccessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ values: table.rows }),
    });
    if (!response.ok) throw new Error(await response.text());
  }

  state.connection.lastSyncedAt = new Date().toISOString();
  saveLocalState();
  return true;
}

function buildBackendTables() {
  return [
    {
      sheet: "房間資料",
      rows: state.rooms.map((item) => [item.id, item.name, item.status, item.rent, item.deposit, item.rentDay, item.equipment, item.note, state.connection.lastSyncedAt]),
    },
    {
      sheet: "租客資料",
      rows: state.tenants.map((item) => [item.id, item.name, item.phone, item.line, item.job, item.emergency, item.roomId, item.moveIn, item.moveOut, item.note, state.connection.lastSyncedAt]),
    },
    {
      sheet: "租約資料",
      rows: state.leases.map((item) => [item.id, item.roomId, item.tenantId, item.startDate, item.endDate, item.rent, item.deposit, item.rentDay, item.status, attachmentLinksFor(item.id), item.note, state.connection.lastSyncedAt]),
    },
    {
      sheet: "收款紀錄",
      rows: state.rentPayments.map((item) => [item.id, item.month, item.roomId, item.tenantId, item.dueDate, item.amountDue, item.amountPaid, item.paymentDate, item.method, item.status, attachmentLinksFor(item.id), item.note]),
    },
    {
      sheet: "臺電帳單",
      rows: state.bills.map((item) => [item.id, item.name, item.startDate, item.endDate, item.commonFee, item.indoorFee, item.dueDate, item.paidToTaipower ? "是" : "否", item.payDate, attachmentLinksFor(item.id), item.note]),
    },
    {
      sheet: "電表抄表",
      rows: state.meterReadings.map((item) => [item.id, item.billId, item.roomId, item.previous, item.current, item.usage, item.payer, attachmentLinksFor(item.id), item.note]),
    },
    {
      sheet: "支出紀錄",
      rows: state.expenses.map((item) => [item.id, item.date, item.month, item.type, item.scope, item.amount, item.payer, item.paid, item.maintenanceId, attachmentLinksFor(item.id), item.note]),
    },
    {
      sheet: "代租代管",
      rows: state.managementFees.map((item) => [item.id, item.date, item.month, item.year, item.feeType, item.roomId, item.tenantId, item.contractId, item.baseAmount, item.rate, item.amount, item.vendor, item.status, attachmentLinksFor(item.id), item.note]),
    },
    {
      sheet: "維修紀錄",
      rows: state.maintenance.map((item) => [item.id, item.roomId, item.reporter, item.requestDate, item.type, item.status, item.cost, item.owner, item.vendor, item.doneDate, attachmentLinksFor(item.id), item.note]),
    },
    {
      sheet: "事件紀錄",
      rows: state.incidents.map((item) => [item.id, item.date, item.month, item.type, item.roomId, item.tenantId, item.leaseId, item.severity, item.status, item.action, item.evidenceId, item.note]),
    },
    {
      sheet: "附件中心",
      rows: state.attachments.map((item) => [item.id, item.date, item.module, item.type, item.roomId, item.tenant, item.recordId, item.name, item.url, item.note]),
    },
    {
      sheet: "總帳明細",
      rows: state.ledgerEntries.map((item) => [item.id, item.date, item.month, item.roomId, item.item, item.category, item.income, item.expense, attachmentLinksFor(item.id), item.note]),
    },
    {
      sheet: "年度索引",
      rows: buildRelationRows().map((item) => [item.month?.slice(0, 4), item.month, item.roomId, item.tenant, "", item.label.split(" ")[0], item.label, item.attachments, getWorkspace().mainFolderUrl, ""]),
    },
  ];
}

function attachmentLinksFor(recordId) {
  return state.attachments
    .filter((item) => item.recordId === recordId)
    .map((item) => item.url)
    .filter(Boolean)
    .join("\n");
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
  const metaResponse = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets.properties.title`, {
    headers: { Authorization: `Bearer ${googleAccessToken}` },
  });
  if (!metaResponse.ok) throw new Error(await metaResponse.text());
  const meta = await metaResponse.json();
  const existingSheets = new Set((meta.sheets || []).map((sheet) => sheet.properties?.title).filter(Boolean));

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
    ["代租代管", ["服務費編號", "日期", "月份", "年度", "費用類型", "房號", "租客編號", "租約編號", "計算基礎", "比例", "金額", "業者", "狀態", "附件連結", "備註"]],
    ["維修紀錄", ["維修編號", "房號", "通報者", "通報日", "類型", "狀態", "費用", "負擔者", "廠商", "完成日", "附件連結", "備註"]],
    ["事件紀錄", ["事件編號", "日期", "月份", "事件類型", "房號", "租客編號", "租約編號", "嚴重度", "狀態", "處理摘要", "附件編號", "備註"]],
    ["附件中心", ["附件編號", "日期", "模組", "類型", "房號", "租客", "關聯紀錄", "檔名", "Drive連結", "備註"]],
    ["總帳明細", ["明細編號", "日期", "月份", "房號", "項目", "分類", "收入", "支出", "附件連結", "備註"]],
    ["年度索引", ["年度", "月份", "房號", "租客編號", "租約編號", "資料類型", "資料編號", "附件數", "保存位置", "備註"]],
  ];
  sheets
    .filter((item) => !existingSheets.has(item[0]))
    .forEach((item) => requests.push({ addSheet: { properties: { title: item[0] } } }));

  if (requests.length) {
    const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${googleAccessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ requests }),
    });
    if (!response.ok) throw new Error(await response.text());
  }

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
    managementFees: {
      title: "新增代租代管費用",
      fields: [
        inputField("id", "服務費編號", "text"),
        inputField("date", "發生日", "date"),
        inputField("month", "歸屬月份", "month"),
        inputField("year", "年度", "number"),
        selectField("feeType", "費用類型", state.settings.managementFeeTypes),
        selectField("roomId", "房號", [{ value: "", label: "公共 / 不指定" }, ...roomIds.map((room) => ({ value: room, label: room }))]),
        selectField("tenantId", "租客", [{ value: "", label: "不指定" }, ...state.tenants.map((item) => ({ value: item.id, label: item.name }))]),
        selectField("contractId", "關聯租約", [{ value: "", label: "不指定" }, ...state.leases.map((item) => ({ value: item.id, label: item.id }))]),
        inputField("baseAmount", "計算基礎金額", "number"),
        inputField("rate", "比例（%）", "number"),
        inputField("amount", "實際金額", "number"),
        inputField("vendor", "業者 / 管理者", "text"),
        selectField("status", "狀態", ["待支付", "已支付", "爭議中", "免付"]),
        inputField("note", "備註", "text", true),
      ],
      save(values) {
        const amount = Number(values.amount || 0);
        const record = {
          ...values,
          year: String(values.year || values.month?.slice(0, 4) || selectedYear()),
          baseAmount: Number(values.baseAmount || 0),
          rate: Number(values.rate || 0),
          amount,
        };
        state.managementFees.push(record);
        state.ledgerEntries.push({
          id: `LG-${record.id}`,
          date: record.date,
          month: record.month,
          roomId: record.roomId || "公共",
          item: record.feeType,
          category: record.feeType.includes("代管") ? "代管費" : "代租費",
          income: 0,
          expense: amount,
          note: record.vendor || record.note || "",
        });
      },
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
    incidents: {
      title: "新增事件紀錄",
      fields: [
        inputField("id", "事件編號", "text"),
        inputField("date", "發生日", "date"),
        inputField("month", "歸屬月份", "month"),
        selectField("type", "事件類型", state.settings.incidentTypes),
        selectField("roomId", "房號", [{ value: "", label: "公共 / 不指定" }, ...roomIds.map((room) => ({ value: room, label: room }))]),
        selectField("tenantId", "租客", [{ value: "", label: "不指定" }, ...state.tenants.map((item) => ({ value: item.id, label: item.name }))]),
        selectField("leaseId", "關聯租約", [{ value: "", label: "不指定" }, ...state.leases.map((item) => ({ value: item.id, label: item.id }))]),
        selectField("severity", "嚴重度", ["低", "中", "高"]),
        selectField("status", "處理狀態", ["待處理", "追蹤中", "已結案", "取消"]),
        inputField("action", "處理摘要", "text"),
        inputField("evidenceId", "關聯附件編號", "text"),
        inputField("note", "備註", "text", true),
      ],
      save(values) { state.incidents.push(values); },
    },
    attachments: {
      title: "新增附件紀錄",
      fields: [
        inputField("id", "附件編號", "text"),
        inputField("name", "附件名稱", "text"),
        inputField("type", "附件類型", "text"),
        selectField("module", "關聯模組", ["租約", "臺電", "電表", "維修", "收款", "租客", "事件", "代租代管", "其他"]),
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

function aggregateBy(items, key, valueKey) {
  return Array.from(items.reduce((map, item) => {
    const groupKey = item[key] || "未分類";
    const current = map.get(groupKey) || { key: groupKey, total: 0, count: 0 };
    current.total += Number(item[valueKey] || 0);
    current.count += 1;
    map.set(groupKey, current);
    return map;
  }, new Map()).values());
}

function statusSummaryCard(label, items) {
  return `<article class="summary-card"><span class="overline">${label}</span><strong>${items.length} 件</strong><span class="muted">${items.slice(0, 2).map((item) => `${item.roomId} ${item.type}`).join("、") || "目前無案件"}</span></article>`;
}

function statusTag(status) {
  const tone = ["已收", "生效中", "出租中", "保管中", "已完成", "已支付", "已結案", "低"].includes(status)
    ? "success"
    : ["部分收", "整修中", "處理中", "已派工", "即將到期", "中"].includes(status)
      ? "warning"
      : ["未收", "逾期", "空房", "取消", "高"].includes(status)
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
  const regularExpenses = state.expenses.filter((item) => item.month === state.ui.month && item.payer === "房東").reduce((sum, item) => sum + Number(item.amount), 0);
  const managementFees = state.managementFees.filter((item) => item.month === state.ui.month).reduce((sum, item) => sum + Number(item.amount), 0);
  const expenses = regularExpenses + managementFees;
  const electricDue = sumBy(allocation, "finalDue");
  const electricPaid = sumBy(allocation, "paid");
  const publicExpenses = state.expenses.filter((item) => item.month === state.ui.month && item.scope === "公共").reduce((sum, item) => sum + Number(item.amount), 0);
  const expiringLeases = state.leases.filter((item) => item.status === "即將到期");
  const reminders = [];

  payments.filter((item) => item.amountPaid < item.amountDue).forEach((item) => reminders.push(`${item.roomId} 房租尚差 ${money.format(item.amountDue - item.amountPaid)}`));
  allocation.filter((item) => item.paid < item.finalDue).forEach((item) => reminders.push(`${item.roomId} 電費尚差 ${money.format(item.finalDue - item.paid)}`));
  state.maintenance.filter((item) => ["待處理", "已派工", "處理中"].includes(item.status)).forEach((item) => reminders.push(`${item.roomId} ${item.type} 維修目前為 ${item.status}`));
  state.incidents.filter((item) => !["已結案", "取消"].includes(item.status)).forEach((item) => reminders.push(`${item.roomId || "公共"} ${item.type} 事件目前為 ${item.status}`));

  return {
    rentDue,
    rentPaid,
    electricDue,
    electricPaid,
    expenses,
    regularExpenses,
    managementFees,
    netProfit: rentPaid - expenses,
    publicExpenseShare: publicExpenses / 4,
    reminders: reminders.slice(0, 8),
    expiringLeases,
    unpaidRentRooms: payments.filter((item) => item.amountPaid < item.amountDue),
    unpaidElectricRooms: allocation.filter((item) => item.paid < item.finalDue),
  };
}

function workspaceStatusText() {
  const workspace = getWorkspace();
  if (!workspace.sheetId) return "尚未指定雲端主表";
  if (!googleProfile?.email) return "尚未登入 Google，資料仍以本機狀態為主";
  if (workspace.ownerEmail && workspace.ownerEmail !== googleProfile.email) return "目前登入帳號不是資料擁有者，需確認共享權限";
  return "Google 授權完成後，變更會同步回寫主表";
}

function topExpenseLabel(month) {
  const top = aggregateLedgerCategories(state.ledgerEntries.filter((item) => item.month === month && Number(item.expense) > 0))[0];
  return top ? `${top.category} ${money.format(top.expense)}` : "目前沒有主要支出";
}

function operatingScore(summary) {
  const occupancy = state.rooms.length ? Math.round((state.tenants.length / state.rooms.length) * 100) : 0;
  const rentRate = summary.rentDue ? Math.round((summary.rentPaid / summary.rentDue) * 100) : 100;
  const relationRate = relationScore();
  const riskPenalty = Math.min(30, summary.reminders.length * 6);
  return Math.max(52, Math.min(99, Math.round((occupancy * 0.28) + (rentRate * 0.38) + (relationRate * 0.24) + 10 - riskPenalty)));
}

function scoreLabel(score) {
  if (score >= 90) return "穩定營運";
  if (score >= 80) return "狀態良好";
  if (score >= 70) return "可持續追蹤";
  return "需優先整理";
}

function scoreTone(score) {
  if (score >= 90) return "tone-good";
  if (score >= 80) return "tone-bright";
  if (score >= 70) return "tone-warm";
  return "tone-alert";
}

function buildScoreRows(summary) {
  const occupancy = state.rooms.length ? Math.round((state.tenants.length / state.rooms.length) * 100) : 0;
  const rentRate = summary.rentDue ? Math.round((summary.rentPaid / summary.rentDue) * 100) : 100;
  const relationRate = relationScore();
  const issueRate = Math.max(35, 100 - Math.min(65, summary.reminders.length * 12));
  return [
    { label: "收租完成率", value: rentRate },
    { label: "滿租程度", value: occupancy },
    { label: "資料完整度", value: relationRate },
    { label: "待辦風險", value: issueRate },
  ];
}

function buildMonthlyTrend(year) {
  const monthly = Array.from({ length: 12 }, (_, index) => {
    const month = `${year}-${String(index + 1).padStart(2, "0")}`;
    const entries = state.ledgerEntries.filter((item) => item.month === month);
    const feeEntries = state.managementFees.filter((item) => item.month === month || item.year === year && month === `${year}-04` && item.month === "2026-04");
    return {
      month: month.slice(5),
      income: sumBy(entries, "income"),
      expense: sumBy(entries, "expense") + sumBy(feeEntries, "amount"),
    };
  });
  const maxIncome = Math.max(...monthly.map((item) => item.income), 1);
  const maxExpense = Math.max(...monthly.map((item) => item.expense), 1);
  return monthly.map((item) => ({
    ...item,
    incomeWidth: Math.max(item.income ? 14 : 0, Math.round((item.income / maxIncome) * 100)),
    expenseWidth: Math.max(item.expense ? 14 : 0, Math.round((item.expense / maxExpense) * 100)),
  }));
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

function selectedYear() {
  return (state.ui.month || defaultMonth).slice(0, 4);
}

function recordCountByMonth(month) {
  return [
    state.rentPayments,
    state.expenses,
    state.managementFees,
    state.ledgerEntries,
    state.incidents,
  ].reduce((count, list) => count + list.filter((item) => item.month === month).length, 0);
}

function buildRelationRows() {
  const rows = [
    ...state.leases.map((item) => ({
      label: `租約 ${item.id}`,
      month: item.startDate?.slice(0, 7) || "",
      roomId: item.roomId,
      tenant: tenantName(item.tenantId),
      recordId: item.id,
    })),
    ...state.managementFees.map((item) => ({
      label: `服務費 ${item.id}`,
      month: item.month,
      roomId: item.roomId,
      tenant: tenantName(item.tenantId),
      recordId: item.id,
    })),
    ...state.incidents.map((item) => ({
      label: `事件 ${item.id}`,
      month: item.month,
      roomId: item.roomId,
      tenant: tenantName(item.tenantId),
      recordId: item.id,
    })),
  ];

  return rows.slice(-8).reverse().map((item) => ({
    ...item,
    attachments: state.attachments.filter((file) => file.recordId === item.recordId).length,
  }));
}

function relationScore() {
  const checks = [
    ...state.rentPayments.map((item) => [item.id, item.month, item.roomId, item.tenantId]),
    ...state.expenses.map((item) => [item.id, item.month, item.type, item.scope]),
    ...state.managementFees.map((item) => [item.id, item.month, item.year, item.feeType, item.roomId]),
    ...state.incidents.map((item) => [item.id, item.month, item.date, item.type, item.status]),
    ...state.attachments.map((item) => [item.id, item.module, item.recordId, item.url]),
  ];
  if (!checks.length) return 100;
  const complete = checks.filter((fields) => fields.every((value) => String(value || "").trim())).length;
  return Math.round((complete / checks.length) * 100);
}

function chartColor(index) {
  return ["#ff7a59", "#ff4fa3", "#6f7cff", "#24c6a5", "#ffc857", "#32a8ff", "#7bd88f", "#f26d85"][index % 8];
}

function pieChart(items) {
  const filtered = items.filter((item) => item.value > 0);
  if (!filtered.length) return `<div class="empty-chart">目前沒有可繪製的資料</div>`;
  const total = sumBy(filtered, "value");
  let offset = 0;
  const segments = filtered.map((item) => {
    const size = (item.value / total) * 100;
    const segment = `${item.color} ${offset}% ${offset + size}%`;
    offset += size;
    return segment;
  }).join(", ");
  return `
    <div class="chart-layout">
      <div class="pie-chart" style="background: conic-gradient(${segments});"></div>
      <div class="chart-legend">
        ${filtered.map((item) => `<span><i style="background:${item.color}"></i>${item.label} ${money.format(item.value)}</span>`).join("")}
      </div>
    </div>
  `;
}

function radarChart(items) {
  const center = 100;
  const maxRadius = 74;
  const points = items.map((item, index) => {
    const angle = (-90 + (360 / items.length) * index) * Math.PI / 180;
    const radius = maxRadius * (Math.max(0, Math.min(100, item.value)) / 100);
    return `${center + Math.cos(angle) * radius},${center + Math.sin(angle) * radius}`;
  }).join(" ");
  const axes = items.map((item, index) => {
    const angle = (-90 + (360 / items.length) * index) * Math.PI / 180;
    const x = center + Math.cos(angle) * maxRadius;
    const y = center + Math.sin(angle) * maxRadius;
    const labelX = center + Math.cos(angle) * 92;
    const labelY = center + Math.sin(angle) * 92;
    return `<line x1="${center}" y1="${center}" x2="${x}" y2="${y}" /><text x="${labelX}" y="${labelY}">${item.label}</text>`;
  }).join("");
  return `
    <div class="radar-wrap">
      <svg class="radar-chart" viewBox="0 0 200 200" role="img" aria-label="營運雷達圖">
        <polygon class="radar-grid" points="100,26 170,76 143,158 57,158 30,76"></polygon>
        <polygon class="radar-grid inner" points="100,55 142,85 126,134 74,134 58,85"></polygon>
        <g class="radar-axis">${axes}</g>
        <polygon class="radar-area" points="${points}"></polygon>
      </svg>
      <div class="chart-legend">${items.map((item) => `<span><i style="background:${chartColor(item.value)}"></i>${item.label} ${item.value}</span>`).join("")}</div>
    </div>
  `;
}

function buildRadarMetrics(summary) {
  const rentRate = summary.rentDue ? Math.round((summary.rentPaid / summary.rentDue) * 100) : 100;
  const electricRate = summary.electricDue ? Math.round((summary.electricPaid / summary.electricDue) * 100) : 100;
  const incidentPenalty = Math.min(60, state.incidents.filter((item) => !["已結案", "取消"].includes(item.status)).length * 15);
  const attachmentRate = relationScore();
  const profitRate = summary.rentPaid ? Math.max(0, Math.min(100, Math.round((summary.netProfit / summary.rentPaid) * 100))) : 0;
  return [
    { label: "收租", value: rentRate },
    { label: "電費", value: electricRate },
    { label: "淨利", value: profitRate },
    { label: "關聯", value: attachmentRate },
    { label: "風險", value: 100 - incidentPenalty },
  ];
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
