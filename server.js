require("dotenv").config();

const express = require("express");
const line = require("@line/bot-sdk");
const { google } = require("googleapis");

const app = express();

const config = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

const client = new line.messagingApi.MessagingApiClient({
  channelAccessToken: config.channelAccessToken,
});

const DEFAULT_SHEET_ID = process.env.GOOGLE_SHEET_ID;
const MASTER_SHEET_ID = process.env.MASTER_SHEET_ID;
const ACCOUNT_MASTER_SHEET_ID = process.env.ACCOUNT_MASTER_SHEET_ID;
// 測試版強制隔離：啟動時必須明確啟用，且不得使用正式多人主檔。
if (process.env.TEST_MODE !== "1" || MASTER_SHEET_ID || process.env.CENTRAL_SHEET_ID) {
  throw new Error("測試版安全鎖：需設定 TEST_MODE=1，且不可設定 MASTER_SHEET_ID 或 CENTRAL_SHEET_ID");
}
if (!DEFAULT_SHEET_ID) throw new Error("測試版缺少 GOOGLE_SHEET_ID");

// ===== 場別 =====
const FARM_ALIASES = {
  "測試": "測試場",
  "測試場": "測試場",
  "東平": "東平場",
  "東平場": "東平場",
  "草湖": "草湖場",
  "草湖場": "草湖場",
  "仁愛": "仁愛場",
  "仁愛場": "仁愛場",
  "東勢": "東勢場",
  "東勢場": "東勢場",
  "埤北": "埤北場",
  "埤北場": "埤北場",
  "賜福": "賜福場",
  "賜福場": "賜福場",
  "永興": "永興場",
  "永興場": "永興場",
  "秉夆": "秉夆場",
  "秉夆場": "秉夆場",
  "後寮": "後寮場",
  "後寮場": "後寮場",
  "鎮平": "鎮平場",
  "鎮平場": "鎮平場",
  "東陽": "東陽場",
  "東陽場": "東陽場",
  "潭墘": "潭墘場",
  "潭墘場": "潭墘場",
  "龍潭": "龍潭場",
  "龍潭場": "龍潭場",
  "泰順": "泰順場",
  "泰順場": "泰順場",
  "清屎": "清屎部門",
  "清屎部門": "清屎部門",

  "班神": "班神公司",
  "班神公司": "班神公司",
  "班神場": "班神公司",

  "昊陽": "昊陽公司",
  "昊陽公司": "昊陽公司",
  "昊陽場": "昊陽公司",

  "共用": "共用",
  "共同": "共用",
  "全部": "全部",
};

// ===== 原 Excel「資料」分頁 =====
const ACCOUNT_ITEMS = [
  ["購入中雞、二春雞", "511101", "銷貨成本"],
  ["蛋紙-大浪-小浪", "511101", "銷貨成本"],
  ["勞務支出-清屎人力派遣", "561199", "勞務成本-其他"],
  ["勞務支出-防疫隊", "561199", "勞務成本-其他"],
  ["疫苗", "511101", "銷貨成本"],
  ["藥品", "511101", "銷貨成本"],
  ["營養品", "511101", "銷貨成本"],
  ["飼料費", "511101", "銷貨成本"],
  ["蚵殼粉", "511101", "銷貨成本"],
  ["間接人工", "515101", "勞務成本-其他"],
  ["機器設備", "142101", "機器設備－成本"],
  ["薪資支出-顧問費", "611101", "營業費用"],
  ["薪資支出-廠長", "611101", "營業費用"],
  ["薪資支出-員工1", "611101", "營業費用"],
  ["薪資支出-員工2", "611101", "營業費用"],
  ["薪資支出-員工3", "611101", "營業費用"],
  ["薪資支出-員工4", "611101", "營業費用"],
  ["薪資支出-加班費", "611101", "營業費用"],
  ["租金支出-雞場", "611201", "營業費用"],
  ["水費-雞場", "611901", "營業費用"],
  ["電費-雞場", "611901", "營業費用"],
  ["瓦斯費-雞場", "611901", "營業費用"],
  ["設備修繕費", "611701", "營業費用"],
  ["建物修繕費", "611702", "營業費用"],
  ["其他修繕費", "611799", "營業費用"],
  ["保險費", "612004", "營業費用"],
  ["雜項費用", "613402", "營業費用"],
  ["租金支出-宿舍", "611209", "營業費用"],
  ["宿舍-水費", "611901", "營業費用"],
  ["宿舍-電費", "611901", "營業費用"],
  ["宿舍-瓦斯費", "611901", "營業費用"],
  ["交通費-油資(車號)", "611401", "營業費用"],
  ["車輛保養維修費", "611799", "營業費用"],
  ["伙食費", "6128", "營業費用"],
  ["勞保費", "612001", "營業費用"],
  ["健保費", "612002", "營業費用"],
  ["生化-廢棄物清運費", "613403", "營業費用"],
  ["獸醫師年度合約", "613301", "營業費用"],
  ["檢驗費", "516904", "營業費用"],
];

const KEYWORD_RULES = [
  { keys: ["清屎"], item: "勞務支出-清屎人力派遣" },
  { keys: ["防疫隊", "防疫"], item: "勞務支出-防疫隊" },
  { keys: ["中雞", "二春雞", "購入雞"], item: "購入中雞、二春雞" },
  { keys: ["蛋紙", "大浪", "小浪"], item: "蛋紙-大浪-小浪" },
  { keys: ["疫苗"], item: "疫苗" },
  { keys: ["營養品", "營養"], item: "營養品" },
  { keys: ["飼料"], item: "飼料費" },
  { keys: ["蚵殼"], item: "蚵殼粉" },
  { keys: ["藥品", "藥"], item: "藥品" },

  { keys: ["顧問"], item: "薪資支出-顧問費" },
  { keys: ["廠長薪資", "廠長薪水"], item: "薪資支出-廠長" },
  { keys: ["加班"], item: "薪資支出-加班費" },

  { keys: ["雞場租金", "雞場房租", "房租"], item: "租金支出-雞場" },
  { keys: ["宿舍租金", "宿舍房租"], item: "租金支出-宿舍" },

  { keys: ["宿舍水費"], item: "宿舍-水費" },
  { keys: ["宿舍電費"], item: "宿舍-電費" },
  { keys: ["宿舍瓦斯"], item: "宿舍-瓦斯費" },
  { keys: ["水費"], item: "水費-雞場" },
  { keys: ["電費"], item: "電費-雞場" },
  { keys: ["瓦斯"], item: "瓦斯費-雞場" },

  { keys: ["水泥", "屋頂", "牆壁", "建物", "鐵皮"], item: "建物修繕費" },
  { keys: ["電線", "燈座", "葉片", "輪胎", "手推車", "馬達維修", "設備維修", "機器維修", "山貓維修", "修理"], item: "設備修繕費" },
  { keys: ["新機器", "新設備", "機器設備"], item: "機器設備" },

  { keys: ["員工油資", "汽油", "柴油", "加油", "油資"], item: "交通費-油資(車號)" },
  { keys: ["車輛保養", "汽車保養", "車輛維修", "汽車維修"], item: "車輛保養維修費" },

  { keys: ["礦泉水", "飲料", "便當", "餐費", "沙拉油", "伙食"], item: "伙食費" },
  { keys: ["勞保"], item: "勞保費" },
  { keys: ["健保"], item: "健保費" },
  { keys: ["清運", "廢棄物"], item: "生化-廢棄物清運費" },
  { keys: ["獸醫"], item: "獸醫師年度合約" },
  { keys: ["檢驗"], item: "檢驗費" },
  { keys: ["保險"], item: "保險費" },

  // 依原支出表常用項目
  { keys: ["電風扇", "潤滑油", "衛生紙", "抹布", "抺布", "漂白水", "洗碗精", "噴霧器油"], item: "雜項費用" },
];

function getGoogleAuth() {
  const privateKey = (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n");

  return new google.auth.JWT({
    email: process.env.GOOGLE_CLIENT_EMAIL,
    key: privateKey,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
}

async function getSheets() {
  return google.sheets({ version: "v4", auth: getGoogleAuth() });
}

function taipeiNow() {
  const parts = new Intl.DateTimeFormat("zh-TW", {
    timeZone: "Asia/Taipei",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  }).formatToParts(new Date());

  const get = (type) => parts.find(p => p.type === type)?.value || "";

  return {
    year: get("year"),
    month: get("month"),
    day: get("day"),
    dateText: `${get("year")}/${get("month")}/${get("day")}`,
    monthSheet: `${get("year")}${get("month")}月`,
  };
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function makeDateInfo(year, month, day) {
  const d = new Date(year, month - 1, day);
  if (d.getFullYear() !== year || d.getMonth() + 1 !== month || d.getDate() !== day) {
    return null;
  }
  return {
    year, month, day,
    dateText: `${year}/${pad2(month)}/${pad2(day)}`,
    monthSheet: `${year}${pad2(month)}月`,
  };
}

function parseOptionalDatePrefix(text) {
  const clean = String(text || "").trim().replace(/\s+/g, " ");
  const now = taipeiNow();

  let m = clean.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\s+(.+)$/);
  if (m) {
    const dateInfo = makeDateInfo(Number(m[1]), Number(m[2]), Number(m[3]));
    if (!dateInfo) return { error: "日期格式不正確，例如：2026/9/28 東平 支出 500 電風扇" };
    return { dateInfo, body: m[4].trim(), isCustomDate: true };
  }

  m = clean.match(/^(\d{1,2})[\/\-](\d{1,2})\s+(.+)$/);
  if (m) {
    const dateInfo = makeDateInfo(Number(now.year), Number(m[1]), Number(m[2]));
    if (!dateInfo) return { error: "日期格式不正確，例如：9/28 東平 支出 500 電風扇" };
    return { dateInfo, body: m[3].trim(), isCustomDate: true };
  }

  return {
    dateInfo: {
      year: Number(now.year),
      month: Number(now.month),
      day: Number(now.day),
      dateText: now.dateText,
      monthSheet: now.monthSheet,
    },
    body: clean,
    isCustomDate: false,
  };
}

function parseAmount(v) {
  return Number(String(v || "").replace(/,/g, "").trim());
}

function farmFromToken(token) {
  return FARM_ALIASES[String(token || "").trim()] || null;
}


let ACCOUNT_MASTER_CACHE = {
  loadedAt: 0,
  rows: null,
};

async function getAccountMasterRows() {
  const now = Date.now();

  // 5 分鐘快取，避免每一筆都重新讀 Google Sheet
  if (
    ACCOUNT_MASTER_CACHE.rows &&
    now - ACCOUNT_MASTER_CACHE.loadedAt < 5 * 60 * 1000
  ) {
    return ACCOUNT_MASTER_CACHE.rows;
  }

  // 尚未設定中央科目主檔時，暫時沿用程式內建資料
  if (!ACCOUNT_MASTER_SHEET_ID) {
    const fallback = ACCOUNT_ITEMS.map(r => ({
      item: String(r[0]),
      code: String(r[1]),
      className: String(r[2]),
      keywords: [],
    }));

    ACCOUNT_MASTER_CACHE = {
      loadedAt: now,
      rows: fallback,
    };

    return fallback;
  }

  const sheets = await getSheets();

  const meta = await sheets.spreadsheets.get({
    spreadsheetId: ACCOUNT_MASTER_SHEET_ID,
    fields: "sheets.properties(title,index)",
  });

  const tabs = (meta.data.sheets || [])
    .map(s => s.properties)
    .sort((a, b) => (a.index || 0) - (b.index || 0));

  // 優先讀「資料」分頁；沒有就用第一個分頁
  const targetTab = tabs.find(t => t.title === "資料") || tabs[0];

  if (!targetTab?.title) {
    throw new Error("中央科目主檔沒有可讀取的分頁");
  }

  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: ACCOUNT_MASTER_SHEET_ID,
    range: `'${targetTab.title}'!A2:D`,
    valueRenderOption: "FORMATTED_VALUE",
  });

  const rows = (result.data.values || [])
    .map(r => ({
      item: String(r[0] || "").trim(),
      code: String(r[1] || "").trim(),
      className: String(r[2] || "").trim(),
      keywords: String(r[3] || "")
        .split(/[,，、]/)
        .map(x => x.trim())
        .filter(Boolean),
    }))
    .filter(r => r.item);

  if (!rows.length) {
    throw new Error("中央科目主檔沒有科目資料");
  }

  ACCOUNT_MASTER_CACHE = {
    loadedAt: now,
    rows,
  };

  return rows;
}

async function accountByItem(item) {
  const rows = await getAccountMasterRows();
  return rows.find(r => r.item === item) || null;
}

function normalizeMatchText(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, "")
    .toLowerCase();
}

async function autoChooseItem(description) {
  const rawText = String(description || "").trim();
  const text = normalizeMatchText(rawText);
  const rows = await getAccountMasterRows();

  // 1) 先看 A 欄「品項」
  // 完整命中優先
  let matched = rows.find(r =>
    normalizeMatchText(r.item) === text
  );
  if (matched) return matched;

  // A 欄品項名稱包含於描述中，較長的品項優先
  matched = rows
    .filter(r => {
      const itemText = normalizeMatchText(r.item);
      return itemText && text.includes(itemText);
    })
    .sort((a, b) =>
      normalizeMatchText(b.item).length - normalizeMatchText(a.item).length
    )[0];

  if (matched) return matched;

  // 2) 再看 D 欄「辨識關鍵字」
  // 命中較長的關鍵字優先，避免「租金」蓋過「宿舍租金」
  const keywordMatches = [];

  for (const row of rows) {
    for (const keyword of row.keywords || []) {
      const k = normalizeMatchText(keyword);
      if (k && text.includes(k)) {
        keywordMatches.push({
          row,
          keywordLength: k.length,
        });
      }
    }
  }

  if (keywordMatches.length) {
    keywordMatches.sort((a, b) => b.keywordLength - a.keywordLength);
    return keywordMatches[0].row;
  }

  // 3) A、D 都找不到才歸「雜項費用」
  return rows.find(r => r.item === "雜項費用") || rows[0] || null;
}

async function getMasterSheetTitle() {
  if (!MASTER_SHEET_ID) return null;

  const sheets = await getSheets();
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: MASTER_SHEET_ID,
    fields: "sheets.properties(title,index)",
  });

  const tabs = (meta.data.sheets || [])
    .map(s => s.properties)
    .sort((a, b) => (a.index || 0) - (b.index || 0));

  return tabs[0]?.title || null;
}

async function resolveUserSheet(event) {
  const userId = event.source?.userId || "";

  // 如果尚未設定主檔，暫時沿用原本單人版 Sheet。
  if (!MASTER_SHEET_ID) {
    if (!DEFAULT_SHEET_ID) {
      throw new Error("缺少 MASTER_SHEET_ID 與 GOOGLE_SHEET_ID");
    }

    return {
      sheetId: DEFAULT_SHEET_ID,
      name: await getProfileName(event),
      source: "default",
    };
  }

  if (!userId) {
    throw new Error("無法取得 LINE User ID");
  }

  const masterTitle = await getMasterSheetTitle();
  if (!masterTitle) {
    throw new Error("使用者主檔沒有可讀取的分頁");
  }

  const sheets = await getSheets();
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: MASTER_SHEET_ID,
    range: `'${masterTitle}'!A2:F`,
    valueRenderOption: "FORMATTED_VALUE",
  });

  const rows = result.data.values || [];

  const matched = rows.find(r =>
    String(r[2] || "").trim() === userId
  );

  if (!matched) {
    return {
      error: "❌ 你的帳號尚未登記在使用者主檔。\n請先聯絡管理者完成綁定。",
    };
  }

  const name = String(matched[0] || "").trim() || "未命名";
  const sheetId = String(matched[3] || "").trim();
  const enabled = String(matched[4] || "").trim().toUpperCase();

  if (!["Y", "YES", "TRUE", "1", "啟用"].includes(enabled)) {
    return {
      error: `❌ ${name} 的帳號目前尚未啟用。`,
    };
  }

  if (!sheetId) {
    return {
      error: `❌ ${name} 尚未設定 Google Sheet ID。`,
    };
  }

  return {
    sheetId,
    name,
    source: "master",
  };
}

async function ensureSheet(sheetId, sheetName, headers) {
  const sheets = await getSheets();

  const meta = await sheets.spreadsheets.get({
    spreadsheetId: sheetId,
    fields: "sheets.properties(sheetId,title)",
  });

  const exists = (meta.data.sheets || []).some(s => s.properties.title === sheetName);

  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: {
        requests: [{
          addSheet: {
            properties: { title: sheetName },
          },
        }],
      },
    });
  }

  if (headers?.length) {
    const current = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: `'${sheetName}'!A1:${columnLetter(headers.length)}1`,
    });

    const firstRow = current.data.values?.[0] || [];

    if (firstRow.join("|") !== headers.join("|")) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: sheetId,
        range: `'${sheetName}'!A1:${columnLetter(headers.length)}1`,
        valueInputOption: "RAW",
        requestBody: { values: [headers] },
      });
    }
  }
}

function columnLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

async function ensureBaseStructure(sheetId) {
  // 科目統一由中央科目主檔提供。
  // 每位使用者自己的帳本不再需要維護「資料」分頁。
  await getAccountMasterRows();
}


async function migrateMonthSheetColumns(sheetId, monthSheet) {
  const sheets = await getSheets();

  const desiredHeaders = [
    "場別",
    "收支類型",
    "付款來源",
    "科目代號",
    "科目分類",
    "日期",
    "品項",
    "用途說明",
    "廠商名稱",
    "數量",
    "單價",
    "金額",
    "發票或憑證",
    "代墊對象",
    "填表人",
  ];

  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `'${monthSheet}'!A1:O`,
    valueRenderOption: "FORMATTED_VALUE",
  });

  const values = result.data.values || [];
  const currentHeaders = values[0] || [];

  if (currentHeaders.join("|") === desiredHeaders.join("|")) {
    return;
  }

  const oldHeaders = [
    "場別",
    "科目代號",
    "科目分類",
    "日期",
    "品項",
    "用途說明",
    "廠商名稱",
    "數量",
    "單價",
    "金額",
    "發票或憑證",
    "付款來源",
    "收支類型",
    "代墊對象",
    "填表人",
  ];

  // 只有辨識到舊版欄位順序時才自動搬移，避免誤改其他格式。
  const looksLikeOldLayout =
    currentHeaders.slice(0, oldHeaders.length).join("|") === oldHeaders.join("|");

  if (!looksLikeOldLayout) {
    return;
  }

  const rows = values.slice(1).map(r => [
    r[0] || "",   // 場別
    r[12] || "支出", // 收支類型
    r[11] || "",  // 付款來源
    r[1] || "",   // 科目代號
    r[2] || "",   // 科目分類
    r[3] || "",   // 日期
    r[4] || "",   // 品項
    r[5] || "",   // 用途說明
    r[6] || "",   // 廠商名稱
    r[7] || "",   // 數量
    r[8] || "",   // 單價
    r[9] || "",   // 金額
    r[10] || "",  // 發票或憑證
    r[13] || "",  // 代墊對象
    r[14] || "",  // 填表人
  ]);

  await sheets.spreadsheets.values.clear({
    spreadsheetId: sheetId,
    range: `'${monthSheet}'!A:O`,
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `'${monthSheet}'!A1:O${Math.max(1, rows.length + 1)}`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [desiredHeaders, ...rows],
    },
  });
}

async function ensureMonthSheet(sheetId, monthSheet) {
  await migrateMonthSheetColumns(sheetId, monthSheet);
  await ensureSheet(sheetId, monthSheet, [
    "場別",
    "收支類型",
    "付款來源",
    "科目代號",
    "科目分類",
    "日期",
    "品項",
    "用途說明",
    "廠商名稱",
    "數量",
    "單價",
    "金額",
    "發票或憑證",
    "代墊對象",
    "填表人",
  ]);
}

async function parseExpenseCommand(text) {
  const clean = String(text || "").trim().replace(/\s+/g, " ");
  const parts = clean.split(" ");

  const farm = farmFromToken(parts[0]);
  let paymentSource = "雞場帳戶";
  let actionIndex = 1;

  if (!farm) return null;

  if (farm === "全部") {
    return { error: "記帳時不能使用「全部」，請指定實際場別。" };
  }

  // 零用金：草湖 零 支出 500 ...
  if (parts[1] === "零") {
    paymentSource = "零用金";
    actionIndex = 2;
  }

  const txType = String(parts[actionIndex] || "").trim();

  // 正式簡化版只保留「支出 / 收入」
  if (!["支出", "收入"].includes(txType)) {
    return null;
  }

  const rest = parts.slice(actionIndex + 1).join(" ");

  let qty = 1;
  let unitPrice;
  let amount;
  let description = "";

  // 支援 2x600
  let m = rest.match(/^(\d+(?:\.\d+)?)\s*[xX×*]\s*([\d,]+(?:\.\d+)?)\s+(.+)$/);

  if (m) {
    qty = Number(m[1]);
    unitPrice = parseAmount(m[2]);
    amount = qty * unitPrice;
    description = m[3].trim();
  } else {
    m = rest.match(/^([\d,]+(?:\.\d+)?)(?:\s+(.+))?$/);

    if (!m) {
      return { error: `格式例如：草湖 ${txType} 500 電風扇 / 三豐` };
    }

    amount = parseAmount(m[1]);
    unitPrice = amount;
    description = String(m[2] || "").trim();
  }

  if (!Number.isFinite(amount) || amount <= 0) {
    return { error: "金額格式不正確" };
  }

  if (!description) {
    return { error: "請輸入用途或品項。" };
  }

  // 廠商用 / 分隔
  let vendor = "";
  const splitMatch = description.match(/^(.+?)\s*[\/｜|]\s*(.+)$/);

  if (splitMatch) {
    description = splitMatch[1].trim();
    vendor = splitMatch[2].trim();
  } else {
    const vendorMatch = description.match(/\s+廠商[:：]\s*(.+)$/);
    if (vendorMatch) {
      vendor = vendorMatch[1].trim();
      description = description.slice(0, vendorMatch.index).trim();
    }
  }

  let selected = null;
  let purpose = description;

  const accountRows = await getAccountMasterRows();
  const itemNames = accountRows
    .map(r => r.item)
    .sort((a, b) => b.length - a.length);

  for (const item of itemNames) {
    if (description === item || description.startsWith(item + " ")) {
      selected = await accountByItem(item);
      purpose = description.slice(item.length).trim() || item;
      break;
    }
  }

  if (!selected) {
    selected = await autoChooseItem(description);
  }

  if (!selected) {
    return { error: "中央科目主檔找不到可使用的科目，請先檢查科目主檔。" };
  }

  return {
    farm,
    qty,
    unitPrice,
    amount,
    accountItem: selected.item,
    code: selected.code,
    className: selected.className,
    description: purpose,
    vendor,
    paymentSource,
    txType,
    advanceTarget: "",
  };
}

async function writeExpense(sheetId, monthSheet, expense, dateText, userName, receiptMarker = "") {
  const sheets = await getSheets();

  const appendResult = await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId,
    range: `'${monthSheet}'!A:O`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [[
        expense.farm,
        expense.txType || "支出",
        expense.paymentSource || "",
        expense.code,
        expense.className,
        dateText,
        expense.accountItem,
        expense.description,
        expense.vendor,
        expense.qty,
        expense.unitPrice,
        expense.txType === "收入" ? -Math.abs(expense.amount) : Math.abs(expense.amount),
        receiptMarker,
        expense.advanceTarget || "",
        userName || "",
      ]],
    },
  });

  // 顏色只是方便查看，不參與任何加減計算
  try {
    const updatedRange =
      appendResult.data?.updates?.updatedRange ||
      appendResult.data?.tableRange ||
      "";

    const rowMatch = String(updatedRange).match(/![A-Z]+(\d+):[A-Z]+(\d+)$/);

    if (rowMatch) {
      const rowNumber = Number(rowMatch[1]);

      const meta = await sheets.spreadsheets.get({
        spreadsheetId: sheetId,
        fields: "sheets.properties(sheetId,title)",
      });

      const tab = (meta.data.sheets || []).find(
        s => s.properties.title === monthSheet
      );

      if (tab?.properties?.sheetId != null) {
        const colorMap = {
          "支出": { red: 1.0, green: 0.92, blue: 0.92 },
          "收入": { red: 0.90, green: 0.98, blue: 0.90 },
        };

        const backgroundColor =
          colorMap[expense.txType] || { red: 1, green: 1, blue: 1 };

        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: sheetId,
          requestBody: {
            requests: [{
              repeatCell: {
                range: {
                  sheetId: tab.properties.sheetId,
                  startRowIndex: rowNumber - 1,
                  endRowIndex: rowNumber,
                  startColumnIndex: 0,
                  endColumnIndex: 15,
                },
                cell: {
                  userEnteredFormat: {
                    backgroundColor,
                  },
                },
                fields: "userEnteredFormat.backgroundColor",
              },
            }],
          },
        });
      }
    }
  } catch (formatErr) {
    console.warn("套用列顏色失敗，不影響記帳：", formatErr.message);
  }
}

function headerIndex(headers, names) {
  for (const name of names) {
    const idx = headers.findIndex(h => String(h || "").trim() === name);
    if (idx >= 0) return idx;
  }
  return -1;
}

async function readMonthRows(sheetId, monthSheet) {
  const sheets = await getSheets();

  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `'${monthSheet}'!A1:Z`,
    valueRenderOption: "FORMATTED_VALUE",
  });

  const values = result.data.values || [];
  if (!values.length) return [];

  const headers = values[0] || [];

  // 不再假設固定欄位位置，直接依標題找欄位。
  const iFarm = headerIndex(headers, ["場別"]);
  const iType = headerIndex(headers, ["收支類型"]);
  const iPayment = headerIndex(headers, ["付款來源"]);
  const iCode = headerIndex(headers, ["科目代號"]);
  const iClass = headerIndex(headers, ["科目分類"]);
  const iDate = headerIndex(headers, ["日期"]);
  const iItem = headerIndex(headers, ["品項"]);
  const iDesc = headerIndex(headers, ["用途說明"]);
  const iVendor = headerIndex(headers, ["廠商名稱"]);
  const iQty = headerIndex(headers, ["數量"]);
  const iUnit = headerIndex(headers, ["單價"]);
  const iAmount = headerIndex(headers, ["金額"]);
  const iReceipt = headerIndex(headers, ["發票或憑證"]);
  const iRecorder = headerIndex(headers, ["填表人"]);

  const cell = (r, idx) => idx >= 0 ? (r[idx] || "") : "";
  const num = v => Number(String(v || "0").replace(/,/g, "")) || 0;

  return values.slice(1).map(r => ({
    farm: cell(r, iFarm),
    txType: cell(r, iType) || "支出",
    paymentSource: cell(r, iPayment),
    code: cell(r, iCode),
    className: cell(r, iClass),
    date: cell(r, iDate),
    item: cell(r, iItem),
    description: cell(r, iDesc),
    vendor: cell(r, iVendor),
    qty: num(cell(r, iQty)),
    unitPrice: num(cell(r, iUnit)),
    amount: num(cell(r, iAmount)),
    receipt: cell(r, iReceipt),
    recorder: cell(r, iRecorder),
  })).filter(r => r.date || r.item || r.amount);
}

// 將舊資料的「金額」欄正負號整理成：支出 +、收入 -。
// 只改金額欄，不刪除任何歷史資料。
async function normalizeMonthAmountSigns(sheetId, monthSheet) {
  const sheets = await getSheets();

  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `'${monthSheet}'!A1:Z`,
    valueRenderOption: "FORMATTED_VALUE",
  });

  const values = result.data.values || [];
  if (values.length <= 1) return;

  const headers = values[0] || [];
  const iType = headerIndex(headers, ["收支類型"]);
  const iAmount = headerIndex(headers, ["金額"]);

  if (iAmount < 0) return;

  let changed = false;
  const amounts = values.slice(1).map(r => {
    const raw = String(r[iAmount] || "").replace(/,/g, "").trim();
    if (!raw) return [""];

    const n = Number(raw);
    if (!Number.isFinite(n)) return [r[iAmount] || ""];

    const txType = iType >= 0 ? String(r[iType] || "支出").trim() : "支出";
    let normalized = n;

    if (txType === "收入") normalized = -Math.abs(n);
    else if (txType === "支出" || !txType) normalized = Math.abs(n);
    else {
      // 已停用的舊代墊類型不自動改動，避免破壞歷史資料。
      return [r[iAmount] || ""];
    }

    if (normalized !== n) changed = true;
    return [normalized];
  });

  if (!changed) return;

  const col = columnLetter(iAmount + 1);
  await sheets.spreadsheets.values.update({
    spreadsheetId: sheetId,
    range: `'${monthSheet}'!${col}2:${col}${amounts.length + 1}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: amounts },
  });
}

function parseDateText(s) {
  const m = String(s || "").match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (!m) return null;

  return {
    year: Number(m[1]),
    month: Number(m[2]),
    day: Number(m[3]),
  };
}

function parseQuery(text) {
  const clean = String(text || "").trim().replace(/\s+/g, " ");
  const parts = clean.split(" ");
  const now = taipeiNow();

  if (["今天", "本月", "今年"].includes(clean)) {
    return { farm: "全部", mode: clean };
  }

  // 9月
  let m = clean.match(/^(\d{1,2})月$/);
  if (m) {
    const month = Number(m[1]);
    if (month >= 1 && month <= 12) {
      return { farm: "全部", mode: "指定月份", year: now.year, month };
    }
  }

  // 2026/9 或 2026-9
  m = clean.match(/^(\d{4})[\/\-](\d{1,2})$/);
  if (m) {
    const year = Number(m[1]);
    const month = Number(m[2]);
    if (month >= 1 && month <= 12) {
      return { farm: "全部", mode: "指定月份", year, month };
    }
  }

  if (parts.length === 2) {
    const farm = farmFromToken(parts[0]);

    if (farm && ["今天", "本月", "今年"].includes(parts[1])) {
      return { farm, mode: parts[1] };
    }

    // 草湖 9月
    m = parts[1].match(/^(\d{1,2})月$/);
    if (farm && m) {
      const month = Number(m[1]);
      if (month >= 1 && month <= 12) {
        return { farm, mode: "指定月份", year: now.year, month };
      }
    }

    // 東平 2026/9
    m = parts[1].match(/^(\d{4})[\/\-](\d{1,2})$/);
    if (farm && m) {
      const year = Number(m[1]);
      const month = Number(m[2]);
      if (month >= 1 && month <= 12) {
        return { farm, mode: "指定月份", year, month };
      }
    }
  }

  return null;
}

function summaryText(rows, farm, title) {
  const filteredAll = farm === "全部" ? rows : rows.filter(r => r.farm === farm);

  // 目前正式版只統計支出 / 收入。
  // 舊的代墊、代分配資料保留在 Sheet，但不列入目前收支統計。
  const filtered = filteredAll.filter(r =>
    ["支出", "收入"].includes(r.txType || "支出")
  );

  const expenseRows = filtered.filter(r => (r.txType || "支出") === "支出");
  const incomeRows = filtered.filter(r => r.txType === "收入");

  const expenseTotal = expenseRows.reduce((sum, r) => sum + Math.abs(r.amount), 0);
  const incomeTotal = incomeRows.reduce((sum, r) => sum + Math.abs(r.amount), 0);
  const net = expenseTotal - incomeTotal;

  const byItem = {};
  for (const r of expenseRows) {
    if (!r.item) continue;
    byItem[r.item] = (byItem[r.item] || 0) + Math.abs(r.amount);
  }

  const top = Object.entries(byItem)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  const lines = [
    `📒 ${farm === "全部" ? "全部場" : farm}｜${title}`,
    "",
    `💸 支出：＋${expenseTotal.toLocaleString("zh-TW")} 元`,
    `💰 收入：－${incomeTotal.toLocaleString("zh-TW")} 元`,
    `📊 淨支出：${net.toLocaleString("zh-TW")} 元`,
    `🧾 筆數：${filtered.length} 筆`,
  ];

  const ignoredCount = filteredAll.length - filtered.length;
  if (ignoredCount > 0) {
    lines.push(`ℹ️ 舊代墊類資料：${ignoredCount} 筆（目前不計入）`);
  }

  if (top.length) {
    lines.push("", "支出分類：");
    for (const [name, value] of top) {
      lines.push(`・${name}：${value.toLocaleString("zh-TW")} 元`);
    }
  }

  return lines.join("\n");
}

async function querySummary(sheetId, query) {
  const now = taipeiNow();

  if (query.mode === "指定月份") {
    const monthSheet = `${query.year}${pad2(query.month)}月`;

    const sheets = await getSheets();
    const meta = await sheets.spreadsheets.get({
      spreadsheetId: sheetId,
      fields: "sheets.properties.title",
    });

    const exists = (meta.data.sheets || []).some(
      s => s.properties.title === monthSheet
    );

    if (!exists) {
      return [
        `📒 ${query.farm === "全部" ? "全部場" : query.farm}｜${query.year}年${query.month}月支出`,
        "",
        "目前沒有這個月份的資料。",
      ].join("\n");
    }

    await normalizeMonthAmountSigns(sheetId, monthSheet);
    const rows = await readMonthRows(sheetId, monthSheet);
    return summaryText(
      rows,
      query.farm,
      `${query.year}年${query.month}月收支`
    );
  }

  if (query.mode === "今天") {
    await ensureMonthSheet(sheetId, now.monthSheet);
    await normalizeMonthAmountSigns(sheetId, now.monthSheet);
    const rows = await readMonthRows(sheetId, now.monthSheet);

    const todayRows = rows.filter(r => {
      const d = parseDateText(r.date);
      return d &&
        d.year === Number(now.year) &&
        d.month === Number(now.month) &&
        d.day === Number(now.day);
    });

    return summaryText(todayRows, query.farm, "今日收支");
  }

  if (query.mode === "本月") {
    await ensureMonthSheet(sheetId, now.monthSheet);
    await normalizeMonthAmountSigns(sheetId, now.monthSheet);
    const rows = await readMonthRows(sheetId, now.monthSheet);
    return summaryText(rows, query.farm, `${Number(now.month)}月收支`);
  }

  // 今年
  const sheets = await getSheets();
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: sheetId,
    fields: "sheets.properties.title",
  });

  const monthSheets = (meta.data.sheets || [])
    .map(s => s.properties.title)
    .filter(title => new RegExp(`^${now.year}\\d{2}月$`).test(title))
    .sort();

  let rows = [];

  for (const monthSheet of monthSheets) {
    await normalizeMonthAmountSigns(sheetId, monthSheet);
    rows = rows.concat(await readMonthRows(sheetId, monthSheet));
  }

  return summaryText(rows, query.farm, `${now.year}年收支`);
}

async function getProfileName(event) {
  try {
    if (!event.source?.userId) return "未知使用者";
    const profile = await client.getProfile(event.source.userId);
    return profile.displayName || "未知使用者";
  } catch (_) {
    return "未知使用者";
  }
}

function helpText() {
  return [
    "📒 雞場記帳正式簡化版",
    "",
    "【支出】",
    "草湖 支出 500 電風扇 / 三豐",
    "草湖 零 支出 500 電風扇 / 三豐",
    "",
    "【收入】",
    "仁愛 收入 5000 雞蛋銷售 / 客戶",
    "",
    "付款來源：",
    "未寫＝雞場帳戶",
    "加「零」＝零用金",
    "",
    "【補登日期】",
    "9/28 草湖 支出 500 電風扇 / 三豐",
    "2026/9/28 仁愛 收入 5000 雞蛋銷售 / 客戶",
    "",
    "【跨場代付測試】",
    "草湖幫埤北付5000飼料",
    "埤北還草湖3000",
    "查代付",
    "同步中央測試（只寫入中央主控表的測試分頁）",
    "",
    "【查詢】",
    "今天",
    "本月",
    "今年",
    "草湖 本月",
    "仁愛 9月",
    "2026/9",
    "",
    "【查自己的 LINE ID】",
    "我的ID",
  ].join("\n");
}

// ===== 中央主控表測試彙整：僅允許原中央主控表中的「測試_」分頁 =====
// TEST_CENTRAL_SHEET_ID 必須是預先核准的中央表 ID；不得寫入正式分頁。
const TEST_CENTRAL_SHEET_ID = process.env.TEST_CENTRAL_SHEET_ID || "";
const PRODUCTION_CENTRAL_SHEET_ID = "1xy5sqJuR585wyE3Urq5VtSawa8_lY3FkykEitEVmx3c";
const CENTRAL_TEST_HEADERS = {
  "測試_全場收支彙整": ["交易編號","日期","帳務歸屬","付款場別","收支類型","付款來源","科目代號","科目分類","品項","用途說明","廠商名稱","數量","單價","金額","填表人","來源試算表ID"],
  "測試_代付彙整": ["交易編號","日期","付款場別","帳務歸屬","科目分類","品項","代付金額","填表人","來源試算表ID"],
  "測試_還款彙整": ["交易編號","日期","還款場別","收款場別","還款金額","付款來源","填表人","備註"]
};

function requireSafeCentralTarget(sourceId) {
  if (TEST_CENTRAL_SHEET_ID !== PRODUCTION_CENTRAL_SHEET_ID || TEST_CENTRAL_SHEET_ID === sourceId)
    throw new Error("安全鎖：TEST_CENTRAL_SHEET_ID 必須是指定中央主控表，且不可等於來源帳本");
}

async function buildCentralTestSnapshot(sourceId) {
  const sheets = await getSheets();
  const meta = await sheets.spreadsheets.get({spreadsheetId:sourceId,fields:"sheets.properties.title"});
  const monthTabs = (meta.data.sheets || []).map(x=>x.properties.title)
    .filter(x=>/^\d{6}月$/.test(x)).sort();
  const incomeExpense = [];
  for (const tab of monthTabs) {
    const response = await sheets.spreadsheets.values.get({spreadsheetId:sourceId,range:`'${tab}'!A1:O`,valueRenderOption:"FORMATTED_VALUE"});
    const values = response.data.values || [];
    const headers = values[0] || [];
    const idx = name => headers.indexOf(name);
    const cell = (r,name) => {const i=idx(name);return i<0?"":(r[i]??"");};
    values.slice(1).forEach((r,i)=>{
      const kind=cell(r,"收支類型");
      if (!(["支出","收入"].includes(kind))) return;
      const raw=Number(String(cell(r,"金額")).replace(/,/g,""));
      if (!Number.isFinite(raw) || !cell(r,"日期")) return;
      const sourceMarker=cell(r,"發票或憑證");
      const txId=sourceMarker.startsWith("TEST-TRANSFER:") ? sourceMarker : `${sourceId}:${tab}:${i+2}`;
      const payment=String(cell(r,"付款來源"));
      incomeExpense.push([txId,cell(r,"日期"),cell(r,"場別"),payment.startsWith("代付：")?payment.slice(3):cell(r,"場別"),kind,payment,cell(r,"科目代號"),cell(r,"科目分類"),cell(r,"品項"),cell(r,"用途說明"),cell(r,"廠商名稱"),cell(r,"數量"),cell(r,"單價"),kind==="收入"?-Math.abs(raw):Math.abs(raw),cell(r,"填表人"),sourceId]);
    });
  }
  const transfers = (await transferRows(sourceId)).filter(r=>r.status==="完成");
  const advances = transfers.filter(r=>r.type==="代付").map(r=>[r.id,r.date,r.payer,r.owner,"",r.description,Math.abs(r.amount),"",sourceId]);
  const repayments = transfers.filter(r=>r.type==="還款").map(r=>[r.id,r.date,r.owner,r.payer,-Math.abs(r.amount),"","","還款（測試資料）"]);
  return {"測試_全場收支彙整":incomeExpense,"測試_代付彙整":advances,"測試_還款彙整":repayments};
}

async function syncCentralTest(sourceId) {
  requireSafeCentralTarget(sourceId);
  const snapshot = await buildCentralTestSnapshot(sourceId);
  const sheets = await getSheets();
  // 先確認測試副本可存取，避免把不存在的 ID 當作成功。
  await sheets.spreadsheets.get({spreadsheetId:TEST_CENTRAL_SHEET_ID,fields:"spreadsheetId"});
  for (const [tab, rows] of Object.entries(snapshot)) {
    // 白名單：僅允許程式內定義的三個測試分頁。
    if (!Object.prototype.hasOwnProperty.call(CENTRAL_TEST_HEADERS, tab) || !tab.startsWith("測試_"))
      throw new Error("拒絕寫入非測試分頁");
    const expectedHeaders = CENTRAL_TEST_HEADERS[tab];
    const current = await sheets.spreadsheets.values.get({
      spreadsheetId: TEST_CENTRAL_SHEET_ID, range: `'${tab}'!A1:P`, valueRenderOption: "FORMATTED_VALUE"
    });
    const existing = current.data.values || [];
    if (existing[0]?.some(x => String(x).trim()) && existing[0].join("|") !== expectedHeaders.join("|"))
      throw new Error(`${tab} 的欄位與程式不同，已停止同步以免覆蓋資料`);
    if (!existing[0]?.length) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: TEST_CENTRAL_SHEET_ID, range: `'${tab}'!A1`, valueInputOption: "RAW",
        requestBody: {values: [expectedHeaders]}
      });
    }
    // 僅新增未出現過的交易編號，不清除、不覆寫任何既有列。
    const existingIds = new Set(existing.slice(1).map(r => String(r[0] || "")).filter(Boolean));
    const missing = rows.filter(r => !existingIds.has(String(r[0])));
    if (missing.length) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: TEST_CENTRAL_SHEET_ID, range: `'${tab}'!A:P`,
        valueInputOption: "RAW", insertDataOption: "INSERT_ROWS",
        requestBody: {values: missing}
      });
    }
  }
  return `✅ 中央測試彙整完成\n全場收支：${snapshot["測試_全場收支彙整"].length} 筆\n代付：${snapshot["測試_代付彙整"].length} 筆\n還款：${snapshot["測試_還款彙整"].length} 筆\n只新增中央主控表「測試_」分頁中尚未同步的交易，不覆蓋原資料。`;
}

// ===== 測試專用跨場往來 =====
const TRANSFER_TAB = "跨場往來測試";
const TRANSFER_HEADERS = ["訊息ID", "日期", "類型", "付款場別", "帳務歸屬", "金額", "用途", "狀態", "填表人"];
function parseTransferCommand(body) {
  const t = String(body || "").replace(/\s+/g, "").replace(/，/g, "");
  let m = t.match(/^(.+?)幫(.+?)付([\d,]+(?:\.\d+)?)(.+)$/);
  if (m) {
    const payer = farmFromToken(m[1]), owner = farmFromToken(m[2]);
    const amount = parseAmount(m[3]), description = m[4].trim();
    if (!payer || !owner || payer === "全部" || owner === "全部" || payer === owner) return {error:"代付場別不正確，付款場與歸屬場必須不同。"};
    if (!(amount > 0) || !description) return {error:"請填入有效金額及用途。"};
    return {type:"代付",payer,owner,amount,description};
  }
  m = t.match(/^(.+?)還(.+?)([\d,]+(?:\.\d+)?)$/);
  if (m) {
    const owner = farmFromToken(m[1]), payer = farmFromToken(m[2]), amount = parseAmount(m[3]);
    if (!payer || !owner || payer === "全部" || owner === "全部" || payer === owner || !(amount > 0)) return {error:"還款格式或金額不正確。"};
    return {type:"還款",payer,owner,amount,description:"還款"};
  }
  return null;
}
async function transferRows(sheetId) {
  const sheets = await getSheets();
  await ensureSheet(sheetId, TRANSFER_TAB, TRANSFER_HEADERS);
  const result = await sheets.spreadsheets.values.get({spreadsheetId:sheetId,range:`'${TRANSFER_TAB}'!A2:I`,valueRenderOption:"FORMATTED_VALUE"});
  return (result.data.values || []).map((r,i)=>({row:i+2,id:r[0]||"",date:r[1]||"",type:r[2]||"",payer:r[3]||"",owner:r[4]||"",amount:parseAmount(r[5])||0,description:r[6]||"",status:r[7]||""}));
}
// 相容舊資料：過去還款是正數，新版還款為負數；均以類型決定方向。
function signedTransferAmount(row) {
  if (row.type === "代付") return Math.abs(row.amount);
  if (row.type === "還款") return -Math.abs(row.amount);
  return 0;
}
function transferBalance(rows,payer,owner) {
  return rows.filter(r=>r.status === "完成" && r.payer===payer && r.owner===owner)
    .reduce((n,r)=>n+signedTransferAmount(r),0);
}
async function hasExpenseMarker(sheetId,monthSheet,marker) {
  const sheets=await getSheets();
  const result=await sheets.spreadsheets.values.get({spreadsheetId:sheetId,range:`'${monthSheet}'!M2:M`,valueRenderOption:"FORMATTED_VALUE"});
  return (result.data.values||[]).some(r=>r[0]===marker);
}
async function processTransfer(sheetId,event,dated,cmd,userName) {
  const sheets=await getSheets();
  const id=String(event.message?.id||"");
  if (!id) return "❌ 無法取得訊息編號，為避免重複記帳，已取消操作。";
  const rows=await transferRows(sheetId);
  const previous=rows.find(r=>r.id===id);
  if (previous?.status==="完成") return `✅ 這筆${cmd.type}已經記錄過，沒有重複寫入。`;
  if (cmd.type==="還款" && !previous && cmd.amount>transferBalance(rows,cmd.payer,cmd.owner)) {
    return `❌ 還款金額超過未結清代付：${transferBalance(rows,cmd.payer,cmd.owner).toLocaleString("zh-TW")} 元`;
  }
  let transferRow=previous?.row;
  if (!previous) {
    const added=await sheets.spreadsheets.values.append({spreadsheetId:sheetId,range:`'${TRANSFER_TAB}'!A:I`,valueInputOption:"USER_ENTERED",insertDataOption:"INSERT_ROWS",requestBody:{values:[[id,dated.dateInfo.dateText,cmd.type,cmd.payer,cmd.owner,cmd.type==="還款"?-Math.abs(cmd.amount):Math.abs(cmd.amount),cmd.description,"處理中",userName]]}});
    const updated=added.data.updates?.updatedRange||"";
    const match=updated.match(/![A-Z]+(\d+):/);
    if (!match) throw new Error("跨場紀錄已新增但無法確認列號，請先檢查測試表，不要重送。");
    transferRow=Number(match[1]);
  }
  if (cmd.type==="代付") {
    const marker=`TEST-TRANSFER:${id}`;
    await ensureMonthSheet(sheetId,dated.dateInfo.monthSheet);
    if (!(await hasExpenseMarker(sheetId,dated.dateInfo.monthSheet,marker))) {
      const selected=await autoChooseItem(cmd.description);
      if (!selected) throw new Error("找不到對應科目");
      await writeExpense(sheetId,dated.dateInfo.monthSheet,{
        farm:cmd.owner,txType:"支出",paymentSource:"代付："+cmd.payer,
        code:selected.code,className:selected.className,accountItem:selected.item,
        description:cmd.description,vendor:"",qty:1,unitPrice:cmd.amount,amount:cmd.amount,
        advanceTarget:cmd.payer
      },dated.dateInfo.dateText,userName,marker);
    }
  }
  await sheets.spreadsheets.values.update({spreadsheetId:sheetId,range:`'${TRANSFER_TAB}'!H${transferRow}`,valueInputOption:"RAW",requestBody:{values:[["完成"]]}});
  const newRows=await transferRows(sheetId);
  const balance=transferBalance(newRows,cmd.payer,cmd.owner);
  return [`✅ ${cmd.type}測試記錄完成`,`付款場別：${cmd.payer}`,`帳務歸屬：${cmd.owner}`,`金額：${cmd.amount.toLocaleString("zh-TW")} 元`,cmd.type==="代付"?`用途：${cmd.description}`:null,`目前 ${cmd.owner} 欠 ${cmd.payer}：${balance.toLocaleString("zh-TW")} 元`,`紀錄：${TRANSFER_TAB}`].filter(Boolean).join("\n");
}
async function transferSummary(sheetId) {
  const rows=(await transferRows(sheetId)).filter(r=>r.status==="完成");
  const pairs=new Map();
  for(const r of rows){const key=`${r.owner} → ${r.payer}`;const item=pairs.get(key)||{paid:0,repaid:0};if(r.type==="代付")item.paid+=Math.abs(r.amount);if(r.type==="還款")item.repaid+=Math.abs(r.amount);pairs.set(key,item);}
  const lines=["📒 測試版跨場代付餘額"];
  for(const [key,v] of pairs) {const n=v.paid-v.repaid;if(n!==0) lines.push(`${key}\n代付：${v.paid.toLocaleString("zh-TW")} 元\n還款：${v.repaid.toLocaleString("zh-TW")} 元\n尚欠：${n.toLocaleString("zh-TW")} 元`);}
  if(lines.length===1) lines.push("目前沒有未結清代付款。");
  return lines.join("\n");
}

async function handleTextMessage(event) {
  const text = String(event.message.text || "").trim();

  // ===== 查詢自己的 LINE User ID =====
  if (text === "我的ID") {
    const userId = event.source?.userId || "";
    const userName = await getProfileName(event);

    if (!userId) {
      return "❌ 目前無法取得你的 LINE User ID。";
    }

    return [
      "🪪 LINE 使用者資料",
      "",
      `名稱：${userName}`,
      `LINE User ID：${userId}`,
    ].join("\n");
  }

  if (["說明", "幫助", "help", "HELP", "?"].includes(text)) {
    return helpText();
  }

  const userSheet = await resolveUserSheet(event);
  if (userSheet.error) {
    return userSheet.error;
  }

  const targetSheetId = userSheet.sheetId;

  await ensureBaseStructure(targetSheetId);
  if (text === "查代付") return await transferSummary(targetSheetId);
  if (text === "同步中央測試") return await syncCentralTest(targetSheetId);

  const query = parseQuery(text);
  if (query) {
    return await querySummary(targetSheetId, query);
  }

  const dated = parseOptionalDatePrefix(text);
  if (dated.error) {
    return `❌ ${dated.error}`;
  }

  const transfer = parseTransferCommand(dated.body);
  if (transfer?.error) return `❌ ${transfer.error}`;
  if (transfer) return await processTransfer(targetSheetId,event,dated,transfer,await getProfileName(event));

  const expense = await parseExpenseCommand(dated.body);

  if (expense?.error) {
    return `❌ ${expense.error}`;
  }

  if (expense) {
    const userName = await getProfileName(event);

    await ensureMonthSheet(targetSheetId, dated.dateInfo.monthSheet);
    await writeExpense(
      targetSheetId,
      dated.dateInfo.monthSheet,
      expense,
      dated.dateInfo.dateText,
      userName
    );

    return [
      `✅ ${expense.txType}記錄完成`,
      "",
      dated.isCustomDate ? "🗓️ 補登日期" : null,
      `場別：${expense.farm}`,
      `日期：${dated.dateInfo.dateText}`,
      `品項：${expense.accountItem}`,
      `用途：${expense.description}`,
      expense.vendor ? `廠商：${expense.vendor}` : null,
      `數量：${expense.qty}`,
      `單價：${expense.unitPrice.toLocaleString("zh-TW")} 元`,
      `金額：${expense.amount.toLocaleString("zh-TW")} 元`,
      expense.paymentSource ? `付款來源：${expense.paymentSource}` : null,
      `收支類型：${expense.txType}`,
      `科目代號：${expense.code}`,
      `科目分類：${expense.className}`,
      `填表人：${userName}`,
      `帳本：${userSheet.name}`,
      `寫入：${dated.dateInfo.monthSheet}`,
    ].filter(Boolean).join("\n");
  }

  return helpText();
}

async function handleEvent(event) {
  if (event.type !== "message" || event.message.type !== "text") return null;

  try {
    const replyText = await handleTextMessage(event);

    await client.replyMessage({
      replyToken: event.replyToken,
      messages: [{ type: "text", text: replyText }],
    });

    return "OK";
  } catch (err) {
    console.error("處理訊息錯誤：", err);

    try {
      await client.replyMessage({
        replyToken: event.replyToken,
        messages: [{
          type: "text",
          text: "❌ 系統發生錯誤，請稍後再試。",
        }],
      });
    } catch (_) {}

    return "ERROR";
  }
}

app.get("/", (req, res) => {
  res.send("Chicken Farm Accounting Bot - Simplified Income Expense + Keyword Master Version is running.");
});

app.post("/webhook", line.middleware(config), async (req, res) => {
  try {
    await Promise.all(req.body.events.map(handleEvent));
    res.status(200).end();
  } catch (err) {
    console.error("Webhook Error:", err);
    res.status(500).end();
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Chicken Farm Accounting Bot - Simplified Income Expense + Keyword Master Version running on port ${PORT}`);
});
