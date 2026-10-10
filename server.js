require("dotenv").config();

const express = require("express");
const line = require("@line/bot-sdk");
const { google } = require("googleapis");
const { randomUUID } = require("node:crypto");

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

// ===== 場別 =====
const FARM_ALIASES = {
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
  requireCentralTestConfig();
  const userId = event.source?.userId || "";

  // 如果尚未設定主檔，暫時沿用原本單人版 Sheet。
  if (!TEST_MASTER_ID) {
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

  const masterTitle = "使用者名單";
  if (!masterTitle) {
    throw new Error("使用者主檔沒有可讀取的分頁");
  }

  const sheets = await getSheets();
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: TEST_MASTER_ID,
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
  // 先確認分頁是否存在，避免對不存在的月份讀取 A1:O 而發生 400。
  // 已存在的月份完全不更動欄位或歷史資料。
  const sheets = await getSheets();
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: sheetId,
    fields: "sheets.properties.title",
  });
  const exists = (meta.data.sheets || []).some(
    s => s.properties.title === monthSheet
  );
  if (exists) return;

  // 只在缺少月份時建立新分頁及原程式的欄位標題。
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

async function writeExpense(sheetId, monthSheet, expense, dateText, userName, userId) {
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
        "",
        expense.advanceTarget || "",
        userName || "",
      ]],
    },
  });

  // 記錄實際寫入時間，跨月份查詢依此排序；記帳成功後才寫日誌。
  // 日誌失敗不應讓使用者重複記帳。
  let auditWarning = false;
  try {
    const updatedRange = appendResult.data?.updates?.updatedRange || "";
    const rowMatch = updatedRange.match(/![A-Z]+(\d+):[A-Z]+(\d+)$/);
    const rowNumber = rowMatch ? Number(rowMatch[1]) : null;
    if (!rowNumber) throw new Error("Google Sheets 未回傳寫入列號");
    await appendAccountingLog(sheetId, {
      id: `TX-${randomUUID()}`, userId: userId || "",
      timestamp: new Date().toISOString(), monthSheet, rowNumber,
      farm: expense.farm, txType: expense.txType || "支出",
      item: expense.accountItem, amount: expense.amount, recorder: userName || "",
    });
  } catch (auditErr) {
    auditWarning = true;
    console.warn("記帳已成功，但寫入排序紀錄失敗：", auditErr.message);
  }

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
  return { auditWarning };
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
    "【查詢】",
    "今天",
    "本月",
    "今年",
    "草湖 本月",
    "仁愛 9月",
    "2026/9",
    "",
    "最近10筆",
    "刪除上一筆（限本人新記帳）",
    "刪除第2筆（須先查最近10筆，5分鐘內有效）",
    "刪除 TX-完整記帳ID（限本人）",
    "確認刪除 / 取消",
    "",
    "【中央查帳（僅測試管理員）】",
    "中央 本月 / 中央 202609 / 中央 今日明細",
    "中央 場別統計 / 中央 同步",
    "",
    "【查自己的 LINE ID】",
    "我的ID",
  ].join("\n");
}


// 最近10筆專用：獨立記錄分頁，不改動每月 A～O 15 欄。
const RECENT_LOG_TAB = "系統記帳順序";
const RECENT_LOG_HEADERS = ["寫入時間UTC", "月份分頁", "列號", "場別", "收支類型", "品項", "金額", "填表人", "記帳ID", "LINE User ID", "狀態", "刪除時間UTC"];

async function appendAccountingLog(sheetId, record) {
  await ensureSheet(sheetId, RECENT_LOG_TAB, RECENT_LOG_HEADERS);
  const sheets = await getSheets();
  await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId,
    range: `'${RECENT_LOG_TAB}'!A:L`,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: [[record.timestamp, record.monthSheet, record.rowNumber,
      record.farm, record.txType, record.item, record.amount, record.recorder,
      record.id, record.userId, "有效", ""]] },
  });
}

// 所有新資料均有永久記帳ID。舊日誌的 I、J 欄為空，不能用於本人刪除。
async function readAccountingLog(sheetId) {
  const sheets = await getSheets();
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: sheetId, fields: "sheets.properties(title)"
  });
  if (!(meta.data.sheets || []).some(s => s.properties.title === RECENT_LOG_TAB)) return [];
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId, range: `'${RECENT_LOG_TAB}'!A2:L`,
    valueRenderOption: "FORMATTED_VALUE"
  });
  return (result.data.values || []).map((r, index) => ({
    logRow: index + 2, timestamp: String(r[0] || ""), month: String(r[1] || ""),
    row: Number(r[2]), farm: String(r[3] || ""), txType: String(r[4] || ""),
    item: String(r[5] || ""), amount: Number(String(r[6] || "").replace(/,/g, "")),
    recorder: String(r[7] || ""), id: String(r[8] || ""),
    userId: String(r[9] || ""), status: String(r[10] || "有效"),
    deletedAt: String(r[11] || "")
  }));
}

// 刪除確認狀態存在記憶體中，服務重新部署後需重新下指令。
const pendingDeletions = new Map();
const DELETE_CONFIRM_MS = 5 * 60 * 1000;
// 最近10筆編號快照：編號只在本人最近一次查詢後5分鐘有效。
const recentSnapshots = new Map();
const RECENT_SNAPSHOT_MS = 5 * 60 * 1000;
const TX_ID_PATTERN = /^TX-[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

async function lastOwnEntry(sheetId, userId) {
  if (!userId) return null;
  const entries = await readAccountingLog(sheetId);
  return entries.filter(e => e.id && e.userId === userId && e.status !== "已刪除" &&
    /^\d{6}月$/.test(e.month) && e.row >= 2 && !Number.isNaN(Date.parse(e.timestamp)))
    .sort((a, b) => b.timestamp.localeCompare(a.timestamp))[0] || null;
}

function formatDeletionEntry(e) {
  return [`日期所在分頁：${e.month}`, `場別：${e.farm}`,
    `品項：${e.item}`, `收支類型：${e.txType}`,
    `金額：${Math.abs(e.amount).toLocaleString("zh-TW")} 元`,
    `填表人：${e.recorder}`, `記帳ID：${e.id}`].join("\n");
}

async function requestDeleteLast(sheetId, userId) {
  const entry = await lastOwnEntry(sheetId, userId);
  if (!entry) return "目前沒有可刪除的本人記帳。舊資料沒有 LINE User ID，為安全起見不能透過此指令刪除。";
  pendingDeletions.set(`${sheetId}:${userId}`, { id: entry.id, mode: "own-last", expiresAt: Date.now() + DELETE_CONFIRM_MS });
  return `⚠️ 即將永久刪除你最後寫入的記帳\n\n${formatDeletionEntry(entry)}\n\n請於5分鐘內輸入「確認刪除」，或輸入「取消」。\n此操作會真正移除月份分頁中的整列。`;
}

async function requestSelectedDelete(sheetId, userId, id) {
  if (!userId || !TX_ID_PATTERN.test(id)) return "❌ 記帳ID格式不正確。";
  const entries = await readAccountingLog(sheetId);
  const entry = entries.find(e => e.id.toLowerCase() === id.toLowerCase() &&
    e.userId === userId && e.status !== "已刪除" &&
    /^\d{6}月$/.test(e.month) && e.row >= 2);
  if (!entry) return "❌ 找不到本人可刪除的有效記帳；舊資料或其他人的帳不能刪除。";
  pendingDeletions.set(`${sheetId}:${userId}`, {
    id: entry.id, mode: "selected", expiresAt: Date.now() + DELETE_CONFIRM_MS
  });
  return `⚠️ 即將永久刪除指定記帳\n\n${formatDeletionEntry(entry)}\n\n請於5分鐘內輸入「確認刪除」，或輸入「取消」。\n此操作會真正移除月份分頁中的整列。`;
}

async function requestDeleteByNumber(sheetId, userId, number) {
  const snapshot = recentSnapshots.get(`${sheetId}:${userId}`);
  if (!snapshot || Date.now() > snapshot.expiresAt) {
    return "⚠️ 查詢清單已過期，請先輸入「最近10筆」再指定刪除。";
  }
  const id = snapshot.ids[number - 1];
  if (!id) return "❌ 這個編號沒有可刪除的本人記帳（舊帳或其他人記帳不可刪）。";
  return requestSelectedDelete(sheetId, userId, id);
}

function sheetCell(value) {
  return { userEnteredValue: typeof value === "number" ? { numberValue: value } : { stringValue: String(value) } };
}

async function confirmDeleteLast(sheetId, userId) {
  const key = `${sheetId}:${userId}`;
  const pending = pendingDeletions.get(key);
  pendingDeletions.delete(key); // 一次性確認，避免重送
  if (!pending || !["own-last", "selected"].includes(pending.mode) || Date.now() > pending.expiresAt) return "⚠️ 沒有相符的待確認刪除，或已超過5分鐘。請重新下指令。";

  const sheets = await getSheets();
  const entries = await readAccountingLog(sheetId);
  const entry = entries.find(e => e.id === pending.id && e.userId === userId && e.status !== "已刪除");
  if (!entry) return "❌ 找不到本人待刪除的有效記帳，已取消。";
  if (pending.mode === "own-last") {
    const newest = await lastOwnEntry(sheetId, userId);
    if (!newest || newest.id !== entry.id) return "⚠️ 期間又有新記帳，請重新輸入「刪除上一筆」。";
  }

  const meta = await sheets.spreadsheets.get({
    spreadsheetId: sheetId, fields: "sheets.properties(sheetId,title)"
  });
  const tabs = (meta.data.sheets || []).map(s => s.properties);
  const monthTab = tabs.find(t => t.title === entry.month);
  const logTab = tabs.find(t => t.title === RECENT_LOG_TAB);
  if (!monthTab || !logTab) return "❌ 分頁不存在，未執行刪除。";
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId, range: `'${entry.month}'!A${entry.row}:O${entry.row}`,
    valueRenderOption: "FORMATTED_VALUE"
  });
  const row = result.data.values?.[0] || [];
  const headResult = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId, range: `'${entry.month}'!A1:O1`,
    valueRenderOption: "FORMATTED_VALUE"
  });
  const headers = headResult.data.values?.[0] || [];
  const value = (name) => { const i = headerIndex(headers, [name]); return i >= 0 ? String(row[i] || "").trim() : ""; };
  const rowAmount = Number(value("金額").replace(/,/g, ""));
  if (!row.length || value("場別") !== entry.farm ||
      (value("收支類型") || "支出") !== entry.txType || value("品項") !== entry.item ||
      Math.abs(rowAmount) !== Math.abs(entry.amount) || value("填表人") !== entry.recorder) {
    return "❌ 月份分頁資料與記帳紀錄不一致，為避免刪錯帳已停止。請人工檢查。";
  }

  // 一個 batchUpdate 同時刪除月份列、標記日誌刪除、調整後續列號。
  // 記帳ID不變；列號僅作定位，不能當作中央同步主鍵。
  const requests = [{ deleteDimension: {
    range: { sheetId: monthTab.sheetId, dimension: "ROWS", startIndex: entry.row - 1, endIndex: entry.row }
  }}];
  for (const e of entries) {
    if (e.id === entry.id) {
      requests.push({ updateCells: {
        range: { sheetId: logTab.sheetId, startRowIndex: e.logRow - 1,
          endRowIndex: e.logRow, startColumnIndex: 10, endColumnIndex: 12 },
        rows: [{ values: [sheetCell("已刪除"), sheetCell(new Date().toISOString())] }],
        fields: "userEnteredValue"
      }});
    } else if (e.month === entry.month && e.row > entry.row && e.status !== "已刪除") {
      requests.push({ updateCells: {
        range: { sheetId: logTab.sheetId, startRowIndex: e.logRow - 1,
          endRowIndex: e.logRow, startColumnIndex: 2, endColumnIndex: 3 },
        rows: [{ values: [sheetCell(e.row - 1)] }], fields: "userEnteredValue"
      }});
    }
  }
  await sheets.spreadsheets.batchUpdate({ spreadsheetId: sheetId, requestBody: { requests } });
  return `✅ 已永久刪除本人記帳\n\n${formatDeletionEntry(entry)}\n\n月份分頁已移除該列，系統日誌保留刪除狀態供日後中央同步。`;
}

async function recentTenText(sheetId, userId) {
  const sheets = await getSheets();
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: sheetId,
    fields: "sheets.properties.title",
  });
  const titles = (meta.data.sheets || []).map(s => s.properties.title);
  const months = titles.filter(t => /^\d{6}月$/.test(t)).sort().reverse();
  const hasLog = titles.includes(RECENT_LOG_TAB);
  let logRows = [];
  if (hasLog) {
    const logResult = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: `'${RECENT_LOG_TAB}'!A2:L`,
      valueRenderOption: "FORMATTED_VALUE",
    });
    logRows = (logResult.data.values || []).map(r => ({
      timestamp: String(r[0] || ""), month: String(r[1] || ""),
      row: Number(r[2]), status: String(r[10] || "有效"),
      id: String(r[8] || ""), userId: String(r[9] || ""),
    })).filter(r => /^\d{6}月$/.test(r.month) && r.row >= 2 && !Number.isNaN(Date.parse(r.timestamp)) && r.status !== "已刪除");
  }

  // 先依實際寫入時間排序；舊資料沒有時間紀錄時只能用月份及列號近似。
  const ordered = logRows.sort((a, b) => b.timestamp.localeCompare(a.timestamp));
  const byMonth = new Map();
  async function monthRows(month) {
    if (byMonth.has(month)) return byMonth.get(month);
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId, range: `'${month}'!A1:O`,
      valueRenderOption: "FORMATTED_VALUE",
    });
    const values = result.data.values || [];
    const headers = values[0] || [];
    const idx = names => headerIndex(headers, names);
    const indices = {
      farm: idx(["場別", "帳務歸屬"]), date: idx(["日期"]),
      txType: idx(["收支類型"]), item: idx(["品項"]),
      amount: idx(["金額"]), recorder: idx(["填表人"]),
    };
    const rows = values.slice(1).map((r, i) => {
      const get = key => indices[key] < 0 ? "" : String(r[indices[key]] || "").trim();
      return { month, row: i + 2, farm: get("farm"), date: get("date"),
        txType: get("txType"), item: get("item"), amount: get("amount"),
        recorder: get("recorder") };
    }).filter(r => r.date || r.item || r.amount);
    byMonth.set(month, rows);
    return rows;
  }

  const recent = [];
  const used = new Set();
  for (const log of ordered) {
    if (recent.length >= 10) break;
    if (!months.includes(log.month)) continue;
    const key = `${log.month}:${log.row}`;
    if (used.has(key)) continue;
    const row = (await monthRows(log.month)).find(r => r.row === log.row);
    if (!row) continue;
    used.add(key);
    recent.push({ ...row, approximate: false, id: log.id, ownerId: log.userId });
  }

  // 補齊沒有排序紀錄的舊資料。不能宣稱這些舊資料有精確寫入先後。
  if (recent.length < 10) {
    for (const month of months) {
      const rows = await monthRows(month);
      for (const row of rows.slice().reverse()) {
        if (recent.length >= 10) break;
        const key = `${month}:${row.row}`;
        if (used.has(key)) continue;
        used.add(key);
        recent.push({ ...row, approximate: true, id: "", ownerId: "" });
      }
      if (recent.length >= 10) break;
    }
  }

  if (!recent.length) {
    recentSnapshots.delete(`${sheetId}:${userId}`);
    return "📒 最近10筆記帳\n\n目前沒有記帳資料。";
  }
  recentSnapshots.set(`${sheetId}:${userId}`, {
    ids: recent.map(r => r.ownerId === userId && TX_ID_PATTERN.test(r.id) ? r.id : null),
    expiresAt: Date.now() + RECENT_SNAPSHOT_MS
  });
  const lines = ["📒 最近10筆記帳（同帳本所有人）", ""];
  for (let i = 0; i < recent.length; i++) {
    const r = recent[i];
    const amount = Number(r.amount.replace(/,/g, ""));
    const formatted = Number.isFinite(amount) ? Math.abs(amount).toLocaleString("zh-TW") : r.amount;
    const date = r.date.replace(/^\d{4}[\/-]/, "");
    lines.push(`${i + 1}. ${date} ${r.farm}｜${r.txType || "支出"}｜${r.item}｜${formatted}元${r.approximate ? " ※" : ""}`);
    lines.push(`   填表人：${r.recorder || "未填"}`);
    if (r.ownerId === userId && TX_ID_PATTERN.test(r.id)) {
      lines.push(`   可刪除：刪除第${i + 1}筆`);
      lines.push(`   記帳ID：${r.id}`);
    }
  }
  lines.push("", `共顯示 ${recent.length} 筆`);
  lines.push("指定編號須在查詢後5分鐘內使用；只能刪除本人有記帳ID的紀錄。");
  if (recent.some(r => r.approximate)) {
    lines.push("※ 舊資料沒有寫入時間，只能按月份及列順序推估；補登可能不是實際先後。 ");
  }
  return lines.join("\n");
}


// ===== 中央查帳（候選版；僅供隔離 TESTING） =====
// 所有跨帳本操作只讀來源帳本。只有明確下「中央 同步」且配置
// TEST_CENTRAL_REPORT_SHEET_ID 才會寫入隔離的測試報表，不寫入正式中央主控表。
const CENTRAL_TEST_MODE = process.env.APP_ENV === "TESTING";
const TEST_ALLOWED_IDS = new Set(String(process.env.TEST_ALLOWED_SHEET_IDS || "")
  .split(",").map(s => s.trim()).filter(Boolean));
const TEST_MASTER_ID = process.env.TEST_USER_MASTER_SHEET_ID || "";
const TEST_REPORT_ID = process.env.TEST_CENTRAL_REPORT_SHEET_ID || "";

function requireCentralTestConfig() {
  if (!CENTRAL_TEST_MODE || !TEST_MASTER_ID || !TEST_ALLOWED_IDS.size) {
    throw new Error("中央測試未隔離：需 APP_ENV=TESTING、TEST_USER_MASTER_SHEET_ID、TEST_ALLOWED_SHEET_IDS");
  }
  if (TEST_MASTER_ID === MASTER_SHEET_ID) {
    throw new Error("測試使用者主檔不可與正式 MASTER_SHEET_ID 相同");
  }
  // 測試帳本白名單不得包含中央主控表或中央科目主檔（含測試主控表）。
  const protectedIds = [MASTER_SHEET_ID, ACCOUNT_MASTER_SHEET_ID, TEST_MASTER_ID]
    .filter(Boolean).map(id => String(id).trim());
  if (protectedIds.some(id => TEST_ALLOWED_IDS.has(id))) {
    throw new Error("TEST_ALLOWED_SHEET_IDS 含中央主控表或科目主檔 ID，已拒絕中央測試操作");
  }
}

async function centralMasterRows() {
  requireCentralTestConfig();
  const sheets = await getSheets();
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: TEST_MASTER_ID,
    range: "'使用者名單'!A2:F",
    valueRenderOption: "FORMATTED_VALUE",
  });
  return (response.data.values || []).map(r => ({
    name: String(r[0] || "").trim(),
    userId: String(r[2] || "").trim(),
    sheetId: String(r[3] || "").trim(),
    enabled: ["Y", "YES", "TRUE", "1", "啟用"].includes(String(r[4] || "").trim().toUpperCase()),
  }));
}

async function assertCentralAdmin(userId) {
  requireCentralTestConfig();
  const sheets = await getSheets();
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: TEST_MASTER_ID,
    range: "'權限設定'!A2:D",
    valueRenderOption: "FORMATTED_VALUE",
  });
  const matches = (result.data.values || []).filter(r => String(r[0] || "").trim() === userId);
  if (matches.length !== 1) return false;
  const r = matches[0];
  return String(r[1] || "").trim().toUpperCase() === "ADMIN" &&
    String(r[2] || "").trim().toUpperCase() === "ALL" &&
    ["Y", "YES", "TRUE", "1", "啟用"].includes(String(r[3] || "").trim().toUpperCase());
}

async function centralRead(months) {
  const users = (await centralMasterRows()).filter(r => r.enabled && r.sheetId);
  const unique = new Map();
  for (const u of users) {
    if (!TEST_ALLOWED_IDS.has(u.sheetId)) continue; // 禁止測試版讀取未核准帳本
    if (!unique.has(u.sheetId)) unique.set(u.sheetId, u.name || "未命名帳本");
  }
  const sheets = await getSheets();
  const all = [], failures = [];
  for (const [sheetId, owner] of unique) {
    try {
      const meta = await sheets.spreadsheets.get({ spreadsheetId: sheetId, fields: "sheets.properties.title" });
      const exists = new Set((meta.data.sheets || []).map(t => t.properties.title));
      for (const month of months) {
        if (!exists.has(month)) continue;
        const rows = await readMonthRows(sheetId, month); // 純讀取，不呼叫 normalizeMonthAmountSigns
        for (const r of rows) all.push({ ...r, owner, month });
      }
    } catch (e) {
      failures.push(`${owner}：讀取失敗（${e.code || "權限或網路問題"}）`);
    }
  }
  return { all, failures, accounts: unique.size, registered: users.length };
}

function centralTotals(rows) {
  const valid = rows.filter(r => ["收入", "支出"].includes(r.txType || "支出"));
  const expense = valid.filter(r => (r.txType || "支出") === "支出")
    .reduce((n, r) => n + Math.abs(r.amount), 0);
  const income = valid.filter(r => r.txType === "收入")
    .reduce((n, r) => n + Math.abs(r.amount), 0);
  return { expense, income, net: expense - income, count: valid.length };
}

function centralFmt(n) { return Number(n).toLocaleString("zh-TW"); }
function centralMonthFromCommand(text) {
  const now = taipeiNow();
  if (text === "中央 本月" || text === "中央 場別統計" || text === "中央 同步") return `${now.year}${now.month}月`;
  const m = text.match(/^中央\s+(\d{4})(\d{2})$/);
  if (m && Number(m[2]) >= 1 && Number(m[2]) <= 12) return `${m[1]}${m[2]}月`;
  return null;
}

async function centralQuery(text) {
  const now = taipeiNow();
  const month = centralMonthFromCommand(text) || now.monthSheet;
  const { all, failures, accounts, registered } = await centralRead([month]);
  // 任一帳本讀取失敗時，不輸出可能誤導使用者的部分總額。
  if (failures.length) {
    return {
      text: [
        `❌ 中央查帳未完成：${failures.length} 份帳本讀取失敗，已停止顯示統計總額。`,
        ...failures.slice(0, 5),
        "請確認測試帳本讀取權限後重試。",
      ].join("\n"),
      rows: [], failures, month,
    };
  }
  let rows = all;
  if (text === "中央 今日明細") {
    rows = all.filter(r => r.date.replace(/-/g, "/") === now.dateText);
  }
  const totals = centralTotals(rows);
  const lines = [
    `📊 中央查帳｜${text === "中央 今日明細" ? now.dateText : month}`,
    `帳本：${accounts} 份（啟用使用者 ${registered} 位，重複帳本僅計一次）`,
    `支出：${centralFmt(totals.expense)} 元`,
    `收入：${centralFmt(totals.income)} 元`,
    `淨支出：${centralFmt(totals.net)} 元`,
    `筆數：${totals.count}`,
  ];
  if (text === "中央 場別統計") {
    const farms = new Map();
    for (const r of rows) {
      const k = r.farm || "未指定";
      if (!farms.has(k)) farms.set(k, []);
      farms.get(k).push(r);
    }
    lines.push("", "【場別統計】");
    for (const [farm, records] of [...farms].sort((a,b) => a[0].localeCompare(b[0], "zh-TW"))) {
      const t = centralTotals(records);
      lines.push(`${farm}：支出 ${centralFmt(t.expense)}／收入 ${centralFmt(t.income)} 元`);
    }
  } else if (text === "中央 今日明細") {
    lines.push("", "【今日明細（最多20筆）】");
    for (const r of rows.slice(0, 20)) {
      lines.push(`${r.owner}｜${r.farm}｜${r.txType}｜${r.item}｜${centralFmt(Math.abs(r.amount))}`);
    }
    if (rows.length > 20) lines.push(`另有 ${rows.length - 20} 筆未顯示，請至中央報表查看。`);
  }
  if (!accounts) lines.push("⚠️ 尚未有核准的測試帳本，請設定 TEST_ALLOWED_SHEET_IDS。");
  return { text: lines.join("\n"), rows, failures, month };
}

async function centralSyncTestReport() {
  requireCentralTestConfig();
  if (!TEST_REPORT_ID || TEST_REPORT_ID === MASTER_SHEET_ID || TEST_REPORT_ID === TEST_MASTER_ID ||
      TEST_ALLOWED_IDS.has(TEST_REPORT_ID)) {
    throw new Error("未設定獨立的 TEST_CENTRAL_REPORT_SHEET_ID，或報表與帳本/主檔衝突");
  }
  const now = taipeiNow();
  const { all, failures, accounts } = await centralRead([now.monthSheet]);
  if (failures.length || !accounts) return "❌ 有帳本讀取失敗或無核准帳本，已停止同步，避免產生不完整報表。";
  const sheets = await getSheets();
  const tabs = [
    ["中央流水帳", ["帳本持有人", "月份", "場別", "收支類型", "付款來源", "日期", "品項", "用途說明", "廠商名稱", "金額", "填表人"],
      all.map(r => [r.owner, r.month, r.farm, r.txType, r.paymentSource, r.date, r.item, r.description, r.vendor, r.amount, r.recorder])],
    ["中央月報", ["月份", "場別", "支出", "收入", "淨支出", "筆數"], (() => {
      const groups = new Map();
      for (const r of all) { const key = r.farm || "未指定"; if (!groups.has(key)) groups.set(key, []); groups.get(key).push(r); }
      return [...groups].map(([farm, rows]) => { const t = centralTotals(rows); return [now.monthSheet, farm, t.expense, t.income, t.net, t.count]; });
    })()],
    ["同步紀錄", ["同步時間UTC", "月份", "帳本數", "明細筆數", "狀態"], [[new Date().toISOString(), now.monthSheet, accounts, all.length, "成功"]]],
  ];
  // 只寫入明確指定的隔離測試報表；不修改原始帳本。
  for (const [name, headers, data] of tabs) {
    await ensureSheet(TEST_REPORT_ID, name, headers);
    if (name !== "同步紀錄") {
      await sheets.spreadsheets.values.clear({ spreadsheetId: TEST_REPORT_ID, range: `'${name}'!A2:Z` });
      if (data.length) await sheets.spreadsheets.values.update({
        spreadsheetId: TEST_REPORT_ID, range: `'${name}'!A2`, valueInputOption: "RAW",
        requestBody: { values: data },
      });
    } else {
      await sheets.spreadsheets.values.append({ spreadsheetId: TEST_REPORT_ID, range: `'${name}'!A:E`,
        valueInputOption: "RAW", insertDataOption: "INSERT_ROWS", requestBody: { values: data } });
    }
  }
  return `✅ 已同步 ${accounts} 份測試帳本、${all.length} 筆到隔離測試報表（${now.monthSheet}）。`;
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

  // 測試版保護：所有既有寫入、刪除、查詢均不得觸及正式帳本。
  requireCentralTestConfig();
  if (!TEST_ALLOWED_IDS.has(userSheet.sheetId)) {
    return "❌ TESTING 安全限制：你的帳本不在測試白名單，未執行操作。";
  }

  if (text.startsWith("中央 ")) {
    if (!(await assertCentralAdmin(event.source?.userId || ""))) return "❌ 僅管理員可使用中央查帳。";
    if (text === "中央 同步") return await centralSyncTestReport();
    if (centralMonthFromCommand(text) || text === "中央 今日明細") {
      return (await centralQuery(text)).text;
    }
    return "中央指令：中央 本月／中央 202609／中央 今日明細／中央 場別統計／中央 同步";
  }

  const targetSheetId = userSheet.sheetId;

  await ensureBaseStructure(targetSheetId);

  const userId = event.source?.userId || "";
  if (text === "刪除上一筆") return await requestDeleteLast(targetSheetId, userId);
  const numberDelete = text.match(/^刪除第\s*(\d{1,2})\s*筆$/);
  if (numberDelete) return await requestDeleteByNumber(targetSheetId, userId, Number(numberDelete[1]));
  const idDelete = text.match(/^刪除\s+(TX-[0-9a-f-]+)$/i);
  if (idDelete) return await requestSelectedDelete(targetSheetId, userId, idDelete[1]);
  if (text === "確認刪除") return await confirmDeleteLast(targetSheetId, userId);
  if (text === "取消") {
    pendingDeletions.delete(`${targetSheetId}:${userId}`);
    return "已取消刪除。";
  }

  if (text === "最近10筆") {
    return await recentTenText(targetSheetId, userId);
  }

  const query = parseQuery(text);
  if (query) {
    return await querySummary(targetSheetId, query);
  }

  const dated = parseOptionalDatePrefix(text);
  if (dated.error) {
    return `❌ ${dated.error}`;
  }

  const expense = await parseExpenseCommand(dated.body);

  if (expense?.error) {
    return `❌ ${expense.error}`;
  }

  if (expense) {
    const userName = await getProfileName(event);

    await ensureMonthSheet(targetSheetId, dated.dateInfo.monthSheet);
    const writeResult = await writeExpense(
      targetSheetId,
      dated.dateInfo.monthSheet,
      expense,
      dated.dateInfo.dateText,
      userName,
      event.source?.userId || ""
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
      writeResult?.auditWarning ? "⚠️ 記帳成功，但排序紀錄未完成，請勿重複送出。" : null,
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

if (!CENTRAL_TEST_MODE) {
  console.error("安全停止：此候選檔僅允許 APP_ENV=TESTING，禁止部署 MAIN。");
  process.exit(1);
}
app.listen(PORT, () => {
  console.log(`Chicken Farm Accounting Bot - Simplified Income Expense + Keyword Master Version running on port ${PORT}`);
});
