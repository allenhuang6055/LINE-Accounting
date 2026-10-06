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

  "班神": "班神場",
  "班神場": "班神場",
  "昊陽": "昊陽場",
  "昊陽場": "昊陽場",

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

function accountByItem(item) {
  const row = ACCOUNT_ITEMS.find(r => r[0] === item);
  return row ? { item: row[0], code: String(row[1]), className: row[2] } : null;
}

function autoChooseItem(description) {
  const text = String(description || "").trim();

  // 使用者直接輸入完整 Excel 品項時優先
  const exact = ACCOUNT_ITEMS
    .map(r => r[0])
    .sort((a, b) => b.length - a.length)
    .find(item => text === item || text.startsWith(item + " "));

  if (exact) return accountByItem(exact);

  for (const rule of KEYWORD_RULES) {
    if (rule.keys.some(k => text.includes(k))) {
      return accountByItem(rule.item);
    }
  }

  return accountByItem("雜項費用");
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
  // 資料分頁
  await ensureSheet(sheetId, "資料", ["品項", "科目代號", "科目分類"]);

  const sheets = await getSheets();
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: "'資料'!A2:C",
  });

  if (!(result.data.values || []).length) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: sheetId,
      range: `'資料'!A2:C${ACCOUNT_ITEMS.length + 1}`,
      valueInputOption: "RAW",
      requestBody: { values: ACCOUNT_ITEMS },
    });
  }
}

async function ensureMonthSheet(sheetId, monthSheet) {
  await ensureSheet(sheetId, monthSheet, [
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
  ]);
}

function parseExpenseCommand(text) {
  const clean = String(text || "").trim().replace(/\s+/g, " ");
  const parts = clean.split(" ");

  let farm = farmFromToken(parts[0]);
  let txType = "";
  let paymentSource = "雞場帳戶";
  let rest;

  // 支援：
  // 草湖 支出 500 電風扇 / 三豐
  // 草湖 零 支出 500 電風扇 / 三豐
  // 仁愛 收入 5000 雞蛋銷售 / 客戶
  // 仁愛 代墊 1500 油資 / 台塑
  // 仁愛 代墊收回 1500
  // 仁愛 收回代墊 1500
  if (farm) {
    if (parts[1] === "零") {
      paymentSource = "零用金";
      txType = parts[2];
      rest = parts.slice(3).join(" ");
    } else {
      txType = parts[1];
      rest = parts.slice(2).join(" ");
    }
  } else if (["支出", "收入", "代墊", "代墊收回", "收回代墊"].includes(parts[0])) {
    farm = farmFromToken(parts[1]);
    if (!farm) return { error: "請先輸入場別，例如：仁愛 代墊 1500 油資" };
    txType = parts[0];
    rest = parts.slice(2).join(" ");
  } else {
    return null;
  }

  if (txType === "收回代墊") txType = "代墊收回";

  if (!["支出", "收入", "代墊", "代墊收回"].includes(txType)) {
    return null;
  }

  if (farm === "全部") {
    return { error: "記帳時不能使用「全部」，請指定實際場別。" };
  }

  // 代墊與代墊收回先不歸類到雞場帳戶/零用金
  if (["代墊", "代墊收回"].includes(txType)) {
    paymentSource = "";
  }

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
    // 代墊收回可只打金額；其他類型需有用途
    m = rest.match(/^([\d,]+(?:\.\d+)?)(?:\s+(.+))?$/);
    if (!m) {
      return { error: `格式例如：仁愛 ${txType} 1500${txType === "代墊收回" ? "" : " 油資"}` };
    }

    amount = parseAmount(m[1]);
    unitPrice = amount;
    description = String(m[2] || "").trim();
  }

  if (!Number.isFinite(amount) || amount <= 0) {
    return { error: "金額格式不正確" };
  }

  if (!description && txType !== "代墊收回") {
    return { error: `請輸入用途，例如：仁愛 ${txType} 1500 油資` };
  }

  if (!description && txType === "代墊收回") {
    description = "代墊還款";
  }

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

  const itemNames = ACCOUNT_ITEMS.map(r => r[0]).sort((a, b) => b.length - a.length);

  for (const item of itemNames) {
    if (description === item || description.startsWith(item + " ")) {
      selected = accountByItem(item);
      purpose = description.slice(item.length).trim() || item;
      break;
    }
  }

  if (!selected) selected = autoChooseItem(description);

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
    advanceTarget: ["代墊", "代墊收回"].includes(txType) ? farm : "",
  };
}

async function writeExpense(sheetId, monthSheet, expense, dateText, userName) {
  const sheets = await getSheets();

  await sheets.spreadsheets.values.append({
    spreadsheetId: sheetId,
    range: `'${monthSheet}'!A:O`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [[
        expense.farm,
        expense.code,
        expense.className,
        dateText,
        expense.accountItem,
        expense.description,
        expense.vendor,
        expense.qty,
        expense.unitPrice,
        expense.amount,
        "",
        expense.paymentSource || "",
        expense.txType || "支出",
        expense.advanceTarget || "",
        userName || "",
      ]],
    },
  });
}

async function readMonthRows(sheetId, monthSheet) {
  const sheets = await getSheets();

  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range: `'${monthSheet}'!A2:O`,
    valueRenderOption: "FORMATTED_VALUE",
  });

  return (result.data.values || []).map(r => ({
    farm: r[0] || "",
    code: r[1] || "",
    className: r[2] || "",
    date: r[3] || "",
    item: r[4] || "",
    description: r[5] || "",
    vendor: r[6] || "",
    qty: Number(String(r[7] || "0").replace(/,/g, "")) || 0,
    unitPrice: Number(String(r[8] || "0").replace(/,/g, "")) || 0,
    amount: Number(String(r[9] || "0").replace(/,/g, "")) || 0,
    receipt: r[10] || "",
    paymentSource: r[11] || "",
    txType: r[12] || "支出",
    advanceTarget: r[13] || "",
    recorder: r[14] || "",
  })).filter(r => r.date || r.item || r.amount);
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
  const filtered = farm === "全部" ? rows : rows.filter(r => r.farm === farm);

  const expenseRows = filtered.filter(r => (r.txType || "支出") === "支出");
  const incomeRows = filtered.filter(r => r.txType === "收入");
  const advanceRows = filtered.filter(r => r.txType === "代墊");
  const recoveredRows = filtered.filter(r => r.txType === "代墊收回");

  const expenseTotal = expenseRows.reduce((sum, r) => sum + r.amount, 0);
  const incomeTotal = incomeRows.reduce((sum, r) => sum + r.amount, 0);
  const advanceTotal = advanceRows.reduce((sum, r) => sum + r.amount, 0);
  const recoveredTotal = recoveredRows.reduce((sum, r) => sum + r.amount, 0);
  const outstandingAdvance = advanceTotal - recoveredTotal;
  const net = incomeTotal - expenseTotal;

  const byItem = {};
  for (const r of expenseRows) {
    if (!r.item) continue;
    byItem[r.item] = (byItem[r.item] || 0) + r.amount;
  }

  const top = Object.entries(byItem)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  const lines = [
    `📒 ${farm === "全部" ? "全部場" : farm}｜${title}`,
    "",
    `💸 支出合計：${expenseTotal.toLocaleString("zh-TW")} 元`,
    `💰 收入合計：${incomeTotal.toLocaleString("zh-TW")} 元`,
    `📊 收支差額：${net.toLocaleString("zh-TW")} 元`,
    `🤝 代墊合計：${advanceTotal.toLocaleString("zh-TW")} 元`,
    `↩️ 代墊收回：${recoveredTotal.toLocaleString("zh-TW")} 元`,
    `⏳ 尚未收回：${outstandingAdvance.toLocaleString("zh-TW")} 元`,
    `🧾 筆數：${filtered.length} 筆`,
  ];

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

    const rows = await readMonthRows(sheetId, monthSheet);
    return summaryText(
      rows,
      query.farm,
      `${query.year}年${query.month}月收支`
    );
  }

  if (query.mode === "今天") {
    await ensureMonthSheet(sheetId, now.monthSheet);
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
    "📒 雞場支出記帳",
    "",
    "【今天的帳】",
    "東平 支出 500 電風扇 / 三豐",
    "草湖 零 支出 2313 電費 / 台電",
    "仁愛 收入 5000 雞蛋銷售 / 客戶",
    "",
    "付款來源：未寫＝雞場帳戶；「零」＝零用金",
    "",
    "【代墊】",
    "仁愛 代墊 1500 油資 / 台塑",
    "仁愛 代墊 800 工具 / 五金行",
    "仁愛 代墊收回 1500",
    "",
    "【補登以前日期】",
    "9/28 東平 支出 4528 電費 / 台電",
    "2026/9/28 草湖 支出 500 電風扇 / 三豐",
    "",
    "【指定月份查詢】",
    "9月",
    "草湖 9月",
    "全部 9月",
    "2026/9",
    "東平 2026/9",
    "",
    "【其他查詢】",
    "今天",
    "本月",
    "今年",
    "東平 今天",
    "草湖 本月",
    "全部 本月",
  ].join("\\n");
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

  const query = parseQuery(text);
  if (query) {
    return await querySummary(targetSheetId, query);
  }

  const dated = parseOptionalDatePrefix(text);
  if (dated.error) {
    return `❌ ${dated.error}`;
  }

  const expense = parseExpenseCommand(dated.body);

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
      expense.advanceTarget ? `代墊對象：${expense.advanceTarget}` : null,
      expense.txType === "代墊" ? `代墊人：${userName}` : null,
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
  res.send("Chicken Farm Expense Bot - Multi User Test Version is running.");
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
  console.log(`Chicken Farm Expense Bot - Multi User Test Version running on port ${PORT}`);
});
