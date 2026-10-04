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

const SHEET_ID = process.env.GOOGLE_SHEET_ID;

// ===== 場別 =====
const FARM_ALIASES = {
  "東平": "東平場",
  "東平場": "東平場",
  "草湖": "草湖場",
  "草湖場": "草湖場",
  "潭墘": "潭墘場",
  "潭墘場": "潭墘場",
  "後寮": "後寮場",
  "後寮場": "後寮場",
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

async function ensureSheet(sheetName, headers) {
  const sheets = await getSheets();

  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SHEET_ID,
    fields: "sheets.properties(sheetId,title)",
  });

  const exists = (meta.data.sheets || []).some(s => s.properties.title === sheetName);

  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
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
      spreadsheetId: SHEET_ID,
      range: `'${sheetName}'!A1:${columnLetter(headers.length)}1`,
    });

    const firstRow = current.data.values?.[0] || [];

    if (firstRow.join("|") !== headers.join("|")) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
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

async function ensureBaseStructure() {
  // 資料分頁
  await ensureSheet("資料", ["品項", "科目代號", "科目分類"]);

  const sheets = await getSheets();
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: "'資料'!A2:C",
  });

  if (!(result.data.values || []).length) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `'資料'!A2:C${ACCOUNT_ITEMS.length + 1}`,
      valueInputOption: "RAW",
      requestBody: { values: ACCOUNT_ITEMS },
    });
  }
}

async function ensureMonthSheet(monthSheet) {
  await ensureSheet(monthSheet, [
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
  ]);
}

function parseExpenseCommand(text) {
  const clean = String(text || "").trim().replace(/\s+/g, " ");
  const parts = clean.split(" ");

  // 格式一：東平 支出 500 電風扇
  let farm = farmFromToken(parts[0]);
  let rest;

  if (farm && parts[1] === "支出") {
    rest = parts.slice(2).join(" ");
  } else if (parts[0] === "支出") {
    // 格式二：支出 東平 500 電風扇
    farm = farmFromToken(parts[1]);
    if (!farm) return { error: "請先輸入場別，例如：東平 支出 500 電風扇" };
    rest = parts.slice(2).join(" ");
  } else {
    return null;
  }

  if (farm === "全部") {
    return { error: "記帳時不能使用「全部」，請指定東平、草湖、潭墘或後寮。" };
  }

  // 支援 2x600
  let m = rest.match(/^(\d+(?:\.\d+)?)\s*[xX×*]\s*([\d,]+(?:\.\d+)?)\s+(.+)$/);

  let qty = 1;
  let unitPrice;
  let amount;
  let description;

  if (m) {
    qty = Number(m[1]);
    unitPrice = parseAmount(m[2]);
    amount = qty * unitPrice;
    description = m[3].trim();
  } else {
    m = rest.match(/^([\d,]+(?:\.\d+)?)\s+(.+)$/);
    if (!m) return { error: "格式例如：東平 支出 500 電風扇" };

    amount = parseAmount(m[1]);
    unitPrice = amount;
    description = m[2].trim();
  }

  if (!Number.isFinite(amount) || amount <= 0) {
    return { error: "金額格式不正確" };
  }

  let vendor = "";

  // 方式 1：原本格式仍支援
  // 草湖 支出 500 電風扇 廠商:振豐五金
  const vendorMatch = description.match(/\s+廠商[:：]\s*(.+)$/);

  if (vendorMatch) {
    vendor = vendorMatch[1].trim();
    description = description.slice(0, vendorMatch.index).trim();
  } else {
    // 方式 2：簡易格式
    // 草湖 支出 500 電風扇 振豐五金
    // 9/28 草湖 支出 4528 電費 台電
    //
    // 為避免把「山貓 維修」之類用途誤認成廠商，
    // 只有最後一段看起來像廠商名稱時才自動帶入。
    const words = description.split(/\s+/).filter(Boolean);

    if (words.length >= 2) {
      const last = words[words.length - 1];

      const looksLikeVendor =
        /(公司|企業|商行|五金|水電|工程|電機|材料|行|店|廠|中心|台電|中油|農會|合作社|藥局|診所|醫院|牧場|蛋行|農產|實業|股份|有限公司|工作室|加油站)$/.test(last);

      if (looksLikeVendor) {
        vendor = last;
        description = words.slice(0, -1).join(" ").trim();
      }
    }
  }

  let selected = null;
  let purpose = description;

  // 可明確指定 Excel 品項：
  // 東平 支出 500 雜項費用 電風扇
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
  };
}

async function writeExpense(monthSheet, expense, dateText) {
  const sheets = await getSheets();

  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `'${monthSheet}'!A:K`,
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
      ]],
    },
  });
}

async function readMonthRows(monthSheet) {
  const sheets = await getSheets();

  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `'${monthSheet}'!A2:K`,
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

  // 「今天」「本月」「今年」= 全部場
  if (["今天", "本月", "今年"].includes(clean)) {
    return { farm: "全部", mode: clean };
  }

  // 東平 今天 / 草湖 本月 / 全部 今年
  if (parts.length === 2) {
    const farm = farmFromToken(parts[0]);
    if (farm && ["今天", "本月", "今年"].includes(parts[1])) {
      return { farm, mode: parts[1] };
    }
  }

  return null;
}

function summaryText(rows, farm, title) {
  const filtered = farm === "全部" ? rows : rows.filter(r => r.farm === farm);
  const total = filtered.reduce((sum, r) => sum + r.amount, 0);

  const byItem = {};
  for (const r of filtered) {
    if (!r.item) continue;
    byItem[r.item] = (byItem[r.item] || 0) + r.amount;
  }

  const top = Object.entries(byItem)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  const lines = [
    `📒 ${farm === "全部" ? "全部場" : farm}｜${title}`,
    "",
    `💸 支出合計：${total.toLocaleString("zh-TW")} 元`,
    `🧾 筆數：${filtered.length} 筆`,
  ];

  if (top.length) {
    lines.push("", "分類：");
    for (const [name, value] of top) {
      lines.push(`・${name}：${value.toLocaleString("zh-TW")} 元`);
    }
  }

  return lines.join("\n");
}

async function querySummary(query) {
  const now = taipeiNow();

  if (query.mode === "今天") {
    await ensureMonthSheet(now.monthSheet);
    const rows = await readMonthRows(now.monthSheet);

    const todayRows = rows.filter(r => {
      const d = parseDateText(r.date);
      return d &&
        d.year === Number(now.year) &&
        d.month === Number(now.month) &&
        d.day === Number(now.day);
    });

    return summaryText(todayRows, query.farm, "今日支出");
  }

  if (query.mode === "本月") {
    await ensureMonthSheet(now.monthSheet);
    const rows = await readMonthRows(now.monthSheet);
    return summaryText(rows, query.farm, `${Number(now.month)}月支出`);
  }

  // 今年
  const sheets = await getSheets();
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SHEET_ID,
    fields: "sheets.properties.title",
  });

  const monthSheets = (meta.data.sheets || [])
    .map(s => s.properties.title)
    .filter(title => new RegExp(`^${now.year}\\d{2}月$`).test(title))
    .sort();

  let rows = [];

  for (const monthSheet of monthSheets) {
    rows = rows.concat(await readMonthRows(monthSheet));
  }

  return summaryText(rows, query.farm, `${now.year}年支出`);
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
    "東平 支出 500 電風扇",
    "草湖 支出 2313 電費",
    "",
    "【補登以前日期】",
    "9/28 東平 支出 4528 電費",
    "2026/9/28 草湖 支出 500 電風扇",
    "",
    "【數量 × 單價】",
    "9/20 草湖 支出 2x600 手推車輪胎",
    "",
    "【直接輸入廠商】",
    "草湖 支出 500 電風扇 振豐五金",
    "9/28 草湖 支出 4528 電費 台電",
    "",
    "原本格式也能用：",
    "草湖 支出 500 電風扇 廠商:振豐五金",
    "",
    "也可指定 Excel 品項：",
    "東平 支出 500 雜項費用 電風扇",
    "",
    "【查詢】",
    "今天",
    "本月",
    "今年",
    "東平 今天",
    "草湖 本月",
    "潭墘 今年",
    "全部 本月",
  ].join("\n");
}

async function handleTextMessage(event) {
  const text = String(event.message.text || "").trim();

  if (["說明", "幫助", "help", "HELP", "?"].includes(text)) {
    return helpText();
  }

  await ensureBaseStructure();

  const query = parseQuery(text);
  if (query) {
    return await querySummary(query);
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
    await ensureMonthSheet(dated.dateInfo.monthSheet);
    await writeExpense(dated.dateInfo.monthSheet, expense, dated.dateInfo.dateText);

    const userName = await getProfileName(event);

    return [
      "✅ 支出記錄完成",
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
      `科目代號：${expense.code}`,
      `科目分類：${expense.className}`,
      `填表人：${userName}`,
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
  res.send("Chicken Farm Expense Bot - Backdate Version is running.");
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
  console.log(`Chicken Farm Expense Bot - Backdate Version running on port ${PORT}`);
});
