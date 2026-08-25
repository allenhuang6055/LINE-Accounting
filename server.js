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
const SHEET_NAME = process.env.SHEET_NAME || "02_LINE資料庫";

function getGoogleAuth() {
  const privateKey = (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n");

  return new google.auth.JWT({
    email: process.env.GOOGLE_CLIENT_EMAIL,
    key: privateKey,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
}

async function getSheets() {
  const auth = getGoogleAuth();
  return google.sheets({ version: "v4", auth });
}

function nowTaipei() {
  const now = new Date();

  const date = new Intl.DateTimeFormat("zh-TW", {
    timeZone: "Asia/Taipei",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(now).replace(/\//g, "-");

  const time = new Intl.DateTimeFormat("zh-TW", {
    timeZone: "Asia/Taipei",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  }).format(now);

  return { date, time };
}

function parseAmount(text) {
  if (!text) return null;

  const normalized = text
    .replace(/,/g, "")
    .replace(/元/g, "")
    .trim();

  const n = Number(normalized);
  return Number.isFinite(n) ? n : null;
}

function classifyExpense(item) {
  const s = (item || "").toLowerCase();

  const rules = [
    { keys: ["維修", "修理", "保養"], category: "維修費" },
    { keys: ["板手", "扳手", "螺絲", "工具"], category: "工具耗材" },
    { keys: ["油", "柴油", "汽油"], category: "油料費" },
    { keys: ["飼料"], category: "飼料" },
    { keys: ["藥", "疫苗"], category: "藥品" },
    { keys: ["水費"], category: "水費" },
    { keys: ["電費"], category: "電費" },
    { keys: ["電話", "網路"], category: "通訊費" },
    { keys: ["薪資", "薪水", "工資"], category: "薪資" },
    { keys: ["運費", "貨運", "清運"], category: "運輸費" },
  ];

  for (const rule of rules) {
    if (rule.keys.some(k => s.includes(k))) {
      return rule.category;
    }
  }

  return "其他支出";
}

function classifyIncome(item) {
  const s = (item || "").toLowerCase();

  if (s.includes("蛋")) return "銷貨收入";
  if (s.includes("租")) return "租金收入";
  if (s.includes("補助")) return "補助收入";

  return "其他收入";
}

async function getDisplayName(event) {
  try {
    if (event.source?.userId) {
      const profile = await client.getProfile(event.source.userId);
      return profile.displayName || event.source.userId;
    }
  } catch (err) {
    console.error("取得使用者名稱失敗：", err.message);
  }

  return "未知使用者";
}

async function appendAccountingRow({
  date,
  time,
  userName,
  type,
  category,
  item,
  amount,
  paymentMethod = "",
  note = "",
  status = "有效",
  messageId = "",
}) {
  const sheets = await getSheets();

  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A:K`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [[
        date,
        time,
        userName,
        type,
        category,
        item,
        amount,
        paymentMethod,
        note,
        status,
        messageId,
      ]],
    },
  });
}

async function readRows() {
  const sheets = await getSheets();

  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A2:K`,
  });

  return result.data.values || [];
}

function parseTaiwanDateString(s) {
  if (!s) return null;

  // 預期 YYYY-MM-DD
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;

  return {
    year: Number(m[1]),
    month: Number(m[2]),
    day: Number(m[3]),
  };
}

function getCurrentTaipeiParts() {
  const { date } = nowTaipei();
  return parseTaiwanDateString(date);
}

function rowToObj(row) {
  return {
    date: row[0] || "",
    time: row[1] || "",
    userName: row[2] || "",
    type: row[3] || "",
    category: row[4] || "",
    item: row[5] || "",
    amount: Number(String(row[6] || "0").replace(/,/g, "")) || 0,
    paymentMethod: row[7] || "",
    note: row[8] || "",
    status: row[9] || "",
    messageId: row[10] || "",
  };
}

function inRangeByMode(dateStr, mode) {
  const d = parseTaiwanDateString(dateStr);
  if (!d) return false;

  const current = getCurrentTaipeiParts();
  if (!current) return false;

  if (mode === "today") {
    return d.year === current.year &&
           d.month === current.month &&
           d.day === current.day;
  }

  if (mode === "month") {
    return d.year === current.year &&
           d.month === current.month;
  }

  if (mode === "year") {
    return d.year === current.year;
  }

  return false;
}

async function makeSummary(mode) {
  const rows = (await readRows())
    .map(rowToObj)
    .filter(r => r.status !== "刪除" && r.status !== "作廢")
    .filter(r => inRangeByMode(r.date, mode));

  let income = 0;
  let expense = 0;

  for (const row of rows) {
    if (row.type === "收入") income += row.amount;
    if (row.type === "支出") expense += row.amount;
  }

  const profit = income - expense;

  const title =
    mode === "today" ? "今日統計" :
    mode === "month" ? "本月統計" :
    "今年統計";

  return [
    title,
    "",
    `💰 收入：${income.toLocaleString("zh-TW")} 元`,
    `💸 支出：${expense.toLocaleString("zh-TW")} 元`,
    `${profit >= 0 ? "📈" : "📉"} 盈餘：${profit.toLocaleString("zh-TW")} 元`,
    "",
    `🧾 筆數：${rows.length} 筆`,
  ].join("\n");
}

function helpText() {
  return [
    "📒 LINE 記帳",
    "",
    "記帳方式：",
    "收入 5000 賣蛋",
    "支出 1200 山貓維修",
    "",
    "查詢：",
    "今天",
    "本月",
    "今年",
    "",
    "也可以輸入：",
    "說明",
    "幫助",
  ].join("\n");
}

function parseBookkeepingCommand(text) {
  const clean = (text || "").trim().replace(/\s+/g, " ");
  if (!clean) return null;

  // 格式：收入 5000 賣蛋
  // 格式：支出 1200 山貓維修
  const m = clean.match(/^(收入|支出)\s+([\d,]+(?:\.\d+)?)\s*(.*)$/);

  if (!m) return null;

  const type = m[1];
  const amount = parseAmount(m[2]);
  const item = (m[3] || "").trim() || "未填品項";

  if (amount === null || amount <= 0) {
    return { error: "金額格式錯誤" };
  }

  const category =
    type === "支出"
      ? classifyExpense(item)
      : classifyIncome(item);

  return {
    type,
    amount,
    item,
    category,
  };
}

async function handleTextMessage(event) {
  const text = event.message.text.trim();

  if (["說明", "幫助", "help", "HELP", "?"].includes(text)) {
    return helpText();
  }

  if (["今天", "今日", "今日統計"].includes(text)) {
    return await makeSummary("today");
  }

  if (["本月", "本月統計"].includes(text)) {
    return await makeSummary("month");
  }

  if (["今年", "年度", "今年統計"].includes(text)) {
    return await makeSummary("year");
  }

  const cmd = parseBookkeepingCommand(text);

  if (cmd?.error) {
    return `❌ ${cmd.error}\n\n例如：\n支出 1200 山貓維修`;
  }

  if (cmd) {
    const userName = await getDisplayName(event);
    const { date, time } = nowTaipei();

    await appendAccountingRow({
      date,
      time,
      userName,
      type: cmd.type,
      category: cmd.category,
      item: cmd.item,
      amount: cmd.amount,
      messageId: event.message.id || "",
    });

    return [
      `✅ ${cmd.type}記錄完成`,
      "",
      `品項：${cmd.item}`,
      `類別：${cmd.category}`,
      `金額：${cmd.amount.toLocaleString("zh-TW")} 元`,
      `填表人：${userName}`,
      `日期：${date}`,
    ].join("\n");
  }

  return [
    "我目前支援：",
    "",
    "收入 5000 賣蛋",
    "支出 1200 山貓維修",
    "",
    "查詢：今天／本月／今年",
    "",
    "輸入「說明」可看完整格式。",
  ].join("\n");
}

async function handleEvent(event) {
  if (event.type !== "message") return null;
  if (event.message.type !== "text") return null;

  try {
    const replyText = await handleTextMessage(event);

    if (!replyText) return null;

    await client.replyMessage({
      replyToken: event.replyToken,
      messages: [{
        type: "text",
        text: replyText,
      }],
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
  res.send("LINE Accounting Bot V1 is running.");
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
  console.log(`LINE Accounting Bot V1 running on port ${PORT}`);
});
