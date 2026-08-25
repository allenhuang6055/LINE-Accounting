# LINE 記帳機器人 V1

第一版功能：

- LINE 輸入 `收入 5000 賣蛋`
- LINE 輸入 `支出 1200 山貓維修`
- 自動分類基本支出／收入
- 寫入 Google 試算表
- 查詢 `今天`
- 查詢 `本月`
- 查詢 `今年`

---

## 1. Google 試算表

建立一個分頁：

`02_LINE資料庫`

第一列欄位請依序放：

| A | B | C | D | E | F | G | H | I | J | K |
|---|---|---|---|---|---|---|---|---|---|---|
| 日期 | 時間 | 填表人 | 類型 | 類別 | 品項 | 金額 | 付款方式 | 備註 | 狀態 | LINE訊息ID |

---

## 2. Google Service Account

Google Cloud 建立 Service Account 後：

1. 啟用 Google Sheets API
2. 建立 Service Account
3. 建立 JSON Key
4. 取得：
   - client_email
   - private_key
5. 把 Google 試算表分享給 Service Account 的 email，權限設「編輯者」

---

## 3. LINE Developers

建立 Messaging API Channel，取得：

- Channel access token
- Channel secret

Webhook URL：

`https://你的網址/webhook`

並開啟：

- Use webhook

---

## 4. 環境變數

將 `.env.example` 複製成 `.env`

填入：

- LINE_CHANNEL_ACCESS_TOKEN
- LINE_CHANNEL_SECRET
- GOOGLE_SHEET_ID
- GOOGLE_CLIENT_EMAIL
- GOOGLE_PRIVATE_KEY

---

## 5. 本機啟動

```bash
npm install
npm start
```

看到：

```text
LINE Accounting Bot V1 running on port 3000
```

代表程式已啟動。

---

## 6. LINE 測試

輸入：

```text
支出 1200 山貓維修
```

預期：

```text
✅ 支出記錄完成

品項：山貓維修
類別：維修費
金額：1,200 元
填表人：你的 LINE 名稱
日期：2026-08-20
```

再輸入：

```text
收入 5000 賣蛋
```

以及：

```text
今天
本月
今年
```

---

## 下一版建議

V1.1 可加入：

- 收款
- 代收款
- 自訂類別表
- 依關鍵字自動分類
- 8月／7月等指定月份
- 日期區間查詢
- 刪除上一筆
- 權限管理
- Flex Message 選單
