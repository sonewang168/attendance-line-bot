# 📚 學生簽到系統

LINE BOT + GPS 定位 + Google Sheets 雲端簽到系統

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/YOUR_USERNAME/attendance-line-bot)

---

## ✨ 功能特色

| 功能 | 說明 |
|------|------|
| 📱 LINE BOT 簽到 | 學生用 LINE 掃碼即可簽到 |
| 📍 GPS 定位驗證 | 必須在教室範圍內才能簽到 |
| ⏰ 自動判定遲到 | 超過設定時間自動標記 |
| ❌ 缺席自動追蹤 | 課程結束自動標記未簽到學生 |
| 🔔 即時通知 | LINE 推送簽到結果 |
| 📊 Google Sheets | 雲端資料儲存與統計 |
| 🏫 多班級管理 | 支援多班級多課程 |

---

## 🚀 快速部署

### 步驟 1：Fork 此專案

點擊右上角 `Fork` 按鈕

### 步驟 2：建立 LINE Bot

1. 前往 [LINE Developers](https://developers.line.biz/)
2. 建立 Messaging API Channel
3. 記下 `Channel Secret` 和 `Channel Access Token`

### 步驟 3：設定 Google Sheets

1. 建立新的 [Google Sheets](https://sheets.google.com/)
2. 前往 [Google Cloud Console](https://console.cloud.google.com/)
3. 建立專案並啟用 Google Sheets API
4. 建立服務帳號並下載 JSON 金鑰
5. 將服務帳號 Email 加入試算表共用（編輯權限）

### 步驟 4：部署到 Render

1. 前往 [Render](https://render.com/)
2. New → Web Service → 連結 GitHub
3. 選擇此 Repository
4. 設定環境變數（見下方）
5. Deploy！

### 步驟 5：設定 Webhook

回到 LINE Developers，設定 Webhook URL：
```
https://你的render網址.onrender.com/webhook
```

---

## ⚙️ 環境變數

| 變數名稱 | 說明 |
|----------|------|
| `LINE_CHANNEL_ACCESS_TOKEN` | LINE Bot Access Token |
| `LINE_CHANNEL_SECRET` | LINE Bot Channel Secret |
| `GOOGLE_SHEET_ID` | Google Sheets ID |
| `GOOGLE_SERVICE_ACCOUNT_EMAIL` | 服務帳號 Email |
| `GOOGLE_PRIVATE_KEY` | 服務帳號私鑰 |

---

## 📱 學生使用方式

### 首次使用
1. 掃描 LINE Bot QR Code 加好友
2. 輸入「註冊」
3. 依序輸入：學號 → 姓名 → 班級

### 簽到流程
1. 掃描教師的簽到 QR Code
2. 點擊「分享位置簽到」
3. 收到結果：✅ 已報到 / ⚠️ 遲到 / 🚫 位置錯誤

### 指令列表
- `註冊` - 綁定學號
- `我的資料` - 查看個人資訊
- `出席紀錄` - 查看簽到記錄
- `說明` - 顯示使用說明

---

## 📊 Google Sheets 結構

系統會自動建立以下工作表：

- **學生名單** - 學號、姓名、班級、LINE ID
- **班級列表** - 班級代碼、名稱、導師
- **課程列表** - 課程資訊、GPS 座標、簽到範圍
- **簽到活動** - 每日簽到活動
- **簽到紀錄** - 所有簽到記錄
- **出席統計** - 自動統計出席率

---

## 🔧 本地開發

```bash
# 安裝相依套件
npm install

# 複製環境變數
cp .env.example .env

# 編輯 .env 填入你的設定

# 啟動開發伺服器
npm run dev
```

---

## 📁 專案結構

```
attendance-line-bot/
├── server.js           # 主程式
├── package.json        # 相依套件
├── .env.example        # 環境變數範本
├── render.yaml         # Render 部署設定
├── railway.json        # Railway 部署設定
├── public/
│   └── index.html      # 教師管理介面
└── README.md           # 說明文件
```

---

## 📄 授權

MIT License

---

## 🙋 常見問題

**Q: 學生無法分享位置？**
A: 確認 LINE 已開啟定位權限

**Q: GPS 不準確？**
A: 建議簽到範圍設 50 公尺以上

**Q: 如何補簽到？**
A: 直接在 Google Sheets 新增紀錄

---

Made with ❤️ for Teachers
