# 班級 LINE Bot 回覆小幫手 (ClassLineBot)

本專案為結合 **LINE Messaging API**、**Google Apps Script (GAS)**、**Google Sheets** 與 **Google Gemini AI** 的智慧班級聯絡與回覆小幫手機器人。提供家長學生姓名自動綁定、文字與語音訊息辨識、常見問答 (QA) 自動回覆、學校與班級行事曆同步、以及學校最新公告自動擷取等功能。

---

## 📌 雲端專案資源

| 項目 | ID / 連結 | 說明 |
| :--- | :--- | :--- |
| **Google 試算表** | 1HH9KwUUtMYqBs8JVdWQC-1boHsuRzGQYF5mT1A7xwNc | [開啟 Google 試算表](https://docs.google.com/spreadsheets/d/1HH9KwUUtMYqBs8JVdWQC-1boHsuRzGQYF5mT1A7xwNc/edit) |
| **Google Apps Script** | 1BxuxNQMgRHFaZ3NNoW1XovwMUQDAFFMrVNtk9Y_leVvokG6-nAojY3U6 | [開啟 Apps Script 專案](https://script.google.com/d/1BxuxNQMgRHFaZ3NNoW1XovwMUQDAFFMrVNtk9Y_leVvokG6-nAojY3U6/edit) |
| **GitHub 儲存庫** | ClassLineBot | [huahua718-creator/ClassLineBot](https://github.com/huahua718-creator/ClassLineBot) |

---

## 📂 專案檔案結構

`	ext
ClassLineBot/
├── AppsScript/                     # Google Apps Script 原始碼與同步工具
│   ├── .clasp.json                 # Google Clasp 設定檔（已綁定 Script ID）
│   ├── appsscript.json             # GAS 專案資訊清單（時區 Asia/Taipei 與執行環境）
│   ├── 程式碼.js                    # LINE Webhook、Gemini AI、行事曆與公告核心邏輯
│   ├── pull.bat                    # Windows 一鍵從 GAS 拉取最新代碼
│   └── push.bat                    # Windows 一鍵將代碼推送到 GAS
├── GoogleSheets/                   # Google 試算表資料備份與分頁導出
│   ├── Line_Bot_回覆小幫手.xlsx    # 完整 Excel 試算表（含 5 個工作表）
│   ├── QA.csv                      # 常見問答庫（關鍵字與參考回應）
│   ├── 使用者對照.csv               # LINE 使用者 ID、暱稱、家長學生姓名對照表
│   ├── 行事曆Cache.csv              # 學校與班級行事曆快取紀錄
│   ├── 學校公告.csv                 # 學校官網最新公告快取
│   ├── 系統設定.csv                 # 系統參數設定檔
│   └── 系統設定_範本.csv            # 系統設定檔範本（供新環境部署使用）
├── .gitignore                      # Git 忽略檔案設定
└── README.md                       # 專案說明文件
`

---

## 🚀 核心功能特色

1. **家長與學生自動對照綁定**
   - 家長加入 LINE 好友時，系統自動發送親切歡迎詞並引導提供學生姓名。
   - 支援輸入「學生：王小明」、「孩子：王小明」自動記錄至「使用者對照」分頁。

2. **智慧訊息辨識 (Gemini AI)**
   - 支援文字訊息與**語音訊息辨識**，長輩與忙碌家長可直接使用語音詢問。
   - 結合 Google Gemini 2.5 Flash / Flash-Lite 動態最新模型，提供具同理心且溫暖的導師級回覆。

3. **關鍵字 QA 自動回覆與線上動態擴充**
   - 優先比對「QA」分頁關鍵字，提供成績查詢、大屏系統、處室分機、校慶補假等資訊。
   - 導師可直接在 LINE 對話中輸入指令擴充資料庫：
     `	ext
     新增QA 聯絡簿 聯絡簿請於每日 18:00 前簽名確認
     `

4. **行事曆同步與快速新增行程**
   - 整合學校 Google 日曆與教師個人主日曆，即時查詢本週與近期重要日程。
   - 導師可直接在 LINE 中新增行程：
     `	ext
     新增行程 12/7 09:00 校慶大會
     `

5. **學校官網公告即時擷取**
   - 自動解析學校公告網頁，定時快取並過濾重要訊息，方便家長與師生隨時掌握校園動態。

---

## 🛠️ 開發與同步方式 (使用 Clasp)

本專案已配置 @google/clasp，可在本地端使用 VS Code 等編輯器修改代碼後直接同步至 Google Apps Script：

### 1. 取得最新程式碼
`ash
cd AppsScript
npx @google/clasp pull
# 或在 Windows 直接雙擊執行 pull.bat
`

### 2. 推送本地修改至 Google Apps Script
`ash
cd AppsScript
npx @google/clasp push -f
# 或在 Windows 直接雙擊執行 push.bat
`

---

## ⚙️ 系統設定參數說明

於 Google 試算表的「系統設定」分頁中設置：

| 設定項目 | 說明 | 範例 |
| :--- | :--- | :--- |
| LINE_ACCESS_TOKEN | LINE Messaging API Channel Access Token | LUpC... |
| ADMIN_USER_ID | 導師管理員個人 LINE User ID | U1d1... |
| GEMINI_API_KEY | Google Gemini API Key | AIzaSy... |
| TEACHER_EMAIL | 教師 Google 個人日曆（填 primary 即可） | primary |
| SCHOOL_CALENDAR_ID | 學校公開 Google 日曆 ID | clshlib7@gmail.com |
| SCHOOL_ANNOUNCEMENT_URL | 學校官方公告網頁 URL | https://www.clsh.ntpc.edu.tw/... |
| CLASS_NAME | 班級名稱 | 國二愛 |

---

## 📄 授權條款

本專案採 MIT 授權條款，歡迎自由使用與擴充。