# 🚀 快速開始指南

## 第 1 步：安裝依賴

在 PowerShell 或 CMD 中運行：

```bash
cd "C:\Users\ㄋ\Desktop\後台"
npm install
```

**預期時間**: 2-5 分鐘

---

## 第 2 步：啟動後端服務器

```bash
npm start
```

或使用開發模式（自動重啟）：

```bash
npm run dev
```

**預期輸出**:
```
✅ 後端服務器運行於 http://localhost:3000
📊 數據庫路徑: ./data/orders.db
```

---

## 第 3 步：打開前端頁面

在瀏覽器打開：

```
file:///C:/Users/ㄋ/Desktop/後台/index.html
```

或直接雙擊 `index.html` 文件

---

## 第 4 步：導入測試數據

### 方式 A：生成測試 Excel（推薦）

```bash
node generate-test-data.js
```

這會生成 `test-data/sample-orders.xlsx` 文件

然後在前端頁面：
1. 點擊左側菜單的 **📥 匯入訂單**
2. 選擇 `test-data/sample-orders.xlsx`
3. 點擊 **開始匯入**

### 方式 B：使用 curl 測試（需要 curl 工具）

```bash
curl -X POST http://localhost:3000/api/health
```

---

## 📊 功能測試清單

### ✅ 儀表板 (Dashboard)
- [ ] 查看訂單統計卡片（總數、總金額、平均金額）
- [ ] 查看訂單趨勢圖表
- [ ] 查看客戶排名表格

### ✅ 匯入訂單 (Import)
- [ ] 上傳 Excel 文件
- [ ] 查看導入結果
- [ ] 驗證數據完整性

### ✅ 訂單列表 (Orders)
- [ ] 查看訂單列表
- [ ] 搜索訂單（發票、名稱、備註）
- [ ] 分頁瀏覽
- [ ] 刪除訂單

---

## 🧪 API 端點測試

### 1. 健康檢查
```bash
curl http://localhost:3000/api/health
```

### 2. 獲取統計數據
```bash
curl http://localhost:3000/api/statistics
```

### 3. 獲取訂單列表
```bash
curl http://localhost:3000/api/orders?page=1&pageSize=10
```

### 4. 搜索訂單
```bash
curl "http://localhost:3000/api/orders?search=張三"
```

---

## 📝 資料夾結構

```
後台/
├── server.js              # 後端主程序 ⭐
├── index.html             # 前端頁面 ⭐
├── package.json           # npm 依賴配置
├── .env                   # 環境變量
├── start.bat              # 快速啟動腳本
├── test-api.js            # API 測試工具
├── generate-test-data.js  # 生成測試數據
├── README.md              # 項目文檔
├── data/                  # 📁 SQLite 數據庫存儲
│   └── orders.db          # 自動生成
├── uploads/               # 📁 Excel 上傳臨時目錄
│   └── (自動生成)
└── test-data/             # 📁 測試數據
    └── sample-orders.xlsx # 生成的測試 Excel
```

---

## ⚙️ 配置說明

編輯 `.env` 文件可修改：

```env
PORT=3000                      # 服務器端口
NODE_ENV=development           # 運行環境
DB_PATH=./data/orders.db       # 數據庫位置
UPLOAD_DIR=./uploads           # 上傳文件目錄
```

---

## 🐛 常見問題

### Q: 啟動失敗 "Module not found"
**A**: 運行 `npm install` 重新安裝依賴

### Q: 前端無法連接到後端
**A**: 確保後端服務器運行在 `http://localhost:3000`

### Q: Excel 導入失敗
**A**: 確保 Excel 包含必填列：發票、金額、名稱、日期

### Q: 數據不顯示
**A**: 檢查瀏覽器控制台（F12）查看錯誤信息

---

## 📚 相關文件

- `server.js` - API 實現和數據庫邏輯
- `index.html` - 所有前端頁面和功能
- `generate-test-data.js` - 生成測試數據
- `README.md` - API 文檔

---

## 💡 下一步優化

- [ ] 👤 添加用戶認證登錄
- [ ] 📊 實現更多圖表（餅圖、柱狀圖）
- [ ] 📥 批量導出訂單為 Excel
- [ ] 🗄️ 數據備份和恢復功能
- [ ] 🌐 部署到雲服務（Vercel/Heroku）

---

## ✨ 完成！

現在你可以開始使用訂單管理系統了！ 🎉

有問題嗎? 查看 `README.md` 或檢查瀏覽器控制台的錯誤信息。
