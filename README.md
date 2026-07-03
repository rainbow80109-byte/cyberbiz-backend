# 訂單管理後台系統

一個類似 CyberBiz 的訂單管理系統，支持 Excel 導入、數據統計分析。

## 項目結構

```
後台/
├── server.js              # 後端主文件
├── package.json           # 依賴配置
├── .env                   # 環境變量
├── data/                  # 數據庫存儲
├── uploads/               # 臨時上傳文件
└── public/                # 前端 build 文件（待創建）
```

## 技術棧

- **後端**: Node.js + Express
- **數據庫**: SQLite (better-sqlite3)
- **文件上傳**: Multer
- **Excel 處理**: XLSX
- **前端**: Vue.js (待創建)

## 快速開始

### 安裝依賴
```bash
npm install
```

### 運行開發服務器
```bash
npm run dev
```

或直接運行：
```bash
npm start
```

服務器將運行於 `http://localhost:3000`

## API 端點

### 1. 導入 Excel 訂單
**POST** `/api/import`
- Content-Type: `multipart/form-data`
- 文件字段: `file` (Excel 文件)
- Excel 列名支持: 發票、金額、名稱、日期、備註

**響應示例:**
```json
{
  "success": true,
  "successCount": 100,
  "errorCount": 2,
  "errors": ["第 5 行: 金額無效"],
  "message": "成功導入 100 條訂單，失敗 2 條"
}
```

### 2. 獲取訂單列表
**GET** `/api/orders`

**查詢參數:**
- `page` (默認: 1) - 頁碼
- `pageSize` (默認: 20) - 每頁數量
- `search` - 搜索關鍵詞 (發票、名稱、備註)
- `sortBy` (默認: created_at) - 排序字段
- `sortOrder` (默認: DESC) - 排序順序 (ASC/DESC)

**響應示例:**
```json
{
  "data": [
    {
      "id": 1,
      "invoice": "INV-001",
      "amount": 1500.50,
      "name": "張三",
      "date": "2026-06-02",
      "notes": "首單",
      "created_at": "2026-06-02T10:00:00Z"
    }
  ],
  "pagination": {
    "page": 1,
    "pageSize": 20,
    "total": 150,
    "pages": 8
  }
}
```

### 3. 獲取統計數據
**GET** `/api/statistics`

**響應示例:**
```json
{
  "totalOrders": 150,
  "totalAmount": "225750.50",
  "avgAmount": "1505.00",
  "trend": [
    {
      "date": "2026-06-02",
      "count": 12,
      "total": 18500.50
    }
  ],
  "topCustomers": [
    {
      "name": "張三",
      "count": 25,
      "total": 35000.00
    }
  ]
}
```

### 4. 刪除訂單
**DELETE** `/api/orders/:id`

**響應示例:**
```json
{
  "success": true,
  "message": "訂單已刪除"
}
```

### 5. 健康檢查
**GET** `/api/health`

**響應:**
```json
{
  "status": "ok",
  "message": "Server is running"
}
```

## Excel 導入格式

Excel 文件應包含以下列 (支持中文或英文列名):

| 發票 | 金額 | 名稱 | 日期 | 備註 |
|------|------|------|------|------|
| INV-001 | 1500.50 | 張三 | 2026-06-02 | 首單 |
| INV-002 | 2000.00 | 李四 | 2026-06-02 | 續購 |

**必填字段**: 發票、金額、名稱、日期

## 數據庫架構

### orders 表
```sql
CREATE TABLE orders (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  invoice TEXT NOT NULL,              -- 發票號
  amount REAL NOT NULL,               -- 金額
  name TEXT NOT NULL,                 -- 客戶名稱
  date TEXT NOT NULL,                 -- 訂單日期
  notes TEXT,                         -- 備註
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP,  -- 創建時間
  updated_at DATETIME DEFAULT CURRENT_TIMESTAMP   -- 更新時間
);
```

## 環境變量

在 `.env` 文件中配置:

```env
PORT=3000                      # 服務器端口
NODE_ENV=development           # 環境 (development/production)
DB_PATH=./data/orders.db       # 數據庫文件路徑
UPLOAD_DIR=./uploads           # 上傳文件臨時目錄
```

## 開發中的功能

- [ ] 前端 Vue.js 頁面
- [ ] 批量導出訂單 (CSV/Excel)
- [ ] 用戶認證和授權
- [ ] 數據備份和恢復
- [ ] 圖表分析和報表

## 許可證

ISC
