# App Runner 4 步驟 30 分鐘上線

這份是給「會持續更新迭代」的部署流程，重點是 push 程式就能自動部署。

## 先決條件

1. 專案已推到 GitHub（建議 main 分支）。
2. AWS 帳號可用，區域建議固定在 ap-northeast-1（東京）或你主要客群附近。

---

## 步驟 1：建立 App Runner 服務

1. 進 AWS Console -> App Runner -> Create service。
2. Source 選擇 Source code repository。
3. 連接 GitHub，選你的 repo 與 branch（例如 main）。
4. Deployment settings 選 Automatic（每次 push 自動部署）。

---

## 步驟 2：使用 apprunner.yaml

你專案已新增 [apprunner.yaml](apprunner.yaml)。

建立服務時選擇「Configuration file（apprunner.yaml）」，讓 App Runner 用這份設定：

1. Build：npm ci
2. Run：npm start
3. Port：3000
4. Env：NODE_ENV=production

---

## 步驟 3：設定環境變數與健康檢查

在 App Runner Service settings 補上：

1. Environment variable：
   - NODE_ENV=production
   - UPLOAD_DIR=./uploads
2. Health check path：
   - /api/health

部署完成後，先測：

1. https://你的服務網址/api/health
2. https://你的服務網址/

---

## 步驟 4：把資料持久化（很重要）

你目前程式資料放在本機檔案（data 與 uploads），App Runner 屬於容器執行環境，檔案系統不是長期資料庫。

請至少做下面兩件：

1. 訂單/會員/商品資料搬到資料庫（RDS PostgreSQL 或 DynamoDB）。
2. 上傳檔案搬到 S3（不要留在本機 uploads）。

否則服務重建、擴縮或重新部署時，可能出現資料遺失或資料不一致。

---

## 更新流程（之後每天會用到）

1. 本機改程式。
2. git add .
3. git commit -m "feat: ..."
4. git push origin main
5. App Runner 自動部署完成。

---

## 失敗排查（最常見）

1. Build fail：確認 package-lock.json 存在，並且 npm ci 可在本機成功。
2. Health check fail：確認 server 有綁在 process.env.PORT（你目前有，預設 3000）。
3. 上傳或資料突然不見：這是本機檔案儲存造成，改 S3 + DB 才是正解。

---

## 成本與是否適合持續迭代

適合你現在狀況，因為：

1. 不用自己維護 EC2/Nginx/PM2。
2. 每次 push 自動部署。
3. 對小到中型後台系統，維運負擔比 EC2 低很多。

不過如果流量變大、需要更細緻網路與成本控制，再評估 ECS Fargate。
