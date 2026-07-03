# AWS 部署指南（Node.js 後台）

本專案目前是 `Node.js + Express`，資料寫在本機檔案（`data/*.json`、`uploads/`）。

建議先用 **EC2 + PM2** 上線（最直覺、可控性高），之後再升級到 RDS/S3。

---

## 方案 A（推薦）：EC2 + Nginx + PM2

### 1. 建立 EC2

1. 到 AWS Console 建立 EC2（Ubuntu 24.04 LTS）。
2. Security Group 至少開放：
   - `22`（SSH，只開你自己的 IP）
   - `80`（HTTP）
   - `443`（HTTPS，之後憑證會用）
3. 建立或下載 `.pem` 金鑰。

---

### 2. 連線到 EC2（Windows PowerShell）

```powershell
ssh -i "C:\path\to\your-key.pem" ubuntu@<EC2_PUBLIC_IP>
```

---

### 3. 在 EC2 安裝 Node、Nginx、PM2

```bash
sudo apt update
sudo apt install -y nginx git curl
curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -
sudo apt install -y nodejs
sudo npm install -g pm2
```

檢查版本：

```bash
node -v
npm -v
pm2 -v
```

---

### 4. 上傳專案到 EC2

你可以用 `git clone`（推薦）或 `scp`。

#### 方式 1：git clone

```bash
git clone <你的repo網址> app
cd app
```

#### 方式 2：scp（如果你沒有放 GitHub）

在本機 PowerShell：

```powershell
scp -i "C:\path\to\your-key.pem" -r "C:\Users\user\Desktop\後台\*" ubuntu@<EC2_PUBLIC_IP>:/home/ubuntu/app/
```

再 SSH 進去：

```bash
cd /home/ubuntu/app
```

---

### 5. 安裝依賴並設定環境變數

```bash
npm ci
```

建立 `.env`：

```bash
cat > .env << 'EOF'
PORT=3000
NODE_ENV=production
UPLOAD_DIR=./uploads
EOF
```

> 備註：你程式目前讀寫 `data/*.json`，會自動建立資料夾。

---

### 6. 用 PM2 啟動

```bash
pm2 start server.js --name order-backend
pm2 save
pm2 startup systemd -u ubuntu --hp /home/ubuntu
```

依畫面提示再貼一次 `sudo ...` 指令完成開機自啟。

檢查：

```bash
pm2 status
pm2 logs order-backend --lines 100
curl http://127.0.0.1:3000/api/health
```

---

### 7. 設定 Nginx 反向代理

```bash
sudo tee /etc/nginx/sites-available/order-backend > /dev/null << 'EOF'
server {
    listen 80;
    server_name _;

    client_max_body_size 20M;

    location / {
        proxy_pass http://127.0.0.1:3000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_cache_bypass $http_upgrade;
    }
}
EOF

sudo ln -s /etc/nginx/sites-available/order-backend /etc/nginx/sites-enabled/
sudo rm -f /etc/nginx/sites-enabled/default
sudo nginx -t
sudo systemctl restart nginx
```

打開：

- `http://<EC2_PUBLIC_IP>/api/health`

---

### 8. 後續更新流程

```bash
cd /home/ubuntu/app
git pull
npm ci
pm2 restart order-backend
```

---

## 方案 B：Elastic Beanstalk（更省維運）

如果你想少碰伺服器設定，可用 EB：

1. AWS Console 建立 Elastic Beanstalk（Node.js 平台）。
2. 專案根目錄打包 zip（排除 `node_modules`）。
3. 上傳 zip 到 EB。
4. 在 EB 環境變數設定：
   - `NODE_ENV=production`
   - `PORT=8081`（可選，若平台指定）
   - `UPLOAD_DIR=./uploads`
5. 部署後打開 EB URL 測試 `/api/health`。

> 注意：EB 多數情況為短暫磁碟，重建後本機檔案可能遺失。

---

## 資料持久化（你這個專案很重要）

你目前資料在本機檔案：

- `data/orders.json`
- `data/customer_feedback.json`
- `data/products.json`
- `data/product_categories.json`
- `uploads/`

在雲端要避免資料遺失，建議至少做一項：

1. 改用 RDS（MySQL/PostgreSQL）存訂單。
2. 上傳檔案改存 S3。
3. EC2 先做每日備份（cron + S3）。

---

## 最快可上線清單（10 分鐘版）

1. 建 EC2（開 22/80）。
2. SSH 進去安裝 Node + PM2 + Nginx。
3. 上傳專案 + `npm ci`。
4. 建 `.env`，`NODE_ENV=production`。
5. `pm2 start server.js --name order-backend`。
6. 設定 Nginx 反代到 `127.0.0.1:3000`。
7. 開 `http://<IP>/api/health` 驗證。

---

## 常見錯誤

### 1. 連不到服務

- 檢查 SG 有開 `80`。
- 檢查 `pm2 status` 是否在線。
- 檢查 `sudo systemctl status nginx`。

### 2. npm 指令在 PowerShell 被擋

在 Windows 本機可改用：

```powershell
npm.cmd install
npm.cmd run dev
```

### 3. 上傳檔案失敗

- 檢查 Nginx `client_max_body_size`。
- 檢查 `uploads/` 權限。

---

## 建議下一步

1. 綁定網域（Route 53）。
2. 加 HTTPS（Let's Encrypt + certbot）。
3. 將資料層改為 RDS + S3，避免重開機或重建時遺失。
