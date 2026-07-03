require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

// 中間件
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname)));

// 根目錄直接回傳前端首頁
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// 確保必要目錄存在
const uploadDir = process.env.UPLOAD_DIR || './uploads';
const dataDir = './data';
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

// 簡單的內存數據存儲 (用 JSON 代替 SQLite)
const dbFile = path.join(dataDir, 'orders.json');
const feedbackFile = path.join(dataDir, 'customer_feedback.json');
const productsFile = path.join(dataDir, 'products.json');
const productCategoriesFile = path.join(dataDir, 'product_categories.json');

const defaultFeedbacks = [
  {
    id: 1,
    category: '付款問題',
    created_at: '2026-06-24',
    name: '張睿隆',
    email: 'dlcscoffice@gmail.com',
    phone: '0424823900#206',
    gender: '',
    content: '想請問可以打公司統編嗎？會員是否已經含稅'
  },
  {
    id: 2,
    category: '付款問題',
    created_at: '2026-06-02',
    name: '蔡佳彤',
    email: 'juliachua77@gmail.com',
    phone: '0952622653',
    gender: '',
    content: '您好，請問需要開立統編發票該怎麼進行？'
  },
  {
    id: 3,
    category: '其他',
    created_at: '2026-05-26',
    name: 'Lin',
    email: 'pt2008168@gmail.com',
    phone: '09880000000',
    gender: '',
    content: '您好我在酷澎有買貴公司產品椅子，因有些螺絲他會自行脫落找不到，這邊能補螺絲嗎'
  },
  {
    id: 4,
    category: '商品問題',
    created_at: '2026-05-14',
    name: 'CHOU HSIEN-MIN',
    email: 'docchou0527@gmail.com',
    phone: '0953546809',
    gender: '',
    content: '升降圓椅有沒有五爪固定，不要輪子'
  },
  {
    id: 5,
    category: '其他',
    created_at: '2026-03-05',
    name: '陳映儒',
    email: 'atshopacutek@gmail.com',
    phone: '02-86926616#224',
    gender: '',
    content: '請問 CFLY 御貂摺疊美容床 Pro 超輕量旗艦版有售價 *10 張嗎？可以提供報價單嗎？'
  },
  {
    id: 6,
    category: '商品問題',
    created_at: '2026-02-06',
    name: '陳群',
    email: 'james3511321@gmail.com',
    phone: '0987065146',
    gender: '',
    content: '您好，我這裡是安樂高中運動傷害防護室，近期在考慮採購護理椅，於網路上看到你們的相關資訊。'
  }
];

const defaultProducts = [
  {
    id: 1,
    name: 'CFLY Titan Rest 泰坦人體工學辦公椅',
    style: '曜石黑',
    sku: 'a095011',
    salePrice: 2180,
    listPrice: 3280,
    stock: 1,
    sold: 0,
    keyword: '人體工學,辦公椅',
    status: '已上架'
  },
  {
    id: 2,
    name: 'C-FLY Cloud Master 雲境人體工學高背辦公椅',
    style: '曜石黑',
    sku: 'a094011',
    salePrice: 2580,
    listPrice: 3880,
    stock: 1,
    sold: 0,
    keyword: '高背,辦公椅',
    status: '已上架'
  },
  {
    id: 3,
    name: 'C-FLY Night Ranger 夜巡者人體工學電競椅 腰靠+頸枕',
    style: '曜石黑',
    sku: 'a092011',
    salePrice: 1980,
    listPrice: 2980,
    stock: 1,
    sold: 0,
    keyword: '電競椅,頸枕',
    status: '已上架'
  },
  {
    id: 4,
    name: 'C-FLY Black Falcon 黑隼人體工學電競椅 腰靠+頸枕',
    style: '曜石黑',
    sku: 'a092111',
    salePrice: 2980,
    listPrice: 4480,
    stock: 1,
    sold: 0,
    keyword: '黑隼,電競椅',
    status: '已上架'
  },
  {
    id: 5,
    name: '鋁製天花板',
    style: '單一款式',
    sku: 'a122251',
    salePrice: 888888,
    listPrice: 8888890,
    stock: 0,
    sold: 0,
    keyword: '天花板',
    status: '已上架'
  },
  {
    id: 6,
    name: '電動麻將桌',
    style: '單一款式',
    sku: 'c027001',
    salePrice: 25000,
    listPrice: 25000,
    stock: 0,
    sold: 0,
    keyword: '麻將桌',
    status: '已上架'
  }
];

const defaultProductCategories = [
  { id: 1, name: '人體工學椅', isVisible: true, created_at: '2026-06-30T16:00:14.000Z' },
  { id: 2, name: '家飾建材', isVisible: true, created_at: '2026-06-30T15:56:52.000Z' },
  { id: 3, name: '精品家具', isVisible: true, created_at: '2026-06-30T15:50:09.000Z' },
  { id: 4, name: '精品沙發', isVisible: true, created_at: '2026-06-29T18:10:18.000Z' },
  { id: 5, name: '懶人沙發/居家小物', isVisible: true, created_at: '2026-06-23T10:10:32.000Z' },
  { id: 6, name: '餐椅/餐桌', isVisible: true, created_at: '2026-06-22T17:21:44.000Z' },
  { id: 7, name: '辦公椅/電腦椅', isVisible: true, created_at: '2026-06-18T14:42:51.000Z' },
  { id: 8, name: '收納Cabinet', isVisible: true, created_at: '2026-06-17T11:51:48.000Z' },
  { id: 9, name: '起居Living', isVisible: true, created_at: '2026-06-17T11:36:32.000Z' },
  { id: 10, name: '精選商品', isVisible: true, created_at: '2026-06-17T11:20:41.000Z' }
];

function loadData() {
  try {
    if (fs.existsSync(dbFile)) {
      return JSON.parse(fs.readFileSync(dbFile, 'utf-8'));
    }
  } catch (err) {
    console.warn('無法讀取數據庫文件，使用空數據');
  }
  return { orders: [], nextId: 1 };
}

function saveData(data) {
  fs.writeFileSync(dbFile, JSON.stringify(data, null, 2), 'utf-8');
}

function saveProductsData(data) {
  fs.writeFileSync(productsFile, JSON.stringify(data, null, 2), 'utf-8');
}

function saveProductCategoriesData(data) {
  fs.writeFileSync(productCategoriesFile, JSON.stringify(data, null, 2), 'utf-8');
}

function loadFeedbackData() {
  try {
    if (fs.existsSync(feedbackFile)) {
      const parsed = JSON.parse(fs.readFileSync(feedbackFile, 'utf-8'));
      if (Array.isArray(parsed)) return parsed;
    }
  } catch (err) {
    console.warn('無法讀取顧客回饋資料，使用預設內容');
  }

  fs.writeFileSync(feedbackFile, JSON.stringify(defaultFeedbacks, null, 2), 'utf-8');
  return [...defaultFeedbacks];
}

function loadProductsData() {
  try {
    if (fs.existsSync(productsFile)) {
      const parsed = JSON.parse(fs.readFileSync(productsFile, 'utf-8'));
      if (Array.isArray(parsed)) return parsed;
    }
  } catch (err) {
    console.warn('無法讀取商品資料，使用預設內容');
  }

  fs.writeFileSync(productsFile, JSON.stringify(defaultProducts, null, 2), 'utf-8');
  return [...defaultProducts];
}

function loadProductCategoriesData() {
  try {
    if (fs.existsSync(productCategoriesFile)) {
      const parsed = JSON.parse(fs.readFileSync(productCategoriesFile, 'utf-8'));
      if (Array.isArray(parsed)) return parsed;
    }
  } catch (err) {
    console.warn('無法讀取商品分類資料，使用預設內容');
  }

  fs.writeFileSync(productCategoriesFile, JSON.stringify(defaultProductCategories, null, 2), 'utf-8');
  return [...defaultProductCategories];
}

function normalizeText(value) {
  return String(value || '').trim();
}

function toNonNegativeInt(value, fallback = 0) {
  const n = parseInt(value, 10);
  if (Number.isNaN(n) || n < 0) return fallback;
  return n;
}

function parseExcelBoolean(value, fallback = true) {
  if (typeof value === 'boolean') return value;
  const text = normalizeText(value).toLowerCase();
  if (!text) return fallback;
  const truthy = ['1', 'true', 'yes', 'y', 'on', '是', '啟用', '開啟'];
  const falsy = ['0', 'false', 'no', 'n', 'off', '否', '停用', '關閉'];
  if (truthy.includes(text)) return true;
  if (falsy.includes(text)) return false;
  return fallback;
}

function buildMemberKey(name, phone, email) {
  return `${normalizeText(name)}__${normalizeText(phone)}__${normalizeText(email)}`;
}

function toOrderDateIso(input) {
  const ts = orderDateToTs(input);
  if (!ts) return '';
  return new Date(ts).toISOString();
}

function updateMemberCreatedAtFromOrder(member, orderDateValue) {
  if (!member) return false;
  const orderTs = orderDateToTs(orderDateValue);
  if (!orderTs) return false;

  const currentTs = orderDateToTs(member.created_at);
  if (!currentTs || orderTs < currentTs) {
    member.created_at = new Date(orderTs).toISOString();
    return true;
  }
  return false;
}

function ensureDbShape(data) {
  const safe = data && typeof data === 'object' ? data : {};
  safe.orders = Array.isArray(safe.orders) ? safe.orders : [];
  safe.nextId = Number.isInteger(safe.nextId) ? safe.nextId : ((safe.orders.at(-1)?.id || 0) + 1);
  safe.members = Array.isArray(safe.members) ? safe.members : [];
  safe.nextMemberId = Number.isInteger(safe.nextMemberId) ? safe.nextMemberId : ((safe.members.at(-1)?.id || 0) + 1);
  return safe;
}

let db = ensureDbShape(loadData());
let feedbackData = loadFeedbackData();
let productsData = loadProductsData();
let productCategoriesData = loadProductCategoriesData();

function findMemberByIdentity(name, phone, email) {
  const key = buildMemberKey(name, phone, email);
  return db.members.find((m) => buildMemberKey(m.name, m.phone, m.email) === key) || null;
}

function syncMembersFromOrders() {
  let changed = false;
  const existingKeys = new Set(db.members.map((m) => buildMemberKey(m.name, m.phone, m.email)));

  const earliestOrderTsByKey = {};
  db.orders.forEach((order) => {
    const key = buildMemberKey(order.name, order.phone, order.email);
    if (!key.replace(/_/g, '')) return;
    const ts = orderDateToTs(order.date || order.created_at);
    if (!ts) return;
    if (!earliestOrderTsByKey[key] || ts < earliestOrderTsByKey[key]) {
      earliestOrderTsByKey[key] = ts;
    }
  });

  db.orders.forEach((order) => {
    const name = normalizeText(order.name);
    const phone = normalizeText(order.phone);
    const email = normalizeText(order.email);
    const key = buildMemberKey(name, phone, email);
    if (!key.replace(/_/g, '')) return;

    let member = db.members.find((m) => buildMemberKey(m.name, m.phone, m.email) === key);
    if (!member && !existingKeys.has(key)) {
      const earliestTs = earliestOrderTsByKey[key] || orderDateToTs(order.date || order.created_at) || Date.now();
      member = {
        id: db.nextMemberId++,
        name,
        phone,
        email,
        created_at: new Date(earliestTs).toISOString()
      };
      db.members.push(member);
      existingKeys.add(key);
      changed = true;
    }

    if (member) {
      if (
        order.member_id !== member.id ||
        order.name !== member.name ||
        order.phone !== member.phone ||
        order.email !== member.email
      ) {
        changed = true;
      }
      order.member_id = member.id;
      order.name = member.name;
      order.phone = member.phone;
      order.email = member.email;

      if (updateMemberCreatedAtFromOrder(member, order.date || order.created_at)) {
        changed = true;
      }
    }
  });

  return changed;
}

function requireExistingMember(name, phone, email) {
  const member = findMemberByIdentity(name, phone, email);
  if (!member) {
    throw new Error('會員不存在，請先建立會員再建立訂單');
  }
  return member;
}

function upsertMemberForImport(name, phone, email, orderDateValue = '') {
  const normalizedName = normalizeText(name);
  const normalizedPhone = normalizeText(phone);
  const normalizedEmail = normalizeText(email);

  const existing = findMemberByIdentity(normalizedName, normalizedPhone, normalizedEmail);
  if (existing) {
    updateMemberCreatedAtFromOrder(existing, orderDateValue);
    return existing;
  }

  // 匯入時允許自動建立會員，避免「先建會員」流程阻擋大量匯入
  const fallbackName = normalizedName || `匯入會員-${db.nextMemberId}`;
  const createdAt = toOrderDateIso(orderDateValue) || new Date().toISOString();
  const member = {
    id: db.nextMemberId++,
    name: fallbackName,
    phone: normalizedPhone,
    email: normalizedEmail,
    created_at: createdAt
  };
  db.members.push(member);
  return member;
}

function orderDateToTs(value) {
  if (value == null || value === '') return 0;
  if (typeof value === 'number' && !isNaN(value)) {
    const excelEpoch = new Date(1899, 11, 30);
    return excelEpoch.getTime() + value * 86400000;
  }
  const parsed = Date.parse(String(value));
  return isNaN(parsed) ? 0 : parsed;
}

function applyOrderFilters(orders, query) {
  const search = normalizeText(query.search || '');
  const specificOrder = normalizeText(query.specificOrder || '');
  const paymentMethod = normalizeText(query.paymentMethod || '');
  const deliveryMethod = normalizeText(query.deliveryMethod || '');
  const deliveryStatus = normalizeText(query.deliveryStatus || '');
  const startDate = normalizeText(query.startDate || '');
  const endDate = normalizeText(query.endDate || '');

  const startTs = startDate ? orderDateToTs(startDate) : 0;
  const endTs = endDate ? orderDateToTs(endDate) : 0;

  return orders.filter((order) => {
    if (search) {
      const matched =
        normalizeText(order.invoice).includes(search) ||
        normalizeText(order.name).includes(search) ||
        normalizeText(order.notes).includes(search);
      if (!matched) return false;
    }

    if (specificOrder) {
      const hit =
        String(order.id) === specificOrder ||
        normalizeText(order.invoice).includes(specificOrder) ||
        normalizeText(order.name).includes(specificOrder);
      if (!hit) return false;
    }

    if (paymentMethod && normalizeText(order.payment_method) !== paymentMethod) return false;
    if (deliveryMethod && normalizeText(order.delivery_method) !== deliveryMethod) return false;
    if (deliveryStatus && normalizeText(order.delivery_status) !== deliveryStatus) return false;

    if (startTs || endTs) {
      const orderTs = orderDateToTs(order.date || order.created_at);
      if (startTs && orderTs < startTs) return false;
      if (endTs && orderTs > endTs + 86400000 - 1) return false;
    }

    return true;
  });
}

function dateToKey(input) {
  const ts = orderDateToTs(input);
  if (!ts) return '';
  return new Date(ts).toISOString().slice(0, 10);
}

function parseYmdToTs(ymd) {
  if (!ymd) return 0;
  const ts = Date.parse(`${ymd}T00:00:00Z`);
  return Number.isNaN(ts) ? 0 : ts;
}

function eachDateKey(startYmd, endYmd) {
  const startTs = parseYmdToTs(startYmd);
  const endTs = parseYmdToTs(endYmd);
  if (!startTs || !endTs || endTs < startTs) return [];

  const result = [];
  for (let ts = startTs; ts <= endTs; ts += 86400000) {
    result.push(new Date(ts).toISOString().slice(0, 10));
  }
  return result;
}

function monthKey(input) {
  const ts = orderDateToTs(input);
  if (!ts) return '';
  return new Date(ts).toISOString().slice(0, 7);
}

function eachMonthKey(startYmd, endYmd) {
  const startTs = parseYmdToTs(startYmd);
  const endTs = parseYmdToTs(endYmd);
  if (!startTs || !endTs || endTs < startTs) return [];

  const start = new Date(startTs);
  const end = new Date(endTs);
  const result = [];

  let year = start.getUTCFullYear();
  let month = start.getUTCMonth();
  const endYear = end.getUTCFullYear();
  const endMonth = end.getUTCMonth();

  while (year < endYear || (year === endYear && month <= endMonth)) {
    result.push(`${String(year)}-${String(month + 1).padStart(2, '0')}`);
    month += 1;
    if (month > 11) {
      month = 0;
      year += 1;
    }
  }

  return result;
}

function monthDays(month) {
  const [yearStr, monthStr] = String(month).split('-');
  const year = parseInt(yearStr, 10);
  const monthIdx = parseInt(monthStr, 10);
  if (!year || !monthIdx) return 30;
  return new Date(Date.UTC(year, monthIdx, 0)).getUTCDate();
}

function parseQtyFromProduct(product) {
  const text = String(product || '');
  if (!text.trim()) return 1;

  let qty = 0;
  const regex = /\*\s*(\d+)/g;
  let match = regex.exec(text);
  while (match) {
    qty += parseInt(match[1], 10) || 0;
    match = regex.exec(text);
  }

  return qty > 0 ? qty : 1;
}

function amountBucket(amount) {
  const val = parseFloat(amount || 0) || 0;
  if (val <= 1000) return '0~1000';
  if (val <= 2000) return '1001~2000';
  if (val <= 3000) return '2001~3000';
  if (val <= 5000) return '3001~5000';
  if (val <= 10000) return '5001~10000';
  return '10000以上';
}

function qtyBucket(qty) {
  const val = parseInt(qty, 10) || 1;
  if (val <= 1) return '1';
  if (val === 2) return '2';
  if (val <= 5) return '3~5';
  if (val <= 10) return '6~10';
  return '11~50';
}

const startupSynced = syncMembersFromOrders();
if (startupSynced) {
  saveData(db);
}

// 配置 multer
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});
const upload = multer({ storage });

// ==================== API 路由 ====================

// 上傳並導入 Excel
app.post('/api/import', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '未上傳文件' });
    }

    const filePath = req.file.path;
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    if (!data || data.length === 0) {
      return res.status(400).json({ error: 'Excel 文件為空' });
    }

    // 驗證並插入數據
    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    data.forEach((row, index) => {
      try {
        const invoice = row['訂單編號'] || row['發票'] || row['invoice'] || '';
        const name = row['會員名稱'] || row['名稱'] || row['name'] || '';
        const phone = row['會員電話'] || row['電話'] || row['phone'] || '';
        const email = row['會員信箱'] || row['信箱'] || row['email'] || '';
        const product = row['商品詳情'] || row['商品'] || row['product'] || '';
        const tag = row['訂單標籤'] || row['標籤'] || row['tag'] || '';
        const rawDate = row['成立日期'] || row['日期'] || row['date'] || '';
        const payment_status = row['付款狀態'] || row['payment_status'] || '已收到款項';
        const payment_method = row['付款方式'] || row['payment_method'] || '信用卡';
        const delivery_status = row['配送狀態'] || row['delivery_status'] || '已出貨';
        const delivery_method = row['配送方式'] || row['delivery_method'] || '宅配到府';
        const temperature = row['配送溫層'] || row['temperature'] || '常溫';
        const return_status = row['退貨狀態'] || row['return_status'] || '不需退貨';
        const notes = row['訂單備註'] || row['備註'] || row['notes'] || '';
        const rawAmount = row['訂單總額'] || row['金額'] || row['amount'] || 0;
        const amount = parseFloat(rawAmount) || 0;

        let date = rawDate;
        if (typeof rawDate === 'number' && !isNaN(rawDate)) {
          const excelEpoch = new Date(1899, 11, 30);
          const d = new Date(excelEpoch.getTime() + rawDate * 86400000);
          date = d.toISOString().slice(0, 10);
        } else if (rawDate) {
          date = String(rawDate).slice(0, 10);
        } else {
          date = new Date().toISOString().slice(0, 10);
        }

        const member = upsertMemberForImport(name, phone, email, date);

        db.orders.push({
          id: db.nextId++,
          member_id: member.id,
          invoice,
          name: member.name,
          phone: member.phone,
          email: member.email,
          product,
          tag,
          amount,
          date,
          payment_status,
          payment_method,
          delivery_status,
          delivery_method,
          temperature,
          return_status,
          notes,
          created_at: new Date().toISOString()
        });
        successCount++;
      } catch (err) {
        errorCount++;
        errors.push(`第 ${index + 2} 行: ${err.message}`);
      }
    });

    saveData(db);

    // 刪除上傳文件
    fs.unlinkSync(filePath);

    res.json({
      success: true,
      successCount,
      errorCount,
      errors: errors.slice(0, 10),
      message: `成功導入 ${successCount} 條訂單，失敗 ${errorCount} 條`
    });
  } catch (err) {
    console.error('導入錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 匯入商品 Excel（同 SKU 更新，無 SKU 則新增）
app.post('/api/products/import', upload.single('file'), (req, res) => {
  let filePath = '';
  try {
    if (!req.file) {
      return res.status(400).json({ error: '未上傳文件' });
    }

    filePath = req.file.path;
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    if (!rows || rows.length === 0) {
      return res.status(400).json({ error: 'Excel 文件為空' });
    }

    let createdCount = 0;
    let updatedCount = 0;
    let errorCount = 0;
    const errors = [];
    let nextId = productsData.reduce((max, item) => Math.max(max, Number(item.id) || 0), 0) + 1;

    rows.forEach((row, index) => {
      try {
        const name = normalizeText(row['商品名稱'] || row['name'] || row['Name']);
        const sku = normalizeText(row['商品編號'] || row['商品編號(SKU)'] || row['SKU'] || row['sku']);
        const style = normalizeText(row['款式'] || row['style']) || '單一款式';

        if (!name) {
          throw new Error('商品名稱不可為空');
        }
        if (!sku) {
          throw new Error('商品編號(SKU)不可為空');
        }

        const payload = {
          name,
          style,
          sku,
          salePrice: toNonNegativeInt(row['售價'] ?? row['salePrice'] ?? row['price'] ?? 0, 0),
          listPrice: toNonNegativeInt(row['定價'] ?? row['listPrice'] ?? row['originalPrice'] ?? 0, 0),
          stock: toNonNegativeInt(row['庫存數量'] ?? row['庫存'] ?? row['stock'] ?? 0, 0),
          sold: toNonNegativeInt(row['已售出數'] ?? row['sold'] ?? 0, 0),
          keyword: normalizeText(row['商品搜字'] || row['商品搜尋字'] || row['keyword']),
          status: normalizeText(row['上架狀態'] || row['status']) || '已上架',
          tag: normalizeText(row['商品標語'] || row['tag']),
          description: normalizeText(row['商品簡述'] || row['description']),
          urlPath: normalizeText(row['商品網址'] || row['urlPath']) || sku,
          searchEnabled: parseExcelBoolean(row['啟用商品搜尋功能'] ?? row['啟用搜尋'] ?? row['searchEnabled'], true)
        };

        const existing = productsData.find((item) => normalizeText(item.sku) === sku);
        if (existing) {
          Object.assign(existing, payload);
          updatedCount++;
        } else {
          productsData.push({
            id: nextId++,
            ...payload
          });
          createdCount++;
        }
      } catch (err) {
        errorCount++;
        errors.push(`第 ${index + 2} 行: ${err.message}`);
      }
    });

    saveProductsData(productsData);

    res.json({
      success: true,
      createdCount,
      updatedCount,
      errorCount,
      errors: errors.slice(0, 20),
      message: `商品匯入完成：新增 ${createdCount} 筆、更新 ${updatedCount} 筆、失敗 ${errorCount} 筆`
    });
  } catch (err) {
    console.error('匯入商品錯誤:', err);
    res.status(500).json({ error: err.message });
  } finally {
    if (filePath && fs.existsSync(filePath)) {
      try {
        fs.unlinkSync(filePath);
      } catch (_) {
      }
    }
  }
});

// 建立新訂單
app.post('/api/orders', (req, res) => {
  try {
    const {
      invoice,
      name,
      phone,
      email,
      product,
      tag,
      amount,
      date,
      payment_status,
      payment_method,
      delivery_status,
      delivery_method,
      temperature,
      return_status,
      notes
    } = req.body;

    const normalizedAmount = typeof amount === 'number' && !isNaN(amount) && amount > 0 ? amount : 0;
    const normalizedDate = date || new Date().toISOString().slice(0, 10);

    const member = requireExistingMember(name, phone, email);
    updateMemberCreatedAtFromOrder(member, normalizedDate);

    const newOrder = {
      id: db.nextId++,
      member_id: member.id,
      invoice: invoice || '',
      name: member.name,
      phone: member.phone,
      email: member.email,
      product: product || '',
      tag: tag || '',
      amount: normalizedAmount,
      date: normalizedDate,
      payment_status: payment_status || '已收到款項',
      payment_method: payment_method || '信用卡',
      delivery_status: delivery_status || '已出貨',
      delivery_method: delivery_method || '宅配到府',
      temperature: temperature || '常溫',
      return_status: return_status || '不需退貨',
      notes: notes || '',
      created_at: new Date().toISOString()
    };

    db.orders.push(newOrder);
    saveData(db);

    res.status(201).json({ success: true, data: newOrder });
  } catch (err) {
    console.error('建立訂單錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 獲取單筆訂單詳情
app.get('/api/orders/:id(\\d+)', (req, res) => {
  try {
    const id = parseInt(req.params.id, 10);
    const order = db.orders.find(item => item.id === id);
    if (!order) {
      return res.status(404).json({ error: '訂單不存在' });
    }
    res.json(order);
  } catch (err) {
    console.error('查詢訂單詳情錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 更新訂單內容
app.put('/api/orders/:id(\\d+)', (req, res) => {
  try {
    const id = parseInt(req.params.id, 10);
    const order = db.orders.find(item => item.id === id);
    if (!order) {
      return res.status(404).json({ error: '訂單不存在' });
    }

    const {
      invoice,
      name,
      phone,
      email,
      product,
      tag,
      amount,
      date,
      payment_status,
      payment_method,
      delivery_status,
      delivery_method,
      temperature,
      return_status,
      notes
    } = req.body;

    const normalizedAmount = typeof amount === 'number' && !isNaN(amount) && amount > 0 ? amount : 0;
    const normalizedDate = date || order.date || new Date().toISOString().slice(0, 10);

    const member = requireExistingMember(name, phone, email);
    updateMemberCreatedAtFromOrder(member, normalizedDate);

    order.invoice = invoice || '';
    order.member_id = member.id;
    order.name = member.name;
    order.phone = member.phone;
    order.email = member.email;
    order.product = product || '';
    order.tag = tag || '';
    order.amount = normalizedAmount;
    order.date = normalizedDate;
    order.payment_status = payment_status || '已收到款項';
    order.payment_method = payment_method || '信用卡';
    order.delivery_status = delivery_status || '已出貨';
    order.delivery_method = delivery_method || '宅配到府';
    order.temperature = temperature || '常溫';
    order.return_status = return_status || '不需退貨';
    order.notes = notes || '';

    saveData(db);

    res.json({ success: true, data: order });
  } catch (err) {
    console.error('更新訂單錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 獲取訂單列表（分頁、搜索、排序）
app.get('/api/orders', (req, res) => {
  try {
    const page = parseInt(req.query.page) || 1;
    const pageSize = parseInt(req.query.pageSize) || 20;
    const sortBy = req.query.sortBy || 'date';
    const sortOrder = req.query.sortOrder || 'DESC';

    let filtered = applyOrderFilters(db.orders, req.query);

    // 排序
    filtered.sort((a, b) => {
      if (sortBy === 'date') {
        const aVal = orderDateToTs(a.date || a.created_at);
        const bVal = orderDateToTs(b.date || b.created_at);
        if (sortOrder === 'ASC') return aVal - bVal;
        return bVal - aVal;
      }

      let aVal = a[sortBy];
      let bVal = b[sortBy];
      
      if (typeof aVal === 'string') {
        aVal = aVal.toLowerCase();
        bVal = bVal.toLowerCase();
      }

      if (sortOrder === 'ASC') {
        return aVal > bVal ? 1 : -1;
      } else {
        return aVal < bVal ? 1 : -1;
      }
    });

    const total = filtered.length;
    const offset = (page - 1) * pageSize;
    const orders = filtered.slice(offset, offset + pageSize);

    res.json({
      data: orders,
      pagination: {
        page,
        pageSize,
        total,
        pages: Math.ceil(total / pageSize)
      }
    });
  } catch (err) {
    console.error('查詢錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 匯出訂單 Excel
app.get('/api/orders/export', (req, res) => {
  try {
    let exportOrders = applyOrderFilters(db.orders, req.query);

    const rows = exportOrders.map((order) => ({
      訂單ID: order.id,
      訂單編號: order.invoice || '',
      會員名稱: order.name || '',
      會員電話: order.phone || '',
      會員信箱: order.email || '',
      商品詳情: order.product || '',
      訂單標籤: order.tag || '',
      成立日期: order.date || '',
      付款狀態: order.payment_status || '',
      付款方式: order.payment_method || '',
      配送狀態: order.delivery_status || '',
      配送方式: order.delivery_method || '',
      配送溫層: order.temperature || '',
      退貨狀態: order.return_status || '',
      訂單備註: order.notes || '',
      訂單總額: parseFloat(order.amount || 0) || 0,
      建立時間: order.created_at || ''
    }));

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, '訂單列表');

    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    const filename = `orders-${new Date().toISOString().slice(0, 10)}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.send(buffer);
  } catch (err) {
    console.error('匯出訂單錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 獲取統計數據
app.get('/api/statistics', (req, res) => {
  try {
    const totalOrders = db.orders.length;
    const totalAmount = db.orders.reduce((sum, order) => sum + order.amount, 0);
    const avgAmount = totalOrders > 0 ? (totalAmount / totalOrders).toFixed(2) : 0;

    // 按日期統計
    const trendMap = {};
    db.orders.forEach(order => {
      const dateKey = dateToKey(order.date || order.created_at);
      if (!dateKey) return;

      if (!trendMap[dateKey]) {
        trendMap[dateKey] = { date: dateKey, count: 0, total: 0 };
      }
      trendMap[dateKey].count++;
      trendMap[dateKey].total += order.amount;
    });

    const sortedTrend = Object.values(trendMap)
      .sort((a, b) => a.date.localeCompare(b.date));

    const limit = parseInt(req.query.limit, 10);
    const trend = Number.isInteger(limit) && limit > 0
      ? sortedTrend.slice(-limit)
      : sortedTrend;

    // 按客戶統計
    const customerMap = {};
    db.orders.forEach(order => {
      if (!customerMap[order.name]) {
        customerMap[order.name] = { name: order.name, count: 0, total: 0 };
      }
      customerMap[order.name].count++;
      customerMap[order.name].total += order.amount;
    });

    const topCustomers = Object.values(customerMap)
      .sort((a, b) => b.total - a.total)
      .slice(0, 10);

    res.json({
      totalOrders,
      totalAmount: totalAmount.toFixed(2),
      avgAmount,
      trend,
      topCustomers
    });
  } catch (err) {
    console.error('統計錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 銷售概況（仿 Cyberbiz 後台日報表）
app.get('/api/sales-overview', (req, res) => {
  try {
    const allDateKeys = db.orders
      .map((order) => dateToKey(order.date || order.created_at))
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b));

    const lastDate = allDateKeys.length > 0 ? allDateKeys.at(-1) : new Date().toISOString().slice(0, 10);
    const defaultEndTs = parseYmdToTs(lastDate);
    const defaultStartTs = defaultEndTs - 31 * 86400000;

    const requestedStart = normalizeText(req.query.startDate || '');
    const requestedEnd = normalizeText(req.query.endDate || '');
    const startDate = requestedStart || new Date(defaultStartTs).toISOString().slice(0, 10);
    const endDate = requestedEnd || lastDate;

    const dateList = eachDateKey(startDate, endDate);
    if (dateList.length === 0) {
      return res.status(400).json({ error: '日期區間格式錯誤或開始日期晚於結束日期' });
    }

    const dayMap = Object.fromEntries(dateList.map((date) => [date, {
      date,
      traffic: 0,
      orderCount: 0,
      totalAmount: 0,
      avgOrderValue: 0,
      conversionRate: 0,
      newMembers: 0
    }]));

    db.orders.forEach((order) => {
      const key = dateToKey(order.date || order.created_at);
      if (!key || !dayMap[key]) return;
      dayMap[key].orderCount += 1;
      dayMap[key].totalAmount += parseFloat(order.amount || 0) || 0;
    });

    // 無流量原始資料時，建立穩定可重現的估算流量，避免 UI 顯示全 0。
    dateList.forEach((date) => {
      const seed = date.split('-').reduce((sum, p) => sum + parseInt(p, 10), 0);
      const day = dayMap[date];
      const baseline = 8 + (seed % 11);
      day.traffic = Math.max(day.orderCount * 18 + baseline, day.orderCount);
      day.avgOrderValue = day.orderCount > 0 ? Math.round(day.totalAmount / day.orderCount) : 0;
      day.conversionRate = day.traffic > 0
        ? Number(((day.orderCount / day.traffic) * 100).toFixed(2))
        : 0;
    });

    db.members.forEach((member) => {
      const key = dateToKey(member.created_at);
      if (key && dayMap[key]) dayMap[key].newMembers += 1;
    });

    const rows = dateList.map((date) => {
      const item = dayMap[date];
      return {
        date,
        traffic: item.traffic,
        conversionRate: item.conversionRate,
        orderCount: item.orderCount,
        avgOrderValue: item.avgOrderValue,
        revenue: Math.round(item.totalAmount),
        newMembers: item.newMembers
      };
    });

    const totalTraffic = rows.reduce((sum, row) => sum + row.traffic, 0);
    const totalOrders = rows.reduce((sum, row) => sum + row.orderCount, 0);
    const totalRevenue = rows.reduce((sum, row) => sum + row.revenue, 0);
    const totalNewMembers = rows.reduce((sum, row) => sum + row.newMembers, 0);

    const summary = {
      totalTraffic,
      conversionRate: totalTraffic > 0 ? Number(((totalOrders / totalTraffic) * 100).toFixed(2)) : 0,
      avgOrderValue: totalOrders > 0 ? Math.round(totalRevenue / totalOrders) : 0,
      totalOrders,
      totalRevenue,
      totalNewMembers
    };

    const prevStartTs = parseYmdToTs(startDate) - dateList.length * 86400000;
    const prevEndTs = parseYmdToTs(startDate) - 86400000;
    const prevDates = eachDateKey(
      new Date(prevStartTs).toISOString().slice(0, 10),
      new Date(prevEndTs).toISOString().slice(0, 10)
    );

    const prevRevenueMap = Object.fromEntries(prevDates.map((date) => [date, 0]));
    db.orders.forEach((order) => {
      const key = dateToKey(order.date || order.created_at);
      if (!key || !(key in prevRevenueMap)) return;
      prevRevenueMap[key] += parseFloat(order.amount || 0) || 0;
    });

    const chart = {
      labels: rows.map((row) => row.date),
      revenue: rows.map((row) => row.revenue),
      orders: rows.map((row) => row.orderCount),
      prevRevenue: prevDates.map((date) => Math.round(prevRevenueMap[date] || 0))
    };

    res.json({
      range: { startDate, endDate },
      summary,
      chart,
      rows
    });
  } catch (err) {
    console.error('銷售概況錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 訂單分析（仿 Cyberbiz 訂單分析頁）
app.get('/api/order-analysis', (req, res) => {
  try {
    const now = new Date();
    const defaultStart = `${now.getUTCFullYear()}-01-01`;
    const defaultEnd = now.toISOString().slice(0, 10);
    const startDate = normalizeText(req.query.startDate || '') || defaultStart;
    const endDate = normalizeText(req.query.endDate || '') || defaultEnd;

    const startTs = parseYmdToTs(startDate);
    const endTs = parseYmdToTs(endDate) + 86400000 - 1;
    if (!startTs || !endTs || endTs < startTs) {
      return res.status(400).json({ error: '日期區間格式錯誤或開始日期晚於結束日期' });
    }

    const inRangeOrders = db.orders.filter((order) => {
      const ts = orderDateToTs(order.date || order.created_at);
      return ts >= startTs && ts <= endTs;
    });

    const monthLabels = eachMonthKey(startDate, endDate);
    const monthMap = Object.fromEntries(monthLabels.map((month) => [month, {
      orderCount: 0,
      totalAmount: 0,
      cancelCount: 0,
      returnCount: 0,
      amountBuckets: {
        '0~1000': 0,
        '1001~2000': 0,
        '2001~3000': 0,
        '3001~5000': 0,
        '5001~10000': 0,
        '10000以上': 0
      }
    }]));

    const qtyDist = {
      '1': 0,
      '2': 0,
      '3~5': 0,
      '6~10': 0,
      '11~50': 0
    };

    const amountShare = {
      '0~1000': 0,
      '1001~2000': 0,
      '2001~3000': 0,
      '3001~5000': 0,
      '5001~10000': 0,
      '10000以上': 0
    };

    inRangeOrders.forEach((order) => {
      const mKey = monthKey(order.date || order.created_at);
      if (!mKey || !monthMap[mKey]) return;

      const amount = parseFloat(order.amount || 0) || 0;
      const bucket = amountBucket(amount);
      const qty = parseQtyFromProduct(order.product);
      const qtyB = qtyBucket(qty);

      const paymentStatus = normalizeText(order.payment_status);
      const deliveryStatus = normalizeText(order.delivery_status);
      const returnStatus = normalizeText(order.return_status);

      monthMap[mKey].orderCount += 1;
      monthMap[mKey].totalAmount += amount;
      monthMap[mKey].amountBuckets[bucket] += 1;

      if (paymentStatus.includes('取消') || deliveryStatus.includes('取消')) {
        monthMap[mKey].cancelCount += 1;
      }
      if (returnStatus.includes('退貨') && !returnStatus.includes('不需')) {
        monthMap[mKey].returnCount += 1;
      }

      qtyDist[qtyB] += 1;
      amountShare[bucket] += 1;
    });

    const monthlyOrderTrend = monthLabels.map((m) => monthMap[m].orderCount);
    const dailyAvgOrderTrend = monthLabels.map((m) => {
      const days = monthDays(m);
      return Number((monthMap[m].orderCount / days).toFixed(2));
    });
    const dailyAvgAmountTrend = monthLabels.map((m) => {
      const days = monthDays(m);
      return Number((monthMap[m].totalAmount / days).toFixed(2));
    });
    const cancelRateTrend = monthLabels.map((m) => {
      const total = monthMap[m].orderCount;
      return total > 0 ? Number(((monthMap[m].cancelCount / total) * 100).toFixed(2)) : 0;
    });
    const returnRateTrend = monthLabels.map((m) => {
      const total = monthMap[m].orderCount;
      return total > 0 ? Number(((monthMap[m].returnCount / total) * 100).toFixed(2)) : 0;
    });

    const amountBucketLabels = ['0~1000', '1001~2000', '2001~3000', '3001~5000', '5001~10000', '10000以上'];
    const qtyBucketLabels = ['1', '2', '3~5', '6~10', '11~50'];

    const totalOrders = inRangeOrders.length;
    const thirtyDaysAgo = Date.now() - 30 * 86400000;
    const last30Orders = inRangeOrders.filter((o) => orderDateToTs(o.date || o.created_at) >= thirtyDaysAgo).length;

    const latestIdx = monthLabels.length - 1;
    const prevIdx = monthLabels.length - 2;
    const safePrev = (arr) => (prevIdx >= 0 ? arr[prevIdx] : 0);

    const kpis = {
      last30AvgDailyOrders: Number((last30Orders / 30).toFixed(2)),
      deltaDailyOrders: Number((dailyAvgOrderTrend[latestIdx] - safePrev(dailyAvgOrderTrend)).toFixed(2)),
      last30AvgDailyAmount: Number((inRangeOrders
        .filter((o) => orderDateToTs(o.date || o.created_at) >= thirtyDaysAgo)
        .reduce((sum, o) => sum + (parseFloat(o.amount || 0) || 0), 0) / 30).toFixed(2)),
      deltaDailyAmount: Number((dailyAvgAmountTrend[latestIdx] - safePrev(dailyAvgAmountTrend)).toFixed(2)),
      last30CancelRate: Number((cancelRateTrend[latestIdx] || 0).toFixed(2)),
      deltaCancelRate: Number(((cancelRateTrend[latestIdx] || 0) - safePrev(cancelRateTrend)).toFixed(2)),
      last30ReturnRate: Number((returnRateTrend[latestIdx] || 0).toFixed(2)),
      deltaReturnRate: Number(((returnRateTrend[latestIdx] || 0) - safePrev(returnRateTrend)).toFixed(2))
    };

    res.json({
      range: { startDate, endDate },
      summary: { totalOrders, last30Orders },
      labels: monthLabels,
      monthlyOrderTrend,
      amountRangeTrend: amountBucketLabels.map((label) => ({
        name: label,
        values: monthLabels.map((m) => monthMap[m].amountBuckets[label] || 0)
      })),
      qtyDistribution: qtyBucketLabels.map((label) => ({ name: label, value: qtyDist[label] || 0 })),
      amountShare: amountBucketLabels.map((label) => ({ name: label, value: amountShare[label] || 0 })),
      dailyAvgOrderTrend,
      dailyAvgAmountTrend,
      cancelRateTrend,
      returnRateTrend,
      kpis
    });
  } catch (err) {
    console.error('訂單分析錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

function customerKeyFromRecord(name, phone, email, fallback = '') {
  const n = normalizeText(name);
  if (n) return `name:${n.toLowerCase()}`;

  const p = normalizeText(phone);
  if (p) return `phone:${p}`;

  const e = normalizeText(email);
  if (e) return `email:${e.toLowerCase()}`;

  return fallback || 'unknown';
}

// 客戶分析（同名視為同一客戶）
app.get('/api/customer-analysis', (req, res) => {
  try {
    const now = new Date();
    const defaultStart = `${now.getUTCFullYear()}-01-01`;
    const defaultEnd = now.toISOString().slice(0, 10);
    const startDate = normalizeText(req.query.startDate || '') || defaultStart;
    const endDate = normalizeText(req.query.endDate || '') || defaultEnd;

    const startTs = parseYmdToTs(startDate);
    const endTs = parseYmdToTs(endDate) + 86400000 - 1;
    if (!startTs || !endTs || endTs < startTs) {
      return res.status(400).json({ error: '日期區間格式錯誤或開始日期晚於結束日期' });
    }

    const labels = eachMonthKey(startDate, endDate);
    const monthStats = Object.fromEntries(labels.map((m) => [m, {
      newCustomers: 0,
      activeCustomers: new Set(),
      orderCount: 0,
      amount: 0
    }]));

    const customers = new Map();
    const allOrders = db.orders
      .map((order) => {
        const ts = orderDateToTs(order.date || order.created_at);
        return { ...order, __ts: ts };
      })
      .filter((order) => order.__ts >= startTs && order.__ts <= endTs)
      .sort((a, b) => a.__ts - b.__ts);

    allOrders.forEach((order) => {
      const key = customerKeyFromRecord(order.name, order.phone, order.email, `order:${order.id}`);
      const mKey = monthKey(order.date || order.created_at);
      if (!mKey || !monthStats[mKey]) return;

      if (!customers.has(key)) {
        customers.set(key, {
          key,
          displayName: normalizeText(order.name) || '未命名客戶',
          firstTs: order.__ts,
          firstMonth: mKey,
          orders: []
        });
        monthStats[mKey].newCustomers += 1;
      }

      const c = customers.get(key);
      c.orders.push(order);
      monthStats[mKey].activeCustomers.add(key);
      monthStats[mKey].orderCount += 1;
      monthStats[mKey].amount += parseFloat(order.amount || 0) || 0;
    });

    const memberDedup = new Map();
    db.members.forEach((m) => {
      const key = customerKeyFromRecord(m.name, m.phone, m.email, `member:${m.id}`);
      if (!memberDedup.has(key)) {
        memberDedup.set(key, {
          key,
          createdTs: orderDateToTs(m.created_at),
          firstMonth: monthKey(m.created_at)
        });
      }
    });

    const totalCustomers = customers.size;
    const thirtyDaysAgo = Date.now() - 30 * 86400000;
    const last30NewCustomers = Array.from(customers.values()).filter((c) => c.firstTs >= thirtyDaysAgo).length;

    const newCustomerTrend = labels.map((m) => monthStats[m].newCustomers);
    let cumulative = 0;
    const totalCustomerTrend = labels.map((m) => {
      cumulative += monthStats[m].newCustomers;
      return cumulative;
    });

    const freq = {
      '1次購買': 0,
      '2次購買': 0,
      '3次以上': 0,
      '無購買紀錄': 0
    };

    customers.forEach((c) => {
      const count = c.orders.length;
      if (count <= 1) freq['1次購買'] += 1;
      else if (count === 2) freq['2次購買'] += 1;
      else freq['3次以上'] += 1;
    });

    memberDedup.forEach((m) => {
      if (!customers.has(m.key)) {
        freq['無購買紀錄'] += 1;
      }
    });

    const newMemberOrderRateAll = labels.map((m) => {
      const newcomers = Array.from(memberDedup.values()).filter((item) => item.firstMonth === m);
      if (newcomers.length === 0) return { month: m, hasOrder: 0, noOrder: 0 };

      let hasOrder = 0;
      newcomers.forEach((item) => {
        const customer = customers.get(item.key);
        if (!customer) return;
        const monthOrder = customer.orders.some((o) => monthKey(o.date || o.created_at) === m);
        if (monthOrder) hasOrder += 1;
      });

      const ratio = Number(((hasOrder / newcomers.length) * 100).toFixed(2));
      return { month: m, hasOrder: ratio, noOrder: Number((100 - ratio).toFixed(2)) };
    });

    const newMemberOrderRatePaid = labels.map((m) => {
      const newcomers = Array.from(memberDedup.values()).filter((item) => item.firstMonth === m);
      if (newcomers.length === 0) return { month: m, hasOrder: 0, noOrder: 0 };

      let hasOrder = 0;
      newcomers.forEach((item) => {
        const customer = customers.get(item.key);
        if (!customer) return;
        const monthPaid = customer.orders.some((o) => {
          const mm = monthKey(o.date || o.created_at);
          const paid = normalizeText(o.payment_status).includes('已收到款項');
          return mm === m && paid;
        });
        if (monthPaid) hasOrder += 1;
      });

      const ratio = Number(((hasOrder / newcomers.length) * 100).toFixed(2));
      return { month: m, hasOrder: ratio, noOrder: Number((100 - ratio).toFixed(2)) };
    });

    const avgRecognizedOrderTrend = labels.map((m) => {
      const active = monthStats[m].activeCustomers.size;
      if (active === 0) return 0;
      return Number((monthStats[m].orderCount / active).toFixed(2));
    });

    const avgSpendTrend = labels.map((m) => {
      const active = monthStats[m].activeCustomers.size;
      if (active === 0) return 0;
      return Number((monthStats[m].amount / active).toFixed(2));
    });

    const avgRecognizedOrders = totalCustomers > 0
      ? Number((allOrders.length / totalCustomers).toFixed(2))
      : 0;

    let repurchaseCustomer = 0;
    const repurchaseGaps = [];
    customers.forEach((c) => {
      const sortedTs = c.orders.map((o) => o.__ts).sort((a, b) => a - b);
      if (sortedTs.length >= 2) {
        repurchaseCustomer += 1;
        for (let i = 1; i < sortedTs.length; i++) {
          repurchaseGaps.push((sortedTs[i] - sortedTs[i - 1]) / 86400000);
        }
      }
    });

    const avgRepurchaseDays = repurchaseGaps.length > 0
      ? Number((repurchaseGaps.reduce((sum, d) => sum + d, 0) / repurchaseGaps.length).toFixed(2))
      : 0;

    const repurchaseRate = totalCustomers > 0
      ? Number(((repurchaseCustomer / totalCustomers) * 100).toFixed(2))
      : 0;

    res.json({
      range: { startDate, endDate },
      summary: {
        totalCustomers,
        last30NewCustomers
      },
      labels,
      memberCountTrend: {
        newMembers: newCustomerTrend,
        totalMembers: totalCustomerTrend
      },
      purchaseFreqDistribution: Object.entries(freq).map(([name, value]) => ({ name, value })),
      newMemberOrderRateAll,
      newMemberOrderRatePaid,
      avgRecognizedOrderTrend,
      avgSpendTrend,
      kpis: {
        avgRecognizedOrders,
        avgRepurchaseDays,
        repurchaseRate
      }
    });
  } catch (err) {
    console.error('客戶分析錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 顧客回饋與建議列表
app.get('/api/customer-feedback', (req, res) => {
  try {
    const page = parseInt(req.query.page, 10) || 1;
    const pageSize = parseInt(req.query.pageSize, 10) || 10;
    const search = normalizeText(req.query.search || '').toLowerCase();

    let list = [...feedbackData];
    if (search) {
      list = list.filter((item) =>
        normalizeText(item.name).toLowerCase().includes(search) ||
        normalizeText(item.email).toLowerCase().includes(search) ||
        normalizeText(item.phone).toLowerCase().includes(search) ||
        normalizeText(item.content).toLowerCase().includes(search)
      );
    }

    list.sort((a, b) => String(b.created_at || '').localeCompare(String(a.created_at || '')));

    const total = list.length;
    const pages = Math.max(1, Math.ceil(total / pageSize));
    const safePage = Math.min(Math.max(page, 1), pages);
    const offset = (safePage - 1) * pageSize;
    const data = list.slice(offset, offset + pageSize);

    res.json({
      data,
      pagination: {
        page: safePage,
        pageSize,
        total,
        pages
      }
    });
  } catch (err) {
    console.error('查詢顧客回饋錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 商品列表（搜尋/分頁）
app.get('/api/products', (req, res) => {
  try {
    const page = parseInt(req.query.page, 10) || 1;
    const pageSize = parseInt(req.query.pageSize, 10) || 10;
    const search = normalizeText(req.query.search || '').toLowerCase();

    let list = [...productsData];
    if (search) {
      list = list.filter((item) =>
        normalizeText(item.name).toLowerCase().includes(search) ||
        normalizeText(item.style).toLowerCase().includes(search) ||
        normalizeText(item.sku).toLowerCase().includes(search) ||
        normalizeText(item.keyword).toLowerCase().includes(search)
      );
    }

    const total = list.length;
    const pages = Math.max(1, Math.ceil(total / pageSize));
    const safePage = Math.min(Math.max(page, 1), pages);
    const offset = (safePage - 1) * pageSize;
    const data = list.slice(offset, offset + pageSize);

    res.json({
      data,
      pagination: {
        page: safePage,
        pageSize,
        total,
        pages
      }
    });
  } catch (err) {
    console.error('查詢商品列表錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 庫存列表（分頁/搜尋）
app.get('/api/products/inventory', (req, res) => {
  try {
    const page = parseInt(req.query.page, 10) || 1;
    const pageSize = parseInt(req.query.pageSize, 10) || 10;
    const search = normalizeText(req.query.search || '').toLowerCase();

    let list = productsData.map((item) => ({
      id: item.id,
      name: normalizeText(item.name),
      style: normalizeText(item.style) || '單一款式',
      sku: normalizeText(item.sku),
      stock: Math.max(0, parseInt(item.stock, 10) || 0),
      lowStockAction: normalizeText(item.lowStockAction) || '拒絕客戶購買',
      image: normalizeText(item.image)
    }));

    if (search) {
      list = list.filter((item) =>
        item.name.toLowerCase().includes(search) ||
        item.style.toLowerCase().includes(search) ||
        item.sku.toLowerCase().includes(search)
      );
    }

    list.sort((a, b) => {
      if (a.stock !== b.stock) return a.stock - b.stock;
      return a.name.localeCompare(b.name, 'zh-Hant');
    });

    const total = list.length;
    const pages = Math.max(1, Math.ceil(total / pageSize));
    const safePage = Math.min(Math.max(page, 1), pages);
    const offset = (safePage - 1) * pageSize;
    const data = list.slice(offset, offset + pageSize);

    res.json({
      data,
      pagination: {
        page: safePage,
        pageSize,
        total,
        pages
      }
    });
  } catch (err) {
    console.error('查詢庫存列表錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 商品分類頁資料（左商品清單 + 右群組）
app.get('/api/product-categories/overview', (req, res) => {
  try {
    const page = parseInt(req.query.page, 10) || 1;
    const pageSize = parseInt(req.query.pageSize, 10) || 10;
    const search = normalizeText(req.query.search || '').toLowerCase();
    const sort = normalizeText(req.query.sort || 'recent-desc');

    const categories = [...productCategoriesData].sort((a, b) => {
      const at = Date.parse(a.created_at || 0) || 0;
      const bt = Date.parse(b.created_at || 0) || 0;
      return sort === 'recent-asc' ? at - bt : bt - at;
    });

    const fallbackCategory = categories[0]?.name || '未分類';
    let list = productsData.map((item) => ({
      id: item.id,
      name: normalizeText(item.name),
      price: Math.max(0, parseInt(item.salePrice, 10) || 0),
      type: normalizeText(item.style) || '單一款式',
      store: 'C-FLY家具生活館',
      vendor: 'C-FLY家具生活館',
      created_at: item.created_at || new Date(Date.now() - (Number(item.id) || 0) * 86400000).toISOString(),
      category: normalizeText(item.category) || fallbackCategory,
      sku: normalizeText(item.sku)
    }));

    if (search) {
      list = list.filter((item) =>
        item.name.toLowerCase().includes(search) ||
        item.type.toLowerCase().includes(search) ||
        item.store.toLowerCase().includes(search) ||
        item.vendor.toLowerCase().includes(search) ||
        item.sku.toLowerCase().includes(search)
      );
    }

    list.sort((a, b) => String(b.created_at).localeCompare(String(a.created_at)));

    const total = list.length;
    const pages = Math.max(1, Math.ceil(total / pageSize));
    const safePage = Math.min(Math.max(page, 1), pages);
    const offset = (safePage - 1) * pageSize;
    const data = list.slice(offset, offset + pageSize);

    res.json({
      products: {
        data,
        pagination: {
          page: safePage,
          pageSize,
          total,
          pages
        }
      },
      categories
    });
  } catch (err) {
    console.error('查詢商品分類頁資料錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 新增商品群組
app.post('/api/product-categories', (req, res) => {
  try {
    const name = normalizeText(req.body.name);
    if (!name) {
      return res.status(400).json({ error: '群組名稱不可為空' });
    }

    const duplicate = productCategoriesData.find((item) => normalizeText(item.name) === name);
    if (duplicate) {
      return res.status(409).json({ error: '商品群組名稱已存在' });
    }

    const nextId = productCategoriesData.reduce((max, item) => Math.max(max, Number(item.id) || 0), 0) + 1;
    const group = {
      id: nextId,
      name,
      isVisible: true,
      created_at: new Date().toISOString()
    };

    productCategoriesData.push(group);
    saveProductCategoriesData(productCategoriesData);

    res.status(201).json({ success: true, data: group });
  } catch (err) {
    console.error('新增商品群組錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 商品群組內商品列表（供前端展開群組）
app.get('/api/product-categories/:id(\\d+)/products', (req, res) => {
  try {
    const id = parseInt(req.params.id, 10);
    const group = productCategoriesData.find((item) => item.id === id);
    if (!group) {
      return res.status(404).json({ error: '商品群組不存在' });
    }

    const page = parseInt(req.query.page, 10) || 1;
    const pageSize = parseInt(req.query.pageSize, 10) || 10;
    const search = normalizeText(req.query.search || '').toLowerCase();
    const groupName = normalizeText(group.name);
    const fallbackCategory = normalizeText(productCategoriesData[0]?.name);

    let list = productsData
      .filter((item) => {
        const itemCategory = normalizeText(item.category) || fallbackCategory;
        return itemCategory === groupName;
      })
      .map((item) => ({
        id: item.id,
        name: normalizeText(item.name),
        sku: normalizeText(item.sku),
        type: normalizeText(item.style) || '單一款式',
        store: 'C-FLY家具生活館',
        temperature: normalizeText(item.temperature) || '常溫'
      }));

    if (search) {
      list = list.filter((item) =>
        item.name.toLowerCase().includes(search) ||
        item.sku.toLowerCase().includes(search) ||
        item.type.toLowerCase().includes(search)
      );
    }

    const total = list.length;
    const pages = Math.max(1, Math.ceil(total / pageSize));
    const safePage = Math.min(Math.max(page, 1), pages);
    const offset = (safePage - 1) * pageSize;
    const data = list.slice(offset, offset + pageSize);

    res.json({
      group: {
        id: group.id,
        name: group.name,
        url: `https://cfly777.cyberbiz.co/collections/${encodeURIComponent(group.name)}`
      },
      data,
      pagination: {
        page: safePage,
        pageSize,
        total,
        pages
      }
    });
  } catch (err) {
    console.error('查詢群組商品清單錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 修改商品群組名稱
app.put('/api/product-categories/:id(\\d+)', (req, res) => {
  try {
    const id = parseInt(req.params.id, 10);
    const group = productCategoriesData.find((item) => item.id === id);
    if (!group) {
      return res.status(404).json({ error: '商品群組不存在' });
    }

    const name = normalizeText(req.body.name);
    if (!name) {
      return res.status(400).json({ error: '群組名稱不可為空' });
    }

    const duplicate = productCategoriesData.find((item) => item.id !== id && normalizeText(item.name) === name);
    if (duplicate) {
      return res.status(409).json({ error: '商品群組名稱已存在' });
    }

    const oldName = normalizeText(group.name);
    group.name = name;

    productsData.forEach((item) => {
      if (normalizeText(item.category) === oldName) {
        item.category = name;
      }
    });

    saveProductCategoriesData(productCategoriesData);
    saveProductsData(productsData);

    res.json({ success: true, data: group });
  } catch (err) {
    console.error('修改商品群組錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 刪除商品群組
app.delete('/api/product-categories/:id(\\d+)', (req, res) => {
  try {
    const id = parseInt(req.params.id, 10);
    const index = productCategoriesData.findIndex((item) => item.id === id);
    if (index === -1) {
      return res.status(404).json({ error: '商品群組不存在' });
    }

    if (productCategoriesData.length <= 1) {
      return res.status(400).json({ error: '至少需保留一個商品群組' });
    }

    const removedName = normalizeText(productCategoriesData[index].name);
    productCategoriesData.splice(index, 1);

    const fallback = normalizeText(productCategoriesData[0]?.name) || '未分類';
    productsData.forEach((item) => {
      if (normalizeText(item.category) === removedName) {
        item.category = fallback;
      }
    });

    saveProductCategoriesData(productCategoriesData);
    saveProductsData(productsData);

    res.json({ success: true, message: '商品群組已刪除' });
  } catch (err) {
    console.error('刪除商品群組錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 商品單筆
app.get('/api/products/:id(\\d+)', (req, res) => {
  try {
    const id = parseInt(req.params.id, 10);
    const product = productsData.find((item) => item.id === id);
    if (!product) {
      return res.status(404).json({ error: '商品不存在' });
    }

    res.json({
      data: {
        ...product,
        tag: product.tag || '',
        description: product.description || '',
        urlPath: product.urlPath || product.sku || '',
        searchEnabled: typeof product.searchEnabled === 'boolean' ? product.searchEnabled : true
      }
    });
  } catch (err) {
    console.error('查詢商品錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 新增商品
app.post('/api/products', (req, res) => {
  try {
    const name = normalizeText(req.body.name);
    const style = normalizeText(req.body.style) || '單一款式';
    const sku = normalizeText(req.body.sku);

    if (!name) {
      return res.status(400).json({ error: '商品名稱不可為空' });
    }
    if (!sku) {
      return res.status(400).json({ error: '商品編號(SKU)不可為空' });
    }

    const duplicateSku = productsData.find((item) => normalizeText(item.sku) === sku);
    if (duplicateSku) {
      return res.status(409).json({ error: '商品編號(SKU)已存在' });
    }

    const nextId = productsData.reduce((max, item) => Math.max(max, Number(item.id) || 0), 0) + 1;
    const product = {
      id: nextId,
      name,
      style,
      sku,
      salePrice: Math.max(0, parseInt(req.body.salePrice, 10) || 0),
      listPrice: Math.max(0, parseInt(req.body.listPrice, 10) || 0),
      stock: Math.max(0, parseInt(req.body.stock, 10) || 0),
      sold: Math.max(0, parseInt(req.body.sold, 10) || 0),
      keyword: normalizeText(req.body.keyword),
      status: normalizeText(req.body.status) || '已上架',
      tag: normalizeText(req.body.tag),
      description: normalizeText(req.body.description),
      urlPath: normalizeText(req.body.urlPath) || sku,
      searchEnabled: typeof req.body.searchEnabled === 'boolean' ? req.body.searchEnabled : true
    };

    productsData.push(product);
    saveProductsData(productsData);

    res.status(201).json({ success: true, data: product });
  } catch (err) {
    console.error('新增商品錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 更新商品
app.put('/api/products/:id(\\d+)', (req, res) => {
  try {
    const id = parseInt(req.params.id, 10);
    const product = productsData.find((item) => item.id === id);
    if (!product) {
      return res.status(404).json({ error: '商品不存在' });
    }

    const name = normalizeText(req.body.name);
    const style = normalizeText(req.body.style);
    const sku = normalizeText(req.body.sku);

    if (!name) {
      return res.status(400).json({ error: '商品名稱不可為空' });
    }
    if (!sku) {
      return res.status(400).json({ error: '商品編號(SKU)不可為空' });
    }

    const duplicateSku = productsData.find((item) => item.id !== id && normalizeText(item.sku) === sku);
    if (duplicateSku) {
      return res.status(409).json({ error: '商品編號(SKU)已存在' });
    }

    product.name = name;
    product.style = style || '單一款式';
    product.sku = sku;
    product.salePrice = Math.max(0, parseInt(req.body.salePrice, 10) || 0);
    product.listPrice = Math.max(0, parseInt(req.body.listPrice, 10) || 0);
    product.stock = Math.max(0, parseInt(req.body.stock, 10) || 0);
    product.sold = Math.max(0, parseInt(req.body.sold, 10) || 0);
    product.keyword = normalizeText(req.body.keyword);
    product.status = normalizeText(req.body.status) || '已上架';
    product.tag = normalizeText(req.body.tag);
    product.description = normalizeText(req.body.description);
    product.urlPath = normalizeText(req.body.urlPath) || sku;
    product.searchEnabled = typeof req.body.searchEnabled === 'boolean' ? req.body.searchEnabled : true;

    saveProductsData(productsData);
    res.json({ success: true, data: product });
  } catch (err) {
    console.error('更新商品錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 刪除商品
app.delete('/api/products/:id(\\d+)', (req, res) => {
  try {
    const id = parseInt(req.params.id, 10);
    const index = productsData.findIndex((item) => item.id === id);
    if (index === -1) {
      return res.status(404).json({ error: '商品不存在' });
    }

    productsData.splice(index, 1);
    saveProductsData(productsData);
    res.json({ success: true, message: '商品已刪除' });
  } catch (err) {
    console.error('刪除商品錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 會員列表
app.get('/api/members', (req, res) => {
  try {
    const page = parseInt(req.query.page, 10) || 1;
    const pageSize = parseInt(req.query.pageSize, 10) || 20;
    const search = normalizeText(req.query.search || '').toLowerCase();

    let list = db.members.map((member) => {
      const relatedOrders = db.orders.filter((o) => o.member_id === member.id);
      const totalAmount = relatedOrders.reduce((sum, o) => sum + (parseFloat(o.amount || 0) || 0), 0);
      const lastOrderDate = relatedOrders
        .map((o) => o.date || o.created_at || '')
        .filter(Boolean)
        .sort((a, b) => String(b).localeCompare(String(a)))[0] || '';

      return {
        ...member,
        orderCount: relatedOrders.length,
        totalAmount,
        lastOrderDate
      };
    });

    if (search) {
      list = list.filter((m) =>
        normalizeText(m.name).toLowerCase().includes(search) ||
        normalizeText(m.phone).toLowerCase().includes(search) ||
        normalizeText(m.email).toLowerCase().includes(search)
      );
    }

    list.sort((a, b) => String(b.lastOrderDate || '').localeCompare(String(a.lastOrderDate || '')));

    const total = list.length;
    const offset = (page - 1) * pageSize;
    const data = list.slice(offset, offset + pageSize);

    res.json({
      data,
      pagination: {
        page,
        pageSize,
        total,
        pages: Math.ceil(total / pageSize)
      }
    });
  } catch (err) {
    console.error('查詢會員列表錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 建立會員
app.post('/api/members', (req, res) => {
  try {
    const name = normalizeText(req.body.name);
    const phone = normalizeText(req.body.phone);
    const email = normalizeText(req.body.email);

    if (!name && !phone && !email) {
      return res.status(400).json({ error: '請至少填寫姓名、電話或信箱其中一項' });
    }

    const existing = findMemberByIdentity(name, phone, email);
    if (existing) {
      return res.status(409).json({ error: '會員已存在' });
    }

    const member = {
      id: db.nextMemberId++,
      name,
      phone,
      email,
      created_at: new Date().toISOString()
    };

    db.members.push(member);
    saveData(db);
    res.status(201).json({ success: true, data: member });
  } catch (err) {
    console.error('建立會員錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 更新會員
app.put('/api/members/:id', (req, res) => {
  try {
    const id = parseInt(req.params.id, 10);
    const member = db.members.find((m) => m.id === id);
    if (!member) {
      return res.status(404).json({ error: '會員不存在' });
    }

    const name = normalizeText(req.body.name);
    const phone = normalizeText(req.body.phone);
    const email = normalizeText(req.body.email);

    if (!name && !phone && !email) {
      return res.status(400).json({ error: '請至少填寫姓名、電話或信箱其中一項' });
    }

    const duplicate = db.members.find((m) => m.id !== id && buildMemberKey(m.name, m.phone, m.email) === buildMemberKey(name, phone, email));
    if (duplicate) {
      return res.status(409).json({ error: '已存在相同資料的會員' });
    }

    member.name = name;
    member.phone = phone;
    member.email = email;

    db.orders.forEach((order) => {
      if (order.member_id === id) {
        order.name = name;
        order.phone = phone;
        order.email = email;
      }
    });

    saveData(db);
    res.json({ success: true, data: member });
  } catch (err) {
    console.error('更新會員錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 會員訂單明細
app.get('/api/members/:id/orders', (req, res) => {
  try {
    const id = parseInt(req.params.id, 10);
    const member = db.members.find((m) => m.id === id);
    if (!member) {
      return res.status(404).json({ error: '會員不存在' });
    }

    const orders = db.orders
      .filter((o) => o.member_id === id)
      .sort((a, b) => String(b.date || b.created_at || '').localeCompare(String(a.date || a.created_at || '')));

    res.json({ data: orders });
  } catch (err) {
    console.error('查詢會員訂單明細錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 刪除訂單
app.delete('/api/orders/:id(\\d+)', (req, res) => {
  try {
    const id = parseInt(req.params.id);
    const index = db.orders.findIndex(order => order.id === id);

    if (index === -1) {
      return res.status(404).json({ error: '訂單不存在' });
    }

    db.orders.splice(index, 1);
    saveData(db);

    res.json({ success: true, message: '訂單已刪除' });
  } catch (err) {
    console.error('刪除錯誤:', err);
    res.status(500).json({ error: err.message });
  }
});

// 健康檢查
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', message: 'Server is running' });
});

// 啟動服務器
app.listen(PORT, () => {
  console.log(`✅ 後端服務器運行於 http://localhost:${PORT}`);
  console.log(`📊 數據文件路徑: ${dbFile}`);
  console.log(`📥 上傳目錄: ${uploadDir}`);
});

process.on('SIGINT', () => {
  saveData(db);
  process.exit();
});
