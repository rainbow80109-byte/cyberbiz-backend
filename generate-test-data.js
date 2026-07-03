/**
 * 生成測試 Excel 文件
 * 運行: node generate-test-data.js
 */
const XLSX = require('xlsx');
const fs = require('fs');

// 生成測試數據
const testData = [
  { 發票: 'INV-001', 金額: 1500.50, 名稱: '張三', 日期: '2026-05-01', 備註: '首單' },
  { 發票: 'INV-002', 金額: 2000.00, 名稱: '李四', 日期: '2026-05-02', 備註: '續購' },
  { 發票: 'INV-003', 金額: 800.00, 名稱: '王五', 日期: '2026-05-03', 備註: '' },
  { 發票: 'INV-004', 金額: 3500.00, 名稱: '張三', 日期: '2026-05-05', 備註: '大單' },
  { 發票: 'INV-005', 金額: 1200.00, 名稱: '趙六', 日期: '2026-05-06', 備註: '' },
  { 發票: 'INV-006', 金額: 2500.50, 名稱: '李四', 日期: '2026-05-08', 備註: '續購' },
  { 發票: 'INV-007', 金額: 1800.00, 名稱: '孫七', 日期: '2026-05-10', 備註: '' },
  { 發票: 'INV-008', 金額: 4200.00, 名稱: '張三', 日期: '2026-05-12', 備註: '大單' },
  { 發票: 'INV-009', 金額: 950.00, 名稱: '周八', 日期: '2026-05-15', 備註: '小單' },
  { 發票: 'INV-010', 金額: 3100.00, 名稱: '吳九', 日期: '2026-05-18', 備註: '' },
];

// 創建 workbook
const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.json_to_sheet(testData);
XLSX.utils.book_append_sheet(workbook, worksheet, 'Orders');

// 寫入文件
const testDir = './test-data';
if (!fs.existsSync(testDir)) {
  fs.mkdirSync(testDir, { recursive: true });
}

XLSX.writeFile(workbook, `${testDir}/sample-orders.xlsx`);
console.log('✅ 測試 Excel 文件已生成: test-data/sample-orders.xlsx');
