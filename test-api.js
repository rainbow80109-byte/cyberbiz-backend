/**
 * 簡單測試 API 端點
 * 運行: node test-api.js
 */

const http = require('http');

// 測試 API 列表
const tests = [
  { path: '/api/health', method: 'GET', name: '健康檢查' },
  { path: '/api/statistics', method: 'GET', name: '獲取統計數據' },
  { path: '/api/orders?page=1&pageSize=10', method: 'GET', name: '獲取訂單列表' },
];

const API_HOST = 'localhost';
const API_PORT = 3000;

function makeRequest(path, method) {
  return new Promise((resolve, reject) => {
    const options = {
      hostname: API_HOST,
      port: API_PORT,
      path: path,
      method: method,
    };

    const req = http.request(options, (res) => {
      let data = '';
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', () => {
        resolve({
          status: res.statusCode,
          data: data ? JSON.parse(data) : null
        });
      });
    });

    req.on('error', reject);
    req.setTimeout(5000);
    req.end();
  });
}

async function runTests() {
  console.log('🧪 API 測試開始...\n');
  
  for (const test of tests) {
    try {
      console.log(`⏳ 測試: ${test.name} (${test.method} ${test.path})`);
      const result = await makeRequest(test.path, test.method);
      
      if (result.status === 200) {
        console.log(`✅ 成功 (狀態碼: ${result.status})`);
        console.log(`📊 響應數據:`, result.data);
      } else {
        console.log(`⚠️ 狀態碼: ${result.status}`);
      }
    } catch (err) {
      console.log(`❌ 失敗: ${err.message}`);
    }
    console.log('');
  }

  console.log('✨ 測試完成！');
}

// 確保伺服器運行中
console.log('📡 連接到', `http://${API_HOST}:${API_PORT}`);
console.log('⚠️ 請確保服務器正在運行 (npm start)\n');

setTimeout(runTests, 1000);
