const fs = require('fs');
const XLSX = require('xlsx');

const USERNAME = process.env.USERNAME;
const PASSWORD = process.env.PASSWORD;
// 🌟 這裡使用 CHAIN_CODE 讀取 GitHub Secret，但存進名稱為 chainCode 的變數中，完美對齊 API 文件
const chainCode = process.env.chainCod; 
const PLATFORM_KEY = process.env.PLATFORM_KEY; 

const ERROR_THRESHOLD = 100; 
let errorCount = 0;
let successCount = 0; 

async function getToken() {
    if (!USERNAME || !PASSWORD) {
        console.log("❌ 嚴重錯誤：GitHub Secrets 的 USERNAME 或 PASSWORD 沒有正確讀取！");
        process.exit(1); 
    }

    const params = new URLSearchParams();
    params.append('grant_type', 'client_credentials');
    params.append('username', USERNAME); 
    params.append('password', PASSWORD); 

    try {
        // 🌟 登入網址
        const res = await fetch('https://integration-middleware.as.restaurant-partners.com/v2/login', {
            method: 'POST', 
            body: params,
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        const data = await res.json();
        if (data.access_token) return data.access_token;
        
        console.log("❌ 取得 Token 失敗，系統回傳：", data);
        process.exit(1); 
    } catch (err) {
        console.log("❌ 請求 Token 時發生網路錯誤：", err.message);
        process.exit(1);
    }
}

async function updateVendor(token, posVendorId) { // 🌟 變數也對齊文件改成 posVendorId
    // 🌟 網址修正為 STG，並且變數直接使用 ${chainCode} 和 ${posVendorId}，跟官方文件一模一樣！
    const url = `https://integration-middleware.as.restaurant-partners.com/v2/chains/${chainCode}/remoteVendors/${posVendorId}/availability`;
    try {
        const res = await fetch(url, {
            method: 'PUT',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify({ 
                availabilityState: "OPEN",
                platformKey: PLATFORM_KEY,
                platformRestaurantId: posVendorId // 🌟 對應店家編號
            })
        });
        
        if (res.ok) {
            console.log(`✅ [${posVendorId}] 更新成功`);
            successCount++; 
        } else {
            console.log(`❌ [${posVendorId}] 更新失敗: ${await res.text()}`);
            errorCount++; 
        }
    } catch (err) {
        console.log(`❌ [${posVendorId}] 網路錯誤: ${err.message}`);
        errorCount++; 
    }
}

async function main() {
    const token = await getToken();
    const workbook = XLSX.readFile('Vendor711Resto.xlsx');
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    let vendors = [];
    for (const row of rawData) {
        const id = row[0] ? row[0].toString().trim() : '';
        if (id && id.toLowerCase() !== 'vendor_id') vendors.push(id);
    }

    console.log(`共讀取到 ${vendors.length} 家店家，準備開始執行...`);

    for (let i = 0; i < vendors.length; i++) {
        const vid = vendors[i];
        await updateVendor(token, vid);
        await new Promise(r => setTimeout(r, 500)); 
    }
    
    if (process.env.GITHUB_OUTPUT) {
        fs.appendFileSync(process.env.GITHUB_OUTPUT, `success_count=${successCount}\n`);
        fs.appendFileSync(process.env.GITHUB_OUTPUT, `error_count=${errorCount}\n`);
    }

    if (errorCount > ERROR_THRESHOLD) {
        console.log(`\n🚨 警告：共有 ${errorCount} 家店鋪更新失敗，已超過容忍值！`);
        process.exit(1); 
    } else {
        console.log(`\n🎉 執行完畢！成功: ${successCount}, 失敗: ${errorCount}`);
        process.exit(0);
    }
}

main();
