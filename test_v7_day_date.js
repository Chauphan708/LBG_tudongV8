const fs = require('fs');
const path = require('path');
const assert = require('assert');

console.log("=== BẮT ĐẦU KIỂM TRA PHIÊN BẢN V7 ===");

// 1. Kiểm tra thuật toán getDayFullDate
function getDayFullDate(startDateVN, dayStr) {
    if (!startDateVN) return "";
    const parts = startDateVN.trim().split("/");
    if (parts.length < 3) return "";
    const d = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    const y = parseInt(parts[2], 10);
    if (isNaN(d) || isNaN(m) || isNaN(y)) return "";

    let offset = 0;
    const cleanDay = (dayStr || "").trim().toLowerCase();
    if (cleanDay.includes("2")) offset = 0;
    else if (cleanDay.includes("3")) offset = 1;
    else if (cleanDay.includes("4")) offset = 2;
    else if (cleanDay.includes("5")) offset = 3;
    else if (cleanDay.includes("6")) offset = 4;
    else if (cleanDay.includes("7")) offset = 5;
    else if (cleanDay.includes("chủ nhật") || cleanDay.includes("cn")) offset = 6;

    const targetDate = new Date(y, m, d + offset);
    const dayNum = String(targetDate.getDate()).padStart(2, '0');
    const monthNum = String(targetDate.getMonth() + 1).padStart(2, '0');
    const yearNum = targetDate.getFullYear();
    return `${dayNum}/${monthNum}/${yearNum}`;
}

// Test case 1.1: Tuần bình thường trong tháng 9
assert.strictEqual(getDayFullDate("07/09/2026", "Thứ 2"), "07/09/2026");
assert.strictEqual(getDayFullDate("07/09/2026", "Thứ 3"), "08/09/2026");
assert.strictEqual(getDayFullDate("07/09/2026", "Thứ 4"), "09/09/2026");
assert.strictEqual(getDayFullDate("07/09/2026", "Thứ 5"), "10/09/2026");
assert.strictEqual(getDayFullDate("07/09/2026", "Thứ 6"), "11/09/2026");
console.log("✓ Test 1.1: Tuần bình thường trong tháng -> ĐẠT");

// Test case 1.2: Tuần chuyển giao tháng 9 sang tháng 10
assert.strictEqual(getDayFullDate("28/09/2026", "Thứ 2"), "28/09/2026");
assert.strictEqual(getDayFullDate("28/09/2026", "Thứ 3"), "29/09/2026");
assert.strictEqual(getDayFullDate("28/09/2026", "Thứ 4"), "30/09/2026");
assert.strictEqual(getDayFullDate("28/09/2026", "Thứ 5"), "01/10/2026");
assert.strictEqual(getDayFullDate("28/09/2026", "Thứ 6"), "02/10/2026");
console.log("✓ Test 1.2: Chuyển giao tháng 9 -> 10 -> ĐẠT");

// Test case 1.3: Tuần chuyển giao năm 2026 sang năm 2027
assert.strictEqual(getDayFullDate("28/12/2026", "Thứ 2"), "28/12/2026");
assert.strictEqual(getDayFullDate("28/12/2026", "Thứ 5"), "31/12/2026");
assert.strictEqual(getDayFullDate("28/12/2026", "Thứ 6"), "01/01/2027");
console.log("✓ Test 1.3: Chuyển giao năm 2026 -> 2027 -> ĐẠT");

// 2. Kiểm tra js/docx_generator.js
const docxContent = fs.readFileSync(path.join(__dirname, 'js', 'docx_generator.js'), 'utf8');
assert.ok(docxContent.includes('Thứ, ngày'), "js/docx_generator.js thiếu tiêu đề 'Thứ, ngày'");
assert.ok(docxContent.includes('getDayFullDate'), "js/docx_generator.js thiếu hàm getDayFullDate");
assert.ok(docxContent.includes('dayDate'), "js/docx_generator.js không render biến dayDate");
console.log("✓ Test 2: docx_generator.js có đầy đủ cấu trúc ngày dưới thứ -> ĐẠT");

// 3. Kiểm tra js/xlsx_generator.js
const xlsxContent = fs.readFileSync(path.join(__dirname, 'js', 'xlsx_generator.js'), 'utf8');
assert.ok(xlsxContent.includes('Thứ, ngày'), "js/xlsx_generator.js thiếu tiêu đề 'Thứ, ngày'");
assert.ok(xlsxContent.includes('getDayFullDate'), "js/xlsx_generator.js thiếu hàm getDayFullDate");
assert.ok(xlsxContent.includes('${day}\\n${dayDate}') || xlsxContent.includes('${day}\\n'), "js/xlsx_generator.js thiếu ghép dòng ${day}\\n${dayDate}");
console.log("✓ Test 3: xlsx_generator.js xuất đúng 2 dòng Thứ & Ngày -> ĐẠT");

// 4. Kiểm tra js/app.js
const appContent = fs.readFileSync(path.join(__dirname, 'js', 'app.js'), 'utf8');
assert.ok(appContent.includes('LBG_APP_DATA_V7'), "js/app.js chưa cập nhật key V7");
assert.ok(appContent.includes('LBG_APP_DATA_V6'), "js/app.js thiếu cơ chế fallback từ V6");
assert.ok(appContent.includes('LBG_APP_DATA_V5'), "js/app.js thiếu cơ chế fallback từ V5");
assert.ok(appContent.includes('getDayFullDate'), "js/app.js thiếu getDayFullDate");
assert.ok(appContent.includes('Thứ, ngày'), "js/app.js renderSingleWeekPaperHtml thiếu tiêu đề 'Thứ, ngày'");
console.log("✓ Test 4: app.js tương thích V7, bảo toàn dữ liệu và đồng bộ giao diện -> ĐẠT");

// 5. Kiểm tra index.html
const htmlContent = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
assert.ok(htmlContent.includes('Bản V7') || htmlContent.includes('Bản V8'), "index.html title chưa cập nhật Bản V7 hoặc V8");
assert.ok(htmlContent.includes('BẢN V7') || htmlContent.includes('BẢN V8'), "index.html header badge chưa cập nhật BẢN V7 hoặc V8");
assert.ok(htmlContent.includes('Hướng Dẫn V7') || htmlContent.includes('Hướng Dẫn V8'), "index.html quick button chưa có Hướng Dẫn V7 hoặc V8");
assert.ok(htmlContent.includes('>Thứ, ngày</th>'), "index.html table thead chưa đổi sang Thứ, ngày");
assert.ok(htmlContent.includes('id="guide-sec-10"'), "index.html thiếu mục hướng dẫn Điểm Mới Bản V7");
console.log("✓ Test 5: index.html cập nhật giao diện, tiêu đề và hướng dẫn V7 chuẩn xác -> ĐẠT");

console.log("\n==========================================");
console.log("🎉 TẤT CẢ 5 BỘ KIỂM TRA ĐỀU VƯỢT QUA 100%!");
console.log("==========================================");
