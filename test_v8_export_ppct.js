const fs = require('fs');
const path = require('path');
const assert = require('assert');

console.log('=== BẮT ĐẦU KIỂM TRA TÍNH NĂNG XUẤT EXCEL PPCT & HƯỚNG DẪN V8 ===\n');

const appDir = __dirname;
const indexHtml = fs.readFileSync(path.join(appDir, 'index.html'), 'utf8');
const appJs = fs.readFileSync(path.join(appDir, 'js/app.js'), 'utf8');
const guideV8 = fs.readFileSync(path.join(appDir, 'Huong_Dan_Su_Dung_V8.txt'), 'utf8');
const detailGuide = fs.readFileSync(path.join(appDir, 'Huong_Dan_Chi_Tiet_V8.txt'), 'utf8');

let passCount = 0;
function test(name, condition) {
    if (condition) {
        console.log(`✅ PASS: ${name}`);
        passCount++;
    } else {
        console.error(`❌ FAIL: ${name}`);
        process.exit(1);
    }
}

// 1. Check index.html has btn-export-ppct-excel
test('index.html có nút #btn-export-ppct-excel', indexHtml.includes('id="btn-export-ppct-excel"'));
test('Nút #btn-export-ppct-excel có nhãn "Xuất Excel PPCT"', indexHtml.includes('Xuất Excel PPCT'));

// 2. Check app.js defines exportPpctToExcel
test('app.js định nghĩa hàm exportPpctToExcel', appJs.includes('function exportPpctToExcel('));
test('app.js gán sự kiện click cho btn-export-ppct-excel', appJs.includes('document.getElementById("btn-export-ppct-excel").addEventListener("click", exportPpctToExcel);'));

// 3. Check export structure in exportPpctToExcel
test('exportPpctToExcel xuất đầy đủ 6 cột dữ liệu', 
    appJs.includes('["Tuần", "Môn học", "Tiết/Tuần", "Tiết PPCT", "Tên bài dạy", "Nội dung tích hợp / Điều chỉnh"]'));
test('exportPpctToExcel bao gồm cột integration', appJs.includes('p.integration || ""'));
test('exportPpctToExcel có bộ lọc môn học', appJs.includes('filterSubject'));
test('exportPpctToExcel có bộ lọc tuần', appJs.includes('filterWeek'));
test('exportPpctToExcel có tìm kiếm từ khóa', appJs.includes('searchInput'));

// 4. Check Tab 7 Guide in index.html
test('index.html có mục ⭐ QUY TRÌNH 4 BƯỚC THIẾT THỰC', indexHtml.includes('QUY TRÌNH 4 BƯỚC THIẾT THỰC ĐỂ BẮT ĐẦU SỬ DỤNG PHẦN MỀM HIỆU QUẢ'));
test('index.html Bước 3 có câu "Kiểm tra kho bài dạy đã nạp sẵn. Chỉnh sửa, thêm bớt cho phù hợp với thực tế."', 
    indexHtml.includes('Kiểm tra kho bài dạy đã nạp sẵn. Chỉnh sửa, thêm bớt cho phù hợp với thực tế.'));
test('index.html Tab Guide hướng dẫn nút Xuất Excel PPCT', indexHtml.includes('Xuất Excel PPCT (V8 Mới)'));
test('index.html Tab Guide có mục Điểm Mới V8 với Xuất Excel PPCT', indexHtml.includes('Xuất Excel PPCT Toàn Diện'));

// 5. Check Huong_Dan_Chi_Tiet_V8.txt
test('Huong_Dan_Chi_Tiet_V8.txt tồn tại và có đầy đủ mục lục A-F', 
    detailGuide.includes('PHẦN A.') && detailGuide.includes('PHẦN B.') && 
    detailGuide.includes('PHẦN C.') && detailGuide.includes('PHẦN D.') && 
    detailGuide.includes('PHẦN E.') && detailGuide.includes('PHẦN F.'));
test('Huong_Dan_Chi_Tiet_V8.txt hướng dẫn nút Xuất Excel PPCT', detailGuide.includes('D9. XUẤT FILE EXCEL PHÂN PHỐI CHƯƠNG TRÌNH (PPCT)'));

// 6. Check Huong_Dan_Su_Dung_V8.txt
test('Huong_Dan_Su_Dung_V8.txt ghi nhận tính năng Xuất Excel PPCT tại mục 9', guideV8.includes('9. XUẤT FILE EXCEL PHÂN PHỐI CHƯƠNG TRÌNH (PPCT) TOÀN DIỆN:'));

console.log(`\n🎉 TẤT CẢ ${passCount} BÀI KIỂM TRA ĐỀU VƯỢT QUA XUẤT SẮC! (100% PASS)`);
