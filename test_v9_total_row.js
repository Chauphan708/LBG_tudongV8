const fs = require('fs');
const path = require('path');
const assert = require('assert');
const JSZip = require('./lib/jszip.min.js');

console.log("=== BẮT ĐẦU KIỂM THỬ TÍNH NĂNG V9: HÀNG TỔNG SỐ TIẾT/TUẦN ===");

// 1. Kiểm tra mã nguồn app.js
const appJsContent = fs.readFileSync(path.join(__dirname, 'js', 'app.js'), 'utf8');
assert.ok(appJsContent.includes('Tổng số tiết/tuần'), "app.js thiếu text 'Tổng số tiết/tuần'");
assert.ok(appJsContent.includes('lbg-total-periods-row'), "app.js thiếu class 'lbg-total-periods-row'");
assert.ok(appJsContent.includes('tiết/ tuần'), "app.js thiếu text 'tiết/ tuần'");
console.log("✔ Kiểm tra 1: app.js đã tích hợp dòng Tổng số tiết/tuần trên Web và Bản In Giấy.");

// 2. Kiểm tra mã nguồn docx_generator.js
const docxGenContent = fs.readFileSync(path.join(__dirname, 'js', 'docx_generator.js'), 'utf8');
assert.ok(docxGenContent.includes('Tổng số tiết/tuần'), "docx_generator.js thiếu text 'Tổng số tiết/tuần'");
assert.ok(docxGenContent.includes('gridSpan w:val="3"'), "docx_generator.js thiếu gridSpan 3 cho ô Tổng số tiết/tuần");
console.log("✔ Kiểm tra 2: docx_generator.js đã tích hợp hàng Tổng số tiết/tuần với gridSpan chuẩn.");

// 3. Kiểm tra mã nguồn xlsx_generator.js
const xlsxGenContent = fs.readFileSync(path.join(__dirname, 'js', 'xlsx_generator.js'), 'utf8');
assert.ok(xlsxGenContent.includes('Tổng số tiết/tuần'), "xlsx_generator.js thiếu text 'Tổng số tiết/tuần'");
assert.ok(xlsxGenContent.includes('mergeCellsList.push(`A${currentRow}:C${currentRow}`)'), "xlsx_generator.js thiếu merge cell A:C");
console.log("✔ Kiểm tra 3: xlsx_generator.js đã tích hợp hàng Tổng số tiết/tuần và merge cells.");

// 4. Kiểm tra xuất file Word thực tế
global.window = {
    JSZip: JSZip
};
global.JSZip = JSZip;
global.escapeXml = (str) => String(str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// Nạp docx_generator.js
eval(docxGenContent);

const mockWeekInfo = {
    week: 1,
    startDateVN: '07/09/2026',
    endDateVN: '11/09/2026',
    month: 'Tháng 9 năm 2026'
};

const mockSchedule = [
    { day: "Thứ 2", session: "Sáng", period: 1, subject: "Chào cờ", ppct: 1, lessonName: "Chào cờ đầu tuần", isOff: false },
    { day: "Thứ 2", session: "Sáng", period: 2, subject: "Toán", ppct: 1, lessonName: "Ôn tập khái niệm về phân số", isOff: false },
    { day: "Thứ 2", session: "Sáng", period: 3, subject: "Tiếng Việt", ppct: 1, lessonName: "Thư gửi các học sinh", isOff: false }
];

const mockSettings = {
    governingBody: "UBND Phường Trung Nhứt",
    schoolName: "Trường Tiểu học Trung Nhứt",
    grade: "Khối 5",
    className: "Lớp 5A",
    principal: "Phạm Quốc Hùng",
    homeroomTeacher: "Nguyễn Thị Thu Hà",
    headOfGrade: "Trần Thị Mai"
};

const mockStats = { gvcnCount: 2, specialistCount: 1, enhancedCount: 0, total: 3 };

async function getBuffer(blob) {
    if (blob && typeof blob.arrayBuffer === 'function') {
        return Buffer.from(await blob.arrayBuffer());
    }
    return blob;
}

async function testDocxGeneration() {
    // Test bật dòng tổng số tiết (mặc định)
    const blobOn = await window.DocxGenerator.generateLbgDocx(
        false, 1, mockWeekInfo, mockSchedule, mockSettings, mockStats, "portrait", { showTotalRow: true }
    );
    assert.ok(blobOn, "Không sinh được blob Word khi bật");
    const zipOn = await JSZip.loadAsync(await getBuffer(blobOn));
    const docXmlOn = await zipOn.file("word/document.xml").async("text");
    
    assert.ok(docXmlOn.includes("Tổng số tiết/tuần"), "File Word document.xml thiếu 'Tổng số tiết/tuần'");
    assert.ok(docXmlOn.includes("3 tiết/ tuần"), "File Word document.xml thiếu '3 tiết/ tuần'");
    assert.ok(docXmlOn.includes('w:gridSpan w:val="3"'), "File Word document.xml thiếu w:gridSpan 3");

    // Test tắt dòng tổng số tiết (showTotalRow = false)
    const blobOff = await window.DocxGenerator.generateLbgDocx(
        false, 1, mockWeekInfo, mockSchedule, mockSettings, mockStats, "portrait", { showTotalRow: false }
    );
    assert.ok(blobOff, "Không sinh được blob Word khi tắt");
    const zipOff = await JSZip.loadAsync(await getBuffer(blobOff));
    const docXmlOff = await zipOff.file("word/document.xml").async("text");
    assert.ok(!docXmlOff.includes("Tổng số tiết/tuần"), "File Word document.xml vẫn còn 'Tổng số tiết/tuần' khi showTotalRow = false");

    console.log("✔ Kiểm tra 4: Đã sinh file DOCX thực tế và xác nhận tùy chọn showTotalRow hoạt động chính xác (bật/tắt)!");
}

// 5. Kiểm tra xuất file Excel thực tế
eval(xlsxGenContent);

async function testXlsxGeneration() {
    // Test bật dòng tổng số tiết
    const blobOn = await window.XlsxGenerator.generateLbgXlsx({
        settings: mockSettings,
        weekNum: 1,
        weekInfo: mockWeekInfo,
        schedule: mockSchedule,
        stats: mockStats,
        isCtlop: false,
        orientation: "portrait",
        showColSign: false,
        showColNote: true,
        showTotalRow: true
    });
    assert.ok(blobOn, "Không sinh được blob Excel khi bật");
    const zipOn = await JSZip.loadAsync(await getBuffer(blobOn));
    const sheetXmlOn = await zipOn.file("xl/worksheets/sheet1.xml").async("text");

    assert.ok(sheetXmlOn.includes("Tổng số tiết/tuần"), "File Excel sheet1.xml thiếu 'Tổng số tiết/tuần'");
    assert.ok(sheetXmlOn.includes("3 tiết/ tuần"), "File Excel sheet1.xml thiếu '3 tiết/ tuần'");
    assert.ok(sheetXmlOn.includes("A") && sheetXmlOn.includes(":C"), "File Excel sheet1.xml thiếu merge A:C");

    // Test tắt dòng tổng số tiết
    const blobOff = await window.XlsxGenerator.generateLbgXlsx({
        settings: mockSettings,
        weekNum: 1,
        weekInfo: mockWeekInfo,
        schedule: mockSchedule,
        stats: mockStats,
        isCtlop: false,
        orientation: "portrait",
        showColSign: false,
        showColNote: true,
        showTotalRow: false
    });
    assert.ok(blobOff, "Không sinh được blob Excel khi tắt");
    const zipOff = await JSZip.loadAsync(await getBuffer(blobOff));
    const sheetXmlOff = await zipOff.file("xl/worksheets/sheet1.xml").async("text");
    assert.ok(!sheetXmlOff.includes("Tổng số tiết/tuần"), "File Excel sheet1.xml vẫn còn 'Tổng số tiết/tuần' khi showTotalRow = false");

    console.log("✔ Kiểm tra 5: Đã sinh file XLSX thực tế và xác nhận tùy chọn showTotalRow hoạt động chính xác (bật/tắt)!");
}

async function run() {
    await testDocxGeneration();
    await testXlsxGeneration();
    console.log("\n🎉 TOÀN BỘ 5 BƯỚC KIỂM THỬ V9 ĐỀU ĐẠT 100%!");
}

run().catch(err => {
    console.error("❌ LỖI KIỂM THỬ:", err);
    process.exit(1);
});
