// test_v8_ctlop_sync.js
const fs = require('fs');
const path = require('path');

console.log("=== BẮT ĐẦU KIỂM TRA ĐỒNG BỘ CTLOP & ĐỊNH DẠNG TÊN BÀI DẠY ===");

let errors = 0;

// 1. Kiểm tra CSS
const cssContent = fs.readFileSync(path.join(__dirname, 'css/styles.css'), 'utf8');
if (cssContent.includes('padding-left: 1pt') && cssContent.includes('text-indent: 0.5pt')) {
    console.log("✔ [CSS] styles.css chứa quy tắc padding-left: 1pt và text-indent: 0.5pt");
} else {
    console.error("❌ [CSS] Thiếu padding-left: 1pt hoặc text-indent: 0.5pt trong styles.css");
    errors++;
}

// 2. Kiểm tra HTML Tab 2
const htmlContent = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
const requiredIds = [
    'ctlop-opt-col-sign',
    'ctlop-opt-col-note',
    'ctlop-custom-cols-list',
    'btn-ctlop-add-custom-col',
    'ctlop-opt-sig-bgh',
    'ctlop-opt-sig-gvcn',
    'ctlop-opt-sig-head',
    'ctlop-sig-box-bgh',
    'ctlop-sig-box-head',
    'ctlop-sig-box-gvcn'
];

requiredIds.forEach(id => {
    if (htmlContent.includes(`id="${id}"`)) {
        console.log(`✔ [HTML] index.html chứa phần tử #${id}`);
    } else {
        console.error(`❌ [HTML] index.html thiếu phần tử #${id}`);
        errors++;
    }
});

// 3. Kiểm tra JS DOCX Generator logic
const docxContent = fs.readFileSync(path.join(__dirname, 'js/docx_generator.js'), 'utf8');
if (docxContent.includes('w:left w:w="20"') && docxContent.includes('w:ind w:firstLine="10"')) {
    console.log("✔ [DOCX] docx_generator.js có lề trái 1pt (w:w=20) và thụt đầu dòng 0.5pt (w:firstLine=10)");
} else {
    console.error("❌ [DOCX] docx_generator.js thiếu w:left w:w=20 hoặc w:ind w:firstLine=10");
    errors++;
}

// 4. Test thực thi sinh DOCX CTLOP với custom columns
const JSZip = require(path.join(__dirname, 'lib/jszip.min.js'));

// Giả lập môi trường browser tối thiểu cho DocxGenerator
global.JSZip = JSZip;
global.window = { JSZip: JSZip };
require(path.join(__dirname, 'js/docx_generator.js'));

async function testDocxGeneration() {
    try {
        const weekInfo = {
            startDateVN: "08/09/2026",
            endDateVN: "12/09/2026"
        };
        const schedule = [
            {
                day: "Thứ 2", session: "Sáng", period: 1, subject: "Tiếng Việt", ppct: "1",
                lessonName: "Bài 1: Khởi đầu mới cho năm học", integration: "Tích hợp QCN",
                customCols: { col_test: "Đã chuẩn bị ĐDDH" },
                note: "Lưu ý bài tập 2"
            }
        ];
        const settings = {
            governingBody: "PHÒNG GD&ĐT QUẬN BÌNH THỦY",
            schoolName: "TRƯỜNG TIỂU HỌC TRUNG NHỨT",
            className: "5A",
            grade: "Khối 5",
            homeroomTeacher: "Nguyễn Thị Thu Hà",
            headOfGrade: "Trần Thị Mai",
            vicePrincipal: "Phan Văn Minh"
        };
        const stats = {};
        const options = {
            showColSign: true,
            showColNote: true,
            showColCustom: true,
            customCols: [
                { id: "col_test", name: "Đồ dùng DH", pos: "end", enabled: true }
            ],
            showBghSign: true,
            showGvcnSign: true,
            showHeadSign: true
        };

        const blob = await window.DocxGenerator.generateLbgDocx(
            true, // isCtlop = true
            1,
            weekInfo,
            schedule,
            settings,
            stats,
            "portrait",
            options
        );

        // Đọc blob/buffer giải nén XML
        let inputData = blob;
        if (blob && typeof blob.arrayBuffer === 'function') {
            inputData = Buffer.from(await blob.arrayBuffer());
        }
        const zip = await JSZip.loadAsync(inputData);
        const docXml = await zip.file("word/document.xml").async("string");

        // Kiểm tra XML
        if (docXml.includes("LỊCH BÁO GIẢNG TÍCH HỢP TUẦN 1")) {
            console.log("✔ [DOCX Test] Tiêu đề LỊCH BÁO GIẢNG TÍCH HỢP TUẦN 1 có trong document.xml");
        } else {
            console.error("❌ [DOCX Test] Không tìm thấy tiêu đề CTLOP trong docx");
            errors++;
        }

        if (docXml.includes("Đồ dùng DH") && docXml.includes("Đã chuẩn bị ĐDDH")) {
            console.log("✔ [DOCX Test] Cột tùy chỉnh và dữ liệu xuất hiện thành công trong CTLOP DOCX");
        } else {
            console.error("❌ [DOCX Test] Thiếu dữ liệu cột tùy chỉnh trong CTLOP DOCX");
            errors++;
        }

        if (docXml.includes('w:firstLine="10"') && docXml.includes('w:left w:w="20"')) {
            console.log("✔ [DOCX Test] Định dạng w:left=20 dxa (1pt) và w:firstLine=10 dxa (0.5pt) xuất hiện trong bảng");
        } else {
            console.error("❌ [DOCX Test] Thiếu định dạng lề ô hoặc thụt đầu dòng trong bảng DOCX");
            errors++;
        }

        if (docXml.includes("Nội dung tích hợp / Điều chỉnh")) {
            console.log("✔ [DOCX Test] Cột tích hợp xuất hiện đầy đủ trong CTLOP");
        } else {
            console.error("❌ [DOCX Test] Thiếu cột tích hợp");
            errors++;
        }

        if (docXml.includes("Kí tên") && docXml.includes("Ghi chú")) {
            console.log("✔ [DOCX Test] Cột Kí tên và Ghi chú xuất hiện trong CTLOP");
        } else {
            console.error("❌ [DOCX Test] Thiếu cột Kí tên hoặc Ghi chú trong CTLOP");
            errors++;
        }

    } catch (e) {
        console.error("❌ [DOCX Test Exception]", e);
        errors++;
    }
}

// 5. Test XLSX Generator
async function testXlsxGeneration() {
    try {
        const xlsxContent = fs.readFileSync(path.join(__dirname, 'js/xlsx_generator.js'), 'utf8');
        if (xlsxContent.includes('getXlsxLbgOrderedColumns')) {
            console.log("✔ [XLSX] xlsx_generator.js chứa hàm getXlsxLbgOrderedColumns hỗ trợ CTLOP đồng bộ");
        } else {
            console.error("❌ [XLSX] Thiếu getXlsxLbgOrderedColumns trong xlsx_generator.js");
            errors++;
        }
    } catch (e) {
        console.error("❌ [XLSX Test Exception]", e);
        errors++;
    }
}

// 6. Kiểm tra JS app.js syntax
const appContent = fs.readFileSync(path.join(__dirname, 'js/app.js'), 'utf8');
try {
    // Kiểm tra parse syntax app.js
    new Function(appContent);
    console.log("✔ [JS App] app.js cú pháp hoàn toàn hợp lệ (Syntax OK)");
} catch (e) {
    console.error("❌ [JS App Syntax Error]", e);
    errors++;
}

testDocxGeneration().then(() => {
    testXlsxGeneration().then(() => {
        console.log("\n==========================================");
        if (errors === 0) {
            console.log("🎉 TẤT CẢ KIỂM TRA ĐÃ THÀNH CÔNG VỚI 0 LỖI!");
        } else {
            console.log(`⚠️ CÓ ${errors} LỖI CẦN XỬ LÝ!`);
        }
        console.log("==========================================\n");
    });
});
