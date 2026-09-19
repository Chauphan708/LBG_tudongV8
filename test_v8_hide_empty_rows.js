// Automated test for V8 hideEmptyRows option and 2pt indent
const fs = require('fs');
const path = require('path');

const ROOT_DIR = __dirname;
const JSZip = require(path.join(ROOT_DIR, 'lib/jszip.min.js'));

async function runTests() {
  console.log('=== RUNNING TEST V8: HIDE EMPTY ROWS & 2PT INDENT ===\n');
  let failures = 0;

  function assert(condition, message) {
    if (!condition) {
      console.error(`❌ FAIL: ${message}`);
      failures++;
    } else {
      console.log(`✅ PASS: ${message}`);
    }
  }

  // 1. Check CSS text-indent: 2pt
  const cssContent = fs.readFileSync(path.join(ROOT_DIR, 'css', 'styles.css'), 'utf8');
  assert(cssContent.includes('.table-lbg .col-lesson') && cssContent.includes('text-indent: 2pt;'), 'styles.css has text-indent: 2pt for .table-lbg .col-lesson');
  assert(cssContent.includes('.paper-table td.cell-lesson') && cssContent.includes('text-indent: 2pt !important;'), 'styles.css has text-indent: 2pt !important for .paper-table td.cell-lesson');

  // 2. Check index.html checkboxes
  const htmlContent = fs.readFileSync(path.join(ROOT_DIR, 'index.html'), 'utf8');
  assert(htmlContent.includes('id="lbg-opt-hide-empty-rows"'), 'index.html contains lbg-opt-hide-empty-rows');
  assert(htmlContent.includes('id="ctlop-opt-hide-empty-rows"'), 'index.html contains ctlop-opt-hide-empty-rows');
  assert(htmlContent.includes('id="modal-opt-hide-empty-rows"'), 'index.html contains modal-opt-hide-empty-rows');
  assert(htmlContent.includes('id="batch-opt-hide-empty-rows"'), 'index.html contains batch-opt-hide-empty-rows');

  // 3. Check docx_generator.js firstLine="40" (2pt)
  const docxGenContent = fs.readFileSync(path.join(ROOT_DIR, 'js', 'docx_generator.js'), 'utf8');
  assert(docxGenContent.includes('w:firstLine="40"'), 'docx_generator.js uses w:firstLine="40" for 2pt first-line indent');
  assert(docxGenContent.includes('hideEmptyRows'), 'docx_generator.js supports hideEmptyRows option');

  // Load docx_generator into Node environment
  global.JSZip = JSZip;
  global.window = { JSZip: JSZip };
  require(path.join(ROOT_DIR, 'js/docx_generator.js'));

  const weekInfo = {
    startDateVN: "08/09/2026",
    endDateVN: "12/09/2026"
  };

  const schedule = [
    { day: "Thứ 2", session: "Sáng", period: 1, subject: "Toán", ppct: "1", lessonName: "Ôn tập phân số", isOff: false },
    { day: "Thứ 2", session: "Sáng", period: 2, subject: "Tiếng Việt", ppct: "1", lessonName: "Tập đọc: Thư gửi các học sinh", isOff: false },
    { day: "Thứ 2", session: "Chiều", period: 1, subject: "Khoa học", ppct: "1", lessonName: "Sự sinh sản", isOff: false },
    { day: "Thứ 6", session: "Sáng", period: 1, subject: "Toán", ppct: "5", lessonName: "Luyện tập chung", isOff: false },
    { day: "Thứ 6", session: "Chiều", period: 1, subject: "-- Nghỉ / Để trống --", ppct: "", lessonName: "", isOff: true },
    { day: "Thứ 6", session: "Chiều", period: 2, subject: "-- Nghỉ / Để trống --", ppct: "", lessonName: "", isOff: true }
  ];

  const settings = {
    governingBody: "PHÒNG GD&ĐT HUYỆN TEST",
    schoolName: "TRƯỜNG TIỂU HỌC TEST",
    className: "5A1",
    grade: "Khối 5",
    homeroomTeacher: "Nguyễn Văn A",
    headOfGrade: "Trần Thị Mai",
    vicePrincipal: "Phan Văn Minh"
  };

  async function getZipXml(blob) {
    let inputData = blob;
    if (blob && typeof blob.arrayBuffer === 'function') {
      inputData = Buffer.from(await blob.arrayBuffer());
    }
    const zip = await JSZip.loadAsync(inputData);
    return await zip.file('word/document.xml').async('string');
  }

  // Test DOCX with hideEmptyRows = false
  const blobDefault = await window.DocxGenerator.generateLbgDocx(
    false, 1, weekInfo, schedule, settings, {}, "portrait", { hideEmptyRows: false }
  );
  const xmlDefault = await getZipXml(blobDefault);
  assert(xmlDefault.includes('Nghỉ / Để trống'), 'DOCX with hideEmptyRows=false keeps empty rows');
  assert(xmlDefault.includes('w:firstLine="40"'), 'DOCX includes 2pt indent (w:firstLine="40")');

  // Test DOCX with hideEmptyRows = true
  const blobFiltered = await window.DocxGenerator.generateLbgDocx(
    false, 1, weekInfo, schedule, settings, {}, "portrait", { hideEmptyRows: true }
  );
  const xmlFiltered = await getZipXml(blobFiltered);
  assert(!xmlFiltered.includes('Nghỉ / Để trống'), 'DOCX with hideEmptyRows=true removes empty rows');
  assert(xmlFiltered.includes('Ôn tập phân số'), 'DOCX with hideEmptyRows=true keeps regular morning slots');
  assert(xmlFiltered.includes('Sự sinh sản'), 'DOCX with hideEmptyRows=true keeps regular afternoon slots of Thứ 2');

  // Test multi-week DOCX with hideEmptyRows = true
  const calculateWeekScheduleFn = (w) => ({ weekInfo, schedule, stats: {} });
  const blobMulti = await window.DocxGenerator.generateMultiWeekLbgDocx(
    false, 1, 1, calculateWeekScheduleFn, settings, "portrait", { hideEmptyRows: true }
  );
  const xmlMulti = await getZipXml(blobMulti);
  assert(!xmlMulti.includes('Nghỉ / Để trống'), 'Multi-week DOCX with hideEmptyRows=true removes empty rows');

  // 4. Check xlsx_generator.js
  const xlsxGenContent = fs.readFileSync(path.join(ROOT_DIR, 'js', 'xlsx_generator.js'), 'utf8');
  assert(xlsxGenContent.includes('hideEmptyRows'), 'xlsx_generator.js supports hideEmptyRows option');

  // 5. Check app.js integration
  const appJsContent = fs.readFileSync(path.join(ROOT_DIR, 'js', 'app.js'), 'utf8');
  assert(appJsContent.includes('state.lbgHideEmptyRows'), 'app.js manages state.lbgHideEmptyRows');
  assert(appJsContent.includes('lbg-opt-hide-empty-rows'), 'app.js binds lbg-opt-hide-empty-rows');
  assert(appJsContent.includes('ctlop-opt-hide-empty-rows'), 'app.js binds ctlop-opt-hide-empty-rows');
  assert(appJsContent.includes('modal-opt-hide-empty-rows'), 'app.js binds modal-opt-hide-empty-rows');
  assert(appJsContent.includes('batch-opt-hide-empty-rows'), 'app.js binds batch-opt-hide-empty-rows');
  assert(appJsContent.includes('hideEmptyRows: !!state.lbgHideEmptyRows'), 'app.js passes hideEmptyRows to export functions');

  if (failures === 0) {
    console.log('\n🎉 ALL TESTS PASSED SUCCESSFULLY! (100% OK)');
  } else {
    console.error(`\n⚠️ ENCOUNTERED ${failures} FAILURES!`);
    process.exit(1);
  }
}

runTests().catch(err => {
  console.error('Test execution error:', err);
  process.exit(1);
});
