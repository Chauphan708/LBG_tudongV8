/**
 * DOCX Generator for KHDH (Kế hoạch dạy học) and LBG (Lịch báo giảng)
 * Conforming strictly to Decree 30/2020/ND-CP
 * Uses JSZip and FileSaver for 100% offline generation in browser
 */

window.DocxGenerator = (function() {
    
    function escapeXml(unsafe) {
        if (!unsafe) return "";
        return unsafe.toString()
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    function createContentTypes() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;
    }

    function createRels() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
    }

    function createWordRels() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
    }

    function createStyles() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:docDefaults>
        <w:rPrDefault>
            <w:rPr>
                <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
                <w:sz w:val="26"/>
                <w:szCs w:val="26"/>
                <w:lang w:val="vi-VN"/>
            </w:rPr>
        </w:rPrDefault>
        <w:pPrDefault>
            <w:pPr>
                <w:spacing w:line="276" w:lineRule="auto" w:after="100"/>
            </w:pPr>
        </w:pPrDefault>
    </w:docDefaults>
</w:styles>`;
    }

    function formatIntegrationXml(rawInteg) {
        const text = (rawInteg || '').replace(/System\.Xml\.XmlElement/g, '').trim();
        const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
        if (lines.length === 0) {
            return `<w:p><w:pPr><w:spacing w:after="0" w:line="240"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t></w:t></w:r></w:p>`;
        }
        return lines.map(line => {
            return `<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:after="30" w:line="240"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(line)}</w:t></w:r></w:p>`;
        }).join("");
    }

    function formatContentXml(rawContent) {
        const text = (rawContent || '').trim();
        const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
        if (lines.length === 0) {
            return `<w:p><w:pPr><w:spacing w:after="0" w:line="240"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t></w:t></w:r></w:p>`;
        }
        return lines.map(line => {
            return `<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:after="20" w:line="240"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(line)}</w:t></w:r></w:p>`;
        }).join("");
    }

    function getBaseLesson(lessonStr) {
        if (!lessonStr) return '';
        return lessonStr.replace(/\s*[\(\,\–\-]\s*(?:tiết|Tiết)\s*\d+[\)\.]?/gi, '').trim();
    }

    /**
     * Compute vertical row spans for KHDH table cells (vMerge)
     * Handles hierarchical week, topic, lesson group (continuation rows where lesson is empty) and duration
     */
    function computeKhdhSpans(rows, cat) {
        if (!rows || !rows.length) return { week: [], topic: [], lesson: [], duration: [] };
        const n = rows.length;
        const spans = {
            week: new Array(n).fill(0),
            topic: new Array(n).fill(0),
            lesson: new Array(n).fill(0),
            duration: new Array(n).fill(0)
        };

        // 1. Week span
        let i = 0;
        while (i < n) {
            let j = i + 1;
            while (j < n && rows[j].week === rows[i].week) {
                j++;
            }
            spans.week[i] = j - i;
            for (let k = i + 1; k < j; k++) spans.week[k] = -1;
            i = j;
        }

        // 2. Topic span (within same week)
        i = 0;
        while (i < n) {
            let j = i + 1;
            const topic_i = (rows[i].topic || rows[i].theme || '').trim();
            while (j < n && rows[j].week === rows[i].week && (rows[j].topic || rows[j].theme || '').trim() === topic_i) {
                j++;
            }
            spans.topic[i] = j - i;
            for (let k = i + 1; k < j; k++) spans.topic[k] = -1;
            i = j;
        }

        // 3. Lesson and Duration span (within same week)
        i = 0;
        while (i < n) {
            let j = i + 1;
            if (cat === 'tieng_viet' || cat === 'hdtn') {
                while (j < n && rows[j].week === rows[i].week && !(rows[j].lesson && rows[j].lesson.trim())) {
                    j++;
                }
            } else {
                const durMatch = (rows[i].duration || '').match(/(\d+)\s*tiết/i);
                const expectedCount = durMatch ? parseInt(durMatch[1], 10) : 1;
                const base_i = getBaseLesson(rows[i].lesson);

                while (j < n && rows[j].week === rows[i].week && (j - i < expectedCount)) {
                    const base_j = getBaseLesson(rows[j].lesson);
                    if (base_j === base_i || !(rows[j].duration && rows[j].duration.trim())) {
                        j++;
                    } else {
                        break;
                    }
                }
            }

            const count = j - i;
            spans.lesson[i] = count;
            spans.duration[i] = count;
            for (let k = i + 1; k < j; k++) {
                spans.lesson[k] = -1;
                spans.duration[k] = -1;
            }
            i = j;
        }

        return spans;
    }

    /**
     * Generate KHDH DOCX Document
     * @param {Object} options { settings, subjectsData: [{ subjectName, rows: [...] }], isMultiSubject: boolean }
     */
    function generateKhdhDocx(options) {
        const { settings, subjectsData, isMultiSubject } = options;
        const zip = new JSZip();

        zip.file("[Content_Types].xml", createContentTypes());
        zip.file("_rels/.rels", createRels());
        zip.file("word/_rels/document.xml.rels", createWordRels());
        zip.file("word/styles.xml", createStyles());

        let docBody = "";

        // 1. Header (UBND / Trường - Quốc hiệu / Tiêu ngữ) theo Nghị định 30/2020/NĐ-CP
        docBody += `
        <w:tbl>
            <w:tblPr>
                <w:tblW w:w="9600" w:type="dxa"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
                    <w:insideH w:val="none"/><w:insideV w:val="none"/>
                </w:tblBorders>
                <w:tblCellMar>
                    <w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>
                    <w:left w:w="40" w:type="dxa"/><w:right w:w="40" w:type="dxa"/>
                </w:tblCellMar>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="4200"/>
                <w:gridCol w:w="5400"/>
            </w:tblGrid>
            <w:tr>
                <w:tc>
                    <w:tcPr><w:tcW w:w="4200" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:after="40"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="23"/></w:rPr><w:t>${escapeXml(settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT')}</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:after="60"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="23"/></w:rPr><w:t>${escapeXml(settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT')}</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:after="100"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="23"/></w:rPr><w:t>Số:       /KH-TH-K5</w:t></w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="5400" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:after="40"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="23"/></w:rPr><w:t>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:after="60"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="24"/><w:u w:val="single"/></w:rPr><w:t>Độc lập - Tự do - Hạnh phúc</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="right"/><w:spacing w:line="240" w:after="100"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="23"/></w:rPr><w:t>${escapeXml(settings.location || 'Trung Nhứt')}, ${escapeXml(settings.dateString || 'ngày     tháng 8 năm 2026')}</w:t></w:r>
                    </w:p>
                </w:tc>
            </w:tr>
        </w:tbl>`;

        // 2. Tiêu đề chung & Căn cứ pháp lý
        const mainTitle = isMultiSubject
            ? `KẾ HOẠCH DẠY HỌC CÁC MÔN HỌC ${escapeXml((settings.grade || 'KHỐI 5').toUpperCase())}`
            : `KẾ HOẠCH DẠY HỌC MÔN ${escapeXml(subjectsData[0].subjectName.toUpperCase())} ${escapeXml((settings.grade || 'KHỐI 5').toUpperCase())}`;

        docBody += `
        <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="60" w:line="280"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="30"/><w:color w:val="002060"/></w:rPr><w:t>${mainTitle}</w:t></w:r>
        </w:p>
        <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="180" w:line="280"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/></w:rPr><w:t>NĂM HỌC ${escapeXml(settings.academicYear || '2026 - 2027')}</w:t></w:r>
        </w:p>`;

        const defaultBases = [
            "Căn cứ Thông tư số 32/2018/TT-BGDĐT ngày 26/12/2018 Ban hành Chương trình Giáo dục phổ thông tổng thể.",
            "Căn cứ Thông tư 27/2020/TT-BGDĐT ngày 04/9/2020 về việc ban hành Quy định đánh giá học sinh tiểu học;",
            "Căn cứ Công văn số 2345/BGDĐT-GDTH ngày 07/6/2021 về việc hướng dẫn xây dựng kế hoạch giáo dục nhà trường."
        ];

        const basesList = (Array.isArray(settings.bases) && settings.bases.length > 0 && !settings.bases.some(b => b.includes("28/2020")))
            ? settings.bases
            : defaultBases;

        basesList.forEach(b => {
            docBody += `
        <w:p>
            <w:pPr><w:jc w:val="both"/><w:spacing w:after="60" w:line="260"/><w:ind w:firstLine="360"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/></w:rPr><w:t>${escapeXml(b)}</w:t></w:r>
        </w:p>`;
        });

        docBody += `
        <w:p>
            <w:pPr><w:jc w:val="both"/><w:spacing w:after="200" w:line="260"/><w:ind w:firstLine="360"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:t>Tổ chuyên môn ${escapeXml(settings.grade || 'Khối 5')} xây dựng Kế hoạch dạy học chi tiết ${isMultiSubject ? 'các môn học' : 'môn ' + escapeXml(subjectsData[0].subjectName)} như sau:</w:t></w:r>
        </w:p>`;

        // Helper to determine subject category
        function getSubjectCategory(name) {
            const s = (name || '').toUpperCase().trim();
            if (s.includes('TIẾNG VIỆT') || s.includes('TIENG VIET')) return 'tieng_viet';
            if (s.includes('TRẢI NGHIỆM') || s.includes('TRAI NGHIEM') || s.includes('HĐTN') || s.includes('HDTN')) return 'hdtn';
            return 'khmh'; // TOÁN, KHOA HỌC, LS&ĐL, CÔNG NGHỆ, ĐẠO ĐỨC
        }

        // 3. Từng môn học (Lặp qua subjectsData)
        subjectsData.forEach((sub, subIdx) => {
            const partHeader = isMultiSubject
                ? `PHẦN ${subIdx + 1}: KẾ HOẠCH DẠY HỌC MÔN ${escapeXml(sub.subjectName.toUpperCase())}`
                : `BẢNG KẾ HOẠCH DẠY HỌC CHI TIẾT (35 TUẦN)`;

            docBody += `
            <w:p>
                <w:pPr><w:jc w:val="left"/><w:spacing w:before="240" w:after="120" w:line="260"/></w:pPr>
                <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:color w:val="004080"/></w:rPr><w:t>${partHeader}</w:t></w:r>
            </w:p>`;

            const cat = getSubjectCategory(sub.subjectName);

            // Bảng kế hoạch dạy học tùy biến theo môn học (Khớp chuẩn 9600 dxa theo lề A4)
            if (cat === 'tieng_viet') {
                // 1. TIẾNG VIỆT (7 Cột: Tuần, Chủ đề/mạch nội dung, Bài học, Tên hoạt động/Nội dung, Tiết/Thời lượng, Điều chỉnh bổ sung, Ghi chú)
                const spans = computeKhdhSpans(sub.rows, 'tieng_viet');

                docBody += `
                <w:tbl>
                    <w:tblPr>
                        <w:tblW w:w="9600" w:type="dxa"/>
                        <w:jc w:val="center"/>
                        <w:tblLayout w:type="fixed"/>
                        <w:tblBorders>
                            <w:top w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>
                            <w:left w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>
                            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>
                            <w:right w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>
                            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
                            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
                        </w:tblBorders>
                        <w:tblCellMar>
                            <w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/>
                            <w:left w:w="90" w:type="dxa"/><w:right w:w="90" w:type="dxa"/>
                        </w:tblCellMar>
                    </w:tblPr>
                    <w:tblGrid>
                        <w:gridCol w:w="550"/>
                        <w:gridCol w:w="1350"/>
                        <w:gridCol w:w="1450"/>
                        <w:gridCol w:w="2250"/>
                        <w:gridCol w:w="650"/>
                        <w:gridCol w:w="2650"/>
                        <w:gridCol w:w="700"/>
                    </w:tblGrid>
                    <w:tr>
                        <w:trPr><w:tblHeader/></w:trPr>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="550" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Tuần</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="1350" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Chủ đề / Mạch nội dung</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="1450" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Bài học</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="2250" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Tên hoạt động / Nội dung dạy học</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="650" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Tiết / Thời lượng</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="2650" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Nội dung điều chỉnh, bổ sung (nếu có)</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="700" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Ghi chú</w:t></w:r></w:p>
                        </w:tc>
                    </w:tr>`;

                sub.rows.forEach((r, rIdx) => {
                    const integXml = formatIntegrationXml(r.integration);
                    const contentXml = formatContentXml(r.content);
                    const lessonText = r.lesson || '';
                    const topicText = r.topic || r.theme || '';

                    // Cột 1: Tuần (Gộp ô vMerge)
                    let weekCellXml = (spans.week[rIdx] > 0)
                        ? `<w:tc><w:tcPr><w:tcW w:w="550" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(r.week || '')}</w:t></w:r></w:p></w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="550" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr></w:p></w:tc>`;

                    // Cột 2: Chủ đề / Mạch nội dung (Gộp ô vMerge)
                    let topicCellXml = (spans.topic[rIdx] > 0)
                        ? `<w:tc><w:tcPr><w:tcW w:w="1350" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(topicText)}</w:t></w:r></w:p></w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="1350" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="220"/></w:pPr></w:p></w:tc>`;

                    // Cột 3: Bài học (Gộp ô vMerge)
                    let lessonCellXml = (spans.lesson[rIdx] > 0)
                        ? `<w:tc><w:tcPr><w:tcW w:w="1450" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(lessonText)}</w:t></w:r></w:p></w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="1450" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="220"/></w:pPr></w:p></w:tc>`;

                    // Cột 5: Thời lượng (Gộp ô vMerge)
                    let durCellXml = (spans.duration[rIdx] > 0)
                        ? `<w:tc><w:tcPr><w:tcW w:w="650" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(r.duration || (spans.duration[rIdx] + ' tiết'))}</w:t></w:r></w:p></w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="650" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr></w:p></w:tc>`;

                    docBody += `
                    <w:tr>
                        <w:trPr><w:cantSplit/></w:trPr>
                        ${weekCellXml}
                        ${topicCellXml}
                        ${lessonCellXml}
                        <w:tc>
                            <w:tcPr><w:tcW w:w="2250" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            ${contentXml}
                        </w:tc>
                        ${durCellXml}
                        <w:tc>
                            <w:tcPr><w:tcW w:w="2650" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                            ${integXml}
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="700" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(r.notes || '')}</w:t></w:r></w:p>
                        </w:tc>
                    </w:tr>`;
                });
                docBody += `</w:tbl>`;

            } else if (cat === 'hdtn') {
                // 2. HOẠT ĐỘNG TRẢI NGHIỆM (6 Cột: Tuần, Chủ đề/mạch nội dung, tên bài học, Tên hoạt động, Thời lượng, Nội dung điều chỉnh, bổ sung)
                const spans = computeKhdhSpans(sub.rows, 'hdtn');

                docBody += `
                <w:tbl>
                    <w:tblPr>
                        <w:tblW w:w="9600" w:type="dxa"/>
                        <w:jc w:val="center"/>
                        <w:tblLayout w:type="fixed"/>
                        <w:tblBorders>
                            <w:top w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>
                            <w:left w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>
                            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>
                            <w:right w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>
                            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
                            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
                        </w:tblBorders>
                        <w:tblCellMar>
                            <w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/>
                            <w:left w:w="90" w:type="dxa"/><w:right w:w="90" w:type="dxa"/>
                        </w:tblCellMar>
                    </w:tblPr>
                    <w:tblGrid>
                        <w:gridCol w:w="550"/>
                        <w:gridCol w:w="1550"/>
                        <w:gridCol w:w="1700"/>
                        <w:gridCol w:w="1900"/>
                        <w:gridCol w:w="700"/>
                        <w:gridCol w:w="3200"/>
                    </w:tblGrid>
                    <w:tr>
                        <w:trPr><w:tblHeader/></w:trPr>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="550" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Tuần</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="1550" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Chủ đề / Mạch nội dung</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="1700" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Tên bài học</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="1900" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Tên hoạt động</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="700" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Thời lượng</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="3200" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Nội dung điều chỉnh, bổ sung (nếu có)</w:t></w:r></w:p>
                        </w:tc>
                    </w:tr>`;

                sub.rows.forEach((r, rIdx) => {
                    const integXml = formatIntegrationXml(r.integration);
                    const contentXml = formatContentXml(r.content);
                    const lessonText = r.lesson || '';
                    const topicText = r.topic || r.theme || '';

                    // Cột 1: Tuần (Gộp ô vMerge)
                    let weekCellXml = (spans.week[rIdx] > 0)
                        ? `<w:tc><w:tcPr><w:tcW w:w="550" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(r.week || '')}</w:t></w:r></w:p></w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="550" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr></w:p></w:tc>`;

                    // Cột 2: Chủ đề / Mạch nội dung (Gộp ô vMerge)
                    let topicCellXml = (spans.topic[rIdx] > 0)
                        ? `<w:tc><w:tcPr><w:tcW w:w="1550" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(topicText)}</w:t></w:r></w:p></w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="1550" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="220"/></w:pPr></w:p></w:tc>`;

                    // Cột 3: Tên bài học (Gộp ô vMerge)
                    let lessonCellXml = (spans.lesson[rIdx] > 0)
                        ? `<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(lessonText)}</w:t></w:r></w:p></w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="220"/></w:pPr></w:p></w:tc>`;

                    // Cột 5: Thời lượng (Gộp ô vMerge)
                    let durCellXml = (spans.duration[rIdx] > 0)
                        ? `<w:tc><w:tcPr><w:tcW w:w="700" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(r.duration || (spans.duration[rIdx] + ' tiết'))}</w:t></w:r></w:p></w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="700" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr></w:p></w:tc>`;

                    docBody += `
                    <w:tr>
                        <w:trPr><w:cantSplit/></w:trPr>
                        ${weekCellXml}
                        ${topicCellXml}
                        ${lessonCellXml}
                        <w:tc>
                            <w:tcPr><w:tcW w:w="1900" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            ${contentXml}
                        </w:tc>
                        ${durCellXml}
                        <w:tc>
                            <w:tcPr><w:tcW w:w="3200" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                            ${integXml}
                        </w:tc>
                    </w:tr>`;
                });
                docBody += `</w:tbl>`;

            } else {
                // 3. TOÁN, KHOA HỌC, LỊCH SỬ VÀ ĐỊA LÍ, CÔNG NGHỆ, ĐẠO ĐỨC (6 Cột: Tuần, Chủ đề/mạch nội dung, tên bài học, Thời lượng, tiết theo KHMH, Nội dung điều chỉnh, bổ sung)
                const spans = computeKhdhSpans(sub.rows, 'khmh');

                docBody += `
                <w:tbl>
                    <w:tblPr>
                        <w:tblW w:w="9600" w:type="dxa"/>
                        <w:jc w:val="center"/>
                        <w:tblLayout w:type="fixed"/>
                        <w:tblBorders>
                            <w:top w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>
                            <w:left w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>
                            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>
                            <w:right w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/>
                            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
                            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>
                        </w:tblBorders>
                        <w:tblCellMar>
                            <w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/>
                            <w:left w:w="90" w:type="dxa"/><w:right w:w="90" w:type="dxa"/>
                        </w:tblCellMar>
                    </w:tblPr>
                    <w:tblGrid>
                        <w:gridCol w:w="550"/>
                        <w:gridCol w:w="1700"/>
                        <w:gridCol w:w="2150"/>
                        <w:gridCol w:w="800"/>
                        <w:gridCol w:w="900"/>
                        <w:gridCol w:w="3500"/>
                    </w:tblGrid>
                    <w:tr>
                        <w:trPr><w:tblHeader/></w:trPr>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="550" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Tuần</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="1700" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Chủ đề / Mạch nội dung</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="2150" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Tên bài học</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="800" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Thời lượng</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="900" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Tiết theo KHMH</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="3500" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EBF1F5"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Nội dung điều chỉnh, bổ sung (nếu có)</w:t></w:r></w:p>
                        </w:tc>
                    </w:tr>`;

                sub.rows.forEach((r, rIdx) => {
                    const integXml = formatIntegrationXml(r.integration);
                    const lessonText = (spans.lesson[rIdx] > 1) ? (getBaseLesson(r.lesson) || r.lesson) : (r.lesson || '');
                    const topicText = r.topic || r.theme || '';
                    const durText = r.duration || (spans.duration[rIdx] > 0 ? (spans.duration[rIdx] + ' tiết') : '');
                    const khmhPeriodText = r.khmhPeriod || r.content || (rIdx + 1).toString();

                    // Cột 1: Tuần (Gộp ô vMerge)
                    let weekCellXml = (spans.week[rIdx] > 0)
                        ? `<w:tc><w:tcPr><w:tcW w:w="550" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(r.week || '')}</w:t></w:r></w:p></w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="550" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr></w:p></w:tc>`;

                    // Cột 2: Chủ đề / Mạch nội dung (Gộp ô vMerge)
                    let topicCellXml = (spans.topic[rIdx] > 0)
                        ? `<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(topicText)}</w:t></w:r></w:p></w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="220"/></w:pPr></w:p></w:tc>`;

                    // Cột 3: Tên bài học (Gộp ô vMerge)
                    let lessonCellXml = (spans.lesson[rIdx] > 0)
                        ? `<w:tc><w:tcPr><w:tcW w:w="2150" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(lessonText)}</w:t></w:r></w:p></w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="2150" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="0" w:line="220"/></w:pPr></w:p></w:tc>`;

                    // Cột 4: Thời lượng (Gộp ô vMerge)
                    let durCellXml = (spans.duration[rIdx] > 0)
                        ? `<w:tc><w:tcPr><w:tcW w:w="800" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(durText)}</w:t></w:r></w:p></w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="800" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr></w:p></w:tc>`;

                    docBody += `
                    <w:tr>
                        <w:trPr><w:cantSplit/></w:trPr>
                        ${weekCellXml}
                        ${topicCellXml}
                        ${lessonCellXml}
                        ${durCellXml}
                        <w:tc>
                            <w:tcPr><w:tcW w:w="900" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(khmhPeriodText)}</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="3500" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                            ${integXml}
                        </w:tc>
                    </w:tr>`;
                });
                docBody += `</w:tbl>`;
            }
        });

        // 4. Phần kết thúc chung & Ký duyệt chuẩn Nghị định 30/2020/NĐ-CP
        docBody += `
        <w:p><w:pPr><w:spacing w:before="240" w:after="120" w:line="240"/></w:pPr></w:p>
        <w:tbl>
            <w:tblPr>
                <w:tblW w:w="9600" w:type="dxa"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
                    <w:insideH w:val="none"/><w:insideV w:val="none"/>
                </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="3800"/>
                <w:gridCol w:w="5800"/>
            </w:tblGrid>
            <w:tr>
                <w:tc>
                    <w:tcPr><w:tcW w:w="3800" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p>
                        <w:pPr><w:jc w:val="left"/><w:spacing w:line="240" w:after="40"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>Nơi nhận:</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="left"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>- Ban Giám hiệu (để b/c);</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="left"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>- Tổ chuyên môn ${escapeXml(settings.grade || 'Khối 5')} (để t/h);</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="left"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>- Giáo viên giảng dạy (để t/h);</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="left"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>- Lưu: VT, Hồ sơ Tổ CM.</w:t></w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="5800" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:after="40"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>TỔ TRƯỞNG CHUYÊN MÔN</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:after="800"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>(Ký và ghi rõ họ tên)</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:after="200"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(settings.headOfGrade || 'Trần Thị Mai')}</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:before="200" w:after="40"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>PHÊ DUYỆT CỦA ${escapeXml((settings.bghSignerKhdhType || settings.bghSignerType || 'HT') === 'HT' ? 'HIỆU TRƯỞNG' : 'PHÓ HIỆU TRƯỞNG')}</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:after="900"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>(Ký, ghi rõ họ tên và đóng dấu)</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:after="100"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(((settings.bghSignerKhdhType || settings.bghSignerType || 'HT') === 'HT') ? (settings.principal || 'Phạm Quốc Hùng') : (Array.isArray(settings.vicePrincipals) && settings.vicePrincipals[(settings.bghSignerKhdhIndex !== undefined ? settings.bghSignerKhdhIndex : (settings.bghSignerIndex || 0))] ? settings.vicePrincipals[(settings.bghSignerKhdhIndex !== undefined ? settings.bghSignerKhdhIndex : (settings.bghSignerIndex || 0))] : (settings.vicePrincipal || 'Lê Văn Tám')))}</w:t></w:r>
                    </w:p>
                </w:tc>
            </w:tr>
        </w:tbl>`;

        // Page setup: A4 Portrait, Margins: Top 20mm (1134 dxa), Bottom 20mm (1134 dxa), Left 25mm (1417 dxa), Right 15mm (850 dxa)
        const fullDoc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        ${docBody}
        <w:sectPr>
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1417" w:header="720" w:footer="720" w:gutter="0"/>
        </w:sectPr>
    </w:body>
</w:document>`;

        zip.file("word/document.xml", fullDoc);

        return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    }

    function getDayFullDate(startDateVN, dayStr) {
        if (!startDateVN) return "";
        const parts = startDateVN.split('/');
        if (parts.length < 3) return "";
        const startD = parseInt(parts[0], 10);
        const startM = parseInt(parts[1], 10) - 1;
        const startY = parseInt(parts[2], 10);
        const dayOffsets = { "Thứ 2": 0, "Thứ 3": 1, "Thứ 4": 2, "Thứ 5": 3, "Thứ 6": 4, "Thứ 7": 5, "Chủ nhật": 6 };
        const offset = dayOffsets[dayStr] !== undefined ? dayOffsets[dayStr] : 0;
        const targetDate = new Date(startY, startM, startD + offset);
        const dStr = String(targetDate.getDate()).padStart(2, '0');
        const mStr = String(targetDate.getMonth() + 1).padStart(2, '0');
        const yStr = targetDate.getFullYear();
        return `${dStr}/${mStr}/${yStr}`;
    }

    function generateLbgDocx(isCtlop, weekNum, weekInfo, schedule, settings, stats, orientation = "portrait", options = {}) {
        const zip = new JSZip();

        zip.file("[Content_Types].xml", createContentTypes());
        zip.file("_rels/.rels", createRels());
        zip.file("word/_rels/document.xml.rels", createWordRels());
        zip.file("word/styles.xml", createStyles());

        const isLandscape = (orientation === "landscape");
        const headerColWidths = isLandscape ? [7500, 7500] : [4800, 4800];
        const sigColWidth = isLandscape ? "5000" : "3200";

        const bghSignerName = ((settings.bghSignerLbgType || settings.bghSignerType || 'PHT') === 'HT')
            ? (settings.principal || 'Phạm Quốc Hùng')
            : (Array.isArray(settings.vicePrincipals) && settings.vicePrincipals[(settings.bghSignerLbgIndex !== undefined ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0))]
                ? settings.vicePrincipals[(settings.bghSignerLbgIndex !== undefined ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0))]
                : (settings.vicePrincipal || 'Lê Văn Tám'));
        const bghSignerRole = ((settings.bghSignerLbgType || settings.bghSignerType || 'PHT') === 'HT') ? 'HIỆU TRƯỞNG' : 'BAN GIÁM HIỆU';

        let docBody = `
        <w:tbl>
            <w:tblPr>
                <w:tblW w:w="0" w:type="auto"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
                    <w:insideH w:val="none"/><w:insideV w:val="none"/>
                </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="${headerColWidths[0]}"/>
                <w:gridCol w:w="${headerColWidths[1]}"/>
            </w:tblGrid>
            <w:tr>
                <w:tc>
                    <w:tcPr><w:tcW w:w="${headerColWidths[0]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr><w:t>${escapeXml((settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT').toUpperCase())}</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>${escapeXml((settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT').toUpperCase())}</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="40"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>${escapeXml((settings.grade || 'KHỐI 5').toUpperCase())} - ${escapeXml((settings.className || 'LỚP 5A').toUpperCase())}</w:t></w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="${headerColWidths[1]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="40"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:u w:val="single"/><w:sz w:val="23"/></w:rPr><w:t>Độc lập - Tự do - Hạnh phúc</w:t></w:r>
                    </w:p>
                </w:tc>
            </w:tr>
        </w:tbl>

        <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="40"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="30"/></w:rPr><w:t>${isCtlop ? 'LỊCH BÁO GIẢNG TÍCH HỢP' : 'LỊCH BÁO GIẢNG'} TUẦN ${weekNum}</w:t></w:r>
        </w:p>
        <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:after="200"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="22"/></w:rPr><w:t>(Thời gian thực hiện: Từ ngày ${escapeXml(weekInfo.startDateVN || '')} đến ngày ${escapeXml(weekInfo.endDateVN || '')})</w:t></w:r>
        </w:p>
        `;

        const showColSign = !isCtlop && !!(options && options.showColSign);
        const showColNote = !isCtlop && !!(options && options.showColNote);
        const showColCustom = !isCtlop && !!(options && options.showColCustom);
        const colCustomName = (options && options.colCustomName) ? options.colCustomName : "Ghi chú";
        const showBghSign = (options && options.showBghSign !== undefined) ? options.showBghSign : true;
        const showHeadSign = (options && options.showHeadSign !== undefined) ? options.showHeadSign : true;
        const showGvcnSign = (options && options.showGvcnSign !== undefined) ? options.showGvcnSign : true;

        let colWidths = [];
        let headers = [];

        if (isCtlop) {
            colWidths = isLandscape
                ? [1300, 850, 600, 2400, 1000, 4350, 4500]
                : [1050, 700, 500, 1600, 750, 2550, 2450];
            headers = ["Thứ, ngày", "Buổi", "Tiết", "Môn học", "Tiết PPCT", "Tên bài dạy", "Nội dung tích hợp / Điều chỉnh"];
        } else {
            headers = ["Thứ, ngày", "Buổi", "Tiết", "Môn học", "Tiết PPCT", "Tên bài dạy"];
            if (isLandscape) {
                let wLesson = 7950;
                let wSign = 1300;
                let wNote = 2000;
                let wCustom = 2000;
                if (showColSign) wLesson -= wSign;
                if (showColNote) wLesson -= wNote;
                if (showColCustom) wLesson -= wCustom;
                colWidths = [1450, 950, 650, 2800, 1200, wLesson];
                if (showColSign) {
                    headers.push("Kí tên");
                    colWidths.push(wSign);
                }
                if (showColNote) {
                    headers.push("Ghi chú");
                    colWidths.push(wNote);
                }
                if (showColCustom) {
                    headers.push(colCustomName);
                    colWidths.push(wCustom);
                }
            } else {
                let wLesson = 4450;
                let wSign = 850;
                let wNote = 1400;
                let wCustom = 1400;
                if (showColSign) wLesson -= wSign;
                if (showColNote) wLesson -= wNote;
                if (showColCustom) wLesson -= wCustom;
                colWidths = [1150, 750, 550, 1900, 800, wLesson];
                if (showColSign) {
                    headers.push("Kí tên");
                    colWidths.push(wSign);
                }
                if (showColNote) {
                    headers.push("Ghi chú");
                    colWidths.push(wNote);
                }
                if (showColCustom) {
                    headers.push(colCustomName);
                    colWidths.push(wCustom);
                }
            }
        }

        docBody += `
        <w:tbl>
            <w:tblPr>
                <w:tblW w:w="0" w:type="auto"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>
                ${colWidths.map(w => `<w:gridCol w:w="${w}"/>`).join('')}
            </w:tblGrid>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                ${headers.map((h, i) => `
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${colWidths[i]}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:vAlign w:val="center"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(h)}</w:t></w:r></w:p>
                    </w:tc>
                `).join('')}
            </w:tr>
        `;

        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];
        days.forEach(day => {
            const daySlots = schedule.filter(s => s.day === day);
            if (daySlots.length === 0) return;

            const dayDate = getDayFullDate(weekInfo.startDateVN, day);
            const morningSlots = daySlots.filter(s => s.session === "Sáng");
            const afternoonSlots = daySlots.filter(s => s.session === "Chiều");

            daySlots.forEach((slot, idx) => {
                const isDayStart = (idx === 0);
                const isMornStart = (idx === 0 && morningSlots.length > 0);
                const isAftStart = (idx === morningSlots.length && afternoonSlots.length > 0);
                let dayCellXml = isDayStart
                    ? `<w:tc>
                        <w:tcPr><w:tcW w:w="${colWidths[0]}" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr>
                        <w:p>
                            <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="${dayDate ? '20' : '0'}"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(day)}</w:t></w:r>
                        </w:p>
                        ${dayDate ? `
                        <w:p>
                            <w:pPr><w:jc w:val="center"/><w:spacing w:line="200" w:after="0"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="333333"/></w:rPr><w:t>${escapeXml(dayDate)}</w:t></w:r>
                        </w:p>` : ''}
                       </w:tc>`
                    : `<w:tc><w:tcPr><w:tcW w:w="${colWidths[0]}" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr></w:p></w:tc>`;

                let sessCellXml = (isMornStart || isAftStart)
                    ? `<w:tc><w:tcPr><w:tcW w:w="${colWidths[1]}" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${isMornStart ? 'Sáng' : 'Chiều'}</w:t></w:r></w:p></w:tc>`
                    : `<w:tc><w:tcPr><w:tcW w:w="${colWidths[1]}" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr></w:p></w:tc>`;

                let signCellXml = showColSign ? `<w:tc><w:tcPr><w:tcW w:w="${isLandscape ? 1300 : 850}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr></w:p></w:tc>` : '';
                let noteCellXml = showColNote ? `<w:tc><w:tcPr><w:tcW w:w="${isLandscape ? 2000 : 1400}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:line="200" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.note || '')}</w:t></w:r></w:p></w:tc>` : '';
                let customCellXml = showColCustom ? `<w:tc><w:tcPr><w:tcW w:w="${isLandscape ? 2000 : 1400}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:line="200" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.customCol || slot.note || '')}</w:t></w:r></w:p></w:tc>` : '';

                docBody += `
                <w:tr>
                    <w:trPr><w:cantSplit/></w:trPr>
                    ${dayCellXml}
                    ${sessCellXml}
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${colWidths[2]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${slot.period}</w:t></w:r></w:p>
                    </w:tc>
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${colWidths[3]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="left"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.subject)}</w:t></w:r></w:p>
                    </w:tc>
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${colWidths[4]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.ppct || '')}</w:t></w:r></w:p>
                    </w:tc>
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${colWidths[5]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="left"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.lessonName || '')}</w:t></w:r></w:p>
                    </w:tc>
                    ${isCtlop ? `
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${colWidths[6]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                        ${formatIntegrationXml(slot.integration)}
                    </w:tc>` : ''}
                    ${signCellXml}
                    ${noteCellXml}
                    ${customCellXml}
                </w:tr>
                `;
            });
        });

        docBody += `</w:tbl>`;

        const signers = [];
        if (showBghSign) {
            signers.push({
                role: `DUYỆT CỦA ${bghSignerRole}`,
                name: bghSignerName
            });
        }
        if (showHeadSign) {
            signers.push({
                role: "TỔ TRƯỞNG CHUYÊN MÔN",
                name: settings.headOfGrade || "Trần Thị Mai"
            });
        }
        if (showGvcnSign) {
            signers.push({
                role: "GIÁO VIÊN CHỦ NHIỆM",
                name: settings.homeroomTeacher || "Nguyễn Thị Thu Hà"
            });
        }

        if (signers.length > 0) {
            const totalSigWidth = isLandscape ? 15000 : 9600;
            const singleSigWidth = Math.floor(totalSigWidth / signers.length);
            docBody += `
            <w:p><w:pPr><w:spacing w:before="240" w:after="100"/></w:pPr></w:p>
            <w:tbl>
                <w:tblPr>
                    <w:tblW w:w="0" w:type="auto"/>
                    <w:jc w:val="center"/>
                    <w:tblBorders>
                        <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
                        <w:insideH w:val="none"/><w:insideV w:val="none"/>
                    </w:tblBorders>
                </w:tblPr>
                <w:tblGrid>
                    ${signers.map(() => `<w:gridCol w:w="${singleSigWidth}"/>`).join('')}
                </w:tblGrid>
                <w:tr>
                    ${signers.map(s => `
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${singleSigWidth}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(s.role)}</w:t></w:r></w:p>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="600"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>(Ký và ghi rõ họ tên)</w:t></w:r></w:p>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(s.name)}</w:t></w:r></w:p>
                    </w:tc>
                    `).join('')}
                </w:tr>
            </w:tbl>`;
        }

        const sectPr = isLandscape
            ? `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`
            : `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1417" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;

        const fullDoc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        ${docBody}
        ${sectPr}
    </w:body>
</w:document>`;

        zip.file("word/document.xml", fullDoc);

        return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    }

    function generateMultiWeekLbgDocx(isCtlop, startWeek, endWeek, calculateWeekScheduleFn, settings, orientation = "portrait", options = {}) {
        const zip = new JSZip();

        zip.file("[Content_Types].xml", createContentTypes());
        zip.file("_rels/.rels", createRels());
        zip.file("word/_rels/document.xml.rels", createWordRels());
        zip.file("word/styles.xml", createStyles());

        const isLandscape = (orientation === "landscape");
        const headerColWidths = isLandscape ? [7200, 7800] : [4600, 5000];

        const bghSignerName = ((settings.bghSignerLbgType || settings.bghSignerType || 'PHT') === 'HT')
            ? (settings.principal || 'Phạm Quốc Hùng')
            : (Array.isArray(settings.vicePrincipals) && settings.vicePrincipals[(settings.bghSignerLbgIndex !== undefined ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0))]
                ? settings.vicePrincipals[(settings.bghSignerLbgIndex !== undefined ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0))]
                : (settings.vicePrincipal || 'Lê Văn Tám'));
        const bghSignerRole = ((settings.bghSignerLbgType || settings.bghSignerType || 'PHT') === 'HT') ? 'HIỆU TRƯỞNG' : 'BAN GIÁM HIỆU';

        const showColSign = !isCtlop && !!(options && options.showColSign);
        const showColNote = !isCtlop && !!(options && options.showColNote);
        const showColCustom = !isCtlop && !!(options && options.showColCustom);
        const colCustomName = (options && options.colCustomName) ? options.colCustomName : "Ghi chú";
        const showBghSign = (options && options.showBghSign !== undefined) ? options.showBghSign : true;
        const showHeadSign = (options && options.showHeadSign !== undefined) ? options.showHeadSign : true;
        const showGvcnSign = (options && options.showGvcnSign !== undefined) ? options.showGvcnSign : true;

        let colWidths = [];
        let headers = [];

        if (isCtlop) {
            colWidths = isLandscape
                ? [1300, 850, 600, 2400, 1000, 4350, 4500]
                : [1050, 700, 500, 1600, 750, 2550, 2450];
            headers = ["Thứ, ngày", "Buổi", "Tiết", "Môn học", "Tiết PPCT", "Tên bài dạy", "Nội dung tích hợp / Điều chỉnh"];
        } else {
            headers = ["Thứ, ngày", "Buổi", "Tiết", "Môn học", "Tiết PPCT", "Tên bài dạy"];
            if (isLandscape) {
                let wLesson = 7950;
                let wSign = 1300;
                let wNote = 2000;
                let wCustom = 2000;
                if (showColSign) wLesson -= wSign;
                if (showColNote) wLesson -= wNote;
                if (showColCustom) wLesson -= wCustom;
                colWidths = [1450, 950, 650, 2800, 1200, wLesson];
                if (showColSign) {
                    headers.push("Kí tên");
                    colWidths.push(wSign);
                }
                if (showColNote) {
                    headers.push("Ghi chú");
                    colWidths.push(wNote);
                }
                if (showColCustom) {
                    headers.push(colCustomName);
                    colWidths.push(wCustom);
                }
            } else {
                let wLesson = 4450;
                let wSign = 850;
                let wNote = 1400;
                let wCustom = 1400;
                if (showColSign) wLesson -= wSign;
                if (showColNote) wLesson -= wNote;
                if (showColCustom) wLesson -= wCustom;
                colWidths = [1150, 750, 550, 1900, 800, wLesson];
                if (showColSign) {
                    headers.push("Kí tên");
                    colWidths.push(wSign);
                }
                if (showColNote) {
                    headers.push("Ghi chú");
                    colWidths.push(wNote);
                }
                if (showColCustom) {
                    headers.push(colCustomName);
                    colWidths.push(wCustom);
                }
            }
        }

        let docBody = "";

        for (let w = startWeek; w <= endWeek; w++) {
            const { weekInfo, schedule, stats } = calculateWeekScheduleFn(w);

            docBody += `
            <w:tbl>
                <w:tblPr>
                    <w:tblW w:w="0" w:type="auto"/>
                    <w:jc w:val="center"/>
                    <w:tblBorders>
                        <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
                        <w:insideH w:val="none"/><w:insideV w:val="none"/>
                    </w:tblBorders>
                </w:tblPr>
                <w:tblGrid>
                    <w:gridCol w:w="${headerColWidths[0]}"/>
                    <w:gridCol w:w="${headerColWidths[1]}"/>
                </w:tblGrid>
                <w:tr>
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${headerColWidths[0]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                        <w:p>
                            <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr><w:t>${escapeXml((settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT').toUpperCase())}</w:t></w:r>
                        </w:p>
                        <w:p>
                            <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>${escapeXml((settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT').toUpperCase())}</w:t></w:r>
                        </w:p>
                        <w:p>
                            <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="40"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>${escapeXml((settings.grade || 'KHỐI 5').toUpperCase())} - ${escapeXml((settings.className || 'LỚP 5A').toUpperCase())}</w:t></w:r>
                        </w:p>
                    </w:tc>
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${headerColWidths[1]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                        <w:p>
                            <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</w:t></w:r>
                        </w:p>
                        <w:p>
                            <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="40"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:u w:val="single"/><w:sz w:val="23"/></w:rPr><w:t>Độc lập - Tự do - Hạnh phúc</w:t></w:r>
                        </w:p>
                    </w:tc>
                </w:tr>
            </w:tbl>

            <w:p>
                <w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="40"/></w:pPr>
                <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="30"/></w:rPr><w:t>${isCtlop ? 'LỊCH BÁO GIẢNG TÍCH HỢP' : 'LỊCH BÁO GIẢNG'} TUẦN ${w}</w:t></w:r>
            </w:p>
            <w:p>
                <w:pPr><w:jc w:val="center"/><w:spacing w:after="200"/></w:pPr>
                <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="22"/></w:rPr><w:t>(Thời gian thực hiện: Từ ngày ${escapeXml(weekInfo.startDateVN || '')} đến ngày ${escapeXml(weekInfo.endDateVN || '')})</w:t></w:r>
            </w:p>
            `;

            docBody += `
            <w:tbl>
                <w:tblPr>
                    <w:tblW w:w="0" w:type="auto"/>
                    <w:jc w:val="center"/>
                    <w:tblBorders>
                        <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                        <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                        <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    </w:tblBorders>
                </w:tblPr>
                <w:tblGrid>
                    ${colWidths.map(wd => `<w:gridCol w:w="${wd}"/>`).join('')}
                </w:tblGrid>
                <w:tr>
                    <w:trPr><w:tblHeader/></w:trPr>
                    ${headers.map((h, i) => `
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[i]}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(h)}</w:t></w:r></w:p>
                        </w:tc>
                    `).join('')}
                </w:tr>
            `;

            const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];
            days.forEach(day => {
                const daySlots = schedule.filter(s => s.day === day);
                if (daySlots.length === 0) return;

                const dayDate = getDayFullDate(weekInfo.startDateVN, day);
                const morningSlots = daySlots.filter(s => s.session === "Sáng");
                const afternoonSlots = daySlots.filter(s => s.session === "Chiều");

                daySlots.forEach((slot, idx) => {
                    const isDayStart = (idx === 0);
                    const isMornStart = (idx === 0 && morningSlots.length > 0);
                    const isAftStart = (idx === morningSlots.length && afternoonSlots.length > 0);

                    let dayCellXml = isDayStart
                        ? `<w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[0]}" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p>
                                <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="${dayDate ? '20' : '0'}"/></w:pPr>
                                <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(day)}</w:t></w:r>
                            </w:p>
                            ${dayDate ? `
                            <w:p>
                                <w:pPr><w:jc w:val="center"/><w:spacing w:line="200" w:after="0"/></w:pPr>
                                <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="333333"/></w:rPr><w:t>${escapeXml(dayDate)}</w:t></w:r>
                            </w:p>` : ''}
                           </w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="${colWidths[0]}" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr></w:p></w:tc>`;

                    let sessCellXml = (isMornStart || isAftStart)
                        ? `<w:tc><w:tcPr><w:tcW w:w="${colWidths[1]}" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${isMornStart ? 'Sáng' : 'Chiều'}</w:t></w:r></w:p></w:tc>`
                        : `<w:tc><w:tcPr><w:tcW w:w="${colWidths[1]}" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr></w:p></w:tc>`;

                    let signCellXml = showColSign ? `<w:tc><w:tcPr><w:tcW w:w="${isLandscape ? 1300 : 850}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr></w:p></w:tc>` : '';
                    let noteCellXml = showColNote ? `<w:tc><w:tcPr><w:tcW w:w="${isLandscape ? 2000 : 1400}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:line="200" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.note || '')}</w:t></w:r></w:p></w:tc>` : '';
                    let customCellXml = showColCustom ? `<w:tc><w:tcPr><w:tcW w:w="${isLandscape ? 2000 : 1400}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="left"/><w:spacing w:line="200" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.customCol || slot.note || '')}</w:t></w:r></w:p></w:tc>` : '';

                    docBody += `
                    <w:tr>
                        <w:trPr><w:cantSplit/></w:trPr>
                        ${dayCellXml}
                        ${sessCellXml}
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[2]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${slot.period}</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[3]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="left"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.subject)}</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[4]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.ppct || '')}</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[5]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="left"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.lessonName || '')}</w:t></w:r></w:p>
                        </w:tc>
                        ${isCtlop ? `
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[6]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                            ${formatIntegrationXml(slot.integration)}
                        </w:tc>` : ''}
                        ${signCellXml}
                        ${noteCellXml}
                        ${customCellXml}
                    </w:tr>
                    `;
                });
            });

            docBody += `</w:tbl>`;

            const signers = [];
            if (showBghSign) {
                signers.push({
                    role: `DUYỆT CỦA ${bghSignerRole}`,
                    name: bghSignerName
                });
            }
            if (showHeadSign) {
                signers.push({
                    role: "TỔ TRƯỞNG CHUYÊN MÔN",
                    name: settings.headOfGrade || "Trần Thị Mai"
                });
            }
            if (showGvcnSign) {
                signers.push({
                    role: "GIÁO VIÊN CHỦ NHIỆM",
                    name: settings.homeroomTeacher || "Nguyễn Thị Thu Hà"
                });
            }

            if (signers.length > 0) {
                const totalSigWidth = isLandscape ? 15000 : 9600;
                const singleSigWidth = Math.floor(totalSigWidth / signers.length);
                docBody += `
                <w:p><w:pPr><w:spacing w:before="240" w:after="100"/></w:pPr></w:p>
                <w:tbl>
                    <w:tblPr>
                        <w:tblW w:w="0" w:type="auto"/>
                        <w:jc w:val="center"/>
                        <w:tblBorders>
                            <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
                            <w:insideH w:val="none"/><w:insideV w:val="none"/>
                        </w:tblBorders>
                    </w:tblPr>
                    <w:tblGrid>
                        ${signers.map(() => `<w:gridCol w:w="${singleSigWidth}"/>`).join('')}
                    </w:tblGrid>
                    <w:tr>
                        ${signers.map(s => `
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${singleSigWidth}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(s.role)}</w:t></w:r></w:p>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="600"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>(Ký và ghi rõ họ tên)</w:t></w:r></w:p>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(s.name)}</w:t></w:r></w:p>
                        </w:tc>
                        `).join('')}
                    </w:tr>
                </w:tbl>
                `;
            }

            if (w < endWeek) {
                docBody += `
                <w:p>
                    <w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>
                    <w:r><w:br w:type="page"/></w:r>
                </w:p>
                `;
            }
        }

        const sectPr = isLandscape
            ? `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`
            : `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1417" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;

        const fullDoc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        ${docBody}
        ${sectPr}
    </w:body>
</w:document>`;

        zip.file("word/document.xml", fullDoc);

        return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    }

    function generateTimetableDocx(slots, settings, customTitle, customSubtitle) {
        const zip = new JSZip();

        zip.file("[Content_Types].xml", createContentTypes());
        zip.file("_rels/.rels", createRels());
        zip.file("word/_rels/document.xml.rels", createWordRels());
        zip.file("word/styles.xml", createStyles());

        const bghSignerName = ((settings.bghSignerLbgType || settings.bghSignerType || 'PHT') === 'HT')
            ? (settings.principal || 'Phạm Quốc Hùng')
            : (Array.isArray(settings.vicePrincipals) && settings.vicePrincipals[(settings.bghSignerLbgIndex !== undefined ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0))]
                ? settings.vicePrincipals[(settings.bghSignerLbgIndex !== undefined ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0))]
                : (settings.vicePrincipal || 'Lê Văn Tám'));
        const bghSignerRole = ((settings.bghSignerLbgType || settings.bghSignerType || 'PHT') === 'HT') ? 'HIỆU TRƯỞNG' : 'BAN GIÁM HIỆU';

        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];
        const colWidths = [1700, 1200, 2320, 2320, 2320, 2320, 2320];

        // Find max periods for morning and afternoon
        let maxMorn = 4;
        let maxAft = 3;
        slots.forEach(s => {
            if (s.session === "Sáng" && s.period > maxMorn) maxMorn = s.period;
            if (s.session === "Chiều" && s.period > maxAft) maxAft = s.period;
        });

        let docBody = `
        <w:tbl>
            <w:tblPr>
                <w:tblW w:w="0" w:type="auto"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
                    <w:insideH w:val="none"/><w:insideV w:val="none"/>
                </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="6800"/>
                <w:gridCol w:w="7700"/>
            </w:tblGrid>
            <w:tr>
                <w:tc>
                    <w:tcPr><w:tcW w:w="6800" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="23"/></w:rPr><w:t>${escapeXml((settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT').toUpperCase())}</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="23"/></w:rPr><w:t>${escapeXml((settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT').toUpperCase())}</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="40"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="23"/></w:rPr><w:t>${escapeXml((settings.grade || 'KHỐI 5').toUpperCase())} - ${escapeXml((settings.className || 'LỚP 5A').toUpperCase())}</w:t></w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="7700" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="23"/></w:rPr><w:t>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="40"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:u w:val="single"/><w:sz w:val="24"/></w:rPr><w:t>Độc lập - Tự do - Hạnh phúc</w:t></w:r>
                    </w:p>
                </w:tc>
            </w:tr>
        </w:tbl>

        <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="40"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="32"/></w:rPr><w:t>${escapeXml((customTitle || ('THỜI KHÓA BIỂU ' + (settings.grade || 'KHỐI 5') + ' - ' + (settings.className || 'LỚP 5A'))).toUpperCase())}</w:t></w:r>
        </w:p>
        <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:after="200"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="23"/></w:rPr><w:t>${escapeXml(customSubtitle || ('Năm học ' + (settings.academicYear || '2026 - 2027')))}</w:t></w:r>
        </w:p>

        <w:tbl>
            <w:tblPr>
                <w:tblW w:w="0" w:type="auto"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>
                ${colWidths.map(wd => `<w:gridCol w:w="${wd}"/>`).join('')}
            </w:tblGrid>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc>
                    <w:tcPr><w:tcW w:w="${colWidths[0]}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:vAlign w:val="center"/></w:tcPr>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Buổi</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="${colWidths[1]}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:vAlign w:val="center"/></w:tcPr>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Tiết</w:t></w:r></w:p>
                </w:tc>
                ${days.map((d, i) => `
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${colWidths[i + 2]}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:vAlign w:val="center"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(d)}</w:t></w:r></w:p>
                    </w:tc>
                `).join('')}
            </w:tr>
        `;

        // Morning rows
        for (let p = 1; p <= maxMorn; p++) {
            const vMerge = (p === 1) ? `<w:vMerge w:val="restart"/>` : `<w:vMerge/>`;
            docBody += `
            <w:tr>
                <w:trPr><w:cantSplit/></w:trPr>
                <w:tc>
                    <w:tcPr><w:tcW w:w="${colWidths[0]}" w:type="dxa"/>${vMerge}<w:vAlign w:val="center"/></w:tcPr>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${(p === 1 ? 'Sáng' : '')}</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="${colWidths[1]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${p}</w:t></w:r></w:p>
                </w:tc>
                ${days.map((d, i) => {
                    const isAltDay = (d === "Thứ 2" || d === "Thứ 4" || d === "Thứ 6");
                    const dayMornSlots = slots.filter(s => s.day === d && s.session === "Sáng").sort((a, b) => (a.period || 0) - (b.period || 0));
                    const slot = dayMornSlots[p - 1];
                    const sub = slot ? slot.subject : "";
                    const shdXml = isAltDay ? `<w:shd w:val="clear" w:color="auto" w:fill="F8FAFC"/>` : ``;
                    return `
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${colWidths[i + 2]}" w:type="dxa"/>${shdXml}<w:vAlign w:val="center"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(sub)}</w:t></w:r></w:p>
                    </w:tc>`;
                }).join('')}
            </w:tr>
            `;
        }

        // Afternoon rows
        for (let p = 1; p <= maxAft; p++) {
            const vMerge = (p === 1) ? `<w:vMerge w:val="restart"/>` : `<w:vMerge/>`;
            docBody += `
            <w:tr>
                <w:trPr><w:cantSplit/></w:trPr>
                <w:tc>
                    <w:tcPr><w:tcW w:w="${colWidths[0]}" w:type="dxa"/>${vMerge}<w:vAlign w:val="center"/></w:tcPr>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${(p === 1 ? 'Chiều' : '')}</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="${colWidths[1]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${p}</w:t></w:r></w:p>
                </w:tc>
                ${days.map((d, i) => {
                    const isAltDay = (d === "Thứ 2" || d === "Thứ 4" || d === "Thứ 6");
                    const dayAftSlots = slots.filter(s => s.day === d && s.session === "Chiều").sort((a, b) => (a.period || 0) - (b.period || 0));
                    const slot = dayAftSlots[p - 1];
                    const sub = slot ? slot.subject : "";
                    const shdXml = isAltDay ? `<w:shd w:val="clear" w:color="auto" w:fill="F8FAFC"/>` : ``;
                    return `
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${colWidths[i + 2]}" w:type="dxa"/>${shdXml}<w:vAlign w:val="center"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(sub)}</w:t></w:r></w:p>
                    </w:tc>`;
                }).join('')}
            </w:tr>
            `;
        }

        docBody += `</w:tbl>`;

        // Signatures (Landscape 3 columns: 4800, 4800, 4900)
        docBody += `
        <w:p><w:pPr><w:spacing w:before="240" w:after="100"/></w:pPr></w:p>
        <w:tbl>
            <w:tblPr>
                <w:tblW w:w="0" w:type="auto"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
                    <w:insideH w:val="none"/><w:insideV w:val="none"/>
                </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="4800"/>
                <w:gridCol w:w="4800"/>
                <w:gridCol w:w="4900"/>
            </w:tblGrid>
            <w:tr>
                <w:tc>
                    <w:tcPr><w:tcW w:w="4800" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>DUYỆT CỦA ${escapeXml(bghSignerRole)}</w:t></w:r></w:p>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="600"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>(Ký và ghi rõ họ tên)</w:t></w:r></w:p>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(bghSignerName)}</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="4800" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>TỔ TRƯỞNG CHUYÊN MÔN</w:t></w:r></w:p>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="600"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>(Ký và ghi rõ họ tên)</w:t></w:r></w:p>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(settings.headOfGrade || 'Trần Thị Mai')}</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="4900" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>GIÁO VIÊN CHỦ NHIỆM</w:t></w:r></w:p>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="600"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>(Ký và ghi rõ họ tên)</w:t></w:r></w:p>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(settings.homeroomTeacher || 'Nguyễn Thị Thu Hà')}</w:t></w:r></w:p>
                </w:tc>
            </w:tr>
        </w:tbl>
        `;

        const fullDoc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        ${docBody}
        <w:sectPr>
            <w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>
            <w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/>
        </w:sectPr>
    </w:body>
</w:document>`;

        zip.file("word/document.xml", fullDoc);

        return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    }


    function getDayDateStrBySubject(startDateVN, dayStr) {
        if (!startDateVN) return dayStr;
        const parts = startDateVN.split('/');
        if (parts.length < 3) return dayStr;
        const startD = parseInt(parts[0], 10);
        const startM = parseInt(parts[1], 10) - 1;
        const startY = parseInt(parts[2], 10);
        const baseDate = new Date(startY, startM, startD);
        const dayOffsets = { "Thứ 2": 0, "Thứ 3": 1, "Thứ 4": 2, "Thứ 5": 3, "Thứ 6": 4, "Thứ 7": 5, "Chủ nhật": 6 };
        const offset = dayOffsets[dayStr] !== undefined ? dayOffsets[dayStr] : 0;
        const targetDate = new Date(baseDate.getTime() + offset * 86400000);
        const dStr = String(targetDate.getDate()).padStart(2, '0');
        const mStr = String(targetDate.getMonth() + 1).padStart(2, '0');
        return `${dayStr} (${dStr}/${mStr})`;
    }

    function generateLbgBySubjectDocx(data, settings, isCtlop = true, orientation = "portrait", filterSubject = "all") {
        const zip = new JSZip();

        zip.file("[Content_Types].xml", createContentTypes());
        zip.file("_rels/.rels", createRels());
        zip.file("word/_rels/document.xml.rels", createWordRels());
        zip.file("word/styles.xml", createStyles());

        const isLandscape = (orientation === "landscape");
        const headerColWidths = isLandscape ? [7200, 7800] : [4600, 5000];
        const sigColWidth = isLandscape ? "5000" : "3200";

        const bghSignerName = ((settings.bghSignerLbgType || settings.bghSignerType || 'PHT') === 'HT')
            ? (settings.principal || 'Phạm Quốc Hùng')
            : (Array.isArray(settings.vicePrincipals) && settings.vicePrincipals[(settings.bghSignerLbgIndex !== undefined ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0))]
                ? settings.vicePrincipals[(settings.bghSignerLbgIndex !== undefined ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0))]
                : (settings.vicePrincipal || 'Lê Văn Tám'));
        const bghSignerRole = ((settings.bghSignerLbgType || settings.bghSignerType || 'PHT') === 'HT') ? 'HIỆU TRƯỞNG' : 'BAN GIÁM HIỆU';

        const weekNum = data.weekNum || 1;
        const weekInfo = data.weekInfo || {};
        const subjectGroups = data.subjectGroups || [];

        let titleText = `LỊCH BÁO GIẢNG THEO MÔN HỌC TUẦN ${weekNum}`;
        if (filterSubject && filterSubject !== "all") {
            titleText = `LỊCH BÁO GIẢNG MÔN ${filterSubject.toUpperCase()} TUẦN ${weekNum}`;
        }

        let docBody = `
        <w:tbl>
            <w:tblPr>
                <w:tblW w:w="0" w:type="auto"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
                    <w:insideH w:val="none"/><w:insideV w:val="none"/>
                </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="${headerColWidths[0]}"/>
                <w:gridCol w:w="${headerColWidths[1]}"/>
            </w:tblGrid>
            <w:tr>
                <w:tc>
                    <w:tcPr><w:tcW w:w="${headerColWidths[0]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr><w:t>${escapeXml((settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT').toUpperCase())}</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>${escapeXml((settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT').toUpperCase())}</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="40"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>${escapeXml((settings.grade || 'KHỐI 5').toUpperCase())} - ${escapeXml((settings.className || 'LỚP 5A').toUpperCase())}</w:t></w:r>
                    </w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="${headerColWidths[1]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</w:t></w:r>
                    </w:p>
                    <w:p>
                        <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="40"/></w:pPr>
                        <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:u w:val="single"/><w:sz w:val="23"/></w:rPr><w:t>Độc lập - Tự do - Hạnh phúc</w:t></w:r>
                    </w:p>
                </w:tc>
            </w:tr>
        </w:tbl>

        <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="40"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="30"/></w:rPr><w:t>${escapeXml(titleText)}</w:t></w:r>
        </w:p>
        <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:after="200"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="22"/></w:rPr><w:t>(Thời gian thực hiện: Từ ngày ${escapeXml(weekInfo.startDateVN || '')} đến ngày ${escapeXml(weekInfo.endDateVN || '')})</w:t></w:r>
        </w:p>
        `;

        const colWidths = isLandscape
            ? (isCtlop ? [2200, 1500, 1000, 800, 900, 1100, 3500, 3600] : [2600, 1800, 1100, 900, 1000, 1200, 6000])
            : (isCtlop ? [1600, 1100, 700, 550, 650, 750, 2150, 2100] : [1800, 1300, 800, 600, 750, 900, 3500]);
        const headers = isCtlop
            ? ["Môn học", "Thứ / Ngày", "Buổi", "Tiết", "Tiết/tuần", "Tiết PPCT", "Tên bài dạy", "Nội dung tích hợp / Điều chỉnh"]
            : ["Môn học", "Thứ / Ngày", "Buổi", "Tiết", "Tiết/tuần", "Tiết PPCT", "Tên bài dạy"];

        docBody += `
        <w:tbl>
            <w:tblPr>
                <w:tblW w:w="0" w:type="auto"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>
                ${colWidths.map(w => `<w:gridCol w:w="${w}"/>`).join('')}
            </w:tblGrid>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                ${headers.map((h, i) => `
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${colWidths[i]}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:vAlign w:val="center"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(h)}</w:t></w:r></w:p>
                    </w:tc>
                `).join('')}
            </w:tr>
        `;

        if (subjectGroups.length === 0) {
            docBody += `
            <w:tr>
                <w:tc>
                    <w:tcPr><w:gridSpan w:val="${headers.length}"/><w:vAlign w:val="center"/></w:tcPr>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Không có tiết học nào phù hợp với bộ lọc đã chọn.</w:t></w:r></w:p>
                </w:tc>
            </w:tr>
            `;
        } else {
            subjectGroups.forEach(group => {
                group.slots.forEach((slot, idx) => {
                    const isGroupStart = (idx === 0);
                    const dayDateText = getDayDateStrBySubject(weekInfo.startDateVN, slot.day);

                    let subCellXml = isGroupStart
                        ? `<w:tc>
                            <w:tcPr>
                                <w:tcW w:w="${colWidths[0]}" w:type="dxa"/>
                                <w:vMerge w:val="restart"/>
                                <w:vAlign w:val="center"/>
                                <w:shd w:val="clear" w:color="auto" w:fill="FAFAFA"/>
                            </w:tcPr>
                            <w:p>
                                <w:pPr><w:jc w:val="left"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                                <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(group.subjectName)}</w:t></w:r>
                            </w:p>
                            <w:p>
                                <w:pPr><w:jc w:val="left"/><w:spacing w:line="200" w:after="0"/></w:pPr>
                                <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="475569"/></w:rPr><w:t>(${group.slots.length} tiết)</w:t></w:r>
                            </w:p>
                           </w:tc>`
                        : `<w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[0]}" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr></w:p>
                           </w:tc>`;

                    docBody += `
                    <w:tr>
                        <w:trPr><w:cantSplit/></w:trPr>
                        ${subCellXml}
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[1]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(dayDateText)}</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[2]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.session || '')}</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[3]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${slot.period || ''}</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[4]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${slot.periodInWeek || (idx + 1)}</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[5]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.ppct || '')}</w:t></w:r></w:p>
                        </w:tc>
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[6]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="left"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.lessonName || '')}</w:t></w:r></w:p>
                        </w:tc>
                        ${isCtlop ? `
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[7]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                            ${formatIntegrationXml(slot.integration)}
                        </w:tc>` : ''}
                    </w:tr>
                    `;
                });
            });
        }

        docBody += `</w:tbl>`;

        docBody += `
        <w:p><w:pPr><w:spacing w:before="240" w:after="100"/></w:pPr></w:p>
        <w:tbl>
            <w:tblPr>
                <w:tblW w:w="0" w:type="auto"/>
                <w:jc w:val="center"/>
                <w:tblBorders>
                    <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
                    <w:insideH w:val="none"/><w:insideV w:val="none"/>
                </w:tblBorders>
            </w:tblPr>
            <w:tblGrid>
                <w:gridCol w:w="${sigColWidth}"/>
                <w:gridCol w:w="${sigColWidth}"/>
                <w:gridCol w:w="${sigColWidth}"/>
            </w:tblGrid>
            <w:tr>
                <w:tc>
                    <w:tcPr><w:tcW w:w="${sigColWidth}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>DUYỆT CỦA ${escapeXml(bghSignerRole)}</w:t></w:r></w:p>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="600"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>(Ký và ghi rõ họ tên)</w:t></w:r></w:p>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(bghSignerName)}</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="${sigColWidth}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>TỔ TRƯỞNG CHUYÊN MÔN</w:t></w:r></w:p>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="600"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>(Ký và ghi rõ họ tên)</w:t></w:r></w:p>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(settings.headOfGrade || 'Trần Thị Mai')}</w:t></w:r></w:p>
                </w:tc>
                <w:tc>
                    <w:tcPr><w:tcW w:w="${sigColWidth}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>GIÁO VIÊN CHỦ NHIỆM</w:t></w:r></w:p>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="600"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>(Ký và ghi rõ họ tên)</w:t></w:r></w:p>
                    <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(settings.homeroomTeacher || 'Nguyễn Thị Thu Hà')}</w:t></w:r></w:p>
                </w:tc>
            </w:tr>
        </w:tbl>
        `;

        const sectPr = isLandscape
            ? `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`
            : `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1417" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;

        const fullDoc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        ${docBody}
        ${sectPr}
    </w:body>
</w:document>`;

        zip.file("word/document.xml", fullDoc);

        return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    }

    function generateBatchLbgBySubjectDocx(startWeek, endWeek, calculateWeekScheduleBySubjectFn, settings, orientation = "portrait", isCtlop = true, filterSubject = "all", filterCategory = "all") {
        const zip = new JSZip();

        zip.file("[Content_Types].xml", createContentTypes());
        zip.file("_rels/.rels", createRels());
        zip.file("word/_rels/document.xml.rels", createWordRels());
        zip.file("word/styles.xml", createStyles());

        const isLandscape = (orientation === "landscape");
        const headerColWidths = isLandscape ? [7200, 7800] : [4600, 5000];
        const sigColWidth = isLandscape ? "5000" : "3200";

        const bghSignerName = ((settings.bghSignerLbgType || settings.bghSignerType || 'PHT') === 'HT')
            ? (settings.principal || 'Phạm Quốc Hùng')
            : (Array.isArray(settings.vicePrincipals) && settings.vicePrincipals[(settings.bghSignerLbgIndex !== undefined ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0))]
                ? settings.vicePrincipals[(settings.bghSignerLbgIndex !== undefined ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0))]
                : (settings.vicePrincipal || 'Lê Văn Tám'));
        const bghSignerRole = ((settings.bghSignerLbgType || settings.bghSignerType || 'PHT') === 'HT') ? 'HIỆU TRƯỞNG' : 'BAN GIÁM HIỆU';

        const colWidths = isLandscape
            ? (isCtlop ? [2200, 1500, 1000, 800, 900, 1100, 3500, 3600] : [2600, 1800, 1100, 900, 1000, 1200, 6000])
            : (isCtlop ? [1600, 1100, 700, 550, 650, 750, 2150, 2100] : [1800, 1300, 800, 600, 750, 900, 3500]);
        const headers = isCtlop
            ? ["Môn học", "Thứ / Ngày", "Buổi", "Tiết", "Tiết/tuần", "Tiết PPCT", "Tên bài dạy", "Nội dung tích hợp / Điều chỉnh"]
            : ["Môn học", "Thứ / Ngày", "Buổi", "Tiết", "Tiết/tuần", "Tiết PPCT", "Tên bài dạy"];

        let docBody = "";

        for (let w = startWeek; w <= endWeek; w++) {
            const data = calculateWeekScheduleBySubjectFn(w, filterSubject, filterCategory);
            const weekInfo = data.weekInfo || {};
            const subjectGroups = data.subjectGroups || [];

            let titleText = `LỊCH BÁO GIẢNG THEO MÔN HỌC TUẦN ${w}`;
            if (filterSubject && filterSubject !== "all") {
                titleText = `LỊCH BÁO GIẢNG MÔN ${filterSubject.toUpperCase()} TUẦN ${w}`;
            }

            docBody += `
            <w:tbl>
                <w:tblPr>
                    <w:tblW w:w="0" w:type="auto"/>
                    <w:jc w:val="center"/>
                    <w:tblBorders>
                        <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
                        <w:insideH w:val="none"/><w:insideV w:val="none"/>
                    </w:tblBorders>
                </w:tblPr>
                <w:tblGrid>
                    <w:gridCol w:w="${headerColWidths[0]}"/>
                    <w:gridCol w:w="${headerColWidths[1]}"/>
                </w:tblGrid>
                <w:tr>
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${headerColWidths[0]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                        <w:p>
                            <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr><w:t>${escapeXml((settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT').toUpperCase())}</w:t></w:r>
                        </w:p>
                        <w:p>
                            <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>${escapeXml((settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT').toUpperCase())}</w:t></w:r>
                        </w:p>
                        <w:p>
                            <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="40"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>${escapeXml((settings.grade || 'KHỐI 5').toUpperCase())} - ${escapeXml((settings.className || 'LỚP 5A').toUpperCase())}</w:t></w:r>
                        </w:p>
                    </w:tc>
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${headerColWidths[1]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                        <w:p>
                            <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</w:t></w:r>
                        </w:p>
                        <w:p>
                            <w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="40"/></w:pPr>
                            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:u w:val="single"/><w:sz w:val="23"/></w:rPr><w:t>Độc lập - Tự do - Hạnh phúc</w:t></w:r>
                        </w:p>
                    </w:tc>
                </w:tr>
            </w:tbl>

            <w:p>
                <w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="40"/></w:pPr>
                <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="30"/></w:rPr><w:t>${escapeXml(titleText)}</w:t></w:r>
            </w:p>
            <w:p>
                <w:pPr><w:jc w:val="center"/><w:spacing w:after="200"/></w:pPr>
                <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="22"/></w:rPr><w:t>(Thời gian thực hiện: Từ ngày ${escapeXml(weekInfo.startDateVN || '')} đến ngày ${escapeXml(weekInfo.endDateVN || '')})</w:t></w:r>
            </w:p>
            `;

            docBody += `
            <w:tbl>
                <w:tblPr>
                    <w:tblW w:w="0" w:type="auto"/>
                    <w:jc w:val="center"/>
                    <w:tblBorders>
                        <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                        <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                        <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                    </w:tblBorders>
                </w:tblPr>
                <w:tblGrid>
                    ${colWidths.map(wdx => `<w:gridCol w:w="${wdx}"/>`).join('')}
                </w:tblGrid>
                <w:tr>
                    <w:trPr><w:tblHeader/></w:trPr>
                    ${headers.map((h, i) => `
                        <w:tc>
                            <w:tcPr><w:tcW w:w="${colWidths[i]}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:vAlign w:val="center"/></w:tcPr>
                            <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(h)}</w:t></w:r></w:p>
                        </w:tc>
                    `).join('')}
                </w:tr>
            `;

            if (subjectGroups.length === 0) {
                docBody += `
                <w:tr>
                    <w:tc>
                        <w:tcPr><w:gridSpan w:val="${headers.length}"/><w:vAlign w:val="center"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>Không có tiết học nào phù hợp với bộ lọc đã chọn.</w:t></w:r></w:p>
                    </w:tc>
                </w:tr>
                `;
            } else {
                subjectGroups.forEach(group => {
                    group.slots.forEach((slot, idx) => {
                        const isGroupStart = (idx === 0);
                        const dayDateText = getDayDateStrBySubject(weekInfo.startDateVN, slot.day);

                        let subCellXml = isGroupStart
                            ? `<w:tc>
                                <w:tcPr>
                                    <w:tcW w:w="${colWidths[0]}" w:type="dxa"/>
                                    <w:vMerge w:val="restart"/>
                                    <w:vAlign w:val="center"/>
                                    <w:shd w:val="clear" w:color="auto" w:fill="FAFAFA"/>
                                </w:tcPr>
                                <w:p>
                                    <w:pPr><w:jc w:val="left"/><w:spacing w:line="220" w:after="20"/></w:pPr>
                                    <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(group.subjectName)}</w:t></w:r>
                                </w:p>
                                <w:p>
                                    <w:pPr><w:jc w:val="left"/><w:spacing w:line="200" w:after="0"/></w:pPr>
                                    <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="475569"/></w:rPr><w:t>(${group.slots.length} tiết)</w:t></w:r>
                                </w:p>
                               </w:tc>`
                            : `<w:tc>
                                <w:tcPr><w:tcW w:w="${colWidths[0]}" w:type="dxa"/><w:vMerge/><w:vAlign w:val="center"/></w:tcPr>
                                <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr></w:p>
                               </w:tc>`;

                        docBody += `
                        <w:tr>
                            <w:trPr><w:cantSplit/></w:trPr>
                            ${subCellXml}
                            <w:tc>
                                <w:tcPr><w:tcW w:w="${colWidths[1]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                                <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(dayDateText)}</w:t></w:r></w:p>
                            </w:tc>
                            <w:tc>
                                <w:tcPr><w:tcW w:w="${colWidths[2]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                                <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.session || '')}</w:t></w:r></w:p>
                            </w:tc>
                            <w:tc>
                                <w:tcPr><w:tcW w:w="${colWidths[3]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                                <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${slot.period || ''}</w:t></w:r></w:p>
                            </w:tc>
                            <w:tc>
                                <w:tcPr><w:tcW w:w="${colWidths[4]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                                <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${slot.periodInWeek || (idx + 1)}</w:t></w:r></w:p>
                            </w:tc>
                            <w:tc>
                                <w:tcPr><w:tcW w:w="${colWidths[5]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                                <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.ppct || '')}</w:t></w:r></w:p>
                            </w:tc>
                            <w:tc>
                                <w:tcPr><w:tcW w:w="${colWidths[6]}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                                <w:p><w:pPr><w:jc w:val="left"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(slot.lessonName || '')}</w:t></w:r></w:p>
                            </w:tc>
                            ${isCtlop ? `
                            <w:tc>
                                <w:tcPr><w:tcW w:w="${colWidths[7]}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                                ${formatIntegrationXml(slot.integration)}
                            </w:tc>` : ''}
                        </w:tr>
                        `;
                    });
                });
            }

            docBody += `</w:tbl>`;

            docBody += `
            <w:p><w:pPr><w:spacing w:before="240" w:after="100"/></w:pPr></w:p>
            <w:tbl>
                <w:tblPr>
                    <w:tblW w:w="0" w:type="auto"/>
                    <w:jc w:val="center"/>
                    <w:tblBorders>
                        <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
                        <w:insideH w:val="none"/><w:insideV w:val="none"/>
                    </w:tblBorders>
                </w:tblPr>
                <w:tblGrid>
                    <w:gridCol w:w="${sigColWidth}"/>
                    <w:gridCol w:w="${sigColWidth}"/>
                    <w:gridCol w:w="${sigColWidth}"/>
                </w:tblGrid>
                <w:tr>
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${sigColWidth}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>DUYỆT CỦA ${escapeXml(bghSignerRole)}</w:t></w:r></w:p>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="600"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>(Ký và ghi rõ họ tên)</w:t></w:r></w:p>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(bghSignerName)}</w:t></w:r></w:p>
                    </w:tc>
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${sigColWidth}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>TỔ TRƯỞNG CHUYÊN MÔN</w:t></w:r></w:p>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="600"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>(Ký và ghi rõ họ tên)</w:t></w:r></w:p>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(settings.headOfGrade || 'Trần Thị Mai')}</w:t></w:r></w:p>
                    </w:tc>
                    <w:tc>
                        <w:tcPr><w:tcW w:w="${sigColWidth}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>GIÁO VIÊN CHỦ NHIỆM</w:t></w:r></w:p>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="600"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:i/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>(Ký và ghi rõ họ tên)</w:t></w:r></w:p>
                        <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="220" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr><w:t>${escapeXml(settings.homeroomTeacher || 'Nguyễn Thị Thu Hà')}</w:t></w:r></w:p>
                    </w:tc>
                </w:tr>
            </w:tbl>
            `;

            if (w < endWeek) {
                docBody += `
                <w:p>
                    <w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>
                    <w:r><w:br w:type="page"/></w:r>
                </w:p>
                `;
            }
        }

        const sectPr = isLandscape
            ? `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`
            : `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1417" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;

        const fullDoc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        ${docBody}
        ${sectPr}
    </w:body>
</w:document>`;

        zip.file("word/document.xml", fullDoc);

        return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    }


    return {
        generateKhdhDocx: generateKhdhDocx,
        generateLbgDocx: generateLbgDocx,
        generateMultiWeekLbgDocx: generateMultiWeekLbgDocx,
        generateTimetableDocx: generateTimetableDocx,
        generateLbgBySubjectDocx: generateLbgBySubjectDocx,
        generateBatchLbgBySubjectDocx: generateBatchLbgBySubjectDocx
    };
})();
