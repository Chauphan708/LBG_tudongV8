/**
 * XLSX Generator for Lịch Báo Giảng & Báo Giảng Tích Hợp (CTLOP)
 * Conforming strictly to Decree 30/2020/ND-CP & Standard Print Layout
 * Generates styled OpenXML .xlsx files with full formatting, merges, borders, column widths, and A4 print setup
 * Uses JSZip and FileSaver for 100% offline generation in browser
 */

window.XlsxGenerator = (function() {

    function escapeXml(unsafe) {
        if (unsafe === undefined || unsafe === null) return "";
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
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlpackage.core-properties+xml"/>
</Types>`;
    }

    function createRels() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
    }

    function createAppXml() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
    <Application>Phần Mềm Quản Lý Lịch Báo Giảng Lớp 5</Application>
</Properties>`;
    }

    function createCoreXml(title) {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlpackage.core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:title>${escapeXml(title)}</dc:title>
    <dc:creator>Phần Mềm LBG Lớp 5</dc:creator>
</cp:coreProperties>`;
    }

    function createWbRels() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
    }

    function createWbXml(sheetName) {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="${escapeXml(sheetName)}" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>`;
    }

    function createStylesXml() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="9">
        <!-- 0: Times New Roman 13pt Regular -->
        <font><name val="Times New Roman"/><sz val="13"/><color rgb="FF000000"/></font>
        <!-- 1: Times New Roman 13pt Bold -->
        <font><name val="Times New Roman"/><sz val="13"/><b/><color rgb="FF000000"/></font>
        <!-- 2: Times New Roman 13pt Bold Underline -->
        <font><name val="Times New Roman"/><sz val="13"/><b/><u/><color rgb="FF000000"/></font>
        <!-- 3: Times New Roman 13pt Bold -->
        <font><name val="Times New Roman"/><sz val="13"/><b/><color rgb="FF000000"/></font>
        <!-- 4: Times New Roman 13pt Italic -->
        <font><name val="Times New Roman"/><sz val="13"/><i/><color rgb="FF000000"/></font>
        <!-- 5: Times New Roman 13pt Regular -->
        <font><name val="Times New Roman"/><sz val="13"/><color rgb="FF000000"/></font>
        <!-- 6: Times New Roman 13pt Bold Underline -->
        <font><name val="Times New Roman"/><sz val="13"/><b/><u/><color rgb="FF000000"/></font>
        <!-- 7: Times New Roman 12pt Italic -->
        <font><name val="Times New Roman"/><sz val="12"/><i/><color rgb="FF000000"/></font>
        <!-- 8: Times New Roman 13pt Bold Underline (for School Header) -->
        <font><name val="Times New Roman"/><sz val="13"/><b/><u/><color rgb="FF000000"/></font>
    </fonts>
    <fills count="3">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFF2F4F8"/></patternFill></fill>
    </fills>
    <borders count="2">
        <!-- 0: None -->
        <border><left/><right/><top/><bottom/></border>
        <!-- 1: Thin border all 4 sides -->
        <border>
            <left style="thin"><color rgb="FF000000"/></left>
            <right style="thin"><color rgb="FF000000"/></right>
            <top style="thin"><color rgb="FF000000"/></top>
            <bottom style="thin"><color rgb="FF000000"/></bottom>
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="20">
        <!-- 0: Default -->
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
        <!-- 1: Header Gov Body (Left/Center, 10.5pt) -->
        <xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
        <!-- 2: Header School (Bold, Underline, 10.5pt) -->
        <xf numFmtId="0" fontId="8" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
        <!-- 3: Header Grade-Class (Bold, 10.5pt) -->
        <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
        <!-- 4: Header Nation (Bold, Center) -->
        <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
        <!-- 5: Header Slogan (Bold, Underline, Center) -->
        <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
        <!-- 6: Title Main (13pt Bold, Center) -->
        <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
        <!-- 7: Subtitle (Italic, Center) -->
        <xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
        <!-- 8: Table Header (Bold, Gray Fill, Border, Center H/V, Wrap) -->
        <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
        <!-- 9: Table Day (Bold, Border, Center H/V, Wrap) -->
        <xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
        <!-- 10: Table Session (Border, Center H/V, Wrap) -->
        <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
        <!-- 11: Table Period (Bold, Border, Center H/V) -->
        <xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
        <!-- 12: Table Subject (Bold, Border, Center H, Center V, Wrap) -->
        <xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
        <!-- 13: Table PPCT (Bold, Border, Center H/V) -->
        <xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
        <!-- 14: Table Lesson (Border, Left H, Center V, Wrap) -->
        <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
        <!-- 15: Table Integration (10pt, Border, Left H, Center V, Wrap) -->
        <xf numFmtId="0" fontId="5" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
        <!-- 16: Footer Title (Bold, Center) -->
        <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
        <!-- 17: Footer Sub (10pt Italic, Center) -->
        <xf numFmtId="0" fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
        <!-- 18: Footer Name (Bold, Center) -->
        <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
        <!-- 19: Total Row (Bold, Border, Left H, Center V) -->
        <xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    </cellXfs>
</styleSheet>`;
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

    function getColLetter(colIdx) {
        let letter = "";
        let n = colIdx;
        while (n > 0) {
            const mod = (n - 1) % 26;
            letter = String.fromCharCode(65 + mod) + letter;
            n = Math.floor((n - mod) / 26);
        }
        return letter;
    }

    function getXlsxLbgOrderedColumns(customColsList, showSign, showNote, isCtlop) {
        const enabledCustomCols = Array.isArray(customColsList) 
            ? customColsList.filter(c => c && c.enabled !== false) 
            : [];
        
        const cols = [];
        
        const pushCustom = (pos) => {
            enabledCustomCols.filter(c => String(c.pos) === String(pos)).forEach(c => {
                cols.push({
                    key: 'custom_' + c.id,
                    id: c.id,
                    title: c.name || 'Cột mới',
                    isCustom: true,
                    pos: c.pos,
                    width: 22
                });
            });
        };

        // Pos 1: before day
        pushCustom('1');

        // Day
        cols.push({ key: 'day', title: 'Thứ, ngày', isCustom: false, width: 13 });

        // Pos 2: after day
        pushCustom('2');

        // Session
        cols.push({ key: 'session', title: 'Buổi', isCustom: false, width: 9 });

        // Pos 3: after session
        pushCustom('3');

        // Period
        cols.push({ key: 'period', title: 'Tiết', isCustom: false, width: 7 });

        // Pos 4: after period
        pushCustom('4');

        // Subject
        cols.push({ key: 'subject', title: 'Môn học', isCustom: false, width: isCtlop ? 18 : 20 });

        // Pos 5: after subject
        pushCustom('5');

        // PPCT
        cols.push({ key: 'ppct', title: 'Tiết PPCT', isCustom: false, width: 10 });

        // Pos 6: after ppct
        pushCustom('6');

        // Lesson
        cols.push({ key: 'lesson', title: 'Tên bài dạy', isCustom: false, width: 44 });

        // Pos 7: after lesson
        pushCustom('7');

        if (isCtlop) {
            cols.push({ key: 'integ', title: 'Nội dung tích hợp / Điều chỉnh', isCustom: false, width: 30 });
        }
        if (showSign) {
            cols.push({ key: 'sign', title: 'Kí tên', isCustom: false, width: 12 });
        }
        if (showNote) {
            cols.push({ key: 'note', title: 'Ghi chú', isCustom: false, width: 22 });
        }

        // Pos end: at the end of table
        enabledCustomCols.filter(c => !['1', '2', '3', '4', '5', '6', '7'].includes(String(c.pos))).forEach(c => {
            cols.push({
                key: 'custom_' + c.id,
                id: c.id,
                title: c.name || 'Cột mới',
                isCustom: true,
                pos: c.pos || 'end',
                width: 22
            });
        });

        const customCount = enabledCustomCols.length;
        const lessonCol = cols.find(c => c.key === 'lesson');
        const integCol = cols.find(c => c.key === 'integ');
        if (lessonCol) {
            if (isCtlop) {
                let extraCount = (showSign ? 1 : 0) + (showNote ? 1 : 0) + customCount;
                if (extraCount >= 4) {
                    lessonCol.width = 24;
                    if (integCol) integCol.width = 22;
                } else if (extraCount >= 2) {
                    lessonCol.width = 28;
                    if (integCol) integCol.width = 25;
                } else if (extraCount === 1) {
                    lessonCol.width = 30;
                    if (integCol) integCol.width = 28;
                } else {
                    lessonCol.width = 34;
                    if (integCol) integCol.width = 30;
                }
            } else {
                let extraCount = (showSign ? 1 : 0) + (showNote ? 1 : 0) + customCount;
                if (extraCount >= 4) lessonCol.width = 28;
                else if (extraCount === 3) lessonCol.width = 30;
                else if (extraCount === 2) lessonCol.width = 34;
                else if (extraCount === 1) lessonCol.width = 40;
                else lessonCol.width = 50;
            }
        }

        cols.forEach((col, idx) => {
            col.colIdx = idx + 1;
            col.colLetter = getColLetter(col.colIdx);
        });

        return cols;
    }

    /**
     * Generate standard LBG / CTLOP Excel document
     */
    function generateLbgXlsx(options) {
        const { settings, weekNum, weekInfo, schedule, isCtlop } = options;
        const zip = new JSZip();

        const sheetName = `Tuan_${weekNum}`;
        const mainTitle = isCtlop 
            ? `LỊCH BÁO GIẢNG TÍCH HỢP TUẦN ${weekNum}` 
            : `LỊCH BÁO GIẢNG TUẦN ${weekNum}`;
        const subtitle = `(Thời gian thực hiện: Từ ngày ${weekInfo.startDateVN} đến ngày ${weekInfo.endDateVN})`;

        zip.file("[Content_Types].xml", createContentTypes());
        zip.file("_rels/.rels", createRels());
        zip.file("docProps/app.xml", createAppXml());
        zip.file("docProps/core.xml", createCoreXml(mainTitle));
        zip.file("xl/_rels/workbook.xml.rels", createWbRels());
        zip.file("xl/workbook.xml", createWbXml(sheetName));
        zip.file("xl/styles.xml", createStylesXml());

        const showColSign = !!(options.showColSign);
        const showColNote = !!(options.showColNote);
        let customCols = [];
        if (Array.isArray(options.customCols) && options.customCols.length > 0) {
            customCols = options.customCols;
        } else if (options.showColCustom) {
            customCols = [{
                id: 'col_1',
                name: (options.colCustomName && options.colCustomName.trim()) ? options.colCustomName.trim() : 'Ghi chú',
                pos: options.colCustomPos || "end",
                enabled: true
            }];
        }
        const showBghSign = (options.showBghSign !== false);
        const showGvcnSign = (options.showGvcnSign !== false);
        const showHeadSign = (options.showHeadSign !== false);
        const isLandscape = (options.orientation === 'landscape' || isCtlop);

        const orderedCols = getXlsxLbgOrderedColumns(customCols, showColSign, showColNote, isCtlop);
        const totalCols = orderedCols.length;
        const lastColLetter = getColLetter(totalCols);

        function makeCellsRow(startColIdx, endColIdx, rowNum, styleId) {
            let res = "";
            for (let c = startColIdx; c <= endColIdx; c++) {
                res += `<c r="${getColLetter(c)}${rowNum}" s="${styleId}"/>`;
            }
            return res;
        }

        let colsXml = "";
        orderedCols.forEach(col => {
            colsXml += `        <col min="${col.colIdx}" max="${col.colIdx}" width="${col.width}" customWidth="1"/>\n`;
        });

        let sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
        <pageSetUpPr fitToPage="1"/>
    </sheetPr>
    <cols>
${colsXml}    </cols>
    <sheetData>
`;

        const mergeCellsList = [];

        const splitCol = Math.max(3, Math.min(4, Math.floor(totalCols / 2)));
        const splitColLetter = getColLetter(splitCol);
        const nextColLetter = getColLetter(splitCol + 1);

        // Row 1: Header UBND & Quoc Hieu
        sheetXml += `
        <row r="1" ht="22" customHeight="1">
            <c r="A1" s="1" t="inlineStr"><is><t>${escapeXml(settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT')}</t></is></c>
            ${makeCellsRow(2, splitCol, 1, 1)}
            <c r="${nextColLetter}1" s="4" t="inlineStr"><is><t>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</t></is></c>
            ${makeCellsRow(splitCol + 2, totalCols, 1, 4)}
        </row>`;
        mergeCellsList.push(`A1:${splitColLetter}1`);
        mergeCellsList.push(`${nextColLetter}1:${lastColLetter}1`);

        // Row 2: School & Tieu Ngu
        sheetXml += `
        <row r="2" ht="22" customHeight="1">
            <c r="A2" s="2" t="inlineStr"><is><t>${escapeXml(settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT')}</t></is></c>
            ${makeCellsRow(2, splitCol, 2, 2)}
            <c r="${nextColLetter}2" s="5" t="inlineStr"><is><t>Độc lập - Tự do - Hạnh phúc</t></is></c>
            ${makeCellsRow(splitCol + 2, totalCols, 2, 5)}
        </row>`;
        mergeCellsList.push(`A2:${splitColLetter}2`);
        mergeCellsList.push(`${nextColLetter}2:${lastColLetter}2`);

        // Row 3: Grade & Class
        sheetXml += `
        <row r="3" ht="20" customHeight="1">
            <c r="A3" s="3" t="inlineStr"><is><t>${escapeXml(settings.grade || 'KHỐI 5')} - ${escapeXml(settings.className || 'LỚP 5A')}</t></is></c>
            ${makeCellsRow(2, splitCol, 3, 3)}
            ${makeCellsRow(splitCol + 1, totalCols, 3, 0)}
        </row>`;
        mergeCellsList.push(`A3:${splitColLetter}3`);

        // Row 4: Space
        sheetXml += `<row r="4" ht="10" customHeight="1"/>`;

        // Row 5: Title
        sheetXml += `
        <row r="5" ht="26" customHeight="1">
            <c r="A5" s="6" t="inlineStr"><is><t>${escapeXml(mainTitle)}</t></is></c>
            ${makeCellsRow(2, totalCols, 5, 6)}
        </row>`;
        mergeCellsList.push(`A5:${lastColLetter}5`);

        // Row 6: Subtitle
        sheetXml += `
        <row r="6" ht="18" customHeight="1">
            <c r="A6" s="7" t="inlineStr"><is><t>${escapeXml(subtitle)}</t></is></c>
            ${makeCellsRow(2, totalCols, 6, 7)}
        </row>`;
        mergeCellsList.push(`A6:${lastColLetter}6`);

        // Row 7: Space
        sheetXml += `<row r="7" ht="10" customHeight="1"/>`;

        // Row 8: Table Header
        let headerColsXml = "";
        orderedCols.forEach(col => {
            headerColsXml += `\n            <c r="${col.colLetter}8" s="8" t="inlineStr"><is><t>${escapeXml(col.title)}</t></is></c>`;
        });

        sheetXml += `
        <row r="8" ht="26" customHeight="1">${headerColsXml}
        </row>`;

        // Table Data Rows
        let currentRow = 9;
        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];

        const dayCol = orderedCols.find(c => c.key === 'day') || { colLetter: 'A' };
        const sessCol = orderedCols.find(c => c.key === 'session') || { colLetter: 'B' };

        days.forEach(day => {
            let daySlots = schedule.filter(s => s.day === day);
            if (options && options.hideEmptyRows) {
                daySlots = daySlots.filter(s => !s.isOff && s.subject !== '-- Nghỉ / Để trống --');
            }
            if (daySlots.length === 0) return;

            const dayDate = getDayFullDate(weekInfo.startDateVN, day);
            const dayStartRow = currentRow;
            const morningSlots = daySlots.filter(s => s.session === "Sáng");
            const afternoonSlots = daySlots.filter(s => s.session === "Chiều");

            const morningStartRow = currentRow;
            const morningEndRow = morningSlots.length > 0 ? (morningStartRow + morningSlots.length - 1) : morningStartRow;

            const afternoonStartRow = morningSlots.length > 0 ? (morningEndRow + 1) : currentRow;
            const afternoonEndRow = afternoonSlots.length > 0 ? (afternoonStartRow + afternoonSlots.length - 1) : afternoonStartRow;

            daySlots.forEach((slot, idx) => {
                const r = currentRow;
                const isFirstOfDay = (idx === 0);
                const isFirstOfMorning = (morningSlots.length > 0 && idx === 0);
                const isFirstOfAfternoon = (afternoonSlots.length > 0 && idx === morningSlots.length);

                const dayVal = isFirstOfDay ? (dayDate ? `${day}\n${dayDate}` : day) : "";
                const sessionVal = isFirstOfMorning ? "Sáng" : (isFirstOfAfternoon ? "Chiều" : "");

                // Calculate comfortable row height based on text length
                const lessonLen = (slot.lessonName || '').length;
                const integLen = (slot.integration || '').length;
                const lessonLines = Math.ceil(lessonLen / (isCtlop ? 30 : 44));
                const integLines = isCtlop ? Math.ceil(integLen / 26) : 1;
                let maxCustomLines = 1;
                orderedCols.filter(c => c.isCustom).forEach(c => {
                    let cVal = "";
                    if (slot.customCols && slot.customCols[c.id] !== undefined) cVal = slot.customCols[c.id];
                    else if (slot[`customCol_${c.id}`] !== undefined) cVal = slot[`customCol_${c.id}`];
                    else if (c.id === 'col_1') cVal = slot.customCol !== undefined ? slot.customCol : (slot.note || '');
                    const lines = Math.ceil((cVal || '').length / 22);
                    if (lines > maxCustomLines) maxCustomLines = lines;
                });
                const numLines = Math.max(isFirstOfDay && dayDate ? 2 : 1, lessonLines, integLines, maxCustomLines);

                let rowHt = 24;
                if (numLines === 2) rowHt = 38;
                else if (numLines === 3) rowHt = 54;
                else if (numLines >= 4) rowHt = numLines * 18;

                let rowCellsXml = "";
                orderedCols.forEach(col => {
                    if (col.key === 'day') {
                        rowCellsXml += `\n            <c r="${col.colLetter}${r}" s="9" ${dayVal ? 't="inlineStr"' : ''}>${dayVal ? `<is><t xml:space="preserve">${escapeXml(dayVal)}</t></is>` : ''}</c>`;
                    } else if (col.key === 'session') {
                        rowCellsXml += `\n            <c r="${col.colLetter}${r}" s="10" ${sessionVal ? 't="inlineStr"' : ''}>${sessionVal ? `<is><t>${escapeXml(sessionVal)}</t></is>` : ''}</c>`;
                    } else if (col.key === 'period') {
                        rowCellsXml += `\n            <c r="${col.colLetter}${r}" s="11" t="inlineStr"><is><t>${escapeXml(slot.period || '')}</t></is></c>`;
                    } else if (col.key === 'subject') {
                        const subjVal = slot.isOff ? '-- Nghỉ / Để trống --' : (slot.subject || '');
                        rowCellsXml += `\n            <c r="${col.colLetter}${r}" s="12" t="inlineStr"><is><t>${escapeXml(subjVal)}</t></is></c>`;
                    } else if (col.key === 'ppct') {
                        rowCellsXml += `\n            <c r="${col.colLetter}${r}" s="13" t="inlineStr"><is><t>${escapeXml(slot.ppct || '')}</t></is></c>`;
                    } else if (col.key === 'lesson') {
                        rowCellsXml += `\n            <c r="${col.colLetter}${r}" s="14" t="inlineStr"><is><t>${escapeXml(slot.lessonName || '')}</t></is></c>`;
                    } else if (col.key === 'integ') {
                        rowCellsXml += `\n            <c r="${col.colLetter}${r}" s="15" t="inlineStr"><is><t>${escapeXml(slot.integration || '')}</t></is></c>`;
                    } else if (col.key === 'sign') {
                        rowCellsXml += `\n            <c r="${col.colLetter}${r}" s="10"/>`;
                    } else if (col.key === 'note') {
                        const noteText = slot.note || '';
                        rowCellsXml += `\n            <c r="${col.colLetter}${r}" s="14" ${noteText ? 't="inlineStr"' : ''}>${noteText ? `<is><t>${escapeXml(noteText)}</t></is>` : ''}</c>`;
                    } else if (col.isCustom) {
                        let cVal = "";
                        if (slot.customCols && slot.customCols[col.id] !== undefined) {
                            cVal = slot.customCols[col.id];
                        } else if (slot[`customCol_${col.id}`] !== undefined) {
                            cVal = slot[`customCol_${col.id}`];
                        } else if (col.id === 'col_1') {
                            cVal = slot.customCol !== undefined ? slot.customCol : (slot.note || '');
                        }
                        rowCellsXml += `\n            <c r="${col.colLetter}${r}" s="14" ${cVal ? 't="inlineStr"' : ''}>${cVal ? `<is><t>${escapeXml(cVal)}</t></is>` : ''}</c>`;
                    }
                });

                sheetXml += `
        <row r="${r}" ht="${rowHt}" customHeight="1">${rowCellsXml}
        </row>`;
                currentRow++;
            });

            const dayEndRow = currentRow - 1;
            if (dayEndRow > dayStartRow) {
                mergeCellsList.push(`${dayCol.colLetter}${dayStartRow}:${dayCol.colLetter}${dayEndRow}`);
            }
            if (morningSlots.length > 1) {
                mergeCellsList.push(`${sessCol.colLetter}${morningStartRow}:${sessCol.colLetter}${morningEndRow}`);
            }
            if (afternoonSlots.length > 1) {
                mergeCellsList.push(`${sessCol.colLetter}${afternoonStartRow}:${sessCol.colLetter}${afternoonEndRow}`);
            }
        });

        // Hàng tổng kết cuối bảng: Tổng số tiết/tuần
        const totalPeriodsXlsx = (options.stats && options.stats.total !== undefined) 
            ? options.stats.total 
            : schedule.filter(s => !s.isOff && s.subject !== '-- Nghỉ / Để trống --').length;
        const totalRowHt = 26;
        let totalRowXml = "";
        totalRowXml += `\n            <c r="A${currentRow}" s="19" t="inlineStr"><is><t>Tổng số tiết/tuần</t></is></c>`;
        totalRowXml += `\n            <c r="B${currentRow}" s="19"/>`;
        totalRowXml += `\n            <c r="C${currentRow}" s="19"/>`;
        totalRowXml += `\n            <c r="D${currentRow}" s="19" t="inlineStr"><is><t>${totalPeriodsXlsx} tiết/ tuần</t></is></c>`;
        for (let c = 5; c <= totalCols; c++) {
            totalRowXml += `\n            <c r="${getColLetter(c)}${currentRow}" s="19"/>`;
        }
        sheetXml += `\n        <row r="${currentRow}" ht="${totalRowHt}" customHeight="1">${totalRowXml}\n        </row>`;
        mergeCellsList.push(`A${currentRow}:C${currentRow}`);
        mergeCellsList.push(`D${currentRow}:${lastColLetter}${currentRow}`);
        currentRow++;

        // Space before footer
        sheetXml += `<row r="${currentRow}" ht="14" customHeight="1"/>`;
        currentRow++;

        // Footer Signatures
        const bghLbgType = settings.bghSignerLbgType || settings.bghSignerType || 'PHT';
        const bghLbgIdx = (settings.bghSignerLbgIndex !== undefined) ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0);

        const bghSignerName = (bghLbgType === 'HT') 
            ? (settings.principal || 'Phạm Quốc Hùng') 
            : (Array.isArray(settings.vicePrincipals) && settings.vicePrincipals[bghLbgIdx] 
                ? settings.vicePrincipals[bghLbgIdx] 
                : (settings.vicePrincipal || 'Lê Văn Tám'));
        const bghSignerTitle = bghLbgType === 'HT' ? 'HIỆU TRƯỞNG' : 'BAN GIÁM HIỆU';

        const signers = [];
        if (showBghSign) {
            signers.push({
                role: `DUYỆT CỦA ${bghSignerTitle}`,
                sub: `(Ký và ghi rõ họ tên)`,
                name: bghSignerName
            });
        }
        if (showHeadSign) {
            signers.push({
                role: `KHỐI TRƯỞNG`,
                sub: `(Ký và ghi rõ họ tên)`,
                name: settings.headOfGrade || 'Trần Thị Mai'
            });
        }
        if (showGvcnSign) {
            signers.push({
                role: `GIÁO VIÊN CHỦ NHIỆM`,
                sub: `(Ký và ghi rõ họ tên)`,
                name: settings.homeroomTeacher || 'Nguyễn Thị Thu Hà'
            });
        }

        if (signers.length > 0) {
            const footerTitleRow = currentRow;
            const footerSubRow = currentRow + 1;
            const footerSpaceRow1 = currentRow + 2;
            const footerSpaceRow2 = currentRow + 3;
            const footerSpaceRow3 = currentRow + 4;
            const footerNameRow = currentRow + 5;

            const numSigners = signers.length;
            let currentCol = 1;
            let titleCells = "";
            let subCells = "";
            let nameCells = "";

            for (let i = 0; i < numSigners; i++) {
                const s = signers[i];
                const remainingCols = totalCols - currentCol + 1;
                const remainingSigners = numSigners - i;
                const span = Math.round(remainingCols / remainingSigners);
                const startCol = currentCol;
                const endCol = (i === numSigners - 1) ? totalCols : (startCol + span - 1);
                const startColLetter = getColLetter(startCol);
                const endColLetter = getColLetter(endCol);

                titleCells += `\n            <c r="${startColLetter}${footerTitleRow}" s="16" t="inlineStr"><is><t>${escapeXml(s.role)}</t></is></c>`;
                for (let c = startCol + 1; c <= endCol; c++) {
                    titleCells += `\n            <c r="${getColLetter(c)}${footerTitleRow}" s="16"/>`;
                }

                subCells += `\n            <c r="${startColLetter}${footerSubRow}" s="17" t="inlineStr"><is><t>${escapeXml(s.sub)}</t></is></c>`;
                for (let c = startCol + 1; c <= endCol; c++) {
                    subCells += `\n            <c r="${getColLetter(c)}${footerSubRow}" s="17"/>`;
                }

                nameCells += `\n            <c r="${startColLetter}${footerNameRow}" s="18" t="inlineStr"><is><t>${escapeXml(s.name)}</t></is></c>`;
                for (let c = startCol + 1; c <= endCol; c++) {
                    nameCells += `\n            <c r="${getColLetter(c)}${footerNameRow}" s="18"/>`;
                }

                if (endCol > startCol) {
                    mergeCellsList.push(`${startColLetter}${footerTitleRow}:${endColLetter}${footerTitleRow}`);
                    mergeCellsList.push(`${startColLetter}${footerSubRow}:${endColLetter}${footerSubRow}`);
                    mergeCellsList.push(`${startColLetter}${footerNameRow}:${endColLetter}${footerNameRow}`);
                }

                currentCol = endCol + 1;
            }

            sheetXml += `
        <row r="${footerTitleRow}" ht="20" customHeight="1">${titleCells}
        </row>
        <row r="${footerSubRow}" ht="18" customHeight="1">${subCells}
        </row>
        <row r="${footerSpaceRow1}" ht="18" customHeight="1"/>
        <row r="${footerSpaceRow2}" ht="18" customHeight="1"/>
        <row r="${footerSpaceRow3}" ht="18" customHeight="1"/>
        <row r="${footerNameRow}" ht="22" customHeight="1">${nameCells}
        </row>`;
        }

        sheetXml += `
    </sheetData>
    <mergeCells count="${mergeCellsList.length}">
        ${mergeCellsList.map(m => `<mergeCell ref="${m}"/>`).join('\n        ')}
    </mergeCells>
    <pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>
    <pageSetup orientation="${isLandscape ? 'landscape' : 'portrait'}" paperSize="9" fitToWidth="1" fitToHeight="0"/>
</worksheet>`;

        zip.file("xl/worksheets/sheet1.xml", sheetXml);

        return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    }

    function generateTimetableXlsx(slots, settings) {
        const zip = new JSZip();

        zip.file("[Content_Types].xml", createContentTypes());
        zip.file("_rels/.rels", createRels());
        zip.file("docProps/app.xml", createAppXml());
        zip.file("docProps/core.xml", createCoreXml("Thoi_Khoa_Bieu"));
        zip.file("xl/_rels/workbook.xml.rels", createWbRels());
        zip.file("xl/workbook.xml", createWbXml("Thoi_Khoa_Bieu"));
        zip.file("xl/styles.xml", createStylesXml());

        const bghSignerName = ((settings.bghSignerLbgType || settings.bghSignerType || 'PHT') === 'HT')
            ? (settings.principal || 'Phạm Quốc Hùng')
            : (Array.isArray(settings.vicePrincipals) && settings.vicePrincipals[(settings.bghSignerLbgIndex !== undefined ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0))]
                ? settings.vicePrincipals[(settings.bghSignerLbgIndex !== undefined ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0))]
                : (settings.vicePrincipal || 'Lê Văn Tám'));
        const bghSignerRole = ((settings.bghSignerLbgType || settings.bghSignerType || 'PHT') === 'HT') ? 'HIỆU TRƯỞNG' : 'BAN GIÁM HIỆU';

        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];
        const dayCols = ["C", "D", "E", "F", "G"];

        let maxMorn = 4;
        let maxAft = 3;
        slots.forEach(s => {
            if (s.session === "Sáng" && s.period > maxMorn) maxMorn = s.period;
            if (s.session === "Chiều" && s.period > maxAft) maxAft = s.period;
        });

        const mergeCellsList = [];
        let currentRow = 1;

        let sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <cols>
        <col min="1" max="1" width="12" customWidth="1"/>
        <col min="2" max="2" width="8" customWidth="1"/>
        <col min="3" max="7" width="22" customWidth="1"/>
    </cols>
    <sheetData>
        <row r="${currentRow}">
            <c r="A${currentRow}" s="1" t="s"><v>${escapeXml((settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT').toUpperCase())}</v></c>
            <c r="E${currentRow}" s="2" t="s"><v>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</v></c>
        </row>`;
        mergeCellsList.push(`A${currentRow}:C${currentRow}`);
        mergeCellsList.push(`E${currentRow}:G${currentRow}`);
        currentRow++;

        sheetXml += `
        <row r="${currentRow}">
            <c r="A${currentRow}" s="2" t="s"><v>${escapeXml((settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT').toUpperCase())}</v></c>
            <c r="E${currentRow}" s="2" t="s"><v>Độc lập - Tự do - Hạnh phúc</v></c>
        </row>`;
        mergeCellsList.push(`A${currentRow}:C${currentRow}`);
        mergeCellsList.push(`E${currentRow}:G${currentRow}`);
        currentRow++;

        sheetXml += `
        <row r="${currentRow}">
            <c r="A${currentRow}" s="2" t="s"><v>${escapeXml((settings.grade || 'KHỐI 5').toUpperCase())} - ${escapeXml((settings.className || 'LỚP 5A').toUpperCase())}</v></c>
        </row>`;
        mergeCellsList.push(`A${currentRow}:C${currentRow}`);
        currentRow += 2;

        sheetXml += `
        <row r="${currentRow}">
            <c r="A${currentRow}" s="3" t="s"><v>${escapeXml(('THỜI KHÓA BIỂU ' + (settings.grade || 'KHỐI 5') + ' - ' + (settings.className || 'LỚP 5A')).toUpperCase())}</v></c>
        </row>`;
        mergeCellsList.push(`A${currentRow}:G${currentRow}`);
        currentRow++;

        sheetXml += `
        <row r="${currentRow}">
            <c r="A${currentRow}" s="4" t="s"><v>${escapeXml('Năm học ' + (settings.academicYear || '2026 - 2027'))}</v></c>
        </row>`;
        mergeCellsList.push(`A${currentRow}:G${currentRow}`);
        currentRow += 2;

        // Header table
        sheetXml += `
        <row r="${currentRow}">
            <c r="A${currentRow}" s="5" t="s"><v>Buổi</v></c>
            <c r="B${currentRow}" s="5" t="s"><v>Tiết</v></c>
            <c r="C${currentRow}" s="5" t="s"><v>Thứ 2</v></c>
            <c r="D${currentRow}" s="5" t="s"><v>Thứ 3</v></c>
            <c r="E${currentRow}" s="5" t="s"><v>Thứ 4</v></c>
            <c r="F${currentRow}" s="5" t="s"><v>Thứ 5</v></c>
            <c r="G${currentRow}" s="5" t="s"><v>Thứ 6</v></c>
        </row>`;
        currentRow++;

        // Morning
        const mornStartRow = currentRow;
        for (let p = 1; p <= maxMorn; p++) {
            sheetXml += `
            <row r="${currentRow}">
                <c r="A${currentRow}" s="6" t="s"><v>${p === 1 ? 'Sáng' : ''}</v></c>
                <c r="B${currentRow}" s="6" t="s"><v>${p}</v></c>`;
            days.forEach((d, idx) => {
                const dayMornSlots = slots.filter(s => s.day === d && s.session === "Sáng").sort((a, b) => (a.period || 0) - (b.period || 0));
                const slot = dayMornSlots[p - 1];
                const sub = slot ? slot.subject : "";
                sheetXml += `<c r="${dayCols[idx]}${currentRow}" s="6" t="s"><v>${escapeXml(sub)}</v></c>`;
            });
            sheetXml += `</row>`;
            currentRow++;
        }
        if (maxMorn > 1) {
            mergeCellsList.push(`A${mornStartRow}:A${currentRow - 1}`);
        }

        // Afternoon
        const aftStartRow = currentRow;
        for (let p = 1; p <= maxAft; p++) {
            sheetXml += `
            <row r="${currentRow}">
                <c r="A${currentRow}" s="6" t="s"><v>${p === 1 ? 'Chiều' : ''}</v></c>
                <c r="B${currentRow}" s="6" t="s"><v>${p}</v></c>`;
            days.forEach((d, idx) => {
                const dayAftSlots = slots.filter(s => s.day === d && s.session === "Chiều").sort((a, b) => (a.period || 0) - (b.period || 0));
                const slot = dayAftSlots[p - 1];
                const sub = slot ? slot.subject : "";
                sheetXml += `<c r="${dayCols[idx]}${currentRow}" s="6" t="s"><v>${escapeXml(sub)}</v></c>`;
            });
            sheetXml += `</row>`;
            currentRow++;
        }
        if (maxAft > 1) {
            mergeCellsList.push(`A${aftStartRow}:A${currentRow - 1}`);
        }

        currentRow += 2;
        const footerTitleRow = currentRow;
        sheetXml += `
        <row r="${footerTitleRow}">
            <c r="A${footerTitleRow}" s="2" t="s"><v>DUYỆT CỦA ${escapeXml(bghSignerRole)}</v></c>
            <c r="D${footerTitleRow}" s="2" t="s"><v>TỔ TRƯỞNG CHUYÊN MÔN</v></c>
            <c r="F${footerTitleRow}" s="2" t="s"><v>GIÁO VIÊN CHỦ NHIỆM</v></c>
        </row>`;
        mergeCellsList.push(`A${footerTitleRow}:C${footerTitleRow}`);
        mergeCellsList.push(`D${footerTitleRow}:E${footerTitleRow}`);
        mergeCellsList.push(`F${footerTitleRow}:G${footerTitleRow}`);

        currentRow++;
        const footerSubRow = currentRow;
        sheetXml += `
        <row r="${footerSubRow}">
            <c r="A${footerSubRow}" s="4" t="s"><v>(Ký và ghi rõ họ tên)</v></c>
            <c r="D${footerSubRow}" s="4" t="s"><v>(Ký và ghi rõ họ tên)</v></c>
            <c r="F${footerSubRow}" s="4" t="s"><v>(Ký và ghi rõ họ tên)</v></c>
        </row>`;
        mergeCellsList.push(`A${footerSubRow}:C${footerSubRow}`);
        mergeCellsList.push(`D${footerSubRow}:E${footerSubRow}`);
        mergeCellsList.push(`F${footerSubRow}:G${footerSubRow}`);

        currentRow += 4;
        const footerNameRow = currentRow;
        sheetXml += `
        <row r="${footerNameRow}">
            <c r="A${footerNameRow}" s="2" t="s"><v>${escapeXml(bghSignerName)}</v></c>
            <c r="D${footerNameRow}" s="2" t="s"><v>${escapeXml(settings.headOfGrade || 'Trần Thị Mai')}</v></c>
            <c r="F${footerNameRow}" s="2" t="s"><v>${escapeXml(settings.homeroomTeacher || 'Nguyễn Thị Thu Hà')}</v></c>
        </row>`;
        mergeCellsList.push(`A${footerNameRow}:C${footerNameRow}`);
        mergeCellsList.push(`D${footerNameRow}:E${footerNameRow}`);
        mergeCellsList.push(`F${footerNameRow}:G${footerNameRow}`);

        sheetXml += `
    </sheetData>
    <mergeCells count="${mergeCellsList.length}">
        ${mergeCellsList.map(m => `<mergeCell ref="${m}"/>`).join('\n        ')}
    </mergeCells>
    <pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>
    <pageSetup orientation="landscape" paperSize="9" fitToWidth="1" fitToHeight="0"/>
</worksheet>`;

        zip.file("xl/worksheets/sheet1.xml", sheetXml);

        return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    }


    function getDayDateStrBySubjectXlsx(startDateVN, dayStr) {
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

    function generateLbgBySubjectXlsx(data, settings, isCtlop = true, filterSubject = "all") {
        const zip = new JSZip();

        zip.file("[Content_Types].xml", createContentTypes());
        zip.file("_rels/.rels", createRels());
        zip.file("docProps/app.xml", createAppXml());

        const weekNum = data.weekNum || 1;
        const weekInfo = data.weekInfo || {};
        const subjectGroups = data.subjectGroups || [];

        let mainTitle = `LỊCH BÁO GIẢNG THEO MÔN HỌC TUẦN ${weekNum}`;
        if (filterSubject && filterSubject !== "all") {
            mainTitle = `LỊCH BÁO GIẢNG MÔN ${filterSubject.toUpperCase()} TUẦN ${weekNum}`;
        }

        zip.file("docProps/core.xml", createCoreXml(mainTitle));
        zip.file("xl/_rels/workbook.xml.rels", createWbRels());
        zip.file("xl/workbook.xml", createWbXml(`Tuan ${weekNum} Theo Mon`));
        zip.file("xl/styles.xml", createStylesXml());

        const subtitle = `(Thời gian thực hiện: Từ ngày ${weekInfo.startDateVN || ''} đến ngày ${weekInfo.endDateVN || ''})`;
        const lastColLetter = isCtlop ? "H" : "G";
        const mergeCellsList = [];

        let sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <dimension ref="A1:${lastColLetter}60"/>
    <sheetViews>
        <sheetView tabSelected="1" workbookViewId="0">
            <pane ySplit="8" topLeftCell="A9" activePane="bottomLeft" state="frozen"/>
        </sheetView>
    </sheetViews>
    <sheetFormatPr defaultRowHeight="20"/>
    <cols>
        <col min="1" max="1" width="${isCtlop ? 20 : 22}" customWidth="1"/>
        <col min="2" max="2" width="16" customWidth="1"/>
        <col min="3" max="3" width="9" customWidth="1"/>
        <col min="4" max="4" width="7" customWidth="1"/>
        <col min="5" max="5" width="10" customWidth="1"/>
        <col min="6" max="6" width="11" customWidth="1"/>
        <col min="7" max="7" width="${isCtlop ? 36 : 48}" customWidth="1"/>
        ${isCtlop ? '<col min="8" max="8" width="36" customWidth="1"/>' : ''}
    </cols>
    <sheetData>
        <!-- Row 1: Header UBND & Quoc Hieu -->
        <row r="1" ht="22" customHeight="1">
            <c r="A1" s="1" t="inlineStr"><is><t>${escapeXml(settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT')}</t></is></c>
            <c r="B1" s="1"/><c r="C1" s="1"/>
            <c r="D1" s="4" t="inlineStr"><is><t>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</t></is></c>
            <c r="E1" s="4"/><c r="F1" s="4"/><c r="G1" s="4"/>${isCtlop ? '<c r="H1" s="4"/>' : ''}
        </row>
        <!-- Row 2: School & Tieu Ngu -->
        <row r="2" ht="22" customHeight="1">
            <c r="A2" s="2" t="inlineStr"><is><t>${escapeXml(settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT')}</t></is></c>
            <c r="B2" s="2"/><c r="C2" s="2"/>
            <c r="D2" s="5" t="inlineStr"><is><t>Độc lập - Tự do - Hạnh phúc</t></is></c>
            <c r="E2" s="5"/><c r="F2" s="5"/><c r="G2" s="5"/>${isCtlop ? '<c r="H2" s="5"/>' : ''}
        </row>
        <!-- Row 3: Grade & Class -->
        <row r="3" ht="20" customHeight="1">
            <c r="A3" s="3" t="inlineStr"><is><t>${escapeXml(settings.grade || 'KHỐI 5')} - ${escapeXml(settings.className || 'LỚP 5A')}</t></is></c>
            <c r="B3" s="3"/><c r="C3" s="3"/>
            <c r="D3" s="0"/><c r="E3" s="0"/><c r="F3" s="0"/><c r="G3" s="0"/>${isCtlop ? '<c r="H3" s="0"/>' : ''}
        </row>
        <!-- Row 4: Space -->
        <row r="4" ht="10" customHeight="1"/>
        <!-- Row 5: Title -->
        <row r="5" ht="26" customHeight="1">
            <c r="A5" s="6" t="inlineStr"><is><t>${escapeXml(mainTitle)}</t></is></c>
            <c r="B5" s="6"/><c r="C5" s="6"/><c r="D5" s="6"/><c r="E5" s="6"/><c r="F5" s="6"/><c r="G5" s="6"/>${isCtlop ? '<c r="H5" s="6"/>' : ''}
        </row>
        <!-- Row 6: Subtitle -->
        <row r="6" ht="18" customHeight="1">
            <c r="A6" s="7" t="inlineStr"><is><t>${escapeXml(subtitle)}</t></is></c>
            <c r="B6" s="7"/><c r="C6" s="7"/><c r="D6" s="7"/><c r="E6" s="7"/><c r="F6" s="7"/><c r="G6" s="7"/>${isCtlop ? '<c r="H6" s="7"/>' : ''}
        </row>
        <!-- Row 7: Space -->
        <row r="7" ht="10" customHeight="1"/>
        <!-- Row 8: Table Header -->
        <row r="8" ht="26" customHeight="1">
            <c r="A8" s="8" t="inlineStr"><is><t>Môn học</t></is></c>
            <c r="B8" s="8" t="inlineStr"><is><t>Thứ / Ngày</t></is></c>
            <c r="C8" s="8" t="inlineStr"><is><t>Buổi</t></is></c>
            <c r="D8" s="8" t="inlineStr"><is><t>Tiết</t></is></c>
            <c r="E8" s="8" t="inlineStr"><is><t>Tiết/tuần</t></is></c>
            <c r="F8" s="8" t="inlineStr"><is><t>Tiết PPCT</t></is></c>
            <c r="G8" s="8" t="inlineStr"><is><t>Tên bài dạy</t></is></c>
            ${isCtlop ? '<c r="H8" s="8" t="inlineStr"><is><t>Nội dung tích hợp / Điều chỉnh</t></is></c>' : ''}
        </row>`;

        mergeCellsList.push("A1:C1");
        mergeCellsList.push(`D1:${lastColLetter}1`);
        mergeCellsList.push("A2:C2");
        mergeCellsList.push(`D2:${lastColLetter}2`);
        mergeCellsList.push("A3:C3");
        mergeCellsList.push(`A5:${lastColLetter}5`);
        mergeCellsList.push(`A6:${lastColLetter}6`);

        let currentRow = 9;

        if (subjectGroups.length === 0) {
            sheetXml += `
        <row r="${currentRow}" ht="26" customHeight="1">
            <c r="A${currentRow}" s="12" t="inlineStr"><is><t>Không có tiết học nào phù hợp với bộ lọc đã chọn.</t></is></c>
            <c r="B${currentRow}" s="12"/><c r="C${currentRow}" s="12"/><c r="D${currentRow}" s="12"/><c r="E${currentRow}" s="12"/><c r="F${currentRow}" s="12"/><c r="G${currentRow}" s="12"/>${isCtlop ? `<c r="H${currentRow}" s="12"/>` : ''}
        </row>`;
            mergeCellsList.push(`A${currentRow}:${lastColLetter}${currentRow}`);
            currentRow++;
        } else {
            subjectGroups.forEach(group => {
                const groupStartRow = currentRow;
                const groupCount = group.slots.length;

                group.slots.forEach((slot, idx) => {
                    const r = currentRow;
                    const isFirstOfGroup = (idx === 0);
                    const subVal = isFirstOfGroup ? `${group.subjectName} (${groupCount} tiết)` : "";
                    const dayDateText = getDayDateStrBySubjectXlsx(weekInfo.startDateVN, slot.day);

                    const lessonLen = (slot.lessonName || '').length;
                    const integLen = (slot.integration || '').length;
                    const lessonLines = Math.ceil(lessonLen / (isCtlop ? 30 : 44));
                    const integLines = isCtlop ? Math.ceil(integLen / 26) : 1;
                    const numLines = Math.max(1, lessonLines, integLines);

                    let rowHt = 24;
                    if (numLines === 2) rowHt = 38;
                    else if (numLines === 3) rowHt = 54;
                    else if (numLines >= 4) rowHt = numLines * 18;

                    sheetXml += `
        <row r="${r}" ht="${rowHt}" customHeight="1">
            <c r="A${r}" s="9" ${subVal ? 't="inlineStr"' : ''}>${subVal ? `<is><t>${escapeXml(subVal)}</t></is>` : ''}</c>
            <c r="B${r}" s="10" t="inlineStr"><is><t>${escapeXml(dayDateText)}</t></is></c>
            <c r="C${r}" s="10" t="inlineStr"><is><t>${escapeXml(slot.session || '')}</t></is></c>
            <c r="D${r}" s="11" t="inlineStr"><is><t>${escapeXml(slot.period || '')}</t></is></c>
            <c r="E${r}" s="11" t="inlineStr"><is><t>${escapeXml(slot.periodInWeek || (idx + 1))}</t></is></c>
            <c r="F${r}" s="13" t="inlineStr"><is><t>${escapeXml(slot.ppct || '')}</t></is></c>
            <c r="G${r}" s="14" t="inlineStr"><is><t>${escapeXml(slot.lessonName || '')}</t></is></c>
            ${isCtlop ? `<c r="H${r}" s="15" t="inlineStr"><is><t>${escapeXml(slot.integration || '')}</t></is></c>` : ''}
        </row>`;
                    currentRow++;
                });

                if (groupCount > 1) {
                    mergeCellsList.push(`A${groupStartRow}:A${currentRow - 1}`);
                }
            });
        }

        // Space before footer
        sheetXml += `<row r="${currentRow}" ht="14" customHeight="1"/>`;
        currentRow++;

        // Footer Signatures
        const footerTitleRow = currentRow;
        const footerSubRow = currentRow + 1;
        const footerSpaceRow1 = currentRow + 2;
        const footerSpaceRow2 = currentRow + 3;
        const footerSpaceRow3 = currentRow + 4;
        const footerNameRow = currentRow + 5;

        const bghLbgType = settings.bghSignerLbgType || settings.bghSignerType || 'PHT';
        const bghLbgIdx = (settings.bghSignerLbgIndex !== undefined) ? settings.bghSignerLbgIndex : (settings.bghSignerIndex || 0);

        const bghSignerName = (bghLbgType === 'HT') 
            ? (settings.principal || 'Phạm Quốc Hùng') 
            : (Array.isArray(settings.vicePrincipals) && settings.vicePrincipals[bghLbgIdx] 
                ? settings.vicePrincipals[bghLbgIdx] 
                : (settings.vicePrincipal || 'Lê Văn Tám'));
        const bghSignerTitle = bghLbgType === 'HT' ? 'HIỆU TRƯỞNG' : 'BAN GIÁM HIỆU';

        sheetXml += `
        <row r="${footerTitleRow}" ht="20" customHeight="1">
            <c r="A${footerTitleRow}" s="16" t="inlineStr"><is><t>DUYỆT CỦA ${escapeXml(bghSignerTitle)}</t></is></c>
            <c r="B${footerTitleRow}" s="16"/>
            <c r="C${footerTitleRow}" s="16" t="inlineStr"><is><t>TỔ TRƯỞNG CHUYÊN MÔN</t></is></c>
            <c r="D${footerTitleRow}" s="16"/><c r="E${footerTitleRow}" s="16"/>
            <c r="F${footerTitleRow}" s="16" t="inlineStr"><is><t>GIÁO VIÊN CHỦ NHIỆM</t></is></c>
            <c r="G${footerTitleRow}" s="16"/>${isCtlop ? `<c r="H${footerTitleRow}" s="16"/>` : ''}
        </row>
        <row r="${footerSubRow}" ht="18" customHeight="1">
            <c r="A${footerSubRow}" s="17" t="inlineStr"><is><t>(Ký và ghi rõ họ tên)</t></is></c>
            <c r="B${footerSubRow}" s="17"/>
            <c r="C${footerSubRow}" s="17" t="inlineStr"><is><t>(Ký và ghi rõ họ tên)</t></is></c>
            <c r="D${footerSubRow}" s="17"/><c r="E${footerSubRow}" s="17"/>
            <c r="F${footerSubRow}" s="17" t="inlineStr"><is><t>(Ký và ghi rõ họ tên)</t></is></c>
            <c r="G${footerSubRow}" s="17"/>${isCtlop ? `<c r="H${footerSubRow}" s="17"/>` : ''}
        </row>
        <row r="${footerSpaceRow1}" ht="18" customHeight="1"/>
        <row r="${footerSpaceRow2}" ht="18" customHeight="1"/>
        <row r="${footerSpaceRow3}" ht="18" customHeight="1"/>
        <row r="${footerNameRow}" ht="22" customHeight="1">
            <c r="A${footerNameRow}" s="18" t="inlineStr"><is><t>${escapeXml(bghSignerName)}</t></is></c>
            <c r="B${footerNameRow}" s="18"/>
            <c r="C${footerNameRow}" s="18" t="inlineStr"><is><t>${escapeXml(settings.headOfGrade || 'Trần Thị Mai')}</t></is></c>
            <c r="D${footerNameRow}" s="18"/><c r="E${footerNameRow}" s="18"/>
            <c r="F${footerNameRow}" s="18" t="inlineStr"><is><t>${escapeXml(settings.homeroomTeacher || 'Nguyễn Thị Thu Hà')}</t></is></c>
            <c r="G${footerNameRow}" s="18"/>${isCtlop ? `<c r="H${footerNameRow}" s="18"/>` : ''}
        </row>`;

        mergeCellsList.push(`A${footerTitleRow}:B${footerTitleRow}`);
        mergeCellsList.push(`C${footerTitleRow}:E${footerTitleRow}`);
        mergeCellsList.push(`F${footerTitleRow}:${lastColLetter}${footerTitleRow}`);

        mergeCellsList.push(`A${footerSubRow}:B${footerSubRow}`);
        mergeCellsList.push(`C${footerSubRow}:E${footerSubRow}`);
        mergeCellsList.push(`F${footerSubRow}:${lastColLetter}${footerSubRow}`);

        mergeCellsList.push(`A${footerNameRow}:B${footerNameRow}`);
        mergeCellsList.push(`C${footerNameRow}:E${footerNameRow}`);
        mergeCellsList.push(`F${footerNameRow}:${lastColLetter}${footerNameRow}`);

        sheetXml += `
    </sheetData>
    <mergeCells count="${mergeCellsList.length}">
        ${mergeCellsList.map(m => `<mergeCell ref="${m}"/>`).join('\n        ')}
    </mergeCells>
    <pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>
    <pageSetup orientation="${isCtlop ? 'landscape' : 'portrait'}" paperSize="9" fitToWidth="1" fitToHeight="0"/>
</worksheet>`;

        zip.file("xl/worksheets/sheet1.xml", sheetXml);

        return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    }


    return {
        generateLbgXlsx: generateLbgXlsx,
        generateTimetableXlsx: generateTimetableXlsx,
        generateLbgBySubjectXlsx: generateLbgBySubjectXlsx
    };
})();
