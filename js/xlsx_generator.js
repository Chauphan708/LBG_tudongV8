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
        <!-- 0: Times New Roman 11pt Regular -->
        <font><name val="Times New Roman"/><sz val="11"/><color rgb="FF000000"/></font>
        <!-- 1: Times New Roman 11pt Bold -->
        <font><name val="Times New Roman"/><sz val="11"/><b/><color rgb="FF000000"/></font>
        <!-- 2: Times New Roman 12pt Bold Underline -->
        <font><name val="Times New Roman"/><sz val="12"/><b/><u/><color rgb="FF000000"/></font>
        <!-- 3: Times New Roman 14pt Bold -->
        <font><name val="Times New Roman"/><sz val="14"/><b/><color rgb="FF000000"/></font>
        <!-- 4: Times New Roman 11pt Italic -->
        <font><name val="Times New Roman"/><sz val="11"/><i/><color rgb="FF000000"/></font>
        <!-- 5: Times New Roman 10pt Regular -->
        <font><name val="Times New Roman"/><sz val="10"/><color rgb="FF000000"/></font>
        <!-- 6: Times New Roman 11pt Bold Underline -->
        <font><name val="Times New Roman"/><sz val="11"/><b/><u/><color rgb="FF000000"/></font>
        <!-- 7: Times New Roman 10pt Italic -->
        <font><name val="Times New Roman"/><sz val="10"/><i/><color rgb="FF000000"/></font>
        <!-- 8: Times New Roman 10pt Bold Underline (for School Header) -->
        <font><name val="Times New Roman"/><sz val="10.5"/><b/><u/><color rgb="FF000000"/></font>
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
        <!-- 6: Title Main (14pt Bold, Center) -->
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
        <!-- 12: Table Subject (Bold, Border, Left H, Center V, Wrap) -->
        <xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
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

        const lastColLetter = isCtlop ? "G" : "F";

        let sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
        <pageSetUpPr fitToPage="1"/>
    </sheetPr>
    <cols>
        <col min="1" max="1" width="13" customWidth="1"/>
        <col min="2" max="2" width="9" customWidth="1"/>
        <col min="3" max="3" width="7" customWidth="1"/>
        <col min="4" max="4" width="${isCtlop ? 18 : 20}" customWidth="1"/>
        <col min="5" max="5" width="10" customWidth="1"/>
        <col min="6" max="6" width="${isCtlop ? 34 : 50}" customWidth="1"/>
        ${isCtlop ? '<col min="7" max="7" width="30" customWidth="1"/>' : ''}
    </cols>
    <sheetData>
`;

        const mergeCellsList = [];

        // Row 1: Header UBND & Quoc Hieu
        sheetXml += `
        <row r="1" ht="22" customHeight="1">
            <c r="A1" s="1" t="inlineStr"><is><t>${escapeXml(settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT')}</t></is></c>
            <c r="B1" s="1"/><c r="C1" s="1"/>
            <c r="D1" s="4" t="inlineStr"><is><t>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</t></is></c>
            <c r="E1" s="4"/><c r="F1" s="4"/>${isCtlop ? '<c r="G1" s="4"/>' : ''}
        </row>`;
        mergeCellsList.push("A1:C1");
        mergeCellsList.push(`D1:${lastColLetter}1`);

        // Row 2: School & Tieu Ngu
        sheetXml += `
        <row r="2" ht="22" customHeight="1">
            <c r="A2" s="2" t="inlineStr"><is><t>${escapeXml(settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT')}</t></is></c>
            <c r="B2" s="2"/><c r="C2" s="2"/>
            <c r="D2" s="5" t="inlineStr"><is><t>Độc lập - Tự do - Hạnh phúc</t></is></c>
            <c r="E2" s="5"/><c r="F2" s="5"/>${isCtlop ? '<c r="G2" s="5"/>' : ''}
        </row>`;
        mergeCellsList.push("A2:C2");
        mergeCellsList.push(`D2:${lastColLetter}2`);

        // Row 3: Grade & Class
        sheetXml += `
        <row r="3" ht="20" customHeight="1">
            <c r="A3" s="3" t="inlineStr"><is><t>${escapeXml(settings.grade || 'KHỐI 5')} - ${escapeXml(settings.className || 'LỚP 5A')}</t></is></c>
            <c r="B3" s="3"/><c r="C3" s="3"/>
            <c r="D3" s="0"/><c r="E3" s="0"/><c r="F3" s="0"/>${isCtlop ? '<c r="G3" s="0"/>' : ''}
        </row>`;
        mergeCellsList.push("A3:C3");

        // Row 4: Space
        sheetXml += `<row r="4" ht="10" customHeight="1"/>`;

        // Row 5: Title
        sheetXml += `
        <row r="5" ht="26" customHeight="1">
            <c r="A5" s="6" t="inlineStr"><is><t>${escapeXml(mainTitle)}</t></is></c>
            <c r="B5" s="6"/><c r="C5" s="6"/><c r="D5" s="6"/><c r="E5" s="6"/><c r="F5" s="6"/>${isCtlop ? '<c r="G5" s="6"/>' : ''}
        </row>`;
        mergeCellsList.push(`A5:${lastColLetter}5`);

        // Row 6: Subtitle
        sheetXml += `
        <row r="6" ht="18" customHeight="1">
            <c r="A6" s="7" t="inlineStr"><is><t>${escapeXml(subtitle)}</t></is></c>
            <c r="B6" s="7"/><c r="C6" s="7"/><c r="D6" s="7"/><c r="E6" s="7"/><c r="F6" s="7"/>${isCtlop ? '<c r="G6" s="7"/>' : ''}
        </row>`;
        mergeCellsList.push(`A6:${lastColLetter}6`);

        // Row 7: Space
        sheetXml += `<row r="7" ht="10" customHeight="1"/>`;

        // Row 8: Table Header
        sheetXml += `
        <row r="8" ht="26" customHeight="1">
            <c r="A8" s="8" t="inlineStr"><is><t>Thứ, ngày</t></is></c>
            <c r="B8" s="8" t="inlineStr"><is><t>Buổi</t></is></c>
            <c r="C8" s="8" t="inlineStr"><is><t>Tiết</t></is></c>
            <c r="D8" s="8" t="inlineStr"><is><t>Môn học</t></is></c>
            <c r="E8" s="8" t="inlineStr"><is><t>Tiết PPCT</t></is></c>
            <c r="F8" s="8" t="inlineStr"><is><t>Tên bài dạy</t></is></c>
            ${isCtlop ? '<c r="G8" s="8" t="inlineStr"><is><t>Nội dung tích hợp / Điều chỉnh</t></is></c>' : ''}
        </row>`;

        // Table Data Rows
        let currentRow = 9;
        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];

        days.forEach(day => {
            const daySlots = schedule.filter(s => s.day === day);
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
                const isFirstOfMorning = (slot.session === "Sáng" && idx === 0);
                const isFirstOfAfternoon = (slot.session === "Chiều" && idx === morningSlots.length);

                const dayVal = isFirstOfDay ? (dayDate ? `${day}\n${dayDate}` : day) : "";
                const sessionVal = isFirstOfMorning ? "Sáng" : (isFirstOfAfternoon ? "Chiều" : "");

                // Calculate comfortable row height based on text length
                const lessonLen = (slot.lessonName || '').length;
                const integLen = (slot.integration || '').length;
                const lessonLines = Math.ceil(lessonLen / (isCtlop ? 30 : 44));
                const integLines = isCtlop ? Math.ceil(integLen / 26) : 1;
                const numLines = Math.max(isFirstOfDay && dayDate ? 2 : 1, lessonLines, integLines);

                let rowHt = 24;
                if (numLines === 2) rowHt = 38;
                else if (numLines === 3) rowHt = 54;
                else if (numLines >= 4) rowHt = numLines * 18;

                sheetXml += `
        <row r="${r}" ht="${rowHt}" customHeight="1">
            <c r="A${r}" s="9" ${dayVal ? 't="inlineStr"' : ''}>${dayVal ? `<is><t xml:space="preserve">${escapeXml(dayVal)}</t></is>` : ''}</c>
            <c r="B${r}" s="10" ${sessionVal ? 't="inlineStr"' : ''}>${sessionVal ? `<is><t>${escapeXml(sessionVal)}</t></is>` : ''}</c>
            <c r="C${r}" s="11" t="inlineStr"><is><t>${escapeXml(slot.period || '')}</t></is></c>
            <c r="D${r}" s="12" t="inlineStr"><is><t>${escapeXml(slot.isOff ? '-- Nghỉ / Để trống --' : (slot.subject || ''))}</t></is></c>
            <c r="E${r}" s="13" t="inlineStr"><is><t>${escapeXml(slot.ppct || '')}</t></is></c>
            <c r="F${r}" s="14" t="inlineStr"><is><t>${escapeXml(slot.lessonName || '')}</t></is></c>
            ${isCtlop ? `<c r="G${r}" s="15" t="inlineStr"><is><t>${escapeXml(slot.integration || '')}</t></is></c>` : ''}
        </row>`;
                currentRow++;
            });

            const dayEndRow = currentRow - 1;
            if (dayEndRow > dayStartRow) {
                mergeCellsList.push(`A${dayStartRow}:A${dayEndRow}`);
            }
            if (morningSlots.length > 1) {
                mergeCellsList.push(`B${morningStartRow}:B${morningEndRow}`);
            }
            if (afternoonSlots.length > 1) {
                mergeCellsList.push(`B${afternoonStartRow}:B${afternoonEndRow}`);
            }
        });

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
            <c r="C${footerTitleRow}" s="16" t="inlineStr"><is><t>KHỐI TRƯỞNG</t></is></c>
            <c r="D${footerTitleRow}" s="16"/>
            <c r="E${footerTitleRow}" s="16" t="inlineStr"><is><t>GIÁO VIÊN CHỦ NHIỆM</t></is></c>
            <c r="F${footerTitleRow}" s="16"/>${isCtlop ? `<c r="G${footerTitleRow}" s="16"/>` : ''}
        </row>
        <row r="${footerSubRow}" ht="18" customHeight="1">
            <c r="A${footerSubRow}" s="17" t="inlineStr"><is><t>(Ký và ghi rõ họ tên)</t></is></c>
            <c r="B${footerSubRow}" s="17"/>
            <c r="C${footerSubRow}" s="17" t="inlineStr"><is><t>(Ký và ghi rõ họ tên)</t></is></c>
            <c r="D${footerSubRow}" s="17"/>
            <c r="E${footerSubRow}" s="17" t="inlineStr"><is><t>(Ký và ghi rõ họ tên)</t></is></c>
            <c r="F${footerSubRow}" s="17"/>${isCtlop ? `<c r="G${footerSubRow}" s="17"/>` : ''}
        </row>
        <row r="${footerSpaceRow1}" ht="18" customHeight="1"/>
        <row r="${footerSpaceRow2}" ht="18" customHeight="1"/>
        <row r="${footerSpaceRow3}" ht="18" customHeight="1"/>
        <row r="${footerNameRow}" ht="22" customHeight="1">
            <c r="A${footerNameRow}" s="18" t="inlineStr"><is><t>${escapeXml(bghSignerName)}</t></is></c>
            <c r="B${footerNameRow}" s="18"/>
            <c r="C${footerNameRow}" s="18" t="inlineStr"><is><t>${escapeXml(settings.headOfGrade || 'Trần Thị Mai')}</t></is></c>
            <c r="D${footerNameRow}" s="18"/>
            <c r="E${footerNameRow}" s="18" t="inlineStr"><is><t>${escapeXml(settings.homeroomTeacher || 'Nguyễn Thị Thu Hà')}</t></is></c>
            <c r="F${footerNameRow}" s="18"/>${isCtlop ? `<c r="G${footerNameRow}" s="18"/>` : ''}
        </row>`;

        mergeCellsList.push(`A${footerTitleRow}:B${footerTitleRow}`);
        mergeCellsList.push(`C${footerTitleRow}:D${footerTitleRow}`);
        mergeCellsList.push(`E${footerTitleRow}:${lastColLetter}${footerTitleRow}`);

        mergeCellsList.push(`A${footerSubRow}:B${footerSubRow}`);
        mergeCellsList.push(`C${footerSubRow}:D${footerSubRow}`);
        mergeCellsList.push(`E${footerSubRow}:${lastColLetter}${footerSubRow}`);

        mergeCellsList.push(`A${footerNameRow}:B${footerNameRow}`);
        mergeCellsList.push(`C${footerNameRow}:D${footerNameRow}`);
        mergeCellsList.push(`E${footerNameRow}:${lastColLetter}${footerNameRow}`);

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
