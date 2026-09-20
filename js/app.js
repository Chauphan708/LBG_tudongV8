/**
 * LỊCH BÁO GIẢNG & KẾ HOẠCH DẠY HỌC LỚP 5 - MAIN APPLICATION
 * Features: 5 Tabs, Realtime Search & Calculation, Friday Afternoon Customization,
 * Smart 35-Week Generator, PPCT Excel Import/Export Template, Decree 30/2020 Preview & DOCX/XLSX Export
 */

(function() {
    const STORAGE_KEY = "LBG_APP_DATA_V8";
    let currentOrientation = "portrait";

    function getBaseLesson(lessonStr) {
        if (!lessonStr) return '';
        return lessonStr.replace(/\s*[\(\,\–\-]\s*(?:tiết|Tiết)\s*\d+[\)\.]?/gi, '').trim();
    }

    /**
     * Compute vertical row spans for KHDH table cells (rowspan)
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

    const DEFAULT_GRADE_5_SUBJECTS = [
        { name: "Tiếng Việt", category: "GVCN", defaultPeriods: 7, isIncluded: true },
        { name: "Toán", category: "GVCN", defaultPeriods: 5, isIncluded: true },
        { name: "HĐ Trải nghiệm", category: "GVCN", defaultPeriods: 3, isIncluded: true },
        { name: "Khoa học", category: "GVCN", defaultPeriods: 2, isIncluded: true },
        { name: "LS&ĐL", category: "GVCN", defaultPeriods: 2, isIncluded: true },
        { name: "Đạo đức", category: "GVCN", defaultPeriods: 1, isIncluded: true },
        { name: "Công nghệ", category: "GVCN", defaultPeriods: 1, isIncluded: true },
        { name: "Tin học", category: "Chuyên trách", defaultPeriods: 1, isIncluded: false },
        { name: "Tiếng Anh", category: "Chuyên trách", defaultPeriods: 4, isIncluded: false },
        { name: "GD Thể chất", category: "Chuyên trách", defaultPeriods: 2, isIncluded: false },
        { name: "Âm nhạc", category: "Chuyên trách", defaultPeriods: 1, isIncluded: false },
        { name: "Mĩ thuật", category: "Chuyên trách", defaultPeriods: 1, isIncluded: false },
        { name: "TC Tiếng Việt", category: "Tăng cường", defaultPeriods: 2, isIncluded: true },
        { name: "TC Toán", category: "Tăng cường", defaultPeriods: 2, isIncluded: true },
        { name: "KNS", category: "Tăng cường", defaultPeriods: 1, isIncluded: false },
        { name: "STEM", category: "Tăng cường", defaultPeriods: 1, isIncluded: false },
        { name: "CD Số", category: "Tăng cường", defaultPeriods: 1, isIncluded: false },
        { name: "Đọc Thư viện", category: "Tăng cường", defaultPeriods: 1, isIncluded: false }
    ];
    
    const DEFAULT_BASES = [
        "Căn cứ Thông tư số 32/2018/TT-BGDĐT ngày 26/12/2018 Ban hành Chương trình Giáo dục phổ thông tổng thể.",
        "Căn cứ Thông tư 27/2020/TT-BGDĐT ngày 04/9/2020 về việc ban hành Quy định đánh giá học sinh tiểu học;",
        "Căn cứ Công văn số 2345/BGDĐT-GDTH ngày 07/6/2021 về việc hướng dẫn xây dựng kế hoạch giáo dục nhà trường."
    ];

    let state = {
        currentTab: "tab-lbg",
        currentWeek: 1,
        lbgMonFilterSubject: "all",
        lbgMonFilterCategory: "all",
        lbgMonShowIntegration: true,
        lbgMonOrientation: "portrait",
        selectedKhdhSubjects: ["TIẾNG VIỆT", "TOÁN", "KHOA HỌC", "LS&ĐL", "HĐ TRẢI NGHIỆM", "ĐẠO ĐỨC", "CÔNG NGHỆ"],
        includedSubjects: [],
        subjectList: [],
        lbgExcludedMode: "off",
        lbgShowColSign: false,
        lbgShowColNote: false,
        lbgShowColCustom: false,
        lbgColCustomName: "Ghi chú",
        lbgColCustomPos: "end",
        lbgCustomCols: [],
        lbgShowBghSign: true,
        lbgShowGvcnSign: true,
        lbgShowHeadSign: true,
        settings: {
            governingBody: "UBND PHƯỜNG TRUNG NHỨT",
            schoolName: "TRƯỜNG TIỂU HỌC TRUNG NHỨT",
            grade: "KHỐI 5",
            className: "LỚP 5A",
            academicYear: "2026 - 2027",
            homeroomTeacher: "Nguyễn Thị Thu Hà",
            headOfGrade: "Trần Thị Mai",
            principal: "Phạm Quốc Hùng",
            vicePrincipal: "Lê Văn Tám",
            vicePrincipals: ["Lê Văn Tám"],
            bghSignerType: "PHT",
            bghSignerIndex: 0,
            bghSignerLbgType: "PHT",
            bghSignerLbgIndex: 0,
            bghSignerKhdhType: "HT",
            bghSignerKhdhIndex: 0,
            location: "Trung Nhứt",
            dateString: "ngày 28 tháng 8 năm 2026",
            bases: [...DEFAULT_BASES]
        },
        fridayPeriodsDefault: 3,
        weeklyFridayPeriods: {},
        weeklyCustomSlots: {},
        weeks: [],
        ppct: [],
        khdh: {},
        timetable: [],
        weeklyScheduleOverrides: {}
    };

    let currentGrade = parseInt(localStorage.getItem("LBG_CURRENT_GRADE") || "5", 10);
    if (isNaN(currentGrade) || currentGrade < 1 || currentGrade > 5) currentGrade = 5;
    state.currentGrade = currentGrade;

    function getStorageKey(grade) {
        return "LBG_APP_DATA_V8_G" + grade;
    }

    function getGradeDefaultSubjects(grade) {
        if (grade === 5) {
            return JSON.parse(JSON.stringify(DEFAULT_GRADE_5_SUBJECTS));
        }
        const gData = window.APP_GRADE_DATA && window.APP_GRADE_DATA[grade];
        if (gData && gData.subjects) {
            return JSON.parse(JSON.stringify(gData.subjects));
        }
        return JSON.parse(JSON.stringify(DEFAULT_GRADE_5_SUBJECTS));
    }

    function getGradeDefaultTimetable(grade) {
        if (grade === 5) {
            if (window.APP_INITIAL_DATA && window.APP_INITIAL_DATA.settings && window.APP_INITIAL_DATA.settings.timetable) {
                return JSON.parse(JSON.stringify(window.APP_INITIAL_DATA.settings.timetable));
            }
        }
        const gData = window.APP_GRADE_DATA && window.APP_GRADE_DATA[grade];
        if (gData && gData.timetable) {
            return JSON.parse(JSON.stringify(gData.timetable));
        }
        return [];
    }

    function sanitizeTimetable(timetable, grade) {
        const defaultTt = getGradeDefaultTimetable(grade);
        if (!Array.isArray(timetable) || timetable.length === 0) {
            return defaultTt;
        }
        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];
        const cleaned = [];

        // Check if Monday is corrupt (e.g. has Tiếng Anh in Sáng, or != 7 slots, or Sáng > 4 slots)
        const monSlots = timetable.filter(s => s.day === "Thứ 2");
        const hasMonCorrupt = monSlots.length !== 7 || 
            monSlots.some(s => s.session === "Sáng" && s.subject && s.subject.toLowerCase().includes("tiếng anh")) ||
            !monSlots.some(s => s.session === "Chiều" && s.period === 1 && s.subject && s.subject.toLowerCase().includes("tiếng anh"));

        days.forEach(day => {
            let daySlots = timetable.filter(s => s.day === day);
            if (day === "Thứ 2" && hasMonCorrupt) {
                daySlots = defaultTt.filter(s => s.day === "Thứ 2");
            } else if (daySlots.length === 0) {
                daySlots = defaultTt.filter(s => s.day === day);
            }

            // Group & Sort Sáng
            const morn = daySlots.filter(s => s.session === "Sáng").sort((a, b) => (a.period || 0) - (b.period || 0));
            morn.forEach((s, i) => { s.period = i + 1; s.session = "Sáng"; s.day = day; });

            // Group & Sort Chiều
            const aft = daySlots.filter(s => s.session === "Chiều").sort((a, b) => (a.period || 0) - (b.period || 0));
            aft.forEach((s, i) => { s.period = i + 1; s.session = "Chiều"; s.day = day; });

            cleaned.push(...morn, ...aft);
        });

        return cleaned;
    }

    function sanitizeSubjectList(subjectList, grade) {
        if (!Array.isArray(subjectList) || subjectList.length === 0) {
            return getGradeDefaultSubjects(grade);
        }
        subjectList.forEach(s => {
            const norm = normalizeSubjectName(s.name);
            if (norm === "Toán") {
                // Theo Chương trình GDPT 2018 (Thông tư 32/2018): Lớp 1 là 3 tiết/tuần (105 tiết); Lớp 2, 3, 4, 5 là 5 tiết/tuần (175 tiết)
                s.defaultPeriods = (grade === 1) ? 3 : 5;
            } else if (norm === "HĐ Trải nghiệm") {
                // Hoạt động trải nghiệm là 3 tiết/tuần (105 tiết/năm)
                s.defaultPeriods = 3;
            }
        });
        return subjectList;
    }

    function loadStateForGrade(grade) {
        state.currentGrade = grade;

        // Reset temporary overrides and state to avoid cross-contamination between grades
        state.weeklyScheduleOverrides = {};
        state.weeklyFridayPeriods = {};
        state.weeklyCustomSlots = {};

        // 1. Initial base data for this grade
        if (grade === 5) {
            if (window.APP_INITIAL_DATA) {
                state.weeks = JSON.parse(JSON.stringify(window.APP_INITIAL_DATA.weeks || []));
                state.ppct = JSON.parse(JSON.stringify(window.APP_INITIAL_DATA.ppct || []));
                state.khdh = JSON.parse(JSON.stringify(window.APP_INITIAL_DATA.khdh || {}));
                state.timetable = JSON.parse(JSON.stringify(window.APP_INITIAL_DATA.settings?.timetable || []));
                Object.assign(state.settings, window.APP_INITIAL_DATA.settings || {});
            }
            state.subjectList = JSON.parse(JSON.stringify(DEFAULT_GRADE_5_SUBJECTS));
            state.selectedKhdhSubjects = ["TIẾNG VIỆT", "TOÁN", "KHOA HỌC", "LS&ĐL", "HĐ TRẢI NGHIỆM", "ĐẠO ĐỨC", "CÔNG NGHỆ"];
        } else {
            const gData = window.APP_GRADE_DATA && window.APP_GRADE_DATA[grade];
            if (gData) {
                state.ppct = JSON.parse(JSON.stringify(gData.ppct || []));
                state.timetable = JSON.parse(JSON.stringify(gData.timetable || []));
                state.settings.grade = gData.gradeName;
                state.settings.className = gData.className;
                state.subjectList = JSON.parse(JSON.stringify(gData.subjects || []));
                if (window.APP_INITIAL_DATA && window.APP_INITIAL_DATA.weeks) {
                    state.weeks = JSON.parse(JSON.stringify(window.APP_INITIAL_DATA.weeks));
                }
                state.khdh = buildKhdhForGrade(grade, state.ppct);
                if (grade === 1 || grade === 2) {
                    state.selectedKhdhSubjects = ["TIẾNG VIỆT", "TOÁN", "TỰ NHIÊN VÀ XÃ HỘI", "HĐ TRẢI NGHIỆM", "ĐẠO ĐỨC"];
                } else if (grade === 3) {
                    state.selectedKhdhSubjects = ["TIẾNG VIỆT", "TOÁN", "TỰ NHIÊN VÀ XÃ HỘI", "HĐ TRẢI NGHIỆM", "ĐẠO ĐỨC", "TIN HỌC", "CÔNG NGHỆ"];
                } else if (grade === 4) {
                    state.selectedKhdhSubjects = ["TIẾNG VIỆT", "TOÁN", "KHOA HỌC", "LS&ĐL", "HĐ TRẢI NGHIỆM", "ĐẠO ĐỨC", "TIN HỌC", "CÔNG NGHỆ"];
                }
            }
        }

        // 2. Load custom modifications from localStorage
        try {
            const gradeKey = getStorageKey(grade);
            let saved = localStorage.getItem(gradeKey);
            if (!saved) {
                saved = localStorage.getItem("LBG_APP_DATA_V7_G" + grade) || localStorage.getItem("LBG_APP_DATA_V6_G" + grade) || localStorage.getItem("LBG_APP_DATA_V5_G" + grade) || localStorage.getItem("LBG_APP_DATA_V4_G" + grade);
            }
            if (!saved && grade === 5) {
                // Seamless migration from earlier keys
                saved = localStorage.getItem("LBG_APP_DATA_V7") || localStorage.getItem("LBG_APP_DATA_V6") || localStorage.getItem("LBG_APP_DATA_V5") || localStorage.getItem("LBG_APP_DATA_V4") || localStorage.getItem("LBG_APP_DATA_V3") || localStorage.getItem("LBG_APP_DATA_V2") || localStorage.getItem("LBG_APP_DATA_V1");
            }
            if (saved) {
                const parsed = JSON.parse(saved);
                if (parsed.settings) {
                    Object.assign(state.settings, parsed.settings);
                    if (!state.settings.bghSignerLbgType) {
                        state.settings.bghSignerLbgType = parsed.settings.bghSignerType || "PHT";
                        state.settings.bghSignerLbgIndex = parsed.settings.bghSignerIndex || 0;
                    }
                    if (!state.settings.bghSignerKhdhType) {
                        state.settings.bghSignerKhdhType = "HT";
                        state.settings.bghSignerKhdhIndex = 0;
                    }
                }
                if (parsed.weeks && parsed.weeks.length) state.weeks = parsed.weeks;
                if (parsed.ppct && parsed.ppct.length) {
                    state.ppct = parsed.ppct;
                    if (grade !== 5) {
                        state.khdh = buildKhdhForGrade(grade, state.ppct);
                    }
                }
                // Smart Auto-Healing for Grade 5 Integration: Restore full official 621 integration items if wiped in earlier sessions
                if (grade === 5 && window.APP_INITIAL_DATA && Array.isArray(window.APP_INITIAL_DATA.ppct)) {
                    const defaultPpct = window.APP_INITIAL_DATA.ppct;
                    const validCount = state.ppct.filter(x => x.integration && x.integration.trim()).length;
                    if (validCount < 500) {
                        state.ppct.forEach(item => {
                            if (!item.integration || !item.integration.trim()) {
                                const normSub = normalizeSubjectName(item.subject);
                                const found = defaultPpct.find(d =>
                                    d.week === item.week &&
                                    normalizeSubjectName(d.subject) === normSub &&
                                    (d.ppct === item.ppct || d.periodInWeek === item.periodInWeek)
                                );
                                if (found && found.integration) {
                                    item.integration = found.integration;
                                }
                            }
                        });
                    }
                }
                if (parsed.timetable && parsed.timetable.length) state.timetable = parsed.timetable;
                if (parsed.weeklyCustomSlots) state.weeklyCustomSlots = parsed.weeklyCustomSlots;
                if (parsed.weeklyScheduleOverrides) {
                    // Clean up any stale overrides carried over from Grade 5 into other grades
                    if (grade !== 5) {
                        for (const k in parsed.weeklyScheduleOverrides) {
                            const ov = parsed.weeklyScheduleOverrides[k];
                            if (ov && ov.integration && (
                                ov.integration.includes("Thanh âm của gió") || 
                                ov.integration.includes("Lý tưởng cách mạng, đạo đức, lối sống: Giáo dục tình yêu quê hương, biết lắng nghe")
                            )) {
                                delete ov.integration;
                            }
                        }
                    }
                    state.weeklyScheduleOverrides = parsed.weeklyScheduleOverrides;
                }
                if (parsed.weeklyFridayPeriods) state.weeklyFridayPeriods = parsed.weeklyFridayPeriods;
                if (parsed.currentWeek) state.currentWeek = parsed.currentWeek;
                if (parsed.fridayPeriodsDefault !== undefined) state.fridayPeriodsDefault = parsed.fridayPeriodsDefault;
                if (parsed.includedSubjects && parsed.includedSubjects.length) state.includedSubjects = parsed.includedSubjects;
                if (parsed.subjectList && Array.isArray(parsed.subjectList) && parsed.subjectList.length) {
                    state.subjectList = parsed.subjectList;
                }
                if (parsed.lbgExcludedMode) state.lbgExcludedMode = parsed.lbgExcludedMode;
                if (parsed.lbgShowColSign !== undefined) state.lbgShowColSign = parsed.lbgShowColSign;
                if (parsed.lbgShowColNote !== undefined) state.lbgShowColNote = parsed.lbgShowColNote;
                if (parsed.lbgShowColCustom !== undefined) state.lbgShowColCustom = parsed.lbgShowColCustom;
                if (parsed.lbgColCustomName) state.lbgColCustomName = parsed.lbgColCustomName;
                if (parsed.lbgColCustomPos) state.lbgColCustomPos = parsed.lbgColCustomPos;
                if (parsed.lbgCustomCols && Array.isArray(parsed.lbgCustomCols)) {
                    state.lbgCustomCols = parsed.lbgCustomCols;
                } else if (parsed.lbgShowColCustom) {
                    state.lbgCustomCols = [
                        { id: "col_1", name: parsed.lbgColCustomName || "Ghi chú", pos: parsed.lbgColCustomPos || "end", enabled: true }
                    ];
                } else {
                    state.lbgCustomCols = [];
                }
                if (parsed.lbgShowBghSign !== undefined) state.lbgShowBghSign = parsed.lbgShowBghSign;
                if (parsed.lbgShowGvcnSign !== undefined) state.lbgShowGvcnSign = parsed.lbgShowGvcnSign;
                if (parsed.lbgShowHeadSign !== undefined) state.lbgShowHeadSign = parsed.lbgShowHeadSign;
                if (parsed.lbgHideEmptyRows !== undefined) state.lbgHideEmptyRows = parsed.lbgHideEmptyRows;
            }
        } catch (e) {
            console.warn("Could not load localStorage for Grade " + grade + ":", e);
        }

        // Auto migrate weeks if unshifted (e.g. week 23 starting 08/02/2027 during Tet)
        if (state.weeks && state.weeks.length === 35) {
            const w23 = state.weeks.find(w => w.week === 23);
            if (w23 && w23.startDateVN === "08/02/2027") {
                autoGenerateCalendar(
                    state.settings.academicYear || "2026 - 2027",
                    state.settings.startDate || "2026-09-07",
                    state.settings.tetStartDate || "2027-02-08",
                    state.settings.tetEndDate || "2027-02-21"
                );
            }
        }

        // Sanitize and heal Timetable: ensure all days Sáng (1..4) then Chiều (1..3), Monday has Tiếng Anh at Chiều T1
        state.timetable = sanitizeTimetable(state.timetable, grade);

        // Sanitize Subject List: ensure Toán is 5 periods/week for Grade 2, 3 and all grades
        state.subjectList = sanitizeSubjectList(state.subjectList, grade);

        // Sanitize weeklyCustomSlots if any
        if (state.weeklyCustomSlots) {
            for (const wk in state.weeklyCustomSlots) {
                if (Array.isArray(state.weeklyCustomSlots[wk])) {
                    state.weeklyCustomSlots[wk] = sanitizeTimetable(state.weeklyCustomSlots[wk], grade);
                }
            }
        }

        // Standardize all PPCT & Timetable subjects
        if (state.ppct) {
            state.ppct.forEach(p => {
                p.subject = normalizeSubjectName(p.subject);
                if (p.lessonName) p.lessonName = normalizePunctuationSpacing(p.lessonName);
                if (p.integration) p.integration = normalizePunctuationSpacing(p.integration);
            });
        }
        if (state.settings) {
            if (!Array.isArray(state.settings.bases) || state.settings.bases.length === 0 || state.settings.bases.some(b => b.includes("28/2020"))) {
                state.settings.bases = [...DEFAULT_BASES];
            } else {
                state.settings.bases = state.settings.bases.map(b => normalizePunctuationSpacing(b));
            }
        }
        if (state.timetable) {
            state.timetable.forEach(t => {
                t.subject = normalizeSubjectName(t.subject);
            });
        }

        // Sync state.includedSubjects with state.subjectList
        if (!state.includedSubjects || state.includedSubjects.length === 0) {
            state.includedSubjects = state.subjectList.filter(s => s.isIncluded !== false).map(s => s.name);
        } else {
            const inclSet = new Set(state.includedSubjects);
            state.subjectList.forEach(s => {
                if (s.isIncluded !== false) inclSet.add(s.name);
            });
            state.includedSubjects = Array.from(inclSet);
        }
    }

    function loadState() {
        loadStateForGrade(state.currentGrade);
    }

    function saveState() {
        try {
            const payload = {
                settings: state.settings,
                weeks: state.weeks,
                ppct: state.ppct,
                timetable: state.timetable,
                weeklyScheduleOverrides: state.weeklyScheduleOverrides,
                weeklyFridayPeriods: state.weeklyFridayPeriods,
                currentWeek: state.currentWeek,
                fridayPeriodsDefault: state.fridayPeriodsDefault,
                includedSubjects: state.includedSubjects,
                subjectList: state.subjectList,
                lbgExcludedMode: state.lbgExcludedMode,
                lbgShowColSign: state.lbgShowColSign,
                lbgShowColNote: state.lbgShowColNote,
                lbgShowColCustom: (Array.isArray(state.lbgCustomCols) && state.lbgCustomCols.some(c => c.enabled !== false)) || state.lbgShowColCustom,
                lbgColCustomName: (state.lbgCustomCols && state.lbgCustomCols[0] && state.lbgCustomCols[0].name) || state.lbgColCustomName,
                lbgColCustomPos: (state.lbgCustomCols && state.lbgCustomCols[0] && state.lbgCustomCols[0].pos) || state.lbgColCustomPos || "end",
                lbgCustomCols: state.lbgCustomCols || [],
                lbgShowBghSign: state.lbgShowBghSign,
                lbgShowGvcnSign: state.lbgShowGvcnSign,
                lbgShowHeadSign: state.lbgShowHeadSign,
                lbgHideEmptyRows: !!state.lbgHideEmptyRows,
                currentGrade: state.currentGrade
            };
            const gradeKey = getStorageKey(state.currentGrade);
            localStorage.setItem(gradeKey, JSON.stringify(payload));
            if (state.currentGrade === 5) {
                localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
                localStorage.setItem("LBG_APP_DATA_V8", JSON.stringify(payload));
                localStorage.setItem("LBG_APP_DATA_V7", JSON.stringify(payload));
                localStorage.setItem("LBG_APP_DATA_V6", JSON.stringify(payload));
                localStorage.setItem("LBG_APP_DATA_V5", JSON.stringify(payload));
                localStorage.setItem("LBG_APP_DATA_V4", JSON.stringify(payload));
            }
        } catch (e) {
            console.warn("Could not save to localStorage:", e);
        }
    }

    function switchGrade(newGrade) {
        newGrade = parseInt(newGrade, 10);
        if (isNaN(newGrade) || newGrade < 1 || newGrade > 5) return;
        if (newGrade === state.currentGrade) return;

        // 1. Save current grade state
        saveState();

        // 2. Switch to new grade
        state.currentGrade = newGrade;
        localStorage.setItem("LBG_CURRENT_GRADE", newGrade.toString());
        loadStateForGrade(newGrade);

        // 3. Update UI
        const gradeSelect = document.getElementById("global-grade-selector");
        if (gradeSelect) gradeSelect.value = newGrade;

        const ppctTitle = document.getElementById("ppct-grade-title");
        if (ppctTitle) ppctTitle.textContent = `Khối ${newGrade}`;

        const badge = document.getElementById("top-badge-school");
        if (badge) badge.textContent = `🏫 ${state.settings.schoolName || 'TH TRUNG NHỨT'} — ${state.settings.className} (${state.settings.grade})`;

        renderLbgInclusionChips();
        renderTabLbg(state.currentWeek);
        renderTabCtlop(state.currentWeek);
        renderTabLbgMon();
        renderTabLichtuan();
        renderTabPpct();
        renderTabSettings();
        updatePpctCountBadge();

        showToast(`Đã chuyển sang Khối ${newGrade} (${state.settings.className})! Dữ liệu được lưu độc lập.`, 'info');
    }

    function updatePpctCountBadge() {
        const badge = document.getElementById("ppct-count-badge");
        if (badge) {
            const count = (state.ppct || []).length;
            badge.innerText = `${count} tiết`;
        }
    }

    function showToast(message, type = "info") {
        const container = document.getElementById("toast-container");
        if (!container) return;
        const toast = document.createElement("div");
        toast.className = `toast toast-${type}`;
        toast.innerHTML = `<span>✓</span> <span>${message}</span>`;
        container.appendChild(toast);
        setTimeout(() => {
            toast.style.opacity = "0";
            toast.style.transform = "translateY(10px)";
            toast.style.transition = "all 0.3s ease";
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    }

    // 2. Canonical Subject Definition & Robust Normalization
    const SUBJECT_LIST = [
        "HĐ Trải nghiệm",
        "Tiếng Việt",
        "Toán",
        "TNXH",
        "Khoa học",
        "LS&ĐL",
        "Đạo đức",
        "Tin học",
        "Công nghệ",
        "Tiếng Anh",
        "GD Thể chất",
        "Âm nhạc",
        "Mĩ thuật",
        "Đọc Thư viện",
        "TC Tiếng Việt",
        "TC Toán",
        "KNS",
        "STEM",
        "CD Số",
        "-- Nghỉ / Để trống --"
    ];

    function normalizePunctuationSpacing(str) {
        if (!str || typeof str !== "string") return str;
        
        let res = str;

        // 1. Remove spaces before punctuation marks: , . ? ! ; : ) ] }
        res = res.replace(/\s+([,\.\?\!;:\)\]\}])/g, "$1");
        res = res.replace(/([\(\[\{])\s+/g, "$1");

        // 2. Ensure 1 space after comma (,), semicolon (;), question mark (?), exclamation mark (!)
        res = res.replace(/([,\?;!])(?=[a-zA-ZÀ-ỹ0-9])/g, "$1 ");

        // 3. Ensure 1 space after colon (:) if followed by a non-space (unless part of URL http://)
        res = res.replace(/(:)(?=[^\s\/\\])/g, "$1 ");

        // 4. Ensure 1 space after period (.) when ending a word or sentence
        res = res.replace(/([a-zA-ZÀ-ỹ]{2,})\.(?=[A-ZÀ-Ỹa-zà-ỹ0-9])/g, "$1. ");
        res = res.replace(/(Bài \d+)\.(?=[^\s\d])/gi, "$1. ");
        res = res.replace(/(Tiết \d+)\.(?=[^\s\d])/gi, "$1. ");
        res = res.replace(/(Tuần \d+)\.(?=[^\s\d])/gi, "$1. ");
        res = res.replace(/(Điều \d+)\.(?=[^\s\d])/gi, "$1. ");
        res = res.replace(/(Khoản \d+)\.(?=[^\s\d])/gi, "$1. ");

        // 5. Ensure 1 space after hyphen / dash
        // Bullet dash at beginning of line / string: "-Mục tiêu" -> "- Mục tiêu"
        res = res.replace(/(^|\n)\s*[-–—]\s*(?=[a-zA-ZÀ-ỹ0-9])/g, "$1- ");
        // Dash between words or numbers (excluding legal doc numbers like TT-BGDĐT with all-caps acronyms)
        res = res.replace(/([a-zA-ZÀ-ỹ0-9]{3,})\s*[-–—]\s*([a-zA-ZÀ-ỹ0-9])/g, "$1 - $2");
        res = res.replace(/(\d{4})\s*[-–—]\s*(\d{4})/g, "$1 - $2");

        // 6. Clean multiple consecutive spaces (preserve newlines)
        res = res.replace(/[^\S\r\n]{2,}/g, " ");

        return res.trim();
    }

    function normalizeSubjectName(sub) {
        if (!sub) return "";
        const s = sub.toString().normalize("NFC").trim().toUpperCase()
            .replace(/À|Á|Ạ|Ả|Ã|Â|Ầ|Ấ|Ậ|Ẩ|Ẫ|Ă|Ằ|Ắ|Ặ|Ẳ|Ẵ/g, "A")
            .replace(/È|É|Ẹ|Ẻ|Ẽ|Ê|Ề|Ế|Ệ|Ể|Ễ/g, "E")
            .replace(/Ì|Í|Ị|Ỉ|Ĩ/g, "I")
            .replace(/Ò|Ó|Ọ|Ỏ|Õ|Ô|Ồ|Ố|Ộ|Ổ|Ỗ|Ơ|Ờ|Ớ|Ợ|Ở|Ỡ/g, "O")
            .replace(/Ù|Ú|Ụ|Ủ|Ũ|Ư|Ừ|Ứ|Ự|Ử|Ữ/g, "U")
            .replace(/Ỳ|Ý|Ỵ|Ỷ|Ỹ/g, "Y")
            .replace(/Đ/g, "D");
        
        // Priority 1: HĐ Trải nghiệm (MUST check before any off/empty substring checks!)
        if (s.includes("TRAI NGHIEM") || s.includes("HDTN") || s.includes("HOAT DONG") || s.includes("CHAO CO") || s.includes("SINH HOAT")) {
            return "HĐ Trải nghiệm";
        }

        // Priority 2: TC Tiếng Việt / Tiếng Việt
        if (s.includes("TC TIENG VIET") || s.includes("TANG CUONG TIENG VIET") || s.includes("TC TV")) return "TC Tiếng Việt";
        if (s.includes("TIENG VIET") || s.includes("TV")) return "Tiếng Việt";

        // Priority 3: TC Toán / Toán
        if (s.includes("TC TOAN") || s.includes("TANG CUONG TOAN")) return "TC Toán";
        if (s.includes("TOAN")) return "Toán";

        // Priority 4: Other Subjects
        if (s.includes("TIENG ANH") || s.includes("TA") || s.includes("ENGLISH")) return "Tiếng Anh";
        if (s.includes("TNXH") || s.includes("TU NHIEN VA XA HOI") || s.includes("TU NHIEN XA HOI") || s.includes("TU NHIEN")) return "TNXH";
        if (s.includes("KHOA HOC") || s === "KH") return "Khoa học";
        if (s.includes("LICH SU") || s.includes("DIA LI") || s.includes("LS") || s.includes("DL")) return "LS&ĐL";
        if (s.includes("DAO DUC") || s === "DD") return "Đạo đức";
        if (s.includes("TIN HOC") || s === "TH" || s.includes("TIN")) return "Tin học";
        if (s.includes("CONG NGHE") || s === "CN") return "Công nghệ";
        if (s.includes("THE CHAT") || s.includes("GDTC") || s.includes("THE DUC")) return "GD Thể chất";
        if (s.includes("AM NHAC") || s.includes("AN") || s.includes("NHAC")) return "Âm nhạc";
        if (s.includes("MI THUAT") || s.includes("MT") || s.includes("MY THUAT") || s.includes("VE")) return "Mĩ thuật";
        if (s.includes("THU VIEN")) return "Đọc Thư viện";
        if (s.includes("KNS") || s.includes("KY NANG")) return "KNS";
        if (s.includes("STEM")) return "STEM";
        if (s.includes("CD SO") || s.includes("CONG DAN")) return "CD Số";

        // Priority 5: Blank / Off Periods
        if (s.includes("DE TRONG") || s === "NGHI" || s.includes("NGHI HOC") || s.includes("NGHI TIET") || s.startsWith("--")) {
            return "-- Nghỉ / Để trống --";
        }

        return sub.toString().trim();
    }

    function getAllUniqueSubjects() {
        const list = (state.subjectList && state.subjectList.length > 0)
            ? state.subjectList.map(s => s.name)
            : DEFAULT_GRADE_5_SUBJECTS.map(s => s.name);
        
        const map = new Map();
        // 1. Add subjects from subjectList
        list.forEach(s => {
            map.set(normalizeSubjectName(s), s);
        });

        // 2. Add any custom subjects from PPCT
        if (state.ppct) {
            state.ppct.forEach(p => {
                if (p.subject && p.subject !== "-- Nghỉ / Để trống --") {
                    const norm = normalizeSubjectName(p.subject);
                    if (!map.has(norm)) {
                        map.set(norm, p.subject);
                    }
                }
            });
        }

        // 3. Add any custom subjects from timetable
        if (state.timetable) {
            state.timetable.forEach(t => {
                if (t.subject && t.subject !== "-- Nghỉ / Để trống --") {
                    const norm = normalizeSubjectName(t.subject);
                    if (!map.has(norm)) {
                        map.set(norm, t.subject);
                    }
                }
            });
        }

        return Array.from(map.values()).filter(s => s && s !== "-- Nghỉ / Để trống --");
    }

    function renderLbgInclusionChips() {
        const container = document.getElementById("lbg-inclusion-chips");
        if (!container) return;

        const allSubs = getAllUniqueSubjects();
        if (!state.includedSubjects || state.includedSubjects.length === 0) {
            state.includedSubjects = [...allSubs];
        }

        container.innerHTML = allSubs.map(sub => {
            const normSub = normalizeSubjectName(sub);
            const isChecked = state.includedSubjects.some(s => s === sub || normalizeSubjectName(s) === normSub);
            return `
                <label class="subject-chip ${isChecked ? 'selected' : ''}" data-subject="${sub}" style="font-size:0.78rem; padding:0.2rem 0.55rem; cursor:pointer;">
                    <input type="checkbox" value="${sub}" ${isChecked ? 'checked' : ''} style="cursor:pointer;">
                    <span>${isChecked ? '✓ ' : '+ '}${sub}</span>
                </label>
            `;
        }).join('');

        container.querySelectorAll(".subject-chip").forEach(chip => {
            chip.addEventListener("click", (e) => {
                e.preventDefault();
                const sub = chip.dataset.subject;
                const normSub = normalizeSubjectName(sub);
                const isCurrentlySelected = state.includedSubjects.some(s => s === sub || normalizeSubjectName(s) === normSub);
                
                if (isCurrentlySelected) {
                    state.includedSubjects = state.includedSubjects.filter(s => s !== sub && normalizeSubjectName(s) !== normSub);
                } else {
                    state.includedSubjects.push(sub);
                }
                saveState();
                renderLbgInclusionChips();
                renderTabLbg();
                renderTabCtlop();
            });
        });

        const modeSelect = document.getElementById("lbg-excluded-mode");
        if (modeSelect) {
            modeSelect.value = state.lbgExcludedMode || "off";
            modeSelect.onchange = (e) => {
                state.lbgExcludedMode = e.target.value;
                saveState();
                renderTabLbg();
                renderTabCtlop();
            };
        }
    }

    function getGradeSpecificIntegration(grade, subject, week, periodInWeek, lessonName) {
        const norm = normalizeSubjectName(subject);
        const lName = (lessonName || "").toLowerCase();

        // 1. TIẾNG VIỆT (Khối 1, 2, 3, 4)
        if (norm === "Tiếng Việt") {
            if (grade === 1) {
                if (week <= 4) return "Tích hợp kĩ năng sống: Rèn tư thế ngồi học đúng, cách cầm bút; giáo dục nề nếp học tập, giao tiếp lễ phép với thầy cô và bạn bè.";
                if (week <= 10) return "Tích hợp giáo dục bảo vệ môi trường: Giữ gìn vệ sinh lớp học, yêu quý thiên nhiên, bảo vệ giọng nói và đôi mắt.";
                if (week <= 18) return "Tích hợp giáo dục tình cảm gia đình, kính yêu ông bà cha mẹ, yêu thương bạn bè; an toàn giao thông trên đường đến trường.";
                if (week <= 27) return "Tích hợp Quyền trẻ em: Quyền được học tập, vui chơi và chăm sóc sức khỏe; rèn luyện kĩ năng tự phục vụ bản thân.";
                return "Tích hợp bồi dưỡng tình yêu quê hương đất nước, Bác Hồ với thiếu nhi; phát triển năng lực tự chủ và giao tiếp tự tin.";
            } else if (grade === 2) {
                if (week <= 9) return "Tích hợp giáo dục nề nếp học tập, tình bạn, xây dựng môi trường học tập thân thiện; rèn kĩ năng lắng nghe và chia sẻ.";
                if (week <= 18) return "Tích hợp giáo dục bảo vệ môi trường, yêu quý cỏ cây hoa lá và loài vật xung quanh; Quyền con người (QCN).";
                if (week <= 27) return "Tích hợp tình cảm gia đình, lòng biết ơn cha mẹ, thầy cô; rèn luyện kĩ năng giải quyết vấn đề và tự quản.";
                return "Tích hợp giáo dục tình yêu thiên nhiên, đất nước con người Việt Nam; giáo dục an toàn giao thông và văn hóa ứng xử.";
            } else if (grade === 3) {
                if (week <= 9) return "Tích hợp Quyền con người (QCN), phát triển năng lực số (NLS), giáo dục tình bạn và ý thức tổ chức kỷ luật.";
                if (week <= 18) return "Tích hợp giáo dục bảo vệ môi trường, biến đổi khí hậu; giáo dục kĩ năng phòng chống tai nạn thương tích và đuối nước.";
                if (week <= 27) return "Tích hợp giáo dục truyền thống quê hương, lòng nhân ái, sự sẻ chia với người khuyết tật và người có hoàn cảnh khó khăn.";
                return "Tích hợp tình yêu biển đảo, chủ quyền quê hương; giáo dục di sản văn hóa và phát triển năng lực thẩm mĩ.";
            } else if (grade === 4) {
                if (week <= 9) return "Tích hợp Quyền con người (QCN), phát triển năng lực số (NLS), lan tỏa tình yêu thương và sự thấu cảm trong học đường.";
                if (week <= 18) return "Tích hợp giáo dục bảo vệ môi trường, tài nguyên thiên nhiên; Quốc phòng và An ninh (QPAN): Ý thức bảo vệ sự bình yên quê hương.";
                if (week <= 27) return "Tích hợp giáo dục bình đẳng giới, kỹ năng tự nhận thức và giao tiếp hợp tác; phòng chống bạo lực học đường.";
                return "Tích hợp QPAN: Lòng tự hào về truyền thống anh hùng của dân tộc, chủ quyền biên giới và biển đảo Việt Nam.";
            }
        }

        // 2. TOÁN (Khối 1, 2, 3, 4)
        if (norm === "Toán") {
            if (lName.includes("thực hành") || lName.includes("trải nghiệm") || lName.includes("stem")) {
                return "Tích hợp Bài học STEM: Vận dụng kiến thức toán học vào thiết kế, chế tạo mô hình và giải quyết vấn đề thực tế.";
            }
            if (lName.includes("hình") || lName.includes("góc") || lName.includes("đo") || lName.includes("xăng-ti-mét") || lName.includes("mét") || lName.includes("diện tích")) {
                return "Tích hợp năng lực hình học và đo đạc thực tế: Quan sát đồ vật xung quanh, ước lượng và đo đạc trong cuộc sống hàng ngày.";
            }
            if (lName.includes("tiền") || lName.includes("thời gian") || lName.includes("giờ") || lName.includes("đồng hồ") || lName.includes("lịch")) {
                return "Tích hợp kĩ năng tài chính và quản lý thời gian: Nhận biết giá trị thời gian và tiền tệ, rèn thói quen chi tiêu tiết kiệm, đúng giờ.";
            }
            if (grade === 1) return "Tích hợp toán học gắn liền với cuộc sống: Sử dụng đồ dùng học tập trực quan, phát triển tư duy logic và thao tác đếm, tính toán nhanh.";
            if (grade === 2) return "Tích hợp làm quen bài học STEM: Phát triển tư duy logic, ước lượng và liên hệ phép tính với các tình huống mua sắm, chia sẻ hàng ngày.";
            if (grade === 3) return "Tích hợp Bài học STEM, năng lực số (NLS): Sử dụng bảng số liệu, giải bài toán có lời văn gắn với thực tiễn đời sống gia đình, nhà trường.";
            if (grade === 4) return "Tích hợp Bài học STEM, năng lực số (NLS): Thu thập, phân tích số liệu biểu đồ, tính toán vận dụng giải quyết bài toán thực tế.";
        }

        // 3. TỰ NHIÊN VÀ XÃ HỘI (TNXH - Khối 1, 2, 3)
        if (norm === "TNXH" || norm === "Tự nhiên và Xã hội") {
            if (lName.includes("gia đình") || lName.includes("nhà")) {
                return "Tích hợp giáo dục tình cảm gia đình, phòng tránh tai nạn thương tích tại nhà (bỏng, điện giật, vật sắc nhọn).";
            }
            if (lName.includes("trường") || lName.includes("lớp")) {
                return "Tích hợp an toàn trường học, giữ gìn vệ sinh lớp học, phòng chống bạo lực và thương tích trong giờ ra chơi.";
            }
            if (lName.includes("cộng đồng") || lName.includes("giao thông") || lName.includes("đường")) {
                return "Tích hợp Giáo dục An toàn giao thông, tuân thủ biển báo, văn hóa xếp hàng và ứng xử văn minh nơi công cộng.";
            }
            if (lName.includes("thực vật") || lName.includes("động vật") || lName.includes("cây") || lName.includes("con")) {
                return "Tích hợp Giáo dục bảo vệ môi trường, bảo vệ động thực vật, giữ gìn cảnh quan thiên nhiên và cân bằng sinh thái.";
            }
            if (lName.includes("cơ thể") || lName.includes("sức khỏe") || lName.includes("vệ sinh") || lName.includes("ăn")) {
                return "Tích hợp chăm sóc sức khỏe ban đầu, vệ sinh an toàn thực phẩm, phòng chống dịch bệnh và rèn kĩ năng tự bảo vệ.";
            }
            if (lName.includes("trái đất") || lName.includes("thời tiết") || lName.includes("bầu trời") || lName.includes("nước") || lName.includes("không khí")) {
                return "Tích hợp ứng phó biến đổi khí hậu, tiết kiệm nước sạch và năng lượng, phòng chống ô nhiễm môi trường.";
            }
            return "Tích hợp giáo dục bảo vệ môi trường, an toàn cuộc sống, chăm sóc sức khỏe và phát triển kĩ năng tự nhận thức.";
        }

        // 4. HOẠT ĐỘNG TRẢI NGHIỆM (Khối 1, 2, 3, 4)
        if (norm === "HĐ Trải nghiệm") {
            if (periodInWeek === 1 || lName.includes("dưới cờ") || lName.includes("khai giảng")) {
                return "Tích hợp Lý tưởng cách mạng, đạo đức lối sống: Nghi thức Đội - Sao nhi đồng, giáo dục truyền thống yêu nước và lòng tự hào dân tộc.";
            }
            if (periodInWeek === 3 || lName.includes("sinh hoạt lớp")) {
                return "Tích hợp kĩ năng tự đánh giá, xây dựng tinh thần hợp tác, trách nhiệm với tập thể và kỹ năng giải quyết mâu thuẫn học đường.";
            }
            return "Tích hợp rèn luyện kĩ năng sống, tự chủ, gắn kết tình bạn, bảo vệ môi trường xanh - sạch - đẹp trong và ngoài nhà trường.";
        }

        // 5. ĐẠO ĐỨC (Khối 1, 2, 3, 4)
        if (norm === "Đạo đức") {
            if (lName.includes("bác hồ")) {
                return "Tích hợp Học tập và làm theo tư tưởng, đạo đức, phong cách Hồ Chí Minh: Bác Hồ với thiếu nhi, rèn luyện 5 điều Bác dạy.";
            }
            if (lName.includes("quê hương") || lName.includes("đất nước") || lName.includes("truyền thống")) {
                return "Tích hợp Giáo dục Quốc phòng và An ninh (QPAN), tình yêu quê hương đất nước, biết ơn người có công.";
            }
            return "Tích hợp Quyền và bổn phận trẻ em, giáo dục lòng nhân ái, trung thực, trách nhiệm với bản thân, gia đình và cộng đồng.";
        }

        // 6. KHOA HỌC (Khối 4)
        if (norm === "Khoa học") {
            return "Tích hợp Giáo dục bảo vệ môi trường, sử dụng năng lượng tiết kiệm và hiệu quả, giữ gìn nguồn nước sạch và phòng tránh tai nạn.";
        }

        // 7. LỊCH SỬ VÀ ĐỊA LÍ (Khối 4)
        if (norm === "LS&ĐL") {
            return "Tích hợp Giáo dục Quốc phòng và An ninh (QPAN), chủ quyền lãnh thổ, biên giới quốc gia và bảo tồn di sản văn hóa dân tộc.";
        }

        // 8. TIN HỌC / CÔNG NGHỆ / CD SỐ (Khối 1..4)
        if (norm === "Tin học" || norm === "Công nghệ" || norm === "CD Số") {
            return "Tích hợp Năng lực số (NLS), giáo dục STEM, văn hóa ứng xử trên không gian mạng và bảo vệ an toàn thông tin cá nhân.";
        }

        // 9. STEM
        if (norm === "STEM") {
            return "Tích hợp Giáo dục STEM: Khơi dậy niềm đam mê sáng tạo, tư duy giải quyết vấn đề và năng lực hợp tác.";
        }

        // 10. KNS
        if (norm === "KNS") {
            return "Tích hợp Kĩ năng sống: Rèn luyện kỹ năng tự phục vụ, ứng phó với các tình huống nguy hiểm và giao tiếp văn minh.";
        }

        return "";
    }

    function buildKhdhForGrade(grade, ppctList) {
        const khdhObj = {};
        if (!ppctList || !Array.isArray(ppctList)) return khdhObj;

        const subGroups = {};
        ppctList.forEach(item => {
            const norm = normalizeSubjectName(item.subject);
            let kKey = norm.toUpperCase();
            if (norm === "TNXH" || norm === "Tự nhiên và Xã hội") kKey = "TỰ NHIÊN VÀ XÃ HỘI";
            if (!subGroups[kKey]) subGroups[kKey] = [];
            subGroups[kKey].push(item);
        });

        Object.keys(subGroups).forEach(kKey => {
            khdhObj[kKey] = subGroups[kKey].map(p => {
                const autoInteg = (p.integration && p.integration.trim()) 
                    ? p.integration.trim() 
                    : getGradeSpecificIntegration(grade, p.subject, p.week, p.periodInWeek, p.lessonName);
                return {
                    week: p.week,
                    topic: p.subject,
                    lesson: p.lessonName,
                    content: p.lessonName,
                    duration: p.duration || "1 tiết",
                    integration: autoInteg,
                    notes: ""
                };
            });
        });

        return khdhObj;
    }

    function getIntegrationFromKhdh(subject, week, periodInWeek) {
        const normSub = normalizeSubjectName(subject);
        if (!normSub || normSub === "-- Nghỉ / Để trống --") return "";

        // 1. Grade 5: use Grade 5's rich KHDH
        if (state.currentGrade === 5) {
            if (!state.khdh) return "";
            let khdhKey = null;
            if (normSub === "Toán") khdhKey = "TOÁN";
            else if (normSub === "Tiếng Việt") khdhKey = "TIẾNG VIỆT";
            else if (normSub === "Khoa học") khdhKey = "KHOA HỌC";
            else if (normSub === "LS&ĐL") khdhKey = "LS&ĐL";
            else if (normSub === "Đạo đức") khdhKey = "ĐẠO ĐỨC";
            else if (normSub === "Công nghệ") khdhKey = "CÔNG NGHỆ";
            else if (normSub === "HĐ Trải nghiệm") khdhKey = "HĐ TRẢI NGHIỆM";

            if (khdhKey && state.khdh[khdhKey]) {
                const weekRows = state.khdh[khdhKey].filter(r => r.week === week);
                if (weekRows.length > 0) {
                    let matchedRow = null;
                    const pInW = periodInWeek || 1;
                    if (khdhKey === "TOÁN" || khdhKey === "KHOA HỌC" || khdhKey === "LS&ĐL" || khdhKey === "HĐ TRẢI NGHIỆM") {
                        matchedRow = weekRows[Math.min(pInW - 1, weekRows.length - 1)];
                    } else if (khdhKey === "ĐẠO ĐỨC" || khdhKey === "CÔNG NGHỆ") {
                        matchedRow = weekRows[0];
                    } else if (khdhKey === "TIẾNG VIỆT") {
                        if (weekRows.length === 6) {
                            if (pInW === 1) matchedRow = weekRows[0];
                            else if (pInW === 2) matchedRow = weekRows[1];
                            else if (pInW === 3) matchedRow = weekRows[2];
                            else if (pInW === 4 || pInW === 5) matchedRow = weekRows[3];
                            else if (pInW === 6) matchedRow = weekRows[4];
                            else if (pInW === 7) matchedRow = weekRows[5];
                        } else {
                            const mappedIdx = Math.min(Math.floor((pInW - 1) * weekRows.length / 7), weekRows.length - 1);
                            matchedRow = weekRows[mappedIdx];
                        }
                    }
                    if (matchedRow && matchedRow.integration && matchedRow.integration.trim()) {
                        return matchedRow.integration.trim();
                    }
                }
            }
        }

        // 2. Grades 1 - 4: Check state.khdh
        let subKey = normSub.toUpperCase();
        if (normSub === "TNXH" || normSub === "Tự nhiên và Xã hội") subKey = "TỰ NHIÊN VÀ XÃ HỘI";
        if (state.khdh && state.khdh[subKey]) {
            const weekRows = state.khdh[subKey].filter(r => r.week === week);
            if (weekRows.length > 0) {
                const matched = weekRows[Math.min((periodInWeek || 1) - 1, weekRows.length - 1)];
                if (matched && matched.integration && matched.integration.trim()) {
                    return matched.integration.trim();
                }
            }
        }

        // 3. Fallback: Grade-specific automatic integration matching the actual lesson name
        const ppctItem = findPpctItem(week, subject, periodInWeek);
        const lName = ppctItem ? ppctItem.lessonName : "";
        return getGradeSpecificIntegration(state.currentGrade || 1, subject, week, periodInWeek, lName);
    }

    function findPpctItem(week, subject, periodInWeek) {
        const normSub = normalizeSubjectName(subject);
        if (!normSub || normSub === "-- Nghỉ / Để trống --") return null;
        return state.ppct.find(p => p.week === week && normalizeSubjectName(p.subject) === normSub && p.periodInWeek === periodInWeek);
    }

    function getWeekSlots(weekNum) {
        if (state.weeklyCustomSlots && state.weeklyCustomSlots[weekNum] && Array.isArray(state.weeklyCustomSlots[weekNum]) && state.weeklyCustomSlots[weekNum].length > 0) {
            return JSON.parse(JSON.stringify(state.weeklyCustomSlots[weekNum]));
        }
        return JSON.parse(JSON.stringify(state.timetable));
    }

    function addSlotToWeek(weekNum, day, session, insertAfterPeriod, defaultSub) {
        if (!state.weeklyCustomSlots) state.weeklyCustomSlots = {};
        let slots = getWeekSlots(weekNum);
        
        let targetIdx = -1;
        for (let i = 0; i < slots.length; i++) {
            if (slots[i].day === day && slots[i].session === session && slots[i].period === insertAfterPeriod) {
                targetIdx = i;
                break;
            }
        }

        const newId = `slot_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`;
        const newSlot = {
            day: day,
            session: session,
            period: (insertAfterPeriod || 0) + 1,
            subject: defaultSub || "Tiếng Việt",
            id: newId
        };

        if (targetIdx >= 0) {
            slots.splice(targetIdx + 1, 0, newSlot);
        } else {
            let lastIdx = -1;
            for (let i = 0; i < slots.length; i++) {
                if (slots[i].day === day && slots[i].session === session) {
                    lastIdx = i;
                }
            }
            if (lastIdx >= 0) {
                slots.splice(lastIdx + 1, 0, newSlot);
            } else {
                slots.push(newSlot);
            }
        }

        let pCounter = 1;
        slots.forEach(s => {
            if (s.day === day && s.session === session) {
                s.period = pCounter++;
            }
        });

        state.weeklyCustomSlots[weekNum] = slots;
        saveState();
        if (state.currentTab === "tab-lbg") renderTabLbg();
        else if (state.currentTab === "tab-ctlop") renderTabCtlop();
        showToast(`Đã thêm 1 tiết mới vào ${day} (${session}) Tuần ${weekNum}!`, "success");
    }

    function removeSlotFromWeek(weekNum, day, session, periodNum) {
        if (!state.weeklyCustomSlots) state.weeklyCustomSlots = {};
        let slots = getWeekSlots(weekNum);
        
        const targetIdx = slots.findIndex(s => s.day === day && s.session === session && s.period === periodNum);
        if (targetIdx < 0) return;

        slots.splice(targetIdx, 1);

        let pCounter = 1;
        slots.forEach(s => {
            if (s.day === day && s.session === session) {
                s.period = pCounter++;
            }
        });

        state.weeklyCustomSlots[weekNum] = slots;
        saveState();
        if (state.currentTab === "tab-lbg") renderTabLbg();
        else if (state.currentTab === "tab-ctlop") renderTabCtlop();
        showToast(`Đã xóa Tiết ${periodNum} của ${day} (${session}) Tuần ${weekNum}!`, "success");
    }

    function resetWeekSlotsToDefault(weekNum) {
        if (state.weeklyCustomSlots && state.weeklyCustomSlots[weekNum]) {
            delete state.weeklyCustomSlots[weekNum];
        }
        const prefix = `w${weekNum}_`;
        for (let k in state.weeklyScheduleOverrides) {
            if (k.startsWith(prefix)) {
                delete state.weeklyScheduleOverrides[k];
            }
        }
        saveState();
        if (state.currentTab === "tab-lbg") renderTabLbg();
        else if (state.currentTab === "tab-ctlop") renderTabCtlop();
        showToast(`Đã khôi phục Thời khóa biểu gốc cho Tuần ${weekNum}!`, "success");
    }

    function renderMasterTimetableMatrix() {
        const tbody = document.getElementById("master-timetable-matrix-tbody");
        if (!tbody) return;

        const slots = state.timetable || [];
        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];

        let maxMorn = 4;
        let maxAft = 3;
        slots.forEach(s => {
            if (s.session === "Sáng" && s.period > maxMorn) maxMorn = s.period;
            if (s.session === "Chiều" && s.period > maxAft) maxAft = s.period;
        });

        let html = "";

        // Morning rows
        for (let p = 1; p <= maxMorn; p++) {
            html += `<tr>`;
            if (p === 1) {
                html += `<td rowspan="${maxMorn}" style="font-weight: 700; background: #f8fafc; vertical-align: middle; color: var(--primary);">Sáng</td>`;
            }
            html += `<td style="font-weight: 700; background: #fafafa; vertical-align: middle;">${p}</td>`;
            days.forEach(d => {
                const isAltDay = (d === "Thứ 2" || d === "Thứ 4" || d === "Thứ 6");
                const dayMornSlots = slots.filter(s => s.day === d && s.session === "Sáng").sort((a, b) => (a.period || 0) - (b.period || 0));
                const slot = dayMornSlots[p - 1];
                const sub = slot ? slot.subject : "";
                const isOff = sub === "-- Nghỉ / Để trống --" || !sub;
                html += `<td style="vertical-align: middle; font-weight: 600; font-size: 0.88rem; ${isAltDay ? 'background-color: #f0f7ff;' : 'background-color: #ffffff;'} ${isOff ? 'color: #94a3b8; font-style: italic;' : 'color: var(--text-main);'}">${escapeHtml(sub || '--')}</td>`;
            });
            html += `</tr>`;
        }

        // Afternoon rows
        for (let p = 1; p <= maxAft; p++) {
            html += `<tr>`;
            if (p === 1) {
                html += `<td rowspan="${maxAft}" style="font-weight: 700; background: #f8fafc; vertical-align: middle; color: #b45309;">Chiều</td>`;
            }
            html += `<td style="font-weight: 700; background: #fafafa; vertical-align: middle;">${p}</td>`;
            days.forEach(d => {
                const isAltDay = (d === "Thứ 2" || d === "Thứ 4" || d === "Thứ 6");
                const dayAftSlots = slots.filter(s => s.day === d && s.session === "Chiều").sort((a, b) => (a.period || 0) - (b.period || 0));
                const slot = dayAftSlots[p - 1];
                const sub = slot ? slot.subject : "";
                const isOff = sub === "-- Nghỉ / Để trống --" || !sub;
                html += `<td style="vertical-align: middle; font-weight: 600; font-size: 0.88rem; ${isAltDay ? 'background-color: #f0f7ff;' : 'background-color: #ffffff;'} ${isOff ? 'color: #94a3b8; font-style: italic;' : 'color: var(--text-main);'}">${escapeHtml(sub || '--')}</td>`;
            });
            html += `</tr>`;
        }

        tbody.innerHTML = html;
    }

    function renderMasterTimetableEditor() {
        const tbody = document.getElementById("master-timetable-tbody");
        if (!tbody) return;

        tbody.innerHTML = "";
        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];
        const dynamicSubjectOptions = getAllUniqueSubjects().concat(["-- Nghỉ / Để trống --"]);

        days.forEach(day => {
            const isAltDay = (day === "Thứ 2" || day === "Thứ 4" || day === "Thứ 6");
            const daySlots = state.timetable.filter(s => s.day === day);
            const morningSlots = daySlots.filter(s => s.session === "Sáng").sort((a, b) => (a.period || 0) - (b.period || 0));
            const afternoonSlots = daySlots.filter(s => s.session === "Chiều").sort((a, b) => (a.period || 0) - (b.period || 0));
            const orderedDaySlots = [...morningSlots, ...afternoonSlots];
            const totalRows = orderedDaySlots.length;

            orderedDaySlots.forEach((slot, idx) => {
                const globalIdx = state.timetable.indexOf(slot);
                const tr = document.createElement("tr");
                if (isAltDay) tr.classList.add("row-day-alt");

                let dayCellHtml = "";
                if (idx === 0) {
                    dayCellHtml = `<td rowspan="${totalRows}" class="col-day ${isAltDay ? 'day-alt' : 'day-normal'}" style="font-weight:bold; text-align:center;">${day}</td>`;
                }

                let sessionCellHtml = "";
                if (idx === 0) {
                    sessionCellHtml = `<td rowspan="${morningSlots.length}" style="text-align:center; font-weight:600; background:#fafafa;">Sáng</td>`;
                } else if (idx === morningSlots.length) {
                    sessionCellHtml = `<td rowspan="${afternoonSlots.length}" style="text-align:center; font-weight:600; background:#fafafa;">Chiều</td>`;
                }

                tr.innerHTML = `
                    ${dayCellHtml}
                    ${sessionCellHtml}
                    <td style="text-align:center; font-weight:bold;">${slot.period}</td>
                    <td>
                        <select class="form-select form-select-sm master-tt-subject-sel" data-index="${globalIdx}">
                            ${dynamicSubjectOptions.map(s => `<option value="${s}" ${normalizeSubjectName(s) === normalizeSubjectName(slot.subject) || s === slot.subject ? 'selected' : ''}>${s}</option>`).join('')}
                        </select>
                    </td>
                    <td style="text-align:center;">
                        <div class="row-actions-group">
                            <button type="button" class="btn-row-action btn-master-add-slot" data-day="${slot.day}" data-session="${slot.session}" data-period="${slot.period}" title="Chèn thêm 1 tiết ngay dưới dòng này">➕</button>
                            <button type="button" class="btn-row-action btn-master-del-slot" data-index="${globalIdx}" title="Xóa tiết này khỏi TKB gốc">🗑️</button>
                        </div>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        });

        renderMasterTimetableMatrix();

        tbody.querySelectorAll(".master-tt-subject-sel").forEach(sel => {
            sel.addEventListener("change", (e) => {
                const idx = parseInt(e.target.dataset.index);
                if (idx >= 0 && idx < state.timetable.length) {
                    state.timetable[idx].subject = normalizeSubjectName(e.target.value);
                    saveState();
                    renderMasterTimetableMatrix();
                    renderTabLbg();
                    renderTabCtlop();
                }
            });
        });

        tbody.querySelectorAll(".btn-master-add-slot").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const day = e.currentTarget.dataset.day;
                const session = e.currentTarget.dataset.session;
                const period = parseInt(e.currentTarget.dataset.period) || 1;
                addSlotToMasterTimetable(day, session, period);
            });
        });

        tbody.querySelectorAll(".btn-master-del-slot").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const idx = parseInt(e.currentTarget.dataset.index);
                removeSlotFromMasterTimetable(idx);
            });
        });
    }

    function addSlotToMasterTimetable(day, session, insertAfterPeriod) {
        let targetIdx = -1;
        for (let i = 0; i < state.timetable.length; i++) {
            if (state.timetable[i].day === day && state.timetable[i].session === session && state.timetable[i].period === insertAfterPeriod) {
                targetIdx = i;
                break;
            }
        }

        const newSlot = {
            day: day,
            session: session,
            period: (insertAfterPeriod || 0) + 1,
            subject: "Tiếng Việt"
        };

        if (targetIdx >= 0) {
            state.timetable.splice(targetIdx + 1, 0, newSlot);
        } else {
            state.timetable.push(newSlot);
        }

        let pCounter = 1;
        state.timetable.forEach(s => {
            if (s.day === day && s.session === session) {
                s.period = pCounter++;
            }
        });

        saveState();
        renderMasterTimetableEditor();
        renderTabLbg();
        renderTabCtlop();
        showToast(`Đã thêm 1 tiết mới vào ${day} (${session}) và tự động đồng bộ sang Lịch Báo Giảng!`, "success");
    }

    function removeSlotFromMasterTimetable(idx) {
        if (idx < 0 || idx >= state.timetable.length) return;
        const removed = state.timetable[idx];
        state.timetable.splice(idx, 1);

        let pCounter = 1;
        state.timetable.forEach(s => {
            if (s.day === removed.day && s.session === removed.session) {
                s.period = pCounter++;
            }
        });

        saveState();
        renderMasterTimetableEditor();
        renderTabLbg();
        renderTabCtlop();
        showToast(`Đã xóa Tiết ${removed.period} của ${removed.day} (${removed.session}) và tự động đồng bộ sang Lịch Báo Giảng!`, "success");
    }

    // Render HTML for Timetable Paper (Landscape Preview & Print)
    function renderTimetablePaperHtml(slots, customTitle, customSubtitle) {
        const s = state.settings;
        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];

        let maxMorn = 4;
        let maxAft = 3;
        slots.forEach(slot => {
            if (slot.session === "Sáng" && slot.period > maxMorn) maxMorn = slot.period;
            if (slot.session === "Chiều" && slot.period > maxAft) maxAft = slot.period;
        });

        const bghSigner = getBghSignerInfo("LBG");

        let html = `
        <div class="paper-page paper-landscape">
            <div class="paper-header">
                <table class="paper-header-table" style="width: 100%;">
                    <tr>
                        <td style="width: 50%; text-align: center;">
                            <div style="font-size: 11pt; text-transform: uppercase;">${escapeHtml((s.governingBody || 'UBND PHƯỜNG TRUNG NHỨT').toUpperCase())}</div>
                            <div style="font-size: 11.5pt; font-weight: bold; text-transform: uppercase;">${escapeHtml((s.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT').toUpperCase())}</div>
                            <div style="font-size: 11.5pt; font-weight: bold; text-transform: uppercase;">${escapeHtml((s.grade || 'KHỐI 5').toUpperCase())} - ${escapeHtml((s.className || 'LỚP 5A').toUpperCase())}</div>
                        </td>
                        <td style="width: 50%; text-align: center;">
                            <div style="font-size: 11pt; font-weight: bold;">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
                            <div style="font-size: 12pt; font-weight: bold; text-decoration: underline;">Độc lập - Tự do - Hạnh phúc</div>
                        </td>
                    </tr>
                </table>
                <div class="paper-title" style="margin-top: 15px; font-size: 16pt;">${escapeHtml(customTitle || ('THỜI KHÓA BIỂU ' + (s.grade || 'KHỐI 5') + ' - ' + (s.className || 'LỚP 5A')).toUpperCase())}</div>
                <div class="paper-subtitle" style="font-size: 11.5pt;">${escapeHtml(customSubtitle || ('Năm học ' + (s.academicYear || '2026 - 2027')))}</div>
            </div>

            <table class="paper-table" style="width: 100%; text-align: center;">
                <thead>
                    <tr>
                        <th style="width: 10%;">Buổi</th>
                        <th style="width: 8%;">Tiết</th>
                        <th style="width: 16.4%; background-color: #f1f5f9;">Thứ 2</th>
                        <th style="width: 16.4%;">Thứ 3</th>
                        <th style="width: 16.4%; background-color: #f1f5f9;">Thứ 4</th>
                        <th style="width: 16.4%;">Thứ 5</th>
                        <th style="width: 16.4%; background-color: #f1f5f9;">Thứ 6</th>
                    </tr>
                </thead>
                <tbody>
        `;

        // Morning
        for (let p = 1; p <= maxMorn; p++) {
            html += `<tr>`;
            if (p === 1) {
                html += `<td rowspan="${maxMorn}" style="font-weight: bold; vertical-align: middle; background: #fbfbfb;">Sáng</td>`;
            }
            html += `<td style="font-weight: bold; vertical-align: middle;">${p}</td>`;
            days.forEach(d => {
                const isAltDay = (d === "Thứ 2" || d === "Thứ 4" || d === "Thứ 6");
                const dayMornSlots = slots.filter(item => item.day === d && item.session === "Sáng").sort((a, b) => (a.period || 0) - (b.period || 0));
                const slot = dayMornSlots[p - 1];
                const sub = slot ? slot.subject : "";
                const isOff = sub === "-- Nghỉ / Để trống --" || !sub;
                html += `<td style="vertical-align: middle; font-weight: 600; padding: 6px 4px; ${isAltDay ? 'background-color: #f8fafc;' : ''} ${isOff ? 'color: #94a3b8; font-style: italic;' : ''}">${escapeHtml(isOff ? '' : sub)}</td>`;
            });
            html += `</tr>`;
        }

        // Afternoon
        for (let p = 1; p <= maxAft; p++) {
            html += `<tr>`;
            if (p === 1) {
                html += `<td rowspan="${maxAft}" style="font-weight: bold; vertical-align: middle; background: #fbfbfb;">Chiều</td>`;
            }
            html += `<td style="font-weight: bold; vertical-align: middle;">${p}</td>`;
            days.forEach(d => {
                const isAltDay = (d === "Thứ 2" || d === "Thứ 4" || d === "Thứ 6");
                const dayAftSlots = slots.filter(item => item.day === d && item.session === "Chiều").sort((a, b) => (a.period || 0) - (b.period || 0));
                const slot = dayAftSlots[p - 1];
                const sub = slot ? slot.subject : "";
                const isOff = sub === "-- Nghỉ / Để trống --" || !sub;
                html += `<td style="vertical-align: middle; font-weight: 600; padding: 6px 4px; ${isAltDay ? 'background-color: #f8fafc;' : ''} ${isOff ? 'color: #94a3b8; font-style: italic;' : ''}">${escapeHtml(isOff ? '' : sub)}</td>`;
            });
            html += `</tr>`;
        }

        html += `
                </tbody>
            </table>

            <div class="paper-footer" style="margin-top: 25px;">
                <table class="paper-footer-table" style="width: 100%;">
                    <tr>
                        <td style="width: 33.3%; text-align: center;">
                            <div style="font-size: 11.5pt; font-weight: bold;">DUYỆT CỦA ${escapeHtml(bghSigner.title)}</div>
                            <div style="font-size: 10.5pt; font-style: italic;">(Ký và ghi rõ họ tên)</div>
                            <div style="height: 60px;"></div>
                            <div style="font-size: 11.5pt; font-weight: bold;">${escapeHtml(bghSigner.name)}</div>
                        </td>
                        <td style="width: 33.3%; text-align: center;">
                            <div style="font-size: 11.5pt; font-weight: bold;">TỔ TRƯỞNG CHUYÊN MÔN</div>
                            <div style="font-size: 10.5pt; font-style: italic;">(Ký và ghi rõ họ tên)</div>
                            <div style="height: 60px;"></div>
                            <div style="font-size: 11.5pt; font-weight: bold;">${escapeHtml(s.headOfGrade || 'Trần Thị Mai')}</div>
                        </td>
                        <td style="width: 33.3%; text-align: center;">
                            <div style="font-size: 11.5pt; font-weight: bold;">GIÁO VIÊN CHỦ NHIỆM</div>
                            <div style="font-size: 10.5pt; font-style: italic;">(Ký và ghi rõ họ tên)</div>
                            <div style="height: 60px;"></div>
                            <div style="font-size: 11.5pt; font-weight: bold;">${escapeHtml(s.homeroomTeacher || 'Nguyễn Thị Thu Hà')}</div>
                        </td>
                    </tr>
                </table>
            </div>
        </div>
        `;
        return html;
    }

    function exportMasterTimetableDocx() {
        if (window.DocxGenerator && window.DocxGenerator.generateTimetableDocx) {
            showToast("Đang tạo file Word Thời khóa biểu (khổ ngang)...", "info");
            window.DocxGenerator.generateTimetableDocx(state.timetable, state.settings).then(blob => {
                const filename = `Thoi_Khoa_Bieu_${(state.settings.className || 'Lop_5A').replace(/\s+/g, '_')}_${(state.settings.academicYear || '2026_2027').replace(/\s+/g, '')}.docx`;
                saveAs(blob, filename);
                showToast(`Đã xuất file Word Thời khóa biểu: ${filename}`, "success");
            }).catch(err => {
                console.error("Timetable DOCX error:", err);
                alert("Lỗi xuất file Word: " + err.message);
            });
        } else {
            alert("Bộ tạo file Word chưa sẵn sàng!");
        }
    }

    function exportMasterTimetableXlsx() {
        if (window.XlsxGenerator && window.XlsxGenerator.generateTimetableXlsx) {
            window.XlsxGenerator.generateTimetableXlsx(state.timetable, state.settings).then(blob => {
                const filename = `Thoi_Khoa_Bieu_${(state.settings.className || 'Lop_5A').replace(/\s+/g, '_')}_${(state.settings.academicYear || '2026_2027').replace(/\s+/g, '')}.xlsx`;
                saveAs(blob, filename);
                showToast(`Đã xuất file Excel Thời khóa biểu: ${filename}`, "success");
            }).catch(err => {
                console.error("Timetable XLSX error:", err);
                alert("Lỗi xuất file Excel: " + err.message);
            });
        } else {
            alert("Bộ tạo file Excel chưa sẵn sàng!");
        }
    }

    function previewMasterTimetablePrint() {
        openPreviewModal(
            `Xem trước & In Thời Khóa Biểu ${(state.settings.className || 'Lớp 5A')}`,
            renderTimetablePaperHtml(state.timetable),
            exportMasterTimetableDocx,
            exportMasterTimetableXlsx,
            () => window.print()
        );
    }

    function calculateWeekSchedule(weekNum) {
        const rawTimetable = getWeekSlots(weekNum);
        const daysOrder = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];
        const timetable = [];
        daysOrder.forEach(d => {
            const morn = rawTimetable.filter(s => s.day === d && s.session === "Sáng").sort((a, b) => (a.period || 0) - (b.period || 0));
            const aft = rawTimetable.filter(s => s.day === d && s.session === "Chiều").sort((a, b) => (a.period || 0) - (b.period || 0));
            timetable.push(...morn, ...aft);
        });
        const subjectCounts = {};
        const schedule = [];
        
        let gvcnCount = 0;
        let specialistCount = 0;
        let enhancedCount = 0;

        const fridayPeriods = state.weeklyFridayPeriods[weekNum] !== undefined 
            ? state.weeklyFridayPeriods[weekNum] 
            : state.fridayPeriodsDefault;

        timetable.forEach((slot) => {
            const key = slot.id || `w${weekNum}_d${slot.day}_s${slot.session}_p${slot.period}`;
            const override = state.weeklyScheduleOverrides[key] || {};
            
            let rawSubject = override.subject !== undefined ? override.subject : slot.subject;
            
            // Check Friday afternoon period cutoff
            if (slot.day === "Thứ 6" && slot.session === "Chiều") {
                if (slot.period > fridayPeriods) {
                    rawSubject = "-- Nghỉ / Để trống --";
                }
            }

            const canonicalSub = normalizeSubjectName(rawSubject);
            const isOff = canonicalSub === "-- Nghỉ / Để trống --";

            let periodInWeek = "";
            let ppct = "";
            let lessonName = "";
            let integration = "";

            const customColsData = (override.customCols && typeof override.customCols === 'object') ? { ...override.customCols } : {};
            if (override.customCol !== undefined && customColsData['col_1'] === undefined) {
                customColsData['col_1'] = override.customCol;
            }

            if (!isOff) {
                const isIncluded = !state.includedSubjects || state.includedSubjects.length === 0 || 
                    state.includedSubjects.some(s => s === canonicalSub || normalizeSubjectName(s) === canonicalSub);

                if (!isIncluded) {
                    // Môn không chọn đưa vào LBG: Vẫn hiển thị tên môn, để trống Tiết PPCT và Tên bài dạy
                    schedule.push({
                        day: slot.day,
                        session: slot.session,
                        period: slot.period,
                        subject: canonicalSub,
                        isOff: false,
                        periodInWeek: "",
                        ppct: "",
                        lessonName: override.lessonName !== undefined ? override.lessonName : "",
                        integration: override.integration !== undefined ? override.integration : "",
                        note: override.note !== undefined ? override.note : "",
                        customCol: override.customCol !== undefined ? override.customCol : (customColsData['col_1'] !== undefined ? customColsData['col_1'] : (override.note !== undefined ? override.note : "")),
                        customCols: customColsData,
                        key: key
                    });
                    return;
                }

                if (!subjectCounts[canonicalSub]) subjectCounts[canonicalSub] = 0;
                subjectCounts[canonicalSub]++;
                periodInWeek = subjectCounts[canonicalSub];

                const ppctItem = findPpctItem(weekNum, canonicalSub, periodInWeek);
                ppct = override.ppct !== undefined ? override.ppct : (ppctItem ? ppctItem.ppct : "");
                lessonName = override.lessonName !== undefined ? override.lessonName : (ppctItem ? ppctItem.lessonName : "");
                
                // Strict 100% Pedagogical Consistency: Override -> PPCT Item Integration -> Empty
                if (override.integration !== undefined) {
                    integration = override.integration;
                } else if (ppctItem && ppctItem.integration && ppctItem.integration.trim()) {
                    integration = ppctItem.integration.trim();
                } else {
                    integration = "";
                }

                const sUpper = canonicalSub.toUpperCase();
                if (["TIẾNG VIỆT", "TOÁN", "ĐẠO ĐỨC", "HĐ TRẢI NGHIỆM", "KHOA HỌC", "LS&ĐL"].includes(sUpper)) {
                    gvcnCount++;
                } else if (["TIẾNG ANH", "TIN HỌC", "CÔNG NGHỆ", "GD THỂ CHẤT", "ÂM NHẠC", "MĨ THUẬT"].includes(sUpper)) {
                    specialistCount++;
                } else {
                    enhancedCount++;
                }
            }

            schedule.push({
                day: slot.day,
                session: slot.session,
                period: slot.period,
                subject: canonicalSub,
                isOff: isOff,
                periodInWeek: periodInWeek,
                ppct: ppct,
                lessonName: lessonName,
                integration: integration,
                note: override.note !== undefined ? override.note : "",
                customCol: override.customCol !== undefined ? override.customCol : (customColsData['col_1'] !== undefined ? customColsData['col_1'] : (override.note !== undefined ? override.note : "")),
                customCols: customColsData,
                key: key
            });
        });

        const weekInfo = state.weeks.find(w => w.week === weekNum) || {
            week: weekNum,
            startDateVN: "07/09/2026",
            endDateVN: "11/09/2026",
            month: "Tháng 9 năm 2026"
        };

        return {
            weekInfo,
            schedule,
            fridayPeriods,
            stats: { gvcnCount, specialistCount, enhancedCount, total: (gvcnCount + specialistCount + enhancedCount) }
        };
    }

    // Helper: Convert week startDateVN (dd/mm/yyyy) + day name to "Thứ X (dd/mm)"
    function getDayDateStr(startDateVN, dayStr) {
        if (!startDateVN) return dayStr;
        const parts = startDateVN.split("/");
        if (parts.length < 3) return dayStr;
        const d = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10) - 1;
        const y = parseInt(parts[2], 10);
        const dayOffsets = { "Thứ 2": 0, "Thứ 3": 1, "Thứ 4": 2, "Thứ 5": 3, "Thứ 6": 4, "Thứ 7": 5, "Chủ nhật": 6 };
        const offset = dayOffsets[dayStr] !== undefined ? dayOffsets[dayStr] : 0;
        const slotDate = new Date(y, m, d + offset);
        const dd = String(slotDate.getDate()).padStart(2, '0');
        const mm = String(slotDate.getMonth() + 1).padStart(2, '0');
        return `${dayStr} (${dd}/${mm})`;
    }

    // Helper: Calculate exact date (dd/mm/yyyy) for each day of week based on week startDateVN
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

    // VERSION 6.0: Calculate schedule grouped by subject
    function calculateWeekScheduleBySubject(weekNum, filterSubject = "all", filterCategory = "all") {
        const { weekInfo, schedule, stats } = calculateWeekSchedule(weekNum);
        const activeSlots = schedule.filter(s => !s.isOff);

        // Order subjects according to state.subjectList
        const definedOrder = (state.subjectList || []).map(s => normalizeSubjectName(s.name));
        const presentSubjects = [];
        const seen = new Set();

        definedOrder.forEach(subName => {
            if (!seen.has(subName) && activeSlots.some(s => normalizeSubjectName(s.subject) === subName)) {
                seen.add(subName);
                presentSubjects.push(subName);
            }
        });

        activeSlots.forEach(s => {
            const norm = normalizeSubjectName(s.subject);
            if (!seen.has(norm)) {
                seen.add(norm);
                presentSubjects.push(norm);
            }
        });

        const dayOrder = { "Thứ 2": 1, "Thứ 3": 2, "Thứ 4": 3, "Thứ 5": 4, "Thứ 6": 5 };
        const sessOrder = { "Sáng": 1, "Chiều": 2 };

        const subjectGroups = [];
        let totalFilteredPeriods = 0;
        let gvcnFilteredCount = 0;
        let specialistFilteredCount = 0;

        presentSubjects.forEach(subName => {
            const subMeta = (state.subjectList || []).find(s => normalizeSubjectName(s.name) === subName) || {
                name: subName,
                category: "GVCN"
            };

            if (filterCategory && filterCategory !== "all") {
                if (subMeta.category !== filterCategory) return;
            }

            if (filterSubject && filterSubject !== "all") {
                if (normalizeSubjectName(filterSubject) !== subName) return;
            }

            const slots = activeSlots.filter(s => normalizeSubjectName(s.subject) === subName);
            slots.sort((a, b) => {
                const dA = dayOrder[a.day] || 99;
                const dB = dayOrder[b.day] || 99;
                if (dA !== dB) return dA - dB;
                const sA = sessOrder[a.session] || 99;
                const sB = sessOrder[b.session] || 99;
                if (sA !== sB) return sA - sB;
                return (a.period || 0) - (b.period || 0);
            });

            slots.forEach((slot, idx) => {
                slot.periodInWeek = idx + 1;
            });

            if (slots.length > 0) {
                totalFilteredPeriods += slots.length;
                if (subMeta.category === "GVCN") gvcnFilteredCount += slots.length;
                else if (subMeta.category === "Chuyên trách") specialistFilteredCount += slots.length;

                subjectGroups.push({
                    subjectName: subMeta.name || subName,
                    canonicalName: subName,
                    category: subMeta.category || "GVCN",
                    totalPeriods: slots.length,
                    slots: slots
                });
            }
        });

        return {
            weekInfo,
            subjectGroups,
            stats: {
                totalSubjects: subjectGroups.length,
                totalPeriods: totalFilteredPeriods,
                gvcnCount: gvcnFilteredCount,
                specialistCount: specialistFilteredCount,
                originalStats: stats
            }
        };
    }

    // 3. Renderers
    function renderWeekToolbar(containerId, onChangeCallback) {
        const container = document.getElementById(containerId);
        if (!container) return;

        const current = state.currentWeek;
        const weekInfo = state.weeks.find(w => w.week === current);
        const dateStr = weekInfo ? `Từ ngày ${weekInfo.startDateVN} đến ngày ${weekInfo.endDateVN}` : "";

        let optionsHtml = "";
        for (let i = 1; i <= 35; i++) {
            const term = i <= 18 ? "HK I" : "HK II";
            optionsHtml += `<option value="${i}" ${i === current ? 'selected' : ''}>Tuần ${i} (${term})</option>`;
        }

        container.innerHTML = `
            <div class="week-nav">
                <button class="btn btn-secondary btn-sm" id="${containerId}-prev" ${current <= 1 ? 'disabled' : ''}>◀ Tuần trước</button>
                <select class="form-select" id="${containerId}-select" style="font-weight: bold; color: var(--primary);">
                    ${optionsHtml}
                </select>
                <button class="btn btn-secondary btn-sm" id="${containerId}-next" ${current >= 35 ? 'disabled' : ''}>Tuần sau ▶</button>
            </div>
            <div class="week-date-info">
                📅 <strong>TUẦN ${current}</strong>: ${dateStr}
            </div>
        `;

        document.getElementById(`${containerId}-prev`).addEventListener("click", () => {
            if (state.currentWeek > 1) {
                state.currentWeek--;
                saveState();
                onChangeCallback();
            }
        });

        document.getElementById(`${containerId}-next`).addEventListener("click", () => {
            if (state.currentWeek < 35) {
                state.currentWeek++;
                saveState();
                onChangeCallback();
            }
        });

        document.getElementById(`${containerId}-select`).addEventListener("change", (e) => {
            state.currentWeek = parseInt(e.target.value);
            saveState();
            onChangeCallback();
        });
    }

    // Helper: Determine ordered columns for LBG table and export
    function getLbgOrderedColumns(customColsList, showSign, showNote, isCtlop) {
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
                    pos: c.pos
                });
            });
        };

        // Pos 1: before day (Cột đầu tiên)
        pushCustom('1');

        // Day (Thứ, ngày)
        cols.push({ key: 'day', title: 'Thứ, ngày', isCustom: false });

        // Pos 2: after day
        pushCustom('2');

        // Session (Buổi)
        cols.push({ key: 'session', title: 'Buổi', isCustom: false });

        // Pos 3: after session
        pushCustom('3');

        // Period (Tiết)
        cols.push({ key: 'period', title: 'Tiết', isCustom: false });

        // Pos 4: after period
        pushCustom('4');

        // Subject (Môn học)
        cols.push({ key: 'subject', title: 'Môn học', isCustom: false });

        // Pos 5: after subject
        pushCustom('5');

        // PPCT (Tiết PPCT)
        cols.push({ key: 'ppct', title: 'Tiết PPCT', isCustom: false });

        // Pos 6: after ppct
        pushCustom('6');

        // Lesson (Tên bài dạy)
        cols.push({ key: 'lesson', title: 'Tên bài dạy', isCustom: false });

        // Pos 7: after lesson
        pushCustom('7');

        if (isCtlop) {
            cols.push({ key: 'integ', title: 'Nội dung tích hợp / Điều chỉnh', isCustom: false });
        }
        if (showSign) {
            cols.push({ key: 'sign', title: 'Kí tên', isCustom: false });
        }
        if (showNote) {
            cols.push({ key: 'note', title: 'Ghi chú', isCustom: false });
        }

        // Pos end: at the end of the table
        enabledCustomCols.filter(c => !['1', '2', '3', '4', '5', '6', '7'].includes(String(c.pos))).forEach(c => {
            cols.push({
                key: 'custom_' + c.id,
                id: c.id,
                title: c.name || 'Cột mới',
                isCustom: true,
                pos: c.pos || 'end'
            });
        });

        return cols;
    }

    // Helper: Render custom columns manager chips in toolbar or preview modal
    function renderCustomColsManager(containerId, isModal = false) {
        const container = document.getElementById(containerId);
        if (!container) return;
        container.innerHTML = "";

        if (!Array.isArray(state.lbgCustomCols)) state.lbgCustomCols = [];

        state.lbgCustomCols.forEach(col => {
            const chip = document.createElement("div");
            chip.className = "custom-col-chip";
            chip.setAttribute("data-id", col.id);

            const chk = document.createElement("input");
            chk.type = "checkbox";
            chk.className = "chk-custom-col-toggle";
            chk.checked = (col.enabled !== false);
            chk.title = "Bật/Tắt hiển thị cột này";
            chk.onchange = (e) => {
                col.enabled = e.target.checked;
                saveState();
                renderCustomColsManager("lbg-custom-cols-list", false);
                renderCustomColsManager("ctlop-custom-cols-list", false);
                renderCustomColsManager("modal-custom-cols-list", true);
                renderTabLbg();
                renderTabCtlop();
                if (currentPreviewRefreshFn) {
                    const body = document.getElementById("modal-preview-body");
                    if (body) body.innerHTML = currentPreviewRefreshFn();
                }
            };

            const txt = document.createElement("input");
            txt.type = "text";
            txt.className = "txt-custom-col-name";
            txt.value = col.name || "Ghi chú";
            txt.placeholder = "Tên cột";
            txt.title = "Nhập tên cột tuỳ chỉnh";
            txt.oninput = (e) => {
                col.name = e.target.value;
                saveState();
                ["lbg-custom-cols-list", "ctlop-custom-cols-list", "modal-custom-cols-list"].forEach(listId => {
                    const otherInput = document.querySelector(`#${listId} .custom-col-chip[data-id="${col.id}"] .txt-custom-col-name`);
                    if (otherInput && otherInput.value !== e.target.value) otherInput.value = e.target.value;
                });
                const theadCell1 = document.querySelector(`#tab-lbg .col-custom-${col.id}`);
                if (theadCell1) theadCell1.innerText = col.name || "Cột mới";
                const theadCell2 = document.querySelector(`#tab-ctlop .col-custom-${col.id}`);
                if (theadCell2) theadCell2.innerText = col.name || "Cột mới";
                if (currentPreviewRefreshFn) {
                    const body = document.getElementById("modal-preview-body");
                    if (body) body.innerHTML = currentPreviewRefreshFn();
                }
            };
            txt.onblur = () => {
                if (!col.name || !col.name.trim()) col.name = "Cột mới";
                saveState();
                renderTabLbg();
                renderTabCtlop();
                if (currentPreviewRefreshFn) {
                    const body = document.getElementById("modal-preview-body");
                    if (body) body.innerHTML = currentPreviewRefreshFn();
                }
            };

            const sel = document.createElement("select");
            sel.className = "sel-custom-col-pos";
            sel.title = "Chọn vị trí đặt cột trong bảng";
            const posOptions = [
                { val: "end", label: "Cuối bảng" },
                { val: "1", label: "Vị trí 1 (Đầu tiên)" },
                { val: "2", label: "Vị trí 2 (Sau Thứ)" },
                { val: "3", label: "Vị trí 3 (Sau Buổi)" },
                { val: "4", label: "Vị trí 4 (Sau Tiết)" },
                { val: "5", label: "Vị trí 5 (Sau Môn)" },
                { val: "6", label: "Vị trí 6 (Sau PPCT)" },
                { val: "7", label: "Vị trí 7 (Sau Bài)" }
            ];
            posOptions.forEach(opt => {
                const optEl = document.createElement("option");
                optEl.value = opt.val;
                optEl.innerText = opt.label;
                if (String(col.pos || "end") === opt.val) optEl.selected = true;
                sel.appendChild(optEl);
            });
            sel.onchange = (e) => {
                col.pos = e.target.value;
                saveState();
                renderCustomColsManager("lbg-custom-cols-list", false);
                renderCustomColsManager("ctlop-custom-cols-list", false);
                renderCustomColsManager("modal-custom-cols-list", true);
                renderTabLbg();
                renderTabCtlop();
                if (currentPreviewRefreshFn) {
                    const body = document.getElementById("modal-preview-body");
                    if (body) body.innerHTML = currentPreviewRefreshFn();
                }
            };

            const btnDel = document.createElement("button");
            btnDel.type = "button";
            btnDel.className = "btn-del-custom-col";
            btnDel.innerText = "✕";
            btnDel.title = "Xóa cột này";
            btnDel.onclick = () => {
                state.lbgCustomCols = state.lbgCustomCols.filter(c => c.id !== col.id);
                saveState();
                renderCustomColsManager("lbg-custom-cols-list", false);
                renderCustomColsManager("ctlop-custom-cols-list", false);
                renderCustomColsManager("modal-custom-cols-list", true);
                renderTabLbg();
                renderTabCtlop();
                if (currentPreviewRefreshFn) {
                    const body = document.getElementById("modal-preview-body");
                    if (body) body.innerHTML = currentPreviewRefreshFn();
                }
            };

            chip.appendChild(chk);
            chip.appendChild(txt);
            chip.appendChild(sel);
            chip.appendChild(btnDel);
            container.appendChild(chip);
        });
    }

    // Render Tab 1: Lịch Báo Giảng Thường
    function renderTabLbg() {
        renderWeekToolbar("lbg-week-toolbar", () => {
            renderTabLbg();
            if (state.currentTab === "tab-ctlop") renderTabCtlop();
        });

        renderLbgInclusionChips();

        const { weekInfo, schedule, stats, fridayPeriods } = calculateWeekSchedule(state.currentWeek);

        const fridayOptSelect = document.getElementById("lbg-friday-opt");
        if (fridayOptSelect) {
            fridayOptSelect.value = fridayPeriods.toString();
            fridayOptSelect.onchange = (e) => {
                state.weeklyFridayPeriods[state.currentWeek] = parseInt(e.target.value);
                saveState();
                renderTabLbg();
                if (state.currentTab === "tab-ctlop") renderTabCtlop();
            };
        }

        document.getElementById("lbg-gov-body").innerText = (state.settings.governingBody || "UBND PHƯỜNG TRUNG NHỨT").toUpperCase();
        document.getElementById("lbg-school-name").innerText = (state.settings.schoolName || "TRƯỜNG TIỂU HỌC TRUNG NHỨT").toUpperCase();
        document.getElementById("lbg-grade-class").innerText = `${state.settings.grade || "KHỐI 5"} - ${state.settings.className || "LỚP 5A"}`;
        document.getElementById("lbg-week-title").innerText = `LỊCH BÁO GIẢNG TUẦN ${state.currentWeek}`;
        document.getElementById("lbg-date-range").innerText = `(Thời gian thực hiện: Từ ngày ${weekInfo.startDateVN} đến ngày ${weekInfo.endDateVN})`;

        // Sync LBG Options checkboxes
        const optColSign = document.getElementById("lbg-opt-col-sign");
        if (optColSign) optColSign.checked = !!state.lbgShowColSign;
        const optColNote = document.getElementById("lbg-opt-col-note");
        if (optColNote) optColNote.checked = !!state.lbgShowColNote;
        const optSigBgh = document.getElementById("lbg-opt-sig-bgh");
        if (optSigBgh) optSigBgh.checked = (state.lbgShowBghSign !== false);
        const optSigGvcn = document.getElementById("lbg-opt-sig-gvcn");
        if (optSigGvcn) optSigGvcn.checked = (state.lbgShowGvcnSign !== false);
        const optSigHead = document.getElementById("lbg-opt-sig-head");
        if (optSigHead) optSigHead.checked = (state.lbgShowHeadSign !== false);
        const optHideEmpty = document.getElementById("lbg-opt-hide-empty-rows");
        if (optHideEmpty) optHideEmpty.checked = !!state.lbgHideEmptyRows;

        // Render dynamic custom column chips in Tab 1
        renderCustomColsManager("lbg-custom-cols-list", false);

        // Compute ordered columns for Tab 1
        const orderedCols = getLbgOrderedColumns(state.lbgCustomCols, state.lbgShowColSign, state.lbgShowColNote, false);

        // Dynamically update Table Header columns
        const theadTr = document.querySelector("#tab-lbg .table-lbg thead tr");
        if (theadTr) {
            let colsHtml = "";
            orderedCols.forEach(col => {
                if (col.key === 'day') {
                    colsHtml += `<th style="width: 105px;">Thứ, ngày</th>`;
                } else if (col.key === 'session') {
                    colsHtml += `<th style="width: 65px;">Buổi</th>`;
                } else if (col.key === 'period') {
                    colsHtml += `<th style="width: 45px;">Tiết</th>`;
                } else if (col.key === 'subject') {
                    colsHtml += `<th style="width: 180px;">Môn học</th>`;
                } else if (col.key === 'ppct') {
                    colsHtml += `<th style="width: 75px;">Tiết PPCT</th>`;
                } else if (col.key === 'lesson') {
                    colsHtml += `<th>Tên bài dạy (Nhấp đúp vào ô để sửa trực tiếp)</th>`;
                } else if (col.key === 'sign') {
                    colsHtml += `<th style="width: 80px;" class="col-sign">Kí tên</th>`;
                } else if (col.key === 'note') {
                    colsHtml += `<th style="width: 120px;" class="col-note">Ghi chú</th>`;
                } else if (col.isCustom) {
                    colsHtml += `<th style="width: 120px;" class="col-custom col-custom-${col.id}">${escapeHtml(col.title)}</th>`;
                }
            });
            colsHtml += `<th style="width: 75px;" class="no-print">Thao tác</th>`;
            theadTr.innerHTML = colsHtml;
        }

        const tbody = document.getElementById("lbg-table-body");
        tbody.innerHTML = "";

        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];
        const dynamicSubjectOptions = getAllUniqueSubjects().concat(["-- Nghỉ / Để trống --"]);
        
        days.forEach(day => {
            const isAltDay = (day === "Thứ 2" || day === "Thứ 4" || day === "Thứ 6");
            const dayDate = getDayFullDate(weekInfo.startDateVN, day);
            const daySlots = schedule.filter(s => s.day === day);
            const morningSlots = daySlots.filter(s => s.session === "Sáng").sort((a, b) => (a.period || 0) - (b.period || 0));
            const afternoonSlots = daySlots.filter(s => s.session === "Chiều").sort((a, b) => (a.period || 0) - (b.period || 0));
            const orderedDaySlots = [...morningSlots, ...afternoonSlots];
            const totalRows = orderedDaySlots.length;

            orderedDaySlots.forEach((slot, idx) => {
                const tr = document.createElement("tr");
                if (isAltDay) tr.classList.add("row-day-alt");
                if (slot.isOff) tr.classList.add("row-empty-period");

                let dayCellHtml = "";
                if (idx === 0) {
                    dayCellHtml = `<td rowspan="${totalRows}" class="col-day ${isAltDay ? 'day-alt' : 'day-normal'}" style="line-height:1.35;">
                        <div style="font-weight:700;">${day}</div>
                        ${dayDate ? `<div style="font-size:0.78rem; font-weight:500; color:var(--text-muted); margin-top:2px;">${dayDate}</div>` : ''}
                    </td>`;
                }

                let sessionCellHtml = "";
                if (idx === 0) {
                    sessionCellHtml = `<td rowspan="${morningSlots.length}" style="text-align:center; font-weight:600; background:#fafafa;">Sáng</td>`;
                } else if (idx === morningSlots.length) {
                    sessionCellHtml = `<td rowspan="${afternoonSlots.length}" style="text-align:center; font-weight:600; background:#fafafa;">Chiều</td>`;
                }

                let rowHtml = "";
                orderedCols.forEach(col => {
                    if (col.key === 'day') {
                        if (dayCellHtml) rowHtml += dayCellHtml;
                    } else if (col.key === 'session') {
                        if (sessionCellHtml) rowHtml += sessionCellHtml;
                    } else if (col.key === 'period') {
                        rowHtml += `<td class="col-period">${slot.period}</td>`;
                    } else if (col.key === 'subject') {
                        rowHtml += `
                            <td class="col-subject" style="text-align:center;">
                                <select class="form-select form-select-sm subject-selector" data-key="${slot.key}" style="text-align:center; text-align-last:center;">
                                    ${dynamicSubjectOptions.map(s => `<option value="${s}" ${normalizeSubjectName(s) === normalizeSubjectName(slot.subject) || s === slot.subject ? 'selected' : ''}>${s}</option>`).join('')}
                                </select>
                            </td>
                        `;
                    } else if (col.key === 'ppct') {
                        rowHtml += `<td class="col-ppct editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="ppct">${slot.ppct}</td>`;
                    } else if (col.key === 'lesson') {
                        rowHtml += `<td class="col-lesson editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="lessonName" style="padding-top:1pt; padding-bottom:1pt; padding-left:1pt; text-indent:2pt; text-align:left; line-height:1.35;">${slot.lessonName}</td>`;
                    } else if (col.key === 'sign') {
                        rowHtml += `<td class="col-sign" style="text-align:center; color:#94a3b8; font-size:0.8rem;"></td>`;
                    } else if (col.key === 'note') {
                        rowHtml += `<td class="col-note editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="note" style="font-size:0.85rem; text-align:left;">${escapeHtml(slot.note || '')}</td>`;
                    } else if (col.isCustom) {
                        const customVal = (slot.customCols && slot.customCols[col.id] !== undefined)
                            ? slot.customCols[col.id]
                            : (col.id === 'col_1' ? (slot.customCol || slot.note || '') : (slot['customCol_' + col.id] || ''));
                        rowHtml += `<td class="col-custom editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="customCol_${col.id}" data-col-id="${col.id}" style="font-size:0.85rem; text-align:left;">${escapeHtml(customVal)}</td>`;
                    }
                });

                rowHtml += `
                    <td class="no-print" style="text-align: center;">
                        <div class="row-actions-group">
                            <button type="button" class="btn-row-action btn-add-week-slot" data-day="${slot.day}" data-session="${slot.session}" data-period="${slot.period}" title="Chèn thêm 1 tiết ngay phía dưới">+</button>
                            <button type="button" class="btn-row-action btn-del-week-slot" data-day="${slot.day}" data-session="${slot.session}" data-period="${slot.period}" title="Xóa tiết học này">🗑️</button>
                        </div>
                    </td>
                `;
                tr.innerHTML = rowHtml;
                tbody.appendChild(tr);
            });
        });

        document.getElementById("lbg-stat-gvcn").innerText = `${stats.gvcnCount} tiết`;
        document.getElementById("lbg-stat-specialist").innerText = `${stats.specialistCount} tiết`;
        document.getElementById("lbg-stat-enhanced").innerText = `${stats.enhancedCount} tiết`;
        document.getElementById("lbg-stat-total").innerText = `${stats.total} tiết`;

        const bghSigner = getBghSignerInfo();
        document.getElementById("lbg-sig-teacher").innerText = state.settings.homeroomTeacher || "Nguyễn Thị Thu Hà";
        document.getElementById("lbg-sig-head").innerText = state.settings.headOfGrade || "Trần Thị Mai";
        document.getElementById("lbg-sig-pht").innerText = bghSigner.name;
        const bghRoleElem = document.getElementById("lbg-bgh-role");
        if (bghRoleElem) bghRoleElem.innerText = `DUYỆT CỦA ${bghSigner.title}`;

        // Toggle signatures visibility in Tab 1
        const boxBgh = document.getElementById("lbg-sig-box-bgh");
        const boxHead = document.getElementById("lbg-sig-box-head");
        const boxGvcn = document.getElementById("lbg-sig-box-gvcn");
        if (boxBgh) boxBgh.style.display = (state.lbgShowBghSign !== false) ? "block" : "none";
        if (boxHead) boxHead.style.display = (state.lbgShowHeadSign !== false) ? "block" : "none";
        if (boxGvcn) boxGvcn.style.display = (state.lbgShowGvcnSign !== false) ? "block" : "none";

        tbody.querySelectorAll(".subject-selector").forEach(sel => {
            sel.addEventListener("change", (e) => {
                const key = e.target.dataset.key;
                if (!state.weeklyScheduleOverrides[key]) state.weeklyScheduleOverrides[key] = {};
                state.weeklyScheduleOverrides[key].subject = e.target.value;
                delete state.weeklyScheduleOverrides[key].lessonName;
                delete state.weeklyScheduleOverrides[key].ppct;
                delete state.weeklyScheduleOverrides[key].integration;
                saveState();
                renderTabLbg();
                renderTabCtlop();
            });
        });

        tbody.querySelectorAll(".editable-cell").forEach(cell => {
            cell.addEventListener("blur", (e) => {
                const key = e.target.dataset.key;
                const field = e.target.dataset.field;
                let val = e.target.innerText.trim();
                val = normalizePunctuationSpacing(val);
                if (!state.weeklyScheduleOverrides[key]) state.weeklyScheduleOverrides[key] = {};
                state.weeklyScheduleOverrides[key][field] = val;
                const colId = e.target.dataset.colId;
                if (colId) {
                    if (!state.weeklyScheduleOverrides[key].customCols) state.weeklyScheduleOverrides[key].customCols = {};
                    state.weeklyScheduleOverrides[key].customCols[colId] = val;
                    if (state.lbgCustomCols && state.lbgCustomCols[0] && state.lbgCustomCols[0].id === colId) {
                        state.weeklyScheduleOverrides[key].customCol = val;
                    }
                }
                saveState();
            });
        });

        tbody.querySelectorAll(".btn-add-week-slot").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const day = e.currentTarget.dataset.day;
                const session = e.currentTarget.dataset.session;
                const period = parseInt(e.currentTarget.dataset.period) || 1;
                addSlotToWeek(state.currentWeek, day, session, period);
            });
        });

        tbody.querySelectorAll(".btn-del-week-slot").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const day = e.currentTarget.dataset.day;
                const session = e.currentTarget.dataset.session;
                const period = parseInt(e.currentTarget.dataset.period) || 1;
                removeSlotFromWeek(state.currentWeek, day, session, period);
            });
        });
    }

    // =========================================================================
    // VERSION 6.0: TAB LỊCH BÁO GIẢNG THEO MÔN HỌC (TAB 2B)
    // =========================================================================
    function renderTabLbgMon() {
        renderWeekToolbar("lbgmon-week-toolbar", () => {
            renderTabLbgMon();
            if (state.currentTab === "tab-lbg") renderTabLbg();
            if (state.currentTab === "tab-ctlop") renderTabCtlop();
        });

        const weekNum = state.currentWeek;
        const { weekInfo } = calculateWeekSchedule(weekNum);

        // Update header
        const govEl = document.getElementById("lbgmon-gov-body");
        if (govEl) govEl.innerText = (state.settings.governingBody || "UBND PHƯỜNG TRUNG NHỨT").toUpperCase();
        const schEl = document.getElementById("lbgmon-school-name");
        if (schEl) schEl.innerText = (state.settings.schoolName || "TRƯỜNG TIỂU HỌC TRUNG NHỨT").toUpperCase();
        const gcEl = document.getElementById("lbgmon-grade-class");
        if (gcEl) gcEl.innerText = `${state.settings.grade || "KHỐI 5"} - ${state.settings.className || "LỚP 5A"}`;
        const titleEl = document.getElementById("lbgmon-week-title");
        if (titleEl) {
            if (state.lbgMonFilterSubject && state.lbgMonFilterSubject !== "all") {
                titleEl.innerText = `LỊCH BÁO GIẢNG MÔN ${state.lbgMonFilterSubject.toUpperCase()} - TUẦN ${weekNum}`;
            } else {
                titleEl.innerText = `LỊCH BÁO GIẢNG THEO MÔN HỌC TUẦN ${weekNum}`;
            }
        }
        const rangeEl = document.getElementById("lbgmon-date-range");
        if (rangeEl) rangeEl.innerText = `(Thời gian thực hiện: Từ ngày ${weekInfo.startDateVN} đến ngày ${weekInfo.endDateVN})`;

        // Populate Subject Filter dropdown
        const subFilterSel = document.getElementById("lbgmon-subject-filter");
        if (subFilterSel) {
            const rawSchedule = calculateWeekSchedule(weekNum).schedule.filter(s => !s.isOff);
            const activeSubs = Array.from(new Set(rawSchedule.map(s => s.subject)));
            const currentVal = state.lbgMonFilterSubject || "all";
            
            let opts = `<option value="all">🌟 Tất cả các môn học (Tuần tự từng môn)</option>`;
            activeSubs.forEach(s => {
                opts += `<option value="${escapeHtml(s)}" ${s === currentVal ? 'selected' : ''}>${escapeHtml(s)}</option>`;
            });
            subFilterSel.innerHTML = opts;

            subFilterSel.onchange = (e) => {
                state.lbgMonFilterSubject = e.target.value;
                renderTabLbgMon();
            };
        }

        // Category Filter
        const catFilterSel = document.getElementById("lbgmon-category-filter");
        if (catFilterSel) {
            catFilterSel.value = state.lbgMonFilterCategory || "all";
            catFilterSel.onchange = (e) => {
                state.lbgMonFilterCategory = e.target.value;
                renderTabLbgMon();
            };
        }

        // Show/Hide Integration
        const showIntegCb = document.getElementById("lbgmon-show-integration");
        if (showIntegCb) {
            showIntegCb.checked = (state.lbgMonShowIntegration !== false);
            showIntegCb.onchange = (e) => {
                state.lbgMonShowIntegration = e.target.checked;
                renderTabLbgMon();
            };
        }

        const showIntegration = (state.lbgMonShowIntegration !== false);
        const thInteg = document.querySelector("#table-lbg-mon .col-mon-integration");
        if (thInteg) {
            thInteg.style.display = showIntegration ? "" : "none";
        }

        // Calculate data
        const { subjectGroups, stats } = calculateWeekScheduleBySubject(
            weekNum,
            state.lbgMonFilterSubject,
            state.lbgMonFilterCategory
        );

        const tbody = document.getElementById("lbgmon-table-body");
        if (!tbody) return;
        tbody.innerHTML = "";

        if (subjectGroups.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="${showIntegration ? 9 : 8}" style="text-align: center; padding: 2.5rem 1rem; color: var(--text-muted); font-style: italic;">
                        Không có tiết học nào phù hợp với bộ lọc đã chọn trong Tuần ${weekNum}.
                    </td>
                </tr>
            `;
        } else {
            subjectGroups.forEach((group, gIdx) => {
                let badgeClass = "badge-cat-gvcn";
                let badgeText = "GVCN Dạy";
                if (group.category === "Chuyên trách") {
                    badgeClass = "badge-cat-specialist";
                    badgeText = "GV Chuyên Trách";
                } else if (group.category === "Tăng cường") {
                    badgeClass = "badge-cat-enhanced";
                    badgeText = "Tăng Cường";
                }

                group.slots.forEach((slot, idx) => {
                    const tr = document.createElement("tr");
                    if (gIdx % 2 === 1) tr.classList.add("row-day-alt");

                    let subjectCellHtml = "";
                    if (idx === 0) {
                        subjectCellHtml = `
                            <td rowspan="${group.slots.length}" class="col-subject-header" style="vertical-align: middle; text-align: center; font-weight: 700; background: #f8fafc; border-right: 2px solid #cbd5e1; padding: 0.75rem 0.6rem;">
                                <div style="font-size: 0.98rem; font-weight: 800; color: var(--primary);">${escapeHtml(group.subjectName)}</div>
                                <div style="font-size: 0.8rem; color: #475569; font-weight: 600; margin-top: 2px;">Tổng: ${group.slots.length} tiết/tuần</div>
                                <span class="badge-cat ${badgeClass}" style="margin-top: 4px; font-size: 0.72rem; display: inline-block;">${badgeText}</span>
                            </td>
                        `;
                    }

                    const dayDateStr = getDayDateStr(weekInfo.startDateVN, slot.day);

                    let integrationCellHtml = "";
                    if (showIntegration) {
                        integrationCellHtml = `
                            <td class="col-integration editable-cell" contenteditable="true" data-key="${slot.key}" data-field="integration" style="line-height: 1.45; font-size: 0.88rem; text-align: left;">
                                ${escapeHtml(slot.integration || '').replace(/\n/g, '<br>')}
                            </td>
                        `;
                    }

                    tr.innerHTML = `
                        ${subjectCellHtml}
                        <td style="text-align: center; font-weight: 600; font-size: 0.88rem; vertical-align: middle;">${dayDateStr}</td>
                        <td style="text-align: center; font-weight: 600; vertical-align: middle; color: ${slot.session === 'Sáng' ? 'var(--primary)' : '#b45309'};">${slot.session}</td>
                        <td class="col-period" style="vertical-align: middle; font-weight: 700;">${slot.period}</td>
                        <td style="text-align: center; font-weight: 700; color: var(--primary); vertical-align: middle;">Tiết ${slot.periodInWeek}</td>
                        <td class="col-ppct editable-cell" contenteditable="true" data-key="${slot.key}" data-field="ppct" style="vertical-align: middle; font-weight: 700;">${slot.ppct}</td>
                        <td class="col-lesson editable-cell" contenteditable="true" data-key="${slot.key}" data-field="lessonName" style="vertical-align: middle; text-align: left; padding-top: 1pt; padding-bottom: 1pt; padding-left: 1pt; text-indent: 2pt; line-height: 1.35;">${escapeHtml(slot.lessonName || '')}</td>
                        ${integrationCellHtml}
                        <td class="no-print" style="text-align: center; vertical-align: middle;">
                            <button type="button" class="btn-row-action btn-clear-mon-slot" data-key="${slot.key}" title="Xóa nội dung tiết này">✕</button>
                        </td>
                    `;

                    tbody.appendChild(tr);
                });
            });
        }

        // Stats summary
        const statSubsEl = document.getElementById("lbgmon-stat-subjects");
        if (statSubsEl) statSubsEl.innerText = `${stats.totalSubjects} môn`;
        const statGvcnEl = document.getElementById("lbgmon-stat-gvcn");
        if (statGvcnEl) statGvcnEl.innerText = `${stats.gvcnCount} tiết`;
        const statSpecEl = document.getElementById("lbgmon-stat-specialist");
        if (statSpecEl) statSpecEl.innerText = `${stats.specialistCount} tiết`;
        const statTotalEl = document.getElementById("lbgmon-stat-total");
        if (statTotalEl) statTotalEl.innerText = `${stats.totalPeriods} tiết`;

        // Signatures
        const bghSigner = getBghSignerInfo();
        const sigTeacherEl = document.getElementById("lbgmon-sig-teacher");
        if (sigTeacherEl) sigTeacherEl.innerText = state.settings.homeroomTeacher || "Nguyễn Thị Thu Hà";
        const sigHeadEl = document.getElementById("lbgmon-sig-head");
        if (sigHeadEl) sigHeadEl.innerText = state.settings.headOfGrade || "Trần Thị Mai";
        const sigPhtEl = document.getElementById("lbgmon-sig-pht");
        if (sigPhtEl) sigPhtEl.innerText = bghSigner.name;
        const bghRoleElem = document.getElementById("lbgmon-bgh-role");
        if (bghRoleElem) bghRoleElem.innerText = `DUYỆT CỦA ${bghSigner.title}`;

        // Bind editable cells blur
        tbody.querySelectorAll(".editable-cell").forEach(cell => {
            cell.addEventListener("blur", (e) => {
                const key = e.target.dataset.key;
                const field = e.target.dataset.field;
                let val = e.target.innerText.trim();
                val = normalizePunctuationSpacing(val);
                if (field === "integration") {
                    e.target.innerHTML = escapeHtml(val).replace(/\n/g, '<br>');
                } else {
                    e.target.innerText = val;
                }
                if (!state.weeklyScheduleOverrides[key]) state.weeklyScheduleOverrides[key] = {};
                state.weeklyScheduleOverrides[key][field] = val;
                saveState();
            });
        });

        // Clear slot button
        tbody.querySelectorAll(".btn-clear-mon-slot").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const key = e.currentTarget.dataset.key;
                if (!state.weeklyScheduleOverrides[key]) state.weeklyScheduleOverrides[key] = {};
                state.weeklyScheduleOverrides[key].lessonName = "";
                state.weeklyScheduleOverrides[key].ppct = "";
                state.weeklyScheduleOverrides[key].integration = "";
                saveState();
                renderTabLbgMon();
                showToast("Đã xóa nội dung bài dạy của tiết!", "info");
            });
        });
    }

    function previewLbgMonA4() {
        const orient = state.lbgMonOrientation || "portrait";
        const isCtlop = (state.lbgMonShowIntegration !== false);
        setAppOrientation(orient);
        const refreshFn = () => renderSingleWeekPaperBySubjectHtml(state.currentWeek, isCtlop, orient, state.lbgMonFilterSubject, state.lbgMonFilterCategory);
        openPreviewModal(
            `Xem trước Lịch Báo Giảng Theo Môn (Tuần ${state.currentWeek})`,
            refreshFn(),
            exportLbgMonDocx,
            exportLbgMonXlsx,
            () => printWithOrientation(orient),
            refreshFn
        );
    }

    function exportLbgMonDocx() {
        const orient = state.lbgMonOrientation || "portrait";
        const isCtlop = (state.lbgMonShowIntegration !== false);
        if (window.DocxGenerator && window.DocxGenerator.generateLbgBySubjectDocx) {
            showToast(`Đang tạo file Word (${orient === 'landscape' ? 'Khổ ngang' : 'Khổ đứng'}) theo môn học...`, "info");
            const data = calculateWeekScheduleBySubject(state.currentWeek, state.lbgMonFilterSubject, state.lbgMonFilterCategory);
            window.DocxGenerator.generateLbgBySubjectDocx(data, state.settings, isCtlop, orient, state.lbgMonFilterSubject).then(blob => {
                const subSuffix = (state.lbgMonFilterSubject && state.lbgMonFilterSubject !== 'all') ? `_${state.lbgMonFilterSubject.replace(/\s+/g, '_')}` : '';
                const orientSuffix = orient === "landscape" ? "_Kho_Ngang" : "";
                const filename = `Lich_Bao_Giang_Theo_Mon${subSuffix}_Tuan_${state.currentWeek}_${(state.settings.className || 'Lop_5A').replace(/\s+/g, '_')}${orientSuffix}.docx`;
                saveAs(blob, filename);
                showToast(`Đã xuất file Word thành công: ${filename}`, "success");
            }).catch(err => {
                console.error("DOCX By Subject error:", err);
                alert("Lỗi xuất file Word: " + err.message);
            });
        } else {
            alert("Bộ tạo file Word chưa sẵn sàng!");
        }
    }

    function exportLbgMonXlsx() {
        const isCtlop = (state.lbgMonShowIntegration !== false);
        if (window.XlsxGenerator && window.XlsxGenerator.generateLbgBySubjectXlsx) {
            showToast("Đang tạo file Excel theo môn học...", "info");
            const data = calculateWeekScheduleBySubject(state.currentWeek, state.lbgMonFilterSubject, state.lbgMonFilterCategory);
            window.XlsxGenerator.generateLbgBySubjectXlsx(data, state.settings, isCtlop, state.lbgMonFilterSubject).then(blob => {
                const subSuffix = (state.lbgMonFilterSubject && state.lbgMonFilterSubject !== 'all') ? `_${state.lbgMonFilterSubject.replace(/\s+/g, '_')}` : '';
                const filename = `Lich_Bao_Giang_Theo_Mon${subSuffix}_Tuan_${state.currentWeek}_${(state.settings.className || 'Lop_5A').replace(/\s+/g, '_')}.xlsx`;
                saveAs(blob, filename);
                showToast(`Đã xuất file Excel thành công: ${filename}`, "success");
            }).catch(err => {
                console.error("XLSX By Subject error:", err);
                alert("Lỗi xuất file Excel: " + err.message);
            });
        } else {
            alert("Bộ tạo file Excel chưa sẵn sàng!");
        }
    }

    // Render Tab 2: Lịch Báo Giảng Tích Hợp (CTLOP)
    function renderTabCtlop() {
        renderWeekToolbar("ctlop-week-toolbar", () => {
            renderTabCtlop();
            if (state.currentTab === "tab-lbg") renderTabLbg();
        });

        const { weekInfo, schedule, fridayPeriods } = calculateWeekSchedule(state.currentWeek);

        const fridayOptSelect = document.getElementById("ctlop-friday-opt");
        if (fridayOptSelect) {
            fridayOptSelect.value = fridayPeriods.toString();
            fridayOptSelect.onchange = (e) => {
                state.weeklyFridayPeriods[state.currentWeek] = parseInt(e.target.value);
                saveState();
                renderTabCtlop();
                if (state.currentTab === "tab-lbg") renderTabLbg();
            };
        }

        document.getElementById("ctlop-gov-body").innerText = (state.settings.governingBody || "UBND PHƯỜNG TRUNG NHỨT").toUpperCase();
        document.getElementById("ctlop-school-name").innerText = (state.settings.schoolName || "TRƯỜNG TIỂU HỌC TRUNG NHỨT").toUpperCase();
        document.getElementById("ctlop-grade-class").innerText = `${state.settings.grade || "KHỐI 5"} - ${state.settings.className || "LỚP 5A"}`;
        document.getElementById("ctlop-week-title").innerText = `LỊCH BÁO GIẢNG TÍCH HỢP TUẦN ${state.currentWeek}`;
        document.getElementById("ctlop-date-range").innerText = `(Thời gian thực hiện: Từ ngày ${weekInfo.startDateVN} đến ngày ${weekInfo.endDateVN})`;

        // Sync CTLOP Options checkboxes with state
        const ctlopOptColSign = document.getElementById("ctlop-opt-col-sign");
        if (ctlopOptColSign) ctlopOptColSign.checked = !!state.lbgShowColSign;
        const ctlopOptColNote = document.getElementById("ctlop-opt-col-note");
        if (ctlopOptColNote) ctlopOptColNote.checked = !!state.lbgShowColNote;
        const ctlopOptSigBgh = document.getElementById("ctlop-opt-sig-bgh");
        if (ctlopOptSigBgh) ctlopOptSigBgh.checked = (state.lbgShowBghSign !== false);
        const ctlopOptSigGvcn = document.getElementById("ctlop-opt-sig-gvcn");
        if (ctlopOptSigGvcn) ctlopOptSigGvcn.checked = (state.lbgShowGvcnSign !== false);
        const ctlopOptSigHead = document.getElementById("ctlop-opt-sig-head");
        if (ctlopOptSigHead) ctlopOptSigHead.checked = (state.lbgShowHeadSign !== false);
        const ctlopOptHideEmpty = document.getElementById("ctlop-opt-hide-empty-rows");
        if (ctlopOptHideEmpty) ctlopOptHideEmpty.checked = !!state.lbgHideEmptyRows;

        // Render dynamic custom column chips in Tab 2
        renderCustomColsManager("ctlop-custom-cols-list", false);

        // Compute ordered columns for Tab 2 (CTLOP)
        const orderedCols = getLbgOrderedColumns(state.lbgCustomCols, state.lbgShowColSign, state.lbgShowColNote, true);

        // Dynamically update Table Header columns for Tab 2
        const theadTr = document.querySelector("#tab-ctlop .table-lbg thead tr");
        if (theadTr) {
            let colsHtml = "";
            orderedCols.forEach(col => {
                if (col.key === 'day') {
                    colsHtml += `<th style="width: 105px;">Thứ, ngày</th>`;
                } else if (col.key === 'session') {
                    colsHtml += `<th style="width: 65px;">Buổi</th>`;
                } else if (col.key === 'period') {
                    colsHtml += `<th style="width: 45px;">Tiết</th>`;
                } else if (col.key === 'subject') {
                    colsHtml += `<th style="width: 170px;">Môn học</th>`;
                } else if (col.key === 'ppct') {
                    colsHtml += `<th style="width: 75px;">Tiết PPCT</th>`;
                } else if (col.key === 'lesson') {
                    colsHtml += `<th style="min-width: 250px;">Tên bài dạy</th>`;
                } else if (col.key === 'integ') {
                    colsHtml += `<th style="min-width: 220px;" class="col-integration">Nội dung tích hợp / Điều chỉnh</th>`;
                } else if (col.key === 'sign') {
                    colsHtml += `<th style="width: 80px;" class="col-sign">Kí tên</th>`;
                } else if (col.key === 'note') {
                    colsHtml += `<th style="width: 120px;" class="col-note">Ghi chú</th>`;
                } else if (col.isCustom) {
                    colsHtml += `<th style="width: 120px;" class="col-custom col-custom-${col.id}">${escapeHtml(col.title)}</th>`;
                }
            });
            colsHtml += `<th style="width: 75px;" class="no-print">Thao tác</th>`;
            theadTr.innerHTML = colsHtml;
        }

        const tbody = document.getElementById("ctlop-table-body");
        tbody.innerHTML = "";

        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];
        const dynamicSubjectOptions = getAllUniqueSubjects().concat(["-- Nghỉ / Để trống --"]);
        
        days.forEach(day => {
            const isAltDay = (day === "Thứ 2" || day === "Thứ 4" || day === "Thứ 6");
            const dayDate = getDayFullDate(weekInfo.startDateVN, day);
            const daySlots = schedule.filter(s => s.day === day);
            const morningSlots = daySlots.filter(s => s.session === "Sáng").sort((a, b) => (a.period || 0) - (b.period || 0));
            const afternoonSlots = daySlots.filter(s => s.session === "Chiều").sort((a, b) => (a.period || 0) - (b.period || 0));
            const orderedDaySlots = [...morningSlots, ...afternoonSlots];
            const totalRows = orderedDaySlots.length;

            orderedDaySlots.forEach((slot, idx) => {
                const tr = document.createElement("tr");
                if (isAltDay) tr.classList.add("row-day-alt");
                if (slot.isOff) tr.classList.add("row-empty-period");

                let dayCellHtml = "";
                if (idx === 0) {
                    dayCellHtml = `<td rowspan="${totalRows}" class="col-day ${isAltDay ? 'day-alt' : 'day-normal'}" style="line-height:1.35;">
                        <div style="font-weight:700;">${day}</div>
                        ${dayDate ? `<div style="font-size:0.78rem; font-weight:500; color:var(--text-muted); margin-top:2px;">${dayDate}</div>` : ''}
                    </td>`;
                }

                let sessionCellHtml = "";
                if (idx === 0) {
                    sessionCellHtml = `<td rowspan="${morningSlots.length}" style="text-align:center; font-weight:600; background:#fafafa;">Sáng</td>`;
                } else if (idx === morningSlots.length) {
                    sessionCellHtml = `<td rowspan="${afternoonSlots.length}" style="text-align:center; font-weight:600; background:#fafafa;">Chiều</td>`;
                }

                let rowHtml = "";
                orderedCols.forEach(col => {
                    if (col.key === 'day') {
                        if (dayCellHtml) rowHtml += dayCellHtml;
                    } else if (col.key === 'session') {
                        if (sessionCellHtml) rowHtml += sessionCellHtml;
                    } else if (col.key === 'period') {
                        rowHtml += `<td class="col-period">${slot.period}</td>`;
                    } else if (col.key === 'subject') {
                        rowHtml += `
                            <td class="col-subject" style="text-align:center;">
                                <select class="form-select form-select-sm subject-selector" data-key="${slot.key}" style="text-align:center; text-align-last:center;">
                                    ${dynamicSubjectOptions.map(s => `<option value="${s}" ${normalizeSubjectName(s) === normalizeSubjectName(slot.subject) || s === slot.subject ? 'selected' : ''}>${s}</option>`).join('')}
                                </select>
                            </td>
                        `;
                    } else if (col.key === 'ppct') {
                        rowHtml += `<td class="col-ppct editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="ppct">${slot.ppct}</td>`;
                    } else if (col.key === 'lesson') {
                        rowHtml += `<td class="col-lesson editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="lessonName" style="padding-top:1pt; padding-bottom:1pt; padding-left:1pt; text-indent:2pt; text-align:left; line-height:1.35;">${slot.lessonName}</td>`;
                    } else if (col.key === 'integ') {
                        rowHtml += `<td class="col-integration editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="integration" style="line-height:1.45;">${escapeHtml(slot.integration || '').replace(/\n/g, '<br>')}</td>`;
                    } else if (col.key === 'sign') {
                        rowHtml += `<td class="col-sign" style="text-align:center; color:#94a3b8; font-size:0.8rem;"></td>`;
                    } else if (col.key === 'note') {
                        rowHtml += `<td class="col-note editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="note" style="font-size:0.85rem; text-align:left;">${escapeHtml(slot.note || '')}</td>`;
                    } else if (col.isCustom) {
                        const customVal = (slot.customCols && slot.customCols[col.id] !== undefined)
                            ? slot.customCols[col.id]
                            : (col.id === 'col_1' ? (slot.customCol || slot.note || '') : (slot['customCol_' + col.id] || ''));
                        rowHtml += `<td class="col-custom editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="customCol_${col.id}" data-col-id="${col.id}" style="font-size:0.85rem; text-align:left;">${escapeHtml(customVal)}</td>`;
                    }
                });

                rowHtml += `
                    <td class="no-print" style="text-align: center;">
                        <div class="row-actions-group">
                            <button type="button" class="btn-row-action btn-add-week-slot" data-day="${slot.day}" data-session="${slot.session}" data-period="${slot.period}" title="Chèn thêm 1 tiết ngay phía dưới">+</button>
                            <button type="button" class="btn-row-action btn-del-week-slot" data-day="${slot.day}" data-session="${slot.session}" data-period="${slot.period}" title="Xóa tiết học này">🗑️</button>
                        </div>
                    </td>
                `;
                tr.innerHTML = rowHtml;
                tbody.appendChild(tr);
            });
        });

        // Update signature box in Tab 2
        const bghSigner = getBghSignerInfo();
        const ctlopTeacher = document.getElementById("ctlop-sig-teacher");
        if (ctlopTeacher) ctlopTeacher.innerText = state.settings.homeroomTeacher || "Nguyễn Thị Thu Hà";
        const ctlopHead = document.getElementById("ctlop-sig-head");
        if (ctlopHead) ctlopHead.innerText = state.settings.headOfGrade || "Trần Thị Mai";
        const ctlopPht = document.getElementById("ctlop-sig-pht");
        if (ctlopPht) ctlopPht.innerText = bghSigner.name;
        const ctlopBghRoleElem = document.getElementById("ctlop-bgh-role");
        if (ctlopBghRoleElem) ctlopBghRoleElem.innerText = `DUYỆT CỦA ${bghSigner.title}`;

        const boxCtlopBgh = document.getElementById("ctlop-sig-box-bgh");
        const boxCtlopHead = document.getElementById("ctlop-sig-box-head");
        const boxCtlopGvcn = document.getElementById("ctlop-sig-box-gvcn");
        if (boxCtlopBgh) boxCtlopBgh.style.display = (state.lbgShowBghSign !== false) ? "block" : "none";
        if (boxCtlopHead) boxCtlopHead.style.display = (state.lbgShowHeadSign !== false) ? "block" : "none";
        if (boxCtlopGvcn) boxCtlopGvcn.style.display = (state.lbgShowGvcnSign !== false) ? "block" : "none";

        tbody.querySelectorAll(".subject-selector").forEach(sel => {
            sel.addEventListener("change", (e) => {
                const key = e.target.dataset.key;
                if (!state.weeklyScheduleOverrides[key]) state.weeklyScheduleOverrides[key] = {};
                state.weeklyScheduleOverrides[key].subject = e.target.value;
                delete state.weeklyScheduleOverrides[key].lessonName;
                delete state.weeklyScheduleOverrides[key].ppct;
                delete state.weeklyScheduleOverrides[key].integration;
                saveState();
                renderTabCtlop();
                renderTabLbg();
            });
        });

        tbody.querySelectorAll(".editable-cell").forEach(cell => {
            cell.addEventListener("blur", (e) => {
                const key = e.target.dataset.key;
                const field = e.target.dataset.field;
                let val = e.target.innerText.trim();
                val = normalizePunctuationSpacing(val);
                if (field === "integration") {
                    e.target.innerHTML = escapeHtml(val).replace(/\n/g, '<br>');
                } else {
                    e.target.innerText = val;
                }
                if (!state.weeklyScheduleOverrides[key]) state.weeklyScheduleOverrides[key] = {};
                state.weeklyScheduleOverrides[key][field] = val;
                const colId = e.target.dataset.colId;
                if (colId) {
                    if (!state.weeklyScheduleOverrides[key].customCols) state.weeklyScheduleOverrides[key].customCols = {};
                    state.weeklyScheduleOverrides[key].customCols[colId] = val;
                    if (state.lbgCustomCols && state.lbgCustomCols[0] && state.lbgCustomCols[0].id === colId) {
                        state.weeklyScheduleOverrides[key].customCol = val;
                    }
                }
                saveState();
            });
        });

        tbody.querySelectorAll(".btn-add-week-slot").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const day = e.currentTarget.dataset.day;
                const session = e.currentTarget.dataset.session;
                const period = parseInt(e.currentTarget.dataset.period) || 1;
                addSlotToWeek(state.currentWeek, day, session, period);
            });
        });

        tbody.querySelectorAll(".btn-del-week-slot").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const day = e.currentTarget.dataset.day;
                const session = e.currentTarget.dataset.session;
                const period = parseInt(e.currentTarget.dataset.period) || 1;
                removeSlotFromWeek(state.currentWeek, day, session, period);
            });
        });
    }

    // 4. Smart Calendar Generator (Tab 3)
    function getNextMonday(date) {
        const d = new Date(date.getTime());
        const day = d.getDay(); // 0 is Sunday, 1 is Monday...
        const diff = (day === 0) ? 1 : (8 - day);
        d.setDate(d.getDate() + diff);
        d.setHours(0, 0, 0, 0);
        return d;
    }

    // 4. Smart Calendar Generator (Tab 3 - Tự động tạo 35 tuần theo ngày nghỉ Tết tùy chỉnh)
    function autoGenerateCalendar(yearString, startDateString, tetStartStr, tetEndStr) {
        if (!startDateString) return;
        const startDt = new Date(startDateString + "T00:00:00");
        if (isNaN(startDt.getTime())) return;

        let currStart = new Date(startDt.getTime());
        const startDay = currStart.getDay();
        if (startDay === 0) currStart.setDate(currStart.getDate() + 1);
        else if (startDay > 1) currStart.setDate(currStart.getDate() - (startDay - 1));
        currStart.setHours(0, 0, 0, 0);

        let customTetStart = null;
        let customTetEnd = null;
        if (tetStartStr && tetEndStr) {
            customTetStart = new Date(tetStartStr + "T00:00:00");
            customTetEnd = new Date(tetEndStr + "T23:59:59");
        }

        const formatVN = (d) => {
            const day = String(d.getDate()).padStart(2, '0');
            const month = String(d.getMonth() + 1).padStart(2, '0');
            const yr = d.getFullYear();
            return `${day}/${month}/${yr}`;
        };

        const newWeeks = [];
        let tetShiftApplied = false;
        let tetBreakWeekNum = -1;

        for (let w = 1; w <= 35; w++) {
            const term = w <= 18 ? 1 : 2;
            let currEnd = new Date(currStart.getTime() + 4 * 24 * 60 * 60 * 1000); // Thứ Sáu
            currEnd.setHours(23, 59, 59, 999);

            // Kiểm tra nếu tuần học dự kiến trùng/giao với kỳ nghỉ Tết Âm lịch
            if (customTetStart && customTetEnd && !tetShiftApplied) {
                const overlapsTet = (currStart <= customTetEnd && currEnd >= customTetStart);
                if (overlapsTet) {
                    tetBreakWeekNum = w;
                    currStart = getNextMonday(new Date(tetEndStr + "T00:00:00"));
                    currEnd = new Date(currStart.getTime() + 4 * 24 * 60 * 60 * 1000);
                    currEnd.setHours(23, 59, 59, 999);
                    tetShiftApplied = true;
                }
            }

            const monthStr = `Tháng ${currStart.getMonth() + 1} năm ${currStart.getFullYear()}`;
            let holiday = "";
            let note = "";

            const m = currStart.getMonth() + 1;
            const sDay = currStart.getDate();
            const eDay = currEnd.getDate();

            // Ngày nghỉ lễ chuẩn Quốc gia
            if (m === 9 && sDay <= 5 && eDay >= 2) {
                holiday = "Lễ Quốc khánh (02/09)";
                note = "Khai giảng năm học (05/09)";
            } else if (m === 11 && sDay <= 20 && eDay >= 20) {
                holiday = "Ngày Nhà giáo Việt Nam (20/11)";
            } else if ((m === 12 && eDay >= 31) || (m === 1 && sDay <= 1)) {
                holiday = "Tết Dương lịch (01/01)";
            } else if (m === 4 && sDay <= 16 && eDay >= 16) {
                holiday = "Giỗ Tổ Hùng Vương (10/3 ÂL - 16/04)";
            } else if ((m === 4 && eDay >= 30) || (m === 5 && sDay <= 1)) {
                holiday = "Nghỉ lễ 30/4 và 1/5";
            }

            // Ghi chú chi tiết kỳ nghỉ Tết Âm lịch vào tuần trở lại trường
            if (tetShiftApplied && w === tetBreakWeekNum && customTetStart && customTetEnd) {
                const sTetStr = `${String(customTetStart.getDate()).padStart(2, '0')}/${String(customTetStart.getMonth() + 1).padStart(2, '0')}`;
                const eTetStr = `${String(customTetEnd.getDate()).padStart(2, '0')}/${String(customTetEnd.getMonth() + 1).padStart(2, '0')}/${customTetEnd.getFullYear()}`;
                holiday = `Nghỉ Tết Âm lịch (${sTetStr} - ${eTetStr})`;
                note = "Thực học sau kỳ nghỉ Tết Âm lịch";
            }

            if (w === 18) {
                note = note ? `${note} - Kết thúc HK I` : "Kết thúc Học kỳ I";
            } else if (w === 19) {
                note = note ? `${note} - Bắt đầu HK II` : "Bắt đầu Học kỳ II";
            } else if (w === 35) {
                note = "Tổng kết năm học & Hoàn thành hồ sơ";
            }

            newWeeks.push({
                week: w,
                term: term,
                startDateVN: formatVN(currStart),
                endDateVN: formatVN(currEnd),
                month: monthStr,
                holiday: holiday,
                note: note
            });

            currStart = new Date(currStart.getTime() + 7 * 24 * 60 * 60 * 1000);
            currStart.setHours(0, 0, 0, 0);
        }

        state.weeks = newWeeks;
        state.settings.academicYear = yearString;
        state.settings.startDate = startDateString;
        state.settings.tetStartDate = tetStartStr;
        state.settings.tetEndDate = tetEndStr;
        saveState();
        renderTabLichtuan();
        renderTabLbg();
        renderTabCtlop();
        showToast(`Đã tự động tính toán lại toàn bộ 35 tuần năm học ${yearString}!`, "success");
    }

    // Render Tab 3: Kế hoạch Thời gian (Lịch Tuần)
    function renderTabLichtuan() {
        const genYearInput = document.getElementById("gen-year");
        if (genYearInput && state.settings.academicYear) {
            genYearInput.value = state.settings.academicYear;
        }
        const genStartInput = document.getElementById("gen-start-date");
        if (genStartInput && state.settings.startDate) {
            genStartInput.value = state.settings.startDate;
        }
        const genTetStartInput = document.getElementById("gen-tet-start-date");
        if (genTetStartInput && state.settings.tetStartDate) {
            genTetStartInput.value = state.settings.tetStartDate;
        }
        const genTetEndInput = document.getElementById("gen-tet-end-date");
        if (genTetEndInput && state.settings.tetEndDate) {
            genTetEndInput.value = state.settings.tetEndDate;
        }

        const tbody = document.getElementById("lichtuan-table-body");
        if (!tbody) return;
        tbody.innerHTML = "";

        state.weeks.forEach(w => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td style="text-align:center; font-weight:bold;">${w.term === 1 ? 'Học kỳ I' : 'Học kỳ II'}</td>
                <td style="text-align:center; font-weight:bold; color:var(--primary);">Tuần ${w.week}</td>
                <td style="text-align:center;" class="editable-cell week-edit" contenteditable="true" data-week="${w.week}" data-field="startDateVN">${w.startDateVN}</td>
                <td style="text-align:center;" class="editable-cell week-edit" contenteditable="true" data-week="${w.week}" data-field="endDateVN">${w.endDateVN}</td>
                <td class="editable-cell week-edit" contenteditable="true" data-week="${w.week}" data-field="month">${w.month}</td>
                <td style="color:#b45309; font-weight:500;" class="editable-cell week-edit" contenteditable="true" data-week="${w.week}" data-field="holiday">${w.holiday || ''}</td>
                <td class="editable-cell week-edit" contenteditable="true" data-week="${w.week}" data-field="note">${w.note || ''}</td>
            `;
            tbody.appendChild(tr);
        });

        tbody.querySelectorAll(".week-edit").forEach(cell => {
            cell.addEventListener("blur", (e) => {
                const wNum = parseInt(e.target.dataset.week);
                const field = e.target.dataset.field;
                const weekObj = state.weeks.find(w => w.week === wNum);
                if (weekObj) {
                    weekObj[field] = e.target.innerText.trim();
                    saveState();
                }
            });
        });
    }

    // 5. PPCT Management & Excel Import/Export
    function downloadPpctTemplate() {
        if (typeof XLSX === "undefined") {
            alert("Thư viện SheetJS chưa sẵn sàng!");
            return;
        }

        const wb = XLSX.utils.book_new();
        const headers = ["Tuần", "Môn học", "Tiết trong tuần", "Tiết PPCT", "Tên bài dạy", "Nội dung tích hợp / Điều chỉnh"];
        
        const sampleRows = [
            headers,
            [1, "HĐ Trải nghiệm", 1, 1, "Sinh hoạt dưới cờ: Chào năm học mới", "Tích hợp LTCM, QCN"],
            [1, "Tiếng Việt", 1, 1, "Đọc: Thanh âm của gió", "Tích hợp Năng lực số (NLS), Giáo dục bảo vệ môi trường"],
            [1, "Tin học", 1, 1, "Bài 1. Em có thể làm gì với máy tính (Tiết 1)", "Tích hợp NLS"],
            [1, "Tiếng Việt", 2, 2, "Luyện từ và câu: Luyện tập về danh từ, động từ, tính từ", ""],
            [1, "Tiếng Anh", 1, 1, "Make a rule board and have practise", ""],
            [1, "Toán", 1, 1, "Bài 1. Ôn tập số tự nhiên (tiết 1): Luyện tập (Trang 6)", "Tích hợp NLS"],
            [1, "TC Toán", 1, 1, "Ôn tập Toán - Tiết 1", "Rèn kỹ năng tính toán"],
            [1, "Khoa học", 1, 1, "Bài 1: Thành phần và vai trò của đất đối với cây trồng (Tiết 1)", "Giáo dục môi trường"],
            [1, "LS&ĐL", 1, 1, "Bài 1: Vị trí địa lí, lãnh thổ, đơn vị hành chính, Quốc kì, Quốc huy (Tiết 1)", "Tích hợp QPAN, QCN"],
            [1, "Đạo đức", 1, 1, "Bài 1. Biết ơn những người có công với quê hương, đất nước (tiết 1)", "Tích hợp LTCM, QPAN"],
            [1, "Công nghệ", 1, 1, "Bài 1. Vai trò của công nghệ (Tiết 1)", "Tích hợp STEM, AI"],
            [1, "GD Thể chất", 1, 1, "Bài 1: Bài tập phối hợp ĐHĐN", ""],
            [1, "Âm nhạc", 1, 1, "Đọc nhạc: Bài số 1", ""],
            [1, "Mĩ thuật", 1, 1, "Yếu tố tạo hình trong thực hành, sáng tạo theo chủ đề (tiết 1)", ""],
            [1, "Đọc Thư viện", 1, 1, "Đọc sách trong Thư viện", ""],
            [1, "TC Tiếng Việt", 1, 1, "Ôn tập Tiếng Việt - Tiết 1", ""]
        ];

        const ws = XLSX.utils.aoa_to_sheet(sampleRows);
        ws['!cols'] = [{ wch: 8 }, { wch: 18 }, { wch: 15 }, { wch: 12 }, { wch: 45 }, { wch: 35 }];
        XLSX.utils.book_append_sheet(wb, ws, "Phan_Phoi_CT");

        XLSX.writeFile(wb, "Mau_Phan_Phoi_Chuong_Trinh_Lop_5.xlsx");
        showToast("Đã tải về file Excel mẫu phân phối chương trình!", "success");
    }

    function exportPpctToExcel() {
        if (typeof XLSX === "undefined") {
            alert("Thư viện SheetJS chưa sẵn sàng!");
            return;
        }

        const ppctData = state.ppct || [];
        if (ppctData.length === 0) {
            alert("Không có dữ liệu Phân phối chương trình để xuất!");
            return;
        }

        // Get current filter selections
        const filterSubject = document.getElementById("ppct-filter-subject");
        const filterWeek = document.getElementById("ppct-filter-week");
        const searchInput = document.getElementById("ppct-search");

        const selSubject = filterSubject ? filterSubject.value : "ALL";
        const selWeek = filterWeek ? filterWeek.value : "ALL";
        const searchTerm = searchInput ? searchInput.value.trim().toLowerCase() : "";

        // Filter data same as table display
        let filteredData = ppctData;
        if (selSubject !== "ALL") {
            filteredData = filteredData.filter(p => normalizeSubjectName(p.subject) === normalizeSubjectName(selSubject));
        }
        if (selWeek !== "ALL") {
            filteredData = filteredData.filter(p => String(p.week) === String(selWeek));
        }
        if (searchTerm) {
            filteredData = filteredData.filter(p => {
                const combined = [p.subject, p.lessonName, p.integration, String(p.week), String(p.ppct)].join(" ").toLowerCase();
                return combined.includes(searchTerm);
            });
        }

        if (filteredData.length === 0) {
            alert("Không có tiết học nào phù hợp với bộ lọc hiện tại để xuất!");
            return;
        }

        const wb = XLSX.utils.book_new();
        const gradeName = state.settings.grade || ("Khối " + state.currentGrade);
        const className = state.settings.className || ("Lớp " + state.currentGrade + "A");
        const schoolName = state.settings.schoolName || "TRƯỜNG TIỂU HỌC";

        // Header rows
        const headerRows = [
            [schoolName, "", "", "", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", ""],
            [gradeName.toUpperCase() + " - " + className.toUpperCase(), "", "", "", "Độc lập - Tự do - Hạnh phúc", ""],
            [""],
            ["PHÂN PHỐI CHƯƠNG TRÌNH " + gradeName.toUpperCase(), "", "", "", "", ""],
            ["Năm học: " + (state.settings.academicYear || "2026 - 2027"), "", "", "", "", ""],
            [""]
        ];

        // Column headers
        const colHeaders = ["Tuần", "Môn học", "Tiết/Tuần", "Tiết PPCT", "Tên bài dạy", "Nội dung tích hợp / Điều chỉnh"];
        headerRows.push(colHeaders);

        // Data rows
        const dataRows = filteredData.map(p => [
            p.week || "",
            p.subject || "",
            p.periodInWeek || "",
            p.ppct || "",
            p.lessonName || "",
            p.integration || ""
        ]);

        const allRows = [...headerRows, ...dataRows];
        const ws = XLSX.utils.aoa_to_sheet(allRows);

        // Column widths
        ws['!cols'] = [
            { wch: 8 },   // Tuần
            { wch: 20 },  // Môn học
            { wch: 10 },  // Tiết/Tuần
            { wch: 10 },  // Tiết PPCT
            { wch: 50 },  // Tên bài dạy
            { wch: 40 }   // Nội dung tích hợp / Điều chỉnh
        ];

        // Merge cells for header area
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } },  // School name
            { s: { r: 0, c: 4 }, e: { r: 0, c: 5 } },  // Quốc hiệu
            { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } },  // Grade-Class
            { s: { r: 1, c: 4 }, e: { r: 1, c: 5 } },  // Tiêu ngữ
            { s: { r: 3, c: 0 }, e: { r: 3, c: 5 } },  // Title
            { s: { r: 4, c: 0 }, e: { r: 4, c: 5 } }   // Academic year
        ];

        const filterNote = [];
        if (selSubject !== "ALL") filterNote.push("Môn: " + selSubject);
        if (selWeek !== "ALL") filterNote.push("Tuần: " + selWeek);
        if (searchTerm) filterNote.push("Tìm kiếm: " + searchTerm);
        const sheetName = filterNote.length > 0
            ? "PPCT_Loc"
            : "PPCT_" + gradeName.replace(/[^a-zA-Z0-9_\u00C0-\u024F\u1EA0-\u1EF9]/g, "_");

        XLSX.utils.book_append_sheet(wb, ws, sheetName.substring(0, 31));

        const safeClassName = className.replace(/[^a-zA-Z0-9_\u00C0-\u024F\u1EA0-\u1EF9]/g, "_");
        const fileName = `Phan_Phoi_Chuong_Trinh_${safeClassName}_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, fileName);

        const totalInfo = filteredData.length === ppctData.length
            ? `toàn bộ ${filteredData.length} tiết`
            : `${filteredData.length}/${ppctData.length} tiết (đã lọc)`;
        showToast(`Đã xuất file Excel Phân phối chương trình ${gradeName} (${totalInfo}) thành công!`, "success");
    }

    // Global variable for pending import
    window._pendingPpctImport = null;

    function promptImportModeModal(parsedRows, fileName) {
        if (!parsedRows || parsedRows.length === 0) {
            alert("Không tìm thấy dữ liệu tiết học hợp lệ trong file!");
            return;
        }

        window._pendingPpctImport = {
            rows: parsedRows,
            fileName: fileName,
            targetGrade: state.currentGrade
        };

        const modal = document.getElementById("modal-import-mode");
        const countEl = document.getElementById("import-row-count");
        const fileEl = document.getElementById("import-file-name");
        const gradeEl = document.getElementById("import-target-grade");

        if (countEl) countEl.textContent = parsedRows.length.toLocaleString('vi-VN');
        if (fileEl) fileEl.textContent = fileName;
        if (gradeEl) gradeEl.textContent = state.currentGrade;

        if (modal) {
            modal.style.display = "flex";
        }
    }

    function executePendingImport(mode) {
        if (!window._pendingPpctImport || !window._pendingPpctImport.rows) {
            return;
        }

        const { rows, fileName, targetGrade } = window._pendingPpctImport;

        if (mode === "replace") {
            state.ppct = rows;
            saveState();
            renderTabPpct();
            renderTabLbg();
            renderTabCtlop();
            updatePpctCountBadge();
            showToast(`Đã thay thế toàn bộ bằng ${rows.length} tiết PPCT mới Khối ${targetGrade} từ file!`, "success");
        } else if (mode === "merge") {
            let updatedCount = 0;
            let insertedCount = 0;

            rows.forEach(item => {
                const existingIdx = state.ppct.findIndex(p => 
                    p.week === item.week && 
                    normalizeSubjectName(p.subject) === normalizeSubjectName(item.subject) && 
                    p.periodInWeek === item.periodInWeek
                );

                if (existingIdx >= 0) {
                    state.ppct[existingIdx].ppct = item.ppct;
                    state.ppct[existingIdx].lessonName = item.lessonName;
                    if (item.integration) state.ppct[existingIdx].integration = item.integration;
                    if (item.duration) state.ppct[existingIdx].duration = item.duration;
                    updatedCount++;
                } else {
                    state.ppct.push(item);
                    insertedCount++;
                }
            });

            saveState();
            renderTabPpct();
            renderTabLbg();
            renderTabCtlop();
            updatePpctCountBadge();
            showToast(`Đã cập nhật ${updatedCount} tiết và bổ sung mới ${insertedCount} tiết PPCT từ file!`, "success");
        }

        window._pendingPpctImport = null;
        const modal = document.getElementById("modal-import-mode");
        if (modal) modal.style.display = "none";
    }

    function matchPpctColumn(cellText) {
        const t = (cellText || "").toString().normalize("NFC").toLowerCase().trim();
        if (t.includes("ppct") || t.includes("tiết ppct") || t.includes("tiết theo") || t.includes("tiết phân phối") || t.includes("tiet ppct")) return "ppct";
        if (t.includes("tiết/tuần") || t.includes("tiết / tuần") || t.includes("tiết trong tuần") || t.includes("tiết thứ") || t.includes("tiet thu") || t.includes("số tiết") || t.includes("tiet/t") || t.includes("tiết/ t") || t === "tiết" || t === "tiet") return "periodInWeek";
        if (t.includes("tuần") || t.includes("tuan") || t === "w") return "week";
        if (t.includes("môn") || t.includes("mon") || t.includes("phân môn") || t.includes("subject")) return "subject";
        if (t.includes("tên bài") || t.includes("bài học") || t.includes("tên bài dạy") || t.includes("nội dung bài") || t.includes("nội dung dạy học") || t.includes("hoạt động dạy") || t.includes("tên bài/hoạt động") || (t.includes("bài") && !t.includes("căn cứ"))) return "lesson";
        if (t.includes("tích hợp") || t.includes("điều chỉnh") || t.includes("bổ sung") || t.includes("ghi chú") || t.includes("integration") || t.includes("lưu ý")) return "integration";
        return null;
    }

    function checkMissingPpctColumns(colMap) {
        const missing = [];
        if (colMap.week === -1) missing.push("Tuần (hoặc Tuần học)");
        if (colMap.subject === -1) missing.push("Môn học (hoặc Phân môn)");
        if (colMap.periodInWeek === -1) missing.push("Tiết trong tuần (hoặc Tiết/tuần, Tiết thứ)");
        if (colMap.ppct === -1) missing.push("Tiết PPCT (hoặc Tiết theo CT)");
        if (colMap.lesson === -1) missing.push("Tên bài dạy (hoặc Bài học, Tên bài)");
        return missing;
    }

    function showMissingColumnsAlert(fileTypeName, missingCols) {
        alert(`⚠️ FILE CHƯA ĐÚNG MẪU QUY ĐỊNH!\n\nFile ${fileTypeName} vừa tải lên không đảm bảo các cột bắt buộc:\n• ${missingCols.join('\n• ')}\n\n📌 Yêu cầu cấu trúc bảng phân phối chương trình cần có đủ các cột:\n 1. Tuần (Bắt buộc)\n 2. Môn học (Bắt buộc)\n 3. Tiết trong tuần / Tiết thứ (Bắt buộc)\n 4. Tiết PPCT (Bắt buộc)\n 5. Tên bài dạy (Bắt buộc)\n 6. Nội dung tích hợp / Điều chỉnh (Không bắt buộc, có thì thêm vào)\n\n👉 Thầy/Cô vui lòng kiểm tra lại bảng trong file hoặc nhấn "📥 Tải Mẫu Excel" để lấy file mẫu chuẩn.`);
    }

    function uploadPpctFromFile(file) {
        if (!file) return;
        const fileName = (file.name || "").toLowerCase();

        if (fileName.endsWith(".docx")) {
            if (typeof JSZip === "undefined") {
                alert("Thư viện JSZip chưa sẵn sàng để đọc file Word!");
                return;
            }
            const reader = new FileReader();
            reader.onload = async function(e) {
                try {
                    const zip = await JSZip.loadAsync(e.target.result);
                    const docXmlFile = zip.file("word/document.xml");
                    if (!docXmlFile) {
                        alert("File Word không đúng định dạng chuẩn (không tìm thấy document.xml)!");
                        return;
                    }
                    const xmlText = await docXmlFile.async("string");
                    const parser = new DOMParser();
                    const xmlDoc = parser.parseFromString(xmlText, "application/xml");
                    const tables = xmlDoc.getElementsByTagName("w:tbl");

                    if (!tables || tables.length === 0) {
                        alert("Không tìm thấy bảng biểu phân phối chương trình trong file Word này!");
                        return;
                    }

                    // Extract rows from the best matching table
                    let bestTableRows = null;
                    let bestColMap = { week: -1, subject: -1, periodInWeek: -1, ppct: -1, lesson: -1, integration: -1 };
                    let maxMatchCount = 0;

                    for (let t = 0; t < tables.length; t++) {
                        const tbl = tables[t];
                        const rows = tbl.getElementsByTagName("w:tr");
                        if (rows.length < 2) continue;

                        const headerRow = rows[0];
                        const tcs = headerRow.getElementsByTagName("w:tc");
                        const currentMap = { week: -1, subject: -1, periodInWeek: -1, ppct: -1, lesson: -1, integration: -1 };
                        let matchScore = 0;

                        for (let c = 0; c < tcs.length; c++) {
                            const matchedType = matchPpctColumn(tcs[c].textContent);
                            if (matchedType && currentMap[matchedType] === -1) {
                                currentMap[matchedType] = c;
                                matchScore++;
                            }
                        }

                        if (matchScore > maxMatchCount) {
                            maxMatchCount = matchScore;
                            bestTableRows = rows;
                            bestColMap = currentMap;
                        }
                    }

                    if (!bestTableRows) {
                        alert("Không nhận diện được bảng biểu nào trong file Word!");
                        return;
                    }

                    // Check missing mandatory columns
                    const missingCols = checkMissingPpctColumns(bestColMap);
                    if (missingCols.length > 0) {
                        showMissingColumnsAlert("Word (.docx)", missingCols);
                        return;
                    }

                    const parsedRows = [];
                    let currentW = 1;
                    let currentSub = "Tiếng Việt";

                    for (let r = 1; r < bestTableRows.length; r++) {
                        const tr = bestTableRows[r];
                        const tcs = tr.getElementsByTagName("w:tc");
                        if (tcs.length === 0) continue;

                        const getCellText = (idx) => {
                            if (idx >= 0 && idx < tcs.length) {
                                return (tcs[idx].textContent || "").trim();
                            }
                            return "";
                        };

                        const wVal = parseInt(getCellText(bestColMap.week), 10);
                        if (!isNaN(wVal) && wVal >= 1 && wVal <= 35) currentW = wVal;

                        const sVal = getCellText(bestColMap.subject);
                        if (sVal && !sVal.toLowerCase().includes("môn")) currentSub = normalizeSubjectName(sVal);

                        const lessonVal = normalizePunctuationSpacing(getCellText(bestColMap.lesson));
                        if (!lessonVal || lessonVal.toLowerCase().includes("tên bài") || lessonVal.toLowerCase().includes("bài học")) continue;

                        let pInWeek = parseInt(getCellText(bestColMap.periodInWeek), 10);
                        if (isNaN(pInWeek) || pInWeek <= 0) pInWeek = 1;

                        let ppctNum = parseInt(getCellText(bestColMap.ppct), 10);
                        if (isNaN(ppctNum) || ppctNum <= 0) ppctNum = parsedRows.length + 1;

                        const integVal = bestColMap.integration !== -1 
                            ? normalizePunctuationSpacing(getCellText(bestColMap.integration)) 
                            : "";

                        parsedRows.push({
                            week: currentW,
                            subject: currentSub,
                            periodInWeek: pInWeek,
                            ppct: ppctNum,
                            lessonName: lessonVal,
                            integration: integVal
                        });
                    }

                    promptImportModeModal(parsedRows, file.name);
                } catch (err) {
                    console.error("Lỗi parse Word PPCT:", err);
                    alert("Lỗi đọc file Word: " + err.message);
                }
            };
            reader.readAsArrayBuffer(file);
        } else if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls") || fileName.endsWith(".csv")) {
            if (typeof XLSX === "undefined") {
                alert("Thư viện SheetJS chưa sẵn sàng!");
                return;
            }

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // Priority: Check if CTLOP sheet exists
                    const ctlopSheet = workbook.Sheets['CTLOP'] || workbook.Sheets['ctlop'];
                    if (ctlopSheet) {
                        const jsonRows = XLSX.utils.sheet_to_json(ctlopSheet, { header: 1 });
                        const parsedRows = [];
                        let currentW = 1;

                        for (let r = 5; r < jsonRows.length; r++) {
                            const row = jsonRows[r];
                            if (!row || !row.length) continue;
                            const subRaw = row[2] ? String(row[2]).trim() : '';
                            const weekVal = row[3] !== undefined ? parseInt(row[3]) : null;
                            const periodInWeek = row[4] !== undefined ? parseInt(row[4]) : 1;
                            const ppctNum = row[5] !== undefined ? parseInt(row[5]) : null;
                            const lesson = row[6] ? String(row[6]).trim() : '';
                            const dur = row[7] ? String(row[7]).trim() : '';

                            if (!subRaw || subRaw === 'Môn' || subRaw.includes('PHÂN PHỐI') || subRaw === 'Phân môn') continue;
                            if (!ppctNum && !lesson) continue;

                            if (weekVal && !isNaN(weekVal) && weekVal >= 1 && weekVal <= 35) {
                                currentW = weekVal;
                            } else if (currentW > 35) {
                                continue;
                            }

                            const normSub = normalizeSubjectName(subRaw);
                            const lessonClean = normalizePunctuationSpacing(lesson || 'Bài học');

                            parsedRows.push({
                                week: currentW,
                                subject: normSub,
                                periodInWeek: isNaN(periodInWeek) ? 1 : periodInWeek,
                                ppct: ppctNum || (parsedRows.length + 1),
                                lessonName: lessonClean,
                                duration: dur,
                                integration: ''
                            });
                        }

                        if (parsedRows.length === 0) {
                            alert("Không tìm thấy dữ liệu bài dạy hợp lệ trong sheet CTLOP!");
                            return;
                        }

                        promptImportModeModal(parsedRows, file.name);
                        return;
                    }

                    // Fallback to standard sheet
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    if (!jsonRows || jsonRows.length <= 1) {
                        alert("File Excel không có dữ liệu!");
                        return;
                    }

                    let headerRowIdx = -1;
                    let bestColMap = { week: -1, subject: -1, periodInWeek: -1, ppct: -1, lesson: -1, integration: -1 };
                    let maxMatchCount = 0;

                    for (let r = 0; r < Math.min(10, jsonRows.length); r++) {
                        const row = jsonRows[r] || [];
                        const currentMap = { week: -1, subject: -1, periodInWeek: -1, ppct: -1, lesson: -1, integration: -1 };
                        let matchScore = 0;

                        for (let c = 0; c < row.length; c++) {
                            const matchedType = matchPpctColumn(row[c]);
                            if (matchedType && currentMap[matchedType] === -1) {
                                currentMap[matchedType] = c;
                                matchScore++;
                            }
                        }

                        if (matchScore > maxMatchCount) {
                            maxMatchCount = matchScore;
                            bestColMap = currentMap;
                            headerRowIdx = r;
                        }
                    }

                    // Validate mandatory columns
                    const missingCols = checkMissingPpctColumns(bestColMap);
                    if (missingCols.length > 0) {
                        showMissingColumnsAlert("Excel (.xlsx)", missingCols);
                        return;
                    }

                    const startRow = headerRowIdx !== -1 ? headerRowIdx + 1 : 1;
                    const parsedRows = [];
                    let currentW = 1;
                    let currentSub = "Tiếng Việt";

                    for (let r = startRow; r < jsonRows.length; r++) {
                        const row = jsonRows[r];
                        if (!row || row.length === 0) continue;

                        const wVal = row[bestColMap.week];
                        if (wVal && !isNaN(parseInt(wVal))) currentW = parseInt(wVal);

                        const subVal = row[bestColMap.subject];
                        if (subVal && subVal.toString().trim()) {
                            currentSub = normalizeSubjectName(subVal.toString().trim());
                        }

                        const lessonVal = row[bestColMap.lesson] ? normalizePunctuationSpacing(row[bestColMap.lesson].toString().trim()) : "";
                        if (!lessonVal || lessonVal.toLowerCase().includes("tên bài") || lessonVal.toLowerCase().includes("bài học")) continue;

                        let pInWeekVal = row[bestColMap.periodInWeek] ? parseInt(row[bestColMap.periodInWeek]) : 1;
                        if (isNaN(pInWeekVal) || pInWeekVal <= 0) pInWeekVal = 1;

                        let ppctVal = row[bestColMap.ppct] ? parseInt(row[bestColMap.ppct]) : (parsedRows.length + 1);
                        if (isNaN(ppctVal)) ppctVal = parsedRows.length + 1;

                        const integVal = bestColMap.integration !== -1 && row[bestColMap.integration] 
                            ? normalizePunctuationSpacing(row[bestColMap.integration].toString().trim()) 
                            : "";

                        parsedRows.push({
                            week: currentW,
                            subject: currentSub,
                            periodInWeek: pInWeekVal,
                            ppct: ppctVal,
                            lessonName: lessonVal,
                            integration: integVal
                        });
                    }

                    promptImportModeModal(parsedRows, file.name);
                } catch (err) {
                    console.error("Lỗi đọc file Excel:", err);
                    alert("Lỗi đọc file Excel: " + err.message);
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            alert("Vui lòng chọn file Word (.docx) hoặc Excel (.xlsx, .xls, .csv)!");
        }
    }

    // Render Tab 4: Phân Phối Chương Trình (PPCT & Xuất KHDH)
    function renderTabPpct() {
        const chipsContainer = document.getElementById("khdh-subject-chips");
        if (chipsContainer) {
            const availableKhdhSubs = Object.keys(state.khdh);
            chipsContainer.innerHTML = availableKhdhSubs.map(sub => {
                const isSelected = state.selectedKhdhSubjects.includes(sub);
                return `
                    <label class="subject-chip ${isSelected ? 'selected' : ''}" data-subject="${sub}">
                        <input type="checkbox" value="${sub}" ${isSelected ? 'checked' : ''}>
                        <span>${isSelected ? '✓ ' : '+ '}${sub}</span>
                    </label>
                `;
            }).join('');

            chipsContainer.querySelectorAll(".subject-chip").forEach(chip => {
                chip.addEventListener("click", (e) => {
                    e.preventDefault();
                    const sub = chip.dataset.subject;
                    if (state.selectedKhdhSubjects.includes(sub)) {
                        state.selectedKhdhSubjects = state.selectedKhdhSubjects.filter(s => s !== sub);
                    } else {
                        state.selectedKhdhSubjects.push(sub);
                    }
                    renderTabPpct();
                });
            });
        }

        renderTabPpctBases();

        const subSelect = document.getElementById("ppct-filter-subject");
        const weekSelect = document.getElementById("ppct-filter-week");
        const searchInput = document.getElementById("ppct-search");

        const currentSelectedSub = subSelect ? subSelect.value : "ALL";
        if (subSelect) {
            const allSubs = getAllUniqueSubjects();
            subSelect.innerHTML = `<option value="ALL">-- Tất cả các môn --</option>` + 
                allSubs.map(s => `<option value="${s}" ${s === currentSelectedSub ? 'selected' : ''}>${s}</option>`).join('');
        }

        const selectedSubject = subSelect ? subSelect.value : "ALL";
        const selectedWeek = weekSelect ? weekSelect.value : "ALL";
        const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : "";

        let filtered = state.ppct;
        if (selectedSubject !== "ALL") {
            filtered = filtered.filter(p => normalizeSubjectName(p.subject) === normalizeSubjectName(selectedSubject) || p.subject === selectedSubject);
        }
        if (selectedWeek !== "ALL") {
            filtered = filtered.filter(p => p.week === parseInt(selectedWeek));
        }
        if (searchTerm) {
            filtered = filtered.filter(p => p.lessonName.toLowerCase().includes(searchTerm) || (p.integration && p.integration.toLowerCase().includes(searchTerm)));
        }

        document.getElementById("ppct-count-badge").innerText = `${filtered.length} tiết`;

        const tbody = document.getElementById("ppct-table-body");
        tbody.innerHTML = "";

        const displayRows = filtered;
        displayRows.forEach(p => {
            const globalIdx = state.ppct.indexOf(p);
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td class="editable-cell ppct-cell-edit" contenteditable="true" data-index="${globalIdx}" data-field="week" title="Nhấp để sửa tuần" style="text-align:center; font-weight:bold;">${p.week}</td>
                <td class="editable-cell ppct-cell-edit" contenteditable="true" data-index="${globalIdx}" data-field="subject" title="Nhấp để sửa môn" style="font-weight:600; color:var(--primary);">${p.subject}</td>
                <td class="editable-cell ppct-cell-edit" contenteditable="true" data-index="${globalIdx}" data-field="periodInWeek" title="Nhấp để sửa tiết trong tuần" style="text-align:center;">${p.periodInWeek}</td>
                <td class="editable-cell ppct-cell-edit" contenteditable="true" data-index="${globalIdx}" data-field="ppct" title="Nhấp để sửa số tiết PPCT" style="text-align:center; font-weight:bold; color:var(--primary);">${p.ppct}</td>
                <td class="editable-cell ppct-cell-edit" contenteditable="true" data-index="${globalIdx}" data-field="lessonName" title="Nhấp để sửa tên bài dạy" style="font-weight:500;">${p.lessonName}</td>
                <td class="editable-cell ppct-cell-edit" contenteditable="true" data-index="${globalIdx}" data-field="integration" title="Nhấp để sửa nội dung tích hợp / điều chỉnh" style="font-size:0.86rem; color:var(--text-main);">${p.integration || ''}</td>
                <td style="text-align: center;">
                    <div class="row-actions-group">
                        <button class="btn-row-action btn-insert-row" data-index="${globalIdx}" title="Chèn thêm 1 dòng ngay phía dưới">➕</button>
                        <button class="btn-row-action btn-delete-row" data-index="${globalIdx}" title="Xóa dòng bài dạy này">🗑️</button>
                    </div>
                </td>
            `;
            tbody.appendChild(tr);
        });

        // Bind editable cell events for PPCT
        tbody.querySelectorAll(".ppct-cell-edit").forEach(cell => {
            cell.addEventListener("blur", (e) => {
                const idx = parseInt(e.target.dataset.index);
                const field = e.target.dataset.field;
                let val = e.target.innerText.trim();

                if (idx >= 0 && idx < state.ppct.length) {
                    if (field === "week" || field === "periodInWeek" || field === "ppct") {
                        val = parseInt(val) || 1;
                    } else if (field === "subject") {
                        val = normalizeSubjectName(val);
                    } else if (field === "lessonName" || field === "integration") {
                        val = normalizePunctuationSpacing(val);
                        e.target.innerText = val;
                    }
                    state.ppct[idx][field] = val;
                    saveState();
                }
            });

            cell.addEventListener("keydown", (e) => {
                if (e.key === "Enter" && !e.shiftKey) {
                    e.preventDefault();
                    cell.blur();
                }
            });
        });

        // Bind insert row buttons
        tbody.querySelectorAll(".btn-insert-row").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const idx = parseInt(e.currentTarget.dataset.index);
                if (idx >= 0 && idx < state.ppct.length) {
                    const cur = state.ppct[idx];
                    const newRow = {
                        week: cur ? cur.week : 1,
                        subject: cur ? cur.subject : "Tiếng Việt",
                        periodInWeek: cur ? (parseInt(cur.periodInWeek) + 1) : 1,
                        ppct: cur ? (parseInt(cur.ppct) + 1) : (state.ppct.length + 1),
                        lessonName: "Nhập tên bài dạy mới...",
                        integration: ""
                    };
                    state.ppct.splice(idx + 1, 0, newRow);
                    saveState();
                    renderTabPpct();
                    renderTabLbg();
                    renderTabCtlop();
                    showToast("Đã chèn thêm 1 dòng tiết học mới!", "success");
                }
            });
        });

        // Bind delete row buttons
        tbody.querySelectorAll(".btn-delete-row").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const idx = parseInt(e.currentTarget.dataset.index);
                if (idx >= 0 && idx < state.ppct.length) {
                    const cur = state.ppct[idx];
                    if (confirm(`Bạn có chắc chắn muốn xóa tiết PPCT ${cur.ppct}: "${cur.lessonName}"?`)) {
                        state.ppct.splice(idx, 1);
                        saveState();
                        renderTabPpct();
                        renderTabLbg();
                        renderTabCtlop();
                        showToast("Đã xóa 1 dòng tiết học!", "info");
                    }
                }
            });
        });
    }

    // Bulk Rename Subject functions
    function openRenameModal() {
        const modal = document.getElementById("rename-modal");
        const select = document.getElementById("rename-old-subject");
        const input = document.getElementById("rename-new-subject");
        if (!modal || !select || !input) return;

        const allSubs = getAllUniqueSubjects();
        select.innerHTML = allSubs.map(s => `<option value="${s}">${s}</option>`).join('');
        input.value = "";
        modal.classList.add("active");
    }

    function closeRenameModal() {
        const modal = document.getElementById("rename-modal");
        if (modal) modal.classList.remove("active");
    }

    function applyBulkRenameSubject() {
        const select = document.getElementById("rename-old-subject");
        const input = document.getElementById("rename-new-subject");
        const syncCb = document.getElementById("rename-sync-timetable");

        const oldName = select ? select.value.trim() : "";
        const newName = input ? input.value.trim() : "";
        const syncTimetable = syncCb ? syncCb.checked : true;

        if (!oldName) {
            alert("Vui lòng chọn môn học cần đổi tên!");
            return;
        }
        if (!newName) {
            alert("Vui lòng nhập tên môn học mới!");
            return;
        }
        if (oldName === newName) {
            alert("Tên mới trùng với tên cũ!");
            return;
        }

        let ppctChangedCount = 0;
        const normOld = normalizeSubjectName(oldName);

        // 1. Rename in PPCT
        state.ppct.forEach(p => {
            if (p.subject === oldName || normalizeSubjectName(p.subject) === normOld) {
                p.subject = newName;
                ppctChangedCount++;
            }
        });

        // 2. Rename in Timetable & Overrides & KHDH if requested
        if (syncTimetable) {
            if (state.timetable) {
                state.timetable.forEach(t => {
                    if (t.subject === oldName || normalizeSubjectName(t.subject) === normOld) {
                        t.subject = newName;
                    }
                });
            }

            if (state.weeklyScheduleOverrides) {
                Object.keys(state.weeklyScheduleOverrides).forEach(key => {
                    const ov = state.weeklyScheduleOverrides[key];
                    if (ov.subject && (ov.subject === oldName || normalizeSubjectName(ov.subject) === normOld)) {
                        ov.subject = newName;
                    }
                });
            }

            if (state.khdh && state.khdh[oldName]) {
                state.khdh[newName] = state.khdh[oldName];
                delete state.khdh[oldName];
            }
        }

        // 3. Update included subjects list
        if (state.includedSubjects) {
            state.includedSubjects = state.includedSubjects.map(s => (s === oldName || normalizeSubjectName(s) === normOld) ? newName : s);
        }

        // 4. Update selected KHDH subjects list
        if (state.selectedKhdhSubjects) {
            state.selectedKhdhSubjects = state.selectedKhdhSubjects.map(s => (s === oldName || normalizeSubjectName(s) === normOld) ? newName : s);
        }

        saveState();
        closeRenameModal();
        renderTabPpct();
        renderLbgInclusionChips();
        renderTabLbg();
        renderTabCtlop();
        showToast(`Đã đổi tên môn "${oldName}" thành "${newName}" cho toàn bộ ${ppctChangedCount} tiết học!`, "success");
    }

    // Export Lịch Tuần (35 tuần) to Excel
    function exportLichtuanToExcel() {
        if (typeof XLSX === "undefined") {
            alert("Thư viện Excel chưa sẵn sàng!");
            return;
        }

        const wb = XLSX.utils.book_new();
        const titleText = `KẾ HOẠCH THỜI GIAN NĂM HỌC ${(state.settings.academicYear || "2026 - 2027").toUpperCase()} (35 TUẦN THỰC HỌC)`;
        
        const wsData = [
            [state.settings.governingBody || "UBND PHƯỜNG TRUNG NHỨT", "", "", "", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"],
            [state.settings.schoolName || "TRƯỜNG TIỂU HỌC TRUNG NHỨT", "", "", "", "Độc lập - Tự do - Hạnh phúc"],
            [`${state.settings.grade || "KHỐI 5"} - ${state.settings.className || "LỚP 5A"}`],
            [""],
            ["", "", titleText],
            [""],
            ["Học kỳ", "Tuần", "Từ ngày (dd/mm/yyyy)", "Đến ngày (dd/mm/yyyy)", "Tháng", "Số ngày lễ trong tháng (Luật LĐ)", "Ghi chú"]
        ];

        state.weeks.forEach(w => {
            wsData.push([
                w.semester || (w.week <= 18 ? "Học kỳ 1" : "Học kỳ 2"),
                `Tuần ${w.week}`,
                w.startDateVN || "",
                w.endDateVN || "",
                w.month || "",
                w.holidays || "",
                w.notes || ""
            ]);
        });

        const ws = XLSX.utils.aoa_to_sheet(wsData);
        ws['!cols'] = [
            { wch: 14 },
            { wch: 12 },
            { wch: 22 },
            { wch: 22 },
            { wch: 16 },
            { wch: 34 },
            { wch: 32 }
        ];

        XLSX.utils.book_append_sheet(wb, ws, "Lich_35_Tuan");
        const filename = `Ke_Hoach_Thoi_Gian_35_Tuan_${(state.settings.academicYear || '2026_2027').replace(/\s+/g, '')}.xlsx`;
        XLSX.writeFile(wb, filename);
        showToast(`Đã xuất file Excel Kế hoạch 35 tuần: ${filename}`, "success");
    }

    // Bases (Căn cứ pháp lý) Management
    function renderTabPpctBases() {
        const textarea = document.getElementById("textarea-khdh-bases");
        if (!textarea) return;

        if (!Array.isArray(state.settings.bases) || state.settings.bases.length === 0 || state.settings.bases.some(b => b.includes("28/2020"))) {
            state.settings.bases = [...DEFAULT_BASES];
        }

        textarea.value = state.settings.bases.join('\n');
    }

    function saveBasesFromTextarea() {
        const textarea = document.getElementById("textarea-khdh-bases");
        if (!textarea) return;

        const lines = textarea.value.split('\n').map(l => normalizePunctuationSpacing(l.trim())).filter(Boolean);
        if (lines.length === 0) {
            alert("Vui lòng nhập ít nhất 1 căn cứ!");
            return;
        }

        state.settings.bases = lines;
        saveState();
        showToast("Đã lưu danh sách căn cứ xây dựng KHDH thành công!", "success");
    }

    function importBasesFromFile(file) {
        if (!file) return;
        const fileName = file.name.toLowerCase();

        if (fileName.endsWith('.txt')) {
            const reader = new FileReader();
            reader.onload = e => {
                const text = e.target.result;
                const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
                applyImportedBases(lines);
            };
            reader.readAsText(file, 'UTF-8');
        } else if (fileName.endsWith('.docx')) {
            if (typeof JSZip === 'undefined') {
                alert("Thư viện đọc file Word chưa sẵn sàng!");
                return;
            }
            const reader = new FileReader();
            reader.onload = e => {
                const zip = new JSZip();
                zip.loadAsync(e.target.result).then(doc => {
                    return doc.file("word/document.xml").async("string");
                }).then(xmlStr => {
                    const parser = new DOMParser();
                    const xmlDoc = parser.parseFromString(xmlStr, "application/xml");
                    const pNodes = xmlDoc.getElementsByTagName("w:p");
                    const extracted = [];
                    for (let p of pNodes) {
                        const tNodes = p.getElementsByTagName("w:t");
                        let pText = "";
                        for (let t of tNodes) {
                            pText += t.textContent;
                        }
                        pText = pText.trim();
                        if (pText.toLowerCase().startsWith("căn cứ") || pText.includes("Thông tư") || pText.includes("Quyết định") || pText.includes("Kế hoạch")) {
                            extracted.push(pText);
                        }
                    }
                    if (extracted.length > 0) {
                        applyImportedBases(extracted);
                    } else {
                        alert("Không tìm thấy đoạn văn bản 'Căn cứ...' trong file Word!");
                    }
                }).catch(err => {
                    alert("Lỗi đọc file Word: " + err.message);
                });
            };
            reader.readAsArrayBuffer(file);
        } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
            if (typeof XLSX === 'undefined') {
                alert("Thư viện đọc Excel chưa sẵn sàng!");
                return;
            }
            const reader = new FileReader();
            reader.onload = e => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const extracted = [];
                workbook.SheetNames.forEach(sheetName => {
                    const sheet = workbook.Sheets[sheetName];
                    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                    json.forEach(row => {
                        row.forEach(cell => {
                            if (cell && typeof cell === 'string') {
                                const val = cell.trim();
                                if (val.toLowerCase().startsWith("căn cứ") || val.includes("Thông tư") || val.includes("Quyết định")) {
                                    extracted.push(val);
                                }
                            }
                        });
                    });
                });
                if (extracted.length > 0) {
                    applyImportedBases(extracted);
                } else {
                    alert("Không tìm thấy nội dung 'Căn cứ...' trong file Excel!");
                }
            };
            reader.readAsArrayBuffer(file);
        }
    }

    function applyImportedBases(lines) {
        state.settings.bases = (lines || []).map(l => normalizePunctuationSpacing(l.trim())).filter(Boolean);
        renderTabPpctBases();
        saveState();
        showToast(`Đã nạp thành công ${state.settings.bases.length} căn cứ từ file!`, "success");
    }

    function escapeHtml(unsafe) {
        if (!unsafe) return "";
        return unsafe.toString()
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    function getBghSignerInfo(docType = "LBG") {
        const s = state.settings;
        let vps = s.vicePrincipals;
        if (!Array.isArray(vps) || vps.length === 0) {
            vps = [s.vicePrincipal || "Lê Văn Tám"];
            s.vicePrincipals = vps;
        }

        const isKhdh = (docType === "KHDH");
        const signerType = isKhdh
            ? (s.bghSignerKhdhType || "HT")
            : (s.bghSignerLbgType || s.bghSignerType || "PHT");
        
        const signerIdx = isKhdh
            ? (s.bghSignerKhdhIndex !== undefined ? s.bghSignerKhdhIndex : 0)
            : (s.bghSignerLbgIndex !== undefined ? s.bghSignerLbgIndex : (s.bghSignerIndex || 0));

        if (signerType === "HT") {
            return {
                role: "HIỆU TRƯỞNG",
                title: "HIỆU TRƯỞNG",
                name: s.principal || "Phạm Quốc Hùng",
                type: "HT",
                index: 0
            };
        } else {
            const idx = (signerIdx >= 0 && signerIdx < vps.length) ? signerIdx : 0;
            return {
                role: "PHÓ HIỆU TRƯỞNG",
                title: isKhdh ? "PHÓ HIỆU TRƯỞNG" : "BAN GIÁM HIỆU",
                name: vps[idx] || "Lê Văn Tám",
                type: "PHT",
                index: idx
            };
        }
    }

    function updateTopHeaderBadge() {
        const badge = document.getElementById("top-badge-school");
        if (!badge) return;
        const school = (state.settings.schoolName || "TRƯỜNG TIỂU HỌC").trim();
        const className = (state.settings.className || "LỚP 5A").trim();
        
        let shortSchool = school;
        if (shortSchool.toUpperCase().startsWith("TRƯỜNG TIỂU HỌC ")) {
            shortSchool = "TH " + shortSchool.substring(16);
        } else if (shortSchool.toUpperCase().startsWith("TIỂU HỌC ")) {
            shortSchool = "TH " + shortSchool.substring(9);
        }
        badge.innerText = `🏫 ${shortSchool.toUpperCase()} — ${className.toUpperCase()}`;
    }

    // =========================================================================
    // SUBJECT MANAGEMENT (PHÂN HỆ QUẢN LÝ MÔN HỌC - VERSION 3.0)
    // =========================================================================
    function renderSubjectManagementTable() {
        const tbody = document.getElementById("subject-management-tbody");
        if (!tbody) return;

        if (!state.subjectList || state.subjectList.length === 0) {
            state.subjectList = JSON.parse(JSON.stringify(DEFAULT_GRADE_5_SUBJECTS));
        }

        // Count statistics
        let gvcnCount = 0;
        let specialistCount = 0;
        let enhancedCount = 0;
        state.subjectList.forEach(sub => {
            if (sub.category === "GVCN") gvcnCount++;
            else if (sub.category === "Chuyên trách") specialistCount++;
            else if (sub.category === "Tăng cường") enhancedCount++;
        });

        const totalPpctCount = state.ppct ? state.ppct.length : 0;

        const statGvcnEl = document.getElementById("stat-count-gvcn");
        if (statGvcnEl) statGvcnEl.innerText = `${gvcnCount} môn`;
        const statSpecEl = document.getElementById("stat-count-specialist");
        if (statSpecEl) statSpecEl.innerText = `${specialistCount} môn`;
        const statEnhEl = document.getElementById("stat-count-enhanced");
        if (statEnhEl) statEnhEl.innerText = `${enhancedCount} môn`;
        const statPpctEl = document.getElementById("stat-count-total-ppct");
        if (statPpctEl) statPpctEl.innerText = `${totalPpctCount.toLocaleString('vi-VN')} tiết`;

        tbody.innerHTML = "";

        state.subjectList.forEach((sub, idx) => {
            const tr = document.createElement("tr");

            let badgeClass = "badge-cat-gvcn";
            let badgeIcon = "👨‍🏫";
            let badgeText = "GVCN Dạy";
            if (sub.category === "Chuyên trách") {
                badgeClass = "badge-cat-specialist";
                badgeIcon = "🎨";
                badgeText = "GV Chuyên Trách";
            } else if (sub.category === "Tăng cường") {
                badgeClass = "badge-cat-enhanced";
                badgeIcon = "⚡";
                badgeText = "Tăng Cường / Tự Chọn";
            }

            const normSub = normalizeSubjectName(sub.name);
            const ppctLessons = state.ppct ? state.ppct.filter(p => normalizeSubjectName(p.subject) === normSub).length : 0;
            const isIncl = (sub.isIncluded !== false);

            tr.innerHTML = `
                <td style="text-align: center; font-weight: 600; color: var(--text-muted);">${idx + 1}</td>
                <td style="text-align: left; font-weight: 700; color: var(--text-main); font-size: 0.92rem;">
                    ${escapeHtml(sub.name)}
                </td>
                <td style="text-align: center;">
                    <span class="badge-cat ${badgeClass}">
                        <span>${badgeIcon}</span> <span>${badgeText}</span>
                    </span>
                </td>
                <td style="text-align: center; font-weight: 700;">
                    <input type="number" min="1" max="15" value="${sub.defaultPeriods || 1}" class="form-control form-control-sm sub-period-input" data-index="${idx}" style="width: 65px; margin: 0 auto; text-align: center; font-weight: 700; padding: 0.2rem 0.4rem;">
                </td>
                <td style="text-align: center; font-weight: 600; color: ${ppctLessons > 0 ? 'var(--primary)' : '#94a3b8'};">
                    ${ppctLessons > 0 ? `${ppctLessons} tiết` : '--'}
                </td>
                <td style="text-align: center;">
                    <label class="toggle-switch-wrap" title="Bật/Tắt đưa môn này vào Lịch Báo Giảng">
                        <input type="checkbox" class="sub-toggle-incl" data-index="${idx}" ${isIncl ? 'checked' : ''}>
                    </label>
                </td>
                <td style="text-align: center;">
                    <div style="display: inline-flex; gap: 0.35rem;">
                        <button type="button" class="table-sub-action-btn btn-edit-sub" data-index="${idx}" title="Chỉnh sửa môn học">✏️</button>
                        <button type="button" class="table-sub-action-btn btn-del-sub" data-index="${idx}" title="Xóa môn học">🗑️</button>
                    </div>
                </td>
            `;

            tbody.appendChild(tr);
        });

        // Bind events
        tbody.querySelectorAll(".sub-period-input").forEach(inp => {
            inp.addEventListener("change", (e) => {
                const idx = parseInt(e.target.dataset.index);
                if (idx >= 0 && idx < state.subjectList.length) {
                    state.subjectList[idx].defaultPeriods = parseInt(e.target.value) || 1;
                    saveState();
                }
            });
        });

        tbody.querySelectorAll(".sub-toggle-incl").forEach(chk => {
            chk.addEventListener("change", (e) => {
                const idx = parseInt(e.target.dataset.index);
                if (idx >= 0 && idx < state.subjectList.length) {
                    const sub = state.subjectList[idx];
                    sub.isIncluded = e.target.checked;
                    
                    const norm = normalizeSubjectName(sub.name);
                    if (e.target.checked) {
                        if (!state.includedSubjects.includes(sub.name)) {
                            state.includedSubjects.push(sub.name);
                        }
                    } else {
                        state.includedSubjects = state.includedSubjects.filter(s => normalizeSubjectName(s) !== norm);
                    }

                    saveState();
                    renderLbgInclusionChips();
                    renderTabLbg();
                    renderTabCtlop();
                    showToast(`Đã ${e.target.checked ? 'bật đưa' : 'tắt đưa'} môn '${sub.name}' vào Lịch Báo Giảng!`, "success");
                }
            });
        });

        tbody.querySelectorAll(".btn-edit-sub").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const idx = parseInt(e.currentTarget.dataset.index);
                if (idx >= 0 && idx < state.subjectList.length) {
                    openSubjectEditorModal(state.subjectList[idx]);
                }
            });
        });

        tbody.querySelectorAll(".btn-del-sub").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const idx = parseInt(e.currentTarget.dataset.index);
                if (idx >= 0 && idx < state.subjectList.length) {
                    deleteSubject(state.subjectList[idx].name);
                }
            });
        });
    }

    function openSubjectEditorModal(subjectObj = null) {
        const modal = document.getElementById("modal-subject-editor");
        if (!modal) return;

        const titleEl = document.getElementById("modal-subject-title");
        const origInput = document.getElementById("modal-sub-editing-original-name");
        const nameInput = document.getElementById("modal-sub-name");
        const catSelect = document.getElementById("modal-sub-category");
        const perInput = document.getElementById("modal-sub-periods");
        const inclCheck = document.getElementById("modal-sub-included");

        if (subjectObj) {
            titleEl.innerText = "✏️ Chỉnh Sửa Môn Học";
            origInput.value = subjectObj.name;
            nameInput.value = subjectObj.name;
            catSelect.value = subjectObj.category || "GVCN";
            perInput.value = subjectObj.defaultPeriods || 1;
            inclCheck.checked = (subjectObj.isIncluded !== false);
        } else {
            titleEl.innerText = "➕ Thêm Môn Học Mới";
            origInput.value = "";
            nameInput.value = "";
            catSelect.value = "GVCN";
            perInput.value = 1;
            inclCheck.checked = true;
        }

        modal.classList.add("show");
    }

    function closeSubjectEditorModal() {
        const modal = document.getElementById("modal-subject-editor");
        if (modal) modal.classList.remove("show");
    }

    function saveSubjectFromModal() {
        const origName = document.getElementById("modal-sub-editing-original-name").value.trim();
        const newName = document.getElementById("modal-sub-name").value.trim();
        const category = document.getElementById("modal-sub-category").value;
        const periods = parseInt(document.getElementById("modal-sub-periods").value) || 1;
        const isIncluded = document.getElementById("modal-sub-included").checked;

        if (!newName) {
            alert("Vui lòng nhập tên môn học!");
            return;
        }

        const normNew = normalizeSubjectName(newName);

        if (!origName) {
            // Adding new subject
            const exists = state.subjectList.some(s => normalizeSubjectName(s.name) === normNew);
            if (exists) {
                alert(`Môn học '${newName}' đã tồn tại trong danh sách!`);
                return;
            }

            state.subjectList.push({
                name: newName,
                category: category,
                defaultPeriods: periods,
                isIncluded: isIncluded
            });

            if (isIncluded && !state.includedSubjects.includes(newName)) {
                state.includedSubjects.push(newName);
            }

            saveState();
            closeSubjectEditorModal();
            renderSubjectManagementTable();
            renderMasterTimetableEditor();
            renderTabLbg();
            renderTabCtlop();
            renderTabPpct();
            showToast(`Đã thêm môn học mới: ${newName}!`, "success");
        } else {
            // Editing existing subject
            const idx = state.subjectList.findIndex(s => s.name === origName);
            if (idx >= 0) {
                const normOrig = normalizeSubjectName(origName);
                if (normOrig !== normNew) {
                    // Name changed - check conflict
                    const exists = state.subjectList.some((s, i) => i !== idx && normalizeSubjectName(s.name) === normNew);
                    if (exists) {
                        alert(`Môn học '${newName}' đã tồn tại trong danh sách!`);
                        return;
                    }
                    bulkRenameSubjectInternal(origName, newName);
                }

                state.subjectList[idx].name = newName;
                state.subjectList[idx].category = category;
                state.subjectList[idx].defaultPeriods = periods;
                state.subjectList[idx].isIncluded = isIncluded;

                // Sync includedSubjects
                if (isIncluded) {
                    if (!state.includedSubjects.includes(newName)) state.includedSubjects.push(newName);
                } else {
                    state.includedSubjects = state.includedSubjects.filter(s => normalizeSubjectName(s) !== normNew);
                }

                saveState();
                closeSubjectEditorModal();
                renderSubjectManagementTable();
                renderMasterTimetableEditor();
                renderTabLbg();
                renderTabCtlop();
                renderTabPpct();
                showToast(`Đã cập nhật thông tin môn học: ${newName}!`, "success");
            }
        }
    }

    function deleteSubject(subjectName) {
        if (!confirm(`Bạn có chắc chắn muốn xóa môn '${subjectName}' khỏi danh mục môn học không?`)) {
            return;
        }

        const norm = normalizeSubjectName(subjectName);
        state.subjectList = state.subjectList.filter(s => normalizeSubjectName(s.name) !== norm);
        state.includedSubjects = state.includedSubjects.filter(s => normalizeSubjectName(s) !== norm);

        saveState();
        renderSubjectManagementTable();
        renderMasterTimetableEditor();
        renderTabLbg();
        renderTabCtlop();
        renderTabPpct();
        showToast(`Đã xóa môn '${subjectName}' khỏi danh sách!`, "success");
    }

    function resetDefaultSubjects() {
        if (!confirm("Bạn có chắc chắn muốn khôi phục danh mục môn học chuẩn theo Chương trình GDPT 2018 (Kết nối tri thức) không?")) {
            return;
        }

        state.subjectList = getGradeDefaultSubjects(state.currentGrade);
        state.includedSubjects = state.subjectList.filter(s => s.isIncluded !== false).map(s => s.name);

        saveState();
        renderSubjectManagementTable();
        renderMasterTimetableEditor();
        renderTabLbg();
        renderTabCtlop();
        renderTabPpct();
        showToast("Đã khôi phục danh mục môn học chuẩn thành công!", "success");
    }

    function openRenameSubjectModal() {
        const modal = document.getElementById("modal-rename-subject");
        if (!modal) return;

        const selectOld = document.getElementById("modal-rename-old-subject");
        const inputNew = document.getElementById("modal-rename-new-name");

        const allSubjects = getAllUniqueSubjects();
        selectOld.innerHTML = allSubjects.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');

        if (allSubjects.length > 0) {
            inputNew.value = allSubjects[0];
        }

        selectOld.onchange = (e) => {
            inputNew.value = e.target.value;
        };

        modal.classList.add("show");
    }

    function closeRenameSubjectModal() {
        const modal = document.getElementById("modal-rename-subject");
        if (modal) modal.classList.remove("show");
    }

    function bulkRenameSubjectInternal(oldName, newName) {
        const normOld = normalizeSubjectName(oldName);
        const normNew = normalizeSubjectName(newName);
        if (normOld === normNew) return;

        // 1. Rename in subjectList
        if (state.subjectList) {
            state.subjectList.forEach(s => {
                if (normalizeSubjectName(s.name) === normOld) {
                    s.name = newName;
                }
            });
        }

        // 2. Rename in timetable
        if (state.timetable) {
            state.timetable.forEach(t => {
                if (normalizeSubjectName(t.subject) === normOld) {
                    t.subject = newName;
                }
            });
        }

        // 3. Rename in weeklyCustomSlots
        if (state.weeklyCustomSlots) {
            for (let w in state.weeklyCustomSlots) {
                state.weeklyCustomSlots[w].forEach(s => {
                    if (normalizeSubjectName(s.subject) === normOld) {
                        s.subject = newName;
                    }
                });
            }
        }

        // 4. Rename in weeklyScheduleOverrides
        if (state.weeklyScheduleOverrides) {
            for (let k in state.weeklyScheduleOverrides) {
                if (state.weeklyScheduleOverrides[k].subject && normalizeSubjectName(state.weeklyScheduleOverrides[k].subject) === normOld) {
                    state.weeklyScheduleOverrides[k].subject = newName;
                }
            }
        }

        // 5. Rename in ppct
        if (state.ppct) {
            state.ppct.forEach(p => {
                if (normalizeSubjectName(p.subject) === normOld) {
                    p.subject = newName;
                }
            });
        }

        // 6. Rename in khdh
        if (state.khdh && state.khdh[oldName]) {
            state.khdh[newName] = state.khdh[oldName];
            delete state.khdh[oldName];
        }

        // 7. Rename in includedSubjects
        if (state.includedSubjects) {
            state.includedSubjects = state.includedSubjects.map(s => normalizeSubjectName(s) === normOld ? newName : s);
        }
    }

    function applyBulkRenameSubject() {
        const oldName = document.getElementById("modal-rename-old-subject").value.trim();
        const newName = document.getElementById("modal-rename-new-name").value.trim();

        if (!oldName || !newName) {
            alert("Vui lòng nhập đầy đủ tên môn học cũ và tên môn học mới!");
            return;
        }

        if (oldName === newName) {
            alert("Tên môn mới trùng với tên môn cũ!");
            return;
        }

        bulkRenameSubjectInternal(oldName, newName);
        saveState();
        closeRenameSubjectModal();
        renderSubjectManagementTable();
        renderMasterTimetableEditor();
        renderTabLbg();
        renderTabCtlop();
        renderTabPpct();
        showToast(`Đã đổi tên môn '${oldName}' thành '${newName}' trên toàn bộ hệ thống!`, "success");
    }

    // Expose subject management helpers to window for global access
    window.openSubjectEditorModal = openSubjectEditorModal;
    window.closeSubjectEditorModal = closeSubjectEditorModal;
    window.saveSubjectFromModal = saveSubjectFromModal;
    window.openRenameSubjectModal = openRenameSubjectModal;
    window.closeRenameSubjectModal = closeRenameSubjectModal;
    window.applyBulkRenameSubject = applyBulkRenameSubject;
    window.resetDefaultSubjects = resetDefaultSubjects;
    window.renderSubjectManagementTable = renderSubjectManagementTable;

    function syncAcademicYear(newYear) {
        if (!newYear || !newYear.trim()) return;
        const cleanYear = newYear.trim();
        state.settings.academicYear = cleanYear;

        // Auto-update bases in settings
        if (Array.isArray(state.settings.bases)) {
            state.settings.bases = state.settings.bases.map(b => {
                return b.replace(/năm học \d{4}\s*[-–—]\s*\d{4}/gi, `năm học ${cleanYear}`)
                        .replace(/NĂM HỌC \d{4}\s*[-–—]\s*\d{4}/g, `NĂM HỌC ${cleanYear}`);
            });
        }

        // Auto-update starting year in date string if present (e.g. "ngày 28 tháng 8 năm 2026")
        const m = cleanYear.match(/^(\d{4})/);
        if (m && state.settings.dateString) {
            state.settings.dateString = state.settings.dateString.replace(/năm \d{4}/g, `năm ${m[1]}`);
            const dateInput = document.getElementById("set-date");
            if (dateInput) dateInput.value = state.settings.dateString;
        }

        const inputYear = document.getElementById("set-academic-year");
        if (inputYear) inputYear.value = cleanYear;
        const quickSelect = document.getElementById("set-academic-year-quick");
        if (quickSelect) quickSelect.value = cleanYear;

        saveState();
        showToast(`Đã đồng bộ Năm học ${cleanYear} sang KHDH, các căn cứ pháp lý và toàn hệ thống!`, "success");
        renderTabSettings();
        renderTabLbg();
        renderTabCtlop();
    }

    // Render Tab 5: Cài đặt
    function renderTabSettings() {
        document.getElementById("set-governing-body").value = state.settings.governingBody || "";
        document.getElementById("set-school-name").value = state.settings.schoolName || "";
        document.getElementById("set-grade").value = state.settings.grade || "KHỐI 5";
        document.getElementById("set-class-name").value = state.settings.className || "LỚP 5A";
        document.getElementById("set-academic-year").value = state.settings.academicYear || "2026 - 2027";
        const quickSelect = document.getElementById("set-academic-year-quick");
        if (quickSelect) {
            quickSelect.value = state.settings.academicYear || "2026 - 2027";
            if (!quickSelect.value) quickSelect.value = "";
        }
        document.getElementById("set-teacher").value = state.settings.homeroomTeacher || "";
        document.getElementById("set-head").value = state.settings.headOfGrade || "";
        document.getElementById("set-principal").value = state.settings.principal || "Phạm Quốc Hùng";
        document.getElementById("set-location").value = state.settings.location || "Trung Nhứt";
        document.getElementById("set-date").value = state.settings.dateString || "ngày 28 tháng 8 năm 2026";

        renderPhtListInputs();
        renderBghSignerDropdown();
        renderSubjectManagementTable();
        renderMasterTimetableEditor();
        updateTopHeaderBadge();

        const principalInput = document.getElementById("set-principal");
        if (principalInput && !principalInput.dataset.bound) {
            principalInput.dataset.bound = "true";
            const handlePrincipalChange = (e) => {
                state.settings.principal = e.target.value.trim();
                renderBghSignerDropdown();
                renderTabLbg();
                renderTabCtlop();
                renderTabLbgMon();
            };
            principalInput.addEventListener("input", handlePrincipalChange);
            principalInput.addEventListener("change", handlePrincipalChange);
        }
    }

    function renderPhtListInputs() {
        const container = document.getElementById("pht-list-container");
        if (!container) return;

        let vps = state.settings.vicePrincipals;
        if (!Array.isArray(vps) || vps.length === 0) {
            vps = [state.settings.vicePrincipal || "Lê Văn Tám"];
            state.settings.vicePrincipals = vps;
        }

        container.innerHTML = vps.map((phtName, idx) => `
            <div class="pht-input-row" style="display: flex; gap: 0.5rem; align-items: center;">
                <span style="font-size: 0.82rem; font-weight: 600; color: var(--text-muted); width: 60px; white-space: nowrap;">PHT ${idx + 1}:</span>
                <input type="text" class="form-control form-control-sm pht-name-input" data-index="${idx}" value="${escapeHtml(phtName)}" placeholder="Nhập họ tên Phó Hiệu trưởng ${idx + 1}...">
                <button type="button" class="btn btn-sm btn-delete-pht" data-index="${idx}" title="Xóa PHT này" style="padding: 0.2rem 0.5rem; color: #dc2626; border: 1px solid #fca5a5; background: #fff; cursor: pointer;">🗑️</button>
            </div>
        `).join('');

        // Bind PHT inputs
        container.querySelectorAll(".pht-name-input").forEach(input => {
            input.addEventListener("input", (e) => {
                const idx = parseInt(e.target.dataset.index);
                state.settings.vicePrincipals[idx] = e.target.value.trim();
                if (idx === 0) state.settings.vicePrincipal = state.settings.vicePrincipals[idx];
                renderBghSignerDropdown();
            });
        });

        // Bind delete buttons
        container.querySelectorAll(".btn-delete-pht").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const idx = parseInt(e.currentTarget.dataset.index);
                if (state.settings.vicePrincipals.length <= 1) {
                    alert("Cần có ít nhất 1 Phó Hiệu trưởng trong danh sách!");
                    return;
                }
                state.settings.vicePrincipals.splice(idx, 1);
                state.settings.vicePrincipal = state.settings.vicePrincipals[0] || "";
                if (state.settings.bghSignerLbgIndex >= state.settings.vicePrincipals.length) {
                    state.settings.bghSignerLbgIndex = 0;
                }
                if (state.settings.bghSignerKhdhIndex >= state.settings.vicePrincipals.length) {
                    state.settings.bghSignerKhdhIndex = 0;
                }
                renderPhtListInputs();
                renderBghSignerDropdown();
            });
        });
    }

    function renderBghSignerDropdown() {
        const selectLbg = document.getElementById("set-bgh-signer-lbg-select");
        const selectKhdh = document.getElementById("set-bgh-signer-khdh-select");

        const htName = state.settings.principal || "Phạm Quốc Hùng";
        const vps = state.settings.vicePrincipals || ["Lê Văn Tám"];

        // 1. LBG Signer Select
        if (selectLbg) {
            const curVal = selectLbg.value;
            const lbgType = state.settings.bghSignerLbgType || state.settings.bghSignerType || "PHT";
            const lbgIdx = (state.settings.bghSignerLbgIndex !== undefined) ? state.settings.bghSignerLbgIndex : (state.settings.bghSignerIndex || 0);
            const isHTSelected = curVal ? (curVal === 'HT:0') : (lbgType === 'HT');

            let lbgHtml = `<option value="HT:0" ${isHTSelected ? 'selected' : ''}>👨‍💼 Hiệu trưởng: ${escapeHtml(htName)}</option>`;
            vps.forEach((pht, idx) => {
                const isSel = curVal ? (curVal === `PHT:${idx}`) : (lbgType === 'PHT' && lbgIdx === idx);
                lbgHtml += `<option value="PHT:${idx}" ${isSel ? 'selected' : ''}>👥 Phó Hiệu trưởng ${idx + 1}: ${escapeHtml(pht || ('PHT ' + (idx + 1)))}</option>`;
            });
            selectLbg.innerHTML = lbgHtml;
            selectLbg.onchange = (e) => {
                const [type, idxStr] = e.target.value.split(':');
                state.settings.bghSignerLbgType = type;
                state.settings.bghSignerLbgIndex = parseInt(idxStr) || 0;
                state.settings.bghSignerType = type;
                state.settings.bghSignerIndex = parseInt(idxStr) || 0;
                saveState();
                renderTabLbg();
                renderTabCtlop();
                renderTabLbgMon();
            };
        }

        // 2. KHDH Signer Select
        if (selectKhdh) {
            const curVal = selectKhdh.value;
            const khdhType = state.settings.bghSignerKhdhType || "HT";
            const khdhIdx = (state.settings.bghSignerKhdhIndex !== undefined) ? state.settings.bghSignerKhdhIndex : 0;
            const isHTSelected = curVal ? (curVal === 'HT:0') : (khdhType === 'HT');

            let khdhHtml = `<option value="HT:0" ${isHTSelected ? 'selected' : ''}>👨‍💼 Hiệu trưởng: ${escapeHtml(htName)}</option>`;
            vps.forEach((pht, idx) => {
                const isSel = curVal ? (curVal === `PHT:${idx}`) : (khdhType === 'PHT' && khdhIdx === idx);
                khdhHtml += `<option value="PHT:${idx}" ${isSel ? 'selected' : ''}>👥 Phó Hiệu trưởng ${idx + 1}: ${escapeHtml(pht || ('PHT ' + (idx + 1)))}</option>`;
            });
            selectKhdh.innerHTML = khdhHtml;
            selectKhdh.onchange = (e) => {
                const [type, idxStr] = e.target.value.split(':');
                state.settings.bghSignerKhdhType = type;
                state.settings.bghSignerKhdhIndex = parseInt(idxStr) || 0;
                saveState();
            };
        }
    }

    function restoreGrade5DefaultIntegration() {
        if (state.currentGrade !== 5) {
            showToast("Chức năng này áp dụng cho Khối lớp 5.", "info");
            return;
        }
        if (!window.APP_INITIAL_DATA || !Array.isArray(window.APP_INITIAL_DATA.ppct)) {
            alert("Không tìm thấy dữ liệu tích hợp gốc của Khối 5!");
            return;
        }
        const defaultPpct = window.APP_INITIAL_DATA.ppct;
        let countRestored = 0;
        state.ppct.forEach(item => {
            const normSub = normalizeSubjectName(item.subject);
            const found = defaultPpct.find(d => 
                d.week === item.week && 
                normalizeSubjectName(d.subject) === normSub && 
                (d.ppct === item.ppct || d.periodInWeek === item.periodInWeek)
            );
            if (found && found.integration && found.integration.trim()) {
                if (item.integration !== found.integration) {
                    item.integration = found.integration;
                    countRestored++;
                }
            }
        });
        saveState();
        renderTabPpct();
        renderTabCtlop();
        renderTabLbgMon();
        showToast(`Đã phục hồi thành công ${countRestored} nội dung tích hợp cho Khối 5!`, "success");
    }

    function saveSettingsFromForm() {
        state.settings.governingBody = document.getElementById("set-governing-body").value.trim();
        state.settings.schoolName = document.getElementById("set-school-name").value.trim();
        state.settings.grade = document.getElementById("set-grade").value.trim();
        state.settings.className = document.getElementById("set-class-name").value.trim();
        state.settings.academicYear = document.getElementById("set-academic-year").value.trim();
        state.settings.homeroomTeacher = document.getElementById("set-teacher").value.trim();
        state.settings.headOfGrade = document.getElementById("set-head").value.trim();
        state.settings.principal = document.getElementById("set-principal").value.trim();
        state.settings.location = document.getElementById("set-location").value.trim();
        state.settings.dateString = document.getElementById("set-date").value.trim();

        // Collect PHT names
        const phtInputs = document.querySelectorAll(".pht-name-input");
        const newVps = [];
        phtInputs.forEach(inp => {
            const val = inp.value.trim();
            if (val) newVps.push(val);
        });
        if (newVps.length > 0) {
            state.settings.vicePrincipals = newVps;
            state.settings.vicePrincipal = newVps[0];
        }

        const selectLbg = document.getElementById("set-bgh-signer-lbg-select");
        if (selectLbg && selectLbg.value) {
            const [type, idxStr] = selectLbg.value.split(':');
            state.settings.bghSignerLbgType = type;
            state.settings.bghSignerLbgIndex = parseInt(idxStr) || 0;
            state.settings.bghSignerType = type;
            state.settings.bghSignerIndex = parseInt(idxStr) || 0;
        }

        const selectKhdh = document.getElementById("set-bgh-signer-khdh-select");
        if (selectKhdh && selectKhdh.value) {
            const [type, idxStr] = selectKhdh.value.split(':');
            state.settings.bghSignerKhdhType = type;
            state.settings.bghSignerKhdhIndex = parseInt(idxStr) || 0;
        }
        
        updateTopHeaderBadge();
        saveState();
        showToast("Đã lưu thông tin cài đặt thành công!", "success");
        renderTabLbg();
        renderTabCtlop();
    }

    // 6. Modal Preview System & Orientation Management
    function updateModalOrientationUI(orient) {
        const modal = document.getElementById("preview-modal");
        if (!modal) return;
        const btnPort = document.getElementById("btn-modal-orient-portrait");
        const btnLand = document.getElementById("btn-modal-orient-landscape");
        const paperPages = modal.querySelectorAll(".paper-page");

        if (orient === "landscape") {
            modal.classList.add("is-landscape");
            paperPages.forEach(p => p.classList.add("landscape"));
            if (btnLand) btnLand.classList.add("active");
            if (btnPort) btnPort.classList.remove("active");
        } else {
            modal.classList.remove("is-landscape");
            paperPages.forEach(p => p.classList.remove("landscape"));
            if (btnPort) btnPort.classList.add("active");
            if (btnLand) btnLand.classList.remove("active");
        }
    }

    function printWithOrientation(orient = null) {
        const activeOrient = orient || currentOrientation || "portrait";
        if (activeOrient === "landscape") {
            document.body.classList.add("print-landscape");
        } else {
            document.body.classList.remove("print-landscape");
        }
        const cleanUp = () => {
            document.body.classList.remove("print-landscape");
            window.removeEventListener("afterprint", cleanUp);
        };
        window.addEventListener("afterprint", cleanUp);
        window.print();
    }

    function setAppOrientation(newOrient) {
        currentOrientation = newOrient;

        // Sync Tab 1 buttons
        const lbgPort = document.getElementById("btn-orient-portrait-lbg") || document.getElementById("btn-orient-portrait");
        const lbgLand = document.getElementById("btn-orient-landscape-lbg") || document.getElementById("btn-orient-landscape");
        if (lbgPort && lbgLand) {
            if (newOrient === "landscape") {
                lbgLand.classList.add("active");
                lbgPort.classList.remove("active");
            } else {
                lbgPort.classList.add("active");
                lbgLand.classList.remove("active");
            }
        }

        // Sync Tab 2 buttons
        const ctlopPort = document.getElementById("btn-ctlop-orient-portrait");
        const ctlopLand = document.getElementById("btn-ctlop-orient-landscape");
        if (ctlopPort && ctlopLand) {
            if (newOrient === "landscape") {
                ctlopLand.classList.add("active");
                ctlopPort.classList.remove("active");
            } else {
                ctlopPort.classList.add("active");
                ctlopLand.classList.remove("active");
            }
        }

        // Sync modal if open
        updateModalOrientationUI(newOrient);
        if (currentPreviewRefreshFn) {
            const modalEl = document.getElementById("preview-modal");
            if (modalEl && modalEl.classList.contains("show")) {
                document.getElementById("modal-preview-body").innerHTML = currentPreviewRefreshFn();
            }
        }
    }

    let currentPreviewRefreshFn = null;

    function openPreviewModal(title, contentHtml, onDownloadDocx, onDownloadXlsx, onPrint, refreshFn = null) {
        currentPreviewRefreshFn = refreshFn;
        const modal = document.getElementById("preview-modal");
        document.getElementById("modal-preview-title").innerText = title;
        document.getElementById("modal-preview-body").innerHTML = contentHtml;

        updateModalOrientationUI(currentOrientation);

        // Handle modal toolbar options bar
        const optBar = document.getElementById("modal-preview-options-bar");
        if (optBar) {
            optBar.style.display = refreshFn ? "flex" : "none";
            const cbSign = document.getElementById("modal-opt-col-sign");
            const cbNote = document.getElementById("modal-opt-col-note");
            const cbBgh = document.getElementById("modal-opt-sig-bgh");
            const cbGvcn = document.getElementById("modal-opt-sig-gvcn");
            const cbHead = document.getElementById("modal-opt-sig-head");
            if (cbSign) cbSign.checked = !!state.lbgShowColSign;
            if (cbNote) cbNote.checked = !!state.lbgShowColNote;
            if (cbBgh) cbBgh.checked = (state.lbgShowBghSign !== false);
            if (cbGvcn) cbGvcn.checked = (state.lbgShowGvcnSign !== false);
            if (cbHead) cbHead.checked = (state.lbgShowHeadSign !== false);
            const cbHide = document.getElementById("modal-opt-hide-empty-rows");
            if (cbHide) cbHide.checked = !!state.lbgHideEmptyRows;
            renderCustomColsManager("modal-custom-cols-list", true);
        }

        const btnDocx = document.getElementById("modal-btn-docx");
        const btnXlsx = document.getElementById("modal-btn-xlsx");
        const btnPrint = document.getElementById("modal-btn-print");

        btnDocx.style.display = onDownloadDocx ? "inline-flex" : "none";
        btnXlsx.style.display = onDownloadXlsx ? "inline-flex" : "none";
        btnPrint.style.display = onPrint ? "inline-flex" : "none";

        btnDocx.onclick = onDownloadDocx || null;
        btnXlsx.onclick = onDownloadXlsx || null;
        btnPrint.onclick = onPrint || (() => printWithOrientation(currentOrientation));

        modal.classList.add("show");
    }

    function closePreviewModal() {
        document.getElementById("preview-modal").classList.remove("show");
    }

    // Preview KHDH (Single or Multi-subject)
    function previewKhdh(selectedSubjects) {
        if (!selectedSubjects || selectedSubjects.length === 0) {
            alert("Vui lòng chọn ít nhất 1 môn học để xem trước hoặc xuất Kế hoạch dạy học!");
            return;
        }

        const isMulti = selectedSubjects.length > 1;
        const subjectsData = selectedSubjects.map(subName => {
            const rows = state.khdh[subName] || [];
            return { subjectName: subName, rows: rows };
        });

        const mainTitle = isMulti 
            ? `KẾ HOẠCH DẠY HỌC CÁC MÔN HỌC ${(state.settings.grade || 'KHỐI 5').toUpperCase()}`
            : `KẾ HOẠCH DẠY HỌC MÔN ${selectedSubjects[0].toUpperCase()} ${(state.settings.grade || 'KHỐI 5').toUpperCase()}`;

        let paperHtml = `
        <div class="paper-page">
            <div class="paper-header">
                <table class="paper-header-table">
                    <tr>
                        <td style="width:48%; text-align:center;">
                            <div style="font-size:12pt; text-transform:uppercase;">${state.settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT'}</div>
                            <div style="font-size:12pt; font-weight:bold; text-transform:uppercase; text-decoration:underline;">${state.settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT'}</div>
                            <div style="font-size:12pt; margin-top:5px;">Số: &nbsp; &nbsp; &nbsp; /KH-TH-K5</div>
                        </td>
                        <td style="width:52%; text-align:center;">
                            <div style="font-size:12pt; font-weight:bold;">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
                            <div style="font-size:13pt; font-weight:bold; text-decoration:underline;">Độc lập - Tự do - Hạnh phúc</div>
                            <div style="font-size:12pt; font-style:italic; margin-top:5px; text-align:right;">${state.settings.location || 'Trung Nhứt'}, ${state.settings.dateString || 'ngày 28 tháng 8 năm 2026'}</div>
                        </td>
                    </tr>
                </table>
            </div>

            <div class="paper-title">${mainTitle}</div>
            <div class="paper-subtitle">NĂM HỌC ${state.settings.academicYear || '2026 - 2027'}</div>
        `;

        const basesList = (Array.isArray(state.settings.bases) && state.settings.bases.length > 0)
            ? state.settings.bases
            : DEFAULT_BASES;

        basesList.forEach(b => {
            paperHtml += `<div class="paper-basis">${escapeHtml(b)}</div>`;
        });

        paperHtml += `
            <div style="font-size:12pt; text-indent:25px; margin: 10px 0;">Tổ chuyên môn ${state.settings.grade || 'Khối 5'} xây dựng Kế hoạch dạy học chi tiết ${isMulti ? 'các môn học' : 'môn ' + escapeHtml(selectedSubjects[0])} như sau:</div>
        `;

        // Helper to determine subject category
        function getSubjectCategory(name) {
            const s = (name || '').toUpperCase().trim();
            if (s.includes('TIẾNG VIỆT') || s.includes('TIENG VIET')) return 'tieng_viet';
            if (s.includes('TRẢI NGHIỆM') || s.includes('TRAI NGHIEM') || s.includes('HĐTN') || s.includes('HDTN')) return 'hdtn';
            return 'khmh'; // TOÁN, KHOA HỌC, LS&ĐL, CÔNG NGHỆ, ĐẠO ĐỨC
        }

        subjectsData.forEach((sub, subIdx) => {
            const partHeader = isMulti 
                ? `PHẦN ${subIdx + 1}: KẾ HOẠCH DẠY HỌC MÔN ${sub.subjectName.toUpperCase()}`
                : `BẢNG KẾ HOẠCH DẠY HỌC CHI TIẾT (35 TUẦN)`;

            const cat = getSubjectCategory(sub.subjectName);

            paperHtml += `
            <div style="font-size:13pt; font-weight:bold; color:#002060; margin: 25px 0 10px 0;">${partHeader}</div>
            `;

            if (cat === 'tieng_viet') {
                paperHtml += `
                <table class="paper-table">
                    <thead>
                        <tr>
                            <th style="width:6%;">Tuần</th>
                            <th style="width:15%;">Chủ đề / Mạch nội dung</th>
                            <th style="width:17%;">Bài học</th>
                            <th style="width:22%;">Tên hoạt động / Nội dung dạy học</th>
                            <th style="width:8%;">Tiết / Thời lượng</th>
                            <th style="width:24%;">Nội dung điều chỉnh, bổ sung (nếu có)</th>
                            <th style="width:8%;">Ghi chú</th>
                        </tr>
                    </thead>
                    <tbody>
                `;

                const spans = computeKhdhSpans(sub.rows, 'tieng_viet');

                sub.rows.forEach((r, rIdx) => {
                    let rawInteg = r.integration || '';
                    rawInteg = rawInteg.replace(/System\.Xml\.XmlElement/g, '').trim();

                    let integFormatted = escapeHtml(rawInteg)
                        .replace(/(Lý tưởng cách mạng[^\:]*\:|Quyền con người[^\:]*\:|Tích hợp QCN[^\:]*\:|NLS[^\:]*\:|Năng lực số[^\:]*\:|Tích hợp NLS[^\:]*\:|AI[^\:]*\:|Tích hợp AI[^\:]*\:|BVMT[^\:]*\:|ANQP[^\:]*\:|QPAN[^\:]*\:|Tiết kiệm và bảo vệ nguồn nước[^\:]*\:|Đạo đức, lối sống[^\:]*\:)/g, '<br><strong>$1</strong>')
                        .replace(/^<br>/, '')
                        .replace(/\n/g, '<br>');

                    const contentFormatted = escapeHtml(r.content || '').replace(/\n/g, '<br>');
                    const lessonFormatted = escapeHtml(r.lesson || '').replace(/\n/g, '<br>');
                    const topicFormatted = escapeHtml(r.topic || r.theme || '').replace(/\n/g, '<br>');

                    // Gộp ô tuần
                    let weekTd = '';
                    if (spans.week[rIdx] > 0) {
                        weekTd = `<td rowspan="${spans.week[rIdx]}" style="text-align:center; font-weight:bold; vertical-align:middle; background:#fff;">${escapeHtml(r.week || '')}</td>`;
                    }

                    // Gộp ô chủ đề
                    let topicTd = '';
                    if (spans.topic[rIdx] > 0) {
                        topicTd = `<td rowspan="${spans.topic[rIdx]}" style="font-weight:600; text-align:left; vertical-align:middle; background:#fff;">${topicFormatted}</td>`;
                    }

                    // Gộp ô bài học
                    let lessonTd = '';
                    if (spans.lesson[rIdx] > 0) {
                        lessonTd = `<td rowspan="${spans.lesson[rIdx]}" style="font-weight:600; text-align:left; vertical-align:middle; background:#fff;">${lessonFormatted}</td>`;
                    }

                    // Gộp ô thời lượng
                    let durTd = '';
                    if (spans.duration[rIdx] > 0) {
                        const durText = r.duration || (spans.duration[rIdx] + ' tiết');
                        durTd = `<td rowspan="${spans.duration[rIdx]}" style="text-align:center; vertical-align:middle; background:#fff;">${escapeHtml(durText)}</td>`;
                    }

                    paperHtml += `
                        <tr>
                            ${weekTd}
                            ${topicTd}
                            ${lessonTd}
                            <td class="khdh-content-cell" style="text-align:justify; vertical-align:middle; line-height:1.5;">${contentFormatted}</td>
                            ${durTd}
                            <td style="font-size:9.5pt; line-height:1.45; text-align:left; vertical-align:top;">${integFormatted}</td>
                            <td style="font-size:9.5pt; text-align:left; vertical-align:middle;">${escapeHtml(r.notes || '')}</td>
                        </tr>
                    `;
                });

                paperHtml += `</tbody></table>`;

            } else if (cat === 'hdtn') {
                paperHtml += `
                <table class="paper-table">
                    <thead>
                        <tr>
                            <th style="width:6%;">Tuần</th>
                            <th style="width:16%;">Chủ đề / Mạch nội dung</th>
                            <th style="width:18%;">Tên bài học</th>
                            <th style="width:20%;">Tên hoạt động</th>
                            <th style="width:7%;">Thời lượng</th>
                            <th style="width:33%;">Nội dung điều chỉnh, bổ sung (nếu có)</th>
                        </tr>
                    </thead>
                    <tbody>
                `;

                const spans = computeKhdhSpans(sub.rows, 'hdtn');

                sub.rows.forEach((r, rIdx) => {
                    let rawInteg = r.integration || '';
                    rawInteg = rawInteg.replace(/System\.Xml\.XmlElement/g, '').trim();

                    let integFormatted = escapeHtml(rawInteg)
                        .replace(/(Lý tưởng cách mạng[^\:]*\:|Quyền con người[^\:]*\:|Tích hợp QCN[^\:]*\:|NLS[^\:]*\:|Năng lực số[^\:]*\:|Tích hợp NLS[^\:]*\:|AI[^\:]*\:|Tích hợp AI[^\:]*\:|BVMT[^\:]*\:|ANQP[^\:]*\:|QPAN[^\:]*\:|Tiết kiệm và bảo vệ nguồn nước[^\:]*\:|Đạo đức, lối sống[^\:]*\:)/g, '<br><strong>$1</strong>')
                        .replace(/^<br>/, '')
                        .replace(/\n/g, '<br>');

                    const contentFormatted = escapeHtml(r.content || '').replace(/\n/g, '<br>');
                    const lessonFormatted = escapeHtml(r.lesson || '').replace(/\n/g, '<br>');
                    const topicFormatted = escapeHtml(r.topic || r.theme || '').replace(/\n/g, '<br>');

                    let weekTd = '';
                    if (spans.week[rIdx] > 0) {
                        weekTd = `<td rowspan="${spans.week[rIdx]}" style="text-align:center; font-weight:bold; vertical-align:middle; background:#fff;">${escapeHtml(r.week || '')}</td>`;
                    }

                    let topicTd = '';
                    if (spans.topic[rIdx] > 0) {
                        topicTd = `<td rowspan="${spans.topic[rIdx]}" style="font-weight:600; text-align:left; vertical-align:middle; background:#fff;">${topicFormatted}</td>`;
                    }

                    let lessonTd = '';
                    if (spans.lesson[rIdx] > 0) {
                        lessonTd = `<td rowspan="${spans.lesson[rIdx]}" style="font-weight:600; text-align:left; vertical-align:middle; background:#fff;">${lessonFormatted}</td>`;
                    }

                    let durTd = '';
                    if (spans.duration[rIdx] > 0) {
                        const durText = r.duration || (spans.duration[rIdx] + ' tiết');
                        durTd = `<td rowspan="${spans.duration[rIdx]}" style="text-align:center; vertical-align:middle; background:#fff;">${escapeHtml(durText)}</td>`;
                    }

                    paperHtml += `
                        <tr>
                            ${weekTd}
                            ${topicTd}
                            ${lessonTd}
                            <td class="khdh-content-cell" style="text-align:justify; vertical-align:middle; line-height:1.5;">${contentFormatted}</td>
                            ${durTd}
                            <td style="font-size:9.5pt; line-height:1.45; text-align:left; vertical-align:top;">${integFormatted}</td>
                        </tr>
                    `;
                });

                paperHtml += `</tbody></table>`;

            } else {
                paperHtml += `
                <table class="paper-table">
                    <thead>
                        <tr>
                            <th style="width:6%;">Tuần</th>
                            <th style="width:18%;">Chủ đề / Mạch nội dung</th>
                            <th style="width:23%;">Tên bài học</th>
                            <th style="width:8%;">Thời lượng</th>
                            <th style="width:9%;">Tiết theo KHMH</th>
                            <th style="width:36%;">Nội dung điều chỉnh, bổ sung (nếu có)</th>
                        </tr>
                    </thead>
                    <tbody>
                `;

                const spans = computeKhdhSpans(sub.rows, 'khmh');

                sub.rows.forEach((r, rIdx) => {
                    let rawInteg = r.integration || '';
                    rawInteg = rawInteg.replace(/System\.Xml\.XmlElement/g, '').trim();

                    let integFormatted = escapeHtml(rawInteg)
                        .replace(/(Lý tưởng cách mạng[^\:]*\:|Quyền con người[^\:]*\:|Tích hợp QCN[^\:]*\:|NLS[^\:]*\:|Năng lực số[^\:]*\:|Tích hợp NLS[^\:]*\:|AI[^\:]*\:|Tích hợp AI[^\:]*\:|BVMT[^\:]*\:|ANQP[^\:]*\:|QPAN[^\:]*\:|Tiết kiệm và bảo vệ nguồn nước[^\:]*\:|Đạo đức, lối sống[^\:]*\:)/g, '<br><strong>$1</strong>')
                        .replace(/^<br>/, '')
                        .replace(/\n/g, '<br>');

                    const topicFormatted = escapeHtml(r.topic || r.theme || '').replace(/\n/g, '<br>');
                    const khmhPeriodText = r.khmhPeriod || r.content || (rIdx + 1).toString();

                    let weekTd = '';
                    if (spans.week[rIdx] > 0) {
                        weekTd = `<td rowspan="${spans.week[rIdx]}" style="text-align:center; font-weight:bold; vertical-align:middle; background:#fff;">${escapeHtml(r.week || '')}</td>`;
                    }

                    let topicTd = '';
                    if (spans.topic[rIdx] > 0) {
                        topicTd = `<td rowspan="${spans.topic[rIdx]}" style="font-weight:600; text-align:left; vertical-align:middle; background:#fff;">${topicFormatted}</td>`;
                    }

                    let lessonTd = '';
                    if (spans.lesson[rIdx] > 0) {
                        const displayLesson = (spans.lesson[rIdx] > 1) ? (getBaseLesson(r.lesson) || r.lesson) : (r.lesson || '');
                        lessonTd = `<td rowspan="${spans.lesson[rIdx]}" style="text-align:left; vertical-align:middle; background:#fff;">${escapeHtml(displayLesson).replace(/\n/g, '<br>')}</td>`;
                    }

                    let durTd = '';
                    if (spans.duration[rIdx] > 0) {
                        const durText = r.duration || (spans.duration[rIdx] + ' tiết');
                        durTd = `<td rowspan="${spans.duration[rIdx]}" style="text-align:center; vertical-align:middle; background:#fff;">${escapeHtml(durText)}</td>`;
                    }

                    paperHtml += `
                        <tr>
                            ${weekTd}
                            ${topicTd}
                            ${lessonTd}
                            ${durTd}
                            <td style="text-align:center; font-weight:bold; vertical-align:middle;">${escapeHtml(khmhPeriodText)}</td>
                            <td style="font-size:9.5pt; line-height:1.45; text-align:left; vertical-align:top;">${integFormatted}</td>
                        </tr>
                    `;
                });

                paperHtml += `</tbody></table>`;
            }
        });

        const bghSigner = getBghSignerInfo('KHDH');
        paperHtml += `
            <div class="paper-footer">
                <table class="paper-footer-table">
                    <tr>
                        <td style="width:40%; font-size:11pt;">
                            <div style="font-weight:bold; font-style:italic;">Nơi nhận:</div>
                            <div>- Ban Giám hiệu (để b/c);</div>
                            <div>- Tổ CM ${state.settings.grade || 'Khối 5'} (để t/h);</div>
                            <div>- Giáo viên giảng dạy;</div>
                            <div>- Lưu: VT, Hồ sơ Tổ.</div>
                        </td>
                        <td style="width:60%; text-align:center;">
                            <div style="font-size:12pt; font-weight:bold;">TỔ TRƯỞNG CHUYÊN MÔN</div>
                            <div style="font-size:11pt; font-style:italic;">(Ký và ghi rõ họ tên)</div>
                            <div style="height:65px;"></div>
                            <div style="font-size:12pt; font-weight:bold;">${state.settings.headOfGrade || 'Trần Thị Mai'}</div>

                            <div style="font-size:12pt; font-weight:bold; margin-top:20px;">PHÊ DUYỆT CỦA ${bghSigner.title}</div>
                            <div style="font-size:11pt; font-style:italic;">(Ký, ghi rõ họ tên và đóng dấu)</div>
                            <div style="height:70px;"></div>
                            <div style="font-size:12pt; font-weight:bold;">${bghSigner.name}</div>
                        </td>
                    </tr>
                </table>
            </div>
        </div>
        `;

        openPreviewModal(
            `Xem trước Kế hoạch dạy học (${isMulti ? selectedSubjects.length + ' môn gộp' : selectedSubjects[0]})`,
            paperHtml,
            () => {
                if (window.DocxGenerator && typeof window.DocxGenerator.generateKhdhDocx === 'function') {
                    window.DocxGenerator.generateKhdhDocx({
                        settings: state.settings,
                        subjectsData: subjectsData,
                        isMultiSubject: isMulti
                    }).then(blob => {
                        const fileName = isMulti 
                            ? `KHDH_Khoi_5_${selectedSubjects.length}_Mon.docx`
                            : `KHDH_Mon_${selectedSubjects[0]}_Lop_5.docx`;
                        if (typeof saveAs === 'function') {
                            saveAs(blob, fileName);
                        } else {
                            const url = URL.createObjectURL(blob);
                            const a = document.createElement('a');
                            a.href = url;
                            a.download = fileName;
                            document.body.appendChild(a);
                            a.click();
                            document.body.removeChild(a);
                            URL.revokeObjectURL(url);
                        }
                        showToast(`Đã xuất file Word: ${fileName}`, "success");
                    }).catch(err => {
                        console.error("DOCX Export error:", err);
                        alert("Lỗi xuất file Word: " + err.message);
                    });
                } else {
                    alert("Bộ sinh file Word chưa sẵn sàng, vui lòng tải lại trang!");
                }
            },
            null,
            () => {
                window.print();
            }
        );
    }

    // Export LBG to Excel
    function exportLbgToExcel(isCtlop = false) {
        const { weekInfo, schedule, stats } = calculateWeekSchedule(state.currentWeek);
        const filename = isCtlop 
            ? `Lich_Bao_Giang_Tich_Hop_Tuan_${state.currentWeek}_Lop_5A.xlsx`
            : `Lich_Bao_Giang_Tuan_${state.currentWeek}_Lop_5A.xlsx`;

        if (window.XlsxGenerator) {
            window.XlsxGenerator.generateLbgXlsx({
                settings: state.settings,
                weekNum: state.currentWeek,
                weekInfo: weekInfo,
                schedule: schedule,
                isCtlop: isCtlop,
                orientation: currentOrientation,
                showColSign: state.lbgShowColSign,
                showColNote: state.lbgShowColNote,
                showColCustom: (Array.isArray(state.lbgCustomCols) && state.lbgCustomCols.some(c => c.enabled !== false)) || state.lbgShowColCustom,
                colCustomName: (state.lbgCustomCols && state.lbgCustomCols[0] && state.lbgCustomCols[0].name) || state.lbgColCustomName || "Ghi chú",
                colCustomPos: (state.lbgCustomCols && state.lbgCustomCols[0] && state.lbgCustomCols[0].pos) || state.lbgColCustomPos || "end",
                customCols: state.lbgCustomCols || [],
                showBghSign: state.lbgShowBghSign,
                showGvcnSign: state.lbgShowGvcnSign,
                showHeadSign: state.lbgShowHeadSign,
                hideEmptyRows: !!state.lbgHideEmptyRows
            }).then(blob => {
                saveAs(blob, filename);
                showToast(`Đã xuất file Excel chuẩn in A4: ${filename}`, "success");
            }).catch(err => {
                console.error("XLSX Generator error:", err);
                alert("Lỗi xuất file Excel: " + err.message);
            });
        } else {
            alert("Bộ tạo file Excel chưa sẵn sàng!");
        }
    }

    // Export LBG to Word (.docx)
    function exportLbgToDocx(isCtlop = false, customOrientation = null) {
        const orient = customOrientation || currentOrientation || "portrait";
        if (window.DocxGenerator && window.DocxGenerator.generateLbgDocx) {
            showToast(`Đang tạo file Word (${orient === 'landscape' ? 'Khổ ngang' : 'Khổ đứng'}) chuẩn Nghị định 30...`, "info");
            const { weekInfo, schedule, stats } = calculateWeekSchedule(state.currentWeek);
            const options = {
                showColSign: state.lbgShowColSign,
                showColNote: state.lbgShowColNote,
                showColCustom: (Array.isArray(state.lbgCustomCols) && state.lbgCustomCols.some(c => c.enabled !== false)) || state.lbgShowColCustom,
                colCustomName: (state.lbgCustomCols && state.lbgCustomCols[0] && state.lbgCustomCols[0].name) || state.lbgColCustomName || "Ghi chú",
                colCustomPos: (state.lbgCustomCols && state.lbgCustomCols[0] && state.lbgCustomCols[0].pos) || state.lbgColCustomPos || "end",
                customCols: state.lbgCustomCols || [],
                showBghSign: state.lbgShowBghSign,
                showGvcnSign: state.lbgShowGvcnSign,
                showHeadSign: state.lbgShowHeadSign,
                hideEmptyRows: !!state.lbgHideEmptyRows
            };
            window.DocxGenerator.generateLbgDocx(isCtlop, state.currentWeek, weekInfo, schedule, state.settings, stats, orient, options).then(blob => {
                const prefix = isCtlop ? "Lich_Bao_Giang_Tich_Hop" : "Lich_Bao_Giang";
                const orientSuffix = orient === "landscape" ? "_Kho_Ngang" : "";
                const filename = `${prefix}_Tuan_${state.currentWeek}_${(state.settings.className || 'Lop_5A').replace(/\s+/g, '_')}${orientSuffix}.docx`;
                saveAs(blob, filename);
                showToast(`Đã xuất file Word thành công: ${filename}`, "success");
            }).catch(err => {
                console.error("DOCX Generator error:", err);
                alert("Lỗi xuất file Word: " + err.message);
            });
        } else {
            alert("Bộ tạo file Word chưa sẵn sàng!");
        }
    }

    // Helper to render HTML for a single week's paper page
    function renderSingleWeekPaperHtml(weekNum, isCtlop, customOrientation = null) {
        const { weekInfo, schedule, stats } = calculateWeekSchedule(weekNum);
        const bghSigner = getBghSignerInfo();
        const activeOrient = customOrientation || currentOrientation || "portrait";
        const isLand = (activeOrient === "landscape");
        const paperClass = isLand ? "paper-page landscape" : "paper-page";
        const maxWidth = isLand ? "1100px" : (isCtlop ? "960px" : "900px");
        const showSign = !!state.lbgShowColSign;
        const showNote = !!state.lbgShowColNote;

        const orderedCols = getLbgOrderedColumns(state.lbgCustomCols, showSign, showNote, isCtlop);

        // Compute adaptive column widths
        const customCount = Array.isArray(state.lbgCustomCols) ? state.lbgCustomCols.filter(c => c && c.enabled !== false).length : 0;
        let customColWidth = isLand ? '12%' : '10%';
        if (customCount === 2) customColWidth = isLand ? '10%' : '8%';
        else if (customCount >= 3) customColWidth = isLand ? '8%' : '7%';

        let theadColsHtml = "";
        orderedCols.forEach(col => {
            if (col.key === 'day') {
                theadColsHtml += `<th style="width:${isLand ? '10%' : '11%'};">Thứ, ngày</th>`;
            } else if (col.key === 'session') {
                theadColsHtml += `<th style="width:${isLand ? '6%' : '7%'};">Buổi</th>`;
            } else if (col.key === 'period') {
                theadColsHtml += `<th style="width:${isLand ? '4%' : '5%'};">Tiết</th>`;
            } else if (col.key === 'subject') {
                theadColsHtml += `<th style="width:${isLand ? '17%' : '18%'};">Môn học</th>`;
            } else if (col.key === 'ppct') {
                theadColsHtml += `<th style="width:${isLand ? '7%' : '8%'};">Tiết PPCT</th>`;
            } else if (col.key === 'lesson') {
                theadColsHtml += `<th>Tên bài dạy</th>`;
            } else if (col.key === 'integ') {
                theadColsHtml += `<th style="width:${isLand ? '30%' : '23%'};">Nội dung tích hợp / Điều chỉnh</th>`;
            } else if (col.key === 'sign') {
                theadColsHtml += `<th style="width:${isLand ? '7%' : '8%'};">Kí tên</th>`;
            } else if (col.key === 'note') {
                theadColsHtml += `<th style="width:${isLand ? '10%' : '11%'};">Ghi chú</th>`;
            } else if (col.isCustom) {
                theadColsHtml += `<th style="width:${customColWidth};">${escapeHtml(col.title)}</th>`;
            }
        });

        let html = `
        <div class="${paperClass}" style="max-width:${maxWidth}; margin-bottom: 2.5rem; page-break-after: always;">
            <div class="paper-header">
                <table class="paper-header-table">
                    <tr>
                        <td style="width:${isLand ? '48%' : '50%'}; text-align:center;">
                            <div style="font-size:13pt; text-transform:uppercase;">${state.settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT'}</div>
                            <div style="font-size:13pt; font-weight:bold; text-transform:uppercase;">${state.settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT'}</div>
                            <div style="font-size:13pt; font-weight:bold; margin-top:2px;">${state.settings.grade || 'KHỐI 5'} - ${state.settings.className || 'LỚP 5A'}</div>
                        </td>
                        <td style="width:${isLand ? '52%' : '50%'}; text-align:center;">
                            <div style="font-size:13pt; font-weight:bold;">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
                            <div style="font-size:13pt; font-weight:bold; text-decoration:underline;">Độc lập - Tự do - Hạnh phúc</div>
                        </td>
                    </tr>
                </table>
            </div>

            <div class="paper-title" style="font-size:13pt;">${isCtlop ? 'LỊCH BÁO GIẢNG TÍCH HỢP' : 'LỊCH BÁO GIẢNG'} TUẦN ${weekNum}</div>
            <div class="paper-subtitle" style="font-size:13pt;">(Thời gian thực hiện: Từ ngày ${weekInfo.startDateVN} đến ngày ${weekInfo.endDateVN})</div>

            <table class="paper-table">
                <thead>
                    <tr>
                        ${theadColsHtml}
                    </tr>
                </thead>
                <tbody>
        `;

        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];
        days.forEach(day => {
            const dayDate = getDayFullDate(weekInfo.startDateVN, day);
            let daySlots = schedule.filter(s => s.day === day);
            if (state.lbgHideEmptyRows) {
                daySlots = daySlots.filter(s => !s.isOff && s.subject !== '-- Nghỉ / Để trống --');
            }
            if (daySlots.length === 0) return;

            const morningSlots = daySlots.filter(s => s.session === "Sáng").sort((a, b) => (a.period || 0) - (b.period || 0));
            const afternoonSlots = daySlots.filter(s => s.session === "Chiều").sort((a, b) => (a.period || 0) - (b.period || 0));
            const orderedDaySlots = [...morningSlots, ...afternoonSlots];
            orderedDaySlots.forEach((slot, idx) => {
                let dayCell = idx === 0 
                    ? `<td rowspan="${orderedDaySlots.length}" style="text-align:center; font-weight:bold; vertical-align:middle; line-height:1.35;">
                        <div>${day}</div>
                        ${dayDate ? `<div style="font-size:8.5pt; font-weight:normal; color:#475569; margin-top:3px;">${dayDate}</div>` : ''}
                       </td>` 
                    : "";
                let sessionCell = "";
                if (morningSlots.length > 0 && idx === 0) sessionCell = `<td rowspan="${morningSlots.length}" style="text-align:center; vertical-align:middle;">Sáng</td>`;
                else if (afternoonSlots.length > 0 && idx === morningSlots.length) sessionCell = `<td rowspan="${afternoonSlots.length}" style="text-align:center; vertical-align:middle;">Chiều</td>`;

                let integrationCell = "";
                if (isCtlop) {
                    const formattedInteg = escapeHtml(slot.integration || '')
                        .replace(/(Lý tưởng cách mạng[^\:]*\:|Quyền con người[^\:]*\:|Tích hợp QCN[^\:]*\:|NLS[^\:]*\:|Năng lực số[^\:]*\:|Tích hợp NLS[^\:]*\:|AI[^\:]*\:|Tích hợp AI[^\:]*\:|BVMT[^\:]*\:|ANQP[^\:]*\:|QPAN[^\:]*\:|Tiết kiệm và bảo vệ nguồn nước[^\:]*\:|Đạo đức, lối sống[^\:]*\:)/g, '<br><strong>$1</strong>')
                        .replace(/^<br>/, '')
                        .replace(/\n/g, '<br>');
                    integrationCell = `<td style="font-size:9.5pt; line-height:1.45; text-align:left; vertical-align:top;">${formattedInteg}</td>`;
                }

                let rowCellsHtml = "";
                orderedCols.forEach(col => {
                    if (col.key === 'day') {
                        if (dayCell) rowCellsHtml += dayCell;
                    } else if (col.key === 'session') {
                        if (sessionCell) rowCellsHtml += sessionCell;
                    } else if (col.key === 'period') {
                        rowCellsHtml += `<td style="text-align:center; font-weight:bold; vertical-align:middle;">${slot.period}</td>`;
                    } else if (col.key === 'subject') {
                        rowCellsHtml += `<td style="font-weight:bold; text-align:center; vertical-align:middle;">${escapeHtml(slot.subject)}</td>`;
                    } else if (col.key === 'ppct') {
                        rowCellsHtml += `<td style="text-align:center; font-weight:bold; vertical-align:middle;">${escapeHtml(slot.ppct || '')}</td>`;
                    } else if (col.key === 'lesson') {
                        rowCellsHtml += `<td class="cell-lesson" style="text-align:left; vertical-align:middle; padding-top:1pt; padding-bottom:1pt; padding-left:1pt; text-indent:2pt; line-height:1.35;">${escapeHtml(slot.lessonName || '')}</td>`;
                    } else if (col.key === 'integ') {
                        rowCellsHtml += integrationCell;
                    } else if (col.key === 'sign') {
                        rowCellsHtml += `<td style="text-align:center; vertical-align:middle;"></td>`;
                    } else if (col.key === 'note') {
                        rowCellsHtml += `<td style="text-align:left; vertical-align:middle; font-size:8.5pt;">${escapeHtml(slot.note || '')}</td>`;
                    } else if (col.isCustom) {
                        const customVal = (slot.customCols && slot.customCols[col.id] !== undefined)
                            ? slot.customCols[col.id]
                            : (col.id === 'col_1' ? (slot.customCol || slot.note || '') : (slot['customCol_' + col.id] || ''));
                        rowCellsHtml += `<td style="text-align:left; vertical-align:middle; font-size:8.5pt;">${escapeHtml(customVal)}</td>`;
                    }
                });

                html += `
                    <tr>
                        ${rowCellsHtml}
                    </tr>
                `;
            });
        });

        html += `
                </tbody>
            </table>

            ${renderPaperFooterSignatures(bghSigner)}
        </div>
        `;
        return html;
    }

    // Helper: Generate dynamic signatures footer according to user preferences
    function renderPaperFooterSignatures(bghSigner) {
        const signers = [];
        if (state.lbgShowBghSign !== false) {
            signers.push({
                title: `DUYỆT CỦA ${bghSigner.title}`,
                sub: "(Ký và ghi rõ họ tên)",
                name: bghSigner.name
            });
        }
        if (state.lbgShowHeadSign !== false) {
            signers.push({
                title: "KHỐI TRƯỞNG",
                sub: "(Ký và ghi rõ họ tên)",
                name: state.settings.headOfGrade || 'Trần Thị Mai'
            });
        }
        if (state.lbgShowGvcnSign !== false) {
            signers.push({
                title: "GIÁO VIÊN CHỦ NHIỆM",
                sub: "(Ký và ghi rõ họ tên)",
                name: state.settings.homeroomTeacher || 'Nguyễn Thị Thu Hà'
            });
        }

        if (signers.length === 0) return '';

        const colWidthPercent = Math.floor(100 / signers.length);
        return `
            <div class="paper-footer">
                <table class="paper-footer-table">
                    <tr>
                        ${signers.map((s, idx) => {
                            const w = (idx === signers.length - 1) ? (100 - colWidthPercent * (signers.length - 1)) : colWidthPercent;
                            return `
                                <td style="width:${w}%; text-align:center;">
                                    <div style="font-size:12pt; font-weight:bold;">${escapeHtml(s.title)}</div>
                                    <div style="font-size:11pt; font-style:italic;">${escapeHtml(s.sub)}</div>
                                    <div style="height:65px;"></div>
                                    <div style="font-size:12pt; font-weight:bold;">${escapeHtml(s.name)}</div>
                                </td>
                            `;
                        }).join('')}
                    </tr>
                </table>
            </div>
        `;
    }

    // VERSION 6.0: Single week paper page for Subject-wise schedule (HTML for Print & Preview)
    function renderSingleWeekPaperBySubjectHtml(weekNum, isCtlop, customOrientation = null, filterSubject = "all", filterCategory = "all") {
        const { weekInfo, subjectGroups, stats } = calculateWeekScheduleBySubject(weekNum, filterSubject, filterCategory);
        const bghSigner = getBghSignerInfo();
        const activeOrient = customOrientation || state.lbgMonOrientation || currentOrientation || "portrait";
        const isLand = (activeOrient === "landscape");
        const maxWidth = isLand ? "1100px" : (isCtlop ? "960px" : "900px");
        const paperClass = isLand ? "paper-page landscape" : "paper-page";

        let titleStr = `LỊCH BÁO GIẢNG THEO MÔN HỌC TUẦN ${weekNum}`;
        if (filterSubject && filterSubject !== "all") {
            titleStr = `LỊCH BÁO GIẢNG MÔN ${filterSubject.toUpperCase()} TUẦN ${weekNum}`;
        }

        let html = `
        <div class="${paperClass}" style="max-width:${maxWidth}; margin-bottom: 2.5rem; page-break-after: always;">
            <div class="paper-header">
                <table class="paper-header-table">
                    <tr>
                        <td style="width:${isLand ? '48%' : '50%'}; text-align:center;">
                            <div style="font-size:13pt; text-transform:uppercase;">${state.settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT'}</div>
                            <div style="font-size:13pt; font-weight:bold; text-transform:uppercase;">${state.settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT'}</div>
                            <div style="font-size:13pt; font-weight:bold; margin-top:2px;">${state.settings.grade || 'KHỐI 5'} - ${state.settings.className || 'LỚP 5A'}</div>
                        </td>
                        <td style="width:${isLand ? '52%' : '50%'}; text-align:center;">
                            <div style="font-size:13pt; font-weight:bold;">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
                            <div style="font-size:13pt; font-weight:bold; text-decoration:underline;">Độc lập - Tự do - Hạnh phúc</div>
                        </td>
                    </tr>
                </table>
            </div>

            <div class="paper-title" style="font-size:13pt;">${titleStr}</div>
            <div class="paper-subtitle" style="font-size:13pt;">(Thời gian thực hiện: Từ ngày ${weekInfo.startDateVN} đến ngày ${weekInfo.endDateVN})</div>

            <table class="paper-table">
                <thead>
                    <tr>
                        <th style="width:${isCtlop ? (isLand ? '14%' : '16%') : (isLand ? '18%' : '20%')};">Môn học</th>
                        <th style="width:${isCtlop ? (isLand ? '10%' : '11%') : (isLand ? '12%' : '13%')};">Thứ / Ngày</th>
                        <th style="width:${isCtlop ? (isLand ? '6%' : '7%') : (isLand ? '7%' : '8%')};">Buổi</th>
                        <th style="width:${isCtlop ? (isLand ? '4%' : '5%') : (isLand ? '5%' : '6%')};">Tiết</th>
                        <th style="width:${isCtlop ? (isLand ? '6%' : '7%') : (isLand ? '7%' : '8%')};">Tiết/tuần</th>
                        <th style="width:${isCtlop ? (isLand ? '7%' : '8%') : (isLand ? '8%' : '9%')};">Tiết PPCT</th>
                        <th style="width:${isCtlop ? (isLand ? '25%' : '24%') : (isLand ? '43%' : '36%')};">Tên bài dạy</th>
                        ${isCtlop ? `<th style="width:${isLand ? '28%' : '22%'};">Nội dung tích hợp / Điều chỉnh</th>` : ''}
                    </tr>
                </thead>
                <tbody>
        `;

        if (subjectGroups.length === 0) {
            html += `
                <tr>
                    <td colspan="${isCtlop ? 8 : 7}" style="text-align:center; padding:20px; font-style:italic;">Không có tiết học nào phù hợp với bộ lọc đã chọn.</td>
                </tr>
            `;
        } else {
            subjectGroups.forEach(group => {
                group.slots.forEach((slot, idx) => {
                    const dayDateStr = getDayDateStr(weekInfo.startDateVN, slot.day);
                    let subCell = idx === 0 ? `<td rowspan="${group.slots.length}" style="text-align:center; font-weight:bold; vertical-align:middle; background:#fafafa;">${escapeHtml(group.subjectName)}<br><small style="font-weight:normal; color:#475569;">(${group.slots.length} tiết)</small></td>` : '';

                    let integrationCell = "";
                    if (isCtlop) {
                        const formattedInteg = escapeHtml(slot.integration || '')
                            .replace(/(Lý tưởng cách mạng[^\:]*\:|Quyền con người[^\:]*\:|Tích hợp QCN[^\:]*\:|NLS[^\:]*\:|Năng lực số[^\:]*\:|Tích hợp NLS[^\:]*\:|AI[^\:]*\:|Tích hợp AI[^\:]*\:|BVMT[^\:]*\:|ANQP[^\:]*\:|QPAN[^\:]*\:|Tiết kiệm và bảo vệ nguồn nước[^\:]*\:|Đạo đức, lối sống[^\:]*\:)/g, '<br><strong>$1</strong>')
                            .replace(/^<br>/, '')
                            .replace(/\n/g, '<br>');
                        integrationCell = `<td style="font-size:9pt; line-height:1.4; text-align:left; vertical-align:top;">${formattedInteg}</td>`;
                    }

                    html += `
                        <tr>
                            ${subCell}
                            <td style="text-align:center; font-weight:bold; vertical-align:middle;">${dayDateStr}</td>
                            <td style="text-align:center; vertical-align:middle;">${slot.session}</td>
                            <td style="text-align:center; font-weight:bold; vertical-align:middle;">${slot.period}</td>
                            <td style="text-align:center; font-weight:bold; vertical-align:middle;">${slot.periodInWeek}</td>
                            <td style="text-align:center; font-weight:bold; vertical-align:middle;">${escapeHtml(slot.ppct || '')}</td>
                            <td class="cell-lesson" style="text-align:left; vertical-align:middle; padding-top:1pt; padding-bottom:1pt; padding-left:1pt; text-indent:2pt; line-height:1.35;">${escapeHtml(slot.lessonName || '')}</td>
                            ${integrationCell}
                        </tr>
                    `;
                });
            });
        }

        html += `
                </tbody>
            </table>

            ${renderPaperFooterSignatures(bghSigner)}
        </div>
        `;
        return html;
    }

    // Batch Export / Print Modal Handlers
    function openBatchExportModal(defaultType = 'lbg') {
        const radios = document.querySelectorAll('input[name="batch-doc-type"]');
        radios.forEach(r => {
            r.checked = (r.value === defaultType);
        });

        // Set batch orientation from currentOrientation
        const batchOrientRadios = document.querySelectorAll('input[name="batch-orientation"]');
        batchOrientRadios.forEach(r => {
            r.checked = (r.value === currentOrientation);
        });

        const selStart = document.getElementById("batch-start-week");
        const selEnd = document.getElementById("batch-end-week");
        if (selStart && selEnd) {
            let optHtml = "";
            for (let i = 1; i <= 35; i++) {
                optHtml += `<option value="${i}">Tuần ${i}</option>`;
            }
            selStart.innerHTML = optHtml;
            selEnd.innerHTML = optHtml;

            selStart.value = "1";
            selEnd.value = "35";
        }

        const batchHideChk = document.getElementById("batch-opt-hide-empty-rows");
        if (batchHideChk) batchHideChk.checked = !!state.lbgHideEmptyRows;

        document.getElementById("batch-export-modal").classList.add("show");
    }

    function closeBatchExportModal() {
        document.getElementById("batch-export-modal").classList.remove("show");
    }

    function exportBatchLbgToDocxDirect(isCtlop, startWeek, endWeek, orientation = "portrait") {
        if (window.DocxGenerator && window.DocxGenerator.generateMultiWeekLbgDocx) {
            showToast(`Đang tạo file Word (${orientation === 'landscape' ? 'Khổ ngang' : 'Khổ đứng'}) từ Tuần ${startWeek} đến Tuần ${endWeek}...`, "info");
            const options = {
                showColSign: state.lbgShowColSign,
                showColNote: state.lbgShowColNote,
                showColCustom: (Array.isArray(state.lbgCustomCols) && state.lbgCustomCols.some(c => c.enabled !== false)) || state.lbgShowColCustom,
                colCustomName: (state.lbgCustomCols && state.lbgCustomCols[0] && state.lbgCustomCols[0].name) || state.lbgColCustomName || "Ghi chú",
                colCustomPos: (state.lbgCustomCols && state.lbgCustomCols[0] && state.lbgCustomCols[0].pos) || state.lbgColCustomPos || "end",
                customCols: state.lbgCustomCols || [],
                showBghSign: state.lbgShowBghSign,
                showGvcnSign: state.lbgShowGvcnSign,
                showHeadSign: state.lbgShowHeadSign,
                hideEmptyRows: !!state.lbgHideEmptyRows
            };
            window.DocxGenerator.generateMultiWeekLbgDocx(isCtlop, startWeek, endWeek, calculateWeekSchedule, state.settings, orientation, options).then(blob => {
                const prefix = isCtlop ? "Lich_Bao_Giang_Tich_Hop" : "Lich_Bao_Giang";
                const orientSuffix = orientation === "landscape" ? "_Kho_Ngang" : "";
                const filename = `${prefix}_Tuan_${startWeek}_den_${endWeek}_${(state.settings.className || 'Lop_5A').replace(/\s+/g, '_')}${orientSuffix}.docx`;
                saveAs(blob, filename);
                showToast(`Đã xuất file Word nhiều tuần thành công: ${filename}`, "success");
            }).catch(err => {
                console.error("Multi-week DOCX error:", err);
                alert("Lỗi xuất file Word: " + err.message);
            });
        } else {
            alert("Bộ tạo file Word chưa sẵn sàng!");
        }
    }

    function exportBatchLbgBySubjectToDocxDirect(startWeek, endWeek, orientation = "portrait") {
        if (window.DocxGenerator && window.DocxGenerator.generateBatchLbgBySubjectDocx) {
            showToast(`Đang tạo file Word (${orientation === 'landscape' ? 'Khổ ngang' : 'Khổ đứng'}) theo môn học từ Tuần ${startWeek} đến Tuần ${endWeek}...`, "info");
            const isCtlop = (state.lbgMonShowIntegration !== false);
            window.DocxGenerator.generateBatchLbgBySubjectDocx(startWeek, endWeek, calculateWeekScheduleBySubject, state.settings, orientation, isCtlop, state.lbgMonFilterSubject, state.lbgMonFilterCategory).then(blob => {
                const orientSuffix = orientation === "landscape" ? "_Kho_Ngang" : "";
                const filename = `Lich_Bao_Giang_Theo_Mon_Tuan_${startWeek}_den_${endWeek}_${(state.settings.className || 'Lop_5A').replace(/\s+/g, '_')}${orientSuffix}.docx`;
                saveAs(blob, filename);
                showToast(`Đã xuất file Word nhiều tuần theo môn học thành công: ${filename}`, "success");
            }).catch(err => {
                console.error("Batch LBG By Subject DOCX error:", err);
                alert("Lỗi xuất file Word: " + err.message);
            });
        } else {
            alert("Bộ tạo file Word chưa sẵn sàng!");
        }
    }

    function exportBatchLbgToDocx() {
        const docType = document.querySelector('input[name="batch-doc-type"]:checked')?.value || 'lbg';
        const isCtlop = docType === 'ctlop';
        const startWeek = parseInt(document.getElementById("batch-start-week").value) || 1;
        const endWeek = parseInt(document.getElementById("batch-end-week").value) || 35;
        const batchOrient = document.querySelector('input[name="batch-orientation"]:checked')?.value || currentOrientation || 'portrait';

        if (startWeek > endWeek) {
            alert("Tuần bắt đầu phải nhỏ hơn hoặc bằng tuần kết thúc!");
            return;
        }

        if (docType === 'lbg-mon') {
            exportBatchLbgBySubjectToDocxDirect(startWeek, endWeek, batchOrient);
        } else {
            exportBatchLbgToDocxDirect(isCtlop, startWeek, endWeek, batchOrient);
        }
        closeBatchExportModal();
    }

    function previewBatchLbg() {
        const isCtlop = document.querySelector('input[name="batch-doc-type"]:checked')?.value === 'ctlop';
        const startWeek = parseInt(document.getElementById("batch-start-week").value) || 1;
        const endWeek = parseInt(document.getElementById("batch-end-week").value) || 35;
        const batchOrient = document.querySelector('input[name="batch-orientation"]:checked')?.value || currentOrientation || 'portrait';

        if (startWeek > endWeek) {
            alert("Tuần bắt đầu phải nhỏ hơn hoặc bằng tuần kết thúc!");
            return;
        }

        setAppOrientation(batchOrient);

        let allHtml = "";
        const docType = document.querySelector('input[name="batch-doc-type"]:checked')?.value || 'lbg';
        for (let w = startWeek; w <= endWeek; w++) {
            if (docType === 'lbg-mon') {
                allHtml += renderSingleWeekPaperBySubjectHtml(w, true, batchOrient, 'all', 'all');
            } else {
                allHtml += renderSingleWeekPaperHtml(w, isCtlop, batchOrient);
            }
        }

        const docTitle = isCtlop ? "Lịch Báo Giảng Tích Hợp" : "Lịch Báo Giảng";
        closeBatchExportModal();
        const exportBatchFn = (docType === 'lbg-mon')
            ? () => exportBatchLbgBySubjectToDocxDirect(startWeek, endWeek, batchOrient)
            : () => exportBatchLbgToDocxDirect(isCtlop, startWeek, endWeek, batchOrient);
        openPreviewModal(
            `Xem trước ${docTitle} từ Tuần ${startWeek} đến Tuần ${endWeek} (${batchOrient === 'landscape' ? 'Khổ ngang' : 'Khổ đứng'})`,
            allHtml,
            exportBatchFn,
            null,
            () => printWithOrientation(batchOrient)
        );
    }

    // Global Tab Switching Helper Function
    function switchActiveTab(targetTab) {
        if (!targetTab) return;
        document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
        document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));

        const btn = document.querySelector(`.tab-btn[data-tab="${targetTab}"]`);
        if (btn) btn.classList.add("active");

        state.currentTab = targetTab;
        const targetEl = document.getElementById(targetTab);
        if (targetEl) {
            targetEl.classList.add("active");
            try {
                window.scrollTo({ top: 0, behavior: 'smooth' });
            } catch (e) {}
        }

        if (targetTab === "tab-lbg") renderTabLbg();
        else if (targetTab === "tab-ctlop") renderTabCtlop();
        else if (targetTab === "tab-lbg-mon") renderTabLbgMon();
        else if (targetTab === "tab-lichtuan") renderTabLichtuan();
        else if (targetTab === "tab-ppct") renderTabPpct();
        else if (targetTab === "tab-settings") renderTabSettings();
    }

    // 7. App Initialization & Event Bindings
    function initApp() {
        loadState();

        document.querySelectorAll(".tab-btn").forEach(btn => {
            btn.addEventListener("click", () => {
                const targetTab = btn.dataset.tab;
                switchActiveTab(targetTab);
            });
        });

        // Top Quick Access Buttons (Settings & Guide)
        const btnTopSettings = document.getElementById("btn-top-settings");
        if (btnTopSettings) {
            btnTopSettings.addEventListener("click", () => switchActiveTab("tab-settings"));
        }
        const btnTopGuide = document.getElementById("btn-top-guide");
        if (btnTopGuide) {
            btnTopGuide.addEventListener("click", () => switchActiveTab("tab-guide"));
        }

        // Tab 1 Actions & Orientation
        const btnOrientPort = document.getElementById("btn-orient-portrait-lbg") || document.getElementById("btn-orient-portrait");
        const btnOrientLand = document.getElementById("btn-orient-landscape-lbg") || document.getElementById("btn-orient-landscape");
        if (btnOrientPort) btnOrientPort.addEventListener("click", () => setAppOrientation("portrait"));
        if (btnOrientLand) btnOrientLand.addEventListener("click", () => setAppOrientation("landscape"));

        // Tab 1 & Tab 2 Options Checkboxes
        const bindLbgOption = (id, prop, isDefaultTrue = false) => {
            const el = document.getElementById(id);
            if (!el) return;
            el.addEventListener("change", (e) => {
                state[prop] = isDefaultTrue ? e.target.checked : !!e.target.checked;
                saveState();
                renderTabLbg();
                renderTabCtlop();
            });
        };
        bindLbgOption("lbg-opt-col-sign", "lbgShowColSign", false);
        bindLbgOption("lbg-opt-col-note", "lbgShowColNote", false);
        bindLbgOption("lbg-opt-sig-bgh", "lbgShowBghSign", true);
        bindLbgOption("lbg-opt-sig-gvcn", "lbgShowGvcnSign", true);
        bindLbgOption("lbg-opt-sig-head", "lbgShowHeadSign", true);
        bindLbgOption("lbg-opt-hide-empty-rows", "lbgHideEmptyRows", false);

        // Bind CTLOP Checkboxes
        bindLbgOption("ctlop-opt-col-sign", "lbgShowColSign", false);
        bindLbgOption("ctlop-opt-col-note", "lbgShowColNote", false);
        bindLbgOption("ctlop-opt-sig-bgh", "lbgShowBghSign", true);
        bindLbgOption("ctlop-opt-sig-gvcn", "lbgShowGvcnSign", true);
        bindLbgOption("ctlop-opt-sig-head", "lbgShowHeadSign", true);
        bindLbgOption("ctlop-opt-hide-empty-rows", "lbgHideEmptyRows", false);

        // Modal Preview Options Checkboxes
        const bindModalOption = (id, prop, isDefaultTrue = false) => {
            const el = document.getElementById(id);
            if (!el) return;
            el.addEventListener("change", (e) => {
                state[prop] = isDefaultTrue ? e.target.checked : !!e.target.checked;
                saveState();
                renderTabLbg();
                renderTabCtlop();
                if (currentPreviewRefreshFn) {
                    const body = document.getElementById("modal-preview-body");
                    if (body) body.innerHTML = currentPreviewRefreshFn();
                }
            });
        };
        bindModalOption("modal-opt-col-sign", "lbgShowColSign", false);
        bindModalOption("modal-opt-col-note", "lbgShowColNote", false);
        bindModalOption("modal-opt-sig-bgh", "lbgShowBghSign", true);
        bindModalOption("modal-opt-sig-gvcn", "lbgShowGvcnSign", true);
        bindModalOption("modal-opt-sig-head", "lbgShowHeadSign", true);
        bindModalOption("modal-opt-hide-empty-rows", "lbgHideEmptyRows", false);

        const batchHideChkEl = document.getElementById("batch-opt-hide-empty-rows");
        if (batchHideChkEl) {
            batchHideChkEl.addEventListener("change", (e) => {
                state.lbgHideEmptyRows = !!e.target.checked;
                saveState();
                renderTabLbg();
                renderTabCtlop();
                if (currentPreviewRefreshFn) {
                    const body = document.getElementById("modal-preview-body");
                    if (body) body.innerHTML = currentPreviewRefreshFn();
                }
            });
        }

        // Add Custom Column Handlers
        const handleAddNewCustomCol = () => {
            if (!Array.isArray(state.lbgCustomCols)) state.lbgCustomCols = [];
            const colNum = state.lbgCustomCols.length + 1;
            const newId = 'col_' + Date.now();
            state.lbgCustomCols.push({
                id: newId,
                name: `Cột mới ${colNum}`,
                pos: 'end',
                enabled: true
            });
            saveState();
            renderCustomColsManager("lbg-custom-cols-list", false);
            renderCustomColsManager("ctlop-custom-cols-list", false);
            renderCustomColsManager("modal-custom-cols-list", true);
            renderTabLbg();
            renderTabCtlop();
            if (currentPreviewRefreshFn) {
                const body = document.getElementById("modal-preview-body");
                if (body) body.innerHTML = currentPreviewRefreshFn();
            }
        };

        const btnAddCol1 = document.getElementById("btn-lbg-add-custom-col");
        if (btnAddCol1) btnAddCol1.onclick = handleAddNewCustomCol;
        const btnAddCol2 = document.getElementById("btn-modal-add-custom-col");
        if (btnAddCol2) btnAddCol2.onclick = handleAddNewCustomCol;
        const btnAddCol3 = document.getElementById("btn-ctlop-add-custom-col");
        if (btnAddCol3) btnAddCol3.onclick = handleAddNewCustomCol;

        document.getElementById("btn-lbg-reset-tkb").addEventListener("click", () => {
            if (confirm(`Bạn có muốn khôi phục Thời khóa biểu của Tuần ${state.currentWeek} về mặc định không?`)) {
                resetWeekSlotsToDefault(state.currentWeek);
            }
        });

        const btnLbgBatch = document.getElementById("btn-lbg-batch");
        if (btnLbgBatch) {
            btnLbgBatch.addEventListener("click", () => openBatchExportModal('lbg'));
        }

        document.getElementById("btn-lbg-preview").addEventListener("click", () => {
            const refreshFn = () => renderSingleWeekPaperHtml(state.currentWeek, false);
            openPreviewModal(
                `Xem trước Lịch Báo Giảng Tuần ${state.currentWeek} (${currentOrientation === 'landscape' ? 'Khổ ngang' : 'Khổ đứng'})`,
                refreshFn(),
                () => exportLbgToDocx(false),
                () => exportLbgToExcel(false),
                () => printWithOrientation(currentOrientation),
                refreshFn
            );
        });

        document.getElementById("btn-lbg-excel").addEventListener("click", () => exportLbgToExcel(false));
        const btnLbgPrint = document.getElementById("btn-lbg-print");
        if (btnLbgPrint) {
            btnLbgPrint.addEventListener("click", () => {
                const refreshFn = () => renderSingleWeekPaperHtml(state.currentWeek, false);
                openPreviewModal(
                    `Xem trước Lịch Báo Giảng Tuần ${state.currentWeek} (${currentOrientation === 'landscape' ? 'Khổ ngang' : 'Khổ đứng'})`,
                    refreshFn(),
                    () => exportLbgToDocx(false),
                    () => exportLbgToExcel(false),
                    () => printWithOrientation(currentOrientation),
                    refreshFn
                );
                setTimeout(() => {
                    printWithOrientation(currentOrientation);
                }, 250);
            });
        }

        // Tab 2 Actions (CTLOP) & Orientation
        const btnCtlopOrientPort = document.getElementById("btn-ctlop-orient-portrait");
        const btnCtlopOrientLand = document.getElementById("btn-ctlop-orient-landscape");
        if (btnCtlopOrientPort) btnCtlopOrientPort.addEventListener("click", () => setAppOrientation("portrait"));
        if (btnCtlopOrientLand) btnCtlopOrientLand.addEventListener("click", () => setAppOrientation("landscape"));

        const btnCtlopBatch = document.getElementById("btn-ctlop-batch");
        if (btnCtlopBatch) {
            btnCtlopBatch.addEventListener("click", () => openBatchExportModal('ctlop'));
        }

        document.getElementById("btn-ctlop-preview").addEventListener("click", () => {
            const refreshFn = () => renderSingleWeekPaperHtml(state.currentWeek, true);
            openPreviewModal(
                `Xem trước Lịch Báo Giảng Tích Hợp Tuần ${state.currentWeek} (${currentOrientation === 'landscape' ? 'Khổ ngang' : 'Khổ đứng'})`,
                refreshFn(),
                () => exportLbgToDocx(true),
                () => exportLbgToExcel(true),
                () => printWithOrientation(currentOrientation),
                refreshFn
            );
        });

        document.getElementById("btn-ctlop-excel").addEventListener("click", () => exportLbgToExcel(true));
        const btnCtlopPrint = document.getElementById("btn-ctlop-print");
        if (btnCtlopPrint) {
            btnCtlopPrint.addEventListener("click", () => {
                const refreshFn = () => renderSingleWeekPaperHtml(state.currentWeek, true);
                openPreviewModal(
                    `Xem trước Lịch Báo Giảng Tích Hợp Tuần ${state.currentWeek} (${currentOrientation === 'landscape' ? 'Khổ ngang' : 'Khổ đứng'})`,
                    refreshFn(),
                    () => exportLbgToDocx(true),
                    () => exportLbgToExcel(true),
                    () => printWithOrientation(currentOrientation),
                    refreshFn
                );
                setTimeout(() => {
                    printWithOrientation(currentOrientation);
                }, 250);
            });
        }

        // Modal Preview Orientation Buttons
        const btnModalOrientPort = document.getElementById("btn-modal-orient-portrait");
        const btnModalOrientLand = document.getElementById("btn-modal-orient-landscape");
        if (btnModalOrientPort) btnModalOrientPort.addEventListener("click", () => setAppOrientation("portrait"));
        if (btnModalOrientLand) btnModalOrientLand.addEventListener("click", () => setAppOrientation("landscape"));

        // Academic Year Sync Listeners
        const btnSyncYear = document.getElementById("btn-sync-academic-year");
        if (btnSyncYear) {
            btnSyncYear.addEventListener("click", () => {
                const yearVal = document.getElementById("set-academic-year").value;
                syncAcademicYear(yearVal);
            });
        }
        const quickYearSelect = document.getElementById("set-academic-year-quick");
        if (quickYearSelect) {
            quickYearSelect.addEventListener("change", (e) => {
                if (e.target.value) {
                    syncAcademicYear(e.target.value);
                }
            });
        }

        // Batch Modal Bindings
        const modalBatchCloseBtn = document.getElementById("modal-batch-close-btn");
        if (modalBatchCloseBtn) modalBatchCloseBtn.addEventListener("click", closeBatchExportModal);
        
        const modalBatchCancelBtn = document.getElementById("modal-batch-cancel-btn");
        if (modalBatchCancelBtn) modalBatchCancelBtn.addEventListener("click", closeBatchExportModal);

        document.querySelectorAll(".batch-range-btn").forEach(btn => {
            btn.addEventListener("click", (e) => {
                const s = e.currentTarget.dataset.start;
                const end = e.currentTarget.dataset.end;
                const selStart = document.getElementById("batch-start-week");
                const selEnd = document.getElementById("batch-end-week");
                if (selStart && selEnd) {
                    selStart.value = s;
                    selEnd.value = end;
                }
            });
        });

        const btnBatchDownloadDocx = document.getElementById("btn-batch-download-docx");
        if (btnBatchDownloadDocx) btnBatchDownloadDocx.addEventListener("click", exportBatchLbgToDocx);

        const btnBatchPreviewPrint = document.getElementById("btn-batch-preview-print");
        if (btnBatchPreviewPrint) btnBatchPreviewPrint.addEventListener("click", previewBatchLbg);

        // Global Grade Selector
        const globalGradeSel = document.getElementById("global-grade-selector");
        if (globalGradeSel) {
            globalGradeSel.value = (state.currentGrade || 5).toString();
            globalGradeSel.addEventListener("change", (e) => {
                switchGrade(parseInt(e.target.value, 10));
            });
        }

        // Tab 3 Smart Generator Actions
        document.getElementById("btn-run-week-generator").addEventListener("click", () => {
            const yr = document.getElementById("gen-year").value.trim() || "2026 - 2027";
            const sDate = document.getElementById("gen-start-date").value;
            const tetStart = document.getElementById("gen-tet-start-date") ? document.getElementById("gen-tet-start-date").value : "";
            const tetEnd = document.getElementById("gen-tet-end-date") ? document.getElementById("gen-tet-end-date").value : "";

            if (!sDate) {
                alert("Vui lòng chọn Ngày bắt đầu Tuần 1!");
                return;
            }
            if (tetStart && tetEnd && tetStart > tetEnd) {
                alert("Ngày kết thúc nghỉ Tết phải sau hoặc bằng ngày bắt đầu nghỉ Tết!");
                return;
            }

            autoGenerateCalendar(yr, sDate, tetStart, tetEnd);
        });

        document.getElementById("btn-save-lichtuan").addEventListener("click", () => {
            saveState();
            showToast("Đã lưu bảng Kế hoạch thời gian 35 tuần!", "success");
            renderTabLbg();
            renderTabCtlop();
        });

        const btnExportLichtuanExcel = document.getElementById("btn-export-lichtuan-excel");
        if (btnExportLichtuanExcel) {
            btnExportLichtuanExcel.addEventListener("click", exportLichtuanToExcel);
        }

        const btnPrintLichtuan = document.getElementById("btn-print-lichtuan");
        if (btnPrintLichtuan) {
            btnPrintLichtuan.addEventListener("click", () => window.print());
        }

        // Tab 4 (PPCT & KHDH) Actions
        document.getElementById("btn-preview-khdh").addEventListener("click", () => {
            previewKhdh(state.selectedKhdhSubjects);
        });

        const btnSaveBases = document.getElementById("btn-save-bases");
        if (btnSaveBases) {
            btnSaveBases.addEventListener("click", saveBasesFromTextarea);
        }

        const btnResetBases = document.getElementById("btn-reset-default-bases");
        if (btnResetBases) {
            btnResetBases.addEventListener("click", () => {
                if (confirm("Khôi phục danh sách căn cứ pháp lý mặc định?")) {
                    state.settings.bases = [...DEFAULT_BASES];
                    renderTabPpctBases();
                    saveState();
                    showToast("Đã khôi phục căn cứ pháp lý mặc định!", "success");
                }
            });
        }

        const inputImportBases = document.getElementById("input-import-bases-file");
        if (inputImportBases) {
            inputImportBases.addEventListener("change", (e) => {
                const file = e.target.files[0];
                if (file) {
                    importBasesFromFile(file);
                    inputImportBases.value = "";
                }
            });
        }

        document.getElementById("btn-export-khdh-docx").addEventListener("click", () => {
            if (!state.selectedKhdhSubjects || state.selectedKhdhSubjects.length === 0) {
                alert("Vui lòng chọn ít nhất 1 môn học để xuất file Word!");
                return;
            }
            const isMulti = state.selectedKhdhSubjects.length > 1;
            const subjectsData = state.selectedKhdhSubjects.map(subName => ({
                subjectName: subName,
                rows: state.khdh[subName] || []
            }));

            if (window.DocxGenerator && typeof window.DocxGenerator.generateKhdhDocx === 'function') {
                window.DocxGenerator.generateKhdhDocx({
                    settings: state.settings,
                    subjectsData: subjectsData,
                    isMultiSubject: isMulti
                }).then(blob => {
                    const fileName = isMulti 
                        ? `KHDH_Khoi_5_${state.selectedKhdhSubjects.length}_Mon.docx`
                        : `KHDH_Mon_${state.selectedKhdhSubjects[0]}_Lop_5.docx`;
                    if (typeof saveAs === 'function') {
                        saveAs(blob, fileName);
                    } else {
                        const url = URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = fileName;
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                        URL.revokeObjectURL(url);
                    }
                    showToast(`Đã xuất file Word: ${fileName}`, "success");
                }).catch(err => {
                    console.error("DOCX Export error:", err);
                    alert("Lỗi xuất file Word: " + err.message);
                });
            } else {
                alert("Bộ sinh file Word chưa sẵn sàng, vui lòng tải lại trang!");
            }
        });

        document.getElementById("btn-select-all-khdh").addEventListener("click", () => {
            state.selectedKhdhSubjects = Object.keys(state.khdh);
            renderTabPpct();
        });

        document.getElementById("btn-deselect-all-khdh").addEventListener("click", () => {
            state.selectedKhdhSubjects = [];
            renderTabPpct();
        });

        document.getElementById("ppct-filter-subject").addEventListener("change", renderTabPpct);
        document.getElementById("ppct-filter-week").addEventListener("change", renderTabPpct);
        document.getElementById("ppct-search").addEventListener("input", renderTabPpct);

        // PPCT Add Row and Save
        const btnAddPpct = document.getElementById("btn-add-ppct-row");
        if (btnAddPpct) {
            btnAddPpct.addEventListener("click", () => {
                const subSelect = document.getElementById("ppct-filter-subject");
                const weekSelect = document.getElementById("ppct-filter-week");
                const targetSub = subSelect && subSelect.value !== "ALL" ? subSelect.value : "Tiếng Việt";
                const targetWeek = weekSelect && weekSelect.value !== "ALL" ? parseInt(weekSelect.value) : 1;

                const subItems = state.ppct.filter(p => normalizeSubjectName(p.subject) === normalizeSubjectName(targetSub));
                const maxPpct = subItems.reduce((max, p) => Math.max(max, parseInt(p.ppct) || 0), 0);
                const weekItems = subItems.filter(p => p.week === targetWeek);
                const nextPeriodInWeek = weekItems.length + 1;

                const newRow = {
                    week: targetWeek,
                    subject: targetSub,
                    periodInWeek: nextPeriodInWeek,
                    ppct: maxPpct + 1,
                    lessonName: "Nhập tên bài dạy mới...",
                    integration: ""
                };

                state.ppct.push(newRow);
                saveState();
                renderTabPpct();
                renderTabLbg();
                renderTabCtlop();
                showToast(`Đã thêm tiết mới cho môn ${targetSub} (Tuần ${targetWeek})!`, "success");
            });
        }

        const btnSavePpct = document.getElementById("btn-save-ppct");
        if (btnSavePpct) {
            btnSavePpct.addEventListener("click", () => {
                saveState();
                renderTabPpct();
                renderTabLbg();
                renderTabCtlop();
                showToast("Đã lưu toàn bộ thay đổi phân phối chương trình thành công!", "success");
            });
        }

        // PPCT Excel Upload, Template Download, Export, and Reset
        document.getElementById("btn-download-ppct-template").addEventListener("click", downloadPpctTemplate);
        document.getElementById("btn-export-ppct-excel").addEventListener("click", exportPpctToExcel);
        
        const fileInput = document.getElementById("ppct-upload-file");
        document.getElementById("btn-upload-ppct").addEventListener("click", () => {
            fileInput.click();
        });

        fileInput.addEventListener("change", (e) => {
            const file = e.target.files[0];
            if (file) {
                uploadPpctFromFile(file);
                fileInput.value = "";
            }
        });

        document.getElementById("btn-reset-ppct").addEventListener("click", () => {
            if (confirm(`Bạn có chắc chắn muốn khôi phục lại phân phối chương trình gốc của Khối ${state.currentGrade}?`)) {
                let initPpct = [];
                if (state.currentGrade === 5) {
                    initPpct = (window.APP_INITIAL_DATA && window.APP_INITIAL_DATA.ppct) ? window.APP_INITIAL_DATA.ppct : [];
                } else if (window.APP_GRADE_DATA && window.APP_GRADE_DATA[state.currentGrade]) {
                    initPpct = window.APP_GRADE_DATA[state.currentGrade].ppct || [];
                }
                if (initPpct && initPpct.length) {
                    state.ppct = JSON.parse(JSON.stringify(initPpct));
                    saveState();
                    renderTabPpct();
                    renderTabLbg();
                    renderTabCtlop();
                    updatePpctCountBadge();
                    showToast(`Đã khôi phục dữ liệu phân phối chương trình chuẩn gốc Khối ${state.currentGrade}!`, "success");
                }
            }
        });

        const btnRecoverInteg = document.getElementById("btn-recover-integration");
        if (btnRecoverInteg) {
            btnRecoverInteg.addEventListener("click", () => {
                if (confirm("Bạn có chắc chắn muốn phục hồi toàn bộ nội dung tích hợp gốc (NLS, AI, QCN, QPAN, BVMT) cho Khối 5 không?")) {
                    restoreGrade5DefaultIntegration();
                }
            });
        }

        // PPCT Import Mode Modal Listeners
        const btnCloseImport = document.getElementById("btn-close-import-modal");
        const btnCancelImport = document.getElementById("btn-cancel-import");
        const btnConfirmImport = document.getElementById("btn-confirm-import");
        const modalImport = document.getElementById("modal-import-mode");

        if (btnCloseImport) {
            btnCloseImport.addEventListener("click", () => {
                window._pendingPpctImport = null;
                if (modalImport) modalImport.style.display = "none";
            });
        }
        if (btnCancelImport) {
            btnCancelImport.addEventListener("click", () => {
                window._pendingPpctImport = null;
                if (modalImport) modalImport.style.display = "none";
            });
        }
        if (btnConfirmImport) {
            btnConfirmImport.addEventListener("click", () => {
                const selectedMode = document.querySelector('input[name="import-mode"]:checked');
                const mode = selectedMode ? selectedMode.value : "replace";
                executePendingImport(mode);
            });
        }
        if (modalImport) {
            modalImport.addEventListener("click", (e) => {
                if (e.target === modalImport) {
                    window._pendingPpctImport = null;
                    modalImport.style.display = "none";
                }
            });
        }

        // Tab 5 Settings Buttons
        const btnAddPht = document.getElementById("btn-add-pht");
        if (btnAddPht) {
            btnAddPht.addEventListener("click", () => {
                if (!Array.isArray(state.settings.vicePrincipals)) {
                    state.settings.vicePrincipals = [state.settings.vicePrincipal || "Lê Văn Tám"];
                }
                const newIdx = state.settings.vicePrincipals.length + 1;
                state.settings.vicePrincipals.push(`Phó Hiệu trưởng ${newIdx}`);
                renderPhtListInputs();
                renderBghSignerDropdown();
                showToast("Đã thêm 1 Phó Hiệu trưởng mới vào danh sách!", "info");
            });
        }

        // Tab 5 Master Timetable Print & Export Buttons
        const btnPrintMasterTt = document.getElementById("btn-print-master-tkb");
        if (btnPrintMasterTt) {
            btnPrintMasterTt.addEventListener("click", previewMasterTimetablePrint);
        }

        const btnDocxMasterTt = document.getElementById("btn-docx-master-tkb");
        if (btnDocxMasterTt) {
            btnDocxMasterTt.addEventListener("click", exportMasterTimetableDocx);
        }

        const btnXlsxMasterTt = document.getElementById("btn-xlsx-master-tkb");
        if (btnXlsxMasterTt) {
            btnXlsxMasterTt.addEventListener("click", exportMasterTimetableXlsx);
        }

        const btnSaveMasterTt = document.getElementById("btn-save-master-timetable");
        if (btnSaveMasterTt) {
            btnSaveMasterTt.addEventListener("click", () => {
                saveState();
                renderTabLbg();
                renderTabCtlop();
                showToast("Đã lưu Thời khóa biểu gốc thành công và áp dụng cho 35 tuần!", "success");
            });
        }

        const btnResetMasterTt = document.getElementById("btn-reset-master-timetable");
        if (btnResetMasterTt) {
            btnResetMasterTt.addEventListener("click", () => {
                if (confirm("Bạn có chắc chắn muốn khôi phục lại Thời khóa biểu gốc chuẩn mặc định cho " + (state.settings.grade || `Khối ${state.currentGrade}`) + " không?")) {
                    state.timetable = getGradeDefaultTimetable(state.currentGrade);
                    state.weeklyCustomSlots = {};
                    saveState();
                    renderMasterTimetableEditor();
                    renderTabLbg();
                    renderTabCtlop();
                    showToast("Đã khôi phục Thời khóa biểu gốc chuẩn mặc định!", "success");
                }
            });
        }

        document.getElementById("btn-save-settings").addEventListener("click", saveSettingsFromForm);
        document.getElementById("btn-backup-data").addEventListener("click", () => {
            const currentGradeKey = getStorageKey(state.currentGrade);
            const rawData = localStorage.getItem(currentGradeKey) || localStorage.getItem(STORAGE_KEY) || localStorage.getItem("LBG_APP_DATA_V7_G" + state.currentGrade) || localStorage.getItem("LBG_APP_DATA_V7") || localStorage.getItem("LBG_APP_DATA_V6_G" + state.currentGrade) || localStorage.getItem("LBG_APP_DATA_V6");
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(rawData || "{}");
            const dlAnchor = document.createElement('a');
            dlAnchor.setAttribute("href", dataStr);
            const safeClassName = (state.settings.className || `Khoi_${state.currentGrade}`).replace(/[^a-zA-Z0-9_\u00C0-\u024F\u1EA0-\u1EF9]/g, '_');
            dlAnchor.setAttribute("download", `Sao_Luu_LBG_${safeClassName}_${new Date().toISOString().slice(0,10)}.json`);
            dlAnchor.click();
            dlAnchor.remove();
            showToast(`Đã xuất file sao lưu toàn bộ dữ liệu Khối ${state.currentGrade} (${state.settings.className})!`, "success");
        });

        document.getElementById("btn-restore-data").addEventListener("click", () => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.json';
            input.onchange = e => {
                const file = e.target.files[0];
                const reader = new FileReader();
                reader.readAsText(file, 'UTF-8');
                reader.onload = readerEvent => {
                    try {
                        const content = readerEvent.target.result;
                        const parsed = JSON.parse(content);
                        if (!parsed || typeof parsed !== 'object') {
                            throw new Error("Cấu trúc file JSON không đúng định dạng!");
                        }
                        const currentGradeKey = getStorageKey(state.currentGrade);
                        localStorage.setItem(currentGradeKey, content);
                        if (state.currentGrade === 5) {
                            localStorage.setItem(STORAGE_KEY, content);
                        }
                        loadStateForGrade(state.currentGrade);
                        renderTabLbg();
                        renderTabCtlop();
                        renderTabSettings();
                        renderTabLichtuan();
                        renderTabPpct();
                        updatePpctCountBadge();
                        showToast(`Khôi phục toàn bộ dữ liệu Khối ${state.currentGrade} thành công!`, "success");
                    } catch (err) {
                        alert("File sao lưu không hợp lệ: " + err.message);
                    }
                };
            };
            input.click();
        });

        // Tab 5 Subject Management Buttons
        const btnOpenAddSub = document.getElementById("btn-open-add-subject");
        if (btnOpenAddSub) btnOpenAddSub.addEventListener("click", () => openSubjectEditorModal(null));

        const btnSubClose = document.getElementById("modal-subject-close-btn");
        if (btnSubClose) btnSubClose.addEventListener("click", closeSubjectEditorModal);

        const btnSubCancel = document.getElementById("modal-subject-cancel-btn");
        if (btnSubCancel) btnSubCancel.addEventListener("click", closeSubjectEditorModal);

        const btnSubSave = document.getElementById("modal-subject-save-btn");
        if (btnSubSave) btnSubSave.addEventListener("click", saveSubjectFromModal);

        const btnOpenRenameSub = document.getElementById("btn-open-rename-subject");
        if (btnOpenRenameSub) btnOpenRenameSub.addEventListener("click", openRenameSubjectModal);

        const btnResetDefSubs = document.getElementById("btn-reset-default-subjects");
        if (btnResetDefSubs) btnResetDefSubs.addEventListener("click", resetDefaultSubjects);

        const btnRenameSubClose = document.getElementById("modal-rename-close-btn");
        if (btnRenameSubClose) btnRenameSubClose.addEventListener("click", closeRenameSubjectModal);

        const btnRenameSubCancel = document.getElementById("modal-rename-cancel-btn");
        if (btnRenameSubCancel) btnRenameSubCancel.addEventListener("click", closeRenameSubjectModal);

        const btnRenameSubSave = document.getElementById("modal-rename-save-btn");
        if (btnRenameSubSave) btnRenameSubSave.addEventListener("click", applyBulkRenameSubject);

        // VERSION 6.0: Tab Báo Giảng Theo Môn Actions & Orientation
        const btnOrientPortMon = document.getElementById("btn-orient-portrait-lbgmon");
        const btnOrientLandMon = document.getElementById("btn-orient-landscape-lbgmon");
        if (btnOrientPortMon) {
            btnOrientPortMon.addEventListener("click", () => {
                state.lbgMonOrientation = "portrait";
                btnOrientPortMon.classList.add("active");
                if (btnOrientLandMon) btnOrientLandMon.classList.remove("active");
                renderTabLbgMon();
            });
        }
        if (btnOrientLandMon) {
            btnOrientLandMon.addEventListener("click", () => {
                state.lbgMonOrientation = "landscape";
                btnOrientLandMon.classList.add("active");
                if (btnOrientPortMon) btnOrientPortMon.classList.remove("active");
                renderTabLbgMon();
            });
        }

        const btnLbgMonPreview = document.getElementById("btn-lbgmon-preview");
        if (btnLbgMonPreview) btnLbgMonPreview.addEventListener("click", previewLbgMonA4);

        const btnLbgMonDocx = document.getElementById("btn-lbgmon-docx");
        if (btnLbgMonDocx) btnLbgMonDocx.addEventListener("click", exportLbgMonDocx);

        const btnLbgMonExcel = document.getElementById("btn-lbgmon-excel");
        if (btnLbgMonExcel) btnLbgMonExcel.addEventListener("click", exportLbgMonXlsx);

        const btnLbgMonPrint = document.getElementById("btn-lbgmon-print");
        if (btnLbgMonPrint) btnLbgMonPrint.addEventListener("click", () => {
            previewLbgMonA4();
        });

        const btnLbgMonBatch = document.getElementById("btn-lbgmon-batch");
        if (btnLbgMonBatch) btnLbgMonBatch.addEventListener("click", () => {
            openBatchExportModal('lbg-mon');
        });

        // Quick Navigation buttons
        document.querySelectorAll(".btn-nav-to-lbgmon").forEach(btn => {
            btn.addEventListener("click", () => {
                switchActiveTab("tab-lbg-mon");
            });
        });

        document.querySelectorAll(".btn-nav-to-settings").forEach(btn => {
            btn.addEventListener("click", () => {
                switchActiveTab("tab-settings");
            });
        });

        document.querySelectorAll(".btn-nav-to-guide").forEach(btn => {
            btn.addEventListener("click", () => {
                switchActiveTab("tab-guide");
            });
        });

        // Bulk Rename Modal Events (Tab 4 & General)
        const btnRenameBulk = document.getElementById("btn-rename-subject-bulk");
        if (btnRenameBulk) btnRenameBulk.addEventListener("click", openRenameSubjectModal);

        // Subject Inclusion Quick Filter Buttons
        const btnInclAll = document.getElementById("btn-incl-all");
        if (btnInclAll) {
            btnInclAll.addEventListener("click", () => {
                state.includedSubjects = getAllUniqueSubjects();
                saveState();
                renderLbgInclusionChips();
                renderTabLbg();
                renderTabCtlop();
                showToast("Đã chọn đưa TẤT CẢ các môn học vào Lịch Báo Giảng!", "success");
            });
        }

        const btnInclGvcn = document.getElementById("btn-incl-gvcn");
        if (btnInclGvcn) {
            btnInclGvcn.addEventListener("click", () => {
                const gvcnSubs = ["Tiếng Việt", "Toán", "HĐ Trải nghiệm", "Khoa học", "LS&ĐL", "Đạo đức", "Công nghệ", "Đọc Thư viện", "TC Tiếng Việt", "TC Toán", "KNS", "STEM", "CD Số"];
                state.includedSubjects = getAllUniqueSubjects().filter(s => gvcnSubs.some(g => normalizeSubjectName(g) === normalizeSubjectName(s)));
                saveState();
                renderLbgInclusionChips();
                renderTabLbg();
                renderTabCtlop();
                showToast("Đã chọn chỉ đưa các môn GVCN dạy vào Lịch Báo Giảng!", "success");
            });
        }

        const btnInclSpec = document.getElementById("btn-incl-specialist");
        if (btnInclSpec) {
            btnInclSpec.addEventListener("click", () => {
                const specSubs = ["Tiếng Anh", "Tin học", "GD Thể chất", "Âm nhạc", "Mĩ thuật"];
                state.includedSubjects = getAllUniqueSubjects().filter(s => specSubs.some(g => normalizeSubjectName(g) === normalizeSubjectName(s)));
                saveState();
                renderLbgInclusionChips();
                renderTabLbg();
                renderTabCtlop();
                showToast("Đã chọn chỉ đưa các môn GV Chuyên dạy vào Lịch Báo Giảng!", "success");
            });
        }

        // Modal Close
        document.getElementById("modal-close-btn").addEventListener("click", closePreviewModal);
        document.getElementById("modal-close-footer").addEventListener("click", closePreviewModal);

        // Realtime Top Header Badge Update Listeners
        const inputSchool = document.getElementById("set-school-name");
        if (inputSchool) {
            inputSchool.addEventListener("input", (e) => {
                state.settings.schoolName = e.target.value;
                updateTopHeaderBadge();
            });
        }

        const inputClass = document.getElementById("set-class-name");
        if (inputClass) {
            inputClass.addEventListener("input", (e) => {
                state.settings.className = e.target.value;
                updateTopHeaderBadge();
            });
        }

        // Initial Render
        updateTopHeaderBadge();
        renderTabLbg();
        renderTabSettings();
        renderTabLichtuan();
        renderTabPpct();
        // Expose functions for debugging / testing
        window.calculateWeekSchedule = calculateWeekSchedule;
        window.calculateWeekScheduleBySubject = calculateWeekScheduleBySubject;
        window.renderTabLbgMon = renderTabLbgMon;
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", initApp);
    } else {
        initApp();
    }
})();
