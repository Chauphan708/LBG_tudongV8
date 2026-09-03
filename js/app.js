/**
 * LỊCH BÁO GIẢNG & KẾ HOẠCH DẠY HỌC LỚP 5 - MAIN APPLICATION
 * Features: 5 Tabs, Realtime Search & Calculation, Friday Afternoon Customization,
 * Smart 35-Week Generator, PPCT Excel Import/Export Template, Decree 30/2020 Preview & DOCX/XLSX Export
 */

(function() {
    const STORAGE_KEY = "LBG_APP_DATA_V3";

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
    
    let state = {
        currentTab: "tab-lbg",
        currentWeek: 1,
        selectedKhdhSubjects: ["TIẾNG VIỆT", "TOÁN", "KHOA HỌC", "LS&ĐL", "HĐ TRẢI NGHIỆM", "ĐẠO ĐỨC", "CÔNG NGHỆ"],
        includedSubjects: [],
        subjectList: [],
        lbgExcludedMode: "off",
        settings: {
            governingBody: "UBND PHƯỜNG TRUNG NHỨT",
            schoolName: "TRƯỜNG TIỂU HỌC TRUNG NHỨT",
            grade: "KHỐI 5",
            className: "LỚP 5A",
            academicYear: "2026 - 2027",
            homeroomTeacher: "Nguyễn Thị Thu Hà",
            headOfGrade: "Trần Thị Mai",
            vicePrincipal: "Lê Văn Tám",
            vicePrincipals: ["Lê Văn Tám"],
            principal: "Phạm Quốc Hùng",
            bghSignerType: "PHT",
            bghSignerIndex: 0,
            bghSignerLbgType: "PHT",
            bghSignerLbgIndex: 0,
            bghSignerKhdhType: "HT",
            bghSignerKhdhIndex: 0,
            location: "Trung Nhứt",
            dateString: "ngày 28 tháng 8 năm 2026",
            bases: [
                "Căn cứ Thông tư số 28/2020/TT-BGDĐT ngày 04/9/2020 của Bộ trưởng Bộ Giáo dục và Đào tạo ban hành Điều lệ Trường tiểu học;",
                "Căn cứ Thông tư số 32/2018/TT-BGDĐT ngày 26/12/2018 của Bộ trưởng Bộ Giáo dục và Đào tạo ban hành Chương trình Giáo dục phổ thông 2018;",
                "Căn cứ Quyết định ban hành Khung kế hoạch thời gian năm học 2026 - 2027 của Ủy ban nhân dân Thành phố/Tỉnh;",
                "Căn cứ Kế hoạch giáo dục nhà trường năm học 2026 - 2027 của Trường Tiểu học Trung Nhứt;"
            ]
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

    function loadState() {
        if (window.APP_INITIAL_DATA) {
            state.weeks = JSON.parse(JSON.stringify(window.APP_INITIAL_DATA.weeks || []));
            state.ppct = JSON.parse(JSON.stringify(window.APP_INITIAL_DATA.ppct || []));
            state.khdh = JSON.parse(JSON.stringify(window.APP_INITIAL_DATA.khdh || {}));
            state.timetable = JSON.parse(JSON.stringify(window.APP_INITIAL_DATA.settings?.timetable || []));
            Object.assign(state.settings, window.APP_INITIAL_DATA.settings || {});
        }

        try {
            let saved = localStorage.getItem(STORAGE_KEY);
            if (!saved) {
                // Seamless migration from V2 or V1
                saved = localStorage.getItem("LBG_APP_DATA_V2") || localStorage.getItem("LBG_APP_DATA_V1");
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
                if (parsed.ppct && parsed.ppct.length) state.ppct = parsed.ppct;
                if (parsed.timetable && parsed.timetable.length) state.timetable = parsed.timetable;
                if (parsed.weeklyCustomSlots) state.weeklyCustomSlots = parsed.weeklyCustomSlots;
                if (parsed.weeklyScheduleOverrides) state.weeklyScheduleOverrides = parsed.weeklyScheduleOverrides;
                if (parsed.weeklyFridayPeriods) state.weeklyFridayPeriods = parsed.weeklyFridayPeriods;
                if (parsed.currentWeek) state.currentWeek = parsed.currentWeek;
                if (parsed.fridayPeriodsDefault !== undefined) state.fridayPeriodsDefault = parsed.fridayPeriodsDefault;
                if (parsed.includedSubjects && parsed.includedSubjects.length) state.includedSubjects = parsed.includedSubjects;
                if (parsed.subjectList && Array.isArray(parsed.subjectList) && parsed.subjectList.length) {
                    state.subjectList = parsed.subjectList;
                }
                if (parsed.lbgExcludedMode) state.lbgExcludedMode = parsed.lbgExcludedMode;
            }
        } catch (e) {
            console.warn("Could not load localStorage:", e);
        }

        if (!state.subjectList || state.subjectList.length === 0) {
            state.subjectList = JSON.parse(JSON.stringify(DEFAULT_GRADE_5_SUBJECTS));
        }

        // Standardize all PPCT & Timetable subjects to clean Title Case & Auto-hydrate integration & normalize punctuation spacing
        if (state.ppct) {
            const initPpct = (window.APP_INITIAL_DATA && Array.isArray(window.APP_INITIAL_DATA.ppct)) ? window.APP_INITIAL_DATA.ppct : [];
            state.ppct.forEach(p => {
                p.subject = normalizeSubjectName(p.subject);
                if (p.lessonName) p.lessonName = normalizePunctuationSpacing(p.lessonName);
                if (p.integration) p.integration = normalizePunctuationSpacing(p.integration);
                if (!p.integration || !p.integration.trim()) {
                    const normSub = normalizeSubjectName(p.subject);
                    const matched = initPpct.find(ip => ip.week === p.week && normalizeSubjectName(ip.subject) === normSub && ip.periodInWeek === p.periodInWeek);
                    if (matched && matched.integration) {
                        p.integration = normalizePunctuationSpacing(matched.integration);
                    }
                }
            });
        }
        if (state.settings && state.settings.bases) {
            state.settings.bases = state.settings.bases.map(b => normalizePunctuationSpacing(b));
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
                lbgExcludedMode: state.lbgExcludedMode
            };
            localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
        } catch (e) {
            console.warn("Could not save to localStorage:", e);
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

    function getIntegrationFromKhdh(subject, week, periodInWeek) {
        if (!state.khdh) return "";
        const normSub = normalizeSubjectName(subject);
        let khdhKey = null;
        if (normSub === "Toán") khdhKey = "TOÁN";
        else if (normSub === "Tiếng Việt") khdhKey = "TIẾNG VIỆT";
        else if (normSub === "Khoa học") khdhKey = "KHOA HỌC";
        else if (normSub === "LS&ĐL") khdhKey = "LS&ĐL";
        else if (normSub === "Đạo đức") khdhKey = "ĐẠO ĐỨC";
        else if (normSub === "Công nghệ") khdhKey = "CÔNG NGHỆ";
        else if (normSub === "HĐ Trải nghiệm") khdhKey = "HĐ TRẢI NGHIỆM";

        if (!khdhKey || !state.khdh[khdhKey]) return "";

        const weekRows = state.khdh[khdhKey].filter(r => r.week === week);
        if (weekRows.length === 0) return "";

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

        return matchedRow && matchedRow.integration ? matchedRow.integration.trim() : "";
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
                const slot = slots.find(s => s.day === d && s.session === "Sáng" && s.period === p);
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
                const slot = slots.find(s => s.day === d && s.session === "Chiều" && s.period === p);
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
            const morningSlots = daySlots.filter(s => s.session === "Sáng");
            const afternoonSlots = daySlots.filter(s => s.session === "Chiều");
            const totalRows = daySlots.length;

            daySlots.forEach((slot, idx) => {
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
                const slot = slots.find(item => item.day === d && item.session === "Sáng" && item.period === p);
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
                const slot = slots.find(item => item.day === d && item.session === "Chiều" && item.period === p);
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
        const timetable = getWeekSlots(weekNum);
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
                
                // Absolute 100% accurate integration resolution: Override -> PPCT -> KHDH Fallback
                if (override.integration !== undefined) {
                    integration = override.integration;
                } else if (ppctItem && ppctItem.integration && ppctItem.integration.trim()) {
                    integration = ppctItem.integration.trim();
                } else {
                    integration = getIntegrationFromKhdh(canonicalSub, weekNum, periodInWeek);
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

        const tbody = document.getElementById("lbg-table-body");
        tbody.innerHTML = "";

        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];
        const dynamicSubjectOptions = getAllUniqueSubjects().concat(["-- Nghỉ / Để trống --"]);
        
        days.forEach(day => {
            const isAltDay = (day === "Thứ 2" || day === "Thứ 4" || day === "Thứ 6");
            const daySlots = schedule.filter(s => s.day === day);
            const morningSlots = daySlots.filter(s => s.session === "Sáng");
            const afternoonSlots = daySlots.filter(s => s.session === "Chiều");
            const totalRows = daySlots.length;

            daySlots.forEach((slot, idx) => {
                const tr = document.createElement("tr");
                if (isAltDay) tr.classList.add("row-day-alt");
                if (slot.isOff) tr.classList.add("row-empty-period");

                let dayCellHtml = "";
                if (idx === 0) {
                    dayCellHtml = `<td rowspan="${totalRows}" class="col-day ${isAltDay ? 'day-alt' : 'day-normal'}">${day}</td>`;
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
                    <td class="col-period">${slot.period}</td>
                    <td class="col-subject">
                        <select class="form-select form-select-sm subject-selector" data-key="${slot.key}">
                            ${dynamicSubjectOptions.map(s => `<option value="${s}" ${normalizeSubjectName(s) === normalizeSubjectName(slot.subject) || s === slot.subject ? 'selected' : ''}>${s}</option>`).join('')}
                        </select>
                    </td>
                    <td class="col-ppct editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="ppct">${slot.ppct}</td>
                    <td class="col-lesson editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="lessonName">${slot.lessonName}</td>
                    <td class="no-print" style="text-align: center;">
                        <div class="row-actions-group">
                            <button type="button" class="btn-row-action btn-add-week-slot" data-day="${slot.day}" data-session="${slot.session}" data-period="${slot.period}" title="Chèn thêm 1 tiết ngay phía dưới">+</button>
                            <button type="button" class="btn-row-action btn-del-week-slot" data-day="${slot.day}" data-session="${slot.session}" data-period="${slot.period}" title="Xóa tiết học này">🗑️</button>
                        </div>
                    </td>
                `;
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
                e.target.innerText = val;
                if (!state.weeklyScheduleOverrides[key]) state.weeklyScheduleOverrides[key] = {};
                state.weeklyScheduleOverrides[key][field] = val;
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

        const tbody = document.getElementById("ctlop-table-body");
        tbody.innerHTML = "";

        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];
        const dynamicSubjectOptions = getAllUniqueSubjects().concat(["-- Nghỉ / Để trống --"]);
        
        days.forEach(day => {
            const isAltDay = (day === "Thứ 2" || day === "Thứ 4" || day === "Thứ 6");
            const daySlots = schedule.filter(s => s.day === day);
            const morningSlots = daySlots.filter(s => s.session === "Sáng");
            const afternoonSlots = daySlots.filter(s => s.session === "Chiều");
            const totalRows = daySlots.length;

            daySlots.forEach((slot, idx) => {
                const tr = document.createElement("tr");
                if (isAltDay) tr.classList.add("row-day-alt");
                if (slot.isOff) tr.classList.add("row-empty-period");

                let dayCellHtml = "";
                if (idx === 0) {
                    dayCellHtml = `<td rowspan="${totalRows}" class="col-day ${isAltDay ? 'day-alt' : 'day-normal'}">${day}</td>`;
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
                    <td class="col-period">${slot.period}</td>
                    <td class="col-subject">
                        <select class="form-select form-select-sm subject-selector" data-key="${slot.key}">
                            ${dynamicSubjectOptions.map(s => `<option value="${s}" ${normalizeSubjectName(s) === normalizeSubjectName(slot.subject) || s === slot.subject ? 'selected' : ''}>${s}</option>`).join('')}
                        </select>
                    </td>
                    <td class="col-ppct editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="ppct">${slot.ppct}</td>
                    <td class="col-lesson editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="lessonName">${slot.lessonName}</td>
                    <td class="col-integration editable-cell" contenteditable="${!slot.isOff}" data-key="${slot.key}" data-field="integration" style="line-height:1.45;">${escapeHtml(slot.integration || '').replace(/\n/g, '<br>')}</td>
                    <td class="no-print" style="text-align: center;">
                        <div class="row-actions-group">
                            <button type="button" class="btn-row-action btn-add-week-slot" data-day="${slot.day}" data-session="${slot.session}" data-period="${slot.period}" title="Chèn thêm 1 tiết ngay phía dưới">+</button>
                            <button type="button" class="btn-row-action btn-del-week-slot" data-day="${slot.day}" data-session="${slot.session}" data-period="${slot.period}" title="Xóa tiết học này">🗑️</button>
                        </div>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        });

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
    function autoGenerateCalendar(yearString, startDateString, tetBreakWeeks) {
        if (!startDateString) return;
        const startDt = new Date(startDateString);
        if (isNaN(startDt.getTime())) return;

        const newWeeks = [];
        let currStart = new Date(startDt.getTime());

        for (let w = 1; w <= 35; w++) {
            const term = w <= 18 ? 1 : 2;
            const currEnd = new Date(currStart.getTime() + 4 * 24 * 60 * 60 * 1000);

            const formatVN = (d) => {
                const day = String(d.getDate()).padStart(2, '0');
                const month = String(d.getMonth() + 1).padStart(2, '0');
                const yr = d.getFullYear();
                return `${day}/${month}/${yr}`;
            };

            const monthStr = `Tháng ${currStart.getMonth() + 1} năm ${currStart.getFullYear()}`;
            
            let holiday = "";
            let note = "";

            const m = currStart.getMonth() + 1;
            const startDay = currStart.getDate();
            const endDay = currEnd.getDate();

            if (m === 9 && startDay <= 5 && endDay >= 2) {
                holiday = "Lễ Quốc khánh (02/09)";
                note = "Khai giảng năm học (05/09)";
            } else if (m === 11 && startDay <= 20 && endDay >= 20) {
                holiday = "Ngày Nhà giáo Việt Nam (20/11)";
            } else if ((m === 12 && endDay >= 31) || (m === 1 && startDay <= 1)) {
                holiday = "Tết Dương lịch (01/01)";
            } else if (m === 4 && startDay <= 30 && endDay >= 30) {
                holiday = "Nghỉ lễ 30/4 và 1/5";
            }

            if (w === 35) {
                note = "Tổng kết năm học & Báo cáo hồ sơ";
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

            if (w === 20 && tetBreakWeeks > 0) {
                currStart = new Date(currStart.getTime() + tetBreakWeeks * 7 * 24 * 60 * 60 * 1000);
            }
        }

        state.weeks = newWeeks;
        state.settings.academicYear = yearString;
        saveState();
        renderTabLichtuan();
        renderTabLbg();
        renderTabCtlop();
        showToast(`Đã tự động tính toán lại toàn bộ 35 tuần năm học ${yearString}!`, "success");
    }

    // Render Tab 3: Kế hoạch Thời gian (Lịch Tuần)
    function renderTabLichtuan() {
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

    function uploadPpctFromExcel(file) {
        if (typeof XLSX === "undefined") {
            alert("Thư viện SheetJS chưa sẵn sàng!");
            return;
        }

        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                if (!jsonRows || jsonRows.length <= 1) {
                    alert("File Excel không có dữ liệu!");
                    return;
                }

                // Identify header row
                let headerRowIdx = -1;
                let colMap = { week: -1, subject: -1, periodInWeek: -1, ppct: -1, lesson: -1, integration: -1 };

                for (let r = 0; r < Math.min(10, jsonRows.length); r++) {
                    const row = jsonRows[r] || [];
                    for (let c = 0; c < row.length; c++) {
                        const val = (row[c] || "").toString().normalize("NFC").toLowerCase();
                        if (val.includes("tuần") || val.includes("tuan") || val === "w") colMap.week = c;
                        if (val.includes("môn") || val.includes("mon") || val.includes("subject")) colMap.subject = c;
                        if (val.includes("tiết/tuần") || val.includes("tiết trong tuần") || val.includes("tiet/t") || val.includes("tiet thu")) colMap.periodInWeek = c;
                        if (val.includes("ppct") || val.includes("tiết ppct") || val.includes("tiết thứ")) colMap.ppct = c;
                        if (val.includes("bài") || val.includes("tên bài") || val.includes("lesson") || val.includes("nội dung")) colMap.lesson = c;
                        if (val.includes("tích hợp") || val.includes("điều chỉnh") || val.includes("ghi chú") || val.includes("integration")) colMap.integration = c;
                    }
                    if (colMap.subject !== -1 && colMap.lesson !== -1) {
                        headerRowIdx = r;
                        break;
                    }
                }

                // Default column indices fallback
                if (colMap.week === -1) colMap.week = 0;
                if (colMap.subject === -1) colMap.subject = 1;
                if (colMap.periodInWeek === -1) colMap.periodInWeek = 2;
                if (colMap.ppct === -1) colMap.ppct = 3;
                if (colMap.lesson === -1) colMap.lesson = 4;
                if (colMap.integration === -1) colMap.integration = 5;

                const startRow = headerRowIdx !== -1 ? headerRowIdx + 1 : 1;
                let addedCount = 0;
                let currentW = 1;
                let currentSub = "";

                for (let r = startRow; r < jsonRows.length; r++) {
                    const row = jsonRows[r];
                    if (!row || row.length === 0) continue;

                    const wVal = row[colMap.week];
                    if (wVal && !isNaN(parseInt(wVal))) currentW = parseInt(wVal);

                    const subVal = row[colMap.subject];
                    if (subVal && subVal.toString().trim()) {
                        currentSub = normalizeSubjectName(subVal.toString().trim());
                    }

                    const lessonVal = row[colMap.lesson] ? normalizePunctuationSpacing(row[colMap.lesson].toString().trim()) : "";
                    if (!lessonVal || lessonVal.includes("Tên bài")) continue;

                    let pInWeekVal = row[colMap.periodInWeek] ? parseInt(row[colMap.periodInWeek]) : 1;
                    if (isNaN(pInWeekVal) || pInWeekVal <= 0) pInWeekVal = 1;

                    let ppctVal = row[colMap.ppct] ? parseInt(row[colMap.ppct]) : (state.ppct.length + 1);
                    if (isNaN(ppctVal)) ppctVal = state.ppct.length + 1;

                    const integVal = row[colMap.integration] ? normalizePunctuationSpacing(row[colMap.integration].toString().trim()) : "";

                    // Upsert or Add
                    const existingIdx = state.ppct.findIndex(p => p.week === currentW && normalizeSubjectName(p.subject) === currentSub && p.periodInWeek === pInWeekVal);
                    if (existingIdx >= 0) {
                        state.ppct[existingIdx].ppct = ppctVal;
                        state.ppct[existingIdx].lessonName = lessonVal;
                        if (integVal) state.ppct[existingIdx].integration = integVal;
                    } else {
                        state.ppct.push({
                            week: currentW,
                            subject: currentSub,
                            periodInWeek: pInWeekVal,
                            ppct: ppctVal,
                            lessonName: lessonVal,
                            integration: integVal
                        });
                    }
                    addedCount++;
                }

                saveState();
                renderTabPpct();
                renderTabLbg();
                renderTabCtlop();
                showToast(`Đã nạp thành công ${addedCount} tiết phân phối chương trình từ Excel!`, "success");
            } catch (err) {
                console.error(err);
                alert("Lỗi đọc file Excel: " + err.message);
            }
        };
        reader.readAsArrayBuffer(file);
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

        document.getElementById("ppct-count-badge").innerText = `${filtered.length} tiết (Tuần 1 - 35)`;

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

        const defaultBases = [
            "Căn cứ Thông tư số 28/2020/TT-BGDĐT ngày 04/9/2020 của Bộ trưởng Bộ Giáo dục và Đào tạo ban hành Điều lệ Trường tiểu học;",
            "Căn cứ Thông tư số 32/2018/TT-BGDĐT ngày 26/12/2018 của Bộ trưởng Bộ Giáo dục và Đào tạo ban hành Chương trình Giáo dục phổ thông 2018;",
            `Căn cứ Quyết định ban hành Khung kế hoạch thời gian năm học ${state.settings.academicYear || '2026 - 2027'} của Ủy ban nhân dân Thành phố/Tỉnh;`,
            `Căn cứ Kế hoạch giáo dục nhà trường năm học ${state.settings.academicYear || '2026 - 2027'} của ${state.settings.schoolName || 'Trường Tiểu học Trung Nhứt'};`
        ];

        if (!Array.isArray(state.settings.bases) || state.settings.bases.length === 0) {
            state.settings.bases = defaultBases;
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

        state.subjectList = JSON.parse(JSON.stringify(DEFAULT_GRADE_5_SUBJECTS));
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

    // Render Tab 5: Cài đặt
    function renderTabSettings() {
        document.getElementById("set-governing-body").value = state.settings.governingBody || "";
        document.getElementById("set-school-name").value = state.settings.schoolName || "";
        document.getElementById("set-grade").value = state.settings.grade || "KHỐI 5";
        document.getElementById("set-class-name").value = state.settings.className || "LỚP 5A";
        document.getElementById("set-academic-year").value = state.settings.academicYear || "2026 - 2027";
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
            const lbgType = state.settings.bghSignerLbgType || state.settings.bghSignerType || "PHT";
            const lbgIdx = (state.settings.bghSignerLbgIndex !== undefined) ? state.settings.bghSignerLbgIndex : (state.settings.bghSignerIndex || 0);

            let lbgHtml = `<option value="HT:0" ${lbgType === 'HT' ? 'selected' : ''}>👨‍💼 Hiệu trưởng: ${htName}</option>`;
            vps.forEach((pht, idx) => {
                const isSel = (lbgType === 'PHT' && lbgIdx === idx);
                lbgHtml += `<option value="PHT:${idx}" ${isSel ? 'selected' : ''}>👥 Phó Hiệu trưởng ${idx + 1}: ${pht || ('PHT ' + (idx + 1))}</option>`;
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
            };
        }

        // 2. KHDH Signer Select
        if (selectKhdh) {
            const khdhType = state.settings.bghSignerKhdhType || "HT";
            const khdhIdx = (state.settings.bghSignerKhdhIndex !== undefined) ? state.settings.bghSignerKhdhIndex : 0;

            let khdhHtml = `<option value="HT:0" ${khdhType === 'HT' ? 'selected' : ''}>👨‍💼 Hiệu trưởng: ${htName}</option>`;
            vps.forEach((pht, idx) => {
                const isSel = (khdhType === 'PHT' && khdhIdx === idx);
                khdhHtml += `<option value="PHT:${idx}" ${isSel ? 'selected' : ''}>👥 Phó Hiệu trưởng ${idx + 1}: ${pht || ('PHT ' + (idx + 1))}</option>`;
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

    // 6. Modal Preview System
    function openPreviewModal(title, contentHtml, onDownloadDocx, onDownloadXlsx, onPrint) {
        const modal = document.getElementById("preview-modal");
        document.getElementById("modal-preview-title").innerText = title;
        document.getElementById("modal-preview-body").innerHTML = contentHtml;

        const btnDocx = document.getElementById("modal-btn-docx");
        const btnXlsx = document.getElementById("modal-btn-xlsx");
        const btnPrint = document.getElementById("modal-btn-print");

        btnDocx.style.display = onDownloadDocx ? "inline-flex" : "none";
        btnXlsx.style.display = onDownloadXlsx ? "inline-flex" : "none";
        btnPrint.style.display = onPrint ? "inline-flex" : "none";

        btnDocx.onclick = onDownloadDocx || null;
        btnXlsx.onclick = onDownloadXlsx || null;
        btnPrint.onclick = onPrint || null;

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
            : [
                "Căn cứ Thông tư số 28/2020/TT-BGDĐT ngày 04/9/2020 của Bộ trưởng Bộ Giáo dục và Đào tạo ban hành Điều lệ Trường tiểu học;",
                "Căn cứ Thông tư số 32/2018/TT-BGDĐT ngày 26/12/2018 của Bộ trưởng Bộ Giáo dục và Đào tạo ban hành Chương trình Giáo dục phổ thông 2018;",
                `Căn cứ Quyết định ban hành Khung kế hoạch thời gian năm học ${state.settings.academicYear || '2026 - 2027'} của Ủy ban nhân dân Thành phố/Tỉnh;`,
                `Căn cứ Kế hoạch giáo dục nhà trường năm học ${state.settings.academicYear || '2026 - 2027'} của ${state.settings.schoolName || 'Trường Tiểu học Trung Nhứt'};`
            ];

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

                sub.rows.forEach(r => {
                    let rawInteg = r.integration || '';
                    rawInteg = rawInteg.replace(/System\.Xml\.XmlElement/g, '').trim();

                    let integFormatted = escapeHtml(rawInteg)
                        .replace(/(Lý tưởng cách mạng[^\:]*\:|Quyền con người[^\:]*\:|Tích hợp QCN[^\:]*\:|NLS[^\:]*\:|Năng lực số[^\:]*\:|Tích hợp NLS[^\:]*\:|AI[^\:]*\:|Tích hợp AI[^\:]*\:|BVMT[^\:]*\:|ANQP[^\:]*\:|QPAN[^\:]*\:|Tiết kiệm và bảo vệ nguồn nước[^\:]*\:|Đạo đức, lối sống[^\:]*\:)/g, '<br><strong>$1</strong>')
                        .replace(/^<br>/, '')
                        .replace(/\n/g, '<br>');

                    const contentFormatted = escapeHtml(r.content || '').replace(/\n/g, '<br>');
                    const lessonFormatted = escapeHtml(r.lesson || '').replace(/\n/g, '<br>');
                    const topicFormatted = escapeHtml(r.topic || r.theme || '').replace(/\n/g, '<br>');

                    paperHtml += `
                        <tr>
                            <td style="text-align:center; font-weight:bold; vertical-align:middle;">${escapeHtml(r.week || '')}</td>
                            <td style="font-weight:600; text-align:left; vertical-align:middle;">${topicFormatted}</td>
                            <td style="font-weight:600; text-align:left; vertical-align:middle;">${lessonFormatted}</td>
                            <td style="text-align:left; vertical-align:top;">${contentFormatted}</td>
                            <td style="text-align:center; vertical-align:middle;">${escapeHtml(r.duration || r.periods || '')}</td>
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

                sub.rows.forEach(r => {
                    let rawInteg = r.integration || '';
                    rawInteg = rawInteg.replace(/System\.Xml\.XmlElement/g, '').trim();

                    let integFormatted = escapeHtml(rawInteg)
                        .replace(/(Lý tưởng cách mạng[^\:]*\:|Quyền con người[^\:]*\:|Tích hợp QCN[^\:]*\:|NLS[^\:]*\:|Năng lực số[^\:]*\:|Tích hợp NLS[^\:]*\:|AI[^\:]*\:|Tích hợp AI[^\:]*\:|BVMT[^\:]*\:|ANQP[^\:]*\:|QPAN[^\:]*\:|Tiết kiệm và bảo vệ nguồn nước[^\:]*\:|Đạo đức, lối sống[^\:]*\:)/g, '<br><strong>$1</strong>')
                        .replace(/^<br>/, '')
                        .replace(/\n/g, '<br>');

                    const contentFormatted = escapeHtml(r.content || '').replace(/\n/g, '<br>');
                    const lessonFormatted = escapeHtml(r.lesson || '').replace(/\n/g, '<br>');
                    const topicFormatted = escapeHtml(r.topic || r.theme || '').replace(/\n/g, '<br>');

                    paperHtml += `
                        <tr>
                            <td style="text-align:center; font-weight:bold; vertical-align:middle;">${escapeHtml(r.week || '')}</td>
                            <td style="font-weight:600; text-align:left; vertical-align:middle;">${topicFormatted}</td>
                            <td style="font-weight:600; text-align:left; vertical-align:middle;">${lessonFormatted}</td>
                            <td style="text-align:left; vertical-align:top;">${contentFormatted}</td>
                            <td style="text-align:center; vertical-align:middle;">${escapeHtml(r.duration || '')}</td>
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

                sub.rows.forEach((r, rIdx) => {
                    let rawInteg = r.integration || '';
                    rawInteg = rawInteg.replace(/System\.Xml\.XmlElement/g, '').trim();

                    let integFormatted = escapeHtml(rawInteg)
                        .replace(/(Lý tưởng cách mạng[^\:]*\:|Quyền con người[^\:]*\:|Tích hợp QCN[^\:]*\:|NLS[^\:]*\:|Năng lực số[^\:]*\:|Tích hợp NLS[^\:]*\:|AI[^\:]*\:|Tích hợp AI[^\:]*\:|BVMT[^\:]*\:|ANQP[^\:]*\:|QPAN[^\:]*\:|Tiết kiệm và bảo vệ nguồn nước[^\:]*\:|Đạo đức, lối sống[^\:]*\:)/g, '<br><strong>$1</strong>')
                        .replace(/^<br>/, '')
                        .replace(/\n/g, '<br>');

                    const lessonFormatted = escapeHtml(r.lesson || '').replace(/\n/g, '<br>');
                    const topicFormatted = escapeHtml(r.topic || r.theme || '').replace(/\n/g, '<br>');
                    const khmhPeriodText = r.khmhPeriod || r.content || (rIdx + 1).toString();

                    paperHtml += `
                        <tr>
                            <td style="text-align:center; font-weight:bold; vertical-align:middle;">${escapeHtml(r.week || '')}</td>
                            <td style="font-weight:600; text-align:left; vertical-align:middle;">${topicFormatted}</td>
                            <td style="text-align:left; vertical-align:middle;">${lessonFormatted}</td>
                            <td style="text-align:center; vertical-align:middle;">${escapeHtml(r.duration || '')}</td>
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
                isCtlop: isCtlop
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
    function exportLbgToDocx(isCtlop = false) {
        if (window.DocxGenerator && window.DocxGenerator.generateLbgDocx) {
            showToast("Đang tạo file Word (.docx) chuẩn Nghị định 30...", "info");
            const { weekInfo, schedule, stats } = calculateWeekSchedule(state.currentWeek);
            window.DocxGenerator.generateLbgDocx(isCtlop, state.currentWeek, weekInfo, schedule, state.settings, stats).then(blob => {
                const prefix = isCtlop ? "Lich_Bao_Giang_Tich_Hop" : "Lich_Bao_Giang";
                const filename = `${prefix}_Tuan_${state.currentWeek}_${(state.settings.className || 'Lop_5A').replace(/\s+/g, '_')}.docx`;
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
    function renderSingleWeekPaperHtml(weekNum, isCtlop) {
        const { weekInfo, schedule, stats } = calculateWeekSchedule(weekNum);
        const bghSigner = getBghSignerInfo();
        const maxWidth = isCtlop ? "960px" : "900px";

        let html = `
        <div class="paper-page" style="max-width:${maxWidth}; margin-bottom: 2.5rem; page-break-after: always;">
            <div class="paper-header">
                <table class="paper-header-table">
                    <tr>
                        <td style="width:50%; text-align:center;">
                            <div style="font-size:12pt; text-transform:uppercase;">${state.settings.governingBody || 'UBND PHƯỜNG TRUNG NHỨT'}</div>
                            <div style="font-size:12pt; font-weight:bold; text-transform:uppercase;">${state.settings.schoolName || 'TRƯỜNG TIỂU HỌC TRUNG NHỨT'}</div>
                            <div style="font-size:12pt; font-weight:bold; margin-top:2px;">${state.settings.grade || 'KHỐI 5'} - ${state.settings.className || 'LỚP 5A'}</div>
                        </td>
                        <td style="width:50%; text-align:center;">
                            <div style="font-size:12pt; font-weight:bold;">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
                            <div style="font-size:13pt; font-weight:bold; text-decoration:underline;">Độc lập - Tự do - Hạnh phúc</div>
                        </td>
                    </tr>
                </table>
            </div>

            <div class="paper-title" style="font-size:15pt;">${isCtlop ? 'LỊCH BÁO GIẢNG TÍCH HỢP' : 'LỊCH BÁO GIẢNG'} TUẦN ${weekNum}</div>
            <div class="paper-subtitle">(Thời gian thực hiện: Từ ngày ${weekInfo.startDateVN} đến ngày ${weekInfo.endDateVN})</div>

            <table class="paper-table">
                <thead>
                    <tr>
                        <th style="width:${isCtlop ? '10%' : '12%'};">Thứ</th>
                        <th style="width:${isCtlop ? '7%' : '8%'};">Buổi</th>
                        <th style="width:${isCtlop ? '5%' : '6%'};">Tiết</th>
                        <th style="width:${isCtlop ? '17%' : '20%'};">Môn học</th>
                        <th style="width:${isCtlop ? '8%' : '10%'};">Tiết PPCT</th>
                        <th style="width:${isCtlop ? '30%' : '44%'};">Tên bài dạy</th>
                        ${isCtlop ? '<th style="width:23%;">Nội dung tích hợp / Điều chỉnh</th>' : ''}
                    </tr>
                </thead>
                <tbody>
        `;

        const days = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6"];
        days.forEach(day => {
            const daySlots = schedule.filter(s => s.day === day);
            const morningSlots = daySlots.filter(s => s.session === "Sáng");
            const afternoonSlots = daySlots.filter(s => s.session === "Chiều");
            daySlots.forEach((slot, idx) => {
                let dayCell = idx === 0 ? `<td rowspan="${daySlots.length}" style="text-align:center; font-weight:bold; vertical-align:middle;">${day}</td>` : "";
                let sessionCell = "";
                if (idx === 0) sessionCell = `<td rowspan="${morningSlots.length}" style="text-align:center; vertical-align:middle;">Sáng</td>`;
                else if (idx === morningSlots.length) sessionCell = `<td rowspan="${afternoonSlots.length}" style="text-align:center; vertical-align:middle;">Chiều</td>`;

                let integrationCell = "";
                if (isCtlop) {
                    const formattedInteg = escapeHtml(slot.integration || '')
                        .replace(/(Lý tưởng cách mạng[^\:]*\:|Quyền con người[^\:]*\:|Tích hợp QCN[^\:]*\:|NLS[^\:]*\:|Năng lực số[^\:]*\:|Tích hợp NLS[^\:]*\:|AI[^\:]*\:|Tích hợp AI[^\:]*\:|BVMT[^\:]*\:|ANQP[^\:]*\:|QPAN[^\:]*\:|Tiết kiệm và bảo vệ nguồn nước[^\:]*\:|Đạo đức, lối sống[^\:]*\:)/g, '<br><strong>$1</strong>')
                        .replace(/^<br>/, '')
                        .replace(/\n/g, '<br>');
                    integrationCell = `<td style="font-size:9.5pt; line-height:1.45; text-align:left; vertical-align:top;">${formattedInteg}</td>`;
                }

                html += `
                    <tr>
                        ${dayCell}
                        ${sessionCell}
                        <td style="text-align:center; font-weight:bold; vertical-align:middle;">${slot.period}</td>
                        <td style="font-weight:bold; text-align:left; vertical-align:middle;">${escapeHtml(slot.subject)}</td>
                        <td style="text-align:center; font-weight:bold; vertical-align:middle;">${escapeHtml(slot.ppct || '')}</td>
                        <td style="text-align:left; vertical-align:middle;">${escapeHtml(slot.lessonName || '')}</td>
                        ${integrationCell}
                    </tr>
                `;
            });
        });

        html += `
                </tbody>
            </table>

            <div class="paper-footer">
                <table class="paper-footer-table">
                    <tr>
                        <td style="width:33%; text-align:center;">
                            <div style="font-size:12pt; font-weight:bold;">DUYỆT CỦA ${bghSigner.title}</div>
                            <div style="font-size:11pt; font-style:italic;">(Ký và ghi rõ họ tên)</div>
                            <div style="height:65px;"></div>
                            <div style="font-size:12pt; font-weight:bold;">${bghSigner.name}</div>
                        </td>
                        <td style="width:33%; text-align:center;">
                            <div style="font-size:12pt; font-weight:bold;">KHỐI TRƯỞNG</div>
                            <div style="font-size:11pt; font-style:italic;">(Ký và ghi rõ họ tên)</div>
                            <div style="height:65px;"></div>
                            <div style="font-size:12pt; font-weight:bold;">${state.settings.headOfGrade || 'Trần Thị Mai'}</div>
                        </td>
                        <td style="width:34%; text-align:center;">
                            <div style="font-size:12pt; font-weight:bold;">GIÁO VIÊN CHỦ NHIỆM</div>
                            <div style="font-size:11pt; font-style:italic;">(Ký và ghi rõ họ tên)</div>
                            <div style="height:65px;"></div>
                            <div style="font-size:12pt; font-weight:bold;">${state.settings.homeroomTeacher || 'Nguyễn Thị Thu Hà'}</div>
                        </td>
                    </tr>
                </table>
            </div>
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

        document.getElementById("batch-export-modal").classList.add("show");
    }

    function closeBatchExportModal() {
        document.getElementById("batch-export-modal").classList.remove("show");
    }

    function exportBatchLbgToDocxDirect(isCtlop, startWeek, endWeek) {
        if (window.DocxGenerator && window.DocxGenerator.generateMultiWeekLbgDocx) {
            showToast(`Đang tạo file Word từ Tuần ${startWeek} đến Tuần ${endWeek}...`, "info");
            window.DocxGenerator.generateMultiWeekLbgDocx(isCtlop, startWeek, endWeek, calculateWeekSchedule, state.settings).then(blob => {
                const prefix = isCtlop ? "Lich_Bao_Giang_Tich_Hop" : "Lich_Bao_Giang";
                const filename = `${prefix}_Tuan_${startWeek}_den_${endWeek}_${(state.settings.className || 'Lop_5A').replace(/\s+/g, '_')}.docx`;
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

    function exportBatchLbgToDocx() {
        const isCtlop = document.querySelector('input[name="batch-doc-type"]:checked')?.value === 'ctlop';
        const startWeek = parseInt(document.getElementById("batch-start-week").value) || 1;
        const endWeek = parseInt(document.getElementById("batch-end-week").value) || 35;

        if (startWeek > endWeek) {
            alert("Tuần bắt đầu phải nhỏ hơn hoặc bằng tuần kết thúc!");
            return;
        }

        exportBatchLbgToDocxDirect(isCtlop, startWeek, endWeek);
        closeBatchExportModal();
    }

    function previewBatchLbg() {
        const isCtlop = document.querySelector('input[name="batch-doc-type"]:checked')?.value === 'ctlop';
        const startWeek = parseInt(document.getElementById("batch-start-week").value) || 1;
        const endWeek = parseInt(document.getElementById("batch-end-week").value) || 35;

        if (startWeek > endWeek) {
            alert("Tuần bắt đầu phải nhỏ hơn hoặc bằng tuần kết thúc!");
            return;
        }

        let allHtml = "";
        for (let w = startWeek; w <= endWeek; w++) {
            allHtml += renderSingleWeekPaperHtml(w, isCtlop);
        }

        const docTitle = isCtlop ? "Lịch Báo Giảng Tích Hợp" : "Lịch Báo Giảng";
        closeBatchExportModal();
        openPreviewModal(
            `Xem trước ${docTitle} từ Tuần ${startWeek} đến Tuần ${endWeek}`,
            allHtml,
            () => exportBatchLbgToDocxDirect(isCtlop, startWeek, endWeek),
            null,
            () => window.print()
        );
    }

    // 7. App Initialization & Event Bindings
    function initApp() {
        loadState();

        document.querySelectorAll(".tab-btn").forEach(btn => {
            btn.addEventListener("click", () => {
                document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
                document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));

                btn.classList.add("active");
                const targetTab = btn.dataset.tab;
                state.currentTab = targetTab;
                document.getElementById(targetTab).classList.add("active");

                if (targetTab === "tab-lbg") renderTabLbg();
                else if (targetTab === "tab-ctlop") renderTabCtlop();
                else if (targetTab === "tab-lichtuan") renderTabLichtuan();
                else if (targetTab === "tab-ppct") renderTabPpct();
                else if (targetTab === "tab-settings") renderTabSettings();
            });
        });

        // Tab 1 Actions
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
            openPreviewModal(
                `Xem trước Lịch Báo Giảng Tuần ${state.currentWeek}`,
                renderSingleWeekPaperHtml(state.currentWeek, false),
                () => exportLbgToDocx(false),
                () => exportLbgToExcel(false),
                () => window.print()
            );
        });

        document.getElementById("btn-lbg-excel").addEventListener("click", () => exportLbgToExcel(false));
        document.getElementById("btn-lbg-print").addEventListener("click", () => {
            openPreviewModal(
                `Xem trước Lịch Báo Giảng Tuần ${state.currentWeek}`,
                renderSingleWeekPaperHtml(state.currentWeek, false),
                () => exportLbgToDocx(false),
                () => exportLbgToExcel(false),
                () => window.print()
            );
            setTimeout(() => {
                window.print();
            }, 250);
        });

        // Tab 2 Actions (CTLOP)
        const btnCtlopBatch = document.getElementById("btn-ctlop-batch");
        if (btnCtlopBatch) {
            btnCtlopBatch.addEventListener("click", () => openBatchExportModal('ctlop'));
        }

        document.getElementById("btn-ctlop-preview").addEventListener("click", () => {
            openPreviewModal(
                `Xem trước Lịch Báo Giảng Tích Hợp Tuần ${state.currentWeek}`,
                renderSingleWeekPaperHtml(state.currentWeek, true),
                () => exportLbgToDocx(true),
                () => exportLbgToExcel(true),
                () => window.print()
            );
        });

        document.getElementById("btn-ctlop-excel").addEventListener("click", () => exportLbgToExcel(true));
        document.getElementById("btn-ctlop-print").addEventListener("click", () => {
            openPreviewModal(
                `Xem trước Lịch Báo Giảng Tích Hợp Tuần ${state.currentWeek}`,
                renderSingleWeekPaperHtml(state.currentWeek, true),
                () => exportLbgToDocx(true),
                () => exportLbgToExcel(true),
                () => window.print()
            );
            setTimeout(() => {
                window.print();
            }, 250);
        });

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

        document.getElementById("btn-ctlop-excel").addEventListener("click", () => exportLbgToExcel(true));
        document.getElementById("btn-ctlop-print").addEventListener("click", () => window.print());

        // Tab 3 Smart Generator Actions
        document.getElementById("btn-run-week-generator").addEventListener("click", () => {
            const yr = document.getElementById("gen-year").value.trim() || "2026 - 2027";
            const sDate = document.getElementById("gen-start-date").value;
            const tetOpt = parseInt(document.getElementById("gen-tet-opt").value);
            autoGenerateCalendar(yr, sDate, tetOpt);
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
                    state.settings.bases = [
                        "Căn cứ Thông tư số 28/2020/TT-BGDĐT ngày 04/9/2020 của Bộ trưởng Bộ Giáo dục và Đào tạo ban hành Điều lệ Trường tiểu học;",
                        "Căn cứ Thông tư số 32/2018/TT-BGDĐT ngày 26/12/2018 của Bộ trưởng Bộ Giáo dục và Đào tạo ban hành Chương trình Giáo dục phổ thông 2018;",
                        `Căn cứ Quyết định ban hành Khung kế hoạch thời gian năm học ${state.settings.academicYear || '2026 - 2027'} của Ủy ban nhân dân Thành phố/Tỉnh;`,
                        `Căn cứ Kế hoạch giáo dục nhà trường năm học ${state.settings.academicYear || '2026 - 2027'} của ${state.settings.schoolName || 'Trường Tiểu học Trung Nhứt'};`
                    ];
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

        // PPCT Excel Upload, Template Download, and Reset
        document.getElementById("btn-download-ppct-template").addEventListener("click", downloadPpctTemplate);
        
        const fileInput = document.getElementById("ppct-upload-file");
        document.getElementById("btn-upload-ppct").addEventListener("click", () => {
            fileInput.click();
        });

        fileInput.addEventListener("change", (e) => {
            const file = e.target.files[0];
            if (file) {
                uploadPpctFromExcel(file);
                fileInput.value = "";
            }
        });

        document.getElementById("btn-reset-ppct").addEventListener("click", () => {
            if (confirm("Bạn có chắc chắn muốn khôi phục lại phân phối chương trình gốc chuẩn?")) {
                if (window.APP_INITIAL_DATA && window.APP_INITIAL_DATA.ppct) {
                    state.ppct = JSON.parse(JSON.stringify(window.APP_INITIAL_DATA.ppct));
                    saveState();
                    renderTabPpct();
                    renderTabLbg();
                    renderTabCtlop();
                    showToast("Đã khôi phục dữ liệu phân phối chương trình chuẩn gốc!", "success");
                }
            }
        });

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
                if (confirm("Bạn có chắc chắn muốn khôi phục lại Thời khóa biểu gốc chuẩn 35 tiết mặc định không?")) {
                    if (window.APP_INITIAL_DATA && window.APP_INITIAL_DATA.settings && window.APP_INITIAL_DATA.settings.timetable) {
                        state.timetable = JSON.parse(JSON.stringify(window.APP_INITIAL_DATA.settings.timetable));
                        state.weeklyCustomSlots = {};
                        saveState();
                        renderMasterTimetableEditor();
                        renderTabLbg();
                        renderTabCtlop();
                        showToast("Đã khôi phục Thời khóa biểu gốc chuẩn 35 tiết!", "success");
                    }
                }
            });
        }

        document.getElementById("btn-save-settings").addEventListener("click", saveSettingsFromForm);
        document.getElementById("btn-backup-data").addEventListener("click", () => {
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(localStorage.getItem(STORAGE_KEY) || "{}");
            const dlAnchor = document.createElement('a');
            dlAnchor.setAttribute("href", dataStr);
            dlAnchor.setAttribute("download", `Sao_Luu_LBG_Lop_5A_${new Date().toISOString().slice(0,10)}.json`);
            dlAnchor.click();
            dlAnchor.remove();
            showToast("Đã xuất file sao lưu dữ liệu!", "success");
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
                        localStorage.setItem(STORAGE_KEY, content);
                        loadState();
                        renderTabLbg();
                        renderTabSettings();
                        renderTabLichtuan();
                        renderTabPpct();
                        showToast("Khôi phục dữ liệu thành công!", "success");
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
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", initApp);
    } else {
        initApp();
    }
})();
