// CSV Parser
class CSVParser {
    static parse(text) {
        const lines = text.split('\n').filter((line) => line.trim());
        if (lines.length === 0) return [];

        const headers = this.parseCSVLine(lines[0]);
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = this.parseCSVLine(lines[i]);
            const row = {};
            headers.forEach((header, index) => {
                row[header.trim()] = values[index] ? values[index].trim() : '';
            });
            data.push(row);
        }

        return data;
    }

    static parseCSVLine(line) {
        const result = [];
        let current = '';
        let insideQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            const nextChar = line[i + 1];

            if (char === '"') {
                if (insideQuotes && nextChar === '"') {
                    current += '"';
                    i++;
                } else {
                    insideQuotes = !insideQuotes;
                }
            } else if (char === ',' && !insideQuotes) {
                result.push(current);
                current = '';
            } else {
                current += char;
            }
        }

        result.push(current);
        return result;
    }
}

class SimpleXLSXParser {
    static libraryPromise = null;

    static loadScript(src) {
        return new Promise((resolve, reject) => {
            const existing = document.querySelector(`script[data-xlsx-src="${src}"]`);
            if (existing) {
                if (typeof XLSX !== 'undefined') {
                    resolve();
                } else {
                    existing.addEventListener('load', () => resolve(), { once: true });
                    existing.addEventListener('error', () => reject(new Error('Không thể tải thư viện đọc Excel.')), { once: true });
                }
                return;
            }

            const script = document.createElement('script');
            script.src = src;
            script.async = true;
            script.dataset.xlsxSrc = src;
            script.onload = () => resolve();
            script.onerror = () => reject(new Error(`Không thể tải thư viện từ ${src}`));
            document.head.appendChild(script);
        });
    }

    static async ensureLibraryLoaded() {
        if (typeof XLSX !== 'undefined') return;

        if (!this.libraryPromise) {
            this.libraryPromise = (async () => {
                const sources = [
                    'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.min.js',
                    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
                ];

                let lastError = null;
                for (const source of sources) {
                    try {
                        await this.loadScript(source);
                        if (typeof XLSX !== 'undefined') return;
                    } catch (error) {
                        lastError = error;
                    }
                }

                throw new Error(lastError ? lastError.message : 'Không thể tải thư viện đọc Excel.');
            })();
        }

        await this.libraryPromise;
    }

    static async parse(arrayBuffer) {
        await this.ensureLibraryLoaded();

        try {
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            return XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        } catch (e) {
            throw new Error('Không thể đọc file Excel. Vui lòng lưu file dưới dạng CSV và tải lên lại.');
        }
    }
}

class GradeManager {
    constructor(options = {}) {
        this.studentInfo = [];
        this.gradeData = [];
        this.mergedData = [];
        this.studentMap = new Map();
        this.recordsByClass = new Map();
        this.recordsBySubject = new Map();
        this.subjects = [];
        this.classes = [];

        this.reportSubjectCodes = ['190036', '190008', '195001'];
        this.persistEnabled = options.persistEnabled !== false;
        this.subjectMeta = {
            '195001': {
                title: 'Giáo dục quốc phòng và an ninh (195001)',
                detail: '(Trình độ: Liên thông Cao đẳng, Thời lượng: 30 tiết)'
            },
            '190008': {
                title: 'Giáo dục quốc phòng và an ninh (190008)',
                detail: '(Trình độ: Trung cấp, Thời lượng: 45 tiết)'
            },
            '190036': {
                title: 'Giáo dục quốc phòng và an ninh (190036)',
                detail: '(Trình độ: Cao đẳng, Thời lượng: 75 tiết)'
            }
        };

        this.schoolCodeMap = {
            DVT: 'Trường Đại học Trà Vinh',
            CDD5704: 'Trường Cao đẳng Vĩnh Long',
            CDD5802: 'Trường Cao đẳng Y tế Trà Vinh',
            CDD5302: 'Trường Cao đẳng Y tế Tiền Giang',
            CDD5701: 'Trường Cao đẳng Nghề Vĩnh Long',
            CDD5801: 'Trường Cao đẳng Nghề Trà Vinh',
            CDD5301: 'Trường Cao đẳng Tiền Giang',
            CDT5301: 'Trường Cao đẳng Nông nghiệp Nam Bộ',
            KSV: 'Phân hiệu Trường Đại học Kinh tế TP. HCM tại Vĩnh Long',
            MTU: 'Trường ĐH XD Miền Tây',
            CDT5601: 'Trường Cao đẳng Đồng Khởi Bến Tre.',
            C56: 'Trường Cao đẳng Bến Tre.'
        };

        this.loadFromStorage();
    }

    normalizeSubjectCode(code) {
        const text = String(code || '').trim();
        if (text === '1190036') return '190036';
        return text;
    }

    static normalizeFieldName(name) {
        return String(name || '')
            .trim()
            .toUpperCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/[^A-Z0-9]/g, '');
    }

    static normalizeRow(row) {
        const normalized = {};
        Object.keys(row || {}).forEach((key) => {
            normalized[this.normalizeFieldName(key)] = row[key];
        });
        return normalized;
    }

    static pickValue(normalizedRow, keys, fallback = '') {
        for (const key of keys) {
            const value = normalizedRow[this.normalizeFieldName(key)];
            if (value !== undefined && value !== null && String(value).trim() !== '') {
                return String(value).trim();
            }
        }
        return fallback;
    }

    static pickClassName(normalizedRow) {
        const direct = this.pickValue(normalizedRow, [
            'TenLop', 'Ten lop', 'Tên lớp', 'TenNhomLop', 'TenLopHoc',
            'LopHoc', 'ClassName', 'Class name', 'TenNganhLop'
        ], '');
        if (direct && !/^\d+(?:[.,]\d+)?$/.test(direct)) return direct;

        const keys = Object.keys(normalizedRow || {});
        for (const key of keys) {
            const value = normalizedRow[key];
            const text = String(value == null ? '' : value).trim();
            if (!text) continue;
            if (/^\d+(?:[.,]\d+)?$/.test(text)) continue;
            if (/^CA\d/i.test(text) || /^(CDD|CDT)\d{4}$/i.test(text)) continue;
            if (key.includes('TENLOP') || key.includes('LOPHOC') || key.includes('TENNHOM')) {
                return text;
            }
        }
        return '';
    }

    loadFromStorage() {
        if (!this.persistEnabled) return;
        const stored = localStorage.getItem('gradeData');
        if (!stored) return;

        try {
            const data = JSON.parse(stored);
            this.studentInfo = data.studentInfo || [];
            this.gradeData = data.gradeData || [];
            this.subjects = [...new Set(this.gradeData.map((g) => g.MaMH).filter(Boolean))].sort();
            this.classes = this.collectClasses();
            this.mergeData();
            // Regenerate KetQua/XepLoai for all merged rows to ensure correct categorization
            this.regenerateResultColumns();
        } catch (e) {
            console.error('Error loading from storage:', e);
        }
    }

    regenerateResultColumns() {
        this.mergedData.forEach((row) => {
            const rawScore = row.DiemTB != null && String(row.DiemTB).trim() !== ''
                ? row.DiemTB
                : row.T1_DTK;
            const specialScoreCode = this.getSpecialScoreCode(rawScore);
            const numericScore = Number(String(rawScore == null ? '' : rawScore).replace(',', '.'));
            const safeScore = Number.isFinite(numericScore) ? numericScore : 0;
            
            row.XepLoai = specialScoreCode ? 'Không đạt' : this.getClassification(safeScore);
            row.KetQua = specialScoreCode || this.getResult(safeScore);
        });
        
        // Rebuild recordsBySubject with updated KetQua
        this.recordsBySubject.clear();
        this.mergedData.forEach((row) => {
            if (row.MaMH) {
                const normalizedCode = this.normalizeSubjectCode(row.MaMH);
                if (!this.recordsBySubject.has(normalizedCode)) {
                    this.recordsBySubject.set(normalizedCode, []);
                }
                this.recordsBySubject.get(normalizedCode).push(row);
            }
        });
        
        this.saveToStorage();
    }

    saveToStorage() {
        if (!this.persistEnabled) return;
        const data = {
            studentInfo: this.studentInfo,
            gradeData: this.gradeData,
            version: 2
        };
        try {
            localStorage.setItem('gradeData', JSON.stringify(data));
        } catch (e) {
            console.error('Error saving to storage:', e);
        }
    }

    importStudentInfo(data) {
        this.studentInfo = data.map((row) => {
            const normalized = GradeManager.normalizeRow(row);
            return {
                MaSV: GradeManager.pickValue(normalized, ['MaSV', 'MSSV']),
                HoLotSV: GradeManager.pickValue(normalized, ['HoLotSV', 'HoVa']),
                TenSV: GradeManager.pickValue(normalized, ['TenSV', 'Ten']),
                NgaySinhC: GradeManager.pickValue(normalized, ['NgaySinhC', 'NgaySinh']),
                NoiSinh: GradeManager.pickValue(normalized, ['NoiSinh']),
                TenNganh: GradeManager.pickValue(normalized, ['TenNganh', 'NganhHoc', 'TenChNg', 'Nganh', 'ChuyenNganh']),
                NganhHoc: GradeManager.pickValue(normalized, ['NganhHoc', 'TenNganh', 'Nganh', 'ChuyenNganh']),
                MaLop: GradeManager.pickValue(normalized, ['MaLop', 'NhomLop']),
                TenLop: GradeManager.pickClassName(normalized),
                TenDT: GradeManager.pickValue(normalized, ['TenDT']),
                Phai: GradeManager.pickValue(normalized, ['Phai', 'GioiTinh'])
            };
        });

        this.classes = this.collectClasses();
        this.mergeData();
        this.saveToStorage();
    }

    importGradeData(data) {
        this.gradeData = data.map((row) => {
            const normalized = GradeManager.normalizeRow(row);
            const maMHRaw = GradeManager.pickValue(normalized, ['MaMH', 'MonHoc']);
            return {
                MaSV: GradeManager.pickValue(normalized, ['MaSV', 'MSSV']),
                HoLotSV: GradeManager.pickValue(normalized, ['HoLotSV', 'HoVa']),
                TenSV: GradeManager.pickValue(normalized, ['TenSV', 'Ten']),
                NgaySinhC: GradeManager.pickValue(normalized, ['NgaySinhC', 'NgaySinh']),
                MaLop: GradeManager.pickValue(normalized, ['MaLop', 'NhomLop']),
                TenLop: GradeManager.pickClassName(normalized),
                TenNganh: GradeManager.pickValue(normalized, ['TenNganh', 'NganhHoc', 'TenChNg', 'Nganh', 'ChuyenNganh']),
                NganhHoc: GradeManager.pickValue(normalized, ['NganhHoc', 'TenNganh', 'TenChNg', 'Nganh', 'ChuyenNganh']),
                NhomHoc: GradeManager.pickValue(normalized, ['NhomHoc', 'NhomLop']),
                ToTH: GradeManager.pickValue(normalized, ['ToTH']),
                MaMH: this.normalizeSubjectCode(maMHRaw),
                QT: parseFloat(GradeManager.pickValue(normalized, ['QT', 'DiemQT']) || 0) || 0,
                Thi: parseFloat(GradeManager.pickValue(normalized, ['Thi', 'DiemThi']) || 0) || 0,
                T1_DTK: parseFloat(GradeManager.pickValue(normalized, ['T1_DTK', 'T1ĐTK', 'T1DTK', 'DiemTB']) || 0) || 0,
                L1: parseFloat(GradeManager.pickValue(normalized, ['L1', 'T1_DTK_L1', 'T1DTKL1']) || 0) || 0,
                T2_DTK: parseFloat(GradeManager.pickValue(normalized, ['T2_DTK', 'T2ĐTK', 'T2DTK']) || 0) || 0,
                L2: parseFloat(GradeManager.pickValue(normalized, ['L2']) || 0) || 0,
                T3_T3: parseFloat(GradeManager.pickValue(normalized, ['T3_T3', 'T3T3']) || 0) || 0,
                DiemTB: GradeManager.pickValue(normalized, ['DiemTB', 'T1_ĐTK L1', 'T1_DTK L1', 'T1ĐTK L1', 'T1DTKL1'])
            };
        });

        this.subjects = [...new Set(this.gradeData.map((g) => g.MaMH).filter(Boolean))].sort();
        this.classes = this.collectClasses();
        this.mergeData();
        this.saveToStorage();
    }

    collectClasses() {
        return [
            ...new Set([
                ...this.studentInfo.map((s) => s.MaLop),
                ...this.gradeData.map((g) => g.MaLop)
            ].filter(Boolean))
        ].sort();
    }

    mergeData() {
        this.studentMap = new Map(this.studentInfo.map((student) => [student.MaSV, student]));
        this.recordsByClass = new Map();
        this.recordsBySubject = new Map();

        this.mergedData = this.gradeData.map((grade) => {
            const student = this.studentMap.get(grade.MaSV);
            const mergedRow = {
                ...grade,
                ...(student || {}),
                MaSV: grade.MaSV || (student ? student.MaSV : ''),
                HoLotSV: grade.HoLotSV || (student ? student.HoLotSV : ''),
                TenSV: grade.TenSV || (student ? student.TenSV : ''),
                NgaySinhC: grade.NgaySinhC || (student ? student.NgaySinhC : ''),
                MaLop: grade.MaLop || (student ? student.MaLop : ''),
                TenLop: grade.TenLop
                    || (student ? student.TenLop : '')
                    || grade.TenNganh
                    || grade.NganhHoc
                    || (student ? (student.TenNganh || student.NganhHoc || '') : ''),
                TenNganh: grade.TenNganh || grade.NganhHoc || (student ? (student.TenNganh || student.NganhHoc || '') : ''),
                NganhHoc: grade.NganhHoc || grade.TenNganh || (student ? (student.NganhHoc || student.TenNganh || '') : '')
            };

            if (mergedRow.MaLop) {
                if (!this.recordsByClass.has(mergedRow.MaLop)) {
                    this.recordsByClass.set(mergedRow.MaLop, []);
                }
                this.recordsByClass.get(mergedRow.MaLop).push(mergedRow);
            }

            if (mergedRow.MaMH) {
                if (!this.recordsBySubject.has(mergedRow.MaMH)) {
                    this.recordsBySubject.set(mergedRow.MaMH, []);
                }
                this.recordsBySubject.get(mergedRow.MaMH).push(mergedRow);
            }

            return mergedRow;
        });
    }

    getDataByClass(className) {
        if (!className) return this.mergedData;
        return this.recordsByClass.get(className) || [];
    }

    getDataBySubject(subjectCode) {
        if (!subjectCode) return this.mergedData;
        return this.recordsBySubject.get(subjectCode) || [];
    }

    getStatistics() {
        const stats = {
            totalStudents: [...new Set(this.mergedData.map((d) => d.MaSV))].length,
            totalSubjects: this.subjects.length,
            totalClasses: this.classes.length,
            averageScore: 0,
            subjectStats: {}
        };

        if (this.mergedData.length > 0) {
            const scores = this.mergedData.map((d) => Number(d.Thi)).filter((s) => !Number.isNaN(s) && s > 0);
            stats.averageScore = scores.length > 0 ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2) : 0;
        }

        this.subjects.forEach((subject) => {
            const subjectData = this.getDataBySubject(subject);
            const scores = subjectData.map((d) => Number(d.Thi)).filter((s) => !Number.isNaN(s) && s > 0);
            stats.subjectStats[subject] = {
                count: subjectData.length,
                average: scores.length > 0 ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2) : 0,
                max: scores.length > 0 ? Math.max(...scores) : 0,
                min: scores.length > 0 ? Math.min(...scores) : 0
            };
        });

        return stats;
    }

    getReportSubjectOptions() {
        return [...this.reportSubjectCodes];
    }

    getSubjectDisplayText(subjectCode) {
        const code = this.normalizeSubjectCode(subjectCode);
        const meta = this.subjectMeta[code];
        if (!meta) {
            return `Giáo dục quốc phòng và an ninh (${code || ''})`;
        }
        return `${meta.title} ${meta.detail}`;
    }

    extractSchoolCodeFromClass(classCode) {
        const text = String(classCode || '').trim();
        if (!text) return '';
        const normalized = text.toUpperCase();
        const match = normalized.match(/(CDD\d{4}|CDT\d{4}|DVT|KSV|MTU|C56)\b/);
        if (match && match[1]) {
            return match[1];
        }

        const parts = normalized.split('_').filter(Boolean);
        return (parts.length ? parts[parts.length - 1] : normalized).trim();
    }

    resolveSchoolName(row = {}) {
        const schoolCode = this.extractSchoolCodeFromClass(row.MaLop);
        if (schoolCode && this.schoolCodeMap[schoolCode]) {
            return this.schoolCodeMap[schoolCode];
        }

        const tenDT = String(row.TenDT || '').trim();
        if (tenDT) {
            if (/trường/i.test(tenDT)) return tenDT;
            const mappedByTenDT = this.schoolCodeMap[tenDT.toUpperCase()];
            if (mappedByTenDT) return mappedByTenDT;
        }

        return '';
    }

    getDominantSchoolName(rows = []) {
        const counts = new Map();
        rows.forEach((row) => {
            const school = String(row.SchoolName || row.Truong || '').trim();
            if (!school) return;
            counts.set(school, (counts.get(school) || 0) + 1);
        });

        let dominant = '';
        let maxCount = 0;
        counts.forEach((count, school) => {
            if (count > maxCount) {
                dominant = school;
                maxCount = count;
            }
        });
        return dominant;
    }

    getDisplayGroupName(rows = []) {
        const candidates = rows
            .map((row) => String(row.TenLop || row.TenNhomLop || row.TenNganh || row.NganhHoc || row.NhomLop || row.NhomHoc || row.ToTH || '').trim())
            .filter((value) => {
                if (!value) return false;
                // Numeric-only labels (e.g. "15") are group numbers, not class names.
                return !/^\d+(?:[.,]\d+)?$/.test(value);
            });
        if (!candidates.length) return '';
        return [...new Set(candidates)][0];
    }

    formatClassDisplay(classCode, groupName) {
        const code = String(classCode || '').split(',')[0].trim();
        const name = String(groupName || '').split(',')[0].trim();
        if (code && name) return `${code}(${name})`;
        if (code) return code;
        return name;
    }

    getSchoolDisplayText(classCode, fallbackName = '') {
        const schoolCode = this.extractSchoolCodeFromClass(classCode);
        const mappedName = schoolCode ? (this.schoolCodeMap[schoolCode] || '') : '';
        const name = mappedName || String(fallbackName || '').trim();
        return name || '';
    }

    getSchoolOptions(rows = this.mergedData) {
        const options = new Set();
        Object.values(this.schoolCodeMap || {}).forEach((name) => {
            if (name) options.add(name);
        });
        rows.forEach((row) => {
            const school = String(row.SchoolName || row.Truong || '').trim();
            if (school) options.add(school);
            const tenDT = String(row.TenDT || '').trim();
            if (tenDT && /trường/i.test(tenDT)) options.add(tenDT);
        });
        return [...options].sort((a, b) => a.localeCompare(b));
    }

    getReportClassOptions(subjectCode = '') {
        const normalizedSubjectCode = this.normalizeSubjectCode(subjectCode);
        const source = normalizedSubjectCode
            ? (this.recordsBySubject.get(normalizedSubjectCode) || [])
            : this.gradeData;

        return [...new Set(source.map((row) => row.MaLop).filter(Boolean))].sort();
    }

    getSpecialScoreCode(scoreValue) {
        const text = String(scoreValue == null ? '' : scoreValue).trim().toUpperCase();
        if (!text) return '';
        if (text === 'CT' || text === 'CD' || text === 'VT') return text;
        return '';
    }

    getClassification(score) {
        if (!Number.isFinite(score) || score < 5) return 'Không đạt';
        if (score >= 9) return 'Xuất sắc';
        if (score >= 8) return 'Giỏi';
        if (score >= 6.5) return 'Khá';
        return 'Trung bình';
    }

    getResult(score) {
        return Number.isFinite(score) && score >= 5 ? 'Đạt' : 'Hỏng';
    }

    calculateStatistics(rows) {
        const stats = {
            total: rows.length,
            passed: 0,
            banned: 0,
            failed: 0,
            absent: 0,
            suspended: 0,
            notStudied: 0
        };

        rows.forEach((row) => {
            const ketQua = String(row.KetQua || '').trim();
            const diemTB = row.DiemTB;
            const diemStr = String(diemTB || '').trim();
            const diemNumeric = diemStr ? Number(diemStr.replace(',', '.')) : NaN;
            
            if (ketQua === 'Đạt') {
                stats.passed++;
            } else if (ketQua === 'CT' || ketQua === 'CD') {
                stats.banned++;
            } else if (ketQua === 'VT') {
                stats.absent++;
            } else if (ketQua === 'Đình chỉ') {
                stats.suspended++;
            } else if (diemNumeric === 0 || (!diemStr || Number.isNaN(diemNumeric))) {
                stats.notStudied++;
            } else if (diemNumeric > 0 && diemNumeric < 5.0) {
                stats.failed++;
            } else {
                stats.failed++;
            }
        });

        return stats;
    }

    getScoreSheetData(subjectCode, classCode) {
        const normalizedSubjectCode = this.normalizeSubjectCode(subjectCode);
        const gradeRows = normalizedSubjectCode
            ? (this.recordsBySubject.get(normalizedSubjectCode) || [])
            : this.mergedData;

        const rowsByStudent = new Map();
        gradeRows.forEach((gradeRow) => {
            if (classCode && gradeRow.MaLop !== classCode) return;

            const student = this.studentMap.get(gradeRow.MaSV) || {};
            const hoLot = gradeRow.HoLotSV || student.HoLotSV || '';
            const ten = gradeRow.TenSV || student.TenSV || '';
            const rawScore = gradeRow.DiemTB != null && String(gradeRow.DiemTB).trim() !== ''
                ? gradeRow.DiemTB
                : gradeRow.T1_DTK;
            const specialScoreCode = this.getSpecialScoreCode(rawScore);
            const numericScore = Number(String(rawScore == null ? '' : rawScore).replace(',', '.'));
            const safeScore = Number.isFinite(numericScore) ? numericScore : 0;
            const scoreForDisplay = specialScoreCode || safeScore;
            const scoreRankValue = specialScoreCode ? -1 : safeScore;
            const schoolName = this.resolveSchoolName({
                ...student,
                ...gradeRow,
                MaLop: gradeRow.MaLop || student.MaLop || '',
                TenDT: gradeRow.TenDT || student.TenDT || ''
            });

            const row = {
                MaSV: gradeRow.MaSV || student.MaSV || '',
                HoVa: hoLot,
                Ten: ten,
                HoTen: [hoLot, ten].join(' ').replace(/\s+/g, ' ').trim(),
                NgaySinhC: gradeRow.NgaySinhC || student.NgaySinhC || '',
                DiemTB: scoreForDisplay,
                XepLoai: specialScoreCode ? 'Không đạt' : this.getClassification(safeScore),
                KetQua: specialScoreCode || this.getResult(safeScore),
                _scoreRank: scoreRankValue,
                MaLop: gradeRow.MaLop || student.MaLop || '',
                MaMH: this.normalizeSubjectCode(gradeRow.MaMH || normalizedSubjectCode || ''),
                TenNganh: gradeRow.TenNganh || gradeRow.NganhHoc || student.TenNganh || student.NganhHoc || '',
                NganhHoc: gradeRow.NganhHoc || gradeRow.TenNganh || student.NganhHoc || student.TenNganh || '',
                NhomHoc: gradeRow.NhomHoc || gradeRow.ToTH || '',
                SchoolName: schoolName,
                Truong: schoolName
            };

            const existing = rowsByStudent.get(row.MaSV);
            if (!existing || row._scoreRank > existing._scoreRank) {
                rowsByStudent.set(row.MaSV, row);
            }
        });

        const rows = Array.from(rowsByStudent.values()).sort((a, b) => a.MaSV.localeCompare(b.MaSV));
        const total = rows.length;
        const passed = rows.filter((row) => row.KetQua === 'Đạt').length;

        return {
            rows,
            total,
            passed,
            failed: total - passed,
            majorName: rows.find((row) => row.TenNganh)?.TenNganh || '',
            schoolName: this.getDominantSchoolName(rows),
            groupName: this.getDisplayGroupName(rows),
            subjectCode: normalizedSubjectCode || (rows[0] ? rows[0].MaMH : ''),
            classCode: classCode || (rows[0] ? rows[0].MaLop : '')
        };
    }

    clearAll() {
        this.studentInfo = [];
        this.gradeData = [];
        this.mergedData = [];
        this.studentMap = new Map();
        this.recordsByClass = new Map();
        this.recordsBySubject = new Map();
        this.subjects = [];
        this.classes = [];
        localStorage.removeItem('gradeData');
    }
}

class UIManager {
    constructor(manager) {
        this.manager = manager;
        this.selLop = new Set();
        this.selectedStudentCodes = new Set();
        this.allKeys = [];
        this._searchHandler = null;
        this.initEventListeners();
        this.render();
    }

    initEventListeners() {
        const bind = (id, eventName, handler) => {
            const el = document.getElementById(id);
            if (el) {
                el.addEventListener(eventName, handler);
            }
        };

        document.querySelectorAll('.tab-btn').forEach((btn) => {
            btn.addEventListener('click', (e) => this.switchTab(e.target.dataset.tab));
        });

        bind('studentFile', 'change', (e) => this.handleStudentFileUpload(e));
        bind('gradeFile', 'change', (e) => this.handleGradeFileUpload(e));
        bind('clearData', 'click', () => this.clearAllData());

        bind('classFilter', 'change', () => this.renderByClass());
        bind('subjectFilter', 'change', () => this.renderBySubject());
        bind('reportSubjectFilter', 'change', () => {
            this.updateFilters();
            this.renderScoreSheet();
        });
        bind('reportClassFilter', 'change', () => this.renderScoreSheet());
        bind('reportSchoolFilter', 'change', () => {
            document.getElementById('reportClassFilter').value = '';
            this.renderScoreSheet();
        });
        bind('selectAllClassBtn', 'click', () => this.selectAll());
        bind('clearAllClassBtn', 'click', () => this.clearAll());
        bind('studentSearchToggleBtn', 'click', () => this.toggleStudentSearchPanel());
        bind('applyStudentCodesBtn', 'click', () => this.applyStudentCodesFilter());
        bind('clearStudentCodesBtn', 'click', () => this.clearStudentCodesFilter());
        bind('exportScoreSheetBtn', 'click', () => this.exportScoreSheetExcel());
        bind('exportDecisionBtn', 'click', () => this.exportDecisionExcel());

        this.initSearch();

        document.addEventListener('click', (event) => {
            const input = document.getElementById('searchInput');
            const listEl = document.getElementById('classSearchList');
            if (!input || !listEl) return;
            if (event.target === input || listEl.contains(event.target)) return;
            listEl.classList.remove('show');

            const panel = document.getElementById('studentSearchPanel');
            const toggleBtn = document.getElementById('studentSearchToggleBtn');
            if (!panel || !toggleBtn) return;
            if (panel.contains(event.target) || toggleBtn.contains(event.target)) return;
            panel.classList.remove('show');
        });
    }

    parseStudentCodes(text) {
        const parts = String(text || '').split(/[\s,;]+/);
        const unique = new Set();
        parts.forEach((part) => {
            const code = String(part || '').trim();
            if (!code) return;
            unique.add(code);
        });
        return [...unique];
    }

    updateStudentCodeFilterInfo() {
        const infoEl = document.getElementById('studentCodeFilterInfo');
        if (!infoEl) return;
        if (!this.selectedStudentCodes.size) {
            infoEl.textContent = 'Chưa áp dụng lọc MSSV.';
            return;
        }
        infoEl.textContent = `Đang lọc theo ${this.selectedStudentCodes.size} MSSV.`;
    }

    toggleStudentSearchPanel() {
        const panel = document.getElementById('studentSearchPanel');
        const input = document.getElementById('studentCodesInput');
        if (!panel) return;

        panel.classList.toggle('show');
        if (panel.classList.contains('show') && input) {
            input.focus();
        }
        this.updateStudentCodeFilterInfo();
    }

    applyStudentCodesFilter() {
        const input = document.getElementById('studentCodesInput');
        if (!input) return;

        const codes = this.parseStudentCodes(input.value);
        this.selectedStudentCodes = new Set(codes);
        this.updateStudentCodeFilterInfo();
        this.renderStudentCodeTags();
        this.renderTabs();
        this.renderQDTabs();
    }

    clearStudentCodesFilter() {
        this.selectedStudentCodes = new Set();
        const input = document.getElementById('studentCodesInput');
        if (input) input.value = '';
        this.updateStudentCodeFilterInfo();
        this.renderStudentCodeTags();
        this.renderTabs();
        this.renderQDTabs();
    }

    renderStudentCodeTags() {
        const tagsEl = document.getElementById('studentCodeTags');
        if (!tagsEl) return;

        if (!this.selectedStudentCodes.size) {
            tagsEl.innerHTML = '';
            return;
        }

        tagsEl.innerHTML = Array.from(this.selectedStudentCodes)
            .map((code) => `<span class="class-tag">${code}<button type="button" data-rm-mssv="${code}">x</button></span>`)
            .join('');

        tagsEl.querySelectorAll('button[data-rm-mssv]').forEach((btn) => {
            btn.addEventListener('click', () => this.removeStudentCode(btn.dataset.rmMssv || ''));
        });
    }

    removeStudentCode(code) {
        const key = String(code || '').trim();
        if (!key) return;
        this.selectedStudentCodes.delete(key);

        const input = document.getElementById('studentCodesInput');
        if (input) {
            input.value = Array.from(this.selectedStudentCodes).join('\n');
        }

        this.updateStudentCodeFilterInfo();
        this.renderStudentCodeTags();
        this.renderTabs();
        this.renderQDTabs();
    }

    initSearch() {
        const inp = document.getElementById('searchInput');
        if (!inp) return;
        inp.value = '';
        if (this._searchHandler) {
            inp.removeEventListener('input', this._searchHandler);
        }
        this._searchHandler = () => this.renderList(inp.value, true);
        inp.addEventListener('input', this._searchHandler);
        inp.addEventListener('focus', () => this.renderList(inp.value, true));
        inp.addEventListener('click', () => this.renderList(inp.value, true));
    }

    saveSelLop() {}

    renderList(keyword = '', forceShow = false) {
        const listEl = document.getElementById('classSearchList');
        if (!listEl) return;

        const q = String(keyword || '').trim().toLowerCase();
        const filtered = this.allKeys
            .filter((k) => String(k).toLowerCase().includes(q))
            .sort((a, b) => String(a).localeCompare(String(b)));

        if (!filtered.length) {
            listEl.innerHTML = '<div class="class-search-item">Không có lớp phù hợp</div>';
            if (forceShow) {
                listEl.classList.add('show');
            }
            return;
        }

        listEl.innerHTML = filtered
            .map((k) => {
                const checked = this.selLop.has(k) ? 'checked' : '';
                const schoolCode = this.manager.extractSchoolCodeFromClass(k);
                const schoolName = schoolCode ? this.manager.schoolCodeMap[schoolCode] || '' : '';
                const detail = schoolCode ? `<div class="class-search-sub">${schoolCode}${schoolName ? ` - ${schoolName}` : ''}</div>` : '';
                return `<label class="class-search-item"><input type="checkbox" data-k="${k}" ${checked}><div class="class-search-meta"><div class="class-search-code">${k}</div>${detail}</div></label>`;
            })
            .join('');

        listEl.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
            cb.addEventListener('change', () => {
                const k = cb.dataset.k || '';
                if (!k) return;
                if (cb.checked) {
                    this.selLop.add(k);
                } else {
                    this.selLop.delete(k);
                }
                this.saveSelLop();
                this.renderTags();
                this.renderTabs();
                this.renderQDTabs();
            });
        });

        if (forceShow) {
            listEl.classList.add('show');
        }
    }

    renderTags() {
        const tagsEl = document.getElementById('classTags');
        if (!tagsEl) return;
        if (!this.selLop.size) {
            tagsEl.innerHTML = '';
            return;
        }

        tagsEl.innerHTML = Array.from(this.selLop)
            .map((k) => `<span class="class-tag">${k}<button type="button" data-rm="${k}">x</button></span>`)
            .join('');

        tagsEl.querySelectorAll('button[data-rm]').forEach((btn) => {
            btn.addEventListener('click', () => this.removeLop(btn.dataset.rm || ''));
        });
    }

    removeLop(k) {
        this.selLop.delete(k);
        this.saveSelLop();
        this.renderList(document.getElementById('searchInput').value, true);
        this.renderTags();
        this.renderTabs();
        this.renderQDTabs();
    }

    selectAll() {
        this.selLop = new Set(this.allKeys);
        this.saveSelLop();
        this.renderList(document.getElementById('searchInput').value, true);
        this.renderTags();
        this.renderTabs();
        this.renderQDTabs();
    }

    clearAll() {
        this.selLop = new Set();
        this.saveSelLop();
        this.renderList(document.getElementById('searchInput').value, true);
        this.renderTags();
        this.renderTabs();
        this.renderQDTabs();
    }

    renderTabs() {
        this.renderScoreSheet();
    }

    renderQDTabs() { }

    applySelectedClasses(report) {
        if (!this.selLop || this.selLop.size === 0) {
            return report;
        }

        const rows = report.rows.filter((row) => this.selLop.has(row.MaLop));
        const total = rows.length;
        const passed = rows.filter((row) => row.KetQua === 'Đạt').length;
        const classCodes = [...new Set(rows.map((row) => String(row.MaLop || '').trim()).filter(Boolean))];

        return {
            ...report,
            rows,
            total,
            passed,
            failed: total - passed,
            classCode: classCodes.length === 1 ? classCodes[0] : '',
            schoolName: this.manager.getDominantSchoolName(rows) || report.schoolName || '',
            groupName: this.manager.getDisplayGroupName(rows) || report.groupName || ''
        };
    }

    applySelectedStudentCodes(report) {
        if (!this.selectedStudentCodes || this.selectedStudentCodes.size === 0) {
            return report;
        }

        const rows = report.rows.filter((row) => this.selectedStudentCodes.has(String(row.MaSV || '').trim()));
        const total = rows.length;
        const passed = rows.filter((row) => row.KetQua === 'Đạt').length;
        const classCodes = [...new Set(rows.map((row) => String(row.MaLop || '').trim()).filter(Boolean))];
        const isSingleClass = classCodes.length === 1;

        return {
            ...report,
            rows,
            total,
            passed,
            failed: total - passed,
            classCode: isSingleClass ? classCodes[0] : '',
            schoolName: this.manager.getDominantSchoolName(rows) || report.schoolName || '',
            groupName: isSingleClass ? (this.manager.getDisplayGroupName(rows) || report.groupName || '') : ''
        };
    }

    switchTab(tabName) {
        document.querySelectorAll('.tab-content').forEach((tab) => tab.classList.remove('active'));
        document.querySelectorAll('.tab-btn').forEach((btn) => btn.classList.remove('active'));

        document.getElementById(tabName).classList.add('active');
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        if (tabName === 'by-class') this.renderByClass();
        if (tabName === 'by-subject') this.renderBySubject();
        if (tabName === 'score-sheet') this.renderScoreSheet();
        if (tabName === 'statistics') this.renderStatistics();
    }

    handleStudentFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;
        this.importStudentFile(file);
    }

    handleGradeFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;
        this.importGradeFile(file);
    }

    importStudentFile(file) {
        if (!file) return;
        this.showStatus('Đang xử lý file, vui lòng chờ...', 'info', 'studentStatus', false);

        if (this.isCSVFile(file)) {
            this.readAsText(file, (text) => {
                try {
                    const data = CSVParser.parse(text);
                    if (!data.length) {
                        this.showStatus('✗ File không có dữ liệu', 'error', 'studentStatus');
                        return;
                    }
                    this.manager.importStudentInfo(data);
                    this.showStatus(`✓ Nhập danh sách học sinh thành công! (${data.length} học sinh)`, 'success', 'studentStatus');
                    this.render();
                } catch (error) {
                    this.showStatus(`✗ Lỗi: ${error.message}`, 'error', 'studentStatus');
                }
            }, 'studentStatus');
            return;
        }

        if (this.isExcelFile(file)) {
            this.readAsArrayBuffer(file, (buffer) => {
                SimpleXLSXParser.parse(buffer)
                    .then((data) => {
                        if (!data.length) {
                            this.showStatus('✗ File không có dữ liệu', 'error', 'studentStatus');
                            return;
                        }
                        this.manager.importStudentInfo(data);
                        this.showStatus(`✓ Nhập danh sách học sinh thành công! (${data.length} học sinh)`, 'success', 'studentStatus');
                        this.render();
                    })
                    .catch((error) => this.showStatus(`✗ Lỗi: ${error.message}`, 'error', 'studentStatus'));
            }, 'studentStatus');
            return;
        }

        this.showStatus('✗ Định dạng file không hỗ trợ. Vui lòng chọn .csv, .xlsx hoặc .xls', 'error', 'studentStatus');
    }

    importBirthFile(file) {
        if (!file) return;
        this.showStatus('Đang xử lý file Nơi sinh, vui lòng chờ...', 'info', 'studentStatus', false);

        const onSuccess = (data) => {
            if (!data.length) {
                this.showStatus('✗ File Nơi sinh không có dữ liệu', 'error', 'studentStatus');
                return;
            }
            this.manager.importStudentInfo(data);
            this.showStatus(`✓ Nhập file Nơi sinh thành công! (${data.length} bản ghi)`, 'success', 'studentStatus');
            this.render();
        };

        if (this.isCSVFile(file)) {
            this.readAsText(file, (text) => {
                try {
                    onSuccess(CSVParser.parse(text));
                } catch (error) {
                    this.showStatus(`✗ Lỗi: ${error.message}`, 'error', 'studentStatus');
                }
            }, 'studentStatus');
            return;
        }

        if (this.isExcelFile(file)) {
            this.readAsArrayBuffer(file, (buffer) => {
                SimpleXLSXParser.parse(buffer)
                    .then(onSuccess)
                    .catch((error) => this.showStatus(`✗ Lỗi: ${error.message}`, 'error', 'studentStatus'));
            }, 'studentStatus');
            return;
        }

        this.showStatus('✗ Định dạng file không hỗ trợ. Vui lòng chọn .csv, .xlsx hoặc .xls', 'error', 'studentStatus');
    }

    importGradeFile(file) {
        if (!file) return;
        this.showStatus('Đang xử lý file, vui lòng chờ...', 'info', 'gradeStatus', false);

        if (this.isCSVFile(file)) {
            this.readAsText(file, (text) => {
                try {
                    const data = CSVParser.parse(text);
                    if (!data.length) {
                        this.showStatus('✗ File không có dữ liệu', 'error', 'gradeStatus');
                        return;
                    }
                    this.manager.importGradeData(data);
                    this.showStatus(`✓ Nhập bảng điểm thành công! (${data.length} bản ghi)`, 'success', 'gradeStatus');
                    this.render();
                } catch (error) {
                    this.showStatus(`✗ Lỗi: ${error.message}`, 'error', 'gradeStatus');
                }
            }, 'gradeStatus');
            return;
        }

        if (this.isExcelFile(file)) {
            this.readAsArrayBuffer(file, (buffer) => {
                SimpleXLSXParser.parse(buffer)
                    .then((data) => {
                        if (!data.length) {
                            this.showStatus('✗ File không có dữ liệu', 'error', 'gradeStatus');
                            return;
                        }
                        this.manager.importGradeData(data);
                        this.showStatus(`✓ Nhập bảng điểm thành công! (${data.length} bản ghi)`, 'success', 'gradeStatus');
                        this.render();
                    })
                    .catch((error) => this.showStatus(`✗ Lỗi: ${error.message}`, 'error', 'gradeStatus'));
            }, 'gradeStatus');
            return;
        }

        this.showStatus('✗ Định dạng file không hỗ trợ. Vui lòng chọn .csv, .xlsx hoặc .xls', 'error', 'gradeStatus');
    }

    isCSVFile(file) {
        return file.name.toLowerCase().endsWith('.csv');
    }

    isExcelFile(file) {
        const lowerName = file.name.toLowerCase();
        return lowerName.endsWith('.xlsx') || lowerName.endsWith('.xls');
    }

    readAsText(file, callback, statusElementId) {
        const reader = new FileReader();
        reader.onload = (e) => callback(e.target.result);
        reader.onerror = () => this.showStatus('✗ Không thể đọc file', 'error', statusElementId);
        reader.readAsText(file, 'UTF-8');
    }

    readAsArrayBuffer(file, callback, statusElementId) {
        const reader = new FileReader();
        reader.onload = (e) => callback(e.target.result);
        reader.onerror = () => this.showStatus('✗ Không thể đọc file', 'error', statusElementId);
        reader.readAsArrayBuffer(file);
    }

    showStatus(message, type, elementId, autoHide = true) {
        const statusEl = document.getElementById(elementId);
        if (!statusEl) return;
        statusEl.textContent = message;
        statusEl.className = `status-message ${type}`;
        if (autoHide) {
            setTimeout(() => {
                statusEl.className = 'status-message';
            }, 3000);
        }
    }

    clearAllData() {
        if (!confirm('Bạn chắc chắn muốn xóa tất cả dữ liệu?')) return;
        this.manager.clearAll();
        this.showStatus('✓ Đã xóa tất cả dữ liệu', 'success', 'studentStatus');
        this.render();
    }

    render() {
        this.updateFilters();
        this.renderActiveTab();
    }

    renderActiveTab() {
        const activeTab = document.querySelector('.tab-content.active');
        if (!activeTab) return;

        if (activeTab.id === 'by-class') this.renderByClass();
        if (activeTab.id === 'by-subject') this.renderBySubject();
        if (activeTab.id === 'score-sheet') this.renderScoreSheet();
        if (activeTab.id === 'statistics') this.renderStatistics();
    }

    updateFilters() {
        const classFilter = document.getElementById('classFilter');
        const currentClass = classFilter.value;
        classFilter.innerHTML = '<option value="">-- Tất Cả Lớp --</option>';
        this.manager.classes.forEach((cls) => {
            const option = document.createElement('option');
            option.value = cls;
            option.textContent = cls;
            classFilter.appendChild(option);
        });
        classFilter.value = currentClass;

        const subjectFilter = document.getElementById('subjectFilter');
        const currentSubject = subjectFilter.value;
        subjectFilter.innerHTML = '<option value="">-- Tất Cả Môn --</option>';
        this.manager.subjects.forEach((subject) => {
            const option = document.createElement('option');
            option.value = subject;
            option.textContent = subject;
            subjectFilter.appendChild(option);
        });
        subjectFilter.value = currentSubject;

        const reportSubjectFilter = document.getElementById('reportSubjectFilter');
        const currentReportSubject = reportSubjectFilter.value;
        reportSubjectFilter.innerHTML = '<option value="">-- Chọn Mã Môn --</option>';
        this.manager.getReportSubjectOptions().forEach((subjectCode) => {
            const option = document.createElement('option');
            option.value = subjectCode;
            option.textContent = subjectCode;
            reportSubjectFilter.appendChild(option);
        });
        reportSubjectFilter.value = currentReportSubject;

        const reportClassFilter = document.getElementById('reportClassFilter');
        const currentReportClass = reportClassFilter.value;
        const selectedReportSubject = reportSubjectFilter.value;
        reportClassFilter.innerHTML = '<option value="">-- Tất Cả Lớp --</option>';
        const reportClasses = this.manager.getReportClassOptions(selectedReportSubject);
        this.allKeys = [...reportClasses];
        this.selLop.forEach((k) => {
            if (!this.allKeys.includes(k)) {
                this.selLop.delete(k);
            }
        });
        reportClasses.forEach((cls) => {
            const option = document.createElement('option');
            option.value = cls;
            option.textContent = cls;
            reportClassFilter.appendChild(option);
        });
        reportClassFilter.value = reportClasses.includes(currentReportClass) ? currentReportClass : '';

        const reportSchoolFilter = document.getElementById('reportSchoolFilter');
        if (reportSchoolFilter) {
            const currentReportSchool = reportSchoolFilter.value;
            const schoolOptions = this.manager.getSchoolOptions();
            reportSchoolFilter.innerHTML = '<option value="">-- Tự động theo dữ liệu --</option>';
            schoolOptions.forEach((school) => {
                const option = document.createElement('option');
                option.value = school;
                option.textContent = school;
                reportSchoolFilter.appendChild(option);
            });
            reportSchoolFilter.value = schoolOptions.includes(currentReportSchool) ? currentReportSchool : '';
        }

        this.renderList(document.getElementById('searchInput').value, false);
        this.renderTags();
        this.updateStudentCodeFilterInfo();
        this.renderStudentCodeTags();
    }

    renderByClass() {
        const selectedClass = document.getElementById('classFilter').value;
        const data = this.manager.getDataByClass(selectedClass);
        const content = document.getElementById('byClassContent');

        if (!data.length) {
            content.innerHTML = '<p class="empty-message">Chưa có dữ liệu</p>';
            return;
        }

        const studentMap = {};
        data.forEach((row) => {
            if (!studentMap[row.MaSV]) {
                studentMap[row.MaSV] = {
                    MaSV: row.MaSV,
                    HoLotSV: row.HoLotSV,
                    TenSV: row.TenSV,
                    NgaySinhC: row.NgaySinhC,
                    MaLop: row.MaLop,
                    grades: {}
                };
            }
            studentMap[row.MaSV].grades[row.MaMH] = {
                QT: row.QT,
                Thi: row.Thi,
                T1_DTK: row.T1_DTK,
                L1: row.L1,
                T2_DTK: row.T2_DTK,
                L2: row.L2,
                T3_T3: row.T3_T3
            };
        });

        let html = '<table><thead><tr><th>MaSV</th><th>Họ Lót</th><th>Tên</th><th>Ngày Sinh</th>';
        this.manager.subjects.forEach((subject) => {
            html += `<th colspan="7" style="text-align: center;">${subject}</th>`;
        });
        html += '</tr><tr><th></th><th></th><th></th><th></th>';
        this.manager.subjects.forEach(() => {
            html += '<th>QT</th><th>Thi</th><th>T1_ĐTK</th><th>L1</th><th>T2_ĐTK</th><th>L2</th><th>T3_T3</th>';
        });
        html += '</tr></thead><tbody>';

        Object.values(studentMap).forEach((student) => {
            html += `<tr>
                <td>${student.MaSV}</td>
                <td>${student.HoLotSV}</td>
                <td>${student.TenSV}</td>
                <td>${student.NgaySinhC}</td>`;

            this.manager.subjects.forEach((subject) => {
                const grade = student.grades[subject] || {};
                html += `<td>${grade.QT || '-'}</td>
                         <td>${grade.Thi || '-'}</td>
                         <td>${grade.T1_DTK || '-'}</td>
                         <td>${grade.L1 || '-'}</td>
                         <td>${grade.T2_DTK || '-'}</td>
                         <td>${grade.L2 || '-'}</td>
                         <td>${grade.T3_T3 || '-'}</td>`;
            });

            html += '</tr>';
        });

        html += '</tbody></table>';
        content.innerHTML = html;
    }

    renderBySubject() {
        const selectedSubject = document.getElementById('subjectFilter').value;
        const data = this.manager.getDataBySubject(selectedSubject);
        const content = document.getElementById('bySubjectContent');

        if (!data.length) {
            content.innerHTML = '<p class="empty-message">Chưa có dữ liệu</p>';
            return;
        }

        let html = '<table><thead><tr><th>MaSV</th><th>Họ Lót</th><th>Tên</th><th>Ngày Sinh</th><th>Lớp</th><th>QT</th><th>Thi</th><th>T1_ĐTK</th><th>L1</th><th>T2_ĐTK</th><th>L2</th><th>T3_T3</th></tr></thead><tbody>';
        data.forEach((row) => {
            html += `<tr>
                <td>${row.MaSV}</td>
                <td>${row.HoLotSV}</td>
                <td>${row.TenSV}</td>
                <td>${row.NgaySinhC}</td>
                <td>${row.MaLop}</td>
                <td>${row.QT}</td>
                <td>${row.Thi}</td>
                <td>${row.T1_DTK}</td>
                <td>${row.L1}</td>
                <td>${row.T2_DTK}</td>
                <td>${row.L2}</td>
                <td>${row.T3_T3}</td>
            </tr>`;
        });
        html += '</tbody></table>';
        content.innerHTML = html;
    }

    renderStatistics() {
        const stats = this.manager.getStatistics();
        const content = document.getElementById('statisticsContent');

        if (stats.totalStudents === 0) {
            content.innerHTML = '<p class="empty-message">Chưa có dữ liệu</p>';
            return;
        }

        let html = `
            <div class="stat-card">
                <h3>Tổng Học Sinh</h3>
                <div class="value">${stats.totalStudents}</div>
            </div>
            <div class="stat-card">
                <h3>Tổng Môn Học</h3>
                <div class="value">${stats.totalSubjects}</div>
            </div>
            <div class="stat-card">
                <h3>Tổng Lớp</h3>
                <div class="value">${stats.totalClasses}</div>
            </div>
            <div class="stat-card">
                <h3>Điểm Thi Trung Bình</h3>
                <div class="value">${stats.averageScore}</div>
            </div>
        `;

        html += '<h3 style="grid-column: 1/-1; margin-top: 20px; color: #333;">Thống Kê Theo Môn</h3>';
        this.manager.subjects.forEach((subject) => {
            const stat = stats.subjectStats[subject];
            html += `
                <div class="stat-card">
                    <h3>${subject}</h3>
                    <div class="detail">Số học sinh: ${stat.count}</div>
                    <div class="detail">Trung bình: ${stat.average}</div>
                    <div class="detail">Cao nhất: ${stat.max}</div>
                    <div class="detail">Thấp nhất: ${stat.min}</div>
                </div>
            `;
        });

        content.innerHTML = html;
    }

    formatScore(score) {
        if (score === null || score === undefined || String(score).trim() === '') {
            return '-';
        }

        if (typeof score === 'string') {
            const scoreText = score.trim();
            const parsed = Number(scoreText.replace(',', '.'));
            if (Number.isNaN(parsed)) {
                return scoreText;
            }
            return parsed.toFixed(1);
        }

        return Number(score).toFixed(1);
    }

    getCurrentReport() {
        const selectedSubject = document.getElementById('reportSubjectFilter').value;
        const selectedClass = document.getElementById('reportClassFilter').value;

        if (!selectedSubject) {
            this.showStatus('✗ Vui lòng chọn Mã Môn trước khi xuất file', 'error', 'gradeStatus');
            return null;
        }

        let report = this.manager.getScoreSheetData(selectedSubject, selectedClass);
        report = this.applySelectedClasses(report);
        report = this.applySelectedStudentCodes(report);
        const schoolOverride = this.getSchoolOverride();
        if (schoolOverride) {
            report = { ...report, schoolName: schoolOverride, schoolNameOverride: schoolOverride };
        }
        if (!report.rows.length) {
            this.showStatus('✗ Không có dữ liệu để xuất theo bộ lọc hiện tại', 'error', 'gradeStatus');
            return null;
        }

        return report;
    }

    getTodayText() {
        const now = new Date();
        const day = String(now.getDate()).padStart(2, '0');
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const year = now.getFullYear();
        return `Vĩnh Long, ngày ${day} tháng ${month} năm ${year}`;
    }

    getDecisionIssueText() {
        const now = new Date();
        const day = String(now.getDate()).padStart(2, '0');
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const year = now.getFullYear();
        return `(Ban hành kèm theo Quyết định số      /QĐ-GDQP ngày ${day} tháng ${month} năm ${year})`;
    }

    getSchoolOverride() {
        const select = document.getElementById('reportSchoolFilter');
        return select ? String(select.value || '').trim() : '';
    }

    sanitizeFileName(text) {
        return String(text || '')
            .trim()
            .replace(/[\\/:*?"<>|]/g, '_')
            .replace(/\s+/g, '_');
    }

    shouldShowClassColumn(subjectCode) {
        return this.manager.normalizeSubjectCode(subjectCode) !== '190036';
    }

    buildScoreSheetAOA(report, options = {}) {
        const includeClassColumn = options.includeClassColumn !== false;
        const groupLabelOverride = options.groupLabelOverride;
        const colCount = includeClassColumn ? 10 : 9;
        const blankRow = () => new Array(colCount).fill('');
        const rowPad = (arr) => {
            const out = arr.slice(0, colCount);
            while (out.length < colCount) out.push('');
            return out;
        };
        const aoa = [];
        const displaySchoolOverride = String(report.schoolNameOverride || '').trim();
        const schoolDisplay = displaySchoolOverride || this.manager.getSchoolDisplayText(report.classCode, report.schoolName || report.majorName || '');
        const displayGroup = groupLabelOverride != null
            ? String(groupLabelOverride)
            : this.manager.formatClassDisplay(report.classCode, report.groupName);

        aoa.push(rowPad(['ĐẠI HỌC TRÀ VINH', '', '', '', 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM']));
        aoa.push(rowPad(['TRUNG TÂM GIÁO DỤC QUỐC PHÒNG', '', '', '', 'Độc lập - Tự do - Hạnh phúc']));
        aoa.push(rowPad(['VÀ AN NINH ĐẠI HỌC TRÀ VINH', '', '', '', this.getTodayText()]));
        aoa.push(blankRow());
        aoa.push(rowPad(['DANH SÁCH GHI ĐIỂM MÔN GIÁO DỤC QUỐC PHÒNG VÀ AN NINH']));

        aoa.push(rowPad([`Trường: ${schoolDisplay || report.schoolName || report.majorName || ''}`]));
        aoa.push(rowPad([`Nhóm/ lớp: ${displayGroup}`]));
        const _subjectFull = this.manager.getSubjectDisplayText(report.subjectCode);
        const _detailIdx = _subjectFull.indexOf('(Trình độ');
        const _monHocTitle = _detailIdx >= 0
            ? `Môn học: ${_subjectFull.substring(0, _detailIdx).trim()}`
            : `Môn học: ${_subjectFull}`;
        const _monHocDetail = _detailIdx >= 0 ? _subjectFull.substring(_detailIdx).trim() : '';
        aoa.push(rowPad([_monHocTitle, '', '', '', '', _monHocDetail]));
        aoa.push(blankRow());

        const header = ['TT', 'MSSV', 'HỌ VÀ', 'TÊN', 'NGÀY SINH', 'NGÀNH HỌC', 'ĐIỂM TB', 'XẾP LOẠI', 'KẾT QUẢ', 'GHI CHÚ'];
        if (includeClassColumn) header.push('Mã lớp');
        aoa.push(header);

        report.rows.forEach((row, index) => {
            const dataRow = [
                index + 1,
                row.MaSV,
                row.HoVa,
                row.Ten,
                row.NgaySinhC,
                row.NganhHoc || row.TenNganh || '',
                this.formatScore(row.DiemTB),
                row.XepLoai,
                row.KetQua,
                ''
            ];
            if (includeClassColumn) dataRow.push(row.MaLop);
            aoa.push(dataRow);
        });

        aoa.push(rowPad(['Ghi chú: "CT" - Cấm thi.']));
        
        const stats = this.manager.calculateStatistics(report.rows);
        aoa.push(rowPad(['', 'Tổng số SV trên danh sách:', '', stats.total.toString(), 'sinh viên/học sinh']));
        aoa.push(rowPad(['', 'Số sinh viên đạt:', '', stats.passed.toString(), 'sinh viên/học sinh']));
        aoa.push(rowPad(['', 'Cấm thi:', '', stats.banned.toString(), 'sinh viên/học sinh']));
        aoa.push(rowPad(['', 'Vắng thi:', '', stats.absent.toString(), 'sinh viên/học sinh']));
        aoa.push(rowPad(['', 'Số sinh viên hỏng:', '', stats.failed.toString(), 'sinh viên/học sinh']));
        aoa.push(rowPad(['', 'Đình chỉ:', '', stats.suspended.toString(), 'sinh viên/học sinh']));
        aoa.push(rowPad(['', 'Chưa học:', '', stats.notStudied.toString(), 'sinh viên/học sinh']));

        aoa.push(blankRow());
        aoa.push(rowPad(['Cán bộ ghi điểm', '', 'Phòng ĐT, QLSV', '', '', '', 'KT. GIÁM ĐỐC']));
        aoa.push(rowPad(['', '', '', '', '', '', 'PHÓ GIÁM ĐỐC']));
        for (let i = 0; i < 5; i++) {
            aoa.push(blankRow());
        }
        aoa.push(rowPad(['Trương Tấn Tài', '', 'Cao Nguyên Ty', '', '', '', 'Trương Minh Hải']));

        return aoa;
    }

    buildReportFromRows(baseReport, rows, explicitClassCode = '') {
        const total = rows.length;
        const passed = rows.filter((row) => row.KetQua === 'Đạt').length;
        const classCodes = [...new Set(rows.map((row) => String(row.MaLop || '').trim()).filter(Boolean))];
        const classCode = explicitClassCode || (classCodes.length === 1 ? classCodes[0] : '');
        const schoolNameOverride = String(baseReport.schoolNameOverride || '').trim();
        const schoolName = schoolNameOverride || this.manager.getDominantSchoolName(rows) || baseReport.schoolName || '';
        const groupName = this.manager.getDisplayGroupName(rows) || baseReport.groupName || '';
        return {
            ...baseReport,
            rows,
            total,
            passed,
            failed: total - passed,
            classCode,
            schoolName,
            groupName
        };
    }

    styleScoreSheetWorksheet(worksheet, reportRowsCount, includeClassColumn) {
        const allBorders = {
            top: { style: 'thin' }, bottom: { style: 'thin' },
            left: { style: 'thin' }, right: { style: 'thin' }
        };
        const mkFont = (sz, bold, italic) => ({ name: 'Times New Roman', sz, bold: !!bold, italic: !!italic });
        const N = reportRowsCount;
        const lastCol = includeClassColumn ? 9 : 8;

        const range = XLSX.utils.decode_range(worksheet['!ref']);
        for (let r = 0; r <= range.e.r; r++) {
            for (let c = 0; c <= range.e.c; c++) {
                const cell = XLSX.utils.encode_cell({ r, c });
                if (!worksheet[cell]) continue;
                worksheet[cell].s = {};
                const s = worksheet[cell].s;

                if (r === 0) {
                    s.font = c >= 4 ? mkFont(12, true, false) : mkFont(12, false, false);
                    s.alignment = { horizontal: 'center', vertical: 'center' };
                } else if (r === 1) {
                    s.font = c >= 4
                        ? { name: 'Arial', sz: 12, bold: true, underline: true }
                        : mkFont(12, true, false);
                    s.alignment = { horizontal: 'center', vertical: 'center' };
                } else if (r === 2) {
                    if (c < 4) {
                        s.font = { name: 'Arial', sz: 12, bold: true, underline: true };
                        s.alignment = { horizontal: 'center', vertical: 'center' };
                    } else {
                        s.font = mkFont(12, false, true);
                        s.alignment = { horizontal: 'right', vertical: 'center' };
                    }
                } else if (r === 4) {
                    s.font = mkFont(14, true, false);
                    s.alignment = { horizontal: 'center', vertical: 'center' };
                } else if (r === 5) {
                    s.font = mkFont(12, true, false);
                    s.alignment = { horizontal: 'left', vertical: 'center' };
                } else if (r === 6) {
                    s.font = mkFont(11, true, false);
                    s.alignment = { horizontal: 'left', vertical: 'center' };
                } else if (r === 7) {
                    s.alignment = { horizontal: 'left', vertical: 'center' };
                    if (c <= 4) {
                        s.font = mkFont(12, true, false);
                    } else {
                        s.font = mkFont(11, false, true);
                    }
                } else if (r === 9) {
                    s.font = mkFont(10, true, false);
                    s.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
                    s.border = allBorders;
                } else if (r >= 10 && r <= 9 + N) {
                    s.font = mkFont(11, false, false);
                    s.border = allBorders;
                    s.alignment = (c === 2 || c === 3)
                        ? { horizontal: 'left', vertical: 'center' }
                        : { horizontal: 'center', vertical: 'center' };
                } else if (r === 10 + N) {
                    s.font = mkFont(12, false, false);
                    s.alignment = { horizontal: 'left', vertical: 'center' };
                } else if (r === 11 + N || r === 12 + N || r === 13 + N) {
                    if (c === 3) s.font = mkFont(12, true, false);
                    else if (c === 4 || c === 5) s.font = mkFont(12, false, true);
                    else s.font = mkFont(12, false, false);
                    s.alignment = { horizontal: 'left', vertical: 'center', wrapText: true };
                } else if (r === 15 + N) {
                    s.font = mkFont(12, true, false);
                    s.alignment = { horizontal: 'center', vertical: 'center' };
                } else if (r === 16 + N) {
                    s.font = mkFont(11, true, false);
                    s.alignment = { horizontal: 'center', vertical: 'center' };
                } else if (r === 22 + N) {
                    s.font = mkFont(11, true, false);
                    s.alignment = { horizontal: 'center', vertical: 'center' };
                } else {
                    s.font = mkFont(11, false, false);
                }
            }
        }

        worksheet['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } },
            { s: { r: 0, c: 4 }, e: { r: 0, c: lastCol - 1 } },
            { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } },
            { s: { r: 1, c: 4 }, e: { r: 1, c: lastCol - 1 } },
            { s: { r: 2, c: 0 }, e: { r: 2, c: 3 } },
            { s: { r: 2, c: 4 }, e: { r: 2, c: lastCol - 1 } },
            { s: { r: 4, c: 0 }, e: { r: 4, c: lastCol } },
            { s: { r: 5, c: 0 }, e: { r: 5, c: lastCol } },
            { s: { r: 6, c: 0 }, e: { r: 6, c: lastCol } },
            { s: { r: 7, c: 0 }, e: { r: 7, c: 4 } },
            { s: { r: 7, c: 5 }, e: { r: 7, c: lastCol } },
            { s: { r: 10 + N, c: 0 }, e: { r: 10 + N, c: lastCol } },
            { s: { r: 10 + N, c: 1 }, e: { r: 10 + N, c: 2 } },
            { s: { r: 11 + N, c: 1 }, e: { r: 11 + N, c: 2 } },
            { s: { r: 11 + N, c: 4 }, e: { r: 11 + N, c: 5 } },
            { s: { r: 12 + N, c: 1 }, e: { r: 12 + N, c: 2 } },
            { s: { r: 12 + N, c: 4 }, e: { r: 12 + N, c: 5 } },
            { s: { r: 13 + N, c: 1 }, e: { r: 13 + N, c: 2 } },
            { s: { r: 13 + N, c: 4 }, e: { r: 13 + N, c: 5 } },
            { s: { r: 15 + N, c: 0 }, e: { r: 15 + N, c: 1 } },
            { s: { r: 15 + N, c: 2 }, e: { r: 15 + N, c: 5 } },
            { s: { r: 15 + N, c: 6 }, e: { r: 15 + N, c: lastCol } },
            { s: { r: 16 + N, c: 6 }, e: { r: 16 + N, c: lastCol } },
            { s: { r: 22 + N, c: 0 }, e: { r: 22 + N, c: 1 } },
            { s: { r: 22 + N, c: 2 }, e: { r: 22 + N, c: 5 } },
            { s: { r: 22 + N, c: 6 }, e: { r: 22 + N, c: lastCol } }
        ];

        worksheet['!cols'] = includeClassColumn
            ? [
                { wch: 4 }, { wch: 14 }, { wch: 20 }, { wch: 10 }, { wch: 12 },
                { wch: 8 }, { wch: 11 }, { wch: 8 }, { wch: 10 }, { wch: 24 }
            ]
            : [
                { wch: 4 }, { wch: 14 }, { wch: 20 }, { wch: 10 }, { wch: 12 },
                { wch: 8 }, { wch: 11 }, { wch: 8 }, { wch: 10 }
            ];
    }

    async exportScoreSheetExcel() {
        const report = this.getCurrentReport();
        if (!report) return;

        try {
            await SimpleXLSXParser.ensureLibraryLoaded();
            const workbook = XLSX.utils.book_new();
            const usedSheetNames = new Set();
            const makeUniqueSheetName = (rawName) => {
                const base = this.sanitizeFileName(rawName || 'Sheet').substring(0, 31) || 'Sheet';
                if (!usedSheetNames.has(base)) {
                    usedSheetNames.add(base);
                    return base;
                }
                let i = 1;
                while (i < 1000) {
                    const suffix = `_${i}`;
                    const candidate = `${base.substring(0, Math.max(1, 31 - suffix.length))}${suffix}`;
                    if (!usedSheetNames.has(candidate)) {
                        usedSheetNames.add(candidate);
                        return candidate;
                    }
                    i++;
                }
                return `Sheet_${Date.now()}`.substring(0, 31);
            };
            const includeClassColumn = this.shouldShowClassColumn(report.subjectCode);
            const summaryReport = this.buildReportFromRows(report, [...report.rows], '');
            const summaryAoa = this.buildScoreSheetAOA(summaryReport, {
                includeClassColumn,
                groupLabelOverride: 'tổng hợp'
            });
            const summaryWs = XLSX.utils.aoa_to_sheet(summaryAoa);
            this.styleScoreSheetWorksheet(summaryWs, summaryReport.rows.length, includeClassColumn);
            XLSX.utils.book_append_sheet(workbook, summaryWs, makeUniqueSheetName('TongHop'));

            const rowsByClass = new Map();
            report.rows.forEach((row) => {
                const classCode = String(row.MaLop || '').trim() || 'KhongRo';
                if (!rowsByClass.has(classCode)) rowsByClass.set(classCode, []);
                rowsByClass.get(classCode).push(row);
            });

            [...rowsByClass.keys()].sort((a, b) => a.localeCompare(b)).forEach((classCode) => {
                const classRows = rowsByClass.get(classCode) || [];
                const classReport = this.buildReportFromRows(report, classRows, classCode);
                const classAoa = this.buildScoreSheetAOA(classReport, { includeClassColumn });
                const classWs = XLSX.utils.aoa_to_sheet(classAoa);
                this.styleScoreSheetWorksheet(classWs, classReport.rows.length, includeClassColumn);
                XLSX.utils.book_append_sheet(
                    workbook,
                    classWs,
                    makeUniqueSheetName(classCode || 'KhongRo')
                );
            });

            const fileName = this.sanitizeFileName(`Bang_diem_${report.subjectCode || 'MonHoc'}_TongHop_va_theo_lop.xlsx`);
            
            // Add sheets for each category
            const categories = {
                'CT': (row) => {
                    const ketQua = String(row.KetQua || '').trim();
                    // CHỈ CT - không lẫn VT
                    return ketQua === 'CT';
                },
                'VT': (row) => {
                    const ketQua = String(row.KetQua || '').trim();
                    // CHỈ VT
                    return ketQua === 'VT';
                },
                'Hong': (row) => {
                    const ketQua = String(row.KetQua || '').trim();
                    const diemTB = row.DiemTB;
                    const diemStr = String(diemTB || '').trim();
                    const diemNumeric = diemStr ? Number(diemStr.replace(',', '.')) : NaN;
                    // Hỏng = KetQua là Hỏng hoặc điểm < 5.0 (nhưng > 0)
                    return (ketQua === 'Hỏng') || (diemNumeric > 0 && diemNumeric < 5.0 && !Number.isNaN(diemNumeric));
                },
                'ChuaHoc': (row) => {
                    const diemTB = row.DiemTB;
                    const diemStr = String(diemTB || '').trim();
                    const diemNumeric = diemStr ? Number(diemStr.replace(',', '.')) : NaN;
                    // Chưa học = không có điểm (0 hoặc empty/NaN)
                    return diemNumeric === 0 || (!diemStr || Number.isNaN(diemNumeric));
                },
                'DinhChi': (row) => row.KetQua === 'Đình chỉ'
            };
            
            Object.entries(categories).forEach(([categoryName, filterFn]) => {
                const categoryRows = report.rows.filter(filterFn);
                if (categoryRows.length === 0) return;
                
                const categoryReport = this.buildReportFromRows(report, categoryRows, '');
                const categoryAoa = this.buildScoreSheetAOA(categoryReport, {
                    includeClassColumn,
                    groupLabelOverride: categoryName
                });
                const categoryWs = XLSX.utils.aoa_to_sheet(categoryAoa);
                this.styleScoreSheetWorksheet(categoryWs, categoryRows.length, includeClassColumn);
                XLSX.utils.book_append_sheet(workbook, categoryWs, makeUniqueSheetName(categoryName));
            });
            
            XLSX.writeFile(workbook, fileName);
            this.showStatus('✓ Đã xuất file Excel thành công', 'success', 'gradeStatus');
        } catch (error) {
            this.showStatus(`✗ Xuất Excel thất bại: ${error.message}`, 'error', 'gradeStatus');
        }
    }

    buildDecisionAOA(report) {
        const aoa = [];
        const normalizeClassCode = (value) => String(value || '').split(',')[0].trim();
        const classCodes = [...new Set(report.rows.map((row) => normalizeClassCode(row.MaLop)).filter(Boolean))];
        const reportClassCode = classCodes.join(', ') || normalizeClassCode(report.classCode) || normalizeClassCode(report.rows[0] ? report.rows[0].MaLop : '');
        aoa.push(['ĐẠI HỌC TRÀ VINH', '', '', '', 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', '', '', '', '', '']);
        aoa.push(['TRUNG TÂM GIÁO DỤC QUỐC PHÒNG', '', '', '', 'Độc lập - Tự do - Hạnh phúc', '', '', '', '', '']);
        aoa.push(['VÀ AN NINH ĐẠI HỌC TRÀ VINH', '', '', '', this.getTodayText(), '', '', '', '', '']);
        aoa.push(['', '', '', '', '', '', '', '', '', '']);
        aoa.push(['DANH SÁCH CẤP CHỨNG CHỈ GIÁO DỤC QUỐC PHÒNG VÀ AN NINH', '', '', '', '', '', '', '', '', '']);
        aoa.push([this.getDecisionIssueText(), '', '', '', '', '', '', '', '', '']);
        aoa.push(['', '', '', '', '', '', '', '', '', '']);
        aoa.push(['', '', '', '', '', '', '', '', '', '']);

        aoa.push(['TT', 'Mã SV', 'Họ Và', 'Tên', 'Ngày Sinh', 'Ngành học', 'Điểm TB', 'Xếp Loại', 'Ghi Chú', 'Mã Lớp']);

        const uniqueByStudent = new Map();
        report.rows.forEach((row) => {
            const key = String(row.MaSV || '').trim();
            if (!key) return;
            if (!uniqueByStudent.has(key)) {
                uniqueByStudent.set(key, row);
            }
        });
        const rowsForDecision = Array.from(uniqueByStudent.values());
        const dataRowsCount = rowsForDecision.length;
        for (let i = 0; i < dataRowsCount; i++) {
            const row = rowsForDecision[i];
            aoa.push([
                i + 1,
                row.MaSV,
                row.HoVa,
                row.Ten,
                row.NgaySinhC,
                row.NganhHoc || row.TenNganh || '',
                this.formatScore(row.DiemTB),
                row.XepLoai,
                '',
                row.MaLop
            ]);
        }

        aoa.push([`Trên danh sách có ${dataRowsCount} sinh viên`, '', '', '', '', '', '', '', '', '']);
        return aoa;
    }

    async exportDecisionExcel() {
        const report = this.getCurrentReport();
        if (!report) return;

        try {
            await SimpleXLSXParser.ensureLibraryLoaded();
            const workbook = XLSX.utils.book_new();
            const aoa = this.buildDecisionAOA(report);
            const worksheet = XLSX.utils.aoa_to_sheet(aoa);
            
            const uniqueDecisionRows = new Map();
            report.rows.forEach((row) => {
                const key = String(row.MaSV || '').trim();
                if (!key || uniqueDecisionRows.has(key)) return;
                uniqueDecisionRows.set(key, row);
            });
            const dataRowsCount = uniqueDecisionRows.size;

            worksheet['!cols'] = [
                { wch: 4 }, { wch: 10 }, { wch: 16 }, { wch: 8 }, { wch: 12 },
                { wch: 14 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 24 }
            ];
            worksheet['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } },
                { s: { r: 0, c: 4 }, e: { r: 0, c: 8 } },
                { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } },
                { s: { r: 1, c: 4 }, e: { r: 1, c: 8 } },
                { s: { r: 2, c: 0 }, e: { r: 2, c: 3 } },
                { s: { r: 2, c: 4 }, e: { r: 2, c: 8 } },
                { s: { r: 4, c: 0 }, e: { r: 4, c: 9 } },
                { s: { r: 5, c: 0 }, e: { r: 5, c: 9 } },
                { s: { r: 6, c: 0 }, e: { r: 6, c: 9 } },
                { s: { r: 9 + dataRowsCount, c: 0 }, e: { r: 9 + dataRowsCount, c: 9 } }
            ];

            const allBorders = {
                top: { style: 'thin' }, bottom: { style: 'thin' },
                left: { style: 'thin' }, right: { style: 'thin' }
            };
            const mkFont = (sz, bold, italic) => ({ name: 'Arial', sz, bold: !!bold, italic: !!italic });
            // AOA layout: rows 0-2=header, 3=empty, 4=title, 5=(Ban hành...), 6-7=empty,
            //             8=table header, 9..(8+M)=data,
            //             (9+M)=empty, (10+M)=Trên danh sách, (11+M)=empty,
            //             (12+M)=chức vụ ký, (13-15+M)=empty, (16+M)=tên ký

            const range = XLSX.utils.decode_range(worksheet['!ref']);
            for (let r = 0; r <= range.e.r; r++) {
                for (let c = 0; c <= range.e.c; c++) {
                    const cell = XLSX.utils.encode_cell({ r, c });
                    if (!worksheet[cell]) continue;
                    worksheet[cell].s = {};
                    const s = worksheet[cell].s;

                    if (r === 0) {
                        // ĐẠI HỌC TRÀ VINH (không đậm) | CỘNG HÒA (đậm)
                        s.font = c >= 4 ? mkFont(13, true, false) : mkFont(13, false, false);
                        s.alignment = { horizontal: 'center', vertical: 'center' };
                    } else if (r === 1) {
                        s.font = c >= 4
                            ? { name: 'Times New Roman', sz: 13, bold: true, underline: true }
                            : mkFont(13, true, false);
                        s.alignment = { horizontal: 'center', vertical: 'center' };
                    } else if (r === 2) {
                        if (c < 4) {
                            // VÀ AN NINH ĐẠI HỌC TRÀ VINH (đậm + gạch chân)
                            s.font = { name: 'Arial', sz: 13, bold: true, underline: true };
                            s.alignment = { horizontal: 'center', vertical: 'center' };
                        } else {
                            // Ngày tháng năm (nghiêng, canh giữa)
                            s.font = mkFont(12, false, true);
                            s.alignment = { horizontal: 'center', vertical: 'center' };
                        }
                    } else if (r === 4) {
                        s.font = mkFont(14, true, false);
                        s.alignment = { horizontal: 'center', vertical: 'center' };
                    } else if (r === 5) {
                        // (Ban hành kèm theo...)
                        s.font = mkFont(12, false, false);
                        s.alignment = { horizontal: 'center', vertical: 'center' };
                    } else if (r === 8) {
                        // Header bảng
                        s.font = mkFont(11, true, false);
                        s.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
                        s.border = allBorders;
                    } else if (r >= 9 && r <= 8 + dataRowsCount) {
                        s.font = mkFont(11, false, false);
                        s.border = allBorders;
                        s.alignment = (c === 2 || c === 3 || c === 5)
                            ? { horizontal: 'left',   vertical: 'center' }
                            : { horizontal: 'center', vertical: 'center' };
                    } else if (r === 9 + dataRowsCount) {
                        // "Trên danh sách có X sinh viên" – đậm + nghiêng
                        s.font = mkFont(13, true, true);
                        s.alignment = { horizontal: 'left', vertical: 'center' };
                    } else {
                        s.font = mkFont(11, false, false);
                    }
                }
            }

            XLSX.utils.book_append_sheet(workbook, worksheet, 'QuyetDinh');
            const fileName = this.sanitizeFileName(`Danh_sach_cap_chung_chi_${report.subjectCode || 'MonHoc'}_${report.classCode || 'TatCaLop'}.xlsx`);
            XLSX.writeFile(workbook, fileName);
            this.showStatus('✓ Đã xuất file quyết định/chứng chỉ thành công', 'success', 'gradeStatus');
        } catch (error) {
            this.showStatus(`✗ Xuất file quyết định/chứng chỉ thất bại: ${error.message}`, 'error', 'gradeStatus');
        }
    }

    renderScoreSheet() {
        const content = document.getElementById('reportContent');
        const selectedSubject = document.getElementById('reportSubjectFilter').value;

        if (!selectedSubject) {
            content.innerHTML = '<p class="empty-message">Vui lòng chọn Mã Môn để tạo bảng ghi điểm</p>';
            return;
        }

        const selectedClass = document.getElementById('reportClassFilter').value;
        let report = this.manager.getScoreSheetData(selectedSubject, selectedClass);
        report = this.applySelectedClasses(report);
        report = this.applySelectedStudentCodes(report);
        const schoolOverride = this.getSchoolOverride();
        if (schoolOverride) {
            report = { ...report, schoolName: schoolOverride, schoolNameOverride: schoolOverride };
        }

        if (!report.rows.length) {
            content.innerHTML = '<p class="empty-message">Không có dữ liệu phù hợp với bộ lọc đã chọn</p>';
            return;
        }

        const displayClass = this.manager.formatClassDisplay(report.classCode, report.groupName);
        const displaySchoolOverride = String(report.schoolNameOverride || '').trim();
        const displaySchool = displaySchoolOverride || this.manager.getSchoolDisplayText(report.classCode, report.schoolName || report.majorName || '');
        const includeClassColumn = this.shouldShowClassColumn(report.subjectCode);

        const headerClassCol = includeClassColumn ? '<th>Mã lớp</th>' : '';

        let html = `
            <div class="report-meta">
                <p><strong>Trường:</strong> ${displaySchool || report.schoolName || report.majorName || '-'}</p>
                <p><strong>Nhóm/lớp:</strong> ${displayClass || 'Tất cả lớp'}</p>
                <p><strong>Môn học:</strong> ${this.manager.getSubjectDisplayText(report.subjectCode) || '-'}</p>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>TT</th>
                        <th>MSSV</th>
                        <th>HỌ VÀ</th>
                        <th>TÊN</th>
                        <th>NGÀY SINH</th>
                        <th>ĐIỂM TB</th>
                        <th>XẾP LOẠI</th>
                        <th>KẾT QUẢ</th>
                        ${headerClassCol}
                    </tr>
                </thead>
                <tbody>
        `;

        report.rows.forEach((row, index) => {
            const classCell = includeClassColumn ? `<td>${row.MaLop || ''}</td>` : '';
            html += `
                <tr>
                    <td>${index + 1}</td>
                    <td>${row.MaSV}</td>
                    <td>${row.HoVa || ''}</td>
                    <td>${row.Ten || ''}</td>
                    <td>${row.NgaySinhC || ''}</td>
                    <td>${this.formatScore(row.DiemTB)}</td>
                    <td>${row.XepLoai || ''}</td>
                    <td>${row.KetQua || ''}</td>
                    ${classCell}
                </tr>
            `;
        });

        html += `
                </tbody>
            </table>
            <div class="report-summary">
        `;

        const stats = this.manager.calculateStatistics(report.rows);
        html += `
                <span>Tổng số SV: ${stats.total}</span>
                <span>Số SV đạt: ${stats.passed}</span>
                <span>Cấm thi: ${stats.banned}</span>
                <span>Vắng thi: ${stats.absent}</span>
                <span>Hỏng: ${stats.failed}</span>
                <span>Đình chỉ: ${stats.suspended}</span>
                <span>Chưa học: ${stats.notStudied}</span>
            </div>
        `;

        content.innerHTML = html;
    }
}

document.addEventListener('DOMContentLoaded', () => {
    const params = new URLSearchParams(window.location.search);
    const isEmbedMode = params.get('embed') === '1';
    const isBldPage = /bangdiembld\.html$/i.test(window.location.pathname);
    const enablePersistence = !isEmbedMode && !isBldPage;

    const manager = new GradeManager({ persistEnabled: enablePersistence });
    const ui = new UIManager(manager);

    if (isEmbedMode) {
        document.body.classList.add('embed-mode');
    }

    window.addEventListener('message', (event) => {
        const data = event.data || {};
        if (data.type === 'gdqp-load-main' && data.file) {
            ui.importGradeFile(data.file);
        }
        if (data.type === 'gdqp-load-main-rows' && Array.isArray(data.rows)) {
            manager.importGradeData(data.rows);
            ui.showStatus(`✓ Đồng bộ bảng điểm thành công! (${data.rows.length} bản ghi)`, 'success', 'gradeStatus');
            ui.render();
        }
        if (data.type === 'gdqp-load-birth' && data.file) {
            ui.importBirthFile(data.file);
        }
        if (data.type === 'gdqp-load-birth-rows' && Array.isArray(data.rows)) {
            manager.importStudentInfo(data.rows);
            ui.showStatus(`✓ Đồng bộ Nơi sinh thành công! (${data.rows.length} bản ghi)`, 'success', 'studentStatus');
            ui.render();
        }
        if (data.type === 'gdqp-load-student' && data.file) {
            ui.importStudentFile(data.file);
        }
    });
});
