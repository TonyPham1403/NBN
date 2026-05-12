// Convert either 535.txt (TSV exported from sheet1) or data.xlsx -> data.json
// Usage: node convert_data_to_json.js
// Requires: npm install xlsx

const fs = require("fs");
const path = require("path");
let XLSX = null;
try { XLSX = require("xlsx"); } catch (e) { /* fallback if not needed */ }

// Look for 535.txt in several likely locations (proj/, parent folder, current working dir)
const txtCandidates = [
    path.join(__dirname, "535.txt"),
    path.join(__dirname, "..", "535.txt"),
    path.join(process.cwd(), "535.txt")
];
const txtPath = txtCandidates.find(p => fs.existsSync(p)) || path.join(__dirname, "535.txt");

// Look for .xlsm and .xlsx in several likely locations (proj/, parent folder, cwd)
const xlsmCandidates = [
    path.join(__dirname, "535.xlsm"),
    path.join(__dirname, "..", "535.xlsm"),
    path.join(process.cwd(), "535.xlsm"),
    path.join(__dirname, "data.xlsm"),
    path.join(__dirname, "..", "data.xlsm"),
    path.join(process.cwd(), "data.xlsm")
];
const xlsmPath = xlsmCandidates.find(p => fs.existsSync(p));

const xlsxCandidates = [
    path.join(__dirname, "535.xlsx"),
    path.join(__dirname, "..", "535.xlsx"),
    path.join(process.cwd(), "535.xlsx"),
    path.join(__dirname, "data.xlsx"),
    path.join(__dirname, "..", "data.xlsx"),
    path.join(process.cwd(), "data.xlsx")
];
const xlsxPath = xlsxCandidates.find(p => fs.existsSync(p));

const outputPath = path.join(__dirname, "data.json");

function parseIdNumber(idRaw) {
    const digits = String(idRaw || '').trim().match(/\d+$/);
    if (!digits) return null;
    const num = parseInt(digits[0], 10);
    return Number.isNaN(num) ? null : num;
}

function normalizeDateText(value) {
    const text = String(value || '').trim();
    const parts = text.split(/[\/\-.]/).map(p => p.trim()).filter(Boolean);
    if (parts.length !== 3) return '';
    const dd = parts[0].padStart(2, '0');
    const mm = parts[1].padStart(2, '0');
    const yyyy = parts[2].length === 2 ? `20${parts[2]}` : parts[2];
    if (!/^\d{2}$/.test(dd) || !/^\d{2}$/.test(mm) || !/^\d{4}$/.test(yyyy)) return '';
    return `${dd}/${mm}/${yyyy}`;
}

function addOneDay(dateText) {
    const normalized = normalizeDateText(dateText);
    if (!normalized) return '';
    const [dd, mm, yyyy] = normalized.split('/').map(n => parseInt(n, 10));
    const dt = new Date(Date.UTC(yyyy, mm - 1, dd));
    if (Number.isNaN(dt.getTime())) return '';
    dt.setUTCDate(dt.getUTCDate() + 1);
    const nextDd = String(dt.getUTCDate()).padStart(2, '0');
    const nextMm = String(dt.getUTCMonth() + 1).padStart(2, '0');
    const nextYyyy = String(dt.getUTCFullYear());
    return `${nextDd}/${nextMm}/${nextYyyy}`;
}

function computeTrailingBlankDate(rows) {
    if (!rows || rows.length === 0) return '';
    const last = rows[rows.length - 1] || {};
    const lastDate = normalizeDateText(last.date || last.Date || '');
    if (!lastDate) return '';
    if (rows.length < 2) return lastDate;

    const prev = rows[rows.length - 2] || {};
    const prevDate = normalizeDateText(prev.date || prev.Date || '');
    if (!prevDate) return lastDate;

    // Rule: if 2 previous records share same date => blank row date = +1 day, else keep last date.
    if (prevDate === lastDate) {
        return addOneDay(lastDate) || lastDate;
    }
    return lastDate;
}

function readExistingJson(pathJson) {
    if (!fs.existsSync(pathJson)) return null;
    try {
        const raw = fs.readFileSync(pathJson, 'utf8');
        const parsed = JSON.parse(raw);
        if (!Array.isArray(parsed)) return null;

        // Keep only meaningful rows (result exists); trailing editable row is regenerated later.
        const realRows = parsed
            .filter(row => {
                const result = String((row && row.result) || '').trim();
                return result.replace(/[,|\s]/g, '').length > 0;
            })
            .map(row => ({
                date: String((row && row.date) || '').trim(),
                id: String((row && row.id) || '').trim(),
                result: String((row && row.result) || '').trim()
            }));

        let lastIdNumber = null;
        for (let i = realRows.length - 1; i >= 0; i--) {
            const n = parseIdNumber(realRows[i] && realRows[i].id);
            if (n !== null) {
                lastIdNumber = n;
                break;
            }
        }

        return { realRows, lastIdNumber };
    } catch (e) {
        console.warn('⚠️  data.json hiện có không hợp lệ, sẽ build lại toàn bộ.');
        return null;
    }
}

function parseTxt(pathTxt) {
    const raw = fs.readFileSync(pathTxt, "utf8");
    const lines = raw.split(/\r?\n/).filter(l => l.trim().length > 0);
    if (lines.length === 0) return [];
    // detect delimiter (prefer tab)
    const header = lines[0];
    const delim = header.indexOf('\t') !== -1 ? '\t' : '\t';

    const out = [];
    // assume first line is header
    for (let i = 1; i < lines.length; i++) {
        const line = lines[i];
        // split into at most 5 columns; if more tabs exist, join extras into last
        const parts = line.split('\t');
        // sometimes the file may be space-separated; if only 1 part, try splitting by multiple spaces
        let cols = parts;
        if (parts.length === 1) cols = line.split(/\s{2,}/);
        // normalize to 5 columns
        while (cols.length < 5) cols.push('');
        if (cols.length > 5) {
            // join extras into the 5th column
            cols = cols.slice(0, 4).concat([cols.slice(4).join('\t')]);
        }
        const entry = {
            date: String(cols[0] || '').trim(),
            id: String(cols[1] || '').trim(),
            result: String(cols[2] || '').trim()
        };
        // treat rows without a meaningful result as skip
        if (entry.result && entry.result.replace(/[,|\s]/g, '').length > 0) out.push(entry);
    }
    return out;
}

function parseSpreadsheet(pathSpreadsheet, minIdExclusive = null) {
    if (!XLSX) throw new Error('xlsx package not installed');
    const workbook = XLSX.readFile(pathSpreadsheet, { cellDates: true });
    const sheet = workbook.Sheets['Sheet1'] || workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (!rows || rows.length === 0) return [];

    const out = [];
    // Row 1 in 535.xlsm is the label row, so start from row 2.
    // Stop immediately when the result column is blank.
    for (let i = 1; i < rows.length; i++) {
        const r = rows[i] || [];
        const cell = sheet[`A${i + 1}`] || {};
        const rawDate = r[0];
        const id = String(r[1] || '').trim();
        const result = String(r[2] || '').trim();

        let dateStr = '';
        if (cell && String(cell.w || '').trim()) {
            const display = String(cell.w).trim();
            const parts = display.split(/[\/\-.]/).map(part => part.trim()).filter(Boolean);
            if (parts.length === 3) {
                const [monthPart, dayPart, yearPart] = parts;
                const year = yearPart.length === 2 ? `20${yearPart}` : yearPart;
                dateStr = `${String(dayPart).padStart(2, '0')}/${String(monthPart).padStart(2, '0')}/${year}`;
            } else {
                dateStr = display;
            }
        } else if (rawDate instanceof Date) {
            const dd = String(rawDate.getUTCDate()).padStart(2, '0');
            const mm = String(rawDate.getUTCMonth() + 1).padStart(2, '0');
            const yyyy = rawDate.getUTCFullYear();
            dateStr = `${dd}/${mm}/${yyyy}`;
        } else {
            dateStr = String(rawDate || '').trim();
        }

        if (!result || result.replace(/[,|\s]/g, '').length === 0) break;

        if (minIdExclusive !== null) {
            const idNum = parseIdNumber(id);
            if (idNum !== null && idNum <= minIdExclusive) {
                continue;
            }
        }

        out.push({ date: dateStr, id, result });
    }
    return out;
}

function main() {
    let rows = [];
    const existingJson = readExistingJson(outputPath);
    if (xlsmPath) {
        console.log(`ℹ️  Found ${path.basename(xlsmPath)} — parsing Sheet1 columns A:E`);
        try {
            if (existingJson) {
                const newRows = parseSpreadsheet(xlsmPath, existingJson.lastIdNumber);
                rows = existingJson.realRows.concat(newRows);
                console.log(`ℹ️  Incremental mode: +${newRows.length} dòng mới`);
            } else {
                rows = parseSpreadsheet(xlsmPath);
            }
        } catch (e) {
            console.error('❌ Failed to parse XLSM:', e.message);
            process.exit(1);
        }
    } else if (xlsxPath) {
        console.log(`ℹ️  Found ${path.basename(xlsxPath)} — parsing first sheet`);
        try {
            if (existingJson) {
                const newRows = parseSpreadsheet(xlsxPath, existingJson.lastIdNumber);
                rows = existingJson.realRows.concat(newRows);
                console.log(`ℹ️  Incremental mode: +${newRows.length} dòng mới`);
            } else {
                rows = parseSpreadsheet(xlsxPath);
            }
        } catch (e) {
            console.error('❌ Failed to parse XLSX:', e.message);
            process.exit(1);
        }
    } else if (fs.existsSync(txtPath)) {
        console.log('ℹ️  Found 535.txt — parsing as TSV (first 5 columns)');
        rows = parseTxt(txtPath);
    } else {
        console.error('❌ Không tìm thấy 535.txt, 535.xlsm, hoặc data.xlsx');
        process.exit(1);
    }

    // Append a trailing blank record so the right pane can render the editable last row.
    // Keep the id sequence continuous by incrementing the last parsed id.
    let nextId = '';
    for (let i = rows.length - 1; i >= 0; i--) {
        const candidate = rows[i] || {};
        const idRaw = String(candidate.id || candidate.ID || '').trim();
        if (!idRaw) continue;

        const match = idRaw.match(/(\d+)$/);
        if (match) {
            const width = match[1].length;
            nextId = String(parseInt(match[1], 10) + 1).padStart(width, '0');
        } else {
            nextId = idRaw + '1';
        }
        break;
    }

    const nextDate = computeTrailingBlankDate(rows);
    rows.push({ date: nextDate, id: nextId, result: '' });

    fs.writeFileSync(outputPath, JSON.stringify(rows, null, 2), 'utf8');
    console.log(`✅ Đã tạo data.json (${rows.length} dòng)`);
}

main();
