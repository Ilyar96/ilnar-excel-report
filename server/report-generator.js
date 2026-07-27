const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx-js-style");
const express = require("express");
const multer = require("multer");
const cors = require("cors");

const INPUT_DIR = path.join(__dirname, "input");
const OUTPUT_DIR = path.join(__dirname, "output");
const OUTPUT_FILE = path.join(OUTPUT_DIR, "result.xlsx");
const SHOULD_ROUND_NUMBERS = true;
const SHOULD_FORMAT_NUMBERS = true;

const RESULT_COLUMNS = {
    department: 0,
    articleName: 1,
    contractor: 2,
    amountWithoutVat: 3,
    amountWithVat: 4,
    period: 5
};

/**
 * Чтение всех xlsx файлов
 */
function readInputFiles() {
    if (!fs.existsSync(INPUT_DIR)) {
        throw new Error("Папка input не найдена.");
    }

    return fs
        .readdirSync(INPUT_DIR)
        .filter(file => file.toLowerCase().endsWith(".xlsx"));
}

/**
 * Преобразование значения Excel в число
 */
function toNumber(value) {
    if (value === undefined || value === null || value === "") {
        return 0;
    }

    if (typeof value === "number") {
        return value;
    }

    return Number(
        String(value)
            .replace(/\s/g, "")
            .replace(",", ".")
    ) || 0;
}

function normalizeAmount(value) {
    const number = toNumber(value);

    if (!SHOULD_ROUND_NUMBERS) {
        return number;
    }

    return Math.round(number * 100) / 100;
}

function getNumericFormat() {
    if (SHOULD_ROUND_NUMBERS && SHOULD_FORMAT_NUMBERS) {
        return "# ##0.##";
    }

    if (SHOULD_ROUND_NUMBERS) {
        return "0.##";
    }

    if (SHOULD_FORMAT_NUMBERS) {
        return "# ##0.############";
    }

    return null;
}

function encodeCellAddress(rowIndex, columnIndex) {
    return XLSX.utils.encode_cell({
        r: rowIndex,
        c: columnIndex
    });
}

function setCellStyle(sheet, rowIndex, columnIndex, style) {
    const address = encodeCellAddress(rowIndex, columnIndex);

    if (!sheet[address]) {
        sheet[address] = {
            t: "s",
            v: ""
        };
    }

    sheet[address].s = style;
}

function getBorderStyle() {
    return {
        top: { style: "thin", color: { rgb: "000000" } },
        bottom: { style: "thin", color: { rgb: "000000" } },
        left: { style: "thin", color: { rgb: "000000" } },
        right: { style: "thin", color: { rgb: "000000" } }
    };
}

function applySheetStyles(sheet, rowCount, columnCount, totalRowIndex) {
    const numericFormat = getNumericFormat();
    const baseStyle = {
        border: getBorderStyle(),
        alignment: {
            horizontal: "right",
            vertical: "center"
        },
        font: {
            sz: 11
        }
    };

    const headerStyle = {
        ...baseStyle,
        font: {
            ...baseStyle.font,
            bold: true
        },
        alignment: {
            horizontal: "center",
            vertical: "center"
        }
    };

    const totalStyle = {
        font: {
            ...baseStyle.font,
            bold: true
        },
        alignment: {
            horizontal: "center",
            vertical: "center"
        },
        border: getBorderStyle()
    };

    for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
        for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
            const isHeaderRow = rowIndex === 0;
            const isTotalAmountCell =
                rowIndex === totalRowIndex &&
                columnIndex === RESULT_COLUMNS.amountWithVat;
            const isTotalRow = rowIndex === totalRowIndex;

            let style = baseStyle;

            if (isHeaderRow) {
                style = headerStyle;
            }

            if (isTotalAmountCell) {
                style = totalStyle;
            }

            if (isTotalRow && !isTotalAmountCell) {
                style = {
                    alignment: {
                        horizontal: "right",
                        vertical: "center"
                    }
                };
            }

            if (
                numericFormat &&
                rowIndex > 0 &&
                [
                    RESULT_COLUMNS.articleName,
                    RESULT_COLUMNS.amountWithoutVat,
                    RESULT_COLUMNS.amountWithVat
                ].includes(columnIndex)
            ) {
                style = {
                    ...style,
                    numFmt: numericFormat
                };
            }

            setCellStyle(sheet, rowIndex, columnIndex, style);
        }
    }
}

/**
 * Получение месяца из даты
 */
function getMonth(value) {
    if (!value) {
        return "";
    }

    if (value instanceof Date) {
        return value.getMonth() + 1;
    }

    if (typeof value === "number") {
        const date = XLSX.SSF.parse_date_code(value);

        if (date) {
            return date.m;
        }
    }

    const str = String(value).trim();

    let match = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);

    if (match) {
        return Number(match[2]);
    }

    match = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

    if (match) {
        return Number(match[1]);
    }

    const date = new Date(str);

    if (!Number.isNaN(date.getTime())) {
        return date.getMonth() + 1;
    }

    return "";
}

function getCurrentMonthNameRu() {
    const months = [
        "январь",
        "февраль",
        "март",
        "апрель",
        "май",
        "июнь",
        "июль",
        "август",
        "сентябрь",
        "октябрь",
        "ноябрь",
        "декабрь"
    ];

    const d = new Date();
    return months[d.getMonth()];
}

/**
 * Получение кода подразделения
 */
function extractDepartmentCode(objectName) {
    if (!objectName) {
        return null;
    }

    const match = String(objectName).trim().match(/^0*(\d+)/);

    return match ? String(Number(match[1])) : null;
}

/**
 * Нормализация заголовка для гибкого поиска
 */
function normalizeHeader(value) {
    return String(value ?? "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, " ")
        .replace(/\(руб\)/g, "")
        .replace(/[:.,]/g, "")
        .trim();
}

/**
 * Поиск строки заголовков
 */
function findHeaderRows(rows) {
    const indices = [];

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i].map(v => normalizeHeader(v));

        if (
            row.includes("объект") &&
            row.includes("дата закрытия заявки")
        ) {
            indices.push(i);
        }
    }

    return indices;
}

/**
 * Создание словаря "Название столбца -> индекс"
 */
function buildColumnMap(header) {
    const map = {};

    header.forEach((value, index) => {
        map[normalizeHeader(value)] = index;
    });

    return map;
}

/**
 * Поиск индекса колонки по вариантам названия
 */
function findColumnIndex(columns, candidates) {
    for (const candidate of candidates) {
        const index = columns[normalizeHeader(candidate)];

        if (index !== undefined) {
            return index;
        }
    }

    return -1;
}

/**
 * Получение итоговой суммы
 */
function getTotalAmount(rows, startRow, columns) {
    for (let i = startRow; i < rows.length; i++) {
        const row = rows[i];

        for (const cell of row) {
            if (
                String(cell ?? "")
                    .trim()
                    .toLowerCase()
                    .startsWith("итого")
            ) {
                for (let j = row.length - 1; j >= 0; j--) {
                    const num = toNumber(row[j]);

                    if (num !== 0) {
                        return num;
                    }
                }
            }
        }
    }

    let total = 0;

    const amountColumn =
        columns["сумма материала и услуги"];

    for (let i = startRow; i < rows.length; i++) {
        total += toNumber(rows[i][amountColumn]);
    }

    return total;
}

/**
 * Обработка одного файла
 */
function processFile(filePath) {
    const workbook = XLSX.readFile(filePath, {
        cellFormula: true,
        cellDates: true
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: ""
    });

    const headerRows = findHeaderRows(rows);

    if (headerRows.length === 0) {
        return [];
    }

    const fileMap = new Map();

    for (let h = 0; h < headerRows.length; h++) {
        const headerRow = headerRows[h];
        const nextHeaderRow = h + 1 < headerRows.length ? headerRows[h + 1] : rows.length;

        const columns = buildColumnMap(rows[headerRow]);

        const objectColumn = findColumnIndex(columns, ["объект"]);
        const dateColumn = findColumnIndex(columns, ["дата закрытия заявки"]);
        const amountColumn = findColumnIndex(columns, [
            "сумма материала и услуги",
            "сумма материала и услуги (руб)",
            "сумма материала и услуги(руб)"
        ]);

        if (objectColumn === -1 || dateColumn === -1 || amountColumn === -1) {
            console.warn(
                `Пропускаю таблицу, начиная с строки ${headerRow + 1}: не найдены обязательные колонки.`
            );
            continue;
        }

        let currentCode = null;
        let currentMonth = "";

        for (let i = headerRow + 1; i < nextHeaderRow; i++) {
            const row = rows[i];

            if (!row || row.every(cell => String(cell ?? "").trim() === "")) {
                continue;
            }

            const rowText = row.map(cell => String(cell ?? "").trim().toLowerCase());

            if (rowText.some(cell => cell.startsWith("итого"))) {
                continue;
            }

            const rowCode = extractDepartmentCode(row[objectColumn]);
            const rowMonth = getMonth(row[dateColumn]);

            if (rowCode) {
                currentCode = rowCode;
            }

            if (rowMonth) {
                currentMonth = rowMonth;
            }

            const code = rowCode || currentCode;

            if (!code) {
                continue;
            }

            const month = rowMonth || currentMonth || "";
            const amount = toNumber(row[amountColumn]);

            if (!fileMap.has(code)) {
                fileMap.set(code, {
                    code,
                    month: month || "",
                    total: 0
                });
            }

            const bucket = fileMap.get(code);

            if (!bucket.month && month) {
                bucket.month = month;
            }

            bucket.total += amount;
        }
    }

    return Array.from(fileMap.values());
}

/**
 * Формирование итогового массива
 */
function buildResultRows(resultMap, period) {
    const rows = [
        [
            "Подразделение",
            "Наименование статьи ДиР",
            "Контрагент",
            "Сумма начисления (без НДС)",
            "Сумма начисления (с НДС)",
            "Период расхода"
        ]
    ];
    let totalWithVatSum = 0;

    for (const item of resultMap.values()) {
        const totalWithVat = normalizeAmount(item.total);

        rows.push([
            item.code,
            totalWithVat,
            "",
            totalWithVat,
            totalWithVat,
            period || item.month || ""
        ]);
        totalWithVatSum += totalWithVat;
    }

    rows.push(["", "", "", "", normalizeAmount(totalWithVatSum), ""]);

    return rows;
}

/**
 * Запись Excel
 */
function writeResult(resultMap, options = {}) {
    if (!fs.existsSync(OUTPUT_DIR)) {
        fs.mkdirSync(OUTPUT_DIR);
    }

    const workbook = XLSX.utils.book_new();

    const rows = buildResultRows(resultMap, options.period);

    const sheet = XLSX.utils.aoa_to_sheet(rows);
    for (let i = 2; i < rows.length; i++) {
        const totalWithVat = normalizeAmount(rows[i - 1][RESULT_COLUMNS.amountWithVat]);

        sheet[`D${i}`] = {
            t: "n",
            v: normalizeAmount(totalWithVat / 1.2),
            f: `E${i}/1.2`
        };
    }

    sheet[`E${rows.length}`] = {
        t: "n",
        v: normalizeAmount(rows[rows.length - 1][RESULT_COLUMNS.amountWithVat])
    };

    // apply font sizes if provided
    const titleFontSize = options.titleFontSize ?? 11;
    const cellFontSize = options.cellFontSize ?? 11;

    // update default font sizes used in styles
    // applySheetStyles uses default of 11; we need to set sizes in cell styles after creation
    applySheetStyles(sheet, rows.length, rows[0].length, rows.length - 1);

    // walk through cells and set font sizes
    for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < rows[0].length; c++) {
            const addr = encodeCellAddress(r, c);
            if (!sheet[addr]) continue;
            sheet[addr].s = sheet[addr].s || {};
            sheet[addr].s.font = sheet[addr].s.font || {};
            if (r === 0) {
                sheet[addr].s.font.sz = titleFontSize;
            } else {
                sheet[addr].s.font.sz = cellFontSize;
            }
        }
    }
    sheet["!cols"] = [
        { wch: 18 },
        { wch: 30 },
        { wch: 24 },
        { wch: 26 },
        { wch: 24 },
        { wch: 18 }
    ];

    XLSX.utils.book_append_sheet(workbook, sheet, "Result");

    let outputName = options.resultName;
    if (!outputName) {
        outputName = `Акруалы за ${getCurrentMonthNameRu()}`;
    }

    if (!outputName.toLowerCase().endsWith(".xlsx")) {
        outputName += ".xlsx";
    }

    const outputPath = path.join(OUTPUT_DIR, outputName);

    XLSX.writeFile(workbook, outputPath);

    return outputPath;
}

function deleteAllFilesInDir(dir) {
    if (!fs.existsSync(dir)) {
        return [];
    }

    const deleted = [];
    for (const name of fs.readdirSync(dir)) {
        const filePath = path.join(dir, name);
        try {
            if (fs.lstatSync(filePath).isFile()) {
                fs.unlinkSync(filePath);
                deleted.push(name);
            }
        } catch (err) {
            console.error(`Не удалось удалить файл ${filePath}:`, err);
        }
    }
    return deleted;
}

/**
 * Главная функция
 */
// --- Express API ---
const app = express();
app.use(cors());
app.use(express.json());

// ensure input/output dirs
if (!fs.existsSync(INPUT_DIR)) fs.mkdirSync(INPUT_DIR);
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR);

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, INPUT_DIR);
    },
    filename: function (req, file, cb) {
        cb(null, path.basename(file.originalname));
    }
});

const upload = multer({ storage });

app.get("/files", (req, res) => {
    try {
        const files = readInputFiles();
        res.json(files);
    } catch (err) {
        res.status(500).send(err.message);
    }
});

app.post("/files", upload.array("files"), (req, res) => {
    const saved = [];
    for (const f of req.files || []) {
        if (!f.originalname.toLowerCase().endsWith('.xlsx')) {
            // remove invalid
            fs.unlinkSync(f.path);
            continue;
        }
        saved.push(path.basename(f.originalname));
    }

    res.json({ saved });
});

app.delete("/files/:name", (req, res) => {
    const name = path.basename(req.params.name);
    if (!name.toLowerCase().endsWith('.xlsx')) {
        return res.status(400).send('Неверное имя файла');
    }

    const p = path.join(INPUT_DIR, name);
    if (!fs.existsSync(p)) return res.status(404).send('Файл не найден');
    try {
        fs.unlinkSync(p);
        res.sendStatus(204);
    } catch (err) {
        res.status(500).send(err.message);
    }
});

app.post("/generate", async (req, res) => {
    const { files, resultName, titleFontSize, cellFontSize } = req.body || {};

    let toProcess = [];

    try {
        const all = readInputFiles();

        if (files && Array.isArray(files) && files.length > 0) {
            for (const f of files) {
                const b = path.basename(f);
                if (!all.includes(b)) {
                    return res.status(400).send(`Файл ${b} не найден в input`);
                }
                toProcess.push(b);
            }
        } else {
            toProcess = all;
        }

        const resultMap = new Map();

        for (const file of toProcess) {
            const filePath = path.join(INPUT_DIR, file);
            const fileRows = processFile(filePath);

            for (const result of fileRows) {
                if (!resultMap.has(result.code)) {
                    resultMap.set(result.code, {
                        code: result.code,
                        month: result.month,
                        total: 0
                    });
                }

                const bucket = resultMap.get(result.code);

                if (!bucket.month && result.month) {
                    bucket.month = result.month;
                }

                bucket.total += result.total;
            }
        }

let periodValue = undefined;
    if (typeof req.body?.period !== 'undefined' && req.body.period !== null) {
        periodValue = Number(req.body.period);
        if (!Number.isInteger(periodValue) || periodValue < 1 || periodValue > 12) {
            return res.status(400).send('Период расхода должен быть целым числом от 1 до 12.');
        }
    }

    const outputPath = writeResult(resultMap, {
        resultName,
        titleFontSize: Number(titleFontSize) || 11,
        cellFontSize: Number(cellFontSize) || 11,
        period: periodValue
    });

        const filename = path.basename(outputPath);

        res.download(outputPath, filename, err => {
            if (err) {
                console.error('Ошибка отправки файла:', err);
                return res.status(500).send(err.message);
            }

            // on success remove input files used and delete generated output
            for (const f of toProcess) {
                try {
                    fs.unlinkSync(path.join(INPUT_DIR, f));
                } catch (e) {
                    // ignore
                }
            }

            try {
                fs.unlinkSync(outputPath);
            } catch (e) {
                console.error('Не удалось удалить сгенерированный файл после отправки:', e);
            }
        });
    } catch (err) {
        console.error(err);
        res.status(500).send(String(err.message || err));
    }
});

app.get('/result/:name', (req, res) => {
    const name = path.basename(req.params.name);
    const p = path.join(OUTPUT_DIR, name);
    if (!fs.existsSync(p)) return res.status(404).send('Файл не найден');
    res.download(p, name);
});

app.delete('/result/:name', (req, res) => {
    const name = path.basename(req.params.name);
    const p = path.join(OUTPUT_DIR, name);
    if (!fs.existsSync(p)) return res.status(404).send('Файл не найден');
    try {
        fs.unlinkSync(p);
        res.sendStatus(204);
    } catch (err) {
        res.status(500).send(err.message);
    }
});

app.delete('/cleanup', (req, res) => {
    const deletedInput = deleteAllFilesInDir(INPUT_DIR);
    const deletedOutput = deleteAllFilesInDir(OUTPUT_DIR);

    res.json({
        deletedInput,
        deletedOutput
    });
});

app.post('/result/cleanup', (req, res) => {
    const { name } = req.body || {};
    if (!name) {
        return res.sendStatus(204);
    }

    const safeName = path.basename(String(name));
    const p = path.join(OUTPUT_DIR, safeName);

    if (fs.existsSync(p)) {
        try {
            fs.unlinkSync(p);
        } catch (err) {
            console.error('Не удалось удалить результат при очистке:', err);
        }
    }

    res.sendStatus(204);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Report generator server listening on http://localhost:${PORT}`);
});