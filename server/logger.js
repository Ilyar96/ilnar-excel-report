const fs = require("fs");
const path = require("path");

const LOG_DIR = path.join(__dirname, "logs");
const RETENTION_DAYS = 7;
const RETENTION_MS = RETENTION_DAYS * 24 * 60 * 60 * 1000;

function ensureLogDir() {
    if (!fs.existsSync(LOG_DIR)) {
        fs.mkdirSync(LOG_DIR, { recursive: true });
    }
}

function formatMeta(meta) {
    if (!meta || Object.keys(meta).length === 0) {
        return "";
    }

    try {
        return " " + JSON.stringify(meta);
    } catch (err) {
        return " [meta serialization failed]";
    }
}

function formatMessage(level, message, meta) {
    const timestamp = new Date().toISOString();
    const normalizedMessage = typeof message === "string" ? message : String(message);
    return `${timestamp} [${level}] ${normalizedMessage}${formatMeta(meta)}\n`;
}

function getLogFilePath() {
    const fileName = `${new Date().toISOString().slice(0, 10)}.log`;
    return path.join(LOG_DIR, fileName);
}

function writeSeparator() {
    try {
        ensureLogDir();
        fs.appendFileSync(getLogFilePath(), "==============================\n", "utf8");
    } catch (err) {
        console.error("Logger separator failed:", err);
    }
}

function writeLog(level, message, meta) {
    try {
        ensureLogDir();
        const entry = formatMessage(level, message, meta);
        fs.appendFileSync(getLogFilePath(), entry, "utf8");
    } catch (err) {
        // don't throw from logger
        console.error("Logger failed:", err);
    }
}

function cleanOldLogs() {
    try {
        if (!fs.existsSync(LOG_DIR)) {
            return;
        }

        const threshold = Date.now() - RETENTION_MS;
        for (const fileName of fs.readdirSync(LOG_DIR)) {
            const filePath = path.join(LOG_DIR, fileName);
            try {
                const stat = fs.statSync(filePath);
                if (stat.isFile() && stat.mtimeMs < threshold && path.extname(fileName) === ".log") {
                    fs.unlinkSync(filePath);
                }
            } catch (err) {
                console.error(`Failed to cleanup log file ${filePath}:`, err);
            }
        }
    } catch (err) {
        console.error("Failed to cleanup logs directory:", err);
    }
}

cleanOldLogs();
writeSeparator();
writeLog("INFO", "=== Новый запуск логгера ===", { action: "loggerStart" });

module.exports = {
    info: (message, meta) => writeLog("INFO", message, meta),
    warn: (message, meta) => writeLog("WARN", message, meta),
    error: (message, meta) => writeLog("ERROR", message, meta),
    debug: (message, meta) => writeLog("DEBUG", message, meta),
    cleanOldLogs
};
