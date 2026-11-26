require("dotenv").config();

const fs = require("fs");
const path = require("path");
const express = require("express");
const helmet = require("helmet");
const multer = require("multer");
const ExcelJS = require("exceljs");
const { rateLimit } = require("express-rate-limit");

const app = express();
const PORT = Number(process.env.PORT || 3000);
const ROOT = __dirname;
const BACKUP_DIR = path.join(ROOT, "backups");
fs.mkdirSync(BACKUP_DIR, { recursive: true });

app.disable("x-powered-by");
app.set("trust proxy", 1);
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      scriptSrc: ["'self'", "https://cdn.jsdelivr.net", "https://fonts.googleapis.com"],
      styleSrc: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com"],
      fontSrc: ["'self'", "https://fonts.gstatic.com", "data:"],
      imgSrc: ["'self'", "data:", "https:"],
      connectSrc: ["'self'"]
    }
  }
}));
app.use(express.json({ limit: "1mb" }));
app.use("/api/", rateLimit({ windowMs: 60 * 1000, limit: 90, standardHeaders: "draft-8", legacyHeaders: false }));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024, files: 1 }
});

function norm(value) {
  return String(value ?? "").trim().toUpperCase();
}

function headerIndex(headers, wanted) {
  const target = norm(wanted).replace(/[^A-Z0-9]/g, "");
  return headers.findIndex(h => norm(h).replace(/[^A-Z0-9]/g, "") === target);
}

function readSheetRows(worksheet) {
  if (!worksheet || worksheet.rowCount < 1) throw new Error("El Excel no contiene una hoja con encabezados.");
  const headerRow = worksheet.getRow(1);
  const headers = [];
  for (let c = 1; c <= Math.max(headerRow.cellCount, worksheet.actualColumnCount); c++) {
    headers.push(String(headerRow.getCell(c).text || headerRow.getCell(c).value || "").trim());
  }
  while (headers.length && !headers[headers.length - 1]) headers.pop();
  if (headers.length !== 41) {
    throw new Error(`Se esperaban exactamente 41 columnas en el Excel; se encontraron ${headers.length}.`);
  }
  const rows = [];
  for (let r = 2; r <= worksheet.rowCount; r++) {
    const row = worksheet.getRow(r);
    const values = headers.map((_, i) => row.getCell(i + 1).value ?? null);
    if (values.every(v => v === null || v === "")) continue;
    rows.push({ excelRow: r, values });
  }
  return { headers, rows };
}

function rowKey(values, supplierIdx, allocationIdx) {
  const supplier = norm(values[supplierIdx]);
  if (supplier) return `SUPPLIER:${supplier}`;
  const allocation = norm(values[allocationIdx]);
  return allocation ? `ALLOCATION:${allocation}` : "";
}

function compareAndApply(destinationWs, sourceData) {
  const destination = readSheetRows(destinationWs);
  const source = sourceData;
  const normalizedDestination = destination.headers.map(h => norm(h).replace(/[^A-Z0-9]/g, ""));
  const normalizedSource = source.headers.map(h => norm(h).replace(/[^A-Z0-9]/g, ""));
  if (normalizedDestination.some((h, i) => h !== normalizedSource[i])) {
    throw new Error("Los encabezados o el orden de las 41 columnas no coinciden entre ambos Excel.");
  }

  const supplierIdx = headerIndex(source.headers, "Supplier ID");
  const allocationIdx = headerIndex(source.headers, "Allocation Number");
  const actionIdx = headerIndex(source.headers, "Action Code");
  if ([supplierIdx, allocationIdx, actionIdx].some(i => i < 0)) {
    throw new Error("No se encontraron los encabezados Supplier ID, Allocation Number y Action Code en el Excel base.");
  }

  const destinationMap = new Map();
  for (const item of destination.rows) {
    const key = rowKey(item.values, supplierIdx, allocationIdx);
    if (!key) continue;
    if (destinationMap.has(key)) {
      throw new Error(`El Excel de OneDrive contiene una clave duplicada (${key}). Se canceló la sincronización.`);
    }
    destinationMap.set(key, item.excelRow);
  }

  const sourceKeys = new Set();
  const changes = [];
  let added = 0, updated = 0, deleted = 0, unchanged = 0, skipped = 0;

  for (const item of source.rows) {
    const values = item.values;
    const key = rowKey(values, supplierIdx, allocationIdx);
    const action = norm(values[actionIdx]).toLowerCase();
    if (!key) {
      if (action === "deleven" || action === "sysdel") {
        throw new Error(`La fila ${item.excelRow} tiene acción ${action} pero no contiene Supplier ID ni Allocation Number.`);
      }
      skipped++;
      continue;
    }
    if (sourceKeys.has(key)) {
      throw new Error(`El Excel base contiene una clave duplicada (${key}). Se canceló la sincronización.`);
    }
    sourceKeys.add(key);

    const targetRowNumber = destinationMap.get(key);
    if (action === "deleven" || action === "sysdel") {
      if (targetRowNumber) {
        destinationWs.spliceRows(targetRowNumber, 1);
        deleted++;
        changes.push({ action: "Eliminado", key, sourceRow: item.excelRow, fields: [] });
        for (const [mapKey, rowNum] of destinationMap.entries()) {
          if (rowNum > targetRowNumber) destinationMap.set(mapKey, rowNum - 1);
        }
        destinationMap.delete(key);
      } else {
        unchanged++;
      }
      continue;
    }

    if (!targetRowNumber) {
      const addedRow = destinationWs.addRow(values);
      addedRow.commit?.();
      added++;
      destinationMap.set(key, addedRow.number);
      changes.push({ action: "Agregado", key, sourceRow: item.excelRow, fields: [] });
      continue;
    }

    const targetRow = destinationWs.getRow(targetRowNumber);
    const changedFields = [];
    for (let c = 0; c < source.headers.length; c++) {
      const oldValue = targetRow.getCell(c + 1).value ?? null;
      const newValue = values[c] ?? null;
      const oldCompare = oldValue instanceof Date ? oldValue.toISOString() : JSON.stringify(oldValue);
      const newCompare = newValue instanceof Date ? newValue.toISOString() : JSON.stringify(newValue);
      if (oldCompare !== newCompare) {
        targetRow.getCell(c + 1).value = newValue;
        changedFields.push(source.headers[c]);
      }
    }
    if (changedFields.length) {
      targetRow.commit?.();
      updated++;
      changes.push({ action: "Actualizado", key, sourceRow: item.excelRow, fields: changedFields });
    } else {
      unchanged++;
    }
  }

  return { added, updated, deleted, unchanged, skipped, changes, totalBase: source.rows.length, totalDestino: destination.rows.length };
}

function buildOneDriveDownloadCandidates(rawUrl) {
  const cleanUrl = String(rawUrl || "").trim();
  let parsed;
  try {
    parsed = new URL(cleanUrl);
  } catch {
    throw new Error("El enlace de OneDrive no tiene un formato de URL válido.");
  }
  if (!["http:", "https:"].includes(parsed.protocol)) {
    throw new Error("El enlace debe comenzar con https://");
  }

  const candidates = [];

  // 1. Endpoint oficial de Graph para enlaces compartidos codificados en base64url (u!...)
  const encodedShare = Buffer.from(cleanUrl, "utf8")
    .toString("base64")
    .replace(/=+$/g, "")
    .replace(/\//g, "_")
    .replace(/\+/g, "-");
  candidates.push(`https://api.onedrive.com/v1.0/shares/u!${encodedShare}/root/content`);

  // 2. Parámetro download=1 en enlaces 1drv.ms / onedrive.live.com / sharepoint
  const withDownload = new URL(cleanUrl);
  withDownload.searchParams.set("download", "1");
  candidates.push(withDownload.toString());

  // 3. Conversión clásica de onedrive.live.com/redir -> /download
  if (parsed.hostname.includes("onedrive.live.com") && parsed.pathname.includes("/redir")) {
    const redir = new URL(cleanUrl);
    redir.pathname = redir.pathname.replace("/redir", "/download");
    candidates.push(redir.toString());
  }

  // 4. URL original tal cual
  candidates.push(cleanUrl);

  return [...new Set(candidates)];
}

async function fetchExcelFromOneDriveLink(rawUrl) {
  const candidates = buildOneDriveDownloadCandidates(rawUrl);
  let lastError = "No se pudo descargar el archivo desde el enlace.";

  for (const url of candidates) {
    try {
      const response = await fetch(url, {
        redirect: "follow",
        headers: {
          "User-Agent": "Mozilla/5.0 (QA-MANTO-Sync)",
          "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*"
        }
      });

      if (!response.ok) {
        lastError = `OneDrive respondió con estado ${response.status}. Verifica que el enlace sea público ("Cualquier persona con el vínculo").`;
        continue;
      }

      const contentType = String(response.headers.get("content-type") || "").toLowerCase();
      const buffer = Buffer.from(await response.arrayBuffer());

      // Verificar firma ZIP (PK\x03\x04) que corresponde a un archivo .xlsx válido y no a una página HTML de login
      const isXlsxZip = buffer.length > 4 && buffer[0] === 0x50 && buffer[1] === 0x4b && buffer[2] === 0x03 && buffer[3] === 0x04;
      if (!isXlsxZip || contentType.includes("text/html")) {
        lastError = "El enlace devolvió una página web en lugar del archivo .xlsx. Asegúrate de compartirlo como 'Cualquier persona que tenga el vínculo'.";
        continue;
      }

      // Extraer nombre de archivo si viene en Content-Disposition
      let fileName = "Proveedor_OneDrive.xlsx";
      const disp = response.headers.get("content-disposition") || "";
      const matchStar = disp.match(/filename\*=UTF-8''([^;]+)/i);
      const matchSimple = disp.match(/filename="?([^";]+)"?/i);
      if (matchStar) fileName = decodeURIComponent(matchStar[1]);
      else if (matchSimple) fileName = matchSimple[1];
      if (!/\.xlsx$/i.test(fileName)) fileName += ".xlsx";

      return { buffer, fileName };
    } catch (err) {
      lastError = err.message;
    }
  }

  throw new Error(lastError);
}

app.get("/api/proveedor/config", (req, res) => {
  res.json({
    defaultOneDriveUrl: process.env.ONEDRIVE_EXCEL_URL || "",
    maxUploadMB: 25
  });
});

app.post("/api/proveedor/test-link", async (req, res, next) => {
  try {
    const link = String(req.body?.oneDriveUrl || process.env.ONEDRIVE_EXCEL_URL || "").trim();
    if (!link) return res.status(400).json({ error: "Pega el enlace compartido de OneDrive." });

    const { buffer, fileName } = await fetchExcelFromOneDriveLink(link);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];
    const parsed = readSheetRows(ws);

    res.json({
      ok: true,
      fileName,
      rowsCount: parsed.rows.length,
      columnsCount: parsed.headers.length
    });
  } catch (error) { next(error); }
});

app.post("/api/proveedor/sync-link", upload.single("baseExcel"), async (req, res, next) => {
  try {
    if (!req.file) return res.status(400).json({ error: "Selecciona el Excel base desde tu computadora." });
    if (!/\.xlsx$/i.test(req.file.originalname)) {
      return res.status(400).json({ error: "El Excel base debe estar en formato .xlsx." });
    }

    const link = String(req.body?.oneDriveUrl || process.env.ONEDRIVE_EXCEL_URL || "").trim();
    if (!link) return res.status(400).json({ error: "Ingresa el enlace compartido del Excel en OneDrive." });

    // 1. Leer Excel base local
    const sourceWb = new ExcelJS.Workbook();
    await sourceWb.xlsx.load(req.file.buffer);
    const sourceWs = sourceWb.worksheets[0];
    const sourceData = readSheetRows(sourceWs);

    // 2. Descargar Excel destino directamente desde el enlace compartido de OneDrive
    const { buffer: destinationBuffer, fileName } = await fetchExcelFromOneDriveLink(link);
    const destinationWb = new ExcelJS.Workbook();
    await destinationWb.xlsx.load(destinationBuffer);
    const destinationWs = destinationWb.worksheets[0];

    // 3. Comparar y aplicar reglas
    const report = compareAndApply(destinationWs, sourceData);

    // Guardar copia de respaldo local en carpeta backups/
    const safeName = fileName.replace(/[\\/:*?"<>|]/g, "_");
    const backupName = `${safeName.replace(/\.xlsx$/i, "")}.backup-${new Date().toISOString().replace(/[:.]/g, "-")}.xlsx`;
    fs.writeFileSync(path.join(BACKUP_DIR, backupName), destinationBuffer);

    const updatedBuffer = Buffer.from(await destinationWb.xlsx.writeBuffer());
    const hasChanges = (report.added + report.updated + report.deleted) > 0;

    res.json({
      ...report,
      fileName: safeName,
      backupName,
      completedAt: new Date().toISOString(),
      message: hasChanges
        ? "Sincronización completada. Se descargó el Excel actualizado listo para reemplazar en OneDrive."
        : "No se detectaron cambios entre el Excel base y el archivo de OneDrive.",
      updatedFileBase64: hasChanges ? updatedBuffer.toString("base64") : null
    });
  } catch (error) { next(error); }
});

app.use(express.static(ROOT));

app.get(/.*/, (req, res, next) => {
  if (req.path.startsWith("/api/")) return next();
  res.sendFile(path.join(ROOT, "index.html"));
});

app.use((error, req, res, next) => {
  console.error("[QA MANTO]", error);
  const status = Number(error.status || 500);
  res.status(status).json({
    error: error.message || "Ocurrió un error interno."
  });
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`QA MANTO escuchando en el puerto ${PORT}`);
});
