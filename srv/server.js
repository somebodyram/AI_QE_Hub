require('dotenv').config();
const cds = require("@sap/cds");
const ExcelJS = require("exceljs");
const multer = require("multer");
const express = require("express");
const hdb = require('hdb');        
const crypto = require('crypto');  
const { initCData, queryData, cdataRequest } = require("./cdata");
const upload = multer();

// ─── HELPER FUNCTIONS ────────────────────────────────────────────────────────
function extractOrderNumber(text) {
  if (!text) return null;
  const strText = String(text);

  // PRIORITY 1: Always hunt for the AB series (Cancellation) first
  // If an AB number exists anywhere in the text, it grabs this one and stops.
  const cancelMatch = strText.match(/\bAB\d{10,}[A-Z0-9]{0,5}\b/);
  if (cancelMatch) {
    return cancelMatch[0];
  }

  // PRIORITY 2: Fallback to the standard AD/AE series
  const match = strText.match(/\b[A-Z]{2}\d{10,}[A-Z0-9]{0,5}\b/);
  return match ? match[0] : null;
}

function isInitialPurchaseSheet(sheetName) {
  if (!sheetName) return false;
  const name = String(sheetName).trim().toLowerCase();
  return name.includes("initial purchase") || name.includes("ip -") || name.includes("ip-") || /\bip\b/.test(name);
}

// function getCellText(rawValue) {
//   if (rawValue === null || rawValue === undefined) return null;
//   if (typeof rawValue === "object") {
//     if (rawValue.richText) return rawValue.richText.map(r => r.text).join("");
//     if (rawValue.result !== undefined) return String(rawValue.result);
//     if (rawValue.text !== undefined) return String(rawValue.text);       // hyperlink cell { text, hyperlink }
//     if (rawValue.hyperlink !== undefined) return String(rawValue.hyperlink);
//     return null; // unknown object shape — avoid leaking "[object Object]"
//   }
//   return String(rawValue);
// }

// Extract exactly 10 digits, ignoring emails and connected text
function extractEccContract(text) {
  if (!text) return null;
  // (?<!...) ensures NO letter, number, +, @, _, or - is BEFORE the 10 digits
  // (?!=...) ensures NO letter, number, +, @, _, or - is AFTER the 10 digits
  const match = String(text).match(/(?<![A-Za-z0-9+@_\-])\d{10}(?![A-Za-z0-9+@_\-])/);
  return match ? match[0] : null;
}

function getCellText(rawValue) {
  if (rawValue === null || rawValue === undefined) return null;
  if (typeof rawValue === "object") {
    if (rawValue.richText) return rawValue.richText.map(r => r.text).join("");
    if (rawValue.result !== undefined) return String(rawValue.result);
    if (rawValue.text !== undefined) return String(rawValue.text);
    if (rawValue.hyperlink !== undefined) return String(rawValue.hyperlink);
    return null;
  }
  return String(rawValue);
}

// function resolveSharePointFileUrl(url, folderPathOverride) {
//   try {
//     const parsed = new URL(url);
//     if (parsed.pathname.includes("/_layouts/")) {
//       const fileName = parsed.searchParams.get("file");
//       if (fileName) {
//         const folderPath = folderPathOverride || process.env.SP_FOLDER_PATH || "ABS - AI AGENT";
//         return `${folderPath}/${fileName}`;
//       }
//     }
//     return url;
//   } catch {
//     return url;
//   }
// }

cds.on("bootstrap", (app) => {

  app.use(express.json());

  /* ================= 1. READ EXCEL (LOCAL UPLOAD) ================= */
  app.post("/excel/read", upload.single("file"), async (req, res) => {
    try {
      const actionType = req.body.action || "INITIAL_PURCHASE";
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(req.file.buffer);
      const result = [];
      
      // Extract file name and remove .xlsx extension
      const rawFileName = req.file.originalname || "Local_Upload";
      const cleanFileName = rawFileName.replace(/\.[^/.]+$/, "");

      workbook.worksheets.forEach(ws => {
        const headerRow = ws.getRow(1);
        let cols = { scenario: 2, country: null, actualResult: null };
        headerRow.eachCell((cell, colNumber) => {
        const headerName = (getCellText(cell.value) || "").trim().toLowerCase();
        if (headerName === "scenario" && cols.scenario === 2) cols.scenario = colNumber;
        if (headerName === "country" && !cols.country) cols.country = colNumber;
        if (["actual result", "adobe id", "result"].includes(headerName) && !cols.actualResult) {
          cols.actualResult = colNumber;
        }
      });

        // --- ROW PROCESSING ---
        ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (rowNumber <= 1) return; // skip header
          
          const rawScenario = row.getCell(cols.scenario).value;
          const scenario = rawScenario ? String(rawScenario).trim().toLowerCase() : "";

          const IGNORE_TEXTS = "VIP, 3ds, change plan, switch plan"; 
          const excludeList = IGNORE_TEXTS.split(',').map(s => s.trim().toLowerCase());
          if (excludeList.some(kw => scenario.includes(kw))) return;

          const IP_KEYWORDS = ["initial purchase", "ip -", "ip-", "ip " , " cancel ", "cancellation", "cancel plan"]; 
          const matchesKeyword = IP_KEYWORDS.some(kw => scenario.includes(kw));
          const matchesSheet = isInitialPurchaseSheet(ws.name);

          if (matchesKeyword || matchesSheet) {
            let actualResultText = null;
          if (cols.actualResult) {
  //           if (ws.name.includes("TwP")) {
  // headerRow.eachCell((cell, colNumber) => {
  //   console.log(`[DEBUG-HEADERS] TwP col=${colNumber} value=`, JSON.stringify(getCellText(cell.value)));
  // });
// }
            const rawActualResult = row.getCell(cols.actualResult).value;
            actualResultText = getCellText(rawActualResult);
              // ? (typeof rawActualResult === 'object' && rawActualResult.richText)
              //   ? rawActualResult.richText.map(r => r.text).join('')
              //   : (typeof rawActualResult === 'object' && rawActualResult.result !== undefined)
              //     ? String(rawActualResult.result)
              //     : String(rawActualResult)
              // : null;
          }

          const extractedOrderNumber = extractOrderNumber(actualResultText);
          if (!extractedOrderNumber) return;

            result.push({
              scenario:     row.getCell(cols.scenario).value,
            country:      cols.country ? row.getCell(cols.country).value : null,
            actualResult: actualResultText,
            orderNumber:  extractedOrderNumber,
            eccContract:  extractEccContract(actualResultText), 
              sheet:        ws.name,
              rowNumber:    rowNumber,
              fileName:     cleanFileName, 
              runStatus:    "", 
              sapDb:        "", 
              sourceData:   "", 
              sapOutbound:  ""  
            });
          }
        });
      });
      res.json(result);
    } catch (err) {
      console.error("Read Excel error:", err); 
      res.status(500).json({ error: err.message });
    }
  });

  /* ================= 4. WRITE BACK TO LOCAL UPLOADED EXCEL ================= */
  app.post("/excel/writeback", upload.single("file"), async (req, res) => {
    try {
      const updates = JSON.parse(req.body.updates || "[]");
      if (!req.file || !updates.length) {
        return res.status(400).json({ error: "File data and table row updates are required" });
      }

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(req.file.buffer);

      // Create a rapid lookup map: { "SheetName:rowNumber": { status, reason } }
      const updateMap = {};
      updates.forEach(u => {
        updateMap[`${u.sheet}:${u.rowNumber}`] = {
          status: u.status || "",
          reason: u.reason || ""
        };
      });

      // Loop through the uploaded worksheets and update the modified rows
      workbook.worksheets.forEach(ws => {
        const headerRow = ws.getRow(1);
        let statusCol = null;
        let reasonCol = null;

        // Dynamically find position indexes of target columns
        headerRow.eachCell((cell, colNumber) => {
          const val = String(cell.value || "").trim().toLowerCase();
          if (val === "status")  statusCol = colNumber;
          if (val === "remarks" || val === "reason") reasonCol = colNumber;
        });

        if (!statusCol && !reasonCol) return; // Skip if worksheet lacks core columns

        ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (rowNumber <= 1) return; // Skip header row
          
          const key = `${ws.name}:${rowNumber}`;
          if (!updateMap[key]) return;

          const { status, reason } = updateMap[key];
          
          if (statusCol && status) {
            row.getCell(statusCol).value = status;
          }
          if (reasonCol && reason) {
            row.getCell(reasonCol).value = reason;
          }
          row.commit();
        });
      });

      // Compile Workbook directly back into an OpenXML buffer array
      const outBuffer = await workbook.xlsx.writeBuffer();
      
      // Send download stream headers back down to browser client
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", "attachment; filename=updated_execution.xlsx");
      res.send(outBuffer);

    } catch (err) {
      console.error("Local file writeback processing failed:", err);
      res.status(500).json({ error: err.message });
    }
  });


  /* =========================================================== */
  /* SHAREPOINT INTEGRATION (CDATA)                              */
  /* =========================================================== */
  const IP_KEYWORDS = ["initial purchase", "ip -", "ip-", "ip " , " cancel ", "cancellation", "cancel plan"];
  const CATALOG     = "ConsumerCommerce_SharePoint_ess_commerce_india";

  async function cdataTool(toolName, args) {
    const results = await cdataRequest("tools/call", {
      name: toolName,
      arguments: args
    });
    for (const r of results) {
      const content = r?.result?.content || [];
      if (r?.result?.isError) {
        const msg = content.find(c => c.type === "text")?.text || "Unknown CData error";
        throw new Error(`CData tool error: ${msg}`);
      }
      const text = content.find(c => c.type === "text")?.text;
      if (text) return text;
    }
    throw new Error("No content returned from CData tool");
  }

  // ─── Resolve SharePoint file path by unique document GUID ────────────────
  // "sourcedoc" in a Doc.aspx link is SharePoint's UniqueId for that file.
  // It doesn't change even if the file gets moved to a different folder,
  // so we look up the real path instead of guessing it from env vars.

  function extractSourceDocId(url) {
    try {
      const parsed = new URL(url);
      const raw = parsed.searchParams.get("sourcedoc"); // e.g. "{A7EA4DE7-...}"
      return raw ? raw.replace(/[{}]/g, "") : null;
    } catch {
      return null;
    }
  }

  // Minimal CSV line parser — handles quoted fields that contain commas,
  // since folder/file names like "ABS QA - Regression Agent, E2E" could
  // theoretically contain a comma once quoted by the driver.
  function parseCsvLine(line) {
    const result = [];
    let cur = "";
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQuotes) {
        if (ch === '"') {
          if (line[i + 1] === '"') { cur += '"'; i++; }
          else inQuotes = false;
        } else {
          cur += ch;
        }
      } else {
        if (ch === '"') inQuotes = true;
        else if (ch === ",") { result.push(cur); cur = ""; }
        else cur += ch;
      }
    }
    result.push(cur);
    return result.map(v => v.trim());
  }

  async function resolveRemoteFileByGuid(sourceDocId, library = "Shared Documents") {
    const raw = await cdataTool("queryData", {
      query: `SELECT Url FROM [${CATALOG}].[REST].[Files] WHERE Id = '${sourceDocId}'`
    });

    if (!raw) return null;

    const lines = raw.trim().split("\n");
    if (lines.length < 2) return null; // header only, no match found

    const headers = parseCsvLine(lines[0]).map(h => h.toLowerCase());
    const values = parseCsvLine(lines[1]);
    const idx = headers.indexOf("url");
    if (idx === -1) return null;

    const fullUrl = values[idx];
    // e.g. "https://adobe.sharepoint.com/sites/ess-commerce-india/Shared Documents/ABS QA - Regression Agent/E2E Sample Files/CMEINTAKE-2976 - Kakao Pay.xlsx"

    // @RemoteFile must be relative to @Library ("Shared Documents"), so strip
    // everything up to and including "<library>/"
    const marker = `${library}/`;
    const markerIdx = fullUrl.indexOf(marker);
    if (markerIdx === -1) return null;

    return decodeURIComponent(fullUrl.substring(markerIdx + marker.length));
    // -> "ABS QA - Regression Agent/E2E Sample Files/CMEINTAKE-2976 - Kakao Pay.xlsx"
  }

  async function resolveSharePointFileUrl(url) {
    try {
      const parsed = new URL(url);
      if (parsed.pathname.includes("/_layouts/")) {
        // Fallback: old folder-guess behavior (kept only as a safety net)
        const fileName = parsed.searchParams.get("file");
        if (fileName) {
          const folderPath = process.env.SP_FOLDER_PATH || "ABS - AI AGENT";
          return `${folderPath}/${fileName}`;
        }
      }
      return url;
    } catch {
      return url;
    }
  }

  // ─── Fetch Data from SharePoint ───────────────────────────────────────────────
  app.post("/sharepoint/data", async (req, res) => {
    try {
      const actionType = req.body.action || "INITIAL_PURCHASE";
      const { sharepointUrl } = req.body;
      if (!sharepointUrl) {
        return res.status(400).json({ error: "sharepointUrl is required" });
      }

      const remoteFile = await resolveSharePointFileUrl(sharepointUrl);

      // Extract file name from URL for DB
      let cleanFileName = "SharePoint_File";
      try {
        const parsedUrl = new URL(sharepointUrl);
        const fileParam = parsedUrl.searchParams.get("file");
        if (fileParam) cleanFileName = fileParam.replace(/\.[^/.]+$/, "");
      } catch (e) {}

      await initCData();

      const raw = await cdataTool("executeProcedure", {
        catalogName:   CATALOG,
        schemaName:    "REST",
        procedureName: "DownloadDocument",
        parameters: {
          "@Library":    "Shared Documents",
          "@RemoteFile": remoteFile
        }
      });

      let buffer;
      const lines = raw?.trim().split("\n");

      if (lines && lines.length >= 2) {
        const dataLine = lines[1].trim();
        const parts = dataLine.split(",");
        const success = parts[0].trim().toLowerCase();
        
        if (success !== "true") {
          throw new Error(`CData DownloadDocument returned success=false: ${raw}`);
        }

        const base64 = dataLine.substring(dataLine.indexOf(",") + 1).trim();
        
        if (!base64 || base64.length < 100) {
          throw new Error(`Base64 content too short or empty: "${base64?.substring(0, 50)}"`);
        }

        buffer = Buffer.from(base64, "base64");
        console.log(`[SP/CData] File decoded successfully, ${buffer.length.toLocaleString()} bytes`);

      } else {
        throw new Error(`Unexpected CData response format: ${raw?.substring(0, 200)}`);
      }

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      const result = [];
      workbook.worksheets.forEach(ws => {
        const headerRow = ws.getRow(1);
        let cols = { scenario: 2, country: null, actualResult: null };
        headerRow.eachCell((cell, colNumber) => {
        const headerName = (getCellText(cell.value) || "").trim().toLowerCase();
        if (headerName === "scenario" && cols.scenario === 2) cols.scenario = colNumber;
        if (headerName === "country" && !cols.country) cols.country = colNumber;
        if (["actual result", "adobe id", "result"].includes(headerName) && !cols.actualResult) {
          cols.actualResult = colNumber;
        }
      });

        // --- ROW PROCESSING ---
        ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (rowNumber <= 1) return;
          
          const rawScenario = row.getCell(cols.scenario).value;
          const scenario = rawScenario ? String(rawScenario).trim().toLowerCase() : "";

          const IGNORE_TEXTS = "VIP, 3ds, change plan, switch plan"; 
          const excludeList = IGNORE_TEXTS.split(',').map(s => s.trim().toLowerCase());
          if (excludeList.some(kw => scenario.includes(kw))) return;
          
          if (!IP_KEYWORDS.some(kw => scenario.includes(kw)) && !isInitialPurchaseSheet(ws.name)) return;

          let actualResultText = null;
          if (cols.actualResult) {
            const rawActualResult = row.getCell(cols.actualResult).value;
            actualResultText = getCellText(rawActualResult);
              // ? (typeof rawActualResult === 'object' && rawActualResult.richText)
              //   ? rawActualResult.richText.map(r => r.text).join('')
              //   : (typeof rawActualResult === 'object' && rawActualResult.result !== undefined)
              //     ? String(rawActualResult.result)
              //     : String(rawActualResult)
              // : null;
          }
          const extractedOrderNumber = extractOrderNumber(actualResultText);
          if (!extractedOrderNumber) return;

          result.push({
            scenario:     row.getCell(cols.scenario).value, 
            country:      cols.country ? row.getCell(cols.country).value : null,
            actualResult: actualResultText,
            orderNumber:  extractedOrderNumber,
            eccContract:  extractEccContract(actualResultText),
            sheet:        ws.name,
            rowNumber:    rowNumber,
            fileName:     cleanFileName, 
            runStatus:    "", 
            sapDb:        "",            
            sourceData:   "",            
            sapOutbound:  "" 
          });
        });
      });

      res.json({ status: "success", data: result });

    } catch (err) {
      console.error("[SP/CData] Error:", err.message);
      res.status(500).json({ error: err.message });
    }
  });

  // ─── Write Back to SharePoint ─────────────────────────────────────────────────
  app.post("/sharepoint/save", async (req, res) => {
    try {
      const { sharepointUrl, updates } = req.body;

      if (!sharepointUrl || !updates || !updates.length) {
        return res.status(400).json({ error: "sharepointUrl and updates are required" });
      }

      const remoteFile = await resolveSharePointFileUrl(sharepointUrl);
      console.log("[SP/Save] Writing back to:", remoteFile);

      await initCData();

      const raw = await cdataTool("executeProcedure", {
        catalogName:   CATALOG,
        schemaName:    "REST",
        procedureName: "DownloadDocument",
        parameters: {
          "@Library":    "Shared Documents",
          "@RemoteFile": remoteFile
        }
      });

      const lines = raw?.trim().split("\n");
      if (!lines || lines.length < 2) {
        throw new Error("Failed to download file for writeback");
      }
      const dataLine = lines[1].trim();
      if (!dataLine.toLowerCase().startsWith("true,")) {
        throw new Error("DownloadDocument returned failure");
      }
      const base64In = dataLine.substring(dataLine.indexOf(",") + 1).trim();
      const buffer   = Buffer.from(base64In, "base64");

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      const updateMap = {};
      updates.forEach(u => {
        updateMap[`${u.sheet}:${u.rowNumber}`] = {
          status: u.status || "",
          reason: u.reason || ""
        };
      });

      workbook.worksheets.forEach(ws => {
        const headerRow = ws.getRow(1);
        let statusCol = null;
        let reasonCol = null;

        headerRow.eachCell((cell, colNumber) => {
          const val = String(cell.value || "").trim().toLowerCase();
          if (val === "status")  statusCol = colNumber;
          if (val === "remarks" || val === "reason") reasonCol = colNumber;
        });

        if (!statusCol && !reasonCol) return; 

        ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (rowNumber <= 1) return;
          const key = `${ws.name}:${rowNumber}`;
          if (!updateMap[key]) return;

          const { status, reason } = updateMap[key];
          if (statusCol && status) {
            row.getCell(statusCol).value = status;
          }
          if (reasonCol && reason) {
            row.getCell(reasonCol).value = reason;
          }
          row.commit();
        });
      });

      const outBuffer = await workbook.xlsx.writeBuffer();
      const base64Out = Buffer.from(outBuffer).toString("base64");

      const uploadResult = await cdataTool("executeProcedure", {
        catalogName:   CATALOG,
        schemaName:    "REST",
        procedureName: "UploadDocument",
        parameters: {
          "@Library":     "Shared Documents",
          "@RelativeUrl": remoteFile,
          "@FileData":    base64Out
        }
      });

      console.log("[SP/Save] Upload result:", uploadResult?.substring(0, 100));
      res.json({ status: "success", updated: updates.length });

    } catch (err) {
      console.error("[SP/Save] Error:", err.message);
      res.status(500).json({ error: err.message });
    }
  });


  /* ================= 2. SAVE TO HANA DB ================= */
  app.post("/excel/save", async (req, res) => {
    try {
      // const action = row.action ? String(row.action) : 'INITIAL_PURCHASE';
      const rows = req.body;
      console.log("Received rows:", JSON.stringify(rows));

      const client = hdb.createClient({
        host:     process.env.HANA_HOST,
        port:     443,
        user:     process.env.HANA_USER,
        password: process.env.HANA_PASSWORD,
        encrypt:  true,
        sslValidateCertificate: false
      });

      await new Promise((resolve, reject) => {
        client.connect(err => err ? reject(err) : resolve());
      });

      console.log("Connected to HANA");

      const savedData = []; 

      for (const row of rows) {
        const scenario     = row.scenario     ? String(row.scenario)     : null;
        const country      = row.country      ? String(row.country)      : null;
        const actualResult = row.actualResult ? String(row.actualResult) : null;
        const orderNumber  = row.orderNumber  ? String(row.orderNumber)  : null; 
        const eccContract  = row.eccContract  ? String(row.eccContract)  : null; 
        const fileName     = row.fileName     ? String(row.fileName)     : null;
        const sheet        = row.sheet        ? String(row.sheet)        : null;
        const rowNumber    = row.rowNumber    ? parseInt(row.rowNumber)  : 0;
        const action       = row.action       ? String(row.action)       : 'INITIAL_PURCHASE';
        
        let status         = row.status       ? String(row.status)       : 'PENDING';
        let reason         = row.reason       ? String(row.reason)       : null; 
        let finalId        = null;
        let finalCreatedAt = null;
        
        const currentDateTime = new Date().toISOString().slice(0, 19).replace('T', ' ');

        // STEP 1: Check if this specific File, Sheet, and Row already exists
        let existingRecord = null;
        if (fileName && sheet && rowNumber) {
          existingRecord = await new Promise((resolve, reject) => {
            client.exec(
              `SELECT "ID", "STATUS", TO_NVARCHAR("REASON") AS "REASON", "CREATEDAT" 
               FROM "TRNFRM_BTP_CONSCOMM"."EXCELDATA_TESTRESULTS" 
               WHERE "FILENAME" = '${fileName.replace(/'/g,"''")}'
               AND "SHEETNAME" = '${sheet.replace(/'/g,"''")}' 
               AND "ROWNUMBER" = ${rowNumber}`,
              (err, result) => {
                if (err) { console.error("Select error:", err.message); reject(err); }
                else { resolve(result.length > 0 ? result[0] : null); }
              }
            );
          });
        }

        // STEP 2: Decide whether to Update or Insert
        if (existingRecord) {
          console.log(`File ${fileName}, Sheet ${sheet}, Row ${rowNumber} exists. Updating fields...`);
          
          finalId        = existingRecord.ID; // Keep the original UUID
          status         = existingRecord.STATUS;
          reason         = existingRecord.REASON;
          finalCreatedAt = existingRecord.CREATEDAT; 

          await new Promise((resolve, reject) => {
            client.exec(
              `UPDATE "TRNFRM_BTP_CONSCOMM"."EXCELDATA_TESTRESULTS"
               SET "SCENARIO" = ${scenario ? `'${scenario.replace(/'/g,"''")}'` : 'NULL'},
                   "COUNTRY" = ${country ? `'${country.replace(/'/g,"''")}'` : 'NULL'},
                   "ACTUALRESULT" = ${actualResult ? `'${actualResult.replace(/'/g,"''")}'` : 'NULL'},
                   "ORDERNUMBER" = ${orderNumber ? `'${orderNumber.replace(/'/g,"''")}'` : 'NULL'},
                   "UPDATEDAT" = '${currentDateTime}',
                   "ECC_CONTRACT" = ${eccContract ? `'${eccContract}'` : 'NULL'},
                   "Action" = '${action.replace(/'/g,"''")}'
               WHERE "FILENAME" = '${fileName.replace(/'/g,"''")}' 
               AND "SHEETNAME" = '${sheet.replace(/'/g,"''")}' 
               AND "ROWNUMBER" = ${rowNumber}`,
              (err) => {
                if (err) { console.error("Update error:", err.message); reject(err); }
                else { console.log("Updated record:", finalId); resolve(); }
              }
            );
          });

        } else {
          console.log(`File ${fileName}, Sheet ${sheet}, Row ${rowNumber} is new. Inserting...`);
          
          finalId        = crypto.randomUUID();
          finalCreatedAt = currentDateTime; 
          
          await new Promise((resolve, reject) => {
            client.exec(
              `INSERT INTO "TRNFRM_BTP_CONSCOMM"."EXCELDATA_TESTRESULTS"
               ("ID","SCENARIO","COUNTRY","ACTUALRESULT","ORDERNUMBER","STATUS","SHEETNAME","ROWNUMBER","REASON","CREATEDAT","UPDATEDAT","ECC_CONTRACT","FILENAME", "Action")
               VALUES ('${finalId}',
                       ${scenario     ? `'${scenario.replace(/'/g,"''")}'`     : 'NULL'},
                       ${country      ? `'${country.replace(/'/g,"''")}'`      : 'NULL'},
                       ${actualResult ? `'${actualResult.replace(/'/g,"''")}'` : 'NULL'},
                       ${orderNumber  ? `'${orderNumber.replace(/'/g,"''")}'`  : 'NULL'},
                       '${status.replace(/'/g,"''")}',
                       ${sheet        ? `'${sheet.replace(/'/g,"''")}'`        : 'NULL'},
                       ${rowNumber},
                       ${reason       ? `'${reason.replace(/'/g,"''")}'`       : 'NULL'},
                       '${finalCreatedAt}',
                       '${currentDateTime}',
                       ${eccContract  ? `'${eccContract}'` : 'NULL'},
                       ${fileName     ? `'${fileName.replace(/'/g,"''")}'`     : 'NULL'},
                       '${action.replace(/'/g,"''")}')`,
              (err) => {
                if (err) { console.error("Insert error:", err.message); reject(err); }
                else { console.log("Inserted brand new record:", finalId); resolve(); }
              }
            );
          });
        }

        savedData.push({
          ...row, 
          id: finalId,
          status: status,
          reason: reason,
          createdAt: finalCreatedAt,
          updatedAt: currentDateTime,
          eccContract: eccContract,
          fileName: fileName
        });
      }

      client.end();
      res.json({ success: true, saved: rows.length, data: savedData });

    } catch (err) {
      console.error("Save error:", err);
      res.status(500).json({ error: err.message });
    }
  });


  /* ================= 3. READ FROM HANA ================= */
  app.get("/excel/results", async (req, res) => {
    try {
      const client = hdb.createClient({
        host:     process.env.HANA_HOST,
        port:     443,
        user:     process.env.HANA_USER,
        password: process.env.HANA_PASSWORD,
        encrypt:  true,
        sslValidateCertificate: false
      });

      await new Promise((resolve, reject) => {
        client.connect(err => err ? reject(err) : resolve());
      });

      // Fetch the data, strictly converting NCLOB tracking fields for UI display
      const rows = await new Promise((resolve, reject) => {
        client.exec(`SELECT 
            "ID", "SCENARIO", "COUNTRY", "ACTUALRESULT", "ORDERNUMBER", 
            "STATUS", "SHEETNAME", "ROWNUMBER", "CREATEDAT", "ECC_CONTRACT", "FILENAME",
            TO_NVARCHAR("REASON") AS "REASON", 
            TO_NVARCHAR("RUN_STATUS") AS "RUN_STATUS"
          FROM "TRNFRM_BTP_CONSCOMM"."EXCELDATA_TESTRESULTS" 
          ORDER BY "CREATEDAT" DESC`, (err, result) => {
            if (err) reject(err);
            else resolve(result);
          });
      });

      client.end();

      const mappedRows = rows.map(row => ({
        id:           row.ID,
        scenario:     row.SCENARIO,
        country:      row.COUNTRY,
        actualResult: row.ACTUALRESULT,
        orderNumber:  row.ORDERNUMBER,
        status:       row.STATUS,
        sheet:        row.SHEETNAME,
        rowNumber:    row.ROWNUMBER,
        reason:       row.REASON,
        eccContract:  row.ECC_CONTRACT,
        fileName:     row.FILENAME,
        runStatus:    row.RUN_STATUS
      }));

      res.json(mappedRows);

    } catch (err) {
      console.error("Fetch error:", err);
      res.status(500).json({ error: err.message });
    }
  });


  /* ================= CDATA TESTS ================= */
  app.get("/cdata/init", async (req, res) => {
    try {
      const result = await initCData();
      res.json({ success: true, result });
    } catch (err) {
      console.error("[CData] Init error:", err.message);
      res.status(500).json({ success: false, error: err.message });
    }
  });

  app.get("/cdata/test", async (req, res) => {
    try {
      await initCData(); 
      const result = await queryData(
        `SELECT COUNT(*) AS total 
         FROM [iPaaS_SAPERP_PostgresWire_RS2].[public].[vbak] 
         WHERE [AUART] = 'ZCSB' 
         AND [ERDAT] >= '2026-06-01'`
      );
      res.json({ success: true, data: result });
    } catch (err) {
      console.error("[CData] Query error:", err.message);
      res.status(500).json({ success: false, error: err.message });
    }
  });

  // app.get("/cdata/test-resolve/:guid", async (req, res) => {
  //   try {
  //     await initCData();
  //     const resolved = await resolveRemoteFileByGuid(req.params.guid);
  //     res.json({ resolved });
  //   } catch (err) {
  //     res.status(500).json({ error: err.message });
  //   }
  // });

});