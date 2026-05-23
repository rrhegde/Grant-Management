/**
 * Configuration Object: Map folder names to Sheet names and their Unique ID column index.
 * Index 0 = Column A, Index 2 = Column C, etc.
 */
const CONFIG = {
  "Benevity": { sheetName: "Benevity", idIndex: 12 },     // Transaction ID is Col M (Index 12)
  "CyberGrants": { sheetName: "CyberGrants", idIndex: 2 }, // Donation ID is Col C
  "GoodStack": { sheetName: "GoodStack", idIndex: 0 },    // donationId is Col A
  "ArchiveFolderName": "Archives"
};

/**
 * Master Sheet column mapping per platform.
 * Each key maps to the 0-based column index in the platform's sheet row (after processing).
 * null = empty string in Master sheet.
 * Master columns: Disbursement ID | Company Name | Project | Activity | Currency | Total Amount | Disbursed Date | Source
 */
const MASTER_MAPPING = {
  "Benevity": {
    DisbursementID: 25,   // Disbursement ID (appended column Z)
    CompanyName: 0,       // Company
    Project: 1,           // Project
    Activity: 10,         // Activity
    Currency: 14,         // Currency
    TotalAmount: 23,      // Total Donation Received (calculated column X)
    DisbursedDate: 24     // Period Ending (appended column Y)
  },
  "CyberGrants": {
    DisbursementID: null,
    CompanyName: 0,       // Company Name
    Project: 3,           // Program Name
    Activity: 6,          // Donation Designation
    Currency: 20,         // Payment Net Amount (Currency Code)
    TotalAmount: 19,      // Payment Net Amount
    DisbursedDate: 22     // Payment Date
  },
  "GoodStack": {
    DisbursementID: null,
    CompanyName: 4,       // partner
    Project: null,
    Activity: null,
    Currency: 2,          // currency code
    TotalAmount: 12,      // Total Donation Received (calculated, appended as column 13 / index 12)
    DisbursedDate: 10     // created at (UTC)
  }
};

/**
 * Main function to be triggered by your Web App button.
 */
function processAllFolders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetParent  = DriveApp.getFileById(ss.getId()).getParents().next();
  const platformsFolder = sheetParent.getFoldersByName("GrantManagementPlatforms").next();
  const archiveFolder = platformsFolder.getFoldersByName(CONFIG.ArchiveFolderName).next();
  const logSheet = ss.getSheetByName("Processing_Logs") || ss.insertSheet("Processing_Logs");
  const masterSheet = ss.getSheetByName("Master") || ss.insertSheet("Master");
  
  // Ensure Master sheet has headers
  ensureMasterHeaders(masterSheet);
  
  let reportSummary = [];
  let finalLogs = [];

  for (let folderName in CONFIG) {
    if (folderName === "ArchiveFolderName") continue;

    const folder = platformsFolder.getFoldersByName(folderName).next();
    const files = folder.getFiles();
    const targetSheet = ss.getSheetByName(CONFIG[folderName].sheetName);
    const idIndex = CONFIG[folderName].idIndex;
    
    // Ensure sheet headers
    if (folderName === "Benevity") {
      ensureBenevityHeaders(targetSheet);
    } else if (folderName === "GoodStack") {
      ensureGoodStackHeaders(targetSheet);
    }

    // PERFORMANCE OPTIMIZATION: 
    // Load IDs into a Set ONCE per folder, not once per file.
    const existingIds = new Set();
    const existingDisbursementIds = new Set();
    const lastRow = targetSheet.getLastRow();
    if (lastRow > 1) {
      if (folderName === "Benevity") {
        // Load Disbursement IDs for Benevity (Column Z is column index 26)
        const disbData = targetSheet.getRange(2, 26, lastRow - 1, 1).getValues();
        disbData.forEach(row => { if (row[0]) existingDisbursementIds.add(row[0].toString()); });
      } else {
        // Load Transaction/Donation IDs for other platforms
        const data = targetSheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues();
        data.forEach(row => { if (row[0]) existingIds.add(row[0].toString()); });
      }
    }

    let folderResults = { name: folderName, filesProcessed: 0, rowsAdded: 0, duplicatesSkipped: 0, errors: [] };

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();

      try {
        let content = file.getBlob().getDataAsString('UTF-8').trim();
        if (!content) throw new Error("File empty");

        let csvData = parseMessyCsv(content);
        if (csvData.length <= 1) throw new Error("Only headers/No data");

        let rowsToImport = [];
        let firstRowId = "";
        let disbursementId = "";

        if (folderName === "Benevity") {
          // Metadata and custom layout extraction for Benevity
          let periodEnding = "";
          
          for (let i = 0; i < Math.min(csvData.length, 11); i++) {
            const firstVal = csvData[i][0] ? csvData[i][0].toString().trim() : "";
            if (firstVal === "Period Ending") {
              periodEnding = csvData[i][1] ? csvData[i][1].toString().trim() : "";
            } else if (firstVal === "Disbursement ID") {
              disbursementId = csvData[i][1] ? csvData[i][1].toString().trim() : "";
            }
          }

          // FAST PATH: Check if this Disbursement ID has already been imported
          if (disbursementId && existingDisbursementIds.has(disbursementId)) {
            let msg = `[DUPLICATE] Skipping file (Disbursement ID '${disbursementId}' already processed): ${file.getName()}`;
            finalLogs.push(msg);
            file.moveTo(archiveFolder);
            continue;
          }
          
          // Format periodEnding date to DD-MM-YYYY
          const formattedPeriodEnding = formatBenevityDate(periodEnding);
          
          // Dynamically find the header row by searching for "Company" in the first column
          let headerIndex = -1;
          for (let i = 0; i < csvData.length; i++) {
            if (csvData[i][0] && csvData[i][0].toString().trim() === "Company") {
              headerIndex = i;
              break;
            }
          }
          
          if (headerIndex === -1) {
            throw new Error("Could not find the header row (starting with 'Company')");
          }
          
          for (let i = headerIndex + 1; i < csvData.length; i++) {
            const row = csvData[i];
            if (!row || row.length === 0) continue;
            
            const firstCell = row[0] ? row[0].toString().trim() : "";
            if (firstCell.toLowerCase().startsWith("totals")) {
              break;
            }
            
            // Pad or slice row to exactly 23 columns
            let cleanRow = row.slice(0, 23);
            while (cleanRow.length < 23) {
              cleanRow.push("");
            }
            
            // Calculate Total Donation Received
            const ack = parseAmount(cleanRow[18]);
            const match = parseAmount(cleanRow[19]);
            const causeSupportFee = parseAmount(cleanRow[20]);
            const merchantFee = parseAmount(cleanRow[21]);
            const totalDonationReceived = Number((ack + match - causeSupportFee - merchantFee).toFixed(2));
            
            // Append the 3 new columns: Total Donation Received, Period Ending, Disbursement ID
            cleanRow.push(totalDonationReceived);   // index 23
            cleanRow.push(formattedPeriodEnding);    // index 24
            cleanRow.push(disbursementId);           // index 25
            
            rowsToImport.push(cleanRow);
          }
          
          if (rowsToImport.length === 0) {
            throw new Error("No valid donation rows processed");
          }

        } else if (folderName === "GoodStack") {
          // GoodStack: standard CSV with header in row 1
          rowsToImport = csvData.slice(1);
          if (rowsToImport.length === 0) throw new Error("No data rows found");

          firstRowId = rowsToImport[0][idIndex] ? rowsToImport[0][idIndex].toString() : "";

          // FAST LOOKUP: Check Transaction ID duplicate
          if (firstRowId && existingIds.has(firstRowId)) {
            let msg = `[DUPLICATE] Skipping file: ${file.getName()}`;
            finalLogs.push(msg);
            file.moveTo(archiveFolder);
            continue;
          }

          // Calculate and append Total Donation Received = gross donation (1) - fees (8)
          rowsToImport = rowsToImport.map(row => {
            let cleanRow = row.slice(0, 12);
            while (cleanRow.length < 12) {
              cleanRow.push("");
            }
            const grossDonation = parseAmount(cleanRow[1]);
            const fees = parseAmount(cleanRow[8]);
            const totalDonationReceived = Number((grossDonation - fees).toFixed(2));
            cleanRow.push(totalDonationReceived); // index 12
            return cleanRow;
          });

        } else {
          // CyberGrants: standard CSV with header in row 1
          rowsToImport = csvData.slice(1);
          if (rowsToImport.length > 0) {
            firstRowId = rowsToImport[0][idIndex].toString();
          } else {
            throw new Error("No data rows found");
          }

          // FAST LOOKUP: Check Transaction ID duplicate for CyberGrants
          if (firstRowId && existingIds.has(firstRowId)) {
            let msg = `[DUPLICATE] Skipping file: ${file.getName()}`;
            finalLogs.push(msg);
            file.moveTo(archiveFolder);
            continue;
          }
        }

        // Write to the individual platform sheet
        const appendIdsSet = folderName === "Benevity" ? new Set() : existingIds;
        const importStats = smartAppend(targetSheet, rowsToImport, idIndex, appendIdsSet);
        
        // Build and write Master sheet rows from successfully imported data
        if (importStats.appendedRows.length > 0) {
          const masterRows = importStats.appendedRows.map(row => mapToMasterRow(row, folderName));
          appendToMaster(masterSheet, masterRows);
        }

        // Log success to Sheet and Console
        logSheet.appendRow([new Date(), folderName, fileName, file.getId(), importStats.added]);
        let msg = `[SUCCESS] Processed '${fileName}' from folder '${folderName}'. Rows added: ${importStats.added}, Skipped: ${importStats.skipped}.`;
        finalLogs.push(msg);

        folderResults.filesProcessed++;
        folderResults.rowsAdded += importStats.added;
        folderResults.duplicatesSkipped += importStats.skipped;

        // Update the Sets so the NEXT file in this loop knows about these new IDs
        if (folderName === "Benevity") {
          if (disbursementId) {
            existingDisbursementIds.add(disbursementId);
          }
        } else {
          rowsToImport.forEach(row => {
            if (row[idIndex]) {
              existingIds.add(row[idIndex].toString());
            }
          });
        }

      } catch (e) {
        let msg = `Error in ${fileName}: ${e.message}`;
        finalLogs.push(msg);
        folderResults.errors.push(`Error in ${fileName}: ${e.message}`);
      }

      // MOVE TO ARCHIVE ALWAYS (at the end of the loop, successful or not)
      file.moveTo(archiveFolder);
      let msg = `[ARCHIVED] Moved '${fileName}' to '${CONFIG.ArchiveFolderName}'.`;
      finalLogs.push(msg);
    }
    reportSummary.push(folderResults);
  }

  finalLogs.push("Process Finished Successfully.");
  return finalLogs;
}

// ============================================================
// HELPER: Parse currency/amount strings to float
// ============================================================
function parseAmount(val) {
  if (!val) return 0;
  const cleanVal = val.toString().replace(/[$,]/g, '').trim();
  const parsed = parseFloat(cleanVal);
  return isNaN(parsed) ? 0 : parsed;
}

// ============================================================
// MASTER SHEET: Map a platform row to the Master sheet format
// ============================================================
/**
 * Transforms a single platform row into a Master sheet row using MASTER_MAPPING.
 * Master columns: Disbursement ID | Company Name | Project | Activity | Currency | Total Amount | Disbursed Date | Source
 */
function mapToMasterRow(row, folderName) {
  const mapping = MASTER_MAPPING[folderName];
  if (!mapping) return [];

  const getVal = (index) => {
    if (index === null || index === undefined) return "";
    return (row[index] !== undefined && row[index] !== null) ? row[index] : "";
  };

  return [
    getVal(mapping.DisbursementID),
    getVal(mapping.CompanyName),
    getVal(mapping.Project),
    getVal(mapping.Activity),
    getVal(mapping.Currency),
    getVal(mapping.TotalAmount),
    getVal(mapping.DisbursedDate),
    folderName  // Source
  ];
}

// ============================================================
// MASTER SHEET: Batch append rows (no dedup needed, already filtered)
// ============================================================
function appendToMaster(masterSheet, rows) {
  if (!rows || rows.length === 0) return;
  const colCount = 8; // Master sheet always has exactly 8 columns
  masterSheet.getRange(masterSheet.getLastRow() + 1, 1, rows.length, colCount)
             .setValues(rows);
}

// ============================================================
// SMART APPEND: Now returns appendedRows for Master sheet population
// ============================================================
function smartAppend(sheet, rows, idIndex, existingIdsSet) {
  if (!rows || rows.length === 0) {
    return { added: 0, skipped: 0, appendedRows: [] };
  }
  
  const TARGET_COL_COUNT = sheet.getLastColumn() || rows[0].length;
  const lastRow = sheet.getLastRow();
  
  // Get existing IDs from Column (idIndex + 1)
  let existingIds = existingIdsSet;
  if (!existingIds) {
    existingIds = new Set();
    if (lastRow > 1) {
      const data = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues();
      data.forEach(r => { if(r[0]) existingIds.add(r[0].toString()) });
    }
  }

  const rowsToAppend = [];
  const originalRows = []; // Track original rows before padding/trimming for Master mapping
  rows.forEach(row => {
    // 1. Ensure ID exists
    const id = row[idIndex] ? row[idIndex].toString() : null;
    if (!id || id === "") return;

    // 2. Prevent Duplicates
    if (existingIds.has(id)) return;

    // 3. Force Pad to sheet column count
    let cleanRow = [...row];
    while (cleanRow.length < TARGET_COL_COUNT) {
      cleanRow.push("");
    }
    if (cleanRow.length > TARGET_COL_COUNT) {
      cleanRow = cleanRow.slice(0, TARGET_COL_COUNT);
    }
    
    rowsToAppend.push(cleanRow);
    originalRows.push([...row]); // Keep the full original row for Master mapping
    existingIds.add(id);
  });

  // Batch Write
  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, TARGET_COL_COUNT)
         .setValues(rowsToAppend);
  }

  return { added: rowsToAppend.length, skipped: (rows.length - rowsToAppend.length), appendedRows: originalRows };
}

/**
 * Helper to check if a single ID exists in the target column
 */
function isIdAlreadyInSheet(sheet, idIndex, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;
  
  const existingIds = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues().flat();
  return existingIds.includes(id.toString());
}

/**
 * Required for Web App functionality
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Grant Data Processor')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * IMPROVED: Robust Parsing that handles uneven columns and empty lines
 */
function parseMessyCsv(content) {
  const result = [];
  let row = [];
  let field = "";
  let inQuotes = false;

  const data = content.trim() + "\n";

  for (let i = 0; i < data.length; i++) {
    const char = data[i];
    
    if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === ',' && !inQuotes) {
      row.push(field.trim());
      field = "";
    } else if (char === '\n' && !inQuotes) {
      row.push(field.trim());
      if (row.length > 1 || (row.length === 1 && row[0] !== "")) {
        result.push(row);
      }
      row = [];
      field = "";
    } else {
      field += char;
    }
  }
  return result;
}

// ============================================================
// HEADER SETUP FUNCTIONS
// ============================================================

/**
 * Ensures the Benevity sheet has the proper 26-column headers.
 */
function ensureBenevityHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    const headers = [
      "Company", "Project", "Donation Date", "Donor First Name", "Donor Last Name", "Email", 
      "Address", "City", "State/Province", "Postal Code", "Activity", "Comment", "Transaction ID", 
      "Donation Frequency", "Currency", "Project Remote ID", "Source", "Reason", 
      "Total Donation to be Acknowledged", "Match Amount", "Cause Support Fee", "Merchant Fee", "Fee Comment",
      "Total Donation Received", "Period Ending", "Disbursement ID"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (lastCol < 26) {
    const headers = ["Total Donation Received", "Period Ending", "Disbursement ID"];
    sheet.getRange(1, lastCol + 1, 1, headers.length).setValues([headers]);
  }
}

/**
 * Ensures the GoodStack sheet has the "Total Donation Received" column appended.
 */
function ensureGoodStackHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    const headers = [
      "donation id", "gross donation", "currency code", "marketing consent", "partner",
      "donor first name", "donor last name", "donor email", "fees", "fees currency",
      "created at (UTC)", "gift note", "Total Donation Received"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (lastCol < 13) {
    sheet.getRange(1, lastCol + 1, 1, 1).setValues([["Total Donation Received"]]);
  }
}

/**
 * Ensures the Master sheet has the correct 8-column headers.
 */
function ensureMasterHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    const headers = [
      "Disbursement ID", "Company Name", "Project", "Activity",
      "Currency", "Total Amount", "Disbursed Date", "Source"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

/**
 * Formats a Date string from the report metadata (e.g. "Mon 4 May 2026 0:00:00") to "DD-MM-YYYY".
 */
function formatBenevityDate(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) {
    // Fallback: parse using Regex if Date constructor fails
    const match = dateStr.match(/(\d+)\s+([A-Za-z]+)\s+(\d{4})/);
    if (match) {
      const day = parseInt(match[1], 10);
      const monthStr = match[2].toLowerCase();
      const year = match[3];
      const months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
      const monthIndex = months.findIndex(m => monthStr.startsWith(m));
      if (monthIndex !== -1) {
        const dd = day < 10 ? "0" + day : day;
        const mm = (monthIndex + 1) < 10 ? "0" + (monthIndex + 1) : (monthIndex + 1);
        return `${dd}-${mm}-${year}`;
      }
    }
    return dateStr;
  }
  const day = d.getDate();
  const month = d.getMonth() + 1;
  const year = d.getFullYear();
  const dd = day < 10 ? "0" + day : day;
  const mm = month < 10 ? "0" + month : month;
  return `${dd}-${mm}-${year}`;
}
