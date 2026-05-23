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

const MASTER_MAPPING = {
  "Benevity": { DisbursementID: 25, CompanyName: 0, Project: 1, Activity: 10, Currency: 14,
   amount:(row) => {
    const donationAmount = parseFloat(row[18]) || 0; // Total Donation to be Acknowledged
    const matchAmount = parseFloat(row[19]) || 0; // Match Amount
    const fee = parseFloat(row[21]) || 0; // Merchant Fee
    const causeSupportFee = parseFloat(row[20]) || 0; // Cause Support Fee
    return donationAmount + matchAmount - fee - causeSupportFee;
   }
   },
  "CyberGrants": { DisbursementID: 2, CompanyName: 0, Project: 3, Activity: 6, Currency: 9 },
  "GoodStack": { DisbursementID: 0, CompanyName: 4, Project: null, Activity: null, Currency: 9 }
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
  
  let reportSummary = [];
  let finalLogs = [];

  for (let folderName in CONFIG) {
    if (folderName === "ArchiveFolderName") continue;

    const folder = platformsFolder.getFoldersByName(folderName).next();
    const files = folder.getFiles();
    const targetSheet = ss.getSheetByName(CONFIG[folderName].sheetName);
    const idIndex = CONFIG[folderName].idIndex;
    
    // Ensure sheet headers for Benevity
    if (folderName === "Benevity") {
      ensureBenevityHeaders(targetSheet);
    }

    // PERFORMANCE OPTIMIZATION: 
    // Load IDs into a Set ONCE per folder, not once per file.
    const existingIds = new Set();
    const lastRow = targetSheet.getLastRow();
    if (lastRow > 1) {
      const data = targetSheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues();
      data.forEach(row => { if (row[0]) existingIds.add(row[0].toString()); });
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

        if (folderName === "Benevity") {
          // Metadata and custom layout extraction for Benevity
          // 1. Period Ending (Row 5 - index 4)
          // 2. Disbursement ID (Row 8 - index 7)
          let periodEnding = "";
          let disbursementId = "";
          
          for (let i = 0; i < Math.min(csvData.length, 11); i++) {
            const firstVal = csvData[i][0] ? csvData[i][0].toString().trim() : "";
            if (firstVal === "Period Ending") {
              periodEnding = csvData[i][1] ? csvData[i][1].toString().trim() : "";
            } else if (firstVal === "Disbursement ID") {
              disbursementId = csvData[i][1] ? csvData[i][1].toString().trim() : "";
            }
          }
          
          // Format periodEnding date to DD-MM-YYYY
          const formattedPeriodEnding = formatBenevityDate(periodEnding);
          
          // Parse data rows starting from row 13 (index 12)
          if (csvData.length <= 12) {
            throw new Error("No data rows found below headers (starting at row 13)");
          }
          
          for (let i = 12; i < csvData.length; i++) {
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
            
            // Calculate Total Donation Received = Total Donation to be Acknowledged (18) + Match Amount (19) - Cause Support Fee (20) - Merchant Fee (21)
            const parseAmount = (val) => {
              if (!val) return 0;
              const cleanVal = val.toString().replace(/[$,]/g, '').trim();
              const parsed = parseFloat(cleanVal);
              return isNaN(parsed) ? 0 : parsed;
            };
            
            const ack = parseAmount(cleanRow[18]);
            const match = parseAmount(cleanRow[19]);
            const causeSupportFee = parseAmount(cleanRow[20]);
            const merchantFee = parseAmount(cleanRow[21]);
            
            const totalDonationReceived = Number((ack + match - causeSupportFee - merchantFee).toFixed(2));
            
            // Append the 3 new columns
            cleanRow.push(totalDonationReceived);
            cleanRow.push(formattedPeriodEnding);
            cleanRow.push(disbursementId);
            
            rowsToImport.push(cleanRow);
          }
          
          if (rowsToImport.length === 0) {
            throw new Error("No valid donation rows processed");
          }
          
          firstRowId = rowsToImport[0][idIndex].toString();
          
        } else {
          // For other platforms, standard data extraction starting from index 1 (second row)
          rowsToImport = csvData.slice(1);
          if (rowsToImport.length > 0) {
            firstRowId = rowsToImport[0][idIndex].toString();
          } else {
            throw new Error("No data rows found");
          }
        }

        // FAST LOOKUP: .has() on a Set is near-instant
        if (firstRowId && existingIds.has(firstRowId)) {
          let msg = `[DUPLICATE] Skipping file: ${file.getName()}`;
          finalLogs.push(msg);
          file.moveTo(archiveFolder);
          continue;
        }

        // Process data if not a duplicate
        const importStats = smartAppend(targetSheet, rowsToImport, idIndex, existingIds);
        
        // Log success to Sheet and Console
        logSheet.appendRow([new Date(), folderName, fileName, file.getId(), importStats.added]);
        let msg = `[SUCCESS] Processed '${fileName}' from folder '${folderName}'. Rows added: ${importStats.added}, Skipped: ${importStats.skipped}.`;
        finalLogs.push(msg);

        folderResults.filesProcessed++;
        folderResults.rowsAdded += importStats.added;
        folderResults.duplicatesSkipped += importStats.skipped;

        // Update the Set so the NEXT file in this loop knows about these new IDs
        rowsToImport.forEach(row => {
          if (row[idIndex]) {
            existingIds.add(row[idIndex].toString());
          }
        });

      } catch (e) {
        let msg = `Error in ${fileName}: ${e.message}`;
        finalLogs.push(msg);
        folderResults.errors.push(`Error in ${fileName}: ${e.message}`);
      }

      // 2. MOVE TO ARCHIVE ALWAYS (at the end of the loop, successful or not)
      file.moveTo(archiveFolder);
      let msg = `[ARCHIVED] Moved '${fileName}' to '${CONFIG.ArchiveFolderName}'.`;
      finalLogs.push(msg);
    }
    reportSummary.push(folderResults);
  }

  finalLogs.push("Process Finished Successfully.");
  return finalLogs;
}

/**
 * IMPROVED: Smart Append with protection against uneven column lengths and caching of existing IDs
 */
function smartAppend(sheet, rows, idIndex, existingIdsSet) {
  if (!rows || rows.length === 0) {
    return { added: 0, skipped: 0 };
  }
  
  const TARGET_COL_COUNT = sheet.getLastColumn() || rows[0].length; // Fallback if getLastColumn is 0
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
  rows.forEach(row => {
    // 1. Ensure ID exists
    const id = row[idIndex] ? row[idIndex].toString() : null;
    if (!id || id === "") return; // Skip empty rows

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
    existingIds.add(id);
  });

  // 4. Batch Write
  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, TARGET_COL_COUNT)
         .setValues(rowsToAppend);
  }

  return { added: rowsToAppend.length, skipped: (rows.length - rowsToAppend.length) };
}

/**
 * Helper to check if a single ID exists in the target column
 */
function isIdAlreadyInSheet(sheet, idIndex, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;
  
  // Get all IDs in the column
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

  // Add a newline to the end of content to ensure the loop processes the final row
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
      // Only push if the row isn't just empty (handles trailing newlines)
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

/**
 * Automatically ensures the target sheet has the proper columns for Benevity.
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
 * Formats a Date string from the report metadata (e.g. "Mon 4 May 2026 0:00:00") to "DD-MM-YYYY".
 */
function formatBenevityDate(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) {
    // Fallback: parse using Regex if Date constructor fails
    // e.g. Mon 4 May 2026 0:00:00
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
    return dateStr; // fallback to original if parsing completely fails
  }
  const day = d.getDate();
  const month = d.getMonth() + 1;
  const year = d.getFullYear();
  const dd = day < 10 ? "0" + day : day;
  const mm = month < 10 ? "0" + month : month;
  return `${dd}-${mm}-${year}`;
}
