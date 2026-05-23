/**
 * Configuration Object: Map folder names to Sheet names and their Unique ID column index.
 * Index 0 = Column A, Index 2 = Column C, etc.
 */
const CONFIG = {
  "Benevity": { sheetName: "Benevity", idIndex: 14 },     // Transaction ID is Col O
  "CyberGrants": { sheetName: "CyberGrants", idIndex: 2 }, // Donation ID is Col C
  "GoodStack": { sheetName: "GoodStack", idIndex: 0 },    // donationId is Col A
  "ArchiveFolderName": "Archives"
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
    
    // PERFORMANCE OPTIMIZATION: 
    // Load IDs into a Set ONCE per folder, not once per file.
    const existingIds = new Set();
    const lastRow = targetSheet.getLastRow();
    if (lastRow > 1) {
      const data = targetSheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues();
      data.forEach(row => existingIds.add(row[0].toString()));
    }

    let folderResults = { name: folderName, filesProcessed: 0, rowsAdded: 0, duplicatesSkipped: 0, errors: [] };

    while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const idIndex = CONFIG[folderName].idIndex;

    // // 1. Check if already processed via Logs
    // if (isFileAlreadyProcessed(logSheet, fileName)) {
    //   console.log(`[ALREADY PROCESSED] File '${fileName}' from folder '${folderName}' found in logs. Moving to Archive.`);
    //   file.moveTo(archiveFolder);
    //   continue;
    // }

    try {
      let content = file.getBlob().getDataAsString('UTF-8').trim();
      if (!content) throw new Error("File empty");

      let csvData = parseMessyCsv(content);
      if (csvData.length <= 1) throw new Error("Only headers/No data");


      // const firstRowId = csvData[1][idIndex]; 

      // // CHECK IF THIS ID EXISTS IN THE SHEET
      //   if (isIdAlreadyInSheet(targetSheet, idIndex, firstRowId)) {
      //     console.log(`[DUPLICATE DETECTED] File '${fileName}' (First ID: ${firstRowId}) already exists in sheet. Moving to Archive.`);
      //     file.moveTo(archiveFolder);
      //     continue;
      //   }

      // GET THE UNIQUE ID OF THE FIRST DATA ROW
      // FAST LOOKUP: .has() on a Set is near-instant
      const firstRowId = csvData[1][idIndex].toString();
      if (existingIds.has(firstRowId)) {
        //console.log(`[DUPLICATE] Skipping file: ${file.getName()}`);
        let msg = `[DUPLICATE] Skipping file: ${file.getName()}`;
        finalLogs.push(msg);
        file.moveTo(archiveFolder);
        continue;
      }

      // Process data if not a duplicate
      const rows = csvData.slice(1); 
      const importStats = smartAppend(targetSheet, rows, CONFIG[folderName].idIndex);
      
      // Log success to Sheet and Console
      logSheet.appendRow([new Date(), folderName, fileName, file.getId(), importStats.added]);
      //console.log(`[SUCCESS] Processed '${fileName}' from folder '${folderName}'. Rows added: ${importStats.added}, Skipped: ${importStats.skipped}.`);
      let msg = `[SUCCESS] Processed '${fileName}' from folder '${folderName}'. Rows added: ${importStats.added}, Skipped: ${importStats.skipped}.`;
      finalLogs.push(msg);

      folderResults.filesProcessed++;
      folderResults.rowsAdded += importStats.added;
      folderResults.duplicatesSkipped += importStats.skipped;

      // Update the Set so the NEXT file in this loop knows about these new IDs
      csvData.slice(1).forEach(row => existingIds.add(row[idIndex].toString()));

    } catch (e) {
      //console.error(`Error in ${fileName}: ${e.message}`);
      let msg = `Error in ${fileName}: ${e.message}`;
      finalLogs.push(msg);
      folderResults.errors.push(`Error in ${fileName}: ${e.message}`);
    }

      // 2. MOVE TO ARCHIVE ALWAYS (at the end of the loop, successful or not)
      // This ensures your folders stay clean. 
      // If you only want to archive successful ones, keep this inside the 'try' block.
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
 * IMPROVED: Smart Append with protection against uneven column lengths
 */
function smartAppend(sheet, rows, idIndex) {
  const TARGET_COL_COUNT = sheet.getLastColumn(); // Set this to exactly what your sheet has
  const lastRow = sheet.getLastRow();
  
  // Get existing IDs from Column (idIndex + 1)
  let existingIds = new Set();
  if (lastRow > 1) {
    const data = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues();
    data.forEach(r => { if(r[0]) existingIds.add(r[0].toString()) });
  }

  const rowsToAppend = [];
  rows.forEach(row => {
    // 1. Ensure ID exists
    const id = row[idIndex] ? row[idIndex].toString() : null;
    if (!id || id === "") return; // Skip empty rows

    // 2. Prevent Duplicates
    if (existingIds.has(id)) return;

    // 3. Force Pad to 34 columns
    while (row.length < TARGET_COL_COUNT) {
      row.push("");
    }
    
    rowsToAppend.push(row);
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
 * Checks the Processing_Logs sheet for the filename.
 */
// function isFileAlreadyProcessed(logSheet, fileName) {
//   const lastRow = logSheet.getLastRow();
//   if (lastRow < 2) return false;
//   // Only get Column C
//   const filenames = logSheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
//   return filenames.includes(fileName);
// }

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

