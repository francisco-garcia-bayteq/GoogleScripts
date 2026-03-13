const Importer = (function () {
  function importAllCsv(folderId) {
    if (!folderId) {
      throw new Error('Import folder ID is not configured.');
    }
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const batches = [];

    while (files.hasNext()) {
      const file = files.next();
      if (!file.getName().toLowerCase().endsWith('.csv')) {
        continue;
      }
      const parsed = parseCsvFile(file);
      if (!parsed || !parsed.rows.length) {
        continue;
      }
      batches.push({
        file,
        fileName: file.getName(),
        header: parsed.header,
        records: parsed.records,
      });
    }
    return batches;
  }

  function parseCsvFile(file) {
    const content = file.getBlob().getDataAsString();
    const parsed = Utilities.parseCsv(content);
    if (!parsed || !parsed.length) {
      return null;
    }
    const header = parsed[0];
    const rows = parsed.slice(1).filter((row) => row.some((cell) => cell != null && cell.toString().trim() !== ''));
    const records = rows.map((row, index) => Utils.mapRowToRecord(header, row, file.getName(), index + 2));
    return { header, rows, records };
  }

  function importCsvFolderToRaw(importFolderId, spreadsheetId, processedFolderId) {
    if (!importFolderId) {
      throw new Error('Import folder ID is required for RAW import.');
    }
    if (!spreadsheetId) {
      throw new Error('Spreadsheet ID is required for RAW import.');
    }

    const folder = DriveApp.getFolderById(importFolderId);
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName('RAW_IMPORT');
    if (!sheet) {
      sheet = spreadsheet.insertSheet('RAW_IMPORT');
    }

    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (!file.getName().toLowerCase().endsWith('.csv')) {
        continue;
      }

      const csv = Utilities.parseCsv(file.getBlob().getDataAsString());
      if (!csv || csv.length < 2) {
        moveToProcessed(file, processedFolderId);
        continue;
      }

      const rowsToAppend = csv
        .slice(1)
        .filter((row) => row.some((cell) => cell != null && cell.toString().trim() !== ''));
      if (!rowsToAppend.length) {
        moveToProcessed(file, processedFolderId);
        continue;
      }

      const maxCols = rowsToAppend.reduce((acc, row) => Math.max(acc, row.length), 0);
      const paddedRows = rowsToAppend.map((row) => {
        const copy = row.slice();
        while (copy.length < maxCols) {
          copy.push('');
        }
        return copy;
      });

      const startRow = Math.max(sheet.getLastRow(), 0) + 1;
      sheet.getRange(startRow, 1, paddedRows.length, maxCols).setValues(paddedRows);

      moveToProcessed(file, processedFolderId);
    }
  }

  function moveToProcessed(file, folderId) {
    if (!file || !folderId) {
      return;
    }
    Utils.moveFileToFolder(file, folderId);
  }

  return {
    importAllCsv,
    importCsvFolderToRaw,
    moveToProcessed,
  };
})();
