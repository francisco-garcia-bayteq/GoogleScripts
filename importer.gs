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

  function moveToProcessed(file, folderId) {
    if (!file || !folderId) {
      return;
    }
    Utils.moveFileToFolder(file, folderId);
  }

  return {
    importAllCsv,
    moveToProcessed,
  };
})();
