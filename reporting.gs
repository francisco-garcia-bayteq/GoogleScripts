const Reporting = (function () {
  function ensureReportSpreadsheet(config) {
    if (config.reportSpreadsheetId) {
      return config.reportSpreadsheetId;
    }
    const spreadsheet = Utils.createTimestampedSpreadsheet(config.reportFolderId);
    return spreadsheet.getId();
  }

  function logImportedData(spreadsheetId, batch) {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = 'Import - ' + batch.fileName;
    const sheet = Utils.insertOrResetSheet(spreadsheet, sheetName);
    const headers = batch.header.map((value) => (value == null ? '' : value));
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const values = batch.records.map((record) => record.rawRow.map((cell) => (cell == null ? '' : cell)));
    if (values.length) {
      sheet.getRange(2, 1, values.length, headers.length).setValues(values);
    }
    sheet.autoResizeColumns(1, headers.length);
  }

  function publishFinalReport(spreadsheetId, records, errors) {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const finalSheet = Utils.insertOrResetSheet(spreadsheet, 'Final Grades');
    const finalHeaders = ['Student ID', 'Name', 'Course', 'Grade 1', 'Grade 2', 'Grade 3', 'Final Grade', 'Status', 'Source File', 'Row'];
    const finalValues = records.map((record) => [
      record.id_student,
      record.name,
      record.course,
      record.grade1,
      record.grade2,
      record.grade3,
      record.finalGrade,
      record.status,
      record.sourceFile,
      record.rowNumber,
    ]);
    finalSheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
    if (finalValues.length) {
      finalSheet.getRange(2, 1, finalValues.length, finalHeaders.length).setValues(finalValues);
    }
    finalSheet.autoResizeColumns(1, finalHeaders.length);

    const errorSheet = Utils.insertOrResetSheet(spreadsheet, 'Validation Issues');
    const errorHeaders = ['File', 'Row', 'Issue'];
    const errorValues = errors.map((entry) => [
      entry.fileName,
      entry.rowNumber,
      entry.issues.join('; '),
    ]);
    errorSheet.getRange(1, 1, 1, errorHeaders.length).setValues([errorHeaders]);
    if (errorValues.length) {
      errorSheet.getRange(2, 1, errorValues.length, errorHeaders.length).setValues(errorValues);
    }
    errorSheet.autoResizeColumns(1, errorHeaders.length);
  }

  return {
    ensureReportSpreadsheet,
    logImportedData,
    publishFinalReport,
  };
})();
