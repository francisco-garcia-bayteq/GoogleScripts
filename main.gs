function runGradeAutomation() {
  const config = Utils.getConfig();
  const batches = Importer.importAllCsv(config.importFolderId);
  if (!batches.length) {
    Logger.log('No CSV files found in the import folder.');
    return;
  }

  const reportSpreadsheetId = Reporting.ensureReportSpreadsheet(config);
  batches.forEach((batch) => Reporting.logImportedData(reportSpreadsheetId, batch));

  const allRecords = batches.reduce((acc, batch) => acc.concat(batch.records), []);
  const validationResult = Validation.validateRecords(allRecords);
  const calculatedRecords = validationResult.validRecords.map((record) =>
    Calculation.applyWeightedAverage(record, config.weights, config.approvalThreshold)
  );

  const recordsWithDocs = Documents.createCertificates(calculatedRecords, config);
  Reporting.publishFinalReport(reportSpreadsheetId, recordsWithDocs, validationResult.errors);
  Email.sendCertificates(recordsWithDocs, config);

  batches.forEach((batch) => Importer.moveToProcessed(batch.file, config.processedFolderId));

  Logger.log('Processed ' + calculatedRecords.length + ' grade records with ' + validationResult.errors.length + ' validation issues.');
}

function manualRun() {
  runGradeAutomation();
}
