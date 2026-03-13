/**
 * Executes the full grade-processing pipeline: imports, validates, calculates,
 * builds certificates, emails results, and archives processed CSVs.
 */
function runGradeAutomation() {
  const config = Utils.getConfig();
  const batches = Importer.importAllCsv(config.importFolderId);
  if (!batches.length) {
    Logger.log('No CSV files found in the import folder.');
    return;
  }

  const reportSpreadsheetId = Reporting.ensureReportSpreadsheet(config);
  batches.forEach(function (batch) {
    Reporting.logImportedData(reportSpreadsheetId, batch);
  });

  const allRecords = batches.reduce(function (acc, batch) {
    return acc.concat(batch.records);
  }, []);
  const validationResult = Validation.validateRecords(allRecords);
  const calculatedRecords = validationResult.validRecords.map(function (record) {
    return Calculation.applyWeightedAverage(record, config.weights, config.approvalThreshold);
  });

  const recordsWithDocs = Documents.createCertificates(calculatedRecords, config);
  Reporting.publishFinalReport(reportSpreadsheetId, recordsWithDocs, validationResult.errors);
  Email.sendCertificates(recordsWithDocs, config);

  batches.forEach(function (batch) {
    Importer.moveToProcessed(batch.file, config.processedFolderId);
  });

  Logger.log('Processed ' + calculatedRecords.length + ' grade records with ' + validationResult.errors.length + ' validation issues.');
}

/**
 * Convenient entry point for manual execution or trigger wiring.
 */
function manualRun() {
  runGradeAutomation();
}
