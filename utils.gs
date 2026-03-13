const Utils = (function () {
  const DEFAULT_CONFIG = {
    importFolderId: 'IMPORT_FOLDER_ID',
    processedFolderId: 'PROCESSED_FOLDER_ID',
    reportFolderId: 'REPORT_FOLDER_ID',
    reportSpreadsheetId: '',
    certificateFolderId: 'CERTIFICATE_FOLDER_ID',
    approvalThreshold: 7,
    weights: { grade1: 0.3, grade2: 0.3, grade3: 0.4 },
    emailSubject: 'Your grade certificate',
    emailBodyTemplate: 'Hello {{name}},\n\nYour final grade for {{course}} is {{finalGrade}} ({{status}}).\n\nBest regards,\nAcademic Office',
    defaultRecipient: '',
    emailSenderName: 'Academic Office',
  };

  function getConfig() {
    return JSON.parse(JSON.stringify(DEFAULT_CONFIG));
  }

  function normalizeHeader(headerRow) {
    return headerRow.map((cell) => {
      if (cell == null) {
        return '';
      }
      return cell.toString().trim().toLowerCase();
    });
  }

  function mapRowToRecord(headerRow, dataRow, fileName, line) {
    const normalizedHeaders = normalizeHeader(headerRow);
    const record = {
      rawRow: dataRow,
      rowNumber: line,
      sourceFile: fileName,
    };

    normalizedHeaders.forEach((key, index) => {
      if (!key) {
        return;
      }
      record[key] = dataRow[index] == null ? '' : dataRow[index].toString().trim();
    });

    return record;
  }

  function toNumber(value) {
    const num = parseFloat(value);
    return Number.isFinite(num) ? num : null;
  }

  function formatDateISO(date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  function insertOrResetSheet(spreadsheet, sheetName) {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      sheet.clearContents().clearFormats();
    } else {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    sheet.setFrozenRows(1);
    return sheet;
  }

  function moveFileToFolder(file, folderId) {
    if (!folderId) {
      return;
    }
    const targetFolder = DriveApp.getFolderById(folderId);
    targetFolder.addFile(file);
    const parents = file.getParents();
    while (parents.hasNext()) {
      const parent = parents.next();
      if (parent.getId() !== targetFolder.getId()) {
        parent.removeFile(file);
      }
    }
  }

  function createTimestampedSpreadsheet(folderId) {
    const today = formatDateISO(new Date());
    const spreadsheet = SpreadsheetApp.create('Academic Grade Report ' + today);
    if (folderId) {
      const file = DriveApp.getFileById(spreadsheet.getId());
      moveFileToFolder(file, folderId);
    }
    return spreadsheet;
  }

  function renderTemplate(template, data) {
    return template.replace(/{{\s*([a-zA-Z0-9_]+)\s*}}/g, function (match, key) {
      return data[key] != null ? data[key] : '';
    });
  }

  return {
    getConfig,
    normalizeHeader,
    mapRowToRecord,
    toNumber,
    insertOrResetSheet,
    moveFileToFolder,
    createTimestampedSpreadsheet,
    renderTemplate,
  };
})();
