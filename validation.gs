const Validation = (function () {
  const requiredFields = ['id_student', 'name', 'course', 'grade1', 'grade2', 'grade3'];

  function validateRecords(records) {
    const errors = [];
    const validRecords = [];
    const seenStudentIds = new Set();

    records.forEach((record) => {
      const issues = [];
      const normalized = {};

      requiredFields.forEach((field) => {
        const rawValue = record[field] == null ? '' : record[field].toString().trim();
        normalized[field] = rawValue;
        if (!rawValue) {
          issues.push('Missing ' + field);
        }
      });

      for (let i = 1; i <= 3; i += 1) {
        const gradeKey = 'grade' + i;
        const numeric = Utils.toNumber(normalized[gradeKey]);
        if (numeric == null) {
          issues.push('Invalid ' + gradeKey);
          normalized[gradeKey] = null;
        } else {
          normalized[gradeKey] = numeric;
          if (numeric < 0 || numeric > 10) {
            issues.push(gradeKey + ' out of range');
          }
        }
      }

      const studentId = normalized.id_student;
      if (studentId && seenStudentIds.has(studentId)) {
        issues.push('Duplicate student ID');
      } else if (studentId) {
        seenStudentIds.add(studentId);
      }

      if (issues.length) {
        errors.push({
          rowNumber: record.rowNumber,
          fileName: record.sourceFile,
          studentId: normalized.id_student,
          name: normalized.name,
          issues,
        });
      } else {
        validRecords.push(
          Object.assign({}, normalized, {
            rowNumber: record.rowNumber,
            sourceFile: record.sourceFile,
          })
        );
      }
    });

    return {
      validRecords,
      errors,
    };
  }

  function logErrors(spreadsheetId, errors) {
    if (!spreadsheetId || !errors.length) {
      return;
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = Utils.insertOrResetSheet(spreadsheet, 'LOG_ERRORS');
    const headers = ['Timestamp', 'Student ID', 'Name', 'Row', 'File', 'Issues'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const rows = errors.map((entry) => [
      timestamp,
      entry.studentId || '',
      entry.name || '',
      entry.rowNumber || '',
      entry.fileName || '',
      entry.issues.join('; '),
    ]);
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  function validateAndLog(records, spreadsheetId) {
    const result = validateRecords(records);
    logErrors(spreadsheetId, result.errors);
    return {
      cleanData: result.validRecords,
      errors: result.errors,
    };
  }

  return {
    validateRecords,
    validateAndLog,
  };
})();
