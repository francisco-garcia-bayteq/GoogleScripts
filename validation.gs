const Validation = (function () {
  const requiredFields = ['id_student', 'name', 'course', 'grade1', 'grade2', 'grade3'];

  function validateRecords(records) {
    const errors = [];
    const validRecords = [];
    const seenKeys = new Set();

    records.forEach((record) => {
      const issues = [];
      const normalized = {};

      requiredFields.forEach((field) => {
        const rawValue = record[field] == null ? '' : record[field].toString().trim();
        normalized[field] = rawValue;
        if (!rawValue) {
          issues.push(Missing );
        }
      });

      const grades = [];
      for (let i = 1; i <= 3; i += 1) {
        const gradeKey = grade;
        const numeric = Utils.toNumber(normalized[gradeKey]);
        if (numeric == null) {
          issues.push(Invalid );
          normalized[gradeKey] = null;
        } else {
          normalized[gradeKey] = numeric;
          if (numeric < 0 || numeric > 10) {
            issues.push(${gradeKey} out of range);
          }
          grades.push(numeric);
        }
      }

      const duplicateKey = ${normalized.id_student}|;
      if (duplicateKey && seenKeys.has(duplicateKey)) {
        issues.push('Duplicate student/course combination');
      } else if (duplicateKey) {
        seenKeys.add(duplicateKey);
      }

      if (issues.length) {
        errors.push({
          rowNumber: record.rowNumber,
          fileName: record.sourceFile,
          issues,
          record: record.rawRow,
        });
      } else {
        validRecords.push(Object.assign({}, normalized, {
          rowNumber: record.rowNumber,
          sourceFile: record.sourceFile,
        }));
      }
    });

    return {
      validRecords,
      errors,
    };
  }

  return {
    validateRecords,
  };
})();
