const Documents = (function () {
  function createCertificates(records, config) {
    return records.map((record) => {
      const docName = record.name + ' - Grade Certificate';
      const doc = DocumentApp.create(docName);
      const body = doc.getBody();
      body.appendParagraph('Academic Performance Certificate').setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph('Student: ' + record.name);
      body.appendParagraph('Student ID: ' + record.id_student);
      body.appendParagraph('Course: ' + record.course);
      body.appendParagraph('Grade 1: ' + record.grade1);
      body.appendParagraph('Grade 2: ' + record.grade2);
      body.appendParagraph('Grade 3: ' + record.grade3);
      body.appendParagraph('');
      body.appendParagraph('Final Grade: ' + record.finalGrade);
      body.appendParagraph('Status: ' + record.status);
      body.appendParagraph('Approved if ' + record.approvalThreshold + ' or higher.');
      body.appendParagraph('');
      body.appendParagraph('Issued on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy') + '.');
      doc.saveAndClose();

      const file = DriveApp.getFileById(doc.getId());
      if (config.certificateFolderId) {
        Utils.moveFileToFolder(file, config.certificateFolderId);
      }

      return Object.assign({}, record, {
        certificateDocId: file.getId(),
        certificatePdf: file.getAs(MimeType.PDF),
      });
    });
  }

  return {
    createCertificates,
  };
})();
