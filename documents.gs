const Documents = (function () {
  const CERTIFICATE_TEMPLATE = [
    'ACTA DE CALIFICACIONES',
    '-----------------------',
    '',
    'Nombre: {{name}}',
    'Curso: {{course}}',
    'Promedio ponderado: {{average}}',
    'Estado: {{status}}',
    '',
    'Esta acta certifica el desempeño académico correspondiente.',
    '',
    'Emitido el {{issuedOn}}',
  ].join('\n');

  function createCertificates(records, config) {
    const folderId = config.certificateFolderId;
    const folder = folderId ? DriveApp.getFolderById(folderId) : null;
    const issuedOn = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy');

    return records.map((record) => {
      const docName = Acta - ;
      const doc = DocumentApp.create(docName);
      const body = doc.getBody();
      body.setText(
        Utils.renderTemplate(CERTIFICATE_TEMPLATE, {
          name: record.name,
          course: record.course,
          average: record.average != null ? record.average : record.finalGrade,
          status: record.status,
          issuedOn,
        })
      );
      doc.saveAndClose();

      const file = DriveApp.getFileById(doc.getId());
      if (folder) {
        folder.addFile(file);
        const parents = file.getParents();
        while (parents.hasNext()) {
          const parent = parents.next();
          if (parent.getId() !== folder.getId()) {
            parent.removeFile(file);
          }
        }
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
