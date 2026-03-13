const Email = (function () {
  function sendCertificates(records, config) {
    const subject = config.emailSubject || 'Grade certificate';
    records.forEach((record) => {
      const recipient = record.email || config.defaultRecipient;
      if (!recipient) {
        Logger.log('Skipping email for ' + (record.name || 'unknown student') + ': no recipient email configured.');
        return;
      }
      if (!record.certificatePdf) {
        Logger.log('Skipping email for ' + (record.name || 'unknown student') + ': certificate PDF is missing.');
        return;
      }

      const bodyText = Utils.renderTemplate(config.emailBodyTemplate, {
        name: record.name,
        course: record.course,
        finalGrade: record.finalGrade,
        status: record.status,
      });

      const options = {
        attachments: [record.certificatePdf],
        htmlBody: bodyText.replace(/\n/g, '<br>'),
        name: config.emailSenderName,
      };

      MailApp.sendEmail(recipient, subject, bodyText, options);
    });
  }

  return {
    sendCertificates,
  };
})();
