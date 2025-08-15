Office.onReady(() => {
  console.log("Add-in ready");
});

function forwardToSecurity(event) {
  const item = Office.context.mailbox.item;

  // Get EML and then create a new compose window
  item.getAsFileAsync("eml", (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emlBytes = _base64ToUint8Array(result.value);
      const fileName = `${item.subject || "ReportedEmail"}.eml`;

      Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["security@mycompany.com"], // change to your SOC mailbox
        subject: "Reported Suspicious Email",
        htmlBody: "<p>See attached reported email.</p>",
        attachments: [
          {
            type: "file",
            name: fileName,
            content: emlBytes
          }
        ]
      });
    }
    event.completed();
  });
}

function _base64ToUint8Array(base64) {
  const binary = atob(base64);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

Office.actions.associate("forwardToSecurity", forwardToSecurity);
