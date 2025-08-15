Office.onReady();

function onSpamReport(spamEvent) {
  Office.context.mailbox.item.getAsFileAsync({ asyncContext: spamEvent }, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error(`Error processing message: ${asyncResult.error.message}`);
      spamEvent.completed({ onErrorDeleteItem: true });
      return;
    }

    const emlFile = asyncResult.value; // Base64-encoded EML
    const fileName = `${Office.context.mailbox.item.subject || "ReportedEmail"}.eml`;

    // 📌 Step 1: Convert Base64 to byte array for attachment
    const byteArray = _base64ToUint8Array(emlFile);

    // 📌 Step 2: Open new compose window with To, subject, and attachment
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ["security@mycompany.com"], // Change to your address
      subject: "Reported Suspicious Email",
      htmlBody: "<p>See attached reported email.</p>",
      attachments: [
        {
          type: "file",
          name: fileName,
          content: byteArray
        }
      ]
    });

    // 📌 Step 3: Continue spam reporting completion
    spamEvent.completed({
      onErrorDeleteItem: true,
      moveItemTo: Office.MailboxEnums.MoveSpamItemTo.JunkFolder,
      showPostProcessingDialog: {
        title: "Mawarid Spam Reporting",
        description: "Thank you for reporting this message.",
      },
    });
  });
}

// Helper: Base64 → Uint8Array
function _base64ToUint8Array(base64) {
  const binaryString = atob(base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes;
}

Office.actions.associate("onSpamReport", onSpamReport);
