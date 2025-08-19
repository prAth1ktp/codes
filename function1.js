Office.onReady();

function onSpamReport(spamEvent) {
  Office.context.mailbox.item.getAsFileAsync({ asyncContext: spamEvent }, async (asyncResult) => {
    const emlFile = asyncResult.value;
    const spamEvent = asyncResult.asyncContext;

    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error(`Error processing message: ${asyncResult.error.message}`);
      spamEvent.completed({ onErrorDeleteItem: true });
      return;
    }

    const fileName = `${Office.context.mailbox.item.subject || "ReportedEmail"}.eml`;
    const byteArray = _base64ToUint8Array(emlFile);

    // Open compose first
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ["security@mycompany.com"],
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

    // Wait a bit before completing spam event
    setTimeout(() => {
      spamEvent.completed({
        moveItemTo: Office.MailboxEnums.MoveSpamItemTo.JunkFolder,
        showPostProcessingDialog: {
          title: "Yahoo Spam Reporting",
          description: "Thank you for reporting this message.",
        }
      });
    }, 500); // half a second delay
  });
}

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
