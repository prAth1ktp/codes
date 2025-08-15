Office.onReady();

function onSpamReport(event) {
  Office.context.mailbox.item.getAsFileAsync({ asyncContext: event }, asyncResult => {
    const spamEvent = asyncResult.asyncContext;

    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error(`Error processing message: ${asyncResult.error.message}`);
      spamEvent.completed({ onErrorDeleteItem: true });
      return;
    }

    const emlFile = asyncResult.value;
    const fileName = `${Office.context.mailbox.item.subject || "ReportedEmail"}.eml`;
    const byteArray = _base64ToUint8Array(emlFile);

    // Open compose window
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ["security@mycompany.com"], // change to your security email
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

    // Optionally move original message to Junk
    Office.context.mailbox.item.moveAsync(Office.MailboxEnums.MoveItemTo.JunkFolder, result => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to move message to Junk:", result.error.message);
      }
    });

    // Complete the event
    spamEvent.completed();
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

// Associate function
Office.actions.associate("onSpamReport", onSpamReport);
