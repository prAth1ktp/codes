Office.onReady();

function onSpamReport(spamEvent) {

    console.log("forwardEmail function started6.");

    Office.context.mailbox.item.getAsFileAsync((asyncResult) => {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log(`Error encountered during processing: ${asyncResult.error.message}`);
    return;
  }

  const base64Msg = asyncResult.value;

//console.log(asyncResult.value);
const dataUrl = "data:application/vnd.ms-outlook;base64," + base64Msg;

   Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["prathik@novanodit.com"],
    subject: "Fwd: Selected Mail",
    htmlBody: "<p>Please see attached mail.</p>",

     attachments: [
    {
      type: "file",
      name: "forwarded.eml",
      url: dataUrl,
      isInline: false
    }
  ]
});

});

    // Wait a bit before completing spam event
    setTimeout(() => {
      spamEvent.completed({
        moveItemTo: Office.MailboxEnums.MoveSpamItemTo.JunkFolder,
        showPostProcessingDialog: {
          title: "Yahoo1 Spam Reporting",
          description: "Thank you for reporting this message.",
        }
      });
    }, 500); // half a second delay
}


Office.actions.associate("onSpamReport", onSpamReport);
