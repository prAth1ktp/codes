Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        console.log("Office is ready in Outlook.");
    } else {
        console.log("Office is ready, but not in Outlook. Host:", info.host);
    }
});

function action(event) {
    console.log("forwardEmail function started6.");



    const itemType = Office.context.mailbox.item;

    Office.context.mailbox.item.getAsFileAsync((asyncResult) => {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log(`Error encountered during processing: ${asyncResult.error.message}`);
    return;
  }

  const base64Msg = asyncResult.value;

//  console.log(asyncResult.value);
const dataUrl = "data:application/vnd.ms-outlook;base64," + base64Msg;

   Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["prathik@novanodit.com"],
    subject: "Fwd: Spam / Phishing Reporting",
    htmlBody: "<p>I would like to report the attached spam / phishing Email.</p>",

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



}

// **This line is critical:**
if (typeof Office !== "undefined" && Office.actions) {
    Office.actions.associate("action", action);
} else {
    console.error("Office.actions is not available");
}


//change links to weblink
//show in main pane
