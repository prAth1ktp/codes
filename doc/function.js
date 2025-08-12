Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        // Office is ready
    }
});

function forwardEmail(event) {
    const item = Office.context.mailbox.item;
    const forwardAddress = "prathik@novanodit.com"; // change to your target email

    item.displayReplyForm({
        htmlBody: "",
        toRecipients: [forwardAddress]
    });

    event.completed();
}