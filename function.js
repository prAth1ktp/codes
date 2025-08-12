Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        console.log("Office is ready in Outlook.");
    } else {
        console.log("Office is ready, but not in Outlook. Host:", info.host);
    }
});

function forwardEmail(event) {
    console.log("forwardEmail function started.");

    try {
        const item = Office.context.mailbox.item;
        const forwardAddress = "prathik@novanodit.com"; // change as needed

        item.displayReplyForm({
            htmlBody: "",
            toRecipients: [forwardAddress]
        });

        console.log("displayReplyForm called.");
    } catch (error) {
        console.error("Error in forwardEmail:", error);
    } finally {
        event.completed();
        console.log("event.completed() called.");
    }
}
