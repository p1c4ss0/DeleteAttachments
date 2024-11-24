Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("removeButton").onclick = countAttachments;
    }
});

function countAttachments() {
    var item = Office.context.mailbox.item;
    var totalAttachments = item.attachments.length;
    console.log('Total attachments:', totalAttachments);
}
