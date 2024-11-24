Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("removeButton").onclick = deleteAttachments;
    }
});

function deleteAttachments() {
    var item = Office.context.mailbox.item;
    var totalAttachments = item.attachments.length;
    
    if (totalAttachments > 0) {
        console.log('Total attachments:', totalAttachments);
        item.attachments.forEach(function (attachment) {
            item.removeAttachmentAsync(attachment.id, function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error('Failed to remove attachment:', result.error.message);
                } else {
                    console.log('Attachment removed:', attachment.name);
                }
            });
        });
    } else {
        console.log("No attachments to remove.");
    }
}
