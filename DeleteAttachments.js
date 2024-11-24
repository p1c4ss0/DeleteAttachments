Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("removeButton").onclick = deleteAttachments;
    }
});

function deleteAttachments() {
    var item = Office.context.mailbox.item.attachments[0].id;
    if (item.attachments && item.attachments.length > 0) {
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
