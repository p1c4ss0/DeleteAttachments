Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("deleteAttachmentsButton").onclick = deleteAttachments;
    }
});

async function deleteAttachments() {
    try {
        const selectedItems = Office.context.mailbox.item;
        const itemIds = selectedItems ? [selectedItems.itemId] : await getSelectedEmailIds();

        if (itemIds.length === 0) {
            throw new Error("No emails selected.");
        }

        for (const itemId of itemIds) {
            await deleteAllAttachments(itemId);
        }

        showDialog("Success", "All attachments have been successfully deleted.");
    } catch (error) {
        showDialog("Error", error.message);
    }
}

async function getSelectedEmailIds() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(new Error("Failed to retrieve selected emails."));
            }
        });
    });
}

async function deleteAllAttachments(itemId) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentsAsync(itemId, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const attachments = result.value;
                if (attachments.length === 0) {
                    resolve();
                } else {
                    const deletePromises = attachments.map(attachment => {
                        return new Promise((res, rej) => {
                            Office.context.mailbox.item.removeAttachmentAsync(attachment.id, (removeResult) => {
                                if (removeResult.status === Office.AsyncResultStatus.Succeeded) {
                                    res();
                                } else {
                                    rej(new Error(`Failed to delete attachment: ${attachment.name}`));
                                }
                            });
                        });
                    });
                    Promise.all(deletePromises).then(resolve).catch(reject);
                }
            } else {
                reject(new Error("Failed to retrieve attachments."));
            }
        });
    });
}

function showDialog(title, message) {
    const dialog = document.createElement("div");
    dialog.className = "dialog";
    dialog.innerHTML = `<h2>${title}</h2><p>${message}</p><button onclick="this.parentElement.remove()">Close</button>`;
    document.body.appendChild(dialog);
}
