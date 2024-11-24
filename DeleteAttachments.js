Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        // Binding the countAttachments function to an element with id "countButton"
        document.getElementById("countButton").onclick = countAttachments;
    }
});

function countAttachments() {
    var item = Office.context.mailbox.item;
    var totalAttachments = item.attachments.length;

    // Display the total number of attachments on the screen
    var displayElement = document.getElementById("attachmentCount");
    if (!displayElement) {
        displayElement = document.createElement("p");
        displayElement.id = "attachmentCount";
        document.body.appendChild(displayElement);
    }
    displayElement.innerText = 'Total attachments: ' + totalAttachments;
}
