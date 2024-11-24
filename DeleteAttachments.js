Office.onReady(function(() => {

}):
  // ... (button click event handler)
async function deleteAttachments() {
    Office.context.mailbox.item.load('attachments');
    Office.context.mailbox.item.retrieveAsync().then(function() {
      var attachments = Office.context.mailbox.item.attachments;
      attachments.forEach(function(attachment) {
        attachment.deleteAsync().then(function() {
          console.log('Attachment deleted: ' + attachment.name);
        }, function(error) {
          console.error('Error deleting attachment: ' + error);
        });
      });
    });

office.actions.associate("deleteAttachments" , deleteAttachments) ;
