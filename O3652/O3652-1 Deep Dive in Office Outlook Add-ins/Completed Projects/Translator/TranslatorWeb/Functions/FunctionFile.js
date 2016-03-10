Office.initialize = function () {
}

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!");
}

function translateEnglish2Spanish(event) {
  translate('en', 'es', function (error) {
    if (error) {
      Office.context.mailbox.item.notificationMessages.addAsync('translateError', {
        type: 'errorMessage',
        message: error
      });
    }
    else {
      Office.context.mailbox.item.notificationMessages.addAsync('success', {
        type: 'informationalMessage',
        icon: 'icon-16',
        message: 'Translated successfully',
        persistent: false
      });
    }
  });

  event.completed();
}