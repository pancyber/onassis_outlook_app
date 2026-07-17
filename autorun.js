Office.onReady(() => {
    // Office is ready
});

function onNewAppointmentComposeHandler(event) {
    const notification = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Your calendar policy tip text goes here.", 
        icon: "Icon.16x16",
        persistent: true
    };
    
    Office.context.mailbox.item.notificationMessages.replaceAsync("alwaysOnBanner", notification, () => {
        event.completed(); 
    });
}

Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
