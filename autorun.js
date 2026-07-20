Office.onReady(() => {});

function onNewAppointmentComposeHandler(event) {
    const notification = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Your calendar policy tip text goes here.", 
        icon: "Icon.16x16",
        persistent: true
    };
    
    Office.context.mailbox.item.notificationMessages.replaceAsync("calendarBanner", notification, () => {
        event.completed(); 
    });
}

function onNewMessageComposeHandler(event) {
    const notification = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Please keep the email brief, clear and concise. Max 200 words - Max six (6) recipients!", 
        icon: "Icon.16x16",
        persistent: true
    };
    
    Office.context.mailbox.item.notificationMessages.replaceAsync("emailBanner", notification, () => {
        event.completed(); 
    });
}

Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
