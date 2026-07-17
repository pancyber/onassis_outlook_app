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
        message: "Please keep all emails concise and professional. Aim for under 200 words.", 
        icon: "Icon.16x16",
        persistent: true
    };
    
    Office.context.mailbox.item.notificationMessages.replaceAsync("emailBanner", notification, () => {
        event.completed(); 
    });
}

Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
