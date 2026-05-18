Office.onReady(() => {
    // Office is ready
});

function showPersistentBanner(event) {
    const notification = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Please keep all emails concise and professional. Aim for under 200 words.",
        icon: "Icon.16x16",
        persistent: true
    };
    
    Office.context.mailbox.item.notificationMessages.replaceAsync("alwaysOnBanner", notification, () => {
        event.completed(); 
    });
}

// Crucial: Associate the function name from the manifest with the actual JS function
Office.actions.associate("showPersistentBanner", showPersistentBanner);
