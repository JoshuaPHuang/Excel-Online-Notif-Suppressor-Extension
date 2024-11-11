// Listen for messages from content script
chrome.runtime.onMessage.addListener(function(message, sender, sendResponse) {
    if (message.type === 'custom_notif') {
        chrome.notifications.create({
            type: 'basic',
            iconUrl: 'custom_excel_suppressor.png',
            title: 'Excel Suppressor Notification',
            message: message.text
        })
    }
});