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


// chrome.tabs.query({}, (tabs) => { // Get all tabs
//     tabs.forEach((tab) => {
//         chrome.webNavigation.getAllFrames({ tabId: tab.id }, (frames) => {
//             if (frames) {
//                 frames.forEach((frame) => {
//                     // if (frame.url.includes("officeapps.live.com") && frame.frameId && frame.parentFrameId !== 0) {
//                     console.log(`FrameId: ${frame.frameId}, URL: ${frame.url}, TabId: ${tab.id}`);
//                     // }
//                 });
//             }
//         });
//     });
// });