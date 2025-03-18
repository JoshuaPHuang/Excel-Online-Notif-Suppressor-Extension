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

// Listen for the request to find all iframes with frame.url starting with "https://usc-excel.officeapps.live.com"
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "findIframes") {
        // Get all open tabs
        chrome.tabs.query({}, (tabs) => {
            let allFrames = [];
            let pendingPromises = tabs.map(tab => {
                return new Promise((resolve) => {
                    chrome.webNavigation.getAllFrames({ tabId: tab.id }, (frames) => {
                        if (chrome.runtime.lastError || !frames) {
                            return resolve([]);
                        }
                        // Filter iframes and add tab.id as frame.tabId to each frame
                        // console.log("Frames:");
                        // console.log(frames);
                        let officeIframes = frames
                        .filter(frame => frame.url.startsWith("https://usc-excel.officeapps.live.com"))
                        .map(frame => ({ ...frame, customTabId: tab.id })); // Add tabId to each frame
                        allFrames.push(...officeIframes);
                        resolve();
                    });
                });
            });
            Promise.all(pendingPromises).then(() => {
                sendResponse({ frames: allFrames });
            });
        });
        return true; // Holds open for asynchronous response
    }
});