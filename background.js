// Create a queue class that you can queue.add(() => {...resolve();}) so that each task is resolved before starting the next in queue 
class PromiseQueue {
    constructor() {
        this.queue = [];
        this.processing = false;
    }
    async add(task) {
        return new Promise((resolve, reject) => {
            this.queue.push({ task, resolve, reject });
            this.process();
        });
    }
    async process() {
        if (this.processing || this.queue.length === 0) {
            return;
        }
        this.processing = true;
        const { task, resolve, reject } = this.queue.shift();
        try {
            const result = await task();
            resolve(result);
        } catch (error) {
            reject(error);
        } finally {
            this.processing = false;
            this.process();
        }
    }
}


// Create a queue for notifications to make sure they are not sent too quickly
let notifQueue = new PromiseQueue();
const notifDelay = 1000; // 1 second


// Listen for notification requests from content scripts
chrome.runtime.onMessage.addListener(function(message, sender, sendResponse) {
    if (message.type === 'xlpbNotif') {
        notifQueue.add(() => new Promise((resolve) => {
            chrome.notifications.create(`notif_${Date.now()}`, { // Create and send the notification
                type: 'basic',
                iconUrl: 'custom_excel_suppressor.png',
                title: 'Excel Suppressor Notification',
                message: message.text
            });
            setTimeout(resolve, notifDelay); // Call resolve to finish this task after notifDelay has elapsed
        }));
    }
});


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