// Debouncer function that runs the function instantly once if only called once; sets timeout to prevent duplicate calls; calls the function at the very end of the timeout and resets behavior
function advDebounce(func, delay) {
    let timer;
    let isFirstCall = true; // Flag to track the first call
    return function(...args) {
        if (isFirstCall) {
            func.apply(this, args); // Execute immediately on the first call
            isFirstCall = false;
        } else {
            // Clear any existing timer and set a new one for subsequent calls
            clearTimeout(timer);
            timer = setTimeout(() => {
                func.apply(this, args); // Execute after the delay
                isFirstCall = true;
            }, delay);
        }
    };
}

// Function to send notifications (throttled in background.js)
function createNotif(notifTxt) {
    chrome.runtime.sendMessage( {type: 'xlpbNotif', text: notifTxt} )    
}

// Function to look for an xpath in a certain context (document for most cases)
const searchXPath = (context, xPath) => {
    return document.evaluate(xPath, context, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
}

// Function to look for the nearest instance of an xpath surrounding an element
function nearestXPath(element, xPath) {
    // Check to see that the xPath will look in subtrees
    if (!xPath.startsWith('.//')) {
        throw new Error("xPath does not start with './/'; nearestXPath will not search in subtrees, please modify the xPath to continue.")
    }
    // Check the current element's subtree and return if true
    let foundXPath = searchXPath(element, xPath);
    if (foundXPath) return foundXPath; 
    // Look upwards through the document to find the xPath
    while (element.parentElement) {
        foundXPath = searchXPath(element.parentElement, xPath);
        if (foundXPath) return foundXPath;
        element = element.parentElement; // Set element to the parentElement to continue searching upwards
    }
    return null;
}

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






// Define the observers and promise queues globally so they can be managed
let observer = null;
let readQueue = new PromiseQueue();



// Run the functions to start the observer
const iframePattern = new RegExp('^https://(usc-excel\\.officeapps\\.live\\.com|excel\\.officeapps\\.live\\.com|officeonline\\.sfcollab\\.org)/')
// if (document.location.origin.match('https://usc-excel.officeapps.live.com'))
if (document.location.origin.match(iframePattern))
{
    // document.addEventListener("DOMContentLoaded", function() {
    // start_excel_observer();
    initXlObserver();
    // Add listener to re-initialize the excel_observer
    chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
        if (request.action === "refreshMemory") {
            console.log("Refreshing values...");
            // initXlObserver();
            readQueue.add(() => initXlObserver());
            sendResponse({ status: "success" });  // Respond back immediately to indicate receipt; no need to hold as long as this side is scheduled to run
        }
        return true;
    });
    // });
}






function initXlObserver() {
    return new Promise((resolve, reject) => {
        // Regenerate local memory by syncing defaults.json and local memory)
        let defaultMemPromise = fetch(chrome.runtime.getURL("defaults.json")).then(response => response.json())// Get defaults dict from defaults.json file
        let localMemPromise = new Promise((resolve) => { // Get the local memory from chrome local storage
            chrome.storage.local.get(["xlpbMem"], (result) => {
                resolve(result.xlpbMem || {});
            });
        });
        // Wait until both local and default memories are ready
        Promise.all([localMemPromise, defaultMemPromise]).then(([localMem, defaultMem]) => {
            // Create list of local names with state=true
            let trueNameArr = [];
            for (let key in localMem) {
                if (localMem[key].state === true) {
                    trueNameArr.push(localMem[key].name);
                }
            }
            // Compare chrome local storage non-appended rows with defaults and update any rows which have differences (set state to false when updating)
            let localModified = false;
            let keySet = new Set([ // Get all keys to compare
                ...Object.keys(localMem).filter(key => localMem[key].appended === false),
                ...Object.keys(defaultMem),
            ]);
            keySet.forEach(key => {
                let localItem = localMem[key];
                let defaultItem = defaultMem[key];   
                if (!localItem || JSON.stringify(localItem) !== JSON.stringify(defaultItem)) {
                    localMem[key] = defaultItem; // Overwrite
                    localModified = true;
                }
            });
            // Set all local names that had state=true back to state=true
            for (let key in localMem) {
                if (trueNameArr.includes(localMem[key].name)) {
                    localMem[key].state = true;
                }
            }
            // Filter out all entries which have state: false
            localMem = Object.fromEntries(Object.entries(localMem).filter(([key, value]) => value.state === true));
            // Start the actual mutation observer with the new data
            refreshMutObserver(localMem);
            resolve();
        }).catch(error => {
            console.log(error);
            console.error("Error fetching data!");
            reject(error);
        });
    });
}


function refreshMutObserver(newMem) {
    if (observer) {
        observer.disconnect();
        console.log('Disconnected previous observer (attempt type 2).')
    }
    // Create new observer and assign it to the global var
    observer = new MutationObserver(function(mutations) {
        // Check to see if there were any nodes added
        let addedNodesFlag = false;
        mutations.forEach(function(mutation) {
            if (mutation.addedNodes.length > 0) { // If child elements were added
                addedNodesFlag = true;
            }
        });
        if (addedNodesFlag) {
            suppressor(newMem);
            // debouncedSuppressor(newMem);
        }
    })
    observer.observe(document.body, {
        childList: true, // Only listen to when child elements are added/removed
        subtree: true // Listen to all descendants of document.body
    });
}

const debouncedSuppressor = advDebounce(function(arg) { suppressor(arg); }, 10); // 10ms

function suppressor(mem) {
    let notifFlag = false;
    // METHOD 1
    Object.keys(mem).forEach(key => {
        if (mem[key].name == "Enable Desktop Notifications") {
            notifFlag = true;
        }
    });
    // METHOD 1
    Object.values(mem).forEach(item => {
        // console.log(`Looking for ${mem[key].name}: ${mem[key].xpath}`);
        let foundElem = null;
        if (item.state !== true) {
            console.error(`Key ${item.idNumStr} .state !== true, please check the filter`);
            return;
        } else if (item.xpath == "" || item.method == "") {
            return;
        }
        foundElem = searchXPath(document, item.xpath);
        if (!foundElem) return;
        console.log(`Found element with xpath ${item.xpath}, suppressing...`)
        let suppressedFlag = false;
        if (item.method == "REMOVE") { // Remove the element entirely if remove is selected
            foundElem.remove();
            suppressedFlag = true;
        } else {
            try {
                nearestXPath(foundElem, item.method).click(); // Click on the method button to interact w/ the dialog
                suppressedFlag = true;
            } catch (error) {
                console.log(`Error trying to click ${item.method}: ${error}`);
            }
        }
        if (notifFlag && suppressedFlag) {
            createNotif(`Suppressed: ${item.name}`);
        }
    })
}