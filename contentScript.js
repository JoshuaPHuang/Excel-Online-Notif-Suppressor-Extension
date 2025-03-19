// // Debouncer function
// function debounce(func, delay) {
//     let timer;
//     return function(...args) {
//         clearTimeout(timer);
//         timer = setTimeout(() => func.apply(this, args), delay);
//     };
// }


// function advDebounce(func, delay) {
//     let timer;
//     let isFirstCall = true; // Track whether it's the first call after the timer ran out
//     return function(...args) {
//         if (isFirstCall) {
//             func.apply(this, args); // Execute immediately on the first call
//             isFirstCall = false;
//         }
//         clearTimeout(timer); // Clear any previous timer
//         timer = setTimeout(() => {
//             func.apply(this, args); // Execute after the delay
//         }, delay);
//     };
// }
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







// Enable throttled notifications
let last_notif_time = 0;
function custom_notif(notif_txt)
{
    const now_time = Date.now();
    if (now_time - last_notif_time >= 1000) {
        chrome.runtime.sendMessage({type: 'custom_notif', text: notif_txt});
        last_notif_time = now_time;
    }
};

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
const iframePattern = new RegExp('https://.*\\.officeapps\\.live\\.com');
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
            debouncedSuppressor(newMem);
        }
    })
    observer.observe(document.body, {
        childList: true, // Only listen to when child elements are added/removed
        subtree: true // Listen to all descendants of document.body
    });
}

const debouncedSuppressor = advDebounce(function(arg) { suppressor(arg); }, 100);

function suppressor(mem) {
    for (let key in mem) {
        console.log(`Looking for ${mem[key].name}: ${mem[key].xpath}`);
        let foundElem = null;
        if (mem[key].state !== true) {
            console.error(`Key ${key} .state !== true, please check the filter`);
            continue;
        }
        foundElem = searchXPath(document, mem[key].xpath);
        if (!foundElem) continue;
        console.log(`Found element with xpath ${mem[key].xpath}, suppresing...`)
        if (mem[key].method == "REMOVE") { // Remove the element entirely if remove is selected
            foundElem.remove();
        } else {
            try {
                nearestXPath(foundElem, mem[key].method).click(); // Click on the method button to interact w/ the dialog
            } catch (error) {
                console.log(`Error trying to click ${mem[key].method}: ${error}`);
            }
        }
    }
}


// // Function to start the MutationObserver for the excel iframe
// function start_excel_observer() {
//     var observer = new MutationObserver(function(mutations) {
//         mutations.forEach(function(mutation) {
//             if (mutation.addedNodes.length > 0) { // If child elements were added
//                 // Define an array of XPaths to check and remove elements
//                 let rm_xpaths = [
//                     // "//div[@id='fluent-default-layer-host']",
//                     "//div[@class='RenamePromptCalloutHeader']",
//                     "//div[contains(@aria-label, 'Selected date')]",
//                     "//button[@id='KeyboardShortcutAwarenessCalloutCloseButton']/parent::*/parent::*/parent::*",
//                     "//div[contains(text(), 'Editing session in progress')]",
//                     "//div[contains(text(), 'Someone has this workbook locked')]",
//                     "//div[contains(text(), 'Sorry, your session has expired')]"
//                 ];
//                 // Call removeElementByXPath with the array of XPaths
//                 removeElementByXPath(rm_xpaths);
//             }
//         });
//     });
//     observer.observe(document.body, {
//         childList: true, // Only listen to when child elements are added/removed
//         subtree: true // Listen to all descendants of document.body
//     });
// }



// // Function to check for the element using XPath and remove it if found
// function removeElementByXPath(xpathArray) {
//     for (let xpath of xpathArray) {
//         let elem = searchXPath(document, xpath);
//         if (!elem) {
//             continue
//         }
//         if (xpath.includes("Editing session in progress"))
//         {
//             let actionElem = searchXPath(document, "//button[@aria-label='Yes'][@id='DialogActionButton']");
//             if (actionElem)
//             {
//                 console.log("EXCEL SUPPRESSOR: Editing session in progress Dialogue has been auto-approved...");
//                 custom_notif("Editing in progress dialogue");
//                 actionElem.click();
//                 return true; // Indicate that the element was found and removed
//             }
//         }
//         else if (xpath.includes("Someone has this workbook locked"))
//         {
//             let actionElem = searchXPath(document, "//button[@aria-label='No'][@id='DialogSecondaryActionButton']");
//             if (actionElem)
//             {
//                 console.log("EXCEL SUPPRESSOR: Someone has this workbook locked Dialogue has been auto-disapproved...");
//                 custom_notif("Workbook locked dialogue");
//                 actionElem.click();
//                 return true; // Indicate that the element was found and removed
//             }
//         }
//         else if (xpath.includes("Sorry, your session has expired"))
//         {
//             let actionElem = searchXPath(document, "//button[@aria-label='Refresh'][@id='DialogActionButton']");
//             if (actionElem)
//             {
//                 console.log("EXCEL SUPPRESSOR: Someone has this workbook locked Dialogue has been auto-disapproved...");
//                 custom_notif("Workbook locked dialogue");
//                 actionElem.click();
//                 return true; // Indicate that the element was found and removed
//             }
//         }
//         else // 
//         {
//             elem.remove();
//             console.log("EXCEL SUPPRESSOR: Element removed for XPath: " + xpath);
//             custom_notif(xpath);
//             return true; // Indicate that the element was found and removed
//         }
//     }
//     return false; // Indicate that none of the XPaths matched any element
// }