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


// ALLOW ACCESS TO MICROSOFT 365 ACCOUNT
// document.evaluate("//span[contains(text(), 'Allow access to Microsoft 365 account')]", document, null, XPathelem.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
// <span id="8gdli5-Title" class="messageTitle-365">Allow access to Microsoft 365 account</span>
// nearestXPath(elem, ".//button[@aria-label='Close']").click()

// TRUST WORKBOOK LINKS?
// searchXPath(document, "//*[contains(text(), 'Trust workbook links')]")
// nearestXPath(elem, ".//button[@aria-label='Close']").click()

// UNABLE TO REFRESH
// searchXPath(document, "//*[contains(text(), 'UNABLE TO REFRESH')]")
// nearestXPath(elem, ".//button[@aria-label='Close']").click()

// STAY SIGNED IN (original)
// "//span[contains(text(), 'Stay signed in')]"
// click the closest button

// STAY SIGNED IN (new? with more lenient button xpath)
// "//*[contains(text(), 'Your session is about to expire')]"
// nearestXPath(elem, ".//button[contains(@aria-label, 'Stay signed in') or contains(text(), 'Stay signed in')]")

// Excel shortcuts enabled
// "//label[contains(text(), 'Excel shortcuts enabled')]"
// nearestXPath(elem, ".//button[@aria-label='Got it']").click()

// EDITING SESSION IN PROGRESS
// elem = searchXPath(document, "//div[contains(text(), 'Editing session in progress')]")
// nearestXPath(elem, ".//button[@aria-label='Yes']").click()

// SESSION EXPIRED (Sign In Again)
// elem = searchXPath(document, "//span[contains(text(), 'Please sign in again to continue working')]")
// nearestXPath(elem, ".//button[@aria-label='Sign in again']").click()



// 1-3: NEED MORE WORKING SPACE?
// elem = searchXPath(document, "//label[contains(text(), 'Need more working space')]")
// nearestXPath(elem, ".//button[@aria-label='Close']").click()

// 1-5: DISABLE CTRL MENU
// elem = searchXPath(document, "//div[@id='PasteRecoveryUIGroup']")
// nearestXPath(elem, ".//div[@id='ReactFloatie']") OR okay to just remove PasteRecoveryUIGroup

// 2-1: SOMEONE HAS THIS WORKBOOK LOCKED
// elem = searchXPath(document, "//*[contains(text(), 'Someone has this workbook locked')]")
// nearestXPath(elem, ".//button[@aria-label='Continue in reading view']")

// 2-2: FROZEN ROWS WON'T SCROLL
// elem = searchXPath(document, "//div[contains(text(), 'Frozen rows won')]")
// nearestXPath(elem, ".//button[@aria-label='OK']").click()









// Define the observer globally so it can be managed
let observer = null;



// Run the functions to start the observer
const iframePattern = new RegExp('https://.*\\.officeapps\\.live\\.com');
// if (document.location.origin.match('https://usc-excel.officeapps.live.com'))
if (document.location.origin.match(iframePattern))
{
    start_excel_observer();




    // Add listener to re-initialize the excel_observer
    chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
        if (request.action === "refreshMemory") {
            console.log("Refreshing values...");
            initXlObserver();
            sendResponse({ status: "success" });  // Respond back if needed
        }
        return true;
    });



}







function initXlObserver() {
    // Add all rows to popup.html (after syncing defaults.json and local memory)
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
        // Update chrome local storage if needed
        if (localModified) {
            chrome.storage.local.set({ xlpbMem: localMem }, () => {
                if (chrome.runtime.lastError) {
                    console.error("Error saving to local storage:", chrome.runtime.lastError);
                }
            });
        }
        // Order keys numerically upwards
        let sortedKeys = Object.keys(localMem).sort((a, b) => {
            let numA = parseInt(a.split("-")[1], 10);
            let numB = parseInt(b.split("-")[1], 10);
            return numA - numB;
        });
        // Populate popup.html with all items from the current updated memory
        for (let key of sortedKeys) {
            // console.log(key)
            // console.log(localMem[key])
            appendRow(localMem[key]);
        }
        // Start the actual mutation observer
        refreshMutObserver(localMem);
    }).catch(error => {
        console.log(error)
        console.error("Error fetching data!")
    });
}


function refreshMutObserver(memory) {
    if (observer) {
        observer.disconnect();
        console.log('Disconnected previous observer.')
    }
    // Create new observer and assign it to the global var
    observer = new MutationObserver(function(mutations) {

    })
}


// Function to start the MutationObserver for the excel iframe
function start_excel_observer() {
    var observer = new MutationObserver(function(mutations) {
        mutations.forEach(function(mutation) {
            if (mutation.addedNodes.length > 0) { // If child elements were added

                // Define an array of XPaths to check and remove elements
                let rm_xpaths = [
                    // "//div[@id='fluent-default-layer-host']",
                    "//div[@class='RenamePromptCalloutHeader']",
                    "//div[contains(@aria-label, 'Selected date')]",
                    "//button[@id='KeyboardShortcutAwarenessCalloutCloseButton']/parent::*/parent::*/parent::*",
                    "//div[contains(text(), 'Editing session in progress')]",
                    "//div[contains(text(), 'Someone has this workbook locked')]",
                    "//div[contains(text(), 'Sorry, your session has expired')]"
                ];
                // Call removeElementByXPath with the array of XPaths
                removeElementByXPath(rm_xpaths);
            }
        });
    });
    observer.observe(document.body, {
        childList: true, // Only listen to when child elements are added/removed
        subtree: true // Listen to all descendants of document.body
    });
}





// Function to check for the element using XPath and remove it if found
function removeElementByXPath(xpathArray) {
    for (let xpath of xpathArray) {
        let elem = searchXPath(document, xpath);
        if (!elem) {
            continue
        }
        if (xpath.includes("Editing session in progress"))
        {
            let actionElem = searchXPath(document, "//button[@aria-label='Yes'][@id='DialogActionButton']");
            if (actionElem)
            {
                console.log("EXCEL SUPPRESSOR: Editing session in progress Dialogue has been auto-approved...");
                custom_notif("Editing in progress dialogue");
                actionElem.click();
                return true; // Indicate that the element was found and removed
            }
        }
        else if (xpath.includes("Someone has this workbook locked"))
        {
            let actionElem = searchXPath(document, "//button[@aria-label='No'][@id='DialogSecondaryActionButton']");
            if (actionElem)
            {
                console.log("EXCEL SUPPRESSOR: Someone has this workbook locked Dialogue has been auto-disapproved...");
                custom_notif("Workbook locked dialogue");
                actionElem.click();
                return true; // Indicate that the element was found and removed
            }
        }
        else if (xpath.includes("Sorry, your session has expired"))
        {
            let actionElem = searchXPath(document, "//button[@aria-label='Refresh'][@id='DialogActionButton']");
            if (actionElem)
            {
                console.log("EXCEL SUPPRESSOR: Someone has this workbook locked Dialogue has been auto-disapproved...");
                custom_notif("Workbook locked dialogue");
                actionElem.click();
                return true; // Indicate that the element was found and removed
            }
        }
        else // 
        {
            elem.remove();
            console.log("EXCEL SUPPRESSOR: Element removed for XPath: " + xpath);
            custom_notif(xpath);
            return true; // Indicate that the element was found and removed
        }
    }
    return false; // Indicate that none of the XPaths matched any element
}

