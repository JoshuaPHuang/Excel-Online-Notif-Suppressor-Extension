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

// FROZEN ROWS WON'T SCROLL
// elem = searchXPath(document, "//div[contains(text(), 'Frozen rows won')]")
// nearestXPath(elem, ".//button[@aria-label='OK']").click()

// NEED MORE WORKING SPACE?
// elem = searchXPath(document, "//label[contains(text(), 'Need more working space')]")
// nearestXPath(elem, ".//button[@aria-label='Close']").click()

// DISABLE CTRL MENU
// elem = searchXPath(document, "//div[@id='PasteRecoveryUIGroup']")
// nearestXPath(elem, ".//div[@id='ReactFloatie']") OR okay to just remove PasteRecoveryUIGroup

// SOMEONE HAS THIS WORKBOOK LOCKED
// elem = searchXPath(document, "//*[contains(text(), 'Someone has this workbook locked')]")
// nearestXPath(elem, ".//button[@aria-label='Continue in reading view']")





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

// Run the functions to start the observer
const iframePattern = new RegExp('https://.*\\.officeapps\\.live\\.com');
// if (document.location.origin.match('https://usc-excel.officeapps.live.com'))
if (document.location.origin.match(iframePattern))
{
    start_excel_observer();
}