// Function to check for the element using XPath and remove it if found
function removeElementByXPath(xpathArray) {
    for (let xpath of xpathArray) {
        let result = document.evaluate(
            xpath,
            document,
            null,
            XPathResult.FIRST_ORDERED_NODE_TYPE,
            null
        ).singleNodeValue;
        if (result) {
            if (xpath.includes("Editing session in progress"))
            {
                let temp_result = document.evaluate(
                    "//button[@aria-label='Yes'][@id='DialogActionButton']",
                    document,
                    null,
                    XPathResult.FIRST_ORDERED_NODE_TYPE,
                    null
                ).singleNodeValue;
                if (temp_result)
                {
                    console.log("EXCEL SUPPRESSOR: Editing session in progress Dialogue has been auto-approved...")
                    custom_notif("Editing in progress dialogue");
                    temp_result.click();
                    return true; // Indicate that the element was found and removed
                }
            }
            else if (xpath.includes("Someone has this workbook locked"))
            {
                let temp_result = document.evaluate(
                    "//button[@aria-label='No'][@id='DialogSecondaryActionButton']",
                    document,
                    null,
                    XPathResult.FIRST_ORDERED_NODE_TYPE,
                    null
                ).singleNodeValue;
                if (temp_result)
                {
                    console.log("EXCEL SUPPRESSOR: Someone has this workbook locked Dialogue has been auto-disapproved...")
                    custom_notif("Workbook locked dialogue");
                    temp_result.click();
                    return true; // Indicate that the element was found and removed
                }
            }
            else
            {
            result.remove();
            console.log("EXCEL SUPPRESSOR: Element removed for XPath: " + xpath);
            custom_notif(xpath)
            return true; // Indicate that the element was found and removed
            }
        }
    }
    return false; // Indicate that none of the XPaths matched any element
}

/*
// Function to start the MutationObserver for the main page
function start_main_pg_observer() {
    var observer = new MutationObserver(function(mutations) {
        mutations.forEach(function(mutation) {
            let result = document.evaluate(
                "//span[contains(text(), 'Stay signed in')]",
                document,
                null,
                XPathResult.FIRST_ORDERED_NODE_TYPE,
                null
            ).singleNodeValue;
            if (result) {
                let temp_result = result.closest('button')
                console.log(temp_result)
                if (temp_result)
                {
                    console.log("EXCEL SUPPRESSOR: 'Stay signed in' Dialogue has been auto-approved...")
                    temp_result.click();
                    // temp_result.dispatchEvent(enter_down_event)
                    return true; // Indicate that the element was found and removed
                }
            }
        });
    });
    observer.observe(document.body, {
        childList: true,
        subtree: true
    });
}
*/

// Function to start the MutationObserver for the excel iframe
function start_excel_observer() {
    var observer = new MutationObserver(function(mutations) {
        mutations.forEach(function(mutation) {
            // Define an array of XPaths to check and remove elements
            let rm_xpaths = [
                // "//div[@id='fluent-default-layer-host']",
                "//div[@class='RenamePromptCalloutHeader']",
                "//div[contains(@aria-label, 'Selected date')]",
                "//button[@id='KeyboardShortcutAwarenessCalloutCloseButton']/parent::*/parent::*/parent::*",
                "//div[contains(text(), 'Editing session in progress')]",
                "//div[contains(text(), 'Someone has this workbook locked')]",
            ];
            // Call removeElementByXPath with the array of XPaths
            removeElementByXPath(rm_xpaths);
        });
    });
    observer.observe(document.body, {
        childList: true,
        subtree: true
    });
}

// enable throttled notifications
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
if (document.location.origin.match('https://usc-excel.officeapps.live.com'))
{
    start_excel_observer();
}
/*
else if (document.location.origin.match('sharepoint.com'))
{
    start_main_pg_observer();
}
*/