// Storage for original method to apply an observer to the main page instead of simply to the iframe


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


/*
else if (document.location.origin.match('sharepoint.com'))
{
    start_main_pg_observer();
}
*/

