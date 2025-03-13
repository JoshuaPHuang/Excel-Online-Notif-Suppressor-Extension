// Debouncer function to delay saves...
function debounce(func, delay) {
    let timer;
    return function(...args) {
        clearTimeout(timer);
        timer = setTimeout(() => func.apply(this, args), delay);
    };
}

// Function to generate [mm/dd/YYYY HH:MM:SS] timestamp
function timestamp() {
    const now = new Date();
    const pad = (num) => String(num).padStart(2, '0');
    const month = pad(now.getMonth() + 1); // getMonth() zero-based
    const day = pad(now.getDate());
    const year = now.getFullYear();
    const hours = pad(now.getHours());
    const minutes = pad(now.getMinutes());
    const seconds = pad(now.getSeconds());
    return `[${month}/${day}/${year} ${hours}:${minutes}:${seconds}]`;
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

// Function to show one element and hide another
function showHideElems(showElem, hideElem) {
    showElem.style.display = "block";
    hideElem.style.display = "none";
}

function resetErrorClass(elem) {
    elem.classList.remove('error');
}

function setErrorClass(elem) {
    elem.classList.add('error');
}

// MAIN FUNCTION
document.addEventListener("DOMContentLoaded", function () {
    // KEY
    // 2: Auto-Approve Pop-Ups section
    // 3: Auto-Close Pop-Ups section
    const container1 = document.getElementById('blocked-pop-ups')
    const container2 = document.getElementById('auto-approved-pop-ups')
    const container3 = document.getElementById('auto-closed-pop-ups')
    const container4 = document.getElementById('settings')
    const containerD = {"1": container1, "2": container2, "3": container3, "4": container4}


    const newRow2 = document.getElementById('new-row-2')
    const addRow2 = document.getElementById('add-row-2')
    const addButton2 = document.getElementById('add-button-2')
    const saveButton2 = document.getElementById('new-save-button-2')
    const cancelButton2 = document.getElementById('cancel-new-row-2')
    const nameInput2 = document.getElementById('new-name-2')
    const textInput2 = document.getElementById('new-text-2')
    const labelInput2 = document.getElementById('new-label-2')
    const inputsArr2 = [nameInput2, textInput2, labelInput2];

    const newRow3 = document.getElementById('new-row-3')
    const addRow3 = document.getElementById('add-row-3')
    const addButton3 = document.getElementById('add-button-3')
    const saveButton3 = document.getElementById('new-save-button-3')

    const beforeElemD = {"1": null, "2": newRow2, "3": newRow3, "4": null}

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
        // Populate popup.html with all items from the current updated memory
        for (let key in localMem) {
            // console.log(key)
            // console.log(localMem[key])
            appendRow(localMem[key]);
        }
    }).catch(error => {
        console.log(error)
        console.error("Error fetching data!")
    });








    // Set an event listener to reset the error class of an input element every time an input occurs on that element
    let inputsArr = Array.from(document.getElementsByTagName('input'));
    inputsArr.forEach(inputElem => inputElem.addEventListener("input", () => resetErrorClass(inputElem)));
    
    // ADDING NEW ROWS
    addButton2.addEventListener("click", function() { // Show new approve row input, hide "Add New Row", reset block input
        showHideElems(newRow2, addRow2);
        showHideElems(addRow3, newRow3);
        inputsArr.forEach(resetErrorClass);
    });
    cancelButton2.addEventListener("click", function() { // Cancel and hide new approve row input, show "Add New Row"
        showHideElems(addRow2, newRow2);
        clearInputs();
    });
    saveButton2.addEventListener("click", function() {
        // Check first to see if any are empty
        let hasEmpty = false;
        inputsArr2.forEach(elem => {
            if (elem.value == "") { // Set classes to error if empty
                setErrorClass(elem);
                hasEmpty = true;
            }
        });
        if (hasEmpty) return; // Return if any of them are empty
        // Get the index for the 2nd part of idNumStr
        let newIdInd = 101;
        chrome.storage.local.get(["xlpbMem"], (result) => {
            let localMem = result.xlpbMem || {};
            appendedKeys = new Set([...Object.keys(localMem).filter(key => (localMem[key].appended === true && localMem[key].section == "2"))]); // Get set of keys that are appended and are in the correct section
            let takenIndSet = new Set();
            appendedKeys.forEach(key => { // Find all currently taken indices
                let idNumParts = key.split("-");
                let idNumInd = parseInt(idNumParts[1], 10);
                takenIndSet.add(idNumInd);
            })
            while (takenIndSet.has(newIdInd)) {
                newIdInd++;
            }
            // console.log(`newIdInd:${String(newIdInd)}`)
            // Create the rowItem to append and add to local storage
            rowItem = {
                idNumStr: `2-${newIdInd}`,
                section: "2",
                name: nameInput2.value,
                state: true,
                appended: true,
                text: textInput2.value,
                label: labelInput2.value,
            };
            // Append the row to popup.html
            appendRow(rowItem);
            // Add the rowItem to local memory and write back to storage
            localMem[rowItem.idNumStr] = rowItem;
            chrome.storage.local.set({ xlpbMem: localMem }, () => {
                if (chrome.runtime.lastError) {
                    console.error("Error saving to local storage:", chrome.runtime.lastError);
                }
            });
            // Clean up
            showHideElems(addRow2, newRow2);
            clearInputs();
        });
    });

    // Function to add a new row before the new-row input row
    function appendRow(rowItem) {
        let container = containerD[rowItem.section];
        let beforeElem = beforeElemD[rowItem.section];
        const newDiv = document.createElement('div');
        newDiv.className = "row";
        newDiv.id = `row-${rowItem.idNumStr}`;
        let htmlText, htmlLabel, htmlDelete;
        if (rowItem.appended) {
            htmlText = `Pop-Up Text: '${rowItem.text}'`;
            htmlLabel = `Check to auto-select Button Label: '${rowItem.label}'.`;
            htmlDelete = `<button class="delete-button" id="delete-row-${rowItem.idNumStr}">&times;</button>`;
        } else {
            htmlText = rowItem.text;
            htmlLabel = rowItem.label;
            htmlDelete = "";
        }
        newDiv.innerHTML = `
            <label for="checkbox-${rowItem.idNumStr}">
                <input type="checkbox" id="checkbox-${rowItem.idNumStr}" style="margin-right: 0.5rem">${rowItem.name}
                <div class="info-container">
                    <div class="hoverable">🛈</div>
                    <div class="tooltip">${htmlText}<br><br>${htmlLabel}</div>
                </div>
            </label>
            ${htmlDelete}
        `;
        container.insertBefore(newDiv, beforeElem); // Insert the element
        document.getElementById(`checkbox-${rowItem.idNumStr}`).checked = rowItem.state; // Set the checked state of this new checkbox
        if (rowItem.appended) { // Add delete button if this is an appended row
            document.getElementById(`delete-row-${rowItem.idNumStr}`).addEventListener("click", deleteRow); // Add a listener on the delete button to delete the row
        }
    }

    // Function to delete the nearest item with class="row" to an event
    function deleteRow(event) {
        const rowToDel = event.target.closest(".row"); // Closest element with class row
        if (rowToDel) {
            let idNumStr = rowToDel.id.replace("row-", "");
            chrome.storage.local.get(["xlpbMem"], (result) => {
                let localMem = result.xlpbMem || {};
                // Delete the rowItem from local storage and write back
                delete localMem[idNumStr];
                chrome.storage.local.set({ xlpbMem: localMem }, () => {
                    if (chrome.runtime.lastError) {
                        console.error("Error saving to local storage:", chrome.runtime.lastError);
                    }
                });
            });
            // Remove the row from .html
            rowToDel.remove();
        }
    }




    xlpbMem = {
        "1-1": {
            idNumStr: "1-1",
            section: "1",
            name: "Calendar Pop-Up",
            state: false,
            appended: false,
            text: "This calendar blocks view of everything to the right of the current cell.<br>This pop-up appears when pressing Ctrl+; in Excel Online.",
            label: "Check to block the pop-up.",
        },
        "1-2": {
            idNumStr: "1-2",
            section: "1",
            name: "Renaming Prompt",
            state: false,
            appended: false,
            text: "'We have autosaved your file. Ready to name it?'<br>This pop-up sometimes appears when pressing Ctrl+S in Excel Online.",
            label: "Check to block the pop-up.",
        },
        "1-3": {
            idNumStr: "1-3",
            section: "1",
            name: "Need More Working Space? Pop-Up",
            state: false,
            appended: false,
            text: "'Need more working space?'<br>'Looks like your display is not optimised.'<br>'Toggle FullScreen to maximise screen space.'<br>This pop-up sometimes appears when the window is smaller.",
            label: "Check to block the pop-up.",
        },
        "1-4": {
            idNumStr: "1-4",
            section: "1",
            name: "Reserved Shortcut Pop-Up",
            state: false,
            appended: false,
            text: "'Your browser reserves some shortcuts.'<br>'Find the list of alternatives here.'<br>This pop-up sometimes appears when using hotkeys in Excel Online.",
            label: "Check to block the pop-up.",
        },
        "1-5": {
            idNumStr: "1-5",
            section: "1",
            name: "Ctrl Paste Menu",
            state: false,
            appended: false,
            text: "Paste options menu that opens when pressing Ctrl.<br>This pop-up appears after pasting.",
            label: "Check to block the pop-up.",
        },
        "2-1": {
            idNumStr: "2-1",
            section: "2",
            name: "Workbook Locked",
            state: false,
            appended: false,
            text: "'Someone has this workbook locked'<br>'_ has locked this file for editing. This may happen if they're using a non-subscription version of Excel. Ask them to close the file or check it in.'",
            label: "Check to auto-select 'Continue in reading view'.",
        },
        "2-2": {
            idNumStr: "2-2",
            section: "2",
            name: "Frozen Rows Won't Scroll",
            state: false,
            appended: false,
            text: "'Frozen rows won't scroll'<br>'The visible area of your grid contains only frozen rows, which will not scroll. To scroll, unfreeze the rows, change the size of the Excel window, or zoom out.'<br>This pop-up appears when trying to scroll on a screen with no unfrozen cells visible.",
            label: "Check to auto-select 'OK'.",
        },
        "2-3": {
            idNumStr: "2-3",
            section: "2",
            name: "Editing Session in Progress",
            state: false,
            appended: false,
            text: "'Editing session in progress'<br>'Other people are already editing this workbook. You can join their editing session, but you will lose any changes you have made so far. Would you like to join the existing editing session anyway?'",
            label: "Check to auto-select 'Yes'.",
        },
        "2-4": {
            idNumStr: "2-4",
            section: "2",
            name: "Session About to Expire",
            state: false,
            appended: false,
            text: "'Your session is about to expire'<br>'Your organization's policy enforces automatic sign out after a period of inactivity on Office 365 web applications.'<br>'Do you want to stay signed in?'",
            label: "Check to auto-select 'Stay signed in'.",
        },
        "2-5": {
            idNumStr: "2-5",
            section: "2",
            name: "Session Expired",
            state: false,
            appended: false,
            text: "'Sorry, your session has expired. Please refresh the page to continue.'",
            label: "Check to auto-select 'Refresh'.",
        },
        "3-1": {
            idNumStr: "3-1",
            section: "3",
            name: "Can't Edit Workbook Dialog",
            state: false,
            appended: false,
            text: "'Can't Edit Workbook'<br>'Someone has this workbook locked'<br>'We're sorry. We couldn't lock this file for editing. Would you like to try again?'<br>This pop-up sometimes appears when trying to change to Editing mode.",
            label: "Check to auto-select 'No'.",
        },
        "3-2": {
            idNumStr: "3-2",
            section: "3",
            name: "Trust Workbook Links Banner",
            state: false,
            appended: false,
            text: "'Trust workbook links? This workbook links to data in external workbooks.'<br>This pop-up appears when external workbooks are referenced.",
            label: "Check to auto-close (ignore and close the dialog).",
        },
        "3-3": {
            idNumStr: "3-3",
            section: "3",
            name: "Unable to Refresh Links Banner",
            state: false,
            appended: false,
            text: "'UNABLE TO REFRESH. We couldn't get updated values from a linked workbook.'<br>This pop-up sometimes appears when external workbooks cannot be reached.",
            label: "Check to auto-close (ignore and close the dialog).",
        },
        "3-4": {
            idNumStr: "3-4",
            section: "3",
            name: "Microsoft 365 Access Banner",
            state: false,
            appended: false,
            text: "'Allow access to Microsoft 365 account. Your browser settings are preventing an optimal experience with Microsoft 365. Allow access to improve your experience.'",
            label: "Check to auto-close (ignore and close the dialog).",
        },
        "4-1": {
            idNumStr: "4-1",
            section: "4",
            name: "Enable Notifications",
            state: false,
            appended: false,
            text: "",
            label: "",
        },
    };





    // NEED TO STORE
    // Name of the option
    // idNumStr of the option (1-1, etc.)
    // Current state of the option
    // Pop-Up Text of the option (only for appended)
    // Button Label of the option (only for appended)
    // Boolean of whether or not this option is appended









    // Function to clear the contents of all name/text/label inputs
    function clearInputs() {
        nameInput2.value = "";
        textInput2.value = "";
        labelInput2.value = "";
        // nameInput3.value = "";
        // textInput3.value = "";
        // labelInput4.value = "";
    }

});