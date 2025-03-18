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

    // Create a queue class to hold read and write functions to local storage
    const memQueue = new PromiseQueue();

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
    }).catch(error => {
        console.log(error)
        console.error("Error fetching data!")
    });






    // Set an event listener to reset the error class of an input element every time an input occurs on that element
    let inputsArr = Array.from(document.getElementsByTagName('input'));
    inputsArr.forEach(inputElem => inputElem.addEventListener("input", () => resetErrorClass(inputElem)));
    
    // ADDING NEW ROWS
    addButton2.addEventListener("click", function() { // Show new approve row input, hide "Add New Row", reset block input
        refreshCsMem();
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
        return memQueue.add(() => {
            return new Promise((resolve, reject) => {
                chrome.storage.local.get(["xlpbMem"], (result) => {
                    let newIdInd = 101;
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
                        xpath: `//*[contains(text(), '${textInput2.value}')]`,
                        method: `//button[contains(@aria-label, '${labelInput2}')]`,
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
                    resolve();
                });
            })
        })
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
            return memQueue.add(() => {
                return new Promise((resolve, reject) => {
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
                    resolve();
                })
            })
        }
    }

    // Function to send a refresh signal to the contentScript on each Excel iframe so it re-reads the current contents of local memory
    function refreshCsMem() {
        return memQueue.add(() => {
            return new Promise ((resolve, reject) => {
                chrome.runtime.sendMessage({ action: "findIframes" }, (response) => {
                    if (!response || !response.frames || response.frames.length == 0) {
                        // console.error("No iframes found."); // Just stop; it's fine if there are no Excel iframes open
                        resolve();
                    }
                    // console.log(response)
                    response.frames.forEach(frame => {
                        console.log(`Found iframe: Frame ID = ${frame.frameId}, URL = ${frame.url}, customTabId = ${frame.customTabId}`);
                        chrome.tabs.sendMessage(frame.customTabId, { action: "refreshMemory" }, { frameId: frame.frameId }, (response) => {
                            if (chrome.runtime.lastError) {
                                console.log(`Error sending message to frame ${frame.frameId}:`, chrome.runtime.lastError.message);
                            } else {
                                console.log(`Response from iframe ${frame.frameId}:`, response);
                            }
                            resolve();
                        });
                    });
                });
            });
        });
    }










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