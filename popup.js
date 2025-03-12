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
            if (elem.value == "") { // Set class to error if empty
                setErrorClass(elem);
                hasEmpty = true;
            }
        });
        if (hasEmpty) return;
        appendRow(nameInput2.value, textInput2.value, labelInput2.value, '2-8');
        showHideElems(addRow2, newRow2);
        clearInputs();
    });

    // Function to add a new row before the new-row input row
    function appendRow(nameVal, textVal, labelVal, idNumStr) {
        let container, beforeElem
        if (idNumStr[0] == "2") {
            container = container2;
            beforeElem = newRow2;
        } else if (idNumStr[0] == "3") {
            container = container3;
            beforeElem = newRow3;
        } else {
            throw new Error(`idNumStr needs to start with '2' or '3': ${idNumStr}`);
        }
        const newDiv = document.createElement('div');
        newDiv.className = "row";
        newDiv.id = `row-${idNumStr}`;
        newDiv.innerHTML = `
            <label for="checkbox-${idNumStr}">
                <input type="checkbox" id="checkbox-${idNumStr}" style="margin-right: 0.5rem">${nameVal}
                <div class="info-container">
                    <div class="hoverable">🛈</div>
                    <div class="tooltip">Pop-Up Text: '${textVal}'<br><br>Check to auto-select Button Label: '${labelVal}'.</div>
                </div>
            </label>
            <button class="delete-button" id="delete-row-${idNumStr}">&times;</button>
        `;
        container.insertBefore(newDiv, beforeElem); // Insert the element
        document.getElementById(`delete-row-${idNumStr}`).addEventListener("click", deleteRow); // Add a listener on the delete button to delete the row
    }

    // Function to delete the nearest item with class="row" to an event
    function deleteRow(event) {
        const rowToDel = event.target.closest(".row"); // Closest element with class row
        if (rowToDel) {
            rowToDel.remove();
        }
    }








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
        nameInput3.value = "";
        textInput3.value = "";
        labelInput4.value = "";
    }

});