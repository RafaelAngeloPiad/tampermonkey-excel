// ==UserScript==
// @name         New Userscript
// @namespace    http://tampermonkey.net/
// @version      2024-07-31
// @description  Try to take over the world!
// @author       You
// @match        http://127.0.0.1:5500/index.html
// @icon         https://www.google.com/s2/favicons?sz=64&domain=0.1
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // Function to load an external script
    function loadScript(url, callback) {
        let script = document.createElement('script');
        script.type = 'text/javascript';
        script.src = url;
        script.onload = callback;
        document.head.appendChild(script);
    }

    // Function to convert Excel serial date number to yyyy-MM-dd format (Mac 1904 system)
    function convertExcelDate(serial) {
        if (typeof serial === 'number') {
            const epoch = new Date(1904, 0, 1); // Mac's epoch date
            const date = new Date(epoch.getTime() + serial * 86400000);
            return date.toISOString().split('T')[0]; // Format as YYYY-MM-DD
        }
        return serial;
    }

    // Function to normalize date formats
    function normalizeDate(date) {
        if (typeof date === 'number') {
            return convertExcelDate(date);
        } else {
            return null;
        }
    }

    // Load the xlsx library
    loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', function() {
        console.log('xlsx library loaded');

        // Create floating menu
        let menu = document.createElement('div');
        menu.style.position = 'fixed';
        menu.style.top = '10px';
        menu.style.right = '10px';
        menu.style.backgroundColor = 'white';
        menu.style.border = '1px solid black';
        menu.style.padding = '10px';
        menu.style.zIndex = '10000';

        let fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.accept = '.xlsx';
        menu.appendChild(fileInput);

        // Create a line break element
        let lineBreak = document.createElement('br');
        // Append the line break element to the menu
        menu.appendChild(lineBreak);

        let addButton = document.createElement('button');
        addButton.textContent = 'Add';
        menu.appendChild(addButton);

        let prevButton = document.createElement('button');
        prevButton.textContent = 'Prev';
        menu.appendChild(prevButton);

        let nextButton = document.createElement('button');
        nextButton.textContent = 'Next';
        menu.appendChild(nextButton);

        // Add an input to show current index
        let indexDisplay = document.createElement('input');
        indexDisplay.type = 'number';
        indexDisplay.min = '0';
        indexDisplay.style.marginTop = '10px';
        indexDisplay.style.width = '90px';
        menu.appendChild(indexDisplay);

         // add spaces
        let whitespace = document.createTextNode('       ');
        menu.appendChild(whitespace);

        // Add alert button
        let alertButton = document.createElement('button');
        alertButton.textContent = 'Guide';
        menu.appendChild(alertButton);

        // Create a line break element
        let lineBreak2 = document.createElement('br');
        // Append the line break element to the menu
        menu.appendChild(lineBreak2);

        // Add a text area to show data
        let dataDisplay = document.createElement('textarea');
        dataDisplay.style.width = '300px';
        dataDisplay.style.height = '150px';
        dataDisplay.style.marginTop = '10px';
        menu.appendChild(dataDisplay);

        document.body.appendChild(menu);

        let data = [];
        let currentIndex = 0;

        // Handle file upload
        fileInput.addEventListener('change', (e) => {
            let file = e.target.files[0];
            let reader = new FileReader();

            reader.onload = function(event) {
                let arrayBuffer = event.target.result;
                let workbook = XLSX.read(arrayBuffer, {type: 'array'});

                let sheetName = workbook.SheetNames[0];
                let sheet = workbook.Sheets[sheetName];
                let rows = XLSX.utils.sheet_to_json(sheet, {header: 1});

                // Convert rows to an array of objects with column indices
                data = rows.map(row => ({
                    firstName: row[0] || '',
                    middleName: row[1] || '',
                    lastName: row[2] || '',
                    ext: row[3] || '',
                    dob: normalizeDate(row[4]) || '',
                    brgy: row[5] || '',
                    cityOrMunicipality: row[6] || '',
                    province: row[7] || '',
                    district: row[8] || '',
                    idType: row[9] || '',
                    idNumber: row[10] || '',
                    contactNumber: row[11] || '',
                    epayment: row[12] || '',
                    beneficiarytpe: row[13] || '',
                    occupation: row[14] || '',
                    sex: row[15] || '',
                    civilStatus: row[16] || '',
                    age: row[17] || '',
                    aveMonthIncome: row[18] || '',
                    dependent: row[19] || '',
                    forPBeneficiary: row[20] || '',
                    interestedwage: row[21] || '',
                    skillTrainNeed: row[22] || '',
                }));

                console.log('Parsed data:', data);
                displayData(data[currentIndex]); // Display the first record
                updateIndexDisplay();
            };

            reader.readAsArrayBuffer(file);
        });

        // Function to display data in the text area
        function displayData(rowData) {
            dataDisplay.value = JSON.stringify(rowData, null, 2);
        }

        // Function to update the index display
        function updateIndexDisplay() {
            indexDisplay.value = currentIndex;
        }

        // Add button functionality
        addButton.addEventListener('click', () => {
            if (data.length > 0) {
                populateForm(data[currentIndex]);
            }
        });

        // Prev button functionality
        prevButton.addEventListener('click', () => {
            if (data.length > 0) {
                currentIndex = (currentIndex - 1 + data.length) % data.length;
                displayData(data[currentIndex]);
                updateIndexDisplay();
            }
        });

        // Next button functionality
        nextButton.addEventListener('click', () => {
            if (data.length > 0) {
                currentIndex = (currentIndex + 1) % data.length;
                displayData(data[currentIndex]);
                updateIndexDisplay();
            }
        });

        // Handle custom index input
        indexDisplay.addEventListener('input', () => {
            let newIndex = parseInt(indexDisplay.value, 10);
            if (!isNaN(newIndex) && newIndex >= 0 && newIndex < data.length) {
                currentIndex = newIndex;
                displayData(data[currentIndex]);
            } else {
                if (newIndex < 0) {
                    currentIndex = 0;
                } else if (newIndex >= data.length) {
                    currentIndex = data.length - 1;
                }
                displayData(data[currentIndex]);
                updateIndexDisplay();
            }
        });

        // Alert button functionality
        alertButton.addEventListener('click', () => {
            alert(`1. Please convert the Excel cell format to text and the dates to date yyyy-mm-dd but first what i do is copy paste the dates to a notepad and paste them back again to fix some format issues then proceed with the date format cell

2.data = rows.map(row => ({
                    firstName: row[0] || '',
                    middleName: row[1] || '',
                    lastName: row[2] || '',
                    ext: row[3] || '',
                    dob: normalizeDate(row[4]) || '',
                    brgy: row[5] || '',
                    cityOrMunicipality: row[6] || '',
                    province: row[7] || '',
                    district: row[8] || '',
                    idType: row[9] || '',
                    idNumber: row[10] || '',
                    contactNumber: row[11] || '',
                    epayment: row[12] || '',
                    beneficiarytpe: row[13] || '',
                    occupation: row[14] || '',
                    sex: row[15] || '',
                    civilStatus: row[16] || '',
                    age: row[17] || '',
                    aveMonthIncome: row[18] || '',
                    dependent: row[19] || '',
                    forPBeneficiary: row[20] || '',
                    interestedwage: row[21] || '',
                    skillTrainNeed: row[22] || '',
                }));

                  `);
        });

        function populateForm(rowData) {
            document.getElementById('firstName').value = rowData.firstName || '';
            document.getElementById('middleName').value = rowData.middleName || '';
            document.getElementById('lastName').value = rowData.lastName || '';
            document.getElementById('dob').value = rowData.dob || '';
            document.getElementById('brgy').value = rowData.brgy || '';
            document.getElementById('idType').value = rowData.idType || '';
            document.getElementById('idNumber').value = rowData.idNumber || '';
            document.getElementById('contactNumber').value = rowData.contactNumber || '';
            document.getElementById('occupation').value = rowData.occupation || '';
            document.getElementById('sex').value = rowData.sex || '';
            document.getElementById('civilStatus').value = rowData.civilStatus || '';
            document.getElementById('dependent').value = rowData.dependent || '';
        }
    });
})();
