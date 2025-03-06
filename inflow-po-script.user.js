// ==UserScript==
// @name         InFlow Master Script (All Enhancements)
// @namespace    http://yourdomain.com
// @version      3.0.0
// @description  Combined script with PO Enhancements, Auto-Assign, Product Options Remover, and Search Automation
// @match        https://app.inflowinventory.com/*
// @grant        GM_xmlhttpRequest
// @connect      docs.google.com
// @connect      googleusercontent.com
// @downloadURL  https://raw.githubusercontent.com/HemDog/InFlow/main/inflow-master-script.user.js
// @updateURL    https://raw.githubusercontent.com/HemDog/InFlow/main/inflow-master-script.user.js
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    // ====================================
    // Log Helper - Used by both scripts
    // ====================================
    function log(message) {
        console.log(`[InFlow Script] ${message}`);
    }

    // ====================================
    // PART 1: PO ENHANCEMENTS & AUTO-ASSIGN
    // ====================================

    // -----------------------------
    // Global flags and variables
    // -----------------------------
    let creditHoldPopupDismissed = false;
    let lastCardCustomerName = "";
    let inFlightCustomerName = "";
    let autoAssigned = false; // This flag will be set if the current sales order already has an assignment

    // Flag to disable quote page behavior after conversion
    let quoteConversionHappened = false;
    document.addEventListener('click', function(e) {
        let convertBtn = e.target.closest('#convertSalesOrder');
        if (convertBtn) {
            quoteConversionHappened = true;
            log("Convert to order button clicked. Disabling quote page behavior and reloading page shortly.");
            // Delay reload to allow conversion to complete
            setTimeout(() => {
                window.location.reload();
            }, 2000);
        }
    });

    // -----------------------------
    // URLs for your published CSVs
    // -----------------------------
    const sheetUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ-iA8XppRLS2s-30ElJSdvNbrgOQB-kaYIS39uC1TwvHYS4HgmZKMbSR1vUuekeiYHlQaYsFuJk3P9/pub?output=csv";
    const cardSheetUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSyRPtuexxvaAtYT6t0aTJ-I8f4vYIvd5TxkY-HU6XeTSIltadeXH6Y5N0L97lvHJidpoROBSXPob1Z/pub?output=csv";

    // -----------------------------
    // 1) Custom parser for MM/YY
    // -----------------------------
    function parseExpirationDate(expStr) {
        let match = expStr.match(/^(\d{1,2})\/(\d{2})$/);
        if (!match) {
            return new Date(expStr);
        }
        let mm = parseInt(match[1], 10);
        let yy = parseInt(match[2], 10);
        let pivot = 50;
        let fullYear = (yy < pivot) ? (2000 + yy) : (1900 + yy);
        let nextMonth = (mm === 12) ? 1 : mm + 1;
        let nextYear = (mm === 12) ? (fullYear + 1) : fullYear;
        let firstOfNext = new Date(nextYear, nextMonth - 1, 1);
        return new Date(firstOfNext.getTime() - 86400000);
    }

    // -----------------------------
    // 2) CSV Parsing Helper
    // -----------------------------
    function parseCSV(csvText) {
        let lines = csvText.trim().split('\n');
        let headers = lines[0].split(',').map(h => h.replace(/^\uFEFF/, '').trim());
        let result = [];
        for (let i = 1; i < lines.length; i++) {
            let obj = {};
            let row = lines[i].split(',');
            headers.forEach((h, index) => {
                obj[h] = (row[index] || "").trim();
            });
            result.push(obj);
        }
        return result;
    }

    // -----------------------------
    // 3) Insert Blanket PO
    // -----------------------------
    function tryInsertPO() {
        if (!window.location.href.includes('/sales-orders/')) return;
        let customerField = document.querySelector("#so_customer input[type='text']");
        let poField = document.querySelector('textarea[name="remarks"]');
        if (!customerField || !poField) return;
        let customerName = customerField.value.trim();
        if (!customerName) return;

        GM_xmlhttpRequest({
            method: "GET",
            url: sheetUrl,
            onload: function(response) {
                if (response.status !== 200) {
                    console.error("Failed to fetch Blanket PO CSV. Status:", response.status, response.statusText);
                    return;
                }
                let data = parseCSV(response.responseText);
                let row = data.find(r => {
                    let csvName = (r["Customer Name"] || r["Customer_Name"] || r["Customer"] || "").trim().toLowerCase();
                    return csvName === customerName.toLowerCase();
                });
                if (!row) {
                    log("No matching customer found in the Blanket PO sheet for: " + customerName);
                    return;
                }
                let blanketPO = row["Blanket PO"] || "";
                if (blanketPO) {
                    let formattedPO = `(BLANKET PO#:${blanketPO})`;
                    if (poField.value.includes("BLANKET PO#")) {
                        log(`"BLANKET PO#" already found in remarks. Not modifying text.`);
                    } else {
                        poField.value = formattedPO + " " + poField.value;
                        log(`Prepended PO for ${customerName} with ${formattedPO}`);
                    }
                } else {
                    log(`No Blanket PO found for ${customerName}`);
                }
            },
            onerror: function(err) {
                console.error("Error fetching Blanket PO CSV:", err);
            }
        });
    }

    // -----------------------------
    // 4) Insert the Card-on-file row
    // -----------------------------
    function insertCardOnFileRow(status) {
        const existingRow = document.getElementById('card-on-file-row');
        if (existingRow) existingRow.remove();

        let labelText = "No card";
        let bgColor = "grey";
        if (status === "on-file") {
            labelText = "On file";
            bgColor = "green";
        } else if (status === "expired") {
            labelText = "Expired";
            bgColor = "red";
        }

        const container = document.createElement('div');
        container.id = 'card-on-file-row';
        container.innerHTML = `
          <div class="sc-a3cffb35-0 ifEDhR">
            <dl class="sc-a3cffb35-1 Jiqon">
              <dt class="sc-a3cffb35-3 cDizGq" style="margin-bottom: auto; margin-top: 4px;">
                <p class="sc-a3cffb35-4 botIdl">Card on file</p>
              </dt>
              <dd class="sc-a3cffb35-6 kXBFJe">
                <span id="card-on-file-indicator"
                      style="background-color: ${bgColor}; color: white; padding: 2px 6px; border-radius: 4px; font-weight: bold;">
                  ${labelText}
                </span>
              </dd>
            </dl>
            <div class="sc-a3cffb35-2 iqSCmF"></div>
          </div>
        `;

        const shippingAddressDiv = document.querySelector('div.sc-8cb329e6-8.eLvnWB');
        if (!shippingAddressDiv) {
            console.error("Could not find shipping address container to insert card on file row.");
            return;
        }
        const nextSibling = shippingAddressDiv.nextElementSibling;
        if (!nextSibling) {
            console.error("Could not find Payment Terms container to insert card on file row.");
            return;
        }
        shippingAddressDiv.parentNode.insertBefore(container, nextSibling);
    }

    // -----------------------------
    // 5) Card On File Indicator (logic)
    // -----------------------------
    function addCardOnFileIndicator() {
        let customerField = document.querySelector("#so_customer input[type='text']");
        if (!customerField) return;
        let customerName = customerField.value.trim();
        if (!customerName) return;

        if (customerName === lastCardCustomerName) return;
        if (inFlightCustomerName === customerName) {
            log("A card on file request is already in flight for this customer. Skipping...");
            return;
        }

        inFlightCustomerName = customerName;
        GM_xmlhttpRequest({
            method: "GET",
            url: cardSheetUrl,
            onload: function(response) {
                inFlightCustomerName = "";
                if (response.status !== 200) {
                    console.error("Failed to fetch Card On File CSV. Status:", response.status, response.statusText);
                    return;
                }
                let data = parseCSV(response.responseText);
                if (!data || data.length === 0) {
                    console.error("No data found in Card On File CSV. Treating all as no card.");
                    insertCardOnFileRow("no-card");
                    lastCardCustomerName = customerName;
                    return;
                }

                let row = data.find(r => {
                    let csvName = (r["Customer Name"] || r["Customer_Name"] || r["Customer"] || "").trim().toLowerCase();
                    return csvName === customerName.toLowerCase();
                });

                if (!row) {
                    insertCardOnFileRow("no-card");
                    lastCardCustomerName = customerName;
                    return;
                }

                let expDateStr = row["Experation Date"] || row["Expiration Date"];
                let isExpired = false;
                if (expDateStr) {
                    let expDate = parseExpirationDate(expDateStr);
                    let now = new Date();
                    if (expDate.toString() === "Invalid Date") {
                        console.error("Invalid date format for expiration date:", expDateStr);
                        isExpired = true;
                    } else if (expDate < now) {
                        isExpired = true;
                    }
                }
                if (isExpired) {
                    insertCardOnFileRow("expired");
                } else {
                    insertCardOnFileRow("on-file");
                }
                lastCardCustomerName = customerName;
            },
            onerror: function(err) {
                inFlightCustomerName = "";
                console.error("Error fetching Card On File CSV:", err);
            }
        });
    }

    // -----------------------------
    // 6) Rename Elements on Sales Orders
    // -----------------------------
    function renamePOElements() {
        if (!window.location.href.includes('/sales-orders/')) return;
        document.querySelectorAll('p').forEach(p => {
            if (p.textContent.trim() === "PO #") {
                p.textContent = "Look Up Number";
                log('Renamed "PO #" to "Look Up Number"');
            }
            if (p.textContent.trim() === "Include shipping") {
                p.textContent = "Fulfill/Pick";
                log('Renamed "Include shipping" to "Fulfill/Pick"');
            }
            if (p.textContent.trim() === "Assigned To") {
                p.textContent = "Converted/Made by";
                log('Renamed "Assigned To" to "Converted/Made by"');
            }
        });
    }

    // -----------------------------
    // 7) Remove Fulfill Button
    // -----------------------------
    function removeFulfillButton() {
        let fulfillButton = document.getElementById("so-btn-Fulfill");
        if (fulfillButton) {
            fulfillButton.remove();
            log("Fulfill button removed.");
        }
    }

    // -----------------------------
    // 8) Credit Hold Check
    // -----------------------------
    function checkCreditHold() {
        let creditHoldElement = document.querySelector('input[title="CREDIT HOLD"]');
        if (!creditHoldElement) {
            creditHoldPopupDismissed = false;
            return;
        }
        let printButton = document.getElementById("sales-order-print-dropdown");
        if (printButton) {
            printButton.remove();
            log("Print button removed because the customer is on credit hold.");
        }
        let emailButton = document.getElementById("so-email-dropdown");
        if (emailButton) {
            emailButton.remove();
            log("Email button removed because the customer is on credit hold.");
        }
        if (!creditHoldPopupDismissed && !document.getElementById("credit-hold-popup")) {
            let overlay = document.createElement('div');
            overlay.id = "credit-hold-popup";
            overlay.style.position = "fixed";
            overlay.style.top = "0";
            overlay.style.left = "0";
            overlay.style.width = "100%";
            overlay.style.height = "100%";
            overlay.style.backgroundColor = "rgba(0, 0, 0, 0.5)";
            overlay.style.display = "flex";
            overlay.style.alignItems = "center";
            overlay.style.justifyContent = "center";
            overlay.style.zIndex = "10000";

            let popup = document.createElement('div');
            popup.style.backgroundColor = "white";
            popup.style.padding = "20px";
            popup.style.borderRadius = "8px";
            popup.style.textAlign = "center";
            popup.style.boxShadow = "0 0 10px rgba(0,0,0,0.25)";
            popup.style.maxWidth = "90%";

            let title = document.createElement('h1');
            title.textContent = "This customer is on credit hold!";
            title.style.fontSize = "24px";
            title.style.fontWeight = "bold";
            title.style.margin = "0 0 10px 0";

            let message = document.createElement('p');
            message.textContent = "Overdue balance must be paid in full before order can be fulfilled.";
            message.style.margin = "0";

            let closeButton = document.createElement('button');
            closeButton.textContent = "Close";
            closeButton.style.marginTop = "15px";
            closeButton.addEventListener('click', function() {
                creditHoldPopupDismissed = true;
                overlay.remove();
            });

            popup.appendChild(title);
            popup.appendChild(message);
            popup.appendChild(closeButton);
            overlay.appendChild(popup);
            document.body.appendChild(overlay);
        }
    }

    // -----------------------------
    // 9) Handle Quote Page
    // -----------------------------
    function handleQuotePage() {
        // Only apply quote behavior if conversion has not occurred.
        if (quoteConversionHappened) return;

        let convertButton = document.getElementById("convertSalesOrder");
        if (convertButton) {
            // Quote page behavior:
            document.querySelectorAll('p.sc-a3cffb35-4.botIdl').forEach(elem => {
                if (elem.textContent.trim() === "Sales rep") {
                    elem.textContent = "Quote by";
                    log("Renamed Sales rep to Quote by.");
                }
            });
            let container = document.querySelector("div.sc-3e2e0cb2-0.gLiEEA");
            if (container) {
                let sections = container.querySelectorAll("div.sc-3e2e0cb2-1.gtsOkx");
                if (sections.length >= 2) {
                    sections[1].remove();
                    log("Removed the second (assigned-to) section for quotes.");
                }
            }
        }
    }

    // -----------------------------
    // 10) Helper: Get Logged-In User Name
    // -----------------------------
    function getLoggedInUserName(callback) {
        let nameElem = document.querySelector("#user-menu-container p.sc-8b7d6573-10");
        if (nameElem && nameElem.textContent.trim() !== "") {
            let name = nameElem.textContent.trim();
            log("Logged in user (from menu): " + name);
            document.body.click();
            callback(name);
        } else {
            let menuButton = document.getElementById("user-menu-button");
            if (menuButton) {
                menuButton.click();
            } else {
                log("Menu button not found.");
            }
            setTimeout(() => {
                nameElem = document.querySelector("#user-menu-container p.sc-8b7d6573-10");
                if (nameElem && nameElem.textContent.trim() !== "") {
                    let name = nameElem.textContent.trim();
                    log("Logged in user (after opening menu): " + name);
                    document.body.click();
                    callback(name);
                } else {
                    log("Still could not find logged in user name.");
                    callback(null);
                }
            }, 1000);
        }
    }

    // -----------------------------
    // 11) Auto-Assign if Sales Rep (Only on Sales Orders)
    // -----------------------------
    function autoAssignIfSalesRep() {
        // Skip if we're on a quote page (convert button exists) or not on a sales order page.
        let convertButton = document.getElementById("convertSalesOrder");
        if (convertButton) {
            log("Quote page detected. Skipping auto-assign logic.");
            return;
        }
        if (!window.location.href.includes('/sales-orders/')) {
            log("Not on a sales order page. Skipping auto-assign logic.");
            return;
        }

        // Check that the "Assigned To" section exists.
        let assignedSection = document.querySelector("div.sc-3e2e0cb2-1.gtsOkx");
        if (!assignedSection) {
            log("Assigned To section not found.");
            return;
        }

        // Check the assigned input—if someone is already assigned, skip automation.
        let assignedInput = assignedSection.querySelector("input");
        if (assignedInput && assignedInput.value.trim() !== "") {
            log("Someone is already assigned. Skipping auto-assign logic.");
            autoAssigned = true;
            return;
        } else {
            autoAssigned = false;
        }

        // Use a specific selector to look for the unique icon in the Assigned To section.
        let iconElem = document.querySelector('div.sc-3e2e0cb2-1.gtsOkx div.sc-130ca08d-2.sc-5768b3d2-2.kauNDR.eRHrWR h4[color="#58698d"]');
        if (!iconElem) {
            log("Icon button not found; likely a user is already assigned. Skipping auto-assign logic.");
            return;
        }
        // If the icon is found but its text is not the default "Ŝ", then someone is assigned.
        if (iconElem.textContent.trim() !== "Ŝ") {
            log("A user is already assigned (icon text is not 'Ŝ'). Skipping auto-assign logic.");
            autoAssigned = true;
            return;
        }

        // Get logged-in user and compare with sales rep.
        getLoggedInUserName(function(loggedInName) {
            if (!loggedInName) {
                log("Logged in user not found.");
                return;
            }
            log("Logged in user: " + loggedInName);

            let salesRepInput = document.querySelector("#salesRep input");
            if (!salesRepInput) {
                log("Sales rep element not found.");
                return;
            }
            let salesRepName = salesRepInput.getAttribute("title").trim();
            log("Sales rep name: " + salesRepName);

            if (loggedInName !== salesRepName) {
                log("Logged in user is not the sales rep. No auto assignment.");
                return;
            }

            // Target the unique icon based on its text "Ŝ" and attribute color="#58698d"
            let iconButton = document.querySelector('h4[color="#58698d"]');
            if (iconButton && iconButton.textContent.trim() === "Ŝ") {
                iconButton.click();
                log("Clicked the icon with Ŝ and color #58698d.");
            } else {
                log("Could not find the icon button by text and color.");
                return;
            }

            // Wait for the assignment menu to appear, then select the logged-in user.
            setTimeout(() => {
                let teamMembers = document.querySelectorAll("div[id^='assign-to-user-listing-']");
                let matchingMember = null;
                teamMembers.forEach(member => {
                    let nameElem = member.querySelector("div.sc-f6227726-5");
                    if (nameElem && nameElem.textContent.trim() === loggedInName) {
                        matchingMember = member;
                    }
                });
                if (matchingMember) {
                    matchingMember.click();
                    log("Auto-assigned to " + loggedInName);
                    autoAssigned = true;

                    // Final step: Click the "Assign" button (id="modalOK")
                    setTimeout(() => {
                        let assignButton = document.getElementById("modalOK");
                        if (assignButton && assignButton.textContent.trim() === "Assign") {
                            assignButton.click();
                            log("Clicked the 'Assign' button at the bottom.");
                            // Wait a bit and then dispatch an Escape key event to close the logged-in user menu.
                            setTimeout(() => {
                                document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
                                log("Dispatched Escape key event to close the logged-in user menu.");
                            }, 1000);
                        } else {
                            log("Could not find the 'Assign' button (modalOK).");
                        }
                    }, 500);
                } else {
                    log("Could not find the logged in user in the assignment list.");
                }
            }, 1000);
        });
    }

    // ====================================
    // PART 2: PRODUCT OPTIONS REMOVER & SEARCH
    // ====================================

    // Store the last created product name
    let lastCreatedProductName = '';

    // Function to hide the "Enable serial tracking" option
    function hideSerialTrackingOption() {
        // Look for the serial tracking option using the specific classes and text content
        const serialTrackingElements = document.querySelectorAll('.SwitchSlideWithLabel-LabelText, .sc-8b7d6573-11, .jamuPn');

        let found = false;
        serialTrackingElements.forEach(element => {
            if (element.textContent.trim() === 'Enable serial tracking') {
                // Find the parent container to hide
                let parentToHide = element.closest('.sc-1d26e07a-1, .foKSlp');
                if (!parentToHide) {
                    // Try getting the parent manually if the closest method doesn't work
                    let parent = element.parentElement;
                    for (let i = 0; i < 4; i++) { // Don't go too deep
                        if (!parent) break;
                        if (parent.classList.contains('sc-1d26e07a-1') || parent.classList.contains('foKSlp')) {
                            parentToHide = parent;
                            break;
                        }
                        parent = parent.parentElement;
                    }
                }

                if (parentToHide) {
                    parentToHide.style.display = 'none';
                    log('Hidden "Enable serial tracking" option');
                    found = true;
                } else {
                    // If we can't find the exact parent container, hide the row containing the text
                    let container = element.parentElement;
                    if (container) {
                        container.style.display = 'none';
                        log('Hidden "Enable serial tracking" option (parent container)');
                        found = true;
                    }
                }
            }
        });

        return found;
    }

    // Function to select the Non-stocked product option using the exact HTML structure
    function selectNonStockedProduct() {
        // Find the non-stocked product text element
        const nonStockedLabels = Array.from(document.querySelectorAll('p.RadioButtonWithLabelContainer-LabelText, p.jamuPn'))
            .filter(el => el.textContent.trim() === 'Non-stocked product');

        if (nonStockedLabels.length === 0) {
            log('Could not find Non-stocked product label');
            return false;
        }

        log(`Found ${nonStockedLabels.length} Non-stocked product labels`);

        // Try to work with the first label we found
        const labelElement = nonStockedLabels[0];

        // Find the label element that contains the radio button
        const labelContainer = labelElement.closest('label.RadioButtonWithLabelContainer, label.lfDODJ');
        if (!labelContainer) {
            log('Could not find label container');
            return false;
        }

        // First, we'll try to directly click the label to activate the radio
        try {
            labelContainer.click();
            log('Clicked the label container');
        } catch (e) {
            log('Error clicking label container: ' + e.message);
        }

        // Find the radio input
        const radioInput = labelContainer.querySelector('input[type="radio"]');
        if (!radioInput) {
            log('Could not find radio input');
            return false;
        }

        // Remove readonly if present
        if (radioInput.hasAttribute('readonly')) {
            radioInput.removeAttribute('readonly');
        }

        // Set checked attribute
        radioInput.checked = true;
        log('Set radio input checked = true');

        // Find the div that contains the radio (should be direct parent or grandparent)
        let radioContainer = radioInput.parentElement;
        while (radioContainer && !radioContainer.classList.contains('sc-8793bcb8-0')) {
            radioContainer = radioContainer.parentElement;
        }

        // Based on your HTML, we need to:
        // 1. Change the container class from eiTaSL to hjmWMW (indicating selection)
        // 2. Make sure it has the inner divs with classes eUBSlS and oLYUG
        if (radioContainer) {
            // Remove any existing classes that might indicate unselected state
            radioContainer.classList.remove('eiTaSL');

            // Add the class that indicates selected state
            radioContainer.classList.add('hjmWMW');
            log('Modified radio container classes for visual selection');

            // Check if we need to add the inner divs
            if (!radioContainer.querySelector('.sc-8793bcb8-1')) {
                // Create the inner structure for the orange circle
                const innerDiv1 = document.createElement('div');
                innerDiv1.className = 'sc-8793bcb8-1 eUBSlS';

                const innerDiv2 = document.createElement('div');
                innerDiv2.className = 'sc-8793bcb8-2 oLYUG';

                innerDiv1.appendChild(innerDiv2);

                // Insert before the radio input
                radioContainer.insertBefore(innerDiv1, radioInput);
                log('Added inner divs for visual selection');
            }

            // Simulate user activity
            radioContainer.dispatchEvent(new MouseEvent('click', {bubbles: true}));
            radioInput.dispatchEvent(new Event('change', {bubbles: true}));

            // Find the outer container div and set its attributes
            const outerContainer = labelElement.closest('div.iiNbRK, div[data-hasfocus]');
            if (outerContainer) {
                // Set data-hasfocus to true
                outerContainer.setAttribute('data-hasfocus', 'true');
                log('Set outer container data-hasfocus to true');
            }

            // As a final measure, check if there's a special form handler in the window object
            try {
                if (window.dispatchEvent) {
                    // Dispatch a custom event that the app might be listening for
                    window.dispatchEvent(new CustomEvent('radioButtonSelected', {
                        detail: { element: radioInput, value: 'Non-stocked product' }
                    }));
                    log('Dispatched custom event to window');
                }
            } catch (e) {
                log('Error dispatching custom event: ' + e.message);
            }

            return true;
        } else {
            log('Could not find the radio container div');
            return false;
        }
    }

    // The main function that checks for and modifies product options
    function checkForProductTypeOptions() {
        // Look for the option containers using the specific class from your HTML
        const options = document.querySelectorAll('div.iiNbRK, div[data-hasfocus]');

        // Always try to hide the serial tracking option
        hideSerialTrackingOption();

        // Only try to modify options if we found some
        let optionsModified = false;

        if (options.length > 0) {
            // Hide unwanted options first
            options.forEach(option => {
                // Find the text element using the class from your HTML
                const textElement = option.querySelector('p.RadioButtonWithLabelContainer-LabelText, p.jamuPn');

                if (!textElement) return;

                const text = textElement.textContent.trim();

                // Hide Stocked product and Service options
                if (text.includes('Stocked product') || text === 'Service') {
                    option.style.display = 'none';
                    log(`Hidden option: "${text}"`);
                    optionsModified = true;
                }
            });

            // After hiding, try to select the Non-stocked product
            const nonStockedSelected = selectNonStockedProduct();
            if (nonStockedSelected) {
                log('Successfully selected Non-stocked product option');
                optionsModified = true;
            }
        }

        if (optionsModified) {
            // If we modified options, watch for the Create button
            watchForCreateButtonClick();
            return true;
        }

        return false;
    }

    // Watch for the Create button click in the product creation popup
    function watchForCreateButtonClick() {
        // Get all buttons in the document
        const buttons = document.querySelectorAll('button');
        let createButton = null;

        // Find the Create button by its text content
        for (const button of buttons) {
            if (button.textContent.trim() === 'Create') {
                createButton = button;
                break;
            }
        }

        if (createButton && !createButton.hasAttribute('data-listener-added')) {
            log('Found Create button in product popup');

            // Mark the button so we don't add multiple listeners
            createButton.setAttribute('data-listener-added', 'true');

            // Add click listener
            createButton.addEventListener('click', function() {
                // Get the product name from the input field
                const nameInput = document.querySelector('#new-product-name');
                if (nameInput) {
                    lastCreatedProductName = nameInput.value.trim();
                    log(`Captured product name: "${lastCreatedProductName}"`);

                    // Store the product name in localStorage for use in the new tab
                    try {
                        localStorage.setItem('lastCreatedProduct', lastCreatedProductName);
                        log('Stored product name in localStorage: ' + lastCreatedProductName);
                    } catch (e) {
                        log('Failed to store product name in localStorage');
                    }

                    // Wait for the creation process to complete, then navigate
                    setTimeout(openProductsInNewTab, 1500);
                } else {
                    log('Could not find product name input');
                }
            });

            log('Added listener to Create button');
        }
    }

    // Function to open the Products page in a new tab
    function openProductsInNewTab() {
        log('Opening Products page in new tab');

        // Simply open the products URL directly
        const productsUrl = 'https://app.inflowinventory.com/products';
        const newTab = window.open(productsUrl, '_blank');

        if (newTab) {
            log('Successfully opened Products in a new tab');
            try {
                newTab.focus();
                log('Focused on new tab');
            } catch (e) {
                log('Could not focus new tab (browser security restriction)');
            }
        } else {
            log('Popup blocked or unable to open new tab');
        }
    }

    // Function to handle search for the newly created product
    function handleProductSearch() {
        // Check if we're on the products page
        if (!window.location.href.includes('/products')) {
            return;
        }

        // Check if we have a stored product name
        const productName = localStorage.getItem('lastCreatedProduct');
        if (!productName) {
            return;
        }

        log(`Found product name in localStorage: "${productName}"`);

        // Wait for the page to fully load, then proceed with search
        setTimeout(() => {
            // ========== ENHANCED BUTTON DETECTION ==========
            let searchButton = null;

            // Method 1: Use character code lookup - looking for 'e' (Lookup_icon contains 'e')
            const elementsWithE = document.querySelectorAll('h4, span, div');
            for (const element of elementsWithE) {
                if (element.textContent.trim() === 'e') {
                    log('Found element with text "e" - potential icon');
                    // Check if this is in a button
                    const buttonParent = element.closest('button');
                    if (buttonParent) {
                        searchButton = buttonParent;
                        log('Found search button via text content "e"');
                        break;
                    }
                }
            }

            // Method 2: Find by Lookup_icon ID
            if (!searchButton) {
                const lookupIcon = document.getElementById('Lookup_icon');
                if (lookupIcon) {
                    // Try to find the parent button
                    const button = lookupIcon.closest('button');
                    if (button) {
                        searchButton = button;
                        log('Found search button via Lookup_icon');
                    }
                }
            }

            // Method 3: Try by specific button class
            if (!searchButton) {
                const buttons = document.querySelectorAll('button.knmauO, button.sc-4a547e52-3');
                if (buttons.length > 0) {
                    searchButton = buttons[0];
                    log('Found search button by class');
                }
            }

            // Method 4: Check if search input is already active
            const searchInput = document.getElementById('searchListing');
            if (searchInput) {
                log('Search input already active, skipping button click');
                useNativeCommands(searchInput, productName);
                return;
            }

            // If we found the search button, click it
            if (searchButton) {
                log('Found search button, clicking it');
                try {
                    searchButton.click();
                    log('Clicked search button');

                    // Wait for the input field to become active
                    setTimeout(() => checkForSearchInput(productName), 300);
                } catch (e) {
                    log('Error clicking search button: ' + e.message);
                }
            } else {
                log('Could not find search button by any method');

                // Last resort: try to find any input that might be the search
                const possibleSearchInputs = document.querySelectorAll('input[placeholder*="search"]');
                if (possibleSearchInputs.length > 0) {
                    log('Found possible search input by placeholder');
                    useNativeCommands(possibleSearchInputs[0], productName);
                } else {
                    log('Could not find any search input');

                    // Try again in 1 second as a fallback
                    setTimeout(() => handleProductSearch(), 1000);
                }
            }
        }, 1000);
    }

    // Function to check for the search input after clicking the button
    function checkForSearchInput(productName) {
        const searchInput = document.getElementById('searchListing');
        if (searchInput) {
            log('Found search input field after button click');
            useNativeCommands(searchInput, productName);
        } else {
            // Try looking for the mini icon then the input
            const miniIcon = document.getElementById('mini_Lookup_icon');
            if (miniIcon) {
                const parentDiv = miniIcon.parentElement;
                if (parentDiv) {
                    const input = parentDiv.querySelector('input');
                    if (input) {
                        log('Found search input via mini icon');
                        useNativeCommands(input, productName);
                        return;
                    }
                }
            }

            // Try to find any input with a search placeholder
            const searchInputs = document.querySelectorAll('input[placeholder*="search" i]');
            if (searchInputs.length > 0) {
                log('Found search input by placeholder');
                useNativeCommands(searchInputs[0], productName);
                return;
            }

            log('Could not find search input after button click, scheduling retry');
            // Retry after a short delay
            setTimeout(() => checkForSearchInput(productName), 300);
        }
    }

    // Function to use browser native commands for input
    function useNativeCommands(inputElement, text) {
        log('Using browser native commands for text input');

        // Ensure the input element is focused first
        inputElement.focus();

        // Clear any existing value using selectAll + delete
        inputElement.select();
        document.execCommand('delete', false);

        // First approach: use execCommand for insertText
        try {
            const success = document.execCommand('insertText', false, text);
            if (success) {
                log('Successfully inserted text using execCommand');
            } else {
                throw new Error('execCommand returned false');
            }
        } catch (e) {
            log('execCommand failed: ' + e.message);

            // Fallback: try to set the value directly and trigger events
            try {
                inputElement.value = text;
                log('Set input value directly');

                // Create a new InputEvent that more closely resembles a real input event
                try {
                    const inputEvent = new InputEvent('input', {
                        bubbles: true,
                        cancelable: true,
                        inputType: 'insertText',
                        data: text
                    });
                    inputElement.dispatchEvent(inputEvent);
                    log('Dispatched InputEvent with inputType and data');
                } catch (e2) {
                    // Fallback for older browsers or environments that don't support InputEvent constructor
                    log('InputEvent constructor failed, using Event instead');
                    inputElement.dispatchEvent(new Event('input', { bubbles: true, cancelable: true }));
                }

                // Also dispatch a change event
                inputElement.dispatchEvent(new Event('change', { bubbles: true, cancelable: true }));
            } catch (e2) {
                log('All input methods failed: ' + e2.message);
            }
        }

        // Add a slight delay, then press Enter
        setTimeout(() => {
            try {
                // Try to use execCommand to execute an Enter keypress
                document.execCommand('insertParagraph', false);
                log('Executed insertParagraph command (Enter key)');
            } catch (e) {
                log('insertParagraph command failed: ' + e.message);

                // Fallback: simulate Enter key event
                const enterEvent = new KeyboardEvent('keydown', {
                    bubbles: true,
                    cancelable: true,
                    key: 'Enter',
                    code: 'Enter',
                    keyCode: 13,
                    which: 13
                });
                inputElement.dispatchEvent(enterEvent);
                log('Dispatched Enter key event');
            }

            // Wait a bit longer, then look for the product in search results
            setTimeout(() => {
                // Try to directly call any search or filter functions on the window
                try {
                    // Look for common search function names in the global scope
                    const possibleFunctions = [
                        'search', 'doSearch', 'performSearch', 'filterResults',
                        'filterProducts', 'searchProducts', 'findProducts'
                    ];

                    for (const funcName of possibleFunctions) {
                        if (typeof window[funcName] === 'function') {
                            log(`Found global search function: ${funcName}`);
                            window[funcName](text);
                            break;
                        }
                    }
                } catch (e) {
                    log('Failed to call global search function: ' + e.message);
                }

                // Remove the item from localStorage and look for product
                localStorage.removeItem('lastCreatedProduct');
                findAndClickProduct(text);
            }, 800);
        }, 300);
    }

    // Function to find and click on the product in the search results
    function findAndClickProduct(productName) {
        log(`Looking for product "${productName}" in the results`);

        // Try multiple strategies to find the product
        let foundElement = null;

        // Method 1: Look for elements with name attribute matching our product
        const namedElements = document.querySelectorAll(`[name="${productName}"]`);
        if (namedElements.length > 0) {
            log('Found product by name attribute');
            foundElement = namedElements[0];

            // Try to find a clickable parent element
            let parent = foundElement;
            for (let i = 0; i < 5; i++) {  // Don't go too deep
                if (!parent) break;

                if (parent.tagName === 'LI' || parent.tagName === 'DIV' && parent.getAttribute('id')?.includes('listing-line')) {
                    foundElement = parent;
                    log('Found clickable parent element');
                    break;
                }
                parent = parent.parentElement;
            }
        }

        // Method 2: Look for any row or item containing the product name
        if (!foundElement) {
            // First try list items
            const listItems = document.querySelectorAll('li, tr, div[role="row"]');
            for (const item of listItems) {
                if (item.textContent.includes(productName)) {
                    foundElement = item;
                    log('Found product in list/table by text content');
                    break;
                }
            }

            // If still not found, look for any element with the exact product name
            if (!foundElement) {
                const allElements = document.querySelectorAll('div, span, td, h4');
                for (const el of allElements) {
                    if (el.textContent.trim() === productName) {
                        log('Found product by exact text match');

                        // Look for a clickable parent
                        let parent = el;
                        for (let i = 0; i < 5; i++) {  // Don't go too deep
                            if (!parent) break;
                            if (parent.tagName === 'LI' || parent.tagName === 'TR' ||
                                (parent.tagName === 'DIV' && (
                                    parent.getAttribute('role') === 'row' ||
                                    parent.getAttribute('id')?.includes('listing-line')
                                ))) {
                                foundElement = parent;
                                log('Found clickable parent element');
                                break;
                            }
                            parent = parent.parentElement;
                        }

                        // If no good parent was found, use the element itself
                        if (!foundElement) {
                            foundElement = el;
                        }
                        break;
                    }
                }
            }
        }

        if (foundElement) {
            log('Found product element, clicking it');
            try {
                foundElement.click();
                log('Successfully clicked on product: ' + productName);
            } catch (e) {
                log('Error clicking product: ' + e.message);
            }
        } else {
            log('Could not find the product in the results, will retry in 1 second');
            // Try again after a short delay in case results are still loading
            setTimeout(() => findAndClickProduct(productName), 1000);
        }
    }

    // ====================================
    // COMBINED INITIALIZATION
    // ====================================

    // Master initialization function that calls both parts
    function masterInit() {
        // Part 1: PO Enhancements & Auto-Assign
        tryInsertPO();
        renamePOElements();
        removeFulfillButton();
        checkCreditHold();
        handleQuotePage();
        addCardOnFileIndicator();
        autoAssignIfSalesRep();

        // Part 2: Product Options & Search
        const url = window.location.href;

        if (url.includes('/products')) {
            // We're on the products page
            const marker = document.getElementById('product-search-initialized');
            if (!marker) {
                log('On products page, setting up product search');

                // Add a marker to avoid duplicate initialization
                const marker = document.createElement('div');
                marker.id = 'product-search-initialized';
                marker.style.display = 'none';
                document.body.appendChild(marker);

                // Handle product search
                setTimeout(handleProductSearch, 800);
            }
        } else {
            // Check for product creation dialogs
            checkForProductTypeOptions();
        }
    }

    // Combined observer setup for both parts of the script
    let observer = new MutationObserver(() => {
        masterInit();
    });
    observer.observe(document.body, { childList: true, subtree: true });

    // History change handlers (for SPA navigation)
    const originalPushState = history.pushState;
    const originalReplaceState = history.replaceState;

    function handleHistoryChange() {
        masterInit();
    }

    history.pushState = function() {
        let result = originalPushState.apply(this, arguments);
        handleHistoryChange();
        return result;
    };

    history.replaceState = function() {
        let result = originalReplaceState.apply(this, arguments);
        handleHistoryChange();
        return result;
    };

    window.addEventListener('popstate', handleHistoryChange);

    // Start initialization process with retry logic
    function startScript() {
        // Start the script with a delay to ensure page is loaded
        setTimeout(masterInit, 800);

        // Add additional checks after longer delays to catch late-loading elements
        setTimeout(masterInit, 1500);
        setTimeout(masterInit, 3000);
    }

    startScript();

    // Also run initialization periodically to catch any missed changes
    setInterval(masterInit, 3000);

    // Watch for page navigation within SPA
    let lastUrl = window.location.href;
    const urlObserver = new MutationObserver(() => {
        if (lastUrl !== window.location.href) {
            lastUrl = window.location.href;
            startScript(); // Restart the initialization process
        }
    });

    urlObserver.observe(document.body, { childList: true, subtree: true });

    log('InFlow Master Script loaded - All enhancements active');
})();
