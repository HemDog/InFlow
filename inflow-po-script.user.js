// ==UserScript==
// @name         InFlow Combined (PO Enhancements + Auto-Assign)
// @namespace    http://yourdomain.com
// @version      2.3.6
// @description  Inserts a "Card on file" row, prepends Blanket PO info, renames elements, removes the Fulfill button, checks credit hold, handles quote pages, AND auto-assigns if the logged-in user matches the sales rep (only on Sales Orders). If someone is already assigned, auto-assign is skipped. After auto-assignment, the logged-in user menu is closed via an Escape key event.
// @match        https://app.inflowinventory.com/*
// @grant        GM_xmlhttpRequest
// @connect      docs.google.com
// @connect      googleusercontent.com
// @downloadURL  https://raw.githubusercontent.com/HemDog/InFlow/main/inflow-po-script.user.js
// @updateURL    https://raw.githubusercontent.com/HemDog/InFlow/main/inflow-po-script.user.js
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

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
            console.log("Convert to order button clicked. Disabling quote page behavior and reloading page shortly.");
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
                    console.log("No matching customer found in the Blanket PO sheet for:", customerName);
                    return;
                }
                let blanketPO = row["Blanket PO"] || "";
                if (blanketPO) {
                    let formattedPO = `(BLANKET PO#:${blanketPO})`;
                    if (poField.value.includes("BLANKET PO#")) {
                        console.log(`"BLANKET PO#" already found in remarks. Not modifying text.`);
                    } else {
                        poField.value = formattedPO + " " + poField.value;
                        console.log(`Prepended PO for ${customerName} with ${formattedPO}`);
                    }
                } else {
                    console.log(`No Blanket PO found for ${customerName}`);
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
            console.log("A card on file request is already in flight for this customer. Skipping...");
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
                console.log('Renamed "PO #" to "Look Up Number"');
            }
            if (p.textContent.trim() === "Include shipping") {
                p.textContent = "Fulfill/Pick";
                console.log('Renamed "Include shipping" to "Fulfill/Pick"');
            }
            if (p.textContent.trim() === "Assigned To") {
                p.textContent = "Converted/Made by";
                console.log('Renamed "Assigned To" to "Converted/Made by"');
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
            console.log("Fulfill button removed.");
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
            console.log("Print button removed because the customer is on credit hold.");
        }
        let emailButton = document.getElementById("so-email-dropdown");
        if (emailButton) {
            emailButton.remove();
            console.log("Email button removed because the customer is on credit hold.");
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
                    console.log("Renamed Sales rep to Quote by.");
                }
            });
            let container = document.querySelector("div.sc-3e2e0cb2-0.gLiEEA");
            if (container) {
                let sections = container.querySelectorAll("div.sc-3e2e0cb2-1.gtsOkx");
                if (sections.length >= 2) {
                    sections[1].remove();
                    console.log("Removed the second (assigned-to) section for quotes.");
                }
            }
        }
    }

    // -----------------------------
    // 10) Helper: Get Logged-In User Name (Duplicate for auto-assign)
    // -----------------------------
    function getLoggedInUserName(callback) {
        let nameElem = document.querySelector("#user-menu-container p.sc-8b7d6573-10");
        if (nameElem && nameElem.textContent.trim() !== "") {
            let name = nameElem.textContent.trim();
            console.log("Logged in user (from menu):", name);
            document.body.click();
            callback(name);
        } else {
            let menuButton = document.getElementById("user-menu-button");
            if (menuButton) {
                menuButton.click();
            } else {
                console.log("Menu button not found.");
            }
            setTimeout(() => {
                nameElem = document.querySelector("#user-menu-container p.sc-8b7d6573-10");
                if (nameElem && nameElem.textContent.trim() !== "") {
                    let name = nameElem.textContent.trim();
                    console.log("Logged in user (after opening menu):", name);
                    document.body.click();
                    callback(name);
                } else {
                    console.log("Still could not find logged in user name.");
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
            console.log("Quote page detected. Skipping auto-assign logic.");
            return;
        }
        if (!window.location.href.includes('/sales-orders/')) {
            console.log("Not on a sales order page. Skipping auto-assign logic.");
            return;
        }

        // Check that the "Assigned To" section exists.
        let assignedSection = document.querySelector("div.sc-3e2e0cb2-1.gtsOkx");
        if (!assignedSection) {
            console.log("Assigned To section not found.");
            return;
        }

        // Check the assigned input—if someone is already assigned, skip automation.
        let assignedInput = assignedSection.querySelector("input");
        if (assignedInput && assignedInput.value.trim() !== "") {
            console.log("Someone is already assigned. Skipping auto-assign logic.");
            autoAssigned = true;
            return;
        } else {
            autoAssigned = false;
        }

        // Use a specific selector to look for the unique icon in the Assigned To section.
        let iconElem = document.querySelector('div.sc-3e2e0cb2-1.gtsOkx div.sc-130ca08d-2.sc-5768b3d2-2.kauNDR.eRHrWR h4[color="#58698d"]');
        if (!iconElem) {
            console.log("Icon button not found; likely a user is already assigned. Skipping auto-assign logic.");
            return;
        }
        // If the icon is found but its text is not the default "Ŝ", then someone is assigned.
        if (iconElem.textContent.trim() !== "Ŝ") {
            console.log("A user is already assigned (icon text is not 'Ŝ'). Skipping auto-assign logic.");
            autoAssigned = true;
            return;
        }

        // Get logged-in user and compare with sales rep.
        getLoggedInUserName(function(loggedInName) {
            if (!loggedInName) {
                console.log("Logged in user not found.");
                return;
            }
            console.log("Logged in user:", loggedInName);

            let salesRepInput = document.querySelector("#salesRep input");
            if (!salesRepInput) {
                console.log("Sales rep element not found.");
                return;
            }
            let salesRepName = salesRepInput.getAttribute("title").trim();
            console.log("Sales rep name:", salesRepName);

            if (loggedInName !== salesRepName) {
                console.log("Logged in user is not the sales rep. No auto assignment.");
                return;
            }

            // Target the unique icon based on its text "Ŝ" and attribute color="#58698d"
            let iconButton = document.querySelector('h4[color="#58698d"]');
            if (iconButton && iconButton.textContent.trim() === "Ŝ") {
                iconButton.click();
                console.log("Clicked the icon with Ŝ and color #58698d.");
            } else {
                console.log("Could not find the icon button by text and color.");
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
                    console.log("Auto-assigned to", loggedInName);
                    autoAssigned = true;

                    // Final step: Click the "Assign" button (id="modalOK")
                    setTimeout(() => {
                        let assignButton = document.getElementById("modalOK");
                        if (assignButton && assignButton.textContent.trim() === "Assign") {
                            assignButton.click();
                            console.log("Clicked the 'Assign' button at the bottom.");
                            // Wait a bit and then dispatch an Escape key event to close the logged-in user menu.
                            setTimeout(() => {
                                document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
                                console.log("Dispatched Escape key event to close the logged-in user menu.");
                            }, 1000);
                        } else {
                            console.log("Could not find the 'Assign' button (modalOK).");
                        }
                    }, 500);
                } else {
                    console.log("Could not find the logged in user in the assignment list.");
                }
            }, 1000);
        });
    }

    // -----------------------------
    // 12) Main Init
    // -----------------------------
    function init() {
        tryInsertPO();
        renamePOElements();
        removeFulfillButton();
        checkCreditHold();
        handleQuotePage();
        addCardOnFileIndicator();

        // Run auto-assign logic last
        autoAssignIfSalesRep();
    }

    // -----------------------------
    // 13) Observe DOM + Route Changes
    // -----------------------------
    let observer = new MutationObserver(() => { init(); });
    observer.observe(document.body, { childList: true, subtree: true });
    init();

    const originalPushState = history.pushState;
    const originalReplaceState = history.replaceState;
    function handleHistoryChange() { init(); }
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

    setInterval(() => { init(); }, 3000);

})();
