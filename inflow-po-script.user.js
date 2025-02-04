// ==UserScript==
// @name         InFlow Auto PO Filler & Rename PO Continuous Attempts
// @namespace    http://yourdomain.com
// @version      1.21
// @description  Automatically fills the PO field, renames elements, and makes quote- and sales order–specific modifications.
// @match        https://app.inflowinventory.com/*
// @grant        GM_xmlhttpRequest
// @connect      docs.google.com
// @connect      googleusercontent.com
// @downloadURL  https://github.com/HemDog/InFlow/raw/refs/heads/main/inflow-po-script.user.js
// @updateURL    https://github.com/HemDog/InFlow/raw/refs/heads/main/inflow-po-script.user.js
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    // Flag to track if the credit hold popup has been dismissed
    let creditHoldPopupDismissed = false;

    // URL to your published Google Sheet CSV
    const sheetUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ-iA8XppRLS2s-30ElJSdvNbrgOQB-kaYIS39uC1TwvHYS4HgmZKMbSR1vUuekeiYHlQaYsFuJk3P9/pub?output=csv";

    // Parse CSV data into an array of objects
    function parseCSV(csvText) {
        let lines = csvText.trim().split('\n');
        let headers = lines[0].split(',');
        let result = [];
        for (let i = 1; i < lines.length; i++) {
            let obj = {};
            let row = lines[i].split(',');
            headers.forEach((h, index) => {
                obj[h.trim()] = (row[index] || "").trim();
            });
            result.push(obj);
        }
        return result;
    }

    // Insert PO if we're on the Sales Orders page
    function tryInsertPO() {
        // Only run if the URL includes '/sales-orders/'
        if (!window.location.href.includes('/sales-orders/')) {
            return;
        }

        let customerField = document.querySelector("#so_customer input[type='text']");
        let poField = document.querySelector('textarea[name="remarks"]');

        if (!customerField || !poField) {
            return;
        }

        let customerName = customerField.value.trim();
        if (!customerName) return; // no customer name, nothing to do

        GM_xmlhttpRequest({
            method: "GET",
            url: sheetUrl,
            onload: function(response) {
                if (response.status !== 200) {
                    console.error("Failed to fetch CSV. Status:", response.status, response.statusText);
                    return;
                }
                let data = parseCSV(response.responseText);
                let row = data.find(r =>
                    r.Customer && r.Customer.trim().toLowerCase() === customerName.toLowerCase()
                );
                if (!row) {
                    console.log("No matching customer found in the sheet for:", customerName);
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
                console.error("Error fetching CSV:", err);
            }
        });
    }

    // Rename elements on Sales Orders pages
    function renamePOElements() {
        // Only run if the URL includes '/sales-orders/'
        if (!window.location.href.includes('/sales-orders/')) {
            return;
        }

        let paragraphs = document.querySelectorAll('p');
        paragraphs.forEach(p => {
            // Rename "PO #" to "Look Up Number"
            if (p.textContent.trim() === "PO #") {
                p.textContent = "Look Up Number";
                console.log('Renamed "PO #" to "Look Up Number"');
            }

            // Rename "Include shipping" to "Fulfill/Pick"
            if (p.textContent.trim() === "Include shipping") {
                p.textContent = "Fulfill/Pick";
                console.log('Renamed "Include shipping" to "Fulfill/Pick"');
            }

            // Rename "Assigned To" to "Converted/Made by"
            if (p.textContent.trim() === "Assigned To") {
                p.textContent = "Converted/Made by";
                console.log('Renamed "Assigned To" to "Converted/Made by"');
            }
        });
    }

    // Remove the Fulfill button regardless of customer status
    function removeFulfillButton() {
        let fulfillButton = document.getElementById("so-btn-Fulfill");
        if (fulfillButton) {
            fulfillButton.remove();
            console.log("Fulfill button removed.");
        }
    }

    // Check for Credit Hold status and remove additional elements if needed
    function checkCreditHold() {
        // Look for an input element whose title equals "CREDIT HOLD"
        let creditHoldElement = document.querySelector('input[title="CREDIT HOLD"]');

        // If the credit hold element is not found, reset the dismissed flag.
        if (!creditHoldElement) {
            creditHoldPopupDismissed = false;
            return;
        }

        // Remove the Print button if it exists
        let printButton = document.getElementById("sales-order-print-dropdown");
        if (printButton) {
            printButton.remove();
            console.log("Print button removed because the customer is on credit hold.");
        }

        // Remove the Email button if it exists
        let emailButton = document.getElementById("so-email-dropdown");
        if (emailButton) {
            emailButton.remove();
            console.log("Email button removed because the customer is on credit hold.");
        }

        // Show the credit hold popup if it hasn't been dismissed
        if (creditHoldPopupDismissed) return;
        if (!document.getElementById("credit-hold-popup")) {
            // Create an overlay for the popup
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

            // Create the popup container
            let popup = document.createElement('div');
            popup.style.backgroundColor = "white";
            popup.style.padding = "20px";
            popup.style.borderRadius = "8px";
            popup.style.textAlign = "center";
            popup.style.boxShadow = "0 0 10px rgba(0,0,0,0.25)";
            popup.style.maxWidth = "90%";

            // Title in larger, bold letters
            let title = document.createElement('h1');
            title.textContent = "This customer is on credit hold!";
            title.style.fontSize = "24px";
            title.style.fontWeight = "bold";
            title.style.margin = "0 0 10px 0";

            // Message text
            let message = document.createElement('p');
            message.textContent = "Overdue balance must be paid in full before order can be fulfilled.";
            message.style.margin = "0";

            // Close button for the popup
            let closeButton = document.createElement('button');
            closeButton.textContent = "Close";
            closeButton.style.marginTop = "15px";
            closeButton.addEventListener('click', function() {
                creditHoldPopupDismissed = true;
                overlay.remove();
            });

            // Assemble the popup
            popup.appendChild(title);
            popup.appendChild(message);
            popup.appendChild(closeButton);
            overlay.appendChild(popup);
            document.body.appendChild(overlay);
        }
    }

    // Handle Quote Page Modifications
    // A page is considered a quote page if the "Convert to order" button is present.
    function handleQuotePage() {
        let convertButton = document.getElementById("convertSalesOrder");
        if (convertButton) {
            // Rename any Sales rep element with text "Sales rep" to "Quote by"
            let salesRepElements = document.querySelectorAll('p.sc-a3cffb35-4.botIdl');
            salesRepElements.forEach(elem => {
                if (elem.textContent.trim() === "Sales rep") {
                    elem.textContent = "Quote by";
                    console.log("Renamed Sales rep to Quote by.");
                }
            });
            // Remove the assigned-to section
            // The assigned-to section is the second .gtsOkx element inside the parent container.
            let container = document.querySelector("div.sc-3e2e0cb2-0.gLiEEA");
            if (container) {
                let sections = container.querySelectorAll("div.sc-3e2e0cb2-1.gtsOkx");
                if (sections.length >= 2) {
                    sections[1].remove();
                    console.log("Removed the second (assigned-to) section.");
                }
            }
        }
    }

    // Combine all main calls into one function
    function init() {
        tryInsertPO();
        renamePOElements();
        removeFulfillButton(); // Always remove the fulfill button
        checkCreditHold();     // Remove Print/Email and show popup if on credit hold
        handleQuotePage();     // Apply quote page modifications if the Convert to order button exists
    }

    // Use a MutationObserver to detect DOM changes
    let observer = new MutationObserver(() => {
        init();
    });
    observer.observe(document.body, { childList: true, subtree: true });

    // Initial call when the script first runs
    init();

    // Override History API to detect route changes (for single-page app transitions)
    const originalPushState = history.pushState;
    const originalReplaceState = history.replaceState;

    function handleHistoryChange() {
        init();
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

    // Fallback polling every 3 seconds in case the above methods miss an update
    setInterval(() => {
        init();
    }, 3000);

    //test
})();
