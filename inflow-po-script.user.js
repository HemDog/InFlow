// ==UserScript==
// @updateURL    https://raw.githubusercontent.com/HemDog/InFlow/main/inflow-po-script.user.js
// @downloadURL  https://raw.githubusercontent.com/HemDog/Inflow/main/inflow-po-script.user.js
// @name         InFlow Auto PO Filler & Rename PO with History API Detection
// @namespace    http://yourdomain.com
// @version      1.7
// @description  Automatically fills the PO field and renames PO-related text even with SPA navigation
// @match        https://app.inflowinventory.com/sales-orders/*
// @grant        GM_xmlhttpRequest
// @connect      docs.google.com
// @connect      googleusercontent.com
// ==/UserScript==

(function() {
    'use strict';

    const sheetUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ-iA8XppRLS2s-30ElJSdvNbrgOQB-kaYIS39uC1TwvHYS4HgmZKMbSR1vUuekeiYHlQaYsFuJk3P9/pub?output=csv";
    let lastProcessedCustomer = null;

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

    function tryInsertPO() {
        let customerField = document.querySelector("#so_customer input[type='text']");
        let poField = document.querySelector('textarea[name="remarks"]');

        if (!customerField || !poField) {
            return;
        }

        let customerName = customerField.value.trim();

        if (lastProcessedCustomer === customerName) return;

        GM_xmlhttpRequest({
            method: "GET",
            url: sheetUrl,
            onload: function(response) {
                if (response.status !== 200) {
                    console.error("Failed to fetch CSV. Status:", response.status, response.statusText);
                    return;
                }
                let data = parseCSV(response.responseText);
                let row = data.find(r => r.Customer.trim().toLowerCase() === customerName.toLowerCase());
                if (!row) {
                    console.log("No matching customer found in the sheet for:", customerName);
                    lastProcessedCustomer = customerName;
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

                lastProcessedCustomer = customerName;
            },
            onerror: function(err) {
                console.error("Error fetching CSV:", err);
            }
        });
    }

    function renamePOElements() {
        // Rename "PO #" in <p> elements to "Customer Notes"
        let paragraphs = document.querySelectorAll('p');
        paragraphs.forEach(p => {
            if (p.textContent.trim() === "PO #") {
                p.textContent = "Customer Notes";
                console.log('Renamed "PO #" to "Customer Notes"');
            }
        });
    }

    // MutationObserver to detect DOM changes
    let observer = new MutationObserver(() => {
        setTimeout(() => {
            tryInsertPO();
            renamePOElements();
        }, 500);
    });
    observer.observe(document.body, { childList: true, subtree: true });

    // Initial attempts
    tryInsertPO();
    renamePOElements();

    // History API override to detect route changes
    const originalPushState = history.pushState;
    const originalReplaceState = history.replaceState;

    function handleHistoryChange() {
        // Wait a bit for DOM changes to occur before running
        setTimeout(() => {
            tryInsertPO();
            renamePOElements();
        }, 500);
    }

    history.pushState = function(state) {
        let result = originalPushState.apply(this, arguments);
        handleHistoryChange();
        return result;
    };

    history.replaceState = function(state) {
        let result = originalReplaceState.apply(this, arguments);
        handleHistoryChange();
        return result;
    };

    window.addEventListener('popstate', () => {
        handleHistoryChange();
    });
})();
