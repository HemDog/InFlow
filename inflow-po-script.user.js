// ==UserScript==
// @name         InFlow Auto PO Filler & Rename PO Continuous Attempts
// @namespace    http://yourdomain.com
// @version      1.10
// @description  Automatically fills the PO field and renames PO continuously as long as you're on a sales-orders page
// @match        https://app.inflowinventory.com/sales-orders/*
// @grant        GM_xmlhttpRequest
// @connect      docs.google.com
// @connect      googleusercontent.com
// @downloadURL  https://raw.githubusercontent.com/HemDog/inflow-po-script.user.js/main/inflow-po-script.user.js
// @updateURL    https://raw.githubusercontent.com/HemDog/inflow-po-script.user.js/main/inflow-po-script.user.js
// ==/UserScript==

(function() {
    'use strict';

    const sheetUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ-iA8XppRLS2s-30ElJSdvNbrgOQB-kaYIS39uC1TwvHYS4HgmZKMbSR1vUuekeiYHlQaYsFuJk3P9/pub?output=csv";

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
                let row = data.find(r => r.Customer && r.Customer.trim().toLowerCase() === customerName.toLowerCase());
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

    function renamePOElements() {
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
        }, 1000);
    });
    observer.observe(document.body, { childList: true, subtree: true });

    // Initial attempts
    tryInsertPO();
    renamePOElements();

    // History API override to detect route changes
    const originalPushState = history.pushState;
    const originalReplaceState = history.replaceState;

    function handleHistoryChange() {
        // Wait a moment for the new page content to load
        setTimeout(() => {
            tryInsertPO();
            renamePOElements();
        }, 1000);
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

    // Continuous fallback polling: every 3 seconds, try again to ensure PO is always inserted
    setInterval(() => {
        tryInsertPO();
        renamePOElements();
    }, 3000);
})();
