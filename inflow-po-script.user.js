// ==UserScript==
// @name         InFlow Auto PO Filler & Rename PO Continuous Attempts
// @namespace    http://yourdomain.com
// @version      1.11
// @description  Automatically fills the PO field and renames PO continuously as long as you're on a sales-orders page
// @match        https://app.inflowinventory.com/*
// @grant        GM_xmlhttpRequest
// @connect      docs.google.com
// @connect      googleusercontent.com
// @downloadURL  https://raw.githubusercontent.com/HemDog/inflow-po-script.user.js/main/inflow-po-script.user.js
// @updateURL    https://raw.githubusercontent.com/HemDog/inflow-po-script.user.js/main/inflow-po-script.user.js
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

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

    // Rename PO elements if we're on the Sales Orders page
    function renamePOElements() {
        // Only run if the URL includes '/sales-orders/'
        if (!window.location.href.includes('/sales-orders/')) {
            return;
        }

        let paragraphs = document.querySelectorAll('p');
        paragraphs.forEach(p => {
            if (p.textContent.trim() === "PO #") {
                p.textContent = "Customer Notes";
                console.log('Renamed "PO #" to "Customer Notes"');
            }
        });
    }

    // Combine the two main calls into one function
    function init() {
        tryInsertPO();
        renamePOElements();
    }

    // MutationObserver to detect DOM changes
    let observer = new MutationObserver(() => {
        // Call immediately (no 1-second delay)
        init();
    });
    observer.observe(document.body, { childList: true, subtree: true });

    // Initial attempt when the script first runs
    init();

    // Override History API to detect route changes (single-page app transitions)
    const originalPushState = history.pushState;
    const originalReplaceState = history.replaceState;

    function handleHistoryChange() {
        // Call immediately, or with a small delay if needed
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

    // Fallback polling every 3 seconds, in case the above methods miss an update
    setInterval(() => {
        init();
    }, 3000);

})();
