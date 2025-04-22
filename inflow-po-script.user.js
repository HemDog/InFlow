// ==UserScript==
// @name         InFlow Combined Script (Converter + Enhancements)
// @namespace    http://yourdomain.com
// @version      3.1.2
// @description  Combines Quote-to-Order Converter with Google Sheets and PO/Sales Order/Product Enhancements. Allows multi-line/indenting in Internal Note.
// @match        https://app.inflowinventory.com/*
// @grant        GM_xmlhttpRequest
// @connect      script.google.com
// @connect      docs.google.com
// @connect      googleusercontent.com
// @downloadURL  https://raw.githubusercontent.com/HemDog/InFlow/main/inflow-combined-script.user.js // Replace if hosted elsewhere
// @updateURL    https://raw.githubusercontent.com/HemDog/InFlow/main/inflow-combined-script.user.js // Replace if hosted elsewhere
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';
    // ** ADD BASIC LOG **
    console.log('--- InFlow Combined Script START ---');

    // ====================================
    // Shared Helper Functions & Variables
    // ====================================

    // Simple Log Helper (from Master Script)
    function log(message, data = null, level = 'info') { // Added data/level params like debugLog
        const prefix = '[InFlow Combined]';
        if (level === 'error') {
            console.error(`${prefix} ERROR: ${message}`, data || '');
        } else if (level === 'warn') {
            console.warn(`${prefix} WARNING: ${message}`, data || '');
        } else {
            console.log(`${prefix} ${message}`, data || '');
        }
    }


    // Helper function for debouncing
    function debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func.apply(this, args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    }

    // Enhanced Debugging function (from Quote Converter)
    function debugLog(message, data = null, level = 'info') {
        const prefix = '[InFlow Debug]';
        if (level === 'error') {
            console.error(`${prefix} ERROR: ${message}`, data || '');
            // Show in UI too
            showError(`Debug Error: ${message} ${data ? JSON.stringify(data) : ''}`);
        } else if (level === 'warn') {
            console.warn(`${prefix} WARNING: ${message}`, data || '');
        } else {
            console.log(`${prefix} ${message}`, data || '');
        }
    }

    // Error notification function (from Quote Converter)
    function showError(message) {
        const existingError = document.getElementById('quote-converter-error');
        if (existingError) {
            existingError.innerHTML = message;
            return;
        }

        const errorDiv = document.createElement('div');
        errorDiv.id = 'quote-converter-error';
        errorDiv.style.position = 'fixed';
        errorDiv.style.top = '10px';
        errorDiv.style.right = '10px';
        errorDiv.style.backgroundColor = '#f44336';
        errorDiv.style.color = 'white';
        errorDiv.style.padding = '10px';
        errorDiv.style.borderRadius = '4px';
        errorDiv.style.zIndex = '9999';
        errorDiv.innerHTML = message;

        const closeBtn = document.createElement('button');
        closeBtn.textContent = '×';
        closeBtn.style.marginLeft = '10px';
        closeBtn.style.background = 'none';
        closeBtn.style.border = 'none';
        closeBtn.style.color = 'white';
        closeBtn.style.cursor = 'pointer';
        closeBtn.onclick = function() { errorDiv.remove(); };
        errorDiv.appendChild(closeBtn);

        document.body.appendChild(errorDiv);
    }

    // Status notification function (from Quote Converter)
    function showStatusNotification(message, type = 'loading') {
        // Remove any existing notification
        const existingNotification = document.getElementById('google-sheets-status');
        if (existingNotification) {
            existingNotification.remove();
        }

        // Create notification element
        const notification = document.createElement('div');
        notification.id = 'google-sheets-status';
        notification.style.position = 'fixed';
        notification.style.bottom = '10px';
        notification.style.right = '10px';
        notification.style.padding = '10px 15px';
        notification.style.borderRadius = '4px';
        notification.style.zIndex = '9999';
        notification.style.display = 'flex';
        notification.style.alignItems = 'center';
        notification.style.boxShadow = '0 2px 5px rgba(0,0,0,0.2)';
        notification.style.fontSize = '14px';

        // Set style based on type
        if (type === 'loading') {
            notification.style.backgroundColor = '#2196F3';
            notification.style.color = 'white';
        } else if (type === 'success') {
            notification.style.backgroundColor = '#4CAF50';
            notification.style.color = 'white';
        } else if (type === 'error') {
            notification.style.backgroundColor = '#F44336';
            notification.style.color = 'white';
        }

        // Create loader if loading
        if (type === 'loading') {
            const loader = document.createElement('div');
            loader.style.width = '20px';
            loader.style.height = '20px';
            loader.style.border = '3px solid rgba(255,255,255,0.3)';
            loader.style.borderTop = '3px solid white';
            loader.style.borderRadius = '50%';
            loader.style.marginRight = '10px';
            loader.style.animation = 'spin 1s linear infinite';

            // Add keyframes for spinner animation
            const style = document.createElement('style');
            style.textContent = `
                @keyframes spin {
                    0% { transform: rotate(0deg); }
                    100% { transform: rotate(360deg); }
                }
            `;
            document.head.appendChild(style);

            notification.appendChild(loader);
        } else {
            // Add icon for success/error
            const icon = document.createElement('span');
            icon.style.marginRight = '10px';
            icon.style.fontSize = '20px';

            if (type === 'success') {
                icon.innerHTML = '✓';
            } else if (type === 'error') {
                icon.innerHTML = '✗';
            }

            notification.appendChild(icon);
        }

        // Add message
        const messageSpan = document.createElement('span');
        messageSpan.textContent = message;
        notification.appendChild(messageSpan);

        // Add close button
        const closeBtn = document.createElement('span');
        closeBtn.textContent = '×';
        closeBtn.style.marginLeft = '15px';
        closeBtn.style.cursor = 'pointer';
        closeBtn.style.fontSize = '20px';
        closeBtn.onclick = function() { notification.remove(); };
        notification.appendChild(closeBtn);

        // Add to document
        document.body.appendChild(notification);

        // Auto-remove successful notifications after 5 seconds
        if (type === 'success') {
            setTimeout(() => {
                if (document.getElementById('google-sheets-status')) {
                    notification.remove();
                }
            }, 5000);
        }

        return notification;
    }

    // CSV Parsing Helper (from Master Script)
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

    // Custom parser for MM/YY (from Master Script)
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

    // Get User Initials (from Quote Converter - Placeholder)
    function cacheUserName() {
        getLoggedInUserName(function(name) {
            if (name) {
                localStorage.setItem('cachedUserName', name);
                debugLog(`Cached user name: ${name}`);
            }
        });
    }

    function getUserInitials() {
        // Try to get from cache first
        const cachedName = localStorage.getItem('cachedUserName');
        if (cachedName) {
            const parts = cachedName.split(' ').filter(part => part.length > 0);
            if (parts.length > 0) {
                const firstInitial = parts[0][0];
                const lastInitial = parts.length > 1 ? parts[parts.length - 1][0] : '';
                const initials = (firstInitial + lastInitial).toUpperCase();
                debugLog(`Using cached name "${cachedName}" for initials "${initials}"`);
                return initials;
            }
        }

        // If no cache, try direct DOM access
        const nameElem = document.querySelector("#user-menu-container p.sc-8b7d6573-10");
        if (nameElem && nameElem.textContent?.trim()) {
            const fullName = nameElem.textContent.trim();
            const parts = fullName.split(' ').filter(part => part.length > 0);
            if (parts.length > 0) {
                const firstInitial = parts[0][0];
                const lastInitial = parts.length > 1 ? parts[parts.length - 1][0] : '';
                const initials = (firstInitial + lastInitial).toUpperCase();
                debugLog(`Got initials "${initials}" directly from DOM name "${fullName}"`);
                // Cache this name for future use
                localStorage.setItem('cachedUserName', fullName);
                return initials;
            }
        }

        // If we still don't have a name, trigger a cache update for next time
        debugLog("No user name found, triggering cache update");
        cacheUserName();

        // Return ?? while we wait for cache to update
        return "??";
    }

    // Format Note (from Quote Converter)
    function formatNote(quoteNumber) {
        const now = new Date();
        const userInitials = getUserInitials();
        const date = `${now.getMonth() + 1}/${now.getDate()}`;
        let hours = now.getHours();
        const minutes = now.getMinutes().toString().padStart(2, '0');
        const ampm = hours >= 12 ? 'PM' : 'AM';
        hours = hours % 12;
        hours = hours ? hours : 12;
        const time = `${hours}:${minutes}${ampm}`;
        return `${quoteNumber}-${userInitials}-${date}-${time}`;
    }

    // Format Sales Order Note (from Quote Converter)
    function formatSalesOrderNote(orderNumber) {
        const now = new Date();
        const userInitials = getUserInitials();
        const date = `${now.getMonth() + 1}/${now.getDate()}`;
        let hours = now.getHours();
        const minutes = now.getMinutes().toString().padStart(2, '0');
        const ampm = hours >= 12 ? 'PM' : 'AM';
        hours = hours % 12;
        hours = hours ? hours : 12;
        const time = `${hours}:${minutes}${ampm}`;
        return `SO-${orderNumber}-${userInitials}-${date}-${time}`;
    }

    // Highlight Element (from Quote Converter)
    function highlightElement(element, color = '#2196F3') {
        try {
            const originalOutline = element.style.outline;
            const originalBoxShadow = element.style.boxShadow;

            element.style.outline = `2px solid ${color}`;
            element.style.boxShadow = `0 0 10px ${color}`;
            element.style.transition = 'all 0.3s ease';

            const pulseAnimation = element.animate(
                [
                    { boxShadow: `0 0 5px ${color}` },
                    { boxShadow: `0 0 15px ${color}` },
                    { boxShadow: `0 0 5px ${color}` }
                ],
                {
                    duration: 1000,
                    iterations: 2
                }
            );

            setTimeout(() => {
                element.style.outline = originalOutline;
                element.style.boxShadow = originalBoxShadow;
            }, 2500);

        } catch (e) {
            debugLog(`Error highlighting element: ${e.message}`, null, 'error');
        }
    }

    // Create Visual Notice (from Quote Converter)
    function createVisualNotice(targetElement, message = 'Filling this field...') {
        try {
            const rect = targetElement.getBoundingClientRect();
            const notice = document.createElement('div');
            notice.id = 'user-interaction-notice';
            notice.style.position = 'absolute';
            notice.style.top = `${window.scrollY + rect.top - 40}px`; // Adjust for scroll position
            notice.style.left = `${window.scrollX + rect.left}px`;  // Adjust for scroll position
            notice.style.backgroundColor = '#FFC107';
            notice.style.color = 'black';
            notice.style.padding = '5px 10px';
            notice.style.borderRadius = '3px';
            notice.style.fontSize = '12px';
            notice.style.zIndex = '9999';
            notice.style.pointerEvents = 'none';
            notice.style.animation = 'pulse-attention 1s infinite';
            notice.textContent = message;

            const style = document.createElement('style');
            style.textContent = `
                @keyframes pulse-attention {
                    0% { opacity: 0.8; }
                    50% { opacity: 1; }
                    100% { opacity: 0.8; }
                }
            `;
            document.head.appendChild(style);
            document.body.appendChild(notice);

            const originalOutline = targetElement.style.outline;
            targetElement.style.outline = '2px solid #FFC107';

            setTimeout(() => {
                if (document.body.contains(notice)) {
                    notice.remove();
                }
                targetElement.style.outline = originalOutline;
            }, 5000);

        } catch (e) {
            debugLog(`Error creating user interaction notice: ${e.message}`, null, 'error');
        }
    }

    // Generate a unique tab ID (from Quote Converter)
    const TAB_ID = Date.now().toString() + Math.random().toString().substring(2, 8);

    // Global flags and variables (from Master Script)
    let creditHoldPopupDismissed = false;
    let lastCardCustomerName = "";
    let inFlightCustomerName = "";
    let autoAssigned = false; // Flag if the current sales order already has an assignment
    let quoteConversionHappened = false; // Flag to disable quote behavior after conversion

    // Store the last created product name (from Master Script Part 2)
    let lastCreatedProductName = '';

    // URLs for published CSVs (from Master Script)
    const blanketPOSheetUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ-iA8XppRLS2s-30ElJSdvNbrgOQB-kaYIS39uC1TwvHYS4HgmZKMbSR1vUuekeiYHlQaYsFuJk3P9/pub?output=csv";
    const cardSheetUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSyRPtuexxvaAtYT6t0aTJ-I8f4vYIvd5TxkY-HU6XeTSIltadeXH6Y5N0L97lvHJidpoROBSXPob1Z/pub?output=csv";
    // Google Apps Script Web App URL (from Quote Converter)
    const webAppUrl = 'https://script.google.com/macros/s/AKfycbxy4Ed1EXZ-bE-oYWSOu3v8DJ78TH-ptQdRQ9gsU0Nz844dTV6DZN5dIjwawNYSZFI7lA/exec';
    // Google Sheets CSV URL for verification (from Quote Converter)
    const verificationCsvUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR8DsfQHPdd7RK47qTj7BvakCgfDxEBI50kqLFQLeozYlHziNZH3HQXCKMcViPm3mh-o_bZGs1ssrR0/pub?output=csv";


    // =========================================
    // PART 1: Quote Converter Functionality
    // =========================================

    // Function to verify if data was added to the spreadsheet
    function verifySheetUpdate(quoteNumber, orderNumber, showResults = false) {
        debugLog(`Verifying data in sheet: Q#${quoteNumber} & SO#${orderNumber}`);
        if (showResults) {
            showStatusNotification(`Verifying data in Google Sheets...`, 'loading');
        }

        GM_xmlhttpRequest({
            method: 'GET',
            url: verificationCsvUrl, // Use the specific CSV URL for verification
            onload: function(response) {
                try {
                    debugLog(`CSV fetch status: ${response.status}`);
                    if (response.status !== 200) {
                        debugLog(`Failed to fetch CSV: Status ${response.status}`, response, 'error');
                        if (showResults) showStatusNotification(`Failed to fetch CSV: Status ${response.status}`, 'error');
                        return;
                    }
                    const csvData = response.responseText;
                    debugLog(`Received CSV data (${csvData.length} bytes)`, csvData.substring(0, 200) + '...');
                    const rows = csvData.split(/\r?\n/);
                    debugLog(`CSV contains ${rows.length} rows`);
                    rows.slice(0, Math.min(5, rows.length)).forEach((row, i) => debugLog(`CSV Row ${i}: ${row}`));

                    let found = false;
                    for (const row of rows) {
                        const cells = row.split(',');
                        if (cells.length >= 2 && cells[0]?.trim() === quoteNumber?.trim() && cells[1]?.trim() === orderNumber?.trim()) {
                            found = true;
                            debugLog(`FOUND match in row: ${row}`);
                            break;
                        }
                    }

                    if (found) {
                        debugLog('Data verified in Google Sheets ✓');
                        if (showResults) showStatusNotification(`Successfully verified Q#${quoteNumber} & SO#${orderNumber} in sheet`, 'success');
                    } else {
                        debugLog('Data NOT found in Google Sheets ✗', null, 'warn');
                        const debugInfo = `Looking for: Q#${quoteNumber} & SO#${orderNumber}\nRows: ${rows.length}\nFirst few rows: ${rows.slice(0, 3).join('\n')}`;
                        debugLog('Debug CSV info:', debugInfo);
                        if (showResults) showStatusNotification(`Data not found in Google Sheets. Check the console for details.`, 'error');
                    }
                } catch (e) {
                    debugLog('Error parsing CSV data', e, 'error');
                    if (showResults) showStatusNotification(`Error verifying data: ${e.message}`, 'error');
                }
            },
            onerror: function(error) {
                debugLog('Error fetching CSV', error, 'error');
                if (showResults) showStatusNotification(`Error fetching CSV: ${error.message}`, 'error');
            }
        });
    }

    // Function to track and prevent duplicate submissions
    function trackSubmission(quoteNumber, orderNumber) {
        try {
            const submissions = JSON.parse(localStorage.getItem('sheetsSubmissions') || '{}');
            const submissionKey = `${quoteNumber}_${orderNumber}`;
            if (submissions[submissionKey]) {
                debugLog(`Preventing duplicate submission of Q#${quoteNumber} & SO#${orderNumber}`);
                return false; // Already submitted
            }
            submissions[submissionKey] = Date.now();
            localStorage.setItem('sheetsSubmissions', JSON.stringify(submissions));
            debugLog(`Tracking new submission: Q#${quoteNumber} & SO#${orderNumber}`);
            return true; // New submission
        } catch (e) {
            debugLog('Error tracking submission', e, 'error');
            return true; // On error, allow submission
        }
    }

    // Function to create iframe for Google Sheets integration with URLs
    function createGoogleSheetsIframeWithUrls(quoteNumber, orderNumber, quoteUrl, orderUrl) {
        debugLog(`Creating iframe for Google Sheets with URLs: Q#${quoteNumber} -> SO#${orderNumber}`);
        debugLog(`Quote URL: ${quoteUrl}`);
        debugLog(`Order URL: ${orderUrl}`);

        showStatusNotification(`Sending data to Google Sheets...`, 'loading');

        // Store attempt
        try {
            const attempts = JSON.parse(localStorage.getItem('sheetSubmitAttempts') || '[]');
            attempts.push({ timestamp: new Date().toISOString(), quoteNumber, orderNumber, quoteUrl, orderUrl });
            localStorage.setItem('sheetSubmitAttempts', JSON.stringify(attempts));
        } catch (e) { debugLog('Error storing attempt info', e, 'error'); }

        let fullUrl = `${webAppUrl}?quoteNumber=${encodeURIComponent(quoteNumber)}&orderNumber=${encodeURIComponent(orderNumber)}`;
        if (quoteUrl && orderUrl) {
            fullUrl += `&quoteUrl=${encodeURIComponent(quoteUrl)}&orderUrl=${encodeURIComponent(orderUrl)}`;
        }
        debugLog(`Prepared URL for submission: ${fullUrl}`);

        function createIframeForSubmission(url) {
            debugLog('Creating hidden iframe for submission');
            const iframe = document.createElement('iframe');
            iframe.id = 'google-sheets-iframe';
            iframe.style.width = '1px'; // Keep hidden unless debugging
            iframe.style.height = '1px';
            iframe.style.position = 'fixed';
            iframe.style.bottom = '-10px'; // Off screen
            iframe.style.left = '-10px';
            iframe.style.border = 'none'; // No border
            iframe.style.zIndex = '9999';
            iframe.src = url;

            iframe.addEventListener('load', () => {
                debugLog(`Google Sheets iframe loaded - URL: ${iframe.src}`);
                try {
                    const iframeContent = iframe.contentDocument || iframe.contentWindow.document;
                    debugLog('Attempted to access iframe content (may fail due to CORS)', iframeContent?.body?.innerHTML?.substring(0, 100));
                } catch (e) { debugLog('Cannot access iframe content due to security restrictions', e, 'warn'); }
                showStatusNotification(`Data sent to Google Sheets (iframe loaded)`, 'success');
                setTimeout(() => { if (document.body.contains(iframe)) document.body.removeChild(iframe); }, 5000); // Remove after 5s
                setTimeout(() => verifySheetUpdate(quoteNumber, orderNumber, true), 3000); // Verify after 3s
            });

            iframe.addEventListener('error', (e) => {
                debugLog(`Google Sheets iframe error`, e, 'error');
                showStatusNotification(`Error loading Google Sheets iframe`, 'error');
                if (document.body.contains(iframe)) document.body.removeChild(iframe);
            });

            document.body.appendChild(iframe);
            debugLog('Added iframe to document body');
            setTimeout(() => { debugLog('Timeout-based verification starting'); verifySheetUpdate(quoteNumber, orderNumber, true); }, 8000); // Backup verification
        }

        // Try direct fetch first
        try {
            debugLog('Attempting direct fetch first');
            fetch(fullUrl, { method: 'GET', mode: 'no-cors' }) // Use no-cors for web app
                .then(response => {
                    // no-cors responses are opaque, status is 0, but it might have worked
                    debugLog(`Direct fetch (no-cors) completed. Status: ${response.status}, Type: ${response.type}`);
                    if (response.type === 'opaque') {
                       showStatusNotification(`Data sent directly to Google Sheets (no-cors)!`, 'success');
                       setTimeout(() => verifySheetUpdate(quoteNumber, orderNumber, true), 3000); // Verify after 3s
                    } else {
                       // Fallback if direct fetch wasn't opaque or failed unexpectedly
                       debugLog('Direct fetch (no-cors) did not result in opaque response, falling back to iframe', response, 'warn');
                       createIframeForSubmission(fullUrl);
                    }
                })
                .catch(err => {
                    debugLog('Direct fetch (no-cors) failed, falling back to iframe', err, 'warn');
                    createIframeForSubmission(fullUrl);
                });
        } catch (e) {
            debugLog('Error in fetch attempt', e, 'error');
            createIframeForSubmission(fullUrl);
        }
        return true;
    }

     // Function to send mappings to Google Sheets using the appropriate method
    function sendMappingToGoogleSheets(quoteNumber, orderNumber, quoteUrl, orderUrl) {
        try {
            if (!trackSubmission(quoteNumber, orderNumber)) {
                debugLog(`Skipping duplicate submission to Google Sheets: Q#${quoteNumber} -> SO#${orderNumber}`);
                showStatusNotification(`Data already sent to Google Sheets`, 'success');
                return;
            }
            log(`Automatically sending mapping to Google Sheets: Q#${quoteNumber} -> SO#${orderNumber}`);

            // Save mapping locally
            const mappings = JSON.parse(localStorage.getItem('pendingSheetUpdates') || '[]');
            mappings.push({ quoteNumber, orderNumber, quoteUrl, orderUrl, timestamp: Date.now() });
            localStorage.setItem('pendingSheetUpdates', JSON.stringify(mappings));

            // Use the iframe method with URLs
            createGoogleSheetsIframeWithUrls(quoteNumber, orderNumber, quoteUrl, orderUrl);

        } catch (e) {
            debugLog(`Error sending to Google Sheets: ${e.message}`, null, 'error');
            showStatusNotification(`Error sending to Google Sheets: ${e.message}`, 'error');
        }
    }

    // Add debug button
    /* function addDebugButton() {
        if (document.getElementById('inflow-debug-button')) return; // Avoid duplicates
        const button = document.createElement('button');
        button.id = 'inflow-debug-button';
        button.textContent = 'Debug Sheet';
        button.style.position = 'fixed';
        button.style.bottom = '5px';
        button.style.left = '5px';
        button.style.zIndex = '9999';
        button.style.padding = '5px';
        button.style.backgroundColor = '#f44336';
        button.style.color = 'white';
        button.style.border = 'none';
        button.style.borderRadius = '4px';
        button.style.cursor = 'pointer';
        button.addEventListener('click', showDebugPanel);
        document.body.appendChild(button);
    } */

    // Show debug panel
    /* function showDebugPanel() {
        if (document.getElementById('inflow-debug-panel')) return;
        const panel = document.createElement('div');
        panel.id = 'inflow-debug-panel'; // Add ID to prevent duplicates
        panel.style.position = 'fixed';
        panel.style.top = '50px';
        panel.style.left = '50px';
        panel.style.width = '80%';
        panel.style.height = '80%';
        panel.style.backgroundColor = 'white';
        panel.style.border = '1px solid black';
        panel.style.padding = '10px';
        panel.style.zIndex = '10000';
        panel.style.overflow = 'auto';

        const closeBtn = document.createElement('button');
        closeBtn.textContent = 'Close';
        closeBtn.style.position = 'absolute';
        closeBtn.style.top = '10px';
        closeBtn.style.right = '10px';
        closeBtn.addEventListener('click', () => panel.remove());

        const testBtn = document.createElement('button');
        testBtn.textContent = 'Test Sheet Integration';
        testBtn.style.marginRight = '10px';
        testBtn.addEventListener('click', () => {
            const testQuote = `TEST-${Date.now()}`;
            const testOrder = `TEST-ORDER-${Date.now()}`;
            debugLog(`Running manual test with Q#${testQuote} and SO#${testOrder}`);
            sendMappingToGoogleSheets(testQuote, testOrder, 'http://test.quote.url', 'http://test.order.url');
        });

        const verifyBtn = document.createElement('button');
        verifyBtn.textContent = 'Verify Latest Entry';
        verifyBtn.addEventListener('click', () => {
            try {
                const attempts = JSON.parse(localStorage.getItem('sheetSubmitAttempts') || '[]');
                if (attempts.length > 0) {
                    const latest = attempts[attempts.length - 1];
                    debugLog(`Manually verifying latest entry: Q#${latest.quoteNumber} -> SO#${latest.orderNumber}`);
                    verifySheetUpdate(latest.quoteNumber, latest.orderNumber, true);
                } else {
                    debugLog('No previous attempts found', null, 'warn'); alert('No previous submission attempts found');
                }
            } catch (e) { debugLog('Error loading attempts', e, 'error'); }
        });

        const directUrlBtn = document.createElement('button');
        directUrlBtn.textContent = 'Test Direct URL';
        directUrlBtn.style.marginLeft = '10px';
        directUrlBtn.addEventListener('click', () => {
            const testQuote = `DIRECT-${Date.now()}`;
            const testOrder = `DIRECT-ORDER-${Date.now()}`;
            const testQuoteUrl = `https://app.inflowinventory.com/quotes/test-${Date.now()}`;
            const testOrderUrl = `https://app.inflowinventory.com/sales-orders/test-${Date.now()}`;
            let fullUrl = `${webAppUrl}?quoteNumber=${encodeURIComponent(testQuote)}&orderNumber=${encodeURIComponent(testOrder)}`;
            fullUrl += `&quoteUrl=${encodeURIComponent(testQuoteUrl)}&orderUrl=${encodeURIComponent(testOrderUrl)}`;
            window.open(fullUrl, '_blank');
        });

        const content = document.createElement('div');
        content.innerHTML = '<h3>Debug Logs (Last 20)</h3>';
        try {
            const logs = JSON.parse(localStorage.getItem('inflowDebugLogs') || '[]');
            if (logs.length > 0) {
                content.innerHTML += logs.slice(-20).map(log => {
                    const color = log.level === 'error' ? 'red' : (log.level === 'warn' ? 'orange' : 'black');
                    const dataStr = log.data ? `<pre style="margin: 5px 0; font-size: 11px; background: #f5f5f5; padding: 5px; white-space: pre-wrap; word-wrap: break-word;">${typeof log.data === 'string' ? log.data : JSON.stringify(log.data, null, 2)}</pre>` : '';
                    return `<div style="margin-bottom: 5px; color: ${color}"><strong>${new Date(log.timestamp).toLocaleTimeString()}</strong>: ${log.message} ${dataStr}</div>`;
                }).join('');
            } else { content.innerHTML += '<p>No logs found</p>'; }
        } catch (e) { content.innerHTML += `<p style="color: red">Error loading logs: ${e.message}</p>`; }

        content.innerHTML += '<h3>Sheet Submission Attempts (Last 10)</h3>';
        try {
            const attempts = JSON.parse(localStorage.getItem('sheetSubmitAttempts') || '[]');
            if (attempts.length > 0) {
                content.innerHTML += attempts.slice(-10).map(attempt =>
                    `<div style="margin-bottom: 5px;"><strong>${new Date(attempt.timestamp).toLocaleString()}</strong>: Q: ${attempt.quoteNumber} → O: ${attempt.orderNumber}${attempt.quoteUrl ? `<br><small>Q URL: ${attempt.quoteUrl}</small>` : ''}${attempt.orderUrl ? `<br><small>O URL: ${attempt.orderUrl}</small>` : ''}</div>`
                ).join('');
            } else { content.innerHTML += '<p>No submission attempts found</p>'; }
        } catch (e) { content.innerHTML += `<p style="color: red">Error loading attempts: ${e.message}</p>`; }

        const buttonRow = document.createElement('div');
        buttonRow.style.marginBottom = '15px';
        buttonRow.appendChild(testBtn);
        buttonRow.appendChild(verifyBtn);
        buttonRow.appendChild(directUrlBtn);

        panel.appendChild(closeBtn);
        panel.appendChild(buttonRow);
        panel.appendChild(content);
        document.body.appendChild(panel);
    } */

    // Capture Quote Number
    function captureQuoteNumber() {
        const quoteNumberInput = document.getElementById('entity-header-input');
        if (!quoteNumberInput) {
            debugLog("Quote number input field not found", null, 'warn');
            return null;
        }
        const quoteNumber = quoteNumberInput.value;
        debugLog(`Captured quote number from field: ${quoteNumber}`);
        return quoteNumber;
    }

    // Click "As sales order" option
    function clickAsSalesOrderOption() {
        let asSalesOrderOption = document.getElementById('salesOrder') || document.getElementById('option-item-salesOrder');
        if (!asSalesOrderOption) {
            const spans = document.querySelectorAll('span');
            for (const span of spans) {
                if (span.textContent.includes('As sales order')) {
                    asSalesOrderOption = span.closest('div[role="button"], button, a') || span; // Find clickable parent
                    break;
                }
            }
        }

        if (asSalesOrderOption) {
            debugLog("Found 'As sales order' option, clicking it now");
            asSalesOrderOption.click();

            let attempts = 0;
            const maxAttempts = 15;
            const checkForOpenButton = setInterval(() => {
                attempts++;
                const openNewTabButton = document.querySelector('#modalOK');
                if (openNewTabButton) {
                    clearInterval(checkForOpenButton);
                    debugLog("Found 'Open in new tab' button, clicking it");
                    const newWindowName = 'orderWindow_' + Date.now();
                    localStorage.setItem('newOrderWindowName', newWindowName);
                    openNewTabButton.setAttribute('target', newWindowName); // Ensure it opens in the named window
                    openNewTabButton.click();

                    setTimeout(() => {
                        const closeButton = document.querySelector('#modalCancel');
                        if (closeButton) {
                            debugLog("Found 'Close' button, clicking it");
                            closeButton.click();
                        } else { debugLog("Close button not found after clicking Open in new tab", null, 'warn'); }
                    }, 500);
                } else if (attempts >= maxAttempts) {
                    clearInterval(checkForOpenButton);
                    debugLog("Failed to find 'Open in new tab' button", null, 'error');
                    showError("Error: 'Open in new tab' button not found");
                }
            }, 300);
        } else {
            debugLog("'As sales order' option not found yet, will retry in 500ms");
            setTimeout(clickAsSalesOrderOption, 500); // Retry
        }
    }

    // Handle click on Copy & Convert button
    function handleCopyConvertClick(event) {
        if (event) { event.preventDefault(); event.stopPropagation(); }
        debugLog(`Starting Copy & Convert process from original tab (ID: ${TAB_ID})`);
        localStorage.setItem('originalTabId', TAB_ID);
        localStorage.setItem('isOriginalQuote_' + window.location.pathname, 'true');
        localStorage.setItem('originalQuoteUrl', window.location.href);

        const displayedQuoteNumber = captureQuoteNumber();
        if (!displayedQuoteNumber) { showError("Error: Could not capture quote number from page"); return; }

        const formattedNote = formatNote(displayedQuoteNumber);
        const conversionInfo = {
            tabId: TAB_ID,
            quoteNumber: displayedQuoteNumber,
            formattedNote: formattedNote,
            originalQuoteId: window.location.pathname.split('/').pop(),
            originalQuoteUrl: window.location.href,
            originalWindowName: window.name || 'originalQuoteWindow',
            timestamp: Date.now()
        };
        localStorage.setItem('pendingConversion', JSON.stringify(conversionInfo));
        debugLog(`Stored conversion data: ${JSON.stringify(conversionInfo)}`);

        const copyButtons = Array.from(document.querySelectorAll('button.sc-dd788c24-0'));
        let copyButton = copyButtons.find(btn => btn.textContent.includes('Copy') && !btn.textContent.includes('Copy & Convert'));

        if (!copyButton) {
            // Fallback selector if the class changed
             copyButton = Array.from(document.querySelectorAll('button')).find(btn => btn.textContent.includes('Copy') && !btn.textContent.includes('Copy & Convert'));
        }

        if (!copyButton) { debugLog("Copy button not found", null, 'error'); showError("Error: Copy button not found"); return; }

        debugLog("Clicking Copy button to open dropdown");
        copyButton.click();

        const marker = document.createElement('div');
        marker.id = 'automation-already-started'; // Marker used in both processCopiedQuote and here
        marker.style.display = 'none';
        document.body.appendChild(marker);

        setTimeout(clickAsSalesOrderOption, 700);
    }

    // Helper to find input field (robustly)
    function findInputField() {
        debugLog("Attempting to find input field..."); // Log start of find
        const selectors = [
            '#custom1-input', // Direct ID
            'input[name="custom_1"]', // Common name attribute
             'input[placeholder="Enter data"][aria-label*="Converted Quote Note"]', // Placeholder + Aria Label
             'textarea[aria-label*="Converted Quote Note"]', // Try Textarea too
        ];
        for (const selector of selectors) {
            const input = document.querySelector(selector);
            if (input) { debugLog(`Found field by selector: ${selector}`); return input; }
        }

        // Find by label text proximity
        const labels = Array.from(document.querySelectorAll('div, p, label'));
        for (const label of labels) {
            // Look for the specific label text
            if (label.textContent.trim().toUpperCase() === 'CONVERTED QUOTE NOTE') {
                debugLog("Found label with exact text: CONVERTED QUOTE NOTE");
                // Try to find the input within the label's container or siblings
                let container = label.closest('.sc-5bf8f203-2, .form-field, .input-group') || label.parentElement;
                for (let i = 0; i < 3 && container; i++) { // Check siblings/children within 3 levels
                    const input = container.querySelector('input[placeholder="Enter data"], textarea[placeholder="Enter data"]');
                    if (input) {
                        debugLog(`Found input near the label (Selector: ${input.tagName} placeholder="Enter data")`);
                        return input;
                    }
                    container = container.parentElement;
                }
            }
        }
        debugLog("Could not find CONVERTED QUOTE NOTE input field with any method", null, 'warn');
        // ** ADD LOG **
        // debugLog("[findInputField Diagnostics] Returning null."); // Keep this if detailed logging is useful
        return null;
    }


     // Helper function to dispatch all necessary events
    function dispatchAllEvents(input) {
        const events = [
            new Event('input', {bubbles: true, cancelable: true}),
            new Event('change', {bubbles: true, cancelable: true}),
            new KeyboardEvent('keydown', {bubbles: true}), // <<< RESTORED
            new KeyboardEvent('keypress', {bubbles: true}), // <<< RESTORED
            new KeyboardEvent('keyup', {bubbles: true}), // <<< RESTORED
            new Event('focus', {bubbles: true}), // <<< RESTORED
            new Event('blur', {bubbles: true})
        ];
        for (const event of events) { input.dispatchEvent(event); }

        try {
            const reactKey = Object.keys(input).find(key => key.startsWith('__reactProps') || key.startsWith('__reactFiber'));
            if (reactKey) {
                const reactProps = input[reactKey];
                const handler = reactProps?.onChange || reactProps?.memoizedProps?.onChange || reactProps?.props?.onChange;
                if (handler && typeof handler === 'function') {
                     // Create a minimal event object that might satisfy the handler
                     const eventArg = {
                        target: input,
                        currentTarget: input,
                        type: 'change', // Mimic a change event
                        bubbles: true,
                        cancelable: true,
                        isTrusted: false, // Indicate it's synthetic
                        nativeEvent: { target: input }, // Provide nested nativeEvent if needed
                        preventDefault: () => {},
                        stopPropagation: () => {}
                        // Add other properties if console errors indicate they are needed
                    };
                    handler(eventArg);
                    debugLog("Attempted to call React onChange handler directly");
                } else {
                    log("React onChange handler not found or not a function.");
                }
            } else {
                 log("React properties key not found on input element.");
            }
        } catch (e) {
            log("Error calling React handler: " + e.message, null, 'warn');
        }

        setTimeout(() => { input.blur(); setTimeout(() => { input.focus(); document.body.click(); }, 100); }, 100);
    }


    // Populate Converted Quote Note Field (robustly) - Returns true if field found & attempt made
    function populateConvertedQuoteNote(noteText, isUpdateWindow = false) {
        // ** ADD LOG **
        debugLog("[populateConvertedQuoteNote Start] Received noteText: " + noteText + ` (isUpdateWindow: ${isUpdateWindow})`); // <<< LOG PARAMETER
        debugLog("Attempting to fill note field with: \"" + noteText + "\" using multiple methods");
        const input = findInputField();
        if (!input) {
            // REMOVED: showError("Could not find the 'CONVERTED QUOTE NOTE' field.");
             // ** ADD LOG **
             debugLog("[populateConvertedQuoteNote Error] findInputField returned null.");
            return false; // Indicate failure: field not found
        }

        createVisualNotice(input, "Setting note...");
        // window.focus(); // REMOVE THIS - Avoid potentially disruptive global focus changes
        // input.focus(); // Let the caller manage focus if needed
        // input.click(); // Avoid extra clicks

        let valueSetAndVerified = false; // Track if value is confirmed

        // Method 1: Direct set + events (Simplified)
        try {
            debugLog("Trying Simplified Method: Direct set and dispatch input/change");
            input.value = ''; // Clear first
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.value = noteText;
            // --- REVERT: Call original dispatchAllEvents ---
            dispatchAllEvents(input);
            // input.dispatchEvent(new Event('input', { bubbles: true })); // Covered by dispatchAllEvents
            // input.dispatchEvent(new Event('change', { bubbles: true })); // Covered by dispatchAllEvents

            // --- REMOVED: React handler call attempt (now part of dispatchAllEvents logic conceptually) ---
            /*
            try {
                const reactKey = Object.keys(input).find(key => key.startsWith('__reactProps') || key.startsWith('__reactFiber'));
                if (reactKey) {
                    const reactProps = input[reactKey];
                    const handler = reactProps?.onChange || reactProps?.memoizedProps?.onChange || reactProps?.props?.onChange;
                    if (handler && typeof handler === 'function') {
                         const eventArg = { target: input, currentTarget: input, type: 'change', bubbles: true, cancelable: true, isTrusted: false, nativeEvent: { target: input }, preventDefault: () => {}, stopPropagation: () => {} };
                        handler(eventArg);
                        debugLog("Attempted to call React onChange handler directly in populateConvertedQuoteNote");
                    }
                }
            } catch (e) {
                debugLog("Error calling React handler in populateConvertedQuoteNote: " + e.message, null, 'warn');
            }
            */
            // --- END REVERT ---
            
            // --- REMOVED: Explicit blur dispatch (now part of dispatchAllEvents) ---
            // input.dispatchEvent(new Event('blur', { bubbles: true }));
            // debugLog("Dispatched blur event after setting note value");
            // --- END REMOVED ---

            // Verification (Crucial) - Use a small delay to allow UI to potentially update
            setTimeout(() => {
                 if (input.value === noteText) {
                     debugLog("Simplified Method: Value verified after delay.");
                     valueSetAndVerified = true; // Keep this flag update

                     // --- REMOVED: Secondary delay, call check immediately ---
                     // setTimeout(() => {
                          if (valueSetAndVerified) { // Re-check flag (good practice)
                             debugLog("Verification succeeded, calling checkForSaveButton immediately.");
                             checkForSaveButton(isUpdateWindow); // <<< Call immediately, pass flag
                          }
                     // }, 250); // Removed 250ms delay
                     // --- END REMOVED ---

                 } else {
                      debugLog("Simplified Method: Value verification failed after delay.", input.value, 'warn');
                      valueSetAndVerified = false; // Explicitly set false
                      showError("Failed to automatically fill and verify the note field."); // Show error on verification failure
                 }
            }, 400); // INCREASED verification delay to 400ms

            // For flow control, we initially assume it *might* work and rely on the timeout verification
             valueSetAndVerified = true; // Tentative true for return value, verification happens async


        } catch (e) {
            debugLog(`Simplified Method Error: ${e.message}`, null, 'error');
            showError("Error automatically filling the note field.");
            valueSetAndVerified = false; // Ensure false on error
        }




        // This log might be misleading now due to async verification
        // debugLog("[populateConvertedQuoteNote Finish] Returning tentative: " + valueSetAndVerified);
        return valueSetAndVerified; // Return tentative success, actual save triggered by verification timeout
    }

    // Check for Save Button
    function checkForSaveButton(isQuoteUpdate = false) { // Added isQuoteUpdate flag
        debugLog(`Checking for Save button (isQuoteUpdate: ${isQuoteUpdate})`);
        // **REMOVED**: localStorage.setItem('saveInProgress', 'true'); // Set this only if triggering completion check

        const saveSelectors = [
            '#save', '#saveButton', '#btnSave', // IDs
            'button.sc-dfdd4e89-1.sc-dfdd4e89-2', // Class from user HTML
            'button[data-testid="save-button"]', // Test ID
            'button span:contains("Save")', // Span text (case sensitive)
            'button:contains("Save")', // Button text (case insensitive - jQuery style)
             'button[aria-label*="Save"]' // Aria label
        ];

        let saveButton = null;
        for (const selector of saveSelectors) {
             // Use querySelector for IDs and standard CSS, handle :contains separately
             if (selector.includes(':contains')) {
                 const text = selector.match(/:contains\("(.+?)"\)/)[1];
                 const buttons = Array.from(document.querySelectorAll('button'));
                 saveButton = buttons.find(btn => btn.textContent.includes(text) && btn.offsetParent !== null); // Check visible
             } else {
                 saveButton = document.querySelector(selector);
             }

             if (saveButton && saveButton.offsetParent !== null) { // Check if visible
                 debugLog(`Found Save button using selector: ${selector}`);
                 break;
             } else {
                 saveButton = null; // Reset if not found or not visible
             }
         }

        if (!saveButton) {
            // Last resort: Look for any button in the lower part of the screen containing "Save"
            const allButtons = document.querySelectorAll('button');
            for (const btn of allButtons) {
                const rect = btn.getBoundingClientRect();
                if (rect.top > (window.innerHeight * 0.7) && btn.textContent.includes('Save') && btn.offsetParent !== null) {
                    saveButton = btn;
                    debugLog(`Found possible Save button in bottom area: ${btn.textContent}`);
                    break;
                }
            }
        }

        if (saveButton) {
            highlightElement(saveButton, '#4CAF50');
            debugLog("Save button found - will click it");

            setTimeout(() => {
                try {
                    if (saveButton.offsetParent === null || saveButton.disabled) {
                         debugLog("Save button became hidden or disabled before click.", null, 'warn');
                         // If it's the quote update, just close after delay
                         if (isQuoteUpdate) {
                             setTimeout(() => closeWindowAfterSave(true), 1500);
                         }
                         return;
                     }
                    saveButton.click();
                    debugLog("Clicked Save button");
                    showStatusNotification("Changes saving...", "loading"); // Indicate saving started

                    // --- MODIFIED: Bypass detection for quote update ---
                    if (isQuoteUpdate) {
                        debugLog("Quote update context: Bypassing save completion detection, closing after fixed delay.");
                        // Use a reasonable fixed delay to allow save before closing
                        setTimeout(() => closeWindowAfterSave(true), 4000); // Close after 4 seconds
                    } else {
                         // Original logic for other contexts (e.g., copied quote initial save)
                         localStorage.setItem('saveInProgress', 'true'); // Set flag only when using detection
                         detectSaveCompletion(); // Start monitoring for completion
                    }
                    // --- END MODIFIED ---

                } catch (e) {
                    debugLog(`Error clicking Save button: ${e.message}`, null, 'error');
                    showStatusNotification("Failed to click Save button", "error");
                    // Only close automatically if it's the quote update context
                    if (isQuoteUpdate) {
                         setTimeout(closeWindowAfterSave, 1500); // Close quote update window anyway after delay
                    }
                }
            }, 800); // Delay before click
        } else {
            debugLog("Save button not yet found, will try again");
            if (document.getElementById('save-button-clicked')) {
                debugLog("Save has already been clicked once, won't search again");
                setTimeout(closeWindowAfterSave, 1500); // Close after delay
                return;
            }
            const currentTries = parseInt(localStorage.getItem('saveButtonSearchTries') || '0');
            if (currentTries < 10) {
                localStorage.setItem('saveButtonSearchTries', (currentTries + 1).toString());
                setTimeout(checkForSaveButton, 800);
            } else {
                debugLog("Gave up searching for save button after 10 attempts", null, 'error');
                showStatusNotification("Could not find Save button. Please save manually.", "error");
                localStorage.removeItem('saveButtonSearchTries');
                setTimeout(closeWindowAfterSave, 1500); // Close anyway
            }
        }
    }

    // Detect Save Completion
    function detectSaveCompletion() {
        debugLog("Monitoring for save completion");
        // ** ADDED: Reference to the originating window if this is the quote update window **
        const isQuoteUpdateWin = window.name === 'quoteUpdateWindow';
        let closeHandled = false;
        let observer = null; // Define observer variable

        const handleClose = () => {
            if (!closeHandled) {
                closeHandled = true;
                if (observer) observer.disconnect();
                debugLog("Save monitoring complete/timed out - closing window shortly");
                showStatusNotification("Save complete!", "success"); // Update status
                // ** CHANGE: Call closeWindowAfterSave directly here, passing the flag **
                closeWindowAfterSave(isQuoteUpdateWin);
            }
        };

        // Method 1: Look for visual indicators (button disabled/gone, success message)
        observer = new MutationObserver((mutations) => {
            if (closeHandled) return;
            let saveComplete = false;
            const saveButton = document.querySelector('#save, #saveButton, #btnSave, button:contains("Save")'); // Re-find button
            if (!saveButton || saveButton.disabled || saveButton.offsetParent === null) {
                debugLog("Save button is now disabled/hidden - save likely complete");
                saveComplete = true;
            }
            const successElements = document.querySelectorAll('.toast, .notification, .alert, [class*="success"]');
            for (const el of successElements) {
                if (el.textContent.toLowerCase().includes('saved') || el.textContent.toLowerCase().includes('success')) {
                    debugLog("Found success message - save is complete");
                    saveComplete = true;
                    break;
                }
            }
             // Check for specific loading spinners disappearing
             const spinners = document.querySelectorAll('.spinner, .loader, [class*="loading"]');
             if (localStorage.getItem('saveInProgress') === 'true' && spinners.length === 0) {
                  debugLog("Loading indicators disappeared - save likely complete");
                 // saveComplete = true; // This might be too aggressive, rely on other indicators
             }

            if (saveComplete) handleClose();
        });

        observer.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['disabled', 'style', 'class'] });

        // Method 2: Backup timeout
        setTimeout(() => {
             if (!closeHandled) {
                // **CHANGE**: Increased timeout duration
                debugLog("Closing window after save timeout (7s)");
                handleClose();
             }
        }, 7000); // Increased timeout to 7 seconds
    }

    // Close Window After Save
    function closeWindowAfterSave(isQuoteUpdate = false) {
        debugLog("Preparing to close window after save");
        // This function is called by detectSaveCompletion. The `isQuoteUpdate` flag
        // indicates if this is the specific quote update window.

        if (isQuoteUpdate) { // Check the passed flag
            debugLog("Closing the dedicated quote update window now.");

            // Clean up local storage related to this specific update attempt if needed
            // (These might be cleaned elsewhere too, but doesn't hurt to ensure)
            localStorage.removeItem('saveInProgress');
            localStorage.removeItem('saveCompleted');
            localStorage.removeItem('saveButtonSearchTries');
            // pendingQuoteUpdate should be removed earlier, but clear just in case
            localStorage.removeItem('pendingQuoteUpdate');

            // --- ADDED: Set flag for Sales Order refresh ---
            debugLog("Setting quote completion flag in localStorage before closing update window.");
            localStorage.setItem('quoteConversionCompletedTimestamp', Date.now().toString());
            // --- END ADDED ---

            try {
                window.close();
            } catch (e) {
                debugLog(`Error closing window: ${e.message}`, null, 'warn');
            }
        } else {
            debugLog(`Window name is "${window.name}", not "quoteUpdateWindow". Not closing this tab.`);
            // Clean up general save flags if they were set in a non-update context
            // (e.g., saving the copied quote note before conversion)
            localStorage.removeItem('saveInProgress');
            localStorage.removeItem('saveCompleted');
            localStorage.removeItem('saveButtonSearchTries');
        }
    }

    // Check if this is an original quote page
    function isOriginalQuotePage() {
        if (localStorage.getItem('originalTabId') === TAB_ID) return true;
        if (localStorage.getItem('isOriginalQuote_' + window.location.pathname) === 'true') return true;
        if (localStorage.getItem('originalQuoteUrl') === window.location.href) return true;
        try {
            const updateInfoStr = localStorage.getItem('pendingQuoteUpdate');
            if (updateInfoStr) {
                const updateInfo = JSON.parse(updateInfoStr);
                if (window.location.pathname.split('/').pop() === updateInfo.originalQuoteId) return true;
            }
        } catch (e) { debugLog("Error parsing pendingQuoteUpdate in isOriginalQuotePage: " + e.message, null, 'warn'); }
        return false;
    }

    // Get Stored Conversion Info
    function getStoredConversionInfo() {
        try {
            const conversionInfoStr = localStorage.getItem('pendingConversion');
            if (!conversionInfoStr) return null;
            const conversionInfo = JSON.parse(conversionInfoStr);
            if ((Date.now() - conversionInfo.timestamp) > 300000) { // 5 min timeout
                debugLog("Conversion data is too old, clearing it", null, 'warn');
                localStorage.removeItem('pendingConversion');
                return null;
            }
            if (isOriginalQuotePage()) { // Safety check
                 debugLog("Attempted to get conversion info on original quote page. Aborting.", null, 'warn');
                 return null;
            }
            return conversionInfo;
        } catch (e) { debugLog(`Error getting conversion info: ${e.message}`, null, 'error'); return null; }
    }

    // Click Convert Button
    function clickConvertButton() {
        if (isOriginalQuotePage()) { // Critical safety check
            debugLog("PREVENTED CONVERSION ATTEMPT on original quote page!", null, 'error');
            showError("⚠️ This is an original quote page. Operation aborted.");
            return false;
        }
        debugLog("Looking for Convert to Order button");
        const convertBtn = document.querySelector('#convertSalesOrder');
        if (convertBtn) {
            debugLog("Found Convert to Order button, clicking it");
            setTimeout(() => {
                try {
                    localStorage.setItem('orderWindowName', window.name || 'orderWindow');
                    convertBtn.click();
                    debugLog("Clicked Convert to Order button");
                    quoteConversionHappened = true; // Set flag from Master script

                    setTimeout(() => {
                        const orderNumberInput = document.getElementById('entity-header-input');
                        if (!orderNumberInput) { debugLog("Order number input field not found after conversion", null, 'warn'); return; }

                        const fullOrderNumber = orderNumberInput.value; // e.g., SO-12345
                        debugLog(`Found sales order number: ${fullOrderNumber}`);
                        const orderNumberMatch = fullOrderNumber.match(/SO-(\d+)/);
                        const orderNumberForNote = orderNumberMatch ? orderNumberMatch[1] : fullOrderNumber; // Just number for note
                        const orderUrl = window.location.href;
                        debugLog(`Order URL: ${orderUrl}`);

                        const conversionInfo = getStoredConversionInfo(); // Get original quote info
                         if (!conversionInfo) { // Check if conversion info still exists (it should)
                             debugLog("Conversion info missing after conversion. Cannot update original quote or send to sheets.", null, 'error');
                             showError("Error: Conversion data lost. Cannot link quote/order.");
                             return;
                         }
                        localStorage.removeItem('pendingConversion'); // Clear pending conversion as it's done

                        const salesOrderNote = formatSalesOrderNote(orderNumberForNote);
                        const updateInfo = {
                            originalQuoteUrl: conversionInfo.originalQuoteUrl,
                            originalQuoteId: conversionInfo.originalQuoteId,
                            salesOrderNote: salesOrderNote,
                            originalWindowName: conversionInfo.originalWindowName,
                            currentOrderWindowName: window.name || 'currentOrderWindow',
                            quoteNumber: conversionInfo.quoteNumber, // Original quote number
                            orderNumber: orderNumberForNote, // Numeric order number
                            fullOrderNumber: fullOrderNumber, // Full SO-xxx number
                            orderUrl: orderUrl, // Sales order URL
                            timestamp: Date.now()
                        };
                        localStorage.setItem('pendingQuoteUpdate', JSON.stringify(updateInfo));
                        debugLog(`Stored update info for original quote: ${JSON.stringify(updateInfo)}`);

                        // Log mapping locally (simple key-value)
                        try {
                            const mappings = JSON.parse(localStorage.getItem('quoteToOrderMappings') || '{}');
                            mappings[conversionInfo.quoteNumber] = fullOrderNumber; // Store full order number
                            localStorage.setItem('quoteToOrderMappings', JSON.stringify(mappings));
                            debugLog(`Saved mapping locally: Quote ${conversionInfo.quoteNumber} → Order ${fullOrderNumber}`);
                        } catch (e) { debugLog(`Error saving local mapping: ${e}`, null, 'error'); }

                        // Send to Google Sheets immediately (don't wait for quote update)
                        sendMappingToGoogleSheets(updateInfo.quoteNumber, updateInfo.fullOrderNumber, updateInfo.originalQuoteUrl, updateInfo.orderUrl);

                        // Automatically open the quote window to update it
                        openQuoteWindowAndUpdate(updateInfo.originalQuoteUrl, updateInfo);

                    }, 3000); // Wait for conversion/page load
                } catch (e) {
                    debugLog("Error clicking Convert button: " + e.message, null, 'error');
                    showError(`Error clicking Convert button: ${e.message}`);
                }
            }, 500); // Delay before click
            return true;
        } else { debugLog("Convert to Order button not found", null, 'warn'); return false; }
    }

    // Open Quote Window and Update
    function openQuoteWindowAndUpdate(originalQuoteUrl, updateInfo) {
        try {
            debugLog("Automatically opening original quote to update");
            showStatusNotification("Opening original quote to update...", "loading");

            const quoteWindow = window.open(originalQuoteUrl, 'quoteUpdateWindow', 'width=1200,height=800,menubar=no,toolbar=no,location=no,resizable=yes,scrollbars=yes,status=no');

            if (!quoteWindow) {
                showError("Could not open quote window. Please check popup blocker settings.");
                 // Still attempt to send to Google Sheets if window fails to open
                 // Ensure mapping is sent even if quote update fails
                 sendMappingToGoogleSheets(updateInfo.quoteNumber, updateInfo.fullOrderNumber, updateInfo.originalQuoteUrl, updateInfo.orderUrl);
                return;
            }

            // Reset save flags for the new window context (handled within updateOriginalQuote)
            localStorage.removeItem('saveButtonSearchTries');
            localStorage.removeItem('saveCompleted');
            localStorage.removeItem('saveInProgress');

            // Backup timeout to close the quote window - reduced to 15s
            setTimeout(() => {
                try {
                    if (quoteWindow && !quoteWindow.closed) {
                        debugLog("Closing quote window after backup timeout (15s)");
                        quoteWindow.close();
                         // Ensure mapping sent if update process failed/timed out
                         sendMappingToGoogleSheets(updateInfo.quoteNumber, updateInfo.fullOrderNumber, updateInfo.originalQuoteUrl, updateInfo.orderUrl);
                    }
                } catch (e) { debugLog("Error closing quote window via timeout: " + e.message, null, 'warn'); }
            }, 15000);

        } catch (e) {
            debugLog(`Error opening quote window: ${e.message}`, null, 'error');
            showStatusNotification(`Error opening quote window: ${e.message}`, "error");
             // Still attempt to send to Google Sheets if window fails to open
             sendMappingToGoogleSheets(updateInfo.quoteNumber, updateInfo.fullOrderNumber, updateInfo.originalQuoteUrl, updateInfo.orderUrl);
        }
    }

    // Check for Pending Quote Update (to be run in the opened quote window)
    function checkForPendingQuoteUpdate() {
        try {
            const updateInfoStr = localStorage.getItem('pendingQuoteUpdate');
            if (!updateInfoStr) return null;
            const updateInfo = JSON.parse(updateInfoStr);
            if ((Date.now() - updateInfo.timestamp) > 300000) { // 5 min timeout
                debugLog("Update info is too old, clearing it", null, 'warn');
                localStorage.removeItem('pendingQuoteUpdate');
                return null;
            }
            // Check if the *current* page matches the expected original quote ID
            const currentQuoteId = window.location.pathname.split('/').pop();
            if (currentQuoteId !== updateInfo.originalQuoteId) {
                // debugLog(`Not the target quote page. Expected ${updateInfo.originalQuoteId}, Current ${currentQuoteId}`);
                return null; // Not the page we intend to update
            }
            return updateInfo;
        } catch (e) { debugLog(`Error checking for quote update: ${e.message}`, null, 'error'); return null; }
    }

    // Update Original Quote (runs in the opened quote window)
    function updateOriginalQuote() {
        const updateInfo = checkForPendingQuoteUpdate();
        if (!updateInfo) return; // Not the right page or no update info

        debugLog(`Found update info for this quote: ${JSON.stringify(updateInfo)}`);
        if (document.getElementById('quote-already-updated')) { debugLog("This quote has already been updated"); return; }

        const marker = document.createElement('div');
        marker.id = 'quote-already-updated'; marker.style.display = 'none';
        document.body.appendChild(marker);

        debugLog("Starting update process for original quote note...");
        showStatusNotification(`Updating quote with SO note...`, 'loading');

        // --- REINTRODUCE RETRY LOGIC ---
        let updateAttempt = 0;
        const maxUpdateAttempts = 8; // Increased attempts slightly
        const attemptInterval = 750; // Retry interval

        function tryUpdateNote() {
            updateAttempt++;
            debugLog(`Attempting to update original quote note (attempt ${updateAttempt}/${maxUpdateAttempts})`);

            const populateSuccess = populateConvertedQuoteNote(updateInfo.salesOrderNote, true); // PASS TRUE
            debugLog(`[tryUpdateNote] populateConvertedQuoteNote returned (tentative): ${populateSuccess}`);

            if (populateSuccess) {
                // Success: Note population initiated, rely on its internal verification/save trigger
                debugLog("populateConvertedQuoteNote succeeded (tentatively). Relying on its async verification/save.");
                localStorage.removeItem('pendingQuoteUpdate'); // Remove flag assuming process started

                // Send to Google Sheets
                setTimeout(() => {
                    sendMappingToGoogleSheets(updateInfo.quoteNumber, updateInfo.fullOrderNumber, updateInfo.originalQuoteUrl, updateInfo.orderUrl);
                }, 500);
                // Do NOT close the window here, let the save process handle it via closeWindowAfterSave

            } else if (updateAttempt < maxUpdateAttempts) {
                 // Failure, but more attempts left: Schedule retry
                 debugLog("Field likely not ready, scheduling retry...");
                 setTimeout(tryUpdateNote, attemptInterval);

            } else {
                 // Failure and max attempts reached: Show error, send mapping, close window
                 debugLog("[updateOriginalQuote Error] Failed to find/update original quote note after multiple attempts.", null, 'error');
                 showStatusNotification("Could not update quote automatically", 'error');
                 showError("Could not find/update the 'CONVERTED QUOTE NOTE' field after several attempts.");
                 // Still send mapping and close window after delay
                 sendMappingToGoogleSheets(updateInfo.quoteNumber, updateInfo.fullOrderNumber, updateInfo.originalQuoteUrl, updateInfo.orderUrl);
                 setTimeout(() => closeWindowAfterSave(true), 1500); // Pass true to close quote window
             }
        }

        // Initial attempt
        tryUpdateNote();
        // --- END RETRY LOGIC ---
    }

    // Process Copied Quote (runs in the new tab)
    function processCopiedQuote() { // <<< REMOVED async
        const conversionInfo = getStoredConversionInfo();
        if (!conversionInfo) return; // No info or original tab

        if (isOriginalQuotePage()) { // Double safety check
            debugLog("SAFETY CHECK: This is an original quote page - canceling auto-conversion!", null, 'error');
            return;
        }
        // Ensure we are on a page with a convert button (likely a copied quote)
        const hasConvertButton = !!document.querySelector('#convertSalesOrder');
        if (!hasConvertButton) { debugLog("Not a copied quote page (no convert button)", null, 'warn'); return; }

        if (document.getElementById('automation-already-started')) { debugLog("Automation already started on this page"); return; }
        const marker = document.createElement('div');
        marker.id = 'automation-already-started'; marker.style.display = 'none';
        document.body.appendChild(marker);

        debugLog("Starting automated conversion process on the copied quote");
        debugLog("Attempting to find and populate note field...");

        // *** START RETRY LOGIC ***
        let populateAttempt = 0;
        const maxPopulateAttempts = 5;
        const populateIntervalTime = 1000;

        const attemptToPopulateInterval = setInterval(() => {
            populateAttempt++;
            debugLog(`Attempting to populate copied quote note (attempt ${populateAttempt}/${maxPopulateAttempts})`);

            const populateSuccess = populateConvertedQuoteNote(conversionInfo.formattedNote, false); // <<< PASS FALSE (or rely on default)

            if (populateSuccess) {
                clearInterval(attemptToPopulateInterval);
                debugLog("Successfully populated note field. Proceeding to convert.");
                // Wait a bit for UI to settle after note population, then click convert
                setTimeout(() => {
                    if (!isOriginalQuotePage()) { // Final safety check
                        clickConvertButton(); // *** Call CONVERT here ***
                    } else {
                        debugLog("FINAL SAFETY CHECK PREVENTED CONVERSION on original quote page!", null, 'error');
                        showError("⚠️ Prevented conversion attempt on original quote page");
                    }
                }, 500); // Delay before clicking convert
            } else if (populateAttempt >= maxPopulateAttempts) {
                 clearInterval(attemptToPopulateInterval);
                 debugLog("Failed to populate note field for copied quote after multiple attempts. Attempting conversion anyway.", null, 'warn');
                 // ** ADD ERROR MESSAGE HERE **
                 showError("Could not find/update the 'CONVERTED QUOTE NOTE' field after several attempts. Trying to convert anyway.");
                 setTimeout(() => {
                     if (!isOriginalQuotePage()) { // Final safety check
                         clickConvertButton();
                     } else {
                         debugLog("FINAL SAFETY CHECK PREVENTED CONVERSION on original quote page (after note failure)!", null, 'error');
                         showError("⚠️ Prevented conversion attempt on original quote page");
                     }
                 }, 1000);
            }
            // If not successful and not max attempts, the interval continues...
        }, populateIntervalTime);
         // *** END RETRY LOGIC ***
    }

    // Check if this is a Quote Page (more robust)
    function isThisAQuotePage() {
        if (window.location.pathname.includes('/quotes/')) return true; // URL check first
        if (document.querySelector('#convertSalesOrder')) return true; // Convert button exists
        const entityHeaderInput = document.getElementById('entity-header-input');
        if (entityHeaderInput && (entityHeaderInput.value?.includes('SQ-') || entityHeaderInput.title?.includes('SQ-'))) return true; // Header input SQ-
         const title = document.title.toLowerCase();
         if (title.includes('quote') && !title.includes('order')) return true; // Page title check
        const prominentElements = document.querySelectorAll('h1, h2, h3, .header, .title');
        for (const el of prominentElements) {
            if (el.textContent.includes('Quote') && !el.textContent.includes('Order')) return true; // Prominent text
        }
        return false;
    }

    // Add Copy & Convert Button to Header
    function addHeaderButton() {
        if (document.getElementById('copyAndConvertButton')) return; // Already exists
        if (!isThisAQuotePage()) return; // Only on quote pages

        debugLog("Attempting to add Copy & Convert button");
        const headerSelectors = [ '.sc-371bb816-2', // Original selector
            '[class*="-header"] [class*="buttons"]', 'div[class*="header"] div[class*="actions"]',
            'div[class*="toolbar"]' ];
        let headerButtonsContainer = null;
        for (const selector of headerSelectors) {
            const container = document.querySelector(selector);
            if (container) { headerButtonsContainer = container; debugLog(`Found header container using selector: ${selector}`); break; }
        }
         // Fallback: Find container near Copy button
         if (!headerButtonsContainer) {
            const copyBtn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.includes('Copy') && !b.textContent.includes('Copy & Convert'));
            if (copyBtn) {
                 let potentialContainer = copyBtn.parentElement;
                 for (let i = 0; i < 3 && potentialContainer; i++) {
                     if (potentialContainer.children.length > 1 && potentialContainer.tagName === 'DIV') { // Look for a div container
                         headerButtonsContainer = potentialContainer;
                         debugLog(`Found header container near Copy button (level ${i+1})`);
                         break;
                     }
                     potentialContainer = potentialContainer.parentElement;
                 }
            }
         }

        if (!headerButtonsContainer) { debugLog("Header buttons container not found - will retry later", null, 'warn'); return; }

        // --- Create Button Elements (Copying styles observed in original script) ---
        const outerDiv = document.createElement('div');
        // Use specific classes if known, otherwise generic styles
        outerDiv.style.display = 'inline-block'; // Adjust as needed
        outerDiv.style.marginRight = '8px'; // Add spacing

        const button = document.createElement('button');
        button.id = 'copyAndConvertButton';
        // Apply classes observed on other buttons for styling consistency
        button.className = 'sc-dd788c24-0 jepwNu sc-3e671f2a-0 hsFHlp'; // Example classes - adjust if needed
        button.style.padding = '8px 12px';
        // **REMOVE BORDER**
        button.style.border = 'none'; // Set border to none
        button.style.borderRadius = '4px';
        button.style.cursor = 'pointer';
        button.style.display = 'flex';
        button.style.alignItems = 'center';

        const textSpan = document.createElement('span');
        textSpan.textContent = 'Copy & Convert';
        textSpan.style.marginRight = '5px'; // Space between text and icon
         // Apply text styling classes if known
         textSpan.className = 'sc-8b7d6573-19 sc-dd788c24-2 kCNWRT bNlzzM'; // Example

        const icon = document.createElement('h4'); // Or span, div etc.
        icon.textContent = '⟳'; // Refresh-like icon
        icon.style.fontSize = '16px';
        // Apply icon styling classes if known
        icon.className = 'sc-8b7d6573-4 sc-4a378d6f-0 kTHGAT kkZzUh'; // Example

        button.appendChild(textSpan);
        button.appendChild(icon);
        outerDiv.appendChild(button);
        // --- End Button Element Creation ---


        button.addEventListener('click', handleCopyConvertClick, true); // Use capture phase

        // Find insertion point (after Copy button ideally)
        let insertBeforeElement = null;
        const copyButtonContainer = Array.from(headerButtonsContainer.children).find(el => el.textContent.includes('Copy') && !el.textContent.includes('Copy & Convert'));
        if (copyButtonContainer) {
            insertBeforeElement = copyButtonContainer.nextSibling;
            debugLog("Found Copy button container to position after");
        } else {
            // Fallback: insert before Convert button or append
            const convertButtonContainer = Array.from(headerButtonsContainer.children).find(el => el.textContent.includes('Convert'));
            if (convertButtonContainer) {
                insertBeforeElement = convertButtonContainer;
                debugLog("Found Convert button container to position before");
            }
        }

        if (insertBeforeElement) {
            headerButtonsContainer.insertBefore(outerDiv, insertBeforeElement);
        } else {
            headerButtonsContainer.appendChild(outerDiv); // Append if no better position found
        }
        debugLog("Successfully added Copy & Convert button to header");
        localStorage.setItem('buttonAddedToQuote_' + window.location.pathname, 'true'); // Mark as added for this page
    }

    // Schedule Button Addition Retries
    function scheduleButtonAdditionRetries() {
        if (localStorage.getItem('buttonAddedToQuote_' + window.location.pathname) === 'true' && document.getElementById('copyAndConvertButton')) {
            return; // Already added successfully
        }
        const retryTimes = [1000, 2000, 3000, 5000, 8000, 13000, 21000];
        for (const time of retryTimes) {
            setTimeout(() => {
                if (!document.getElementById('copyAndConvertButton') && isThisAQuotePage()) { // Check page type again
                    debugLog(`Retry adding button at ${time}ms`);
                    addHeaderButton();
                }
            }, time);
        }
    }


    // ===================================================
    // PART 2: Master Script Functionality (PO, SO, Product)
    // ===================================================

    // ---- INTERNAL NOTE ENHANCEMENT START ----
    function enhanceInternalNoteField() {
        const inputField = document.querySelector('#custom2-input'); // Find original input
        if (!inputField) return; // Not found

        const container = inputField.closest('.sc-2c12f804-0.eALfvH.input-host');
        if (!container) {
            log("INTERNAL NOTE: Could not find container for input.", null, 'warn');
            return;
        }

        const newlinePlaceholder = '¶'; // Pilcrow sign U+00B6
        const newlineRegex = /\n/g;
        const placeholderRegex = new RegExp(newlinePlaceholder.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g'); // Ensure placeholder is regex-safe

        // Check if already enhanced
        if (container.querySelector('#custom2-textarea-overlay')) {
            const textareaOverlay = container.querySelector('#custom2-textarea-overlay');
            // ** CHANGE: Convert placeholder back to newline for update check **
            const inputFieldValueWithNewlines = inputField.value.replace(placeholderRegex, '\n');
            if (textareaOverlay.value !== inputFieldValueWithNewlines) {
                log("INTERNAL NOTE: Updating overlay from hidden input.");
                textareaOverlay.value = inputFieldValueWithNewlines;
            }
            return;
        }

        log("INTERNAL NOTE: Enhancing input field with overlay and newline substitution.");

        container.style.position = 'relative';

        const textarea = document.createElement('textarea');
        textarea.id = 'custom2-textarea-overlay';
        textarea.placeholder = inputField.placeholder;
        // ** CHANGE: Convert placeholder back to newline for initial display **
        textarea.value = inputField.value.replace(placeholderRegex, '\n');
        textarea.className = 'internal-note-overlay';
        textarea.setAttribute('rows', '5');

        // Add listener to sync textarea value TO hidden input
        textarea.addEventListener('input', function() {
            // ** CHANGE: Convert newline to placeholder before setting hidden input **
            const valueWithPlaceholders = textarea.value.replace(newlineRegex, newlinePlaceholder);

            if (inputField.value !== valueWithPlaceholders) {
                inputField.value = valueWithPlaceholders;

                // Try simpler event dispatch first
                inputField.dispatchEvent(new Event('input', { bubbles: true }));
                inputField.dispatchEvent(new Event('change', { bubbles: true }));

                // --- REINSTATED: Attempt React handler call ---
                try {
                    const reactKey = Object.keys(inputField).find(key => key.startsWith('__reactProps') || key.startsWith('__reactFiber'));
                    if (reactKey) {
                        const reactProps = inputField[reactKey];
                        const handler = reactProps?.onChange || reactProps?.memoizedProps?.onChange || reactProps?.props?.onChange;
                        if (handler && typeof handler === 'function') {
                             const eventArg = { target: inputField, currentTarget: inputField, type: 'change', bubbles: true, cancelable: true, isTrusted: false, nativeEvent: { target: inputField }, preventDefault: () => {}, stopPropagation: () => {} };
                            handler(eventArg);
                            debugLog("Attempted to call React onChange handler directly in populateConvertedQuoteNote");
                        }
                    }
                } catch (e) {
                    debugLog("Error calling React handler in populateConvertedQuoteNote: " + e.message, null, 'warn');
                }
                // --- END REINSTATED ---
            }
            // ** REMOVED: Cursor restoration - focus on core sync **
        });

        // Double-check: Only append if not already there
        if (!container.querySelector('#custom2-textarea-overlay')) {
            container.appendChild(textarea);
            container.setAttribute('data-note-enhanced', 'true');
             log("INTERNAL NOTE: Overlay textarea added.");
        } else {
            log("INTERNAL NOTE: Overlay already present, skipped appending.");
        }


        // Add CSS for hiding the original input and styling/positioning the overlay
        const styleCheck = document.getElementById('internal-note-enhancement-style');
        if (!styleCheck) {
             const style = document.createElement('style');
             style.id = 'internal-note-enhancement-style';
             // ** CHANGE: Use absolute positioning for overlay, adjust container **
             style.textContent = `
                /* Ensure container is relative for absolute child positioning */\n\
                .sc-2c12f804-0.eALfvH.input-host {\n\
                    position: relative !important; \n\
                    width: 100% !important; /* Ensure container takes full width */\n\
                }\n\
\n\
                /* Hide the original input effectively */\n\
                #custom2-input {\n\
                    position: absolute;\n\
                    top: 0;\n\
                    left: 0;\n\
                    width: 100%;\n\
                    height: 100%;\n\
                    opacity: 0;\n\
                    pointer-events: none; \n\
                    z-index: -1; \n\
                    border: none; \n\
                    padding: 0; \n\
                    margin: 0; \n\
                    /* Crucially, allow the container to dictate size */\n\
                    min-height: unset !important;\n\
                }\n\
\n\
                /* Style the overlay textarea */\n\
                .internal-note-overlay {\n\
                    position: absolute; /* Position absolutely within the container */\n\
                    top: 0;\n\
                    left: 0;\n\
                    z-index: 1; /* Ensure it\'s above the hidden input */\n\
                    width: 100%;\n\
                    height: 100%; /* Fill the container */\n\
                    min-height: 90px; /* Set desired minimum height FOR THE TEXTAREA */\n\
                    box-sizing: border-box;\n\
                    font-family: inherit;\n\
                    font-size: inherit;\n\
                    line-height: inherit;\n\
                    padding: 6px 8px; /* Match typical input padding */\n\
                    border: 1px solid #ccc; /* Basic border */\n\
                    border-radius: 4px; /* Rounded corners */\n\
                    white-space: pre-wrap !important;\n\
                    resize: vertical !important;\n\
                    overflow: auto; /* Show scrollbar if needed */\n\
                    background-color: inherit; \n\
                    color: inherit;\n\
                }\n\
\n\
                /* Ensure container allows height adjustment AND sets the base size */\n\
                .sc-2c12f804-0.eALfvH.input-host[data-note-enhanced="true"] {\n\
                     height: auto !important; \n\
                     min-height: 90px !important; /* Set min-height ON THE CONTAINER */\n\
                     padding: 0 !important; /* Remove container padding if it conflicts */\n\
                     overflow: visible !important; /* Allow content to overflow if needed */\n\
                }\n\
\n\
                /* Force height auto and overflow visible on higher parents */\n\
                 .sc-5bf8f203-4.fcolsK, /* Grandparent */\n\
                 .sc-5bf8f203-2.DYtF /* Great-Grandparent */ {\n\
                     height: auto !important; \n\
                     min-height: unset !important; \n\
                     overflow: visible !important; \n\
                 }\n\
             `;
             document.head.appendChild(style);
        }
    }
    // ---- INTERNAL NOTE ENHANCEMENT END ----

    // --- ADDED Save Handler Function ---
    // --- Refined Save Handler Function ---
    async function handleSaveClick(event) {
        const saveButton = event.currentTarget;
        log("Save button clicked, handling potential Look Up Number update (Bubbling Phase).");

        // Only run on Sales Order pages
        if (!window.location.href.includes('/sales-orders/')) {
            log("Not on Sales Order page, skipping Look Up Number update.");
            return;
        }

        const remarksField = document.querySelector('textarea[name="remarks"]');
        const lookUpNumberField = document.getElementById('so_PO');

        if (!remarksField || !lookUpNumberField) {
            log("Could not find Remarks or Look Up Number field for save handling.", null, 'warn');
            // Don't interfere with save if fields are missing
            return;
        }

        const remarksText = remarksField.value;
        const jobNameMatch = remarksText.match(/JOB NAME\/REF:\s*(.*?)(?:\s*PICKED UP BY:|\n|$)/i);
        const jobNameValue = jobNameMatch && jobNameMatch[1] ? jobNameMatch[1].trim() : '';
        const currentLookUpValue = lookUpNumberField.value;

        if (jobNameValue !== currentLookUpValue) {
            log(`Look Up Number mismatch. Current: "${currentLookUpValue}", Expected from Remarks: "${jobNameValue}". Intercepting save to update.`);

            // --- Intercept the event ONLY if update is needed ---
            event.preventDefault();
            event.stopPropagation();

            // Disable button and show status
            saveButton.disabled = true;
            const originalButtonText = saveButton.textContent;
            saveButton.textContent = 'Updating...';
            log("Button disabled, text set to Updating..."); // Add log

            try {
                log("Calling useNativeCommands..."); // Add log
                lookUpNumberField.focus(); // <<< ADD FOCUS before update
                useNativeCommands(lookUpNumberField, jobNameValue, false);
                // lookUpNumberField.blur(); // <<< REMOVED: useNativeCommands now handles blur
                log("useNativeCommands finished."); // Add log

                log("Waiting for longer delay..."); // Add log
                await new Promise(resolve => setTimeout(resolve, 1000)); // <<< INCREASED DELAY to 1000ms
                log("Delay finished."); // Add log

                log("Look Up Number updated, preparing to re-trigger save.");

            } catch (error) {
                log("Error occurred while updating Look Up Number before save.", error, 'error');
                showError("Error updating Look Up Number. Please check and save manually.");
                // Restore button state on error, but DO NOT proceed with save automatically
                saveButton.disabled = false;
                saveButton.textContent = originalButtonText;
                log("Button restored after error."); // Add log
                return; // Stop the process if update failed
            }
            // No finally block needed here as we only restore state on error or success before re-click

            // If update was successful (no error thrown/caught), restore button and click save
            log("Update successful, restoring button and re-triggering save click.");
            saveButton.disabled = false;
            saveButton.textContent = originalButtonText;
            saveButton.click(); // Trigger the actual save action
            log("Re-triggered save click."); // Add log

        } else {
            log("Look Up Number matches remarks value or no JOB NAME/REF found. Allowing default save action.");
            // Do nothing, let the original click event proceed.
        }
    }
    // --- END Refined Save Handler Function ---

    // Insert Blanket PO
    function tryInsertPO() {
        if (!window.location.href.includes('/sales-orders/')) return;
        let customerField = document.querySelector("#so_customer input[type='text']");
        let remarksField = document.querySelector('textarea[name="remarks"]'); // Changed from poField
        if (!customerField || !remarksField) return;

        // Prevent duplicate formatting if already done
        if (remarksField.value.includes("JOB NAME/REF:")) {
            log("Remarks already formatted. Skipping.");
            // Ensure correct status is shown if BPO was previously applied
            if (remarksField.value.startsWith("BPO#:")) {
                insertBPOStatusIndicator("applied");
            } else if (remarksField.value.startsWith("PO#:")) {
                insertBPOStatusIndicator("no-bpo"); // Or perhaps a different status?
            }
            return;
        }
        // --- ADDED: Stronger duplicate check ---
        if (remarksField.value.includes("JOB NAME/REF:") && remarksField.value.includes("PICKED UP BY:")) {
            log("Remarks already fully formatted (JOB NAME & PICKED UP BY). Skipping.");
            return;
        }
        // --- ADDED: Check formatting flag ---
        if (remarksField.dataset.formatting === 'true') {
            log("Remarks formatting already in progress. Skipping.");
            return;
        }

        let customerName = customerField.value.trim();
        if (!customerName) return;

        // Avoid duplicate calls if customer name hasn't changed
        if (customerName === localStorage.getItem('lastPOCustomer')) return;
        localStorage.setItem('lastPOCustomer', customerName);

        // Show loading state
        insertBPOStatusIndicator("loading");

        // --- ADDED: Set formatting flag ---
        remarksField.dataset.formatting = 'true';

        // --- ADDED: Delay before fetching/formatting ---
        setTimeout(() => {
             // Re-check elements in case they disappeared during the timeout
             customerField = document.querySelector("#so_customer input[type='text']");
             remarksField = document.querySelector('textarea[name="remarks"]');
             if (!customerField || !remarksField) {
                 log("Elements disappeared during tryInsertPO delay. Aborting.");
                 return;
             }
             // Re-verify customer name hasn't changed during timeout
             if (customerField.value.trim() !== customerName) {
                 log("Customer changed during tryInsertPO delay. Aborting.");
                 // Reset lastPOCustomer so the change triggers a new run
                 localStorage.removeItem('lastPOCustomer');
                 return;
             }


             GM_xmlhttpRequest({
                 method: "GET",
                 url: blanketPOSheetUrl, // Use correct URL
                 onload: function(response) {
                     let blanketPOValue = null;
                     if (response.status === 200) {
                         try {
                             let data = parseCSV(response.responseText);
                             let row = data.find(r => (r["Customer Name"] || r["Customer_Name"] || r["Customer"] || "").trim().toLowerCase() === customerName.toLowerCase());
                             if (row && row["Blanket PO"]) {
                                 blanketPOValue = row["Blanket PO"];
                             }
                         } catch (e) {
                             console.error("Error parsing Blanket PO CSV:", e);
                             // Proceed without blanketPO
                         }
                     } else {
                         console.error("Failed to fetch Blanket PO CSV. Status:", response.status);
                         // Proceed without blanketPO
                     }
                     // Call helper to format and insert
                     formatAndInsertRemarks(blanketPOValue, remarksField);
                 },
                 onerror: function(err) {
                     console.error("Error fetching Blanket PO CSV:", err);
                     insertBPOStatusIndicator("error"); // Show error status
                     // Still try to format with page PO# even on sheet error
                     formatAndInsertRemarks(null, remarksField);
                     // --- ADDED: Clear formatting flag on error ---
                     remarksField.dataset.formatting = 'false'; // Clear flag if format attempt aborted by error
                 }
             });
        }, 1000); // 1 second delay
    }

    // Helper function to format and insert remarks
    function formatAndInsertRemarks(blanketPOValue, remarksField) {
        // 1. Find Look Up Number field value
        let lookUpNumber = '';
        try {
            const lookUpLabelElement = Array.from(document.querySelectorAll('dt')).find(dt => dt.textContent.trim() === 'Look Up Number');
            if (lookUpLabelElement) {
                const ddElement = lookUpLabelElement.nextElementSibling;
                const inputElement = ddElement ? ddElement.querySelector('input') : null;
                if (inputElement) {
                    lookUpNumber = inputElement.value.trim();
                    log(`Found Look Up Number: ${lookUpNumber}`);
                } else {
                    log("Could not find input field for Look Up Number.", null, 'warn');
                }
            } else {
                log("Could not find 'Look Up Number' label.", null, 'warn');
            }
        } catch (e) {
            log(`Error getting Look Up Number: ${e.message}`, null, 'error');
        }

        // 2. Determine Prefix and update indicator
        let poPrefix = '';
        let indicatorStatus = '';
        if (blanketPOValue) {
            poPrefix = `BPO#:${blanketPOValue}`;
            indicatorStatus = 'applied'; // BPO was found and applied
        } else {
            poPrefix = `PO#:${lookUpNumber || ''}`; // Use Look Up number or empty string
            indicatorStatus = 'no-bpo'; // No BPO found in sheet
        }
        insertBPOStatusIndicator(indicatorStatus); // Update indicator

        // 3. Extract existing user-entered values for JOB NAME/REF and PICKED UP BY
        const originalValue = remarksField.value;
        const jobMatch = originalValue.match(/JOB NAME\/REF:\s*(.*)/i);
        const existingJobValue = jobMatch && jobMatch[1] ? jobMatch[1].trim() : '';
        const pickMatch = originalValue.match(/PICKED UP BY:\s*(.*)/i);
        const existingPickValue = pickMatch && pickMatch[1] ? pickMatch[1].trim() : '';

        // 4. Clean other content from remarks
        let otherContent = originalValue;
        otherContent = otherContent.replace(/\(BLANKET PO#:[^)]+\)\s*\n?/g, '');
        otherContent = otherContent.replace(/BPO#:[^\s]+\s*\n?/g, '');
        otherContent = otherContent.replace(/PO#:[^\s]*\n?/g, '');
        otherContent = otherContent.replace(/JOB NAME\/REF:.*\n?/gi, '');
        otherContent = otherContent.replace(/PICKED UP BY:.*\n?/gi, '');
        otherContent = otherContent.trim();

        // 5. Construct new remarks block with preserved user values
        const jobLine = existingJobValue ? `JOB NAME/REF:${existingJobValue}` : 'JOB NAME/REF:';
        const pickLine = existingPickValue ? `PICKED UP BY:${existingPickValue}` : 'PICKED UP BY:';
        const newRemarksBlock = `${poPrefix}\n${jobLine}\n${pickLine}`;
        const textToInsert = otherContent ? `${newRemarksBlock}\n\n${otherContent}` : newRemarksBlock;

        // 6. Only update if there is a change
        if (textToInsert !== originalValue) {
            useNativeCommands(remarksField, textToInsert);
            log(`Updated remarks field. Prefix used: ${poPrefix}`);
        }
        // --- ADDED: Clear formatting flag ---
        remarksField.dataset.formatting = 'false'; // Ensure flag is cleared after update attempt
    }

    // Function to insert BPO status indicator
    function insertBPOStatusIndicator(status) {
        const existingIndicator = document.getElementById('bpo-status-indicator');
        if (existingIndicator) existingIndicator.remove();

        let labelText = "No BPO"; let bgColor = "grey";
        if (status === "applied") { labelText = "BPO Applied"; bgColor = "green"; }
        else if (status === "loading") { labelText = "Checking For BPO"; bgColor = "#2196F3"; }
        else if (status === "error") { labelText = "Error"; bgColor = "red"; }

        const indicator = document.createElement('span');
        indicator.id = 'bpo-status-indicator';
        indicator.style.cssText = `
            background-color: ${bgColor};
            color: white;
            padding: 2px 6px;
            border-radius: 4px;
            font-weight: bold;
            margin-left: 10px;
            font-size: 12px;
            display: inline-block;
            vertical-align: middle;
        `;
        indicator.textContent = labelText;

        // Find the Remarks title
        const remarksTitle = document.querySelector('h3.sc-8b7d6573-2.sc-2a74c12b-1.bTGyIl.ETPKj');
        if (remarksTitle) {
            remarksTitle.appendChild(indicator);
        }
    }

    // Insert the Card-on-file row
    function insertCardOnFileRow(status) {
        const existingRow = document.getElementById('card-on-file-row');
        if (existingRow) existingRow.remove();

        let labelText = "No card"; let bgColor = "grey";
        if (status === "on-file") { labelText = "On file"; bgColor = "green"; }
        else if (status === "expired") { labelText = "Expired"; bgColor = "red"; }
        else if (status === "loading") { labelText = "Checking Sheet"; bgColor = "#2196F3"; }

        const container = document.createElement('div');
        container.id = 'card-on-file-row';
        // Use known class names if possible, otherwise style directly
        container.className = 'sc-a3cffb35-0 ifEDhR'; // Example class
        container.style.marginBottom = '10px'; // Add some spacing

        container.innerHTML = `
            <dl class="sc-a3cffb35-1 Jiqon" style="display: flex; align-items: center;">
              <dt class="sc-a3cffb35-3 cDizGq" style="flex-basis: 120px; margin-right: 10px;">
                <p class="sc-a3cffb35-4 botIdl" style="margin: 0; font-weight: bold;">Card on file</p>
              </dt>
              <dd class="sc-a3cffb35-6 kXBFJe">
                <span id="card-on-file-indicator" style="background-color: ${bgColor}; color: white; padding: 2px 6px; border-radius: 4px; font-weight: bold;">
                  ${labelText}
                </span>
              </dd>
            </dl>
          `;
         // Find insertion point more robustly
         const customerDiv = document.querySelector('#so_customer')?.closest('div.sc-a3cffb35-0'); // Find customer row container
         const paymentTermsDiv = Array.from(document.querySelectorAll('p.sc-a3cffb35-4')).find(p => p.textContent.includes('Payment terms'))?.closest('div.sc-a3cffb35-0');

         let insertBeforeElement = paymentTermsDiv; // Default to inserting before payment terms

         if (!insertBeforeElement) {
             // Fallback: try inserting after customer div if payment terms not found
             if(customerDiv && customerDiv.nextElementSibling) {
                insertBeforeElement = customerDiv.nextElementSibling;
             }
         }

         if (insertBeforeElement && insertBeforeElement.parentNode) {
             insertBeforeElement.parentNode.insertBefore(container, insertBeforeElement);
             log("Inserted Card on File row");
         } else if (customerDiv && customerDiv.parentNode) {
             // Fallback: Append after customer div if nothing else works
             customerDiv.parentNode.appendChild(container);
              log("Inserted Card on File row (fallback append)");
         }
          else {
             console.error("Could not find suitable location to insert card on file row.");
         }
    }

    // Card On File Indicator (logic)
    function addCardOnFileIndicator() {
        if (!window.location.href.includes('/sales-orders/')) return; // Only on SO pages
        let customerField = document.querySelector("#so_customer input[type='text']");
        if (!customerField) return;
        let customerName = customerField.value.trim();
        if (!customerName) return;

        if (customerName === lastCardCustomerName) return; // Avoid redundant checks
        if (inFlightCustomerName === customerName) { log("Card on file request already in flight. Skipping..."); return; }

        // Show loading state immediately
        insertCardOnFileRow("loading");
        inFlightCustomerName = customerName; // Set flag

        GM_xmlhttpRequest({
            method: "GET",
            url: cardSheetUrl,
            onload: function(response) {
                inFlightCustomerName = ""; // Clear flag
                lastCardCustomerName = customerName; // Update last checked name *after* response
                if (response.status !== 200) { console.error("Failed to fetch Card On File CSV. Status:", response.status); return; }
                let data = parseCSV(response.responseText);
                if (!data || data.length === 0) { insertCardOnFileRow("no-card"); return; }

                let row = data.find(r => (r["Customer Name"] || r["Customer_Name"] || r["Customer"] || "").trim().toLowerCase() === customerName.toLowerCase());
                if (!row) { insertCardOnFileRow("no-card"); return; }

                let expDateStr = row["Experation Date"] || row["Expiration Date"]; // Handle typo
                let isExpired = false;
                if (expDateStr) {
                    try {
                         let expDate = parseExpirationDate(expDateStr);
                         let now = new Date();
                         now.setHours(0, 0, 0, 0); // Compare dates only
                         if (isNaN(expDate.getTime())) { // Check if Invalid Date
                             console.error("Invalid date format for expiration date:", expDateStr);
                             isExpired = true; // Treat invalid date as expired
                         } else if (expDate < now) {
                             isExpired = true;
                         }
                    } catch (e) {
                         console.error("Error parsing expiration date:", expDateStr, e);
                         isExpired = true; // Treat parsing error as expired
                    }
                } else {
                    // If no expiration date is provided, assume it's okay unless explicitly marked otherwise
                     isExpired = false;
                     log(`No expiration date found for ${customerName}, assuming card is valid.`);
                }

                insertCardOnFileRow(isExpired ? "expired" : "on-file");
            },
            onerror: function(err) {
                inFlightCustomerName = ""; // Clear flag
                lastCardCustomerName = customerName; // Still update last checked name on error to avoid loops
                console.error("Error fetching Card On File CSV:", err);
                 insertCardOnFileRow("error"); // Indicate an error occurred
            }
        });
    }

    // Rename Elements on Sales Orders
    function renamePOElements() {
        if (!window.location.href.includes('/sales-orders/')) return;
        let changesMade = false;
        document.querySelectorAll('p, label, dt').forEach(elem => { // Check more element types
            const text = elem.textContent.trim();
            if (text === "PO #") { elem.textContent = "Look Up Number"; changesMade = true; }
            if (text === "Include shipping") { elem.textContent = "Fulfill/Pick"; changesMade = true; }
            if (text === "Assigned To") { elem.textContent = "Converted/Made by"; changesMade = true; }
        });
        // if (changesMade) log('Renamed elements on Sales Order page');
    }

    // Remove Fulfill Button
    function removeFulfillButton() {
         if (!window.location.href.includes('/sales-orders/')) return;
        let fulfillButton = document.getElementById("so-btn-Fulfill");
        if (fulfillButton && fulfillButton.offsetParent !== null) { // Check if visible
            fulfillButton.remove();
            log("Fulfill button removed.");
        }
    }

    // Credit Hold Check
    function checkCreditHold() {
         if (!window.location.href.includes('/sales-orders/')) return;
        let creditHoldElement = document.querySelector('input[title*="CREDIT HOLD"]'); // Use contains selector for title
        if (!creditHoldElement) { creditHoldPopupDismissed = false; return; } // Reset dismiss flag if no hold

        // Remove Print/Email buttons if on credit hold
        let printButton = document.getElementById("sales-order-print-dropdown");
        if (printButton) { printButton.remove(); log("Print button removed (Credit Hold)."); }
        let emailButton = document.getElementById("so-email-dropdown");
        if (emailButton) { emailButton.remove(); log("Email button removed (Credit Hold)."); }

        // Show popup if not dismissed
        if (!creditHoldPopupDismissed && !document.getElementById("credit-hold-popup")) {
            let overlay = document.createElement('div');
            overlay.id = "credit-hold-popup";
            overlay.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background-color:rgba(0,0,0,0.6); display:flex; align-items:center; justify-content:center; z-index:10000;";
            let popup = document.createElement('div');
            popup.style.cssText = "background-color:white; padding:30px; border-radius:8px; text-align:center; box-shadow:0 0 15px rgba(0,0,0,0.3); max-width:400px;";
            popup.innerHTML = `
                <h1 style="font-size:24px; font-weight:bold; margin:0 0 15px 0; color: red;">CREDIT HOLD</h1>
                <p style="margin:0 0 20px 0;">Overdue balance must be paid before fulfilling.</p>
                <button id="credit-hold-close" style="padding: 8px 15px; cursor: pointer;">Close</button>
            `;
            overlay.appendChild(popup);
            document.body.appendChild(overlay);
            document.getElementById('credit-hold-close').addEventListener('click', function() {
                creditHoldPopupDismissed = true;
                overlay.remove();
            });
            log("Displayed Credit Hold popup.");
        }
    }

    // Handle Quote Page specific UI changes
    function handleQuotePage() {
        if (!isThisAQuotePage()) return; // Ensure it's a quote page
        // ** REMOVE quoteConversionHappened check - let the observer handle re-applying if needed after conversion **
        // if (quoteConversionHappened) return; // Don't run if conversion started/happened

        let changesMade = false;
        // Rename "Sales rep" to "Quote by"
        document.querySelectorAll('p.sc-a3cffb35-4.botIdl, dt').forEach(elem => {
            if (elem.textContent.trim() === "Sales rep") {
                elem.textContent = "Quote by";
                changesMade = true;
            }
        });

        // Remove "Assigned To" section
        const container = document.querySelector("div.sc-3e2e0cb2-0.gLiEEA"); // Main container for sales rep/assigned to
        if (container) {
             // Find the "Assigned To" section more reliably
             const assignedToLabel = Array.from(container.querySelectorAll('p, dt')).find(el => el.textContent.trim() === 'Assigned To' || el.textContent.trim() === 'Converted/Made by');
             if (assignedToLabel) {
                 const sectionToRemove = assignedToLabel.closest('div.sc-3e2e0cb2-1, div.form-section'); // Find parent section
                 if (sectionToRemove) {
                     sectionToRemove.remove();
                      changesMade = true;
                 }
             }
        }
        // if (changesMade) log("Applied Quote-specific UI changes (Renamed Sales Rep, Removed Assigned To)");
    }


    // Get Logged-In User Name
    function getLoggedInUserName(callback) {
        // Add CSS to hide the menu but keep it functional
        const style = document.createElement('style');
        style.textContent = `
            #portal .sc-627ee65b-1 {
                opacity: 0 !important;
                pointer-events: none !important;
            }
        `;
        document.head.appendChild(style);

        // Try direct selector first
        let nameElem = document.querySelector("#user-menu-container p.sc-8b7d6573-10");
        if (nameElem && nameElem.textContent.trim()) {
            let name = nameElem.textContent.trim();
            log("Logged in user (from already open menu): " + name);
            callback(name);
            return;
        }

        // If not found, try clicking the menu button
        let menuButton = document.getElementById("user-menu-button");
        if (menuButton) {
            log("Attempting to open user menu to get name...");
            menuButton.click();
            setTimeout(() => {
                nameElem = document.querySelector("#user-menu-container p.sc-8b7d6573-10");
                if (nameElem && nameElem.textContent.trim()) {
                    let name = nameElem.textContent.trim();
                    log("Logged in user (after opening menu): " + name);
                    // Close the menu by clicking the button again
                    if (menuButton) menuButton.click();
                    callback(name);
                } else {
                    log("Still could not find logged in user name after opening menu.", null, 'warn');
                    // Close menu if open
                    if (menuButton) menuButton.click();
                    callback(null);
                }
            }, 500);
        } else {
            log("User menu button not found.", null, 'warn');
            callback(null);
        }
    }

    // Auto-Assign if Sales Rep (Only on Sales Orders)
    function autoAssignIfSalesRep() {
        // --- ADDED FOCUS GUARD ---
        if (document.activeElement && (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA')) {
            log("Skipping auto-assign: An input field has focus.");
            return;
        }
        // --- END FOCUS GUARD ---


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

        // log("Running autoAssignIfSalesRep check..."); // Add a log to see when it runs

        // Check that the "Assigned To" section exists.
        let assignedSection = document.querySelector("div.sc-3e2e0cb2-1.gtsOkx");
        if (!assignedSection) {
            // log("Assigned To section not found yet."); // Expected sometimes, don't log unless debugging
            return; // Exit if section not found, rely on more frequent calls
        }

        // More thorough check for existing assignment - PRIORITIZE ICON CHECK
        let isAlreadyAssigned = false;
        const iconSelector = 'div.sc-3e2e0cb2-1.gtsOkx div.sc-130ca08d-2.sc-5768b3d2-2.kauNDR.eRHrWR h4[color="#58698d"]';
        let iconElem = document.querySelector(iconSelector);

        if (iconElem) {
            const iconText = iconElem.textContent.trim();
            if (iconText !== "Ŝ") {
                // Icon is present and NOT the default "unassigned" icon
                isAlreadyAssigned = true;
                log(`Assignment detected via icon text: "${iconText}"`);
            } else {
                // Icon is present and IS the default "unassigned" icon
                isAlreadyAssigned = false;
                // log("Assignment icon shows default state (Ŝ)."); // Reduce noise
            }
        } else {
            // Icon element itself wasn't found. This might happen if the structure changes
            // OR if someone *is* assigned and the icon element is replaced entirely.
            // Let's fallback check for the assignment text div just in case.
            log("Assignment icon element not found. Falling back to text check.");
            let assignmentText = assignedSection.querySelector('div.sc-f6227726-5');
            // Check if the text element exists, is visible, and has actual content
            if (assignmentText && assignmentText.offsetParent !== null && assignmentText.textContent.trim() !== "") {
                isAlreadyAssigned = true;
                log(`Assignment detected via fallback text check: "${assignmentText.textContent.trim()}"`);
            } else {
                 isAlreadyAssigned = false;
                 log("Fallback text check indicates no assignment.");
            }
        }

        if (isAlreadyAssigned || autoAssigned) { // Check global flag too
            if (!autoAssigned && isAlreadyAssigned) { // Only log if we just detected it
                log("Someone is already assigned based on DOM checks. Skipping auto-assign logic.");
                autoAssigned = true; // Ensure flag is set if detected
            }
            return;
        } else {
            // log("No existing assignment found, proceeding with auto-assign attempt...");
        }

        // Get logged-in user and compare with sales rep.
        getLoggedInUserName(function(loggedInName) {
            if (!loggedInName) {
                // log("Logged in user not found yet."); // Expected sometimes
                return; // Exit, rely on next call
            }
            // log("Logged in user: " + loggedInName); // Reduce logging

            let salesRepInput = document.querySelector("#salesRep input");
            if (!salesRepInput) {
                // log("Sales rep element not found yet."); // Expected sometimes
                return; // Exit, rely on next call
            }
            let salesRepName = salesRepInput.getAttribute("title")?.trim(); // Added optional chaining
            if (!salesRepName) {
                 log("Sales rep name (title attribute) not found or empty.");
                 return; // Exit if name attribute is missing
            }
            // log("Sales rep name: " + salesRepName); // Reduce logging

            if (loggedInName !== salesRepName) {
                // log("Logged in user is not the sales rep. No auto assignment."); // Reduce logging
                return;
            }

            // Target the unique icon based on its text "Ŝ" and attribute color="#58698d"
            let iconButton = document.querySelector('h4[color="#58698d"]');
            if (iconButton && iconButton.textContent.trim() === "Ŝ") {
                log("Attempting to click assignment icon...");
                iconButton.click();
                log("Clicked the icon with Ŝ and color #58698d.");
            } else {
                // log("Could not find the icon button (or it\'s not the default icon)."); // Expected sometimes
                return; // Exit, rely on next call
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
                    log("Found matching team member, clicking...");
                    matchingMember.click();
                    log("Auto-assigned to " + loggedInName);
                    autoAssigned = true; // Set flag on successful assignment

                    // Final step: Click the "Assign" button (id="modalOK")
                    setTimeout(() => {
                        let assignButton = document.getElementById("modalOK");
                        if (assignButton && assignButton.textContent.trim() === "Assign") {
                            log("Clicking final \'Assign\' button...");
                            assignButton.click();
                            log("Clicked the \'Assign\' button at the bottom.");
                            // Wait a bit and then dispatch an Escape key event to close the logged-in user menu.
                            setTimeout(() => {
                                try { // Wrap in try-catch as menu might already be closed
                                    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
                                    log("Dispatched Escape key event to close the logged-in user menu.");
                                } catch(e) { log("Error dispatching Escape key: " + e.message, null, 'warn'); }
                            }, 1000);
                        } else {
                            log("Could not find the \'Assign\' button (modalOK) or text mismatch.");
                            // If assign button isn't found, maybe the modal closed automatically? Assume success.
                        }
                    }, 500);
                } else {
                    log("Could not find the logged in user in the assignment list.");
                     // Don't retry here, as the list was present but the user wasn't.
                }
            }, 1000);
        });
    }

    // -----------------------------
    // Product Options Remover & Search (Master Script Part 2)
    // -----------------------------

    // Hide Serial Tracking Option
    function hideSerialTrackingOption() {
        const serialTrackingElements = document.querySelectorAll('.SwitchSlideWithLabel-LabelText, .sc-8b7d6573-11, .jamuPn, label span'); // Add generic span inside label
        let found = false;
        serialTrackingElements.forEach(element => {
            if (element.textContent.trim() === 'Enable serial tracking') {
                // Find parent container more reliably
                let parentToHide = element.closest('.sc-1d26e07a-1, .foKSlp, .form-field, .settings-row'); // Add more potential parent classes
                 if (!parentToHide) parentToHide = element.parentElement?.parentElement; // Go up two levels as fallback

                if (parentToHide) {
                    parentToHide.style.display = 'none';
                    log('Hidden "Enable serial tracking" option');
                    found = true;
                }
            }
        });
        return found;
    }

    // Select Non-Stocked Product
    function selectNonStockedProduct() {
        const nonStockedLabels = Array.from(document.querySelectorAll('p.RadioButtonWithLabelContainer-LabelText, p.jamuPn, label span')) // Add generic span inside label
            .filter(el => el.textContent.trim() === 'Non-stocked product');
        if (nonStockedLabels.length === 0) { log('Could not find Non-stocked product label'); return false; }

        const labelElement = nonStockedLabels[0];
        const labelContainer = labelElement.closest('label.RadioButtonWithLabelContainer, label.lfDODJ');
        if (!labelContainer) { log('Could not find label container for Non-stocked'); return false; }

        const radioInput = labelContainer.querySelector('input[type="radio"]');
        if (!radioInput) { log('Could not find radio input for Non-stocked'); return false; }

        // Check if already selected
        if (radioInput.checked) {
            log("Non-stocked product already selected.");
            return true;
        }

        // Select it
        try {
            // Method 1: Click the label
            labelContainer.click();
            log('Clicked the label container for Non-stocked');

            // Method 2: Directly manipulate radio (fallback/reinforcement)
            if (!radioInput.checked) {
                radioInput.checked = true;
                log('Set radio input checked = true');
                // Dispatch events
                 radioInput.dispatchEvent(new Event('change', { bubbles: true }));
                 radioInput.dispatchEvent(new Event('input', { bubbles: true })); // Some frameworks use input
            }

            // Ensure visual state matches (if applicable - classes might change)
             const radioContainerDiv = radioInput.closest('div.sc-8793bcb8-0'); // Container with visual state
             if (radioContainerDiv) {
                 radioContainerDiv.classList.remove('eiTaSL'); // Unselected class?
                 radioContainerDiv.classList.add('hjmWMW'); // Selected class?
             }

            return true; // Assume success after clicking/setting
        } catch (e) {
            log('Error selecting Non-stocked product: ' + e.message, null, 'error');
            return false;
        }
    }

    // Check for and Modify Product Type Options
    function checkForProductTypeOptions() {
        // Only run if we are likely in a product creation/edit context
        const nameInput = document.querySelector('#new-product-name, #product-name-input'); // Check for name input
        const typeOptionsContainer = document.querySelector('.product-type-options, div[role="radiogroup"]'); // Check for options container
        if (!nameInput && !typeOptionsContainer) return false; // Not in the right context

        let optionsModified = hideSerialTrackingOption(); // Always try to hide serial tracking

        const options = document.querySelectorAll('div.iiNbRK, div[data-hasfocus], label.RadioButtonWithLabelContainer'); // More selectors for options

        if (options.length > 0) {
            log("Found product type options, modifying...");
            options.forEach(option => {
                const textElement = option.querySelector('p, span'); // Get text element
                if (!textElement) return;
                const text = textElement.textContent.trim();

                // Hide Stocked/Service
                if (text.includes('Stocked product') || text === 'Service') {
                    option.style.display = 'none';
                    log(`Hidden option: "${text}"`);
                    optionsModified = true;
                }
            });

            // Select Non-stocked
            if (selectNonStockedProduct()) {
                log('Successfully selected Non-stocked product option');
                optionsModified = true;
            }
        }

        if (optionsModified) {
            watchForCreateButtonClick(); // Watch for create button if options were modified
            return true;
        }
        return false;
    }

    // Watch for Create Button Click (Product Creation)
    function watchForCreateButtonClick() {
        const createButton = Array.from(document.querySelectorAll('button')).find(
             // More robust search: Check modal footers, text content, potential test IDs
             btn => (btn.textContent.trim() === 'Create' || btn.textContent.trim() === 'Save') &&
                    btn.offsetParent !== null && // Visible
                    btn.closest('.modal-footer, .popup-actions') // Likely inside a modal/popup
         );

        if (createButton && !createButton.hasAttribute('data-listener-added')) {
            log('Found Create/Save button in product popup');
            createButton.setAttribute('data-listener-added', 'true');
            createButton.addEventListener('click', function() {
                const nameInput = document.querySelector('#new-product-name, #product-name-input'); // Check both possible IDs
                if (nameInput) {
                    lastCreatedProductName = nameInput.value.trim();
                    log(`Captured product name: "${lastCreatedProductName}"`);
                    try {
                        localStorage.setItem('lastCreatedProduct', lastCreatedProductName);
                        log('Stored product name in localStorage');
                    } catch (e) { log('Failed to store product name in localStorage', null, 'error'); }
                    setTimeout(openProductsInNewTab, 1500); // Wait for save, then navigate
                } else { log('Could not find product name input when Create clicked', null, 'warn'); }
            });
            log('Added listener to Create/Save button');
        }
    }

    // Open Products Page in New Tab
    function openProductsInNewTab() {
        log('Opening Products page in new tab');
        const productsUrl = 'https://app.inflowinventory.com/products';
        const newTab = window.open(productsUrl, '_blank');
        if (newTab) {
            log('Successfully opened Products in a new tab');
            try { newTab.focus(); } catch (e) { log('Could not focus new tab', null, 'warn'); }
        } else {
            log('Popup blocked or unable to open new tab', null, 'error');
            showError("Could not open Products tab. Please disable popup blocker.");
        }
    }

    // Handle Product Search on Products Page
    function handleProductSearch() {
        if (!window.location.href.includes('/products')) return;

        const productName = localStorage.getItem('lastCreatedProduct');
        if (!productName) return; // No product stored

        log(`Found product name in localStorage: "${productName}"`);
        localStorage.removeItem('lastCreatedProduct'); // Consume the item

        setTimeout(() => {
            // Find search input/button more reliably
            let searchInput = document.getElementById('searchListing');
            let searchButton = null;

            if (!searchInput) {
                 // Look for button first (magnifying glass icon)
                 searchButton = document.querySelector('button [class*="lookup"], button [class*="search"]'); // Icon inside button
                 if (!searchButton) searchButton = document.querySelector('button[aria-label*="Search"]'); // Aria label
            }

            if (searchInput && searchInput.offsetParent !== null) {
                 log('Search input found and visible');
                 useNativeCommands(searchInput, productName, true); // Pass true to simulate Enter
            } else if (searchButton && searchButton.offsetParent !== null) {
                log('Search button found, clicking it');
                searchButton.click();
                setTimeout(() => checkForSearchInput(productName), 500); // Wait for input to appear
            } else {
                log('Could not find search input or button, retrying...', null, 'warn');
                setTimeout(() => handleProductSearch(), 1000); // Retry after delay
            }
        }, 1500); // Longer initial delay for products page load
    }

    // Check for Search Input after Click
    function checkForSearchInput(productName) {
        const searchInput = document.getElementById('searchListing') || document.querySelector('input[type="search"], input[placeholder*="Search"]');
        if (searchInput && searchInput.offsetParent !== null) {
            log('Found search input field after button click');
            useNativeCommands(searchInput, productName, true); // Pass true to simulate Enter
        } else {
            log('Could not find search input after button click, retrying', null, 'warn');
            setTimeout(() => checkForSearchInput(productName), 500); // Retry
        }
    }

    // Use Native Commands for Input (with optional Enter simulation)
    function useNativeCommands(inputElement, text, simulateEnter = false) {
        log(`Using native commands for text input: "${text}" (Simulate Enter: ${simulateEnter})`);
        // inputElement.focus(); // Let caller manage focus

        // Method 1: Direct value set + events (Simplified)
        inputElement.value = ''; // Clear first
        inputElement.dispatchEvent(new Event('input', { bubbles: true })); // Trigger input event for clearing
        inputElement.value = text;
        inputElement.dispatchEvent(new Event('input', { bubbles: true }));
        inputElement.dispatchEvent(new Event('change', { bubbles: true }));
        log('Set value directly and dispatched input/change events');

        // --- REINSTATED: Attempt to call React handler directly ---
        try {
            const reactKey = Object.keys(inputElement).find(key => key.startsWith('__reactProps') || key.startsWith('__reactFiber'));
            if (reactKey) {
                const reactProps = inputElement[reactKey];
                // Look for common handler names
                const handler = reactProps?.onChange || reactProps?.memoizedProps?.onChange || reactProps?.props?.onChange;
                if (handler && typeof handler === 'function') {
                     // Create a minimal event object that might satisfy the handler
                     const eventArg = {
                        target: inputElement,
                        currentTarget: inputElement,
                        type: 'change', // Mimic a change event
                        bubbles: true,
                        cancelable: true,
                        isTrusted: false, // Indicate it's synthetic
                        nativeEvent: { target: inputElement }, // Provide nested nativeEvent if needed
                        preventDefault: () => {},
                        stopPropagation: () => {}
                        // Add other properties if console errors indicate they are needed
                    };
                    handler(eventArg);
                    log("Attempted to call React onChange handler directly");
                } else {
                    log("React onChange handler not found or not a function.");
                }
            } else {
                 log("React properties key not found on input element.");
            }
        } catch (e) {
            log("Error calling React handler: " + e.message, null, 'warn');
        }
        // --- END REINSTATED ---



        // Blur/Refocus and Enter simulation (Keep this part, but maybe simplify delay)
        setTimeout(() => {
            inputElement.blur(); // <<< RE-ADD BLUR
            // setTimeout(() => {
                 // inputElement.focus(); // Avoid unnecessary blur/focus cycle if possible
                 // Press Enter (if requested) - Keep this part
                 if (simulateEnter) {
                     log('Simulating Enter key');
                     inputElement.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
                     inputElement.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
                     // Wait for search results (if applicable)
                     setTimeout(() => findAndClickProduct(text), 1000);
                 }
            // }, 100);
        }, 100); // Reduced delay slightly, removed nested timeout/blur/focus
    }


    // Find and Click Product in Search Results
    function findAndClickProduct(productName) {
        log(`Looking for product "${productName}" in the results`);
        let foundElement = null;

        // Robust search for product row/link
        const potentialElements = document.querySelectorAll(
            'li[role="option"], tr[data-id], div.product-row, a[href*="/products/"]' // More selectors
        );

        for (const item of potentialElements) {
            // Check text content, aria-label, title, etc.
             const textContent = item.textContent || "";
             const ariaLabel = item.getAttribute('aria-label') || "";
             const title = item.getAttribute('title') || "";

            if (textContent.includes(productName) || ariaLabel.includes(productName) || title.includes(productName)) {
                // Prefer a direct link if available within the item
                foundElement = item.querySelector(`a[href*="/products/"]`) || item;
                log(`Found potential product element: ${item.tagName}`);
                break;
            }
        }

         // Fallback: Look for exact text match in common elements if specific rows not found
         if (!foundElement) {
             const textElements = document.querySelectorAll('span, div, td, h4');
             for (const el of textElements) {
                 if (el.textContent.trim() === productName && el.offsetParent !== null) {
                     // Find a clickable ancestor
                     let clickableParent = el;
                     for(let i=0; i<5 && clickableParent; i++) {
                          if (clickableParent.tagName === 'A' || clickableParent.tagName === 'LI' || clickableParent.tagName === 'TR' || clickableParent.hasAttribute('onclick') || window.getComputedStyle(clickableParent).cursor === 'pointer') {
                              break;
                          }
                         clickableParent = clickableParent.parentElement;
                     }
                     foundElement = clickableParent || el; // Use clickable parent or element itself
                     log(`Found product by exact text match: ${foundElement.tagName}`);
                     break;
                 }
             }
         }


        if (foundElement && foundElement.offsetParent !== null) { // Check visibility
            log('Found product element, clicking it');
            highlightElement(foundElement); // Highlight before click
            setTimeout(() => {
                 try {
                    foundElement.click();
                    log('Successfully clicked on product: ' + productName);
                 } catch (e) { log('Error clicking product: ' + e.message, null, 'error'); }
            }, 300); // Small delay after highlight
        } else {
            log('Could not find the product in the results, will retry in 1.5 seconds', null, 'warn');
            setTimeout(() => findAndClickProduct(productName), 1500); // Retry after delay
        }
    }

    // ====================================
    // COMBINED INITIALIZATION & OBSERVER
    // ====================================

    let quoteUpdateCheckInterval = null; // Variable to hold the interval ID
    const MAX_QUOTE_UPDATE_CHECKS = 20; // Try for 10 seconds (20 * 500ms)
    let quoteUpdateCheckCount = 0;

    function initializeCombinedScript() {
        const url = window.location.href;
        log(`Initializing Combined Script on URL: ${url}`);

        // Cache the user name immediately
        cacheUserName();

        // --- Quote Converter Initializations ---
        // ** ADD: Try updating original quote immediately **
        updateOriginalQuote();

        if (isThisAQuotePage()) {
            addHeaderButton(); // Add main button
            scheduleButtonAdditionRetries(); // Set up retries
            handleQuotePage(); // Apply quote-specific UI tweaks from Master Script part
        }
        // ** ADD: Explicit call to handleQuotePage on initialization **
        handleQuotePage();

        // **REMOVED**: updateOriginalQuote(); // Don't call directly here, handled by interval/observer
        // Check if we need to process a *copied* quote (runs in the new tab after copy)
        if (!isOriginalQuotePage()) { // Safety check crucial here
             processCopiedQuote();
        }
        // REMOVED: addDebugButton(); // Add debug button always


        // --- Master Script Initializations ---
        if (url.includes('/sales-orders/')) {
            tryInsertPO();
            renamePOElements();
            removeFulfillButton();
            checkCreditHold();
            addCardOnFileIndicator();
            autoAssignIfSalesRep();
            // Reset customer tracking on SO page load
            localStorage.removeItem('lastPOCustomer');
            lastCardCustomerName = ""; // Reset card check cache
        } else if (url.includes('/products')) {
            // Handle product search if needed (triggered by localStorage item)
            const marker = document.getElementById('product-search-initialized');
            if (!marker) {
                log('On products page, setting up potential product search');
                const markerDiv = document.createElement('div');
                markerDiv.id = 'product-search-initialized'; markerDiv.style.display = 'none';
                document.body.appendChild(markerDiv);
                setTimeout(handleProductSearch, 1000); // Delay search initiation
            }
        } else {
            // Check for product creation/edit context on other pages
            checkForProductTypeOptions();
        }

        // --- Enhancement Initializations ---
        enhanceInternalNoteField(); // Attempt to enhance the internal note field on load
    }

    // Combined Mutation Observer (DEBOUNCED)
    const observer = new MutationObserver(debounce((mutations) => {
        // Use requestAnimationFrame to avoid running too often during rapid changes
        window.requestAnimationFrame(() => {
             let reInitNeeded = false;
             let runProductCheck = false;
             let runQuoteUpdateCheck = false;
             let runCopiedQuoteCheck = false;
              let runSalesOrderCheck = false;
              let runAddButtonCheck = false;
              let runAsSalesOrderCheck = false;
              let runInternalNoteCheck = false; // <-- Add flag for internal note
              let runSaveButtonListenerCheck = false; // <-- Add flag for save button listener
              let runQuotePageUICheck = false; // <-- Add flag for quote page UI check
              let runAssignCheck = false; // <-- ADDED: Specific flag for assignment section changes


             for (const mutation of mutations) {
                 if (mutation.addedNodes.length > 0 || mutation.removedNodes.length > 0) {
                      reInitNeeded = true; // General re-check on node changes

                      // More specific checks based on added nodes
                      mutation.addedNodes.forEach(node => {
                          if (node.nodeType === 1) { // Check elements only
                               // Product related checks
                               if (node.querySelector('#new-product-name, .product-type-options')) runProductCheck = true;
                               if (node.matches && node.matches('#modalOK, #modalCancel')) runProductCheck = true; // Modal buttons appear

                               // Quote/Order related checks
                               if (node.querySelector('#convertSalesOrder, #copyAndConvertButton, #entity-header-input')) runAddButtonCheck = true;
                               if (node.querySelector('#custom1-input, input[placeholder="Enter data"]')) runQuoteUpdateCheck = true;
                               if (node.matches && (node.matches('#salesOrder') || node.matches('#option-item-salesOrder'))) runAsSalesOrderCheck = true;
                               if (node.querySelector('#so_customer, #salesRep, textarea[name="remarks"]')) runSalesOrderCheck = true;

                               // Internal Note Check <-- Add this check
                               if (node.matches && node.matches('#custom2-input') || node.querySelector('#custom2-input')) {
                                   runInternalNoteCheck = true;
                               }
                               // ---- ADDED Check for Assigned To section ----
                               if (node.matches && node.matches('div.sc-3e2e0cb2-1.gtsOkx') || node.querySelector('div.sc-3e2e0cb2-1.gtsOkx')) {
                                   log("Observer detected Assigned To section added/modified.");
                                   runSalesOrderCheck = true; // Still relevant for SO context
                                   runAssignCheck = true; // Set specific flag
                               }
                               // ---- ADDED Check ----
                               if (node.matches && node.matches('#save') || node.querySelector('#save')) { runSaveButtonListenerCheck = true; }
                               // ---- Check for Quote Page UI elements ----
                               let foundQuoteUIElement = false;
                               // Check for the buttons first using valid selectors
                               if (node.querySelector('#convertSalesOrder, #copyAndConvertButton')) {
                                   foundQuoteUIElement = true;
                               } else {
                                   // If buttons not found, check for the specific paragraph text
                                   const pElements = node.querySelectorAll('p');
                                   for (const p of pElements) {
                                       // Use textContent.includes for robust matching
                                       if (p.textContent?.includes('Sales rep')) {
                                           foundQuoteUIElement = true;
                                           break; // Found one, no need to check further
                                       }
                                   }
                               }
                               // Set the flag if any relevant element was found
                               if (foundQuoteUIElement) {
                                   runQuotePageUICheck = true;
                               }
                          }
                      });
                 }
                  // Check for attribute changes that might signify state changes
                  if (mutation.type === 'attributes') {
                      if (mutation.target.id === 'entity-header-input') reInitNeeded = true; // Quote/Order number changed?
                      if (mutation.target.matches && mutation.target.matches('#so_customer input')) runSalesOrderCheck = true; // Customer changed?
                      // Also check if attributes change on the internal note input (e.g., value)
                      if (mutation.target.id === 'custom2-input') {
                           runInternalNoteCheck = true;
                      }
                      // ---- ADDED Check ----
                      if (mutation.target.id === 'save') { runSaveButtonListenerCheck = true; }
                      // Check for attributes relevant to quote page UI
                      if (mutation.attributeName === 'title' && mutation.target.matches('#entity-header-input') && mutation.target.title?.includes('SQ-')) {
                          runQuotePageUICheck = true;
                      }
                  }
             }

            // Run checks based on flags
             if (runAddButtonCheck || reInitNeeded) {
                 if (isThisAQuotePage() && !document.getElementById('copyAndConvertButton')) {
                     addHeaderButton();
                 }
             }
             if (runQuoteUpdateCheck || reInitNeeded) {
                  // --- TEMPORARILY DISABLED IN OBSERVER ---
                  /*
                   if (checkForPendingQuoteUpdate() && !document.getElementById('quote-already-updated')) {
                      updateOriginalQuote();
                  }
                  */
                  // --- END TEMPORARY DISABLE ---
              }
              if (runCopiedQuoteCheck || reInitNeeded) {
                  // --- TEMPORARILY DISABLED IN OBSERVER ---
                  
                   if (getStoredConversionInfo() && !document.getElementById('automation-already-started') && !isOriginalQuotePage()) {
                       setTimeout(processCopiedQuote, 500); // Small delay before processing copied quote
                   }
                  
                   // --- END TEMPORARY DISABLE ---
              }
              if (runSalesOrderCheck || reInitNeeded) {
                   if (window.location.href.includes('/sales-orders/')) {
                       tryInsertPO(); // This still runs, but no longer attaches the listener
                       addCardOnFileIndicator();
                       renamePOElements(); // Re-apply renaming
                       checkCreditHold(); // Re-check credit hold
                         // --- TEMPORARILY DISABLED IN OBSERVER ---
                        // Call autoAssign only if the specific section changed and focus guard allows
                        if (runAssignCheck) {
                           log("Assign section change detected, attempting auto-assign check.");
                           autoAssignIfSalesRep(); // Re-check assignment
                        }
                         // --- END TEMPORARY DISABLE ---
                   }
              }
              if (runProductCheck || reInitNeeded) {
                  checkForProductTypeOptions();
             }
             if (runAsSalesOrderCheck) {
                 // If the "As sales order" option appears and we have pending conversion info, click it
                 const salesOrderOption = document.getElementById('salesOrder') || document.getElementById('option-item-salesOrder');
                 if (salesOrderOption && getStoredConversionInfo()) {
                     debugLog("Detected 'As sales order' option appeared, clicking via observer");
                     clickAsSalesOrderOption();
                 }
             }
             // Run Internal Note Check <-- Add this call
             if (runInternalNoteCheck || reInitNeeded) { // Also check on reInitNeeded
                 enhanceInternalNoteField();
             }
             // Run Quote Page UI Check
             if (runQuotePageUICheck || reInitNeeded) {
                  if (isThisAQuotePage()) {
                      handleQuotePage();
                  }
             }
             // ---- ADDED Listener Attachment ----
             if ((runSaveButtonListenerCheck || reInitNeeded) && window.location.href.includes('/sales-orders/')) { // Only attach on SO pages
                  const saveButton = document.getElementById('save');
                  // Check visibility (offsetParent) and if listener is already attached
                  if (saveButton && saveButton.offsetParent !== null && !saveButton.hasAttribute('data-lookup-save-listener')) {
                      log("Attaching Look Up Number save handler to Save button (Bubble Phase).");
                      saveButton.addEventListener('click', handleSaveClick, false); // Use bubble phase (default)
                      saveButton.setAttribute('data-lookup-save-listener', 'true');
                  }
             }
        });
    }, 250)); // Debounce observer callback for 250ms

    // Observe the body for changes
    observer.observe(document.body, {
        childList: true,
        subtree: true,
        attributes: true,
         // Observe specific attributes that often change with state
         attributeFilter: ['style', 'class', 'disabled', 'title', 'id'] // Removed 'value' to reduce noise
    });

    // --- History change handling (SPA navigation) ---
    let lastUrl = window.location.href;

    function handleHistoryChange() {
        // Use requestAnimationFrame to ensure DOM is likely updated after navigation
        requestAnimationFrame(() => {
             if (lastUrl !== window.location.href) {
                 log("URL changed, re-initializing script...");
                 lastUrl = window.location.href;
                  // Reset relevant flags/states on navigation
                  creditHoldPopupDismissed = false;
                  lastCardCustomerName = "";
                  inFlightCustomerName = "";
                  autoAssigned = false;
                  quoteConversionHappened = false;
                  localStorage.removeItem('lastPOCustomer'); // Reset PO check cache
                  localStorage.removeItem('cachedUserName'); // Clear user name cache on navigation
                  // ** ADD ** Clear quote update interval on navigation
                  if (quoteUpdateCheckInterval) {
                       clearInterval(quoteUpdateCheckInterval);
                       quoteUpdateCheckInterval = null;
                       debugLog("Cleared quote update check interval due to navigation.");
                  }
                  quoteUpdateCheckCount = 0; // Reset counter

                 // Full re-initialization
                 setTimeout(startScriptChecks, 500); // Rerun the startup checks
                 // ** ADD: Clear lastPOCustomer on navigation **
                 localStorage.removeItem('lastPOCustomer');
                  // ** ADD: Call handleQuotePage after a short delay on navigation **
                  setTimeout(handleQuotePage, 600); // Ensure it runs after re-init
                  // ** ADD: Specific delayed check for auto-assign on SO pages after navigation **
                  if (window.location.href.includes('/sales-orders/')) {
                      log("Scheduling delayed auto-assign check after navigation to SO page.");
                      setTimeout(autoAssignIfSalesRep, 1500); // Delay check by 1.5 seconds
                  }
             }
        });
    }

    // Check and patch history.pushState
    try {
        if (history.pushState && typeof history.pushState === 'function') {
            const originalPushState = history.pushState; // Capture original function
            history.pushState = function() {
                const args = arguments;
                const self = this;
                // Check if originalPushState is still a function before calling
                if (typeof originalPushState === 'function') {
                    try {
                        let result = originalPushState.apply(self, args);
                        handleHistoryChange(); // Handle potential URL change
                        return result;
                    } catch (e) {
                         console.error('[InFlow Combined Script] Error calling original history.pushState:', e);
                         handleHistoryChange(); // Still try to handle the change
                         throw e; // Re-throw error to allow other handlers to see it
                    }
                } else {
                    console.error('[InFlow Combined Script] Original history.pushState seems to have been modified or removed!');
                    handleHistoryChange(); // Still try to handle the change
                    // Cannot reliably call the original, so we might break navigation here.
                    // Return undefined or throw an error? For now, just log and handle our part.
                    return undefined;
                }
            };
             log("Patched history.pushState");
        } else {
            log("Could not patch history.pushState - function not found or not a function.", null, 'warn');
        }
    } catch (e) {
        log("Error during history.pushState patching:", e, 'error');
    }

    // Check and patch history.replaceState
    try {
        if (history.replaceState && typeof history.replaceState === 'function') {
            const originalReplaceState = history.replaceState; // Capture original function
            history.replaceState = function() {
                const args = arguments;
                const self = this;
                // Check if originalReplaceState is still a function before calling
                if (typeof originalReplaceState === 'function') {
                    try {
                        let result = originalReplaceState.apply(self, args);
                         handleHistoryChange(); // Handle potential URL change
                        return result;
                    } catch (e) {
                        console.error('[InFlow Combined Script] Error calling original history.replaceState:', e);
                        handleHistoryChange(); // Still try to handle the change
                        throw e; // Re-throw error
                    }
                } else {
                    console.error('[InFlow Combined Script] Original history.replaceState seems to have been modified or removed!');
                    handleHistoryChange(); // Still try to handle the change
                    return undefined;
                }
            };
             log("Patched history.replaceState");
        } else {
             log("Could not patch history.replaceState - function not found or not a function.", null, 'warn');
        }
    } catch (e) {
        log("Error during history.replaceState patching:", e, 'error');
    }


    window.addEventListener('popstate', handleHistoryChange);

     // Listen for the convert button click to set the flag
     document.addEventListener('click', function(e) {
         let convertBtn = e.target.closest('#convertSalesOrder');
         if (convertBtn) {
             quoteConversionHappened = true;
             log("Convert to order button clicked. Setting flag.");
             // Store a flag in localStorage to indicate we need to auto-assign after navigation
             localStorage.setItem('pendingAutoAssign', 'true');
             localStorage.setItem('autoAssignTimestamp', Date.now().toString());
         }
     }, true); // Use capture phase

    // Add a function to check and handle auto-assign after navigation
    function checkForPendingAutoAssign() {
        const pendingAutoAssign = localStorage.getItem('pendingAutoAssign');
        const autoAssignTimestamp = localStorage.getItem('autoAssignTimestamp');

        if (pendingAutoAssign && autoAssignTimestamp) {
            const timestamp = parseInt(autoAssignTimestamp);
            const now = Date.now();

            // Only process if within last 30 seconds
            if (now - timestamp < 30000) {
                if (window.location.href.includes('/sales-orders/')) {
                    log("Found pending auto-assign request, attempting to assign");
                    // Clear the pending flag
                    localStorage.removeItem('pendingAutoAssign');
                    localStorage.removeItem('autoAssignTimestamp');

                    // Function to attempt auto-assign with retries
                    function attemptAutoAssign(attempts = 0, maxAttempts = 10) {
                        if (attempts >= maxAttempts) {
                            log("Max attempts reached for auto-assign, giving up");
                            return;
                        }

                        // Check if the assigned section exists
                        const assignedSection = document.querySelector("div.sc-3e2e0cb2-1.gtsOkx");
                        if (assignedSection) {
                            log("Found assigned section, triggering auto-assign");
                            autoAssignIfSalesRep();
                        } else {
                            log(`Attempt ${attempts + 1}: Assigned section not found, retrying...`);
                            setTimeout(() => attemptAutoAssign(attempts + 1), 1000);
                        }
                    }

                    // Start the auto-assign attempt process
                    setTimeout(() => attemptAutoAssign(), 1000);
                }
            } else {
                // Clear old pending auto-assign
                localStorage.removeItem('pendingAutoAssign');
                localStorage.removeItem('autoAssignTimestamp');
            }
        }
    }

    // Check for pending auto-assign on page load
    window.addEventListener('load', checkForPendingAutoAssign);

    // Also check when URL changes (for single-page app navigation)
    let autoAssignLastUrl = window.location.href;
    new MutationObserver(() => {
        const currentUrl = window.location.href;
        if (currentUrl !== autoAssignLastUrl) {
            autoAssignLastUrl = currentUrl;
            checkForPendingAutoAssign();
        }
    }).observe(document, { subtree: true, childList: true });

    // --- Initial Load ---
    function startScriptChecks() {
        // ** REMOVE: Special handling for quoteUpdateWindow **
        /* ... */
        // ** REPLACE with direct call **
        initializeCombinedScript();
        // ** ADD: Schedule multiple handleQuotePage checks after initial load **
        setTimeout(handleQuotePage, 1000);
        setTimeout(handleQuotePage, 2500);
        setTimeout(handleQuotePage, 5000);
    }

    function startScript() {
        log('Combined InFlow Script starting...');
        // ** CHANGE: Call the check function instead of direct init **
        setTimeout(startScriptChecks, 800); // Initial delay
        // Subsequent checks might not be needed if interval/observer works
        // setTimeout(startScriptChecks, 2000);
        // setTimeout(startScriptChecks, 5000);
    }

    // Start on load/ready
    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        startScript();
    } else {
        window.addEventListener('DOMContentLoaded', startScript);
    }

    // Periodic check for Copy & Convert button as a final fallback
    setInterval(() => {
        if (isThisAQuotePage() && !document.getElementById('copyAndConvertButton')) {
            log("Periodic check: Copy & Convert button missing, attempting add.");
            addHeaderButton();
        }
         // Periodic check for credit hold status
         if (window.location.href.includes('/sales-orders/')) {
             checkCreditHold();
         }
         // Periodic check for internal note field enhancement <-- REMOVED THIS CHECK
         // enhanceInternalNoteField();
         // ** ADD: Explicit periodic check for quote page UI **
         if (isThisAQuotePage()) {
             handleQuotePage();
         }

    }, 10000); // Check every 10 seconds

    log('InFlow Combined Script loaded and running.');

    // Add periodic cache update
    setInterval(() => {
        cacheUserName(); // Update the cached name periodically
    }, 60000); // Every minute

    // --- ADDED: Listen for storage changes for SO refresh ---
    window.addEventListener('storage', handleStorageChange);
    // --- END ADDED ---

    // --- ADDED: Handle localStorage changes for SO Refresh ---
    function handleStorageChange(event) {
        // Only act if we are on a sales order page (where refresh makes sense)
        if (!window.location.href.includes('/sales-orders/')) return;

        if (event.key === 'quoteConversionCompletedTimestamp' && event.newValue !== null) {
            debugLog("Received quote conversion completion signal via localStorage.", event);
            const completionTimestamp = parseInt(event.newValue, 10);
            const now = Date.now();
            // Check if the signal is recent (e.g., within 15 seconds)
            if (!isNaN(completionTimestamp) && (now - completionTimestamp < 15000)) {
                debugLog("Completion signal is recent. Refreshing Sales Order page.");
                // Clean up the flag immediately to prevent loops
                localStorage.removeItem('quoteConversionCompletedTimestamp');
                // Show a brief notification
                showStatusNotification("Quote conversion complete. Refreshing page...", "success");
                // Reload after a short delay to allow notification to show
                setTimeout(() => {
                    window.location.reload();
                }, 1500);
            } else {
                debugLog("Ignoring old or invalid quote completion signal.", { now, completionTimestamp }, 'warn');
                // Clean up potentially stale flag
                localStorage.removeItem('quoteConversionCompletedTimestamp');
            }
        }
    }
    // --- END ADDED ---

})(); 
