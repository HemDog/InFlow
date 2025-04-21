// ==UserScript==
// @name         InFlow Non-Stocked Product Automation
// @namespace    http://yourdomain.com
// @version      1.0.0
// @description  Automates selecting Non-stocked product type, hides other options, and searches for the created product.
// @match        https://app.inflowinventory.com/*
// @grant        none
// @downloadURL  https://raw.githubusercontent.com/HemDog/InFlow/main/inflow-nonstocked-automation.user.js
// @updateURL    https://raw.githubusercontent.com/HemDog/InFlow/main/inflow-nonstocked-automation.user.js
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    // ====================================
    // Log Helper
    // ====================================
    function log(message) {
        console.log(`[InFlow Script] ${message}`);
    }

    // ====================================
    // PRODUCT OPTIONS REMOVER & SEARCH AUTOMATION
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
    // INITIALIZATION & OBSERVATION
    // ====================================

    // Main initialization function for the focused features
    function initProductAutomation() {
        const url = window.location.href;

        if (url.includes('/products')) {
            // We're on the products page - check if we need to search
            const marker = document.getElementById('product-search-initialized');
            if (!marker) {
                log('On products page, setting up product search');

                // Add a marker to avoid duplicate initialization in this tab session
                const markerElem = document.createElement('div');
                markerElem.id = 'product-search-initialized';
                markerElem.style.display = 'none';
                document.body.appendChild(markerElem);

                // Handle product search after a delay
                setTimeout(handleProductSearch, 800);
            }
        } else {
            // Check for product creation dialogs on other pages
            checkForProductTypeOptions();
        }
    }

    // Use MutationObserver to detect dynamically loaded content
    let observer = new MutationObserver(() => {
        initProductAutomation();
    });
    observer.observe(document.body, { childList: true, subtree: true });

    // History change handlers (for SPA navigation)
    const originalPushState = history.pushState;
    const originalReplaceState = history.replaceState;

    function handleHistoryChange() {
        log("History changed, re-initializing product automation checks.");
        // Reset the search marker on navigation
        const marker = document.getElementById('product-search-initialized');
        if (marker) {
            marker.remove();
        }
        // Re-run initialization logic
        startScript();
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
        log("Starting product automation script checks...");
        // Start the script with a delay to ensure page elements might be loaded
        setTimeout(initProductAutomation, 800);

        // Add additional checks after longer delays to catch late-loading elements
        setTimeout(initProductAutomation, 1500);
        setTimeout(initProductAutomation, 3000);
    }

    startScript();

    // Optional: Run checks periodically to catch missed changes (can be adjusted or removed if observer is reliable)
    // setInterval(initProductAutomation, 5000); // Reduced frequency

    log('InFlow Non-Stocked Product Automation script loaded.');

})();
