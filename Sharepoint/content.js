// Content Script for SharePoint JSOM Field Creator
// This script runs in the context of SharePoint pages and handles JSOM operations

// Inject the page context script
injectPageContextScript();

// Listen for messages from popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    console.log('Content script received message:', request);

    if (request.action === 'connect') {
        handleConnect(request, sendResponse);
        return true; // Keep message channel open for async response
    }

    if (request.action === 'createField') {
        handleCreateField(request, sendResponse);
        return true; // Keep message channel open for async response
    }

    if (request.action === 'createList') {
        handleCreateList(request, sendResponse);
        return true; // Keep message channel open for async response
    }

    if (request.action === 'getAllLists') {
        handleGetAllLists(request, sendResponse);
        return true; // Keep message channel open for async response
    }

    if (request.action === 'previewListDeleterItems') {
        handlePreviewListDeleterItems(request, sendResponse);
        return true; // Keep message channel open for async response
    }

    if (request.action === 'deleteList') {
        handleDeleteList(request, sendResponse);
        return true; // Keep message channel open for async response
    }

    if (request.action === 'deleteListItems') {
        handleDeleteListItems(request, sendResponse);
        return true; // Keep message channel open for async response
    }

    if (request.action === 'ping') {
        sendResponse({ success: true, message: 'Content script is ready' });
        return true;
    }

    if (request.action === 'checkAllCheckboxes') {
        // Directly check all checkboxes on the page
        const checkboxes = document.querySelectorAll('input[type="checkbox"]');
        let checkedCount = 0;
        checkboxes.forEach(checkbox => {
            if (!checkbox.disabled) {
                checkbox.checked = true;
                checkbox.dispatchEvent(new Event('change', { bubbles: true }));
                checkedCount++;
            }
        });
        sendResponse({ success: true, message: `Checked ${checkedCount} checkboxes` });
        return true;
    }

    if (request.action === 'getListFields') {
        handleGetListFields(request, sendResponse);
        return true; // Keep message channel open for async response
    }

    if (request.action === 'filterListItems') {
        handleFilterListItems(request, sendResponse);
        return true; // Keep message channel open for async response
    }

    if (request.action === 'uncheckAllCheckboxes') {
        // Directly uncheck all checkboxes on the page
        const checkboxes = document.querySelectorAll('input[type="checkbox"]');
        let uncheckedCount = 0;
        checkboxes.forEach(checkbox => {
            if (!checkbox.disabled) {
                checkbox.checked = false;
                checkbox.dispatchEvent(new Event('change', { bubbles: true }));
                uncheckedCount++;
            }
        });
        sendResponse({ success: true, message: `Unchecked ${uncheckedCount} checkboxes` });
        return true;
    }
});

// Inject script into page context to access SharePoint JSOM
function injectPageContextScript() {
    const script = document.createElement('script');
    script.src = chrome.runtime.getURL('jsom_injector.js');
    script.onload = function() {
        this.remove();
    };
    (document.head || document.documentElement).appendChild(script);
}

// Handle connection test
function handleConnect(request, sendResponse) {
    const { siteUrl, listName } = request;

    // Listen for response from page context FIRST (before sending message)
    const messageHandler = (event) => {
        if (event.data.type === 'SP_FIELD_CREATOR_CONNECT_RESPONSE') {
            window.removeEventListener('message', messageHandler);
            sendResponse(event.data.response);
        }
    };
    window.addEventListener('message', messageHandler);

    // Send message to page context script AFTER listener is attached
    window.postMessage({
        type: 'SP_FIELD_CREATOR_CONNECT',
        siteUrl: siteUrl,
        listName: listName
    }, '*');
}

// Handle field creation
function handleCreateField(request, sendResponse) {
    const { siteUrl, listName, fieldData } = request;

    // Listen for response from page context FIRST
    const messageHandler = (event) => {
        if (event.data.type === 'SP_FIELD_CREATOR_CREATE_FIELD_RESPONSE') {
            window.removeEventListener('message', messageHandler);
            sendResponse(event.data.response);
        }
    };
    window.addEventListener('message', messageHandler);

    // Send message to page context script AFTER listener is attached
    window.postMessage({
        type: 'SP_FIELD_CREATOR_CREATE_FIELD',
        siteUrl: siteUrl,
        listName: listName,
        fieldData: fieldData
    }, '*');
}

// Handle list creation
function handleCreateList(request, sendResponse) {
    const { siteUrl, listName } = request;

    // Listen for response from page context FIRST
    const messageHandler = (event) => {
        if (event.data.type === 'SP_FIELD_CREATOR_CREATE_LIST_RESPONSE') {
            window.removeEventListener('message', messageHandler);
            sendResponse(event.data.response);
        }
    };
    window.addEventListener('message', messageHandler);

    // Send message to page context script AFTER listener is attached
    window.postMessage({
        type: 'SP_FIELD_CREATOR_CREATE_LIST',
        siteUrl: siteUrl,
        listName: listName
    }, '*');
}

// Handle get all lists
function handleGetAllLists(request, sendResponse) {
    // Listen for response from page context FIRST
    const messageHandler = (event) => {
        if (event.data.type === 'SP_FIELD_CREATOR_GET_ALL_LISTS_RESPONSE') {
            window.removeEventListener('message', messageHandler);
            sendResponse(event.data.response);
        }
    };
    window.addEventListener('message', messageHandler);

    // Send message to page context script AFTER listener is attached
    window.postMessage({
        type: 'SP_FIELD_CREATOR_GET_ALL_LISTS'
    }, '*');
}

// Handle preview list deleter items
function handlePreviewListDeleterItems(request, sendResponse) {
    const { listTitle, camlQuery, folderPath } = request;

    // Listen for response from page context FIRST
    const messageHandler = (event) => {
        if (event.data.type === 'SP_FIELD_CREATOR_PREVIEW_ITEMS_RESPONSE') {
            window.removeEventListener('message', messageHandler);
            sendResponse(event.data.response);
        }
    };
    window.addEventListener('message', messageHandler);

    // Send message to page context script AFTER listener is attached
    window.postMessage({
        type: 'SP_FIELD_CREATOR_PREVIEW_ITEMS',
        listTitle: listTitle,
        camlQuery: camlQuery,
        folderPath: folderPath
    }, '*');
}

// Handle delete list
function handleDeleteList(request, sendResponse) {
    const { listTitle } = request;

    // Listen for response from page context FIRST
    const messageHandler = (event) => {
        if (event.data.type === 'SP_FIELD_CREATOR_DELETE_LIST_RESPONSE') {
            window.removeEventListener('message', messageHandler);
            sendResponse(event.data.response);
        }
    };
    window.addEventListener('message', messageHandler);

    // Send message to page context script AFTER listener is attached
    window.postMessage({
        type: 'SP_FIELD_CREATOR_DELETE_LIST',
        listTitle: listTitle
    }, '*');
}

// Handle delete list items
function handleDeleteListItems(request, sendResponse) {
    const { listTitle, camlQuery, folderPath } = request;

    // Listen for response from page context FIRST
    const messageHandler = (event) => {
        if (event.data.type === 'SP_FIELD_CREATOR_DELETE_ITEMS_RESPONSE') {
            window.removeEventListener('message', messageHandler);
            sendResponse(event.data.response);
        }
    };
    window.addEventListener('message', messageHandler);

    // Send message to page context script AFTER listener is attached
    window.postMessage({
        type: 'SP_FIELD_CREATOR_DELETE_ITEMS',
        listTitle: listTitle,
        camlQuery: camlQuery,
        folderPath: folderPath
    }, '*');
}

// Log that content script is loaded
console.log('SharePoint Field Creator - Content Script Loaded');

// Handle get list fields
function handleGetListFields(request, sendResponse) {
    const { listTitle } = request;

    // Listen for response from page context FIRST
    const messageHandler = (event) => {
        if (event.data.type === 'SP_FIELD_CREATOR_GET_LIST_FIELDS_RESPONSE') {
            window.removeEventListener('message', messageHandler);
            sendResponse(event.data.response);
        }
    };
    window.addEventListener('message', messageHandler);

    // Send message to page context script AFTER listener is attached
    window.postMessage({
        type: 'SP_FIELD_CREATOR_GET_LIST_FIELDS',
        listTitle: listTitle
    }, '*');
}

// Handle filter list items
function handleFilterListItems(request, sendResponse) {
    const { listTitle, camlQuery, rowLimit } = request;

    // Listen for response from page context FIRST
    const messageHandler = (event) => {
        if (event.data.type === 'SP_FIELD_CREATOR_FILTER_ITEMS_RESPONSE') {
            window.removeEventListener('message', messageHandler);
            sendResponse(event.data.response);
        }
    };
    window.addEventListener('message', messageHandler);

    // Send message to page context script AFTER listener is attached
    window.postMessage({
        type: 'SP_FIELD_CREATOR_FILTER_ITEMS',
        listTitle: listTitle,
        camlQuery: camlQuery,
        rowLimit: rowLimit
    }, '*');
}
