// Global state
let csvData = [];
let isConnected = false;
let currentTabId = null;

// DOM Elements
const elements = {
    welcomeScreen: document.getElementById('welcomeScreen'),
    mainForm: document.getElementById('mainForm'),
    checkUncheckView: document.getElementById('checkUncheckView'),
    openFormBtn: document.getElementById('openFormBtn'),
    checkUncheckBtn: document.getElementById('checkUncheckBtn'),
    backBtn: document.getElementById('backBtn'),
    backToWelcomeBtn: document.getElementById('backToWelcomeBtn'),
    checkAllBtn: document.getElementById('checkAllBtn'),
    uncheckAllBtn: document.getElementById('uncheckAllBtn'),
    checkboxResult: document.getElementById('checkboxResult'),
    checkboxResultText: document.getElementById('checkboxResultText'),
    connectionStatus: document.getElementById('connectionStatus'),
    siteUrl: document.getElementById('siteUrl'),
    listName: document.getElementById('listName'),
    connectBtn: document.getElementById('connectBtn'),
    csvFile: document.getElementById('csvFile'),
    dropZone: document.getElementById('dropZone'),
    fileInfo: document.getElementById('fileInfo'),
    fileName: document.getElementById('fileName'),
    clearFile: document.getElementById('clearFile'),
    previewSection: document.getElementById('previewSection'),
    previewBody: document.getElementById('previewBody'),
    fieldCount: document.getElementById('fieldCount'),
    progressSection: document.getElementById('progressSection'),
    progressFill: document.getElementById('progressFill'),
    progressText: document.getElementById('progressText'),
    resultsSection: document.getElementById('resultsSection'),
    resultsContent: document.getElementById('resultsContent'),
    createFieldsBtn: document.getElementById('createFieldsBtn'),
    resetBtn: document.getElementById('resetBtn')
};

// Data Type Mappings for SharePoint JSOM
const dataTypeMappings = {
    'text': 'SP.FieldText',
    'note': 'SP.FieldMultiLineText',
    'number': 'SP.FieldNumber',
    'integer': 'SP.FieldNumber',
    'currency': 'SP.FieldCurrency',
    'datetime': 'SP.FieldDateTime',
    'date': 'SP.FieldDateTime',
    'boolean': 'SP.FieldBoolean',
    'yesno': 'SP.FieldBoolean',
    'user': 'SP.FieldUser',
    'lookup': 'SP.FieldLookup',
    'choice': 'SP.FieldChoice',
    'url': 'SP.FieldURL',
    'hyperlink': 'SP.FieldURL',
    'counter': 'SP.FieldCounter',
    'calculated': 'SP.FieldCalculated',
    'guid': 'SP.FieldGuid',
    'attachment': 'SP.FieldAttachments'
};

// Initialize
document.addEventListener('DOMContentLoaded', initialize);

async function initialize() {
    // Setup event listeners first
    setupEventListeners();

    // Get current tab and auto-detect SharePoint info
    await autoDetectSharePointInfo();

    // Load saved configuration as backup
    loadSavedConfig();
}

function setupEventListeners() {
    // Open Form button (SharePoint Field Creator)
    if (elements.openFormBtn) {
        elements.openFormBtn.addEventListener('click', openMainForm);
    }

    // Check and Uncheck button
    if (elements.checkUncheckBtn) {
        elements.checkUncheckBtn.addEventListener('click', openCheckUncheckView);
    }

    // Back button
    if (elements.backBtn) {
        elements.backBtn.addEventListener('click', closeMainForm);
    }

    // Back to Welcome button (from Check and Uncheck view)
    if (elements.backToWelcomeBtn) {
        elements.backToWelcomeBtn.addEventListener('click', closeCheckUncheckView);
    }

    // Check All button
    if (elements.checkAllBtn) {
        elements.checkAllBtn.addEventListener('click', checkAllCheckboxes);
    }

    // Uncheck All button
    if (elements.uncheckAllBtn) {
        elements.uncheckAllBtn.addEventListener('click', uncheckAllCheckboxes);
    }

    // Connect button
    elements.connectBtn.addEventListener('click', connectToSharePoint);

    // Refresh button
    const refreshBtn = document.getElementById('refreshBtn');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', async () => {
            await autoDetectSharePointInfo();
        });
    }

    // File input
    elements.csvFile.addEventListener('change', handleFileSelect);

    // Drag and drop
    elements.dropZone.addEventListener('dragover', handleDragOver);
    elements.dropZone.addEventListener('dragleave', handleDragLeave);
    elements.dropZone.addEventListener('drop', handleDrop);

    // Clear file
    elements.clearFile.addEventListener('click', clearFile);

    // Create fields button
    elements.createFieldsBtn.addEventListener('click', createFields);

    // Reset button
    elements.resetBtn.addEventListener('click', resetForm);

    // Save config on change (removed - fields are now auto-detected)
    // elements.siteUrl.addEventListener('change', saveConfig);
    // elements.listName.addEventListener('change', saveConfig);
}

function getCurrentTab() {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
        if (tabs[0]) {
            currentTabId = tabs[0].id;
        }
    });
}

// Auto-detect SharePoint Site URL and List Name from current tab URL
async function autoDetectSharePointInfo() {
    try {
        const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tabs[0]) {
            showConnectionStatus('error', 'Could not access current tab');
            return;
        }

        const currentUrl = tabs[0].url;
        currentTabId = tabs[0].id;

        // Check if we're on a SharePoint page
        if (!isSharePointUrl(currentUrl)) {
            showConnectionStatus('error', 'Please navigate to a SharePoint page first');
            elements.siteUrl.placeholder = 'Not a SharePoint page';
            elements.listName.placeholder = 'Not a SharePoint page';
            return;
        }

        // Parse SharePoint URL to extract site URL and list name
        const { siteUrl, listName } = parseSharePointUrl(currentUrl);

        if (siteUrl) {
            elements.siteUrl.value = siteUrl;
        }

        if (listName) {
            elements.listName.value = listName;
            elements.listName.readOnly = true;
            elements.listName.style.background = '#f5f5f5';
            elements.listName.style.cursor = 'not-allowed';
            showConnectionStatus('connected', `Auto-detected: ${listName}`);
        } else {
            // No list detected - make field editable for manual entry
            elements.listName.value = '';
            elements.listName.readOnly = false;
            elements.listName.placeholder = 'Enter list name manually...';
            elements.listName.style.background = 'white';
            elements.listName.style.cursor = 'text';
            elements.listName.style.border = '1px solid #667eea';

            // Update badge
            const badgeText = document.getElementById('badgeText');
            badgeText.textContent = 'Site URL detected - enter list name manually';

            showConnectionStatus('connected', 'Site URL detected - enter list name');
        }

    } catch (error) {
        console.error('Auto-detect error:', error);
        showConnectionStatus('error', 'Failed to auto-detect SharePoint info');
    }
}

// Check if URL is a SharePoint URL
function isSharePointUrl(url) {
    if (!url) return false;

    const sharePointPatterns = [
        // SharePoint on-premises patterns
        /\/Pages\/.*\.aspx$/i,
        /\/Lists\//i,
        /\/_layouts\//i,
        /\/SitePages\//i,
        /\/sites\//i,
        // Additional patterns
        /\/default\.aspx$/i,
        /\/allitems\.aspx$/i,
        /\/listedit\.aspx$/i,
        /\/viewedit\.aspx$/i
    ];

    return sharePointPatterns.some(pattern => pattern.test(url));
}

// Parse SharePoint URL to extract site URL and list name
function parseSharePointUrl(url) {
    try {
        const urlObj = new URL(url);
        let siteUrl = '';
        let listName = '';

        const pathname = urlObj.pathname;

        // Pattern 1: /Lists/ListName/...
        // Example: https://server/sites/sitename/Lists/CustomList/AllItems.aspx
        const listsMatch = pathname.match(/(.+)\/Lists\/([^\/]+)/i);
        if (listsMatch) {
            siteUrl = urlObj.origin + listsMatch[1];
            listName = decodeURIComponent(listsMatch[2]).replace(/_/g, ' ');
        }
        // Pattern 2: Site pages with List parameter
        // Example: https://server/sites/sitename/Pages/Home.aspx?List=...
        else if (pathname.includes('/Pages/') || pathname.includes('/SitePages/')) {
            // Extract site URL (everything before /Pages or /SitePages)
            const pagesIndex = pathname.indexOf('/Pages');
            const sitePagesIndex = pathname.indexOf('/SitePages');

            let basePath = pathname;
            if (pagesIndex !== -1) {
                basePath = pathname.substring(0, pagesIndex);
            } else if (sitePagesIndex !== -1) {
                basePath = pathname.substring(0, sitePagesIndex);
            }

            siteUrl = urlObj.origin + basePath;

            // Try to get list name from URL parameter
            const listParam = urlObj.searchParams.get('List');
            if (listParam) {
                // List parameter might be GUID or encoded name
                // We'll set a placeholder for GUID-based lists
                listName = listParam.match(/^[{]?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}[}]?$/)
                    ? 'List (GUID detected)'
                    : listParam;
            }
        }
        // Pattern 3: List settings page
        // Example: https://server/sites/sitename/_layouts/15/listedit.aspx?List=...
        else if (pathname.includes('/_layouts/') && pathname.includes('listedit.aspx')) {
            // Get the path before /_layouts/
            const layoutsIndex = pathname.indexOf('/_layouts/');
            siteUrl = urlObj.origin + pathname.substring(0, layoutsIndex);

            const listParam = urlObj.searchParams.get('List');
            if (listParam) {
                listName = 'List (from settings)';
            }
        }
        // Pattern 4: Site home page
        // Example: https://server/sites/sitename/default.aspx
        // Example: https://server/sites/sitename
        else if (pathname.endsWith('/default.aspx') || pathname.match(/\/sites\/[^\/]+\/?$/)) {
            const cleanPath = pathname.replace('/default.aspx', '').replace(/\/$/, '');
            siteUrl = urlObj.origin + cleanPath;
            listName = ''; // No specific list
        }
        // Pattern 5: Root site collection
        // Example: https://server/Lists/ListName/AllItems.aspx
        else if (pathname.startsWith('/Lists/')) {
            siteUrl = urlObj.origin;
            const rootListsMatch = pathname.match(/\/Lists\/([^\/]+)/i);
            if (rootListsMatch) {
                listName = decodeURIComponent(rootListsMatch[1]).replace(/_/g, ' ');
            }
        }
        // Fallback: Just use the origin for any SharePoint-looking URL
        else {
            siteUrl = urlObj.origin;
            listName = '';
        }

        return { siteUrl, listName };

    } catch (error) {
        console.error('Error parsing SharePoint URL:', error);
        return { siteUrl: '', listName: '' };
    }
}

function loadSavedConfig() {
    chrome.storage.local.get(['siteUrl', 'listName'], (result) => {
        if (result.siteUrl) elements.siteUrl.value = result.siteUrl;
        if (result.listName) elements.listName.value = result.listName;
    });
}

function saveConfig() {
    chrome.storage.local.set({
        siteUrl: elements.siteUrl.value,
        listName: elements.listName.value
    });
}

// Connection Functions
async function connectToSharePoint() {
    const siteUrl = elements.siteUrl.value.trim();
    const listName = elements.listName.value.trim();

    if (!siteUrl || !listName) {
        showConnectionStatus('error', 'Please enter Site URL and List Name');
        return;
    }

    elements.connectBtn.disabled = true;
    elements.connectBtn.textContent = 'Connecting...';

    try {
        // Send message to content script to test connection
        const response = await sendMessageToContentScript({
            action: 'connect',
            siteUrl: siteUrl,
            listName: listName
        });

        console.log('[Popup] Connection response:', response);
        if (response && response.success) {
            isConnected = true;
            showConnectionStatus('connected', 'Connected to SharePoint');
            elements.createFieldsBtn.disabled = csvData.length === 0;
        } else if (response && response.listNotFound) {
            // List doesn't exist - prompt user to create it
            console.log('[Popup] List not found - prompting user');
            const createList = confirm(`${response.message}\n\nClick OK to create the list, or Cancel to abort.`);
            if (createList) {
                await createNewList(siteUrl, listName);
            } else {
                isConnected = false;
                showConnectionStatus('error', 'Connection cancelled - list not created');
            }
        } else {
            console.log('[Popup] Connection failed:', response);
            isConnected = false;
            showConnectionStatus('error', response?.message || 'Connection failed');
        }
    } catch (error) {
        isConnected = false;
        showConnectionStatus('error', 'Connection error: ' + error.message);
    } finally {
        elements.connectBtn.disabled = false;
        elements.connectBtn.textContent = 'Connect';
    }
}

// Create a new list in SharePoint
async function createNewList(siteUrl, listName) {
    elements.connectBtn.disabled = true;
    elements.connectBtn.textContent = 'Creating list...';

    try {
        const response = await sendMessageToContentScript({
            action: 'createList',
            siteUrl: siteUrl,
            listName: listName
        });

        if (response && response.success) {
            isConnected = true;
            showConnectionStatus('connected', `List "${listName}" created and connected!`);
            elements.createFieldsBtn.disabled = csvData.length === 0;
        } else {
            isConnected = false;
            showConnectionStatus('error', response?.message || 'Failed to create list');
        }
    } catch (error) {
        isConnected = false;
        showConnectionStatus('error', 'Error creating list: ' + error.message);
    } finally {
        elements.connectBtn.disabled = false;
        elements.connectBtn.textContent = 'Connect';
    }
}

function showConnectionStatus(status, message) {
    elements.connectionStatus.className = 'status-indicator ' + status;
    elements.connectionStatus.querySelector('.status-text').textContent = message;
}

// File Handling Functions
function handleDragOver(e) {
    e.preventDefault();
    elements.dropZone.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    elements.dropZone.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    elements.dropZone.classList.remove('dragover');

    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        processFile(file);
    }
}

function processFile(file) {
    if (!file.name.endsWith('.csv')) {
        alert('Please upload a CSV file.');
        return;
    }

    elements.fileName.textContent = file.name;
    elements.fileInfo.classList.remove('hidden');
    elements.dropZone.classList.add('hidden');

    const reader = new FileReader();
    reader.onload = (e) => {
        parseCSV(e.target.result);
    };
    reader.readAsText(file);
}

function clearFile() {
    csvData = [];
    elements.csvFile.value = '';
    elements.fileInfo.classList.add('hidden');
    elements.dropZone.classList.remove('hidden');
    elements.previewSection.classList.add('hidden');
    elements.createFieldsBtn.disabled = true;
}

// CSV Parsing
function parseCSV(csvText) {
    const lines = csvText.split('\n').filter(line => line.trim());
    csvData = [];

    // Skip header if present (check if first row contains "FieldName" or similar)
    let startIndex = 0;
    if (lines.length > 0) {
        const firstLine = lines[0].toLowerCase();
        if (firstLine.includes('fieldname') || firstLine.includes('field name')) {
            startIndex = 1;
        }
    }

    for (let i = startIndex; i < lines.length; i++) {
        const values = parseCSVLine(lines[i]);
        if (values.length >= 1 && values[0].trim()) {
            csvData.push({
                fieldName: sanitizeFieldName(values[0]),
                dataType: values[1] ? values[1].trim() : 'Text',
                displayName: values[2] ? values[2].trim() : values[0].trim(),
                description: values[3] ? values[3].trim() : '',
                required: values[4] ? values[4].trim().toLowerCase() === 'yes' || values[4].trim().toLowerCase() === 'true' : false
            });
        }
    }

    if (csvData.length > 0) {
        displayPreview();
        elements.createFieldsBtn.disabled = !isConnected;
    } else {
        alert('No valid field data found in CSV. Please check the format.');
        clearFile();
    }
}

function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        const nextChar = line[i + 1];

        if (char === '"') {
            if (inQuotes && nextChar === '"') {
                current += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (char === ',' && !inQuotes) {
            result.push(current);
            current = '';
        } else {
            current += char;
        }
    }

    result.push(current);
    return result;
}

function sanitizeFieldName(name) {
    // Remove spaces and special characters, ensure it starts with a letter or underscore
    return name.trim()
        .replace(/[^a-zA-Z0-9_]/g, '_')
        .replace(/^[0-9]/, '_$&')
        .substring(0, 32); // SharePoint field names max 32 chars for internal name
}

function displayPreview() {
    elements.previewBody.innerHTML = '';

    csvData.forEach((field, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${escapeHtml(field.fieldName)}</td>
            <td>${escapeHtml(field.dataType)}</td>
            <td>${escapeHtml(field.displayName)}</td>
            <td>${escapeHtml(field.description || '-')}</td>
            <td>${field.required ? 'Yes' : 'No'}</td>
        `;
        elements.previewBody.appendChild(row);
    });

    elements.fieldCount.textContent = `${csvData.length} field${csvData.length > 1 ? 's' : ''} to create`;
    elements.previewSection.classList.remove('hidden');
}

// Create Fields
async function createFields() {
    if (!isConnected || csvData.length === 0) {
        alert('Please connect to SharePoint and upload a CSV file first.');
        return;
    }

    elements.createFieldsBtn.disabled = true;
    elements.resetBtn.disabled = true;
    elements.progressSection.classList.remove('hidden');
    elements.resultsSection.classList.remove('hidden');
    elements.resultsContent.innerHTML = '';

    const siteUrl = elements.siteUrl.value.trim();
    const listName = elements.listName.value.trim();

    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < csvData.length; i++) {
        const field = csvData[i];
        updateProgress(i + 1, csvData.length);

        try {
            const response = await sendMessageToContentScript({
                action: 'createField',
                siteUrl: siteUrl,
                listName: listName,
                fieldData: field
            });

            if (response && response.success) {
                successCount++;
                addResultItem('success', `✓ Created: ${field.displayName} (${field.fieldName})`);
            } else {
                errorCount++;
                addResultItem('error', `✗ Failed: ${field.displayName} - ${response?.message || 'Unknown error'}`);
            }
        } catch (error) {
            errorCount++;
            addResultItem('error', `✗ Failed: ${field.displayName} - ${error.message}`);
        }

        // Small delay between requests
        await delay(300);
    }

    // Summary
    addResultItem('info', `\n=== Summary ===`);
    addResultItem(successCount > 0 ? 'success' : 'info', `Success: ${successCount} fields`);
    addResultItem(errorCount > 0 ? 'error' : 'info', `Failed: ${errorCount} fields`);

    elements.createFieldsBtn.disabled = false;
    elements.resetBtn.disabled = false;
}

function updateProgress(current, total) {
    const percentage = (current / total) * 100;
    elements.progressFill.style.width = percentage + '%';
    elements.progressText.textContent = `${current} / ${total} (${Math.round(percentage)}%)`;
}

function addResultItem(type, message) {
    const item = document.createElement('div');
    item.className = 'result-item ' + type;
    item.textContent = message;
    elements.resultsContent.appendChild(item);
    elements.resultsContent.scrollTop = elements.resultsContent.scrollHeight;
}

function resetForm() {
    csvData = [];
    clearFile();
    elements.progressSection.classList.add('hidden');
    elements.resultsSection.classList.add('hidden');
    elements.progressFill.style.width = '0%';
    elements.resultsContent.innerHTML = '';
}

// Communication Functions
function sendMessageToContentScript(message) {
    return new Promise((resolve, reject) => {
        chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
            if (tabs[0]) {
                chrome.tabs.sendMessage(tabs[0].id, message, (response) => {
                    if (chrome.runtime.lastError) {
                        reject(new Error(chrome.runtime.lastError.message));
                    } else {
                        resolve(response);
                    }
                });
            } else {
                reject(new Error('No active tab found'));
            }
        });
    });
}

// Utility Functions
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Screen Navigation Functions
function openMainForm() {
    if (elements.welcomeScreen && elements.mainForm) {
        elements.welcomeScreen.classList.add('hidden');
        elements.mainForm.classList.remove('hidden');
    }
}

function closeMainForm() {
    if (elements.welcomeScreen && elements.mainForm) {
        elements.mainForm.classList.add('hidden');
        elements.welcomeScreen.classList.remove('hidden');
    }
}

function openCheckUncheckView() {
    if (elements.welcomeScreen && elements.checkUncheckView) {
        elements.welcomeScreen.classList.add('hidden');
        elements.checkUncheckView.classList.remove('hidden');
    }
}

function closeCheckUncheckView() {
    if (elements.welcomeScreen && elements.checkUncheckView) {
        elements.checkUncheckView.classList.add('hidden');
        elements.welcomeScreen.classList.remove('hidden');
    }
}

// Checkbox Functions
async function checkAllCheckboxes() {
    try {
        const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tabs[0]) {
            showCheckboxResult('error', 'No active tab found');
            return;
        }

        const results = await chrome.scripting.executeScript({
            target: { tabId: tabs[0].id },
            func: () => {
                const checkboxes = document.querySelectorAll('input[type="checkbox"]');
                let checkedCount = 0;
                checkboxes.forEach(checkbox => {
                    if (!checkbox.disabled) {
                        checkbox.checked = true;
                        checkbox.dispatchEvent(new Event('change', { bubbles: true }));
                        checkedCount++;
                    }
                });
                return { checkedCount, totalCheckboxes: checkboxes.length };
            }
        });

        if (results && results[0] && results[0].result) {
            const { checkedCount, totalCheckboxes } = results[0].result;
            showCheckboxResult('success', `Checked ${checkedCount} of ${totalCheckboxes} checkboxes`);
        } else {
            showCheckboxResult('error', 'Failed to check checkboxes');
        }
    } catch (error) {
        showCheckboxResult('error', 'Error: ' + error.message);
    }
}

async function uncheckAllCheckboxes() {
    try {
        const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tabs[0]) {
            showCheckboxResult('error', 'No active tab found');
            return;
        }

        const results = await chrome.scripting.executeScript({
            target: { tabId: tabs[0].id },
            func: () => {
                const checkboxes = document.querySelectorAll('input[type="checkbox"]');
                let uncheckedCount = 0;
                checkboxes.forEach(checkbox => {
                    if (!checkbox.disabled) {
                        checkbox.checked = false;
                        checkbox.dispatchEvent(new Event('change', { bubbles: true }));
                        uncheckedCount++;
                    }
                });
                return { uncheckedCount, totalCheckboxes: checkboxes.length };
            }
        });

        if (results && results[0] && results[0].result) {
            const { uncheckedCount, totalCheckboxes } = results[0].result;
            showCheckboxResult('success', `Unchecked ${uncheckedCount} of ${totalCheckboxes} checkboxes`);
        } else {
            showCheckboxResult('error', 'Failed to uncheck checkboxes');
        }
    } catch (error) {
        showCheckboxResult('error', 'Error: ' + error.message);
    }
}

function showCheckboxResult(type, message) {
    if (elements.checkboxResult && elements.checkboxResultText) {
        elements.checkboxResultText.textContent = message;
        elements.checkboxResult.className = 'checkbox-result ' + type;
        elements.checkboxResult.classList.remove('hidden');

        // Auto-hide after 3 seconds
        setTimeout(() => {
            elements.checkboxResult.classList.add('hidden');
        }, 3000);
    }
}
