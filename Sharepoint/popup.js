// Global state
let csvData = [];
let isConnected = false;
let currentTabId = null;
let filterAdditionalConditions = [];

// DOM Elements
const elements = {
    welcomeScreen: document.getElementById('welcomeScreen'),
    mainForm: document.getElementById('mainForm'),
    checkUncheckView: document.getElementById('checkUncheckView'),
    listCountView: document.getElementById('listCountView'),
    listDeleterView: document.getElementById('listDeleterView'),
    listFilterView: document.getElementById('listFilterView'),
    openFormBtn: document.getElementById('openFormBtn'),
    checkUncheckBtn: document.getElementById('checkUncheckBtn'),
    listCountBtn: document.getElementById('listCountBtn'),
    listDeleterBtn: document.getElementById('listDeleterBtn'),
    listFilterBtn: document.getElementById('listFilterBtn'),
    backBtn: document.getElementById('backBtn'),
    backToWelcomeBtn: document.getElementById('backToWelcomeBtn'),
    backToWelcomeFromListCountBtn: document.getElementById('backToWelcomeFromListCountBtn'),
    backToWelcomeFromListDeleterBtn: document.getElementById('backToWelcomeFromListDeleterBtn'),
    backToWelcomeFromListFilterBtn: document.getElementById('backToWelcomeFromListFilterBtn'),
    // List Filter elements
    lfGetLists: document.getElementById('lf-get-lists'),
    lfListDropdown: document.getElementById('lf-list-dropdown'),
    lfGetColumns: document.getElementById('lf-get-columns'),
    lfColumnDropdown: document.getElementById('lf-column-dropdown'),
    lfColumnTypeGroup: document.getElementById('lf-column-type-group'),
    lfColumnType: document.getElementById('lf-column-type'),
    lfOperator: document.getElementById('lf-operator'),
    lfValue: document.getElementById('lf-value'),
    lfAddFilter: document.getElementById('lf-add-filter'),
    lfAdditionalFilters: document.getElementById('lf-additional-filters'),
    lfRowLimit: document.getElementById('lf-row-limit'),
    lfApplyFilter: document.getElementById('lf-apply-filter'),
    lfClearFilter: document.getElementById('lf-clear-filter'),
    lfResultsSection: document.getElementById('lf-results-section'),
    lfResultsCount: document.getElementById('lf-results-count'),
    lfResultsThead: document.getElementById('lf-results-thead'),
    lfResultsTbody: document.getElementById('lf-results-tbody'),
    checkAllBtn: document.getElementById('checkAllBtn'),
    uncheckAllBtn: document.getElementById('uncheckAllBtn'),
    checkboxResult: document.getElementById('checkboxResult'),
    checkboxResultText: document.getElementById('checkboxResultText'),
    runListCountBtn: document.getElementById('runListCountBtn'),
    listCountResult: document.getElementById('listCountResult'),
    listCountContent: document.getElementById('listCountContent'),
    spListType: document.getElementById('sp-list-type'),
    spWorkflowUrlInfo: document.getElementById('sp-workflow-url-info'),
    spWorkflowUrlDisplay: document.getElementById('sp-workflow-url-display'),
    spGetLists: document.getElementById('sp-get-lists'),
    spListName: document.getElementById('sp-list-name'),
    spListDropdown: document.getElementById('sp-list-dropdown'),
    spWfAll: document.getElementById('sp-wf-all'),
    spWfCompleted: document.getElementById('sp-wf-completed'),
    spWfPending: document.getElementById('sp-wf-pending'),
    spWfCancelled: document.getElementById('sp-wf-cancelled'),
    spOperationMode: document.getElementById('sp-operation-mode'),
    spToggleBuilder: document.getElementById('sp-toggle-builder'),
    spCamlBuilder: document.getElementById('sp-caml-builder'),
    spBuilderField: document.getElementById('sp-builder-field'),
    spBuilderOperator: document.getElementById('sp-builder-operator'),
    spBuilderType: document.getElementById('sp-builder-type'),
    spBuilderValue: document.getElementById('sp-builder-value'),
    spBuildCaml: document.getElementById('sp-build-caml'),
    spCamlQuery: document.getElementById('sp-caml-query'),
    spRowLimit: document.getElementById('sp-row-limit'),
    spFolderPath: document.getElementById('sp-folder-path'),
    spConfirmDelete: document.getElementById('sp-confirm-delete'),
    spExecuteBtn: document.getElementById('sp-execute-btn'),
    spPreviewBtn: document.getElementById('sp-preview-btn'),
    spOutput: document.getElementById('sp-output'),
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

    // List Count Aggregator button
    if (elements.listCountBtn) {
        elements.listCountBtn.addEventListener('click', openListCountView);
    }

    // Back to Welcome button (from List Count view)
    if (elements.backToWelcomeFromListCountBtn) {
        elements.backToWelcomeFromListCountBtn.addEventListener('click', closeListCountView);
    }

    // Run List Count button
    if (elements.runListCountBtn) {
        elements.runListCountBtn.addEventListener('click', runListCountAggregation);
    }

    // List Deleter button
    if (elements.listDeleterBtn) {
        elements.listDeleterBtn.addEventListener('click', openListDeleterView);
    }

    // Back to Welcome button (from List Deleter view)
    if (elements.backToWelcomeFromListDeleterBtn) {
        elements.backToWelcomeFromListDeleterBtn.addEventListener('click', closeListDeleterView);
    }

    // List Type selection
    if (elements.spListType) {
        elements.spListType.addEventListener('change', handleListTypeChange);
    }

    // Get All Lists button
    if (elements.spGetLists) {
        elements.spGetLists.addEventListener('click', getAllLists);
    }

    // WorkflowTasks Quick Actions
    if (elements.spWfAll) {
        elements.spWfAll.addEventListener('click', () => setupWorkflowTaskList('deleteItems', 'all'));
    }
    if (elements.spWfCompleted) {
        elements.spWfCompleted.addEventListener('click', () => setupWorkflowTaskList('deleteItems', 'completed'));
    }
    if (elements.spWfPending) {
        elements.spWfPending.addEventListener('click', () => setupWorkflowTaskList('deleteItems', 'pending'));
    }
    if (elements.spWfCancelled) {
        elements.spWfCancelled.addEventListener('click', () => setupWorkflowTaskList('deleteItems', 'cancelled'));
    }

    // Operation Mode change
    if (elements.spOperationMode) {
        elements.spOperationMode.addEventListener('change', handleOperationModeChange);
    }

    // CAML Builder toggle
    if (elements.spToggleBuilder) {
        elements.spToggleBuilder.addEventListener('click', toggleCamlBuilder);
    }

    // Generate CAML from builder
    if (elements.spBuildCaml) {
        elements.spBuildCaml.addEventListener('click', buildCamlQuery);
    }

    // Execute button
    if (elements.spExecuteBtn) {
        elements.spExecuteBtn.addEventListener('click', executeListDeleterOperation);
    }

    // Preview button
    if (elements.spPreviewBtn) {
        elements.spPreviewBtn.addEventListener('click', previewListDeleterItems);
    }

    // List Filter button
    if (elements.listFilterBtn) {
        elements.listFilterBtn.addEventListener('click', openListFilterView);
    }

    // Back to Welcome button (from List Filter view)
    if (elements.backToWelcomeFromListFilterBtn) {
        elements.backToWelcomeFromListFilterBtn.addEventListener('click', closeListFilterView);
    }

    // List Filter - Get Lists
    if (elements.lfGetLists) {
        elements.lfGetLists.addEventListener('click', handleLfGetLists);
    }

    // List Filter - Get Columns
    if (elements.lfGetColumns) {
        elements.lfGetColumns.addEventListener('click', handleLfGetColumns);
    }

    // List Filter - Add Additional Filter
    if (elements.lfAddFilter) {
        elements.lfAddFilter.addEventListener('click', addAdditionalFilter);
    }

    // List Filter - Apply Filter
    if (elements.lfApplyFilter) {
        elements.lfApplyFilter.addEventListener('click', applyListFilter);
    }

    // List Filter - Clear Filter
    if (elements.lfClearFilter) {
        elements.lfClearFilter.addEventListener('click', clearListFilter);
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

function openListCountView() {
    if (elements.welcomeScreen && elements.listCountView) {
        elements.welcomeScreen.classList.add('hidden');
        elements.listCountView.classList.remove('hidden');
    }
}

function closeListCountView() {
    if (elements.welcomeScreen && elements.listCountView) {
        elements.listCountView.classList.add('hidden');
        elements.welcomeScreen.classList.remove('hidden');
    }
}

function openListDeleterView() {
    if (elements.welcomeScreen && elements.listDeleterView) {
        elements.welcomeScreen.classList.add('hidden');
        elements.listDeleterView.classList.remove('hidden');
    }
}

function closeListDeleterView() {
    if (elements.welcomeScreen && elements.listDeleterView) {
        elements.listDeleterView.classList.add('hidden');
        elements.welcomeScreen.classList.remove('hidden');
    }
}

function openListFilterView() {
    if (elements.welcomeScreen && elements.listFilterView) {
        elements.welcomeScreen.classList.add('hidden');
        elements.listFilterView.classList.remove('hidden');
    }
}

function closeListFilterView() {
    if (elements.welcomeScreen && elements.listFilterView) {
        elements.listFilterView.classList.add('hidden');
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

// List Count Aggregator Functions
async function runListCountAggregation() {
    try {
        const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tabs[0]) {
            showListCountError('No active tab found');
            return;
        }

        // Check if current page is a SharePoint page
        const currentUrl = tabs[0].url;
        const isSharePointPage = currentUrl.includes('/sites/') ||
                                currentUrl.includes('/Lists/') ||
                                currentUrl.includes('/_layouts/') ||
                                currentUrl.includes('.aspx');

        if (!isSharePointPage) {
            showListCountError(`
                <p style="margin: 0 0 10px 0; font-weight: bold;">Not on a SharePoint Page</p>
                <p style="margin: 0;">Please navigate to a SharePoint page (e.g., a site home, list view, or any .aspx page) before running the analysis.</p>
                <p style="margin: 10px 0 0 0; font-size: 11px; color: #666;">Current URL: ${currentUrl.substring(0, 60)}${currentUrl.length > 60 ? '...' : ''}</p>
            `);
            return;
        }

        // Show loading state
        elements.listCountResult.classList.remove('hidden');
        elements.listCountContent.innerHTML = '<p style="text-align: center; color: #667eea;">Running analysis... Please wait.</p>';

        const results = await chrome.scripting.executeScript({
            target: { tabId: tabs[0].id },
            func: () => {
                // Configuration
                const MIN_ITEM_COUNT_THRESHOLD = 3500;

                // Helper function to encode HTML
                const htmlEncode = (str) => {
                    if (str === null || str === undefined) return '';
                    return String(str).replace(/[&<>"']/g, function(s) {
                        return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[s];
                    });
                };

                // Function to get request digest
                const getRequestDigest = async (siteUrl) => {
                    try {
                        const response = await fetch(siteUrl + "/_api/contextinfo", {
                            method: "POST",
                            headers: { "Accept": "application/json; odata=verbose" }
                        });
                        const data = await response.json();
                        return data.d.GetContextWebInformation.FormDigestValue;
                    } catch (error) {
                        console.warn(`Could not get digest for ${siteUrl}`, error);
                        return null;
                    }
                };

                // Function to recursively find all subsites
                async function getSubsiteUrls(rootWebUrl) {
                    const subsiteUrls = [rootWebUrl];
                    let queue = [rootWebUrl];

                    while (queue.length > 0) {
                        const currentUrl = queue.shift();

                        try {
                            const restUrl = `${currentUrl}/_api/web/webs?$select=Title,Url`;
                            const response = await fetch(restUrl, {
                                method: "GET",
                                headers: { "Accept": "application/json; odata=verbose" }
                            });
                            const data = await response.json();
                            const webs = data.d.results;

                            webs.forEach(web => {
                                const subsiteUrl = web.Url.replace(/\/$/, '');
                                if (!subsiteUrls.includes(subsiteUrl)) {
                                    subsiteUrls.push(subsiteUrl);
                                    queue.push(subsiteUrl);
                                }
                            });
                        } catch (error) {
                            console.warn(`Could not access webs for URL: ${currentUrl}. Skipping.`, error);
                        }
                    }
                    return subsiteUrls;
                }

                // Function to get filtered lists for a single site
                async function getFilteredListsForSite(siteUrl) {
                    const digest = await getRequestDigest(siteUrl);
                    if (!digest) {
                        return [{
                            siteUrl: siteUrl,
                            title: 'ERROR FETCHING LISTS',
                            count: 'Digest Error/No Access',
                            type: ''
                        }];
                    }

                    try {
                        const restUrl = `${siteUrl}/_api/web/lists?$select=Title,ItemCount,BaseType`;
                        const response = await fetch(restUrl, {
                            method: "GET",
                            headers: {
                                "Accept": "application/json; odata=verbose",
                                "X-RequestDigest": digest
                            }
                        });
                        const data = await response.json();
                        const allLists = data.d.results;

                        return allLists
                            .filter(list => list.ItemCount > MIN_ITEM_COUNT_THRESHOLD)
                            .map(list => ({
                                siteUrl: siteUrl,
                                title: list.Title,
                                count: list.ItemCount,
                                type: list.BaseType === 1 ? 'Library' : 'List'
                            }));

                    } catch (error) {
                        console.error(`Error fetching lists for ${siteUrl}`, error);
                        return [{
                            siteUrl: siteUrl,
                            title: 'ERROR FETCHING LISTS',
                            count: `Error (${error.status || 'Unknown'})`,
                            type: ''
                        }];
                    }
                }

                // Main aggregation logic
                async function runAggregation() {
                    let rootWebUrl = '';

                    // Try to get SharePoint context first
                    if (typeof _spPageContextInfo !== 'undefined') {
                        rootWebUrl = _spPageContextInfo.webAbsoluteUrl.replace(/\/$/, '');
                    } else {
                        // Fall back to using the current page URL
                        const currentUrl = window.location.href;
                        // Check if it looks like a SharePoint URL
                        if (currentUrl.includes('/sites/') || currentUrl.includes('/Lists/') || currentUrl.includes('/_layouts/')) {
                            // Extract the base site URL
                            const urlObj = new URL(currentUrl);
                            // Try to find the site base URL
                            const sitesMatch = currentUrl.match(/(https?:\/\/[^\/]+\/sites\/[^\/]+)/);
                            if (sitesMatch) {
                                rootWebUrl = sitesMatch[1];
                            } else {
                                // Use origin as fallback
                                rootWebUrl = urlObj.origin;
                            }
                        } else {
                            return {
                                error: 'SharePoint context not found. Please navigate to a SharePoint page.',
                                html: ''
                            };
                        }
                    }

                    if (!rootWebUrl) {
                        return {
                            error: 'Could not determine SharePoint site URL. Please navigate to a SharePoint page.',
                            html: ''
                        };
                    }

                    // Get all subsite URLs recursively
                    const allSiteUrls = await getSubsiteUrls(rootWebUrl);

                    // Fetch filtered lists for all sites concurrently
                    const fetchPromises = allSiteUrls.map(siteUrl => getFilteredListsForSite(siteUrl));
                    const resultsArray = await Promise.all(fetchPromises);

                    // Flatten and filter out errors
                    const allFilteredLists = resultsArray.flat().filter(res => res.title !== 'ERROR FETCHING LISTS');

                    // Build HTML output
                    let htmlOutput = `
                        <div style="margin-bottom: 15px; font-weight: bold;">
                            Found ${allFilteredLists.length} lists/libraries with >${MIN_ITEM_COUNT_THRESHOLD} items.
                        </div>
                        <div style="border: 1px solid #ccc; padding: 10px; background-color: #f9f9f9; max-height: 400px; overflow-y: auto;">
                    `;

                    if (allFilteredLists.length === 0) {
                        htmlOutput += `<p style="color: green; font-weight: bold;">No lists found with more than ${MIN_ITEM_COUNT_THRESHOLD} items across all accessible subsites.</p>`;
                    } else {
                        htmlOutput += `
                            <table style="border-collapse: collapse; width: 100%; font-size: 12px;">
                                <thead>
                                    <tr style="background-color: #eee; position: sticky; top: 0;">
                                        <th style="border: 1px solid #ccc; padding: 8px; text-align: left; width: 45%;">Site URL</th>
                                        <th style="border: 1px solid #ccc; padding: 8px; text-align: left; width: 35%;">List Name (Type)</th>
                                        <th style="border: 1px solid #ccc; padding: 8px; text-align: right; width: 20%;">Item Count</th>
                                    </tr>
                                </thead>
                                <tbody>
                        `;

                        allFilteredLists.forEach(res => {
                            const safeUrl = htmlEncode(res.siteUrl);
                            const safeTitle = htmlEncode(res.title);
                            const safeCount = htmlEncode(res.count);
                            const typeInfo = ` (${res.type})`;

                            htmlOutput += `
                                <tr>
                                    <td style="border: 1px solid #ccc; padding: 8px;">
                                        <a href="${res.siteUrl}" target="_blank" style="color: #667eea; text-decoration: none;">${safeUrl}</a>
                                    </td>
                                    <td style="border: 1px solid #ccc; padding: 8px;">${safeTitle} ${typeInfo}</td>
                                    <td style="border: 1px solid #ccc; padding: 8px; text-align: right; font-weight: bold; color: #CC0000;">${safeCount}</td>
                                </tr>
                            `;
                        });
                        htmlOutput += '</tbody></table>';
                    }

                    htmlOutput += '</div>';

                    return {
                        error: null,
                        html: htmlOutput,
                        totalSites: allSiteUrls.length
                    };
                }

                return runAggregation();
            }
        });

        if (results && results[0] && results[0].result) {
            const { error, html, totalSites } = results[0].result;

            if (error) {
                showListCountError(error);
            } else {
                showListCountSuccess(html, totalSites);
            }
        } else {
            showListCountError('Failed to run analysis');
        }
    } catch (error) {
        showListCountError('Error: ' + error.message);
    }
}

function showListCountError(message) {
    if (elements.listCountResult && elements.listCountContent) {
        elements.listCountContent.innerHTML = `
            <div style="padding: 15px; background: #ffebee; border: 1px solid #ef9a9a; border-radius: 6px; color: #c62828;">
                <p style="margin: 0; font-weight: bold;">Error</p>
                <p style="margin: 5px 0 0 0;">${message}</p>
            </div>
        `;
    }
}

function showListCountSuccess(html, totalSites) {
    if (elements.listCountResult && elements.listCountContent) {
        elements.listCountContent.innerHTML = `
            <div style="margin-bottom: 10px; padding: 10px; background: #e8f5e9; border: 1px solid #a5d6a7; border-radius: 6px; color: #2e7d32;">
                <p style="margin: 0; font-weight: bold;">Analysis Complete</p>
                <p style="margin: 5px 0 0 0;">Scanned ${totalSites} sites successfully.</p>
            </div>
            ${html}
        `;
    }
}

// List Deleter Functions
function showDeleterOutput(message, type = 'info') {
    if (elements.spOutput) {
        elements.spOutput.style.display = 'block';
        elements.spOutput.className = 'sp-deleter-output sp-deleter-output-' + type;
        elements.spOutput.textContent = message;
    }
}

function handleListTypeChange(e) {
    if (e.target.value === 'workflow') {
        elements.spListName.value = 'WorkflowTasks';
        const workflowUrl = 'WorkflowTasks';
        elements.spWorkflowUrlDisplay.textContent = workflowUrl;
        elements.spWorkflowUrlInfo.style.display = 'block';
        showDeleterOutput('Workflow List selected. Using /WorkflowTasks list.', 'info');
    } else {
        elements.spListName.value = '';
        elements.spWorkflowUrlInfo.style.display = 'none';
    }
}

async function getAllLists() {
    try {
        showDeleterOutput('Fetching all lists from site...', 'info');
        const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tabs[0]) {
            showDeleterOutput('No active tab found', 'error');
            return;
        }

        // Check if on SharePoint page
        const currentUrl = tabs[0].url;
        const isSharePointPage = currentUrl.includes('/sites/') ||
                                currentUrl.includes('/Lists/') ||
                                currentUrl.includes('/_layouts/') ||
                                currentUrl.includes('.aspx');

        if (!isSharePointPage) {
            showDeleterOutput('❌ Not on a SharePoint page.\n\nPlease navigate to a SharePoint site first (e.g., /sites/yoursite or a list page).', 'error');
            return;
        }

        showDeleterOutput('Verifying content script is loaded...', 'info');

        // First verify content script is loaded
        try {
            await sendMessageToContentScript({
                action: 'ping'
            });
        } catch (pingError) {
            showDeleterOutput('❌ Content script not available.\n\nPlease refresh the SharePoint page and try again.', 'error');
            console.error('Content script ping error:', pingError);
            return;
        }

        showDeleterOutput('Fetching lists using SharePoint JSOM...', 'info');

        // Send message to content script
        const response = await sendMessageToContentScript({
            action: 'getAllLists'
        });

        if (response && response.success) {
            const lists = response.lists;
            elements.spListDropdown.innerHTML = '<option value="">-- Select a list --</option>';
            lists.forEach(list => {
                const option = document.createElement('option');
                option.value = list.title;
                option.textContent = `${list.title} (${list.itemCount} items)`;
                elements.spListDropdown.appendChild(option);
            });

            elements.spListDropdown.style.display = 'block';
            showDeleterOutput(`✓ Found ${lists.length} list(s). Select from dropdown.`, 'success');

            // Handle dropdown selection
            elements.spListDropdown.onchange = () => {
                elements.spListName.value = elements.spListDropdown.value;
            };
        } else {
            showDeleterOutput(`❌ Error: ${response?.message || 'Failed to fetch lists'}`, 'error');
        }
    } catch (error) {
        console.error('getAllLists error:', error);
        const errorMsg = error.message.includes('Receiving end does not exist') 
            ? '❌ Content script not loaded.\n\nPlease refresh the SharePoint page (F5) and try again.' 
            : `❌ Error: ${error.message}`;
        showDeleterOutput(errorMsg, 'error');
    }
}

function setupWorkflowTaskList(mode = 'deleteItems', filterType = 'all') {
    elements.spListName.value = 'WorkflowTasks';
    elements.spOperationMode.value = mode;

    let camlQuery = '';
    switch(filterType) {
        case 'completed':
            camlQuery = `<View>
  <Query>
    <Where>
      <Eq>
        <FieldRef Name='Status'/>
        <Value Type='Text'>Completed</Value>
      </Eq>
    </Where>
  </Query>
</View>`;
            break;
        case 'pending':
            camlQuery = `<View>
  <Query>
    <Where>
      <Neq>
        <FieldRef Name='Status'/>
        <Value Type='Text'>Completed</Value>
      </Neq>
    </Where>
  </Query>
</View>`;
            break;
        case 'cancelled':
            camlQuery = `<View>
  <Query>
    <Where>
      <Eq>
        <FieldRef Name='Status'/>
        <Value Type='Text'>Cancelled</Value>
      </Eq>
    </Where>
  </Query>
</View>`;
            break;
        default:
            camlQuery = `<View></View>`;
    }
    elements.spCamlQuery.value = camlQuery;

    // Trigger change event to update button text
    elements.spOperationMode.dispatchEvent(new Event('change'));

    showDeleterOutput(`Setup: WorkflowTasks - ${filterType.charAt(0).toUpperCase() + filterType.slice(1)} filter applied. Check checkbox and click Execute.`, 'info');
}

function handleOperationModeChange(e) {
    const btn = elements.spExecuteBtn;
    if (e.target.value === 'deleteList') {
        btn.textContent = 'Delete Entire List';
        btn.className = 'btn btn-danger';
    } else if (e.target.value === 'deleteItems') {
        btn.textContent = 'Delete Filtered Items';
        btn.className = 'btn btn-danger';
    } else {
        btn.textContent = 'Preview Only';
        btn.className = 'btn btn-secondary';
    }
}

function toggleCamlBuilder() {
    if (elements.spCamlBuilder.style.display === 'none') {
        elements.spCamlBuilder.style.display = 'block';
        elements.spToggleBuilder.textContent = 'Hide Builder';
    } else {
        elements.spCamlBuilder.style.display = 'none';
        elements.spToggleBuilder.textContent = 'Show Builder';
    }
}

function buildCamlQuery() {
    const field = elements.spBuilderField.value;
    const operator = elements.spBuilderOperator.value;
    const valueType = elements.spBuilderType.value;
    const value = elements.spBuilderValue.value.trim();

    if (!field) {
        showDeleterOutput('Please select a field', 'error');
        return;
    }

    let camlXml = `<View>\n  <Query>\n    <Where>\n      `;

    if (operator === 'IsNull') {
        camlXml += `<IsNull>\n        <FieldRef Name='${field}'/>\n      </IsNull>`;
    } else if (operator === 'IsNotNull') {
        camlXml += `<IsNotNull>\n        <FieldRef Name='${field}'/>\n      </IsNotNull>`;
    } else if (operator === 'Contains') {
        camlXml += `<Contains>\n        <FieldRef Name='${field}'/>\n        <Value Type='${valueType}'>${value}</Value>\n      </Contains>`;
    } else if (operator === 'BeginsWith') {
        camlXml += `<BeginsWith>\n        <FieldRef Name='${field}'/>\n        <Value Type='${valueType}'>${value}</Value>\n      </BeginsWith>`;
    } else {
        camlXml += `<${operator}>\n        <FieldRef Name='${field}'/>\n        <Value Type='${valueType}'>${value}</Value>\n      </${operator}>`;
    }

    camlXml += `\n    </Where>\n  </Query>\n</View>`;

    elements.spCamlQuery.value = camlXml;
    showDeleterOutput('CAML query generated!', 'success');
}

async function previewListDeleterItems() {
    const listTitle = elements.spListName.value.trim();
    const folderPath = elements.spFolderPath.value.trim();

    if (!listTitle) {
        showDeleterOutput('❌ Please enter a list name', 'error');
        return;
    }

    try {
        showDeleterOutput('Fetching items...', 'info');
        const camlQuery = getCAMLQuery();

        // Send message to content script
        const response = await sendMessageToContentScript({
            action: 'previewListDeleterItems',
            listTitle: listTitle,
            camlQuery: camlQuery,
            folderPath: folderPath
        });

        if (response && response.success) {
            const items = response.items;
            if (items.length === 0) {
                showDeleterOutput('No items found matching the query.', 'info');
            } else {
                const preview = items.map(item => {
                    const type = item.isFolder ? '[FOLDER]' : '[ITEM]';
                    return `${type} ID:${item.id} | Title:${item.title}`;
                }).join('\n');

                showDeleterOutput(`✓ Found ${items.length} item(s):\n${preview}`, 'success');
            }
        } else {
            showDeleterOutput(`❌ Error: ${response?.message || 'Failed to preview items'}`, 'error');
        }
    } catch (error) {
        console.error('previewListDeleterItems error:', error);
        const errorMsg = error.message.includes('Receiving end does not exist') 
            ? '❌ Content script not loaded.\n\nPlease refresh the SharePoint page (F5) and try again.' 
            : `❌ Error: ${error.message}`;
        showDeleterOutput(errorMsg, 'error');
    }
}

async function executeListDeleterOperation() {
    const listTitle = elements.spListName.value.trim();
    const mode = elements.spOperationMode.value;
    const confirmed = elements.spConfirmDelete.checked;
    const folderPath = elements.spFolderPath.value.trim();

    if (!listTitle) {
        showDeleterOutput('❌ Please enter a list name', 'error');
        return;
    }

    if (!confirmed) {
        showDeleterOutput('❌ Please confirm the action by checking the checkbox', 'error');
        return;
    }

    try {
        if (mode === 'deleteList') {
            // Show confirmation dialog
            const confirmDelete = confirm(`⚠️ WARNING!\n\nYou are about to PERMANENTLY DELETE the entire list:\n"${listTitle}"\n\nThis action CANNOT be undone. All items and data will be lost.\n\nAre you absolutely sure? Click OK to delete, or Cancel to abort.`);
            
            if (!confirmDelete) {
                showDeleterOutput('❌ Deletion cancelled by user', 'info');
                return;
            }

            showDeleterOutput('Deleting list...', 'info');

            const response = await sendMessageToContentScript({
                action: 'deleteList',
                listTitle: listTitle
            });

            if (response && response.success) {
                showDeleterOutput(`✓ List "${listTitle}" deleted successfully!`, 'success');
            } else {
                showDeleterOutput(`❌ Error: ${response?.message || 'Failed to delete list'}`, 'error');
            }
        } else if (mode === 'deleteItems') {
            // Show confirmation dialog for items
            const confirmDelete = confirm(`⚠️ WARNING!\n\nYou are about to PERMANENTLY DELETE items from list:\n"${listTitle}"\n\nThis action CANNOT be undone.\n\nAre you sure? Click OK to delete matching items, or Cancel to abort.`);
            
            if (!confirmDelete) {
                showDeleterOutput('❌ Deletion cancelled by user', 'info');
                return;
            }

            showDeleterOutput('Fetching items to delete...', 'info');
            const camlQuery = getCAMLQuery();

            const response = await sendMessageToContentScript({
                action: 'deleteListItems',
                listTitle: listTitle,
                camlQuery: camlQuery,
                folderPath: folderPath
            });

            if (response && response.success) {
                const deleted = response.deleted;
                showDeleterOutput(`✓ Successfully deleted ${deleted} item(s) from list "${listTitle}"`, 'success');
            } else {
                showDeleterOutput(`❌ Error: ${response?.message || 'Failed to delete items'}`, 'error');
            }
        } else {
            // Preview mode
            await previewListDeleterItems();
        }
    } catch (error) {
        console.error('executeListDeleterOperation error:', error);
        const errorMsg = error.message.includes('Receiving end does not exist') 
            ? '❌ Content script not loaded.\n\nPlease refresh the SharePoint page (F5) and try again.' 
            : `❌ Error: ${error.message}`;
        showDeleterOutput(errorMsg, 'error');
    }
}

function getCAMLQuery() {
    let camlQuery = elements.spCamlQuery.value.trim();
    const rowLimit = elements.spRowLimit.value;

    if (!camlQuery) {
        return `<View><RowLimit>${rowLimit}</RowLimit></View>`;
    }

    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(camlQuery, 'text/xml');

    let viewNode = xmlDoc.getElementsByTagName('View')[0];
    if (!viewNode) {
        const root = xmlDoc.documentElement;
        if (root.tagName === 'View') {
            viewNode = root;
        } else {
            viewNode = xmlDoc.createElement('View');
            const queryNode = xmlDoc.getElementsByTagName('Query')[0];
            if (queryNode) {
                viewNode.appendChild(queryNode.cloneNode(true));
            }
            xmlDoc.appendChild(viewNode);
        }
    }

    if (!viewNode.getElementsByTagName('RowLimit').length) {
        const rowLimitElem = xmlDoc.createElement('RowLimit');
        rowLimitElem.textContent = rowLimit;
        viewNode.appendChild(rowLimitElem);
    }

    const serializer = new XMLSerializer();
    return serializer.serializeToString(viewNode);
}

// ============================================================================
// List Filter Functions
// ============================================================================

let lfAvailableColumns = [];
let lfAvailableLists = [];

// Handle Get Lists for List Filter
async function handleLfGetLists() {
    try {
        showLfOutput('Loading lists...', 'info');

        const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tabs[0]) {
            showLfOutput('No active tab found', 'error');
            return;
        }

        const response = await chrome.tabs.sendMessage(tabs[0].id, {
            action: 'getAllLists'
        });

        if (response && response.success) {
            lfAvailableLists = response.lists || [];
            populateLfListDropdown(lfAvailableLists);
            showLfOutput(`Found ${lfAvailableLists.length} lists`, 'success');
        } else {
            showLfOutput(`Failed to get lists: ${response?.message || 'Unknown error'}`, 'error');
        }
    } catch (error) {
        console.error('handleLfGetLists error:', error);
        const errorMsg = error.message.includes('Receiving end does not exist')
            ? 'Content script not loaded.\n\nPlease refresh the SharePoint page (F5) and try again.'
            : `Error: ${error.message}`;
        showLfOutput(errorMsg, 'error');
    }
}

function populateLfListDropdown(lists) {
    elements.lfListDropdown.innerHTML = '<option value="">-- Select a list --</option>';
    lists.forEach(list => {
        const option = document.createElement('option');
        option.value = list.title;
        option.textContent = list.title;
        elements.lfListDropdown.appendChild(option);
    });
}

// Handle Get Columns
async function handleLfGetColumns() {
    const listTitle = elements.lfListDropdown.value;
    if (!listTitle) {
        showLfOutput('Please select a list first', 'error');
        return;
    }

    try {
        showLfOutput('Loading columns...', 'info');

        const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tabs[0]) {
            showLfOutput('No active tab found', 'error');
            return;
        }

        const response = await chrome.tabs.sendMessage(tabs[0].id, {
            action: 'getListFields',
            listTitle: listTitle
        });

        if (response && response.success) {
            lfAvailableColumns = response.fields || [];
            populateLfColumnDropdown(lfAvailableColumns);
            showLfOutput(`Found ${lfAvailableColumns.length} columns`, 'success');
        } else {
            showLfOutput(`Failed to get columns: ${response?.message || 'Unknown error'}`, 'error');
        }
    } catch (error) {
        console.error('handleLfGetColumns error:', error);
        const errorMsg = error.message.includes('Receiving end does not exist')
            ? 'Content script not loaded.\n\nPlease refresh the SharePoint page (F5) and try again.'
            : `Error: ${error.message}`;
        showLfOutput(errorMsg, 'error');
    }
}

function populateLfColumnDropdown(fields) {
    elements.lfColumnDropdown.innerHTML = '<option value="">-- Select a column --</option>';
    fields.forEach(field => {
        const option = document.createElement('option');
        option.value = field.internalName;
        option.textContent = `${field.title} (${field.type})`;
        option.dataset.type = field.type;
        elements.lfColumnDropdown.appendChild(option);
    });
    elements.lfColumnDropdown.disabled = false;
}

// Handle column selection change
elements.lfColumnDropdown.addEventListener('change', function() {
    const selectedOption = this.options[this.selectedIndex];
    if (selectedOption.value) {
        elements.lfColumnTypeGroup.style.display = 'block';
        elements.lfColumnType.value = selectedOption.dataset.type;
    } else {
        elements.lfColumnTypeGroup.style.display = 'none';
    }
});

// Add Additional Filter
function addAdditionalFilter() {
    const container = elements.lfAdditionalFilters;

    // Clear placeholder if it exists
    if (container.querySelector('p')) {
        container.innerHTML = '';
    }

    const filterIndex = filterAdditionalConditions.length;
    filterAdditionalConditions.push({ id: Date.now() });

    const filterDiv = document.createElement('div');
    filterDiv.className = 'lf-additional-filter-row';
    filterDiv.dataset.index = filterIndex;
    filterDiv.style.cssText = 'display:flex;gap:5px;margin-bottom:5px;align-items:center;';

    filterDiv.innerHTML = `
        <select class="lf-additional-field sp-deleter-select" style="flex:1;">
            <option value="">Select Field...</option>
            ${lfAvailableColumns.map(f => `<option value="${f.internalName}">${f.title}</option>`).join('')}
        </select>
        <select class="lf-additional-operator sp-deleter-select" style="flex:1;">
            <option value="Eq">Equals (=)</option>
            <option value="Neq">Not Equals (!=)</option>
            <option value="Contains">Contains</option>
            <option value="BeginsWith">Begins With</option>
            <option value="Lt">Less Than (&lt;)</option>
            <option value="Le">Less Than Equal (&lt;=)</option>
            <option value="Gt">Greater Than (&gt;)</option>
            <option value="Ge">Greater Than Equal (&gt;=)</option>
            <option value="IsNull">Is Null</option>
            <option value="IsNotNull">Is Not Null</option>
        </select>
        <input type="text" class="lf-additional-value sp-deleter-input" placeholder="Value..." style="flex:1;">
        <button type="button" class="lf-remove-filter btn" style="padding:5px 10px;background:#d13438;color:white;">×</button>
    `;

    container.appendChild(filterDiv);

    // Add remove handler
    filterDiv.querySelector('.lf-remove-filter').addEventListener('click', function() {
        filterDiv.remove();
        if (container.children.length === 0) {
            container.innerHTML = '<p style="color:#888;font-style:italic;font-size:12px;">No additional filters. Click "+ Add Filter" to add more conditions.</p>';
        }
    });
}

// Apply List Filter
async function applyListFilter() {
    const listTitle = elements.lfListDropdown.value;
    const columnName = elements.lfColumnDropdown.value;
    const operator = elements.lfOperator.value;
    const value = elements.lfValue.value;
    const rowLimit = parseInt(elements.lfRowLimit.value) || 0;

    if (!listTitle) {
        showLfOutput('Please select a list', 'error');
        return;
    }

    if (!columnName) {
        showLfOutput('Please select a column', 'error');
        return;
    }

    // Build CAML query
    const camlQuery = buildFilterCAML(columnName, operator, value);

    try {
        showLfOutput('Applying filter...', 'info');

        const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tabs[0]) {
            showLfOutput('No active tab found', 'error');
            return;
        }

        const response = await chrome.tabs.sendMessage(tabs[0].id, {
            action: 'filterListItems',
            listTitle: listTitle,
            camlQuery: camlQuery,
            rowLimit: rowLimit
        });

        if (response && response.success) {
            displayFilterResults(response.results);
            showLfOutput(`Found ${response.results.length} items`, 'success');
        } else {
            showLfOutput(`Failed to filter: ${response?.message || 'Unknown error'}`, 'error');
        }
    } catch (error) {
        console.error('applyListFilter error:', error);
        const errorMsg = error.message.includes('Receiving end does not exist')
            ? 'Content script not loaded.\n\nPlease refresh the SharePoint page (F5) and try again.'
            : `Error: ${error.message}`;
        showLfOutput(errorMsg, 'error');
    }
}

function buildFilterCAML(columnName, operator, value) {
    let whereClause = '';

    // Get additional filters
    const additionalFilters = [];
    document.querySelectorAll('.lf-additional-filter-row').forEach(row => {
        const field = row.querySelector('.lf-additional-field').value;
        const op = row.querySelector('.lf-additional-operator').value;
        const val = row.querySelector('.lf-additional-value').value;
        if (field && op) {
            additionalFilters.push({ field, operator: op, value: val });
        }
    });

    // Build where clause
    if (operator === 'IsNull') {
        whereClause = `<IsNull><FieldRef Name='${columnName}'/></IsNull>`;
    } else if (operator === 'IsNotNull') {
        whereClause = `<IsNotNull><FieldRef Name='${columnName}'/></IsNotNull>`;
    } else if (operator === 'Contains') {
        whereClause = `<Contains><FieldRef Name='${columnName}'/><Value Type='Text'>${escapeXml(value)}</Value></Contains>`;
    } else if (operator === 'BeginsWith') {
        whereClause = `<BeginsWith><FieldRef Name='${columnName}'/><Value Type='Text'>${escapeXml(value)}</Value></BeginsWith>`;
    } else if (operator === 'In') {
        const values = value.split(',').map(v => v.trim());
        const orParts = values.map(v => `<Eq><FieldRef Name='${columnName}'/><Value Type='Text'>${escapeXml(v)}</Value></Eq>`).join('');
        whereClause = `<Or>${orParts}</Or>`;
    } else {
        whereClause = `<${operator}><FieldRef Name='${columnName}'/><Value Type='Text'>${escapeXml(value)}</Value></${operator}>`;
    }

    // Add additional filters with And logic
    if (additionalFilters.length > 0) {
        const additionalClauses = additionalFilters.map(f => {
            if (f.operator === 'IsNull') {
                return `<IsNull><FieldRef Name='${f.field}'/></IsNull>`;
            } else if (f.operator === 'IsNotNull') {
                return `<IsNotNull><FieldRef Name='${f.field}'/></IsNotNull>`;
            } else if (f.operator === 'Contains') {
                return `<Contains><FieldRef Name='${f.field}'/><Value Type='Text'>${escapeXml(f.value)}</Value></Contains>`;
            } else if (f.operator === 'BeginsWith') {
                return `<BeginsWith><FieldRef Name='${f.field}'/><Value Type='Text'>${escapeXml(f.value)}</Value></BeginsWith>`;
            } else {
                return `<${f.operator}><FieldRef Name='${f.field}'/><Value Type='Text'>${escapeXml(f.value)}</Value></${f.operator}>`;
            }
        });

        // Combine all with And
        const allClauses = [whereClause, ...additionalClauses];
        while (allClauses.length > 1) {
            const first = allClauses.shift();
            const second = allClauses.shift();
            allClauses.unshift(`<And>${first}${second}</And>`);
        }
        whereClause = allClauses[0];
    }

    const rowLimit = elements.lfRowLimit.value || '100';

    return `<View><Query><Where>${whereClause}</Where></Query><RowLimit>${rowLimit}</RowLimit></View>`;
}

function escapeXml(str) {
    return str.replace(/&/g, '&amp;')
              .replace(/</g, '&lt;')
              .replace(/>/g, '&gt;')
              .replace(/"/g, '&quot;')
              .replace(/'/g, '&apos;');
}

function displayFilterResults(results) {
    if (!results || results.length === 0) {
        elements.lfResultsSection.style.display = 'block';
        elements.lfResultsCount.textContent = 'No items found';
        elements.lfResultsThead.innerHTML = '';
        elements.lfResultsTbody.innerHTML = '<tr><td colspan="100%">No matching items found</td></tr>';
        return;
    }

    // Get all unique columns from results
    const columns = Object.keys(results[0]);

    // Build header
    elements.lfResultsThead.innerHTML = '<tr>' +
        columns.map(col => `<th>${col}</th>`).join('') +
        '</tr>';

    // Build body
    elements.lfResultsTbody.innerHTML = results.map(item =>
        '<tr>' +
        columns.map(col => `<td>${item[col] !== null ? item[col] : ''}</td>`).join('') +
        '</tr>'
    ).join('');

    elements.lfResultsCount.textContent = `${results.length} item(s) found`;
    elements.lfResultsSection.style.display = 'block';
}

function clearListFilter() {
    elements.lfListDropdown.value = '';
    elements.lfColumnDropdown.value = '';
    elements.lfColumnDropdown.disabled = true;
    elements.lfColumnTypeGroup.style.display = 'none';
    elements.lfOperator.value = 'Eq';
    elements.lfValue.value = '';
    elements.lfRowLimit.value = '100';
    elements.lfAdditionalFilters.innerHTML = '<p style="color:#888;font-style:italic;font-size:12px;">No additional filters. Click "+ Add Filter" to add more conditions.</p>';
    elements.lfResultsSection.style.display = 'none';
    filterAdditionalConditions = [];
}

function showLfOutput(message, type = 'info') {
    const outputDiv = document.createElement('div');
    outputDiv.className = `lf-output-message`;
    outputDiv.style.cssText = `
        margin-top: 10px;
        padding: 10px;
        border-radius: 4px;
        background: ${type === 'error' ? '#fde8e8' : type === 'success' ? '#e6f4ea' : '#e8f0fe'};
        border: 1px solid ${type === 'error' ? '#d13438' : type === 'success' ? '#107c10' : '#0078d4'};
        color: ${type === 'error' ? '#d13438' : type === 'success' ? '#107c10' : '#0078d4'};
        white-space: pre-wrap;
    `;
    outputDiv.textContent = message;

    // Remove previous messages
    const prevMessages = document.querySelectorAll('.lf-output-message');
    prevMessages.forEach(msg => msg.remove());

    // Insert after the results section
    elements.lfResultsSection.parentNode.insertBefore(outputDiv, elements.lfResultsSection);
}
