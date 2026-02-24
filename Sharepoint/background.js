// Background Service Worker for SharePoint Field Creator
// Manifest V3 requires service worker instead of background pages

// Install event - set up any initial state
self.addEventListener('install', (event) => {
    console.log('SharePoint Field Creator - Service Worker installing...');
    self.skipWaiting();
});

// Activate event - clean up old versions
self.addEventListener('activate', (event) => {
    console.log('SharePoint Field Creator - Service Worker activating...');
    event.waitUntil(self.clients.claim());
});

// Handle extension icon click
chrome.action.onClicked.addListener((tab) => {
    console.log('Extension icon clicked on tab:', tab.url);
});

// Handle messages from content scripts
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    console.log('Background received message:', request);

    if (request.action === 'log') {
        console.log('From content script:', request.data);
    }

    return true;
});

// Handle extension installation or update
chrome.runtime.onInstalled.addListener((details) => {
    if (details.reason === 'install') {
        console.log('Extension installed for the first time');
    } else if (details.reason === 'update') {
        console.log('Extension updated');
    }
});

console.log('SharePoint Field Creator - Service Worker loaded');
