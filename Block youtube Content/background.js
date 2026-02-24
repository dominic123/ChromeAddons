// Background service worker for YouTube Channel Blocker

// Handle installation
chrome.runtime.onInstalled.addListener((details) => {
  console.log('YouTube Channel Blocker installed/updated');

  // Initialize storage if empty
  chrome.storage.local.get(['blockedChannels'], (result) => {
    if (!result.blockedChannels) {
      chrome.storage.local.set({ blockedChannels: [] });
    }
  });

  // Show important security notification on install/update
  if (details.reason === 'install' || details.reason === 'update') {
    chrome.notifications.create({
      type: 'basic',
      iconUrl: 'icon48.png',
      title: 'YouTube Channel Blocker - Security Notice',
      message: '⚠️ For full parental control: Create a Supervised Account in Chrome settings to prevent extension removal. Open extension popup for details.',
      priority: 2
    });
  }
});

// Listen for messages from content script
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'blockChannel') {
    blockChannel(request.channelId, request.channelName);
    sendResponse({ success: true });
  } else if (request.action === 'unblockChannel') {
    unblockChannel(request.channelId);
    sendResponse({ success: true });
  } else if (request.action === 'getBlockedChannels') {
    chrome.storage.local.get(['blockedChannels'], (result) => {
      sendResponse({ blockedChannels: result.blockedChannels || [] });
    });
    return true; // Keep message channel open for async response
  } else if (request.action === 'isChannelBlocked') {
    chrome.storage.local.get(['blockedChannels'], (result) => {
      const blockedChannels = result.blockedChannels || [];
      const isBlocked = blockedChannels.some(ch => ch.id === request.channelId);
      sendResponse({ isBlocked });
    });
    return true;
  }
  return true;
});

function blockChannel(channelId, channelName) {
  chrome.storage.local.get(['blockedChannels'], (result) => {
    const blockedChannels = result.blockedChannels || [];

    // Check if already blocked
    if (!blockedChannels.some(ch => ch.id === channelId)) {
      blockedChannels.push({
        id: channelId,
        name: channelName,
        blockedAt: Date.now()
      });
      chrome.storage.local.set({ blockedChannels });

      // Notify all tabs to refresh
      chrome.tabs.query({ url: '*://www.youtube.com/*' }, (tabs) => {
        tabs.forEach(tab => {
          chrome.tabs.sendMessage(tab.id, { action: 'refreshBlocks' });
        });
      });
    }
  });
}

function unblockChannel(channelId) {
  chrome.storage.local.get(['blockedChannels'], (result) => {
    const blockedChannels = result.blockedChannels || [];
    const updated = blockedChannels.filter(ch => ch.id !== channelId);
    chrome.storage.local.set({ blockedChannels: updated });

    // Notify all tabs to refresh
    chrome.tabs.query({ url: '*://www.youtube.com/*' }, (tabs) => {
      tabs.forEach(tab => {
        chrome.tabs.sendMessage(tab.id, { action: 'refreshBlocks' });
      });
    });
  });
}
