// Popup script for YouTube Channel Blocker

let blockedChannels = [];
let currentPin = null;

// Load blocked channels and PIN on popup open
function loadSettings() {
  chrome.runtime.sendMessage({ action: 'getBlockedChannels' }, (response) => {
    if (response && response.blockedChannels) {
      blockedChannels = response.blockedChannels;
      renderChannelList();
    }
  });

  // Load PIN
  chrome.storage.local.get(['parentalPin'], (result) => {
    currentPin = result.parentalPin || null;
    updatePinUI();
  });
}

// Update PIN UI based on current state
function updatePinUI() {
  const pinStatus = document.getElementById('pinStatus');
  const pinSetup = document.getElementById('pinSetup');
  const pinChange = document.getElementById('pinChange');

  if (currentPin) {
    pinStatus.textContent = 'Active';
    pinStatus.className = 'pin-status set';
    pinSetup.classList.add('hidden');
    pinChange.classList.remove('hidden');
  } else {
    pinStatus.textContent = 'Not Set';
    pinStatus.className = 'pin-status not-set';
    pinSetup.classList.remove('hidden');
    pinChange.classList.add('hidden');
  }
}

// Set new PIN
function setPin() {
  const pinInput = document.getElementById('newPinInput');
  const newPin = pinInput.value.trim();

  // Validate PIN
  if (!newPin) {
    alert('Please enter a PIN');
    return;
  }

  if (!/^\d{4,6}$/.test(newPin)) {
    alert('PIN must be 4-6 digits');
    return;
  }

  // Save PIN
  chrome.storage.local.set({ parentalPin: newPin }, () => {
    currentPin = newPin;
    pinInput.value = '';
    updatePinUI();

    // Notify all tabs about PIN update
    chrome.tabs.query({ url: '*://www.youtube.com/*' }, (tabs) => {
      tabs.forEach(tab => {
        chrome.tabs.sendMessage(tab.id, { action: 'pinUpdated' }).catch(() => {});
      });
    });

    alert('PIN set successfully! Blocked content will now require PIN to unlock.');
  });
}

// Update existing PIN
function updatePin() {
  const currentPinInput = document.getElementById('currentPinInput');
  const newPinInput = document.getElementById('newPinInput2');
  const enteredCurrentPin = currentPinInput.value.trim();
  const newPin = newPinInput.value.trim();

  // Validate
  if (!enteredCurrentPin || !newPin) {
    alert('Please fill in all fields');
    return;
  }

  if (enteredCurrentPin !== currentPin) {
    alert('Current PIN is incorrect');
    return;
  }

  if (!/^\d{4,6}$/.test(newPin)) {
    alert('New PIN must be 4-6 digits');
    return;
  }

  // Save new PIN
  chrome.storage.local.set({ parentalPin: newPin }, () => {
    currentPin = newPin;
    currentPinInput.value = '';
    newPinInput.value = '';
    updatePinUI();

    // Notify all tabs about PIN update
    chrome.tabs.query({ url: '*://www.youtube.com/*' }, (tabs) => {
      tabs.forEach(tab => {
        chrome.tabs.sendMessage(tab.id, { action: 'pinUpdated' }).catch(() => {});
      });
    });

    alert('PIN updated successfully!');
  });
}

// Remove PIN
function removePin() {
  if (!currentPin) return;

  if (confirm('Are you sure you want to remove the parental lock PIN? Blocked content will be completely hidden without a PIN.')) {
    chrome.storage.local.remove(['parentalPin'], () => {
      currentPin = null;
      updatePinUI();

      // Notify all tabs about PIN update
      chrome.tabs.query({ url: '*://www.youtube.com/*' }, (tabs) => {
        tabs.forEach(tab => {
          chrome.tabs.sendMessage(tab.id, { action: 'pinUpdated' }).catch(() => {});
        });
      });

      alert('PIN removed. Blocked content is now completely hidden.');
    });
  }
}

// Render the list of blocked channels
function renderChannelList() {
  const listElement = document.getElementById('channelList');
  const countElement = document.getElementById('channelCount');

  countElement.textContent = `${blockedChannels.length} channel${blockedChannels.length !== 1 ? 's' : ''} blocked`;

  if (blockedChannels.length === 0) {
    listElement.innerHTML = '<p class="empty-message">No blocked channels yet</p>';
    return;
  }

  listElement.innerHTML = blockedChannels.map(channel => `
    <div class="channel-item">
      <div class="channel-info">
        <div class="channel-name">${escapeHtml(channel.name || 'Unknown Channel')}</div>
        <div class="channel-id">${escapeHtml(channel.id)}</div>
      </div>
      <button class="unblock-btn" data-id="${escapeHtml(channel.id)}">Unblock</button>
    </div>
  `).join('');

  // Add event listeners to unblock buttons
  document.querySelectorAll('.unblock-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const channelId = btn.getAttribute('data-id');
      unblockChannel(channelId);
    });
  });
}

// Escape HTML to prevent XSS
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// Block a channel manually
function blockChannel() {
  const channelInput = document.getElementById('channelInput');
  const channelNameInput = document.getElementById('channelNameInput');

  const channelId = channelInput.value.trim();
  const channelName = channelNameInput.value.trim() || channelId;

  if (!channelId) {
    alert('Please enter a channel ID or URL');
    return;
  }

  // Extract channel ID from URL if needed
  let extractedId = channelId;
  if (channelId.includes('youtube.com')) {
    const match = channelId.match(/\/channel\/([a-zA-Z0-9_-]+)/) ||
                  channelId.match(/\/(@[a-zA-Z0-9_.-]+)/) ||
                  channelId.match(/\/c\/([a-zA-Z0-9_-]+)/) ||
                  channelId.match(/\/user\/([a-zA-Z0-9_-]+)/);
    if (match) {
      extractedId = match[1];
    }
  }

  // Check if already blocked
  if (blockedChannels.some(ch => ch.id === extractedId)) {
    alert('This channel is already blocked');
    return;
  }

  chrome.runtime.sendMessage({
    action: 'blockChannel',
    channelId: extractedId,
    channelName: channelName
  }, () => {
    channelInput.value = '';
    channelNameInput.value = '';
    loadSettings();
  });
}

// Unblock a channel
function unblockChannel(channelId) {
  chrome.runtime.sendMessage({
    action: 'unblockChannel',
    channelId: channelId
  }, () => {
    loadSettings();
  });
}

// Unblock all channels
function unblockAllChannels() {
  if (blockedChannels.length === 0) {
    alert('No channels to unblock');
    return;
  }

  if (confirm(`Are you sure you want to unblock all ${blockedChannels.length} channels?`)) {
    chrome.storage.local.set({ blockedChannels: [] }, () => {
      // Notify all tabs to refresh
      chrome.tabs.query({ url: '*://www.youtube.com/*' }, (tabs) => {
        tabs.forEach(tab => {
          chrome.tabs.sendMessage(tab.id, { action: 'refreshBlocks' }).catch(() => {});
        });
      });
      loadSettings();
    });
  }
}

// Event listeners
document.addEventListener('DOMContentLoaded', () => {
  loadSettings();

  // PIN button listeners
  document.getElementById('setPinBtn').addEventListener('click', setPin);
  document.getElementById('updatePinBtn').addEventListener('click', updatePin);
  document.getElementById('removePinBtn').addEventListener('click', removePin);

  // Enter key for PIN inputs
  document.getElementById('newPinInput').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') setPin();
  });

  document.getElementById('newPinInput2').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') updatePin();
  });

  // Channel blocking
  document.getElementById('addBtn').addEventListener('click', blockChannel);

  document.getElementById('channelInput').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      blockChannel();
    }
  });

  document.getElementById('channelNameInput').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      blockChannel();
    }
  });

  // Unblock all
  document.getElementById('unblockAllBtn').addEventListener('click', unblockAllChannels);
});
