// Content script for YouTube Channel Blocker

let blockedChannels = [];
let observer = null;
let userPin = null;

// Load blocked channels and PIN on startup
function loadBlockedChannels() {
  chrome.runtime.sendMessage({ action: 'getBlockedChannels' }, (response) => {
    if (response && response.blockedChannels) {
      blockedChannels = response.blockedChannels;
    }
    // Also load PIN
    chrome.storage.local.get(['parentalPin'], (result) => {
      userPin = result.parentalPin || null;
      // Check URL again after loading settings
      if (!checkCurrentUrl()) {
        processPage();
      }
    });
  });
}

// Check if a channel ID is blocked
function isChannelBlocked(channelId) {
  return blockedChannels.some(ch => ch.id === channelId);
}

// Get blocked channel info
function getBlockedChannelInfo(channelId) {
  return blockedChannels.find(ch => ch.id === channelId);
}

// Extract channel ID from URL
function extractChannelIdFromUrl(href) {
  if (!href) return null;

  // Handle @channel URLs
  const atMatch = href.match(/\/(@[a-zA-Z0-9_.-]+)/);
  if (atMatch) return atMatch[1];

  // Handle /channel/ URLs
  const channelMatch = href.match(/\/channel\/([a-zA-Z0-9_-]+)/);
  if (channelMatch) return channelMatch[1];

  // Handle /c/ URLs
  const cMatch = href.match(/\/c\/([a-zA-Z0-9_-]+)/);
  if (cMatch) return 'c/' + cMatch[1];

  // Handle /user/ URLs
  const userMatch = href.match(/\/user\/([a-zA-Z0-9_-]+)/);
  if (userMatch) return 'user/' + userMatch[1];

  return null;
}

// Extract channel ID from element
function extractChannelId(element) {
  // Try to find channel link
  const link = element.querySelector('a[href*="/channel/"], a[href*="/@"], a[href*="/c/"], a[href*="/user/"]');
  if (!link) return null;
  return extractChannelIdFromUrl(link.getAttribute('href'));
}

// Extract channel name
function extractChannelName(element) {
  const nameElement = element.querySelector('#channel-name a, #text a, .ytd-channel-name a, #avatar-link, ytd-channel-name #text');
  if (nameElement) {
    return nameElement.textContent.trim();
  }
  return null;
}

// Show PIN entry modal
function showPinModal(channelName, channelId, onSuccess) {
  // Remove existing modal if any
  const existingModal = document.getElementById('yt-blocker-modal');
  if (existingModal) existingModal.remove();

  // Create modal
  const modal = document.createElement('div');
  modal.id = 'yt-blocker-modal';
  modal.innerHTML = `
    <div class="yt-blocker-overlay">
      <div class="yt-blocker-content">
        <div class="yt-blocker-icon">🔒</div>
        <h2>Parental Lock</h2>
        <p>This channel has been blocked</p>
        <p class="channel-info-text">${escapeHtml(channelName)}</p>
        <div class="pin-container">
          <input type="password" id="yt-pin-input" class="pin-input" placeholder="Enter PIN" maxlength="6">
          <p id="pin-error" class="pin-error">Incorrect PIN</p>
        </div>
        <div class="button-row">
          <button class="yt-blocker-cancel">Cancel</button>
          <button class="yt-blocker-submit">Unlock</button>
        </div>
      </div>
    </div>
  `;

  // Add styles
  modal.querySelector('.yt-blocker-overlay').style.cssText = `
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(0, 0, 0, 0.7);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 100000;
    animation: fadeIn 0.2s ease-out;
  `;

  modal.querySelector('.yt-blocker-content').style.cssText = `
    background: white;
    padding: 32px;
    border-radius: 12px;
    text-align: center;
    max-width: 380px;
    box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
    animation: slideUp 0.3s ease-out;
  `;

  modal.querySelector('.yt-blocker-icon').style.cssText = `
    font-size: 64px;
    margin-bottom: 16px;
  `;

  modal.querySelector('h2').style.cssText = `
    color: #065fd4;
    font-size: 24px;
    margin: 0 0 12px 0;
    font-family: Roboto, Arial, sans-serif;
  `;

  const paragraphs = modal.querySelectorAll('p');
  if (paragraphs[0]) paragraphs[0].style.cssText = `color: #666; font-size: 14px; margin: 0 0 8px 0; font-family: Roboto, Arial, sans-serif;`;
  if (paragraphs[1]) paragraphs[1].style.cssText = `color: #333; background: #f1f3f4; padding: 12px; border-radius: 8px; margin: 12px 0 20px 0; font-size: 14px; font-family: Roboto, Arial, sans-serif;`;

  modal.querySelector('.pin-container').style.cssText = `
    margin-bottom: 20px;
  `;

  const pinInput = modal.querySelector('#yt-pin-input');
  pinInput.style.cssText = `
    width: 100%;
    padding: 14px 16px;
    border: 2px solid #ddd;
    border-radius: 8px;
    font-size: 18px;
    text-align: center;
    letter-spacing: 4px;
    font-family: Roboto, Arial, sans-serif;
    box-sizing: border-box;
  `;
  pinInput.focus();

  const pinError = modal.querySelector('#pin-error');
  pinError.style.cssText = `
    color: #d93025;
    font-size: 13px;
    margin-top: 8px;
    display: none;
    font-family: Roboto, Arial, sans-serif;
  `;

  modal.querySelector('.button-row').style.cssText = `
    display: flex;
    gap: 12px;
    justify-content: center;
  `;

  const cancelBtn = modal.querySelector('.yt-blocker-cancel');
  cancelBtn.style.cssText = `
    background: #f1f3f4;
    color: #333;
    border: none;
    padding: 12px 24px;
    border-radius: 24px;
    cursor: pointer;
    font-size: 14px;
    font-weight: 500;
    font-family: Roboto, Arial, sans-serif;
  `;

  const submitBtn = modal.querySelector('.yt-blocker-submit');
  submitBtn.style.cssText = `
    background: #065fd4;
    color: white;
    border: none;
    padding: 12px 24px;
    border-radius: 24px;
    cursor: pointer;
    font-size: 14px;
    font-weight: 500;
    font-family: Roboto, Arial, sans-serif;
  `;

  // Add keyframe animations
  const style = document.createElement('style');
  style.id = 'yt-blocker-animations';
  style.textContent = `
    @keyframes fadeIn {
      from { opacity: 0; }
      to { opacity: 1; }
    }
    @keyframes slideUp {
      from { transform: translateY(20px); opacity: 0; }
      to { transform: translateY(0); opacity: 1; }
    }
  `;
  if (!document.getElementById('yt-blocker-animations')) {
    document.head.appendChild(style);
  }

  // Function to verify PIN
  function verifyPin() {
    const enteredPin = pinInput.value;

    if (!userPin) {
      // No PIN set - allow access and show setup hint
      modal.remove();
      onSuccess();
      alert('No PIN set! Go to extension settings to set up a parental PIN.');
      return;
    }

    if (enteredPin === userPin) {
      modal.remove();
      onSuccess();
    } else {
      pinError.style.display = 'block';
      pinInput.style.borderColor = '#d93025';
      pinInput.value = '';
      pinInput.focus();
    }
  }

  // Event listeners
  submitBtn.onclick = verifyPin;
  cancelBtn.onclick = () => modal.remove();
  modal.querySelector('.yt-blocker-overlay').onclick = (e) => {
    if (e.target === modal.querySelector('.yt-blocker-overlay')) {
      modal.remove();
    }
  };

  pinInput.onkeypress = (e) => {
    pinError.style.display = 'none';
    pinInput.style.borderColor = '#ddd';
    if (e.key === 'Enter') {
      verifyPin();
    }
  };

  // Close on Escape key
  const escapeHandler = (e) => {
    if (e.key === 'Escape') {
      modal.remove();
      document.removeEventListener('keydown', escapeHandler);
    }
  };
  document.addEventListener('keydown', escapeHandler);

  document.body.appendChild(modal);
}

// Show blocked modal (for when PIN is not set)
function showBlockedModal(channelName, channelId) {
  // If PIN is set, show PIN modal instead
  if (userPin) {
    showPinModal(channelName, channelId, () => {
      // Allow access for this session
      // Store temporary access grant
      sessionStorage.setItem('yt-unlocked-' + channelId, Date.now());
    });
    return;
  }

  // No PIN - show simple blocked message
  showPinModal(channelName, channelId, () => {
    // Allow access
    sessionStorage.setItem('yt-unlocked-' + channelId, Date.now());
  });
}

// Check if channel is temporarily unlocked
function isChannelUnlocked(channelId) {
  const unlocked = sessionStorage.getItem('yt-unlocked-' + channelId);
  if (unlocked) {
    // Check if unlock is still valid (30 minutes)
    const unlockTime = parseInt(unlocked);
    if (Date.now() - unlockTime < 30 * 60 * 1000) {
      return true;
    } else {
      sessionStorage.removeItem('yt-unlocked-' + channelId);
    }
  }
  return false;
}

// Escape HTML
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// Intercept clicks on links
function interceptLinkClicks() {
  document.addEventListener('click', (e) => {
    // Find closest anchor tag
    const link = e.target.closest('a[href*="/channel/"], a[href*="/@"], a[href*="/c/"], a[href*="/user/"]');
    if (!link) return;

    const href = link.getAttribute('href');
    const channelId = extractChannelIdFromUrl(href);

    if (channelId && isChannelBlocked(channelId) && !isChannelUnlocked(channelId)) {
      e.preventDefault();
      e.stopPropagation();

      // Get channel name
      const channelInfo = getBlockedChannelInfo(channelId);
      const channelName = channelInfo ? channelInfo.name : channelId;

      showBlockedModal(channelName, channelId);
    }
  }, true);
}

// Check current URL and show modal if blocked
function checkCurrentUrl() {
  const pathname = window.location.pathname;
  let channelId = null;

  // Check if on a watch page - need to extract channel from page
  if (pathname.startsWith('/watch') || pathname.startsWith('/shorts')) {
    // Try to find channel link on the page after a short delay
    setTimeout(() => {
      checkWatchPageChannel();
    }, 1000);
    return false;
  }

  // Don't block on homepage, search, or feeds
  if (pathname === '/' ||
      pathname.startsWith('/results') ||
      pathname.startsWith('/feed') ||
      pathname.startsWith('/playlist') ||
      pathname.startsWith('/premium') ||
      pathname.startsWith('/gaming') ||
      pathname.startsWith('/music')) {
    return false;
  }

  // Extract channel ID from current URL
  if (pathname.startsWith('/@')) {
    const match = pathname.match(/^\/(@[a-zA-Z0-9_.-]+)/);
    if (match) channelId = match[1];
  } else if (pathname.startsWith('/channel/')) {
    const match = pathname.match(/^\/channel\/([a-zA-Z0-9_-]+)/);
    if (match) channelId = match[1];
  } else if (pathname.startsWith('/c/')) {
    const match = pathname.match(/^\/c\/([a-zA-Z0-9_-]+)/);
    if (match) channelId = 'c/' + match[1];
  } else if (pathname.startsWith('/user/')) {
    const match = pathname.match(/^\/user\/([a-zA-Z0-9_-]+)/);
    if (match) channelId = 'user/' + match[1];
  }

  // If on a blocked channel page, show PIN modal
  if (channelId && isChannelBlocked(channelId) && !isChannelUnlocked(channelId)) {
    const channelInfo = getBlockedChannelInfo(channelId);
    const channelName = channelInfo ? channelInfo.name : channelId;

    // Replace page with PIN prompt
    document.documentElement.innerHTML = `
      <head>
        <title>Parental Lock</title>
        <style>
          * { margin: 0; padding: 0; box-sizing: border-box; }
          body {
            font-family: Roboto, Arial, sans-serif;
            background: #f9f9f9;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
          }
          .blocked-container {
            text-align: center;
            padding: 40px;
            background: white;
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
            max-width: 400px;
          }
          .icon { font-size: 80px; margin-bottom: 20px; }
          h1 { color: #065fd4; font-size: 32px; margin-bottom: 16px; }
          p { color: #666; font-size: 16px; margin-bottom: 12px; }
          .channel-name {
            background: #f1f3f4;
            padding: 12px 20px;
            border-radius: 8px;
            margin: 20px 0;
            font-weight: 500;
            color: #333;
          }
          .pin-input {
            width: 100%;
            padding: 14px 16px;
            border: 2px solid #ddd;
            border-radius: 8px;
            font-size: 18px;
            text-align: center;
            letter-spacing: 4px;
            margin-bottom: 8px;
            box-sizing: border-box;
          }
          .pin-error {
            color: #d93025;
            font-size: 13px;
            display: none;
            margin-bottom: 16px;
          }
          .button-row { display: flex; gap: 12px; justify-content: center; }
          .btn {
            border: none;
            padding: 12px 32px;
            border-radius: 24px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
          }
          .btn-primary { background: #065fd4; color: white; }
          .btn-primary:hover { background: #0550b8; }
          .btn-secondary { background: #f1f3f4; color: #333; }
          .btn-secondary:hover { background: #e0e0e0; }
        </style>
      </head>
      <body>
        <div class="blocked-container">
          <div class="icon">🔒</div>
          <h1>Parental Lock</h1>
          <p>This channel has been blocked</p>
          <div class="channel-name">${escapeHtml(channelName)}</div>
          <input type="password" id="page-pin-input" class="pin-input" placeholder="Enter PIN" maxlength="6">
          <p id="page-pin-error" class="pin-error">Incorrect PIN</p>
          <div class="button-row">
            <button class="btn btn-secondary" onclick="window.location.href='https://www.youtube.com'">Home</button>
            <button class="btn btn-primary" id="unlock-btn">Unlock</button>
          </div>
        </div>
        <script>
          const userPin = ${JSON.stringify(userPin || 'null')};
          const channelId = ${JSON.stringify(channelId)};
          const pinInput = document.getElementById('page-pin-input');
          const pinError = document.getElementById('page-pin-error');
          const unlockBtn = document.getElementById('unlock-btn');

          pinInput.focus();

          function verifyPin() {
            const enteredPin = pinInput.value;

            if (!userPin) {
              sessionStorage.setItem('yt-unlocked-' + channelId, Date.now());
              location.reload();
              return;
            }

            if (enteredPin === userPin) {
              sessionStorage.setItem('yt-unlocked-' + channelId, Date.now());
              location.reload();
            } else {
              pinError.style.display = 'block';
              pinInput.style.borderColor = '#d93025';
              pinInput.value = '';
              pinInput.focus();
            }
          }

          unlockBtn.onclick = verifyPin;
          pinInput.onkeypress = function(e) {
            pinError.style.display = 'none';
            pinInput.style.borderColor = '#ddd';
            if (e.key === 'Enter') verifyPin();
          };
        </script>
      </body>
    `;

    window.stop();
    return true;
  }
  return false;
}

// Check if current video is from a blocked channel
function checkWatchPageChannel() {
  // Check if already unlocked
  const channelLink = document.querySelector('#owner #channel-name a, ytd-video-owner-renderer a, #top-row ytd-channel-name a');
  if (channelLink) {
    const href = channelLink.getAttribute('href');
    const channelId = extractChannelIdFromUrl(href);

    if (channelId && isChannelBlocked(channelId) && !isChannelUnlocked(channelId)) {
      const channelInfo = getBlockedChannelInfo(channelId);
      const channelName = channelInfo ? channelInfo.name : channelId;

      // Replace the entire page content
      document.documentElement.innerHTML = `
        <head>
          <title>Parental Lock</title>
          <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body {
              font-family: Roboto, Arial, sans-serif;
              background: #f9f9f9;
              display: flex;
              align-items: center;
              justify-content: center;
              min-height: 100vh;
            }
            .blocked-container {
              text-align: center;
              padding: 40px;
              background: white;
              border-radius: 12px;
              box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
              max-width: 400px;
            }
            .icon { font-size: 80px; margin-bottom: 20px; }
            h1 { color: #065fd4; font-size: 32px; margin-bottom: 16px; }
            p { color: #666; font-size: 16px; margin-bottom: 12px; }
            .channel-name {
              background: #f1f3f4;
              padding: 12px 20px;
              border-radius: 8px;
              margin: 20px 0;
              font-weight: 500;
              color: #333;
            }
            .pin-input {
              width: 100%;
              padding: 14px 16px;
              border: 2px solid #ddd;
              border-radius: 8px;
              font-size: 18px;
              text-align: center;
              letter-spacing: 4px;
              margin-bottom: 8px;
              box-sizing: border-box;
            }
            .pin-error {
              color: #d93025;
              font-size: 13px;
              display: none;
              margin-bottom: 16px;
            }
            .button-row { display: flex; gap: 12px; justify-content: center; }
            .btn {
              border: none;
              padding: 12px 32px;
              border-radius: 24px;
              cursor: pointer;
              font-size: 14px;
              font-weight: 500;
            }
            .btn-primary { background: #065fd4; color: white; }
            .btn-primary:hover { background: #0550b8; }
            .btn-secondary { background: #f1f3f4; color: #333; }
            .btn-secondary:hover { background: #e0e0e0; }
          </style>
        </head>
        <body>
          <div class="blocked-container">
            <div class="icon">🔒</div>
            <h1>Parental Lock</h1>
            <p>This video is from a blocked channel</p>
            <div class="channel-name">${escapeHtml(channelName)}</div>
            <input type="password" id="page-pin-input" class="pin-input" placeholder="Enter PIN" maxlength="6">
            <p id="page-pin-error" class="pin-error">Incorrect PIN</p>
            <div class="button-row">
              <button class="btn btn-secondary" onclick="history.back()">← Back</button>
              <button class="btn btn-primary" id="unlock-btn">Unlock</button>
            </div>
          </div>
          <script>
            const userPin = ${JSON.stringify(userPin || 'null')};
            const channelId = ${JSON.stringify(channelId)};
            const pinInput = document.getElementById('page-pin-input');
            const pinError = document.getElementById('page-pin-error');
            const unlockBtn = document.getElementById('unlock-btn');

            pinInput.focus();

            function verifyPin() {
              const enteredPin = pinInput.value;

              if (!userPin) {
                sessionStorage.setItem('yt-unlocked-' + channelId, Date.now());
                location.reload();
                return;
              }

              if (enteredPin === userPin) {
                sessionStorage.setItem('yt-unlocked-' + channelId, Date.now());
                location.reload();
              } else {
                pinError.style.display = 'block';
                pinInput.style.borderColor = '#d93025';
                pinInput.value = '';
                pinInput.focus();
              }
            }

            unlockBtn.onclick = verifyPin;
            pinInput.onkeypress = function(e) {
              pinError.style.display = 'none';
              pinInput.style.borderColor = '#ddd';
              if (e.key === 'Enter') verifyPin();
            };
          </script>
        </body>
      `;
    }
  }
}

// Process page - hide blocked content
function processPage() {
  // Hide videos from blocked channels
  hideBlockedVideos('ytd-rich-item-renderer, ytd-grid-video-renderer, ytd-video-renderer');
  hideBlockedVideos('ytd-video-renderer', 'ytd-search');
  hideBlockedVideos('ytd-compact-video-renderer', 'ytd-watch-next');

  // Hide comments from blocked channels
  hideBlockedComments();

  // Add block button to channel pages
  addBlockButton();
}

// Hide videos from blocked channels
function hideBlockedVideos(selector, container = 'body') {
  const containers = document.querySelectorAll(container);
  containers.forEach(cont => {
    const videos = cont.querySelectorAll(selector);
    videos.forEach(video => {
      const channelId = extractChannelId(video);
      if (channelId && isChannelBlocked(channelId) && !isChannelUnlocked(channelId)) {
        video.style.display = 'none';
        video.classList.add('yt-blocker-hidden');
      }
    });
  });
}

// Hide comments from blocked channels
function hideBlockedComments() {
  const comments = document.querySelectorAll('ytd-comment-thread-renderer');
  comments.forEach(comment => {
    const channelId = extractChannelId(comment);
    if (channelId && isChannelBlocked(channelId) && !isChannelUnlocked(channelId)) {
      comment.style.display = 'none';
      comment.classList.add('yt-blocker-hidden');
    }
  });
}

// Add block button to channel pages
function addBlockButton() {
  if (document.querySelector('#yt-channel-blocker-btn')) return;

  const subscribeButton = document.querySelector('#subscribe-button, #subscribe-button-shape, paper-button[sub-button]');
  if (!subscribeButton) return;

  const blockBtn = document.createElement('button');
  blockBtn.id = 'yt-channel-blocker-btn';
  blockBtn.textContent = 'Block Channel';
  blockBtn.style.cssText = `
    margin-left: 8px;
    padding: 10px 16px;
    background: #ff0000;
    color: white;
    border: none;
    border-radius: 18px;
    cursor: pointer;
    font-size: 14px;
    font-weight: 500;
  `;

  const channelId = extractChannelId(document.body);
  const channelName = extractChannelName(document.body);

  blockBtn.onclick = () => {
    if (channelId) {
      chrome.runtime.sendMessage({
        action: 'blockChannel',
        channelId: channelId,
        channelName: channelName || 'Unknown Channel'
      }, () => {
        blockBtn.textContent = 'Blocked!';
        blockBtn.style.background = '#666';
        blockBtn.disabled = true;
      });
    }
  };

  subscribeButton.parentNode.appendChild(blockBtn);
}

// Watch for page changes (YouTube is a SPA)
function observePageChanges() {
  if (observer) observer.disconnect();

  observer = new MutationObserver(() => {
    processPage();
  });

  observer.observe(document.body, {
    childList: true,
    subtree: true
  });
}

// Listen for messages from background (including PIN updates)
chrome.runtime.onMessage.addListener((request) => {
  if (request.action === 'refreshBlocks') {
    loadBlockedChannels();
  } else if (request.action === 'pinUpdated') {
    chrome.storage.local.get(['parentalPin'], (result) => {
      userPin = result.parentalPin || null;
    });
  }
});

// ============================================
// CRITICAL: Check URL IMMEDIATELY on script load
// This catches direct URL entries (typing in address bar)
// ============================================
(function immediateUrlCheck() {
  const pathname = window.location.pathname;

  // Quick check if this might be a channel URL or watch page
  if (pathname.startsWith('/@') ||
      pathname.startsWith('/channel/') ||
      pathname.startsWith('/c/') ||
      pathname.startsWith('/user/') ||
      pathname.startsWith('/watch') ||
      pathname.startsWith('/shorts')) {

    // Need to load blocked channels and PIN first, then check
    chrome.runtime.sendMessage({ action: 'getBlockedChannels' }, (response) => {
      if (response && response.blockedChannels) {
        blockedChannels = response.blockedChannels;
      }
      // Also load PIN
      chrome.storage.local.get(['parentalPin'], (result) => {
        userPin = result.parentalPin || null;

        // Now check if current URL is blocked
        const blocked = checkCurrentUrl();
        if (!blocked) {
          // Not blocked, continue with normal initialization
          initializeExtension();
        }
      });
    });
  } else {
    // Not a channel URL, initialize normally
    initializeExtension();
  }
})();

// Main initialization function
function initializeExtension() {
  loadBlockedChannels();
  interceptLinkClicks();

  // Wait for page to load
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      processPage();
      observePageChanges();
    });
  } else {
    processPage();
    observePageChanges();
  }

  // Also run on navigation (YouTube SPA)
  let lastUrl = location.href;
  new MutationObserver(() => {
    if (location.href !== lastUrl) {
      lastUrl = location.href;
      setTimeout(() => {
        checkCurrentUrl();
        processPage();
      }, 500);
    }
  }).observe(document, { subtree: true, childList: true });
}
