# YouTube Channel Blocker - Chrome Extension

A Chrome extension that allows you to block YouTube channels you don't want to see in your feed, search results, and recommendations.

## Features

- **Block channels from any YouTube page** - Videos from blocked channels are hidden
- **Quick block button** - One-click blocking on channel pages
- **Manage blocked channels** - View and unblock channels through the popup
- **Works everywhere** - Hides blocked channels from:
  - Home page feed
  - Search results
  - Recommended videos sidebar
  - Comments section
  - Channel pages (redirects to blocked page)

## How to Test Locally

### Step 1: Open Chrome Extensions Page
1. Open Google Chrome
2. Navigate to `chrome://extensions/` (type this in the address bar)
3. Alternatively: Chrome Menu (⋮) → More Tools → Extensions

### Step 2: Enable Developer Mode
1. Look for the **"Developer mode"** toggle in the top-right corner
2. Click it to enable (it should turn blue)

### Step 3: Load the Extension
1. Click the **"Load unpacked"** button that appears (top-left)
2. A file picker will open
3. Navigate to and select the folder containing this extension (the folder with `manifest.json`)
4. Click "Select Folder"

### Step 4: Verify Installation
You should see "YouTube Channel Blocker" appear in your extensions list with its icon.

### Step 5: Test the Extension

#### Option A: Test on YouTube Homepage
1. Go to [YouTube](https://www.youtube.com)
2. Videos from blocked channels should be hidden
3. Click the extension icon to see your blocked channels list

#### Option B: Test on a Channel Page
1. Visit any YouTube channel (e.g., `https://www.youtube.com/@SomeChannel`)
2. You should see a red **"Block Channel"** button next to the Subscribe button
3. Click it to block the channel
4. Refresh the page - you should see the "Channel Blocked" message

#### Option C: Manual Block via Popup
1. Click the extension icon in Chrome toolbar
2. Enter a channel ID or URL in the input field
3. Click "Block Channel"
4. The channel will appear in your blocked list

#### Option D: Test Blocking & Unblocking
1. Click the extension icon
2. Block a channel using any method above
3. Go to the popup and click "Unblock" to remove it
4. The channel's content should reappear

## How to Find Channel IDs

There are several ways to get a channel ID:

1. **From the channel URL:**
   - `https://www.youtube.com/@ChannelName` - Use `@ChannelName`
   - `https://www.youtube.com/channel/UC...` - Use the `UC...` part
   - `https://www.youtube.com/c/ChannelName` - Use `c/ChannelName`
   - `https://www.youtube.com/user/Username` - Use `user/Username`

2. **Right-click method:**
   - Go to the channel page
   - Right-click and "View Page Source"
   - Search for `"channelId"` to find the ID

3. **Use the block button:**
   - Just visit the channel page and click the "Block Channel" button

## Project Structure

```
youtube-channel-blocker/
├── manifest.json       # Extension configuration
├── background.js       # Service worker (handles storage & messaging)
├── content.js          # Content script (runs on YouTube pages)
├── popup.html          # Popup interface
├── popup.js            # Popup logic
├── popup.css           # Popup styling
├── styles.css          # Content script styles
└── README.md           # This file
```

## Development Notes

- The extension uses Manifest V3 (latest Chrome extension format)
- Blocked channels are stored in Chrome's local storage
- The content script observes DOM changes to handle YouTube's SPA navigation
- The service worker manages the block list and communicates between components

## Troubleshooting

**Extension not working?**
1. Make sure Developer Mode is enabled
2. Try clicking the "Refresh" icon on the extension card
3. Check that all files are in the same folder
4. Refresh YouTube pages after making changes

**Changes not appearing?**
1. Go to `chrome://extensions/`
2. Click the refresh icon on your extension card
3. Reload YouTube pages

**Can't see block button?**
1. Make sure you're on a channel page (not homepage)
2. Try scrolling down so the subscribe button is visible
3. Refresh the page

## Future Enhancements

Potential improvements:
- Export/import blocked channels list
- Block by keywords in channel names
- Whitelist mode (only show specified channels)
- Block entire categories/topics
- Sync across devices (requires Chrome sync storage)

## License

MIT License - Free to use and modify
