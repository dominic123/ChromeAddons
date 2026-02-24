# SharePoint Field Creator - Chrome Extension

A Chrome extension for SharePoint 2019 on-premises that creates list fields from CSV files using JSOM (JavaScript Object Model).

## Features

- **Auto-Detect Configuration**: Automatically detects SharePoint site URL and list name from the current browser URL
- **CSV Upload**: Drag and drop or click to upload CSV files containing field definitions
- **Bulk Field Creation**: Create multiple fields in a SharePoint list at once
- **Progress Tracking**: Real-time progress bar and detailed results
- **Data Type Support**: Supports all common SharePoint field types

## Installation

### Load as Unpacked Extension (Development)

1. Open Chrome and navigate to `chrome://extensions/`
2. Enable "Developer mode" using the toggle in the top-right corner
3. Click "Load unpacked" button
4. Select the folder containing these extension files

### For Distribution

1. Package the extension files into a ZIP file
2. Distribute the ZIP file for users to unpack and load as unpacked extension
3. Or publish to Chrome Web Store (requires developer account)

## Usage

### Step 1: Navigate to SharePoint

1. Open your SharePoint 2019 on-premises site in Chrome
2. Navigate to any list (e.g., `https://yourserver/sites/sitename/Lists/YourList/AllItems.aspx`)

### Step 2: Open the Extension

1. Click the extension icon in Chrome's toolbar
2. The extension will automatically detect:
   - SharePoint Site URL
   - List Name (from current list page)

### Step 3: Connect to SharePoint

1. Click "Connect to SharePoint" button
2. Connection status will show as "Connected" if successful

### Step 4: Upload CSV File

**CSV Format:**
```
FieldName,DataType,DisplayName,Description,Required
EmployeeID,Text,Employee ID,Unique identifier,Yes
FirstName,Text,First Name,Employee first name,Yes
```

**Columns:**
- `FieldName`: Internal field name (no spaces, max 32 characters)
- `DataType`: Field type (see supported types below)
- `DisplayName: Name shown to users
- `Description`: Field description (optional)
- `Required`: "Yes" or "No" (optional, default: No)

**Upload Methods:**
- Drag and drop CSV file onto the drop zone
- Click the drop zone to browse for file

### Step 5: Create Fields

1. Preview the fields in the table
2. Click "Create Fields" button
3. Monitor progress and results

## Supported Data Types

| Type | Description |
|------|-------------|
| `Text` | Single line of text |
| `Note` | Multiple lines of text |
| `Number` | Decimal numbers |
| `Integer` | Whole numbers |
| `Currency` | Currency values |
| `DateTime` / `Date` | Date and time |
| `Boolean` / `YesNo` | Yes/No checkbox |
| `User` | Person or Group |
| `Lookup` | Lookup to another list |
| `Choice` | Choice field |
| `URL` / `Hyperlink` | Hyperlink or picture |

## Auto-Detection Rules

The extension automatically detects SharePoint information from various URL patterns:

| URL Pattern | Detection |
|-------------|-----------|
| `/Lists/ListName/AllItems.aspx` | Detects both site and list |
| `/Pages/Page.aspx?List=...` | Detects site and list |
| `/SitePages/Home.aspx` | Detects site only |
| `/_layouts/15/listedit.aspx?List=...` | Detects site from settings page |
| `/default.aspx` | Detects site home page |

## Troubleshooting

### "Not connected to SharePoint"
- Ensure you're on a SharePoint page
- Refresh the page and try again
- Check that SharePoint libraries are loaded

### "Connection failed"
- Verify you have permissions to the list
- Ensure the list name is correct
- Check browser console for detailed errors

### "Field already exists"
- The extension checks for existing fields
- Skip existing fields or use different field names

### Fields not visible after creation
- Refresh the SharePoint page
- Check list settings to verify field was created
- Some field types may require additional configuration

## File Structure

```
Sharepoint/
├── manifest.json           # Extension manifest (V3)
├── popup.html             # Popup UI
├── popup.js               # Popup logic and CSV parsing
├── styles.css             # UI styling
├── content.js             # JSOM execution in page context
├── background.js          # Service worker
├── sample_fields.csv      # Sample CSV file
├── icons/                 # Extension icons (create folder)
│   ├── icon16.png
│   ├── icon32.png
│   ├── icon48.png
│   └── icon128.png
└── README.md             # This file
```

## Creating Icons

Create or add icons to the `icons/` folder:
- `icon16.png` - 16x16 pixels
- `icon32.png` - 32x32 pixels
- `icon48.png` - 48x48 pixels
- `icon128.png` - 128x128 pixels

Or you can use any icon generator tool to create these icons.

## Permissions

The extension requires:
- `activeTab` - Access current tab information
- `storage` - Save configuration preferences
- `<all_urls>` - Work with any SharePoint site

## Browser Compatibility

- **Chrome/Edge**: Fully supported (Manifest V3)
- **SharePoint Version**: Designed for SharePoint 2019 on-premises
- **JSOM**: Uses SharePoint JavaScript Object Model

## Security Notes

- Extension only works when explicitly opened by user
- JSOM operations use current user's SharePoint permissions
- No data is sent to external servers
- All processing happens locally in browser

## Limitations

- Requires user to be logged into SharePoint
- Works only with sites accessible via current browser session
- Some field types (Lookup, Calculated) may need additional configuration after creation
- Cannot modify existing fields (only create new ones)

## Example CSV Files

### Simple Fields
```csv
FieldName,DataType,DisplayName,Description,Required
Title,Text,Title,Item title,Yes
Description,Note,Description,Detailed description,No
Status,Choice,Status,Select status,No
```

### Employee List
```csv
FieldName,DataType,DisplayName,Description,Required
EmployeeID,Text,Employee ID,Unique employee ID,Yes
FirstName,Text,First Name,Employee first name,Yes
LastName,Text,Last Name,Employee last name,Yes
Email,Text,Email,Email address,Yes
Department,Text,Department,Department name,No
Manager,User,Manager,Direct manager,No
HireDate,DateTime,Hire Date,Date hired,No
```

## License

This extension is provided as-is for SharePoint 2019 on-premises environments.

## Support

For issues or questions:
1. Check browser console (F12) for error messages
2. Verify SharePoint URL and list name
3. Ensure you have necessary permissions
4. Test with sample_fields.csv first
