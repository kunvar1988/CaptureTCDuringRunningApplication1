# Enhanced Auto Test Case Capture Tool

A Python GUI application that **automatically captures test cases** with browser URL monitoring, automatic module detection, and organized Excel output.

## 🚀 New Enhanced Features

- **🌐 Browser URL Monitoring**: Automatically monitors browser URLs and navigation using Chrome DevTools Protocol
- **🎯 URL Filtering**: Only captures actions for pages with base URL: `https://qa-exchange.doceree.com`
- **🤖 Automatic Module Detection**: Automatically identifies modules from URL patterns:
  - Login
  - Advertiser Dashboard
  - Brand Dashboard
  - Manage Payments
  - Manage Users
  - Manage Accounts
  - Target
  - Plan
  - Activate
  - Measure
- **📊 Module-Based Organization**: Test cases are organized by module in separate Excel sheets
- **🔄 Automatic Navigation Capture**: Captures page navigation and URL changes automatically
- **📝 Smart Test Case Generation**: Automatically generates descriptions and expected results based on navigation

## Features

- **🔄 Automatic Action Capture**: Automatically monitors and captures:
  - Mouse clicks (left, right, middle)
  - Mouse scrolling
  - Browser URL changes and navigation
  - Window/application switches
  - Keyboard input (periodic capture)
- **⚡ Real-Time Capture**: Actions are captured instantly as you perform them
- **💾 Auto-Save Option**: Automatically saves test cases to Excel after every 5 actions
- **📝 Manual Override**: Add manual actions if needed
- **Auto-Export to Excel**: Automatically saves test cases to Excel organized by module
- **Formatted Excel Output**: Professional Excel formatting with:
  - Separate sheets for each module
  - Colored headers
  - Status-based color coding (Green for Pass, Red for Fail, Yellow for Blocked)
  - Auto-adjusted column widths
  - Text wrapping for long descriptions
  - Frozen header row
- **Two-Tab Interface**: 
  - **Setup Tab**: Start browser monitoring
  - **Auto Capture Tab**: Start monitoring and watch actions get captured automatically

## Installation

1. Make sure you have Python 3.6 or higher installed
2. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   
   **Note**: On Windows, you may need to install `pywin32` separately:
   ```bash
   pip install pywin32
   ```

## Setup for Browser URL Monitoring

To enable full URL monitoring, you need to start your browser with remote debugging enabled:

### Option 1: Use Helper Scripts (Recommended)

**For Chrome:**
1. Double-click `start_chrome_with_debugging.bat`
2. Chrome will start with remote debugging enabled
3. Navigate to `https://qa-exchange.doceree.com`

**For Microsoft Edge:**
1. Double-click `start_edge_with_debugging.bat`
2. Edge will start with remote debugging enabled
3. Navigate to `https://qa-exchange.doceree.com`

### Option 2: Manual Start

**For Chrome:**
```bash
chrome.exe --remote-debugging-port=9222 --user-data-dir="%TEMP%\chrome_debug_profile"
```

**For Microsoft Edge:**
```bash
msedge.exe --remote-debugging-port=9222 --user-data-dir="%TEMP%\edge_debug_profile"
```

**Important:** Close all existing Chrome/Edge instances before starting with remote debugging, or use a separate user data directory.

## Usage

### Step 1: Start Browser with Debugging

1. Close all Chrome/Edge browsers
2. Run `start_chrome_with_debugging.bat` or `start_edge_with_debugging.bat`
3. Navigate to `https://qa-exchange.doceree.com` in the browser

### Step 2: Start Test Case Capture Tool

1. Run the application:
   ```bash
   python test_case_capture.py
   ```

2. Go to **"Setup"** tab and click **"Start Browser Monitoring & Auto-Capture"**

3. The tool will start monitoring browser URLs automatically

### Step 3: Start Auto-Capture

1. Switch to **"Auto Capture"** tab (automatically switches after setup)

2. **Click "Start Auto-Capture"** button

3. **Start testing your application** - The tool will automatically:
   - Detect when you navigate to pages on `https://qa-exchange.doceree.com`
   - Identify the module and page from the URL
   - Capture all your actions (clicks, scrolling, typing)
   - Display current URL, module, and page in real-time

4. **Actions are automatically captured:**
   - `[10:30:15] Navigated to: https://qa-exchange.doceree.com/login` - URL navigation
   - `[10:30:16] Module: Login | Page: Login` - Module and page detection
   - `[10:30:18] Mouse Button.left click at (450, 320)` - Mouse clicks
   - `[10:30:20] Text input entered` - Keyboard input
   - All actions appear in real-time in the "Captured Actions" list

5. **Optional**: Enable "Auto-save after 5 actions" to automatically save test cases

6. After completing a test scenario:
   - Review captured actions (you can add manual actions if needed)
   - Fill in **Expected Result** (auto-generated if left empty)
   - Fill in **Actual Result** (what actually happened)
   - Select **Status** (Pass, Fail, Blocked, Not Executed)

7. Click **"Save Test Case to Excel"** (or it auto-saves if enabled)

8. The test case is immediately saved to the appropriate module sheet in `test_cases.xlsx`

9. Click **"Stop Auto-Capture"** when done, or **"Clear Actions"** to start next test case

## Module Detection

The tool automatically identifies modules from URL patterns:

| Module | URL Patterns |
|--------|-------------|
| Login | `/login`, `login`, `signin` |
| Advertiser Dashboard | `/advertiser`, `/dashboard` |
| Brand Dashboard | `/brand` |
| Manage Payments | `/payment`, `/payments` |
| Manage Users | `/user`, `/users` |
| Manage Accounts | `/account`, `/accounts` |
| Target | `/target`, `targeting` |
| Plan | `/plan`, `/planning` |
| Activate | `/activate`, `activation` |
| Measure | `/measure`, `/measurement`, `/analytics`, `/report` |

## Example Workflow

1. **Start Browser:**
   - Run `start_chrome_with_debugging.bat`
   - Navigate to `https://qa-exchange.doceree.com`

2. **Start Tool:**
   - Run `python test_case_capture.py`
   - Click "Start Browser Monitoring & Auto-Capture"
   - Click "Start Auto-Capture"

3. **Testing Login:**
   - Navigate to login page
   - Tool automatically detects: `Module: Login | Page: Login`
   - Perform login actions - all automatically captured
   - Expected Result: Auto-generated as "User should be successfully logged in and redirected to [page]"
   - Status: "Pass"
   - Click "Save Test Case to Excel"
   - Test case saved to "Login" sheet in Excel

4. **Testing Dashboard:**
   - Navigate to advertiser dashboard
   - Tool automatically detects: `Module: Advertiser Dashboard | Page: Dashboard`
   - Perform actions - all automatically captured
   - Save test case - saved to "Advertiser Dashboard" sheet

5. **Continue for other modules...**

## Output

Test cases are automatically saved to `test_cases.xlsx` in the same directory as the application. The Excel file includes:

- **Separate sheets for each module** (Login, Advertiser Dashboard, etc.)
- Professional formatting with colored headers
- Status-based color coding
- Auto-sized columns
- Text wrapping for readability
- Frozen header row for easy navigation
- All test cases with:
  - Test Case ID (format: `TC_MODULENAME_001`)
  - Test Case Name
  - Description
  - Preconditions (includes URL)
  - Test Steps
  - Expected Result
  - Actual Result
  - Status
  - Priority
  - Module
  - Page
  - URL
  - Created Date

## Key Benefits

- ✅ **Fully Automatic**: No manual entry needed - actions and URLs captured automatically
- ✅ **Real-time capture**: Actions and navigation appear instantly as you perform them
- ✅ **Zero interruption**: Test your application normally, capture happens in background
- ✅ **Auto-save option**: Test cases saved automatically after N actions
- ✅ **Manual override**: Add custom actions if automatic capture misses something
- ✅ **Immediate Excel export**: Test cases saved instantly to appropriate module sheet
- ✅ **Accumulative**: All test cases are saved in one Excel file, organized by module
- ✅ **URL Filtering**: Only captures actions for `https://qa-exchange.doceree.com`
- ✅ **Smart Organization**: Test cases automatically organized by module

## Notes

- **Browser Monitoring**: Requires Chrome/Edge to be started with `--remote-debugging-port=9222`
- **URL Filtering**: Only actions on pages with base URL `https://qa-exchange.doceree.com` are captured
- **Module Detection**: Modules are automatically identified from URL patterns
- **Privacy**: All monitoring happens locally - no data is sent anywhere
- **Performance**: Minimal impact on system performance
- Test cases are automatically exported to Excel after each save
- The Excel file accumulates all test cases (appends new ones)
- Test Case IDs are automatically generated with format: `TC_MODULENAME_001`, `TC_MODULENAME_002`, etc.
- You can remove captured actions by selecting them and clicking "Remove Selected"
- You can add manual actions if automatic capture doesn't capture something specific
- **Windows**: Window switching detection requires `pywin32` (included in requirements)

## Troubleshooting

### URL Monitoring Not Working

1. Make sure Chrome/Edge is started with remote debugging (use helper scripts)
2. Check that you're navigating to `https://qa-exchange.doceree.com`
3. Try restarting the browser with debugging enabled
4. Check firewall settings - port 9222 should be accessible locally

### Actions Not Being Captured

1. Make sure "Start Auto-Capture" is clicked
2. Verify you're on a page with base URL `https://qa-exchange.doceree.com`
3. Check that the browser window is active/focused

### Module Not Detected Correctly

- The tool uses URL patterns to detect modules
- If a module is not detected, you can manually edit the test case after saving
- Module patterns can be customized in the code (MODULE_PATTERNS dictionary)

## Requirements

- Python 3.6+
- openpyxl 3.1.2+
- pynput 1.7.6+
- psutil 5.9.5+
- selenium 4.0.0+
- requests 2.28.0+
- pywin32 (Windows only)

## License

This tool is provided as-is for test case management purposes.
