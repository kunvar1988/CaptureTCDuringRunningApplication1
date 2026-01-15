# Quick Start Guide

## Step-by-Step Instructions

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Start Browser with Debugging

**Option A: Use Helper Script (Easiest)**
- Double-click `start_chrome_with_debugging.bat` (for Chrome)
- OR double-click `start_edge_with_debugging.bat` (for Edge)

**Option B: Manual Start**
- Close all Chrome/Edge browsers
- Open Command Prompt
- Run:
  ```
  chrome.exe --remote-debugging-port=9222 --user-data-dir="%TEMP%\chrome_debug_profile"
  ```
  OR
  ```
  msedge.exe --remote-debugging-port=9222 --user-data-dir="%TEMP%\edge_debug_profile"
  ```

### 3. Navigate to Application
- In the browser, navigate to: `https://qa-exchange.doceree.com`
- Login or access the pages you want to test

### 4. Start Test Case Capture Tool
```bash
python test_case_capture.py
```

### 5. Setup
- Click **"Start Browser Monitoring & Auto-Capture"** button
- Tool will start monitoring browser URLs

### 6. Start Capturing
- Switch to **"Auto Capture"** tab
- Click **"Start Auto-Capture"** button
- Start testing your application normally

### 7. Watch Magic Happen! ✨
- As you navigate, the tool automatically:
  - Detects URL changes
  - Identifies module and page from URL
  - Captures all your actions (clicks, typing, scrolling)
  - Displays everything in real-time

### 8. Save Test Cases
- After performing actions, click **"Save Test Case to Excel"**
- Test case is automatically saved to the correct module sheet
- Or enable **"Auto-save after 5 actions"** for automatic saving

## What Gets Captured?

✅ **Automatically Captured:**
- Browser URL changes and navigation
- Module identification (Login, Dashboard, Payments, etc.)
- Page identification from URL
- Mouse clicks
- Mouse scrolling
- Keyboard input (periodic)
- Window switches

✅ **Auto-Generated:**
- Test Case ID (format: TC_MODULENAME_001)
- Module name (from URL)
- Page name (from URL)
- Description
- Expected Result (based on navigation)
- Preconditions (includes URL)

## Excel Output

Test cases are saved to `test_cases.xlsx` with:
- **Separate sheet for each module** (Login, Advertiser Dashboard, etc.)
- Professional formatting
- All test case details
- URL included in preconditions

## Tips

1. **Keep browser with debugging enabled** - Don't close it during testing
2. **Navigate normally** - The tool captures everything automatically
3. **Check current URL/Module** - Always visible in the "Current Browser State" section
4. **Auto-save is helpful** - Enable it to save test cases automatically
5. **Manual actions** - You can add manual actions if needed

## Troubleshooting

**URL not detected?**
- Make sure browser was started with `--remote-debugging-port=9222`
- Check that you're on `https://qa-exchange.doceree.com`
- Try restarting browser with debugging

**Actions not captured?**
- Make sure "Start Auto-Capture" is clicked
- Verify you're on the target application (not other windows)
- Check that browser window is active

**Module not detected?**
- The tool uses URL patterns to detect modules
- Some URLs might not match patterns - you can manually edit after saving
