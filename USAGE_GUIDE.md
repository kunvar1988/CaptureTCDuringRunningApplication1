# Complete Usage Guide - Enhanced Auto Test Case Capture Tool

## 📋 Table of Contents
1. [Quick Start](#quick-start)
2. [Step-by-Step Workflow](#step-by-step-workflow)
3. [Features Explained](#features-explained)
4. [Best Practices](#best-practices)
5. [Troubleshooting](#troubleshooting)

---

## 🚀 Quick Start

### Prerequisites
1. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the application:
   ```bash
   python test_case_capture.py
   ```

---

## 📝 Step-by-Step Workflow

### Step 1: Start the Application
1. Open terminal/command prompt
2. Navigate to the project folder
3. Run: `python test_case_capture.py`
4. The application window will open with two tabs: **Setup** and **Auto Capture**

### Step 2: Set Up Browser Monitoring
1. Click on **"Setup"** tab
2. Click **"Start Browser Monitoring & Auto-Capture"** button
3. A message will appear confirming monitoring has started
4. You'll be automatically switched to the **"Auto Capture"** tab

### Step 3: Set the Current URL (Two Methods)

#### Method A: Manual Entry (Recommended for manually opened browsers)
1. Open your browser manually (Chrome/Edge/Firefox)
2. Navigate to: `https://qa-exchange.doceree.com/login` (or any page)
3. Copy the URL from the browser address bar (Ctrl+C)
4. In the tool, go to **"Manual Override"** section
5. Paste the URL in the URL field (Ctrl+V)
6. Click **"Set URL"** button
7. ✅ Module and Page will be automatically detected!

#### Method B: Auto-Detection (If browser started with debugging)
1. Start browser with debugging enabled (use helper .bat files)
2. Navigate to the target page
3. Click **"🔄 Update from Browser"** button
4. URL will be automatically detected

### Step 4: Start Capturing Actions
1. In **"Monitoring Controls"** section, click **"Start Auto-Capture"** button
2. Status will change to **"Status: Monitoring ON"** (green)
3. ✅ Now all your actions will be automatically captured!

### Step 5: Perform Your Testing
1. **Go to your browser** and start testing the application
2. **Perform actions** like:
   - Clicking buttons
   - Entering text
   - Navigating between pages
   - Scrolling
3. **Watch the tool** - Actions appear automatically in real-time!

### Step 6: Add Manual Actions (Optional)
If automatic capture misses something:
1. Type the action in **"Add Manual Action"** field
2. Click **"Add"** or press Enter
3. Or use **Quick Templates** buttons:
   - "Navigate to page"
   - "Click button"
   - "Enter text"
   - "Verify element"

### Step 7: Fill Test Case Details
1. **Expected Result**: What should happen (auto-generated if left empty)
2. **Actual Result**: What actually happened
3. **Status**: Select from dropdown:
   - Not Executed
   - Pass
   - Fail
   - Blocked

### Step 8: Save Test Case
1. Click **"Save Test Case to Excel"** button
2. Test case is saved to `test_cases.xlsx`
3. Test cases are organized by module in separate Excel sheets
4. ✅ Success message will appear!

### Step 9: Continue Testing
1. Click **"Clear Actions"** to start a new test case
2. Navigate to a new page (URL updates automatically)
3. Repeat steps 5-8 for more test cases

---

## 🎯 Features Explained

### 1. Current Browser State
- **URL**: Shows the currently detected/monitored URL
- **Module | Page**: Auto-detected module and page name
- **🔄 Update from Browser**: Button to manually refresh URL detection

### 2. Manual Override Section
Use this when:
- Browser was opened manually (not with debugging)
- Auto-detection isn't working
- You want to manually set URL/Module/Page

**Workflow:**
1. Copy URL from browser → 2. Paste here → 3. Click "Set URL" → 4. Module/Page auto-detected

### 3. Monitoring Controls
- **Status**: Shows if monitoring is ON (green) or OFF (red)
- **Start Auto-Capture**: Begins capturing mouse/keyboard actions
- **Stop Auto-Capture**: Stops capturing
- **Auto-save after 5 actions**: Automatically saves test cases every 5 actions

### 4. Automatically Captured Actions
- Shows all captured actions in real-time
- Each action has a timestamp
- Actions are numbered sequentially
- Scroll to see all actions

### 5. Quick Templates
Fast way to add common test actions:
- **Navigate to page**: Adds navigation step
- **Click button**: Adds click action
- **Enter text**: Adds text input action
- **Verify element**: Adds verification step

### 6. Activity Log & Processing Status
- **Real-time log** showing:
  - What the tool is doing
  - URL detection attempts
  - Action captures
  - Processing status
  - Errors and warnings
- **Color-coded messages** for easy reading
- **Clear Logs** button to reset

---

## 💡 Best Practices

### 1. Before Starting
- ✅ Set the URL first before starting auto-capture
- ✅ Verify module and page are correctly detected
- ✅ Check the Activity Log to see if monitoring is working

### 2. During Testing
- ✅ Keep the tool window visible to see captured actions
- ✅ Watch the Activity Log for real-time status
- ✅ Use Quick Templates for common actions
- ✅ Add manual actions if automatic capture misses something

### 3. After Testing
- ✅ Review captured actions before saving
- ✅ Fill in Expected and Actual Results
- ✅ Set appropriate Status (Pass/Fail/Blocked)
- ✅ Save test case to Excel

### 4. Organization Tips
- ✅ Test cases are automatically organized by module
- ✅ Each module has its own Excel sheet
- ✅ Test Case IDs are auto-generated (TC_MODULENAME_001, etc.)
- ✅ All test cases are saved in one Excel file

---

## 🔧 Troubleshooting

### Issue: URL Not Detected
**Solution:**
1. Use **Manual Override** section
2. Copy URL from browser and paste it
3. Click "Set URL"
4. Check Activity Log for detection attempts

### Issue: Actions Not Being Captured
**Possible Causes:**
1. **Monitoring not started**: Click "Start Auto-Capture"
2. **URL not set**: Set URL first (see above)
3. **Wrong browser window**: Make sure you're clicking in the target application
4. **Not on target URL**: Actions only captured for `https://qa-exchange.doceree.com`

**Check Activity Log** - it will show why actions aren't being captured!

### Issue: Module Not Detected Correctly
**Solution:**
1. Use **Manual Override** section
2. Select correct module from dropdown
3. Enter page name manually
4. Click "Set Module" and "Set Page"

### Issue: Browser Monitoring Not Working
**For Manually Opened Browsers:**
- This is normal! Use Manual Override section instead
- Copy/paste URL from browser

**For Auto-Detection:**
- Start browser with `--remote-debugging-port=9222`
- Use helper .bat files provided
- Check Activity Log for connection attempts

### Issue: Excel File Not Saving
**Check:**
1. Activity Log for errors
2. File permissions (make sure you can write to the folder)
3. Excel file is not open in another program

---

## 📊 Understanding the Activity Log

The Activity Log shows real-time status:

- **🟢 SUCCESS** (Green): Successful operations
- **🔵 INFO** (Cyan): General information
- **🟡 WARNING** (Yellow): Warnings (e.g., URL not matching)
- **🔴 ERROR** (Red): Errors that need attention
- **🔵 ACTION** (Blue): Actions being captured
- **🟠 URL** (Orange): URL detection events

**Example Log Messages:**
```
[19:16:42.123] [INFO] Starting browser monitoring...
[19:16:42.456] [SUCCESS] Browser monitoring started successfully
[19:16:50.789] [URL] URL detected: https://qa-exchange.doceree.com/login
[19:16:50.890] [SUCCESS] Module identified: Login | Page: Login
[19:17:15.234] [ACTION] Action captured: Mouse Button.left click at (450, 320)
[19:17:20.567] [ACTION] Action captured: Text input entered
```

---

## 🎓 Example Workflow

### Testing Login Functionality

1. **Setup:**
   - Start application
   - Click "Start Browser Monitoring"
   - Set URL: `https://qa-exchange.doceree.com/login`
   - Module auto-detected: "Login"

2. **Start Capturing:**
   - Click "Start Auto-Capture"
   - Status turns green: "Monitoring ON"

3. **Test Login:**
   - Go to browser
   - Click username field → ✅ Captured automatically
   - Type email → ✅ Captured automatically
   - Click password field → ✅ Captured automatically
   - Type password → ✅ Captured automatically
   - Click login button → ✅ Captured automatically

4. **After Login:**
   - URL changes to dashboard
   - Tool detects new URL automatically
   - Module changes to "Advertiser Dashboard"

5. **Save Test Case:**
   - Review captured actions (should show all steps)
   - Expected Result: "User should be logged in successfully"
   - Actual Result: "Login successful, redirected to dashboard"
   - Status: "Pass"
   - Click "Save Test Case to Excel"
   - ✅ Saved to "Login" module sheet!

---

## 📁 Output

Test cases are saved to **`test_cases.xlsx`** with:
- **Separate sheets** for each module (Login, Advertiser Dashboard, etc.)
- **Professional formatting** with colored headers
- **All test details**: ID, Name, Steps, Expected/Actual Results, Status, URL, etc.
- **Auto-organized** by module

---

## 🆘 Need Help?

1. **Check Activity Log** - It shows what's happening in real-time
2. **Read error messages** - They provide specific guidance
3. **Use Manual Override** - Works even if auto-detection fails
4. **Watch the status indicators** - Green = working, Red = stopped

---

## ✨ Tips & Tricks

1. **Quick URL Update**: Just copy URL and paste in Manual Override, then click "Set URL"
2. **Use Templates**: Click Quick Template buttons for common actions
3. **Auto-Save**: Enable "Auto-save after 5 actions" to save automatically
4. **Multiple Test Cases**: Click "Clear Actions" after saving to start a new test case
5. **Module Organization**: Test cases are automatically sorted by module in Excel

---

**Happy Testing! 🎉**
