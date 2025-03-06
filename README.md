# Microsoft 365 Group Calendar Event Script

## Overview
This PowerShell script automates the creation of calendar events for **all members of a Microsoft 365 group** using the **Microsoft Graph API**. It ensures that events are added to users' calendars without flooding their inboxes with email notifications.
The script ensures that duplicate events are not created and seamlessly syncs events for all users.

## Required Permissions
To run this script, you need the following Microsoft Graph API permissions:
- **Microsoft Graph API**
- `Group.Read.All`
- `GroupMember.Read.All`
# Microsoft 365 Group Event Pusher

## Description
This script automates the process of adding calendar events to all members of a Microsoft 365 group. Additionally, when a new member joins the group, they will receive an email notification prompting them to add all past events to their calendar. 

## Prerequisites
Before running the script, ensure that you have the following installed and configured:

### **1. Install Python 3 and Required Libraries**
The script includes a Python script that converts Excel files (.xls) to CSV format. You need to install Python and the required dependencies:
```bash
# Install Python 3 (if not installed)
brew install python  # macOS
sudo apt install python3  # Ubuntu/Debian

# Install required Python libraries
pip install pandas xlrd
```

### **2. Install PowerShell (if not installed)**
Ensure PowerShell is installed on your system:
- **Windows**: PowerShell is pre-installed.
- **macOS/Linux**: Install it using:
  ```bash
  brew install --cask powershell  # macOS
  sudo apt install powershell  # Ubuntu/Debian
  ```

### **3. Configure Microsoft Graph API**
This script requires access to the Microsoft Graph API. Follow these steps to set up authentication:
1. Register an application in **Microsoft Entra ID (Azure AD)**.
2. Note down the following credentials:
   - **Tenant ID**
   - **Client ID**
   - **Client Secret**

### **4. Enable Execution of PowerShell Scripts**
If running on Windows, allow execution of scripts:
```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```
For macOS/Linux, run PowerShell as:
```bash
pwsh -File script.ps1
```

## How the Script Works
### **1. Convert Excel to CSV**
- The Python script (`convert_excel.py`) reads an Excel file containing event details and converts it into a CSV file.

### **2. Fetch Microsoft 365 Group Members**
- The script retrieves all members of the specified Microsoft 365 group.

### **3. Process Events for Each Member**
- Reads the CSV file containing event details.
- Parses the `StartTime` field to ensure correct date format.
- Checks if the event already exists in the user’s calendar.
- If the event does not exist, it creates it.

## Usage
### **Run the Script**
```powershell
./script.ps1
```
Ensure the script has executable permissions if running on macOS/Linux:
```bash
chmod +x script.ps1
./script.ps1
```

### **Logging**
- The script logs all activities to `EventPusher.log` for troubleshooting.

## Future Enhancements
- Add support for event updates.
- Improve error handling and logging.
- Provide a web-based UI for easier event management.

## Author
Asaf Apel

