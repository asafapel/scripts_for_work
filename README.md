# Microsoft 365 Group Calendar Event Script

## Overview
This PowerShell script automates the creation of calendar events for **all members of a Microsoft 365 group** using the **Microsoft Graph API**. It ensures that events are added to users' calendars without flooding their inboxes with email notifications.

## Required Permissions
To run this script, you need the following Microsoft Graph API permissions:
- **Microsoft Graph API**
- `Group.Read.All`
- `GroupMember.Read.All`

## Prerequisites
Before running the script, ensure that you have:
1. Installed the **Microsoft Graph PowerShell SDK**
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```
2. Registered an **Azure AD Application** with the required permissions.
3. Generated a **Client Secret** for authentication.
4. Assigned **Admin Consent** to the required permissions.
5. Created a CSV file containing the event details.

## Installation & Usage
### 1️⃣ Clone the Repository
```sh
git clone https://github.com/your-username/your-repository.git
cd your-repository
```

### 2️⃣ Configure Authentication
Update the script with your Azure AD credentials:
```powershell
$TenantId = "your-tenant-id"
$ClientId = "your-client-id"
$ClientSecret = "your-client-secret"
$CsvFilePath = "path-to-your-csv-file.csv"
```

### Prepare the CSV File
Ensure your CSV file follows this format:
| Subject | StartTime | EndTime | AttendeeEmails | IsAllDay
|---------|----------|---------|----------------|----------|------|
| Meeting | 01/03/2024 10:00 | 01/03/2024 12:00 | group@example.com | True/False

### Run the Script
```powershell
.\calendar_events_script.ps1
```

### Verify Events
After execution, check Microsoft 365 users' calendars to confirm the event creation.

## Features
- **Bulk event creation** for Microsoft 365 group members.
- **Secure authentication** via Microsoft Graph API.
- **Customizable event details** including subject, time, location, and attendees.
- **Error handling and logging** for efficient troubleshooting.

## Troubleshooting
If you encounter errors, consider:
- Verifying the **Azure AD App permissions**.
- Checking that the **CSV file is correctly formatted**.
- Running `Get-MgContext` to confirm successful authentication.

## Contributing
Feel free to submit pull requests for improvements or bug fixes!

## Author
  Asaf Apel

