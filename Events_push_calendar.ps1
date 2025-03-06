# Microsoft 365 Group Event Pusher
$TenantId = 
$ClientId = 
$ClientSecret = 
$groupId = 
$excelFilePath = "/Users/username/Downloads/Meetings_name.xls"
$csvFilePath = "/Users/username/Downloads/Meetings_name.csv"
$pythonScriptPath = "/Users/username/Downloads/name_excel.py"
$logFilePath = "/Users/username/Downloads/EventPusher.log"

# Create Python conversion script
@"
#!/usr/bin/env python3

import sys
import pandas as pd

def convert_excel_to_csv(input_file, output_file):
    try:
        # Read the Excel file
        df = pd.read_excel(input_file)
        
        # Save as CSV
        df.to_csv(output_file, index=False)
        
        print(f"Successfully converted {input_file} to {output_file}")
        
        # Print column names for debugging
        print("Columns in the file:")
        print(df.columns.tolist())
        
        # Print first few rows
        print("\nFirst few rows:")
        print(df.head())
        
    except Exception as e:
        print(f"Error converting file: {e}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python convert_excel.py input.xls output.csv")
        sys.exit(1)
    
    convert_excel_to_csv(sys.argv[1], sys.argv[2])
"@ | Out-File -FilePath $pythonScriptPath -Encoding UTF8

# Logging Function
function Write-CustomLog {
    param (
        [string]$Message,
        [string]$LogPath = $logFilePath
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Add-Content -Path $LogPath -Value $logEntry
    Write-Host $logEntry
}

# Function to Get Access Token
function Get-MicrosoftGraphAccessToken {
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $env:CLIENTID
        client_secret = $env:CLIENT_SECRET
        scope         = "https://graph.microsoft.com/.default"
    }

    try {
        $response = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body
        return $response.access_token
    }
    catch {
        Write-CustomLog "Failed to obtain access token: $($_.Exception.Message)"
        throw
    }
}

# Function to Get Group Members
function Get-GroupMembers {
    param (
        [string]$AccessToken
    )

    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }

    try {
        $response = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/members" -Headers $headers
        return $response.value
    }
    catch {
        Write-CustomLog "Failed to retrieve group members: $($_.Exception.Message)"
        throw
    }
}

# Function to Check if Event Exists
function Test-EventExists {
    param (
        [string]$AccessToken,
        [string]$UserId,
        [string]$Subject,
        [datetime]$EventDate
    )

    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }

    $startDate = $EventDate.ToString("yyyy-MM-ddT00:00:00Z")
    $endDate = $EventDate.AddDays(1).ToString("yyyy-MM-ddT00:00:00Z")

    try {
        $filter = "`$filter=subject eq '$Subject' and start/dateTime ge '$startDate' and start/dateTime lt '$endDate'"
        $response = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/users/$UserId/calendar/events?$filter" -Headers $headers
        
        return ($response.value.Count -gt 0)
    }
    catch {
        Write-CustomLog ('Failed to check event existence for user {0}: {1}' -f $UserId, $_.Exception.Message)
        return $false
    }
}

# Function to Create Event
function New-CalendarEvent {
    param (
        [string]$AccessToken,
        [string]$UserId,
        [string]$Subject,
        [datetime]$EventDate
    )

    $headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }

    $eventBody = @{
        subject = $Subject
        start = @{
            dateTime = $EventDate.ToString("yyyy-MM-ddT00:00:00Z")
            timeZone = "UTC"
        }
        end = @{
            dateTime = $EventDate.AddDays(1).ToString("yyyy-MM-ddT00:00:00Z")
            timeZone = "UTC"
        }
        isAllDay = $true
        showAs = "free"
    } | ConvertTo-Json

    try {
        Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$UserId/calendar/events" -Headers $headers -Body $eventBody
        Write-CustomLog "Event '$Subject' created for user $UserId on $EventDate"
    }
    catch {
        Write-CustomLog ('Failed to check event existence for user {0}: {1}' -f $UserId, $_.Exception.Message)
    }
}

# Main Event Processing Function
function Process-Events {
    # Convert Excel to CSV
    try {
        & python3 $pythonScriptPath $excelFilePath $csvFilePath
    }
    catch {
        Write-CustomLog "Excel to CSV conversion failed: $($_.Exception.Message)"
        return
    }

    # Verify CSV file exists
    if (-not (Test-Path -Path $csvFilePath)) {
        Write-CustomLog "CSV file was not created"
        return
    }

    # Read CSV events
    $events = Import-Csv -Path $csvFilePath

    # Get Access Token
    $accessToken = Get-MicrosoftGraphAccessToken

    # Get Group Members
    $groupMembers = Get-GroupMembers -AccessToken $accessToken

    # Process each event
    foreach ($event in $events) {
        $subject = $event.Subject
        
       # Attempt to parse StartTime
    if (-not $event.StartTime -or $event.StartTime -eq "") {
        Write-CustomLog "Skipping event '$subject' due to missing StartTime."
        continue
    }

    try {
        $eventDate = [datetime]::ParseExact($event.StartTime, "yyyy-MM-dd", $null)
    }
    catch {
        Write-CustomLog "Failed to parse StartTime for event '$subject': $($_.Exception.Message)"
        continue
    }


        Write-CustomLog "Processing event: $subject on $eventDate"

        # Process for each group member
        foreach ($member in $groupMembers) {
            $userId = $member.id

            # Check if event already exists
            $eventExists = Test-EventExists -AccessToken $accessToken -UserId $userId -Subject $subject -EventDate $eventDate

            if ($eventExists) {
                Write-CustomLog "Event '$subject' already exists for user $userId. Skipping."
            }
            else {
                # Create event if it doesn't exist
                New-CalendarEvent -AccessToken $accessToken -UserId $userId -Subject $subject -EventDate $eventDate
            }
            Start-Sleep -Seconds 1
        }
    }
}

# Execute the event processing
Process-Events