# To the application you will need to give those permissions:
# Microsoft Graph and Group.Read.All and GroupMember.Read.All
# This script will be pushed to all members of the Microsoft 365 group events without getting a lot of emails in their mailboxes.

# Load Microsoft Graph Module
Import-Module Microsoft.Graph

# Define credentials and permissions
$TenantId = 
$ClientId = 
$ClientSecret = 
$CsvFilePath = 

# Convert client secret to secure credential
$SecureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $SecureClientSecret

# Connect to Microsoft Graph with required permissions
try {
    Connect-MgGraph -ClientSecretCredential $ClientSecretCredential `
                   -TenantId $TenantId `
                   -ErrorAction Stop
    
    Get-MgContext -ErrorAction Stop | Out-Null
    Write-Host "Successfully connected to Microsoft Graph API"
} catch {
    Write-Error "Failed to connect to Microsoft Graph API: $_"
    exit 1
}

# Function to create event in a user's calendar
function New-GraphCalendarEventForUser {
    param (
        [string]$Subject,
        [datetime]$StartTime,
        [datetime]$EndTime,
        [string[]]$AttendeeEmails,  # This will be the Microsoft 365 group email
        [string]$OrganizerEmail,
        [string]$TargetUserEmail,
        [string]$Location = "",
        [string]$Body = "",
        [string]$TimeZone = "UTC"
    )

    try {
        # Create base event parameters
        $EventParams = @{
            Subject = $Subject
            Start = @{
                DateTime = $StartTime.Date.ToString("yyyy-MM-dd")
                TimeZone = $TimeZone
            }
            End = @{
                DateTime = $StartTime.Date.AddDays(1).ToString("yyyy-MM-dd")
                TimeZone = $TimeZone
            }
            IsAllDay = $true
            ShowAs = "Free"
            ResponseRequested = $false
            Body = @{
                ContentType = "Text"
                Content = if ([string]::IsNullOrEmpty($Body)) { " " } else { $Body }
            }
        }

        if (-not [string]::IsNullOrEmpty($Location)) {
            $EventParams.Location = @{
                DisplayName = $Location
            }
        }

        # Process each attendee (Microsoft 365 group)
        foreach ($GroupEmail in $AttendeeEmails) {
            Write-Host "Processing Microsoft 365 Group: $GroupEmail" -ForegroundColor Yellow
            
            # Get group details
            $GroupName = $GroupEmail.Split('@')[0]
            $AllGroups = Get-MgGroup -All
            $Group = $AllGroups | Where-Object { 
                $_.Mail -eq $GroupEmail -or 
                $_.MailNickname -eq $GroupName
            }
            
            if ($Group) {
                Write-Host "Found group: $($Group.DisplayName) with ID: $($Group.Id)" -ForegroundColor Green
                
                # Get all members of the group
                $GroupMembers = Get-MgGroupMember -GroupId $Group.Id -All
                Write-Host "Found $($GroupMembers.Count) members in the group" -ForegroundColor Yellow
                
                foreach ($Member in $GroupMembers) {
                    # Get detailed user info
                    $UserDetails = Get-MgUser -UserId $Member.Id
                    if ($UserDetails -and $UserDetails.Mail) {
                        Write-Host "Creating event for group member: $($UserDetails.Mail)" -ForegroundColor Yellow
                        
                        # Create individual event for each member
                        $MemberEvent = New-MgUserEvent -UserId $UserDetails.Id -BodyParameter $EventParams
                        if ($MemberEvent) {
                            Write-Host "Event created for $Subject in $($UserDetails.Mail)'s calendar" -ForegroundColor Green
                        }
                    }
                }
            }
            else {
                Write-Warning "Group not found: $GroupEmail"
            }
        }

        # Create event for organizer
        # $OrganizerDetails = Get-MgUser -Filter "userPrincipalName eq '$OrganizerEmail'"
        # if ($OrganizerDetails) {
            # $NewEvent = New-MgUserEvent -UserId $OrganizerDetails.Id -BodyParameter $EventParams
            # if ($NewEvent) {
                # Write-Host "Event created for $Subject in organizer's calendar ($OrganizerEmail)" -ForegroundColor Green
            # }
        # }

        return $true
    }
    catch {
        Write-Error "Failed to create events: $_"
        Write-Error $_.Exception.Message
        return $false
    }
}

# Read CSV
if (-not (Test-Path -Path $CsvFilePath)) {
    Write-Error "CSV file not found: $CsvFilePath"
    exit 1
}

try {
    $EventData = Import-Csv -Path $CsvFilePath -ErrorAction Stop
    if ($EventData.Count -eq 0) { Write-Error "No events found in the CSV file"; exit 1 }
} catch { Write-Error "Failed to import CSV: $_"; exit 1 }

# Process events
foreach ($Event in $EventData) {
    try {
        # Add more flexible date parsing formats
        $DateFormats = @(
            "dd/MM/yyyy HH:mm",
            "dd/MM/yyyy H:mm",
            "d/M/yyyy HH:mm",
            "d/M/yyyy H:mm",
            "dd/MM/yyyy",  # For dates without time
            "d/M/yyyy"     # For dates without time
        )

        $StartTime = if ($Event.StartTime -match '\d{2}:\d{2}$') {
            [DateTime]::ParseExact(
                $Event.StartTime, 
                $DateFormats, 
                [System.Globalization.CultureInfo]::InvariantCulture, 
                [System.Globalization.DateTimeStyles]::None
            )
        } else {
            # If no time specified, use ParseExactMultiple instead
            [DateTime]::ParseExact(
                $Event.StartTime,
                "dd/MM/yyyy",  # Single format for dates without time
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None
            ).AddHours(9)
        }

        $EndTime = if ($Event.EndTime -match '\d{2}:\d{2}$') {
            [DateTime]::ParseExact(
                $Event.EndTime, 
                $DateFormats, 
                [System.Globalization.CultureInfo]::InvariantCulture, 
                [System.Globalization.DateTimeStyles]::None
            )
        } else {
            # If no time specified, use ParseExactMultiple instead
            [DateTime]::ParseExact(
                $Event.EndTime,
                "dd/MM/yyyy",  # Single format for dates without time
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None
            ).AddHours(10)
        }

        $AttendeeEmails = $Event.AttendeeEmails -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }

        New-GraphCalendarEventForUser `
            -Subject $Event.Subject `
            -StartTime $StartTime `
            -EndTime $EndTime `
            -AttendeeEmails $AttendeeEmails `
            -OrganizerEmail $OrganizerEmail `
            -TargetUserEmail $OrganizerEmail `
            -Location $Event.Location `
            -Body $Event.Body

    } catch {
        Write-Error "Failed to process event $($Event.Subject): $_"
        continue
    }
}

# Disconnect from Microsoft Graph when done
Disconnect-MgGraph