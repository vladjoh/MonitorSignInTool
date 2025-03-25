#Requires -RunAsAdministrator

# Script configuration
$config = @{
    ClientId         = ""
    ClientSecret     = ""
    TenantId         = ""
    SenderEmail      = ""
    RecipientEmail   = ""
    MonitoredAccount = ""  # Object ID works best
    MonitoredAccountName = ""  # Display name for alerts
    TimeWindowMinutes = 10  # Look back this many minutes for sign-ins (reduced from 24 hours)
    StateFilePath    = "C:\temp\SignInMonitor_state.json"
}

# Log file path
$logFile = "C:\temp\MonitorSignIn_log.txt"

# Function to write to log file
function Write-Log {
    param (
        [string]$Message
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -Append -FilePath $logFile
    Write-Host "$timestamp - $Message"
}

# Function to get access token for Microsoft Graph API
function Get-AccessToken {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $body = @{
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $ClientId
        client_secret = $ClientSecret
    }

    try {
        $response = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $body
        Write-Log "Access token obtained successfully."
        return $response.access_token
    }
    catch {
        Write-Log "Failed to get access token: $_"
        exit 1
    }
}

# Function to get sign-in events by object ID - optimized for quicker response
function Get-SignInEvents {
    param (
        [string]$AccessToken,
        [int]$MinutesBack = 10,
        [string]$UserId
    )

    $headers = @{
        Authorization = "Bearer $AccessToken"
    }

    try {
        # Get sign-ins for the specified time window using userId
        $lookBackTime = [DateTime]::UtcNow.AddMinutes(-$MinutesBack)
        $timeFilter = [System.Web.HttpUtility]::UrlEncode("createdDateTime ge $($lookBackTime.ToString('yyyy-MM-ddTHH:mm:ssZ')) and userId eq '$UserId'")
        
        $uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=$timeFilter&`$top=50&`$orderby=createdDateTime desc"
        Write-Log "Fetching sign-ins with URL: $uri"
        
        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
        Write-Log "Retrieved $($response.value.Count) sign-in events for user ID $UserId."
        return $response.value
    }
    catch {
        Write-Log "Failed to get sign-in events: $_"
        return $null
    }
}

# Function to read the last processed state
function Get-LastProcessedState {
    param (
        [string]$StateFilePath
    )
    
    if (Test-Path $StateFilePath) {
        try {
            $state = Get-Content $StateFilePath -Raw | ConvertFrom-Json
            return $state
        }
        catch {
            Write-Log "Error reading state file: $_"
            return @{
                LastRunTime = [DateTime]::UtcNow.AddMinutes(-10).ToString('o')
                ProcessedIds = @()
            }
        }
    }
    else {
        # Return default state if file doesn't exist
        return @{
            LastRunTime = [DateTime]::UtcNow.AddMinutes(-10).ToString('o')
            ProcessedIds = @()
        }
    }
}

# Function to save the processed state
function Save-ProcessedState {
    param (
        [string]$StateFilePath,
        [object]$State
    )
    
    try {
        $State | ConvertTo-Json | Out-File $StateFilePath -Force
        Write-Log "Saved state file with $($State.ProcessedIds.Count) processed IDs."
    }
    catch {
        Write-Log "Error saving state file: $_"
    }
}

# Function to format fields for better readability
function Format-SignInField {
    param(
        [Parameter(Mandatory=$false)]
        $Value,
        
        [Parameter(Mandatory=$false)]
        [string]$DefaultValue = "Not available"
    )
    
    if ($null -eq $Value -or $Value -eq "" -or $Value -eq 0) {
        return $DefaultValue
    }
    
    return $Value
}

# Function to send email notification using REST API directly
function Send-EmailNotification {
    param (
        [string]$AccessToken,
        [string]$SenderEmail,
        [string]$RecipientEmail,
        [string]$Subject,
        [string]$Body
    )

    try {
        # Create the message payload
        $messageJson = @{
            message = @{
                subject = $Subject
                body = @{
                    contentType = "HTML"
                    content = $Body
                }
                toRecipients = @(
                    @{
                        emailAddress = @{
                            address = $RecipientEmail
                        }
                    }
                )
            }
            saveToSentItems = $false
        } | ConvertTo-Json -Depth 4
        
        # Set up request headers
        $headers = @{
            "Authorization" = "Bearer $AccessToken"
            "Content-Type" = "application/json"
        }
        
        # Send the email using REST API
        Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$SenderEmail/sendMail" -Headers $headers -Body $messageJson
        Write-Log "Email notification sent successfully."
        return $true
    }
    catch {
        Write-Log "Failed to send email via Graph API: $_"
        
        # Try using Send-MailMessage as last resort
        try {
            Send-MailMessage -From $SenderEmail -To $RecipientEmail -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer "smtp.office365.com" -Port 587 -UseSsl -Credential (New-Object System.Management.Automation.PSCredential -ArgumentList $SenderEmail, (ConvertTo-SecureString -String $config.ClientSecret -AsPlainText -Force))
            Write-Log "Email notification sent successfully using Send-MailMessage."
            return $true
        }
        catch {
            Write-Log "All email notification methods failed: $_"
            return $false
        }
    }
}

# Function to create HTML email body
function Get-SignInAlertEmailBody {
    param (
        [PSCustomObject]$SignIn,
        [string]$UserDisplayName,
        [string]$UserPrincipalName
    )
    
    # Format the data
    $timeFormatted = try {
        [DateTime]::Parse($SignIn.createdDateTime).ToString("yyyy-MM-dd HH:mm:ss")
    } catch {
        $SignIn.createdDateTime
    }
    
    $ipAddress = Format-SignInField -Value $SignIn.ipAddress -DefaultValue "Unknown"
    $location = Format-SignInField -Value $(if ($SignIn.location -and $SignIn.location.city) { $SignIn.location.city } else { "" }) -DefaultValue "Unknown"
    $country = Format-SignInField -Value $(if ($SignIn.location -and $SignIn.location.countryOrRegion) { $SignIn.location.countryOrRegion } else { "" }) -DefaultValue "Unknown"
    $status = if ($SignIn.status -and $SignIn.status.errorCode -eq 0) { "<span style='color:green'>Successful</span>" } else { "<span style='color:red'>Failed (Code: $($SignIn.status.errorCode))</span>" }
    $clientApp = Format-SignInField -Value $SignIn.clientAppUsed -DefaultValue "Unknown"
    $device = Format-SignInField -Value $(if ($SignIn.deviceDetail -and $SignIn.deviceDetail.displayName) { $SignIn.deviceDetail.displayName } else { "" }) -DefaultValue "Unknown"
    $browser = Format-SignInField -Value $(if ($SignIn.deviceDetail -and $SignIn.deviceDetail.browser) { $SignIn.deviceDetail.browser } else { "" }) -DefaultValue "Unknown"
    $os = Format-SignInField -Value $(if ($SignIn.deviceDetail -and $SignIn.deviceDetail.operatingSystem) { $SignIn.deviceDetail.operatingSystem } else { "" }) -DefaultValue "Unknown"
    
    # HTML formatted email
    $body = @"
<html>
<body style='font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px;'>
    <h2 style='color: #d9534f; text-align: center; padding: 10px; background-color: #f9f9f9; border-radius: 3px;'>⚠️ SECURITY ALERT - ACCOUNT SIGN-IN DETECTED ⚠️</h2>
    
    <table style='width: 100%; border-collapse: collapse; margin-top: 20px;'>
        <tr>
            <td style='font-weight: bold; padding: 8px; border-bottom: 1px solid #ddd;'>Account Name:</td>
            <td style='padding: 8px; border-bottom: 1px solid #ddd;'>$UserDisplayName</td>
        </tr>
        <tr>
            <td style='font-weight: bold; padding: 8px; border-bottom: 1px solid #ddd;'>Email:</td>
            <td style='padding: 8px; border-bottom: 1px solid #ddd;'>$UserPrincipalName</td>
        </tr>
        <tr>
            <td style='font-weight: bold; padding: 8px; border-bottom: 1px solid #ddd;'>Time (UTC):</td>
            <td style='padding: 8px; border-bottom: 1px solid #ddd;'>$timeFormatted</td>
        </tr>
        <tr>
            <td style='font-weight: bold; padding: 8px; border-bottom: 1px solid #ddd;'>Status:</td>
            <td style='padding: 8px; border-bottom: 1px solid #ddd;'>$status</td>
        </tr>
        <tr>
            <td style='font-weight: bold; padding: 8px; border-bottom: 1px solid #ddd;'>IP Address:</td>
            <td style='padding: 8px; border-bottom: 1px solid #ddd;'>$ipAddress</td>
        </tr>
        <tr>
            <td style='font-weight: bold; padding: 8px; border-bottom: 1px solid #ddd;'>Location:</td>
            <td style='padding: 8px; border-bottom: 1px solid #ddd;'>$location, $country</td>
        </tr>
        <tr>
            <td style='font-weight: bold; padding: 8px; border-bottom: 1px solid #ddd;'>Client App:</td>
            <td style='padding: 8px; border-bottom: 1px solid #ddd;'>$clientApp</td>
        </tr>
        <tr>
            <td style='font-weight: bold; padding: 8px; border-bottom: 1px solid #ddd;'>Device:</td>
            <td style='padding: 8px; border-bottom: 1px solid #ddd;'>$device</td>
        </tr>
        <tr>
            <td style='font-weight: bold; padding: 8px; border-bottom: 1px solid #ddd;'>Browser:</td>
            <td style='padding: 8px; border-bottom: 1px solid #ddd;'>$browser</td>
        </tr>
        <tr>
            <td style='font-weight: bold; padding: 8px; border-bottom: 1px solid #ddd;'>Operating System:</td>
            <td style='padding: 8px; border-bottom: 1px solid #ddd;'>$os</td>
        </tr>
        <tr>
            <td style='font-weight: bold; padding: 8px; border-bottom: 1px solid #ddd;'>Sign-in ID:</td>
            <td style='padding: 8px; border-bottom: 1px solid #ddd;'>$($SignIn.id)</td>
        </tr>
    </table>
    
    <div style='margin-top: 20px; padding: 10px; background-color: #f2dede; border: 1px solid #ebccd1; border-radius: 3px; color: #a94442;'>
        <strong>If this was not you, please contact your system administrator immediately!</strong>
    </div>
</body>
</html>
"@

    return $body
}

# Main function to monitor sign-ins
function Monitor-SignIns {
    param (
        [hashtable]$Config
    )
    
    Write-Log "Starting sign-in monitoring for account $($Config.MonitoredAccount)..."
    
    # Get access token
    $accessToken = Get-AccessToken -TenantId $Config.TenantId -ClientId $Config.ClientId -ClientSecret $Config.ClientSecret
    
    # Get current state
    $state = Get-LastProcessedState -StateFilePath $Config.StateFilePath
    Write-Log "Last run time: $($state.LastRunTime), Processed IDs count: $($state.ProcessedIds.Count)"
    
    # Get sign-in events for the monitored account
    $signInEvents = Get-SignInEvents -AccessToken $accessToken -MinutesBack $Config.TimeWindowMinutes -UserId $Config.MonitoredAccount
    
    # Use hardcoded name or try to get UPN from sign-in events
    $userDisplayName = $Config.MonitoredAccountName
    $userPrincipalName = $Config.MonitoredAccount  # Default to object ID
    
    if ($signInEvents -and $signInEvents.Count -gt 0 -and $signInEvents[0].userPrincipalName) {
        $userPrincipalName = $signInEvents[0].userPrincipalName
        Write-Log "Using UPN from sign-in events: $userPrincipalName"
    }
    
    if ($signInEvents -and $signInEvents.Count -gt 0) {
        Write-Log "Found $($signInEvents.Count) sign-ins for monitored account."
        
        # Track new processed IDs
        $newProcessedIds = @()
        $newSignInsDetected = $false
        
        # Process sign-ins in chronological order (oldest first)
        foreach ($signIn in ($signInEvents | Sort-Object createdDateTime)) {
            # Check if we've already processed this sign-in
            if ($state.ProcessedIds -contains $signIn.id) {
                Write-Log "Sign-in ID $($signIn.id) already processed, skipping."
                $newProcessedIds += $signIn.id
                continue
            }
            
            $newSignInsDetected = $true
            $newProcessedIds += $signIn.id
            
            # Use the current sign-in's UPN if available
            $currentUPN = if ($signIn.userPrincipalName) { $signIn.userPrincipalName } else { $userPrincipalName }
            
            # Get formatted timestamp for logging
            $timeFormatted = try {
                [DateTime]::Parse($signIn.createdDateTime).ToString("yyyy-MM-dd HH:mm:ss")
            } catch {
                $signIn.createdDateTime
            }
            
            # Create email body
            $emailBody = Get-SignInAlertEmailBody -SignIn $signIn -UserDisplayName $userDisplayName -UserPrincipalName $currentUPN
            
            # Send the notification
            $emailSubject = "⚠️ Security Alert - Sign-In Detected for $userDisplayName"
            $emailSent = Send-EmailNotification -AccessToken $accessToken -SenderEmail $Config.SenderEmail -RecipientEmail $Config.RecipientEmail -Subject $emailSubject -Body $emailBody
            
            if ($emailSent) {
                Write-Log "Alert sent for sign-in at $timeFormatted from $($signIn.ipAddress) (ID: $($signIn.id))"
            } else {
                Write-Log "Failed to send alert for sign-in at $timeFormatted from $($signIn.ipAddress) (ID: $($signIn.id))"
            }
        }
        
        # Update state to include all IDs we've seen (including previously processed ones)
        $allProcessedIds = @($state.ProcessedIds) + @($newProcessedIds) | Select-Object -Unique
        
        # Limit the number of stored IDs to prevent file growth
        if ($allProcessedIds.Count -gt 100) {
            $allProcessedIds = $allProcessedIds | Select-Object -Last 100
        }
        
        $newState = @{
            LastRunTime = [DateTime]::UtcNow.ToString('o')
            ProcessedIds = $allProcessedIds
        }
        
        Save-ProcessedState -StateFilePath $Config.StateFilePath -State $newState
        
        if (-not $newSignInsDetected) {
            Write-Log "No new sign-ins detected for the monitored account."
        }
    } else {
        Write-Log "No sign-ins detected for the monitored account in the last $($Config.TimeWindowMinutes) minutes."
        
        # Update last run time but keep processed IDs
        $newState = @{
            LastRunTime = [DateTime]::UtcNow.ToString('o')
            ProcessedIds = $state.ProcessedIds
        }
        
        Save-ProcessedState -StateFilePath $Config.StateFilePath -State $newState
    }
    
    Write-Log "Monitoring complete."
}

# Make sure we can use HttpUtility
Add-Type -AssemblyName System.Web

# Only check for required modules the first time
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Reports")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Log "Installing $module module..."
        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $module -ErrorAction SilentlyContinue
}
Write-Log "Required modules verified."

# Run the monitor function with the configuration
Monitor-SignIns -Config $config