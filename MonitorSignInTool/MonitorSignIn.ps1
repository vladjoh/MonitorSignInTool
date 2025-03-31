# Script configuration
$config = @{
    ClientId              = ""
    CertificateThumbprint = "" 
    TenantId              = ""
    SenderEmail           = ""
    RecipientEmail        = ""
    MonitoredAccount      = ""  # Using Object ID
    MonitoredAccountName  = ""
    TimeWindowMinutes     = 60
    StateFilePath         = "C:\temp\SignInMonitor_state.json"
}

# Log file path
$logFile = "C:\temp\MonitorSignIn_log.txt"

# Function to write to log file with color
function Write-Log {
    param (
        [string]$Message,
        [switch]$Error,
        [switch]$Success
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -Append -FilePath $logFile
    
    if ($Error) {
        Write-Host "$timestamp - $Message" -ForegroundColor Red
    }
    elseif ($Success) {
        Write-Host "$timestamp - $Message" -ForegroundColor Green
    }
    else {
        Write-Host "$timestamp - $Message"
    }
}

# Function to get access token using certificate authentication
function Get-AccessToken {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint
    )

    try {
        # Get certificate from local store
        $certificate = Get-ChildItem -Path "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction Stop
        
        if (-not $certificate) {
            throw "Certificate with thumbprint $CertificateThumbprint not found"
        }
        
        Write-Log "Found certificate: $($certificate.Subject), Issued by: $($certificate.Issuer)"
        
        # Get token using MSAL.PS module
        if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
            Write-Log "Installing MSAL.PS module..."
            Install-Module -Name MSAL.PS -Scope CurrentUser -Force -AllowClobber -Confirm:$false
        }
        
        Import-Module MSAL.PS -ErrorAction Stop
        
        $msalParams = @{
            ClientId = $ClientId
            TenantId = $TenantId
            ClientCertificate = $certificate
            Scopes = "https://graph.microsoft.com/.default"
        }
        
        $authResult = Get-MsalToken @msalParams
        Write-Log "Access token obtained successfully." -Success
        return $authResult.AccessToken
    }
    catch {
        Write-Log "Failed to get access token using certificate: $_" -Error
        
        # Log detailed error for certificate retrieval
        try {
            $certsAvailable = Get-ChildItem -Path "Cert:\LocalMachine\My" | Select-Object -Property Subject, Thumbprint
            Write-Log "Available certificates: $($certsAvailable | ConvertTo-Json)" -Error
        } catch {
            Write-Log "Could not list available certificates: $_" -Error
        }
        
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
        
        if ($response.value.Count -gt 0) {
            # Count successful and failed sign-ins
            $successfulSignIns = 0
            $failedSignIns = 0
            
            foreach ($signIn in $response.value) {
                if ($signIn.status -and $signIn.status.errorCode -eq 0) {
                    $successfulSignIns++
                } else {
                    $failedSignIns++
                }
            }
            
            $message = "Retrieved $($response.value.Count) sign-in event"
            if ($response.value.Count -gt 1) {
                $message += "s"
            }
            $message += ":"
            
            if ($successfulSignIns -gt 0) {
                $message += " $successfulSignIns successful sign-in"
                if ($successfulSignIns -gt 1) {
                    $message += "s"
                }
                
                if ($failedSignIns -gt 0) {
                    $message += " and"
                }
            }
            
            if ($failedSignIns -gt 0) {
                $message += " $failedSignIns failed sign-in attempt"
                if ($failedSignIns -gt 1) {
                    $message += "s"
                }
            }
            
            Write-Log $message -Success
        }
        
        return $response.value
    }
    catch {
        Write-Log "Failed to get sign-in events: $_" -Error
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
            Write-Log "Error reading state file: $_" -Error
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
        Write-Log "Error saving state file: $_" -Error
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

# Updated email sending function that should fix single email delivery issues
function Send-EmailNotification {
    param (
        [string]$AccessToken,
        [string]$SenderEmail,
        [string]$RecipientEmail,
        [string]$Subject,
        [string]$Body
    )

    try {
        # Add a timestamp to subject line to help ensure uniqueness
        $uniqueSubject = "$Subject - [$(Get-Date -Format 'HH:mm:ss.fff')]"
        
        # Create the message payload with simpler format
        $messageJson = @{
            message = @{
                subject = $uniqueSubject
                importance = "high"  # Add high importance flag
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
            saveToSentItems = $true  # Changed to save to sent items for verification
        } | ConvertTo-Json -Depth 4
        
        # Set up request headers
        $headers = @{
            "Authorization" = "Bearer $AccessToken"
            "Content-Type" = "application/json"
            "Prefer" = "outlook.timezone=""UTC"""
        }
        
        # Send the email using REST API with additional parameters
        $apiVersion = "v1.0"  # Using v1.0 for maximum compatibility
        $uri = "https://graph.microsoft.com/$apiVersion/users/$SenderEmail/sendMail"
        
        Write-Log "Sending email to $RecipientEmail with subject '$uniqueSubject'"
        
        # Using Invoke-RestMethod with explicit parameters
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $messageJson -ContentType "application/json" -ErrorAction Stop
        
        # Force a small delay to ensure processing
        Start-Sleep -Seconds 1
        
        Write-Log "Email notification sent successfully." -Success
        return $true
    }
    catch {
        Write-Log "Failed to send email via Graph API: $_" -Error
        Write-Log "Request failed with error details: $($_.ErrorDetails)" -Error
        return $false
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
    
    # Determine if this was a successful sign-in or a failed attempt
    $isSuccessful = $SignIn.status -and $SignIn.status.errorCode -eq 0
    
    if ($isSuccessful) {
        $alertType = "SUCCESSFUL SIGN-IN DETECTED"
        $status = "<span style='color:green'>Successful</span>"
    } else {
        $alertType = "FAILED SIGN-IN ATTEMPT DETECTED"
        $status = "<span style='color:red'>Failed (Code: $($SignIn.status.errorCode))</span>"
    }
    
    $clientApp = Format-SignInField -Value $SignIn.clientAppUsed -DefaultValue "Unknown"
    $device = Format-SignInField -Value $(if ($SignIn.deviceDetail -and $SignIn.deviceDetail.displayName) { $SignIn.deviceDetail.displayName } else { "" }) -DefaultValue "Unknown"
    $browser = Format-SignInField -Value $(if ($SignIn.deviceDetail -and $SignIn.deviceDetail.browser) { $SignIn.deviceDetail.browser } else { "" }) -DefaultValue "Unknown"
    $os = Format-SignInField -Value $(if ($SignIn.deviceDetail -and $SignIn.deviceDetail.operatingSystem) { $SignIn.deviceDetail.operatingSystem } else { "" }) -DefaultValue "Unknown"
    
    # HTML formatted email with updated alert type - simplified for maximum compatibility
    $body = @"
<html>
<body style='font-family: Arial, sans-serif; margin: 0 auto; padding: 20px;'>
    <h2 style='color: #d9534f; text-align: center;'>⚠️ SECURITY ALERT - $alertType ⚠️</h2>
    
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
    
    <div style='margin-top: 20px; padding: 10px; background-color: #f2dede; color: #a94442;'>
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
    
    # Get access token using certificate authentication
    $accessToken = Get-AccessToken -TenantId $Config.TenantId -ClientId $Config.ClientId -CertificateThumbprint $Config.CertificateThumbprint
    
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
        # Filter out sign-ins we've already processed
        $newSignIns = $signInEvents | Where-Object { $state.ProcessedIds -notcontains $_.id }
        
        # Only show the "Found" message for new sign-ins
        if ($newSignIns.Count -gt 0) {
            # Count successful and failed sign-ins for the log message
            $successfulSignIns = 0
            $failedSignIns = 0
            
            foreach ($signIn in $newSignIns) {
                if ($signIn.status -and $signIn.status.errorCode -eq 0) {
                    $successfulSignIns++
                } else {
                    $failedSignIns++
                }
            }
            
            $message = "Found $($newSignIns.Count) new sign-in event"
            if ($newSignIns.Count -gt 1) {
                $message += "s"
            }
            $message += ":"
            
            if ($successfulSignIns -gt 0) {
                $message += " $successfulSignIns successful sign-in"
                if ($successfulSignIns -gt 1) {
                    $message += "s"
                }
                
                if ($failedSignIns -gt 0) {
                    $message += " and"
                }
            }
            
            if ($failedSignIns -gt 0) {
                $message += " $failedSignIns failed sign-in attempt"
                if ($failedSignIns -gt 1) {
                    $message += "s"
                }
            }
            
            Write-Log $message -Success
        }
        else {
            Write-Log "All sign-in events have already been processed. No new alerts needed."
            
            # Update last run time but keep processed IDs
            $newState = @{
                LastRunTime = [DateTime]::UtcNow.ToString('o')
                ProcessedIds = $state.ProcessedIds
            }
            
            Save-ProcessedState -StateFilePath $Config.StateFilePath -State $newState
            Write-Log "Monitoring complete."
            return
        }
        
        # Track new processed IDs
        $newProcessedIds = @()
        $emailsSent = 0
        
        # Process sign-ins in chronological order (oldest first) - FOR SINGLE EVENTS, CRUCIAL CHANGE
        # First update the state, then send the email
        foreach ($signIn in ($newSignIns | Sort-Object createdDateTime)) {
            $newProcessedIds += $signIn.id
            
            # CRITICAL: Update state FIRST, then send email
            $curState = Get-LastProcessedState -StateFilePath $Config.StateFilePath
            $curIds = $curState.ProcessedIds + @($signIn.id) | Select-Object -Unique
            $immediate = @{
                LastRunTime = [DateTime]::UtcNow.ToString('o')
                ProcessedIds = $curIds
            }
            Save-ProcessedState -StateFilePath $Config.StateFilePath -State $immediate
            Write-Log "Updated state with sign-in ID: $($signIn.id) before sending email"
            
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
            
            # Determine if this was a successful sign-in or a failed attempt
            $isSuccessful = $signIn.status -and $signIn.status.errorCode -eq 0
            
            # Send the notification with appropriate subject based on sign-in status
            if ($isSuccessful) {
                $emailSubject = "Security Alert - Successful Sign-In Detected for $userDisplayName"
                $logMessage = "Successful sign-in"
            } else {
                $emailSubject = "Security Alert - Failed Sign-In Attempt Detected for $userDisplayName"
                $logMessage = "Failed sign-in attempt"
            }
            
            # PROCESS ONE AT A TIME
            $emailSent = Send-EmailNotification -AccessToken $accessToken -SenderEmail $Config.SenderEmail -RecipientEmail $Config.RecipientEmail -Subject $emailSubject -Body $emailBody
            
            if ($emailSent) {
                Write-Log "Alert sent for $logMessage at $timeFormatted from $($signIn.ipAddress) (ID: $($signIn.id))" -Success
                $emailsSent++
                
                # Wait a bit after sending each email to avoid any rate limits
                Start-Sleep -Seconds 2
            } else {
                Write-Log "Failed to send alert for $logMessage at $timeFormatted from $($signIn.ipAddress)" -Error
            }
        }
        
        # Final state update
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
        
        if ($emailsSent -gt 0) {
            Write-Log "Successfully sent $emailsSent email alerts." -Success
        } else {
            Write-Log "No email alerts were sent successfully." -Error
        }
    } else {
        Write-Log "No sign-ins or sign-in attempts detected for the monitored account in the last $($Config.TimeWindowMinutes) minutes." -Error
        
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
        Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -Confirm:$false
    }
    Import-Module $module -ErrorAction SilentlyContinue
}
Write-Log "Required modules verified."

# Run the monitor function with the configuration
Monitor-SignIns -Config $config