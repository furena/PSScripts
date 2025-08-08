<#
.SYNOPSIS
    Adds secondary alias addresses from source to migrated mailboxes in target Exchange Online tenant.

.DESCRIPTION
    This script processes a CSV file containing user mappings and adds the original  email addresses 
    as secondary proxy addresses to the corresponding mailboxes in the target tenant. This ensures users can 
    continue receiving mail at their old email addresses after the tenant migration.

.PARAMETER CsvPath
    Path to the CSV file containing user mappings. Required columns:
    - UserPrincipalName: Current UPN in Contoso tenant (e.g., john.doe@contoso.com)
    - SecondaryAlias: Original Fabrikam email address (e.g., john.doe@fabrikam.com)

.PARAMETER LogPath
    Optional path for log file. Defaults to script directory with timestamp.

.PARAMETER WhatIf
    Shows what would be done without making actual changes.

.EXAMPLE
    .\Update-emailAliases.ps1 -CsvPath "C:\Migration\emailliases.csv"

.EXAMPLE
    .\Update-emailAliases.ps1 -CsvPath "C:\Migration\emailAliases.csv" -WhatIf

.NOTES
    - Requires Exchange Online PowerShell V2 module
    - Must be connected to Exchange Online before running
    - Run Connect-ExchangeOnline before executing this script
    - Author: Migration Script for mail migrations following acquisition
    - Date: 2025-08-08
#>

param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_})]
    [string]$CsvPath,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = ".\emailAliasUpdate_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

# Initialize logging
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage
    Add-Content -Path $LogPath -Value $logMessage
}

# Check if Exchange Online module is available
try {
    if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "Exchange Online Management module not found. Install with: Install-Module ExchangeOnlineManagement"
    }
    
    # Check if connected to Exchange Online
    $session = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if (!$session) {
        throw "Not connected to Exchange Online. Run Connect-ExchangeOnline first."
    }
    
    Write-Log "Connected to Exchange Online tenant: $($session.TenantId)"
}
catch {
    Write-Log "Prerequisites check failed: $_" -Level "ERROR"
    exit 1
}

# Import and validate CSV
try {
    Write-Log "Importing CSV file: $CsvPath"
    $users = Import-Csv -Path $CsvPath
    
    if ($users.Count -eq 0) {
        throw "CSV file is empty or invalid"
    }
    
    # Validate required columns
    $requiredColumns = @('UserPrincipalName', 'SecondaryAlias')
    foreach ($column in $requiredColumns) {
        if ($column -notin $users[0].PSObject.Properties.Name) {
            throw "Required column '$column' not found in CSV"
        }
    }
    
    Write-Log "Successfully imported $($users.Count) user records"
}
catch {
    Write-Log "Failed to import CSV: $_" -Level "ERROR"
    exit 1
}

# Process each user
$successCount = 0
$errorCount = 0
$skippedCount = 0

Write-Log "Starting processing of $($users.Count) users..."

foreach ($user in $users) {
    $upn = $user.UserPrincipalName?.Trim()
    $SecondaryAlias = $user.SecondaryAlias?.Trim()
    
    # Validate row data
    if ([string]::IsNullOrWhiteSpace($upn) -or [string]::IsNullOrWhiteSpace($SecondaryAlias)) {
        Write-Log "Skipping row with missing data - UPN: '$upn', SecondaryAlias: '$SecondaryAlias'" -Level "WARNING"
        $skippedCount++
        continue
    }
    
    try {
        # Get mailbox
        Write-Log "Processing user: $upn"
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop
        
        # Get current proxy addresses
        $currentProxies = $mailbox.EmailAddresses
        
        # Format the alias as secondary SMTP address (lowercase 'smtp:')
        $newProxyAddress = "smtp:$SecondaryAlias"
        
        # Check if alias already exists
        $existingProxy = $currentProxies | Where-Object { $_.ToString().ToLower() -eq $newProxyAddress.ToLower() }
        
        if ($existingProxy) {
            Write-Log "Secondary alias '$SecondaryAlias' already exists for $upn - skipping" -Level "INFO"
            $skippedCount++
            continue
        }
        
        # Add the new proxy address
        if ($WhatIf) {
            Write-Log "[WHATIF] Would add '$SecondaryAlias' as secondary alias to $upn" -Level "INFO"
        }
        else {
            $updatedProxies = $currentProxies + $newProxyAddress
            Set-Mailbox -Identity $upn -EmailAddresses $updatedProxies -ErrorAction Stop
            Write-Log "Successfully added '$SecondaryAlias' as secondary alias to $upn" -Level "INFO"
        }
        
        $successCount++
    }
    catch {
        Write-Log "Failed to process $upn : $_" -Level "ERROR"
        $errorCount++
    }
}

# Summary
Write-Log "=== PROCESSING COMPLETE ===" -Level "INFO"
Write-Log "Total users processed: $($users.Count)" -Level "INFO"
Write-Log "Successful updates: $successCount" -Level "INFO"
Write-Log "Errors: $errorCount" -Level "INFO"
Write-Log "Skipped: $skippedCount" -Level "INFO"

if ($WhatIf) {
    Write-Log "WhatIf mode was enabled - no actual changes were made" -Level "INFO"
}

Write-Log "Log file saved to: $LogPath" -Level "INFO"