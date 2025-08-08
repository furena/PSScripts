<#
.SYNOPSIS
    Microsoft 365 Groups Domain Cleanup Script
    Sets onmicrosoft.com addresses as primary and removes old domain aliases.

.DESCRIPTION
    This script specifically handles Microsoft 365 Groups during tenant-to-tenant migrations.
    It will:
    1. Set the @tenant.onmicrosoft.com address as the primary SMTP address
    2. Remove all @olddomain.com aliases from the groups
    3. Log all changes for audit purposes

.PARAMETER OldDomain
    The old domain(s) to be removed from groups (e.g., 'contoso.com')

.PARAMETER NewDomain
    The new domain to use as primary (defaults to tenant's onmicrosoft.com domain)

.PARAMETER Identity
    Test against a single group. Useful for testing before running full cleanup.

.PARAMETER LogPath
    Custom path for log files. If not specified, uses current directory.

.EXAMPLE
    .\M365GroupsDomainCleanup.ps1 -OldDomain "contoso.com"
    # Clean up all groups, remove contoso.com aliases

.EXAMPLE
    .\M365GroupsDomainCleanup.ps1 -OldDomain "contoso.com" -Identity "sales@contoso.com"
    # Test cleanup on single group

.EXAMPLE
    .\M365GroupsDomainCleanup.ps1 -OldDomain "contoso.com" -NewDomain "newcompany.onmicrosoft.com"
    # Use specific onmicrosoft.com domain as primary

.NOTES
    Author: furena
    Date: 2025-08-08
    Version: 1.0
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory=$true, HelpMessage="Old domain(s) to be removed (e.g., 'contoso.com')")]
    [string[]]$OldDomain,
    
    [Parameter(Mandatory=$false, HelpMessage="New domain to use as primary (defaults to tenant's onmicrosoft.com domain)")]
    [string]$NewDomain,
    
    [Parameter(Mandatory=$false, HelpMessage="Test against a single group (group email address)")]
    [string]$Identity,
    
    [Parameter(Mandatory=$false, HelpMessage="Path for log files (defaults to current directory)")]
    [string]$LogPath = (Get-Location).Path
)

#region SETUP AND VALIDATION
Write-Host "=== MICROSOFT 365 GROUPS DOMAIN CLEANUP ===" -ForegroundColor Magenta
Write-Host "Current User: $env:USERNAME" -ForegroundColor Yellow
Write-Host "Current Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')" -ForegroundColor Yellow

if ($Identity) {
    Write-Host "SINGLE GROUP TEST MODE" -ForegroundColor Yellow
    Write-Host "Target Group: $Identity" -ForegroundColor Cyan
}

# Validate domains
foreach ($domain in $OldDomain) {
    if ($domain -notmatch '^[a-zA-Z0-9][a-zA-Z0-9-]{0,61}[a-zA-Z0-9]?\.[a-zA-Z]{2,}$') {
        Write-Error "Invalid domain format: $domain"
        exit 1
    }
}

Write-Host "`nDomains to clean up:" -ForegroundColor Cyan
$OldDomain | ForEach-Object { Write-Host "  - $_" -ForegroundColor White }

# Setup logging
$TimeStamp = Get-Date -UFormat '%Y-%m-%d_%H-%M-%S'
$ModePrefix = if ($Identity) { "SingleGroup" } else { "AllGroups" }
$ChangeLog = Join-Path $LogPath "$ModePrefix`_GroupCleanup_ChangeLog_$TimeStamp.csv"
$ErrorLog = Join-Path $LogPath "$ModePrefix`_GroupCleanup_Errors_$TimeStamp.log"

# Initialize log files
@"
Timestamp,GroupName,GroupEmail,Action,OldValue,NewValue,Status
"@ | Out-File -FilePath $ChangeLog -Encoding UTF8

Write-Host "`nLog files:" -ForegroundColor Green
Write-Host "  Changes: $ChangeLog" -ForegroundColor White
Write-Host "  Errors: $ErrorLog" -ForegroundColor White
#endregion

#region CONNECT TO SERVICES
Write-Host "`nConnecting to Exchange Online and Microsoft Graph..." -ForegroundColor Yellow

try {
    # Increase function capacity to avoid Graph module issues
    $MaximumFunctionCount = 8192
    
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop
    
    if (-not $WhatIfPreference) {
        # Connect to Exchange Online first
        Connect-ExchangeOnline -ErrorAction Stop
        Write-Host "✓ Connected to Exchange Online" -ForegroundColor Green
        
        # Get tenant info and connect to Graph
        try {
            $OrgConfig = Get-OrganizationConfig -ErrorAction Stop
            $TenantDomain = $OrgConfig.Name
            Write-Host "Detected tenant: $TenantDomain" -ForegroundColor Cyan
            
            Connect-MgGraph -TenantId $TenantDomain -Scopes 'Group.ReadWrite.All','Directory.Read.All' -ErrorAction Stop
            Write-Host "✓ Connected to Microsoft Graph" -ForegroundColor Green
        } catch {
            Write-Warning "Could not auto-detect tenant. Manual connection required."
            $ManualTenant = Read-Host "Enter the tenant ID or domain for Microsoft Graph connection"
            Connect-MgGraph -TenantId $ManualTenant -Scopes 'Group.ReadWrite.All','Directory.Read.All' -ErrorAction Stop
            Write-Host "✓ Connected to Microsoft Graph" -ForegroundColor Green
        }
    } else {
        Write-Host "WhatIf mode: Skipping actual connections" -ForegroundColor Yellow
    }
} catch {
    $ErrorMessage = "Failed to connect to required services: $($_.Exception.Message)"
    Write-Error $ErrorMessage
    $ErrorMessage | Out-File -FilePath $ErrorLog -Append
    exit 1
}
#endregion

#region DETERMINE NEW DOMAIN
if (-not $NewDomain) {
    if (-not $WhatIfPreference) {
        try {
            $NewDomain = (Get-MgDomain | Where-Object {$_.IsInitial -eq $true}).Id
            if (-not $NewDomain) {
                throw "Could not find tenant's onmicrosoft.com domain"
            }
            Write-Host "Auto-detected new domain: $NewDomain" -ForegroundColor Cyan
        } catch {
            Write-Error "Could not auto-detect new domain. Please specify -NewDomain parameter."
            exit 1
        }
    } else {
        $NewDomain = "example.onmicrosoft.com"  # Placeholder for WhatIf mode
    }
}
#endregion

#region HELPER FUNCTIONS
function Write-LogEntry {
    param(
        [string]$GroupName,
        [string]$GroupEmail,
        [string]$Action,
        [string]$OldValue,
        [string]$NewValue,
        [string]$Status
    )
    
    $LogEntry = @{
        Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC'
        GroupName = $GroupName
        GroupEmail = $GroupEmail
        Action = $Action
        OldValue = $OldValue
        NewValue = $NewValue
        Status = $Status
    }
    
    "$($LogEntry.Timestamp),$($LogEntry.GroupName),$($LogEntry.GroupEmail),$($LogEntry.Action),$($LogEntry.OldValue),$($LogEntry.NewValue),$($LogEntry.Status)" | 
        Out-File -FilePath $ChangeLog -Append -Encoding UTF8
}

function Update-GroupAddresses {
    param(
        [object]$Group
    )
    
    $GroupName = $Group.DisplayName
    $GroupEmail = $Group.WindowsEmailAddress
    $CurrentAddresses = $Group.EmailAddresses
    
    Write-Host "`nProcessing group: $GroupName ($GroupEmail)" -ForegroundColor Cyan
    Write-Host "Current addresses: $($CurrentAddresses -join ', ')" -ForegroundColor Gray
    
    # Find addresses to remove (old domain)
    $AddressesToRemove = @()
    $CurrentPrimary = $null
    $NewPrimary = $null
    
    foreach ($addr in $CurrentAddresses) {
        if ($addr -like "SMTP:*") {
            $CurrentPrimary = $addr
        }
        
        foreach ($oldDom in $OldDomain) {
            if ($addr -like "*@$oldDom") {
                $AddressesToRemove += $addr
                Write-Host "  Will remove: $addr" -ForegroundColor Red
            }
        }
        
        # Find potential new primary (onmicrosoft.com address)
        if ($addr -like "*@$NewDomain" -and $addr -like "smtp:*") {
            $NewPrimary = $addr.Replace("smtp:", "SMTP:")
        }
    }
    
    if ($AddressesToRemove.Count -eq 0) {
        Write-Host "  No addresses to remove for this group" -ForegroundColor Green
        Write-LogEntry -GroupName $GroupName -GroupEmail $GroupEmail -Action "No Changes" -OldValue "N/A" -NewValue "N/A" -Status "Skipped"
        return
    }
    
    # Build new address list
    $NewAddresses = @()
    
    foreach ($addr in $CurrentAddresses) {
        $ShouldRemove = $false
        foreach ($removeAddr in $AddressesToRemove) {
            if ($addr -eq $removeAddr) {
                $ShouldRemove = $true
                break
            }
        }
        
        if (-not $ShouldRemove) {
            # If this was the primary and we're removing it, make it secondary
            if ($addr -like "SMTP:*") {
                foreach ($oldDom in $OldDomain) {
                    if ($addr -like "*@$oldDom") {
                        $NewAddresses += $addr.Replace("SMTP:", "smtp:")
                        $ShouldRemove = $true
                        break
                    }
                }
            }
            
            if (-not $ShouldRemove) {
                $NewAddresses += $addr
            }
        }
    }
    
    # Set new primary if we found one
    if ($NewPrimary) {
        # Remove the old primary designation from onmicrosoft.com address
        $NewAddresses = $NewAddresses | Where-Object { $_ -ne $NewPrimary.Replace("SMTP:", "smtp:") }
        # Add it as primary
        $NewAddresses += $NewPrimary
        Write-Host "  New primary: $NewPrimary" -ForegroundColor Green
    }
    
    # Remove old domain addresses completely
    $FinalAddresses = @()
    foreach ($addr in $NewAddresses) {
        $ShouldAdd = $true
        foreach ($oldDom in $OldDomain) {
            if ($addr -like "*@$oldDom") {
                $ShouldAdd = $false
                break
            }
        }
        if ($ShouldAdd) {
            $FinalAddresses += $addr
        }
    }
    
    Write-Host "  Final addresses: $($FinalAddresses -join ', ')" -ForegroundColor Yellow
    
    # Apply changes
    if ($PSCmdlet.ShouldProcess($GroupName, "Update email addresses")) {
        try {
            Set-UnifiedGroup -Identity $Group.PrimarySmtpAddress -EmailAddresses $FinalAddresses -ErrorAction Stop
            Write-Host "  ✓ Successfully updated group addresses" -ForegroundColor Green
            
            Write-LogEntry -GroupName $GroupName -GroupEmail $GroupEmail -Action "Updated Addresses" `
                -OldValue ($CurrentAddresses -join ';') -NewValue ($FinalAddresses -join ';') -Status "Success"
            
            return $true
        } catch {
            $ErrorMessage = "Failed to update group $GroupName`: $($_.Exception.Message)"
            Write-Host "  ✗ $ErrorMessage" -ForegroundColor Red
            $ErrorMessage | Out-File -FilePath $ErrorLog -Append
            
            Write-LogEntry -GroupName $GroupName -GroupEmail $GroupEmail -Action "Update Failed" `
                -OldValue ($CurrentAddresses -join ';') -NewValue "FAILED" -Status "Error"
            
            return $false
        }
    } else {
        Write-Host "  [WhatIf] Would update addresses" -ForegroundColor Yellow
        Write-LogEntry -GroupName $GroupName -GroupEmail $GroupEmail -Action "WhatIf - Would Update" `
            -OldValue ($CurrentAddresses -join ';') -NewValue ($FinalAddresses -join ';') -Status "WhatIf"
        return $true
    }
}
#endregion

#region MAIN PROCESSING
Write-Host "`n=== PROCESSING MICROSOFT 365 GROUPS ===" -ForegroundColor Magenta

$ProcessedCount = 0
$SuccessCount = 0
$ErrorCount = 0

try {
    if ($Identity) {
        # Single group mode
        Write-Host "Getting single group: $Identity" -ForegroundColor Yellow
        $Groups = @(Get-UnifiedGroup -Identity $Identity -ErrorAction Stop)
    } else {
        # All groups mode
        Write-Host "Getting all Microsoft 365 groups..." -ForegroundColor Yellow
        $Groups = Get-UnifiedGroup -ResultSize Unlimited
    }
    
    Write-Host "Found $($Groups.Count) group(s) to process" -ForegroundColor Cyan
    
    foreach ($Group in $Groups) {
        $ProcessedCount++
        Write-Progress -Activity "Processing Microsoft 365 Groups" -Status "Group $ProcessedCount of $($Groups.Count)" -PercentComplete (($ProcessedCount / $Groups.Count) * 100)
        
        $Success = Update-GroupAddresses -Group $Group
        
        if ($Success) {
            $SuccessCount++
        } else {
            $ErrorCount++
        }
        
        Start-Sleep -Milliseconds 500  # Brief pause to avoid throttling
    }
    
} catch {
    $ErrorMessage = "Failed to retrieve groups: $($_.Exception.Message)"
    Write-Error $ErrorMessage
    $ErrorMessage | Out-File -FilePath $ErrorLog -Append
    exit 1
}
#endregion

#region SUMMARY
Write-Host "`n=== SUMMARY ===" -ForegroundColor Magenta
Write-Host "Groups processed: $ProcessedCount" -ForegroundColor White
Write-Host "Successful updates: $SuccessCount" -ForegroundColor Green
Write-Host "Failed updates: $ErrorCount" -ForegroundColor Red

if ($ErrorCount -gt 0) {
    Write-Host "`nErrors logged to: $ErrorLog" -ForegroundColor Yellow
}

Write-Host "`nChange log: $ChangeLog" -ForegroundColor Green
Write-Host "`nMicrosoft 365 Groups domain cleanup completed!" -ForegroundColor Green

# Disconnect sessions
if (-not $WhatIfPreference) {
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "Disconnected from services" -ForegroundColor Yellow
    } catch {
        # Ignore disconnect errors
    }
}
#endregion