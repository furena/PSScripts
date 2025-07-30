<#
 .SYNOPSIS
   Universal domain cleanup script for tenant-to-tenant migrations.
   Removes ALL references to specified old domain(s) to prepare for domain transfer.

 .DESCRIPTION
   This script provides a complete solution for cleaning up domain references during
   tenant-to-tenant migrations. It can handle single or multiple domains and ensures
   complete removal of all references including UPNs, SMTP addresses, SIP addresses,
   and proxy addresses.

 .PARAMETER OldDomain
   The old domain(s) to be cleaned up. Can be a single domain or array of domains.

 .PARAMETER NewDomain
   The new domain to use. If not specified, uses the tenant's onmicrosoft.com domain.

 .PARAMETER Identity
   Test against a single user/mailbox. Useful for testing before running full migration.
   Can be UPN, email address, or display name.

 .PARAMETER LogPath
   Custom path for log files. If not specified, uses current directory.

 .EXAMPLE
   .\UniversalDomainCleanup.ps1
   # Interactive mode - prompts for domain

 .EXAMPLE
   .\UniversalDomainCleanup.ps1 -OldDomain "contoso.com"
   # Clean up single domain

 .EXAMPLE
   .\UniversalDomainCleanup.ps1 -OldDomain "contoso.com" -Identity "allison@bigcatlabs.com"
   # Test cleanup on single user

 .EXAMPLE
   .\UniversalDomainCleanup.ps1 -OldDomain @("contoso.com","fabrikam.com") -NewDomain "newcompany.onmicrosoft.com"
   # Clean up multiple domains to specific new domain

 .EXAMPLE
   .\UniversalDomainCleanup.ps1 -OldDomain "contoso.com" -Identity "test.user@contoso.com" -WhatIf
   # Preview changes for single user without making them

 .DISCLAIMER
   • This logs every change and validates thoroughly
   • Test in a lab environment first!
   • Backup your tenant before running!
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory=$false, HelpMessage="Domain(s) to be cleaned up (e.g., 'contoso.com' or @('contoso.com','fabrikam.com'))")]
    [string[]]$OldDomain,
    
    [Parameter(Mandatory=$false, HelpMessage="New domain to use (defaults to tenant's onmicrosoft.com domain)")]
    [string]$NewDomain,
    
    [Parameter(Mandatory=$false, HelpMessage="Test against a single user/mailbox (UPN, email, or display name)")]
    [string]$Identity,
    
    [Parameter(Mandatory=$false, HelpMessage="Path for log files (defaults to current directory)")]
    [string]$LogPath = (Get-Location).Path
)

#region 0. PARAMETER VALIDATION AND SETUP
Write-Host "=== UNIVERSAL DOMAIN CLEANUP SCRIPT ===" -ForegroundColor Magenta
Write-Host "For Tenant-to-Tenant Migrations" -ForegroundColor Cyan
Write-Host "Current User: $env:USERNAME" -ForegroundColor Yellow
Write-Host "Current Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')" -ForegroundColor Yellow

if ($Identity) {
    Write-Host "SINGLE USER TEST MODE" -ForegroundColor Yellow
    Write-Host "Target Identity: $Identity" -ForegroundColor Cyan
}

# Interactive domain input if not provided
if (-not $OldDomain) {
    Write-Host "`nNo domain specified. Let's set this up interactively." -ForegroundColor Yellow
    
    do {
        $DomainInput = Read-Host "`nEnter the old domain to clean up (e.g., contoso.com)"
        if ([string]::IsNullOrWhiteSpace($DomainInput)) {
            Write-Host "Domain cannot be empty. Please try again." -ForegroundColor Red
        }
    } while ([string]::IsNullOrWhiteSpace($DomainInput))
    
    $OldDomain = @($DomainInput.Trim())
    
    # Ask if there are additional domains (only if not testing single identity)
    if (-not $Identity) {
        $AddMore = Read-Host "`nDo you have additional domains to clean up? (y/n)"
        while ($AddMore -eq 'y' -or $AddMore -eq 'Y') {
            $AdditionalDomain = Read-Host "Enter additional domain"
            if (-not [string]::IsNullOrWhiteSpace($AdditionalDomain)) {
                $OldDomain += $AdditionalDomain.Trim()
            }
            $AddMore = Read-Host "Add another domain? (y/n)"
        }
    }
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

# Confirm before proceeding (skip confirmation for single user tests with WhatIf)
if (-not $WhatIfPreference -and -not ($Identity -and $WhatIfPreference)) {
    if ($Identity) {
        $Confirmation = Read-Host "`nThis will update the specified user ($Identity). Continue? (y/n)"
    } else {
        $Confirmation = Read-Host "`nThis will remove ALL references to the above domain(s) from ALL users. Continue? (y/n)"
    }
    
    if ($Confirmation -ne 'y' -and $Confirmation -ne 'Y') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit 0
    }
}
#endregion

#region 1. CONNECT TO SERVICES
Write-Host "`nConnecting to Exchange Online and Microsoft Graph..." -ForegroundColor Yellow

try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Import-Module Microsoft.Graph -ErrorAction Stop
    
    if (-not $WhatIfPreference) {
        Connect-ExchangeOnline -ErrorAction Stop
        Connect-MgGraph -Scopes 'User.ReadWrite.All','Group.ReadWrite.All','Directory.ReadWrite.All','Application.Read.All' -ErrorAction Stop
    } else {
        Write-Host "WhatIf mode: Skipping actual connections" -ForegroundColor Yellow
    }
} catch {
    Write-Error "Failed to connect to required services: $($_.Exception.Message)"
    exit 1
}
#endregion

#region 2. VARIABLES AND SETUP
# Auto-detect new domain if not specified
if (-not $NewDomain) {
    if (-not $WhatIfPreference) {
        try {
            $NewDomain = (Get-MgDomain | Where-Object {$_.IsInitial -eq $true}).Id
            if (-not $NewDomain) {
                throw "Could not find tenant's onmicrosoft.com domain"
            }
        } catch {
            Write-Error "Could not auto-detect new domain. Please specify -NewDomain parameter."
            exit 1
        }
    } else {
        $NewDomain = "example.onmicrosoft.com"  # Placeholder for WhatIf mode
    }
}

$TimeStamp = Get-Date -UFormat '%Y-%m-%d_%H-%M-%S'
$ModePrefix = if ($Identity) { "SingleUser" } else { "FullTenant" }
$ChangeLog = Join-Path $LogPath "$ModePrefix`_DomainCleanup_ChangeLog_$TimeStamp.csv"
$ErrorLog  = Join-Path $LogPath "$ModePrefix`_DomainCleanup_Errors_$TimeStamp.log"
$ValidationLog = Join-Path $LogPath "$ModePrefix`_DomainCleanup_Validation_$TimeStamp.log"

Write-Host "`nConfiguration:" -ForegroundColor Green
Write-Host "  Mode: $(if ($Identity) { "Single User Test" } else { "Full Tenant" })" -ForegroundColor White
Write-Host "  Old Domain(s): $($OldDomain -join ', ')" -ForegroundColor White
Write-Host "  New Domain: $NewDomain" -ForegroundColor White
if ($Identity) { Write-Host "  Target Identity: $Identity" -ForegroundColor White }
Write-Host "  Log Path: $LogPath" -ForegroundColor White
Write-Host "  WhatIf Mode: $WhatIfPreference" -ForegroundColor White

if (-not $WhatIfPreference) {
    # Initialize log files
    $LogHeader = if ($Identity) {
        "Starting SINGLE USER domain cleanup for [$Identity] from [$($OldDomain -join ', ')] to [$NewDomain] at $(Get-Date)"
    } else {
        "Starting FULL TENANT domain cleanup from [$($OldDomain -join ', ')] to [$NewDomain] at $(Get-Date)"
    }
    $LogHeader | Out-File $ErrorLog
    $LogHeader | Out-File $ValidationLog
}
#endregion

#region 3. HELPER FUNCTIONS
function Replace-Domain {
    param(
        [string]$Address,
        [string[]]$OldDomains,
        [string]$NewDomainTarget
    )
    
    if ([string]::IsNullOrEmpty($Address)) { return $Address }
    
    $UpdatedAddress = $Address
    foreach ($oldDom in $OldDomains) {
        $UpdatedAddress = $UpdatedAddress -replace "@$oldDom", "@$NewDomainTarget"
    }
    return $UpdatedAddress
}

function Test-ContainsOldDomain {
    param(
        [string]$Address,
        [string[]]$OldDomains
    )
    
    foreach ($oldDom in $OldDomains) {
        if ($Address -match "@$oldDom") {
            return $true
        }
    }
    return $false
}

function Remove-OldDomainProxies {
    param(
        [array]$ProxyAddresses,
        [string[]]$OldDomains
    )
    
    $CleanProxies = @()
    foreach ($proxy in $ProxyAddresses) {
        $ContainsOld = $false
        foreach ($oldDom in $OldDomains) {
            if ($proxy -match "@$oldDom") {
                $ContainsOld = $true
                break
            }
        }
        
        if (-not $ContainsOld) {
            $CleanProxies += $proxy
        }
    }
    
    return $CleanProxies
}

function Write-LogEntry {
    param(
        [string]$Message,
        [string]$LogFile = $ValidationLog,
        [switch]$Error
    )
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$TimeStamp - $Message"
    
    if ($Error) {
        Write-Host $Message -ForegroundColor Red
        if (-not $WhatIfPreference) { $LogEntry | Out-File $ErrorLog -Append }
    } else {
        Write-Host $Message -ForegroundColor Yellow
    }
    
    if (-not $WhatIfPreference) { $LogEntry | Out-File $LogFile -Append }
}

function Get-ObjectsWithOldDomain {
    param(
        [string]$ObjectType,
        [string[]]$OldDomains,
        [string]$SingleIdentity = $null
    )
    
    $Objects = @()
    
    try {
        switch ($ObjectType) {
            'Mailbox' {
                if (-not $WhatIfPreference) {
                    if ($SingleIdentity) {
                        # Single user mode - get specific mailbox
                        try {
                            $SingleMailbox = Get-Mailbox -Identity $SingleIdentity -ErrorAction Stop
                            
                            # Check if this mailbox has old domain references
                            $HasOldDomain = $false
                            
                            # Check UPN
                            foreach ($oldDom in $OldDomains) {
                                if ($SingleMailbox.UserPrincipalName -match "@$oldDom") {
                                    $HasOldDomain = $true
                                    break
                                }
                            }
                            
                            # Check email addresses
                            if (-not $HasOldDomain) {
                                foreach ($email in $SingleMailbox.EmailAddresses) {
                                    foreach ($oldDom in $OldDomains) {
                                        if ($email -match "@$oldDom") {
                                            $HasOldDomain = $true
                                            break
                                        }
                                    }
                                    if ($HasOldDomain) { break }
                                }
                            }
                            
                            if ($HasOldDomain) {
                                $Objects += $SingleMailbox
                                Write-LogEntry "Found mailbox $($SingleMailbox.DisplayName) with old domain references"
                            } else {
                                Write-LogEntry "Mailbox $($SingleMailbox.DisplayName) has no old domain references - skipping"
                            }
                        } catch {
                            Write-LogEntry "ERROR: Could not find mailbox with identity '$SingleIdentity': $($_.Exception.Message)" -Error
                        }
                    } else {
                        # Full tenant mode - get all mailboxes
                        $AllMailboxes = Get-Mailbox -ResultSize Unlimited
                        foreach ($mbx in $AllMailboxes) {
                            $HasOldDomain = $false
                            
                            # Check UPN
                            foreach ($oldDom in $OldDomains) {
                                if ($mbx.UserPrincipalName -match "@$oldDom") {
                                    $HasOldDomain = $true
                                    break
                                }
                            }
                            
                            # Check email addresses
                            if (-not $HasOldDomain) {
                                foreach ($email in $mbx.EmailAddresses) {
                                    foreach ($oldDom in $OldDomains) {
                                        if ($email -match "@$oldDom") {
                                            $HasOldDomain = $true
                                            break
                                        }
                                    }
                                    if ($HasOldDomain) { break }
                                }
                            }
                            
                            if ($HasOldDomain) {
                                $Objects += $mbx
                            }
                        }
                    }
                } else {
                    # WhatIf mode - return sample data
                    if ($SingleIdentity) {
                        $Objects += [PSCustomObject]@{
                            DisplayName = "Test User ($SingleIdentity)"
                            UserPrincipalName = $SingleIdentity
                            PrimarySmtpAddress = $SingleIdentity
                            EmailAddresses = @("SMTP:$SingleIdentity", "smtp:user.old@$($OldDomains[0])")
                            RecipientTypeDetails = "UserMailbox"
                        }
                    } else {
                        $Objects += [PSCustomObject]@{
                            DisplayName = "Sample User"
                            UserPrincipalName = "user@$($OldDomains[0])"
                            PrimarySmtpAddress = "user@$($OldDomains[0])"
                            EmailAddresses = @("SMTP:user@$($OldDomains[0])", "smtp:user.old@$($OldDomains[0])")
                            RecipientTypeDetails = "UserMailbox"
                        }
                    }
                }
            }
            
            'DistributionGroup' {
                if (-not $WhatIfPreference -and -not $SingleIdentity) {  # Skip groups in single user mode
                    $AllDGs = Get-DistributionGroup -ResultSize Unlimited
                    foreach ($dg in $AllDGs) {
                        foreach ($email in $dg.EmailAddresses) {
                            foreach ($oldDom in $OldDomains) {
                                if ($email -match "@$oldDom") {
                                    $Objects += $dg
                                    break
                                }
                            }
                            if ($Objects -contains $dg) { break }
                        }
                    }
                }
            }
            
            'UnifiedGroup' {
                if (-not $WhatIfPreference -and -not $SingleIdentity) {  # Skip groups in single user mode
                    $AllM365Groups = Get-UnifiedGroup -ResultSize Unlimited
                    foreach ($group in $AllM365Groups) {
                        foreach ($email in $group.EmailAddresses) {
                            foreach ($oldDom in $OldDomains) {
                                if ($email -match "@$oldDom") {
                                    $Objects += $group
                                    break
                                }
                            }
                            if ($Objects -contains $group) { break }
                        }
                    }
                }
            }
        }
    } catch {
        Write-LogEntry "Error getting $ObjectType objects: $($_.Exception.Message)" -Error
    }
    
    return $Objects
}
#endregion

#region 4. COMPREHENSIVE USER/MAILBOX CLEANUP
Write-Host "`n=== PHASE 1: USER MAILBOXES ===" -ForegroundColor Magenta

$AllMailboxes = Get-ObjectsWithOldDomain -ObjectType 'Mailbox' -OldDomains $OldDomain -SingleIdentity $Identity

if ($Identity -and $AllMailboxes.Count -eq 0) {
    Write-Host "No old domain references found for identity '$Identity'" -ForegroundColor Green
    Write-Host "This user either doesn't exist or already has clean domain references." -ForegroundColor Yellow
    
    # Still try to show the mailbox details for verification
    if (-not $WhatIfPreference) {
        try {
            $TestMailbox = Get-Mailbox -Identity $Identity -ErrorAction Stop
            Write-Host "`nCurrent mailbox details for '$Identity':" -ForegroundColor Cyan
            Write-Host "  Display Name: $($TestMailbox.DisplayName)" -ForegroundColor White
            Write-Host "  UPN: $($TestMailbox.UserPrincipalName)" -ForegroundColor White
            Write-Host "  Primary SMTP: $($TestMailbox.PrimarySmtpAddress)" -ForegroundColor White
            Write-Host "  Email Addresses:" -ForegroundColor White
            $TestMailbox.EmailAddresses | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
        } catch {
            Write-Host "Could not retrieve mailbox details: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    # Exit early if no work to do
    if (-not $WhatIfPreference) {
        Disconnect-ExchangeOnline -Confirm:$false
        Disconnect-MgGraph
    }
    exit 0
}

Write-LogEntry "Found $($AllMailboxes.Count) mailbox$(if($AllMailboxes.Count -ne 1){'es'}) with old domain references"

$UserChanges = @()
foreach ($mbx in $AllMailboxes) {
    try {
        Write-Host "Processing mailbox: $($mbx.DisplayName) ($($mbx.UserPrincipalName)) - Type: $($mbx.RecipientTypeDetails)" -ForegroundColor Cyan
        
        $Changes = @{
            ObjectType = $mbx.RecipientTypeDetails
            Identity = $mbx.UserPrincipalName
            DisplayName = $mbx.DisplayName
            OldUPN = $mbx.UserPrincipalName
            OldPrimarySMTP = $mbx.PrimarySmtpAddress
            OldProxyCount = ($mbx.EmailAddresses | Where-Object { Test-ContainsOldDomain -Address $_ -OldDomains $OldDomain }).Count
            NewUPN = $null
            NewPrimarySMTP = $null
            NewProxyCount = 0
            Status = "Processing"
        }

        # Show current state for single user tests
        if ($Identity) {
            Write-Host "`nCurrent state:" -ForegroundColor Yellow
            Write-Host "  UPN: $($mbx.UserPrincipalName)" -ForegroundColor White
            Write-Host "  Primary SMTP: $($mbx.PrimarySmtpAddress)" -ForegroundColor White
            Write-Host "  Email Addresses with old domain:" -ForegroundColor White
            $mbx.EmailAddresses | Where-Object { Test-ContainsOldDomain -Address $_ -OldDomains $OldDomain } | 
                ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
        }

        # Step 1: Update UPN if it uses old domain
        $UPNNeedsUpdate = $false
        foreach ($oldDom in $OldDomain) {
            if ($mbx.UserPrincipalName -match "@$oldDom") {
                $UPNNeedsUpdate = $true
                break
            }
        }
        
        if ($UPNNeedsUpdate) {
            $NewUPN = Replace-Domain -Address $mbx.UserPrincipalName -OldDomains $OldDomain -NewDomainTarget $NewDomain
            
            if ($WhatIfPreference -or $PSCmdlet.ShouldProcess($mbx.UserPrincipalName, "Update UPN to $NewUPN")) {
                if ($WhatIfPreference) {
                    Write-Host "WHATIF: Would update UPN: $($mbx.UserPrincipalName) -> $NewUPN" -ForegroundColor Green
                } else {
                    # Update via Graph (more reliable for UPN changes)
                    $GraphUser = Get-MgUser -Filter "userPrincipalName eq '$($mbx.UserPrincipalName)'" -ErrorAction SilentlyContinue
                    if ($GraphUser) {
                        Update-MgUser -UserId $GraphUser.Id -UserPrincipalName $NewUPN
                        $Changes.NewUPN = $NewUPN
                        Write-LogEntry "Updated UPN: $($mbx.UserPrincipalName) -> $NewUPN"
                        
                        # Wait for AD sync
                        Start-Sleep -Seconds 3
                    } else {
                        Write-LogEntry "Warning: Could not find Graph user for $($mbx.UserPrincipalName)"
                        $Changes.NewUPN = $mbx.UserPrincipalName
                    }
                }
            }
        } else {
            $Changes.NewUPN = $mbx.UserPrincipalName
        }

        # Step 2: Clean up ALL email addresses and set new primary SMTP
        $OriginalProxies = $mbx.EmailAddresses
        $CleanProxies = Remove-OldDomainProxies -ProxyAddresses $OriginalProxies -OldDomains $OldDomain
        
        # Generate new primary SMTP address
        $NewPrimarySmtp = if ($Changes.NewUPN) {
            ($Changes.NewUPN).ToLower()
        } else {
            # Replace domain in current UPN
            $TempUPN = $mbx.UserPrincipalName
            foreach ($oldDom in $OldDomain) {
                $TempUPN = $TempUPN -replace "@$oldDom", "@$NewDomain"
            }
            $TempUPN.ToLower()
        }
        
        # Ensure primary SMTP is in the proxy list
        $PrimaryProxySmtp = "SMTP:$NewPrimarySmtp"
        if ($CleanProxies -notcontains $PrimaryProxySmtp) {
            $CleanProxies += $PrimaryProxySmtp
        }
        
        if ($WhatIfPreference -or $PSCmdlet.ShouldProcess($mbx.DisplayName, "Update email addresses")) {
            if ($WhatIfPreference) {
                Write-Host "`nWHATIF: Would update mailbox $($mbx.DisplayName):" -ForegroundColor Green
                Write-Host "  Primary SMTP: $($mbx.PrimarySmtpAddress) -> $NewPrimarySmtp" -ForegroundColor Green
                Write-Host "  Would remove $($Changes.OldProxyCount) old domain proxies" -ForegroundColor Green
                Write-Host "  New proxy count: $($CleanProxies.Count)" -ForegroundColor Green
                Write-Host "  Remaining email addresses:" -ForegroundColor Green
                $CleanProxies | ForEach-Object { Write-Host "    $_" -ForegroundColor Cyan }
            } else {
                # Determine identity to use for mailbox operations
                $IdentityToUse = if ($Changes.NewUPN) { $Changes.NewUPN } else { $mbx.UserPrincipalName }
                
                # FIXED CODE SECTION - Improved handling for different mailbox types
                try {
                    # Try first with EmailAddresses only (this works for all mailbox types)
                    Set-Mailbox -Identity $IdentityToUse -EmailAddresses $CleanProxies -Force -ErrorAction Stop
                    
                    # Then try to set the primary SMTP using either PrimarySmtpAddress or WindowsEmailAddress
                    try {
                        Set-Mailbox -Identity $IdentityToUse -PrimarySmtpAddress $NewPrimarySmtp -Force -ErrorAction Stop
                        Write-LogEntry "Updated primary SMTP address using PrimarySmtpAddress parameter"
                    } catch {
                        # If PrimarySmtpAddress fails, try WindowsEmailAddress instead
                        try {
                            Set-Mailbox -Identity $IdentityToUse -WindowsEmailAddress $NewPrimarySmtp -Force -ErrorAction Stop
                            Write-LogEntry "Updated primary SMTP address using WindowsEmailAddress parameter"
                        } catch {
                            # If both methods fail, try setting it through the EmailAddresses collection
                            Write-LogEntry "Warning: Could not set primary SMTP address directly. Using EmailAddresses collection method."
                            
                            # Get updated mailbox to see current email addresses
                            $UpdatedMailbox = Get-Mailbox -Identity $IdentityToUse
                            $CurrentProxies = $UpdatedMailbox.EmailAddresses
                            
                            # Remove any existing primary SMTP
                            $NonPrimaryProxies = $CurrentProxies | Where-Object { $_ -notmatch "^SMTP:" }
                            
                            # Add new primary SMTP
                            $NewProxyList = @("SMTP:$NewPrimarySmtp") + $NonPrimaryProxies
                            
                            # Update with new proxy list
                            Set-Mailbox -Identity $IdentityToUse -EmailAddresses $NewProxyList -Force
                        }
                    }
                } catch {
                    # If Set-Mailbox fails completely, try Set-MailUser as fallback
                    try {
                        Set-MailUser -Identity $IdentityToUse -EmailAddresses $CleanProxies -ErrorAction Stop
                        
                        # Try to set primary SMTP for mail user
                        try {
                            Set-MailUser -Identity $IdentityToUse -PrimarySmtpAddress $NewPrimarySmtp -ErrorAction Stop
                        } catch {
                            Write-LogEntry "Warning: Could not set primary SMTP for mail user $($mbx.DisplayName)"
                        }
                    } catch {
                        # If both Set-Mailbox and Set-MailUser fail, log the error
                        throw $_  # Re-throw the exception to be caught by the outer try-catch
                    }
                }
                
                # Show updated state for single user tests
                if ($Identity) {
                    Start-Sleep -Seconds 2  # Allow time for updates
                    try {
                        $UpdatedMailbox = Get-Mailbox -Identity $IdentityToUse
                        Write-Host "`nUpdated state:" -ForegroundColor Green
                        Write-Host "  UPN: $($UpdatedMailbox.UserPrincipalName)" -ForegroundColor White
                        Write-Host "  Primary SMTP: $($UpdatedMailbox.PrimarySmtpAddress)" -ForegroundColor White
                        Write-Host "  All email addresses:" -ForegroundColor White
                        $UpdatedMailbox.EmailAddresses | ForEach-Object { 
                            $Color = if (Test-ContainsOldDomain -Address $_ -OldDomains $OldDomain) { "Red" } else { "Green" }
                            Write-Host "    $_" -ForegroundColor $Color 
                        }
                    } catch {
                        Write-LogEntry "Could not retrieve updated mailbox state: $($_.Exception.Message)" -Error
                    }
                }
            }
        }
        
        $Changes.NewPrimarySMTP = $NewPrimarySmtp
        $Changes.NewProxyCount = $CleanProxies.Count
        $Changes.Status = if ($WhatIfPreference) { "WhatIf" } else { "Success" }
        
        if (-not $WhatIfPreference) {
            Write-LogEntry "Updated mailbox $($mbx.DisplayName): Primary SMTP -> $NewPrimarySmtp, Removed $($Changes.OldProxyCount) old proxies"
        }
        
        $UserChanges += [PSCustomObject]$Changes
    }
    catch {
        $Changes.Status = "Failed: $($_.Exception.Message)"
        $UserChanges += [PSCustomObject]$Changes
        Write-LogEntry "ERROR processing $($mbx.DisplayName): $($_.Exception.Message)" -Error
        
        # Additional debugging info
        Write-LogEntry "Mailbox details - Type: $($mbx.RecipientTypeDetails), UPN: $($mbx.UserPrincipalName), Primary: $($mbx.PrimarySmtpAddress)" -Error
    }
}
#endregion

#region 5. DISTRIBUTION GROUPS CLEANUP (Skip in single user mode)
if (-not $Identity) {
    Write-Host "`n=== PHASE 2: DISTRIBUTION GROUPS ===" -ForegroundColor Magenta

    $DistributionGroups = Get-ObjectsWithOldDomain -ObjectType 'DistributionGroup' -OldDomains $OldDomain

    Write-LogEntry "Found $($DistributionGroups.Count) distribution groups with old domain references"

    $DGChanges = @()
    foreach ($dg in $DistributionGroups) {
        try {
            Write-Host "Processing DG: $($dg.DisplayName)" -ForegroundColor Cyan
            
            # Clean proxy addresses
            $CleanProxies = Remove-OldDomainProxies -ProxyAddresses $dg.EmailAddresses -OldDomains $OldDomain
            
            # Generate new primary SMTP
            $NewPrimarySmtp = ($dg.Alias + "@$NewDomain").ToLower()
            $CleanProxies += "SMTP:$NewPrimarySmtp"
            
            if ($WhatIfPreference -or $PSCmdlet.ShouldProcess($dg.DisplayName, "Update distribution group email addresses")) {
                if ($WhatIfPreference) {
                    Write-Host "WHATIF: Would update DG $($dg.DisplayName): $($dg.PrimarySmtpAddress) -> $NewPrimarySmtp" -ForegroundColor Green
                } else {
                    Set-DistributionGroup -Identity $dg.Identity `
                                         -PrimarySmtpAddress $NewPrimarySmtp `
                                         -EmailAddresses $CleanProxies
                }
            }
            
            $DGChanges += [PSCustomObject]@{
                ObjectType = 'DistributionGroup'
                Identity = $dg.DisplayName
                OldPrimarySMTP = $dg.PrimarySmtpAddress
                NewPrimarySMTP = $NewPrimarySmtp
                Status = if ($WhatIfPreference) { "WhatIf" } else { "Success" }
            }
            
            if (-not $WhatIfPreference) {
                Write-LogEntry "Updated DG $($dg.DisplayName): $($dg.PrimarySmtpAddress) -> $NewPrimarySmtp"
            }
        }
        catch {
            $DGChanges += [PSCustomObject]@{
                ObjectType = 'DistributionGroup'
                Identity = $dg.DisplayName
                Status = "Failed: $($_.Exception.Message)"
            }
            Write-LogEntry "ERROR processing DG $($dg.DisplayName): $($_.Exception.Message)" -Error
        }
    }
} else {
    Write-Host "`n=== SKIPPING DISTRIBUTION GROUPS (Single User Mode) ===" -ForegroundColor Yellow
    $DGChanges = @()
}
#endregion

#region 6. MICROSOFT 365 GROUPS CLEANUP (Skip in single user mode)
if (-not $Identity) {
    Write-Host "`n=== PHASE 3: MICROSOFT 365 GROUPS ===" -ForegroundColor Magenta

    $M365Groups = Get-ObjectsWithOldDomain -ObjectType 'UnifiedGroup' -OldDomains $OldDomain

    Write-LogEntry "Found $($M365Groups.Count) Microsoft 365 groups with old domain references"

    $M365Changes = @()
    foreach ($group in $M365Groups) {
        try {
            Write-Host "Processing M365 Group: $($group.DisplayName)" -ForegroundColor Cyan
            
            # Clean proxy addresses
            $CleanProxies = Remove-OldDomainProxies -ProxyAddresses $group.EmailAddresses -OldDomains $OldDomain
            
            # Generate new primary SMTP
            $NewPrimarySmtp = ($group.Alias + "@$NewDomain").ToLower()
            $CleanProxies += "SMTP:$NewPrimarySmtp"
            
            if ($WhatIfPreference -or $PSCmdlet.ShouldProcess($group.DisplayName, "Update M365 group email addresses")) {
                if ($WhatIfPreference) {
                    Write-Host "WHATIF: Would update M365 Group $($group.DisplayName): $($group.PrimarySmtpAddress) -> $NewPrimarySmtp" -ForegroundColor Green
                } else {
                    Set-UnifiedGroup -Identity $group.Identity `
                                    -PrimarySmtpAddress $NewPrimarySmtp `
                                    -EmailAddresses $CleanProxies
                    
                    # Also update via Graph
                    try {
                        $GraphGroup = Get-MgGroup -Filter "mailNickname eq '$($group.Alias)'"
                        if ($GraphGroup) {
                            Update-MgGroup -GroupId $GraphGroup.Id -Mail $NewPrimarySmtp
                        }
                    }
                    catch {
                        Write-LogEntry "Warning: Could not update Graph group mail for $($group.DisplayName): $($_.Exception.Message)"
                    }
                }
            }
            
            $M365Changes += [PSCustomObject]@{
                ObjectType = 'M365Group'
                Identity = $group.DisplayName
                OldPrimarySMTP = $group.PrimarySmtpAddress
                NewPrimarySMTP = $NewPrimarySmtp
                Status = if ($WhatIfPreference) { "WhatIf" } else { "Success" }
            }
            
            if (-not $WhatIfPreference) {
                Write-LogEntry "Updated M365 Group $($group.DisplayName): $($group.PrimarySmtpAddress) -> $NewPrimarySmtp"
            }
        }
        catch {
            $M365Changes += [PSCustomObject]@{
                ObjectType = 'M365Group'
                Identity = $group.DisplayName
                Status = "Failed: $($_.Exception.Message)"
            }
            Write-LogEntry "ERROR processing M365 Group $($group.DisplayName): $($_.Exception.Message)" -Error
        }
    }
} else {
    Write-Host "`n=== SKIPPING MICROSOFT 365 GROUPS (Single User Mode) ===" -ForegroundColor Yellow
    $M365Changes = @()
}
#endregion

#region 7. EXPORT CHANGE LOGS
Write-Host "`n=== EXPORTING CHANGE LOGS ===" -ForegroundColor Magenta

$AllChanges = @()
$AllChanges += $UserChanges
$AllChanges += $DGChanges
$AllChanges += $M365Changes

if ($AllChanges.Count -gt 0 -and -not $WhatIfPreference) {
    $AllChanges | Export-Csv $ChangeLog -NoTypeInformation
    Write-LogEntry "Exported $($AllChanges.Count) changes to $ChangeLog"
} elseif ($WhatIfPreference) {
    Write-Host "WhatIf mode: $($AllChanges.Count) changes would be made" -ForegroundColor Green
    if ($AllChanges.Count -gt 0) {
        $AllChanges | Format-Table -AutoSize
    }
}
#endregion

#region 8. FINAL SUMMARY
Write-Host "`n=== SUMMARY ===" -ForegroundColor Magenta

if ($WhatIfPreference) {
    Write-Host "WhatIf Mode Summary:" -ForegroundColor Green
    Write-Host "  Mailboxes to update: $($UserChanges.Count)" -ForegroundColor White
    if (-not $Identity) {
        Write-Host "  Distribution Groups to update: $($DGChanges.Count)" -ForegroundColor White
        Write-Host "  M365 Groups to update: $($M365Changes.Count)" -ForegroundColor White
    }
    Write-Host "  Total objects: $($AllChanges.Count)" -ForegroundColor White
    Write-Host "`nTo execute these changes, run the script without -WhatIf parameter." -ForegroundColor Yellow
} else {
    $SuccessCount = ($AllChanges | Where-Object { $_.Status -eq "Success" }).Count
    $FailureCount = ($AllChanges | Where-Object { $_.Status -like "Failed*" }).Count
    
    Write-Host "Execution Summary:" -ForegroundColor Green
    Write-Host "  Successful updates: $SuccessCount" -ForegroundColor Green
    Write-Host "  Failed updates: $FailureCount" -ForegroundColor $(if ($FailureCount -gt 0) { "Red" } else { "Green" })
    Write-Host "  Total processed: $($AllChanges.Count)" -ForegroundColor White
    
    if ($Identity) {
        Write-Host "`nSingle User Test Mode Complete!" -ForegroundColor Cyan
        Write-Host "If this test was successful, you can run the full tenant cleanup by removing the -Identity parameter." -ForegroundColor Yellow
    }
    
    Write-Host "`nLog files created:" -ForegroundColor Cyan
    Write-Host "  Changes: $ChangeLog" -ForegroundColor White
    Write-Host "  Errors: $ErrorLog" -ForegroundColor White
    Write-Host "  Validation: $ValidationLog" -ForegroundColor White
    
    if ($FailureCount -eq 0) {
        Write-Host "`nAll operations completed successfully!" -ForegroundColor Green
        if (-not $Identity) {
            Write-Host "Domain cleanup is complete. You should now be able to remove the old domain(s) from your tenant." -ForegroundColor Green
        }
    } else {
        Write-Host "`nSome operations failed. Please review the error log before proceeding." -ForegroundColor Yellow
    }
}

# Cleanup connections
if (-not $WhatIfPreference) {
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
}

Write-Host "`nDomain cleanup script completed!" -ForegroundColor Magenta
#endregion