
# Universal Domain Cleanup Script

**Version:** 1.0  
**Last Updated:** 2025-07-30  
**Author:** furena  

## Overview

The Universal Domain Cleanup Script is a comprehensive PowerShell solution for tenant-to-tenant migrations. It completely removes all references to specified old domain(s) from a Microsoft 365 tenant, preparing the domains for transfer to a new tenant.

## Features

- **Complete Domain Cleanup**: Removes ALL references to old domain(s) across your tenant
- **Multiple Domain Support**: Clean up single or multiple domains in one operation
- **Comprehensive Scope**: Updates UPNs, SMTP addresses, SIP addresses, and all proxy addresses
- **Identity Types**: Processes user mailboxes, shared mailboxes, distribution groups, and Microsoft 365 Groups
- **Test Mode**: Supports running against a single user for testing before full tenant execution
- **WhatIf Support**: Preview all changes before making them
- **Detailed Logging**: Comprehensive logs for changes, errors, and validation

## Prerequisites

- Exchange Online PowerShell Module
- Microsoft Graph PowerShell Module
- Appropriate administrative permissions in your Microsoft 365 tenant

```powershell
# Required modules
Install-Module -Name ExchangeOnlineManagement -Force
Install-Module -Name Microsoft.Graph -Force
```

## Usage Examples

### Interactive Mode (Prompts for Domain)

```powershell
.\UniversalDomainCleanup.ps1
```

### Clean Up Single Domain

```powershell
.\UniversalDomainCleanup.ps1 -OldDomain "contoso.com"
```

### Test on a Single User

```powershell
.\UniversalDomainCleanup.ps1 -OldDomain "contoso.com" -Identity "user@contoso.com"
```

### Preview Changes Without Making Them

```powershell
.\UniversalDomainCleanup.ps1 -OldDomain "contoso.com" -WhatIf
```

### Clean Multiple Domains to a Specific New Domain

```powershell
.\UniversalDomainCleanup.ps1 -OldDomain @("contoso.com","fabrikam.com") -NewDomain "newcompany.onmicrosoft.com"
```

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| OldDomain | string[] | No | Domain(s) to be cleaned up (e.g., 'contoso.com' or @('contoso.com','fabrikam.com')) |
| NewDomain | string | No | New domain to use (defaults to tenant's onmicrosoft.com domain) |
| Identity | string | No | Test against a single user/mailbox (UPN, email, or display name) |
| LogPath | string | No | Path for log files (defaults to current directory) |

## Troubleshooting

If you encounter an error when using single identity test mode related to the `PrimarySmtpAddress` parameter, this is due to mailbox type incompatibility. The script includes fallback methods to handle these scenarios:

1. It first tries to update using the EmailAddresses parameter
2. Then attempts to set the primary address using PrimarySmtpAddress
3. Falls back to WindowsEmailAddress if needed
4. As a last resort, manipulates the EmailAddresses collection directly

## Logs Generated

The script creates three log files in the specified LogPath:

1. **ChangeLog** - Records all successful changes made
2. **ErrorLog** - Records any errors encountered
3. **ValidationLog** - Detailed validation steps performed

## Important Notes

- Always run with `-WhatIf` first to preview changes
- Back up your tenant before running in full mode
- Test on a single user before performing a complete tenant cleanup
- Domain transfer should be performed soon after cleanup

## Disclaimer

This script makes significant changes to your Microsoft 365 tenant. Always:
- Test in a lab environment first
- Back up your tenant before running
- Run in WhatIf mode first
- Test on a single user before running on the entire tenant

## License

This script is provided as-is with no warranties or guarantees.
