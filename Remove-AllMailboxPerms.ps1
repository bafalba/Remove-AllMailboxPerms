<# 
.SYNOPSIS
Remove ALL Exchange Online mailbox permissions (FullAccess, SendAs, SendOnBehalf) 
for a single user across every mailbox.

.PARAMETER User
The user to remove permissions from (email/UPN).

.PARAMETER WhatIf
Preview only; shows what would be removed without making changes.

.EXAMPLE
.\Remove-AllMailboxPerms.ps1 -User "someone@contoso.com"
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$User,

    [switch]$WhatIf
)

# 1) Import & connect
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Connect-ExchangeOnline -ShowBanner:$false

if (-not $User -or $User -eq "") {
    $User = Read-Host "Enter the user (email/UPN) to remove permissions from"
}

try {
    # Resolve the trustee to a recipient (more reliable than raw string)
    $trustee = Get-Recipient -Identity $User -ErrorAction Stop

    Write-Host "Removing permissions for trustee: $($trustee.PrimarySmtpAddress)" -ForegroundColor Cyan

    # 2) Get all mailboxes (shared + user + resource; adjust filter if you want only shared)
    $mailboxes = Get-Mailbox -ResultSize Unlimited

    # Tracking results
    $results = @()

    foreach ($mbx in $mailboxes) {
        $changed = $false
        $mbxId = $mbx.Identity

        # --- FullAccess ---
        $faPerms = Get-MailboxPermission -Identity $mbxId | Where-Object { $_.User -eq $trustee.PrimarySmtpAddress -and $_.AccessRights -contains 'FullAccess' }
        if ($faPerms) {
            try {
                if ($WhatIf) {
                    Write-Host "[WhatIf] Remove FullAccess from $mbxId" -ForegroundColor Yellow
                } else {
                    Remove-MailboxPermission -Identity $mbxId -User $trustee.PrimarySmtpAddress `
                        -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop
                }
                $changed = $true
                $results += [pscustomobject]@{ Mailbox=$mbxId; Permission='FullAccess'; Status= $(if($WhatIf){'Would Remove'} else {'Removed'}) }
            } catch { }
        }

        # --- SendAs ---
        $saPerms = Get-RecipientPermission -Identity $mbxId | Where-Object { $_.Trustee -eq $trustee.PrimarySmtpAddress -and $_.AccessRights -contains 'SendAs' }
        if ($saPerms) {
            try {
                if ($WhatIf) {
                    Write-Host "[WhatIf] Remove SendAs from $mbxId" -ForegroundColor Yellow
                } else {
                    Remove-RecipientPermission -Identity $mbxId -Trustee $trustee.PrimarySmtpAddress `
                        -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                }
                $changed = $true
                $results += [pscustomobject]@{ Mailbox=$mbxId; Permission='SendAs'; Status= $(if($WhatIf){'Would Remove'} else {'Removed'}) }
            } catch { }
        }

        # --- SendOnBehalf ---
        if ($mbx.GrantSendOnBehalfTo -contains $trustee.PrimarySmtpAddress) {
            try {
                if ($WhatIf) {
                    Write-Host "[WhatIf] Remove SendOnBehalf from $mbxId" -ForegroundColor Yellow
                } else {
                    Set-Mailbox -Identity $mbxId -GrantSendOnBehalfTo @{remove=$trustee.PrimarySmtpAddress} -ErrorAction Stop
                }
                $changed = $true
                $results += [pscustomobject]@{ Mailbox=$mbxId; Permission='SendOnBehalf'; Status= $(if($WhatIf){'Would Remove'} else {'Removed'}) }
            } catch { }
        }

        if (-not $changed) {
            $results += [pscustomobject]@{ Mailbox=$mbxId; Permission='(none matched)'; Status='No change' }
        }
    }

    # 3) Output a nice summary

    $results | Sort-Object Mailbox, Permission | Format-Table -AutoSize

    # --- Show all mailbox permissions for the user ---
    Write-Host "\nCurrent permissions for $($trustee.PrimarySmtpAddress):" -ForegroundColor Cyan

    # FullAccess
    Write-Host "\nFullAccess permissions:" -ForegroundColor Yellow
    Get-MailboxPermission -ResultSize Unlimited | Where-Object { $_.User -eq $trustee.PrimarySmtpAddress } | Format-Table Identity, AccessRights, Deny, IsInherited -AutoSize

    # SendAs
    Write-Host "\nSendAs permissions:" -ForegroundColor Yellow
    Get-RecipientPermission -ResultSize Unlimited | Where-Object { $_.Trustee -eq $trustee.PrimarySmtpAddress } | Format-Table Identity, AccessRights, Trustee -AutoSize

    # SendOnBehalf
    Write-Host "\nSendOnBehalf permissions:" -ForegroundColor Yellow
    Get-Mailbox -ResultSize Unlimited | Where-Object { $_.GrantSendOnBehalfTo -contains $trustee.PrimarySmtpAddress } | Select-Object Identity, GrantSendOnBehalfTo | Format-Table -AutoSize

    # --- Wait for user input to exit ---
    Write-Host "\nType Q to exit..." -ForegroundColor Green
    do {
        $input = Read-Host ""
    } while ($input.ToUpper() -ne "Q")

    } finally {
        Disconnect-ExchangeOnline -Confirm:$false
    }
