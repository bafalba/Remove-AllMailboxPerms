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
    # Resolve the Trainee to a recipient (more reliable than raw string)
    $Trainee = Get-EXORecipient -Identity $User -ErrorAction Stop

    Write-Host "Removing permissions for Trainee: $($Trainee.PrimarySmtpAddress)" -ForegroundColor Cyan

    # Get all shared mailboxes (fast EXO cmdlet)
    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox

    # Tracking results
    $results = @()

    foreach ($mbx in $mailboxes) {
        $changed = $false
        $mbxId = $mbx.Identity
        Write-Host "Processing mailbox: $($mbx.PrimarySmtpAddress) ($mbxId)" -ForegroundColor Cyan


        # --- FullAccess ---
        $faPerms = Get-EXOMailboxPermission -Identity $mbxId -User $Trainee.PrimarySmtpAddress -ErrorAction SilentlyContinue
        if ($faPerms) {
            try {
                if ($WhatIf) {
                    Write-Host "[WhatIf] Remove FullAccess from $mbxId" -ForegroundColor Yellow
                } else {
                    Remove-MailboxPermission -Identity $mbxId -User $Trainee.PrimarySmtpAddress `
                        -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop
                }
                $changed = $true
                $results += [pscustomobject]@{ Mailbox=$mbx.PrimarySmtpAddress; Identity=$mbxId; Permission='FullAccess'; Status= $(if($WhatIf){'Would Remove'} else {'Removed'}) }
            } catch {
                Write-Host ("Error removing FullAccess from {0}: {1}" -f $mbxId, $_) -ForegroundColor Red
            }
        }

        # --- SendAs ---
        $saPerms = Get-EXORecipientPermission -Identity $mbxId -Trustee $Trainee.PrimarySmtpAddress -ErrorAction SilentlyContinue
        if ($saPerms) {
            try {
                if ($WhatIf) {
                    Write-Host "[WhatIf] Remove SendAs from $mbxId" -ForegroundColor Yellow
                } else {
                    Remove-RecipientPermission -Identity $mbxId -Trustee $Trainee.PrimarySmtpAddress `
                        -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                }
                $changed = $true
                $results += [pscustomobject]@{ Mailbox=$mbx.PrimarySmtpAddress; Identity=$mbxId; Permission='SendAs'; Status= $(if($WhatIf){'Would Remove'} else {'Removed'}) }
            } catch {
                Write-Host ("Error removing SendAs from {0}: {1}" -f $mbxId, $_) -ForegroundColor Red
            }
        }

        # --- SendOnBehalf ---
        $mbxDetails = Get-EXOMailbox -Identity $mbxId -Properties GrantSendOnBehalfTo -ErrorAction SilentlyContinue
        if ($mbxDetails -and $mbxDetails.GrantSendOnBehalfTo -contains $Trainee.PrimarySmtpAddress) {
            try {
                if ($WhatIf) {
                    Write-Host "[WhatIf] Remove SendOnBehalf from $mbxId" -ForegroundColor Yellow
                } else {
                    Set-EXOMailbox -Identity $mbxId -GrantSendOnBehalfTo @{remove=$Trainee.PrimarySmtpAddress} -ErrorAction Stop
                }
                $changed = $true
                $results += [pscustomobject]@{ Mailbox=$mbx.PrimarySmtpAddress; Identity=$mbxId; Permission='SendOnBehalf'; Status= $(if($WhatIf){'Would Remove'} else {'Removed'}) }
            } catch {
                Write-Host ("Error removing SendOnBehalf from {0}: {1}" -f $mbxId, $_) -ForegroundColor Red
            }
        }

        if (-not $changed) {
            $results += [pscustomobject]@{ Mailbox=$mbx.PrimarySmtpAddress; Identity=$mbxId; Permission='(none matched)'; Status='No change' }
        }
    }


    # 3) Output a nice summary
    $results | Sort-Object Mailbox, Permission | Format-Table -AutoSize

    # Export results to CSV
    $logPath = "C:\Users\bafalba\Documents\Scripts\Logs"
    if (!(Test-Path $logPath)) { New-Item -Path $logPath -ItemType Directory | Out-Null }
    $csvFile = Join-Path $logPath ("RemovedMailboxPermissions_$($Trainee.PrimarySmtpAddress -replace '[^\w@.-]','_').csv")
    $results | Export-Csv $csvFile -NoTypeInformation
    Write-Host "Results exported to $csvFile" -ForegroundColor Green


    # --- Show all mailbox permissions for the user ---
    Write-Host "\nCurrent permissions for $($Trainee.PrimarySmtpAddress):" -ForegroundColor Cyan

    foreach ($mbx in $mailboxes) {
        Write-Host ("\nMailbox: {0}" -f $mbx.PrimarySmtpAddress) -ForegroundColor Yellow
        $fa = Get-EXOMailboxPermission -Identity $mbx.Identity -User $Trainee.PrimarySmtpAddress -ErrorAction SilentlyContinue
        if ($fa) {
            Write-Host "FullAccess:" -ForegroundColor Yellow
            $fa | Format-Table Identity, AccessRights, Deny, IsInherited -AutoSize
        }
        $sa = Get-EXORecipientPermission -Identity $mbx.Identity -Trustee $Trainee.PrimarySmtpAddress -ErrorAction SilentlyContinue
        if ($sa) {
            Write-Host "SendAs:" -ForegroundColor Yellow
            $sa | Format-Table Identity, AccessRights, Trustee -AutoSize
        }
        $mbxDetails = Get-EXOMailbox -Identity $mbx.Identity -Properties GrantSendOnBehalfTo -ErrorAction SilentlyContinue
        if ($mbxDetails -and $mbxDetails.GrantSendOnBehalfTo -contains $Trainee.PrimarySmtpAddress) {
            Write-Host "SendOnBehalf:" -ForegroundColor Yellow
            Write-Host ($mbxDetails.GrantSendOnBehalfTo | Out-String)
        }
    }

    # --- Wait for user input to exit ---
    Write-Host "\nPress Y to exit..." -ForegroundColor Green
    do {
        $input = Read-Host ""
    } while ($input.ToUpper() -ne "Y")

    } finally {
        Disconnect-ExchangeOnline -Confirm:$false
    }
