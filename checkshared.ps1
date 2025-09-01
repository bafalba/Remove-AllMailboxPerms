
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$false

# Prompt for email input
$user = Read-Host "Enter the user email (UPN) to check permissions for"
$user = Read-Host "Enter the user email (UPN) to check permissions for"
Write-Host "Checking shared mailboxes for: $user" -ForegroundColor Cyan

$shared = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox
$shared = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox
Write-Host "Found $($shared.Count) shared mailboxes." -ForegroundColor Yellow

$rows = @()

# FullAccess
foreach ($m in $shared) {
Write-Host "Checking FullAccess for $($m.PrimarySmtpAddress)..." -ForegroundColor DarkGray
Write-Host "Checking SendAs for $($m.PrimarySmtpAddress)..." -ForegroundColor DarkGray
Write-Host "Checking SendOnBehalf for $($m.PrimarySmtpAddress)..." -ForegroundColor DarkGray
  $p = Get-EXOMailboxPermission -Identity $m.Identity -User $user -ErrorAction SilentlyContinue
  if ($p) { $rows += [pscustomobject]@{Mailbox=$m.PrimarySmtpAddress; Permission='FullAccess'} }
}

# SendAs
foreach ($m in $shared) {
  $p = Get-EXORecipientPermission -Identity $m.Identity -Trustee $user -ErrorAction SilentlyContinue
  if ($p) { $rows += [pscustomobject]@{Mailbox=$m.PrimarySmtpAddress; Permission='SendAs'} }
}

# SendOnBehalf
foreach ($m in $shared) {
  $mbx = Get-EXOMailbox -Identity $m.Identity -Properties GrantSendOnBehalfTo -ErrorAction SilentlyContinue
  if ($mbx -and $mbx.GrantSendOnBehalfTo -contains $user) {
    $rows += [pscustomobject]@{Mailbox=$m.PrimarySmtpAddress; Permission='SendOnBehalf'}
  }
}


$rows | Sort-Object Mailbox, Permission | Format-Table -Auto
Write-Host "Permission checks complete. Preparing output..." -ForegroundColor Cyan

# Export results to CSV in specified folder
$logPath = "C:\Users\bafalba\Documents\Scripts\Logs"
if (!(Test-Path $logPath)) { New-Item -Path $logPath -ItemType Directory | Out-Null }
$csvFile = Join-Path $logPath ("SharedMailboxPermissions_$($user -replace '[^\w@.-]','_').csv")
$rows | Export-Csv $csvFile -NoTypeInformation
Write-Host "Exporting results to CSV..." -ForegroundColor Cyan
Write-Host "Results exported to $csvFile" -ForegroundColor Green

# Wait for user to press Y to close
do {
  $close = Read-Host "Press Y to close the script"
} while ($close.ToUpper() -ne "Y")

Disconnect-ExchangeOnline -Confirm:$false
