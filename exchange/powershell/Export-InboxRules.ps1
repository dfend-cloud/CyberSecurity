<#
.SYNOPSIS
  Export Inbox rules for all mailboxes in Exchange Online. 
  
.DESCRIPTION
  Export InboxRules for all mailboxes to a csv file. Requires ExchangeOnlineManagement module. You must call Connect-ExchangeOnline first. 
  
.PARAMETER csvPath
  This parameter is the path where the csv will be saved. This path should not inlcude a file name or trailing backslash.  

.NOTES
  Author:         Declan Turley 
    
.EXAMPLE
  Export-InboxRules.ps1 -csvPath C:\temp
#>

#Parameter to specify CSV Output
param(
    [Parameter(Mandatory = $true)]
    [string] $csvPath
  )

$allMailboxes = Get-Mailbox -ResultSize Unlimited
$Count = 0
foreach ($Mailbox in $allMailboxes){
    $Count++
    Write-Progress -Activity "Collecting InboxRules for $($Mailbox.UserPrincipalName)" -CurrentOperation $Mailbox -PercentComplete (($Count / $allMailboxes.count) * 100)
    Get-InboxRule -Mailbox $($Mailbox.UserPrincipalName) | Select-Object -Property @{Name="MailboxUPN"; Expression={$($Mailbox.UserPrincipalName)}},* `| Export-Csv $csvPath\InboxRuleExport.csv -Append -NoTypeInformation
}
