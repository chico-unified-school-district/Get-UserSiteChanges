<# 
.SYNOPSIS
Pull data from Employee Database (Escape Online) and 
compare to previous entries for updated site codes. 
Create a KACE helpdesk ticket via email when a site code is updated.
.DESCRIPTION
.EXAMPLE
.\Get-UserSiteChanges.PS1 -SQLServer EmpServer -SQLDatabase EmpDB -SQLCredential $SQLCredObject -EmailCredential $EmailCredObject -TargetAddress servicedesk@kbox1.company.com -WhatIf
.INPUTS
 SQL Server Name
 SQL Server Database
 SQL Server Credential object
 Email Credential object for account that is allowed to send to an target address such as 'servicedesk@kbox1.company.com'
 Target Address that is allowed to create helpdesk tickets in KACE. Example: 'servicedesk@kbox1.company.com'
.NOTES
#>
#Requires -Version 5.0

[cmdletbinding()]
param ( 
 [Parameter(Position = 0, Mandatory = $True)]
 [Alias('DC')]
 [ValidateScript( { Test-Connection -ComputerName $_ -Quiet -Count 1 })]
 [string]$DomainController,
 [Parameter(Position = 1, Mandatory = $True)]
 [System.Management.Automation.PSCredential]$ADCredential,
 [Parameter(Mandatory = $True)]
 [Alias('Server')]
 [string]$SQLServer,
 [Alias('Database')]
 [Parameter(Mandatory = $True)]
 [string]$SQLDatabase,
 # Credential object with database select permission.
 [Parameter(Mandatory = $True)]
 [System.Management.Automation.PSCredential]$SQLCredential,
 [Parameter(Mandatory = $True)]
 [System.Management.Automation.PSCredential]$EmailCredential,
 [string]$TargetAddress,
 [Parameter(Mandatory = $True)]
 [string]$TicketOwner,
 [SWITCH]$WhatIf
)
Clear-Host


# AD Domain Controller Session
$adCmdLets = 'Get-ADUser'
$adSession = New-PSSession -ComputerName $DomainController -Credential $ADCredential
Import-PSSession -Session $adSession -Module ActiveDirectory -CommandName $adCmdLets -AllowClobber | Out-Null

$lookUpTable = Get-Content -Path .\json\lookup-table.json -Raw | ConvertFrom-Json

# Imported Functions
. .\lib\Add-Log.ps1
. .\lib\Compare-Object2.ps1
. .\lib\Invoke-SqlCommand.PS1

# =========================
# Processing
# =========================

$query = Get-Content -Path .\sql\lookup-employee-job-data.sql -Raw

$currentData = Invoke-SqlCommand -Server $SQLServer -Database $SQLDatabase -Credential $SQLCredential -Query $query

if (Test-Path -Path .\data\previous-data.csv) {
 $previousData = Import-Csv -Path .\data\previous-data.csv
}
else {
 Add-Log warning 'No Previous Data. Creating file for next run and exiting...'
 if (!(Test-Path -Path .\data)) { New-Item -Name data -ItemType Directory -Confirm:$false }
 $currentData | Select-Object -Property employeeId, departmentNumber | Export-Csv -Path .\data\previous-data.csv -Force
 EXIT
}


$departmentNumberChanges = Compare-Object -ReferenceObject $previousData -DifferenceObject $currentData -Property employeeId, departmentNumber | 
Where-Object { $_.SideIndicator -eq '=>' }

# $departmentNumberChanges
$messageTemplate = Get-Content -Path .\messages\site-move.txt -Raw

# Check if the sitecode is a proper integer and catch any blank entries
foreach ( $item in $departmentNumberChanges ) {
 write-Debug ( "========`nProcess {0}?" -f $item.employeeID )
 try {
  if ([int]$item.departmentNumber.trim() -is 'Int') {
   $newSite = $lookUpTable.Where( { [int]$_.departmentNumber -eq [int]$item.departmentNumber } )
  } 
 }
 catch { Write-Verbose ('Escape Department Number ({0}) for {1} is not an integer.' -f $item.departmentNumber, $item.employeeID) ; continue }

 # Begin Process Rows
 $employeeId = $item.employeeId

 $userObj = Get-ADUser -LDAPFilter "(employeeId=$employeeId)" -Properties 'Title', 'departmentNumber', 'employeeID'
 if (!$userObj) { continue }

 $escapeData = $currentData.Where( { [int]$_.employeeID -eq [int]$employeeId } )

 # Check if User Object Department Number is an integer
 try {
  if ( ([int]$userObj.departmentNumber.trim() -is 'Int') ) {
   $oldSite = $lookUpTable.Where( { [int]$_.departmentNumber -eq [int]$userObj.departmentNumber[0] })
  } 
 }
 catch { Write-Verbose ('User Object Department Number ({0}) for {1} is not an integer.' -f $userObj.departmentNumber, $item.employeeID) }


 if (($null -eq $escapeData.description) -or ($escapeData.description -match "^\s")) {
  Add-Log warning ('No Job Description for {0} {1}. Skipping.' -f $employeeId, $userObj.name) 
  continue
 }
 if ($oldSite.SiteDescr -eq $newSite.SiteDescr) {
  Write-Verbose 'Site Descriptions already match. Skipping.'
  continue
 }

 $messageData = @(
  $employeeId
  $userObj.Name
  $userObj.samAccountName
  $userObj.Title
  $escapeData.description
  $oldSite.SiteDescr
  $newSite.SiteDescr
  $oldSite.FileServer
  $newSite.FileServer
  $userObj.DistinguishedName
  $ENV:USERNAME
  $ENV:COMPUTERNAME
  $TicketOwner
 )
 $messageBody = $messageTemplate -f $messageData

 $mailParams = @{
  To         = ('<{0}>' -f $TargetAddress)
  From       = $EmailCredential.Username
  Subject    = ( "User Site Change - {0} moved to {1}" -f $user.name, $siteDescr)
  Body       = $messageBody
  BodyAsHtml = $false
  SMTPServer = 'smtp.office365.com'
  Cred       = $EmailCredential
  UseSSL     = $True
  Port       = 587
 }
 Add-Log email ("{0},Helpdesk Ticket Created" -f ( "User Site Change - {0} moved to {1}" -f $userObj.name, $newSite.SiteDescr)) -WhatIf:$WhatIf
 if ( $WhatIf ) { $to, $messageBody }
 else { Send-MailMessage @mailParams }
} # End Process Rows
 
# Export current data as a reference for the next run
if (!$WhatIf) {
 $currentData | Select-Object -Property employeeId, departmentNumber | Export-Csv -Path .\data\previous-data.csv -Force
}