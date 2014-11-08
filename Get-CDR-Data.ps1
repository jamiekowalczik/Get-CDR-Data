#=============================================================================# 
#                                                                             # 
# Get-CDR-Data.ps1                           		                      # 
# Powershell Script automate SSH sessions to Cisco UCM.		              # 
# This script will get a list of CDR data and send the                        #
# results through email and/or and sftp session.                              #
# Author: James Kowalczik                                                     # 
# Creation Date: 11.05.2014                                                   # 
# Modified Date: 11.05.2014                                                   # 
# Version: 1.0.0                                                              # 
#                                                                             # 
#=============================================================================# 

Param(
   [String]$WhichData = "LastMonth",
   [String[]]$Extensions = @("1111","1112"),
   [HashTable]$MaskExtension = @{"1111" = "Me"; "1112" = "You"},
   [String]$MyTimezone = "Eastern Standard Time",
   [String]$SSHHostname = "10.10.10.10",
   [String]$SSHUsername = "ccmadmin",
   [String]$SSHPassword = "ccmpassword",
   [String]$OutputResultsFile = "cdr_output.csv",
   [String]$OutputResultsFileSummary = "cdr_summary.txt",
   [Bool]$SendEmailAlert = $true,
   [String]$EmailTo = "alert@me.here",
   [String]$EmailFrom = "CDR_Report@me.here",
   [String]$EmailSubject = "CDR Data and a Summary",
   [String]$EmailServer = "mail.server.com",
   [Bool]$SftpTransferFiles = $false,
   [String]$SftpHostname = "10.10.10.11",
   [String]$SftpUsername = "cdruser",
   [String]$SftpPassword = "cdrpassword",
   [String]$SftpFingerprint = "ssh-rsa 2048 11:11:22:11:22:11:22:33:11:22:33:11:22:33:22:11",
   [String[]]$SftpFiles = @("c:\ucm_reports\cdr_output.csv","c:\ucm_reports\cdr_summary.txt"),
   [String]$SftpDestination = "/home/cdruser/"
)

############ HELPERS ##################
## Import SSH Shell helper ###
Try{
   . .\Helpers\New-SSHShellSession.ps1 
}Catch{
   Write-Host $_.Exception.Message
   Write-Host $_.Exception.ItemName
   Write-Host "Please source the SSH Shell helper script"
   Exit 1
}

## Import SSH File Transfer helper ###
Try{
   . .\Helpers\New-SSHTransferSession.ps1 
}Catch{
   Write-Host $_.Exception.Message
   Write-Host $_.Exception.ItemName
   Write-Host "Please source the SSH File Transfer helper script"
   Exit 1
}

## Import Pivot Table helper ###
Try{
   . .\Helpers\New-PSPivotTable.ps1 
}Catch{
   Write-Host $_.Exception.Message
   Write-Host $_.Exception.ItemName
   Write-Host "Please source the Pivot Table helper script"
   Exit 1
}
#####################################

Add-Type -TypeDefinition @"
    public struct CDRData {
       public string TimeStamp;
       public string Date;
       public string CallingParty;
       public string CalledParty;
       public override string ToString() { return CallingParty; }
    }
"@

$ExtensionString = ""
$iExtensions = 0
ForEach($aExtension in $Extensions){
   If($iExtensions -lt $Extensions.Count - 1){
      $ExtensionString += "finalCalledPartyNumber like '$aExtension' OR "
   }Else{
      $ExtensionString += "finalCalledPartyNumber like '$aExtension'"
   }
   $iExtensions += 1
}

switch($WhichData){
   "LastMonth" {
      ## Are we starting a new year?? ##
      If((get-date).Month -eq 1){ 
         $yearToReportOn = (Get-Date).Year - 1
      }Else{
         $yearToReportOn = (Get-Date).Year
      }
      ####
   
      $lastMonth = (get-date).Month - 1
      $lastDayOfLastMonth = [System.DateTime]::DaysInMonth($(get-date).Year, $($(get-date).Month - 1))
      $unixTimeFirstOfLastMonth = [Math]::Floor([decimal](Get-Date(Get-Date "$lastMonth/1/$yearToReportOn 12:00 AM").ToUniversalTime()-uformat "%s"))
      $unixTimeLastOfLastMonth = [Math]::Floor([decimal](Get-Date(Get-Date "$lastMonth/$lastDayOfLastMonth/$yearToReportOn 11:59 PM").ToUniversalTime()-uformat "%s"))
      $SSHCommand = "run sql car select datetimeOrigination,callingPartyNumber,finalCalledPartyNumber FROM tbl_billing_data WHERE datetimeOrigination > $unixTimeFirstOfLastMonth AND datetimeOrigination < $unixTimeLastOfLastMonth AND (($ExtensionString) AND callingPartyNumber NOT LIKE 'NULL') ORDER BY datetimeOrigination"
      $EmailSubject += " for $lastMonth/$yearToReportOn"
   }
   default {
      $thisMonth = (Get-Date).Month
      $thisYear = (Get-Date).Year
      $SSHCommand = "run sql car select datetimeOrigination,callingPartyNumber,finalCalledPartyNumber FROM tbl_billing_data WHERE ($ExtensionString) AND callingPartyNumber NOT LIKE 'NULL' ORDER BY datetimeOrigination"
      $EmailSubject += " for $thisMonth/$thisYear"
   }
}

$Results = New-SSHShellSession -Hostname $SSHHostname -Username $SSHUsername -Password $SSHPassword -Command $SSHCommand

$CDRs = @()
$Result = $Results.Split("`n")
ForEach($aLine in $Result){
   If($aLine[0] -match "^[0-9]+$"){
      $tmp = $aLine -replace "\s+","," 
      $CDR = New-Object CDRData
      $origin = New-Object -Type DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, 0
      $convertedTime = $origin.AddSeconds($tmp.Split(",")[0])
      $zone = [system.timezoneinfo]::FindSystemTimeZoneById($MyTimeZone)
      $convertedTime = [system.timezoneinfo]::ConvertTimeFromUtc($convertedTime,$zone)
      $aDate = [String]$convertedTime
      $CDR.Date = $aDate.Split(" ")[0]
      $CDR.TimeStamp = $convertedTime 
      $CDR.CallingParty = $tmp.Split(",")[1]
      $CDR.CalledParty = $tmp.Split(",")[2]
      $MaskExtension.GetEnumerator() | % { 
         If($($_.key) -eq $CDR.CalledParty){ $CDR.CalledParty = $($_.value) }
      }
      $CDRs += $CDR
   }
}

If($SendEmailAlert) {
   $EmailBody = "Attached is the $EmailSubject"
   $CDRs | Select TimeStamp,CallingParty,CalledParty | Export-CSV $OutputResultsFile -NoTypeInformation
   $pivot = New-PSPivotTable $CDRs -yProperty Date -xlabel CalledParty -Count
   $pivot | ft -auto | Out-String | Out-File $OutputResultsFileSummary
   Send-MailMessage -To $EmailTo -Subject "$EmailSubject" -Body "$EmailBody" -SmtpServer $EmailServer -From $EmailFrom -Attachments $OutputResultsFile, $OutputResultsFileSummary
}

If($SftpTransferFiles){
   ForEach($aFile in $SftpFiles){
      New-SSHTransferSession -Username $SftpUsername -Password $SftpPassword -Hostname $SftpHostname -Fingerprint $SftpFingerprint -SourceFile $aFile -Destination $SftpDestination
   }
}