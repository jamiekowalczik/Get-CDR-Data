Function New-SSHTransferSession {
   <#
      comment based help omitted here
   #>

   [cmdletbinding(DefaultParameterSetName="Property")]

   Param(
      [String]$PreReqModuleFile = ".\Modules\winscp556automation\WinSCPnet.dll",
      [String]$Username,
      [String]$Password,
      [String]$Hostname,
      [String]$Fingerprint,
      [String]$SourceFile,
      [String]$Destination
   )
   
   Try{
      Add-Type -Path $PreReqModuleFile
   }Catch{
      Write-Host $_.Exception.Message
      Write-Host $_.Exception.ItemName
      Write-Host "The module can be downloaded from: http://winscp.net/eng/docs/library_install"
      Exit 1
   }
   
   $sessionOptions = New-Object WinSCP.SessionOptions
   $sessionOptions.Protocol = [WinSCP.Protocol]::Sftp
   $sessionOptions.HostName = $Hostname
   $sessionOptions.UserName = $Username
   $sessionOptions.Password = $Password
   $sessionOptions.SshHostKeyFingerprint = $Fingerprint

   $session = New-Object WinSCP.Session
   $session.Open($sessionOptions)

   $transferOptions = New-Object WinSCP.TransferOptions
   $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
 
   $transferResult = $session.PutFiles($SourceFile, $Destination, $False, $transferOptions)
 
   $transferResult.Check()
 
   # Print results
   #foreach ($transfer in $transferResult.Transfers){
   #   Write-Host ("Upload of {0} succeeded" -f $transfer.FileName)
   #}

   $session.Dispose()
}