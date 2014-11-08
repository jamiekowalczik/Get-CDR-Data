Function New-SSHShellSession {
   <#
      comment based help omitted here
   #>

   [cmdletbinding(DefaultParameterSetName="Property")]

   Param(
      [String]$Hostname,
      [String]$Username,
      [String]$Password,
      [String]$Command,
      [String]$Expect = "admin:",
      [String]$PreReqModuleFile = ".\Modules\SSH-Sessions\Renci.SshNet.dll"
   )
   
   Try{
      [void][reflection.assembly]::LoadFrom( (Resolve-Path $PreReqModuleFile) ) 
   }Catch{
      Write-Host $_.Exception.Message
      Write-Host $_.Exception.ItemName
      Write-Host "The module can be downloaded from: http://www.powershelladmin.com/w/images/a/a5/SSH-SessionsPSv3.zip"
      Write-Host "Make sure to unblock the SSH-SessionsPSv3.zip and then uncompress it's contents to a valid Module Path or specify the module file with the PreReqModuleFile command line option."
      Exit 1
   }
   
   Try{
      $SshClient = New-Object Renci.SshNet.SshClient("$Hostname", 22, "$Username", "$Password")
      $SshClient.Connect()
   }Catch{
      Write-Host $_.Exception.Message
      Write-Host $_.Exception.ItemName
      Exit 1
   }

   $SshStream = $SshClient.CreateShellStream("dumb", 80, 24, 800, 600, 1024)
   $reader = new-object System.IO.StreamReader($sshStream) 
   $writer = new-object System.IO.StreamWriter($sshStream) 
   $writer.AutoFlush = $true

   $sshStream.Expect($Expect) | Out-Null
   Start-Sleep -s 3  
   $sshStream.Write("$Command" + "`n") | Out-Null
   Start-Sleep -s 3  
   $writer.WriteLine("exit" + "`n") | Out-Null

   $SshClient.Disconnect() | out-null
   
   $Results = $reader.ReadtoEnd() 
   Return $Results
}