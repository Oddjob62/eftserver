function new-sshkeys {
    Param(
    [Parameter(Mandatory=$true)]
    [string]$keyfilepath = "C:\Users\Administrator\Desktop\thisCompany"
    )

$supressPassphrase = @'
 -N '""' -q
'@

    $sshkeygenCommand = "ssh-keygen -t rsa -b 2048 -f "
#   ssh-keygen -t rsa -b 2048 -f "C:\Users\Administrator\Desktop\newkeys" -N '""'
    $fullcommand = $sshkeygenCommand + '"' + $keyFilePath + '"' + $supressPassphrase
    Invoke-Expression -Command $fullcommand
}


$VerbosePreference="Continue"
$SSHKeyDir = "C:\Users\Administrator\Desktop"

#Server Details
$serverHost = "localhost"
$serverPort = "1100"
$adminUserName = "administrator"
$adminPassword = "************"

#New SFTP User Details
$companyname = "thatCompany"
$whiteListIP = "192.168.1.98","192.168.1.199","192.168.1.200","192.168.1.201"
$companyDescription = ""
$fullName = $companyname
$authenticationMethod = "key"      # inherited, password, key, both, either (v6.3 and later)


Write-Verbose "Creating SSH Keys"
$keyFilePath = Join-path -Path $SSHKeyDir -ChildPath $companyName
if(Test-path -Path $keyFilePath){
    write-host "Key files already exist at $keyFilePath, please delete them or press return to use these keys" -ForegroundColor white
    $confirmation = Read-Host "[Return]"
}

if(!(Test-path -Path $keyFilePath)){
new-sshkeys -keyfilepath $keyFilePath
Write-Verbose "Creating SSH Keys and saving to $keyFilePath"
}else{
    write-host "Using existing Key files"
}

write-verbose "Connect to server"
$EFTServer = New-Object -ComObject "SFTPCOMInterface.CIServer"
$EFTServer.connect($ServerHost, $serverPort, $adminUserName,$adminPassword);

write-verbose "Get Site"
$mySite =  New-Object -ComObject "SFTPCOMInterface.CISite"
$mySite = ($EFTServer.Sites()).item(0)
#$mySite.GetIPAccessRules() 
#$mysite.IPAccessAllowedDefault


$pubkeyFilePath = $keyFilePath + ".pub"
$pubkey = [System.IO.File]::ReadAllBytes($pubkeyFilePath)
Write-Verbose "Importing Public key to $companyname account"
if($mySite.sshkeys | ?{$_.Name -eq $companyName}){
    write-host "public Key already exist for $companyName, please delete them or press return to use these keys" -ForegroundColor white
    $confirmation = Read-Host "[Return]"
}

if($mySite.sshkeys | ?{$_.Name -eq $companyName}){
    Write-Host "Using Existing public key"
}else{
    try{
        $sshKeyID = $mySite.ImportSSHPublicKey($companyName, $pubkey)
    }catch [System.Runtime.InteropServices.COMException]{
        write-error "Failed to import public key."
    }
}

Write-Verbose "Whitelisting IPs"
$whiteListIP | foreach {
    $thisIP = $_
    try{
        $mysite.AddIPAccessRule($thisIP, $true, 0, $companyname)
        Write-Verbose "Whitelisted $thisIP"
    }catch [System.Runtime.InteropServices.COMException]{
        Write-host "The specified IP $thisIP already exists or has an invalid format" -ForegroundColor Red
        $rule = $mySite.GetIPAccessRules() | ?{$_.Address -eq $thisIP}
        Write-host "Rule added on $($rule.Added). Reason: $($rule.Reason)" -ForegroundColor Red
    }catch{
        Write-host $Error[0].Exception.Message -ForegroundColor Red
    }
}

Write-Verbose "Creating Password according to policy"
$myPassword = $mysite.CreateComplexPassword() # Create according to policy
write-host "New password is: $myPassword"

write-verbose "Creating new account for $companyName"
$mySite.CreateUserEx($companyname,$myPassword,0, $companyDescription, $fullName, $true, $true, "Default settings", -2)
write-verbose "Setting Authenticaiton for new account for $companyName to $authenticationMethod"
$mySite.GetUserSettings($companyname).SetSFTPAuthenticationType($authenticationMethod)

$mySite.GetUserSettings($companyname).setSSHKeyID($sshKeyID)








