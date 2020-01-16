$MyDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$EncryptedPasswordFile = "$mydir\email@domin.net.securestring"
Read-Host "Type the password for email@domin" -AsSecureString | ConvertFrom-SecureString | Out-File -FilePath $EncryptedPasswordFile
