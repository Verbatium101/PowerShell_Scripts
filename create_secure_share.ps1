


$credential = Get-Credential

#Location to store password
$credential.Password | ConvertFrom-SecureString | Set-Content -Path C:\secure\password.txt