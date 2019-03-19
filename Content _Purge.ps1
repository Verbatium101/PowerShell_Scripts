

# Connection to Security and Compliance
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force
$UserCredential = Get-Credential
$Proxy = New-PSSessionOption -ProxyAccessType IEConfig
$ComplianceSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $ComplianceSession


# Set Initial Variables
$SearchName = Read-Host -prompt "Name of Search"
$Purge = Read-Host -prompt "Type of purge (H)ard (S)oft"


# Set Secondary Variables based on Initial
If ($Purge -eq "S") {
	Write-Host "Content will be recoverable for 30 days" -ForegroundColor Cyan
    $PurgeType = Softdelete
	$ActionName = "$SearchName_Purge_Soft"
	
} ElseIf ($Purge -eq "H") {
    Write-Host "Content will not be recoverable" -ForegroundColor Red
	$ActionName = "$SearchName_Purge_Hard"
} Else {
Write-Host "Invalid type. Exiting" -ForegroundColor Red
Exit
}

Start-Sleep -s 5

# Process Compliance Action
If ($Purge -eq "S") {
	Write-Host "Searching and Removing Content" -ForegroundColor Green
    New-ComplianceSearchAction -SearchName $SearchName -PurgeType $PurgeType -Confirm -Actionname $ActionName
	
}
Elseif ($Purge -eq "H") {
    Write-Host "Searching and Removing Content" -ForegroundColor Green
	New-ComplianceSearchAction -SearchName $SearchName -Confirm -Actionname $ActionName
}

Write-Host "Retreiving Results" -ForegroundColor Green

Start-Sleep -s 5

Get-ComplianceSearchAction -Identity $ActionName |FL Searchname, Results, Errors

Write-Host "For More detailed results use Command: Get-ComplianceSearchAction -Identity $ActionName |FL " -ForegroundColor Magenta

#Clean up Session and Variables

Remove-PSSession $ComplianceSession


Write-Host"Compiance Action and Cleanup Completed" -ForegroundColor Green





