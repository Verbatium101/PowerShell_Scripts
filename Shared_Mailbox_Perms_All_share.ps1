<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.155
	 Created on:   	02/18/2018 2:50 PM
	 Created by:   	Darren Jochims
	 Filename:     Shared_Mailbox_Perms_All	
	===========================================================================
	.DESCRIPTION
		Produces a CSV with each permission for every shared mailbox.
#>

# Credentials and Connection
$adminName = #"Admin Account username"
$secpasswd = Get-Content # Location of secure password (User Create_secure.ps1) | ConvertTo-SecureString
$Credential = New-Object System.Management.Automation.PSCredential ($adminName, $secpasswd)
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/?proxyMethod=RPS -Authentication Basic -Credential $Credential -AllowRedirection
Import-PSSession $Session

# Create PSdrive Connection ** Location to save reports **
New-PSDrive -Name "SaveLoc" -PSProvider FileSystem -Root "\\netshare\location" 

# Set universal variables
$date = (get-date).ToString("M-d-yyyy")

# Declare arrays
$PermResult = @()
$Count = 0

# Get all shared mailboxes
$Mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited

# Cycle through data for ending array
foreach ($Identity in $Mailboxes) {
	$Alias = $identity.Alias
	
	$Perms = Get-MailboxPermission -Identity $Alias
	
	foreach ($entry in $Perms) {
		try {
            #Add exclusions as they are needed for our environment
			if ($entry.user -like '*NT AUTHORITY*' -or $entry.user -like '*NAMPR20A002*' -or $entry.user -like '*NAMPRD20*' -or $entry.user -like '*S-1-5*' -or $entry.user -like '*JitUsers*') {
				throw
			}
			else {
				$upn = $entry.user
				$userinfo = Get-User $upn | Select-Object FirstName, LastName, WindowsEmailAddress
				$PermResult += New-Object PsObject -Property @{
					"Mailbox"	    = "$($Identity.Identity)";
					"MailboxEmailAddress"  = "$($Identity.WindowsLiveID)";
					"DelegatedUserUpn"    = "$($entry.User)";
					"DelegatedUserFirstName"    = "$($userinfo.FirstName)";
					"DelegatedUserLastName"    = "$($userinfo.LastName)";
					"DelegatedUserEmail"    = "$($userinfo.WindowsEmailAddress)";
					"AccessRights"  = "$($entry.AccessRights)";
				}
			}
		}
		Catch {
			$Count++
		}
	}
}

# Write results
$PermResult | Select-Object Mailbox, MailboxEmailAddress, DelegatedUserUpn, DelegatedUserFirstName, DelegatedUserLastName, DelegatedUserEmail, AccessRights | Export-Csv "SaveLoc:\Shared_Mailbox_delegations.csv" -NoTypeInformation

# Remove Session and PSdrive
Get-PSSession | Remove-PSSession
Get-PSDrive SaveLoc | Remove-PSDrive
