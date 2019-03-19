#####################################################################################################################################################
# Export Office 365 Mail logs
#
# Description :Batch Export of Office 365 Mail logs to a CSV file.
# Office 365 only keeps 30 days of logs 
# Office 365 when running a get-messagetrace will only output 5000 max
# 
# This script outputs 5000 entries at a time into a csv file, when all done, it will combine all CSV files into 1 main one.
#
# To run the script:
#
# Connect to office 365 via powershell
# Run the command, no inputs required
# Author: 				Darren Jochims
# Last Modified Date: 	3/19/2019
# Last Modified By: 	Darren Jochims
#####################################################################################################################################################


# User Credentials
$adminName = #"Admin Username"
$secpasswd = Get-Content #"Location of secure password (user create_secure.ps1)" | ConvertTo-SecureString
$seccredential = New-Object System.Management.Automation.PSCredential ($adminName, $secpasswd)

# Mount Drives Location for report
New-PSDrive -Name ReportLoc -Root "D:\Test" -PSProvider FileSystem

#connect to Exchange Online
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $seccredential -Authentication Basic -AllowRedirection
Import-PSSession $Session


	#Remove all existing Powershell sessions
	#Get-PSSession | Remove-PSSession
	
	
	$Batchfile = $null  
	$Page = 1  
	do  
	{  
		Write-Host "Processing - Page $Page..."  
		
		# by default it will just get the 6th day back, to get more change the values below up to -30
		$Batchfile = Get-MessageTrace -StartDate ((Get-Date).AddDays(-6).ToString('MM-dd-yyyy')) -EndDate ((Get-Date).AddDays(-5).ToString('MM-dd-yyyy')) -PageSize 5000 -Page $Page | Select Received,*Address,Subject,Status,Size
		
		# Change to only do previous day
		#$Batchfile = Get-MessageTrace -StartDate (Get-Date).AddDays(-1).ToString('MM-dd-yyyy') -EndDate (Get-Date).ToString('MM-dd-yyyy') -Status Delivered -PageSize 5000 -Page $Page| Select Received,*Address,Subject
		
		$Batchfile | Export-Csv "ReportLoc:\Temp-$PAGE.csv" -NoTypeInformation
		
		$Page++  
	}  
	until ($Batchfile -eq $null)  

	# Naming of the actual file name to dump out.
	
	# if you change the number of days in line 44, you will also have to
	$Pfilename = "O365 Mail Logs - {0:MM-dd-yyyy}" -f (Get-Date).AddDays(-6) 
	$Pfilename2 = " to {0:MM-dd-yyyy}" -f (Get-Date).AddDays(-5)
	$filename = $Pfilename + $Pfilename2


#get all temp CSV's and dump to 1 CSV file.
Get-ChildItem "ReportLoc:\Temp-*.csv" |   ForEach-Object {Import-Csv $_} | Export-Csv -NoTypeInformation "ReportLoc:\$filename.csv"

#remove all the temp CSVs (brutal , any file called "Temp-*.csv" gets purged!)
Remove-item "ReportLoc:\Temp-*.csv" –recurse

#Copy file to network location.
New-PSDrive -Name "Archive" -PSProvider "FileSystem" -Root #"Archive Location"

Copy-Item "ReportLoc:\$filename.csv" -Destination "Archive:\$filename.csv"


#Clean up session
Remove-PSDrive -Name "ReportLoc"
Remove-PSDrive -Name "Archive"
Get-PSSession | Remove-PSSession



