<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.155
	 Created on:   	3/19/2019 1:12 PM
	 Created by:   	Darren Jochims
	 Organization: 	
	 Filename:     	PBI_Data_Catalog_Pull_Auto.ps1
	===========================================================================
	.DESCRIPTION
		Uses Power BI REST API to get data on all groups, reports, datasets, 
		datasources, Pro Users so a data catalog can correlate the resultant 
		data in a logical manner.
	===========================================================================
	.REQUIREMENTS
		Requires 64bit powershell. If useing 32bit you will need to change the 
		paths for $adal and $adalforms.

		Microsoft PowerBI Management:
        Run - Install-Module MicrosoftPowerBIMgmt

		Microsoft ADAL:
		Run - Install-Module Microsoft.ADAL.PowerShell

        
#>

# Functions
function Renew-Token {
	$TimeString = $Stopwatch.Elapsed.ToString('mm')
	[int]$TimeInt = [convert]::ToInt32($TimeString, 10)
	if ($TimeInt -ge "50") {
		
		# Auth to API to get Token
		$auth2 = Invoke-RestMethod -Uri $pbiAuthorityUrl -Body $authBody -Method POST -Verbose
		
		# Building Rest API header with authorization token
		$authHeader = @{
			'Content-Type'  = 'application/json'
			'Authorization' = 'Bearer ' + $auth.access_token
		}
		# Restart Stopwatch
		$Stopwatch.Restart()
	}
	else {
	}
}

# Load Active Directory Authentication Library (ADAL) Assemblies
$adal = "${env:ProgramFiles}\WindowsPowerShell\Modules\Microsoft.ADAL.PowerShell\1.12\Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
$adalforms = "${env:ProgramFiles}\WindowsPowerShell\Modules\Microsoft.ADAL.PowerShell\1.12\Microsoft.IdentityModel.Clients.ActiveDirectory.WindowsForms.dll"
[System.Reflection.Assembly]::LoadFrom($adal)
[System.Reflection.Assembly]::LoadFrom($adalforms)

# Load MSOL and PBI Service modules
Import-Module MSOnline
Import-Module MicrosoftPowerBIMgmt

# Tenant and API connection information
$tenantID = #"Enter Tenant ID"
$pbiAuthorityUrl = "https://login.windows.net/$tenantID/oauth2/token"
$pbiResourceUrl = "https://analysis.windows.net/powerbi/api"

# Mount Drives ** Location to save reports **
New-PSDrive -Name ReportLoc -Root "\\netshare\location" -PSProvider FileSystem

# Set Variables
$date = (Get-Date).ToString('MM/dd/yyyy')
$authHeader = @()

# Client ID and Client Secret 
$clientId = #"Enter Client ID for app registration"
$client_secret = #"Enter Client Secret"

# User Credentials
$adminName = #"Admin Account username"
$unsecpasswd = #"Clear text Password"
$secpasswd = Get-Content # Location of secure password (User Create_secure.ps1) | ConvertTo-SecureString
$seccredential = New-Object System.Management.Automation.PSCredential ($adminName, $secpasswd)

# Connect MsolService and PBI Service
Connect-MsolService -Credential $seccredential
Connect-PowerBIServiceAccount -Credential $seccredential

# Authenticate to Azure/PBI
$authBody = @{
	'resource'	    = $pbiResourceUrl
	'client_id'	    = $clientId
	'grant_type'    = "password"
	'username'	    = $adminName
	'password'	    = $unsecpasswd
	'scope'		    = "openid"
	'client_secret' = $client_secret
}

# Auth to API to get Token
$auth = Invoke-RestMethod -Uri $pbiAuthorityUrl -Body $authBody -Method POST -Verbose
$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Building Rest API header with authorization token
$authHeader = @{
	'Content-Type'  = 'application/json'
	'Authorization' = 'Bearer ' + $auth.access_token
}

# Get all Groups
$groups = Get-PowerBIWorkspace -Scope Organization -All

# Declare final Arrays
$PBIGroups = @()
$PBIReports = @()
$PBIDatasources = @()
$PBIUsers = @()

# Create Pro User List
$prousers = Get-MsolUser -All | Where-Object {
	($_.licenses).AccountSkuId -match "POWER_BI_PRO"
} | Select-Object DisplayName, UserPrincipalName

foreach ($pro in $prousers) {
	$PBIUsers += New-Object PsObject -Property @{
		"Date" = "$date";
		"Name" = "$($pro.DisplayName)";
		"UPN"  = "$($pro.UserPrincipalName)";
	}
}


# Loop through Group, dataset, report, and datasources; add to arrays
foreach ($group in $groups) {
	Renew-Token
	$groupID = $group.Id
	$uri = "https://api.powerbi.com/v1.0/myorg/admin/groups/$groupID/datasets"
	$uri2 = "https://api.powerbi.com/v1.0/myorg/admin/groups/$groupID/reports"
	$datasets = Invoke-RestMethod -Uri $uri -Headers $authHeader -Method GET -Verbose -ErrorAction SilentlyContinue
	$reports = Invoke-RestMethod -Uri $uri2 -Headers $authHeader -Method GET -Verbose -ErrorAction SilentlyContinue
	
	$PBIGroups += New-Object PsObject -Property @{
		"Date"			    = "$date";
		"GroupName"		    = "$($group.Name)";
		"GroupID"		    = "$($group.id)";
		"ReadOnly"		    = "$($group.isReadOnly)";
		"DedicatedCapacity" = "$($group.isOnDedicatedCapacity)";
		"CapacityId"	    = "$($group.CapacityId)";
		"Description"	    = "$($group.description)";
		"GroupType"		    = "$($group.type)";
		"State"			    = "$($group.state)";
	}
	
	foreach ($report in $reports.value) {
		Renew-Token
		$PBIReports += New-Object PsObject -Property @{
			"Date"	     = "$date";
			"GroupName"  = "$($group.name)";
			"GroupID"    = "$($group.id)";
			"ReportName" = "$($report.name)";
			"ReportID"   = "$($report.id)";
			"DatasetID"  = "$($report.datasetId)";
			"WebURL"	 = "$($report.webUrl)";
			"EmbedURL"   = "$($report.embedUrl)";
		}
	}
	
	
	foreach ($dataset in $datasets.value) {
		Renew-Token
		$datasetID 		= $dataset.id
		$uri3 			= "https://api.powerbi.com/v1.0/myorg/groups/$groupID/datasets/$datasetID/refreshes"
		$uri4 			= "https://api.powerbi.com/v1.0/myorg/admin/datasets/$datasetID/datasources"
		$refreshes		= Invoke-RestMethod -Uri $uri3 -Headers $authHeader -Method GET -Verbose -ErrorAction SilentlyContinue
		$datasources 	= Invoke-RestMethod -Uri $uri4 -Headers $authHeader -Method GET -Verbose -ErrorAction SilentlyContinue
		
		foreach ($datasource in $datasources.value) {
			Renew-Token
			$PBIDatasources += New-Object PsObject -Property @{
				"Date"			    = "$date";
				"GroupName"	       	= "$($group.name)";
				"GroupID"		   	= "$($group.id)";
				"DatasetName"	   	= "$($dataset.name)";
				"DatasetID"	       	= "$($dataset.id)";
				"DatasetOwner"	   	= "$($dataset.configuredBy)";
				"DataSourceName"   	= "$($datasource.name)";
				"ConnectionString" 	= "$($datasource.connectionString)";
				"DatasourceType"   	= "$($datasource.datasourceType)";
				"GatewayID"	       	= "$($datasource.gatewayId)";
				"Server"		   	= "$($datasource.connectionDetails.server)";
				"Database"		   	= "$($datasource.connectionDetails.database)";
				"LastRefreshType"   = "$($refreshes.value.refreshType[0])";
				"LastRefreshStart"  = "$($refreshes.value.startTime[0])";
				"LastRefreshEnd"    = "$($refreshes.value.endTime[0])";
				"LastRefreshStatus" = "$($refreshes.value.status[0])";
			}
		}
		$datasources = $null
	}
	$reports 	= $null
	$datasets 	= $null
}

# Print Arrays
$PBIDatasources | Select-Object Date, GroupName, GroupID, DatasetName, DatasetID, DatasetOwner, DataSourceName, ConnectionString, DatasourceType, GatewayID, Server, Database, LastRefreshType, LastRefreshStart, LastRefreshEnd, LastRefreshStatus | Export-Csv -Path "ReportLoc:\datasources.csv" -NoTypeInformation -Append
$PBIReports | Select-Object Date, GroupName, GroupID, ReportName, ReportID, DatasetID, WebURL, EmbedURL | Export-Csv -Path "ReportLoc:\Reports.csv" -NoTypeInformation -Append
$PBIGroups | Select-Object Date, GroupName, GroupID, ReadOnly, DedicatedCapacity, Description, GroupType, State, Users | Export-Csv -Path "ReportLoc:\Groups.csv" -NoTypeInformation -Append
$PBIUsers | Select-Object Date, Name, UPN | Export-Csv -Path "ReportLoc:\PBI_Pro_Users.csv" -NoTypeInformation -Append

# Remove PS Drive
Remove-PSDrive -Name ReportLoc

