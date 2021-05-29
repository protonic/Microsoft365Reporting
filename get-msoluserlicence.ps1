<#PSScriptInfo 
 
.VERSION 1.2
 
.AUTHOR TechnologyWanderers Youtube Channel
 
.COMPANYNAME 
 
.COPYRIGHT 
 
.TAGS O365 Reprt 
 
.LICENSEURI 
 
.PROJECTURI 
 
.ICONURI 
 
.EXTERNALMODULEDEPENDENCIES 
 
.REQUIREDSCRIPTS 
 
.EXTERNALSCRIPTDEPENDENCIES 
 
.RELEASENOTES https://www.youtube.com/c/technologywanderers
 
 #>

# The Output will be written to this file in the current working directory
$LogFile = "Office_365_Licenses.csv"

##Provide O365 admin creds
#create crdential here this in only one time command # it once xml is generated.
#Get-Credential "automation@techwanderers.in" | Export-Clixml .\creds.xml

#Authenticate with XML file
$cred = Import-Clixml "C:\Automation\credentials\automation.xml"
#create Session
connect-msolservice -Credential $cred

$css = @"
<style>
h1, h5, th { text-align: center; font-family: Segoe UI; }
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }
td { font-size: 11px; padding: 5px 20px; color: #000; }
tr { background: #b8d1f3; }
tr:nth-child(even) { background: #dae5f4; }
tr:nth-child(odd) { background: #b8d1f3; }
</style>
"@

$date = ( get-date ).ToString('yyyyMMdd')
# Pathes to export reports
$HtmlReportFilePath = Join-Path -Path $PSScriptRoot -ChildPath Office365LicensedUsers.html


write-host "Connecting to Office 365..."

# Get a list of all licences that exist within the tenant
$licensetype = Get-MsolAccountSku | Where {$_.ConsumedUnits -ge 1}

# Loop through all licence types found in the tenant
foreach ($license in $licensetype) 
{	
	# Build and write the Header for the CSV file
	$headerstring = "UserPrincipalName,LicenseAssigned"
	
	foreach ($row in $($license.ServiceStatus)) 
	{
		$headerstring = ($headerstring + "," + $row.ServicePlan.servicename)
	}
	
	Out-File -FilePath $LogFile -InputObject $headerstring -Encoding UTF8 -append
	
	write-host ("Gathering users with the following subscription: " + $license.accountskuid)

	# Gather users for this particular AccountSku
	$users = Get-MsolUser -all | where {$_.isLicensed -eq "True" -and $_.licenses.accountskuid -contains $license.accountskuid}

	# Loop through all users and write them to the CSV file
	foreach ($user in $users) {
		
		write-host ("Processing " + $user.displayname)

        $thislicense = $user.licenses | Where-Object {$_.accountskuid -eq $license.accountskuid}

		$datastring = ($user.userprincipalname + "," + $license.SkuPartNumber)
		
		foreach ($row in $($thislicense.servicestatus)) {
			
			# Build data string
			$datastring = ($datastring + "," + $($row.provisioningstatus))
		}
		
		Out-File -FilePath $LogFile -InputObject $datastring -Encoding UTF8 -append
	}

    Out-File -FilePath $LogFile -InputObject " " -Encoding UTF8 -append
}			

write-host ("Script Completed.  Results available in " + $LogFile)

###### tying to make it as HTML ######

Import-CSV ".\Office_365_Licenses.csv" | ConvertTo-Html -Head $css -Body "<h1>Email Report for Licensed users</h1>`n<h5>Generated on $(Get-Date)</h5>" | Out-File ".\Office365LicensedUsers.html" -Encoding utf8 -Force



######################## Make sure changing the properties of From and To ###########################
Send-MailMessage -from "automation@techwanderers.in" -To "automation@techwanderers.in"  -Subject "Automation: Office 365 Licensed users from Techwanderers" -Body "<h1>Email Report for Licensed users is attached</h1>`n<h5>Generated on $(Get-Date)</h5></p></p></p></p></p></p></p></p></p>Best Regards,<br>Office 365 Automation Support<br>URL: https://www.youtube.com/c/technologywanderers<br>automation@yourdomain.in<br>India<br>" -BodyAsHtml -Priority High -DeliveryNotificationOption OnSuccess, OnFailure  -Credential $cred -SmtpServer smtp.office365.com -Port 587 -UseSsl -Attachments .\Office365LicensedUsers.html


Remove-Item -Path .\Office_365_Licenses.csv -Force
Remove-Item -Path .\Office365LicensedUsers.html -Force

exit
