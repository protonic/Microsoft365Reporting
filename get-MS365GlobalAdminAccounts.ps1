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

#requires -Version 3.0 -Modules MSOnline



#Provide O365 admin creds
#create crdential here this in only one time command # it once xml is generated.
#Get-Credential "automation@techwanderers.in" | Export-Clixml .\creds.xml

#Authenticate with XML file
$cred = Import-Clixml "C:\Automation\credentials\automation.xml"

#create Session
connect-msolservice -Credential $cred



# Creating a function that generates the reports



function Get-Office365AdminReport {
# Several reports
$a=Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Company Administrator").ObjectId | Select-Object -Property RoleMemberType,DisplayName,EmailAddress,isLicensed,LastDirSyncTime | ConvertTo-Html 

Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Company Administrator").ObjectId | Select-Object -Property RoleMemberType,DisplayName,EmailAddress,isLicensed | export-csv ".\MS365AdminAccountSummary.csv"


# Creating the HTML Page 
$Html = @"
<!DOCTYPE html>
<head>
 <title>MS 365 License & Unlicensed users created last 7 days</title>
  <style>
   body {
   font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
   }
   table, th, td {
    white-space: nowrap;
    border: 1px solid black;
    text-align: center;
    border-collapse: collapse;
    padding: 4px;
   }
   th {
    background-color: midnightblue;
    color: white;
   }
  </style>
</head>
<body>
<br>
<centre>
<br>
<br>
<br>
<br>
MS 365 Global Admin Roles Report
</centre>
<br>
<centre>
<table class="Techwanderers">
  <tr>
    <td class="Techwanderers-1" colspan="26">$a</td>
  </tr>
</table>
</centre>
<br>
<br>
<br>
<br>
<br>
<br>
Best Regards,<br>
Office 365 Automation Support<br>
URL: https://www.youtube.com/c/technologywanderers<br>
automation@yourdomain.in<br>
India<br>

"@
$date=(get-date).ToShortDateString()



##########################change your email address and other values here##########################
Send-Mailmessage  -from "automation@techwanderers.in" -To "automation@techwanderers.in" -Credential $cred -SmtpServer smtp.office365.com -Port 587 -Subject "Automation: MS 365 Global Admin Roles Report  - $date" -body $html -BodyAsHtml -Attachment .\MS365AdminAccountSummary.csv -UseSsl
}
# Running the function
Get-Office365AdminReport


Remove-Item -Path .\MS365AdminAccountSummary.csv -Force

exit
