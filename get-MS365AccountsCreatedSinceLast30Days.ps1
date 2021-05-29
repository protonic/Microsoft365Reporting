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
function Get-Office365IdentityReport {
# Several reports
$a=Get-MsolUser -All | where-object {$_.whenCreated -gt (Get-Date).AddDays(-30)} | select UserPrincipalName,isLicensed,WhenCreated | ConvertTo-Html 
#$b=Get-MsolUser -All | Where {$_.IsLicensed -eq $True -AND $_.BlockCredential -eq $True} | Select UserPrincipalName,isLicensed,WhenCreated,BlockCredential | ConvertTo-Html 

# Creating the HTML Page 
$Html = @"
<!DOCTYPE html>
<head>
 <title>MS 365 License & Unlicensed users created last 30 days</title>
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
Dear Team,<br>
<br>
This report will help:<br>
To identify MS 365 Licensed & Unlicensed users created last <b>30 days</b><br>
To identify MS 365 Overall Licensed but SignIn Blocked Users, so license can be <b>raclaimed</b> if missed during termination.<br>

<br>
<centre>
MS 365 Licensed & Unlicensed users created last 30 days
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
Send-Mailmessage  -from "automation@techwanderers.in" -To "automation@techwanderers.in" -Subject "Automation: MS 365 Licensed & Unlicensed users created last 30 days - $date" -body $html -BodyAsHtml -Credential $cred -SmtpServer smtp.office365.com -Port 587 -UseSsl
}
# Running the function
Get-Office365IdentityReport

exit

