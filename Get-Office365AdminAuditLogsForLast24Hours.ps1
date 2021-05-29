
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


################# Import Module ExchangeOnlineManagement ###################
Import-Module ExchangeOnlineManagement


#Provide O365 admin creds
#create crdential here this in only one time command # it once xml is generated.
#Get-Credential "automation@techwanderers.in" | Export-Clixml .\creds.xml

#Authenticate with XML file
$cred = Import-Clixml "C:\Automation\credentials\automation.xml"
#create Session
Connect-IPPSSession -Credential $cred

# Creating a function that generates the reports
function Get-Office365AuditReport {
# Several reports
$body=Search-AdminAuditLog -Cmdlets Set-TransportRule -StartDate (Get-Date).AddHours(-24)  -IsSuccess $true | select ObjectModified,CmdletName,CmdletParameters,Caller,ExternalAccess,Succeeded,Error,ClientIP,RunDate,IsValid,ObjectState | ConvertTo-Html
#$body2=Search-AdminAuditLog -Cmdlets Set-Mailbox -StartDate (Get-Date).AddHours(-24)  -IsSuccess $true | select ObjectModified,CmdletName,Caller,CmdletParameters,Error,Succeeded,RunDate,ClientIP  | ConvertTo-Html

# Creating the HTML Page 
$Html = @"
<!DOCTYPE html>
<head>
 <title>Office 365 admin Audit Logs for last 24 Hours</title>
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
To identify Office 365 admin Audit Logs for last  <b>24 Hours</b><br>


<br>
<centre>
Office 365 admin Audit Logs for last 24 hours
</centre>
<br>
$body
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

################ Change your SMTP server and sender receiver with subjectline ###############
Send-Mailmessage -from "automation@techwanderers.in" -To "automation@techwanderers.in" -Subject "Automation: Office 365 admin Audit Logs for last 24 hours - $date" -body $html -BodyAsHtml -Credential $cred -SmtpServer smtp.office365.com -Port 587 -UseSsl
}
# Running the function
Get-Office365AuditReport

Disconnect-ExchangeOnline -Confirm:$false

exit


