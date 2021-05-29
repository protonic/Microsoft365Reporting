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

<# 
 
.DESCRIPTION 
 The script gets all licensed users accounts and creates reports of licenses usage by Country. 
 
#> 


#region Parameters

#Provide O365 admin creds
#create crdential here this in only one time command # it once xml is generated.
#Get-Credential "automation@techwanderers.in" | Export-Clixml C:\Automation\credentials\automation.xml

#Authenticate with XML file
$cred = Import-Clixml "C:\Automation\credentials\automation.xml"

#Connect
connect-msolservice -Credential $cred

$AccountSkuIdDecodeData = @{
    'AAD_BASIC'                  = 'Azure Active Directory Basic'
    'AX7_USER_TRIAL'             = 'Microsoft Dynamics AX7 User Trial'
    'CRM_HYBRIDCONNECTOR'        = 'CRM_HYBRIDCONNECTOR'
    'DESKLESS'                   = 'Microsoft StaffHub'
    'DESKLESSPACK'               = 'Office 365 F1'
    'DYN365_ENTERPRISE_P1_IW'    = 'Dynamics 365 P1 Trial for Information Workers'
    'DYN365_ENTERPRISE_PLAN1'    = 'Dynamics 365 Customer Engagement Plan Enterprise Edition'
    'ENTERPRISEPACK'             = 'Office 365 Enterprise E3'
    'ENTERPRISEPREMIUM'          = 'Office 365 Enterprise E5'
    'ENTERPRISEPREMIUM_NOPSTNCONF' = 'Office 365 Enterprise E5 without PSTN Conferencing'
    'ENTERPRISEWITHSCAL'         = 'Office 365 Enterprise E4'
    'EOP_ENTERPRISE'             = 'Exchange Online Protection'
    'EXCHANGEENTERPRISE'         = 'Exchange Online Plan 2'
    'EXCHANGESTANDARD'           = 'Exchange Online Plan 1'
    'FLOW_FREE'                  = 'Microsoft Flow Free'
    'GLOBAL_SERVICE_MONITOR'     = 'Global Service Monitor Online Service'
    'INTUNE_A'                   = 'Intune'
    'POWER_BI_PRO'               = 'Power BI Pro'
    'POWER_BI_STANDARD'          = 'Power BI'
    'POWERAPPS_INDIVIDUAL_USER'  = 'Microsoft PowerApps and Logic flows'
    'POWERAPPS_VIRAL'            = 'Microsoft Power Apps & Flow'
    'PROJECTCLIENT'              = 'Project Pro for Office 365'
    'PROJECTESSENTIALS'          = 'Project Online Essentials'
    'PROJECTONLINE_PLAN_1'       = 'Project Online Premium without Project Client'
    'PROJECTPREMIUM'             = 'Project Online Premium'
    'PROJECTPROFESSIONAL'        = 'Project Online Professional'
    'RIGHTSMANAGEMENT_ADHOC'     = 'Rights Management Adhoc'
    'STANDARDPACK'               = 'Office 365 Enterprise E1'
    'STREAM'                     = 'Microsoft Stream'
    'VISIOCLIENT'                = 'Visio Pro for Office 365'
	'DYN365_RETAIL_TRIAL'        = 'Dynamics 365 Trial'
	'MCOMEETADV'                 = 'AUDIO CONFERENCING'	
	'SPE_E3'                     = 'MS 365 E3'	
	'TEAMS_COMMERCIAL_TRIAL'     = 'Teams Commercial Cloud Trial'
    'PROJECT_MADEIRA_PREVIEW_IW_SKU'     = 'Dynamics 365 for Financials for IW'	
	'MCOSTANDARD'     			 = 'Skype for Business Online (Plan 2)'
}

# Pathes to export reports
$SkuReportExportPath = Join-Path -Path $PSScriptRoot -ChildPath MicrosoftLicenseReportTotal.csv
$ResultReportExportPath = Join-Path -Path $PSScriptRoot -ChildPath ConsumedLicensesReport.csv
$HtmlReportFilePath = Join-Path -Path $PSScriptRoot -ChildPath Office365LicenseSummary.html

# Mail html body settings
$HeadresultReport = '<div id=HeadresultReport>Consumed Licenses Report in the context of Countries</div>'

$HeadSkuReport = '<br><div id=HeadSkuReport>Summary License Report according Microsoft</div>'

$CssStyle = @" 
<style type="text/css"> 
body 
{ 
 font-family: Verdana, Arial; 
 font-size: 9pt; 
} 
 
#HeadresultReport 
{ 
    font-family: Verdana, Arial; 
    font-size: 11pt; 
    color: green; 
} 
#HeadSkuReport 
{ 
    font-family: Verdana, Arial; 
    font-size: 11pt; 
    color: green; 
} 
table 
{ 
 border-collapse: collapse; 
 font-size: 8pt; 
      
} 
table, td, th 
{ 
 border-color: gray; 
 border-style: solid; 
 border-width: 1px; 
 padding: 2px; 
 white-space: nowrap; 
} 
 
</style> 
"@

#endregion Parameters


#region Functions

function Convert-AccountSkuIdToName
{
    <# 
            .Synopsis 
            Converts AccountSkuId into license name 
            .DESCRIPTION 
            Converts AccountSkuId into Name e.g. 'TENANT:ENTERPRISEPACK' --> 'Office 365 Enterprise E3' 
            .EXAMPLE 
            'TENANT:ENTERPRISEPACK' | Convert-AccountSkuId 
    #>
       
    Process
    {
        $local:skuName = $_.Split(':')[-1]

        if ($Script:AccountSkuIdDecodeData[$local:skuName])
        {
            Write-Output -InputObject $Script:AccountSkuIdDecodeData[$local:skuName]
        }
        else
        {
            Write-Output -InputObject $_
        }        
    }  
}

#endregion Functions


#region Main

# Remove old reports
$null = $(Get-ChildItem -Path $ResultReportExportPath, $SkuReportExportPath, $HtmlReportFilePath -File | Remove-Item -Confirm:$false) 2>&1

# Get all licensed users in organization
#Connect-MsolService -Credential $O365Credentials -ErrorAction Stop

$licensedUsers = Get-MsolUser -All | Where-Object -FilterScript {$_.isLicensed -eq $true}


#region Obtaining of initial data

# Create PSCustom objects
$skusByLocation = @()
  
foreach ($user in $licensedUsers)
{
    foreach ($license in $user.Licenses.AccountSkuId)
    {
        $skusByLocation += [PSCustomObject]@{
            License  = $license
            Location = $user.UsageLocation
        }
    }
}
  
  
# Group PSCustom objects by UsageLocation, License
$licensesByLocation = $skusByLocation |
Group-Object -Property Location, License -NoElement |
ForEach-Object -Process {
    [PSCustomObject]@{
        Count    = $_.Count
        Location = $_.Name -replace ',.*$'
        License  = ($_.Name -replace '^.*,\s*' |
        Convert-AccountSkuIdToName)
    }
} |
Sort-Object -Property License
  
  
# Store all UsageLocation in a hashtable e.g. @{RU = 0; BY = 0; US = 0; .. ; XX = 0} 
$licensesByLocation.Location |
Sort-Object -Unique |
ForEach-Object -Begin {
    $usageLocations = [Ordered]@{}
} -Process {
    $usageLocations.$_ = 0
}

#endregion Obtaining of initial data


#region Generate data for "Consumed Licenses Report"

$resultReport = @()

foreach ($license in ($licensesByLocation | Group-Object -Property License))
{    
    # Create a new object
    $reportObject = [PSCustomObject]@{
        LicenseName = $license.Name
        Total       = ($license.Group | Measure-Object -Property Count -Sum).Sum
    }

    # Add all Locations as properties into the object, so the object will have properties: LicenseName, Total, RU, BY, US .. XX
    $reportObject | Add-Member -NotePropertyMembers $usageLocations  

    # Set nambers of consumed licenses for each location in a group 
    $license.Group | ForEach-Object -Process {
        $reportObject.$($_.Location) = $_.Count
    }      
      
    $resultReport += $reportObject
}

#endregion Generate data for "Consumed Licenses Report"


#region Generate data for "Summary License Report"

$summaryReport = Get-MsolAccountSku |
Select-Object -Property @{Name = 'LicenseName'; Expression = {$_.AccountSkuId | Convert-AccountSkuIdToName}}, ActiveUnits, WarningUnits, ConsumedUnits |
Sort-Object -Property LicenseName

#endregion Generate data for "Summary License Report"


# Export reports 
$summaryReport | Export-Csv -Path $SkuReportExportPath -NoTypeInformation
$resultReport | Export-Csv -Path $ResultReportExportPath -NoTypeInformation


# Create Html tables fragments
$summaryReport = $summaryReport | ConvertTo-Html -Fragment | Out-String
$resultReport = $resultReport | ConvertTo-Html -Fragment | Out-String


# Generate mail body report
$finalReport = ConvertTo-Html -Head $CssStyle -PostContent $HeadresultReport, $resultReport, $HeadSkuReport, $summaryReport | Out-String

# remove empty tables
$finalReport = $finalReport -replace '<table>\s*</table>' 

# Export mail body report
Out-File -InputObject $finalReport -FilePath $HtmlReportFilePath -Encoding utf8 -Force

[string]$body=$finalReport | convertto-html

# Change your SMTP server and sender receiver with subjectline #

Send-MailMessage -from "automation2@techwanderers.in" -To "automation@techwanderers.in" -Subject "Automation: Office 365 License Summary for Techwanderers" -Body "$finalReport </p></p><br></p></p><br></p></p></p><br></p></p>Best Regards,<br>Office 365 Automation Support<br>URL: https://www.youtube.com/c/technologywanderers<br>automation@yourdomain.in<br>India<br>"   -BodyAsHtml  -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -Credential $cred -SmtpServer smtp.office365.com -Port 587 -UseSsl -Attachments .\Office365LicenseSummary.html

Remove-Item -Path .\Office365LicenseSummary.html -Force

exit
