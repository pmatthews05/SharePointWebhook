<#
.SYNOPSIS
- You need to be already signed in with az login
- Creates:
    -Two SharePoint Lists
    -Adds links to lists in Left Navigation
    -Connect up Lists to Webhooks
.EXAMPLE
.\Install-WebhookLists.ps1 -Environment "cfcodedev" -Name:"sharepointwebhook" -SharePointSiteurl:"https://cfcodedev.sharepoint.com/sites/webhooks"
#>
param(
    [Parameter(Mandatory)]
    [string]
    $Environment,
    [Parameter(Mandatory)]
    [string]
    $Name,
    [Parameter(Mandatory)]
    [string]
    $SharePointSiteUrl
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$Identifier = "$Environment-$Name"

az configure --defaults location=uksouth group=$Identifier

Import-Module -Name:"$PSScriptRoot\SharePointWebHooks" -Force -ArgumentList:@($ErrorActionPreference, $InformationPreference, $VerbosePreference)


$KeyVaultName = ConvertTo-KeyVaultAccountName -name:$Identifier
#Installs SharePoint Powershell
Install-PnPPowerShell

Write-Information -MessageData:"Getting the Azure AD App registration to access SharePoint Online with..."
$AppRegistrations = az ad app list --query "[?starts_with(displayName, '$Identifier')]" | ConvertFrom-Json
$AppRegistration = $AppRegistrations | Select-Object -first 1

$Secret = Invoke-AzCommand -Command:"az keyvault secret show --vault-name ""$KeyVaultName"" --name ""$Identifier""" | ConvertFrom-Json

Write-Information -MessageData:"Connecting to SharePoint Online $SharePointSiteUrl with App Token..."
Connect-SharePointAsAppToken `
    -Tenant:$AppRegistration.publisherDomain `
    -ClientId:$AppRegistration.appId `
    -CertificateString:$Secret.value `
    -SiteUrl:$SharePointSiteUrl | Out-Null

#Create SharePoint Lists
Write-Information -MessageData:"Applying SharePoint Template..."
Set-Template

$WebServerRelativeUrl = (Get-PnpWeb).ServerRelativeUrl
#Add to Left navigation
Write-Information -MessageData:"Updating Quicklaunch Navigation..."
Set-NavigationNode -Title:"WebHook Example 1" -Location:"QuickLaunch" -Url:$WebServerRelativeUrl"/Lists/WebHook Example1"
Set-NavigationNode -Title:"WebHook Example 2" -Location:"QuickLaunch" -Url:$WebServerRelativeUrl"/Lists/WebHook Example2"

$webHookUrl = "https://$Identifier.azurewebsites.net/api/SharePointWebHook"

Write-Information -MessageData:"Pinging Azure Function to wake up..."
$Success = Get-ServerPing -Url:"https://$Identifier.azurewebsites.net"

if ($Success) {
    Write-Information -MessageData "Setting the webhook"
    Set-WebHook -List:"WebHook Example1" -ServerNotificationUrl $webHookUrl -ClientState "1241fe65-ce63-4ca3-8827-014ea2f93bd5" -ExpiresInDays 90
    Set-WebHook -List:"WebHook Example2" -ServerNotificationUrl $webHookUrl -ClientState "1241fe65-ce63-4ca3-8827-014ea2f93bd5" -ExpiresInDays 90
}
else {
    Write-Error "Unable to ping Azure WebSite https://$Identifier.azurewebsites.net to call the /api/SharePointWebHook."
}

Disconnect-PnPOnline
Write-Information "Complete"