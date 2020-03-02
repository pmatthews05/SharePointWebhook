<#
.SYNOPSIS
- You need to be already signed in with PNP to your Site (Connect-PnpOnline -url:) this will not work with MFA. Use Install-WebHookList.ps1 instead.
- Creates:
    -Two SharePoint Lists
    -Adds links to lists in Left Navigation
    -Connect up Lists to Webhooks
.EXAMPLE
.\Install-WebhookLists.ps1 -Environment "cfcodedev" -Name:"sharepointwebhook"
#>
param(
    [Parameter(Mandatory)]
    [string]
    $Environment,
    [Parameter(Mandatory)]
    [string]
    $Name
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$identity = "$Environment-$Name"

Import-Module -Name:"$PSScriptRoot\SharePointWebHooks" -Force -ArgumentList:@($ErrorActionPreference, $InformationPreference, $VerbosePreference)

#Create SharePoint Lists
Write-Information -MessageData:"Applying SharePoint Template..."
Set-Template

$WebServerRelativeUrl = (Get-PnpWeb).ServerRelativeUrl
#Add to Left navigation
Write-Information -MessageData:"Updating Quicklaunch Navigation..."
Set-NavigationNode -Title:"WebHook Example1" -Location:"QuickLaunch" -Url:$WebServerRelativeUrl"/Lists/WebHook Example1"
Set-NavigationNode -Title:"WebHook Example2" -Location:"QuickLaunch" -Url:$WebServerRelativeUrl"/Lists/WebHook Example2"

$webHookUrl = "https://$identity.azurewebsites.net/api/SharePointWebHook"

Write-Information -MessageData:"Pinging Azure Function to wake up..."
$Success = Get-ServerPing -Url:"https://$identity.azurewebsites.net"

if ($Success) {
    Write-Information -MessageData "Setting the webhook"
    Set-WebHook -List:"WebHook Example1" -ServerNotificationUrl $webHookUrl -ClientState "1241fe65-ce63-4ca3-8827-014ea2f93bd5" -ExpiresInDays 90
    Set-WebHook -List:"WebHook Example2" -ServerNotificationUrl $webHookUrl -ClientState "1241fe65-ce63-4ca3-8827-014ea2f93bd5" -ExpiresInDays 90
}
else {
    Write-Error "Unable to ping Azure WebSite https://$identity.azurewebsites.net to call the /api/SharePointWebHook."
}
Write-Information "Complete"