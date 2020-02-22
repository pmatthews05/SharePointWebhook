<#
.SYNOPSIS
- You need to be already signed in with PNP to your Site (Connect-PnpOnline -url:)
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
Set-RequestList

#Add to Left navigation
Set-NavigationNode -Title:"WebHook Example 1" -Location:"QuickLaunch" -Url:"/Lists/WebHook Example1"
Set-NavigationNode -Title:"WebHook Example 2" -Location:"QuickLaunch" -Url:"/Lists/WebHook Example2"

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