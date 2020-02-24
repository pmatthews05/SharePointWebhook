<#
.SYNOPSIS
- You need to be already signed in with AZ CLI
- Creates:
	-Resource Group
	-Storage Account
	-Function App with Application Insights
	-KeyVault
	-Self Signed Certificate
	-Application Registration
	-Grants Access
 Defaults to the UKSouth location. This does not check if the name already exists.
.EXAMPLE
.\Install-AzureEnvironment.ps1 -Environment "cfcodedev" -Name:"sharepointwebhook"
.EXAMPLE
.\Install-AzureEnvironment.ps1 -Environment "cfcodedev" -Name:"sharepointwebhook" -Location:"westus"
#>


param(
	[Parameter(Mandatory)]
	[string]
	$Environment,
	[Parameter(Mandatory)]
	[string]
	$Name,
	[string]
	$Location = "uksouth"
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

[string]$Identifier = "$Environment-$Name"


Import-Module -Name:"$PSScriptRoot\SharePointWebHooks" -Force -ArgumentList:@(
	$ErrorActionPreference,
	$InformationPreference,
	$VerbosePreference
)

Write-Information "Setting the Azure CLI defaults..."
Invoke-AzCommand -command:"az configure --defaults location=$Location"

Write-Information -Message:"Creating the $Identifier resource group..."
Invoke-AzCommand -command:"az group create --name $Identifier"

Write-Information "Setting the Azure CLI defaults..."
Invoke-AzCommand -command:"az configure --defaults location=$Location group=$Identifier"

[string]$StorageAccountName = ConvertTo-StorageAccountName -Name:$Identifier
Write-Information "Create a Storage Account..."
Invoke-AzCommand -Command:"az storage account create --name ""$StorageAccountName"" --access-tier ""Hot"" --sku ""Standard_LRS"" --kind ""StorageV2"" --https-only $true"

[string]$KeyVaultName = ConvertTo-KeyVaultAccountName -name:$Identifier
Write-Information "Create a KeyVault..."
Invoke-AzCommand -Command:"az keyvault create --name ""$KeyVaultName"""

Write-Information "Create an Application Insight and Azure Function..."
$appInsights = Invoke-AzCommand -Command:"az resource create --resource-type 'Microsoft.Insights/components' --name '$Identifier' --properties '{\""Application_Type\"":\""web\""}'" | ForEach-Object { $PSItem -join '' } | ConvertFrom-Json

Invoke-AzCommand -Command:"az functionapp create --name ""$Identifier"" --storage-account ""$StorageAccountName"" --consumption-plan-location ""$Location"" --runtime ""dotnet"" --app-insights '$Identifier' --app-insights-key '$($appInsights.properties.InstrumentationKey)'"


Write-Information -MessageData:"Assigning the $Identifier function app identity."
$identityJson = Invoke-AzCommand -Command:"az webapp identity assign --name ""$Identifier""" | ConvertFrom-Json

Write-Information -MessageData:"Setting the $KeyVaultName key vault access policy for the Azure Function passed in..."
Invoke-AzCommand -Command:"az keyvault set-policy --name $KeyVaultName --key-permissions get list --secret-permissions get list --certificate-permissions get list --object-id $($identityJson.principalId)" | Out-Null

Write-Information "Updating the Azure Function Configuration Settings"
az webapp config appsettings set --name $Identifier --settings KeyVaultName=$KeyVaultName Tenant=$Environment CertificateName=$Identifier FUNCTIONS_EXTENSION_VERSION=~1 | Out-Null

Write-Information "Updating Azure Function Origins"
[string[]]$origins = "https://$Environment.sharepoint.com"
Set-OriginForAzureFunction -FunctionAppIdentifier:$Identifier -Origins:$origins
 
Write-Information "Create an Application Registration"
$AppReg = Invoke-AzCommand -Command:"az ad app create --display-name ""$($Identifier)""" | ConvertFrom-Json

Write-Information "Store the Application Client ID in Keyvault"
Set-ApplicationIdToKeyVault -ApplicationName:"$($AppReg.DisplayName)" -ApplicationAppId:"$($AppReg.appId)" -KeyVaultName:"$KeyVaultName"

Write-Information "Create a self signed Certificate and put in KeyVault"
Set-SelfSignedCertificate -ApplicationRegistration:$AppReg -Identifier:$Identifier -KeyVaultName:$KeyVaultName

#Only seems to work with SharePoint Only.
Write-Information "Giving Application SharePoint Full Control Permission"
Set-AppPermissionsAndGrant -ApplicationRegistration:$AppReg -RequestedPermissions:"Application.SharePoint.FullControl.All"

Write-Information "Complete"