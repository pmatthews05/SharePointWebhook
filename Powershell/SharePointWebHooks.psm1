param (
    [Parameter(Position = 0)]
    [string]
    $ErrorActionOverride = $(throw "You must supply an error action preference"),

    [Parameter(Position = 1)]
    [string]
    $InformationOverride = $(throw "You must supply an information preference"),

    [Parameter(Position = 2)]
    [string]
    $VerboseOverride = $(throw "You must supply a verbose preference")
)

$ErrorActionPreference = $ErrorActionOverride
$InformationPreference = $InformationOverride
$VerbosePreference = $VerboseOverride


function Invoke-AzCommand {
    param(
        # The command to execute
        [Parameter(Mandatory)]
        [string]
        $Command,

        # Output that overrides displaying the command, e.g. when it contains a plain text password
        [string]
        $Message = $Command
    )

    Write-Information -MessageData:$Message

    # Az can output WARNINGS on STD_ERR which PowerShell interprets as Errors
    $CurrentErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue" 

    Invoke-Expression -Command:$Command
    $ExitCode = $LastExitCode
    Write-Information -MessageData:"Exit Code: $ExitCode"
    $ErrorActionPreference = $CurrentErrorActionPreference

    switch ($ExitCode) {
        0 {
            Write-Debug -Message:"Last exit code: $ExitCode"
        }
        default {
            throw $ExitCode
        }
    }
}

function Install-PnPPowerShell {
    Write-Information -MessageData:"Getting if the SharePoint PnP PowerShell Online module is available..."
    If (-not (Get-Module SharePointPnPPowerShellOnline)) {
        Write-Information -MessageData:"Installing the NuGet package provider..."
        Install-PackageProvider -Name:NuGet -Force -Scope:CurrentUser

        Write-Information -MessageData:"Installing the SharePoint PnP PowerShell Online module..."
        Install-Module SharePointPnPPowerShellOnline -Scope:CurrentUser -Force    
    }
    
    if ($VerbosePreference) {
        Set-PnPTraceLog -On -Level:Debug
    }
    else {
        Set-PnPTraceLog -Off
    }
}

function Connect-SharePointAsAppToken {
    param(
        [Parameter(Mandatory)]
        [String]
        $ClientId,

        [Parameter(Mandatory)]
        [String]
        $Tenant,
        
        [Parameter(Mandatory)]
        [String]
        $CertificateString,
        
        [Parameter(Mandatory)]
        [String]
        $SiteUrl
    )

    New-Item -Path:"$PSScriptRoot" -Name:"certificates" -ItemType:"directory" -Force

    [string]$CertificateFilePath = "$PSScriptRoot/certificates/$([System.IO.Path]::GetRandomFileName()).pfx"

    Write-Information -MessageData:"Storing the certificate..."
    # Store App Principal certificate as a .pfx file
    # Use code from https://docs.microsoft.com/en-us/azure/devops/pipelines/tasks/deploy/azure-key-vault?view=azure-devops
    $CertificateBase64String = [System.Convert]::FromBase64String("$CertificateString")
    $CertificateCollection = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
    $CertificateCollection.Import($CertificateBase64String, $null, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
    $Password = [System.Guid]::NewGuid().ToString()
    $CertificateExport = $CertificateCollection.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12, $Password)
    [System.IO.File]::WriteAllBytes($CertificateFilePath, $CertificateExport)

    $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force

    Write-Information -MessageData:"Connecting to the $SiteURL SharePoint site via the $ClientId App Token..."
    $Connection = Connect-PnPOnline `
        -Url:$SiteUrl `
        -ClientId:$ClientId `
        -Tenant:$Tenant `
        -CertificatePath:$CertificateFilePath `
        -CertificatePassword:$SecurePassword `
        -ReturnConnection

    Write-Output $Connection
}

function Set-SelfSignedCertificate {
    param(
        [Parameter(Mandatory)]
        $ApplicationRegistration,
        [Parameter(Mandatory)]
        [string]
        $Identifier,
        [Parameter(Mandatory)]
        [string]
        $KeyVaultName
    )
    # Must be 24 if used against an AD App Registration
    [int]$MonthsValidityGranted = 24
    [int]$MonthsValidityRequired = 21

    $keyVaultNameLower = $KeyVaultName.ToLower()
    [DateTime]$EndDate = $(Get-Date).AddMonths($MonthsValidityRequired)
    Write-Information -MessageData:"A certificate must be valid after $(Get-Date $EndDate -Format 'O' )."

    Write-Information -MessageData:"Checking for existing $Identifier certificate"
    $Certificates = Invoke-AzCommand -Command:"az keyvault certificate list --vault-name $KeyVaultName --include-pending $true --query ""[?id == 'https://$keyVaultNameLower.vault.azure.net/certificates/$Identifier']""" | ConvertFrom-Json

    $ValidCertificates = @()

    $Certificates | ForEach-Object {
        $Certificate = $PSItem

        $CertificateAttributes = $Certificate | Select-Object -ExpandProperty attributes
        $Expires = Get-Date $CertificateAttributes.expires

        Write-Information -MessageData:"The $($Certificate.id) certificate is valid until $(Get-Date $Expires -Format 'O')."
        if ($Expires -gt $EndDate) {
            $ValidCertificates += $Certificate
        }
    }

    if ($ValidCertificates.length -eq 0) {
        Write-Information -MessageData:"Creating a $Identifier certificate of 24 months validity, the minimum needed for an Azure AD App Registration..."
        Invoke-AzCommand -Command:"az keyvault certificate create --vault-name $KeyVaultName --name $Identifier --policy ""`@$PSScriptRoot\defaultpolicy.json"" --validity $MonthsValidityGranted" | Out-Null
    }
     
    Write-Information -MessageData:"Getting the $Identifier certificate..."
    $KeyVaultCertificate = Invoke-AzCommand -Command:"az keyvault certificate show --vault-name ""$KeyVaultName"" --name ""$Identifier""" | ConvertFrom-Json
        
    Write-Information -MessageData:"Updating the $($ApplicationRegistration.appId) App Registration key credentials..."
    Invoke-AzCommand -Command:"az ad app update --id ""$($ApplicationRegistration.appId)"" --key-type ""AsymmetricX509Cert"" --key-usage ""Verify"" --key-value ""$($KeyVaultCertificate.cer)""" -Message "az ad app update --id ""$($ApplicationRegistration.appId)"" --key-type ""AsymmetricX509Cert"" --key-usage ""Verify"" --key-value""****""" | Out-Null
    
    Write-Information -MessageData:"Listing the enabled versions of the $Identifier certificate..."
    $Versions = Invoke-AzCommand -Command:"az keyvault certificate list-versions --name $Identifier --vault-name $KeyVaultName --query ""[?attributes.enabled]""" | ConvertFrom-Json

    $Versions | Select-Object -Property id -ExpandProperty attributes | Sort-Object -Property created -Descending | Select-Object -Skip 1 | ForEach-Object {
        $Version = $PSItem
        $VersionId = $Version.id -split '/' | Select-Object -Last 1
    
        Write-Information -MessageData:"Disabling the $VersionId version..."
        Invoke-AzCommand -Command:"az keyvault certificate set-attributes --name ""$Identifier"" --vault-name ""$KeyVaultName"" --version ""$VersionId"" --enabled $false " | Out-Null
    } 
}

function Set-AzureKeyVaultSecrets {
    param(
        # The Key Vault Name
        [Parameter(Mandatory)]
        [string]
        $Name,

        # The secrets
        [Parameter(Mandatory)]
        $Secrets
    )

    $Secrets.Keys | ForEach-Object {
        $Key = $PSItem
    
        Write-Information -Message:"Adding the $Name key vault $Key secret..."
        az keyvault secret set --name $Key --vault-name $Name --value $Secrets[$Key] | Out-Null

        Write-Information -MessageData:"Listing the enabled versions of the $Key Secret..."
        $Versions = az keyvault secret list-versions --name $Key --vault-name $Name --query "[?attributes.enabled]" | ConvertFrom-Json
        $Versions | Select-Object -Property id -ExpandProperty attributes | Sort-Object -Property created -Descending | Select-Object -Skip 1 | ForEach-Object {
            $Version = $PSItem
            $VersionId = $Version.id -split '/' | Select-Object -Last 1

            Write-Information -MessageData:"Disabling the $VersionId version..."
            az keyvault secret set-attributes --name $key --vault-name $Name --version $VersionId --enabled $false | Out-Null
        }
    }
}
function Set-ApplicationIdToKeyVault {
    param(
        [Parameter(Mandatory)]
        [string]
        $ApplicationName,
        [Parameter(Mandatory)]
        [string]
        $ApplicationAppId,
        [Parameter(Mandatory)]
        [string]
        $KeyVaultName
    )
    $secrets = @{ };
    #Can't store characters in Secrets.
    $AppName = $ApplicationName -replace '_', '' -replace '-', ''
    $secrets.Add("$AppName", "$ApplicationAppId");

    Write-Information -Message:"Setting the Keyvault Secrets..."
    Set-AzureKeyVaultSecrets -Name:$KeyVaultName -Secrets:$secrets
} 
function ConvertTo-StorageAccountName {
    param(
        # The generic name to use for all related assets
        [Parameter(Mandatory)]
        [string]
        $Name
    )

    $StorageAccountName = $($Name -replace '-', '').ToLowerInvariant()
    if ($StorageAccountName.Length -gt 24) {
        $StorageAccountName = $StorageAccountName.Substring(0, 24);
    }
    

    Write-Output $StorageAccountName
}
function ConvertTo-KeyVaultAccountName {
    param(
        # The generic name to use for all related assets
        [Parameter(Mandatory)]
        [string]
        $Name
    )

    if ($Name.Length -gt 24) {
        $Name = $Name.Substring(0, 24);
    }
    
    Write-Output $Name
}
function Set-NavigationNode {
    param(
        [Parameter(Mandatory)]
        [string]
        $Title,
        [Parameter(Mandatory)]
        [string]
        $Location,
        [Parameter(Mandatory)]
        [string]
        $Url
    )

    $NavigtionItem = Get-PnPNavigationNode -Location:$Location | Where-Object { $_.Title -eq $Title }
    if ($null -eq $NavigtionItem) {
        $NavigationItem = Add-PnPNavigationNode -Title:$Title -Location:$Location -Url:$Url 
    }

    Write-Output $NavigationItem
}

function Set-Template {
    Write-Information -MessageData:"Applying Template For WebHook Example Lists"
    Apply-PnPProvisioningTemplate -Path:"$PSScriptRoot\Templates\ExampleLists.xml" -verbose
}
function Get-ServerPing {
    param(
        [Parameter(Mandatory)]
        [string]$Url,
        [int]$retryCount = 5,
        [int]$delaySeconds = 1
    )

    $retryAfterInterval = 0;
    $retryAttempts = 0;
    $backoffInterval = $delaySeconds;
    $Success = $false;
    while ($retryAttempts -lt $retryCount) {

        try {
            Write-Information -messageData:"Pinging url - $Url Attempt:$($retryAttempts+1)"
            
            $Response = Invoke-WebRequest $Url -UseBasicParsing 
            
            $StatusCode = $Response.StatusDescription
            $StatusCodeInt = $Response.StatusCode
            $ExceptionOccurred = $false
            Write-Information -messageData:"StatusCode: $StatusCode - ($StatusCodeInt)"
        }
        catch {
            if ($_.Exception.Response.StatusCode) {
                $StatusCode = $_.Exception.Response.StatusDescription
                $StatusCodeInt = $_.Exception.Response.StatusCode
                Write-Information -messageData:"StatusCode: $StatusCode - ($StatusCodeInt)"
            }
            else {
                $StatusText = $_.Exception.Message
                Write-Information -messageData:"Exception Message: $StatusText"
            }
            $ExceptionOccurred = $true
        }

        $Success = (($StatusCodeInt -eq 200) -and -not($ExceptionOccurred))

        if ($Success) { break; }

        $retryAfterInterval = $backoffInterval
        Start-Sleep -Seconds $retryAfterInterval
        $retryAttempts++
        $backoffInterval *= 2
    }
    
    Write-Output $Success
}                 
function Set-AppPermissionsAndGrant {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $ApplicationRegistration,        
        [Parameter(Mandatory)]
        [string[]]
        $RequestedPermissions
    )

    $Permissions = Get-Permissions -RequestedPermissions:$RequestedPermissions
    
    Write-Information -MessageData:"Checking the $($ApplicationRegistration.DisplayName) App Registration..."

    Write-Information -MessageData:"Listing all the permissions on the $($ApplicationRegistration.AppId) App Registration as further attempts to list will fail once a permission is requested until the permissions are granted..."  
    $CurrentPermissions = Invoke-AzCommand "az ad app permission list --id ""$($ApplicationRegistration.AppId)""" | ConvertFrom-Json

    $Permissions | ForEach-Object {
        $Permission = $PSItem
    
        Write-Information -MessageData:"Checking if the $($Permission.apiId) $($Permission.PermissionId) permission is already present..."
        $ExistingPermissions = $CurrentPermissions |
        Where-Object { $PSItem.resourceAppId -eq $Permission.apiId } |
        Select-Object -ExpandProperty "resourceAccess" |
        Where-Object { $PSItem.id -eq $Permission.PermissionId }
    
        if (-not $ExistingPermissions) {
            Write-Information -MessageData:"Adding the $($Permission.apiId) $($Permission.PermissionId) permission..."
            # =Role makes the permission an application permission
            # =Scope is used to indicate a delegate permission
            Invoke-AzCommand -Command:"az ad app permission add --id ""$($ApplicationRegistration.AppId)"" --api ""$($Permission.apiId)"" --api-permissions ""$($Permission.PermissionId)=$($Permission.Type)"""
        }
        Write-Information -MessageData:"Granting admin consent to the $($ApplicationRegistration.appId) Azure AD App Registration..."
        Invoke-AZCommand -Command:"az ad app permission admin-consent --id ""$($ApplicationRegistration.appId)"""    
    }
}

function Get-Permissions {
    param(
        # The generic name to use for all related assets
        [Parameter(Mandatory)]
        [string[]]
        $RequestedPermissions
    )

    $Permissions = $RequestedPermissions | ForEach-Object {
        @{
            "Application.SharePoint.FullControl.All" = @{
                apiId        = "00000003-0000-0ff1-ce00-000000000000"
                permissionId = "678536fe-1083-478a-9c59-b99265e6b0d3"
                type         = "Role"
            }
        }[$PSItem]
    }

    Write-Output $Permissions
}

function Set-OriginForAzureFunction {
    param (
        [Parameter(Mandatory)]
        [string]
        $FunctionAppIdentifier,
        [Parameter(Mandatory)]
        [string[]]
        $Origins

    )
    $Origins | ForEach-Object {
        $origin = $PSItem
        $ExistOrigin = Invoke-AzCommand -Command:"az functionapp cors show --name $FunctionAppIdentifier --query ""allowedOrigins | [?contains(@, '$origin')]""" | ConvertFrom-Json

        if (-not $ExistOrigin) {
            Write-Verbose -Message:"Adding the $FunctionAppIdentifier CORS Value $origin"
            Invoke-AzCommand -Command:"az functionapp cors add --name $FunctionAppIdentifier --allowed-origins $origin" | Out-Null
        }
    }
}

function Set-WebHook {
    param(
        [Parameter(Mandatory)]
        [string]$List,
        [Parameter(Mandatory)]
        [string]$ServerNotificationUrl,
        [Parameter(Mandatory)]
        [string]$ClientState,
        [Parameter(Mandatory)]
        [int]$ExpiresInDays
    )

    Write-Information -MessageData:"Looking for existing Webhook"
    $WebHooks = Get-PnPWebhookSubscriptions -List:$list 
    $WebHook = $webHooks | Where-Object { $_.NotificationUrl -eq $ServerNotificationUrl }

    $ExpiredDate = (Get-Date).AddDays($ExpiresInDays)

    #Set
    if ($WebHook) {
        Write-Information -MessageData:"WebHook found, updating Expired Date to $ExpiredDate"
        Set-PnPWebhookSubscription -List:$List -Subscription:$WebHook -NotificationUrl:$ServerNotificationUrl -ExpirationDate:$ExpiredDate
    }
    else {
        #Add
        Write-Information -MessageData:"Adding new Webhook, with expired date $ExpiredDate"
        Add-PnPWebhookSubscription -List:$List -NotificationUrl:$ServerNotificationUrl -ExpirationDate:$ExpiredDate -ClientState:$ClientState
    }
}