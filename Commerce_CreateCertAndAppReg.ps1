#Requires -RunAsAdministrator

function Invoke-D365CommerceDevWizard {

    [CmdletBinding()]
    param (
        [ValidateNotNullOrEmpty()]
        [string]$ComputerName = $env:computerName,
        [ValidateNotNullOrEmpty()]
        [int]$CsuPort = 446,
        [ValidateNotNullOrEmpty()]
        [int]$HsPort = 450,
        [ValidateNotNullOrEmpty()]
        [string]$CsuAppName = "app-d365-csu-$ComputerName",
        [ValidateNotNullOrEmpty()]
        [string]$CposAppName = "app-d365-cpos-$ComputerName"
    )

    try
    {
        Set-Location 'C:\Temp'
        
        $CsuUrl = "https://$($ComputerName):$($CsuPort)/RetailServer/Commerce"
        $CsuHealthCheckUrl = "https://$($ComputerName):$($CsuPort)/RetailServer/healthcheck?testname=ping"
        $CposUrl = "https://$($ComputerName):$($CsuPort)/POS"
        $HsPingUrl = "https://$($ComputerName):$($HsPort)/HardwareStation/ping"

        Write-Header

        if (-not (Prompt "Are you ready to start? (Y/N)")) {
            return
        }

        $IsUDE = Prompt "Are you connecting to UDE (Unified Developer Experience) environment? (Y/N)"

        if (-not (Install-Prerequisites $IsUDE)) {
            Write-Error "Failed to install required PowerShell modules"
            return
        }

        if ($IsUDE) {
            $D365FOBaseURL = Read-Host "Please provide base D365FO URL where to perform data configuration"
        } else {
            $D365FOBaseURL = (Get-D365Url).Url
        }

        Connect-MgGraph -Scopes "Application.ReadWrite.All","User.Read" -NoWelcome -ErrorAction Stop

        Write-Host "`nAuthentification succeed. Press any key to start process...`n"
        $null = [Console]::ReadKey($true)

        # Step 1
        $CsuCert = New-CSUCertificate -ComputerName $ComputerName

        # Step 2
        $HsCert = New-HSCertificate -ComputerName $ComputerName

        # Step 3
        $CsuApp = New-CSUAppRegistration -ComputerName $ComputerName `
            -CsuCert $CsuCert `
            -CsuAppName $CsuAppName

        # Step 4
        $CposApp = New-CPOSAppRegistration -ComputerName $ComputerName `
            -CsuApp $CsuApp `
            -CposAppName $CposAppName

        # Step 5
        New-D365DataConfiguration -ComputerName $ComputerName `
            -HsPort $HsPort `
            -CsuAppId $CsuApp.AppId `
            -CposAppId $CposApp.AppId `
            -CsuUrl $CsuUrl `
            -CposUrl $CposUrl `
            -IsUDE $IsUDE `
            -D365FOBaseURL $D365FOBaseURL

        # Manual steps instructions
        Write-ManualSteps -ComputerName $ComputerName `
            -HsPort $HsPort `
            -CsuAppId $CsuApp.AppId `
            -CposAppId $CposApp.AppId `
            -CsuUrl $CsuUrl `
            -CposUrl $CposUrl

        # Step 6
        Invoke-CommerceInstallers -ComputerName $ComputerName `
            -CposApp $CposApp `
            -CsuApp $CsuApp `
            -CsuCert $CsuCert `
            -CsuUrl $CsuUrl `
            -CsuPort $CsuPort `
            -HsCert $HsCert `
            -HsPort $HsPort

        Write-Host "`nCongratulations! Check the following URLs for the status of installation

    CSU HealthCheck: $CsuHealthCheckUrl
    HS Ping: $HsPingUrl

    CSU URL: $CsuUrl
    Store Commerce for Web: $CposUrl
    Store Commerce for Windows: Check Start menu...`n" -ForegroundColor Green

    Start-Process $CsuHealthCheckUrl
    Start-Process $HsPingUrl

    }
    catch [Exception]
    {
        Write-Host "Something threw an exception or used Write-Error"
        Write-Error $_
    }
    finally
    {
        Write-Host "Press any key to exit..."

        $null = [Console]::ReadKey($true)

        # Stay in powershell
        powershell -NoLogo
    }
}

function Write-Header {
    Write-Host '
 /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\ 
( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )
 > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ < 
 /\_/\                                                                                             /\_/\ 
( o.o )                                                                                           ( o.o )
 > ^ <                                  .....:::.                                                  > ^ < 
 /\_/\                     .::^~~!7?JYY5555P?^BG!                                                  /\_/\ 
( o.o )                    ^5G&&&&&@&@&&@@@@5 GG~                                                 ( o.o )
 > ^ <                     .5G@@@@@@@@@@@@@@5 GP~                                                  > ^ < 
 /\_/\                      P?P55Y5555555PPG? GP7                                                  /\_/\ 
( o.o )                     :^::::::GJG~^^~~~~G5!                                                 ( o.o )
 > ^ <                    ...       P.P      ..                                                    > ^ < 
 /\_/\                  :!!^^~~~~~~~G^G::....               /$$$$$$$   /$$$$$$   /$$$$$$           /\_/\ 
( o.o )               .JJ~.....   . !Y?.:^^^!JG!           | $$__  $$ /$$__  $$ /$$__  $$         ( o.o )
 > ^ <                ?7G#BGPYJ??!^^:::....~7!P!           | $$  \ $$| $$  \ $$| $$  \__/          > ^ < 
 /\_/\              .!?Y@@@@@@&&@&.!?^?!.!Y!~^P^           | $$$$$$$/| $$  | $$|  $$$$$$           /\_/\ 
( o.o )        .:~7?Y57J55PB##&@@5.?^77. 7!^^:P:           | $$____/ | $$  | $$ \____  $$         ( o.o )
 > ^ <      ~!?YJ5?J??77?!J77?7J5:^7!7:  J^^^^G:           | $$      | $$  | $$ /$$  \ $$          > ^ < 
 /\_/\      G~~??57??!?!?!77??~?!~!~!~:.~?^^^^G:           | $$      |  $$$$$$/|  $$$$$$/          /\_/\ 
( o.o )     5.  ..:~!?J7J!?J7YYJY?J!J~~?!^::^^G.           |__/       \______/  \______/          ( o.o )
 > ^ <      ?7:.  .~:..^~!!Y??7J?7?7!~^::::^~?5                                                    > ^ < 
 /\_/\       .^~!~?P?     ..:^Y7!~^^::::^~7?7^.                                                    /\_/\ 
( o.o )          .:~!~^:.     J^^:::^~!?!^.                                                       ( o.o )
 > ^ <               ..^!!~:. 7~~7!!7^.                                                            > ^ < 
 /\_/\                    .^~~?J!^.                                                                /\_/\ 
( o.o )                                                                                           ( o.o )
 > ^ <                                                                                             > ^ < 
 /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\ 
( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )
 > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ < 
' -ForegroundColor Green

    Write-Host "Welcome to the Commerce configuration script, following steps will be executed:
    1. Create self-service certificate for CSU (cloud scale unit)
    2. Create self-service certificate for HS (hardware station)
    3. Create and configure Microsoft EntraID application for CSU named $CsuAppName
    4. Create and configure Microsoft EntraID application for Store Commerce for Web (Cloud POS) named $CposAppName
    5. Run installers of CSU, Store Commerce App for Windows and HS (you need to download them manually from LCS)

    Official Docs: https://learn.microsoft.com/en-us/dynamics365/commerce/dev-itpro/install-csu-dev-env`n"

}

function Prompt {
    param (
        [Parameter(Position = 0)]
        [string]$promptText
    )

    do
    {
        $answer = Read-Host $promptText
        if ($answer -eq "N") {
            return $False
        }
    } while ($answer -ne "Y")

    return $True
}

function Install-Prerequisites {

    param (
        [Parameter(Position = 0)]
        [string]$IsUDE
    )

    $success = $True

    if (-not (Get-Module Microsoft.Graph.Applications -ListAvailable)) {
        $graph = $True
        Write-Host "`nMicrosoft.Graph.Applications module is missing..." -ForegroundColor Yellow
        Write-Host "(Used for creating App registrations in Microsoft Entra ID)"
    }

    if (-not($IsUDE) -and -not(Get-Module d365fo.tools -ListAvailable)) {
        $tools = $True
        Write-Host "`nd365fo.tools module is missing..." -ForegroundColor Yellow
        Write-Host "(Used for getting current D365FO URL)"
    }

    if (-not (Get-Module d365fo.integrations -ListAvailable)) {
        $integrations = $True
        Write-Host "`nd365fo.integrations module is missing..." -ForegroundColor Yellow
        Write-Host "(Used for creating configuration records in D365FO)"
    }

    if ($graph -or ($tools -or $integrations))
    {
        if (-not (Prompt "`nDo you want to proceed with installation of missing PowerShell modules for current user? (Y/N)")) {
            return $False
        }

        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

        if ($graph) {
            Install-Module Microsoft.Graph.Applications -Scope CurrentUser
            if (-not $?) {
                $success = $False
                Write-Error "Failed to install module, please check the error and try to install manually"
                Write-Error "Install-Module Microsoft.Graph.Applications -Scope CurrentUser"
            }
        }

        if ($tools) {
            Install-Module d365fo.tools -Scope CurrentUser
            if (-not $?) {
                $success = $False
                Write-Error "Failed to install module, please check the error and try to install manually"
                Write-Error "Install-Module d365fo.tools -Scope CurrentUser"
            }
        }

        if ($integrations) {
            Install-Module d365fo.integrations -Scope CurrentUser
            if (-not $?) {
                $success = $False
                Write-Error "Failed to install module, please check the error and try to install manually"
                Write-Error "Install-Module d365fo.integrations -Scope CurrentUser"
            }
        }
    }

    return $success
}

function New-CSUCertificate {

    param (
        [ValidateNotNullOrEmpty()]
        [string]$ComputerName = $env:computerName,
        [ValidateNotNullOrEmpty()]
        [string]$CsuCertFriendlyName = "CSU2024"
    )

    Write-Host "[1] Generate self-signed certificate for CSU named $CsuCertFriendlyName"

    $CsuCert = Get-ChildItem cert: -Recurse | Where-Object { $_.FriendlyName -eq $CsuCertFriendlyName } | Select-Object -first 1
    
    if ($null -eq $CsuCert) {
        
        $CsuCert = New-SelfSignedCertificate -FriendlyName $CsuCertFriendlyName `
            -DnsName $ComputerName `
            -CertStoreLocation Cert:\LocalMachine\My `
            -KeyUsage DigitalSignature,DataEncipherment,KeyEncipherment

        Write-Host "    Install certificate into root storage"
        $mypwd = ConvertTo-SecureString -String '1234' -Force -AsPlainText
        Export-PfxCertificate -Cert $CsuCert -FilePath "$env:temp\csu2024.pfx" -Password $mypwd
        Import-PfxCertificate -CertStoreLocation cert:\LocalMachine\Root -FilePath "$env:temp\csu2024.pfx" -Password $mypwd
    } else {
        Write-Host "    Self-signed certificate for CSU already exists, skipping..." -ForegroundColor Yellow
    }

    Write-Host "    $CsuCertFriendlyName Cert Thumbprint: $($CsuCert.Thumbprint)"
    return $CsuCert
}

function New-HSCertificate {

    param (
        [ValidateNotNullOrEmpty()]
        [string]$ComputerName = $env:computerName,
        [ValidateNotNullOrEmpty()]
        [string]$HsCertFriendlyName = "HS2024"
    )

    Write-Host "[2] Generate self-signed certificate for HS named $HsCertFriendlyName"

    $HsCert = Get-ChildItem cert: -Recurse | Where-Object { $_.FriendlyName -eq $HsCertFriendlyName } | Select-Object -first 1
    
    if ($null -eq $HsCert) {

        $HsCert = New-SelfSignedCertificate -FriendlyName $HsCertFriendlyName `
            -DnsName $ComputerName `
            -CertStoreLocation Cert:\LocalMachine\My `
            -KeyUsage DigitalSignature,DataEncipherment,KeyEncipherment

        Write-Host "    Install certificate into root storage"
        $mypwd = ConvertTo-SecureString -String '1234' -Force -AsPlainText
        Export-PfxCertificate -Cert $HsCert -FilePath "$env:temp\hs2024.pfx" -Password $mypwd
        Import-PfxCertificate -CertStoreLocation cert:\LocalMachine\Root -FilePath "$env:temp\hs2024.pfx" -Password $mypwd

    } else {
        Write-Host "    Self-signed certificate for HS already exists, skipping..." -ForegroundColor Yellow
    }

    Write-Host "    $HsCertFriendlyName Cert Thumbprint: $($HsCert.Thumbprint)"
    return $HsCert
}

function New-CSUAppRegistration {

    param (
        [ValidateNotNullOrEmpty()]
        [string]$ComputerName = $env:computerName,
        [Parameter(Mandatory)]
        [string]$CsuAppName,
        [Parameter(Mandatory)]
        [object]$CsuCert
    )

    Write-Host "[3] Create new Azure App registration for CSU with certificate: $CsuAppName"

    $CsuApp = Get-MgApplication -Filter "DisplayName eq '$CsuAppName'"
    if ($null -eq $CsuApp) {

        [byte[]] $certData = $CsuCert.RawData

        $LegacyAccessFullApiId = $([guid]::NewGuid())
        $api = @{
            oauth2PermissionScopes = @(
                @{
                AdminConsentDescription = "Gives access to CSU APIs on $($ComputerName)"
                AdminConsentDisplayName = "Access CSU on $($ComputerName)"
                Type = "User"
                Value = "Legacy.Access.Full"
                Id = $LegacyAccessFullApiId
            }
            )
        }

        $CsuApp = New-MgApplication -DisplayName $CsuAppName `
            -SignInAudience "AzureADMyOrg" `
            -Api $api `
            -KeyCredentials @(@{ Type = "AsymmetricX509Cert"; Usage = "Verify"; Key = $certData }) `
            -ErrorAction Stop

        Write-Host "    Add app URI"
        Update-MgApplication -ApplicationId $CsuApp.id `
            -IdentifierUris @("api://$($CsuApp.AppId)") `
            -ErrorAction Stop

        Write-Host "    Create corresponding service principal for CSU app"
        New-MgServicePrincipal -AppId $CsuApp.AppId -ErrorAction Stop > $null
    } else {
        Write-Host "    CSU App Registration already exists, skipping..." -ForegroundColor Yellow
    }
          
    Write-Host "    CSU AppId: $($CsuApp.AppId)"

    return $CsuApp
}

function New-CPOSAppRegistration {

    param (
        [ValidateNotNullOrEmpty()]
        [string]$ComputerName = $env:computerName,
        [Parameter(Mandatory)]
        [string]$CposAppName,
        [Parameter(Mandatory)]
        [object]$CsuApp
    )

    Write-Host "[4] Create new Azure App registration for Store Commerce for Web: $CposAppName"

    $CposApp = Get-MgApplication -Filter "DisplayName eq '$CposAppName'"
    if ($null -eq $CposApp) {

        $spa = @{
            RedirectUris = @("https://$($ComputerName):$($csuPort)/POS/")
        }
        
        $web = @{
            ImplicitGrantSettings = @{
                EnableIdTokenIssuance = $true
                EnableAccessTokenIssuance = $true
            }
        }

        $scopeId_UserRead = (Find-MgGraphPermission User.Read -ExactMatch -PermissionType Delegated).Id

        $scopeId_LegacyAccessFull = ($CsuApp.Api.Oauth2PermissionScopes |
            Where-Object Value -eq "Legacy.Access.Full" |
            Select-Object -First 1).Id

        if ($null -eq $scopeId_LegacyAccessFull) {
            throw "Cannot find API scope for CSU App: Legacy.Access.Full, please check if CSU App configured correctly"
        }

        $requiredResourceAccess = @(
            @{
                ResourceAppId = $($CsuApp.AppId)
                ResourceAccess = @(@{ 
                    Id = $scopeId_LegacyAccessFull
                    Type = "Scope"
                })
            }
            @{
                ResourceAppId = "00000003-0000-0000-c000-000000000000"
                ResourceAccess = @(@{ 
                    Id = $scopeId_UserRead
                    Type = "Scope"
                })
            }
        )
        
        $optionalClaims = @{
            AccessToken = @(
                @{
                    Name = "sid"
                }
            )
            IdToken = @(
                @{
                    Name = "sid"
                }
            )
        }

        $CposApp = New-MgApplication -DisplayName $CposAppName `
            -SignInAudience "AzureADMyOrg" `
            -Spa $spa `
            -Web $web `
            -RequiredResourceAccess $requiredResourceAccess `
            -OptionalClaims $optionalClaims `
            -ErrorAction Stop

        Write-Host "    Create corresponding service principal for Store Commerce for Web app"
        New-MgServicePrincipal -AppId $CposApp.AppId -ErrorAction Stop > $null

    } else {
        Write-Host "    CSU App Registration already exists, skipping..." -ForegroundColor Yellow
    }

    Write-Host "    CPOS AppId: $($CposApp.AppId)"

    return $CposApp
}

function Write-ManualSteps {
    param (
        [ValidateNotNullOrEmpty()]
        [string]$ComputerName = $env:computerName
    )
    

    Write-Host "

Manual steps are required in HQ (D365FO):

    1. Open 'Channel database' form, find newly created record (e.g. DevSealedCSU or $($ComputerName))
      Select Download > Configuration file and save it under C:\Temp\StoreSystemSetup.xml" -ForegroundColor yellow

    Write-Host "
Download following installers:

    2. Download and install http://monroecs.com/oposccos_current.htm for Hardware Station"

    if (Prompt "Open the link in default browser? (Y/N)") {
        Start-Process "http://monroecs.com/oposccos_current.htm"
    }

    Write-Host "
    3. Download and install two .NET 6.0 runtimes: https://dotnet.microsoft.com/en-us/download/dotnet/6.0
      In the 'ASP.Net Core Runtime 6.0.X section', select the Hosting Bundle installer for Windows.
      In the '.NET Desktop Runtime 6.0.X' section, select the x64 installer for Windows.

    Some of the could be already installed depending on the environment."

    if (Prompt "Open the link in default browser? (Y/N)") {
        Start-Process "https://dotnet.microsoft.com/en-us/download/dotnet/6.0"
        Start-Process "https://dotnet.microsoft.com/en-us/download/dotnet/6.0"
    }

    Write-Host "
    4. Open LCS -> 'Shared Asset Library' -> 'Retail Self-service package' and download sealed installers (check the version):
      https://eu.lcs.dynamics.com/V2/SharedAssetLibrary
      Download following installers:
        10.0.XX - Commerce Peripheral Simulator
        10.0.XX - Hardware Station (SEALED)
        10.0.XX - Commerce Scale Unit (SEALED)
        10.0.XX - Store Commerce"

    if (Prompt "Open the link in default browser? (Y/N)") {
        Start-Process "https://eu.lcs.dynamics.com/V2/SharedAssetLibrary"
    }
}

function New-D365DataConfiguration {
    param (
        [ValidateNotNullOrEmpty()]
        [string]$ComputerName = $env:computerName,
        [Parameter(Mandatory)]
        [string]$HsPort,
        [Parameter(Mandatory)]
        [string]$CposAppId,
        [Parameter(Mandatory)]
        [string]$CsuAppId,
        [Parameter(Mandatory)]
        [string]$CsuUrl,
        [Parameter(Mandatory)]
        [string]$CposUrl,
        [Parameter(Mandatory)]
        [string]$IsUDE,
        [Parameter(Mandatory)]
        [string]$D365FOBaseUrl
    )

    Write-Host "`n[5] Create configuration records in D365FO with d365fo.integrations"

    $tenantId = (Get-MgContext).TenantId

    $syncRequired = $False

    if ($IsUDE) {
        $channelDatabaseName = $ComputerName
        $retailChannelProfileName = "$($ComputerName)CSUProfile"
    } else {
        $channelDatabaseName = 'DevSealedCSU'
        $retailChannelProfileName = 'DevSealedCSUProfile'
    }

    $storeNumber = 'Houston'
    $workerPersonnelNumber = '000160'

    $config = Get-D365ODataConfig -Name $ComputerName
    if ($null -eq $config) {
        Add-D365ODataConfig -Name $ComputerName -Tenant $tenantId -Url $D365FOBaseUrl
    }
    Set-D365ActiveODataConfig -Name $ComputerName
    
    Get-D365ODataTokenInteractive | Set-D365ODataTokenInSession

    Write-Host "[5.1] Create Microsoft Entra ID application for CSU"

    $aadClient = Get-D365ODataEntityData `
        -EntityName SysAADClients `
        -ODataQuery "`$filter=AADClientId eq '$($CsuAppId)'" -ErrorAction Stop

    if ($null -eq $aadClient) {
        $Payload = @{
            '@odata.type' = "Microsoft.Dynamics.DataEntities.SysAADClient"
            "AADClientId" = "$($CsuAppId)"
            "Name" = "Cloud Scale Unit at $($ComputerName)"
            "UserId" = "RetailServiceAccount"
        } | ConvertTo-Json -Depth 4

        Import-D365ODataEntity -EntityName "SysAADClients" -Payload $Payload -ErrorAction Stop

        $syncRequired = $True
    } else {
        Write-Host "    Microsoft Entra ID application record already exists, skipping..." -ForegroundColor Yellow
    }

    Write-Host "[5.2] Create 'Channel database' record named $channelDatabaseName"
    
    $channel = Get-D365ODataEntityData `
        -EntityName RetailStores `
        -ODataQuery "`$filter=StoreNumber eq '$storeNumber'" -ErrorAction Stop

    if ($null -eq $channel) {
        throw "    Cannot find 'Channel' record with '$storeNumber' name..."
    }
    
    $channelDatabase = Get-D365ODataEntityData `
        -EntityName RetailConnDatabaseProfiles `
        -ODataQuery "`$filter=Name eq '$channelDatabaseName'" -ErrorAction Stop

    if ($null -eq $channelDatabase) {

        $Payload = @{
            '@odata.type' = "Microsoft.Dynamics.DataEntities.RetailConnDatabaseProfile"
            "Name" = "$($channelDatabaseName)"
            "RetailCDXDataGroup_Name" = "Default"
        } | ConvertTo-Json -Depth 4

        Import-D365ODataEntity -EntityName RetailConnDatabaseProfiles -Payload $Payload -ErrorAction Stop

        $syncRequired = $True
    } else {
        Write-Host "    'Channel database' record already exists, skipping..." -ForegroundColor Yellow
    }

    Write-Host "[5.3] Link '$storeNumber' store to newly created 'Channel database'"
    
    $channelDatabaseCDXStore = Get-D365ODataEntityData `
        -EntityName CDXDataStoreChannels `
        -ODataQuery "`$filter=ChannelId eq '$($channel.RetailChannelId)' and ChannelDatabaseId eq '$channelDatabaseName'" -ErrorAction Stop

    if ($null -eq $channelDatabaseCDXStore) {
        $Payload = @{
            '@odata.type' = "Microsoft.Dynamics.DataEntities.CDXDataStoreChannel"
            "ChannelId" = "$($channel.RetailChannelId)"
            "ChannelDatabaseId" = "$channelDatabaseName"
        } | ConvertTo-Json -Depth 4

        Import-D365ODataEntity -EntityName CDXDataStoreChannels -Payload $Payload -ErrorAction Stop

        $syncRequired = $True
    } else {
        Write-Host "    Store is already linked, skipping..." -ForegroundColor Yellow
    }

    Write-Host "[5.4] Update Default channel database to Legacy group if exists"

    $channelDatabaseDefault = Get-D365ODataEntityData `
        -EntityName RetailConnDatabaseProfiles `
        -ODataQuery "`$filter=Name eq 'Default'" -ErrorAction Stop

    if (($null -ne $channelDatabaseDefault) -and ($channelDatabaseDefault.RetailCDXDataGroup_Name -ne 'Legacy')) {

        Write-Host "[5.4.1] Create 'Channel database group' record named 'Legacy' if 'Default' database exists"
        
        $channelDatabaseGroupLegacy = Get-D365ODataEntityData `
            -EntityName RetailCdxDataGroups `
            -ODataQuery "`$filter=Name eq 'Legacy'" -ErrorAction Stop

        if ($null -eq $channelDatabaseGroupLegacy) {

            $Payload = @{
                '@odata.type' = "Microsoft.Dynamics.DataEntities.RetailCdxDataGroup"
                "Name" = "Legacy"
                "Description" = "Legacy data group"
                "RetailConnChannelSchema_SchemaName" = "AX7"
            } | ConvertTo-Json -Depth 4

            Import-D365ODataEntity -EntityName RetailCdxDataGroups -Payload $Payload -ErrorAction Stop

            $syncRequired = $True

        } else {
            Write-Host "    'Channel database group' record already exists, skipping..." -ForegroundColor Yellow
        }

        $Payload = @{
            "RetailCDXDataGroup_Name" = "Legacy"
        } | ConvertTo-Json -Depth 4

        Update-D365ODataEntity -EntityName RetailConnDatabaseProfiles `
            -Key "Name='Default'" `
            -Payload $Payload -ErrorAction Stop

        Write-Host "    Default 'Channel database' record has been updated to Legacy group" -ForegroundColor Green

        $syncRequired = $True
    } else {
        Write-Host "    Default 'Channel database' doesn't exist or record already updated to Legacy group, skipping..." -ForegroundColor Yellow
    }

    Write-Host "[5.5] Create 'Channel profile' record named '$retailChannelProfileName'"
    
    $retailChannelProfile = Get-D365ODataEntityData `
        -EntityName RetailChannelProfiles `
        -ODataQuery "`$filter=Name eq '$retailChannelProfileName'" -ErrorAction Stop

    if ($null -eq $retailChannelProfile) {

        $Payload = @{
            '@odata.type' = "Microsoft.Dynamics.DataEntities.RetailChannelProfile"
            "Name" = "$retailChannelProfileName"
            "ChannelProfileType" = "RetailServer"
        } | ConvertTo-Json -Depth 4

        Import-D365ODataEntity -EntityName RetailChannelProfiles -Payload $Payload -ErrorAction Stop

        $syncRequired = $True

        # Add value for Cloud POS Url to channel profile
        $Payload = @{
            '@odata.type' = "Microsoft.Dynamics.DataEntities.RetailChannelProfileProperty"
            "RetailChannelProfile_Name" = "$retailChannelProfileName"
            "Key" = "7" # Cloud POS URL
            "Value" = "$CposUrl"
        } | ConvertTo-Json -Depth 4

        Import-D365ODataEntity -EntityName RetailChannelProfileProperties -Payload $Payload -ErrorAction Stop

        # Add value for Retail Server Url to channel profile
        $Payload = @{
            '@odata.type' = "Microsoft.Dynamics.DataEntities.RetailChannelProfileProperty"
            "RetailChannelProfile_Name" = "$retailChannelProfileName"
            "Key" = "1" # Retail Server URL
            "Value" = "$CsuUrl"
        } | ConvertTo-Json -Depth 4

        Import-D365ODataEntity -EntityName RetailChannelProfileProperties -Payload $Payload -ErrorAction Stop

        # Receive standard value for Media Server Base URL
        $mediaServerUrl = (Get-D365ODataEntityData `
            -EntityName RetailChannelProfileProperties `
            -ODataQuery "`$filter=RetailChannelProfile_Name eq 'Default' and Key eq 2" -ErrorAction Stop).Value

        # Add value for Media Server Base URL to channel profile if exists
        if (($null -ne $mediaServerUrl) -and ($mediaServerUrl -ne "https://MIGRATION_VALUE")) {
            $Payload = @{
                '@odata.type' = "Microsoft.Dynamics.DataEntities.RetailChannelProfileProperty"
                "RetailChannelProfile_Name" = "$retailChannelProfileName"
                "Key" = "2" # Media Server Base URL
                "Value" = "$mediaServerUrl"
            } | ConvertTo-Json -Depth 4

            Import-D365ODataEntity -EntityName RetailChannelProfileProperties -Payload $Payload -ErrorAction Stop

            $syncRequired = $True
        }

    } else {
        Write-Host "    'Channel profile' record already exists, skipping..." -ForegroundColor Yellow
    }

    Write-Host "[5.6] Update '$storeNumber' store record with Live database and Channel profile value"

    if (($channel.LiveDatabaseConnectionProfileName -ne $channelDatabaseName) -or ($channel.ChannelProfileName -ne $retailChannelProfileName)) {
        $Payload = @{
            "ClosingMethod" = "PosBatch"
            "LiveDatabaseConnectionProfileName" = "$channelDatabaseName"
            "ChannelProfileName" = "$retailChannelProfileName"
        } | ConvertTo-Json -Depth 4

        Update-D365ODataEntity -EntityName RetailStores `
            -Key "RetailChannelId='$($channel.RetailChannelId)'" `
            -Payload $Payload -ErrorAction Stop

        Write-Host "    Store '$storeNumber' record has been updated" -ForegroundColor Green

        $syncRequired = $True
    } else {
        Write-Host "    Store '$storeNumber' record already updated, skipping..." -ForegroundColor Yellow
    }

    Write-Host "[5.7] Create new Hardware Station for '$storeNumber' store record"

    $channelHardwareStation = Get-D365ODataEntityData `
        -EntityName RetailStoreHardwareStations `
        -ODataQuery "`$filter=StoreNumber eq '$storeNumber' and HostName eq '$ComputerName'" -ErrorAction Stop

    if ($null -eq $channelHardwareStation) {
        $Payload = @{
            '@odata.type' = "Microsoft.Dynamics.DataEntities.RetailStoreHardwareStation"
            "StoreNumber" = "$storeNumber"
            "HostName" = "$ComputerName"
            "Description" = "$ComputerName"
            "Port" = "$HsPort"
            "HardwareStationType" = "Shared"
            "HardwareProfileId" = "Virtual"
        } | ConvertTo-Json -Depth 4

        Import-D365ODataEntity -EntityName RetailStoreHardwareStations -Payload $Payload -ErrorAction Stop

        $syncRequired = $True
    } else {
        Write-Host "    'Channel hardware station' record already exists, skipping..." -ForegroundColor Yellow
    }

    Write-Host "[5.8] Update 'Worker' with number '$workerPersonnelNumber' with external identity of the logged user"

    $workerIdentity = Get-D365ODataEntityData `
        -EntityName RetailStaffs `
        -ODataQuery "`$filter=PersonnelNumber eq '$workerPersonnelNumber'" -ErrorAction Stop

    $signedInUser = Invoke-MgGraphRequest -Method GET "/v1.0/me?`$select=mailNickname,userPrincipalName,id" -ErrorAction Stop

    if ($null -eq $workerIdentity) {
        throw "Cannot find 'Worker' record with personnel number '$workerPersonnelNumber'"
    } elseif ($workerIdentity.ExternalSubIdentifier -eq $signedInUser.id) {
        Write-Host "    'Worker' record already updated, skipping..." -ForegroundColor Yellow
    } else {
        $Payload = @{
            "ExternalIdentityAlias" = "$($signedInUser.mailNickname)"
            "ExternalName" = "$($signedInUser.userPrincipalName)"
            "ExternalSubIdentifier" = "$($signedInUser.id)"
            "ExternalIdentifier" = "$($tenantId)"
        } | ConvertTo-Json -Depth 4

        Update-D365ODataEntity -EntityName RetailStaffs `
            -Key "PersonnelNumber='$workerPersonnelNumber'" `
            -Payload $Payload -ErrorAction Stop

        Write-Host "    'Worker' record has been updated" -ForegroundColor Green

        $syncRequired = $True
    }

    Write-Host "[5.9] Create 'External Identity' relying party in Commerce Shared Parameters"

    $Payload = @{
        "identityProviderContract" = @{
            "RetailIdentityProviderName" = "Microsoft Entra ID"
            "RetailIdentityProviderType" = "Aad"
            "RetailIdentityProviderUrl" = "https://sts.windows.net/$($tenantId)/"
            "RetailRelyingParties" = @(
                @{
                    "RetailClientId" = "$($CposAppId)"
                    "RetailClientName" = "Store Commerce for Web"
                    "RetailRelyingPartyType" = "Public"
                    "RetailRelyingUserPartyType" = "Worker"
                    "RetailServerResourceIds" = @(
                        @{
                            "RetailServerResourceId" = "api://$($CsuAppId)"
                            "RetailServerResourceName" = "Cloud Scale Unit"
                        }
                    )
                }
            )
        }
    } | ConvertTo-Json -Depth 5

    # Could be run several times as implementation uses findOrCreate method
    Invoke-D365RestEndpoint -ServiceName "RetailCDXRealTimeService/RetailCDXChannelService/AddRetailServerIdentity" `
        -Payload $Payload `
        -ErrorAction Stop

    if ($syncRequired -eq $True) {

        if (-not($IsUDE)) {
            Write-Host "[5.10] Start Batch Service to run the CDX job"
            net start DynamicsAxBatch
        }

        Write-Host "[5.11] Create a batch job to run '9999' schedule job for $($channelDatabaseName)"
        $Payload = @{
            "scheduleName" = "9999"
            "dataStoreName" = "$($channelDatabaseName)"
        } | ConvertTo-Json -Depth 4

        Invoke-D365RestEndpoint -ServiceName "RetailCDXRealTimeService/RetailCDXChannelService/RunCDXScheduleFullSync" -Payload $Payload
    }
}

function Invoke-CommerceInstallers {
    param (
        [ValidateNotNullOrEmpty()]
        [string]$ComputerName = $env:computerName,
        [Parameter(Mandatory)]
        [object]$CsuApp,
        [Parameter(Mandatory)]
        [string]$CsuUrl,
        [Parameter(Mandatory)]
        [string]$CsuPort,
        [Parameter(Mandatory)]
        [object]$CsuCert,
        [Parameter(Mandatory)]
        [object]$CposApp,
        [Parameter(Mandatory)]
        [object]$HsCert,
        [Parameter(Mandatory)]
        [string]$HsPort
    )

    $install = Prompt "Do you want to proceed with installers? (Y/N)"

    if ($install -eq $True) {

        Write-Host "`n[6.1] Running CSU installer`n" -ForegroundColor yellow

        do {
            # Install CSU
            ./CommerceStoreScaleUnitSetup.exe install `
              --port $CsuPort `
              --SSLCertThumbprint $CsuCert.Thumbprint `
              --RetailServerCertThumbprint $CsuCert.Thumbprint `
              --AsyncClientCertThumbprint $CsuCert.Thumbprint `
              --AsyncClientAADClientID $CsuApp.AppId `
              --RetailServerAADClientID $CsuApp.AppId `
              --CPOSAADClientID $CposApp.AppId `
              --RetailServerAADResourceID "api://$($CsuApp.AppId)" `
              --Config "c:\temp\StoreSystemSetup.xml" `
              --SkipSChannelCheck --trustSqlservercertificate

            if ($LASTEXITCODE -ne 0) {
                Write-Error "`nCSU installer failed. Please check the logs above.`n"
                $retry = Read-Host "Do you want to retry? (Y/N)"
            } else {
                $retry = "N";
            }
        } while ($retry -eq "Y")

        Write-Host "`n[6.2] Running HS installer`n" -ForegroundColor yellow

        do {
            # Install HS
            ./CommerceHardwareStationSetup.exe install `
              --CsuUrl $CsuUrl `
              --CertThumbprint $HsCert.Thumbprint `
              --port $HsPort

            if ($LASTEXITCODE -ne 0) {
                Write-Error "`nHS installer failed. Please check the logs above.`n"
                $retry = Read-Host "Do you want to retry? (Y/N)"
            } else {
                $retry = "N";
            }
        } while ($retry -eq "Y")

        Write-Host "`n[6.3] Running Store Commerce App installer`n" -ForegroundColor yellow

        do {
            # Install Store Commerce App for Windows
            ./StoreCommerce.Installer.exe install --enablewebviewdevtools

            if ($LASTEXITCODE -ne 0) {
                Write-Error "`nStore Commerce App for Windows installer failed. Please check the logs above.`n"
                $retry = Read-Host "Do you want to retry? (Y/N)"
            } else {
                $retry = "N";
            }
        } while ($retry -eq "Y")
    }
}

Invoke-D365CommerceDevWizard