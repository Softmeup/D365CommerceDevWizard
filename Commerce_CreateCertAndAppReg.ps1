#Requires -RunAsAdministrator
#Requires -Modules Microsoft.Graph.Authentication,Microsoft.Graph.Applications

function D365-CommerceDevWizardRun {

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
        cd 'C:\Temp'
        
        $CsuUrl = "https://$($ComputerName):$($CsuPort)/RetailServer/Commerce"
        $CsuHealthCheckUrl = "https://$($ComputerName):$($CsuPort)/RetailServer/healthcheck?testname=ping"
        $CposUrl = "https://$($ComputerName):$($CsuPort)/POS"
        $HsPingUrl = "https://$($ComputerName):$($HsPort)/HardwareStation/ping"
        
        Write-Header

        do
        {
            $answer = Read-Host "Are you ready to start? (Y/N)"
            if ($answer -eq "N") {
                return
            }
        } while ($answer -ne "Y")

        Connect-MgGraph -Scopes "Application.ReadWrite.All","User.Read" -NoWelcome -ErrorAction Stop

        Write-Host "`nAuthentification succeed. Press any key to start process...`n"
        $k = [Console]::ReadKey($true)

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

        # Manual steps instructions
        Write-ManualSteps -ComputerName $ComputerName `
            -HsPort $HsPort `
            -CsuAppId $CsuApp.AppId `
            -CposAppId $CposApp.AppId `
            -CsuUrl $CsuUrl `
            -CposUrl $CposUrl

        # Step 5
        Run-Installers -ComputerName $ComputerName `
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

    }
    catch [Exception]
    {
        Write-Host "Something threw an exception or used Write-Error"
        Write-Error $_
    }
    finally
    {
        Write-Host "Press any key to exit..."

        $k = [Console]::ReadKey($true)

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
    5. Run installers of CSU, Store Commerce App for Windows and HS (you need to download them manually from LCS)`n"

}

function New-CSUCertificate {

    param (
        [ValidateNotNullOrEmpty()]
        [string]$ComputerName = $env:computerName,
        [ValidateNotNullOrEmpty()]
        [string]$CsuCertFriendlyName = "CSU2024"
    )

    Write-Host "[1] Generate self-signed certificate for CSU named $CsuCertFriendlyName"

    $CsuCert = dir cert: -Recurse | Where-Object { $_.FriendlyName -eq $CsuCertFriendlyName } | Select-Object -first 1
    
    if ($CsuCert -eq $null) {
        
        $CsuCert = New-SelfSignedCertificate -FriendlyName $CsuCertFriendlyName `
            -DnsName $ComputerName `
            -CertStoreLocation Cert:\LocalMachine\My `
            -KeyUsage DigitalSignature,DataEncipherment,KeyEncipherment

        Write-Host "    Install certificate into root storage"
        $mypwd = ConvertTo-SecureString -String '1234' -Force -AsPlainText
        Export-PfxCertificate -Cert $CsuCert -FilePath "$env:temp\csu2024.pfx" -Password $mypwd
        Import-PfxCertificate -CertStoreLocation cert:\LocalMachine\Root -FilePath "$env:temp\csu2024.pfx" -Password $mypwd
    } else {
        Write-Host "    Self-signed certificate for CSU already exists, skipping..."
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

    $HsCert = dir cert: -Recurse | Where-Object { $_.FriendlyName -eq $HsCertFriendlyName } | Select-Object -first 1
    
    if ($HsCert -eq $null) {

        $HsCert = New-SelfSignedCertificate -FriendlyName $HsCertFriendlyName `
            -DnsName $ComputerName `
            -CertStoreLocation Cert:\LocalMachine\My `
            -KeyUsage DigitalSignature,DataEncipherment,KeyEncipherment

        Write-Host "    Install certificate into root storage"
        $mypwd = ConvertTo-SecureString -String '1234' -Force -AsPlainText
        Export-PfxCertificate -Cert $HsCert -FilePath "$env:temp\hs2024.pfx" -Password $mypwd
        Import-PfxCertificate -CertStoreLocation cert:\LocalMachine\Root -FilePath "$env:temp\hs2024.pfx" -Password $mypwd

    } else {
        Write-Host "    Self-signed certificate for HS already exists, skipping..."
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
    if ($CsuApp -eq $null) {

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
        $csuSP = New-MgServicePrincipal -AppId $CsuApp.AppId -ErrorAction Stop
    } else {
        Write-Host "    CSU App Registration already exists, skipping..."
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
    if ($CposApp -eq $null) {

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

        if ($scopeId_LegacyAccessFull -eq $null) {
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
        $CposSP = New-MgServicePrincipal -AppId $CposApp.AppId -ErrorAction Stop

    } else {
        Write-Host "    CSU App Registration already exists, skipping..."
    }

    Write-Host "    CPOS AppId: $($CposApp.AppId)"

    return $CposApp
}

function Write-ManualSteps {
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
        [string]$CposUrl
    )
    
    $signedInUser = Invoke-MgGraphRequest -Method GET "/v1.0/me?`$select=mailNickname,userPrincipalName,id" -ErrorAction Stop

    Write-Host "

Manual steps are required in HQ (D365FO):

Check more here: https://learn.microsoft.com/en-us/dynamics365/commerce/dev-itpro/install-csu-dev-env

    1. Open 'Microsoft Entra ID Applications' form, create new record with CSU appId and RetailServiceAccount
      Client Id = $CsuAppId
      Name = <Cloud Scale Unit or any you like>
      User ID = RetailServiceAccount

    2. Open 'Channel database' form, create new record (e.g. DevSealedCSU)
      Add 'Houston channel' to the 'Retail channel' list
      Select Download > Configuration file and save it under C:\temp\StoreSystemSetup.xml

    3. Open 'Channel database group' form, create record named 'Legacy'
      Open 'Default' channel database record and set it's group to 'Legacy'

    4. Open 'Channel profiles' form, create new record with following values
      Retail Server URL = $CsuUrl
      Cloud POS URL = $CposUrl
      Media Server URL = <Copy from existing profile>

    5. Open 'All stores' form and find 'Houston' store, replace 'Live Channel Database' and 'Channel Profile' with newly created records
      Expand 'Hardware stations' fast tab and create new record with following values:
        Hardware station type = Shared
        Host name = $ComputerName
        Port = $HsPort
        Hardware profile = Virtual

    6. Open 'Workers' form and find 'Alexander Eggerer' record, on action pane's Commerce tab run 'Clear external identity'
      On Commerce tab of the form, set the following values of 'External identities' group
        Alias = $($signedInUser.mailNickname)
        UPN = $($signedInUser.userPrincipalName)
        External sub identifier = $($signedInUser.id)

    7. Open 'Commerce shared parameters' form, select 'Identity Providers' tab
      Select record that starts 'https://sts.windows.net/*'
      Create new 'Relying Parties' record and provide following values:
        ClientId = $CposAppId
        Type = Public
        UserType = Worker
      Add new 'Server Resource Ids' record and provide following values (don't forget to save first, so New button will be active):
        ServerResourceId = api://$($CsuAppId)

    8. Open 'Distribution schedule' form, select '9999' record and click 'Run now'

    9. Download and install http://monroecs.com/oposccos_current.htm for Hardware Station

    10. Download and install two .NET 6.0 runtimes: https://dotnet.microsoft.com/en-us/download/dotnet/6.0
      In the 'ASP.Net Core Runtime 6.0.X section', select the Hosting Bundle installer for Windows.
      In the '.NET Desktop Runtime 6.0.X' section, select the x64 installer for Windows.

    11. Open LCS -> 'Asset Library', click 'Import' button and pick sealed installers (check the version):
      Download following installers:
        10.0.XX - Commerce Peripheral Simulator
        10.0.XX - Hardware Station (SEALED)
        10.0.XX - Commerce Scale Unit (SEALED)
        10.0.XX - Store Commerce" -ForegroundColor yellow

}

function Run-Installers {
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

    $answer = Read-Host "Do you want to proceed with installers? (Y/N)"

    if ($answer -eq "Y") {

        Write-Host "`n[5.1] Running CSU installer`n" -ForegroundColor yellow

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

        Write-Host "`n[5.2] Running HS installer`n" -ForegroundColor yellow

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

        Write-Host "`n[5.3] Running Store Commerce App installer`n" -ForegroundColor yellow

        do {
            # Install Store Commerce App for Windows
            ./StoreCommerce.Installer.exe install

            if ($LASTEXITCODE -ne 0) {
                Write-Error "`nStore Commerce App for Windows installer failed. Please check the logs above.`n"
                $retry = Read-Host "Do you want to retry? (Y/N)"
            } else {
                $retry = "N";
            }
        } while ($retry -eq "Y")
    }
}

D365-CommerceDevWizardRun