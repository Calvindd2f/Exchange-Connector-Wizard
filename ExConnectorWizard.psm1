# Import the functions from the script file
#########################
#        Menu-ing       #
#########################
# Main
function Show-Menu {
    Write-Output "                 __    __       __          __ __                      ";
    Write-Output ">| ||  | >> |  /|     |  \ />> |<<>> | || ||  |<<>>|<<>> |<< |   |>>>> ";
    Write-Output "||\||  ||  ||<< |<<-<-|<< <|  || |  ||\||\||<<|    | |  ||>>|| < || /  ";
    Write-Output "|| | \/  << |  \|__   |__/ \<< |__<< | || ||__|__  |  << |  \|/ \||/<< ";
    Write-Output ''
    $menu = @"
=======================================
   Welcome to ExoConnector Wizard.

   Calvin Bergin 24th December
   Calvindd2f
=======================================
0. Authentication
1. SMTP Submission
2. Direct Send
3. SMTP Relay
4. SPF Generator
5. Interactive Wizard
6. How to documenation... 
=======================================
"@

    Write-Host $menu -ForegroundColor Green
}

function Show-Intro {
    $logo = @"
    ExoConnector Wizard

    By Calvindd2f - CalvinBergin 24.12.2023
"@

    $introLines = @(
        " "
        " "
        "Welcome to the ExoConnector Wizard!",
        "Initializing...",
        "Loading resources...",
        "Configuring settings...",
        "Preparing for launch..."
    )

    Clear-Host

    foreach ($line in $introLines) {
        Write-Host $line -ForegroundColor Yellow
        Start-Sleep -Seconds 1
    }

    Clear-Host

    foreach ($char in $logo.ToCharArray()) {
        Write-Host -NoNewline $char -ForegroundColor Yellow
        #Start-Sleep -Milliseconds 50
    }

    Write-Host "`n`n`n"
}
function Invoke-ExoConnectorWiz {
    Show-Menu

    $userChoice = Read-Host "Select an option (0-6)"

    switch ($userChoice) {
        '0' { Start-Authentication }
        '1' { Invoke-SMTPSubmissionFunction }
        '2' { Invoke-DirectSendFunction }
        '3' { Invoke-SMTPRelayFunction }
        '4' { Invoke-SPFGeneratorFunction }
        '5' { Start-ExoWizard }
        '6' { Show-DocumentationMenu }
        default { 
            Write-Host "Invalid choice." -ForegroundColor Red 
            Invoke-ExoConnectorWiz
        }
    }
}


#########################
#    Authentication     #
#########################
function Start-Authentication {
    try {
        Import-Module -Name Microsoft.Graph
        Import-Module -Name ExchangeOnlineManagement
    }
    catch {
        Write-Host "Failed to import the required modules. Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Trying to install modules ; then import." -ForegroundColor Yellow
        Install-Module -Name Microsoft.Graph -Scope CurrentUser -AllowClobber -Force
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -AllowClobber -Force
        Import-Module -Name Microsoft.Graph
        Import-Module -Name ExchangeOnlineManagement
    }
    finally {
        # Authenticate to Exchange Online
        Connect-ExchangeOnline -ShowBanner:$false
        # Authenticate to Microsoft Graph
        Connect-MgGraph -Scopes 'User.Read.All', 'Group.Read.All', 'Directory.Read.All'
        Invoke-ExoConnectorWiz
        Write-Host "Authenticated to Exchange and Graph" -ForegroundColor Green
        }
    }




#Invoke-SPFGenerator Functions
function Invoke-SPFGeneratorFunction {
    param (
        [Parameter(Mandatory=$false)]
        [string]$Domain,
        [bool]$mx,
        [bool]$a,
        [IPAddress]$ip4,
        [IPAddress]$ip6,
        [string]$aDelegate,
        [string]$include,
        [string]$Policy
    )
    $Domain=Read-Host "Enter Domain: "
    $mx=Read-Host "Include MX Record? (Y/N)"
    $a=Read-Host "Include A Record? (Y/N)"
    $ip4=Read-Host "Enter IPv4 Address: " -ErrorAction SilentlyContinue
    #$ip6=Read-Host "Enter IPv6 Address [can be empty]: " -ErrorAction SilentlyContinue
    $aDelegate=Read-Host "Enter A Record Delegate [can be empty] : "
    $include=Read-Host "Enter Include Record [can be empty] : "
    $Policy=Read-Host "Enter Policy [can be empty, will= default to softfail] : "

    if([string]::IsNullOrEmpty($Domain)) { 
        throw "Domain are required parameters."
    }

    if([string]::IsNullOrEmpty($include)) { 
        $include = "_spf.microsoft.com"
    }
    if([string]::IsNullOrEmpty($aDelegate)) {
        $i=Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/domains" |ConvertTo-Json
        $i=$i|ConvertFrom-json
        $delegateDomains = [PSCustomObject]@{
            Domain1=$i.value.id[0]
            Domain2=$i.value.id[1]
            Domain3=$i.value.id[2]
            Domain4=$i.value.id[3] 
            Domain5=$i.value.id[4]
        }
        $aDelegate=$delegateDomains.Domain1;
    }
    
    if([string]::IsNullOrEmpty($mx)) {
        $mx=$true;
    }
    if([string]::IsNullOrEmpty($a)) {
        $a=$true;
    }
    if([string]::IsNullOrEmpty($Policy)) {
        $Policy="~all";
    }
    #if([IPAddress]::IsNullOrEmpty($ip6)) {
    #    $ip6=$false;
    #}
    #if([IPAddress]::IsNullOrEmpty($ip4)) {
    #    $ip4=$false;
    #}
    # Output the domain and its actual SPF record
    $currentSPF = ((Resolve-DnsName -Name $Domain -Type TXT) | Where-Object { $_.Strings -match "v=spf1" }).Strings
    Write-Host "Current SPF is: $currentSPF"
    # Generate the SPF record
    Write-Host "Generated SPF:"
    $spfRecord = "v=spf1"
    if ($mx) { $spfRecord += " mx" }
    if ($a) { $spfRecord += " a" }
    if ($ip4) { $spfRecord += " ip4:$ip4" }
    #if ($ip6) { $spfRecord += " ip6:$ip6" }
    if ($aDelegate) { $spfRecord += " a:$aDelegate"}
    if ($include) { $spfRecord += " include:$include" }
    $spfRecord += " $Policy"
    Write-Host $spfRecord

    Write-Host ""
    Start-Sleep 2
    Write-Host "Please make the changes to the TXT record at your domain host" -ForegroundColor Cyan
    Write-Host ""
    Start-Sleep 1
    $Menu=Read-Host "Continue (Y/N)?"
    if ($menu){
        Invoke-ExoConnectorWiz
    } else {
        Write-Host "Exiting..." -ForegroundColor Blue
        exit
    }
}

# New-SMTPConnector Functions
function New-SMTPConnector {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConnectorName,
        [Parameter(Mandatory=$true)]
        [string]$ServerName,
        [Parameter(Mandatory=$true)]
        [string]$AddressSpaces,
        [Parameter(Mandatory=$true)]
        [string]$RemoteIPRanges,
        [Parameter(Mandatory=$true)]
        [SecureString]$AuthCredentials,
        [Parameter(Mandatory=$true)]
        [string]$DomainSecure,
        [Parameter(Mandatory=$true)]
        [string]$Usage
    )

    # Define a hash table to hold the properties for the new SMTP connector
    $SMTPConnectorProps = @{
        Name                   = $ConnectorName
        Usage                 = $Usage
        AuthMechanism          = 'Credssp'
        Bindings               = '0.0.0.0:25:25'
        BareLinefeedRejection = $true
        BinaryMimeEnabled      = $true
        ChunkingEnabled        = $true
        Comment                = 'Created by New-SMTPConnector script'
        DeliveryStatusNotificationEnabled = $true
        DomainSecureEnabled    = $DomainSecure
        Enabled                = $true
        EnhancedStatusCodesEnabled = $true
        LongAddressesEnabled   = $true
        MaxMessageSize         = 25MB
        MaxAcknowledgementDelay = '00:00:00'
        MaxDeliveryDelay       = '1.00:00:00'
        MaxHeaderSize          = 64KB
        MessageRateLimit       = 'Unlimited'
        MessageRateSource      = 'Remote'
        OrarEnabled            = $true
        PermissionGroups       = 'Anonymous Users'
        ProtocolLoggingLevel   = 'Verbose'
        RequireTLS             = $true
        SmartHosts             = $ServerName
        SmtpUtf8Enabled        = $true
        SuppressXAnonymousTls = $true
        TlsDomainCapabilities = 'DomainAuthNoLogin', 'DomainAuthLogin'
        TlsCertificateName     = '*.example.com'
        TlsDomainName          = 'example.com'
        TlsDomainNames         = @('example.com', '*.example.com')
        TlsLoggingLevel        = 'None'
        TlsReceiptsEnabled     = $true
        UseExternalDNSServersEnabled = $true
        ValidAuthMechanisms    = 'Ntlm', 'Credssp'
    }

    # Create the new SMTP connector using the defined properties
    try {
        $SMTPConnector = New-SendConnector @SMTPConnectorProps ; $SMTPConnector 
        Write-Host "SMTP Connector '$ConnectorName' created successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to create SMTP Connector '$ConnectorName'. Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    Invoke-ExoConnectorWiz
}

# New-DirectSendConector Functions
function Invoke-DirectSendFunction {
    [CmdletBinding()]
    param (

    )
    
    begin {
        
    }
    
    process {
        Write-Host "You have selected : Direct Send"-ForegroundColor Cyan
        Write-Host " "
        $Domain = Read-Host -Prompt "What domain are you sending from:"
        Write-Host ""
        Write-Host "Verifying $Domain..."

        if([string]::IsNullOrEmpty($Domain)) { 
            throw "Domain is required parameters."
            $Domain = Read-Host -Prompt "What domain are you sending from:"
        }

        try {
            $mxRecord = (Resolve-DnsName -Name $Domain -Type MX).NameExchange
            Write-Host "Your MX record is $mxRecord" -ForegroundColor DarkMagenta

            $currentSPF = ((Resolve-DnsName -Name $Domain -Type TXT) | Where-Object { $_.Strings -match "v=spf1" }).Strings
            Write-Host "Current SPF is: $currentSPF" -ForegroundColor DarkMagenta

            Write-Host "Defining WAN IP as the IP used in connector - because fuck you." -ForegroundColor Blue
            $WANip = (Invoke-RestMethod -Uri "https://ipinfo.io/ip").Trim()
            Write-Host "WAN IP: $WANip" -ForegroundColor Blue
            Write-Host ""

            Write-Host "Validating if the SPF record contains an IP address after the prefix 'ip4:'" -ForegroundColor DarkMagenta
            $IP_exists = $currentSPF -match "ip4:$WANip"

            if ($IP_exists) {
                Write-Host "$WANip is already in the SPF Record" -ForegroundColor DarkMagenta
                Write-Host "Proceeding to creating the connector" -ForegroundColor DarkMagenta
                New-SendConnector -Internal -DNSRoutingEnabled:$true -Enabled:$true -Name "Direct Send Connector" -Comment "Direct Send Connector Created by $env:username on $(Get-Date)"
            } else {
                Write-Host "$WANip is not in the record."
                $confirmation = Read-Host -Prompt "Do you want to generate it? [y]/n"

                if ($confirmation -eq 'y') {
                    $newSPF = "v=spf1 ip4:$WANip include:_spf.microsoft.com ~all"
                    Write-Host "Adding $WANIp to the SPF Record"
                    Write-Host "This is the SPF record you need to add: $newSPF"

                    # Set-DkimSigningConfig -Identity Default -SpfRecord $newSPF

                    while ($currentSPF -ne $newSPF) {
                        Start-Sleep -Seconds 10
                        Write-Host "Current SPF record is not equal to $newSPF. Checking again in 10 seconds."
                        $currentSPF = ((Resolve-DnsName -Name $Domain -Type TXT) | Where-Object { $_.Strings -match "v=spf1" }).Strings
                    }

                    Write-Host "Changes Propagated - creating connector."
                    New-SendConnector -Internal -DNSRoutingEnabled:$true -Enabled:$true -Name "Direct Send Connector" -Comment "Direct Send Connector Created by $env:username on $(Get-Date)"
                    Write-Host "Connector Created! Please test it out."
                }
            }
        } catch {
            Write-Host "Failed to verify the domain. Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    end {
        Invoke-ExoConnectorWiz
    }
}

# Invoke-SMTPSubmission Functions
function Invoke-SMTPSubmissionFunction {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        
    }
    
    process {
        Write-Host "You have selected : SMTP Client Submission"-ForegroundColor Cyan
        Write-Host " "
        # Prompt for domain input
        Write-Host "What domain are you sending from:"
        $domain = Read-Host -Prompt "Domain : "
        Start-Sleep 2
        # Validate domain presence
        Write-Host "Validating $domain is present in your tenant"
        Write-Progress -Activity $domainExists
        # Verify that $domain exists in the tenant
        if(!(Get-Module -Name @("Microsoft.Graph","ExchangeOnlineManagement"))){
            Start-Authentication
        }
        $tenantDomains = Get-MgDomain
        $domainExists = $tenantDomains.Name -contains $domain

        if (-not $domainExists) {
            Write-Host "The specified domain does not exist in the tenant."
            Write-Host "Option 1. Use the domain name that you are currently authenticated with."
            $Option1 = {
                Write-Host "Using authenticated domain of current user"
                Start-Sleep 2
                $currentUser = Get-MgMe
                $domain = $currentUser.UserPrincipalName.Split("@")[1]
                Write-Host "Domain = $domain"
            }
            Write-Host "Option 2. Read host again for domain input."
            $Option2 = {
                Write-Host "Read host again for domain input"
                Start-Sleep 2
                $domain = Read-Host -Prompt "Domain :"
                Write-Host "Domain = $domain"
            }
            Write-Host "Option 3. Return to menu."
            $Option3 = {
                Invoke-ExoConnectorWiz
            }
            $SMTPSubChoice = Read-Host "Select an option (1-3)"
            switch ($SMTPSubChoice) {
                '1' { $Option1 }
                '2' { $Option2 }
                '3' { $Option3 }
                default { Write-Host "Invalid choice." -ForegroundColor Red }
            }
        }
        else {
            Write-Host "ENDOF: Validate domain presence" -ForegroundColor DarkGreen
        }

        # Entra License Check
        Write-Host "Begin: Entra License Check" -ForegroundColor DarkGreen
        $licenses = @{
            'Microsoft Entra Basic'               = 'AAD_BASIC'
            'Microsoft Entra ID P1'               = 'AAD_PREMIUM'
            'Microsoft Entra ID P1 for faculty'   = 'AAD_PREMIUM_FACULTY'
            'Microsoft Entra ID P1_USGOV_GCCHIGH' = 'AAD_PREMIUM_USGOV_GCCHIGH'
            'Microsoft Entra ID P2'               = 'AAD_PREMIUM_P2'
        }

        # Check if the Entra AD license is AAD_BASIC
        $entraLicense = Get-MgUserLicenseDetail -UserId $currentUser.Id
        $isBasicLicense = $entraLicense.Licenses -contains $licenses['Microsoft Entra Basic']

        if ($isBasicLicense) {
            Write-Host "The Entra AD license is AAD_BASIC." -ForegroundColor DarkGreen
            Write-Host "SMTP Client Submission is not available for this license." -ForegroundColor Red
            Write-Host "Return to menu, Buy a better license or choose a different method." -ForegroundColor Red
        }
        else {
            Write-Host "The Entra AD license is not AAD_BASIC." -ForegroundColor DarkGreen
            Write-Host "Continuing Script." -ForegroundColor DarkGreen
        }

        # Security Defaults check
        $defaultsEnabled = Get-MgOrganizationPolicy | Where-Object { $_.Name -eq "SecurityDefaults" } | Select-Object -ExpandProperty Value

        if ($defaultsEnabled) {
            if ($isBasicLicense) {
                Write-Host "Security defaults are enabled, but the Entra AD license is AAD_BASIC."
                Write-Host "Please upgrade your license or select a different send method."
                Write-Host "You may also consider disabling security defaults after upgrading."
            }
            else {
                Write-Host "Security defaults are enabled, but the Entra AD license is not AAD_BASIC."
                Write-Host "Please disable security defaults and use Conditional Access policies for better security."
            }
        }
        else {
            if ($isBasicLicense) {
                Write-Host "Security defaults are disabled, but the Entra AD license is AAD_BASIC."
                Write-Host "Please consider enabling security defaults or selecting a different send method."
                Write-Host "You may also consider upgrading your license for additional features."
            }
            else {
                Write-Host "Security defaults are disabled, and the Entra AD license is not AAD_BASIC."
                Write-Host "Continuing script."
            }
        }

        # Enable SMTP Client Authentication & verify target mailbox license.
        Write-Host "Enabling SMTP Client Authentication." -ForegroundColor Green
        Start-Sleep 2
        $targetSMTPmailbox = Read-Host -Prompt "Target SMTP Mailbox :"
        Set-CASMailbox -Identity $targetSMTPmailbox -SmtpClientAuthenticationDisabled $false

        # Verify if the targetSMTPmailbox has at least Exchange Online license
        Write-Host "Verify if the targetSMTPmailbox has at least Exchange Online license" -ForegroundColor Green
        $targetMailboxLicense = Get-MgUserLicenseDetail -UserId $targetSMTPmailbox
        $isExchangeOnlineLicense = $targetMailboxLicense.Licenses -contains $licenses['Microsoft Exchange Online']

        if (-not $isExchangeOnlineLicense) {
            Write-Host "The targetSMTPmailbox does not have the required Exchange Online license." -ForegroundColor Red
            Write-Host "Please assign the appropriate license to the mailbox."  -ForegroundColor Red
            Write-Host "SMTP Client Submission will not work without the Exchange Online license." -ForegroundColor Red

        # Disable Multi Factor Authentication (MFA) on the licensed mailbox being used.
        # Check if MFA is enabled or enforced for the targetSMTP mailbox
        $mfaStatus = Get-MgUser -UserId $targetSMTP | Select-Object -ExpandProperty StrongAuthenticationMethods

        if ($null -ne $mfaStatus) {
            # Disable MFA for the targetSMTP mailbox
            Disable-MgUserStrongAuthentication -UserId $targetSMTP
            Write-Host "Multi-factor authentication has been disabled for the targetSMTP mailbox." -ForegroundColor Green
        }
        else {
            Write-Host "Multi-factor authentication is not enabled or enforced for the targetSMTP mailbox." -ForegroundColor Green
        }

        # Exclude user from each CA policy
        Write-Host "Excluding target mailbox from CA Policies." -ForegroundColor Green
        $caPolicies = Get-MgConditionalAccessPolicy
        foreach ($policy in $caPolicies) {
            $excludedUsers = $policy.ExcludedUsers
            $excludedUsers += $targetSMTPuser
            Set-MgConditionalAccessPolicy -Id $policy.Id -ExcludedUsers $excludedUsers
        }
    }
    end {
        Invoke-ExoConnectorWiz
    }
}}

function Invoke-SMTPDeviceInform {
    begin {}
    process {
    $smtpconfig = @{

        'Server Smarthost'       = 'smtp.office365.com'
        'Port'                   = '587 (recommended) or 25'
        'TLS/StartTLS'           = 'This must be enabled, and only TLS 1.2 is supported.'
        'Username/email address' = '{$}smtpTargetUser'
        'password'               = '{$}smtpTargetUsers password'
        'Note'                   = 'If you are using a Microsoft 365 mailbox, you can use your Microsoft 365 username and password here.'
    }
    Write-Host "Enter the following settings on your device or application."
    Write-Host $smtpconfig
    Write-Host ""
    Write-Host "Test out with application. Swaks on linux , SMTPer or whatever. I don't think Send-MailMessage works anymore."
    Write-Host ""
    }
    end {Invoke-ExoConnectorWiz}
}

function Invoke-TransportInfo {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        
    }
    
    process {
        Write-Host "Creating & Enabling Exchange Rule for internal spam filter bypass"
        Write-Host ""
        Write-Host "Sender is: $targetSMTPuser "
        Write-Host "Recipient is: Inside the organization"
        Write-Host "Modify the message properties: Set spam confidence level"
        Write-Host "Set the spam confidence level to: -1"
    }
    
    end {
        Invoke-ExoConnectorWiz
    }
}

function Set-TransportInfo {
    [CmdletBinding()]
    param (
        
    )
    
    begin {
        
    }
    
    process {
        # Create and enable Exchange rule for internal spam filter bypass
        $targetSMTPuser = Read-Host "TargetSMTPuser : "
        $ruleName = "Bypass Internal Spam Filter"
        $spamConfidenceLevel = -1
        
        $rule = New-TransportRule -Name $ruleName -From $targetSMTPuser -SentToScope "InOrganization" -ApplyClassification "ModifyMessage" -SetHeaderName "X-SpamConfidenceLevel" -SetHeaderValue $spamConfidenceLevel
        Enable-TransportRule -Identity $rule.Identity
    }
    
    end {
        Invoke-ExoConnectorWiz
    }
}


# Invoke-SMTPRelay Functions
function Invoke-SMTPRelayFunction {
    [CmdletBinding()]
    param ()
    
    begin {
        
    }
    
    process {
        Write-Host "You have selected : SMTP Relay"-ForegroundColor Cyan
        Write-Host " "
        Write-Host "Going to use your current WAN IP as the value in the SPF Record and Connector. Please indicate if this is ok or you want to manually input IP."
        $useCurrentWANIP = $null
        while ($null -eq $useCurrentWANIP) {
            $useCurrentWANIP = Read-Host "Use current WAN IP? (Y/N)"
            if ($useCurrentWANIP -eq "Y") {
                $useCurrentWANIP = $false
            } elseif ($useCurrentWANIP -eq "N") {
                $useCurrentWANIP = $true
                $WANIP = Read-Host "Enter Custom WAN IP:"
            } else {
                Write-Host "Invalid Input"
                $useCurrentWANIP = $null
            }
        }
        
        if ($useCurrentWANIP -eq $false) {
            $WANIP = (Invoke-RestMethod http://ipinfo.io/json).ip
        } else {
            $WANIP = Read-Host -Prompt "Enter your desired IPv4 address where the mail sends from. It has to be static."
        }
        
        Write-Host "Your connector and SPF choice: $WANIP"
        Write-Host ""
        
        $domain = Read-Host "Enter your domain:"
        if([string]::IsNullOrEmpty($Domain)) { 
            throw "Domain and Policy are required parameters."
            $Domain = Read-Host -Prompt "What domain are you sending from:"
        }

        Write-Host "Parsing MX for $domain"
        $mxRecord = (Resolve-DnsName -Name $domain -Type MX).NameExchange
        Write-Host "MX record for $domain is $mxRecord"
        Write-Host ""
        
        $currentSPF = ((Resolve-DnsName -Name $domain -Type TXT) | Where-Object { $_.Strings -match "v=spf1" }).Strings
        Write-Host ""
        Write-Host "Please create a TXT record in your DNS with the following value: v=spf1 a mx ip4:$WANIP include:spf.protection.outlook.com -all"
        Write-Host ""
        Write-Host ""
        
        Write-Host "Creating a connector"
        $parameters = @{
            Name = "$domain SMTP Relay"
            ConnectorType = 'OnPremises'
            SenderDomains = '*'
            SenderIPAddresses = $WANIP
            RestrictDomainsToIPAddresses = $true
        }
        
        # Create the SMTP Relay connector
        New-InboundConnector @parameters
        Write-Host "Created SMTP Relay connector with the following parameters: $parameters"
        
        # Validate SPF record
        if ($currentSPF -notcontains $WANIP) { 
            Write-Host "SPF Invalid or does not contain server IP"
        }
        
        # Verify connector is enabled
        Get-InboundConnector -Name "$domain SMTP Relay"
        
        # Write SMTP Configuration
        $smtpConfig = @"
Server Smarthost: smtp.office365.com
Port: 587 (recommended) or 25
TLS/StartTLS: This must be enabled, and only TLS 1.2 is supported.
Email Address to send email from: [Enter email address here]
Note:
"@
        Write-Host $smtpConfig
        Write-Host ''
        # Write SPF record
        Write-Host $currentSPF
        Write-Host ''
        # Write MX record
        Write-Host $mxRecord
    }
    
    end {
        Invoke-ExoConnectorWiz
    }
}


function Start-ExoWizard {
    # Start-Wizard Interactive aka fuck you I don't know what I need...
    # The wizard tells you what to pick.
    #    Fucking kill me this was conditional agony.
    # calvindd2f 21-12-2023

    $q1 = Read-Host "Do you need to send more than 10k msgs/day or faster than 30 msgs/min? (Y/N)"

    if ($q1 -eq "Y" -or $q1 -eq "y") {
        # Set the a1 to true if the q1 is Yes
        $a1 = $true
        Write-Host "Yes."
    } elseif ($q1 -eq "N" -or $q1 -eq "n") {
        # Set the a1 to false if the q1 is No
        $a1 = $false
        Write-Host "No."
    } else {
        # Throw an error for invalid input
        throw "Invalid input. Please enter 'Y' or 'N'."
    }

    $q2 = Read-Host "Do you need to send from more than one email address? (Y/N)"
    if ($q2 -eq "Y" -or $q2 -eq "y") {
        # Set the a2 to true if the q2 is Yes
        $a2 = $true
        Write-Host "Yes."
    } elseif ($q2 -eq "N" -or $q2 -eq "n") {
        # Set the a2 to false if the q2 is No
        $a2 = $false
        Write-Host "No."
    } else {
        # Throw an error for invalid input
        throw "Invalid input. Please enter 'Y' or 'N'."
    }

    $q3 = Read-Host "Do you need to send to recipients outside your organization? (Y/N)"
    if ($q3 -eq "Y" -or $q3 -eq "y") {
        # Set the a3 to true if the q3 is Yes
        $a3 = $true
        Write-Host "Yes."
    } elseif ($q3 -eq "N" -or $q3 -eq "n") {
        # Set the a3 to false if the q3 is No
        $a3 = $false
        Write-Host "No."
    } else {
        # Throw an error for invalid input
        throw "Invalid input. Please enter 'Y' or 'N'."
    }

    $q4 = Read-Host "Do you have a licensed mailbox to send mail through? (Y/N)"
    if ($q4 -eq "Y" -or $q4 -eq "y") {
        # Set the a4 to true if the q4 is Yes
        $a4 = $true
        Write-Host "Yes."
    } elseif ($q4 -eq "N" -or $q4 -eq "n") {
        # Set the a4 to false if the q4 is No
        $a4 = $false
        Write-Host "No."
    } else {
        # Throw an error for invalid input
        throw "Invalid input. Please enter 'Y' or 'N'."
    }

    $q5 = Read-Host "Can your device or application be set up with the user name and password of the mailbox you'll use to send email from? (Y/N)"
    if ($q5 -eq "Y" -or $q5 -eq "y") {
        # Set the a5 to true if the q5 is Yes
        $a5 = $true
        Write-Host "Yes."
    } elseif ($q5 -eq "N" -or $q5 -eq "n") {
        # Set the a5 to false if the q5 is No
        $a5 = $false
        Write-Host "No."
    } else {
        # Throw an error for invalid input
        throw "Invalid input. Please enter 'Y' or 'N'."
    }
    
    Write-Host "Processing..."
    Start-Sleep -Seconds 3
    # Insert Logic for assessing answers, then output the users choice + return
    #################################################################
    if (-not $a3) {
        Write-Host "Use Direct Send method" -ForegroundColor Green
        Write-Host "----------------------------" -ForegroundColor Gray
        Write-Host "Dumping Documenation for doing this."
        Write-Host $DirectSend
        Write-Host ""
        $runDirectSend=Read-Host "Run Direct Send? (Y/N)"
        if($runDirectSend){
            Invoke-DirectSendFunction
        } else {
            Write-Host "If you exit to the menu ; you can get assistance in this scenario." -ForegroundColor Blue
            Write-Host "There are functions to generate the Connectors based on your inputs, show your current spf and what spf you should use" -ForegroundColor Blue
            Write-Host "Returning to menu in 3 seconds." -ForegroundColor Red
            Start-Sleep -Seconds 3
            Invoke-ExoConnectorWiz
        }
    } elseif ($a1 -eq $false -and $a2 -eq $false -and $a3 -eq $true -and $a4 -eq $true -and $a5 -eq $true) {
        Write-Host "Use SMTP Client Submission method" -ForegroundColor Green
        Write-Host "----------------------------" -ForegroundColor Gray
        Write-Host "Dumping Documenation for doing this."
        Write-Host $SMTPClient
        Write-Host ""
        $runSMTPCS=Read-Host "Run SMTP Client Submission? (Y/N)"
        if($runSMTPCS){
            Invoke-SMTPClientSubmissionFunction
        } else {
            Write-Host "If you exit to the menu ; you can get assistance in this scenario." -ForegroundColor Blue
            Write-Host "There are functions to generate the Connectors based on your inputs, show your current spf and what spf you should use" -ForegroundColor Blue
            Write-Host "Returning to menu in 3 seconds." -ForegroundColor Red
            Start-Sleep -Seconds 3
            Invoke-ExoConnectorWiz
        }
    } else {
        Write-Host "Use SMTP Relay method" -ForegroundColor Green
        Write-Host "----------------------------" -ForegroundColor Gray
        Write-Host "Dumping Documenation for doing this."
        Write-Host $SMTPRelay
        Write-Host ""
        $runSMTPRelay=Read-Host "Run SMTP Relay? (Y/N)"
        if($runSMTPRelay){
            Invoke-SMTPRelayFunction
        } else {
            Write-Host "If you exit to the menu ; you can get assistance in this scenario." -ForegroundColor Blue
            Write-Host "There are functions to generate the Connectors based on your inputs, show your current spf and what spf you should use" -ForegroundColor Blue
            Write-Host "Returning to menu in 3 seconds." -ForegroundColor Red
            Start-Sleep -Seconds 3
            Invoke-ExoConnectorWiz
        }
    }
}



#########################
#     Documentation     #
#########################
function Show-DocumentationMenu {
    Write-Host "`nChoose an option from the documentation menu:"
    Write-Host "0) Dumpdoc SMTP Relay"
    Write-Host "1) Dumpdoc SMTP Client Substitution"
    Write-Host "2) Dumpdoc Direct Send"
    Write-Host "q) Quit"

    $userDocChoice = Read-Host "Enter the number of your option"

    switch ($userDocChoice) {
        '0' { DumpdocSMTPRelay }
        '1' { DumpdocSMTPClientSub }
        '2' { DumpdocDirectSend }
        'q' { Invoke-ExoConnectorWiz }
        default { Write-Host "Invalid choice. Please enter a valid number or letter." }
    }
}

Function DumpdocSMTPRelay {
    Clear-Host
    Write-Host "Use the SMTP Relay method" -ForegroundColor Green
    Write-Output ''
    Write-Output ''
    Write-Output ''
    Write-Host "Option 1 (preferred): By verifying the Subject Alternative Name or Common Name on the TLS certificate sent by the sending server or device." -ForegroundColor Green
    Write-Output ''
    Write-Host "Note:" -ForegroundColor Red
    Write-Host "For security reasons, we recommend the sender's domain match one of your Accepted domains in Microsoft 365, but with this option, you can use any domain in the sender address."
    Write-Output ''
    Write-Host "Option 2: By verifying the IP address of the sending server or device." -ForegroundColor Green
    Write-Output ''
    Write-Host "Note: With this option, the sender's domain must match one of your Accepted domains in Microsoft 365. " -ForegroundColor Red
    Write-Output ''
    Write-Host "To determine which option is best for you, please answer a few more questions" 
    Write-Output ''
    Write-Host "Do you have, or can you get, a certificate whose Common Name or Subject Alternative Name contains one of your Accepted domains?"
    Write-Host "NO BECAUSE MOST PEOPLE GET FUCKING RETARDED ABOUT CERTIFICATES AND IF CUSTOMER IS USING DYNAMIC IP FUCK OFF."
    Write-Host "Do you have a dedicated and static public IP address which will be used to send email to Office 365?"
    Write-Host "YES"
    Write-Host "Use the SMTP Relay option by verifying the sending server's or device's IP address."  -ForegroundColor Yellow
    Write-Output ''
    Write-Host "1. First find the MX endpoint for your domain " -ForegroundColor Green
    Write-Output ''
    Write-Host "Sign in to the Microsoft 365 admin center" -ForegroundColor Yellow
    Write-Host "Go to Settings > Domains"  -ForegroundColor Yellow
    Write-Host "Select your domain (for example, contoso.com)"  -ForegroundColor Yellow 
    Write-Host "Select the DNS records tab and locate the MX record in the "Microsoft Exchange" table. "  -ForegroundColor Yellow
    Write-Output ''
    Write-Output ''
    Write-Host "2. Enter the following settings on your server, device or application" -ForegroundColor Green
    Write-Output ''
    Write-Host "Server Smarthost: Enter the MX endpoint for your domain which you located in the step above. For example: contoso-com.mail.protection.outlook.com" -ForegroundColor Cyan
    Write-Host "Port: 25" -ForegroundColor Cyan
    Write-Host "TLS/StartTLS : This is optional, however if you use TLS or StartTLS, only TLS 1.2 version is supported." -ForegroundColor Cyan
    Write-Host "Email Address to send email from: This should be any email address matching your Office 365 Accepted Domain. This email address does not need to have a mailbox." -ForegroundColor Cyan
    Write-Output ''
    Write-Host "Note:" -ForegroundColor Yellow
    Write-Host "To avoid having messages flagged as spam, we recommend adding an SPF record for your domain in the DNS settings at your domain registrar. Add the static IP address you''re sending from to the SPF record." -ForegroundColor Red
    Write-Output ''
    Write-Output ''
    Write-Output ''
    Write-Host "3. Create and Configure an Inbound Connector in your Microsoft 365 Organization"  -ForegroundColor Green
    Write-Output ''
    Write-Host "From: Your organization's email server"  -ForegroundColor Yellow
    Write-Host "To: Office 365" -ForegroundColor Yellow
    Write-Host "Name:  Any descriptive name you wish. " -ForegroundColor Yellow
    Write-Host "What do you want to do after connector is saved: Leave the checkboxes selected.  " -ForegroundColor Yellow
    Write-Host "How should Office 365 identify email from your email server:  By verifying that the IP address of the sending server matches one of the following IP addresses, which belong exclusively to your organization" -ForegroundColor Yellow
    Write-Host "(Enter the dedicated public IP address of your server/device here.) " -ForegroundColor Yellow
    Write-Host ''
    $wait=Read-Host "Press Any to return to menu..."
    if(!([string]::IsNullOrEmpty($wait))) { 
        Invoke-ExoConnectorWiz
    }
}

Function DumpdocSMTPClientSub {
    Clear-Host
    #Write-Host "Dumping Tutorial" -ForegroundColor Green.\_Automations2024Write-Host ""
    Write-Host  'Use the SMTP Client Submission method' -ForegroundColor Green
    Write-Host  ""
    Write-Host  'Use the following instructions to configure SMTP Client Submission:'
    Write-Host  ""
    Write-Host ' 1. Disable the Azure Security Defaults by toggling the “Enable Security Defaults” to “No”. See the following link for details:' -ForegroundColor Green
    Write-Host    'Azure Security Defaults' -ForegroundColor Yellow
    Write-Host  ""
    Write-Host  '2. Run the following remote powershell command to enable SMTP Client Authentication on the licensed mailbox being used.'   -ForegroundColor Green
    Write-Host    'Set-CASMailbox -Identity user@contoso.com -SmtpClientAuthenticationDisabled $false' -ForegroundColor Yellow
    Write-Host
    Write-Host  '3. Disable Multi Factor Authentication (MFA) on the licensed mailbox being used.'  -ForegroundColor Green
    Write-Host    'In the Microsoft 365 admin center, in the left navigation menu choose Users > Active users.' -ForegroundColor Yellow
    Write-Host   ' On the Active users page, choose Multi-factor authentication.' -ForegroundColor Yellow
    Write-Host    'On the Multi-factor authentication page, select the user and set the Multi-factor auth status to Disabled.' -ForegroundColor Yellow
    Write-Host  ""
    Write-Host  '4. Enter the following settings on your device or application.'    -ForegroundColor Green
    Write-Host  ""
    Write-Host   'Server Smarthost: smtp.office365.com (This is the default endpoint for Office 365 client submission)' -ForegroundColor   Cyan
    Write-Host   ' Port: 587 (recommended) or 25'   -ForegroundColor   Cyan
    Write-Host    'TLS/StartTLS: This must be enabled, and only TLS 1.2 is supported.'  -ForegroundColor   Cyan
    Write-Host    'Username/email address and password: Enter the sign in credentials of the hosted mailbox being used. '   -ForegroundColor   Cyan
    Write-Host     ""
    Write-Host     "Note:" -ForegroundColor Red 
    Write-Host    'To avoid having messages flagged as spam, we recommend adding an SPF record for your domain in the DNS settings at your domain registrar. Additionally if you are sending from a static IP address, add that address to your SPF record.'
    Write-Host    'For more information see this article - How to set up a multifunction device or application to send email using Microsoft 365 or Office 365 .'
}

function DumpdocDirectSend {
    Clear-Host
    Write-Host "Use the Direct Send Method " -ForegroundColor Green
    Write-Output ''
    Write-Host "Use the following instructions to configure Direct Send:"
    Write-Output ''
    Write-Host "1. First find the MX endpoint for your domain" -ForegroundColor Green
    Write-Host "Sign in to the Microsoft 365 admin center" -ForegroundColor Yellow
    Write-Host "Go to Settings > Domains" -ForegroundColor Yellow 
    Write-Host "Select your domain (for example, contoso.com)" -ForegroundColor Yellow 
    Write-Host "Select the DNS records tab and locate the MX record in the "Microsoft Exchange" table." -ForegroundColor Yellow
    Write-Output ''
    Write-Host "2. Enter the following settings on your device or application" -ForegroundColor Green
    Write-Host "Server Smarthost: Enter the MX endpoint for your domain which you located above. For example: contoso-com.mail.protection.outlook.com" -ForegroundColor   Cyan
    Write-Host "Port: 25" -ForegroundColor   Cyan
    Write-Host "TLS/StartTLS : This is optional, however if you use TLS or StartTLS, only TLS 1.2 version is supported." -ForegroundColor   Cyan
    Write-Host "Email Address to send email from: This could be an email address matching any of your your Microsoft 365 Accepted domains. This email address does not need to have a mailbox." -ForegroundColor   Cyan
    Write-Outpute-Output ''
    Write-Host "Note:"  -ForegroundColor Red 
    Write-Host "To avoid having messages flagged as spam, we recommend adding an SPF record for your domain in the DNS settings at your domain registrar. Additionally if you are sending from a static IP address, add that address to your SPF record."
}
