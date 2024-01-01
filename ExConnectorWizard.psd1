# Initialize-XOConnectors.psd1

@{
    ModuleVersion = '1.2.4'
    GUID = '4eca504e-977a-4d86-af69-1b48ba20769f'   #[System.Guid]::NewGuid()
    Author = 'Calvin Bergin'
    Copyright = '(c) 2023 Calvin Bergin All rights reserved.'
    Description = 'Assistance with setting up SMTP accounts the correct way. Interactive.'
    FunctionsToExport = @("Show-Menu","Show-Intro","Invoke-ExoConnectorWiz","Start-Authentication","Invoke-SPFGeneratorFunction","New-SMTPConnector","Invoke-DirectSendFunction","Invoke-SMTPSubmissionFunction","Invoke-SMTPDeviceInform","Invoke-TransportInfo","Set-TransportInfo","Invoke-SMTPRelayFunction","Start-ExoWizard","Show-DocumentationMenu","DumpdocSMTPRelay","DumpdocSMTPClientSub","DumpdocDirectSend")  # List of functions to export
    CmdletsToExport = @()  # List of cmdlets to export
    VariablesToExport = @()  # List of variables to export
    AliasesToExport = @('*')  # List of aliases to export
    PrivateData = @{
        PSData = @{
            # Tags applied to your module
            Tags = @('ExchangeOnline', 'Connectors','SMTP','Interactive','DirectSend','SMTPSubmission','SMTPRelay')

            # Prerelease string for pre-release versions
            Prerelease = 'alpha'

            # License URI for your module
            LicenseUri = 'https://opensource.org/licenses/MIT'

            # Project URI for your module
            ProjectUri = 'https://github.com/Calvindd2f/pxwershell'

            # Release notes for your module
            ReleaseNotes = 'First major release.'
        }
    }
    FormatsToProcess = @()  # List of custom format files to process
    TypesToProcess = @()  # List of custom type files to process
    RequiredModules = @('Microsoft.Graph','ExchangeOnlineManagement')  # List of modules required by your module
    NestedModules = @()  # List of nested modules
    # FunctionsToExport = @()  # List of functions to export  # Export all functions
}

