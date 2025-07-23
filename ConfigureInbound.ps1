# This script configures Microsoft Exchange Online for MXGuardian.

# Define parameters for the script
param(
 [switch]$Help,
 [switch]$Force
)

# Initialize variables
$connectorName = "MXGuardian Inbound Connector"
$comment = "Created by $scriptName on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$mxguardianIPs = @(
    "52.202.8.242",
    "54.84.80.162",
    "54.84.69.64",
    "54.84.72.194",
    "54.84.55.19",
    "54.84.86.117",
    "54.84.107.183",
    "54.84.110.207",
    "54.173.91.73"
)
$trustedArcSealers = @(
    "mxguardian.net"
)
$scriptName = $MyInvocation.MyCommand.Name
$warningText = @"
This script will:
 - Install the ExchangeOnlineManagement module if not present.
 - Connect to Exchange Online.
 - Create an inbound connector for MXGuardian
 - Configure the connector to only accept email from MXGuardian's IP addresses.
 - Enable Enhanced Filtering for Connectors (i.e. skip-listing)
 - Set trusted ARC sealers for MXGuardian.

Before running this script, ensure that:
 - All domains in your organization are intended to use MXGuardian for email filtering.
 - You have updated the MX records for your domains to point to MXGuardian.
 - You have waited for DNS propagation to complete.
 - There are no existing connectors that conflict with this configuration.

If you have manually created any inbound connectors for MXGuardian, please delete them before running this script.

Proceed only if you understand the impact.
"@

# Display help information if the -Help switch is used
if ($Help) {
    Write-Host "Usage: $scriptName [-Help] [-Force]"
    Write-Host ""
    Write-Host "Parameters:"
    Write-Host "  -Help   : Display this help information."
    Write-Host "  -Force  : Bypass confirmation prompts and warnings."
    Write-Host ""
    Write-Host $warningText
    exit
}

# Check if the Force switch is used, if not, display a warning
if (-not $Force) {

    Write-Warning $warningText

    $response = Read-Host "Do you want to continue? (Y/N)"
    if ($response -notin @('Y', 'y')) {
        Write-Host "Operation cancelled by user."
        exit 1
    }
}

# Ensure ExchangeOnlineManagement module is installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module..."
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
}

# Import the module
Import-Module ExchangeOnlineManagement -Force

# Connect to Exchange Online only if not already connected
if (-not (Get-ConnectionInformation)) {
    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline
} else {
    Write-Host "Exchange Online session already established."
}

# Check if the connector already exists
$existing = Get-InboundConnector -Identity $connectorName -ErrorAction SilentlyContinue

if ($existing) {
    Write-Host "Connector '$connectorName' already exists. Updating..."
    Set-InboundConnector -Identity $connectorName `
        -SenderIPAddresses $mxguardianIPs `
        -RequireTls $true `
        -RestrictDomainsToIPAddresses $true `
        -RestrictDomainsToCertificate $false `
        -EFSkipLastIP $true `
        -ConnectorType Partner `
        -SenderDomains @("smtp:*;1") `
        -Enabled $true
} else {
    Write-Host "Creating new connector '$connectorName'..."
    New-InboundConnector -Name $connectorName `
        -SenderIPAddresses $mxguardianIPs `
        -RequireTls $true `
        -RestrictDomainsToIPAddresses $true `
        -RestrictDomainsToCertificate $false `
        -EFSkipLastIP $true `
        -ConnectorType Partner `
        -SenderDomains @("smtp:*;1") `
        -Enabled $true `
        -Comment $comment
}

# Set the trusted ARC sealers for MXGuardian
Write-Host "Setting trusted ARC sealers for MXGuardian..."
Set-ArcConfig -Identity Default -ArcTrustedSealers $trustedArcSealers | Out-Null

Write-Host "Configuration complete. The inbound connector for MXGuardian has been set up successfully."
