# microsoft-tools
A collection of PowerShell scripts to automate provisioning Microsoft 365 Exchange Online to work with MXGuardian

# Usage

## ConfigureInbound.ps1

    ConfigureInbound.ps1 [-Help] [-Force]

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
