<#
.SYNOPSIS
    RelayCourier setup: stores the Azure Communication Services (ACS) SMTP credential (SMTP
    username + Entra application client secret) for later use by
    Send-RelayMail.ps1.

.DESCRIPTION
    Supports two storage back ends, selected via the parameter set:

      ClixmlFile  Encrypts the credential with Export-Clixml. The secret is
                  protected by the Windows Data Protection API (DPAPI) and is
                  only readable by the same Windows user account on the same
                  machine it was created on. No extra module required.

      KeyVault    Stores the credential as two secrets in an Azure Key Vault
                  (centralized storage, usable from any machine/account that
                  has "get"/"set" permission on the vault). Requires the
                  Az.KeyVault module and an already-authenticated Az session
                  (Connect-AzAccount, a managed identity, or a service
                  principal context) in the current PowerShell session.

    The client secret is always entered interactively (Read-Host
    -AsSecureString) and is never written to disk, a script file, or
    console output in plain text.

.PARAMETER Username
    The ACS "SMTP Username" created in the Azure Portal (NOT the Entra
    application/client ID).

.PARAMETER Path
    (ClixmlFile set) File path for the encrypted credential file.

.PARAMETER VaultName
    (KeyVault set) Name of the Azure Key Vault.

.PARAMETER SecretName
    (KeyVault set) Name of the Key Vault secret that will hold the client
    secret (password).

.PARAMETER UsernameSecretName
    (KeyVault set) Name of the Key Vault secret that will hold the SMTP
    username. Defaults to "<SecretName>-Username".

.EXAMPLE
    .\Set-RelayCredential.ps1 -Username "<SMTP Username>" -Path .\relaycourier.cred.xml

.EXAMPLE
    Connect-AzAccount
    .\Set-RelayCredential.ps1 -Username "<SMTP Username>" -VaultName "my-vault" -SecretName "relaycourier-secret"

.NOTES
    Product: RelayCourier
    Author:  RaptileBytez
    Version: 1.0.1
    Created: 2026-10-01
    Modified: 2026-10-01
#>

[CmdletBinding(DefaultParameterSetName = 'ClixmlFile')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'ClixmlFile')]
    [Parameter(Mandatory = $true, ParameterSetName = 'KeyVault')]
    [string]$Username,

    [Parameter(Mandatory = $true, ParameterSetName = 'ClixmlFile')]
    [string]$Path,

    [Parameter(Mandatory = $true, ParameterSetName = 'KeyVault')]
    [string]$VaultName,

    [Parameter(Mandatory = $true, ParameterSetName = 'KeyVault')]
    [string]$SecretName,

    [Parameter(ParameterSetName = 'KeyVault')]
    [string]$UsernameSecretName
)

$securePassword = Read-Host -Prompt "Entra Application Client Secret" -AsSecureString
if ($securePassword.Length -eq 0) {
    Write-Error "No secret entered - aborting."
    exit 1
}

switch ($PSCmdlet.ParameterSetName) {

    'ClixmlFile' {
        $credential = New-Object System.Management.Automation.PSCredential($Username, $securePassword)
        try {
            $credential | Export-Clixml -LiteralPath $Path -Force
            Write-Host "Credential saved to '$Path'." -ForegroundColor Green
            Write-Host "This file can only be decrypted by the current Windows user account on this machine." -ForegroundColor Yellow
            exit 0
        }
        catch {
            Write-Error "Failed to save credential file: $($_.Exception.Message)"
            exit 1
        }
    }

    'KeyVault' {
        if (-not (Get-Module -ListAvailable -Name Az.KeyVault)) {
            Write-Error "The Az.KeyVault module is not installed. Install it with: Install-Module Az.KeyVault -Scope CurrentUser"
            exit 1
        }
        if (-not $UsernameSecretName) { $UsernameSecretName = "$SecretName-Username" }

        try {
            $secureUsername = ConvertTo-SecureString -String $Username -AsPlainText -Force
            Set-AzKeyVaultSecret -VaultName $VaultName -Name $SecretName -SecretValue $securePassword | Out-Null
            Set-AzKeyVaultSecret -VaultName $VaultName -Name $UsernameSecretName -SecretValue $secureUsername | Out-Null
            Write-Host "Secrets stored in Key Vault '$VaultName':" -ForegroundColor Green
            Write-Host "  Password secret : $SecretName"
            Write-Host "  Username secret : $UsernameSecretName"
            exit 0
        }
        catch {
            Write-Error "Failed to store secrets in Key Vault '$VaultName': $($_.Exception.Message)"
            exit 1
        }
    }
}