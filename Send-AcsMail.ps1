<#
.SYNOPSIS
    Dynamic blat.exe replacement built on Azure Communication Services (ACS)
    SMTP relay, with a choice of two credential storage back ends.

.DESCRIPTION
    Sends an email via ACS SMTP (smtp.azurecomm.net). All message data
    (sender, recipients, subject, body, attachments) are passed as
    parameters instead of being hard-coded, and the credential is never
    stored in plain text inside this script.

    Two credential sources are supported, chosen by which parameter set you
    use (created with Set-AcsSmtpSecret.ps1):

      ClixmlFile  -CredentialPath <file>
                  Reads a DPAPI-encrypted credential file created with
                  Set-AcsSmtpSecret.ps1 (ClixmlFile mode). Only readable by
                  the same Windows user account on the same machine.

      KeyVault    -VaultName <vault> -SecretName <secret> [-UsernameSecretName <secret>]
                  Reads the SMTP username and client secret from an Azure
                  Key Vault created with Set-AcsSmtpSecret.ps1 (KeyVault
                  mode). Requires the Az.KeyVault module and an already
                  authenticated Az session in the current process
                  (Connect-AzAccount, a managed identity, or a service
                  principal). This is the "central storage" option: any
                  machine/account with read access to the vault can send
                  mail without a local credential file.

    Exit codes match blat.exe (0 = success, 1 = failure) so existing batch
    files / scheduled tasks that check %ERRORLEVEL% keep working unchanged.

.PARAMETER SmtpServer
    SMTP host. Default: smtp.azurecomm.net

.PARAMETER Port
    SMTP port. Default: 587 (STARTTLS)

.PARAMETER CredentialPath
    (ClixmlFile set) Path to the encrypted credential file.

.PARAMETER VaultName
    (KeyVault set) Azure Key Vault name.

.PARAMETER SecretName
    (KeyVault set) Name of the Key Vault secret holding the client secret.

.PARAMETER UsernameSecretName
    (KeyVault set) Name of the Key Vault secret holding the SMTP username.
    Defaults to "<SecretName>-Username".

.PARAMETER From
    Sender address.

.PARAMETER To
    One or more recipient addresses. Accepts an array or a single
    comma-separated string (useful when called from a batch file).

.PARAMETER Cc
    Optional CC recipients, same format as -To.

.PARAMETER Bcc
    Optional BCC recipients, same format as -To.

.PARAMETER Subject
    Mail subject.

.PARAMETER Body
    Mail body text. Ignored if -BodyFile is given.

.PARAMETER BodyFile
    Path to a text file whose content becomes the mail body (blat.exe style
    body-from-file).

.PARAMETER Html
    Treat the body as HTML instead of plain text.

.PARAMETER Attachment
    One or more file paths to attach.

.PARAMETER LogFile
    Optional path to append a timestamped log line for each send attempt.

.EXAMPLE
    .\Send-AcsMail.ps1 -CredentialPath .\acs-smtp.cred.xml `
        -From "noreply@example.com" -To "user@example.com" `
        -Subject "Test" -Body "This is a test."

.EXAMPLE
    .\Send-AcsMail.ps1 -VaultName "my-vault" -SecretName "acs-smtp-secret" `
        -From "noreply@example.com" -To "a@example.com","b@example.com" `
        -Cc "manager@example.com" -Subject "Report" -BodyFile .\mailbody.txt `
        -Attachment ".\report.pdf", ".\log.txt" -LogFile .\mail.log

.NOTES
    Author:  RaptileBytez
    Version: 1.0.0
    Created: 2026-10-01
#>

[CmdletBinding(DefaultParameterSetName = 'ClixmlFile')]
param(
    [string]$SmtpServer = "smtp.azurecomm.net",
    [int]$Port = 587,

    [Parameter(Mandatory = $true, ParameterSetName = 'ClixmlFile')]
    [string]$CredentialPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'KeyVault')]
    [string]$VaultName,

    [Parameter(Mandatory = $true, ParameterSetName = 'KeyVault')]
    [string]$SecretName,

    [Parameter(ParameterSetName = 'KeyVault')]
    [string]$UsernameSecretName,

    [Parameter(Mandatory = $true)]
    [string]$From,

    [Parameter(Mandatory = $true)]
    [string[]]$To,

    [string[]]$Cc,
    [string[]]$Bcc,

    [Parameter(Mandatory = $true)]
    [string]$Subject,

    [string]$Body,
    [string]$BodyFile,
    [switch]$Html,

    [string[]]$Attachment,

    [string]$LogFile
)

function Write-MailLog {
    param([string]$Message)
    $line = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    if ($LogFile) {
        try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 } catch { }
    }
    Write-Verbose $line
}

function Expand-AddressList {
    param([string[]]$Addresses)
    if (-not $Addresses) { return @() }
    $Addresses |
        ForEach-Object { $_ -split ',' } |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ }
}

function ConvertFrom-SecureStringPlain {
    # Works on both Windows PowerShell 5.1 and PowerShell 7+, unlike
    # ConvertFrom-SecureString -AsPlainText which needs PS 6+.
    param([System.Security.SecureString]$SecureString)
    if (-not $SecureString) { return $null }
    return [System.Net.NetworkCredential]::new('', $SecureString).Password
}

# --- Resolve the SMTP credential from the selected back end ----------------

$smtpUsername = $null
$smtpPasswordSecure = $null

switch ($PSCmdlet.ParameterSetName) {

    'ClixmlFile' {
        if (-not (Test-Path -LiteralPath $CredentialPath)) {
            Write-MailLog "ERROR: credential file not found: $CredentialPath"
            Write-Error "Credential file not found: $CredentialPath. Run Set-AcsSmtpSecret.ps1 first."
            exit 1
        }
        try {
            $credential = Import-Clixml -Path $CredentialPath
            $smtpUsername = $credential.UserName
            $smtpPasswordSecure = $credential.Password
        }
        catch {
            Write-MailLog "ERROR: could not decrypt credential file: $($_.Exception.Message)"
            Write-Error "Could not decrypt the credential file. It is bound to the user account and machine it was created on - re-run Set-AcsSmtpSecret.ps1 there if needed."
            exit 1
        }
    }

    'KeyVault' {
        if (-not (Get-Module -ListAvailable -Name Az.KeyVault)) {
            Write-MailLog "ERROR: Az.KeyVault module not available"
            Write-Error "The Az.KeyVault module is not installed. Install it with: Install-Module Az.KeyVault -Scope CurrentUser"
            exit 1
        }
        if (-not $UsernameSecretName) { $UsernameSecretName = "$SecretName-Username" }

        try {
            $passwordSecretObj = Get-AzKeyVaultSecret -VaultName $VaultName -Name $SecretName -ErrorAction Stop
            $usernameSecretObj = Get-AzKeyVaultSecret -VaultName $VaultName -Name $UsernameSecretName -ErrorAction Stop
            $smtpPasswordSecure = $passwordSecretObj.SecretValue
            $smtpUsername = ConvertFrom-SecureStringPlain -SecureString $usernameSecretObj.SecretValue
        }
        catch {
            Write-MailLog "ERROR: could not read secrets from Key Vault '$VaultName': $($_.Exception.Message)"
            Write-Error "Could not read secrets from Key Vault '$VaultName'. Check that you are authenticated (Connect-AzAccount / managed identity) and have 'get' permission on secrets '$SecretName' and '$UsernameSecretName'."
            exit 1
        }
    }
}

# --- Validate message inputs ------------------------------------------------

if ($BodyFile) {
    if (-not (Test-Path -LiteralPath $BodyFile)) {
        Write-MailLog "ERROR: body file not found: $BodyFile"
        Write-Error "Body file not found: $BodyFile"
        exit 1
    }
    $Body = Get-Content -LiteralPath $BodyFile -Raw
}
elseif (-not $Body) {
    $Body = ""
}

$toList = Expand-AddressList $To
if ($toList.Count -eq 0) {
    Write-MailLog "ERROR: no valid recipient given"
    Write-Error "At least one valid recipient (-To) is required."
    exit 1
}
$ccList  = Expand-AddressList $Cc
$bccList = Expand-AddressList $Bcc

$attachmentFiles = @()
if ($Attachment) {
    foreach ($file in $Attachment) {
        if (-not (Test-Path -LiteralPath $file)) {
            Write-MailLog "ERROR: attachment not found: $file"
            Write-Error "Attachment not found: $file"
            exit 1
        }
        $attachmentFiles += (Resolve-Path -LiteralPath $file).ProviderPath
    }
}

# --- Build and send the message ---------------------------------------------

$mail = $null
$smtp = $null
$attachmentObjects = @()

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $mail = New-Object System.Net.Mail.MailMessage
    $mail.From = $From
    foreach ($addr in $toList)  { $mail.To.Add($addr) }
    foreach ($addr in $ccList)  { $mail.CC.Add($addr) }
    foreach ($addr in $bccList) { $mail.Bcc.Add($addr) }
    $mail.Subject    = $Subject
    $mail.Body       = $Body
    $mail.IsBodyHtml = $Html.IsPresent

    foreach ($file in $attachmentFiles) {
        $att = New-Object System.Net.Mail.Attachment($file)
        $attachmentObjects += $att
        $mail.Attachments.Add($att)
    }

    $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $Port)
    $smtp.EnableSsl   = $true
    $smtp.Credentials = New-Object System.Net.NetworkCredential(
        $smtpUsername,
        (ConvertFrom-SecureStringPlain -SecureString $smtpPasswordSecure)
    )

    $smtp.Send($mail)

    Write-MailLog "OK: mail sent to [$($toList -join ', ')] subject='$Subject' (credential source: $($PSCmdlet.ParameterSetName))"
    Write-Host "Email sent successfully." -ForegroundColor Green
    exit 0
}
catch {
    Write-MailLog "ERROR while sending: $($_.Exception.Message)"
    Write-Error "Send failed: $($_.Exception.Message)"
    exit 1
}
finally {
    foreach ($att in $attachmentObjects) { $att.Dispose() }
    if ($mail) { $mail.Dispose() }
    if ($smtp) { $smtp.Dispose() }
}