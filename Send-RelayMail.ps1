<#
.SYNOPSIS
    RelayCourier: send mail from scripts and scheduled tasks through Azure
    Communication Services (ACS) SMTP relay by default, or through any SMTP
    server via -SmtpServer / -Port / -TlsMode / -NoAuth.

.DESCRIPTION
    Sends an email via ACS SMTP (smtp.azurecomm.net:587, STARTTLS) by default.
    Any other SMTP server can be used with -SmtpServer, -Port, -TlsMode and
    -NoAuth (for example an internal relay on port 25 without TLS or
    authentication). All message data (sender, recipients, subject, body,
    attachments) are passed as parameters instead of being hard-coded, and
    the credential is never stored in plain text inside this script.

    Three authentication modes are supported, chosen by which parameter set
    you use (credentials are created with Set-RelayCredential.ps1):

      ClixmlFile  -CredentialPath <file>
                  Reads a DPAPI-encrypted credential file created with
                  Set-RelayCredential.ps1 (ClixmlFile mode). Only readable by
                  the same Windows user account on the same machine.

      KeyVault    -VaultName <vault> -SecretName <secret> [-UsernameSecretName <secret>]
                  Reads the SMTP username and client secret from an Azure
                  Key Vault created with Set-RelayCredential.ps1 (KeyVault
                  mode). Requires the Az.KeyVault module and an already
                  authenticated Az session in the current process
                  (Connect-AzAccount, a managed identity, or a service
                  principal). This is the "central storage" option: any
                  machine/account with read access to the vault can send
                  mail without a local credential file.

      Anonymous   -NoAuth
                  No credential is loaded or sent. For relays that accept
                  anonymous or IP-authenticated mail. Cannot be combined
                  with the ClixmlFile or KeyVault parameters.

    Implicit TLS (SMTPS, port 465) is not supported by System.Net.Mail.

    Exit codes match blat.exe (0 = success, 1 = failure) so existing batch
    files / scheduled tasks that check %ERRORLEVEL% keep working unchanged.

.PARAMETER SmtpServer
    SMTP host. Default: smtp.azurecomm.net
    Alias: -MailServer. A single "host:port" value (blat style) is split
    into host and port; giving both host:port and -Port is an error.
    For IPv6 literals use -Port.

.PARAMETER Port
    SMTP port. Default: 587 (STARTTLS)

.PARAMETER TlsMode
    StartTls (default): upgrade the connection with STARTTLS (TLS 1.2, and
    TLS 1.3 where the platform offers it). None: plain SMTP without TLS, e.g.
    an internal relay on port 25. None cannot be combined with a credential
    (ClixmlFile or KeyVault set): the script fails with exit code 1 before
    loading the credential, because the password would be sent in clear
    text. Use None only together with -NoAuth.

.PARAMETER NoAuth
    (Anonymous set) Send without authentication. Cannot be combined with
    -CredentialPath, -VaultName, -SecretName or -UsernameSecretName.

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
    Sender address. Alias: -MailFrom.

.PARAMETER To
    One or more recipient addresses. Accepts an array or a single
    comma-separated string (useful when called from a batch file).
    Alias: -Recipient.

.PARAMETER Cc
    Optional CC recipients, same format as -To. Alias: -CopyTo.

.PARAMETER Bcc
    Optional BCC recipients, same format as -To. Alias: -BlindCopyTo.

.PARAMETER Subject
    Mail subject. Alias: -Title.

.PARAMETER Body
    Mail body text. Ignored if -BodyFile is given. Alias: -Message.

.PARAMETER BodyFile
    Path to a text file whose content becomes the mail body (blat.exe style
    body-from-file). Alias: -MessageFile.

.PARAMETER Html
    Treat the body as HTML instead of plain text.

.PARAMETER Attachment
    One or more file paths to attach. Alias: -Attach.

.PARAMETER LogFile
    Optional path to append a timestamped log line for each send attempt.
    Alias: -Log.

.EXAMPLE
    .\Send-RelayMail.ps1 -CredentialPath .\relaycourier.cred.xml `
        -From "noreply@example.com" -To "user@example.com" `
        -Subject "Test" -Body "This is a test."

.EXAMPLE
    .\Send-RelayMail.ps1 -VaultName "my-vault" -SecretName "relaycourier-secret" `
        -From "noreply@example.com" -To "a@example.com","b@example.com" `
        -Cc "manager@example.com" -Subject "Report" -BodyFile .\mailbody.txt `
        -Attachment ".\report.pdf", ".\log.txt" -LogFile .\mail.log

.EXAMPLE
    # Anonymous internal relay on port 25, no TLS, blat-style host:port
    .\Send-RelayMail.ps1 -SmtpServer "smtp.example.local:25" -TlsMode None -NoAuth `
        -From "app@example.local" -To "ops@example.local" `
        -Subject "Test" -Body "This is a test."

.EXAMPLE
    # Custom server with STARTTLS and credentials
    .\Send-RelayMail.ps1 -SmtpServer "mail.example.com" -Port 587 `
        -CredentialPath .\mail.cred.xml `
        -From "app@example.com" -To "ops@example.com" `
        -Subject "Test" -Body "This is a test."

.NOTES
    Product: RelayCourier
    Author:  RaptileBytez
    Version: 1.1.0
    Created: 2026-10-01
    Modified: 2026-10-01
#>

[CmdletBinding(DefaultParameterSetName = 'ClixmlFile')]
param(
    [Alias('MailServer')]
    [string]$SmtpServer = "smtp.azurecomm.net",
    [int]$Port = 587,

    [ValidateSet('StartTls', 'None')]
    [string]$TlsMode = 'StartTls',

    [Parameter(Mandatory = $true, ParameterSetName = 'Anonymous')]
    [switch]$NoAuth,

    [Parameter(Mandatory = $true, ParameterSetName = 'ClixmlFile')]
    [string]$CredentialPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'KeyVault')]
    [string]$VaultName,

    [Parameter(Mandatory = $true, ParameterSetName = 'KeyVault')]
    [string]$SecretName,

    [Parameter(ParameterSetName = 'KeyVault')]
    [string]$UsernameSecretName,

    [Parameter(Mandatory = $true)]
    [Alias('MailFrom')]
    [string]$From,

    [Parameter(Mandatory = $true)]
    [Alias('Recipient')]
    [string[]]$To,

    [Alias('CopyTo')]
    [string[]]$Cc,
    [Alias('BlindCopyTo')]
    [string[]]$Bcc,

    [Parameter(Mandatory = $true)]
    [Alias('Title')]
    [string]$Subject,

    [Alias('Message')]
    [string]$Body,
    [Alias('MessageFile')]
    [string]$BodyFile,
    [switch]$Html,

    [Alias('Attach')]
    [string[]]$Attachment,

    [Alias('Log')]
    [string]$LogFile
)

function Write-MailLog {
    param([string]$Message)
    # Flatten line breaks so echoed input cannot forge extra log lines.
    $Message = $Message -replace '[\r\n]+', ' '
    $line = "{0}  {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    if ($LogFile) {
        try { Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8 } catch { }
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

# --- Split "host:port" (blat style) when -Port was not given ----------------

$SmtpServer = $SmtpServer.Trim()

if ($SmtpServer.Split(':').Count -eq 2) {
    if ($PSBoundParameters.ContainsKey('Port')) {
        Write-MailLog "ERROR: port given twice: -SmtpServer '$SmtpServer' and -Port $Port"
        Write-Error "Give the port either in -SmtpServer (host:port) or via -Port, not both."
        exit 1
    }
    $serverParts = $SmtpServer.Split(':')
    $portValue = 0
    if (-not $serverParts[0] -or
        $serverParts[1] -notmatch '^\d{1,5}$' -or
        -not [int]::TryParse($serverParts[1], [ref]$portValue) -or
        $portValue -lt 1 -or $portValue -gt 65535) {
        Write-MailLog "ERROR: invalid host:port value: $SmtpServer"
        Write-Error "Invalid -SmtpServer value '$SmtpServer': expected host:port with a numeric port from 1 to 65535."
        exit 1
    }
    $SmtpServer = $serverParts[0]
    $Port = $portValue
}

# --- Refuse clear-text authentication ----------------------------------------
# Checked before any credential is loaded or decrypted.

if ($TlsMode -eq 'None' -and $PSCmdlet.ParameterSetName -ne 'Anonymous') {
    Write-MailLog "ERROR: -TlsMode None is not allowed with a credential (password would be sent in clear text)"
    Write-Error "-TlsMode None cannot be combined with a credential because the password would be sent in clear text. Use STARTTLS (the default) or -NoAuth."
    exit 1
}

# --- Resolve the SMTP credential from the selected back end ----------------

$smtpUsername = $null
$smtpPasswordSecure = $null

switch ($PSCmdlet.ParameterSetName) {

    'ClixmlFile' {
        if (-not (Test-Path -LiteralPath $CredentialPath)) {
            Write-MailLog "ERROR: credential file not found: $CredentialPath"
            Write-Error "Credential file not found: $CredentialPath. Run Set-RelayCredential.ps1 first."
            exit 1
        }
        try {
            $credential = Import-Clixml -LiteralPath $CredentialPath
            $smtpUsername = $credential.UserName
            $smtpPasswordSecure = $credential.Password
        }
        catch {
            Write-MailLog "ERROR: could not decrypt credential file: $($_.Exception.Message)"
            Write-Error "Could not decrypt the credential file. It is bound to the user account and machine it was created on - re-run Set-RelayCredential.ps1 there if needed."
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

    'Anonymous' {
        # No credential is loaded; the message is sent unauthenticated.
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
    if ($TlsMode -eq 'StartTls') {
        # TLS 1.2 always; add TLS 1.3 only where the platform defines it.
        $tlsProtocols = [Net.SecurityProtocolType]::Tls12
        if ([Enum]::GetNames([Net.SecurityProtocolType]) -contains 'Tls13') {
            $tlsProtocols = $tlsProtocols -bor [Net.SecurityProtocolType]::Tls13
        }
        [Net.ServicePointManager]::SecurityProtocol = $tlsProtocols
    }

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
    $smtp.EnableSsl = ($TlsMode -eq 'StartTls')
    if ($PSCmdlet.ParameterSetName -eq 'Anonymous') {
        $smtp.UseDefaultCredentials = $false
        $smtp.Credentials = $null
    }
    else {
        $smtp.Credentials = New-Object System.Net.NetworkCredential(
            $smtpUsername,
            (ConvertFrom-SecureStringPlain -SecureString $smtpPasswordSecure)
        )
    }

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