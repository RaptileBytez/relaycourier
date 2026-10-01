# ACS SMTP Mailer — blat.exe replacement

PowerShell scripts for sending email through **Azure Communication
Services (ACS) SMTP relay**, written as a drop-in-style replacement for
`blat.exe` in batch files and scheduled tasks. Credentials are never
hard-coded or stored in plain text; two interchangeable storage back ends
are supported.

## Files

| File | Purpose |
|---|---|
| `Set-AcsSmtpSecret.ps1` | One-time setup: stores the SMTP credential (encrypted file **or** Azure Key Vault). |
| `Send-AcsMail.ps1` | Sends an email. Mirrors the parameters you'd pass to `blat.exe` (from, to, subject, body, attachments, …). |

## 1. Background: what the credential actually is

Azure Communication Services SMTP relay authenticates with **Basic Auth**
over TLS (STARTTLS on port 587), but the two values are *not* what the
original script assumed:

| Value | What it really is |
|---|---|
| **Username** | An **"SMTP Username"** resource created in the Azure Portal and linked to a Microsoft Entra application. It is **not** the Entra application (client) ID. |
| **Password** | The **client secret** of that linked Microsoft Entra application. |

### One-time Azure setup (reference)

1. In the ACS resource, make sure an Email domain is provisioned and
   connected.
2. Create a Microsoft Entra **application registration** and generate a
   **client secret** for it.
3. Assign the app the **"Communication and Email Service Owner"** role on
   the ACS resource (Access control (IAM)).
4. In the ACS resource, under **SMTP Username**, create a new SMTP
   username and link it to that Entra application.
5. You now have two values: the **SMTP Username** (from step 4) and the
   **client secret** (from step 2). These are what `Set-AcsSmtpSecret.ps1`
   stores — never the raw application/client ID.

Server/port for all scripts here: `smtp.azurecomm.net`, port `587`,
STARTTLS (`EnableSsl = $true`). TLS 1.2 is enforced explicitly in the
script.

## 2. Choosing a credential storage back end

Both scripts support two mutually exclusive modes, selected by which
parameters you pass (PowerShell parameter sets):

### Option A — `ClixmlFile` (local, no extra dependencies)

The credential is written to a file with `Export-Clixml`. PowerShell
encrypts the password inside that file using the **Windows Data
Protection API (DPAPI)**. The file can only be decrypted by **the same
Windows user account on the same machine** that created it.

Pros: no extra module, no Azure role assignment needed, works offline.
Cons: tied to one machine/account; if the script runs under a service
account or on several servers, you must run the setup once per
account/machine.

### Option B — `KeyVault` (centralized)

The SMTP username and client secret are stored as two secrets in an
**Azure Key Vault**. Any machine or identity with `get` permission on
those secrets can retrieve them — no local file to distribute or
regenerate per machine.

Requirements:
- The `Az.KeyVault` PowerShell module (`Install-Module Az.KeyVault -Scope CurrentUser`).
- An already-authenticated Az session in the process that runs the
  script: interactive `Connect-AzAccount`, a VM/Automation **managed
  identity**, or a service principal login. The scripts do not perform
  login themselves — they use whatever Az context is already active.
- `get`/`set` access policy (or RBAC role `Key Vault Secrets
  User`/`Key Vault Secrets Officer`) on the vault for the identity that
  runs the script.

Pros: single source of truth, easy credential rotation, works across
many machines/accounts, no local secret file to manage.
Cons: requires network access to Azure and an authenticated Az session at
runtime.

You can set up and use **both** at the same time (e.g. a local file for a
workstation, Key Vault for production servers) — the two scripts simply
treat them as alternative parameter sets.

## 3. `Set-AcsSmtpSecret.ps1` — one-time setup

The client secret is always entered interactively
(`Read-Host -AsSecureString`); it is never written to the script, to
console output, or to any file in plain text.

### Store in a local encrypted file

```powershell
.\Set-AcsSmtpSecret.ps1 -Username "<SMTP Username>" -Path .\acs-smtp.cred.xml
```

Run this once per Windows user account and machine that will send mail.
If a scheduled task runs under a dedicated service account, run this
script once while logged in as (or impersonating) that account, on the
machine the task will run on.

### Store in Azure Key Vault

```powershell
Connect-AzAccount                      # or rely on an existing managed-identity/service-principal context
.\Set-AcsSmtpSecret.ps1 -Username "<SMTP Username>" `
    -VaultName "my-vault" -SecretName "acs-smtp-secret"
```

This creates two secrets in the vault: `acs-smtp-secret` (the client
secret/password) and `acs-smtp-secret-Username` (the SMTP username). The
username secret name can be overridden with `-UsernameSecretName`.

### Parameters

| Parameter | Set | Required | Description |
|---|---|---|---|
| `-Username` | both | yes | The ACS "SMTP Username" (not the Entra client ID). |
| `-Path` | ClixmlFile | yes | Output path for the encrypted credential file. |
| `-VaultName` | KeyVault | yes | Azure Key Vault name. |
| `-SecretName` | KeyVault | yes | Secret name for the client secret/password. |
| `-UsernameSecretName` | KeyVault | no | Secret name for the username. Default: `<SecretName>-Username`. |

Exit code: `0` on success, `1` on failure.

## 4. `Send-AcsMail.ps1` — sending mail

### Using the local encrypted file

```powershell
.\Send-AcsMail.ps1 -CredentialPath .\acs-smtp.cred.xml `
    -From "noreply@example.com" -To "user@example.com" `
    -Subject "Test" -Body "This is a test."
```

### Using Azure Key Vault

```powershell
.\Send-AcsMail.ps1 -VaultName "my-vault" -SecretName "acs-smtp-secret" `
    -From "noreply@example.com" -To "user@example.com" `
    -Subject "Test" -Body "This is a test."
```

### Full example: multiple recipients, CC, body file, attachments, log

```powershell
.\Send-AcsMail.ps1 -CredentialPath .\acs-smtp.cred.xml `
    -From "noreply@example.com" `
    -To "a@example.com","b@example.com" `
    -Cc "manager@example.com" `
    -Subject "Report" `
    -BodyFile .\mailbody.txt `
    -Attachment ".\report.pdf", ".\log.txt" `
    -LogFile .\mail.log
```

Comma-separated address strings also work (useful when called from a
`.cmd`/`.bat` file where building a PowerShell array is awkward):

```powershell
.\Send-AcsMail.ps1 -CredentialPath .\acs-smtp.cred.xml `
    -From "noreply@example.com" -To "a@example.com,b@example.com" `
    -Subject "Test" -Body "Test"
```

### Parameters

| Parameter | Set | Required | Description |
|---|---|---|---|
| `-SmtpServer` | both | no | Default `smtp.azurecomm.net`. |
| `-Port` | both | no | Default `587`. |
| `-CredentialPath` | ClixmlFile | yes | Path to the encrypted credential file. |
| `-VaultName` | KeyVault | yes | Azure Key Vault name. |
| `-SecretName` | KeyVault | yes | Secret name for the client secret/password. |
| `-UsernameSecretName` | KeyVault | no | Secret name for the username. Default: `<SecretName>-Username`. |
| `-From` | both | yes | Sender address. |
| `-To` | both | yes | One or more recipients (array or comma-separated string). |
| `-Cc` | both | no | Same format as `-To`. |
| `-Bcc` | both | no | Same format as `-To`. |
| `-Subject` | both | yes | Mail subject. |
| `-Body` | both | no | Inline body text. Ignored if `-BodyFile` is given. |
| `-BodyFile` | both | no | Path to a text file used as the body (blat.exe-style). |
| `-Html` | both | no | Switch: treat the body as HTML. |
| `-Attachment` | both | no | One or more file paths to attach. |
| `-LogFile` | both | no | Path to append a timestamped line per send attempt. |

### Exit codes (for batch files / scheduled tasks)

- `0` — email sent successfully.
- `1` — any failure (missing credential, invalid recipient/attachment
  path, SMTP error, etc.). A message is written with `Write-Error` and,
  if `-LogFile` is set, to the log.

Example batch-file check, same pattern as with `blat.exe`:

```bat
powershell -NoProfile -File Send-AcsMail.ps1 -CredentialPath acs-smtp.cred.xml -From ... -To ... -Subject ... -Body ...
if %ERRORLEVEL% neq 0 (
    echo Mail send failed
)
```

## 5. Security notes

- **At rest:** the client secret is never embedded in script source. It
  lives either DPAPI-encrypted in a file scoped to one user+machine, or
  in Azure Key Vault with its own access control and audit log.
- **In transit:** STARTTLS on port 587 (`EnableSsl = $true`, TLS 1.2
  enforced) encrypts the SMTP session itself — this was already correct
  in the original script.
- **In memory:** the password is kept as a `SecureString` for as long as
  possible and only converted to a plain string at the point
  `System.Net.NetworkCredential` requires it (a limitation of
  `System.Net.Mail`, which has no `SecureString`-based credential type).
- Rotating the client secret: generate a new secret in Entra ID, then
  re-run `Set-AcsSmtpSecret.ps1` (either back end) with the new value —
  no script changes needed.
- `System.Net.Mail` (used here for attachment support and parity with
  the original script) is marked legacy by Microsoft in favor of
  libraries such as MailKit. It still works correctly with ACS's Basic
  Auth SMTP and needs no OAuth/XOAUTH2 flow, since Basic Auth (SMTP
  Username + client secret) is the supported mechanism for this relay —
  but if you later need OAuth-based SMTP auth or want to move off a
  legacy API, MailKit is the recommended migration path.

## 6. Troubleshooting

| Symptom | Likely cause |
|---|---|
| `5.7.3 Authentication unsuccessful` | Username is the Entra client ID instead of the ACS "SMTP Username", or the client secret expired/was rotated. |
| `Could not decrypt the credential file` | The `.cred.xml` file was copied to another machine or is being read under a different user account than the one that created it. Re-run `Set-AcsSmtpSecret.ps1` on that machine/account. |
| `Could not read secrets from Key Vault` | No authenticated Az session in the current process, or the identity lacks `get` permission on the secrets. Run `Connect-AzAccount` (or verify the managed identity/service principal) and check the vault's access policy/RBAC. |
| `Az.KeyVault module not installed` | Run `Install-Module Az.KeyVault -Scope CurrentUser`. |
| Attachment/body-file errors | Path is checked with `Test-Path` before sending; verify the path is correct relative to the working directory the script runs from (e.g. a scheduled task's working directory is not always what you expect). |
