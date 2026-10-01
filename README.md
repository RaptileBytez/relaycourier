# RelayCourier

**Send mail from scripts and scheduled tasks through any SMTP relay, with
secure credential storage.**

RelayCourier is a pair of PowerShell scripts that send email through the
**Azure Communication Services (ACS) SMTP relay** (default) or any other
SMTP server, with or without authentication. They are built for batch files
and scheduled tasks (including as a replacement for `blat.exe`; see the
migration notes below), with predictable exit codes. Credentials are never
hard-coded or stored in plain text; two interchangeable storage back ends
are supported.

## Files

| File | Purpose |
|---|---|
| `Set-RelayCredential.ps1` | One-time setup: stores the SMTP credential (encrypted file **or** Azure Key Vault). |
| `Send-RelayMail.ps1` | Sends an email. Mirrors the parameters you'd pass to `blat.exe` (from, to, subject, body, attachments, …). |

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
   **client secret** (from step 2). These are what `Set-RelayCredential.ps1`
   stores — never the raw application/client ID.

Default server/port for `Send-RelayMail.ps1`: `smtp.azurecomm.net`, port
`587`, STARTTLS (`EnableSsl = $true`). Only TLS 1.2 or TLS 1.3 is
allowed (1.3 only when the .NET runtime supports it); the script sets this
process-wide for the PowerShell session it runs in. Any other
SMTP server can be used with `-SmtpServer`, `-Port`, `-TlsMode` and
`-NoAuth` (see section 4). Implicit TLS (SMTPS, port 465) is not
supported, because `System.Net.Mail` cannot do it; ACS itself only offers
587 and 25 with STARTTLS.

## 2. Choosing a credential storage back end

Both scripts support two mutually exclusive modes, selected by which
parameters you pass (PowerShell parameter sets). `Send-RelayMail.ps1`
additionally has a third, credential-free `Anonymous` mode (`-NoAuth`,
see section 4):

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

## 3. `Set-RelayCredential.ps1` — one-time setup

The client secret is always entered interactively
(`Read-Host -AsSecureString`); it is never written to the script, to
console output, or to any file in plain text.

### Store in a local encrypted file

```powershell
.\Set-RelayCredential.ps1 -Username "<SMTP Username>" -Path .\relaycourier.cred.xml
```

Run this once per Windows user account and machine that will send mail.
If a scheduled task runs under a dedicated service account, run this
script once while logged in as (or impersonating) that account, on the
machine the task will run on.

### Store in Azure Key Vault

```powershell
Connect-AzAccount                      # or rely on an existing managed-identity/service-principal context
.\Set-RelayCredential.ps1 -Username "<SMTP Username>" `
    -VaultName "my-vault" -SecretName "relaycourier-secret"
```

This creates two secrets in the vault: `relaycourier-secret` (the client
secret/password) and `relaycourier-secret-Username` (the SMTP username). The
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

## 4. `Send-RelayMail.ps1` — sending mail

### Using the local encrypted file

```powershell
.\Send-RelayMail.ps1 -CredentialPath .\relaycourier.cred.xml `
    -From "noreply@example.com" -To "user@example.com" `
    -Subject "Test" -Body "This is a test."
```

### Using Azure Key Vault

```powershell
.\Send-RelayMail.ps1 -VaultName "my-vault" -SecretName "relaycourier-secret" `
    -From "noreply@example.com" -To "user@example.com" `
    -Subject "Test" -Body "This is a test."
```

### Full example: multiple recipients, CC, body file, attachments, log

```powershell
.\Send-RelayMail.ps1 -CredentialPath .\relaycourier.cred.xml `
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
.\Send-RelayMail.ps1 -CredentialPath .\relaycourier.cred.xml `
    -From "noreply@example.com" -To "a@example.com,b@example.com" `
    -Subject "Test" -Body "Test"
```

### Using any other SMTP server

```powershell
# Anonymous internal relay on port 25, no TLS, blat-style host:port
.\Send-RelayMail.ps1 -SmtpServer "smtp.example.local:25" -TlsMode None -NoAuth `
    -From "app@example.local" -To "ops@example.local" `
    -Subject "Test" -Body "This is a test."

# Custom server with STARTTLS and credentials
.\Send-RelayMail.ps1 -SmtpServer "mail.example.com" -Port 587 `
    -CredentialPath .\mail.cred.xml `
    -From "app@example.com" -To "ops@example.com" `
    -Subject "Test" -Body "This is a test."
```

- `-NoAuth` is its own parameter set and cannot be combined with
  `-CredentialPath`, `-VaultName`, `-SecretName` or `-UsernameSecretName`
  (PowerShell rejects the call).
- `-TlsMode None` is only allowed together with `-NoAuth`. Combined with
  `-CredentialPath` or `-VaultName`/`-SecretName` the script exits with
  `1` before loading any credential, because the password would travel
  in clear text. There is no override switch.
- Credentials created with `Set-RelayCredential.ps1` are ACS credentials.
  Do not point them at other servers via `-SmtpServer`; the server would
  receive your ACS SMTP username and client secret.
- `-SmtpServer host:port` (surrounding whitespace is trimmed) is split into
  host and port. The port must be numeric, from 1 to 65535, otherwise the
  script exits with `1`. Giving a `host:port` value together with an
  explicit `-Port` is an error (exit `1`): give the port in one place only.
  Values with more than one `:` (IPv6 literals) are not parsed; use
  `-Port` for them.
- If the server presents a certificate from a private CA or a self-signed
  one, install that certificate in the Windows certificate store of the
  sending machine. There is deliberately no switch to skip certificate
  validation.

### Migrating from `blat.exe`

| blat | `Send-RelayMail.ps1` |
|---|---|
| `-server host[:port]` | `-SmtpServer host[:port]` (or `-SmtpServer host -Port n`) |
| `-f` / `-t` / `-cc` / `-bcc` | `-From` / `-To` / `-Cc` / `-Bcc` (aliases `-MailFrom`, `-Recipient`, `-CopyTo`, `-BlindCopyTo`) |
| `-subject` / `-body` / `-bodyF` | `-Subject` / `-Body` / `-BodyFile` (aliases `-Title`, `-Message`, `-MessageFile`) |
| `-attach` / `-log` | `-Attachment` / `-LogFile` (aliases `-Attach`, `-Log`) |
| no `-u` / `-pw` (anonymous) | `-NoAuth` |
| `-u` / `-pw` | `-CredentialPath` or `-VaultName` / `-SecretName` (the password is never passed on the command line) |

The script is blat-inspired, not a drop-in clone: blat's single-letter
flags are not supported. Parameter abbreviations: `-T` no longer resolves
to `-To` (it is ambiguous with `-TlsMode` and `-Title`), so always write
`-To` in full.

### Parameter aliases

Existing parameter names stay the primary names. The aliases are
alternatives with clearer names:

| Parameter | Alias |
|---|---|
| `-SmtpServer` | `-MailServer` |
| `-From` | `-MailFrom` |
| `-To` | `-Recipient` |
| `-Cc` | `-CopyTo` |
| `-Bcc` | `-BlindCopyTo` |
| `-Subject` | `-Title` |
| `-Body` | `-Message` |
| `-BodyFile` | `-MessageFile` |
| `-Attachment` | `-Attach` |
| `-LogFile` | `-Log` |

### Parameters

| Parameter | Set | Required | Description |
|---|---|---|---|
| `-SmtpServer` | all | no | Default `smtp.azurecomm.net`. Accepts `host:port` (surrounding whitespace is trimmed); giving a port both here and via `-Port` is an error (exit 1). |
| `-Port` | all | no | Default `587`. |
| `-TlsMode` | all | no | `StartTls` (default) or `None` (plain SMTP). `None` is rejected (exit `1`) in combination with a credential; use it only with `-NoAuth`. |
| `-NoAuth` | Anonymous | yes | Switch: send without authentication. Cannot be combined with the ClixmlFile/KeyVault parameters. |
| `-CredentialPath` | ClixmlFile | yes | Path to the encrypted credential file. |
| `-VaultName` | KeyVault | yes | Azure Key Vault name. |
| `-SecretName` | KeyVault | yes | Secret name for the client secret/password. |
| `-UsernameSecretName` | KeyVault | no | Secret name for the username. Default: `<SecretName>-Username`. |
| `-From` | all | yes | Sender address. |
| `-To` | all | yes | One or more recipients (array or comma-separated string). |
| `-Cc` | all | no | Same format as `-To`. |
| `-Bcc` | all | no | Same format as `-To`. |
| `-Subject` | all | yes | Mail subject. |
| `-Body` | all | no | Inline body text. Ignored if `-BodyFile` is given. |
| `-BodyFile` | all | no | Path to a text file used as the body (blat.exe-style). |
| `-Html` | all | no | Switch: treat the body as HTML. |
| `-Attachment` | all | no | One or more file paths to attach. |
| `-LogFile` | all | no | Path to append a timestamped line per send attempt. |

### Exit codes (for batch files / scheduled tasks)

- `0` — email sent successfully.
- `1` — any failure (missing credential, invalid recipient/attachment
  path, SMTP error, etc.). A message is written with `Write-Error` and,
  if `-LogFile` is set, to the log.

Example batch-file check, same pattern as with `blat.exe`:

```bat
powershell -NoProfile -File Send-RelayMail.ps1 -CredentialPath relaycourier.cred.xml -From ... -To ... -Subject ... -Body ...
if %ERRORLEVEL% neq 0 (
    echo Mail send failed
)
```

## 5. Security notes

- **At rest:** the client secret is never embedded in script source. It
  lives either DPAPI-encrypted in a file scoped to one user+machine, or
  in Azure Key Vault with its own access control and audit log.
- **In transit:** STARTTLS on port 587 (`EnableSsl = $true`, TLS 1.2 or
  TLS 1.3 only, set process-wide for the PowerShell session) encrypts the
  SMTP session itself — this was already correct in the original script.
  With `-TlsMode None` nothing is encrypted and message content is
  readable on the network. To protect the password, the script refuses
  `-TlsMode None` in combination with a credential (exit `1`, before any
  credential is loaded); it is only accepted with `-NoAuth`, and should
  be used only on trusted internal networks. Older protocols (SSL 3.0,
  TLS 1.0, TLS 1.1) are never enabled, and server certificates are always
  validated.
- **In memory:** the password is kept as a `SecureString` for as long as
  possible and only converted to a plain string at the point
  `System.Net.NetworkCredential` requires it (a limitation of
  `System.Net.Mail`, which has no `SecureString`-based credential type).
- Rotating the client secret: generate a new secret in Entra ID, then
  re-run `Set-RelayCredential.ps1` (either back end) with the new value —
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
| `Could not decrypt the credential file` | The `.cred.xml` file was copied to another machine or is being read under a different user account than the one that created it. Re-run `Set-RelayCredential.ps1` on that machine/account. |
| `Could not read secrets from Key Vault` | No authenticated Az session in the current process, or the identity lacks `get` permission on the secrets. Run `Connect-AzAccount` (or verify the managed identity/service principal) and check the vault's access policy/RBAC. |
| `Az.KeyVault module not installed` | Run `Install-Module Az.KeyVault -Scope CurrentUser`. |
| `Server does not support secure connections` | The target server does not offer STARTTLS on that port (typical for internal relays on port 25). Use a port that offers STARTTLS, or, for an anonymous internal relay, `-TlsMode None -NoAuth`. |
| `-TlsMode None cannot be combined with a credential` | Plain SMTP would send the password in clear text. Use STARTTLS (the default) or `-NoAuth`. |
| `Invalid -SmtpServer value` | A `host:port` value with a non-numeric or out-of-range port. Fix the value or pass `-Port` separately. |
| `Give the port either in -SmtpServer` | Both `host:port` and `-Port` were given. Use only one. |
| `Parameter set cannot be resolved` | `-NoAuth` was combined with `-CredentialPath`, `-VaultName` or other credential parameters. Use one or the other. |
| Attachment/body-file errors | Path is checked with `Test-Path` before sending; verify the path is correct relative to the working directory the script runs from (e.g. a scheduled task's working directory is not always what you expect). |

## 7. License

MIT, see [LICENSE](LICENSE).
