# Parallels RAS 21.2 PowerShell Toolkit

Forty small administration scripts for the `RASAdmin` API v5 module.

## Authentication

Every numbered script loads `_Connect-RASFarm.ps1`. This private helper defines the `Connect-RASFarm` function, imports the `RASAdmin` module, collects connection details, and opens the RAS administrative session. The numbered script closes that session in a `finally` block when its work finishes.

Do not run `_Connect-RASFarm.ps1` directly; it only defines the helper function.

### Interactive mode

Without `-UseEnvironment`, the helper prompts for the Connection Broker, administrator UPN, and secure password:

```powershell
.\02-Get-RASAgentHealth.ps1
```

Supply `-Server` to skip only the Connection Broker prompt. The administrator UPN and password are still requested interactively:

```powershell
.\02-Get-RASAgentHealth.ps1 -Server ras-broker.example.com
```

### Environment-variable mode

With `-UseEnvironment`, the helper reads all three values from `RAS_SERVER`, `RAS_USERNAME`, and `RAS_PASSWORD`. It stops with an error if any value is missing:

For unattended use, set process-scoped environment variables and add `-UseEnvironment`:

```powershell
$env:RAS_SERVER = 'ras-broker.example.com'
$env:RAS_USERNAME = 'administrator@example.com'
$env:RAS_PASSWORD = Read-Host 'RAS password'
.\02-Get-RASAgentHealth.ps1 -UseEnvironment
```

`RAS_PASSWORD` is plain text in the process environment. Prefer interactive mode unless unattended execution is required, and clear it afterward.

```powershell
Remove-Item Env:RAS_PASSWORD
```

## Scripts

1. `01-Get-RASFarmOverview.ps1` — version, sites, and farm settings.
2. `02-Get-RASAgentHealth.ps1` — all agent versions and states.
3. `03-Get-RASSiteHealth.ps1` — site health.
4. `04-Get-RASBrokerHealth.ps1` — Connection Broker health.
5. `05-Get-RASGatewayHealth.ps1` — Secure Gateway health.
6. `06-Get-RASRDSHostHealth.ps1` — RDS host health and load.
7. `07-Get-RASProviderHealth.ps1` — provider health.
8. `08-Get-RASSessions.ps1` — sessions filtered by source and state.
9. `09-Get-RASDisconnectedSessions.ps1` — disconnected sessions.
10. `10-Get-RASPublishedResources.ps1` — published resources.
11. `11-Get-RASMFAUsers.ps1` — enrolled TOTP and Email OTP users.
12. `12-Get-RASCertificates.ps1` — farm certificates.
13. `13-Get-RASLicense.ps1` — license information.
14. `14-Get-RASAdministrators.ps1` — RAS administrator accounts.
15. `15-Get-RASClientPolicies.ps1` — client policies.
16. `16-Get-RASReportingConfiguration.ps1` — farm reporting settings.
17. `17-Export-RASConfiguration.ps1` — timestamped `.dat2` configuration backup.
18. `18-Invoke-RASApplyChanges.ps1` — apply pending farm changes with confirmation.
19. `19-Reset-RASMFAUser.ps1` — reset selected MFA enrollments with confirmation.
20. `20-Clear-RASSessionCache.ps1` — force clients to reauthenticate, with confirmation.
21. `21-Get-RASCertificateExpiry.ps1` — certificates approaching expiration.
22. `22-Test-RASGatewayConnectivity.ps1` — DNS and configurable TCP checks.
23. `23-Compare-RASAgentVersions.ps1` — group agents by installed version.
24. `24-Update-RASAgent.ps1` — update a selected agent with confirmation.
25. `25-Get-RASHostCapacity.ps1` — RDS load and session counts.
26. `26-Set-RASRDSHostAvailability.ps1` — enable or disable an RDS host safely.
27. `27-Logoff-RASDisconnectedSessions.ps1` — log off old disconnected RDS sessions.
28. `28-Send-RASSessionMessage.ps1` — message a user's RDS sessions.
29. `29-Get-RASUserSessions.ps1` — find sessions for a user.
30. `30-Get-RASPublishedResourceAccess.ps1` — effective access for a user (RAS 21.2).
31. `31-Find-RASDisabledResources.ps1` — disabled or maintenance publications.
32. `32-Export-RASPublishedResources.ps1` — publication inventory CSV.
33. `33-Export-RASFarmHealth.ps1` — combined HTML health report.
34. `34-Compare-RASSites.ps1` — compare component counts between two sites.
35. `35-Test-RASMFAUserADAccess.ps1` — validate that a Connection Broker can read a user and optional AD custom attribute. This is not a general MFA or RADIUS connectivity test and is normally not applicable to standard TOTP providers.
36. `36-Get-RASMFAConfiguration.ps1` — MFA providers and criteria.
37. `37-Get-RASLogonHourRules.ps1` — logon-hours inventory.
38. `38-Get-RASTrustedDomains.ps1` — trusted domains and AD integration.
39. `39-Get-RASNotifications.ps1` — notification inventory.
40. `40-Invoke-RASDailyAudit.ps1` — compact daily exception report.

### MFA AD access and RADIUS examples

Validate that the Connection Broker can access an AD user:

```powershell
.\35-Test-RASMFAUserADAccess.ps1 -UseEnvironment `
    -UserPrincipalName 'user01@example.com' `
    -SiteId 1
```

Validate access to the AD attribute configured for an MFA workflow:

```powershell
.\35-Test-RASMFAUserADAccess.ps1 -UseEnvironment `
    -UserPrincipalName 'user01@example.com' `
    -SiteId 1 `
    -ADCustomAttribute 'extensionAttribute1'
```

Script 35 does not test a RADIUS server. From an active RAS session, use the native connection test instead:

```powershell
Invoke-RASMFA -CheckConnection `
    -Server 'radius.example.com' `
    -Port 1812 `
    -Timeout 5 `
    -Retries 2
```

Supply any additional parameters required by the configured RADIUS provider, such as its backup server, password encoding, shared secret, or attribute list.

### Certificate expiration report

Script 21 checks the next 30 days by default. Empty output means no certificates expire within that window. Use `-Days` to change it:

Parameters:

| Parameter | Required | Default | Purpose |
|---|---:|---|---|
| `-Server` | No | Prompted | RAS Connection Broker name or address. |
| `-UseEnvironment` | No | Off | Use `RAS_SERVER`, `RAS_USERNAME`, and `RAS_PASSWORD`. |
| `-Days` | No | `30` | Show certificates expiring within this many days. |

```powershell
.\21-Get-RASCertificateExpiry.ps1 -UseEnvironment -Days 365
```

Use script 12 to display every certificate regardless of expiration date:

```powershell
.\12-Get-RASCertificates.ps1 -UseEnvironment
```

The three write scripts support `-WhatIf`. Use it before making a change:

```powershell
.\19-Reset-RASMFAUser.ps1 -User user@example.com -Type GAuthTOTP -WhatIf
```

## Examples for all scripts

All examples use environment-variable authentication. Omit `-UseEnvironment` to enter the Connection Broker and credentials interactively.

### 01 — Farm overview

```powershell
.\01-Get-RASFarmOverview.ps1 -UseEnvironment
```

### 02 — Agent health

```powershell
.\02-Get-RASAgentHealth.ps1 -UseEnvironment
```

### 03 — Site health

```powershell
.\03-Get-RASSiteHealth.ps1 -UseEnvironment
```

### 04 — Connection Broker health

```powershell
.\04-Get-RASBrokerHealth.ps1 -UseEnvironment
```

### 05 — Secure Gateway health

```powershell
.\05-Get-RASGatewayHealth.ps1 -UseEnvironment
```

### 06 — RDS host health

```powershell
.\06-Get-RASRDSHostHealth.ps1 -UseEnvironment
```

### 07 — Provider health

```powershell
.\07-Get-RASProviderHealth.ps1 -UseEnvironment
```

### 08 — Filter sessions

```powershell
.\08-Get-RASSessions.ps1 -UseEnvironment -Source RDS -State Active
```

`-Source` accepts `RDS`, `VDI`, `AVD`, or `All`. `-State` accepts `Active`, `Connected`, `ConnectQuery`, `Shadow`, `Disconnected`, `Idle`, `Listen`, `Reset`, `Down`, `Init`, or `All`.

### 09 — Disconnected sessions

```powershell
.\09-Get-RASDisconnectedSessions.ps1 -UseEnvironment
```

### 10 — Published resources

```powershell
.\10-Get-RASPublishedResources.ps1 -UseEnvironment
```

### 11 — Enrolled MFA users

```powershell
.\11-Get-RASMFAUsers.ps1 -UseEnvironment
```

Lists enrolled TOTP and Email OTP users. It does not enumerate users held by an external RADIUS provider.

### 12 — Certificates

```powershell
.\12-Get-RASCertificates.ps1 -UseEnvironment
```

### 13 — License details

```powershell
.\13-Get-RASLicense.ps1 -UseEnvironment
```

### 14 — RAS administrators

```powershell
.\14-Get-RASAdministrators.ps1 -UseEnvironment
```

### 15 — Client policies

```powershell
.\15-Get-RASClientPolicies.ps1 -UseEnvironment
```

### 16 — Reporting configuration

```powershell
.\16-Get-RASReportingConfiguration.ps1 -UseEnvironment
```

### 17 — Export farm configuration

```powershell
.\17-Export-RASConfiguration.ps1 -UseEnvironment `
    -Path 'C:\RAS-Backups\Farm-Backup.dat2' `
    -TimeoutInSecs 900
```

The destination folder must already exist.

### 18 — Apply pending changes

Preview first, then remove `-WhatIf` to execute:

```powershell
.\18-Invoke-RASApplyChanges.ps1 -UseEnvironment -WhatIf
```

### 19 — Reset an MFA enrollment

```powershell
.\19-Reset-RASMFAUser.ps1 -UseEnvironment `
    -User 'user01@example.com' `
    -Type GAuthTOTP `
    -SiteId 1 `
    -WhatIf
```

For Email OTP, also pass its provider ID with `-MFAId`.

### 20 — Clear the session ID cache

This forces clients to authenticate again:

```powershell
.\20-Clear-RASSessionCache.ps1 -UseEnvironment -WhatIf
```

### 21 — Certificate expiration

```powershell
.\21-Get-RASCertificateExpiry.ps1 -UseEnvironment -Days 90
```

Empty output means no certificate expires within the selected window.

### 22 — Gateway connectivity

```powershell
.\22-Test-RASGatewayConnectivity.ps1 -UseEnvironment -Port 443, 20009
```

Choose ports that apply to your deployment; the script does not assume a complete RAS firewall profile.

### 23 — Compare agent versions

```powershell
.\23-Compare-RASAgentVersions.ps1 -UseEnvironment
```

### 24 — Update an agent

```powershell
.\24-Update-RASAgent.ps1 -UseEnvironment `
    -AgentServer 'rds01.example.com' `
    -SiteId 1 `
    -WhatIf
```

Add `-Force` only when a normal update cannot proceed.

### 25 — RDS host capacity

```powershell
.\25-Get-RASHostCapacity.ps1 -UseEnvironment
```

### 26 — Enable or disable an RDS host

Disable a host after previewing the change:

```powershell
.\26-Set-RASRDSHostAvailability.ps1 -UseEnvironment `
    -RDSHost 'rds01.example.com' `
    -SiteId 1 `
    -Enabled $false `
    -WhatIf
```

Use `-Enabled $true` to return it to service.

### 27 — Log off old disconnected sessions

```powershell
.\27-Logoff-RASDisconnectedSessions.ps1 -UseEnvironment `
    -IdleHours 12 `
    -WhatIf
```

### 28 — Send a session message

```powershell
.\28-Send-RASSessionMessage.ps1 -UseEnvironment `
    -User 'user01@example.com' `
    -Title 'Maintenance' `
    -Message 'Please save your work and sign out by 18:00.' `
    -WhatIf
```

### 29 — Find a user's sessions

```powershell
.\29-Get-RASUserSessions.ps1 -UseEnvironment -User 'user01@example.com'
```

### 30 — Published-resource effective access

```powershell
.\30-Get-RASPublishedResourceAccess.ps1 -UseEnvironment `
    -User 'user01@example.com' `
    -ClientDeviceOS Windows `
    -SiteId 1
```

Optional access context includes `-ClientDeviceName`, `-PrivateIPAddress`, `-HardwareID`, `-GatewayIP`, and `-ThemeId`.

List valid client OS values with:

```powershell
[enum]::GetNames((Get-Command Get-RASPubEffectiveAccess).Parameters.ClientDeviceOS.ParameterType)
```

### 31 — Disabled published resources

```powershell
.\31-Find-RASDisabledResources.ps1 -UseEnvironment
```

### 32 — Export published resources

```powershell
.\32-Export-RASPublishedResources.ps1 -UseEnvironment `
    -Path 'C:\RAS-Reports\PublishedResources.csv'
```

The destination folder must already exist.

### 33 — Export farm-health HTML

```powershell
.\33-Export-RASFarmHealth.ps1 -UseEnvironment `
    -Path 'C:\RAS-Reports\FarmHealth.html'
```

### 34 — Compare two sites

```powershell
.\34-Compare-RASSites.ps1 -UseEnvironment -SiteId1 1 -SiteId2 2
```

Both IDs must identify existing RAS sites.

### 35 — Validate MFA AD access

```powershell
.\35-Test-RASMFAUserADAccess.ps1 -UseEnvironment `
    -UserPrincipalName 'user01@example.com' `
    -SiteId 1 `
    -ADCustomAttribute 'extensionAttribute1'
```

Use this only when the MFA workflow reads an AD custom attribute. It is not a TOTP enrollment or RADIUS connectivity test.

### 36 — MFA configuration

```powershell
.\36-Get-RASMFAConfiguration.ps1 -UseEnvironment
```

### 37 — Logon-hour rules

```powershell
.\37-Get-RASLogonHourRules.ps1 -UseEnvironment
```

### 38 — Trusted domains and AD integration

```powershell
.\38-Get-RASTrustedDomains.ps1 -UseEnvironment
```

### 39 — Notifications

```powershell
.\39-Get-RASNotifications.ps1 -UseEnvironment
```

### 40 — Daily audit

```powershell
.\40-Invoke-RASDailyAudit.ps1 -UseEnvironment -CertificateWarningDays 60 |
    Format-List
```
