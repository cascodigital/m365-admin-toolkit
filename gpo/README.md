# GPO backup: LOGS

Group Policy backup (GPMC format) that prepares domain computers for the scripts in
`audit/` (`Get-LogonEvents.ps1`, `Search-Events.ps1`). It enables security auditing
and opens remote management access (WinRM, Remote Event Log Management, WMI).

All environment identifiers were replaced with placeholders: domain `corp.local`
(NetBIOS `CORP`), domain controller `DC01.corp.local`, domain GUID
`{00000000-0000-0000-0000-000000000000}` and a zeroed domain SID in the security
descriptors. The built-in group names in the backup are from a pt-BR domain
(`Administradores de empresa`, `Admins. do dominio`).

## What it configures (Computer Configuration only; User section is empty)

### Advanced Audit Policy (Success and Failure)

- Logon, Special Logon
- Security Group Management, User Account Management
- Process Creation
- Removable Storage
- Audit Policy Change, Authentication Policy Change
- Sensitive Privilege Use
- System Integrity

### Services

- WinRM: startup type Automatic

### Administrative Templates

- Windows Components > Windows Remote Management (WinRM) > WinRM Service >
  Allow remote server management through WinRM: Enabled, IPv4 filter `*`, IPv6 filter `*`
  (WinRM HTTP listener on all addresses)

### Windows Defender Firewall inbound rules (Allow, TCP)

| Rule | Port / service | Profiles |
|---|---|---|
| Remote Event Log Management (RPC-EPMAP) | RPC endpoint mapper (135) | All |
| Remote Event Log Management (NP-In) | 445 (SMB named pipes) | All |
| Remote Event Log Management (RPC) | Dynamic RPC (Eventlog service) | All |
| WMI (DCOM-In) | 135 | All |
| WMI (WMI-In) | Dynamic (winmgmt) | All |
| WMI (ASync-In) | Dynamic (unsecapp.exe) | All |
| Windows Remote Management (HTTP-In) | 5985 | Domain, Private (any address) |
| Windows Remote Management (HTTP-In) | 5985 | Public (local subnet only) |

## Security notes

- Port 5985 is also opened on the **Public** profile (restricted to the local subnet).
  Recommended: delete that rule and keep WinRM on the **Domain** profile only.
- The Event Log and WMI rules apply to all profiles. Consider limiting them to the
  Domain profile and restricting remote addresses to management hosts.
- WinRM listens on all addresses (`*`). Narrow the IPv4/IPv6 filters if possible.
- Increase the Security log size if you rely on these events; the defaults roll over quickly.

## Import

1. GPMC > Group Policy Objects > New (e.g. `LOGS`).
2. Right-click the GPO > Import Settings > select this `gpo` folder.
3. Review the settings (especially the firewall rules above), then link the GPO to the
   OUs that contain the computers to audit.
