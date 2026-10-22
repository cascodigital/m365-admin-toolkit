# M365 Admin Toolkit

PowerShell scripts for Microsoft 365 support and administration, written for MSP work and day-to-day L1/L2 tickets. Each one exists because the portal or the native cmdlet left a gap: a limit, a blind spot or a manual step that kept coming back. MIT licensed.

Tested on and recommended for **PowerShell 7 (`pwsh`) on Windows**.

## Exchange Online

| Script | What it does |
|---|---|
| [Get-MessageTraceGUI](exchange/Get-MessageTraceGUI.ps1) | Message trace GUI with hop-by-hop detail. Searches up to 89 days in real time through `Get-MessageTraceV2`, with no Historical Search queue. Splits the range into windows and subdivides any window that hits the 5,000-row cap, so nothing goes missing without a warning. |
| [Invoke-MailFlowForensics](exchange/Invoke-MailFlowForensics.ps1) | Answers "my email went to quarantine" with an evidence package covering trace, quarantine, allow/block lists, junk settings and inbox rules. Correlates by MessageId, because quarantine indexes the header From and trace indexes the envelope sender, so searching by sender returns nothing for shared mailboxes, aliases and automations. |
| [Audit-InboxRules](exchange/Audit-InboxRules.ps1) | Read-only inbox rule audit. Finds what the Outlook rule list hides: rules made dead by an earlier rule with the same condition and *stop processing*, hidden user rules, rules that delete NDRs and quota warnings, corrupt rules, and external forwarding checked against the tenant's accepted domains. |
| [Remove-CorruptInboxRules](exchange/Remove-CorruptInboxRules.ps1) | Removes rules that `Get-InboxRule` reports only as a "contains errors" warning and never returns as objects. Counts the warnings before and after as proof. Takes the rule names as a parameter. Supports `-DryRun`. |
| [Get-MailboxFolderSize](exchange/Get-MailboxFolderSize.ps1) | Read-only diagnosis for a full mailbox. Shows why archiving is not freeing space (holds, MRM disabled, orphaned tags, overdue folders), estimates GB per day and projects each folder's size under a 30-day tag. |
| [Manage-ArchiveAutoExpand](exchange/Manage-ArchiveAutoExpand.ps1) | Checks licence, archive and current state, enables auto-expanding archive only when it can help, then re-queries to prove the change. Loops over several mailboxes. |
| [Manage-SharedMailbox](exchange/Manage-SharedMailbox.ps1) | Full Access, Send As and Send on Behalf in one menu, listing current holders before any change. |
| [Enable-SharedSentItemsCopy](exchange/Enable-SharedSentItemsCopy.ps1) | Makes a shared mailbox keep its own copy of mail sent as or on behalf of it. Refuses non-shared targets, writes a before/after CSV, simulates by default. |
| [Set-CatchAllMailbox](exchange/Set-CatchAllMailbox.ps1) | Catch-all for one or more domains: InternalRelay, a dynamic exception group covering every mail-enabled recipient type (users, shared mailboxes, groups, contacts) and a transport rule, so only mail to unknown addresses is redirected. Supports `-WhatIf`. |
| [Remove-Email](exchange/Remove-Email.ps1) | Search & Purge of a specific message across all mailboxes through Compliance Search. Shows the hit count, asks before purging, waits for the purge to finish and reports the result (soft delete, 10 items per mailbox per run). |

## Security

| Script | What it does |
|---|---|
| [Set-M365Passwords](security/Set-M365Passwords.ps1) | Bulk password reset by domain, with memorable passwords, throttling and a CSV of results. |

## Audit

| Script | What it does |
|---|---|
| [Get-LogonEvents](audit/Get-LogonEvents.ps1) | Human logons only (event 4624, types 2, 7, 10 and 11) on one machine or the whole domain, tolerating offline hosts. Excel output. |
| [Search-Events](audit/Search-Events.ps1) | Several event IDs in one pass, including archived and rotated `.evtx` files, not just the live log. Remote mode over the Event Log RPC service. One Excel sheet per ID. |
| [Search-Files](audit/Search-Files.ps1) | Finds files by name or wildcard in one user's OneDrive or across every OneDrive in a domain, through the Microsoft Graph drive search API with full paging. |

## SharePoint

| Script | What it does |
|---|---|
| [Sync-SharePointDeletion](sharepoint/Sync-SharePointDeletion.ps1) | After a SharePoint Migration Tool run, deletes from SharePoint whatever no longer exists at the source. SPMT copies but never deletes. Simulates by default; aborts if the source is unreachable or empty, and skips any library where planned deletions exceed a threshold (`-MaxDeletePercent`, `-MaxDelete`). |

## Endpoint

| Script | What it does |
|---|---|
| [Reset-Teams](tools/Reset-Teams.ps1) | Removes every trace of Teams (new, classic, Machine-Wide Installer, Outlook add-in, caches) and reinstalls the new client. First run writes a plan, second run applies it. Custom backgrounds are saved first. |
| [Reset-NewOutlook](tools/Reset-NewOutlook.ps1) | Same plan-then-apply removal and reinstall for the new Outlook (olk). |
| [Reset-OfficeTenantCache](tools/Reset-OfficeTenantCache.ps1) | Clears the user's Office, OneDrive and WAM identity cache when apps keep signing in to an old tenant after a migration. Office stays installed; Outlook data folders containing PST files are skipped. |
| [Install-Visio](tools/Install-Visio.ps1) | Installs Visio matching the installed Office bitness. Fetches the Office Deployment Tool from four sources and verifies the Microsoft signature before running it. |
| [Show-DiskUsage](tools/Show-DiskUsage.ps1) | TreeSize-style viewer with no install. Reports physical allocated size, so compression, sparse files, hard links and cloud-only OneDrive files are counted correctly. |
| [Fix-KeyboardLayout](tools/Fix-KeyboardLayout.ps1) | Forces the ABNT2 layout on English Windows and disables the Alt+Shift and Ctrl+Shift switch hotkeys. |

The `gpo/` folder holds a GPO backup (logon auditing plus WinRM and firewall rules for remote collection), ready for `Import-GPO`. See [gpo/README.md](gpo/README.md) before importing.

## Getting started

```powershell
Install-Module ExchangeOnlineManagement, Microsoft.Graph, ImportExcel -Scope CurrentUser

git clone https://github.com/cascodigital/m365-admin-toolkit.git
cd m365-admin-toolkit
pwsh -File .\exchange\Audit-InboxRules.ps1 -MailboxAddress user@contoso.com
```

Every script has built-in help: `Get-Help .\exchange\Audit-InboxRules.ps1 -Full`.

## Before you run

> [!WARNING]
> **Not every script previews first.** Check the safety mode below before running anything that changes the tenant or the machine.

| Safety mode | Scripts |
|---|---|
| **Simulates by default.** Apply with `$global:APPLY = $true; .\script.ps1` | Reset-Teams, Reset-NewOutlook, Enable-SharedSentItemsCopy |
| **Simulates by default.** Apply with `-Execute` | Sync-SharePointDeletion |
| **Supports `-WhatIf`** | Set-CatchAllMailbox, Reset-OfficeTenantCache, Sync-SharePointDeletion |
| **Applies by default.** Preview with `-DryRun` | Remove-CorruptInboxRules |
| **Asks for confirmation** before changing anything | Remove-Email, Set-M365Passwords, Manage-ArchiveAutoExpand, Manage-SharedMailbox |
| **Applies immediately, no preview** | Fix-KeyboardLayout, Install-Visio |
| **Read-only** | Every other script |

- **Bulk deletes:** `Remove-Email` purges the message from every mailbox (soft delete: only an admin can recover it, from Recoverable Items). `Sync-SharePointDeletion` deletes files in bulk; they go to the site recycle bin.
- `Set-M365Passwords` writes passwords to a plain-text CSV.
- Exchange scripts ask for the admin UPN at run time. Several need an active PIM role (Exchange Recipient Administrator); without it, Exchange reports existing mailboxes as not found.
