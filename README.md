# M365 Admin Toolkit

![Status](https://img.shields.io/badge/Status-Active-brightgreen)
![License](https://img.shields.io/badge/License-MIT-blue)
![Author](https://img.shields.io/badge/Author-Casco%20Digital-orange)

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-5391FE?style=flat-square&logo=powershell&logoColor=white)
![Microsoft Graph](https://img.shields.io/badge/Microsoft_Graph-SDK-00D9FF?style=flat-square&logo=microsoft&logoColor=white)
![Exchange Online](https://img.shields.io/badge/Exchange-Online-0078D4?style=flat-square&logo=microsoftexchange&logoColor=white)

Scripts PowerShell para administracao de tenants Microsoft 365. Voltado para MSPs e admins que precisam de ferramentas prontas para auditoria, seguranca e gestao de usuarios.

## Estrutura

```
m365-admin-toolkit/
├── exchange/          # Gestao de correio e aliases
├── security/          # MFA, senhas e Security Defaults
├── audit/             # Logs de logon, busca de arquivos e eventos
├── sharepoint/        # Gestao e sincronizacao de SharePoint/OneDrive
├── tools/             # Utilitarios (keyboard, ping, Office removal)
└── gpo/               # GPO exportada para auditoria de logons (importavel)
```

## Exchange

| Script | Descricao |
|--------|-----------|
| `Set-CatchAllMailbox.ps1` | Cria grupo dinamico + regra catch-all para emails sem destino |
| `Remove-Email.ps1` | Search & Purge de mensagens especificas no tenant |
| `Set-MailAlias.ps1` | Gerencia aliases e habilita SendFromAliasEnabled |
| `Get-Emails.ps1` | Lista todos os enderecos do tenant (usuarios, grupos, aliases) |

## Security

| Script | Descricao |
|--------|-----------|
| `Set-M365Passwords.ps1` | Gera e aplica senhas seguras em massa com export CSV |
| `Manage-SecurityDefaults-SMTP.ps1` | Toggle de Security Defaults (MFA) + SMTP AUTH por usuario |
| `Get-MfaComplianceReport.ps1` | Relatorio de conformidade MFA filtrando usuarios reais |

## Audit

| Script | Descricao |
|--------|-----------|
| `Get-LogonEvents.ps1` | Analise forense de eventos 4624 em multiplos hosts |
| `Search-Files.ps1` | Busca arquivos em OneDrive por nome, extensao ou conteudo |
| `Search-Events.ps1` | Analisa logs Windows por Event ID com export Excel |

## SharePoint

| Script | Descricao |
|--------|-----------|
| `Sync-SharePointDeletion.ps1` | Deleta do SharePoint arquivos ausentes na fonte local. Util apos migracoes via SharePoint Migration Tool para manter destino em sincronia com a origem (pasta local, share de rede, OneDrive etc). Possui modo simulacao antes de executar de verdade. |

## Tools

| Script | Descricao |
|--------|-----------|
| `Fix-KeyboardLayout.ps1` | Forca ABNT2 e desativa hotkeys de troca de layout |
| `Monitor-Ping.ps1` | Monitor de latencia ICMP com gravacao CSV em tempo real |
| `Remove-Office.ps1` | Remocao completa de instalacoes Office |
| `Show-DiskUsage.ps1` | GUI estilo TreeSize/WizTree: tamanho de pastas/arquivos (maior no topo, % do pai), arvore navegavel e delete via botao direito. Auto-eleva para Admin. Sem instalar software de terceiros. |

## GPO

A pasta `gpo/` contem um backup de GPO para auditoria de logons, pronto para importar via `Import-GPO` em qualquer dominio. Sem dados de cliente.

## Requisitos

- PowerShell 5.1+ (recomendado 7+)
- Modulos: `Microsoft.Graph`, `ExchangeOnlineManagement`, `ImportExcel`
- Permissoes adequadas (Global Admin / Exchange Admin / Password Admin)

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
```

## Uso

```powershell
git clone https://github.com/cascodigital/m365-admin-toolkit.git
cd m365-admin-toolkit

# Execute o script desejado
PowerShell -ExecutionPolicy Bypass -File .\exchange\Set-CatchAllMailbox.ps1
```

Todos os scripts possuem prompts interativos e documentacao interna.

## Avisos

- **`Set-M365Passwords.ps1`** gera CSV com senhas em texto claro — armazene com seguranca
- **`Remove-Email.ps1`** executa purge irreversivel — teste antes em ambiente controlado
- **`Sync-SharePointDeletion.ps1`** deleta arquivos permanentemente — sempre rode em modo simulacao primeiro (`$ModoReal = $false`)
- **`Remove-Office.ps1`** faz remocao agressiva — avise usuarios antes de executar
- **`Show-DiskUsage.ps1`** o delete via botao direito e permanente (`Remove-Item -Recurse -Force`, sem Lixeira) — confirme o caminho antes

---

Desenvolvido com 🐢 (e cafe) por **Casco Digital**.
