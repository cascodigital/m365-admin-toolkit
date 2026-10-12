<#
.SYNOPSIS
    Gerencia acesso a caixa de correio compartilhada no Exchange Online.

.DESCRIPTION
    Conecta ao EXO com a conta admin informada, pergunta qual a shared mailbox,
    lista quem tem Full Access / Send As / Send on Behalf, e oferece menu para
    adicionar (ou remover) usuarios em qualquer uma das tres permissoes.

    PRE-REQUISITO: PIM "Exchange Recipient Administrator" ATIVO.
    Sem PIM, o EXO mente e diz que a mailbox nao existe.

.NOTES
    Autor: Andre Kittler da Costa
.EXAMPLE
    .\Manage-SharedMailbox.ps1 -AdminUPN admin@contoso.com
    Pergunta a shared e abre o menu.
.EXAMPLE
    .\Manage-SharedMailbox.ps1 -AdminUPN admin@contoso.com -SharedMailbox financeiro@contoso.com
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$AdminUPN,
    [string]$SharedMailbox
)

$ErrorActionPreference = 'Stop'

#region ---------- Helpers ----------

function Write-Head {
    param([string]$Text)
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
}

function Write-Ok   { param($m) Write-Host "[ OK ] $m"  -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "[WARN] $m"  -ForegroundColor Yellow }
function Write-Err  { param($m) Write-Host "[FAIL] $m"  -ForegroundColor Red }
function Write-Info { param($m) Write-Host "[INFO] $m"  -ForegroundColor Gray }

function Connect-EXOIfNeeded {
    param([string]$UPN)

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "Modulo ExchangeOnlineManagement nao instalado. Rode: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $conn = Get-ConnectionInformation -ErrorAction SilentlyContinue |
            Where-Object { $_.State -eq 'Connected' }

    if ($conn) {
        Write-Ok "Ja conectado ao EXO como $($conn[0].UserPrincipalName)"
        return
    }

    Write-Info "Conectando ao Exchange Online como $UPN ..."
    Connect-ExchangeOnline -UserPrincipalName $UPN -ShowBanner:$false
    Write-Ok "Conectado."
}

# Resolve qualquer coisa (email, alias, nome) para um objeto de recipient
function Resolve-Recipient {
    param(
        [Parameter(Mandatory)][string]$Identity,
        [switch]$MailboxOnly
    )

    $r = $null
    try { $r = Get-Recipient -Identity $Identity -ErrorAction Stop } catch { }

    if (-not $r) {
        try {
            $r = @(Get-Recipient -ResultSize 20 -Filter "EmailAddresses -like '*$Identity*'" -ErrorAction Stop)
        } catch { }
    }

    if (-not $r) {
        try {
            $r = @(Get-Recipient -ResultSize 20 -Filter "DisplayName -like '*$Identity*'" -ErrorAction Stop)
        } catch { }
    }

    $r = @($r)
    if ($r.Count -eq 0) {
        Write-Err "Nada encontrado para '$Identity'. (Se TUDO esta dando not found -> PIM expirado.)"
        return $null
    }

    if ($r.Count -gt 1) {
        Write-Warn "Mais de um resultado para '$Identity':"
        for ($i = 0; $i -lt $r.Count; $i++) {
            "{0,3}) {1,-38} {2,-45} {3}" -f ($i + 1), $r[$i].DisplayName, $r[$i].PrimarySmtpAddress, $r[$i].RecipientTypeDetails |
                Write-Host
        }
        $pick = Read-Host "Escolha o numero (Enter = cancelar)"
        if ([string]::IsNullOrWhiteSpace($pick)) { return $null }
        $idx = 0
        if (-not [int]::TryParse($pick, [ref]$idx) -or $idx -lt 1 -or $idx -gt $r.Count) {
            Write-Err "Escolha invalida."
            return $null
        }
        $r = $r[$idx - 1]
    } else {
        $r = $r[0]
    }

    if ($MailboxOnly -and $r.RecipientTypeDetails -notlike '*Mailbox*') {
        Write-Err "'$($r.PrimarySmtpAddress)' e um $($r.RecipientTypeDetails), nao uma mailbox."
        return $null
    }

    return $r
}

#endregion

#region ---------- Leitura de permissoes ----------

# Os cmdlets REST do EXO V3 devolvem booleanos ora como [bool], ora como a string
# "False" -- e no PowerShell qualquer string nao-vazia e truthy. Normaliza aqui.
function ConvertTo-Bool {
    param($Value)
    if ($null -eq $Value)      { return $false }
    if ($Value -is [bool])     { return $Value }
    return ([string]$Value).Trim() -eq 'True'
}

function Get-SharedMailboxAccess {
    param([Parameter(Mandatory)]$Mailbox)

    $mbx = Get-Mailbox -Identity $Mailbox.Identity

    $full = @(
        Get-MailboxPermission -Identity $mbx.Identity |
        Where-Object {
            $_.User -notlike 'NT AUTHORITY\*' -and
            $_.User -notlike 'S-1-5-21-*'      -and
            -not (ConvertTo-Bool $_.IsInherited) -and
            $_.AccessRights -contains 'FullAccess'
        } |
        Select-Object @{n='User';e={[string]$_.User}},
                      @{n='Rights';e={($_.AccessRights -join ', ')}},
                      @{n='Deny';e={ConvertTo-Bool $_.Deny}}
    )

    $sendAs = @(
        Get-RecipientPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Trustee -notlike 'NT AUTHORITY\*' -and
            $_.AccessRights -contains 'SendAs'
        } |
        Select-Object @{n='User';e={[string]$_.Trustee}},
                      @{n='Deny';e={[string]$_.AccessControlType -eq 'Deny'}}
    )

    $onBehalf = @(
        foreach ($g in $mbx.GrantSendOnBehalfTo) {
            $res = Get-Recipient -Identity $g -ErrorAction SilentlyContinue
            [pscustomobject]@{
                User = if ($res) { [string]$res.PrimarySmtpAddress } else { [string]$g }
            }
        }
    )

    [pscustomobject]@{
        Mailbox      = $mbx
        FullAccess   = $full
        SendAs       = $sendAs
        SendOnBehalf = $onBehalf
    }
}

function Show-SharedMailboxAccess {
    param([Parameter(Mandatory)]$Access)

    $mbx = $Access.Mailbox

    Write-Head "CAIXA: $($mbx.DisplayName)  <$($mbx.PrimarySmtpAddress)>   [$($mbx.RecipientTypeDetails)]"

    Write-Host ''
    Write-Host '--- FULL ACCESS (abrir a caixa) ---' -ForegroundColor White
    if ($Access.FullAccess.Count -eq 0) {
        Write-Host '    (ninguem)' -ForegroundColor DarkGray
    } else {
        $Access.FullAccess | ForEach-Object {
            if ($_.Deny) {
                Write-Host ("    - {0} [DENY]" -f $_.User) -ForegroundColor Red
            } else {
                Write-Host ("    - {0}" -f $_.User) -ForegroundColor Gray
            }
        }
    }

    Write-Host ''
    Write-Host '--- SEND AS (envia COMO a caixa) ---' -ForegroundColor White
    if ($Access.SendAs.Count -eq 0) {
        Write-Host '    (ninguem)' -ForegroundColor DarkGray
    } else {
        $Access.SendAs | ForEach-Object {
            if ($_.Deny) {
                Write-Host ("    - {0} [DENY]" -f $_.User) -ForegroundColor Red
            } else {
                Write-Host ("    - {0}" -f $_.User) -ForegroundColor Gray
            }
        }
    }

    Write-Host ''
    Write-Host '--- SEND ON BEHALF (envia EM NOME DE) ---' -ForegroundColor White
    if ($Access.SendOnBehalf.Count -eq 0) {
        Write-Host '    (ninguem)' -ForegroundColor DarkGray
    } else {
        $Access.SendOnBehalf | ForEach-Object {
            Write-Host ("    - {0}" -f $_.User) -ForegroundColor Gray
        }
    }
    Write-Host ''
}

#endregion

#region ---------- Escrita de permissoes ----------

function Add-FullAccess {
    param($Mailbox, $User, [bool]$AutoMapping = $true)
    try {
        Add-MailboxPermission -Identity $Mailbox.Identity `
                              -User $User.Identity `
                              -AccessRights FullAccess `
                              -InheritanceType All `
                              -AutoMapping:$AutoMapping `
                              -Confirm:$false -ErrorAction Stop | Out-Null
        Write-Ok "Full Access concedido a $($User.PrimarySmtpAddress) (AutoMapping: $AutoMapping)"
    } catch {
        Write-Err "Full Access falhou: $($_.Exception.Message)"
    }
}

function Add-SendAs {
    param($Mailbox, $User)
    try {
        Add-RecipientPermission -Identity $Mailbox.Identity `
                                -Trustee $User.Identity `
                                -AccessRights SendAs `
                                -Confirm:$false -ErrorAction Stop | Out-Null
        Write-Ok "Send As concedido a $($User.PrimarySmtpAddress)"
    } catch {
        Write-Err "Send As falhou: $($_.Exception.Message)"
    }
}

function Add-SendOnBehalf {
    param($Mailbox, $User)
    try {
        Set-Mailbox -Identity $Mailbox.Identity `
                    -GrantSendOnBehalfTo @{Add = $User.Identity} `
                    -Confirm:$false -ErrorAction Stop
        Write-Ok "Send on Behalf concedido a $($User.PrimarySmtpAddress)"
    } catch {
        Write-Err "Send on Behalf falhou: $($_.Exception.Message)"
    }
}

function Remove-AnyAccess {
    param($Mailbox, $User, [ValidateSet('Full','SendAs','OnBehalf')]$Kind)
    try {
        switch ($Kind) {
            'Full' {
                Remove-MailboxPermission -Identity $Mailbox.Identity -User $User.Identity `
                    -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                Write-Ok "Full Access removido de $($User.PrimarySmtpAddress)"
            }
            'SendAs' {
                Remove-RecipientPermission -Identity $Mailbox.Identity -Trustee $User.Identity `
                    -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                Write-Ok "Send As removido de $($User.PrimarySmtpAddress)"
            }
            'OnBehalf' {
                Set-Mailbox -Identity $Mailbox.Identity `
                    -GrantSendOnBehalfTo @{Remove = $User.Identity} -Confirm:$false -ErrorAction Stop
                Write-Ok "Send on Behalf removido de $($User.PrimarySmtpAddress)"
            }
        }
    } catch {
        Write-Err "Remocao falhou: $($_.Exception.Message)"
    }
}

# Pergunta o extra depois de adicionar Full Access
function Invoke-SendRightsPrompt {
    param($Mailbox, $User)

    Write-Host ''
    Write-Host "Voce quer que $($User.PrimarySmtpAddress) tambem possa ENVIAR?" -ForegroundColor Yellow
    Write-Host '   1) Send As        - o email sai COMO a caixa compartilhada'
    Write-Host '   2) Send on Behalf - sai como "Fulano em nome de Caixa"'
    Write-Host '   3) Os dois       - NAO recomendado: quando o user tem os dois,'
    Write-Host '                      o Send As ganha e o on-behalf nunca e usado'
    Write-Host '   4) Nenhum (so Full Access)'
    $opt = Read-Host 'Opcao [4]'

    switch ($opt) {
        '1' { Add-SendAs        -Mailbox $Mailbox -User $User }
        '2' { Add-SendOnBehalf  -Mailbox $Mailbox -User $User }
        '3' { Add-SendAs        -Mailbox $Mailbox -User $User
              Add-SendOnBehalf  -Mailbox $Mailbox -User $User }
        default { Write-Info 'Nenhuma permissao de envio adicionada.' }
    }
}

#endregion

#region ---------- Main ----------

Write-Head 'GERENCIADOR DE CAIXA COMPARTILHADA - EXCHANGE ONLINE'
Write-Warn 'Requer PIM "Exchange Recipient Administrator" ATIVO. Sem ele, o EXO diz que a caixa nao existe.'

Connect-EXOIfNeeded -UPN $AdminUPN

# ---- Qual caixa? (sem assumir dominio nenhum) ----
$mbxObj = $null
while (-not $mbxObj) {
    if ([string]::IsNullOrWhiteSpace($SharedMailbox)) {
        $SharedMailbox = Read-Host "`nDigite o e-mail COMPLETO da caixa compartilhada (Enter vazio = sair)"
    }
    if ([string]::IsNullOrWhiteSpace($SharedMailbox)) {
        Write-Info 'Saindo.'
        return
    }
    $mbxObj = Resolve-Recipient -Identity $SharedMailbox.Trim() -MailboxOnly
    $SharedMailbox = $null
}

$access = Get-SharedMailboxAccess -Mailbox $mbxObj
Show-SharedMailboxAccess -Access $access

# ---- Menu ----
do {
    Write-Host ''
    Write-Host '--------------------- O QUE VOCE QUER FAZER? ---------------------' -ForegroundColor Cyan
    Write-Host '   1) Adicionar usuario com FULL ACCESS (e depois pergunto sobre envio)'
    Write-Host '   2) Adicionar SOMENTE Send As'
    Write-Host '   3) Adicionar SOMENTE Send on Behalf'
    Write-Host '   4) Remover permissao de um usuario'
    Write-Host '   5) Reexibir a lista de acesso'
    Write-Host '   6) Trocar de caixa compartilhada'
    Write-Host '   0) Sair'
    $choice = Read-Host 'Opcao'

    switch ($choice) {

        '1' {
            do {
                $u = Read-Host "`nE-mail do usuario a adicionar (Enter = parar)"
                if ([string]::IsNullOrWhiteSpace($u)) { break }

                $usr = Resolve-Recipient -Identity $u.Trim()
                if (-not $usr) { continue }

                $am = Read-Host "AutoMapping (caixa aparece sozinha no Outlook)? [S/n]"
                $autoMap = -not ($am -match '^[nN]')

                Add-FullAccess -Mailbox $mbxObj -User $usr -AutoMapping $autoMap
                Invoke-SendRightsPrompt -Mailbox $mbxObj -User $usr

                $more = Read-Host "`nAdicionar OUTRO usuario? [s/N]"
            } while ($more -match '^[sSyY]')

            $access = Get-SharedMailboxAccess -Mailbox $mbxObj
            Show-SharedMailboxAccess -Access $access
        }

        '2' {
            do {
                $u = Read-Host "`nE-mail do usuario para Send As (Enter = parar)"
                if ([string]::IsNullOrWhiteSpace($u)) { break }
                $usr = Resolve-Recipient -Identity $u.Trim()
                if ($usr) { Add-SendAs -Mailbox $mbxObj -User $usr }
                $more = Read-Host "Adicionar outro? [s/N]"
            } while ($more -match '^[sSyY]')

            $access = Get-SharedMailboxAccess -Mailbox $mbxObj
            Show-SharedMailboxAccess -Access $access
        }

        '3' {
            do {
                $u = Read-Host "`nE-mail do usuario para Send on Behalf (Enter = parar)"
                if ([string]::IsNullOrWhiteSpace($u)) { break }
                $usr = Resolve-Recipient -Identity $u.Trim()
                if ($usr) { Add-SendOnBehalf -Mailbox $mbxObj -User $usr }
                $more = Read-Host "Adicionar outro? [s/N]"
            } while ($more -match '^[sSyY]')

            $access = Get-SharedMailboxAccess -Mailbox $mbxObj
            Show-SharedMailboxAccess -Access $access
        }

        '4' {
            $u = Read-Host "`nE-mail do usuario a REMOVER (Enter = cancelar)"
            if (-not [string]::IsNullOrWhiteSpace($u)) {
                $usr = Resolve-Recipient -Identity $u.Trim()
                if ($usr) {
                    Write-Host '   1) Full Access'
                    Write-Host '   2) Send As'
                    Write-Host '   3) Send on Behalf'
                    Write-Host '   4) TODAS'
                    $k = Read-Host 'Remover qual?'
                    switch ($k) {
                        '1' { Remove-AnyAccess -Mailbox $mbxObj -User $usr -Kind Full }
                        '2' { Remove-AnyAccess -Mailbox $mbxObj -User $usr -Kind SendAs }
                        '3' { Remove-AnyAccess -Mailbox $mbxObj -User $usr -Kind OnBehalf }
                        '4' { Remove-AnyAccess -Mailbox $mbxObj -User $usr -Kind Full
                              Remove-AnyAccess -Mailbox $mbxObj -User $usr -Kind SendAs
                              Remove-AnyAccess -Mailbox $mbxObj -User $usr -Kind OnBehalf }
                        default { Write-Info 'Cancelado.' }
                    }
                    $access = Get-SharedMailboxAccess -Mailbox $mbxObj
                    Show-SharedMailboxAccess -Access $access
                }
            }
        }

        '5' {
            $access = Get-SharedMailboxAccess -Mailbox $mbxObj
            Show-SharedMailboxAccess -Access $access
        }

        '6' {
            $novo = Read-Host "`nE-mail da nova caixa compartilhada"
            if (-not [string]::IsNullOrWhiteSpace($novo)) {
                $tmp = Resolve-Recipient -Identity $novo.Trim() -MailboxOnly
                if ($tmp) {
                    $mbxObj = $tmp
                    $access = Get-SharedMailboxAccess -Mailbox $mbxObj
                    Show-SharedMailboxAccess -Access $access
                }
            }
        }

        '0' { Write-Info 'Saindo. A sessao EXO continua aberta (Disconnect-ExchangeOnline se quiser fechar).' }

        default { Write-Warn 'Opcao invalida.' }
    }

} while ($choice -ne '0')

#endregion
