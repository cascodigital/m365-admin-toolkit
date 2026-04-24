# ============================================================
# Sync-SharePointDeletion.ps1
# Deleta do SharePoint arquivos que não existem mais na fonte local.
# Útil após migrações (ex: SharePoint Migration Tool) para manter
# o destino em sincronia com a fonte (pasta local, rede, OneDrive etc).
#
# MODO ATUAL: SIMULAÇÃO — nenhum arquivo será deletado.
#
# Para executar de verdade:
#   1. Altere $ModoReal abaixo de $false para $true
# ============================================================

$ModoReal = $false   # <-- mude para $true para deletar de verdade

# ============================================================
# CONFIGURAÇÃO — ajuste para cada cliente/projeto
# ============================================================

$TenantId   = "contoso.onmicrosoft.com"          # tenant do cliente
$SiteRelUrl = "/sites/NomeDaSite"                 # URL relativa da site collection
$LogFile    = "C:\temp\Sync-SharePointDeletion_$(Get-Date -Format 'yyyyMMdd_HHmm').log"

# Mapeamento: nome da Document Library no SharePoint → pasta fonte local
$Mapeamento = @(
    @{ Library = "Documents"; LocalPath = "\\servidor\share\Documentos" },
    @{ Library = "RH";        LocalPath = "\\servidor\share\RH" }
)

# ============================================================

function Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $msg"
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Sites)) {
    Write-Host "Instalando Microsoft.Graph.Sites..."
    Install-Module Microsoft.Graph.Sites -Scope CurrentUser -Force
}

$scopes = if ($ModoReal) { "Sites.FullControl.All","Files.ReadWrite.All" } else { "Sites.Read.All","Files.Read.All" }

try {
    Connect-MgGraph -TenantId $TenantId -Scopes $scopes -ErrorAction Stop
    Log "Conectado ao Microsoft Graph — tenant: $TenantId | Modo: $(if ($ModoReal) { 'REAL' } else { 'SIMULAÇÃO' })"
} catch {
    Log "ERRO: Falha na conexão com Graph — $_"
    exit 1
}

try {
    $spHost    = ($TenantId -replace "\.onmicrosoft\.com$", "") + ".sharepoint.com"
    $site      = Get-MgSite -SiteId "${spHost}:${SiteRelUrl}" -ErrorAction Stop
    Log "Site encontrado: $($site.DisplayName) [$($site.Id)]"
} catch {
    Log "ERRO: Site não encontrado — $_"
    exit 1
}

foreach ($item in $Mapeamento) {
    $libName   = $item.Library
    $localRoot = $item.LocalPath

    Log "=== Iniciando: $libName ==="

    try {
        $drive = Get-MgSiteDrive -SiteId $site.Id -ErrorAction Stop |
                 Where-Object { $_.Name -eq $libName }
        if (-not $drive) { throw "Biblioteca '$libName' não encontrada." }
    } catch {
        Log "ERRO ao localizar biblioteca '$libName': $_"
        Log "=== Pulando: $libName ==="
        continue
    }

    try {
        $cloudItems = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId "root" -ErrorAction Stop
        $allFiles   = @()

        $queue = [System.Collections.Queue]::new()
        foreach ($i in $cloudItems) { $queue.Enqueue(@{ Item = $i; Path = "" }) }

        while ($queue.Count -gt 0) {
            $entry    = $queue.Dequeue()
            $current  = $entry.Item
            $basePath = $entry.Path

            if ($null -ne $current.Folder -and $current.Folder.ChildCount -gt 0) {
                $children = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId $current.Id -ErrorAction SilentlyContinue
                foreach ($c in $children) {
                    $queue.Enqueue(@{ Item = $c; Path = "$basePath\$($current.Name)" })
                }
            } elseif ($null -ne $current.File) {
                $allFiles += @{ Id = $current.Id; Name = $current.Name; RelPath = "$basePath\$($current.Name)".TrimStart("\") }
            }
        }
    } catch {
        Log "ERRO ao listar arquivos de '$libName': $_"
        Log "=== Pulando: $libName ==="
        continue
    }

    Log "Arquivos na nuvem ($libName): $($allFiles.Count)"

    foreach ($f in $allFiles) {
        $localFile = Join-Path $localRoot $f.RelPath

        if (-not (Test-Path -LiteralPath $localFile)) {
            if ($ModoReal) {
                try {
                    Remove-MgDriveItem -DriveId $drive.Id -DriveItemId $f.Id -ErrorAction Stop
                    Log "DELETADO: $($f.RelPath)"
                } catch {
                    Log "ERRO ao deletar '$($f.RelPath)': $_"
                }
            } else {
                Log "SIMULAÇÃO - SERIA DELETADO: $($f.RelPath)"
            }
        }
    }

    Log "=== Concluído: $libName ==="
}

Disconnect-MgGraph | Out-Null
Log "Script finalizado. Modo: $(if ($ModoReal) { 'REAL — arquivos deletados.' } else { 'SIMULAÇÃO — nenhum arquivo foi deletado.' })"
