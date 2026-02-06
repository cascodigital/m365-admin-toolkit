<#
.SYNOPSIS
    Localizador de arquivos no OneDrive via Microsoft Graph API.

.DESCRIPTION
    Busca arquivos em OneDrive for Business com dois modos de operacao:

    1. Usuario unico  - Busca no OneDrive de um usuario especifico por email
    2. Dominio inteiro - Varre todos os OneDrives provisionados de um dominio

    Suporta wildcards no filtro de busca (ex: *.pdf, relatorio*, *.xlsx).
    Resultados incluem nome, caminho, tamanho e proprietario.

.EXAMPLE
    .\Search-Files.ps1
    # Modo interativo: escolha usuario unico ou dominio, informe filtro

.NOTES
    Autor         : Andre Kittler / Casco Digital
    Versao        : 4.0
    Requisitos    : PowerShell 5.1+, Microsoft.Graph
    Permissoes    : Files.Read.All, User.Read.All, Directory.Read.All
    Limitacao     : API retorna maximo 1000 itens por OneDrive
#>

# Limpar console
Clear-Host

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "           LOCALIZADOR DE ARQUIVOS NO ONEDRIVE" -ForegroundColor Cyan
Write-Host "               VERSAO MULTIPLOS USUARIOS" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Funcao para capturar entrada do usuario
function Get-UserInput {
    param(
        [string]$Prompt,
        [switch]$Required
    )
    
    do {
        Write-Host $Prompt -NoNewline -ForegroundColor Yellow
        Write-Host ": " -NoNewline
        $userInput = [Console]::ReadLine()
        
        if ($Required -and [string]::IsNullOrWhiteSpace($userInput)) {
            Write-Host "Entrada nao pode estar vazia. Tente novamente." -ForegroundColor Red
            Write-Host ""
        }
    } while ($Required -and [string]::IsNullOrWhiteSpace($userInput))
    
    return $userInput.Trim()
}

# Funcao para verificar se OneDrive esta provisionado
function Test-OneDriveProvisioned {
    param($UserId)
    
    try {
        $userDrive = Get-MgUserDefaultDrive -UserId $UserId -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Funcao para buscar arquivos em um usuario
function Search-UserFiles {
    param(
        $Usuario,
        $FiltroArquivo
    )
    
    $arquivosEncontrados = @()
    
    if (-not (Test-OneDriveProvisioned -UserId $Usuario.Id)) {
        return @{
            Usuario = $Usuario
            Status = "OneDrive nao provisionado"
            Arquivos = @()
        }
    }
    
    try {
        $drives = Get-MgUserDrive -UserId $Usuario.Id -All -ErrorAction Stop
        
        foreach ($drive in $drives) {
            try {
                $todosItens = Get-MgUserDriveItem -UserId $Usuario.Id -DriveId $drive.Id -All -Filter "file ne null" -ErrorAction Stop
                $arquivos = $todosItens | Where-Object { $_.Name -like $FiltroArquivo }
                
                foreach ($arquivo in $arquivos) {
                    $caminhoCompleto = ""
                    if ($arquivo.ParentReference -and $arquivo.ParentReference.Path) {
                        $caminhoRelativo = $arquivo.ParentReference.Path -replace "^.+:", ""
                        $caminhoCompleto = "$caminhoRelativo/$($arquivo.Name)" -replace "^/", ""
                    } else {
                        $caminhoCompleto = $arquivo.Name
                    }
                    
                    $arquivosEncontrados += [PSCustomObject]@{
                        Usuario = $Usuario.DisplayName
                        Email = $Usuario.Mail
                        NomeArquivo = $arquivo.Name
                        CaminhoCompleto = $caminhoCompleto
                        Drive = $drive.Name
                        Tamanho = if ($arquivo.Size) { $arquivo.Size } else { 0 }
                        TamanhoKB = if ($arquivo.Size) { [math]::Round($arquivo.Size / 1KB, 2) } else { 0 }
                        UltimaModificacao = $arquivo.LastModifiedDateTime
                        ID = $arquivo.Id
                    }
                }
            }
            catch {
                # Ignorar drives inacessiveis
            }
        }
        
        return @{
            Usuario = $Usuario
            Status = "Sucesso"
            Arquivos = $arquivosEncontrados
        }
    }
    catch {
        return @{
            Usuario = $Usuario
            Status = "Erro: $($_.Exception.Message)"
            Arquivos = @()
        }
    }
}

# Escolher modo de operacao
Write-Host "ESCOLHA O MODO DE BUSCA:" -ForegroundColor Cyan
Write-Host "1. Buscar em OneDrive de usuario especifico" -ForegroundColor White
Write-Host "2. Buscar em todos os OneDrives de um dominio" -ForegroundColor White
Write-Host ""

$opcao = Get-UserInput -Prompt "Digite 1 ou 2" -Required

while ($opcao -notin @("1", "2")) {
    Write-Host "Opcao invalida. Digite 1 ou 2." -ForegroundColor Red
    $opcao = Get-UserInput -Prompt "Digite 1 ou 2" -Required
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan

# Verificar conexao ao Graph
try {
    $conexao = Get-MgContext
    if (-not $conexao) {
        Write-Host "Conectando ao Microsoft Graph..." -ForegroundColor Yellow
        
        $scopes = @(
            "https://graph.microsoft.com/Files.Read.All",
            "https://graph.microsoft.com/User.Read.All",
            "https://graph.microsoft.com/Directory.Read.All"
        )
        
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop
        Write-Host "Conectado com sucesso ao Microsoft Graph" -ForegroundColor Green
    } else {
        Write-Host "Ja conectado ao Microsoft Graph" -ForegroundColor Green
    }
}
catch {
    Write-Host "ERRO ao conectar ao Microsoft Graph:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}

Write-Host ""

if ($opcao -eq "1") {
    # MODO 1: Usuario especifico
    Write-Host "MODO: BUSCA EM USUARIO ESPECIFICO" -ForegroundColor Cyan
    Write-Host ""
    
    $usuarioAlvoUPN = Get-UserInput -Prompt "Digite o email do usuario para pesquisar" -Required
    
    Write-Host ""
    Write-Host "OPCOES DE FILTRO PARA BUSCA DE ARQUIVO:" -ForegroundColor Yellow
    Write-Host "  Exemplo 1: mikrotik          (encontra qualquer arquivo com 'mikrotik' no nome)"
    Write-Host "  Exemplo 2: Mikrotik_V2.docx  (busca pelo nome exato)"
    Write-Host "  Exemplo 3: Mikrotik_V2.*     (qualquer Mikrotik_V2 com qualquer extensao)"
    Write-Host "  Exemplo 4: *v2*              (qualquer arquivo com 'v2' em qualquer parte do nome)"
    Write-Host "  Exemplo 5: *.docx            (todos os arquivos .docx)"
    Write-Host ""
    
    $nomeArquivo = Get-UserInput -Prompt "Digite o filtro de busca para o arquivo" -Required
    
    if (-not ($nomeArquivo.Contains("*"))) {
        $nomeArquivoFiltro = "*$nomeArquivo*"
        Write-Host "Filtro aplicado: $nomeArquivoFiltro" -ForegroundColor Gray
    } else {
        $nomeArquivoFiltro = $nomeArquivo
    }
    
    try {
        Write-Host ""
        Write-Host "Buscando usuario $usuarioAlvoUPN..." -ForegroundColor Yellow
        $usuarioAlvo = Get-MgUser -UserId $usuarioAlvoUPN -ErrorAction Stop
        
        Write-Host "Usuario encontrado: $($usuarioAlvo.DisplayName)" -ForegroundColor Green
        Write-Host ""
        Write-Host "Iniciando busca..." -ForegroundColor Yellow
        
        $resultado = Search-UserFiles -Usuario $usuarioAlvo -FiltroArquivo $nomeArquivoFiltro
        
        if ($resultado.Arquivos.Count -gt 0) {
            Write-Host ""
            Write-Host "ARQUIVOS ENCONTRADOS: $($resultado.Arquivos.Count)" -ForegroundColor Green
            Write-Host "============================================================" -ForegroundColor Green
            
            foreach ($arquivo in $resultado.Arquivos) {
                Write-Host ""
                Write-Host "ARQUIVO: $($arquivo.NomeArquivo)" -ForegroundColor Yellow
                Write-Host "CAMINHO: $($arquivo.CaminhoCompleto)" -ForegroundColor White
                Write-Host "DRIVE: $($arquivo.Drive)" -ForegroundColor Cyan
                Write-Host "TAMANHO: $($arquivo.TamanhoKB) KB"
                Write-Host "MODIFICADO: $($arquivo.UltimaModificacao)"
            }
        } else {
            Write-Host ""
            Write-Host "NENHUM ARQUIVO ENCONTRADO" -ForegroundColor Yellow
            Write-Host "Status: $($resultado.Status)" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "ERRO: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    # MODO 2: Todos usuarios do dominio
    Write-Host "MODO: BUSCA EM TODOS OS USUARIOS DO DOMINIO" -ForegroundColor Cyan
    Write-Host ""
    
    $dominio = Get-UserInput -Prompt "Digite o dominio (ex: cascodigital.com.br)" -Required
    
    Write-Host ""
    Write-Host "OPCOES DE FILTRO PARA BUSCA DE ARQUIVO:" -ForegroundColor Yellow
    Write-Host "  Exemplo 1: mikrotik          (encontra qualquer arquivo com 'mikrotik' no nome)"
    Write-Host "  Exemplo 2: Mikrotik_V2.docx  (busca pelo nome exato)"
    Write-Host "  Exemplo 3: *.pdf             (todos arquivos PDF)"
    Write-Host ""
    
    $nomeArquivo = Get-UserInput -Prompt "Digite o filtro de busca para o arquivo" -Required
    
    if (-not ($nomeArquivo.Contains("*"))) {
        $nomeArquivoFiltro = "*$nomeArquivo*"
        Write-Host "Filtro aplicado: $nomeArquivoFiltro" -ForegroundColor Gray
    } else {
        $nomeArquivoFiltro = $nomeArquivo
    }
    
    try {
        Write-Host ""
        Write-Host "Buscando usuarios do dominio $dominio..." -ForegroundColor Yellow
        
        # Buscar todos usuarios do dominio
        $todosUsuarios = Get-MgUser -All | Where-Object { 
            $_.Mail -like "*@$dominio" -or $_.UserPrincipalName -like "*@$dominio" 
        }
        
        if (-not $todosUsuarios -or $todosUsuarios.Count -eq 0) {
            Write-Host "Nenhum usuario encontrado no dominio $dominio" -ForegroundColor Red
            return
        }
        
        Write-Host "Usuarios encontrados: $($todosUsuarios.Count)" -ForegroundColor Green
        Write-Host ""
        Write-Host "============================================================"
        Write-Host "INICIANDO BUSCA EM TODOS OS USUARIOS..."
        Write-Host "============================================================"
        Write-Host ""
        
        $todosResultados = @()
        $contador = 0
        
        foreach ($usuario in $todosUsuarios) {
            $contador++
            Write-Host "[$contador/$($todosUsuarios.Count)] Processando: $($usuario.DisplayName) ($($usuario.Mail))" -ForegroundColor Yellow
            
            $resultado = Search-UserFiles -Usuario $usuario -FiltroArquivo $nomeArquivoFiltro
            
            Write-Host "  Status: $($resultado.Status)" -ForegroundColor Gray
            if ($resultado.Arquivos.Count -gt 0) {
                Write-Host "  Encontrados: $($resultado.Arquivos.Count) arquivo(s)" -ForegroundColor Green
                $todosResultados += $resultado.Arquivos
            }
            
            Write-Host ""
        }
        
        # Gerar relatorio
        Write-Host "============================================================"
        Write-Host "BUSCA CONCLUIDA!" -ForegroundColor Green
        Write-Host "============================================================"
        Write-Host "Total de arquivos encontrados: $($todosResultados.Count)" -ForegroundColor Green
        
        if ($todosResultados.Count -gt 0) {
            # Gerar arquivo CSV
            $nomeRelatorio = "OneDrive_Busca_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').csv"
            $todosResultados | Export-Csv -Path $nomeRelatorio -NoTypeInformation -Encoding UTF8
            
            Write-Host "Relatorio salvo em: $($PWD)\$nomeRelatorio" -ForegroundColor Cyan
            Write-Host ""
            
            # Mostrar resumo por usuario
            $resumoPorUsuario = $todosResultados | Group-Object Usuario | Sort-Object Count -Descending
            
            Write-Host "RESUMO POR USUARIO:" -ForegroundColor Cyan
            foreach ($grupo in $resumoPorUsuario) {
                Write-Host "  $($grupo.Name): $($grupo.Count) arquivo(s)" -ForegroundColor White
            }
            
            Write-Host ""
            Write-Host "PRIMEIROS 10 ARQUIVOS ENCONTRADOS:" -ForegroundColor Cyan
            $todosResultados | Select-Object -First 10 | ForEach-Object {
                Write-Host "  $($_.Usuario): $($_.NomeArquivo)" -ForegroundColor White
            }
            
            if ($todosResultados.Count -gt 10) {
                Write-Host "  ... e mais $($todosResultados.Count - 10) arquivo(s). Veja o relatorio completo no CSV." -ForegroundColor Gray
            }
        } else {
            Write-Host "Nenhum arquivo encontrado em todos os usuarios pesquisados." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "ERRO: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "============================================================"
Write-Host "Pressione ENTER para fechar..." -ForegroundColor Gray
$null = [Console]::ReadLine()
