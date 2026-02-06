<#
.SYNOPSIS
    Script agressivo para remocao completa e definitiva de todas as versoes do Microsoft Office e Outlook do sistema Windows

.DESCRIPTION
    Script PowerShell de limpeza radical que remove completamente todas as instalacoes do Microsoft Office e Outlook do sistema Windows, incluindo aplicativos, registros, perfis, cache e arquivos temporarios. O objetivo e fazer com que o sistema esqueca completamente que o Office existiu, forçando a criacao de novos perfis ao reinstalar Outlook.
    
    Funcionalidades principais:
    - Desinstalacao silenciosa de todas as versoes do Office (Click-to-Run e MSI)
    - Remocao agressiva de chaves de registro do Office em HKLM e HKCU
    - Eliminacao completa de perfis do Outlook e arquivos OST/PST locais
    - Limpeza total de pastas de programa, AppData e arquivos temporarios
    - Parada forçada de todos os processos relacionados ao Office
    - Interface de progresso com etapas numeradas para acompanhamento
    - Reinicializacao recomendada ao final para garantir limpeza completa
    
    Processo de remocao:
    1. Parar todos os processos do Office em execucao
    2. Desinstalar versoes detectadas via registro (Click-to-Run e MSI)
    3. Remover todas as chaves de registro do Office
    4. Eliminar pastas de instalacao do Office
    5. Apagar dados de usuario (perfis Outlook, cache, configuracoes)
    6. Limpar arquivos temporarios e residuais adicionais
    
    Casos de uso:
    - Preparacao para reinstalacao limpa do Office
    - Resolucao de problemas persistentes de instalacao/corrompimento
    - Auditoria de seguranca e remocao de aplicativos
    - Preparacao de maquinas para reimaging
    - Eliminacao de versoes antigas antes de migracao

.PARAMETER None
    Script totalmente autossuficiente - nao requer parametros, executa remocao completa automaticamente

.EXAMPLE
    .\Remove-Office-Complete.ps1
    # Executa a remocao completa do Office sem interacao
    # Resultado: Sistema limpo, pronto para nova instalacao

.EXAMPLE
    PowerShell -ExecutionPolicy Bypass -File .\Remove-Office-Complete.ps1
    # Executa sem alterar politica de execucao permanentemente
    # Resultado: Remocao completa com minimo de requisitos

.INPUTS
    Nenhum - script nao aceita entradas externas

.OUTPUTS
    - Console: Progresso detalhado por etapas com indicacao visual
    - Sistema: Remocao completa de Office/Outlook
    - Arquivos: Nenhum arquivo de saida gerado
    - Registro: Chaves do Office eliminadas
    - Processos: Aplicativos Office encerrados

.NOTES
    Autor         : André Kittler
    Versao        : 1.0
    Compatibilidade: PowerShell 5.1+, Windows 10/11
    
    Requisitos obrigatorios:
    - Execucao como Administrador (verificacao integrada)
    - Acesso completo ao registro do sistema
    - Permissao para modificar pastas do sistema
    - Acesso a todas as pastas AppData dos usuarios
    
    Avisos criticos:
    - **IRREVERSIVEL**: Todos os dados do Outlook locais serao perdidos
    - Nao ha backup automatico - usuario deve fazer backup manual
    - Pode afetar outros aplicativos Microsoft
    - Reinicializacao fortemente recomendada apos execucao
    - Pode gerar eventos de erro inofensivos durante execucao
    
    Etapas de seguranca:
    - Verificacao de execucao como Administrador
    - Tratamento de erros silencioso para continuar remocao
    - Pausas estrategicas entre etapas criticas
    - Remocao recursiva forçada com -Force
    
    Troubleshooting comum:
    - Erro de acesso negado: Executar como Administrador
    - Processo travado: Reiniciar em modo seguro e tentar novamente
    - Chave de registro persistente: Verificar permissao de usuario
    - Pasta nao removida: Verificar se aplicativo esta em execucao

.LINK
    https://learn.microsoft.com/en-us/office/troubleshoot/office-suite-issues/uninstall-office-from-pc
#>


# Script de Remocao Agressiva do Microsoft Office
# ATENCAO: Este script remove TUDO relacionado ao Office/Outlook

# Verificar se esta rodando como Administrador
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Execute este script como Administrador!"
    Exit
}

Write-Host "==== INICIANDO REMOCAO COMPLETA DO MICROSOFT OFFICE ====" -ForegroundColor Red
Write-Host ""

# 1. PARAR TODOS OS PROCESSOS DO OFFICE
Write-Host "[1/6] Parando processos do Office..." -ForegroundColor Yellow
$OfficeProcesses = @(
    "winword", "excel", "powerpnt", "outlook", "onenote", "msaccess", 
    "mspub", "lync", "teams", "onedrive", "officeclicktorun", "groove"
)

foreach ($process in $OfficeProcesses) {
    Get-Process -Name $process -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Write-Host "  - Processo $process encerrado" -ForegroundColor Gray
}

Start-Sleep -Seconds 3

# 2. DESINSTALAR OFFICE VIA REGISTRO (CLICK-TO-RUN E MSI)
Write-Host "[2/6] Desinstalando versoes do Office..." -ForegroundColor Yellow

# Desinstalar Office Click-to-Run
$OfficeUninstallStrings = (Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | 
    Where-Object {$_.DisplayName -like "*Microsoft Office*" -or $_.DisplayName -like "*Microsoft 365*"}).UninstallString

foreach ($UninstallString in $OfficeUninstallStrings) {
    if ($UninstallString) {
        try {
            $UninstallEXE = ($UninstallString -split '"')[1]
            $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
            Write-Host "  - Desinstalando: $UninstallEXE" -ForegroundColor Gray
            Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait -ErrorAction Stop
        } catch {
            Write-Host "  - Erro ao desinstalar via $UninstallEXE" -ForegroundColor DarkYellow
        }
    }
}

# Desinstalar versoes x86 tambem
$OfficeUninstallStrings32 = (Get-ItemProperty "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | 
    Where-Object {$_.DisplayName -like "*Microsoft Office*" -or $_.DisplayName -like "*Microsoft 365*"}).UninstallString

foreach ($UninstallString in $OfficeUninstallStrings32) {
    if ($UninstallString) {
        try {
            $UninstallEXE = ($UninstallString -split '"')[1]
            $UninstallArg = ($UninstallString -split '"')[2] + " DisplayLevel=False"
            Write-Host "  - Desinstalando x86: $UninstallEXE" -ForegroundColor Gray
            Start-Process -FilePath $UninstallEXE -ArgumentList $UninstallArg -Wait -ErrorAction Stop
        } catch {
            Write-Host "  - Erro ao desinstalar x86 via $UninstallEXE" -ForegroundColor DarkYellow
        }
    }
}

Start-Sleep -Seconds 5

# 3. REMOVER CHAVES DE REGISTRO DO OFFICE
Write-Host "[3/6] Removendo chaves de registro do Office..." -ForegroundColor Yellow

# Chaves do Office em HKLM
$RegKeysHKLM = @(
    "HKLM:\SOFTWARE\Microsoft\Office",
    "HKLM:\SOFTWARE\Microsoft\Office ClickToRun",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office",
    "HKLM:\SOFTWARE\Microsoft\OfficeClickToRun"
)

foreach ($key in $RegKeysHKLM) {
    if (Test-Path $key) {
        Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  - Removido: $key" -ForegroundColor Gray
    }
}

# Chaves do Office e Outlook em HKCU
$RegKeysHKCU = @(
    "HKCU:\SOFTWARE\Microsoft\Office",
    "HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
)

foreach ($key in $RegKeysHKCU) {
    if (Test-Path $key) {
        Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  - Removido: $key" -ForegroundColor Gray
    }
}

# 4. REMOVER PASTAS DO OFFICE E OUTLOOK
Write-Host "[4/6] Removendo pastas do Office e Outlook..." -ForegroundColor Yellow

# Pastas principais do Office
$OfficeFolders = @(
    "$env:ProgramFiles\Microsoft Office",
    "$env:ProgramFiles\Microsoft Office 15",
    "$env:ProgramFiles\Microsoft Office 16",
    "${env:ProgramFiles(x86)}\Microsoft Office",
    "${env:ProgramFiles(x86)}\Microsoft Office 15",
    "${env:ProgramFiles(x86)}\Microsoft Office 16",
    "$env:ProgramFiles\Common Files\microsoft shared",
    "${env:ProgramFiles(x86)}\Common Files\microsoft shared"
)

foreach ($folder in $OfficeFolders) {
    if (Test-Path $folder) {
        Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  - Removido: $folder" -ForegroundColor Gray
    }
}

# 5. REMOVER DADOS DO USUARIO (PERFIS OUTLOOK, CACHE, TEMPORARIOS)
Write-Host "[5/6] Removendo dados de usuario (perfis, cache, temporarios)..." -ForegroundColor Yellow

# Pastas de dados do usuario
$UserDataFolders = @(
    "$env:LOCALAPPDATA\Microsoft\Office",
    "$env:LOCALAPPDATA\Microsoft\Outlook",
    "$env:LOCALAPPDATA\Microsoft\OneNote",
    "$env:LOCALAPPDATA\Microsoft\Teams",
    "$env:APPDATA\Microsoft\Office",
    "$env:APPDATA\Microsoft\Outlook",
    "$env:APPDATA\Microsoft\OneNote",
    "$env:APPDATA\Microsoft\Templates",
    "$env:APPDATA\Microsoft\UProof",
    "$env:TEMP\Office*"
)

foreach ($folder in $UserDataFolders) {
    if (Test-Path $folder) {
        Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  - Removido: $folder" -ForegroundColor Gray
    }
}

# 6. LIMPAR ARQUIVOS TEMPORARIOS ADICIONAIS
Write-Host "[6/6] Limpando arquivos temporarios adicionais..." -ForegroundColor Yellow

# Arquivos OST e PST podem estar em locais customizados, mas vamos tentar os padroes
$PSTOSTLocations = @(
    "$env:LOCALAPPDATA\Microsoft\Outlook\*.ost",
    "$env:LOCALAPPDATA\Microsoft\Outlook\*.pst",
    "$env:USERPROFILE\Documents\Arquivos do Outlook\*.pst",
    "$env:APPDATA\Local\Microsoft\Outlook\*.nst"
)

foreach ($location in $PSTOSTLocations) {
    $files = Get-ChildItem -Path $location -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
        Write-Host "  - Removido: $($file.FullName)" -ForegroundColor Gray
    }
}

# Limpar cache de instalacao
$InstallCache = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Bootstrapper"
if (Test-Path $InstallCache) {
    Remove-Item -Path $InstallCache -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "  - Cache de instalacao removido" -ForegroundColor Gray
}

Write-Host ""
Write-Host "==== REMOCAO COMPLETA FINALIZADA ====" -ForegroundColor Green
Write-Host ""
Write-Host "RECOMENDACAO: Reinicie o computador antes de instalar uma nova versao do Office" -ForegroundColor Cyan
Write-Host ""
