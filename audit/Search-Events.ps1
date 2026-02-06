<#
.SYNOPSIS
    Analisador abrangente de logs de eventos Windows com busca multi-ID e exportacao Excel avancada (Local e Remoto)

.DESCRIPTION
    Script robusto para analise forense e investigacao de eventos Windows em ambiente corporativo.
    Realiza busca simultanea de multiplos Event IDs em todos os arquivos de log (.evtx) do sistema,
    incluindo logs arquivados e rotacionados (apenas local). Suporta consulta remota via WinRM.
    Exporta resultados estruturados para Excel com abas separadas por tipo de evento.

    Funcionalidades principais:
    - Busca simultanea de multiplos Event IDs com operador OR logico
    - Modo LOCAL: Varredura completa em todos os logs do sistema incluindo arquivados
    - Modo REMOTO: Consulta em logs ativos de computador remoto via WinRM
    - Filtragem precisa por intervalo de datas com timezone UTC
    - Exportacao Excel multi-aba com agrupamento por Event ID
    - Extracao completa de propriedades XML dos eventos
    - Otimizacao de performance com skip de arquivos antigos
    - Tratamento robusto de arquivos corrompidos ou em uso

    Casos de uso tipicos:
    - Investigacao de incidentes de seguranca
    - Auditoria de logons e acessos 
    - Analise de falhas de sistema
    - Monitoramento de atividades suspeitas
    - Compliance e relatoria regulatoria

.PARAMETER None
    Script interativo - solicita Event IDs, local/remoto e intervalo de datas durante execucao

.EXAMPLE
    .\Analyze-WindowsEventLogs.ps1
    # Tipo de busca: Local
    # Event IDs: 4624,4625,4648
    # Data inicial: 2024-12-01  
    # Data final: 2024-12-31
    # Resultado: Analise completa de logons locais

.EXAMPLE
    .\Analyze-WindowsEventLogs.ps1
    # Tipo de busca: Remoto
    # Nome do computador: SRV-DC01
    # Event IDs: 1074,1076,6005,6006
    # Data inicial: 2024-11-15
    # Data final: 2024-11-30
    # Resultado: Historico remoto de shutdowns e reboots

.INPUTS
    String - Tipo de busca (L para Local, R para Remoto)
    String - Nome do computador remoto (apenas se Remoto)
    String - Lista de Event IDs separados por virgula (ex: 4624,4625,4648)
    DateTime - Data inicial da busca no formato AAAA-MM-DD
    DateTime - Data final da busca no formato AAAA-MM-DD

.OUTPUTS
    - Arquivo Excel (.xlsx) em C:\Temp\EventLog_Exports\
    - Abas separadas por Event ID para analise organizada
    - Aba consolidada "Todos os Eventos" com visao geral
    - Console: Progresso detalhado e estatisticas de coleta

.NOTES
    Autor         : Andre Kittler
    Versao        : 8.0
    Compatibilidade: PowerShell 5.1+, Windows Server/Client

    Requisitos obrigatorios:
    - Privilegios de Administrador local (busca local)
    - Modulo ImportExcel instalado (Install-Module ImportExcel)
    - Acesso de leitura aos logs de eventos do Windows
    - Para busca remota: WinRM habilitado no computador de destino
    - Para busca remota: Credenciais com permissoes no computador remoto

    Configuracoes tecnicas:
    - Busca LOCAL: Recursiva em C:\Windows\System32\winevt\Logs\
    - Busca REMOTA: Consulta logs ativos via Get-WinEvent -ComputerName
    - Extracao completa de propriedades XML EventData
    - Skip automatico de arquivos fora do intervalo temporal

    Logs pesquisados (Local):
    - Application.evtx (eventos de aplicacoes)
    - System.evtx (eventos do sistema)  
    - Security.evtx (eventos de seguranca)
    - Todos os logs adicionais (.evtx) encontrados
    - Logs arquivados e rotacionados historicos

    Logs pesquisados (Remoto):
    - Application, System, Security
    - Microsoft-Windows-PowerShell/Operational
    - Microsoft-Windows-TaskScheduler/Operational
    - E outros logs ativos configurados

    Limitacoes conhecidas:
    - Busca LOCAL: Requer privilegios administrativos para logs de Security
    - Busca REMOTA: Nao acessa logs arquivados, apenas logs ativos
    - Busca REMOTA: Requer WinRM configurado e firewall liberado
    - Performance varia com quantidade de logs e periodo pesquisado

.LINK
    https://docs.microsoft.com/en-us/windows/win32/wes/consuming-events

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.diagnostics/get-winevent
#>

# --- Verificacao de Privilegios de Administrador ---
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "AVISO: Este script nao esta sendo executado com privilegios de Administrador."
    Write-Warning "Para busca LOCAL de eventos de Security, sao necessarios privilegios administrativos."
    Write-Warning "Busca REMOTA pode funcionar se voce tiver credenciais apropriadas."
    Write-Host ""
}

# --- Inicio da Configuracao e Entrada do Usuario ---

# Tenta importar o modulo ImportExcel
try {
    Import-Module ImportExcel -ErrorAction Stop
}
catch {
    Write-Error "ERRO: O modulo 'ImportExcel' nao foi encontrado ou nao pode ser carregado."
    Write-Error "Execute 'Install-Module -Name ImportExcel -Scope CurrentUser' e tente novamente."
    if ($Host.Name -eq "ConsoleHost") {
        Read-Host "Pressione Enter para sair."
    }
    return
}

# Define o diretorio para salvar os arquivos exportados
$exportDirectory = "C:\Temp\EventLog_Exports"
if (-NOT (Test-Path -Path $exportDirectory)) {
    New-Item -ItemType Directory -Path $exportDirectory | Out-Null
    Write-Host "Diretorio de exportacao criado em '$exportDirectory'" -ForegroundColor Cyan
}

# --- NOVA LOGICA: ESCOLHER LOCAL OU REMOTO ---
$searchType = $null
$computerName = $null
$isRemote = $false

while ($null -eq $searchType) {
    Write-Host "`nEscolha o tipo de busca:" -ForegroundColor Cyan
    Write-Host "  [L] Local - Busca completa em todos os logs deste computador (incluindo arquivados)"
    Write-Host "  [R] Remoto - Busca em logs ativos de outro computador via WinRM"
    $input = Read-Host "Digite L para Local ou R para Remoto"

    if ($input -match '^[Ll]$') {
        $searchType = "Local"
        $computerName = $env:COMPUTERNAME
        Write-Host "Modo selecionado: BUSCA LOCAL em '$computerName'" -ForegroundColor Green
    }
    elseif ($input -match '^[Rr]$') {
        $searchType = "Remoto"
        $isRemote = $true

        while ([string]::IsNullOrWhiteSpace($computerName)) {
            $computerName = Read-Host "Digite o nome ou IP do computador remoto"
            if ([string]::IsNullOrWhiteSpace($computerName)) {
                Write-Warning "Nome do computador nao pode ser vazio."
            }
        }

        Write-Host "Modo selecionado: BUSCA REMOTA em '$computerName'" -ForegroundColor Green
        Write-Host "AVISO: Certifique-se de que WinRM esta habilitado no computador de destino." -ForegroundColor Yellow
        Write-Host "AVISO: Busca remota acessa apenas logs ativos, nao logs arquivados." -ForegroundColor Yellow
    }
    else {
        Write-Warning "Opcao invalida. Digite L ou R."
    }
}

# --- LOGICA PARA MULTIPLOS EVENT IDs ---
$eventIds = $null
while ($null -eq $eventIds) {
    try {
        $input = Read-Host "`nDigite o(s) Event ID(s) que deseja buscar, separados por virgula (ex: 4624,4625,4648)"

        # Remove espacos e divide por virgula
        $eventIdStrings = $input -replace '\s+', '' -split ','

        # Converte cada string para inteiro e valida
        $eventIds = @()
        foreach ($idString in $eventIdStrings) {
            if ([string]::IsNullOrWhiteSpace($idString)) {
                continue
            }
            $eventIds += [int]$idString
        }

        if ($eventIds.Count -eq 0) {
            throw "Nenhum Event ID valido foi fornecido"
        }

        Write-Host "Event IDs selecionados: $($eventIds -join ', ')" -ForegroundColor Green

    } catch { 
        Write-Warning "Entrada invalida. Por favor, digite apenas numeros separados por virgula para os Event IDs."
        $eventIds = $null
    }
}

$inputStartDate = $null
$inputEndDate = $null
while ($null -eq $inputStartDate -or $null -eq $inputEndDate) {
    try {
        if ($null -eq $inputStartDate) { $input = Read-Host "Digite a DATA INICIAL da busca no formato AAAA-MM-DD"; $inputStartDate = [datetime]$input }
        if ($null -eq $inputEndDate) { $input = Read-Host "Digite a DATA FINAL da busca no formato AAAA-MM-DD"; $inputEndDate = [datetime]$input }
        if ($inputEndDate -lt $inputStartDate) { Write-Warning "A data final nao pode ser anterior a data inicial."; $inputStartDate = $null; $inputEndDate = $null }
    } catch { Write-Warning "Formato de data invalido. Use AAAA-MM-DD."; $inputStartDate = $null; $inputEndDate = $null }
}

$startDateFormatted = $inputStartDate.Date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.000Z')
$endDateObject = $inputEndDate.Date.AddDays(1).AddTicks(-1)
$endDateFormatted = $endDateObject.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.000Z')

# --- XPATH QUERY PARA MULTIPLOS EVENT IDs ---
$eventIdConditions = @()
foreach ($id in $eventIds) {
    $eventIdConditions += "EventID=$id"
}
$eventIdQuery = "(" + ($eventIdConditions -join " or ") + ")"

$xPathQuery = "*[System[$eventIdQuery and TimeCreated[@SystemTime>='$startDateFormatted' and @SystemTime<='$endDateFormatted']]]"

$allFoundEvents = @()

Write-Host "`nIniciando busca pelos Event IDs '$($eventIds -join ', ')' para o periodo de $($inputStartDate.ToString('yyyy-MM-dd')) ate $($inputEndDate.ToString('yyyy-MM-dd'))" -ForegroundColor Green

# --- LOGICA DE BUSCA: LOCAL vs REMOTO ---

if ($isRemote) {
    # ========== BUSCA REMOTA ==========
    Write-Host "[MODO REMOTO] Consultando logs ativos em '$computerName'..." -ForegroundColor Cyan
    Write-Host "NOTA: Apenas logs ativos serao pesquisados (logs arquivados nao sao acessiveis remotamente)." -ForegroundColor Yellow

    # Lista de logs comuns para consultar remotamente
    $logsToQuery = @(
        'Application',
        'System',
        'Security',
        'Microsoft-Windows-PowerShell/Operational',
        'Microsoft-Windows-TaskScheduler/Operational',
        'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational',
        'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational',
        'Microsoft-Windows-Windows Defender/Operational',
        'Microsoft-Windows-WindowsUpdateClient/Operational'
    )

    $logCount = 0
    foreach ($logName in $logsToQuery) {
        $logCount++
        Write-Host "  [$logCount/$($logsToQuery.Count)] Consultando log: $logName" -ForegroundColor Gray

        try {
            $events = Get-WinEvent -ComputerName $computerName -LogName $logName -FilterXPath $xPathQuery -ErrorAction Stop
            if ($null -ne $events) {
                $allFoundEvents += $events
                Write-Host "      -> Encontrados $($events.Count) eventos." -ForegroundColor White
            }
        }
        catch {
            # Ignora logs que nao existem ou nao podem ser acessados
            if ($_.Exception.Message -notmatch "No events were found") {
                Write-Host "      -> Log inacessivel ou inexistente." -ForegroundColor DarkGray
            }
        }
    }

    # Tenta obter lista de todos os logs disponiveis e consultar os demais
    Write-Host "`n  [Extra] Verificando logs adicionais no sistema remoto..." -ForegroundColor Gray
    try {
        $allRemoteLogs = Get-WinEvent -ComputerName $computerName -ListLog * -ErrorAction Stop | Where-Object { $_.RecordCount -gt 0 -and $_.LogName -notin $logsToQuery }

        foreach ($log in $allRemoteLogs) {
            Write-Host "      Consultando: $($log.LogName)" -ForegroundColor DarkGray
            try {
                $events = Get-WinEvent -ComputerName $computerName -LogName $log.LogName -FilterXPath $xPathQuery -ErrorAction Stop
                if ($null -ne $events) {
                    $allFoundEvents += $events
                    Write-Host "      -> Encontrados $($events.Count) eventos." -ForegroundColor White
                }
            }
            catch {
                # Silenciosamente ignora logs que nao podem ser consultados
            }
        }
    }
    catch {
        Write-Host "      -> Nao foi possivel listar logs adicionais." -ForegroundColor DarkGray
    }
}
else {
    # ========== BUSCA LOCAL ==========
    Write-Host "[MODO LOCAL] Buscando em todos os arquivos de log (.evtx) em '$env:COMPUTERNAME'..." -ForegroundColor Cyan

    $logsPath = "C:\Windows\System32\winevt\Logs"
    Write-Warning "AVISO: A busca sera realizada em TODOS os arquivos de log em '$logsPath'. Este processo pode ser demorado."

    Write-Host "[1/2] Procurando em todos os arquivos de log (.evtx)..." -ForegroundColor Cyan
    $allLogFiles = Get-ChildItem -Path $logsPath -Filter "*.evtx" -File -Recurse -ErrorAction SilentlyContinue

    if ($null -ne $allLogFiles) {
        Write-Host "    -> Encontrados $($allLogFiles.Count) arquivos de log (.evtx) para verificar."

        $fileCount = 0
        foreach ($file in $allLogFiles) {
            $fileCount++

            # Performance: Pula arquivos que estao completamente fora do intervalo de datas
            if ($file.LastWriteTime -lt $inputStartDate.Date) {
                continue
            }

            Write-Host "    [$fileCount/$($allLogFiles.Count)] Verificando: $($file.Name)" -ForegroundColor Gray
            try {
                $archivedEvents = Get-WinEvent -Path $file.FullName -FilterXPath $xPathQuery -ErrorAction Stop
                if ($null -ne $archivedEvents) {
                    $allFoundEvents += $archivedEvents
                    Write-Host "        -> Encontrados $($archivedEvents.Count) eventos." -ForegroundColor White
                }
            }
            catch {
                # Ignora arquivos que nao podem ser processados (em uso, corrompidos, etc)
            }
        }
    }
    else {
        Write-Host "    -> Nenhum arquivo de log (.evtx) encontrado no diretorio especificado." -ForegroundColor Gray
    }
}

# --- Processamento e Exportacao dos Resultados ---

Write-Host "`n[2/2] Processando e exportando resultados..." -ForegroundColor Cyan

if ($allFoundEvents.Count -gt 0) {
    Write-Host "Total de $($allFoundEvents.Count) eventos brutos encontrados. Extraindo detalhes..." -ForegroundColor Green

    $detailedResults = @()

    foreach ($event in $allFoundEvents) {
        $xml = [xml]$event.ToXml()
        $properties = [ordered]@{
            LogName      = $event.LogName
            TimeCreated  = $event.TimeCreated
            EventID      = $event.Id
            ProviderName = $event.ProviderName
            ComputerName = $event.MachineName
            Level        = $event.LevelDisplayName
            UserID       = $event.UserId
        }

        $xml.Event.EventData.Data | ForEach-Object {
            $properties[$_.Name] = $_.'#text'
        }

        $detailedResults += New-Object -TypeName PSObject -Property $properties
    }

    # --- NOME DE ARQUIVO PARA MULTIPLOS EVENT IDs ---
    $eventIdsList = $eventIds -join '-'
    $computerNameSafe = $computerName -replace '[^a-zA-Z0-9]', '_'
    $fileName = "EventLog_$($searchType)_$($computerNameSafe)_IDs-$($eventIdsList)_$($inputStartDate.ToString('yyyyMMdd'))-a-$($inputEndDate.ToString('yyyyMMdd')).xlsx"
    $outputPath = Join-Path -Path $exportDirectory -ChildPath $fileName

    Write-Host "Exportando $($detailedResults.Count) eventos detalhados para '$outputPath'..." -ForegroundColor Cyan

    # Agrupa os resultados por Event ID para criar abas separadas
    $groupedResults = $detailedResults | Group-Object -Property EventID

    # Se houver multiplos Event IDs, cria uma aba para cada um
    if ($groupedResults.Count -gt 1) {
        foreach ($group in $groupedResults) {
            $worksheetName = "Event ID $($group.Name)"
            $group.Group | Export-Excel -Path $outputPath -AutoSize -TableName "EventData_$($group.Name)" -TableStyle Medium9 -WorksheetName $worksheetName
        }

        # Cria tambem uma aba com todos os resultados consolidados
        $detailedResults | Export-Excel -Path $outputPath -AutoSize -TableName "AllEvents" -TableStyle Medium9 -WorksheetName "Todos os Eventos"
    }
    else {
        # Se houver apenas um Event ID, usa o formato original
        $detailedResults | Export-Excel -Path $outputPath -AutoSize -TableName "EventData" -TableStyle Medium9 -WorksheetName "Event ID $($eventIds[0])" -ClearSheet
    }

    Write-Host "`nExportacao concluida com sucesso!" -ForegroundColor Green
    Write-Host "Arquivo salvo em: $outputPath"

    # Exibe resumo por Event ID
    Write-Host "`nResumo dos eventos encontrados:" -ForegroundColor Cyan
    Write-Host "  Computador: $computerName ($searchType)" -ForegroundColor White
    foreach ($group in $groupedResults) {
        Write-Host "  Event ID $($group.Name): $($group.Count) eventos" -ForegroundColor White
    }
}
else {
    Write-Host "`nBusca finalizada. Nenhum evento com os IDs '$($eventIds -join ', ')' foi encontrado no periodo especificado." -ForegroundColor Yellow
    if ($isRemote) {
        Write-Host "DICA: Verifique se WinRM esta habilitado e se voce tem permissoes no computador remoto." -ForegroundColor Yellow
    }
}

Write-Host "`nScript finalizado." -ForegroundColor Cyan
