<#
.SYNOPSIS
    Monitoramento contínuo de latência via ICMP (ping) para múltiplos destinos.

.DESCRIPTION
    Realiza pings em um conjunto de endereços IP (até 5), exibe resultados coloridos
    no console e grava um relatório CSV em tempo real. Usa FileStream/StreamWriter
    para permitir abertura do arquivo enquanto o monitor roda. Classifica latências
    por níveis (VERDE, AMARELO, VERMELHO) e reporta TIMEOUTs.

.PARAMETER Destinos
    Lista interativa de IPs a serem monitorados (entrada via Read-Host).

.PARAMETER Limites
    Limites interativos para determinar cores/status:
      - VERDE  : latência aceitável
      - AMARELO: latência elevada (atenção)
      - VERMELHO: latência alta (problema)
    (Solicitados durante execução.)

.EXAMPLE
    .\monitor-ping.ps1
    Executa em modo interativo, solicitando IPs e limites e iniciando monitoramento contínuo.

.INPUTS
    Nenhum — entradas via prompt interativo.

.OUTPUTS
    - Arquivo CSV com timestamp, latência (ms), status por IP e alerta geral.
    - Saída colorida no console a cada ping.
    - Estatísticas resumidas ao encerrar (Ctrl+C).

.NOTES
    Autor: André Kittler
    Frequência de ping: 1 segundo (configuração atual)
    Arquivo CSV: ping-monitor-<timestamp>.csv
#>

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Monitor de Latencia de Rede" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Solicitar IPs do usuario
$destinos = @()
$indiceIP = 1

Write-Host "Digite os IPs que deseja monitorar (pressione Enter sem digitar nada para finalizar):" -ForegroundColor Yellow
Write-Host ""

while ($true) {
    $ip = Read-Host "IP $indiceIP"

    if ([string]::IsNullOrWhiteSpace($ip)) {
        if ($destinos.Count -eq 0) {
            Write-Host "Voce precisa informar pelo menos 1 IP!" -ForegroundColor Red
            continue
        }
        break
    }

    # Validacao basica de IP
    if ($ip -match "^(\d{1,3}\.){3}\d{1,3}$") {
        $destinos += $ip
        $indiceIP++

        if ($destinos.Count -eq 5) {
            Write-Host ""
            Write-Host "Limite de 5 IPs atingido." -ForegroundColor Yellow
            break
        }
    } else {
        Write-Host "IP invalido. Use o formato: 192.168.0.1" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Configuracao de Limites de Latencia" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Configure os limites para codigo de cores:" -ForegroundColor Yellow
Write-Host "  VERDE   = Latencia normal (boa)" -ForegroundColor Green
Write-Host "  AMARELO = Latencia elevada (atencao)" -ForegroundColor Yellow
Write-Host "  VERMELHO = Latencia alta (problema)" -ForegroundColor Red
Write-Host "  MAGENTA = TIMEOUT (sem resposta)" -ForegroundColor Magenta
Write-Host ""

# Solicitar limite para verde (bom)
while ($true) {
    $limiteVerde = Read-Host "Latencia maxima para VERDE (exemplo: 50 para redes cabeadas, 200 para WiFi)"
    if ($limiteVerde -match "^\d+$" -and [int]$limiteVerde -gt 0) {
        $limiteVerde = [int]$limiteVerde
        break
    }
    Write-Host "Digite apenas numeros maiores que zero!" -ForegroundColor Red
}

# Solicitar limite para amarelo (atencao)
while ($true) {
    $limiteAmarelo = Read-Host "Latencia maxima para AMARELO (deve ser maior que $limiteVerde)"
    if ($limiteAmarelo -match "^\d+$" -and [int]$limiteAmarelo -gt $limiteVerde) {
        $limiteAmarelo = [int]$limiteAmarelo
        break
    }
    Write-Host "Digite um numero maior que $limiteVerde!" -ForegroundColor Red
}

Write-Host ""
Write-Host "Configuracao de limites:" -ForegroundColor Cyan
Write-Host "  VERDE: ate $($limiteVerde)ms" -ForegroundColor Green
Write-Host "  AMARELO: $($limiteVerde + 1)ms ate $($limiteAmarelo)ms" -ForegroundColor Yellow
Write-Host "  VERMELHO: acima de $($limiteAmarelo)ms" -ForegroundColor Red
Write-Host "  MAGENTA: TIMEOUT (sempre destacado)" -ForegroundColor Magenta

Write-Host ""
Write-Host "IPs que serao monitorados:" -ForegroundColor Green
for ($i = 0; $i -lt $destinos.Count; $i++) {
    Write-Host "  IP$($i+1): $($destinos[$i])" -ForegroundColor White
}

$arquivoCSV = "ping-monitor-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"

# Criar cabecalho dinamico do CSV
$cabecalhoCSV = "Timestamp"
for ($i = 0; $i -lt $destinos.Count; $i++) {
    $cabecalhoCSV += ",IP$($i+1)_$($destinos[$i])_ms,IP$($i+1)_Status"
}
$cabecalhoCSV += ",Alerta_Geral"

# Usar StreamWriter para evitar lock de arquivo
$FileMode = [System.IO.FileMode]::Append
$FileAccess = [System.IO.FileAccess]::Write
$FileShare = [IO.FileShare]::ReadWrite

$FileStream = New-Object IO.FileStream($arquivoCSV, [System.IO.FileMode]::Create, $FileAccess, $FileShare)
$StreamWriter = New-Object System.IO.StreamWriter($FileStream)
$StreamWriter.AutoFlush = $true

# Escrever cabecalho
$StreamWriter.WriteLine($cabecalhoCSV)

Write-Host ""
Write-Host "Arquivo de saida: $arquivoCSV" -ForegroundColor Green
Write-Host "Salvando continuamente a cada ping" -ForegroundColor Green
Write-Host "Arquivo pode ser aberto enquanto o script executa" -ForegroundColor Green
Write-Host "Pressione Ctrl+C para parar o monitoramento" -ForegroundColor Gray
Write-Host ""

# Criar cabecalho dinamico da tela
$headerLine = "{0,-23}" -f "Timestamp"
$separatorLine = "{0,-23}" -f "-----------------------"
for ($i = 0; $i -lt $destinos.Count; $i++) {
    $headerLine += " {$($i+1),-15}"
    $separatorLine += " {0,-15}" -f "---------------"
}

# Substituir placeholders pelos IPs
$headerDisplay = $headerLine
for ($i = 0; $i -lt $destinos.Count; $i++) {
    $headerDisplay = $headerDisplay -replace "\{$($i+1),-15\}", ("{0,-15}" -f "IP$($i+1)")
}

Write-Host $headerDisplay -ForegroundColor White
Write-Host $separatorLine -ForegroundColor Gray

$contador = 0
$resultadosTemp = @()

try {
    while ($true) {
        # Timestamp com segundos
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Objeto para armazenar resultados
        $linha = @{
            Timestamp = $timestamp
        }

        # Array para display
        $displayValues = @()

        # CSV line
        $csvLine = "=$([char]34)$timestamp$([char]34)"

        # Flag de alerta geral
        $temTimeout = $false
        $temAltaLatencia = $false

        # Ping em cada destino
        for ($i = 0; $i -lt $destinos.Count; $i++) {
            $ip = $destinos[$i]
            $nomeIP = "IP$($i+1)"

            try {
                $resultado = Test-Connection -ComputerName $ip -Count 1 -ErrorAction Stop

                # Obter latencia de diferentes propriedades possiveis
                if ($null -ne $resultado.ResponseTime) {
                    $latencia = $resultado.ResponseTime
                } elseif ($null -ne $resultado.Latency) {
                    $latencia = $resultado.Latency
                } else {
                    $latencia = ($resultado | Select-Object -ExpandProperty ResponseTime -ErrorAction SilentlyContinue)
                    if ($null -eq $latencia) {
                        $latencia = 0
                    }
                }

                $linha["$($nomeIP)_ms"] = $latencia

                # Determinar status individual baseado nos limites
                if ($latencia -gt $limiteAmarelo) {
                    $statusIP = "ALTA_LATENCIA"
                    $temAltaLatencia = $true
                    $cor = "Red"
                } elseif ($latencia -gt $limiteVerde) {
                    $statusIP = "ATENCAO"
                    $cor = "Yellow"
                } else {
                    $statusIP = "OK"
                    $cor = "Green"
                }

                $linha["$($nomeIP)_Status"] = $statusIP

                # Adicionar ao CSV com status individual
                $csvLine += ",$latencia,$statusIP"

                $texto = "$($latencia)ms"
                $displayValues += @{ Text = $texto; Color = $cor }

            } catch {
                $linha["$($nomeIP)_ms"] = -1
                $linha["$($nomeIP)_Status"] = "TIMEOUT"
                $csvLine += ",-1,TIMEOUT"
                $displayValues += @{ Text = "TIMEOUT"; Color = "Magenta" }
                $temTimeout = $true
            }
        }

        # Determinar alerta geral
        $alertaGeral = ""
        if ($temTimeout) {
            $alertaGeral = "TIMEOUT_DETECTADO"
        } elseif ($temAltaLatencia) {
            $alertaGeral = "ALTA_LATENCIA_DETECTADA"
        }
        $csvLine += ",$alertaGeral"

        # Exibir linha na tela
        Write-Host -NoNewline ($timestamp.PadRight(23))
        foreach ($display in $displayValues) {
            Write-Host -NoNewline (" {0,-15}" -f $display.Text) -ForegroundColor $display.Color
        }
        Write-Host ""

        # Salvar no CSV usando StreamWriter (nao bloqueia arquivo)
        $StreamWriter.WriteLine($csvLine)

        # Armazenar para estatisticas
        $resultadosTemp += $linha
        $contador++

        Start-Sleep -Seconds 1
    }
} catch {
    Write-Host ""
    Write-Host "Monitoramento interrompido." -ForegroundColor Yellow
} finally {
    # Fechar StreamWriter e FileStream
    $StreamWriter.Close()
    $StreamWriter.Dispose()
    $FileStream.Close()
    $FileStream.Dispose()

    Write-Host ""
    Write-Host "Arquivo final: $arquivoCSV" -ForegroundColor Green
    Write-Host ""
    Write-Host "Estatisticas:" -ForegroundColor Cyan
    Write-Host "  Total de pings: $contador" -ForegroundColor White
    Write-Host "  Limites usados: Verde ate $($limiteVerde)ms | Amarelo ate $($limiteAmarelo)ms | Vermelho acima" -ForegroundColor White

    # Calcular estatisticas para cada IP
    for ($i = 0; $i -lt $destinos.Count; $i++) {
        $nomeIP = "IP$($i+1)"
        $ip = $destinos[$i]

        $valores = $resultadosTemp | Where-Object { $_["$($nomeIP)_Status"] -ne "TIMEOUT" } | ForEach-Object { $_["$($nomeIP)_ms"] }
        $timeouts = ($resultadosTemp | Where-Object { $_["$($nomeIP)_Status"] -eq "TIMEOUT" }).Count
        $oks = ($resultadosTemp | Where-Object { $_["$($nomeIP)_Status"] -eq "OK" }).Count
        $atencoes = ($resultadosTemp | Where-Object { $_["$($nomeIP)_Status"] -eq "ATENCAO" }).Count
        $altas = ($resultadosTemp | Where-Object { $_["$($nomeIP)_Status"] -eq "ALTA_LATENCIA" }).Count

        if ($valores.Count -gt 0) {
            $media = ($valores | Measure-Object -Average).Average
            $min = ($valores | Measure-Object -Minimum).Minimum
            $max = ($valores | Measure-Object -Maximum).Maximum

            Write-Host ""
            Write-Host "  $nomeIP ($ip):" -ForegroundColor Yellow
            Write-Host "    Media: $([math]::Round($media, 2))ms" -ForegroundColor White
            Write-Host "    Min: $($min)ms | Max: $($max)ms" -ForegroundColor White
            Write-Host "    Distribuicao:" -ForegroundColor White
            Write-Host "      OK (ate $($limiteVerde)ms): $oks" -ForegroundColor Green
            Write-Host "      ATENCAO ($($limiteVerde + 1)-$($limiteAmarelo)ms): $atencoes" -ForegroundColor Yellow
            Write-Host "      ALTA_LATENCIA (acima $($limiteAmarelo)ms): $altas" -ForegroundColor Red
            Write-Host "      TIMEOUT: $timeouts" -ForegroundColor Magenta
            if ($contador -gt 0) {
                Write-Host "    Taxa de sucesso: $([math]::Round((($valores.Count) / $contador) * 100, 2))%" -ForegroundColor White
            }
        }
    }

    Write-Host ""
}
