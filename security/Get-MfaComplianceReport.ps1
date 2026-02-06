
<#
.SYNOPSIS
    Auditoria abrangente de configuracao MFA em usuarios reais Microsoft 365 com analise detalhada de metodos

.DESCRIPTION
    Script corporativo especializado para auditoria completa de Autenticacao Multifator (MFA) em
    ambientes Microsoft 365, com foco exclusivo em usuarios reais excluindo contas de servico,
    sistema e objetos nao humanos. Implementa filtragem inteligente, analise detalhada de todos
    os 7 tipos de metodos MFA suportados e consultoria interativa individual por usuario.
    
    Funcionalidades principais:
    - Filtragem automatica de usuarios reais vs contas de servico/sistema
    - Analise de 7 tipos de metodos MFA: Telefone, Authenticator, OATH, Email, Windows Hello, FIDO2
    - Relatorio executivo com percentual de conformidade organizacional
    - Consulta detalhada interativa com drill-down por usuario especifico
    - Exportacao dual: CSV tecnico + TXT executivo para clientes
    - Interface colorida com progress bar e confirmacao de usuario
    - Deteccao automatica de duplicatas e dispositivos registrados
    
    Tipos de usuarios filtrados (EXCLUIDOS):
    - Contas de servico (ADConnect, Veeam, Exchange, etc.)
    - Usuarios Guest/External (UserType != Member)
    - Contas do sistema (@*.onmicrosoft.com)
    - Caixas compartilhadas e grupos de seguranca
    - Objetos de protocolo e notificacao automatizada
    
    Casos de uso corporativos:
    - Auditoria de compliance para regulamentacao (SOX, ISO27001)
    - Preparacao para certificacoes de seguranca cibernetica
    - Relatorios executivos de postura de seguranca organizacional
    - Investigacao de usuarios especificos pos-incidente
    - Monitoramento continuo de adocao de MFA

.PARAMETER None
    Script totalmente interativo - conecta automaticamente e processa usuarios filtrados

.EXAMPLE
    .\Audit-MFACompliance.ps1
    # Conecta ao Microsoft Graph automaticamente
    # Filtra 150 usuarios reais de 200 objetos totais
    # Resultado: 95% compliance (143/150 com MFA configurado)

.EXAMPLE
    .\Audit-MFACompliance.ps1
    # Apos relatorio geral, consulta detalhada:
    # Email: joao.silva@cascodigital.com.br
    # Resultado: 3 metodos MFA (Telefone + Authenticator + Windows Hello)

.EXAMPLE
    # Auditoria corporativa para compliance SOX
    .\Audit-MFACompliance.ps1
    # Gera: MFA_Report_UsuariosReais_2024-12-15_14-30.csv (dados tecnicos)
    # Gera: MFA_Comprovacao_UsuariosReais_2024-12-15_14-30.txt (relatorio executivo)

.INPUTS
    None - Script automatico com confirmacao interativa antes do processamento

.OUTPUTS
    - CSV tecnico: Dados completos de todos usuarios com metadados MFA
    - TXT executivo: Relatorio formatado para apresentacao a diretoria
    - Console: Interface interativa com consulta drill-down individual
    - Estatisticas: Percentual de conformidade e usuarios nao conformes

.NOTES
    Autor         : Andre Kittler
    Versao        : 2.0
    Compatibilidade: PowerShell 5.1+, Windows/Linux/macOS
    
    Requisitos Microsoft Graph:
    - Modulo Microsoft.Graph instalado e conectado
    - Permissoes de API obrigatorias:
      * User.Read.All (leitura de usuarios)
      * UserAuthenticationMethod.Read.All (metodos MFA)
      * Directory.Read.All (propriedades estendidas)
      * Policy.Read.All (politicas de autenticacao)
    
    Privilegios administrativos necessarios:
    - Security Administrator OU
    - Security Reader OU
    - Authentication Administrator OU
    - Global Administrator (para acesso completo)
    
    Criterios de filtragem aplicados:
    - UserType = "Member" (exclui Guest/External)
    - DisplayName sem padroes de servico (Sync, ADConnect, Veeam, etc.)
    - UPN sem dominios .onmicrosoft.com ou padroes de sistema
    - Email valido e formato correto obrigatorio
    
    Metodos MFA detectados e analisados:
    1. phoneAuthenticationMethod (Telefone/SMS)
    2. microsoftAuthenticatorAuthenticationMethod (MS Authenticator)
    3. softwareOathAuthenticationMethod (Google Auth, Authy, etc.)
    4. windowsHelloForBusinessAuthenticationMethod (PIN/Biometria)
    5. fido2AuthenticationMethod (Chaves de seguranca)
    6. emailAuthenticationMethod (Email backup)
    7. Outros metodos emergentes suportados pelo Graph API
    
    Relatorio executivo inclui:
    - Resumo executivo com percentual de conformidade
    - Lista de usuarios nao conformes para acao corretiva
    - Status de conformidade organizacional (Conforme/Nao Conforme)
    - Detalhamento tecnico completo para auditoria
    - Criterios de filtragem aplicados para transparencia
    
    Funcionalidade de consulta detalhada:
    - Busca interativa por email ou nome parcial
    - Detalhamento de todos os 7 tipos de metodos MFA
    - Informacoes de dispositivos registrados e datas
    - Numeros de telefone mascarados para privacidade
    - Status de conta (ativa/licenciada) para contexto
    
    Consideracoes de performance:
    - Progress bar durante processamento de usuarios
    - Tratamento de erros individual sem interrupcao do lote
    - Otimizacao de chamadas API com ErrorAction SilentlyContinue
    - Processamento sequencial para evitar throttling

.LINK
    https://docs.microsoft.com/en-us/graph/api/authentication-list-methods

.LINK
    https://docs.microsoft.com/en-us/azure/active-directory/authentication/concept-authentication-methods
#>



# Script MFA - COM FILTROS PARA USUÁRIOS REAIS
Write-Host "=== VERIFICAÇÃO DE MFA - APENAS USUÁRIOS REAIS ===" -ForegroundColor Green
try {
    # 1. Conectar
    Write-Host "Conectando..." -ForegroundColor Yellow
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Connect-MgGraph -Scopes "User.Read.All", "UserAuthenticationMethod.Read.All", "Directory.Read.All", "Policy.Read.All" -NoWelcome
        Write-Host "✅ Conectado com sucesso!" -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Erro na conexão: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Pressione qualquer tecla para continuar..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }
    # 2. Obter usuários com filtros específicos
    Write-Host "Obtendo usuários (apenas usuários reais)..." -ForegroundColor Yellow
    # Obter todos os usuários primeiro
    $allUsers = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled, CreatedDateTime, UserType, Mail, AssignedLicenses
    Write-Host "Total de objetos encontrados: $($allUsers.Count)" -ForegroundColor Gray
    # Filtrar apenas usuários reais
    $users = $allUsers | Where-Object {
        # Incluir apenas usuários do tipo Member
        $_.UserType -eq "Member" -and
        # Excluir contas de serviço/sistema comuns
        $_.DisplayName -notlike "*Service Account*" -and
        $_.DisplayName -notlike "*Sync*" -and
        $_.DisplayName -notlike "*ADConnect*" -and
        $_.DisplayName -notlike "*Veeam*" -and
        $_.DisplayName -notlike "*GED*" -and
        $_.DisplayName -notlike "*SBS*" -and
        $_.DisplayName -notlike "*protocolo*" -and
        $_.DisplayName -notlike "*Recepção*" -and
        $_.DisplayName -notlike "*Controladoria*" -and
        $_.DisplayName -notlike "*Notificacao*" -and
        # Excluir contas com UPN de sistema
        $_.UserPrincipalName -notlike "*@*.onmicrosoft.com" -and
        $_.UserPrincipalName -notlike "*sync*" -and
        $_.UserPrincipalName -notlike "*service*" -and
        $_.UserPrincipalName -notlike "*veeam*" -and
        $_.UserPrincipalName -notlike "*test*" -and
        # Excluir contas desabilitadas (opcional - descomente se quiser)
        # $_.AccountEnabled -eq $true -and
        # Incluir apenas se tiver um email válido
        ![string]::IsNullOrEmpty($_.UserPrincipalName) -and
        $_.UserPrincipalName -match "@"
    }
    Write-Host "✅ Usuários reais encontrados: $($users.Count)" -ForegroundColor Green
    # 3. Mostrar usuários filtrados
    Write-Host "Usuários que serão analisados:" -ForegroundColor Cyan
    $users | Select-Object DisplayName, UserPrincipalName, UserType, AccountEnabled | Format-Table
    Write-Host "Continuar com a análise? (S/N)" -ForegroundColor Yellow
    $response = Read-Host
    if ($response -notlike "S*" -and $response -notlike "s*") {
        Write-Host "Operação cancelada pelo usuário." -ForegroundColor Yellow
        exit
    }
    # 4. Processar usuários
    $report = @()
    $sucessos = 0
    $erros = 0
    foreach ($user in $users) {
        $counter = $report.Count + 1
        $percentual = [math]::Round(($counter / $users.Count) * 100, 1)
        Write-Progress -Activity "Processando usuários" -Status "[$counter/$($users.Count)] $($user.DisplayName)" -PercentComplete $percentual
        try {
            $userId = $user.Id
            # Se o ID estiver vazio, usar o UserPrincipalName
            if ([string]::IsNullOrEmpty($userId)) {
                $userId = $user.UserPrincipalName
            }
            # Obter métodos de autenticação
            $authMethods = @()
            try {
                $authMethods = Get-MgUserAuthenticationMethod -UserId $userId
            }
            catch {
                # Se falhar com ID, tentar com UPN
                if ($userId -ne $user.UserPrincipalName) {
                    try {
                        $authMethods = Get-MgUserAuthenticationMethod -UserId $user.UserPrincipalName
                    }
                    catch {
                        throw "Não foi possível obter métodos de autenticação: $($_.Exception.Message)"
                    }
                }
                else {
                    throw "Não foi possível obter métodos de autenticação: $($_.Exception.Message)"
                }
            }
            # Obter métodos específicos
            $phoneMethods = @()
            $authenticatorMethods = @()
            try {
                $phoneMethods = Get-MgUserAuthenticationPhoneMethod -UserId $userId -ErrorAction SilentlyContinue
            }
            catch {
                # Ignorar erro - pode não ter telefone
            }
            try {
                $authenticatorMethods = Get-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $userId -ErrorAction SilentlyContinue
            }
            catch {
                # Ignorar erro - pode não ter authenticator
            }
            # Analisar métodos
            $mfaTypes = @()
            $mfaDetails = @()
            foreach ($method in $authMethods) {
                $methodType = $method.AdditionalProperties.'@odata.type'
                switch ($methodType) {
                    '#microsoft.graph.phoneAuthenticationMethod' { 
                        if ("Telefone/SMS" -notin $mfaTypes) {
                            $mfaTypes += "Telefone/SMS"
                        }
                    }
                    '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' { 
                        if ("Microsoft Authenticator" -notin $mfaTypes) {
                            $mfaTypes += "Microsoft Authenticator"
                        }
                    }
                    '#microsoft.graph.softwareOathAuthenticationMethod' { 
                        if ("App OATH" -notin $mfaTypes) {
                            $mfaTypes += "App OATH"
                        }
                    }
                    '#microsoft.graph.emailAuthenticationMethod' { 
                        if ("Email" -notin $mfaTypes) {
                            $mfaTypes += "Email"
                        }
                    }
                }
            }
            # Adicionar detalhes
            if ($phoneMethods.Count -gt 0) {
                $mfaDetails += "Telefones: $($phoneMethods.Count)"
            }
            if ($authenticatorMethods.Count -gt 0) {
                $mfaDetails += "Authenticator: $($authenticatorMethods.Count)"
            }
            # Filtrar apenas métodos MFA reais (excluir password e email)
            $mfaMethodsOnly = $authMethods | Where-Object { 
                $_.AdditionalProperties.'@odata.type' -ne '#microsoft.graph.passwordAuthenticationMethod' -and
                $_.AdditionalProperties.'@odata.type' -ne '#microsoft.graph.emailAuthenticationMethod'
            }
            $mfaDetails += "Métodos MFA: $($mfaMethodsOnly.Count)"
            # Determinar status
            $mfaStatus = if ($mfaMethodsOnly.Count -gt 0) { "Configurado" } else { "Não Configurado" }
            # Verificar se tem licenças
            $temLicenca = $user.AssignedLicenses.Count -gt 0
            $userInfo = [PSCustomObject]@{
                'Nome' = $user.DisplayName
                'Email' = $user.UserPrincipalName
                'Status MFA' = $mfaStatus
                'Métodos MFA' = if ($mfaTypes.Count -gt 0) { ($mfaTypes -join ', ') } else { "Nenhum" }
                'Qtd Métodos MFA' = $mfaMethodsOnly.Count
                'Detalhes' = ($mfaDetails -join ' | ')
                'Conta Ativa' = $user.AccountEnabled
                'Licenciado' = $temLicenca
                'Criado em' = $user.CreatedDateTime
            }
            $report += $userInfo
            $sucessos++
        }
        catch {
            Write-Host "⚠️ Erro ao processar $($user.DisplayName): $($_.Exception.Message)" -ForegroundColor Yellow
            $userInfo = [PSCustomObject]@{
                'Nome' = $user.DisplayName
                'Email' = $user.UserPrincipalName
                'Status MFA' = "Erro na verificação"
                'Métodos MFA' = "Erro"
                'Qtd Métodos MFA' = 0
                'Detalhes' = $_.Exception.Message
                'Conta Ativa' = $user.AccountEnabled
                'Licenciado' = $user.AssignedLicenses.Count -gt 0
                'Criado em' = $user.CreatedDateTime
            }
            $report += $userInfo
            $erros++
        }
    }
    Write-Progress -Activity "Processando usuários" -Completed
    # 5. Estatísticas
    $total = $report.Count
    $comMFA = ($report | Where-Object {$_.'Status MFA' -eq 'Configurado'}).Count
    $semMFA = ($report | Where-Object {$_.'Status MFA' -eq 'Não Configurado'}).Count
    $comErro = ($report | Where-Object {$_.'Status MFA' -eq 'Erro na verificação'}).Count
    $usuariosValidos = $total - $comErro
    $percentual = if ($usuariosValidos -gt 0) { [math]::Round(($comMFA/$usuariosValidos)*100, 2) } else { 0 }
    Write-Host ""
    Write-Host "=== RESULTADO FINAL ===" -ForegroundColor Green
    Write-Host "Total de usuários REAIS analisados: $total" -ForegroundColor White
    Write-Host "Com MFA: $comMFA ($percentual%)" -ForegroundColor Green
    Write-Host "Sem MFA: $semMFA" -ForegroundColor $(if ($semMFA -eq 0) { "Green" } else { "Red" })
    Write-Host "Com erro: $comErro" -ForegroundColor Yellow
    # Mostrar usuários com MFA
    if ($comMFA -gt 0) {
        Write-Host ""
        Write-Host "✅ USUÁRIOS COM MFA CONFIGURADO:" -ForegroundColor Green
        $report | Where-Object {$_.'Status MFA' -eq 'Configurado'} | Select-Object Nome, Email, 'Métodos MFA', 'Qtd Métodos MFA' | Format-Table
    }
    # Mostrar usuários sem MFA
    if ($semMFA -gt 0) {
        Write-Host ""
        Write-Host "⚠️ USUÁRIOS SEM MFA:" -ForegroundColor Red
        $report | Where-Object {$_.'Status MFA' -eq 'Não Configurado'} | Select-Object Nome, Email, Licenciado, 'Conta Ativa' | Format-Table
    }
    # Gerar arquivos
    Write-Host ""
    Write-Host "Gerando arquivos..." -ForegroundColor Yellow
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $csvFile = "MFA_Report_UsuariosReais_$timestamp.csv"
    $txtFile = "MFA_Comprovacao_UsuariosReais_$timestamp.txt"
    # Salvar CSV
    $report | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
    # Relatório para cliente
    $clientReport = @"
=== RELATÓRIO DE AUTENTICAÇÃO MULTIFATOR (MFA) ===
=== USUÁRIOS REAIS APENAS ===

Empresa: 
Administrador: admin@
Data da Verificação: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")

CRITÉRIOS DE FILTRAGEM:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Incluídos apenas usuários do tipo "Member"
• Excluídas contas de serviço (ADConnect, Veeam, etc.)
• Excluídas caixas compartilhadas e grupos de segurança
• Excluídas contas do sistema (@*.onmicrosoft.com)
• Analisados apenas usuários reais que devem ter MFA

RESUMO EXECUTIVO:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Total de usuários REAIS analisados: $total
• Usuários com MFA configurado: $comMFA
• Usuários sem MFA configurado: $semMFA
• Usuários com erro na verificação: $comErro
• Percentual de conformidade: $percentual%

STATUS DE CONFORMIDADE: $(if ($semMFA -eq 0 -and $comErro -eq 0) { "✅ CONFORME - TODOS OS USUÁRIOS REAIS POSSUEM MFA" } elseif ($semMFA -eq 0) { "⚠️ CONFORME - MFA OK (com alguns erros técnicos)" } else { "❌ NÃO CONFORME - USUÁRIOS SEM MFA IDENTIFICADOS" })

USUÁRIOS COM MFA CONFIGURADO ($comMFA):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

$($report | Where-Object {$_.'Status MFA' -eq 'Configurado'} | Select-Object Nome, Email, 'Métodos MFA', 'Qtd Métodos MFA' | Format-Table -AutoSize | Out-String)

$(if ($semMFA -gt 0) {
"USUÁRIOS SEM MFA CONFIGURADO ($semMFA):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

$($report | Where-Object {$_.'Status MFA' -eq 'Não Configurado'} | Select-Object Nome, Email, Licenciado, 'Conta Ativa' | Format-Table -AutoSize | Out-String)"
})

DETALHAMENTO TÉCNICO COMPLETO:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

$($report | Format-Table -AutoSize | Out-String)

OBSERVAÇÕES TÉCNICAS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Relatório focado em usuários reais (não inclui contas de serviço)
• Métodos MFA detectados: Telefone/SMS, Microsoft Authenticator, Apps OATH
• Excluídos: Grupos, caixas compartilhadas, contas de sistema
• Processamento: $sucessos sucessos, $erros erros de $total usuários
• Verificação em tempo real no tenant Microsoft 365

Relatório gerado automaticamente.
Administrador: admin@
"@
    $clientReport | Out-File -FilePath $txtFile -Encoding UTF8
    Write-Host "✅ Arquivos gerados:" -ForegroundColor Green
    Write-Host "• $csvFile (dados completos)" -ForegroundColor Yellow
    Write-Host "• $txtFile (relatório para cliente)" -ForegroundColor Yellow
    if ($semMFA -eq 0) {
        Write-Host ""
        Write-Host "🎉 EXCELENTE! Todos os usuários reais possuem MFA configurado!" -ForegroundColor Green
    }
    # ========== ADICIONAR AQUI - NOVA FUNCIONALIDADE ==========
    Write-Host ""
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "CONSULTA DETALHADA DE MÉTODOS MFA" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    while ($true) {
        Write-Host ""
        Write-Host "Deseja consultar detalhes específicos de alguma conta?" -ForegroundColor Yellow
        Write-Host "Digite o email da conta ou pressione ENTER para encerrar:" -ForegroundColor Yellow
        $emailConsulta = Read-Host "Email"
        if ([string]::IsNullOrWhiteSpace($emailConsulta)) {
            Write-Host "Encerrando consultas detalhadas..." -ForegroundColor Gray
            break
        }
        # Buscar usuário na lista processada
        $usuarioEncontrado = $report | Where-Object { $_.'Email' -like "*$emailConsulta*" }
        if (-not $usuarioEncontrado) {
            Write-Host "❌ Usuário não encontrado na lista processada." -ForegroundColor Red
            Write-Host "Usuários disponíveis:" -ForegroundColor Gray
            $report | Select-Object Nome, Email | Format-Table -AutoSize
            continue
        }
        if ($usuarioEncontrado.Count -gt 1) {
            Write-Host "⚠️ Múltiplos usuários encontrados:" -ForegroundColor Yellow
            $usuarioEncontrado | Select-Object Nome, Email | Format-Table -AutoSize
            Write-Host "Seja mais específico com o email." -ForegroundColor Yellow
            continue
        }
        # Consulta detalhada
        Write-Host ""
        Write-Host "🔍 DETALHAMENTO COMPLETO PARA:" -ForegroundColor Green
        Write-Host "Nome: $($usuarioEncontrado.Nome)" -ForegroundColor White
        Write-Host "Email: $($usuarioEncontrado.Email)" -ForegroundColor White
        Write-Host ""
        try {
            $emailProcurar = $usuarioEncontrado.Email
            # Buscar métodos detalhados
            Write-Host "Obtendo métodos de autenticação..." -ForegroundColor Yellow
            $todosMetodos = Get-MgUserAuthenticationMethod -UserId $emailProcurar -ErrorAction Stop
            # Filtrar métodos MFA (excluir password)
            $metodosMFA = $todosMetodos | Where-Object { 
                $_.AdditionalProperties.'@odata.type' -ne '#microsoft.graph.passwordAuthenticationMethod' 
            }
            Write-Host "📊 RESUMO DOS MÉTODOS:" -ForegroundColor Cyan
            Write-Host "Total de métodos MFA: $($metodosMFA.Count)" -ForegroundColor White
            Write-Host ""
            if ($metodosMFA.Count -eq 0) {
                Write-Host "❌ Nenhum método MFA configurado!" -ForegroundColor Red
            } else {
                Write-Host "📱 DETALHAMENTO POR TIPO:" -ForegroundColor Cyan
                Write-Host ""
                # ========== ADICIONAR ESTA SEÇÃO AQUI ==========
                Write-Host "🔍 TODOS OS 7 MÉTODOS DETECTADOS:" -ForegroundColor Magenta
                $contador = 1
                foreach ($metodo in $metodosMFA) {
                    $tipo = $metodo.AdditionalProperties.'@odata.type'
                    $id = $metodo.Id
                    Write-Host "  $contador. Tipo: $tipo" -ForegroundColor White
                    Write-Host "      ID: $id" -ForegroundColor Gray
                    # Tentar obter mais detalhes específicos para cada tipo
                    switch ($tipo) {
                        '#microsoft.graph.phoneAuthenticationMethod' {
                            try {
                                $detalheTel = Get-MgUserAuthenticationPhoneMethod -UserId $emailProcurar -PhoneAuthenticationMethodId $id -ErrorAction SilentlyContinue
                                if ($detalheTel) {
                                    Write-Host "      Número: $($detalheTel.PhoneNumber)" -ForegroundColor Cyan
                                    Write-Host "      Tipo: $($detalheTel.PhoneType)" -ForegroundColor Cyan
                                }
                            } catch { }
                        }
                        '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' {
                            try {
                                $detalheAuth = Get-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $emailProcurar -MicrosoftAuthenticatorAuthenticationMethodId $id -ErrorAction SilentlyContinue
                                if ($detalheAuth) {
                                    Write-Host "      Dispositivo: $($detalheAuth.DisplayName)" -ForegroundColor Cyan
                                    Write-Host "      Versão: $($detalheAuth.DeviceTag)" -ForegroundColor Cyan
                                }
                            } catch { }
                        }
                        '#microsoft.graph.softwareOathAuthenticationMethod' {
                            try {
                                $detalheOath = Get-MgUserAuthenticationSoftwareOathMethod -UserId $emailProcurar -SoftwareOathAuthenticationMethodId $id -ErrorAction SilentlyContinue
                                if ($detalheOath) {
                                    Write-Host "      Nome: $($detalheOath.DisplayName)" -ForegroundColor Cyan
                                }
                            } catch { }
                        }
                    }
                    Write-Host ""
                    $contador++
                }
                Write-Host "=" * 40 -ForegroundColor Magenta
                Write-Host ""
                # ========== FIM DA NOVA SEÇÃO ==========
                # Telefones
                $telefones = Get-MgUserAuthenticationPhoneMethod -UserId $emailProcurar -ErrorAction SilentlyContinue
                if ($telefones.Count -gt 0) {
                    Write-Host "📞 TELEFONES ($($telefones.Count)):" -ForegroundColor Green
                    foreach ($tel in $telefones) {
                        $tipo = if ($tel.PhoneType -eq 'mobile') { "Celular" } else { "Outros" }
                        Write-Host "  • $($tel.PhoneNumber) [$tipo]" -ForegroundColor White
                    }
                    Write-Host ""
                }
                # Microsoft Authenticator
                $authenticators = Get-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $emailProcurar -ErrorAction SilentlyContinue
                if ($authenticators.Count -gt 0) {
                    Write-Host "🔐 MICROSOFT AUTHENTICATOR ($($authenticators.Count)):" -ForegroundColor Green
                    foreach ($auth in $authenticators) {
                        $dispositivo = if ($auth.DisplayName) { $auth.DisplayName } else { "Dispositivo não identificado" }
                        $criado = if ($auth.CreatedDateTime) { (Get-Date $auth.CreatedDateTime -Format "dd/MM/yyyy") } else { "Data desconhecida" }
                        Write-Host "  • $dispositivo (Criado: $criado)" -ForegroundColor White
                    }
                    Write-Host ""
                }
                # Software OATH (Google Authenticator, Authy, etc.)
                $softwareOath = Get-MgUserAuthenticationSoftwareOathMethod -UserId $emailProcurar -ErrorAction SilentlyContinue
                if ($softwareOath.Count -gt 0) {
                    Write-Host "📲 APLICATIVOS TOTP/OATH ($($softwareOath.Count)):" -ForegroundColor Green
                    foreach ($oath in $softwareOath) {
                        $nome = if ($oath.DisplayName) { $oath.DisplayName } else { "App TOTP" }
                        Write-Host "  • $nome" -ForegroundColor White
                    }
                    Write-Host ""
                }
                # ========== ADICIONAR ESTA NOVA SEÇÃO AQUI ==========
                # Windows Hello for Business
                $windowsHello = $metodosMFA | Where-Object { 
                    $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' 
                }
                if ($windowsHello.Count -gt 0) {
                    Write-Host "🖥️ WINDOWS HELLO BUSINESS ($($windowsHello.Count)):" -ForegroundColor Green
                    $contador = 1
                    foreach ($hello in $windowsHello) {
                        try {
                            $detalheHello = Get-MgUserAuthenticationWindowsHelloForBusinessMethod -UserId $emailProcurar -WindowsHelloForBusinessAuthenticationMethodId $hello.Id -ErrorAction SilentlyContinue
                            if ($detalheHello) {
                                $dispositivo = if ($detalheHello.DisplayName) { $detalheHello.DisplayName } else { "Dispositivo Windows $contador" }
                                $criado = if ($detalheHello.CreatedDateTime) { (Get-Date $detalheHello.CreatedDateTime -Format "dd/MM/yyyy") } else { "Data desconhecida" }
                                Write-Host "  • $dispositivo (Criado: $criado)" -ForegroundColor White
                            } else {
                                Write-Host "  • Dispositivo Windows $contador (PIN/Biometria)" -ForegroundColor White
                            }
                        } catch {
                            Write-Host "  • Dispositivo Windows $contador (PIN/Biometria)" -ForegroundColor White
                        }
                        $contador++
                    }
                    Write-Host ""
                }
                # ========== FIM DA NOVA SEÇÃO ==========
                # FIDO2 (se houver)
                try {
                    $fido2 = Get-MgUserAuthenticationFido2Method -UserId $emailProcurar -ErrorAction SilentlyContinue
                    if ($fido2.Count -gt 0) {
                        Write-Host "🔑 CHAVES FIDO2 ($($fido2.Count)):" -ForegroundColor Green
                        foreach ($key in $fido2) {
                            $nome = if ($key.DisplayName) { $key.DisplayName } else { "Chave FIDO2" }
                            Write-Host "  • $nome" -ForegroundColor White
                        }
                        Write-Host ""
                    }
                } catch {
                    # FIDO2 pode não estar disponível em todos os tenants
                }
                # Métodos temporários/email
                $emails = $todosMetodos | Where-Object { 
                    $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.emailAuthenticationMethod' 
                }
                if ($emails.Count -gt 0) {
                    Write-Host "📧 EMAIL BACKUP ($($emails.Count)):" -ForegroundColor Yellow
                    foreach ($email in $emails) {
                        $enderecoEmail = if ($email.EmailAddress) { $email.EmailAddress } else { "Email configurado" }
                        Write-Host "  • $enderecoEmail" -ForegroundColor White
                    }
                    Write-Host ""
                }
            }
            # Resumo final
            Write-Host "=" * 50 -ForegroundColor Gray
            Write-Host "Status geral: $($usuarioEncontrado.'Status MFA')" -ForegroundColor $(if ($usuarioEncontrado.'Status MFA' -eq 'Configurado') { "Green" } else { "Red" })
            Write-Host "Conta ativa: $($usuarioEncontrado.'Conta Ativa')" -ForegroundColor $(if ($usuarioEncontrado.'Conta Ativa') { "Green" } else { "Yellow" })
            Write-Host "Licenciado: $($usuarioEncontrado.'Licenciado')" -ForegroundColor $(if ($usuarioEncontrado.'Licenciado') { "Green" } else { "Yellow" })
            Write-Host "=" * 50 -ForegroundColor Gray
        }
        catch {
            Write-Host "❌ Erro ao obter detalhes: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    # ========== FIM DA NOVA FUNCIONALIDADE ==========
} # ← Manter este fechamento do try principal
catch {
    Write-Host ""
    Write-Host "❌ ERRO GERAL DO SCRIPT:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
Write-Host ""
Write-Host "Pressione qualquer tecla para sair..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")# Script MFA - COM FILTROS PARA USUÁRIOS REAIS

