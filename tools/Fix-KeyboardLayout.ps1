<#
.SYNOPSIS
    Forca layout ABNT2 (PT-BR) em Windows com idioma EN-US e desativa hotkeys de troca.

.DESCRIPTION
    Resolve o problema de troca automatica de layout de teclado em Windows 10/11
    com idioma ingles instalado. Aplica tres correcoes:

    1. Configura Language List com EN-US + input ABNT2 (codigo 0416:00000416)
    2. Desativa atalhos Ctrl+Shift e Alt+Shift via registro (Hotkey=3 = desabilitado)
    3. Forca InputMethodOverride no perfil do usuario para ABNT2

    Requer logoff/login apos execucao para ativar as mudancas de registro.

.EXAMPLE
    .\Fix-KeyboardLayout.ps1
    # Aplica correcoes e solicita logoff

.EXAMPLE
    PowerShell -ExecutionPolicy Bypass -File .\Fix-KeyboardLayout.ps1
    # Execucao direta sem alterar politica permanentemente

.NOTES
    Autor         : Andre Kittler / Casco Digital
    Versao        : 1.0
    Requisitos    : PowerShell 5.1+, Windows 10/11
    Acao          : Logoff/Login obrigatorio apos execucao
#>


# 1. Configure Language List (Display English + Input ABNT2)
# Creates the language object for English (US)
$LangList = New-WinUserLanguageList en-US

# Clears default input methods (which would include US keyboard)
$LangList[0].InputMethodTips.Clear()

# Adds specifically ABNT2 keyboard (Code: 0416:00000416) to English language
$LangList[0].InputMethodTips.Add('0416:00000416')

# Applies the new list (This removes phantom US keyboards)
Set-WinUserLanguageList $LangList -Force

# 2. Disable Keyboard Switching Shortcuts (Ctrl+Shift / Alt+Shift)
# Prevents accidental switching during use
$RegPath = "HKCU:\Keyboard Layout\Toggle"
if (!(Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
Set-ItemProperty -Path $RegPath -Name "Hotkey" -Value "3"
Set-ItemProperty -Path $RegPath -Name "Language Hotkey" -Value "3"
Set-ItemProperty -Path $RegPath -Name "Layout Hotkey" -Value "3"

# 3. Force Override Input Method in User Profile
# Ensures Windows uses ABNT2 regardless of window language
$ProfilePath = "HKCU:\Control Panel\International\User Profile"
Set-ItemProperty -Path $ProfilePath -Name "InputMethodOverride" -Value "0416:00000416"

Write-Host "Configuration applied. Logoff/Login recommended to activate registry changes." -ForegroundColor Green
