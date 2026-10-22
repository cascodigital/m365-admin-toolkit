<#
.SYNOPSIS
    Forces the ABNT2 (pt-BR) keyboard layout on Windows with en-US display language and disables layout-switch hotkeys.

.DESCRIPTION
    Fixes unwanted automatic keyboard layout switching on Windows 10/11 when the
    English display language is installed. Applies three changes:

    1. Sets the user language list to en-US with the ABNT2 input method only (0416:00000416)
    2. Disables the Ctrl+Shift and Alt+Shift switch hotkeys in the registry (Hotkey=3 = disabled)
    3. Sets InputMethodOverride in the user profile to ABNT2

    A sign-out/sign-in is required after running for the registry changes to take effect.

.EXAMPLE
    .\Fix-KeyboardLayout.ps1
    # Applies the changes; sign out and back in afterwards.

.EXAMPLE
    PowerShell -ExecutionPolicy Bypass -File .\Fix-KeyboardLayout.ps1
    # Runs once without permanently changing the execution policy.

.NOTES
    Version       : 1.0
    Requirements  : PowerShell 5.1+, Windows 10/11
    Action        : Sign-out/sign-in required after running
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
