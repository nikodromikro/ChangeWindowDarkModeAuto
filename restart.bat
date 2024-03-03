@echo off

# Set the theme name you want to apply
set themeName="Windows Light"

# Change the theme
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes" /v "CurrentTheme" /t REG_SZ /d "%themeName%" /f

# Toggle system-wide light theme usage setting
for /f "tokens=3" %%i in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" /v "SystemUsesLightTheme"') do set systemUsesLightTheme=%%i
if %systemUsesLightTheme% equ 1 (
    reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" /v "SystemUsesLightTheme" /t REG_DWORD /d 0 /f
) else (
    reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" /v "SystemUsesLightTheme" /t REG_DWORD /d 1 /f
)

# Check current app mode
for /f "tokens=3" %%i in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" /v "AppsUseLightTheme"') do set currentMode=%%i

# Toggle app mode
if %currentMode% equ 1 (
    reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" /v "AppsUseLightTheme" /t REG_DWORD /d 0 /f
) else (
    reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" /v "AppsUseLightTheme" /t REG_DWORD /d 1 /f
)

# Notify the system of the theme and mode change
RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters ,1 ,True

# Refresh Windows Explorer windows to apply the theme changes
powershell -Command "& { (New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Refresh() } }"
