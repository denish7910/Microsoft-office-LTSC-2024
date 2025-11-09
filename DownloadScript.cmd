@echo off
setlocal EnableExtensions EnableDelayedExpansion

:: ============================================================
:: Project : Office LTSC Professional Plus 2024 Installer
:: Author  : Denish Borad
:: Version : 1.1 (Color UI)
:: Date    : %date%
:: ============================================================

:: ---- Enable ANSI escape sequences (for colors in modern Windows 10/11) ----
for /F "tokens=1,2 delims==" %%a in ('"echo prompt $E^|cmd"') do (
  if "%%a"=="PROMPT" set "ESC=%%b"
)
:: Fallback to classic COLOR command if ESC isn't available
if not defined ESC (
  color 0A
)

:: ---- Color helpers (ANSI). If ESC not defined, these collapse to nothing ----
set "C_RESET=%ESC%[0m"
set "C_TITLE=%ESC%[38;2;180;200;255m"
set "C_SUB=%ESC%[38;2;140;160;255m"
set "C_OK=%ESC%[92m"
set "C_WARN=%ESC%[93m"
set "C_ERR=%ESC%[91m"
set "C_INFO=%ESC%[96m"
set "C_DIM=%ESC%[90m"
set "C_BOLD=%ESC%[1m"

title Office LTSC 2024 Installer - Denish Borad
cls

echo %C_TITLE%============================================================%C_RESET%
echo %C_TITLE%   MICROSOFT OFFICE LTSC PROFESSIONAL PLUS 2024 INSTALLER%C_RESET%
echo %C_TITLE%============================================================%C_RESET%
echo %C_SUB%   Developed By  :%C_RESET% Denish Borad
echo %C_SUB%   GitHub        :%C_RESET% https://github.com/denish7910
echo %C_SUB%   Version       :%C_RESET% 1.1
echo %C_TITLE%------------------------------------------------------------%C_RESET%
echo.

:: =================== Admin Check ===================
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo %C_WARN%[!] Admin privileges required. Requesting elevation...%C_RESET%
  powershell -NoProfile -Command "Start-Process '%~f0' -Verb RunAs"
  exit /b
)
echo %C_OK%[✓]%C_RESET% %C_INFO%Running with administrative privileges.%C_RESET%
echo.

:: =================== Spinner (nice touch) ===================
set "SPIN=\|/-"
set /a idx=0
set "DO_SPIN=call :spin"

:: =================== Paths ===================
set "SRC_DIR=%~dp0office2024"
set "DST_DIR=C:\office2024"

:: =================== Copy Files ===================
echo %C_INFO%Preparing setup environment...%C_RESET%
if not exist "%SRC_DIR%" (
  echo %C_ERR%[x] Source folder not found:%C_RESET% "%SRC_DIR%"
  goto :end_fail
)

<nul set /p "= %C_INFO%Copying setup files to %DST_DIR% ...%C_RESET%"
%DO_SPIN% & xcopy "%SRC_DIR%" "%DST_DIR%" /E /I /Y >nul
if %errorlevel% neq 0 (
  echo.
  echo %C_ERR%[x] Failed to copy files to "%DST_DIR%".%C_RESET%
  goto :end_fail
)
echo %C_OK%  [done]%C_RESET%

:: =================== Validate & Run ===================
cd /d "%DST_DIR%" >nul 2>&1
if not exist "%DST_DIR%\setup.exe" (
  echo %C_ERR%[x] setup.exe not found in "%DST_DIR%".%C_RESET%
  goto :end_fail
)

echo.
echo %C_INFO%Launching Office setup...%C_RESET%
echo %C_DIM%Command:%C_RESET% setup.exe /configure configuration.xml
%DO_SPIN% & setup.exe /configure configuration.xml
set "RC=%errorlevel%"

echo.
if %RC% equ 0 (
  echo %C_OK%[✓] Installation process completed or launched successfully.%C_RESET%
  goto :end_ok
) else (
  echo %C_ERR%[x] Setup exited with code %RC%.%C_RESET%
  goto :end_fail
)

:: =================== Spinner Subroutine ===================
:spin
  set /a idx=(idx+1)%%4
  set "ch=!SPIN:~%idx%,1!"
  <nul set /p "=?"
  >nul timeout /t 1 /nobreak
  <nul set /p "="  :: backspace
  exit /b

:: =================== End Sections ===================
:end_ok
echo %C_DIM%Press any key to exit...%C_RESET%
pause >nul
exit /b 0

:end_fail
echo %C_WARN%Tip:%C_RESET% Check that configuration.xml is valid and paths are correct.
echo %C_DIM%Press any key to exit...%C_RESET%
pause >nul
exit /b 1
