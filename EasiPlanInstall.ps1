@echo off

:: ===========================================================================
:: --- Configuration: To Install EasiPlan You MUST Edit the 4 items below ---
:: ===========================================================================

: 1. Your unique API key
SET "APIKEY_VALUE=c******************"

:: 2. The *network share* (UNC path) where "EasiPlanAgent.msi" is located
SET "INSTALLER_SHARE=\\YourServer\YourShare"

:: 3. The full path to a file that proves the agent is installed
::    (This prevents the script from running on every reboot)
::    (*** You MUST check and update this path ***)
SET "CHECK_FILE=C:\Program Files\EasiPlanAgent\agent.exe"

:: 4. A folder for the installation log files (optional, but recommended)
SET "LOG_PATH=C:\Temp\InstallLogs"

:: ==================================================================
:: --- Do Not Edit Below This Line ---
:: ==================================================================

:: --- Set up full paths ---
SET "MsiFile=%INSTALLER_SHARE%\EasiPlanAgent.msi"
SET "MsiLogFile=%LOG_PATH%\EasiPlanAgent_install_log.txt"
SET "ScriptLogFile=%LOG_PATH%\EasiPlanAgent_script_log.txt"

:: --- Create log directory ---
if not exist "%LOG_PATH%" mkdir "%LOG_PATH%"

:: --- Logging function ---
:log
echo %DATE% %TIME%: %* >> "%ScriptLogFile%"
goto :eof

:: --- Start Script ---
call :log "--- EasiPlanAgent Install Script Started ---"

:: --- 1. Check if already installed ---
if exist "%CHECK_FILE%" (
    call :log "Check file found: %CHECK_FILE%"
    call :log "Software is already installed. Exiting."
    goto :End
)

:: --- 2. Check if installer exists ---
if not exist "%MsiFile%" (
    call :log "ERROR: Installer not found at: %MsiFile%"
    call :log "Script cannot continue. Exiting."
    goto :End
)

:: --- 3. Run the installation ---
call :log "Starting installation from: %MsiFile%"
call :log "Applying APIKEY..."

:: This runs your command silently (/qn) and creates a verbose log (/L*v)
msiexec /i "%MsiFile%" APIKEY="%APIKEY_VALUE%" /qn /L*v "%MsiLogFile%"

:: --- 4. Check result ---
if %ERRORLEVEL% == 0 (
    call :log "Installer finished successfully (Code 0)."
) else if %ERRORLEVEL% == 3010 (
    call :log "Installer finished with Code 3010 (Reboot Required)."
) else (
    call :log "INSTALLER FAILED with Error Code: %ERRORLEVEL%."
    call :log "Check the MSI log for details: %MsiLogFile%"
)

:End
call :log "--- Script Finished ---"
goto :eof