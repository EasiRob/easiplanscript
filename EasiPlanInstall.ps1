<#
.SYNOPSIS
    Installs the EasiPlanAgent from an MSI package on a network share.
    Also ensures the scheduled task is created with a daily trigger.
    This script is designed to run as a GPO Computer Startup Script.
#>

# ============================================================================
# --- Configuration: To Install EasiPlan You MUST Edit items 1 and 2 below ---
# ============================================================================

# 1. Your unique API key
$APIKEY_VALUE = "c******************"

# 2. The *network share* (UNC path) where "EasiPlanAgent.msi" is located
$INSTALLER_SHARE = "\\YourServer\YourShare"

# 3. The full path to a file that proves the agent is installed
#    (*** You MUST check and update this path ***)
$CHECK_FILE = "C:\Program Files\EasiPlanDeviceAgent\DeviceInfoAgent.exe"

# 4. A folder for the installation log files (optional, but recommended)
$LOG_PATH = "C:\Temp\InstallLogs"

# ==================================================================
# --- Do Not Edit Below This Line ---
# ==================================================================

# --- Set up full paths ---
$MsiFile = Join-Path -Path $INSTALLER_SHARE -ChildPath "EasiPlanAgent.msi"
$MsiLogFile = Join-Path -Path $LOG_PATH -ChildPath "EasiPlanAgent_install_log.txt"
$ScriptLogFile = Join-Path -Path $LOG_PATH -ChildPath "EasiPlanAgent_script_log.txt"

# --- Create log directory ---
if (-not (Test-Path -Path $LOG_PATH)) {
    try {
        New-Item -ItemType Directory -Path $LOG_PATH -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Error "Failed to create log directory at $LOG_PATH. Exiting."
        exit 1
    }
}

# --- Logging function ---
function Write-Log {
    param (
        [string]$Message
    )
    $Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $LogLine = "${Timestamp}: $Message"
    try {
        Add-Content -Path $ScriptLogFile -Value $LogLine -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to write to log file: $ScriptLogFile"
    }
}

# --- Start Script ---
Write-Log -Message "--- EasiPlanAgent Install Script Started (PowerShell) ---"

# --- 1. Check if software is installed ---
if (Test-Path -Path $CHECK_FILE) {
    Write-Log -Message "Check file found: $CHECK_FILE. Software is installed."
}
else {
    Write-Log -Message "Check file not found. Starting installation..."

    # --- 2. Check if installer exists ---
    if (-not (Test-Path -Path $MsiFile)) {
        Write-Log -Message "ERROR: Installer not found at: $MsiFile"
        Write-Log -Message "Script cannot continue. Exiting."
        Write-Log -Message "--- Script Finished ---"
        exit 1
    }

    # --- 3. Run the installation ---
    Write-Log -Message "Starting installation from: $MsiFile"
    Write-Log -Message "Applying APIKEY..."

    # Build the argument list for msiexec.exe
    $MsiArguments = @(
        "/i",
        "`"$MsiFile`"",
        "APIKEY=`"$APIKEY_VALUE`"",
        "/qn",
        "/L*v",
        "`"$MsiLogFile`""
    )

    try {
        # Start msiexec, wait for it to complete (-Wait), and get the process object (-PassThru)
        $InstallProcess = Start-Process -FilePath "msiexec.exe" -ArgumentList $MsiArguments -Wait -PassThru -ErrorAction Stop
        $ExitCode = $InstallProcess.ExitCode
    }
    catch {
        Write-Log -Message "INSTALLER FAILED to launch. Error: $_"
        Write-Log -Message "--- Script Finished ---"
        exit 1
    }

    # --- 4. Check install result ---
    if ($ExitCode -eq 0) {
        Write-Log -Message "Installer finished successfully (Code 0)."
    }
    elseif ($ExitCode -eq 3010) {
        Write-Log -Message "Installer finished with Code 3010 (Reboot Required)."
    }
    else {
        Write-Log -Message "INSTALLER FAILED with Error Code: $ExitCode."
        Write-Log -Message "Check the MSI log for details: $MsiLogFile"
        Write-Log -Message "--- Script Finished ---"
        exit $ExitCode # Exit if the install failed, no point trying to make the task
    }
}

# --- 5. Check and create Scheduled Task ---
$TaskName = "EasiPlan Agent Task"
$TaskPath = "\EasiPlan" # Optional: Puts it in a folder in Task Scheduler
$TaskExists = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue

if ($TaskExists) {
    Write-Log -Message "Scheduled Task '$TaskName' already exists. No action needed."
}
else {
    Write-Log -Message "Scheduled Task '$TaskName' not found. Creating it..."
    
    # Ensure the check file exists before trying to create a task for it
    if (-not (Test-Path -Path $CHECK_FILE)) {
        Write-Log -Message "ERROR: Cannot create task because the executable is missing: $CHECK_FILE"
        Write-Log -Message "--- Script Finished ---"
        exit 1
    }
    
    try {
        # --- Define Task Action with Arguments and "Start In" ---
        
        # 1. Define Arguments (using the API key from the top of the script)
        $TaskArgs = "--apiurl `"https://ingestdevicedata-plleqvdknq-nw.a.run.app`" --apikey `"$APIKEY_VALUE`""
        
        # 2. Define "Start In" (Working Directory) by getting the parent folder of the exe
        $TaskWorkDir = Split-Path -Path $CHECK_FILE -Parent
        
        # 3. Create the Action
        $TaskAction = New-ScheduledTaskAction -Execute $CHECK_FILE -Argument $TaskArgs -WorkingDirectory $TaskWorkDir
        
        # --- Define Trigger (Daily at 10am) ---
        $TaskTrigger = New-ScheduledTaskTrigger -Daily -At "10:00am"
        
        # --- Define Principal (who to run as) ---
        $TaskPrincipal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -RunLevel Highest

        # --- Define Settings (to run if missed) ---
        $TaskSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable
        
        # --- Register the task ---
        Register-ScheduledTask -TaskName $TaskName `
                               -TaskPath $TaskPath `
                               -Action $TaskAction `
                               -Trigger $TaskTrigger `
                               -Principal $TaskPrincipal `
                               -Settings $TaskSettings `
                               -ErrorAction Stop
        
        Write-Log -Message "Successfully created Scheduled Task: '$TaskName'."
    }
    catch {
        Write-Log -Message "ERROR: Failed to create Scheduled Task. Error: $_"
    }
}

Write-Log -Message "--- Script Finished ---"
exit 0
