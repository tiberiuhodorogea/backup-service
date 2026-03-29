#Requires -Version 5.1
<#
.SYNOPSIS
    Builds a self-contained release of Backup Service.

.DESCRIPTION
    Run this ONCE on the development machine (internet required).
    It downloads a portable Python 3.12 runtime, installs all
    dependencies into it, then copies everything into .\Release\.

    Copy the Release\ folder to a USB stick.
    On the target machine: right-click install.bat -> Run as administrator.

.NOTES
    No Python needs to be installed on the target machine.
    The release is fully offline once built.
#>

$ErrorActionPreference = "Stop"
$ProgressPreference    = "SilentlyContinue"

# ---- Config -----------------------------------------------------------------
$PYTHON_VERSION   = "3.12.8"
$PYTHON_ARCH      = "amd64"
$PYTHON_EMBED_URL = "https://www.python.org/ftp/python/$PYTHON_VERSION/python-$PYTHON_VERSION-embed-$PYTHON_ARCH.zip"
$GET_PIP_URL      = "https://bootstrap.pypa.io/get-pip.py"

# Where install.bat will place the service on the TARGET machine
$INSTALL_DIR = "C:\BackupService"

# ---- Paths (on this build machine) ------------------------------------------
$ROOT    = $PSScriptRoot
$RELEASE = Join-Path $ROOT "Release"
$PY_DIR  = Join-Path $RELEASE "python"

# ---- Header -----------------------------------------------------------------
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Backup Service - Build Release" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Python : $PYTHON_VERSION ($PYTHON_ARCH)"
Write-Host "  Output : $RELEASE"
Write-Host "================================================================"
Write-Host ""

# ---- Clean previous build ---------------------------------------------------
if (Test-Path $RELEASE) {
    Write-Host "[1/6] Cleaning previous release..."
    Remove-Item $RELEASE -Recurse -Force
}
New-Item -ItemType Directory $RELEASE | Out-Null
New-Item -ItemType Directory $PY_DIR  | Out-Null

# ---- Download embedded Python -----------------------------------------------
Write-Host "[2/6] Downloading Python $PYTHON_VERSION embeddable (~10 MB)..."
$zipPath = Join-Path $env:TEMP "python-$PYTHON_VERSION-embed-$PYTHON_ARCH.zip"
try {
    Invoke-WebRequest $PYTHON_EMBED_URL -OutFile $zipPath
} catch {
    throw "Failed to download Python: $_`nURL: $PYTHON_EMBED_URL"
}
Expand-Archive $zipPath $PY_DIR -Force
Remove-Item $zipPath

# Embedded Python ships with import site disabled - re-enable it so pip works.
$pthFile = Get-ChildItem $PY_DIR -Filter "python*._pth" | Select-Object -First 1
if (-not $pthFile) { throw "Cannot find python*._pth inside embedded Python zip." }
Write-Host "       Patching $($pthFile.Name) to enable site-packages..."
(Get-Content $pthFile.FullName) -replace '#import site', 'import site' | Set-Content $pthFile.FullName

# ---- Bootstrap pip ----------------------------------------------------------
Write-Host "[3/6] Bootstrapping pip..."
$getPipPath = Join-Path $env:TEMP "get-pip.py"
try {
    Invoke-WebRequest $GET_PIP_URL -OutFile $getPipPath
} catch {
    throw "Failed to download get-pip.py: $_"
}
& "$PY_DIR\python.exe" $getPipPath --no-warn-script-location --quiet
if ($LASTEXITCODE -ne 0) { throw "get-pip.py failed (exit $LASTEXITCODE)" }
Remove-Item $getPipPath

# ---- Install packages -------------------------------------------------------
Write-Host "[4/6] Installing packages (may take a minute)..."
& "$PY_DIR\python.exe" -m pip install `
    -r "$ROOT\requirements.txt" `
    --no-warn-script-location `
    --quiet
if ($LASTEXITCODE -ne 0) { throw "pip install failed (exit $LASTEXITCODE)" }

# ---- Fix pywin32 DLLs -------------------------------------------------------
# pythoncom312.dll and pywintypes312.dll must be on the DLL search path at
# service start-up. Copying them next to python.exe is the simplest fix.
Write-Host "[5/6] Copying pywin32 DLLs to Python root..."
$pywin32sys = Join-Path $PY_DIR "Lib\site-packages\pywin32_system32"
if (Test-Path $pywin32sys) {
    Get-ChildItem "$pywin32sys\*.dll" | ForEach-Object {
        Copy-Item $_.FullName $PY_DIR
        Write-Host "       $($_.Name)"
    }
} else {
    Write-Warning "pywin32_system32 directory not found - DLLs may be missing on the target."
}

# ---- Copy source files ------------------------------------------------------
Write-Host "[6/6] Copying source files..."
@("main.py", "backup_engine.py", "database.py", "service.py", "requirements.txt") | ForEach-Object {
    $src = Join-Path $ROOT $_
    if (-not (Test-Path $src)) { throw "Source file not found: $src" }
    Copy-Item $src $RELEASE
}
Copy-Item (Join-Path $ROOT "static") (Join-Path $RELEASE "static") -Recurse

# ---- Write install.bat ------------------------------------------------------
Write-Host "       Writing install.bat..."
$installBat = @"
@echo off
setlocal
:: ============================================================
::  Backup Service - Install
::
::  Run this as Administrator.
::  Copies the service to $INSTALL_DIR, registers it as a
::  Windows Service that starts automatically on boot, and
::  opens port 8550 in the Windows Firewall.
::
::  Dashboard will be available at:
::    http://localhost:8550
::    http://<this-PC-name>:8550  (LAN access)
:: ============================================================

set "INSTALL_DIR=$INSTALL_DIR"
set "PY=%INSTALL_DIR%\python\python.exe"
set "SVC=%INSTALL_DIR%\service.py"
set "LOGDIR=%INSTALL_DIR%\logs"
set "LOGFILE=%INSTALL_DIR%\logs\install.log"

echo.
echo ============================================================
echo   Backup Service - Install
echo ============================================================
echo.

:: -- Admin check ----------------------------------------------------------
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo   ERROR: This script must be run as Administrator.
    echo   Right-click install.bat and select "Run as administrator".
    echo.
    pause
    exit /b 1
)

:: -- Remove any previous installation ------------------------------------
sc query BackupService >nul 2>&1
if %errorlevel% equ 0 (
    echo   Previous installation found - removing it first...
    sc stop BackupService >nul 2>&1
    timeout /t 5 /nobreak >nul
    if exist "%PY%" "%PY%" "%SVC%" remove >nul 2>&1
)

:: -- Copy files -----------------------------------------------------------
:: Note: rmdir happens before any log file exists, so nothing is lost.
echo   Copying files to %INSTALL_DIR% ...
if exist "%INSTALL_DIR%" rmdir /s /q "%INSTALL_DIR%"
xcopy /e /i /q "%~dp0" "%INSTALL_DIR%\" >nul
if %errorlevel% neq 0 (
    echo   ERROR: File copy failed. Make sure no files are open or locked.
    echo.
    pause
    exit /b 1
)

:: -- Create logs dir and open install log (safe now - files are in place) -
mkdir "%LOGDIR%" >nul 2>&1
echo ============================================================ >> "%LOGFILE%"
echo Install started : %DATE% %TIME%                              >> "%LOGFILE%"
echo Machine         : %COMPUTERNAME%                             >> "%LOGFILE%"
echo Source          : %~dp0                                      >> "%LOGFILE%"
echo Destination     : %INSTALL_DIR%                              >> "%LOGFILE%"
echo ============================================================ >> "%LOGFILE%"
echo [%TIME%] Files copied to %INSTALL_DIR% >> "%LOGFILE%"

:: -- Register Windows Service ---------------------------------------------
echo   Registering Windows Service...
echo [%TIME%] Registering Windows Service >> "%LOGFILE%"
"%INSTALL_DIR%\python\python.exe" "%INSTALL_DIR%\service.py" install >> "%LOGFILE%" 2>&1
:: Verify via sc query - more reliable than pywin32 exit code
sc query BackupService >nul 2>&1
if %errorlevel% neq 0 (
    echo   ERROR: Service registration failed.
    echo   Check: %LOGFILE%
    echo [%TIME%] ERROR: Service not found after install attempt >> "%LOGFILE%"
    echo.
    pause
    exit /b 1
)
echo [%TIME%] Service registered and verified >> "%LOGFILE%"

:: -- Ensure auto-start on boot --------------------------------------------
sc config BackupService start= auto >nul
echo [%TIME%] Service set to auto-start >> "%LOGFILE%"

:: -- Firewall rule --------------------------------------------------------
echo   Adding Windows Firewall rule for port 8550...
echo [%TIME%] Adding firewall rule for port 8550 >> "%LOGFILE%"
netsh advfirewall firewall delete rule name="BackupService Dashboard" >nul 2>&1
netsh advfirewall firewall add rule name="BackupService Dashboard" dir=in action=allow protocol=TCP localport=8550 description="Backup Service web dashboard" >> "%LOGFILE%" 2>&1

:: -- Start ----------------------------------------------------------------
echo   Starting service...
echo [%TIME%] Starting service >> "%LOGFILE%"
"%INSTALL_DIR%\python\python.exe" "%INSTALL_DIR%\service.py" start >> "%LOGFILE%" 2>&1
if %errorlevel% neq 0 (
    echo   WARNING: Service registered but did not start immediately.
    echo   Check: %LOGFILE%
    echo   Or:    Event Viewer ^> Windows Logs ^> Application
    echo [%TIME%] WARNING: Service start returned non-zero exit code >> "%LOGFILE%"
) else (
    echo [%TIME%] Service started successfully >> "%LOGFILE%"
)

echo [%TIME%] Install complete >> "%LOGFILE%"
echo ============================================================ >> "%LOGFILE%"

echo.
echo ============================================================
echo   Installation complete!
echo.
echo   Dashboard : http://localhost:8550
echo   LAN access: http://%COMPUTERNAME%:8550
echo.
echo   Logs folder: %LOGDIR%
echo     install.log  - this install session
echo     service.log  - runtime log (created on first service start)
echo.
echo   The service starts automatically on every boot.
echo   To manage: Services (services.msc) ^> Backup Service
echo ============================================================
echo.
pause
endlocal
"@
$installBat | Set-Content (Join-Path $RELEASE "install.bat") -Encoding ASCII

# ---- Write uninstall.bat ----------------------------------------------------
Write-Host "       Writing uninstall.bat..."
$uninstallBat = @"
@echo off
setlocal
:: ============================================================
::  Backup Service - Uninstall
::
::  Run this as Administrator.
::  Stops the service, removes its Windows registration,
::  removes the firewall rule, and deletes $INSTALL_DIR.
::
::  Your backup archives on destination drives are NOT touched.
:: ============================================================

set "INSTALL_DIR=$INSTALL_DIR"

echo.
echo ============================================================
echo   Backup Service - Uninstall
echo ============================================================
echo.

:: -- Admin check ----------------------------------------------------------
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo   ERROR: This script must be run as Administrator.
    pause
    exit /b 1
)

:: -- Stop service ---------------------------------------------------------
echo   Stopping service...
sc stop BackupService >nul 2>&1
timeout /t 5 /nobreak >nul

:: -- Remove service registration ------------------------------------------
echo   Removing service registration...
if exist "%INSTALL_DIR%\python\python.exe" (
    "%INSTALL_DIR%\python\python.exe" "%INSTALL_DIR%\service.py" remove >nul 2>&1
) else (
    sc delete BackupService >nul 2>&1
)

:: -- Firewall rule --------------------------------------------------------
echo   Removing firewall rule...
netsh advfirewall firewall delete rule name="BackupService Dashboard" >nul 2>&1

:: -- Delete files ---------------------------------------------------------
echo   Deleting %INSTALL_DIR% ...
if exist "%INSTALL_DIR%" (
    rmdir /s /q "%INSTALL_DIR%"
    if exist "%INSTALL_DIR%" (
        echo   WARNING: Some files could not be deleted.
        echo   Delete %INSTALL_DIR% manually after reboot if needed.
    )
)

echo.
echo ============================================================
echo   Backup Service uninstalled.
echo   Your backup archives on destination drives are untouched.
echo ============================================================
echo.
pause
endlocal
"@
$uninstallBat | Set-Content (Join-Path $RELEASE "uninstall.bat") -Encoding ASCII

# ---- Write instructions.txt -------------------------------------------------
Write-Host "       Writing instructions.txt..."
$instructions = @"
Backup Service
==============
Automatic file backup service with a web dashboard accessible over LAN.


REQUIREMENTS
------------
- Windows 10 or 11 (64-bit)
- No Python or other software needed - everything is included


INSTALL
-------
1. Right-click install.bat
2. Select "Run as administrator"
3. Click Yes on the UAC prompt
4. Wait for "Installation complete!" then press any key

The service starts immediately and will start automatically on every boot.


ACCESS THE DASHBOARD
--------------------
Open a browser and go to:

  http://localhost:8550          (on this machine)
  http://<computer-name>:8550    (from any other PC on the same network)

The computer name is shown at the end of the install output.


FIRST-TIME SETUP
----------------
1. Open the dashboard
2. Under SOURCES: add the folders you want to back up (use Browse or type the path)
3. Under DESTINATIONS: add where backups should go (local drive, USB, or UNC path like \\NAS\backup)
4. Under SETTINGS: set how often to run and how many backups to keep
5. Click "Save Settings"
6. Optionally click "Run Now" to start an immediate backup


UNINSTALL
---------
1. Right-click uninstall.bat
2. Select "Run as administrator"

This removes the service and deletes C:\BackupService.
Your backup archives on destination drives are NOT touched.


TROUBLESHOOT
------------
- Dashboard not loading: open Services (services.msc), find "Backup Service", check its status
- Service won't start: open Event Viewer > Windows Logs > Application and look for BackupService errors
- Can't reach from another PC: make sure the firewall rule was added (install.bat does this automatically)
  To add it manually:
    netsh advfirewall firewall add rule name="BackupService Dashboard" dir=in action=allow protocol=TCP localport=8550
"@
$instructions | Set-Content (Join-Path $RELEASE "instructions.txt") -Encoding ASCII

# ---- Zip the release --------------------------------------------------------
$ZIP_PATH = Join-Path $ROOT "backup-service-release.zip"
Write-Host "Zipping release..."
if (Test-Path $ZIP_PATH) { Remove-Item $ZIP_PATH -Force }
Compress-Archive -Path "$RELEASE\*" -DestinationPath $ZIP_PATH
$zipMB = [math]::Round((Get-Item $ZIP_PATH).Length / 1MB, 1)
Write-Host "       $ZIP_PATH ($zipMB MB)"

# ---- Summary ----------------------------------------------------------------
$sizeMB = [math]::Round(
    (Get-ChildItem $RELEASE -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1MB, 1
)

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "  BUILD COMPLETE" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host "  Release folder : $RELEASE ($sizeMB MB)"
Write-Host "  Release zip    : $ZIP_PATH ($zipMB MB)"
Write-Host ""
Write-Host "  NEXT STEPS:"
Write-Host "    1. Copy Release\ or the .zip to a USB stick."
Write-Host "    2. On the target machine:"
Write-Host "       Right-click install.bat -> Run as administrator"
Write-Host "    3. Open a browser to http://localhost:8550"
Write-Host ""
Write-Host "  To uninstall: right-click uninstall.bat -> Run as administrator"
Write-Host "================================================================"
Write-Host ""
