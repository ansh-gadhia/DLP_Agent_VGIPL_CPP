# CyberSentinel DLP Agent - Uninstall Script
# This script removes the agent service and optionally deletes all files

#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

# Configuration
$SERVICE_NAME = "CyberSentinelDLP"
$INSTALL_DIR = "C:\Program Files\CyberSentinel"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   CyberSentinel DLP Agent - Uninstall Script" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Check if running as Administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    pause
    exit 1
}

# Check if service exists
Write-Host "Checking for installed service..." -ForegroundColor Yellow
$service = Get-Service -Name $SERVICE_NAME -ErrorAction SilentlyContinue

if (-not $service) {
    Write-Host "[ERROR] Service '$SERVICE_NAME' not found!" -ForegroundColor Yellow
    Write-Host "The service may have already been uninstalled." -ForegroundColor Yellow
    
    # Check if installation directory exists
    if (Test-Path $INSTALL_DIR) {
        Write-Host "`nInstallation directory still exists: $INSTALL_DIR" -ForegroundColor Yellow
        $response = Read-Host "Do you want to delete the installation directory? (Y/N)"
        if ($response -eq 'Y' -or $response -eq 'y') {
            try {
                Remove-Item -Path $INSTALL_DIR -Recurse -Force
                Write-Host "[OK] Installation directory deleted" -ForegroundColor Green
            }
            catch {
                Write-Host "[ERROR] Failed to delete directory: $_" -ForegroundColor Red
            }
        }
    }
    
    pause
    exit 0
}

Write-Host "[OK] Service found: $SERVICE_NAME" -ForegroundColor Green
Write-Host "  Status: $($service.Status)" -ForegroundColor White
Write-Host "  Display Name: $($service.DisplayName)" -ForegroundColor White
Write-Host ""

# Confirm uninstallation
Write-Host "WARNING: This will remove the CyberSentinel DLP Agent service!" -ForegroundColor Red
Write-Host ""
$response = Read-Host "Are you sure you want to continue? (Y/N)"

if ($response -ne 'Y' -and $response -ne 'y') {
    Write-Host "Uninstallation cancelled." -ForegroundColor Yellow
    pause
    exit 0
}

# Step 1: Stop the service
Write-Host "`n[Step 1/3] Stopping service..." -ForegroundColor Cyan
try {
    if ($service.Status -eq 'Running') {
        Stop-Service -Name $SERVICE_NAME -Force
        Start-Sleep -Seconds 2
        Write-Host "[OK] Service stopped" -ForegroundColor Green
    }
    else {
        Write-Host "[OK] Service already stopped" -ForegroundColor Green
    }
}
catch {
    Write-Host "[WARNING] Could not stop service: $_" -ForegroundColor Yellow
}

# Step 2: Remove the service
Write-Host "`n[Step 2/3] Removing service..." -ForegroundColor Cyan
$nssmPath = Join-Path $INSTALL_DIR "nssm.exe"

if (Test-Path $nssmPath) {
    try {
        & $nssmPath remove $SERVICE_NAME confirm
        Write-Host "[OK] Service removed successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "[WARNING] NSSM removal had issues: $_" -ForegroundColor Yellow
        
        # Fallback: try using sc.exe
        Write-Host "Attempting fallback removal method..." -ForegroundColor Yellow
        $output = sc.exe delete $SERVICE_NAME
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] Service removed using sc.exe" -ForegroundColor Green
        }
        else {
            Write-Host "[ERROR] Failed to remove service: $output" -ForegroundColor Red
        }
    }
}
else {
    # Try using sc.exe directly
    Write-Host "NSSM not found, using sc.exe..." -ForegroundColor Yellow
    $output = sc.exe delete $SERVICE_NAME
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[OK] Service removed" -ForegroundColor Green
    }
    else {
        Write-Host "[ERROR] Failed to remove service: $output" -ForegroundColor Red
    }
}

# Wait a moment for service to be removed
Start-Sleep -Seconds 2

# Verify service removal
$serviceCheck = Get-Service -Name $SERVICE_NAME -ErrorAction SilentlyContinue
if ($serviceCheck) {
    Write-Host "[WARNING] Service still exists in registry" -ForegroundColor Yellow
    Write-Host "You may need to restart your computer for complete removal" -ForegroundColor Yellow
}
else {
    Write-Host "[OK] Service successfully removed from system" -ForegroundColor Green
}

# Step 3: Remove installation directory
Write-Host "`n[Step 3/3] Removing installation files..." -ForegroundColor Cyan

if (Test-Path $INSTALL_DIR) {
    Write-Host "Installation directory: $INSTALL_DIR" -ForegroundColor White
    Write-Host ""
    Write-Host "Options:" -ForegroundColor Yellow
    Write-Host "1. Delete all files and directory (recommended)" -ForegroundColor White
    Write-Host "2. Keep configuration file (agent_config.json) and logs" -ForegroundColor White
    Write-Host "3. Keep everything (skip deletion)" -ForegroundColor White
    Write-Host ""
    
    $deleteChoice = Read-Host "Enter your choice (1-3)"
    
    switch ($deleteChoice) {
        "1" {
            # Delete everything
            try {
                Remove-Item -Path $INSTALL_DIR -Recurse -Force
                Write-Host "[OK] All files and directory deleted" -ForegroundColor Green
            }
            catch {
                Write-Host "[ERROR] Error deleting directory: $_" -ForegroundColor Red
                Write-Host "You may need to manually delete: $INSTALL_DIR" -ForegroundColor Yellow
            }
        }
        
        "2" {
            # Keep config and logs
            try {
                $configPath = Join-Path $INSTALL_DIR "agent_config.json"
                $logPath = Join-Path $INSTALL_DIR "*.log"
                $backupDir = Join-Path $env:USERPROFILE "Desktop\CyberSentinel_Backup"
                
                # Create backup directory
                New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
                
                # Copy config and logs
                if (Test-Path $configPath) {
                    Copy-Item -Path $configPath -Destination $backupDir -Force
                    Write-Host "[OK] Configuration backed up to: $backupDir" -ForegroundColor Green
                }
                
                Get-ChildItem -Path $INSTALL_DIR -Filter "*.log" | ForEach-Object {
                    Copy-Item -Path $_.FullName -Destination $backupDir -Force
                }
                Write-Host "[OK] Logs backed up to: $backupDir" -ForegroundColor Green
                
                # Delete installation directory
                Remove-Item -Path $INSTALL_DIR -Recurse -Force
                Write-Host "[OK] Installation directory deleted" -ForegroundColor Green
            }
            catch {
                Write-Host "[ERROR] Error during selective deletion: $_" -ForegroundColor Red
            }
        }
        
        "3" {
            # Keep everything
            Write-Host "[OK] Installation files kept at: $INSTALL_DIR" -ForegroundColor Yellow
        }
        
        default {
            Write-Host "[ERROR] Invalid choice, keeping all files" -ForegroundColor Yellow
        }
    }
}
else {
    Write-Host "[OK] Installation directory not found (already deleted)" -ForegroundColor Green
}

# Summary
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "   Uninstallation Complete!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Check if anything is left
$serviceStillExists = Get-Service -Name $SERVICE_NAME -ErrorAction SilentlyContinue
$dirStillExists = Test-Path $INSTALL_DIR

if ($serviceStillExists) {
    Write-Host "[WARNING] Service: Still exists (may require restart)" -ForegroundColor Yellow
}
else {
    Write-Host "[OK] Service: Removed" -ForegroundColor Green
}

if ($dirStillExists) {
    Write-Host "[WARNING] Files: Still present at $INSTALL_DIR" -ForegroundColor Yellow
}
else {
    Write-Host "[OK] Files: Deleted" -ForegroundColor Green
}

Write-Host ""

if ($serviceStillExists -or ($dirStillExists -and $deleteChoice -eq "1")) {
    Write-Host "Note: A system restart may be required for complete removal." -ForegroundColor Yellow
}

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

pause