# CyberSentinel DLP Agent - Upgrade/Modify Script
# Requires Administrator privileges

#Requires -RunAsAdministrator

# Configuration
$GITHUB_REPO = "ansh-gadhia/DLP_Agent_VGIPL_CPP"
$INSTALL_DIR = "C:\Program Files\CyberSentinel"
$EXE_NAME = "cybersentinel_agent.exe"
$CONFIG_NAME = "agent_config.json"
$VBS_NAME = "launch_agent.vbs"
$TASK_NAME = "CyberSentinel DLP Agent"
$PROCESS_NAME = "cybersentinel_agent"

# Colors for output
function Write-ColorOutput {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [ValidateSet("Info", "Success", "Warning", "Error")]
        [string]$Type = "Info"
    )
    
    switch ($Type) {
        "Info"    { Write-Host $Message -ForegroundColor Cyan }
        "Success" { Write-Host $Message -ForegroundColor Green }
        "Warning" { Write-Host $Message -ForegroundColor Yellow }
        "Error"   { Write-Host $Message -ForegroundColor Red }
    }
}

# Function to validate IP address
function Test-IPAddress {
    param([string]$IP)
    
    if ($IP -eq "localhost" -or $IP -eq "") {
        return $true
    }
    
    $isValid = $IP -match '^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
    return $isValid
}

# Function to validate positive integer
function Test-PositiveInteger {
    param([string]$Value)
    
    $num = 0
    if ([int]::TryParse($Value, [ref]$num)) {
        return $num -gt 0
    }
    return $false
}

# Function to load current configuration
function Get-CurrentConfig {
    $configPath = Join-Path $INSTALL_DIR $CONFIG_NAME
    
    if (Test-Path $configPath) {
        try {
            $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
            return $config
        } catch {
            Write-ColorOutput "Error reading configuration file: $($_.Exception.Message)" -Type "Error"
            return $null
        }
    }
    return $null
}

# Function to regenerate VBScript launcher
function Update-VBScriptLauncher {
    param([string]$ExePath)
    
    try {
        $vbsPath = Join-Path $INSTALL_DIR $VBS_NAME
        $vbsContent = @"
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "$ExePath", "--background", "$INSTALL_DIR", "runas", 0
"@
        
        $vbsContent | Out-File -FilePath $vbsPath -Encoding ASCII -Force
        Write-ColorOutput "VBScript launcher updated" -Type "Success"
        return $true
    } catch {
        Write-ColorOutput "Error updating VBScript launcher: $($_.Exception.Message)" -Type "Warning"
        return $false
    }
}

# Banner
Clear-Host
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  CyberSentinel DLP Agent - Upgrade/Modify Script          " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Check if agent is installed
if (-not (Test-Path $INSTALL_DIR)) {
    Write-ColorOutput "CyberSentinel Agent is not installed." -Type "Error"
    Write-ColorOutput "Please run the installation script first." -Type "Info"
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# Load current configuration
Write-ColorOutput "Loading current configuration..." -Type "Info"
$currentConfig = Get-CurrentConfig

if ($currentConfig) {
    Write-Host ""
    Write-Host "Current Configuration:" -ForegroundColor Yellow
    Write-Host "  Server URL: $($currentConfig.server_url)"
    Write-Host "  Agent ID: $($currentConfig.agent_id)"
    Write-Host "  Agent Name: $($currentConfig.agent_name)"
    Write-Host "  Heartbeat Interval: $($currentConfig.heartbeat_interval) seconds"
    Write-Host "  Policy Sync Interval: $($currentConfig.policy_sync_interval) seconds"
} else {
    Write-ColorOutput "Could not load current configuration" -Type "Warning"
}

Write-Host ""

# Main menu
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "What would you like to do?" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "1. Upgrade agent executable to latest version"
Write-Host "2. Modify configuration settings"
Write-Host "3. Both (upgrade + modify configuration)"
Write-Host "4. Cancel"
Write-Host ""

$choice = Read-Host "Enter your choice (1-4)"

switch ($choice) {
    "1" {
        $mode = "upgrade"
        Write-ColorOutput "Mode: Upgrade executable only" -Type "Info"
    }
    "2" {
        $mode = "modify"
        Write-ColorOutput "Mode: Modify configuration only" -Type "Info"
    }
    "3" {
        $mode = "both"
        Write-ColorOutput "Mode: Upgrade executable and modify configuration" -Type "Info"
    }
    "4" {
        Write-ColorOutput "Operation cancelled by user." -Type "Info"
        Read-Host "Press Enter to exit"
        exit 0
    }
    default {
        Write-ColorOutput "Invalid choice. Exiting." -Type "Error"
        Read-Host "Press Enter to exit"
        exit 1
    }
}

Write-Host ""

# ===========================
# UPGRADE EXECUTABLE
# ===========================
if ($mode -eq "upgrade" -or $mode -eq "both") {
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host "UPGRADE EXECUTABLE" -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host ""
    
    # Step 1: Stop the agent
    Write-ColorOutput "Step 1: Stopping agent..." -Type "Info"
    
    try {
        $process = Get-Process -Name $PROCESS_NAME -ErrorAction SilentlyContinue
        
        if ($process) {
            Write-ColorOutput "Stopping running process (PID: $($process.Id))..." -Type "Info"
            Stop-Process -Name $PROCESS_NAME -Force -ErrorAction Stop
            Start-Sleep -Seconds 3
            Write-ColorOutput "Agent stopped successfully" -Type "Success"
        } else {
            Write-ColorOutput "Agent is not currently running" -Type "Info"
        }
    } catch {
        Write-ColorOutput "Error stopping agent: $($_.Exception.Message)" -Type "Warning"
        Write-ColorOutput "Continuing with upgrade..." -Type "Info"
    }
    
    Write-Host ""
    
    # Step 2: Backup current executable
    Write-ColorOutput "Step 2: Backing up current executable..." -Type "Info"
    
    $exePath = Join-Path $INSTALL_DIR $EXE_NAME
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = Join-Path $INSTALL_DIR "$EXE_NAME.backup_$timestamp"
    
    try {
        if (Test-Path $exePath) {
            Copy-Item -Path $exePath -Destination $backupPath -Force
            Write-ColorOutput "Backup created: $backupPath" -Type "Success"
        }
    } catch {
        Write-ColorOutput "Error creating backup: $($_.Exception.Message)" -Type "Warning"
        Write-ColorOutput "Continuing without backup..." -Type "Warning"
    }
    
    Write-Host ""
    
    # Step 3: Download latest version
    Write-ColorOutput "Step 3: Downloading latest version from GitHub..." -Type "Info"
    
    try {
        $releaseUrl = "https://api.github.com/repos/$GITHUB_REPO/releases/latest"
        Write-ColorOutput "Fetching release information..." -Type "Info"
        
        $release = Invoke-RestMethod -Uri $releaseUrl -Headers @{
            "User-Agent" = "CyberSentinel-Upgrader"
        }
        
        $version = $release.tag_name
        Write-ColorOutput "Latest version: $version" -Type "Success"
        
        $asset = $release.assets | Where-Object { $_.name -eq $EXE_NAME }
        
        if (-not $asset) {
            Write-ColorOutput "Error: Could not find $EXE_NAME in release assets." -Type "Error"
            
            # Restore backup if exists
            if (Test-Path $backupPath) {
                Copy-Item -Path $backupPath -Destination $exePath -Force
                Write-ColorOutput "Restored from backup" -Type "Info"
            }
            
            exit 1
        }
        
        $downloadUrl = $asset.browser_download_url
        $tempFile = Join-Path $env:TEMP $EXE_NAME
        
        Write-ColorOutput "Downloading from: $downloadUrl" -Type "Info"
        
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($downloadUrl, $tempFile)
        
        Write-ColorOutput "Download completed!" -Type "Success"
        
    } catch {
        Write-ColorOutput "Error downloading: $($_.Exception.Message)" -Type "Error"
        
        # Restore backup if exists
        if (Test-Path $backupPath) {
            Copy-Item -Path $backupPath -Destination $exePath -Force
            Write-ColorOutput "Restored from backup" -Type "Info"
        }
        
        exit 1
    }
    
    Write-Host ""
    
    # Step 4: Replace executable
    Write-ColorOutput "Step 4: Installing new version..." -Type "Info"
    
    try {
        Start-Sleep -Seconds 2  # Wait for file handles to release
        Copy-Item -Path $tempFile -Destination $exePath -Force
        Write-ColorOutput "Executable updated successfully!" -Type "Success"
        
        # Clean up temp file
        Remove-Item -Path $tempFile -Force
        
    } catch {
        Write-ColorOutput "Error installing new version: $($_.Exception.Message)" -Type "Error"
        
        # Restore backup
        if (Test-Path $backupPath) {
            Copy-Item -Path $backupPath -Destination $exePath -Force
            Write-ColorOutput "Restored from backup due to error" -Type "Warning"
        }
        
        exit 1
    }
    
    Write-Host ""
    
    # Step 5: Update VBScript launcher
    Write-ColorOutput "Step 5: Updating VBScript launcher..." -Type "Info"
    Update-VBScriptLauncher -ExePath $exePath
    
    Write-Host ""
}

# ===========================
# MODIFY CONFIGURATION
# ===========================
if ($mode -eq "modify" -or $mode -eq "both") {
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host "MODIFY CONFIGURATION" -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host ""
    
    # Backup current config
    Write-ColorOutput "Backing up current configuration..." -Type "Info"
    
    $configPath = Join-Path $INSTALL_DIR $CONFIG_NAME
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $configBackup = Join-Path $INSTALL_DIR "$CONFIG_NAME.backup_$timestamp"
    
    if (Test-Path $configPath) {
        Copy-Item -Path $configPath -Destination $configBackup -Force
        Write-ColorOutput "Configuration backed up" -Type "Success"
    }
    
    Write-Host ""
    
    # Get new configuration
    Write-ColorOutput "Enter new configuration values (press Enter to keep current value):" -Type "Info"
    Write-Host ""
    
    # Server IP
    Write-Host "Current Server URL: $($currentConfig.server_url)" -ForegroundColor Yellow
    do {
        $serverIP = Read-Host "Enter new server IP address (or press Enter to keep current)"
        
        if ([string]::IsNullOrWhiteSpace($serverIP)) {
            $serverURL = $currentConfig.server_url
            break
        }
        
        if (Test-IPAddress $serverIP) {
            $serverURL = "http://$serverIP:55000/api/v1"
            break
        } else {
            Write-ColorOutput "Invalid IP address format. Please try again." -Type "Error"
        }
    } while ($true)
    Write-ColorOutput "Server URL: $serverURL" -Type "Success"
    Write-Host ""
    
    # Agent ID
    Write-Host "Current Agent ID: $($currentConfig.agent_id)" -ForegroundColor Yellow
    $agentID = Read-Host "Enter new Agent ID (or press Enter to keep current)"
    if ([string]::IsNullOrWhiteSpace($agentID)) {
        $agentID = $currentConfig.agent_id
    }
    Write-ColorOutput "Agent ID: $agentID" -Type "Success"
    Write-Host ""
    
    # Agent Name
    Write-Host "Current Agent Name: $($currentConfig.agent_name)" -ForegroundColor Yellow
    $agentName = Read-Host "Enter new Agent Name (or press Enter to keep current)"
    if ([string]::IsNullOrWhiteSpace($agentName)) {
        $agentName = $currentConfig.agent_name
    }
    Write-ColorOutput "Agent Name: $agentName" -Type "Success"
    Write-Host ""
    
    # Heartbeat Interval
    Write-Host "Current Heartbeat Interval: $($currentConfig.heartbeat_interval) seconds" -ForegroundColor Yellow
    do {
        $heartbeatInput = Read-Host "Enter new heartbeat interval in seconds (or press Enter to keep current)"
        
        if ([string]::IsNullOrWhiteSpace($heartbeatInput)) {
            $heartbeatInterval = $currentConfig.heartbeat_interval
            break
        }
        
        if (Test-PositiveInteger $heartbeatInput) {
            $heartbeatInterval = [int]$heartbeatInput
            break
        } else {
            Write-ColorOutput "Please enter a valid positive number." -Type "Error"
        }
    } while ($true)
    Write-ColorOutput "Heartbeat Interval: $heartbeatInterval seconds" -Type "Success"
    Write-Host ""
    
    # Policy Sync Interval
    Write-Host "Current Policy Sync Interval: $($currentConfig.policy_sync_interval) seconds" -ForegroundColor Yellow
    do {
        $policySyncInput = Read-Host "Enter new policy sync interval in seconds (or press Enter to keep current)"
        
        if ([string]::IsNullOrWhiteSpace($policySyncInput)) {
            $policySyncInterval = $currentConfig.policy_sync_interval
            break
        }
        
        if (Test-PositiveInteger $policySyncInput) {
            $policySyncInterval = [int]$policySyncInput
            break
        } else {
            Write-ColorOutput "Please enter a valid positive number." -Type "Error"
        }
    } while ($true)
    Write-ColorOutput "Policy Sync Interval: $policySyncInterval seconds" -Type "Success"
    Write-Host ""
    
    # Summary
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host "New Configuration Summary:" -ForegroundColor Yellow
    Write-Host "  Server URL: $serverURL"
    Write-Host "  Agent ID: $agentID"
    Write-Host "  Agent Name: $agentName"
    Write-Host "  Heartbeat Interval: $heartbeatInterval seconds"
    Write-Host "  Policy Sync Interval: $policySyncInterval seconds"
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host ""
    
    $confirmConfig = Read-Host "Save new configuration? (Y/N)"
    
    if ($confirmConfig -eq "Y" -or $confirmConfig -eq "y") {
        try {
            $newConfig = @{
                server_url = $serverURL
                agent_id = $agentID
                agent_name = $agentName
                heartbeat_interval = $heartbeatInterval
                policy_sync_interval = $policySyncInterval
            }
            
            $configJson = $newConfig | ConvertTo-Json -Depth 10
            $configJson | Out-File -FilePath $configPath -Encoding UTF8 -Force
            
            Write-ColorOutput "Configuration updated successfully!" -Type "Success"
            
        } catch {
            Write-ColorOutput "Error saving configuration: $($_.Exception.Message)" -Type "Error"
            
            # Restore backup
            if (Test-Path $configBackup) {
                Copy-Item -Path $configBackup -Destination $configPath -Force
                Write-ColorOutput "Restored previous configuration" -Type "Warning"
            }
        }
    } else {
        Write-ColorOutput "Configuration changes discarded" -Type "Info"
    }
    
    Write-Host ""
}

# ===========================
# RESTART AGENT
# ===========================
Write-Host "============================================================" -ForegroundColor Green
Write-Host "FINALIZATION" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""

$restartAgent = Read-Host "Restart the agent now to apply changes? (Y/N)"

if ($restartAgent -eq "Y" -or $restartAgent -eq "y") {
    Write-ColorOutput "Restarting agent..." -Type "Info"
    
    try {
        # Stop if running
        $process = Get-Process -Name $PROCESS_NAME -ErrorAction SilentlyContinue
        if ($process) {
            Stop-Process -Name $PROCESS_NAME -Force
            Start-Sleep -Seconds 2
        }
        
        # Start with --background flag and admin privileges
        $exePath = Join-Path $INSTALL_DIR $EXE_NAME
        Start-Process -FilePath $exePath -ArgumentList "--background" -WorkingDirectory $INSTALL_DIR -Verb RunAs
        Start-Sleep -Seconds 3
        
        # Verify
        $newProcess = Get-Process -Name $PROCESS_NAME -ErrorAction SilentlyContinue
        if ($newProcess) {
            Write-ColorOutput "Agent restarted successfully! (PID: $($newProcess.Id))" -Type "Success"
            Write-ColorOutput "Running in background mode with administrator privileges" -Type "Success"
        } else {
            Write-ColorOutput "Agent started, but process not detected yet." -Type "Warning"
            Write-ColorOutput "It may take a few moments to initialize in background mode." -Type "Info"
        }
        
    } catch {
        Write-ColorOutput "Error restarting agent: $($_.Exception.Message)" -Type "Error"
        Write-ColorOutput "You may need to restart manually or reboot the system" -Type "Warning"
        Write-Host ""
        Write-Host "Manual start command:" -ForegroundColor Yellow
        Write-Host "  Start-Process '$exePath' -ArgumentList '--background' -Verb RunAs"
    }
} else {
    Write-ColorOutput "Changes will take effect on next agent start" -Type "Info"
    Write-ColorOutput "The agent will auto-start at next logon" -Type "Info"
}

Write-Host ""

# Final Summary
Write-Host "============================================================" -ForegroundColor Green
Write-Host "         Operation Completed Successfully!                 " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""

if ($mode -eq "upgrade" -or $mode -eq "both") {
    Write-Host "Executable Upgrade:" -ForegroundColor Yellow
    Write-Host "  New version installed: $version"
    Write-Host "  VBScript launcher updated"
    Write-Host "  Backup available: $backupPath"
    Write-Host ""
}

if ($mode -eq "modify" -or $mode -eq "both") {
    Write-Host "Configuration Changes:" -ForegroundColor Yellow
    Write-Host "  Configuration updated"
    Write-Host "  Backup available: $configBackup"
    Write-Host ""
}

Write-Host "Management Commands:" -ForegroundColor Yellow
Write-Host "  Start Agent:   Start-Process '$exePath' -ArgumentList '--background' -Verb RunAs"
Write-Host "  Stop Agent:    Stop-Process -Name '$PROCESS_NAME' -Force"
Write-Host "  Check Status:  Get-Process -Name '$PROCESS_NAME'"
Write-Host "  View Logs:     Check $INSTALL_DIR\cybersentinel_agent.log"
Write-Host ""
Write-Host "Note: The agent runs in background mode with administrator privileges" -ForegroundColor Cyan
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green

# Pause before exit
Write-Host ""
Read-Host "Press Enter to exit"