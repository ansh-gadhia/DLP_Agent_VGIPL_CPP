# CyberSentinel DLP Agent - Installation Script
# This script installs the agent as a Windows service using NSSM

#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

# Configuration
$SERVICE_NAME = "CyberSentinelDLP"
$INSTALL_DIR = "C:\Program Files\CyberSentinel"
$EXE_URL = "https://github.com/ansh-gadhia/DLP_Agent_VGIPL_CPP/releases/download/1.0.0/cybersentinel_agent.exe"
$NSSM_URL = "https://nssm.cc/release/nssm-2.24.zip"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   CyberSentinel DLP Agent - Installation Script" -ForegroundColor Cyan
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

# Function to download file
function Download-File {
    param (
        [string]$Url,
        [string]$OutputPath
    )
    
    Write-Host "Downloading from: $Url" -ForegroundColor Yellow
    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $Url -OutFile $OutputPath -UseBasicParsing
        $ProgressPreference = 'Continue'
        Write-Host "[OK] Download completed: $OutputPath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "[ERROR] Download failed: $_" -ForegroundColor Red
        return $false
    }
}

# Function to create agent_config.json
function Create-ConfigFile {
    param (
        [string]$ServerUrl,
        [string]$AgentName,
        [string]$AgentId,
        [int]$HeartbeatInterval,
        [int]$PolicySyncInterval
    )
    
    $configPath = Join-Path $INSTALL_DIR "agent_config.json"
    
    $config = @{
        server_url = $ServerUrl
        agent_id = $AgentId
        agent_name = $AgentName
        heartbeat_interval = $HeartbeatInterval
        policy_sync_interval = $PolicySyncInterval
    }
    
    $configJson = $config | ConvertTo-Json -Depth 10
    $configJson | Out-File -FilePath $configPath -Encoding UTF8
    
    Write-Host "[OK] Configuration file created: $configPath" -ForegroundColor Green
    return $configPath
}

# Function to generate UUID
function New-Guid-String {
    return [System.Guid]::NewGuid().ToString()
}

# Step 1: Create installation directory
Write-Host "`n[Step 1/6] Creating installation directory..." -ForegroundColor Cyan
if (-not (Test-Path $INSTALL_DIR)) {
    New-Item -ItemType Directory -Path $INSTALL_DIR -Force | Out-Null
    Write-Host "[OK] Created: $INSTALL_DIR" -ForegroundColor Green
}
else {
    Write-Host "[OK] Directory already exists: $INSTALL_DIR" -ForegroundColor Green
}

# Step 2: Check if service already exists
Write-Host "`n[Step 2/6] Checking for existing service..." -ForegroundColor Cyan
$existingService = Get-Service -Name $SERVICE_NAME -ErrorAction SilentlyContinue
if ($existingService) {
    Write-Host "[WARNING] Service '$SERVICE_NAME' already exists!" -ForegroundColor Yellow
    $response = Read-Host "Do you want to reinstall? This will stop and remove the existing service. (Y/N)"
    if ($response -ne 'Y' -and $response -ne 'y') {
        Write-Host "Installation cancelled." -ForegroundColor Yellow
        exit 0
    }
    
    # Stop and remove existing service
    Write-Host "Stopping existing service..." -ForegroundColor Yellow
    Stop-Service -Name $SERVICE_NAME -Force -ErrorAction SilentlyContinue
    
    $nssmPath = Join-Path $INSTALL_DIR "nssm.exe"
    if (Test-Path $nssmPath) {
        & $nssmPath remove $SERVICE_NAME confirm
        Write-Host "[OK] Existing service removed" -ForegroundColor Green
    }
}

# Step 3: Download and install NSSM
Write-Host "`n[Step 3/6] Installing NSSM (Service Manager)..." -ForegroundColor Cyan
$nssmPath = Join-Path $INSTALL_DIR "nssm.exe"

if (-not (Test-Path $nssmPath)) {
    $nssmZip = Join-Path $env:TEMP "nssm.zip"
    $nssmExtract = Join-Path $env:TEMP "nssm"
    
    if (Download-File -Url $NSSM_URL -OutputPath $nssmZip) {
        Expand-Archive -Path $nssmZip -DestinationPath $nssmExtract -Force
        
        # Find nssm.exe in extracted folder (it's in win64 subfolder)
        $nssmExe = Get-ChildItem -Path $nssmExtract -Filter "nssm.exe" -Recurse | Where-Object { $_.Directory.Name -eq "win64" } | Select-Object -First 1
        
        if ($nssmExe) {
            Copy-Item -Path $nssmExe.FullName -Destination $nssmPath -Force
            Write-Host "[OK] NSSM installed: $nssmPath" -ForegroundColor Green
        }
        else {
            Write-Host "[ERROR] Could not find nssm.exe in downloaded archive" -ForegroundColor Red
            exit 1
        }
        
        # Cleanup
        Remove-Item -Path $nssmZip -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $nssmExtract -Recurse -Force -ErrorAction SilentlyContinue
    }
    else {
        Write-Host "[ERROR] Failed to download NSSM" -ForegroundColor Red
        exit 1
    }
}
else {
    Write-Host "[OK] NSSM already installed" -ForegroundColor Green
}

# Step 4: Download agent executable
Write-Host "`n[Step 4/6] Downloading CyberSentinel Agent..." -ForegroundColor Cyan
$agentExePath = Join-Path $INSTALL_DIR "cybersentinel_agent.exe"

if (Download-File -Url $EXE_URL -OutputPath $agentExePath) {
    Write-Host "[OK] Agent executable downloaded" -ForegroundColor Green
}
else {
    Write-Host "[ERROR] Failed to download agent executable" -ForegroundColor Red
    exit 1
}

# Step 5: Collect configuration from user
Write-Host "`n[Step 5/6] Agent Configuration" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Server IP/Hostname
$serverIp = ""
while ([string]::IsNullOrWhiteSpace($serverIp)) {
    $serverIp = Read-Host "Enter Server IP Address or Hostname (e.g., 192.168.1.100 or localhost)"
    if ([string]::IsNullOrWhiteSpace($serverIp)) {
        Write-Host "[ERROR] Server IP/Hostname is required!" -ForegroundColor Red
    }
}

# Remove any protocol if user accidentally added it
$serverIp = $serverIp -replace '^https?://', ''
# Remove any trailing slashes or paths
$serverIp = $serverIp -replace '/.*$', ''
# Remove any port if user accidentally added it
$serverIp = $serverIp -replace ':\d+$', ''

# Construct the full server URL
$serverUrl = "http://" + $serverIp + ":55000/api/v1"

Write-Host "[OK] Server IP: $serverIp" -ForegroundColor Green

# Agent Name
$agentName = ""
while ([string]::IsNullOrWhiteSpace($agentName)) {
    $defaultName = $env:COMPUTERNAME
    $input = Read-Host "Enter Agent Name (default: $defaultName)"
    if ([string]::IsNullOrWhiteSpace($input)) {
        $agentName = $defaultName
    }
    else {
        $agentName = $input
    }
}
Write-Host "[OK] Agent Name: $agentName" -ForegroundColor Green

# Agent ID
$agentId = ""
while ([string]::IsNullOrWhiteSpace($agentId)) {
    $defaultId = New-Guid-String
    Write-Host "Generated Agent ID: $defaultId" -ForegroundColor Yellow
    $input = Read-Host "Press Enter to use generated ID, or enter custom Agent ID"
    if ([string]::IsNullOrWhiteSpace($input)) {
        $agentId = $defaultId
    }
    else {
        $agentId = $input
    }
}
Write-Host "[OK] Agent ID: $agentId" -ForegroundColor Green

# Heartbeat Interval
$heartbeatInterval = 30
$input = Read-Host "Enter Heartbeat Interval in seconds (default: 30)"
if (-not [string]::IsNullOrWhiteSpace($input)) {
    try {
        $heartbeatInterval = [int]$input
    }
    catch {
        Write-Host "[WARNING] Invalid input, using default: 30" -ForegroundColor Yellow
        $heartbeatInterval = 30
    }
}
Write-Host "[OK] Heartbeat Interval: $heartbeatInterval seconds" -ForegroundColor Green

# Policy Sync Interval
$policySyncInterval = 60
$input = Read-Host "Enter Policy Sync Interval in seconds (default: 60)"
if (-not [string]::IsNullOrWhiteSpace($input)) {
    try {
        $policySyncInterval = [int]$input
    }
    catch {
        Write-Host "[WARNING] Invalid input, using default: 60" -ForegroundColor Yellow
        $policySyncInterval = 60
    }
}
Write-Host "[OK] Policy Sync Interval: $policySyncInterval seconds" -ForegroundColor Green

# Create configuration file
Write-Host "`nCreating configuration file..." -ForegroundColor Yellow
$configFile = Create-ConfigFile -ServerUrl $serverUrl -AgentName $agentName -AgentId $agentId -HeartbeatInterval $heartbeatInterval -PolicySyncInterval $policySyncInterval

# Step 6: Install and start service
Write-Host "`n[Step 6/6] Installing Windows Service..." -ForegroundColor Cyan

# Install service
& $nssmPath install $SERVICE_NAME $agentExePath

# Configure service
& $nssmPath set $SERVICE_NAME AppDirectory $INSTALL_DIR
& $nssmPath set $SERVICE_NAME DisplayName "CyberSentinel DLP Agent"
& $nssmPath set $SERVICE_NAME Description "Data Loss Prevention agent for monitoring and protecting sensitive data"
& $nssmPath set $SERVICE_NAME Start SERVICE_AUTO_START
& $nssmPath set $SERVICE_NAME AppStdout (Join-Path $INSTALL_DIR "service.log")
& $nssmPath set $SERVICE_NAME AppStderr (Join-Path $INSTALL_DIR "service_error.log")
& $nssmPath set $SERVICE_NAME AppRotateFiles 1
& $nssmPath set $SERVICE_NAME AppRotateBytes 10485760

Write-Host "[OK] Service installed: $SERVICE_NAME" -ForegroundColor Green

# Start service
Write-Host "`nStarting service..." -ForegroundColor Yellow
Start-Service -Name $SERVICE_NAME

# Wait a moment for service to start
Start-Sleep -Seconds 2

# Check service status
$service = Get-Service -Name $SERVICE_NAME
if ($service.Status -eq 'Running') {
    Write-Host "[OK] Service started successfully!" -ForegroundColor Green
}
else {
    Write-Host "[WARNING] Service status: $($service.Status)" -ForegroundColor Yellow
    Write-Host "Check logs at: $INSTALL_DIR\service.log" -ForegroundColor Yellow
}

# Summary
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "   Installation Complete!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Installation Directory: $INSTALL_DIR" -ForegroundColor White
Write-Host "Service Name: $SERVICE_NAME" -ForegroundColor White
Write-Host "Configuration File: $configFile" -ForegroundColor White
Write-Host "Service Status: $($service.Status)" -ForegroundColor White
Write-Host ""
Write-Host "Service Management Commands:" -ForegroundColor Cyan
Write-Host "  Start:   Start-Service -Name $SERVICE_NAME" -ForegroundColor White
Write-Host "  Stop:    Stop-Service -Name $SERVICE_NAME" -ForegroundColor White
Write-Host "  Restart: Restart-Service -Name $SERVICE_NAME" -ForegroundColor White
Write-Host "  Status:  Get-Service -Name $SERVICE_NAME" -ForegroundColor White
Write-Host ""
Write-Host "Logs Location: $INSTALL_DIR\service.log" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

pause
