# CyberSentinel DLP Agent - Update Script
# This script updates the agent executable and/or configuration

#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

# Configuration
$SERVICE_NAME = "CyberSentinelDLP"
$INSTALL_DIR = "C:\Program Files\CyberSentinel"
$EXE_URL = "https://github.com/ansh-gadhia/DLP_Agent_VGIPL_CPP/releases/download/1.0.0/cybersentinel_agent.exe"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   CyberSentinel DLP Agent - Update Script" -ForegroundColor Cyan
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
$service = Get-Service -Name $SERVICE_NAME -ErrorAction SilentlyContinue
if (-not $service) {
    Write-Host "[ERROR] Service '$SERVICE_NAME' not found!" -ForegroundColor Red
    Write-Host "Please install the agent first using Install-DLPAgent.ps1" -ForegroundColor Yellow
    pause
    exit 1
}

# Check if installation directory exists
if (-not (Test-Path $INSTALL_DIR)) {
    Write-Host "[ERROR] Installation directory not found: $INSTALL_DIR" -ForegroundColor Red
    pause
    exit 1
}

# Main menu
Write-Host "What would you like to update?" -ForegroundColor Cyan
Write-Host "1. Update Agent Executable (from GitHub releases)" -ForegroundColor White
Write-Host "2. Update Configuration (agent_config.json)" -ForegroundColor White
Write-Host "3. Update Both" -ForegroundColor White
Write-Host "4. View Current Configuration" -ForegroundColor White
Write-Host "5. Exit" -ForegroundColor White
Write-Host ""

$choice = Read-Host "Enter your choice (1-5)"

switch ($choice) {
    "1" {
        # Update executable only
        Write-Host "`nUpdating agent executable..." -ForegroundColor Yellow
        
        # Stop service
        Write-Host "Stopping service..." -ForegroundColor Yellow
        Stop-Service -Name $SERVICE_NAME -Force
        Start-Sleep -Seconds 2
        
        # Backup current executable
        $agentExePath = Join-Path $INSTALL_DIR "cybersentinel_agent.exe"
        $backupPath = Join-Path $INSTALL_DIR "cybersentinel_agent.exe.backup"
        
        if (Test-Path $agentExePath) {
            Copy-Item -Path $agentExePath -Destination $backupPath -Force
            Write-Host "[OK] Backup created: $backupPath" -ForegroundColor Green
        }
        
        # Download new executable
        try {
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $EXE_URL -OutFile $agentExePath -UseBasicParsing
            $ProgressPreference = 'Continue'
            Write-Host "[OK] New executable downloaded" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERROR] Download failed: $_" -ForegroundColor Red
            
            # Restore backup
            if (Test-Path $backupPath) {
                Copy-Item -Path $backupPath -Destination $agentExePath -Force
                Write-Host "[OK] Backup restored" -ForegroundColor Yellow
            }
            
            Start-Service -Name $SERVICE_NAME
            pause
            exit 1
        }
        
        # Start service
        Write-Host "Starting service..." -ForegroundColor Yellow
        Start-Service -Name $SERVICE_NAME
        Start-Sleep -Seconds 2
        
        $service = Get-Service -Name $SERVICE_NAME
        if ($service.Status -eq 'Running') {
            Write-Host "[OK] Service restarted successfully!" -ForegroundColor Green
            Write-Host "[OK] Agent executable updated!" -ForegroundColor Green
        }
        else {
            Write-Host "[WARNING] Service status: $($service.Status)" -ForegroundColor Yellow
        }
    }
    
    "2" {
        # Update configuration only
        Write-Host "`nUpdating configuration..." -ForegroundColor Yellow
        
        $configPath = Join-Path $INSTALL_DIR "agent_config.json"
        
        # Backup current config
        if (Test-Path $configPath) {
            $backupConfigPath = Join-Path $INSTALL_DIR "agent_config.json.backup"
            Copy-Item -Path $configPath -Destination $backupConfigPath -Force
            Write-Host "[OK] Configuration backup created" -ForegroundColor Green
            
            # Read current config
            try {
                $currentConfig = Get-Content -Path $configPath -Raw | ConvertFrom-Json
                Write-Host "`nCurrent Configuration:" -ForegroundColor Cyan
                Write-Host "  Server URL: $($currentConfig.server_url)" -ForegroundColor White
                Write-Host "  Agent Name: $($currentConfig.agent_name)" -ForegroundColor White
                Write-Host "  Agent ID: $($currentConfig.agent_id)" -ForegroundColor White
                Write-Host "  Heartbeat Interval: $($currentConfig.heartbeat_interval) seconds" -ForegroundColor White
                Write-Host "  Policy Sync Interval: $($currentConfig.policy_sync_interval) seconds" -ForegroundColor White
                Write-Host ""
            }
            catch {
                $currentConfig = $null
            }
        }
        
        # Collect new configuration
        Write-Host "Enter new configuration (press Enter to keep current value):" -ForegroundColor Cyan
        Write-Host ""
        
        # Server URL
        $currentServerUrl = if ($currentConfig) { $currentConfig.server_url } else { "" }
        $input = Read-Host "Server URL [$currentServerUrl]"
        $serverUrl = if ([string]::IsNullOrWhiteSpace($input)) { $currentServerUrl } else { $input }
        
        # Agent Name
        $currentAgentName = if ($currentConfig) { $currentConfig.agent_name } else { $env:COMPUTERNAME }
        $input = Read-Host "Agent Name [$currentAgentName]"
        $agentName = if ([string]::IsNullOrWhiteSpace($input)) { $currentAgentName } else { $input }
        
        # Agent ID
        $currentAgentId = if ($currentConfig) { $currentConfig.agent_id } else { [System.Guid]::NewGuid().ToString() }
        $input = Read-Host "Agent ID [$currentAgentId]"
        $agentId = if ([string]::IsNullOrWhiteSpace($input)) { $currentAgentId } else { $input }
        
        # Heartbeat Interval
        $currentHeartbeat = if ($currentConfig) { $currentConfig.heartbeat_interval } else { 30 }
        $input = Read-Host "Heartbeat Interval in seconds [$currentHeartbeat]"
        $heartbeatInterval = if ([string]::IsNullOrWhiteSpace($input)) { $currentHeartbeat } else { [int]$input }
        
        # Policy Sync Interval
        $currentPolicySync = if ($currentConfig) { $currentConfig.policy_sync_interval } else { 60 }
        $input = Read-Host "Policy Sync Interval in seconds [$currentPolicySync]"
        $policySyncInterval = if ([string]::IsNullOrWhiteSpace($input)) { $currentPolicySync } else { [int]$input }
        
        # Create new config
        $newConfig = @{
            server_url = $serverUrl
            agent_id = $agentId
            agent_name = $agentName
            heartbeat_interval = $heartbeatInterval
            policy_sync_interval = $policySyncInterval
        }
        
        $configJson = $newConfig | ConvertTo-Json -Depth 10
        $configJson | Out-File -FilePath $configPath -Encoding UTF8
        
        Write-Host "`n[OK] Configuration updated!" -ForegroundColor Green
        
        # Restart service to apply changes
        Write-Host "Restarting service to apply changes..." -ForegroundColor Yellow
        Restart-Service -Name $SERVICE_NAME
        Start-Sleep -Seconds 2
        
        $service = Get-Service -Name $SERVICE_NAME
        if ($service.Status -eq 'Running') {
            Write-Host "[OK] Service restarted successfully!" -ForegroundColor Green
        }
        else {
            Write-Host "[WARNING] Service status: $($service.Status)" -ForegroundColor Yellow
        }
    }
    
    "3" {
        # Update both
        Write-Host "`nUpdating agent executable and configuration..." -ForegroundColor Yellow
        
        # First update config (same as option 2)
        $configPath = Join-Path $INSTALL_DIR "agent_config.json"
        
        if (Test-Path $configPath) {
            $backupConfigPath = Join-Path $INSTALL_DIR "agent_config.json.backup"
            Copy-Item -Path $configPath -Destination $backupConfigPath -Force
            
            try {
                $currentConfig = Get-Content -Path $configPath -Raw | ConvertFrom-Json
                Write-Host "`nCurrent Configuration:" -ForegroundColor Cyan
                Write-Host "  Server URL: $($currentConfig.server_url)" -ForegroundColor White
                Write-Host "  Agent Name: $($currentConfig.agent_name)" -ForegroundColor White
                Write-Host "  Agent ID: $($currentConfig.agent_id)" -ForegroundColor White
                Write-Host "  Heartbeat Interval: $($currentConfig.heartbeat_interval) seconds" -ForegroundColor White
                Write-Host "  Policy Sync Interval: $($currentConfig.policy_sync_interval) seconds" -ForegroundColor White
                Write-Host ""
            }
            catch {
                $currentConfig = $null
            }
        }
        
        Write-Host "Enter new configuration (press Enter to keep current value):" -ForegroundColor Cyan
        Write-Host ""
        
        $currentServerUrl = if ($currentConfig) { $currentConfig.server_url } else { "" }
        $input = Read-Host "Server URL [$currentServerUrl]"
        $serverUrl = if ([string]::IsNullOrWhiteSpace($input)) { $currentServerUrl } else { $input }
        
        $currentAgentName = if ($currentConfig) { $currentConfig.agent_name } else { $env:COMPUTERNAME }
        $input = Read-Host "Agent Name [$currentAgentName]"
        $agentName = if ([string]::IsNullOrWhiteSpace($input)) { $currentAgentName } else { $input }
        
        $currentAgentId = if ($currentConfig) { $currentConfig.agent_id } else { [System.Guid]::NewGuid().ToString() }
        $input = Read-Host "Agent ID [$currentAgentId]"
        $agentId = if ([string]::IsNullOrWhiteSpace($input)) { $currentAgentId } else { $input }
        
        $currentHeartbeat = if ($currentConfig) { $currentConfig.heartbeat_interval } else { 30 }
        $input = Read-Host "Heartbeat Interval in seconds [$currentHeartbeat]"
        $heartbeatInterval = if ([string]::IsNullOrWhiteSpace($input)) { $currentHeartbeat } else { [int]$input }
        
        $currentPolicySync = if ($currentConfig) { $currentConfig.policy_sync_interval } else { 60 }
        $input = Read-Host "Policy Sync Interval in seconds [$currentPolicySync]"
        $policySyncInterval = if ([string]::IsNullOrWhiteSpace($input)) { $currentPolicySync } else { [int]$input }
        
        $newConfig = @{
            server_url = $serverUrl
            agent_id = $agentId
            agent_name = $agentName
            heartbeat_interval = $heartbeatInterval
            policy_sync_interval = $policySyncInterval
        }
        
        $configJson = $newConfig | ConvertTo-Json -Depth 10
        $configJson | Out-File -FilePath $configPath -Encoding UTF8
        
        Write-Host "`n[OK] Configuration updated!" -ForegroundColor Green
        
        # Stop service
        Write-Host "Stopping service..." -ForegroundColor Yellow
        Stop-Service -Name $SERVICE_NAME -Force
        Start-Sleep -Seconds 2
        
        # Update executable
        $agentExePath = Join-Path $INSTALL_DIR "cybersentinel_agent.exe"
        $backupPath = Join-Path $INSTALL_DIR "cybersentinel_agent.exe.backup"
        
        if (Test-Path $agentExePath) {
            Copy-Item -Path $agentExePath -Destination $backupPath -Force
        }
        
        try {
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $EXE_URL -OutFile $agentExePath -UseBasicParsing
            $ProgressPreference = 'Continue'
            Write-Host "[OK] New executable downloaded" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERROR] Download failed: $_" -ForegroundColor Red
            
            if (Test-Path $backupPath) {
                Copy-Item -Path $backupPath -Destination $agentExePath -Force
            }
            
            Start-Service -Name $SERVICE_NAME
            pause
            exit 1
        }
        
        # Start service
        Write-Host "Starting service..." -ForegroundColor Yellow
        Start-Service -Name $SERVICE_NAME
        Start-Sleep -Seconds 2
        
        $service = Get-Service -Name $SERVICE_NAME
        if ($service.Status -eq 'Running') {
            Write-Host "[OK] Service restarted successfully!" -ForegroundColor Green
            Write-Host "[OK] Agent fully updated!" -ForegroundColor Green
        }
        else {
            Write-Host "[WARNING] Service status: $($service.Status)" -ForegroundColor Yellow
        }
    }
    
    "4" {
        # View current configuration
        $configPath = Join-Path $INSTALL_DIR "agent_config.json"
        
        if (Test-Path $configPath) {
            try {
                $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
                
                Write-Host "`nCurrent Configuration:" -ForegroundColor Cyan
                Write-Host "============================================================" -ForegroundColor Cyan
                Write-Host "Server URL:           $($config.server_url)" -ForegroundColor White
                Write-Host "Agent Name:           $($config.agent_name)" -ForegroundColor White
                Write-Host "Agent ID:             $($config.agent_id)" -ForegroundColor White
                Write-Host "Heartbeat Interval:   $($config.heartbeat_interval) seconds" -ForegroundColor White
                Write-Host "Policy Sync Interval: $($config.policy_sync_interval) seconds" -ForegroundColor White
                Write-Host "============================================================" -ForegroundColor Cyan
                
                # Show service status
                $service = Get-Service -Name $SERVICE_NAME
                Write-Host "`nService Status: $($service.Status)" -ForegroundColor $(if ($service.Status -eq 'Running') { 'Green' } else { 'Yellow' })
            }
            catch {
                Write-Host "[ERROR] Error reading configuration file" -ForegroundColor Red
            }
        }
        else {
            Write-Host "[ERROR] Configuration file not found: $configPath" -ForegroundColor Red
        }
    }
    
    "5" {
        Write-Host "Exiting..." -ForegroundColor Yellow
        exit 0
    }
    
    default {
        Write-Host "[ERROR] Invalid choice!" -ForegroundColor Red
        pause
        exit 1
    }
}

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "Update Complete!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

pause