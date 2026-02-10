<#
.SYNOPSIS
    Collects hardware and configuration data from ESXi hosts via SSH.

.DESCRIPTION
    Connects to ESXi hosts via PowerCLI, enables SSH, runs diagnostic commands,
    and outputs results to a CSV file with proper RFC 4180 quoting.

.PARAMETER HostFile
    Path to the input file containing ESXi hostnames or IP addresses (one per line).

.PARAMETER ThrottleLimit
    Number of concurrent hosts to process. Default: 10

.PARAMETER Timeout
    Seconds before giving up on SSH command. Default: 30

.PARAMETER Retries
    Number of retry attempts for failed hosts. Default: 2

.PARAMETER OutputFile
    Custom output CSV filename. Default: auto-generated with timestamp.

.PARAMETER PreserveSSHState
    If set, restore SSH to its original state instead of always disabling.

.PARAMETER DebugOutput
    If set, enables detailed diagnostic output including full exception types,
    inner exceptions, stack traces, and variable state at point of failure.

.EXAMPLE
    .\Collect-ESXiData.ps1 -HostFile .\hosts.txt
    .\Collect-ESXiData.ps1 -HostFile .\hosts.txt -ThrottleLimit 5 -PreserveSSHState
#>

#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$HostFile,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 50)]
    [int]$ThrottleLimit = 10,

    [Parameter(Mandatory = $false)]
    [ValidateRange(5, 300)]
    [int]$Timeout = 30,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 10)]
    [int]$Retries = 2,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile,

    [Parameter(Mandatory = $false)]
    [switch]$PreserveSSHState,

    [Parameter(Mandatory = $false)]
    [switch]$DebugOutput
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptVersion = "1.6.0"

# Always print version at startup
Write-Host "`nESXi SSH Data Collection Script v$ScriptVersion" -ForegroundColor Cyan

# Validate required parameter and show usage if missing
if (-not $HostFile) {
    Write-Host @"

Usage: .\Collect-ESXiData.ps1 -HostFile <path> [options]

Required:
  -HostFile <path>       Path to file containing ESXi hostnames (one per line)

Options:
  -ThrottleLimit <n>     Concurrent hosts (default: 10)
  -Timeout <seconds>     SSH command timeout (default: 30)
  -Retries <n>           Retry attempts for failed hosts (default: 2)
  -OutputFile <path>     Custom output CSV filename
  -PreserveSSHState      Restore SSH to original state instead of disabling
  -DebugOutput           Enable detailed diagnostic output for troubleshooting

Examples:
  .\Collect-ESXiData.ps1 -HostFile .\hosts.txt
  .\Collect-ESXiData.ps1 -HostFile .\hosts.txt -ThrottleLimit 5 -PreserveSSHState
  .\Collect-ESXiData.ps1 -HostFile .\hosts.txt -DebugOutput

"@ -ForegroundColor Cyan
    exit 1
}

# Validate host file exists
if (-not (Test-Path $HostFile -PathType Leaf)) {
    Write-Host "ERROR: Host file not found: $HostFile" -ForegroundColor Red
    exit 1
}

# Check for VMware.PowerCLI module
if (-not (Get-Module -ListAvailable -Name VMware.PowerCLI)) {
    Write-Host "ERROR: VMware.PowerCLI module is not installed." -ForegroundColor Red
    Write-Host "Install it with: Install-Module -Name VMware.PowerCLI -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

# Check for Posh-SSH module
if (-not (Get-Module -ListAvailable -Name Posh-SSH)) {
    Write-Host "ERROR: Posh-SSH module is not installed." -ForegroundColor Red
    Write-Host "Install it with: Install-Module -Name Posh-SSH -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

# Static SSH commands to execute on each host
$SSHCommands = @(
    'vmware -vl'
    'vsish -e get /hardware/bios/dmiInfo'
    'vsish -e get /hardware/cpu/cpuModelName'
    'vsish -e get /hardware/cpu/cpuInfo'
    'vsish -e get /memory/comprehensive'
    'esxcfg-scsidevs -A'
    'esxcfg-scsidevs -c'
    'esxcfg-scsidevs -l'
    'esxcli storage core adapter list'
    'esxcli network nic list'
    'lspci -v |grep -i Ethernet -A2'
)

# Commands whose output is used for driver discovery
$StorageDriverCmd = 'esxcli storage core adapter list'
$NetworkDriverCmd = 'esxcli network nic list'

# Generate timestamp for filenames
$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

# Set output and log file paths
if (-not $OutputFile) {
    $OutputFile = "esxi_inventory_$Timestamp.csv"
}
$LogFile = "esxi_collector_$Timestamp.log"

# Thread-safe logging function using synchronized hashtable
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO',
        [string]$Hostname = '',
        [hashtable]$SyncHash = $null
    )

    # Skip DEBUG messages unless DebugOutput is enabled
    if ($Level -eq 'DEBUG' -and -not $script:DebugOutput) { return }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $hostPrefix = if ($Hostname) { "[$Hostname] " } else { '' }
    $logEntry = "[$timestamp] [$Level] $hostPrefix$Message"

    # Console output
    switch ($Level) {
        'ERROR' { Write-Host $logEntry -ForegroundColor Red }
        'WARN'  { Write-Host $logEntry -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
        'DEBUG' { Write-Host $logEntry -ForegroundColor Magenta }
        default { Write-Host $logEntry }
    }

    # File output (thread-safe)
    if ($SyncHash) {
        $SyncHash.LogLock.EnterWriteLock()
        try {
            Add-Content -Path $SyncHash.LogFile -Value $logEntry
        }
        finally {
            $SyncHash.LogLock.ExitWriteLock()
        }
    }
    else {
        Add-Content -Path $script:LogFile -Value $logEntry
    }
}

# RFC 4180 compliant CSV field escaping
function ConvertTo-CsvField {
    param([string]$Value)

    if ([string]::IsNullOrEmpty($Value)) {
        return '""'
    }

    # Check if escaping is needed (contains comma, quote, or newline)
    if ($Value -match '[,"\r\n]') {
        # Escape quotes by doubling them and wrap in quotes
        return '"' + ($Value -replace '"', '""') + '"'
    }
    return $Value
}

# Main execution
try {
    Write-Log -Message "=== ESXi Data Collection Started ===" -Level 'INFO'
    if ($DebugOutput) {
        Write-Host "[DEBUG] Debug output ENABLED - detailed exception info will be shown" -ForegroundColor Magenta
        Write-Log -Message "Debug output enabled" -Level 'INFO'
    }
    Write-Log -Message "Host file: $HostFile" -Level 'INFO'
    Write-Log -Message "Output file: $OutputFile" -Level 'INFO'
    Write-Log -Message "Log file: $LogFile" -Level 'INFO'
    Write-Log -Message "Throttle limit: $ThrottleLimit" -Level 'INFO'
    Write-Log -Message "Timeout: $Timeout seconds" -Level 'INFO'
    Write-Log -Message "Retries: $Retries" -Level 'INFO'
    Write-Log -Message "Preserve SSH state: $PreserveSSHState" -Level 'INFO'

    # Load hosts from file (force array to handle single-host files)
    $hosts = @(Get-Content -Path $HostFile | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim() })

    if ($hosts.Count -eq 0) {
        throw "No hosts found in $HostFile"
    }

    Write-Log -Message "Loaded $($hosts.Count) host(s) from file" -Level 'INFO'

    # Prompt for credentials
    Write-Host "`nEnter credentials for ESXi hosts:" -ForegroundColor Cyan
    $credential = Get-Credential -Message "Enter ESXi credentials (used for all hosts)"

    if (-not $credential) {
        throw "Credentials are required"
    }

    # Check for existing output file and offer resume option
    $processedHosts = @()
    if (Test-Path $OutputFile) {
        # Validate that the file is a CSV (not a log file)
        $fileExtension = [System.IO.Path]::GetExtension($OutputFile)
        if ($fileExtension -ne '.csv') {
            Write-Log -Message "Output file exists but is not a CSV file ($fileExtension). Cannot resume from this file." -Level 'WARN'
            Write-Log -Message "To resume, specify a CSV output file with -OutputFile parameter" -Level 'WARN'
        }
        else {
            Write-Host "`nExisting CSV output file detected: $OutputFile" -ForegroundColor Yellow
            $resume = Read-Host "Do you want to resume and skip already-processed hosts? (Y/N)"

            if ($resume -eq 'Y' -or $resume -eq 'y') {
                Write-Log -Message "Resume mode: Loading previously processed hosts from $OutputFile" -Level 'INFO'
                try {
                    # Parse CSV to get already-processed hostnames
                    $csvContent = Import-Csv -Path $OutputFile
                    $processedHosts = @($csvContent | Select-Object -ExpandProperty Hostname)
                    Write-Log -Message "Found $($processedHosts.Count) already-processed host(s)" -Level 'INFO'

                # Filter hosts list to only unprocessed hosts
                $originalCount = $hosts.Count
                $hosts = @($hosts | Where-Object { $_ -notin $processedHosts })

                if ($hosts.Count -eq 0) {
                    Write-Host "`nAll hosts have already been processed. Nothing to do." -ForegroundColor Green
                    Write-Log -Message "All hosts already processed - exiting" -Level 'INFO'
                    exit 0
                }

                Write-Log -Message "Resuming: $($hosts.Count) host(s) remaining (skipped $($originalCount - $hosts.Count))" -Level 'SUCCESS'
                }
                catch {
                    Write-Log -Message "Failed to parse existing CSV: $_" -Level 'WARN'
                    Write-Log -Message "Proceeding with full host list" -Level 'WARN'
                    $processedHosts = @()
                }
            }
            else {
                Write-Log -Message "Overwriting existing output file" -Level 'WARN'
            }
        }
    }

    # Suppress PowerCLI certificate warnings
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -Scope Session | Out-Null
    Set-PowerCLIConfiguration -ParticipateInCeip $false -Confirm:$false -Scope Session 2>$null | Out-Null

    # Initialize CSV file with header (append mode for resume, create mode for new)
    $csvHeaders = @('Hostname') + $SSHCommands + @('lspci_output')
    if ($processedHosts.Count -eq 0) {
        # New file - write header
        $headerRow = ($csvHeaders | ForEach-Object { ConvertTo-CsvField -Value $_ }) -join ','
        Set-Content -Path $OutputFile -Value $headerRow -NoNewline
        Add-Content -Path $OutputFile -Value "`r`n" -NoNewline
        Write-Log -Message "Created new CSV file with header" -Level 'INFO'
    }
    else {
        # Resuming - CSV already exists with header
        Write-Log -Message "Appending to existing CSV file" -Level 'INFO'
    }

    # Create thread-safe synchronized hashtable for logging and CSV writing
    $syncHash = [hashtable]::Synchronized(@{
        LogFile = $LogFile
        LogLock = [System.Threading.ReaderWriterLockSlim]::new()
        CsvFile = $OutputFile
        CsvLock = [System.Threading.ReaderWriterLockSlim]::new()
        CsvHeaders = $csvHeaders
        CounterLock = [System.Threading.ReaderWriterLockSlim]::new()
        ProcessedCount = 0
        TotalHosts = $hosts.Count
        UpdateInterval = 10
    })

    Write-Log -Message "Starting parallel processing of hosts" -Level 'INFO'

    # Process hosts in parallel
    $hosts | ForEach-Object -Parallel {
        $hostname = $_
        $debugOutput = $using:DebugOutput
        if ($debugOutput) { Write-Host "[DEBUG][$hostname] Setting hostname=$hostname" -ForegroundColor Magenta }
        if ($debugOutput) { Write-Host "[DEBUG][$hostname] Setting cred from using:credential" -ForegroundColor Magenta }
        $cred = $using:credential
        if ($debugOutput) { Write-Host "[DEBUG][$hostname] Setting cmds from using:SSHCommands" -ForegroundColor Magenta }
        $cmds = $using:SSHCommands
        if ($debugOutput) { Write-Host "[DEBUG][$hostname] Setting timeout from using:Timeout" -ForegroundColor Magenta }
        $timeout = $using:Timeout
        if ($debugOutput) { Write-Host "[DEBUG][$hostname] Setting retries from using:Retries" -ForegroundColor Magenta }
        $retries = $using:Retries
        if ($debugOutput) { Write-Host "[DEBUG][$hostname] Setting preserveSSH from using:PreserveSSHState" -ForegroundColor Magenta }
        $preserveSSH = $using:PreserveSSHState
        if ($debugOutput) { Write-Host "[DEBUG][$hostname] Setting sync from using:syncHash" -ForegroundColor Magenta }
        $sync = $using:syncHash
        if ($debugOutput) { Write-Host "[DEBUG][$hostname] Setting storageDriverCmd from using:StorageDriverCmd" -ForegroundColor Magenta }
        $storageDriverCmd = $using:StorageDriverCmd
        if ($debugOutput) { Write-Host "[DEBUG][$hostname] Setting networkDriverCmd from using:NetworkDriverCmd" -ForegroundColor Magenta }
        $networkDriverCmd = $using:NetworkDriverCmd
        if ($debugOutput) { Write-Host "[DEBUG][$hostname] All using: variables set successfully" -ForegroundColor Magenta }

        # Helper: write detailed exception info when -DebugOutput is enabled
        function Write-DebugException {
            param(
                [System.Management.Automation.ErrorRecord]$ErrorRecord,
                [string]$Context = '',
                [hashtable]$SyncHash = $null,
                [string]$Hostname = ''
            )
            $ex = $ErrorRecord.Exception
            $lines = @(
                "===== DEBUG EXCEPTION DETAIL ====="
                "Context: $Context"
                "ErrorRecord: $($ErrorRecord.ToString())"
                "Exception Type: $($ex.GetType().FullName)"
                "Exception Message: $($ex.Message)"
                "Exception Source: $($ex.Source)"
                "Target Object: $($ErrorRecord.TargetObject)"
                "Category: $($ErrorRecord.CategoryInfo)"
                "Fully Qualified Error ID: $($ErrorRecord.FullyQualifiedErrorId)"
                "Script Stack Trace: $($ErrorRecord.ScriptStackTrace)"
                "Exception Stack Trace: $($ex.StackTrace)"
            )
            # Walk inner exceptions
            $inner = $ex.InnerException
            $depth = 1
            while ($inner) {
                $lines += "--- Inner Exception (depth $depth) ---"
                $lines += "  Type: $($inner.GetType().FullName)"
                $lines += "  Message: $($inner.Message)"
                $lines += "  Stack Trace: $($inner.StackTrace)"
                $inner = $inner.InnerException
                $depth++
            }
            $lines += "===== END DEBUG EXCEPTION DETAIL ====="
            $debugMsg = $lines -join "`n"
            Write-Host $debugMsg -ForegroundColor Magenta
            if ($SyncHash) {
                try {
                    $SyncHash['LogLock'].EnterWriteLock()
                    try {
                        Add-Content -Path $SyncHash['LogFile'] -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [DEBUG] [$Hostname] $debugMsg"
                    }
                    finally {
                        $SyncHash['LogLock'].ExitWriteLock()
                    }
                } catch {
                    Write-Host "WARNING: Could not write debug info to log file: $_" -ForegroundColor Yellow
                }
            }
        }

        try {
        # Import required modules in parallel runspace
        if ($debugOutput) { Write-Host "[DEBUG][$hostname] Importing VMware.PowerCLI module..." -ForegroundColor Magenta }
        Import-Module VMware.PowerCLI -ErrorAction SilentlyContinue
        if ($debugOutput) { Write-Host "[DEBUG][$hostname] Importing Posh-SSH module..." -ForegroundColor Magenta }
        Import-Module Posh-SSH -ErrorAction SilentlyContinue
        if ($debugOutput) { Write-Host "[DEBUG][$hostname] Modules imported successfully" -ForegroundColor Magenta }

        if ($debugOutput) { Write-Host "[DEBUG][$hostname] Defining Write-Log function..." -ForegroundColor Magenta }
        # Define functions in parallel scope
        function Write-Log {
            param(
                [string]$Message,
                [string]$Level = 'INFO',
                [string]$Hostname = '',
                [hashtable]$SyncHash = $null
            )

            # Skip DEBUG messages unless debugOutput is enabled
            if ($Level -eq 'DEBUG' -and -not $debugOutput) { return }

            $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            $hostPrefix = if ($Hostname) { "[$Hostname] " } else { '' }
            $logEntry = "[$timestamp] [$Level] $hostPrefix$Message"

            switch ($Level) {
                'ERROR' { Write-Host $logEntry -ForegroundColor Red }
                'WARN'  { Write-Host $logEntry -ForegroundColor Yellow }
                'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
                'DEBUG' { Write-Host $logEntry -ForegroundColor Magenta }
                default { Write-Host $logEntry }
            }

            if ($SyncHash) {
                $SyncHash['LogLock'].EnterWriteLock()
                try {
                    Add-Content -Path $SyncHash['LogFile'] -Value $logEntry
                }
                finally {
                    $SyncHash['LogLock'].ExitWriteLock()
                }
            }
        }

        Write-Log -Message "Write-Log defined. Defining ConvertTo-CsvField..." -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
        # RFC 4180 compliant CSV field escaping
        function ConvertTo-CsvField {
            param([string]$Value)

            if ([string]::IsNullOrEmpty($Value)) {
                return '""'
            }

            # Check if escaping is needed (contains comma, quote, or newline)
            if ($Value -match '[,"\r\n]') {
                # Escape quotes by doubling them and wrap in quotes
                return '"' + ($Value -replace '"', '""') + '"'
            }
            return $Value
        }

        Write-Log -Message "ConvertTo-CsvField defined. Defining Write-CsvRow..." -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
        # Thread-safe CSV row writing
        function Write-CsvRow {
            param(
                [hashtable]$Result,
                [hashtable]$SyncHash,
                [array]$Commands
            )

            $SyncHash['CsvLock'].EnterWriteLock()
            try {
                $rowValues = @(ConvertTo-CsvField -Value $Result.Hostname)
                foreach ($cmd in $Commands) {
                    $rowValues += ConvertTo-CsvField -Value $Result[$cmd]
                }
                $rowValues += ConvertTo-CsvField -Value $Result['lspci_output']
                $csvRow = ($rowValues -join ',') + "`r`n"
                Add-Content -Path $SyncHash['CsvFile'] -Value $csvRow -NoNewline
            }
            finally {
                $SyncHash['CsvLock'].ExitWriteLock()
            }
        }

        Write-Log -Message "Write-CsvRow defined. All functions defined." -Level 'DEBUG' -Hostname $hostname -SyncHash $sync

        # Thread-safe progress counter using lock
        Write-Log -Message "About to access sync['CounterLock'] for EnterWriteLock" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
        $sync['CounterLock'].EnterWriteLock()
        Write-Log -Message "CounterLock acquired" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
        try {
            Write-Log -Message "Reading sync['ProcessedCount']" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
            $currentCount = $sync['ProcessedCount'] + 1
            Write-Log -Message "Setting sync['ProcessedCount'] = $currentCount" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
            $sync['ProcessedCount'] = $currentCount
        }
        finally {
            Write-Log -Message "Releasing CounterLock" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
            $sync['CounterLock'].ExitWriteLock()
        }
        Write-Log -Message "Reading sync['TotalHosts']" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
        $totalHosts = $sync['TotalHosts']
        Write-Log -Message "Reading sync['UpdateInterval']" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
        $interval = $sync['UpdateInterval']
        Write-Log -Message "Counter: $currentCount / $totalHosts, interval=$interval" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
        if ($currentCount % $interval -eq 0) {
            $percent = ($currentCount / $totalHosts) * 100
            Write-Progress -Id 1 -Activity "Collecting data..." -Status "[$currentCount of $totalHosts, $([math]::Round($percent, 2))%]" -PercentComplete $percent
        }

        # Note: SSH session is managed at the host level, not per-command

        # Process the host
        Write-Log -Message "Creating result ordered hashtable" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
        $result = [ordered]@{
            Hostname = $hostname
            Success = $false
        }

        Write-Log -Message "Initializing result keys for $($cmds.Count) commands" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
        foreach ($cmd in $cmds) {
            $result[$cmd] = ''
        }
        # Add column for combined lspci driver output
        $result['lspci_output'] = ''
        Write-Log -Message "Result hashtable initialized with $($result.Count) keys" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync

        Write-Log -Message "Setting viConnection=null, sshSession=null, sshWasRunning=false, attempt=0" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
        $viConnection = $null
        $sshSession = $null
        $sshWasRunning = $false
        $attempt = 0

        Write-Log -Message "Entering retry while loop (retries=$retries)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
        while ($attempt -le $retries) {
            $attempt++
            $sshSession = $null  # Reset for each attempt
            Write-Log -Message "Retry loop iteration: attempt=$attempt, retries=$retries" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync

            try {
                Write-Log -Message "Processing host (attempt $attempt/$($retries + 1))" -Hostname $hostname -SyncHash $sync

                Write-Log -Message "About to call Connect-VIServer for $hostname" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                Write-Log -Message "Connecting via PowerCLI" -Hostname $hostname -SyncHash $sync
                $viConnection = Connect-VIServer -Server $hostname -Credential $cred -ErrorAction Stop
                Write-Log -Message "Connect-VIServer returned: $($viConnection)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                Write-Log -Message "Connected successfully to VI server" -Hostname $hostname -SyncHash $sync

                # Get the VMHost object (required for Get-VMHostService)
                Write-Log -Message "About to call Get-VMHost" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                Write-Log -Message "Getting VMHost object..." -Hostname $hostname -SyncHash $sync
                $vmHost = Get-VMHost -Server $viConnection
                Write-Log -Message "Get-VMHost returned: Name=$($vmHost.Name), Type=$($vmHost.GetType().FullName)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                Write-Log -Message "VMHost: $($vmHost.Name), ConnectionState: $($vmHost.ConnectionState), PowerState: $($vmHost.PowerState)" -Hostname $hostname -SyncHash $sync

                # Get all services and find SSH
                Write-Log -Message "About to call Get-VMHostService" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                Write-Log -Message "Getting VMHost services..." -Hostname $hostname -SyncHash $sync
                $allServices = Get-VMHostService -VMHost $vmHost
                Write-Log -Message "Get-VMHostService returned $($allServices.Count) services" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                Write-Log -Message "Found $($allServices.Count) services" -Hostname $hostname -SyncHash $sync

                Write-Log -Message "Filtering for TSM-SSH service" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                $sshService = $allServices | Where-Object { $_.Key -eq 'TSM-SSH' }
                if (-not $sshService) {
                    Write-Log -Message "ERROR: TSM-SSH service not found! Available services: $($allServices.Key -join ', ')" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                    throw "TSM-SSH service not found on host"
                }

                Write-Log -Message "SSH Service details - Key: $($sshService.Key), Label: $($sshService.Label), Running: $($sshService.Running), Policy: $($sshService.Policy)" -Hostname $hostname -SyncHash $sync

                Write-Log -Message "Setting sshWasRunning from sshService.Running" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                $sshWasRunning = [bool]$sshService.Running
                Write-Log -Message "sshWasRunning=$sshWasRunning" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                Write-Log -Message "SSH service Running property value: '$($sshService.Running)' (type: $($sshService.Running.GetType().Name)), evaluated as: $sshWasRunning" -Hostname $hostname -SyncHash $sync

                if (-not $sshWasRunning) {
                    Write-Log -Message "About to call Start-VMHostService" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    Write-Log -Message "Starting SSH service..." -Hostname $hostname -SyncHash $sync
                    $startResult = Start-VMHostService -HostService $sshService -Confirm:$false
                    Write-Log -Message "Start-VMHostService returned: $($startResult)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    Write-Log -Message "Start-VMHostService returned: Running=$($startResult.Running)" -Hostname $hostname -SyncHash $sync
                    Start-Sleep -Seconds 3

                    # Verify SSH actually started
                    Write-Log -Message "Verifying SSH started - calling Get-VMHostService again" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    $sshServiceAfter = Get-VMHostService -VMHost $vmHost | Where-Object { $_.Key -eq 'TSM-SSH' }
                    Write-Log -Message "Post-start SSH Running=$($sshServiceAfter.Running)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    Write-Log -Message "After start - SSH Running: $($sshServiceAfter.Running)" -Hostname $hostname -SyncHash $sync
                    if (-not $sshServiceAfter.Running) {
                        Write-Log -Message "WARNING: SSH service may not have started properly!" -Level 'WARN' -Hostname $hostname -SyncHash $sync
                    }
                } else {
                    Write-Log -Message "SSH service is already running, skipping start" -Hostname $hostname -SyncHash $sync
                }

                # Create SSH session using Posh-SSH
                Write-Log -Message "About to call New-SSHSession" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                Write-Log -Message "Creating SSH session via Posh-SSH..." -Hostname $hostname -SyncHash $sync
                $sshSession = $null
                try {
                    $sshSession = New-SSHSession -ComputerName $hostname -Credential $cred -AcceptKey -ConnectionTimeout $timeout -ErrorAction Stop
                    Write-Log -Message "New-SSHSession returned: SessionId=$($sshSession.SessionId), Connected=$($sshSession.Connected)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    Write-Log -Message "SSH session established (SessionId: $($sshSession.SessionId))" -Hostname $hostname -SyncHash $sync
                }
                catch {
                    Write-Log -Message "New-SSHSession FAILED: $($_.Exception.GetType().FullName): $($_.Exception.Message)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    Write-Log -Message "Failed to create SSH session: $_" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                    throw "SSH session failed: $_"
                }

                $sshFailed = $false
                Write-Log -Message "Starting SSH command loop ($($cmds.Count) commands)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                foreach ($cmd in $cmds) {
                    Write-Log -Message "About to call Invoke-SSHCommand: $cmd" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    Write-Log -Message "Executing: $cmd" -Hostname $hostname -SyncHash $sync
                    try {
                        $sshResult = Invoke-SSHCommand -SSHSession $sshSession -Command $cmd -TimeOut $timeout -ErrorAction Stop
                        Write-Log -Message "Invoke-SSHCommand returned: ExitStatus=$($sshResult.ExitStatus), OutputLines=$($sshResult.Output.Count)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                        Write-Log -Message "SSH exit code: $($sshResult.ExitStatus)" -Hostname $hostname -SyncHash $sync
                        if ($sshResult.ExitStatus -ne 0 -and $sshResult.Error) {
                            Write-Log -Message "SSH stderr: $($sshResult.Error)" -Level 'WARN' -Hostname $hostname -SyncHash $sync
                        }
                        Write-Log -Message "Setting result['$cmd'] from SSH output" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                        $result[$cmd] = $sshResult.Output -join "`n"
                        Write-Log -Message "result['$cmd'] set (length=$($result[$cmd].Length))" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    }
                    catch {
                        Write-Log -Message "SSH command FAILED: $cmd - $($_.Exception.GetType().FullName): $($_.Exception.Message)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                        Write-Log -Message "SSH command failed: $cmd - $_" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                        Write-Log -Message "Aborting remaining commands for this host" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                        $result[$cmd] = "ERROR: $_"
                        $sshFailed = $true
                        break  # Stop trying more commands on this host
                    }
                }

                # Dynamic lspci driver discovery (only if static commands succeeded)
                Write-Log -Message "sshFailed=$sshFailed - checking if lspci discovery should run" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                if (-not $sshFailed) {
                    Write-Log -Message "Creating lspciOutputParts List and discoveredDrivers HashSet" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    $lspciOutputParts = [System.Collections.Generic.List[string]]::new()
                    $discoveredDrivers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

                    # Parse storage drivers from esxcli storage core adapter list output
                    # Format: HBA Name  Driver  Link State  UID  ...  (with header row)
                    Write-Log -Message "Reading storageOutput from result['$storageDriverCmd']" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    $storageOutput = $result[$storageDriverCmd]
                    Write-Log -Message "storageOutput length=$($storageOutput.Length), isNull=$([string]::IsNullOrEmpty($storageOutput))" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    if ($storageOutput -and $storageOutput -notmatch '^ERROR:') {
                        $storageLines = $storageOutput -split "`n" | Where-Object { $_ -match '\S' -and $_ -notmatch '^\s*$' }
                        $headerFound = $false
                        $driverColIndex = -1
                        foreach ($line in $storageLines) {
                            if (-not $headerFound) {
                                # Look for header line containing "Driver"
                                if ($line -match 'Driver') {
                                    $headerFound = $true
                                    # Find column index for Driver (split by two or more spaces)
                                    $headerParts = $line -split '\s{2,}'
                                    for ($i = 0; $i -lt $headerParts.Count; $i++) {
                                        if ($headerParts[$i] -match '^Driver$') {
                                            $driverColIndex = $i
                                            break
                                        }
                                    }
                                }
                                continue
                            }
                            # Skip separator lines
                            if ($line -match '^[-\s]+$') { continue }
                            # Parse data line
                            $fields = $line -split '\s{2,}' | Where-Object { $_ }
                            if ($fields.Count -gt $driverColIndex -and $driverColIndex -ge 0) {
                                $driver = $fields[$driverColIndex].Trim()
                                if ($driver -and $driver -notmatch '^-+$') {
                                    [void]$discoveredDrivers.Add($driver)
                                }
                            }
                        }
                        Write-Log -Message "Discovered storage drivers: $($discoveredDrivers -join ', ')" -Hostname $hostname -SyncHash $sync
                    }

                    # Parse network drivers from esxcli network nic list output
                    # Format: Name    PCI Device    Driver    Admin Status    Link Status    Speed    Duplex    MAC Address    MTU    Description
                    Write-Log -Message "Reading networkOutput from result['$networkDriverCmd']" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    $networkOutput = $result[$networkDriverCmd]
                    Write-Log -Message "networkOutput length=$($networkOutput.Length), isNull=$([string]::IsNullOrEmpty($networkOutput))" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    if ($networkOutput -and $networkOutput -notmatch '^ERROR:') {
                        $networkLines = $networkOutput -split "`n" | Where-Object { $_ -match '\S' -and $_ -notmatch '^\s*$' }
                        $headerFound = $false
                        $driverColIndex = -1
                        foreach ($line in $networkLines) {
                            if (-not $headerFound) {
                                # Look for header line containing "Driver"
                                if ($line -match 'Driver') {
                                    $headerFound = $true
                                    # Find column index for Driver (split by two or more spaces)
                                    $headerParts = $line -split '\s{2,}'
                                    for ($i = 0; $i -lt $headerParts.Count; $i++) {
                                        if ($headerParts[$i] -match 'Driver') {
                                            $driverColIndex = $i
                                            break
                                        }
                                    }
                                }
                                continue
                            }
                            # Parse data line
                            $fields = $line -split '\s{2,}' | Where-Object { $_ }
                            if ($fields.Count -gt $driverColIndex -and $driverColIndex -ge 0) {
                                $driver = $fields[$driverColIndex].Trim()
                                if ($driver -and $driver -notmatch '^-+$') {
                                    [void]$discoveredDrivers.Add($driver)
                                }
                            }
                        }
                        Write-Log -Message "Total unique drivers (storage + network): $($discoveredDrivers -join ', ')" -Hostname $hostname -SyncHash $sync
                    }

                    # Run lspci -p for each discovered driver
                    Write-Log -Message "Running lspci for $($discoveredDrivers.Count) discovered drivers" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    foreach ($driver in $discoveredDrivers) {
                        $lspciCmd = "lspci -p |grep -i $driver"
                        Write-Log -Message "About to call Invoke-SSHCommand for lspci driver=$driver" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                        Write-Log -Message "Executing dynamic: $lspciCmd" -Hostname $hostname -SyncHash $sync
                        try {
                            $lspciResult = Invoke-SSHCommand -SSHSession $sshSession -Command $lspciCmd -TimeOut $timeout -ErrorAction Stop
                            $output = $lspciResult.Output -join "`n"
                            # Add labeled output to collection
                            $lspciOutputParts.Add("=== $lspciCmd ===`n$output")
                            Write-Log -Message "lspci for $driver completed (exit: $($lspciResult.ExitStatus))" -Hostname $hostname -SyncHash $sync
                        }
                        catch {
                            Write-Log -Message "lspci command failed for driver $driver`: $_" -Level 'WARN' -Hostname $hostname -SyncHash $sync
                            $lspciOutputParts.Add("=== $lspciCmd ===`nERROR: $_")
                        }
                    }

                    # Combine all lspci outputs into single column
                    Write-Log -Message "Combining $($lspciOutputParts.Count) lspci outputs into result['lspci_output']" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    $result['lspci_output'] = $lspciOutputParts -join "`n`n"
                    Write-Log -Message "lspci_output set (length=$($result['lspci_output'].Length))" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                }

                Write-Log -Message "Checking sshFailed=$sshFailed for final status" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                if ($sshFailed) {
                    Write-Log -Message "SSH commands failed - host marked as failed" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                } else {
                    Write-Log -Message "Setting result.Success = true" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    $result.Success = $true
                    Write-Log -Message "Successfully collected data" -Level 'SUCCESS' -Hostname $hostname -SyncHash $sync
                }
                Write-Log -Message "Breaking out of retry loop (success path)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                break  # Exit retry loop - we connected, no point retrying
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log -Message "CATCH in retry loop: $($_.Exception.GetType().FullName): $errorMsg" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                Write-Log -Message "ScriptStackTrace: $($_.ScriptStackTrace)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                Write-Log -Message "Exception StackTrace: $($_.Exception.StackTrace)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                $innerEx = $_.Exception.InnerException
                $d = 1
                while ($innerEx) {
                    Write-Log -Message "InnerException($d): $($innerEx.GetType().FullName): $($innerEx.Message)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    $innerEx = $innerEx.InnerException
                    $d++
                }

                if ($errorMsg -match 'authentication|credential|password|login' -or $_.Exception.GetType().Name -match 'Auth') {
                    Write-Log -Message "Auth failure detected, breaking" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    Write-Log -Message "Authentication failed: $errorMsg" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                    break  # Don't retry auth failures
                }

                if ($attempt -le $retries) {
                    Write-Log -Message "Will retry (attempt=$attempt, retries=$retries)" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    Write-Log -Message "Attempt $attempt failed: $errorMsg. Retrying..." -Level 'WARN' -Hostname $hostname -SyncHash $sync
                    Start-Sleep -Seconds 5
                }
                else {
                    Write-Log -Message "No more retries, giving up" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    Write-Log -Message "All attempts failed: $errorMsg" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                }
            }
            finally {
                Write-Log -Message "Entering finally block" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                # Always close SSH session if open
                if ($sshSession) {
                    Write-Log -Message "Closing SSH session (SessionId=$($sshSession.SessionId))" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    Write-Log -Message "Closing SSH session..." -Hostname $hostname -SyncHash $sync
                    Remove-SSHSession -SSHSession $sshSession -ErrorAction SilentlyContinue | Out-Null
                    $sshSession = $null
                }

                if ($viConnection) {
                    Write-Log -Message "viConnection exists, cleaning up" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    try {
                        Write-Log -Message "Calling Get-VMHost for cleanup" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                        $vmHostCleanup = Get-VMHost -Server $viConnection -ErrorAction SilentlyContinue
                        if ($vmHostCleanup) {
                            Write-Log -Message "Getting SSH service for cleanup" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                            $sshServiceCleanup = Get-VMHostService -VMHost $vmHostCleanup | Where-Object { $_.Key -eq 'TSM-SSH' }
                            Write-Log -Message "sshServiceCleanup found=$($null -ne $sshServiceCleanup), preserveSSH=$preserveSSH, sshWasRunning=$sshWasRunning" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync

                            if ($preserveSSH -and -not $sshWasRunning) {
                                Write-Log -Message "Restoring SSH to original state (stopping)" -Hostname $hostname -SyncHash $sync
                                Stop-VMHostService -HostService $sshServiceCleanup -Confirm:$false | Out-Null
                            }
                            elseif (-not $preserveSSH) {
                                Write-Log -Message "Stopping SSH service (default behavior)" -Hostname $hostname -SyncHash $sync
                                Stop-VMHostService -HostService $sshServiceCleanup -Confirm:$false | Out-Null
                            }
                        }
                    }
                    catch {
                        Write-Log -Message "Failed to manage SSH service: $_" -Level 'WARN' -Hostname $hostname -SyncHash $sync
                    }

                    try {
                        Write-Log -Message "Calling Disconnect-VIServer" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                        Disconnect-VIServer -Server $viConnection -Confirm:$false | Out-Null
                        Write-Log -Message "Disconnect-VIServer completed" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                    }
                    catch {
                        Write-Log -Message "Disconnect-VIServer FAILED: $($_.Exception.GetType().FullName): $_" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
                        Write-Log -Message "Failed to disconnect: $_" -Level 'WARN' -Hostname $hostname -SyncHash $sync
                    }
                }
            }
        }

        # Write results to CSV immediately (thread-safe)
        Write-Log -Message "Writing results to CSV" -Level 'DEBUG' -Hostname $hostname -SyncHash $sync
        Write-CsvRow -Result $result -SyncHash $sync -Commands $cmds

        # Simple progress message - no counter needed, CSV is the source of truth
        Write-Log -Message "Host completed and written to CSV" -Level 'SUCCESS' -Hostname $hostname -SyncHash $sync

        } # end outer try
        catch {
            # This catches ANY unhandled error in the entire parallel scriptblock
            Write-Host "[$hostname] UNHANDLED PARALLEL ERROR: $_" -ForegroundColor Red
            Write-Host "[$hostname] ScriptStackTrace: $($_.ScriptStackTrace)" -ForegroundColor Red
            if ($debugOutput) {
                Write-DebugException -ErrorRecord $_ -Context "Unhandled error in parallel scriptblock for $hostname" -SyncHash $sync -Hostname $hostname
            }
            # Still try to log it
            try {
                $sync['LogLock'].EnterWriteLock()
                try {
                    Add-Content -Path $sync['LogFile'] -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] [$hostname] UNHANDLED PARALLEL ERROR: $_"
                    Add-Content -Path $sync['LogFile'] -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] [$hostname] ScriptStackTrace: $($_.ScriptStackTrace)"
                    if ($debugOutput) {
                        Add-Content -Path $sync['LogFile'] -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] [$hostname] Exception Type: $($_.Exception.GetType().FullName)"
                        Add-Content -Path $sync['LogFile'] -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] [$hostname] Exception StackTrace: $($_.Exception.StackTrace)"
                        $inner = $_.Exception.InnerException
                        $depth = 1
                        while ($inner) {
                            Add-Content -Path $sync['LogFile'] -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] [$hostname] Inner Exception ($depth): $($inner.GetType().FullName): $($inner.Message)"
                            $inner = $inner.InnerException
                            $depth++
                        }
                    }
                }
                finally {
                    $sync['LogLock'].ExitWriteLock()
                }
            } catch {
                Write-Host "[$hostname] WARNING: Could not write error to log file: $_" -ForegroundColor Yellow
            }
        }

    } -ThrottleLimit $ThrottleLimit

    Write-Log -Message "Parallel processing complete" -Level 'INFO'

    # Read CSV to generate summary (CSV was written incrementally during processing)
    $csvData = Import-Csv -Path $OutputFile -ErrorAction SilentlyContinue
    $totalInCsv = if ($csvData) { $csvData.Count } else { 0 }

    # Calculate based on what was expected vs what's in CSV
    $expectedTotal = $processedHosts.Count + $hosts.Count
    $successCount = $totalInCsv
    $failCount = $expectedTotal - $successCount

    Write-Host "`n" -NoNewline
    Write-Log -Message "=== Collection Complete ===" -Level 'INFO'
    Write-Log -Message "Total hosts processed: $successCount of $expectedTotal" -Level 'INFO'
    Write-Log -Message "Successful: $successCount" -Level $(if ($successCount -gt 0) { 'SUCCESS' } else { 'INFO' })
    Write-Log -Message "Failed: $failCount" -Level $(if ($failCount -gt 0) { 'WARN' } else { 'INFO' })
    Write-Log -Message "Output saved to: $OutputFile" -Level 'INFO'
    Write-Log -Message "Log saved to: $LogFile" -Level 'INFO'

    # Cleanup
    $syncHash['LogLock'].Dispose()
    $syncHash['CsvLock'].Dispose()
    $syncHash['CounterLock'].Dispose()
}
catch {
    Write-Log -Message "Fatal error: $_" -Level 'ERROR'
    Write-Log -Message "ScriptStackTrace: $($_.ScriptStackTrace)" -Level 'ERROR'
    if ($DebugOutput) {
        Write-Host "`n===== FATAL ERROR DEBUG DETAIL =====" -ForegroundColor Magenta
        Write-Host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Magenta
        Write-Host "Exception Message: $($_.Exception.Message)" -ForegroundColor Magenta
        Write-Host "Exception Source: $($_.Exception.Source)" -ForegroundColor Magenta
        Write-Host "Target Object: $($_.TargetObject)" -ForegroundColor Magenta
        Write-Host "Category Info: $($_.CategoryInfo)" -ForegroundColor Magenta
        Write-Host "Fully Qualified Error ID: $($_.FullyQualifiedErrorId)" -ForegroundColor Magenta
        Write-Host "Script Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Magenta
        Write-Host "Exception Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Magenta
        $inner = $_.Exception.InnerException
        $depth = 1
        while ($inner) {
            Write-Host "--- Inner Exception (depth $depth) ---" -ForegroundColor Magenta
            Write-Host "  Type: $($inner.GetType().FullName)" -ForegroundColor Magenta
            Write-Host "  Message: $($inner.Message)" -ForegroundColor Magenta
            Write-Host "  Stack Trace: $($inner.StackTrace)" -ForegroundColor Magenta
            $inner = $inner.InnerException
            $depth++
        }
        Write-Host "===== END FATAL ERROR DEBUG DETAIL =====" -ForegroundColor Magenta
    }
    exit 1
}
