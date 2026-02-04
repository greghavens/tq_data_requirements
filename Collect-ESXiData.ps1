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
    [switch]$PreserveSSHState
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptVersion = "1.5.2"

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

Examples:
  .\Collect-ESXiData.ps1 -HostFile .\hosts.txt
  .\Collect-ESXiData.ps1 -HostFile .\hosts.txt -ThrottleLimit 5 -PreserveSSHState

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

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $hostPrefix = if ($Hostname) { "[$Hostname] " } else { '' }
    $logEntry = "[$timestamp] [$Level] $hostPrefix$Message"

    # Console output
    switch ($Level) {
        'ERROR' { Write-Host $logEntry -ForegroundColor Red }
        'WARN'  { Write-Host $logEntry -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
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
        Write-Host "`nExisting output file detected: $OutputFile" -ForegroundColor Yellow
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

    # Suppress PowerCLI certificate warnings
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -Scope Session | Out-Null
    Set-PowerCLIConfiguration -ParticipateInCeip $false -Confirm:$false -Scope Session 2>$null | Out-Null

    # Create thread-safe synchronized hashtable for logging and progress
    $syncHash = [hashtable]::Synchronized(@{
        LogFile = $LogFile
        LogLock = [System.Threading.ReaderWriterLockSlim]::new()
        CompletedCount = 0
        TotalHosts = $hosts.Count
    })

    # Thread-safe collection for results
    $results = [System.Collections.Concurrent.ConcurrentBag[object]]::new()

    Write-Log -Message "Starting parallel processing of hosts" -Level 'INFO'

    # Process hosts in parallel
    $hosts | ForEach-Object -Parallel {
        $hostname = $_
        $cred = $using:credential
        $cmds = $using:SSHCommands
        $timeout = $using:Timeout
        $retries = $using:Retries
        $preserveSSH = $using:PreserveSSHState
        $sync = $using:syncHash
        $resultsBag = $using:results
        $storageDriverCmd = $using:StorageDriverCmd
        $networkDriverCmd = $using:NetworkDriverCmd

        # Import required modules in parallel runspace
        Import-Module VMware.PowerCLI -ErrorAction SilentlyContinue
        Import-Module Posh-SSH -ErrorAction SilentlyContinue

        # Define functions in parallel scope
        function Write-Log {
            param(
                [string]$Message,
                [string]$Level = 'INFO',
                [string]$Hostname = '',
                [hashtable]$SyncHash = $null
            )

            $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            $hostPrefix = if ($Hostname) { "[$Hostname] " } else { '' }
            $logEntry = "[$timestamp] [$Level] $hostPrefix$Message"

            switch ($Level) {
                'ERROR' { Write-Host $logEntry -ForegroundColor Red }
                'WARN'  { Write-Host $logEntry -ForegroundColor Yellow }
                'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
                default { Write-Host $logEntry }
            }

            if ($SyncHash) {
                $SyncHash.LogLock.EnterWriteLock()
                try {
                    Add-Content -Path $SyncHash.LogFile -Value $logEntry
                }
                finally {
                    $SyncHash.LogLock.ExitWriteLock()
                }
            }
        }

        # Note: SSH session is managed at the host level, not per-command

        # Process the host
        $result = [ordered]@{
            Hostname = $hostname
            Success = $false
        }

        foreach ($cmd in $cmds) {
            $result[$cmd] = ''
        }
        # Add column for combined lspci driver output
        $result['lspci_output'] = ''

        $viConnection = $null
        $sshSession = $null
        $sshWasRunning = $false
        $attempt = 0

        while ($attempt -le $retries) {
            $attempt++
            $sshSession = $null  # Reset for each attempt

            try {
                Write-Log -Message "Processing host (attempt $attempt/$($retries + 1))" -Hostname $hostname -SyncHash $sync

                Write-Log -Message "Connecting via PowerCLI" -Hostname $hostname -SyncHash $sync
                $viConnection = Connect-VIServer -Server $hostname -Credential $cred -ErrorAction Stop
                Write-Log -Message "Connected successfully to VI server" -Hostname $hostname -SyncHash $sync

                # Get the VMHost object (required for Get-VMHostService)
                Write-Log -Message "Getting VMHost object..." -Hostname $hostname -SyncHash $sync
                $vmHost = Get-VMHost -Server $viConnection
                Write-Log -Message "VMHost: $($vmHost.Name), ConnectionState: $($vmHost.ConnectionState), PowerState: $($vmHost.PowerState)" -Hostname $hostname -SyncHash $sync

                # Get all services and find SSH
                Write-Log -Message "Getting VMHost services..." -Hostname $hostname -SyncHash $sync
                $allServices = Get-VMHostService -VMHost $vmHost
                Write-Log -Message "Found $($allServices.Count) services" -Hostname $hostname -SyncHash $sync

                $sshService = $allServices | Where-Object { $_.Key -eq 'TSM-SSH' }
                if (-not $sshService) {
                    Write-Log -Message "ERROR: TSM-SSH service not found! Available services: $($allServices.Key -join ', ')" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                    throw "TSM-SSH service not found on host"
                }

                Write-Log -Message "SSH Service details - Key: $($sshService.Key), Label: $($sshService.Label), Running: $($sshService.Running), Policy: $($sshService.Policy)" -Hostname $hostname -SyncHash $sync

                $sshWasRunning = [bool]$sshService.Running
                Write-Log -Message "SSH service Running property value: '$($sshService.Running)' (type: $($sshService.Running.GetType().Name)), evaluated as: $sshWasRunning" -Hostname $hostname -SyncHash $sync

                if (-not $sshWasRunning) {
                    Write-Log -Message "Starting SSH service..." -Hostname $hostname -SyncHash $sync
                    $startResult = Start-VMHostService -HostService $sshService -Confirm:$false
                    Write-Log -Message "Start-VMHostService returned: Running=$($startResult.Running)" -Hostname $hostname -SyncHash $sync
                    Start-Sleep -Seconds 3

                    # Verify SSH actually started
                    $sshServiceAfter = Get-VMHostService -VMHost $vmHost | Where-Object { $_.Key -eq 'TSM-SSH' }
                    Write-Log -Message "After start - SSH Running: $($sshServiceAfter.Running)" -Hostname $hostname -SyncHash $sync
                    if (-not $sshServiceAfter.Running) {
                        Write-Log -Message "WARNING: SSH service may not have started properly!" -Level 'WARN' -Hostname $hostname -SyncHash $sync
                    }
                } else {
                    Write-Log -Message "SSH service is already running, skipping start" -Hostname $hostname -SyncHash $sync
                }

                # Create SSH session using Posh-SSH
                Write-Log -Message "Creating SSH session via Posh-SSH..." -Hostname $hostname -SyncHash $sync
                $sshSession = $null
                try {
                    $sshSession = New-SSHSession -ComputerName $hostname -Credential $cred -AcceptKey -ConnectionTimeout $timeout -ErrorAction Stop
                    Write-Log -Message "SSH session established (SessionId: $($sshSession.SessionId))" -Hostname $hostname -SyncHash $sync
                }
                catch {
                    Write-Log -Message "Failed to create SSH session: $_" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                    throw "SSH session failed: $_"
                }

                $sshFailed = $false
                foreach ($cmd in $cmds) {
                    Write-Log -Message "Executing: $cmd" -Hostname $hostname -SyncHash $sync
                    try {
                        $sshResult = Invoke-SSHCommand -SSHSession $sshSession -Command $cmd -TimeOut $timeout -ErrorAction Stop
                        Write-Log -Message "SSH exit code: $($sshResult.ExitStatus)" -Hostname $hostname -SyncHash $sync
                        if ($sshResult.ExitStatus -ne 0 -and $sshResult.Error) {
                            Write-Log -Message "SSH stderr: $($sshResult.Error)" -Level 'WARN' -Hostname $hostname -SyncHash $sync
                        }
                        $result[$cmd] = $sshResult.Output -join "`n"
                    }
                    catch {
                        Write-Log -Message "SSH command failed: $cmd - $_" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                        Write-Log -Message "Aborting remaining commands for this host" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                        $result[$cmd] = "ERROR: $_"
                        $sshFailed = $true
                        break  # Stop trying more commands on this host
                    }
                }

                # Dynamic lspci driver discovery (only if static commands succeeded)
                if (-not $sshFailed) {
                    $lspciOutputParts = [System.Collections.Generic.List[string]]::new()
                    $discoveredDrivers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

                    # Parse storage drivers from esxcli storage core adapter list output
                    # Format: HBA Name  Driver  Link State  UID  ...  (with header row)
                    $storageOutput = $result[$storageDriverCmd]
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
                    $networkOutput = $result[$networkDriverCmd]
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
                    foreach ($driver in $discoveredDrivers) {
                        $lspciCmd = "lspci -p |grep -i $driver"
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
                    $result['lspci_output'] = $lspciOutputParts -join "`n`n"
                }

                if ($sshFailed) {
                    Write-Log -Message "SSH commands failed - host marked as failed" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                } else {
                    $result.Success = $true
                    Write-Log -Message "Successfully collected data" -Level 'SUCCESS' -Hostname $hostname -SyncHash $sync
                }
                break  # Exit retry loop - we connected, no point retrying
            }
            catch {
                $errorMsg = $_.Exception.Message

                if ($errorMsg -match 'authentication|credential|password|login' -or $_.Exception.GetType().Name -match 'Auth') {
                    Write-Log -Message "Authentication failed: $errorMsg" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                    break  # Don't retry auth failures
                }

                if ($attempt -le $retries) {
                    Write-Log -Message "Attempt $attempt failed: $errorMsg. Retrying..." -Level 'WARN' -Hostname $hostname -SyncHash $sync
                    Start-Sleep -Seconds 5
                }
                else {
                    Write-Log -Message "All attempts failed: $errorMsg" -Level 'ERROR' -Hostname $hostname -SyncHash $sync
                }
            }
            finally {
                # Always close SSH session if open
                if ($sshSession) {
                    Write-Log -Message "Closing SSH session..." -Hostname $hostname -SyncHash $sync
                    Remove-SSHSession -SSHSession $sshSession -ErrorAction SilentlyContinue | Out-Null
                    $sshSession = $null
                }

                if ($viConnection) {
                    try {
                        $vmHostCleanup = Get-VMHost -Server $viConnection -ErrorAction SilentlyContinue
                        if ($vmHostCleanup) {
                            $sshServiceCleanup = Get-VMHostService -VMHost $vmHostCleanup | Where-Object { $_.Key -eq 'TSM-SSH' }

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
                        Disconnect-VIServer -Server $viConnection -Confirm:$false | Out-Null
                    }
                    catch {
                        Write-Log -Message "Failed to disconnect: $_" -Level 'WARN' -Hostname $hostname -SyncHash $sync
                    }
                }
            }
        }

        # Update progress counter (thread-safe)
        $completedCount = [System.Threading.Interlocked]::Increment([ref]$sync.CompletedCount)
        $percentComplete = [math]::Round(($completedCount / $sync.TotalHosts) * 100, 1)
        $remaining = $sync.TotalHosts - $completedCount
        Write-Log -Message "Progress: $completedCount/$($sync.TotalHosts) hosts completed ($percentComplete%) - $remaining remaining" -Level 'INFO' -Hostname $hostname -SyncHash $sync

        $resultsBag.Add([PSCustomObject]$result)

    } -ThrottleLimit $ThrottleLimit

    Write-Log -Message "Parallel processing complete" -Level 'INFO'

    # Convert results to array
    $newResults = $results.ToArray()

    # If resuming, merge with existing results
    if ($processedHosts.Count -gt 0 -and (Test-Path $OutputFile)) {
        Write-Log -Message "Merging new results with existing data" -Level 'INFO'
        try {
            # Load existing CSV data as objects
            $existingCsv = Import-Csv -Path $OutputFile
            $existingResults = @()

            # Convert CSV rows back to result objects
            foreach ($row in $existingCsv) {
                $resultObj = [ordered]@{
                    Hostname = $row.Hostname
                    Success = $true  # Assume existing entries were successful
                }
                # Add all command columns
                foreach ($cmd in $SSHCommands) {
                    $resultObj[$cmd] = $row.$cmd
                }
                $resultObj['lspci_output'] = $row.lspci_output
                $existingResults += [PSCustomObject]$resultObj
            }

            # Combine existing and new results
            $allResults = @($existingResults + $newResults) | Sort-Object -Property Hostname
            Write-Log -Message "Merged $($existingResults.Count) existing + $($newResults.Count) new results" -Level 'SUCCESS'
        }
        catch {
            Write-Log -Message "Failed to merge with existing CSV: $_" -Level 'WARN'
            $allResults = $newResults | Sort-Object -Property Hostname
        }
    }
    else {
        $allResults = $newResults | Sort-Object -Property Hostname
    }

    # Build CSV with RFC 4180 compliant formatting
    $csvHeaders = @('Hostname') + $SSHCommands + @('lspci_output')
    $csvLines = [System.Collections.Generic.List[string]]::new()

    # Add header row
    $headerRow = ($csvHeaders | ForEach-Object { ConvertTo-CsvField -Value $_ }) -join ','
    $csvLines.Add($headerRow)

    # Add data rows
    foreach ($result in $allResults) {
        $rowValues = @(ConvertTo-CsvField -Value $result.Hostname)
        foreach ($cmd in $SSHCommands) {
            $rowValues += ConvertTo-CsvField -Value $result.$cmd
        }
        $rowValues += ConvertTo-CsvField -Value $result.lspci_output
        $csvLines.Add($rowValues -join ',')
    }

    # Write CSV file
    $csvContent = $csvLines -join "`r`n"
    Set-Content -Path $OutputFile -Value $csvContent -NoNewline

    # Summary
    $successCount = ($allResults | Where-Object { $_.Success }).Count
    $failCount = $allResults.Count - $successCount

    Write-Host "`n" -NoNewline
    Write-Log -Message "=== Collection Complete ===" -Level 'INFO'
    Write-Log -Message "Total hosts: $($allResults.Count)" -Level 'INFO'
    Write-Log -Message "Successful: $successCount" -Level $(if ($successCount -gt 0) { 'SUCCESS' } else { 'INFO' })
    Write-Log -Message "Failed: $failCount" -Level $(if ($failCount -gt 0) { 'WARN' } else { 'INFO' })
    Write-Log -Message "Output saved to: $OutputFile" -Level 'INFO'
    Write-Log -Message "Log saved to: $LogFile" -Level 'INFO'

    # Cleanup
    $syncHash.LogLock.Dispose()
}
catch {
    Write-Log -Message "Fatal error: $_" -Level 'ERROR'
    Write-Log -Message $_.ScriptStackTrace -Level 'ERROR'
    exit 1
}
