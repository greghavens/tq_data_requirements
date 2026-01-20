# ESXi SSH Data Collection Script

A PowerShell 7+ script that collects hardware and configuration data from multiple ESXi hosts via SSH for VMware Technical Qualification submissions.

## Requirements

- **PowerShell 7.0 or later**
- **VMware.PowerCLI module** - For connecting to ESXi hosts and managing SSH service
- **Posh-SSH module** - For SSH command execution

### Installing Required Modules

```powershell
Install-Module -Name VMware.PowerCLI -Scope CurrentUser
Install-Module -Name Posh-SSH -Scope CurrentUser
```

## Usage

```powershell
.\Collect-ESXiData.ps1 -HostFile <path> [options]
```

### Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-HostFile` | Yes | - | Path to a text file containing ESXi hostnames or IPs (one per line) |
| `-ThrottleLimit` | No | 10 | Number of concurrent hosts to process (1-50) |
| `-Timeout` | No | 30 | Seconds before giving up on an SSH command (5-300) |
| `-Retries` | No | 2 | Number of retry attempts for failed hosts (0-10) |
| `-OutputFile` | No | Auto-generated | Custom output CSV filename |
| `-PreserveSSHState` | No | $false | If set, restore SSH to its original state instead of always disabling |

### Examples

Basic usage with default settings:
```powershell
.\Collect-ESXiData.ps1 -HostFile .\hosts.txt
```

Process 5 hosts concurrently and preserve SSH state:
```powershell
.\Collect-ESXiData.ps1 -HostFile .\hosts.txt -ThrottleLimit 5 -PreserveSSHState
```

Specify custom output file and longer timeout:
```powershell
.\Collect-ESXiData.ps1 -HostFile .\hosts.txt -OutputFile inventory.csv -Timeout 60
```

## Host File Format

Create a text file with one ESXi hostname or IP address per line:

```
esxi-host-01.example.com
esxi-host-02.example.com
192.168.1.100
192.168.1.101
```

## Collected Data

The script collects the following information from each ESXi host:

### Static Commands
| Command | Description |
|---------|-------------|
| `vmware -vl` | VMware version information |
| `vsish -e get /hardware/bios/dmiInfo` | BIOS/DMI information |
| `vsish -e get /hardware/cpu/cpuModelName` | CPU model name |
| `vsish -e get /hardware/cpu/cpuInfo` | CPU details |
| `vsish -e get /memory/comprehensive` | Memory information |
| `esxcfg-scsidevs -a` | SCSI adapters |
| `esxcfg-scsidevs -A` | SCSI adapter details |
| `esxcfg-scsidevs -c` | SCSI device paths |
| `esxcli network nic list` | Network adapters |
| `lspci -v \|grep -i Ethernet -A2` | PCI Ethernet device details |

### Dynamic Driver Discovery

After running the static commands, the script automatically:

1. Parses `esxcfg-scsidevs -a` output to discover storage driver names
2. Parses `esxcli network nic list` output to discover network driver names
3. Runs `lspci -p |grep -i <driver>` for each unique driver found

All dynamic `lspci` outputs are combined into a single `lspci_output` column with labeled sections.

## Output Files

### CSV File
- Default filename: `esxi_inventory_YYYYMMDD_HHMMSS.csv`
- RFC 4180 compliant formatting (handles embedded commas, newlines, quotes)
- One row per host with all command outputs

### Log File
- Filename: `esxi_collector_YYYYMMDD_HHMMSS.log`
- Contains detailed execution logs including:
  - Connection status for each host
  - SSH command execution results
  - Any errors or warnings encountered

## How It Works

1. **Reads host list** from the specified file
2. **Prompts for credentials** (used for all hosts)
3. **Processes hosts in parallel** up to the throttle limit
4. For each host:
   - Connects via PowerCLI (`Connect-VIServer`)
   - Records current SSH service state
   - Enables SSH service if not running
   - Establishes SSH session via Posh-SSH
   - Executes all static commands
   - Discovers drivers and runs dynamic `lspci` commands
   - Closes SSH session
   - Restores SSH service state (based on `-PreserveSSHState`)
   - Disconnects from host
5. **Writes results** to CSV file

## Error Handling

- **Automatic retries**: Failed hosts are retried up to the configured retry count
- **Authentication failures**: Not retried (immediate failure)
- **Command failures**: If an SSH command fails, remaining commands for that host are skipped
- **Graceful cleanup**: SSH sessions and VI connections are always closed, even on failure

## SSH Service Behavior

By default, the script **disables SSH** on each host after collection (security best practice).

Use `-PreserveSSHState` to restore SSH to its original state:
- If SSH was running before the script, it stays running
- If SSH was stopped before the script, it gets stopped again

## Troubleshooting

### "VMware.PowerCLI module is not installed"
```powershell
Install-Module -Name VMware.PowerCLI -Scope CurrentUser
```

### "Posh-SSH module is not installed"
```powershell
Install-Module -Name Posh-SSH -Scope CurrentUser
```

### Connection timeouts
Increase the timeout value:
```powershell
.\Collect-ESXiData.ps1 -HostFile .\hosts.txt -Timeout 60
```

### Too many concurrent connections
Reduce the throttle limit:
```powershell
.\Collect-ESXiData.ps1 -HostFile .\hosts.txt -ThrottleLimit 5
```

## Version

Current version: **1.5.0**
