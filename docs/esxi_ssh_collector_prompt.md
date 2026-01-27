# ESXi SSH Data Collection Script

Create a PowerShell 7+ script that collects hardware and configuration data from ESXi hosts via SSH.

## Input

- A newline-separated text file containing ESXi hostnames or IP addresses (one per line)
- Credentials prompted at runtime (single username/password for all hosts)

## Workflow per host

1. Connect to the ESXi host via PowerCLI (`Connect-VIServer`)
2. Record whether SSH service is currently enabled
3. Enable the SSH service via `Get-VMHostService` / `Start-VMHostService`
4. Execute SSH commands using the Posh-SSH module (`New-SSHSession` / `Invoke-SSHCommand`):

### Static Commands
   - `vmware -vl`
   - `vsish -e get /hardware/bios/dmiInfo`
   - `vsish -e get /hardware/cpu/cpuModelName`
   - `vsish -e get /hardware/cpu/cpuInfo`
   - `vsish -e get /memory/comprehensive`
   - `esxcfg-scsidevs -A`
   - `esxcfg-scsidevs -c`
   - `esxcfg-scsidevs -l`
   - `esxcli storage core adapter list`
   - `esxcli network nic list`
   - `lspci -v |grep -i Ethernet -A2`

### Dynamic lspci Commands (Driver Discovery)

After running the static commands, parse the output to discover drivers and run additional `lspci` commands:

1. **Storage Drivers**: Parse `esxcli storage core adapter list` output to extract driver names from the "Driver" column, then run `lspci -p |grep -i <driver>` for each unique driver found.

2. **Network Drivers**: Parse `esxcli network nic list` output to extract driver names from the "Driver" column, then run `lspci -p |grep -i <driver>` for each unique driver found.

5. Disable SSH service (default behavior)
6. Disconnect from the host

## Output

- Single CSV file with proper RFC 4180 quoting (handles embedded commas, newlines, quotes)
- Filename auto-generated with timestamp (e.g., `esxi_inventory_20250113_143022.csv`)

### Column Structure

| Column | Description |
|--------|-------------|
| `Hostname` | The ESXi host name or IP |
| `vmware -vl` | VMware version info |
| `vsish -e get /hardware/bios/dmiInfo` | BIOS/DMI information |
| `vsish -e get /hardware/cpu/cpuModelName` | CPU model name |
| `vsish -e get /hardware/cpu/cpuInfo` | CPU details |
| `vsish -e get /memory/comprehensive` | Memory information |
| `esxcfg-scsidevs -A` | SCSI adapter details |
| `esxcfg-scsidevs -c` | SCSI device paths |
| `esxcfg-scsidevs -l` | SCSI device listing |
| `esxcli storage core adapter list` | Storage adapters with driver names |
| `esxcli network nic list` | Network adapters |
| `lspci -v \|grep -i Ethernet -A2` | PCI Ethernet device details |
| `lspci_output` | **Combined output** of all dynamic `lspci -p` commands for both storage and network drivers, with each command and its output clearly labeled |

## Logging

- Log all failures (connection errors, auth failures, command failures, timeouts) to both console and a separate timestamped log file
- Include hostname in all log entries

## Parallelization

- Use `ForEach-Object -Parallel` with a thread-safe approach for collecting results
- Default throttle limit: 10 concurrent hosts

## Command Line Parameters

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `-HostFile` | Yes | - | Path to the input file |
| `-ThrottleLimit` | No | 10 | Number of concurrent hosts |
| `-Timeout` | No | 30 | Seconds before giving up on SSH command |
| `-Retries` | No | 2 | Number of retry attempts for failed hosts |
| `-OutputFile` | No | auto-generated | Custom output CSV filename |
| `-PreserveSSHState` | No (switch) | $false | If set, restore SSH to its original state instead of always disabling |

## Error Handling

- Retry failed hosts up to the configured retry count before logging as failed
- Continue processing remaining hosts on any failure
- Log authentication failures with hostname
- Attempt all hosts regardless of maintenance mode or connection state
