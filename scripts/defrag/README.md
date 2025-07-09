# Exchange 2016 Database Defragmentation Script

**Languages / Языки:**
- [🇺🇸 English](README.md) ← (Current)
- [🇷🇺 Русский](README.ru.md)

---

Automated PowerShell script for safe defragmentation of Exchange Server 2016 databases.

## Description

This script performs complete defragmentation of Exchange Server 2016 databases with automatic safety checks, logging, and service recovery. It is designed to free up space after deleting or moving mailboxes.

## ⚠️ Important Warnings

- **MANDATORY** create a database backup before running
- Defragmentation may take several hours
- Database will be unavailable during execution
- Requires free disk space (≥110% of database size)
- Run during non-business hours

## Installation

1. Copy the script to the folder:
   ```
   C:\Scripts\Defrag\ExchangeDefrag.ps1
   ```

2. Configure parameters at the beginning of the script:
   ```powershell
   # Database name for defragmentation
   $DatabaseName = "Name-DB"

   # Log file path
   $LogPath = "C:\Scripts\Defrag\Logs\DefragLog.txt"

   # Force execution without confirmations (True/False)
   $Force = $True
   ```

## Requirements

- Exchange Server 2016
- PowerShell 5.0 or higher
- Exchange administrator rights
- Local administrator rights on server
- Free disk space ≥110% of database size

## Functionality

### Automatic Checks

- ✅ Check database existence
- ✅ Check free disk space
- ✅ Check active connections
- ✅ Check database integrity after defragmentation

### Safety

- ✅ Automatic stop and start of Exchange services
- ✅ Database dismounting and mounting
- ✅ Recovery in case of errors
- ✅ Detailed logging of all operations

### Monitoring

- ✅ Detailed logging with timestamps
- ✅ Colored console output
- ✅ Calculation of freed space
- ✅ Execution time measurement

## Usage

### Basic Execution

```powershell
# Run from Exchange Management Shell
.\ExchangeDefrag.ps1
```

### Execution with Settings

```powershell
# Change parameters in the script before running
$DatabaseName = "MyDatabase"
$LogPath = "D:\Logs\DefragLog.txt"
$Force = $False  # Enable interactive confirmations
```

### Scheduling via Task Scheduler

```powershell
# Create task for non-business hours execution
$action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument '-ExecutionPolicy Bypass -File "C:\Scripts\Defrag\ExchangeDefrag.ps1"'
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 2:00AM
$settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -WakeToRun
Register-ScheduledTask -TaskName "Exchange DB Defrag" -Action $action -Trigger $trigger -Settings $settings -User "DOMAIN\ExchangeAdmin"
```

## Execution Process

1. **Initialization**
   - Load Exchange Management Shell
   - Create log folder
   - Check parameters

2. **Pre-checks**
   - Check database existence
   - Check free disk space
   - Check active connections

3. **Defragmentation Preparation**
   - Stop Exchange services
   - Dismount database
   - Create backup (recommended)

4. **Defragmentation**
   - Execute `eseutil /d`
   - Monitor progress
   - Handle errors

5. **Completion**
   - Mount database
   - Start Exchange services
   - Check integrity
   - Calculate results

## Log Structure

The log file contains detailed information about each stage:

```
[2024-01-15 02:00:00] [INFO] ========== EXCHANGE DB DEFRAGMENTATION START ==========
[2024-01-15 02:00:01] [INFO] Database: MyDatabase
[2024-01-15 02:00:02] [INFO] Database size before defragmentation: 25.5 GB
[2024-01-15 02:00:03] [SUCCESS] Service MSExchangeIS stopped
[2024-01-15 02:00:04] [SUCCESS] Database dismounted
[2024-01-15 02:00:05] [SUCCESS] Starting defragmentation...
[2024-01-15 04:30:00] [SUCCESS] Defragmentation completed successfully in 150.0 minutes
[2024-01-15 04:30:30] [SUCCESS] Database size after defragmentation: 18.2 GB
[2024-01-15 04:30:31] [SUCCESS] Space freed: 7.3 GB (28.6%)
```

## Error Codes

- **Exit Code 0**: Successful completion
- **Exit Code 1**: Critical error or user cancellation

## Troubleshooting

### Insufficient Disk Space

```powershell
# Check free space
Get-WmiObject -Class Win32_LogicalDisk | Select-Object DeviceID, @{Name="FreeSpace(GB)";Expression={[math]::Round($_.FreeSpace/1GB,2)}}

# Clean temporary files
cleanmgr /sagerun:1

# Move database to another disk (if necessary)
Move-DatabasePath -Identity "MyDatabase" -EdbFilePath "D:\Databases\MyDatabase.edb"
```

### Services Not Starting

```powershell
# Manual start of Exchange services
Start-Service MSExchangeIS
Start-Service MSExchangeRPC
Start-Service MSExchangeTransport

# Check service status
Get-Service | Where-Object {$_.Name -like "*Exchange*"} | Select-Object Name, Status
```

### Database Won't Mount

```powershell
# Check database integrity
eseutil /mh "C:\Database\MyDatabase.edb"

# Restore from logs (if necessary)
eseutil /r E00 /l "C:\Database\Logs"

# Hard recovery (only as last resort)
eseutil /p "C:\Database\MyDatabase.edb"
```

## Recommendations

### Defragmentation Preparation

1. **Create Backup**
   ```powershell
   # Export all mailboxes to PST
   Get-Mailbox -Database "MyDatabase" | New-MailboxExportRequest -FilePath "\\BackupServer\Exports\{0}.pst"
   ```

2. **Notify Users**
   - Send maintenance notification
   - Specify email unavailability time

3. **Choose Appropriate Time**
   - Weekends
   - Non-business hours
   - Low activity periods

### Performance Optimization

```powershell
# Increase defragmentation process priority
# Add to script after starting eseutil:
$defragProcess = Get-Process eseutil
$defragProcess.PriorityClass = "High"
```

### Progress Monitoring

```powershell
# Create task for database size monitoring
while ($true) {
    $size = (Get-Item "C:\Database\MyDatabase.edb").Length / 1GB
    Write-Host "Current database size: $([math]::Round($size, 2)) GB"
    Start-Sleep -Seconds 300  # Check every 5 minutes
}
```

## Alternatives

### Online Defragmentation

```powershell
# Automatic online defragmentation (slower but no downtime)
# Configure in database properties or via PowerShell:
Set-MailboxDatabase -Identity "MyDatabase" -BackgroundDatabaseMaintenance $true
```

### Moving Mailboxes

```powershell
# Alternative: move all mailboxes to new database
New-MailboxDatabase -Name "NewDatabase" -EdbFilePath "D:\NewDB.edb"
Get-Mailbox -Database "MyDatabase" | New-MoveRequest -TargetDatabase "NewDatabase"
```

## Support

For support or bug reports:

1. Check the log file for detailed error information
2. Ensure all requirements are met
3. Refer to Microsoft Exchange Server 2016 documentation

## Compatibility

- ✅ Exchange Server 2016
- ✅ Windows Server 2012 R2 / 2016 / 2019
- ✅ PowerShell 5.0+
- ❌ Exchange Online (Office 365)
- ❌ Exchange Server 2013 (requires adaptation)

## License

This script is provided "as is" for use in corporate environments. Use at your own risk after testing in a test environment.