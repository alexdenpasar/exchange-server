# Exchange Log Cleanup Script

**Languages / Языки:**
- [🇺🇸 English](README.md) ← (Current)
- [🇷🇺 Русский](README.ru.md)

---

PowerShell script for automated cleanup of old Exchange Server transaction logs with database state verification.

## Description

This script is designed for safe cleanup of Exchange Server transaction log files. It checks the state of each database before deleting logs, ensuring operation safety and preventing data loss.

## ⚠️ Important Warnings

- **MANDATORY** create backups before running
- Script temporarily dismounts databases for verification
- Logs are deleted only in "Clean Shutdown" state
- Execute during non-business hours to minimize user impact
- Test in test environment before production use

## Features

### Safety
- ✅ Check database state before deleting logs
- ✅ Automatic database dismounting and mounting
- ✅ Protection against log deletion during Dirty Shutdown
- ✅ Detailed logging of all operations

### Automation
- ✅ Process all databases on server
- ✅ Configurable log retention period
- ✅ Automatic discovery of log paths
- ✅ Colored logging with importance levels

## Installation

1. Copy the script to the folder:
   ```
   C:\Scripts\Maintenance\ExchangeLogCleanup.ps1
   ```

2. Configure parameters at the beginning of the script:
   ```powershell
   $LogRetentionDays = 7  # Log retention period (in days)
   ```

## Configuration Parameters

- **`$LogRetentionDays`** - Number of days to retain logs (default: 7)
- **`$CurrentServer`** - Current server name (automatically determined)
- **`$ScriptPath`** - Script directory (automatically determined)
- **`$LogFile`** - Script log file path (automatically created)

## Usage

### Basic Execution

```powershell
# Run from Exchange Management Shell
.\ExchangeLogCleanup.ps1
```

### Changing Parameters

```powershell
# Change log retention period
# Edit in script:
$LogRetentionDays = 14  # Keep logs for 14 days
```

### Scheduling via Task Scheduler

```powershell
# Create task for weekly execution
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -File C:\Scripts\Maintenance\ExchangeLogCleanup.ps1"
$Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 3:00AM
$Settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -WakeToRun
Register-ScheduledTask -TaskName "Exchange Log Cleanup" -Action $Action -Trigger $Trigger -Settings $Settings -User "DOMAIN\ExchangeAdmin"
```

## Workflow Algorithm

### 1. Initialization
- Create log file with timestamp
- Determine current Exchange server
- Get list of databases

### 2. For Each Database
1. **Dismount database**
   ```powershell
   Dismount-Database -Identity $DB -Confirm:$false
   ```

2. **Check state**
   ```powershell
   eseutil /mh $DBPath  # Check for Clean Shutdown
   ```

3. **Delete logs** (only during Clean Shutdown)
   ```powershell
   Get-ChildItem -Path $LogPath -Filter "*.log" | 
   Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } | 
   Remove-Item -Force
   ```

4. **Mount database**
   ```powershell
   Mount-Database -Identity $DB
   ```

### 3. Completion
- Check status of all databases
- Create summary report

## Log Structure

Log file is created in script folder with name:
```
ExchangeCleanup_YYYY-MM-DD_HH-MM-SS.log
```

### Example Log

```
2024-01-15 03:00:00 [INFO] === Starting Exchange log cleanup on server EXCH01 ===
2024-01-15 03:00:01 [INFO] Found databases on server EXCH01 : DB01, DB02, DB03
2024-01-15 03:00:02 [INFO] 
Processing database: DB01
2024-01-15 03:00:03 [INFO] Dismounting database: DB01
2024-01-15 03:00:08 [INFO] Database DB01 successfully dismounted. Checking state...
2024-01-15 03:00:09 [INFO] Database DB01 in Clean Shutdown state. Deleting old logs...
2024-01-15 03:00:10 [INFO] Deleted 25 logs in folder D:\Logs\DB01
2024-01-15 03:00:11 [INFO] Starting database: DB01
2024-01-15 03:00:16 [INFO] Database DB01 successfully started.
```

### Color Scheme

- **INFO** - White text (normal information)
- **WARNING** - Yellow text (warnings)
- **ERROR** - Red text (errors)

## Database State Verification

### Clean Shutdown vs Dirty Shutdown

**Clean Shutdown:**
- Database was properly dismounted
- All transactions committed
- Logs can be safely deleted

**Dirty Shutdown:**
- Database was improperly dismounted
- Possible uncommitted transactions
- Logs NOT deleted (needed for recovery)

### Verification Command

```cmd
eseutil /mh "D:\Database\DB01.edb"
```

**Result for Clean Shutdown:**
```
State: Clean Shutdown
```

**Result for Dirty Shutdown:**
```
State: Dirty Shutdown
```

## Troubleshooting

### Database Won't Dismount

**Symptoms:**
- Script cannot dismount database
- Error message during Dismount-Database execution

**Solution:**
```powershell
# Check active connections
Get-StoreUsageStatistics -Database "DB01" | Where-Object {$_.TimeInServer -gt 0}

# Force dismount
Dismount-Database -Identity "DB01" -Confirm:$false -Force
```

### Dirty Shutdown After Dismount

**Symptoms:**
- Database shows Dirty Shutdown after proper dismount
- Logs not deleted

**Solution:**
```powershell
# Check database integrity
eseutil /mh "D:\Database\DB01.edb"

# Restore from logs (if necessary)
eseutil /r E00 /l "D:\Logs\DB01"

# Check after recovery
eseutil /mh "D:\Database\DB01.edb"
```

### Database Won't Mount

**Symptoms:**
- Error when trying to mount database
- Database remains in Dismounted state

**Solution:**
```powershell
# Check errors in event logs
Get-EventLog -LogName Application -Source "MSExchange*" -Newest 10

# Check integrity
eseutil /mh "D:\Database\DB01.edb"

# Force mount
Mount-Database -Identity "DB01" -Force
```

## Recommendations

### Cleanup Scheduling

1. **Weekly cleanup** (recommended)
   ```powershell
   $LogRetentionDays = 7
   # Run every Sunday at 3:00 AM
   ```

2. **Monthly cleanup**
   ```powershell
   $LogRetentionDays = 30
   # Run on first day of each month
   ```

### Disk Space Monitoring

```powershell
# Check free space before and after cleanup
Get-WmiObject -Class Win32_LogicalDisk | 
Select-Object DeviceID, @{Name="FreeSpace(GB)";Expression={[math]::Round($_.FreeSpace/1GB,2)}}
```

### Backup

```powershell
# Create backup of logs before deletion
$BackupPath = "\\BackupServer\Exchange\Logs\$(Get-Date -Format 'yyyyMMdd')"
New-Item -ItemType Directory -Path $BackupPath -Force
Copy-Item -Path "D:\Logs\*" -Destination $BackupPath -Recurse
```

## Automation

### Monitoring System Integration

```powershell
# Send notifications about results
if ($ErrorCount -gt 0) {
    Send-MailMessage -To "admin@company.com" -Subject "Exchange Log Cleanup - Errors" -Body "Errors found during log cleanup"
}
```

### Creating Reports

```powershell
# Analyze cleanup results
$LogPath = "C:\Scripts\Maintenance"
$CleanupLogs = Get-ChildItem $LogPath -Filter "ExchangeCleanup_*.log"

foreach ($Log in $CleanupLogs) {
    $Content = Get-Content $Log.FullName
    $DeletedCount = ($Content | Select-String "Deleted \d+ logs").Matches.Count
    $ErrorCount = ($Content | Select-String "\[ERROR\]").Count
    
    Write-Host "Log: $($Log.Name), Deleted: $DeletedCount, Errors: $ErrorCount"
}
```

## Requirements

- Exchange Server 2016/2019
- PowerShell 5.0 or higher
- Exchange administrator rights
- Local administrator rights on server

## Compatibility

- ✅ Exchange Server 2016
- ✅ Exchange Server 2019
- ✅ Windows Server 2012 R2 / 2016 / 2019
- ⚠️ Exchange Server 2013 (requires testing)
- ❌ Exchange Online (not applicable)

## Security

- Script checks database state before deleting logs
- Automatic database recovery after operations
- Detailed logging for audit
- Protection against accidental deletion of important logs

## Support

For help:

1. Check script log file
2. Use Event Viewer for Exchange diagnostics
3. Check database state via Exchange Management Console
4. Refer to Microsoft Exchange documentation

## License

This script is provided "as is" for use in corporate environments. Test before production use.