# Exchange 2016 Migration Scripts

**Languages / Языки:**
- [🇺🇸 English](README.md) ← (Current)
- [🇷🇺 Русский](README.ru.md)

---

Set of automated PowerShell scripts for safe migration of Exchange Server 2016 mailboxes with monitoring and parallelism control.

## Package Contents

### Migration Scripts

1. **`Email_PrimaryOnly_Migrations.ps1`** - Primary mailbox migration
2. **`Email_ArchiveOnly_Migrations.ps1`** - Archive-only migration
3. **`Migration_Monitor.ps1`** - Active migration monitoring

### Configuration Files

4. **`EmailList.txt`** - List of email addresses for migration

## Features

### Safety
- ✅ Control number of parallel migrations
- ✅ Automatic detection of hung migrations
- ✅ Restart of failed migrations
- ✅ Detailed logging of all operations

### Monitoring
- ✅ Real-time status tracking
- ✅ Color-coded migration status indication
- ✅ Progress display
- ✅ Automatic status updates

### Resilience
- ✅ Handle existing migrations
- ✅ Recovery from failures
- ✅ Protection against duplicate runs
- ✅ Automatic cleanup of completed requests

## Installation

### Folder Structure

```
C:\Scripts\Migration\
├── Email_PrimaryOnly_Migrations.ps1
├── Email_ArchiveOnly_Migrations.ps1
├── Migration_Monitor.ps1
├── EmailList.txt
└── Logs\
    ├── MailboxMigration_YYYYMMDD_HHMMSS.log
    └── MailboxMigration_OnlyArchive_YYYYMMDD_HHMMSS.log
```

### Parameter Configuration

Edit parameters in each script:

```powershell
# Main parameters
$TargetDatabase = "DB-Archive"            # Target database
$DomainController = "dc.example.com"      # Domain controller
$EmailListFile = "C:\Scripts\Migration\EmailList.txt"  # Email list file
$MaxParallelMoves = 3                     # Maximum parallel migrations
$BadItemLimit = 100                       # Bad item limit
$CheckInterval = 60                       # Check interval (seconds)
```

### Email List Preparation

Create `EmailList.txt` file with email addresses:

```
user1@example.com
user2@example.com
user3@example.com
# Comments start with #
# user4@example.com - temporarily disabled
```

## Usage

### Primary Mailbox Migration

```powershell
# Run from Exchange Management Shell
.\Email_PrimaryOnly_Migrations.ps1
```

**Features:**
- Moves primary mailboxes (without archives)
- Uses `-PrimaryOnly` parameter
- Recommended for large mailboxes
- Maximum 3 parallel migrations by default

### Archive-Only Migration

```powershell
# Run from Exchange Management Shell
.\Email_ArchiveOnly_Migrations.ps1
```

**Features:**
- Moves only mailbox archives
- Uses `-ArchiveOnly` parameter
- Less performance impact
- Maximum 1 parallel migration by default

### Migration Monitoring

```powershell
# Run monitoring
.\Migration_Monitor.ps1
```

**Monitoring capabilities:**
- Real-time status tracking
- Color-coded status indication
- Automatic updates every 30 seconds
- Stop with Ctrl+C

## Implementation Details

### Parallelism Control

Scripts control the number of simultaneously running migrations:

```powershell
# If limit reached, wait for completion
while ($InProgress.Count -ge $MaxParallelMoves) {
    # Check status of active migrations
    # Remove completed from queue
    # Wait for slot to free up
}
```

### Hung Migration Detection

Automatic detection and restart of hung migrations:

```powershell
# Check execution time
if (((Get-Date) - $InProgress[$ActiveEmail]).TotalHours -gt 24) {
    # Analyze cause of hang
    # Restart if necessary
}
```

### Log Structure

Detailed logs for each migration:

```
2024-01-15 10:30:00 - Mailbox migration script started
2024-01-15 10:30:01 - Target DB: DB-Archive, Max parallel migrations: 3
2024-01-15 10:30:02 - Total email addresses found: 50
2024-01-15 10:30:03 - Starting migration of mailbox user1@example.com
2024-01-15 10:30:04 - Migration request created for user1@example.com
2024-01-15 11:45:30 - Migration for user1@example.com completed (Status: Completed, Percent: 100%)
```

## Status Indication in Monitor

### Status Color Scheme

- 🟢 **Green** - Successful execution
  - `Completed`, `CopyingMessages`, `ScanningForMessages`
- 🟡 **Yellow** - Warnings
  - `InProgress`, `StalledDueToMail_*`, `Suspended`
- 🔴 **Red** - Errors
  - `Failed`, `StalledDueToSource_*`, `StalledDueToTarget_*`
- 🔵 **Blue** - Waiting
  - `Queued`, `WaitingForJobPickup`

### Status Interpretation

```
User: John Doe (john.doe@example.com)
Status: InProgress
Detailed status: CopyingMessages
Progress: 75%
Source: DB-Old
Target: DB-Archive
```

## Troubleshooting

### Hung Migrations

**Symptoms:**
- Migration running for more than 24 hours
- Progress not changing for extended time
- Status `StalledDueToSource_*` or `StalledDueToTarget_*`

**Solution:**
```powershell
# Manual migration check
Get-MoveRequestStatistics -Identity "user@example.com" -IncludeReport

# Restart hung migration
Remove-MoveRequest -Identity "user@example.com" -Confirm:$false
New-MoveRequest -Identity "user@example.com" -TargetDatabase "DB-Archive"
```

### "BadItemLimit" Errors

**Symptoms:**
- Migration stops with limit exceeded error
- Messages about corrupted items in logs

**Solution:**
```powershell
# Increase limit in script
$BadItemLimit = 500

# Or use AcceptLargeDataLoss
New-MoveRequest -Identity "user@example.com" -TargetDatabase "DB-Archive" -BadItemLimit 1000 -AcceptLargeDataLoss
```

### Performance Issues

**Recommendations:**
- Reduce `$MaxParallelMoves` for large mailboxes
- Run migrations during non-business hours
- Monitor disk and network usage

```powershell
# Settings for slow systems
$MaxParallelMoves = 1
$CheckInterval = 120  # Increase check interval
```

## Usage Recommendations

### Migration Planning

1. **Preparation:**
   ```powershell
   # Create backups
   Get-Mailbox -Database "SourceDB" | New-MailboxExportRequest -FilePath "\\BackupServer\{0}.pst"
   
   # Check target database quotas
   Get-MailboxDatabase "TargetDB" | fl ProhibitSendQuota,ProhibitSendReceiveQuota
   ```

2. **Optimization:**
   ```powershell
   # Pre-defragmentation of source
   .\ExchangeDefrag.ps1 -DatabaseName "SourceDB"
   
   # Performance monitoring
   Get-Counter "\MSExchange Database(*)\Database Page Fault Stalls/sec"
   ```

### Batch Migration

```powershell
# Migration by groups
$Groups = @("Sales", "Marketing", "IT")
foreach ($Group in $Groups) {
    Get-ADGroupMember $Group | ForEach-Object { $_.Mail } | Out-File "EmailList_$Group.txt"
    .\Email_PrimaryOnly_Migrations.ps1 -EmailListFile "EmailList_$Group.txt"
}
```

### Task Scheduler Automation

```powershell
# Create task for nightly migrations
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -File C:\Scripts\Migration\Email_PrimaryOnly_Migrations.ps1"
$Trigger = New-ScheduledTaskTrigger -Daily -At 2:00AM
$Settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable
Register-ScheduledTask -TaskName "Exchange Migration" -Action $Action -Trigger $Trigger -Settings $Settings
```

## Monitoring and Reporting

### Migration Statistics

```powershell
# Analyze migration logs
$LogPath = "C:\Scripts\Migration\Logs"
$LogFiles = Get-ChildItem $LogPath -Filter "*.log"

foreach ($LogFile in $LogFiles) {
    $Content = Get-Content $LogFile.FullName
    $Successful = ($Content | Select-String "completed successfully").Count
    $Failed = ($Content | Select-String "failed").Count
    
    Write-Host "File: $($LogFile.Name)"
    Write-Host "Successful: $Successful, Failed: $Failed"
}
```

### Creating Reports

```powershell
# Weekly migration report
$Report = @()
Get-MoveRequest | Get-MoveRequestStatistics | ForEach-Object {
    $Report += [PSCustomObject]@{
        User = $_.DisplayName
        Status = $_.Status
        PercentComplete = $_.PercentComplete
        SourceDB = $_.SourceDatabase
        TargetDB = $_.TargetDatabase
        CompletionTimestamp = $_.CompletionTimestamp
    }
}

$Report | Export-Csv "MigrationReport_$(Get-Date -Format 'yyyyMMdd').csv" -NoTypeInformation
```

## Compatibility

- ✅ Exchange Server 2016
- ✅ Windows Server 2012 R2 / 2016 / 2019
- ✅ PowerShell 5.0+
- ⚠️ Exchange Server 2013 (requires adaptation)
- ❌ Exchange Online (use other tools)

## Security

- Scripts require Exchange Organization Management rights
- Logs contain only necessary information (no passwords)
- Support for operation interruption via Ctrl+C
- Automatic cleanup of temporary files

## Support

For help:

1. Check log files in `Logs\` folder
2. Use monitoring for diagnostics
3. Refer to Microsoft Exchange documentation
4. Check access rights and service status

## License

Scripts are provided "as is" for use in corporate environments. Test in test environment before production use.