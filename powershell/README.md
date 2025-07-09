# Exchange Server 2016 - PowerShell Commands Reference

**Languages / Языки:**
- [🇺🇸 English](README.md) ← (Current)
- [🇷🇺 Русский](README.ru.md)

---

## Migration

### Request to move mailbox to another database (with archive)
```powershell
New-MoveRequest -Identity "email@example.com" -DomainController "dc.example.com" -TargetDatabase "Name_DB" -BadItemLimit 1 -AcceptLargeDataLoss -ErrorAction Stop
```

### Move only primary mailbox to another database without archive
```powershell
New-MoveRequest -Identity "email@example.com" -DomainController "dc.example.com" -TargetDatabase "Name_DB" -BadItemLimit 3 -AcceptLargeDataLoss -PrimaryOnly -ErrorAction Stop
```

### Move only archive to another database
```powershell
New-MoveRequest -Identity "email@example.com" -DomainController "dc.example.com" -ArchiveOnly -ArchiveTargetDatabase "Name_DB-Archive"
```

### View active migrations
```powershell
Get-MoveRequest -DomainController "dc.example.com" -ErrorAction SilentlyContinue | Where-Object { $_.TargetDatabase -eq "Name_DB" }
```

### View status of all migrations
```powershell
Get-MoveRequest | Get-MoveRequestStatistics | ft DisplayName, Status, PercentComplete, TotalMailboxSize
```

### Remove completed and failed migrations
```powershell
Remove-MoveRequest -Identity "email@example.com" -DomainController "dc.example.com" -Confirm:$false -ErrorAction SilentlyContinue
```

### Suspend migration
```powershell
Suspend-MoveRequest -Identity "email@example.com" -SuspendComment "Suspended for maintenance"
```

### Resume migration
```powershell
Resume-MoveRequest -Identity "email@example.com"
```

## Database Management

### Dismount database
```powershell
Dismount-Database "Name_DB" -Confirm:$false
```

### Mount database
```powershell
Mount-Database "Name_DB" -Confirm:$false
```

### Check database status
```cmd
eseutil /mh D:\mailbox\db.edb
```

### Check database logs (must be in database folder)
```cmd
eseutil /ml E00
```

### Restore database from logs (must be in database folder)
```cmd
eseutil /r E00
```

### Hard recovery of database
```cmd
eseutil /p D:\mailbox\db.edb
```

### Defragment database after deleting/moving mailboxes
```cmd
eseutil /d D:\mailbox\db.edb
```

### Check database for errors and repair
```powershell
New-MailboxRepairRequest -Database "Name_DB" -CorruptionType SearchFolder,AggregateCounts,ProvisionedFolder,FolderView
```

### Track progress of database integrity check
```powershell
Get-MailboxRepairRequest -Database "Name_DB"
```

### View mailboxes in database
```powershell
Get-MailboxStatistics -Database "Name_DB" -DomainController "dc.example.com" | ft DisplayName, TotalItemSize, ItemCount
Get-MailboxStatistics -Database "Name_DB" -DomainController "dc.example.com" | ft DisplayName, TotalItemSize
```

### View where mailbox archive is located
```powershell
Get-Mailbox -Identity "email@example.com" -DomainController "dc.example.com" -Archive | Select-Object DisplayName, ArchiveDatabase
```

### More detailed view
```powershell
Get-Mailbox -Database "Name_DB" -DomainController "dc.example.com" -Archive | Select-Object DisplayName, Alias, ArchiveDatabase
```

### View all databases and their status
```powershell
Get-MailboxDatabase | ft Name, Server, Mounted, DatabaseSize
```

### Create new database
```powershell
New-MailboxDatabase -Name "NewDB" -EdbFilePath "D:\mailbox\NewDB.edb" -LogFolderPath "D:\logs\NewDB"
```

## Mailbox Management

### Grant full access to mailbox
```powershell
Add-MailboxPermission -Identity "mailbox_user@domain.com" -User "your_user@domain.com" -DomainController "dc.example.com" -AccessRights FullAccess -InheritanceType All
```

### Grant send on behalf permission
```powershell
Set-Mailbox -Identity "mailbox_user@domain.com" -DomainController "dc.example.com" -GrantSendOnBehalfTo "your_user@domain.com"
```

### Allow send as
```powershell
Add-ADPermission -Identity "mailbox_user" -User "your_user" -DomainController "dc.example.com" -ExtendedRights "Send As"
```

### View permissions on mailbox
```powershell
Get-MailboxPermission -Identity "ceo@domain.com" -DomainController "dc.example.com" | Where-Object { $_.User -like "*alex*" }
```

### Remove permissions (FullAccess example)
```powershell
Remove-MailboxPermission -Identity "ceo@domain.com" -User "alex@domain.com" -DomainController "dc.example.com" -AccessRights FullAccess -InheritanceType All
```

### Create new mailbox
```powershell
New-Mailbox -Name "John Doe" -UserPrincipalName "john.doe@domain.com" -SamAccountName "john.doe" -Database "Name_DB" -Password (ConvertTo-SecureString -String "Password123!" -AsPlainText -Force)
```

### Enable archive for mailbox
```powershell
Enable-Mailbox -Identity "user@domain.com" -Archive -ArchiveDatabase "Archive_DB"
```

### Disable archive
```powershell
Disable-Mailbox -Identity "user@domain.com" -Archive
```

### Set mailbox quotas
```powershell
Set-Mailbox -Identity "user@domain.com" -IssueWarningQuota 1GB -ProhibitSendQuota 1.2GB -ProhibitSendReceiveQuota 1.5GB
```

### Disable mailbox (keep in database)
```powershell
Disable-Mailbox -Identity "user@domain.com" -Confirm:$false
```

### Remove mailbox from database
```powershell
Remove-Mailbox -Identity "user@domain.com" -Confirm:$false
```

## Distribution Groups

### Create distribution group
```powershell
New-DistributionGroup -Name "IT Team" -SamAccountName "ITTeam" -PrimarySmtpAddress "it@domain.com"
```

### Add user to group
```powershell
Add-DistributionGroupMember -Identity "ITTeam" -Member "user@domain.com"
```

### Remove user from group
```powershell
Remove-DistributionGroupMember -Identity "ITTeam" -Member "user@domain.com"
```

### View group members
```powershell
Get-DistributionGroupMember -Identity "ITTeam" | ft Name, PrimarySmtpAddress
```

### Create dynamic distribution group
```powershell
New-DynamicDistributionGroup -Name "All Users" -RecipientFilter "RecipientType -eq 'UserMailbox'"
```

## Public Folders

### Create public folder
```powershell
New-PublicFolder -Name "Company Documents" -Path "\"
```

### Create public folder database
```powershell
New-PublicFolderDatabase -Name "Public Folder DB" -EdbFilePath "D:\PublicFolders\PFDB.edb"
```

### Mail-enable public folder
```powershell
Enable-MailPublicFolder -Identity "\Company Documents" -ExternalEmailAddress "docs@domain.com"
```

### Set permissions on public folder
```powershell
Add-PublicFolderClientPermission -Identity "\Company Documents" -User "user@domain.com" -AccessRights Editor
```

## Transport and Rules

### Create transport rule
```powershell
New-TransportRule -Name "Block External Attachments" -FromScope NotInOrganization -AttachmentHasExecutableContent $true -RejectMessageReasonText "Executable attachments are blocked"
```

### View transport rules
```powershell
Get-TransportRule | ft Name, State, Priority
```

### Disable rule
```powershell
Disable-TransportRule -Identity "Block External Attachments"
```

### Create send connector
```powershell
New-SendConnector -Name "Internet Connector" -Usage Internet -AddressSpaces "SMTP:*;1" -SourceTransportServers "EXCH01"
```

### View message queues
```powershell
Get-Queue | ft Identity, Status, MessageCount, NextHopDomain
```

## Monitoring and Statistics

### View mailbox statistics by size
```powershell
Get-MailboxStatistics | Sort-Object TotalItemSize -Descending | ft DisplayName, TotalItemSize, ItemCount
```

### View top 10 largest mailboxes
```powershell
Get-MailboxStatistics | Sort-Object TotalItemSize -Descending | Select-Object -First 10 | ft DisplayName, TotalItemSize
```

### Check Exchange services status
```powershell
Get-Service | Where-Object {$_.Name -like "*Exchange*"} | ft Name, Status
```

### View Exchange event logs
```powershell
Get-EventLog -LogName Application -Source "MSExchange*" -Newest 50 | ft TimeGenerated, EntryType, Source, Message
```

### Test mailbox connectivity
```powershell
Test-MapiConnectivity -Identity "user@domain.com"
```

### Test mail flow
```powershell
Test-MailFlow -TargetEmailAddress "test@domain.com"
```

## Backup and Restore

### Create database backup
```powershell
New-MailboxExportRequest -Mailbox "user@domain.com" -FilePath "\\backup\exports\user.pst"
```

### Import from PST file
```powershell
New-MailboxImportRequest -Mailbox "user@domain.com" -FilePath "\\backup\imports\user.pst"
```

### View export/import status
```powershell
Get-MailboxExportRequest | ft Name, Status, PercentComplete
Get-MailboxImportRequest | ft Name, Status, PercentComplete
```

## Certificates

### View certificates
```powershell
Get-ExchangeCertificate | ft Thumbprint, Subject, NotAfter, Services
```

### Assign certificate to services
```powershell
Enable-ExchangeCertificate -Thumbprint "THUMBPRINT" -Services IIS,SMTP,POP,IMAP
```

### Create certificate request
```powershell
New-ExchangeCertificate -GenerateRequest -SubjectName "CN=mail.domain.com" -DomainName "mail.domain.com","autodiscover.domain.com" -Path "C:\cert_request.req"
```

## Useful Diagnostic Commands

### Check database replication (for DAG)
```powershell
Get-MailboxDatabaseCopyStatus | ft Name, Status, CopyQueueLength, ReplayQueueLength
```

### View user activity
```powershell
Get-MailboxStatistics -Identity "user@domain.com" | Select-Object DisplayName, LastLogonTime, LastLogoffTime
```

### Search messages in mailboxes
```powershell
New-MailboxSearch -Name "SearchName" -SourceMailboxes "user@domain.com" -SearchQuery "Subject:'Important Meeting'" -TargetMailbox "admin@domain.com" -TargetFolder "SearchResults"
```

### Clear transport logs
```powershell
Set-TransportService -Identity "EXCH01" -MessageTrackingLogMaxAge 30.00:00:00
```

### View virtual directory configuration
```powershell
Get-OwaVirtualDirectory | ft Name, Server, InternalUrl, ExternalUrl
Get-ActiveSyncVirtualDirectory | ft Name, Server, InternalUrl, ExternalUrl
```

---

## Notes

- Replace `dc.example.com` with your domain controller
- Replace `Name_DB` with your database name
- Replace `domain.com` with your domain
- Always test commands in a test environment before applying in production
- Regularly create backups before performing critical operations

## Connecting to Exchange Management Shell

```powershell
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
```

Or use Exchange Management Shell directly from the Start menu.