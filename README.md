# Exchange Server 2016/2019 - Scripts and PowerShell Commands Repository

**Languages / Языки:**
- [🇷🇺 Русский](README.ru.md)
- [🇺🇸 English](README.md) (current)

Comprehensive collection of PowerShell scripts and commands for Exchange Server 2016/2019 administration, automation, and maintenance.

## 🚀 Overview

This repository contains a complete toolkit for Exchange Server administrators, featuring:

- **PowerShell Command Reference** - Essential commands for daily operations
- **Automation Scripts** - Production-ready scripts for complex tasks
- **Monitoring Tools** - Real-time monitoring and alerting solutions
- **Migration Utilities** - Safe and efficient mailbox migration tools
- **Maintenance Scripts** - Database defragmentation and log cleanup tools

## 📁 Repository Structure

```
exchange-server-2016/
├── README.md                           # This file
├── powershell/                         # PowerShell commands reference
│   └── README.md                       # Complete command reference
├── scripts/                            # Automation scripts
│   ├── db/                            # Database monitoring
│   │   ├── exchange_db_discovery.ps1  # Zabbix monitoring script
│   │   ├── databases_info.json        # Database cache file
│   │   └── README.md                  # Database monitoring guide
│   ├── defrag/                        # Database defragmentation
│   │   ├── defrag_db.ps1              # Database defrag script
│   │   └── README.md                  # Defragmentation guide
│   ├── logs/                          # Log management
│   │   ├── ExchangeLogCleanup.ps1     # Transaction log cleanup
│   │   └── README.md                  # Log cleanup guide
│   └── migration/                     # Mailbox migration
│       ├── Email_PrimaryOnly_Migrations.ps1
│       ├── Email_ArchiveOnly_Migrations.ps1
│       ├── Migration_Monitor.ps1
│       ├── EmailList.txt
│       └── README.md                  # Migration guide
```

## 🔧 Quick Start

### Prerequisites

- Exchange Server 2016 or 2019
- PowerShell 5.0 or higher
- Exchange Management Shell
- Administrator privileges

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/your-org/exchange-server-2016.git
   cd exchange-server-2016
   ```

2. **Set execution policy:**
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **Load Exchange Management Shell:**
   ```powershell
   Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
   ```

## 📚 Documentation

### [PowerShell Commands Reference](powershell/README.md)
Complete reference guide with over 100 essential Exchange PowerShell commands organized by category:

- **Migration** - Mailbox and archive migration commands
- **Database Management** - Database operations and maintenance
- **Mailbox Management** - User mailbox administration
- **Distribution Groups** - Group management and membership
- **Public Folders** - Shared folder administration
- **Transport Rules** - Mail flow and security rules
- **Monitoring** - System health and performance monitoring
- **Certificates** - SSL/TLS certificate management
- **Backup/Restore** - Data protection operations

### Script Documentation

#### [Database Monitoring](scripts/db/README.md)
- **Exchange JSON Manager** - Zabbix integration for database monitoring
- Real-time database status and size monitoring
- Automated alerts and reporting
- Caching system for performance optimization

#### [Database Defragmentation](scripts/defrag/README.md)
- **Automated Database Defragmentation** - Safe offline defragmentation
- Disk space optimization after mailbox migrations
- Automatic service management and recovery
- Comprehensive logging and error handling

#### [Log Management](scripts/logs/README.md)
- **Exchange Log Cleanup** - Automated transaction log cleanup
- Safe log deletion with database state verification
- Configurable retention policies
- Clean/Dirty shutdown detection

#### [Migration Tools](scripts/migration/README.md)
- **Primary Mailbox Migration** - Bulk mailbox migration with parallelization
- **Archive-Only Migration** - Separate archive migration for performance
- **Migration Monitor** - Real-time migration status tracking
- Advanced error handling and recovery mechanisms

## 🛠️ Key Features

### Production-Ready Scripts
- ✅ **Error Handling** - Comprehensive error detection and recovery
- ✅ **Logging** - Detailed operation logs with timestamps
- ✅ **Safety Checks** - Pre-execution validation and confirmation
- ✅ **Rollback Support** - Automatic recovery from failures

### Monitoring & Alerting
- ✅ **Zabbix Integration** - Native monitoring system support
- ✅ **Real-time Status** - Live migration and system monitoring
- ✅ **Performance Metrics** - Database size, mount status, and health
- ✅ **Automated Alerts** - Email notifications for critical events

### Migration Excellence
- ✅ **Parallel Processing** - Configurable concurrent migration limits
- ✅ **Progress Tracking** - Real-time migration progress monitoring
- ✅ **Hung Migration Detection** - Automatic detection and restart
- ✅ **Batch Operations** - Bulk migration from email lists

### Maintenance Automation
- ✅ **Scheduled Operations** - Task Scheduler integration
- ✅ **Database Optimization** - Automated defragmentation workflows
- ✅ **Log Cleanup** - Intelligent transaction log management
- ✅ **Health Monitoring** - Continuous system health checks

## 🎯 Common Use Cases

### Daily Operations
```powershell
# Check database status
Get-MailboxDatabase | ft Name, Server, Mounted, DatabaseSize

# Monitor active migrations
.\scripts\migration\Migration_Monitor.ps1

# View mailbox statistics
Get-MailboxStatistics | Sort-Object TotalItemSize -Descending | Select-Object -First 10
```

### Maintenance Tasks
```powershell
# Run database defragmentation
.\scripts\defrag\defrag_db.ps1

# Clean up old transaction logs
.\scripts\logs\ExchangeLogCleanup.ps1

# Update monitoring cache
.\scripts\db\exchange_db_discovery.ps1 -Action forceupdate
```

### Migration Projects
```powershell
# Start primary mailbox migration
.\scripts\migration\Email_PrimaryOnly_Migrations.ps1

# Monitor migration progress
.\scripts\migration\Migration_Monitor.ps1

# Migrate archives separately
.\scripts\migration\Email_ArchiveOnly_Migrations.ps1
```

## 📊 Monitoring Integration

### Zabbix Templates
The repository includes complete Zabbix integration:

```ini
# Add to zabbix_agentd.conf
UserParameter=exchange.db.discovery,powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\db\exchange_db_discovery.ps1" -Action discovery
UserParameter=exchange.db.mounted[*],powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\db\exchange_db_discovery.ps1" -Action mounted -DatabaseName "$1"
UserParameter=exchange.db.size[*],powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\db\exchange_db_discovery.ps1" -Action size -DatabaseName "$1"
```

### Performance Metrics
- Database mount status and health
- Database size and growth trends
- Migration progress and completion rates
- Transaction log accumulation monitoring

## 🔐 Security Best Practices

### Permissions
- Use dedicated service accounts with minimal required privileges
- Implement role-based access control (RBAC)
- Regular audit of administrative access

### Logging & Auditing
- All scripts include comprehensive logging
- Sensitive information is never logged
- Audit trails for all administrative actions

### Data Protection
- Pre-execution backup validation
- Automatic rollback capabilities
- Safe failure modes with service recovery

## 📈 Performance Optimization

### Migration Performance
- Configurable parallel migration limits
- Automatic load balancing across databases
- Network and disk I/O optimization

### Database Optimization
- Automated defragmentation scheduling
- Intelligent log cleanup policies
- Proactive space management

### Monitoring Efficiency
- Cached data collection for reduced server load
- Optimized query patterns
- Minimal impact monitoring intervals

## 🤝 Contributing

### Code Standards
- Follow PowerShell best practices
- Include comprehensive error handling
- Document all parameters and functions
- Provide usage examples

### Testing Requirements
- Test in non-production environments first
- Validate with different Exchange versions
- Performance testing with large datasets
- Documentation updates for new features

### Pull Request Process
1. Fork the repository
2. Create a feature branch
3. Implement changes with tests
4. Update documentation
5. Submit pull request with detailed description

## 📋 Changelog

### Version 2.0.0 (Current)
- ✅ Complete migration script redesign
- ✅ Enhanced monitoring with Zabbix integration
- ✅ Improved error handling and logging
- ✅ Added database defragmentation automation
- ✅ Comprehensive documentation updates

### Version 1.5.0
- ✅ Added migration monitoring tools
- ✅ Implemented parallel migration controls
- ✅ Enhanced log cleanup functionality

### Version 1.0.0
- ✅ Initial PowerShell command collection
- ✅ Basic migration scripts
- ✅ Simple monitoring tools

## 🆘 Support & Troubleshooting

### Common Issues

1. **Permission Denied Errors**
   ```powershell
   # Check Exchange permissions
   Get-ManagementRoleAssignment -RoleAssignee "username"
   
   # Add required permissions
   New-ManagementRoleAssignment -Role "Mailbox Import Export" -User "username"
   ```

2. **Migration Failures**
   ```powershell
   # Check migration status
   Get-MoveRequest | Get-MoveRequestStatistics | ft DisplayName, Status, PercentComplete
   
   # Restart failed migrations
   Get-MoveRequest -MoveStatus Failed | Resume-MoveRequest
   ```

3. **Database Mount Issues**
   ```powershell
   # Check database state
   eseutil /mh "C:\Database\DB.edb"
   
   # Force mount if needed
   Mount-Database -Identity "DB01" -Force
   ```

### Getting Help

1. **Check the logs** - All scripts generate detailed logs
2. **Review documentation** - Each script has comprehensive README
3. **Test in lab environment** - Always test before production use
4. **Community support** - Submit issues for help and improvements

---

**⚠️ Disclaimer**: These scripts are provided "as is" without warranty. Always test in a non-production environment before deploying to production systems. Create backups before running any maintenance operations.

**📚 Quick Links**:
- [PowerShell Commands](powershell/README.md)
- [Database Monitoring](scripts/db/README.md)
- [Migration Tools](scripts/migration/README.md)
- [Defragmentation Guide](scripts/defrag/README.md)
- [Log Management](scripts/logs/README.md)
