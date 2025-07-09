# Exchange JSON Manager Script

**Languages / Языки:**
- [🇺🇸 English](README.md) ← (Current)
- [🇷🇺 Русский](README.ru.md)

---

PowerShell script for monitoring Exchange Server databases via Zabbix with JSON data caching.

## Description

This script is designed to collect information about Exchange Server databases and provide this data to Zabbix monitoring system. It uses caching to reduce load on the Exchange server.

## Installation

1. Copy the script to the folder:
   ```
   C:\Scripts\db\exchange_json_manager.ps1
   ```

2. Ensure that the account running the script has permissions to:
   - Read Exchange configuration
   - Create/modify files in the `C:\Scripts\db\` folder

## Parameters

- `Action` - Action to perform (default: "discovery")
- `DatabaseName` - Database name (required for some actions)
- `CacheLifetime` - Cache lifetime in minutes (default: 30)

## Available Actions

### discovery
Returns JSON for Zabbix Low Level Discovery with information about all databases.

```powershell
.\exchange_json_manager.ps1 -Action discovery
```

**Example output:**
```json
{"data":[{"{#DBNAME}":"DB01","{#DBSERVER}":"EXCH01"},{"{#DBNAME}":"DB02","{#DBSERVER}":"EXCH01"}]}
```

### mounted
Checks the mount status of a specific database.

```powershell
.\exchange_json_manager.ps1 -Action mounted -DatabaseName "DB01"
```

**Returns:**
- `1` - database is mounted
- `0` - database is not mounted or not found

### size
Returns the size of a specific database in bytes.

```powershell
.\exchange_json_manager.ps1 -Action size -DatabaseName "DB01"
```

**Returns:** size in bytes (e.g., `5368709120`)

### status
Returns the total number of databases.

```powershell
.\exchange_json_manager.ps1 -Action status
```

**Returns:** number of databases (e.g., `3`)

### lastupdate
Returns the age of the last data update in minutes.

```powershell
.\exchange_json_manager.ps1 -Action lastupdate
```

**Returns:** 
- number of minutes since last update
- `9999` - if data is missing

### fileage
Returns the age of the JSON file in minutes.

```powershell
.\exchange_json_manager.ps1 -Action fileage
```

**Returns:**
- file age in minutes
- `9999` - if file doesn't exist

### forceupdate
Forces an update of Exchange data.

```powershell
.\exchange_json_manager.ps1 -Action forceupdate
```

**Returns:**
- `1` - update successful
- `0` - update failed

## JSON File Structure

The script creates a file `C:\Scripts\db\databases_info.json` with the following structure:

```json
{
  "LastUpdate": "2024-01-15 14:30:00",
  "LastUpdateUnix": 1705316200,
  "DatabaseCount": 2,
  "Databases": [
    {
      "Name": "DB01",
      "Server": {
        "Name": "EXCH01"
      },
      "Mounted": true,
      "MountedNumeric": 1,
      "DatabaseSize": "5.0 GB (5,368,709,120 bytes)",
      "DatabaseSizeBytes": 5368709120
    }
  ]
}
```

## Zabbix Configuration

### Creating User Parameters

Add to `zabbix_agentd.conf`:

```ini
# Exchange Database Discovery
UserParameter=exchange.db.discovery,powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\db\exchange_json_manager.ps1" -Action discovery

# Exchange Database Status
UserParameter=exchange.db.mounted[*],powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\db\exchange_json_manager.ps1" -Action mounted -DatabaseName "$1"

# Exchange Database Size
UserParameter=exchange.db.size[*],powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\db\exchange_json_manager.ps1" -Action size -DatabaseName "$1"

# Exchange General Status
UserParameter=exchange.db.status,powershell.exe -NoProfile -ExecutionPolicy