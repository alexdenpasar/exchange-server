# Exchange JSON Manager Script

PowerShell скрипт для мониторинга баз данных Exchange Server через Zabbix с кешированием данных в JSON формате.

## Описание

Скрипт предназначен для сбора информации о базах данных Exchange Server и предоставления этих данных системе мониторинга Zabbix. Использует кеширование для снижения нагрузки на сервер Exchange.

## Установка

1. Скопируйте скрипт в папку:
   ```
   C:\Scripts\db\exchange_json_manager.ps1
   ```

2. Убедитесь, что у учетной записи, от которой запускается скрипт, есть права на:
   - Чтение конфигурации Exchange
   - Создание/изменение файлов в папке `C:\Scripts\db\`

## Параметры

- `Action` - Действие для выполнения (по умолчанию: "discovery")
- `DatabaseName` - Имя базы данных (требуется для некоторых действий)
- `CacheLifetime` - Время жизни кеша в минутах (по умолчанию: 30)

## Доступные действия

### discovery
Возвращает JSON для Zabbix Low Level Discovery с информацией о всех базах данных.

```powershell
.\exchange_json_manager.ps1 -Action discovery
```

**Пример вывода:**
```json
{"data":[{"{#DBNAME}":"DB01","{#DBSERVER}":"EXCH01"},{"{#DBNAME}":"DB02","{#DBSERVER}":"EXCH01"}]}
```

### mounted
Проверяет статус монтирования конкретной базы данных.

```powershell
.\exchange_json_manager.ps1 -Action mounted -DatabaseName "DB01"
```

**Возвращает:**
- `1` - база данных примонтирована
- `0` - база данных не примонтирована или не найдена

### size
Возвращает размер конкретной базы данных в байтах.

```powershell
.\exchange_json_manager.ps1 -Action size -DatabaseName "DB01"
```

**Возвращает:** размер в байтах (например: `5368709120`)

### status
Возвращает общее количество баз данных.

```powershell
.\exchange_json_manager.ps1 -Action status
```

**Возвращает:** количество баз данных (например: `3`)

### lastupdate
Возвращает возраст последнего обновления данных в минутах.

```powershell
.\exchange_json_manager.ps1 -Action lastupdate
```

**Возвращает:** 
- количество минут с последнего обновления
- `9999` - если данные отсутствуют

### fileage
Возвращает возраст JSON файла в минутах.

```powershell
.\exchange_json_manager.ps1 -Action fileage
```

**Возвращает:**
- возраст файла в минутах
- `9999` - если файл не существует

### forceupdate
Принудительно обновляет данные Exchange.

```powershell
.\exchange_json_manager.ps1 -Action forceupdate
```

**Возвращает:**
- `1` - обновление успешно
- `0` - обновление неуспешно

## Структура JSON файла

Скрипт создает файл `C:\Scripts\db\databases_info.json` со следующей структурой:

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

## Настройка Zabbix

### Создание пользовательских параметров

Добавьте в `zabbix_agentd.conf`:

```ini
# Exchange Database Discovery
UserParameter=exchange.db.discovery,powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\db\exchange_json_manager.ps1" -Action discovery

# Exchange Database Status
UserParameter=exchange.db.mounted[*],powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\db\exchange_json_manager.ps1" -Action mounted -DatabaseName "$1"

# Exchange Database Size
UserParameter=exchange.db.size[*],powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\db\exchange_json_manager.ps1" -Action size -DatabaseName "$1"

# Exchange General Status
UserParameter=exchange.db.status,powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\db\exchange_json_manager.ps1" -Action status

# Exchange Cache Age
UserParameter=exchange.cache.age,powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\db\exchange_json_manager.ps1" -Action lastupdate
```

### Создание шаблона Zabbix

1. **Discovery Rule:**
   - Key: `exchange.db.discovery`
   - Update interval: 1h

2. **Item prototypes:**
   - Mounted Status: `exchange.db.mounted[{#DBNAME}]`
   - Database Size: `exchange.db.size[{#DBNAME}]`

3. **Trigger prototypes:**
   - Database unmounted: `{Template:exchange.db.mounted[{#DBNAME}].last()}=0`
   - Database size critical: `{Template:exchange.db.size[{#DBNAME}].last()}>50000000000`

## Кеширование

Скрипт использует интеллектуальное кеширование:

- **Время жизни кеша:** 30 минут (настраивается параметром `CacheLifetime`)
- **Защита от одновременных запросов:** использует lock-файл
- **Автоматическое обновление:** данные обновляются при первом запросе после истечения кеша

## Файлы

- `C:\Scripts\db\exchange_json_manager.ps1` - основной скрипт
- `C:\Scripts\db\databases_info.json` - кеш данных
- `C:\Scripts\db\update.lock` - файл блокировки (временный)

## Мониторинг и диагностика

### Проверка работы скрипта

```powershell
# Проверка discovery
.\exchange_json_manager.ps1 -Action discovery

# Проверка статуса конкретной БД
.\exchange_json_manager.ps1 -Action mounted -DatabaseName "DB01"

# Принудительное обновление
.\exchange_json_manager.ps1 -Action forceupdate
```

### Возможные проблемы

1. **Нет прав на Exchange:**
   - Убедитесь, что учетная запись имеет права на чтение конфигурации Exchange

2. **Файл кеша не создается:**
   - Проверьте права на запись в папку `C:\Scripts\db\`

3. **Старые данные:**
   - Используйте `forceupdate` для принудительного обновления

## Производительность

- **Первый запрос:** может занять 5-10 секунд (сбор данных с Exchange)
- **Последующие запросы:** менее 1 секунды (чтение из кеша)
- **Рекомендуемый интервал проверки в Zabbix:** 5-10 минут

## Безопасность

- Скрипт не выводит ошибки в stdout (для совместимости с Zabbix)
- Использует минимальные привилегии для чтения данных Exchange
- Автоматически очищает lock-файлы при ошибках

## Лицензия

Скрипт распространяется свободно для использования в корпоративной среде.

## Поддержка

Для получения поддержки или сообщения об ошибках обращайтесь к администратору системы.