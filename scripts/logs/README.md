# Exchange Log Cleanup Script

PowerShell скрипт для автоматической очистки старых логов транзакций Exchange Server с проверкой состояния баз данных.

## Описание

Скрипт предназначен для безопасной очистки файлов логов транзакций Exchange Server. Он проверяет состояние каждой базы данных перед удалением логов, обеспечивая безопасность операций и предотвращая потерю данных.

## ⚠️ Важные предупреждения

- **ОБЯЗАТЕЛЬНО** создайте резервные копии перед запуском
- Скрипт временно отключает базы данных для проверки
- Логи удаляются только при состоянии "Clean Shutdown"
- Выполняйте в нерабочее время для минимизации влияния на пользователей
- Тестируйте в тестовой среде перед использованием в продакшене

## Возможности

### Безопасность
- ✅ Проверка состояния базы данных перед удалением логов
- ✅ Автоматическое отключение и включение баз данных
- ✅ Защита от удаления логов при Dirty Shutdown
- ✅ Детальное логирование всех операций

### Автоматизация
- ✅ Обработка всех баз данных на сервере
- ✅ Настраиваемый период хранения логов
- ✅ Автоматическое обнаружение путей к логам
- ✅ Цветное логирование с уровнями важности

## Установка

1. Скопируйте скрипт в папку:
   ```
   C:\Scripts\Maintenance\ExchangeLogCleanup.ps1
   ```

2. Настройте параметры в начале скрипта:
   ```powershell
   $LogRetentionDays = 7  # Срок хранения логов (в днях)
   ```

## Параметры конфигурации

- **`$LogRetentionDays`** - Количество дней для хранения логов (по умолчанию: 7)
- **`$CurrentServer`** - Имя текущего сервера (автоматически определяется)
- **`$ScriptPath`** - Директория скрипта (автоматически определяется)
- **`$LogFile`** - Путь к файлу логов скрипта (автоматически создается)

## Использование

### Базовый запуск

```powershell
# Запуск из Exchange Management Shell
.\ExchangeLogCleanup.ps1
```

### Изменение параметров

```powershell
# Изменить срок хранения логов
# Отредактируйте в скрипте:
$LogRetentionDays = 14  # Хранить логи 14 дней
```

### Планирование через Task Scheduler

```powershell
# Создание задачи для еженедельного запуска
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -File C:\Scripts\Maintenance\ExchangeLogCleanup.ps1"
$Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 3:00AM
$Settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -WakeToRun
Register-ScheduledTask -TaskName "Exchange Log Cleanup" -Action $Action -Trigger $Trigger -Settings $Settings -User "DOMAIN\ExchangeAdmin"
```

## Алгоритм работы

### 1. Инициализация
- Создание файла логов с временной меткой
- Определение текущего сервера Exchange
- Получение списка баз данных

### 2. Для каждой базы данных
1. **Отключение базы данных**
   ```powershell
   Dismount-Database -Identity $DB -Confirm:$false
   ```

2. **Проверка состояния**
   ```powershell
   eseutil /mh $DBPath  # Проверка на Clean Shutdown
   ```

3. **Удаление логов** (только при Clean Shutdown)
   ```powershell
   Get-ChildItem -Path $LogPath -Filter "*.log" | 
   Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } | 
   Remove-Item -Force
   ```

4. **Включение базы данных**
   ```powershell
   Mount-Database -Identity $DB
   ```

### 3. Завершение
- Проверка статуса всех баз данных
- Создание итогового отчета

## Структура логов

Файл логов создается в папке скрипта с именем:
```
ExchangeCleanup_YYYY-MM-DD_HH-MM-SS.log
```

### Пример лога

```
2024-01-15 03:00:00 [INFO] === Запуск очистки логов Exchange на сервере EXCH01 ===
2024-01-15 03:00:01 [INFO] Обнаружены базы на сервере EXCH01 : DB01, DB02, DB03
2024-01-15 03:00:02 [INFO] 
Обработка базы: DB01
2024-01-15 03:00:03 [INFO] Отключаем базу данных: DB01
2024-01-15 03:00:08 [INFO] База DB01 успешно отключена. Проверяем состояние...
2024-01-15 03:00:09 [INFO] База DB01 в состоянии Clean Shutdown. Удаляем старые логи...
2024-01-15 03:00:10 [INFO] Удалено 25 логов в папке D:\Logs\DB01
2024-01-15 03:00:11 [INFO] Запускаем базу данных: DB01
2024-01-15 03:00:16 [INFO] База DB01 успешно запущена.
```

### Цветовая схема

- **INFO** - Белый текст (обычная информация)
- **WARNING** - Желтый текст (предупреждения)
- **ERROR** - Красный текст (ошибки)

## Проверка состояния базы данных

### Clean Shutdown vs Dirty Shutdown

**Clean Shutdown:**
- База данных была корректно отключена
- Все транзакции зафиксированы
- Логи можно безопасно удалять

**Dirty Shutdown:**
- База данных была отключена некорректно
- Возможны незафиксированные транзакции
- Логи НЕ удаляются (необходимы для восстановления)

### Команда проверки

```cmd
eseutil /mh "D:\Database\DB01.edb"
```

**Результат для Clean Shutdown:**
```
State: Clean Shutdown
```

**Результат для Dirty Shutdown:**
```
State: Dirty Shutdown
```

## Устранение неполадок

### База данных не отключается

**Симптомы:**
- Скрипт не может отключить базу данных
- Сообщение об ошибке при выполнении Dismount-Database

**Решение:**
```powershell
# Проверка активных подключений
Get-StoreUsageStatistics -Database "DB01" | Where-Object {$_.TimeInServer -gt 0}

# Принудительное отключение
Dismount-Database -Identity "DB01" -Confirm:$false -Force
```

### Dirty Shutdown после отключения

**Симптомы:**
- База показывает Dirty Shutdown после корректного отключения
- Логи не удаляются

**Решение:**
```powershell
# Проверка целостности базы
eseutil /mh "D:\Database\DB01.edb"

# Восстановление из логов (если необходимо)
eseutil /r E00 /l "D:\Logs\DB01"

# Проверка после восстановления
eseutil /mh "D:\Database\DB01.edb"
```

### База данных не монтируется

**Симптомы:**
- Ошибка при попытке включения базы
- База остается в состоянии Dismounted

**Решение:**
```powershell
# Проверка ошибок в логах событий
Get-EventLog -LogName Application -Source "MSExchange*" -Newest 10

# Проверка целостности
eseutil /mh "D:\Database\DB01.edb"

# Принудительное монтирование
Mount-Database -Identity "DB01" -Force
```

## Рекомендации

### Планирование очистки

1. **Еженедельная очистка** (рекомендуется)
   ```powershell
   $LogRetentionDays = 7
   # Запуск каждое воскресенье в 3:00
   ```

2. **Ежемесячная очистка**
   ```powershell
   $LogRetentionDays = 30
   # Запуск первого числа каждого месяца
   ```

### Мониторинг дискового пространства

```powershell
# Проверка свободного места до и после очистки
Get-WmiObject -Class Win32_LogicalDisk | 
Select-Object DeviceID, @{Name="FreeSpace(GB)";Expression={[math]::Round($_.FreeSpace/1GB,2)}}
```

### Резервное копирование

```powershell
# Создание резервной копии логов перед удалением
$BackupPath = "\\BackupServer\Exchange\Logs\$(Get-Date -Format 'yyyyMMdd')"
New-Item -ItemType Directory -Path $BackupPath -Force
Copy-Item -Path "D:\Logs\*" -Destination $BackupPath -Recurse
```

## Автоматизация

### Интеграция с системой мониторинга

```powershell
# Отправка уведомлений о результатах
if ($ErrorCount -gt 0) {
    Send-MailMessage -To "admin@company.com" -Subject "Exchange Log Cleanup - Errors" -Body "Обнаружены ошибки при очистке логов"
}
```

### Создание отчетов

```powershell
# Анализ результатов очистки
$LogPath = "C:\Scripts\Maintenance"
$CleanupLogs = Get-ChildItem $LogPath -Filter "ExchangeCleanup_*.log"

foreach ($Log in $CleanupLogs) {
    $Content = Get-Content $Log.FullName
    $DeletedCount = ($Content | Select-String "Удалено \d+ логов").Matches.Count
    $ErrorCount = ($Content | Select-String "\[ERROR\]").Count
    
    Write-Host "Лог: $($Log.Name), Удалено: $DeletedCount, Ошибок: $ErrorCount"
}
```

## Требования

- Exchange Server 2016/2019
- PowerShell 5.0 или выше
- Права администратора Exchange
- Права локального администратора на сервере

## Совместимость

- ✅ Exchange Server 2016
- ✅ Exchange Server 2019
- ✅ Windows Server 2012 R2 / 2016 / 2019
- ⚠️ Exchange Server 2013 (требует тестирования)
- ❌ Exchange Online (не применимо)

## Безопасность

- Скрипт проверяет состояние базы перед удалением логов
- Автоматическое восстановление баз данных после операций
- Детальное логирование для аудита
- Защита от случайного удаления важных логов

## Поддержка

Для получения помощи:

1. Проверьте файл логов скрипта
2. Используйте Event Viewer для диагностики Exchange
3. Проверьте состояние баз данных через Exchange Management Console
4. Обратитесь к документации Microsoft Exchange

## Лицензия

Скрипт предоставляется "как есть" для использования в корпоративной среде. Тестируйте перед использованием в продакшене.