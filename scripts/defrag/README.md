# Exchange 2016 Database Defragmentation Script

Автоматизированный PowerShell скрипт для безопасной дефрагментации баз данных Exchange Server 2016.

## Описание

Скрипт выполняет полную дефрагментацию базы данных Exchange Server 2016 с автоматическими проверками безопасности, логированием и восстановлением служб. Предназначен для освобождения места после удаления или перемещения почтовых ящиков.

## ⚠️ Важные предупреждения

- **ОБЯЗАТЕЛЬНО** создайте резервную копию базы данных перед запуском
- Дефрагментация может занимать несколько часов
- База данных будет недоступна во время выполнения
- Требуется свободное место на диске (≥110% от размера БД)
- Выполняйте в нерабочее время

## Установка

1. Скопируйте скрипт в папку:
   ```
   C:\Scripts\Defrag\ExchangeDefrag.ps1
   ```

2. Настройте параметры в начале скрипта:
   ```powershell
   # Имя базы данных для дефрагментации
   $DatabaseName = "Name-DB"

   # Путь для лог-файла
   $LogPath = "C:\Scripts\Defrag\Logs\DefragLog.txt"

   # Принудительное выполнение без подтверждений (True/False)
   $Force = $True
   ```

## Требования

- Exchange Server 2016
- PowerShell 5.0 или выше
- Права администратора Exchange
- Права локального администратора на сервере
- Свободное место на диске ≥110% от размера БД

## Функциональность

### Автоматические проверки

- ✅ Проверка существования базы данных
- ✅ Проверка свободного места на диске
- ✅ Проверка активных подключений
- ✅ Проверка целостности БД после дефрагментации

### Безопасность

- ✅ Автоматическая остановка и запуск служб Exchange
- ✅ Демонтирование и монтирование БД
- ✅ Восстановление в случае ошибки
- ✅ Подробное логирование всех операций

### Мониторинг

- ✅ Детальное логирование с временными метками
- ✅ Цветной вывод в консоль
- ✅ Расчет освобожденного места
- ✅ Измерение времени выполнения

## Использование

### Базовый запуск

```powershell
# Запуск из Exchange Management Shell
.\ExchangeDefrag.ps1
```

### Запуск с настройками

```powershell
# Изменить параметры в скрипте перед запуском
$DatabaseName = "MyDatabase"
$LogPath = "D:\Logs\DefragLog.txt"
$Force = $False  # Включить интерактивные подтверждения
```

### Планирование через Task Scheduler

```powershell
# Создание задачи для запуска в нерабочее время
$action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument '-ExecutionPolicy Bypass -File "C:\Scripts\Defrag\ExchangeDefrag.ps1"'
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 2:00AM
$settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -WakeToRun
Register-ScheduledTask -TaskName "Exchange DB Defrag" -Action $action -Trigger $trigger -Settings $settings -User "DOMAIN\ExchangeAdmin"
```

## Процесс выполнения

1. **Инициализация**
   - Загрузка Exchange Management Shell
   - Создание папки для логов
   - Проверка параметров

2. **Предварительные проверки**
   - Проверка существования БД
   - Проверка свободного места
   - Проверка активных подключений

3. **Подготовка к дефрагментации**
   - Остановка служб Exchange
   - Демонтирование БД
   - Создание резервной копии (рекомендуется)

4. **Дефрагментация**
   - Выполнение `eseutil /d`
   - Мониторинг прогресса
   - Обработка ошибок

5. **Завершение**
   - Монтирование БД
   - Запуск служб Exchange
   - Проверка целостности
   - Расчет результатов

## Структура логов

Лог-файл содержит детальную информацию о каждом этапе:

```
[2024-01-15 02:00:00] [INFO] ========== НАЧАЛО ДЕФРАГМЕНТАЦИИ БД EXCHANGE ==========
[2024-01-15 02:00:01] [INFO] База данных: MyDatabase
[2024-01-15 02:00:02] [INFO] Размер БД до дефрагментации: 25.5 GB
[2024-01-15 02:00:03] [SUCCESS] Служба MSExchangeIS остановлена
[2024-01-15 02:00:04] [SUCCESS] База данных демонтирована
[2024-01-15 02:00:05] [SUCCESS] Начало дефрагментации...
[2024-01-15 04:30:00] [SUCCESS] Дефрагментация завершена успешно за 150.0 минут
[2024-01-15 04:30:30] [SUCCESS] Размер БД после дефрагментации: 18.2 GB
[2024-01-15 04:30:31] [SUCCESS] Освобождено места: 7.3 GB (28.6%)
```

## Коды ошибок

- **Exit Code 0**: Успешное завершение
- **Exit Code 1**: Критическая ошибка или отмена пользователем

## Устранение неполадок

### Недостаточно места на диске

```powershell
# Проверка свободного места
Get-WmiObject -Class Win32_LogicalDisk | Select-Object DeviceID, @{Name="FreeSpace(GB)";Expression={[math]::Round($_.FreeSpace/1GB,2)}}

# Очистка временных файлов
cleanmgr /sagerun:1

# Перемещение БД на другой диск (если необходимо)
Move-DatabasePath -Identity "MyDatabase" -EdbFilePath "D:\Databases\MyDatabase.edb"
```

### Службы не запускаются

```powershell
# Ручной запуск служб Exchange
Start-Service MSExchangeIS
Start-Service MSExchangeRPC
Start-Service MSExchangeTransport

# Проверка статуса служб
Get-Service | Where-Object {$_.Name -like "*Exchange*"} | Select-Object Name, Status
```

### БД не монтируется

```powershell
# Проверка целостности БД
eseutil /mh "C:\Database\MyDatabase.edb"

# Восстановление из логов (если необходимо)
eseutil /r E00 /l "C:\Database\Logs"

# Жесткое восстановление (только в крайнем случае)
eseutil /p "C:\Database\MyDatabase.edb"
```

### Активные подключения

```powershell
# Просмотр активных подключений
Get-StoreUsageStatistics -Database "MyDatabase" | Where-Object {$_.TimeInServer -gt 0}

# Принудительное отключение пользователей
Get-LogonStatistics -Database "MyDatabase" | Disable-MailboxImportRequest
```

## Рекомендации

### Подготовка к дефрагментации

1. **Создайте резервную копию**
   ```powershell
   # Экспорт всех ящиков в PST
   Get-Mailbox -Database "MyDatabase" | New-MailboxExportRequest -FilePath "\\BackupServer\Exports\{0}.pst"
   ```

2. **Уведомите пользователей**
   - Отправьте уведомление о запланированном обслуживании
   - Укажите время недоступности почты

3. **Выберите подходящее время**
   - Выходные дни
   - Нерабочие часы
   - Периоды низкой активности

### Оптимизация производительности

```powershell
# Увеличение приоритета процесса дефрагментации
# Добавьте в скрипт после запуска eseutil:
$defragProcess = Get-Process eseutil
$defragProcess.PriorityClass = "High"
```

### Мониторинг прогресса

```powershell
# Создание задачи для мониторинга размера БД
while ($true) {
    $size = (Get-Item "C:\Database\MyDatabase.edb").Length / 1GB
    Write-Host "Текущий размер БД: $([math]::Round($size, 2)) GB"
    Start-Sleep -Seconds 300  # Проверка каждые 5 минут
}
```

## Альтернативы

### Онлайн-дефрагментация

```powershell
# Автоматическая онлайн-дефрагментация (медленнее, но без простоя)
# Настройка в свойствах БД или через PowerShell:
Set-MailboxDatabase -Identity "MyDatabase" -BackgroundDatabaseMaintenance $true
```

### Перемещение ящиков

```powershell
# Альтернатива: перемещение всех ящиков в новую БД
New-MailboxDatabase -Name "NewDatabase" -EdbFilePath "D:\NewDB.edb"
Get-Mailbox -Database "MyDatabase" | New-MoveRequest -TargetDatabase "NewDatabase"
```

## Поддержка

Для получения поддержки или сообщения об ошибках:

1. Проверьте лог-файл на наличие детальной информации об ошибках
2. Убедитесь, что выполнены все требования
3. Обратитесь к документации Microsoft Exchange Server 2016

## Совместимость

- ✅ Exchange Server 2016
- ✅ Windows Server 2012 R2 / 2016 / 2019
- ✅ PowerShell 5.0+
- ❌ Exchange Online (Office 365)
- ❌ Exchange Server 2013 (требуется адаптация)

## Лицензия

Скрипт предоставляется "как есть" для использования в корпоративной среде. Используйте на свой страх и риск после тестирования в тестовой среде.