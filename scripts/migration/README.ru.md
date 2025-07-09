# Exchange 2016 Migration Scripts

## Языки / Languages
- [🇷🇺 Русский](README.ru.md) ← (Текущий)
- [🇺🇸 English](README.md)

Набор автоматизированных PowerShell скриптов для безопасной миграции почтовых ящиков Exchange Server 2016 с мониторингом и контролем параллельности.

## Состав пакета

### Скрипты миграции

1. **`Email_PrimaryOnly_Migrations.ps1`** - Миграция основных почтовых ящиков
2. **`Email_ArchiveOnly_Migrations.ps1`** - Миграция только архивов
3. **`Migration_Monitor.ps1`** - Мониторинг активных миграций

### Конфигурационные файлы

4. **`EmailList.txt`** - Список email-адресов для миграции

## Возможности

### Безопасность
- ✅ Контроль количества параллельных миграций
- ✅ Автоматическое обнаружение зависших миграций
- ✅ Перезапуск неудачных миграций
- ✅ Детальное логирование всех операций

### Мониторинг
- ✅ Реальное время отслеживания статуса
- ✅ Цветная индикация состояния миграций
- ✅ Отображение прогресса выполнения
- ✅ Автоматическое обновление статуса

### Устойчивость
- ✅ Обработка существующих миграций
- ✅ Восстановление после сбоев
- ✅ Защита от повторного запуска
- ✅ Автоматическая очистка завершенных запросов

## Установка

### Структура папок

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

### Настройка параметров

Отредактируйте параметры в каждом скрипте:

```powershell
# Основные параметры
$TargetDatabase = "DB-Archive"            # Целевая база данных
$DomainController = "dc.example.com"      # Домен-контроллер
$EmailListFile = "C:\Scripts\Migration\EmailList.txt"  # Файл со списком email
$MaxParallelMoves = 3                     # Максимальное количество параллельных миграций
$BadItemLimit = 100                       # Лимит плохих элементов
$CheckInterval = 60                       # Интервал проверки (секунды)
```

### Подготовка списка email

Создайте файл `EmailList.txt` с email-адресами:

```
user1@example.com
user2@example.com
user3@example.com
# Комментарии начинаются с #
# user4@example.com - временно отключен
```

## Использование

### Миграция основных почтовых ящиков

```powershell
# Запуск из Exchange Management Shell
.\Email_PrimaryOnly_Migrations.ps1
```

**Особенности:**
- Перемещает основные почтовые ящики (без архивов)
- Использует параметр `-PrimaryOnly`
- Рекомендуется для больших ящиков
- Максимум 3 параллельные миграции по умолчанию

### Миграция только архивов

```powershell
# Запуск из Exchange Management Shell
.\Email_ArchiveOnly_Migrations.ps1
```

**Особенности:**
- Перемещает только архивы почтовых ящиков
- Использует параметр `-ArchiveOnly`
- Меньше влияет на производительность
- Максимум 1 параллельная миграция по умолчанию

### Мониторинг миграций

```powershell
# Запуск мониторинга
.\Migration_Monitor.ps1
```

**Возможности мониторинга:**
- Реальное время отслеживания статуса
- Цветная индикация состояния
- Автоматическое обновление каждые 30 секунд
- Остановка по Ctrl+C

## Детали реализации

### Контроль параллельности

Скрипты контролируют количество одновременно выполняющихся миграций:

```powershell
# Если достигнут лимит, ожидаем завершения
while ($InProgress.Count -ge $MaxParallelMoves) {
    # Проверка статуса активных миграций
    # Удаление завершенных из очереди
    # Ожидание освобождения слота
}
```

### Обнаружение зависших миграций

Автоматическое обнаружение и перезапуск зависших миграций:

```powershell
# Проверка времени выполнения
if (((Get-Date) - $InProgress[$ActiveEmail]).TotalHours -gt 24) {
    # Анализ причины зависания
    # Перезапуск при необходимости
}
```

### Структура логов

Подробные логи для каждой миграции:

```
2024-01-15 10:30:00 - Скрипт миграции почтовых ящиков запущен
2024-01-15 10:30:01 - Целевая БД: DB-Archive, Макс. параллельных миграций: 3
2024-01-15 10:30:02 - Всего найдено email-адресов: 50
2024-01-15 10:30:03 - Начало миграции почтового ящика user1@example.com
2024-01-15 10:30:04 - Запрос на миграцию создан для user1@example.com
2024-01-15 11:45:30 - Миграция для user1@example.com завершена (Статус: Completed, Процент: 100%)
```

## Индикация статуса в мониторе

### Цветовая схема статусов

- 🟢 **Зеленый** - Успешное выполнение
  - `Completed`, `CopyingMessages`, `ScanningForMessages`
- 🟡 **Желтый** - Предупреждения
  - `InProgress`, `StalledDueToMail_*`, `Suspended`
- 🔴 **Красный** - Ошибки
  - `Failed`, `StalledDueToSource_*`, `StalledDueToTarget_*`
- 🔵 **Синий** - Ожидание
  - `Queued`, `WaitingForJobPickup`

### Интерпретация статусов

```
Пользователь: John Doe (john.doe@example.com)
Статус: InProgress
Детальный статус: CopyingMessages
Прогресс: 75%
Источник: DB-Old
Назначение: DB-Archive
```

## Устранение неполадок

### Зависшие миграции

**Симптомы:**
- Миграция выполняется более 24 часов
- Прогресс не изменяется длительное время
- Статус `StalledDueToSource_*` или `StalledDueToTarget_*`

**Решение:**
```powershell
# Ручная проверка миграции
Get-MoveRequestStatistics -Identity "user@example.com" -IncludeReport

# Перезапуск зависшей миграции
Remove-MoveRequest -Identity "user@example.com" -Confirm:$false
New-MoveRequest -Identity "user@example.com" -TargetDatabase "DB-Archive"
```

### Ошибки "BadItemLimit"

**Симптомы:**
- Миграция останавливается с ошибкой превышения лимита
- В логах сообщения о поврежденных элементах

**Решение:**
```powershell
# Увеличение лимита в скрипте
$BadItemLimit = 500

# Или использование AcceptLargeDataLoss
New-MoveRequest -Identity "user@example.com" -TargetDatabase "DB-Archive" -BadItemLimit 1000 -AcceptLargeDataLoss
```

### Проблемы с производительностью

**Рекомендации:**
- Уменьшите `$MaxParallelMoves` для больших ящиков
- Выполняйте миграции в нерабочее время
- Мониторьте использование дисков и сети

```powershell
# Настройка для медленных систем
$MaxParallelMoves = 1
$CheckInterval = 120  # Увеличить интервал проверки
```

## Рекомендации по использованию

### Планирование миграций

1. **Подготовка:**
   ```powershell
   # Создание резервных копий
   Get-Mailbox -Database "SourceDB" | New-MailboxExportRequest -FilePath "\\BackupServer\{0}.pst"
   
   # Проверка квот целевой БД
   Get-MailboxDatabase "TargetDB" | fl ProhibitSendQuota,ProhibitSendReceiveQuota
   ```

2. **Оптимизация:**
   ```powershell
   # Предварительная дефрагментация источника
   .\ExchangeDefrag.ps1 -DatabaseName "SourceDB"
   
   # Мониторинг производительности
   Get-Counter "\MSExchange Database(*)\Database Page Fault Stalls/sec"
   ```

### Пакетная миграция

```powershell
# Миграция по группам
$Groups = @("Sales", "Marketing", "IT")
foreach ($Group in $Groups) {
    Get-ADGroupMember $Group | ForEach-Object { $_.Mail } | Out-File "EmailList_$Group.txt"
    .\Email_PrimaryOnly_Migrations.ps1 -EmailListFile "EmailList_$Group.txt"
}
```

### Автоматизация через Task Scheduler

```powershell
# Создание задачи для ночных миграций
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -File C:\Scripts\Migration\Email_PrimaryOnly_Migrations.ps1"
$Trigger = New-ScheduledTaskTrigger -Daily -At 2:00AM
$Settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable
Register-ScheduledTask -TaskName "Exchange Migration" -Action $Action -Trigger $Trigger -Settings $Settings
```

## Мониторинг и отчетность

### Статистика миграций

```powershell
# Анализ логов миграций
$LogPath = "C:\Scripts\Migration\Logs"
$LogFiles = Get-ChildItem $LogPath -Filter "*.log"

foreach ($LogFile in $LogFiles) {
    $Content = Get-Content $LogFile.FullName
    $Successful = ($Content | Select-String "завершена успешно").Count
    $Failed = ($Content | Select-String "не удалась").Count
    
    Write-Host "Файл: $($LogFile.Name)"
    Write-Host "Успешно: $Successful, Неудачно: $Failed"
}
```

### Создание отчетов

```powershell
# Еженедельный отчет о миграциях
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

## Совместимость

- ✅ Exchange Server 2016
- ✅ Windows Server 2012 R2 / 2016 / 2019
- ✅ PowerShell 5.0+
- ⚠️ Exchange Server 2013 (требует адаптации)
- ❌ Exchange Online (используйте другие инструменты)

## Безопасность

- Скрипты требуют права Exchange Organization Management
- Логи содержат только необходимую информацию (без паролей)
- Поддержка прерывания операций через Ctrl+C
- Автоматическая очистка временных файлов

## Поддержка

Для получения помощи:

1. Проверьте лог-файлы в папке `Logs\`
2. Используйте мониторинг для диагностики
3. Обратитесь к документации Microsoft Exchange
4. Проверьте права доступа и состояние служб

## Лицензия

Скрипты предоставляются "как есть" для использования в корпоративной среде. Тестируйте в тестовой среде перед использованием в продакшене.