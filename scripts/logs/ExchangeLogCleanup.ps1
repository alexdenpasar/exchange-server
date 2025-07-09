# Параметры
$LogRetentionDays = 7  # Срок хранения логов (в днях)
$CurrentServer = $env:COMPUTERNAME  # Имя текущего сервера
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path  # Директория скрипта
$LogFile = "$ScriptPath\ExchangeCleanup_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

# Функция логирования
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $LogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    Write-Host $LogEntry
    Add-Content -Path $LogFile -Value $LogEntry
}

Write-Log "=== Запуск очистки логов Exchange на сервере $CurrentServer ==="

# Получаем список баз, размещенных на текущем сервере
$Databases = Get-MailboxDatabase -Server $CurrentServer | Select-Object -ExpandProperty Name

if ($Databases.Count -eq 0) {
    Write-Log "На сервере $CurrentServer нет баз данных Exchange." "WARNING"
    exit
}

Write-Log "Обнаружены базы на сервере $CurrentServer : $($Databases -join ', ')"

foreach ($DB in $Databases) {
    Write-Log "`nОбработка базы: $DB"

    # Получаем путь к файлу базы
    $DBPath = (Get-MailboxDatabase -Identity $DB).EdbFilePath
    $LogPath = (Get-MailboxDatabase -Identity $DB).LogFolderPath

    if (-not (Test-Path $DBPath)) {
        Write-Log "Файл базы данных не найден: $DBPath. Пропускаем..." "WARNING"
        continue
    }

    # Отключение базы данных
    Write-Log "Отключаем базу данных: $DB"
    Dismount-Database -Identity $DB -Confirm:$false
    Start-Sleep -Seconds 5

    # Проверка статуса базы
    $DBStatus = (Get-MailboxDatabaseCopyStatus -Identity $DB).Status
    if ($DBStatus -ne "Dismounted") {
        Write-Log "Ошибка! Не удалось отключить базу $DB. Текущий статус: $DBStatus" "ERROR"
        continue
    }

    Write-Log "База $DB успешно отключена. Проверяем состояние..."

    # Проверка состояния базы через eseutil
    $EseutilOutput = eseutil /mh $DBPath 2>&1
    $CleanShutdown = $EseutilOutput -match "State:\s+Clean Shutdown"

    if (-not $CleanShutdown) {
        Write-Log "База $DB в состоянии Dirty Shutdown! Логи не удаляем!" "ERROR"
    } else {
        Write-Log "База $DB в состоянии Clean Shutdown. Удаляем старые логи..."

        if (Test-Path $LogPath) {
            $LogCount = (Get-ChildItem -Path $LogPath -Filter "*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) }).Count
            Get-ChildItem -Path $LogPath -Filter "*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } | Remove-Item -Force
            Write-Log "Удалено $LogCount логов в папке $LogPath"
        } else {
            Write-Log "Папка с логами не найдена: $LogPath" "WARNING"
        }
    }

    # Включение базы данных
    Write-Log "Запускаем базу данных: $DB"
    Mount-Database -Identity $DB
    Start-Sleep -Seconds 5

    # Проверка статуса после включения
    $DBStatus = (Get-MailboxDatabaseCopyStatus -Identity $DB).Status
    if ($DBStatus -eq "Mounted") {
        Write-Log "База $DB успешно запущена."
    } else {
        Write-Log "Ошибка! База $DB не запустилась. Требуется проверка!" "ERROR"
    }
}

Write-Log "=== Очистка логов завершена ==="
