# Exchange 2016 Database Defragmentation Script
# Настройка параметров - ИЗМЕНИТЕ ПОД ВАШУ СРЕДУ

# =============================================================================
# ПАРАМЕТРЫ КОНФИГУРАЦИИ
# =============================================================================

# Имя базы данных для дефрагментации
$DatabaseName = "Name-DB"

# Путь для лог-файла
$LogPath = "C:\Scripts\Defrag\Logs\DefragLog.txt"

# Принудительное выполнение без подтверждений (True/False)
$Force = $True

# Создать папку для логов, если не существует
$LogDir = Split-Path $LogPath -Parent
if (!(Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force
}

# =============================================================================
# ОСНОВНОЙ КОД СКРИПТА
# =============================================================================

# Функция логирования
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry -ForegroundColor $(
        switch ($Level) {
            "ERROR" { "Red" }
            "WARNING" { "Yellow" }
            "SUCCESS" { "Green" }
            default { "White" }
        }
    )
    Add-Content -Path $LogPath -Value $logEntry
}

# Функция проверки свободного места
function Check-DiskSpace {
    param([string]$DatabasePath)
    
    $drive = Split-Path $DatabasePath -Qualifier
    $disk = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='$drive'"
    $freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    $dbSizeGB = [math]::Round((Get-Item $DatabasePath).Length / 1GB, 2)
    
    Write-Log "Свободное место на диске: $freeSpaceGB GB"
    Write-Log "Размер базы данных: $dbSizeGB GB"
    Write-Log "Требуется для дефрагментации: $([math]::Round($dbSizeGB * 1.1, 2)) GB"
    
    if ($freeSpaceGB -lt ($dbSizeGB * 1.1)) {
        Write-Log "ОШИБКА: Недостаточно свободного места для дефрагментации!" "ERROR"
        return $false
    }
    return $true
}

# Основной скрипт
try {
    Write-Log "========== НАЧАЛО ДЕФРАГМЕНТАЦИИ БД EXCHANGE ==========" "SUCCESS"
    Write-Log "База данных: $DatabaseName"
    Write-Log "Лог-файл: $LogPath"
    Write-Log "Принудительный режим: $Force"
    
    # Проверка Exchange Management Shell
    if (!(Get-Command Get-MailboxDatabase -ErrorAction SilentlyContinue)) {
        Write-Log "Загрузка Exchange Management Shell..." "WARNING"
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
    }
    
    # Получение информации о БД
    Write-Log "Получение информации о базе данных..."
    $database = Get-MailboxDatabase -Identity $DatabaseName -ErrorAction Stop
    $dbPath = $database.EdbFilePath
    $logFolderPath = $database.LogFolderPath
    
    Write-Log "Путь к БД: $dbPath"
    Write-Log "Путь к логам: $logFolderPath"
    
    # Проверка существования файла БД
    if (!(Test-Path $dbPath)) {
        Write-Log "ОШИБКА: Файл базы данных не найден: $dbPath" "ERROR"
        exit 1
    }
    
    # Проверка свободного места
    Write-Log "Проверка свободного места на диске..."
    if (!(Check-DiskSpace -DatabasePath $dbPath)) {
        if (!$Force) {
            Write-Log "ОШИБКА: Недостаточно места. Установите Force = True для принудительного выполнения" "ERROR"
            exit 1
        }
        Write-Log "ПРЕДУПРЕЖДЕНИЕ: Продолжение с недостаточным местом (Force режим)" "WARNING"
    }
    
    # Проверка активных подключений
    Write-Log "Проверка активных подключений..."
    try {
        $activeConnections = Get-StoreUsageStatistics -Database $DatabaseName | Where-Object {$_.TimeInServer -gt 0}
        if ($activeConnections.Count -gt 0) {
            Write-Log "ПРЕДУПРЕЖДЕНИЕ: Обнаружены активные подключения к БД: $($activeConnections.Count)" "WARNING"
            if (!$Force) {
                Write-Log "Дефрагментация отменена. Установите Force = True для принудительного выполнения" "WARNING"
                exit 1
            }
        }
    } catch {
        Write-Log "Не удалось проверить активные подключения" "WARNING"
    }
    
    # Подтверждение пользователя
    if (!$Force) {
        Write-Log "ВАЖНО: Убедитесь, что у вас есть актуальная резервная копия!" "WARNING"
        Write-Log "Дефрагментация приведет к недоступности БД на несколько часов!" "WARNING"
        $confirm = Read-Host "Продолжить дефрагментацию? (y/N)"
        if ($confirm -ne 'y' -and $confirm -ne 'Y') {
            Write-Log "Дефрагментация отменена пользователем"
            exit 0
        }
    }
    
    # Получение размера БД до дефрагментации
    $sizeBeforeGB = [math]::Round((Get-Item $dbPath).Length / 1GB, 2)
    Write-Log "Размер БД до дефрагментации: $sizeBeforeGB GB" "SUCCESS"
    
    # Остановка служб Exchange
    Write-Log "Остановка служб Exchange..."
    $services = @('MSExchangeIS', 'MSExchangeRPC', 'MSExchangeTransport')
    foreach ($service in $services) {
        try {
            $svc = Get-Service $service -ErrorAction SilentlyContinue
            if ($svc -and $svc.Status -eq 'Running') {
                Stop-Service $service -Force -ErrorAction Stop
                Write-Log "Служба $service остановлена" "SUCCESS"
            } else {
                Write-Log "Служба $service уже остановлена" "WARNING"
            }
        } catch {
            Write-Log "Предупреждение: Не удалось остановить службу $service - $($_.Exception.Message)" "WARNING"
        }
    }
    
    # Демонтирование БД
    Write-Log "Демонтирование базы данных..."
    Dismount-Database $DatabaseName -Confirm:$false -ErrorAction Stop
    Write-Log "База данных демонтирована" "SUCCESS"
    
    # Дефрагментация
    Write-Log "Начало дефрагментации. Это может занять несколько часов..." "SUCCESS"
    $startTime = Get-Date
    
    $defragCmd = "eseutil.exe"
    $defragArgs = @("/d", "`"$dbPath`"", "/o")
    
    Write-Log "Выполнение команды: $defragCmd $($defragArgs -join ' ')"
    
    $process = Start-Process -FilePath $defragCmd -ArgumentList $defragArgs -Wait -PassThru -NoNewWindow
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    if ($process.ExitCode -eq 0) {
        Write-Log "Дефрагментация завершена успешно за $($duration.TotalMinutes.ToString('F1')) минут" "SUCCESS"
    } else {
        Write-Log "ОШИБКА: Дефрагментация завершена с ошибкой. Код выхода: $($process.ExitCode)" "ERROR"
        throw "Ошибка дефрагментации"
    }
    
    # Монтирование БД
    Write-Log "Монтирование базы данных..."
    Mount-Database $DatabaseName -ErrorAction Stop
    Write-Log "База данных смонтирована" "SUCCESS"
    
    # Запуск служб Exchange
    Write-Log "Запуск служб Exchange..."
    foreach ($service in $services) {
        try {
            Start-Service $service -ErrorAction Stop
            Write-Log "Служба $service запущена" "SUCCESS"
        } catch {
            Write-Log "Предупреждение: Не удалось запустить службу $service - $($_.Exception.Message)" "WARNING"
        }
    }
    
    # Ожидание запуска служб
    Write-Log "Ожидание полного запуска служб..."
    Start-Sleep -Seconds 30
    
    # Проверка целостности БД
    Write-Log "Проверка целостности базы данных..."
    try {
        $checkCmd = "eseutil.exe"
        $checkArgs = @("/mh", "`"$dbPath`"")
        $checkResult = & $checkCmd $checkArgs
        
        if ($checkResult -match "State: Clean Shutdown") {
            Write-Log "Проверка целостности: УСПЕШНО (Clean Shutdown)" "SUCCESS"
        } else {
            Write-Log "ПРЕДУПРЕЖДЕНИЕ: Проверьте состояние БД вручную" "WARNING"
        }
    } catch {
        Write-Log "Не удалось выполнить проверку целостности" "WARNING"
    }
    
    # Получение размера БД после дефрагментации
    Start-Sleep -Seconds 10
    $sizeAfterGB = [math]::Round((Get-Item $dbPath).Length / 1GB, 2)
    $savedSpaceGB = [math]::Round($sizeBeforeGB - $sizeAfterGB, 2)
    $savedPercent = [math]::Round(($savedSpaceGB / $sizeBeforeGB) * 100, 1)
    
    Write-Log "Размер БД после дефрагментации: $sizeAfterGB GB" "SUCCESS"
    Write-Log "Освобождено места: $savedSpaceGB GB ($savedPercent%)" "SUCCESS"
    
    # Проверка почтовых ящиков
    Write-Log "Проверка почтовых ящиков..."
    try {
        $mailboxCount = (Get-MailboxStatistics -Database $DatabaseName).Count
        Write-Log "Количество почтовых ящиков: $mailboxCount" "SUCCESS"
    } catch {
        Write-Log "Не удалось получить статистику почтовых ящиков" "WARNING"
    }
    
    # Итоговая информация
    Write-Log "========== ДЕФРАГМЕНТАЦИЯ ЗАВЕРШЕНА УСПЕШНО ==========" "SUCCESS"
    Write-Log "Общее время выполнения: $($duration.TotalMinutes.ToString('F1')) минут" "SUCCESS"
    Write-Log "Результат: Освобождено $savedSpaceGB GB ($savedPercent%)" "SUCCESS"
    
} catch {
    Write-Log "КРИТИЧЕСКАЯ ОШИБКА: $($_.Exception.Message)" "ERROR"
    Write-Log "Попытка восстановления служб..." "WARNING"
    
    # Попытка восстановления
    try {
        Mount-Database $DatabaseName -ErrorAction SilentlyContinue
        Start-Service MSExchangeIS -ErrorAction SilentlyContinue
        Start-Service MSExchangeRPC -ErrorAction SilentlyContinue
        Start-Service MSExchangeTransport -ErrorAction SilentlyContinue
        Write-Log "Службы восстановлены" "SUCCESS"
    } catch {
        Write-Log "Не удалось автоматически восстановить службы. Требуется ручное вмешательство!" "ERROR"
    }
    
    exit 1
}

Write-Log "Проверьте лог-файл для подробной информации: $LogPath"