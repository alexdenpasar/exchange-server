# Скрипт мониторинга активных миграций Exchange
# Автор: PowerShell Migration Monitor
# Версия: 1.0

# Параметры конфигурации
$DomainController = "dc.example.com"
$EmailListFile = "C:\Scripts\Migration\EmailList.txt"
$CheckInterval = 30 # секунды

# Функция для проверки существования файла со списком email
function Test-EmailListFile {
    param($FilePath)
    
    if (-not (Test-Path $FilePath)) {
        Write-Host "ОШИБКА: Файл $FilePath не найден!" -ForegroundColor Red
        Write-Host "Создайте файл $FilePath и добавьте в него email адреса (по одному на строку)" -ForegroundColor Yellow
        return $false
    }
    return $true
}

# Функция для чтения списка email из файла
function Get-EmailList {
    param($FilePath)
    
    try {
        $emails = Get-Content $FilePath -ErrorAction Stop | Where-Object { 
            $_.Trim() -ne "" -and $_.Trim() -match "^[^#].*@.*\." 
        }
        return $emails
    }
    catch {
        Write-Host "ОШИБКА при чтении файла $FilePath : $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# Функция для проверки статуса миграции пользователя
function Get-UserMigrationStatus {
    param(
        [string]$Email,
        [string]$DC
    )
    
    try {
        $moveRequest = Get-MoveRequestStatistics -DomainController $DC -Identity $Email -ErrorAction Stop
        
        if ($moveRequest) {
            return @{
                Email = $Email
                Status = $moveRequest.Status
                StatusDetail = $moveRequest.StatusDetail
                PercentComplete = $moveRequest.PercentComplete
                DisplayName = $moveRequest.DisplayName
                SourceDatabase = $moveRequest.SourceDatabase
                TargetDatabase = $moveRequest.TargetDatabase
                # Добавляем дополнительные поля для диагностики
                Message = $moveRequest.Message
                Report = $moveRequest.Report
                Found = $true
            }
        }
    }
    catch {
        # Если запрос на миграцию не найден, возвращаем статус "не найден"
        return @{
            Email = $Email
            Found = $false
            Error = $_.Exception.Message
        }
    }
}

# Функция для отображения статуса миграций
function Show-MigrationStatus {
    param($MigrationData)
    
    $activeMigrations = $MigrationData | Where-Object { $_.Found -eq $true }
    
    if ($activeMigrations.Count -eq 0) {
        return # Ничего не выводим, если нет активных миграций
    }
    
    # Очищаем экран и выводим заголовок
    Clear-Host
    Write-Host ("=" * 80) -ForegroundColor Cyan
    Write-Host "МОНИТОРИНГ АКТИВНЫХ МИГРАЦИЙ EXCHANGE" -ForegroundColor Cyan
    Write-Host "Время проверки: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')" -ForegroundColor Gray
    Write-Host ("=" * 80) -ForegroundColor Cyan
    Write-Host ""
    
    # Выводим информацию о каждой активной миграции
    foreach ($migration in $activeMigrations) {
        Write-Host "Пользователь: $($migration.DisplayName) ($($migration.Email))" -ForegroundColor White
        Write-Host "Статус: $($migration.Status)" -ForegroundColor $(
            switch ($migration.Status) {
                "InProgress" { "Yellow" }
                "Completed" { "Green" }
                "Failed" { "Red" }
                "Queued" { "Cyan" }
                default { "White" }
            }
        )
        
        # Выводим детальный статус, если он есть и не пустой
        if ($migration.StatusDetail -and $migration.StatusDetail -ne "" -and $migration.StatusDetail -ne $null) {
            $statusDetailText = $migration.StatusDetail.ToString()
            Write-Host "Детальный статус: $statusDetailText" -ForegroundColor $(
                switch -Wildcard ($statusDetailText) {
                    # Проблемные статусы - красный цвет
                    "StalledDueToSource_*" { "Red" }
                    "StalledDueToTarget_*" { "Red" }
                    "StalledDueToCI_*" { "Red" }
                    "StalledDueToHA_*" { "Red" }
                    "StalledDueToReadThrottle" { "Red" }
                    "StalledDueToWriteThrottle" { "Red" }
                    "StalledDueToReadCpu" { "Red" }
                    "StalledDueToWriteCpu" { "Red" }
                    "StalledDueToReadUnknown" { "Red" }
                    "StalledDueToWriteUnknown" { "Red" }
                    "Failed" { "Red" }
                    "FailedOther" { "Red" }
                    "Corrupted" { "Red" }
                    
                    # Предупреждающие статусы - желтый цвет
                    "StalledDueToMail_*" { "Yellow" }
                    "Suspended" { "Yellow" }
                    "AutoSuspended" { "Yellow" }
                    "Relinquished" { "Yellow" }
                    "CompletedWithWarning" { "Yellow" }
                    
                    # Активные статусы - зеленый цвет
                    "CopyingMessages" { "Green" }
                    "ScanningForMessages" { "Green" }
                    "InitialSeedingComplete" { "Green" }
                    "CompletionInProgress" { "Green" }
                    "Completing" { "Green" }
                    "Completed" { "Green" }
                    
                    # Статусы ожидания - голубой цвет
                    "Queued" { "Cyan" }
                    "InProgress" { "Cyan" }
                    "WaitingForJobPickup" { "Cyan" }
                    "CreatingFolderHierarchy" { "Cyan" }
                    "CreatingInitialSyncCheckpoint" { "Cyan" }
                    "LoadingMessages" { "Cyan" }
                    "CopyingMessagesPerUserRead" { "Cyan" }
                    "CopyingMessagesPerUserWrite" { "Cyan" }
                    
                    # Для числовых значений (как в вашем случае) - белый цвет
                    { $_ -match "^\d+$" } { "White" }
                    
                    # Остальные статусы - серый цвет
                    default { "Gray" }
                }
            )
        }
        
        # Если есть сообщение о статусе, выводим его
        if ($migration.Message -and $migration.Message -ne "" -and $migration.Message -ne $null) {
            $messageText = $migration.Message.ToString()
            if ($messageText.Length -gt 100) {
                $messageText = $messageText.Substring(0, 100) + "..."
            }
            Write-Host "Сообщение: $messageText" -ForegroundColor DarkYellow
        }
        
        Write-Host "Прогресс: $($migration.PercentComplete)%" -ForegroundColor $(
            if ($migration.PercentComplete -ge 90) { "Green" }
            elseif ($migration.PercentComplete -ge 50) { "Yellow" }
            else { "Red" }
        )
        
        if ($migration.SourceDatabase) {
            Write-Host "Источник: $($migration.SourceDatabase)" -ForegroundColor Gray
        }
        if ($migration.TargetDatabase) {
            Write-Host "Назначение: $($migration.TargetDatabase)" -ForegroundColor Gray
        }
        
        Write-Host ("-" * 60) -ForegroundColor DarkGray
        Write-Host ""
    }
    
    Write-Host "Следующая проверка через $CheckInterval секунд..." -ForegroundColor Gray
    Write-Host "Для остановки мониторинга нажмите Ctrl+C" -ForegroundColor Yellow
}

# Основная логика скрипта
function Start-MigrationMonitoring {
    Write-Host "Запуск мониторинга миграций Exchange..." -ForegroundColor Green
    Write-Host "Файл со списком email: $EmailListFile" -ForegroundColor Gray
    Write-Host "Domain Controller: $DomainController" -ForegroundColor Gray
    Write-Host "Интервал проверки: $CheckInterval секунд" -ForegroundColor Gray
    Write-Host ""
    
    # Проверяем существование файла со списком email
    if (-not (Test-EmailListFile $EmailListFile)) {
        return
    }
    
    while ($true) {
        try {
            # Читаем список email адресов
            $emailList = Get-EmailList $EmailListFile
            
            if ($emailList.Count -eq 0) {
                Write-Host "ПРЕДУПРЕЖДЕНИЕ: Список email адресов пуст или содержит только комментарии" -ForegroundColor Yellow
                Start-Sleep $CheckInterval
                continue
            }
            
            # Проверяем статус миграции для каждого пользователя
            $migrationStatuses = @()
            foreach ($email in $emailList) {
                $status = Get-UserMigrationStatus -Email $email.Trim() -DC $DomainController
                $migrationStatuses += $status
            }
            
            # Отображаем результаты
            Show-MigrationStatus $migrationStatuses
            
            # Ждем перед следующей проверкой
            Start-Sleep $CheckInterval
        }
        catch {
            Write-Host "ОШИБКА в основном цикле: $($_.Exception.Message)" -ForegroundColor Red
            Start-Sleep $CheckInterval
        }
    }
}

# Запуск мониторинга
Write-Host "PowerShell Migration Monitor v1.0" -ForegroundColor Cyan
Write-Host ""

Start-MigrationMonitoring