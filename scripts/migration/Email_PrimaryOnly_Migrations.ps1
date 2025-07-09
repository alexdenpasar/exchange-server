# Скрипт для переноса почтовых ящиков в Exchange 2016
# из списка email-адресов в целевую БД с ограничением количества параллельных миграций

# Параметры, которые нужно настроить
$TargetDatabase = "DB-Archive" # Целевая база данных
$DomainController = "dc.example.com" # Домен-контроллер
$EmailListFile = "C:\Scripts\Migration\EmailList.txt" # Файл со списком email-адресов
$MaxParallelMoves = 3 # Максимальное количество параллельных перемещений
$BadItemLimit = 100 # Лимит плохих элементов при миграции
$LogFolderPath = "C:\Scripts\Migration\Logs" # Путь для логов
$LogFile = "$LogFolderPath\MailboxMigration_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$CheckInterval = 60 # Интервал проверки статуса в секундах

# Создание папки для логов, если она отсутствует
if (-not (Test-Path -Path $LogFolderPath)) {
    New-Item -ItemType Directory -Path $LogFolderPath -Force | Out-Null
}

# Функция для записи в лог
function Write-Log {
    param (
        [string]$Message
    )
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$TimeStamp - $Message" | Out-File -FilePath $LogFile -Append
    Write-Host "$TimeStamp - $Message"
}

# Функция для проверки статуса миграции
function Get-MigrationStatus {
    param (
        [string]$EmailAddress
    )
    try {
        $stats = Get-MoveRequestStatistics -Identity $EmailAddress -DomainController $DomainController -ErrorAction Stop
        
        # Проверка на процент завершения и статус
        if ($stats.PercentComplete -eq 100 -or $stats.Status -eq "Completed") {
            Write-Log "Миграция для $EmailAddress завершена (Статус: $($stats.Status), Процент: $($stats.PercentComplete)%)"
            return "Completed"
        }
        elseif ($stats.Status -eq "Failed" -or $stats.Status -eq "Error") {
            Write-Log "Миграция для $EmailAddress завершилась с ошибкой (Статус: $($stats.Status))"
            return "Failed"
        }
        else {
            # Добавляем дополнительную проверку для пограничных случаев
            if ($stats.PercentComplete -ge 99) {
                Write-Log "Миграция для $EmailAddress практически завершена ($($stats.PercentComplete)%). Ожидаем финализации..."
            }
            return $stats.Status
        }
    }
    catch [Microsoft.Exchange.Management.MoveRequestNotFoundException] {
        # Если запрос не найден, считаем что он был уже удалён
        Write-Log "Запрос на миграцию для $EmailAddress не найден. Возможно уже удалён."
        return "Removed"
    }
    catch {
        Write-Log "ОШИБКА: Не удалось получить статус миграции для $EmailAddress`: $($_.Exception.Message)"
        return "Error"
    }
}

# Функция для удаления запроса на миграцию
function Remove-Migration {
    param (
        [string]$EmailAddress
    )
    try {
        Remove-MoveRequest -Identity $EmailAddress -DomainController $DomainController -Confirm:$false -ErrorAction Stop
        Write-Log "Запрос на миграцию для $EmailAddress успешно удален"
        return $true
    }
    catch [Microsoft.Exchange.Management.MoveRequestNotFoundException] {
        Write-Log "Запрос на миграцию для $EmailAddress уже был удалён или не найден"
        return $true
    }
    catch {
        Write-Log "ОШИБКА: Не удалось удалить запрос на миграцию для $EmailAddress`: $($_.Exception.Message)"
        return $false
    }
}

Write-Log "Скрипт миграции почтовых ящиков запущен"
Write-Log "Целевая БД: $TargetDatabase, Домен-контроллер: $DomainController, Макс. параллельных миграций: $MaxParallelMoves"

# Проверка наличия файла со списком email-адресов
if (-not (Test-Path -Path $EmailListFile)) {
    Write-Log "ОШИБКА: Файл со списком email-адресов не найден по пути $EmailListFile"
    exit
}

try {
    # Чтение списка email-адресов из файла
    Write-Log "Чтение списка email-адресов из файла $EmailListFile"
    $EmailList = Get-Content -Path $EmailListFile | Where-Object { $_ -match '@' } | ForEach-Object { $_.Trim() }
    $TotalMailboxes = $EmailList.Count
    Write-Log "Всего найдено email-адресов: $TotalMailboxes"
    
    # Счетчики для отслеживания прогресса
    $Completed = 0
    $Failed = 0
    $InProgress = @{}
    
    # Проверка существующих запросов на миграцию
    Write-Log "Проверка существующих запросов на миграцию..."
    try {
        $existingRequests = Get-MoveRequest -DomainController $DomainController -ErrorAction SilentlyContinue | 
                           Where-Object { $_.TargetDatabase -eq $TargetDatabase }
        
        if ($existingRequests -and $existingRequests.Count -gt 0) {
            Write-Log "Найдено $($existingRequests.Count) существующих запросов на миграцию в БД $TargetDatabase"
            
            foreach ($request in $existingRequests) {
                $email = $request.Identity.ToString()
                $status = Get-MigrationStatus -EmailAddress $email
                
                if ($status -eq "InProgress" -or $status -eq "Queued") {
                    Write-Log "Добавление существующей активной миграции для $email в список отслеживаемых"
                    $InProgress[$email] = (Get-Date).AddMinutes(-5) # Небольшой отступ по времени для существующих миграций
                }
                elseif ($status -eq "Completed" -or $status -eq "Removed") {
                    Write-Log "Обнаружен завершенный запрос на миграцию для $email. Удаление..."
                    Remove-Migration -EmailAddress $email
                    $Completed++
                }
                elseif ($status -eq "Failed" -or $status -eq "Error") {
                    Write-Log "Обнаружен запрос на миграцию с ошибкой для $email. Удаление..."
                    Remove-Migration -EmailAddress $email
                    $Failed++
                }
            }
        }
        else {
            Write-Log "Существующих запросов на миграцию не обнаружено"
        }
    }
    catch {
        Write-Log "ПРЕДУПРЕЖДЕНИЕ: Не удалось проверить существующие запросы на миграцию: $($_.Exception.Message)"
    }
    
    # Обработка каждого email-адреса
    foreach ($Email in $EmailList) {
        # Проверка, не является ли этот ящик уже в процессе миграции
        if ($InProgress.ContainsKey($Email)) {
            Write-Log "Почтовый ящик $Email уже находится в процессе миграции. Пропуск."
            continue
        }
        
        # Если список активных миграций достиг максимума, ожидаем завершения хотя бы одной
        while ($InProgress.Count -ge $MaxParallelMoves) {
            Write-Log "Достигнут лимит параллельных миграций ($($InProgress.Count)/$MaxParallelMoves). Ожидание завершения..."
            
            # Проверка статуса каждой активной миграции
            $CompletedEmails = @()
            
            foreach ($ActiveEmail in $InProgress.Keys) {
                $Status = Get-MigrationStatus -EmailAddress $ActiveEmail
                
                # Проверка завершенных миграций
                if ($Status -eq "Completed" -or $Status -eq "Removed") {
                    Write-Log "Миграция для $ActiveEmail завершена успешно или уже удалена"
                    # Попытка удаления запроса, даже если он мог быть уже удалён
                    try {
                        Remove-MoveRequest -Identity $ActiveEmail -DomainController $DomainController -Confirm:$false -ErrorAction SilentlyContinue
                        Write-Log "Запрос на миграцию для $ActiveEmail удалён"
                    }
                    catch {
                        Write-Log "Запрос на миграцию для $ActiveEmail уже был удалён или не найден"
                    }
                    $CompletedEmails += $ActiveEmail
                    $Completed++
                }
                # Проверка неудачных миграций
                elseif ($Status -eq "Failed" -or $Status -eq "Error") {
                    Write-Log "ОШИБКА: Миграция для $ActiveEmail не удалась"
                    Remove-Migration -EmailAddress $ActiveEmail
                    $CompletedEmails += $ActiveEmail
                    $Failed++
                }
                # Проверка зависших миграций (более 24 часов)
                elseif (((Get-Date) - $InProgress[$ActiveEmail]).TotalHours -gt 24) {
                    Write-Log "ПРЕДУПРЕЖДЕНИЕ: Миграция для $ActiveEmail выполняется более 24 часов. Возможно, она зависла."
                    
                    # Получаем детальную информацию о миграции
                    try {
                        $detailedStats = Get-MoveRequestStatistics -Identity $ActiveEmail -DomainController $DomainController -IncludeReport -ErrorAction Stop
                        Write-Log "Детали зависшей миграции: Статус $($detailedStats.Status), Процент $($detailedStats.PercentComplete)%, Этап $($detailedStats.StatusDetail)"
                        
                        # Если миграция не выполнялась долгое время, пробуем перезапустить
                        if ($detailedStats.BytesTransferredPerMinute -eq 0 -or $null -eq $detailedStats.BytesTransferredPerMinute) {
                            Write-Log "Миграция для $ActiveEmail зависла (нет передачи данных). Попытка перезапуска..."
                            Remove-Migration -EmailAddress $ActiveEmail
                            $CompletedEmails += $ActiveEmail
                            # Добавляем обратно в список для обработки (будет создана заново)
                            $EmailList += $ActiveEmail
                            Write-Log "Запрос на миграцию для $ActiveEmail будет создан заново"
                        }
                    }
                    catch {
                        Write-Log "Не удалось получить детальную информацию о миграции $ActiveEmail : $($_.Exception.Message)"
                    }
                }
                # Дополнительная проверка для длительных миграций (более 3 часов)
                elseif (((Get-Date) - $InProgress[$ActiveEmail]).TotalMinutes -gt 180) {
                    # Получаем дополнительные данные о состоянии миграции
                    try {
                        $detailedStats = Get-MoveRequestStatistics -Identity $ActiveEmail -DomainController $DomainController -IncludeReport -ErrorAction Stop
                        $currentPercent = $detailedStats.PercentComplete
                        $currentStage = $detailedStats.StatusDetail
                        Write-Log "Длительная миграция $ActiveEmail : Прогресс $currentPercent%, Статус $($detailedStats.Status), Этап $currentStage"
                    }
                    catch {
                        Write-Log "Не удалось получить детальную информацию о миграции $ActiveEmail : $($_.Exception.Message)"
                    }
                }
            }
            
            # Удаление завершенных и неудачных миграций из списка активных
            foreach ($CompletedEmail in $CompletedEmails) {
                $InProgress.Remove($CompletedEmail)
            }
            
            # Если все еще достигнут максимум, ждем перед следующей проверкой
            if ($InProgress.Count -ge $MaxParallelMoves) {
                Write-Log "Активных миграций: $($InProgress.Count). Следующая проверка через $CheckInterval секунд..."
                Start-Sleep -Seconds $CheckInterval
            }
        }
        
        # Создание запроса на перемещение почтового ящика
        try {
            # Проверяем, существует ли уже запрос на миграцию для этого ящика
            $existingRequest = Get-MoveRequest -Identity $Email -DomainController $DomainController -ErrorAction SilentlyContinue
            
            if ($existingRequest) {
                Write-Log "Запрос на миграцию для $Email уже существует (Статус: $($existingRequest.Status)). Пропуск."
                
                # Если запрос уже существует и активен, добавляем его в список отслеживаемых
                if ($existingRequest.Status -eq "InProgress" -or $existingRequest.Status -eq "Queued") {
                    $InProgress[$Email] = Get-Date
                    Write-Log "Добавлен в список отслеживаемых миграций: $Email"
                }
                else {
                    # Если запрос существует, но не активен (завершен или с ошибкой), удаляем его
                    Write-Log "Удаление существующего неактивного запроса для $Email"
                    Remove-Migration -EmailAddress $Email
                    
                    # Создаем новый запрос
                    Write-Log "Создание нового запроса на миграцию для $Email"
                    New-MoveRequest -Identity $Email -DomainController $DomainController -TargetDatabase $TargetDatabase -BadItemLimit $BadItemLimit -AcceptLargeDataLoss -PrimaryOnly -ErrorAction Stop
                    Write-Log "Запрос на миграцию создан для $Email"
                    $InProgress[$Email] = Get-Date
                }
            }
            else {
                # Создаем новый запрос на миграцию
                Write-Log "Начало миграции почтового ящика $Email"
                New-MoveRequest -Identity $Email -DomainController $DomainController -TargetDatabase $TargetDatabase -BadItemLimit $BadItemLimit -AcceptLargeDataLoss -PrimaryOnly -ErrorAction Stop
                Write-Log "Запрос на миграцию создан для $Email"
                $InProgress[$Email] = Get-Date
            }
        }
        catch {
            Write-Log "ОШИБКА: Не удалось создать запрос на миграцию для $Email`: $($_.Exception.Message)"
            $Failed++
        }
        
        # Пауза между созданием новых запросов
        Start-Sleep -Seconds 2
    }
    
    # Ожидание завершения всех оставшихся миграций
    Write-Log "Все запросы на миграцию созданы. Ожидание завершения оставшихся миграций..."
    while ($InProgress.Count -gt 0) {
        Write-Log "Осталось активных миграций: $($InProgress.Count). Проверка статуса..."
        
        # Проверка статуса каждой активной миграции
        $CompletedEmails = @()
        
        foreach ($ActiveEmail in $InProgress.Keys) {
            $Status = Get-MigrationStatus -EmailAddress $ActiveEmail
            
            # Проверка завершенных миграций
            if ($Status -eq "Completed" -or $Status -eq "Removed") {
                Write-Log "Миграция для $ActiveEmail завершена успешно или уже удалена"
                Remove-Migration -EmailAddress $ActiveEmail
                $CompletedEmails += $ActiveEmail
                $Completed++
            }
            # Проверка неудачных миграций
            elseif ($Status -eq "Failed" -or $Status -eq "Error") {
                Write-Log "ОШИБКА: Миграция для $ActiveEmail не удалась"
                Remove-Migration -EmailAddress $ActiveEmail
                $CompletedEmails += $ActiveEmail
                $Failed++
            }
            # Проверка зависших миграций (более 24 часов)
            elseif (((Get-Date) - $InProgress[$ActiveEmail]).TotalHours -gt 24) {
                Write-Log "ПРЕДУПРЕЖДЕНИЕ: Миграция для $ActiveEmail выполняется более 24 часов."
                
                # Получаем детальную информацию о миграции
                try {
                    $detailedStats = Get-MoveRequestStatistics -Identity $ActiveEmail -DomainController $DomainController -IncludeReport -ErrorAction Stop
                    Write-Log "Детали длительной миграции: Статус $($detailedStats.Status), Процент $($detailedStats.PercentComplete)%, Этап $($detailedStats.StatusDetail)"
                    
                    # Если процент выполнения высокий, просто ждем
                    if ($detailedStats.PercentComplete -gt 95) {
                        Write-Log "Миграция $ActiveEmail близка к завершению ($($detailedStats.PercentComplete)%). Продолжаем ожидание."
                    }
                    # Если миграция не выполнялась долгое время, пробуем перезапустить
                    elseif ($detailedStats.BytesTransferredPerMinute -eq 0 -or $null -eq $detailedStats.BytesTransferredPerMinute) {
                        Write-Log "Миграция для $ActiveEmail зависла (нет передачи данных). Попытка перезапуска..."
                        Remove-Migration -EmailAddress $ActiveEmail
                        $CompletedEmails += $ActiveEmail
                        
                        # Пытаемся создать новый запрос миграции
                        try {
                            Start-Sleep -Seconds 10 # Даем время на завершение удаления
                            Write-Log "Перезапуск миграции для $ActiveEmail"
                            New-MoveRequest -Identity $ActiveEmail -DomainController $DomainController -TargetDatabase $TargetDatabase -BadItemLimit $BadItemLimit -AcceptLargeDataLoss -PrimaryOnly -ErrorAction Stop
                            $InProgress[$ActiveEmail] = Get-Date
                            Write-Log "Миграция для $ActiveEmail перезапущена"
                        }
                        catch {
                            Write-Log "ОШИБКА: Не удалось перезапустить миграцию для $ActiveEmail`: $($_.Exception.Message)"
                            $Failed++
                        }
                    }
                }
                catch {
                    Write-Log "Не удалось получить детальную информацию о миграции $ActiveEmail : $($_.Exception.Message)"
                }
            }
            # Дополнительная проверка для длительных миграций (более 3 часов)
            elseif (((Get-Date) - $InProgress[$ActiveEmail]).TotalMinutes -gt 180) {
                # Получаем дополнительные данные о состоянии миграции
                try {
                    $detailedStats = Get-MoveRequestStatistics -Identity $ActiveEmail -DomainController $DomainController -IncludeReport -ErrorAction Stop
                    $currentPercent = $detailedStats.PercentComplete
                    $currentStage = $detailedStats.StatusDetail
                    $bytesPerMin = if ($detailedStats.BytesTransferredPerMinute) { "$($detailedStats.BytesTransferredPerMinute / 1MB) MB/мин" } else { "0 MB/мин" }
                    Write-Log "Длительная миграция $ActiveEmail : Прогресс $currentPercent%, Статус $($detailedStats.Status), Этап $currentStage, Скорость $bytesPerMin"
                }
                catch {
                    Write-Log "Не удалось получить детальную информацию о миграции $ActiveEmail : $($_.Exception.Message)"
                }
            }
        }
        
        # Удаление завершенных и неудачных миграций из списка активных
        foreach ($CompletedEmail in $CompletedEmails) {
            $InProgress.Remove($CompletedEmail)
        }
        
        # Если остались активные миграции, ждем перед следующей проверкой
        if ($InProgress.Count -gt 0) {
            Write-Log "Активных миграций: $($InProgress.Count). Следующая проверка через $CheckInterval секунд..."
            Start-Sleep -Seconds $CheckInterval
        }
    }
    
    # Итоговая статистика
    Write-Log "Миграция завершена."
    Write-Log "Итог: Всего почтовых ящиков: $TotalMailboxes, Успешно перемещено: $Completed, Не удалось переместить: $Failed"
}
catch {
    Write-Log "КРИТИЧЕСКАЯ ОШИБКА: $($_.Exception.Message)"
    Write-Log "Трассировка стека: $($_.ScriptStackTrace)"
}

Write-Log "Скрипт миграции завершен"