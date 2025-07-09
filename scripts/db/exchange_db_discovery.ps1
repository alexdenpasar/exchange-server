# Exchange JSON Manager Script
# Сохранить как: C:\Scripts\db\exchange_json_manager.ps1

param(
    [string]$Action = "discovery",
    [string]$DatabaseName = "",
    [int]$CacheLifetime = 30  # Время жизни кеша в минутах
)

$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$jsonFile = "C:\Scripts\db\databases_info.json"
$lockFile = "C:\Scripts\db\update.lock"

# Функция для обновления данных Exchange
function Update-ExchangeData {
    try {
        # Проверяем блокировку (защита от одновременных запросов)
        if (Test-Path $lockFile) {
            $lockAge = (Get-Date) - (Get-Item $lockFile).CreationTime
            if ($lockAge.TotalMinutes -lt 5) {
                # НЕ выводим предупреждение в stdout для Zabbix
                # Write-Warning "Update already in progress (lock file exists)"
                return $false
            } else {
                # Удаляем старый lock файл
                Remove-Item $lockFile -Force
            }
        }

        # Создаем lock файл
        New-Item $lockFile -ItemType File -Force | Out-Null

        # Загружаем Exchange Management Shell если не загружен
        if (!(Get-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
        }

        # Получаем данные о базах данных
        $databases = Get-MailboxDatabase -Status | ForEach-Object {
            # Извлекаем размер в байтах из строки
            $sizeBytes = 0
            if ($_.DatabaseSize) {
                $sizeMatch = [regex]::Match($_.DatabaseSize.ToString(), '\(([0-9,]+) bytes\)')
                if ($sizeMatch.Success) {
                    $sizeBytes = [long]($sizeMatch.Groups[1].Value -replace ',', '')
                }
            }

            [PSCustomObject]@{
                Name = $_.Name
                Server = @{Name = $_.Server.Name}
                Mounted = $_.Mounted
                MountedNumeric = if ($_.Mounted) { 1 } else { 0 }
                DatabaseSize = if ($_.DatabaseSize) { $_.DatabaseSize.ToString() } else { "0 GB" }
                DatabaseSizeBytes = $sizeBytes
            }
        }

        # Создаем объект с метаинформацией
        $result = @{
            LastUpdate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            LastUpdateUnix = [int][double]::Parse((Get-Date -UFormat %s))
            DatabaseCount = $databases.Count
            Databases = $databases
        }

        # Создаем директорию если не существует
        $outputDir = Split-Path $jsonFile -Parent
        if (!(Test-Path $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory -Force
        }

        # Сохраняем в JSON файл
        $jsonOutput = $result | ConvertTo-Json -Depth 3
        $jsonOutput | Out-File $jsonFile -Encoding UTF8

        # Удаляем lock файл
        Remove-Item $lockFile -Force

        return $true
        
    } catch {
        # Удаляем lock файл в случае ошибки
        if (Test-Path $lockFile) {
            Remove-Item $lockFile -Force
        }
        
        # Write-Error "Failed to update Exchange data: $($_.Exception.Message)"
        return $false
    }
}

# Функция для проверки нужно ли обновление
function Test-CacheNeedsUpdate {
    param([int]$LifetimeMinutes)
    
    if (!(Test-Path $jsonFile)) {
        return $true
    }
    
    $fileAge = (Get-Date) - (Get-Item $jsonFile).LastWriteTime
    return ($fileAge.TotalMinutes -gt $LifetimeMinutes)
}

# Функция для чтения JSON файла
function Get-DatabasesFromJson {
    try {
        if (!(Test-Path $jsonFile)) {
            return $null
        }

        $jsonContent = Get-Content $jsonFile -Raw -Encoding UTF8
        $data = $jsonContent | ConvertFrom-Json
        
        return $data
    }
    catch {
        # Write-Error "Failed to read JSON file: $($_.Exception.Message)"
        return $null
    }
}

# Основная логика
# Проверяем нужно ли обновление кеша
if (Test-CacheNeedsUpdate -LifetimeMinutes $CacheLifetime) {
    $updateResult = Update-ExchangeData
    if (!$updateResult -and !(Test-Path $jsonFile)) {
        # Если обновление не удалось и файла нет, возвращаем пустой результат
        switch ($Action.ToLower()) {
            "discovery" { Write-Host '{"data":[]}' }
            default { Write-Host "0" }
        }
        exit 1
    }
}

# Обработка различных действий
switch ($Action.ToLower()) {
    "discovery" {
        # LLD Discovery для Zabbix
        $data = Get-DatabasesFromJson
        if (!$data -or !$data.Databases) {
            Write-Host '{"data":[]}'
            return
        }
        
        $discovery = @{
            data = @()
        }
        
        foreach ($db in $data.Databases) {
            $discovery.data += @{
                "{#DBNAME}" = $db.Name
                "{#DBSERVER}" = $db.Server.Name
            }
        }
        
        # Выводим JSON для Zabbix
        ($discovery | ConvertTo-Json -Compress)
    }
    
    "mounted" {
        # Статус конкретной базы (mounted/unmounted)
        if (!$DatabaseName) {
            Write-Host "0"
            return
        }
        
        $data = Get-DatabasesFromJson
        if (!$data -or !$data.Databases) {
            Write-Host "0"
            return
        }
        
        $db = $data.Databases | Where-Object { $_.Name -eq $DatabaseName }
        
        if ($db) {
            Write-Host $db.MountedNumeric
        } else {
            Write-Host "0"
        }
    }
    
    "size" {
        # Размер конкретной базы в байтах
        if (!$DatabaseName) {
            Write-Host "0"
            return
        }
        
        $data = Get-DatabasesFromJson
        if (!$data -or !$data.Databases) {
            Write-Host "0"
            return
        }
        
        $db = $data.Databases | Where-Object { $_.Name -eq $DatabaseName }
        
        if ($db) {
            Write-Host $db.DatabaseSizeBytes
        } else {
            Write-Host "0"
        }
    }
    
    "status" {
        # Общий статус - количество баз данных
        $data = Get-DatabasesFromJson
        if (!$data) {
            Write-Host "0"
            return
        }
        
        Write-Host $data.DatabaseCount
    }
    
    "lastupdate" {
        # Unix timestamp последнего обновления
        $data = Get-DatabasesFromJson
        if (!$data) {
            Write-Host "9999"
            return
        }
        
        if ($data.LastUpdateUnix) {
            $currentUnix = [int][double]::Parse((Get-Date -UFormat %s))
            $ageMinutes = [int](($currentUnix - $data.LastUpdateUnix) / 60)
            Write-Host $ageMinutes
        } else {
            Write-Host "9999"
        }
    }
    
    "fileage" {
        # Возраст файла в минутах
        if (Test-Path $jsonFile) {
            $fileAge = (Get-Date) - (Get-Item $jsonFile).LastWriteTime
            Write-Host ([int]$fileAge.TotalMinutes)
        } else {
            Write-Host "9999"
        }
    }
    
    "forceupdate" {
        # Принудительное обновление
        $updateResult = Update-ExchangeData
        if ($updateResult) {
            Write-Host "1"
        } else {
            Write-Host "0"
        }
    }
    
    default {
        Write-Host "Available actions: discovery, mounted, size, status, lastupdate, fileage, forceupdate"
    }
}