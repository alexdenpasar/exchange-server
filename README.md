# Exchange Server 2016 - Справочник команд PowerShell

Scripts and commands for working with Exchange Server 2016/2019

## Миграция

### Запрос на перемещения ящика в другую БД (вместе с архивом)
```powershell
New-MoveRequest -Identity "email@example.com" -DomainController "dc.example.com" -TargetDatabase "Name_DB" -BadItemLimit 1 -AcceptLargeDataLoss -ErrorAction Stop
```

### Перемещение только ящика в другую бд без архива
 New-MoveRequest -Identity "email@example.com" -DomainController "dc.example.com" -TargetDatabase "Name_DB" -BadItemLimit 3 -AcceptLargeDataLoss -PrimaryOnly -ErrorAction Stop

### Перемещение только архива в другую базу
```powershell
New-MoveRequest -Identity "email@example.com" -DomainController "dc.example.com" -ArchiveOnly -ArchiveTargetDatabase "Name_DB-Archive"
```

### Просмотр активных миграций
```powershell
Get-MoveRequest -DomainController "dc.example.com" -ErrorAction SilentlyContinue | Where-Object { $_.TargetDatabase -eq "Name_DB" }
```

### Просмотр статуса всех миграций
```powershell
Get-MoveRequest | Get-MoveRequestStatistics | ft DisplayName, Status, PercentComplete, TotalMailboxSize
```

### Удаление завершенных и ошибочных миграций
```powershell
Remove-MoveRequest -Identity "email@example.com" -DomainController "dc.example.com" -Confirm:$false -ErrorAction SilentlyContinue
```

### Приостановка миграции
```powershell
Suspend-MoveRequest -Identity "email@example.com" -SuspendComment "Приостановлено на техническое обслуживание"
```

### Возобновление миграции
```powershell
Resume-MoveRequest -Identity "email@example.com"
```

## Работа с БД

### Отключение (отмонтирование) БД
```powershell
Dismount-Database "Name_DB" -Confirm:$false
```

### Включение (примонтирование) БД
```powershell
Mount-Database "Name_DB" -Confirm:$false
```

### Проверка статуса БД
```cmd
eseutil /mh D:\mailbox\db.edb
```

### Проверка логов БД (нужно находиться в папке с БД)
```cmd
eseutil /ml E00
```

### Восстановление БД из логов (нужно находиться в папке с БД)
```cmd
eseutil /r E00
```

### Жесткое восстановление БД
```cmd
eseutil /p D:\mailbox\db.edb
```

### Сжатие (дефрагментация) БД после удаления/переноса почтовых ящиков
```cmd
eseutil /d D:\mailbox\db.edb
```

### Проверка БД на ошибки и восстановление
```powershell
New-MailboxRepairRequest -Database "Name_DB" -CorruptionType SearchFolder,AggregateCounts,ProvisionedFolder,FolderView
```

### Отследить прогресс операции проверки целостности БД
```powershell
Get-MailboxRepairRequest -Database "Name_DB"
```

### Просмотр ящиков в БД
```powershell
Get-MailboxStatistics -Database "Name_DB" -DomainController "dc.example.com" | ft DisplayName, TotalItemSize, ItemCount
Get-MailboxStatistics -Database "Name_DB" -DomainController "dc.example.com" | ft DisplayName, TotalItemSize
```

### Просмотр где находится сетевой архив ящика
```powershell
Get-Mailbox -Identity "email@example.com" -DomainController "dc.example.com" -Archive | Select-Object DisplayName, ArchiveDatabase
```

### Более подробно
```powershell
Get-Mailbox -Database "Name_DB" -DomainController "dc.example.com" -Archive | Select-Object DisplayName, Alias, ArchiveDatabase
```

### Просмотр всех БД и их статуса
```powershell
Get-MailboxDatabase | ft Name, Server, Mounted, DatabaseSize
```

### Создание новой БД
```powershell
New-MailboxDatabase -Name "NewDB" -EdbFilePath "D:\mailbox\NewDB.edb" -LogFolderPath "D:\logs\NewDB"
```

## Управление ящиками

### Дать полный доступ на управление ящиками
```powershell
Add-MailboxPermission -Identity "mailbox_user@domain.com" -User "your_user@domain.com" -DomainController "dc.example.com" -AccessRights FullAccess -InheritanceType All
```

### Дать разрешение отправлять от имени
```powershell
Set-Mailbox -Identity "mailbox_user@domain.com" -DomainController "dc.example.com" -GrantSendOnBehalfTo "your_user@domain.com"
```

### Разрешить отправлять как
```powershell
Add-ADPermission -Identity "mailbox_user" -User "your_user" -DomainController "dc.example.com" -ExtendedRights "Send As"
```

### Посмотреть какие разрешения есть на ящике
```powershell
Get-MailboxPermission -Identity "ceo@domain.com" -DomainController "dc.example.com" | Where-Object { $_.User -like "*alex*" }
```

### Удалить разрешения на примере FullAccess
```powershell
Remove-MailboxPermission -Identity "ceo@domain.com" -User "alex@domain.com" -DomainController "dc.example.com" -AccessRights FullAccess -InheritanceType All
```

### Создание нового ящика
```powershell
New-Mailbox -Name "John Doe" -UserPrincipalName "john.doe@domain.com" -SamAccountName "john.doe" -Database "Name_DB" -Password (ConvertTo-SecureString -String "Password123!" -AsPlainText -Force)
```

### Включение архива для ящика
```powershell
Enable-Mailbox -Identity "user@domain.com" -Archive -ArchiveDatabase "Archive_DB"
```

### Отключение архива
```powershell
Disable-Mailbox -Identity "user@domain.com" -Archive
```

### Установка квот на ящик
```powershell
Set-Mailbox -Identity "user@domain.com" -IssueWarningQuota 1GB -ProhibitSendQuota 1.2GB -ProhibitSendReceiveQuota 1.5GB
```

### Отключение ящика (сохранение в БД)
```powershell
Disable-Mailbox -Identity "user@domain.com" -Confirm:$false
```

### Удаление ящика из БД
```powershell
Remove-Mailbox -Identity "user@domain.com" -Confirm:$false
```

## Группы рассылки

### Создание группы рассылки
```powershell
New-DistributionGroup -Name "IT Team" -SamAccountName "ITTeam" -PrimarySmtpAddress "it@domain.com"
```

### Добавление пользователя в группу
```powershell
Add-DistributionGroupMember -Identity "ITTeam" -Member "user@domain.com"
```

### Удаление пользователя из группы
```powershell
Remove-DistributionGroupMember -Identity "ITTeam" -Member "user@domain.com"
```

### Просмотр членов группы
```powershell
Get-DistributionGroupMember -Identity "ITTeam" | ft Name, PrimarySmtpAddress
```

### Создание динамической группы рассылки
```powershell
New-DynamicDistributionGroup -Name "All Users" -RecipientFilter "RecipientType -eq 'UserMailbox'"
```

## Общие папки

### Создание общей папки
```powershell
New-PublicFolder -Name "Company Documents" -Path "\"
```

### Создание базы данных общих папок
```powershell
New-PublicFolderDatabase -Name "Public Folder DB" -EdbFilePath "D:\PublicFolders\PFDB.edb"
```

### Включение почты для общей папки
```powershell
Enable-MailPublicFolder -Identity "\Company Documents" -ExternalEmailAddress "docs@domain.com"
```

### Установка разрешений на общую папку
```powershell
Add-PublicFolderClientPermission -Identity "\Company Documents" -User "user@domain.com" -AccessRights Editor
```

## Транспорт и правила

### Создание правила транспорта
```powershell
New-TransportRule -Name "Block External Attachments" -FromScope NotInOrganization -AttachmentHasExecutableContent $true -RejectMessageReasonText "Executable attachments are blocked"
```

### Просмотр правил транспорта
```powershell
Get-TransportRule | ft Name, State, Priority
```

### Отключение правила
```powershell
Disable-TransportRule -Identity "Block External Attachments"
```

### Создание соединителя отправки
```powershell
New-SendConnector -Name "Internet Connector" -Usage Internet -AddressSpaces "SMTP:*;1" -SourceTransportServers "EXCH01"
```

### Просмотр очереди сообщений
```powershell
Get-Queue | ft Identity, Status, MessageCount, NextHopDomain
```

## Мониторинг и статистика

### Просмотр статистики ящиков по размеру
```powershell
Get-MailboxStatistics | Sort-Object TotalItemSize -Descending | ft DisplayName, TotalItemSize, ItemCount
```

### Просмотр топ-10 самых больших ящиков
```powershell
Get-MailboxStatistics | Sort-Object TotalItemSize -Descending | Select-Object -First 10 | ft DisplayName, TotalItemSize
```

### Проверка состояния служб Exchange
```powershell
Get-Service | Where-Object {$_.Name -like "*Exchange*"} | ft Name, Status
```

### Просмотр логов событий Exchange
```powershell
Get-EventLog -LogName Application -Source "MSExchange*" -Newest 50 | ft TimeGenerated, EntryType, Source, Message
```

### Тест подключения к ящику
```powershell
Test-MapiConnectivity -Identity "user@domain.com"
```

### Тест потока почты
```powershell
Test-MailFlow -TargetEmailAddress "test@domain.com"
```

## Резервное копирование

### Создание резервной копии БД
```powershell
New-MailboxExportRequest -Mailbox "user@domain.com" -FilePath "\\backup\exports\user.pst"
```

### Импорт из PST файла
```powershell
New-MailboxImportRequest -Mailbox "user@domain.com" -FilePath "\\backup\imports\user.pst"
```

### Просмотр статуса экспорта/импорта
```powershell
Get-MailboxExportRequest | ft Name, Status, PercentComplete
Get-MailboxImportRequest | ft Name, Status, PercentComplete
```

## Сертификаты

### Просмотр сертификатов
```powershell
Get-ExchangeCertificate | ft Thumbprint, Subject, NotAfter, Services
```

### Назначение сертификата службам
```powershell
Enable-ExchangeCertificate -Thumbprint "THUMBPRINT" -Services IIS,SMTP,POP,IMAP
```

### Создание запроса на сертификат
```powershell
New-ExchangeCertificate -GenerateRequest -SubjectName "CN=mail.domain.com" -DomainName "mail.domain.com","autodiscover.domain.com" -Path "C:\cert_request.req"
```

## Полезные команды для диагностики

### Проверка репликации БД (для DAG)
```powershell
Get-MailboxDatabaseCopyStatus | ft Name, Status, CopyQueueLength, ReplayQueueLength
```

### Просмотр активности пользователей
```powershell
Get-MailboxStatistics -Identity "user@domain.com" | Select-Object DisplayName, LastLogonTime, LastLogoffTime
```

### Поиск сообщений в ящиках
```powershell
New-MailboxSearch -Name "SearchName" -SourceMailboxes "user@domain.com" -SearchQuery "Subject:'Important Meeting'" -TargetMailbox "admin@domain.com" -TargetFolder "SearchResults"
```

### Очистка логов транспорта
```powershell
Set-TransportService -Identity "EXCH01" -MessageTrackingLogMaxAge 30.00:00:00
```

### Просмотр конфигурации виртуальных директорий
```powershell
Get-OwaVirtualDirectory | ft Name, Server, InternalUrl, ExternalUrl
Get-ActiveSyncVirtualDirectory | ft Name, Server, InternalUrl, ExternalUrl
```

---

## Примечания

- Замените `dc.example.com` на ваш контроллер домена
- Замените `Name_DB` на имя вашей базы данных
- Замените `domain.com` на ваш домен
- Всегда тестируйте команды в тестовой среде перед применением в продакшене
- Регулярно создавайте резервные копии перед выполнением критических операций

## Подключение к Exchange Management Shell

```powershell
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
```

Или используйте Exchange Management Shell напрямую из меню "Пуск".
