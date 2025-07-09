# ���������
$LogRetentionDays = 7  # ���� �������� ����� (� ����)
$CurrentServer = $env:COMPUTERNAME  # ��� �������� �������
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path  # ���������� �������
$LogFile = "$ScriptPath\ExchangeCleanup_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

# ������� �����������
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $LogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    Write-Host $LogEntry
    Add-Content -Path $LogFile -Value $LogEntry
}

Write-Log "=== ������ ������� ����� Exchange �� ������� $CurrentServer ==="

# �������� ������ ���, ����������� �� ������� �������
$Databases = Get-MailboxDatabase -Server $CurrentServer | Select-Object -ExpandProperty Name

if ($Databases.Count -eq 0) {
    Write-Log "�� ������� $CurrentServer ��� ��� ������ Exchange." "WARNING"
    exit
}

Write-Log "���������� ���� �� ������� $CurrentServer : $($Databases -join ', ')"

foreach ($DB in $Databases) {
    Write-Log "`n��������� ����: $DB"

    # �������� ���� � ����� ����
    $DBPath = (Get-MailboxDatabase -Identity $DB).EdbFilePath
    $LogPath = (Get-MailboxDatabase -Identity $DB).LogFolderPath

    if (-not (Test-Path $DBPath)) {
        Write-Log "���� ���� ������ �� ������: $DBPath. ����������..." "WARNING"
        continue
    }

    # ���������� ���� ������
    Write-Log "��������� ���� ������: $DB"
    Dismount-Database -Identity $DB -Confirm:$false
    Start-Sleep -Seconds 5

    # �������� ������� ����
    $DBStatus = (Get-MailboxDatabaseCopyStatus -Identity $DB).Status
    if ($DBStatus -ne "Dismounted") {
        Write-Log "������! �� ������� ��������� ���� $DB. ������� ������: $DBStatus" "ERROR"
        continue
    }

    Write-Log "���� $DB ������� ���������. ��������� ���������..."

    # �������� ��������� ���� ����� eseutil
    $EseutilOutput = eseutil /mh $DBPath 2>&1
    $CleanShutdown = $EseutilOutput -match "State:\s+Clean Shutdown"

    if (-not $CleanShutdown) {
        Write-Log "���� $DB � ��������� Dirty Shutdown! ���� �� �������!" "ERROR"
    } else {
        Write-Log "���� $DB � ��������� Clean Shutdown. ������� ������ ����..."

        if (Test-Path $LogPath) {
            $LogCount = (Get-ChildItem -Path $LogPath -Filter "*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) }).Count
            Get-ChildItem -Path $LogPath -Filter "*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } | Remove-Item -Force
            Write-Log "������� $LogCount ����� � ����� $LogPath"
        } else {
            Write-Log "����� � ������ �� �������: $LogPath" "WARNING"
        }
    }

    # ��������� ���� ������
    Write-Log "��������� ���� ������: $DB"
    Mount-Database -Identity $DB
    Start-Sleep -Seconds 5

    # �������� ������� ����� ���������
    $DBStatus = (Get-MailboxDatabaseCopyStatus -Identity $DB).Status
    if ($DBStatus -eq "Mounted") {
        Write-Log "���� $DB ������� ��������."
    } else {
        Write-Log "������! ���� $DB �� �����������. ��������� ��������!" "ERROR"
    }
}

Write-Log "=== ������� ����� ��������� ==="
