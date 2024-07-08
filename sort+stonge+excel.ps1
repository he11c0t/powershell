# Указываем список удаленных компьютеров
$computerNames = Get-Content -Path "C:\1\vm1.txt"

# Импортируем модуль ImportExcel для работы с файлами Excel
Import-Module -Name ImportExcel

# Определяем временной интервал (30 дней)
$startTime = (Get-Date).AddDays(-120)

# Формируем хэш-таблицу с фильтром для журнала событий
$filterHashtable = @{
    LogName = 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational'
    ID = 1149
    StartTime = $startTime
}

# Создаем массив результатов
$results = @()

# Получаем записи событий из журнала за последние 30 дней на каждом удаленном компьютере
foreach ($computerName in $computerNames) {
    Write-Output "`n--------------------------`nКомпьютер: $computerName`n--------------------------`n"
    
    if (Test-Connection $computerName -Count 1 -Quiet) {
        $events = Get-WinEvent -ComputerName $computerName -FilterHashtable $filterHashtable -MaxEvents 1000
        
        # Получаем информацию о свободном месте на диске и версии операционной системы
        $diskSpace = Get-WmiObject -ComputerName $computerName -Class Win32_LogicalDisk | Where-Object {$_.DriveType -eq 3} | Select-Object -Property DeviceID, @{Name='FreeSpaceGB';Expression={[math]::Round(($_.FreeSpace/1gb),2)}} 
        $osVersion = Get-WmiObject -ComputerName $computerName -Class Win32_OperatingSystem | Select-Object -Property Caption, Version

        # Обрабатываем каждую запись и добавляем ее в массив результатов
        foreach ($event in $events) {
            $eventObject = [PSCustomObject]@{
                ComputerName = $computerName
                TimeCreated = $event.TimeCreated.ToString('yyyy-MM-dd')
                UserName = $event.Properties[0].Value
                FreeSpaceGB = $diskSpace.FreeSpaceGB
                OSVersion = $osVersion.Caption + ' (' + $osVersion.Version + ')'
                #ClientIPAddress = $event.Properties[1].Value
            }

            $existingResultObject = $results | Where-Object { $_.ComputerName -eq $computerName }
            if ($existingResultObject) {
                if ($eventObject.TimeCreated -gt $existingResultObject.TimeCreated) {
                    $results.Remove($existingResultObject)
                    $results += $eventObject
                }
            } else {
                $results += $eventObject
            }
        }
    } else {
        Write-Output "Компьютер $computerName недоступен."
        $unavailableComputerObject = [PSCustomObject]@{
            ComputerName = $computerName
            TimeCreated = $null
            UserName = "недоступен"
            FreeSpaceGB = $null
            OSVersion = $null
            #ClientIPAddress = $null
        }
        $results += $unavailableComputerObject
    }
}

# Сортируем массив результатов по номеру компьютера и времени создания
$results = $results | Sort-Object ComputerName, TimeCreated

# Создаем новый файл Excel и записываем в него результаты
$resultsFilePath = "C:\1\results.xlsx"
$results | Select-Object -Property ComputerName, TimeCreated, UserName, FreeSpaceGB, OSVersion | Export-Excel -Path $resultsFilePath -AutoSize -FreezeTopRow -BoldTopRow -WorksheetName "Results"

# Выводим сообщение об успешном завершении скрипта
Write-Output "`nВыгрузка завершена. Результаты сохранены в файле: $resultsFilePath`n"
