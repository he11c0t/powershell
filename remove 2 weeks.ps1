$computerName = ""
$outlookFolderPath = "C$\Users\*\AppData\Local\Microsoft\Outlook" # Звездочка (*) указывает на все папки пользователей
$tempFolderPath = "C:\Windows\Temp","C:\Cache\ccmcache","C:\Users\*\Downloads"
$daysToKeep = 14

# Установка службы WinRM на удаленном компьютере
Get-Service -ComputerName $computerName -Name WinRM -ErrorAction Stop | Set-Service -StartupType Manual -Status Running

# Команда удаления файлов старше указанного количества дней
$deleteCommand = {
    Param($path, $days, $extensions)
    $currentDate = Get-Date
    $limitDate = $currentDate.AddDays(-$days)

    Get-ChildItem -Path $path | Where-Object {
        $_.LastWriteTime -lt $limitDate -and ($_.Extension -notin $extensions)
    } | ForEach-Object {
        $file = $_.FullName
        Write-Host "Удаление файла: $file"
        Remove-Item -Path $file -Force -Recurse
    }
}

# Получение списка пользователей на удаленном компьютере
$userFolders = Invoke-Command -ComputerName $computerName -ScriptBlock {
    Get-ChildItem -Path "C:\Users" | Where-Object { $_.PSIsContainer } | Select-Object -ExpandProperty Name
}

# Установка сеанса WinRM на удаленном компьютере
$session = New-PSSession -ComputerName $computerName

# Удаление файлов в папках Outlook каждого пользователя на удаленном компьютере
foreach ($userFolder in $userFolders) {
    $userOutlookFolderPath = "C:\Users\$userFolder\AppData\Local\Microsoft\Outlook"
    $excludeExtensions = @(".pst")
    Invoke-Command -Session $session -ScriptBlock $deleteCommand -ArgumentList $userOutlookFolderPath, $daysToKeep, $excludeExtensions
}

# Удаление файлов из папок на удаленном компьютере
Invoke-Command -Session $session -ScriptBlock $deleteCommand -ArgumentList $tempFolderPath, $daysToKeep, @()

# Завершение сеанса WinRM
Remove-PSSession -Session $session