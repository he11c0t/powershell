$computerName = ""
$isoPath = "C:\1\Windows_10_22h2.iso"
Get-Service -ComputerName $computerName -Name WinRM -ErrorAction Stop | Set-Service -StartupType Manual -Status Running
# Монтируем ISO-образ на удаленном компьютере
Invoke-Command -ComputerName $computerName -ArgumentList $isoPath -ScriptBlock {
    param($isoPath)
    $driveLetter = Mount-DiskImage -ImagePath $isoPath -PassThru | Get-Volume | Where-Object { $_.DriveType -eq "CD-ROM" } | Select-Object -ExpandProperty DriveLetter

    # Запускаем процесс установки Windows на удаленном компьютере с использованием монтированного образа
    $setupExePath = "${driveLetter}:\setup.exe"
    Start-Process -FilePath $setupExePath -ArgumentList "/auto", "upgrade", "/quiet", "/noreboot" -Wait
}