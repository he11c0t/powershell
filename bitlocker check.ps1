do {
    $computerName = Read-Host "Введите имя компьютера"
    Get-ADObject -SearchBase (Get-AdComputer $computerName).DistinguishedName -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -Properties msfve-recoverypassword | fl name, msfve-recoverypassword
    # Остановка скрипта для просмотра результатов
    Read-Host "Нажмите Enter, чтобы продолжить"
}
while ($true)
