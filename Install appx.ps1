$computers = Get-Content "C:\list\test.txt"
$appxPackages = Get-ChildItem -Path "C:\1\appx\"


foreach ($computer in $computers) {
    Get-Service -ComputerName $computer -Name WinRM -ErrorAction Stop | Set-Service -StartupType Manual -Status Running
    $session = New-PSSession -ComputerName $computer
    Try {
        if (-not (Test-Path "\\$computer\C$\temp")) {
            New-Item -Path "\\$computer\C$" -Name 'temp' -ItemType Directory -ErrorAction Stop | Out-Null
        }
    } Catch {
        Write-Host "Не удалось создать папку C:\temp на компьютере $computer"
        continue
    }
    foreach ($appxPackage in $appxPackages) {
        Copy-Item -Path $appxPackage.FullName -Destination "\\$computer\C$\temp\$appxPackage" -Force
        Invoke-Command -ComputerName $computer -ScriptBlock {
            Add-AppxProvisionedPackage -Online -PackagePath "C:\temp\$($using:appxPackage.Name)" -SkipLicense
        }
        del "\\$computer\C$\temp\$appxPackage"
    }
    Remove-PSSession $session
}