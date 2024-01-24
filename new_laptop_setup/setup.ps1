# Prompt for computer name change
$newComputerName = Read-Host -Prompt "Enter the new computer name"
$localUsername = ""
$localPassword = ""
$adminUsername = ""
$adminPin = ""
$wifiName = ""
$wifiKey = ""


# Change computer name
Rename-Computer -NewName $newComputerName -Force -Restart

# 1. Create a local user account with no administrator privileges
New-LocalUser -Name $localUsername -Password (ConvertTo-SecureString $localPassword -AsPlainText) -Description "Standard User"

# 2. Create an administrator account with the name "Makom Ligdol" and PIN number 4279
$securePin = ConvertTo-SecureString -String $adminPin -AsPlainText -Force
New-LocalUser -Name $adminUsername -Password $securePin -Description "Administrator" -AccountNeverExpires -UserMayNotChangePassword
Add-LocalGroupMember -Group "Administrators" -Member $adminUsername

# 3. Connect to a WiFi network
netsh wlan connect name=$wifiName key=$wifiKey

# 4. Install Microsoft Office (Word, Excel, PowerPoint)
# Replace 'YourUsername' and 'YourPassword' with the actual username and password
Start-Process -FilePath "setup.exe" -ArgumentList "/configure OfficeConfiguration.xml" -Wait -PassThru

# 5. Install Google Chrome
Invoke-WebRequest -Uri "https://dl.google.com/chrome/install/latest/chrome_installer.exe" -OutFile "$env:TEMP\chrome_installer.exe"
Start-Process -FilePath "$env:TEMP\chrome_installer.exe" -Wait

# 6. Install VLC Player
Invoke-WebRequest -Uri "https://get.videolan.org/vlc/3.0.11/win64/vlc-3.0.11-win64.exe" -OutFile "$env:TEMP\vlc_installer.exe"
Start-Process -FilePath "$env:TEMP\vlc_installer.exe" -Wait

# 7. Install Zoom
Invoke-WebRequest -Uri "https://zoom.us/client/latest/ZoomInstallerFull.msi" -OutFile "$env:TEMP\zoom_installer.msi"
Start-Process -FilePath "msiexec.exe" -ArgumentList "/i $env:TEMP\zoom_installer.msi /qn" -Wait

# 8. Run Windows Update with service packs and enable automatic updates
Install-WindowsUpdate -AcceptAll -AutoReboot

# 9. Create shortcuts for installed apps
$desktopPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'Shortcuts')
New-Item -ItemType Directory -Path $desktopPath -Force | Out-Null

$shortcuts = @(
    ("$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE", "Word.lnk"),
    ("$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE", "Excel.lnk"),
    ("$env:ProgramFiles\Microsoft Office\root\Office16\POWERPNT.EXE", "PowerPoint.lnk"),
    ("$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "Google Chrome.lnk"),
    ("$env:ProgramFiles\VideoLAN\VLC\vlc.exe", "VLC Media Player.lnk"),
    ("$env:ProgramFiles\Zoom\bin\Zoom.exe", "Zoom.lnk")
)

foreach ($shortcut in $shortcuts) {
    $shell = New-Object -ComObject WScript.Shell
    $shortcutPath = [System.IO.Path]::Combine($desktopPath, $shortcut[1])
    $shortcutObject = $shell.CreateShortcut($shortcutPath)
    $shortcutObject.TargetPath = $shortcut[0]
    $shortcutObject.Save()
}

Write-Host "Script execution completed successfully."
