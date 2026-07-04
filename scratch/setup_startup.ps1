$WshShell = New-Object -ComObject WScript.Shell
$StartupFolder = [System.Environment]::GetFolderPath('Startup')
$ShortcutPath = Join-Path $StartupFolder "TokiwaenShortstaySync.lnk"
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = "C:\Users\tokiw\projects\shortstay-vacancy\silent_start.vbs"
$Shortcut.WorkingDirectory = "C:\Users\tokiw\projects\shortstay-vacancy"
$Shortcut.Save()
Write-Output "Startup shortcut created successfully at: $ShortcutPath"
