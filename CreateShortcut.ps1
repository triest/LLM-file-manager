$ErrorActionPreference = 'Stop'
$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = (Get-Command pythonw.exe -ErrorAction Stop).Source
$desktop = [Environment]::GetFolderPath('Desktop')
$shell = New-Object -ComObject WScript.Shell
foreach ($shortcutPath in @(
    (Join-Path $desktop 'Local File Manager.lnk'),
    (Join-Path $appDir 'Local File Manager.lnk')
)) {
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $python
    $shortcut.Arguments = '"' + (Join-Path $appDir 'app.py') + '"'
    $shortcut.WorkingDirectory = $appDir
    $shortcut.Description = 'Local file search and AI'
    $shortcut.Save()
}
Write-Host 'Shortcut created on the desktop.'
