# Cria o atalho "Escanner em Lote" na Área de Trabalho, apontando para esta pasta.
$ErrorActionPreference = 'Stop'
$pythonw = (Get-Command pythonw).Source
$dir = $PSScriptRoot
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $desktop 'Escanner em Lote.lnk'))
$shortcut.TargetPath = $pythonw
$shortcut.Arguments = '"' + (Join-Path $dir 'escanner.py') + '"'
$shortcut.WorkingDirectory = $dir
$shortcut.IconLocation = "$env:SystemRoot\System32\imageres.dll,111"
$shortcut.Save()
