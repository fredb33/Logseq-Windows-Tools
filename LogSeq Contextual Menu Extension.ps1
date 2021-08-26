

$NewShellKey10 = "Registry::HKEY_CLASSES_ROOT\*\shell\CopyForLogseq"

# Clean everything
Remove-Item -Force -Recurse -LiteralPath $NewShellKey10
New-Item -Path $NewShellKey10 -Value "Copy Path 4 Logseq"
Set-ItemProperty `
    -LiteralPath $NewShellKey10 `
    -Name "Icon" `
    -Type ExpandString `
    -Value "%systemroot%\system32\WindowsPowerShell\v1.0\powershell.exe,0"

# Get-ItemProperty -LiteralPath $NewShellKey20

# The command Key

$NewShellKey20 = $NewShellKey10 + "\command"

Remove-Item -LiteralPath $NewShellKey20
Get-Item -LiteralPath $NewShellKey20
New-Item -Path $NewShellKey20

# Final Value to Set: C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe -NoExit \"file://\" + \"%1\".replace(\"\\\", \"/\") | Set-Clipboard
Set-ItemProperty `
    -LiteralPath $NewShellKey20 `
    -Name ‘(Default)’ `
    -Type ExpandString `
    -Value "%systemroot%\system32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -WindowStyle Hidden \`"file://\`" + \`"%1\`".replace(\`"\\\`", \`"/\`") | Set-Clipboard"

