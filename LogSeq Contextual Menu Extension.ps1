#
#
# References:
#    https://docs.microsoft.com/en-us/windows/win32/shell/context-menu-handlers
#

# Clean everything
Remove-Item -Force -Recurse -LiteralPath "Registry::HKEY_CLASSES_ROOT\*\shell\CopyForLogseq"
Remove-Item -Force -Recurse -LiteralPath "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\Shell\logseq.geturl"
Remove-Item -Force -Recurse -LiteralPath "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\Shell\logseq.getfulllink"


# Create the main entry in the context menu
New-Item -Path "Registry::HKEY_CLASSES_ROOT\*\shell\CopyForLogseq"
Set-ItemProperty `
    -LiteralPath "Registry::HKEY_CLASSES_ROOT\*\shell\CopyForLogseq" `
    -Name "MUIVerb" `
    -Type String `
    -Value "Copy 4 Logseq"
Set-ItemProperty `
    -LiteralPath "Registry::HKEY_CLASSES_ROOT\*\shell\CopyForLogseq" `
    -Name "Icon" `
    -Type ExpandString `
    -Value "C:\Users\GVTD3145\AppData\Local\Logseq\logseq.exe,0"
Set-ItemProperty `
    -LiteralPath "Registry::HKEY_CLASSES_ROOT\*\shell\CopyForLogseq" `
    -Name "SubCommands" `
    -Type String `
    -Value "logseq.geturl;logseq.getfulllink"

# Create the logseq.geturl command
New-Item -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\Shell\logseq.geturl" -Value "Get URL"
New-Item -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\Shell\logseq.geturl\command"
Set-ItemProperty `
    -LiteralPath "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\Shell\logseq.geturl\command" `
    -Name ‘(Default)’ `
    -Type ExpandString `
    -Value "%systemroot%\system32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -WindowStyle Hidden \`"file://\`" + \`"%1\`".replace(\`"\\\`", \`"/\`") | Set-Clipboard"

# Create the logseq.getfulllink command
New-Item -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\Shell\logseq.getfulllink" -Value "Get Full Link"
New-Item -Path "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\Shell\logseq.getfulllink\command"
Set-ItemProperty `
    -LiteralPath "Registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Explorer\CommandStore\Shell\logseq.getfulllink\command" `
    -Name ‘(Default)’ `
    -Type ExpandString `
    -Value "%systemroot%\system32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -WindowStyle Hidden \`"[\`" + (Split-Path \`"%1\`" -leaf) + \`"](file://\`" + \`"%1\`".replace(\`"\\\`", \`"/\`") + \`")\`" | Set-Clipboard"








