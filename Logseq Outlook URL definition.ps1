

# Clean everything
Remove-Item -Force -Recurse -LiteralPath "Registry::HKEY_CLASSES_ROOT\outlookurl"

# Create the URL handler
New-Item -Path "Registry::HKEY_CLASSES_ROOT\outlookurl" -Value "URL:outlookurl Protocol"

Set-ItemProperty `
    -LiteralPath "Registry::HKEY_CLASSES_ROOT\outlookurl" `
    -Name "URL Protocol" `
    -Type String `
    -Value ""

New-Item -Path "Registry::HKEY_CLASSES_ROOT\outlookurl\shell"

New-Item -Path "Registry::HKEY_CLASSES_ROOT\outlookurl\shell\open"

New-Item -Path "Registry::HKEY_CLASSES_ROOT\outlookurl\shell\open\command"

Set-ItemProperty `
    -LiteralPath "Registry::HKEY_CLASSES_ROOT\outlookurl\shell\open\command" `
    -Name ‘(Default)’ `
    -Type ExpandString `
    -Value "cmd.exe /V:ON /C set LINK=%1 && set OID=!LINK:~11!  && set OID=!OID:/=! && `"%ProgramFiles%\Microsoft Office\root\Office16\OUTLOOK.EXE`" /select outlook:!OID!"
