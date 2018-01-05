$SourcePath = "\\server\share$\SystemResources\Fonts";

# http://code.kliu.org/misc/fontreg/
$FontRegExePath = "\\server\share$\SystemResources\Script utilities\FontReg.exe"


Get-ChildItem -Path $SourcePath | % {
    If (!(Test-Path "c:\windows\fonts\$($_.name)"))
    {
        copy-item $_.FullName "c:\windows\fonts"
    }
}

# FontReg.exe, when called without parameters, registers all fonts in the Windows Fonts directory
# and de-registers any that no longer exist
& $FontRegExePath