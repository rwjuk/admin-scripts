$ErrorActionPreference = "SilentlyContinue"

# CONFIGURATION VARIABLES

$installationSource = "\\SERVER\installers\Capita\FMS\Core installation"
$iniSource = "\\SERVER\installers\Capita\FMS\Other files\*.ini"
$destinationPath = "C:\Program Files (x86)\SIMS\FMSSQL"
$crystalReportsSourcePath = "\\SERVER\installers\Capita\FMS\Auxiliary components\Crystal"
$crystalReportsSystemSourcePath = "\\SERVER\installers\Capita\FMS\Auxiliary components\Crystal_System"
$crystalReportsDestPath = "C:\Windows\Crystal"
$sqlncliSourcePath = If ([IntPtr]::size -eq 8) {"\\SERVER\installers\Capita\FMS\Auxiliary components\SQLNCLI_X64.MSI"} else {"\\SERVER\installers\Capita\FMS\Auxiliary components\SQLNCLI_X86.MSI"}
$bdeSourcePath = "\\SERVER\installers\Capita\FMS\Auxiliary components\Borland Database Engine\Borland Database Engine.msi"
$systemPath = If ([IntPtr]::size -eq 8) {"C:\Windows\SysWOW64"} else {"C:\Windows\System32"}
$logSource = "SIMS Deployment"
$targetVersion = (Get-Command "$installationSource\Finance.exe").FileVersionInfo.FileVersion
$installedVersion = (Get-Command "$destinationPath\Finance.exe").FileVersionInfo.FileVersion

# END CONFIGURATION

if ($targetVersion -ne $installedVersion){
    if (![System.Diagnostics.EventLog]::SourceExists($logSource)){
        new-eventlog -LogName Application -Source $logSource
    }
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 11 -Message "FMS not installed or out of date. Installed version: $installedVersion; target version: $targetVersion. Installation starting..."

    if (Test-Path ($destinationPath)){
        Remove-Item $destinationPath -recurse
        write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 12 -Message "Existing FMS installation removed"
    }

    Copy-Item "$installationSource\*" $destinationPath -recurse
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 13 -Message "New FMS files copied"

    Copy-Item $iniSource $destinationPath
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 14 -Message ".ini files copied"
    
    msiexec /i "$sqlncliSourcePath" /qn
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 16 -Message "SQL Native Client installed"
    msiexec /i "\\ad.oaklands.uk.net\shares$\installers\Capita\FMS\Auxiliary components\Borland Database Engine\Borland Database Engine.msi" /qn
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 17 -Message "Borland Database Engine installed"
    
    mkdir $crystalReportsDestPath
    Copy-Item "$crystalReportsSourcePath\*" $crystalReportsDestPath
    Copy-Item "$crystalReportsSystemSourcePath\*" $systemPath
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 18 -Message "Crystal Reports DLLs copied"
    
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 15 -Message "FMS installation complete"
} else{
    Get-Item $iniSource | %{
        $iniFilename = $_.Name
        if ( -not(Test-Path "$destinationPath\$iniFilename") ){
            Copy-Item $_.FullName $destinationPath
            write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 111 -Message "FMS ini file ($iniFilename) was missing, and has been replaced"
        }
    }
}
