$ErrorActionPreference = "SilentlyContinue"

# LOGGING INITIALISATION
$logSource = "Oaklands SIMS Deployment"
if (![System.Diagnostics.EventLog]::SourceExists($logSource)){
        new-eventlog -LogName Application -Source $logSource
}

# END LOGGING

# CONFIGURATION VARIABLES

$installationSource = "\\SERVER\installers\Capita\SIMS\Core installation"
$connectIniSource = "\\SERVER\installers\Capita\SIMS\Other files\connect.ini"
$simsIniSource = "\\SERVER\installers\Capita\SIMS\Other files\sims.ini"
$destinationPath = "C:\Program Files (x86)\SIMS\SIMS .net"
$regAsmPath = "C:\Windows\MICROS~1.NET\FRAMEW~1\V40~1.303\RegAsm.exe"
$installedVersion = (Get-Command "$destinationPath\Pulsar.exe").FileVersionInfo.FileVersion

try{
    $ErrorActionPreference = "Stop"
    $targetVersion = (Get-Command "$installationSource\Pulsar.exe").FileVersionInfo.FileVersion
} catch{
    write-eventlog -LogName Application -Source $logSource -EntryType Error -EventId 900 -Message "Unable to determine target SIMS.net version - check network connectivity or existence of deployment files."
    Exit
}

$ErrorActionPreference = "SilentlyContinue"

# END CONFIGURATION

function Register-SIMSDlls(){
    Get-ChildItem $destinationPath -filter "CES*.dll" | %{ & regsvr32 /s $_.FullName }
    & $regAsmPath "$destinationPath\LoginProcesses.dll" /nologo /silent
    & $regAsmPath "$destinationPath\DocManagementProcesses.dll" /nologo /silent
    New-Item 'HKLM:\Software\SIMS .net' -Force | New-ItemProperty -Name DllsRegistered -Type "DWord" -Value 1 -Force | Out-Null
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 201 -Message "SIMS DLLs registered"
}


if ($targetVersion -ne $installedVersion){
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 1 -Message "SIMS.net not installed or out of date. Installed version: $installedVersion; target version: $targetVersion. Installation starting..."

    if (Test-Path ($destinationPath)){
        Remove-Item $destinationPath -recurse
        write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 2 -Message "Existing SIMS.net installation removed"
    }

    Copy-Item "$installationSource\*" $destinationPath -recurse
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 3 -Message "New SIMS.net files copied"

    Copy-Item $simsIniSource "C:\WINDOWS\"
    Copy-Item $connectIniSource $destinationPath
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 4 -Message "sims.ini & connect.ini file copied"
    
    Register-SIMSDlls
    
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 5 -Message "SIMS.net installation complete"
} elseif ( -not(Test-Path "$destinationPath\connect.ini") ){
    Copy-Item $connectIniSource $destinationPath
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 100 -Message "connect.ini was missing and has been replaced"
} elseif ( -not(Test-Path "C:\WINDOWS\sims.ini") ){
    Copy-Item $simsIniSource "C:\WINDOWS\"
    write-eventlog -LogName Application -Source $logSource -EntryType Information -EventId 101 -Message "sims.ini was missing and has been replaced"
}

$dllsRegistered = (Get-ItemProperty -Path "HKLM:\Software\SIMS .net" -Name DllsRegistered).DllsRegistered 2>$null  # Check DLL registration flag; 2>$null redirects error output (if the key doesn't exist) to the bit bucket
if (($dllsRegistered -eq $null) -or ($dllsRegistered -eq 0)){
    Register-SIMSDlls
}
