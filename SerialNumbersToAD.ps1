# PLEASE NOTE
# "computerModel", "computerManufacturer", "monitorSerialNumbers", and "monitorManufacturer" are custom AD attributes
# that you will need to add to your schema if you want to use this script as-is

# Get DumpEDID from here: https://www.nirsoft.net/utils/dump_edid.html

$getEDIDPath = "\\SERVER\path\DumpEDID.exe"

# Monitor info

$MonitorNames = @()
$MonitorSerials = @()

$EDIDInfo = & $getEDIDPath
$count = 1
$EDIDInfo | select-string "Serial Number\s+:" -context 2,0 | sort-object | get-unique | %{
    $info = $_ -split ">"
    $MonitorNames += "$($count):$($info[0].Split(":")[1].Trim())"
    $MonitorSerials += "$($count):$($info[1].Split(":")[1].Trim())"
    $count++
}

# Computer Info

$compInfo = (Get-CimInstance -ClassName Win32_ComputerSystem)

# Write to AD

$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$root = $domain.GetDirectoryEntry()
$searcher = [System.DirectoryServices.DirectorySearcher] $root

$searcher.Filter = "(sAMAccountName=$env:ComputerName`$)"
$searcher.PropertiesToLoad.Add("distinguishedName") > $Null

$results = $searcher.FindAll()
ForEach ($computer In $results)
{
    
    $dn = $computer.properties.Item("distinguishedName")
    $Computer = [ADSI]"LDAP://$dn"
    $Computer.serialNumber = (gwmi Win32_bios).SerialNumber
    $Computer.computerModel = $compInfo.Model
    $Computer.computerManufacturer = $compInfo.Manufacturer
    $Computer.monitorSerialNumbers = $MonitorSerials
    $Computer.monitorModels = $MonitorNames
    $Computer.SetInfo()
}
