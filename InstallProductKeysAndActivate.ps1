function Get-ActivationStatus {
[CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$DNSHostName = $Env:COMPUTERNAME
    )
    process {
        try {
            $wpa = Get-WmiObject SoftwareLicensingProduct -ComputerName $DNSHostName `
            -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f'" `
            -Property LicenseStatus -ErrorAction Stop
        } catch {
            $status = New-Object ComponentModel.Win32Exception ($_.Exception.ErrorCode)
            $wpa = $null    
        }
        $out = New-Object psobject -Property @{
            ComputerName = $DNSHostName;
            Status = [string]::Empty;
        }
        if ($wpa) {
            :outer foreach($item in $wpa) {
                switch ($item.LicenseStatus) {
                    0 {$out.Status = "Unlicensed"}
                    1 {$out.Status = "Licensed"; break outer}
                    2 {$out.Status = "Out-Of-Box Grace Period"; break outer}
                    3 {$out.Status = "Out-Of-Tolerance Grace Period"; break outer}
                    4 {$out.Status = "Non-Genuine Grace Period"; break outer}
                    5 {$out.Status = "Notification"; break outer}
                    6 {$out.Status = "Extended Grace"; break outer}
                    default {$out.Status = "Unknown value"}
                }
            }
        } else {$out.Status = $status.Message}
        $out
    }
}

#Check Windows activation status
if ((Get-ActivationStatus).Status -ne "Licensed") {

    # Install Windows 7 MAK and activate
    & cscript //B "$($env:windir)\system32\slmgr.vbs" /ipk KEY-KEY-KEY-KEY-KEY
    & cscript //B "$($env:windir)\system32\slmgr.vbs" /ato

    
}

# Check Office activation status
$officeActivationStatus = & cscript "$(${env:ProgramFiles(x86)})\Microsoft Office\Office15\OSPP.VBS" /dstatus 

if (!"$officeActivationStatus".Contains("---LICENSED---")) {
    # Install Office 2013 MAK and activate
    & cscript //B "$(${env:ProgramFiles(x86)})\Microsoft Office\Office15\OSPP.VBS" /inpkey:KEY-KEY-KEY-KEY-KEY
    & cscript //B "$(${env:ProgramFiles(x86)})\Microsoft Office\Office15\OSPP.VBS" /act
}