# CONFIGURATION VARIABLES
$solusLinkedInstallations = @{"SIMS" = "\\SIMS SERVER\c$\Program Files\SIMS\SIMS .net"; "FMS" = "\\SIMS SERVER\c$\Program Files (x86)\SIMS\FMSSQL"}
$deploymentFolders = @{"SIMS" = "\\DEPLOYMENT SERVER\applications$\Capita\SIMS\Core installation"; "FMS" = "\\DEPLOYMENT SERVER\applications$\Capita\FMS\Core installation"}
$executableNames = @{"SIMS" = "Pulsar.exe"; "FMS" = "Finance.exe"}

# END CONFIGURATION

$solusLinkedInstallations.GetEnumerator() | % {
    $product = $_.key
    $ex = $executableNames[$product]
    $deployment = $deploymentFolders[$product]
    $soluslinked = $_.value
    $solusVersion = (Get-Command "$soluslinked\$ex").FileVersionInfo.FileVersion
    $deploymentVersion = (Get-Command "$deployment\$ex").FileVersionInfo.FileVersion
    if ($solusVersion -ne $deploymentVersion){
        Copy-Item "$soluslinked\*" $deployment -recurse
        Send-MailMessage -From "helpdesk@DOMAIN.tld" -To "it@DOMAIN.TLD" -Subject "$product update detected" -Body "<strong>$product</strong> update detected - files have been copied to the <a href='$deployment'>deployment server</a> from the SOLUS-updated installation on the <a href='$soluslinked'>SIMS server</a>.<br>Details are below:<br><br><strong>Old version:</strong>&nbsp;$deploymentVersion<br><strong>New version:</strong>&nbsp;$solusVersion<br><br><em>Automated update staging script (scheduled task on deployment server)</em>" -BodyAsHtml -SmtpServer "EXCHANGE SERVER"
    }
}

