## ================================================
## ==                                            ==
## ==               script version 0.60          ==
## ==                                            ==
## ==                                 Moroz Oleg ==
## ==                                 16/01/2013 ==
## ==                             mod 16/01/2013 ==
## ================================================

$cDir = Get-Location
$serversFile = $cDir.Path + "\srvs.csv"
$servers = Import-Csv $serversFile

Function ConfigAppHost($hostPath) {
    Write-Host "Make Config for $hostPath"
    #$cuDir = Get-Location
    #$IISCfg = $cuDir.Path + "\applicationHost.config"
    $IISCfg = "\\" + $hostPath + "\c$\Windows\System32\inetsrv\config\applicationHost.config"
    #Write-Host $IISCfg
    $cDate = (Get-Date).ToString("yyyyMMdd-hhmms")
    $IISBkp = $IISCfg + "_$cDate"

    $itemToEnable = @(".skin", ".config", ".vb", ".resources", ".mdb", ".java", ".mdf")
    $itemsToDelete = @("bin")

    Copy-Item $IISCfg $IISBkp
    $xml = [xml](Get-Content $IISCfg -Encoding UTF8)
    ForEach ($ext in $itemToEnable) {
        $itm = $xml.configuration."system.webServer".security.requestFiltering.fileExtensions.ChildNodes | where { $_.fileExtension -eq $ext }
        if ($itm -ne $null) { $itm.allowed = "true" }
    }

    ForEach ($fldr in $itemsToDelete) {
        $itm = $xml.configuration."system.webServer".security.requestFiltering.hiddenSegments.ChildNodes | where { $_.segment -eq $fldr }
        if ($itm -ne $null) { $itm.ParentNode.RemoveChild($itm) }
    }

    $xml.Save($IISCfg)

    Write-Host "Stoping IIS service"
    & sc.exe \\$hostPath stop W3SVC
    Start-Sleep -Seconds 5
    Write-Host "Starting IIS service"
    & sc.exe \\$hostPath start W3SVC
}

ForEach ($svr in $servers) {
    Write-Host "----------"
    Write-Host $svr.Server
    $cfgPath = "\\" + $svr.Server + "\c$\Windows\System32\inetsrv\config\"
    Write-Host "Test for network path"
    If (Test-Path $cfgPath) { ConfigAppHost($svr.Server) }
    Write-Host "Done for this host!"
    Start-Sleep -Seconds 1
}

Write-Host "All Done!"
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")