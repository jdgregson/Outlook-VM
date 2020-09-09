$username = "OutlookUser"
$ipFile = "C:\tmp\outlook-ip"

Write-Host "Starting admin tasks ... " -NoNewline
Start-Process -Wait "C:\bin\Start-OutlookVm.lnk"
Write-Host "Done" -ForegroundColor "DarkGreen"
Write-Host "Reading IP address ... " -NoNewline
$vmIp = Get-Content $ipFile
if ($vmIp) {
    Write-Host "Done" -ForegroundColor "DarkGreen"
    cmdkey /generic:"$vmIp" /user:"$username" /pass:"thispasswordwasintentionallyleftblank"
    mstsc /v:$vmIp /w:1440 /h:800
    sleep 5
    cmdkey /delete:"$vmIp"
    "" | Set-Content $ipFile -Encoding "UTF8"
} else {
    Write-Warning "`nCould not find IP in `"$ipFile`""
    Wait-AnyKey
}
