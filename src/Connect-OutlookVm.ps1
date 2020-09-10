Param (
    [String]$IPFile = "C:\tmp\outlook-ip"
)

Write-Host "Starting admin tasks ... " -NoNewline
Start-Process -Wait "C:\bin\Start-OutlookVm.lnk"
Write-Host "Done" -ForegroundColor "DarkGreen"
Write-Host "Reading IP address ... " -NoNewline
$outlookVmIp = Get-Content $IPFile
if ($outlookVmIp) {
    Write-Host "OK" -ForegroundColor "DarkGreen"
    mstsc /v:$outlookVmIp /w:1440 /h:800
    sleep 5
    "" | Set-Content $IPFile -Encoding "UTF8"
} else {
    Write-Host "FAIL" -ForegroundColor "DarkRed"
    Write-Warning "`nIP address not found in `"$IPFile`""
    Wait-AnyKey
}
