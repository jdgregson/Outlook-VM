$vmName = "JDG-OUTLOOK-VM"
$vm = Get-VM $vmName
$ipFile = "C:\tmp\outlook-ip"

if (-not($vm)) {
    Write-Warning "`nCould not get VM `"$vmName`""
    Wait-AnyKey
    exit
}

if ($vm.state -ne "Running") {
    Write-Host "Starting VM `"$vmName`" ... " -NoNewline
    $vm | Start-VM
    if ($vm.state -ne "Running") {
        Write-Warning "`nFailed to start VM `"$vmName`""
        Wait-AnyKey
        exit
    }
    Write-Host "Done" -ForegroundColor "DarkGreen"
}

$ipRetryCount = 15
$vmIp = $null
"" | Set-Content $ipFile -Encoding "UTF8"
Write-Host "Waiting for an IP address ... " -NoNewline
while ($ipRetryCount -ne 0) {
    if (($vm | Select -ExpandProperty NetworkAdapters).IPAddresses.count -gt 0) {
        $_vmIps = ($vm | Select -ExpandProperty NetworkAdapters).IPAddresses
        $_vmIp = $_vmIps | Where-Object {$_ -notmatch ":"}
        if ($_vmIp -and $_vmIp -notmatch "169.254") {
            $vmIp = $_vmIp
            Write-Host "$vmIp" -ForegroundColor "DarkGreen"
            $vmIp | Set-Content $ipFile -Encoding "UTF8"
            break
        }
    }
    sleep 5
    $ipRetryCount--
}
if ($vmIp -eq $null) {
    Write-Warning "`nFailed to get the VM's IP address"
    Wait-AnyKey
    exit
}
