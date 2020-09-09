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

# SIG # Begin signature block
# MIIGPgYJKoZIhvcNAQcCoIIGLzCCBisCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUnDAhosEpRLboZyhMTuCiLIQA
# 3Y+gggPHMIIDwzCCAqugAwIBAgIGAW03AMu+MA0GCSqGSIb3DQEBCwUAMFwxCzAJ
# BgNVBAYTAlVTMRIwEAYDVQQKDAlqZGdyZWdzb24xHTAbBgNVBAsMFHNlY3VyaXR5
# IGVuZ2luZWVyaW5nMRowGAYDVQQDDBFqZGdyZWdzb24gcm9vdCBjYTAeFw0xOTA5
# MTUyMjE3NDJaFw0yOTA5MTIyMjE3MjdaMGkxCzAJBgNVBAYTAlVTMRIwEAYDVQQK
# DAlqZGdyZWdzb24xHTAbBgNVBAsMFFNlY3VyaXR5IEVuZ2luZWVyaW5nMScwJQYD
# VQQDDB5qZGdyZWdzb24gc2VjdXJpdHkgZW5naW5lZXJpbmcwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCZcnvlfZZIFomB63FyAlhMRA8jOBD7z5xFJ3F2
# Nxb9ySSsERj958Id/ub+FwbuCQgn9PLRl8x83mKnh6RjGe4fs0B7E6hJ2yUztIun
# vsbqEy7g9cxsdpzxwklHCC/l0hvyhmJa7KRd9xIUutQENOOLxNEZjrKpRKJmg752
# lZXN4s1RPUdllYgyEcO3PIyo11vtk/zOeo1dU8HmIfShMTFFkYVGgttn9I6KfJAR
# aaIaOz+ZWr84S2S2pPZLwQg/y3eouMJGFzpEr8X8T5SPMzqqHJvBS3zeo3fgV6sO
# KVlFsKL6yUDSDTWMgj4GvVV26DgGB9I9vPYyt6Kg0CSqIIAfAgMBAAGjfjB8MB8G
# A1UdIwQYMBaAFF6o+ZlGzh9Gnw8novkx4xOLXDfLMAkGA1UdEwQCMAAwHwYDVR0R
# BBgwFoEUYWRtaW5zQGpkZ3JlZ3Nvbi5jb20wDgYDVR0PAQH/BAQDAgWgMB0GA1Ud
# DgQWBBSDhvmWBAYP8OX4Ug/cx81OkyqJJjANBgkqhkiG9w0BAQsFAAOCAQEADqI5
# Znq/Azj2ieWD4ENaKU1+FZaXMaY+l2OYVCwziDam5Dn6tfz9g3fpSHLRnVsFefP1
# 4hZ7TpISUmM5752mTTMk0kSkJzSz3ywL3MRuEb2+b2v98vf8VcDZiwmVd5dLbPbj
# YrLa/v7+E2daSMvGMCfxRihlMWMYchok9YBSAO01evtAIMnFmZw1M52EITNVBI+V
# FBV3E9Z1c9j2SlHr7pX1TNS9iFqO6wOVkcfMIGzn7/vEsY29+R1BOYGojj6rATyQ
# P37K5DvNy5wPxapp8MSKfV1V4zxMN4elf3RDqclVwTtEZqYqDIIlZUigAASWSCB1
# IYaISusLCO8aUwRywTGCAeEwggHdAgEBMGYwXDELMAkGA1UEBhMCVVMxEjAQBgNV
# BAoMCWpkZ3JlZ3NvbjEdMBsGA1UECwwUc2VjdXJpdHkgZW5naW5lZXJpbmcxGjAY
# BgNVBAMMEWpkZ3JlZ3NvbiByb290IGNhAgYBbTcAy74wCQYFKw4DAhoFAKBSMBAG
# CisGAQQBgjcCAQwxAjAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMCMGCSqG
# SIb3DQEJBDEWBBTCysoDDO/f28dBB+cYPw86xOXEyTANBgkqhkiG9w0BAQEFAASC
# AQBp3OpWnkWmSJoqcsweIeo9sgwln+6xx6EC8RjgUiD1xt9M7gSriug2iWzVRSD7
# lX05J4+MzItGXx2iNxybVzrLyV7WhEx7eSlNkBIaS62Sx1gSe7BIRYQdsJ4SeVaI
# 86etri/QNF/X0bWmkmL1AzJL44xeHoYNIaAp7IM631Fhd1795x+S0dD1OWJNrw5Q
# p4CyoRSSXeP3yGoPnFwlxdArjgbBZO7GuqPyzWBv/vRSqEdejtuKLBwr1MtsDaHA
# 6QuZr8paAaei2K4AS0Qofqd+N38a46fCHWMgijC7F6mZoVDcz0PX7rXpbOF/LnMC
# TPOggdPstcGw07997+MlhsIK
# SIG # End signature block
