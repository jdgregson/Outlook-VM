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

# SIG # Begin signature block
# MIIGPgYJKoZIhvcNAQcCoIIGLzCCBisCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUye68rhfxFzbtll1woQOYCSLe
# Qt6gggPHMIIDwzCCAqugAwIBAgIGAW03AMu+MA0GCSqGSIb3DQEBCwUAMFwxCzAJ
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
# SIb3DQEJBDEWBBT/y/+zKKzQcWGaQ40jRgi/PxWbmTANBgkqhkiG9w0BAQEFAASC
# AQCTT8wB+WzydZltCywubCZv7kgzjaL1wZMnNBccQy566aqfueO2kyYNmIRwEEEg
# 7pR/fYHjMexo1CqY+ohARYgqnhKVFTHZ/C9GP5LFEee/TX7g9m9ymVyH6PcB13U/
# dSuup2PXxA//O/rR5XubBmuhM2GskngDuvJIDLf8diJL3251GnRBx3dBrlD1v84w
# Oa/U5jU3Wy2hIndzvf6pfNxUa5AVdN9K6kfxqbtAf7l8gkRihHgtICCqEgZ9x5O4
# fLny2y8oT5YwcboUn7JYsQVb/JFGGgUZu0RSXeRdpENWPoFAn8/LQYUdHICCdQ/w
# 8Emy5uJwK5UniC1zc4krCnRS
# SIG # End signature block
