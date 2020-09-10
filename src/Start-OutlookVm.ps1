Param (
    [String]$VmName = "JDG-OUTLOOK-VM",

    [String]$IPFile = "C:\tmp\outlook-ip",

    [String]$SwitchName = "vEthernet (Default Switch)"
)

function Exit-WithError {
    Param (
        [String]$Message
    )

    process {
        Write-Host "FAIL" -ForegroundColor "DarkRed"
        Write-Host "`n$Message" -ForegroundColor "DarkRed"
        Wait-AnyKey
        exit
    }
}

Set-Content -Value "" -Path $IPFile -Encoding "UTF8"
$outlookVm = Get-VM $VmName
if (-not($outlookVm)) {
    Exit-WithError -Message "Could not get VM `"$VmName`""
}

if ($outlookVm.state -ne "Running") {
    Write-Host "Starting VM `"$VmName`" ... " -NoNewline
    Start-VM -VM $outlookVm
    if ($outlookVm.state -ne "Running") {
        Exit-WithError -Message "Failed to start VM `"$VmName`""
    }
    Write-Host "OK" -ForegroundColor "DarkGreen"
    Write-Host "Waiting 30s for VM to boot ... " -NoNewline
    Start-Sleep -Seconds 30
    Write-Host "OK" -ForegroundColor "DarkGreen"
}

Write-Host "Getting switch IP address ... " -NoNewline
$targetSwitch = Get-NetAdapter | Where-Object {$_.Name -like $SwitchName}
if ($targetSwitch) {
    $switchIpAddress = Get-NetIPAddress -InterfaceIndex $targetSwitch.InterfaceIndex -AddressFamily "IPv4"
    if (-not($switchIpAddress)) {
        Exit-WithError -Message "No IPv4 address on switch"
    }
    if ($switchIpAddress.count -gt 1) {
        Exit-WithError -Message "Multiple IPv4 addresses on switch"
    }
    $switchIpAddress = $switchIpAddress.IPAddress
    $newGuestIpAddress = $switchIpAddress -replace "`\d+$", (Get-Random -Min 2 -Max 254)
    Write-Host "OK" -ForegroundColor "DarkGreen"
} else {
    Exit-WithError -Message "Switch not found: `"$SwitchName`""
}

Write-Host "Getting credentials for `"$VmName`" ... " -NoNewline
if (Get-Command "Get-PnPStoredCredential" -ErrorAction SilentlyContinue) {
    $credential = Get-PnPStoredCredential -Name "${VmName}-CREDENTIAL" -Type PSCredential
    if (-not($credential)) {
        $credential = Get-Credential -Message "Please enter admin credentials for `"$VmName`"" -ErrorAction SilentlyContinue
        if ($credential) {
            Add-PnPStoredCredential -Name "${VmName}-CREDENTIAL" -Username $credential.username -Password $credential.password
        } else {
            Exit-WithError -Message "Unable to get credentials"
        }
    }
    Write-Host "OK" -ForegroundColor "DarkGreen"
} else {
    Exit-WithError "Please install SharePointPnPPowerShellOnline:`n`n  Install-Module sharepointpnppowershellonline`n"
}

Write-Host "Connecting to `"$VmName`" ... " -NoNewline
$outlookVmSession = New-PSSession -VMName $VmName -Credential $credential
if ($outlookVmSession) {
    Write-Host "OK" -ForegroundColor "DarkGreen"
    Invoke-Command -Session $outlookVmSession -ArgumentList $switchIpAddress, $newGuestIpAddress -ScriptBlock {
        Param (
            [String]$SwitchIpAddress,

            [String]$GuestIpAddress
        )

        function Exit-WithError {
            Param (
                [String]$Message
            )

            process {
                Write-Host "`n$Message" -ForegroundColor "DarkRed"
                Write-Host 'Press any key to continue ...';
                $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
                Exit
            }
        }

        $adapters = @(Get-NetAdapter)
        if ($adapters.count -eq 0) {
            Exit-WithError "No network adapters on `"$env:computername`""
        }
        if ($adapters.count -gt 1) {
            Exit-WithError "Multiple network adapters on `"$env:computername`""
        }

        Write-Host "Setting IP Address ..." -NoNewline
        Remove-NetIPAddress -IPAddress * -InterfaceIndex $adapters.InterfaceIndex -Confirm:$False
        Remove-NetRoute -InterfaceIndex $adapters.InterfaceIndex -DestinationPrefix 0.0.0.0/0 -Confirm:$False
        New-NetIPAddress -InterfaceIndex $adapters.InterfaceIndex -IPAddress $GuestIpAddress -DefaultGateway $SwitchIpAddress -PrefixLength 24 -AddressFamily "IPv4"
        Set-NetIPAddress -InterfaceIndex $adapters.InterfaceIndex -IPAddress $GuestIpAddress

        if ((Get-NetIPAddress -InterfaceIndex $adapters.InterfaceIndex).IPAddress -eq $GuestIpAddress) {
            Write-Host "OK" -ForegroundColor "DarkGreen"
        } else {
            Write-Host "FAIL" -ForegroundColor "DarkRed"
            Exit-WithError -Message "The new IP address was not applied inside the VM"
        }
    }
} else {
    Exit-WithError -Message "Could not connect to `"$VmName`""
}

if ($outlookVmSession.State -eq "Opened") {
    Remove-PSSession $outlookVmSession
    Set-Content -Value $newGuestIpAddress -Path $IPFile -Encoding "UTF8"
} else {
    Write-Host "There was an error configuring the IP address on `"$VmName`"" -ForegroundColor "DarkRed"
    Wait-AnyKey
}
