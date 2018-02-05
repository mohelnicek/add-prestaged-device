# Import modules
Import-Module ActiveDirectory


# Where WDS stores its files
$wdsRoot = "D:\RemoteInstall"

# Change settings of computer that is already prestaged or
$changeIfExists = $true


Function existsOU {
    <#
    .DESCRIPTION
        Checks if OU exists in local domain.
    .PARAMETER OU
        Expects fully qualified name of OU as parameter.
    .EXAMPLE
        existsOU("OU=existing-ou,OU=desktops,OU=Computers,DC=my,DC=domain,DC=cz")
        OU that exists - returns true
    .EXAMPLE
        existsOU("OU=unknown-ou,OU=desktops,OU=Computers,DC=my,DC=domain,DC=cz")
        Non existatnt OU - returns false
    .NOTES
        Needs to be run as domain user
    #>
    Param(
    $OU
    )
    try {
        Get-ADOrganizationalUnit $OU | Out-Null
    }
    catch {
        return $false
    }
    return $true
}


# Skip commented lines
$items = Get-Content .\add-prestaged-device.csv |  Where-Object { !$_.StartsWith("`;") } | ConvertFrom-Csv -Delimiter ";"

foreach($item in $items) {
    if([string]::IsNullOrWhitespace($($item.Name))) { # Computer name must not be empty
        Write-host "ERROR: missing computer name for prestage."
        Write-host ""
        continue;
    }

    # UUID must have format of UUID or MAC address
    if([string]::IsNullOrWhitespace($($item.UUID)) -or -not (($($item.UUID) -match "^[A-F0-9]{8}(-[A-F0-9]{4}){3}-[A-F0-9]{12}$") -or ($($item.UUID) -match "^[A-F0-9]{2}(-[A-F0-9]{2}){5}$"))) {
        Write-host "ERROR: wrong UUID or MAC pattern"
        Write-host ""
        continue;
    } else {
        # format is good, if its UUID add brackets to beginning and end of the string
        if($($item.UUID) -match "^[A-F0-9]{8}(-[A-F0-9]{4}){3}-[A-F0-9]{12}$") {
            $item.UUID = "`{$($item.UUID)`}"
        }
    }

    # Boot image file must either exist or the entry must be empty
    if(-not [string]::IsNullOrWhitespace($($item.BootImg)) -and -not [System.IO.File]::Exists("$wdsRoot\$($item.BootImg)")) {
        Write-host "CHYBA: Neexistujici boot image."
        Write-host ""
        continue;
    }

    # UnattendXML file must either exist or the entry must be empty string
    if(-not [string]::IsNullOrWhitespace($($item.UnattendXml)) -and -not [System.IO.File]::Exists("$wdsRoot\$($item.UnattendXml)")) {
        Write-host "CHYBA: Neexistujici UnattendXML."
        Write-host ""
        continue;
    }

    # OU must either exist in AD or be empty string
    if(-not [string]::IsNullOrWhitespace($($item.OU)) -and -not (existsOU($($item.OU)))) {
        Write-host "CHYBA: Neexistujici UnattendXML."
        Write-host ""
        continue;
    }


    
    # Generating of command
    $command = "cmd.exe /c C:\Windows\sysnative\wdsutil /add-device "
    $command += "/Device:`"$($item.Name)`" "
    $command += "/ID:`"$($item.UUID)`" "
    if(-not [string]::IsNullOrWhitespace($($item.BootImg))) {
        $command += "/BootImagePath:`"$($item.BootImg)`" "
    }
    if(-not [string]::IsNullOrWhitespace($($item.UnattendXml))) {
        $command += "/WDSClientUnattend:`"$($item.UnattendXml)`" "
    }
    if(-not [string]::IsNullOrWhitespace($($item.Group))) {
        $command += "/Group:`"$($item.Group)`" "
    }
    if(-not [string]::IsNullOrWhitespace($($item.OU))) { # OU can be set only when adding new device
        $commandOU = $command + "/OU:`"$($item.OU)`" "
    }

    # Run generated command
    Invoke-Expression $commandOU | Out-Null
    $errorCode = $LASTEXITCODE

    # Check result of the command
    if($errorCode -eq 0) {
        Write-host -ForegroundColor Green "$($item.Name) Successfully added."
    } else {
        if(($errorCode -eq -1056702156) -or ($errorCode -eq -1056767687)) { # Only change settings if the computer is joined in domain or already prestaged
            if($changeIfExists){
                Write-host -ForegroundColor Yellow "$($item.Name) Device already exists."

                # Attempt to change device instead adding it
                $command = $command -replace "add-device","set-device"
                Invoke-Expression $command | Out-Null
                $errorCode = $LASTEXITCODE

                if(($errorCode -eq 0) -or ($errorCode -eq 87)) {
                        Write-host -ForegroundColor Green "$($item.Name) Successfully changed."
                } else {  # Unknown error
                    Write-host -ForegroundColor Red "$($item.Name) ERROR $errorCode"
                    Write-host -ForegroundColor DarkRed $command
                }
            } else {  # Already exists
                Write-host -ForegroundColor Red "$($item.Name) ERROR - device already exists $errorCode"
                Write-host -ForegroundColor DarkRed $command
            }
        } else {  # Other error
                Write-host -ForegroundColor Red "$($item.Name) ERROR $errorCode"
                Write-host -ForegroundColor DarkRed $command
        }
    }
    Write-host ""
}

# TODO automatically change wdsutil location sys32 - sysnative
# TODO move csv name to variable or pass as parameter
