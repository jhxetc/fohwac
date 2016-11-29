Param (
    [switch]$background,
    [int]$addon
)
function initialize-Config () {
    $gcConfigFile = $MyInvocation.ScriptName.Replace('ps1','json')
    if (!(Test-Path $gcConfigFile)) {
        $gcObj = New-Object -TypeName psobject
        if (${env:ProgramFiles(x86)}) {
            $pfDir = ${env:ProgramFiles(x86)}
        }
        else {
            $pfDir = $env:ProgramFiles
        }
        Add-Member -InputObject $gcObj -Type NoteProperty -Name "InstallPath" -Value "$pfDir\World of Warcraft\Interface\AddOns"
        Add-Member -InputObject $gcObj -Type NoteProperty -Name "TempPath" -Value $env:TEMP
        Add-Member -InputObject $gcObj -Type NoteProperty -Name "Addons" -Value @()
        write-Config $gcObj
    }
    else {
        $gcConfigUri = ([System.Uri]$gcConfigFile).AbsoluteUri
        $gcObj = Invoke-RestMethod -Uri $gcConfigUri
        #write-host -ForegroundColor Green ">>> $($MyInvocation.MyCommand.Name)"
        #out-host -InputObject $gcObj
    }
    return $gcObj
}

function write-Config ($gcObj) {
    ConvertTo-Json -InputObject $gcObj | Out-File -encoding ascii -force $MyInvocation.ScriptName.Replace('ps1','json')
}

function add-Addon ($gcObj) {
    $newAddonsObj = New-Object -TypeName psobject -Property @{"AddonName"=$null; "AddonUrl"=$null; "AddonRel"=$null}
    do {
        $done = $false
        $addonName = Read-Host "`n`tAddon name ('q' to exit)"
        if ($gcObj.Addons.AddonName -contains $addonName.Trim()) {
            Write-Host -ForegroundColor Red "`n`tAddon with name $addonName already exists.`n"
        }
        elseif ($addonName -eq "q") {
            return $null
        }
        else {
            $done = $true
        }
    } until ($done -and $addonName)
    
    do {
        $done = $false
        $addonUrl = Read-Host "`n`tAddon url ('q' to exit)"
        $addonUrl = $addonUrl.Split("?")[0]
        $addonUrl = $addonUrl -Replace "files/?","files"
        if ($addonUrl -match "^https{0,1}:{1}/{2}[a-zA-Z_0-9].*files$") {
            $done = $true
        }
        elseif ($addonUrl -eq "q") {
            return $null
        }
        else {
            Write-Host -ForegroundColor Red "`n`tInvalid url... (url should end with /files)`n"
        }
    } until ($done)
    do {
       $done = $false
       $addonRel = Read-Host "`n`tRelease versions only [Y]es or [N]o? ('q' to exit)"
       switch -regex ($addonRel) {
           "^y" {$addonRel = $true; $done = $true}
           "^n" {$addonRel = $false; $done = $true}
           "^q" {return $null}
           default {write-host -ForegroundColor Red "`n`tInvalid response...`n"}
       }
    } until ($done)
    $newAddonsObj.AddonName = $addonName
    $newAddonsObj.AddonUrl = $addonUrl
    $newAddonsObj.AddonRel = [int]$addonRel
    $gcObj.Addons += $newAddonsObj
    write-Config $gcObj
    Write-Host -ForegroundColor Cyan "`n`tAddon: $addonName Successfully Added."
}

function update-Addon ($gcObj) {
    show-Addons
    try {
        [int]$addonNum = Read-Host "`n`tNumber of Addon to Update (any non-number to exit)"
        if (!$gcObj.Addons[$addonNum].AddonName) {
            write-host -ForegroundColor Red "`n`tAddon number $addonNum not found...`n"
            throw
        }
    }
    catch {
        return $null
    }

    do {
        $done = $false
        $addonName = Read-Host "`n`tUpdate addon name: $($gcObj.Addons[$addonNum].AddonName) (leave blank to retain name or 'q' to exit)"
        if ([string]::IsNullOrEmpty($addonName)) {
            $done = $true
        }
        elseif ($gcObj.Addons.AddonName -contains $addonName) {
            Write-Host -ForegroundColor Red "`n`tName: $addonName already exists`n"
        }
        elseif ($addonName -eq "q") {
            return $null
        }
        else {
            $gcObj.Addons[$addonNum].AddonName = $addonName
            $done = $true
        }
    } until ($done)
    
    do {
        $done = $false
        $addonUrl = Read-Host "`n`tUpdate addon url: $($gcObj.Addons[$addonNum].AddonUrl) (leave blank to retain value or 'q' to exit)"
        $addonUrl = $addonUrl.Split("?")[0]
        $addonUrl = $addonUrl -Replace "files/?","files"
        if ([string]::IsNullOrEmpty($addonUrl)) {
            $done = $true
        }
        elseif ($addonUrl -match "^https{0,1}:{1}/{2}[a-zA-Z_0-9].*files$") {
            $gcObj.Addons[$addonNum].AddonUrl = $addonUrl
            $done = $true
        }
        elseif ($addonUrl -eq "q") {
            return $null
        }
        else {
            Write-Host -ForegroundColor Red "`n`tInvalid url... (url should end with /files)`n"
        }
    } until ($done)
    do {
       $done = $false
       switch ($gcObj.Addons[$addonNum].AddonRel) {
           0 {$curAddonRel = "N"}
           1 {$curAddonRel = "Y"}
           default {$curAddonRel = "N"}
       }
       $addonRel = Read-Host "`n`tDownload release versions only? [Y]es or [N]o? current: $curAddonRel ('q' to exit)"
       switch -regex ($addonRel) {
           "^y" {$addonRel = $true; $gcObj.Addons[$addonNum].AddonRel = [int]$addonRel; $done = $true}
           "^n" {$addonRel = $false; $gcObj.Addons[$addonNum].AddonRel = [int]$addonRel; $done = $true}
           "^q" {return $null}
           default {write-host -ForegroundColor Red "`n`tInvalid response...`n"}
       }
    } until ($done)
    write-Config $gcObj
    Write-Host -ForegroundColor Cyan "`n`t$($gcObj.Addons[$addonNum].AddonName) Successfully Updated.`n"
}

function remove-Addon ($gcObj) {
    show-Addons
    $addonNum = Read-Host "`n`tNumber of Addon to Remove (any non-number to exit)"
    if ([string]::IsNullOrEmpty($addonNum)) {
        $addonNum = "q"
    }
    try {
        [int]$addonNum = $addonNum
        if (!$gcObj.addons[$addonNum].AddonName) {
            write-host -ForegroundColor Red "`n`tAddon number $addonNum not found...`n"
            throw
        }
    }
    catch {
        return $null
    }
    $newAddonsObj = New-Object -TypeName psobject -Property @{"Addons"=@()}
    For ($i = 0; $i -lt $gcObj.Addons.Count; $i++) {
        if ($i -eq $addonNum) {
            $remAddonName = $gcObj.Addons[$i].AddonName
            continue
        }
        else {
            $newAddonsObj.Addons += $gcObj.Addons[$i]
        }
    }
    $gcObj.Addons = $newAddonsObj.Addons
    write-Config $gcObj
    Write-Host -ForegroundColor Yellow "`n`t$remAddonName Successfully Removed`n"
}

function show-Addons () {
    Write-Host "`n"
    For ($i = 0; $i -lt $gcObj.Addons.Count; $i++) {
        Write-Host -NoNewLine -ForegroundColor Cyan "`t$i. "
        Write-Host -NoNewline -ForegroundColor DarkGray "$($gcObj.Addons[$i].AddonName) " 
        Write-Host -NoNewline -ForegroundColor Cyan "("
        Write-Host -NoNewline -ForegroundColor Gray "$($gcObj.Addons[$i].AddonUrl)"
        Write-Host -NoNewline -ForegroundColor Cyan ") " 
        Write-Host -ForegroundColor Cyan "$([bool]$gcObj.Addons[$i].AddonRel)"
    }
}

function update-installPath ($gcObj) {
    do {
        $done = $false
        Write-Host -NoNewLine -ForegroundColor Yellow "`n`tChange Addon Installation Path: "
        Write-Host -NoNewline "$($gcObj.InstallPath)? ('q' to exit): "
        $newInstallPath = read-host
        if ([string]::IsNullOrEmpty($newInstallPath)) {
            if (Test-Path $gcObj.InstallPath) {
                $done = $true
            }
            else {
                Write-Host -ForegroundColor Red "`n`t$($gcObj.InstallPath) is not accessible..."
            }
        }
        elseif ($newInstallPath -eq "q") {
            Write-Host "`n`tExiting."
            $done = $true
        }
        else {
            if (Test-Path $newInstallPath) {
                $gcObj.InstallPath = (Convert-Path $newInstallPath)
                write-Config $gcObj
                Write-Host -ForegroundColor Yellow "`n`tAddon Installation Path Updated Successfully."
                $done = $true
            }
            else {
                Write-Host -ForegroundColor Red "`n`t$newInstallPath is not accessible..."
            }
        }
    } until ($done)
}

function update-tempPath ($gcObj) {
    do {
        $done = $false
        Write-Host -NoNewLine -ForegroundColor Yellow "`n`tChange Temporary Files Path: "
        Write-Host -NoNewline "$($gcObj.TempPath)? ('q' to exit): "
        $newTempPath = read-host
        if ([string]::IsNullOrEmpty($newTempPath)) {
            if (Test-Path $gcObj.TempPath) {
                $done = $true
            }
            else {
                Write-Host -ForegroundColor Red "`n`t$($gcObj.TempPath) is not accessible..."
            }
        }
        elseif ($newTempPath -eq "q") {
            Write-Host "`n`tExiting."
            $done = $true
        }
        else {
            if (Test-Path $newTempPath) {
                $gcObj.TempPath = (Convert-Path $newTempPath)
                write-Config $gcObj
                Write-Host -ForegroundColor Yellow "`n`tTemporary Files Path Updated Successfully."
                $done = $true
            }
            else {
                Write-Host -ForegroundColor Red "`n`t$newTempPath is not accessible..."
            }
        }
    } until ($done)
}

function show-Paths ($gcObj) {
    Write-Host -NoNewLine -ForegroundColor Yellow "`n`tAddon Installation Path: "
    Write-Host $gcObj.InstallPath
    Write-Host -NoNewline -ForegroundColor Yellow "`n`tTemporary Files Path: "
    Write-Host $gcObj.TempPath
}

function install-Addons ($gcObj, $addonNum) {
    $start = $addonNum
    if ($addonNum) {
        $end = $addonNum
    }
    else {
        $end = $gcObj.Addons.Count-1
    }
    foreach ($addonObj in $gcObj.addons[$start..$end]) {
        if ($addonObj.AddonRel -eq $true) {
            $addonUrl = "$($addonObj.AddonUrl)/latest"
        }
        else {
            $addonPage = Invoke-WebRequest $addonObj.AddonUrl
            $addonLinks = $addonPage.Links | Where-Object {$_.Title -eq "download file"} | Select-Object href
            $addonUrlParse = $addonObj.AddonUrl.Split("/")
            $addonUrl = "$($addonUrlParse[0])//$($addonUrlParse[2])$($addonLinks[0].href)"
        }
        Write-Host -ForegroundColor Cyan "`n`tDownloading the latest $($addonObj.AddonName) file..."
        $addonZip = "$($gcObj.TempPath)\$($addonObj.AddonName).zip"
        Invoke-WebRequest -Uri $addonUrl -OutFile $addonZip
        Write-Host -ForegroundColor DarkGray "`t...Download Complete"
        write-Host -ForegroundColor Yellow "`n`tUnzipping $($addonObj.AddonName).zip ..."
        $shell = new-object -com shell.application
        $zip = $shell.NameSpace($addonZip)
        foreach($zipItem in $zip.items())
        {
            $shell.Namespace($gcObj.InstallPath).copyhere($zipItem, 0x14)
        }
        Write-Host -ForegroundColor Green "`t...Installation of $($addonObj.AddonName) Complete"
    }
}

function invoke-Menu ($gcObj) {
    do {
Write-Host -NoNewLine -ForegroundColor DarkGreen @"
`n`t[L]ist addon entries
`t[N]ew addon entry
`t[C]hange addon entry
`t[R]emove addon entry
`t[S]how Paths
`t[E]dit paths 
`t[I]nstall/Update ALL Addons
`t[U]pdate/Install Specific Addon
`t[Q]uit:
"@
        $choice = Read-Host 
        switch ($choice) {
            "l" {show-Addons}
            "n" {add-Addon $gcObj}
            "c" {
                if ($gcObj.Addons.count -gt 0) {
                    update-Addon $gcObj
                } 
                else {
                    write-host -ForegroundColor Red "`n`tNo addon entries found.`n"
                }
            }
            "r" {
                if ($gcObj.Addons.count -gt 0) {
                    remove-Addon $gcObj
                } 
                else {
                    write-host -ForegroundColor Red "`n`tNo addon entries found.`n"
                }
            }
            "s" {show-Paths $gcObj}
            "e" {update-installPath $gcObj; update-tempPath $gcObj}
            "i" {install-Addons $gcObj}
            "u" {
                show-Addons
                $addonNum = Read-Host "`n`tNumber of Specific Addon to Install/Update (any non-number to exit)"
                try {
                    if ([string]::IsNullOrEmpty($addonNum)) {
                        throw
                    }
                    else {
                        [int]$addonNum = $addonNum
                        install-Addons $gcObj $addonNum
                    }
                }
                catch {
                    $addonNum = $null
                }
            }
            "q" {Write-Host -ForegroundColor Yellow "`n`tExiting...`n"}
            default {Write-Host -ForegroundColor Red "`n`tInvalid Option: $choice`n"}
        }    
    } until ($choice -eq "q")
}
 
$gcObj = initialize-Config
if (!(Test-Path $gcObj.InstallPath)) {
    Write-Host -ForegroundColor Red "`n`tInvalid Addon Installation Path"
    if (!$background) {
        update-installPath $gcObj
    }
}
if (!(Test-Path $gcObj.TempPath)) {
    Write-Host -ForegroundColor Red "`n`tInvalid Temp Path"
    if (!$background) {
        update-tempPath $gcObj
    }
}
if ($background) {
    install-Addons $gcObj $addon
}
else {
    invoke-Menu $gcObj
}


