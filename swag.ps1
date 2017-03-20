Param (
    [switch]$batch,
    [int]$addon,
    [string]$json="$PSScriptRoot\addons.json"
)

$curseUrl = "https://wow.curseforge.com/projects"
$wowAceUrl = "https://www.wowace.com/projects"
$urlSuffix = "files"

function initialize-Config () {
    $configFile = "$PSScriptRoot\config.json"
    if (!(Test-Path $configFile)) {
        $configObj = New-Object -TypeName psobject
        if (${env:ProgramFiles(x86)}) {
            $pfDir = ${env:ProgramFiles(x86)}
        }
        else {
            $pfDir = $env:ProgramFiles
        }
        Add-Member -InputObject $configObj -Type NoteProperty -Name "InstallPath" -Value "$pfDir\World of Warcraft\Interface\AddOns"
        Add-Member -InputObject $configObj -Type NoteProperty -Name "TempPath" -Value $env:TEMP
        Add-Member -InputObject $configObj -Type NoteProperty -Name "WowVers" -Value $null
        #Add-Member -InputObject $configObj -Type NoteProperty -Name "DownloadOld" -Value 0
        write-configJson $configObj
    }
    else {
        $configUri = ([System.Uri]$configFile).AbsoluteUri
        $configObj = Invoke-RestMethod -Uri $configUri
        #write-host -ForegroundColor Green ">>> $($MyInvocation.MyCommand.Name)"
        #out-host -InputObject $configObj
    }
    return $configObj
}

function initialize-Addons ($subJson) {
    if ($subJson) {
        $json = $subJson
    }
    if ($json -match "^https{0,1}:{1}/{2}[a-zA-Z_0-9]") {
        try {
            $addonsObj = Invoke-RestMethod -Uri $json
        }
        catch {
            Write-Host -ForegroundColor Red "`n`tError loading addons from url: $json"
            return $null
        }
    }
    elseif (Test-Path $json) {
        $addonsFile = Convert-Path $json
        $addonsUri = ([System.Uri]$json).AbsoluteUri
        $addonsObj = Invoke-RestMethod -Uri $addonsUri
    }
    else {
        $addonsObj = New-Object -TypeName psobject -Property @{"Addons"=@()}
        write-addonsJson $json
    }
    return $addonsObj 
}

function write-configJson () {
    ConvertTo-Json -InputObject $configObj | Out-File -encoding ascii -force "$PSScriptRoot\config.json"
}

function write-addonsJson ($json) {
    if ($json -match "^https{0,1}:{1}/{2}[a-zA-Z_0-9]") {
        Write-Host -ForegroundColor Red "`n`t$json is not a local file can't save changes."        
    }
    else {
        ConvertTo-Json -InputObject $addonsObj | Out-File -encoding ascii -force $json
    }
}

function add-Addon () {
    if ($json -match "^https{0,1}:{1}/{2}[a-zA-Z_0-9]") {
        Write-Host -ForegroundColor Red -BackgroundColor Black "`n`t!!$json is not a local file CHANGES WILL NOT BE SAVED!!."        
    }
    $newAddonsObj = New-Object -TypeName psobject -Property @{"AddonName"=$null; "AddonUrl"=$null; "AddonRel"=$null; "DownloadOld"=0}
    do {
        $done = $false
        $addonName = Read-Host "`n`tAddon name ('q' to exit)"
        if ($addonsObj.Addons.AddonName -contains $addonName.Trim()) {
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
    do {
       $done = $false
       $downloadOld = Read-Host "`n`tDownload older versions? [Y]es or [N]o? current: $curDownloadOld ('q' to exit)"
       switch -regex ($downloadOld) {
           "^y" {$downloadOld = $true; $done = $true}
           "^n" {$downloadOld = $false; $done = $true}
           "^q" {return $null}
           default {write-host -ForegroundColor Red "`n`tInvalid response...`n"}
       }
    } until ($done)
    $newAddonsObj.AddonName = $addonName
    $newAddonsObj.AddonUrl = $addonUrl
    $newAddonsObj.AddonRel = [int]$addonRel
    $newAddonsObj.DownloadOld = [int]$downloadOld
    $addonsObj.Addons += $newAddonsObj
    write-addonsJson $json
    Write-Host -ForegroundColor Cyan "`n`tAddon: $addonName Successfully Added."
}

function update-Addon () {
    if ($addonsObj.Addons.count -le 0) {
        write-host -ForegroundColor Red "`n`tNo addon entries found.`n"
        return $null
    }
    if ($json -match "^https{0,1}:{1}/{2}[a-zA-Z_0-9]") {
        Write-Host -ForegroundColor Red -BackgroundColor Black "`n`t!!$json is not a local file CHANGES WILL NOT BE SAVED!!."        
    }
    show-Addons
    try {
        [int]$addonNum = Read-Host "`n`tNumber of Addon to Update (any non-number to exit)"
        if (!$addonsObj.Addons[$addonNum].AddonName) {
            write-host -ForegroundColor Red "`n`tAddon number $addonNum not found...`n"
            throw
        }
    }
    catch {
        return $null
    }

    do {
        $done = $false
        $addonName = Read-Host "`n`tUpdate addon name: $($addonsObj.Addons[$addonNum].AddonName) (leave blank to retain name or 'q' to exit)"
        if ([string]::IsNullOrEmpty($addonName)) {
            $done = $true
        }
        elseif ($addonsObj.Addons.AddonName -contains $addonName) {
            Write-Host -ForegroundColor Red "`n`tName: $addonName already exists`n"
        }
        elseif ($addonName -eq "q") {
            return $null
        }
        else {
            $addonsObj.Addons[$addonNum].AddonName = $addonName
            $done = $true
        }
    } until ($done)
    
    do {
        $done = $false
        $addonUrl = Read-Host "`n`tUpdate addon url: $($addonsObj.Addons[$addonNum].AddonUrl) (leave blank to retain value or 'q' to exit)"
        $addonUrl = $addonUrl.Split("?")[0]
        $addonUrl = $addonUrl -Replace "files/?","files"
        if ([string]::IsNullOrEmpty($addonUrl)) {
            $done = $true
        }
        elseif ($addonUrl -match "^https{0,1}:{1}/{2}[a-zA-Z_0-9].*files$") {
            $addonsObj.Addons[$addonNum].AddonUrl = $addonUrl
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
       switch ($addonsObj.Addons[$addonNum].AddonRel) {
           0 {$curAddonRel = "N"}
           1 {$curAddonRel = "Y"}
           default {$curAddonRel = "N"}
       }
       $addonRel = Read-Host "`n`tDownload release versions only? [Y]es or [N]o? current: $curAddonRel ('q' to exit)"
       switch -regex ($addonRel) {
           "^y" {$addonRel = $true; $addonsObj.Addons[$addonNum].AddonRel = [int]$addonRel; $done = $true}
           "^n" {$addonRel = $false; $addonsObj.Addons[$addonNum].AddonRel = [int]$addonRel; $done = $true}
           "^q" {return $null}
           default {write-host -ForegroundColor Red "`n`tInvalid response...`n"}
       }
    } until ($done)
    do {
       $done = $false
       switch ($addonsObj.Addons[$addonNum].DownloadOld) {
           0 {$curDownloadOld = "N"}
           1 {$curDownloadOld = "Y"}
           default {$curDownloadOld = "N"}
       }
       $downloadOld = Read-Host "`n`tDownload older versions? [Y]es or [N]o? current: $curDownloadOld ('q' to exit)"
       switch -regex ($downloadOld) {
           "^y" {$downloadOld = $true; $addonsObj.Addons[$addonNum].DownloadOld = [int]$downloadOld; $done = $true}
           "^n" {$downloadOld = $false; $addonsObj.Addons[$addonNum].DownloadOld = [int]$downloadOld; $done = $true}
           "^q" {return $null}
           default {write-host -ForegroundColor Red "`n`tInvalid response...`n"}
       }
    } until ($done)
    write-addonsJson $json
    Write-Host -ForegroundColor Cyan "`n`t$($addonsObj.Addons[$addonNum].AddonName) Successfully Updated.`n"
}

function remove-Addon () {
    if ($addonsObj.Addons.count -le 0) {
        write-host -ForegroundColor Red "`n`tNo addon entries found.`n"
        return $null
    }
    if ($json -match "^https{0,1}:{1}/{2}[a-zA-Z_0-9]") {
        Write-Host -ForegroundColor Red -BackgroundColor Black "`n`t!!$json is not a local file CHANGES WILL NOT BE SAVED!!."        
    }
    show-Addons
    $addonNum = Read-Host "`n`tNumber of Addon to Remove (any non-number to exit)"
    if ([string]::IsNullOrEmpty($addonNum)) {
        $addonNum = "q"
    }
    try {
        [int]$addonNum = $addonNum
        if (!$addonsObj.Addons[$addonNum].AddonName) {
            write-host -ForegroundColor Red "`n`tAddon number $addonNum not found...`n"
            throw
        }
    }
    catch {
        return $null
    }
    $newAddonsObj = New-Object -TypeName psobject -Property @{"Addons"=@()}
    For ($i = 0; $i -lt $addonsObj.Addons.Count; $i++) {
        if ($i -eq $addonNum) {
            $remAddonName = $addonsObj.Addons[$i].AddonName
            continue
        }
        else {
            $newAddonsObj.Addons += $addonsObj.Addons[$i]
        }
    }
    $addonsObj.Addons = $newAddonsObj.Addons
    write-addonsJson $json
    Write-Host -ForegroundColor Yellow "`n`t$remAddonName Successfully Removed`n"
}

function show-Addons () {
    Write-Host "`n`tCurrent addons file: $json"
    if ($json -match "^https{0,1}:{1}/{2}[a-zA-Z_0-9]") {
        Write-Host -ForegroundColor Red -BackgroundColor Black "`n`t!!$json is not a local file CHANGES WILL NOT BE SAVED!!."        
    }
    Write-Host "`n"
    For ($i = 0; $i -lt $addonsObj.Addons.Count; $i++) {
        Write-Host -NoNewLine -ForegroundColor Cyan "`t$i. "
        Write-Host -NoNewline -ForegroundColor DarkGray "$($addonsObj.Addons[$i].AddonName) " 
        Write-Host -NoNewline -ForegroundColor Cyan "("
        Write-Host -NoNewline -ForegroundColor Gray "$($addonsObj.Addons[$i].AddonUrl)"
        Write-Host -NoNewline -ForegroundColor Cyan ") " 
        Write-Host -ForegroundColor Cyan "$([bool]$addonsObj.Addons[$i].AddonRel)"
    }
}

function update-installPath () {
    do {
        $done = $false
        Write-Host -NoNewLine -ForegroundColor Yellow "`n`tChange Addon Installation Path: "
        Write-Host -NoNewline "$($configObj.InstallPath)? ('q' to exit): "
        $newInstallPath = read-host
        if ([string]::IsNullOrEmpty($newInstallPath)) {
            if (Test-Path $configObj.InstallPath) {
                $done = $true
            }
            else {
                Write-Host -ForegroundColor Red "`n`t$($configObj.InstallPath) is not accessible..."
            }
        }
        elseif ($newInstallPath -eq "q") {
            Write-Host "`n`tExiting."
            $done = $true
        }
        else {
            if (Test-Path $newInstallPath) {
                $configObj.InstallPath = (Convert-Path $newInstallPath)
                write-configJson
                Write-Host -ForegroundColor Yellow "`n`tAddon Installation Path Updated Successfully."
                $done = $true
            }
            else {
                Write-Host -ForegroundColor Red "`n`t$newInstallPath is not accessible..."
            }
        }
    } until ($done)
}

function update-tempPath () {
    do {
        $done = $false
        Write-Host -NoNewLine -ForegroundColor Yellow "`n`tChange Temporary Files Path: "
        Write-Host -NoNewline "$($configObj.TempPath)? ('q' to exit): "
        $newTempPath = read-host
        if ([string]::IsNullOrEmpty($newTempPath)) {
            if (Test-Path $configObj.TempPath) {
                $done = $true
            }
            else {
                Write-Host -ForegroundColor Red "`n`t$($configObj.TempPath) is not accessible..."
            }
        }
        elseif ($newTempPath -eq "q") {
            Write-Host "`n`tExiting."
            $done = $true
        }
        else {
            if (Test-Path $newTempPath) {
                $configObj.TempPath = (Convert-Path $newTempPath)
                write-configJson
                Write-Host -ForegroundColor Yellow "`n`tTemporary Files Path Updated Successfully."
                $done = $true
            }
            else {
                Write-Host -ForegroundColor Red "`n`t$newTempPath is not accessible..."
            }
        }
    } until ($done)
}

function update-wowVers () {
    do {
        $done = $false
        Write-Host -NoNewLine -ForegroundColor Yellow "`n`tChange WoW Release Version: "
        Write-Host -NoNewline "$($configObj.WowVers)? ('q' to exit): "
        $newWowVers = read-host
        if ($newWowVers -match "^[0-9]{1,2}\.[0-9]{1,2}\.[0-9]{1,2}$") {
            $configObj.WowVers = $newWowVers
            write-configJson
            Write-Host -ForegroundColor Yellow "`n`tWoW Release Version Updated Successfully"
            $done = $true
        }
        elseif ($newWowVers -eq "q") {
            Write-Host "`n`tExiting."
            $done = $true
        }
        else {
            Write-Host -ForegroundColor Red "`n`t$newWowVers is an invalid version (format is #.#.#)" 
        }
    } until ($done)

}
<#
function update-downloadOld () {
    do {
        $done = $false
        Write-Host -NoNewLine -ForegroundColor Yellow "`n`tDownload out of date Addons: "
        Write-Host -NoNewline "[Y]es or [N]o? ('q' to exit): "
        $newDownloadOld = read-host 
        switch -Regex ($newDownloadOld) {
            "^y" {$configObj.DownloadOld = 1; write-configJson; Write-Host -ForegroundColor Yellow "`n`tDownload out of date Addons Updated Successfully"; $done = $true}
            "^n" {$configObj.DownloadOld = 0; write-configJson; Write-Host -ForegroundColor Yellow "`n`tDownload out of date Addons Updated Successfully"; $done = $true}
            "^q" {Write-Host "`n`tExiting."; $done = $true}
            default {Write-Host -ForegroundColor Red "`n`tInvalid Option...`n"}
        }
    } until ($done)
}
#>

function update-Settings () {
    do {
        $done = $false
        show-Settings
        $choice = Read-Host "`n`tNumber of Setting to Change ('q' to exit)"
        switch ($choice) {
            0 {update-installPath; $done = $true}
            1 {update-tempPath; $done = $true}
            2 {update-wowVers; $done = $true}
            #3 {update-downloadOld; $done = $true}
            "q" {$done = $true}
            default {Write-Host -ForegroundColor Red "`n`tInvalid Option..."}
        }
    } until ($done)
}

function show-Settings () {
    Write-Host -NoNewLine -ForegroundColor Yellow "`n`t0. Addon Installation Path: "
    Write-Host $configObj.InstallPath
    Write-Host -NoNewline -ForegroundColor Yellow "`n`t1. Temporary Files Path: "
    Write-Host $configObj.TempPath
    Write-Host -NoNewline -ForegroundColor Yellow "`n`t2. Current WoW release: "
    Write-Host $configObj.WowVers
    #Write-Host -NoNewline -ForegroundColor Yellow "`n`t3. Download Out of Date Addons: "
    #Write-Host ([bool]$configObj.DownloadOld)
}

function install-Addons ($addonNum) {
    [int]$start = $addonNum
    if (!([string]::IsNullOrEmpty($addonNum))) {
        [int]$end = $addonNum
    }
    else {
        $end = $addonsObj.Addons.Count-1
    }
    foreach ($addonObj in $addonsObj.Addons[$start..$end]) {
        if ($addonObj.AddonRel -eq $true) {
            $addonUrl = "$($addonObj.AddonUrl)/latest"
        }
        else {
            try {
                $addonPage = Invoke-WebRequest $addonObj.AddonUrl
            }
            catch {
                Write-Host -ForegroundColor Red "`n`t$($addonObj.AddonUrl) is not valid.`n"
                continue
            }
            
            $addonVers = $addonPage.parsedhtml.body.getElementsByClassName("version-label")[0].innerText
            $addonLinks = $addonPage.Links | Where-Object {$_.Title -eq "download file"} | Select-Object href
            $addonUrlParse = $addonObj.AddonUrl.Split("/")
            $addonUrl = "$($addonUrlParse[0])//$($addonUrlParse[2])$($addonLinks[0].href)"
        }
        if (($addonVers.Replace('.','') -ge $configObj.WowVers.Replace('.','')) -or ($addonObj.DownloadOld -eq 1)) {
            Write-Host -ForegroundColor Cyan "`n`tDownloading the latest $($addonObj.AddonName) file..."
            $addonZip = "$($configObj.TempPath)\$($addonObj.AddonName).zip"
            Invoke-WebRequest -Uri $addonUrl -OutFile $addonZip
            Write-Host -ForegroundColor DarkGray "`t...Download Complete"
            write-Host -ForegroundColor Yellow "`n`tUnzipping $($addonObj.AddonName).zip ..."
            $shell = new-object -com shell.application
            $zip = $shell.NameSpace($addonZip)
            foreach($zipItem in $zip.items())
            {
                $shell.Namespace($configObj.InstallPath).copyhere($zipItem, 0x14)
            }
            Write-Host -ForegroundColor Green "`t...Installation of $($addonObj.AddonName) Complete"
        }
        else {
            Write-Host -ForegroundColor Red "`n`tAddon $($addonObj.addonname) is out of date, skipping..."
        }
    }
}

function get-objFromString ($encString) {
    $stringObj = New-Object -TypeName psobject -Property @{"Addons"=@()}
    $decString = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($encString))
    $decArray = $decString.Split(',')
    ForEach ($dec in $decArray) {
        if ($dec -match "^[0-1]:") {
            $subArray = $dec.Split(':')
            if ($subArray[0] -eq 0) {
                $subObj = New-Object -TypeName psobject -Property @{"AddonName"=$subArray[1]; "AddonUrl"="$curseUrl/$($subArray[1])/$urlSuffix"; "AddonRel"=0; "DownloadOld"=0}
            }
            else {
                $subObj = New-Object -TypeName psobject -Property @{"AddonName"=$subArray[1]; "AddonUrl"="$wowAceUrl/$($subArray[1])/$urlSuffix"; "AddonRel"=0; "DownloadOld"=0}
            }
            $stringObj.Addons += $subObj          
        }
        else {
            Write-Host -ForegroundColor Red "`n`tError decoding an addon in string."
        }  
    }
    return $stringObj
}

function invoke-Menu () {
    $exit = $false
    do {
Write-Host -NoNewLine -ForegroundColor DarkGreen @"
`n`t1. Show All Addons
`t2. Add New Addon
`t3. Change/Update Addon
`t4. Delete Addon
`t5. Load Addon File
`t6. Write Addon File
`t7. Enter Encoded Addon String
`t8. Show All Program Settings
`t9. Change/Update Program Setting
`t10. Install/Update ALL Addons
`t11. Install/Update SINGLE Addon
`t12. Uninstall SINGLE Addon
`t0. Exit Program
"@
        $choice = Read-Host 
        switch ($choice) {
            1 {show-Addons}
            2 {add-Addon}
            3 {update-Addon}
            4 {remove-Addon}
            5 {
                Write-Host -NoNewline -ForegroundColor Magenta "`n`tNew Addon File Path (Local file path or web address - leave blank for default): "
                $subJson = Read-Host
                if (!([string]::IsNullOrEmpty($subJson))) {
                    $subAddonsObj = initialize-Addons $subJson
                    if ($subAddonsObj.Addons.Count -gt 0) {
                        $addonsObj = $subAddonsObj
                        $json = $subJson
                        Write-Host -ForegroundColor Cyan "`n`tNew Addon File Loaded Successfully."
                    } 
                    else {
                        Write-Host -ForegroundColor Red "`n`tInvalid Addon File/Path..."
                    } 
                }
                else {
                    $json = $origJson
                    $addonsObj = initialize-Addons $json
                    Write-Host -ForegroundColor Cyan "`n`tDefault Addons File Loaded Successfully."
                }
            }
            6 {
                Write-Host -NoNewline -ForegroundColor DarkCyan "`n`tEnter new or existing file name to write data: "
                $newJson = Read-Host
                if (Test-Path -IsValid $newJson) {
                    write-addonsJson $newJson
                    Write-Host -ForegroundColor Cyan "`n`tAddon Data Written Successfully to $(convert-path $newJson)"
                }
                else {
                    Write-Host -ForegroundColor Red "`n`t$newJson is not formatted properly...`n"
                }
            }
            7 {
                Write-Host -NoNewline -ForegroundColor Cyan "`n`tPaste encoded download string: "
                $encString = Read-Host
                $stringAddonsObj = get-objFromString $encString
                if ($stringAddonsObj.Addons.Count -gt 0) {
                    $addonsObj = $stringAddonsObj
                    $json = "http://string.obj"
                    Write-Host -ForegroundColor Cyan "`n`tAddons From String Loaded Successfully."
                }
                else {
                    Write-Host -ForegroundColor Red "`n`tString Did Not Contain Addon Data..."
                }
            }
            8 {show-Settings}
            9 {update-Settings}
            10 {install-Addons}
            11 {
                show-Addons
                $addonNum = Read-Host "`n`tNumber of Specific Addon to Install/Update (any non-number to exit)"
                try {
                    if ([string]::IsNullOrEmpty($addonNum)) {
                        throw
                    }
                    else {
                        [int]$addonNum = $addonNum
                        install-Addons $addonNum
                    }
                }
                catch {
                    $addonNum = $null
                }
            }
            0 {Write-Host -ForegroundColor Yellow "`n`tExiting...`n"; $exit = $true}
            default {Write-Host -ForegroundColor Red "`n`tInvalid Option: $choice`n"}
        }    
    } until ($exit)
}

if ($batch) {
    Start-Transcript -Path ($MyInvocation.MyCommand.Path).Replace('ps1','log') -IncludeInvocationHeader
}
$configObj = initialize-Config
$addonsObj = initialize-Addons
$origJson = $json
if (!(Test-Path $configObj.InstallPath)) {
    Write-Host -ForegroundColor Red "`n`tInvalid Addon Installation Path"
    if ($batch) {
        exit 1
    }
    else {
        update-installPath
    }
}
if (!(Test-Path $configObj.TempPath)) {
    Write-Host -ForegroundColor Red "`n`tInvalid Temp Path"
    if ($batch) {
        exit 1
    }
    else {
        update-tempPath
    }
}
if (($configObj.WowVers -notmatch "^[0-9]{1,2}\.[0-9]{1,2}\.[0-9]{1,2}$") -and ($configObj.DownloadOld -eq 0)) {
    Write-Host -ForegroundColor Red "`n`tInvalid WoW Release Version"
    if ($batch) {
        exit 1
    }
    else {
        update-wowVers
    }
}
if ($batch) {
    install-Addons $addon
    Stop-Transcript
}
else {
    invoke-Menu
}