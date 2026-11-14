[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$ApplicationName,
    [Parameter(Mandatory=$true)][string]$Version,
    [Parameter(Mandatory=$true)][string]$Vendor,
    [ValidateSet('MSI','EXE')][string]$InstallerType = 'MSI',
    [ValidateSet('x86','x64','ARM64')][string]$Architecture = 'x64',
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [string]$InstallFile,
    [string]$InstallParameters = '',
    [ValidateSet('Main EXE / Installer','Uninstall from different location','MSI ProductCode')]
    [string]$UninstallMethod = 'Main EXE / Installer',
    [string]$UninstallFile = '',
    [string]$UninstallParameters = '',
    [switch]$UseRegistryUninstall,
    [string]$PreInstallSteps = '',
    [string]$InstallSteps = '',
    [string]$PostInstallSteps = '',
    [string]$PreUninstallSteps = '',
    [string]$UninstallSteps = '',
    [string]$PostUninstallSteps = '',
    [string]$InstallCloseApps = '',
    [string]$UninstallCloseApps = '',
    [switch]$GenerateMST,
    [switch]$GeneratePID,
    [switch]$IncludeServiceUI,
    [switch]$GenerateIntuneWin,
    [switch]$PrepareOnly,
    [switch]$ContinuePrepared,
    [string]$TemplateRoot = '',
    [string]$OutputRoot = '',
    [string]$AutomationLogRoot = 'C:\Temp\AutomationLogs',
    [ValidateSet('3.10','4.x')][string]$PSADTVersion = '3.10',
    [bool]$RemoveDesktopShortcut = $false,
    [string]$AppLogoSource = ''
)

$ErrorActionPreference = 'Stop'

# V1 internal folder structure.
# Script location: <ToolRoot>\Scripts\Invoke-PackagingEngine.ps1
$ToolRoot = Split-Path -Parent $PSScriptRoot
$DefaultTemplateRoot = Join-Path $ToolRoot 'Templates'
$DefaultOutputRoot   = Join-Path $ToolRoot 'Output'

if([string]::IsNullOrWhiteSpace($TemplateRoot)){
    $TemplateRoot = $DefaultTemplateRoot
}
if([string]::IsNullOrWhiteSpace($OutputRoot)){
    $OutputRoot = $DefaultOutputRoot
}

# Resolve relative paths against the tool root.
if(-not [IO.Path]::IsPathRooted($TemplateRoot)){
    $TemplateRoot = Join-Path $ToolRoot $TemplateRoot
}
if(-not [IO.Path]::IsPathRooted($OutputRoot)){
    $OutputRoot = Join-Path $ToolRoot $OutputRoot
}
$TemplateRoot = [IO.Path]::GetFullPath($TemplateRoot)
$OutputRoot   = [IO.Path]::GetFullPath($OutputRoot)
$Vendor = ($Vendor.Trim() -replace '\s+', '-')
$ApplicationName = ($ApplicationName.Trim() -replace '\s+', '-')
$Version = ($Version.Trim() -replace '\s+', '-')
$PackageName = if($Architecture -eq 'x86'){ '{0}_{1}_{2}_P1' -f $Vendor,$ApplicationName,$Version } else { '{0}_{1}_{2}_P1_{3}' -f $Vendor,$ApplicationName,$Version,$Architecture }
$PackageRoot = Join-Path $OutputRoot $PackageName
$PackageDir = Join-Path $PackageRoot 'Package'
$DocDir = Join-Path $PackageRoot 'Docs'
$IntuneDir = Join-Path $PackageRoot 'Intune'
# Automation logs are kept in the fixed global folder requested by the user.
$RunLogDir = $AutomationLogRoot
$EngineParameterSummary = @(
    "ApplicationName=$ApplicationName"
    "Version=$Version"
    "Vendor=$Vendor"
    "InstallerType=$InstallerType"
    "Architecture=$Architecture"
    "PSADTVersion=$PSADTVersion"
    "RemoveDesktopShortcut=$RemoveDesktopShortcut"
    "AppLogoSource=$AppLogoSource"
    "SourcePath=$SourcePath"
    "InstallFile=$InstallFile"
    "UninstallMethod=$UninstallMethod"
    "UninstallFile=$UninstallFile"
    "UninstallParameters=$UninstallParameters"
    "TemplateRoot=$TemplateRoot"
    "OutputRoot=$OutputRoot"
    "AutomationLogRoot=$AutomationLogRoot"
) -join [Environment]::NewLine

New-Item -ItemType Directory -Force -Path $PackageDir,$DocDir,$IntuneDir,$RunLogDir | Out-Null
$LogPath = Join-Path $RunLogDir "$PackageName.log"
$RuntimeLog = $LogPath

function Write-AutoLog {
    param([string]$Message,[ValidateSet('INFO','SUCCESS','WARNING','ERROR')][string]$Level='INFO')
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Add-Content -LiteralPath $LogPath -Value $line
    Add-Content -LiteralPath $RuntimeLog -Value $line
    Write-Host $line
}


function Get-MsiProperty {
    param([Parameter(Mandatory=$true)][string]$MsiPath,[Parameter(Mandatory=$true)][string]$PropertyName)
    $installer = $null
    $database = $null
    $view = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database = $installer.OpenDatabase($MsiPath, 0)
        $view = $database.OpenView("SELECT `Value` FROM Property WHERE `Property`='$PropertyName'")
        $view.Execute()
        $record = $view.Fetch()
        if($record){ return ([string]$record.StringData(1)).Trim() }
        return [string]''
    } finally {
        # Suppress every COM cleanup return value.  If these values are allowed
        # to escape the function, PowerShell can turn the function result into
        # an Object[]; that later produces values such as System.Object[] in
        # the generated Execute-MSI uninstall command.
        try { if($view){ [void]$view.Close() } } catch {}
        try { if($database){ [void][Runtime.InteropServices.Marshal]::ReleaseComObject($database) } } catch {}
        try { if($installer){ [void][Runtime.InteropServices.Marshal]::ReleaseComObject($installer) } } catch {}
    }
}

function Find-PSADTTemplate {
   param([string]$Root)
   # Check the root itself
   if (
       (Test-Path -LiteralPath (Join-Path $Root 'AppDeployToolkit') -PathType Container) -and
       (Test-Path -LiteralPath (Join-Path $Root 'Deploy-Application.ps1') -PathType Leaf)
   ) {
       return $Root
   }
   # Search recursively for the PSADT template
   $candidates = Get-ChildItem -LiteralPath $Root -Directory -Recurse -ErrorAction SilentlyContinue
   foreach ($d in $candidates) {
       $toolkitFolder = Join-Path $d.FullName 'AppDeployToolkit'
       $deployScript  = Join-Path $d.FullName 'Deploy-Application.ps1'
       if (
           (Test-Path -LiteralPath $toolkitFolder -PathType Container) -and
           (Test-Path -LiteralPath $deployScript -PathType Leaf)
       ) {
           return $d.FullName
       }
   }
   throw "PSADT template folder was not found under $Root. Expected Deploy-Application.ps1 and AppDeployToolkit."
}

function Insert-PSADTSectionContent {
    param(
        [string]$FilePath,
        [string]$Marker,
        [string]$Block,
        [string]$Tag
    )

    if([string]::IsNullOrWhiteSpace($Block)){ return }

    $text = Get-Content -LiteralPath $FilePath -Raw -ErrorAction Stop

    # Insert directly into the existing PSADT lifecycle section.
    # The generated command(s) inherit the indentation of the template marker,
    # so Deploy-Application.ps1 keeps the original PSADT alignment.
    $idx = $text.IndexOf($Marker, [System.StringComparison]::OrdinalIgnoreCase)
    if($idx -lt 0){
        $markerRegex = '(?im)^[ \t]*#+[ \t]*' + [regex]::Escape($Marker.Substring(2)) + '[ \t]*$'
        $m = [regex]::Match($text, $markerRegex)
        if($m.Success){ $idx = $m.Index }
    }

    if($idx -lt 0){
        $fallbackRegex = switch -Regex ($Tag) {
            '^Pre-Installation$' {
                '(?im)^[ \t]*(?:##\*+[ \t]*INSTALLATION|##[ \t]*={3,}[ \t]*\r?\n[ \t]*##\*+[ \t]*INSTALLATION)'
            }
            '^Installation$' {
                '(?im)^[ \t]*(?:##\*+[ \t]*POST-INSTALLATION|##[ \t]*MARK:[ \t]*Post-Install)'
            }
            '^Post-Installation$' {
                '(?im)^[ \t]*(?:\}[ \t]*ElseIf\s*\(\$deploymentType\s+-ieq\s+[''\"]Uninstall[''\"]|##\*+[ \t]*PRE-UNINSTALLATION|##[ \t]*MARK:[ \t]*Pre-Uninstall)'
            }
            '^Pre-Uninstallation$' {
                '(?im)^[ \t]*(?:##\*+[ \t]*UNINSTALLATION|##[ \t]*MARK:[ \t]*Uninstall)'
            }
            '^Uninstallation$' {
                '(?im)^[ \t]*(?:##\*+[ \t]*POST-UNINSTALLATION|##[ \t]*MARK:[ \t]*Post-Uninstallation)'
            }
            '^Post-Uninstallation$' {
                '(?im)^[ \t]*(?:\}[ \t]*ElseIf\s*\(\$deploymentType|##\*+[ \t]*REPAIR|##[ \t]*MARK:[ \t]*Repair)'
            }
            default { $null }
        }
        if($fallbackRegex){
            $fm = [regex]::Match($text, $fallbackRegex)
            if($fm.Success){ $idx = $fm.Index }
        }
    }

    if($idx -lt 0){
        throw "PSADT section marker not found: $Marker`r`nTag: $Tag`r`nTemplate: $FilePath`r`nThe template is not a supported PSADT 3.10 Deploy-Application.ps1 layout."
    }

    # Find the complete marker line and use its leading whitespace as the
    # canonical indentation for the generated block.
    $lineStart = $idx
    while($lineStart -gt 0 -and $text[$lineStart-1] -ne "`n"){ $lineStart-- }
    $lineEnd = $text.IndexOf("`n", $idx)
    if($lineEnd -lt 0){ $lineEnd = $text.Length }
    $markerLine = $text.Substring($lineStart, $lineEnd - $lineStart).TrimEnd("`r")
    $baseIndent = [regex]::Match($markerLine, '^[ \t]*').Value

    # Normalize CR/LF and apply the marker's indentation to every generated
    # line, while preserving intentional relative indentation within a block.
    $blockLines = $Block -split "`r?`n", -1
    while($blockLines.Count -gt 0 -and [string]::IsNullOrWhiteSpace($blockLines[0])){ $blockLines = $blockLines[1..($blockLines.Count-1)] }
    while($blockLines.Count -gt 0 -and [string]::IsNullOrWhiteSpace($blockLines[$blockLines.Count-1])){ $blockLines = $blockLines[0..($blockLines.Count-2)] }
    if($blockLines.Count -eq 0){ return }

    $commonIndent = $null
    foreach($bl in $blockLines){
        if([string]::IsNullOrWhiteSpace($bl)){ continue }
        $ind = [regex]::Match($bl, '^[ \t]*').Value
        if($null -eq $commonIndent -or $ind.Length -lt $commonIndent.Length){ $commonIndent = $ind }
    }
    if($null -eq $commonIndent){ $commonIndent = '' }

    $normalized = foreach($bl in $blockLines){
        if([string]::IsNullOrWhiteSpace($bl)){ '' }
        else {
            $relative = $bl
            if($commonIndent.Length -gt 0 -and $relative.StartsWith($commonIndent)){ $relative = $relative.Substring($commonIndent.Length) }
            $baseIndent + $relative
        }
    }
    $formattedBlock = $normalized -join "`r`n"

    # Insert immediately after the marker line, preserving the template's
    # existing lifecycle structure and indentation.
    $insert = "`r`n$formattedBlock"
    $text = $text.Insert($lineEnd + 1, $insert)
    Set-Content -LiteralPath $FilePath -Value $text -Encoding UTF8
}

function Set-PSADTVariableValue {
    param(
        [string]$FilePath,
        [string]$Name,
        [string]$Value
    )

    $text = Get-Content -LiteralPath $FilePath -Raw -ErrorAction Stop
    $exactValue = [string]$Value
    $escapedValue = $exactValue -replace "'", "''"
    $pattern = '(?im)^(?<indent>[ \t]*)(?:\[[^\]]+\][ \t]*)?\$' + [regex]::Escape($Name) + '\s*=\s*.*$'

    $m = [regex]::Match($text,$pattern)
    if($m.Success){
        $indent = $m.Groups['indent'].Value
        $line = $indent + "[String]`$$Name = '$escapedValue'"
        $text = [regex]::Replace($text,$pattern,[System.Text.RegularExpressions.MatchEvaluator]{param($match) $line},1)
    } else {
        $anchor = '(?im)^(?<indent>[ \t]*)##\s*Variables:\s*Application\s*$'
        $a = [regex]::Match($text,$anchor)
        if($a.Success){
            $lineEnd = $text.IndexOf("`n",$a.Index)
            if($lineEnd -lt 0){$lineEnd=$text.Length-1}
            $line = $a.Groups['indent'].Value + "[String]`$$Name = '$escapedValue'"
            $text=$text.Insert($lineEnd+1, "`r`n$line")
        } else {
            $text = "[String]`$$Name = '$escapedValue'`r`n$text"
        }
    }
    Set-Content -LiteralPath $FilePath -Value $text -Encoding UTF8
}

function Update-ExistingShowInstallationWelcomeCloseApps {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [string]$CloseApps = '',
        [ValidateSet('Pre-Installation','Pre-Uninstallation')][string]$SectionName = 'Pre-Installation'
    )
    $text = Get-Content -LiteralPath $FilePath -Raw -ErrorAction Stop
    $lines = $text -split "`r?`n", -1
    if($SectionName -eq 'Pre-Installation'){
        $startPatterns=@('(?i)^\s*##\*\s*PRE-INSTALLATION\s*$','(?i)^\s*##\s*PRE-INSTALLATION\s*$')
        $endPattern='(?i)^\s*##\s*<Perform Installation tasks here>\s*$'
    } else {
        $startPatterns=@('(?i)^\s*##\*\s*PRE-UNINSTALLATION\s*$','(?i)^\s*##\s*PRE-UNINSTALLATION\s*$')
        $endPattern='(?i)^\s*##\s*<Perform Uninstallation tasks here>\s*$'
    }
    $start=-1
    for($i=0;$i -lt $lines.Count;$i++){ foreach($p in $startPatterns){ if($lines[$i] -match $p){$start=$i;break} }; if($start -ge 0){break} }
    if($start -lt 0){throw "PSADT $SectionName section marker was not found in $FilePath."}
    $end=$lines.Count
    for($i=$start+1;$i -lt $lines.Count;$i++){if($lines[$i] -match $endPattern){$end=$i;break}}
    $commandIndex=-1
    for($i=$start;$i -lt $end;$i++){if($lines[$i] -match '(?i)\bShow-InstallationWelcome\b'){$commandIndex=$i;break}}
    if($commandIndex -lt 0){throw "The PSADT template does not contain Show-InstallationWelcome in the $SectionName section."}
    $line=$lines[$commandIndex]
    if([string]::IsNullOrWhiteSpace($CloseApps)){
        if($line -notmatch '^\s*#'){
            $indent=[regex]::Match($line,'^\s*').Value
            $lines[$commandIndex]=$indent+'# '+$line.Substring($indent.Length)
            Write-AutoLog "CloseApps blank for $SectionName. Existing Show-InstallationWelcome commented." 'INFO'
        }
    } else {
        $value=$CloseApps.Trim().Replace("'","''")
        if($line -match '(?i)(-CloseApps\s+)(?:''[^'']*''|"[^"]*"|[^\s]+)'){
            $replacement='${1}'+"'"+$value+"'"
            $newLine=[regex]::Replace($line,'(?i)(-CloseApps\s+)(?:''[^'']*''|"[^"]*"|[^\s]+)',$replacement,1)
        } else {
            $newLine=$line -replace '(?i)(\bShow-InstallationWelcome\b)',('$1 -CloseApps '+"'"+$value+"'")
        }
        $lines[$commandIndex]=$newLine
        Write-AutoLog "Show-InstallationWelcome updated only in $SectionName with CloseApps='$CloseApps'." 'SUCCESS'
    }
    Set-Content -LiteralPath $FilePath -Value ($lines -join "`r`n") -Encoding UTF8
}

function Test-PSADTTemplateLayout {
    param([Parameter(Mandatory=$true)][string]$FilePath)

    $text = Get-Content -LiteralPath $FilePath -Raw -ErrorAction Stop
    $required = @(
        'Pre-Installation',
        'Installation',
        'Post-Installation',
        'Pre-Uninstallation',
        'Uninstallation',
        'Post-Uninstallation'
    )

    $missing = New-Object System.Collections.Generic.List[string]
    foreach($name in $required){
        $needle = '<Perform ' + $name + ' tasks here>'
        if($text.IndexOf($needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0){ continue }

        $lifecycle = switch($name){
            'Pre-Installation' { 'INSTALLATION' }
            'Installation' { 'POST-INSTALLATION' }
            'Post-Installation' { 'PRE-UNINSTALLATION' }
            'Pre-Uninstallation' { 'UNINSTALLATION' }
            'Uninstallation' { 'POST-UNINSTALLATION' }
            'Post-Uninstallation' { 'REPAIR' }
        }
        $pattern = '(?im)^[ \t]*(?:##\*+[ \t]*' + [regex]::Escape($lifecycle) + '|##[ \t]*MARK:[ \t]*' + [regex]::Escape($lifecycle) + ')'
        if(-not [regex]::IsMatch($text, $pattern)){ $missing.Add($name) }
    }

    if($missing.Count -gt 0){
        throw "PSADT 3.10 template layout validation failed. Missing lifecycle sections: $($missing -join ', ').`r`nTemplate: $FilePath"
    }
}

function Get-RelativeFileName {
    param([string]$Base,[string]$File)
    if([string]::IsNullOrWhiteSpace($File)){ return '' }
    if([IO.Path]::IsPathRooted($File)){
        $fullBase = [IO.Path]::GetFullPath($Base).TrimEnd('\') + '\'
        $fullFile = [IO.Path]::GetFullPath($File)
        if($fullFile.StartsWith($fullBase,[StringComparison]::OrdinalIgnoreCase)){ return $fullFile.Substring($fullBase.Length) }
        return [IO.Path]::GetFileName($fullFile)
    }
    return $File
}


function Save-AppLogo {
    param(
        [string]$Source,
        [string]$IntuneDirectory,
        [string]$ToolkitDirectory
    )

    if([string]::IsNullOrWhiteSpace($Source)){ return }
    if(-not (Test-Path -LiteralPath $Source -PathType Leaf)){
        throw "AppLogo source file was not found: $Source"
    }

    Add-Type -AssemblyName System.Drawing
    $icon = $null
    $bitmap = $null
    try {
        $ext = [IO.Path]::GetExtension($Source).ToLowerInvariant()
        if($ext -eq '.ico'){
            $icon = New-Object System.Drawing.Icon($Source)
        } else {
            $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Source)
        }
        if($null -eq $icon){
            throw "No extractable icon was found in: $Source"
        }
        $bitmap = $icon.ToBitmap()
        $intuneLogo = Join-Path $IntuneDirectory 'AppLogo.png'
        $toolkitLogo = Join-Path $ToolkitDirectory 'AppLogo.png'
        $bitmap.Save($intuneLogo, [System.Drawing.Imaging.ImageFormat]::Png)
        if(Test-Path -LiteralPath $ToolkitDirectory -PathType Container){
            $bitmap.Save($toolkitLogo, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        Write-AutoLog "AppLogo.png created: $intuneLogo" 'SUCCESS'
        if(Test-Path -LiteralPath $ToolkitDirectory -PathType Container){
            Write-AutoLog "AppLogo.png created: $toolkitLogo" 'SUCCESS'
        }
    } finally {
        if($bitmap){ $bitmap.Dispose() }
        if($icon){ $icon.Dispose() }
    }
}

try {
    Write-AutoLog '============================================================'
    Write-AutoLog 'APPLICATION PACKAGING AUTOMATION'
    Write-AutoLog "Package: $PackageName"
    Write-AutoLog "Installer type: $InstallerType"
    Write-AutoLog '============================================================'

    if(-not $ContinuePrepared){
    if(-not (Test-Path -LiteralPath $SourcePath -PathType Container)){ throw "Source folder not found: $SourcePath" }
    if(-not (Test-Path -LiteralPath $TemplateRoot -PathType Container)){ throw "Template folder not found: $TemplateRoot" }
    $pidTemplate = Get-ChildItem -LiteralPath $TemplateRoot -Filter 'Package_Template.docx' -File -Recurse -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
    if([string]::IsNullOrWhiteSpace($pidTemplate)){
        $pidTemplate = Get-ChildItem -LiteralPath $TemplateRoot -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -ieq '.docx' -and ($_.Name -match 'PID|Package|Template') } |
            Select-Object -First 1 -ExpandProperty FullName
    }
    if([string]::IsNullOrWhiteSpace($pidTemplate)){
        $pidTemplate = Get-ChildItem -LiteralPath $TemplateRoot -Filter '*.docx' -File -Recurse -ErrorAction SilentlyContinue |
            Select-Object -First 1 -ExpandProperty FullName
    }
    if($GeneratePID -and [string]::IsNullOrWhiteSpace($pidTemplate)){ throw "Package_Template.docx was not found anywhere under $TemplateRoot." }

    $psadtVersionRoot = Join-Path (Join-Path $TemplateRoot 'PSADT') $PSADTVersion
    if(-not(Test-Path -LiteralPath $psadtVersionRoot)){ throw "PSADT template version '$PSADTVersion' was not found: $psadtVersionRoot" }
    $psadtSource = Find-PSADTTemplate -Root $psadtVersionRoot
    Write-AutoLog "PSADT $PSADTVersion template: $psadtSource" 'SUCCESS'

    # Copy every item from the selected PSADT template root explicitly, except
    # ServiceUI.exe. ServiceUI is opt-in and must never arrive through the generic
    # template copy when the checkbox is not selected.
    Get-ChildItem -LiteralPath $psadtSource -Force -ErrorAction Stop | ForEach-Object {
        if($_.Name -ieq 'ServiceUI.exe'){
            return
        }
        Copy-Item -LiteralPath $_.FullName -Destination $PackageDir -Recurse -Force -ErrorAction Stop
    }

    # Remove a stale ServiceUI.exe from an existing output package when the option
    # is not selected. This makes the opt-in behaviour deterministic across rebuilds.
    $serviceUiTarget = Join-Path $PackageDir 'ServiceUI.exe'
    if(-not $IncludeServiceUI -and (Test-Path -LiteralPath $serviceUiTarget -PathType Leaf)){
        Remove-Item -LiteralPath $serviceUiTarget -Force -ErrorAction Stop
        Write-AutoLog 'ServiceUI.exe removed from package because IncludeServiceUI was not selected.' 'INFO'
    }

    $deployFromTemplate = Join-Path $psadtSource 'Deploy-Application.ps1'
    $deployInPackage = Join-Path $PackageDir 'Deploy-Application.ps1'
    if(-not (Test-Path -LiteralPath $deployInPackage -PathType Leaf)){
        if(Test-Path -LiteralPath $deployFromTemplate -PathType Leaf){
            Copy-Item -LiteralPath $deployFromTemplate -Destination $deployInPackage -Force -ErrorAction Stop
        }
    }
    if(-not (Test-Path -LiteralPath $deployInPackage -PathType Leaf)){
        throw "PSADT template was detected at '$psadtSource' but Deploy-Application.ps1 could not be copied to '$PackageDir'. Verify that the template contains Deploy-Application.ps1 and that the destination is writable."
    }

    # Optional ServiceUI.exe: copy only when explicitly requested.
    if($IncludeServiceUI){
        $serviceUiTemplate = Get-ChildItem -LiteralPath $psadtSource -File -Recurse -Filter 'ServiceUI.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
        if(-not $serviceUiTemplate){
            throw "ServiceUI.exe was requested, but it was not found under the selected PSADT template: $psadtSource"
        }
        Copy-Item -LiteralPath $serviceUiTemplate.FullName -Destination (Join-Path $PackageDir 'ServiceUI.exe') -Force -ErrorAction Stop
        Write-AutoLog "ServiceUI.exe copied to Package root from template: $($serviceUiTemplate.FullName)" 'SUCCESS'
    }

    $filesDir = Join-Path $PackageDir 'Files'
    New-Item -ItemType Directory -Force -Path $filesDir | Out-Null
    Copy-Item (Join-Path $SourcePath '*') $filesDir -Recurse -Force
    Write-AutoLog 'Vendor source copied to Package\Files.' 'SUCCESS'

    # Never keep an AutomationLogs folder inside the generated package.
    foreach($localLogDir in @(Get-ChildItem -LiteralPath $PackageDir -Directory -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq 'AutomationLogs' })){
        try { Remove-Item -LiteralPath $localLogDir.FullName -Recurse -Force -ErrorAction Stop } catch { Write-AutoLog ("Could not remove template/source AutomationLogs folder: {0}" -f $_.Exception.Message) 'WARNING' }
    }

    $installer = Get-RelativeFileName -Base $SourcePath -File $InstallFile
    if([string]::IsNullOrWhiteSpace($installer)){
        $found = Get-ChildItem -LiteralPath $filesDir -File | Where-Object {$_.Extension -in '.msi','.exe'} | Select-Object -First 1
        if(-not $found){ throw 'No MSI/EXE installer was found in the source folder. Specify Installer File.' }
        $installer = $found.Name
    }
    $installerFull = Join-Path $filesDir $installer
    if(-not (Test-Path -LiteralPath $installerFull -PathType Leaf)){ throw "Installer file not found after source copy: $installer" }
    Write-AutoLog "Installer: $installer" 'SUCCESS'

    $msiProductCode = ''
    $msiProductName = ''
    $msiProductVersion = ''
    $msiManufacturer = ''
    $msiUpgradeCode = ''

    if($InstallerType -eq 'MSI'){
        # Force each MSI property to a scalar string. This prevents any COM
        # cleanup/output from becoming an Object[] and guarantees the ProductCode
        # is written into the generated uninstall command.
        # Normalize every MSI property immediately. Some MSI COM calls can return
        # leading/trailing whitespace; ProductCode must never carry that whitespace
        # into the generated Execute-MSI command.
        $msiProductCode = ([string](Get-MsiProperty -MsiPath $installerFull -PropertyName 'ProductCode')).Trim()
        $msiProductName = ([string](Get-MsiProperty -MsiPath $installerFull -PropertyName 'ProductName')).Trim()
        $msiProductVersion = ([string](Get-MsiProperty -MsiPath $installerFull -PropertyName 'ProductVersion')).Trim()
        $msiManufacturer = ([string](Get-MsiProperty -MsiPath $installerFull -PropertyName 'Manufacturer')).Trim()
        $msiUpgradeCode = ([string](Get-MsiProperty -MsiPath $installerFull -PropertyName 'UpgradeCode')).Trim()
        if([string]::IsNullOrWhiteSpace($msiProductCode)) { throw "Unable to read ProductCode directly from MSI: $installerFull" }
        if($msiProductCode -notmatch '^\{[0-9A-Fa-f-]{36}\}$'){
            throw "Invalid MSI ProductCode detected: '$msiProductCode'"
        }
        Write-AutoLog "MSI ProductCode auto-detected: $msiProductCode" 'SUCCESS'
        Write-AutoLog "MSI ProductName: $msiProductName" 'INFO'
        Write-AutoLog "MSI ProductVersion: $msiProductVersion" 'INFO'
        Write-AutoLog "MSI Manufacturer: $msiManufacturer" 'INFO'
        Write-AutoLog "MSI UpgradeCode: $msiUpgradeCode" 'INFO'
    }

    # An EXE bootstrapper may install an embedded MSI. In that case its uninstall
    # registry entry can legitimately be an MSI ProductCode. Treat uninstall
    # method independently from the installer file type.
    if($InstallerType -eq 'EXE' -and $UninstallMethod -eq 'MSI ProductCode'){
        $msiProductCode = ([string]$UninstallFile).Trim()
        if([string]::IsNullOrWhiteSpace($msiProductCode)){
            throw 'MSI ProductCode uninstall was selected for this EXE package, but the ProductCode field is empty.'
        }
        if($msiProductCode -notmatch '^\{[0-9A-Fa-f-]{36}\}$'){
            throw "Invalid MSI ProductCode supplied for EXE uninstall: '$msiProductCode'"
        }
        Write-AutoLog "EXE installer uses MSI ProductCode for uninstall: $msiProductCode" 'SUCCESS'
    }

    if($InstallerType -eq 'MSI'){
        if([IO.Path]::GetExtension($installerFull) -ine '.msi'){ throw 'Installer Type is MSI but the selected installer is not an MSI.' }
        if($GenerateMST){
            # V1 standard name is Create-StandardMST.ps1. Keep backward
            # compatibility with the previous Schema300 filename.
            $mstScript = Join-Path $PSScriptRoot 'Create-StandardMST.ps1'
            if(-not (Test-Path -LiteralPath $mstScript -PathType Leaf)){
                $mstScript = Join-Path $PSScriptRoot 'Create-StandardMST-Schema300.ps1'
            }
            if(-not (Test-Path -LiteralPath $mstScript -PathType Leaf)){
                throw 'Create-StandardMST.ps1 was not found.'
            }
            Write-AutoLog ("Calling Standard MST script: {0}" -f $mstScript)

            # Compatibility: different Standard MST script revisions may not
            # expose the optional -LogDirectory parameter. Do not pass an
            # unsupported named parameter, otherwise PowerShell aborts before
            # the MST generator can run.
            $mstCommand = Get-Command -Name $mstScript -ErrorAction Stop
            $mstParams = @{
                MsiPath      = $installerFull
                Vendor       = $Vendor
                AppName      = $ApplicationName
                Version      = $Version
                Architecture = $Architecture
            }

            if($mstCommand.Parameters.ContainsKey('LogDirectory')){
                $mstParams['LogDirectory'] = $RunLogDir
                Write-AutoLog 'Standard MST script supports -LogDirectory; package log directory supplied.'
            } else {
                Write-AutoLog 'Standard MST script does not expose -LogDirectory; continuing without it.' 'INFO'
            }

            & $mstScript @mstParams
            # The supplied Standard MST script always includes the architecture
            # in the MST filename, even when the x86 package root omits it.
            $mstFileName = '{0}_{1}_{2}_P1_{3}.mst' -f $Vendor,$ApplicationName,$Version,$Architecture
            $expectedMst = Join-Path $filesDir $mstFileName
            if(-not (Test-Path -LiteralPath $expectedMst -PathType Leaf)){
                throw "Standard MST script did not create the expected MST: $expectedMst"
            }
            Write-AutoLog 'Standard MST generated successfully.' 'SUCCESS'
        }
    } else {
        if([IO.Path]::GetExtension($installerFull) -ine '.exe'){ throw 'Installer Type is EXE but the selected installer is not an EXE.' }
    }


    $deploy = Join-Path $PackageDir 'Deploy-Application.ps1'
    if(-not (Test-Path -LiteralPath $deploy -PathType Leaf)){ throw "Deploy-Application.ps1 was not copied to the package root from PSADT template '$psadtSource'." }
    Test-PSADTTemplateLayout -FilePath $deploy

    # Populate the standard PSADT Application variables in the generated
    # Deploy-Application.ps1. These values are derived from the UI/MSI and
    # are not required as manual input.
    # PSADT application variables are human-readable values. Keep the original
    # UI values for package/folder naming, but convert common package separators
    # (hyphen/underscore) to spaces when writing appVendor/appName.
    # Example: keyshot-studio -> keyshot studio.
    # IMPORTANT: normalize the values before writing them to Deploy-Application.ps1.
    # Package/folder names continue to use the original UI values.
    function ConvertTo-PSADTDisplayValue {
        param([AllowNull()][string]$Value)
        if($null -eq $Value){ return '' }
        $normalized = $Value.Trim()
        $normalized = $normalized -replace '[-_]+', ' '
        $normalized = $normalized -replace '\s+', ' '
        return $normalized.Trim()
    }

    $psadtVendor = ConvertTo-PSADTDisplayValue -Value $Vendor
    $psadtAppName = ConvertTo-PSADTDisplayValue -Value $ApplicationName

    Write-AutoLog ("PSADT display values: appVendor='{0}', appName='{1}' (package name remains '{2}')." -f $psadtVendor,$psadtAppName,$PackageName) 'INFO'
    Set-PSADTVariableValue -FilePath $deploy -Name 'appVendor' -Value $psadtVendor
    Set-PSADTVariableValue -FilePath $deploy -Name 'appName' -Value $psadtAppName

    # Verify the values written to the generated script. This prevents the package
    # from silently continuing with the original hyphenated UI values.
    $verifyText = Get-Content -LiteralPath $deploy -Raw -ErrorAction Stop
    $verifyAppName = [regex]::Match($verifyText,"(?im)^\s*\[String\]\`$appName\s*=\s*'([^']*)'\s*$")
    $verifyAppVendor = [regex]::Match($verifyText,"(?im)^\s*\[String\]\`$appVendor\s*=\s*'([^']*)'\s*$")
    if(-not $verifyAppName.Success -or $verifyAppName.Groups[1].Value -cne $psadtAppName){
        throw "Deploy-Application.ps1 appName verification failed. Expected '$psadtAppName'."
    }
    if(-not $verifyAppVendor.Success -or $verifyAppVendor.Groups[1].Value -cne $psadtVendor){
        throw "Deploy-Application.ps1 appVendor verification failed. Expected '$psadtVendor'."
    }
    Write-AutoLog 'PSADT appName/appVendor values verified in Deploy-Application.ps1.' 'SUCCESS'
    Set-PSADTVariableValue -FilePath $deploy -Name 'appVersion' -Value $Version
    Set-PSADTVariableValue -FilePath $deploy -Name 'appArch' -Value $Architecture
    Set-PSADTVariableValue -FilePath $deploy -Name 'appLang' -Value 'EN'
    Set-PSADTVariableValue -FilePath $deploy -Name 'appRevision' -Value 'P1'
    # PSADT installName/installTitle must be the complete package name.
    Set-PSADTVariableValue -FilePath $deploy -Name 'installName' -Value $PackageName
    Set-PSADTVariableValue -FilePath $deploy -Name 'installTitle' -Value $PackageName
    Set-PSADTVariableValue -FilePath $deploy -Name 'appScriptVersion' -Value '1.0.0'
    Set-PSADTVariableValue -FilePath $deploy -Name 'appScriptDate' -Value (Get-Date -Format 'MM/dd/yyyy')
    Set-PSADTVariableValue -FilePath $deploy -Name 'appScriptAuthor' -Value 'Dinesh Dalavayi'
    Write-AutoLog 'PSADT Application variables populated.' 'SUCCESS'
    $mstFileName = '{0}_{1}_{2}_P1_{3}.mst' -f $Vendor,$ApplicationName,$Version,$Architecture

    function ConvertTo-PSADTPath([string]$Path) {
        if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
        $p = $Path.Trim().Replace('/', '\')
        if ($p -notmatch '^[A-Za-z]:\\') { return $p }

        # Use literal, case-insensitive prefix checks. Do not build regular
        # expressions from Windows paths: sequences such as \ProgramData can
        # be interpreted as invalid regex escapes (for example \p).
        $rules = @(
            @{ Prefix='ProgramData\Microsoft\Windows\Start Menu\Programs'; Replacement='$envCommonStartMenuPrograms' },
            @{ Prefix='Users\Public\Desktop'; Replacement='$envCommonDesktop' },
            @{ Prefix='Program Files (x86)'; Replacement='$envProgramFilesX86' },
            @{ Prefix='Program Files'; Replacement='$envProgramFiles' },
            @{ Prefix='ProgramData'; Replacement='$envProgramData' },
            @{ Prefix='Windows'; Replacement='$envWinDir' }
        )
        $root = $p.Substring(3)
        foreach ($rule in $rules) {
            $prefix = [string]$rule.Prefix
            if ($root.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) -and
                ($root.Length -eq $prefix.Length -or $root[$prefix.Length] -eq '\')) {
                return ([string]$rule.Replacement) + $root.Substring($prefix.Length)
            }
        }
        return $p
    }

    # Build PSADT commands according to the selected PSADT version.
    # PSADT 3.10 uses Execute-MSI / Execute-Process.
    # PSADT 4.x uses Start-ADTProcess. MSI operations are executed through msiexec.exe
    # so the generated command does not depend on a removed v3 cmdlet.
    if($PSADTVersion -eq '3.10'){
        if($InstallerType -eq 'MSI'){
            $effectiveInstallParameters = if([string]::IsNullOrWhiteSpace($InstallParameters)){ '/qn /norestart' } else { $InstallParameters.Trim() }
            if($GenerateMST){
                $installPsadt = 'Execute-MSI -Action Install -Path "$dirFiles\{0}" -Transform "$dirFiles\{1}" -Parameters "{2}"' -f $installer,$mstFileName,$effectiveInstallParameters
            } else {
                $installPsadt = 'Execute-MSI -Action Install -Path "$dirFiles\{0}" -Parameters "{1}"' -f $installer,$effectiveInstallParameters
            }
        } else {
            if([string]::IsNullOrWhiteSpace($InstallParameters)){
                $installPsadt = 'Execute-Process -Path "$dirFiles\{0}" -WindowStyle Hidden -Wait' -f $installer
            } else {
                $installPsadt = 'Execute-Process -Path "$dirFiles\{0}" -Parameters "{1}" -WindowStyle Hidden -Wait' -f $installer,$InstallParameters
            }
        }

        if($InstallerType -eq 'MSI'){
            $UninstallMethod = 'MSI ProductCode'
            $msiProductCode = $msiProductCode.Trim()
            if([string]::IsNullOrWhiteSpace($msiProductCode)){ throw 'MSI uninstall cannot be generated because ProductCode is empty.' }
            if($msiProductCode -notmatch '^\{[0-9A-Fa-f-]{36}\}$'){ throw "MSI uninstall cannot be generated because ProductCode is invalid: '$msiProductCode'" }
            $UninstallParameters = '/qn /norestart'
            $uninstallPsadt = 'Execute-MSI -Action Uninstall -Path "{0}" -Parameters "{1}"' -f $msiProductCode,$UninstallParameters
        } else {
            $uninstallPsadt = switch($UninstallMethod){
                'MSI ProductCode' {
                    if([string]::IsNullOrWhiteSpace($msiProductCode)){ throw 'MSI ProductCode is required for uninstall.' }
                    'Execute-MSI -Action Uninstall -Path "{0}" -Parameters "/qn /norestart"' -f $msiProductCode
                }
                'Main EXE / Installer' {
                    if([string]::IsNullOrWhiteSpace($UninstallParameters)){
                        'Execute-Process -Path "$dirFiles\{0}" -WindowStyle Hidden -Wait' -f $installer
                    } else {
                        'Execute-Process -Path "$dirFiles\{0}" -Parameters "{1}" -WindowStyle Hidden -Wait' -f $installer,$UninstallParameters.Trim()
                    }
                }
                'Uninstall from different location' {
                    if([string]::IsNullOrWhiteSpace($UninstallFile)){throw 'Uninstall executable/path is required.'}
                    $uninstallPath = ConvertTo-PSADTPath $UninstallFile.Trim()
                    if(-not $uninstallPath.StartsWith('$env', [System.StringComparison]::OrdinalIgnoreCase) -and -not [IO.Path]::IsPathRooted($uninstallPath)){ $uninstallPath = '$dirFiles\' + $uninstallPath }
                    if([string]::IsNullOrWhiteSpace($UninstallParameters)){
                        'Execute-Process -Path "{0}" -WindowStyle Hidden -Wait' -f $uninstallPath
                    } else {
                        'Execute-Process -Path "{0}" -Parameters "{1}" -WindowStyle Hidden -Wait' -f $uninstallPath,$UninstallParameters.Trim()
                    }
                }
                default { throw "Unsupported Uninstall Method: $UninstallMethod" }
            }
        }
    }
    else {
        $filePrefix = '$PSScriptRoot\Files\'
        if($InstallerType -eq 'MSI'){
            $effectiveInstallParameters = if([string]::IsNullOrWhiteSpace($InstallParameters)){ '/qn /norestart' } else { $InstallParameters.Trim() }
            $msiArgs = '/i "{0}{1}" {2}' -f $filePrefix,$installer,$effectiveInstallParameters
            if($GenerateMST){ $msiArgs += ' TRANSFORMS="{0}{1}"' -f $filePrefix,$mstFileName }
            $installPsadt = 'Start-ADTProcess -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList ''{0}'' -WindowStyle Hidden -Wait' -f $msiArgs.Replace("'","''")
            if([string]::IsNullOrWhiteSpace($msiProductCode)){ throw 'MSI ProductCode is required for PSADT 4.x uninstall.' }
            $uninstallPsadt = 'Start-ADTProcess -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList ''/x {0} /qn /norestart'' -WindowStyle Hidden -Wait' -f $msiProductCode
        }
        else {
            $installPath = $filePrefix + $installer
            $installPsadt = if([string]::IsNullOrWhiteSpace($InstallParameters)){
                'Start-ADTProcess -FilePath "{0}" -WindowStyle Hidden -Wait' -f $installPath
            } else {
                'Start-ADTProcess -FilePath "{0}" -ArgumentList ''{1}'' -WindowStyle Hidden -Wait' -f $installPath,$InstallParameters.Trim().Replace("'","''")
            }
            $uninstallPsadt = switch($UninstallMethod){
                'MSI ProductCode' {
                    if([string]::IsNullOrWhiteSpace($msiProductCode)){ throw 'MSI ProductCode is required for uninstall.' }
                    'Start-ADTProcess -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList ''/x {0} /qn /norestart'' -WindowStyle Hidden -Wait' -f $msiProductCode
                }
                'Main EXE / Installer' {
                    if([string]::IsNullOrWhiteSpace($UninstallParameters)){
                        'Start-ADTProcess -FilePath "{0}" -WindowStyle Hidden -Wait' -f $installPath
                    } else {
                        'Start-ADTProcess -FilePath "{0}" -ArgumentList ''{1}'' -WindowStyle Hidden -Wait' -f $installPath,$UninstallParameters.Trim().Replace("'","''")
                    }
                }
                'Uninstall from different location' {
                    if([string]::IsNullOrWhiteSpace($UninstallFile)){ throw 'Uninstall executable/path is required.' }
                    $up = ConvertTo-PSADTPath $UninstallFile.Trim()
                    if(-not $up.StartsWith('$env', [System.StringComparison]::OrdinalIgnoreCase) -and -not [IO.Path]::IsPathRooted($up)){ $up = '$PSScriptRoot\Files\' + $up }
                    if([string]::IsNullOrWhiteSpace($UninstallParameters)){
                        'Start-ADTProcess -FilePath "{0}" -WindowStyle Hidden -Wait' -f $up
                    } else {
                        'Start-ADTProcess -FilePath "{0}" -ArgumentList ''{1}'' -WindowStyle Hidden -Wait' -f $up,$UninstallParameters.Trim().Replace("'","''")
                    }
                }
                default { throw "Unsupported Uninstall Method: $UninstallMethod" }
            }
        }
    }
    if($UseRegistryUninstall -and $InstallerType -eq 'EXE' -and $UninstallMethod -ne 'MSI ProductCode'){ $uninstallPsadt = '' ; Write-AutoLog 'EXE uninstall will use the registry-based command configured in the Uninstallation editor.' 'INFO' }
    Write-AutoLog "PSADT command syntax selected for version: $PSADTVersion" 'SUCCESS' 

    Write-AutoLog 'BUILD PHASE: PSADT CONFIGURATION COMPLETED; preparing PID document generation.' 'INFO'

    # Close Apps: update the Show-InstallationWelcome command that already exists
    # in the PSADT 3.10 Deploy-Application.ps1 template. In PSADT 3.10 the
    # Show-InstallationWelcome line is located AFTER the PRE-INSTALLATION /
    # PRE-UNINSTALLATION header but BEFORE the <Perform ... tasks here> marker.
    # Therefore the search must start at the lifecycle header, not at the task marker.
    # Do not create a second Show-InstallationWelcome command.
    # If no EXE is supplied, comment the existing command in both lifecycle sections.
    if($PSADTVersion -eq '3.10'){
        Update-ExistingShowInstallationWelcomeCloseApps -FilePath $deploy -CloseApps $InstallCloseApps -SectionName 'Pre-Installation'
    } elseif(-not [string]::IsNullOrWhiteSpace($InstallCloseApps)){
        Write-AutoLog 'PSADT 4.x CloseApps requires a compatible template implementation; no v3 Show-InstallationWelcome command was injected.' 'WARNING'
    }

    Insert-PSADTSectionContent -FilePath $deploy -Marker '## <Perform Pre-Installation tasks here>' -Block $PreInstallSteps -Tag 'Pre-Installation'
    $installationBlock = @($installPsadt, $InstallSteps) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
    $installationBlock = $installationBlock -join "`r`n`r`n"
    Insert-PSADTSectionContent -FilePath $deploy -Marker '## <Perform Installation tasks here>' -Block $installationBlock -Tag 'Installation'
    Insert-PSADTSectionContent -FilePath $deploy -Marker '## <Perform Post-Installation tasks here>' -Block $PostInstallSteps -Tag 'Post-Installation'

    if($PSADTVersion -eq '3.10'){
        Update-ExistingShowInstallationWelcomeCloseApps -FilePath $deploy -CloseApps $UninstallCloseApps -SectionName 'Pre-Uninstallation'
    } elseif(-not [string]::IsNullOrWhiteSpace($UninstallCloseApps)){
        Write-AutoLog 'PSADT 4.x Uninstall CloseApps requires a compatible template implementation.' 'WARNING'
    }

    Insert-PSADTSectionContent -FilePath $deploy -Marker '## <Perform Pre-Uninstallation tasks here>' -Block $PreUninstallSteps -Tag 'Pre-Uninstallation'
    $uninstallationBlock = @($uninstallPsadt, $UninstallSteps) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
    $uninstallationBlock = $uninstallationBlock -join "`r`n`r`n"
    Insert-PSADTSectionContent -FilePath $deploy -Marker '## <Perform Uninstallation tasks here>' -Block $uninstallationBlock -Tag 'Uninstallation'
    Insert-PSADTSectionContent -FilePath $deploy -Marker '## <Perform Post-Uninstallation tasks here>' -Block $PostUninstallSteps -Tag 'Post-Uninstallation'
    if($RemoveDesktopShortcut -and $InstallerType -eq 'EXE'){
        $shortcutBlock = @'
# Optional standard: remove desktop shortcuts matching the application name.
$desktopPaths = @("$env:PUBLIC\Desktop")
$desktopPaths += Get-ChildItem -Path 'C:\Users' -Directory -ErrorAction SilentlyContinue | ForEach-Object { Join-Path $_.FullName 'Desktop' }
Get-ChildItem -Path $desktopPaths -Filter '*.lnk' -ErrorAction SilentlyContinue |
    Where-Object { $_.BaseName -match [regex]::Escape($ApplicationName) } |
    Remove-Item -Force -ErrorAction SilentlyContinue
'@
        Insert-PSADTSectionContent -FilePath $deploy -Marker '## <Perform Post-Installation tasks here>' -Block $shortcutBlock -Tag 'Desktop Shortcut Removal'
        Write-AutoLog 'Optional desktop shortcut removal enabled.' 'INFO'
    }
    Write-AutoLog 'PSADT deployment sections configured.' 'SUCCESS'

    if(-not [string]::IsNullOrWhiteSpace($AppLogoSource)){
        $toolkitDir = Join-Path $PackageDir 'AppDeployToolkit'
        Save-AppLogo -Source $AppLogoSource -IntuneDirectory $IntuneDir -ToolkitDirectory $toolkitDir
    }

    # PID MUST run after Deploy-Application.ps1 has been fully populated. The PID generator reads the generated package variables and PSADT commands.
    # PID is generated after the package staging/source and before Deploy-Application.ps1 is populated.
    if($GeneratePID){
        # PID generation is a mandatory build phase whenever the UI requests it.
        # Do not silently skip it and do not assume parameter names that may differ
        # between PID-script revisions. Discover the supported parameters first.
        Write-AutoLog '============================================================' 'INFO'
        Write-AutoLog 'PID DOCUMENT GENERATION - STARTING' 'INFO'
        Write-AutoLog "PID PackageRoot : $PackageRoot" 'INFO'
        Write-AutoLog "PID Template    : $pidTemplate" 'INFO'
        Write-AutoLog "PID Output      : $DocDir" 'INFO'
        Write-AutoLog "PID Log Folder  : $RunLogDir" 'INFO'

        $pidScript = Join-Path $PSScriptRoot 'Generate-PIDDocument.ps1'
        if(-not (Test-Path -LiteralPath $pidScript -PathType Leaf)){
            throw "PID script not found beside Invoke-PackagingEngine.ps1: $pidScript"
        }

        if(-not (Test-Path -LiteralPath $pidTemplate -PathType Leaf)){
            throw "PID Word template was not found: $pidTemplate"
        }

        if(-not (Test-Path -LiteralPath $PackageRoot -PathType Container)){
            throw "PID PackageRoot does not exist: $PackageRoot"
        }

        $pidCommand = Get-Command -Name $pidScript -ErrorAction Stop
        $pidSupported = $pidCommand.Parameters.Keys
        $pidArgs = @{}

        if($pidSupported -contains 'PackageRoot'){ $pidArgs['PackageRoot'] = $PackageRoot }
        elseif($pidSupported -contains 'PackageRootPath'){ $pidArgs['PackageRootPath'] = $PackageRoot }
        elseif($pidSupported -contains 'PackageRootPathName'){ $pidArgs['PackageRootPathName'] = $PackageRoot }
        else { throw "PID script does not expose a supported PackageRoot parameter." }

        if($pidSupported -contains 'TemplatePath'){ $pidArgs['TemplatePath'] = $pidTemplate }
        elseif($pidSupported -contains 'PIDTemplate'){ $pidArgs['PIDTemplate'] = $pidTemplate }
        elseif($pidSupported -contains 'Template'){ $pidArgs['Template'] = $pidTemplate }
        else { throw "PID script does not expose a supported template parameter." }

        if($pidSupported -contains 'OutputDirectory'){ $pidArgs['OutputDirectory'] = $DocDir }
        elseif($pidSupported -contains 'OutputPath'){ $pidArgs['OutputPath'] = $DocDir }
        elseif($pidSupported -contains 'DocsDirectory'){ $pidArgs['DocsDirectory'] = $DocDir }
        else { throw "PID script does not expose a supported output-directory parameter." }

        if($pidSupported -contains 'LogDirectory'){ $pidArgs['LogDirectory'] = $RunLogDir }

        Write-AutoLog ("PID supported parameters: " + (($pidSupported | Sort-Object) -join ', ')) 'INFO'
        Write-AutoLog 'PID script invocation starting now.' 'INFO'

        try {
            # Generate-PIDDocument.ps1 is a PowerShell script. A successful
            # script invocation may leave LASTEXITCODE empty, so use invocation status
            # when no native exit code was returned.
            $pidOutput = @(& $pidScript @pidArgs 2>&1)
            $pidInvocationSucceeded = $?
            $pidExitCode = $LASTEXITCODE
            if($null -eq $pidExitCode -or "$pidExitCode" -eq ''){
                $pidExitCode = if($pidInvocationSucceeded){ 0 } else { 1 }
            }
        }
        catch {
            Write-AutoLog ("PID script threw an exception: " + $_.Exception.Message) 'ERROR'
            throw
        }

        foreach($pidLine in $pidOutput){
            if($null -ne $pidLine -and "$pidLine".Trim()){
                Write-AutoLog ("PID: " + "$pidLine") 'INFO'
            }
        }

        if($pidExitCode -ne 0){
            throw "PID document script exited with code $pidExitCode. Review the PID messages above and $RunLogDir."
        }

        $pidOutputPath = Join-Path $DocDir ($PackageName + '_PID.docx')
        if(-not (Test-Path -LiteralPath $pidOutputPath -PathType Leaf)){
            $pidCandidates = @(Get-ChildItem -LiteralPath $DocDir -Filter '*_PID.docx' -File -ErrorAction SilentlyContinue)
            if($pidCandidates.Count -eq 1){
                $pidOutputPath = $pidCandidates[0].FullName
                Write-AutoLog "PID document found using generated-document discovery: $pidOutputPath" 'INFO'
            }
        }

        if(-not (Test-Path -LiteralPath $pidOutputPath -PathType Leaf)){
            throw "PID script returned successfully but no PID document was found in '$DocDir'. Expected: $PackageName`_PID.docx"
        }

        $pidSize = (Get-Item -LiteralPath $pidOutputPath).Length
        if($pidSize -le 0){
            throw "PID document was created but is empty: $pidOutputPath"
        }

        Write-AutoLog "PID document generated successfully: $pidOutputPath" 'SUCCESS'
        Write-AutoLog "PID document size: $pidSize bytes" 'SUCCESS'
        Write-AutoLog 'PID DOCUMENT GENERATION - COMPLETED' 'SUCCESS'
        Write-AutoLog '============================================================' 'INFO'
    }
    else {
        Write-AutoLog 'PID DOCUMENT GENERATION - SKIPPED (GeneratePID switch was not supplied).' 'WARNING'
    }


        if($PrepareOnly){
            Write-AutoLog 'PREPARE PACKAGE - COMPLETED. PID generated and Deploy-Application.ps1 populated.' 'SUCCESS'
            Write-AutoLog 'PACKAGE READY FOR REVIEW - Please check and confirm Deploy-Application.ps1, then click CONFIRM & BUILD INTUNEWIN.' 'SUCCESS'
            return
        }
    }

    # IntuneWinAppUtil is required both for a fresh package build and for the
    # CONFIRM & BUILD step after the user reviews/edits Deploy-Application.ps1.
    # Previously it was resolved only inside the !ContinuePrepared branch, so
    # CONFIRM & BUILD passed a null -FilePath to Start-Process.
    if($GenerateIntuneWin){
        if(-not (Test-Path -LiteralPath $TemplateRoot -PathType Container)){
            throw "Template folder does not exist: $TemplateRoot"
        }
        $intuneUtil = Get-ChildItem -LiteralPath $TemplateRoot -Filter 'IntuneWinAppUtil.exe' -File -Recurse -ErrorAction SilentlyContinue |
            Select-Object -First 1 -ExpandProperty FullName
        if([string]::IsNullOrWhiteSpace($intuneUtil)){
            throw "IntuneWinAppUtil.exe was not found anywhere under $TemplateRoot."
        }
        Write-AutoLog "IntuneWinAppUtil: $intuneUtil" 'SUCCESS'
    }

    if($ContinuePrepared){
        if(-not (Test-Path -LiteralPath $PackageDir -PathType Container)){
            throw "Prepared package folder was not found: $PackageDir"
        }
        $deploy = Join-Path $PackageDir 'Deploy-Application.ps1'
        if(-not (Test-Path -LiteralPath $deploy -PathType Leaf)){
            throw "Prepared Deploy-Application.ps1 was not found: $deploy"
        }
        if(-not (Test-Path -LiteralPath (Join-Path $PackageDir 'Deploy-Application.exe') -PathType Leaf)){
            throw "Prepared package does not contain Deploy-Application.exe. Re-run PREPARE PACKAGE."
        }
        Write-AutoLog 'CHECK & EDIT: using the existing staged Deploy-Application.ps1 exactly as saved by the user.' 'SUCCESS'
        Write-AutoLog 'CHECK & EDIT: PID generation was completed during PREPARE PACKAGE; no PID regeneration will overwrite the edited script.' 'INFO'
    }

    if($GenerateIntuneWin){
        Write-AutoLog 'PACKAGE VALIDATION - STARTING' 'INFO'
        if(-not (Test-Path -LiteralPath (Join-Path $PackageDir 'Deploy-Application.ps1') -PathType Leaf)){ throw 'Deploy-Application.ps1 is missing from the final package.' }
        if(-not (Test-Path -LiteralPath (Join-Path $PackageDir 'Deploy-Application.exe') -PathType Leaf)){ throw 'Deploy-Application.exe is missing from the final package.' }
        Write-AutoLog 'PACKAGE VALIDATION - COMPLETED' 'SUCCESS'
        Write-AutoLog 'INTUNEWIN GENERATION - STARTING' 'INFO'
        $intuneLog = Join-Path $RunLogDir "$PackageName-IntuneWin.log"
        $intuneErrLog = $intuneLog + '.err'
        Write-AutoLog 'Starting IntuneWinAppUtil.exe.'
        $buildStarted = Get-Date
        Remove-Item -LiteralPath $intuneLog,$intuneErrLog -Force -ErrorAction SilentlyContinue
        $p = Start-Process -FilePath $intuneUtil -WorkingDirectory (Split-Path -Parent $intuneUtil) -ArgumentList @('-c',$PackageDir,'-s','Deploy-Application.exe','-o',$IntuneDir,'-q') -Wait -PassThru -NoNewWindow -RedirectStandardOutput $intuneLog -RedirectStandardError $intuneErrLog
        if($p.ExitCode -ne 0){
            $details = ''
            if(Test-Path -LiteralPath $intuneErrLog){ $details = (Get-Content -LiteralPath $intuneErrLog -Raw -ErrorAction SilentlyContinue).Trim() }
            if([string]::IsNullOrWhiteSpace($details) -and (Test-Path -LiteralPath $intuneLog)){ $details = (Get-Content -LiteralPath $intuneLog -Raw -ErrorAction SilentlyContinue).Trim() }
            throw "IntuneWinAppUtil.exe failed with exit code $($p.ExitCode). $details"
        }
        $iw = @(Get-ChildItem -LiteralPath $IntuneDir -Filter '*.intunewin' -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $buildStarted.AddSeconds(-2) } | Sort-Object LastWriteTime -Descending | Select-Object -First 1)[0]
        if(-not $iw){
            $iw = @(Get-ChildItem -LiteralPath $PackageDir -Filter '*.intunewin' -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $buildStarted.AddSeconds(-2) } | Sort-Object LastWriteTime -Descending | Select-Object -First 1)[0]
        }
        if(-not $iw){
            $details = ''
            if(Test-Path -LiteralPath $intuneLog){ $details += (Get-Content -LiteralPath $intuneLog -Raw -ErrorAction SilentlyContinue) }
            if(Test-Path -LiteralPath $intuneErrLog){ $details += [Environment]::NewLine + (Get-Content -LiteralPath $intuneErrLog -Raw -ErrorAction SilentlyContinue) }
            throw "IntuneWinAppUtil completed but no .intunewin file was found. Utility output: $details"
        }
        $target = Join-Path $IntuneDir "$PackageName.intunewin"
        if($iw.FullName -ne $target){ Move-Item -LiteralPath $iw.FullName -Destination $target -Force }
        Write-AutoLog "IntuneWin created: $target" 'SUCCESS'
        Write-AutoLog 'INTUNEWIN GENERATION - COMPLETED' 'SUCCESS'
    }

    $summary = @("APPLICATION PACKAGING SUMMARY","Package : $PackageName","Installer: $InstallerType","Status  : SUCCESS","Output  : $PackageRoot","Log     : $LogPath","Log Root: $AutomationLogRoot")
    Set-Content -LiteralPath (Join-Path $DocDir 'PackagingSummary.txt') -Value $summary -Encoding UTF8

    # Keep all automation logs exclusively under C:\Temp\AutomationLogs.
    # Remove any package-local AutomationLogs folders/log files that may have been
    # copied from an older template or source package.
    foreach($localLogDir in @(Get-ChildItem -LiteralPath $PackageRoot -Directory -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq 'AutomationLogs' })){
        try { Remove-Item -LiteralPath $localLogDir.FullName -Recurse -Force -ErrorAction Stop } catch { Write-AutoLog ("Could not remove local AutomationLogs folder: {0}" -f $_.Exception.Message) 'WARNING' }
    }
    foreach($localLogFile in @(Get-ChildItem -LiteralPath $DocDir -File -Filter '*.log' -ErrorAction SilentlyContinue)){
        try { Remove-Item -LiteralPath $localLogFile.FullName -Force -ErrorAction Stop } catch { Write-AutoLog ("Could not remove local DOC log: {0}" -f $_.Exception.Message) 'WARNING' }
    }
    Write-AutoLog 'PACKAGING COMPLETED SUCCESSFULLY' 'SUCCESS'
} catch {
    Write-AutoLog "PACKAGING FAILED: $($_.Exception.Message)" 'ERROR'
    throw
}
