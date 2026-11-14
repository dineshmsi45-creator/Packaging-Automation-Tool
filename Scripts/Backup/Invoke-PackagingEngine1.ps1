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
    [Alias('UninstallWorkingDir','UninstallWD')][string]$UninstallWorkingDirectory = '',
    [string]$PreInstallSteps = '',
    [string]$PostInstallSteps = '',
    [string]$PreUninstallSteps = '',
    [string]$PostUninstallSteps = '',
    [string]$InstallCloseApps = '',
    [string]$UninstallCloseApps = '',
    [switch]$GenerateMST,
    [switch]$GeneratePID,
    [switch]$IncludeServiceUI,
    [switch]$GenerateIntuneWin,
    [switch]$PrepareOnly,
    [switch]$ContinuePrepared,
    [string]$TemplateRoot = 'C:\Temp\template',
    [string]$OutputRoot = 'C:\Packaging\Output',
    [string]$AutomationLogRoot = 'C:\Temp\AutomationLogs'
)

$ErrorActionPreference = 'Stop'
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
    "SourcePath=$SourcePath"
    "InstallFile=$InstallFile"
    "UninstallMethod=$UninstallMethod"
    "UninstallFile=$UninstallFile"
    "UninstallParameters=$UninstallParameters"
    "UninstallWorkingDirectory=$UninstallWorkingDirectory"
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
        if($record){ return [string]$record.StringData(1) }
        return ''
    } finally {
        try { if($view){ $view.Close() } } catch {}
        try { if($database){ [void][Runtime.InteropServices.Marshal]::ReleaseComObject($database) } } catch {}
        try { if($installer){ [void][Runtime.InteropServices.Marshal]::ReleaseComObject($installer) } } catch {}
    }
}

function Find-PSADTTemplate {
    param([string]$Root)
    $candidates = Get-ChildItem -LiteralPath $Root -Directory -ErrorAction SilentlyContinue
    foreach($d in $candidates){
        if((Test-Path (Join-Path $d.FullName 'AppDeployToolkit')) -and (Test-Path (Join-Path $d.FullName 'Deploy-Application.ps1'))){ return $d.FullName }
    }
    if((Test-Path (Join-Path $Root 'AppDeployToolkit')) -and (Test-Path (Join-Path $Root 'Deploy-Application.ps1'))){ return $Root }
    throw "PSADT 3.10 template folder was not found under $Root. Expected Deploy-Application.ps1 and AppDeployToolkit."
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

    $psadtSource = Find-PSADTTemplate -Root $TemplateRoot
    Write-AutoLog "PSADT 3.10 template: $psadtSource" 'SUCCESS'

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
        $msiProductCode = Get-MsiProperty -MsiPath $installerFull -PropertyName 'ProductCode'
        $msiProductName = Get-MsiProperty -MsiPath $installerFull -PropertyName 'ProductName'
        $msiProductVersion = Get-MsiProperty -MsiPath $installerFull -PropertyName 'ProductVersion'
        $msiManufacturer = Get-MsiProperty -MsiPath $installerFull -PropertyName 'Manufacturer'
        $msiUpgradeCode = Get-MsiProperty -MsiPath $installerFull -PropertyName 'UpgradeCode'
        if([string]::IsNullOrWhiteSpace($msiProductCode)) { throw "Unable to read ProductCode directly from MSI: $installerFull" }
        Write-AutoLog "MSI ProductCode auto-detected: $msiProductCode" 'SUCCESS'
        Write-AutoLog "MSI ProductName: $msiProductName" 'INFO'
        Write-AutoLog "MSI ProductVersion: $msiProductVersion" 'INFO'
        Write-AutoLog "MSI Manufacturer: $msiManufacturer" 'INFO'
        Write-AutoLog "MSI UpgradeCode: $msiUpgradeCode" 'INFO'
    }

    if($InstallerType -eq 'MSI'){
        if([IO.Path]::GetExtension($installerFull) -ine '.msi'){ throw 'Installer Type is MSI but the selected installer is not an MSI.' }
        if($GenerateMST){
            $mstScript = Join-Path $PSScriptRoot 'Create-StandardMST-Schema300.ps1'
            Write-AutoLog 'Calling the supplied Standard MST script (unchanged).'
            & $mstScript -MsiPath $installerFull -Vendor $Vendor -AppName $ApplicationName -Version $Version -Architecture $Architecture -LogDirectory $RunLogDir
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
    # Preserve the exact values entered in the UI for the PSADT application variables.
    # Package/folder naming uses its own package-name convention; it must not alter
    # the values written into Deploy-Application.ps1.
    Set-PSADTVariableValue -FilePath $deploy -Name 'appVendor' -Value $Vendor.Trim()
    Set-PSADTVariableValue -FilePath $deploy -Name 'appName' -Value $ApplicationName.Trim()
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

    # Build PSADT commands. MSI commands always use the standard silent defaults unless
    # the user explicitly supplies different parameters. Optional UI fields are not
    # required and are never passed to the engine as empty command-line arguments.
    if($InstallerType -eq 'MSI'){
        $effectiveInstallParameters = if([string]::IsNullOrWhiteSpace($InstallParameters)){ '/qn /norestart' } else { $InstallParameters.Trim() }
        if($GenerateMST){
            # Build the command using a single-quoted format string so paths/quotes are
            # never represented with C-style backslash escaping in PowerShell source.
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

    if($InstallerType -eq 'MSI' -and [string]::IsNullOrWhiteSpace($UninstallParameters)){
        $UninstallParameters = '/qn /norestart'
    }

    $uninstallPsadt = switch($UninstallMethod){
        'Main EXE / Installer' {
            if($InstallerType -eq 'MSI'){
                'Execute-MSI -Action Uninstall -Path "{0}" -Parameters "{1}"' -f $msiProductCode,$UninstallParameters
            } elseif([string]::IsNullOrWhiteSpace($UninstallParameters)){
                'Execute-Process -Path "$dirFiles\{0}" -WindowStyle Hidden -Wait' -f $installer
            } else {
                'Execute-Process -Path "$dirFiles\{0}" -Parameters "{1}" -WindowStyle Hidden -Wait' -f $installer,$UninstallParameters
            }
        }
        'Uninstall from different location' {
            if([string]::IsNullOrWhiteSpace($UninstallFile)){throw 'Uninstall executable/path is required.'}
            $uninstallPath = $UninstallFile
            if(-not [IO.Path]::IsPathRooted($uninstallPath)){ $uninstallPath = '$dirFiles\' + $uninstallPath }
            $wd = if($UninstallWorkingDirectory){' -WorkingDirectory "{0}"' -f $UninstallWorkingDirectory}else{''}
            if([string]::IsNullOrWhiteSpace($UninstallParameters)){
                'Execute-Process -Path "{0}" -WindowStyle Hidden -Wait{1}' -f $uninstallPath,$wd
            } else {
                'Execute-Process -Path "{0}" -Parameters "{1}" -WindowStyle Hidden -Wait{2}' -f $uninstallPath,$UninstallParameters,$wd
            }
        }
        'MSI ProductCode' {
            if($InstallerType -ne 'MSI'){throw 'MSI ProductCode uninstall is available only for MSI packages.'}
            'Execute-MSI -Action Uninstall -Path "{0}" -Parameters "{1}"' -f $msiProductCode,$UninstallParameters
        }
    }


    Write-AutoLog 'BUILD PHASE: PSADT CONFIGURATION COMPLETED; preparing PID document generation.' 'INFO'

    # Close Apps: update the Show-InstallationWelcome command that already exists
    # in the PSADT 3.10 Deploy-Application.ps1 template. In PSADT 3.10 the
    # Show-InstallationWelcome line is located AFTER the PRE-INSTALLATION /
    # PRE-UNINSTALLATION header but BEFORE the <Perform ... tasks here> marker.
    # Therefore the search must start at the lifecycle header, not at the task marker.
    # Do not create a second Show-InstallationWelcome command.
    # If no EXE is supplied, comment the existing command in both lifecycle sections.
    Update-ExistingShowInstallationWelcomeCloseApps `
    -FilePath $deploy `
    -CloseApps $InstallCloseApps `
    -SectionName 'Pre-Installation'

    Insert-PSADTSectionContent -FilePath $deploy -Marker '## <Perform Pre-Installation tasks here>' -Block $PreInstallSteps -Tag 'Pre-Installation'
    Insert-PSADTSectionContent -FilePath $deploy -Marker '## <Perform Installation tasks here>' -Block $installPsadt -Tag 'Installation'
    Insert-PSADTSectionContent -FilePath $deploy -Marker '## <Perform Post-Installation tasks here>' -Block $PostInstallSteps -Tag 'Post-Installation'

    Update-ExistingShowInstallationWelcomeCloseApps `
    -FilePath $deploy `
    -CloseApps $UninstallCloseApps `
    -SectionName 'Pre-Uninstallation'

    Insert-PSADTSectionContent -FilePath $deploy -Marker '## <Perform Pre-Uninstallation tasks here>' -Block $PreUninstallSteps -Tag 'Pre-Uninstallation'
    Insert-PSADTSectionContent -FilePath $deploy -Marker '## <Perform Uninstallation tasks here>' -Block $uninstallPsadt -Tag 'Uninstallation'
    Insert-PSADTSectionContent -FilePath $deploy -Marker '## <Perform Post-Uninstallation tasks here>' -Block $PostUninstallSteps -Tag 'Post-Uninstallation'
    Write-AutoLog 'PSADT 3.10 deployment sections configured.' 'SUCCESS'

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

        $pidScript = Join-Path $PSScriptRoot 'Generate-PIDDocument_Final_V2_5.ps1'
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
            # Generate-PIDDocument_Final_V2_5.ps1 is a PowerShell script. A successful
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
        Write-AutoLog 'Starting IntuneWinAppUtil.exe.'
        $p = Start-Process -FilePath $intuneUtil -ArgumentList @('-c',$PackageDir,'-s','Deploy-Application.exe','-o',$IntuneDir,'-q') -Wait -PassThru -NoNewWindow -RedirectStandardOutput $intuneLog -RedirectStandardError ($intuneLog+'.err')
        if($p.ExitCode -ne 0){ throw "IntuneWinAppUtil.exe failed with exit code $($p.ExitCode). See $intuneLog" }
        $iw = Get-ChildItem -LiteralPath $IntuneDir -Filter '*.intunewin' -File | Select-Object -First 1
        if(-not $iw){ throw 'IntuneWinAppUtil completed but no .intunewin file was found.' }
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
