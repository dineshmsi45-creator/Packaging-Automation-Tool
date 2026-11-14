# Generate-PIDDocument V2.4 - fixed PSCustomObject registry property error
<#
.SYNOPSIS
    Generates the PID Word document from an Application Packaging package.

.DESCRIPTION
    Expected package structure:

    <PackageRoot>\
        Package\
            AppDeployToolkit\
            Files\              # vendor EXE/MSI/source files
            SupportFiles\
            Deploy-Application.ps1
            Deploy-Application.exe
            ServiceUI.exe        # optional; may also be under Package subfolders
        Docs\

    This version deliberately does NOT use .Count, .Item(), Find, Replace,
    ClearFormatting, or other fragile Word COM collection access patterns.
    Word tables/paragraphs are first enumerated into normal PowerShell arrays.

    Output:
        <PackageRoot>\Docs\<PackageRootName>_PID.docx
#>

[CmdletBinding()]
param(
    [string]$PackageRoot,
    [string]$TemplatePath,
    [string]$OutputDirectory,
    [string]$LogDirectory = 'C:\Temp\AutomationLogs',
    [switch]$OpenDocument
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$StartTime = Get-Date
$LogLines = [System.Collections.Generic.List[string]]::new()

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')]
        [string]$Level = 'INFO'
    )
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    [void]$LogLines.Add($line)
    Write-Host $line
}

function Fail {
    param([string]$Message)
    Write-Log $Message 'ERROR'
    throw $Message
}

# -----------------------------------------------------------------------------
# Helpers: normal PowerShell collection counting only (never Word COM Count)
# -----------------------------------------------------------------------------
function Get-Array {
    param($Value)
    if ($null -eq $Value) { return @() }
    return @($Value)
}

# -----------------------------------------------------------------------------
# Package-root discovery
# -----------------------------------------------------------------------------
function Test-PackageRoot {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return $false }

    $package = Join-Path $Path 'Package'
    $files   = Join-Path $package 'Files'

    return (
        (Test-Path -LiteralPath $package -PathType Container) -and
        (Test-Path -LiteralPath $files -PathType Container)
    )
}

function Resolve-PackageRoot {
    param([string]$RequestedPath)
    if ([string]::IsNullOrWhiteSpace($RequestedPath)) {
        Fail 'PackageRoot is required. The packaging engine must pass the exact generated package root.'
    }
    try { $resolved = (Resolve-Path -LiteralPath $RequestedPath -ErrorAction Stop).Path }
    catch { Fail "PackageRoot does not exist or cannot be resolved: $RequestedPath" }
    if (-not (Test-PackageRoot $resolved)) {
        Fail "Invalid PackageRoot: $resolved. Expected Package\Files under the supplied package root."
    }
    return $resolved
}

$PackageRoot = Resolve-PackageRoot $PackageRoot
$PackageRoot = (Resolve-Path -LiteralPath $PackageRoot).Path
$PackageName = Split-Path -Leaf $PackageRoot.TrimEnd('\')
$PackageDir  = Join-Path $PackageRoot 'Package'
$FilesDir    = Join-Path $PackageDir 'Files'
$DocsDir     = if ($OutputDirectory) { $OutputDirectory } else { Join-Path $PackageRoot 'Docs' }
$LogDir      = if ($LogDirectory) { $LogDirectory } else { 'C:\Temp\AutomationLogs' }

if (-not (Test-Path -LiteralPath $DocsDir -PathType Container)) {
    New-Item -Path $DocsDir -ItemType Directory -Force | Out-Null
}
if (-not (Test-Path -LiteralPath $LogDir -PathType Container)) {
    New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
}

Write-Log "Package root: $PackageRoot" 'OK'
Write-Log "Package folder: $PackageDir" 'OK'
Write-Log "Files folder: $FilesDir" 'OK'
Write-Log "Package/document base name: $PackageName" 'OK'
Write-Log "Output folder: $DocsDir" 'OK'

# -----------------------------------------------------------------------------
# Template discovery
# -----------------------------------------------------------------------------
function Resolve-Template {
    param([string]$Requested)

    $candidates = [System.Collections.Generic.List[string]]::new()
    if ($Requested) { [void]$candidates.Add($Requested) }

    [void]$candidates.Add((Join-Path $PackageRoot 'Docs\Package_Template.docx'))
    [void]$candidates.Add((Join-Path $PackageRoot 'Package_Template.docx'))
    [void]$candidates.Add((Join-Path (Get-Location).Path 'Package_Template.docx'))
    [void]$candidates.Add((Join-Path $PSScriptRoot 'Package_Template.docx'))

    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath $c -PathType Leaf)) {
            return (Resolve-Path -LiteralPath $c).Path
        }
    }

    Fail "Word template was not found. Supply -TemplatePath 'C:\Path\Package_Template.docx'."
}

$TemplatePath = Resolve-Template $TemplatePath
Write-Log "Template: $TemplatePath" 'OK'

# -----------------------------------------------------------------------------
# Deployment script
# -----------------------------------------------------------------------------
$DeployScriptPath = Join-Path $PackageDir 'Deploy-Application.ps1'
if (-not (Test-Path -LiteralPath $DeployScriptPath -PathType Leaf)) {
    Fail "Deploy-Application.ps1 was not found at: $DeployScriptPath"
}
$DeployScript = Get-Content -LiteralPath $DeployScriptPath -Raw -ErrorAction Stop
Write-Log "Deploy-Application.ps1 loaded." 'OK'

function Get-ScriptVariable {
    param(
        [string]$Text,
        [string[]]$Names
    )

    foreach ($name in $Names) {
        # Use hex character classes for single/double quotes so that
        # PowerShell string quoting cannot break the regular expression.
        $n = [regex]::Escape($name)

        $patterns = @(
            ('(?im)^\s*\$' + $n + '\s*=\s*[\x27\x22]([^\x27\x22]*)[\x27\x22]'),
            ('(?im)\$' + $n + '\s*=\s*[\x27\x22]([^\x27\x22]*)[\x27\x22]')
        )

        foreach ($pattern in $patterns) {
            try {
                $m = [regex]::Match($Text, $pattern)
                if ($m.Success) {
                    $value = $m.Groups[1].Value.Trim()
                    if ($value) {
                        return $value
                    }
                }
            }
            catch {
                Write-Log "Could not parse variable '$name': $($_.Exception.Message)" 'WARN'
            }
        }
    }

    return $null
}

$appVendor  = Get-ScriptVariable $DeployScript @('appVendor','AppVendor','vendor','Vendor')
$appName    = Get-ScriptVariable $DeployScript @('appName','AppName','applicationName','ApplicationName')
$appVersion = Get-ScriptVariable $DeployScript @('appVersion','AppVersion','applicationVersion','ApplicationVersion')
$appArch    = Get-ScriptVariable $DeployScript @('appArch','AppArch','architecture','Architecture')

# Folder-name fallback. Handles Vendor_AppName_4.17.0_P1_x64 and similar.
$parts = @($PackageName -split '_')
if (-not $appVendor -and $parts.Count -ge 1) { $appVendor = $parts[0] }
if (-not $appName -and $parts.Count -ge 2) { $appName = $parts[1] }
if (-not $appVersion) {
    foreach ($part in $parts) {
        if ($part -match '^\d+(?:\.\d+){1,4}$') {
            $appVersion = $part
            break
        }
    }
}
if (-not $appArch) {
    if ($parts -contains 'x64' -or $parts -contains '64') { $appArch = 'x64' }
    elseif ($parts -contains 'x86' -or $parts -contains '32') { $appArch = 'x86' }
}

if (-not $appVendor)  { $appVendor = 'Unknown' }
if (-not $appName)    { $appName = 'Unknown' }
if (-not $appVersion) { $appVersion = 'Unknown' }
if (-not $appArch)    { $appArch = '' }

Write-Log "Vendor: $appVendor" 'INFO'
Write-Log "Application: $appName" 'INFO'
Write-Log "Version: $appVersion" 'INFO'
Write-Log "Architecture: $(if($appArch){$appArch}else{'Not specified'})" 'INFO'

# -----------------------------------------------------------------------------
# Source analysis
# -----------------------------------------------------------------------------
$sourceFiles = @(Get-ChildItem -LiteralPath $FilesDir -File -Recurse -ErrorAction SilentlyContinue)
Write-Log "Source files found in Package\Files: $($sourceFiles.Count)" 'OK'

$msis = @($sourceFiles | Where-Object { $_.Extension -ieq '.msi' })
$exes = @($sourceFiles | Where-Object { $_.Extension -ieq '.exe' })

# ServiceUI is NOT expected inside Files. Search Package but explicitly exclude Files.
$serviceUI = $null
$serviceCandidates = @(Get-ChildItem -LiteralPath $PackageDir -File -Recurse -Filter 'ServiceUI.exe' -ErrorAction SilentlyContinue)
foreach ($candidate in $serviceCandidates) {
    if ($candidate.FullName -notlike ($FilesDir.TrimEnd('\') + '\*')) {
        $serviceUI = $candidate
        break
    }
}

if ($serviceUI) {
    Write-Log "ServiceUI.exe detected: $($serviceUI.FullName)" 'OK'
} else {
    Write-Log "ServiceUI.exe not detected outside Package\Files." 'INFO'
}

if ($msis.Count -gt 0 -and $exes.Count -gt 0) { $packageType = 'MSI + EXE + Intune' }
elseif ($msis.Count -gt 0) { $packageType = 'MSI + Intune' }
elseif ($exes.Count -gt 0) { $packageType = 'EXE + Intune' }
else { $packageType = 'Intune' }

# -----------------------------------------------------------------------------
# Installed application / registry discovery - explicitly check BOTH registry views
# -----------------------------------------------------------------------------
function Get-UninstallEntriesByArchitecture {
    $entries = [System.Collections.Generic.List[object]]::new()
    $subPath = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'

    foreach ($viewName in @('64-bit','32-bit')) {
        $view = if ($viewName -eq '64-bit') {
            [Microsoft.Win32.RegistryView]::Registry64
        } else {
            [Microsoft.Win32.RegistryView]::Registry32
        }

        $baseKey = $null
        $uninstallKey = $null
        try {
            $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                [Microsoft.Win32.RegistryHive]::LocalMachine,
                $view
            )
            $uninstallKey = $baseKey.OpenSubKey($subPath)
            if ($null -eq $uninstallKey) { continue }

            foreach ($subName in $uninstallKey.GetSubKeyNames()) {
                $subKey = $null
                try {
                    $subKey = $uninstallKey.OpenSubKey($subName)
                    if ($null -eq $subKey) { continue }

                    $displayName = [string]$subKey.GetValue('DisplayName','')
                    if ([string]::IsNullOrWhiteSpace($displayName)) { continue }

                    $hivePath = 'HKEY_LOCAL_MACHINE\' + $subPath + '\' + $subName
                    [void]$entries.Add([pscustomobject]@{
                        Architecture = $viewName
                        KeyName       = $subName
                        RegistryPath  = $hivePath
                        DisplayName   = $displayName
                        DisplayVersion= [string]$subKey.GetValue('DisplayVersion','')
                        Publisher     = [string]$subKey.GetValue('Publisher','')
                        InstallLocation = [string]$subKey.GetValue('InstallLocation','')
                        UninstallString = [string]$subKey.GetValue('UninstallString','')
                        DisplayIcon   = [string]$subKey.GetValue('DisplayIcon','')
                    })
                }
                catch {
                    Write-Log "Could not read uninstall key '$subName' from $viewName registry: $($_.Exception.Message)" 'WARN'
                }
                finally {
                    if ($subKey) { try { $subKey.Dispose() } catch {} }
                }
            }
        }
        catch {
            Write-Log "Could not enumerate $viewName uninstall registry: $($_.Exception.Message)" 'WARN'
        }
        finally {
            if ($uninstallKey) { try { $uninstallKey.Dispose() } catch {} }
            if ($baseKey) { try { $baseKey.Dispose() } catch {} }
        }
    }

    return @($entries)
}

$uninstallEntries = @(Get-UninstallEntriesByArchitecture)
Write-Log "Uninstall registry entries checked: 64-bit=$(@($uninstallEntries | Where-Object Architecture -eq '64-bit').Count), 32-bit=$(@($uninstallEntries | Where-Object Architecture -eq '32-bit').Count)" 'OK'

function Test-ApplicationMatch {
    param($Entry)

    $display = [string]$Entry.DisplayName
    $publisher = [string]$Entry.Publisher
    $version = [string]$Entry.DisplayVersion

    $nameMatch = ($appName -ne 'Unknown' -and $display -like "*$appName*")
    $vendorMatch = ($appVendor -ne 'Unknown' -and (
        $display -like "*$appVendor*" -or $publisher -like "*$appVendor*"
    ))
    $versionMatch = ($appVersion -ne 'Unknown' -and $version -eq $appVersion)

    return [pscustomobject]@{
        NameMatch = $nameMatch
        VendorMatch = $vendorMatch
        VersionMatch = $versionMatch
        IsMatch = (($nameMatch -or $vendorMatch) -and $versionMatch)
    }
}

$matchingEntries = @(
    $uninstallEntries |
        ForEach-Object {
            $m = Test-ApplicationMatch $_
            if ($m.IsMatch) {
                $score = 0
                if ($m.NameMatch) { $score += 100 }
                if ($m.VendorMatch) { $score += 50 }
                if ($m.VersionMatch) { $score += 100 }
                [pscustomobject]@{
                    Score = $score
                    Entry = $_
                }
            }
        } |
        Sort-Object Score -Descending |
        ForEach-Object { $_.Entry }
)

# If exact package-version matches are unavailable, retain a softer match only
# for installation-folder discovery; Intune detection still reports exact matches.
$softMatches = @(
    $uninstallEntries |
        ForEach-Object {
            $display = [string]$_.DisplayName
            $publisher = [string]$_.Publisher
            $nameMatch = ($appName -ne 'Unknown' -and $display -like "*$appName*")
            $vendorMatch = ($appVendor -ne 'Unknown' -and ($display -like "*$appVendor*" -or $publisher -like "*$appVendor*"))
            if ($nameMatch -or $vendorMatch) { $_ }
        }
)

$installed = $null
if ($matchingEntries.Count -gt 0) {
    $installed = $matchingEntries[0]
    Write-Log "Installed application match: [$($installed.Architecture)] $($installed.DisplayName) | Version: $($installed.DisplayVersion)" 'OK'
} elseif ($softMatches.Count -gt 0) {
    $installed = $softMatches[0]
    Write-Log "Exact package-version registry match was not found; using best name/vendor match for installation-folder discovery: [$($installed.Architecture)] $($installed.DisplayName) | Version: $($installed.DisplayVersion)" 'WARN'
} else {
    Write-Log 'Application was not found in either 64-bit or 32-bit machine uninstall registry.' 'WARN'
}

# -----------------------------------------------------------------------------
# Installation directory discovery - MAIN folder only
# Priority:
#   1. Registry InstallLocation from the matched uninstall entry.
#   2. Exact application-name folder directly under Program Files / x86.
#   3. Vendor\Application folder.
#   4. Best matching application folder one level below a vendor folder.
#   5. Vendor folder itself only as a last resort.
# This deliberately avoids selecting arbitrary deep child folders.
# -----------------------------------------------------------------------------
$installFolders = [System.Collections.Generic.List[string]]::new()

function Normalize-InstallPath {
    param([string]$PathText)
    if ([string]::IsNullOrWhiteSpace($PathText)) { return '' }
    try {
        return ([System.IO.Path]::GetFullPath($PathText.Trim().Trim('"'))).TrimEnd('\')
    } catch {
        return $PathText.Trim().Trim('"').TrimEnd('\')
    }
}

function Add-InstallFolderIfValid {
    param([string]$Candidate, [string]$Reason)
    if ([string]::IsNullOrWhiteSpace($Candidate)) { return $false }
    try {
        $normalized = Normalize-InstallPath $Candidate
        if (Test-Path -LiteralPath $normalized -PathType Container) {
            [void]$installFolders.Add((Resolve-Path -LiteralPath $normalized).Path)
            Write-Log "Main installation folder detected: $normalized [$Reason]" 'OK'
            return $true
        }
    } catch {
        Write-Log "Unable to validate installation folder '$Candidate': $($_.Exception.Message)" 'WARN'
    }
    return $false
}

# 1. Registry InstallLocation is authoritative when it exists and is valid.
if ($installed -and $installed.InstallLocation) {
    if (-not (Add-InstallFolderIfValid $installed.InstallLocation 'Registry InstallLocation')) {
        Write-Log "Registry InstallLocation is not a valid existing folder: $($installed.InstallLocation)" 'WARN'
    }
}

$programRoots = @(
    ${env:ProgramFiles},
    ${env:ProgramFiles(x86)}
) | Where-Object {
    $_ -and (Test-Path -LiteralPath $_ -PathType Container)
} | Select-Object -Unique

# 2. Exact application-name folder directly under Program Files / x86.
if ($installFolders.Count -eq 0 -and $appName -ne 'Unknown') {
    foreach ($root in $programRoots) {
        $candidate = Join-Path $root $appName
        if (Add-InstallFolderIfValid $candidate 'Exact application folder') { break }
    }
}

# 3. Vendor\Application folder.
if ($installFolders.Count -eq 0 -and $appVendor -ne 'Unknown' -and $appName -ne 'Unknown') {
    foreach ($root in $programRoots) {
        $candidate = Join-Path (Join-Path $root $appVendor) $appName
        if (Add-InstallFolderIfValid $candidate 'Vendor\\Application folder') { break }
    }
}

# 4/5. Search only one level below vendor folders and score candidates.
# This prevents selecting a random bin/data/config child directory.
if ($installFolders.Count -eq 0) {
    $folderCandidates = [System.Collections.Generic.List[object]]::new()

    foreach ($root in $programRoots) {
        try {
            foreach ($vendorDir in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
                # Application folder directly below a vendor folder.
                foreach ($dir in @(Get-ChildItem -LiteralPath $vendorDir.FullName -Directory -ErrorAction SilentlyContinue)) {
                    $score = 0
                    $n = $dir.Name

                    if ($appName -ne 'Unknown' -and $n -ieq $appName) { $score += 1000 }
                    elseif ($appName -ne 'Unknown' -and $n -like "*$appName*") { $score += 500 }

                    if ($appVendor -ne 'Unknown' -and $vendorDir.Name -ieq $appVendor) { $score += 300 }
                    elseif ($appVendor -ne 'Unknown' -and $vendorDir.Name -like "*$appVendor*") { $score += 150 }

                    if ($score -gt 0) {
                        [void]$folderCandidates.Add([pscustomobject]@{
                            Score = $score
                            Path  = $dir.FullName
                            Type  = 'Vendor\\Application'
                        })
                    }
                }

                # Vendor folder itself is only a last-resort candidate.
                $vendorScore = 0
                if ($appVendor -ne 'Unknown' -and $vendorDir.Name -ieq $appVendor) { $vendorScore = 100 }
                elseif ($appVendor -ne 'Unknown' -and $vendorDir.Name -like "*$appVendor*") { $vendorScore = 50 }

                if ($vendorScore -gt 0) {
                    [void]$folderCandidates.Add([pscustomobject]@{
                        Score = $vendorScore
                        Path  = $vendorDir.FullName
                        Type  = 'Vendor folder (last resort)'
                    })
                }
            }
        } catch {
            Write-Log "Unable to inspect $root : $($_.Exception.Message)" 'WARN'
        }
    }

    $best = $folderCandidates |
        Sort-Object @{Expression='Score';Descending=$true}, @{Expression={($_.Path -split '\\').Count};Ascending=$true} |
        Select-Object -First 1

    if ($best) {
        [void]$installFolders.Add((Resolve-Path -LiteralPath $best.Path).Path)
        Write-Log "Main installation folder detected: $($best.Path) [$($best.Type)]" 'OK'
    }
}

if ($installFolders.Count -eq 0) {
    Write-Log 'Main installation folder was not found in Program Files or Program Files (x86).' 'WARN'
}

# Installation-directory size, NOT Package\Files source size.
function Get-FolderSizeBytes {
    param([string]$Path)
    [int64]$total = 0
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return 0L }
    try {
        foreach ($file in @(Get-ChildItem -LiteralPath $Path -File -Recurse -ErrorAction SilentlyContinue)) {
            $total += [int64]$file.Length
        }
    } catch {}
    return $total
}

$installSizeBytes = if ($installFolders.Count -gt 0) { Get-FolderSizeBytes $installFolders[0] } else { 0L }
if ($installSizeBytes -gt 0) {
    if ($installSizeBytes -ge 1GB) { $installationSize = '{0:N2} GB' -f ($installSizeBytes / 1GB) }
    elseif ($installSizeBytes -ge 1MB) { $installationSize = '{0:N2} MB' -f ($installSizeBytes / 1MB) }
    elseif ($installSizeBytes -ge 1KB) { $installationSize = '{0:N2} KB' -f ($installSizeBytes / 1KB) }
    else { $installationSize = '{0} Bytes' -f $installSizeBytes }
} else {
    $installationSize = 'NA'
}
Write-Log "Installation directory size: $installationSize" 'OK'
$packageSize = $installationSize
Write-Log "PID package size (installation directory): $packageSize" 'INFO'

# -----------------------------------------------------------------------------
# Installed executable properties
# -----------------------------------------------------------------------------
$filePropertyLines = [System.Collections.Generic.List[string]]::new()
foreach ($folder in @($installFolders | Select-Object -Unique)) {
    try {
        $appFiles = @(Get-ChildItem -LiteralPath $folder -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -ieq '.exe' -or $_.Extension -ieq '.dll' } |
            Sort-Object Length -Descending |
            Select-Object -First 10)

        foreach ($file in $appFiles) {
            try {
                $vi = $file.VersionInfo
                $line = "{0} | Product: {1} | FileVersion: {2} | Company: {3}" -f `
                    $file.FullName,
                    $(if($vi.ProductName){$vi.ProductName}else{'NA'}),
                    $(if($vi.FileVersion){$vi.FileVersion}else{'NA'}),
                    $(if($vi.CompanyName){$vi.CompanyName}else{'NA'})
                [void]$filePropertyLines.Add($line)
            } catch {}
        }
    } catch {}
}

# -----------------------------------------------------------------------------
# Start Menu shortcut discovery
# -----------------------------------------------------------------------------
$shortcutList = [System.Collections.Generic.List[object]]::new()
$shortcutRoots = @(
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -Unique

foreach ($root in $shortcutRoots) {
    try {
        $lnks = @(Get-ChildItem -LiteralPath $root -Filter '*.lnk' -File -Recurse -ErrorAction SilentlyContinue)
        foreach ($lnk in $lnks) {
            $base = $lnk.BaseName
            if (
                ($appName -ne 'Unknown' -and $base -like "*$appName*") -or
                ($appVendor -ne 'Unknown' -and $base -like "*$appVendor*")
            ) {
                [void]$shortcutList.Add([pscustomobject]@{
                    Name = $base
                    Location = $lnk.DirectoryName
                    FullPath = $lnk.FullName
                })
            }
        }
    } catch {}
}

$shortcuts = @($shortcutList | Sort-Object FullPath -Unique)
foreach ($s in $shortcuts) { Write-Log "Shortcut: $($s.Name) -> $($s.Location)" 'OK' }

# -----------------------------------------------------------------------------
# MSI ProductCode / properties
# -----------------------------------------------------------------------------
$productCode = if ($installed -and $installed.KeyName) { [string]$installed.KeyName } else { '' }
$uninstallHive = ''
$msiPropertyLines = [System.Collections.Generic.List[string]]::new()

if ($installed -and $productCode) {
    if ([string]$installed.Architecture -eq '32-bit') {
        $uninstallHive = 'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\' + $productCode
    } else {
        $uninstallHive = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' + $productCode
    }
}

foreach ($msi in $msis) {
    $installer = $null
    $database = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database = $installer.OpenDatabase($msi.FullName, 0)
        foreach ($propertyName in @('ProductName','ProductVersion','Manufacturer','ProductCode','UpgradeCode')) {
            try {
                $view = $database.OpenView("SELECT `Value` FROM Property WHERE `Property`='$propertyName'")
                $view.Execute()
                $record = $view.Fetch()
                if ($record) {
                    $value = [string]$record.StringData(1)
                    [void]$msiPropertyLines.Add("$propertyName = $value")
                    if ($propertyName -eq 'ProductCode' -and -not $productCode) { $productCode = $value }
                }
                $view.Close()
            } catch {}
        }
    } catch {
        Write-Log "Could not read MSI properties from $($msi.Name): $($_.Exception.Message)" 'WARN'
    } finally {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($database) } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($installer) } catch {}
    }
}

if ($productCode -and -not $uninstallHive) {
    $uninstallHive = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' + $productCode
}

# -----------------------------------------------------------------------------
# Package commands
# -----------------------------------------------------------------------------
$installCommand = 'Deploy-Application.exe'
$uninstallCommand = 'Deploy-Application.exe Uninstall'

if ($serviceUI) {
    $installCommand = 'ServiceUI.exe -Process:explorer.exe Deploy-Application.exe'
    Write-Log 'ServiceUI installation command selected.' 'OK'
}

# -----------------------------------------------------------------------------
# Word helpers. IMPORTANT: no Word COM .Count and no .Item() calls.
# -----------------------------------------------------------------------------
function Get-WordTables {
    param($Document)
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($table in $Document.Tables) { [void]$list.Add($table) }
    return @($list)
}

function Get-WordParagraphs {
    param($Document)
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($paragraph in $Document.Paragraphs) { [void]$list.Add($paragraph) }
    return @($list)
}

# Font color used for generated VALUE text. This is taken from the
# existing Application Name value in the template, so generated values
# use exactly the same font color as the template's application name.
$script:GeneratedValueFontColor = $null

function Get-TemplateValueFontColor {
    param($Document)

    # Safely read the font colour from the value cell next to the
    # "Application Name" label. Do not use Row.Cells.Count or row indexes;
    # those COM properties can fail with merged Word-table cells.
    try {
        foreach ($table in $Document.Tables) {
            $cells = @($table.Range.Cells)

            for ($i = 0; $i -lt $cells.Count; $i++) {
                try {
                    $label = (($cells[$i].Range.Text -replace '[\r\a]+$', '') -replace '\s+', ' ').Trim()
                    if ($label -ne 'Application Name') {
                        continue
                    }

                    # In the Application Details table the value cell is the
                    # immediately following cell in document order.
                    if (($i + 1) -lt $cells.Count) {
                        $valueRange = $cells[$i + 1].Range.Duplicate()
                        $color = $valueRange.Font.Color

                        # Word can return wdUndefined (-9999999) when a range
                        # contains mixed formatting. In that case use the
                        # first character's colour instead.
                        if ($color -eq -9999999) {
                            $charRange = $valueRange.Duplicate()
                            if ($charRange.End -gt $charRange.Start) {
                                $charRange.End = $charRange.Start + 1
                                $color = $charRange.Font.Color
                            }
                        }

                        if ($color -ne -9999999 -and $null -ne $color) {
                            return $color
                        }
                    }
                }
                catch {
                    # Ignore one malformed/merged cell and continue scanning.
                    continue
                }
            }
        }
    }
    catch {
        Write-Log "Could not read Application Name font color: $($_.Exception.Message)" 'WARN'
    }

    # Safe fallback: Word automatic font colour.
    return -16777216
}

function Apply-GeneratedValueColor {
    param($Range)
    try {
        if ($null -ne $script:GeneratedValueFontColor) {
            $Range.Font.Color = $script:GeneratedValueFontColor
        }
    }
    catch {
        Write-Log "Could not apply generated-value font color: $($_.Exception.Message)" 'WARN'
    }
}

function Set-CellText {
    param($Cell, [AllowNull()][string]$Text)
    if ($null -eq $Text) { $Text = '' }
    $range = $Cell.Range.Duplicate()
    $range.End = $range.End - 1
    $range.Text = $Text
    Apply-GeneratedValueColor $range
}

function Set-ParagraphText {
    param($Paragraph, [AllowNull()][string]$Text)

    if ($null -eq $Text) { $Text = '' }

    try {
        $range = $Paragraph.Range.Duplicate()

        # Word throws "This is not a valid action for the end of a row"
        # when text is assigned to a range ending on a table row marker.
        # Do not modify table-contained paragraphs here; table cells are
        # handled separately by Set-CellText.
        $wdWithInTable = 12
        if ($range.Information($wdWithInTable)) {
            return $false
        }

        if ($range.End -gt $range.Start) {
            $range.End = $range.End - 1
        }

        $range.Text = $Text
        Apply-GeneratedValueColor $range
        return $true
    }
    catch {
        Write-Log "Paragraph update skipped: $($_.Exception.Message)" 'WARN'
        return $false
    }
}

function Set-ParagraphAfterHeading {
    param(
        $Paragraphs,
        [string]$Heading,
        [string]$Value
    )

    $normalizedHeading = ($Heading -replace '\s+', ' ').Trim()

    for ($i = 0; $i -lt $Paragraphs.Count; $i++) {
        try {
            $current = (($Paragraphs[$i].Range.Text -replace '[\r\a]+$', '') -replace '\s+', ' ').Trim()
            if ($current -ne $normalizedHeading) {
                continue
            }

            # Look ahead for the first standalone value paragraph.
            # The TOC occurrence is ignored because its following text is
            # not "NA" and is normally a page number/dotted leader.
            $max = [Math]::Min($Paragraphs.Count - 1, $i + 8)

            for ($j = $i + 1; $j -le $max; $j++) {
                $nextRange = $Paragraphs[$j].Range.Duplicate()

                if ($nextRange.Information(12)) {
                    continue
                }

                $nextText = (($nextRange.Text -replace '[\r\a]+$', '') -replace '\s+', ' ').Trim()

                if ($nextText -eq 'NA' -or [string]::IsNullOrWhiteSpace($nextText)) {
                    if (Set-ParagraphText $Paragraphs[$j] $Value) {
                        Write-Log "Updated section '$Heading'." 'OK'
                    }
                    return
                }

                # Stop if another section heading has started.
                if ($nextText -match '^(Dependencies|Environment variables|ODBC|Rights/Permissions|MST|MSI Properties|Custom Actions|Known Errors|INTUNE Detection Method|Supported Installation & Uninstallation Method|Package Remarks|Files|Shortcuts)$') {
                    break
                }
            }
        }
        catch {
            Write-Log "Could not update section '$Heading': $($_.Exception.Message)" 'WARN'
        }
    }
}

# ----------------------------------------------------------------------------- 
# Word generation
# -----------------------------------------------------------------------------
$Word = $null
$Document = $null
$outputPath = Join-Path $DocsDir ($PackageName + '_PID.docx')

try {
    if (Test-Path -LiteralPath $outputPath -PathType Leaf) {
        Remove-Item -LiteralPath $outputPath -Force
    }

    Copy-Item -LiteralPath $TemplatePath -Destination $outputPath -Force
    Write-Log "Template copied to: $outputPath" 'OK'

    try {
        $Word = New-Object -ComObject Word.Application
    } catch {
        Fail 'Microsoft Word could not be started. Microsoft Word desktop must be installed.'
    }

    $Word.Visible = $false
    $Word.DisplayAlerts = 0

    $Document = $Word.Documents.Open($outputPath, $false, $false)
    Write-Log 'Word document opened.' 'OK'

    # Read the template's Application Name value color before replacing content.
    # All generated VALUE text will use this same color; template headings remain unchanged.
    $script:GeneratedValueFontColor = Get-TemplateValueFontColor $Document
    Write-Log 'Generated value font color initialized from template Application Name.' 'OK'

    # Convert Word COM collections to ordinary arrays ONCE.
    $Tables = Get-WordTables $Document
    $Paragraphs = Get-WordParagraphs $Document

    Write-Log "Word tables detected: $($Tables.Count)" 'OK'
    Write-Log "Word paragraphs detected: $($Paragraphs.Count)" 'OK'

    if ($Tables.Count -lt 9) {
        Fail "The supplied template contains $($Tables.Count) tables. The uploaded PID template is expected to contain at least 9 tables."
    }

    # Table 1: List of Changes
    $t = $Tables[0]
    Set-CellText $t.Cell(2,1) '1.0'
    Set-CellText $t.Cell(2,2) (Get-Date -Format 'dd/MM/yyyy')
    Set-CellText $t.Cell(2,3) $PackageName
    Set-CellText $t.Cell(2,4) 'HDS RMS'

    # Table 2: Application Details
    $t = $Tables[1]
    Set-CellText $t.Cell(2,2) $PackageName
    Set-CellText $t.Cell(3,2) $appVendor
    Set-CellText $t.Cell(4,2) $appName
    Set-CellText $t.Cell(5,2) $appVersion
    Set-CellText $t.Cell(6,2) $(if($installed -and $installed.DisplayIcon){[string]$installed.DisplayIcon}else{'NA'})

    # Table 3: Contacts - leave existing template data intact unless script contains values.
    $requester = Get-ScriptVariable $DeployScript @('requesterName','RequesterName','requester','Requester')
    $owner     = Get-ScriptVariable $DeployScript @('applicationOwner','ApplicationOwner','ownerName','OwnerName')
    $phone     = Get-ScriptVariable $DeployScript @('phoneNumber','PhoneNumber','phone','Phone')
    $email     = Get-ScriptVariable $DeployScript @('email','Email','requesterEmail','RequesterEmail')
    $t = $Tables[2]
    if ($owner) { Set-CellText $t.Cell(3,2) $owner }
    if ($requester) { Set-CellText $t.Cell(7,2) $requester }
    if ($phone) { Set-CellText $t.Cell(4,2) $phone; Set-CellText $t.Cell(8,2) $phone }
    if ($email) { Set-CellText $t.Cell(5,2) $email; Set-CellText $t.Cell(9,2) $email }

    # Table 4: OS
    $t = $Tables[3]
    Set-CellText $t.Cell(2,1) 'Windows 11 64 bit'
    Set-CellText $t.Cell(2,2) '6.3'

    # Table 5: Package details
    $t = $Tables[4]
    Set-CellText $t.Cell(2,2) $packageSize
    Set-CellText $t.Cell(3,2) $packageType
    Set-CellText $t.Cell(4,2) 'NA'
    Set-CellText $t.Cell(5,2) 'NA'

    # Files table
    $t = $Tables[5]
    $applicationText = if ($installFolders.Count -gt 0) { $installFolders[0] } else { 'NA - Application not detected on this machine' }
    Set-CellText $t.Cell(2,2) $applicationText
    Set-CellText $t.Cell(3,2) 'NA'
    Set-CellText $t.Cell(4,2) 'NA'

    # Shortcuts table. Keep only the header + one data row, then add rows as needed.
    $t = $Tables[6]
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $t.Rows) { [void]$rows.Add($row) }
    while ($rows.Count -gt 2) {
        $rows[$rows.Count - 1].Delete()
        $rows.RemoveAt($rows.Count - 1)
    }

    if ($shortcuts.Count -eq 0) {
        Set-CellText $t.Cell(2,1) 'NA'
        Set-CellText $t.Cell(2,2) 'NA'
    } else {
        $first = $true
        foreach ($shortcut in $shortcuts) {
            if ($first) {
                $row = $rows[1]
                $first = $false
            } else {
                $row = $t.Rows.Add()
            }
            Set-CellText $row.Cells.Item(1) $shortcut.Name
            Set-CellText $row.Cells.Item(2) $shortcut.Location
        }
    }

    # Section values are located by heading text instead of hard-coded
    # paragraph numbers. This prevents edits to a table-row paragraph from
    # triggering Word's "end of a row" COM exception.
    Set-ParagraphAfterHeading $Paragraphs 'Dependencies' 'NA'
    Set-ParagraphAfterHeading $Paragraphs 'Environment variables' 'NA'
    Set-ParagraphAfterHeading $Paragraphs 'ODBC' 'NA'
    Set-ParagraphAfterHeading $Paragraphs 'Rights/Permissions' 'NA'
    Set-ParagraphAfterHeading $Paragraphs 'MST' 'NA'
    Set-ParagraphAfterHeading $Paragraphs 'MSI Properties' $(if($msiPropertyLines.Count -gt 0){$msiPropertyLines -join "`r"}else{'NA'})
    Set-ParagraphAfterHeading $Paragraphs 'Custom Actions' 'NA'
    Set-ParagraphAfterHeading $Paragraphs 'Known Errors' 'NA'
    Set-ParagraphAfterHeading $Paragraphs 'Package Remarks' 'NA'
    Set-ParagraphAfterHeading $Paragraphs 'LOGS LOCATION:' 'C:\Windows\Logs\Software'

    # Intune detection table - show BOTH 64-bit and 32-bit registry matches
    # whose DisplayVersion exactly equals the package/application version.
    $t = $Tables[7]
    $detectEntries = @($matchingEntries | Sort-Object @{Expression={if($_.Architecture -eq '64-bit'){0}else{1}}})

    if ($detectEntries.Count -gt 0) {
        $productLines = [System.Collections.Generic.List[string]]::new()
        $registryLines = [System.Collections.Generic.List[string]]::new()

        foreach ($entry in $detectEntries) {
            $displayVersion = if ($entry.DisplayVersion) { $entry.DisplayVersion } else { 'NA' }
            $productKey = if ($entry.KeyName) { $entry.KeyName } else { 'NA' }
            $arch = $entry.Architecture

            # Intune detection contains only the registry key/product code and
            # the uninstall hive/inventory key. Display Name is intentionally
            # not written to the document.
            [void]$productLines.Add("${arch}: $productKey")
            [void]$registryLines.Add("${arch}: $($entry.RegistryPath)")
            [void]$registryLines.Add("DisplayVersion = `"$displayVersion`"")

            if ($entry.UninstallString) {
                [void]$registryLines.Add("UninstallString: $($entry.UninstallString)")
            }
        }

        Set-CellText $t.Cell(2,1) ($productLines -join "`r")
        Set-CellText $t.Cell(2,2) ($registryLines -join "`r")
    } else {
        # Explicitly state that both registry views were checked and that no
        # exact package-version match was found.
        Set-CellText $t.Cell(2,1) "64-bit: No exact match for DisplayVersion $appVersion`r32-bit: No exact match for DisplayVersion $appVersion"
        Set-CellText $t.Cell(2,2) "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*`rBoth 64-bit and 32-bit registry views checked"
    }

    # Installation/uninstallation table
    $t = $Tables[8]
    Set-CellText $t.Cell(2,2) 'Yes'
    Set-CellText $t.Cell(3,2) 'Yes'
    Set-CellText $t.Cell(2,3) $installCommand
    Set-CellText $t.Cell(3,3) $uninstallCommand

    # Save without updating Word fields/TOC. This avoids another class of Word COM
    # errors. The existing TOC can be refreshed manually in Word if required.
    $Document.Save()
    Write-Log "Document saved: $outputPath" 'OK'

    if ($OpenDocument) {
        $Document.Close($false)
        $Document = $null
        $Word.Quit()
        $Word = $null
        Start-Process -FilePath $outputPath
    }
}
catch {
    Write-Log "DOCUMENT GENERATION FAILED: $($_.Exception.Message)" 'ERROR'
    Write-Log "Script line: $($_.InvocationInfo.ScriptLineNumber)" 'ERROR'
    throw
}
finally {
    if ($Document) { try { $Document.Close($false) } catch {} }
    if ($Word) { try { $Word.Quit() } catch {} }
    try { [GC]::Collect(); [GC]::WaitForPendingFinalizers() } catch {}

    try {
        $logPath = Join-Path $LogDir ($PackageName + '_PID_Generation.log')
        $LogLines | Set-Content -LiteralPath $logPath -Encoding UTF8
        Write-Host "Log: $logPath"
    } catch {}
}

$seconds = [math]::Round(((Get-Date) - $StartTime).TotalSeconds, 2)
Write-Log "PID generation completed in $seconds seconds." 'OK'
Write-Host "OUTPUT: $outputPath" -ForegroundColor Green
