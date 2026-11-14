[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$PackageRoot
)

$ErrorActionPreference = 'Stop'
$script:results = @()

function Add-QA {
    param([string]$Name,[string]$Value,[bool]$Pass,[string]$Details='')
    $script:results += [pscustomobject]@{
        Check=$Name; Value=$Value; Status=if($Pass){'PASS'}else{'FAIL'}; Details=$Details
    }
}

function Get-DeployVariable {
    param([string]$Text,[string]$Name)
    $pattern = '(?im)^\s*(?:\[[^\]]+\]\s*)?\$' + [regex]::Escape($Name) + '\s*=\s*[''"]([^''"]*)[''"]'
    $m = [regex]::Match($Text,$pattern)
    if($m.Success){ return $m.Groups[1].Value.Trim() }
    return ''
}

function Get-DocxText {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry('word/document.xml')
        if(-not $entry){ return '' }
        $reader = New-Object IO.StreamReader($entry.Open())
        try {
            $xmlText = $reader.ReadToEnd()
        } finally { $reader.Dispose() }
        $xml = New-Object Xml.XmlDocument
        $xml.LoadXml($xmlText)
        return (($xml.InnerText -replace '\s+',' ').Trim())
    } finally { $zip.Dispose() }
}

try {
    $selectedRoot = (Resolve-Path -LiteralPath $PackageRoot).Path
    if(Test-Path -LiteralPath (Join-Path $selectedRoot 'Deploy-Application.ps1')){
        $pkg=$selectedRoot; $packageRoot=Split-Path -Parent $pkg
    } elseif(Test-Path -LiteralPath (Join-Path $selectedRoot 'Package')){
        $packageRoot=$selectedRoot; $pkg=Join-Path $packageRoot 'Package'
    } else {
        throw "Invalid package folder. Select the package root or its Package folder. Selected: $selectedRoot"
    }

    $packageName = Split-Path -Leaf $packageRoot
    $docs=Join-Path $packageRoot 'Docs'
    $deploy=Join-Path $pkg 'Deploy-Application.ps1'
    $files=Join-Path $pkg 'Files'
    $toolkit=Join-Path $pkg 'AppDeployToolkit'

    Add-QA 'Selected folder' $selectedRoot $true 'QA target'
    Add-QA 'Package name' $packageName ($packageName -match '.+_.+_.+_P1') 'Package folder naming format'
    Add-QA 'Package folder' $pkg (Test-Path -LiteralPath $pkg) 'Package folder exists'
    Add-QA 'Deploy-Application.ps1' $deploy (Test-Path -LiteralPath $deploy) 'Required deployment script'
    Add-QA 'Deploy-Application.exe' (Join-Path $pkg 'Deploy-Application.exe') (Test-Path -LiteralPath (Join-Path $pkg 'Deploy-Application.exe')) 'Required launcher'
    Add-QA 'AppDeployToolkit folder' $toolkit (Test-Path -LiteralPath $toolkit) 'PSADT runtime'
    Add-QA 'Files folder' $files (Test-Path -LiteralPath $files) 'Installer/source files'

    if(Test-Path -LiteralPath $deploy){
        $raw=Get-Content -LiteralPath $deploy -Raw
        $values=@{}
        foreach($v in 'appVendor','appName','appVersion','appArch','appLang','appRevision','appScriptAuthor','installName','installTitle'){
            $values[$v]=Get-DeployVariable -Text $raw -Name $v
            Add-QA "Deploy variable: $v" $values[$v] (-not [string]::IsNullOrWhiteSpace($values[$v])) 'Read from Deploy-Application.ps1'
        }

        Add-QA 'InstallName equals package name' $values['installName'] ($values['installName'] -eq $packageName) "Expected: $packageName"
        Add-QA 'InstallTitle equals package name' $values['installTitle'] ($values['installTitle'] -eq $packageName) "Expected: $packageName"

        if($values['appVendor']){
            Add-QA 'Vendor appears in package name' $values['appVendor'] ($packageName -match [regex]::Escape(($values['appVendor'] -replace '\s+','_'))) 'Package/Deploy consistency'
        }
        if($values['appName']){
            Add-QA 'Application appears in package name' $values['appName'] ($packageName -match [regex]::Escape(($values['appName'] -replace '\s+','_'))) 'Package/Deploy consistency'
        }
        if($values['appVersion']){
            Add-QA 'Version appears in package name' $values['appVersion'] ($packageName -match [regex]::Escape($values['appVersion'])) 'Package/Deploy consistency'
        }

        $tokens=$null;$errs=$null
        [System.Management.Automation.Language.Parser]::ParseFile($deploy,[ref]$tokens,[ref]$errs)|Out-Null
        Add-QA 'PowerShell syntax' "$($errs.Count) error(s)" ($errs.Count -eq 0) 'Parser validation'

        $installCmd = [regex]::Match($raw,'(?im)^\s*(?:Execute-MSI|Execute-Process|Start-ADTProcess).*$')
        Add-QA 'Installation command detected' $installCmd.Value ($installCmd.Success) 'PSADT installation command'

        $uninstallMarker = [regex]::Match($raw,'(?is)Pre-Uninstallation.*?(?:Execute-MSI|Execute-Process|Start-ADTProcess).+')
        Add-QA 'Uninstallation command detected' $(if($uninstallMarker.Success){'Detected'}else{''}) $uninstallMarker.Success 'PSADT uninstallation area'

        if(Test-Path -LiteralPath $files){
            $installerFiles=@(Get-ChildItem -LiteralPath $files -File -Recurse | Where-Object { $_.Extension -in '.msi','.exe' })
            Add-QA 'Installer file count' "$($installerFiles.Count)" ($installerFiles.Count -ge 1) 'MSI/EXE found under Package\Files'
            $referencedInstallers=@($installerFiles | Where-Object { $raw -match [regex]::Escape($_.Name) })
            $names=($referencedInstallers | Select-Object -ExpandProperty Name) -join '; '
            Add-QA 'Installer name referenced in Deploy' $names ($referencedInstallers.Count -ge 1) 'At least one installer filename under Package\Files must be referenced by Deploy-Application.ps1'
        }
    }

    $pidDocument = if(Test-Path -LiteralPath $docs){ Get-ChildItem -LiteralPath $docs -Filter '*.docx' -File | Select-Object -First 1 } else { $null }
    Add-QA 'PID document' $(if($pidDocument){$pidDocument.FullName}else{''}) ($null -ne $pidDocument) 'PID document exists'

    if($pidDocument){
        $pidText=Get-DocxText -Path $pidDocument.FullName
        Add-QA 'PID readable' "$($pidDocument.Length) bytes" (-not [string]::IsNullOrWhiteSpace($pidText)) 'Word document XML read'
        if($values){
            foreach($pair in @(
                @{N='PID vendor matches Deploy';V=$values['appVendor']},
                @{N='PID application matches Deploy';V=$values['appName']},
                @{N='PID version matches Deploy';V=$values['appVersion']},
                @{N='PID package name matches folder';V=$packageName}
            )){
                if([string]::IsNullOrWhiteSpace($pair.V)){ Add-QA $pair.N '' $false 'Expected value is empty'; continue }
                $found=$pidText.IndexOf($pair.V,[StringComparison]::OrdinalIgnoreCase) -ge 0
                Add-QA $pair.N $pair.V $found 'PID text compared with package/deployment data'
            }
        }
        if(Test-Path -LiteralPath $files){
            $mainInstaller=Get-ChildItem -LiteralPath $files -File -Recurse | Where-Object { $_.Extension -in '.msi','.exe' } | Select-Object -First 1
            if($mainInstaller){
                $found=$pidText.IndexOf($mainInstaller.Name,[StringComparison]::OrdinalIgnoreCase) -ge 0
                Add-QA 'PID installer name matches package' $mainInstaller.Name $found 'Installer filename compared with PID'
            }
        }
    }

    $intuneLogo=Join-Path $packageRoot 'Intune\AppLogo.png'
    $toolkitLogo=Join-Path $toolkit 'AppLogo.png'
    if((Test-Path -LiteralPath $intuneLogo) -or (Test-Path -LiteralPath $toolkitLogo)){
        Add-QA 'AppLogo consistency' "$intuneLogo | $toolkitLogo" ((Test-Path $intuneLogo) -and (Test-Path $toolkitLogo)) 'Optional logo should exist in both locations when enabled'
    } else {
        Add-QA 'AppLogo' 'Not configured' $true 'Optional feature'
    }

    if($script:results.Count -eq 0){ throw 'QA produced no results.' }
    $script:results
}
catch {
    $script:results += [pscustomobject]@{Check='QA runtime';Value='';Status='FAIL';Details=$_.Exception.Message}
    $script:results
}
