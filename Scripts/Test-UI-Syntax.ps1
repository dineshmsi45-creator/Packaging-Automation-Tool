$path=Join-Path $PSScriptRoot 'Start-PackageUI.ps1'
$tokens=$null;$errors=$null
[void][System.Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
if($errors.Count -eq 0){Write-Host 'Start-PackageUI.ps1 syntax OK' -ForegroundColor Green;exit 0}
$errors|Format-List;exit 1
