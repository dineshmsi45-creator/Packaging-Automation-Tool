[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$MsiPath,

    [Parameter(Mandatory = $true)]
    [string]$Vendor,

    [Parameter(Mandatory = $true)]
    [string]$AppName,

    [Parameter(Mandatory = $true)]
    [string]$Version,

    [Parameter(Mandatory = $true)]
    [ValidateSet("x86", "x64", "ARM64")]
    [string]$Architecture
)

$ErrorActionPreference = "Stop"

# ------------------------------------------------------------
# HDS RMS STANDARD MST GENERATOR
# The original vendor MSI is NEVER modified.
# Output:
#   Vendor.msi                         <- original file, unchanged
#   Vendor_App_Version_P1_x64.mst     <- generated transform
#   Vendor_App_Version_P1_x64.log     <- processing log
# ------------------------------------------------------------

if (-not (Test-Path -LiteralPath $MsiPath -PathType Leaf)) {
    throw "MSI file not found: $MsiPath"
}

$MsiPath = [System.IO.Path]::GetFullPath($MsiPath)

if ([System.IO.Path]::GetExtension($MsiPath) -ine ".msi") {
    throw "Input file must be an MSI file."
}

$SourceFolder = Split-Path -Path $MsiPath -Parent
$PackageName = "{0}_{1}_{2}_P1_{3}" -f $Vendor, $AppName, $Version, $Architecture

$OutputMst = Join-Path $SourceFolder ($PackageName + ".mst")
$LogPath   = Join-Path $SourceFolder ($PackageName + ".log")

$TempMsi = Join-Path $env:TEMP ("HDSRMS_" + [guid]::NewGuid().ToString("N") + ".msi")

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Line = "{0} [{1}] {2}" -f $Time, $Level, $Message

    Add-Content -LiteralPath $LogPath -Value $Line

    switch ($Level) {
        "ERROR" {
            Write-Host $Line -ForegroundColor Red
        }
        "WARNING" {
            Write-Host $Line -ForegroundColor Yellow
        }
        "SUCCESS" {
            Write-Host $Line -ForegroundColor Green
        }
        default {
            Write-Host $Line
        }
    }
}

if (Test-Path -LiteralPath $LogPath) {
    Remove-Item -LiteralPath $LogPath -Force
}

Write-Log "============================================================"
Write-Log "HDS RMS STANDARD MST GENERATOR"
Write-Log "============================================================"
Write-Log ("Vendor MSI : {0}" -f $MsiPath)
Write-Log ("MST        : {0}" -f $OutputMst)
Write-Log ("Log        : {0}" -f $LogPath)
Write-Log ("Subject    : {0}" -f $PackageName)
Write-Log "Author     : HDS RMS"
Write-Log "============================================================"

# Standard properties to add/update in the transform.
$Properties = [ordered]@{
    REBOOT                   = "ReallySuppress"
    REINSTALLMODE            = "omus"
    ALLUSERS                 = "1"
    ARPNOMODIFY              = "NO"
    ARPNOREMOVE              = "Yes"
    MSIRESTARTMANAGERCONTROL = "Disable"
    PROMPTROLLBACKCOST       = "D"
    DISABLEMEDIA             = "1"
    LIMITUI                  = "1"
    ROOTDRIVE                = "C:\"
}

$Added = 0
$Updated = 0
$Skipped = 0
$DesktopRemoved = 0
$AdvertisedFound = 0

$Installer = $null
$OriginalDatabase = $null
$ModifiedDatabase = $null
$View = $null

try {
    # --------------------------------------------------------
    # Create a working copy of the vendor MSI.
    # --------------------------------------------------------
    Write-Log "Creating temporary working copy of vendor MSI."
    Copy-Item -LiteralPath $MsiPath -Destination $TempMsi -Force
    Write-Log ("Temporary MSI: {0}" -f $TempMsi)

    $Installer = New-Object -ComObject WindowsInstaller.Installer

    # Mode 0 = read-only
    $OriginalDatabase = $Installer.OpenDatabase($MsiPath, 0)

    # Mode 2 = direct read/write
    $ModifiedDatabase = $Installer.OpenDatabase($TempMsi, 2)

    # --------------------------------------------------------
    # STANDARD MSI PROPERTIES
    # --------------------------------------------------------
    Write-Log "------------------------------------------------------------"
    Write-Log "PROCESSING STANDARD MSI PROPERTIES"
    Write-Log "------------------------------------------------------------"

    foreach ($PropertyName in $Properties.Keys) {

        $DesiredValue = [string]$Properties[$PropertyName]
        $View = $null
        $Record = $null
        $Exists = $false
        $ExistingValue = ""

        try {
            $Query = "SELECT ``Value`` FROM ``Property`` WHERE ``Property`` = ?"

            $View = $ModifiedDatabase.OpenView($Query)

            $Record = $Installer.CreateRecord(1)
            $Record.StringData(1) = $PropertyName

            $View.Execute($Record)
            $Record = $View.Fetch()

            if ($null -ne $Record) {
                $Exists = $true
                $ExistingValue = [string]$Record.StringData(1)
            }
        }
        finally {
            if ($null -ne $View) {
                $View.Close()
            }

            $View = $null
            $Record = $null
        }

        if (-not $Exists) {

            $View = $ModifiedDatabase.OpenView(
                "INSERT INTO ``Property`` (``Property``, ``Value``) VALUES (?, ?)"
            )

            $NewRecord = $Installer.CreateRecord(2)
            $NewRecord.StringData(1) = $PropertyName
            $NewRecord.StringData(2) = $DesiredValue

            $View.Execute($NewRecord)
            $View.Close()
            $View = $null

            $Added++

            Write-Log ("[ADD] {0} = {1}" -f $PropertyName, $DesiredValue)
        }
        elseif ($ExistingValue -ieq $DesiredValue) {

            $Skipped++

            Write-Log (
                "[SKIP] {0} already has value '{1}'." -f
                $PropertyName,
                $ExistingValue
            )
        }
        else {

            $View = $ModifiedDatabase.OpenView(
                "UPDATE ``Property`` SET ``Value`` = ? WHERE ``Property`` = ?"
            )

            $UpdateRecord = $Installer.CreateRecord(2)
            $UpdateRecord.StringData(1) = $DesiredValue
            $UpdateRecord.StringData(2) = $PropertyName

            $View.Execute($UpdateRecord)
            $View.Close()
            $View = $null

            $Updated++

            Write-Log (
                "[UPDATE] {0}: '{1}' -> '{2}'" -f
                $PropertyName,
                $ExistingValue,
                $DesiredValue
            )
        }
    }

    # --------------------------------------------------------
    # FORCE ADVERTISED SHORTCUTS TO REGULAR SHORTCUTS
    #
    # DISABLEADVTSHORTCUTS=1 is the supported Windows Installer
    # method. It causes advertised shortcuts to be created as
    # regular shortcuts instead of installation-on-demand
    # shortcuts.
    # --------------------------------------------------------
    Write-Log "------------------------------------------------------------"
    Write-Log "CHECKING SHORTCUTS"
    Write-Log "------------------------------------------------------------"

    $ShortcutView = $null
    $FeatureView = $null
    $ShortcutRows = @()
    $FeatureNames = @{}

    # Read Feature table.
    try {
        $FeatureView = $ModifiedDatabase.OpenView(
            "SELECT ``Feature`` FROM ``Feature``"
        )

        $FeatureView.Execute()

        while ($FeatureRecord = $FeatureView.Fetch()) {
            $FeatureName = [string]$FeatureRecord.StringData(1)

            if (-not [string]::IsNullOrWhiteSpace($FeatureName)) {
                $FeatureNames[$FeatureName] = $true
            }
        }
    }
    catch {
        Write-Log (
            "Could not read Feature table: {0}" -f $_.Exception.Message
        ) "WARNING"
    }
    finally {
        if ($null -ne $FeatureView) {
            $FeatureView.Close()
        }

        $FeatureView = $null
    }

    # Read Shortcut table.
    try {
        $ShortcutView = $ModifiedDatabase.OpenView(
            "SELECT ``Shortcut``, ``Directory_``, ``Target``, ``Name`` FROM ``Shortcut``"
        )

        $ShortcutView.Execute()

        while ($ShortcutRecord = $ShortcutView.Fetch()) {

            $ShortcutRows += [PSCustomObject]@{
                Shortcut  = [string]$ShortcutRecord.StringData(1)
                Directory = [string]$ShortcutRecord.StringData(2)
                Target    = [string]$ShortcutRecord.StringData(3)
                Name      = [string]$ShortcutRecord.StringData(4)
            }
        }
    }
    catch {
        Write-Log (
            "Could not read Shortcut table: {0}" -f $_.Exception.Message
        ) "WARNING"
    }
    finally {
        if ($null -ne $ShortcutView) {
            $ShortcutView.Close()
        }

        $ShortcutView = $null
    }

    # --------------------------------------------------------
    # Remove Desktop shortcuts.
    # --------------------------------------------------------
    foreach ($Shortcut in $ShortcutRows) {

        if ($Shortcut.Directory -ieq "DesktopFolder") {

            Write-Log ("Desktop shortcut found: {0}" -f $Shortcut.Name)

            $DeleteView = $null

            try {
                $DeleteView = $ModifiedDatabase.OpenView(
                    "DELETE FROM ``Shortcut`` WHERE ``Shortcut`` = ?"
                )

                $DeleteRecord = $Installer.CreateRecord(1)
                $DeleteRecord.StringData(1) = $Shortcut.Shortcut

                $DeleteView.Execute($DeleteRecord)
            }
            finally {
                if ($null -ne $DeleteView) {
                    $DeleteView.Close()
                }

                $DeleteView = $null
            }

            $DesktopRemoved++

            Write-Log (
                "Desktop shortcut removed: {0}" -f $Shortcut.Name
            ) "SUCCESS"
        }
    }

    if ($DesktopRemoved -eq 0) {
        Write-Log "No Desktop shortcuts found."
    }

    # --------------------------------------------------------
    # Detect advertised shortcuts.
    #
    # In the Shortcut table, an advertised shortcut has Target
    # equal to a Feature identifier. A non-advertised shortcut
    # normally contains a formatted string such as [#FileKey].
    # --------------------------------------------------------
    foreach ($Shortcut in $ShortcutRows) {

        if ($Shortcut.Directory -ieq "DesktopFolder") {
            continue
        }

        $Target = $Shortcut.Target

        if ([string]::IsNullOrWhiteSpace($Target)) {
            continue
        }

        if ($FeatureNames.ContainsKey($Target)) {

            $AdvertisedFound++

            Write-Log (
                "Advertised shortcut found: {0}" -f $Shortcut.Name
            ) "WARNING"

            Write-Log (
                "Advertised target/feature: {0}" -f $Target
            ) "WARNING"

            Write-Log (
                "DISABLEADVTSHORTCUTS=1 will disable advertisement for this shortcut."
            ) "SUCCESS"
        }
    }

    if ($AdvertisedFound -eq 0) {
        Write-Log "No advertised shortcuts found."
    }
    else {
        Write-Log (
            "Total advertised shortcuts found: {0}" -f $AdvertisedFound
        )

        # Add/update DISABLEADVTSHORTCUTS explicitly.
        $PropertyName = "DISABLEADVTSHORTCUTS"
        $DesiredValue = "1"

        $View = $ModifiedDatabase.OpenView(
            "SELECT ``Value`` FROM ``Property`` WHERE ``Property`` = ?"
        )

        $Record = $Installer.CreateRecord(1)
        $Record.StringData(1) = $PropertyName

        $View.Execute($Record)
        $Record = $View.Fetch()
        $View.Close()
        $View = $null

        if ($null -eq $Record) {

            $View = $ModifiedDatabase.OpenView(
                "INSERT INTO ``Property`` (``Property``, ``Value``) VALUES (?, ?)"
            )

            $NewRecord = $Installer.CreateRecord(2)
            $NewRecord.StringData(1) = $PropertyName
            $NewRecord.StringData(2) = $DesiredValue

            $View.Execute($NewRecord)
            $View.Close()
            $View = $null

            $Added++

            Write-Log "[ADD] DISABLEADVTSHORTCUTS = 1" "SUCCESS"
        }
        elseif ([string]$Record.StringData(1) -ne "1") {

            $View = $ModifiedDatabase.OpenView(
                "UPDATE ``Property`` SET ``Value`` = ? WHERE ``Property`` = ?"
            )

            $UpdateRecord = $Installer.CreateRecord(2)
            $UpdateRecord.StringData(1) = $DesiredValue
            $UpdateRecord.StringData(2) = $PropertyName

            $View.Execute($UpdateRecord)
            $View.Close()
            $View = $null

            $Updated++

            Write-Log "[UPDATE] DISABLEADVTSHORTCUTS -> 1" "SUCCESS"
        }
        else {
            Write-Log "[SKIP] DISABLEADVTSHORTCUTS is already 1."
        }
    }

    # --------------------------------------------------------
    # SAVE TEMPORARY MSI
    # --------------------------------------------------------
    
    Write-Log "Saving temporary MSI database."

    $ModifiedDatabase.Commit()

    $ModifiedDatabase = $null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # --------------------------------------------------------
    # SUMMARY INFORMATION
    # --------------------------------------------------------
    Write-Log "------------------------------------------------------------"
    Write-Log "UPDATING MSI SUMMARY INFORMATION"
    Write-Log "------------------------------------------------------------"

    $Summary = $Installer.SummaryInformation($TempMsi, 20)

    # PID_TITLE = 2
    $OldTitle = [string]$Summary.Property(2)

    if ($OldTitle -ne $PackageName) {
        $Summary.Property(2) = $PackageName
        Write-Log ("[UPDATE] Title: '{0}' -> '{1}'" -f $OldTitle, $PackageName)
    }
    else {
        Write-Log ("[SKIP] Title already '{0}'." -f $PackageName)
    }

    # PID_SUBJECT = 3
   $OldSubject = [string]$Summary.Property(3)
   if ($OldSubject -ne $PackageName) {
       $Summary.Property(3) = $PackageName
       Write-Log ("[UPDATE] Subject: '{0}' -> '{1}'" -f $OldSubject, $PackageName)
   }
   else {
       Write-Log ("[SKIP] Subject already '{0}'." -f $PackageName)
   }

   # PID_AUTHOR = 4
   $OldAuthor = [string]$Summary.Property(4)
   if ($OldAuthor -ne "HDS RMS") {
   $Summary.Property(4) = "HDS RMS"
   Write-Log ("[UPDATE] Author: '{0}' -> 'HDS RMS'" -f $OldAuthor)
    }
    else {
   Write-Log "[SKIP] Author already 'HDS RMS'."
    }
    # ------------------------------------------------------------
    # MSI SUMMARY INFORMATION SCHEMA
    # PID_PAGECOUNT = 14
    # Required value = 300
    # ------------------------------------------------------------
    $OldSchemaRaw = $Summary.Property(14)
    if ($null -eq $OldSchemaRaw -or [string]::IsNullOrWhiteSpace([string]$OldSchemaRaw)) {
   $OldSchema = 0
    }
    else {
   $OldSchema = [int]$OldSchemaRaw
    }
    if ($OldSchema -eq 300) {
   Write-Log "[SKIP] Schema already 300."
    }
    else {
   $Summary.Property(14) = 300
   Write-Log ("[UPDATE] Schema: '{0}' -> '300'" -f $OldSchema)
    }

    $Summary.Persist()
    $Summary = $null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # --------------------------------------------------------
    # GENERATE MST
    # --------------------------------------------------------
    Write-Log "------------------------------------------------------------"
    Write-Log "GENERATING MST"
    Write-Log "------------------------------------------------------------"

    $OriginalDatabase = $Installer.OpenDatabase($MsiPath, 0)
    $ModifiedDatabase = $Installer.OpenDatabase($TempMsi, 0)

    if (Test-Path -LiteralPath $OutputMst) {
        Remove-Item -LiteralPath $OutputMst -Force
        Write-Log "Existing MST removed."
    }

    $ModifiedDatabase.GenerateTransform(
        $OriginalDatabase,
        $OutputMst
    )

    Write-Log "MST generated successfully." "SUCCESS"

    try {
        $ModifiedDatabase.CreateTransformSummaryInfo(
            $OriginalDatabase,
            $OutputMst,
            0,
            2
        )

        Write-Log "Transform summary information created."
    }
    catch {
        Write-Log (
            "Transform summary information could not be created: {0}" -f
            $_.Exception.Message
        ) "WARNING"
    }

    if (-not (Test-Path -LiteralPath $OutputMst -PathType Leaf)) {
        throw "MST file was not created."
    }

    # --------------------------------------------------------
    # FINAL SUMMARY
    # --------------------------------------------------------
    Write-Log "============================================================"
    Write-Log "PROCESSING SUMMARY"
    Write-Log "============================================================"
    Write-Log ("Properties Added           : {0}" -f $Added)
    Write-Log ("Properties Updated         : {0}" -f $Updated)
    Write-Log ("Properties Skipped         : {0}" -f $Skipped)
    Write-Log ("Desktop Shortcuts Removed  : {0}" -f $DesktopRemoved)
    Write-Log ("Advertised Shortcuts Found : {0}" -f $AdvertisedFound)
    Write-Log "------------------------------------------------------------"
    Write-Log ("Original MSI : {0}" -f $MsiPath)
    Write-Log ("MST          : {0}" -f $OutputMst)
    Write-Log ("Log          : {0}" -f $LogPath)
    Write-Log "Original vendor MSI was NOT modified."
    Write-Log "Temporary MSI will be deleted."
    Write-Log "============================================================"
    Write-Log "PROCESSING COMPLETED SUCCESSFULLY" "SUCCESS"
    Write-Log "============================================================"

    Write-Host "-----------------------------------------------------------"
    Write-Host "PROCESSING COMPLETED SUCCESSFULLY" -ForegroundColor Green
    Write-Host "-----------------------------------------------------------"
    Write-Host ("MST : {0}" -f $OutputMst) -ForegroundColor Green
    Write-Host ("LOG : {0}" -f $LogPath) -ForegroundColor Green
    Write-Host "-----------------------------------------------------------"
}
catch {
    $ErrorMessage = $_.Exception.Message

    if (-not (Test-Path -LiteralPath $LogPath)) {
        New-Item -ItemType File -Path $LogPath -Force | Out-Null
    }

    Write-Log "============================================================" "ERROR"
    Write-Log "PROCESSING FAILED" "ERROR"
    Write-Log ("Error: {0}" -f $ErrorMessage) "ERROR"
    Write-Log "Original vendor MSI was NOT modified." "ERROR"
    Write-Log "============================================================" "ERROR"

    Write-Host "-----------------------------------------------------------"
    Write-Host "PROCESSING FAILED" -ForegroundColor Red
    Write-Host $ErrorMessage -ForegroundColor Red
    Write-Host "-----------------------------------------------------------"

    throw
}
finally {
    $View = $null
    $ModifiedDatabase = $null
    $OriginalDatabase = $null
    $Installer = $null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if (Test-Path -LiteralPath $TempMsi) {
        Remove-Item -LiteralPath $TempMsi -Force -ErrorAction SilentlyContinue
    }
}
