Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Root = Split-Path -Parent $PSScriptRoot
$script:Engine = Join-Path $script:Root 'Scripts\Invoke-PackagingEngine.ps1'

function New-Label {
    param($Parent,[string]$Text,[int]$X,[int]$Y)
    $c = New-Object System.Windows.Forms.Label
    $c.Text = $Text
    $c.Location = New-Object System.Drawing.Point($X,$Y)
    $c.AutoSize = $true
    [void]$Parent.Controls.Add($c)
    return $c
}
function New-TextBox {
    param($Parent,[int]$X,[int]$Y,[int]$W,[int]$H=28)
    $c = New-Object System.Windows.Forms.TextBox
    $c.Location = New-Object System.Drawing.Point($X,$Y)
    $c.Size = New-Object System.Drawing.Size($W,$H)
    [void]$Parent.Controls.Add($c)
    return $c
}
function New-Combo {
    param($Parent,[int]$X,[int]$Y,[int]$W,[string[]]$Items)
    $c = New-Object System.Windows.Forms.ComboBox
    $c.Location = New-Object System.Drawing.Point($X,$Y)
    $c.Size = New-Object System.Drawing.Size($W,28)
    $c.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$c.Items.AddRange($Items)
    $c.SelectedIndex = 0
    [void]$Parent.Controls.Add($c)
    return $c
}
function Add-FolderButton {
    param($Parent,[int]$X,[int]$Y,[System.Windows.Forms.TextBox]$Target,[string]$Description)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = 'Browse...'
    $b.Location = New-Object System.Drawing.Point($X,$Y)
    $b.Size = New-Object System.Drawing.Size(90,28)
    $b.Tag = $Target
    $b.Add_Click({
        $box = [System.Windows.Forms.TextBox]$this.Tag
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = $Description
        $d.ShowNewFolderButton = $true
        if ($box.Text -and (Test-Path -LiteralPath $box.Text -PathType Container)) {
            $d.SelectedPath = $box.Text
        }
        try {
            if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $box.Text = $d.SelectedPath
                $box.Focus()
            }
        } finally { $d.Dispose() }
    })
    [void]$Parent.Controls.Add($b)
    return $b
}
function Add-FileButton {
    param($Parent,[int]$X,[int]$Y,[System.Windows.Forms.TextBox]$Target,[System.Windows.Forms.TextBox]$Source,[string]$Filter)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = 'Browse...'
    $b.Location = New-Object System.Drawing.Point($X,$Y)
    $b.Size = New-Object System.Drawing.Size(90,28)
    $b.Tag = [pscustomobject]@{Target=$Target;Source=$Source;Filter=$Filter}
    $b.Add_Click({
        $ctx = $this.Tag
        $box = [System.Windows.Forms.TextBox]$ctx.Target
        $src = [System.Windows.Forms.TextBox]$ctx.Source
        $d = New-Object System.Windows.Forms.OpenFileDialog
        $d.Filter = $ctx.Filter
        $d.Multiselect = $false
        if ($src.Text -and (Test-Path -LiteralPath $src.Text -PathType Container)) {
            $d.InitialDirectory = $src.Text
        }
        try {
            if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $selected = [IO.Path]::GetFullPath($d.FileName)
                $base = ''
                if ($src.Text -and (Test-Path -LiteralPath $src.Text -PathType Container)) {
                    $base = [IO.Path]::GetFullPath($src.Text).TrimEnd('\')
                }
                if ($base -and $selected.StartsWith($base + '\',[StringComparison]::OrdinalIgnoreCase)) {
                    $box.Text = $selected.Substring($base.Length + 1)
                } else { $box.Text = $selected }
                $box.Focus()
            }
        } finally { $d.Dispose() }
    })
    [void]$Parent.Controls.Add($b)
    return $b
}

# ---------------- Window ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Application Packaging Automation - V1'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(1280,850)
$form.MinimumSize = New-Object System.Drawing.Size(1050,700)
$form.Font = New-Object System.Drawing.Font('Segoe UI',10)
$form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)

# Header is a fixed 78px band at the top.
$header = New-Object System.Windows.Forms.Panel
$header.Name = 'Header'
$header.Location = New-Object System.Drawing.Point(0,0)
$header.Size = New-Object System.Drawing.Size($form.ClientSize.Width,78)
$header.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
                 [System.Windows.Forms.AnchorStyles]::Left -bor
                 [System.Windows.Forms.AnchorStyles]::Right
$header.BackColor = [System.Drawing.Color]::FromArgb(30,43,55)
$form.Controls.Add($header)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'Application Packaging Automation'
$title.Font = New-Object System.Drawing.Font('Segoe UI Semibold',19)
$title.ForeColor = [System.Drawing.Color]::White
$title.Location = New-Object System.Drawing.Point(24,12)
$title.AutoSize = $true
$header.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = 'MSI / EXE  |  PSADT 3.10  |  Standard MST  |  PID  |  IntuneWin'
$subtitle.ForeColor = [System.Drawing.Color]::Gainsboro
$subtitle.Location = New-Object System.Drawing.Point(27,48)
$subtitle.AutoSize = $true
$header.Controls.Add($subtitle)

$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Location = New-Object System.Drawing.Point(($form.ClientSize.Width-150),0)
$statusPanel.Size = New-Object System.Drawing.Size(150,78)
$statusPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
                      [System.Windows.Forms.AnchorStyles]::Right
$statusPanel.BackColor = [System.Drawing.Color]::FromArgb(24,35,45)
$header.Controls.Add($statusPanel)

$status = New-Object System.Windows.Forms.Label
$status.Text = 'READY'
$status.Font = New-Object System.Drawing.Font('Segoe UI Semibold',11)
$status.ForeColor = [System.Drawing.Color]::LightGreen
$status.Dock = [System.Windows.Forms.DockStyle]::Fill
$status.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$statusPanel.Controls.Add($status)

# Main TabControl: explicit position/size. No TableLayoutPanel and no Dock.
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Name = 'MainTabs'
$tabs.Location = New-Object System.Drawing.Point(0,78)
$tabs.Size = New-Object System.Drawing.Size($form.ClientSize.Width,($form.ClientSize.Height-78))
$tabs.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
               [System.Windows.Forms.AnchorStyles]::Bottom -bor
               [System.Windows.Forms.AnchorStyles]::Left -bor
               [System.Windows.Forms.AnchorStyles]::Right
$tabs.Alignment = [System.Windows.Forms.TabAlignment]::Top
$tabs.Multiline = $false
$tabs.DrawMode = [System.Windows.Forms.TabDrawMode]::Normal
$tabs.ItemSize = New-Object System.Drawing.Size(140,30)
$tabs.Padding = New-Object System.Drawing.Point(14,5)
$form.Controls.Add($tabs)
$tabs.BringToFront()
$header.BringToFront()

$form.Add_Resize({
    if($form.ClientSize.Width -gt 0 -and $form.ClientSize.Height -gt 100){
        $header.Location = New-Object System.Drawing.Point(0,0)
        $header.Size = New-Object System.Drawing.Size($form.ClientSize.Width,78)
        $tabs.Location = New-Object System.Drawing.Point(0,78)
        $tabs.Size = New-Object System.Drawing.Size($form.ClientSize.Width,($form.ClientSize.Height-78))
    }
})

# ---------------- Tabs ----------------

function New-Tab([string]$Text) {
    $p = New-Object System.Windows.Forms.TabPage
    $p.Text = $Text
    $p.BackColor = [System.Drawing.Color]::White
    [void]$tabs.TabPages.Add($p)
    return $p
}
$general = New-Tab 'Package'
$pidTab = New-Tab 'PID Document'
$install = New-Tab 'Installation'
$uninstall = New-Tab 'Uninstallation'
$advanced = New-Tab 'Templates and Output'
$runTab = New-Tab 'Run Log'
$buildTab = New-Tab 'Build and Output'
$nonAutoTab = New-Tab 'Non-Automated IntuneWin'

# ---------------- Package ----------------
New-Label $general 'Application Name' 28 25
$app = New-TextBox $general 28 48 300
New-Label $general 'Version' 350 25
$ver = New-TextBox $general 350 48 160
New-Label $general 'Vendor' 535 25
$vendor = New-TextBox $general 535 48 240
New-Label $general 'Architecture' 805 25
$arch = New-Combo $general 805 48 180 @('x64','x86','ARM64')

New-Label $general 'Installer Type' 28 94
$type = New-Combo $general 28 117 180 @('MSI','EXE')
New-Label $general 'Source Folder' 225 94
$source = New-TextBox $general 225 117 680
[void](Add-FolderButton $general 915 117 $source 'Select the vendor source folder')

New-Label $general 'Installer File (optional; relative to source folder)' 28 163
$installer = New-TextBox $general 28 186 870
[void](Add-FileButton $general 915 186 $installer $source 'Installer files (*.msi;*.exe)|*.msi;*.exe|All files (*.*)|*.*')

$mst = New-Object System.Windows.Forms.CheckBox
$mst.Text = 'Generate Standard MST (MSI only)'
$mst.Location = New-Object System.Drawing.Point(28,235)
$mst.AutoSize = $true
$mst.Checked = $true
$general.Controls.Add($mst)

$pid = New-Object System.Windows.Forms.CheckBox
$pid.Text = 'Generate PID document automatically'
$pid.Location = New-Object System.Drawing.Point(28,270)
$pid.AutoSize = $true
$pid.Checked = $true
$pid.Enabled = $false
$general.Controls.Add($pid)

$intune = New-Object System.Windows.Forms.CheckBox
$intune.Text = 'Generate .intunewin'
$intune.Location = New-Object System.Drawing.Point(28,305)
$intune.AutoSize = $true
$intune.Checked = $true
$general.Controls.Add($intune)

$serviceUI = New-Object System.Windows.Forms.CheckBox
$serviceUI.Text = 'Include ServiceUI.exe (copy from PSADT template)'
$serviceUI.Location = New-Object System.Drawing.Point(28,340)
$serviceUI.AutoSize = $true
$serviceUI.Checked = $false
$general.Controls.Add($serviceUI)

$hint = New-Object System.Windows.Forms.Label
$hint.Text = 'Select the source folder first. MSI ProductCode and PID information are read automatically from the selected MSI.'
$hint.ForeColor = [System.Drawing.Color]::DimGray
$hint.Location = New-Object System.Drawing.Point(28,375)
$hint.AutoSize = $true
$general.Controls.Add($hint)

# ---------------- PID Document ----------------
New-Label $pidTab 'PID Document Generation' 28 25
$pidInfo = New-Object System.Windows.Forms.Label
$pidInfo.Text = 'PID is generated automatically for every build using your existing PID script. MSI ProductCode and MSI properties are read directly from the selected MSI.'
$pidInfo.Location = New-Object System.Drawing.Point(28,52)
$pidInfo.AutoSize = $true
$pidInfo.ForeColor = [System.Drawing.Color]::DimGray
$pidTab.Controls.Add($pidInfo)

New-Label $pidTab 'PID Template' 28 92
$pidTemplateDisplay = New-TextBox $pidTab 28 115 850
$pidTemplateDisplay.Text = 'Auto-detected under C:\Temp\template\Package_Template.docx'
$pidTemplateDisplay.ReadOnly = $true
$pidTemplateDisplay.BackColor = [System.Drawing.Color]::WhiteSmoke

New-Label $pidTab 'PID Output' 28 162
$pidOutputDisplay = New-TextBox $pidTab 28 185 850
$pidOutputDisplay.Text = 'Automatically created: <Output>\<PackageName>\Docs\<PackageName>_PID.docx'
$pidOutputDisplay.ReadOnly = $true
$pidOutputDisplay.BackColor = [System.Drawing.Color]::WhiteSmoke

New-Label $pidTab 'MSI Product Information (automatic)' 28 235
New-Label $pidTab 'Product Name' 28 268
$msiProductName = New-TextBox $pidTab 28 291 430
$msiProductName.ReadOnly = $true
New-Label $pidTab 'ProductCode' 480 268
$msiProductCode = New-TextBox $pidTab 480 291 430
$msiProductCode.ReadOnly = $true
New-Label $pidTab 'Product Version' 28 333
$msiProductVersion = New-TextBox $pidTab 28 356 430
$msiProductVersion.ReadOnly = $true
New-Label $pidTab 'Manufacturer' 480 333
$msiManufacturer = New-TextBox $pidTab 480 356 430
$msiManufacturer.ReadOnly = $true
New-Label $pidTab 'UpgradeCode' 28 398
$msiUpgradeCode = New-TextBox $pidTab 28 421 430
$msiUpgradeCode.ReadOnly = $true

$pidStatus = New-Object System.Windows.Forms.Label
$pidStatus.Text = 'PID Status: READY - ProductCode will be read directly from the MSI.'
$pidStatus.Location = New-Object System.Drawing.Point(28,475)
$pidStatus.AutoSize = $true
$pidStatus.ForeColor = [System.Drawing.Color]::DarkGreen
$pidTab.Controls.Add($pidStatus)

# ---------------- Installation ----------------
New-Label $install 'Install Parameters' 28 25
$installParams = New-TextBox $install 28 48 1020
$installParams.Text = '/qn /norestart'
$installSwitchStatus=New-Object System.Windows.Forms.Label; $installSwitchStatus.Text='Silent Switch Finder: waiting for installer selection...'; $installSwitchStatus.Location=New-Object System.Drawing.Point(28,72); $installSwitchStatus.Size=New-Object System.Drawing.Size(1020,20); $installSwitchStatus.ForeColor=[System.Drawing.Color]::DimGray; $install.Controls.Add($installSwitchStatus)
New-Label $install 'Pre-Installation Steps (PSADT 3.10 commands)' 28 92
$preInstall = New-TextBox $install 28 118 1020 135
$preInstall.Multiline = $true
$preInstall.ScrollBars = 'Vertical'
$preInstall.AcceptsReturn = $true
New-Label $install 'Post-Installation Steps (PSADT 3.10 commands)' 28 260
$postInstall = New-TextBox $install 28 286 1020 110
$postInstall.Multiline = $true
$postInstall.ScrollBars = 'Vertical'
$postInstall.AcceptsReturn = $true
New-Label $install 'Close Apps (optional - e.g. WINWORD, EXCEL)' 28 410
$installCloseApps = New-TextBox $install 28 434 1020
$install.AutoScroll = $true

# ---------------- Uninstallation ----------------
New-Label $uninstall 'Uninstallation Method' 28 25
$uMethod = New-Combo $uninstall 28 48 360 @('Main EXE / Installer','Uninstall from different location','MSI ProductCode')
New-Label $uninstall 'Uninstall EXE / Path / ProductCode' 420 25
$uFile = New-TextBox $uninstall 420 48 520
[void](Add-FileButton $uninstall 950 48 $uFile $source 'Executable files (*.exe;*.cmd;*.bat)|*.exe;*.cmd;*.bat|All files (*.*)|*.*')

New-Label $uninstall 'Uninstall Parameters' 28 92
$uParams = New-TextBox $uninstall 28 115 480
$uParams.Text = '/qn /norestart'
$uninstallSwitchStatus=New-Object System.Windows.Forms.Label; $uninstallSwitchStatus.Text='Silent Switch Finder: waiting for uninstall selection...'; $uninstallSwitchStatus.Location=New-Object System.Drawing.Point(28,138); $uninstallSwitchStatus.Size=New-Object System.Drawing.Size(1010,20); $uninstallSwitchStatus.ForeColor=[System.Drawing.Color]::DimGray; $uninstall.Controls.Add($uninstallSwitchStatus)
New-Label $uninstall 'Working Directory (optional)' 530 92
$uWd = New-TextBox $uninstall 530 115 508
New-Label $uninstall 'Pre-Uninstallation Steps (PSADT 3.10 commands)' 28 190
$preUninstall = New-TextBox $uninstall 28 216 1010 110
$preUninstall.Multiline = $true
$preUninstall.ScrollBars = 'Vertical'
$preUninstall.AcceptsReturn = $true
New-Label $uninstall 'Post-Uninstallation Steps (PSADT 3.10 commands)' 28 340
$postUninstall = New-TextBox $uninstall 28 366 1010 110
$postUninstall.Multiline = $true
$postUninstall.ScrollBars = 'Vertical'
$postUninstall.AcceptsReturn = $true
New-Label $uninstall 'Close Apps (optional - e.g. WINWORD, EXCEL)' 28 490
$uninstallCloseApps = New-TextBox $uninstall 28 514 1010
$uninstall.AutoScroll = $true

# ---------------- Templates & Output ----------------
New-Label $advanced 'Template Root' 28 25
$templateBox = New-TextBox $advanced 28 48 850
$templateBox.Text = 'C:\Temp\template'
[void](Add-FolderButton $advanced 890 48 $templateBox 'Select template root')

New-Label $advanced 'Output Root' 28 95
$outputBox = New-TextBox $advanced 28 118 850
$outputBox.Text = 'C:\Packaging\Output'
[void](Add-FolderButton $advanced 890 118 $outputBox 'Select output root')

New-Label $advanced 'Automation Logs' 28 165
$logBox = New-TextBox $advanced 28 188 850
$logBox.Text = 'C:\Temp\AutomationLogs'
$logBox.ReadOnly = $true
$logBox.BackColor = [System.Drawing.Color]::WhiteSmoke

$validate = New-Object System.Windows.Forms.Button
$validate.Text = 'Validate Templates'
$validate.Location = New-Object System.Drawing.Point(28,240)
$validate.Size = New-Object System.Drawing.Size(180,38)
$advanced.Controls.Add($validate)

$validation = New-Object System.Windows.Forms.TextBox
$validation.Location = New-Object System.Drawing.Point(28,295)
$validation.Size = New-Object System.Drawing.Size(1010,220)
$validation.Multiline = $true
$validation.ReadOnly = $true
$validation.ScrollBars = 'Vertical'
$validation.Font = New-Object System.Drawing.Font('Consolas',9)
$advanced.Controls.Add($validation)

$validate.Add_Click({
    $items = @()
    $rootPath = $templateBox.Text.Trim()

    if (-not (Test-Path -LiteralPath $rootPath -PathType Container)) {
        $items += ('Template root NOT found: ' + $rootPath)
        $validation.Lines = $items
        return
    }
    $items += ('Template root found: ' + $rootPath)

    $packageTemplate = Join-Path $rootPath 'Package_Template.docx'
    if (Test-Path -LiteralPath $packageTemplate -PathType Leaf) {
        $items += 'Package_Template.docx found'
    } else {
        $items += ('Package_Template.docx NOT found: ' + $packageTemplate)
    }

    $intuneUtil = Join-Path $rootPath 'IntuneWinAppUtil.exe'
    if (Test-Path -LiteralPath $intuneUtil -PathType Leaf) {
        $items += 'IntuneWinAppUtil.exe found'
    } else {
        $items += ('IntuneWinAppUtil.exe NOT found: ' + $intuneUtil)
    }

    $psadt = $null
    $dirs = Get-ChildItem -LiteralPath $rootPath -Directory -ErrorAction SilentlyContinue
    foreach ($d in $dirs) {
        if ((Test-Path -LiteralPath (Join-Path $d.FullName 'Deploy-Application.ps1')) -and
            (Test-Path -LiteralPath (Join-Path $d.FullName 'AppDeployToolkit'))) {
            $psadt = $d
            break
        }
    }
    if (-not $psadt -and
        (Test-Path -LiteralPath (Join-Path $rootPath 'Deploy-Application.ps1')) -and
        (Test-Path -LiteralPath (Join-Path $rootPath 'AppDeployToolkit'))) {
        $psadt = Get-Item -LiteralPath $rootPath
    }

    if ($psadt) {
        $items += ('PSADT 3.10 template found: ' + $psadt.FullName)
    } else {
        $items += 'PSADT 3.10 template NOT found.'
    }

    $validation.Lines = $items
})

# ---------------- Run Log ----------------
$runBox = New-Object System.Windows.Forms.TextBox
$runBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$runBox.Multiline = $true
$runBox.ReadOnly = $true
$runBox.ScrollBars = 'Vertical'
$runBox.Font = New-Object System.Drawing.Font('Consolas',9)
$runTab.Controls.Add($runBox)

# ---------------- BUILD & OUTPUT TAB ----------------

# ---------------- Build & Output ----------------
$buildIntro = New-Object System.Windows.Forms.Label
$buildIntro.Text = 'Build controls and final package status  |  V1'
$buildIntro.Font = New-Object System.Drawing.Font('Segoe UI Semibold',14)
$buildIntro.Location = New-Object System.Drawing.Point(28,22)
$buildIntro.AutoSize = $true
$buildTab.Controls.Add($buildIntro)

$buildSummary = New-Object System.Windows.Forms.TextBox
$buildSummary.Location = New-Object System.Drawing.Point(28,62)
$buildSummary.Size = New-Object System.Drawing.Size(720,300)
$buildSummary.Multiline = $true
$buildSummary.ReadOnly = $true
$buildSummary.ScrollBars = 'Vertical'
$buildSummary.Font = New-Object System.Drawing.Font('Consolas',10)
$buildTab.Controls.Add($buildSummary)

$script:PackagePrepared = $false
$script:ScriptConfirmed = $false
$script:PreparedPackageRoot = ''

$buildStatus = New-Object System.Windows.Forms.Label
$buildStatus.Text = 'READY - Workflow: Prepare PSADT -> Populate Deploy-Application.ps1 -> PID -> Check/Edit -> Validate -> IntuneWin.'
$buildStatus.Location = New-Object System.Drawing.Point(28,380)
$buildStatus.Size = New-Object System.Drawing.Size(1000,42)
$buildStatus.ForeColor = [System.Drawing.Color]::DimGray
$buildTab.Controls.Add($buildStatus)

$buildButtonsHeader = New-Object System.Windows.Forms.Label
$buildButtonsHeader.Text = 'BUILD CONTROLS'
$buildButtonsHeader.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$buildButtonsHeader.Location = New-Object System.Drawing.Point(790,38)
$buildButtonsHeader.AutoSize = $true
$buildButtonsHeader.ForeColor = [System.Drawing.Color]::DimGray
$buildTab.Controls.Add($buildButtonsHeader)

$prepareButton = New-Object System.Windows.Forms.Button
$prepareButton.Text = 'PREPARE PACKAGE'
$prepareButton.Font = New-Object System.Drawing.Font('Segoe UI Semibold',11)
$prepareButton.Location = New-Object System.Drawing.Point(790,62)
$prepareButton.Size = New-Object System.Drawing.Size(220,42)
$prepareButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$prepareButton.ForeColor = [System.Drawing.Color]::White
$prepareButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$prepareButton.FlatAppearance.BorderSize = 0
$buildTab.Controls.Add($prepareButton)

$openDeployButton = New-Object System.Windows.Forms.Button
$openDeployButton.Text = 'REVIEW / EDIT DEPLOY-APPLICATION.PS1'
$openDeployButton.Location = New-Object System.Drawing.Point(790,110)
$openDeployButton.Size = New-Object System.Drawing.Size(220,42)
$openDeployButton.Enabled = $false
$buildTab.Controls.Add($openDeployButton)

$confirmBuildButton = New-Object System.Windows.Forms.Button
$confirmBuildButton.Text = 'CONFIRM AND BUILD INTUNEWIN'
$confirmBuildButton.Font = New-Object System.Drawing.Font('Segoe UI Semibold',11)
$confirmBuildButton.Location = New-Object System.Drawing.Point(790,158)
$confirmBuildButton.Size = New-Object System.Drawing.Size(220,42)
$confirmBuildButton.BackColor = [System.Drawing.Color]::FromArgb(40,167,69)
$confirmBuildButton.ForeColor = [System.Drawing.Color]::White
$confirmBuildButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$confirmBuildButton.FlatAppearance.BorderSize = 0
$confirmBuildButton.Enabled = $false
$buildTab.Controls.Add($confirmBuildButton)

$validateBuildButton = New-Object System.Windows.Forms.Button
$validateBuildButton.Text = 'VALIDATE TEMPLATES'
$validateBuildButton.Location = New-Object System.Drawing.Point(790,206)
$validateBuildButton.Size = New-Object System.Drawing.Size(220,38)
$buildTab.Controls.Add($validateBuildButton)

$openOutputButton = New-Object System.Windows.Forms.Button
$openOutputButton.Text = 'OPEN OUTPUT'
$openOutputButton.Location = New-Object System.Drawing.Point(790,252)
$openOutputButton.Size = New-Object System.Drawing.Size(220,38)
$buildTab.Controls.Add($openOutputButton)

$openLogsButton = New-Object System.Windows.Forms.Button
$openLogsButton.Text = 'OPEN LOGS'
$openLogsButton.Location = New-Object System.Drawing.Point(790,298)
$openLogsButton.Size = New-Object System.Drawing.Size(220,38)
$buildTab.Controls.Add($openLogsButton)

$openPackageButton = New-Object System.Windows.Forms.Button
$openPackageButton.Text = 'OPEN PACKAGE FOLDER'
$openPackageButton.Location = New-Object System.Drawing.Point(790,344)
$openPackageButton.Size = New-Object System.Drawing.Size(220,38)
$buildTab.Controls.Add($openPackageButton)

function Update-BuildSummary {
    $pkgName = if($arch.SelectedItem -eq 'x86'){ '{0}_{1}_{2}_P1' -f $vendor.Text.Trim(),$app.Text.Trim(),$ver.Text.Trim() } else { '{0}_{1}_{2}_P1_{3}' -f $vendor.Text.Trim(),$app.Text.Trim(),$ver.Text.Trim(),$arch.SelectedItem }
    $pkgRoot = if($outputBox.Text.Trim()) { Join-Path $outputBox.Text.Trim() $pkgName } else { $pkgName }
    $buildSummary.Text = @(
        'PACKAGE BUILD SUMMARY'
        '====================='
        ('Application : ' + $app.Text.Trim())
        ('Version     : ' + $ver.Text.Trim())
        ('Vendor      : ' + $vendor.Text.Trim())
        ('Architecture: ' + $arch.SelectedItem)
        ('Type        : ' + $type.SelectedItem)
        ('Source      : ' + $source.Text.Trim())
        ('Installer   : ' + $installer.Text.Trim())
        ('Output      : ' + $pkgRoot)
        ('Package     : ' + (Join-Path $pkgRoot 'Package'))
        ('Docs        : ' + (Join-Path $pkgRoot 'Docs'))
        ('Intune      : ' + (Join-Path $pkgRoot 'Intune'))
        ('Automation Logs: C:\Temp\AutomationLogs')
        ('Logs        : C:\Temp\AutomationLogs')
        ''
        ('MST         : ' + $(if($mst.Checked -and $type.SelectedItem -eq 'MSI'){'YES'}else{'NO'}))
        ('PID         : YES (mandatory automatic phase)')
        ('IntuneWin   : BUILD AFTER SCRIPT REVIEW')
        ('PID Output  : ' + (Join-Path (Join-Path $pkgRoot 'Docs') ($pkgName + '_PID.docx')))
    ) -join [Environment]::NewLine
}

$validateBuildButton.Add_Click({
    $tabs.SelectedTab = $advanced
    $validate.PerformClick()
})

$openOutputButton.Add_Click({
    if(-not (Test-Path -LiteralPath $outputBox.Text -PathType Container)){
        New-Item -ItemType Directory -Force -Path $outputBox.Text | Out-Null
    }
    Start-Process explorer.exe -ArgumentList ('"{0}"' -f $outputBox.Text)
})

$openLogsButton.Add_Click({
    $logPath = 'C:\Temp\AutomationLogs'
    if(-not (Test-Path -LiteralPath $logPath -PathType Container)){
        New-Item -ItemType Directory -Force -Path $logPath | Out-Null
    }
    Start-Process explorer.exe -ArgumentList ('"{0}"' -f $logPath)
})

$openPackageButton.Add_Click({
    $pkgName = if($arch.SelectedItem -eq 'x86'){ '{0}_{1}_{2}_P1' -f $vendor.Text.Trim(),$app.Text.Trim(),$ver.Text.Trim() } else { '{0}_{1}_{2}_P1_{3}' -f $vendor.Text.Trim(),$app.Text.Trim(),$ver.Text.Trim(),$arch.SelectedItem }
    $pkgRoot = Join-Path $outputBox.Text.Trim() $pkgName
    if(Test-Path -LiteralPath $pkgRoot -PathType Container){
        Start-Process explorer.exe -ArgumentList ('"{0}"' -f $pkgRoot)
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            ('Package has not been built yet.' + [Environment]::NewLine + $pkgRoot),
            'Packaging Automation'
        )
    }
})

function Read-MsiPropertiesToUI {
    param([string]$MsiPath)
    if([string]::IsNullOrWhiteSpace($MsiPath) -or -not (Test-Path -LiteralPath $MsiPath -PathType Leaf)){ return }
    if([IO.Path]::GetExtension($MsiPath) -ine '.msi'){ return }
    $installerCom=$null; $database=$null
    try {
        $installerCom=New-Object -ComObject WindowsInstaller.Installer
        $database=$installerCom.OpenDatabase($MsiPath,0)
        $values=@{}
        foreach($prop in @('ProductName','ProductVersion','Manufacturer','ProductCode','UpgradeCode')){
            $view=$null
            try {
                $view=$database.OpenView("SELECT `Value` FROM Property WHERE `Property`='$prop'")
                $view.Execute(); $record=$view.Fetch()
                if($record){ $values[$prop]=[string]$record.StringData(1) }
            } finally { try { if($view){$view.Close()} } catch {} }
        }
        $msiProductName.Text = [string]$values['ProductName']
        $msiProductVersion.Text = [string]$values['ProductVersion']
        $msiManufacturer.Text = [string]$values['Manufacturer']
        $msiProductCode.Text = [string]$values['ProductCode']
        $msiUpgradeCode.Text = [string]$values['UpgradeCode']
        if($values['ProductCode']){
            $pidStatus.Text='PID Status: READY - MSI ProductCode detected automatically.'
            $pidStatus.ForeColor=[System.Drawing.Color]::DarkGreen
        } else {
            $pidStatus.Text='PID Status: WARNING - ProductCode could not be read from this MSI.'
            $pidStatus.ForeColor=[System.Drawing.Color]::DarkOrange
        }
    } catch {
        $pidStatus.Text='PID Status: WARNING - Could not read MSI properties: ' + $_.Exception.Message
        $pidStatus.ForeColor=[System.Drawing.Color]::DarkOrange
    } finally {
        try { if($database){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($database)} } catch {}
        try { if($installerCom){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($installerCom)} } catch {}
    }
}

# ---------------- Non-Automated IntuneWin ----------------
New-Label $nonAutoTab 'Non-Automated IntuneWin Creator' 28 25
$nonAutoInfo = New-Object System.Windows.Forms.Label
$nonAutoInfo.Text = 'Create an .intunewin file for an existing / non-automated package without running the packaging automation workflow.'
$nonAutoInfo.Location = New-Object System.Drawing.Point(28,52)
$nonAutoInfo.AutoSize = $true
$nonAutoInfo.ForeColor = [System.Drawing.Color]::DimGray
$nonAutoTab.Controls.Add($nonAutoInfo)

New-Label $nonAutoTab 'Source Folder' 28 95
$nonAutoSource = New-TextBox $nonAutoTab 28 118 850
[void](Add-FolderButton $nonAutoTab 890 118 $nonAutoSource 'Select the existing package/source folder')

New-Label $nonAutoTab 'Setup File (relative to Source Folder)' 28 165
$nonAutoSetup = New-TextBox $nonAutoTab 28 188 850
[void](Add-FileButton $nonAutoTab 890 188 $nonAutoSetup $nonAutoSource 'Setup files (*.exe;*.msi;*.cmd;*.bat)|*.exe;*.msi;*.cmd;*.bat|All files (*.*)|*.*')

New-Label $nonAutoTab 'Output Folder' 28 235
$nonAutoOutput = New-TextBox $nonAutoTab 28 258 850
$nonAutoOutput.Text = 'C:\Packaging\Output\IntuneWin'
[void](Add-FolderButton $nonAutoTab 890 258 $nonAutoOutput 'Select the IntuneWin output folder')

$nonAutoStatus = New-Object System.Windows.Forms.Label
$nonAutoStatus.Text = 'READY'
$nonAutoStatus.Location = New-Object System.Drawing.Point(28,315)
$nonAutoStatus.Size = New-Object System.Drawing.Size(1000,35)
$nonAutoStatus.ForeColor = [System.Drawing.Color]::DimGray
$nonAutoTab.Controls.Add($nonAutoStatus)

$nonAutoLog = New-Object System.Windows.Forms.TextBox
$nonAutoLog.Location = New-Object System.Drawing.Point(28,350)
$nonAutoLog.Size = New-Object System.Drawing.Size(1010,145)
$nonAutoLog.Multiline = $true
$nonAutoLog.ReadOnly = $true
$nonAutoLog.ScrollBars = 'Vertical'
$nonAutoLog.Font = New-Object System.Drawing.Font('Consolas',9)
$nonAutoTab.Controls.Add($nonAutoLog)

$nonAutoBuild = New-Object System.Windows.Forms.Button
$nonAutoBuild.Text = 'CREATE INTUNEWIN'
$nonAutoBuild.Font = New-Object System.Drawing.Font('Segoe UI Semibold',11)
$nonAutoBuild.Location = New-Object System.Drawing.Point(28,515)
$nonAutoBuild.Size = New-Object System.Drawing.Size(240,44)
$nonAutoBuild.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$nonAutoBuild.ForeColor = [System.Drawing.Color]::White
$nonAutoBuild.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$nonAutoBuild.FlatAppearance.BorderSize = 0
$nonAutoTab.Controls.Add($nonAutoBuild)

$nonAutoOpen = New-Object System.Windows.Forms.Button
$nonAutoOpen.Text = 'OPEN OUTPUT'
$nonAutoOpen.Location = New-Object System.Drawing.Point(285,515)
$nonAutoOpen.Size = New-Object System.Drawing.Size(180,44)
$nonAutoTab.Controls.Add($nonAutoOpen)

$nonAutoOpen.Add_Click({
    if(-not (Test-Path -LiteralPath $nonAutoOutput.Text.Trim() -PathType Container)){
        New-Item -ItemType Directory -Force -Path $nonAutoOutput.Text.Trim() | Out-Null
    }
    Start-Process explorer.exe -ArgumentList ('"{0}"' -f $nonAutoOutput.Text.Trim())
})

$nonAutoBuild.Add_Click({
    try {
        $srcPath = $nonAutoSource.Text.Trim()
        $setup = $nonAutoSetup.Text.Trim()
        $outPath = $nonAutoOutput.Text.Trim()
        if([string]::IsNullOrWhiteSpace($srcPath) -or -not (Test-Path -LiteralPath $srcPath -PathType Container)){
            throw 'Select a valid Source Folder.'
        }
        if([string]::IsNullOrWhiteSpace($setup)){ throw 'Select the Setup File.' }
        $setupFull = if([IO.Path]::IsPathRooted($setup)){ $setup } else { Join-Path $srcPath $setup }
        if(-not (Test-Path -LiteralPath $setupFull -PathType Leaf)){ throw "Setup file not found: $setupFull" }
        $setupArgument = $setup
        if([IO.Path]::IsPathRooted($setupArgument)) {
            $sourceRoot = $srcPath.TrimEnd('\')
            if($setupArgument.StartsWith($sourceRoot + '\',[StringComparison]::OrdinalIgnoreCase)) {
                $setupArgument = $setupArgument.Substring($sourceRoot.Length + 1)
            } else {
                throw 'Setup File must be inside the selected Source Folder.'
            }
        }
        if([string]::IsNullOrWhiteSpace($outPath)){ throw 'Select an Output Folder.' }
        New-Item -ItemType Directory -Force -Path $outPath | Out-Null

        $util = Get-ChildItem -LiteralPath $templateBox.Text.Trim() -Filter 'IntuneWinAppUtil.exe' -File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if(-not $util){ throw "IntuneWinAppUtil.exe was not found under Template Root: $($templateBox.Text.Trim())" }

        $nonAutoStatus.Text = 'BUILDING - Creating IntuneWin...'
        $nonAutoStatus.ForeColor = [System.Drawing.Color]::DarkOrange
        $nonAutoLog.Clear()
        $nonAutoLog.AppendText(('Utility: ' + $util.FullName) + [Environment]::NewLine)
        $nonAutoLog.AppendText(('Source : ' + $srcPath) + [Environment]::NewLine)
        $nonAutoLog.AppendText(('Setup  : ' + $setupFull) + [Environment]::NewLine)
        $nonAutoLog.AppendText(('Output : ' + $outPath) + [Environment]::NewLine)

        $p = Start-Process -FilePath $util.FullName -ArgumentList @('-c',$srcPath,'-s',$setupArgument,'-o',$outPath,'-q') -Wait -PassThru -NoNewWindow
        if($p.ExitCode -ne 0){ throw "IntuneWinAppUtil.exe failed with exit code $($p.ExitCode)." }
        $created = Get-ChildItem -LiteralPath $outPath -Filter '*.intunewin' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if(-not $created){ throw 'IntuneWinAppUtil completed but no .intunewin file was found in the output folder.' }

        $nonAutoLog.AppendText(('SUCCESS: ' + $created.FullName) + [Environment]::NewLine)
        $nonAutoStatus.Text = 'COMPLETED - IntuneWin created successfully.'
        $nonAutoStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        [System.Windows.Forms.MessageBox]::Show(('IntuneWin created successfully.' + [Environment]::NewLine + $created.FullName),'Non-Automated IntuneWin',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        $nonAutoStatus.Text = 'FAILED - Review the log below.'
        $nonAutoStatus.ForeColor = [System.Drawing.Color]::Tomato
        $nonAutoLog.AppendText(('ERROR: ' + $_.Exception.Message) + [Environment]::NewLine)
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Non-Automated IntuneWin',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# ---------------- Silent Switch Finder ----------------
$script:SilentSwitchFinderModule=Join-Path $PSScriptRoot '..\Tools\SilentSwitchFinder\SilentSwitchFinder.psd1'
function Find-AndPopulateSilentSwitch {
    param([Parameter(Mandatory=$true)][string]$InstallerPath,[Parameter(Mandatory=$true)][System.Windows.Forms.TextBox]$ParameterBox,[Parameter(Mandatory=$true)][System.Windows.Forms.Label]$StatusLabel,[Parameter(Mandatory=$true)][string]$ContextName)
    if(-not(Test-Path -LiteralPath $InstallerPath -PathType Leaf)){return}
    try{
        $StatusLabel.Text="Silent Switch Finder: analyzing $ContextName...";$StatusLabel.ForeColor=[System.Drawing.Color]::DarkOrange;[System.Windows.Forms.Application]::DoEvents()
        Import-Module $script:SilentSwitchFinderModule -Force -ErrorAction Stop
        $result=Find-SilentSwitch -Path $InstallerPath -ErrorAction Stop
        $silent=[string]$result.Silent
        if([IO.Path]::GetExtension($InstallerPath) -ieq '.msi'){
            if($silent -match '(?i)^msiexec\s+/[ix]\s+.*?\s+(?<params>/.*)$'){$silent=$Matches['params']}elseif($silent -match '(?i)^msiexec\s+/[ix]\s+'){$silent='/qn /norestart'}
        }
        if(-not[string]::IsNullOrWhiteSpace($silent)){$ParameterBox.Text=$silent}
        $StatusLabel.Text="Silent Switch Finder: $($result.Installer) | $silent | Confidence $($result.Confidence)% | $($result.Source)";$StatusLabel.ForeColor=[System.Drawing.Color]::DarkGreen
        if($runBox){$runBox.AppendText("[SilentSwitchFinder] $ContextName : $silent | Confidence $($result.Confidence)% | $($result.Source)"+[Environment]::NewLine);$runBox.SelectionStart=$runBox.TextLength;$runBox.ScrollToCaret()}
    }catch{
        $StatusLabel.Text='Silent Switch Finder: detection failed - enter parameters manually.';$StatusLabel.ForeColor=[System.Drawing.Color]::Tomato
        if($runBox){$runBox.AppendText("[SilentSwitchFinder][ERROR] $ContextName : $($_.Exception.Message)"+[Environment]::NewLine);$runBox.SelectionStart=$runBox.TextLength;$runBox.ScrollToCaret()}
    }
}

# ---------------- Events ----------------
$type.Add_SelectedValueChanged({
    if ($type.SelectedItem -eq 'MSI') {
        $mst.Enabled = $true
        $mst.Checked = $true
        if([string]::IsNullOrWhiteSpace($installParams.Text)){ $installParams.Text = '/qn /norestart' }
        if([string]::IsNullOrWhiteSpace($uParams.Text)){ $uParams.Text = '/qn /norestart' }
        $uMethod.SelectedItem = 'MSI ProductCode'
        if($source.Text -and (Test-Path -LiteralPath $source.Text -PathType Container) -and $installer.Text){
            $candidate = Join-Path $source.Text $installer.Text
            if(Test-Path -LiteralPath $candidate -PathType Leaf){ Read-MsiPropertiesToUI -MsiPath $candidate }
        }
    } else {
        $mst.Checked = $false
        $mst.Enabled = $false
        $installParams.Clear()
        $uParams.Clear()
        $uMethod.SelectedItem = 'Main EXE / Installer'
        $msiProductName.Clear(); $msiProductVersion.Clear(); $msiManufacturer.Clear(); $msiProductCode.Clear(); $msiUpgradeCode.Clear()
        $pidStatus.Text='PID Status: READY - EXE package.'
    }
})

$installer.Add_TextChanged({
    if($source.Text -and (Test-Path -LiteralPath $source.Text -PathType Container) -and $installer.Text){
        $candidate=Join-Path $source.Text $installer.Text.Trim()
        if(Test-Path -LiteralPath $candidate -PathType Leaf){
            Find-AndPopulateSilentSwitch -InstallerPath $candidate -ParameterBox $installParams -StatusLabel $installSwitchStatus -ContextName 'Installation'
            if($type.SelectedItem -eq 'MSI'){ Read-MsiPropertiesToUI -MsiPath $candidate }
        }
    }
})

$source.Add_TextChanged({
    if($source.Text -and (Test-Path -LiteralPath $source.Text -PathType Container) -and $installer.Text){
        $candidate=Join-Path $source.Text $installer.Text.Trim()
        if(Test-Path -LiteralPath $candidate -PathType Leaf){
            Find-AndPopulateSilentSwitch -InstallerPath $candidate -ParameterBox $installParams -StatusLabel $installSwitchStatus -ContextName 'Installation'
            if($type.SelectedItem -eq 'MSI'){ Read-MsiPropertiesToUI -MsiPath $candidate }
        }
    }
})



$uFile.Add_TextChanged({
    if($uFile.Text -and (Test-Path -LiteralPath $uFile.Text.Trim() -PathType Leaf)){
        Find-AndPopulateSilentSwitch -InstallerPath $uFile.Text.Trim() -ParameterBox $uParams -StatusLabel $uninstallSwitchStatus -ContextName 'Uninstallation'
    }
})

function Get-BuildEngineParameters {
    # Use a hashtable for PowerShell parameter splatting.
    # This is important because an argument such as "Beyond Compare.exe"
    # or a command containing spaces must remain ONE parameter value.
    $engineParameters = @{
        ApplicationName = $app.Text.Trim()
        Version         = $ver.Text.Trim()
        Vendor          = $vendor.Text.Trim()
        InstallerType   = [string]$type.SelectedItem
        Architecture    = [string]$arch.SelectedItem
        SourcePath      = $source.Text.Trim()
        InstallFile     = $installer.Text.Trim()
        UninstallMethod = [string]$uMethod.SelectedItem
        TemplateRoot    = $templateBox.Text.Trim()
        OutputRoot      = $outputBox.Text.Trim()
        GeneratePID     = $true
    }

    if(-not [string]::IsNullOrWhiteSpace($installParams.Text.Trim())){
        $engineParameters.InstallParameters = $installParams.Text.Trim()
    }
    if(-not [string]::IsNullOrWhiteSpace($uFile.Text.Trim()) -and $type.SelectedItem -ne 'MSI'){
        $engineParameters.UninstallFile = $uFile.Text.Trim()
    }
    if(-not [string]::IsNullOrWhiteSpace($uParams.Text.Trim())){
        $engineParameters.UninstallParameters = $uParams.Text.Trim()
    }
    if(-not [string]::IsNullOrWhiteSpace($preInstall.Text.Trim())){
        $engineParameters.PreInstallSteps = $preInstall.Text
    }
    if(-not [string]::IsNullOrWhiteSpace($postInstall.Text.Trim())){
        $engineParameters.PostInstallSteps = $postInstall.Text
    }
    if(-not [string]::IsNullOrWhiteSpace($preUninstall.Text.Trim())){
        $engineParameters.PreUninstallSteps = $preUninstall.Text
    }
    if(-not [string]::IsNullOrWhiteSpace($postUninstall.Text.Trim())){
        $engineParameters.PostUninstallSteps = $postUninstall.Text
    }
    if(-not [string]::IsNullOrWhiteSpace($uWd.Text.Trim())){
        $engineParameters.UninstallWorkingDirectory = $uWd.Text.Trim()
    }
    if(-not [string]::IsNullOrWhiteSpace($installCloseApps.Text.Trim())){
        $engineParameters.InstallCloseApps = $installCloseApps.Text.Trim()
    }
    if(-not [string]::IsNullOrWhiteSpace($uninstallCloseApps.Text.Trim())){
        $engineParameters.UninstallCloseApps = $uninstallCloseApps.Text.Trim()
    }

    if($mst.Checked -and $type.SelectedItem -eq 'MSI'){
        $engineParameters.GenerateMST = $true
    }
    if($serviceUI.Checked){
        $engineParameters.IncludeServiceUI = $true
    }

    return $engineParameters
}

function Invoke-PackagingEngineForWorkflow {
    param([hashtable]$EngineParameters)

    # Invoke the engine directly in this PowerShell process so every named
    # parameter remains a single value, including paths containing spaces.
    # IMPORTANT: do not capture the entire engine output into an array first.
    # That makes the Run Log appear empty until the engine has completely
    # finished. Stream each output record directly into the UI instead.
    # Also do NOT use $LASTEXITCODE here; it belongs to the last native process
    # and may contain a stale value from Word/MSI/IntuneWinAppUtil/PID work.
    $engineInvocationSucceeded = $true
    try {
        & $script:Engine @EngineParameters *>&1 | ForEach-Object {
            $engineItem = $_
            $line = ($engineItem | Out-String).TrimEnd()
            if($line){
                $runBox.AppendText($line + [Environment]::NewLine)
                $runBox.SelectionStart=$runBox.TextLength
                $runBox.ScrollToCaret()

                if($line -match '(?i)PID DOCUMENT GENERATION - STARTING'){
                    $buildStatus.Text='PREPARING - Generating PID document...'
                } elseif($line -match '(?i)PID DOCUMENT GENERATION - COMPLETED'){
                    $buildStatus.Text='PREPARING - PID document completed. Deploy-Application.ps1 is ready for review.'
                } elseif($line -match '(?i)PACKAGE VALIDATION - COMPLETED'){
                    $buildStatus.Text='BUILDING - Package validation completed.'
                } elseif($line -match '(?i)INTUNEWIN GENERATION - STARTING'){
                    $buildStatus.Text='BUILDING - Generating IntuneWin package...'
                } elseif($line -match '(?i)INTUNEWIN GENERATION - COMPLETED'){
                    $buildStatus.Text='BUILDING - IntuneWin generated successfully.'
                }
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
        $engineInvocationSucceeded = $?
    } catch {
        $engineInvocationSucceeded = $false
        throw
    }

    if(-not $engineInvocationSucceeded){
        throw 'Packaging engine failed. Review Run Log for the engine error details.'
    }
}

$prepareButton.Add_Click({
    try {
        $appName=$app.Text.Trim(); $versionText=$ver.Text.Trim(); $vendorText=$vendor.Text.Trim()
        $sourcePath=$source.Text.Trim(); $templatePath=$templateBox.Text.Trim()

        if([string]::IsNullOrWhiteSpace($appName) -or [string]::IsNullOrWhiteSpace($versionText) -or
           [string]::IsNullOrWhiteSpace($vendorText) -or [string]::IsNullOrWhiteSpace($sourcePath)){
            throw 'Application Name, Version, Vendor and Source Folder are required.'
        }
        if(-not(Test-Path -LiteralPath $sourcePath -PathType Container)){throw "Source folder does not exist:`n$sourcePath"}
        if(-not(Test-Path -LiteralPath $templatePath -PathType Container)){throw "Template folder does not exist:`n$templatePath"}

        if([string]::IsNullOrWhiteSpace($installer.Text.Trim())){
            $ext=if($type.SelectedItem -eq 'MSI'){'.msi'}else{'.exe'}
            $candidate=Get-ChildItem -LiteralPath $sourcePath -File -ErrorAction SilentlyContinue |
                Where-Object {$_.Extension -ieq $ext}|Select-Object -First 1
            if($candidate){$installer.Text=$candidate.Name}else{throw "No $ext installer was found in the selected source folder."}
        }

        if($uMethod.SelectedItem -eq 'Uninstall from different location' -and [string]::IsNullOrWhiteSpace($uFile.Text.Trim())){
            throw 'Provide the uninstall executable/path for the different-location uninstall option.'
        }

        if($type.SelectedItem -eq 'MSI'){
            $resolvedMsi=Join-Path $sourcePath $installer.Text.Trim()
            if(Test-Path -LiteralPath $resolvedMsi -PathType Leaf){Read-MsiPropertiesToUI -MsiPath $resolvedMsi}
            if([string]::IsNullOrWhiteSpace($msiProductCode.Text.Trim())){throw 'Unable to read ProductCode from the selected MSI.'}
            $uFile.Text=$msiProductCode.Text.Trim()
        }

        $script:PackagePrepared=$false
        $script:ScriptConfirmed=$false
        $buildStatus.Text='PREPARING - Creating PSADT package and populating Deploy-Application.ps1...'
        $status.Text='PREPARING'; $status.ForeColor=[System.Drawing.Color]::Gold
        $prepareButton.Enabled=$false; $openDeployButton.Enabled=$false; $confirmBuildButton.Enabled=$false
        $tabs.SelectedTab=$runTab; $runBox.Clear()

        $engineParameters=Get-BuildEngineParameters
        $engineParameters.PrepareOnly = $true
        Invoke-PackagingEngineForWorkflow -EngineParameters $engineParameters

        $pkgName=if($arch.SelectedItem -eq 'x86'){ '{0}_{1}_{2}_P1' -f $vendor.Text.Trim(),$app.Text.Trim(),$ver.Text.Trim() }else{ '{0}_{1}_{2}_P1_{3}' -f $vendor.Text.Trim(),$app.Text.Trim(),$ver.Text.Trim(),$arch.SelectedItem }
        $script:PreparedPackageRoot=Join-Path $outputBox.Text.Trim() $pkgName
        $deployPath=Join-Path $script:PreparedPackageRoot 'Package\Deploy-Application.ps1'
        if(-not(Test-Path -LiteralPath $deployPath -PathType Leaf)){throw "Prepared Deploy-Application.ps1 was not found: $deployPath"}

        $script:PackagePrepared=$true
        $openDeployButton.Enabled=$true
        $confirmBuildButton.Enabled=$true
        $buildStatus.Text='READY FOR REVIEW - Package and PID document are ready. Check/edit Deploy-Application.ps1, then confirm and build IntuneWin.'
        $tabs.SelectedTab=$buildTab
        Update-BuildSummary
        [System.Windows.Forms.MessageBox]::Show(
            'The package and PID document are ready for review.' + [Environment]::NewLine + [Environment]::NewLine +
            'Please check and edit Deploy-Application.ps1 as required, save your changes, then click CONFIRM & BUILD INTUNEWIN.',
            'Package Ready for Review',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } catch {
        $status.Text='FAILED'; $status.ForeColor=[System.Drawing.Color]::Tomato
        $buildStatus.Text='PREPARE FAILED - Review Run Log.'
        $tabs.SelectedTab=$runTab
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Prepare Package Failed',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $prepareButton.Enabled=$true
    }
})

$openDeployButton.Add_Click({
    try {
        if(-not $script:PackagePrepared){throw 'Prepare the package first.'}
        $deployPath=Join-Path $script:PreparedPackageRoot 'Package\Deploy-Application.ps1'
        if(-not(Test-Path -LiteralPath $deployPath -PathType Leaf)){throw "Deploy-Application.ps1 not found:`n$deployPath"}

        $ise=Get-Command powershell_ise.exe -ErrorAction SilentlyContinue
        if($ise){
            Start-Process -FilePath $ise.Source -ArgumentList @('-File',$deployPath)
        } else {
            Start-Process -FilePath notepad.exe -ArgumentList @($deployPath)
        }
        $buildStatus.Text='EDITING - Save Deploy-Application.ps1, then return here and click CONFIRM AND BUILD INTUNEWIN.'
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Open Deploy-Application.ps1',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$confirmBuildButton.Add_Click({
    try {
        if(-not $script:PackagePrepared){throw 'Prepare the package first.'}
        $deployPath=Join-Path $script:PreparedPackageRoot 'Package\Deploy-Application.ps1'
        if(-not(Test-Path -LiteralPath $deployPath -PathType Leaf)){throw "Deploy-Application.ps1 not found:`n$deployPath"}

        $buildStatus.Text='BUILDING - Using the reviewed Deploy-Application.ps1 exactly as saved.'
        $status.Text='RUNNING'; $status.ForeColor=[System.Drawing.Color]::Gold
        $prepareButton.Enabled=$false; $openDeployButton.Enabled=$false; $confirmBuildButton.Enabled=$false
        $tabs.SelectedTab=$runTab

        $engineParameters=Get-BuildEngineParameters
        $engineParameters.ContinuePrepared = $true
        $engineParameters.GenerateIntuneWin = $true
        Invoke-PackagingEngineForWorkflow -EngineParameters $engineParameters

        $script:ScriptConfirmed=$true
        $status.Text='COMPLETED'; $status.ForeColor=[System.Drawing.Color]::LightGreen
        $buildStatus.Text='BUILD COMPLETED SUCCESSFULLY - Reviewed Deploy-Application.ps1 converted to IntuneWin.'
        $tabs.SelectedTab=$buildTab
        Update-BuildSummary
        [System.Windows.Forms.MessageBox]::Show('Package build completed successfully.','Packaging Complete',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        $status.Text='FAILED'; $status.ForeColor=[System.Drawing.Color]::Tomato
        $buildStatus.Text='BUILD FAILED - Review Run Log and package-specific log.'
        $tabs.SelectedTab=$runTab
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Packaging Failed',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $prepareButton.Enabled=$true
        $openDeployButton.Enabled=$script:PackagePrepared
        $confirmBuildButton.Enabled=$script:PackagePrepared
        Update-BuildSummary
    }
})


# Select Package tab and show the window.
$tabs.SelectedTab = $general
Update-BuildSummary
$tabs.Visible = $true

$null = $form.ShowDialog()
