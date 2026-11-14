Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Root = Split-Path -Parent $PSScriptRoot
$script:Engine = Join-Path $script:Root 'Scripts\Invoke-PackagingEngine.ps1'
$script:SelectedMainInstallerPath = ''

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
                } else {
                    # Installer selection is the first step: use its parent folder as source root.
                    $src.Text = [IO.Path]::GetDirectoryName($selected)
                    $box.Text = [IO.Path]::GetFileName($selected)
                }
                $box.Focus()
            }
        } finally { $d.Dispose() }
    })
    [void]$Parent.Controls.Add($b)
    return $b
}

# ---------------- Window ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Application Packaging Automation - V2'
$form.StartPosition = 'CenterScreen'
# Always fit the application inside the current monitor working area.
$workingArea = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$initialWidth  = [Math]::Min(1280, [Math]::Max(1050, $workingArea.Width - 20))
$initialHeight = [Math]::Min(850,  [Math]::Max(700,  $workingArea.Height - 20))
$form.Size = New-Object System.Drawing.Size($initialWidth,$initialHeight)
$form.MinimumSize = New-Object System.Drawing.Size(1050,700)
$form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
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

# APA logo created entirely in WinForms (no external image file required).
$apaLogo = New-Object System.Windows.Forms.Panel
$apaLogo.Size = New-Object System.Drawing.Size(48,48)
$apaLogo.Location = New-Object System.Drawing.Point(24,15)
$apaLogo.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$header.Controls.Add($apaLogo)
$apaLogoText = New-Object System.Windows.Forms.Label
$apaLogoText.Text = 'APA'
$apaLogoText.Dock = [System.Windows.Forms.DockStyle]::Fill
$apaLogoText.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$apaLogoText.Font = New-Object System.Drawing.Font('Segoe UI Semibold',10)
$apaLogoText.ForeColor = [System.Drawing.Color]::White
$apaLogo.Controls.Add($apaLogoText)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'Application Packaging Automation'
$title.Font = New-Object System.Drawing.Font('Segoe UI Semibold',19)
$title.ForeColor = [System.Drawing.Color]::White
$title.Location = New-Object System.Drawing.Point(84,10)
$title.AutoSize = $true
$header.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = 'MSI / EXE  |  PSADT 3.10 / 4.x  |  Configuration Builder  |  PID  |  IntuneWin'
$subtitle.ForeColor = [System.Drawing.Color]::Gainsboro
$subtitle.Location = New-Object System.Drawing.Point(87,47)
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
$configTab = New-Tab 'Configuration'
# Keep the existing PID metadata controls off-screen; PID generation remains part of Build & Output.
$pidTab = New-Object System.Windows.Forms.TabPage
$install = New-Tab 'Install'
$uninstall = New-Tab 'Uninstall'
$advanced = New-Tab 'Templates / Configuration'
$buildTab = New-Tab 'Build & Output'
$nonAutoTab = New-Tab 'Non-Automated IntuneWin'
$runTab = New-Tab 'Run Log'

# ---------------- Package ----------------
$appLabel = New-Label $general 'Application Name' 28 25
$app = New-TextBox $general 28 48 300
$verLabel = New-Label $general 'Version' 350 25
$ver = New-TextBox $general 350 48 160
$vendorLabel = New-Label $general 'Vendor' 535 25
$vendor = New-TextBox $general 535 48 240
$archLabel = New-Label $general 'Architecture' 805 25
$arch = New-Combo $general 805 48 180 @('x64','x86','ARM64')
# PSADT version selector is kept on the top row so it is always visible.
New-Label $general 'PSADT Version' 1020 94
$psadtVersion = New-Combo $general 1020 117 130 @('3.10','4.x')
$psadtVersion.SelectedIndex = 0

$typeLabel = New-Label $general 'Installer Type' 28 25
$type = New-Combo $general 28 48 180 @('MSI','EXE')
$sourceLabel = New-Label $general 'Source Folder' 225 25
$source = New-TextBox $general 225 48 680
$sourceBrowse = Add-FolderButton $general 915 48 $source 'Select the vendor source folder'

$installerLabel = New-Label $general 'Select Main Installer' 28 94
$installer = New-TextBox $general 28 117 870
$installerBrowse = New-Object System.Windows.Forms.Button
$installerBrowse.Text='Browse...'; $installerBrowse.Location=New-Object System.Drawing.Point(915,117); $installerBrowse.Size=New-Object System.Drawing.Size(90,28)
$installerBrowse.Add_Click({
    $fd=New-Object System.Windows.Forms.OpenFileDialog
    $fd.Filter='Installer files (*.msi;*.exe)|*.msi;*.exe|All files (*.*)|*.*'
    if($source.Text -and (Test-Path -LiteralPath $source.Text -PathType Container)){$fd.InitialDirectory=$source.Text}
    try {
        if($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $full=[IO.Path]::GetFullPath($fd.FileName)
            $isNewMainInstaller = -not [string]::Equals($full,$script:SelectedMainInstallerPath,[StringComparison]::OrdinalIgnoreCase)
            $source.Text=[IO.Path]::GetDirectoryName($full)
            $installer.Text=[IO.Path]::GetFileName($full)
            $ext=[IO.Path]::GetExtension($full)
            if($isNewMainInstaller) { $app.Clear(); $ver.Clear(); $vendor.Clear() }
            if($ext -ieq '.msi') {
                $type.SelectedItem='MSI'; $mst.Enabled=$true; $mst.Checked=$true
                $installParams.Text='/qn /norestart'; $uParams.Text='/qn /norestart'
                Read-MsiPropertiesToUI -MsiPath $full -Force
            } elseif($ext -ieq '.exe') {
                $type.SelectedItem='EXE'; $mst.Enabled=$false; $mst.Checked=$false
                $msiProductName.Clear(); $msiProductVersion.Clear(); $msiManufacturer.Clear(); $msiProductCode.Clear(); $msiUpgradeCode.Clear()
                $fi=[Diagnostics.FileVersionInfo]::GetVersionInfo($full)
                if($isNewMainInstaller) {
                    if($fi.ProductName){$app.Text=$fi.ProductName.Trim()} elseif(-not $app.Text){$app.Text=[IO.Path]::GetFileNameWithoutExtension($full)}
                    if($fi.ProductVersion){$ver.Text=$fi.ProductVersion.Trim()} elseif($fi.FileVersion){$ver.Text=$fi.FileVersion.Trim()}
                    if($fi.CompanyName){$vendor.Text=$fi.CompanyName.Trim()}
                }
                Find-AndPopulateSilentSwitch -InstallerPath $full -ParameterBox $installParams -StatusLabel $installSwitchStatus -ContextName 'Installation'
                Find-InstalledUninstallEntry
            }
            $script:SelectedMainInstallerPath=$full
        }
    } finally {$fd.Dispose()}
})
$general.Controls.Add($installerBrowse)

# Metadata follows installer selection and is populated from the selected installer.
$appLabel.Location = New-Object System.Drawing.Point(28,163)
$app.Location = New-Object System.Drawing.Point(28,186)
$verLabel.Location = New-Object System.Drawing.Point(350,163)
$ver.Location = New-Object System.Drawing.Point(350,186)
$vendorLabel.Location = New-Object System.Drawing.Point(535,163)
$vendor.Location = New-Object System.Drawing.Point(535,186)
$archLabel.Location = New-Object System.Drawing.Point(805,163)
$arch.Location = New-Object System.Drawing.Point(805,186)

$mst = New-Object System.Windows.Forms.CheckBox
$mst.Text = 'Generate Standard MST (MSI only)'
$mst.Location = New-Object System.Drawing.Point(28,230)
$mst.AutoSize = $true
$mst.Checked = $true
$general.Controls.Add($mst)

$pidGenerationCheckBox = New-Object System.Windows.Forms.CheckBox
$pidGenerationCheckBox.Text = 'Generate PID document automatically'
$pidGenerationCheckBox.Location = New-Object System.Drawing.Point(28,270)
$pidGenerationCheckBox.AutoSize = $true
$pidGenerationCheckBox.Checked = $true
$pidGenerationCheckBox.Enabled = $false
$general.Controls.Add($pidGenerationCheckBox)

$intune = New-Object System.Windows.Forms.CheckBox
$intune.Text = 'Generate .intunewin'
$intune.Location = New-Object System.Drawing.Point(28,270)
$intune.AutoSize = $true
$intune.Checked = $true
$general.Controls.Add($intune)

$serviceUI = New-Object System.Windows.Forms.CheckBox
$serviceUI.Text = 'Include ServiceUI.exe (copy from PSADT template)'
$serviceUI.Location = New-Object System.Drawing.Point(28,315)
$serviceUI.AutoSize = $true
$serviceUI.Checked = $false
$general.Controls.Add($serviceUI)

New-Label $general 'AppLogo Source (optional EXE/DLL/ICO)' 28 355
$appLogoSource = New-TextBox $general 28 378 870
$appLogoBrowse = New-Object System.Windows.Forms.Button
$appLogoBrowse.Text = 'Browse...'
$appLogoBrowse.Location = New-Object System.Drawing.Point(915,378)
$appLogoBrowse.Size = New-Object System.Drawing.Size(90,28)
$appLogoBrowse.Add_Click({
    $d = New-Object System.Windows.Forms.OpenFileDialog
    $d.Filter = 'Icon source (*.exe;*.dll;*.ico)|*.exe;*.dll;*.ico|All files (*.*)|*.*'
    try {
        if($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){ $appLogoSource.Text = $d.FileName }
    } finally { $d.Dispose() }
})
$general.Controls.Add($appLogoBrowse)

$hint = New-Object System.Windows.Forms.Label
$hint.Text = 'Select the main installer first. Metadata and MSI ProductCode are read automatically where available; EXE metadata may be incomplete.'
$hint.ForeColor = [System.Drawing.Color]::DimGray
$hint.Location = New-Object System.Drawing.Point(28,485)
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
$pidTemplateDisplay.Text = 'Auto-detected under <ToolRoot>\Templates\PID\Package_Template.docx'
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

# ---------------- Install / Uninstall direct editors ----------------
New-Label $install 'Main Installer Parameters' 28 25
$installParams = New-TextBox $install 28 48 1020
$installParams.Text = '/qn /norestart'
$installSwitchStatus=New-Object System.Windows.Forms.Label; $installSwitchStatus.Text='Silent Switch Finder: waiting for installer selection...'; $installSwitchStatus.Location=New-Object System.Drawing.Point(28,72); $installSwitchStatus.Size=New-Object System.Drawing.Size(1020,20); $installSwitchStatus.ForeColor=[System.Drawing.Color]::DimGray; $install.Controls.Add($installSwitchStatus)
New-Label $install 'Pre-Installation (older-version removal and prerequisite commands)' 28 100
$preInstall = New-TextBox $install 28 124 1020 105
$preInstall.Multiline = $true; $preInstall.ScrollBars = 'Vertical'; $preInstall.AcceptsReturn = $true
New-Label $install 'Installation (additional install commands; main installer is generated automatically)' 28 238
$installSteps = New-TextBox $install 28 262 1020 95
$installSteps.Multiline = $true; $installSteps.ScrollBars = 'Vertical'; $installSteps.AcceptsReturn = $true
New-Label $install 'Post-Installation' 28 366
$postInstall = New-TextBox $install 28 390 1020 85
$postInstall.Multiline = $true; $postInstall.ScrollBars = 'Vertical'; $postInstall.AcceptsReturn = $true
New-Label $install 'Close Apps (optional - e.g. WINWORD, EXCEL)' 28 484
$installCloseApps = New-TextBox $install 28 508 1020
$install.AutoScroll = $true

New-Label $uninstall 'Uninstallation Method' 28 25
$uMethod = New-Combo $uninstall 28 48 360 @('Main EXE / Installer','Uninstall from different location','MSI ProductCode')
New-Label $uninstall 'Uninstall EXE full path / ProductCode (auto-detected where possible)' 420 25
$uFile = New-TextBox $uninstall 420 48 520
$uBrowse = New-Object System.Windows.Forms.Button; $uBrowse.Text='Browse...'; $uBrowse.Location=New-Object System.Drawing.Point(950,48); $uBrowse.Size=New-Object System.Drawing.Size(90,28)
$uBrowse.Add_Click({
    $fd=New-Object System.Windows.Forms.OpenFileDialog
    $fd.Filter='Uninstallers (*.exe;*.cmd;*.bat;*.msi)|*.exe;*.cmd;*.bat;*.msi|All files (*.*)|*.*'
    if($uFile.Text -and (Test-Path -LiteralPath $uFile.Text -PathType Leaf)){$fd.InitialDirectory=[IO.Path]::GetDirectoryName($uFile.Text)}
    elseif($source.Text -and (Test-Path -LiteralPath $source.Text -PathType Container)){$fd.InitialDirectory=$source.Text}
    try { if($fd.ShowDialog() -eq 'OK') {
        $uMethod.SelectedItem='Uninstall from different location'
        if([IO.Path]::GetExtension($fd.FileName) -ieq '.msi'){
            $uMethod.SelectedItem='MSI ProductCode'
            try {
                $wiU=New-Object -ComObject WindowsInstaller.Installer
                $dbU=$wiU.OpenDatabase([IO.Path]::GetFullPath($fd.FileName),0)
                $vwU=$dbU.OpenView("SELECT `Value` FROM Property WHERE `Property`='ProductCode'");$vwU.Execute();$rU=$vwU.Fetch()
                if($rU){$uFile.Text=[string]$rU.StringData(1);$uParams.Text='/qn /norestart';$uninstallSwitchStatus.Text='MSI uninstall ProductCode detected automatically.';$uninstallSwitchStatus.ForeColor=[System.Drawing.Color]::DarkGreen}
                else {$uFile.Text='';$uninstallSwitchStatus.Text='Could not read MSI ProductCode; enter it manually.';$uninstallSwitchStatus.ForeColor=[System.Drawing.Color]::DarkOrange}
                try{$vwU.Close()}catch{};try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($dbU)}catch{};try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($wiU)}catch{}
            } catch {$uFile.Text='';[System.Windows.Forms.MessageBox]::Show('Could not read ProductCode from selected MSI. Enter the ProductCode manually.')}
        } else {
            $uMethod.SelectedItem='Uninstall from different location'
            $uFile.Text=[IO.Path]::GetFullPath($fd.FileName)
            Find-AndPopulateSilentSwitch -InstallerPath $fd.FileName -ParameterBox $uParams -StatusLabel $uninstallSwitchStatus -ContextName 'Uninstallation'
        }
    }} finally {$fd.Dispose()}
})
$uninstall.Controls.Add($uBrowse)
New-Label $uninstall 'Uninstall Parameters' 28 92
$uParams = New-TextBox $uninstall 28 115 480
$uParams.Text = '/qn /norestart'
$uninstallSwitchStatus=New-Object System.Windows.Forms.Label; $uninstallSwitchStatus.Text='Uninstall switch detection: waiting for selection...'; $uninstallSwitchStatus.Location=New-Object System.Drawing.Point(28,138); $uninstallSwitchStatus.Size=New-Object System.Drawing.Size(480,20); $uninstallSwitchStatus.ForeColor=[System.Drawing.Color]::DimGray; $uninstall.Controls.Add($uninstallSwitchStatus)
New-Label $uninstall 'Pre-Uninstallation' 28 170
$preUninstall = New-TextBox $uninstall 28 194 1010 80
$preUninstall.Multiline = $true; $preUninstall.ScrollBars = 'Vertical'; $preUninstall.AcceptsReturn = $true
New-Label $uninstall 'Uninstallation (additional uninstall commands; main uninstall is generated automatically)' 28 282
$uninstallSteps = New-TextBox $uninstall 28 306 1010 80
$uninstallSteps.Multiline = $true; $uninstallSteps.ScrollBars = 'Vertical'; $uninstallSteps.AcceptsReturn = $true
New-Label $uninstall 'Post-Uninstallation' 28 394
$postUninstall = New-TextBox $uninstall 28 418 1010 80
$postUninstall.Multiline = $true; $postUninstall.ScrollBars = 'Vertical'; $postUninstall.AcceptsReturn = $true
New-Label $uninstall 'Close Apps (optional - e.g. WINWORD, EXCEL)' 28 506
$uninstallCloseApps = New-TextBox $uninstall 28 530 1010
$uninstall.AutoScroll = $true

# ---------------- Configuration Builder ----------------
New-Label $configTab 'Additional Installers' 24 18
$additionalInstallerInfo = New-Label $configTab 'Use one list for prerequisites and additional components. Each entry can have its own install/uninstall command and phase.' 24 44
$additionalInstallerInfo.ForeColor = [System.Drawing.Color]::DimGray
$additionalList = New-Object System.Windows.Forms.ListView
$additionalList.Location = New-Object System.Drawing.Point(24,72); $additionalList.Size = New-Object System.Drawing.Size(1090,140)
$additionalList.View = [System.Windows.Forms.View]::Details; $additionalList.FullRowSelect = $true; $additionalList.GridLines = $true
[void]$additionalList.Columns.Add('Installer',220); [void]$additionalList.Columns.Add('Detected Metadata',220); [void]$additionalList.Columns.Add('Install Phase',135); [void]$additionalList.Columns.Add('Uninstall Phase',145); [void]$additionalList.Columns.Add('Detection',155); [void]$additionalList.Columns.Add('Parameters',210)
$configTab.Controls.Add($additionalList)
$addInstallerButton = New-Object System.Windows.Forms.Button; $addInstallerButton.Text = '+ Add Installer'; $addInstallerButton.Location = New-Object System.Drawing.Point(24,220); $addInstallerButton.Size = New-Object System.Drawing.Size(140,32); $configTab.Controls.Add($addInstallerButton)
$removeInstallerButton = New-Object System.Windows.Forms.Button; $removeInstallerButton.Text = 'Remove Selected'; $removeInstallerButton.Location = New-Object System.Drawing.Point(174,220); $removeInstallerButton.Size = New-Object System.Drawing.Size(140,32); $configTab.Controls.Add($removeInstallerButton)
$editInstallerButton = New-Object System.Windows.Forms.Button; $editInstallerButton.Text = 'Edit Selected'; $editInstallerButton.Location = New-Object System.Drawing.Point(324,220); $editInstallerButton.Size = New-Object System.Drawing.Size(130,32); $configTab.Controls.Add($editInstallerButton)
New-Label $configTab 'Configuration Actions' 24 266
New-Label $configTab 'Action' 24 294
$actionType = New-Combo $configTab 24 318 235 @('Copy single file','Copy files / folder recursively','Remove file','Remove folder','Set registry value','Remove registry key','Import .reg file','Create shortcut','Install certificate','Install driver (PnPUtil)','Install driver (DPInst)','Set environment variable','Start service','Stop service','Create firewall rule','Remove firewall rule','Set file/folder permissions','Set HKCU registry value for all users','Copy file to all user profiles','Start-Sleep','Wait-Process','Write-Log','Show restart prompt','Custom PowerShell','Detect and uninstall older version')
New-Label $configTab 'Execution phase' 274 294
$actionPhase = New-Combo $configTab 274 318 220 @('Pre-Installation','Installation','Post-Installation','Pre-Uninstallation','Uninstallation','Post-Uninstallation')
$actionSourceLabel = New-Label $configTab 'Source / Key / Process / Command' 514 294
$actionSource = New-TextBox $configTab 514 318 205
$actionSourceBrowse = New-Object System.Windows.Forms.Button; $actionSourceBrowse.Text='Browse...'; $actionSourceBrowse.Location=New-Object System.Drawing.Point(724,318); $actionSourceBrowse.Size=New-Object System.Drawing.Size(72,28); $configTab.Controls.Add($actionSourceBrowse)
$actionDestinationLabel = New-Label $configTab 'Destination / Value / Parameters' 804 294
$actionDestination = New-TextBox $configTab 804 318 220
$actionDestinationBrowse = New-Object System.Windows.Forms.Button; $actionDestinationBrowse.Text='Browse...'; $actionDestinationBrowse.Location=New-Object System.Drawing.Point(1029,318); $actionDestinationBrowse.Size=New-Object System.Drawing.Size(75,28); $configTab.Controls.Add($actionDestinationBrowse)
$actionExtraLabel = New-Label $configTab 'Additional value / type / timeout / message' 24 354
$actionExtra = New-TextBox $configTab 24 378 350
$actionRegistryTypeLabel = New-Label $configTab 'Registry Type' 390 354
$actionRegistryType = New-Combo $configTab 390 378 150 @('String','DWord')
$actionRegistryType.SelectedItem = 'String'
$actionOption1Label = New-Label $configTab 'Option' 390 354
$actionOption1 = New-Combo $configTab 390 378 150 @('Machine','User')
$actionOption2Label = New-Label $configTab 'Option' 560 354
$actionOption2 = New-Combo $configTab 560 378 150 @('Inbound','Outbound')
$actionOption3Label = New-Label $configTab 'Option' 730 354
$actionOption3 = New-Combo $configTab 730 378 150 @('TCP','UDP','Any')
$actionOption4Label = New-Label $configTab 'Option' 900 354
$actionOption4 = New-Combo $configTab 900 378 150 @('Allow','Block')
New-Label $configTab 'Preview of all configuration actions (edit final commands in phase editors)' 24 416
$actionPreview = New-TextBox $configTab 24 440 1080 105
$actionPreview.Multiline = $true; $actionPreview.ScrollBars = 'Vertical'; $actionPreview.AcceptsReturn = $true; $actionPreview.Font = New-Object System.Drawing.Font('Consolas',9)
$previewActionButton = New-Object System.Windows.Forms.Button; $previewActionButton.Text = 'Generate Preview'; $previewActionButton.Location = New-Object System.Drawing.Point(24,556); $previewActionButton.Size = New-Object System.Drawing.Size(150,34); $configTab.Controls.Add($previewActionButton)
$addActionButton = New-Object System.Windows.Forms.Button; $addActionButton.Text = 'Add to Selected Phase'; $addActionButton.Location = New-Object System.Drawing.Point(184,556); $addActionButton.Size = New-Object System.Drawing.Size(190,34); $configTab.Controls.Add($addActionButton)
$clearConfigButton = New-Object System.Windows.Forms.Button; $clearConfigButton.Text = 'Clear Configuration Actions'; $clearConfigButton.Location = New-Object System.Drawing.Point(384,556); $clearConfigButton.Size = New-Object System.Drawing.Size(210,34); $configTab.Controls.Add($clearConfigButton)
$actionHelp = New-Label $configTab 'Preview lists all added configuration actions plus the current selection. Commands are inserted into phase editors; review/edit them there before building.' 24 602
$actionHelp.ForeColor = [System.Drawing.Color]::DimGray

# ---------------- Templates & Output ----------------
New-Label $advanced 'Template Root' 28 25
$templateBox = New-TextBox $advanced 28 48 850
$templateBox.Text = Join-Path $script:Root 'Templates'
[void](Add-FolderButton $advanced 890 48 $templateBox 'Select template root')

New-Label $advanced 'Output Root' 28 95
$outputBox = New-TextBox $advanced 28 118 850
$outputBox.Text = Join-Path $script:Root 'Output'
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

    $packageTemplate = Join-Path $rootPath 'PID\Package_Template.docx'
    if (Test-Path -LiteralPath $packageTemplate -PathType Leaf) {
        $items += 'Package_Template.docx found'
    } else {
        $items += ('Package_Template.docx NOT found: ' + $packageTemplate)
    }

    $toolsRoot = Join-Path $rootPath 'Tools'
    $intuneUtil = Join-Path $toolsRoot 'IntuneWinAppUtil.exe'
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
        (Test-Path -LiteralPath (Join-Path $rootPath 'PSADT\3.10\Deploy-Application.ps1')) -and
        (Test-Path -LiteralPath (Join-Path $rootPath 'PSADT\3.10\AppDeployToolkit'))) {
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
$buildSummary.Size = New-Object System.Drawing.Size(700,300)
$buildSummary.Multiline = $true
$buildSummary.ReadOnly = $true
$buildSummary.ScrollBars = 'Vertical'
$buildSummary.Font = New-Object System.Drawing.Font('Consolas',10)
$buildTab.Controls.Add($buildSummary)

$script:PackagePrepared = $false
$script:ScriptConfirmed = $false
$script:PreparedPackageRoot = ''

$buildStatus = New-Object System.Windows.Forms.Label
$buildStatus.Text = 'READY - Workflow: Configure -> Prepare PSADT -> Generate PID -> Review/Edit -> Build IntuneWin.'
$buildStatus.Location = New-Object System.Drawing.Point(28,438)
$buildStatus.Size = New-Object System.Drawing.Size(1020,35)
$buildStatus.ForeColor = [System.Drawing.Color]::DimGray
$buildTab.Controls.Add($buildStatus)

$buildButtonsHeader = New-Object System.Windows.Forms.Label
$buildButtonsHeader.Text = 'BUILD CONTROLS'
$buildButtonsHeader.Font = New-Object System.Drawing.Font('Segoe UI Semibold',9)
$buildButtonsHeader.Location = New-Object System.Drawing.Point(755,38)
$buildButtonsHeader.AutoSize = $true
$buildButtonsHeader.ForeColor = [System.Drawing.Color]::DimGray
$buildTab.Controls.Add($buildButtonsHeader)

$prepareButton = New-Object System.Windows.Forms.Button
$prepareButton.Text = 'PREPARE PACKAGE'
$prepareButton.Font = New-Object System.Drawing.Font('Segoe UI Semibold',11)
$prepareButton.Location = New-Object System.Drawing.Point(755,62)
$prepareButton.Size = New-Object System.Drawing.Size(300,42)
$prepareButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$prepareButton.ForeColor = [System.Drawing.Color]::White
$prepareButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$prepareButton.FlatAppearance.BorderSize = 0
$buildTab.Controls.Add($prepareButton)

$openDeployButton = New-Object System.Windows.Forms.Button
$openDeployButton.Text = 'REVIEW / EDIT DEPLOY-APPLICATION.PS1'
$openDeployButton.Location = New-Object System.Drawing.Point(755,110)
$openDeployButton.Size = New-Object System.Drawing.Size(300,42)
$openDeployButton.Enabled = $false
$buildTab.Controls.Add($openDeployButton)

$confirmBuildButton = New-Object System.Windows.Forms.Button
$confirmBuildButton.Text = 'CONFIRM AND BUILD INTUNEWIN'
$confirmBuildButton.Font = New-Object System.Drawing.Font('Segoe UI Semibold',11)
$confirmBuildButton.Location = New-Object System.Drawing.Point(755,158)
$confirmBuildButton.Size = New-Object System.Drawing.Size(300,42)
$confirmBuildButton.BackColor = [System.Drawing.Color]::FromArgb(40,167,69)
$confirmBuildButton.ForeColor = [System.Drawing.Color]::White
$confirmBuildButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$confirmBuildButton.FlatAppearance.BorderSize = 0
$confirmBuildButton.Enabled = $false
$buildTab.Controls.Add($confirmBuildButton)

$validateBuildButton = New-Object System.Windows.Forms.Button
$validateBuildButton.Text = 'VALIDATE TEMPLATES'
$validateBuildButton.Location = New-Object System.Drawing.Point(755,206)
$validateBuildButton.Size = New-Object System.Drawing.Size(300,38)
$buildTab.Controls.Add($validateBuildButton)

$openOutputButton = New-Object System.Windows.Forms.Button
$openOutputButton.Text = 'OPEN OUTPUT'
$openOutputButton.Location = New-Object System.Drawing.Point(755,252)
$openOutputButton.Size = New-Object System.Drawing.Size(300,38)
$buildTab.Controls.Add($openOutputButton)

$openLogsButton = New-Object System.Windows.Forms.Button
$openLogsButton.Text = 'OPEN LOGS'
$openLogsButton.Location = New-Object System.Drawing.Point(755,298)
$openLogsButton.Size = New-Object System.Drawing.Size(300,38)
$buildTab.Controls.Add($openLogsButton)

$openPackageButton = New-Object System.Windows.Forms.Button
$openPackageButton.Text = 'OPEN PACKAGE FOLDER'
$openPackageButton.Location = New-Object System.Drawing.Point(755,344)
$openPackageButton.Size = New-Object System.Drawing.Size(300,38)
$buildTab.Controls.Add($openPackageButton)
$newPackageButton = New-Object System.Windows.Forms.Button
$newPackageButton.Text = 'REFRESH / NEW PACKAGE'
$newPackageButton.Location = New-Object System.Drawing.Point(755,392)
$newPackageButton.Size = New-Object System.Drawing.Size(300,38)
$buildTab.Controls.Add($newPackageButton)
$newPackageButton.Add_Click({
    $dirty = ($app.Text -or $ver.Text -or $vendor.Text -or $source.Text -or $installer.Text -or $preInstall.Text -or $installSteps.Text -or $postInstall.Text -or $preUninstall.Text -or $uninstallSteps.Text -or $postUninstall.Text -or $additionalList.Items.Count -gt 0)
    if($dirty -and [System.Windows.Forms.MessageBox]::Show('Start a new packaging session? Saved settings, templates, logs and previous outputs will be preserved.','New Package',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes){ return }
    foreach($box in @($app,$ver,$vendor,$installer,$installParams,$preInstall,$installSteps,$postInstall,$uFile,$uParams,$preUninstall,$uninstallSteps,$postUninstall,$installCloseApps,$uninstallCloseApps,$actionSource,$actionDestination,$actionExtra,$actionPreview)){ if($box){$box.Clear()} }
    $source.Text=''; $type.SelectedItem='MSI'; $arch.SelectedItem='x64'; $psadtVersion.SelectedItem='3.10'; if($actionRegistryType){$actionRegistryType.SelectedItem='String'}; if($actionType){$actionType.SelectedIndex=0}; $uMethod.SelectedItem='Main EXE / Installer'; $installParams.Text='/qn /norestart'; $uParams.Text='/qn /norestart'
    $additionalList.Items.Clear(); $script:AdditionalInstallers.Clear(); if($script:ConfigurationActionRecords){$script:ConfigurationActionRecords.Clear()}; $script:PackagePrepared=$false; $script:ScriptConfirmed=$false; $script:PreparedPackageRoot=''; $runBox.Clear(); $buildSummary.Clear(); if($uninstallSwitchStatus){$uninstallSwitchStatus.Text='';$uninstallSwitchStatus.ForeColor=[System.Drawing.Color]::DimGray}; $buildStatus.Text='READY - Start a new package.'; $status.Text='READY'; $status.ForeColor=[System.Drawing.Color]::LightGreen
    $openDeployButton.Enabled=$false; $confirmBuildButton.Enabled=$false; $tabs.SelectedTab=$general
})

function Update-BuildSummary {
    $pkgName = if($arch.SelectedItem -eq 'x86'){ '{0}_{1}_{2}_P1' -f ($vendor.Text.Trim() -replace '\s+','-'),($app.Text.Trim() -replace '\s+','-'),($ver.Text.Trim() -replace '\s+','-') } else { '{0}_{1}_{2}_P1_{3}' -f ($vendor.Text.Trim() -replace '\s+','-'),($app.Text.Trim() -replace '\s+','-'),($ver.Text.Trim() -replace '\s+','-'),$arch.SelectedItem }
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
    $pkgName = if($arch.SelectedItem -eq 'x86'){ '{0}_{1}_{2}_P1' -f ($vendor.Text.Trim() -replace '\s+','-'),($app.Text.Trim() -replace '\s+','-'),($ver.Text.Trim() -replace '\s+','-') } else { '{0}_{1}_{2}_P1_{3}' -f ($vendor.Text.Trim() -replace '\s+','-'),($app.Text.Trim() -replace '\s+','-'),($ver.Text.Trim() -replace '\s+','-'),$arch.SelectedItem }
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
    param([string]$MsiPath,[switch]$Force)
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
        $productCode = ([string]$values['ProductCode']).Trim()
        $msiProductName.Text = ([string]$values['ProductName']).Trim()
        $msiProductVersion.Text = ([string]$values['ProductVersion']).Trim()
        $msiManufacturer.Text = ([string]$values['Manufacturer']).Trim()
        # Populate the visible Package metadata fields immediately after selection.
        if(-not [string]::IsNullOrWhiteSpace($values['ProductName']) -and ($Force -or [string]::IsNullOrWhiteSpace($app.Text))) { $app.Text = ([string]$values['ProductName']).Trim() }
        if(-not [string]::IsNullOrWhiteSpace($values['ProductVersion']) -and ($Force -or [string]::IsNullOrWhiteSpace($ver.Text))) { $ver.Text = ([string]$values['ProductVersion']).Trim() }
        if(-not [string]::IsNullOrWhiteSpace($values['Manufacturer']) -and ($Force -or [string]::IsNullOrWhiteSpace($vendor.Text))) { $vendor.Text = ([string]$values['Manufacturer']).Trim() }
        if($Force) {
            try {
                $summary = $database.SummaryInformation(0)
                $templateProperty = [string]$summary.Property(7)
                if($templateProperty -match '(?i)^x64') { $arch.SelectedItem = 'x64' }
                elseif($templateProperty -match '(?i)^Intel') { $arch.SelectedItem = 'x86' }
                elseif($templateProperty -match '(?i)^Arm64') { $arch.SelectedItem = 'ARM64' }
            } catch {}
        }
        $msiProductCode.Text = $productCode
        $msiUpgradeCode.Text = ([string]$values['UpgradeCode']).Trim()
        if(-not [string]::IsNullOrWhiteSpace($productCode)){
            if($productCode -notmatch '^\{[0-9A-Fa-f-]{36}\}$'){
                throw "Invalid MSI ProductCode detected: '$productCode'"
            }
            # Keep the UI and generated PSADT uninstall command in sync as soon
            # as the MSI is selected, not only when Prepare Package is clicked.
            if($Force -or [string]::IsNullOrWhiteSpace($uFile.Text)) {
                $uMethod.SelectedItem = 'MSI ProductCode'
                $uFile.Text = $productCode
                $uParams.Text = '/qn /norestart'
            }
            if($Force -or [string]::IsNullOrWhiteSpace($installParams.Text)) { $installParams.Text = '/qn /norestart' }
            $uninstallSwitchStatus.Text = 'MSI ProductCode detected automatically. Silent parameters: /qn /norestart'
            $uninstallSwitchStatus.ForeColor=[System.Drawing.Color]::DarkGreen
            $installSwitchStatus.Text = 'MSI selected. Silent parameters: /qn /norestart (Silent Switch Finder skipped).'
            $installSwitchStatus.ForeColor=[System.Drawing.Color]::DarkGreen
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

        $srcPath = [IO.Path]::GetFullPath($srcPath).TrimEnd('\')
        $outPath = [IO.Path]::GetFullPath($outPath).TrimEnd('\')
        if($outPath.Equals($srcPath,[StringComparison]::OrdinalIgnoreCase) -or $outPath.StartsWith($srcPath + '\',[StringComparison]::OrdinalIgnoreCase)){
            throw 'Output Folder must be outside the Source Folder. Select a separate output folder.'
        }

        $nonAutoStatus.Text = 'BUILDING - Creating IntuneWin...'
        $nonAutoStatus.ForeColor = [System.Drawing.Color]::DarkOrange
        $nonAutoLog.Clear()
        $nonAutoLog.AppendText(('Utility: ' + $util.FullName) + [Environment]::NewLine)
        $nonAutoLog.AppendText(('Source : ' + $srcPath) + [Environment]::NewLine)
        $nonAutoLog.AppendText(('Setup  : ' + $setupArgument) + [Environment]::NewLine)
        $nonAutoLog.AppendText(('Output : ' + $outPath) + [Environment]::NewLine)

        $buildStarted = Get-Date
        $stdoutLog = Join-Path $outPath 'IntuneWinAppUtil-output.log'
        $stderrLog = Join-Path $outPath 'IntuneWinAppUtil-error.log'
        Remove-Item -LiteralPath $stdoutLog,$stderrLog -Force -ErrorAction SilentlyContinue
        $p = Start-Process -FilePath $util.FullName -WorkingDirectory (Split-Path -Parent $util.FullName) -ArgumentList @('-c',$srcPath,'-s',$setupArgument,'-o',$outPath,'-q') -Wait -PassThru -NoNewWindow -RedirectStandardOutput $stdoutLog -RedirectStandardError $stderrLog
        if($p.ExitCode -ne 0){
            $details = ''
            if(Test-Path -LiteralPath $stderrLog){ $details = (Get-Content -LiteralPath $stderrLog -Raw -ErrorAction SilentlyContinue).Trim() }
            if([string]::IsNullOrWhiteSpace($details) -and (Test-Path -LiteralPath $stdoutLog)){ $details = (Get-Content -LiteralPath $stdoutLog -Raw -ErrorAction SilentlyContinue).Trim() }
            throw "IntuneWinAppUtil.exe failed with exit code $($p.ExitCode). $details"
        }

        # Different IntuneWinAppUtil versions may place the file differently.
        # Search the requested output first, then the source as a safe fallback.
        $created = @(Get-ChildItem -LiteralPath $outPath -Filter '*.intunewin' -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $buildStarted.AddSeconds(-2) } | Sort-Object LastWriteTime -Descending | Select-Object -First 1)[0]
        if(-not $created){
            $created = @(Get-ChildItem -LiteralPath $srcPath -Filter '*.intunewin' -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $buildStarted.AddSeconds(-2) } | Sort-Object LastWriteTime -Descending | Select-Object -First 1)[0]
            if($created){
                $targetFile = Join-Path $outPath $created.Name
                if(-not $created.FullName.Equals($targetFile,[StringComparison]::OrdinalIgnoreCase)){
                    Move-Item -LiteralPath $created.FullName -Destination $targetFile -Force
                    $created = Get-Item -LiteralPath $targetFile
                }
            }
        }
        if(-not $created){
            $details = ''
            if(Test-Path -LiteralPath $stdoutLog){ $details += (Get-Content -LiteralPath $stdoutLog -Raw -ErrorAction SilentlyContinue) }
            if(Test-Path -LiteralPath $stderrLog){ $details += [Environment]::NewLine + (Get-Content -LiteralPath $stderrLog -Raw -ErrorAction SilentlyContinue) }
            throw "IntuneWinAppUtil completed but no .intunewin file was found. Utility output: $details"
        }

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

function Find-InstalledUninstallEntry {
    if([string]::IsNullOrWhiteSpace($app.Text)){ return }
    $name=$app.Text.Trim()
    $roots=@('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    $allEntries=@(foreach($root in $roots){ Get-ItemProperty -Path $root -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName } })
    $matches=@($allEntries | Where-Object { [string]$_.DisplayName -eq $name })
    if(-not $matches.Count){ $matches=@($allEntries | Where-Object { [string]$_.DisplayName -like "*$name*" }) }
    # Match by application name only. Version differences must not block uninstall detection.
    $entry=$matches | Select-Object -First 1
    if($entry){
        $cmd=if($entry.QuietUninstallString){[string]$entry.QuietUninstallString}else{[string]$entry.UninstallString}
        if([string]::IsNullOrWhiteSpace($cmd)){$uninstallSwitchStatus.Text='Unable to find the uninstall string. Please add it manually.';$uninstallSwitchStatus.ForeColor=[System.Drawing.Color]::DarkOrange;return}
        $uninstallSwitchStatus.Text="Registry entry found: $($entry.DisplayName) $($entry.DisplayVersion)";$uninstallSwitchStatus.ForeColor=[System.Drawing.Color]::DarkGreen
        if($cmd -match '(?i)msiexec(?:\.exe)?\s+.*?\{[0-9a-f-]{36}\}'){
            $pc=[regex]::Match($cmd,'\{[0-9a-fA-F-]{36}\}').Value
            if($pc){$uMethod.SelectedItem='MSI ProductCode';$uFile.Text=$pc;$uParams.Text='/qn /norestart'}
        } else {
            if($cmd -match '^\s*"(?<exe>[^"]+)"\s*(?<args>.*)$'){$exe=$Matches.exe;$args=$Matches.args}
            else {$parts=[regex]::Match($cmd,'^\s*(?<exe>\S+)(?:\s+(?<args>.*))?$');$exe=$parts.Groups['exe'].Value;$args=$parts.Groups['args'].Value}
            if($exe){$uMethod.SelectedItem='Uninstall from different location';$uFile.Text=$exe}
            $uParams.Text=([string]$args).Trim()
        }
        $uninstallSwitchStatus.Text += ' — command populated. Verify the parameters before deployment.'
    } else {
        $uninstallSwitchStatus.Text='Unable to find the uninstall string. Please add it manually.'
        $uninstallSwitchStatus.ForeColor=[System.Drawing.Color]::DarkOrange
    }
}

# Context-aware action fields: hide unused inputs and expose browse/options only where applicable.
function Set-ActionComboItems($Combo,[string[]]$Items,[string]$Default) {
    $Combo.Items.Clear(); [void]$Combo.Items.AddRange($Items)
    if($Items -contains $Default){$Combo.SelectedItem=$Default}elseif($Combo.Items.Count -gt 0){$Combo.SelectedIndex=0}
}
function Update-ConfigurationActionFields {
    $kind=[string]$actionType.SelectedItem
    $actionSourceLabel.Text='Source / Key / Process / Command'
    $actionDestinationLabel.Text='Destination / Value / Parameters'
    $actionExtraLabel.Text='Additional value / type / timeout / message'
    $actionSourceLabel.Visible=$true; $actionSource.Visible=$true
    $actionDestinationLabel.Visible=$true; $actionDestination.Visible=$true
    $actionExtraLabel.Visible=$true; $actionExtra.Visible=$true
    $actionSourceBrowse.Visible=$false; $actionDestinationBrowse.Visible=$false
    $actionRegistryTypeLabel.Visible=$false; $actionRegistryType.Visible=$false
    foreach($c in @($actionOption1Label,$actionOption1,$actionOption2Label,$actionOption2,$actionOption3Label,$actionOption3,$actionOption4Label,$actionOption4)){$c.Visible=$false}
    $actionExtra.Width=350
    switch($kind){
        'Copy single file' {
            $actionSourceLabel.Text='Source file (relative to package Files)'; $actionDestinationLabel.Text='Destination folder (PSADT path)'
            $actionSourceBrowse.Visible=$true; $actionDestinationBrowse.Visible=$true
            $actionExtraLabel.Visible=$false; $actionExtra.Visible=$false
            if(-not $actionDestination.Text){$actionDestination.Text='$envProgramData\Vendor\App'}
        }
        'Copy files / folder recursively' {
            $actionSourceLabel.Text='Source folder/file (relative to package Files)'; $actionDestinationLabel.Text='Destination folder (PSADT path)'
            $actionSourceBrowse.Visible=$true; $actionDestinationBrowse.Visible=$true
            $actionExtraLabel.Visible=$false; $actionExtra.Visible=$false
            if(-not $actionDestination.Text){$actionDestination.Text='$envProgramData\Vendor\App'}
        }
        'Remove file' {$actionSourceLabel.Text='File path to remove';$actionDestinationLabel.Visible=$false;$actionDestination.Visible=$false;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false;$actionSourceBrowse.Visible=$true}
        'Remove folder' {$actionSourceLabel.Text='Folder path to remove';$actionDestinationLabel.Visible=$false;$actionDestination.Visible=$false;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false;$actionSourceBrowse.Visible=$true}
        'Set registry value' {$actionSourceLabel.Text='Registry key';$actionDestinationLabel.Text='Value name';$actionExtraLabel.Text='Value data';$actionRegistryTypeLabel.Visible=$true;$actionRegistryType.Visible=$true}
        'Remove registry key' {$actionSourceLabel.Text='Registry key';$actionDestinationLabel.Visible=$false;$actionDestination.Visible=$false;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false}
        'Import .reg file' {$actionSourceLabel.Text='.reg file (relative to package Files)';$actionDestinationLabel.Visible=$false;$actionDestination.Visible=$false;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false;$actionSourceBrowse.Visible=$true}
        'Create shortcut' {$actionSourceLabel.Text='Shortcut target executable';$actionDestinationLabel.Text='Shortcut (.lnk) destination';$actionExtraLabel.Text='Arguments / description (optional)';$actionSourceBrowse.Visible=$true;$actionDestinationBrowse.Visible=$true}
        'Install certificate' {$actionSourceLabel.Text='Certificate file (relative to package Files)';$actionDestinationLabel.Text='Certificate store';$actionSourceBrowse.Visible=$true;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false;$actionOption1Label.Text='Store';Set-ActionComboItems $actionOption1 @('TrustedPublisher','Root','CA','My') 'TrustedPublisher';$actionOption1Label.Visible=$true;$actionOption1.Visible=$true;$actionDestinationLabel.Visible=$false;$actionDestination.Visible=$false}
        'Install driver (PnPUtil)' {$actionSourceLabel.Text='Driver INF (relative to package Files)';$actionDestinationLabel.Text='Extra pnputil parameters (optional)';$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false;$actionSourceBrowse.Visible=$true}
        'Install driver (DPInst)' {$actionSourceLabel.Text='Driver folder / INF (relative to package Files)';$actionDestinationLabel.Text='DPInst parameters (default /SW /SA)';$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false;$actionSourceBrowse.Visible=$true;$actionDestinationBrowse.Visible=$false}
        'Set environment variable' {$actionSourceLabel.Text='Variable name';$actionDestinationLabel.Text='Variable value';$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false;$actionOption1Label.Text='Scope';Set-ActionComboItems $actionOption1 @('Machine','User') 'Machine';$actionOption1Label.Visible=$true;$actionOption1.Visible=$true}
        'Start service' {$actionSourceLabel.Text='Service name';$actionDestinationLabel.Visible=$false;$actionDestination.Visible=$false;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false}
        'Stop service' {$actionSourceLabel.Text='Service name';$actionDestinationLabel.Visible=$false;$actionDestination.Visible=$false;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false}
        'Create firewall rule' {$actionSourceLabel.Text='Firewall rule display name';$actionDestinationLabel.Text='Program path OR local port';$actionDestinationBrowse.Visible=$true;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false;$actionOption1Label.Text='Direction';Set-ActionComboItems $actionOption1 @('Inbound','Outbound') 'Inbound';$actionOption2Label.Text='Protocol';Set-ActionComboItems $actionOption2 @('TCP','UDP','Any') 'TCP';$actionOption3Label.Text='Action';Set-ActionComboItems $actionOption3 @('Allow','Block') 'Allow';$actionOption1Label.Visible=$true;$actionOption1.Visible=$true;$actionOption2Label.Visible=$true;$actionOption2.Visible=$true;$actionOption3Label.Visible=$true;$actionOption3.Visible=$true}
        'Remove firewall rule' {$actionSourceLabel.Text='Firewall rule display name';$actionDestinationLabel.Visible=$false;$actionDestination.Visible=$false;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false}
        'Set file/folder permissions' {$actionSourceLabel.Text='Target file/folder path';$actionDestinationLabel.Text='Account / group';$actionSourceBrowse.Visible=$true;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false;$actionOption1Label.Text='Rights';Set-ActionComboItems $actionOption1 @('FullControl','Modify','ReadAndExecute','Read','Write') 'Modify';$actionOption2Label.Text='Access';Set-ActionComboItems $actionOption2 @('Allow','Deny') 'Allow';$actionOption1Label.Visible=$true;$actionOption1.Visible=$true;$actionOption2Label.Visible=$true;$actionOption2.Visible=$true}
        'Set HKCU registry value for all users' {$actionSourceLabel.Text='HKCU registry key';$actionDestinationLabel.Text='Value name';$actionExtraLabel.Text='Value data';$actionRegistryTypeLabel.Visible=$true;$actionRegistryType.Visible=$true}
        'Copy file to all user profiles' {$actionSourceLabel.Text='Source file (relative to package Files)';$actionDestinationLabel.Text='Relative destination under each profile';$actionSourceBrowse.Visible=$true;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false}
        'Start-Sleep' {$actionSourceLabel.Text='Delay in seconds';$actionDestinationLabel.Visible=$false;$actionDestination.Visible=$false;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false}
        'Wait-Process' {$actionSourceLabel.Text='Process name (without .exe)';$actionDestinationLabel.Text='Timeout seconds (optional)';$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false}
        'Write-Log' {$actionSourceLabel.Text='Log message';$actionDestinationLabel.Visible=$false;$actionDestination.Visible=$false;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false}
        'Show restart prompt' {$actionSourceLabel.Visible=$false;$actionSource.Visible=$false;$actionDestinationLabel.Visible=$false;$actionDestination.Visible=$false;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false}
        'Custom PowerShell' {$actionSourceLabel.Text='PowerShell command';$actionDestinationLabel.Visible=$false;$actionDestination.Visible=$false;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false}
        'Detect and uninstall older version' {$actionSourceLabel.Text='Application display name (optional)';$actionDestinationLabel.Visible=$false;$actionDestination.Visible=$false;$actionExtraLabel.Visible=$false;$actionExtra.Visible=$false}
    }
}
$actionType.Add_SelectedValueChanged({
    $actionSource.Clear(); $actionSource.Tag=$null; $actionDestination.Clear(); $actionExtra.Clear()
    if($actionRegistryType){$actionRegistryType.SelectedItem='String'}
    Update-ConfigurationActionFields
})
Update-ConfigurationActionFields
# Compact path pickers for configuration actions. File selections inside the source root are stored relatively.
$actionSourceBrowse.Add_Click({
    $kind=[string]$actionType.SelectedItem
    if($kind -eq 'Remove folder' -or $kind -eq 'Set file/folder permissions' -or $kind -eq 'Copy files / folder recursively' -or $kind -eq 'Install driver (DPInst)') {
        $d=New-Object System.Windows.Forms.FolderBrowserDialog; $d.Description='Select target folder'
        try {
            if($actionSource.Text -and (Test-Path -LiteralPath $actionSource.Text -PathType Container)){$d.SelectedPath=$actionSource.Text}
            if($d.ShowDialog() -eq 'OK') {
                $selected=[IO.Path]::GetFullPath($d.SelectedPath); $base=''
                if($script:source -and (Test-Path -LiteralPath $script:source.Text -PathType Container)){$base=[IO.Path]::GetFullPath($script:source.Text).TrimEnd('\')}
                if($kind -in @('Copy files / folder recursively','Install driver (DPInst)')) {
                    if($base -and $selected.StartsWith($base+'\',[StringComparison]::OrdinalIgnoreCase)){$actionSource.Text=$selected.Substring($base.Length+1)}
                    else {$script:source.Text=[IO.Path]::GetDirectoryName($selected);$actionSource.Text=[IO.Path]::GetFileName($selected)}
                } else {$actionSource.Text=Convert-ToPortablePath $selected}
            }
        } finally {$d.Dispose()}
    } else {
        $d=New-Object System.Windows.Forms.OpenFileDialog
        if($kind -eq 'Remove file') { $d.DereferenceLinks=$false; $d.Filter='All files, including shortcuts (*.*)|*.*' }
        else { $d.Filter=if($kind -eq 'Import .reg file'){'Registry files (*.reg)|*.reg|All files (*.*)|*.*'} elseif($kind -eq 'Install certificate'){'Certificate files (*.cer;*.crt;*.pfx;*.p12)|*.cer;*.crt;*.pfx;*.p12|All files (*.*)|*.*'} elseif($kind -eq 'Install driver (PnPUtil)'){'Driver INF files (*.inf)|*.inf|All files (*.*)|*.*'} else {'All files (*.*)|*.*'} }
        if($script:source.Text -and (Test-Path -LiteralPath $script:source.Text -PathType Container)){$d.InitialDirectory=$script:source.Text}
        try {
            if($d.ShowDialog() -eq 'OK') {
                $selected=[IO.Path]::GetFullPath($d.FileName); $base=''
                if($script:source.Text -and (Test-Path -LiteralPath $script:source.Text -PathType Container)){$base=[IO.Path]::GetFullPath($script:source.Text).TrimEnd('\')}
                if($kind -eq 'Remove file') {
                    # Remove file must target the exact file selected in the picker,
                    # including .lnk files. Do not resolve shortcut contents or replace
                    # the selected file with the shortcut's target executable.
                    $actionSource.Text=$selected
                    $actionSource.Tag=$selected
                } elseif($kind -eq 'Create shortcut') {
                    # This field is the shortcut target executable, not the .lnk output.
                    $actionSource.Text=Convert-ToPortablePath $selected
                } elseif($base -and $selected.StartsWith($base+'\',[StringComparison]::OrdinalIgnoreCase)) {
                    $actionSource.Text=$selected.Substring($base.Length+1)
                } else {
                    # Keep package-file actions relative to the selected source root.
                    $script:source.Text=[IO.Path]::GetDirectoryName($selected)
                    $actionSource.Text=[IO.Path]::GetFileName($selected)
                }
            }
        } finally {$d.Dispose()}
    }
})
$actionDestinationBrowse.Add_Click({
    $kind=[string]$actionType.SelectedItem
    if($kind -eq 'Create shortcut') {
        $d=New-Object System.Windows.Forms.SaveFileDialog; $d.Filter='Shortcut files (*.lnk)|*.lnk'; $d.DefaultExt='lnk'; $d.AddExtension=$true
        if($actionDestination.Text){$d.FileName=$actionDestination.Text}
        try {if($d.ShowDialog() -eq 'OK'){$actionDestination.Text=Convert-ToPortablePath $d.FileName}} finally {$d.Dispose()}
    } elseif($kind -eq 'Create firewall rule') {
        $d=New-Object System.Windows.Forms.OpenFileDialog; $d.Filter='Executable files (*.exe)|*.exe|All files (*.*)|*.*'
        try {if($d.ShowDialog() -eq 'OK'){$actionDestination.Text=Convert-ToPortablePath $d.FileName}} finally {$d.Dispose()}
    } else {
        $d=New-Object System.Windows.Forms.FolderBrowserDialog; $d.Description='Select destination folder'; $d.ShowNewFolderButton=$true
        try {if($actionDestination.Text -and (Test-Path -LiteralPath $actionDestination.Text -PathType Container)){$d.SelectedPath=$actionDestination.Text};if($d.ShowDialog() -eq 'OK'){$actionDestination.Text=Convert-ToPortablePath $d.SelectedPath}} finally {$d.Dispose()}
    }
})

# ---------------- Configuration Builder behavior ----------------
function Get-PhaseEditor([string]$Phase) {
    switch($Phase){
        'Pre-Installation' { return $preInstall }
        'Installation' { return $installSteps }
        'Post-Installation' { return $postInstall }
        'Pre-Uninstallation' { return $preUninstall }
        'Uninstallation' { return $uninstallSteps }
        'Post-Uninstallation' { return $postUninstall }
        default { throw "Unknown execution phase: $Phase" }
    }
}
function Add-CommandToPhase([string]$Phase,[string]$Command) {
    if([string]::IsNullOrWhiteSpace($Command)){ return }
    $editor=Get-PhaseEditor $Phase
    if(-not [string]::IsNullOrWhiteSpace($editor.Text)){$editor.AppendText([Environment]::NewLine+[Environment]::NewLine)}
    $editor.AppendText($Command.Trim())
}
function Convert-ToPortablePath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    $p = $Path.Trim()
    # Use literal path-prefix comparisons, not regex. Windows paths contain
    # backslashes (for example \ProgramData), which must not be interpreted
    # as regex escapes such as \p.
    $p = $p.Replace('/', '\')
    $rules = @(
        @{ Prefix='ProgramData\Microsoft\Windows\Start Menu\Programs'; Replacement='$envCommonStartMenuPrograms' },
        @{ Prefix='Users\Public\Desktop'; Replacement='$envCommonDesktop' },
        @{ Prefix='Users\Public\Documents'; Replacement='$envPublic\Documents' },
        @{ Prefix='Users\Public'; Replacement='$envPublic' },
        @{ Prefix='Program Files (x86)'; Replacement='$envProgramFilesX86' },
        @{ Prefix='Program Files'; Replacement='$envProgramFiles' },
        @{ Prefix='ProgramData'; Replacement='$envProgramData' },
        @{ Prefix='Windows'; Replacement='$envWinDir' }
    )
    if ($p -match '^[A-Za-z]:\\') {
        $root = $p.Substring(3)
        foreach ($rule in $rules) {
            if ($root.StartsWith($rule.Prefix, [StringComparison]::OrdinalIgnoreCase)) {
                if ($root.Length -eq $rule.Prefix.Length -or $root[$rule.Prefix.Length] -eq '\') {
                    return $rule.Replacement + $root.Substring($rule.Prefix.Length)
                }
            }
        }
    }
    return $p
}

function Get-ConfigurationCommand {
    $src=$actionSource.Text.Trim(); $dst=Convert-ToPortablePath $actionDestination.Text.Trim(); $extra=$actionExtra.Text.Trim()
    $qsrc=$src.Replace('"','`"'); $qdst=$dst.Replace('"','`"')
    switch([string]$actionType.SelectedItem){
        'Copy single file' { if(!$src -or !$dst){throw 'Enter a source file and destination folder.'}; return ('Copy-File -Path "$dirFiles\{0}" -Destination "{1}"' -f $qsrc,$qdst) }
        'Copy files / folder recursively' { if(!$src -or !$dst){throw 'Enter a source file/folder and destination folder.'}; return ('Copy-File -Path "$dirFiles\{0}" -Destination "{1}" -Recurse' -f $qsrc,$qdst) }
        'Remove file' { if(!$src){throw 'Enter a file path.'}; $literalPath=$src; if($actionSource.Tag -and [string]::Equals([string]$actionSource.Tag,$src,[StringComparison]::OrdinalIgnoreCase)){$literalPath=[string]$actionSource.Tag}; $portable=Convert-ToPortablePath $literalPath; return ('Remove-File -Path "{0}"' -f $portable.Replace('"','`"')) }
        'Remove folder' { if(!$src){throw 'Enter a folder path.'}; $portable=Convert-ToPortablePath $src; return ('Remove-Folder -Path "{0}"' -f $portable.Replace('"','`"')) }
        'Set registry value' { if(!$src -or !$dst){throw 'Source field is registry key; destination is value name; enter data in Value data.'}; $v=$extra.Replace("'","''"); $rt=[string]$actionRegistryType.SelectedItem; if($rt -eq 'DWord' -and $v -notmatch '^-?\d+$'){throw 'For DWord, enter a whole-number value such as 0 or 1.'}; return ("Set-RegistryKey -Key '{0}' -Name '{1}' -Value '{2}' -Type {3}" -f $src.Replace("'","''"),$dst.Replace("'","''"),$v,$rt) }
        'Remove registry key' { if(!$src){throw 'Enter a registry key.'}; return ("Remove-RegistryKey -Key '{0}' -Recurse" -f $src.Replace("'","''")) }
        'Import .reg file' { if(!$src){throw 'Enter a .reg filename relative to the source folder.'}; return ('Execute-Process -Path "reg.exe" -Parameters "IMPORT `"$dirFiles\{0}`""' -f $qsrc) }
        'Create shortcut' { if(!$src -or !$dst){throw 'Source field is shortcut target; destination is full .lnk path.'}; return ('New-Shortcut -Path "{0}" -TargetPath "{1}"' -f (Convert-ToPortablePath $dst),(Convert-ToPortablePath $src)) }
        'Install certificate' { if(!$src){throw 'Enter certificate filename relative to source folder.'}; $store=[string]$actionOption1.SelectedItem; if($psadtVersion.SelectedItem -eq '4.x'){return ('Start-ADTProcess -FilePath "$env:SystemRoot\System32\certutil.exe" -ArgumentList ''-f -addstore -enterprise {0} `"$dirFiles\{1}`"'' -Wait' -f $store,$qsrc)}; return ('Execute-Process -Path "certutil.exe" -Parameters "-f -addstore -enterprise {0} `"$dirFiles\{1}`""' -f $store,$qsrc) }
        'Install driver (PnPUtil)' { if(!$src){throw 'Enter the driver INF path relative to source folder.'}; $driverParams=if($dst){$dst}else{'/subdirs /install'}; if($psadtVersion.SelectedItem -eq '4.x'){return ('Start-ADTProcess -FilePath "$env:SystemRoot\System32\pnputil.exe" -ArgumentList ''/add-driver `"$dirFiles\{0}`" {1}'' -Wait' -f $qsrc,$driverParams)}; return ('Execute-Process -Path "pnputil.exe" -Parameters "/add-driver `"$dirFiles\{0}`" {1}"' -f $qsrc,$driverParams) }
        'Install driver (DPInst)' { if(!$src){throw 'Enter the driver folder/INF path relative to source folder.'}; if($psadtVersion.SelectedItem -eq '4.x'){return ('Start-ADTProcess -FilePath "$dirFiles\DPInst.exe" -ArgumentList ''{0}'' -Wait' -f $(if($dst){$dst}else{'/SW /SA'}))}; return ('Execute-Process -Path "$dirFiles\DPInst.exe" -Parameters "{0}"' -f $(if($dst){$dst}else{'/SW /SA'})) }
        'Set environment variable' { if(!$src -or !$dst){throw 'Enter environment variable name in Source and value in Destination.'}; $scope=[string]$actionOption1.SelectedItem; return ('[Environment]::SetEnvironmentVariable("{0}", "{1}", "{2}")' -f $src,$dst,$scope) }
        'Start service' { if(!$src){throw 'Enter service name.'}; return ('Start-Service -Name "{0}" -ErrorAction SilentlyContinue' -f $qsrc) }
        'Stop service' { if(!$src){throw 'Enter service name.'}; return ('Stop-Service -Name "{0}" -Force -ErrorAction SilentlyContinue' -f $qsrc) }
        'Create firewall rule' {
            if(!$src -or !$dst){throw 'Enter a firewall rule name and a program path or local port.'}
            $direction=[string]$actionOption1.SelectedItem; $protocol=[string]$actionOption2.SelectedItem; $fwAction=[string]$actionOption3.SelectedItem
            if($direction -notin @('Inbound','Outbound') -or $protocol -notin @('TCP','UDP','Any') -or $fwAction -notin @('Allow','Block')){throw 'Select valid firewall Direction, Protocol, and Action values.'}
            if($dst -match '^\d+(?:\s*[-,]\s*\d+)*$'){
                $ports=($dst -replace '\s','')
                foreach($portToken in ($ports -split '[,-]')){if([int]$portToken -lt 1 -or [int]$portToken -gt 65535){throw 'Firewall ports must be between 1 and 65535.'}}
                if($ports -match '^(\d+)-(\d+)$' -and [int]$Matches[1] -gt [int]$Matches[2]){throw 'Firewall port range start must be less than or equal to its end.'}
                return ('New-NetFirewallRule -DisplayName "{0}" -Direction {1} -Action {2} -Protocol {3} -LocalPort "{4}"' -f $qsrc,$direction,$fwAction,$protocol,$ports)
            }
            $programPath=Convert-ToPortablePath $dst
            return ('New-NetFirewallRule -DisplayName "{0}" -Direction {1} -Action {2} -Protocol {3} -Program "{4}"' -f $qsrc,$direction,$fwAction,$protocol,$programPath.Replace('"','`"'))
        }
        'Remove firewall rule' { if(!$src){throw 'Enter the firewall rule display name.'}; $nameEsc=$src.Replace("'","''"); return "`$rules = @(Get-NetFirewallRule -DisplayName '$nameEsc' -ErrorAction SilentlyContinue); if (`$rules.Count -gt 0) { `$rules | Remove-NetFirewallRule -ErrorAction Stop } else { Write-Log -Message 'Firewall rule not found: $nameEsc' }" }
        'Set file/folder permissions' {
            if(!$src -or !$dst){throw 'Enter a target path and account/group.'}
            $rights=[string]$actionOption1.SelectedItem; $access=[string]$actionOption2.SelectedItem
            if($rights -notin @('FullControl','Modify','ReadAndExecute','Read','Write') -or $access -notin @('Allow','Deny')){throw 'Rights must be FullControl, Modify, ReadAndExecute, Read, or Write; access must be Allow or Deny.'}
            $path=Convert-ToPortablePath $src
            return ('$aclPath = "{0}"; $acl = Get-Acl -LiteralPath $aclPath; $rule = New-Object System.Security.AccessControl.FileSystemAccessRule("{1}","{2}","ContainerInherit,ObjectInherit","None","{3}"); $acl.SetAccessRule($rule); Set-Acl -LiteralPath $aclPath -AclObject $acl' -f $path.Replace('"','`"'),$dst.Replace('"','`"'),$rights,$access)
        }
        'Set HKCU registry value for all users' {
            if(!$src -or !$dst){throw 'Enter the HKCU key and value name.'}
            if([string]::IsNullOrWhiteSpace($extra)){throw 'Enter the registry value data.'}
            $data=$extra.Trim();$regType=[string]$actionRegistryType.SelectedItem
            if($regType -notin @('String','DWord')){throw 'Select String or DWord as the registry value type.'}
            if($regType -eq 'DWord' -and $data -notmatch '^-?\d+$'){throw 'For DWord, enter a whole-number value such as 0 or 1.'}
            if($regType -eq 'DWord'){$data=[int]$data}
            $key=$src.Replace("'","''");$valueName=$dst.Replace("'","''");$dataEsc=$data.Replace("'","''")
            return @"
[scriptblock]`$HKCURegistrySettings = {
    Set-RegistryKey -Key '$key' -Name '$valueName' -Value '$dataEsc' -Type $regType -ContinueOnError:`$True -SID `$UserProfile.SID
}
Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings `$HKCURegistrySettings
"@
        }
        'Copy file to all user profiles' {
            if(!$src -or !$dst){throw 'Enter a source file and a relative destination path under each user profile.'}
            if([IO.Path]::IsPathRooted($dst) -or $dst -match '(^|[\\/])\.\.([\\/]|$)'){throw 'Destination must be relative to each user profile and cannot contain .. segments.'}
            $sourceRel=$src.Replace('"','`"');$destRel=$dst.Replace('"','`"')
            return @"
`$ProfilePaths = Get-UserProfiles | Select-Object -ExpandProperty 'ProfilePath'
ForEach (`$Profile in `$ProfilePaths) {
    Copy-File -Path "`$dirFiles\$sourceRel" -Destination (Join-Path `$Profile "$destRel")
}
"@
        }
        'Start-Sleep' { $n=if($src){$src}else{'5'}; if($n -notmatch '^\d+$'){throw 'Enter delay seconds in Source.'}; return "Start-Sleep -Seconds $n" }
        'Wait-Process' { if(!$src){throw 'Enter process name (without .exe).' }; if($dst -match '^\d+$'){return ('Wait-Process -Name "{0}" -Timeout {1} -ErrorAction SilentlyContinue' -f $qsrc,$dst)}; return ('Wait-Process -Name "{0}" -ErrorAction SilentlyContinue' -f $qsrc) }
        'Write-Log' { $msg=if($src){$src}else{'Configuration action executed'}; if($psadtVersion.SelectedItem -eq '4.x'){return ("Write-ADTLogEntry -Message '{0}'" -f $msg.Replace("'","''"))}; return ("Write-Log -Message '{0}'" -f $msg.Replace("'","''")) }
        'Show restart prompt' { if($psadtVersion.SelectedItem -eq '4.x'){return 'Show-ADTInstallationRestartPrompt'}; return 'Show-InstallationRestartPrompt' }
        'Custom PowerShell' { if(!$src){throw 'Enter the PowerShell command in Source / Command.'}; return $src }
        'Detect and uninstall older version' { if(!$src){$matchName=$app.Text.Trim()}else{$matchName=$src}; if(!$matchName){throw 'Enter the installed application display name or leave Source blank to use the package application name.'}; $escaped=$matchName.Replace("'","''"); $targetVersion=$ver.Text.Trim().Replace("'","''"); return @"
`$targetName = '$escaped'
`$targetVersion = '$targetVersion'
`$uninstallRoots = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
`$installedMatches = foreach (`$root in `$uninstallRoots) { Get-ItemProperty -Path `$root -ErrorAction SilentlyContinue | Where-Object { `$_.DisplayName -and `$_.DisplayName -like "*`$targetName*" -and `$_.DisplayVersion } }
foreach (`$oldApp in `$installedMatches) {
    try { if ([version]`$oldApp.DisplayVersion -ge [version]`$targetVersion) { continue } } catch { continue }
    `$uninstallCmd = if (`$oldApp.QuietUninstallString) { `$oldApp.QuietUninstallString } else { `$oldApp.UninstallString }
    if (-not `$uninstallCmd) { Write-Log -Message "No uninstall command found for `$(`$oldApp.DisplayName)"; continue }
    if (`$uninstallCmd -match '(?i)msiexec(?:\.exe)?\s+.*?\{[0-9a-f-]{36}\}') {
        `$productCode = [regex]::Match(`$uninstallCmd,'\{[0-9a-fA-F-]{36}\}').Value
        if (`$productCode) { `$p = Start-Process -FilePath "`$env:SystemRoot\System32\msiexec.exe" -ArgumentList "/x `$productCode /qn /norestart" -Wait -PassThru; if (`$p.ExitCode -notin @(0,1605,1614,3010)) { throw "Old MSI uninstall failed with exit code `$(`$p.ExitCode)" } }
    } else {
        `$exe = [regex]::Match(`$uninstallCmd,'^\s*"(?<p>[^"]+)"|^\s*(?<p>\S+)').Groups['p'].Value
        `$argsText = `$uninstallCmd.Substring(`$exe.Length).Trim().Trim('"')
        if (`$oldApp.QuietUninstallString) { `$p = Start-Process -FilePath `$exe -ArgumentList `$argsText -Wait -PassThru; if (`$p.ExitCode -notin @(0,3010)) { throw "Old EXE uninstall failed with exit code `$(`$p.ExitCode)" } } else { Write-Output "Uninstall for `$(`$oldApp.DisplayName) may not be silent; review registry command: `$uninstallCmd" }
    }
}
"@ }
    }
}
$script:ConfigurationActionRecords = New-Object System.Collections.Generic.List[object]
$previewActionButton.Text='Preview All Actions'
$previewActionButton.Add_Click({
    try {
        $draft=Get-ConfigurationCommand
        $all=@()
        foreach($record in $script:ConfigurationActionRecords){$all += ('# [{0}] {1}' -f $record.Phase,$record.Action);$all += $record.Command}
        if($draft){
            $all += ('# [{0}] {1} (current selection; not yet added)' -f [string]$actionPhase.SelectedItem,[string]$actionType.SelectedItem)
            $all += $draft
        }
        $actionPreview.Text=$all -join ([Environment]::NewLine+[Environment]::NewLine)
    } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Configuration Builder') }
})
$addActionButton.Text='Add Current Action'
$addActionButton.Add_Click({
    try {
        $command=Get-ConfigurationCommand
        $phase=[string]$actionPhase.SelectedItem
        $actionName=[string]$actionType.SelectedItem
        $editor=Get-PhaseEditor $phase
        if($script:ConfigurationActionRecords | Where-Object { $_.Phase -eq $phase -and $_.Command -eq $command }){
            [System.Windows.Forms.MessageBox]::Show('This exact action is already in the selected phase. It was not added again.','Configuration Builder')
            return
        }
        if($editor.Text.Contains($command.Trim())){
            [System.Windows.Forms.MessageBox]::Show('This command already exists in the selected phase editor. It was not added again.','Configuration Builder')
            return
        }
        Add-CommandToPhase -Phase $phase -Command $command
        $script:ConfigurationActionRecords.Add([pscustomobject]@{Action=$actionName;Phase=$phase;Command=$command})
        $tabs.SelectedTab= if($phase -match 'Uninstallation'){ $uninstall } else { $install }
    } catch { [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Configuration Builder') }
})
$clearConfigButton.Add_Click({ if([System.Windows.Forms.MessageBox]::Show('Clear the six phase editors? This cannot be undone.','Clear Configuration',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning) -eq [System.Windows.Forms.DialogResult]::Yes){ $preInstall.Clear();$installSteps.Clear();$postInstall.Clear();$preUninstall.Clear();$uninstallSteps.Clear();$postUninstall.Clear();$actionPreview.Clear();$script:ConfigurationActionRecords.Clear() } })
$script:AdditionalInstallers = New-Object System.Collections.Generic.List[object]
$script:EditingAdditionalInstallerIndex = -1
$addInstallerButton.Add_Click({
    $dlg=New-Object System.Windows.Forms.Form; $dlg.Text='Add / Edit Additional Installer'; $dlg.Size=New-Object System.Drawing.Size(760,580); $dlg.StartPosition='CenterParent'; $dlg.Font=$form.Font
    New-Label $dlg 'Installer file (relative to source folder)' 18 16
    $fileBox=New-TextBox $dlg 18 40 590
    $browse=New-Object System.Windows.Forms.Button; $browse.Text='Browse...';$browse.Location=New-Object System.Drawing.Point(620,40);$browse.Size=New-Object System.Drawing.Size(100,28);$dlg.Controls.Add($browse)
    $browse.Add_Click({
        $fd=New-Object System.Windows.Forms.OpenFileDialog
        $fd.Filter='Installer files (*.exe;*.msi)|*.exe;*.msi|All files (*.*)|*.*'
        if($source.Text -and (Test-Path $source.Text)){$fd.InitialDirectory=$source.Text}
        if($fd.ShowDialog() -eq 'OK'){
            $full=[IO.Path]::GetFullPath($fd.FileName);$base=''
            if($source.Text -and (Test-Path $source.Text -PathType Container)){$base=[IO.Path]::GetFullPath($source.Text).TrimEnd('\')}
            if($base -and $full.StartsWith($base+'\',[StringComparison]::OrdinalIgnoreCase)){
                $fileBox.Text=$full.Substring($base.Length+1)
                $ext=[IO.Path]::GetExtension($full)
                if($ext -ieq '.msi'){
                    $paramsBox.Text='/qn /norestart';$unParamsBox.Text='/qn /norestart';$additionalMst.Enabled=$true
                    try {
                        $wi=New-Object -ComObject WindowsInstaller.Installer;$db=$wi.OpenDatabase($full,0);$props=@{}
                        foreach($pn in @('ProductName','ProductVersion','Manufacturer','ProductCode')){$vw=$db.OpenView("SELECT `Value` FROM Property WHERE `Property`='$pn'");$vw.Execute();$r=$vw.Fetch();if($r){$props[$pn]=[string]$r.StringData(1)};$vw.Close()}
                        $additionalMetadata.Text=('MSI: {0} | Version: {1} | Vendor: {2} | ProductCode: {3}' -f $props.ProductName,$props.ProductVersion,$props.Manufacturer,$props.ProductCode)
                        $switchStatus.Text='MSI detected: /qn /norestart applied. Silent Switch Finder is not required for MSI.'
                    } catch {$additionalMetadata.Text='MSI selected; metadata read failed. Enter values manually.'}
                } else {
                    $additionalMst.Checked=$false;$additionalMst.Enabled=$false
                    $fi=[Diagnostics.FileVersionInfo]::GetVersionInfo($full)
                    $additionalMetadata.Text=('EXE: {0} | Version: {1} | Vendor: {2}' -f [IO.Path]::GetFileNameWithoutExtension($full),$fi.FileVersion,$fi.CompanyName)
                    Find-AndPopulateSilentSwitch -InstallerPath $full -ParameterBox $paramsBox -StatusLabel $switchStatus -ContextName 'Additional Installer'
                }
            } else {[System.Windows.Forms.MessageBox]::Show('Choose an installer inside the selected source folder so it can be included in the package.')}
            # Uninstall detection is attempted in the background, but uninstall is optional for additional installers.
            # A missing string is shown in the status label only; it must not block installer selection or save.
            if($fileBox.Text.Trim()){ $detectUninstallButton.PerformClick() }
        }
        $fd.Dispose()
    })
    New-Label $dlg 'Install phase' 18 138; $phaseBox=New-Combo $dlg 18 162 220 @('Pre-Installation','Installation','Post-Installation')
    New-Label $dlg 'Install parameters' 260 138; $paramsBox=New-TextBox $dlg 260 162 460; $paramsBox.Text='/quiet /norestart'
    New-Label $dlg 'Detection condition' 18 198; $detectBox=New-Combo $dlg 18 222 300 @('Always run','Run only if missing','Custom detection (manual)')
    New-Label $dlg 'Uninstall command (optional; parameters included)' 18 258; $uninstallCmdBox=New-TextBox $dlg 18 282 700 28
    New-Label $dlg 'Uninstall phase' 18 318; $unPhaseBox=New-Combo $dlg 18 342 220 @('Pre-Uninstallation','Uninstallation','Post-Uninstallation')
    New-Label $dlg 'Uninstall parameters' 260 318; $unParamsBox=New-TextBox $dlg 260 342 460 28
    $detectUninstallButton=New-Object System.Windows.Forms.Button;$detectUninstallButton.Text='Detect Uninstall String';$detectUninstallButton.Location=New-Object System.Drawing.Point(18,486);$detectUninstallButton.Size=New-Object System.Drawing.Size(180,30);$dlg.Controls.Add($detectUninstallButton)
    $uninstallDetectStatus=New-Label $dlg 'Uninstall detection will use the installer metadata/application name.' 210 490;$uninstallDetectStatus.Size=New-Object System.Drawing.Size(510,24);$uninstallDetectStatus.ForeColor=[System.Drawing.Color]::DimGray
    $additionalMst=New-Object System.Windows.Forms.CheckBox
    $additionalMst.Text='Generate Standard MST (MSI only)'
    $additionalMst.Location=New-Object System.Drawing.Point(18,378);$additionalMst.AutoSize=$true;$additionalMst.Enabled=$false;$additionalMst.Checked=$false
    $dlg.Controls.Add($additionalMst)
    $switchStatus=New-Label $dlg 'Select an EXE to run Silent Switch Finder, or an MSI for /qn /norestart and metadata.' 18 404; $switchStatus.Size=New-Object System.Drawing.Size(700,32);$switchStatus.ForeColor=[System.Drawing.Color]::DimGray
    $detectSwitch=New-Object System.Windows.Forms.Button;$detectSwitch.Text='Detect Silent Switches';$detectSwitch.Location=New-Object System.Drawing.Point(18,440);$detectSwitch.Size=New-Object System.Drawing.Size(170,32);$dlg.Controls.Add($detectSwitch)
    $detectSwitch.Add_Click({
        $f=Join-Path $source.Text $fileBox.Text
        if(Test-Path -LiteralPath $f -PathType Leaf){
            if([IO.Path]::GetExtension($f) -ieq '.exe'){
                $additionalMst.Checked=$false;$additionalMst.Enabled=$false
                Find-AndPopulateSilentSwitch -InstallerPath $f -ParameterBox $paramsBox -StatusLabel $switchStatus -ContextName 'Additional Installer'
            } else {
                $paramsBox.Text='/qn /norestart';$unParamsBox.Text='/qn /norestart'
                $additionalMst.Enabled=$true
                $switchStatus.Text='MSI detected: /qn /norestart applied. Enable Standard MST if needed.'
            }
        } else {[System.Windows.Forms.MessageBox]::Show('Select an installer located inside the source folder.')}
    })
    $save=New-Object System.Windows.Forms.Button;$save.Text='Save Installer';$save.Location=New-Object System.Drawing.Point(540,520);$save.Size=New-Object System.Drawing.Size(180,34);$dlg.Controls.Add($save)
    $detectUninstallButton.Add_Click({
        # Detect uninstall information for both MSI and EXE installers.
        # MSI: ProductCode is the strongest key. EXE: use product metadata, file name,
        # publisher and registry uninstall entries; an EXE cannot provide a registry
        # uninstall string if the product has never been installed on this machine.
        $displayName='';$publisher='';$productCode=''
        if($additionalMetadata.Text -match '(?i)MSI:\s*(.*?)\s*\| Version:'){$displayName=$Matches[1].Trim()}
        elseif($additionalMetadata.Text -match '(?i)EXE:\s*(.*?)\s*\| Version:'){$displayName=$Matches[1].Trim()}
        if($additionalMetadata.Text -match '(?i)Vendor:\s*(.*?)(?:\s*\|\s*ProductCode:|$)'){$publisher=$Matches[1].Trim()}
        if($additionalMetadata.Text -match '(?i)ProductCode:\s*(\{[0-9a-f-]{36}\})'){$productCode=$Matches[1]}
        $installerFullPath=Join-Path $source.Text $fileBox.Text
        $installerExt=[IO.Path]::GetExtension($fileBox.Text)
        if(-not $productCode -and $installerExt -ieq '.msi'){
            try {
                $wiDetect=New-Object -ComObject WindowsInstaller.Installer
                $dbDetect=$wiDetect.OpenDatabase($installerFullPath,0)
                $viewDetect=$dbDetect.OpenView("SELECT `Value` FROM Property WHERE `Property`='ProductCode'")
                $viewDetect.Execute();$rowDetect=$viewDetect.Fetch()
                if($rowDetect){$productCode=[string]$rowDetect.StringData(1)}
                $viewDetect.Close()
            } catch {}
        }
        if(-not $displayName){$displayName=[IO.Path]::GetFileNameWithoutExtension($fileBox.Text)}
        $fileBase=[IO.Path]::GetFileNameWithoutExtension($fileBox.Text)
        $roots=@('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
        $allUninstallEntries=@(foreach($root in $roots){Get-ItemProperty -Path $root -ErrorAction SilentlyContinue | Where-Object {$_.DisplayName -or $_.UninstallString -or $_.QuietUninstallString}})
        $found=$null
        if($productCode){
            $found=$allUninstallEntries | Where-Object { ([string]$_.PSChildName -ieq $productCode) -or ([string]$_.UninstallString -match [regex]::Escape($productCode)) -or ([string]$_.QuietUninstallString -match [regex]::Escape($productCode)) } | Select-Object -First 1
        }
        # Try exact product name first, then filename/product-name overlap. Do not
        # require DisplayVersion: version differences must not block detection.
        if(-not $found -and $displayName){$found=$allUninstallEntries | Where-Object { [string]$_.DisplayName -ieq $displayName } | Select-Object -First 1}
        if(-not $found -and $displayName -and $displayName.Length -ge 4){$found=$allUninstallEntries | Where-Object { ([string]$_.DisplayName -like "*$displayName*") -or ($displayName -like "*$([string]$_.DisplayName)*" -and ([string]$_.DisplayName).Length -ge 4) } | Select-Object -First 1}
        if(-not $found -and $fileBase -and $fileBase.Length -ge 4){$found=$allUninstallEntries | Where-Object { [string]$_.DisplayName -like "*$fileBase*" } | Select-Object -First 1}
        if(-not $found -and $publisher -and $publisher.Length -ge 3 -and $fileBase.Length -ge 4){$found=$allUninstallEntries | Where-Object { ([string]$_.Publisher -like "*$publisher*") -and ([string]$_.DisplayName -like "*$fileBase*") } | Select-Object -First 1}
        if($found){
            $cmd=if($found.QuietUninstallString){[string]$found.QuietUninstallString}else{[string]$found.UninstallString}
            if([string]::IsNullOrWhiteSpace($cmd)){
                $uninstallDetectStatus.Text="Found '$($found.DisplayName)', but its registry entry has no uninstall string. Enter one manually only if needed.";$uninstallDetectStatus.ForeColor=[System.Drawing.Color]::DarkOrange
                return
            }
            if($cmd -match '(?i)msiexec(?:\.exe)?\s+.*?\{[0-9a-f-]{36}\}'){
                $uninstallCmdBox.Text=[regex]::Match($cmd,'\{[0-9a-fA-F-]{36}\}').Value;$unParamsBox.Text='/qn /norestart'
            } elseif($cmd -match '^\s*"(?<exe>[^"]+)"\s*(?<args>.*)$'){
                $uninstallCmdBox.Text=$Matches.exe;$unParamsBox.Text=([string]$Matches.args).Trim()
            } else {
                $parts=[regex]::Match($cmd,'^\s*(?<exe>\S+)(?:\s+(?<args>.*))?$')
                $uninstallCmdBox.Text=$parts.Groups['exe'].Value;$unParamsBox.Text=$parts.Groups['args'].Value.Trim()
            }
            $uninstallDetectStatus.Text="Uninstall string detected for '$($found.DisplayName)'. Review the command and parameters.";$uninstallDetectStatus.ForeColor=[System.Drawing.Color]::DarkGreen
        } elseif($productCode -and $installerExt -ieq '.msi') {
            $uninstallCmdBox.Text=$productCode;$unParamsBox.Text='/qn /norestart'
            $uninstallDetectStatus.Text='No matching registry entry found; MSI ProductCode was read from installer metadata.';$uninstallDetectStatus.ForeColor=[System.Drawing.Color]::DarkOrange
        } else {
            $uninstallDetectStatus.Text='Unable to find an uninstall string for this EXE/MSI. If uninstallation is required, enter the uninstall command and parameters manually; otherwise leave them blank.'
            $uninstallDetectStatus.ForeColor=[System.Drawing.Color]::DarkOrange
        }
    })
    $save.Add_Click({
        if(!$fileBox.Text.Trim()){[System.Windows.Forms.MessageBox]::Show('Select an installer.');return}
        $file=$fileBox.Text.Trim();$ext=[IO.Path]::GetExtension($file).ToLowerInvariant();$phase=[string]$phaseBox.SelectedItem;$unPhase=[string]$unPhaseBox.SelectedItem
        $selectedParams=if($paramsBox.Text.Trim()){$paramsBox.Text.Trim()}else{'/qn /norestart'}
        $transformName=''
        if($ext -eq '.msi' -and $additionalMst.Checked){
            $componentPathForMst=Join-Path $source.Text $file
            $mstScript=Join-Path $PSScriptRoot 'Create-StandardMST.ps1'
            if(-not (Test-Path -LiteralPath $componentPathForMst -PathType Leaf)){[System.Windows.Forms.MessageBox]::Show('The selected MSI must be inside the source folder.');return}
            if(-not (Test-Path -LiteralPath $mstScript -PathType Leaf)){[System.Windows.Forms.MessageBox]::Show('Create-StandardMST.ps1 was not found.');return}
            try {
                $wi2=New-Object -ComObject WindowsInstaller.Installer;$db2=$wi2.OpenDatabase($componentPathForMst,0);$pv=@{}
                foreach($pn in @('ProductName','ProductVersion','Manufacturer')){$vw2=$db2.OpenView("SELECT `Value` FROM Property WHERE `Property`='$pn'");$vw2.Execute();$rr=$vw2.Fetch();if($rr){$pv[$pn]=[string]$rr.StringData(1)};$vw2.Close()}
                $mstVendor=if($pv.Manufacturer){$pv.Manufacturer -replace '\s+','-'}else{'Vendor'}
                $mstApp=if($pv.ProductName){$pv.ProductName -replace '\s+','-'}else{[IO.Path]::GetFileNameWithoutExtension($file) -replace '\s+','-'}
                $mstVer=if($pv.ProductVersion){$pv.ProductVersion}else{'1.0'}
                & $mstScript -MsiPath $componentPathForMst -Vendor $mstVendor -AppName $mstApp -Version $mstVer -Architecture ([string]$arch.SelectedItem)
                if($LASTEXITCODE -and $LASTEXITCODE -ne 0){throw 'Standard MST generator returned a non-zero exit code.'}
                $transformName='{0}_{1}_{2}_P1_{3}.mst' -f $mstVendor,$mstApp,$mstVer,$arch.SelectedItem
                if(-not (Test-Path -LiteralPath (Join-Path $source.Text $transformName))){throw "Expected transform not found: $transformName"}
            } catch {[System.Windows.Forms.MessageBox]::Show(('Could not generate MST: '+$_.Exception.Message),'Additional Installer');return}
        }
        if($psadtVersion.SelectedItem -eq '4.x'){
            if($ext -eq '.msi'){$msiArgs='/i `"$dirFiles\{0}`" {1}' -f $file,$selectedParams;if($transformName){$msiArgs+=' TRANSFORMS=`"$dirFiles\'+$transformName+'`"'};$installCmd='Start-ADTProcess -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList ''{0}'' -WindowStyle Hidden -Wait' -f $msiArgs}
            else{$safeParam=$selectedParams.Replace("'","''");$installCmd='Start-ADTProcess -FilePath "$dirFiles\{0}" -ArgumentList ''{1}'' -WindowStyle Hidden -Wait' -f $file,$safeParam}
        } else {
            if($ext -eq '.msi'){$installCmd='Execute-MSI -Action Install -Path "$dirFiles\'+$file+'" -Parameters "'+$selectedParams+'"';if($transformName){$installCmd='Execute-MSI -Action Install -Path "$dirFiles\'+$file+'" -Parameters "'+$selectedParams+' TRANSFORMS=`"$dirFiles\'+$transformName+'`""'}}
            else{$safeParam=$selectedParams.Replace('"','`"');$installCmd='Execute-Process -Path "$dirFiles\'+$file+'" -Parameters "'+$safeParam+'"'}
        }
        $uninstallTarget=$uninstallCmdBox.Text.Trim();$uninstallArgs=$unParamsBox.Text.Trim();$unCmd=''
        if($uninstallTarget -and $uninstallTarget -notmatch '^\{[0-9A-Fa-f-]{36}\}$'){$uninstallTarget=Convert-ToPortablePath $uninstallTarget}
        if($uninstallTarget){
            if($uninstallTarget -match '^\{[0-9A-Fa-f-]{36}\}$' -or $ext -eq '.msi'){
                if($uninstallTarget -match '^\{[0-9A-Fa-f-]{36}\}$'){$uninstallProductCode=$uninstallTarget}else{$uninstallProductCode=''}
                if($uninstallProductCode){if($psadtVersion.SelectedItem -eq '4.x'){$unCmd='Start-ADTProcess -FilePath "$env:SystemRoot\System32\msiexec.exe" -ArgumentList ''/x {0} /qn /norestart'' -WindowStyle Hidden -Wait' -f $uninstallProductCode}else{$unCmd='Execute-MSI -Action Uninstall -Path "{0}" -Parameters "/qn /norestart"' -f $uninstallProductCode}}
                else {$unCmd=$uninstallTarget}
            } else {
                $safeUninstallArgs=$uninstallArgs.Replace("'","''")
                if($psadtVersion.SelectedItem -eq '4.x'){$unCmd=if($uninstallArgs){'Start-ADTProcess -FilePath "{0}" -ArgumentList ''{1}'' -WindowStyle Hidden -Wait' -f $uninstallTarget,$safeUninstallArgs}else{'Start-ADTProcess -FilePath "{0}" -WindowStyle Hidden -Wait' -f $uninstallTarget}}
                else {$safeUninstallArgs=$uninstallArgs.Replace('"','`"');$unCmd=if($uninstallArgs){'Execute-Process -Path "{0}" -Parameters "{1}" -WindowStyle Hidden -Wait' -f $uninstallTarget,$safeUninstallArgs}else{'Execute-Process -Path "{0}" -WindowStyle Hidden -Wait' -f $uninstallTarget}}
            }
        }
        if($detectBox.SelectedItem -eq 'Run only if missing'){ $displayGuess=[IO.Path]::GetFileNameWithoutExtension($file).Replace("'","''"); $installCmd="`$componentPresent = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | Where-Object { `$_.DisplayName -like '*$displayGuess*' }; if (-not `$componentPresent) {`r`n    $installCmd`r`n}" }
        if($script:EditingAdditionalInstallerIndex -ge 0 -and $script:EditingAdditionalInstallerIndex -lt $script:AdditionalInstallers.Count){
            $oldRecord=$script:AdditionalInstallers[$script:EditingAdditionalInstallerIndex]
            if($oldRecord){try{$oldEditor=Get-PhaseEditor $oldRecord.InstallPhase;$oldEditor.Text=$oldEditor.Text.Replace([string]$oldRecord.InstallCommand,'').Trim()}catch{};if($oldRecord.UninstallCommand){try{$oldUnEditor=Get-PhaseEditor $oldRecord.UninstallPhase;$oldUnEditor.Text=$oldUnEditor.Text.Replace([string]$oldRecord.UninstallCommand,'').Trim()}catch{}}}
        }
        Add-CommandToPhase -Phase $phase -Command $installCmd
        if($unCmd){Add-CommandToPhase -Phase $unPhase -Command $unCmd}
        $metadata=''
        $componentPath=Join-Path $source.Text $file
        if(Test-Path -LiteralPath $componentPath -PathType Leaf){
            if($ext -eq '.exe'){
                $fi=[Diagnostics.FileVersionInfo]::GetVersionInfo($componentPath)
                $displayName=if($fi.ProductName){$fi.ProductName}else{[IO.Path]::GetFileNameWithoutExtension($file)}
                $metadata=('{0} | Version: {1} | Vendor: {2}' -f $displayName,$fi.ProductVersion,$fi.CompanyName).Trim(' ','|')
            } elseif($ext -eq '.msi') {
                try {
                    $wi3=New-Object -ComObject WindowsInstaller.Installer;$db3=$wi3.OpenDatabase($componentPath,0);$p3=@{}
                    foreach($pn3 in @('ProductName','ProductVersion','Manufacturer','ProductCode')){$vw3=$db3.OpenView("SELECT `Value` FROM Property WHERE `Property`='$pn3'");$vw3.Execute();$r3=$vw3.Fetch();if($r3){$p3[$pn3]=[string]$r3.StringData(1)};$vw3.Close()}
                    $metadata=('{0} | Version: {1} | Vendor: {2} | ProductCode: {3}' -f $p3.ProductName,$p3.ProductVersion,$p3.Manufacturer,$p3.ProductCode)
                } catch {$metadata='MSI metadata read failed; review installer manually.'}
            }
        }
        $entry=New-Object System.Windows.Forms.ListViewItem($file);[void]$entry.SubItems.Add($metadata);[void]$entry.SubItems.Add($phase);[void]$entry.SubItems.Add($(if($unCmd){$unPhase}else{'Not configured'}));[void]$entry.SubItems.Add([string]$detectBox.SelectedItem);[void]$entry.SubItems.Add($paramsBox.Text.Trim());$entry.Tag=[pscustomobject]@{InstallPhase=$phase;InstallCommand=$installCmd;UninstallPhase=$unPhase;UninstallCommand=$unCmd}
        $record=[pscustomobject]@{File=$file;Metadata=$metadata;InstallCommand=$installCmd;GenerateMST=[bool]$additionalMst.Checked;Transform=$transformName;InstallPhase=$phase;UninstallPhase=$unPhase;InstallParameters=$paramsBox.Text.Trim();UninstallCommand=$unCmd;UninstallFile=$uninstallCmdBox.Text.Trim();UninstallParameters=$unParamsBox.Text.Trim();Detection=[string]$detectBox.SelectedItem}
        if($script:EditingAdditionalInstallerIndex -ge 0 -and $script:EditingAdditionalInstallerIndex -lt $script:AdditionalInstallers.Count){$idx=$script:EditingAdditionalInstallerIndex;$additionalList.Items.RemoveAt($idx);$additionalList.Items.Insert($idx,$entry);$script:AdditionalInstallers[$idx]=$record}else{[void]$additionalList.Items.Add($entry);$script:AdditionalInstallers.Add($record)}
        $dlg.DialogResult=[System.Windows.Forms.DialogResult]::OK;$dlg.Close()
    })
    if($script:EditingAdditionalInstallerIndex -ge 0 -and $script:EditingAdditionalInstallerIndex -lt $script:AdditionalInstallers.Count){
        $existing=$script:AdditionalInstallers[$script:EditingAdditionalInstallerIndex];$fileBox.Text=$existing.File;$paramsBox.Text=$existing.InstallParameters;$uninstallCmdBox.Text=if($existing.UninstallFile){$existing.UninstallFile}else{$existing.UninstallCommand};$unParamsBox.Text=[string]$existing.UninstallParameters;$unPhaseBox.SelectedItem=$existing.UninstallPhase;$phaseBox.SelectedItem=$existing.InstallPhase;$detectBox.SelectedItem=$existing.Detection;$save.Text='Save Changes'
    }
    [void]$dlg.ShowDialog($form);$dlg.Dispose();$script:EditingAdditionalInstallerIndex=-1
})
$editInstallerButton.Add_Click({ if($additionalList.SelectedItems.Count -eq 0){[System.Windows.Forms.MessageBox]::Show('Select an additional installer to edit.');return}; $script:EditingAdditionalInstallerIndex=$additionalList.SelectedItems[0].Index; $addInstallerButton.PerformClick() })
$removeInstallerButton.Add_Click({
    if($additionalList.SelectedItems.Count -gt 0){
        $item=$additionalList.SelectedItems[0];$tag=$item.Tag
        if($tag){
            try{$ed=Get-PhaseEditor $tag.InstallPhase;$ed.Text=$ed.Text.Replace([string]$tag.InstallCommand,'').Trim()}catch{}
            if($tag.UninstallCommand){try{$ed=Get-PhaseEditor $tag.UninstallPhase;$ed.Text=$ed.Text.Replace([string]$tag.UninstallCommand,'').Trim()}catch{}}
        }
        $idx=$item.Index;$fileName=$item.Text;$item.Remove();if($idx -ge 0 -and $idx -lt $script:AdditionalInstallers.Count){$script:AdditionalInstallers.RemoveAt($idx)}
    }
})

# ---------------- Silent Switch Finder ----------------
$script:SilentSwitchFinderModule=Join-Path $PSScriptRoot '..\Templates\Tools\SilentSwitchFinder\SilentSwitchFinder.psd1'
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

$app.Add_Leave({ if($type.SelectedItem -eq 'EXE' -and $app.Text.Trim()){ Find-InstalledUninstallEntry } })
$ver.Add_Leave({ if($type.SelectedItem -eq 'EXE' -and $app.Text.Trim()){ Find-InstalledUninstallEntry } })

# ---------------- Events ----------------
$type.Add_SelectedValueChanged({
    if ($type.SelectedItem -eq 'MSI') {
        $mst.Enabled = $true
        $mst.Checked = $true
        $installParams.Text = '/qn /norestart'
        $uParams.Text = '/qn /norestart'
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
    if([string]::IsNullOrWhiteSpace($installer.Text)){ return }
    $candidate = $installer.Text.Trim()
    if(-not [IO.Path]::IsPathRooted($candidate) -and $source.Text -and (Test-Path -LiteralPath $source.Text -PathType Container)){
        $candidate = Join-Path $source.Text $candidate
    }
    if(Test-Path -LiteralPath $candidate -PathType Leaf){
        $selectedExt = [IO.Path]::GetExtension($candidate)
        if($selectedExt -ieq '.msi'){
            $type.SelectedItem = 'MSI'
            $mst.Enabled = $true; $mst.Checked = $true
            $installParams.Text = '/qn /norestart'; $uParams.Text = '/qn /norestart'
            Read-MsiPropertiesToUI -MsiPath $candidate
        } elseif($selectedExt -ieq '.exe'){
            $type.SelectedItem = 'EXE'
            $mst.Checked = $false; $mst.Enabled = $false
            if(-not $source.Text -or -not (Test-Path -LiteralPath $source.Text -PathType Container)){
                $source.Text = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($candidate))
            }
            $fi=[Diagnostics.FileVersionInfo]::GetVersionInfo($candidate)
            if([string]::IsNullOrWhiteSpace($app.Text)) { if($fi.ProductName){$app.Text=$fi.ProductName.Trim()} else {$app.Text=[IO.Path]::GetFileNameWithoutExtension($candidate)} }
            if([string]::IsNullOrWhiteSpace($ver.Text)) { if($fi.ProductVersion){$ver.Text=$fi.ProductVersion.Trim()} elseif($fi.FileVersion){$ver.Text=$fi.FileVersion.Trim()} }
            if([string]::IsNullOrWhiteSpace($vendor.Text) -and $fi.CompanyName){$vendor.Text=$fi.CompanyName.Trim()}
            Find-AndPopulateSilentSwitch -InstallerPath $candidate -ParameterBox $installParams -StatusLabel $installSwitchStatus -ContextName 'Installation'
            Find-InstalledUninstallEntry
        }
    }
})

$source.Add_TextChanged({
    if($source.Text -and (Test-Path -LiteralPath $source.Text -PathType Container) -and $installer.Text){
        $candidate=Join-Path $source.Text $installer.Text.Trim()
        if(Test-Path -LiteralPath $candidate -PathType Leaf){
            if($type.SelectedItem -eq 'MSI'){
                Read-MsiPropertiesToUI -MsiPath $candidate
            } else {
                Find-AndPopulateSilentSwitch -InstallerPath $candidate -ParameterBox $installParams -StatusLabel $installSwitchStatus -ContextName 'Installation'
            }
        }
    }
})

$uFile.Add_TextChanged({
    # For MSI, this box contains a ProductCode, not a file path. Never send it to
    # Silent Switch Finder or Test-Path. This prevents MSI auto-detection from being
    # overwritten by EXE uninstall logic.
    if($type.SelectedItem -eq 'MSI' -or $uMethod.SelectedItem -eq 'MSI ProductCode'){
        $uFile.Text = $uFile.Text.Trim()
        if(-not [string]::IsNullOrWhiteSpace($uFile.Text)){
            $uMethod.SelectedItem = 'MSI ProductCode'
            $uParams.Text = '/qn /norestart'
            $uninstallSwitchStatus.Text = 'MSI ProductCode detected automatically. Silent parameters: /qn /norestart'
            $uninstallSwitchStatus.ForeColor=[System.Drawing.Color]::DarkGreen
        }
        return
    }
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
        PSADTVersion    = [string]$psadtVersion.SelectedItem
        AppLogoSource = $appLogoSource.Text.Trim()
    }

    if(-not [string]::IsNullOrWhiteSpace($installParams.Text.Trim())){
        $engineParameters.InstallParameters = $installParams.Text.Trim()
    }
    if($type.SelectedItem -eq 'MSI'){
        # MSI ProductCode is read again by the engine directly from the selected MSI.
        # Do not pass the ProductCode through UninstallFile as if it were an EXE path.
        $engineParameters.UninstallMethod = 'MSI ProductCode'
        $engineParameters.UninstallParameters = '/qn /norestart'
    } else {
        if(-not [string]::IsNullOrWhiteSpace($uFile.Text.Trim())){
            $engineParameters.UninstallFile = $uFile.Text.Trim()
        }
        if(-not [string]::IsNullOrWhiteSpace($uParams.Text.Trim())){
            $engineParameters.UninstallParameters = $uParams.Text.Trim()
        }
    }
    if(-not [string]::IsNullOrWhiteSpace($preInstall.Text.Trim())){
        $engineParameters.PreInstallSteps = $preInstall.Text
    }
    if(-not [string]::IsNullOrWhiteSpace($installSteps.Text.Trim())){
        $engineParameters.InstallSteps = $installSteps.Text
    }
    if(-not [string]::IsNullOrWhiteSpace($postInstall.Text.Trim())){
        $engineParameters.PostInstallSteps = $postInstall.Text
    }
    if(-not [string]::IsNullOrWhiteSpace($preUninstall.Text.Trim())){
        $engineParameters.PreUninstallSteps = $preUninstall.Text
    }
    if(-not [string]::IsNullOrWhiteSpace($uninstallSteps.Text.Trim())){
        $engineParameters.UninstallSteps = $uninstallSteps.Text
    }
    if(-not [string]::IsNullOrWhiteSpace($postUninstall.Text.Trim())){
        $engineParameters.PostUninstallSteps = $postUninstall.Text
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
            $detectedProductCode=$msiProductCode.Text.Trim()
            if([string]::IsNullOrWhiteSpace($detectedProductCode)){throw 'Unable to read ProductCode from the selected MSI.'}
            if($detectedProductCode -notmatch '^\{[0-9A-Fa-f-]{36}\}$'){throw "Invalid MSI ProductCode detected: '$detectedProductCode'"}
            $uMethod.SelectedItem='MSI ProductCode'
            $uFile.Text=$detectedProductCode
            $uParams.Text='/qn /norestart'
            $uninstallSwitchStatus.Text='MSI ProductCode confirmed. Uninstall will use Execute-MSI with /qn /norestart.'
            $uninstallSwitchStatus.ForeColor=[System.Drawing.Color]::DarkGreen
        }

        $script:PackagePrepared=$false
        $script:ScriptConfirmed=$false
        $buildStatus.Text='PREPARING - Creating PSADT package and populating Deploy-Application.ps1...'
        $status.Text='PREPARING'; $status.ForeColor=[System.Drawing.Color]::Gold
        $prepareButton.Enabled=$false; $openDeployButton.Enabled=$false; $confirmBuildButton.Enabled=$false
        $tabs.SelectedTab=$runTab; $runBox.Clear()

        if($type.SelectedItem -eq 'EXE' -and [string]::IsNullOrWhiteSpace($uFile.Text.Trim())){
            $uninstallSwitchStatus.Text='Unable to find the uninstall string. Please add it manually.'
            $uninstallSwitchStatus.ForeColor=[System.Drawing.Color]::DarkOrange
            [System.Windows.Forms.MessageBox]::Show('Unable to find the uninstall string. Please add it manually in the Uninstall tab before preparing the package.','Uninstall Detection',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
            $tabs.SelectedTab=$uninstall; $prepareButton.Enabled=$true; return
        }
        $engineParameters=Get-BuildEngineParameters
        $engineParameters.UseRegistryUninstall = $false
        $engineParameters.PrepareOnly = $true
        Invoke-PackagingEngineForWorkflow -EngineParameters $engineParameters

        $pkgName=if($arch.SelectedItem -eq 'x86'){ '{0}_{1}_{2}_P1' -f ($vendor.Text.Trim() -replace '\s+','-'),($app.Text.Trim() -replace '\s+','-'),($ver.Text.Trim() -replace '\s+','-') }else{ '{0}_{1}_{2}_P1_{3}' -f ($vendor.Text.Trim() -replace '\s+','-'),($app.Text.Trim() -replace '\s+','-'),($ver.Text.Trim() -replace '\s+','-'),$arch.SelectedItem }
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
            Start-Process -FilePath $ise.Source -ArgumentList @($deployPath)
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



# Build/System context controls
$systemTestGroup=New-Object System.Windows.Forms.GroupBox;$systemTestGroup.Text='SYSTEM Context Testing';$systemTestGroup.Location=New-Object System.Drawing.Point(28,478);$systemTestGroup.Size=New-Object System.Drawing.Size(430,86);$buildTab.Controls.Add($systemTestGroup)
$systemInstall=New-Object System.Windows.Forms.Button;$systemInstall.Text='INSTALL AS SYSTEM';$systemInstall.Location=New-Object System.Drawing.Point(16,30);$systemInstall.Size=New-Object System.Drawing.Size(185,36);$systemInstall.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat;$systemTestGroup.Controls.Add($systemInstall)
$systemUninstall=New-Object System.Windows.Forms.Button;$systemUninstall.Text='UNINSTALL AS SYSTEM';$systemUninstall.Location=New-Object System.Drawing.Point(215,30);$systemUninstall.Size=New-Object System.Drawing.Size(195,36);$systemUninstall.FlatStyle=[System.Windows.Forms.FlatStyle]::Flat;$systemTestGroup.Controls.Add($systemUninstall)
function Invoke-SystemPackage([string]$Action){
    if(-not $script:PackagePrepared){ throw 'Prepare the package first.' }
    $pr = Join-Path $script:Root 'Templates\Tools\PowerRun.exe'
    $packageDir = Join-Path $script:PreparedPackageRoot 'Package'
    $deploy = Join-Path $packageDir 'Deploy-Application.exe'
    if(-not(Test-Path -LiteralPath $pr)){ throw "PowerRun.exe not found: $pr" }
    if(-not(Test-Path -LiteralPath $deploy)){ throw "Deploy-Application.exe not found: $deploy" }

    # PowerRun command-line mode:
    # /SYS = run as LocalSystem, /WD = working directory, then executable + command.
    $args = @('/SYS', ('/WD:"{0}"' -f $packageDir), ('"{0}"' -f $deploy), $Action)
    $runBox.AppendText(("[SYSTEM] PowerRun: {0} {1}`r`n" -f $pr, ($args -join ' ')))
    Start-Process -FilePath $pr -ArgumentList $args -Wait
    $runBox.AppendText("[SYSTEM] Completed $Action using PowerRun /SYS.`r`n")
}
$systemInstall.Add_Click({try{Invoke-SystemPackage 'Install'}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'SYSTEM Install')}})
$systemUninstall.Add_Click({try{Invoke-SystemPackage 'Uninstall'}catch{[System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'SYSTEM Uninstall')}})

# Select Package tab and show the window.
$tabs.SelectedTab = $general
Update-BuildSummary
$tabs.Visible = $true

$null = $form.ShowDialog()
