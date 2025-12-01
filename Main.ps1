#region Imports & Config
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $exe = Join-Path $PSHome 'powershell.exe'
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    Start-Process -FilePath $exe -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -File `"$scriptPath`""
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.ComponentModel
try {
    Add-Type -AssemblyName 'Microsoft.VisualBasic' -ErrorAction SilentlyContinue
} catch {}

$ScriptRootPath = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $PSScriptRoot) { $PSScriptRoot = $ScriptRootPath }
$modulesRoot = Join-Path $ScriptRootPath 'Modules'

### Import configuration and other modules
. (Join-Path $modulesRoot 'Config.ps1') -ScriptRoot $ScriptRootPath
. (Join-Path $modulesRoot 'Splash.ps1')
. (Join-Path $modulesRoot 'UiStyling.ps1')
. (Join-Path $modulesRoot 'Logging.ps1')
. (Join-Path $modulesRoot 'Result.ps1')
. (Join-Path $modulesRoot 'DataHelpers.ps1')
# Import control material helper efter övriga moduler så att hjälpfunktioner finns laddade
. (Join-Path $modulesRoot 'ControlMaterialHelper.ps1')
. (Join-Path $modulesRoot 'Compiling.ps1')
. (Join-Path $modulesRoot 'SignatureHelpers.ps1')
. (Join-Path $modulesRoot 'SealTestHelpers.ps1')
. (Join-Path $modulesRoot 'ReportBuilder.ps1')
. (Join-Path $modulesRoot 'ValidateAssayRun.ps1')

$global:StartupReady = $true
$configStatus = $null

try {

    $configStatus = Test-Config
    if ($configStatus) {
        foreach ($err in $configStatus.Errors) { Gui-Log "❌ Konfig-fel: $err" 'Error' }
        foreach ($warn in $configStatus.Warnings) { Gui-Log "⚠️ Konfig-varning: $warn" 'Warn' }
        if (-not $configStatus.Ok) {
            $global:StartupReady = $false
            try { [System.Windows.Forms.MessageBox]::Show("Startkontroll misslyckades. Se logg för detaljer.","Startkontroll") | Out-Null } catch {}
        }
    }
} catch { Gui-Log "❌ Test-Config misslyckades: $($_.Exception.Message)" 'Error'; $global:StartupReady = $false }

$Host.UI.RawUI.WindowTitle = "Startar…"
Show-Splash "Laddar PnP.PowerShell…"
$global:SpConnected = $false
$global:SpError     = $null

try {
    $null = Get-PackageProvider -Name "NuGet" -ForceBootstrap -ErrorAction SilentlyContinue
} catch {}

try {
    Update-Splash "Laddar..."
    Import-Module PnP.PowerShell -ErrorAction Stop
} catch {
    try {
        Gui-Log "ℹ️ PowerShell-modulen saknas, installerar modulen..." 'Info'
        Install-Module PnP.PowerShell -MaximumVersion 1.12.0 -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Update-Splash "Laddar..."
        Import-Module PnP.PowerShell -ErrorAction Stop
    } catch {
        $global:SpError = "PnP-install/import misslyckades: $($_.Exception.Message)"
    }
}
$env:PNPPOWERSHELL_UPDATECHECK = "Off"
try { $null = Ensure-EPPlus -Version '4.5.3.3' } catch { Gui-Log "⚠️ EPPlus-förkontroll misslyckades: $($_.Exception.Message)" 'Warn' }
try {
    if (-not (Load-EPPlus)) {
        Gui-Log "⚠️ EPPlus kunde inte laddas – Excel-åtkomst kan misslyckas." 'Warn'
    }
} catch { Gui-Log "⚠️ EPPlus-laddning misslyckades: $($_.Exception.Message)" 'Warn' }

# Bekräfta att kritiska moduler är tillgängliga i bakgrundslogg
try {
    foreach ($modName in @('PnP.PowerShell')) {
        $isLoaded = Get-Module -Name $modName -ListAvailable -ErrorAction SilentlyContinue
        $msg = if ($isLoaded) { "Modul laddad eller tillgänglig: $modName" } else { "Modul saknas: $modName" }
        $severity = if ($isLoaded) { 'Info' } else { 'Warn' }
        Write-BackendLog -Message $msg -Severity $severity
    }
    $epplusLoaded = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'EPPlus' }
    if ($epplusLoaded) {
        Write-BackendLog -Message 'EPPlus assembly laddad i AppDomain.' -Severity 'Info'
    } else {
        Write-BackendLog -Message 'EPPlus assembly kunde inte verifieras som laddad.' -Severity 'Warn'
    }
} catch {
    Write-BackendError -Context 'Modulkontroll' -ErrorRecord $_
}

# Efter att EPPlus laddats, läs in kontrollmaterial-kartan en gång
try {
    # Kontrollmaterial-kartan kräver att stigen är definierad i Config.ps1 som $Global:ControlMaterialMapPath
    if ($Global:ControlMaterialMapPath) {
        $global:ControlMaterialData = Get-ControlMaterialMap
        if ($global:ControlMaterialData -and $global:ControlMaterialData.PartNoIndex) {
            $pnCount  = $global:ControlMaterialData.PartNoIndex.Count
            $assCount = $global:ControlMaterialData.AssayUsageIndex.Count
            Gui-Log ("ℹ️ ControlMaterialMap: {0} P/N och {1} assays laddade." -f $pnCount, $assCount) 'Info'
        } else {
            Gui-Log "ℹ️ ControlMaterialMap: inga data eller kartfil saknas." 'Info'
        }
    } else {
        Gui-Log "ℹ️ ControlMaterialMapPath inte definierat." 'Info'
    }
} catch {
    Gui-Log ("⚠️ Kunde inte läsa ControlMaterialMap: " + $_.Exception.Message) 'Warn'
}

if (-not $global:SpError) {
    try {
        Update-Splash "Ansluter till SharePoint"
        Connect-PnPOnline -Url $global:SP_SiteUrl `
                          -Tenant $global:SP_Tenant `
                          -ClientId $global:SP_ClientId `
                          -CertificateBase64Encoded $global:SP_CertBase64 `
                          -ErrorAction Stop
        $global:SpConnected = $true
    } catch {
        $msg = "Connect-PnPOnline misslyckades: $($_.Exception.Message)"
        Update-Splash $msg
        $global:SpError = $msg
    }
}

#endregion Imports & Config

#region GUI Construction

Update-Splash "Startar gränssnitt…"
Close-Splash
$form = New-Object System.Windows.Forms.Form
$form.Text = "$ScriptVersion QWERTYUIO Höhö... "
$form.AutoScaleMode = 'Dpi'
$form.Size = New-Object System.Drawing.Size(860,910)
$form.MinimumSize = New-Object System.Drawing.Size(860,910)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.AutoScroll  = $false
$form.MaximizeBox = $false
$form.Padding     = New-Object System.Windows.Forms.Padding(8)
$form.Font        = New-Object System.Drawing.Font('Segoe UI',10)
$form.KeyPreview = $true
$form.add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $form.Close() } })

# ---------- Menyrad ----------
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.Dock='Top'; $menuStrip.GripStyle='Hidden'
$menuStrip.ImageScalingSize = New-Object System.Drawing.Size(20,20)
$menuStrip.Padding = New-Object System.Windows.Forms.Padding(8,6,0,6)
$menuStrip.Font = New-Object System.Drawing.Font('Segoe UI',10)
$miArkiv   = New-Object System.Windows.Forms.ToolStripMenuItem('🗂️ Arkiv')
$miVerktyg = New-Object System.Windows.Forms.ToolStripMenuItem('🛠️ Verktyg')
$miSettings= New-Object System.Windows.Forms.ToolStripMenuItem('⚙️ Inställningar')
$miHelp    = New-Object System.Windows.Forms.ToolStripMenuItem('📖 Instruktioner')
$miAbout   = New-Object System.Windows.Forms.ToolStripMenuItem('ℹ️ Om')
$miScan  = New-Object System.Windows.Forms.ToolStripMenuItem('🔍 Sök filer')
$miBuild = New-Object System.Windows.Forms.ToolStripMenuItem('✅ Skapa rapport')
$miExit  = New-Object System.Windows.Forms.ToolStripMenuItem('❌ Avsluta')

# Rensa ev. gamla undermenyer
$miArkiv.DropDownItems.Clear()
$miVerktyg.DropDownItems.Clear()
$miSettings.DropDownItems.Clear()
$miHelp.DropDownItems.Clear()

# ----- Arkiv -----
$miNew         = New-Object System.Windows.Forms.ToolStripMenuItem('🆕 Nytt')
$miOpenRecent  = New-Object System.Windows.Forms.ToolStripMenuItem('📂 Öppna senaste rapport')
$miArkiv.DropDownItems.AddRange(@(
    $miNew,
    $miOpenRecent,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $miExit
))


# ----- Verktyg -----
$miScript1   = New-Object System.Windows.Forms.ToolStripMenuItem('📜 Kontrollprovsfilskript')
$miScript2   = New-Object System.Windows.Forms.ToolStripMenuItem('📅 Ändra datum-prefix för filnamn')
$miScript3   = New-Object System.Windows.Forms.ToolStripMenuItem('📅 TBD')
$miToggleSign = New-Object System.Windows.Forms.ToolStripMenuItem('✅ Aktivera Seal Test-signatur')
$miVerktyg.DropDownItems.AddRange(@(
    $miScript1,
    $miScript2,
    $miScript3,
    $miToggleSign
))

# ----- Inställningar -----
$miTheme = New-Object System.Windows.Forms.ToolStripMenuItem('🎨 Tema')
$miLightTheme = New-Object System.Windows.Forms.ToolStripMenuItem('☀️ Ljust (default)')
$miDarkTheme  = New-Object System.Windows.Forms.ToolStripMenuItem('🌙 Mörkt')
$miTheme.DropDownItems.AddRange(@($miLightTheme,$miDarkTheme))
$miSettings.DropDownItems.Add($miTheme)

# ----- Instruktioner -----
$miShowInstr   = New-Object System.Windows.Forms.ToolStripMenuItem('📖 Visa instruktioner')
$miFAQ         = New-Object System.Windows.Forms.ToolStripMenuItem('❓ Vanliga frågor (FAQ)')
$miHelpDlg     = New-Object System.Windows.Forms.ToolStripMenuItem('🆘 Hjälp')
$miHelp.DropDownItems.AddRange(@($miShowInstr,$miFAQ,$miHelpDlg))

$miGenvagar = New-Object System.Windows.Forms.ToolStripMenuItem('🔗 Genvägar')
$ShortcutGroups = @{
    '🗂️ IPT-mappar' = @(

        @{ Text='📂 IPT - PÅGÅENDE KÖRNINGAR';        Target='N:\QC\QC-1\IPT\2. IPT - PÅGÅENDE KÖRNINGAR' },
        @{ Text='📂 IPT - KLART FÖR SAMMANSTÄLLNING'; Target='N:\QC\QC-1\IPT\3. IPT - KLART FÖR SAMMANSTÄLLNING' },
        @{ Text='📂 IPT - KLART FÖR GRANSKNING';      Target='N:\QC\QC-1\IPT\4. IPT - KLART FÖR GRANSKNING' },
        @{ Text='📂 SPT Macro Assay';                 Target='N:\QC\QC-0\SPT\SPT macros\Assay' }

    )

    '📄 Dokument' = @(

        @{ Text='🧰 Utrustningslista';    Target=$UtrustningListPath },
        @{ Text='🧪 Kontrollprovsfil';    Target=$RawDataPath }

    )

    '🌐 Länkar' = @(

        @{ Text='⚡ IPT App';              Target='https://apps.powerapps.com/play/e/default-771c9c47-7f24-44dc-958e-34f8713a8394/a/fd340dbd-bbbf-470b-b043-d2af4cb62c83' },
        @{ Text='🌐 MES';                  Target='http://mes.cepheid.pri/camstarportal/?domain=CEPHEID.COM' },
        @{ Text='🌐 CSV Uploader';         Target='http://auw2wgxtpap01.cepaws.com/Welcome.aspx' },
        @{ Text='🌐 BMRAM';                Target='https://cepheid62468.coolbluecloud.com/' },
        @{ Text='🌐 Agile';                Target='https://agileprod.cepheid.com/Agile/default/login-cms.jsp' }

    )

}

foreach ($grp in $ShortcutGroups.GetEnumerator()) {

    $grpMenu = New-Object System.Windows.Forms.ToolStripMenuItem($grp.Key)
    foreach ($entry in $grp.Value) { Add-ShortcutItem -Parent $grpMenu -Text $entry.Text -Target $entry.Target }
    [void]$miGenvagar.DropDownItems.Add($grpMenu)

}

$miOm = New-Object System.Windows.Forms.ToolStripMenuItem('ℹ️ Om det här verktyget'); $miAbout.DropDownItems.Add($miOm)
$menuStrip.Items.AddRange(@($miArkiv,$miVerktyg,$miGenvagar,$miSettings,$miHelp,$miAbout))
$form.MainMenuStrip=$menuStrip

# ---------- Header ----------
$panelHeader = New-Object System.Windows.Forms.Panel
$panelHeader.Dock='Top'; $panelHeader.Height=64
$panelHeader.BackColor=[System.Drawing.Color]::SteelBlue
$panelHeader.Padding = New-Object System.Windows.Forms.Padding(10,8,10,8)

$picLogo = New-Object System.Windows.Forms.PictureBox
$picLogo.Dock='Left'; $picLogo.Width=50; $picLogo.BorderStyle='FixedSingle'
if(Test-Path $ikonSokvag){ $picLogo.Image=[System.Drawing.Image]::FromFile($ikonSokvag); $picLogo.SizeMode='Zoom' }

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text="$ScriptVersion - TESTA ATT ANVÄNDA. KUL!"
$lblTitle.ForeColor=[System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold',13)
$lblTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblTitle.Padding = New-Object System.Windows.Forms.Padding(8,0,0,0)
$lblTitle.Dock='Fill'

$panelHeader.Controls.Add($lblTitle)
$panelHeader.Controls.Add($picLogo)

# ---------- Sök-rad ----------

$tlSearch = New-Object System.Windows.Forms.TableLayoutPanel
$tlSearch.Dock='Top'; $tlSearch.AutoSize=$true; $tlSearch.AutoSizeMode='GrowAndShrink'
$tlSearch.Padding = New-Object System.Windows.Forms.Padding(0,10,0,8)
$tlSearch.ColumnCount=3
[void]$tlSearch.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$tlSearch.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
[void]$tlSearch.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,130)))

$lblLSP = New-Object System.Windows.Forms.Label
$lblLSP.Text='LSP:'; $lblLSP.Anchor='Left'; $lblLSP.AutoSize=$true
$lblLSP.Margin = New-Object System.Windows.Forms.Padding(0,6,8,0)
$txtLSP = New-Object System.Windows.Forms.TextBox
$txtLSP.Dock='Fill'
$txtLSP.Margin = New-Object System.Windows.Forms.Padding(0,2,10,2)
$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text='Sök filer'; $btnScan.Dock='Fill'; Set-AccentButton $btnScan -Primary
$btnScan.Margin= New-Object System.Windows.Forms.Padding(0,2,0,2)

$tlSearch.Controls.Add($lblLSP,0,0)
$tlSearch.Controls.Add($txtLSP,1,0)
$tlSearch.Controls.Add($btnScan,2,0)

$pLog = New-Object System.Windows.Forms.Panel
$pLog.Dock='Top'; $pLog.Height=220; $pLog.Padding=New-Object System.Windows.Forms.Padding(0,0,0,8)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline=$true; $outputBox.ScrollBars='Vertical'; $outputBox.ReadOnly=$true
$outputBox.BackColor='White'; $outputBox.Dock='Fill'
$outputBox.Font = New-Object System.Drawing.Font('Segoe UI',9)
$pLog.Controls.Add($outputBox)
try { Set-LogOutputControl -Control $outputBox } catch {}

$grpPick = New-Object System.Windows.Forms.GroupBox
$grpPick.Text='Välj filer för rapport'
$grpPick.Dock='Top'
$grpPick.Padding = New-Object System.Windows.Forms.Padding(10,12,10,14)
$grpPick.AutoSize=$false
$grpPick.Height = (78*3) + $grpPick.Padding.Top + $grpPick.Padding.Bottom +15

$tlPick = New-Object System.Windows.Forms.TableLayoutPanel
$tlPick.Dock='Fill'; $tlPick.ColumnCount=3; $tlPick.RowCount=3
$tlPick.GrowStyle=[System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize
[void]$tlPick.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$tlPick.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
[void]$tlPick.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,100)))
for($i=0;$i -lt 3;$i++){ [void]$tlPick.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,78))) }

function New-ListRow {
    param([string]$labelText,[ref]$lbl,[ref]$clb,[ref]$btn)
    $lbl.Value = New-Object System.Windows.Forms.Label
    $lbl.Value.Text=$labelText
    $lbl.Value.Anchor='Left'
    $lbl.Value.AutoSize=$true
    $lbl.Value.Margin=New-Object System.Windows.Forms.Padding(0,12,6,0)
    $clb.Value = New-Object System.Windows.Forms.CheckedListBox
    $clb.Value.Dock='Fill'
    $clb.Value.Margin=New-Object System.Windows.Forms.Padding(0,6,8,6)
    $clb.Value.Height=70
    $clb.Value.IntegralHeight=$false
    $clb.Value.CheckOnClick = $true
    $clb.Value.DisplayMember = 'Name'

    $btn.Value = New-Object System.Windows.Forms.Button
    $btn.Value.Text='Bläddra…'
    $btn.Value.Dock='Fill'
    $btn.Value.Margin=New-Object System.Windows.Forms.Padding(0,6,0,6)
    Set-AccentButton $btn.Value
}

$lblCsv=$null;$clbCsv=$null;$btnCsvBrowse=$null
New-ListRow -labelText 'CSV-fil:' -lbl ([ref]$lblCsv) -clb ([ref]$clbCsv) -btn ([ref]$btnCsvBrowse)
$lblNeg=$null;$clbNeg=$null;$btnNegBrowse=$null
New-ListRow -labelText 'Seal Test Neg:' -lbl ([ref]$lblNeg) -clb ([ref]$clbNeg) -btn ([ref]$btnNegBrowse)
$lblPos=$null;$clbPos=$null;$btnPosBrowse=$null
New-ListRow -labelText 'Seal Test Pos:' -lbl ([ref]$lblPos) -clb ([ref]$clbPos) -btn ([ref]$btnPosBrowse)

try {
    if ($tlPick.RowCount -lt 4) {
        $tlPick.RowCount = 4
        for ($i=$tlPick.RowStyles.Count; $i -lt 4; $i++) {
            $null = $tlPick.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 78)) )
        }
        $grpPick.Height = (78*4) + $grpPick.Padding.Top + $grpPick.Padding.Bottom + 15
    }
} catch {}

$lblLsp = $null; $clbLsp = $null; $btnLspBrowse = $null
New-ListRow -labelText 'Worksheet:' -lbl ([ref]$lblLsp) -clb ([ref]$clbLsp) -btn ([ref]$btnLspBrowse)

$tlPick.Controls.Add($lblLsp,  0, 3)
$tlPick.Controls.Add($clbLsp,  1, 3)
$tlPick.Controls.Add($btnLspBrowse, 2, 3)

$clbLsp.add_ItemCheck({
    param($s,$e)
    if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
        for ($i=0; $i -lt $s.Items.Count; $i++) {
            if ($i -ne $e.Index) { $s.SetItemChecked($i, $false) }
        }
    }
})

$btnLspBrowse.Add_Click({
    try {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Excel|*.xlsx;*.xlsm|Alla filer|*.*"
        $dlg.Title  = "Välj LSP Worksheet"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $f = Get-Item -LiteralPath $dlg.FileName
            Add-CLBItems -clb $clbLsp -files @($f) -AutoCheckFirst
            if (Get-Command Update-StatusBar -ErrorAction SilentlyContinue) { Update-StatusBar }
        }
    } catch {
        Gui-Log ("⚠️ LSP-browse fel: " + $_.Exception.Message) 'Warn'
    }
})

# Lägg in i tabellen
$tlPick.Controls.Add($lblCsv,0,0); $tlPick.Controls.Add($clbCsv,1,0); $tlPick.Controls.Add($btnCsvBrowse,2,0)
$tlPick.Controls.Add($lblNeg,0,1); $tlPick.Controls.Add($clbNeg,1,1); $tlPick.Controls.Add($btnNegBrowse,2,1)
$tlPick.Controls.Add($lblPos,0,2); $tlPick.Controls.Add($clbPos,1,2); $tlPick.Controls.Add($btnPosBrowse,2,2)
$grpPick.Controls.Add($tlPick)

# ---------- Signatur ----------
$grpSign = New-Object System.Windows.Forms.GroupBox
$grpSign.Text = "Lägg till signatur i Seal Test-filerna"
$grpSign.Dock='Top'
$grpSign.Padding = New-Object System.Windows.Forms.Padding(10,8,10,10)
$grpSign.AutoSize = $false
$grpSign.Height = 88

$tlSign = New-Object System.Windows.Forms.TableLayoutPanel
$tlSign.Dock='Fill'; $tlSign.ColumnCount=2; $tlSign.RowCount=2
[void]$tlSign.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$tlSign.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
[void]$tlSign.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,28)))
[void]$tlSign.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,28)))

$lblSigner = New-Object System.Windows.Forms.Label
$lblSigner.Text = 'Fullständigt namn, signatur och datum:'
$lblSigner.Anchor='Left'; $lblSigner.AutoSize=$true

$txtSigner = New-Object System.Windows.Forms.TextBox
$txtSigner.Dock='Fill'; $txtSigner.Margin = New-Object System.Windows.Forms.Padding(6,2,0,2)
$chkWriteSign = New-Object System.Windows.Forms.CheckBox
$chkWriteSign.Text = 'Signera Seal Test-Filerna'
$chkWriteSign.Anchor='Left'
$chkWriteSign.AutoSize = $true

$chkOverwriteSign = New-Object System.Windows.Forms.CheckBox
$chkOverwriteSign.Text = 'Aktivera'

$chkOverwriteSign.Anchor='Left'
$chkOverwriteSign.AutoSize = $true
$chkOverwriteSign.Enabled = $false
$chkWriteSign.add_CheckedChanged({ $chkOverwriteSign.Enabled = $chkWriteSign.Checked })

$tlSign.Controls.Add($lblSigner,0,0); $tlSign.Controls.Add($txtSigner,1,0)
$tlSign.Controls.Add($chkWriteSign,0,1); $tlSign.Controls.Add($chkOverwriteSign,1,1)
$grpSign.Controls.Add($tlSign)

$grpSign.Visible = $false
$baseHeight = $form.Height

# ---------- Utdatasparande ----------
$grpSave = New-Object System.Windows.Forms.GroupBox
$grpSave.Text = "Rapport-utdata"
$grpSave.Dock='Top'
$grpSave.Padding = New-Object System.Windows.Forms.Padding(10,8,10,10)
$grpSave.AutoSize = $false
$grpSave.Height = 62

$flSave = New-Object System.Windows.Forms.FlowLayoutPanel
$flSave.Dock='Fill'
$rbSaveInLsp = New-Object System.Windows.Forms.RadioButton
$rbSaveInLsp.Text = "Spara i LSP-mapp"
$rbSaveInLsp.Checked = $true
$rbSaveInLsp.AutoSize = $true
$rbTempOnly = New-Object System.Windows.Forms.RadioButton
$rbTempOnly.Text = "Öppna i temporärt läge"
$rbTempOnly.AutoSize = $true
$flSave.Controls.Add($rbSaveInLsp); $flSave.Controls.Add($rbTempOnly)
$chkSharePointInfo = New-Object System.Windows.Forms.CheckBox
$chkSharePointInfo.Text = "Inkludera SharePoint Info"
$chkSharePointInfo.AutoSize = $true
$chkSharePointInfo.Checked = $false
$flSave.Controls.Add($chkSharePointInfo)

$grpSave.Controls.Add($flSave)

# ---------- Primärknapp ----------
$btnBuild = New-Object System.Windows.Forms.Button
$btnBuild.Text='Skapa rapport'; $btnBuild.Dock='Top'; $btnBuild.Height=40
$btnBuild.Margin = New-Object System.Windows.Forms.Padding(0,16,0,8)
$btnBuild.Enabled=$false; Set-AccentButton $btnBuild -Primary

# ---------- Statusrad ----------
$status = New-Object System.Windows.Forms.StatusStrip
$status.SizingGrip=$false; $status.Dock='Bottom'; $status.Font=New-Object System.Drawing.Font('Segoe UI',9)
$status.ShowItemToolTips = $true
$slCount = New-Object System.Windows.Forms.ToolStripStatusLabel; $slCount.Text='0 filer valda'; $slCount.Spring=$false
$slSpacer= New-Object System.Windows.Forms.ToolStripStatusLabel; $slSpacer.Spring=$true

# --- Klickbar SharePoint-länk ---
$slBatchLink = New-Object System.Windows.Forms.ToolStripStatusLabel
$slBatchLink.IsLink   = $true
$slBatchLink.Text     = 'SharePoint: —'
$slBatchLink.Enabled  = $false
$slBatchLink.Tag      = $null
$slBatchLink.ToolTipText = 'Direktlänk aktiveras när Batch# hittas i filer.'
$slBatchLink.add_Click({
    if ($this.Enabled -and $this.Tag) {
        try { Start-Process $this.Tag } catch {
            [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna:`n$($this.Tag)`n$($_.Exception.Message)","Länk") | Out-Null
        }
    }
})

$status.Items.AddRange(@($slCount,$slSpacer,$slBatchLink))
$tsc = New-Object System.Windows.Forms.ToolStripContainer
$tsc.Dock = 'Fill'
$tsc.LeftToolStripPanelVisible  = $false
$tsc.RightToolStripPanelVisible = $false

$form.SuspendLayout()
$form.Controls.Clear()
$form.Controls.Add($tsc)

# Meny högst upp
$tsc.TopToolStripPanel.Controls.Add($menuStrip)
$form.MainMenuStrip = $menuStrip

# Status längst ner
$tsc.BottomToolStripPanel.Controls.Add($status)

# Content i mitten
$content = New-Object System.Windows.Forms.Panel
$content.Dock='Fill'
$content.BackColor = $form.BackColor
$tsc.ContentPanel.Controls.Add($content)

# Dock=Top: nedersta först
$content.SuspendLayout()
$content.Controls.Add($btnBuild)
$content.Controls.Add($grpSave)
$content.Controls.Add($grpSign)
$content.Controls.Add($grpPick)
$content.Controls.Add($pLog)
$content.Controls.Add($tlSearch)
$content.Controls.Add($panelHeader)
$content.ResumeLayout()
$form.ResumeLayout()
$form.PerformLayout()
$form.AcceptButton = $btnScan

#endregion GUI Construction

function Add-CLBItems {
    param([System.Windows.Forms.CheckedListBox]$clb,[System.IO.FileInfo[]]$files,[switch]$AutoCheckFirst)
    $clb.BeginUpdate()
    $clb.Items.Clear()
    foreach($f in $files){
        if ($f -isnot [System.IO.FileInfo]) { try { $f = Get-Item -LiteralPath $f } catch { continue } }
        [void]$clb.Items.Add($f, $false)
    }
    $clb.EndUpdate()
    if ($AutoCheckFirst -and $clb.Items.Count -gt 0) { $clb.SetItemChecked(0,$true) }
    Update-StatusBar
}

function Get-CheckedFilePath { param([System.Windows.Forms.CheckedListBox]$clb)
    for($i=0;$i -lt $clb.Items.Count;$i++){
        if ($clb.GetItemChecked($i)) {
            $fi = [System.IO.FileInfo]$clb.Items[$i]
            return $fi.FullName
        }
    }
    return $null
}

function Clear-GUI {
    $txtLSP.Text = ''
    $txtSigner.Text = ''
    $chkWriteSign.Checked = $false
    $chkOverwriteSign.Checked = $false
    Add-CLBItems -clb $clbCsv -files @()
    Add-CLBItems -clb $clbNeg -files @()
    Add-CLBItems -clb $clbPos -files @()
    Add-CLBItems -clb $clbLsp -files @() 
    $outputBox.Clear()
    Update-BuildEnabled
    Gui-Log "🧹 GUI rensat." 'Info'
    Update-BatchLink
}

$onExclusive = {
    $clb = $this
    if ($_.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
        for ($i=0; $i -lt $clb.Items.Count; $i++) {
            if ($i -ne $_.Index -and $clb.GetItemChecked($i)) { $clb.SetItemChecked($i, $false) }
        }
    }
    $clb.BeginInvoke([Action]{ Update-BuildEnabled }) | Out-Null
}
$clbCsv.add_ItemCheck($onExclusive)
$clbNeg.add_ItemCheck($onExclusive)
$clbPos.add_ItemCheck($onExclusive)

function Get-SelectedFileCount {
    $n=0
    if (Get-CheckedFilePath $clbCsv) { $n++ }
    if (Get-CheckedFilePath $clbNeg) { $n++ }
    if (Get-CheckedFilePath $clbPos) { $n++ }
    if (Get-CheckedFilePath $clbLsp) { $n++ }
    return $n
}

function Update-StatusBar { $slCount.Text = "$(Get-SelectedFileCount) filer valda" }

function Update-BuildEnabled {
    $btnBuild.Enabled = ((Get-CheckedFilePath $clbNeg) -and (Get-CheckedFilePath $clbPos))
    Update-StatusBar
}

$script:LastScanResult = $null

function Get-BatchLinkInfo {
    param(
        [string]$SealPosPath,
        [string]$SealNegPath,
        [string]$Lsp
    )

    $batch = $null
    try { if ($SealPosPath) { $batch = Get-BatchNumberFromSealFile $SealPosPath } } catch {}
    if (-not $batch) {
        try { if ($SealNegPath) { $batch = Get-BatchNumberFromSealFile $SealNegPath } } catch {}
    }

    $batchEsc = if ($batch) { [uri]::EscapeDataString($batch) } else { '' }
    $lspEsc   = if ($Lsp)   { [uri]::EscapeDataString($Lsp) }   else { '' }

    $url = if ($SharePointBatchLinkTemplate) {
        ($SharePointBatchLinkTemplate -replace '\{BatchNumber\}', $batchEsc) -replace '\{LSP\}', $lspEsc
    } else {
        "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management/Lists/Cepheid%20%20Production%20orders/AllItems.aspx?view=7&q=$batchEsc"
    }
    $linkText = if ($batch) { "Öppna $batch" } else { 'Ingen batch funnen' }

    return [pscustomobject]@{
        Batch    = $batch
        Url      = $url
        LinkText = $linkText
    }
}

function Assert-StartupReady {
    if ($global:StartupReady) { return $true }
    Gui-Log "❌ Startkontroll misslyckades. Åtgärda konfigurationsfel innan du fortsätter." 'Error'
    return $false
}

#region Event Handlers
$miScan.add_Click({ $btnScan.PerformClick() })
$miBuild.add_Click({ if ($btnBuild.Enabled) { $btnBuild.PerformClick() } })
$miExit.add_Click({ $form.Close() })
$miNew.add_Click({ Clear-GUI })

$miOpenRecent.add_Click({
    if ($global:LastReportPath -and (Test-Path -LiteralPath $global:LastReportPath)) {
        try { Start-Process -FilePath $global:LastReportPath } catch {
            [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna rapporten:\n$($_.Exception.Message)","Öppna senaste rapport") | Out-Null
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Ingen rapport har genererats i denna session.","Öppna senaste rapport") | Out-Null
    }
})

# Skript1..3
$miScript1.add_Click({
    $p = $Script1Path
    if ([string]::IsNullOrWhiteSpace($p)) { [System.Windows.Forms.MessageBox]::Show("Ange sökvägen till Skript1 i variabeln `$Script1Path.","Skript1") | Out-Null; return }
    if (-not (Test-Path -LiteralPath $p)) { [System.Windows.Forms.MessageBox]::Show("Filen hittades inte:\n$Script1Path","Skript1") | Out-Null; return }
    $ext=[System.IO.Path]::GetExtension($p).ToLowerInvariant()
    switch ($ext) {
        '.ps1' { Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$p`"" }
        '.bat' { Start-Process cmd.exe -ArgumentList "/c `"$p`"" }
        '.lnk' { Start-Process -FilePath $p }
        default { try { Start-Process -FilePath $p } catch { [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna filen:","Skript1") | Out-Null } }
    }
})

$miScript2.add_Click({
    $p = $Script2Path
    if ([string]::IsNullOrWhiteSpace($p)) { [System.Windows.Forms.MessageBox]::Show("Ange sökvägen till Skript2 i variabeln `$Script2Path.","Skript2") | Out-Null; return }
    if (-not (Test-Path -LiteralPath $p)) { [System.Windows.Forms.MessageBox]::Show("Filen hittades inte:\n$Script2Path","Skript2") | Out-Null; return }
    $ext=[System.IO.Path]::GetExtension($p).ToLowerInvariant()
    switch ($ext) {
        '.ps1' { Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$p`"" }
        '.bat' { Start-Process cmd.exe -ArgumentList "/c `"$p`"" }
        '.lnk' { Start-Process -FilePath $p }
        default { try { Start-Process -FilePath $p } catch { [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna filen:","Skript2") | Out-Null } }
    }
})

$miScript3.add_Click({
    $p = $Script3Path
    if ([string]::IsNullOrWhiteSpace($p)) { [System.Windows.Forms.MessageBox]::Show("...","Skript3") | Out-Null; return }
    if (-not (Test-Path -LiteralPath $p)) { [System.Windows.Forms.MessageBox]::Show("...","Skript3") | Out-Null; return }
    $ext=[System.IO.Path]::GetExtension($p).ToLowerInvariant()
    switch ($ext) {
        '.ps1' { Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$p`"" }
        '.bat' { Start-Process cmd.exe -ArgumentList "/c `"$p`"" }
        '.lnk' { Start-Process -FilePath $p }
        default { try { Start-Process -FilePath $p } catch { [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna filen:","Skript3") | Out-Null } }
    }
})

$miToggleSign.add_Click({
    $lsp = $txtLSP.Text.Trim()
    if (-not $lsp) {
        Gui-Log "⚠️ Ange och sök ett LSP först innan du aktiverar Seal Test-signatur." 'Warn'
        return
    }
    $selNeg = Get-CheckedFilePath $clbNeg
    $selPos = Get-CheckedFilePath $clbPos
    if (-not $selNeg -or -not $selPos) {
        Gui-Log "⚠️ Du måste först välja både Seal Test NEG och POS innan Seal Test-signatur kan aktiveras." 'Warn'
        return
    }
    $grpSign.Visible = -not $grpSign.Visible
    if ($grpSign.Visible) {
        $form.Height = $baseHeight + $grpSign.Height + 40
        $miToggleSign.Text  = '❌ Dölj Seal Test-signatur'
    }
    else {
        $form.Height = $baseHeight
        $miToggleSign.Text  = '✅ Aktivera Seal Test-signatur'
    }
})

function Set-Theme {
    param([string]$Theme)
    if ($Theme -eq 'dark') {
        $global:CurrentTheme = 'dark'
        $form.BackColor        = [System.Drawing.Color]::FromArgb(35,35,35)
        $content.BackColor     = $form.BackColor
        $panelHeader.BackColor = [System.Drawing.Color]::DarkSlateBlue
        $pLog.BackColor        = [System.Drawing.Color]::FromArgb(45,45,45)
        $grpPick.BackColor     = $form.BackColor   
        $grpSign.BackColor     = $form.BackColor
        $grpSave.BackColor     = $form.BackColor
        $tlSearch.BackColor    = $form.BackColor
        $outputBox.BackColor   = [System.Drawing.Color]::FromArgb(55,55,55)
        $outputBox.ForeColor   = [System.Drawing.Color]::White
        $lblLSP.ForeColor      = [System.Drawing.Color]::White
        $lblCsv.ForeColor      = [System.Drawing.Color]::White
        $lblNeg.ForeColor      = [System.Drawing.Color]::White
        $lblPos.ForeColor      = [System.Drawing.Color]::White
        $grpPick.ForeColor     = [System.Drawing.Color]::White
        $grpSign.ForeColor     = [System.Drawing.Color]::White
        $grpSave.ForeColor     = [System.Drawing.Color]::White
        $pLog.ForeColor        = [System.Drawing.Color]::White
        $tlSearch.ForeColor    = [System.Drawing.Color]::White
    } else {
        $global:CurrentTheme = 'light'
        $form.BackColor        = [System.Drawing.Color]::WhiteSmoke
        $content.BackColor     = $form.BackColor
        $panelHeader.BackColor = [System.Drawing.Color]::SteelBlue
        $pLog.BackColor        = [System.Drawing.Color]::White
        $grpPick.BackColor     = $form.BackColor
        $grpSign.BackColor     = $form.BackColor
        $grpSave.BackColor     = $form.BackColor
        $tlSearch.BackColor    = $form.BackColor
        $outputBox.BackColor   = [System.Drawing.Color]::White
        $outputBox.ForeColor   = [System.Drawing.Color]::Black
        $lblLSP.ForeColor      = [System.Drawing.Color]::Black
        $lblCsv.ForeColor      = [System.Drawing.Color]::Black
        $lblNeg.ForeColor      = [System.Drawing.Color]::Black
        $lblPos.ForeColor      = [System.Drawing.Color]::Black
        $grpPick.ForeColor     = [System.Drawing.Color]::Black
        $grpSign.ForeColor     = [System.Drawing.Color]::Black
        $grpSave.ForeColor     = [System.Drawing.Color]::Black
        $pLog.ForeColor        = [System.Drawing.Color]::Black
        $tlSearch.ForeColor    = [System.Drawing.Color]::Black
    }
}

$miLightTheme.add_Click({ Set-Theme 'light' })
$miDarkTheme.add_Click({ Set-Theme 'dark' })

# Instruktioner

$miShowInstr.add_Click({
    $msg = @"
Snabbguide

1. Skriv in ditt LSP och klicka "Sök Filer eller använd Bläddra..."

2. Välj fil:
   • 1x CSV
   • 1x Seal Test NEG
   • 1x Seal Test POS
   • 1x Worksheet

3. Välj Rapport-utdata:
   • Spara i LSP-mapp (default)
   • Öppna endast i temporärt läge
   • Inkludera flik "SharePoint Info"

4. Klicka på "Skapa rapport"

Excelrapport öppnas med följande flikar beroende på valda filer:
   • Information (generell)
   • Seal Test Info
   • STF Sum (och minusvärden)
   • Infinity/GX 
   • Kontrollmaterial
   • SharePoint Info

Felsökning:
   • Filen låst → Stäng Excelfiler.
"@
    [System.Windows.Forms.MessageBox]::Show($msg,"Instruktioner") | Out-Null
})

$miFAQ.add_Click({
    $faq = @"
Vad gör skriptet?

Det skapar en excel-rapport som jämför sökt LSP för Seal Test-Filer,
hämtar utrustningslista, rätt kontrollmaterial och SharePoint Info för sökt LSP.

1) Varför ser jag inte fliken “SharePoint Info”?
   • Kryssrutan “SharePoint Info” måste vara ibockad.
   • Inloggning kan saknas eller SharePoint-listan innehåller inte batchnumret.

2) UI fryser ibland – är det normalt?
   • Nej. PnP-kopplingen och läsningen görs i bakgrunden. Om det ändå känns segt:
     - Testa utan SharePoint först (avbocka) för att isolera.
     - Stäng tunga Excel-instans(er) i bakgrunden.

3) “Filen är låst/kan inte spara”
   • Stäng källfilen i Excel.
   • Kontrollera att OneDrive/SharePoint Sync inte håller filen exklusivt låst.
   • Spara till TEMP för att testa att genereringen fungerar.

4) Var sparas rapporten?
   • Välj “LSP-mapp” (samma mapp som ditt LSP) eller “TEMP” = sparas inte.
"@
    [System.Windows.Forms.MessageBox]::Show($faq,"Vanliga frågor") | Out-Null
})

$miHelpDlg.add_Click({
    $helpForm = New-Object System.Windows.Forms.Form
    $helpForm.Text = 'Skicka meddelande'
    $helpForm.Size = New-Object System.Drawing.Size(400,300)
    $helpForm.StartPosition = 'CenterParent'
    $helpForm.Font = $form.Font
    $helpBox = New-Object System.Windows.Forms.TextBox
    $helpBox.Multiline = $true
    $helpBox.ScrollBars = 'Vertical'
    $helpBox.Dock = 'Fill'
    $helpBox.Font = New-Object System.Drawing.Font('Segoe UI',9)
    $helpBox.Margin = New-Object System.Windows.Forms.Padding(10)
    $panelButtons = New-Object System.Windows.Forms.FlowLayoutPanel
    $panelButtons.Dock = 'Bottom'
    $panelButtons.FlowDirection = 'RightToLeft'
    $panelButtons.Padding = New-Object System.Windows.Forms.Padding(10)
    $btnSend = New-Object System.Windows.Forms.Button
    $btnSend.Text = 'Skicka'
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Avbryt'
    $panelButtons.Controls.Add($btnSend)
    $panelButtons.Controls.Add($btnCancel)
    $helpForm.Controls.Add($helpBox)
    $helpForm.Controls.Add($panelButtons)
    $btnSend.Add_Click({
        $msg = $helpBox.Text.Trim()
        if (-not $msg) { [System.Windows.Forms.MessageBox]::Show('Ange ett meddelande innan du skickar.','Hjälp') | Out-Null; return }
        try {
            $helpDir = Join-Path $PSScriptRoot 'help'
            if (-not (Test-Path $helpDir)) { New-Item -ItemType Directory -Path $helpDir -Force | Out-Null }
            $ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
            $file = Join-Path $helpDir "help_${ts}.txt"
            Set-Content -Path $file -Value $msg -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show('Meddelandet sparades. Tack!','Hjälp') | Out-Null
            $helpForm.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Kunde inte spara meddelandet:\n$($_.Exception.Message)",'Hjälp') | Out-Null
        }
    })
    $btnCancel.Add_Click({ $helpForm.Close() })
    $helpForm.ShowDialog() | Out-Null
})

# Om
$miOm.add_Click({ [System.Windows.Forms.MessageBox]::Show("OBS! Detta verktyg är endast ett hjälpmedel och ersätter inte någon process hos PQC.`n $ScriptVersion`nav Jesper","Om") | Out-Null })

$btnScan.Add_Click({
    if (-not (Assert-StartupReady)) { return }
    Gui-Log '🔎 Söker filer…' -Immediate
    try {
    $lsp = $txtLSP.Text.Trim()
    if (-not $lsp) { Gui-Log "⚠️ Ange ett LSP-nummer" 'Warn'; return }

    if ($script:LastScanResult -and $script:LastScanResult.Lsp -eq $lsp -and $script:LastScanResult.Folder -and (Test-Path -LiteralPath $script:LastScanResult.Folder)) {
        Gui-Log "♻️ Återanvänder senaste sökresultatet för $lsp." 'Info'
        Add-CLBItems -clb $clbCsv -files $script:LastScanResult.Csv -AutoCheckFirst
        Add-CLBItems -clb $clbNeg -files $script:LastScanResult.Neg -AutoCheckFirst
        Add-CLBItems -clb $clbPos -files $script:LastScanResult.Pos -AutoCheckFirst
        Add-CLBItems -clb $clbLsp -files $script:LastScanResult.LspFiles -AutoCheckFirst
        Update-BuildEnabled
        Update-BatchLink
        return
    }

    $folder = $null
    foreach ($path in $RootPaths) {
        $folder = Get-ChildItem $path -Directory -Recurse -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -match "#?$lsp" } |
                  Select-Object -First 1
        if ($folder) { break }
    }
    if (-not $folder) { Gui-Log "❌ Ingen LSP-mapp hittad för $lsp" 'Warn'; return }

    $files = Get-ChildItem $folder.FullName -File -ErrorAction SilentlyContinue
    $candCsv = $files | Where-Object { $_.Extension -ieq '.csv' -and ( $_.Name -match $lsp -or $_.Length -gt 100kb ) } | Sort-Object LastWriteTime -Descending
    $candNeg = $files | Where-Object { $_.Name -match '(?i)Neg.*\.xls[xm]$' -and $_.Name -match $lsp } | Sort-Object LastWriteTime -Descending
    $candPos = $files | Where-Object { $_.Name -match '(?i)Pos.*\.xls[xm]$' -and $_.Name -match $lsp } | Sort-Object LastWriteTime -Descending
    $candLsp = $files | Where-Object {
        ($_.Name -match '(?i)worksheet') -and ($_.Name -match [regex]::Escape($lsp)) -and ($_.Extension -match '^(\.xlsx|\.xlsm|\.xls)$')
    } | Sort-Object LastWriteTime -Descending
    Gui-Log "📂 Hittad mapp: $($folder.FullName)" 'Info'
    Add-CLBItems -clb $clbCsv -files $candCsv -AutoCheckFirst
    Add-CLBItems -clb $clbNeg -files $candNeg -AutoCheckFirst
    Add-CLBItems -clb $clbPos -files $candPos -AutoCheckFirst
    Add-CLBItems -clb $clbLsp -files $candLsp -AutoCheckFirst
    if ($candCsv.Count -eq 0) { Gui-Log "ℹ️ Ingen CSV hittad (endast .csv visas)." 'Info' }
    if ($candNeg.Count -eq 0) { Gui-Log "⚠️ Ingen Seal NEG hittad." 'Warn' }
    if ($candPos.Count -eq 0) { Gui-Log "⚠️ Ingen Seal POS hittad." 'Warn' }
    if ($candLsp.Count -eq 0) { Gui-Log "ℹ️ Ingen LSP Worksheet hittad." 'Info' }
    Update-BuildEnabled
    Update-BatchLink
    $script:LastScanResult = [pscustomobject]@{
        Lsp      = $lsp
        Folder   = $folder.FullName
        Csv      = $candCsv  | ForEach-Object { $_.FullName }
        Neg      = $candNeg  | ForEach-Object { $_.FullName }
        Pos      = $candPos  | ForEach-Object { $_.FullName }
        LspFiles = $candLsp  | ForEach-Object { $_.FullName }
    }
    } finally {
        Gui-Log '✅ Filer laddade'
    }
})

$btnCsvBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "CSV|*.csv|Alla filer|*.*"
    if ($dlg.ShowDialog() -eq 'OK') {
        $f = Get-Item -LiteralPath $dlg.FileName
        Add-CLBItems -clb $clbCsv -files @($f) -AutoCheckFirst
        Update-BuildEnabled
        Update-BatchLink
    }
})

$btnNegBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Excel|*.xlsx;*.xlsm|Alla filer|*.*"
    if ($dlg.ShowDialog() -eq 'OK') {
        $f = Get-Item -LiteralPath $dlg.FileName
        Add-CLBItems -clb $clbNeg -files @($f) -AutoCheckFirst
        Update-BuildEnabled
        Update-BatchLink
    }
})

$btnPosBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Excel|*.xlsx;*.xlsm|Alla filer|*.*"
    if ($dlg.ShowDialog() -eq 'OK') {
        $f = Get-Item -LiteralPath $dlg.FileName
        Add-CLBItems -clb $clbPos -files @($f) -AutoCheckFirst
        Update-BuildEnabled
        Update-BatchLink
    }
})

 

if (-not (Get-Command Write-SPSheet-Safe -ErrorAction SilentlyContinue)) {
    function Write-SPSheet-Safe {
        param(
            [OfficeOpenXml.ExcelPackage]$Pkg,
            [object]$Rows,                    
            [string[]]$DesiredOrder,            
            [string]$Batch
        )
        if (-not $Pkg) { return $false }
        $Rows = @($Rows)
        $name = "SharePoint Info"
        $wsOld = $Pkg.Workbook.Worksheets[$name]
        if ($wsOld) { $Pkg.Workbook.Worksheets.Delete($wsOld) }
        $ws = $Pkg.Workbook.Worksheets.Add($name)
        if ($Rows.Count -eq 0 -or $Rows[0] -eq $null) {
            $ws.Cells[1,1].Value = "No rows found (Batch=$Batch)"
            return $true
        }
        $isKV = ($Rows[0].psobject.Properties.Name -contains 'Rubrik') -and `
                ($Rows[0].psobject.Properties.Name -contains 'Värde')
        if ($isKV) {
            $ws.Cells[1,1].Value = "SharePoint Information"
            $ws.Cells[1,2].Value = ""
            $ws.Cells["A1:B1"].Merge = $true
            $ws.Cells["A1"].Style.Font.Bold = $true
            $ws.Cells["A1"].Style.Font.Size = 12
            $ws.Cells["A1"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
            $ws.Cells["A1"].Style.Fill.PatternType = "Solid"
            $ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkBlue)
            $ws.Cells["A1"].Style.HorizontalAlignment = "Center"
            $ws.Cells["A1"].Style.VerticalAlignment   = "Center"

            $r = 2
            foreach ($row in $Rows) {
                $ws.Cells[$r,1].Value = $row.Rubrik
                $ws.Cells[$r,2].Value = $row.'Värde'
                $r++
            }
            $lastRow = $r-1
            $ws.Cells["A2:A$lastRow"].Style.Font.Bold = $true
            $ws.Cells["A2:A$lastRow"].Style.Fill.PatternType = "Solid"
            $ws.Cells["A2:A$lastRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Gainsboro)
            $ws.Cells["B2:B$lastRow"].Style.Fill.PatternType = "Solid"
            $ws.Cells["B2:B$lastRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::WhiteSmoke)

            $rng = $ws.Cells["A1:B$lastRow"]
            $rng.Style.Font.Name = "Arial"
            $rng.Style.Font.Size = 10
            $rng.Style.HorizontalAlignment = "Left"
            $rng.Style.VerticalAlignment   = "Center"
            $rng.Style.Border.Top.Style    = "Thin"
            $rng.Style.Border.Bottom.Style = "Thin"
            $rng.Style.Border.Left.Style   = "Thin"
            $rng.Style.Border.Right.Style  = "Thin"
            $rng.Style.Border.BorderAround("Medium")
            try { $rng.AutoFitColumns() | Out-Null } catch {}
        }
        else {
            $cols = @()
            if ($DesiredOrder) { $cols += $DesiredOrder }
            foreach ($k in $Rows[0].psobject.Properties.Name) {
                if ($cols -notcontains $k) { $cols += $k }
            }

            for ($c=0; $c -lt $cols.Count; $c++) {
                $ws.Cells[1,$c+1].Value = $cols[$c]
                $ws.Cells[1,$c+1].Style.Font.Bold = $true
            }
            $r = 2
            foreach ($row in $Rows) {
                for ($c=0; $c -lt $cols.Count; $c++) {
                    $ws.Cells[$r,$c+1].Value = $row.$($cols[$c])
                }
                $r++
            }
            try {
                if ($ws.Dimension) {
                    $maxR = [Math]::Min($ws.Dimension.End.Row, 2000)
                    $ws.Cells[$ws.Dimension.Start.Row,$ws.Dimension.Start.Column,$maxR,$ws.Dimension.End.Column].AutoFitColumns() | Out-Null
                }
            } catch {}
        }
        return $true
    }
}

# ============================
# ===== RAPPORTLOGIK =========
# ============================

$btnBuild.Add_Click({
    if (-not (Assert-StartupReady)) { return }
    Gui-Log 'Skapar rapport…' -Immediate

    $runConfig = [pscustomobject]@{
        CsvPath               = Get-CheckedFilePath $clbCsv
        SealNegPath           = Get-CheckedFilePath $clbNeg
        SealPosPath           = Get-CheckedFilePath $clbPos
        WorksheetPath         = Get-CheckedFilePath $clbLsp
        Lsp                   = ($txtLSP.Text + '').Trim()
        WriteSignatures       = $chkWriteSign.Checked
        OverwriteSignature    = $chkOverwriteSign.Checked
        SignatureText         = ($txtSigner.Text + '').Trim()
        SaveInLsp             = $rbSaveInLsp.Checked
        IncludeSharePointInfo = if ($chkSharePointInfo) { $chkSharePointInfo.Checked } else { $true }
        ReportOptions         = $global:ReportOptions
        TemplatePath          = Join-Path $PSScriptRoot "output_template-v4.xlsx"
        RawDataPath           = $RawDataPath
        UtrustningListPath    = $UtrustningListPath
        ScriptRoot            = $PSScriptRoot
        ScriptVersion         = $ScriptVersion
    }

    if ($runConfig.WriteSignatures) {
        if (-not $runConfig.SignatureText) {
            Gui-Log "❌ Ingen signatur angiven (B47). Avbryter." 'Error'
            return
        }
        if (-not (Confirm-SignatureInput -Text $runConfig.SignatureText)) {
            Gui-Log "🛑 Signatur ej bekräftad. Avbryter." 'Warn'
            return
        }
    }

    $result = Invoke-ValidateAssayRun -RunConfig $runConfig -Logger { param($Message,$Severity) Gui-Log $Message $Severity }

    foreach ($err in $result.Errors) { if ($err) { Gui-Log ("❌ {0}" -f $err) 'Error' } }
    foreach ($warn in $result.Warnings) { if ($warn) { Gui-Log ("⚠️ {0}" -f $warn) 'Warn' } }

    if (-not $result.Ok) { return }

    if ($result.Data -and $result.Data.ReportPath) {
        try { Start-Process -FilePath "excel.exe" -ArgumentList "`"$($result.Data.ReportPath)`"" } catch {}
    }
})

#endregion Event Handlers

# === Tooltip-inställningar ===
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 8000
$toolTip.InitialDelay = 500
$toolTip.ReshowDelay  = 500
$toolTip.ShowAlways   = $true
$toolTip.SetToolTip($txtLSP, 'Ange LSP-numret utan ”#” och klicka på Sök filer.')
$toolTip.SetToolTip($btnScan, 'Sök efter LSP och lista tillgängliga filer.')
$toolTip.SetToolTip($clbCsv,  'Välj CSV-fil.')
$toolTip.SetToolTip($clbNeg,  'Välj Seal Test Neg-fil.')
$toolTip.SetToolTip($clbPos,  'Välj Seal Test Pos-fil.')
$toolTip.SetToolTip($btnCsvBrowse, 'Bläddra efter en CSV-fil manuellt.')
$toolTip.SetToolTip($btnNegBrowse, 'Bläddra efter Seal Test Neg-fil manuellt.')
$toolTip.SetToolTip($btnPosBrowse, 'Bläddra efter Seal Test Pos-fil manuellt.')
$toolTip.SetToolTip($txtSigner, 'Skriv fullständigt namn, signatur och datum (separerat med kommatecken).')
$toolTip.SetToolTip($chkWriteSign, 'Signatur appliceras på flikar.')
$toolTip.SetToolTip($chkOverwriteSign, 'Dubbelkontroll för att aktivera signering')
$miToggleSign.ToolTipText = 'Visa eller dölj panelen för att lägga till signatur.'
$toolTip.SetToolTip($rbSaveInLsp, 'Spara rapporten i mappen för ditt LSP.')
$toolTip.SetToolTip($rbTempOnly, 'Skapa rapporten temporär utan att spara.')
$toolTip.SetToolTip($btnBuild, 'Skapa och öppna rapporten baserat på de valda filerna.')
$toolTip.SetToolTip($chkSharePointInfo, 'Exportera med SharePoint Info.')
$txtLSP.add_TextChanged({ Update-BatchLink })

#region Main Run / Orchestration
# =============== SLUT ===============
function Enable-DoubleBuffer {
    $pi = [Windows.Forms.Control].GetProperty('DoubleBuffered',[Reflection.BindingFlags]'NonPublic,Instance')
    foreach($c in @($content,$pLog,$grpPick,$grpSign,$grpSave)) { if ($c) { $pi.SetValue($c,$true,$null) } }
}
try { Set-Theme 'light' } catch {}
Enable-DoubleBuffer
Update-BatchLink
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)

try{ Stop-Transcript | Out-Null }catch{}
#endregion Main Run / Orchestration
