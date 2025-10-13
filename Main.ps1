# ===============================
# DocMerge.ps1  (v8.8)
# ===============================

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Start-Process -FilePath $PSHome\powershell.exe -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -File `"$PSCommandPath`""; exit
}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing        # Ladda in GDI+-typer för grafik
Add-Type -AssemblyName System.ComponentModel  # För BackgroundWorker
try {
    # Det här assemblyt behövs för TextFieldParser (för CSV-läsning)
    Add-Type -AssemblyName 'Microsoft.VisualBasic' -ErrorAction SilentlyContinue
} catch {}

# === Inställningar ===
$ScriptVersion = "v8.8"
$PSScriptRoot  = Split-Path -Parent $MyInvocation.MyCommand.Path

$RootPaths = @(
    "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Tests",
    "N:\QC\QC-1\IPT\3. IPT - KLART FÖR SAMMANSTÄLLNING",
    "N:\QC\QC-1\IPT\4. IPT - KLART FÖR GRANSKNING"
)
$ikonSokvag = Join-Path $PSScriptRoot "icon.png"

$UtrustningListPath = "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project\Utrustninglista2.0.xlsx"
$RawDataPath        = "N:\QC\QC-1\IPT\KONTROLLPROVSFIL - Version 2.4.xlsm"
$SlangAssayPath     = "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project\slangassay.xlsx"

$OtherScriptPath = ''

$Script1Path  = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Main\Batch-Search\Batch-Search.ps1'
$Script2Path  = 'N:\QC\QC-1\IPT\1. IPT - KOMMANDE TESTER\Mappscript\1. mappscript\Script.ps1'
$Script3Path  = ''

# === Dynamisk direktlänk till SharePoint (byggs från Batch# och ev. LSP) ===
# Ersätt med er faktiska vy/filtrering i SharePoint. {BatchNumber} och {LSP} ersätts automatiskt.
$SharePointBatchLinkTemplate = 'https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management/Lists/Cepheid%20%20Production%20orders/ROBAL.aspx?viewid=6c9e53c9-a377-40c1-a154-13a13866b52b&view=7&q={BatchNumber}'

# === Logg: alltid till $PSScriptRoot\Loggar ===
$DevLogDir = Join-Path $PSScriptRoot 'Loggar'
if (-not (Test-Path $DevLogDir)) { New-Item -ItemType Directory -Path $DevLogDir -Force | Out-Null }
$global:LogPath = Join-Path $DevLogDir ("DocMerge_{0:yyyyMMdd_HHmmss}.txt" -f (Get-Date))

# === Ikoner ===
function New-GlyphIcon {
    param(
        [ValidateSet('folder','search','report','tools','settings','help','info','open','exit')]
        [string]$Kind,[int]$Size=20,[string]$Stroke='#34495E',[single]$PenW=1.8
    )
    $bmp = New-Object System.Drawing.Bitmap $Size,$Size,([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g=[System.Drawing.Graphics]::FromImage($bmp); $g.SmoothingMode='AntiAlias'; $g.Clear([System.Drawing.Color]::Transparent)
    $pen=New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml($Stroke)),$PenW
    $cx=$Size/2.0; $cy=$Size/2.0
    switch($Kind){
        'search' {$r=$Size*.32; $g.DrawEllipse($pen,$cx-$r,$cy-$r,2*$r,2*$r); $p1=New-Object Drawing.PointF ([single]($cx+$r*.7)),([single]($cy+$r*.7)); $p2=New-Object Drawing.PointF ([single]($p1.X+$Size*.22)),([single]($p1.Y+$Size*.22)); $g.DrawLine($pen,$p1,$p2)}
        'report' {$g.DrawRectangle($pen,4,3,$Size-8,$Size-6); 6,7,11,15 | % { $g.DrawLine($pen,6,$_,$Size-10,$_)} }
        'folder' {$g.DrawRectangle($pen,3,8,$Size-6,$Size-12); $g.DrawLine($pen,6,8,10,4); $g.DrawLine($pen,10,4,16,4); $g.DrawLine($pen,16,4,16,8)}
        'tools'  {$r=$Size*.18; $g.DrawArc($pen,$cx-$r,4,2*$r,2*$r,200,220); $g.DrawLine($pen,$cx,$Size*.18,$Size-5,$Size-5); $g.DrawEllipse($pen,$Size-7,$Size-7,3,3)}
        'settings'{$y=[int]$cy; $g.DrawLine($pen,3,$y,$Size-3,$y); $g.DrawEllipse($pen,$cx-4,$y-4,8,8)}
        'help'   {$g.DrawEllipse($pen,3,3,$Size-6,$Size-6); $g.DrawCurve($pen,@( (New-Object Drawing.PointF ([single]($cx-3)),([single]($cy-2))), (New-Object Drawing.PointF ([single]$cx),([single]($cy-5))), (New-Object Drawing.PointF ([single]($cx+3)),([single]($cy-2))) )); $g.DrawLine($pen,$cx,$cy,$cx,$cy+4)}
        'info'   {$g.DrawEllipse($pen,3,3,$Size-6,$Size-6); $g.DrawLine($pen,$cx,$cy-2,$cx,$cy+5); $g.DrawEllipse($pen,$cx-0.8,$cy-6.8,1.6,1.6)}
        'open'   {$g.DrawRectangle($pen,4,6,$Size-12,$Size-10); $g.DrawLine($pen,$Size-8,$cy,$Size-4,$cy); $g.DrawLine($pen,$Size-7,$cy-3,$Size-4,$cy); $g.DrawLine($pen,$Size-7,$cy+3,$Size-4,$cy)}
        'exit'   {$m=5; $g.DrawLine($pen,$m,$m,$Size-$m,$Size-$m); $g.DrawLine($pen,$Size-$m,$m,$m,$Size-$m)}
    }
    $pen.Dispose(); $g.Dispose(); return $bmp
}

# === Genvägar (meny) ===
function Add-ShortcutItem {
    param(
        [System.Windows.Forms.ToolStripMenuItem]$Parent,
        [string]$Text,
        [string]$Target
    )
    $it = New-Object System.Windows.Forms.ToolStripMenuItem($Text)
    $it.Tag = $Target

    if ($Target -match '^(?i)https?://') { $it.Image = New-GlyphIcon -Kind 'open' }
    elseif (Test-Path -LiteralPath $Target) {
        try {
            $gi = Get-Item -LiteralPath $Target -ErrorAction Stop
            $it.Image = if ($gi.PSIsContainer) { New-GlyphIcon -Kind 'folder' } else { New-GlyphIcon -Kind 'report' }
        } catch { $it.Image = New-GlyphIcon -Kind 'open' }
    } else { $it.Image = New-GlyphIcon -Kind 'open' }

    $it.add_Click({
        param($s,$e)
        $t = [string]$s.Tag
        try {
            if ($t -match '^(?i)https?://') { Start-Process $t }
            elseif (Test-Path -LiteralPath $t) {
                $gi = Get-Item -LiteralPath $t
                if ($gi.PSIsContainer) { Start-Process explorer.exe -ArgumentList "`"$t`"" }
                else { Start-Process -FilePath $t }
            } else { [System.Windows.Forms.MessageBox]::Show("Hittar inte sökvägen:`n$t","Genväg",'OK','Warning') | Out-Null }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna:`n$t`n$($_.Exception.Message)","Genväg") | Out-Null
        }
    })
    [void]$Parent.DropDownItems.Add($it)
}

# --- Accentfärg & knappstil ---
function Get-WinAccentColor {
    try {
        $p = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\DWM' -ErrorAction Stop
        $argb = if($p.AccentColor){$p.AccentColor}elseif($p.ColorizationColor){$p.ColorizationColor}else{$null}
        if($argb){ return [System.Drawing.Color]::FromArgb([int]$argb) }
    } catch {}
    return [System.Drawing.Color]::FromArgb(38,120,178)
}
function New-Color { param([int]$A,[int]$R,[int]$G,[int]$B) [System.Drawing.Color]::FromArgb($A,$R,$G,$B) }
function Darken  { param([System.Drawing.Color]$c,[double]$f=0.85) New-Color 255 ([int]($c.R*$f)) ([int]($c.G*$f)) ([int]($c.B*$f)) }
function Lighten { param([System.Drawing.Color]$c,[double]$f=0.12) New-Color 255 ([int]([Math]::Min(255,$c.R+(255-$c.R)*$f))) ([int]([Math]::Min(255,$c.G+(255-$c.G)*$f))) ([int]([Math]::Min(255,$c.B+(255-$c.B)*$f))) }
$Accent=Get-WinAccentColor; $AccentBorder=Darken $Accent 0.75; $AccentHover=Lighten $Accent 0.12; $AccentDisabled=New-Color 255 210 210 210
function Set-AccentButton {
    param([System.Windows.Forms.Button]$Btn,[switch]$Primary)
    $Btn.FlatStyle='Flat'
    $Btn.FlatAppearance.BorderSize=1
    $Btn.FlatAppearance.BorderColor=$AccentBorder
    $Btn.FlatAppearance.MouseOverBackColor=$AccentHover
    if($Primary){ $Btn.BackColor=$Accent; $Btn.ForeColor=[System.Drawing.Color]::White; $Btn.UseVisualStyleBackColor=$false }
    else{ $Btn.BackColor=[System.Drawing.Color]::White; $Btn.ForeColor=[System.Drawing.Color]::Black; $Btn.UseVisualStyleBackColor=$false }
    if($Btn.Height -lt 30){ $Btn.Height=30 }
    $Btn.add_EnabledChanged({
        if($this.Enabled){ if($Primary){$this.BackColor=$Accent; $this.ForeColor=[System.Drawing.Color]::White}else{$this.BackColor=[System.Drawing.Color]::White; $this.ForeColor=[System.Drawing.Color]::Black} }
        else{ $this.BackColor=$AccentDisabled; $this.ForeColor=[System.Drawing.Color]::Gray }
    })
}

# ---------- Form ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "DocMerge $ScriptVersion"
$form.AutoScaleMode = 'Dpi'
$form.Size = New-Object System.Drawing.Size(860,860)
$form.MinimumSize = New-Object System.Drawing.Size(860,860)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.AutoScroll  = $false
$form.MaximizeBox = $false
$form.Padding     = New-Object System.Windows.Forms.Padding(8)
$form.Font        = New-Object System.Drawing.Font('Segoe UI',10)

# ---------- Menyrad ----------
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.Dock='Top'; $menuStrip.GripStyle='Hidden'
$menuStrip.ImageScalingSize = New-Object System.Drawing.Size(20,20)
$menuStrip.Padding = New-Object System.Windows.Forms.Padding(8,6,0,6)
$menuStrip.Font = New-Object System.Drawing.Font('Segoe UI',10)

$miArkiv   = New-Object System.Windows.Forms.ToolStripMenuItem('Arkiv')
$miVerktyg = New-Object System.Windows.Forms.ToolStripMenuItem('Verktyg')
$miSettings= New-Object System.Windows.Forms.ToolStripMenuItem('Inställningar')
$miHelp    = New-Object System.Windows.Forms.ToolStripMenuItem('Instruktioner')
$miAbout   = New-Object System.Windows.Forms.ToolStripMenuItem('Om')

$miScan  = New-Object System.Windows.Forms.ToolStripMenuItem('Sök filer')
$miBuild = New-Object System.Windows.Forms.ToolStripMenuItem('Skapa rapport')
$miExit  = New-Object System.Windows.Forms.ToolStripMenuItem('Avsluta')

# Rensa ev. gamla undermenyer
$miArkiv.DropDownItems.Clear()
$miVerktyg.DropDownItems.Clear()
$miSettings.DropDownItems.Clear()
$miHelp.DropDownItems.Clear()

# ----- Arkiv -----
$miNew         = New-Object System.Windows.Forms.ToolStripMenuItem('Nytt')
$miOpenRecent  = New-Object System.Windows.Forms.ToolStripMenuItem('Öppna senaste rapport')
$miArkiv.DropDownItems.AddRange(@(
    $miNew,
    $miOpenRecent,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $miExit
))
$miNew.Image        = New-GlyphIcon -Kind 'report'
$miOpenRecent.Image = New-GlyphIcon -Kind 'open'

# ----- Verktyg -----
$miScript1   = New-Object System.Windows.Forms.ToolStripMenuItem('Sök Batch')
$miScript2   = New-Object System.Windows.Forms.ToolStripMenuItem('Skapa Mapp')
$miScript3   = New-Object System.Windows.Forms.ToolStripMenuItem('TBD')
$miToggleSign = New-Object System.Windows.Forms.ToolStripMenuItem('Aktivera Seal Test-signatur')
$miVerktyg.DropDownItems.AddRange(@(
    $miScript1,
    $miScript2,
    $miScript3,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $miToggleSign
))
$miScript1.Image   = New-GlyphIcon -Kind 'report'
$miScript2.Image   = New-GlyphIcon -Kind 'report'
$miScript3.Image   = New-GlyphIcon -Kind 'report'

# ----- Inställningar -----
$miTheme = New-Object System.Windows.Forms.ToolStripMenuItem('Tema')
$miLightTheme = New-Object System.Windows.Forms.ToolStripMenuItem('☀️ Ljust (default)')
$miDarkTheme  = New-Object System.Windows.Forms.ToolStripMenuItem('🌙 Mörkt')
$miTheme.DropDownItems.AddRange(@($miLightTheme,$miDarkTheme))
$miTheme.Image = New-GlyphIcon -Kind 'settings'
$miSettings.DropDownItems.Add($miTheme)

# ----- Instruktioner -----
$miShowInstr   = New-Object System.Windows.Forms.ToolStripMenuItem('Visa instruktioner')
$miFAQ         = New-Object System.Windows.Forms.ToolStripMenuItem('Vanliga frågor (FAQ)')
$miHelpDlg     = New-Object System.Windows.Forms.ToolStripMenuItem('Hjälp')
$miHelp.DropDownItems.AddRange(@($miShowInstr,$miFAQ,$miHelpDlg))

$miShowInstr.Image = New-GlyphIcon -Kind 'help'
$miFAQ.Image       = New-GlyphIcon -Kind 'info'
$miHelpDlg.Image   = New-GlyphIcon -Kind 'help'

$miGenvagar = New-Object System.Windows.Forms.ToolStripMenuItem('Genvägar'); $miGenvagar.Image = New-GlyphIcon -Kind 'open'
$ShortcutGroups = @{
    'IPT-mappar' = @(
        @{ Text='IPT - PÅGÅENDE KÖRNINGAR';        Target='N:\QC\QC-1\IPT\2. IPT - PÅGÅENDE KÖRNINGAR' },
        @{ Text='IPT - KLART FÖR SAMMANSTÄLLNING'; Target='N:\QC\QC-1\IPT\3. IPT - KLART FÖR SAMMANSTÄLLNING' },
        @{ Text='IPT - KLART FÖR GRANSKNING';      Target='N:\QC\QC-1\IPT\4. IPT - KLART FÖR GRANSKNING' },
        @{ Text='SPT Macro Assay';                 Target='N:\QC\QC-0\SPT\SPT macros\Assay' }
    )
    'Dokument' = @(
        @{ Text='Utrustningslista';    Target=$UtrustningListPath },
        @{ Text='Kontrollprovsfil';    Target=$RawDataPath }
    )
    'Länkar' = @(
        @{ Text='IPT App';              Target='https://apps.powerapps.com/play/e/default-771c9c47-7f24-44dc-958e-34f8713a8394/a/fd340dbd-bbbf-470b-b043-d2af4cb62c83' },
        @{ Text='MES';                  Target='http://mes.cepheid.pri/camstarportal/?domain=CEPHEID.COM' },
        @{ Text='CSV Uploader';         Target='http://auw2wgxtpap01.cepaws.com/Welcome.aspx' },
        @{ Text='BMRAM';                Target='https://cepheid62468.coolbluecloud.com/' },
        @{ Text='Agile';                Target='https://agileprod.cepheid.com/Agile/default/login-cms.jsp' }
    )
}
foreach ($grp in $ShortcutGroups.GetEnumerator()) {
    $grpMenu = New-Object System.Windows.Forms.ToolStripMenuItem($grp.Key)
    $grpMenu.Image = New-GlyphIcon -Kind 'folder'
    foreach ($entry in $grp.Value) { Add-ShortcutItem -Parent $grpMenu -Text $entry.Text -Target $entry.Target }
    [void]$miGenvagar.DropDownItems.Add($grpMenu)
}
$miOm = New-Object System.Windows.Forms.ToolStripMenuItem('Om det här verktyget'); $miAbout.DropDownItems.Add($miOm)

$miArkiv.Image     = New-GlyphIcon -Kind folder
$miVerktyg.Image   = New-GlyphIcon -Kind tools
$miSettings.Image  = New-GlyphIcon -Kind settings
$miHelp.Image      = New-GlyphIcon -Kind help
$miAbout.Image     = New-GlyphIcon -Kind info
$miExit.Image      = New-GlyphIcon -Kind exit
$miOm.Image        = New-GlyphIcon -Kind 'info'

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
$lblTitle.Text="DocMerge $ScriptVersion"
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

# ---------- Loggpanel ----------
$pLog = New-Object System.Windows.Forms.Panel
$pLog.Dock='Top'; $pLog.Height=220; $pLog.Padding=New-Object System.Windows.Forms.Padding(0,0,0,8)

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline=$true; $outputBox.ScrollBars='Vertical'; $outputBox.ReadOnly=$true
$outputBox.BackColor='White'; $outputBox.Dock='Fill'
$outputBox.Font = New-Object System.Drawing.Font('Segoe UI',9)
$pLog.Controls.Add($outputBox)

# ---------- Välj filer ----------
$grpPick = New-Object System.Windows.Forms.GroupBox
$grpPick.Text='Välj filer för rapport'
$grpPick.Dock='Top'
$grpPick.Padding = New-Object System.Windows.Forms.Padding(10,12,10,14)
$grpPick.AutoSize=$false
$grpPick.Height = (78*3) + $grpPick.Padding.Top + $grpPick.Padding.Bottom + 40

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

# CSV
$lblCsv=$null;$clbCsv=$null;$btnCsvBrowse=$null
New-ListRow -labelText 'CSV:' -lbl ([ref]$lblCsv) -clb ([ref]$clbCsv) -btn ([ref]$btnCsvBrowse)
# NEG
$lblNeg=$null;$clbNeg=$null;$btnNegBrowse=$null
New-ListRow -labelText 'Seal NEG:' -lbl ([ref]$lblNeg) -clb ([ref]$clbNeg) -btn ([ref]$btnNegBrowse)
# POS
$lblPos=$null;$clbPos=$null;$btnPosBrowse=$null
New-ListRow -labelText 'Seal POS:' -lbl ([ref]$lblPos) -clb ([ref]$clbPos) -btn ([ref]$btnPosBrowse)

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

# --- NYTT: klickbar SharePoint-länk ---
$slBatchLink = New-Object System.Windows.Forms.ToolStripStatusLabel
$slBatchLink.IsLink   = $true
$slBatchLink.Text     = 'SharePoint: —'
$slBatchLink.Enabled  = $false
$slBatchLink.Tag      = $null
$slBatchLink.ToolTipText = 'Direktlänk aktiveras när Batch# hittas i POS/NEG och inte är mismatch.'
$slBatchLink.add_Click({
    if ($this.Enabled -and $this.Tag) {
        try { Start-Process $this.Tag } catch {
            [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna:`n$($this.Tag)`n$($_.Exception.Message)","Länk") | Out-Null
        }
    }
})

$status.Items.AddRange(@($slCount,$slSpacer,$slBatchLink))

# ================= ToolStripContainer-layout =================
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

# Enter = "Sök filer"
$form.AcceptButton = $btnScan

# === Logg ===
function Gui-Log {
    param([string] $Text,[ValidateSet('Info','Warn','Error')][string] $Severity = 'Info')
    $prefix = switch ($Severity) { 'Warn' {'⚠️'} 'Error' {'❌'} default {'ℹ️'} }
    $timestamp = (Get-Date).ToString('HH:mm:ss')
    $line = "[$timestamp] $prefix $Text"
    $outputBox.AppendText("$line`r`n")
    $outputBox.Refresh()
    if ($global:LogPath) { Add-Content -Path $global:LogPath -Value $line }
}

# === EPPlus ===
function Ensure-EPPlus {
    param(
        [string] $Version = "4.5.3.3",
        [string] $SourceDllPath = "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Modules\EPPlus\EPPlus.4.5.3.3\lib\net35\EPPlus.dll",
        [string] $LocalFolder = "$env:TEMP\EPPlus"
    )
    $candidatePaths = @()
    if ($SourceDllPath) { $candidatePaths += $SourceDllPath }
    $localScriptDll = Join-Path $PSScriptRoot 'EPPlus.dll'
    $candidatePaths += $localScriptDll

    $userModRoot = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules'
    if (Test-Path $userModRoot) {
        Get-ChildItem -Path (Join-Path $userModRoot 'EPPlus') -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $candidatePaths += Join-Path $_.FullName 'lib\net45\EPPlus.dll'
            $candidatePaths += Join-Path $_.FullName 'lib\net40\EPPlus.dll'
        }
    }

    $progFiles = $env:ProgramFiles
    $systemModRoot = Join-Path $progFiles 'WindowsPowerShell\Modules'
    if (Test-Path $systemModRoot) {
        Get-ChildItem -Path (Join-Path $systemModRoot 'EPPlus') -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $candidatePaths += Join-Path $_.FullName 'lib\net45\EPPlus.dll'
            $candidatePaths += Join-Path $_.FullName 'lib\net40\EPPlus.dll'
        }
    }

    foreach ($cand in $candidatePaths) {
        if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path -LiteralPath $cand)) { return $cand }
    }

    $nugetUrl = "https://www.nuget.org/api/v2/package/EPPlus/$Version"
    try {
        $guid = [Guid]::NewGuid().ToString()
        $tempDir = Join-Path $env:TEMP "EPPlus_$guid"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        $zipPath  = Join-Path $tempDir 'EPPlus.zip'

        $reqParams = @{ Uri = $nugetUrl; OutFile = $zipPath; UseBasicParsing = $true; Headers = @{ 'User-Agent' = 'DocMerge/1.0' } }
        Invoke-WebRequest @reqParams -ErrorAction Stop | Out-Null

        if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'System.IO.Compression.FileSystem' })) {
            Add-Type -AssemblyName 'System.IO.Compression.FileSystem' -ErrorAction SilentlyContinue
        }
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $tempDir)

        $extractedRoot = Join-Path $tempDir 'lib'
        if (Test-Path $extractedRoot) {
            $dllCandidates = Get-ChildItem -Path (Join-Path $extractedRoot 'net45'), (Join-Path $extractedRoot 'net40') -Filter 'EPPlus.dll' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($dllCandidates) { return $dllCandidates.FullName }
        }
    } catch {
        Write-Warning "❌ EPPlus: Kunde inte hämta EPPlus ($Version): $($_.Exception.Message)"
    }
    Write-Warning "❌ EPPlus.dll hittades inte. Installera EPPlus $Version manuellt."
    return $null
}

function Load-EPPlus {
    if ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'EPPlus' }) { return $true }
    $dllPath = Ensure-EPPlus -Version '4.5.3.3'
    if (-not $dllPath) { return $false }
    try {
        $bytes = [System.IO.File]::ReadAllBytes($dllPath)
        [System.Reflection.Assembly]::Load($bytes) | Out-Null
        return $true
    } catch {
        Write-Warning "❌ EPPlus-fel: $($_.Exception.Message)"
        return $false
    }
}

# === Style hjälpare ===
function Set-RowBorder {
    param ($ws, [int] $row, [int] $firstRow, [int] $lastRow)
    foreach ($col in 'B','C','D','E','F','G','H') {
        $ws.Cells["$col$row"].Style.Border.Left.Style   = "None"
        $ws.Cells["$col$row"].Style.Border.Right.Style  = "None"
        $ws.Cells["$col$row"].Style.Border.Top.Style    = "None"
        $ws.Cells["$col$row"].Style.Border.Bottom.Style = "None"
    }
    $ws.Cells["B$row"].Style.Border.Left.Style  = "Medium"
    $ws.Cells["H$row"].Style.Border.Right.Style = "Medium"
    foreach ($col in 'B','C','D','E','F','G') { $ws.Cells["$col$row"].Style.Border.Right.Style = "Thin" }
    $topStyle = if ($row -eq $firstRow) { "Medium" } else { "Thin" }
    $bottomStyle = if ($row -eq $lastRow)  { "Medium" } else { "Thin" }
    foreach ($col in 'B','C','D','E','F','G','H') {
        $ws.Cells["$col$row"].Style.Border.Top.Style    = $topStyle
        $ws.Cells["$col$row"].Style.Border.Bottom.Style = $bottomStyle
    }
}
function Style-Cell { param($cell,$bold,$bg,$border,$fontColor)
    if ($bold) { $cell.Style.Font.Bold = $true }
    if ($bg)   { $cell.Style.Fill.PatternType = "Solid"; $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$bg")) }
    if ($fontColor) { $cell.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$fontColor")) }
    if ($border) { $cell.Style.Border.Top.Style=$border; $cell.Style.Border.Bottom.Style=$border; $cell.Style.Border.Left.Style=$border; $cell.Style.Border.Right.Style=$border }
}

# Utility: test if a file is locked (opened in Excel)
function Test-FileLocked { param([Parameter(Mandatory=$true)][string]$Path)
    try { $fs = [IO.File]::Open($Path,'Open','ReadWrite','None'); $fs.Close(); return $false } catch { return $true }
}

# === CSV-hjälpmetoder ===
function Get-CsvDelimiter { param([string]$Path)
    $first = Get-Content -LiteralPath $Path -Encoding Default -TotalCount 30 | Where-Object { $_ -and $_.Trim() } | Select-Object -First 1
    if (-not $first) { return ';' }
    $sc = ($first -split ';').Count; $cc = ($first -split ',').Count
    if ($cc -gt $sc -and $cc -ge 2) { return ',' } else { return ';' }
}
function New-TextFieldParser { param([string]$Path,[string]$Delimiter)
    $tp = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, [System.Text.Encoding]::Default)
    $tp.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $tp.SetDelimiters($Delimiter)
    $tp.HasFieldsEnclosedInQuotes = $true
    $tp.TrimWhiteSpace = $true
    return $tp
}
function Get-AssayFromCsv { param([string]$Path,[int]$StartRow=10)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $tp = $null; $delim=Get-CsvDelimiter $Path; $row=0
    try {
        $tp = New-TextFieldParser -Path $Path -Delimiter $delim
        while (-not $tp.EndOfData) {
            $row++; $f = $tp.ReadFields()
            if ($row -lt $StartRow) { continue }
            if (-not $f -or $f.Length -lt 1) { continue }
            $a=([string]$f[0]).Trim()
            if ($a -and $a -notmatch '^(?i)\s*assay\s*$') { return $a }
        }
    } finally { if ($tp){$tp.Close()} }
    return $null
}
function Import-CsvRows { param([string]$Path,[int]$StartRow=10)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $delim=Get-CsvDelimiter $Path; $tp=$null; $rows=@()
    try {
        $tp = New-TextFieldParser -Path $Path -Delimiter $delim
        $r=0
        while (-not $tp.EndOfData) {
            $r++; $f=$tp.ReadFields()
            if ($r -lt $StartRow) { continue }
            if (-not $f -or ($f -join '').Trim().Length -eq 0) { continue }
            $rows += ,$f
        }
    } finally { if ($tp){$tp.Close()} }
    return ,@($rows)
}

# === Assay-mappning → Control-flik ===
function Normalize-Assay { param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $x=$s.ToLowerInvariant(); $x=[regex]::Replace($x,'[^a-z0-9]+',' '); $x=$x.Trim() -replace '\s+',' '; return $x
}
$AssayMap = @(
    @{ Tab='MTB ULTRA';            Aliases=@('Xpert MTB-RIF Ultra') }
    @{ Tab='MTB RIF';              Aliases=@('Xpert MTB-RIF Assay G4') }
    @{ Tab='MTB JP';               Aliases=@('Xpert MTB-RIF JP IVD') }
    @{ Tab='MTB XDR';              Aliases=@('Xpert MTB-XDR') }
    @{ Tab='FLUVID | FLUVID+';     Aliases=@('Xpress SARS-CoV-2_Flu_RSV plus','Xpert Xpress_SARS-CoV-2_Flu_RSV') }
    @{ Tab='SARS-COV-2 Plus';      Aliases=@('Xpert Xpress CoV-2 plus') }
    @{ Tab='CTNG | CTNG JP';       Aliases=@('Xpert CT_NG','Xpert CT_CE') }
    @{ Tab='C.DIFF | C.DIFF JP';   Aliases=@('Xpert C.difficile G3','Xpert C.difficile BT') }
    @{ Tab='HPV';                  Aliases=@('Xpert HPV HR','Xpert HPV v2 HR') }
    @{ Tab='HBV VL';               Aliases=@('Xpert HBV Viral Load') }
    @{ Tab='HCV VL';               Aliases=@('Xpert HCV Viral Load','Xpert_HCV Viral Load') }
    @{ Tab='HCV VL FS';            Aliases=@('Xpert HCV VL Fingerstick') }
    @{ Tab='HIV VL';               Aliases=@('Xpert HIV-1 Viral Load','Xpert_HIV-1 Viral Load') }
    @{ Tab='HIV VL XC';            Aliases=@('Xpert HIV-1 Viral Load XC') }
    @{ Tab='HIV QA';               Aliases=@('Xpert HIV-1 Qual','Xpert_HIV-1 Qual') }
    @{ Tab='HIV QA XC';            Aliases=@('Xpert HIV-1 Qual XC PQC','Xpert HIV-1 Qual XC') }
    @{ Tab='SARS-COV-2';           Aliases=@('Xpert Xpress SARS-CoV-2 CE-IVD','Xpert Xpress SARS-CoV-2') }
    @{ Tab='FLU RSV';              Aliases=@('Xpert Xpress Flu-RSV','Xpress Flu IPT_EAT off') }
    @{ Tab='MRSA SA';              Aliases=@('Xpert SA Nasal Complete G3','Xpert MRSA-SA SSTI G3') }
    @{ Tab='MRSA NxG';             Aliases=@('Xpert MRSA NxG') }
    @{ Tab='NORO';                 Aliases=@('Xpert Norovirus') }
    @{ Tab='VAN AB';               Aliases=@('Xpert vanA vanB') }
    @{ Tab='GBS';                  Aliases=@('Xpert GBS LB XC','Xpert Xpress GBS','Xpert Xpress GBS US-IVD') }
    @{ Tab='STREP A';              Aliases=@('Xpert Xpress Strep A') }
    @{ Tab='CARBA R';              Aliases=@('Xpert Carba-R','Xpert_Carba-R') }
)
$AssayIndex = @{}
foreach($row in $AssayMap){ foreach($a in $row.Aliases){ $k=Normalize-Assay $a; if($k -and -not $AssayIndex.ContainsKey($k)){ $AssayIndex[$k]=$row.Tab } } }

function Get-ControlTabName {
    param([string]$AssayName)
    $k = Normalize-Assay $AssayName
    if ($k -and $AssayIndex.ContainsKey($k)) { return $AssayIndex[$k] }

    if (Test-Path $SlangAssayPath) {
        try {
            $mapPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($SlangAssayPath))
            $ws = $mapPkg.Workbook.Worksheets['Slang till Assay']; if (-not $ws) { $ws = $mapPkg.Workbook.Worksheets[1] }
            if ($ws -and $ws.Dimension) {
                for ($r=2; $r -le $ws.Dimension.End.Row; $r++){
                    $sheet=$ws.Cells[$r,1].Text.Trim()
                    $aliases=@($ws.Cells[$r,2].Text,$ws.Cells[$r,3].Text,$ws.Cells[$r,4].Text) | Where-Object { $_ -and $_.Trim() }
                    foreach($al in $aliases){ if (Normalize-Assay $AssayName -eq (Normalize-Assay $al)) { $mapPkg.Dispose(); return $sheet } }
                }
            }
            $mapPkg.Dispose()
        } catch {}
    }
    return $null
}

# === Minitab Macro-mappning (Assay → %-kod) ===
$MinitabMap = @(
    @{ Aliases=@('Xpert MTB-RIF Ultra');                           Macro='%D12547-MTBU-SWE' }
    @{ Aliases=@('Xpert MTB-RIF Assay G4');                        Macro='%D12547-MTB-SWE' }
    @{ Aliases=@('Xpress SARS-CoV-2_Flu_RSV plus','Xpert Xpress_SARS-CoV-2_Flu_RSV'); Macro='%D12547-XP3COV2FLURSV-SWE' }
    @{ Aliases=@('Xpert Xpress CoV-2 plus');                        Macro='%D12547-XP3SARSCOV2-SWE' }
    @{ Aliases=@('CT_NG','Xpert CT_CE');                            Macro='%D12547-CTNG-SWE' }
    @{ Aliases=@('Xpert C.difficile G3','Xpert C.difficile BT');    Macro='%D12547-CDCE-SWE' }
    @{ Aliases=@('Xpert HPV HR','Xpert HPV v2 HR');                 Macro='%D12547-HPV-SWE' }
    @{ Aliases=@('Xpert HBV Viral Load');                           Macro='%D12547-HBVVL-SWE' }
    @{ Aliases=@('Xpert HCV Viral Load','Xpert_HCV Viral Load');    Macro='%D12547-HCVVL-SWE' }
    @{ Aliases=@('Xpert HCV VL Fingerstick');                       Macro='%D12547-FSHCV-SWE' }
    @{ Aliases=@('Xpert HIV-1 Viral Load','Xpert_HIV-1 Viral Load'); Macro='%D12547-HIVVL-SWE' }
    @{ Aliases=@('Xpert HIV-1 Qual','Xpert_HIV-1 Qual');            Macro='%D12547-HIVQA-SWE' }
    @{ Aliases=@('Xpert Xpress SARS-CoV-2 CE-IVD','Xpert Xpress SARS-CoV-2'); Macro='%D12547-XPRSARSCOV2-SWE' }
    @{ Aliases=@('Xpert Xpress Flu-RSV');                           Macro='%D12547-XPFLURSV-SWE' }
    @{ Aliases=@('Xpress Flu IPT_EAT off');                         Macro='%D12547-FLUNG-SWE' } 
    @{ Aliases=@('Xpert Norovirus');                                Macro='%D12547-NORO-SWE' }
    @{ Aliases=@('Xpert vanA vanB');                                Macro='%D12547-VAB-SWE' }
    @{ Aliases=@('Xpert Xpress Strep A');                           Macro='%D12547-STREPA-SWE' }
    @{ Aliases=@('Xpert Carba-R','Xpert_Carba-R');                  Macro='%D12547-CARBAR-SWE' }
    @{ Aliases=@('Xpert Ebola EUA','Xpert Ebola CE-IVD');           Macro='%D12547-EBOLA-SWE' }
    @{ Aliases=@('Xpert SA Nasal Complete G3','Xpert MRSA-SA SSTI G3'); Macro='%D12547-SACOMP-SWE' }
    # N/A-gruppen:
    @{ Aliases=@('Xpert GBS LB XC','Xpert Xpress GBS','Xpert Xpress GBS US-IVD'); Macro=$null }
    @{ Aliases=@('Xpert HIV-1 Qual XC PQC','Xpert HIV-1 Qual XC');  Macro=$null }
    @{ Aliases=@('Xpert HIV-1 Viral Load XC');                      Macro=$null }
    @{ Aliases=@('Xpert MTB-RIF JP IVD');                           Macro=$null }
    @{ Aliases=@('Xpert MTB-XDR');                                  Macro=$null }
    @{ Aliases=@('Xpert MRSA NxG');                                 Macro=$null }
)
$MinitabIndex = @{}
foreach ($row in $MinitabMap) { foreach ($a in $row.Aliases) { $k = Normalize-Assay $a; if ($k -and -not $MinitabIndex.ContainsKey($k)) { $MinitabIndex[$k] = $row.Macro } } }
function Get-MinitabMacro { param([string]$AssayName)
    if ([string]::IsNullOrWhiteSpace($AssayName)) { return $null }
    $k = Normalize-Assay $AssayName
    if ($k -and $MinitabIndex.ContainsKey($k)) { return $MinitabIndex[$k] }
    return $null
}

# === Excelbladsdetaljer ===
function Find-ObservationCol { param($ws)
    $default = 13 # M
    if (-not $ws -or -not $ws.Dimension) { return $default }
    $maxR = [Math]::Min(5, $ws.Dimension.End.Row)
    $maxC = $ws.Dimension.End.Column
    for ($r=1; $r -le $maxR; $r++) {
        for ($c=1; $c -le $maxC; $c++) {
            $t = ($ws.Cells[$r,$c].Text + '').Trim()
            if ($t -match '^(?i)\s*(obs|observation)') { return $c }
        }
    }
    return $default
}

# === GUI-utils: CheckedListBox ===
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

# === GUI-hjälp: Clear-GUI ===
function Clear-GUI {
    $txtLSP.Text = ''
    $txtSigner.Text = ''
    $chkWriteSign.Checked = $false
    $chkOverwriteSign.Checked = $false
    Add-CLBItems -clb $clbCsv -files @()
    Add-CLBItems -clb $clbNeg -files @()
    Add-CLBItems -clb $clbPos -files @()
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
    # Uppdatera efter att nya checkstaten har slagit igenom
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
    return $n
}
function Update-StatusBar { $slCount.Text = "$(Get-SelectedFileCount) filer valda" }
function Update-BuildEnabled {
    $btnBuild.Enabled = ((Get-CheckedFilePath $clbNeg) -and (Get-CheckedFilePath $clbPos))
    Update-StatusBar
}

$miScan.add_Click({ $btnScan.PerformClick() })
$miBuild.add_Click({ if ($btnBuild.Enabled) { $btnBuild.PerformClick() } })
$miExit.add_Click({ $form.Close() })

# Nytt – rensa GUI
$miNew.add_Click({ Clear-GUI })

# Öppna senaste rapport
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
    if ([string]::IsNullOrWhiteSpace($p)) { [System.Windows.Forms.MessageBox]::Show("Ange sökvägen till Skript3 i variabeln `$Script3Path.","Skript3") | Out-Null; return }
    if (-not (Test-Path -LiteralPath $p)) { [System.Windows.Forms.MessageBox]::Show("Filen hittades inte:\n$Script3Path","Skript3") | Out-Null; return }
    $ext=[System.IO.Path]::GetExtension($p).ToLowerInvariant()
    switch ($ext) {
        '.ps1' { Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$p`"" }
        '.bat' { Start-Process cmd.exe -ArgumentList "/c `"$p`"" }
        '.lnk' { Start-Process -FilePath $p }
        default { try { Start-Process -FilePath $p } catch { [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna filen:","Skript3") | Out-Null } }
    }
})

# Toggle signatur
$miToggleSign.add_Click({
    $grpSign.Visible = -not $grpSign.Visible
    if ($grpSign.Visible) { $form.Height = $baseHeight + $grpSign.Height + 40; $miToggleSign.Text = 'Dölj Seal Test-signatur' }
    else { $form.Height = $baseHeight; $miToggleSign.Text = 'Aktivera Seal Test-signatur' }
})

# Tema
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
1) Ange LSP och klicka 'Sök filer'.
2) Välj en CSV-fil, en Seal Test Neg- och Pos-fil i listorna (små checkboxar).
3) (Valfritt vid sammanställning) Aktivera "Fyll i 'Print Full Name, Sign, and Date'" via Verktyg.
4) Välj om rapporten ska sparas i LSP-mappen eller bara öppnas temporärt.
5) Klicka 'Skapa rapport'.
"@
    [System.Windows.Forms.MessageBox]::Show($msg,"Instruktioner") | Out-Null
})
$miFAQ.add_Click({
    $faq = @"
Fråga: Vad gör skriptet?
Svar : Det skapar en excel-rapport som jämför sökt LSP för Seal Test-Filer, hämtar utrustningslista och rätt kontrollmaterial för produkten.

Fråga: Ändrar DocMerge mina källfiler?
Svar : Nej. Verktyget skriver endast till NEG/POS om du väljer att skriva in signaturer och sparar rapporten på en ny plats.
"@
    [System.Windows.Forms.MessageBox]::Show($faq,"Vanliga frågor") | Out-Null
})

# Hjälp – enkel dialog
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
$miOm.add_Click({ [System.Windows.Forms.MessageBox]::Show("DocMerge $ScriptVersion`nAv Jesper","Om") | Out-Null })

# === Signaturhjälp ===
function Get-DataSheets { param([OfficeOpenXml.ExcelPackage]$Pkg)
    $all = @($Pkg.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" })
    if ($all.Count -gt 1) { return $all | Select-Object -Skip 1 } else { return @() }
}

function Test-SignatureFormat {
    param([string]$Text)
    $raw = ($Text + '')
    $trim = $raw.Trim()
    $parts = $trim -split '\s*,\s*'
    $name = if ($parts.Count -ge 1) { $parts[0] } else { '' }
    $sign = if ($parts.Count -ge 2) { $parts[1] } else { '' }
    $date = if ($parts.Count -ge 3) { $parts[2] } else { '' }
    $dateOk = $false
    if ($date) { if ($date -match '^\d{4}-\d{2}-\d{2}$' -or $date -match '^\d{8}$') { $dateOk = $true } }
    [pscustomobject]@{ Raw=$raw; Name=$name; Sign=$sign; Date=$date; Parts=$parts.Count; DateOk=$dateOk; LooksOk=($name -ne '' -and $sign -ne '') }
}
function Confirm-SignatureInput { param([string]$Text)
    $info = Test-SignatureFormat $Text
    $hint = @()
    if (-not $info.Name) { $hint += '• Namn verkar saknas' }
    if (-not $info.Sign) { $hint += '• Signatur verkar saknas' }
    if ($info.Date -and -not $info.DateOk) { $hint += "• Datumformat ovanligt: '$($info.Date)'" }
    $msg = "Har du skrivit korrekt 'Print Full Name, Sign, and Date'?

Text: $($info.Raw)

Tolkning:
  • Namn   : $($info.Name)
  • Sign   : $($info.Sign)
  • Datum  : $($info.Date)

" + ($(if ($hint.Count){ "Obs:`n  " + ($hint -join "`n  ") } else { "Ser bra ut." }))
    $res = [System.Windows.Forms.MessageBox]::Show($msg, "Bekräfta signatur", 'YesNo', 'Question')
    return ($res -eq 'Yes')
}
function Get-AnyB47 { param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    if (-not (Load-EPPlus)) { return $null }
    $pkg = $null
    try {
        $pkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($Path))
        foreach ($ws in (Get-DataSheets $pkg)) {
            $txt = ($ws.Cells['B47'].Text + '').Trim()
            if ($txt) { return $txt }
        }
    } catch {} finally { if ($pkg) { try { $pkg.Dispose() } catch {} } }
    return $null
}

function Normalize-Signature {
    param([string]$s)
    if (-not $s) { return '' }
    $x = $s.Trim().ToLowerInvariant()
    # Komprimera whitespace och normalisera kommatecken + blanksteg
    $x = [regex]::Replace($x, '\s+', ' ')
    $x = $x -replace '\s*,\s*', ','
    return $x
}

function Get-SignatureSetForDataSheets {
    param([OfficeOpenXml.ExcelPackage]$Pkg)
    $result = [pscustomobject]@{
        RawFirst  = $null
        NormSet   = New-Object 'System.Collections.Generic.HashSet[string]'
        Occ       = @{}  # normSign -> [List[string]] (bladnamn)
        RawByNorm = @{}  # normSign -> representativ rå text för B47
    }
    if (-not $Pkg) { return $result }

    foreach ($ws in ($Pkg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' })) {
        $h3 = ($ws.Cells['H3'].Text + '').Trim()
        if ($h3 -match '^[0-9]') {
            $raw = ($ws.Cells['B47'].Text + '').Trim()
            if ($raw) {
                $norm = Normalize-Signature $raw
                [void]$result.NormSet.Add($norm)
                if (-not $result.RawFirst) { $result.RawFirst = $raw }
                if (-not $result.Occ.ContainsKey($norm)) {
                    $result.Occ[$norm] = New-Object 'System.Collections.Generic.List[string]'
                }
                if (-not $result.RawByNorm.ContainsKey($norm)) {
                    $result.RawByNorm[$norm] = $raw
                }
                [void]$result.Occ[$norm].Add($ws.Name)
            }
        } elseif ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/?A|NA|Tomt( innehåll)?)$') {
            break
        }
    }
    return $result
}

# === NYTT: SharePoint-batchlänkshjälp ===
function UrlEncode([string]$s){ try { [System.Uri]::EscapeDataString($s) } catch { $s } }

function Get-BatchNumberFromSealFile([string]$Path){
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    if (-not (Load-EPPlus)) { return $null }
    $pkg = $null
    try {
        $pkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($Path))
        foreach ($ws in (Get-DataSheets $pkg)) {
            $txt = ($ws.Cells['D2'].Text + '').Trim()   # "Batch Number"
            if ($txt) { return $txt }
        }
    } catch {} finally { if ($pkg) { try { $pkg.Dispose() } catch {} } }
    return $null
}
function Update-BatchLink {
    $selNeg = Get-CheckedFilePath $clbNeg
    $selPos = Get-CheckedFilePath $clbPos
    $bnNeg  = if ($selNeg) { Get-BatchNumberFromSealFile $selNeg } else { $null }
    $bnPos  = if ($selPos) { Get-BatchNumberFromSealFile $selPos } else { $null }
    $lsp    = $txtLSP.Text.Trim()

    $mismatch = ($bnNeg -and $bnPos -and ($bnNeg -ne $bnPos))
    if ($mismatch) {
        $slBatchLink.Text        = 'SharePoint: mismatch'
        $slBatchLink.Enabled     = $false
        $slBatchLink.Tag         = $null
        $slBatchLink.ToolTipText = "NEG: $bnNeg  |  POS: $bnPos"
        return
    }

    $batch = if ($bnPos) { $bnPos } elseif ($bnNeg) { $bnNeg } else { $null }
    if ($batch) {
        $url = $SharePointBatchLinkTemplate `
            -replace '\{BatchNumber\}', (UrlEncode $batch) `
            -replace '\{LSP\}',         (UrlEncode $lsp)
        $slBatchLink.Text        = "SharePoint: $batch"
        $slBatchLink.Enabled     = $true
        $slBatchLink.Tag         = $url
        $slBatchLink.ToolTipText = $url
    } else {
        $slBatchLink.Text        = 'SharePoint: —'
        $slBatchLink.Enabled     = $false
        $slBatchLink.Tag         = $null
        $slBatchLink.ToolTipText = 'Direktlänk aktiveras när Batch# hittas i POS/NEG.'
    }
}

# === Signatur-läsning (justerad enligt H3-regeln) ===
function Get-SignatureList {
    param([OfficeOpenXml.ExcelPackage]$Pkg)
    $list = @()
    if (-not $Pkg) { return ,@() }
    $sheets = @($Pkg.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" })
    if ($sheets.Count -le 1) { return ,@() }
    for ($i = 1; $i -lt $sheets.Count; $i++) {
        $ws = $sheets[$i]
        $h3 = ($ws.Cells['H3'].Text + '').Trim()
        if ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/?A|NA|Tomt( innehåll)?)$') { break }
        if ($h3 -notmatch '^[0-9]') { continue }  # endast flikar där H3 börjar med siffra
        $sig = ($ws.Cells['B47'].Text + '').Trim()
        if ($sig) { $list += $sig }
    }
    return ,$list
}

# === Sök filer-knapp ===
$btnScan.Add_Click({
    $lsp = $txtLSP.Text.Trim()
    if (-not $lsp) { Gui-Log "⚠️ Ange ett LSP-nummer" 'Warn'; return }

    $folder = $null
    foreach ($path in $RootPaths) {
        $folder = Get-ChildItem $path -Directory -Recurse -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -match "#?$lsp" } |
                  Select-Object -First 1
        if ($folder) { break }
    }
    if (-not $folder) { Gui-Log "❌ Ingen LSP-mapp hittad för $lsp" 'Warn'; return }

    $files = Get-ChildItem $folder.FullName -File -ErrorAction SilentlyContinue
    $candCsv = $files | Where-Object { $_.Extension -ieq '.csv' -and $_.Name -match $lsp } | Sort-Object LastWriteTime -Descending
    $candNeg = $files | Where-Object { $_.Name -match '(?i)Neg.*\.xls[xm]$' -and $_.Name -match $lsp } | Sort-Object LastWriteTime -Descending
    $candPos = $files | Where-Object { $_.Name -match '(?i)Pos.*\.xls[xm]$' -and $_.Name -match $lsp } | Sort-Object LastWriteTime -Descending
    Gui-Log "📂 Hittad mapp: $($folder.FullName)" 'Info'

    Add-CLBItems -clb $clbCsv -files $candCsv -AutoCheckFirst
    Add-CLBItems -clb $clbNeg -files $candNeg -AutoCheckFirst
    Add-CLBItems -clb $clbPos -files $candPos -AutoCheckFirst
    if ($candCsv.Count -eq 0) { Gui-Log "ℹ️ Ingen CSV hittad (endast .csv visas)." 'Info' }
    if ($candNeg.Count -eq 0) { Gui-Log "⚠️ Ingen Seal NEG hittad." 'Warn' }
    if ($candPos.Count -eq 0) { Gui-Log "⚠️ Ingen Seal POS hittad." 'Warn' }
    Update-BuildEnabled
    Update-BatchLink   # NYTT
})

# === Bläddra-knappar ===
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

# ============================
# ===== RAPPORTLOGIK =========
# ============================

$btnBuild.Add_Click({
    if (-not (Load-EPPlus)) { return }

    # Valda filer (CSV valfri)
    $selCsv = Get-CheckedFilePath $clbCsv
    $selNeg = Get-CheckedFilePath $clbNeg
    $selPos = Get-CheckedFilePath $clbPos

    $negWritable = $true; $posWritable = $true
    if ($chkWriteSign.Checked) {
        $negWritable = -not (Test-FileLocked $selNeg); if (-not $negWritable) { Gui-Log "🔒 NEG är låst (öppen i Excel?)." 'Warn' }
        $posWritable = -not (Test-FileLocked $selPos); if (-not $posWritable) { Gui-Log "🔒 POS är låst (öppen i Excel?)." 'Warn' }
    }

    if (-not $selNeg -or -not $selPos) { Gui-Log "❌ Du måste välja en Seal NEG och en Seal POS." 'Error'; return }
    $lsp = $txtLSP.Text.Trim()
    if (-not $lsp) { Gui-Log "⚠️ Ange ett LSP-nummer" 'Warn'; return }

    Gui-Log "📄 Neg-fil: $(Split-Path $selNeg -Leaf)" 'Info'
    Gui-Log "📄 Pos-fil: $(Split-Path $selPos -Leaf)" 'Info'
    if ($selCsv) { Gui-Log "📄 CSV: $(Split-Path $selCsv -Leaf)" 'Info' } else { Gui-Log "ℹ️ Ingen CSV vald." 'Info' }

    # --- Öppna NEG/POS + mall ---
    try {
        $pkgNeg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selNeg))
        $pkgPos = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selPos))
    } catch { Gui-Log "❌ Kunde inte öppna NEG/POS: $($_.Exception.Message)" 'Error'; return }

    $templatePath = Join-Path $PSScriptRoot "Output_Template.xlsx"
    if (-not (Test-Path $templatePath)) { Gui-Log "❌ Mallfilen 'Output_Template.xlsx' saknas!" 'Error'; return }
    try { $pkgOut = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($templatePath)) }
    catch { Gui-Log "❌ Kunde inte läsa mall: $($_.Exception.Message)" 'Error'; return }

    # --- Signera NEG/POS (flik 2→, endast H3 som börjar med siffra) ---
    $signToWrite = ($txtSigner.Text + '').Trim()
    if ($chkWriteSign.Checked) {
        if (-not $signToWrite) { Gui-Log "❌ Ingen signatur angiven (B47). Avbryter."; return }
        if (-not (Confirm-SignatureInput -Text $signToWrite)) { Gui-Log "🛑 Signatur ej bekräftad. Avbryter."; return }

        $negWritten = 0; $posWritten = 0; $negSkipped = 0; $posSkipped = 0

        foreach ($ws in $pkgNeg.Workbook.Worksheets) {
            if ($ws.Name -eq 'Worksheet Instructions') { continue }
            $h3 = ($ws.Cells['H3'].Text + '').Trim()
            if ($h3 -match '^[0-9]') {
                $existing = ($ws.Cells['B47'].Text + '').Trim()
                if ($existing -and -not $chkOverwriteSign.Checked) { $negSkipped++; continue }
                $ws.Cells['B47'].Style.Numberformat.Format = '@'
                $ws.Cells['B47'].Value = $signToWrite
                $negWritten++
            } elseif ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/\?A|NA|Tomt( innehåll)?)$') {
                break
            }
        }
        foreach ($ws in $pkgPos.Workbook.Worksheets) {
            if ($ws.Name -eq 'Worksheet Instructions') { continue }
            $h3 = ($ws.Cells['H3'].Text + '').Trim()
            if ($h3 -match '^[0-9]') {
                $existing = ($ws.Cells['B47'].Text + '').Trim()
                if ($existing -and -not $chkOverwriteSign.Checked) { $posSkipped++; continue }
                $ws.Cells['B47'].Style.Numberformat.Format = '@'
                $ws.Cells['B47'].Value = $signToWrite
                $posWritten++
            } elseif ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/\?A|NA|Tomt( innehåll)?)$') {
                break
            }
        }
        try {
            if ($negWritten -eq 0 -and $negSkipped -eq 0 -and $posWritten -eq 0 -and $posSkipped -eq 0) {
                Gui-Log "ℹ️ Inga databladsflikar efter flik 1 att sätta signatur i (ingen åtgärd)."
            } else {
                if ($negWritten -gt 0 -and $negWritable) { $pkgNeg.Save() } elseif ($negWritten -gt 0) { Gui-Log "🔒 Kunde inte spara NEG (låst)." 'Warn' }
                if ($posWritten -gt 0 -and $posWritable) { $pkgPos.Save() } elseif ($posWritten -gt 0) { Gui-Log "🔒 Kunde inte spara POS (låst)." 'Warn' }
                Gui-Log "🖊️ Signatur satt: NEG $negWritten blad (överhoppade $negSkipped), POS $posWritten blad (överhoppade $posSkipped)."
            }
        } catch { Gui-Log "⚠️ Kunde inte spara signatur i NEG/POS: $($_.Exception.Message)" }
    }

    # --- Läs CSV för Info/Control ---
    $csvRows = @(); $runAssay = $null
    if ($selCsv) {
        try { $csvRows = Import-CsvRows -Path $selCsv -StartRow 10 } catch {}
        try { $runAssay = Get-AssayFromCsv -Path $selCsv -StartRow 10 } catch {}
        if ($runAssay) { Gui-Log "🔎 Assay från CSV: $runAssay" }
    }
    $controlTab = $null
    if ($runAssay) { $controlTab = Get-ControlTabName -AssayName $runAssay }
    if ($controlTab) { Gui-Log "🧪 Control Material-flik: $controlTab" } else { Gui-Log "ℹ️ Ingen control-mappning (fortsätter utan)." }

    # --- Läs avvikelser NEG/POS ---
    $violationsNeg = @(); $violationsPos = @(); $failNegCount = 0; $failPosCount = 0
    foreach ($ws in $pkgNeg.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) {
        if (-not $ws.Dimension) { continue }
        $obsC = Find-ObservationCol $ws
        for ($r = 3; $r -le 45; $r++) {
            $valK = $ws.Cells["K$r"].Value; $textL = $ws.Cells["L$r"].Text
            if ($valK -ne $null -and $valK -is [double]) {
                if ($textL -eq "FAIL" -or $valK -le -2.4) {
                    $obsTxt = $ws.Cells[$r, $obsC].Text
                    $violationsNeg += [PSCustomObject]@{
                        Sheet      = $ws.Name
                        Cartridge  = $ws.Cells["H$r"].Text
                        InitialW   = $ws.Cells["I$r"].Value
                        FinalW     = $ws.Cells["J$r"].Value
                        WeightLoss = $valK
                        Status     = if ($textL -eq "FAIL") { "FAIL" } else { "Minusvärde" }
                        Obs        = $obsTxt
                    }
                    if ($textL -eq "FAIL") { $failNegCount++ }
                }
            }
        }
    }
    foreach ($ws in $pkgPos.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) {
        if (-not $ws.Dimension) { continue }
        $obsC = Find-ObservationCol $ws
        for ($r = 3; $r -le 45; $r++) {
            $valK = $ws.Cells["K$r"].Value; $textL = $ws.Cells["L$r"].Text
            if ($valK -ne $null -and $valK -is [double]) {
                if ($textL -eq "FAIL" -or $valK -le -2.4) {
                    $obsTxt = $ws.Cells[$r, $obsC].Text
                    $violationsPos += [PSCustomObject]@{
                        Sheet      = $ws.Name
                        Cartridge  = $ws.Cells["H$r"].Text
                        InitialW   = $ws.Cells["I$r"].Value
                        FinalW     = $ws.Cells["J$r"].Value
                        WeightLoss = $valK
                        Status     = if ($textL -eq "FAIL") { "FAIL" } else { "Minusvärde" }
                        Obs        = $obsTxt
                    }
                    if ($textL -eq "FAIL") { $failPosCount++ }
                }
            }
        }
    }

    # --- Fyll "Seal Test Output" ---
    $wsOut1 = $pkgOut.Workbook.Worksheets["Seal Test Output"]
    if (-not $wsOut1) { Gui-Log "❌ Fliken 'Seal Test Output' saknas i mallen"; return }

    # Nollställ mismatch-kolumn (D3..D15)
    for ($row = 3; $row -le 15; $row++) {
        $wsOut1.Cells["D$row"].Value = $null
        try { $wsOut1.Cells["D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::None } catch {}
    }

    $fields = @(
        @{ Label = "ROBAL";                         Cell = "F2"  }
        @{ Label = "Part Number";                   Cell = "B2"  }
        @{ Label = "Batch Number";                  Cell = "D2"  }
        @{ Label = "Cartridge Number (LSP)";        Cell = "B6"  }
        @{ Label = "PO Number";                     Cell = "B10" }
        @{ Label = "Assay Family";                  Cell = "D10" }
        @{ Label = "Weight Loss Spec";              Cell = "F10" }
        @{ Label = "Balance ID Number";             Cell = "B14" }
        @{ Label = "Balance Cal Due Date";          Cell = "D14" }
        @{ Label = "Vacuum Oven ID Number";         Cell = "B20" }
        @{ Label = "Vacuum Oven Cal Due Date";      Cell = "D20" }
        @{ Label = "Timer ID Number";               Cell = "B25" }
        @{ Label = "Timer Cal Due Date";            Cell = "D25" }
    )
    $forceText = @(
        "ROBAL","Part Number","Batch Number","Cartridge Number (LSP)",
        "PO Number","Assay Family","Balance ID Number","Vacuum Oven ID Number","Timer ID Number"
    )
    $mismatchFields = $fields[0..6] | ForEach-Object { $_.Label }

    $row = 3
    foreach ($f in $fields) {
        $valNeg=''; $valPos=''
        foreach ($wsN in $pkgNeg.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) {
            $cell = $wsN.Cells[$f.Cell]
            if ($cell.Value -ne $null) { if ($cell.Value -is [datetime]) { $valNeg = $cell.Value.ToString('MMM-yy') } else { $valNeg = $cell.Text }; break }
        }
        foreach ($wsP in $pkgPos.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) {
            $cell = $wsP.Cells[$f.Cell]
            if ($cell.Value -ne $null) { if ($cell.Value -is [datetime]) { $valPos = $cell.Value.ToString('MMM-yy') } else { $valPos = $cell.Text }; break }
        }

        if ($forceText -contains $f.Label) {
            $wsOut1.Cells["B$row"].Style.Numberformat.Format = '@'
            $wsOut1.Cells["C$row"].Style.Numberformat.Format = '@'
        }

        $wsOut1.Cells["B$row"].Value = $valNeg
        $wsOut1.Cells["C$row"].Value = $valPos
        $wsOut1.Cells["B$row"].Style.Border.Right.Style = "Medium"
        $wsOut1.Cells["C$row"].Style.Border.Left.Style  = "Medium"

        if ($mismatchFields -contains $f.Label -and $valNeg -ne $valPos) {
            $wsOut1.Cells["D$row"].Value = "Mismatch"
            Style-Cell $wsOut1.Cells["D$row"] $true "FF0000" "Medium" "FFFFFF"
            Gui-Log "⚠️ Mismatch: $($f.Label) ($valNeg vs $valPos)"
        }
        $row++
    }

    # --- Testare (B43) ---
    $testersNeg = @(); $testersPos = @()
    foreach ($s in $pkgNeg.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) { $t=$s.Cells["B43"].Text; if ($t) { $testersNeg += ($t -split ",") } }
    foreach ($s in $pkgPos.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) { $t=$s.Cells["B43"].Text; if ($t) { $testersPos += ($t -split ",") } }
    $testersNeg = $testersNeg | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique
    $testersPos = $testersPos | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique

    $wsOut1.Cells["B16"].Value = "Name of Tester"
    $wsOut1.Cells["B16:C16"].Merge = $true
    $wsOut1.Cells["B16"].Style.HorizontalAlignment = "Center"

    $maxTesters = [Math]::Max($testersNeg.Count, $testersPos.Count)
    $initialRows = 5
    if ($maxTesters -lt $initialRows) { $wsOut1.DeleteRow(17 + $maxTesters, $initialRows - $maxTesters) }
    if ($maxTesters -gt $initialRows) {
        $rowsToAdd = $maxTesters - $initialRows
        $lastRow = 16 + $initialRows
        for ($i = 1; $i -le $rowsToAdd; $i++) { $wsOut1.InsertRow($lastRow + 1, 1, $lastRow) }
    }
    for ($i = 0; $i -lt $maxTesters; $i++) {
        $rowIndex = 17 + $i
        $wsOut1.Cells["A$rowIndex"].Value = $null
        $wsOut1.Cells["B$rowIndex"].Value = if ($i -lt $testersNeg.Count) { $testersNeg[$i] } else { "N/A" }
        $wsOut1.Cells["C$rowIndex"].Value = if ($i -lt $testersPos.Count) { $testersPos[$i] } else { "N/A" }

        $topStyle    = if ($i -eq 0) { "Medium" } else { "Thin" }
        $bottomStyle = if ($i -eq $maxTesters - 1) { "Medium" } else { "Thin" }
        foreach ($col in @("B","C")) {
            $cell = $wsOut1.Cells["$col$rowIndex"]
            $cell.Style.Border.Top.Style    = $topStyle
            $cell.Style.Border.Bottom.Style = $bottomStyle
            $cell.Style.Border.Left.Style   = "Medium"
            $cell.Style.Border.Right.Style  = "Medium"
            $cell.Style.Fill.PatternType = "Solid"
            $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
        }
    }

    # --- Signatur-mismatchkontroll (B47) med H3-regel ---
$negSigSet = Get-SignatureSetForDataSheets -Pkg $pkgNeg
$posSigSet = Get-SignatureSetForDataSheets -Pkg $pkgPos

# Bygg mängder
$negSet = New-Object 'System.Collections.Generic.HashSet[string]'
$posSet = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($n in $negSigSet.NormSet) { [void]$negSet.Add($n) }
foreach ($p in $posSigSet.NormSet) { [void]$posSet.Add($p) }

$hasNeg = ($negSet.Count -gt 0)
$hasPos = ($posSet.Count -gt 0)

# Ny policy (robust):
# - Om båda sidor har minst EN signatur: mismatch om MÄNGDERNA skiljer sig (dvs. det finns något namn som bara finns på ena sidan)
# - Om en sida saknar signatur helt: räkna INTE som mismatch
# - Skillnader i antal flikar i sig triggar inte mismatch; det är faktisk textskillnad som triggar.
$onlyNeg = @()
$onlyPos = @()
$sigMismatch = $false
if ($hasNeg -and $hasPos) {
    foreach ($n in $negSet) { if (-not $posSet.Contains($n)) { $onlyNeg += $n } }
    foreach ($p in $posSet) { if (-not $negSet.Contains($p)) { $onlyPos += $p } }
    $sigMismatch = ($onlyNeg.Count -gt 0 -or $onlyPos.Count -gt 0)
} else {
    $sigMismatch = $false
}

# Bygg en tydlig lista över vilka signaturer som skiljer sig och på vilka blad de förekommer
$mismatchSheets = @()
if ($sigMismatch) {
    foreach ($k in $onlyNeg) {
        $raw = if ($negSigSet.RawByNorm.ContainsKey($k)) { $negSigSet.RawByNorm[$k] } else { $k }
        $where = if ($negSigSet.Occ.ContainsKey($k)) { ($negSigSet.Occ[$k] -join ', ') } else { '—' }
        $mismatchSheets += ("NEG only: " + $raw + "  [Blad: " + $where + "]")
    }
    foreach ($k in $onlyPos) {
        $raw = if ($posSigSet.RawByNorm.ContainsKey($k)) { $posSigSet.RawByNorm[$k] } else { $k }
        $where = if ($posSigSet.Occ.ContainsKey($k)) { ($posSigSet.Occ[$k] -join ', ') } else { '—' }
        $mismatchSheets += ("POS only: " + $raw + "  [Blad: " + $where + "]")
    }
    Gui-Log "⚠️ Mismatch: Print Full Name, Sign, and Date (NEG vs POS)"
}

# --- Infoga signaturinformation i rad 32 ---
    $signRow = 32
    $displaySignNeg = $null
    $displaySignPos = $null
    if ($signToWrite) { $displaySignNeg = $signToWrite; $displaySignPos = $signToWrite }
    else {
$displaySignNeg = if ($signToWrite) { $signToWrite } elseif ($negSigSet.RawFirst) { $negSigSet.RawFirst } else { '—' }
$displaySignPos = if ($signToWrite) { $signToWrite } elseif ($posSigSet.RawFirst) { $posSigSet.RawFirst } else { '—' }
    }
$wsOut1.Cells["B$signRow"].Value = $null
$wsOut1.Cells["C$signRow"].Value = $null
$wsOut1.Cells["D$signRow"].Value = $null

$wsOut1.Cells["B$signRow"].Style.Numberformat.Format = '@'
$wsOut1.Cells["C$signRow"].Style.Numberformat.Format = '@'
$wsOut1.Cells["B$signRow"].Value = $displaySignNeg
$wsOut1.Cells["C$signRow"].Value = $displaySignPos
foreach ($col in @('B','C')) {
    $cell = $wsOut1.Cells["${col}$signRow"]
    Style-Cell $cell $false 'CCFFFF' 'Medium' $null
    $cell.Style.HorizontalAlignment = 'Center'
}
    try { $wsOut1.Column(2).Width = 40; $wsOut1.Column(3).Width = 40 } catch {}
    try { $wsOut1.Column(4).Width = 10 } catch {}  # ≈ 68 px

# Visa mismatch‐indikation + bladlista endast om sigMismatch = true
if ($sigMismatch) {
    $mismatchCell = $wsOut1.Cells["D$signRow"]
    $mismatchCell.Value = 'Mismatch'
    Style-Cell $mismatchCell $true 'FF0000' 'Medium' 'FFFFFF'

    if ($mismatchSheets.Count -gt 0) {
        for ($j = 0; $j -lt $mismatchSheets.Count; $j++) {
            $rowIdx = $signRow + 1 + $j
            try { $wsOut1.Cells["B$rowIdx:C$rowIdx"].Merge = $true } catch {}
            $wsOut1.Cells["B$rowIdx"].Value = $mismatchSheets[$j]
            foreach ($mc in $wsOut1.Cells["B$rowIdx:C$rowIdx"]) {
                Style-Cell $mc $false 'CCFFFF' 'Medium' $null
            }
            $wsOut1.Cells["B$rowIdx:C$rowIdx"].Style.HorizontalAlignment = 'Center'
        }
    }
}

    # --- STF + Minusvärde ---
    $wsOut2 = $pkgOut.Workbook.Worksheets["STF + Minusvärde"]
    if (-not $wsOut2) { Gui-Log "❌ Fliken 'STF + Minusvärde' saknas i mallen!"; return }

    $totalRows = $violationsNeg.Count + $violationsPos.Count
    $currentRow = 2
    if ($totalRows -eq 0) {
        Gui-Log "✅ Inga avvikelser hittades"
        $wsOut2.Cells["B1:H1"].Value = $null
        $wsOut2.Cells["A1"].Value = "Inga avvikelser hittades!"
        Style-Cell $wsOut2.Cells["A1"] $true "D9EAD3" "Medium" "006100"
        $wsOut2.Cells["A1"].Style.HorizontalAlignment = "Left"
        if ($wsOut2.Dimension -and $wsOut2.Dimension.End.Row -gt 1) { $wsOut2.DeleteRow(2, $wsOut2.Dimension.End.Row - 1) }
    } else {
        Gui-Log "❗ $failNegCount FAIL i NEG, $failPosCount i POS"

        $oldDataRows = 0
        if ($wsOut2.Dimension) { $oldDataRows = $wsOut2.Dimension.End.Row - 1; if ($oldDataRows -lt 0) { $oldDataRows = 0 } }
        if ($totalRows -lt $oldDataRows) { $wsOut2.DeleteRow(2 + $totalRows, $oldDataRows - $totalRows) }
        elseif ($totalRows -gt $oldDataRows) { $wsOut2.InsertRow(2 + $oldDataRows, $totalRows - $oldDataRows, 1 + $oldDataRows) }

        $currentRow = 2

        foreach ($v in $violationsNeg) {
            $wsOut2.Cells["A$currentRow"].Value = "NEG"
            $wsOut2.Cells["B$currentRow"].Value = $v.Sheet
            $wsOut2.Cells["C$currentRow"].Value = $v.Cartridge
            $wsOut2.Cells["D$currentRow"].Value = $v.InitialW
            $wsOut2.Cells["E$currentRow"].Value = $v.FinalW
            $wsOut2.Cells["F$currentRow"].Value = [Math]::Round($v.WeightLoss, 1)
            $wsOut2.Cells["G$currentRow"].Value = $v.Status
            $wsOut2.Cells["H$currentRow"].Value = if ([string]::IsNullOrWhiteSpace($v.Obs)) { 'NA' } else { $v.Obs }

            Style-Cell $wsOut2.Cells["A$currentRow"] $true "B5E6A2" "Medium" $null
            $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.PatternType = "Solid"
            $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
            $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.PatternType = "Solid"
            $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFFF99"))
            $wsOut2.Cells["H$currentRow"].Style.Fill.PatternType = "Solid"
            $wsOut2.Cells["H$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#D9D9D9"))

            if ($v.Status -in @("FAIL","Minusvärde")) {
                $wsOut2.Cells["F$currentRow"].Style.Font.Bold = $true
                $wsOut2.Cells["F$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                $wsOut2.Cells["G$currentRow"].Style.Font.Bold = $true
                $wsOut2.Cells["G$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
            }
            Set-RowBorder -ws $wsOut2 -row $currentRow -firstRow 2 -lastRow ($totalRows + 1)
            $currentRow++
        }

        foreach ($v in $violationsPos) {
            $wsOut2.Cells["A$currentRow"].Value = "POS"
            $wsOut2.Cells["B$currentRow"].Value = $v.Sheet
            $wsOut2.Cells["C$currentRow"].Value = $v.Cartridge
            $wsOut2.Cells["D$currentRow"].Value = $v.InitialW
            $wsOut2.Cells["E$currentRow"].Value = $v.FinalW
            $wsOut2.Cells["F$currentRow"].Value = [Math]::Round($v.WeightLoss, 1)
            $wsOut2.Cells["G$currentRow"].Value = $v.Status
            $wsOut2.Cells["H$currentRow"].Value = if ($v.Obs) { $v.Obs } else { 'NA' }

            Style-Cell $wsOut2.Cells["A$currentRow"] $true "FFB3B3" "Medium" $null
            $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.PatternType = "Solid"
            $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
            $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.PatternType = "Solid"
            $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFFF99"))
            $wsOut2.Cells["H$currentRow"].Style.Fill.PatternType = "Solid"
            $wsOut2.Cells["H$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#D9D9D9"))

            if ($v.Status -in @("FAIL","Minusvärde")) {
                $wsOut2.Cells["F$currentRow"].Style.Font.Bold = $true
                $wsOut2.Cells["F$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                $wsOut2.Cells["G$currentRow"].Style.Font.Bold = $true
                $wsOut2.Cells["G$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
            }
            Set-RowBorder -ws $wsOut2 -row $currentRow -firstRow 2 -lastRow ($totalRows + 1)
            $currentRow++
        }

        $wsOut2.Cells.Style.WrapText = $false
        $wsOut2.Cells["A1"].Style.HorizontalAlignment = "Left"
        try { $wsOut2.Cells[2,6,([Math]::Max($currentRow-1,2)),6].Style.Numberformat.Format = '0.0' } catch {}
        if ($wsOut2.Dimension) { $wsOut2.Cells[$wsOut2.Dimension.Address].AutoFitColumns() }
    }

    # --- Header watermark "UNCONTROLLED" ---
    try {
        foreach ($ws in $pkgOut.Workbook.Worksheets) {
            try {
                $ws.HeaderFooter.OddHeader.CenteredText   = '&"Arial,Bold"&14 UNCONTROLLED'
                $ws.HeaderFooter.EvenHeader.CenteredText  = '&"Arial,Bold"&14 UNCONTROLLED'
                $ws.HeaderFooter.FirstHeader.CenteredText = '&"Arial,Bold"&14 UNCONTROLLED'
            } catch { Write-Warning "Kunde inte sätta header på blad: $($ws.Name)" }
        }
    } catch { Write-Warning "Fel vid vattenstämpling av rapporten." }

    # --- Textformat på strängfält i "Seal Test Output" ---
    try {
        for ($r = 3; $r -le 15; $r++) {
            $label = $wsOut1.Cells["A$r"].Text
            if ($label -match 'ROBAL|PO Number|Batch Number|Number') {
                $wsOut1.Cells["B$r:C$r"].Style.Numberformat.Format = '@'
            }
        }
    } catch {}

    # --- Flik "Information" (utan 'Antal avvikelser') ---
    try {
        $wsInfo = $pkgOut.Workbook.Worksheets["Information"]
        if ($wsInfo) { $pkgOut.Workbook.Worksheets.Delete($wsInfo) }
        $wsInfo = $pkgOut.Workbook.Worksheets.Add("Information")
        try { $wsInfo.Cells.Style.Font.Name='Arial'; $wsInfo.Cells.Style.Font.Size=11 } catch {}

        $csvLeaf = if ($selCsv) { Split-Path $selCsv -Leaf } else { '' }
        $negLeaf = Split-Path $selNeg -Leaf
        $posLeaf = Split-Path $selPos -Leaf

        # Minitab macro, primärt från CSV, annars från D10 i utbladet
        $assayForMacro = if ($runAssay) { $runAssay } else { ($wsOut1.Cells['D10'].Text + '').Trim() }
        $miniVal = Get-MinitabMacro -AssayName $assayForMacro
        if (-not $miniVal) { $miniVal = 'N/A' }

        $meta = [ordered]@{
            'Datum Tid'     = (Get-Date).ToString('yyyy-MM-dd HH:mm')
            'LSP'           = $lsp
            'Skriptversion' = $ScriptVersion
            'Användare'     = $env:USERNAME
            'Vald CSV'      = $csvLeaf
            'Vald Seal NEG' = $negLeaf
            'Vald Seal POS' = $posLeaf
            'Minitab Macro' = $miniVal
        }

        $r=1
        foreach($k in $meta.Keys){
            $wsInfo.Cells[$r,1].Value = $k
            if ($k -eq 'Minitab Macro') { $wsInfo.Cells[$r,2].Style.Numberformat.Format = '@' }
            $wsInfo.Cells[$r,2].Value = $meta[$k]
            $wsInfo.Cells[$r,1].Style.Font.Bold = $true
            $wsInfo.Cells[$r,1].Style.Fill.PatternType = 'Solid'
            $wsInfo.Cells[$r,1].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#D9D9D9"))
            $r++
        }
        $thin=[OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $wsInfo.Cells[1,1,$meta.Keys.Count,2].Style.Border.Left.Style   = $thin
        $wsInfo.Cells[1,1,$meta.Keys.Count,2].Style.Border.Right.Style  = $thin
        $wsInfo.Cells[1,1,$meta.Keys.Count,2].Style.Border.Top.Style    = $thin
        $wsInfo.Cells[1,1,$meta.Keys.Count,2].Style.Border.Bottom.Style = $thin
        try { $wsInfo.Cells[1,1,$meta.Keys.Count,2].AutoFitColumns() | Out-Null } catch {}
    } catch { Gui-Log "⚠️ Kunde inte skapa fliken 'Information': $($_.Exception.Message)" }

    # --- Flik "Equipment" (kolumnbredder bevaras) ---
    try {
        if (Test-Path -LiteralPath $UtrustningListPath) {
            $srcPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($UtrustningListPath))
            $srcWs  = $srcPkg.Workbook.Worksheets[1]
            if ($srcWs) {
                $wsEq = $pkgOut.Workbook.Worksheets["Equipment"]
                if ($wsEq) { $pkgOut.Workbook.Worksheets.Delete($wsEq) }
                $wsEq = $pkgOut.Workbook.Worksheets.Add('Equipment', $srcWs)
                if ($wsEq.Dimension) {
                    foreach ($cell in $wsEq.Cells[$wsEq.Dimension.Address]) {
                        if ($cell.Formula -or $cell.FormulaR1C1) { $val=$cell.Value; $cell.Formula=$null; $cell.FormulaR1C1=$null; $cell.Value=$val }
                    }
                    $colCount = $srcWs.Dimension.End.Column
                    for ($c=1; $c -le $colCount; $c++) { try { $wsEq.Column($c).Width = $srcWs.Column($c).Width } catch {} }
                }
            }
            $srcPkg.Dispose()
        } else { Gui-Log "ℹ️ Utrustningslista saknas: $UtrustningListPath" }
    } catch { Gui-Log "⚠️ Kunde inte kopiera 'Equipment': $($_.Exception.Message)" }

    # --- Control Material-flik ---
    try {
        if ($controlTab -and (Test-Path -LiteralPath $RawDataPath)) {
            $srcPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($RawDataPath))
            try { $srcPkg.Workbook.Calculate() } catch {}
            $candidates = if ($controlTab -match '\|') { $controlTab -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @($controlTab) }
            $srcWs = $null
            foreach ($cand in $candidates) {
                $srcWs = $srcPkg.Workbook.Worksheets | Where-Object { $_.Name -eq $cand } | Select-Object -First 1
                if ($srcWs) { break }
                $srcWs = $srcPkg.Workbook.Worksheets | Where-Object { $_.Name -like "*$cand*" } | Select-Object -First 1
                if ($srcWs) { break }
            }
            if ($srcWs) {
                $safeName = if ($srcWs.Name.Length -gt 31) { $srcWs.Name.Substring(0,31) } else { $srcWs.Name }
                $destName = $safeName; $n=1
                while ($pkgOut.Workbook.Worksheets[$destName]) { $base = if ($safeName.Length -gt 27) { $safeName.Substring(0,27) } else { $safeName }; $destName = "$base($n)"; $n++ }
                $wsCM = $pkgOut.Workbook.Worksheets.Add($destName, $srcWs)
                if ($wsCM.Dimension) {
                    foreach ($cell in $wsCM.Cells[$wsCM.Dimension.Address]) {
                        if ($cell.Formula -or $cell.FormulaR1C1) { $v=$cell.Value; $cell.Formula=$null; $cell.FormulaR1C1=$null; $cell.Value=$v }
                    }
                    try { $wsCM.Cells[$wsCM.Dimension.Address].AutoFitColumns() | Out-Null } catch {}
                }
                Gui-Log "✅ Control Material kopierad: '$($srcWs.Name)' → '$destName'"
            } else { Gui-Log "ℹ️ Hittade inget blad i kontrollfilen som matchar '$controlTab'." }
            $srcPkg.Dispose()
        } else { Gui-Log "ℹ️ Ingen Control-flik skapad (saknar mappning eller kontrollfil)." }
    } catch { Gui-Log "⚠️ Control Material-fel: $($_.Exception.Message)" }

    # Flytta informationsfliken sist
    try { $wsInfoMove = $pkgOut.Workbook.Worksheets["Information"]; if ($wsInfoMove) { $wsInfoMove.MoveToEnd() } } catch {}

    # --- Spara ---
    $nowTs   = Get-Date -Format "yyyyMMdd_HHmmss"
    $baseName= "DocMerge_output_${lsp}_$nowTs.xlsx"
    if ($rbSaveInLsp.Checked) {
        $saveDir = Split-Path -Parent $selNeg
        $SavePath = Join-Path $saveDir $baseName
        Gui-Log "💾 Sparläge: LSP-mapp → $saveDir"
    } else {
        $saveDir = $env:TEMP
        $SavePath = Join-Path $saveDir $baseName
        Gui-Log "💾 Sparläge: Temporärt → $SavePath"
    }

    try {
        $pkgOut.Workbook.View.ActiveTab = 0
        $pkgOut.Workbook.Worksheets["Seal Test Output"].View.TabSelected = $true
        $pkgOut.SaveAs($SavePath)
        Gui-Log "✅ Rapport sparad: $SavePath" 'Info'
        $global:LastReportPath = $SavePath

        # --- Revisionsfil (audit) ---
        try {
            $auditDir = Join-Path $PSScriptRoot 'audit'
            if (-not (Test-Path $auditDir)) { New-Item -ItemType Directory -Path $auditDir -Force | Out-Null }
            $auditObj = [pscustomobject]@{
                DatumTid        = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                Användare        = $env:USERNAME
                LSP              = $lsp
                ValdCSV          = if ($selCsv) { Split-Path $selCsv -Leaf } else { '' }
                ValdSealNEG      = Split-Path $selNeg -Leaf
                ValdSealPOS      = Split-Path $selPos -Leaf
                SignaturSkriven  = if ($chkWriteSign.Checked) { 'Ja' } else { 'Nej' }
                OverwroteSign    = if ($chkOverwriteSign.Checked) { 'Ja' } else { 'Nej' }
                SigMismatch      = if ($sigMismatch) { 'Ja' } else { 'Nej' }
                MismatchSheets   = if ($mismatchSheets -and $mismatchSheets.Count -gt 0) { ($mismatchSheets -join ';') } else { '' }
                Violations       = ($violationsNeg.Count + $violationsPos.Count)
                OutputFile       = $SavePath
                Kommentar        = 'UNCONTROLLED rapport, ingen källfil ändrades automatiskt.'
            }
            $auditFile = Join-Path $auditDir ("DocMerge_audit_${nowTs}.csv")
            $auditObj | Export-Csv -Path $auditFile -NoTypeInformation -Encoding UTF8
        } catch { Gui-Log "⚠️ Kunde inte skriva revisionsfil: $($_.Exception.Message)" }

        # Öppna endast rapporten i Excel (CSV öppnas INTE automatiskt)
        Start-Process -FilePath "excel.exe" -ArgumentList "`"$SavePath`""
    }
    catch { Gui-Log "⚠️ Kunde inte spara/öppna: $($_.Exception.Message)" }
    finally {
        try { $pkgNeg.Dispose() } catch {}
        try { $pkgPos.Dispose() } catch {}
        try { $pkgOut.Dispose() } catch {}
    }

    Gui-Log "🔐 Klar."
})

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

# Uppdatera länk när LSP skrivs in
$txtLSP.add_TextChanged({ Update-BatchLink })

# =============== SLUT ===============
function Enable-DoubleBuffer {
    $pi = [Windows.Forms.Control].GetProperty('DoubleBuffered',[Reflection.BindingFlags]'NonPublic,Instance')
    foreach($c in @($content,$pLog,$grpPick,$grpSign,$grpSave)) { if ($c) { $pi.SetValue($c,$true,$null) } }
}
try { Set-Theme 'light' } catch {}
Enable-DoubleBuffer
Update-BatchLink  # Initiera statuslänk
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)

try{ Stop-Transcript | Out-Null }catch{}