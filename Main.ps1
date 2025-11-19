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

$global:Splash = $null
function Show-Splash([string]$msg="Startar…") {
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing
    $f = New-Object Windows.Forms.Form
    $f.FormBorderStyle = 'None'
    $f.StartPosition   = 'CenterScreen'
    $f.BackColor = [Drawing.Color]::FromArgb(0,120,215)
    $f.ForeColor       = [Drawing.Color]::White
    $f.Size            = New-Object Drawing.Size(420,120)
    $lbl = New-Object Windows.Forms.Label
    $lbl.Dock='Fill'; $lbl.TextAlign='MiddleCenter'
    $lbl.Font = New-Object Drawing.Font('Segoe UI Semibold',12)
    $f.Controls.Add($lbl)
    $f.TopMost = $true
    $f.Show()
    $global:Splash = @{ Form=$f; Label=$lbl }
    Update-Splash $msg
    [Windows.Forms.Application]::DoEvents()
}
function Update-Splash([string]$msg) {
    if ($global:Splash) { $global:Splash.Label.Text = $msg; [Windows.Forms.Application]::DoEvents() }
}
function Close-Splash() {
    if ($global:Splash) { $global:Splash.Form.Close(); $global:Splash.Form.Dispose(); $global:Splash = $null }
}

$Host.UI.RawUI.WindowTitle = "Startar…"
Show-Splash "Laddar PnP.PowerShell…"

# === SharePoint init (före WinForms) ===
$global:SpConnected = $false
$global:SpError     = $null

# 0) NuGet + PnP-modul
try {
    $null = Get-PackageProvider -Name "NuGet" -ForceBootstrap -ErrorAction SilentlyContinue
} catch {}
try {
    Update-Splash "Laddar PnP.PowerShell…"
Import-Module PnP.PowerShell -ErrorAction Stop
} catch {
    try {
        Write-Host "PnP ej hittad – installerar (kan ta någon minut)…"
        Install-Module PnP.PowerShell -MaximumVersion 1.12.0 -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Update-Splash "Laddar PnP.PowerShell…"
Import-Module PnP.PowerShell -ErrorAction Stop
    } catch {
        $global:SpError = "PnP-install/import misslyckades: $($_.Exception.Message)"
    }
}

$env:PNPPOWERSHELL_UPDATECHECK = "Off"
$global:SP_ClientId   = "INSERT MYSELF"
$global:SP_Tenant     = "danaher.onmicrosoft.com"
$global:SP_CertBase64 = "INSERT MYSELF"
$global:SP_SiteUrl    = "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management"

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
        # Om inloggningen misslyckas, logga och visa felet för användaren på splash-skärmen
        $msg = "Connect-PnPOnline misslyckades: $($_.Exception.Message)"
        Update-Splash $msg
        $global:SpError = $msg
    }
}

# === Inställningar ===
$ScriptVersion = "v40.1"
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

$Script1Path  = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Kontrollprovsfil 2025\Script Raw Data\Kontrollprovsfil_EPPlus_2025.ps1'
$Script2Path  = ''
$Script3Path  = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project\rename_CLI.ps1'

# === Centraliserad konfiguration ===
$Config = [ordered]@{
    # Källfiler (kan lämnas tomma; fylls i via GUI)
    CsvPath       = ''           # Sökväg till CSV-fil
    SealNegPath   = ''           # Sökväg till Seal Test NEG
    SealPosPath   = ''           # Sökväg till Seal Test POS
    WorksheetPath = ''           # Sökväg till LSP worksheet (Worksheet.xlsx)

    # SharePoint-inställningar – används för PnP-auth. Ärvd från globala variabler för bakåtkompatibilitet
    SiteUrl      = $global:SP_SiteUrl
    Tenant       = $global:SP_Tenant
    ClientId     = $global:SP_ClientId
    Certificate  = $global:SP_CertBase64

    # EPPlus – ange var lokalt DLL ska ligga och vilken version som ska hämtas om DLL saknas
    EpplusDllPath = (Join-Path $PSScriptRoot 'EPPlus.dll')
    EpplusVersion = '4.5.3.3'
}

$script:GXINF_Map = @{
    'Infinity-VI'   = '847922'
    'Infinity-VIII' = '803094'
    'GX5'           = '750210,750211,750212,750213'
    'GX6'           = '750246,750247,750248,750249'
    'GX1'           = '709863,709864,709865,709866'
    'GX2'           = '709951,709952,709953,709954'
    'GX3'           = '710084,710085,710086,710087'
    'GX7'           = '750170,750171,750172,750213'
    'Infinity-I'    = '802069'
    'Infinity-III'  = '807363'
    'Infinity-V'    = '839032'
}

# === Fallback-funktioner ===

# Filtrering i SharePoint. {BatchNumber} och {LSP} ersätts automatiskt.
$SharePointBatchLinkTemplate = 'https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management/Lists/Cepheid%20%20Production%20orders/ROBAL.aspx?viewid=6c9e53c9-a377-40c1-a154-13a13866b52b&view=7&q={BatchNumber}'

# === Logg: alltid till $PSScriptRoot\Loggar ===
$DevLogDir = Join-Path $PSScriptRoot 'Loggar'
if (-not (Test-Path $DevLogDir)) { New-Item -ItemType Directory -Path $DevLogDir -Force | Out-Null }
$global:LogPath = Join-Path $DevLogDir ("$($env:USERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt")

# === Ikoner ===

function New-GlyphIcon {
    param(
        [ValidateSet('folder','search','report','tools','settings','help','info','open','exit','toggleon','toggleoff')]
        [string]$Kind,[int]$Size=20,[string]$Stroke='#34495E',[single]$PenW=1.8
    )
    $bmp = New-Object System.Drawing.Bitmap $Size,$Size,([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g=[System.Drawing.Graphics]::FromImage($bmp); $g.SmoothingMode='AntiAlias'; $g.Clear([System.Drawing.Color]::Transparent)
    $pen=New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml($Stroke)),$PenW
    $cx=$Size/2.0; $cy=$Size/2.0

    switch($Kind){
        'search' { $r=$Size*.32; $g.DrawEllipse($pen,$cx-$r,$cy-$r,2*$r,2*$r); $p1=New-Object Drawing.PointF ([single]($cx+$r*.7)),([single]($cy+$r*.7)); $p2=New-Object Drawing.PointF ([single]($p1.X+$Size*.22)),([single]($p1.Y+$Size*.22)); $g.DrawLine($pen,$p1,$p2) }
        'report' { $g.DrawRectangle($pen,4,3,$Size-8,$Size-6); 6,7,11,15 | % { $g.DrawLine($pen,6,$_,$Size-10,$_)} }
        'folder' { $g.DrawRectangle($pen,3,8,$Size-6,$Size-12); $g.DrawLine($pen,6,8,10,4); $g.DrawLine($pen,10,4,16,4); $g.DrawLine($pen,16,4,16,8) }
        'tools'  { $r=$Size*.18; $g.DrawArc($pen,$cx-$r,4,2*$r,2*$r,200,220); $g.DrawLine($pen,$cx,$Size*.18,$Size-5,$Size-5); $g.DrawEllipse($pen,$Size-7,$Size-7,3,3) }
        'settings' { $y=[int]$cy; $g.DrawLine($pen,3,$y,$Size-3,$y); $g.DrawEllipse($pen,$cx-4,$y-4,8,8) }
        'help'   { $g.DrawEllipse($pen,3,3,$Size-6,$Size-6); $g.DrawCurve($pen,@( (New-Object Drawing.PointF ([single]($cx-3)),([single]($cy-2))), (New-Object Drawing.PointF ([single]$cx),([single]($cy-5))), (New-Object Drawing.PointF ([single]($cx+3)),([single]($cy-2))) )); $g.DrawLine($pen,$cx,$cy,$cx,$cy+4) }
        'info'   { $g.DrawEllipse($pen,3,3,$Size-6,$Size-6); $g.DrawLine($pen,$cx,$cy-2,$cx,$cy+5); $g.DrawEllipse($pen,$cx-0.8,$cy-6.8,1.6,1.6) }
        'open'   { $g.DrawRectangle($pen,4,6,$Size-12,$Size-10); $g.DrawLine($pen,$Size-8,$cy,$Size-4,$cy); $g.DrawLine($pen,$Size-7,$cy-3,$Size-4,$cy); $g.DrawLine($pen,$Size-7,$cy+3,$Size-4,$cy) }
        'exit'   { $m=5; $g.DrawLine($pen,$m,$m,$Size-$m,$Size-$m); $g.DrawLine($pen,$Size-$m,$m,$m,$Size-$m) }

        # --- Nya: toggle-off / toggle-on ---
        'toggleoff' {
            # En enkel “brytarknapp” OFF (rundad kapsling, cirkel till vänster)
            $r= [single]($Size*0.38)
            $rect = New-Object Drawing.RectangleF ([single]($cx-$r)),([single]($cy-6)),([single](2*$r)),([single]12)
            $g.DrawArc($pen,$rect.X,$rect.Y,12,12,90,180)
            $g.DrawArc($pen,$rect.X+$rect.Width-12,$rect.Y,12,12,270,180)
            $g.DrawLine($pen,$rect.X+6,$rect.Y,$rect.X+$rect.Width-6,$rect.Y)
            $g.DrawLine($pen,$rect.X+6,$rect.Y+$rect.Height,$rect.X+$rect.Width-6,$rect.Y+$rect.Height)
            $g.DrawEllipse($pen,$rect.X+2,$cy-4,8,8)
        }
        'toggleon' {
            # Samma brytare ON (cirkel till höger)
            $r= [single]($Size*0.38)
            $rect = New-Object Drawing.RectangleF ([single]($cx-$r)),([single]($cy-6)),([single](2*$r)),([single]12)
            $g.DrawArc($pen,$rect.X,$rect.Y,12,12,90,180)
            $g.DrawArc($pen,$rect.X+$rect.Width-12,$rect.Y,12,12,270,180)
            $g.DrawLine($pen,$rect.X+6,$rect.Y,$rect.X+$rect.Width-6,$rect.Y)
            $g.DrawLine($pen,$rect.X+6,$rect.Y+$rect.Height,$rect.X+$rect.Width-6,$rect.Y+$rect.Height)
            $g.DrawEllipse($pen,$rect.X+$rect.Width-10,$cy-4,8,8)
        }
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

    $it.add_Click({
        param($s,$e)
        $t = [string]$s.Tag
        try {
            if ($t -match '^(?i)https?://') {
                Start-Process $t
            }
            elseif (Test-Path -LiteralPath $t) {
                $gi = Get-Item -LiteralPath $t
                if ($gi.PSIsContainer) {
                    Start-Process explorer.exe -ArgumentList "`"$t`""
                } else {
                    Start-Process -FilePath $t
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("Hittar inte sökvägen:`n$t","Genväg",'OK','Warning') | Out-Null
            }
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
Update-Splash "Startar gränssnitt…"
Close-Splash
$form = New-Object System.Windows.Forms.Form
$form.Text = "BETA $ScriptVersion - verktyg vid dokumentation"
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
$miScript2   = New-Object System.Windows.Forms.ToolStripMenuItem('❌ TBD')
$miScript3   = New-Object System.Windows.Forms.ToolStripMenuItem('📅 Ändra datum-prefix för filnamn')
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
$lblTitle.Text="BETA $ScriptVersion - verktyg vid dokumentation"
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

# CSV
$lblCsv=$null;$clbCsv=$null;$btnCsvBrowse=$null
New-ListRow -labelText 'CSV:' -lbl ([ref]$lblCsv) -clb ([ref]$clbCsv) -btn ([ref]$btnCsvBrowse)
# NEG
$lblNeg=$null;$clbNeg=$null;$btnNegBrowse=$null
New-ListRow -labelText 'Seal NEG:' -lbl ([ref]$lblNeg) -clb ([ref]$clbNeg) -btn ([ref]$btnNegBrowse)
# POS
$lblPos=$null;$clbPos=$null;$btnPosBrowse=$null
New-ListRow -labelText 'Seal POS:' -lbl ([ref]$lblPos) -clb ([ref]$clbPos) -btn ([ref]$btnPosBrowse)
# Säkerställ att vi har plats för en fjärde rad (idempotent)
try {
    if ($tlPick.RowCount -lt 4) {
        $tlPick.RowCount = 4
        # Lägg till rad-styles upp till 4
        for ($i=$tlPick.RowStyles.Count; $i -lt 4; $i++) {
            $null = $tlPick.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 78)) )
        }
        # Höjdjustering av gruppen (samma stil som hos dig)
        $grpPick.Height = (78*4) + $grpPick.Padding.Top + $grpPick.Padding.Bottom + 15
    }
} catch {}

# Skapa kontroller (återanvänder din helper New-ListRow)
$lblLsp = $null; $clbLsp = $null; $btnLspBrowse = $null
New-ListRow -labelText 'Worksheet:' -lbl ([ref]$lblLsp) -clb ([ref]$clbLsp) -btn ([ref]$btnLspBrowse)

# Lägg in i tabellen, radindex 3 (0-baserat: 0=CSV, 1=NEG, 2=POS, 3=WS)
$tlPick.Controls.Add($lblLsp,  0, 3)
$tlPick.Controls.Add($clbLsp,  1, 3)
$tlPick.Controls.Add($btnLspBrowse, 2, 3)

# Single-select i denna CheckedListBox (utan att röra ev. bef. $onExclusive)
$clbLsp.add_ItemCheck({
    param($s,$e)
    if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
        for ($i=0; $i -lt $s.Items.Count; $i++) {
            if ($i -ne $e.Index) { $s.SetItemChecked($i, $false) }
        }
    }
})

# Browse-knapp: välj valfri worksheet .xlsx/.xlsm
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
    param(
        [string] $Text,
        [ValidateSet('Info','Warn','Error')][string] $Severity = 'Info',
        [switch] $Immediate
    )

    $prefix    = switch ($Severity) { 'Warn' {'⚠️'} 'Error' {'❌'} default {'ℹ️'} }
    $timestamp = (Get-Date).ToString('HH:mm:ss')
    $line      = "[$timestamp] $prefix $Text"

    $append = {
        $outputBox.AppendText("$line`r`n")
        $outputBox.SelectionStart = $outputBox.TextLength
        $outputBox.ScrollToCaret()
        $outputBox.Refresh()
    }

    # UI-safe uppdatering (om du i framtiden loggar från annan tråd)
    if ($outputBox.InvokeRequired) {
        $null = $outputBox.BeginInvoke([System.Windows.Forms.MethodInvoker]$append)
    } else {
        & $append
    }

    if ($Immediate) {
        # Rita om direkt så texten hinner visas innan UI:t ev. fryser
        [System.Windows.Forms.Application]::DoEvents()
    }

    if ($global:LogPath) {
        try { Add-Content -Path $global:LogPath -Value $line -Encoding UTF8 } catch {}
    }
}

# === EPPlus ===
function Ensure-EPPlus {
    param(
        [string] $Version = "4.5.3.3",
        [string] $SourceDllPath = "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Modules\EPPlus\EPPlus.4.5.3.3\.5.3.3\lib\net35\EPPlus.dll",
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
        if ($dllCandidates) {
            try {
                # Om lokala EPPlus.dll inte redan finns, kopiera nedladdad DLL till skriptmappen för att återanvändas nästa gång
                if (-not (Test-Path -LiteralPath $localScriptDll)) {
                    Copy-Item -Path $dllCandidates.FullName -Destination $localScriptDll -Force -ErrorAction SilentlyContinue
                }
            } catch {}
            return $localScriptDll
        }
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

# === CSV Instrument map + stats (robust PS 5.1) ===
if (-not (Get-Variable -Name GXINF_Map -Scope Script -ErrorAction SilentlyContinue)) {
    $script:GXINF_Map = @{
        'Infinity-VI'   = '847922'
        'Infinity-VIII' = '803094'
        'GX5'           = '750210,750211,750212,750213'
        'GX6'           = '750246,750247,750248,750249'
        'GX1'           = '709863,709864,709865,709866'
        'GX2'           = '709951,709952,709953,709954'
        'GX3'           = '710084,710085,710086,710087'
        'GX7'           = '750170,750171,750172,750213'
        'Infinity-I'    = '802069'
        'Infinity-III'  = '807363'
        'Infinity-V'    = '839032'
    }
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

function ConvertTo-CsvFields {
    param([string]$Line)
    # CSV-split som respekterar citat
    return [regex]::Split($Line, ',(?=(?:[^"]*"[^"]*")*[^"]*$)')
}

function Get-CsvStats {
    <#
      Returnerar objekt:
      TestCount, DupCount, Duplicates[], LspValues[], LspOK, InstrumentByType (ordered)
    #>
    param([string]$Path)

    $out = [ordered]@{
        TestCount       = 0
        DupCount        = 0
        Duplicates      = @()
        LspValues       = @()
        LspOK           = $null
        InstrumentByType= [ordered]@{}
    }

    if (-not (Test-Path -LiteralPath $Path)) { return [pscustomobject]$out }

    $lines = Get-Content -LiteralPath $Path
    if (-not $lines -or $lines.Count -lt 8) { return [pscustomobject]$out }

    # Försök tolka headern från rad 8 (index 7)
    $header = ConvertTo-CsvFields $lines[7]
    $dataLines = @()
    if ($lines.Count -gt 9) { $dataLines = $lines[9..($lines.Count-1)] }

    # Primär: använd headernamn om de finns
    $csv = $null
    try {
        if ($dataLines.Count -gt 0) {
            $csv = ConvertFrom-Csv -InputObject ($dataLines -join "`n") -Header $header
        }
    } catch { $csv = $null }

    # Fångare
    $cartSnList = New-Object System.Collections.Generic.List[string]
    $lspList    = New-Object System.Collections.Generic.List[string]
    $instrList  = New-Object System.Collections.Generic.List[string]

    if ($csv) {
        foreach ($row in $csv) {
            # Tolerant hämtning via headernamn OCH fallback på index
            $cart = $row.'Cartridge S/N'
            $lsp  = $row.'Reagent Lot ID'  # fallback kolumn E (0-baserad index 4)
            $ins  = $row.'Instrument S/N'

            if (-not $cart) { try { $cart = $row.Item(3) } catch {} }  # kolumn D
            if (-not $ins)  { try { $ins  = $row.Item(6) } catch {} }  # kolumn G

            if ($cart) { $cartSnList.Add( ($cart + '').Trim() ) }
            if ($lsp)  { $lspList.Add(  ($lsp  + '').Trim() ) }
            if ($ins)  { $instrList.Add(($ins  + '').Trim() ) }
        }
    } else {
        # Helt utan ConvertFrom-Csv: plocka kolumn D/E/G via index
        foreach ($ln in $dataLines) {
            if (-not $ln -or -not $ln.Trim()) { continue }
            $f = ConvertTo-CsvFields $ln
            if ($f.Count -lt 7) { continue }
            $cartSnList.Add( ($f[3] + '').Trim() )
            $lspList.Add(    ($f[4] + '').Trim() )
            $instrList.Add(  ($f[6] + '').Trim() )
        }
    }

    # Testcount = antal icke-tomma Cartridge S/N
    $cartSn = $cartSnList | Where-Object { $_ -and $_ -ne '' }
    $out.TestCount = @($cartSn).Count

    # Dubbletter
    $dups = $cartSn | Group-Object | Where-Object { $_.Count -gt 1 }
    if ($dups) {
        $out.DupCount   = $dups.Count
        $out.Duplicates = $dups | ForEach-Object { "$($_.Name) x$($_.Count)" }
    }

    # LSP (unika femsiffriga — tolerant)
    $lspClean = $lspList | ForEach-Object {
        $m = [regex]::Match($_, '(?<!\d)(\d{5})(?!\d)')
        if ($m.Success) { $m.Groups[1].Value } else { $null }
    } | Where-Object { $_ } | Select-Object -Unique
    $out.LspValues = $lspClean
    if ($lspClean.Count -gt 0) {
        $out.LspOK = ($lspClean.Count -eq 1)
    }

    # Instrument → typ
    # Förbered lookups
    $lut = @{}
    foreach ($k in $script:GXINF_Map.Keys) {
        $codes = $script:GXINF_Map[$k].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        foreach ($code in $codes) { $lut[$code] = $k }
    }
    foreach ($ins in ($instrList | Where-Object { $_ })) {
        $t = $null
        if ($lut.ContainsKey($ins)) { $t = $lut[$ins] } else { $t = 'Unknown' }
        if (-not $out.InstrumentByType.Contains($t)) { $out.InstrumentByType[$t] = 0 }
        $out.InstrumentByType[$t]++
    }

    return [pscustomobject]$out
}

function Format-SpPresenceGrandTotalStrict {
    param([hashtable]$Counts)

    if (-not $Counts -or $Counts.Count -eq 0) { return '—' }

    # Hämta SP som har >0 Infinity-körda test (0..10)
    $present = @(
        $Counts.GetEnumerator() |
        Where-Object { $_.Value -gt 0 } |
        ForEach-Object { [int]$_.Key } |
        Sort-Object
    )
    if ($present.Count -eq 0) { return '—' }

    # Komprimera närliggande SP: #00-05, singlar: #06
    $parts = New-Object System.Collections.Generic.List[string]
    $i = 0
    while ($i -lt $present.Count) {
        $start = $present[$i]; $end = $start
        $j = $i + 1
        while ($j -lt $present.Count -and $present[$j] -eq $end + 1) {
            $end = $present[$j]; $j++
        }
        if ($start -eq $end) { $parts.Add( ('#{0:00}' -f $start) ) }
        else { $parts.Add( ('#{0:00}-{1:00}' -f $start, $end) ) }
        $i = $j
    }

    # Total: summan av alla Infinity-körda test över listade SP
    $total = (
        $Counts.GetEnumerator() |
        Where-Object { $_.Value -gt 0 } |
        Measure-Object -Property Value -Sum
    ).Sum

    return ( ($parts -join '+') + (' x{0}' -f $total) )
}

if (Get-Command Get-InfinitySpFromCsvStrict -ErrorAction SilentlyContinue) {
    Remove-Item Function:\Get-InfinitySpFromCsvStrict -ErrorAction SilentlyContinue
}
function Get-InfinitySpFromCsvStrict {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        # för kompatibilitet med dina anrop
        [Alias('InfinitySerials')][string[]]$InstrumentSerials
    )

    if (-not (Test-Path -LiteralPath $Path)) { return '—' }

    # Använd din robusta ConvertTo-CsvFields om den redan finns; annars fallback-split
    $useConvertFn = $false
    try { $useConvertFn = [bool](Get-Command ConvertTo-CsvFields -ErrorAction Stop) } catch {}

    function Split-CsvSmart([string]$ln) {
        if ($ln -like '*;*' -and $ln -notlike '*,*') {
            return [regex]::Split($ln, ';(?=(?:[^"]*"[^"]*")*[^"]*$)')
        } else {
            return [regex]::Split($ln, ',(?=(?:[^"]*"[^"]*")*[^"]*$)')
        }
    }

    # Normalisera Infinity-serier (Instrument S/N)
    $serSet = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
    foreach ($s in ($InstrumentSerials | Where-Object { $_ })) { $null = $serSet.Add( ($s + '').Trim().Trim('"') ) }
    if ($serSet.Count -eq 0) { return '—' }

    $lines = Get-Content -LiteralPath $Path
    if (-not $lines -or $lines.Count -lt 10) { return '—' }

    # *** FAST LAYOUT enligt dina regler ***
    # header = rad 8 (index 7), data = rad 10 (index 9) och nedåt
    $headerIndex = 7
    $dataStart   = 9

    $hdr = if ($useConvertFn) { ConvertTo-CsvFields $lines[$headerIndex] } else { Split-CsvSmart $lines[$headerIndex] }
    $colCount = $hdr.Count

    # Instrument S/N via header; fallback till G (index 6)
    $idxInstr = -1
    for ($i=0; $i -lt $colCount; $i++) {
        $h = (($hdr[$i] + '') -replace '[\uFEFF\u200B]','').Trim().ToLowerInvariant()
        if ($h -match 'instrument' -and ($h -match 's/?n' -or $h -match 'serial')) { $idxInstr = $i; break }
    }
    if ($idxInstr -lt 0) { $idxInstr = 6 }  # G

    # Sample ID = kolumn C (index 2)
    $idxSample = 2

    # Räkna replikat per SP (endast Infinity-rader)
    $counts = @{}   # nyckel: int (0..10) där 0 motsvarar _00_, värde: antal
    for ($r = $dataStart; $r -lt $lines.Count; $r++) {
        $ln = $lines[$r]; if (-not $ln -or -not $ln.Trim()) { continue }
        $f = if ($useConvertFn) { ConvertTo-CsvFields $ln } else { Split-CsvSmart $ln }
        if ($f.Count -le [Math]::Max($idxInstr,$idxSample)) { continue }

        $instr  = ($f[$idxInstr] + '').Trim().Trim('"')
        if (-not $serSet.Contains($instr)) { continue }  # inte Infinity → skippa

        $sample = ($f[$idxSample] + '').Trim().Trim('"')
        if (-not $sample) { continue }

        # *** Viktigt: offset-regel ***
        # Matcha exakt _00_.._10_ i Sample ID (kolumn C)
        $m = [regex]::Match($sample, '_(\d{2})_')
         if ($m.Success) {
             $nRaw = 0
             if ([int]::TryParse($m.Groups[1].Value, [ref]$nRaw)) {

                $sp = $nRaw            # 00..10 (ingen offset)
                if ($sp -ge 0 -and $sp -le 10) {
                     if (-not $counts.ContainsKey($sp)) { $counts[$sp] = 0 }
                     $counts[$sp]++
                 }
             }
         }
         }

    if ($counts.Count -eq 0) { return '—' }
    return (Format-SpPresenceGrandTotalStrict -Counts $counts)
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

if (-not (Get-Command Extract-WorksheetHeader -ErrorAction SilentlyContinue)) {
    function Extract-WorksheetHeader {
        param([OfficeOpenXml.ExcelPackage]$Pkg)

        $result = [pscustomobject]@{
            PartNo         = $null
            BatchNo        = $null
            CartridgeNo    = $null
            DocumentNumber = $null
            Attachment     = $null
            Rev            = $null
            Effective      = $null
        }

        if (-not $Pkg) { return $result }

        foreach ($ws in $Pkg.Workbook.Worksheets) {
        if ($ws.Name -eq 'Worksheet Instructions' -or
        $ws.Name -match '(?i)for reference only') {
        continue
        }

            # Hämta rå header-text (vänster + höger)
            $left  = (($ws.HeaderFooter.OddHeader.LeftAlignedText  + '') -replace '\r?\n',' ')
            if (-not $left)  { $left  = (($ws.HeaderFooter.EvenHeader.LeftAlignedText  + '') -replace '\r?\n',' ') }
            $right = (($ws.HeaderFooter.OddHeader.RightAlignedText + '') -replace '\r?\n',' ')
            if (-not $right) { $right = (($ws.HeaderFooter.EvenHeader.RightAlignedText + '') -replace '\r?\n',' ') }

            $raw = (($left + ' ' + $right) + '').Trim()
            if (-not $raw) { continue }

            $parsed = Parse-WorksheetHeaderRaw -Raw $raw

            if (-not $result.PartNo         -and $parsed.PartNo)         { $result.PartNo         = $parsed.PartNo }
            if (-not $result.BatchNo        -and $parsed.BatchNo)        { $result.BatchNo        = $parsed.BatchNo }
            if (-not $result.CartridgeNo    -and $parsed.CartridgeNo)    { $result.CartridgeNo    = $parsed.CartridgeNo }
            if (-not $result.DocumentNumber -and $parsed.DocumentNumber) { $result.DocumentNumber = $parsed.DocumentNumber }
            if (-not $result.Attachment     -and $parsed.Attachment)     { $result.Attachment     = $parsed.Attachment }
            if (-not $result.Rev            -and $parsed.Rev)            { $result.Rev            = $parsed.Rev }
            if (-not $result.Effective      -and $parsed.Effective)      { $result.Effective      = $parsed.Effective }
        }

        # Extra fallback: Cartridge om fortfarande null → hitta första fristående 5-siffriga talet
        if (-not $result.CartridgeNo) {
            foreach ($ws in $Pkg.Workbook.Worksheets) {
                if ($ws.Name -eq 'Worksheet Instructions') { continue }
                $left  = (($ws.HeaderFooter.OddHeader.LeftAlignedText  + '') -replace '\r?\n',' ').Trim()
                $right = (($ws.HeaderFooter.OddHeader.RightAlignedText + '') -replace '\r?\n',' ').Trim()
                $m  = [regex]::Match($left,  '(?<!\d)(\d{5})(?!\d)')
                $m2 = if (-not $m.Success) { [regex]::Match($right,'(?<!\d)(\d{5})(?!\d)') } else { $null }
                if ($m.Success) { $result.CartridgeNo = $m.Groups[1].Value; break }
                if ($m2 -and $m2.Success) { $result.CartridgeNo = $m2.Groups[1].Value; break }
            }
        }

        return $result
    }
}

if (-not (Get-Command Get-WorksheetHeaderPerSheet -ErrorAction SilentlyContinue)) {
    function Get-WorksheetHeaderPerSheet {
        param([OfficeOpenXml.ExcelPackage]$Pkg)

        # --- etikettmönster för cell-sök ---
        $rxPart      = '(?i)^\s*Part\s*(?:No|Number)\.?\s*(?:\(s\))?\s*$'
        $rxBatch     = '(?i)^\s*Batch\s*(?:No|Number)(?:\s*\(s\))?\.?\s*$'
        $rxCartridge = '(?i)^\s*Cartridge\s*(?:No|Number)?\s*(?:\(?LSP\)?)?\.?\s*$'
        $rxDoc       = '(?i)^\s*Document\s*(?:No|Number|#)\s*$'
        $rxRev       = '(?i)^\s*Rev(?:ision)?\.?\s*$'
        $rxEff       = '(?i)^\s*Effective(?:\s*Date)?\s*$'

        function Get-HeaderFooterText([Object]$ws) {
            $left  = Normalize-HeaderText ((($ws.HeaderFooter.OddHeader.LeftAlignedText  + '') -replace '\r?\n',' '))
            if (-not $left)  { $left  = Normalize-HeaderText ((($ws.HeaderFooter.EvenHeader.LeftAlignedText  + '') -replace '\r?\n',' ')) }
            $right = Normalize-HeaderText ((($ws.HeaderFooter.OddHeader.RightAlignedText + '') -replace '\r?\n',' '))
            if (-not $right) { $right = Normalize-HeaderText ((($ws.HeaderFooter.EvenHeader.RightAlignedText + '') -replace '\r?\n',' ')) }
            return @($left,$right)
        }


        function Try-FromHeaderFooter([string[]]$lr, [string]$kind) {
            $left,$right = $lr
            $raw = (($left + ' ' + $right) + '').Trim()

            if (-not $raw) {
                return @{ Has = $false; Val = $null }
            }

            $parsed = Parse-WorksheetHeaderRaw -Raw $raw

            $has = $false
            $val = $null

            switch ($kind) {
                'Part' {
                    $has = ($raw -match '(?i)\bPart\b')
                    $val = $parsed.PartNo
                }
                'Batch' {
                    $has = ($raw -match '(?i)\bBatch\b')
                    $val = $parsed.BatchNo
                }
                'Cartridge' {
                    $has = ($raw -match '(?i)\bCartridge\b')
                    $val = $parsed.CartridgeNo
                }
                'Doc' {
                    $has = ($raw -match '(?i)Document\s*Number')
                    $val = $parsed.DocumentNumber
                }
                'REV' {
                    $has = ($raw -match '(?i)\bRev(?:ision)?\b')
                    $val = $parsed.Rev
                }
                'EFF' {
                    $has = ($raw -match '(?i)\bEffective(\s*Date)?\b')
                    $val = $parsed.Effective
                }
            }

            return @{ Has = [bool]$has; Val = $val }
        }

        function Try-FromCells([Object]$ws, [string]$rxLabel, [int]$maxR=30, [int]$maxC=20, [string]$canon='') {
            $has=$false; $val=$null
            for($r=1;$r -le $maxR;$r++){
                for($c=1;$c -le $maxC;$c++){
                    $t     = Normalize-HeaderText (""+$ws.Cells[$r,$c].Text)
                    if([string]::IsNullOrWhiteSpace($t)){ continue }
                    if($t -match $rxLabel){
                        $has=$true
$right = Normalize-HeaderText (""+$ws.Cells[$r,[Math]::Min($c+1,$maxC)].Text)
$down  = Normalize-HeaderText (""+$ws.Cells[[Math]::Min($r+1,$maxR),$c].Text)
                        $val = if($right){$right}elseif($down){$down}else{''}
                        if($canon -eq 'Cartridge' -and $val){
                        elseif ($canon -eq 'EFF' -and $val) {
    $dt = $null
    if (Try-Parse-HeaderDate $val ([ref]$dt)) { $val = $dt.ToString('yyyy-MM-dd') }
}

                            $m=[regex]::Match($val,'(?<!\d)(\d{5})(?!\d)'); if($m.Success){$val=$m.Groups[1].Value}
                        }
                        return @{Has=$has; Val=$val}
                    }
                }
            }
            return @{Has=$has; Val=$val}
        }

        $rows=@()
        if(-not $Pkg){ return $rows }

foreach ($ws in $Pkg.Workbook.Worksheets) {
    if ($ws.Name -eq 'Worksheet Instructions' -or
        $ws.Name -eq 'Test Summary'           -or
        $ws.Name -match '(?i)for reference only') {
        continue
    }

            $partLabel=$false; $batchLabel=$false; $cartLabel=$false; $docLabel=$false; $revLabel=$false; $effLabel=$false
            $part=$null; $batch=$null; $cart=$null; $doc=$null; $rev=$null; $eff=$null

            $lr = Get-HeaderFooterText $ws

            # Header/footer först…
            $r1 = Try-FromHeaderFooter $lr 'Part';      $partLabel=$partLabel -or $r1.Has; if($r1.Val){$part=$r1.Val}
            $r2 = Try-FromHeaderFooter $lr 'Batch';     $batchLabel=$batchLabel -or $r2.Has; if($r2.Val){$batch=$r2.Val}
            $r3 = Try-FromHeaderFooter $lr 'Cartridge'; $cartLabel=$cartLabel -or $r3.Has; if($r3.Val){$cart=$r3.Val}
            $r4 = Try-FromHeaderFooter $lr 'Doc';       $docLabel=$docLabel -or $r4.Has; if($r4.Val){$doc=$r4.Val}
            $r5 = Try-FromHeaderFooter $lr 'REV';       $revLabel=$revLabel -or $r5.Has; if($r5.Val){$rev=$r5.Val}
            $r6 = Try-FromHeaderFooter $lr 'EFF';       $effLabel=$effLabel -or $r6.Has; if($r6.Val){$eff=$r6.Val}

            # …fallback: cells (om inget värde hittades eller etikett saknas)
            if(-not $part){  $c=Try-FromCells $ws $rxPart      30 20 'Part';      $partLabel=$partLabel -or $c.Has; if($c.Val){$part=$c.Val} }
            if(-not $batch){ $c=Try-FromCells $ws $rxBatch     30 20 'Batch';     $batchLabel=$batchLabel -or $c.Has; if($c.Val){$batch=$c.Val} }
            if(-not $cart){  $c=Try-FromCells $ws $rxCartridge 30 20 'Cartridge'; $cartLabel=$cartLabel -or $c.Has; if($c.Val){$cart=$c.Val} }
            if(-not $doc){   $c=Try-FromCells $ws $rxDoc       30 20;             $docLabel=$docLabel -or $c.Has; if($c.Val){$doc=$c.Val} }
            if(-not $rev){   $c=Try-FromCells $ws $rxRev       30 20;             $revLabel=$revLabel -or $c.Has; if($c.Val){$rev=$c.Val} }
            if(-not $eff){   $c=Try-FromCells $ws $rxEff       30 20 'EFF';       $effLabel=$effLabel -or $c.Has; if($c.Val){$eff=$c.Val} }

            $rows += [pscustomobject]@{
                Sheet              = $ws.Name
                PartNo             = $part
                BatchNo            = $batch
                CartridgeNo        = $cart
                DocumentNumber     = $doc
                Rev                = $rev
                Effective          = $eff
                HasPartNoLabel     = [bool]$partLabel
                HasBatchNoLabel    = [bool]$batchLabel
                HasCartridgeLabel  = [bool]$cartLabel
                HasDocumentLabel   = [bool]$docLabel
                HasRevLabel        = [bool]$revLabel
                HasEffectiveLabel  = [bool]$effLabel
            }
        }
        return $rows
    }
}

if (-not (Get-Command Compare-WorksheetHeaderSet -ErrorAction SilentlyContinue)) {
    function Compare-WorksheetHeaderSet {
        param([Parameter(Mandatory)][object[]]$Rows)

        $devIssues  = 0
        $reqIssues  = 0
        $detailsLst = New-Object System.Collections.Generic.List[string]

function _canon([string]$raw, [string]$type) {
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    $txt = Normalize-HeaderText $raw
            switch ($type) {
                'Part'      { $m=[regex]::Match($txt,'(?i)\b(\d{3}-\d{4})\b'); return $(if($m.Success){$m.Groups[1].Value}else{$txt.ToUpper()}) }
                'Batch'     { $m=[regex]::Match($txt,'(?<!\d)(\d{10})(?!\d)');  return $(if($m.Success){$m.Groups[1].Value}else{$txt.ToUpper()}) }
                'Cartridge' { $m=[regex]::Match($txt,'(?<!\d)(\d{5})(?!\d)');   return $(if($m.Success){$m.Groups[1].Value}else{$txt.ToUpper()}) }
                'Doc'       { $m=[regex]::Match($txt,'(?i)\b(D\d+)\b');         return $(if($m.Success){$m.Groups[1].Value.ToUpper()}else{$txt.ToUpper()}) }
                'REV'       { return $txt.ToUpper() }
'EFF' {
    $dt = $null
    if (Try-Parse-HeaderDate $txt ([ref]$dt)) { return $dt.ToString('yyyy-MM-dd') }
    return $txt
}
                default     { return $txt }
            }
        }

        $keys = @(
            @{K='PartNo';         T='Part';      Required=$true;  Label='HasPartNoLabel'    },
            @{K='BatchNo';        T='Batch';     Required=$true;  Label='HasBatchNoLabel'   },
            @{K='CartridgeNo';    T='Cartridge'; Required=$true;  Label='HasCartridgeLabel' },
            @{K='DocumentNumber'; T='Doc';       Required=$false; Label='HasDocumentLabel'  },
            @{K='Rev';            T='REV';       Required=$false; Label='HasRevLabel'       },
            @{K='Effective';      T='EFF';       Required=$false; Label='HasEffectiveLabel' }
        )

        foreach ($entry in $keys) {
            $k=$entry.K; $t=$entry.T; $req=$entry.Required; $labelProp=$entry.Label

            $vals = foreach($r in $Rows){
                [pscustomobject]@{
                    Sheet  = $r.Sheet
                    Canon  = _canon ($r.$k) $t
                    Raw    = if([string]::IsNullOrWhiteSpace($r.$k)){ $null } else { $r.$k.Trim() }
                    HasLbl = [bool]($r.$labelProp)
                }
            }

            # Majoritet (ignorera tomma)
            $nonEmpty = $vals | Where-Object { $_.Canon }
            $useSet   = if($nonEmpty.Count -gt 0){$nonEmpty}else{$vals}
            $byVal    = $useSet | Group-Object Canon | Sort-Object Count -Descending
            $major    = if($byVal.Count -gt 0){ $byVal[0].Name } else { $null }

            # Avvikare = bara blad med värde som != major
            $deviants = @()
            if($major){
                $deviants = $vals | Where-Object { $_.Canon -and ($_.Canon -ne $major) } | Select-Object -ExpandProperty Sheet -Unique
            }

            if($deviants.Count -gt 0){
                $devIssues++
                $majShow = if($major){"'$major'"}else{'(empty)'}
                $line = "- $k majoritet=$majShow | avvikande flikar: " + ($deviants -join ', ')
                [void]$detailsLst.Add($line)
            }

            # Saknas: endast där etikett finns på bladet men värdet saknas
            if($req){
                $missingSheets = $vals | Where-Object { $_.HasLbl -and -not $_.Canon } | Select-Object -ExpandProperty Sheet -Unique
                if($missingSheets.Count -gt 0){
                    $reqIssues++
                    $line = "*Saknas* - $k " + ($missingSheets -join ', ')
                    [void]$detailsLst.Add($line)
                }
            }
        }

        $issuesTotal = $devIssues + $reqIssues
        $summary = if($issuesTotal -eq 0){ 'OK' } else { "Avvikelser: $issuesTotal (avvikare=$devIssues, saknas=$reqIssues)" }

        return [pscustomobject]@{
            Issues  = $issuesTotal
            Summary = $summary
            Details = ($detailsLst -join "`r`n")
        }
    }
}


if (-not (Get-Command Extract-SealTestHeader -ErrorAction SilentlyContinue)) {
    function Extract-SealTestHeader {
        param([OfficeOpenXml.ExcelPackage]$Pkg)

        $result = [pscustomobject]@{ DocumentNumber=$null; Rev=$null; Effective=$null }
        if (-not $Pkg) { return $result }

        foreach ($ws in $Pkg.Workbook.Worksheets) {
            if ($ws.Name -eq 'Worksheet Instructions') { continue }
            $right = (($ws.HeaderFooter.OddHeader.RightAlignedText + '') -replace '\r?\n',' ').Trim()
            if (-not $right) { $right = (($ws.HeaderFooter.EvenHeader.RightAlignedText + '') -replace '\r?\n',' ').Trim() }

            if (-not $result.DocumentNumber -and $right -match '(?i)\bDocument\s*(?:No|Number|#)\s*[:#]?\s*(D\d+)\b') { $result.DocumentNumber = $matches[1] }
            if (-not $result.Rev            -and $right -match '(?i)\bRev(?:ision)?\.?\s*[:#]?\s*([A-Z]{1,3}(?:\.\d+)?)\b') { $result.Rev = $matches[1] }
            if (-not $result.Effective      -and $right -match '(?i)\bEffective\s*[:#]?\s*([0-9]{1,2}[\/\-][0-9]{1,2}[\/\-][0-9]{4}|[0-9]{4}[\/\-][0-9]{2}[\/\-][0-9]{2})') { $result.Effective = $matches[1] }

            if ($result.DocumentNumber -and $result.Rev -and $result.Effective) { break }
        }
        return $result
    }
}

function Normalize-Id {
    param([string]$Value, [ValidateSet('Batch','Part','Cartridge')] [string]$Type)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $v = $Value.Trim()
    switch ($Type) {
        'Batch'     { return ($v -replace '[^\d]', '') }                      # 10 siffror
        'Part'      { return (($v -replace '[^0-9A-Za-z\-]', '')).ToUpper() } # t.ex. 700-5702
        'Cartridge' { return ($v -replace '[^\d]', '') }                      # 5 siffror
        default     { return $v }
    }
}

function Normalize-HeaderText {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return '' }
    # Ta bort BOM/zero-width
    $s = $s -replace "[\uFEFF\u200B]", ""
    # Ta bort inbäddade radbrytningsmarkörer från Excel-export
    # t.ex. 700-4370_x000D_ Batch No(s).: 1001502746_x000D_ ...
    $s = $s -replace "_x000D_", " "
    # NBSP/figur/narrow NBSP → vanligt mellanslag
    $s = $s -replace "[\u00A0\u2007\u202F]", " "

    # Normalisera olika streck till '-'
    $s = $s -replace "[\u2013\u2014\u2212]", "-"

    # Kollapsa whitespace och trimma
    $s = ($s -replace "\s+", " ").Trim()
    return $s
}

function Parse-WorksheetHeaderRaw {
    param(
        [string]$Raw
    )

    $obj = [pscustomobject]@{
        PartNo         = $null
        BatchNo        = $null
        CartridgeNo    = $null
        DocumentNumber = $null
        Attachment     = $null
        Rev            = $null
        Effective      = $null
    }

    if (-not $Raw) { return $obj }

    # Grundstädning
    $txt = $Raw + ''

    # Ta bort header-koder (sidnr m.m.)
    $txt = $txt -replace '&P', ''
    $txt = $txt -replace '&N', ''
    $txt = $txt -replace '&L', ' '

    # Normalisera (tar också bort _x000D_)
    $txt = Normalize-HeaderText $txt

    # Document Number + ev. Attachment
    $m = [regex]::Match(
        $txt,
        '(?i)\bDocument\s*Number\b\s*[:#]?\s*(D\d{5})(?:\s+Attachment[:\s]+([0-9A-Za-z\.]+))?'
    )
    if ($m.Success) {
        $obj.DocumentNumber = $m.Groups[1].Value.Trim()
        if ($m.Groups.Count -ge 3 -and $m.Groups[2].Value) {
            $obj.Attachment = $m.Groups[2].Value.Trim()
        }
    }

    # Fristående "Attachment" om det inte fångades ovan
    if (-not $obj.Attachment) {
        $m = [regex]::Match($txt, '(?i)\bAttachment\b\s*:?\s*([0-9A-Za-z\.]+)')
        if ($m.Success) {
            $obj.Attachment = $m.Groups[1].Value.Trim()
        }
    }

    # Rev (t.ex. F, AW, E.1)
    $m = [regex]::Match($txt, '(?i)\bRev(?:ision)?\b\s*:?\s*([A-Z0-9]+(?:\.[0-9]+)?)')
    if ($m.Success) {
        $obj.Rev = $m.Groups[1].Value.Trim()
    }

    # Effective (t.ex. 4/14/2025 eller 09/20/2023)
    $m = [regex]::Match($txt, '(\d{1,2}/\d{1,2}/\d{4})')
    if ($m.Success) {
        $obj.Effective = $m.Groups[1].Value.Trim()
    }

    # Part No (t.ex. 700-4370)
    $m = [regex]::Match($txt, '(?i)Part\s*No\.\s*:\s*([0-9A-Z\-]+)')
    if ($m.Success) {
        $obj.PartNo = $m.Groups[1].Value.Trim()
    }

    # Batch No (t.ex. 1001502746)
    $m = [regex]::Match($txt, '(?i)Batch\s*No\(s\)\.\s*:\s*(\d{6,12})')
    if ($m.Success) {
        $obj.BatchNo = $m.Groups[1].Value.Trim()
    }

    # Cartridge (LSP, t.ex. 78704)
    $m = [regex]::Match($txt, '(?i)Cartridge\s*No\.\s*\(LSP\)\s*:\s*(\d{5,7})')
    if ($m.Success) {
        $obj.CartridgeNo = $m.Groups[1].Value.Trim()
    }

    return $obj
}

function Try-Parse-HeaderDate {
    param([object]$cellValue, [ref]$out)
    $out.Value = $null
    if ($null -eq $cellValue) { return $false }

    if ($cellValue -is [double] -or $cellValue -is [int]) {
        try { $out.Value = [DateTime]::FromOADate([double]$cellValue); return $true } catch {}
    }
    if ($cellValue -is [DateTime]) { $out.Value = [DateTime]$cellValue; return $true }

    $s = Normalize-HeaderText ([string]$cellValue)
    if ([string]::IsNullOrWhiteSpace($s)) { return $false }

    $cultures = @('sv-SE','en-US','en-GB')
    foreach ($c in $cultures) {
        try {
            $dt = [DateTime]::Parse($s, [Globalization.CultureInfo]::GetCultureInfo($c), [Globalization.DateTimeStyles]::AssumeLocal)
            $out.Value = $dt; return $true
        } catch {}
    }
    $formats = @(
        'yyyy-MM-dd','yyyy/MM/dd','dd/MM/yyyy','d/M/yyyy','MM/dd/yyyy','M/d/yyyy',
        'dd-MM-yyyy','d-M-yyyy','dd.M.yyyy','d.M.yyyy','dd-MMM-yyyy','MMM-yy'
    )
    foreach ($fmt in $formats) {
        try {
            $dt = [DateTime]::ParseExact($s, $fmt, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeLocal)
            $out.Value = $dt; return $true
        } catch {}
    }
    return $false
}

function Get-ConsensusValue {
    param(
        [string]$Ws,  [string]$Pos, [string]$Neg,
        [ValidateSet('Batch','Part','Cartridge')] [string]$Type
    )

    $nWs  = Normalize-Id $Ws  $Type
    $nPos = Normalize-Id $Pos $Type
    $nNeg = Normalize-Id $Neg $Type

    # vilka källor har något värde?
    $present = @{}
    if ($nWs)  { $present['WS']  = $nWs  }
    if ($nPos) { $present['POS'] = $nPos }
    if ($nNeg) { $present['NEG'] = $nNeg }

    $chosenNorm = $null
    $sources    = @()

    # 1) Majoritet (minst två lika)
    $counts = @{}
    foreach ($k in $present.Keys) {
        $val = $present[$k]
        if (-not $counts.ContainsKey($val)) { $counts[$val] = 0 }
        $counts[$val]++
    }
    $maxCount = 0; $maxVal = $null
    foreach ($kv in $counts.GetEnumerator()) {
        if ($kv.Value -gt $maxCount) { $maxCount = $kv.Value; $maxVal = $kv.Key }
    }
    if ($maxCount -ge 2) {
        $chosenNorm = $maxVal
        foreach ($k in $present.Keys) { if ($present[$k] -eq $chosenNorm) { $sources += $k } }
    }

    # 2) Ingen majoritet? Om exakt en av POS/NEG matchar WS ⇒ välj WS
    if (-not $chosenNorm -and $nWs) {
        $posMatch = ($nPos -and $nPos -eq $nWs)
        $negMatch = ($nNeg -and $nNeg -eq $nWs)
        if (($posMatch -and -not $negMatch) -or ($negMatch -and -not $posMatch)) {
            $chosenNorm = $nWs
            $sources = @('WS')
            if ($posMatch) { $sources += 'POS' }
            if ($negMatch) { $sources += 'NEG' }
        }
    }

    # 3) POS==NEG (och WS saknas/avviker)
    if (-not $chosenNorm -and $nPos -and $nNeg -and $nPos -eq $nNeg) {
        $chosenNorm = $nPos
        $sources = @('POS','NEG')
    }

    # 4) Annars välj det som finns (prioritet WS > POS > NEG)
    if (-not $chosenNorm) {
        if ($nWs)      { $chosenNorm = $nWs;  $sources = @('WS')  }
        elseif ($nPos) { $chosenNorm = $nPos; $sources = @('POS') }
        elseif ($nNeg) { $chosenNorm = $nNeg; $sources = @('NEG') }
    }

    # “Pretty” original (behåll format från första källa i ordning WS>POS>NEG)
    $orig = @{ WS=$Ws; POS=$Pos; NEG=$Neg }
    $pretty = $null
    foreach ($pref in @('WS','POS','NEG')) {
        if ($present.ContainsKey($pref) -and $present[$pref] -eq $chosenNorm) { $pretty = $orig[$pref]; break }
    }

    # Note-text (frivillig)
    $note = $null
    if ($sources -and $sources.Count -gt 0) {
        $note = "Consensus: " + ($sources -join '+')
        $others = @()
        foreach ($k in @('WS','POS','NEG')) {
            if ($present.ContainsKey($k) -and ($sources -notcontains $k)) {
                $val = $orig[$k]
                if (-not [string]::IsNullOrEmpty($val)) { $others += ($k + '=' + $val) }
            }
        }
        if ($others.Count -gt 0) { $note += " | Others: " + ($others -join ', ') }
    }

    return @{ Value=$pretty; Source=($sources -join '+'); Note=$note }
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
    if (Get-CheckedFilePath $clbLsp) { $n++ }
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

$miToggleSign.add_Click({
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
   • Information
   • Seal Test Info
   • STF Sum
   • Equipment
   • Control Material
   • SharePoint Info

Tips:
   • Använd "Genvägar" för att snabbt hitta relevanta:
     - Dokument
     - IPT-mappar
     - Länkar
   • Använd "Verktyg" för att Skriv signatur till
     valda Seal Test-filer (sammanställning):
     - Verktyg → Aktivera Seal Test-signatur → Följ instruktion
     - Verktyg → Deaktivera Seal Test-signatur

Felsökning:
   • Filen låst → Stäng Excelfiler.

"@
    [System.Windows.Forms.MessageBox]::Show($msg,"Instruktioner") | Out-Null
})
$miFAQ.add_Click({
    $faq = @"

Vad gör skriptet?

Det skapar en excel-rapport som jämför sökt LSP för Seal Test-Filer,
hämtar utrustningslista och rätt kontrollmaterial för sökt LSP.

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
   • Välj “LSP-mapp” (samma mapp som ditt LSP) eller “TEMP” = sparas ej.

6) Hur fungerar Seal Test-signering?
   • Signatur hamnar på alla flikar med:
     - "Seal Test Data" och "Name of Tester"
   • Flikar utan data, signatur eller/och markerade som N/A hoppas över.

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
$miOm.add_Click({ [System.Windows.Forms.MessageBox]::Show("OBS! Detta verktyg är endast ett hjälpmedel och ersätter inte någon process hos PQC.`n BETA $ScriptVersion`nav Jesper","Om") | Out-Null })

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
# Removed unused function Get-AnyB47

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

# === Sök filer-knapp ===
$btnScan.Add_Click({
    Gui-Log '🔎 Söker filer…' -Immediate
    try {
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
    # Hitta LSP-worksheet (innehåller "worksheet" i filnamnet, matchar LSP och är en Excel-fil)
    # LSP‑worksheet: match files containing "worksheet" and the LSP number. Accept xlsx/xlsm/xls extensions.
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
    Update-BatchLink   # NYTT
    } finally {
        # valfritt
        Gui-Log '✅ Filer laddade'
    }
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

if (-not (Get-Command Write-SPSheet-Safe -ErrorAction SilentlyContinue)) {
    function Write-SPSheet-Safe {
        param(
            [OfficeOpenXml.ExcelPackage]$Pkg,
            [object]$Rows,                       # EN PSCustomObject eller en array
            [string[]]$DesiredOrder,             # används i tabell-läge
            [string]$Batch
        )
        if (-not $Pkg) { return $false }

        # Tvinga alltid array
        $Rows = @($Rows)

        # Skapa/ersätt flik
        $name = "SharePoint Info"
        $wsOld = $Pkg.Workbook.Worksheets[$name]
        if ($wsOld) { $Pkg.Workbook.Worksheets.Delete($wsOld) }
        $ws = $Pkg.Workbook.Worksheets.Add($name)

        # Tomt?
        if ($Rows.Count -eq 0 -or $Rows[0] -eq $null) {
            $ws.Cells[1,1].Value = "No rows found (Batch=$Batch)"
            # Flytta inte SharePoint-bladet till första position här. Ordningen sätts
            # senare i rapportlogiken för att säkerställa att Information blir först.
            return $true
        }

        # Är detta “Rubrik/Värde” (från batch-script)?
        $isKV = ($Rows[0].psobject.Properties.Name -contains 'Rubrik') -and `
                ($Rows[0].psobject.Properties.Name -contains 'Värde')

        if ($isKV) {
            # --- Nyckel/Värde-layout ---
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
            # --- Tabell-layout (kolumnrubriker som "Work Center", "Order#", ...) ---
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

        # Låt flikens ordning bestämmas i rapportlogiken istället för att tvinga första position här.
        return $true
    }
}

# ============================
# ===== RAPPORTLOGIK =========
# ============================

$btnBuild.Add_Click({
    Gui-Log 'Skapar rapport…' -Immediate
    try {
    # --- EPPlus obligatoriskt ---
    if (-not (Load-EPPlus)) { Gui-Log "❌ EPPlus kunde inte laddas – avbryter." 'Error'; return }

    # --- Valda filer (CSV valfri) ---
    $selCsv = Get-CheckedFilePath $clbCsv
    $selNeg = Get-CheckedFilePath $clbNeg
    $selPos = Get-CheckedFilePath $clbPos

    if (-not $selNeg -or -not $selPos) { Gui-Log "❌ Du måste välja en Seal NEG och en Seal POS." 'Error'; return }

    $lsp = ($txtLSP.Text + '').Trim()
    if (-not $lsp) { Gui-Log "⚠️ Ange ett LSP-nummer." 'Warn'; return }

    Gui-Log "📄 Neg-fil: $(Split-Path $selNeg -Leaf)" 'Info'
    Gui-Log "📄 Pos-fil: $(Split-Path $selPos -Leaf)" 'Info'
    if ($selCsv) { Gui-Log "📄 CSV: $(Split-Path $selCsv -Leaf)" 'Info' } else { Gui-Log "ℹ️ Ingen CSV vald." 'Info' }

    # --- Fil-lås om vi ska skriva signatur ---
    $negWritable = $true; $posWritable = $true
    if ($chkWriteSign.Checked) {
        $negWritable = -not (Test-FileLocked $selNeg); if (-not $negWritable) { Gui-Log "🔒 NEG är låst (öppen i Excel?)." 'Warn' }
        $posWritable = -not (Test-FileLocked $selPos); if (-not $posWritable) { Gui-Log "🔒 POS är låst (öppen i Excel?)." 'Warn' }
    }

    # --- Öppna NEG/POS + mall ---
    $pkgNeg = $null; $pkgPos = $null; $pkgOut = $null
    try {
        try {
            $pkgNeg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selNeg))
            $pkgPos = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selPos))
        } catch {
            Gui-Log "❌ Kunde inte öppna NEG/POS: $($_.Exception.Message)" 'Error'
            return
        }

        $templatePath = Join-Path $PSScriptRoot "output_template-v4.xlsx"
        if (-not (Test-Path -LiteralPath $templatePath)) { Gui-Log "❌ Mallfilen 'output_template-v4.xlsx' saknas!" 'Error'; return }
        try {
            $pkgOut = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($templatePath))
        } catch {
            Gui-Log "❌ Kunde inte läsa mall: $($_.Exception.Message)" 'Error'
            return
        }

        # ============================
        # === SIGNATUR I NEG/POS  ====
        # ============================
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
            } catch {
                Gui-Log "⚠️ Kunde inte spara signatur i NEG/POS: $($_.Exception.Message)" 'Warn'
            }
        }

        # ============================
        # === CSV (Info/Control)  ====
        # ============================
        $csvRows = @(); $runAssay = $null
        if ($selCsv) {
            try { $csvRows = Import-CsvRows -Path $selCsv -StartRow 10 } catch {}
            try { $runAssay = Get-AssayFromCsv -Path $selCsv -StartRow 10 } catch {}
            if ($runAssay) { Gui-Log "🔎 Assay från CSV: $runAssay" }
        }
        $controlTab = $null
        if ($runAssay) { $controlTab = Get-ControlTabName -AssayName $runAssay }
        if ($controlTab) { Gui-Log "🧪 Control Material-flik: $controlTab" } else { Gui-Log "ℹ️ Ingen control-mappning (fortsätter utan)." }

        # ============================
        # === Läs avvikelser       ===
        # ============================
        # Läs avvikelser från varje data‑blad i Seal Test NEG/POS. Detta läser in våtförluster
        # och detekterar fail eller minusvärde när Weight Loss (kolumn K) är ≤ -2.4 eller text i L
        # indikerar “FAIL”.
        $violationsNeg = @(); $violationsPos = @(); $failNegCount = 0; $failPosCount = 0

         foreach ($ws in $pkgNeg.Workbook.Worksheets) {
            if ($ws.Name -eq "Worksheet Instructions") { continue }
            if (-not $ws.Dimension) { continue }
            $obsC = Find-ObservationCol $ws
            for ($r = 3; $r -le 45; $r++) {
                $valK = $ws.Cells["K$r"].Value; $textL = $ws.Cells["L$r"].Text
                if ($valK -ne $null -and $valK -is [double]) {
                    if ($textL -eq "FAIL" -or $valK -le -3.0) {
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
         foreach ($ws in $pkgPos.Workbook.Worksheets) {
            if ($ws.Name -eq "Worksheet Instructions") { continue }
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

        # ============================
        # === Seal Test Info (blad) ==
        # ============================
        $wsOut1 = $pkgOut.Workbook.Worksheets["Seal Test Info"]
        if (-not $wsOut1) { Gui-Log "❌ Fliken 'Seal Test Info' saknas i mallen"; return }

        # Rensa mismatch-kolumn (D3..D15)
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
        $forceText = @("ROBAL","Part Number","Batch Number","Cartridge Number (LSP)","PO Number","Assay Family","Balance ID Number","Vacuum Oven ID Number","Timer ID Number")
        $mismatchFields = $fields[0..6] | ForEach-Object { $_.Label }

        $row = 3
        foreach ($f in $fields) {
            $valNeg=''; $valPos=''
             foreach ($wsN in $pkgNeg.Workbook.Worksheets) {
                if ($wsN.Name -eq "Worksheet Instructions") { continue }
                $cell = $wsN.Cells[$f.Cell]
                if ($cell.Value -ne $null) { if ($cell.Value -is [datetime]) { $valNeg = $cell.Value.ToString('MMM-yy') } else { $valNeg = $cell.Text }; break }
            }
             foreach ($wsP in $pkgPos.Workbook.Worksheets) {
                if ($wsP.Name -eq "Worksheet Instructions") { continue }
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

        # ============================
        # === Testare (B43)        ===
        # ============================
        $testersNeg = @(); $testersPos = @()
        foreach ($s in $pkgNeg.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) { $t=$s.Cells["B43"].Text; if ($t) { $testersNeg += ($t -split ",") } }
        foreach ($s in $pkgPos.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) { $t=$s.Cells["B43"].Text; if ($t) { $testersPos += ($t -split ",") } }
        $testersNeg = $testersNeg | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique
        $testersPos = $testersPos | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique

        $wsOut1.Cells["B16"].Value = "Name of Tester"
        $wsOut1.Cells["B16:C16"].Merge = $true
        $wsOut1.Cells["B16"].Style.HorizontalAlignment = "Center"

        $maxTesters = [Math]::Max($testersNeg.Count, $testersPos.Count)
        # Antal standardrader för testare i mallen. Höj till 11 för att rymma fler testare utan att rad  för signaturen hamnar mitt i listan.
        $initialRows = 11
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

        # ============================
        # === Signatur-jämförelse  ===
        # ============================
        $negSigSet = Get-SignatureSetForDataSheets -Pkg $pkgNeg
        $posSigSet = Get-SignatureSetForDataSheets -Pkg $pkgPos

        $negSet = New-Object 'System.Collections.Generic.HashSet[string]'
        $posSet = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($n in $negSigSet.NormSet) { [void]$negSet.Add($n) }
        foreach ($p in $posSigSet.NormSet) { [void]$posSet.Add($p) }

        $hasNeg = ($negSet.Count -gt 0)
        $hasPos = ($posSet.Count -gt 0)

        $onlyNeg = @(); $onlyPos = @(); $sigMismatch = $false
        if ($hasNeg -and $hasPos) {
            foreach ($n in $negSet) { if (-not $posSet.Contains($n)) { $onlyNeg += $n } }
            foreach ($p in $posSet) { if (-not $negSet.Contains($p)) { $onlyPos += $p } }
            $sigMismatch = ($onlyNeg.Count -gt 0 -or $onlyPos.Count -gt 0)
        } else {
            $sigMismatch = $false
        }

        $mismatchSheets = @()
        if ($sigMismatch) {
            foreach ($k in $onlyNeg) {
                $raw = if ($negSigSet.RawByNorm.ContainsKey($k)) { $negSigSet.RawByNorm[$k] } else { $k }
                $where = if ($negSigSet.Occ.ContainsKey($k)) { ($negSigSet.Occ[$k] -join ', ') } else { '—' }
                $mismatchSheets += ("NEG: " + $raw + "  [Blad: " + $where + "]")
            }
            foreach ($k in $onlyPos) {
                $raw = if ($posSigSet.RawByNorm.ContainsKey($k)) { $posSigSet.RawByNorm[$k] } else { $k }
                $where = if ($posSigSet.Occ.ContainsKey($k)) { ($posSigSet.Occ[$k] -join ', ') } else { '—' }
                $mismatchSheets += ("POS: " + $raw + "  [Blad: " + $where + "]")
            }
            Gui-Log "⚠️ Mismatch: Print Full Name, Sign, and Date (NEG vs POS)"
        }

        # Infoga signaturinfo (rad 32)
        function Set-MergedWrapAutoHeight {
            param([OfficeOpenXml.ExcelWorksheet]$Sheet,[int]$RowIndex,[int]$ColStart=2,[int]$ColEnd=3,[string]$Text)
            $rng = $Sheet.Cells[$RowIndex, $ColStart, $RowIndex, $ColEnd]
            $rng.Style.WrapText = $true
            $rng.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::None
            $Sheet.Row($RowIndex).CustomHeight = $false
            try {
                $wChars = [Math]::Floor(($Sheet.Column($ColStart).Width + $Sheet.Column($ColEnd).Width) - 2); if ($wChars -lt 1) { $wChars = 1 }
                $segments = $Text -split "(\r\n|\n|\r)"; $lineCount = 0
                foreach ($seg in $segments) { if (-not $seg) { $lineCount++ } else { $lineCount += [Math]::Ceiling($seg.Length / $wChars) } }
                if ($lineCount -lt 1) { $lineCount = 1 }
                $targetHeight = [Math]::Max(15, [Math]::Ceiling(15 * $lineCount * 2.15))
                if ($Sheet.Row($RowIndex).Height -lt $targetHeight) {
                    $Sheet.Row($RowIndex).Height = $targetHeight
                    $Sheet.Row($RowIndex).CustomHeight = $true
                }
            } catch { $Sheet.Row($RowIndex).CustomHeight = $false }
        }

        # Justera radnumret för signatursektionen dynamiskt.
        # I mallen finns en tom rad mellan den sista testaren och raden med texten
        # "Print Full Name, Sign, and Date", följt av raden "Neg/Pos" och därefter signaturraden.
        # Detta innebär att signraden hamnar tre rader efter listan med testare: 17 + antal testare + 3.
        $signRow = 17 + $maxTesters + 3
        $displaySignNeg = $null; $displaySignPos = $null
        if ($signToWrite) { $displaySignNeg = $signToWrite; $displaySignPos = $signToWrite }
        else {
            $displaySignNeg = if ($negSigSet.RawFirst) { $negSigSet.RawFirst } else { '—' }
            $displaySignPos = if ($posSigSet.RawFirst) { $posSigSet.RawFirst } else { '—' }
        }
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
        try { $wsOut1.Column(4).Width = 10 } catch {}

        if ($sigMismatch) {
            $mismatchCell = $wsOut1.Cells["D$signRow"]
            $mismatchCell.Value = 'Mismatch'
            Style-Cell $mismatchCell $true 'FF0000' 'Medium' 'FFFFFF'
            if ($mismatchSheets.Count -gt 0) {
                for ($j = 0; $j -lt $mismatchSheets.Count; $j++) {
                    $rowIdx = $signRow + 1 + $j
                    try { $wsOut1.Cells["B$rowIdx:C$rowIdx"].Merge = $true } catch {}
                    $text = $mismatchSheets[$j]
                    $wsOut1.Cells["B$rowIdx"].Value = $text
                    foreach ($mc in $wsOut1.Cells["B$rowIdx:C$rowIdx"]) { Style-Cell $mc $false 'CCFFFF' 'Medium' $null }
                    $wsOut1.Cells["B$rowIdx:C$rowIdx"].Style.HorizontalAlignment = 'Center'
                    if ($text -like 'NEG:*' -or $text -like 'POS:*') {
                        Set-MergedWrapAutoHeight -Sheet $wsOut1 -RowIndex $rowIdx -ColStart 2 -ColEnd 3 -Text $text
                    }
                }
            }
        }

        # ============================
        # === STF Sum               ===
        # ============================
        $wsOut2 = $pkgOut.Workbook.Worksheets["STF Sum"]
        if (-not $wsOut2) { Gui-Log "❌ Fliken 'STF Sum' saknas i mallen!"; return }

        $totalRows = $violationsNeg.Count + $violationsPos.Count
        $currentRow = 2
        if ($totalRows -eq 0) {
            Gui-Log "✅ Seal Test hittades"
            $wsOut2.Cells["B1:H1"].Value = $null
            $wsOut2.Cells["A1"].Value = "Inga STF hittades!"
            Style-Cell $wsOut2.Cells["A1"] $true "D9EAD3" "Medium" "006100"
            $wsOut2.Cells["A1"].Style.HorizontalAlignment = "Left"
            if ($wsOut2.Dimension -and $wsOut2.Dimension.End.Row -gt 1) { $wsOut2.DeleteRow(2, $wsOut2.Dimension.End.Row - 1) }
        } else {
            Gui-Log "❗ $failNegCount avvikelser i NEG, $failPosCount i POS"

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

# ============================
# === Information-blad     ===
# ============================
try {
    # -- Helpers (definieras en gång per session om de saknas) --
    if (-not (Get-Command Add-Hyperlink -ErrorAction SilentlyContinue)) {
        function Add-Hyperlink {
            param([OfficeOpenXml.ExcelRange]$Cell,[string]$Text,[string]$Url)
            try {
                $Cell.Value = $Text
                $Cell.Hyperlink = [Uri]$Url
                $Cell.Style.Font.UnderLine = $true
                $Cell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0,102,204))
            } catch {}
        }
    }
    if (-not (Get-Command Find-RegexCell -ErrorAction SilentlyContinue)) {
        function Find-RegexCell {
            param([OfficeOpenXml.ExcelWorksheet]$Ws,[regex]$Rx,[int]$MaxRows=200,[int]$MaxCols=40)
            if (-not $Ws -or -not $Ws.Dimension) { return $null }
            $rMax = [Math]::Min($Ws.Dimension.End.Row, $MaxRows)
            $cMax = [Math]::Min($Ws.Dimension.End.Column, $MaxCols)
            for ($r=1; $r -le $rMax; $r++) {
                for ($c=1; $c -le $cMax; $c++) {
            $t = Normalize-HeaderText ($Ws.Cells[$r,$c].Text + '')
            if ($t -and $Rx.IsMatch($t)) { return @{Row=$r;Col=$c;Text=$t} }
                }
            }
            return $null
        }
    }

    if (-not (Get-Command Get-SealHeaderDocInfo -ErrorAction SilentlyContinue)) {
        function Get-SealHeaderDocInfo {
            param([OfficeOpenXml.ExcelPackage]$Pkg)
            $result = [pscustomobject]@{ Raw=''; DocNo=''; Rev='' }
            if (-not $Pkg) { return $result }
            $ws = $Pkg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
            if (-not $ws) { return $result }
            try {
                $lt = ($ws.HeaderFooter.OddHeader.LeftAlignedText + '').Trim()
                if (-not $lt) { $lt = ($ws.HeaderFooter.EvenHeader.LeftAlignedText + '').Trim() }
                $result.Raw = $lt
                $rx = [regex]'(?i)(?:document\s*(?:no|nr|#|number)\s*[:#]?\s*([A-Z0-9\-_\.\/]+))?.*?(?:rev(?:ision)?\.?\s*[:#]?\s*([A-Z0-9\-_\.]+))?'
                $m = $rx.Match($lt)
                if ($m.Success) {
                    if ($m.Groups[1].Value) { $result.DocNo = $m.Groups[1].Value.Trim() }
                    if ($m.Groups[2].Value) { $result.Rev   = $m.Groups[2].Value.Trim() }
                }
            } catch {}
            return $result
        }
    }

    # -- Förbered fliken "Information" --
    $wsInfo = $pkgOut.Workbook.Worksheets['Information']
    if (-not $wsInfo) {
        $wsInfo = $pkgOut.Workbook.Worksheets.Add('Information')
    }
    # Säkerställ teckensnitt för konsekvent utseende
    try { $wsInfo.Cells.Style.Font.Name='Arial'; $wsInfo.Cells.Style.Font.Size=11 } catch {}

    
    # --- CSV-statistik och info till nya rader i Information (CSV-Info) ---
    try {
        # Läs CSV-statistik
        $csvStats = $null
        if ($selCsv -and (Test-Path -LiteralPath $selCsv)) {
            $csvStats = Get-CsvStats -Path $selCsv
        } else {
            $csvStats = [pscustomobject]@{
                TestCount    = 0
                DupCount     = 0
                Duplicates   = @()
                LspValues    = @()
                LspOK        = $null
                InstrumentByType = [ordered]@{}
            }
            }

# --- Bygg Infinity-serialslistan ---
$infSN = @()
if ($script:GXINF_Map) {
    foreach ($k in $script:GXINF_Map.Keys) {
        if ($k -like 'Infinity-*') {
            $infSN += ($script:GXINF_Map[$k].Split(',') | ForEach-Object { ($_ + '').Trim() } | Where-Object { $_ })
        }
    }
}
$infSN = $infSN | Select-Object -Unique

# --- Räkna SP för Infinity från CSV ---
$infSummary = '—'
try {
    if ($selCsv -and (Test-Path -LiteralPath $selCsv) -and $infSN.Count -gt 0) {
        $infSummary = Get-InfinitySpFromCsvStrict -Path $selCsv -InfinitySerials $infSN
    }
} catch {
    Gui-Log ("Infinity SP fel: " + $_.Exception.Message) 'Warn'
}

if ($selCsv -and (Test-Path -LiteralPath $selCsv)) {
    try { $csvStats = Get-CsvStats -Path $selCsv } catch {}
        }

        # Beräkna dubbletter för Sample ID (rad 11)
        $dupSampleCount = 0
        $dupSampleList  = @()
        if ($selCsv -and (Test-Path -LiteralPath $selCsv)) {
            try {
                $csvLines = Get-Content -LiteralPath $selCsv
                if ($csvLines.Count -gt 8) {
                    $headerFields = ConvertTo-CsvFields $csvLines[7]
                    $sampleIdx = -1
                    for ($i=0; $i -lt $headerFields.Count; $i++) {
                        $hf = ($headerFields[$i] + '').Trim().ToLower()
                        if ($hf -match 'sample') { $sampleIdx = $i; break }
                    }
                    if ($sampleIdx -ge 0) {
                        $samples = @()
                        for ($r=9; $r -lt $csvLines.Count; $r++) {
                            $line = $csvLines[$r]
                            if (-not $line -or -not $line.Trim()) { continue }
                            $fields = ConvertTo-CsvFields $line
                            if ($fields.Count -gt $sampleIdx) {
                                $val = ($fields[$sampleIdx] + '').Trim()
                                if ($val) { $samples += $val }
                            }
                        }
                        if ($samples.Count -gt 0) {
                            $counts = @{}
                            foreach ($s in $samples) { if (-not $counts.ContainsKey($s)) { $counts[$s] = 0 }; $counts[$s]++ }
                            $dupList = @()
                            foreach ($entry in $counts.GetEnumerator()) {
                                if ($entry.Value -gt 1) {
                                    # Kombinera värde och antal, t.ex. "12345 x2"
                                    $dupList += ("$($entry.Key) x$($entry.Value)")
                                }
                            }
                            $dupSampleCount = $dupList.Count
                            $dupSampleList  = $dupList
                        }
                    }
                }
            } catch {
                Gui-Log ("⚠️ Fel vid analys av Sample ID: " + $_.Exception.Message) 'Warn'
            }
        }
        $dupSampleText = if ($dupSampleCount -gt 0) {
            $show = ($dupSampleList | Select-Object -First 8) -join ', '
            "$dupSampleCount ($show)"
        } else { 'N/A' }

        # Beräkna dubbletter för Cartridge S/N (rad 12)
        $dupCartText = if ($csvStats.DupCount -gt 0) {
            $show = ($csvStats.Duplicates | Select-Object -First 8) -join ', '
            "$($csvStats.DupCount) ($show)"
        } else { 'N/A' }

        # ----- LSP från CSV (rad 9) -----
        # Sammanställ unika LSP-nummer från CSV och deras antal. Om inga hittas lämnas
        # $lspSummary tomt så att GUI-värdet används istället längre ned.
        $lspSummary = ''
        try {
            if ($selCsv -and (Test-Path -LiteralPath $selCsv)) {
                $csvLines2 = Get-Content -LiteralPath $selCsv
                if ($csvLines2.Count -gt 8) {
                    $counts = @{}
                    for ($rr = 9; $rr -lt $csvLines2.Count; $rr++) {
                        $ln = $csvLines2[$rr]
                        if (-not $ln -or -not $ln.Trim()) { continue }
                        $fs = ConvertTo-CsvFields $ln
                        if ($fs.Count -gt 4) {
                            $raw = ($fs[4] + '').Trim()
                            if ($raw) {
                                # Extrahera femsiffrigt nummer om möjligt, annars använd råtalet
                                $mLsp = [regex]::Match($raw,'(\\d{5})')
                                $code = if ($mLsp.Success) { $mLsp.Groups[1].Value } else { $raw }
                                if (-not $counts.ContainsKey($code)) { $counts[$code] = 0 }
                                $counts[$code]++
                            }
                        }
                    }
if ($counts.Count -gt 0) {
    $sorted = $counts.GetEnumerator() | Sort-Object Key

    $lspSummaryParts = @()
    foreach ($kvp in $sorted) {
        $part = if ($kvp.Value -gt 1) { "$($kvp.Key) x$($kvp.Value)" } else { $kvp.Key }
        $lspSummaryParts += $part
    }

    $total = $sorted.Count

    if ($total -eq 1) {
        $lspSummary = $sorted[0].Key
    }
    else {
        $lspSummary = "$total (" + ($lspSummaryParts -join ', ') + ")"
    }
}
                        }
                    }

        } catch {
            Gui-Log ("⚠️ Fel vid extraktion av LSP från CSV: " + $_.Exception.Message) 'Warn'
            $lspSummary = ''
        }
        
        $instText = if ($csvStats.InstrumentByType.Keys.Count -gt 0) {
            ($csvStats.InstrumentByType.GetEnumerator() | ForEach-Object { "$($_.Key)" } | Sort-Object) -join '; '
        } else { '' }

        function Find-InfoRow {
            param([OfficeOpenXml.ExcelWorksheet]$Ws, [string]$Label)
            if (-not $Ws -or -not $Ws.Dimension) { return $null }
            $maxRow = [Math]::Min($Ws.Dimension.End.Row, 300)
            for ($ri=1; $ri -le $maxRow; $ri++) {
                $txt = (($Ws.Cells[$ri,1].Text) + '').Trim()
                if (-not $txt) { continue }
                if ($txt.ToLowerInvariant() -eq $Label.ToLowerInvariant()) { return $ri }
            }
            return $null
        }

        $isNewLayout = $true
        try {
            $tmpRow = Find-InfoRow -Ws $wsInfo -Label 'CSV-Info'
            if ($tmpRow) { $isNewLayout = $true }
        } catch {}

        $rowCsvFile    = Find-InfoRow -Ws $wsInfo -Label 'CSV'
        $rowLsp        = Find-InfoRow -Ws $wsInfo -Label 'LSP'
        $rowAntal      = Find-InfoRow -Ws $wsInfo -Label 'Antal tester'
        # Sample ID kan förekomma endast i nya layouten
        $rowDupSample  = Find-InfoRow -Ws $wsInfo -Label 'Dubblett Sample ID'
        if (-not $rowDupSample) { $rowDupSample = Find-InfoRow -Ws $wsInfo -Label 'Dublett Sample ID' }
        # Cartridge S/N – namnet kan stavas med ett eller två b
        $rowDupCart    = Find-InfoRow -Ws $wsInfo -Label 'Dubblett Cartridge S/N'
        if (-not $rowDupCart) { $rowDupCart = Find-InfoRow -Ws $wsInfo -Label 'Dublett Cartridge S/N' }
        # Instrumentrad
        $rowInst       = Find-InfoRow -Ws $wsInfo -Label 'Använda INF/GX'

        $rowBag = Find-InfoRow -Ws $wsInfo -Label 'Bag Numbers Tested Using Infinity'
if (-not $rowBag) { $rowBag = Find-InfoRow -Ws $wsInfo -Label 'Bag Numbers Tested Using Infinity:' }
if (-not $rowBag) { $rowBag = 14 }  # fallback till din standardrad

# --- Skriv exakt en gång till B14 (eller B$rowBag) ---
$wsInfo.Cells["B$rowBag"].Style.Numberformat.Format = '@'
$wsInfo.Cells["B$rowBag"].Value = $infSummary

        if ($isNewLayout) {
            # Nya layoutens LSP-etikett är "LSP"
            $rowLsp = Find-InfoRow -Ws $wsInfo -Label 'LSP'
            # Standardrader om någon etikett saknas
            if (-not $rowCsvFile)   { $rowCsvFile   = 8 }
            if (-not $rowLsp)       { $rowLsp       = 9 }
            if (-not $rowAntal)     { $rowAntal     = 10 }
            if (-not $rowDupSample) { $rowDupSample = 11 }
            if (-not $rowDupCart)   { $rowDupCart   = 12 }
            if (-not $rowInst)      { $rowInst      = 13 }

        }

        # Sätt värden på respektive rad (kolumn B). Nummerformat för text
        if ($selCsv) {
            $wsInfo.Cells["B$rowCsvFile"].Style.Numberformat.Format = '@'
            $wsInfo.Cells["B$rowCsvFile"].Value = (Split-Path $selCsv -Leaf)
        } else {
            $wsInfo.Cells["B$rowCsvFile"].Value = ''
        }
        # LSP från CSV (om tillgänglig); annars använd GUI-värdet
        if ($lspSummary -and $lspSummary -ne '') {
            $wsInfo.Cells["B$rowLsp"].Style.Numberformat.Format = '@'
            $wsInfo.Cells["B$rowLsp"].Value = $lspSummary
        } else {
            $wsInfo.Cells["B$rowLsp"].Style.Numberformat.Format = '@'
            $wsInfo.Cells["B$rowLsp"].Value = $lsp
        }
        # Antal tester
        $wsInfo.Cells["B$rowAntal"].Value = $csvStats.TestCount
        $wsInfo.Cells["B$rowAntal"].Style.Numberformat.Format = '@'
        $wsInfo.Cells["B$rowAntal"].Value = "$($csvStats.TestCount)"
        # Dubblett Sample ID (skriv endast om rad finns i mallen)
        if ($rowDupSample) {
            $wsInfo.Cells["B$rowDupSample"].Value = $dupSampleText
        }
        # Dubblett Cartridge S/N
        if ($rowDupCart) {
            $wsInfo.Cells["B$rowDupCart"].Value = $dupCartText
        }
        # Använda INF/GX
        $wsInfo.Cells["B$rowInst"].Value = $instText
        # Bag Numbers Tested Using Infinity
    } catch {
        Gui-Log ("⚠️ CSV data-fel: " + $_.Exception.Message) 'Warn'
    }

# --- Meta & källor ---
    $csvLeaf = ''
    if ($selCsv) { $csvLeaf = Split-Path $selCsv -Leaf }
    $negLeaf = Split-Path $selNeg -Leaf
    $posLeaf = Split-Path $selPos -Leaf

    $assayForMacro = ''
    if ($runAssay) {
        $assayForMacro = $runAssay
    } elseif ($wsOut1) {
        $assayForMacro = ($wsOut1.Cells['D10'].Text + '').Trim()
    }
    $miniVal = ''
    if (Get-Command Get-MinitabMacro -ErrorAction SilentlyContinue) {
        $miniVal = Get-MinitabMacro -AssayName $assayForMacro
    }
    if (-not $miniVal) { $miniVal = 'N/A' }

    $hdNeg = $null; $hdPos = $null
    try { $hdNeg = Get-SealHeaderDocInfo -Pkg $pkgNeg } catch {}
    try { $hdPos = Get-SealHeaderDocInfo -Pkg $pkgPos } catch {}
    if (-not $hdNeg) { $hdNeg = [pscustomobject]@{Raw='';DocNo='';Rev=''} }
    if (-not $hdPos) { $hdPos = [pscustomobject]@{Raw='';DocNo='';Rev=''} }

    $wsInfo.Cells['B2'].Value = $ScriptVersion
    $wsInfo.Cells['B3'].Value = $env:USERNAME
    $wsInfo.Cells['B4'].Value = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    $wsInfo.Cells['B5'].Value = if ($miniVal) { $miniVal } else { 'N/A' }

    $selLsp = $null
    try {
        if (Get-Variable -Name clbLsp -ErrorAction SilentlyContinue) {
            $selLsp = Get-CheckedFilePath $clbLsp
        }
    } catch {}

    $batch = $null
    # Få batch från POS → NEG fallback
    try { if ($selPos) { $batch = Get-BatchFromSealFile $selPos } } catch {}
    if (-not $batch) {
        try { if ($selNeg) { $batch = Get-BatchFromSealFile $selNeg } } catch {}
    }
    $batchEsc = ''
    if ($batch) { $batchEsc = [uri]::EscapeDataString($batch) }
    $lspEsc = ''
    if ($lsp) { $lspEsc = [uri]::EscapeDataString($lsp) }
    $url = ''
    if ($SharePointBatchLinkTemplate) {
        $url = $SharePointBatchLinkTemplate -replace '\{BatchNumber\}', $batchEsc
        $url = $url -replace '\{LSP\}', $lspEsc
    } else {
        $url = "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management/Lists/Cepheid%20%20Production%20orders/AllItems.aspx?view=7&q=$batchEsc"
    }
    $linkText = 'Ingen batch funnen'
    if ($batch) { $linkText = ("Öppna " + $batch) }

    # --- Klickbar batch-länk (POS -> NEG fallback) ---
    function Get-BatchFromSealFile {
        param([string]$Path)
        if (-not (Test-Path -LiteralPath $Path)) { return $null }
        try {
            $p = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($Path))
            $ws  = $p.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
            $bn  = $null
            if ($ws) { $bn = ($ws.Cells['D2'].Text + '').Trim() }
            $p.Dispose()
            return $bn
        } catch { return $null }
    }
    $batch = $null
    if ($selPos) { $batch = Get-BatchFromSealFile $selPos }
    if (-not $batch -and $selNeg) { $batch = Get-BatchFromSealFile $selNeg }

    # Skriv SharePoint-batchlänk i en fast rad (rad 34)
    $batchEsc = ''
    if ($batch) { $batchEsc = [uri]::EscapeDataString($batch) }
    $lspEsc = ''
    if ($lsp) { $lspEsc = [uri]::EscapeDataString($lsp) }
    $url = ''
    if ($SharePointBatchLinkTemplate) {
        $url = $SharePointBatchLinkTemplate -replace '\{BatchNumber\}', $batchEsc
        $url = $url -replace '\{LSP\}', $lspEsc
    } else {
        $url = "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management/Lists/Cepheid%20%20Production%20orders/AllItems.aspx?view=7&q=$batchEsc"
    }
    $linkText = 'Ingen batch funnen'
    if ($batch) { $linkText = ("Öppna " + $batch) }
    $wsInfo.Cells['A34'].Value = 'SharePoint Batch'
    $wsInfo.Cells['A34'].Style.Font.Bold = $true
    Add-Hyperlink -Cell $wsInfo.Cells['B34'] -Text $linkText -Url $url


    $linkMap = [ordered]@{
        'IPT App'      = 'https://apps.powerapps.com/play/e/default-771c9c47-7f24-44dc-958e-34f8713a8394/a/fd340dbd-bbbf-470b-b043-d2af4cb62c83'
        'MES Login'    = 'http://mes.cepheid.pri/camstarportal/?domain=CEPHEID.COM'
        'CSV Uploader' = 'http://auw2wgxtpap01.cepaws.com/Welcome.aspx'
        'BMRAM'        = 'https://cepheid62468.coolbluecloud.com/'
        'Agile'        = 'https://agileprod.cepheid.com/Agile/default/login-cms.jsp'
    }
    $rowLink = 35
    foreach ($key in $linkMap.Keys) {
        $wsInfo.Cells["A$rowLink"].Value = $key
        # Förkorta texten som visas i cellen till "LÄNK" enligt mallens stil
        Add-Hyperlink -Cell $wsInfo.Cells["B$rowLink"] -Text 'LÄNK' -Url $linkMap[$key]
        $rowLink++
    }

    # ----------------------------------------------------------------
    # WS (LSP Worksheet): hitta fil och skriv in i Information-bladet
    # ----------------------------------------------------------------
    try {
        # 1) Om GUI har clbLsp – använd den. Annars försök hitta *worksheet*.xls* bredvid POS/NEG.
        if (-not $selLsp) {
            $probeDir = $null
            if ($selPos) { $probeDir = Split-Path -Parent $selPos }
            if (-not $probeDir -and $selNeg) { $probeDir = Split-Path -Parent $selNeg }
            if ($probeDir -and (Test-Path -LiteralPath $probeDir)) {
                $cand = Get-ChildItem -LiteralPath $probeDir -File -ErrorAction SilentlyContinue |
                        Where-Object {
                            # Filnamnet ska innehålla ordet "worksheet" (case-insensitive) och LSP-numret
                            ($_.Name -match '(?i)worksheet') -and ($_.Name -match [regex]::Escape($lsp)) -and ($_.Extension -match '^\.(xlsx|xlsm|xls)$')
                        } |
                        Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($cand) {
                    # Tilldela vald LSP-worksheet till variabeln men injicera inte i bladet här.
                    $selLsp = $cand.FullName
                }
            }
        }

        function Find-LabelValueRightward {
        $normLbl = Normalize-HeaderText $Label
        # tillåt valfritt antal mellanslag, valfri punkt/kolon i slutet
        $pat = '^(?i)\s*' + [regex]::Escape($normLbl).Replace('\ ', '\s*') + '\s*[:\.]*\s*$'
        $rx  = [regex]::new($pat, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $hit = Find-RegexCell -Ws $Ws -Rx $rx -MaxRows $MaxRows -MaxCols $MaxCols
    if (-not $hit) { return $null }
    $cMax = [Math]::Min($Ws.Dimension.End.Column, $MaxCols)
    for ($c = $hit.Col + 1; $c -le $cMax; $c++) {
        $t = Normalize-HeaderText ($Ws.Cells[$hit.Row,$c].Text + '')
        if ($t) { return $t }
    }
    return $null
}

        if ($selLsp -and (Test-Path -LiteralPath $selLsp)) {
            Gui-Log ("🔎 WS hittad: " + (Split-Path $selLsp -Leaf)) 'Info'
            # LSP-filen finns. Vi extraherar endast headern via Extract-WorksheetHeader längre ned.
            # Tidigare logik för att extrahera och infoga Test Summary och SPC har tagits bort för att förenkla.
            # Eventuell headerinformation från denna WS används i sammanfattningen nedan.
        } else {
            Gui-Log "ℹ️ Ingen WS-fil vald/hittad (LSP Worksheet). Hoppar över WS-extraktion." 'Info'
        }
    } catch {
        Gui-Log ("⚠️ WS-block fel: " + $_.Exception.Message) 'Warn'
    }

        # --- Header summary table (Worksheet & Seal Test) ---
        try {
            $headerWs  = $null
            $headerNeg = $null
            $headerPos = $null
            # Hämta header från WS (om tillgänglig)
            if ($selLsp -and (Test-Path -LiteralPath $selLsp)) {
                try {
                    $tmpPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selLsp))
                    $headerWs = Extract-WorksheetHeader -Pkg $tmpPkg


            # === NYTT: Konsistenskontroll av LSP-header över ALLA flikar ===
            $wsHeaderRows  = Get-WorksheetHeaderPerSheet -Pkg $tmpPkg
            $wsHeaderCheck = Compare-WorksheetHeaderSet   -Rows $wsHeaderRows

            # Logga resultat av header-kontrollen. Eventuell utskrift sker senare på specifika rader.
            try {
                if ($wsHeaderCheck.Issues -gt 0 -and $wsHeaderCheck.Summary) {
                    Gui-Log ("Worksheet header-avvikelser: {0} – se Information!" -f $wsHeaderCheck.Summary) 'Warn'
                } else {
                    Gui-Log "Worksheet header konsekvent över alla flikar." 'Info'
                }
            } catch {}
                $tmpPkg.Dispose()
                } catch {}
            }
            # Hämta header från NEG/POS
            try { $headerNeg = Extract-SealTestHeader -Pkg $pkgNeg } catch {}
            try { $headerPos = Extract-SealTestHeader -Pkg $pkgPos } catch {}

            # Om vissa fält i headern inte hittades via header/footer, försök att hitta dem i själva
            # arbetsbladet genom att söka efter etiketter och hämta värdet i cellen till höger.
            # Detta gäller för Part No., Batch No(s), Cartridge No. (LSP) och Effective.
            try {
                if ($selLsp -and (Test-Path -LiteralPath $selLsp)) {
                    $tmpPkg2 = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selLsp))
                    $wsLsp   = $tmpPkg2.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                    if ($wsLsp) {
                        # Försök hitta Part No. i cellerna om den saknas. Prova flera varianter inklusive kolon
                        if (-not $headerWs -or -not $headerWs.PartNo) {
                            $val = $null
                            $labels = @(
                                'Part No.', 'Part No.:', 'Part No', 'Part Number', 'Part Number:', 'Part Number.', 'Part Number.:'
                            )
                            foreach ($lbl in $labels) {
                                $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl
                                if ($val) { break }
                            }
                            if ($val) { $headerWs.PartNo = $val }
                        }
                        # Försök hitta Batch No(s) i cellerna. Tillåt olika stavningar och kolon
                        if (-not $headerWs -or -not $headerWs.BatchNo) {
                            $val = $null
                            $labels = @(
                                'Batch No(s)', 'Batch No(s).', 'Batch No(s):', 'Batch No(s).:',
                                'Batch No', 'Batch No.', 'Batch No:', 'Batch No.:' ,
                                'Batch Number', 'Batch Number.', 'Batch Number:', 'Batch Number.:'
                            )
                            foreach ($lbl in $labels) {
                                $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl
                                if ($val) { break }
                            }
                            if ($val) { $headerWs.BatchNo = $val }
                        }
                        # Försök hitta Cartridge No (LSP) i cellerna. Tillåt olika stavningar, parentes (LSP) och kolon
                        # Om CartridgeNo saknas eller är en ensam punkt (från mall) så försöker vi hitta den
                        if (-not $headerWs -or -not $headerWs.CartridgeNo -or $headerWs.CartridgeNo -eq '.') {
                            $val = $null
                            # Först, prova exakta etiketter med olika kolon/blankstegsvariationer
                            $labels = @(
                                'Cartridge No. (LSP)', 'Cartridge No. (LSP):', 'Cartridge No. (LSP) :',
                                'Cartridge No (LSP)', 'Cartridge No (LSP):', 'Cartridge No (LSP) :',
                                'Cartridge Number (LSP)', 'Cartridge Number (LSP):', 'Cartridge Number (LSP) :',
                                'Cartridge No.', 'Cartridge No.:', 'Cartridge No. :', 'Cartridge No :',
                                'Cartridge Number', 'Cartridge Number:', 'Cartridge Number :',
                                'Cartridge No', 'Cartridge No:', 'Cartridge No :'
                            )
                            foreach ($lbl in $labels) {
                                $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl
                                if ($val) { break }
                            }
                            # Om fortfarande null, försök hitta en cell som innehåller "Cartridge" och "LSP"
                            if (-not $val) {
                                $rxCart = [regex]::new('(?i)Cartridge.*\(LSP\)')
                                # Sök igenom fler kolumner (upp till 100) för att hitta cellen med "Cartridge" och "(LSP)"
                                $maxCols = [Math]::Min($wsLsp.Dimension.End.Column, 100)
                                $hitCart = Find-RegexCell -Ws $wsLsp -Rx $rxCart -MaxRows 200 -MaxCols $maxCols
                                if ($hitCart) {
                                    # Leta efter första icke-tomma cellen till höger om etiketten
                                    for ($c = $hitCart.Col + 1; $c -le $wsLsp.Dimension.End.Column; $c++) {
                                        $cellVal = ($wsLsp.Cells[$hitCart.Row, $c].Text + '').Trim()
                                        if ($cellVal) { $val = $cellVal; break }
                                    }
                                }
                            }
                            if ($val) { $headerWs.CartridgeNo = $val }
                        }
                        if (-not $headerWs -or -not $headerWs.Effective) {
                            $val = Find-LabelValueRightward -Ws $wsLsp -Label 'Effective'
                            if (-not $val) { $val = Find-LabelValueRightward -Ws $wsLsp -Label 'Effective Date' }
                            if ($val) { $headerWs.Effective = $val }
                        }
            }
            # Om Cartridge No fortfarande är tomt, försök hitta första femsiffriga tal i filnamnet som fallback.
            try {
                # Om CartridgeNo fortfarande saknas, är en punkt eller tom sträng, försök hämta nummer ur filnamnet.
                if ($selLsp -and (-not $headerWs -or -not $headerWs.CartridgeNo -or $headerWs.CartridgeNo -eq '.' -or $headerWs.CartridgeNo -eq '')) {
                    $fn = Split-Path $selLsp -Leaf
                    # Hitta första tal om 5–7 siffror som står självständigt (ej del av längre tal)
                    $m = [regex]::Matches($fn, '(?<!\d)(\d{5,7})(?!\d)')
                    if ($m.Count -gt 0) {
                        $headerWs.CartridgeNo = $m[0].Groups[1].Value
                    }
                }
            } catch {}
                    $tmpPkg2.Dispose()
                }
            } catch {}

            # Fyll i Effective för POS/NEG om saknas via cellbaserad sökning
            try {
                if ($selPos -and -not $headerPos.Effective) {
                    $tmpPkg3 = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selPos))
                    $wsPos   = $tmpPkg3.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                    if ($wsPos) {
                        $val = Find-LabelValueRightward -Ws $wsPos -Label 'Effective'
                        if (-not $val) { $val = Find-LabelValueRightward -Ws $wsPos -Label 'Effective Date' }
                        if ($val) { $headerPos.Effective = $val }
                    }
                    $tmpPkg3.Dispose()
                }
            } catch {}
            try {
                if ($selNeg -and -not $headerNeg.Effective) {
                    $tmpPkg4 = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selNeg))
                    $wsNeg   = $tmpPkg4.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                    if ($wsNeg) {
                        $val = Find-LabelValueRightward -Ws $wsNeg -Label 'Effective'
                        if (-not $val) { $val = Find-LabelValueRightward -Ws $wsNeg -Label 'Effective Date' }
                        if ($val) { $headerNeg.Effective = $val }
                    }
                    $tmpPkg4.Dispose()
                }
            } catch {}
            # Fastställ batcher
            $wsBatch   = if ($headerWs -and $headerWs.BatchNo) { $headerWs.BatchNo } else { $null }
            $sealBatch = $batch
            if (-not $sealBatch) {
                try { $sealBatch = Get-BatchFromSealFile $selPos } catch {}
                if (-not $sealBatch) { try { $sealBatch = Get-BatchFromSealFile $selNeg } catch {} }
            }
            $batchMatchFlag = $null
            if ($wsBatch -and $sealBatch) { $batchMatchFlag = ($wsBatch -eq $sealBatch) }
            # Fastställ konsistens mellan POS/NEG
            $sealConsistentFlag = $null
            if ($headerNeg -and $headerPos) {
                if ($headerNeg.DocumentNumber -and $headerPos.DocumentNumber -and $headerNeg.Rev -and $headerPos.Rev -and $headerNeg.Effective -and $headerPos.Effective) {
                    $sealConsistentFlag = (($headerNeg.DocumentNumber -eq $headerPos.DocumentNumber) -and ($headerNeg.Rev -eq $headerPos.Rev) -and ($headerNeg.Effective -eq $headerPos.Effective))
                }
            }
            # Sätt noteringar för oväntade Document Number
            $noteStr = ''
            if ($headerNeg -and $headerNeg.DocumentNumber -and $headerNeg.DocumentNumber -ne 'D10552') { $noteStr += ("NEG DocNo (" + $headerNeg.DocumentNumber + ") != D10552; ") }
            if ($headerPos -and $headerPos.DocumentNumber -and $headerPos.DocumentNumber -ne 'D10552') { $noteStr += ("POS DocNo (" + $headerPos.DocumentNumber + ") != D10552; ") }
            # Dynamiskt hitta rader för header-information (Worksheet, POS, NEG) genom att söka efter etiketter i kolumn A.
            # Find-InfoRow-funktionen definieras tidigare i samma event och återanvänds här.

            # Worksheet‑sektion
            $rowWsFile = Find-InfoRow -Ws $wsInfo -Label 'Worksheet'
            if (-not $rowWsFile) { $rowWsFile = 17 }
            $rowPart  = $rowWsFile + 1
            $rowBatch = $rowWsFile + 2
            $rowCart  = $rowWsFile + 3
            $rowDoc   = $rowWsFile + 4
            $rowRev   = $rowWsFile + 5
            $rowEff   = $rowWsFile + 6

            # Seal Test POS‑sektion
            $rowPosFile = Find-InfoRow -Ws $wsInfo -Label 'Seal Test POS'
            if (-not $rowPosFile) {
                # Om ej hittad, anta fyra rader efter Worksheet‑sektionen eller använd tidigare fallback (24)
                if ($rowWsFile) { $rowPosFile = $rowWsFile + 7 } else { $rowPosFile = 24 }
            }
            $rowPosDoc = $rowPosFile + 1
            $rowPosRev = $rowPosFile + 2
            $rowPosEff = $rowPosFile + 3

            # Seal Test NEG‑sektion
            $rowNegFile = Find-InfoRow -Ws $wsInfo -Label 'Seal Test NEG'
            if (-not $rowNegFile) {
                # Om ej hittad, anta fyra rader efter POS‑sektionen
                $rowNegFile = $rowPosFile + 4
            }
            $rowNegDoc = $rowNegFile + 1
            $rowNegRev = $rowNegFile + 2
            $rowNegEff = $rowNegFile + 3

            # Worksheet-filnamn
            if ($selLsp) {
                $wsInfo.Cells["B$rowWsFile"].Style.Numberformat.Format = '@'
                $wsInfo.Cells["B$rowWsFile"].Value = (Split-Path $selLsp -Leaf)
            } else {
                $wsInfo.Cells["B$rowWsFile"].Value = ''
            }
 
            # Konsensus mellan WS/POS/NEG (ändrar inte källvärden)
            $consPart  = Get-ConsensusValue -Type 'Part'      -Ws $headerWs.PartNo      -Pos $headerPos.PartNumber   -Neg $headerNeg.PartNumber
            $consBatch = Get-ConsensusValue -Type 'Batch'     -Ws $headerWs.BatchNo     -Pos $headerPos.BatchNumber  -Neg $headerNeg.BatchNumber
            $consCart  = Get-ConsensusValue -Type 'Cartridge' -Ws $headerWs.CartridgeNo -Pos $headerPos.CartridgeNo  -Neg $headerNeg.CartridgeNo

            # Extra fallback för Cartridge: om konsensus saknas → försök från LSP-filnamn (5–7 siffror)
            if (-not $consCart.Value -and $selLsp) {
                $fnCart = Split-Path $selLsp -Leaf
                $mCart  = [regex]::Match($fnCart,'(?<!\d)(\d{5,7})(?!\d)')
                if ($mCart.Success) {
                    $consCart = @{
                        Value  = $mCart.Groups[1].Value
                        Source = 'FILENAME'
                        Note   = 'Filename fallback'
                    }
                }
            }

            # Skriv ut konsensusvärden enligt ny layout
            if ($consPart.Value)  { $wsInfo.Cells["B$rowPart"].Value = $consPart.Value }  else { $wsInfo.Cells["B$rowPart"].Value = '' }
            if ($consBatch.Value) { $wsInfo.Cells["B$rowBatch"].Value = $consBatch.Value } else { $wsInfo.Cells["B$rowBatch"].Value = '' }
            if ($consCart.Value)  { $wsInfo.Cells["B$rowCart"].Value = $consCart.Value }  else { $wsInfo.Cells["B$rowCart"].Value = '' }

            # Bestäm om batch-mismatch ska visas som notering
            $batchMismatch = $false
            try {
                if ($headerNeg -and $headerPos -and $headerNeg.BatchNumber -and $headerPos.BatchNumber) {
                    $normNeg = Normalize-Id -Value $headerNeg.BatchNumber -Type 'Batch'
                    $normPos = Normalize-Id -Value $headerPos.BatchNumber -Type 'Batch'
                    if ($normNeg -and $normPos -and $normNeg -ne $normPos) { $batchMismatch = $true }
                }
            } catch {}

            # Lägg till kommentarer endast om POS/NEG batch skiljer sig
            if ($batchMismatch) {
                try { if ($consPart.Note)  { [void]$wsInfo.Cells["B$rowPart"].AddComment($consPart.Note,  'DocMerge') } } catch {}
                try { if ($consBatch.Note) { [void]$wsInfo.Cells["B$rowBatch"].AddComment($consBatch.Note, 'DocMerge') } } catch {}
                try { if ($consCart.Note)  { [void]$wsInfo.Cells["B$rowCart"].AddComment($consCart.Note,  'DocMerge') } } catch {}
            }

            # Fyll i header-avvikelser per fält i kolumn C (PartNo, BatchNo, CartridgeNo)
            # Använd informationen från wsHeaderCheck.Details som returneras av Compare-WorksheetHeaderSet.
            try {
                if ($wsHeaderCheck -and $wsHeaderCheck.Details) {
                    $linesDev = ($wsHeaderCheck.Details -split "`r?`n")
                    $devPart  = $null
                    $devBatch = $null
                    $devCart  = $null
                    foreach ($ln in $linesDev) {
                        if ($ln -match '^-\s*PartNo[^:]*:\s*(.+)$') {
                            $devPart = $matches[1].Trim()
                        } elseif ($ln -match '^-\s*BatchNo[^:]*:\s*(.+)$') {
                            $devBatch = $matches[1].Trim()
                        } elseif ($ln -match '^-\s*CartridgeNo[^:]*:\s*(.+)$') {
                            $devCart = $matches[1].Trim()
                        }
                    }
                    if ($devPart) {
                        $wsInfo.Cells["C$rowPart"].Style.Numberformat.Format = '@'
                        $wsInfo.Cells["C$rowPart"].Value = 'Avvikande flik: ' + $devPart
                    }
                    if ($devBatch) {
                        $wsInfo.Cells["C$rowBatch"].Style.Numberformat.Format = '@'
                        $wsInfo.Cells["C$rowBatch"].Value = 'Avvikande flik: ' + $devBatch
                    }
                    if ($devCart) {
                        $wsInfo.Cells["C$rowCart"].Style.Numberformat.Format = '@'
                        $wsInfo.Cells["C$rowCart"].Value = 'Avvikande flik: ' + $devCart
                    }
                }
            } catch {}


            # Dokumentnummer (inkl. attachment), Rev och Effective för Worksheet
            if ($headerWs) {
                $doc = $headerWs.DocumentNumber
                if ($doc) {
                    # klipp av svans som ibland limmas ihop i headern
                    $doc = ($doc -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$', '').Trim()
                }
                if ($headerWs.Attachment -and ($doc -notmatch '(?i)\bAttachment\s+\w+\b')) {
                    $doc = "$doc Attachment $($headerWs.Attachment)"
                }
                $wsInfo.Cells["B$rowDoc"].Value = $doc
                $wsInfo.Cells["B$rowRev"].Value = $headerWs.Rev
                $wsInfo.Cells["B$rowEff"].Value = $headerWs.Effective
            } else {
                $wsInfo.Cells["B$rowDoc"].Value = ''
                $wsInfo.Cells["B$rowRev"].Value = ''
                $wsInfo.Cells["B$rowEff"].Value = ''
            }
            # Seal Test POS fil
            if ($selPos) {
                $wsInfo.Cells["B$rowPosFile"].Style.Numberformat.Format = '@'
                $wsInfo.Cells["B$rowPosFile"].Value = (Split-Path $selPos -Leaf)
            } else {
                $wsInfo.Cells["B$rowPosFile"].Value = ''
            }
            # Seal Test POS metadata
            if ($headerPos) {
                # POS: ta bort ev. "Rev/Effective" som följt med
                $docPos = $headerPos.DocumentNumber
                if ($docPos) { $docPos = ($docPos -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$','').Trim() }
                $wsInfo.Cells["B$rowPosDoc"].Value = $docPos
                $wsInfo.Cells["B$rowPosRev"].Value = $headerPos.Rev
                $wsInfo.Cells["B$rowPosEff"].Value = $headerPos.Effective
            } else {
                $wsInfo.Cells["B$rowPosDoc"].Value = ''
                $wsInfo.Cells["B$rowPosRev"].Value = ''
                $wsInfo.Cells["B$rowPosEff"].Value = ''
            }
            # Seal Test NEG fil
            if ($selNeg) {
                $wsInfo.Cells["B$rowNegFile"].Style.Numberformat.Format = '@'
                $wsInfo.Cells["B$rowNegFile"].Value = (Split-Path $selNeg -Leaf)
            } else {
                $wsInfo.Cells["B$rowNegFile"].Value = ''
            }
            # Seal Test NEG metadata
            if ($headerNeg) {
                # NEG: ta bort ev. "Rev/Effective" som följt med
                $docNeg = $headerNeg.DocumentNumber
                if ($docNeg) { $docNeg = ($docNeg -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$','').Trim() }
                $wsInfo.Cells["B$rowNegDoc"].Value = $docNeg
                $wsInfo.Cells["B$rowNegRev"].Value = $headerNeg.Rev
                $wsInfo.Cells["B$rowNegEff"].Value = $headerNeg.Effective
            } else {
                $wsInfo.Cells["B$rowNegDoc"].Value = ''
                $wsInfo.Cells["B$rowNegRev"].Value = ''
                $wsInfo.Cells["B$rowNegEff"].Value = ''
            }
            # Töm eventuella överflödiga rader nedanför tabellen – ej nödvändig då layout definierad i mall
        } catch {
            Gui-Log ("⚠️ Header summary fel: " + $_.Exception.Message) 'Warn'
        }

} catch {
    Gui-Log "⚠️ Information-blad fel: $($_.Exception.Message)" 'Warn'
}

        # ============================
        # === Equipment-blad       ===
        # ============================
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
            } else { Gui-Log "ℹ️ Utrustningslista saknas: $UtrustningListPath" 'Info' }
        } catch { Gui-Log "⚠️ Kunde inte kopiera 'Equipment': $($_.Exception.Message)" 'Warn' }

        # ============================
        # === Control Material      ===
        # ============================
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
                    Gui-Log "✅ Control Material kopierad: '$($srcWs.Name)' → '$destName'" 'Info'
                } else { Gui-Log "ℹ️ Hittade inget blad i kontrollfilen som matchar '$controlTab'." 'Info' }
                $srcPkg.Dispose()
            } else { Gui-Log "ℹ️ Ingen Control-flik skapad (saknar mappning eller kontrollfil)." 'Info' }
        } catch { Gui-Log "⚠️ Control Material-fel: $($_.Exception.Message)" 'Warn' }

        # ============================
        # === SharePoint Info       ===
        # ============================
        try {
            # 0) checkbox
            if ($chkSharePointInfo -and -not $chkSharePointInfo.Checked) {
                Gui-Log "ℹ️ SharePoint Info ej valt – hoppar över." 'Info'
                # Rensa ev. gammal flik när ej valt:
                try { $old = $pkgOut.Workbook.Worksheets["SharePoint Info"]; if ($old) { $pkgOut.Workbook.Worksheets.Delete($old) } } catch {}
            } else {
                # 1) Säkerställ anslutning
                $spOk = $false
                if ($global:SpConnected) { $spOk = $true }
                elseif (Get-Command Get-PnPConnection -ErrorAction SilentlyContinue) {
                    try { $null = Get-PnPConnection; $spOk = $true } catch { $spOk = $false }
                }

                if (-not $spOk) {
                    $errMsg = if ($global:SpError) { $global:SpError } else { 'Okänt fel' }
                    Gui-Log ("⚠️ SharePoint ej tillgängligt: $errMsg") 'Warn'
                }

                # 2) Hämta Batch # från POS→NEG (D2)
                function Get-BatchFromSealFile {
                    param([string]$Path)
                    if (-not (Test-Path -LiteralPath $Path)) { return $null }
                    try {
                        $p = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($Path))
                        $ws  = $p.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                        $bn  = if ($ws) { ($ws.Cells['D2'].Text + '').Trim() } else { $null }
                        $p.Dispose()
                        return $bn
                    } catch { return $null }
                }
                $batch = $null
                try { $batch = Get-BatchFromSealFile $selPos } catch {}
                if (-not $batch) { try { $batch = Get-BatchFromSealFile $selNeg } catch {} }

                if (-not $batch) {
                    Gui-Log "ℹ️ Inget Batch # i POS/NEG – skriver tom SharePoint Info." 'Info'
                    if (Get-Command Write-SPSheet-Safe -ErrorAction SilentlyContinue) {
                        [void](Write-SPSheet-Safe -Pkg $pkgOut -Rows @() -DesiredOrder @() -Batch '—')
                    } else {
                        # enkel fallback
                        $wsSp = $pkgOut.Workbook.Worksheets["SharePoint Info"]; if ($wsSp) { $pkgOut.Workbook.Worksheets.Delete($wsSp) }
                        $wsSp = $pkgOut.Workbook.Worksheets.Add("SharePoint Info")
                        $wsSp.Cells[1,1].Value = "Rubrik"; $wsSp.Cells[1,2].Value = "Värde"
                        $wsSp.Cells[2,1].Value = "Batch";  $wsSp.Cells[2,2].Value = "—"
                        try { $wsSp.Cells[$wsSp.Dimension.Address].AutoFitColumns() | Out-Null } catch {}
                    }
                } else {
                    Gui-Log "🔎 Batch hittad: $batch" 'Info'

                    # 3) Fält / rubrik / ordning
                    $fields = @(
                        'Work_x0020_Center','Title','Batch_x0023_','SAP_x0020_Batch_x0023__x0020_2',
                        'LSP','Material','BBD_x002f_SLED','Actual_x0020_startdate_x002f__x0',
                        'PAL_x0020__x002d__x0020_Sample_x','Sample_x0020_Reagent_x0020_P_x00',
                        'Order_x0020_quantity','Total_x0020_good','ITP_x0020_Test_x0020_results',
                        'IPT_x0020__x002d__x0020_Testing_0','MES_x0020__x002d__x0020_Order_x0'
                    )
                    $renameMap = @{
                        'Work Center'            = 'Work Center'
                        'Title'                  = 'Order#'
                        'Batch#'                 = 'SAP Batch#'
                        'SAP Batch# 2'           = 'SAP Batch# 2'
                        'LSP'                    = 'LSP'
                        'Material'               = 'Material'
                        'BBD/SLED'               = 'BBD/SLED'
                        'Actual startdate/_x0'   = 'ROBAL - Actual start date/time'
                        'PAL - Sample_x'         = 'Sample Reagent use'
                        'Sample Reagent P'       = 'Sample Reagent P/N'
                        'Order quantity'         = 'Order quantity'
                        'Total good'             = 'ROBAL - Till Packning'
                        'IPT Test results'       = 'IPT Test results'
                        'IPT - Testing_0'        = 'IPT - Testing Finalized'
                        'MES - Order_x0'         = 'MES Order'
                    }
                    $desiredOrder = @(
                        'Work Center','Order#','SAP Batch#','SAP Batch# 2','LSP','Material','BBD/SLED',
                        'ROBAL - Actual start date/time','Sample Reagent use','Sample Reagent P/N',
                        'Order quantity','ROBAL - Till Packning','IPT Test results',
                        'IPT - Testing Finalized','MES Order'
                    )
                    $dateFields      = @('BBD/SLED','ROBAL - Actual start date/time','IPT - Testing Finalized')
                    $shortDateFields = @('BBD/SLED')  # endast datum

                    $rows = @()
                    if ($spOk) {
                        try {
                            $items = Get-PnPListItem -List "Cepheid | Production orders" -Fields $fields -PageSize 2000 -ErrorAction Stop
                            $match = $items | Where-Object {
                                $v1 = $_['Batch_x0023_']; $v2 = $_['SAP_x0020_Batch_x0023__x0020_2']
                                $s1 = if ($null -ne $v1) { ([string]$v1).Trim() } else { '' }
                                $s2 = if ($null -ne $v2) { ([string]$v2).Trim() } else { '' }
                                $s1 -eq $batch -or $s2 -eq $batch
                            } | Select-Object -First 1

                            if ($match) {
                                foreach ($f in $fields) {
                                    $val = $match[$f]

                                    # label norm
                                    $label = $f -replace '_x0020_', ' ' `
                                                 -replace '_x002d_', '-' `
                                                 -replace '_x0023_', '#' `
                                                 -replace '_x002f_', '/' `
                                                 -replace '_x2013_', '–' `
                                                 -replace '_x00',''
                                    $label = $label.Trim()
                                    if ($renameMap.ContainsKey($label)) { $label = $renameMap[$label] }

                                    if ($null -ne $val -and $val -ne '') {
                                        if ($val -eq $true) { $val = 'JA' }
                                        elseif ($val -eq $false) { $val = 'NEJ' }

                                        # datumformat
                                        $dt = $null
                                        if ($val -is [datetime]) { $dt = [datetime]$val }
                                        else { try { $dt = [datetime]::Parse($val) } catch { $dt = $null } }

                                        if ($dt -ne $null -and ($dateFields -contains $label)) {
                                            $fmt = if ($shortDateFields -contains $label) { 'yyyy-MM-dd' } else { 'yyyy-MM-dd HH:mm' }
                                            $val = $dt.ToString($fmt)
                                        }

                                        $rows += [pscustomobject]@{ Rubrik = $label; 'Värde' = $val }
                                    }
                                }

                                # ordning
                                if ($rows.Count -gt 0) {
                                    $ordered = @()
                                    foreach ($label in $desiredOrder) {
                                        $hit = $rows | Where-Object { $_.Rubrik -eq $label } | Select-Object -First 1
                                        if ($hit) { $ordered += $hit }
                                    }
                                    if ($ordered.Count -gt 0) { $rows = $ordered }
                                }
                                Gui-Log "📄 SharePoint-post hittad – skriver blad." 'Info'
                            } else {
                                Gui-Log "ℹ️ Ingen post i SharePoint för Batch=$batch." 'Info'
                            }
                        } catch {
                            Gui-Log "⚠️ SP: Get-PnPListItem misslyckades: $($_.Exception.Message)" 'Warn'
                        }
                    }

                    # Skriv blad (även tomt)
                    if (Get-Command Write-SPSheet-Safe -ErrorAction SilentlyContinue) {
                        [void](Write-SPSheet-Safe -Pkg $pkgOut -Rows $rows -DesiredOrder $desiredOrder -Batch $batch)
                    } else {
                        # enkel fallback
                        $wsSp = $pkgOut.Workbook.Worksheets["SharePoint Info"]; if ($wsSp) { $pkgOut.Workbook.Worksheets.Delete($wsSp) }
                        $wsSp = $pkgOut.Workbook.Worksheets.Add("SharePoint Info")
                        $wsSp.Cells[1,1].Value = "Rubrik"; $wsSp.Cells[1,2].Value = "Värde"
                        if ($rows.Count -gt 0) {
                            $r=2; foreach($rowObj in $rows) { $wsSp.Cells[$r,1].Value = $rowObj.Rubrik; $wsSp.Cells[$r,2].Value = $rowObj.'Värde'; $r++ }
                        } else {
                            $wsSp.Cells[2,1].Value = "Batch";  $wsSp.Cells[2,2].Value = $batch
                            $wsSp.Cells[3,1].Value = "Info";   $wsSp.Cells[3,2].Value = "No matching SharePoint row"
                        }
                        try { $wsSp.Cells[$wsSp.Dimension.Address].AutoFitColumns() | Out-Null } catch {}
                    }

                    try {
                        if ($slBatchLink -and $batch) {
                            $slBatchLink.Text = "SharePoint: $batch"
                            $slBatchLink.Tag  = "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management/Lists/Cepheid%20%20Production%20orders/ROBAL.aspx?viewid=6c9e53c9-a377-40c1-a154-13a13866b52b&view=7&q=$batch"
                            $slBatchLink.Enabled = $true
                        }
                    } catch {}

                    # Wrap "Sample Reagent use"
                    try {
                        $wsSP = $pkgOut.Workbook.Worksheets['SharePoint Info']
                        if ($wsSP -and $wsSP.Dimension) {
                            $labelCol = 1; $valueCol = 2
                            for ($r = 1; $r -le $wsSP.Dimension.End.Row; $r++) {
                                if (($wsSP.Cells[$r,$labelCol].Text).Trim() -eq 'Sample Reagent use') {
                                    $wsSP.Cells[$r,$valueCol].Style.WrapText = $true
                                    $wsSP.Cells[$r,$valueCol].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
                                    try { $wsSP.Column($valueCol).Width = 55 } catch {}
                                    $wsSP.Row($r).CustomHeight = $true
                                    break
                                }
                            }
                        }
                    } catch {
                        Gui-Log "⚠️ WrapText på 'Sample Reagent use' misslyckades: $($_.Exception.Message)" 'Warn'
                    }
                }
            }
        } catch {
            Gui-Log "⚠️ SP-blad: $($_.Exception.Message)" 'Warn'
        }

        # ============================
        # === Header watermark     ===
        # ============================
        try {
            foreach ($ws in $pkgOut.Workbook.Worksheets) {
                try {
                    $ws.HeaderFooter.OddHeader.CenteredText   = '&"Arial,Bold"&14 UNCONTROLLED'
                    $ws.HeaderFooter.EvenHeader.CenteredText  = '&"Arial,Bold"&14 UNCONTROLLED'
                    $ws.HeaderFooter.FirstHeader.CenteredText = '&"Arial,Bold"&14 UNCONTROLLED'
                } catch { Write-Warning "Kunde inte sätta header på blad: $($ws.Name)" }
            }
        } catch { Write-Warning "Fel vid vattenstämpling av rapporten." }

        # ============================
        # === Tab-färger (innan Save)
        # ============================
        try {
            $wsT = $pkgOut.Workbook.Worksheets['Information'];     if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 52, 152, 219) }
            $wsT = $pkgOut.Workbook.Worksheets['Equipment'];       if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 33, 115, 70) }
            $wsT = $pkgOut.Workbook.Worksheets['SharePoint Info']; if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 0, 120, 212) }
        } catch {
            Gui-Log "⚠️ Kunde inte sätta tab-färg: $($_.Exception.Message)" 'Warn'
        }

        # ============================
        # === Spara & Audit        ===
        # ============================
        $nowTs   = Get-Date -Format "yyyyMMdd_HHmmss"
        $baseName = "$($env:USERNAME)_output_${lsp}_$nowTs.xlsx"
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
            $wsInitial = $pkgOut.Workbook.Worksheets["Information"]
            if ($wsInitial) { $wsInitial.View.TabSelected = $true }
            # Innan vi sparar, säkerställ korrekt tabbordning med Information först.
            # Vi avstår från att flytta flikarna vid sparning, se kommentar ovan.
            # Spara arbetsboken efter att flikordningen har säkerställts
            $pkgOut.SaveAs($SavePath)
            Gui-Log "✅ Rapport sparad: $SavePath" 'Info'
            $global:LastReportPath = $SavePath

            # Audit
            try {
                $auditDir = Join-Path $PSScriptRoot 'audit'
                if (-not (Test-Path $auditDir)) { New-Item -ItemType Directory -Path $auditDir -Force | Out-Null }
                $auditObj = [pscustomobject]@{
                    DatumTid        = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    Användare       = $env:USERNAME
                    LSP             = $lsp
                    ValdCSV         = if ($selCsv) { Split-Path $selCsv -Leaf } else { '' }
                    ValdSealNEG     = Split-Path $selNeg -Leaf
                    ValdSealPOS     = Split-Path $selPos -Leaf
                    SignaturSkriven = if ($chkWriteSign.Checked) { 'Ja' } else { 'Nej' }
                    OverwroteSign   = if ($chkOverwriteSign.Checked) { 'Ja' } else { 'Nej' }
                    SigMismatch     = if ($sigMismatch) { 'Ja' } else { 'Nej' }
                    MismatchSheets  = if ($mismatchSheets -and $mismatchSheets.Count -gt 0) { ($mismatchSheets -join ';') } else { '' }
                    ViolationsNEG   = $violationsNeg.Count
                    ViolationsPOS   = $violationsPos.Count
                    Violations      = ($violationsNeg.Count + $violationsPos.Count)
                    Sparläge        = if ($rbSaveInLsp.Checked) { 'LSP-mapp' } else { 'Temporärt' }
                    OutputFile      = $SavePath
                    Kommentar       = 'UNCONTROLLED rapport, ingen källfil ändrades automatiskt.'
                    ScriptVersion   = $ScriptVersion

                }
                $auditFile = Join-Path $auditDir ("$($env:USERNAME)_audit_${nowTs}.csv")
                $auditObj | Export-Csv -Path $auditFile -NoTypeInformation -Encoding UTF8
            } catch { Gui-Log "⚠️ Kunde inte skriva revisionsfil: $($_.Exception.Message)" 'Warn' }

            # Öppna rapport i Excel
            try { Start-Process -FilePath "excel.exe" -ArgumentList "`"$SavePath`"" } catch {}
        }
        catch { Gui-Log "⚠️ Kunde inte spara/öppna: $($_.Exception.Message)" 'Warn' }

    } finally {
        try { if ($pkgNeg) { $pkgNeg.Dispose() } } catch {}
        try { if ($pkgPos) { $pkgPos.Dispose() } } catch {}
        try { if ($pkgOut) { $pkgOut.Dispose() } } catch {}
    }
    } finally {
        # valfritt
        # Gui-Log 'Klar.'
    }
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
$toolTip.SetToolTip($chkSharePointInfo, 'Exportera med SharePoint Info.')

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
