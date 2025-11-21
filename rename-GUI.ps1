[CmdletBinding()]
param(
    [string]$Lsp,
    [string]$HeadFolder,
    [string]$NewDatePrefix
)

# =====================[ KONFIG ]=====================

$RootFolders = @(
    '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Tests',
    '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\3. IPT - KLART FÖR SAMMANSTÄLLNING',
    '\\SE.CEPHEID.PRI\Cepheid Sweden\QC\QC-1\IPT\4. IPT - KLART FÖR GRANSKNING'
) |
Where-Object { $_ -and (Test-Path -LiteralPath $_) } |
Select-Object -Unique

if (-not $RootFolders -or $RootFolders.Count -eq 0) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "Inga giltiga rotmappar hittades i konfigurationen.",
        "LSP-datum-prefix-rename"
    ) | Out-Null
    return
}

# =====================[ WinForms-init ]=====================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:LogTextBox        = $null
$script:CurrentHeadFolder = $null
$script:CurrentItems      = @()
$script:LspHits           = @()

function Add-GuiLog {
    param([string]$Text)

    if ($script:LogTextBox -ne $null -and -not $script:LogTextBox.IsDisposed) {
        if ($script:LogTextBox.TextLength -gt 0) {
            $script:LogTextBox.AppendText("`r`n$Text")
        } else {
            $script:LogTextBox.AppendText($Text)
        }
        $script:LogTextBox.SelectionStart = $script:LogTextBox.TextLength
        $script:LogTextBox.ScrollToCaret()
    }
}

# =====================[ Hjälpfunktioner ]=====================

# Tar bort alla inledande datumblock (inkl. 2025 10 09A) och normaliserar första
# datumet till "yyyy MM dd". Rest är namnet efter datumet.
function Parse-DatePrefixName {
    param(
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }

    $rx = '^\s*(?<Block>(?:\d{4}[-\s]?\d{2}[-\s]?\d{2}\s*)+)(?<Rest>.*)$'
    $m  = [regex]::Match($Name, $rx)
    if (-not $m.Success) { return $null }

    $block = $m.Groups['Block'].Value
    $rest  = $m.Groups['Rest'].Value

    $digits = ($block -replace '\D', '')
    if ($digits.Length -lt 8) { return $null }

    $digits = $digits.Substring(0,8)
    $year   = $digits.Substring(0,4)
    $month  = $digits.Substring(4,2)
    $day    = $digits.Substring(6,2)

    $currentDate = ('{0} {1} {2}' -f $year, $month, $day)

    $rest = $rest.TrimStart(" -_.".ToCharArray())

    if ([string]::IsNullOrWhiteSpace($rest)) {
        $oldName = $currentDate
    } else {
        $oldName = "$currentDate $rest"
    }

    [pscustomobject]@{
        Date    = $currentDate
        Rest    = $rest
        OldName = $oldName
    }
}

function Get-RelativePath {
    param(
        [string]$BaseFolder,
        [string]$FullPath
    )

    if ([string]::IsNullOrWhiteSpace($BaseFolder) -or
        [string]::IsNullOrWhiteSpace($FullPath)) {
        return $FullPath
    }

    $baseResolved = (Resolve-Path -LiteralPath $BaseFolder).ProviderPath
    $baseResolved = $baseResolved.TrimEnd('\')
    $full = $FullPath

    if ($full.StartsWith($baseResolved, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $full.Substring($baseResolved.Length).TrimStart('\')
    }

    return $FullPath
}

function Truncate-Text {
    param(
        [string]$Text,
        [int]$MaxLength = 80
    )

    if (-not $Text) { return "" }
    if ($Text.Length -le $MaxLength) { return $Text }
    $Text.Substring(0, $MaxLength - 3) + "..."
}

function Scan-DatePrefixItems {
    param(
        [string]$BaseFolder,
        [string]$TargetDatePrefix
    )

    $results = @()

    if ([string]::IsNullOrWhiteSpace($BaseFolder)) { return $results }
    if (-not (Test-Path -LiteralPath $BaseFolder -PathType Container)) { return $results }

    $baseResolved = (Resolve-Path -LiteralPath $BaseFolder).ProviderPath

    $scanCount = 0

    # --- Mappar ---
    $dirs = @()
    try {
        $dirs = [System.IO.Directory]::EnumerateDirectories(
            $baseResolved,
            '*',
            [System.IO.SearchOption]::AllDirectories
        )
    }
    catch { $dirs = @() }

    foreach ($dir in $dirs) {
        $scanCount++
        if (($scanCount % 200) -eq 0) { [System.Windows.Forms.Application]::DoEvents() }

        $name   = [System.IO.Path]::GetFileName($dir)
        $parsed = Parse-DatePrefixName -Name $name
        if ($parsed -eq $null) { continue }

        if ([string]::IsNullOrWhiteSpace($parsed.Rest)) {
            $newName = $TargetDatePrefix
        }
        else {
            $newName = "$TargetDatePrefix $($parsed.Rest)"
        }

        $relPath = Get-RelativePath -BaseFolder $BaseFolder -FullPath $dir

        $results += [pscustomobject]@{
            Type        = 'Mapp'
            FullName    = $dir
            RelPath     = $relPath
            CurrentDate = $parsed.Date
            NewDate     = $TargetDatePrefix
            OldName     = $parsed.OldName
            NewName     = $newName
            IsDir       = $true
        }
    }

    # --- Filer ---
    $files = @()
    try {
        $files = [System.IO.Directory]::EnumerateFiles(
            $baseResolved,
            '*',
            [System.IO.SearchOption]::AllDirectories
        )
    }
    catch { $files = @() }

    foreach ($file in $files) {
        $scanCount++
        if (($scanCount % 200) -eq 0) { [System.Windows.Forms.Application]::DoEvents() }

        $name   = [System.IO.Path]::GetFileName($file)
        $parsed = Parse-DatePrefixName -Name $name
        if ($parsed -eq $null) { continue }

        if ([string]::IsNullOrWhiteSpace($parsed.Rest)) {
            $newName = $TargetDatePrefix
        }
        else {
            $newName = "$TargetDatePrefix $($parsed.Rest)"
        }

        $relPath = Get-RelativePath -BaseFolder $BaseFolder -FullPath $file

        $results += [pscustomobject]@{
            Type        = 'Fil'
            FullName    = $file
            RelPath     = $relPath
            CurrentDate = $parsed.Date
            NewDate     = $TargetDatePrefix
            OldName     = $parsed.OldName
            NewName     = $newName
            IsDir       = $false
        }
    }

    $results
}

function Test-FileLocked {
    param(
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }

    try {
        $fs = New-Object System.IO.FileStream(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
        $fs.Close()
        $false
    }
    catch {
        $true
    }
}

function Apply-Rename {
    param(
        [string]$BaseFolder,
        [object[]]$Items,
        [string]$TargetDatePrefix
    )

    $DirOK = 0
    $DirErr = 0
    $FileOK = 0
    $FileErr = 0
    $LockedFiles = @()

    # --- 1) Mappar, djupaste först ---
    $dirItems = @($Items | Where-Object { $_.IsDir -eq $true })
    if ($dirItems.Count -gt 0) {
        $dirItems = $dirItems | Sort-Object {
            if ([string]::IsNullOrWhiteSpace($_.RelPath)) {
                0
            }
            else {
                ($_.RelPath -split '\\').Length
            }
        } -Descending

        $i = 0
        foreach ($item in $dirItems) {
            $i++
            if (($i % 100) -eq 0) { [System.Windows.Forms.Application]::DoEvents() }

            $full = $item.FullName
            if (-not (Test-Path -LiteralPath $full -PathType Container)) {
                $DirErr++
                continue
            }

            $parent     = [System.IO.Path]::GetDirectoryName($full)
            $targetLeaf = $item.NewName
            $target     = Join-Path -Path $parent -ChildPath $targetLeaf

            if ($full.TrimEnd('\') -ieq $target.TrimEnd('\')) { continue }

            if (Test-Path -LiteralPath $target) {
                $altLeaf   = ($item.NewName + " (1)")
                $altTarget = Join-Path -Path $parent -ChildPath $altLeaf

                if (Test-Path -LiteralPath $altTarget) {
                    Add-GuiLog ("Varning: kunde inte byta mapp (mål finns redan två gånger): {0}" -f $full)
                    $DirErr++
                    continue
                }
                else {
                    $targetLeaf = $altLeaf
                    $target     = $altTarget
                }
            }

            try {
                Rename-Item -LiteralPath $full -NewName $targetLeaf -ErrorAction Stop
                $DirOK++
            }
            catch {
                Add-GuiLog ("Varning: kunde inte byta mapp: {0}. Fel: {1}" -f $full, $_.Exception.Message)
                $DirErr++
            }
        }
    }

    # --- 2) Skanna om filer efter mapp-bytet ---
    $fileItems = Scan-DatePrefixItems -BaseFolder $BaseFolder -TargetDatePrefix $TargetDatePrefix |
        Where-Object { $_.IsDir -eq $false }

    # --- 3) Filer ---
    $j = 0
    foreach ($item in $fileItems) {
        $j++
        if (($j % 200) -eq 0) { [System.Windows.Forms.Application]::DoEvents() }

        $full = $item.FullName
        if (-not (Test-Path -LiteralPath $full -PathType Leaf)) {
            $FileErr++
            continue
        }

        if (Test-FileLocked -Path $full) {
            Add-GuiLog ("🔒 Låst fil (hoppar över): {0}" -f $full)
            $FileErr++
            $LockedFiles += $full
            continue
        }

        $parent     = [System.IO.Path]::GetDirectoryName($full)
        $targetLeaf = $item.NewName
        $target     = Join-Path -Path $parent -ChildPath $targetLeaf

        if ($full -ieq $target) { continue }

        if (Test-Path -LiteralPath $target) {
            $nameOnly = [System.IO.Path]::GetFileNameWithoutExtension($item.NewName)
            $ext      = [System.IO.Path]::GetExtension($item.NewName)

            if ($ext) {
                $altLeaf = ("{0} (1){1}" -f $nameOnly, $ext)
            }
            else {
                $altLeaf = ($nameOnly + " (1)")
            }

            $altTarget = Join-Path -Path $parent -ChildPath $altLeaf
            if (Test-Path -LiteralPath $altTarget) {
                Add-GuiLog ("Varning: kunde inte byta fil (mål finns redan två gånger): {0}" -f $full)
                $FileErr++
                continue
            }
            else {
                $targetLeaf = $altLeaf
                $target     = $altTarget
            }
        }

        try {
            Rename-Item -LiteralPath $full -NewName $targetLeaf -ErrorAction Stop
            $FileOK++
        }
        catch {
            Add-GuiLog ("Varning: kunde inte byta fil: {0}. Fel: {1}" -f $full, $_.Exception.Message)
            $FileErr++
        }
    }

    [pscustomobject]@{
        DirOK       = $DirOK
        DirErr      = $DirErr
        FileOK      = $FileOK
        FileErr     = $FileErr
        LockedFiles = $LockedFiles
    }
}

function Find-LspHeadFolders {
    param(
        [string]$LspNumber
    )

    $hits = @()
    if ([string]::IsNullOrWhiteSpace($LspNumber)) { return $hits }

    $pattern = ('(?i)#{0}(?!\d)' -f [regex]::Escape($LspNumber))

    foreach ($root in $RootFolders) {
        try {
            $children = Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue
        }
        catch { $children = @() }

        foreach ($child in $children) {
            if ($child.Name -match $pattern) {
                $full = $child.FullName
                if ($full -match '^[A-Za-z]:\\?$' -or $full -match '^[A-Za-z]:?$') {
                    continue
                }

                $hits += [pscustomobject]@{
                    Name     = $child.Name
                    FullPath = $full
                }
            }
        }
    }

    $hits
}

# =====================[ GUI-layout ]=====================

$form = New-Object System.Windows.Forms.Form
$form.Text = "LSP-datum-prefix-rename"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(1200, 720)
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)
$form.BackColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Övre panel (LSP + datum)
$panelTop = New-Object System.Windows.Forms.TableLayoutPanel
$panelTop.Dock = [System.Windows.Forms.DockStyle]::Top
$panelTop.RowCount = 2
$panelTop.ColumnCount = 1
$panelTop.AutoSize = $true
$panelTop.Padding = New-Object System.Windows.Forms.Padding(10,10,10,5)
$form.Controls.Add($panelTop)

# LSP-grupp
$grpLsp = New-Object System.Windows.Forms.GroupBox
$grpLsp.Text = " LSP "
$grpLsp.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpLsp.Height = 90
$panelTop.Controls.Add($grpLsp, 0, 0)

$lblLsp = New-Object System.Windows.Forms.Label
$lblLsp.Text = "LSP-nummer:"
$lblLsp.AutoSize = $true
$lblLsp.Location = New-Object System.Drawing.Point(10, 25)
$grpLsp.Controls.Add($lblLsp)

$txtLsp = New-Object System.Windows.Forms.TextBox
$txtLsp.Location = New-Object System.Drawing.Point(100, 22)
$txtLsp.Width = 120
if ($Lsp) { $txtLsp.Text = $Lsp }
$grpLsp.Controls.Add($txtLsp)

$btnFindLsp = New-Object System.Windows.Forms.Button
$btnFindLsp.Text = "Sök LSP-mappar"
$btnFindLsp.Location = New-Object System.Drawing.Point(240, 20)
$btnFindLsp.Width = 140
$grpLsp.Controls.Add($btnFindLsp)

$lblLspHits = New-Object System.Windows.Forms.Label
$lblLspHits.Text = "LSP-mappar:"
$lblLspHits.AutoSize = $true
$lblLspHits.Location = New-Object System.Drawing.Point(10, 55)
$grpLsp.Controls.Add($lblLspHits)

$listLspHits = New-Object System.Windows.Forms.ListBox
$listLspHits.Location = New-Object System.Drawing.Point(100, 50)
$listLspHits.Width = 500
$listLspHits.Height = 35
$listLspHits.Anchor = "Top,Left,Right"
$grpLsp.Controls.Add($listLspHits)

# Datum-grupp
$grpDate = New-Object System.Windows.Forms.GroupBox
$grpDate.Text = " Datum-prefix "
$grpDate.Dock = [System.Windows.Forms.DockStyle]::Fill
$grpDate.Height = 60
$panelTop.Controls.Add($grpDate, 0, 1)

$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Text = "Nytt datum-prefix (yyyy MM dd):"
$lblDate.AutoSize = $true
$lblDate.Location = New-Object System.Drawing.Point(10, 25)
$grpDate.Controls.Add($lblDate)

$txtDate = New-Object System.Windows.Forms.TextBox
$txtDate.Location = New-Object System.Drawing.Point(210, 22)
$txtDate.Width = 110
if ($NewDatePrefix) { $txtDate.Text = $NewDatePrefix } else { $txtDate.Text = (Get-Date -Format 'yyyy MM dd') }
$grpDate.Controls.Add($txtDate)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Skanna datum-prefix"
$btnScan.Location = New-Object System.Drawing.Point(335, 20)
$btnScan.Width = 160
$grpDate.Controls.Add($btnScan)

# Mittpanel med preview (2 kolumner: gammalt namn / nytt namn)
$panelMid = New-Object System.Windows.Forms.Panel
$panelMid.Dock = [System.Windows.Forms.DockStyle]::Fill
$panelMid.Padding = New-Object System.Windows.Forms.Padding(10,0,10,0)
$form.Controls.Add($panelMid)

$lvPreview = New-Object System.Windows.Forms.ListView
$lvPreview.View             = [System.Windows.Forms.View]::Details
$lvPreview.FullRowSelect    = $true
$lvPreview.GridLines        = $true
$lvPreview.HideSelection    = $false
$lvPreview.ShowItemToolTips = $true
$lvPreview.Font             = New-Object System.Drawing.Font("Consolas", 9)
$lvPreview.Dock             = [System.Windows.Forms.DockStyle]::Fill
$lvPreview.Scrollable       = $true
$lvPreview.MultiSelect      = $false
$lvPreview.HeaderStyle      = 'Nonclickable'

$colOld = $lvPreview.Columns.Add("Nuvarande namn", 550)
$colNew = $lvPreview.Columns.Add("Nytt namn", 550)

$panelMid.Controls.Add($lvPreview)

$lvPreview.Add_SizeChanged({
    $totalWidth = $lvPreview.ClientSize.Width
    if ($totalWidth -gt 0) {
        $colWidth = [int]([math]::Floor($totalWidth / 2))
        $colOld.Width = $colWidth
        $colNew.Width = $colWidth
    }
})

# Nedre panel – summary/logg/knappar
$panelBottom = New-Object System.Windows.Forms.Panel
$panelBottom.Dock = [System.Windows.Forms.DockStyle]::Bottom
$panelBottom.Height = 120
$panelBottom.Padding = New-Object System.Windows.Forms.Padding(10,5,10,10)
$form.Controls.Add($panelBottom)

$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Text = "Ingen skanning utförd ännu."
$lblSummary.AutoSize = $true
$lblSummary.Location = New-Object System.Drawing.Point(10, 5)
$panelBottom.Controls.Add($lblSummary)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Logg:"
$lblLog.AutoSize = $true
$lblLog.Location = New-Object System.Drawing.Point(10, 30)
$panelBottom.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(60, 28)
$txtLog.Width = 880
$txtLog.Height = 60
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.Anchor = "Top,Left,Right"
$panelBottom.Controls.Add($txtLog)
$script:LogTextBox = $txtLog

$btnRename = New-Object System.Windows.Forms.Button
$btnRename.Text = "Byt datum-prefix"
$btnRename.Width = 150
$btnRename.Location = New-Object System.Drawing.Point(960, 25)
$btnRename.Anchor = "Top,Right"
$panelBottom.Controls.Add($btnRename)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Stäng"
$btnClose.Width = 150
$btnClose.Location = New-Object System.Drawing.Point(960, 65)
$btnClose.Anchor = "Top,Right"
$btnClose.Add_Click({ $form.Close() })
$panelBottom.Controls.Add($btnClose)

# =====================[ Enter-genvägar ]=====================

$txtLsp.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        $btnFindLsp.PerformClick()
    }
})

$txtDate.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        $btnScan.PerformClick()
    }
})

# =====================[ Händelser ]=====================

$btnFindLsp.Add_Click({
    $txtLog.Clear()
    $lvPreview.Items.Clear()
    $lblSummary.Text = "Ingen skanning utförd ännu."
    $script:CurrentItems = @()
    $script:CurrentHeadFolder = $null
    $listLspHits.Items.Clear()
    $script:LspHits = @()

    $lspVal = $txtLsp.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($lspVal)) {
        [System.Windows.Forms.MessageBox]::Show("Ange LSP-nummer först.","LSP") | Out-Null
        return
    }

    Add-GuiLog ("Söker LSP-mappar för #{0}..." -f $lspVal)
    $hits = Find-LspHeadFolders -LspNumber $lspVal

    if (-not $hits -or $hits.Count -eq 0) {
        Add-GuiLog ("Inga LSP-huvudmappar (#{0}) hittades i rotmapparna." -f $lspVal)
        [System.Windows.Forms.MessageBox]::Show(
            "Inga LSP-huvudmappar (#$lspVal) hittades.",
            "LSP"
        ) | Out-Null
        return
    }

    $script:LspHits = $hits
    foreach ($h in $hits) {
        [void]$listLspHits.Items.Add($h.Name)
    }
    $listLspHits.SelectedIndex = 0

    Add-GuiLog ("Hittade {0} LSP-mapp(er)." -f $hits.Count)
})

$btnScan.Add_Click({
    $txtLog.Clear()
    $lvPreview.Items.Clear()
    $script:CurrentItems = @()

    $lspVal = $txtLsp.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($lspVal)) {
        [System.Windows.Forms.MessageBox]::Show("Ange LSP-nummer först.","LSP") | Out-Null
        return
    }

    if ($listLspHits.Items.Count -eq 0) {
        $btnFindLsp.PerformClick()
        if ($listLspHits.Items.Count -eq 0) { return }
    }

    if ($listLspHits.SelectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show("Välj en LSP-mapp i listan.","LSP") | Out-Null
        return
    }

    $head = $script:LspHits[$listLspHits.SelectedIndex].FullPath
    $script:CurrentHeadFolder = $head

    $prefix = $txtDate.Text.Trim()
    if ($prefix -notmatch '^\d{4}\s\d{2}\s\d{2}$') {
        [System.Windows.Forms.MessageBox]::Show(
            "Ogiltigt datum-prefix. Använd formatet 'yyyy MM dd'.",
            "Datum-prefix"
        ) | Out-Null
        return
    }

    Add-GuiLog ("Skannar '{0}' för datum-prefix..." -f (Split-Path -Path $head -Leaf))
    [System.Windows.Forms.Application]::DoEvents()

    $items = Scan-DatePrefixItems -BaseFolder $head -TargetDatePrefix $prefix
    $script:CurrentItems = $items

    if (-not $items -or $items.Count -eq 0) {
        $lblSummary.Text = "Inga filer/mappar med giltigt datum-prefix hittades."
        Add-GuiLog "Inga filer/mappar med giltigt datum-prefix hittades. Ingen åtgärd."
        return
    }

    $dirCount   = (@($items | Where-Object { $_.IsDir -eq $true })).Count
    $fileCount  = (@($items | Where-Object { $_.IsDir -eq $false })).Count
    $totalCount = $items.Count

    $lblSummary.Text = ("Hittade {0} mappar och {1} filer med datum-prefix (totalt {2} poster)." -f $dirCount, $fileCount, $totalCount)

    $lvPreview.BeginUpdate()
    foreach ($item in $items) {
        $oldText = Truncate-Text -Text $item.OldName -MaxLength 80
        $newText = Truncate-Text -Text $item.NewName -MaxLength 80

        $lvi = New-Object System.Windows.Forms.ListViewItem($oldText)
        [void]$lvi.SubItems.Add($newText)
        $lvi.ToolTipText = ("Full väg: {0}" -f $item.FullName)
        [void]$lvPreview.Items.Add($lvi)
    }
    $lvPreview.EndUpdate()

    if ($lvPreview.Items.Count -gt 0) {
        $lvPreview.EnsureVisible(0)
    }

    Add-GuiLog ("Preview visar {0} poster (alla som kommer att bytas)." -f $items.Count)
})

$btnRename.Add_Click({
    $txtLog.Clear()

    if (-not $script:CurrentHeadFolder -or -not (Test-Path -LiteralPath $script:CurrentHeadFolder -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Ingen giltig huvudmapp vald. Kör skanningen först.",
            "Byt datum-prefix"
        ) | Out-Null
        return
    }

    if (-not $script:CurrentItems -or $script:CurrentItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Inga poster att byta. Kör skanningen först.",
            "Byt datum-prefix"
        ) | Out-Null
        return
    }

    $prefix = $txtDate.Text.Trim()
    if ($prefix -notmatch '^\d{4}\s\d{2}\s\d{2}$') {
        [System.Windows.Forms.MessageBox]::Show(
            "Ogiltigt datum-prefix. Använd formatet 'yyyy MM dd'.",
            "Datum-prefix"
        ) | Out-Null
        return
    }

    $dirCount   = (@($script:CurrentItems | Where-Object { $_.IsDir -eq $true })).Count
    $fileCount  = (@($script:CurrentItems | Where-Object { $_.IsDir -eq $false })).Count
    $totalCount = $script:CurrentItems.Count

    $msg = "Genomför datum-byte till '$prefix' för {0} mappar och {1} filer (totalt {2} poster)?" -f $dirCount, $fileCount, $totalCount
    $answer = [System.Windows.Forms.MessageBox]::Show($msg, "Bekräfta datum-byte", "YesNo", "Question")

    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        Add-GuiLog "Avbrutet. Inga ändringar har gjorts."
        return
    }

    Add-GuiLog "Datum byts på filer och mappar..."
    Add-GuiLog "Detta kan ta en stund beroende på antal filer."
    Add-GuiLog "Fönstret kan visa 'Not Responding' under tiden – avvakta tills bytet är klart."
    [System.Windows.Forms.Application]::DoEvents()

    $result = Apply-Rename -BaseFolder $script:CurrentHeadFolder -Items $script:CurrentItems -TargetDatePrefix $prefix

    Add-GuiLog ""
    Add-GuiLog ("Mappar bytta: {0}, fel: {1}" -f $result.DirOK, $result.DirErr)
    Add-GuiLog ("Filer bytta:  {0}, fel: {1}" -f $result.FileOK, $result.FileErr)

    if ($result.LockedFiles -and $result.LockedFiles.Count -gt 0) {
        $lockedCount = $result.LockedFiles.Count
        Add-GuiLog ("Obs: {0} fil(er) var låsta (troligen öppna i Excel) och hoppades över." -f $lockedCount)

        $toShow = $result.LockedFiles | Select-Object -First 5
        foreach ($lf in $toShow) {
            Add-GuiLog ("  🔒 {0}" -f $lf)
        }

        if ($lockedCount -gt 5) {
            $extraLocked = $lockedCount - 5
            Add-GuiLog ("  ...och ytterligare {0} låsta filer." -f $extraLocked)
        }

        Add-GuiLog "Stäng Excel-filerna och kör skriptet igen om de också ska bytas."
    }

    Add-GuiLog ""
    Add-GuiLog "Klar. Om något blev fel, kör skriptet igen och välj rätt datum / LSP-mapp."
})

# Om HeadFolder angivits vid start – använd den direkt
if ($HeadFolder -and (Test-Path -LiteralPath $HeadFolder -PathType Container)) {
    $resolved = (Resolve-Path -LiteralPath $HeadFolder).ProviderPath
    $script:LspHits = @([pscustomobject]@{
        Name     = (Split-Path -Path $resolved -Leaf)
        FullPath = $resolved
    })
    $listLspHits.Items.Clear()
    [void]$listLspHits.Items.Add($script:LspHits[0].Name)
    $listLspHits.SelectedIndex = 0
}

[void]$form.ShowDialog()