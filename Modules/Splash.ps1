$global:Splash = $null

function Show-Splash([string]$msg = "Startar…") {

    Add-Type -AssemblyName System.Windows.Forms, System.Drawing
    $f = New-Object Windows.Forms.Form
    $f.FormBorderStyle = 'None'
    $f.StartPosition   = 'CenterScreen'
    $f.BackColor       = [Drawing.Color]::FromArgb(0,120,215)
    $f.ForeColor       = [Drawing.Color]::White
    $f.Size            = New-Object Drawing.Size(420,120)

    $lbl = New-Object Windows.Forms.Label
    $lbl.Dock = 'Fill'
    $lbl.TextAlign = 'MiddleCenter'
    $lbl.Font = New-Object Drawing.Font('Segoe UI Semibold',12)
    $f.Controls.Add($lbl)
    $f.TopMost = $true
    $f.Show()

    $global:Splash = @{ Form = $f; Label = $lbl }
    Update-Splash $msg
    [Windows.Forms.Application]::DoEvents()
}

function Update-Splash([string]$msg) {
    if ($global:Splash) {
        $global:Splash.Label.Text = $msg
        [Windows.Forms.Application]::DoEvents()
    }
}

function Close-Splash() {
    if ($global:Splash) {
        $global:Splash.Form.Close()
        $global:Splash.Form.Dispose()
        $global:Splash = $null
    }
}