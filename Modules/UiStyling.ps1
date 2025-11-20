function Add-ShortcutItem {
    param(

        [System.Windows.Forms.ToolStripMenuItem]$Parent,
        [string]$Text,
        [string]$Target
    )

    $it = New-Object System.Windows.Forms.ToolStripMenuItem($Text)
    $it.Tag = $Target

    $it.add_Click({
        param($s, $e)
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
                [System.Windows.Forms.MessageBox]::Show("Hittar inte sökvägen:`n$t", "Genväg", 'OK', 'Warning') | Out-Null
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Kunde inte öppna:`n$t`n$($_.Exception.Message)", "Genväg") | Out-Null
        }
    })
    [void]$Parent.DropDownItems.Add($it)
}

function Get-WinAccentColor {
    try {
        $p = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\DWM' -ErrorAction Stop
        $argb = if ($p.AccentColor) { $p.AccentColor } elseif ($p.ColorizationColor) { $p.ColorizationColor } else { $null }
        if ($argb) { return [System.Drawing.Color]::FromArgb([int]$argb) }
    } catch {}
    return [System.Drawing.Color]::FromArgb(38, 120, 178)
}

function New-Color { param([int]$A, [int]$R, [int]$G, [int]$B) [System.Drawing.Color]::FromArgb($A, $R, $G, $B) }
function Darken { param([System.Drawing.Color]$c, [double]$f = 0.85) New-Color 255 ([int]($c.R * $f)) ([int]($c.G * $f)) ([int]($c.B * $f)) }
function Lighten { param([System.Drawing.Color]$c, [double]$f = 0.12) New-Color 255 ([int]([Math]::Min(255, $c.R + (255 - $c.R) * $f))) ([int]([Math]::Min(255, $c.G + (255 - $c.G) * $f))) ([int]([Math]::Min(255, $c.B + (255 - $c.B) * $f))) }
 
$Accent = Get-WinAccentColor
$AccentBorder = Darken $Accent 0.75
$AccentHover = Lighten $Accent 0.12
$AccentDisabled = New-Color 255 210 210 210

function Set-AccentButton {
    param([System.Windows.Forms.Button]$Btn, [switch]$Primary)
    $Btn.FlatStyle = 'Flat'
    $Btn.FlatAppearance.BorderSize = 1
    $Btn.FlatAppearance.BorderColor = $AccentBorder
    $Btn.FlatAppearance.MouseOverBackColor = $AccentHover

    if ($Primary) {
        $Btn.BackColor = $Accent
        $Btn.ForeColor = [System.Drawing.Color]::White
        $Btn.UseVisualStyleBackColor = $false
    }

    else {
        $Btn.BackColor = [System.Drawing.Color]::White
        $Btn.ForeColor = [System.Drawing.Color]::Black
        $Btn.UseVisualStyleBackColor = $false
    }

    if ($Btn.Height -lt 30) { $Btn.Height = 30 }
    $Btn.add_EnabledChanged({
        if ($this.Enabled) {
            if ($Primary) {
                $this.BackColor = $Accent
                $this.ForeColor = [System.Drawing.Color]::White
            }
            else {
                $this.BackColor = [System.Drawing.Color]::White
                $this.ForeColor = [System.Drawing.Color]::Black
            }
        }
        else {
            $this.BackColor = $AccentDisabled
            $this.ForeColor = [System.Drawing.Color]::Gray
        }
    })
}