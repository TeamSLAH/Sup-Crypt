$c = Invoke-RestMethod "https://raw.githubusercontent.com/TeamSLAH/Sup-Crypt/main/script.ps1" -Headers @{"Cache-Control"="no-cache"}
function Sup-Install {
    [CmdletBinding()]
    [Alias("instzert", "zertinst")]
    param(
        $code = ""
    )
    $f = $profile
    if ($code -eq "") {
        Write-Host "Skript in die Zwischenablage kopieren und Enter druecken"
        Read-Host
        $code = Get-Clipboard
    }
    else {
        $code = $code.Split("`n")
    }

    if (!(Test-Path $f)) {
        New-Item $f -Force
    }
    $content = Get-Content -Path $f
    $newContent = @()
    $inscript=$false
    $installed = $false
    foreach($row in $content) {
        if ($row.StartsWith("# Suportis Crypto Script")) {
            if ($row.EndsWith("-START-")) {
                $inscript=$true
            }
            elseif ($row.EndsWith("-END-")) {
                $inscript = $false
                $clip = $code
                foreach($row in $clip) {
                    $newContent += $row
                }
                $installed = $true
                continue
            }
        }
        if (!($inscript)) {
            $newContent += $row
        }
    }
    if (!($installed)) {
        $clip = $code
        foreach($row in $clip) {
            $newContent += $row
        }
    }

    Set-Content -Path $f -Value $newContent
    Write-Host "Suportis Crypto Script installiert/upgedatet" -ForegroundColor Green
}

Sup-Install $c

Invoke-Expression $c

