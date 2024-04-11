# Suportis Crypto Script V1.0 -START-
function Sup-CreateCertificate {
    [CmdletBinding()]
    [Alias("zert","zertifikat", "makecert")]
    param(
        $cn = ""
    )
    if ($cn -eq "") {
        Write-Host "DnsName eingeben."
        Write-Host "Hier kann z.B. Kunde-Projektname verwendet werden" -ForegroundColor Gray
        $cn = Read-Host "DnsName"
    }
    
    $Global:cert = New-SelfSignedCertificate -DnsName $cn -CertStoreLocation "Cert:\CurrentUser\My" -KeyUsage KeyEncipherment,DataEncipherment,KeyAgreement -Type DocumentEncryptionCert
    Write-Host "Zertifikat erstellt, in `$cert gespeichert."
    Write-Host "Thumbprint " -NoNewline
    Write-Host $cert.Thumbprint -NoNewline -ForegroundColor Yellow
    Write-Host " in die Zwischenablage kopiert."
    Set-Clipboard $global:cert.Thumbprint

    $e = Read-Host "Soll dieses gleich exportiert werden? (J/n)"
    if ($e -ne "N") {
        Sup-ExportCertificate
    }
}

function Zertifikat_ArgumentCompleter {
    param(
        $commandName,
        $parameterName,
        $wordToComplete,
        $commandAst,
        $fakeBoundParameters
    )
 
    $posibleItems = @()

    $zertifikate = Get-ChildItem Cert:\CurrentUser\My\

    foreach($zert in $zertifikate) {
        $sub = $zert.Subject
        if ($sub -eq $null) {
            $sub=""
        }
        if ($wordToComplete.ToUpper().StartsWith("CN=")) {
            if ($sub -match "^$($wordToComplete).*") {
                $item = [PSCustomObject]@{
                    shortName = $zert.Subject
                    longName = $zert.Subject
                    toolTip = "Thumbprint: $($zert.Thumbprint)"
                }
                $posibleItems += $item
            }
        }  
        else {
            if ($sub.Length -gt 3) {
                if ($sub.SubString(3) -match "$($wordToComplete).*" -or ($zert.Thumbprint -match "^$($wordToComplete)")) {
                    $item = [PSCustomObject]@{
                        shortName = $zert.Subject.SubString(3)
                        longName = $zert.Subject.SubString(3)
                        toolTip = "Thumbprint: $($zert.Thumbprint)"
                    }
                    $posibleItems += $item
                }
            } 
        }
    }
    # Auswertung
    foreach($item in $posibleItems) {
        New-Object System.Management.Automation.CompletionResult (
            $item.shortName,
            $item.longName,
            "ParameterValue",
            $item.toolTip
        )
    }
}
function GetZertifikateFromSubjectOrCN {
    param(
        $subjectOrCn = ""
    )
 
    $zertifikate = Get-ChildItem Cert:\CurrentUser\My\

    foreach($zert in $zertifikate) {
        $sub = $zert.Subject
        if ($sub -eq $null) {
            $sub=""
        }
        if ($subjectOrCn.ToUpper().StartsWith("CN=")) {
            if ($sub -eq "$($subjectOrCn)") {
                return $zert
            }
        }  
        else {
            if ($sub.Length -gt 3) {
                if ($sub.SubString(3) -eq "$($subjectOrCn)" -or ($zert.Thumbprint -eq "$($subjectOrCn)")) {
                    return $zert
                }
            } 
        }
    }
}
    

function Sup-ExportCertificate {
    [CmdletBinding()]
    [Alias("exportzert","zertexport","exportzertifikat", "exptcert")]
    param(
        [ArgumentCompleter({Zertifikat_ArgumentCompleter @args})]
        $CnOrThumbprint = "",
        $Filename = ""

    )
    if ($CnOrThumbprint -eq "" -and $Global:cert -ne $null) {
        $cert = $Global:cert
    }
    else {
        $cert = (GetZertifikateFromSubjectOrCN $CnOrThumbprint)
    }

    if ($cert -eq $null) {
        Write-Host "Kein gueltiges Zertifikate gefunden! Abbruch" -ForegroundColor Red
        return
    } 

    if ($Filename -eq "") {
        Write-Host "Es wird der Standard-Dateiname " -NoNewline
        Write-Host "cert.pfx" -ForegroundColor Yellow -NoNewline
        Write-Host " verwendet"
        $Filename = "cert.pfx"
    }

    $pwd = Read-Host "Passwort" -AsSecureString
    $pwd2 = Read-Host "Passwort wiederholen" -AsSecureString
    if (Compare-SecureString $pwd $pwd2) {
        Try {
            Export-PfxCertificate "Cert:\CurrentUser\My\$($cert.Thumbprint)" -FilePath $Filename -Password $pwd 
        }
        catch {
            Write-Host "Nicht exportiert. Ggfs. handelt es sich um ein importiertes Zertifikat. Diese koennen nicht exportiert werden!" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Passwoerter stimmen nicht ueberein! Abbruch" -ForegroundColor Red
    }
}

function Compare-SecureString {
    param(
      [Security.SecureString]
      $secureString1,
  
      [Security.SecureString]
      $secureString2
    )
    try {
      $bstr1 = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString1)
      $bstr2 = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString2)
      $length1 = [Runtime.InteropServices.Marshal]::ReadInt32($bstr1,-4)
      $length2 = [Runtime.InteropServices.Marshal]::ReadInt32($bstr2,-4)
      if ( $length1 -ne $length2 ) {
        return $false
      }
      for ( $i = 0; $i -lt $length1; ++$i ) {
        $b1 = [Runtime.InteropServices.Marshal]::ReadByte($bstr1,$i)
        $b2 = [Runtime.InteropServices.Marshal]::ReadByte($bstr2,$i)
        if ( $b1 -ne $b2 ) {
          return $false
        }
      }
      return $true
    }
    finally {
      if ( $bstr1 -ne [IntPtr]::Zero ) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr1)
      }
      if ( $bstr2 -ne [IntPtr]::Zero ) {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr2)
      }
    }
}

function Sup-ImportCertificate {
    [CmdletBinding()]
    [Alias("importzert","zertimport","importzertifikat", "impcert")]
    param(
        $Filename = ""
    )
    if ($Filename -eq "") {
        Write-Host "Es wird der Standard-Dateiname " -NoNewline
        Write-Host "cert.pfx" -ForegroundColor Yellow -NoNewline
        Write-Host " verwendet"
        $Filename = "cert.pfx"
    }
    if (!(Test-Path $Filename)) {
        Write-Host "Zertifikat-Datei " -NoNewline -ForegroundColor Red
        Write-Host $Filename -NoNewline -ForegroundColor Yellow
        Write-Host " nicht gefudnen! Abbruch" -ForegroundColor Red
        return
    }
    
    $pwd = Read-Host "Passwort" -AsSecureString
    try {
        if ($IsMacOS) {
            $Global:cert = Get-PfxCertificate -FilePath $Filename -Password $pwd
        }
        else {
            Import-PfxCertificate -FilePath $Filename Cert:\CurrentUser\My\ -Password $pwd
        }
    }
    catch {
        Write-Host "Import fehlgeschlagen!" -ForegroundColor Red
    }
}

function Sup-Encrpyt {
    [CmdletBinding()]
    [Alias("verschlüsseln", "verschluesseln", "encrypt")]
    param(
        [ArgumentCompleter({Zertifikat_ArgumentCompleter @args})]
        $CnOrThumbprint = "",
        $Filename = ""
    )

    if ($IsMacOS) {
        if ($Global:cert -eq $null) {
            if ($Filename -eq "") {
                Write-Host "Es wird der Standard-Dateiname " -NoNewline
                Write-Host "cert.pfx" -ForegroundColor Yellow -NoNewline
                Write-Host " verwendet"
                $Filename = "cert.pfx"
            }
            if (!(Test-Path $Filename)) {
                Write-Host "Zertifikat-Datei " -NoNewline -ForegroundColor Red
                Write-Host $Filename -NoNewline -ForegroundColor Yellow
                Write-Host " nicht gefunden! Abbruch" -ForegroundColor Red
                return
            }
            Sup-ImportCertificate -Filename $Filename
        }

        $e = Read-Host "Zu verschluesselnden Text (Enter=aus der Zwischenablage)"
        if ($e -eq "") {
            $e = (Get-Clipboard -Raw)
        }
        try {
            $out = Protect-CmsMessage -Content $e -To $Global:cert
            Write-Host $out -ForegroundColor Yellow
            Write-Host "Text verschluesselt, verschluesselter Text in die Zwischenablage kopiert" -ForegroundColor Green
            Set-Clipboard $out
        }
        catch {
            Write-Host "Fehler beim verschluesseln!" -ForegroundColor Red
        }
    }
    else {
        $cert = $null
        if ($CnOrThumbprint -eq "") {
            $cn = Read-Host "DNS-Name eingeben (Enter fuer 'Suportis')"
            if ($cn -eq "") {
                $cn = "Suportis"
            }

            $cert = (GetZertifikateFromSubjectOrCN $cn)
        }
        else {
            $cert = (GetZertifikateFromSubjectOrCN $CnOrThumbprint)
        }

        if ($cert -eq $null) {
            Write-Host "Kein gueltiges Zertifikat gefunden! Abbruch" -ForegroundColor Red
            return
        }
        $e = Read-Host "Zu verschluesselnden Text (Enter=aus der Zwischenablage)"
        if ($e -eq "") {
            $e = (Get-Clipboard -Raw)
        }
        try {
            $out = Protect-CmsMessage -Content $e -to $cert.Subject
            Write-Host $out -ForegroundColor Yellow
            Write-Host "Text verschluesselt, verschluesselter Text in die Zwischenablage kopiert" -ForegroundColor Green
            Set-Clipboard $out
        }
        catch {
            Write-Host "Fehler beim verschluesseln!" -ForegroundColor Red
        }
    }
}


function Sup-Decrpyt {
    [CmdletBinding()]
    [Alias("entschlüsseln", "entschluesseln", "decrypt")]
    param(
        $Filename = ""
    )

    if ($IsMacOS) {
        if ($Global:cert -eq $null) {
            if ($Filename -eq "") {
                Write-Host "Es wird der Standard-Dateiname " -NoNewline
                Write-Host "cert.pfx" -ForegroundColor Yellow -NoNewline
                Write-Host " verwendet"
                $Filename = "cert.pfx"
            }
            if (!(Test-Path $Filename)) {
                Write-Host "Zertifikat-Datei " -NoNewline -ForegroundColor Red
                Write-Host $Filename -NoNewline -ForegroundColor Yellow
                Write-Host " nicht gefudnen! Abbruch" -ForegroundColor Red
                return
            }
        }
        Sup-ImportCertificate -Filename $Filename
        if ($Global:cert -eq $null) {
            Write-Host "Kein Zertifikat geladen/importiert! Abbruch" -ForegroundColor Red
            return
        }

        $e = Read-Host "Zu entschluesselnden Text in die Zwischenablage kopieren und Enter druecken."
        try {
            $out = Unprotect-CmsMessage -Content (Get-Clipboard -Raw) -To $Global:cert
            Write-Host "Entschluesselter Text:"
            Write-Host $out -ForegroundColor Yellow
            Write-Host "Text entschluesselt, entschluesselter Text in die Zwischenablage kopiert" -ForegroundColor Green
            Set-Clipboard $out
        }
        catch {
            Write-Host "Fehler beim entschluesseln!" -ForegroundColor Red
        }
    }
    else {
        $e = Read-Host "Zu entschluesselnden Text in die Zwischenablage kopieren und Enter druecken."
        try {
            $out = Unprotect-CmsMessage -Content (Get-Clipboard -Raw)
            Write-Host "Entschluesselter Text:"
            Write-Host $out -ForegroundColor Yellow
            Write-Host "Text verschluesselt, verschluesselter Text in die Zwischenablage kopiert" -ForegroundColor Green
            Set-Clipboard $out
        }
        catch {
            Write-Host "Fehler beim verschluesseln!" -ForegroundColor Red
        }
    }
}

function Sup-Install {
    [CmdletBinding()]
    [Alias("instzert", "zertinst")]
    param(

    )
    Write-Host "Skript in die Zwischenablage kopieren und Enter druecken"
    Read-Host
    $f = "z:\temp\cert2\abc\def\xyz.txt"

    if (!(Test-Path $f)) {
        New-Item $f -Force
    }
    $content = Get-Content -Path $f
    $newContent = @()
    $inscript=$false
    foreach($row in $content) {
        if ($row.StartsWith("# Suportis Crypto Script")) {
            if ($row.EndsWith("-START-")) {
                $inscript=$true
            }
            elseif ($row.EndsWith("-END-")) {
                $inscript = $false
                $clip = Get-Clipboard
                foreach($row in $clip) {
                    $newContent += $row
                }
                continue
            }
        }
        if (!($inscript)) {
            $newContent += $row
        }
    }


    Set-Content -Path $f -Value $newContent
    Write-Host "Suportis Crypto Script installiert/upgedatet" -ForegroundColor Green
}
# Suportis Crypto Script V1.0 -END-
