# Suportis Crypto Script V1.5 -START-

function Sup-CreateCertificate {
    [CmdletBinding()]
    [Alias("zert","zertifikat", "makecert")]
    param(
        $cn = ""
    )
    if ($cn -eq "") {
        Write-Host "DnsName eingeben."
        Write-Host "Hier kann z.B. Kunde-Projektname oder Mitarbeitername verwendet werden" -ForegroundColor Gray
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
        Sup-ExportCertificate -cert $Global:cert
    }
}

function Sup-ExportCertificate {
    [CmdletBinding()]
    [Alias("exportzert", "expzert", "zertexport","exportzertifikat", "exptcert")]
    param(
        [ArgumentCompleter({Zertifikat_ArgumentCompleter @args})]
        $CnOrThumbprint = "",
        $Filename = "",
        $cert = $null

    )
    if ($cert -eq $null) {
        $cert = (GetZertifikateFromSubjectOrCN $CnOrThumbprint)
        if ($cert -eq $null) {
            $cert = SelectCertificate $CnOrThumbprint
        }
    }

    if ($cert -eq $null) {
        Write-Host "Kein gueltiges Zertifikate gefunden! Abbruch" -ForegroundColor Red
        return
    } 

    if ($Filename -eq "") {
        $folder = FolderSelect
        if ($folder -eq $null) {
            return
        }
        if ($folder -eq "") {
            Write-Host "Es wird der Standard-Dateiname " -NoNewline
            Write-Host "cert.pfx" -ForegroundColor Yellow -NoNewline
            Write-Host " am aktuellen Ordner verwendet"
            $Filename = "cert.pfx"
        }
        else {
            $name = $cert.Subject
            if ($name.ToUpper().StartsWith("CN=")) {
                $name = $name.SubString(3)
            }

            Write-Host "Dateiname (Enter=" -NoNewline
            Write-Host $name -ForegroundColor Yellow -NoNewline
            Write-Host "): " -NoNewline
            $e = Read-Host
            if ($e -eq "" ) {
                $Filename = (Join-Path $folder "$($name).pfx")
            }
            else {
                if (!($e.ToUpper().StartsWith(".PFX"))) {
                    $e += ".pfx"
                }
                $Filename = (Join-Path $folder $e)
            }
        }
    }

    $pwd = Read-Host "Passwort" -AsSecureString
    $pwd2 = Read-Host "Passwort wiederholen" -AsSecureString
    if (Compare-SecureString $pwd $pwd2) {
        try {
            Export-PfxCertificate "Cert:\CurrentUser\My\$($cert.Thumbprint)" -FilePath $Filename -Password $pwd 
        }
        catch {
            Write-Host "Fehler beim Exportieren des Zertifikats!" -ForegroundColor Red
            Write-Host "Ggfs. handelt es sich um ein ehemals importieres Zertifikat."
            Write-Host "Diese können nicht exportiert werden (nur der Ersteller kann exportieren)"
        }
    }
    else {
        Write-Host "Passwoerter stimmen nicht ueberein! Abbruch" -ForegroundColor Red
    }
}


function Sup-ImportCertificate {
    [CmdletBinding()]
    [Alias("importzert", "impzert", "zertimport","importzertifikat", "impcert")]
    param(
        $Filename = ""
    )
    if ($Filename -eq "") {
        $Filename = PfxFileSelect
        if ($Filename -eq $null) {
            return
        }
        if ($Filename -eq "") {
            Write-Host "Es wird der Standard-Dateiname " -NoNewline
            Write-Host "cert.pfx" -ForegroundColor Yellow -NoNewline
            Write-Host " verwendet"
            $Filename = "cert.pfx"
        }
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
    [Alias("ver", "verschlüsseln", "verschluesseln", "encrypt")]
    param(
        [ArgumentCompleter({Zertifikat_ArgumentCompleter @args})]
        $CnOrThumbprint = "",
        $Filename = "",
        [Switch] $Copy
    )

    if ($Global:cert -ne $null) {
        $cert = $Global:cert
        if (!($IsMac)) {
            if (!(Test-Path "Cert:\CurrentUser\My\$($cert.Thumbprint)")) {
                $cert = $null
                $Global:cert = $null
            }
        }
    }
    if ($Global:cert -ne $null) {
        Write-Host "Soll das Zertifikat " -NoNewline
        $sub = $Global:cert.Subject
        if ($sub.ToUpper().StartsWith("CN=")) {
            $sub = $sub.SubString(3)
        }
        Write-Host $sub -ForegroundColor Yellow -NoNewline
        Write-Host " verwendet werden? (J/n): " -NoNewline
        $e = Read-Host
        if ($e -eq "N") {
            $Global:cert = $null
            $cert = $null
        }
    }
    if ($IsMacOS) {
        if ($Global:cert -eq $null) {
            if ($Filename -eq "") {
                $Filename = PfxFileSelect
                if ($Filename -eq $null) {
                    return
                }
            }
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
            if ($Global:cert -eq $null) {
                Sup-ImportCertificate -Filename $Filename
            }
        }
        if ($Global:cert -eq $null) {
            Write-Host "Kein Zertifikat für Verschlüsselung gewählt! Abbruch" -ForegroundColor Red
            return
        }

        $e = Read-Host "Zu verschluesselnden Text (Enter=aus der Zwischenablage)"
        if ($e -eq "") {
            $e = (Get-Clipboard -Raw)
        }
        try {
            $out = Protect-CmsMessage -Content $e -To $Global:cert
            Write-Host $out -ForegroundColor Yellow
            Write-Host "Text verschluesselt" -ForegroundColor Green
            if (!($Copy)) {
                $e = Read-Host "Soll der Text des CMS-String in die Zwischenablage kopiert werden? (J/n)"
                if ($e -ne "N") {
                    $Copy = $true
                }
            }
            if ($Copy) {
                Write-Host "CMS-String in die Zwischenablage kopiert" -ForegroundColor Green
                Set-Clipboard $out
            }
            else {
                Write-Host "Zwischenablage geleert" -ForegroundColor Gray
                Set-Clipboard "-"
            }
        }
        catch {
            Write-Host "Fehler beim verschluesseln!" -ForegroundColor Red
        }
    }
    else {
        if ($Global:cert -eq $null) {
            $cert = SelectCertificate $CnOrThumbprint
        }
        else {
            $cert = $Global:cert
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
            if (!($Copy)) {
                $e = Read-Host "Soll der Text des CMS-String in die Zwischenablage kopiert werden? (J/n)"
                if ($e -ne "N") {
                    $Copy = $true
                }
            }
            if ($Copy) {
                Write-Host "CMS-String in die Zwischenablage kopiert" -ForegroundColor Green
                Set-Clipboard $out
            }
            else {
                Write-Host "Zwischenablage geleert" -ForegroundColor Gray
                Set-Clipboard "-"
            }
        }
        catch {
            Write-Host "Fehler beim verschluesseln!" -ForegroundColor Red
        }
    }
}

function IsCMSInClip {
    $con = Get-Clipboard -Raw
    return ($con.Contains("-----BEGIN CMS-----") -and $con.Contains("-----END CMS-----")) 
}
function Sup-Decrpyt {
    [CmdletBinding()]
    [Alias("ent","entschlüsseln", "entschluesseln", "decrypt")]
    param(
        $Filename = "",
        [Switch] $Copy
    )

    if ($IsMacOS) {
        if (!($Global:cert -eq $null)) {
        Write-Host "Soll das Zertifikat " -NoNewline
        $sub = $Global:cert.Subject
        if ($sub.ToUpper().StartsWith("CN=")) {
            $sub = $sub.SubString(3)
        }
        Write-Host $sub -ForegroundColor Yellow -NoNewline
        Write-Host " verwendet werden? (J/n): " -NoNewline
        $e = Read-Host
        if ($e -eq "N") {
            $Global:cert = $null
            $cert = $null
        }
        }

        if ($Global:cert -eq $null) {
            if ($Filename -eq "") {
                $Filename = PfxFileSelect
                if ($Filename -eq $null) {
                    return
                }
            }
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
            if ($Global:cert -eq $null) {
                Sup-ImportCertificate -Filename $Filename
            }
        }
        if ($Global:cert -eq $null) {
            Write-Host "Kein Zertifikat für Verschlüsselung gewählt! Abbruch" -ForegroundColor Red
            return
        }
        if (!(IsCMSInClip)) {
            $e = Read-Host "Zu entschluesselnden Text in die Zwischenablage kopieren und Enter druecken."
        }
        try {
            $out = Unprotect-CmsMessage -Content (Get-Clipboard -Raw) -To $Global:cert
            Write-Host "Entschluesselter Text:"
            Write-Host $out -ForegroundColor Yellow
            Write-Host "Text entschluesselt, entschluesselter Text in die Zwischenablage kopiert" -ForegroundColor Green
            if (!($Copy)) {
                $e = Read-Host "Soll der entschluesselte Text in die Zwischenablage kopiert werden? (J/n)"
                if ($e -ne "N") {
                    $Copy = $true
                }
            }
            if ($Copy) {
                Write-Host "Entschluesselter Text in die Zwischenablage kopiert" -ForegroundColor Green
                Set-Clipboard $out
            }
            else {
                Write-Host "Zwischenablage geleert" -ForegroundColor Gray
                Set-Clipboard "-"
            }
        }
        catch {
            Write-Host "Fehler beim entschluesseln!" -ForegroundColor Red
        }
    }
    else {
        if (!(IsCMSInClip)) {
            $e = Read-Host "Zu entschluesselnden Text in die Zwischenablage kopieren und Enter druecken."
        }
        try {
            $out = Unprotect-CmsMessage -Content (Get-Clipboard -Raw)
            Write-Host "Entschluesselter Text:"
            Write-Host $out -ForegroundColor Yellow
            Write-Host "Text verschluesselt, verschluesselter Text in die Zwischenablage kopiert" -ForegroundColor Green
            if (!($Copy)) {
                $e = Read-Host "Soll der entschluesselte Text in die Zwischenablage kopiert werden? (J/n)"
                if ($e -ne "N") {
                    $Copy = $true
                }
            }
            if ($Copy) {
                Write-Host "Entschluesselter Text in die Zwischenablage kopiert" -ForegroundColor Green
                Set-Clipboard $out
            }
            else {
                Write-Host "Zwischenablage geleert" -ForegroundColor Gray
                Set-Clipboard "-"
            }
        }
        catch {
            Write-Host "Fehler beim entschluesseln!" -ForegroundColor Red
        }
    }
}

function Sup-RemoveCertificate {
    [CmdletBinding()]
    [Alias("delzert", "zertdel", "remzert", "zertrem", "removezert", "zertremove", "deletezert", "zertdelete" )]
    param(
        [ArgumentCompleter({Zertifikat_ArgumentCompleter @args})]
        $CnOrThumbprint = ""
    )

    if ($IsMacOS) {
        Write-host "Zertifikate sind auf dem Mac nicht in einem Schlüsselbund sondern liegen nur als .pfx Datei vor."
        Write-Host "Diese können mit del <Dateiname> oder im Finder gelöscht werden."
        $e = Read-Host "Soll der Finder für das aktuelle Verzeichnis geöffnet werden? (j/N)"
        if ($e -eq "J") {
            Invoke-Item .
        }
        return
    }
    else {
        while($true) {
            $cert = SelectCertificate $CnOrThumbprint
            if ($cert -eq $null) {
                break
            }

            if ($cert -eq $null) {
                Write-Host "Kein gueltiges Zertifikat gefunden! Abbruch" -ForegroundColor Red
                return
            }
            $e = Read-Host "Zertifikat $($cert.Subject) loeschen? (j/N)"
            if ($e -eq "J") {
                Remove-Item "Cert:\CurrentUser\My\$($cert.Thumbprint)" -Force
                Write-Host "Zertifikat geloescht!"
            }
            else {
                Write-Host "Zertifikat nicht geloescht."
            }
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
    $f = $profile

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
                $clip = Get-Clipboard
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
        $clip = Get-Clipboard
        foreach($row in $clip) {
            $newContent += $row
        }
    }


    Set-Content -Path $f -Value $newContent
    Write-Host "Suportis Crypto Script installiert/upgedatet" -ForegroundColor Green
}
# Hilfsfunktionen

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

function SelectCertificate {
    param(
        $CnOrThumbprint = ""
    )

    $possible = Get-ChildItem Cert:\CurrentUser\My
    if ($CnOrThumbprint -ne "") {
        $possible = $possible | Where-Object { $_.Subject -match $CnOrThumbprint -or ($_.Thumbprint -match $CnOrThumbprint)}
    }
    if ($possible.Count -eq 0) {
        $possible = Get-ChildItem Cert:\CurrentUser\My
    }
    if ($possible.Count -eq 0) {
        return $null
    }
    if ($possible.Count -eq 1) {
        return $possible[0]
    }

    $msg = ""
    while($true) {
        $nr = 0
        Clear-Host
        foreach($p in $possible) 
        {
            Write-Host ($nr+1) -ForegroundColor Yellow -NoNewline
            $sub = $p.Subject
            if ($sub.Length -lt 40) {
                $sub += (" ") * (40 - $sub.Length)
            }
            if ($sub.ToUpper().StartsWith("CN=")) {
                $sub = $sub.SubString(3)
            }
            Write-Host ". $($sub)" -NoNewline
            Write-Host " $($p.Thumbprint)" -ForegroundColor Gray
            $nr++
        }
        if ($msg -ne "") {
            Write-Host "`n$msg" -ForegroundColor Red
        }
        Write-Host "`nNr. eingeben, Text fuer Filterung, keine Eingabe fuer Abbruch: " -NoNewline
        $e = Read-Host 
        $nr = -1
        if ([int]::TryParse($e, [Ref] $nr)) {
            if ($nr -gt 0 -and ($nr -le $possible.Count)) {
                return $possible[$nr - 1]
            }
            else {
                $msg = "Ungueltige Nummer"
            }
        }
        else {
            if ($e -ne "") {
                $possible = $possible | Where-Object { $_.Subject -match $e -or ($_.Thumbprint -match $e)}
                if ($possible.Count -eq 0) {
                    $possible = Get-ChildItem Cert:\CurrentUser\My
                }
                if ($possible.Count -eq 1) {
                    return $possible[0]
                }
            }
            else {
                return $null
            }
        }
    }
}

function PfxFileSelect {
    param(
        $folder = ""
    )

    if ($folder -eq "") {
        $folder = FolderSelect
    }
    if ($folder -eq "" -or $folder -eq $null) {
        Write-Host "Kein oder ungueltiger Ordner! Abbruch" -ForegroundColor Red
        return $null
    }
    $fullFolder = (Join-Path $folder "*.pfx")
    $filesAll = Get-ChildItem $fullFolder
    if ($filesAll -isnot [Object[]]) {
        $filesAll = @($filesAll)
    }
    if ($filesAll.Count -eq 0) {
        Write-Host "Keine .pfx Dateien im angegebenen Ordner! Abbruch" -ForegroundColor Red
        return $null
    }
    $files = $filesAll.Clone()

    $msg = ""
    while($true) {
        $nr = 0
        Clear-Host
        foreach($f in $files) 
        {
            Write-Host ($nr+1) -ForegroundColor Yellow -NoNewline
            $sub = $f.Name
            if ($sub.Length -lt 40) {
                $sub += (" ") * (40 - $sub.Length)
            }
            if ($sub.ToUpper().StartsWith("CN=")) {
                $sub = $sub.SubString(3)
            }
            Write-Host ". $($sub)" -NoNewline
            Write-Host " $($f.LastWriteTime.ToString("dd.MM.yyyy hh:mm:ss"))" -ForegroundColor Gray
            $nr++
        }
        if ($msg -ne "") {
            Write-Host "`n$msg" -ForegroundColor Red
        }
        Write-Host "`nNr. eingeben, Text fuer Filterung, keine Eingabe fuer Abbruch, *=anderer Ordner: " -NoNewline
        $e = Read-Host 
        if ($e -eq "*") {
            $folderNeu = FolderSelect
            if ($folderNeu -eq "" -or $folderNeu -eq $null) {
                $msg = "Kein oder ungueltiger Ordner" 
                continue
            }
            $fullFolderNeu = (Join-Path $folderNeu "*.pfx")
            $filesAllNeu = Get-ChildItem $fullFolderNeu
            if ($filesAllNeu -isnot [Object[]]) {
                $filesAllNeu = @($filesAllNeu)
            }
            if ($filesAllNeu.Count -eq 0) {
                $msg = "Keine .pfx Dateien im angegebenen Ordner!" 
                continue
            }
            $filesAll = $filesAllNeu
            $files = $filesAll.Clone()
            continue
        }
        $nr = -1
        if ([int]::TryParse($e, [Ref] $nr)) {
            if ($nr -gt 0 -and ($nr -le $files.Count)) {
                return $files[$nr - 1].FullName
            }
            else {
                $msg = "Ungueltige Nummer"
            }
        }
        else {
            if ($e -ne "") {
                $files = $files | Where-Object { $_.Name -match $e }
                if ($files.Count -eq 0) {
                    $files = $filesAll.Clone()
                }
                if ($files.Count -eq 1) {
                    return $files.FullName
                }
            }
            else {
                return $null
            }
        }
    }
}
function FolderSelect {
    param(
    
    )

    $homedir = [System.Environment]::GetFolderPath(40)
    $dl = Join-Path $homedir "Downloads"
    $certpath = Join-Path $homedir "Zertifikate"
    if (!(Test-Path $certpath)) {
        New-Item -ItemType Directory -Path $certpath -Force
    }
    $foldersAll = @(
        $certpath,
        $dl,
        [System.Environment]::GetFolderPath(16),
        [System.Environment]::GetFolderPath(5),
        $homedir
    )
    $foldersAll = $foldersAll | Where-Object { $_ -ne "" }
    $folders = $foldersAll.Clone()

    $msg = ""
    while($true) {
        $nr = 0
        Clear-Host
        foreach($f in $folders) 
        {
            Write-Host ($nr+1) -ForegroundColor Yellow -NoNewline
            $sub = $f
            if ($sub.Length -lt 40) {
                $sub += (" ") * (40 - $sub.Length)
            }
            Write-Host ". $($sub)" -NoNewline
            $files = (Get-ChildItem (Join-Path $f "*.pfx"))

            Write-Host " $($files.count) .pfx Datei(en)" -ForegroundColor Gray
            $nr++
        }
        if ($msg -ne "") {
            Write-Host "`n$msg" -ForegroundColor Red
        }
        Write-Host "`nNr. eingeben, Text fuer Filterung, keine Eingabe fuer Abbruch: " -NoNewline
        $e = Read-Host 
        $nr = -1
        if ([int]::TryParse($e, [Ref] $nr)) {
            if ($nr -gt 0 -and ($nr -le $folders.Count)) {
                return $folders[$nr - 1]
            }
            else {
                $msg = "Ungueltige Nummer"
            }
        }
        else {
            if ($e -ne "") {
                $folders = $folders | Where-Object { $_ -match $e }
                if ($folders.Count -eq 0) {
                    $folders = $foldersAll.Clone()
                }
                if ($folders.Count -eq 1) {
                    if ($folders -is [Object[]]) {
                        return $folders[0]
                    }
                    else
                    {
                        return $folders
                    }
                }
            }
            else {
                return $null
            }
        }
    }
}
# Suportis Crypto Script V1.5 -END-
