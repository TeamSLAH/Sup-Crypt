$SUPVERSION = "1.21"
function Sup-Version {
    Write-Host $SUPVERSION
}

Write-Host "SUP-Crypt Version $SUPVERSION"
$c = Invoke-RestMethod "https://raw.githubusercontent.com/TeamSLAH/Sup-Crypt/main/version.json"
if ($c.Version -eq $SUPVERSION) {
    Write-Host "    Aktuelle Version, es gibt keine Updates"
}
else {
    Write-Host "    Updates Verfügbar auf Version " -NoNewline
    Write-Host "$($c.Version)" -NoNewline -ForegroundColor Red
    Write-Host "."
    Write-Host "Sup-Update" -NoNewline -ForegroundColor Yellow
    Write-Host " um neuste Version zu installieren."
    Write-Host "Sup-WhatsNew" -NoNewline -ForegroundColor Yellow
    Write-Host " um Änderungen der neuen Version(en) anzuzeigen."
}
function Sup-WhatsNew {
    param(
        $ask = $true
    )

    $c = Invoke-RestMethod "https://raw.githubusercontent.com/TeamSLAH/Sup-Crypt/main/version.json"
    Write-Host "Installierte Version: " -NoNewline
    Write-Host $SUPVERSION -ForegroundColor Red
    Write-Host "Neuste Version      : " -NoNewline
    Write-Host $c.Version -ForegroundColor Green
    Write-Host
    [int] $major = $SUPVERSION.SubString(0, $SUPVERSION.IndexOf("."))
    [int] $minor = $SUPVERSION.Substring($SUPVERSION.IndexOf(".") + 1)
    $foundNew = $false
    foreach($version in $c.History) {
        [int] $ma = $version.Version.SubString(0, $version.Version.IndexOf("."))
        [int] $mi = $version.Version.Substring($version.Version.IndexOf(".") + 1)
        if ($ma -gt $major -or ($ma -eq $major -and $mi -gt $minor)) {
            $foundNew = $true
            Write-Host "Version : $($version.Version)"
            Write-Host "Stand   : $($version.Datum)"
            Write-Host "Neuerung: $($version.Beschreibung)"
            Write-Host
        }
    }
    if ($foundNew -and $ask) {
        $e = Read-Host "Neue Version installieren? (J/n)"
        if ($e -ne "N") {
            Sup-Update
        }
    }
}
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
    if ($IsMacOS) {
        Write-Host "Export von Zertifikaten auf dem Mac ist vorerst nicht möglich!" -ForegroundColor Red
        return
    }

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
    if ($IsMacOS) {
        Write-Host "Export von Zertifikat $($cert)"
    }
    else {
        Write-Host "Export von Zertigikat $($cert.Subject)"
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

    $pwd = Read-Host "Passwort (leer für automatische generierung)" -MaskInput
    if ($pwd -eq "") {
        $pwd = GeneratePassword
        Write-Host "Erstelltes Passwort in die Zwischenablage kopiert" -ForegroundColor Green
        Write-Host "Erstelltes Passwort ausgeben? (j/N): " -NoNewline
        $e = Read-Host
        if ($e -eq "J") {
            Write-Host "Passwort: $($pwd)" -ForegroundColor Gray
        }
        Set-Clipboard $pwd
        $pwd = ConvertTo-SecureString -String $pwd -AsPlainText
        $pwd2 = $pwd
    }
    else {
        $pwd2 = Read-Host "Passwort wiederholen" -AsSecureString
    }
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
    
    try {
        if ($IsMacOS) {
            security import "$Filename" -k ~/Library/Keychains/login.keychain
            # $Global:cert = Get-PfxCertificate -FilePath $Filename -Password $pwd
        }
        else {
            $pwd = Read-Host "Passwort" -AsSecureString
            $Global:cert = Import-PfxCertificate -FilePath $Filename Cert:\CurrentUser\My\ -Password $pwd
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
        $Path = "",
        $ZertFile = "",
        [Switch] $Copy
    )

    if ($IsMacOS) {
        $CnOrThumbprint = SelectCertificate -CnOrThumbprint $CnOrThumbprint
        if ($CnOrThumbprint -eq "" -or $CnOrThumbprint -eq $null) {
            Write-Host "Kein Zertifikate gewaehlt! Abbruch" -ForegroundColor Red
            return
        }
        Write-Host "Verschluesselung mit Zertifikat " -NoNewline
        Write-Host "$($CnOrThumbprint)" -ForegroundColor Yellow
        if ($ZertFile -ne "") {
            if (Test-Path $ZertFile) {
                Sup-ImportCertificate -Filename $ZertFile
            }
        }

        if ($Path -ne "") {
            if (Test-Path $Path) {
                $Path = (Resolve-Path $Path.ToString()).Path
                $e = [Convert]::ToBase64String([io.file]::ReadAllBytes($Path))
            }
            else {
                $Path=""
            }
        }
        if ($Path -eq "") {
            $e = Read-Host "Zu verschluesselnden Text (Enter=aus der Zwischenablage)"
            if ($e -eq "") {
                $e = (Get-Clipboard -Raw)
            }
        }
        try {
            $out = Protect-CmsMessage -Content $e -To $CnOrThumbprint

            if ($Path -ne "") {
                $filename = (Get-Item $Path).Name
                $out = "###FILENAME###:$filename`n$out"
            }
            if ($out.Length -lt 1kb) {
                Write-Host $out -ForegroundColor Yellow
            }
            $out = "Anleitung/Installation/Git Repo: [Github](https://github.com/TeamSLAH/Sup-Crypt/tree/main)`nEntschluesseln mit : ``ent```n```````n$out`n```````n"

            Write-Host "Text verschluesselt" -ForegroundColor Green
            if (!($Copy)) {
                if ($out.Length -lt 1kb) {
                    $e = Read-Host "Soll der Text des CMS-String in die Zwischenablage kopiert werden? (J/n)"
                }
                else {
                    $e = "J"
                }
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
        if ($ZertFile -eq "") {
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
            if ($ZertFile -ne "") {
                Sup-ImportCertificate -Filename $ZertFile
            }
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
            Write-Host "Verschluesselung mit Zertifikat " -NoNewline
            Write-Host "$($cert.Subject)" -ForegroundColor Yellow
            if ($Path -ne "") {
                if (Test-Path $Path) {
                    $Path = (Resolve-Path $Path.ToString()).Path
                    $e = [Convert]::ToBase64String([io.file]::ReadAllBytes($Path))
                }
                else {
                    $Path=""
                }
            }
            if ($Path -eq "") {
                $e = Read-Host "Zu verschluesselnden Text (Enter=aus der Zwischenablage)"
                if ($e -eq "") {
                    $e = (Get-Clipboard -Raw)
                }
            }
            try {
                $out = Protect-CmsMessage -Content $e -to $cert.Subject
                if ($out.Length -lt 1kb) {
                    Write-Host $out -ForegroundColor Yellow
                }
                Write-Host "Text verschluesselt, verschluesselter Text in die Zwischenablage kopiert" -ForegroundColor Green
                if (!($Copy)) {
                    if ($out.Length -lt 1kb) {
                        $e = Read-Host "Soll der Text des CMS-String in die Zwischenablage kopiert werden? (J/n)"
                    }
                    else {
                        $e = "J"
                    }
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
}

function IsCMSInClip {
    $con = Get-Clipboard -Raw
    return ($con.Contains("-----BEGIN CMS-----") -and $con.Contains("-----END CMS-----")) 
}
function Sup-Decrpyt {
    [CmdletBinding()]
    [Alias("ent","entschlüsseln", "entschluesseln", "decrypt")]
    param(
        [Switch] $Copy
    )
    if (!(IsCMSInClip)) {
        $e = Read-Host "Zu entschluesselnden Text in die Zwischenablage kopieren und Enter druecken."
    }

    try {
        $cms = (Get-Clipboard -Raw)
        $out = Unprotect-CmsMessage -Content ($cms)
        if ($cms.Contains("###FILENAME###:")) {
            $i = $cms.IndexOf("###FILENAME###:")
            $r = $cms.IndexOf("`n", $i)
            $filename = $cms.SubString($i+15, $r-15)

            while( ([int]$filename[$filename.Length - 1]) -eq 13) {
                $filename = $filename.SubString(0, $filename.Length - 1)
            }
            # Download Folder
            $df = Resolve-Path "~/Downloads"
            if (Test-Path $df) {
                $Path = (Join-Path $df $filename)
            }
            else {
                # Temp
                $df = [System.IO.Path]::GetTempPath()
                $Path = (Join-Path $df $filename)
            }



            EncryptFile $Path $out
        }
        else {
            CopyStringPart $out
        }
    }
    catch {
        Write-Host "Fehler beim entschluesseln!" -ForegroundColor Red
    }
}
function EncryptFile {
    param(
        $Path,
        $Out
    )

    $Path = $Path.ToString()
    $destPath = $Path 

    $a = [convert]::FromBase64String($out)
    [io.file]::WriteAllBytes($destPath, $a)
    Write-Host "Entschluesselte Datei " -NoNewline
    Write-Host $destPath -ForegroundColor Yellow

    $p = (Get-Item $destPath).Directory.FullName
    while($true) {
        $e = Read-Host "(O)effnen, (P)fad offnen, Pfad (k)opieren, (D)atei (l)oeschen? Enter=fertig (o/p/k/d/l)"
        if ($e -eq "P") {
            Invoke-Item $p
        }
        elseif ($e -eq "K") {
            Set-Clipboard $p
        }
        elseif ($e -eq "D" -or $e -eq "L") {
            Remove-Item $destPath
            break
        }
        elseif ($e -eq "O") {
            Invoke-Item $destPath
        }
        else {
            break
        }
    }
}
function ParseKeychain {
    $r = security find-identity ~/Library/Keychains/login.keychain
    $m = $r -match '.*\) (?<Thumbprint>.*) "(?<Subject>.*)"'
    $list = @()
    foreach($match in $m) {
        $g = [regex]::Match($match, '.*\) (?<Thumbprint>.*) "(?<Subject>.*)"').Groups
        $item = [PSCustomObject]@{
            Thumbprint = $g["Thumbprint"].Value
            Subject = $g["Subject"].Value
        }
        $list += $item
    }
    return $list
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

function Sup-ListExports {
    [CmdletBinding()]
    [Alias("listexp", "explist")]
    param(
    
    )
    $homedir = [System.Environment]::GetFolderPath(40)
    $dl = Join-Path $homedir "Downloads"
    $certpath = Join-Path $homedir "Zertifikate"
    if (!(Test-Path $certpath)) {
        New-Item -ItemType Directory -Path $certpath -Force
    }
    $folders = @(
        $certpath,
        $dl,
        [System.Environment]::GetFolderPath(16),
        [System.Environment]::GetFolderPath(5),
        $homedir
    )
    $folders = $folders | Where-Object { $_ -ne "" }
    $fileList = @()
    foreach($folder in $folders) {
        $files = Get-ChildItem (Join-Path $folder "*.pfx")
        if ($files.Count -gt 0) {
            $fileList += $files
        }
    }
    $fileList = $fileList | Sort-Object LastWriteTime -Descending
    if ($fileList -isnot [object[]]) {
        $fileList = @($fileList)
    }
    $fileListAll = $fileList.Clone()

    $msg = ""
    $selectedFile = $null
    while($true) {
        $nr = 0
        Clear-Host
        foreach($f in $fileList) 
        {
            if ($nr % 2 -eq 0) {
                Write-Host ($nr+1) -ForegroundColor Yellow -NoNewline
            }
            else {
                Write-Host ($nr+1) -ForegroundColor DarkYellow -NoNewline -BackgroundColor DarkGray
            }
            $sub = $f.Name
            if ($sub.Length -lt 30) {
                $sub += (" ") * (30 - $sub.Length)
            }
            if ($nr % 2 -eq 0) {
                Write-Host ". $($sub)" -NoNewline
            }
            else {
                Write-Host ". $($sub)" -NoNewline -ForegroundColor Black -BackgroundColor DarkGray 
            }
            $diff = [DateTime]::Now - $f.LastWriteTime
            $diff = [DateTime]::Now - $f.LastWriteTime
            if ($diff.TotalHours -ge 24) {
                if ($diff.Days -eq 0) {
                    $diffFormatted = "heute"
                } elseif ($diff.Days -eq 1) {
                    $diffFormatted = "gestern"
                } else {
                    if ($diff.Days -le 7) {
                        $diffFormatted = "vor $($diff.Days) Tagen"
                    }
                    else {
                        $diffFormatted = $f.LastWriteTime("dd.MM.yyyy")
                    }
                }
                $diffFormatted += " um $($f.LastWriteTime.ToString("HH:mm:ss"))"
            } elseif ($diff.TotalHours -ge 1) {
                $diffFormatted = "vor $($diff.Hours) Stunden"
            } elseif ($diff.TotalMinutes -ge 1) {
                $diffFormatted = "vor $($diff.Minutes) Minuten"
            } else {
                $diffFormatted = "vor wenigen Sekunden"
            }

            if ($diffFormatted.Length -lt 27) {
                $diffFormatted += (" ") * (27 - $diffFormatted.Length)
            }
            if ($nr % 2 -eq 0) {
                Write-Host $diffFormatted -ForegroundColor Gray -NoNewline
                Write-Host " in $($f.Directory.FullName)"
            }
            else 
            {
                Write-Host $diffFormatted -ForegroundColor Black -NoNewline -BackgroundColor DarkGray
                Write-Host " in $($f.Directory.FullName)" -ForegroundColor Black -BackgroundColor DarkGray
            }
            $nr++
        }
        if ($msg -ne "") {
            Write-Host "`n$msg" -ForegroundColor Red
            $msg=""
        }
        Write-Host "`nNr. eingeben, Text fuer Filterung, keine Eingabe fuer Abbruch: " -NoNewline
        $e = Read-Host 
        $nr = -1
        if ([int]::TryParse($e, [Ref] $nr)) {
            if ($nr -gt 0 -and ($nr -le $fileList.Count)) {
                $selectedFile = $fileList[$nr - 1]
                break
            }
            else {
                $msg = "Ungueltige Nummer"
            }
        }
        else {
            if ($e -ne "") {
                $fileList = $fileList | Where-Object { $_ -match $e }
                if ($fileList.Count -eq 0) {
                    $fileList = $fileListAll.Clone()
                }
                if ($fileList.Count -eq 1) {
                    if ($fileList -is [Object[]]) {
                        return $fileList[0]
                    }
                    else
                    {
                        return $fileList
                    }
                }
            }
            else {
                return $null
            }
        }
    }
    if ($selectedFile -eq $null) {
        return $null
    }

    $msg = ""
    [System.ConsoleColor] $color = 'Red'
    while($true) {
        Clear-Host
        Write-Host "Export-Datei: $($selectedFile.Name)"
        Write-Host "Vom Ordner  : $($selectedFile.Directory.FullName)"
        Write-Host
        Write-Host "Optionen:"
        if ($IsMacOS)  {
            Write-Host "X/O/F" -ForegroundColor Yellow -NoNewline
            Write-Host ": Ordner im Finder oeffnen"
        }
        else {
            Write-Host "X/O  " -ForegroundColor Yellow -NoNewline
            Write-Host ": Ordner im Explorer oeffnen"
        }



        Write-Host "I    " -ForegroundColor Yellow -NoNewline
        Write-Host ": Importieren"

        Write-Host "D    " -ForegroundColor Yellow -NoNewline
        Write-Host ": Loeschen (delete)"

        if ($IsMacOS) {
            Write-Host "E    " -ForegroundColor Yellow -NoNewline
            Write-Host ": Entschluesseln"
        }

        if ($IsMacOS) {
            Write-Host "V    " -ForegroundColor Yellow -NoNewline
            Write-Host ": Verschluesseln"
        }
        else {
            Write-Host "V    " -ForegroundColor Yellow -NoNewline
            Write-Host ": Verschluesseln (mit vorherigem Import)"
        }

        Write-Host "C    " -ForegroundColor Yellow -NoNewline
        Write-Host ": Dateipfad kopieren"

        Write-Host "Keine Eingabe: Abbruch"

        Write-Host
        if ($msg -ne "") {
            Write-Host "`n$($msg)" -ForegroundColor $color
            $msg=""
        }
        $e = Read-Host "Option"
        if ($e -eq "") {
            return
        }
        if ($e -eq "X" -or $e -eq "O" -or $e -eq "F") {
            Invoke-Item $selectedFile.Directory.FullName
            if ($IsMacOS) {
                $msg = "Finder geoeffnet"
            }
            else {
                $msg = "Explorer geoeffnet"
            }
            $color = 'Green'
        }
        elseif ($e -eq "I") {
           Sup-ImportCertificate -Filename $selectedFile.FullName 
           $msg = "Importvorgang ausgeloest/abgeschlossen"
           $color = 'Green'
        }
        elseif ($e -eq "D") {
            # Delete
            Write-Host "Datei " -NoNewline
            Write-Host $selectedFile.FullName -ForegroundColor Yellow -NoNewline
            Write-Host " wirklich loeschen? (j/N):"
            $e = Read-Host
            if ($e -eq "J") {
                Remove-Item $selectedFile
                Write-Host "Datei geloescht" -ForegroundColor Green
                return
            }
        }
        elseif ($e -eq "E" -and $IsMacOS) {
            # Entschlüsseln (nur Mac)
            Sup-Decrpyt -Filename $selectedFile.FullName
            $msg = "Entschluesselung abgeschlossen"
            $color = 'Green'
        }
        elseif ($e -eq "V") {
            # Verschlüsseln
            Sup-Encrpyt -Filename $selectedFile.FullName
            $msg = "Verschluesselung abgeschlossen"
            $color = 'Green'
        }
        elseif ($e -eq "C") {
            # Copy Path
            Set-Clipboard $selectedFile.FullName
            $msg = "Vollstaendiger Pfad in die Zwischenablage kopiert"
            $color = 'Green'
        }
        else {
            $msg = "Ungueltige Eingabe"
            $color = 'Red'
        }
    }


}
function Sup-Update {
    [CmdletBinding()]
    param(

    )
    Sup-WhatsNew -ask $false
    if ($IsMacOS) {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/TeamSLAH/Sup-Crypt/main/install.ps1'))
    }
    else {
        Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/TeamSLAH/Sup-Crypt/main/install.ps1'))
    }
}
function Sup-ListCertificates {
    [CmdletBinding()]
    [Alias("listzert", "zertlist", "zertlst", "lstzert")]
    param(
        [ArgumentCompleter({Zertifikat_ArgumentCompleter @args})]
        $CnOrThumbprint = ""
    )

    if ($IsMacOS) {
        $CnOrThumbprint = SelectCertificate -CnOrThumbprint $CnOrThumbprint
        if ($CnOrThumbprint -eq "" -or $CnOrThumbprint -eq $null) {
            Write-Host "Kein Zertifikate gewaehlt! Abbruch" -ForegroundColor Red
            return
        }
        $cert = (ParseKeychain | Where-Object { $_.Subject -eq $CnOrThumbprint })
    }
    else {
        $cert = SelectCertificate $CnOrThumbprint
        if ($cert -eq $null) {
            Write-Host "Kein gueltiges Zertifikat gefunden! Abbruch" -ForegroundColor Red
            return
        }
    }
    $msg=""
    while($true) {
        Clear-Host
        Write-Host "Subject    : " -NoNewline
        Write-Host $cert.Subject -ForegroundColor Yellow
        Write-Host "Thumbprint : " -NoNewline
        Write-Host $cert.Thumbprint -ForegroundColor Yellow
        Write-Host
        Write-Host "V" -ForegroundColor Yellow -NoNewline
        Write-Host ". Verschluesseln mit diesem Zertifikat"
        Write-Host "E" -ForegroundColor Yellow -NoNewline
        Write-Host ". Exportieren des Zertifikats"
        Write-Host "L" -ForegroundColor Yellow -NoNewline
        Write-Host ". Loeschen des Zertifikats"
        Write-Host "S" -ForegroundColor Yellow -NoNewline
        Write-Host ". Subject (Name) kopieren"
        Write-Host "T" -ForegroundColor Yellow -NoNewline
        Write-Host ". Thumbprint (Kennung) kopieren"
        Write-Host "X" -ForegroundColor Yellow -NoNewline
        Write-Host ". Name / Kennungs Text kopieren"
        Write-Host
        if ($msg -ne "") {
            Write-Host $msg -ForegroundColor Red
            $msg=""
        }
        $e = Read-Host "VELST oder X; Enter = Beenden"
        if ($e -eq "") {
            break
        }
        if ($e -eq "V") {
            Sup-Encrpyt -CnOrThumbprint $cert.Subject
        }
        elseif ($e -eq "E") {
            Sup-ExportCertificate -CnOrThumbprint $cert.Subject
        }
        elseif ($e -eq "L") {
            Sup-RemoveCertificate -CnOrThumbprint $cert.Subject
        }
        elseif ($e -eq "S") {
            $msg="$($cert.Subject) kopiert"
            Set-Clipboard $cert.Subject
        }
        elseif ($e -eq "T") {
            $msg="$($cert.Thumbprint) kopiert"
            Set-Clipboard $cert.Thumbprint
        }
        elseif ($e -eq "X") {
            Write-Host "Verschluesselt mit $($cert.Subject) ($($cert.Thumbprint)) kopiert"
            Set-Clipboard "Verschluesselt mit $($cert.Subject) ($($cert.Thumbprint))."
        }
        else {
            $msg = "Ungueltige Eingabe"
        }
    }
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
    
    if ($IsMacOS) {
        $list = ParseKeychain
        foreach($item in $list) {
            if ($item.Subject -match "^$($wordToComplete).*") {
                New-Object System.Management.Automation.CompletionResult (
                    $item.Subject,
                    $item.Subject,
                    "ParameterValue",
                    $item.Thumbprint
                )
            }
        }

    }
    else {
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
}
function GetZertifikateFromSubjectOrCN {
    param(
        $subjectOrCn = ""
    )
 
    if (!$IsMacOS) 
    {
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
}

function SelectCertificate {
    param(
        $CnOrThumbprint = ""
    )
    if ($IsMacOS) {
        $possible = ParseKeychain
        if ($CnOrThumbprint -ne "") {
            $possible = $possible | Where-Object { $_.Subject -match $CnOrThumbprint -or ($_.Thumbprint -match $CnOrThumbprint)}
        }
        if ($possible.Count -eq 0) {
            $possible = ParseKeychain
        }
        if ($possible.Count -eq 0) {
            return $null
        }
    }
    else {
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
    }

    if ($possible.Count -eq 1) {
        if ($IsMacOS) {
            return $possible[0].Subject
        }
        else {
            return $possible[0]
        }
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
            $msg=""
        }
        Write-Host "`nNr. eingeben, Text fuer Filterung, keine Eingabe fuer Abbruch: " -NoNewline
        $e = Read-Host 
        $nr = -1
        if ([int]::TryParse($e, [Ref] $nr)) {
            if ($nr -gt 0 -and ($nr -le $possible.Count)) {
                if ($IsMacOS) {
                    return $possible[$nr - 1].Subject
                }
                else {
                    return $possible[$nr - 1]
                }
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
                    if ($IsMacOS) {
                        return $possible[0].Subject
                    }
                    else {
                        return $possible[0]
                    }
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
            $msg=""
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
            $msg=""
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

function CopyStringPart {
    param(
        $text = ""
    )
    if ($text -eq "") {
        $text = (Get-Clipboard -Raw)
    }

    $list = @()
    $arr = $text.Split("`n")
    $maxL = -1
    foreach($z in $arr) {
        if ($z.Contains(":")) {
            $dp = $z.IndexOf(":")
            $pw = $z.SubString($dp+1).Trim()
            if ($pw.Length -gt $maxL) {
                $maxL = $pw.Length
            }
            $desc = $z.SubString(0, $dp).Trim()
            $item = [PSCustomObject]@{
                Desc = $desc
                Pass = $pw
            }
            $list += $item
        }
        else {
            if ($z.Trim() -ne "") {
                $item = [PSCustomObject]@{
                    Desc = ""
                    Pass = $z.Trim()
                }
                $list += $item
            }
        }
    }
    if ($maxL -gt -1) {
        $maxL += 2
    }
    
    if ($list.Count -eq 1) {
        Set-Clipboard $list[0].Pass
        Write-Host $list[0].Pass -NoNewline -ForegroundColor Red
        Write-Host " ($($list[0].Desc))" -NoNewline -ForegroundColor Blue
        Write-Host " in die Zwischenablage kopiert."
        Write-Host "Enter um diese wieder zu leeren (- um Zwischenableg nicht zu leeren): "
        $e = Read-Host
        if ($e -ne "-") {
            Set-Clipboard "-"
        }
        return
    }
    $nr = 0
    $msg=""
    while($true) {
        Clear-Host
        Write-Host "Alles (" -NoNewline
        Write-Host "*" -NoNewline -ForegroundColor Yellow
        Write-Host ")"
        Write-Host $text -ForegroundColor Blue
        Write-Host "--------------------------------------------------------------------------------"
        for($nr = 0; $nr -lt $list.Count; $nr++)
        {
            $pw = $list[$nr].Pass
            if ($MaxL -gt -1 -and $list[$nr].Desc -ne "") {
                $pw += (" "*($MaxL - $pw.Length))
                Write-Host "$($nr + 1)" -ForegroundColor Yellow -NoNewline
                Write-Host ". $pw" -NoNewline
                Write-Host " ($($list[$nr].Desc))" -ForegroundColor Blue
            }
            else {
                Write-Host "$($nr + 1)" -ForegroundColor Yellow -NoNewline
                Write-Host ". $pw" 
            }
        }
        Write-Host
        Write-Host "Zum kopieren: " -NoNewline
        Write-Host "Nr." -NoNewline -ForegroundColor Yellow
        Write-Host " eingeben oder " -NoNewline
        Write-Host "*" -NoNewline -ForegroundColor Yellow
        Write-Host " für alles"
        Write-Host "Keine Eingabe: Auswahl für kopieren beenden (und Zwischenablage leeren)"
        Write-Host "-" -ForegroundColor Yellow -NoNewline
        Write-Host " um Zwischenablage nicht zu leeren"
        Write-Host
        if ($msg -ne "") {
            Write-Host $msg -ForegroundColor Red
        }
        $msg=""
        $e = Read-Host "Nr, * oder keine Eingabe"
        if ($e -eq "") {
            Set-Clipboard "-"
            return
        }
        if ($e -eq "-") {
            return
        }
        $nr = -1
        if ([int]::TryParse($e, [Ref] $nr)) {
            if ($nr -gt 0 -and $nr -le $list.Count) {
                Set-Clipboard $list[$nr-1].Pass
                $msg = "$($list[$nr-1].Pass) in Zwischenablage kopiert"
            }
            else {
                $msg = "Ungültige Nummer"
            }
        }
        elseif ($e -eq "*") {
            Set-Clipboard $text
            $msg = "Kompletter Text in Zwischenablage kopiert"
        }
        else {
            $msg = "Ungültige Eingabe!"
        }
    }
}

$symbols = '?.-_!#*'.ToCharArray()
$characterList = 'a'..'z' + 'A'..'Z' + '0'..'9' + $symbols

function GeneratePassword {
    param(
        [ValidateRange(12, 256)]
        [int] 
        $length = 18
    )

    do {
        $password = -join (0..$length | % { $characterList | Get-Random })
        [int]$hasLowerChar = $password -cmatch '[a-z]'
        [int]$hasUpperChar = $password -cmatch '[A-Z]'
        [int]$hasDigit = $password -match '[0-9]'
        [int]$hasSymbol = $password.IndexOfAny($symbols) -ne -1

    }
    until (($hasLowerChar + $hasUpperChar + $hasDigit + $hasSymbol) -ge 3)

    $password 
}
