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