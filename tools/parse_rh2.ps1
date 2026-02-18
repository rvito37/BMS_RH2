param(
    [string]$RH2Dir = "C:\Users\AVXUser\BMS\RH2",
    [string]$OutFile = "C:\Users\AVXUser\BMS\RH2_PARSED.txt"
)

$allFiles = Get-ChildItem -Path $RH2Dir -Filter "*.RH2" | Sort-Object Name
$allFiles += Get-ChildItem -Path $RH2Dir -Filter "*.rh2" | Sort-Object Name
$allFiles = $allFiles | Sort-Object Name -Unique

$output = @()
$output += "=" * 80
$output += "RH2 REPORT FILES ANALYSIS"
$output += "Generated: $(Get-Date)"
$output += "Total files: $($allFiles.Count)"
$output += "=" * 80

foreach ($file in $allFiles) {
    $data = [System.IO.File]::ReadAllBytes($file.FullName)
    $text = [System.Text.Encoding]::ASCII.GetString($data)

    # Extract all printable strings >= 3 chars
    $strings = [regex]::Matches($text, '[\x20-\x7E]{3,}')
    $strList = @()
    foreach ($m in $strings) {
        $strList += $m.Value.Trim()
    }

    $output += ""
    $output += "=" * 80
    $output += "FILE: $($file.Name)  (Size: $($data.Length) bytes)"
    $output += "=" * 80

    # Find DBF table references
    $dbfRefs = @()
    foreach ($s in $strList) {
        if ($s -match '(?i)\\([A-Z_][A-Z0-9_]*\.DBF)' -or $s -match '(?i)([A-Z_][A-Z0-9_]+\.DBF)') {
            $dbfRefs += $Matches[1]
        }
    }
    $dbfRefs = $dbfRefs | Sort-Object -Unique
    if ($dbfRefs.Count -gt 0) {
        $output += ""
        $output += "  TABLES (DBF references):"
        foreach ($d in $dbfRefs) { $output += "    - $d" }
    }

    # Find CDX/NTX index references
    $idxRefs = @()
    foreach ($s in $strList) {
        if ($s -match '(?i)([A-Z_][A-Z0-9_]+\.CDX)') { $idxRefs += $Matches[1] }
        if ($s -match '(?i)([A-Z_][A-Z0-9_]+\.NTX)') { $idxRefs += $Matches[1] }
    }
    $idxRefs = $idxRefs | Sort-Object -Unique
    if ($idxRefs.Count -gt 0) {
        $output += ""
        $output += "  INDEXES:"
        foreach ($d in $idxRefs) { $output += "    - $d" }
    }

    # Find field names (patterns like FIELD->name or alias->name)
    $fieldRefs = @()
    foreach ($s in $strList) {
        $fmatches = [regex]::Matches($s, '(?i)([A-Z_][A-Z0-9_]*)->([A-Z_][A-Z0-9_]*)')
        foreach ($fm in $fmatches) {
            $fieldRefs += "$($fm.Groups[1].Value)->$($fm.Groups[2].Value)"
        }
    }
    $fieldRefs = $fieldRefs | Sort-Object -Unique
    if ($fieldRefs.Count -gt 0) {
        $output += ""
        $output += "  FIELD REFERENCES (alias->field):"
        foreach ($d in $fieldRefs) { $output += "    - $d" }
    }

    # Find function calls
    $funcRefs = @()
    foreach ($s in $strList) {
        $fmatches = [regex]::Matches($s, '(?i)([A-Z_][A-Za-z0-9_]+)\(')
        foreach ($fm in $fmatches) {
            $fname = $fm.Groups[1].Value
            if ($fname -notmatch '(?i)^(DBFCDX|DBFCDXAX|CDX|DBF|NTX|FOR|AND|NOT|STR|VAL|IIF|IF|ELSE|END|DO|WHILE|CASE|PADR|PADL|PADC|TRIM|ALLTRIM|LTRIM|RTRIM|UPPER|LOWER|SUBSTR|LEFT|RIGHT|LEN|SPACE|DTOC|DTOS|CTOD|DATE|TIME|YEAR|MONTH|DAY|TRANSFORM|CHR|ASC|AT|RAT|REPLICATE|STRTRAN|EMPTY|TYPE|VALTYPE|SELECT|ALIAS|USED|FOUND|EOF|BOF|RECNO|LASTREC|FCOUNT|DELETED|FIELDNAME|FIELDGET|FIELDPUT|INT|ROUND|MOD|ABS|MAX|MIN|SQRT|EXP|LOG|CEILING|FLOOR)$') {
                $funcRefs += $fname + "()"
            }
        }
    }
    $funcRefs = $funcRefs | Sort-Object -Unique
    if ($funcRefs.Count -gt 0) {
        $output += ""
        $output += "  FUNCTIONS/UDFs:"
        foreach ($d in $funcRefs) { $output += "    - $d" }
    }

    # Find report title strings (look for AVX, Report, common title words)
    $titles = @()
    foreach ($s in $strList) {
        if ($s -match '(?i)(AVX|Report|Summary|List|Production|Quality|Marketing|Batch|Order|Control|Design|Engineering|FGS|Accounting)' -and $s.Length -ge 5 -and $s.Length -le 120) {
            $titles += $s
        }
    }
    $titles = $titles | Sort-Object -Unique
    if ($titles.Count -gt 0) {
        $output += ""
        $output += "  TITLE/LABEL STRINGS:"
        foreach ($d in $titles) { $output += "    - $d" }
    }

    # Find driver references
    $drivers = @()
    foreach ($s in $strList) {
        if ($s -match '(?i)(DBFCDXAX|DBFCDX|DBFNTX)') {
            $drivers += $Matches[1]
        }
    }
    $drivers = $drivers | Sort-Object -Unique
    if ($drivers.Count -gt 0) {
        $output += ""
        $output += "  DATABASE DRIVERS:"
        foreach ($d in $drivers) { $output += "    - $d" }
    }

    # Find path references
    $paths = @()
    foreach ($s in $strList) {
        if ($s -match '(?i)([A-Z]:\\[^\x00-\x1F]{3,})') {
            $paths += $Matches[1]
        }
    }
    $paths = $paths | Sort-Object -Unique
    if ($paths.Count -gt 0) {
        $output += ""
        $output += "  FILE PATHS:"
        foreach ($d in $paths) { $output += "    - $d" }
    }

    # Find sort/index key expressions
    $sortKeys = @()
    foreach ($s in $strList) {
        if ($s -match '\+' -and $s -match '(?i)[a-z_]' -and $s.Length -ge 8 -and $s.Length -le 200 -and $s -notmatch '(?i)(DBF|CDX|NTX|\\|Report|AVX|Page|Date|Time|Total|Count|Sum|print)') {
            $sortKeys += $s
        }
    }
    if ($sortKeys.Count -gt 0) {
        $output += ""
        $output += "  POSSIBLE INDEX/SORT EXPRESSIONS:"
        foreach ($d in $sortKeys) { $output += "    - $d" }
    }

    # Dump ALL meaningful strings (skip very short or gibberish)
    $output += ""
    $output += "  ALL EXTRACTED STRINGS:"
    $counter = 0
    foreach ($s in $strList) {
        $trimmed = $s.Trim()
        if ($trimmed.Length -ge 3) {
            $output += "    [$counter] $trimmed"
            $counter++
        }
    }
}

# Write output
$output | Out-File -FilePath $OutFile -Encoding UTF8
Write-Host "Done! Parsed $($allFiles.Count) files. Output: $OutFile"
