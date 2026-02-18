param(
    [string]$FilePath = "C:\Users\AVXUser\BMS\RH2\RPQC01V4.RH2"
)

$bytes = [System.IO.File]::ReadAllBytes($FilePath)
$fname = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)

# Extract printable strings with positions
$strings = New-Object System.Collections.ArrayList
$positions = New-Object System.Collections.ArrayList
$cur = New-Object System.Text.StringBuilder
$startPos = 0

for ($i = 0; $i -lt $bytes.Length; $i++) {
    $b = $bytes[$i]
    if ($b -ge 32 -and $b -le 126) {
        if ($cur.Length -eq 0) { $startPos = $i }
        [void]$cur.Append([char]$b)
    } else {
        if ($cur.Length -ge 3) {
            [void]$strings.Add($cur.ToString())
            [void]$positions.Add($startPos)
        }
        [void]$cur.Clear()
    }
}
if ($cur.Length -ge 3) {
    [void]$strings.Add($cur.ToString())
    [void]$positions.Add($startPos)
}

Write-Host "================================================================"
Write-Host "FILE: $fname  ($($bytes.Length) bytes, $($strings.Count) strings)"
Write-Host "================================================================"

# Categorize
$tables = @{}
$indexes = @{}
$fieldRefs = @{}
$formulas = @{}
$drivers = @{}
$paths = @{}

for ($j = 0; $j -lt $strings.Count; $j++) {
    $s = $strings[$j]
    if ($s -match '(?i)([A-Z_][A-Z0-9_]+\.DBF)') { $tables[$Matches[1]] = $true }
    if ($s -match '(?i)([A-Z_][A-Z0-9_]+\.CDX)') { $indexes[$Matches[1]] = $true }
    if ($s -match '->') { $fieldRefs[$s] = $true }
    if ($s -match '(?i)DBFCDX') { $drivers[$s] = $true }
    if ($s -match '(?i)[A-Z]:\\') { $paths[$s] = $true }
}

Write-Host ""
Write-Host "--- TABLES (DBF) ---"
$tables.Keys | Sort-Object | ForEach-Object { Write-Host "  $_" }

Write-Host ""
Write-Host "--- INDEXES (CDX) ---"
$indexes.Keys | Sort-Object | ForEach-Object { Write-Host "  $_" }

Write-Host ""
Write-Host "--- FIELD REFERENCES (alias->field) ---"
$fieldRefs.Keys | Sort-Object | ForEach-Object { Write-Host "  $_" }

Write-Host ""
Write-Host "--- DRIVERS ---"
$drivers.Keys | Sort-Object | ForEach-Object { Write-Host "  $_" }

Write-Host ""
Write-Host "--- FILE PATHS ---"
$paths.Keys | Sort-Object | ForEach-Object { Write-Host "  $_" }

Write-Host ""
Write-Host "--- ALL STRINGS (offset | content) ---"
for ($j = 0; $j -lt $strings.Count; $j++) {
    $s = $strings[$j]
    $pos = $positions[$j]
    $hexPos = $pos.ToString("X6")
    Write-Host "  [0x$hexPos] $s"
}
