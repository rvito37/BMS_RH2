$RH2Dir = "C:\Users\AVXUser\BMS\RH2"
$OutFile = "C:\Users\AVXUser\BMS\RH2_PARSED.txt"

$allFiles = @()
$allFiles += Get-ChildItem -Path $RH2Dir -Filter "*.RH2" -ErrorAction SilentlyContinue
$allFiles += Get-ChildItem -Path $RH2Dir -Filter "*.rh2" -ErrorAction SilentlyContinue
$allFiles = $allFiles | Sort-Object Name -Unique

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("=" * 80)
[void]$sb.AppendLine("RH2 REPORT FILES ANALYSIS")
[void]$sb.AppendLine("Total files: $($allFiles.Count)")
[void]$sb.AppendLine("=" * 80)

foreach ($file in $allFiles) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

    # Fast string extraction: scan bytes for printable runs >= 3
    $strings = New-Object System.Collections.ArrayList
    $current = New-Object System.Text.StringBuilder
    foreach ($b in $bytes) {
        if ($b -ge 32 -and $b -le 126) {
            [void]$current.Append([char]$b)
        } else {
            if ($current.Length -ge 3) {
                [void]$strings.Add($current.ToString().Trim())
            }
            $current.Clear() | Out-Null
        }
    }
    if ($current.Length -ge 3) {
        [void]$strings.Add($current.ToString().Trim())
    }

    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("=" * 80)
    [void]$sb.AppendLine("FILE: $($file.Name)  (Size: $($bytes.Length) bytes)")
    [void]$sb.AppendLine("-" * 80)

    $counter = 0
    foreach ($s in $strings) {
        if ($s.Length -ge 3) {
            [void]$sb.AppendLine("  [$counter] $s")
            $counter++
        }
    }

    Write-Host "Parsed: $($file.Name) ($counter strings)"
}

[System.IO.File]::WriteAllText($OutFile, $sb.ToString(), [System.Text.Encoding]::UTF8)
Write-Host "Done! Output: $OutFile"
