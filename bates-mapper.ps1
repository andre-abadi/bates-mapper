<#
.SYNOPSIS
Replaces Bates numbers in a .docx with new Bates numbers defined in an .xlsx (col A = old, col B = new).
The .docx is saved only if at least one replacement is made; the .xlsx is never modified. Requires Excel and Word (COM).
#>

# Global Bates format (regex). Default: ABC.001.001.0001
$Global:BatesPattern = '[A-Z]{3}\.\d{3}\.\d{3}\.\d{4}'

function Select-FileByExtension {
    param([string]$Extension, [string]$Prompt)
    $files = @(Get-ChildItem -LiteralPath (Get-Location).Path -Filter "*.$Extension" -File |
        Where-Object { -not $_.Name.StartsWith('~$') } | Sort-Object Name)
    if ($files.Count -eq 0) {
        Write-Host "No .$Extension files found in $(Get-Location). Exiting."
        return $null
    }
    Write-Host "`n$Prompt"
    for ($i = 0; $i -lt $files.Count; $i++) { Write-Host ("  [{0}] {1}" -f ($i + 1), $files[$i].Name) }
    while ($true) {
        $answer = Read-Host "Enter number (1-$($files.Count)), or leave blank to quit"
        if ([string]::IsNullOrWhiteSpace($answer)) { return $null }
        $n = 0
        if ([int]::TryParse($answer, [ref]$n) -and $n -ge 1 -and $n -le $files.Count) { return $files[$n - 1] }
        Write-Host "Invalid selection."
    }
}

function Import-BatesMap([string]$Path) {
    $map = @{}
    $exact = "^$($Global:BatesPattern)$"
    $dupes = 0
    $missing = [System.Collections.Generic.List[string]]::new()

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $null
    try {
        $wb = $excel.Workbooks.Open($Path, 0, $true)
        $ws = $wb.Worksheets.Item(1)
        $lastRow = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
        if ($lastRow -ge 1) {
            # One bulk read of columns A:B into a 2D array
            $data = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($lastRow, 2)).Value2
            for ($r = 1; $r -le $lastRow; $r++) {
                $a = "$($data[$r, 1])".Trim()
                $b = "$($data[$r, 2])".Trim()
                if ($a -eq '' -and $b -eq '') { continue }
                if ($a -notmatch $exact) { continue }
                if ($b -eq '') { $missing.Add("row ${r}: $a"); continue }
                if ($map.ContainsKey($a)) { $dupes++; continue }
                $map[$a] = $b
            }
        }
    }
    finally {
        if ($wb) { $wb.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    Write-Host "`nLoaded $($map.Count) mapping(s)."
    if ($dupes -gt 0) { Write-Warning "$dupes duplicate source number(s) ignored (first occurrence kept)." }
    if ($missing.Count -gt 0) {
        Write-Warning "$($missing.Count) source number(s) have no matching second column (proceeding anyway):"
        $missing | ForEach-Object { Write-Warning "  $_" }
    }
    return $map
}

# Emits every story range (body, footnotes, endnotes, headers, footers, text boxes...)
function Get-DocxStoryRanges($Doc) {
    foreach ($story in $Doc.StoryRanges) {
        $range = $story
        while ($range) {
            $range
            $range = $range.NextStoryRange
        }
    }
}

# --- Main ---
$xlsx = Select-FileByExtension -Extension 'xlsx' -Prompt 'Select the Excel mapping file:'
if (-not $xlsx) { return }
$map = Import-BatesMap $xlsx.FullName

$docx = Select-FileByExtension -Extension 'docx' -Prompt 'Select the Word document:'
if (-not $docx) { return }

$fuzzy = (Read-Host "`nEnable fuzzy mode? Unmapped numbers fall back to the next-lowest mapped number in the same prefix (Y = yes, anything else = no)") -cmatch '^[Yy]$'
Write-Host ("Fuzzy mode: " + $(if ($fuzzy) { 'ON' } else { 'OFF' }))

# Sorted mapped page numbers per prefix (everything before the last 4 digits)
$byPrefix = @{}
if ($fuzzy) {
    foreach ($k in $map.Keys) {
        $p = $k.Substring(0, $k.Length - 4)
        if (-not $byPrefix.ContainsKey($p)) { $byPrefix[$p] = [System.Collections.Generic.List[int]]::new() }
        $byPrefix[$p].Add([int]$k.Substring($k.Length - 4))
    }
    foreach ($p in @($byPrefix.Keys)) { $byPrefix[$p].Sort() }
}

function Find-FuzzyKey([string]$Bates) {
    $p = $Bates.Substring(0, $Bates.Length - 4)
    $n = [int]$Bates.Substring($Bates.Length - 4)
    if (-not $byPrefix.ContainsKey($p)) { return $null }
    $best = $null
    foreach ($v in $byPrefix[$p]) { if ($v -lt $n) { $best = $v } else { break } }
    if ($null -eq $best) { return $null }
    return $p + $best.ToString('0000')
}

$fuzzyCount = 0
$exactRegex = [regex]"^$($Global:BatesPattern)$"
$wildcard = $Global:BatesPattern.Replace('\d', '[0-9]').Replace('\.', '.')  # Word wildcard syntax
$replaced = 0
$total = 0; $changed = 0; $unchanged = 0; $noMap = 0
$unmapped = [System.Collections.Generic.HashSet[string]]::new()

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $null
try {
    Write-Host "`nProcessing $($docx.Name)...`n"
    $doc = $word.Documents.Open($docx.FullName, $false, $false)
    foreach ($story in Get-DocxStoryRanges $doc) {
        $rng = $story.Duplicate
        $find = $rng.Find
        $find.ClearFormatting()
        $find.Text = $wildcard
        $find.MatchWildcards = $true
        $find.Forward = $true
        $find.Wrap = 0
        while ($find.Execute()) {
            $old = $rng.Text
            $before = $rng.Duplicate; $before.Collapse(1); [void]$before.MoveStart(1, -1)
            $after = $rng.Duplicate; $after.Collapse(0); [void]$after.MoveEnd(1, 1)
            $isBates = $exactRegex.IsMatch($old) -and ($before.Text -notmatch '[A-Za-z0-9]$') -and ($after.Text -notmatch '^[A-Za-z0-9]')
            if (-not $isBates) { $rng.Collapse(0); continue }
            $total++
            $newValue = $null
            $fuzzyKey = $null
            if ($fuzzy -and -not $map.ContainsKey($old)) { $fuzzyKey = Find-FuzzyKey $old }
            if ($fuzzyKey) {
                $fuzzyCount++
                $newValue = $map[$fuzzyKey]
                Write-Host "$old  ->  $newValue (fuzzy)" -ForegroundColor Cyan
            }
            elseif (-not $map.ContainsKey($old)) {
                $noMap++; [void]$unmapped.Add($old)
                Write-Host "$old  ->  (no mapping, skipped)" -ForegroundColor Yellow
            }
            elseif ($map[$old] -eq $old) {
                $unchanged++
                Write-Host "$old  ->  $($map[$old]) (unchanged)" -ForegroundColor DarkGray
            }
            else {
                $changed++
                $newValue = $map[$old]
                Write-Host "$old  ->  $newValue" -ForegroundColor Green
            }
            if ($newValue -and $newValue -ne $old) {
                $rng.Text = $newValue
                $replaced++
            }
            $rng.Collapse(0)
        }
    }
    if ($replaced -gt 0) {
        $doc.Save()
        Write-Host "`nSaved $($docx.Name) with $replaced replacement(s)." -ForegroundColor Green
    }
    else {
        Write-Host "`nNo changes made; $($docx.Name) not saved."
    }
}
finally {
    if ($doc) { $doc.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
    $word.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}

Write-Host "`n===== Summary ====="
Write-Host "Bates numbers found:        $total"
Write-Host "Mapped (changed):           $changed"
Write-Host "Mapped (unchanged):         $unchanged"
if ($fuzzy) { Write-Host "Mapped (fuzzy):             $fuzzyCount" }
Write-Host "Skipped (no mapping):       $noMap"
Write-Host ""