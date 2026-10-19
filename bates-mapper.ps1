<#
.SYNOPSIS
Replaces Bates numbers in a .docx with new Bates numbers defined in an .xlsx (col A = old, col B = new).
The .docx is saved only if at least one replacement is made; the .xlsx is never modified. Requires Excel and Word (COM).
#>

# Global Bates format (regex). Default: ABC.001.001.0001
$Global:BatesPattern = '[A-Z]{3}\.\d{3}\.\d{3}\.\d{4}'

# How many columns (per row) and rows to scan to find the first Bates id in the spreadsheet
$Global:traverse_columns = 5
$Global:traverse_rows = 10

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
            # One bulk read covering the scan area plus the column after it (the new number)
            $data = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($lastRow, $Global:traverse_columns + 1)).Value2

            # Scan each row across columns 1..traverse_columns; the first hit sets the start row and old-number column
            $startRow = 0; $oldCol = 0
            $scanRows = [Math]::Min($Global:traverse_rows, $lastRow)
            :scan for ($r = 1; $r -le $scanRows; $r++) {
                for ($c = 1; $c -le $Global:traverse_columns; $c++) {
                    if ("$($data[$r, $c])".Trim() -match $exact) { $startRow = $r; $oldCol = $c; break scan }
                }
            }
            if ($startRow -eq 0) {
                Write-Host "`nNo Bates number found in the first $scanRows row(s) x $($Global:traverse_columns) column(s) of the spreadsheet." -ForegroundColor Yellow
            }
            # No Bates numbers in the next column: it is only the authoritative list of correct numbers
            $listOnly = $false
            if ($startRow -gt 0) {
                $listOnly = $true
                for ($r = $startRow; $r -le $lastRow; $r++) {
                    if ("$($data[$r, ($oldCol + 1)])".Trim() -match $exact) { $listOnly = $false; break }
                }
                if ($listOnly) { Write-Host "`nNo second column of Bates numbers found; treating column $oldCol as the list of correct numbers." -ForegroundColor Cyan }
            }
            for ($r = $startRow; $startRow -gt 0 -and $r -le $lastRow; $r++) {
                $a = "$($data[$r, $oldCol])".Trim()
                $b = "$($data[$r, ($oldCol + 1)])".Trim()
                if ($listOnly) { $b = $a }
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

    Write-Host "`nLoaded $($map.Count) mapping(s)." -ForegroundColor Cyan
    if ($dupes -gt 0) { Write-Host "  $dupes duplicate source number(s) ignored (first occurrence kept)." -ForegroundColor Yellow }
    if ($missing.Count -gt 0) {
        Write-Host "  $($missing.Count) source number(s) have no new number (not mapped):" -ForegroundColor Yellow
        $missing | ForEach-Object { Write-Host "    $_" -ForegroundColor Yellow }
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

Write-Host "`nFuzzy mode: unmapped numbers use the nearest lower mapped number in the same prefix."
$fuzzy = (Read-Host "Enable fuzzy mode? (Y = yes, anything else = no)") -cmatch '^[Yy]$'

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
# Numbers already in column B are final and must never be touched again
$targets = [System.Collections.Generic.HashSet[string]]::new([string[]]@($map.Values))
$total = 0; $changed = 0; $unchanged = 0; $noMap = 0
$unmapped = [System.Collections.Generic.HashSet[string]]::new()

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $null
try {
    Write-Host ("`nProcessing {0} (fuzzy mode {1})`n" -f $docx.Name, $(if ($fuzzy) { 'on' } else { 'off' }))
    Write-Host ('  {0,-18}    {1,-18}  {2}' -f 'FOUND', 'REPLACED WITH', 'NOTE') -ForegroundColor DarkGray
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
            if ($fuzzy -and -not $map.ContainsKey($old) -and -not $targets.Contains($old)) { $fuzzyKey = Find-FuzzyKey $old }
            if ($targets.Contains($old)) {
                $unchanged++
                Write-Host ('  {0,-18} -> {1,-18}  {2}' -f $old, $old, 'unchanged') -ForegroundColor DarkGray
            }
            elseif ($fuzzyKey) {
                $fuzzyCount++
                $newValue = $map[$fuzzyKey]
                Write-Host ('  {0,-18} -> {1,-18}  {2}' -f $old, $newValue, 'fuzzy') -ForegroundColor Cyan
            }
            elseif (-not $map.ContainsKey($old)) {
                $noMap++; [void]$unmapped.Add($old)
                Write-Host ('  {0,-18} -> {1,-18}  {2}' -f $old, '-', 'no mapping, skipped') -ForegroundColor Yellow
            }
            elseif ($map[$old] -eq $old) {
                $unchanged++
                Write-Host ('  {0,-18} -> {1,-18}  {2}' -f $old, $map[$old], 'unchanged') -ForegroundColor DarkGray
            }
            else {
                $changed++
                $newValue = $map[$old]
                Write-Host ('  {0,-18} -> {1,-18}' -f $old, $newValue) -ForegroundColor Green
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

Write-Host "`nSummary" -ForegroundColor Cyan
Write-Host ('  {0,-22}{1,4}' -f 'Found', $total)
Write-Host ('  {0,-22}{1,4}' -f 'Changed', $changed)
if ($fuzzy) { Write-Host ('  {0,-22}{1,4}' -f 'Fuzzy matched', $fuzzyCount) }
Write-Host ('  {0,-22}{1,4}' -f 'Unchanged', $unchanged)
Write-Host ('  {0,-22}{1,4}' -f 'Skipped (no mapping)', $noMap)
Write-Host ""