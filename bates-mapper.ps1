<#
.SYNOPSIS
Maps Bates numbers found in a .docx to new Bates numbers defined in an .xlsx (col A = old, col B = new).
Read-only: neither file is modified. Requires Excel and Word (COM).
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

# Emits the text of every story (body, footnotes, endnotes, headers, footers, text boxes...)
function Get-DocxStoryTexts($Doc) {
    foreach ($story in $Doc.StoryRanges) {
        $range = $story
        while ($range) {
            $range.Text
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

$regex = [regex]"(?<![A-Za-z0-9])$($Global:BatesPattern)(?![A-Za-z0-9])"
$total = 0; $changed = 0; $unchanged = 0; $noMap = 0
$unmapped = [System.Collections.Generic.HashSet[string]]::new()

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $null
try {
    Write-Host "`nProcessing $($docx.Name)...`n"
    $doc = $word.Documents.Open($docx.FullName, $false, $true)
    foreach ($text in Get-DocxStoryTexts $doc) {
        foreach ($m in $regex.Matches($text)) {
            $total++
            $old = $m.Value
            if (-not $map.ContainsKey($old)) {
                $noMap++; [void]$unmapped.Add($old)
                Write-Host "$old  ->  (no mapping, skipped)" -ForegroundColor Yellow
            }
            elseif ($map[$old] -eq $old) {
                $unchanged++
                Write-Host "$old  ->  $($map[$old]) (unchanged)" -ForegroundColor DarkGray
            }
            else {
                $changed++
                Write-Host "$old  ->  $($map[$old])" -ForegroundColor Green
            }
        }
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
Write-Host "Skipped (no mapping):       $noMap"
Write-Host ""