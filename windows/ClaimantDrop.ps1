param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Paths
)

Add-Type -AssemblyName System.Windows.Forms

function Show-Message([string]$Text, [string]$Title, [System.Windows.Forms.MessageBoxIcon]$Icon) {
    # CLAIMANTS_NO_DIALOG=1 prints messages instead, for automated checks.
    if ($env:CLAIMANTS_NO_DIALOG -eq '1') {
        if ($Icon -eq [System.Windows.Forms.MessageBoxIcon]::Information) { [Console]::Out.WriteLine($Text) }
        else { [Console]::Error.WriteLine($Text) }
        return
    }
    [void][System.Windows.Forms.MessageBox]::Show(
        $Text,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $Icon
    )
}

function Csv-Escape([string]$Value) {
    if ($null -eq $Value) { return '' }
    if ($Value.Contains(',') -or $Value.Contains('"') -or $Value.Contains("`r") -or $Value.Contains("`n")) {
        return '"' + $Value.Replace('"', '""') + '"'
    }
    return $Value
}

function Get-ColumnNumber([string]$Letters) {
    $text = ([string]$Letters).Trim().ToUpperInvariant()
    if ($text -notmatch '^[A-Z]+$') { return $null }
    $number = 0
    foreach ($character in $text.ToCharArray()) { $number = $number * 26 + ([int]$character - 64) }
    return $number
}

# CLAIMANTS_COLUMNS names a settings file; otherwise claimant-columns.json beside this script is used if present.
function Get-ColumnMapping {
    $mapping = @{
        AnswerLetter = 'E'
        AnswerColumn = 5
        Headers = @('Admin', 'Admin P#', 'Publisher', 'Publisher P#')
        Columns = @(18, 17, 14, 13)
    }
    $settingsPath = $env:CLAIMANTS_COLUMNS
    if ([string]::IsNullOrWhiteSpace($settingsPath)) {
        $settingsPath = Join-Path $PSScriptRoot 'claimant-columns.json'
        if (-not (Test-Path -LiteralPath $settingsPath)) { return $mapping }
    }
    try {
        $settings = Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw "Could not read column settings: $([System.IO.Path]::GetFileName($settingsPath))."
    }
    $answerColumn = Get-ColumnNumber $settings.answerColumn
    if ($null -eq $answerColumn -or -not $settings.fields -or @($settings.fields).Count -eq 0) {
        throw 'Column settings need an answerColumn letter and at least one field.'
    }
    $headers = @()
    $columns = @()
    foreach ($field in @($settings.fields)) {
        $column = Get-ColumnNumber $field.column
        if ([string]::IsNullOrEmpty($field.header) -or $null -eq $column) {
            throw 'Each column setting needs a header and a column letter.'
        }
        $headers += [string]$field.header
        $columns += $column
    }
    $mapping.AnswerLetter = ([string]$settings.answerColumn).Trim().ToUpperInvariant()
    $mapping.AnswerColumn = $answerColumn
    $mapping.Headers = $headers
    $mapping.Columns = $columns
    return $mapping
}

if (-not $Paths -or $Paths.Count -eq 0) {
    Show-Message 'Drag one .xlsx report onto the desktop icon.' 'Claimant Drop' ([System.Windows.Forms.MessageBoxIcon]::Information)
    exit 0
}

$inputPath = [System.IO.Path]::GetFullPath($Paths[0])
if ([System.IO.Path]::GetExtension($inputPath).ToLowerInvariant() -ne '.xlsx') {
    Show-Message 'Please drop an .xlsx report onto the icon.' 'Claimant Drop' ([System.Windows.Forms.MessageBoxIcon]::Warning)
    exit 1
}

$excel = $null
$workbook = $null
$sheet = $null
$range = $null

try {
    $mapping = Get-ColumnMapping
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($inputPath, 0, $true)
    $sheet = $workbook.Worksheets.Item(1)
    $lastRow = $sheet.UsedRange.Row + $sheet.UsedRange.Rows.Count - 1

    if ($lastRow -lt 2) {
        throw 'No report rows were found.'
    }

    $columns = $mapping.Columns
    $headers = $mapping.Headers
    $lastColumn = ($columns + @($mapping.AnswerColumn) | Measure-Object -Maximum).Maximum
    $range = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item($lastRow, $lastColumn))
    $data = $range.Value2

    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $uniqueRows = New-Object 'System.Collections.Generic.List[object]'
    $matchedRows = 0
    $hasNo = $false
    $unknownRows = New-Object 'System.Collections.Generic.List[int]'

    for ($rowNumber = 2; $rowNumber -le $lastRow; $rowNumber++) {
        $answer = ([string]$data[$rowNumber, $mapping.AnswerColumn]).Trim().ToUpperInvariant()
        $answer = $answer.Replace('(', '').Replace(')', '').Trim()
        if ([string]::IsNullOrWhiteSpace($answer)) { continue }
        if ($answer -eq 'NO' -or $answer -eq 'N') { $hasNo = $true; continue }
        if ($answer -ne 'YES' -and $answer -ne 'Y') { [void]$unknownRows.Add($rowNumber); continue }

        $matchedRows++
        $claimant = @()
        foreach ($column in $columns) {
            $claimant += [string]$data[$rowNumber, $column]
        }

        $keyText = ConvertTo-Json -InputObject $claimant -Compress
        $key = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($keyText))
        if ($seen.Add($key)) {
            [void]$uniqueRows.Add($claimant)
        }
    }

    if ($matchedRows -eq 0) {
        throw "No YES, Yes, Y, or (Y) rows were found in column $($mapping.AnswerLetter)."
    }

    $lines = New-Object 'System.Collections.Generic.List[string]'
    [void]$lines.Add(($headers -join "`t"))
    foreach ($claimant in $uniqueRows) {
        [void]$lines.Add(($claimant -join "`t"))
    }
    $clipboardText = $lines -join "`r`n"

    $csvLines = New-Object 'System.Collections.Generic.List[string]'
    [void]$csvLines.Add((($headers | ForEach-Object { Csv-Escape $_ }) -join ','))
    foreach ($claimant in $uniqueRows) {
        [void]$csvLines.Add((($claimant | ForEach-Object { Csv-Escape ([string]$_) }) -join ','))
    }

    $outputName = [System.IO.Path]::GetFileNameWithoutExtension($inputPath) + '_yes_claimants.csv'
    $outputPath = Join-Path ([System.IO.Path]::GetDirectoryName($inputPath)) $outputName
    [System.IO.File]::WriteAllLines($outputPath, $csvLines, (New-Object System.Text.UTF8Encoding($true)))

    if ($env:CLAIMANTS_NO_CLIPBOARD -ne '1') { Set-Clipboard -Value $clipboardText }
    $classification = if ($hasNo) { 'True Partial' } elseif ($unknownRows.Count -eq 0) { 'True Full' } else { 'Needs review' }
    $warning = if ($unknownRows.Count -gt 0) { "`r`nUnrecognized transfer answers in rows: $($unknownRows -join ', ')." } else { '' }

    $plural = if ($uniqueRows.Count -eq 1) { 'row' } else { 'rows' }
    Show-Message ("$classification.$warning`r`n`r`nFound $matchedRows matching rows and $($uniqueRows.Count) unique claimant $plural.`r`n`r`nThe table is copied—press Ctrl+V to paste it.`r`n`r`nAlso saved:`r`n$outputPath") 'Claimant Drop' ([System.Windows.Forms.MessageBoxIcon]::Information)
}
catch {
    Show-Message ("The report could not be processed:`r`n`r`n$($_.Exception.Message)") 'Claimant Drop' ([System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}
finally {
    if ($workbook -ne $null) { $workbook.Close($false) }
    if ($excel -ne $null) { $excel.Quit() }

    foreach ($comObject in @($range, $sheet, $workbook, $excel)) {
        if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
