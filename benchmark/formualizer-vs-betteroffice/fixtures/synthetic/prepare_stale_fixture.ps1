param(
    [string]$FixtureDir = $PSScriptRoot,
    [string]$SourcePath = (Join-Path $PSScriptRoot 'formula-suite.xlsx'),
    [string]$StaleInputPath = (Join-Path $PSScriptRoot 'formula-suite-stale-cache.xlsx'),
    [string]$SpecPath = (Join-Path $PSScriptRoot 'cases.json'),
    [string]$StaleSpecPath = (Join-Path $PSScriptRoot 'cases-stale.json'),
    [int]$OriginalCachedValue = 10,
    [int]$StaleInputValue = 25,
    [int]$NextInputValue = 35,
    [string]$ChangedOraclePath = (Join-Path $PSScriptRoot 'excel-com-input35-oracle.json')
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$source = $SourcePath
$stale = $StaleInputPath

# Build a reference for a new input value (35) while leaving the original fixture untouched.
$excel = $null
$book = $null
try {
    $spec = Get-Content -LiteralPath $specPath -Raw | ConvertFrom-Json
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($source, 0, $true)
    $book.Worksheets.Item('Data').Range('B2').Value2 = $NextInputValue
    $excel.CalculateFullRebuild()
    $oracle = @{}
    foreach ($target in $spec.targets) {
        $cell = $book.Worksheets.Item([string]$target.sheet).Cells.Item([int]$target.row + 1, [int]$target.col + 1)
        $value = $cell.Value2
        if ($cell.Text -match '^#(DIV/0!|N/A|NAME\?|REF!|VALUE!|NUM!|NULL!)$') { $oracle[$target.address] = "ERROR:$($cell.Text)" }
        elseif ($null -eq $value) { $oracle[$target.address] = 'EMPTY' }
        elseif ($value -is [bool]) { $oracle[$target.address] = "BOOL:$($value.ToString().ToLowerInvariant())" }
        elseif ($value -is [string]) { $oracle[$target.address] = "TEXT:$value" }
        else { $oracle[$target.address] = "NUM:$([double]$value)" }
    }
    $oracle | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $changedOraclePath -Encoding utf8
} finally {
    if ($book) { $book.Close($false) | Out-Null }
    if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# Alter only Data!B2 in its worksheet XML. Formula caches remain at the old input value.
$input = [IO.Compression.ZipFile]::OpenRead($source)
$outStream = [IO.File]::Create($stale)
$output = [IO.Compression.ZipArchive]::new($outStream, [IO.Compression.ZipArchiveMode]::Create, $false)
try {
    $workbookEntry = $input.GetEntry('xl/workbook.xml')
    $reader = [IO.StreamReader]::new($workbookEntry.Open()); [xml]$workbookXml = $reader.ReadToEnd(); $reader.Dispose()
    $ns = [Xml.XmlNamespaceManager]::new($workbookXml.NameTable)
    $ns.AddNamespace('m','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $ns.AddNamespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $sheet = $workbookXml.SelectSingleNode("//m:sheet[@name='Data']", $ns)
    $rid = $sheet.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $relsEntry = $input.GetEntry('xl/_rels/workbook.xml.rels')
    $reader = [IO.StreamReader]::new($relsEntry.Open()); [xml]$rels = $reader.ReadToEnd(); $reader.Dispose()
    $relNs = [Xml.XmlNamespaceManager]::new($rels.NameTable); $relNs.AddNamespace('p','http://schemas.openxmlformats.org/package/2006/relationships')
    $target = $rels.SelectSingleNode("//p:Relationship[@Id='$rid']", $relNs).GetAttribute('Target')
    $sheetPart = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { "xl/$target" }
    foreach ($entry in $input.Entries) {
        $newEntry = $output.CreateEntry($entry.FullName, [IO.Compression.CompressionLevel]::Optimal)
        $dest = $newEntry.Open()
        try {
            if ($entry.FullName -eq $sheetPart) {
                $reader = [IO.StreamReader]::new($entry.Open()); $sheetXml = $reader.ReadToEnd(); $reader.Dispose()
                # Preserve lexical XML offsets used by Formualizer's source-aware
                # cache patcher; change only the numeric <v> payload.
                $pattern = '(<c\b(?=[^>]*\br="B2")[^>]*>\s*<v>)' + [regex]::Escape([string]$OriginalCachedValue) + '(</v>\s*</c>)'
                $matches = [regex]::Matches($sheetXml, $pattern)
                if ($matches.Count -ne 1) { throw "Expected one Data!B2 value cache $OriginalCachedValue in $sheetPart; found $($matches.Count)" }
                $sheetXml = [regex]::Replace($sheetXml, $pattern, '${1}' + $StaleInputValue + '${2}', 1)
                $writer = [IO.StreamWriter]::new($dest, [Text.UTF8Encoding]::new($false))
                try { $writer.Write($sheetXml) } finally { $writer.Dispose() }
            } else {
                $src = $entry.Open(); try { $src.CopyTo($dest) } finally { $src.Dispose() }
            }
        } finally { $dest.Dispose() }
    }
} finally { $output.Dispose(); $outStream.Dispose(); $input.Dispose() }

$staleSpec = Get-Content -LiteralPath $specPath -Raw | ConvertFrom-Json
$staleSpec.mutations[0].value = $NextInputValue
$staleSpec | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $staleSpecPath -Encoding utf8
Write-Output "STALE_INPUT=$stale"
Write-Output "STALE_INPUT_VALUE=$StaleInputValue; NEXT_INPUT_VALUE=$NextInputValue"
Write-Output "EXPECTED_FULL_AFTER_STALE_INPUT=separate oracle for B2=$StaleInputValue"
Write-Output "EXPECTED_AFTER_NEXT_MUTATION=$changedOraclePath (B2=$NextInputValue)"
