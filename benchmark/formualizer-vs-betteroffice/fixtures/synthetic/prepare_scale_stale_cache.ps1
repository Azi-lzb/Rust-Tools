param(
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [Parameter(Mandatory=$true)][string]$StaleInputPath,
    [Parameter(Mandatory=$true)][string]$SpecPath,
    [Parameter(Mandatory=$true)][string]$StaleSpecPath,
    [double]$OriginalInput = 1,
    [double]$StaleInput = 25,
    [double]$NextInput = 35
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Read-ZipEntryText($Archive, [string]$Name) {
    $entry = $Archive.GetEntry($Name)
    if (!$entry) { throw "ZIP entry missing: $Name" }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}
$source = [IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $SourcePath).Path)
$destinationStream = [IO.File]::Create($StaleInputPath)
$destination = [IO.Compression.ZipArchive]::new($destinationStream, [IO.Compression.ZipArchiveMode]::Create, $false)
try {
    [xml]$workbook = (Read-ZipEntryText $source 'xl/workbook.xml')
    $workbookNs = [Xml.XmlNamespaceManager]::new($workbook.NameTable)
    $workbookNs.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $workbookNs.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $sheet = $workbook.SelectSingleNode("//m:sheet[@name='Data']", $workbookNs)
    if (!$sheet) { throw 'Data worksheet not found' }
    $relationshipId = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    [xml]$relationships = Read-ZipEntryText $source 'xl/_rels/workbook.xml.rels'
    $relationshipNs = [Xml.XmlNamespaceManager]::new($relationships.NameTable)
    $relationshipNs.AddNamespace('p', 'http://schemas.openxmlformats.org/package/2006/relationships')
    $relationship = $relationships.SelectSingleNode("//p:Relationship[@Id='$relationshipId']", $relationshipNs)
    if (!$relationship) { throw 'Data worksheet relationship not found' }
    $target = $relationship.GetAttribute('Target')
    $part = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { "xl/$target" }
    $patched = $false
    foreach ($entry in $source.Entries) {
        $newEntry = $destination.CreateEntry($entry.FullName, [IO.Compression.CompressionLevel]::Optimal)
        $dest = $newEntry.Open()
        try {
            if ($entry.FullName -eq $part) {
                $reader = [IO.StreamReader]::new($entry.Open())
                try { $sheetText = $reader.ReadToEnd() } finally { $reader.Dispose() }
                $escapedValue = [regex]::Escape($OriginalInput.ToString([Globalization.CultureInfo]::InvariantCulture))
                $pattern = '(<c\b(?=[^>]*\br="B2")[^>]*>\s*<v>)' + $escapedValue + '(</v>\s*</c>)'
                $updatedText = [regex]::Replace($sheetText, $pattern, {
                    param($match)
                    $script:patched = $true
                    return $match.Groups[1].Value + $StaleInput.ToString([Globalization.CultureInfo]::InvariantCulture) + $match.Groups[2].Value
                }, 1)
                if (!$patched) { throw "Expected one Data!B2 input value $OriginalInput in $part" }
                $writer = [IO.StreamWriter]::new($dest, [Text.UTF8Encoding]::new($false))
                try { $writer.Write($updatedText) } finally { $writer.Dispose() }
            } else {
                $inputStream = $entry.Open()
                try { $inputStream.CopyTo($dest) } finally { $inputStream.Dispose() }
            }
        } finally { $dest.Dispose() }
    }
} finally {
    $destination.Dispose()
    $destinationStream.Dispose()
    $source.Dispose()
}

$spec = Get-Content -LiteralPath $SpecPath -Raw | ConvertFrom-Json
if (!$spec.mutations -or $spec.mutations.Count -eq 0) { throw 'Fixture spec must include an input mutation' }
$spec.mutations[0].value = $NextInput
$spec | ConvertTo-Json -Depth 8 -Compress | Set-Content -LiteralPath $StaleSpecPath -Encoding utf8
Write-Output "STALE_INPUT=$StaleInputPath"
Write-Output "INPUT_VALUE=$StaleInput; formula caches retained from $OriginalInput; next mutation=$NextInput"
