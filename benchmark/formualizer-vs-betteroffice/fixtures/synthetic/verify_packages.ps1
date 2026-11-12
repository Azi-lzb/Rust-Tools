param(
    [string]$FixtureDir = $PSScriptRoot,
    [string]$OutputDir = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'results\synthetic'),
    [string]$OraclePath = (Join-Path $PSScriptRoot 'excel-com-oracle.json'),
    [string]$InputPath = (Join-Path $PSScriptRoot 'formula-suite.xlsx'),
    [string]$SpecPath = (Join-Path $PSScriptRoot 'cases.json')
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$oracle = Get-Content -LiteralPath $OraclePath -Raw | ConvertFrom-Json -AsHashtable
$targets = Get-Content -LiteralPath $SpecPath -Raw | ConvertFrom-Json

function Read-ZipText($Archive, [string]$Name) {
    $entry = $Archive.GetEntry($Name)
    if (!$entry) { return $null }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-SheetTargetMap($Archive) {
    $xmlText = Read-ZipText $Archive 'xl/workbook.xml'
    if (!$xmlText) { return @{} }
    [xml]$xml = $xmlText
    $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
    $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $relsText = Read-ZipText $Archive 'xl/_rels/workbook.xml.rels'
    [xml]$rels = $relsText
    $relNs = [Xml.XmlNamespaceManager]::new($rels.NameTable)
    $relNs.AddNamespace('p', 'http://schemas.openxmlformats.org/package/2006/relationships')
    $relMap = @{}
    foreach ($rel in $rels.SelectNodes('//p:Relationship', $relNs)) { $relMap[$rel.Id] = $rel.Target }
    $map = @{}
    foreach ($sheet in $xml.SelectNodes('//m:sheets/m:sheet', $ns)) {
        $target = $relMap[$sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')]
        if ($target -notmatch '^/') { $target = "xl/$target" }
        $map[$sheet.GetAttribute('name')] = $target.Replace('\', '/')
    }
    return $map
}

function Get-PackageDigest([string]$Path) {
    if (!(Test-Path -LiteralPath $Path)) { return [pscustomobject]@{ path=$Path; status='MISSING' } }
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $sheetMap = Get-SheetTargetMap $archive
        $formulaCaches = @{}
        $loadedSheets = @{}
        foreach ($target in $targets.targets) {
            $part = $sheetMap[[string]$target.sheet]
            if (!$loadedSheets.ContainsKey($part)) {
                [xml]$sheetXml = Read-ZipText $archive $part
                $ns = [Xml.XmlNamespaceManager]::new($sheetXml.NameTable)
                $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
                $cellMap = @{}
                foreach ($sheetCell in $sheetXml.SelectNodes('//m:c', $ns)) {
                    $cellMap[$sheetCell.GetAttribute('r')] = $sheetCell
                }
                $loadedSheets[$part] = [pscustomobject]@{ xml=$sheetXml; ns=$ns; cells=$cellMap }
            }
            $sheetXml = $loadedSheets[$part].xml
            $ns = $loadedSheets[$part].ns
            $address = ([string]$target.address -split '!')[1]
            $cell = $loadedSheets[$part].cells[$address]
            if (!$cell) { $formulaCaches[$target.address] = $null; continue }
            $f = $cell.SelectSingleNode('m:f', $ns)
            $v = $cell.SelectSingleNode('m:v', $ns)
            $type = $cell.GetAttribute('t')
            if (!$v) { $cacheValue = 'EMPTY' }
            elseif ($type -eq 'e') { $cacheValue = "ERROR:$($v.InnerText)" }
            elseif ($type -eq 'str') { $cacheValue = "TEXT:$($v.InnerText)" }
            elseif ($type -eq 'b') { $cacheValue = "BOOL:$(([bool]([int]$v.InnerText)).ToString().ToLowerInvariant())" }
            elseif ($type -eq 's') {
                if (!$sharedStrings) { $sharedStrings = Read-ZipText $Archive 'xl/sharedStrings.xml' }
                $cacheValue = "SHARED_INDEX:$($v.InnerText)"
            }
            else { $cacheValue = "NUM:$($v.InnerText)" }
            $formulaCaches[$target.address] = [pscustomobject]@{ formula=if($f){$f.InnerText}else{$null}; cache=$cacheValue; cell_type=$type }
        }
        $names = @($archive.Entries | ForEach-Object FullName | Sort-Object)
        $parts = @{}
        foreach ($entry in $archive.Entries) {
            $stream = $entry.Open()
            try { $sha = [Security.Cryptography.SHA256]::HashData($stream); $parts[$entry.FullName] = [Convert]::ToHexString($sha) } finally { $stream.Dispose() }
        }
        return [pscustomobject]@{ path=$Path; status='READ'; parts=$parts; entries=$names; formula_caches=$formulaCaches }
    } finally { $archive.Dispose() }
}

$paths = [ordered]@{
    source = $InputPath
    betteroffice = Join-Path $OutputDir 'betteroffice.xlsx'
    formualizer_resident = Join-Path $OutputDir 'formualizer.xlsx'
    formualizer_cache = Join-Path $OutputDir 'formualizer-cache.xlsx'
}
$digests = [ordered]@{}
foreach ($entry in $paths.GetEnumerator()) { $digests[$entry.Key] = Get-PackageDigest $entry.Value }
$baseline = $digests.source.parts
$comparison = [ordered]@{ packages=$digests; cache_correctness=@{}; changed_zip_members=@{} }
foreach ($name in @('betteroffice','formualizer_resident','formualizer_cache')) {
    $digest = $digests[$name]
    if ($digest.status -ne 'READ') { continue }
    $match = 0
    $mismatchCount = 0
    $mismatch = [Collections.Generic.List[object]]::new()
    foreach ($address in $oracle.Keys) {
        $actual = $digest.formula_caches[$address].cache
        $expected = [string]$oracle[$address]
        if ($actual -eq $expected -or ($actual -like 'NUM:*' -and $expected -like 'NUM:*' -and [math]::Abs([double]$actual.Substring(4)-[double]$expected.Substring(4)) -le 1e-9)) { $match++ }
        else {
            $mismatchCount++
            if ($mismatch.Count -lt 100) {
                $mismatch.Add([pscustomobject]@{ address=$address; actual=$actual; expected=$expected })
            }
        }
    }
    $comparison.cache_correctness[$name] = [pscustomobject]@{
        status=if($mismatchCount -eq 0){'MATCH'}else{'MISMATCH'}
        matched=$match
        total=$oracle.Count
        mismatch_count=$mismatchCount
        mismatches=@($mismatch)
        mismatches_truncated=($mismatchCount -gt $mismatch.Count)
    }
    $changed = @()
    foreach ($part in $baseline.Keys) { if (!$digest.parts.ContainsKey($part) -or $digest.parts[$part] -ne $baseline[$part]) { $changed += $part } }
    foreach ($part in $digest.parts.Keys) { if (!$baseline.ContainsKey($part)) { $changed += $part } }
    $comparison.changed_zip_members[$name] = @($changed | Sort-Object -Unique)
}
$destination = Join-Path $OutputDir 'package-verification.json'
$comparison | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $destination -Encoding utf8
Write-Output "PACKAGE_VERIFICATION=$destination"
foreach ($name in $comparison.cache_correctness.Keys) {
    $r = $comparison.cache_correctness[$name]
    Write-Output "$name cache: $($r.status) $($r.matched)/$($r.total); changed parts=$($comparison.changed_zip_members[$name] -join ',')"
}
