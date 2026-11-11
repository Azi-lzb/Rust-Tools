$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$output = Join-Path $root 'reports\PACKAGE_MANIFEST.json'
$files = Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object {
    $_.FullName -notmatch '\\(bin|obj|publish|target|rust-target)\\' -and
    $_.Name -notmatch '\.(dll|so|dylib|lib|a|lic)$' -and
    $_.FullName -notmatch '\\.git\\' -and
    $_.FullName -ne $output
} | Sort-Object FullName | ForEach-Object {
    $relative = [System.IO.Path]::GetRelativePath($root, $_.FullName).Replace('\', '/')
    $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $_.FullName).Hash.ToLowerInvariant()
    [ordered]@{ path = $relative; bytes = $_.Length; sha256 = $hash }
}
$manifest = [ordered]@{
    schema_version = 1
    generated_utc = [DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
    root = 'benchmark/excel-linux-compat'
    note = 'Hashes cover harness, evidence, synthetic fixtures, and outputs; this manifest excludes itself. Commercial SDK binaries, license files, publish directories, and cargo target output are excluded.'
    files = @($files)
}
$manifest | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 $output
Write-Output "Wrote $output ($($files.Count) file entries)"

