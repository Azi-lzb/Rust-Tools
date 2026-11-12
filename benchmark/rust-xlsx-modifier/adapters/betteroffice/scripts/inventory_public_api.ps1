param([string]$Output = (Join-Path $PSScriptRoot '..\results\api-inventory.txt'))
$ErrorActionPreference = 'Stop'
$RegistryRoot = Join-Path $env:USERPROFILE '.cargo\registry\src'
$Crates = [System.Collections.Generic.List[System.IO.DirectoryInfo]]::new()
foreach ($Registry in Get-ChildItem -Path $RegistryRoot -Directory -ErrorAction SilentlyContinue) {
    foreach ($Name in @('betteroffice-xlsx-0.5.1', 'betteroffice-xlsx-ops-0.5.1')) {
        $Candidate = Join-Path $Registry.FullName $Name
        if (Test-Path $Candidate) { $Crates.Add((Get-Item $Candidate)) }
    }
}
$ModelSource = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..\..\..\vendor\betteroffice-xlsx-model'))
if (Test-Path $ModelSource) { $Crates.Add((Get-Item $ModelSource)) }
if (-not $Crates) { throw 'Pinned BetterOffice 0.5.1 sources were not found in Cargo registry cache.' }
$Patterns = 'pub enum Op|PageSetup|PageMargins|page_setup|page_margins|orientation|paper_size|print_area|Print_Area|Print_Titles|header_footer|header|footer|margins|FitToPages|fit_to_page|print_order|horizontal_centered|vertical_centered|scale'
$Lines = [System.Collections.Generic.List[string]]::new()
$Lines.Add('BetterOffice 0.5.1 public API inventory')
$Lines.Add("Source roots: $($Crates.FullName -join '; ')")
foreach ($Crate in $Crates) {
    $Lines.Add("")
    $Lines.Add("## $($Crate.Name)")
    $SourceRoot = if ($Crate.Name -eq 'betteroffice-xlsx-model') { Join-Path $Crate.FullName 'src' } else { Join-Path $Crate.FullName 'src' }
    foreach ($File in Get-ChildItem -Path $SourceRoot -Recurse -Filter '*.rs' -File) {
        foreach ($Hit in Select-String -Path $File.FullName -Pattern $Patterns -CaseSensitive:$false) {
            $Relative = $File.FullName.Substring($Crate.FullName.Length).TrimStart('\')
            $Lines.Add("$Relative`:$($Hit.LineNumber): $($Hit.Line.Trim())")
        }
    }
}
$OpsFile = Join-Path ($Crates | Where-Object Name -eq 'betteroffice-xlsx-ops-0.5.1' | Select-Object -First 1 -ExpandProperty FullName) 'src\op.rs'
$Lines.Add('')
$Lines.Add('## Public Op variants')
$InEnum = $false
foreach ($Line in Get-Content $OpsFile) {
    if ($Line -match '^pub enum Op\s*\{') { $InEnum = $true; continue }
    if ($InEnum -and $Line -match '^}') { break }
    if ($InEnum -and $Line -match '^\s{4}([A-Z][A-Za-z0-9_]*)\s*(\{|,)') { $Lines.Add($Matches[1]) }
}
$Parent = Split-Path -Parent $Output
New-Item -ItemType Directory -Force -Path $Parent | Out-Null
[IO.File]::WriteAllLines([IO.Path]::GetFullPath($Output), $Lines, [Text.UTF8Encoding]::new($false))
$Lines | Select-String -Pattern 'PageSetup|PageMargins|page_setup|page_margins|orientation|paper_size|print_area|Print_Area|Print_Titles|header_footer|margins|FitToPages|fit_to_page|print_order|centered|Public Op variants' | ForEach-Object { $_.Line }
