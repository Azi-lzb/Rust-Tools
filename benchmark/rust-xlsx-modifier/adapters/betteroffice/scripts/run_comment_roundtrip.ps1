param([string]$Results = (Join-Path $PSScriptRoot '..\results'))
$ErrorActionPreference = 'Stop'
$Results = [System.IO.Path]::GetFullPath($Results)
$Base = Join-Path $Results 'audit_fixture.xlsx'
$CommentInput = Join-Path $Results 'comments_input.xlsx'
$Output = Join-Path $Results 'comments_after_edit.xlsx'
$Exe = 'F:\RustBuild\pytools-test\release\betteroffice-modification-audit.exe'
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false
try {
    $Wb = $Excel.Workbooks.Open($Base, 0, $false)
    $Ws = $Wb.Worksheets.Item('Data')
    $Ws.Range('A1').AddComment('audit comment should survive ordinary cell edit') | Out-Null
    $Wb.SaveAs($CommentInput, 51)
    $Wb.Close($false)
    # Excel COM normalizes the package, so append the opaque extension part
    # after COM creates the comment fixture and before the BetterOffice edit.
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $TempZip = "$CommentInput.tmp.zip"
    $SourceZip = [System.IO.Compression.ZipFile]::OpenRead($CommentInput)
    $TargetStream = [System.IO.File]::Open($TempZip, [System.IO.FileMode]::CreateNew)
    $TargetZip = [System.IO.Compression.ZipArchive]::new($TargetStream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
    try {
        foreach ($Entry in $SourceZip.Entries) {
            $NewEntry = $TargetZip.CreateEntry($Entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
            $InStream = $Entry.Open(); $OutStream = $NewEntry.Open()
            try {
                if ($Entry.FullName -eq '[Content_Types].xml') {
                    $Reader = [System.IO.StreamReader]::new($InStream)
                    $Xml = $Reader.ReadToEnd(); $Reader.Dispose()
                    if ($Xml -notmatch 'Extension="xml"') { $Xml = $Xml.Replace('</Types>', '<Default Extension="xml" ContentType="application/xml"/></Types>') }
                    $Bytes = [Text.Encoding]::UTF8.GetBytes($Xml); $OutStream.Write($Bytes, 0, $Bytes.Length)
                } else { $InStream.CopyTo($OutStream) }
            } finally { $InStream.Dispose(); $OutStream.Dispose() }
        }
        $Opaque = $TargetZip.CreateEntry('customXml/item1.xml', [System.IO.Compression.CompressionLevel]::Optimal)
        $OpaqueStream = $Opaque.Open(); $OpaqueBytes = [Text.Encoding]::UTF8.GetBytes('<extension marker="betteroffice-audit">opaque-1</extension>')
        $OpaqueStream.Write($OpaqueBytes, 0, $OpaqueBytes.Length); $OpaqueStream.Dispose()
    } finally { $TargetZip.Dispose(); $TargetStream.Dispose(); $SourceZip.Dispose() }
    Move-Item -LiteralPath $TempZip -Destination $CommentInput -Force
    $null = & $Exe edit $CommentInput $Output
    if ($LASTEXITCODE -ne 0) { throw "BetterOffice edit exited $LASTEXITCODE" }
    $Wb = $Excel.Workbooks.Open($Output, 0, $true)
    $Ws = $Wb.Worksheets.Item('Data')
    $CommentText = $Ws.Range('A1').Comment.Text()
    $EditedValue = $Ws.Range('E1').Value2
    $Wb.Close($false)
    function Get-PartHash([string]$Path, [string]$Part) {
        $Archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $Entry = $Archive.GetEntry($Part)
            if ($null -eq $Entry) { return $null }
            $Stream = $Entry.Open()
            try {
                $Ms = [System.IO.MemoryStream]::new()
                $Stream.CopyTo($Ms)
                return [Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData($Ms.ToArray()))
            } finally { $Stream.Dispose() }
        } finally { $Archive.Dispose() }
    }
    $Report = [ordered]@{
        comment_text = $CommentText
        ordinary_cell_value = $EditedValue
        custom_xml_input_sha256 = Get-PartHash $CommentInput 'customXml/item1.xml'
        custom_xml_output_sha256 = Get-PartHash $Output 'customXml/item1.xml'
        custom_xml_byte_identical = (Get-PartHash $CommentInput 'customXml/item1.xml') -eq (Get-PartHash $Output 'customXml/item1.xml')
    }
    $Report | ConvertTo-Json -Depth 3 | Set-Content -Encoding utf8 (Join-Path $Results 'comment-roundtrip.json')
    $Report | ConvertTo-Json -Depth 3
} finally {
    if ($Wb) { try { $Wb.Close($false) } catch {} }
    $Excel.Quit()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($Excel) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
