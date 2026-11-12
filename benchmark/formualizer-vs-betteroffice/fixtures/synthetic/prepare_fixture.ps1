param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'formula-suite.xlsx'),
    [string]$SpecPath = (Join-Path $PSScriptRoot 'cases.json'),
    [string]$OraclePath = (Join-Path $PSScriptRoot 'excel-com-oracle.json')
)
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Add()
    $data = $book.Worksheets.Item(1)
    $data.Name = 'Data'
    $calc = $book.Worksheets.Add()
    $calc.Name = 'Calc'

    $rows = @(
        @('Region', 'Amount', 'Type'),
        @('East', 10, 'Retail'),
        @('West', 20, 'Retail'),
        @('East', 30, 'Wholesale'),
        @('South', 40, 'Retail'),
        @('North', '15', 'Retail')
    )
    for ($r = 0; $r -lt $rows.Count; $r++) {
        for ($c = 0; $c -lt $rows[$r].Count; $c++) {
            $cell = $data.Cells.Item([int]($r + 1), [int]($c + 1))
            $value = $rows[$r][$c]
            if ($value -is [int] -or $value -is [double]) { $cell.Value2 = [double]$value }
            else { $cell.Value2 = [string]$value }
        }
    }
    # Force a text-typed numeric-looking cell; Excel otherwise coerces "15" to a number.
    $data.Range('B6').Formula = "'15"
    $data.Range('A1:C1').Font.Bold = $true
    $data.Range('B2:B6').NumberFormat = '#,##0.00'
    $data.Range('D1').Value2 = 'Merged fidelity marker'
    $data.Range('D1:E1').Merge()
    $data.Range('D1').Interior.Color = 0xD9EAD3
    $data.Columns.Item(1).ColumnWidth = 17
    $data.Rows.Item(1).RowHeight = 24
    $data.Hyperlinks.Add($data.Range('D2'), 'https://example.com/reference', '', 'synthetic hyperlink', 'reference') | Out-Null
    $data.PageSetup.Orientation = 2 # landscape
    $data.PageSetup.PaperSize = 9 # A4
    $data.PageSetup.PrintArea = '$A$1:$E$6'
    $data.PageSetup.PrintTitleRows = '$1:$1'
    $data.PageSetup.LeftMargin = $excel.InchesToPoints(0.25)
    $data.PageSetup.RightMargin = $excel.InchesToPoints(0.25)
    $data.PageSetup.CenterHorizontally = $true
    $data.PageSetup.LeftHeader = '&BFormula fixture'
    $data.PageSetup.CenterFooter = 'Page &P of &N'

    $book.Names.Add('NamedInput', '=Data!$B$3') | Out-Null
    $formulas = [ordered]@{
        B2 = '=SUM(Data!$B$2:$B$5)'
        B3 = '=SUMIF(Data!$A$2:$A$6,"East",Data!$B$2:$B$6)'
        B4 = '=SUMIFS(Data!$B$2:$B$6,Data!$A$2:$A$6,"East",Data!$C$2:$C$6,"Retail")'
        B5 = '=COUNTIF(Data!$A$2:$A$6,"East")'
        B6 = '=COUNTIFS(Data!$A$2:$A$6,"East",Data!$C$2:$C$6,"Retail")'
        B7 = '=IF(Data!B2>5,"yes","no")'
        B8 = '=IFS(Data!B2>5,"high",Data!B2>0,"low",TRUE,"zero")'
        B9 = '=AND(Data!B2>0,Data!B3>0)'
        B10 = '=OR(Data!B2=0,Data!B3>0)'
        B11 = '=ROUND(2.345,2)'
        B12 = '=ROUNDUP(2.341,2)'
        B13 = '=ROUNDDOWN(2.349,2)'
        B14 = '=VLOOKUP("West",Data!$A$2:$C$6,2,FALSE)'
        B15 = '=XLOOKUP("South",Data!$A$2:$A$6,Data!$B$2:$B$6,"missing")'
        B16 = '=INDEX(Data!$B$2:$B$6,MATCH("South",Data!$A$2:$A$6,0))'
        B17 = '=IFERROR(1/0,"safe")'
        B18 = '=DATE(2024,2,29)'
        B19 = '=YEAR(B18)'
        B20 = '=MONTH(B18)'
        B21 = '=DAY(B18)'
        B22 = '=LEFT("formula",3)'
        B23 = '=RIGHT("formula",3)'
        B24 = '=MID("formula",2,3)'
        B25 = '=TEXT(1234.5,"0.00")'
        B26 = '=CONCAT("rust","Tools")'
        B27 = '=Data!B2+Data!B3*2'
        B28 = '=NamedInput+Data!$B$2'
        B29 = '=SUM(Data!B2:B6)'
        B30 = '=Data!B2+Data!$B2+$B2+$B$2'
        B31 = '=Data!B6+1'
        B32 = '=1/0'
        B33 = '=FORMUALIZER_UNKNOWN_FN(1)'
        B34 = '=SUM(B2:B3)'
    }
    $targets = @()
    foreach ($entry in $formulas.GetEnumerator()) {
        $cell = $calc.Range($entry.Key)
        $cell.Formula = $entry.Value
        $targets += [pscustomobject]@{
            sheet = 'Calc'; row = [int]$cell.Row - 1; col = [int]$cell.Column - 1
            address = "Calc!$($entry.Key)"; formula = $entry.Value
        }
    }
    $calc.Range('A1').Value2 = 'Formula'
    $calc.Range('A1:B1').Font.Bold = $true
    $calc.Range('B2').Interior.Color = 0xFFF2CC
    $calc.PageSetup.Orientation = 1 # portrait on second sheet
    $calc.PageSetup.PaperSize = 9
    $calc.PageSetup.PrintArea = '$A$1:$B$34'
    $calc.PageSetup.PrintTitleRows = '$1:$1'

    $excel.CalculateFullRebuild()
    $book.ForceFullCalculation = $true
    $book.SaveAs($OutputPath, 51) # xlOpenXMLWorkbook
    $oracle = @{}
    foreach ($t in $targets) {
        $cell = $calc.Cells.Item($t.row + 1, $t.col + 1)
        $v = $cell.Value2
        if ($cell.Text -match '^#(DIV/0!|N/A|NAME\?|REF!|VALUE!|NUM!|NULL!)$') {
            $oracle[$t.address] = "ERROR:$($cell.Text)"
        } elseif ($null -eq $v) {
            $oracle[$t.address] = 'EMPTY'
        } elseif ($v -is [bool]) {
            $oracle[$t.address] = "BOOL:$($v.ToString().ToLowerInvariant())"
        } elseif ($v -is [string]) {
            $oracle[$t.address] = "TEXT:$v"
        } else {
            $oracle[$t.address] = "NUM:$([double]$v)"
        }
    }
    $incrementalOraclePath = Join-Path $PSScriptRoot 'excel-com-incremental-oracle.json'
    $data.Range('B2').Value2 = 25
    $excel.CalculateFullRebuild()
    $incrementalOracle = @{}
    foreach ($t in $targets) {
        $cell = $calc.Cells.Item($t.row + 1, $t.col + 1)
        $v = $cell.Value2
        if ($cell.Text -match '^#(DIV/0!|N/A|NAME\?|REF!|VALUE!|NUM!|NULL!)$') {
            $incrementalOracle[$t.address] = "ERROR:$($cell.Text)"
        } elseif ($null -eq $v) {
            $incrementalOracle[$t.address] = 'EMPTY'
        } elseif ($v -is [bool]) {
            $incrementalOracle[$t.address] = "BOOL:$($v.ToString().ToLowerInvariant())"
        } elseif ($v -is [string]) {
            $incrementalOracle[$t.address] = "TEXT:$v"
        } else {
            $incrementalOracle[$t.address] = "NUM:$([double]$v)"
        }
    }
    $incrementalOracle | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $incrementalOraclePath -Encoding utf8
    $data.Range('B2').Value2 = 10
    $excel.CalculateFullRebuild()
    $book.Save()
    $spec = [pscustomobject]@{
        targets = $targets
        mutations = @([pscustomobject]@{ sheet = 'Data'; row = 1; col = 1; value = 25 })
        notes = @('Workbook coordinates in the JSON are zero-based.', 'Excel COM expected values are generated after CalculateFullRebuild.')
    }
    $spec | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $SpecPath -Encoding utf8
    $oracle | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OraclePath -Encoding utf8
    Write-Output "FIXTURE=$OutputPath"
    Write-Output "TARGETS=$($targets.Count)"
    Write-Output "EXCEL_VERSION=$($excel.Version)"
    Write-Output "ORACLE=$OraclePath"
    Write-Output "INCREMENTAL_ORACLE=$incrementalOraclePath"
} finally {
    if ($book) { $book.Close($false) | Out-Null }
    if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
