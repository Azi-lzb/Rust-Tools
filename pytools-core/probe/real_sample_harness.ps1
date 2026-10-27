<#
Development-only artifact comparator for the first PT real-sample batch.
It reads independently captured baseline/Rust JSON artifacts. It does not run
product commands, modify workbooks, manufacture capability evidence, or accept
a feature for production. Requires PowerShell 7.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Manifest,
    [Parameter(Mandatory)][string]$Report
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'PowerShell 7 is required.' }

function Read-Document([string]$Path) {
    return (Get-Content -LiteralPath $Path -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable)
}

function Compare-Value($Left, $Right, [string]$Location, $Differences) {
    if ($null -eq $Left -or $null -eq $Right) {
        if ($null -ne $Left -or $null -ne $Right) { $Differences.Add($Location + ': null mismatch') }
        return
    }
    if ($Left -is [System.Collections.IDictionary] -and $Right -is [System.Collections.IDictionary]) {
        $keys = @(@($Left.Keys) + @($Right.Keys) | Sort-Object -Unique)
        foreach ($key in $keys) {
            if (-not $Left.Contains($key) -or -not $Right.Contains($key)) {
                $Differences.Add($Location + '.' + $key + ': missing field')
            } else {
                Compare-Value $Left[$key] $Right[$key] ($Location + '.' + $key) $Differences
            }
        }
        return
    }
    $leftArray = $Left -is [array]
    $rightArray = $Right -is [array]
    if ($leftArray -or $rightArray) {
        if (-not ($leftArray -and $rightArray) -or $Left.Count -ne $Right.Count) {
            $Differences.Add($Location + ': array shape mismatch')
            return
        }
        for ($i = 0; $i -lt $Left.Count; $i++) {
            Compare-Value $Left[$i] $Right[$i] ($Location + '[' + $i + ']') $Differences
        }
        return
    }
    $numericTypes = @([byte], [int16], [int32], [int64], [uint64], [decimal], [single], [double], [System.Numerics.BigInteger])
    $leftNumber = $numericTypes -contains $Left.GetType()
    $rightNumber = $numericTypes -contains $Right.GetType()
    if ($leftNumber -and $rightNumber) {
        $floating = ($Left -is [double] -or $Left -is [single] -or $Right -is [double] -or $Right -is [single])
        if (-not $floating) {
            if ([string]$Left -cne [string]$Right) { $Differences.Add($Location + ': integer/decimal mismatch') }
        } else {
            $a = [double]$Left
            $b = [double]$Right
            if (-not [double]::IsFinite($a) -or -not [double]::IsFinite($b) -or
                [math]::Abs($a - $b) -gt (1e-15 * [math]::Max([math]::Abs($a), [math]::Abs($b)))) {
                $Differences.Add($Location + ': numeric tolerance exceeded')
            }
        }
        return
    }
    if ($Left.GetType() -ne $Right.GetType() -or $Left -cne $Right) {
        $Differences.Add($Location + ': value/type mismatch')
    }
}

function Assert-Artifact($Artifact, [string]$FeatureId, [string]$CaseId, $RequiredFields) {
    if ($Artifact.schema_version -ne 1 -or $Artifact.feature_id -cne $FeatureId -or $Artifact.case_id -cne $CaseId) {
        throw 'artifact schema or feature/case identity does not match'
    }
    if ($Artifact.origin -cne 'real_sanitized' -or $Artifact.approved_sanitized -ne $true) {
        throw 'artifact is not an explicitly approved sanitized real sample'
    }
    if ($Artifact.run.exit_code -ne 0 -or -not $Artifact.run.command -or
        $Artifact.run.isolated_copy -ne $true) {
        throw 'artifact lacks a successful command on an isolated copy'
    }
    if ($Artifact.inputs.Count -eq 0) { throw 'artifact has no input integrity evidence' }
    foreach ($inputItem in $Artifact.inputs) {
        if (-not $inputItem.id -or $inputItem.sha256_before -notmatch '^[a-fA-F0-9]{64}$' -or
            $inputItem.sha256_after -cne $inputItem.sha256_before) {
            throw 'artifact input hash is missing or changed'
        }
    }
    foreach ($field in $RequiredFields) {
        if (-not $Artifact.result.Contains($field)) { throw ('missing result contract: ' + $field) }
    }
}

$manifestPath = (Resolve-Path -LiteralPath $Manifest).Path
$manifestDir = Split-Path -Parent $manifestPath
$document = Read-Document $manifestPath
if ($document.schema_version -ne 1) { throw 'manifest schema_version must be 1' }
$requiredIds = @('PT-4-1','PT-4-6','PT-4-7','PT-4-8','PT-4-9','PT-4-10','PT-8-1','PT-8-2','PT-8-3','PT-8-4','PT-8-5','PT-7-4','PT-7-5')
$seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$items = [System.Collections.Generic.List[object]]::new()
foreach ($feature in $document.features) {
    $id = [string]$feature.feature_id
    if ($id -cnotin $requiredIds -or -not $seen.Add($id)) { throw ('unknown/duplicate feature_id: ' + $id) }
    if ($feature.cases.Count -eq 0) {
        $items.Add(@{ feature_id=$id; status='未验证'; comparison_status='未验证'; blockers=@('缺少经批准的脱敏真实样本及独立基线/Rust运行产物') })
        continue
    }
    foreach ($case in $feature.cases) {
        $caseId = [string]$case.case_id
        $blockers = [System.Collections.Generic.List[string]]::new()
        $differences = [System.Collections.Generic.List[string]]::new()
        $hashes = @{}
        try {
            if (-not $caseId) { throw 'case_id is required' }
            $baselinePath = [IO.Path]::GetFullPath([string]$case.baseline_artifact, $manifestDir)
            $rustPath = [IO.Path]::GetFullPath([string]$case.rust_artifact, $manifestDir)
            if ($baselinePath -eq $rustPath) { throw 'baseline and Rust artifacts must be independent files' }
            $baseline = Read-Document $baselinePath
            $rust = Read-Document $rustPath
            Assert-Artifact $baseline $id $caseId $feature.required_result_fields
            Assert-Artifact $rust $id $caseId $feature.required_result_fields
            Compare-Value $baseline.inputs $rust.inputs 'inputs' $differences
            Compare-Value $baseline.result $rust.result 'result' $differences
            $hashes.baseline_artifact_sha256 = (Get-FileHash -LiteralPath $baselinePath -Algorithm SHA256).Hash
            $hashes.rust_artifact_sha256 = (Get-FileHash -LiteralPath $rustPath -Algorithm SHA256).Hash
        } catch {
            $blockers.Add($_.Exception.Message)
        }
        $comparison = if ($blockers.Count) { '未验证' } elseif ($differences.Count) { '损坏' } else { '保留' }
        $items.Add(@{
            feature_id=$id; case_id=$caseId; status='未验证'; comparison_status=$comparison
            blockers=@($blockers); differences=@($differences); evidence=$hashes
            remaining_gates=@('真实文件/完整保留项的人工证据复核','同机同契约完整流程性能与预先定义资源预算','麒麟/UOS x86_64 glibc 2.28 deb验收','涉及既有簿编辑时的生产capability登记')
        })
    }
}
foreach ($id in $requiredIds) {
    if (-not $seen.Contains($id)) { throw ('manifest omits required feature: ' + $id) }
}
$output = @{
    schema_version=1; batch_id=$document.batch_id; status='未验证'
    generated_at=(Get-Date -Format o); stage='第二阶段第1步：真实样本产物对拍准备/比较'
    manifest_sha256=(Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
    harness_sha256=(Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash
    comparison_contract='类型与有序数组明确；跨实现浮点相对容差1e-15；空值/日期由产物生成器按功能合同归一；文件SHA只用于输入完整性/文件操作字节保留，不用于跨引擎工作簿逐格判等'
    result_items=@($items)
    limitations=@('本工具比较已独立捕获的运行产物，不代替端到端运行适配器','不读取或执行capability；不改变默认路由','无真实样本/目标平台验收时功能状态始终未验证')
}
$reportPath = [IO.Path]::GetFullPath($Report)
$reportRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'reports')) + [IO.Path]::DirectorySeparatorChar
if (-not $reportPath.StartsWith($reportRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'report must be written under pytools-core/probe/reports; source/sample directories are read-only'
}
if (Test-Path -LiteralPath $reportPath) { throw 'report already exists; use a new dated path' }
[IO.Directory]::CreateDirectory((Split-Path -Parent $reportPath)) | Out-Null
$output | ConvertTo-Json -Depth 40 | Set-Content -LiteralPath $reportPath -Encoding utf8
Write-Output ('已写报告：' + $reportPath)
if (@($items | Where-Object { $_.comparison_status -ne '保留' }).Count) { exit 2 }
exit 0
