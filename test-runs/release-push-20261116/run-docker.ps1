param([string]$Label='docker-workspace')
$ErrorActionPreference='Stop'
$repo=(Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$registry=(Get-Item 'C:/Users/AZI/.cargo/registry').Target
$arguments=@('run','--rm','--network','none','--read-only','-e',"RUSTTOOLS_TEST_LABEL=$Label",
    '--tmpfs','/tmp:rw,exec,nosuid,nodev,size=2g',
    '--tmpfs','/src/pytools-core/probe/reports:rw,nosuid,nodev,size=64m',
    '--mount',"type=bind,source=$repo,target=/src,readonly",
    '--mount',"type=bind,source=$repo/pytools-core/probe/reports,target=/report-source,readonly",
    '--mount',"type=bind,source=F:/RustBuild/pytools/release-push-20261116,target=/evidence",
    '--mount',"type=bind,source=$registry,target=/registry,readonly",
    '--mount','type=bind,source=F:\RustBuild\pytools,target=/build',
    '--mount','type=bind,source=F:\RustBuild\pytools\openpyxl-runtime-src\runtime,target=/runtime-src,readonly',
    '-w','/src','rusttools-excel-linux-compat-rxls:20261112','bash','/src/test-runs/release-push-20261116/run-docker.sh')
$prefix = if ($Label -eq 'docker-workspace') { 'docker' } else { $Label }
$arguments | ConvertTo-Json | Set-Content (Join-Path 'F:/RustBuild/pytools/release-push-20261116' "$prefix-arguments.json")
& docker @arguments *> (Join-Path 'F:/RustBuild/pytools/release-push-20261116' "$prefix-launch.log")
$code=$LASTEXITCODE
$code | Set-Content (Join-Path 'F:/RustBuild/pytools/release-push-20261116' "$prefix-launch.exit")
Get-Content (Join-Path 'F:/RustBuild/pytools/release-push-20261116' "$prefix-launch.log") -Tail 20
exit $code
