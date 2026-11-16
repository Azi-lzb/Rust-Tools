$ErrorActionPreference='Stop'
$repo=(Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$arguments=@('run','--rm','--network','none','--read-only','--tmpfs','/tmp:rw,exec,nosuid,nodev,size=2g',
    '--mount',"type=bind,source=$repo,target=/src,readonly",
    '--mount',"type=bind,source=F:/RustBuild/pytools/release-push-20261116,target=/evidence",
    '--mount','type=bind,source=F:\RustBuild\pytools,target=/build,readonly',
    '--mount','type=bind,source=F:\RustBuild\pytools\openpyxl-runtime-src\runtime,target=/runtime-src,readonly',
    '-w','/src','rusttools-excel-linux-compat-rxls:20261112','bash','/src/test-runs/release-push-20261116/run-docker-python.sh')
& docker @arguments *> (Join-Path 'F:/RustBuild/pytools/release-push-20261116' 'docker-python-launch.log')
$code=$LASTEXITCODE
$code | Set-Content (Join-Path 'F:/RustBuild/pytools/release-push-20261116' 'docker-python-launch.exit')
Get-Content (Join-Path 'F:/RustBuild/pytools/release-push-20261116' 'docker-python-launch.log') -Tail 12
exit $code
