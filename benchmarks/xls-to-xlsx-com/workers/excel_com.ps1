param(
    [switch]$Stdin,
    [string]$InputFile
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)

if (-not ('RustToolsExcelComInterop' -as [type])) {
    $null = Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class RustToolsExcelComInterop {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    [DllImport("ole32.dll")] public static extern int CoInitializeEx(IntPtr reserved, uint coInit);
    [DllImport("ole32.dll")] public static extern void CoUninitialize();
}
'@
}

function Write-ProtocolLine([hashtable]$Payload) {
    [Console]::Out.WriteLine((ConvertTo-Json -InputObject $Payload -Compress -Depth 30))
    [Console]::Out.Flush()
}

function Release-ComObject($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        $null = [Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

$workerPid = $PID
$threadId = [RustToolsExcelComInterop]::GetCurrentThreadId()
Write-ProtocolLine @{ event = 'worker_ready'; worker_pid = $workerPid; sta_thread_id = $threadId }

$fatal = $null
$timings = @{}
$files = @()
$excel = $null
$workbooksCollection = $null
$configurationBook = $null
$excelPid = $null
$excelHandle = $null
$apartment = $false
$excelExitConfirmed = $null

try {
    if ($Stdin) {
        $jsonText = [Console]::In.ReadToEnd()
    } elseif (-not [string]::IsNullOrWhiteSpace($InputFile)) {
        $jsonText = [IO.File]::ReadAllText($InputFile, [Text.Encoding]::UTF8)
    } else {
        throw 'Specify -InputFile or -Stdin'
    }
    $task = ConvertFrom-Json -InputObject $jsonText -ErrorAction Stop

    $timer = [Diagnostics.Stopwatch]::StartNew()
    $hr = [RustToolsExcelComInterop]::CoInitializeEx([IntPtr]::Zero, 2)
    if ($hr -lt 0) { throw ('CoInitializeEx failed: 0x{0:X8}' -f $hr) }
    $apartment = $true
    $timings.sta_init_ms = $timer.Elapsed.TotalMilliseconds

    $timer.Restart()
    $before = [Collections.Generic.HashSet[uint32]]::new()
    foreach ($id in @($task.baseline_pids)) { $null = $before.Add([uint32]$id) }
    $excel = New-Object -ComObject ([string]$task.excel_progid)
    $hwnd = [IntPtr]$excel.Hwnd
    [uint32]$foundPid = 0
    $excelThread = [RustToolsExcelComInterop]::GetWindowThreadProcessId($hwnd, [ref]$foundPid)
    $workbooksCollection = $excel.Workbooks
    $workbookCount = [int]$workbooksCollection.Count
    if ($foundPid -eq 0 -or $before.Contains($foundPid) -or $workbookCount -ne 0) {
        throw "Refusing non-owned Excel instance PID=$foundPid, workbooks=$workbookCount"
    }
    $excelPid = [uint32]$foundPid
    $excelHandle = [Diagnostics.Process]::GetProcessById([int]$excelPid)
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $configurationBook = $workbooksCollection.Add()
    try {
        $excel.Calculation = -4135
        $excel.CalculateBeforeSave = $false
        if ([int]$excel.Calculation -ne -4135) { throw 'Excel did not enter manual calculation mode' }
    } finally {
        if ($null -ne $configurationBook) {
            $null = $configurationBook.Close($false)
            Release-ComObject $configurationBook
            $configurationBook = $null
        }
    }
    $timings.excel_create_config_ms = $timer.Elapsed.TotalMilliseconds

    $index = 0
    foreach ($entry in @($task.files)) {
        $itemTimer = [Diagnostics.Stopwatch]::StartNew()
        $item = @{
            source = [string]$entry.input
            target = [string]$entry.output
            index = $index
            status = 'failed'
        }
        $book = $null
        try {
            $phase = [Diagnostics.Stopwatch]::StartNew()
            $book = $workbooksCollection.Open([string]$entry.input, 0, $true)
            $item.open_ms = $phase.Elapsed.TotalMilliseconds
            $phase.Restart()
            $null = $book.SaveAs([string]$entry.output, 51)
            $item.save_as_ms = $phase.Elapsed.TotalMilliseconds
            $phase.Restart()
            $null = $book.Close($false)
            $item.close_ms = $phase.Elapsed.TotalMilliseconds
            Release-ComObject $book
            $book = $null
            $item.status = 'succeeded'
        } catch {
            $item.error = "Open/SaveAs $($entry.input): $($_.Exception.ToString())"
        } finally {
            if ($null -ne $book) {
                try {
                    $phase.Restart()
                    $null = $book.Close($false)
                    $item.close_ms = $phase.Elapsed.TotalMilliseconds
                } catch {
                    $item.close_error = $_.Exception.ToString()
                }
                Release-ComObject $book
                $book = $null
            }
            $item.total_ms = $itemTimer.Elapsed.TotalMilliseconds
            $files += $item
            $index++
            if (($index % 10) -eq 0) {
                Write-ProtocolLine @{ event = 'progress'; completed = $index; total = @($task.files).Count }
            }
        }
    }
} catch {
    $fatal = $_.Exception.ToString()
} finally {
    if ($null -ne $excel -and $null -ne $excelPid) {
        $timer = [Diagnostics.Stopwatch]::StartNew()
        try { $null = $excel.Quit() } catch { if ($null -eq $fatal) { $fatal = "Excel Quit failed: $($_.Exception)" } }
        $timings.excel_quit_ms = $timer.Elapsed.TotalMilliseconds
    }
    Release-ComObject $configurationBook
    Release-ComObject $workbooksCollection
    Release-ComObject $excel
    $configurationBook = $null
    $workbooksCollection = $null
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($apartment) {
        $timer = [Diagnostics.Stopwatch]::StartNew()
        [RustToolsExcelComInterop]::CoUninitialize()
        $timings.sta_release_ms = $timer.Elapsed.TotalMilliseconds
    }
    if ($null -ne $excelHandle) {
        $timer = [Diagnostics.Stopwatch]::StartNew()
        try { $excelExitConfirmed = $excelHandle.WaitForExit(30000) } catch { $excelExitConfirmed = $true }
        $timings.excel_exit_wait_ms = $timer.Elapsed.TotalMilliseconds
        $excelHandle.Dispose()
    }
}

Write-ProtocolLine @{
    event = 'worker_result'
    worker_pid = $workerPid
    sta_thread_id = $threadId
    excel_pid = $excelPid
    excel_instance_count = [int]($null -ne $excelPid)
    excel_exit_confirmed = $excelExitConfirmed
    timings = $timings
    files = $files
    fatal_error = $fatal
}
if ($null -ne $fatal -or $excelExitConfirmed -eq $false) { exit 2 }
exit 0
