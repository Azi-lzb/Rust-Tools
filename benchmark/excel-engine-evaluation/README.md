# rustTools Excel engine evaluation lab

This isolated experiment compares exactly five implementations: Aspose.Cells C++, Aspose.Cells .NET, Spire.XLS C++, GemBox.Spreadsheet .NET, and LibXL C++. It does not modify rustTools production code. Commercial SDK binaries, license keys, and runtime bundles stay outside Git.

The work proceeds through gated phases: (1) real glibc 2.28 and offline deployment checks, (2) workbook capability validation only for eligible engines, (3) moderate performance tests only after capability results, then (4) runnable Rust bridge prototypes and (5) an evidence-based architecture recommendation. No PDF export or PDF performance test is in scope. The historical rXLS experiment is not rerun or included in this five-engine comparison.

## Current measured result

LibXL 5.2.0 C++, Aspose.Cells .NET 26.10.0 and GemBox.Spreadsheet .NET 2026.10.101 have passed Rocky Linux 8.9 / glibc 2.28 offline read/modify/save/reopen smoke tests. Aspose.Cells C++ 26.10 and Spire.XLS C++ 16.7.0 are incompatible with this target: their Linux C++ smoke paths require GLIBC 2.29 (Aspose also has a link-time GLIBCXX/FreeType mismatch). LibXL's XLS-to-XLSX sample is a cell reconstruction, not a general fidelity-preserving converter; formula calculation is unsupported by design. Kylin V10 and UOS V20 remain untested. See [01_LINUX_COMPATIBILITY.md](reports/01_LINUX_COMPATIBILITY.md) and retained logs under `reports/evidence/`.

Phase 1 is complete for all five implementations on the Rocky 8.9/glibc 2.28 test host. Phase 2 has run on the three compatible implementations; Aspose .NET/GemBox pass the tested XLS/XLSX edits, conversion and formula scenarios, with license and fidelity limits recorded, while LibXL passes bounded modification/print tests and has no calculation engine. Phase 3 has completed 24 synthetic workbooks × 3 Linux rounds plus one 150k and one 1m XLSX capacity run; GemBox batch runs passed but the large cases were blocked by the 150-row free-version cap. An additional same-host Windows conversion run compared Aspose .NET, GemBox and resident Excel COM on 8 XLS files × 3 rounds. Phase 4 has runnable LibXL C ABI, Aspose .NET Worker and openpyxl IPC prototypes; phase 5 has a provisional recommendation. See reports `01`–`05` and raw CSV evidence.

## Reproduce LibXL phase 1

Download the official Linux C++ SDK outside this repository and extract it. In PowerShell, run:

```powershell
.\benchmark\excel-engine-evaluation\scripts\run_libxl_rocky8.ps1 `
  -SdkRoot "$env:TEMP\libxl-sdk\libxl-5.2.0.1"
```

The script builds in a pinned Rocky Linux 8 container, packages only the smoke executable and `libxl.so` into a temporary directory, then runs it in the same clean image with networking disabled, a read-only root filesystem, and no Python, Java, .NET, Office, or LibreOffice runtime. Generated samples and logs are retained in `reports/`.

The evaluation SDK inserts a banner in the first row and limits reads to 300 cells. No commercial license key is included. Obtain and comply with an appropriate license before production or distribution use.

## Reproduce bridge prototypes

Build the pinned Rocky 8.9/glibc 2.28 bridge toolchain once, then run the LibXL C ABI + Rust FFI prototype:

```powershell
.\benchmark\excel-engine-evaluation\scripts\build_bridge_toolchain.ps1
.\benchmark\excel-engine-evaluation\scripts\run_libxl_rust_ffi_rocky8.ps1 `
  -SdkRoot "$env:TEMP\libxl-sdk\libxl-5.2.0.1"
```

The self-contained Aspose .NET resident Worker prototype also builds/runs in the same glibc 2.28 container:

```powershell
.\benchmark\excel-engine-evaluation\scripts\run_aspose_dotnet_rust_worker_rocky8.ps1 `
  -SdkRoot "$env:TEMP\dotnet-sdk-10.0.404-expanded"
```

The test-only Rust→openpyxl JSON-file Worker runs on the Windows development host and uses the locally available Python/openpyxl runtime (Python 3.12.10 / openpyxl 3.1.5 in the measured run):

```powershell
.\benchmark\excel-engine-evaluation\scripts\run_openpyxl_bridge_windows.ps1
```

These bridge samples remain experimental and do not change production routing. Detailed results are in `reports/04_RUST_BRIDGE.md`.

