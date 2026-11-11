# Linux glibc 2.28 Compatibility Results

Date: 2026-11-12. Scope: compatibility, self-contained/offline execution, real XLS/XLSX smoke edits, and a minimal PDF export check where implemented. This is not the broader capability/performance/bridge evaluation.

## Environment

- Runtime: Rocky Linux 8.9, x86_64, glibc 2.28; official Rocky image pinned as `rockylinux@sha256:9794037624aaa6212aeada1d28861ef5e0a935adaf93e4ef79837119f2a2d04c`.
- Container execution used `--network none`; runtime test containers used read-only rootfs with only `/tmp` writable.
- The .NET publish was directory-based, `linux-x64`, Release, self-contained. Apphosts were launched directly, not via `dotnet run`; runtime containers had no `dotnet`, Java, Python, LibreOffice, or `soffice` executable.
- C++ SDK/application build used GCC 8.5.0 and CMake in the glibc 2.28 toolchain image.
- Environment confirmation and per-run logs are retained under `evidence/`.
- Kylin V10 and UOS V20 are `NOT_TESTED`; a Rocky 8 result does not imply either vendor OS has passed.

## Results

| Implementation | SDK/package tested | Build/link | XLSX open/edit/save/reopen | XLS open/edit/save/reopen | XLS→XLSX | Offline/self-contained | PDF smoke | Overall glibc 2.28 |
|---|---|---:|---:|---:|---:|---:|---:|---|
| Aspose.Cells C++ | Linux SDK obtained from Aspose, 26.10 package | **INCOMPATIBLE** | NOT_RUN | NOT_RUN | NOT_RUN | NOT_RUN | NOT_RUN | **INCOMPATIBLE** |
| Aspose.Cells .NET | NuGet 26.10.0; .NET SDK 10.0.404/runtime 10.0.15 | PASS | PASS | PASS | PASS | PASS | PASS (6345 bytes) | **PASS** |
| Spire.XLS C++ | Linux SDK 16.7.0 | PASS, GCC 8.5 | runtime fails on a late-loaded native dependency | runtime fails before workbook load | NOT_RUN | NOT_RUN | NOT_RUN | **INCOMPATIBLE** |
| GemBox.Spreadsheet .NET | NuGet 2026.10.101; .NET SDK 10.0.404/runtime 10.0.15 | PASS | PASS | PASS | PASS | PASS | PASS (12785 bytes) | **PASS** (100-row Free limit applies) |
| LibXL C++ | Linux SDK 5.2.0.1 | PASS, GCC 8.5 | PASS | PASS | **UNSUPPORTED natively**; reconstruction only | PASS | UNSUPPORTED_BY_DESIGN | **PASS** |
| rXLS (historical reference) | benchmark harness at commit `fb05ebc` (harness 0.1.0, rxls dependency 0.1.4) | PASS, Release | generated XLSX validated with openpyxl | no in-place XLS editing | PASS for tested sample | PASS | unsupported by this converter | **PASS for conversion only** |

The Aspose C++ SDK's shared library requires `GLIBC_2.29` math symbols and `GLIBCXX_3.4.26`; link against Rocky 8 failed with unresolved references including `log2@GLIBC_2.29`, `exp2@GLIBC_2.29`, and `std::_Sp_make_shared_tag...@GLIBCXX_3.4.26`. The Linux SDK also documents GCC 9.4+ while this target baseline provides GCC 8.5. These are SDK compatibility failures, not missing test code. See `evidence/aspose-cpp/`.

Spire's main linked libraries do not expose the decisive dependency at startup. Its packaged `libSkiaSharp.so` is loaded dynamically, requires `GLIBC_2.29`, and on the target image `ldd` reports `/lib64/libm.so.6: version 'GLIBC_2.29' not found`. The actual C++ program builds but fails when creating a Workbook (`TypeInitialization_Type_NoTypeAvailable`). See `evidence/spire/libSkiaSharp-rocky8-ldd.txt` and `spire-offline-runtime.log`.

LibXL passes original XLS and XLSX edit/reopen checks on Rocky 8. It has no formula calculation engine. The tested C++ SDK rejected loading a BIFF XLS through its XMLBook API; the sample labeled `libxl-converted.xlsx` is explicitly reconstruction of scalar values/formula, not direct native conversion and not a fidelity-preserving conversion claim.

The LibXL 5.2.0.1 public C/C++ headers contain no PDF API token; `IBookT` exposes `save(filename)` and `saveRaw(data,size)` without an output format selector. The official product page lists Excel formats. PDF is recorded `UNSUPPORTED_BY_DESIGN`, and no PDF operation was attempted. See `evidence/libxl/libxl-pdf-api-review.txt`.

rXLS is retained only as the requested historical reference. It built on glibc 2.28 and converted the sample successfully (`ok=1`, `failed=0`, 112.37 ms reported by the converter); this says nothing about its previously documented formula/style limitations or general editing ability. ELF requirements peak at `GLIBC_2.28`, with only standard system libraries (`libgcc_s`, `libpthread`, `libm`, `libdl`, `libc`).

## Deployment and unresolved acceptance

- `PASS` means the tested sample was processed by a program that actually ran in the isolated glibc 2.28 environment. It does not constitute broad format fidelity or production acceptance.
- Aspose .NET and GemBox were published self-contained and ran with network disabled and no shared .NET installation. The resulting private publish directories were 143 MiB (Aspose) and 150 MiB (GemBox), including bundled runtime, ICU, and fonts. Aspose's PDF path needed private ICU/native assets; GemBox needed bundled DejaVu TTF fonts plus ICU. The custom font fallback is part of the compatibility harness. A production package needs a deliberate font bundle and must retain required font notices/licenses.
- The Aspose .NET run used evaluation mode because no commercial license file was supplied; its output workbooks contain `Evaluation Warning` sheets. Thus the API smoke operations ran, but those files are evaluation artifacts and not fidelity evidence for a licensed production build. The GemBox run operated under the Free-Limited cap; append was capped at 100 rows.
- LibXL's tested package staged its SDK shared library beside the executable; runtime system dependencies resolved on the target image. It is technically packageable, but this smoke did not bundle or test replacing the system `libstdc++`/glibc dependencies, so strict self-containment is `PARTIAL`. This proves neither permission to redistribute under every LibXL license tier nor target-machine dependency availability.
- No production license files/keys were supplied or committed. Production shipping rights are therefore `BLOCKED_LICENSE_REVIEW` for every commercial SDK; obtain and review the correct license before packaging. Evaluation entitlements were not bypassed.
- Kylin V10, UOS V20, `.deb` installation on those systems, dynamic load tracing over representative workbooks, and production font availability remain `NOT_TESTED`.

## Compatibility gate for the proposed Rust routing

The user's proposed layout can proceed as a candidate, subject to these Linux 2.28 results:

```text
Rust main program
├─ Rust native: calamine / rust_xlsxwriter                         NOT_TESTED HERE
├─ C ABI / C++ route: LibXL                                         GLIBC228 PASS
│  ├─ Aspose.Cells C++                                              INCOMPATIBLE
│  └─ Spire.XLS C++                                                 INCOMPATIBLE
└─ isolated worker route: Aspose.Cells .NET / GemBox.Spreadsheet   GLIBC228 PASS
   └─ Python / openpyxl                                             NOT_TESTED HERE
```

This only says which tested SDK distributions clear the glibc/offline smoke gate. It does not validate Rust FFI/IPC, full Excel feature fidelity, business correctness, or relative performance, and it is not the final architecture recommendation.

## Next stage recommendation

Advance **Aspose.Cells .NET, GemBox.Spreadsheet, and LibXL** to a bounded capability gate only after the team supplies/chooses the production license terms and confirms Kylin/UOS acceptance requirements. Keep Aspose C++ and Spire C++ out of glibc 2.28 performance ranking at these tested versions; their SDK binaries already fail the compatibility gate. If a vendor provides a later Linux package built for the 2.28 baseline, repeat this compatibility test first.

For capability tests, compare formula calculation only between engines that implement it; LibXL is excluded from formula correctness/rate ranking because the vendor confirms there is no calculation engine. For performance, first gate on workbook fidelity and business correctness, then use the original bounded workload (24 files, 150k rows, 1m rows) rather than timing the smoke fixtures. Measure .NET worker and Rust bridge costs separately after native engine capability tests; this compatibility project has not measured either.

### License/re-distribution notes

- **Aspose.Cells (.NET and C++):** the test ran without an applied commercial license. The .NET workbook outputs therefore include evaluation-warning sheets. Aspose's published licensing material differentiates developer, SDK, and OEM distribution models; the chosen product/platform license must match deployment to the intended customer locations. No SDK or license file is included here.
- **Spire.XLS C++:** the downloaded SDK is a temporary evaluation package. The vendor lists Small Business and OEM tiers separately; OEM is required for distribution across arbitrary locations and for the listed docker/cloud/public-facing cases. No production key was supplied.
- **GemBox.Spreadsheet:** this run was under Free-Limited behavior (100 append rows). The vendor states redistribution is permitted for files marked redistributable in product documentation and requires developers to be covered; verify the selected license and packaged-file list before shipping.
- **LibXL:** the vendor advertises royalty-free distribution inside commercial applications, while its purchase page distinguishes a per-developer Linux license from an enterprise unlimited developer/server license. Its download page says the unlicensed trial writes a banner and limits reading to 300 cells. This smoke confirms API behavior but does not validate a purchased key or the exact license tier for the rustTools deployment model.
- **rXLS historical harness:** `Cargo.toml` does not declare a license for the harness. The bundled `rxls` crate was resolved from crates.io at 0.1.4; independently verify the license of this historical version and each dependency before redistributing any resulting executable.

## Evidence map

- Structured status: `SMOKE_RESULTS.json`
- Runtime/package entries: `DEPENDENCIES.csv`, `PACKAGE_MANIFEST.json`
- Test harness source: `../scripts/`, `../aspose-cpp/`, `../spire/`, `../libxl/`, `../dotnet/`
- Full logs: `evidence/<engine>/`
- Openable output files: `samples/<engine>/`

## Vendor licensing references

These links summarize official published terms; they are not a legal determination. The exact license agreement delivered with a purchase governs.

- [Aspose.Cells C++ licensing](https://docs.aspose.com/cells/cpp/licensing/) and [Aspose license types](https://purchase.aspose.com/policies/license-types/)
- [Aspose use-license policy](https://purchase.aspose.com/policies/use-license/) and [licensing FAQ](https://purchase.aspose.com/faqs/licensing)
- [Spire.XLS C++ license tiers](https://cdn.e-iceblue.com/Buy/Spire.XLS-CPP.html); the vendor distinguishes internal/small-business and OEM distribution.
- [GemBox EULA](https://www.gemboxsoftware.com/purchase/eula) and [redistribution FAQ](https://www.gemboxsoftware.com/support/faq)
- [LibXL license/pricing](https://www.libxl.com/purchase.html), [distribution statement](https://www.libxl.com/home.html), and [trial limits](https://www.libxl.com/download.html); Linux developer and enterprise/server licenses are distinct.
- [LibXL formula FAQ](https://www.libxl.com/faq.html) confirms its API stores formula expressions but does not calculate them.
- [.NET official support policy](https://dotnet.microsoft.com/en-us/platform/support/policy): .NET 10 is LTS; our recorded runtime patch was 10.0.15.

