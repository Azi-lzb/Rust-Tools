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
| LibXL C++ | Linux SDK 5.2.0.1 | PASS, GCC 8.5 | PASS | PASS | **UNSUPPORTED natively**; reconstruction only | PASS | UNSUPPORTED / not part of SDK smoke | **PASS** |
| rXLS (historical reference) | benchmark harness at commit `fb05ebc` (harness 0.1.0, rxls dependency 0.1.4) | PASS, Release | generated XLSX validated with openpyxl | no in-place XLS editing | PASS for tested sample | PASS | unsupported by this converter | **PASS for conversion only** |

The Aspose C++ SDK's shared library requires `GLIBC_2.29` math symbols and `GLIBCXX_3.4.26`; link against Rocky 8 failed with unresolved references including `log2@GLIBC_2.29`, `exp2@GLIBC_2.29`, and `std::_Sp_make_shared_tag...@GLIBCXX_3.4.26`. The Linux SDK also documents GCC 9.4+ while this target baseline provides GCC 8.5. These are SDK compatibility failures, not missing test code. See `evidence/aspose-cpp/`.

Spire's main linked libraries do not expose the decisive dependency at startup. Its packaged `libSkiaSharp.so` is loaded dynamically, requires `GLIBC_2.29`, and on the target image `ldd` reports `/lib64/libm.so.6: version 'GLIBC_2.29' not found`. The actual C++ program builds but fails when creating a Workbook (`TypeInitialization_Type_NoTypeAvailable`). See `evidence/spire/libSkiaSharp-rocky8-ldd.txt` and `spire-offline-runtime.log`.

LibXL passes original XLS and XLSX edit/reopen checks on Rocky 8. It has no formula calculation engine. The tested C++ SDK rejected loading a BIFF XLS through its XMLBook API; the sample labeled `libxl-converted.xlsx` is explicitly reconstruction of scalar values/formula, not direct native conversion and not a fidelity-preserving conversion claim.

rXLS is retained only as the requested historical reference. It built on glibc 2.28 and converted the sample successfully (`ok=1`, `failed=0`, 112.37 ms reported by the converter); this says nothing about its previously documented formula/style limitations or general editing ability. ELF requirements peak at `GLIBC_2.28`, with only standard system libraries (`libgcc_s`, `libpthread`, `libm`, `libdl`, `libc`).

## Deployment and unresolved acceptance

- `PASS` means the tested sample was processed by a program that actually ran in the isolated glibc 2.28 environment. It does not constitute broad format fidelity or production acceptance.
- Aspose .NET and GemBox were published self-contained and ran with network disabled and no shared .NET installation. The resulting private publish directories were 143 MiB (Aspose) and 150 MiB (GemBox), including bundled runtime, ICU, and fonts. Aspose's PDF path needed private ICU/native assets; GemBox needed bundled DejaVu TTF fonts plus ICU. The custom font fallback is part of the compatibility harness. A production package needs a deliberate font bundle and must retain required font notices/licenses.
- The Aspose .NET run used evaluation mode because no commercial license file was supplied; its output workbooks contain `Evaluation Warning` sheets. Thus the API smoke operations ran, but those files are evaluation artifacts and not fidelity evidence for a licensed production build. The GemBox run operated under the Free-Limited cap; append was capped at 100 rows.
- LibXL's tested package staged its SDK shared library beside the executable; runtime system dependencies resolved on the target image. It is technically packageable, but this smoke did not bundle or test replacing the system `libstdc++`/glibc dependencies, so strict self-containment is `PARTIAL`. This proves neither permission to redistribute under every LibXL license tier nor target-machine dependency availability.
- No production license files/keys were supplied or committed. Production shipping rights are therefore `BLOCKED_LICENSE_REVIEW` for every commercial SDK; obtain and review the correct license before packaging. Evaluation entitlements were not bypassed.
- Kylin V10, UOS V20, `.deb` installation on those systems, dynamic load tracing over representative workbooks, and production font availability remain `NOT_TESTED`.

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
- [LibXL license/pricing](https://www.libxl.com/purchase.html); Linux developer and enterprise/server licenses are distinct.
- [.NET official support policy](https://dotnet.microsoft.com/en-us/platform/support/policy): .NET 10 is LTS; our recorded runtime patch was 10.0.15.

