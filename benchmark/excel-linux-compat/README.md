# Excel Linux glibc 2.28 Compatibility Experiment

This isolated test project evaluates six Linux implementations: Aspose.Cells C++, Aspose.Cells .NET, Spire.XLS C++, GemBox.Spreadsheet .NET, LibXL C++, and the historical rXLS converter source from commit `fb05ebc`.

The container target is Rocky Linux 8.9 x86_64 (glibc 2.28), pinned by image digest in `docker/`. This is a glibc baseline, not a substitute for Kylin V10 or UOS V20 device acceptance. Those OS results remain `NOT_TESTED`.

## Reproduce

PowerShell examples (SDK paths are local and are intentionally not stored in Git):

```powershell
benchmark\excel-linux-compat\scripts\run_aspose_cpp_rocky8.ps1 -SdkRoot <Aspose.Cells Linux SDK root>
benchmark\excel-linux-compat\scripts\run_spire_rocky8.ps1 -SdkRoot <Spire.XLS C++ Linux SDK root>
benchmark\excel-linux-compat\scripts\run_libxl_rocky8.ps1 -SdkRoot <LibXL Linux SDK root> -Mode capabilities
benchmark\excel-linux-compat\scripts\run_dotnet_smoke_rocky8.ps1 -Engine aspose -SdkRoot <expanded .NET SDK root>
benchmark\excel-linux-compat\scripts\run_dotnet_smoke_rocky8.ps1 -Engine gembox -SdkRoot <expanded .NET SDK root>
benchmark\excel-linux-compat\scripts\run_rxls_rocky8.ps1
benchmark\excel-linux-compat\scripts\capture_environment_rocky8.ps1
benchmark\excel-linux-compat\scripts\write_package_manifest.ps1
```

See `reports/evidence/` for captured build, dependency, and offline execution output. `reports/samples/` contains generated files only; commercial SDKs, license keys, and self-contained runtime bundles are not included.

Do not treat an evaluation build as authorization to ship. Review the vendor license selected for the actual deployment model before distributing a deb.

