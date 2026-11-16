// JSONL is UTF-8 regardless of the inherited Windows console code page.
Console.InputEncoding = new System.Text.UTF8Encoding(false, true);
Console.OutputEncoding = new System.Text.UTF8Encoding(false);
if (args.Length > 0 && args[0] == "--memory-save-probe") return MemorySaveProbe.Run(args);
if (args.Length == 3 && args[0] == "--style-save-diagnostic") return StyleSaveDiagnostic.Run(args[1], args[2]);
if (args.Length == 1 && args[0] == "--stdio") return ResidentWorker.Run();
Console.Error.WriteLine("Use --stdio or --memory-save-probe INPUT OUTPUT_DIRECTORY EXPECTED_ACTIVE_INDEX");
return 64;
