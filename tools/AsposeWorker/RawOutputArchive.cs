using System.Security.Cryptography;
using System.Text.Json;

/// Append-only evidence; no worksheet, formula, style or license changes.
internal static class RawOutputArchive
{
    internal sealed record Info(string path, string sha256, string manifest_path, string stage);
    internal const string Stage = "sdk_save_before_worker_preservation";

    internal static Info? Capture(string candidate, string input, string requestId, string runId, string method)
    {
        var directory = Environment.GetEnvironmentVariable("RUSTTOOLS_ASPOSE_RAW_ARCHIVE_DIR");
        // Direct legacy probe invocations retain their existing behavior.
        if (string.IsNullOrEmpty(directory)) return null;
        if (!Path.IsPathFullyQualified(directory)) throw new ArgumentException("Raw SDK archive directory must be absolute");
        directory = Path.GetFullPath(directory);
        Directory.CreateDirectory(directory);
        var archive = Path.Combine(directory, $"sdk-save-{Environment.ProcessId}-{Guid.NewGuid():N}.xlsx");
        var manifest = archive + ".json";
        using (var source = new FileStream(candidate, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (var target = new FileStream(archive + ".partial", FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            source.CopyTo(target);
            target.Flush(true);
        }
        File.Move(archive + ".partial", archive, false);
        string hash;
        using (var stream = File.OpenRead(archive)) hash = Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
        using (var record = new FileStream(manifest, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            JsonSerializer.Serialize(record, new {
                schema_version = 1, event_name = "aspose_sdk_save_archived", stage = Stage,
                timestamp = DateTimeOffset.UtcNow, engine = "Aspose.Cells .NET",
                run_id = runId, request_id = requestId, method, input, candidate,
                archive, archive_sha256 = hash, evaluation_content_removed = false
            });
            record.Flush(true);
        }
        return new Info(archive, hash, manifest, Stage);
    }
}
