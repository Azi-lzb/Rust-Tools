using Aspose.Words;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

Console.InputEncoding = new UTF8Encoding(false, true);
Console.OutputEncoding = new UTF8Encoding(false);
if (args.Length != 1 || args[0] != "--stdio") return 64;
void Emit(object o) => Console.WriteLine(JsonSerializer.Serialize(o));
string Text(JsonElement q, string key) => q.GetProperty(key).GetString() ?? throw new ArgumentException(key);
bool EvaluationParagraph(string text) => text is
    "Evaluation Only. Created with Aspose.Words. Copyright 2003-2026 Aspose Pty Ltd." or
    "Created with an evaluation copy of Aspose.Words. To remove all limitations, you can use Free Temporary License https://products.aspose.com/words/temporary-license/" or
    "Created with an evaluation copy of Aspose.Words. To remove all limitations, you can use Free Temporary License  HYPERLINK \"https://products.aspose.com/words/temporary-license/\" https://products.aspose.com/words/temporary-license/";
const string EvaluationImageHash = "FF0BA8E6F56B061FAC6F1112C43C2C14D9B54153204B140AF9E1FA40C6104B34";
Emit(new { ready = true, protocol_version = 1, engine = "Aspose.Words", pid = Environment.ProcessId, evaluation_mode = true, max_attempts = 99 });
int attempts = 0;
string? line;
while ((line = Console.ReadLine()) is not null)
{
    string? id = null, run = null;
    var clock = Stopwatch.StartNew();
    try
    {
        if (line.Length > 1024 * 1024) throw new ArgumentException("Request too large");
        using var document = JsonDocument.Parse(line); var q = document.RootElement;
        id = Text(q, "request_id"); run = Text(q, "run_id");
        if (q.GetProperty("protocol_version").GetInt32() != 1) throw new NotSupportedException("Protocol version");
        if (Text(q, "method") == "Ping") { Emit(new { request_id = id, ok = true, attempts, pid = Environment.ProcessId }); continue; }
        if (Text(q, "method") != "Convert") throw new NotSupportedException("Only document conversion is supported");
        string input = Text(q, "input"), output = Text(q, "candidate");
        if (!Path.IsPathFullyQualified(input) || !Path.IsPathFullyQualified(output) || Path.GetFullPath(input).Equals(Path.GetFullPath(output), StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Distinct absolute input/candidate paths required");
        if (File.Exists(output) || Directory.Exists(output)) throw new IOException("Candidate already exists");
        if (!new[] { ".doc", ".docx" }.Contains(Path.GetExtension(input).ToLowerInvariant()) || !Path.GetExtension(output).Equals(".docx", StringComparison.OrdinalIgnoreCase)) throw new NotSupportedException("Only DOC/DOCX to DOCX is validated; WPS/WPT unverified");
        attempts++;
        var wb = new Document(input); double load = clock.Elapsed.TotalMilliseconds;
        // Raw evaluation headers/paragraphs remain in the saved artifact.
        var paragraphs = wb.Sections.Cast<Section>().SelectMany(s => s.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>())
            .Select(p => string.Concat(p.GetChildNodes(NodeType.Run, true).Cast<Run>().Select(r => r.Text))).Where(p => !EvaluationParagraph(p)).ToArray();
        if (Environment.GetEnvironmentVariable("ASPOSE_WORDS_DIAGNOSTIC_EVALUATION") == "1")
            foreach (string paragraph in paragraphs.Where(p => p.StartsWith("Created with an evaluation copy of Aspose.Words.", StringComparison.Ordinal)))
                Console.Error.WriteLine(JsonSerializer.Serialize(new { stage = "evaluation_text_diagnostic", text = paragraph }));
        string[] hashes = paragraphs.Select(p => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(p)))).ToArray();
        var auxiliaryHashes = wb.Sections.Cast<Section>().SelectMany(s => s.HeadersFooters.Cast<HeaderFooter>())
            .SelectMany(h => h.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>())
            .Select(p => string.Concat(p.GetChildNodes(NodeType.Run, true).Cast<Run>().Select(r => r.Text))).Where(p => !EvaluationParagraph(p))
            .Select(p => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(p)))).ToArray();
        var allImageHashes = wb.GetChildNodes(NodeType.Shape, true).Cast<Aspose.Words.Drawing.Shape>().Where(s => s.HasImage)
            .Select(s => Convert.ToHexString(SHA256.HashData(s.ImageData.ImageBytes))).ToArray();
        var imageHashes = allImageHashes.Where(h => h != EvaluationImageHash).ToArray();
        using (var stream = new FileStream(output, FileMode.CreateNew, FileAccess.Write, FileShare.None)) wb.Save(stream, SaveFormat.Docx);
        Emit(new { request_id = id, ok = true, candidate = output, attempts, draining = attempts >= 99, pid = Environment.ProcessId, evaluation_output = true,
            source_paragraph_hashes = hashes, source_sections = wb.Sections.Count, source_tables = wb.GetChildNodes(NodeType.Table, true).Count,
            source_auxiliary_paragraph_hashes = auxiliaryHashes, source_image_hashes = imageHashes,
            source_shapes = wb.GetChildNodes(NodeType.Shape, true).Count - (allImageHashes.Length - imageHashes.Length), load_ms = load, save_ms = clock.Elapsed.TotalMilliseconds - load, processing_ms = clock.Elapsed.TotalMilliseconds });
        Console.Error.WriteLine(JsonSerializer.Serialize(new { run_id = run, request_id = id, timestamp = DateTimeOffset.UtcNow.ToString("O"), input, candidate = output, stage = "document_convert", elapsed_ms = clock.Elapsed.TotalMilliseconds }));
    }
    catch (Exception ex)
    {
        Emit(new { request_id = id, ok = false, attempts, draining = attempts >= 99, pid = Environment.ProcessId, code = ex is NotSupportedException ? "UNSUPPORTED" : "ENGINE_ERROR", message = ex.Message });
        Console.Error.WriteLine(JsonSerializer.Serialize(new { run_id = run, request_id = id, exception = ex.ToString() }));
    }
    if (attempts >= 99) break;
}
return 0;
