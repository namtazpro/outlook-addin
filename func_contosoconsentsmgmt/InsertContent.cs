using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Azure;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace Company.Function;

public class InsertContent
{
    private readonly ILogger<InsertContent> _logger;

    public InsertContent(ILogger<InsertContent> logger)
    {
        _logger = logger;
    }

    [Function("InsertContent")]
    public async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req)
    {
        _logger.LogInformation("InsertContent invoked.");

        var sasUrl = Environment.GetEnvironmentVariable("BlobContainerSasUrl");
        if (string.IsNullOrWhiteSpace(sasUrl))
        {
            _logger.LogError("BlobContainerSasUrl app setting is not configured.");
            return new ObjectResult("Storage is not configured.") { StatusCode = 500 };
        }

        // Read and parse the JSON body.
        string body;
        using (var reader = new StreamReader(req.Body, Encoding.UTF8))
        {
            body = await reader.ReadToEndAsync();
        }

        if (string.IsNullOrWhiteSpace(body))
        {
            return new BadRequestObjectResult("Request body is empty.");
        }

        JsonDocument doc;
        try
        {
            doc = JsonDocument.Parse(body);
        }
        catch (JsonException ex)
        {
            _logger.LogWarning(ex, "Invalid JSON payload.");
            return new BadRequestObjectResult($"Invalid JSON: {ex.Message}");
        }

        // Build a deterministic blob name from internetMessageId + capturedAt date prefix.
        var root = doc.RootElement;
        string? messageId = root.TryGetProperty("internetMessageId", out var midEl) && midEl.ValueKind == JsonValueKind.String
            ? midEl.GetString()
            : null;

        DateTime captured = DateTime.UtcNow;
        if (root.TryGetProperty("capturedAt", out var capEl) && capEl.ValueKind == JsonValueKind.String &&
            DateTime.TryParse(capEl.GetString(), null, System.Globalization.DateTimeStyles.RoundtripKind, out var parsed))
        {
            captured = parsed.ToUniversalTime();
        }

        string slug = !string.IsNullOrWhiteSpace(messageId)
            ? Sanitize(messageId!)
            : Guid.NewGuid().ToString("N");

        // Truncate to keep blob name reasonable.
        if (slug.Length > 120) slug = slug[..120];

        // Unique correlation id linking the JSON envelope to its extracted attachments.
        string correlationId = Guid.NewGuid().ToString("N");

        string folder = $"{captured:yyyy/MM/dd}/{correlationId}";
        string jsonBlobName = $"{folder}/{slug}.json";

        try
        {
            var containerClient = new BlobContainerClient(new Uri(sasUrl));

            // 1) Extract attachments first so we can rewrite the JSON envelope to reference them.
            var uploadedAttachments = new List<object>();

            if (root.TryGetProperty("attachments", out var attEl) && attEl.ValueKind == JsonValueKind.Array)
            {
                int index = 0;
                foreach (var att in attEl.EnumerateArray())
                {
                    if (att.ValueKind != JsonValueKind.Object) { index++; continue; }

                    if (!att.TryGetProperty("contentBase64", out var b64El) || b64El.ValueKind != JsonValueKind.String)
                    {
                        index++;
                        continue;
                    }

                    string fileName = att.TryGetProperty("name", out var nameEl) && nameEl.ValueKind == JsonValueKind.String
                        ? (nameEl.GetString() ?? $"attachment-{index}")
                        : $"attachment-{index}";
                    string contentType = att.TryGetProperty("contentType", out var ctEl) && ctEl.ValueKind == JsonValueKind.String
                        ? (ctEl.GetString() ?? "application/octet-stream")
                        : "application/octet-stream";
                    string? attachmentId = att.TryGetProperty("id", out var idEl) && idEl.ValueKind == JsonValueKind.String
                        ? idEl.GetString()
                        : null;

                    byte[] bytes;
                    try
                    {
                        bytes = Convert.FromBase64String(b64El.GetString() ?? "");
                    }
                    catch (FormatException ex)
                    {
                        _logger.LogWarning(ex, "Skipping attachment {Index} due to invalid base64.", index);
                        index++;
                        continue;
                    }

                    string safeName = SanitizeFileName(fileName);
                    string attachmentBlobName = $"{folder}/attachments/{index:D3}-{safeName}";

                    var attBlobClient = containerClient.GetBlobClient(attachmentBlobName);
                    var attMeta = new Dictionary<string, string>
                    {
                        ["correlationid"] = correlationId,
                        ["originalname"] = TruncateForMetadata(fileName),
                    };
                    if (!string.IsNullOrWhiteSpace(attachmentId))
                        attMeta["attachmentid"] = TruncateForMetadata(attachmentId!);

                    using (var attMs = new MemoryStream(bytes))
                    {
                        await attBlobClient.UploadAsync(attMs, new BlobUploadOptions
                        {
                            HttpHeaders = new BlobHttpHeaders { ContentType = contentType },
                            Metadata = attMeta
                        });
                    }

                    uploadedAttachments.Add(new
                    {
                        index,
                        id = attachmentId,
                        name = fileName,
                        contentType,
                        size = bytes.Length,
                        blobName = attachmentBlobName,
                        uri = attBlobClient.Uri.GetLeftPart(UriPartial.Path)
                    });

                    _logger.LogInformation("Uploaded attachment blob {BlobName} ({Bytes} bytes).", attachmentBlobName, bytes.Length);
                    index++;
                }
            }

            // 2) Rewrite the JSON envelope: drop contentBase64, add correlationId + per-attachment blobName/uri.
            var envelope = BuildEnvelope(root, correlationId, uploadedAttachments);
            byte[] envelopeBytes = JsonSerializer.SerializeToUtf8Bytes(envelope, new JsonSerializerOptions { WriteIndented = false });

            var jsonBlobClient = containerClient.GetBlobClient(jsonBlobName);
            var jsonHeaders = new BlobHttpHeaders { ContentType = "application/json; charset=utf-8" };

            var jsonMetadata = new Dictionary<string, string>
            {
                ["correlationid"] = correlationId,
                ["attachmentcount"] = uploadedAttachments.Count.ToString(),
            };
            if (root.TryGetProperty("subject", out var subjEl) && subjEl.ValueKind == JsonValueKind.String)
                jsonMetadata["subject"] = TruncateForMetadata(subjEl.GetString() ?? "");
            if (root.TryGetProperty("from", out var fromEl) &&
                fromEl.ValueKind == JsonValueKind.Object &&
                fromEl.TryGetProperty("address", out var fromAddrEl) &&
                fromAddrEl.ValueKind == JsonValueKind.String)
                jsonMetadata["from"] = TruncateForMetadata(fromAddrEl.GetString() ?? "");
            if (!string.IsNullOrWhiteSpace(messageId))
                jsonMetadata["messageid"] = TruncateForMetadata(messageId!);

            using (var ms = new MemoryStream(envelopeBytes))
            {
                await jsonBlobClient.UploadAsync(ms, new BlobUploadOptions
                {
                    HttpHeaders = jsonHeaders,
                    Metadata = jsonMetadata
                });
            }

            _logger.LogInformation("Uploaded envelope blob {BlobName} ({Bytes} bytes, correlationId={CorrelationId}).",
                jsonBlobName, envelopeBytes.Length, correlationId);

            return new OkObjectResult(new
            {
                correlationId,
                json = new
                {
                    blobName = jsonBlobName,
                    uri = jsonBlobClient.Uri.GetLeftPart(UriPartial.Path),
                    bytes = envelopeBytes.Length
                },
                attachments = uploadedAttachments
            });
        }
        catch (RequestFailedException ex)
        {
            _logger.LogError(ex, "Failed to upload blob(s) for correlation {CorrelationId}. Status {Status}.", correlationId, ex.Status);
            return new ObjectResult($"Storage error: {ex.ErrorCode ?? ex.Message}") { StatusCode = ex.Status };
        }
    }

    // Builds a JSON envelope from the original payload, stripping contentBase64 and adding storage references.
    private static Dictionary<string, object?> BuildEnvelope(
        JsonElement root,
        string correlationId,
        List<object> uploadedAttachments)
    {
        var envelope = new Dictionary<string, object?>
        {
            ["correlationId"] = correlationId
        };

        foreach (var prop in root.EnumerateObject())
        {
            if (prop.NameEquals("attachments")) continue; // replaced below
            envelope[prop.Name] = JsonSerializer.Deserialize<object?>(prop.Value.GetRawText());
        }

        // Rebuild attachments without contentBase64, with blob references merged in.
        var rebuilt = new List<Dictionary<string, object?>>();
        if (root.TryGetProperty("attachments", out var attEl) && attEl.ValueKind == JsonValueKind.Array)
        {
            int index = 0;
            foreach (var att in attEl.EnumerateArray())
            {
                var item = new Dictionary<string, object?>();
                if (att.ValueKind == JsonValueKind.Object)
                {
                    foreach (var prop in att.EnumerateObject())
                    {
                        if (prop.NameEquals("contentBase64")) continue; // strip
                        item[prop.Name] = JsonSerializer.Deserialize<object?>(prop.Value.GetRawText());
                    }
                }

                // Merge upload result for this index, if any.
                var uploaded = uploadedAttachments
                    .Select(o => (dynamic)o)
                    .FirstOrDefault(o => (int)o.index == index);
                if (uploaded != null)
                {
                    item["blobName"] = (string)uploaded.blobName;
                    item["uri"] = (string)uploaded.uri;
                    item["correlationId"] = correlationId;
                }

                rebuilt.Add(item);
                index++;
            }
        }
        envelope["attachments"] = rebuilt;
        return envelope;
    }

    // Strips angle brackets and replaces characters that aren't safe in blob names.
    private static string Sanitize(string input)
    {
        var trimmed = input.Trim().Trim('<', '>');
        var cleaned = Regex.Replace(trimmed, @"[^a-zA-Z0-9._@-]", "_");
        return cleaned.Trim('_', '.');
    }

    private static string SanitizeFileName(string input)
    {
        var cleaned = Regex.Replace(input.Trim(), @"[^a-zA-Z0-9._-]", "_").Trim('_', '.');
        if (string.IsNullOrEmpty(cleaned)) cleaned = "file";
        if (cleaned.Length > 120) cleaned = cleaned[..120];
        return cleaned;
    }

    // Blob metadata values must be ASCII and reasonably short.
    private static string TruncateForMetadata(string value)
    {
        var ascii = Regex.Replace(value, @"[^\x20-\x7E]", "?");
        return ascii.Length > 200 ? ascii[..200] : ascii;
    }
}