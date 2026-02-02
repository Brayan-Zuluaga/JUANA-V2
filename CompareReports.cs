using DocumentFormat.OpenXml;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

namespace OpenXmlFunc;

public class CompareReports
{
    private readonly ILogger _logger;

    public CompareReports(ILoggerFactory loggerFactory)
    {
        _logger = loggerFactory.CreateLogger<CompareReports>();
    }

    // ====== DTOs ======
    public record CompareRequest(
        string baselineDocxBase64,
        string currentDocxBase64,
        Options? options,
        Metadata? metadata
    );

    public record Options(
        string? mode,               // MVP: "delta_document"
        bool? significantOnly,      // true => omite sin cambios
        bool? includeHighlights,    // true => agrega sección de cambios
        int? maxHighlights          // ej: 12
    );

    public record Metadata(
        string? gerente,
        string? mercado,
        string? baselineDate,
        string? currentDate
    );

    public record CompareResponse(
        string fileName,
        string docxBase64,
        Summary summary
    );

    public record Summary(
        int critical,
        int high,
        int medium,
        int low,
        int newRisk,
        int updated,
        int noChange
    );

    // ====== Modelo interno ======
    public record Block(
        string itemKey,
        string title,
        string body,
        bool hasRiskFlag,
        bool isSinNovedad
    );

    public enum Severity { Low, Medium, High, Critical }
    public enum Tag { NoChange, Updated, NewRisk, New }

    public record DeltaItem(
        string itemKey,
        string title,
        Tag tag,
        Severity severity,
        string note,
        Block? previous,
        Block current
    );

    // ====== Function ======
    [Function("CompareReports")]
    public async Task<HttpResponseData> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
    {
        try
        {
            var body = await new StreamReader(req.Body).ReadToEndAsync();
            var input = JsonSerializer.Deserialize<CompareRequest>(body, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (input == null ||
                string.IsNullOrWhiteSpace(input.baselineDocxBase64) ||
                string.IsNullOrWhiteSpace(input.currentDocxBase64))
            {
                var bad = req.CreateResponse(HttpStatusCode.BadRequest);
                await bad.WriteStringAsync("Missing baselineDocxBase64/currentDocxBase64.");
                return bad;
            }

            var opts = input.options ?? new Options(
                mode: "delta_document",
                significantOnly: true,
                includeHighlights: true,
                maxHighlights: 12
            );

            // 1) DOCX -> texto
            var baselineText = ExtractTextFromDocxBase64(input.baselineDocxBase64);
            var currentText  = ExtractTextFromDocxBase64(input.currentDocxBase64);

            // 2) Texto -> bloques
            var baselineBlocks = SplitIntoBlocks(baselineText);
            var currentBlocks  = SplitIntoBlocks(currentText);

            // 3) Comparación
            var deltas = Compare(baselineBlocks, currentBlocks);

            // 4) Generación DOCX resultado
            var outBytes = BuildDeltaDocx(deltas, input.metadata, opts);

            var response = new CompareResponse(
                fileName: BuildFileName(input.metadata),
                docxBase64: Convert.ToBase64String(outBytes),
                summary: BuildSummary(deltas)
            );

            var ok = req.CreateResponse(HttpStatusCode.OK);
            ok.Headers.Add("Content-Type", "application/json; charset=utf-8");
            await ok.WriteStringAsync(JsonSerializer.Serialize(response));
            return ok;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "CompareReports failed");
            var err = req.CreateResponse(HttpStatusCode.InternalServerError);
            await err.WriteStringAsync($"Error: {ex.Message}");
            return err;
        }
    }

    // ====== 1) DOCX -> texto ======
    private static string ExtractTextFromDocxBase64(string base64)
    {
        var bytes = Convert.FromBase64String(base64);
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);

        var sb = new StringBuilder();
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "";

        foreach (var para in body.Elements<Paragraph>())
        {
            var text = para.InnerText?.Trim();
            sb.AppendLine(text ?? "");
        }
        return sb.ToString();
    }

    // ====== 2) Texto -> bloques ======
    private static List<Block> SplitIntoBlocks(string text)
    {
        var rawBlocks = Regex.Split(text, @"\n\s*\n")
            .Select(b => b.Trim())
            .Where(b => b.Length > 0)
            .ToList();

        var blocks = new List<Block>();

        foreach (var raw in rawBlocks)
        {
            var lines = raw.Split('\n')
                .Select(l => l.Trim())
                .Where(l => l.Length > 0)
                .ToList();

            var title = GuessTitle(lines);
            var hasRisk = Regex.IsMatch(raw, @"\b\[?RIESGO\]?\b", RegexOptions.IgnoreCase);
            var sinNovedad = Regex.IsMatch(raw, @"\bsin novedad\b", RegexOptions.IgnoreCase);

            var keySeed = Normalize(title);
            if (keySeed.Length < 8)
                keySeed = Normalize(lines.FirstOrDefault() ?? raw[..Math.Min(raw.Length, 60)]);

            var key = Sha1(keySeed);

            blocks.Add(new Block(key, title, raw, hasRisk, sinNovedad));
        }

        return blocks;
    }

    private static string GuessTitle(List<string> lines)
    {
        var titleParts = new List<string>();

        foreach (var l in lines.Take(6))
        {
            if (Regex.IsMatch(l, @"^(seguimiento|fecha:|inditex|general|vertical|squad|área|datos global|partners|otros varios)",
                RegexOptions.IgnoreCase))
            {
                titleParts.Add(l);
            }
        }

        if (titleParts.Count > 0)
            return string.Join(" · ", titleParts.Take(3));

        return lines.FirstOrDefault() ?? "Bloque";
    }

    // ====== 3) Comparación ======
    private static List<DeltaItem> Compare(List<Block> baseline, List<Block> current)
    {
        var baseMap = baseline.ToDictionary(b => b.itemKey, b => b);

        var deltas = new List<DeltaItem>();
        foreach (var cur in current)
        {
            baseMap.TryGetValue(cur.itemKey, out var prev);
            var (tag, sev, note) = EvaluateDelta(prev, cur);

            deltas.Add(new DeltaItem(cur.itemKey, cur.title, tag, sev, note, prev, cur));
        }

        return deltas
            .OrderByDescending(d => d.severity)
            .ThenBy(d => d.title)
            .ToList();
    }

    private static (Tag tag, Severity severity, string note) EvaluateDelta(Block? prev, Block cur)
    {
        if (prev == null)
            return (Tag.New, Severity.Medium, "Aparece este bloque por primera vez en el informe actual.");

        // NUEVO RIESGO
        if (!prev.hasRiskFlag && cur.hasRiskFlag)
            return (Tag.NewRisk, Severity.High, "Aparece marcado como [RIESGO] en la semana actual.");

        // SIN NOVEDAD -> ahora hay acciones
        if (prev.isSinNovedad && !cur.isSinNovedad)
        {
            var sev = cur.hasRiskFlag ? Severity.High : Severity.Medium;
            return (Tag.Updated, sev, "Pasa de 'Sin novedad' a incluir acciones/seguimiento.");
        }

        // Confirmaciones típicas
        if (HasConfirmations(cur.body) && !HasConfirmations(prev.body))
        {
            var sev = cur.hasRiskFlag ? Severity.High : Severity.Medium;
            return (Tag.Updated, sev, "Se concreta información que antes era preliminar.");
        }

        // Actualización por cambio de contenido
        var similarity = RoughSimilarity(prev.body, cur.body);
        if (similarity < 0.70)
        {
            var sev = cur.hasRiskFlag ? Severity.Medium : Severity.Low;
            return (Tag.Updated, sev, "Contenido actualizado respecto a la semana anterior.");
        }

        return (Tag.NoChange, Severity.Low, "Sin cambios relevantes detectados.");
    }

    private static bool HasConfirmations(string s)
        => Regex.IsMatch(s, @"\b(confirmad|confirmado|arrancan|arranca|nuestro tl será|liderará|movimiento de tl)\b",
            RegexOptions.IgnoreCase);

    private static double RoughSimilarity(string a, string b)
    {
        var ta = Tokenize(a);
        var tb = Tokenize(b);
        if (ta.Count == 0 && tb.Count == 0) return 1.0;
        if (ta.Count == 0 || tb.Count == 0) return 0.0;

        var inter = ta.Intersect(tb).Count();
        var uni = ta.Union(tb).Count();
        return uni == 0 ? 0 : (double)inter / uni;
    }

    private static HashSet<string> Tokenize(string s)
    {
        var norm = Normalize(s);
        var tokens = norm.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return tokens.Where(t => t.Length > 2).ToHashSet();
    }

    // ====== 4) DOCX resultado ======
    private static byte[] BuildDeltaDocx(List<DeltaItem> deltas, Metadata? meta, Options opts)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            var body = main.Document.Body!;

            body.Append(MakeHeading($"Delta semanal – {meta?.mercado ?? "Mercado"} – {meta?.gerente ?? "Gerente"}", 1));
            body.Append(MakeParagraph($"Comparación: {meta?.baselineDate ?? "baseline"} → {meta?.currentDate ?? "current"}"));
            body.Append(MakeParagraph(" "));

            if (opts.includeHighlights == true)
            {
                body.Append(MakeHeading("Cambios relevantes de la semana", 2));

                var highlights = deltas
                    .Where(d => d.severity >= Severity.High && d.tag != Tag.NoChange)
                    .Take(opts.maxHighlights ?? 12)
                    .ToList();

                if (highlights.Count == 0)
                    body.Append(MakeParagraph("No se detectan cambios de severidad alta/crítica."));
                else
                    foreach (var h in highlights)
                        body.Append(MakeBullet($"{FormatTag(h.tag)} {h.title} — {h.note}"));

                body.Append(MakeParagraph(" "));
            }

            body.Append(MakeHeading("Detalle por bloques", 2));

            foreach (var d in deltas)
            {
                if (opts.significantOnly == true && d.tag == Tag.NoChange) continue;

                body.Append(MakeHeading($"{FormatTag(d.tag)} {d.title}", 3));
                body.Append(MakeParagraph($"Severidad: {d.severity}. {d.note}"));
                body.Append(MakeParagraph(" "));
            }

            main.Document.Save();
        }

        return ms.ToArray();
    }

    private static Summary BuildSummary(List<DeltaItem> deltas)
    {
        int critical = deltas.Count(d => d.severity == Severity.Critical);
        int high     = deltas.Count(d => d.severity == Severity.High);
        int medium   = deltas.Count(d => d.severity == Severity.Medium);
        int low      = deltas.Count(d => d.severity == Severity.Low);

        int newRisk  = deltas.Count(d => d.tag == Tag.NewRisk);
        int updated  = deltas.Count(d => d.tag == Tag.Updated || d.tag == Tag.New);
        int noChange = deltas.Count(d => d.tag == Tag.NoChange);

        return new Summary(critical, high, medium, low, newRisk, updated, noChange);
    }

    private static string BuildFileName(Metadata? meta)
    {
        var mercado = SafeFile(meta?.mercado ?? "Mercado");
        var gerente = SafeFile(meta?.gerente ?? "Gerente");
        var date = SafeFile(meta?.currentDate ?? DateTime.UtcNow.ToString("yyyy-MM-dd"));
        return $"Delta_{mercado}_{gerente}_{date}.docx";
    }

    // ====== Helpers ======
    private static Paragraph MakeHeading(string text, int level)
    {
        var p = new Paragraph();
        var pPr = new ParagraphProperties
        {
            ParagraphStyleId = new ParagraphStyleId { Val = $"Heading{level}" }
        };
        p.Append(pPr);
        p.Append(new Run(new Text(text)));
        return p;
    }

    private static Paragraph MakeParagraph(string text)
        => new Paragraph(new Run(new Text(text)));

    private static Paragraph MakeBullet(string text)
        => new Paragraph(new Run(new Text("• " + text)));

    private static string FormatTag(Tag tag) => tag switch
    {
        Tag.NewRisk => "[NUEVO RIESGO]",
        Tag.Updated => "[ACTUALIZACIÓN]",
        Tag.New     => "[NUEVO]",
        _           => "[SIN CAMBIOS]"
    };

    private static string Normalize(string s)
    {
        var lower = s.ToLowerInvariant();
        lower = Regex.Replace(lower, @"[^\p{L}\p{N}\s]", " ");
        lower = Regex.Replace(lower, @"\s+", " ").Trim();
        return lower;
    }

    private static string Sha1(string s)
    {
        var bytes = Encoding.UTF8.GetBytes(s);
        var hash = SHA1.HashData(bytes);
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    private static string SafeFile(string s)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            s = s.Replace(c, '_');
        return s.Replace(" ", "_");
    }
}
