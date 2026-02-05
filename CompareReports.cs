using System.Net;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
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

    public record RequestDto(
        string baselineDocxBase64,
        string currentDocxBase64,
        string? author,
        string? initials,
        bool? includeDeletedComments,   // default true
        double? tokenChangeThreshold,   // default 0.35
        double? itemMatchThreshold      // default 0.72
    );

    private class Item
    {
        public string Title { get; set; } = "";
        public string Client { get; set; } = "";
        public string Description { get; set; } = "";
        public int AnchorTitle { get; set; } = 0;

        public string Signature => $"{Client} {Title}".Trim();
    }

    public record ChangeItem(string Tag, string Message, int AnchorIndex);

    [Function("CompareReports")]
    public async Task<HttpResponseData> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
    {
        try
        {
            var raw = await new StreamReader(req.Body).ReadToEndAsync();
            var input = JsonSerializer.Deserialize<RequestDto>(raw, new JsonSerializerOptions
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

            var author = string.IsNullOrWhiteSpace(input.author) ? "JUANA" : input.author!;
            var initials = string.IsNullOrWhiteSpace(input.initials) ? "J" : input.initials!;
            var includeDeleted = input.includeDeletedComments ?? true;

            var tokenThreshold = (input.tokenChangeThreshold is > 0 and <= 1) ? input.tokenChangeThreshold.Value : 0.35;
            var matchThreshold = (input.itemMatchThreshold is > 0 and <= 1) ? input.itemMatchThreshold.Value : 0.72;

            // ✅ baseline = v1 (antigua), current = v2 (nueva)
            var v1 = Convert.FromBase64String(input.baselineDocxBase64);
            var v2 = Convert.FromBase64String(input.currentDocxBase64);

            var v1Items = ExtractItems_TitlePlusDescription(v1, includeAnchors: false);
            var v2Items = ExtractItems_TitlePlusDescription(v2, includeAnchors: true);

            var matches = MatchItemsBySimilarity(v1Items, v2Items, matchThreshold);
            var matchedV1 = new HashSet<int>(matches.Values);
            var matchedV2 = new HashSet<int>(matches.Keys);

            var changes = new List<ChangeItem>();

            // ✅ NUEVOS
            for (int i2 = 0; i2 < v2Items.Count; i2++)
            {
                if (matchedV2.Contains(i2)) continue;
                var it = v2Items[i2];

                changes.Add(new ChangeItem(
                    "[NUEVO]",
                    $"{SafeTitle(it)}. Este bloque aparece en v2 y no estaba en v1.",
                    it.AnchorTitle
                ));
            }

            // ✅ ELIMINADOS
            if (includeDeleted)
            {
                for (int i1 = 0; i1 < v1Items.Count; i1++)
                {
                    if (matchedV1.Contains(i1)) continue;
                    var it = v1Items[i1];

                    var anchor = FindAnchorByClient(v2Items, it.Client) ?? 0;

                    changes.Add(new ChangeItem(
                        "[ELIMINADO]",
                        $"{SafeTitle(it)}. Estaba en v1 y ya no aparece en v2.",
                        anchor
                    ));
                }
            }

            // ✅ ACTUALIZACIONES
            foreach (var kv in matches)
            {
                var it2 = v2Items[kv.Key];
                var it1 = v1Items[kv.Value];

                // 1) números (% y €) con dirección correcta
                var numericMsgs = CompareNumbers_Directional(it1.Description, it2.Description);

                foreach (var msg in numericMsgs)
                {
                    changes.Add(new ChangeItem(
                        "[ACTUALIZACION]",
                        $"{SafeTitle(it2)}. {msg}",
                        it2.AnchorTitle
                    ));
                }

                // 2) si no hay números, token delta (umbral)
                if (numericMsgs.Count == 0)
                {
                    var delta = TokenDelta(it1.Description, it2.Description);
                    if (delta.ChangeRatio >= tokenThreshold)
                    {
                        changes.Add(new ChangeItem(
                            "[ACTUALIZACION]",
                            $"{SafeTitle(it2)}. Cambios relevantes en el texto. Añadidos: {delta.AddedSummary}. Eliminados: {delta.RemovedSummary}.",
                            it2.AnchorTitle
                        ));
                    }
                }
            }

            // ✅ limpiar comentarios existentes y dejar solo JUANA
            var updated = ReplaceAllCommentsWithJuana(v2, changes, author, initials);

            var res = req.CreateResponse(HttpStatusCode.OK);
            res.Headers.Add("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            res.Headers.Add("Content-Disposition", "attachment; filename=\"v2_con_comentarios.docx\"");
            await res.Body.WriteAsync(updated, 0, updated.Length);
            return res;
        }
        catch (FormatException)
        {
            var bad = req.CreateResponse(HttpStatusCode.BadRequest);
            await bad.WriteStringAsync("Base64 inválido (baselineDocxBase64 o currentDocxBase64).");
            return bad;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "CompareReports failed");
            var err = req.CreateResponse(HttpStatusCode.InternalServerError);
            await err.WriteStringAsync(ex.Message);
            return err;
        }
    }

    // =========================================================
    // NORMALIZACIÓN (guiones raros –/—)
    // =========================================================
    private static string NormalizeDashes(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        return s.Replace('–', '-').Replace('—', '-');
    }

    private static bool HasDashSeparator(string t)
    {
        t = NormalizeDashes(t);
        // " - " o guiones pegados (CRT-)
        return t.Contains(" - ") || t.Contains('-');
    }

    private static string GetClientFromTitle(string t)
    {
        t = NormalizeDashes(t);

        // preferir split por " - "
        var parts = t.Split(new[] { " - " }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length >= 2) return parts[0].Trim();

        // fallback: primer segmento antes del primer '-'
        var idx = t.IndexOf('-');
        if (idx > 0) return t.Substring(0, idx).Trim();

        return "";
    }

    // =========================================================
    // EXTRACCIÓN: título + descripción (1..3 párrafos después)
    // =========================================================
    private static List<Item> ExtractItems_TitlePlusDescription(byte[] docxBytes, bool includeAnchors)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);

        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null) return new();

        var paragraphs = body.Elements<Paragraph>().ToList();
        var items = new List<Item>();

        for (int i = 0; i < paragraphs.Count - 1; i++)
        {
            var title = NormalizeDashes(Clean(paragraphs[i].InnerText));
            if (string.IsNullOrWhiteSpace(title)) continue;

            if (IsSectionOrCategoryTitle(title)) continue;
            if (!LooksLikeRealItemTitle(title)) continue;

            // buscar descripción en los próximos 3 párrafos
            string desc = "";
            for (int j = i + 1; j <= Math.Min(i + 3, paragraphs.Count - 1); j++)
            {
                var candidate = Clean(paragraphs[j].InnerText);
                if (string.IsNullOrWhiteSpace(candidate)) continue;
                if (IsSectionOrCategoryTitle(candidate)) break;

                if (LooksLikeDescription(candidate))
                {
                    desc = candidate;
                    break;
                }
            }

            if (string.IsNullOrWhiteSpace(desc)) continue;

            var it = new Item
            {
                Title = title,
                Description = desc,
                AnchorTitle = includeAnchors ? i : 0,
                Client = GetClientFromTitle(title)
            };

            items.Add(it);
        }

        return items;
    }

    private static bool LooksLikeDescription(string text)
    {
        return text.Length >= 80 && text.Contains(' ');
    }

    private static bool LooksLikeRealItemTitle(string t)
    {
        t = NormalizeDashes(t);

        if (!HasDashSeparator(t)) return false;
        if (t.Length < 12 || t.Length > 220) return false;

        // numeración tipo 1.1
        if (Regex.IsMatch(t, @"^\d+(\.\d+)*\s+", RegexOptions.IgnoreCase)) return false;

        // filtrar categorías típicas
        var low = t.ToLowerInvariant();
        var forbidden = new[]
        {
            "muy importante", "relevante", "ordinaria",
            "estado crítico", "otros proyectos",
            "oportunidades comerciales", "otra actividad comercial",
            "visión general", "proyectos en ejecución"
        };
        if (forbidden.Any(f => low.Contains(f))) return false;

        // evita títulos demasiado “genéricos”
        if (low is "relevante" or "ordinaria" or "muy importante") return false;

        return true;
    }

    private static bool IsSectionOrCategoryTitle(string t)
    {
        var up = NormalizeDashes(t).Trim().ToUpperInvariant();

        if (Regex.IsMatch(up, @"^\d+(\.\d+)*\s+", RegexOptions.IgnoreCase)) return true;

        var exact = new HashSet<string>
        {
            "ESTADO CRÍTICO",
            "OTROS PROYECTOS",
            "MUY IMPORTANTE",
            "RELEVANTE",
            "ORDINARIA"
        };
        if (exact.Contains(up)) return true;

        if (up.Contains("PROYECTOS EN EJECUCIÓN")) return true;
        if (up.Contains("OPORTUNIDADES COMERCIALES")) return true;
        if (up.Contains("OTRA ACTIVIDAD COMERCIAL")) return true;
        if (up.Contains("VISIÓN GENERAL")) return true;

        return false;
    }

    private static string Clean(string s)
    {
        s = (s ?? "").Replace("**", "");
        return Regex.Replace(s, @"\s+", " ").Trim();
    }

    private static string Norm(string s)
    {
        s = (s ?? "").ToLowerInvariant().Trim();
        s = Regex.Replace(s, @"\s+", " ");
        s = Regex.Replace(s, @"[^\p{L}\p{N}\s]", "");
        return s;
    }

    private static string SafeTitle(Item it)
    {
        var t = it.Title?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(t)) t = "Sin título";
        return t;
    }

    private static int? FindAnchorByClient(List<Item> items, string client)
    {
        var c = Norm(client);
        if (string.IsNullOrWhiteSpace(c)) return null;

        foreach (var it in items)
            if (!string.IsNullOrWhiteSpace(it.Client) && Norm(it.Client) == c)
                return it.AnchorTitle;

        return null;
    }

    // =========================================================
    // MATCHING POR SIMILITUD (sin “tokens fuertes” agresivos)
    // =========================================================
    private static Dictionary<int, int> MatchItemsBySimilarity(List<Item> v1, List<Item> v2, double threshold)
    {
        var map = new Dictionary<int, int>(); // v2Index -> v1Index
        var usedV1 = new HashSet<int>();

        for (int i2 = 0; i2 < v2.Count; i2++)
        {
            double best = 0;
            int bestI1 = -1;

            for (int i1 = 0; i1 < v1.Count; i1++)
            {
                if (usedV1.Contains(i1)) continue;

                // similitud principalmente por título (más estable que mezclar description)
                var scoreTitle = Similarity(v2[i2].Title, v1[i1].Title);
                var scoreSig = Similarity(v2[i2].Signature, v1[i1].Signature);

                // mezcla: título pesa más
                var score = (0.75 * scoreTitle) + (0.25 * scoreSig);

                // bonus si cliente coincide (cuando lo tenemos)
                if (!string.IsNullOrWhiteSpace(v2[i2].Client) &&
                    !string.IsNullOrWhiteSpace(v1[i1].Client) &&
                    Norm(v2[i2].Client) == Norm(v1[i1].Client))
                    score += 0.08;

                if (score > best)
                {
                    best = score;
                    bestI1 = i1;
                }
            }

            if (bestI1 >= 0 && best >= threshold)
            {
                map[i2] = bestI1;
                usedV1.Add(bestI1);
            }
        }

        return map;
    }

    private static double Similarity(string a, string b)
    {
        var A = Tokenize(a);
        var B = Tokenize(b);
        if (A.Count == 0 && B.Count == 0) return 1;
        if (A.Count == 0 || B.Count == 0) return 0;

        var inter = A.Intersect(B).Count();
        var uni = A.Union(B).Count();
        return uni == 0 ? 0 : (double)inter / uni;
    }

    private static HashSet<string> Tokenize(string s)
    {
        s = Norm(s);

        var stop = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "de","la","el","y","en","a","por","para","con","sin","un","una","los","las",
            "que","se","su","al","del","lo","como","más","menos","muy","ya","no","si",
            "proyecto","squad","producto","vertical","estado","otros"
        };

        var parts = s.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var p in parts)
        {
            if (p.Length < 4) continue;
            if (stop.Contains(p)) continue;
            set.Add(p);
        }
        return set;
    }

    // =========================================================
    // ACTUALIZACIONES NUMÉRICAS (direccional)
    // =========================================================
    private static List<string> CompareNumbers_Directional(string oldText, string newText)
    {
        var msgs = new List<string>();

        // porcentajes: si cambia el conjunto, reportamos old -> new
        var oldPct = ExtractPercentages(oldText).OrderBy(x => x).ToList();
        var newPct = ExtractPercentages(newText).OrderBy(x => x).ToList();

        if (oldPct.Count > 0 || newPct.Count > 0)
        {
            if (!oldPct.SequenceEqual(newPct))
            {
                var o = oldPct.Count > 0 ? string.Join(", ", oldPct.Select(x => $"{x}%")) : "—";
                var n = newPct.Count > 0 ? string.Join(", ", newPct.Select(x => $"{x}%")) : "—";
                msgs.Add($"Cambio de porcentaje: {o} → {n}.");
            }
        }

        // euros: reporta cambio de tokens € (simple)
        var oldEur = ExtractEuros(oldText).OrderBy(x => x).ToList();
        var newEur = ExtractEuros(newText).OrderBy(x => x).ToList();

        if (oldEur.Count > 0 || newEur.Count > 0)
        {
            if (!oldEur.SequenceEqual(newEur))
            {
                var o = oldEur.Count > 0 ? string.Join(", ", oldEur) : "—";
                var n = newEur.Count > 0 ? string.Join(", ", newEur) : "—";
                msgs.Add($"Cambio de importe: {o} → {n}.");
            }
        }

        return msgs;
    }

    private static List<int> ExtractPercentages(string text)
    {
        var list = new List<int>();
        foreach (Match m in Regex.Matches(text, @"(\d{1,3})\s*%", RegexOptions.IgnoreCase))
        {
            if (int.TryParse(m.Groups[1].Value, out var pct))
                list.Add(pct);
        }
        return list.Distinct().ToList();
    }

    private static List<string> ExtractEuros(string text)
    {
        var list = new List<string>();
        foreach (Match m in Regex.Matches(text, @"(\d{1,3}(\.\d{3})+|\d+)(,\d+)?\s*(k€|k|€)", RegexOptions.IgnoreCase))
        {
            list.Add(m.Value.Trim());
        }
        return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    // =========================================================
    // TOKEN DELTA (fallback)
    // =========================================================
    private record TokenDeltaResult(double ChangeRatio, string AddedSummary, string RemovedSummary);

    private static TokenDeltaResult TokenDelta(string a, string b)
    {
        var A = Tokenize(a);
        var B = Tokenize(b);

        if (A.Count == 0 && B.Count == 0) return new(0, "—", "—");

        var added = B.Except(A).ToList();
        var removed = A.Except(B).ToList();
        var union = A.Union(B).Count();
        var ratio = union == 0 ? 0 : (double)(added.Count + removed.Count) / union;

        return new(
            ratio,
            added.Count == 0 ? "—" : string.Join(", ", added.Take(10)),
            removed.Count == 0 ? "—" : string.Join(", ", removed.Take(10))
        );
    }

    // =========================================================
    // BORRAR TODOS LOS COMENTARIOS Y CREAR SOLO JUANA
    // =========================================================
    private static byte[] ReplaceAllCommentsWithJuana(byte[] v2Bytes, List<ChangeItem> changes, string author, string initials)
    {
        using var ms = new MemoryStream();
        ms.Write(v2Bytes, 0, v2Bytes.Length);
        ms.Position = 0;

        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var main = doc.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart no encontrado.");
            main.Document ??= new Document(new Body());
            main.Document.Body ??= new Body();

            RemoveAllCommentAnchors(main.Document);

            if (main.WordprocessingCommentsPart != null)
                main.DeletePart(main.WordprocessingCommentsPart);

            var commentsPart = main.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments();

            var paragraphs = main.Document.Body.Elements<Paragraph>().ToList();
            if (paragraphs.Count == 0)
            {
                main.Document.Body.AppendChild(new Paragraph(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve })));
                paragraphs = main.Document.Body.Elements<Paragraph>().ToList();
            }

            int id = 1;
            foreach (var ch in changes)
            {
                var idx = Math.Clamp(ch.AnchorIndex, 0, paragraphs.Count - 1);
                var p = paragraphs[idx];

                var idStr = id.ToString();
                id++;

                var comment = new Comment
                {
                    Id = idStr,
                    Author = author,
                    Initials = initials,
                    Date = DateTime.Now
                };

                comment.AppendChild(new Paragraph(
                    new Run(new Text($"{ch.Tag} {ch.Message}") { Space = SpaceProcessingModeValues.Preserve })
                ));

                commentsPart.Comments.AppendChild(comment);
                AnchorCommentToParagraph(p, idStr);
            }

            commentsPart.Comments.Save();
            main.Document.Save();
        }

        return ms.ToArray();
    }

    private static void RemoveAllCommentAnchors(Document document)
    {
        var body = document.Body;
        if (body == null) return;

        foreach (var el in body.Descendants<CommentRangeStart>().ToList()) el.Remove();
        foreach (var el in body.Descendants<CommentRangeEnd>().ToList()) el.Remove();
        foreach (var el in body.Descendants<CommentReference>().ToList()) el.Remove();
    }

    private static void AnchorCommentToParagraph(Paragraph paragraph, string commentId)
    {
        var firstRun = paragraph.Elements<Run>().FirstOrDefault();
        if (firstRun == null)
        {
            firstRun = new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(firstRun);
        }

        paragraph.InsertBefore(new CommentRangeStart { Id = commentId }, firstRun);
        var end = new CommentRangeEnd { Id = commentId };
        paragraph.InsertAfter(end, firstRun);
        paragraph.InsertAfter(new Run(new CommentReference { Id = commentId }), end);
    }
}
