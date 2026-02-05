using System.Net;
using System.Text.Json;
using DocumentFormat.OpenXml; // SpaceProcessingModeValues
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

namespace OpenXmlFunc;

public class DocxProcess
{
    private readonly ILogger _logger;

    public DocxProcess(ILoggerFactory loggerFactory)
    {
        _logger = loggerFactory.CreateLogger<DocxProcess>();
    }

    public record RequestDto(
        string docxBase64,
        string textToAppend
    );

    [Function("DocxProcess")]
    public async Task<HttpResponseData> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
    {
        try
        {
            var body = await new StreamReader(req.Body).ReadToEndAsync();

            var input = JsonSerializer.Deserialize<RequestDto>(body, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (input == null || string.IsNullOrWhiteSpace(input.docxBase64))
            {
                var bad = req.CreateResponse(HttpStatusCode.BadRequest);
                await bad.WriteStringAsync("docxBase64 is required");
                return bad;
            }

            if (string.IsNullOrWhiteSpace(input.textToAppend))
            {
                var bad = req.CreateResponse(HttpStatusCode.BadRequest);
                await bad.WriteStringAsync("textToAppend is required");
                return bad;
            }

            // ✅ Si alguna vez viniera con prefijo data:...;base64, lo limpiamos
            var cleanBase64 = input.docxBase64.Contains(",")
                ? input.docxBase64.Split(',').Last()
                : input.docxBase64;

            byte[] fileBytes = Convert.FromBase64String(cleanBase64);

            // ✅ FIX: stream expandible (evita "Memory stream is not expandable")
            using var ms = new MemoryStream();
            ms.Write(fileBytes, 0, fileBytes.Length);
            ms.Position = 0;

            using (var wordDoc = WordprocessingDocument.Open(ms, true))
            {
                // ✅ Aseguramos que exista MainDocumentPart + Document + Body
                var mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();
                mainPart.Document ??= new Document(new Body());
                mainPart.Document.Body ??= new Body();

                var bodyDoc = mainPart.Document.Body;

                // ------------------------------------------------------------------
                // 1) ✅ Añadir el texto al final (como ya lo hacías)
                // ------------------------------------------------------------------
                var paragraph = new Paragraph();

                var runWithText = new Run(
                    new Text(input.textToAppend)
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    }
                );

                paragraph.Append(runWithText);

                // ✅ Añadimos el párrafo al final del documento
                bodyDoc.AppendChild(paragraph);

                // ------------------------------------------------------------------
                // 2) ✅ Añadir un comentario que señalice que el texto fue añadido por JUANA
                //     Comentario anclado sobre el texto recién insertado.
                // ------------------------------------------------------------------

                // ✅ Crear/asegurar CommentsPart
                var commentsPart = mainPart.GetPartsOfType<WordprocessingCommentsPart>().FirstOrDefault();
                if (commentsPart == null)
                {
                    commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
                    commentsPart.Comments = new Comments();
                }
                else if (commentsPart.Comments == null)
                {
                    commentsPart.Comments = new Comments();
                }

                // ✅ ID único de comentario (evita colisiones)
                var existingIds = commentsPart.Comments.Elements<Comment>()
                    .Select(c => int.TryParse(c.Id?.Value, out var n) ? n : 0);

                var commentId = (existingIds.Any() ? existingIds.Max() : 0) + 1;
                var commentIdStr = commentId.ToString();

                // ✅ Contenido del comentario (lo que se verá en el panel lateral)
                var comment = new Comment()
                {
                    Id = commentIdStr,
                    Author = "JUANA",
                    Date = DateTime.UtcNow
                };

                comment.AppendChild(
                    new Paragraph(
                        new Run(
                            new Text("Texto añadido por JUANA")
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            }
                        )
                    )
                );

                commentsPart.Comments.AppendChild(comment);
                commentsPart.Comments.Save();

                // ✅ “Anclar” el comentario al texto del párrafo insertado:
                //    - Start antes del run
                //    - End después del run
                //    - Reference al final (para que Word lo renderice)
                var start = new CommentRangeStart() { Id = commentIdStr };
                var end = new CommentRangeEnd() { Id = commentIdStr };
                var referenceRun = new Run(new CommentReference() { Id = commentIdStr });

                paragraph.InsertBefore(start, runWithText);
                paragraph.InsertAfter(end, runWithText);
                paragraph.Append(referenceRun);

                // ✅ Guardar documento
                mainPart.Document.Save();
            }

            // ✅ Respuesta JSON con el DOCX modificado en base64
            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/json");

            await response.WriteStringAsync(JsonSerializer.Serialize(new
            {
                docxBase64 = Convert.ToBase64String(ms.ToArray())
            }));

            return response;
        }
        catch (FormatException)
        {
            var bad = req.CreateResponse(HttpStatusCode.BadRequest);
            await bad.WriteStringAsync("docxBase64 is not valid base64.");
            return bad;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing DOCX");

            var error = req.CreateResponse(HttpStatusCode.InternalServerError);
            await error.WriteStringAsync(ex.Message);
            return error;
        }
    }
}
