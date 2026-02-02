using System.Net;
using System.Text.Json;
using DocumentFormat.OpenXml; // ✅ necesario para SpaceProcessingModeValues
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

            byte[] fileBytes = Convert.FromBase64String(input.docxBase64);

            using var ms = new MemoryStream(fileBytes);
            ms.Position = 0;

            using (var wordDoc = WordprocessingDocument.Open(ms, true))
            {
                // ✅ Aseguramos que exista MainDocumentPart + Document + Body
                var mainPart = wordDoc.MainDocumentPart ?? wordDoc.AddMainDocumentPart();
                mainPart.Document ??= new Document(new Body());
                mainPart.Document.Body ??= new Body();

                var bodyDoc = mainPart.Document.Body;

                // ✅ Crear párrafo nuevo (preserva espacios)
                var paragraph = new Paragraph(
                    new Run(
                        new Text(input.textToAppend)
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        }
                    )
                );

                // ✅ Añadir al final
                bodyDoc.AppendChild(paragraph);
                mainPart.Document.Save();
            }

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
