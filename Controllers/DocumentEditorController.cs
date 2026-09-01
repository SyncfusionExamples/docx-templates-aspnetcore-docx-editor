using System.Linq;
using System.Net;
using System.Text;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;
using Syncfusion.DocIORenderer;
using Syncfusion.EJ2.DocumentEditor;
using Syncfusion.EJ2.SpellChecker;
using Syncfusion.Pdf;
using static System.Convert;
using static System.IO.Path;
using static Newtonsoft.Json.Linq.JObject;
using static Syncfusion.EJ2.DocumentEditor.WordDocument;
using WDocument = Syncfusion.DocIO.DLS.WordDocument;
using WFormatType = Syncfusion.DocIO.FormatType;

namespace DocumentTemplateStudioService.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DocumentEditorController : ControllerBase
    {
        private readonly string templatePath;

        // Security Constants
        private static readonly string[] ALLOWED_EXTENSIONS = { ".docx", ".doc", ".rtf", ".txt", ".xml", ".html", ".dotx", ".docm", ".dotm" };

        public DocumentEditorController()
        {
            templatePath = Path.Combine(Program.webRootPath, "Templates");
        }

        /// <summary>
        /// Validates file extension against whitelist
        /// </summary>
        private bool IsAllowedExtension(string fileName)
        {
            string ext = Path.GetExtension(fileName).ToLower();
            return ALLOWED_EXTENSIONS.Contains(ext);
        }

        // Loads the global common-merge-fields catalog from
        // wwwroot/Data/common-merge-fields.json. Returns the field
        // names as a list (NOT the `fields` object — we only need the
        // keys for the merge-field union). Missing file or any
        // parse error returns an empty list so callers never throw.
        private static List<string> LoadCommonMergeFieldNames()
        {
            try
            {
                var dataFolder = Combine(Directory.GetCurrentDirectory(), "wwwroot", "Data");
                var file = Combine(dataFolder, "common-merge-fields.json");
                if (!System.IO.File.Exists(file)) return new List<string>();
                var json = System.IO.File.ReadAllText(file, Encoding.UTF8);
                if (string.IsNullOrWhiteSpace(json)) return new List<string>();
                var doc = Newtonsoft.Json.Linq.JObject.Parse(json);
                var fieldsToken = doc["fields"] as Newtonsoft.Json.Linq.JObject;
                if (fieldsToken == null) return new List<string>();
                return fieldsToken.Properties()
                    .Select(p => p.Name)
                    .Where(n => !string.IsNullOrWhiteSpace(n))
                    .ToList();
            }
            catch
            {
                return new List<string>();
            }
        }

        // Loads the per-template merge-field keys from
        // wwwroot/Data/templates.json. Returns the `fieldKeys` array
        // for the catalog entry whose `id` matches the given
        // templateId. Missing file, missing entry, or any parse error
        // returns an empty list so callers never throw — a template
        // with no persisted fieldKeys simply contributes nothing to
        // the union. We read the file fresh on every Import call so
        // newly-added template-scoped fields are picked up without a
        // service restart.
        private static List<string> LoadTemplateFieldKeys(string templateId)
        {
            if (string.IsNullOrWhiteSpace(templateId)) return new List<string>();
            try
            {
                var dataFolder = Combine(Directory.GetCurrentDirectory(), "wwwroot", "Data");
                var file = Combine(dataFolder, "templates.json");
                if (!System.IO.File.Exists(file)) return new List<string>();
                var json = System.IO.File.ReadAllText(file, Encoding.UTF8);
                if (string.IsNullOrWhiteSpace(json)) return new List<string>();
                var arr = Newtonsoft.Json.Linq.JArray.Parse(json);
                foreach (var entry in arr)
                {
                    var idToken = entry["id"];
                    if (idToken == null) continue;
                    if (string.Equals(idToken.ToString(), templateId, StringComparison.Ordinal))
                    {
                        var keysToken = entry["fieldKeys"] as Newtonsoft.Json.Linq.JArray;
                        if (keysToken == null) return new List<string>();
                        return keysToken
                            .Select(t => t?.ToString())
                            .Where(n => !string.IsNullOrWhiteSpace(n))
                            .ToList();
                    }
                }
                return new List<string>();
            }
            catch
            {
                return new List<string>();
            }
        }

        // Build the deduped, doc-first union of merge-field names
        // returned to the client. THREE sources are combined in a
        // fixed precedence order so the panel renders a stable list
        // regardless of which source happens to contain a given key:
        //   1. MERGEFIELDs actually present in the .docx body
        //      (enumerated by DocIO's MailMerge API)
        //   2. template.fieldKeys from wwwroot/Data/templates.json
        //      (per-template-scoped keys persisted by Add Field /
        //      seeded on upload)
        //   3. keys from wwwroot/Data/common-merge-fields.json
        //      (global keys shared across all templates)
        // Each source is appended in order; deduplication via a
        // HashSet means a key that appears in multiple sources is
        // kept only at its FIRST occurrence (doc first, then
        // template, then common). This is the single source of
        // truth the right-side Merge Fields panel renders — the
        // client no longer unions with anything locally.
        private static string[] BuildMergeFieldsArray(string[] docFieldNames, string templateId)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var ordered = new List<string>();
            if (docFieldNames != null)
            {
                foreach (var name in docFieldNames)
                {
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (seen.Add(name)) ordered.Add(name);
                }
            }
            foreach (var name in LoadTemplateFieldKeys(templateId))
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                if (seen.Add(name)) ordered.Add(name);
            }
            foreach (var name in LoadCommonMergeFieldNames())
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                if (seen.Add(name)) ordered.Add(name);
            }
            return ordered.ToArray();
        }

        // Handles importing a document file and converting it to JSON format.
        // Returns an envelope { sfdt, mergeFields } where mergeFields is
        // the deduped union of THREE sources, in precedence order:
        //   1. MERGEFIELDs present in the uploaded .docx (enumerated via
        //      DocIO's MailMerge API)
        //   2. template.fieldKeys for the template whose `id` matches the
        //      optional `TemplateId` form field (looked up in
        //      wwwroot/Data/templates.json)
        //   3. keys from wwwroot/Data/common-merge-fields.json
        // The right-side Merge Fields panel renders this list as-is — the
        // client does no further unioning. The two-step load (DocIO
        // first for MailMerge, EJ2 second for SFDT) avoids a second
        // download/round-trip and keeps the response payload
        // self-contained. `TemplateId` is OPTIONAL so the endpoint still
        // works for callers that don't have a catalog entry (e.g. a
        // brand-new upload that hasn't been committed yet).
        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("Import")]
        public object Import(IFormCollection data)
        {
            if (data.Files.Count == 0)
                return null;
            Stream stream1 = new MemoryStream();
            IFormFile file = data.Files[0];
            int index = file.FileName.LastIndexOf('.');
            string type = index > -1 && index < file.FileName.Length - 1 ?
                file.FileName.Substring(index) : ".docx";
            file.CopyTo(stream1);
            stream1.Position = 0;

            // Optional template id passed by the client so the server can
            // look up the matching per-template fieldKeys in templates.json.
            // Absent for uploads that aren't tied to a catalog entry yet.
            string templateId = data.TryGetValue("TemplateId", out var tid) && tid.Count > 0
                ? tid.ToString()
                : null;

            // Use Syncfusion DocIO (wDocument) to enumerate the merge
            // field names actually present in the uploaded .docx, then
            // union with template.fieldKeys + the global common-merge-fields
            // catalog.
            WDocument docxDoc = new WDocument(stream1, WFormatType.Docx);
            string[] docFieldNames = docxDoc.MailMerge.GetMergeFieldNames() ?? new string[0];
            docxDoc.Close();
            string[] mergeFieldNames = BuildMergeFieldsArray(docFieldNames, templateId);

            // Rewind the same memory stream and load as EJ2 DocumentEditor
            // WordDocument so we can serialize to SFDT JSON for the client.
            stream1.Position = 0;
            WordDocument document = WordDocument.Load(stream1, GetFormatType(type.ToLower()));
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
            document.Dispose();
            return new { sfdt = json, mergeFields = mergeFieldNames };
        }
        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("ImportFileURL")]
        public object ImportFileURL([FromBody] FileUrlInfo param)
        {
            using (WebClient client = new WebClient())
            {
                MemoryStream stream = new MemoryStream(client.DownloadData(param.fileUrl));
                // Use Syncfusion DocIO (wDocument) to enumerate the merge
                // field names actually present in the .docx, then union
                // with template.fieldKeys + the global common-merge-fields
                // catalog. Doc fields are returned first (in the order
                // Syncfusion reported them), then template-scoped keys,
                // then any common keys not already present — so the
                // right-side panel renders the deduped list as-is without
                // any client-side unioning.
                WDocument docxDoc = new WDocument(stream, WFormatType.Docx);
                string[] docFieldNames = docxDoc.MailMerge.GetMergeFieldNames() ?? new string[0];
                docxDoc.Close();
                string[] mergeFieldNames = BuildMergeFieldsArray(docFieldNames, param.templateId);

                // Rewind the same memory stream and load as EJ2
                // DocumentEditor WordDocument so we can serialize to SFDT
                // JSON for the client. The two-step load (DocIO first for
                // MailMerge, EJ2 second for SFDT) avoids a separate
                // download/round-trip and keeps the response payload
                // self-contained.
                stream.Position = 0;
                WordDocument document = WordDocument.Load(stream, FormatType.Docx);
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
                document.Dispose();
                stream.Dispose();
                return new { sfdt = json, mergeFields = mergeFieldNames };
            }
        }
        public class FileUrlInfo
        {
            public string fileUrl { get; set; }
            // Optional: the id of the catalog entry being opened, so the
            // server can look up the matching template.fieldKeys in
            // templates.json and union them into the mergeFields response.
            // Absent for callers that don't have a catalog entry.
            public string templateId { get; set; }
        }

        // Representing parameters for clipboard operations
        public class CustomRestrictParameter
        {
            public string? passwordBase64 { get; set; }
            public string? saltBase64 { get; set; }
            public int spinCount { get; set; }
        }

        // Handles document editing restrictions
        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("RestrictEditing")]
        public string[]? RestrictEditing([FromBody] CustomRestrictParameter param)
        {
            if (param.passwordBase64 == "" && param.passwordBase64 == null)
                return null;
            return WordDocument.ComputeHash(param.passwordBase64, param.saltBase64, param.spinCount);
        }

        // Determines the document format based on file extension
        internal static FormatType GetFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            switch (format.ToLower())
            {
                case ".dotx":
                case ".docx":
                case ".docm":
                case ".dotm":
                    return FormatType.Docx;
                case ".dot":
                case ".doc":
                    return FormatType.Doc;
                case ".rtf":
                    return FormatType.Rtf;
                case ".txt":
                    return FormatType.Txt;
                case ".xml":
                    return FormatType.WordML;
                case ".html":
                    return FormatType.Html;
                default:
                    throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }
        }

        // Determines the document format type specifically for Word formats
        internal static WFormatType GetWFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            switch (format.ToLower())
            {
                case ".dotx":
                    return WFormatType.Dotx;
                case ".docx":
                    return WFormatType.Docx;
                case ".docm":
                    return WFormatType.Docm;
                case ".dotm":
                    return WFormatType.Dotm;
                case ".dot":
                    return WFormatType.Dot;
                case ".doc":
                    return WFormatType.Doc;
                case ".rtf":
                    return WFormatType.Rtf;
                case ".html":
                    return WFormatType.Html;
                case ".txt":
                    return WFormatType.Txt;
                case ".xml":
                    return WFormatType.WordML;
                case ".odt":
                    return WFormatType.Odt;
                default:
                    throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }
        }

        public class SaveParameter
        {
            public string Content { get; set; }
            public string FileName { get; set; }
            public string Format { get; set; }
            // Optional: catalog entry id. When supplied (the Publish action),
            // the matching templates.json entry gets its docxUrl + updatedAt
            // refreshed so the dashboard reflects the published document.
            public string TemplateId { get; set; }
        }

         /// <summary>
        /// Validates file name for path traversal and special characters
        /// </summary>
        private bool IsValidFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return false;

            // Check for path traversal attempts
            if (fileName.Contains("..") || fileName.Contains("/") || fileName.Contains("\\"))
                return false;

            // Check for invalid filename characters
            return !Path.GetInvalidFileNameChars()
                .Any(c => fileName.Contains(c));
        }

        private string RetrieveFileType(string name)
        {
            int index = name.LastIndexOf('.');
            string format = index > -1 && index < name.Length - 1 ?
                name.Substring(index) : ".doc";
            return format;
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("Export")]
        public FileStreamResult Export([FromBody] SaveParameter data)
        {
            string fileName = data.FileName;
            string format = RetrieveFileType(string.IsNullOrEmpty(data.Format) ? fileName : data.Format);
            if (string.IsNullOrEmpty(fileName))
            {
                fileName = "Document1.docx";
            }
            // Validate filename
            if (!IsValidFileName(fileName))
            {
                throw new ArgumentException("Invalid filename format.");
            }

            WDocument document;
            if (format.ToLower() == ".pdf")
            {
                Stream stream = WordDocument.Save(data.Content, FormatType.Docx);
                document = new Syncfusion.DocIO.DLS.WordDocument(stream, Syncfusion.DocIO.FormatType.Docx);
            }
            else
            {
                document = WordDocument.Save(data.Content);
            }

            return SaveDocument(document, format, fileName);
        }

        private FileStreamResult SaveDocument(WDocument document, string format, string fileName)
        {
            // Validate filename before saving
            if (!IsValidFileName(fileName))
            {
                throw new ArgumentException("Invalid filename format.");
            }

            Stream stream = new MemoryStream();
            string contentType = "";
            try
            {
                if (format.ToLower() == ".pdf")
                {
                    contentType = "application/pdf";
                    DocIORenderer render = new DocIORenderer();
                    PdfDocument pdfDocument = render.ConvertToPDF(document);
                    stream = new MemoryStream();
                    pdfDocument.Save(stream);
                    pdfDocument.Close();
                }
                else
                {
                    WFormatType type = GetWFormatType(format);
                    switch (type)
                    {
                        case WFormatType.Rtf:
                            contentType = "application/rtf";
                            break;
                        case WFormatType.WordML:
                            contentType = "application/xml";
                            break;
                        case WFormatType.Html:
                            contentType = "application/html";
                            break;
                        case WFormatType.Dotx:
                            contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.template";
                            break;
                        case WFormatType.Docx:
                            contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                            break;
                        case WFormatType.Doc:
                            contentType = "application/msword";
                            break;
                        case WFormatType.Dot:
                            contentType = "application/msword";
                            break;
                        case WFormatType.Odt:
                            contentType = "application/vnd.oasis.opendocument.text";
                            break;
                        case WFormatType.Markdown:
                            contentType = "text/markdown";
                            break;
                    }
                    document.Save(stream, type);
                }

                document.Close();

                stream.Position = 0;
                return new FileStreamResult(stream, contentType)
                {
                    FileDownloadName = fileName
                };
            }
            catch (Exception ex)
            {
                stream?.Dispose();
                throw new InvalidOperationException("Error saving documents: " + ex.Message, ex);
            }
        }

         [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("Save")]
        public IActionResult Save([FromBody] SaveParameter data)
        {
            string name = data.FileName;
            string format = !string.IsNullOrWhiteSpace(data.Format)
                ? (data.Format.StartsWith(".") ? data.Format : "." + data.Format)
                : RetrieveFileType(name);

            if (string.IsNullOrEmpty(name))
            {
                name = "Document1" + format;
            }

            // Validate filename
            if (!IsValidFileName(name))
            {
                return BadRequest(new { error = "Invalid filename format." });
            }

            string existingExt = Path.GetExtension(name);
            if (string.IsNullOrEmpty(existingExt) ||
                !string.Equals(existingExt, format, StringComparison.OrdinalIgnoreCase))
            {
                name = name + format;
            }

            try
            {
                WDocument document = WordDocument.Save(data.Content);
                string savePath = Path.Combine(templatePath, Path.GetFileName(name));

                // Ensure templatePath is within allowed directory
                string fullPath = Path.GetFullPath(savePath);
                string allowedPath = Path.GetFullPath(templatePath);
                if (!fullPath.StartsWith(allowedPath, StringComparison.OrdinalIgnoreCase))
                {
                    return StatusCode((int)HttpStatusCode.Forbidden,
                        new { error = "Access denied: Cannot save file outside allowed directory." });
                }

                // FileMode.Create truncates any existing file at savePath so
                // a second Save with the same FileName cleanly OVERWRITES
                // the prior .docx (requirement: "Replace existing file with
                // same name — do not save as new file").
                using (FileStream fileStream = new FileStream(savePath, FileMode.Create, FileAccess.ReadWrite))
                {
                    document.Save(fileStream, GetWFormatType(format));
                }

                document.Close();
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = "Error saving document: " + ex.Message });
            }

            // Publish half of the action: link the saved .docx back into
            // templates.json so the dashboard + a later editor open-load
            // resolve it. Skipped silently when TemplateId is absent (a
            // plain Save from a caller without a catalog entry).
            var docxUrl = "/Templates/" + Path.GetFileName(name);
            if (!string.IsNullOrWhiteSpace(data.TemplateId))
            {
                CatalogService? catalogService = null;
                try
                {
                    catalogService = HttpContext.RequestServices
                        .GetService(typeof(CatalogService)) as CatalogService;
                }
                catch { /* best-effort below */ }
                if (catalogService != null)
                {
                    var updated = catalogService.MarkTemplatePublished(data.TemplateId, docxUrl);
                    if (updated == null)
                    {
                        return Ok(new { ok = true, savedPath = docxUrl, published = false,
                            warning = "Saved to disk, but the catalog entry was not found." });
                    }
                    return Ok(new { ok = true, savedPath = docxUrl, published = true, entry = updated });
                }
            }
            return Ok(new { ok = true, savedPath = docxUrl, published = false });
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("MailMerge")]
        public string MailMerge([FromBody] ExportData exportData)
        {
            byte[] data;
            // Input validation
            if (exportData == null || string.IsNullOrEmpty(exportData.documentData))
            {
                throw new ArgumentException("Document data cannot be null or empty.");
            }

            try
            {
                string cleanBase64 = exportData.documentData.Contains(',') ? exportData.documentData.Split(',')[1] : exportData.documentData;
                data = Convert.FromBase64String(cleanBase64);
                using (MemoryStream stream = new MemoryStream())
                {
                    stream.Write(data, 0, data.Length);
                    stream.Position = 0;

                    using (Syncfusion.DocIO.DLS.WordDocument document = new Syncfusion.DocIO.DLS.WordDocument(stream, Syncfusion.DocIO.FormatType.Docx))
                    {
                        document.MailMerge.RemoveEmptyGroup = true;
                        document.MailMerge.RemoveEmptyParagraphs = true;
                        document.MailMerge.ClearFields = true;
                        document.MailMerge.Execute(GetJsonData(exportData.mailMergeData));
                        document.Save(stream, Syncfusion.DocIO.FormatType.Docx);
                    }

                    stream.Position = 0;
                    Syncfusion.EJ2.DocumentEditor.WordDocument wordDocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(stream, Syncfusion.EJ2.DocumentEditor.FormatType.Docx);
                    string sfdtText = Newtonsoft.Json.JsonConvert.SerializeObject(wordDocument);
                    wordDocument?.Dispose();
                    return sfdtText;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Error processing mail merge: " + ex.Message, ex);
            }
        }

        public class ExportData
        {
            public string fileName { get; set; }
            public string documentData { get; set; }
            public string mailMergeData { get; set; }
        }

        #region Helper methods for Mail Merge JSON Data
        /// <summary>
        /// Prepares the data table from JSON data for processing.
        /// </summary>
        private static List<object> GetJsonData(string mailMergeData)
        {
            //Reads the JSON object from JSON file.
            JObject jsonObject = JObject.Parse(mailMergeData);
            //Converts JSON object to Dictionary.
            IDictionary<string, object> data = GetData(jsonObject);
            return data.Values.First() as List<object>;
        }

        /// <summary>
        /// Gets data from JSON object.
        /// </summary>
        /// <param name="jsonObject">JSON object.</param>
        /// <returns>Dictionary of data.</returns>
        private static IDictionary<string, object> GetData(JObject jsonObject)
        {
            Dictionary<string, object> dictionary = new Dictionary<string, object>();
            foreach (var item in jsonObject)
            {
                object keyValue = null;
                if (item.Value is JArray)
                    keyValue = GetData((JArray)item.Value);
                else if (item.Value is JToken)
                    keyValue = ((JToken)item.Value).ToObject<string>();
                dictionary.Add(item.Key, keyValue);
            }
            return dictionary;
        }
        /// <summary>
        /// Gets array of items from JSON array.
        /// </summary>
        /// <param name="jArray">JSON array.</param>
        /// <returns>List of objects.</returns>
        private static List<object> GetData(JArray jArray)
        {
            List<object> jArrayItems = new List<object>();
            foreach (var item in jArray)
            {
                object keyValue = null;
                if (item is JObject)
                    keyValue = GetData((JObject)item);
                jArrayItems.Add(keyValue);
            }
            return jArrayItems;
        }
        #endregion
    }
}
