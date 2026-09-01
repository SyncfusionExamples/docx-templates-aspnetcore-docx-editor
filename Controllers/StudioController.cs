using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using static System.IO.Path;

namespace DocumentTemplateStudioService.Controllers
{
    /// <summary>
    /// StudioController — owns the server-side template catalog and the
    /// common merge-field catalog for the Document Template Studio React
    /// client. The catalog lives at Server-side/wwwroot/Data/templates.json
    /// (and common-merge-fields.json next to it) so the JSON files are
    /// co-located with the .docx files under wwwroot/Templates/. The
    /// client no longer persists the catalog through the Vite dev plugin;
    /// everything is driven from the ASP.NET Core service now.
    ///
    /// Endpoints (all under /api/studio):
    ///   GET    /api/studio/catalog          -> returns templates.json as JSON
    ///   PUT    /api/studio/catalog          -> persists an updated catalog
    ///   GET    /api/studio/common-fields    -> returns the common merge fields
    ///   POST   /api/studio/mergefield       -> adds a custom merge field
    ///   POST   /api/studio/upload           -> writes a new .docx + catalog entry
    ///   DELETE /api/studio/template/{id}   -> removes a template entry
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [EnableCors("AllowAllOrigins")]
    public class StudioController : ControllerBase
    {
        // Server-side locations for the JSON files. Both live under
        // wwwroot/Data/ so the React app's relative `/Data/...` URLs can
        // be served as static files (a useful escape hatch for debugging
        // — see /Data/templates.json in the browser).
        private static readonly string DataFolder =
            Combine(Directory.GetCurrentDirectory(), "wwwroot", "Data");
        private static readonly string CatalogFile =
            Combine(DataFolder, "templates.json");
        private static readonly string CommonFieldsFile =
            Combine(DataFolder, "common-merge-fields.json");
        private static readonly string TemplatesFolder =
            Combine(Directory.GetCurrentDirectory(), "wwwroot", "Templates");

        // Limit the catalog-write surface to safe filenames + safe JSON
        // shapes. The .docx write path was already validated in the
        // existing DocumentEditorController.Save; we only validate the
        // catalog payload here.
        private static readonly Regex SafeId = new("^[A-Za-z0-9_\\-\\.]+$");

        // Coerce a JSON-element value (the shape returned by
        // `JsonSerializer.Deserialize<object>`) into a List<string> of
        // template field keys. Returns an empty list for missing/non-array/
        // non-string-element inputs instead of throwing — the contract is
        // "best-effort merge field extraction", and a malformed entry
        // should not block the caller's update.
        private static List<string> CoerceFieldKeys(JsonElement el)
        {
            var result = new List<string>();
            if (el.ValueKind != JsonValueKind.Array)
            {
                return result;
            }
            foreach (var item in el.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.String)
                {
                    var s = item.GetString();
                    if (!string.IsNullOrEmpty(s) && !result.Contains(s))
                    {
                        result.Add(s);
                    }
                }
            }
            return result;
        }

        public StudioController()
        {
            // Make sure both folders exist at startup. Use lazy init so a
            // first request that needs the folder creates it on demand.
            Directory.CreateDirectory(DataFolder);
            Directory.CreateDirectory(TemplatesFolder);
        }

        // -------------------------------------------------------------------
        // GET /api/studio/catalog
        // Returns the catalog as a JSON array. Returns an empty array if
        // the file is missing (first run).
        // -------------------------------------------------------------------
        [HttpGet]
        [Route("catalog")]
        public IActionResult GetCatalog()
        {
            try
            {
                if (!System.IO.File.Exists(CatalogFile))
                {
                    return Ok(Array.Empty<object>());
                }
                var json = System.IO.File.ReadAllText(CatalogFile, Encoding.UTF8);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return Ok(Array.Empty<object>());
                }
                // Return the JSON as-is. The client parses it as an array.
                return Content(json, "application/json", Encoding.UTF8);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        // -------------------------------------------------------------------
        // PUT /api/studio/catalog
        // Body: JSON array — the full client catalog to persist. The
        // client always sends the entire array; we just write it to disk.
        // Also detects and deletes orphaned .docx files when templates
        // are removed from the catalog.
        // -------------------------------------------------------------------
        [HttpPut]
        [Route("catalog")]
        public async Task<IActionResult> PutCatalog([FromBody] JsonElement body)
        {
            if (body.ValueKind != JsonValueKind.Array)
            {
                return BadRequest(new { error = "Catalog body must be a JSON array." });
            }
            try
            {
                // 1. Load the existing catalog to detect deleted templates
                var existingCatalog = new List<string>(); // Store docxUrls of existing entries
                if (System.IO.File.Exists(CatalogFile))
                {
                    var existingJson = await System.IO.File.ReadAllTextAsync(CatalogFile, Encoding.UTF8);
                    if (!string.IsNullOrWhiteSpace(existingJson))
                    {
                        var existing = JsonSerializer.Deserialize<JsonElement>(existingJson);
                        if (existing.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var e in existing.EnumerateArray())
                            {
                                if (e.TryGetProperty("docxUrl", out var urlEl) &&
                                    urlEl.ValueKind == JsonValueKind.String)
                                {
                                    var url = urlEl.GetString();
                                    if (!string.IsNullOrEmpty(url))
                                    {
                                        existingCatalog.Add(url);
                                    }
                                }
                            }
                        }
                    }
                }

                // 2. Extract docxUrls from the new catalog
                var newCatalog = new HashSet<string>();
                foreach (var e in body.EnumerateArray())
                {
                    if (e.TryGetProperty("docxUrl", out var urlEl) &&
                        urlEl.ValueKind == JsonValueKind.String)
                    {
                        var url = urlEl.GetString();
                        if (!string.IsNullOrEmpty(url))
                        {
                            newCatalog.Add(url);
                        }
                    }
                }

                // 3. Find deleted entries and clean up orphaned .docx files
                foreach (var oldUrl in existingCatalog)
                {
                    if (!newCatalog.Contains(oldUrl))
                    {
                        // This template was deleted — try to delete its .docx file
                        var docxPath = ExtractDocxFilePath(oldUrl);
                        if (!string.IsNullOrEmpty(docxPath))
                        {
                            try
                            {
                                if (System.IO.File.Exists(docxPath))
                                {
                                    System.IO.File.Delete(docxPath);
                                    // eslint-disable-next-line no-console
                                    Console.WriteLine($"[StudioController] Deleted orphaned file: {docxPath}");
                                }
                            }
                            catch (Exception delEx)
                            {
                                // Log but don't fail the catalog save if file deletion fails
                                // eslint-disable-next-line no-console
                                Console.WriteLine($"[StudioController] Failed to delete orphaned file {docxPath}: {delEx.Message}");
                            }
                        }
                    }
                }

                // 4. Write the new catalog
                var pretty = JsonSerializer.Serialize(
                    body,
                    new JsonSerializerOptions { WriteIndented = true });
                await System.IO.File.WriteAllTextAsync(CatalogFile, pretty, Encoding.UTF8);
                return Ok(new { ok = true, count = body.GetArrayLength() });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        // Helper: Extract the file path from a docxUrl.
        // Handles both absolute URLs (http://localhost:5212/Templates/file.docx)
        // and relative paths (/Templates/file.docx).
        private string? ExtractDocxFilePath(string? docxUrl)
        {
            if (string.IsNullOrEmpty(docxUrl))
                return null;

            string? filename = null;

            // If it's an absolute URL, extract the filename from the end
            if (docxUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                docxUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                // Extract just the filename from the URL
                var uri = new Uri(docxUrl);
                filename = System.IO.Path.GetFileName(uri.AbsolutePath);
            }
            else if (docxUrl.StartsWith("/Templates/", StringComparison.OrdinalIgnoreCase))
            {
                // Relative path — extract filename
                filename = System.IO.Path.GetFileName(docxUrl);
            }

            if (string.IsNullOrEmpty(filename))
                return null;

            // Construct the full path: TemplatesFolder / filename
            return Combine(TemplatesFolder, filename);
        }

        // -------------------------------------------------------------------
        // GET /api/studio/common-fields
        // Returns the common (global) custom merge-field catalog.
        // -------------------------------------------------------------------
        [HttpGet]
        [Route("common-fields")]
        public IActionResult GetCommonFields()
        {
            try
            {
                if (!System.IO.File.Exists(CommonFieldsFile))
                {
                    return Ok(new { fields = new Dictionary<string, object>() });
                }
                var json = System.IO.File.ReadAllText(CommonFieldsFile, Encoding.UTF8);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return Ok(new { fields = new Dictionary<string, object>() });
                }
                return Content(json, "application/json", Encoding.UTF8);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        // -------------------------------------------------------------------
        // POST /api/studio/mergefield
        // Form fields: scope ("template" | "common"), key, templateId (when
        // scope = "template"), templateName, templateType, templateDescription.
        // Persists the new field into common-merge-fields.json (common scope)
        // or appends it to the matching catalog entry's fieldKeys (template
        // scope). Mirrors the Vite plugin's previous behavior.
        // -------------------------------------------------------------------
        [HttpPost]
        [Route("mergefield")]
        public async Task<IActionResult> AddMergeField()
        {
            try
            {
                var form = await Request.ReadFormAsync();
                var scope = form["scope"].ToString();
                var key = (form["key"].ToString() ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(key))
                {
                    return BadRequest(new { error = "key is required" });
                }
                if (scope != "template" && scope != "common")
                {
                    return BadRequest(new { error = "scope must be \"template\" or \"common\"" });
                }
                if (scope == "template" && string.IsNullOrEmpty(form["templateId"]))
                {
                    return BadRequest(new { error = "templateId required for template scope" });
                }

                // A field is just its name. Store `true` so downstream
                // lookups (MERGE_FIELDS[key] || customMap[key]) treat
                // the key as recognised.
                var field = true;

                if (scope == "common")
                {
                    Dictionary<string, object> fields;
                    if (System.IO.File.Exists(CommonFieldsFile))
                    {
                        try
                        {
                            var existing = await System.IO.File.ReadAllTextAsync(CommonFieldsFile, Encoding.UTF8);
                            var parsed = JsonSerializer.Deserialize<JsonElement>(existing);
                            if (parsed.TryGetProperty("fields", out var fieldsEl) &&
                                fieldsEl.ValueKind == JsonValueKind.Object)
                            {
                                fields = JsonSerializer.Deserialize<Dictionary<string, object>>(
                                    fieldsEl.GetRawText()) ?? new();
                            }
                            else
                            {
                                fields = new();
                            }
                        }
                        catch
                        {
                            fields = new();
                        }
                    }
                    else
                    {
                        fields = new();
                    }

                    fields[key] = field;
                    var payload = new
                    {
                        fields,
                        updatedAt = DateTime.UtcNow.ToString("o"),
                    };
                    var pretty = JsonSerializer.Serialize(
                        payload,
                        new JsonSerializerOptions { WriteIndented = true });
                    await System.IO.File.WriteAllTextAsync(CommonFieldsFile, pretty, Encoding.UTF8);
                    return Ok(new { ok = true, scope, key, field });
                }

                // scope == "template": update the matching entry inside
                // templates.json. Bootstrap a fresh entry if missing.
                var templateId = form["templateId"].ToString();
                if (!SafeId.IsMatch(templateId))
                {
                    return BadRequest(new { error = "Invalid templateId format." });
                }

                List<JsonElement> catalog = new();
                if (System.IO.File.Exists(CatalogFile))
                {
                    try
                    {
                        var existing = await System.IO.File.ReadAllTextAsync(CatalogFile, Encoding.UTF8);
                        if (!string.IsNullOrWhiteSpace(existing))
                        {
                            var parsed = JsonSerializer.Deserialize<JsonElement>(existing);
                            if (parsed.ValueKind == JsonValueKind.Array)
                            {
                                foreach (var entry in parsed.EnumerateArray())
                                {
                                    catalog.Add(entry);
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Corrupt file — start fresh.
                    }
                }

                // Find or create the entry.
                JsonElement? existingEntry = null;
                var indexToReplace = -1;
                for (var i = 0; i < catalog.Count; i++)
                {
                    if (catalog[i].TryGetProperty("id", out var idEl) &&
                        idEl.ValueKind == JsonValueKind.String &&
                        idEl.GetString() == templateId)
                    {
                        existingEntry = catalog[i];
                        indexToReplace = i;
                        break;
                    }
                }

                var created = existingEntry == null;
                var entryDict = new Dictionary<string, object>();
                if (existingEntry.HasValue)
                {
                    foreach (var prop in existingEntry.Value.EnumerateObject())
                    {
                        entryDict[prop.Name] = JsonSerializer.Deserialize<object>(prop.Value.GetRawText())!;
                    }
                    // `JsonSerializer.Deserialize<object>` turns a JSON array
                    // into a `JsonElement`, not a `List<string>`. Without
                    // this coercion the `is List<string>` check below always
                    // fails, so the existing fieldKeys array is silently
                    // replaced with [] on every save and any user-added
                    // merge field appears to "disappear" the next reload.
                    if (entryDict.TryGetValue("fieldKeys", out var rawKeys) && rawKeys is JsonElement keysEl)
                    {
                        entryDict["fieldKeys"] = CoerceFieldKeys(keysEl);
                    }
                }
                else
                {
                    entryDict["id"] = templateId;
                    entryDict["name"] = string.IsNullOrEmpty(form["templateName"])
                        ? templateId
                        : form["templateName"].ToString();
                    entryDict["type"] = string.IsNullOrEmpty(form["templateType"])
                        ? "General"
                        : form["templateType"].ToString();
                    entryDict["description"] = string.IsNullOrEmpty(form["templateDescription"])
                        ? "Blank letter template."
                        : form["templateDescription"].ToString();
                    entryDict["fieldKeys"] = new List<string>();
                    entryDict["createdAt"] = DateTime.UtcNow.ToString("o");
                }

                var fieldKeys = entryDict.TryGetValue("fieldKeys", out var fk) && fk is List<string> list
                    ? list
                    : new List<string>();
                if (!fieldKeys.Contains(key)) fieldKeys.Add(key);
                entryDict["fieldKeys"] = fieldKeys;
                entryDict["updatedAt"] = DateTime.UtcNow.ToString("o");

                if (indexToReplace >= 0)
                {
                    catalog[indexToReplace] = JsonDocument.Parse(
                        JsonSerializer.Serialize(entryDict)).RootElement;
                }
                else
                {
                    catalog.Add(JsonDocument.Parse(
                        JsonSerializer.Serialize(entryDict)).RootElement);
                }

                var catalogJson = JsonSerializer.Serialize(
                    catalog,
                    new JsonSerializerOptions { WriteIndented = true });
                await System.IO.File.WriteAllTextAsync(CatalogFile, catalogJson, Encoding.UTF8);

                return Ok(new
                {
                    ok = true,
                    scope,
                    key,
                    field,
                    fieldKeys,
                    entry = entryDict,
                    catalog,
                    created,
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        // -------------------------------------------------------------------
        // POST /api/studio/upload
        // Form fields: name, type, description, file (DOCX).
        // Writes the .docx into wwwroot/Templates/ and adds a new entry
        // into templates.json. Returns the new catalog entry so the
        // client can update its in-memory state.
        // -------------------------------------------------------------------
        [HttpPost]
        [Route("upload")]
        [RequestSizeLimit(50_000_000)]
        public async Task<IActionResult> Upload()
        {
            try
            {
                var form = await Request.ReadFormAsync();
                var file = form.Files.GetFile("file");
                if (file == null || file.Length == 0)
                {
                    return BadRequest(new { error = "file is required" });
                }

                var name = (form["name"].ToString() ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(name))
                {
                    name = (file.FileName ?? "template").Replace(".docx", string.Empty, StringComparison.OrdinalIgnoreCase).Trim();
                }
                if (string.IsNullOrEmpty(name))
                {
                    name = "template";
                }
                var type = (form["type"].ToString() ?? "General").Trim();
                if (string.IsNullOrEmpty(type)) type = "General";
                var description = (form["description"].ToString() ?? "Uploaded .docx template").Trim();
                if (string.IsNullOrEmpty(description)) description = "Uploaded .docx template";

                // Signer roles are sent as repeated form fields ("roles")
                // or as a single comma-separated string. The catalog
                // entry stores them as a string[] so the client's
                // TemplateList ROLES column can render them directly.
                var roles = new List<string>();
                foreach (var r in form["roles"])
                {
                    var v = (r ?? string.Empty).Trim();
                    if (!string.IsNullOrEmpty(v) && !roles.Contains(v))
                    {
                        // A single form value may itself be comma-separated
                        // (clients that build one FormData entry from an
                        // array). Split on commas to be safe.
                        foreach (var part in v.Split(',', StringSplitOptions.RemoveEmptyEntries))
                        {
                            var p = part.Trim();
                            if (!string.IsNullOrEmpty(p) && !roles.Contains(p)) roles.Add(p);
                        }
                    }
                }
                var typeForEntry = type;

                // Build a server-side slug: same scheme the client used to
                // produce, so a later Save can reuse it.
                var safeName = name;
                foreach (var c in System.IO.Path.GetInvalidFileNameChars())
                {
                    safeName = safeName.Replace(c, '_');
                }
                safeName = Regex.Replace(safeName, "[^A-Za-z0-9_\\-]+", "_").Trim('_');
                if (string.IsNullOrEmpty(safeName)) safeName = "template";
                var timestamp = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds().ToString("x");
                var slug = $"{safeName}-{timestamp}";
                var docxName = $"{slug}.docx";
                var docxPath = Combine(TemplatesFolder, docxName);

                // Write the .docx bytes to disk.
                using (var fs = new FileStream(docxPath, FileMode.Create, FileAccess.Write))
                {
                    await file.CopyToAsync(fs);
                }

                // Build the new catalog entry.
                var id = $"tpl-{slug}";
                var entry = new Dictionary<string, object>
                {
                    ["id"] = id,
                    ["name"] = name,
                    ["type"] = typeForEntry,
                    ["description"] = description,
                    ["roles"] = roles,
                    // No template-scoped fieldKeys are seeded on upload.
                    // The right-side Merge Fields panel is driven by the
                    // DocumentEditorController.Import envelope, which
                    // unions the .docx's own MERGEFIELDs +
                    // this template's fieldKeys (empty here) + the
                    // global common-merge-fields.json. So the panel
                    // always shows the right fields for the uploaded
                    // document without us hardcoding a list that may
                    // not match what's actually in the .docx.
                    ["fieldKeys"] = new List<string>(),
                    ["docxUrl"] = $"/Templates/{docxName}",
                    ["updatedAt"] = DateTime.UtcNow.ToString("o"),
                };

                // Append to templates.json.
                List<object> catalog = new();
                if (System.IO.File.Exists(CatalogFile))
                {
                    try
                    {
                        var existing = await System.IO.File.ReadAllTextAsync(CatalogFile, Encoding.UTF8);
                        if (!string.IsNullOrWhiteSpace(existing))
                        {
                            var parsed = JsonSerializer.Deserialize<JsonElement>(existing);
                            if (parsed.ValueKind == JsonValueKind.Array)
                            {
                                foreach (var e in parsed.EnumerateArray())
                                {
                                    catalog.Add(JsonSerializer.Deserialize<object>(e.GetRawText())!);
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Corrupt — start fresh.
                    }
                }
                catalog.Add(entry);
                var catalogJson = JsonSerializer.Serialize(
                    catalog,
                    new JsonSerializerOptions { WriteIndented = true });
                await System.IO.File.WriteAllTextAsync(CatalogFile, catalogJson, Encoding.UTF8);

                return Ok(new
                {
                    ok = true,
                    entry,
                    docxFileName = docxName,
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        // -------------------------------------------------------------------
        // DELETE /api/studio/template/{id}
        // Removes the catalog entry for the given id. The .docx file is
        // left in place (orphaned) since the catalog is the source of
        // truth and cleanup is best done by a separate sweep — but the
        // entry disappears from the UI immediately.
        // -------------------------------------------------------------------
        [HttpDelete]
        [Route("template/{id}")]
        public async Task<IActionResult> DeleteTemplate(string id)
        {
            if (!SafeId.IsMatch(id))
            {
                return BadRequest(new { error = "Invalid id format." });
            }
            try
            {
                if (!System.IO.File.Exists(CatalogFile))
                {
                    return NotFound(new { error = "Catalog file not found." });
                }
                var existing = await System.IO.File.ReadAllTextAsync(CatalogFile, Encoding.UTF8);
                if (string.IsNullOrWhiteSpace(existing))
                {
                    return NotFound(new { error = "Catalog is empty." });
                }
                var parsed = JsonSerializer.Deserialize<JsonElement>(existing);
                if (parsed.ValueKind != JsonValueKind.Array)
                {
                    return BadRequest(new { error = "Catalog is not an array." });
                }

                var remaining = new List<object>();
                var found = false;
                foreach (var e in parsed.EnumerateArray())
                {
                    if (e.TryGetProperty("id", out var idEl) &&
                        idEl.ValueKind == JsonValueKind.String &&
                        idEl.GetString() == id)
                    {
                        found = true;
                        continue;
                    }
                    remaining.Add(JsonSerializer.Deserialize<object>(e.GetRawText())!);
                }
                if (!found)
                {
                    return NotFound(new { error = $"Template '{id}' not found in catalog." });
                }
                var catalogJson = JsonSerializer.Serialize(
                    remaining,
                    new JsonSerializerOptions { WriteIndented = true });
                await System.IO.File.WriteAllTextAsync(CatalogFile, catalogJson, Encoding.UTF8);
                return Ok(new { ok = true, removed = id });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }
    }
}
