using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.AspNetCore.Hosting;
using static System.IO.Path;

namespace DocumentTemplateStudioService
{
    // CatalogService — single source of truth for the server-side
    // catalog state. Both Razor Pages and the existing API controllers
    // go through this service so the same JSON files
    // (wwwroot/Data/templates.json + common-merge-fields.json) are
    // read/written consistently. Scoped lifetime — one instance per
    // HTTP request — so reads inside the same request see the same
    // in-memory snapshot.
    //
    // The catalog shape is intentionally kept identical to what the old
    // React client consumed: a JSON array of objects with camelCase
    // property names (id, name, type, description, fieldKeys, docxUrl,
    // uploadedAt, updatedAt, roles). The shared JsonOptions below
    // enforce camelCase + ignore the display-only helpers
    // (CreatedDisplay / RolesDisplay) so they never leak into the
    // persisted file.
    public class CatalogService
    {
        // Shared serializer options. camelCase matches the existing
        // wwwroot/Data/templates.json + the legacy React client + the
        // StudioController API endpoints. JsonIgnoreCondition.WhenWritingNull
        // keeps optional fields (DocxUrl, UploadedAt, UpdatedAt) out of
        // the JSON when they're unset (e.g. a brand-new blank template).
        public static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        };

        private static readonly string DataFolder =
            Combine(Directory.GetCurrentDirectory(), "wwwroot", "Data");
        public static readonly string CatalogFile =
            Combine(DataFolder, "templates.json");
        public static readonly string CommonFieldsFile =
            Combine(DataFolder, "common-merge-fields.json");
        public static readonly string TemplatesFolder =
            Combine(Directory.GetCurrentDirectory(), "wwwroot", "Templates");

        // Load templates.json into a list of strongly-typed entries.
        // Missing/corrupt file returns an empty list so the dashboard
        // just renders an empty state.
        public List<TemplateEntry> LoadTemplates()
        {
            try
            {
                if (!File.Exists(CatalogFile)) return new();
                var json = File.ReadAllText(CatalogFile, Encoding.UTF8);
                if (string.IsNullOrWhiteSpace(json)) return new();
                var arr = JsonSerializer.Deserialize<JsonElement>(json);
                if (arr.ValueKind != JsonValueKind.Array) return new();
                var list = new List<TemplateEntry>();
                foreach (var e in arr.EnumerateArray())
                {
                    list.Add(TemplateEntry.FromJson(e));
                }
                return list;
            }
            catch
            {
                return new();
            }
        }

        public TemplateEntry? FindTemplate(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            return LoadTemplates().FirstOrDefault(t => t.Id == id);
        }

        // Persist the catalog back to disk. Uses the shared JsonOptions
        // (camelCase + WhenWritingNull) so the file matches the shape
        // the legacy React client and StudioController API both expect.
        // Display-only helpers (CreatedDisplay / RolesDisplay) are
        // decorated with [JsonIgnore] on TemplateEntry so they never
        // leak into the persisted file.
        public void SaveTemplates(List<TemplateEntry> templates)
        {
            Directory.CreateDirectory(DataFolder);
            var json = JsonSerializer.Serialize(templates, JsonOptions);
            File.WriteAllText(CatalogFile, json, Encoding.UTF8);
        }

        public void UpsertTemplate(TemplateEntry entry)
        {
            var list = LoadTemplates();
            var i = list.FindIndex(t => t.Id == entry.Id);
            if (i >= 0) list[i] = entry;
            else list.Add(entry);
            SaveTemplates(list);
        }

        // "Save and Publish" bookkeeping — invoked by
        // DocumentEditorController.Save when a TemplateId accompanies the
        // SFDT payload. Links (or re-links) the entry's DocxUrl to the
        // saved file and stamps UpdatedAt so the dashboard shows the
        // latest publish date. Returns the refreshed entry, or null when
        // the id doesn't match a catalog entry.
        public TemplateEntry? MarkTemplatePublished(string id, string docxUrl)
        {
            if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(docxUrl)) return null;
            var entry = FindTemplate(id);
            if (entry == null) return null;
            entry.DocxUrl = docxUrl;
            entry.UpdatedAt = DateTime.UtcNow.ToString("o");
            if (string.IsNullOrEmpty(entry.UploadedAt)) entry.UploadedAt = entry.UpdatedAt;
            UpsertTemplate(entry);
            return entry;
        }

        public bool DeleteTemplate(string id)
        {
            var list = LoadTemplates();
            var i = list.FindIndex(t => t.Id == id);
            if (i < 0) return false;
            var entry = list[i];
            list.RemoveAt(i);
            SaveTemplates(list);
            // Best-effort: unlink the .docx from disk if it lives in
            // wwwroot/Templates/.
            if (!string.IsNullOrEmpty(entry.DocxUrl))
            {
                try
                {
                    var slug = entry.DocxUrl.Split('/').LastOrDefault() ?? "";
                    if (!string.IsNullOrEmpty(slug))
                    {
                        var path = Combine(TemplatesFolder, slug);
                        if (File.Exists(path)) File.Delete(path);
                    }
                }
                catch { /* ignore — catalog is already updated */ }
            }
            return true;
        }

        // --- common merge fields ---
        public List<string> LoadCommonFieldNames()
        {
            try
            {
                if (!File.Exists(CommonFieldsFile)) return new();
                var json = File.ReadAllText(CommonFieldsFile, Encoding.UTF8);
                if (string.IsNullOrWhiteSpace(json)) return new();
                var doc = JsonDocument.Parse(json);
                if (!doc.RootElement.TryGetProperty("fields", out var fields)) return new();
                if (fields.ValueKind != JsonValueKind.Object) return new();
                return fields.EnumerateObject()
                    .Select(p => p.Name)
                    .Where(n => !string.IsNullOrWhiteSpace(n))
                    .ToList();
            }
            catch
            {
                return new();
            }
        }

        public void AddCommonField(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return;
            var fields = LoadCommonFieldNames();
            if (fields.Contains(key)) return;
            fields.Add(key);
            var payload = new
            {
                fields = fields.ToDictionary(f => f, _ => (object)true),
                updatedAt = DateTime.UtcNow.ToString("o"),
            };
            var json = JsonSerializer.Serialize(payload,
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(CommonFieldsFile, json, Encoding.UTF8);
        }
    }

    // Strongly-typed shape of a templates.json entry.
    public class TemplateEntry
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public string Type { get; set; } = "General";
        public string Description { get; set; } = "";
        public List<string> FieldKeys { get; set; } = new();
        public string? DocxUrl { get; set; }
        public string? UploadedAt { get; set; }
        public string? UpdatedAt { get; set; }
        public List<string> Roles { get; set; } = new();

        public static TemplateEntry FromJson(JsonElement e)
        {
            var t = new TemplateEntry();
            // Read by EITHER camelCase (the canonical shape written by
            // the React client + StudioController + this service) OR
            // PascalCase (the corrupt shape an earlier build of
            // CatalogService wrote before JsonNamingPolicy.CamelCase
            // was applied). Without the PascalCase fallback we'd be
            // unable to load the existing templates.json that was
            // written during the broken interim.
            string Str(string camel, string pascal)
            {
                if (e.TryGetProperty(camel, out var c) && c.ValueKind == JsonValueKind.String) return c.GetString() ?? "";
                if (e.TryGetProperty(pascal, out var p) && p.ValueKind == JsonValueKind.String) return p.GetString() ?? "";
                return "";
            }
            t.Id = Str("id", "Id");
            t.Name = Str("name", "Name");
            t.Type = Str("type", "Type");
            if (string.IsNullOrEmpty(t.Type)) t.Type = "General";
            t.Description = Str("description", "Description");
            // fieldKeys
            if (e.TryGetProperty("fieldKeys", out var keys) && keys.ValueKind == JsonValueKind.Array)
                t.FieldKeys = keys.EnumerateArray().Select(k => k.GetString() ?? "").Where(s => !string.IsNullOrEmpty(s)).ToList();
            else if (e.TryGetProperty("FieldKeys", out var keysP) && keysP.ValueKind == JsonValueKind.Array)
                t.FieldKeys = keysP.EnumerateArray().Select(k => k.GetString() ?? "").Where(s => !string.IsNullOrEmpty(s)).ToList();
            t.DocxUrl = Str("docxUrl", "DocxUrl");
            if (string.IsNullOrEmpty(t.DocxUrl)) t.DocxUrl = null;
            t.UploadedAt = Str("uploadedAt", "UploadedAt");
            if (string.IsNullOrEmpty(t.UploadedAt)) t.UploadedAt = null;
            t.UpdatedAt = Str("updatedAt", "UpdatedAt");
            if (string.IsNullOrEmpty(t.UpdatedAt)) t.UpdatedAt = null;
            // roles
            if (e.TryGetProperty("roles", out var roles) && roles.ValueKind == JsonValueKind.Array)
                t.Roles = roles.EnumerateArray().Select(r => r.GetString() ?? "").Where(s => !string.IsNullOrEmpty(s)).ToList();
            else if (e.TryGetProperty("Roles", out var rolesP) && rolesP.ValueKind == JsonValueKind.Array)
                t.Roles = rolesP.EnumerateArray().Select(r => r.GetString() ?? "").Where(s => !string.IsNullOrEmpty(s)).ToList();
            // Strip accidentally-persisted display helpers if present.
            return t;
        }

        // Render a friendly Created date string for the dashboard.
        // JsonIgnore keeps this display-only helper OUT of the
        // persisted templates.json — only the raw fields are stored.
        [JsonIgnore]
        public string CreatedDisplay
        {
            get
            {
                var s = UploadedAt ?? UpdatedAt;
                if (string.IsNullOrEmpty(s)) return "—";
                if (DateTime.TryParse(s, out var d))
                    return d.ToString("MMM d, yyyy");
                return "—";
            }
        }

        [JsonIgnore]
        public string RolesDisplay =>
            Roles.Count > 0 ? string.Join(", ", Roles) : "—";
    }
}
