using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace DocumentTemplateStudioService.Pages
{
    // Editor page — opens a template in the EJ2 DocumentEditor
    // (<ejs-documenteditorcontainer>) with the server-side SFDT import
    // pipeline. The right-rail Merge Fields panel is rendered as a
    // partial (_MergeFields.cshtml) driven by the server-side union
    // of (a) the .docx's MERGEFIELDs, (b) template.fieldKeys, and
    // (c) the global common-merge-fields.json keys. The page's AJAX
    // save/import endpoints are still served by the existing API
    // controllers (/api/DocumentEditor/*, /api/studio/*).
    public class EditorModel : PageModel
    {
        private readonly CatalogService _catalog;

        public EditorModel(CatalogService catalog)
        {
            _catalog = catalog;
        }

        public TemplateEntry? Template { get; set; }
        public List<string> MergeFields { get; set; } = new();

        public IActionResult OnGet(string? id)
        {
            if (string.IsNullOrEmpty(id))
            {
                return RedirectToPage("/Index");
            }
            Template = _catalog.FindTemplate(id);
            if (Template == null)
            {
                return RedirectToPage("/Index", new { toast = "Template not found." });
            }
            // Seed MergeFields with the union of template.fieldKeys +
            // the global common-merge-fields.json keys so the panel has
            // something to render even before the editor's first
            // Import round-trip resolves the doc MERGEFIELDs. The
            // in-page AJAX call to /api/DocumentEditor/ImportFileURL
            // (fired by the editor's serviceUrl + a JS bootstrapping
            // script) will overwrite this with the full doc∪template∪
            // common union once the .docx body has been parsed.
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var k in Template.FieldKeys) if (seen.Add(k)) MergeFields.Add(k);
            foreach (var k in _catalog.LoadCommonFieldNames()) if (seen.Add(k)) MergeFields.Add(k);
            return Page();
        }

        // Hostname + port the editor's serviceUrl should resolve to.
        // Same value as the legacy DOCUMENT_EDITOR_BASE_URL — points at
        // this same ASP.NET Core service.
        public string ServiceUrl => $"{Request.Scheme}://{Request.Host}/api/DocumentEditor/";
    }
}
