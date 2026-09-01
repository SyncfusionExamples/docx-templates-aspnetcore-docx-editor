using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace DocumentTemplateStudioService.Pages
{
    // Dashboard / Index page — the template list. Mirrors the React
    // TemplateList component: a table with NAME / ROLES / CREATED /
    // ACTIONS columns, plus a "Create Template" button. Clicking a
    // row or the edit icon navigates to /Editor?id=<templateId> which
    // renders the EJ2 DocumentEditor + Merge Fields panel.
    public class IndexModel : PageModel
    {
        private readonly CatalogService _catalog;

        public IndexModel(CatalogService catalog)
        {
            _catalog = catalog;
        }

        public List<TemplateEntry> Templates { get; set; } = new();
        public string? SearchQuery { get; set; }
        public string? Toast { get; set; }

        public void OnGet(string? q = null, string? toast = null)
        {
            var all = _catalog.LoadTemplates();
            SearchQuery = (q ?? "").Trim();
            Templates = string.IsNullOrEmpty(SearchQuery)
                ? all
                : all.Where(t =>
                    t.Name.Contains(SearchQuery, StringComparison.OrdinalIgnoreCase) ||
                    (t.Type?.Contains(SearchQuery, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (t.Description?.Contains(SearchQuery, StringComparison.OrdinalIgnoreCase) ?? false))
                    .ToList();
            Toast = toast;
        }

        // Navigation helper — used by the cshtml to build the edit link.
        public string EditUrl(string id) => Url.Page("/Editor", new { id }) ?? "#";
        public string DeleteUrl(string id) => Url.Page("/Index", new { handler = "Delete", id }) ?? "#";

        // POST: /Index?handler=Delete&id=<id> — removes the template
        // (and its .docx file) from the server-side catalog. Returns
        // to the dashboard with a confirmation toast message.
        public IActionResult OnPostDelete(string id)
        {
            if (string.IsNullOrEmpty(id))
            {
                return RedirectToPage("/Index", new { toast = "Missing template id." });
            }
            var entry = _catalog.FindTemplate(id);
            var ok = _catalog.DeleteTemplate(id);
            return RedirectToPage("/Index",
                new { toast = ok && entry != null
                    ? $"Deleted \"{entry!.Name}\"."
                    : "Template not found." });
        }
    }
}
