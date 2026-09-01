using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace DocumentTemplateStudioService.Pages
{
    // Create Template page — mirrors the React CreateTemplateDialog:
    // collects Name, Description, Signer Roles (chips + suggested list),
    // and an optional .docx "Editable Body" file. On submit:
    //   - WITH a .docx → calls the existing /api/studio/upload endpoint
    //     (server writes the file + a fresh catalog entry), then
    //     redirects to /Editor?id=<newId>.
    //   - WITHOUT a .docx → creates a blank template entry directly in
    //     templates.json via CatalogService, then redirects to /Editor
    //     for the user to author from scratch.
    //
    // Also handles the "Edit metadata" mode (when an id is supplied on
    // GET) for editing the name/description/roles of an existing
    // template without re-uploading a .docx.
    //
    // [IgnoreAntiforgeryToken] must sit at the MODEL level — Razor
    // Pages silently ignores it on handler methods (compiler warning
    // MVC1001), which made the antiforgery validation return an empty
    // 400 before OnPost ever ran (surfaced in the modal as "Unexpected
    // end of JSON input"). This handler is called via fetch() from the
    // dashboard modal, which doesn't carry the antiforgery token.
    [IgnoreAntiforgeryToken]
    public class CreateModel : PageModel
    {
        private readonly CatalogService _catalog;
        private readonly IWebHostEnvironment _env;

        public CreateModel(CatalogService catalog, IWebHostEnvironment env)
        {
            _catalog = catalog;
            _env = env;
        }

        [BindProperty] public string? Id { get; set; }
        [BindProperty] public string Name { get; set; } = "";
        [BindProperty] public string Description { get; set; } = "";
        [BindProperty] public string Type { get; set; } = "General";
        [BindProperty] public string RolesCsv { get; set; } = "";
        [BindProperty] public IFormFile? BodyFile { get; set; }

        public TemplateEntry? Existing { get; set; }
        public bool IsEdit => !string.IsNullOrEmpty(Id);

        public void OnGet(string? id = null)
        {
            if (!string.IsNullOrEmpty(id))
            {
                Id = id;
                Existing = _catalog.FindTemplate(id);
                if (Existing != null)
                {
                    Name = Existing.Name;
                    Description = Existing.Description;
                    Type = Existing.Type;
                    RolesCsv = string.Join(", ", Existing.Roles);
                }
            }
        }

        // Handles the form submit via AJAX from the dashboard's Create
        // Template modal. Returns a JSON envelope so the modal can
        // close + redirect without a full page reload:
        //   { ok: true, id: "<id>", redirect: "/Editor?id=<id>" }
        // On validation failure returns:
        //   { ok: false, error: "<message>" }
        //
        // Two flows:
        //   - WITH a .docx body file → write the .docx to
        //     wwwroot/Templates/ + create the catalog entry, then
        //     redirect to /Editor?id=<newId>.
        //   - WITHOUT a .docx → create a blank template entry directly
        //     in templates.json via CatalogService, then redirect to
        //     /Editor for the user to author from scratch.
        //   - IsEdit mode → update the existing entry's metadata in
        //     place (no .docx re-upload), then redirect to /Editor.
        //
        // POST is called via fetch() from the dashboard modal (no
        // antiforgery token) — see the model-level attribute above.
        public IActionResult OnPost()
        {
            if (string.IsNullOrWhiteSpace(Name))
            {
                return BadRequest(new { ok = false, error = "Template Name is required." });
            }

            // RolesCsv binds NULL when the field is submitted empty —
            // splitting it directly threw NullReferenceException (the old
            // HTTP 500). Treat null as empty, then REQUIRE at least one
            // signer role: the user must provide Signer Roles to create
            // the template.
            var rolesCsv = RolesCsv ?? string.Empty;
            var roles = rolesCsv
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .ToList();
            if (roles.Count == 0)
            {
                return BadRequest(new
                {
                    ok = false,
                    error = "Signer Roles is required. Please provide at least one signer role (e.g. Approver) and try again.",
                });
            }

            if (IsEdit && Existing != null)
            {
                Existing.Name = Name.Trim();
                Existing.Description = Description;
                Existing.Type = Type;
                Existing.Roles = roles;
                Existing.UpdatedAt = DateTime.UtcNow.ToString("o");
                _catalog.UpsertTemplate(Existing);
                return new JsonResult(new
                {
                    ok = true,
                    id = Existing.Id,
                    redirect = $"/Editor?id={Existing.Id}",
                });
            }

            // New template WITH a .docx body — write the file + create
            // the catalog entry in one transaction (mirrors the old
            // /api/studio/upload endpoint inline).
            if (BodyFile != null && BodyFile.Length > 0)
            {
                var safe = Name.Trim().Replace(' ', '_');
                var slug = $"{safe}-{DateTime.UtcNow.Ticks.ToString("x")}";
                var docxName = $"{slug}.docx";
                var docxPath = Path.Combine(CatalogService.TemplatesFolder, docxName);
                Directory.CreateDirectory(CatalogService.TemplatesFolder);
                using (var fs = new FileStream(docxPath, FileMode.Create, FileAccess.Write))
                {
                    BodyFile.CopyTo(fs);
                }
                var id = $"tpl-{slug}";
                var entry = new TemplateEntry
                {
                    Id = id,
                    Name = Name.Trim(),
                    Type = Type,
                    Description = Description,
                    Roles = roles,
                    FieldKeys = new(),
                    DocxUrl = $"/Templates/{docxName}",
                    UploadedAt = DateTime.UtcNow.ToString("o"),
                    UpdatedAt = DateTime.UtcNow.ToString("o"),
                };
                _catalog.UpsertTemplate(entry);
                return new JsonResult(new
                {
                    ok = true,
                    id,
                    redirect = $"/Editor?id={id}",
                });
            }

            // Blank template — no .docx yet. The editor will open in
            // blank mode; Save will publish the first .docx.
            var blankId = $"tpl-{DateTime.UtcNow.Ticks.ToString("x")}-{Guid.NewGuid().ToString("N").Substring(0, 5)}";
            var blank = new TemplateEntry
            {
                Id = blankId,
                Name = Name.Trim(),
                Type = Type,
                Description = string.IsNullOrEmpty(Description)
                    ? "Blank letter template — edit to customize."
                    : Description,
                Roles = roles,
                FieldKeys = new(),
                UpdatedAt = DateTime.UtcNow.ToString("o"),
            };
            _catalog.UpsertTemplate(blank);
            return new JsonResult(new
            {
                ok = true,
                id = blankId,
                redirect = $"/Editor?id={blankId}",
            });
        }
    }
}
