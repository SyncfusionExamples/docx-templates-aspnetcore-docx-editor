
namespace DocumentTemplateStudioService
{
    public class Program
    {
        internal static string webRootPath;

        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            // Razor Pages + EJ2 ASP.NET Core Tag Helpers. The previously
            // React-only client is now a pure server-rendered ASP.NET
            // Core app — the dashboard, create dialog, sidebar and merge
            // fields panel are all Razor Pages / partials, and the
            // document canvas uses Syncfusion's EJ2 DocumentEditor
            // ASP.NET Core control (<ejs-documenteditorcontainer>) via
            // the WordEditor.AspNet.Core NuGet package.
            builder.Services.AddRazorPages();
            builder.Services.AddControllers();

            // CatalogService is a scoped service that owns reads/writes
            // of wwwroot/Data/templates.json + common-merge-fields.json.
            // Both files live in wwwroot/Data/ and are shared with the
            // existing StudioController API endpoints (still mounted so
            // the editor's AJAX save/import paths keep working).
            builder.Services.AddScoped<CatalogService>();

            builder.Services.AddCors(options =>
            {
                options.AddPolicy("AllowAllOrigins", builder =>
                {
                    builder.AllowAnyOrigin()
                    .AllowAnyMethod()
                    .AllowAnyHeader();
                });
            });

            var app = builder.Build();

            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("");


            webRootPath = builder.Environment.WebRootPath;

            app.UseStaticFiles();

            app.UseRouting();

            app.UseCors("AllowAllOrigins");
            app.UseAuthorization();

            app.MapRazorPages();
            app.MapControllers();

            app.Run();
        }
    }
}
