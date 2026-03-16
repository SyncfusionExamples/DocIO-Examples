var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();

var app = builder.Build();

// Set FontManager delay at startup (before any conversions happen)
// Default is 30000ms. Adjust based on your conversion workload.
Syncfusion.Drawing.Fonts.FontManager.Delay = 50000;

// Access the application lifetime service
var lifetime = app.Services.GetRequiredService<IHostApplicationLifetime>();

// Register a callback to run when the app is shutting down
lifetime.ApplicationStopping.Register(() =>
{
    Syncfusion.Drawing.Fonts.FontManager.ClearCache();
});

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
