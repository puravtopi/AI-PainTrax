


using Microsoft.AspNetCore.Diagnostics;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Logging;
using PainTrax.Web.Helper;
using PainTrax.Web.Services;
using PainTrax.Web.AzureServices;


var builder = WebApplication.CreateBuilder(args);



builder.Logging.ClearProviders();
builder.Logging.AddConsole();


builder.Services.AddSession(options =>
{
    options.IdleTimeout = TimeSpan.FromMinutes(30);
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true; // Ensures session is always stored
 
});

builder.Services.ConfigureApplicationCookie(options =>
{
    options.Cookie.HttpOnly = true;
    options.ExpireTimeSpan = TimeSpan.FromMinutes(30);
    options.LoginPath = "/Home/Login";
    options.AccessDeniedPath = "/Home/Login";
    options.SlidingExpiration = true;
});
builder.Services.AddScoped<UserService>();
builder.Services.AddScoped<SignatureService>();

builder.Services.AddTransient<IHttpContextAccessor, HttpContextAccessor>();
builder.Services.AddTransient<ILoggingService, LoggingService>();
builder.Services.AddTransient<IEmailService, EmailService>();
builder.Services.AddScoped<DataTransferService>();
builder.Services.AddHttpClient<SmsService>();

builder.Services.Configure<AzureVisionSettings>(
    builder.Configuration.GetSection("AzureVision"));
builder.Services.AddScoped<AzureAIServices>();


// Add services to the container.
builder.Services.AddControllersWithViews();



builder.Services.AddAutoMapper(AppDomain.CurrentDomain.GetAssemblies());
builder.Services.AddApplicationInsightsTelemetry();

builder.Logging.ClearProviders();
builder.Logging.AddConsole();   // Console
builder.Logging.AddDebug();     // VS Debug window


var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}
app.UseSession();
app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseStaticFiles(new StaticFileOptions()
{
    FileProvider = new
        PhysicalFileProvider(Path.Combine(Directory.GetCurrentDirectory(), @"Templete")),
    RequestPath = new PathString("/Templete")
});

app.UseRouting();
//app.UseExceptionHandler(errorApp =>
//{
//    errorApp.Run(async context =>
//    {
//        var exceptionHandlerPathFeature = context.Features.Get<IExceptionHandlerPathFeature>();
//        var exception = exceptionHandlerPathFeature?.Error;
//        await File.AppendAllTextAsync("errorlog.txt", exception?.ToString());
//        await context.Response.WriteAsync("Something went wrong.");
//    });
//});


app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Login}/{id?}");

app.Run();
