using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using BlazorOffice;
using KristofferStrube.Blazor.FileSystemAccess;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });
builder.Services.AddFileSystemAccessService();
builder.Services.AddScoped<BlazorOffice.Services.WordService>();
builder.Services.AddScoped<BlazorOffice.Services.ExcelService>();

await builder.Build().RunAsync();
