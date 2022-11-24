using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using TeamsTimecardHelperClient;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });
builder.Services.AddMsalAuthentication(options => {
    builder.Configuration.Bind("AzureAd", options.ProviderOptions.Authentication);
    options.ProviderOptions.DefaultAccessTokenScopes.Add("User.Read.All");
    options.ProviderOptions.DefaultAccessTokenScopes.Add("Schedule.ReadWrite.All");
    });

builder.Services.AddGraphClient(
    "https://graph.microsoft.com/User.Read.All",
    "https://graph.microsoft.com/Schedule.ReadWrite.All"
    );

await builder.Build().RunAsync();
