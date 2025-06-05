using Microsoft.Extensions.Configuration;
using SharepointSiteConfiguration.Auth;
using SharepointSiteConfiguration.Models;
using SharepointSiteConfiguration.Services;

class Program {
    static async Task Main(string[] args) {
        var config = new ConfigurationBuilder()
        .AddJsonFile("C:\\Projects\\PnP Powershell\\Project\\CSOM\\SharepointSiteConfiguration\\SharepointSiteConfiguration\\appSettings.Development.json")
        .Build();
        var settings = config.Get<AppSettings>();
        var service = new SharepointServices();
        var accessToken = await AuthProvider.GetAccessTokenAsync(settings);
        var graphClient = AuthProvider.GetGraphClient(settings);

        await service.CreateGroupSite(settings, accessToken, graphClient);
    }
}