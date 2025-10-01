using Azure.Core;
using Microsoft.Extensions.Configuration;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using PnP.Framework.ALM;
using PnP.Framework.Diagnostics;
using PnP.Framework.Enums;
using SharepointSiteConfiguration.Auth;
using SharepointSiteConfiguration.Models;
using SharepointSiteConfiguration.Services;

class Program {
    static async Task Main(string[] args) {
        var config = new ConfigurationBuilder()
        .AddJsonFile("C:\\Projects\\PnP Powershell\\Project\\CSOM\\SharepointSiteConfiguration\\SharepointSiteConfiguration\\appSettings.Development.json")
        .Build();
        var settings = config.Get<AppSettings>();
        var accessToken = await AuthProvider.GetAccessTokenAsync(settings);
        var graphClient = AuthProvider.GetGraphClient(settings);
        var service = new SharepointServices();

        //await service.CreateGroupSite(settings, accessToken, graphClient);
        //service.UploadPackageFileAsync(settings, accessToken) 
    }
}