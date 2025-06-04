using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using SharepointSiteConfiguration.Models;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;


namespace SharepointSiteConfiguration.Auth
{
    internal class AuthProvider
    {

        public class TokenCredentialAuthProvider : IAuthenticationProvider
        {
            private readonly TokenCredential _tokenCredential;
            private readonly string[] _scopes;

            public TokenCredentialAuthProvider(ClientCertificateCredential tokenCredential, string[] scopes)
            {
                _tokenCredential = tokenCredential;
                _scopes = scopes;
            }
            public async Task AuthenticateRequestAsync(HttpRequestMessage request)
            {
                var token = await _tokenCredential.GetTokenAsync(new Azure.Core.TokenRequestContext(_scopes), default);
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            }

            public async Task AuthenticateRequestAsync(HttpRequestMessage requestMessage, System.Threading.CancellationToken cancellationToken)
            {
                await AuthenticateRequestAsync(requestMessage);
            }

        }

        public static async Task<string> GetAccessTokenAsync(AppSettings settings)
        {
            var cert = new X509Certificate2(settings.CertificatePath, settings.CertificatePassword, X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet);

            var app = ConfidentialClientApplicationBuilder.Create(settings.ClientId).WithAuthority($"https://login.microsoftonline.com/{settings.TenantId}").WithCertificate(cert).Build();
            var result = await app.AcquireTokenForClient(new[] { "https://dgneaseteq.sharepoint.com/.default" }).ExecuteAsync();

            return result.AccessToken;
        }

        public static GraphServiceClient GetGraphClient(AppSettings settings)
        {
            // Load certificate
            var cert = new X509Certificate2(settings.CertificatePath, settings.CertificatePassword,
                X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable | X509KeyStorageFlags.PersistKeySet);

            // Create ClientCertificateCredential
            var clientCertCredential = new ClientCertificateCredential(
                settings.TenantId,
                settings.ClientId,
                cert
            );

            var authProvider = new TokenCredentialAuthProvider(clientCertCredential, new[] { "https://graph.microsoft.com/.default" });

            // Create GraphServiceClient with this credential and scopes
            var graphClient = new GraphServiceClient(authProvider);

            return graphClient;
        }

    }
    
}
