namespace PowerBI.Services
{
    using Microsoft.Identity.Client;
    using Microsoft.Extensions.Configuration;

    public class PowerBIAuthService
    {
        private readonly IConfiguration _config;

        public PowerBIAuthService(IConfiguration config)
        {
            _config = config;
        }

        public string GetAdminEmail()
        {
            return _config["PowerBI:AdminEmail"] ?? string.Empty;
        }

        public string GetCapacityId()
        {
            return _config["PowerBI:CapacityId"] ?? string.Empty;
        }

        public async Task<string> GetAccessToken()
        {
            var tenantId = _config["PowerBI:TenantId"];
            var clientId = _config["PowerBI:ClientId"];
            var clientSecret = _config["PowerBI:ClientSecret"];

            // For Service Principal authentication, use ConfidentialClientApplication
            var app = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
                .Build();

            var result = await app
                .AcquireTokenForClient(new[] { "https://analysis.windows.net/powerbi/api/.default" })
                .ExecuteAsync();

            return result.AccessToken;
        }

        public async Task<string> GetFabricToken()
        {
            var tenantId = _config["PowerBI:TenantId"];
            var clientId = _config["PowerBI:ClientId"];
            var clientSecret = _config["PowerBI:ClientSecret"];

            var app = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
                .Build();

            // Fabric API scope
            var result = await app
                .AcquireTokenForClient(new[] { "https://api.fabric.microsoft.com/.default" })
                .ExecuteAsync();

            return result.AccessToken;
        }
    }
}