using Azure.Core;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Azure.KeyVault;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;
using PnP.Framework;
using ILogger = Microsoft.Extensions.Logging.ILogger;

namespace appsvc_fnc_dev_scw_sensitivity_dotnet001
{
    internal class Auth
    {

        internal static X509Certificate2 GetKeyVaultCertificateAsync(string keyVaultUrl, string name, ILogger log)
        {
            log.LogInformation("GetKeyVaultCertificateAsync received a request.");

            var client = new SecretClient(new Uri(keyVaultUrl), new DefaultAzureCredential());
            var secret = client.GetSecret(name).Value;

            X509Certificate2 certificate = new X509Certificate2(Convert.FromBase64String(secret.Value), string.Empty, X509KeyStorageFlags.MachineKeySet);

            log.LogInformation("GetKeyVaultCertificateAsync processed a request.");

            return certificate;
        }

        internal static ClientContext GetContextByCertificate(string siteUrl, string keyVaultUrl,string certificateName, string clientId, string tenantId, ILogger log)
        {
            X509Certificate2 mycert = Auth.GetKeyVaultCertificateAsync(keyVaultUrl, certificateName, log);
            var ctx = new AuthenticationManager(clientId, mycert, tenantId).GetContext(siteUrl);

            log.LogInformation($"Created client connection for {siteUrl}");

            return ctx;
        }
    }

    public class ROPCConfidentialTokenCredential : Azure.Core.TokenCredential
    {
        // Implementation of the Azure.Core.TokenCredential class
        string _clientId;
        string _clientSecret;
        string _password;
        string _tenantId;
        string _tokenEndpoint;
        string _username;
        ILogger _log;

        public ROPCConfidentialTokenCredential(ILogger log)
        {
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string keyVaultUrl = config["keyVaultUrl"];
            string secretName = config["secretName"];
            string secretNamePassword = config["secretNamePassword"];

            _clientId = config["clientId"];
            _tenantId = config["tenantId"];
            _username = config["user_name"];
            _log = log;
            _tokenEndpoint = "https://login.microsoftonline.com/" + _tenantId + "/oauth2/v2.0/token";

            SecretClientOptions options = new SecretClientOptions()
            {
                Retry =
                {
                    Delay= TimeSpan.FromSeconds(2),
                    MaxDelay = TimeSpan.FromSeconds(16),
                    MaxRetries = 5,
                    Mode = RetryMode.Exponential
                 }
            };
            var client = new SecretClient(new Uri(keyVaultUrl), new DefaultAzureCredential(), options);

            KeyVaultSecret secret = client.GetSecret(secretName);
            _clientSecret = secret.Value;

            KeyVaultSecret password = client.GetSecret(secretNamePassword);
            _password = password.Value;
        }

        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            HttpClient httpClient = new HttpClient();

            // Create the request body
            var Parameters = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("client_id", _clientId),
                new KeyValuePair<string, string>("client_secret", _clientSecret),
                new KeyValuePair<string, string>("scope", string.Join(" ", requestContext.Scopes)),
                new KeyValuePair<string, string>("username", _username),
                new KeyValuePair<string, string>("password", _password),
                new KeyValuePair<string, string>("grant_type", "password")
            };

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, _tokenEndpoint)
            {
                Content = new FormUrlEncodedContent(Parameters)
            };
            var response = httpClient.SendAsync(request).Result.Content.ReadAsStringAsync().Result;
            dynamic responseJson = JsonConvert.DeserializeObject(response);
            var expirationDate = DateTimeOffset.UtcNow.AddMinutes(60.0);
            return new AccessToken(responseJson.access_token.ToString(), expirationDate);
        }

        public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            HttpClient httpClient = new HttpClient();

            // Create the request body
            var Parameters = new List<KeyValuePair<string, string>>
            {
                new KeyValuePair<string, string>("client_id", _clientId),
                new KeyValuePair<string, string>("client_secret", _clientSecret),
                new KeyValuePair<string, string>("scope", string.Join(" ", requestContext.Scopes)),
                new KeyValuePair<string, string>("username", _username),
                new KeyValuePair<string, string>("password", _password),
                new KeyValuePair<string, string>("grant_type", "password")
            };

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, _tokenEndpoint)
            {
                Content = new FormUrlEncodedContent(Parameters)
            };
            var response = httpClient.SendAsync(request).Result.Content.ReadAsStringAsync().Result;
            dynamic responseJson = JsonConvert.DeserializeObject(response);
            var expirationDate = DateTimeOffset.UtcNow.AddMinutes(60.0);
            return new ValueTask<AccessToken>(new AccessToken(responseJson.access_token.ToString(), expirationDate));
        }
    }
}
