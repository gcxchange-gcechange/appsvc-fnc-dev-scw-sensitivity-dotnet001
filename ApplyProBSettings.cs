using Azure.Core;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Microsoft.Azure.Functions.Worker;

namespace appsvc_fnc_dev_scw_sensitivity_dotnet001
{
    public class ApplyProBSettings
    {
        private readonly ILogger<ApplyProBSettings> _logger;
        public ApplyProBSettings(ILogger<ApplyProBSettings> logger)
        {
            _logger = logger;
        }

        [Function("ApplyProBSettings")]
        public async Task RunAsync([QueueTrigger("prob", Connection = "AzureWebJobsStorage")] string myQueueItem)
        {
            _logger.LogInformation($"ApplyProBSettings received a request: {myQueueItem}");

            dynamic data = JsonConvert.DeserializeObject(myQueueItem);

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string groupId = data?.groupId;
            string itemId = data?.itemId;
            string labelId = config["proBLabelId"];
            string ownerId = config["ownerId"]; // sv-caupdate@devgcx.ca
            string requestId = data?.Id;
            string SCAGroupName = config["sca_prob_login_name"]; // dgcx-sca-prob
            string sharePointUrl = config["sharePointUrl"] + requestId;
            string spaceNameEn = data?.SpaceName;
            string spaceNameFr = data?.SpaceNameFR;
            string tenantName = config["tenantName"];

            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(_logger);
            var graphClient = new GraphServiceClient(auth);

            var result = Common.ApplyLabel(graphClient, labelId, groupId, itemId, requestId, spaceNameEn, spaceNameFr, _logger);
            
            if (result.Result == true)
            {
                var scopes = new string[] { $"https://{tenantName}.sharepoint.com/.default" };
                var authManager = new PnP.Framework.AuthenticationManager();
                var accessToken = await auth.GetTokenAsync(new TokenRequestContext(scopes), new System.Threading.CancellationToken());
                var ctx = authManager.GetAccessTokenContext(sharePointUrl, accessToken.Token);

                bool result1 = await SetProB(graphClient, groupId, ctx, _logger);
                bool result2 = await Common.UpdateSiteCollectionAdministrator(ctx, SCAGroupName, groupId, _logger);
                bool result3 = await Common.RemoveOwner(graphClient, groupId, ownerId, _logger);

                bool success = result1 && result2 && result3;

                if (success)
                {
                    await Common.AddToStatusQueue(itemId, _logger);
                    await Common.AddToEmailQueue(requestId, "prob", groupId, spaceNameEn, spaceNameFr, (string)data?.RequesterName, (string)data?.RequesterEmail, _logger);
                }
            }

            _logger.LogInformation($"ApplyProBSettings processed a request.");
        }

        private static Task<bool> SetProB(GraphServiceClient graphClient, string groupId, ClientContext ctx, ILogger log)
        {
            log.LogInformation("SetProB received a request.");

            bool result = true;

            try
            {
                // remove the visitor's group
                var avg = ctx.Web.AssociatedVisitorGroup;
                ctx.Load(avg, w => w.Title);
                ctx.ExecuteQuery();

                if (avg != null) {
                    log.LogInformation($"Removing group: {avg.Title}");
                    ctx.Web.RemoveGroup(avg);
                }

                // this prevents the Hub Visitor group from being added to site permissions
                ctx.Load(ctx.Site);
                ctx.Site.CanSyncHubSitePermissions = false;

                // set group visibility to private
                var group = new Microsoft.Graph.Group { Visibility = "Private" };
                graphClient.Groups[groupId].Request().UpdateAsync(group);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
                result = false;
            }

            log.LogInformation("SetProB processed a request.");

            return Task.FromResult(result);
        }
    }
}