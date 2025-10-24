using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using Microsoft.SharePoint.Client;
using Azure.Core;
using Microsoft.Azure.Functions.Worker;

namespace appsvc_fnc_dev_scw_sensitivity_dotnet001
{
    public class ApplyUnclassifiedSettings
    {
        private readonly ILogger<ApplyUnclassifiedSettings> _logger;
        public ApplyUnclassifiedSettings(ILogger<ApplyUnclassifiedSettings> logger)
        {
            _logger = logger;
        }

        [Function("ApplyUnclassifiedSettings")]
        public async Task RunAsync([QueueTrigger("unclassified", Connection = "AzureWebJobsStorage")] string myQueueItem, ExecutionContext functionContext)
        {
            _logger.LogInformation($"ApplyUnclassifiedSettings received a request: {myQueueItem}");

            dynamic data = JsonConvert.DeserializeObject(myQueueItem);

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string groupId = data?.groupId;
            string itemId = data?.itemId;
            string labelId = config["unclassifiedLabelId"];
            string ownerId = config["ownerId"]; // sv-caupdate@devgcx.ca
            string readOnlyGroup = config["readOnlyGroup"]; // dgcx_allusers, dgcx_assigned
            string requestId = data?.Id;
            string SCAGroupName = config["sca_login_name"]; // dgcx_sca
            string sharePointUrl = config["sharePointUrl"] + requestId;
            string spaceNameEn = data?.SpaceName;
            string spaceNameFr = data?.SpaceNameFR;
            string supportGroupName = config["support_group_login_name"];   // dgcx_support
            string tenantName = config["tenantName"];

            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(_logger);
            var graphClient = new GraphServiceClient(auth);

            var result = Common.ApplyLabel(graphClient, labelId, groupId, itemId, requestId, spaceNameEn, spaceNameFr, _logger);

            if (result.Result == true)
            {
                // do not call method to set Visibility = Public
                //await SetUnclassified(graphClient, groupId, log);

                var scopes = new string[] { $"https://{tenantName}.sharepoint.com/.default" };
                var authManager = new PnP.Framework.AuthenticationManager();
                var accessToken = await auth.GetTokenAsync(new TokenRequestContext(scopes), new System.Threading.CancellationToken());
                var ctx = authManager.GetAccessTokenContext(sharePointUrl, accessToken.Token);

                bool result1 = await Common.UpdateSiteCollectionAdministrator(ctx, SCAGroupName, groupId, _logger);
                bool result2 = await AddGroupToFullControl(ctx, supportGroupName, _logger);
                bool result3 = await AddGroupToReadOnly(ctx, readOnlyGroup, _logger);
                bool result4 = await Common.RemoveOwner(graphClient, groupId, ownerId, _logger);

                bool success = result1 && result2 && result3 && result4;

                if (success) {
                    await Common.AddToStatusQueue(itemId, _logger);
                    await Common.AddToEmailQueue(requestId, "unclassified", groupId, spaceNameEn, spaceNameFr, (string)data?.RequesterName, (string)data?.RequesterEmail, _logger);
                }
            }

            _logger.LogInformation($"ApplyUnclassifiedSettings processed a request.");
        }

        public static Task<bool> AddGroupToFullControl(ClientContext ctx, string GroupLoginName, ILogger log)
        {
            var result = true;

            try
            {
                string permissionLevel = "Full Control";

                var adGroup = ctx.Web.EnsureUser(GroupLoginName);
                ctx.Load(adGroup);

                var spGroup = ctx.Web.AssociatedMemberGroup;

                var writeDefinition = ctx.Web.RoleDefinitions.GetByName(permissionLevel);
                var roleDefCollection = new RoleDefinitionBindingCollection(ctx) { writeDefinition};
                var newRoleAssignment = ctx.Web.RoleAssignments.Add(adGroup, roleDefCollection);
                ctx.Load(spGroup, x => x.Users);
                ctx.ExecuteQuery();
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
                result = false;
            }

            return Task.FromResult(result);
        }

        public static Task<bool> AddGroupToReadOnly(ClientContext ctx, string groups, ILogger log)
        {
            string permissionLevel = "Read";
            var result = true;

            try
            {
                // this prevents the Hub Visitor group from being added to site permissions
                ctx.Load(ctx.Site);
                ctx.Site.CanSyncHubSitePermissions = false;

                // break inheritance on the default document library to prevent access to read-only
                ctx.Web.DefaultDocumentLibrary().BreakRoleInheritance(true, true);

                foreach (string group in groups.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var adGroup = ctx.Web.EnsureUser(group);
                    ctx.Load(adGroup);

                    var spGroup = ctx.Web.AssociatedMemberGroup;

                    var writeDefinition = ctx.Web.RoleDefinitions.GetByName(permissionLevel);
                    var roleDefCollection = new RoleDefinitionBindingCollection(ctx) { writeDefinition };
                    var newRoleAssignment = ctx.Web.RoleAssignments.Add(adGroup, roleDefCollection);

                    ctx.Load(spGroup, x => x.Users);
                    ctx.ExecuteQuery();
                }
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
                result = false;
            }

            return Task.FromResult(result);
        }
    }
}