using System;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using PnP.Framework.Entities;
using Azure.Core;

namespace appsvc_fnc_dev_scw_sensitivity_dotnet001
{
    public class ApplyUnclassifiedSettings
    {
        [FunctionName("ApplyUnclassifiedSettings")]
        public async Task RunAsync([QueueTrigger("unclassified", Connection = "AzureWebJobsStorage")] string myQueueItem, ILogger log, ExecutionContext functionContext)
        {
            log.LogInformation($"ApplyUnclassifiedSettings received a request: {myQueueItem}");

            dynamic data = JsonConvert.DeserializeObject(myQueueItem);

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string displayName = data?.DisplayName;
            string groupId = data?.groupId;
            string itemId = data?.itemId;
            string labelId = config["unclassifiedLabelId"];
            string ownerId = config["ownerId"];
            string readOnlyGroup = config["readOnlyGroup"];
            string requestId = data?.Id;
            string SCAGroupName = config["sca_login_name"];
            string sharePointUrl = config["sharePointUrl"] + requestId;
            string supportGroupName = config["support_group_login_name"];
            string tenantName = config["tenantName"];

            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(log);
            var graphClient = new GraphServiceClient(auth);

            var result = Common.ApplyLabel(graphClient, labelId, groupId, itemId, requestId, displayName, log);

            if (result.Result == true)
            {
                // do not call method to set Visibility = Public
                //await SetUnclassified(graphClient, groupId, log);

                var scopes = new string[] { $"https://{tenantName}.sharepoint.com/.default" };
                var authManager = new PnP.Framework.AuthenticationManager();
                var accessToken = await auth.GetTokenAsync(new TokenRequestContext(scopes), new System.Threading.CancellationToken());
                var ctx = authManager.GetAccessTokenContext(sharePointUrl, accessToken.Token);

                await UpdateSiteCollectionAdministrator(ctx, SCAGroupName, groupId, log);   // dgcx_sca
                await AddGroupToFullControl(ctx, supportGroupName, log); // dgcx_support
                await AddGroupToReadOnly(ctx, readOnlyGroup, log); // dgcx_allusers, dgcx_assigned

                await Common.RemoveOwner(graphClient, groupId, ownerId, log); // sv-caupdate@devgcx.ca

                await Common.AddToEmailQueue(requestId, groupId, displayName, (string)data?.RequesterName, (string)data?.RequesterEmail, log);
            }

            log.LogInformation($"ApplyUnclassifiedSettings processed a request.");
        }

        //private static async Task<IActionResult> SetUnclassified(GraphServiceClient graphClient, string groupId, ILogger log)
        //{
        //    log.LogInformation("SetUnclassified received a request.");

        //    try
        //    {
        //        var group = new Microsoft.Graph.Group { Visibility = "Public" };
        //        await graphClient.Groups[groupId].Request().UpdateAsync(group);
        //    }
        //    catch (Exception e)
        //    {
        //        log.LogError($"Message: {e.Message}");
        //        if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
        //        log.LogError($"StackTrace: {e.StackTrace}");
        //    }

        //    log.LogInformation("SetUnclassified processed a request.");

        //    return new OkResult();
        //}

        public static Task<bool> UpdateSiteCollectionAdministrator(ClientContext ctx, string GroupLoginName, string groupId, ILogger log) // ClientContext ctx, 
        {
            log.LogInformation("UpdateSiteCollectionAdministrator received a request.");

            bool result = true;

            try
            {
                ctx.Load(ctx.Web);
                ctx.Load(ctx.Site);
                ctx.Load(ctx.Site.RootWeb);
                ctx.ExecuteQuery();

                // add dgcx_sca as Administrator
                List<UserEntity> admins = new List<UserEntity>();
                UserEntity adminUserEntity = new UserEntity();
                adminUserEntity.LoginName = GroupLoginName;
                admins.Add(adminUserEntity);
                ctx.Site.RootWeb.AddAdministrators(admins, true);

                // remove the owner group
                string loginName = $"c:0o.c|federateddirectoryclaimprovider|{groupId}_o";
                log.LogInformation($"Remove loginName = {loginName}");
                UserEntity ownerGroupEntity = new UserEntity();
                ownerGroupEntity.LoginName = loginName;
                ctx.Site.RootWeb.RemoveAdministrator(ownerGroupEntity);
                log.LogInformation($"Done!");
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
                result = false;
            }

            log.LogInformation("UpdateSiteCollectionAdministrator processed a request.");

            return Task.FromResult(result);
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
                spGroup.Users.AddUser(adGroup);

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
            var result = true;

            try
            {
                string permissionLevel = "Read";

                foreach (string group in groups.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
                {

                    var adGroup = ctx.Web.EnsureUser(group);
                    ctx.Load(adGroup);

                    var spGroup = ctx.Web.AssociatedMemberGroup;
                    spGroup.Users.AddUser(adGroup);

                    var writeDefinition = ctx.Web.RoleDefinitions.GetByName(permissionLevel);
                    var roleDefCollection = new RoleDefinitionBindingCollection(ctx) { writeDefinition };
                    var newRoleAssignment = ctx.Web.RoleAssignments.Add(adGroup, roleDefCollection);

                    ctx.Load(spGroup, x => x.Users);
                    ctx.ExecuteQuery();

                    // this prevents the Hub Visitor group from being added to site permissions
                    ctx.Load(ctx.Site);
                    ctx.Site.CanSyncHubSitePermissions = false;
                }
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            return Task.FromResult(result);
        }
    }
}