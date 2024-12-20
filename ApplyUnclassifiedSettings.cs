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

            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(log);
            var graphClient = new GraphServiceClient(auth);

            var result = Common.ApplyLabel(graphClient, labelId, groupId, itemId, requestId, spaceNameEn, spaceNameFr, log);

            if (result.Result == true)
            {
                // do not call method to set Visibility = Public
                //await SetUnclassified(graphClient, groupId, log);

                var scopes = new string[] { $"https://{tenantName}.sharepoint.com/.default" };
                var authManager = new PnP.Framework.AuthenticationManager();
                var accessToken = await auth.GetTokenAsync(new TokenRequestContext(scopes), new System.Threading.CancellationToken());
                var ctx = authManager.GetAccessTokenContext(sharePointUrl, accessToken.Token);

                bool result1 = await Common.UpdateSiteCollectionAdministrator(ctx, SCAGroupName, groupId, log);
                bool result2 = await AddGroupToFullControl(ctx, supportGroupName, log);
                bool result3 = await AddGroupToReadOnly(ctx, readOnlyGroup, log);
                bool result4 = await Common.RemoveOwner(graphClient, groupId, ownerId, log);

                bool success = result1 && result2 && result3 && result4;

                if (success) {
                    await Common.AddToStatusQueue(itemId, log);
                    await Common.AddToEmailQueue(requestId, "unclassified", groupId, spaceNameEn, spaceNameFr, (string)data?.RequesterName, (string)data?.RequesterEmail, log);
                }
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

        //public static Task<bool> UpdateSiteCollectionAdministrator(ClientContext ctx, string GroupLoginName, string groupId, ILogger log) // ClientContext ctx, 
        //{
        //    log.LogInformation("UpdateSiteCollectionAdministrator received a request.");

        //    bool result = true;

        //    try
        //    {
        //        ctx.Load(ctx.Web);
        //        ctx.Load(ctx.Site);
        //        ctx.Load(ctx.Site.RootWeb);
        //        ctx.Load(ctx.Web.AssociatedOwnerGroup.Users);
        //        ctx.ExecuteQuery();

        //        // add dgcx_sca as Administrator
        //        List<UserEntity> admins = new List<UserEntity>();
        //        UserEntity adminUserEntity = new UserEntity();
        //        adminUserEntity.LoginName = GroupLoginName;
        //        admins.Add(adminUserEntity);
        //        ctx.Site.RootWeb.AddAdministrators(admins, true);

        //        // remove dgcx_sca from the owner group
        //        ctx.Web.AssociatedOwnerGroup.Users.RemoveByLoginName(GroupLoginName);

        //        // remove the owner group
        //        string loginName = $"c:0o.c|federateddirectoryclaimprovider|{groupId}_o";
        //        log.LogInformation($"Remove loginName = {loginName}");
        //        UserEntity ownerGroupEntity = new UserEntity();
        //        ownerGroupEntity.LoginName = loginName;
        //        ctx.Site.RootWeb.RemoveAdministrator(ownerGroupEntity);
        //        log.LogInformation($"Done!");
        //    }
        //    catch (Exception e)
        //    {
        //        log.LogError($"Message: {e.Message}");
        //        if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
        //        log.LogError($"StackTrace: {e.StackTrace}");
        //        result = false;
        //    }

        //    log.LogInformation("UpdateSiteCollectionAdministrator processed a request.");

        //    return Task.FromResult(result);
        //}

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