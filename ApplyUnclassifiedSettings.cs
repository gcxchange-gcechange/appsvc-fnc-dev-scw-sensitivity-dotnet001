using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using PnP.Framework.Entities;

namespace appsvc_fnc_dev_scw_sensitivity_dotnet001
{
    public class ApplyUnclassifiedSettings
    {
        [FunctionName("ApplyUnclassifiedSettings")]
        public async Task RunAsync([QueueTrigger("unclassified", Connection = "AzureWebJobsStorage")]string myQueueItem, ILogger log, ExecutionContext functionContext)
        {
            log.LogInformation($"ApplyUnclassifiedSettings received a request: {myQueueItem}");

            dynamic data = JsonConvert.DeserializeObject(myQueueItem);

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string groupId = data?.groupId;
            string labelId = config["unclassifiedLabelId"];
            string DisplayName = data?.DisplayName;
            string requestId = data?.Id;

            string itemId = data?.itemId;

            string keyVaultUrl = config["keyVaultUrl"];
            string sharePointUrl = config["sharePointUrl"] + requestId;
            string clientId = config["clientId"];
            string certificateName = config["certificateName"];
            string tenantId = config["tenantId"];

            string SCAGroupName = config["sca_login_name"];
            string SupportGroupName = config["support_group_login_name"];

            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(log);
            var graphClient = new GraphServiceClient(auth);
            
            var result = Common.ApplyLabel(graphClient, labelId, groupId, itemId, requestId, DisplayName, log);

            if (result.Result == true)
            {
                await SetUnclassified(graphClient, groupId, log);
                await Common.RemoveOwner(graphClient, groupId, "e4b36075-bb6a-4acf-badb-076b0c3d8d90", log);

                var ctx = Auth.GetContextByCertificate(sharePointUrl, keyVaultUrl, certificateName, clientId, tenantId, log);
                await AddSiteCollectionAdministrator(ctx, SCAGroupName, log);

                await AddPermissionLevel(ctx, SupportGroupName, log);
                await Common.AddToEmailQueue(requestId, groupId, DisplayName, (string)data?.RequesterName, (string)data?.RequesterEmail, log);
            }

            log.LogInformation($"ApplyUnclassifiedSettings processed a request.");
        }

        private static async Task<IActionResult> SetUnclassified(GraphServiceClient graphClient, string groupId, ILogger log)
        {
            log.LogInformation("SetUnclassified received a request.");


            // good to know:
            // var user = ctx.Site.RootWeb.SiteUsers.GetByLoginName($"c:0o.c|federateddirectoryclaimprovider|{groupId}_o");

            try
            {
                var group = new Microsoft.Graph.Group { Visibility = "Public" };
                await graphClient.Groups[groupId].Request().UpdateAsync(group);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation("SetUnclassified processed a request.");

            return new OkResult();
        }

        public static Task<bool> AddSiteCollectionAdministrator(ClientContext ctx, string GroupLoginName, ILogger log)
        {
            var result = true;

            try
            {
                ctx.Load(ctx.Web);
                ctx.Load(ctx.Site);
                ctx.Load(ctx.Site.RootWeb);
                ctx.ExecuteQuery();



                //ctx.Site.RootWeb.AddUserToGroup(ctx.Site.RootWeb.AssociatedMemberGroup, "i:0#.f|tenant|gabriela.morenoramirez@devgcx.onmicrosoft.com");
                //ctx.Site.RootWeb.AddUserToGroup(ctx.Site.RootWeb.AssociatedMemberGroup, "i:0#.f|tenant|vanmathy.raviraj@devgcx.onmicrosoft.com");

                var mems = ctx.Site.RootWeb.GetMembers();
                log.LogInformation("GetMembers()");
                foreach (var member in mems)
                {
                    // member.LoginName: i:0#.f|membership|ilia.salem@devgcx.onmicrosoft.com
                    // member.LoginName: i:0#.f|membership|gabriela.morenoramirez@devgcx.onmicrosoft.com
                    // member.LoginName: i:0#.f|membership|vanmathy.raviraj@devgcx.onmicrosoft.com

                    log.LogInformation($"member.LoginName: {member.LoginName}");

                }

                ////////////////////////////////////////////////////////////////////////////////////////////

                List<UserEntity> admins = new List<UserEntity>();
                UserEntity adminUserEntity = new UserEntity();

                adminUserEntity.LoginName = GroupLoginName;
                admins.Add(adminUserEntity);

                ctx.Site.RootWeb.AddAdministrators(admins, true);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            return Task.FromResult(result);
        }

        public static Task<bool> AddPermissionLevel(ClientContext ctx, string GroupLoginName, ILogger log)
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
                var roleDefCollection = new RoleDefinitionBindingCollection(ctx);
                roleDefCollection.Add(writeDefinition);
                var newRoleAssignment = ctx.Web.RoleAssignments.Add(adGroup, roleDefCollection);

                ctx.Load(spGroup, x => x.Users);
                ctx.ExecuteQuery();
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