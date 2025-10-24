using Azure.Storage.Queues;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PnP.Framework.Entities;

namespace appsvc_fnc_dev_scw_sensitivity_dotnet001
{
    internal class Common
    {
        public static async Task<Boolean> ApplyLabel(GraphServiceClient graphClient, string labelId, string groupId, string itemId, string requestId, string spaceNameEn, string spaceNameFr, ILogger log)
        {
            log.LogInformation($"ApplyLabel - groupId: {groupId} & labelId: {labelId}");

            var group = new Microsoft.Graph.Group
            {
                AssignedLabels = new List<AssignedLabel>()
                {
                    new AssignedLabel { LabelId = labelId }
                }
            };

            try
            {
                var users = await graphClient.Groups[groupId].Request().UpdateAsync(group);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");

                string status = "Failed";

                var listItem = new Microsoft.Graph.ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>()
                        {
                            { "Id", requestId },
                            { "groupId", groupId },
                            { "SpaceName", spaceNameEn },
                            { "SpaceNameFR", spaceNameFr },
                            { "Status", status},
                            { "FunctionApp", "Sensitivity" },
                            { "Method", "ApplyLabel" },
                            { "ErrorMessage", $"{e.Message}" }
                        }
                    }
                };
                await AddQueueMessage("email", JsonConvert.SerializeObject(listItem.Fields.AdditionalData), log);

                listItem = new Microsoft.Graph.ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"Id", itemId},
                            {"Status", status},
                        }
                    }
                };
                await AddQueueMessage("list", JsonConvert.SerializeObject(listItem.Fields.AdditionalData), log);

                return false;
            }

            return true;
        }

        public static async Task<IActionResult> AddToEmailQueue(string requestId, string securityCategory, string groupId, string spaceNameEn, string spaceNameFr, string requesterName, string requesterEmail, ILogger log)
        {
            log.LogInformation("AddToEmailQueue received a request.");

            try
            {
                var listItem = new Microsoft.Graph.ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>()
                        {
                            { "Id", requestId },
                            { "groupId", groupId },
                            { "SpaceName", spaceNameEn },
                            { "SpaceNameFR", spaceNameFr },
                            { "RequesterName", requesterName },
                            { "RequesterEmail", requesterEmail },
                            { "Status", "Team Created" },
                            { "Comment", "" },
                            { "SecurityCategory", securityCategory }
                        }
                    }
                };

                await AddQueueMessage("email", JsonConvert.SerializeObject(listItem.Fields.AdditionalData), log);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation("AddToEmailQueue processed a request.");

            return new OkResult();
        }


        public static Task<bool> UpdateSiteCollectionAdministrator(ClientContext ctx, string GroupLoginName, string groupId, ILogger log) // ClientContext ctx, 
        {
            log.LogInformation("UpdateSiteCollectionAdministrator received a request.");

            bool result = true;

            try
            {
                ctx.Load(ctx.Web);
                ctx.Load(ctx.Site);
                ctx.Load(ctx.Site.RootWeb);
                ctx.Load(ctx.Web.AssociatedOwnerGroup.Users);
                ctx.ExecuteQuery();

                // add dgcx_sca as Administrator
                List<UserEntity> admins = new List<UserEntity>();
                UserEntity adminUserEntity = new UserEntity();
                adminUserEntity.LoginName = GroupLoginName;
                admins.Add(adminUserEntity);

                log.LogInformation($"Add loginName = {GroupLoginName}");
                ctx.Site.RootWeb.AddAdministrators(admins, true);
                log.LogInformation($"Done!");

                // remove dgcx_sca from the owner group
                ctx.Web.AssociatedOwnerGroup.Users.RemoveByLoginName(GroupLoginName);

                // remove the owner group
                string loginName = $"c:0o.c|federateddirectoryclaimprovider|{groupId}_o";
                UserEntity ownerGroupEntity = new UserEntity();
                ownerGroupEntity.LoginName = loginName;

                log.LogInformation($"Remove loginName = {loginName}");
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


        public static async Task<bool> RemoveOwner(GraphServiceClient graphClient, string groupId, string userId, ILogger log)
        {
            log.LogInformation("RemoveOwner received a request.");

            bool result = true;

            try
            {
                await graphClient.Groups[groupId].Owners[userId].Reference.Request().DeleteAsync();
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
                result = false;
            }

            log.LogInformation("RemoveOwner processed a request.");
            
            return result;
        }

        public static async Task AddQueueMessage(string queueName, string serializedMessage, ILogger log)
        {
            log.LogInformation("AddQueueMessage received a request.");

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string connectionString = config["AzureWebJobsStorage"];

            QueueClientOptions options = new QueueClientOptions() { MessageEncoding = QueueMessageEncoding.Base64 };
            QueueClient client = new QueueClient(connectionString, queueName, options);

            await client.SendMessageAsync(serializedMessage);

            log.LogInformation("AddQueueMessage processed a request.");
        }

        public static async Task<IActionResult> AddToStatusQueue(string itemId, ILogger log)
        {
            log.LogInformation("AddToStatusQueue received a request.");

            try
            {
                var listItem = new Microsoft.Graph.ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>()
                        {
                            { "Id", itemId },
                            { "Status", "Complete" },
                        }
                    }
                };

                await AddQueueMessage("status", JsonConvert.SerializeObject(listItem.Fields.AdditionalData), log);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation("AddToStatusQueue processed a request.");

            return new OkResult();
        }
    }
}