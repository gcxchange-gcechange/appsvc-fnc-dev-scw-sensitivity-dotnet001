using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.News.DataModel;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace appsvc_fnc_dev_scw_sensitivity_dotnet001
{
    internal class Common
    {
        public static async Task<Boolean> ApplyLabel(GraphServiceClient graphClient, string labelId, string groupId, string itemId, string requestId, string spaceNameEn, string spaceNameFr, ILogger log)
        {
            // Digital Vault - Vault Digitale Digital Vault                3d277510-cb23-44c1-a9c4-7680fcc237fb
            // PROTECTED B - PROTÉGÉ B        Protect B                    a1ab9d1a-185f-40cc-97d9-e1177019a70b
            // UNCLASSIFIED - NON CLASSIFIÉ   UNCLASSIFIED - NON CLASSIFIÉ d64b0091-505a-4a12-b8e5-9f04b9078a83
            // Protected B MCAS               Protected B - MCAS           e12d86d7-fccd-49e3-8025-027a3c2cbf3a

            log.LogInformation($"ApplyLabel - groupId: {groupId} & labelId: {labelId}");

            var group = new Group
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

                var listItem = new ListItem
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

                listItem = new ListItem
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

        public static async Task<IActionResult> AddToEmailQueue(string requestId, string groupId, string spaceNameEn, string spaceNameFr, string requesterName, string requesterEmail, ILogger log)
        {
            log.LogInformation("AddToEmailQueue received a request.");

            try
            {
                var listItem = new ListItem
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
                            { "Comment", "" }
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

            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
            CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
            CloudQueue queue = queueClient.GetQueueReference(queueName);

            CloudQueueMessage message = new CloudQueueMessage(serializedMessage);
            await queue.AddMessageAsync(message);

            log.LogInformation("AddQueueMessage processed a request.");
        }

        public static async Task<IActionResult> AddToStatusQueue(string itemId, ILogger log)
        {
            log.LogInformation("AddToStatusQueue received a request.");

            try
            {
                var listItem = new ListItem
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