using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
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
        public static async Task<string> ApplyLabel(GraphServiceClient graphClient, string labelId, string groupId, ILogger log)
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
            }

            return "true";
        }

        public static async Task<IActionResult> AddToEmailQueue(string connectionString, string requestId, string groupId, string displayName, string requesterName, string requesterEmail, ILogger log)
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
                            { "SpaceName", displayName },
                            { "RequesterName", requesterName },
                            { "RequesterEmail", requesterEmail },
                            { "Status", "Team Created" },
                            { "Comment", "" }
                        }
                    }
                };

                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
                CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                CloudQueue queue = queueClient.GetQueueReference("email");
                
                string serializedMessage = JsonConvert.SerializeObject(listItem.Fields.AdditionalData); 

                CloudQueueMessage message = new CloudQueueMessage(serializedMessage);
                await queue.AddMessageAsync(message);
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

        public static async Task<IActionResult> RemoveOwner(GraphServiceClient graphClient, string groupId, string userId, ILogger log)
        {
            log.LogInformation("RemoveOwner received a request.");

            try
            {
                await graphClient.Groups[groupId].Owners[userId].Reference.Request().DeleteAsync();
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation("RemoveOwner processed a request.");
            return new OkResult();
        }



       



    }
}