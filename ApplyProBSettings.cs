using System;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;

namespace appsvc_fnc_dev_scw_sensitivity_dotnet001
{
    public class ApplyProBSettings
    {
        [FunctionName("ApplyProBSettings")]
        public async Task RunAsync([QueueTrigger("prob", Connection = "AzureWebJobsStorage")] string myQueueItem, ILogger log)
        {
            log.LogInformation($"ApplyProBSettings received a request: {myQueueItem}");

            dynamic data = JsonConvert.DeserializeObject(myQueueItem);

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string groupId = data?.groupId;
            string itemId = data?.itemId;
            string labelId = config["proBLabelId"];
            string ownerId = config["ownerId"]; // sv-caupdate@devgcx.ca
            string requestId = data?.Id;
            string spaceNameEn = data?.SpaceName;
            string spaceNameFr = data?.SpaceNameFR;

            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(log);
            var graphClient = new GraphServiceClient(auth);

            var result = Common.ApplyLabel(graphClient, labelId, groupId, itemId, requestId, spaceNameEn, spaceNameFr, log);
            
            if (result.Result == true)
            {
                bool result1 = await SetProB(graphClient, groupId, log);
                bool result2 = await Common.RemoveOwner(graphClient, groupId, ownerId, log);

                bool success = result1 && result2;

                if (success)
                {
                    await Common.AddToStatusQueue(itemId, log);
                    await Common.AddToEmailQueue(requestId, groupId, spaceNameEn, spaceNameFr, (string)data?.RequesterName, (string)data?.RequesterEmail, log);
                }
            }

            log.LogInformation($"ApplyProBSettings processed a request.");
        }

        private static Task<bool> SetProB(GraphServiceClient graphClient, string groupId, ILogger log)
        {
            log.LogInformation("SetProB received a request.");

            bool result = true;

            try
            {
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