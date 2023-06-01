//using System;
//using System.Threading.Tasks;
//using Microsoft.AspNetCore.Mvc;
//using Microsoft.Azure.WebJobs;
//using Microsoft.Extensions.Configuration;
//using Microsoft.Extensions.Logging;
//using Microsoft.Graph;
//using Newtonsoft.Json;

//namespace appsvc_fnc_dev_scw_sensitivity_dotnet001
//{
//    public class ApplyProBSettings
//    {
//        [FunctionName("ApplyProBSettings")]
//        public async Task RunAsync([QueueTrigger("prob", Connection = "AzureWebJobsStorage")]string myQueueItem, ILogger log)
//        {
//            log.LogInformation($"ApplyProBSettings received a request: {myQueueItem}");

//            dynamic data = JsonConvert.DeserializeObject(myQueueItem);

//            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

//            string groupId = data?.groupId;
//            string labelId = config["proBLabelId"];
//            string DisplayName = data?.DisplayName;
//            string requestId = data?.Id;

//            string itemId = data?.itemId;

//            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(log);
//            var graphClient = new GraphServiceClient(auth);

//            var result = Common.ApplyLabel(graphClient, labelId, groupId, itemId, requestId, DisplayName, log);

//            if (result.Result == true)
//            {
//                await SetProB(graphClient, groupId, log);
//                await Common.RemoveOwner(graphClient, groupId, "e4b36075-bb6a-4acf-badb-076b0c3d8d90", log);
//                await Common.AddToEmailQueue(requestId, groupId, DisplayName, (string)data?.RequesterName, (string)data?.RequesterEmail, log);
//            }

//            log.LogInformation($"ApplyProBSettings processed a request.");
//        }

//        private static async Task<IActionResult> SetProB(GraphServiceClient graphClient, string groupId, ILogger log)
//        {
//            log.LogInformation("SetProB received a request.");

//            try
//            {
//                var group = new Microsoft.Graph.Group { Visibility = "Private" };
//                await graphClient.Groups[groupId].Request().UpdateAsync(group);
//            }
//            catch (Exception e)
//            {
//                log.LogError($"Message: {e.Message}");
//                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
//                log.LogError($"StackTrace: {e.StackTrace}");
//            }

//            log.LogInformation("SetProB processed a request.");

//            return new OkResult();
//        }
//    }
//}