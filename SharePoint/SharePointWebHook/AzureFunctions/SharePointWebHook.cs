using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Entities;
using SharePointWebHook.WebHooks;

namespace SharePointWebHook.AzureFunctions
{
    public static class SharePointWebHook
    {
        private static CloudQueue SharePointContentUriQueue = null;

        [FunctionName("SharePointWebHook")]
        public static async Task<HttpResponseMessage> RunAsync([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)]HttpRequestMessage req, ILogger log)
        {
            log.LogInformation("WebHook was triggered");

            //Grab the validation Token
            string validationToken = req.GetQueryNameValuePairs().FirstOrDefault(q => string.Compare(q.Key, "validationtoken", true) == 0).Value;

            //If a validation token is present, we need to respond within 5 seconds by
            //returning the given validation token. THis only happens when a new 
            // web hook is being added
            if (validationToken != null)
            {
                log.LogInformation($"Validation token {validationToken} recevied");
                HttpResponseMessage response = req.CreateResponse(HttpStatusCode.OK);
                response.Content = new StringContent(validationToken);
                return response;
            }

            log.LogInformation($"SharePoint triggered the webhook");
            string content = await req.Content.ReadAsStringAsync();
            log.LogInformation($"Received following payload: {content}");

            System.Collections.Generic.List<WebhookNotification> notifications = JsonConvert.DeserializeObject<ResponseModel<WebhookNotification>>(content).Value;
            log.LogInformation($"Found {notifications.Count} notifications");

            if (notifications.Count > 0)
            {
                log.LogInformation("Processing notifications...");
               
                foreach (WebhookNotification notification in notifications)
                {
                    if (SharePointContentUriQueue == null)
                    {
                        string cloudStorageAccountConnectionString = System.Environment.GetEnvironmentVariable("AzureWebJobsStorage", EnvironmentVariableTarget.Process);
                        CloudStorageAccount storageAccount = CloudStorageAccount.Parse(cloudStorageAccountConnectionString);
                        CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                        SharePointContentUriQueue = queueClient.GetQueueReference("processchanges");
                        await SharePointContentUriQueue.CreateIfNotExistsAsync();
                    }
                    string message = JsonConvert.SerializeObject(notification);
                    log.LogInformation($"Adding a message to the queue. Message content: {message}");
                    await SharePointContentUriQueue.AddMessageAsync(new CloudQueueMessage(message));
                    log.LogInformation($"Message added");
                }
            }
            return new HttpResponseMessage(HttpStatusCode.OK);
        }

    }
}
