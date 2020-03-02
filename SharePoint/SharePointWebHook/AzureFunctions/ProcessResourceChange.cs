using System;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;
using SharePointWebHook.Entities;
using SharePointWebHook.HelperClass;
using SharePointWebHook.WebHooks;

namespace SharePointWebHook.AzureFunctions
{
    public static class ProcessResourceChange
    {
        [FunctionName("ProcessResourceChange")]
        public static async Task RunAsync([QueueTrigger("processchanges", Connection = "AzureWebJobsStorage")]WebhookNotification notificationModel, ILogger log)
        {
            log.LogInformation($"ProcesResourceChange queue trigger function process Site:{notificationModel.SiteUrl} Resource{notificationModel.Resource} ");

            string tenant = System.Environment.GetEnvironmentVariable("Tenant", EnvironmentVariableTarget.Process);
            string siteUrl = $"https://{tenant}.sharepoint.com{notificationModel.SiteUrl}";

            log.LogInformation("Getting Azure SharePoint Token");
            LoginEntity loginDetails = await LoginUtil.CertificateLoginDetails();
            OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();

            log.LogInformation("Connecting to SharePoint");
            using (ClientContext clientContext = authManager.GetAzureADAppOnlyAuthenticatedContext(siteUrl, loginDetails.ClientId, loginDetails.Tenant, loginDetails.Certificate))
            {
                log.LogInformation("Getting SharePoint List");
                Guid listId = new Guid(notificationModel.Resource);
                List list = clientContext.Web.Lists.GetById(listId);

                // grab the changes to the provided list using the GetChanges method 
                // on the list. Only request Item changes as that's what's supported via
                // the list web hooks
                ChangeQuery changeQuery = new ChangeQuery(false, false)
                {
                    Item = true,
                    Add = true,
                    DeleteObject = false,
                    Update = true,
                    FetchLimit = 1000
                };

                ChangeToken lastChangeToken = null;
                if (null == notificationModel.ClientState)
                {
                    throw new ApplicationException("Webhook doesn't contain a Client State");
                }

                Guid id = new Guid(notificationModel.ClientState);
                log.LogInformation("Checking Database for Change Token");
                AzureTableSPWebHook listWebHookRow = await AzureTable.GetListWebHookByID(id, listId);

                if (!string.IsNullOrEmpty(listWebHookRow.LastChangeToken))
                {
                    log.LogInformation("Change Token found");
                    lastChangeToken = new ChangeToken
                    {
                        StringValue = listWebHookRow.LastChangeToken
                    };
                }
                //Start pulling down the changes
                bool allChangesRead = false;

                do
                {
                    if (lastChangeToken == null)
                    {
                        log.LogInformation("Change Token not found grabbing the last 5 days of changes");
                        //If none found, grab only the last 5 day changes.
                        lastChangeToken = new ChangeToken
                        {
                            StringValue = string.Format("1;3;{0};{1};-1", notificationModel.Resource, DateTime.Now.AddDays(-5).ToUniversalTime().Ticks.ToString())
                        };
                    }

                    //Assing the change token to the query..this determins from what point in time we'll receive changes
                    changeQuery.ChangeTokenStart = lastChangeToken;

                    ChangeCollection changes = list.GetChanges(changeQuery);
                    clientContext.Load(list);
                    clientContext.Load(changes);
                    await clientContext.ExecuteQueryAsync();

                    if (changes.Count > 0)
                    {
                        log.LogInformation($"Found {changes.Count} changes");
                        foreach (Change change in changes)
                        {
                            lastChangeToken = change.ChangeToken;

                            if (change is ChangeItem item)
                            {
                                log.LogInformation($"Change {change.ChangeType} on Item:{item.ItemId} on List:{list.Title} Site:{notificationModel.SiteUrl}");
                            }
                        }

                        if (changes.Count < changeQuery.FetchLimit)
                        {
                            allChangesRead = true;
                        }
                    }
                    else
                    {
                        allChangesRead = true;
                    }
                } while (allChangesRead == false);

                if (!listWebHookRow.LastChangeToken.Equals(lastChangeToken.StringValue, StringComparison.InvariantCultureIgnoreCase))
                {
                    listWebHookRow.LastChangeToken = lastChangeToken.StringValue;
                    log.LogInformation("Updating change token");
                    await AzureTable.InsertOrReplaceListWebHook(listWebHookRow);
                }

                if (notificationModel.ExpirationDateTime.AddDays(-30) < DateTime.Now)
                {
                    bool updateResult = list.UpdateWebhookSubscription(new Guid(notificationModel.SubscriptionId), DateTime.Now.AddMonths(2));

                    if (!updateResult)
                    {
                        log.LogError($"The expiration date of web hook {notificationModel.SubscriptionId} with endpoint 'https://{Environment.GetEnvironmentVariable("WEBSITE_HOSTNAME")}/SharePointWebHook' could not be updated");
                    }
                    else
                    {
                        log.LogInformation($"The expiration date of web hook {notificationModel.SubscriptionId} with endpoint 'https://{Environment.GetEnvironmentVariable("WEBSITE_HOSTNAME")}/SharePointWebHook' successfully updated until {DateTime.Now.AddMonths(2).ToLongDateString()} ");
                    }
                }
            }
        }
    }
}