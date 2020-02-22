using Microsoft.Extensions.Logging;
using Microsoft.WindowsAzure.Storage.Table;
using SharePointWebHook.WebHooks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointWebHook.HelperClass
{
    internal static class AzureTable
    {
        private static CloudTable Table { get; set; }

        internal static async Task<AzureTableSPWebHook> GetListWebHookByID(Guid stateId, Guid listId)
        {
            //stateID == PartitionKey
            //listId == rowKey

            if (null == Table)
            {
                Table = await Utilities.GetCloudTableByName("SharePointWebHooks");
            }

            TableOperation retrieve = TableOperation.Retrieve<AzureTableSPWebHook>(stateId.ToString(), listId.ToString());

            TableResult tableResult = await Table.ExecuteAsync(retrieve);

            if (tableResult.Result != null)
            {
                return (AzureTableSPWebHook)tableResult.Result;
            }
            else
            {
                return new AzureTableSPWebHook(stateId.ToString(), listId.ToString(), "");
            }

        }

        internal static async Task<bool> InsertOrReplaceListWebHook(AzureTableSPWebHook listWebHookRow)
        {
            try
            {
                if (null == Table)
                {
                    CloudTable table = await Utilities.GetCloudTableByName("SharePointWebHooks");
                }

                listWebHookRow.ETag = "*";
                TableOperation insertOrReplace = TableOperation.InsertOrReplace(listWebHookRow);
                await Table.ExecuteAsync(insertOrReplace);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
