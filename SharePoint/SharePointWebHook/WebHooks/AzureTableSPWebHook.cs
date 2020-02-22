using Microsoft.WindowsAzure.Storage.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointWebHook.WebHooks
{
    public class AzureTableSPWebHook : TableEntity
    {
        public string LastChangeToken { get; set; }

        //stateID == PartitionKey
        //listId == rowKey
        public AzureTableSPWebHook(string stateId, string listId, string lastChangeToken)
        {
            PartitionKey = stateId;
            RowKey = listId;
            LastChangeToken = lastChangeToken;
        }

        public AzureTableSPWebHook() { }
    }
}
