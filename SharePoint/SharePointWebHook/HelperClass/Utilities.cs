using Microsoft.Azure.KeyVault;
using Microsoft.Azure.KeyVault.Models;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace SharePointWebHook.HelperClass
{
    public class Utilities
    {

        internal static async Task<CloudTable> GetCloudTableByName(string name)
        {
            string cloudStorageAccountConnectionString = System.Environment.GetEnvironmentVariable("AzureWebJobsStorage", EnvironmentVariableTarget.Process);

            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(cloudStorageAccountConnectionString);

            CloudTableClient tableClient = storageAccount.CreateCloudTableClient();

            CloudTable table = tableClient.GetTableReference(name);
            await table.CreateIfNotExistsAsync();

            return table;
        }

        internal static async Task<X509Certificate2> GetCertificateFromKeyvault(string certName)
        {

            AzureServiceTokenProvider azureServiceTokenProvider = new AzureServiceTokenProvider();
            string keyVaultName = System.Environment.GetEnvironmentVariable("KeyVaultName", EnvironmentVariableTarget.Process); 

            KeyVaultClient keyVaultClient = new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(azureServiceTokenProvider.KeyVaultTokenCallback));

            SecretBundle kvSecret = await keyVaultClient.GetSecretAsync($"https://{keyVaultName}.vault.azure.net/secrets/{certName}");
            X509Certificate2 certificate = new X509Certificate2(Convert.FromBase64String(kvSecret.Value), string.Empty,
              X509KeyStorageFlags.MachineKeySet |
              X509KeyStorageFlags.PersistKeySet |
              X509KeyStorageFlags.Exportable);

            return certificate;

        }

        internal static async Task<string> GetSecretFromKeyvault(string secretName)
        {
            AzureServiceTokenProvider azureServiceTokenProvider = new AzureServiceTokenProvider();
            string keyVaultName = System.Environment.GetEnvironmentVariable("KeyVaultName", EnvironmentVariableTarget.Process);

            KeyVaultClient keyVaultClient = new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(azureServiceTokenProvider.KeyVaultTokenCallback));

            SecretBundle kvSecret = await keyVaultClient.GetSecretAsync($"https://{keyVaultName}.vault.azure.net/secrets/{secretName}");

            return kvSecret.Value;

        }
    }
}
