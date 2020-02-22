using SharePointWebHook.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePointWebHook.HelperClass
{
    public static class LoginUtil
    {
        public static async Task<LoginEntity> CertificateLoginDetails()
        {
            string certificateName = System.Environment.GetEnvironmentVariable("CertificateName", EnvironmentVariableTarget.Process);
            string clientID = certificateName.Replace("_", "").Replace("-", "");
            string tenant = System.Environment.GetEnvironmentVariable("Tenant", EnvironmentVariableTarget.Process);

            var loginReturn = new LoginEntity
            {
                ClientId = await Utilities.GetSecretFromKeyvault(clientID),
                Tenant = $"{tenant}.onmicrosoft.com",
                Certificate = await Utilities.GetCertificateFromKeyvault(certificateName)
            };

            return loginReturn;
        }
    }
}
