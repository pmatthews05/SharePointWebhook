using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace SharePointWebHook.Entities
{
    public class LoginEntity
    {
        public string ClientId { get; set; }
        public string Tenant { get; set; }
        public X509Certificate2 Certificate { get; set; }
    }
}
