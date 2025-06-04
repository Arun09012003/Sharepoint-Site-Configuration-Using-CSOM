using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharepointSiteConfiguration.Models
{
    internal class AppSettings
    {
        public string TenantId { get; set; }
        public string ClientId { get; set; }
        public string SiteUrl { get; set; }

        public string ListName { get; set; }
        public string CertificatePath { get; set; }
        public string CertificatePassword { get; set; }
    }
}
