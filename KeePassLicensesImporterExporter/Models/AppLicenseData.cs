using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KeePassLicensesImporterExporter.Models
{
    public class AppLicenseData : ILicenseData
    {
        public string LicenseApplicationName { get ; set; }
        public string LicenseApplicationVersion { get; set; }
        public string LicenseNumber { get; set; }
        public string LicenseRegistrationNumber { get; set; }
    }
}
