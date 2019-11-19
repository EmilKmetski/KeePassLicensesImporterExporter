using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KeePassLicensesImporterExporter.Models
{
    public class AppLicense : ILicense
    {
        public string Id { get ; set; }
        public IEnumerable<ILicenseData> LicenseDatas { get ; set ; }
    }
}
