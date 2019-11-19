using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KeePassLicensesImporterExporter
{
    public interface ILicenseData
    {
        string LicenseApplicationName { get; set; }
        string LicenseApplicationVersion { get; set; }
        string LicenseNumber { get; set; }
        string LicenseRegistrationNumber { get; set; }

    }
}
