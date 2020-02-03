using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KeePassLicensesImporterExporter.Interfaces
{
    public interface ILicense
    {
        string Id { get; set; }
        IEnumerable<ILicenseData> LicenseDatas { get; set; }
    }
}
