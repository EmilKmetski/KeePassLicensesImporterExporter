using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KeePassLicensesImporterExporter
{
    public interface ILicense
    {
        string Id { get; set; }
        IEnumerable<ILicenseData> LicenseDatas { get; set; }
    }
}
