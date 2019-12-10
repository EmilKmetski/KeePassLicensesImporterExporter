using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using KeePass.Forms;
using KeePass.Plugins;
using KeePass.Resources;
using KeePass.UI;
using KeePassLib;
using KeePassLib.Security;
using KeePassLib.Utility;
using KeePassLicensesImporterExporter.Models;

namespace KeePassLicensesImporterExporter
{
    public sealed class KeePassLicensesImporterExporterExt : Plugin
    {

        private IPluginHost m_host = null;
        private DataTable licensesTable = new DataTable();
        public override bool Initialize(IPluginHost host)
        {
            if (host == null) return false; // Fail; we need the host
            m_host = host;
            return true; // Initialization successful
        }

        public override ToolStripMenuItem GetMenuItem(PluginMenuType t)
        {
            // Our menu item below is intended for the main location(s),
            // not for other locations like the group or entry menus
            if (t != PluginMenuType.Entry) return null;
            ToolStripMenuItem hardwareLicensesMenu = new ToolStripMenuItem();
            hardwareLicensesMenu.Enabled = true;
            hardwareLicensesMenu.Name = "Licenses";
            hardwareLicensesMenu.AccessibleName = hardwareLicensesMenu.Name;
            hardwareLicensesMenu.Text = hardwareLicensesMenu.Name;
            hardwareLicensesMenu.Visible = true;



            ToolStripMenuItem importLicenses = new ToolStripMenuItem();
            importLicenses.Text = "Import Licenses";
            importLicenses.Click += this.OnMenuImportLicenses;
            hardwareLicensesMenu.DropDownItems.Add(importLicenses);

            ToolStripMenuItem exportLicenses = new ToolStripMenuItem();
            exportLicenses.Text = "Export Licenses";
            exportLicenses.Click += this.OnMenuExportLicenses;
            hardwareLicensesMenu.DropDownItems.Add(exportLicenses);

            ToolStripMenuItem viewLicense = new ToolStripMenuItem();
            viewLicense.Text = "View License";
            viewLicense.Click += this.OnMenuViewLicenses;
            hardwareLicensesMenu.DropDownItems.Add(viewLicense);

            hardwareLicensesMenu.DropDownOpening += delegate (object sender, EventArgs e)
            {
                PwDatabase pd = m_host.Database;
                bool bOpen = ((pd != null) && pd.IsOpen);

                importLicenses.Enabled = bOpen;
                exportLicenses.Enabled = bOpen;
                viewLicense.Enabled = bOpen;
            };

            return hardwareLicensesMenu;
        }
        private void OnMenuImportLicenses(object sender, EventArgs e)
        {
            PwDatabase pd = m_host.Database;
            if ((pd == null) || !pd.IsOpen) { Debug.Assert(false); return; }
            
            //need to add file open dialog here
            string licensesFileFullPath = AppDomain.CurrentDomain.BaseDirectory + "TemplateFileLicenses.xlsx";

            licensesTable = ReadExcelFile.getExcellToDtbl(licensesFileFullPath, "Licenses");
            List<ILicense> licenses = new List<ILicense>();

            var licenseIDs = licensesTable.AsEnumerable().GroupBy(x => x.Field<string>("LicenseId"));

            foreach (var license in licenseIDs)
            {
                AppLicense currentLicense = new AppLicense();
                currentLicense.Id = license.Key;
                List<AppLicenseData> currentLicenseData = new List<AppLicenseData>();

                foreach (var licenseData in licensesTable.AsEnumerable().Where(x => x.Field<string>("LicenseId") == currentLicense.Id))
                {
                    AppLicenseData appLicenseData = new AppLicenseData();
                    appLicenseData.LicenseApplicationName = licenseData.Field<string>(nameof(appLicenseData.LicenseApplicationName));
                    appLicenseData.LicenseApplicationVersion = licenseData.Field<string>(nameof(appLicenseData.LicenseApplicationVersion));
                    appLicenseData.LicenseNumber = licenseData.Field<string>(nameof(appLicenseData.LicenseNumber));
                    appLicenseData.LicenseRegistrationNumber = licenseData.Field<string>(nameof(appLicenseData.LicenseRegistrationNumber));
                    currentLicenseData.Add(appLicenseData);

                }
                currentLicense.LicenseDatas = currentLicenseData;
                licenses.Add(currentLicense);
            }


            PwGroup pgParent = (m_host.MainWindow.GetSelectedGroup() ?? pd.RootGroup);
            // Add a new group licenses if not exist
            PwGroup pg = new PwGroup(true, true, "Licenses", PwIcon.Key);

            licenses.ForEach(x =>
            {

                StringBuilder desctiptionString = new StringBuilder();

                PwEntry pe = new PwEntry(true, true);
                pe.Strings.Set(PwDefs.TitleField, new ProtectedString(pd.MemoryProtection.ProtectTitle, "LicenseId - " + x.Id));
                x.LicenseDatas.ToList().ForEach
                                               (y =>
                                                    {
                                                        desctiptionString.AppendLine(y.LicenseApplicationName + " - " + y.LicenseApplicationVersion);
                                                        pe.Strings.Set(nameof(y.LicenseApplicationName), new ProtectedString(false, y.LicenseApplicationName));
                                                        pe.Strings.Set(nameof(y.LicenseApplicationVersion), new ProtectedString(false, y.LicenseApplicationVersion));
                                                        pe.Strings.Set(nameof(y.LicenseNumber), new ProtectedString(false, y.LicenseNumber));
                                                        pe.Strings.Set(nameof(y.LicenseRegistrationNumber), new ProtectedString(false, y.LicenseRegistrationNumber));

                                                    }
                                               );


                pe.Strings.Set(PwDefs.NotesField, new ProtectedString(pd.MemoryProtection.ProtectNotes, desctiptionString.ToString()));
                pg.AddEntry(pe, true);
            });

            if (!pgParent.Groups.Contains(pg))
                pgParent.AddGroup(pg, true);

            m_host.MainWindow.UpdateUI(false, null, true, pg, true, null, true);
        }
        private void OnMenuExportLicenses(object sender, EventArgs e)
        {
            PwDatabase pd = m_host.Database;
            if ((pd == null) || !pd.IsOpen) { Debug.Assert(false); return; }


        }

        private void OnMenuViewLicenses(object sender, EventArgs e)
        {
            PwDatabase pd = m_host.Database;
            if ((pd == null) || !pd.IsOpen) { Debug.Assert(false); return; }

            PwEntry pe = m_host.MainWindow.GetSelectedEntry(true);
            if (pe.Strings.Get("Title").ReadString().Contains("License"))
            {
                StringBuilder licenseData = new StringBuilder();
                foreach (var item in pe.Strings)
                {
                    licenseData.AppendLine(item.Key + " : " + item.Value.ReadString());
                }
                Clipboard.SetDataObject(licenseData.ToString(), true);
                MessageBox.Show(licenseData.ToString(), "License information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

    }
}
