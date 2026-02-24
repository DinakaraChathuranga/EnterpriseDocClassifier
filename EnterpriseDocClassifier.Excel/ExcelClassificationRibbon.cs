using System;
using System.Windows.Forms;
using EnterpriseDocClassifier.Core; // This connects to your Shared Brain

namespace EnterpriseDocClassifier.Excel
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
        }

        // We use the FULL path here to avoid clashing with your project name
        private void Application_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            var config = ConfigurationManager.LoadConfig();

            if (config == null || !config.EnforceClassification) return;

            bool isClassified = ExcelSecurityService.IsWorkbookClassified(Wb);

            if (!isClassified)
            {
                MessageBox.Show(
                    "Organization Policy: You must select a Document Sensitivity Level from the Ribbon before saving.",
                    "Data Loss Prevention Block",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                Cancel = true;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}