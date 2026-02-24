using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
// using DocClassifier.Shared.Core; // Ensure this points to your ConfigurationManager

namespace EnterpriseDocClassifier.Excel
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Listen for Excel's specific save event
            this.Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
        }

        private void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
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

                Cancel = true; // Blocks the Excel save
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