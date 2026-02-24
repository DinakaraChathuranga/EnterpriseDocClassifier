using System;
using System.Windows.Forms;
using EnterpriseDocClassifier.Core;

namespace EnterpriseDocClassifier.Excel
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;

            // FIX: Explicitly cast to AppEvents_Event to remove the ambiguity error
            ((Microsoft.Office.Interop.Excel.AppEvents_Event)this.Application).NewWorkbook += Application_NewWorkbook;
            this.Application.WorkbookActivate += Application_WorkbookActivate;
        }

        private void Application_NewWorkbook(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            var config = ConfigurationManager.LoadConfig();

            if (config != null && !string.IsNullOrEmpty(config.DefaultClassificationName))
            {
                var defaultTag = config.Classifications.Find(c => c.Name == config.DefaultClassificationName);
                if (defaultTag != null)
                {
                    ExcelSecurityService.ApplyClassification(Wb, defaultTag);
                    Globals.Ribbons.ExcelClassificationRibbon.SyncRibbonUI(Wb);
                }
            }
        }

        private void Application_WorkbookActivate(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            Globals.Ribbons.ExcelClassificationRibbon.SyncRibbonUI(Wb);
        }

        private void Application_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            var config = ConfigurationManager.LoadConfig();
            if (config == null || string.IsNullOrEmpty(config.EnforcementMode) || config.EnforcementMode == "None") return;

            bool isClassified = ExcelSecurityService.IsWorkbookClassified(Wb);

            if (!isClassified)
            {
                if (config.EnforcementMode == "Warn")
                {
                    string msg = string.IsNullOrWhiteSpace(config.CustomWarnMessage)
                        ? "Warning: You are attempting to save an unclassified spreadsheet. Do you wish to proceed?"
                        : config.CustomWarnMessage;

                    DialogResult result = MessageBox.Show(msg, "Data Loss Prevention Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.No) Cancel = true;
                }
                else if (config.EnforcementMode == "Block")
                {
                    string msg = string.IsNullOrWhiteSpace(config.CustomBlockMessage)
                        ? "Organization Policy: You must select a Sensitivity Level before saving."
                        : config.CustomBlockMessage;

                    MessageBox.Show(msg, "Data Loss Prevention Block", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Cancel = true;
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;

            // FIX: Explicitly cast for cleanup
            ((Microsoft.Office.Interop.Excel.AppEvents_Event)this.Application).NewWorkbook -= Application_NewWorkbook;
            this.Application.WorkbookActivate -= Application_WorkbookActivate;
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