using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using EnterpriseDocClassifier.Core;

namespace EnterpriseDocClassifier
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.DocumentBeforeSave += Application_DocumentBeforeSave;
        }

        private void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            var config = ConfigurationManager.LoadConfig();
            if (config == null || !config.EnforceClassification) return;

            bool isClassified = DocumentSecurityService.IsDocumentClassified(Doc);

            if (!isClassified)
            {
                // Custom Error Message implementation
                string msg = string.IsNullOrWhiteSpace(config.CustomBlockMessage)
                    ? "Organization Policy: You must select a Sensitivity Level before saving."
                    : config.CustomBlockMessage;

                MessageBox.Show(msg, "Data Loss Prevention Block", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cancel = true;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.DocumentBeforeSave -= Application_DocumentBeforeSave;
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