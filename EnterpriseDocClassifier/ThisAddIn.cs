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
            // Hook into the Word Application's "BeforeSave" event as soon as the plugin loads
            this.Application.DocumentBeforeSave += Application_DocumentBeforeSave;
        }

        private void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            // 1. Load the current security policy
            var config = ConfigurationManager.LoadConfig();

            // If enforcement is disabled in the config, let them save normally
            if (config == null || !config.EnforceClassification)
            {
                return;
            }

            // 2. Check if the document has the required classification metadata
            bool isClassified = DocumentSecurityService.IsDocumentClassified(Doc);

            // 3. If it is NOT classified, block the save
            if (!isClassified)
            {
                MessageBox.Show(
                    "Organization Policy: You must select a Document Sensitivity Level from the Ribbon before saving.",
                    "Data Loss Prevention Block",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                // This is the critical command that stops Word from actually saving the file
                Cancel = true;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Clean up the event listener when Word closes to prevent memory issues
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