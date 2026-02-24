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
            ((Word.ApplicationEvents4_Event)this.Application).NewDocument += Application_NewDocument;
            ((Word.ApplicationEvents4_Event)this.Application).WindowActivate += Application_WindowActivate;
        }

        private void Application_NewDocument(Word.Document Doc)
        {
            var config = ConfigurationManager.LoadConfig();
            if (config != null && !string.IsNullOrEmpty(config.DefaultClassificationName))
            {
                var defaultTag = config.Classifications.Find(c => c.Name == config.DefaultClassificationName);
                if (defaultTag != null)
                {
                    DocumentSecurityService.ApplyClassification(Doc, defaultTag);
                    Globals.Ribbons.ClassificationRibbon.SyncRibbonUI(Doc);
                }
            }
        }

        private void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        {
            Globals.Ribbons.ClassificationRibbon.SyncRibbonUI(Doc);
        }

        private void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            var config = ConfigurationManager.LoadConfig();

            // Uses the new EnforcementMode (Fixes the EnforceClassification error)
            if (config == null || string.IsNullOrEmpty(config.EnforcementMode) || config.EnforcementMode == "None") return;

            bool isClassified = DocumentSecurityService.IsDocumentClassified(Doc);

            if (!isClassified)
            {
                if (config.EnforcementMode == "Warn")
                {
                    string msg = string.IsNullOrWhiteSpace(config.CustomWarnMessage)
                        ? "Warning: You are attempting to save an unclassified document. Do you wish to proceed?"
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
            this.Application.DocumentBeforeSave -= Application_DocumentBeforeSave;
            ((Word.ApplicationEvents4_Event)this.Application).NewDocument -= Application_NewDocument;
            ((Word.ApplicationEvents4_Event)this.Application).WindowActivate -= Application_WindowActivate;
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