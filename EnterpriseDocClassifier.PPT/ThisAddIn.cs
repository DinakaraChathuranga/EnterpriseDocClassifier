using System;
using System.Windows.Forms;
using EnterpriseDocClassifier.Core;

namespace EnterpriseDocClassifier.PPT
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.PresentationBeforeSave += Application_PresentationBeforeSave;
        }

        private void Application_PresentationBeforeSave(Microsoft.Office.Interop.PowerPoint.Presentation Pres, ref bool Cancel)
        {
            var config = ConfigurationManager.LoadConfig();
            if (config == null || !config.EnforceClassification) return;

            bool isClassified = PPTSecurityService.IsPresentationClassified(Pres);

            if (!isClassified)
            {
                string msg = string.IsNullOrWhiteSpace(config.CustomBlockMessage)
                    ? "Organization Policy: You must select a Sensitivity Level before saving."
                    : config.CustomBlockMessage;

                MessageBox.Show(msg, "Data Loss Prevention Block", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cancel = true;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.PresentationBeforeSave -= Application_PresentationBeforeSave;
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