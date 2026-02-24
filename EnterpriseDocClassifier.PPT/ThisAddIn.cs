using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using EnterpriseDocClassifier.Core;

namespace EnterpriseDocClassifier.PPT
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.PresentationBeforeSave += Application_PresentationBeforeSave;
            this.Application.AfterNewPresentation += Application_AfterNewPresentation;
            this.Application.WindowActivate += Application_WindowActivate;
        }

        private void Application_AfterNewPresentation(PowerPoint.Presentation Pres)
        {
            var config = ConfigurationManager.LoadConfig();

            if (config != null && !string.IsNullOrEmpty(config.DefaultClassificationName))
            {
                var defaultTag = config.Classifications.Find(c => c.Name == config.DefaultClassificationName);
                if (defaultTag != null)
                {
                    PPTSecurityService.ApplyClassification(Pres, defaultTag);
                    Globals.Ribbons.PPTClassificationRibbon.SyncRibbonUI(Pres);
                }
            }
        }

        private void Application_WindowActivate(PowerPoint.Presentation Pres, PowerPoint.DocumentWindow Wn)
        {
            Globals.Ribbons.PPTClassificationRibbon.SyncRibbonUI(Pres);
        }

        private void Application_PresentationBeforeSave(PowerPoint.Presentation Pres, ref bool Cancel)
        {
            var config = ConfigurationManager.LoadConfig();
            if (config == null || string.IsNullOrEmpty(config.EnforcementMode) || config.EnforcementMode == "None") return;

            bool isClassified = PPTSecurityService.IsPresentationClassified(Pres);

            if (!isClassified)
            {
                if (config.EnforcementMode == "Warn")
                {
                    string msg = string.IsNullOrWhiteSpace(config.CustomWarnMessage)
                        ? "Warning: You are attempting to save an unclassified presentation. Do you wish to proceed?"
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
            this.Application.PresentationBeforeSave -= Application_PresentationBeforeSave;
            this.Application.AfterNewPresentation -= Application_AfterNewPresentation;
            this.Application.WindowActivate -= Application_WindowActivate;
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