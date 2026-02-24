using Microsoft.Office.Tools.Ribbon;
using EnterpriseDocClassifier.Core;
using EnterpriseDocClassifier.Models;

namespace EnterpriseDocClassifier
{
    public partial class ClassificationRibbon
    {
        private void ClassificationRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            var config = ConfigurationManager.LoadConfig();

            if (config != null && config.Classifications != null)
            {
                foreach (var label in config.Classifications)
                {
                    if (label.TargetPlatform == "All" || label.TargetPlatform == "Word")
                    {
                        RibbonDropDownItem item = this.Factory.CreateRibbonDropDownItem();
                        item.Label = label.Name;
                        item.Tag = label;
                        dropDownSensitivity.Items.Add(item);
                    }
                }
            }
        }

        private void dropDownSensitivity_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            var selection = dropDownSensitivity.SelectedItem;
            if (selection == null) return;

            var label = (ClassificationLabel)selection.Tag;
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            DocumentSecurityService.ApplyClassification(doc, label);
        }

        // NEW: Ribbon Sync Method
        public void SyncRibbonUI(Microsoft.Office.Interop.Word.Document doc)
        {
            string currentTag = DocumentSecurityService.GetDocumentClassification(doc);

            if (string.IsNullOrEmpty(currentTag))
            {
                dropDownSensitivity.SelectedItem = null;
                return;
            }

            foreach (RibbonDropDownItem item in dropDownSensitivity.Items)
            {
                if (item.Label == currentTag)
                {
                    dropDownSensitivity.SelectedItem = item;
                    break;
                }
            }
        }
    }
}