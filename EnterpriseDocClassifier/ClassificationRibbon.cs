using Microsoft.Office.Tools.Ribbon;
using EnterpriseDocClassifier.Core;
using EnterpriseDocClassifier.Models;

namespace EnterpriseDocClassifier
{
    public partial class ClassificationRibbon
    {
        private void ClassificationRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // 1. Load configuration
            var config = ConfigurationManager.LoadConfig();

            // 2. Create buttons for each tag in the JSON
            if (config != null && config.Classifications != null)
            {
                foreach (var label in config.Classifications)
                {
                    RibbonDropDownItem item = this.Factory.CreateRibbonDropDownItem();
                    item.Label = label.Name;
                    item.Tag = label; // We hide the whole label object inside the button
                    dropDownSensitivity.Items.Add(item);
                }
            }
        }

        private void dropDownSensitivity_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            // 1. Get the selected item
            var selection = dropDownSensitivity.SelectedItem;
            if (selection == null) return;

            // 2. Retrieve the hidden label data
            var label = (ClassificationLabel)selection.Tag;

            // 3. Apply it to the active document
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            DocumentSecurityService.ApplyClassification(doc, label);
        }
    }
}