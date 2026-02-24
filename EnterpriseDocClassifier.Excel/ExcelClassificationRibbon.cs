using System;
using Microsoft.Office.Tools.Ribbon;
using EnterpriseDocClassifier.Models;
using EnterpriseDocClassifier.Core;

namespace EnterpriseDocClassifier.Excel
{
    public partial class ExcelClassificationRibbon
    {
        private void ExcelClassificationRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            var config = ConfigurationManager.LoadConfig();

            if (config != null && config.Classifications != null)
            {
                foreach (var label in config.Classifications)
                {
                    if (label.TargetPlatform == "All" || label.TargetPlatform == "Excel")
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
            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            ExcelSecurityService.ApplyClassification(wb, label);
        }

        // NEW: Ribbon Sync Method for Excel
        public void SyncRibbonUI(Microsoft.Office.Interop.Excel.Workbook wb)
        {
            string currentTag = ExcelSecurityService.GetWorkbookClassification(wb);

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