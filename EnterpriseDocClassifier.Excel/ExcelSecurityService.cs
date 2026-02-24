using System;
using Microsoft.Office.Core;
using EnterpriseDocClassifier.Models;

namespace EnterpriseDocClassifier.Excel
{
    public static class ExcelSecurityService
    {
        private const string MetadataPropertyName = "EnterpriseSensitivityLabel";

        public static bool IsWorkbookClassified(Microsoft.Office.Interop.Excel.Workbook wb)
        {
            return !string.IsNullOrEmpty(GetWorkbookClassification(wb));
        }

        // NEW: Retrieves the current classification tag for UI syncing
        public static string GetWorkbookClassification(Microsoft.Office.Interop.Excel.Workbook wb)
        {
            try
            {
                dynamic properties = wb.CustomDocumentProperties;
                foreach (dynamic property in properties)
                {
                    if (property.Name == MetadataPropertyName) return property.Value;
                }
                return null;
            }
            catch { return null; }
        }

        public static void ApplyClassification(Microsoft.Office.Interop.Excel.Workbook wb, ClassificationLabel label)
        {
            dynamic properties = wb.CustomDocumentProperties;
            try { properties[MetadataPropertyName].Value = label.Name; }
            catch { properties.Add(MetadataPropertyName, false, MsoDocProperties.msoPropertyTypeString, label.Name); }

            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Worksheets)
            {
                string hexColor = label.Marker.FontColor.Replace("#", "");
                string formatCode = $"&{label.Marker.FontSize}&K{hexColor}{label.Marker.Text}";

                sheet.PageSetup.LeftHeader = ""; sheet.PageSetup.CenterHeader = ""; sheet.PageSetup.RightHeader = "";
                sheet.PageSetup.LeftFooter = ""; sheet.PageSetup.CenterFooter = ""; sheet.PageSetup.RightFooter = "";

                switch (label.Marker.Placement)
                {
                    case "Top Left": sheet.PageSetup.LeftHeader = formatCode; break;
                    case "Top Center": sheet.PageSetup.CenterHeader = formatCode; break;
                    case "Top Right": sheet.PageSetup.RightHeader = formatCode; break;
                    case "Bottom Left": sheet.PageSetup.LeftFooter = formatCode; break;
                    case "Bottom Center": sheet.PageSetup.CenterFooter = formatCode; break;
                    case "Bottom Right": sheet.PageSetup.RightFooter = formatCode; break;
                }
            }
        }
    }
}