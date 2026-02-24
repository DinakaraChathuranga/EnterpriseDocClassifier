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
            try
            {
                dynamic properties = wb.CustomDocumentProperties;
                foreach (dynamic property in properties)
                {
                    if (property.Name == MetadataPropertyName) return true;
                }
                return false;
            }
            catch { return false; }
        }

        public static void ApplyClassification(Microsoft.Office.Interop.Excel.Workbook wb, ClassificationLabel label)
        {
            dynamic properties = wb.CustomDocumentProperties;
            try { properties[MetadataPropertyName].Value = label.Name; }
            catch { properties.Add(MetadataPropertyName, false, MsoDocProperties.msoPropertyTypeString, label.Name); }

            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Worksheets)
            {
                // Excel hex code requires stripping the '#'
                string hexColor = label.Marker.FontColor.Replace("#", "");
                string formatCode = $"&{label.Marker.FontSize}&K{hexColor}{label.Marker.Text}";

                // Clear previous headers/footers to avoid duplicates
                sheet.PageSetup.LeftHeader = ""; sheet.PageSetup.CenterHeader = ""; sheet.PageSetup.RightHeader = "";
                sheet.PageSetup.LeftFooter = ""; sheet.PageSetup.CenterFooter = ""; sheet.PageSetup.RightFooter = "";

                // Apply exact positioning
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