using System;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using EnterpriseDocClassifier.Models; // Accesses your shared JSON models

namespace EnterpriseDocClassifier.Excel
{
    public static class ExcelSecurityService
    {
        private const string MetadataPropertyName = "EnterpriseSensitivityLabel";

        public static bool IsWorkbookClassified(Workbook wb)
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
            catch
            {
                return false;
            }
        }

        public static void ApplyClassification(Workbook wb, ClassificationLabel label)
        {
            // 1. Write the hidden Metadata Tag (Crucial for DLP)
            dynamic properties = wb.CustomDocumentProperties;
            try
            {
                properties[MetadataPropertyName].Value = label.Name;
            }
            catch
            {
                properties.Add(MetadataPropertyName, false, MsoDocProperties.msoPropertyTypeString, label.Name);
            }

            // 2. Apply Visual Marker to every sheet in the workbook
            foreach (Worksheet sheet in wb.Worksheets)
            {
                // Excel requires specific formatting strings for headers (e.g., &14 for font size 14)
                string formatCode = $"&{label.Marker.FontSize} ";

                if (label.Marker.Placement == "Header")
                    sheet.PageSetup.CenterHeader = formatCode + label.Marker.Text;
                else
                    sheet.PageSetup.CenterFooter = formatCode + label.Marker.Text;
            }
        }
    }
}