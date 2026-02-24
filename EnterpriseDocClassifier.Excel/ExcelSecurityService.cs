using System;
using Microsoft.Office.Core;
using EnterpriseDocClassifier.Models;

namespace EnterpriseDocClassifier.Excel
{
    public static class ExcelSecurityService
    {
        private const string MetadataPropertyName = "EnterpriseSensitivityLabel";

        // Force C# to use the true Microsoft Excel Workbook type
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
            catch
            {
                return false;
            }
        }

        // Force C# to use the true Microsoft Excel Workbook type
        public static void ApplyClassification(Microsoft.Office.Interop.Excel.Workbook wb, ClassificationLabel label)
        {
            // 1. Write the hidden Metadata Tag 
            dynamic properties = wb.CustomDocumentProperties;
            try
            {
                properties[MetadataPropertyName].Value = label.Name;
            }
            catch
            {
                properties.Add(MetadataPropertyName, false, MsoDocProperties.msoPropertyTypeString, label.Name);
            }

            // 2. Apply Visual Marker to every sheet using the true Microsoft Worksheet type
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Worksheets)
            {
                string formatCode = $"&{label.Marker.FontSize} ";

                if (label.Marker.Placement == "Header")
                    sheet.PageSetup.CenterHeader = formatCode + label.Marker.Text;
                else
                    sheet.PageSetup.CenterFooter = formatCode + label.Marker.Text;
            }
        }
    }
}