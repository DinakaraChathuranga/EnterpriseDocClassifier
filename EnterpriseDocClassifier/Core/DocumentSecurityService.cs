using Microsoft.Office.Interop.Word;
using System;
using Microsoft.Office.Core; // Needed for Document Properties
using EnterpriseDocClassifier.Models;

namespace EnterpriseDocClassifier.Core
{
    public static class DocumentSecurityService
    {
        // We use a hidden property to track classification, not just visual text.
        private const string MetadataPropertyName = "EnterpriseSensitivityLabel";

        public static bool IsDocumentClassified(Document doc)
        {
            try
            {
                // Check if our hidden property exists
                var property = doc.CustomDocumentProperties[MetadataPropertyName];
                return property != null;
            }
            catch
            {
                // If it crashes, the property doesn't exist. Document is unsafe.
                return false;
            }
        }

        public static void ApplyClassification(Document doc, ClassificationLabel label)
        {
            // 1. Write Hidden Metadata (The "Digital" Tag)
            dynamic properties = doc.CustomDocumentProperties;
            try
            {
                properties[MetadataPropertyName].Value = label.Name;
            }
            catch
            {
                properties.Add(MetadataPropertyName, false, MsoDocProperties.msoPropertyTypeString, label.Name);
            }

            // 2. Write Visual Marker (The "Physical" Tag)
            foreach (Section section in doc.Sections)
            {
                // Select Header or Footer based on config
                HeaderFooter target = label.Marker.Placement == "Header"
                    ? section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary]
                    : section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];

                // Clear previous text and set new
                target.Range.Text = label.Marker.Text;
                target.Range.Font.Size = label.Marker.FontSize;
                target.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }
        }
    }
}