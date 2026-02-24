using Microsoft.Office.Interop.Word;
using System;
using Microsoft.Office.Core;
using EnterpriseDocClassifier.Models;

namespace EnterpriseDocClassifier.Core
{
    public static class DocumentSecurityService
    {
        private const string MetadataPropertyName = "EnterpriseSensitivityLabel";

        public static bool IsDocumentClassified(Document doc)
        {
            try
            {
                dynamic properties = doc.CustomDocumentProperties;
                foreach (dynamic property in properties)
                {
                    if (property.Name == MetadataPropertyName) return true;
                }
                return false;
            }
            catch { return false; }
        }

        public static void ApplyClassification(Document doc, ClassificationLabel label)
        {
            dynamic properties = doc.CustomDocumentProperties;
            try { properties[MetadataPropertyName].Value = label.Name; }
            catch { properties.Add(MetadataPropertyName, false, MsoDocProperties.msoPropertyTypeString, label.Name); }

            foreach (Section section in doc.Sections)
            {
                // Determine Header or Footer based on Placement setting
                bool isTop = label.Marker.Placement.StartsWith("Top");
                HeaderFooter targetArea = isTop
                    ? section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary]
                    : section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];

                targetArea.Range.Text = label.Marker.Text;
                targetArea.Range.Font.Size = label.Marker.FontSize;
                System.Drawing.Color sysColor = System.Drawing.ColorTranslator.FromHtml(label.Marker.FontColor);
                targetArea.Range.Font.Color = (WdColor)System.Drawing.ColorTranslator.ToOle(sysColor);

                // Determine Alignment (Left, Center, Right)
                if (label.Marker.Placement.Contains("Left"))
                    targetArea.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                else if (label.Marker.Placement.Contains("Right"))
                    targetArea.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                else
                    targetArea.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }
        }
    }
}