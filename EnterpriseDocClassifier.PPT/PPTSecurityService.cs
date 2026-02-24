using System;
using Microsoft.Office.Core;
using EnterpriseDocClassifier.Models;

namespace EnterpriseDocClassifier.PPT
{
    public static class PPTSecurityService
    {
        private const string MetadataPropertyName = "EnterpriseSensitivityLabel";

        public static bool IsPresentationClassified(Microsoft.Office.Interop.PowerPoint.Presentation pres)
        {
            try
            {
                dynamic properties = pres.CustomDocumentProperties;
                foreach (dynamic property in properties)
                {
                    if (property.Name == MetadataPropertyName) return true;
                }
                return false;
            }
            catch { return false; }
        }

        public static void ApplyClassification(Microsoft.Office.Interop.PowerPoint.Presentation pres, ClassificationLabel label)
        {
            // 1. Write the hidden Metadata Tag 
            dynamic properties = pres.CustomDocumentProperties;
            try { properties[MetadataPropertyName].Value = label.Name; }
            catch { properties.Add(MetadataPropertyName, false, MsoDocProperties.msoPropertyTypeString, label.Name); }

            // 2. Apply Visual Marker to the Slide Master (Displays on all slides)
            foreach (Microsoft.Office.Interop.PowerPoint.Design design in pres.Designs)
            {
                var master = design.SlideMaster;
                Microsoft.Office.Interop.PowerPoint.Shape existingShape = null;

                // Check if we already added a security tag previously
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shp in master.Shapes)
                {
                    if (shp.Name == "DocClassifierLabel") { existingShape = shp; break; }
                }

                // If not, create a new text box
                if (existingShape == null)
                {
                    // Calculate Y position based on Header vs Footer
                    float yPos = label.Marker.Placement == "Header" ? 10 : pres.PageSetup.SlideHeight - 40;

                    existingShape = master.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, yPos, pres.PageSetup.SlideWidth, 30);
                    existingShape.Name = "DocClassifierLabel";
                }

                // Apply text, size, color, and alignment
                existingShape.TextFrame.TextRange.Text = label.Marker.Text;
                existingShape.TextFrame.TextRange.Font.Size = label.Marker.FontSize;

                System.Drawing.Color sysColor = System.Drawing.ColorTranslator.FromHtml(label.Marker.FontColor);
                existingShape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(sysColor);
                existingShape.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            }
        }
    }
}