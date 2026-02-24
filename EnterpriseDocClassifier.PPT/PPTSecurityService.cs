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
            dynamic properties = pres.CustomDocumentProperties;
            try { properties[MetadataPropertyName].Value = label.Name; }
            catch { properties.Add(MetadataPropertyName, false, MsoDocProperties.msoPropertyTypeString, label.Name); }

            foreach (Microsoft.Office.Interop.PowerPoint.Design design in pres.Designs)
            {
                var master = design.SlideMaster;
                Microsoft.Office.Interop.PowerPoint.Shape existingShape = null;

                foreach (Microsoft.Office.Interop.PowerPoint.Shape shp in master.Shapes)
                {
                    if (shp.Name == "DocClassifierLabel") { existingShape = shp; break; }
                }

                if (existingShape == null)
                {
                    // 1. Calculate Y (Top/Bottom)
                    bool isTop = label.Marker.Placement.StartsWith("Top");
                    float yPos = isTop ? 10 : pres.PageSetup.SlideHeight - 40;

                    // 2. Calculate X (Left/Center/Right)
                    float xPos = 0;
                    float width = pres.PageSetup.SlideWidth;
                    var alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;

                    if (label.Marker.Placement.Contains("Left"))
                    {
                        alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                        xPos = 20; // Slight margin from left edge
                        width -= 40;
                    }
                    else if (label.Marker.Placement.Contains("Right"))
                    {
                        alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignRight;
                        xPos = -20; // Slide right edge adjustment
                        width -= 20;
                    }

                    existingShape = master.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, xPos, yPos, width, 30);
                    existingShape.Name = "DocClassifierLabel";

                    // Set Alignment
                    existingShape.TextFrame.TextRange.ParagraphFormat.Alignment = alignment;
                }

                // Keep the standard text/color applying code below this exactly as it was:
                existingShape.TextFrame.TextRange.Text = label.Marker.Text;
                existingShape.TextFrame.TextRange.Font.Size = label.Marker.FontSize;

                System.Drawing.Color sysColor = System.Drawing.ColorTranslator.FromHtml(label.Marker.FontColor);
                existingShape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(sysColor);
            }
        }
    }
}