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
            return !string.IsNullOrEmpty(GetPresentationClassification(pres));
        }

        // NEW: Retrieves the current classification tag for UI syncing
        public static string GetPresentationClassification(Microsoft.Office.Interop.PowerPoint.Presentation pres)
        {
            try
            {
                dynamic properties = pres.CustomDocumentProperties;
                foreach (dynamic property in properties)
                {
                    if (property.Name == MetadataPropertyName) return property.Value;
                }
                return null;
            }
            catch { return null; }
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
                    bool isTop = label.Marker.Placement.StartsWith("Top");
                    float yPos = isTop ? 10 : pres.PageSetup.SlideHeight - 40;

                    float xPos = 0;
                    float width = pres.PageSetup.SlideWidth;
                    var alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;

                    if (label.Marker.Placement.Contains("Left"))
                    {
                        alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                        xPos = 20;
                        width -= 40;
                    }
                    else if (label.Marker.Placement.Contains("Right"))
                    {
                        alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignRight;
                        xPos = -20;
                        width -= 20;
                    }

                    existingShape = master.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, xPos, yPos, width, 30);
                    existingShape.Name = "DocClassifierLabel";
                    existingShape.TextFrame.TextRange.ParagraphFormat.Alignment = alignment;
                }

                existingShape.TextFrame.TextRange.Text = label.Marker.Text;
                existingShape.TextFrame.TextRange.Font.Size = label.Marker.FontSize;
                System.Drawing.Color sysColor = System.Drawing.ColorTranslator.FromHtml(label.Marker.FontColor);
                existingShape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(sysColor);
            }
        }
    }
}