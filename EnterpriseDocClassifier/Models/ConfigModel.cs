using System.Collections.Generic;

namespace EnterpriseDocClassifier.Models
{
    // 1. The Root object: Matches the entire JSON file
    public class PluginConfiguration
    {
        public bool EnforceClassification { get; set; }
        public List<ClassificationLabel> Classifications { get; set; }
    }

    // 2. The Label object: Matches a specific tag (e.g., "Internal")
    public class ClassificationLabel
    {
        public string Name { get; set; }        // Text shown in Ribbon
        public DocumentMarker Marker { get; set; } // The visual watermark settings
    }

    // 3. The Marker object: How the tag looks on the paper
    public class DocumentMarker
    {
        public string Text { get; set; }        // Text to print (e.g. "CONFIDENTIAL")
        public string Placement { get; set; }   // "Header" or "Footer"
        public int FontSize { get; set; }
        public string FontColor { get; set; }   // Hex code like "#FF0000"
    }
}