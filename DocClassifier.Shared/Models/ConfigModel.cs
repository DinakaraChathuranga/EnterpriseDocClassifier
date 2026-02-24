using System.Collections.Generic;

namespace EnterpriseDocClassifier.Models
{
    public class PluginConfiguration
    {
        public bool EnforceClassification { get; set; }
        public string CustomBlockMessage { get; set; } // This fixes the CustomBlockMessage error
        public List<ClassificationLabel> Classifications { get; set; }
    }

    public class ClassificationLabel
    {
        public string Name { get; set; }
        public string TargetPlatform { get; set; } // This fixes the TargetPlatform error
        public DocumentMarker Marker { get; set; }
    }

    public class DocumentMarker
    {
        public string Text { get; set; }
        public string Placement { get; set; }
        public int FontSize { get; set; }
        public string FontColor { get; set; }
    }
}