using System.Collections.Generic;

namespace EnterpriseDocClassifier.Models
{
    public class PluginConfiguration
    {
        // NEW: "None", "Warn", or "Block"
        public string EnforcementMode { get; set; }
        public string CustomBlockMessage { get; set; }
        public string CustomWarnMessage { get; set; } // NEW: Message for "Warn" mode
        public string DefaultClassificationName { get; set; }
        public List<ClassificationLabel> Classifications { get; set; }
    }

    public class ClassificationLabel
    {
        public string Name { get; set; }
        public string TargetPlatform { get; set; }
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