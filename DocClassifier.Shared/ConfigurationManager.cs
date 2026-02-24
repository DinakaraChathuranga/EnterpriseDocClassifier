using System;
using System.IO;
using System.Text.Json;
using EnterpriseDocClassifier.Models;

namespace EnterpriseDocClassifier.Core
{
    public static class ConfigurationManager
    {
        // This path is readable by everyone, but only writable by Admins/ManageEngine.
        private static readonly string ConfigPath = @"C:\ProgramData\YourCompany\DocClassifier\config.json";

        public static PluginConfiguration LoadConfig()
        {
            // Security: Always fail-safe. If config is missing, return a safe default.
            if (!File.Exists(ConfigPath))
            {
                return new PluginConfiguration { EnforceClassification = false };
            }

            try
            {
                string jsonString = File.ReadAllText(ConfigPath);
                return JsonSerializer.Deserialize<PluginConfiguration>(jsonString);
            }
            catch
            {
                // In a real scenario, log this error.
                return new PluginConfiguration { EnforceClassification = false };
            }
        }
    }
}