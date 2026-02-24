using System;
using System.IO;
using System.Text.Json;
using EnterpriseDocClassifier.Models;

namespace EnterpriseDocClassifier.Core
{
    public static class ConfigurationManager
    {
        private static PluginConfiguration _cachedConfig;
        private static DateTime _lastReadTime;
        private static string _configPath;

        public static PluginConfiguration LoadConfig()
        {
            // 1. Registry-Backed Config Path
            if (_configPath == null)
            {
                try
                {
                    // Looks for a registry key pushed by ManageEngine
                    using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\EnterpriseDLP"))
                    {
                        if (key != null) _configPath = key.GetValue("ConfigPath") as string;
                    }
                }
                catch { }

                // Fallback to local path if registry isn't set
                if (string.IsNullOrEmpty(_configPath))
                    _configPath = @"C:\ProgramData\YourCompany\DocClassifier\config.json";
            }

            if (!File.Exists(_configPath)) return null;

            // 2. Memory Caching Logic (Only read disk if file was modified)
            DateTime lastModified = File.GetLastWriteTime(_configPath);
            if (_cachedConfig != null && lastModified == _lastReadTime)
            {
                return _cachedConfig; // Return cached version instantly
            }

            // 3. Load New Config
            try
            {
                string json = File.ReadAllText(_configPath);
                _cachedConfig = JsonSerializer.Deserialize<PluginConfiguration>(json);
                _lastReadTime = lastModified;
                return _cachedConfig;
            }
            catch
            {
                return _cachedConfig; // If file is locked, return old cache
            }
        }
    }
}