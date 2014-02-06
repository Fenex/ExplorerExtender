using System.Configuration;
using System.IO;

namespace ExplorerExtender
{
    public static class SettingHelper
    {
        private static Configuration GetConfiguration()
        {
            string filename = Path.Combine(Path.GetDirectoryName(typeof(SettingHelper).Assembly.Location), "app.config");
            if (!File.Exists(filename))
                return null;

            ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap();
            fileMap.ExeConfigFilename = filename;

            return ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);
        }

        public static string ReadSetting(string key)
        {
            Configuration config = SettingHelper.GetConfiguration();
            if (config == null)
                return null;
            else
                return config.AppSettings.Settings[key].Value;
        }

        public static void WriteSetting(string key, string value)
        {
            Configuration config = SettingHelper.GetConfiguration();
            if (config == null)
                return;

            config.AppSettings.Settings[key].Value = value;
            config.Save(ConfigurationSaveMode.Minimal);
        }        
    }
}
