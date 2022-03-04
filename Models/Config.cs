using System.IO;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Newtonsoft.Json;
using CentralAptitudeTest.Models;

namespace CentralAptitudeTest.Models
{
    public class Config
    {
        // Config.conf 파일 위치
        public static string runPath = Path.GetDirectoryName(typeof(Config).Assembly.Location);
        public static string ConfPath = string.Format("{0}/config", runPath);

        public FilePath FilePath { get; set; }

        public List<Dictionary<string, List<string>>> CollegeDictionaries { get; set; }

        public Config()
        {
            FilePath = new FilePath();
        }

        public static Config GetConfig()
        {
            string confPath = $"{ConfPath}/config.conf";

            if (File.Exists(confPath))
            {
                //ConnectionList = JsonConvert.DeserializeObject<ObservableCollection<TreeNode>>(File.ReadAllText(confPath));
                return JsonConvert.DeserializeObject<Config>(File.ReadAllText(confPath));
            }
            else
            {
                Config config = new Config();
                config.FilePath = new FilePath() { filePath = "파일 없음" };
                Config.SetConfig(config);
                //ConnectionList = new ObservableCollection<TreeNode>();
                //ConnectionList.Add(new TreeNode() { Id = "연결", Type = TreeViewItemType.Root });
                return config;
            }

            //Connections = CollectionViewSource.GetDefaultView(ConnectionList);
        }

        public static void SetConfig(Config conf)
        {
            File.WriteAllText($"{ConfPath}/config.conf", JsonConvert.SerializeObject(conf, Formatting.Indented));
        }

        public static bool InitSystem()
        {
            if (!Directory.Exists(ConfPath))
                Directory.CreateDirectory(ConfPath);

            return true;
        }


    }
}
