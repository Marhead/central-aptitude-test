using System.IO;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Newtonsoft.Json;
using CentralAptitudeTest.Models;

namespace CentralAptitudeTest.Models
{
    public class Config
    {
        public static string runPath = Path.GetDirectoryName(typeof(Config).Assembly.Location);
        public static string ConfPath = string.Format("{0}/config", runPath);

        public List<string> Subjects { get; set; }

        public FilePath FilePath { get; set; }

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
                Config conf = new Config();
                conf.FilePath = new FilePath() { filePath = "파일 없음" };
                Config.SetConfig(conf);
                //ConnectionList = new ObservableCollection<TreeNode>();
                //ConnectionList.Add(new TreeNode() { Id = "연결", Type = TreeViewItemType.Root });
                return conf;
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
