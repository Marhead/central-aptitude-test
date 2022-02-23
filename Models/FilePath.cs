using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CentralAptitudeTest.Models
{
    public class FilePath
    {
        public string filePath { get; set; }

        public List<Dictionary<string, List<string>>> College_Dictionarys { get; set; }
        
        public FilePath()
        {
            
        }
    }
}
