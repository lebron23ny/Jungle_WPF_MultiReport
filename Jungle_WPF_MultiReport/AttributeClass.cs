using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jungle_WPF_MultiReport
{
    public class AttributeClass
    {
        public string FullName { get; set; }
        public string NameTag
        {
            get
            {
                FileInfo file = new FileInfo(FullName);
                string s = file.Name;
                string d = s.Replace("_Jungle_MultiReport.xml", "");
                return d;
            }
        }
        public string Directory
        {
            get
            {
                FileInfo file = new FileInfo(FullName);
                string directory = file.DirectoryName;
                return directory;
            }
        }


        public AttributeClass(string fullName)
        {
            FullName = fullName;
        }

        public override string ToString()
        {
            return FullName;
        }

        public override bool Equals(object obj)
        {
            return obj is AttributeClass @class &&
                   NameTag == @class.NameTag;
        }
    }
}
