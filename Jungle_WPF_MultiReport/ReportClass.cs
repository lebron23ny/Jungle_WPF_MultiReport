

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jungle_WPF_MultiReport
{
    static class Extensions
    {
        public static void Sort<T>(this ObservableCollection<T> collection) where T : IComparable
        {
            List<T> sorted = collection.OrderBy(x => x).ToList();
            for (int i = 0; i < sorted.Count(); i++)
                collection.Move(collection.IndexOf(sorted[i]), i);
        }
    }


    public class ReportClass:IComparable
    {
        public bool Flag { get; set; }
        public string Name { get; set; }

        public int CompareTo(object obj)
        {
            ReportClass a = this;
            ReportClass b = obj as ReportClass;
            return string.Compare(a.Name, b.Name); ;
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
