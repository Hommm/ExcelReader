using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.Model
{
    public class Group : IComparable
    {
        public string Name { get; set; }

        public string Type { get; set; }

        public string Belongs { get; set; }

        public Group() { }

        public Group(string name, string type, string belongs)
        {
            this.Name = name;
            this.Type = type;
            this.Belongs = belongs;
        }

        public int CompareTo(object obj)
        {
            return Name.CompareTo(((Group)obj).Name);
        }

        public override bool Equals(object obj)
        {
            return Name.Equals(((Group)obj).Name);
        }

        public override int GetHashCode()
        {
            int result = 17;
            result = 37 * result + this.Name.GetHashCode();
            return result; 
        }
    }
}
