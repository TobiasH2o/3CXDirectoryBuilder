using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Phone_book_Formatter
{
    internal class Extension
    {
        public string ID
        { get { return No.Substring(0, IDLength); } }

        public string No { get; set; }

        public int IDLength { get; set; }
        public string Name { get; set; }
    }
}