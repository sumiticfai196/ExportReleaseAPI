using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp_ExportAPI
{

    public class RootObject
    {
        public bool morerecords { get; set; }
       // public string paging-cookie-encoded { get; set; }
        public string totalrecords { get; set; }
        public int page { get; set; }
        public List<Product> results { get; set; }
    }


}
