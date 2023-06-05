using System.Collections.Generic;

namespace ConsoleApp
{
    public class ApiResponse
    {
        public bool morerecords { get; set; }
        // public Dictionary<string, string> paging-cookie-encoded { get; set; }
        public string totalrecords { get; set; }
        public int page { get; set; }
        public List<Dictionary<string, string>> results { get; set; }
    }
}