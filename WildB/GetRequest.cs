using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace WildB
{
    public class GetRequest
    {
        HttpWebRequest _request;
        string _url;
        public string Response { get; set; }

        public GetRequest(string url)
        {
            _url = url;
        }
        public void Run()
        {
            _request = (HttpWebRequest)WebRequest.Create(_url);
            _request.Method = "Get";
            try
            {
                HttpWebResponse response = (HttpWebResponse)_request.GetResponse();
                var stream = response.GetResponseStream();
                if (stream != null) Response = new StreamReader(stream).ReadToEnd();
            }
            catch(Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
