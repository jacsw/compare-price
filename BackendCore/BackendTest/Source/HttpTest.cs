using System;
using System.IO;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace BackendTest
{
    [DataContract]
    class JsonRequest
    {
        [DataMember]
        public int Price;

        [DataMember]
        public int Rate;
    }

    [DataContract]
    class JsonSend
    {
        [DataMember]
        public string Com;
        [DataMember]
        public int Rate;
        [DataMember]
        public int FeeM36;
        [DataMember]
        public int FeeM48;
        [DataMember]
        public int FeeM60;
    }

    class HttpTest
    {
        private static string URL = "http://127.0.0.1:40000/";
 
        static string MakeJsonString(JsonRequest request)
        {
            DataContractJsonSerializer js = new DataContractJsonSerializer(typeof(JsonRequest));
            MemoryStream mem = new MemoryStream();
            js.WriteObject(mem, request);
            mem.Position = 0;

            StreamReader sr = new StreamReader(mem);
            string jsonString = sr.ReadToEnd();

            return jsonString;
        }

        static string RequestHttpServer(string url, string json)
        {
            HttpWebRequest request = null;

            // Request using Http Connection JSON Data
            try
            {
                request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "POST";
                request.ContentType = "application/json";
                request.Timeout = 30 * 1000;

                // 2. POST할 데이타를 Request Stream에 쓴다
                byte[] bytes = Encoding.ASCII.GetBytes(json);
                request.ContentLength = bytes.Length; // 바이트수 지정

                Stream reqStream = request.GetRequestStream();
                reqStream.Write(bytes, 0, bytes.Length);
            }
            catch
            {
                System.Console.WriteLine("RequestHttpServer / Request Error!");
            }

            // Process Response Data
            string responseText = String.Empty;
            try
            {
                char[] buffer = new char[1000];
                WebResponse resp = request.GetResponse();

                Stream respStream = resp.GetResponseStream();
                StreamReader sr = new StreamReader(respStream);
                int recvlen = sr.Read(buffer, 0, 1000);
                sr.Close();
                resp.Close();

                responseText = new string(buffer);
            }
            catch
            {
                System.Console.WriteLine("RequestHttpServer / Request Error!");
            }

            return responseText;
        }

        static void Main(string[] args)
        {
            while(true)
            {
                System.Console.Write("Price : ");
                int price = Convert.ToInt32(Console.ReadLine());

                if (price == 0) break;

                System.Console.Write("remain Rate: ");
                int remainRate = Convert.ToInt32(Console.ReadLine());

                // 1. JSON string 만들기
                var reqData = new JsonRequest { Price = price, Rate = remainRate };
                string jsonString = MakeJsonString(reqData);
                System.Console.WriteLine("Json : {0}", jsonString);

                // Request using Http Connection JSON Data
                string response = RequestHttpServer(URL, jsonString);
                System.Console.WriteLine("receive : {0}", response);
            }
        }
    }
}
