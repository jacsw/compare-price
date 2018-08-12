using System;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using BackendCore.Source.Interface;

namespace BackendTest
{
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
                request.Timeout = 60 * 1000;

                // 2. POST할 데이타를 Request Stream에 쓴다
                byte[] bytes = Encoding.UTF8.GetBytes(json);
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

                // System.Console.Write("remain Rate: ");
                // int remainRate = Convert.ToInt32(Console.ReadLine());

                // 1. JSON string 만들기
                JsonRequest reqData = new JsonRequest()
                {
                    RequestID = 1,
                    CarInfo = new JsonReq_CarInfo()
                    {
                        Company = "현대자동차",
                        Model = "그랜저IG",
                        Trim = "가솔린 3.3 셀러브리티"
                    },
                    Cost = new JsonReq_Cost
                    {
                        BasePrice = price, //  41600000,
                        OptionPrice = 2600000,
                        OptionInfo = "HUD, 스마트센트II",
                        Deposit = 10,
                        PrePayment = 0
                    },
                    Commission = new JsonReq_Commission
                    {
                        CMCommission = 2.5,
                        AGCommission = 5.5
                    }
                };


                string jsonString = MakeJsonString(reqData);
                System.Console.WriteLine("Json : {0}", jsonString);

                // Request using Http Connection JSON Data
                string response = RequestHttpServer(URL, jsonString);
                System.Console.WriteLine("receive : {0}", response);
            }
        }
    }
}
