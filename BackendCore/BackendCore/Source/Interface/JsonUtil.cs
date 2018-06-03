using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace BackendCore.Source.Interface
{
    class JsonUtil
    {
        public static string toJsonResponse(JsonResponse[] calculateData)
        {
            string jsonString = string.Empty;

            DataContractJsonSerializer js = new DataContractJsonSerializer(typeof(JsonResponse[]));
            MemoryStream mem = new MemoryStream();
            js.WriteObject(mem, calculateData);
            mem.Position = 0;

            StreamReader sr = new StreamReader(mem);
            jsonString = sr.ReadToEnd();

            return jsonString;
        }

        public static JsonRequest fromJsonRequest(string jsonString)
        {
            JsonRequest parseReceive;
            DataContractJsonSerializer js = new DataContractJsonSerializer(typeof(JsonRequest));

            MemoryStream memJs = new MemoryStream();
            StreamWriter wr = new StreamWriter(memJs);
            wr.Write(jsonString);
            wr.Flush();
            memJs.Position = 0;

            try
            {
                parseReceive = (JsonRequest)js.ReadObject(memJs);
                return parseReceive;
            }
            catch
            {
                System.Console.WriteLine("- Json Parse Error!");
            }

            return null;
        }
    }
}
