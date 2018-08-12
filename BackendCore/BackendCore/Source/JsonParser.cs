using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace BackendCore.Source
{
    class JsonParser
    {
        public static string ComposeJson(Interface.JsonResponse calculateData)
        {
            string jsonString = string.Empty;

            DataContractJsonSerializer js = new DataContractJsonSerializer(typeof(Interface.JsonResponse));
            MemoryStream mem = new MemoryStream();
            js.WriteObject(mem, calculateData);
            mem.Position = 0;

            StreamReader sr = new StreamReader(mem);
            jsonString = sr.ReadToEnd();

            return jsonString;
        }

        public static Interface.JsonRequest ParseJson(string jsonString)
        {
            Interface.JsonRequest parseReceive;
            DataContractJsonSerializer js = new DataContractJsonSerializer(typeof(Interface.JsonRequest));

            MemoryStream memJs = new MemoryStream();
            StreamWriter wr = new StreamWriter(memJs);
            wr.Write(jsonString);
            wr.Flush();
            memJs.Position = 0;

            try
            {
                parseReceive = (Interface.JsonRequest)js.ReadObject(memJs);
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
