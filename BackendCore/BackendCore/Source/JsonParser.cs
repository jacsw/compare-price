using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace BackendCore.Source
{
    [DataContract]
    class JsonReceive
    {
        [DataMember]
        public int Rate;
        [DataMember]
        public int Price;
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

    class JsonParser
    {
        public static string ComposeJson(JsonSend[] calculateData)
        {
            string jsonString = string.Empty;

            DataContractJsonSerializer js = new DataContractJsonSerializer(typeof(JsonSend[]));
            MemoryStream mem = new MemoryStream();
            js.WriteObject(mem, calculateData);
            mem.Position = 0;

            StreamReader sr = new StreamReader(mem);
            jsonString = sr.ReadToEnd();

            return jsonString;
        }

        public static JsonReceive ParseJson(string jsonString)
        {
            JsonReceive parseReceive;
            DataContractJsonSerializer js = new DataContractJsonSerializer(typeof(JsonReceive));

            MemoryStream memJs = new MemoryStream();
            StreamWriter wr = new StreamWriter(memJs);
            wr.Write(jsonString);
            wr.Flush();
            memJs.Position = 0;

            try
            {
                parseReceive = (JsonReceive)js.ReadObject(memJs);
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
