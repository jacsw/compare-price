using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace BackendCore.Source.Config
{
    class CapitalConfig
    {
        public CapitalConfig(string filename)
        {
            mConfigFilename = filename;
        }

        public void LoadConfig()
        {
            StreamReader sr = new StreamReader(mConfigFilename);
            string jsonString = string.Empty;
            jsonString = sr.ReadToEnd();
            sr.Close();

            DataContractJsonSerializer js = new DataContractJsonSerializer(typeof(JsonCapitalData[]));

            MemoryStream memJs = new MemoryStream();
            StreamWriter wr = new StreamWriter(memJs);
            wr.Write(jsonString);
            wr.Flush();
            memJs.Position = 0;
            mCapitalData = (JsonCapitalData[])js.ReadObject(memJs);

            System.Console.WriteLine("Config Size : {0}", mCapitalData.GetLength(0));

            for (int i = 0; i < mCapitalData.GetLength(0); i++)
            {
                System.Console.WriteLine("--------------------------------------------------");
                System.Console.WriteLine("- Com  : {0}", mCapitalData[i].Com);
                System.Console.WriteLine("- File : {0}", mCapitalData[i].File);
                if (mCapitalData[i].Rate != null)
                {
                    System.Console.WriteLine("- Rate : ({0}, {1}) / (Row, Col)",
                        mCapitalData[i].Rate.Row, mCapitalData[i].Rate.Col);
                }
                if (mCapitalData[i].Price != null)
                {
                    System.Console.WriteLine("- Price : ({0}, {1}) / (Row, Col)",
                        mCapitalData[i].Price.Row, mCapitalData[i].Price.Col);
                }
                if (mCapitalData[i].Fee.M36 != null)
                {
                    System.Console.WriteLine("- Fee.M36 : ({0}, {1}) / (Row, Col)",
                        mCapitalData[i].Fee.M36.Row, mCapitalData[i].Fee.M36.Col);
                }
                if (mCapitalData[i].Fee.M48 != null)
                {
                    System.Console.WriteLine("- Fee.M48 : ({0}, {1}) / (Row, Col)",
                        mCapitalData[i].Fee.M48.Row, mCapitalData[i].Fee.M48.Col);
                }
                if (mCapitalData[i].Fee.M60 != null)
                {
                    System.Console.WriteLine("- Fee.M60 : ({0}, {1}) / (Row, Col)",
                        mCapitalData[i].Fee.M60.Row, mCapitalData[i].Fee.M60.Col);
                }
            }

            System.Console.WriteLine("--------------------------------------------------");
        }

        public int GetCount()
        {
            if (mCapitalData == null)
            {
                return -1;
            }

            return mCapitalData.GetLength(0);
        }

        public JsonCapitalData GetCapitalData(int n)
        {
            if (n < 0 || n >= GetCount())
            {
                return null;
            }

            return mCapitalData[n];
        }

        private string mConfigFilename;
        private JsonCapitalData[] mCapitalData;
    }
}
