using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace BackendCore.Source.Config
{
    class CapitalData
    {
        public string CapitalName;
        public string ConfigFile;
        public JsonCapitalConfig Config;
    }

    class CapitalConfig
    {
        public CapitalConfig(string filename)
        {
            mConfigFilename = filename;
        }

        public void LoadConfig() // LoadCapitalList()
        {
            StreamReader sr = new StreamReader(mConfigFilename);
            string jsonString = string.Empty;
            jsonString = sr.ReadToEnd();
            sr.Close();

            DataContractJsonSerializer js = new DataContractJsonSerializer(typeof(JsonCapitalList[]));

            MemoryStream memJs = new MemoryStream();
            StreamWriter wr = new StreamWriter(memJs);
            wr.Write(jsonString);
            wr.Flush();
            memJs.Position = 0;

            JsonCapitalList[] capitalList = (JsonCapitalList[])js.ReadObject(memJs);

            mCapitalData = new CapitalData[capitalList.GetLength(0)];


            System.Console.WriteLine("Config Size : {0}", mCapitalData.GetLength(0));

            System.Console.WriteLine("--------------------------------------------------");
            for (int i = 0; i < mCapitalData.GetLength(0); i++)
            {
                mCapitalData[i] = new CapitalData();
                mCapitalData[i].CapitalName = capitalList[i].Capital;
                mCapitalData[i].ConfigFile  = capitalList[i].File;

                System.Console.WriteLine("* Com : {0} / Config File : {1}", mCapitalData[i].CapitalName, mCapitalData[i].ConfigFile);
                LoadCapitalData(mCapitalData[i]);


            }
            System.Console.WriteLine("--------------------------------------------------");
        }


        private void LoadCapitalData(CapitalData pCapitalData)
        {
            StreamReader sr = new StreamReader(pCapitalData.ConfigFile);
            string jsonString = string.Empty;
            jsonString = sr.ReadToEnd();
            sr.Close();

            DataContractJsonSerializer js = new DataContractJsonSerializer(typeof(JsonCapitalConfig));

            MemoryStream memJs = new MemoryStream();
            StreamWriter wr = new StreamWriter(memJs);
            wr.Write(jsonString);
            wr.Flush();
            memJs.Position = 0;
            pCapitalData.Config = (JsonCapitalConfig)js.ReadObject(memJs);
            System.Console.WriteLine("  - Excel File : {0}", pCapitalData.Config.ExcelFile);
            System.Console.WriteLine("  - Work Sheet : {0}", pCapitalData.Config.Worksheet);
            System.Console.WriteLine("  - CarInfo.Company : ({0}, {1})", pCapitalData.Config.CarInfo.Company.Row,
                pCapitalData.Config.CarInfo.Company.Col);

        }

        public int GetCount()
        {
            if (mCapitalData == null)
            {
                return -1;
            }

            return mCapitalData.GetLength(0);
        }

        public CapitalData GetCapitalData(int n)
        {
            if (n < 0 || n >= GetCount())
            {
                return null;
            }

            return mCapitalData[n];
        }

        private string mConfigFilename;
        private CapitalData[] mCapitalData;
    }
}
