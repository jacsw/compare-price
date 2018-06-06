using BackendCore.Source;
using System;

namespace BackendCore.Source
{
    class MainModule
    {
        static private string config = @"D:\Rent\Excel\CapitalList.json";
        static private string mCarListFile = @"D:\Rent\CarList_20180529.xlsx";
        static private string mCarListSheet = @"현대";

        static void Main(string[] args)
        {
            MainModule main = new MainModule();
            main.Init();
//            main.Run();
            main.Deinit();
        }

        public void Init()
        {
            // TODO : Set by JSON File
            mConfig = new Config.CapitalConfig(config);
            mConfig.LoadConfig();

            System.Console.WriteLine("--------------------------------------------------");

            ExcelParser.ExcelInstance.ModuleInit();

            var CarList = new ExcelParser.ExcelCarList(mCarListFile, mCarListSheet);
            CarList.Init();
            CarList.Deinit();

            mExcels = new ExcelParser.ExcelBase[mConfig.GetCount()];
            for (int i = 0; i < mConfig.GetCount(); i++)
            {
                var config = mConfig.GetCapitalData(i);
                switch(config.CapitalName)
                {
                    case "하나캐피탈" :
                        mExcels[i] = new ExcelParser.ExcelHana(config);
                        mExcels[i].Init();
                        break;
                    case "효성캐피탈":
                        mExcels[i] = new ExcelParser.ExcelHyosung(config);
                        mExcels[i].Init();
                        break;
                    case "JB우리캐피탈":
                        mExcels[i] = new ExcelParser.ExcelJBWoori(config);
                        mExcels[i].Init();
                        break;
                }
            }
        }

        public void Run()
        {
            Console.WriteLine("Running sync server.");
            HttpService service = new HttpService(mExcels);

            System.Console.Write("\nPress Enter Key!!");
            System.Console.ReadKey();
        }

        public void Deinit()
        {
            for (int i = 0; i < mConfig.GetCount(); i++)
            {
                if (mExcels[i] != null) mExcels[i].Deinit();
            }

            ExcelParser.ExcelInstance.ModuleDeInit();
        }

        private ExcelParser.ExcelBase[] mExcels;
        private Config.CapitalConfig mConfig;
    }
}
